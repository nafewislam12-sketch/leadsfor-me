const express = require('express');
const http = require('http');
const { WebSocketServer } = require('ws');
const path = require('path');
const fs = require('fs');
const { runWorker } = require('./zenith_scraper');
const { aggregateResults } = require('./aggregate');

const app = express();
const server = http.createServer(app);
const wss = new WebSocketServer({ server });

app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

let activeJobs = new Map();

function broadcast(msg) {
  const data = JSON.stringify(msg);
  wss.clients.forEach(client => {
    if (client.readyState === 1) client.send(data);
  });
}

// API: Start Scraper Job Locally
app.post('/api/start', async (req, res) => {
  const {
    niches = 'b2b saas founder, real estate coach',
    workers = 2,
    limit = 20,
    requireEmail = false,
    requireDomain = false,
    requireSkool = false,
    excludeThirdWorld = true,
    tier1Only = false,
    minSubs = 0,
    maxSubs = 0,
    minViews = 0,
    maxViews = 0,
    maxUploadDays = 0,
    delay = 500,
    videoCount = 3,
    deepWebsite = true
  } = req.body;

  const jobId = Date.now().toString();
  const nicheList = typeof niches === 'string' ? niches.split(',').map(s => s.trim()).filter(Boolean) : niches;
  const numWorkers = parseInt(workers, 10) || 1;
  const outDir = path.join(__dirname, 'output', `job_${jobId}`);

  fs.mkdirSync(outDir, { recursive: true });

  res.json({ success: true, jobId, message: `Started ${numWorkers} workers for job ${jobId}` });

  broadcast({ type: 'JOB_STARTED', jobId, niches: nicheList, numWorkers });

  // Run workers concurrently
  const workerPromises = [];
  for (let w = 1; w <= numWorkers; w++) {
    const workerPromise = runWorker({
      workerId: w,
      totalWorkers: numWorkers,
      niches: nicheList,
      maxChannelsPerWorker: Math.ceil(limit / numWorkers),
      filters: {
        requireEmail: Boolean(requireEmail),
        requireDomain: Boolean(requireDomain),
        requireSkool: Boolean(requireSkool),
        excludeThirdWorld: Boolean(excludeThirdWorld),
        tier1Only: Boolean(tier1Only),
        minSubscribers: parseInt(minSubs, 10) || 0,
        maxSubscribers: parseInt(maxSubs, 10) || 0,
        minViews: parseInt(minViews, 10) || 0,
        maxViews: parseInt(maxViews, 10) || 0,
        maxUploadAgeDays: parseInt(maxUploadDays, 10) || 0
      },
      scrapingOptions: {
        scanVideoCount: parseInt(videoCount, 10) || 3,
        deepScanWebsite: Boolean(deepWebsite),
        requestDelayMs: parseInt(delay, 10) || 500
      },
      outputDir: outDir,
      onProgress: (p) => {
        broadcast({ type: 'WORKER_PROGRESS', jobId, ...p });
      }
    });
    workerPromises.push(workerPromise);
  }

  try {
    await Promise.all(workerPromises);
    const aggResult = aggregateResults(outDir, outDir);
    broadcast({ type: 'JOB_COMPLETED', jobId, stats: aggResult });
  } catch (err) {
    console.error('Job execution error:', err);
    broadcast({ type: 'JOB_ERROR', jobId, error: err.message });
  }
});

// API: Get latest leads
app.get('/api/leads/latest', (req, res) => {
  const masterPath = path.join(__dirname, 'output', 'zenith_leads_master.json');
  if (fs.existsSync(masterPath)) {
    try {
      const data = JSON.parse(fs.readFileSync(masterPath, 'utf-8'));
      return res.json({ success: true, count: data.length, leads: data });
    } catch {
      return res.json({ success: false, leads: [] });
    }
  }
  res.json({ success: true, leads: [] });
});

// API: Download Master CSV
app.get('/api/download/csv', (req, res) => {
  const file = path.join(__dirname, 'output', 'zenith_leads_master.csv');
  if (fs.existsSync(file)) {
    return res.download(file, 'zenith_leads_master.csv');
  }
  res.status(404).send('No master CSV available yet. Run a scrape first!');
});

// API: Download Master Excel
app.get('/api/download/xlsx', (req, res) => {
  const file = path.join(__dirname, 'output', 'zenith_leads_master.xlsx');
  if (fs.existsSync(file)) {
    return res.download(file, 'zenith_leads_master.xlsx');
  }
  res.status(404).send('No master Excel file available yet. Run a scrape first!');
});

const PORT = process.env.PORT || 4050;
server.listen(PORT, () => {
  console.log(`[Zenith Dashboard] Running on http://localhost:${PORT}`);
});

