/**
 * ZENITH CLI RUNNER
 * Usage:
 *   node run_cli.js --worker 1 --total-workers 5 --niches "b2b saas, coaching" --limit 50
 */

const { runWorker } = require('./zenith_scraper');

function parseArgs() {
  const args = process.argv.slice(2);
  const options = {
    workerId: 1,
    totalWorkers: 1,
    niches: ['b2b saas growth', 'real estate coach'],
    maxChannelsPerWorker: 50,
    filters: {
      requireEmail: false,
      requireDomain: false,
      requireSkool: false,
      excludeThirdWorld: true,
      tier1Only: false,
      targetCountries: [],
      minSubscribers: 0,
      maxSubscribers: 0,
      minViews: 0,
      maxViews: 0,
      maxUploadAgeDays: 0,
      businessKeywords: []
    },
    scrapingOptions: {
      scanVideoCount: 3,
      deepScanWebsite: true,
      requestDelayMs: 600
    },
    outputDir: './output'
  };

  for (let i = 0; i < args.length; i++) {
    const arg = args[i];
    const next = args[i + 1];

    if (arg === '--worker' || arg === '-w') {
      options.workerId = parseInt(next, 10) || 1;
      i++;
    } else if (arg === '--total-workers' || arg === '-tw') {
      options.totalWorkers = parseInt(next, 10) || 1;
      i++;
    } else if (arg === '--niches' || arg === '-n') {
      options.niches = next.split(',').map(s => s.trim()).filter(Boolean);
      i++;
    } else if (arg === '--limit' || arg === '-l') {
      options.maxChannelsPerWorker = parseInt(next, 10) || 50;
      i++;
    } else if (arg === '--require-email') {
      options.filters.requireEmail = next === 'true';
      i++;
    } else if (arg === '--require-domain') {
      options.filters.requireDomain = next === 'true';
      i++;
    } else if (arg === '--require-skool') {
      options.filters.requireSkool = next === 'true';
      i++;
    } else if (arg === '--exclude-third-world') {
      options.filters.excludeThirdWorld = next === 'true';
      i++;
    } else if (arg === '--tier1-only') {
      options.filters.tier1Only = next === 'true';
      i++;
    } else if (arg === '--target-countries' || arg === '-tc') {
      options.filters.targetCountries = next.split(',').map(s => s.trim()).filter(Boolean);
      i++;
    } else if (arg === '--min-subs') {
      options.filters.minSubscribers = parseInt(next, 10) || 0;
      i++;
    } else if (arg === '--max-subs') {
      options.filters.maxSubscribers = parseInt(next, 10) || 0;
      i++;
    } else if (arg === '--min-views') {
      options.filters.minViews = parseInt(next, 10) || 0;
      i++;
    } else if (arg === '--max-views') {
      options.filters.maxViews = parseInt(next, 10) || 0;
      i++;
    } else if (arg === '--max-upload-days') {
      options.filters.maxUploadAgeDays = parseInt(next, 10) || 0;
      i++;
    } else if (arg === '--business-keywords') {
      options.filters.businessKeywords = next.split(',').map(s => s.trim()).filter(Boolean);
      i++;
    } else if (arg === '--delay') {
      options.scrapingOptions.requestDelayMs = parseInt(next, 10) || 600;
      i++;
    } else if (arg === '--video-count') {
      options.scrapingOptions.scanVideoCount = parseInt(next, 10) || 3;
      i++;
    } else if (arg === '--deep-website') {
      options.scrapingOptions.deepScanWebsite = next === 'true';
      i++;
    } else if (arg === '--output-dir' || arg === '-o') {
      options.outputDir = next;
      i++;
    }
  }

  return options;
}

async function main() {
  const config = parseArgs();
  try {
    await runWorker(config);
    process.exit(0);
  } catch (err) {
    console.error('Fatal error in worker execution:', err);
    process.exit(1);
  }
}

main();

