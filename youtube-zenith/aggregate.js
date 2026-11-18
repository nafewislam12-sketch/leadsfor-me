/**
 * ZENITH MASTER AGGREGATOR
 * Combines results from all GitHub Action matrix workers / local workers
 * - De-duplicates channels by channelId / handle
 * - Merges emails and links
 * - Creates master CSV, Excel, and JSON files
 * - Generates markdown summary statistics for GitHub Actions Step Summary
 */

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

function aggregateResults(inputDir = './output', outputDir = './output') {
  console.log(`[Zenith Aggregator] Reading worker results from: ${inputDir}`);

  if (!fs.existsSync(inputDir)) {
    console.error(`Directory not found: ${inputDir}`);
    process.exit(1);
  }

  const files = fs.readdirSync(inputDir).filter(f => f.endsWith('.json') && !f.includes('master'));
  console.log(`Found ${files.length} worker result files:`, files);

  const channelMap = new Map();

  for (const file of files) {
    const fullPath = path.join(inputDir, file);
    try {
      const data = JSON.parse(fs.readFileSync(fullPath, 'utf-8'));
      if (Array.isArray(data)) {
        for (const item of data) {
          const key = item.channelId || item.handle || item.title;
          if (!key) continue;

          if (!channelMap.has(key)) {
            channelMap.set(key, item);
          } else {
            // Merge emails and links if duplicate channel found across workers
            const existing = channelMap.get(key);
            if (!existing.primaryEmail && item.primaryEmail) {
              existing.primaryEmail = item.primaryEmail;
              existing.hasEmail = true;
            }
            if (item.allEmails) {
              const combinedEmails = [...new Set([
                ...(existing.allEmails ? existing.allEmails.split(',').map(e => e.trim()) : []),
                ...(item.allEmails ? item.allEmails.split(',').map(e => e.trim()) : [])
              ])].filter(Boolean);
              existing.allEmails = combinedEmails.join(', ');
            }
            if (!existing.websiteUrl && item.websiteUrl) {
              existing.websiteUrl = item.websiteUrl;
              existing.hasCustomDomain = true;
            }
            if (!existing.skoolLink && item.skoolLink) {
              existing.skoolLink = item.skoolLink;
              existing.hasSkool = true;
            }
          }
        }
      }
    } catch (err) {
      console.warn(`Failed to parse file ${file}:`, err.message);
    }
  }

  const masterLeads = Array.from(channelMap.values());
  console.log(`[Zenith Aggregator] Total unique leads aggregated: ${masterLeads.length}`);

  const masterBasePath = path.join(outputDir, 'zenith_leads_master');

  // 1. Write master JSON
  fs.writeFileSync(`${masterBasePath}.json`, JSON.stringify(masterLeads, null, 2), 'utf-8');

  // 2. Write master CSV
  if (masterLeads.length > 0) {
    const headers = Object.keys(masterLeads[0]);
    const escapeCsv = (val) => {
      if (val === null || val === undefined) return '""';
      const str = String(val).replace(/"/g, '""').replace(/\r?\n/g, ' ');
      return `"${str}"`;
    };
    const csvContent = [
      headers.map(escapeCsv).join(','),
      ...masterLeads.map(row => headers.map(h => escapeCsv(row[h])).join(','))
    ].join('\n');
    fs.writeFileSync(`${masterBasePath}.csv`, csvContent, 'utf-8');
  }

  // 3. Write master XLSX
  const worksheet = xlsx.utils.json_to_sheet(masterLeads);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Zenith Master Leads');
  xlsx.writeFile(workbook, `${masterBasePath}.xlsx`);

  // 4. Update persistent history of seen channels
  try {
    const historyDir = path.join(__dirname, 'data');
    if (!fs.existsSync(historyDir)) {
      fs.mkdirSync(historyDir, { recursive: true });
    }
    const historyFilePath = path.join(historyDir, 'seen_channels.txt');
    const existingSet = new Set();
    if (fs.existsSync(historyFilePath)) {
      fs.readFileSync(historyFilePath, 'utf-8').split(/\r?\n/).forEach(s => {
        const trim = s.trim();
        if (trim) existingSet.add(trim);
      });
    }

    for (const lead of masterLeads) {
      if (lead.channelId) existingSet.add(lead.channelId);
      if (lead.handle) existingSet.add(lead.handle);
    }

    fs.writeFileSync(historyFilePath, Array.from(existingSet).join('\n'), 'utf-8');
    console.log(`[Zenith Aggregator] 🛡️ Recorded ${existingSet.size} total channels in persistent history: ${historyFilePath}`);
  } catch (err) {
    console.warn('[Zenith Aggregator] Failed to update seen_channels.txt:', err.message);
  }

  // Summary Metrics
  const totalLeads = masterLeads.length;
  const withEmails = masterLeads.filter(l => l.hasEmail).length;
  const withWebsites = masterLeads.filter(l => l.hasCustomDomain).length;
  const withSkool = masterLeads.filter(l => l.hasSkool).length;
  const emailRate = totalLeads > 0 ? ((withEmails / totalLeads) * 100).toFixed(1) : '0.0';

  console.log(`\n====================================================`);
  console.log(`  📊 ZENITH AGGREGATION COMPLETE!`);
  console.log(`====================================================`);
  console.log(`Total Leads:       ${totalLeads}`);
  console.log(`With Valid Email:  ${withEmails} (${emailRate}%)`);
  console.log(`With Custom Site:  ${withWebsites}`);
  console.log(`With Skool:        ${withSkool}`);
  console.log(`Saved Files:`);
  console.log(`  - ${masterBasePath}.csv`);
  console.log(`  - ${masterBasePath}.xlsx`);
  console.log(`  - ${masterBasePath}.json`);
  console.log(`====================================================\n`);

  // Output GitHub Actions Step Summary if running in GitHub Actions
  const summaryFile = process.env.GITHUB_STEP_SUMMARY;
  if (summaryFile) {
    const summaryMarkdown = `
# 🚀 Zenith YouTube Scraper Run Report

| Metric | Count | Percentage |
| :--- | :--- | :--- |
| **Total Unique Channels** | **${totalLeads}** | 100% |
| **Channels with Email** | **${withEmails}** | **${emailRate}%** |
| **Channels with Custom Domain** | **${withWebsites}** | ${totalLeads > 0 ? ((withWebsites / totalLeads) * 100).toFixed(1) : 0}% |
| **Channels with Skool Community** | **${withSkool}** | ${totalLeads > 0 ? ((withSkool / totalLeads) * 100).toFixed(1) : 0}% |

### 📁 Output Files Ready for Download:
- \`zenith_leads_master.xlsx\` (Full Excel Spreadsheet)
- \`zenith_leads_master.csv\` (Universal CSV file)
- \`zenith_leads_master.json\` (Raw JSON)

### 🏆 Top 5 Leads Preview:
${masterLeads.slice(0, 5).map(l => `- **${l.title}** (${l.handle || l.channelId}): Email: \`${l.primaryEmail || 'None'}\` | Website: ${l.websiteUrl || 'None'} | Skool: ${l.skoolLink || 'None'}`).join('\n')}
`;
    fs.appendFileSync(summaryFile, summaryMarkdown, 'utf-8');
  }

  return { totalLeads, withEmails, withWebsites, withSkool, emailRate };
}

if (require.main === module) {
  const inputDir = process.argv[2] || './output';
  const outputDir = process.argv[3] || './output';
  aggregateResults(inputDir, outputDir);
}

module.exports = { aggregateResults };

