# ⚡ Zenith YouTube Lead Scraper (Multi-IP Distributed)

High-performance YouTube Lead Generation & ICP Scraper built to run across **GitHub Actions Matrix Workers** to leverage **rotating datacenter IPs** and bypass YouTube rate limits completely.

---

## 🌟 Key Capabilities

1. **Distributed Multi-IP Architecture**:
   - Runs across 1 to 20 parallel GitHub Actions runners.
   - Each runner gets its own fresh Azure/GitHub public IP address.
   - No proxy purchase needed. Zero risk of home/office IP bans.

2. **Full ICP & Lead Extraction**:
   - Channel Title, Handle (`@handle`), Channel URL, Avatar.
   - Subscriber Count, Video Count, Total Views, Country, Joined Date.
   - Full Channel Bio / Description.
   - Recent Video Titles & Full Descriptions.
   - Channel Keywords & Tags.

3. **High-Accuracy Email Hunting Engine**:
   - Scans **Channel Bio** (including obfuscated emails like `name [at] domain.com`).
   - Scans **Recent Video Descriptions** (sponsorship/business contact emails).
   - Scans **External Domain Websites**: Crawls the creator's homepage & `/contact` or `/about` page with a 4-second timeout to find business emails.
   - Filters out junk (like image filenames, sentry.io, wixpress, example.com).

4. **Community & Business Intelligence**:
   - **Skool Community Detector**: Identifies if the creator runs a Skool community.
   - **Custom Domain Detector**: Distinguishes between custom agency/business websites vs generic links.
   - **Social Links**: Extracts Instagram, Twitter/X, LinkedIn, Facebook, Discord, Calendly.

5. **Full Control & Customizable Filters**:
   - `require_email`: Only keep leads with verified email, OR keep all channels and fill emails where found.
   - `require_domain`: Filter to only keep creators with business websites.
   - `require_skool`: Filter to only keep creators running Skool communities.
   - `speed_delay_ms`: Complete speed throttle control (200ms to 2000ms).
   - `workers`: Select 1 to 20 concurrent GitHub Action workers.

6. **Automatic Aggregation**:
   - All worker outputs are automatically downloaded, de-duplicated, and merged into:
     - `zenith_leads_master.xlsx` (Excel Spreadsheet)
     - `zenith_leads_master.csv` (Universal CSV)
     - `zenith_leads_master.json` (Structured JSON)

---

## 🚀 How to Run on GitHub Actions (Multiple IPs)

### Step 1: Push the repository to GitHub
```bash
git add .
git commit -m "Add Zenith Scraper and Workflow"
git push origin main
```

### Step 2: Trigger from GitHub Web Interface
1. Go to your repository on GitHub: `https://github.com/nafewislam11-spec/kineticstudio`
2. Click the **Actions** tab at the top.
3. In the left sidebar, click **Zenith YouTube Lead Scraper (Multi-IP Distributed)**.
4. Click **Run workflow**:
   - **Target Niches**: e.g., `b2b saas founder, real estate coach, ai automation consultant, fitness coach`
   - **Number of Workers**: e.g., `5` (runs 5 runners on 5 distinct IPs simultaneously!)
   - **Target Leads per Worker**: e.g., `50`
   - **Require Email**: Check or uncheck
   - **Require Domain**: Check or uncheck
   - **Require Skool**: Check or uncheck
   - **Speed / Delay**: e.g., `500ms`
5. Click the green **Run workflow** button!

### Step 3: Download Master Leads
When the workflow completes:
- Click the finished workflow run.
- Under **Artifacts** at the bottom, download `zenith-leads-master-<id>`.
- You will get the consolidated `zenith_leads_master.xlsx` and `zenith_leads_master.csv`!

---

## 💻 How to Run Locally (On your Mac / PC)

### 1. Launch the Local Web Dashboard
```bash
cd youtube-zenith
npm start
```
Open [http://localhost:4000](http://localhost:4000) in your browser:
- Configure search queries and filters in a sleek dark UI.
- Watch leads stream in real time.
- Download CSV/Excel with 1 click.

### 2. Run via Terminal CLI
```bash
cd youtube-zenith
node run_cli.js \
  --niches "b2b saas founder, real estate coach" \
  --limit 50 \
  --worker 1 \
  --total-workers 1 \
  --delay 500 \
  --require-email false \
  --require-domain false \
  --require-skool false
```

To aggregate multiple local worker files:
```bash
node aggregate.js output output
```

