/**
 * ZENITH YOUTUBE SCRAPER & LEAD EXTRACTION ENGINE
 * 
 * Features:
 * - Searches YouTube for channels and top videos matching target ICP / niches
 * - Extracts Channel Title, Handle, URL, Subscriber Count, Video Count, Views, Bio
 * - Deep video scanning: Extracts descriptions & links from recent videos
 * - Social link detector: Skool, Instagram, Twitter/X, LinkedIn, Discord, Linktree, Facebook
 * - High-precision Email Extraction: Bio, Video Descriptions, Linktree, and Custom Domains
 * - Deep Website Scraper: Scrapes creator's external website & /contact for business emails
 * - Flexible Filters: require_email, require_domain, require_skool, business_tags
 * - Sharded multi-worker architecture designed for GitHub Actions Matrix (Multiple IPs)
 */

const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Rotating User Agents
const USER_AGENTS = [
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
  'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0',
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 14.4; rv:125.0) Gecko/20100101 Firefox/125.0'
];

function getRandomUserAgent() {
  return USER_AGENTS[Math.floor(Math.random() * USER_AGENTS.length)];
}

// Third-World & Developing Countries Blacklist (Strict Exclusion)
const THIRD_WORLD_COUNTRIES = [
  'india', 'pakistan', 'bangladesh', 'nigeria', 'philippines', 'kenya',
  'egypt', 'vietnam', 'indonesia', 'ghana', 'nepal', 'sri lanka', 'uganda',
  'south africa', 'brazil', 'algeria', 'morocco', 'ethiopia', 'tanzania',
  'afghanistan', 'yemen', 'sudan', 'syria', 'venezuela', 'myanmar', 'cuba',
  'zimbabwe', 'somalia', 'iraq', 'iran', 'cambodia', 'laos', 'bolivia',
  'uzbekistan', 'kazakhstan', 'colombia', 'peru', 'guatemala', 'ecuador'
];

const THIRD_WORLD_TLDS = [
  '.in', '.pk', '.bd', '.ng', '.ph', '.ke', '.eg', '.vn', '.id', '.gh',
  '.np', '.lk', '.ug', '.za', '.br', '.dz', '.ma', '.et', '.tz', '.af', '.ru'
];

const THIRD_WORLD_PHONE_PREFIXES = [
  '+91', '+92', '+880', '+234', '+63', '+254', '+20', '+84', '+62', '+233',
  '+977', '+94', '+256', '+27', '+55', '+7'
];

const TIER_1_COUNTRIES = [
  'united states', 'united kingdom', 'canada', 'australia', 'new zealand',
  'germany', 'france', 'netherlands', 'sweden', 'norway', 'denmark',
  'switzerland', 'ireland', 'singapore', 'united arab emirates', 'israel',
  'spain', 'italy', 'austria', 'belgium', 'finland', 'portugal', 'us', 'uk'
];

function parseNumericCount(text) {
  if (!text) return 0;
  if (typeof text === 'number') return text;
  const clean = text.replace(/,/g, '').trim().toUpperCase();
  const mMatch = clean.match(/([\d.]+)\s*M/);
  if (mMatch) return Math.round(parseFloat(mMatch[1]) * 1000000);
  const kMatch = clean.match(/([\d.]+)\s*K/);
  if (kMatch) return Math.round(parseFloat(kMatch[1]) * 1000);
  const numMatch = clean.match(/[\d.]+/);
  return numMatch ? Math.round(parseFloat(numMatch[0])) : 0;
}

function parseDaysAgo(text) {
  if (!text) return 9999;
  const clean = text.toLowerCase().trim();
  if (clean.includes('premiere') || clean.includes('stream') || clean.includes('live')) return 0;
  const yearMatch = clean.match(/(\d+)\s*(?:y|year|years)/);
  if (yearMatch) return parseInt(yearMatch[1], 10) * 365;
  const monthMatch = clean.match(/(\d+)\s*(?:mo|month|months)/);
  if (monthMatch) return parseInt(monthMatch[1], 10) * 30;
  const weekMatch = clean.match(/(\d+)\s*(?:w|week|weeks)/);
  if (weekMatch) return parseInt(weekMatch[1], 10) * 7;
  const dayMatch = clean.match(/(\d+)\s*(?:d|day|days)/);
  if (dayMatch) return parseInt(dayMatch[1], 10);
  if (clean.includes('yesterday')) return 1;
  return 0;
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function getChannelUrl(handleOrId) {
  if (!handleOrId) return '';
  if (handleOrId.startsWith('http')) return handleOrId;
  if (handleOrId.startsWith('UC') && handleOrId.length === 24) {
    return `https://www.youtube.com/channel/${handleOrId}`;
  }
  const clean = handleOrId.startsWith('/') ? handleOrId : `/${handleOrId}`;
  return `https://www.youtube.com${clean}`;
}

async function safeFetch(url, options = {}, retries = 2) {
  const timeoutMs = options.timeoutMs || 8000;
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      const res = await fetch(url, {
        ...options,
        signal: AbortSignal.timeout(timeoutMs),
        headers: {
          'User-Agent': getRandomUserAgent(),
          ...options.headers
        }
      });
      return res;
    } catch (err) {
      if (attempt === retries) {
        return null;
      }
      await sleep(500 * (attempt + 1));
    }
  }
  return null;
}

// Email blacklist & validation patterns
const EMAIL_BLACKLIST = [
  'example.com', 'domain.com', 'email.com', 'yourcompany.com',
  'sentry.io', 'wixpress.com', 'squarespace.com', 'shopify.com',
  'google.com', 'youtube.com', 'ytimg.com', 'github.com',
  'patreon.com', 'facebook.com', 'instagram.com', 'twitter.com',
  '2x.png', '3x.png', 'default.jpg', 'hqdefault.jpg'
];

const JUNK_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.svg', '.webp', '.gif', '.css', '.js'];

/**
 * Extracts and cleans valid emails from any block of text or HTML
 */
function extractEmails(text) {
  if (!text || typeof text !== 'string') return [];

  // Obfuscated email fixes & clean encoded entities (e.g. \u003e or &gt;)
  const cleanText = text
    .replace(/\\u003c|\\u003e|u003c|u003e|&lt;|&gt;|&amp;/gi, ' ')
    .replace(/\[at\]|\(at\)/gi, '@')
    .replace(/\[dot\]|\(dot\)/gi, '.')
    .replace(/\s*@\s*/g, '@');

  // Standard email regex
  const regex = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,7}\b/g;
  const matches = cleanText.match(regex) || [];

  const validEmails = new Set();
  for (let email of matches) {
    email = email.toLowerCase().trim();
    
    // Clean trailing punctuation
    email = email.replace(/^[<"']+|[>"',.:;]+$/g, '');

    // Check blacklist
    const domain = email.split('@')[1] || '';
    if (EMAIL_BLACKLIST.some(bad => domain.includes(bad))) continue;
    if (JUNK_EXTENSIONS.some(ext => email.endsWith(ext))) continue;
    if (email.length < 6 || email.length > 70) continue;

    validEmails.add(email);
  }

  return Array.from(validEmails);
}

/**
 * Extracts links from text (socials, skool, custom domains)
 */
function extractLinks(text) {
  if (!text) return { urls: [], skool: null, website: null, socials: {} };

  // Match URLs with or without http(s)
  const urlRegex = /(?:https?:\/\/|www\.)[^\s"'<>\)]+/gi;
  const matches = text.match(urlRegex) || [];
  
  // Also match bare skool.com mentions or links (e.g. skool.com/community-name or Skool.com)
  const skoolMatch = text.match(/\b(?:https?:\/\/)?(?:www\.)?skool\.com(?:\/([a-zA-Z0-9_-]+))?/i);

  const cleanUrls = [...new Set(matches.map(u => {
    let url = u.replace(/[\.,;:)]+$/, '');
    if (!url.startsWith('http')) url = `https://${url}`;
    return url;
  }))];

  let skool = null;
  if (skoolMatch) {
    skool = skoolMatch[1] ? `https://www.skool.com/${skoolMatch[1]}` : 'https://www.skool.com';
    if (!cleanUrls.includes(skool)) cleanUrls.push(skool);
  }

  let website = null;
  const socials = {
    instagram: null,
    twitter: null,
    linkedin: null,
    facebook: null,
    skool: skool,
    discord: null,
    calendly: null
  };

  for (const url of cleanUrls) {
    const lower = url.toLowerCase();
    if (lower.includes('skool.com')) {
      if (!skool) skool = url;
      socials.skool = url;
    } else if (lower.includes('instagram.com/')) {
      socials.instagram = url;
    } else if (lower.includes('twitter.com/') || lower.includes('x.com/')) {
      socials.twitter = url;
    } else if (lower.includes('linkedin.com/')) {
      socials.linkedin = url;
    } else if (lower.includes('facebook.com/')) {
      socials.facebook = url;
    } else if (lower.includes('discord.gg/') || lower.includes('discord.com/')) {
      socials.discord = url;
    } else if (lower.includes('calendly.com/') || lower.includes('cal.com/')) {
      socials.calendly = url;
    } else if (
      !lower.includes('youtube.com') &&
      !lower.includes('youtu.be') &&
      !lower.includes('google.com') &&
      !lower.includes('ytimg.com') &&
      !lower.includes('apple.com') &&
      !lower.includes('spotify.com') &&
      !lower.includes('amazon.com') &&
      !lower.includes('wikipedia.org') &&
      !website
    ) {
      website = url;
    }
  }

  return { urls: cleanUrls, skool, website, socials };
}

/**
 * Scrapes an external website & /contact page for emails
 */
async function scrapeDomainForEmails(websiteUrl, timeoutMs = 4000) {
  if (!websiteUrl || !websiteUrl.startsWith('http')) return [];
  const emails = new Set();

  async function fetchWithTimeout(url) {
    try {
      const controller = new AbortController();
      const id = setTimeout(() => controller.abort(), timeoutMs);
      const res = await fetch(url, {
        headers: {
          'User-Agent': getRandomUserAgent(),
          'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8'
        },
        signal: controller.signal,
        redirect: 'follow'
      });
      clearTimeout(id);
      if (!res.ok) return '';
      return await res.text();
    } catch {
      return '';
    }
  }

  try {
    const baseUrl = new URL(websiteUrl).origin;
    
    // Check homepage & /contact in parallel with fast timeout
    const [homeHtml, contactHtml] = await Promise.all([
      fetchWithTimeout(websiteUrl),
      fetchWithTimeout(`${baseUrl}/contact`)
    ]);
    extractEmails(homeHtml).forEach(e => emails.add(e));
    extractEmails(contactHtml).forEach(e => emails.add(e));
  } catch {
    // Ignore URL parse or fetch errors
  }

  return Array.from(emails);
}

/**
 * YouTube Data Scraper: Dual-Engine (InnerTube API + HTML Fallback with Consent Bypass)
 */
async function searchYouTubeChannels(query, limit = 50) {
  const channels = new Map();

  function parseSearchSections(sectionList) {
    for (const section of sectionList) {
      const items = section.itemSectionRenderer?.contents || [];
      for (const item of items) {
        // Channel search result
        if (item.channelRenderer) {
          const c = item.channelRenderer;
          const channelId = c.channelId;
          const handle = c.navigationEndpoint?.browseEndpoint?.canonicalBaseUrl || '';
          const title = c.title?.simpleText || '';
          const desc = c.descriptionSnippet?.runs?.map(r => r.text).join('') || '';
          const subs = c.videoCountText?.simpleText || c.subscriberCountText?.simpleText || '';

          if (channelId && !channels.has(channelId)) {
            channels.set(channelId, {
              channelId,
              handle,
              title,
              subscribers: subs,
              bioSnippet: desc
            });
          }
        }

        // Video search result (extract creator channel)
        if (item.videoRenderer) {
          const v = item.videoRenderer;
          const owner = v.ownerText?.runs?.[0];
          const channelEndpoint = owner?.navigationEndpoint?.browseEndpoint;
          const channelId = channelEndpoint?.browseId;
          const handle = channelEndpoint?.canonicalBaseUrl || '';
          const title = owner?.text || '';

          if (channelId && !channels.has(channelId)) {
            channels.set(channelId, {
              channelId,
              handle,
              title,
              subscribers: '',
              bioSnippet: ''
            });
          }
        }
      }
    }
  }

  // 1. Primary Engine: YouTube InnerTube API (Never blocked by HTML consent redirects!)
  async function executeInnerTubeSearch(params) {
    try {
      const payload = {
        context: {
          client: {
            clientName: "WEB",
            clientVersion: "2.20241101.01.00",
            gl: "US",
            hl: "en"
          }
        },
        query: query
      };
      if (params) payload.params = params;

      const res = await safeFetch('https://www.youtube.com/youtubei/v1/search?prettyPrint=false', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Cookie': 'SOCS=CAESEwgDEgk2OTg0NzgxMDUaAmVuIAEaBgiA_LyaBg; CONSENT=YES+cb; PREF=hl=en&gl=US;',
          'X-YouTube-Client-Name': '1',
          'X-YouTube-Client-Version': '2.20241101.01.00'
        },
        body: JSON.stringify(payload)
      });

      if (!res || !res.ok) return;
      const data = await res.json();
      const sectionList = data.contents?.twoColumnSearchResultsRenderer?.primaryContents?.sectionListRenderer?.contents || [];
      parseSearchSections(sectionList);
    } catch (err) {
      // InnerTube error, fallback will take over
    }
  }

  // 2. Secondary Engine: HTML Search with Cookie Consent bypass
  async function executeHtmlSearch(url) {
    try {
      const res = await safeFetch(url, {
        headers: {
          'Cookie': 'SOCS=CAESEwgDEgk2OTg0NzgxMDUaAmVuIAEaBgiA_LyaBg; CONSENT=YES+cb; PREF=hl=en&gl=US;',
          'Accept-Language': 'en-US,en;q=0.9'
        }
      });
      if (!res || !res.ok) {
        if (res && res.status === 429) {
          console.warn(`[Zenith] ⚠️ YouTube 429 Rate Limit hit. Backing off for 15s...`);
          await sleep(15000);
        }
        return;
      }
      const html = await res.text();
      const match = html.match(/var ytInitialData = ({.*?});<\/script>/s) || html.match(/ytInitialData\s*=\s*({.+?});/s);
      if (!match) return;

      const data = JSON.parse(match[1]);
      const sectionList = data.contents?.twoColumnSearchResultsRenderer?.primaryContents?.sectionListRenderer?.contents || [];
      parseSearchSections(sectionList);
    } catch (err) {
      console.warn(`[Zenith] HTML Search error for "${query}":`, err.message);
    }
  }

  // Step 1: Channels search via InnerTube API (EgIQAg%3D%3D = channels filter)
  await executeInnerTubeSearch("EgIQAg%3D%3D");

  // Step 2: Recent videos uploaded this month (CAISAhAB) if needed
  if (channels.size < limit) {
    await executeInnerTubeSearch("CAISAhAB");
  }

  // Step 3: General keyword search if needed
  if (channels.size < limit) {
    await executeInnerTubeSearch();
  }

  // Step 4: Fallback to HTML Search if InnerTube yielded nothing
  if (channels.size === 0) {
    const channelSearchUrl = `https://www.youtube.com/results?search_query=${encodeURIComponent(query)}&sp=EgIQAg%253D%253D&gl=US&hl=en`;
    await executeHtmlSearch(channelSearchUrl);
    if (channels.size < limit) {
      const recentVideosUrl = `https://www.youtube.com/results?search_query=${encodeURIComponent(query)}&sp=CAISAhAB&gl=US&hl=en`;
      await executeHtmlSearch(recentVideosUrl);
    }
  }

  return Array.from(channels.values()).slice(0, limit);
}

/**
 * Fetches deep channel details: bio, header links, social handles, recent videos
 */
async function fetchChannelDeepDetails(channel, options = {}) {
  const { scanVideoCount = 3, deepScanWebsite = true, requestDelayMs = 500 } = options;
  const channelUrl = getChannelUrl(channel.handle || channel.channelId);

  const lead = {
    channelId: channel.channelId,
    title: channel.title,
    handle: channel.handle || '',
    channelUrl: channelUrl,
    subscribers: channel.subscribers || '',
    subscribersNum: 0,
    videoCount: '',
    views: '',
    viewsNum: 0,
    country: '',
    joinedDate: '',
    bio: channel.bioSnippet || '',
    primaryEmail: '',
    allEmails: '',
    emailSource: '',
    hasEmail: false,
    hasCustomDomain: false,
    websiteUrl: '',
    hasSkool: false,
    skoolLink: '',
    instagram: '',
    twitter: '',
    linkedin: '',
    facebook: '',
    calendly: '',
    tags: '',
    recentVideoTitles: '',
    lastUploadText: '',
    lastUploadDays: 9999,
    scrapedAt: new Date().toISOString()
  };

  const foundEmails = new Map(); // email -> source
  const foundLinks = new Set();

  try {
    await sleep(requestDelayMs);
    const res = await safeFetch(channelUrl, {
      headers: {
        'Cookie': 'SOCS=CAESEwgDEgk2OTg0NzgxMDUaAmVuIAEaBgiA_LyaBg; CONSENT=YES+cb; PREF=hl=en&gl=US;',
        'Accept-Language': 'en-US,en;q=0.9'
      }
    });

    if (res && res.ok) {
      const html = await res.text();
      const match = html.match(/var ytInitialData = ({.*?});<\/script>/s) || html.match(/ytInitialData\s*=\s*({.+?});/s);

      if (match) {
        const data = JSON.parse(match[1]);

        // Channel microformat details
        const mf = data.microformat?.microformatDataRenderer;
        if (mf) {
          if (mf.title) lead.title = mf.title;
          if (mf.tags) lead.tags = mf.tags.join(', ');
          
          const profile = mf.channelProfileMicroformatDetails?.profilePage?.mainEntity;
          if (profile) {
            if (profile.description) lead.bio = profile.description.trim();
            if (profile.alternateName) lead.handle = profile.alternateName;
            if (profile.interactionStatistic?.userInteractionCount) {
              lead.subscribers = Number(profile.interactionStatistic.userInteractionCount).toLocaleString() + ' subscribers';
            }
            if (Array.isArray(profile.sameAs)) {
              profile.sameAs.forEach(u => foundLinks.add(u));
            }
          }
        }

        // Header viewModel metadata
        const headerVm = data.header?.pageHeaderRenderer?.content?.pageHeaderViewModel;
        if (headerVm) {
          // Description preview
          const descContent = headerVm.description?.descriptionPreviewViewModel?.description?.content;
          if (descContent && descContent.length > lead.bio.length) {
            lead.bio = descContent.trim();
          }

          // Attribution links (Website / Skool in header)
          const attr = headerVm.attribution?.attributionViewModel;
          if (attr?.text?.content) {
            foundLinks.add(attr.text.content);
          }
          const commandRuns = attr?.text?.commandRuns || [];
          for (const cr of commandRuns) {
            const u = cr.onTap?.innertubeCommand?.urlEndpoint?.url;
            if (u) foundLinks.add(u);
          }

          // Metadata rows (subscribers & videos)
          const rows = headerVm.metadata?.contentMetadataViewModel?.metadataRows || [];
          for (const row of rows) {
            for (const part of row.metadataParts || []) {
              const text = part.text?.content || '';
              if (text.includes('subscriber')) lead.subscribers = text;
              if (text.includes('video')) lead.videoCount = text;
            }
          }

          // Fetch deep About panel (Country, Total Views, Official Links)
          const ep = headerVm.description?.descriptionPreviewViewModel?.rendererContext?.commandContext?.onTap?.innertubeCommand?.showEngagementPanelEndpoint;
          const token = ep?.engagementPanel?.engagementPanelSectionListRenderer?.content?.sectionListRenderer?.contents?.[0]?.itemSectionRenderer?.contents?.[0]?.continuationItemRenderer?.continuationEndpoint?.continuationCommand?.token;

          if (token) {
            try {
              const bRes = await safeFetch('https://www.youtube.com/youtubei/v1/browse', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                  context: { client: { clientName: 'WEB', clientVersion: '2.20261112.10.00' } },
                  continuation: token
                })
              });
              if (bRes && bRes.ok) {
                const bData = await bRes.json();
                function findAbout(obj) {
                  if (!obj || typeof obj !== 'object') return null;
                  if (obj.aboutChannelViewModel) return obj.aboutChannelViewModel;
                  for (const k of Object.keys(obj)) {
                    const found = findAbout(obj[k]);
                    if (found) return found;
                  }
                  return null;
                }
                const about = findAbout(bData);
                if (about) {
                  if (about.country) lead.country = about.country;
                  if (about.viewCountText) lead.views = about.viewCountText;
                  if (about.subscriberCountText && !lead.subscribers) lead.subscribers = about.subscriberCountText;
                  if (about.joinedDateText?.content) lead.joinedDate = about.joinedDateText.content;
                  if (about.description && about.description.length > lead.bio.length) lead.bio = about.description;

                  if (Array.isArray(about.links)) {
                    for (const lk of about.links) {
                      const extLink = lk.channelExternalLinkViewModel;
                      const linkContent = extLink?.link?.content;
                      if (linkContent) foundLinks.add(linkContent);
                      const redirectUrl = extLink?.link?.commandRuns?.[0]?.onTap?.innertubeCommand?.urlEndpoint?.url;
                      if (redirectUrl) {
                        const qMatch = redirectUrl.match(/[?&]q=([^&]+)/);
                        if (qMatch) foundLinks.add(decodeURIComponent(qMatch[1]));
                      }
                    }
                  }
                }
              }
            } catch {
              // Proceed gracefully if continuation fails
            }
          }
        }
      }

      // Check for emails in channel bio
      const bioEmails = extractEmails(lead.bio);
      bioEmails.forEach(e => foundEmails.set(e, 'Channel Bio'));

      // Check links found in channel page & bio
      const bioLinkInfo = extractLinks(lead.bio + ' ' + Array.from(foundLinks).join(' '));
      if (bioLinkInfo.skool) {
        lead.hasSkool = true;
        lead.skoolLink = bioLinkInfo.skool;
      }
      if (bioLinkInfo.website) {
        lead.hasCustomDomain = true;
        lead.websiteUrl = bioLinkInfo.website;
      }
      if (bioLinkInfo.socials.instagram) lead.instagram = bioLinkInfo.socials.instagram;
      if (bioLinkInfo.socials.twitter) lead.twitter = bioLinkInfo.socials.twitter;
      if (bioLinkInfo.socials.linkedin) lead.linkedin = bioLinkInfo.socials.linkedin;
      if (bioLinkInfo.socials.facebook) lead.facebook = bioLinkInfo.socials.facebook;
      if (bioLinkInfo.socials.calendly) lead.calendly = bioLinkInfo.socials.calendly;
    }

    // Compute preliminary metrics
    lead.subscribersNum = parseNumericCount(lead.subscribers);
    lead.viewsNum = parseNumericCount(lead.views);

    // ⚡ HYPER-SPEED FAST-FAIL GATE:
    // If channel already fails sub count or is from banned 3rd world country,
    // skip slow video downloads and external website crawling immediately (0.05s instead of 5s)!
    const filters = options.filters || {};
    if (filters.minSubscribers && filters.minSubscribers > 0 && lead.subscribersNum > 0 && lead.subscribersNum < filters.minSubscribers) {
      return lead;
    }
    if (filters.maxSubscribers && filters.maxSubscribers > 0 && lead.subscribersNum > filters.maxSubscribers) {
      return lead;
    }
    if (filters.excludeThirdWorld && lead.country) {
      const countryLower = lead.country.toLowerCase().trim();
      if (THIRD_WORLD_COUNTRIES.includes(countryLower)) {
        return lead;
      }
    }

    // 2. Scan recent videos for email in video description
    const targetHandle = lead.handle || channel.handle;
    if (scanVideoCount > 0 && targetHandle) {
      await sleep(requestDelayMs / 2);
      const videosUrl = `${getChannelUrl(targetHandle)}/videos`;
      const vRes = await safeFetch(videosUrl, {
        headers: {
          'Cookie': 'SOCS=CAESEwgDEgk2OTg0NzgxMDUaAmVuIAEaBgiA_LyaBg; CONSENT=YES+cb; PREF=hl=en&gl=US;',
          'Accept-Language': 'en-US,en;q=0.9'
        }
      });

      if (vRes && vRes.ok) {
        const vHtml = await vRes.text();
        const vMatch = vHtml.match(/var ytInitialData = ({.*?});<\/script>/s) || vHtml.match(/ytInitialData\s*=\s*({.+?});/s);
        if (vMatch) {
          const vData = JSON.parse(vMatch[1]);
          const tabContents = vData.contents?.twoColumnBrowseResultsRenderer?.tabs?.find(t => t.tabRenderer?.title === 'Videos' || t.tabRenderer?.selected)?.tabRenderer?.content?.richGridRenderer?.contents || [];
          
          const recentVideoIds = [];
          const videoTitles = [];

          // Scan first 5 items to detect upload recency (handles premiere, live stream, or regular video)
          for (const item of tabContents.slice(0, 5)) {
            const vm = item?.richItemRenderer?.content?.lockupViewModel;
            let dateText = '';
            if (vm) {
              const rows = vm.metadata?.lockupMetadataViewModel?.metadata?.contentMetadataViewModel?.metadataRows || [];
              for (const r of rows) {
                for (const p of r.metadataParts || []) {
                  const txt = p.text?.content || '';
                  if (txt.includes('ago') || txt.includes('Stream') || txt.includes('Premier') || txt.includes('Live')) {
                    dateText = txt;
                    break;
                  }
                }
                if (dateText) break;
              }
            } else {
              const vr = item?.richItemRenderer?.content?.videoRenderer;
              dateText = vr?.publishedTimeText?.simpleText || '';
            }

            if (dateText && !lead.lastUploadText) {
              lead.lastUploadText = dateText;
              lead.lastUploadDays = parseDaysAgo(dateText);
              break;
            }
          }

          for (const item of tabContents.slice(0, scanVideoCount)) {
            const vm = item.richItemRenderer?.content?.lockupViewModel;
            const videoId = vm?.contentImage?.thumbnailBottomOverlayViewModel?.badges?.[0]?.thumbnailBadgeViewModel?.animationActivationTargetId
              || vm?.commandContext?.onTap?.innertubeCommand?.watchEndpoint?.videoId;
            const title = vm?.rendererContext?.accessibilityContext?.label || '';

            if (videoId) {
              recentVideoIds.push(videoId);
              if (title) videoTitles.push(title.replace(/\s+\d+.*$/, ''));
            }
          }

          lead.recentVideoTitles = videoTitles.join(' | ');

          // Fetch descriptions of up to N recent videos in parallel
          const videoFetches = recentVideoIds.slice(0, scanVideoCount).map(async (vid) => {
            const watchUrl = `https://www.youtube.com/watch?v=${vid}`;
            const watchRes = await safeFetch(watchUrl, {
              headers: { 'Accept-Language': 'en-US,en;q=0.9' }
            });
            if (watchRes && watchRes.ok) {
              const watchHtml = await watchRes.text();
              const pMatch = watchHtml.match(/var ytInitialPlayerResponse\s*=\s*(\{.+?\});(?:var|<\/script>)/s);
              if (pMatch) {
                try {
                  const pData = JSON.parse(pMatch[1]);
                  const videoDesc = pData.videoDetails?.shortDescription || '';
                  
                  // Extract emails from video description
                  const vidEmails = extractEmails(videoDesc);
                  vidEmails.forEach(e => {
                    if (!foundEmails.has(e)) foundEmails.set(e, `Video (${vid}) Description`);
                  });

                  // Extract links (Skool, Website) from video description
                  const vLinks = extractLinks(videoDesc);
                  if (vLinks.skool && !lead.skoolLink) {
                    lead.hasSkool = true;
                    lead.skoolLink = vLinks.skool;
                  }
                  if (vLinks.website && !lead.websiteUrl) {
                    lead.hasCustomDomain = true;
                    lead.websiteUrl = vLinks.website;
                  }
                  if (vLinks.socials.calendly && !lead.calendly) lead.calendly = vLinks.socials.calendly;
                  if (vLinks.socials.instagram && !lead.instagram) lead.instagram = vLinks.socials.instagram;
                  if (vLinks.socials.twitter && !lead.twitter) lead.twitter = vLinks.socials.twitter;
                  if (vLinks.socials.linkedin && !lead.linkedin) lead.linkedin = vLinks.socials.linkedin;
                } catch {
                  // Ignore JSON parse error on player
                }
              }
            }
          });
          await Promise.all(videoFetches);
        }
      }
    }

    // 3. Deep Domain Website Scanning for emails (if website found and emails still missing)
    if (deepScanWebsite && lead.websiteUrl && foundEmails.size === 0) {
      const siteEmails = await scrapeDomainForEmails(lead.websiteUrl, 4000);
      siteEmails.forEach(e => {
        if (!foundEmails.has(e)) foundEmails.set(e, 'Website Homepage/Contact');
      });
    }

    // Compute numeric counts for filtering
    lead.subscribersNum = parseNumericCount(lead.subscribers);
    lead.viewsNum = parseNumericCount(lead.views);

    // Consolidate extracted emails
    if (foundEmails.size > 0) {
      const emailList = Array.from(foundEmails.keys());
      lead.primaryEmail = emailList[0];
      lead.allEmails = emailList.join(', ');
      lead.emailSource = Array.from(foundEmails.entries()).map(([e, s]) => `${e} (${s})`).join('; ');
      lead.hasEmail = true;
    }

  } catch (err) {
    console.warn(`[Zenith] Channel inspection failed for ${channel.title}:`, err.message);
  }

  return lead;
}

/**
 * Check if channel matches target countries (e.g. USA, UK, Australia, Austria)
 */
function matchesTargetCountries(lead, targetCountries = []) {
  if (!targetCountries || targetCountries.length === 0) return true;
  const targets = targetCountries.map(c => c.toLowerCase().trim()).filter(Boolean);
  const countryLower = (lead.country || '').toLowerCase().trim();
  const bioLower = (lead.bio || '').toLowerCase();
  const webLower = (lead.websiteUrl || '').toLowerCase();

  // 1. Direct match on official YouTube channel country
  if (countryLower) {
    for (const t of targets) {
      if (countryLower === t || countryLower.includes(t) || t.includes(countryLower)) return true;
      if ((t === 'usa' || t === 'us') && (countryLower === 'united states' || countryLower === 'usa')) return true;
      if ((t === 'uk' || t === 'gb') && (countryLower === 'united kingdom' || countryLower === 'great britain')) return true;
      if ((t === 'au') && countryLower === 'australia') return true;
      if ((t === 'at') && countryLower === 'austria') return true;
    }
  }

  // 2. Infer country from bio or domain TLD if country wasn't explicitly set on YouTube
  for (const t of targets) {
    if (t === 'usa' || t === 'us' || t === 'united states') {
      if (
        bioLower.includes('usa') ||
        bioLower.includes('united states') ||
        /\b(california|texas|new york|florida|chicago|los angeles|miami|austin|seattle|atlanta)\b/i.test(bioLower) ||
        webLower.endsWith('.us')
      ) {
        if (!lead.country) lead.country = 'United States';
        return true;
      }
    } else if (t === 'uk' || t === 'united kingdom' || t === 'great britain') {
      if (
        bioLower.includes('united kingdom') ||
        bioLower.includes('london') ||
        bioLower.includes('england') ||
        bioLower.includes('scotland') ||
        /\buk\b/i.test(bioLower) ||
        webLower.includes('.co.uk') ||
        webLower.endsWith('.uk')
      ) {
        if (!lead.country) lead.country = 'United Kingdom';
        return true;
      }
    } else if (t === 'austria' || t === 'at') {
      if (
        bioLower.includes('austria') ||
        bioLower.includes('vienna') ||
        bioLower.includes('wien') ||
        bioLower.includes('österreich') ||
        webLower.endsWith('.at')
      ) {
        if (!lead.country) lead.country = 'Austria';
        return true;
      }
    } else if (t === 'australia' || t === 'au') {
      if (
        bioLower.includes('australia') ||
        bioLower.includes('sydney') ||
        bioLower.includes('melbourne') ||
        bioLower.includes('brisbane') ||
        webLower.endsWith('.au') ||
        webLower.includes('.com.au')
      ) {
        if (!lead.country) lead.country = 'Australia';
        return true;
      }
    } else {
      if (bioLower.includes(t) || webLower.includes(t)) {
        if (!lead.country) lead.country = t;
        return true;
      }
    }
  }

  return false;
}

/**
 * Filter evaluation based on user preferences
 */
function passesFilters(lead, filters = {}) {
  // 1. Strict 3rd World Exclusion
  if (filters.excludeThirdWorld) {
    const countryLower = (lead.country || '').toLowerCase().trim();
    if (THIRD_WORLD_COUNTRIES.includes(countryLower)) {
      return false;
    }

    // Check website domain TLD for 3rd world domains (.in, .pk, .bd, .ng, etc.)
    if (lead.websiteUrl) {
      try {
        const hostname = new URL(lead.websiteUrl).hostname.toLowerCase();
        if (THIRD_WORLD_TLDS.some(tld => hostname.endsWith(tld))) {
          return false;
        }
      } catch {}
    }

    // Check bio for third-world phone prefixes (+91, +880, +92, etc.)
    const hasThirdWorldPhone = THIRD_WORLD_PHONE_PREFIXES.some(prefix => (lead.bio || '').includes(prefix));
    if (hasThirdWorldPhone) return false;
  }

  // 2. Specific Target Countries Filter (e.g. USA, UK, Australia, Austria)
  if (filters.targetCountries && filters.targetCountries.length > 0) {
    if (!matchesTargetCountries(lead, filters.targetCountries)) {
      return false;
    }
  }

  // 3. Tier-1 Only Mode (strictly USA, UK, Canada, Australia, NZ, Europe Tier-1)
  if (filters.tier1Only) {
    const countryLower = (lead.country || '').toLowerCase().trim();
    if (!countryLower || !TIER_1_COUNTRIES.some(c => countryLower.includes(c))) {
      return false;
    }
  }

  // 3. Minimum & Maximum Subscribers Filter
  if (filters.minSubscribers && filters.minSubscribers > 0) {
    if (lead.subscribersNum < filters.minSubscribers) return false;
  }
  if (filters.maxSubscribers && filters.maxSubscribers > 0) {
    if (lead.subscribersNum > filters.maxSubscribers) return false;
  }

  // 4. Minimum & Maximum Views Filter
  if (filters.minViews && filters.minViews > 0) {
    if (lead.viewsNum < filters.minViews) return false;
  }
  if (filters.maxViews && filters.maxViews > 0) {
    if (lead.viewsNum > filters.maxViews) return false;
  }

  // 5. Upload Recency Filter (e.g. uploaded within last 7, 30, 60, 90, 180 days)
  if (filters.maxUploadAgeDays && filters.maxUploadAgeDays > 0) {
    if (lead.lastUploadDays > filters.maxUploadAgeDays) {
      return false;
    }
  }

  // 6. Email Filter: require email or keep all
  if (filters.requireEmail && !lead.hasEmail) {
    return false;
  }

  // 6. Domain Filter: must have custom website
  if (filters.requireDomain && !lead.hasCustomDomain) {
    return false;
  }

  // 7. Skool Filter: must have Skool community
  if (filters.requireSkool && !lead.hasSkool) {
    return false;
  }

  // 8. Business keyword heuristic (optional)
  if (filters.businessKeywords && filters.businessKeywords.length > 0) {
    const textCorpus = `${lead.title} ${lead.bio} ${lead.tags} ${lead.recentVideoTitles}`.toLowerCase();
    const hasBusinessKeyword = filters.businessKeywords.some(kw => textCorpus.includes(kw.toLowerCase()));
    if (!hasBusinessKeyword) return false;
  }

  return true;
}

/**
 * Export results to CSV, Excel, and JSON
 */
function exportResults(leads, outputBasePath) {
  const dir = path.dirname(outputBasePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  // 1. Export JSON
  fs.writeFileSync(`${outputBasePath}.json`, JSON.stringify(leads, null, 2), 'utf-8');

  // 2. Export CSV
  if (leads.length > 0) {
    const headers = Object.keys(leads[0]);
    const escapeCsv = (val) => {
      if (val === null || val === undefined) return '""';
      const str = String(val).replace(/"/g, '""').replace(/\r?\n/g, ' ');
      return `"${str}"`;
    };
    const csvContent = [
      headers.map(escapeCsv).join(','),
      ...leads.map(row => headers.map(h => escapeCsv(row[h])).join(','))
    ].join('\n');
    fs.writeFileSync(`${outputBasePath}.csv`, csvContent, 'utf-8');
  } else {
    fs.writeFileSync(`${outputBasePath}.csv`, '', 'utf-8');
  }

  // 3. Export XLSX
  const worksheet = xlsx.utils.json_to_sheet(leads);
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, 'Zenith Leads');
  xlsx.writeFile(workbook, `${outputBasePath}.xlsx`);

  console.log(`[Zenith] Exported ${leads.length} leads to:`);
  console.log(`  - ${outputBasePath}.csv`);
  console.log(`  - ${outputBasePath}.xlsx`);
  console.log(`  - ${outputBasePath}.json`);
}

/**
 * Exhaustive Query Expansion Engine
 * Generates hundreds/thousands of laser-focused permutations for any niche
 */
function generateExpandedQueryMatrix(baseNiches = ['coach'], targetCountries = []) {
  const defaultGeos = [
    'USA', 'US', 'United States', 'New York', 'California', 'Texas', 'Florida', 'Austin',
    'Miami', 'Chicago', 'Los Angeles', 'Atlanta', 'Dallas', 'Seattle', 'Denver', 'Boston',
    'UK', 'United Kingdom', 'London', 'Manchester', 'Birmingham', 'Leeds', 'Bristol', 'Edinburgh',
    'Australia', 'Sydney', 'Melbourne', 'Brisbane', 'Perth', 'Adelaide',
    'Austria', 'Vienna', 'Wien', 'Salzburg'
  ];

  const customGeos = (targetCountries && targetCountries.length > 0)
    ? [...new Set([...targetCountries, ...defaultGeos])]
    : defaultGeos;

  const coachingRoles = [
    // General Coaching
    'coach', 'coaching', 'mentor', 'mentorship', 'consultant',
    // Life & Mindset
    'life coach', 'life coaching', 'mindset coach', 'transformation coach',
    'mental performance coach', 'success coach', 'personal development coach',
    'spiritual coach', 'empowerment coach', 'confidence coach',
    // Fitness, Health & Wellness
    'fitness coach', 'online fitness coach', 'health coach', 'nutrition coach',
    'wellness coach', 'weight loss coach', 'strength coach', 'bodybuilding coach',
    // Career, Leadership & Executive
    'career coach', 'executive coach', 'leadership coach', 'interview coach',
    'management coach', 'public speaking coach', 'communication coach',
    // Relationship & Dating
    'relationship coach', 'dating coach', 'marriage coach', 'couples coach',
    // Finance & Wealth
    'financial coach', 'money coach', 'wealth coach', 'real estate coach',
    'crypto coach', 'trading coach',
    // Business, Sales & Marketing
    'business coach', 'business coaching', 'business consultant', 'sales coach',
    'marketing coach', 'scaling coach', 'agency coach', 'ecommerce coach',
    'high ticket coach', 'b2b coach', 'growth consultant',
    // Skill & Content
    'content coach', 'youtube coach', 'vocal coach', 'voice coach'
  ];

  const topics = [
    'coaching client', 'coaching program', 'transformation', 'mindset',
    'habit building', 'mastermind', 'podcast', 'interview', 'q&a',
    'tips & strategies', 'framework & roadmap', 'case study', 'advice',
    'client acquisition', 'scaling to 6 figures', 'scaling to 7 figures',
    'high ticket sales', 'lead generation', 'systems & routines',
    'for beginners', 'for entrepreneurs', 'for creators', 'for professionals'
  ];

  const queries = new Set();

  for (const rawNiche of baseNiches) {
    const cleanNiche = rawNiche.trim();
    if (!cleanNiche) continue;

    const isCoaching = cleanNiche.toLowerCase().includes('coach') || cleanNiche.toLowerCase().includes('consult') || cleanNiche.toLowerCase().includes('mentor');
    const rolesToUse = isCoaching ? coachingRoles : [cleanNiche];

    for (const role of rolesToUse) {
      queries.add(role);

      // Role + Geo
      for (const geo of customGeos) {
        queries.add(`${role} ${geo}`);
      }

      // Role + Topic
      for (const topic of topics) {
        queries.add(`${role} ${topic}`);
        queries.add(`${topic} ${role}`);
      }
    }

    // Role + Topic + Geo permutations
    for (const role of rolesToUse.slice(0, 6)) {
      for (const topic of topics.slice(0, 8)) {
        for (const geo of ['USA', 'UK', 'Australia', 'Austria', 'Austin', 'London', 'Sydney']) {
          queries.add(`${role} ${topic} ${geo}`);
        }
      }
    }
  }

  return Array.from(queries);
}

/**
 * Main Run Function for a Worker
 */
async function runWorker(config) {
  const {
    workerId = 1,
    totalWorkers = 1,
    niches = ['coach'],
    maxChannelsPerWorker = 30,
    filters = {
      requireEmail: false,
      requireDomain: false,
      requireSkool: false,
      businessKeywords: []
    },
    scrapingOptions = {
      scanVideoCount: 3,
      deepScanWebsite: true,
      requestDelayMs: 600
    },
    outputDir = './output',
    onProgress = null
  } = config;

  console.log(`====================================================`);
  console.log(`  🚀 ZENITH YOUTUBE SCRAPER: WORKER #${workerId} / ${totalWorkers}`);
  console.log(`====================================================`);
  console.log(`Target Niches:`, niches.join(' | '));
  console.log(`Max Channels per Worker (Goal):`, maxChannelsPerWorker);
  console.log(`Filters:`, JSON.stringify(filters));
  console.log(`Speed Delay:`, `${scrapingOptions.requestDelayMs}ms`);
  console.log(`Deep Website Scan:`, scrapingOptions.deepScanWebsite);
  console.log(`----------------------------------------------------`);

  // 1. Generate full query pool (hundreds or thousands of unique queries)
  const allQueries = generateExpandedQueryMatrix(niches, filters.targetCountries);

  // 2. Distribute non-overlapping queries to this worker
  const activeNiches = allQueries.filter((_, idx) => (idx % totalWorkers) === (workerId - 1));

  console.log(`Worker #${workerId} assigned ${activeNiches.length} unique search queries (out of ${allQueries.length} total permutations).`);
  console.log(`Sample first 5 queries:`, activeNiches.slice(0, 5).join(' | '));

  const outPath = path.join(outputDir, `zenith_worker_${workerId}`);
  const processedLeads = [];
  const seenChannelIds = new Set();
  let totalInspected = 0;

  // Load persistent history of previously scraped channels across all runs
  const historyFilePath = path.join(__dirname, 'data', 'seen_channels.txt');
  if (fs.existsSync(historyFilePath)) {
    try {
      const savedLines = fs.readFileSync(historyFilePath, 'utf-8').split(/\r?\n/).map(s => s.trim()).filter(Boolean);
      savedLines.forEach(id => seenChannelIds.add(id));
      console.log(`[Worker #${workerId}] 🛡️ Loaded ${seenChannelIds.size} previously scraped channel IDs from history to prevent ANY duplicate.`);
    } catch (err) {
      console.warn(`[Worker #${workerId}] Could not load history file:`, err.message);
    }
  }

  // CONTINUOUS INFINITE STREAMING CRAWL: Grinds through queries until TARGET IS HIT!
  const queryQueue = [...activeNiches];
  let round = 1;

  while (processedLeads.length < maxChannelsPerWorker) {
    if (queryQueue.length === 0) {
      round++;
      console.log(`\n🔄 [Worker ${workerId}] Generating Round ${round} fresh queries to reach target (${processedLeads.length}/${maxChannelsPerWorker} leads)...`);
      
      const subIndustries = [
        'life', 'mindset', 'fitness', 'health', 'nutrition', 'wellness',
        'weight loss', 'strength', 'career', 'executive', 'leadership',
        'relationship', 'dating', 'marriage', 'financial', 'money',
        'wealth', 'sales', 'business', 'real estate', 'marketing',
        'agency', 'ecommerce', 'high ticket', 'transformation', 'success',
        'public speaking', 'performance', 'spiritual', 'consulting'
      ];
      const cities = [
        'New York', 'London', 'Sydney', 'Austin', 'Miami', 'Chicago',
        'Los Angeles', 'Atlanta', 'Dallas', 'Manchester', 'Melbourne',
        'Brisbane', 'Vienna', 'Denver', 'Seattle', 'Boston', 'Phoenix',
        'Houston', 'San Francisco', 'Birmingham', 'Leeds'
      ];
      const topics = [
        'mindset', 'transformation', 'habits', 'client acquisition', 'sales',
        'high ticket', 'operations', 'systems', 'roadmap', 'interview', 'podcast', 'tips'
      ];

      const newBatch = [];
      // 1. Industry + Role (e.g. "life coach", "fitness coach", "mindset coach", "career coach")
      for (const ind of subIndustries) {
        newBatch.push(`${ind} coach`);
        newBatch.push(`${ind} coaching`);
        newBatch.push(`${ind} mentor`);
      }
      // 2. Role + City (e.g. "life coach London", "fitness coach Austin", "career coach Sydney")
      for (const city of cities) {
        for (const ind of ['life', 'fitness', 'mindset', 'career', 'business', 'health', 'executive']) {
          newBatch.push(`${ind} coach ${city}`);
        }
      }
      // 3. Topic + Role
      for (const top of topics) {
        newBatch.push(`${top} coach`);
        newBatch.push(`coaching ${top}`);
      }
      // 4. Industry + Topic + Coach
      for (const ind of subIndustries.slice(0, 15)) {
        for (const top of topics.slice(0, 4)) {
          newBatch.push(`${ind} ${top} coach`);
        }
      }

      // Distribute modulo across workers
      const workerNewBatch = newBatch.filter((_, idx) => (idx % totalWorkers) === (workerId - 1));
      queryQueue.push(...workerNewBatch);

      if (queryQueue.length === 0) {
        // Fallback: alphabet expansion
        const alphabet = 'abcdefghijklmnopqrstuvwxyz'.split('');
        for (const char of alphabet) {
          queryQueue.push(`coach ${char}`);
          queryQueue.push(`coaching ${char}`);
        }
      }
    }

    const currentQuery = queryQueue.shift();
    if (!currentQuery) break;

    console.log(`\n====================================================`);
    console.log(`[Worker ${workerId}] Query: "${currentQuery}"`);
    console.log(`Current Progress: ${processedLeads.length}/${maxChannelsPerWorker} valid leads | Inspected: ${totalInspected}`);
    console.log(`====================================================`);

    // Fetch candidates for this query
    let candidates = [];
    try {
      candidates = await searchYouTubeChannels(currentQuery, 35);
    } catch (err) {
      console.warn(`[Worker ${workerId}] Search failed for "${currentQuery}": ${err.message}`);
      await sleep(1000);
      continue;
    }

    console.log(`[Worker ${workerId}] Found ${candidates.length} candidates from search.`);

    // Filter unseen candidates
    const unseen = [];
    for (const raw of candidates) {
      if (!seenChannelIds.has(raw.channelId)) {
        seenChannelIds.add(raw.channelId);
        unseen.push(raw);
      }
    }

    // Process in concurrent batches of 6 channels for 6x hyper-speed
    const BATCH_SIZE = 6;
    for (let i = 0; i < unseen.length; i += BATCH_SIZE) {
      if (processedLeads.length >= maxChannelsPerWorker) break;

      const batch = unseen.slice(i, i + BATCH_SIZE);
      const batchResults = await Promise.all(batch.map(async (raw) => {
        try {
          const inspected = await Promise.race([
            fetchChannelDeepDetails(raw, { ...scrapingOptions, filters }),
            new Promise((_, reject) => setTimeout(() => reject(new Error('Inspection Timeout')), 12000))
          ]);
          if (inspected) {
            inspected.workerId = workerId;
            inspected.targetNiche = currentQuery;
          }
          return inspected;
        } catch {
          return null;
        }
      }));

      for (const lead of batchResults) {
        if (!lead) continue;
        totalInspected++;
        if (passesFilters(lead, filters)) {
          processedLeads.push(lead);
          const emailStatus = lead.hasEmail ? `✅ EMAIL: ${lead.primaryEmail}` : '❌ No Email';
          const domainStatus = lead.hasCustomDomain ? `🌐 ${lead.websiteUrl}` : 'No Website';
          const skoolStatus = lead.hasSkool ? `🎒 Skool: ${lead.skoolLink}` : '';
          const countryTag = lead.country ? `🌍 ${lead.country}` : '🌍 Target Geo';
          const subsTag = `👥 Subs: ${lead.subscribers || lead.subscribersNum || 'N/A'}`;
          const viewsTag = `👁️ Views: ${lead.views || lead.viewsNum || 'N/A'}`;
          const uploadTag = `📅 Upload: ${lead.lastUploadText || 'N/A'}`;
          console.log(`   🎯 [${processedLeads.length}/${maxChannelsPerWorker}] MATCH! ${countryTag} | ${subsTag} | ${viewsTag} | ${uploadTag} | ${emailStatus} | ${domainStatus} ${skoolStatus}`);

          // Incremental save every 3 leads so no work is ever lost
          if (processedLeads.length % 3 === 0 || processedLeads.length === maxChannelsPerWorker) {
            exportResults(processedLeads, outPath);
          }
        } else {
          console.log(`   └─ Filtered out (${lead.title}: Country: "${lead.country || 'Unknown'}", Subs: ${lead.subscribersNum}, Last Upload: "${lead.lastUploadText || 'None'}")`);
        }

        if (onProgress) {
          onProgress({
            workerId,
            current: totalInspected,
            leadCount: processedLeads.length,
            latestLead: lead
          });
        }
      }

      await sleep(Math.max(100, scrapingOptions.requestDelayMs / 2));
    }

    await sleep(scrapingOptions.requestDelayMs);
  }

  console.log(`\n🎯🎯 [Worker ${workerId}] TARGET COMPLETED! Reached ${processedLeads.length}/${maxChannelsPerWorker} matching leads.`);

  // Final export of results
  exportResults(processedLeads, outPath);

  const stats = {
    workerId,
    totalInspected,
    totalMatched: processedLeads.length,
    withEmails: processedLeads.filter(l => l.hasEmail).length,
    withWebsites: processedLeads.filter(l => l.hasCustomDomain).length,
    withSkool: processedLeads.filter(l => l.hasSkool).length
  };

  console.log(`====================================================`);
  console.log(`Worker #${workerId} Finished:`);
  console.log(`  - Total Inspected: ${stats.totalInspected}`);
  console.log(`  - Total Saved:     ${stats.totalMatched}`);
  console.log(`  - With Emails:     ${stats.withEmails}`);
  console.log(`  - With Websites:   ${stats.withWebsites}`);
  console.log(`  - With Skool:      ${stats.withSkool}`);
  console.log(`====================================================`);

  return { leads: processedLeads, stats };
}

module.exports = {
  runWorker,
  searchYouTubeChannels,
  fetchChannelDeepDetails,
  extractEmails,
  extractLinks,
  scrapeDomainForEmails,
  exportResults,
  passesFilters,
  generateExpandedQueryMatrix
};
