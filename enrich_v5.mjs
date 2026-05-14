#!/usr/bin/env node
/**
 * Website Enrichment V5 - Fast Node.js version using z-ai SDK directly.
 * Processes companies in batches with parallel search + sequential download.
 */
import ZAI from 'z-ai-web-dev-sdk';
import { readFileSync, writeFileSync, existsSync, mkdirSync, appendFileSync } from 'fs';
import { join } from 'path';
import https from 'https';
import http from 'http';

const RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json';
const ENRICH_FILE = '/home/z/my-project/website_enrichment.json';
const HTML_DIR = '/home/z/my-project/download/website_html';
const LOG_FILE = '/home/z/my-project/enrich_v5_log.txt';
const TARGET_CITIES = ['Bamberg', 'Erlangen', 'Nürnberg'];

const SKIP_DOMAINS = [
    'facebook.com','instagram.com','twitter.com','linkedin.com','youtube.com',
    'wikipedia.org','maps.google','gelbeseiten.de','yelp.de','tripadvisor',
    'firmenwissen.de','northdata','kompass','wlw.de','cylex.de','hotfrog',
    'cybo.com','dascleverle','kundeu.com','mojomox','opening-hours',
    'unternehmensregister.de','indofolio','bizdb','de.lusha','checkfacebook',
    'booking.com','11880.com','dasoertliche.de','check24.de','bloomberg.com',
    'crunchbase.com','glassdoor.com','kununu.com','indeed.com','stepstone.de',
    'meisterkarte.de','handwerksuche','diebestenderstadt.de','google.com',
    'provenexpert.com','dastelefonbuch.de','woobi.de','cylex-branchenbuch',
    'branchenbuch24.net','webwiki.de','mittelstandswiki.de','jevee.de',
    'bayerischewirtschaft.de','ratemyarea.com','tellows.de',
    'wlb.de','lokal.blue','meinestadt.de','hellowork.com',
    'stellenonline.de','jobware.de','arbeitsagentur.de','azubiyo.de',
    'ausbildung.de','gebaeudereiniger-portal.de','beeradvocate.com',
    'ratebeer.com','untappd.com','hopfenherz.de','bierland-franken.de',
    'frankentourismus.de','brauerei-map.de',
];

function log(msg) {
    const ts = new Date().toLocaleTimeString('de-DE', {hour12: false});
    console.log(`[${ts}] ${msg}`);
    try { 
        const line = `[${ts}] ${msg}\n`;
        if (existsSync(LOG_FILE)) {
            appendFileSync(LOG_FILE, line);
        } else {
            writeFileSync(LOG_FILE, line);
        }
    } catch(e) {}
}

function safeFilename(name) {
    return name.replace(/[^\w\s\-.]/g, '').replace(/\s+/g, '_').trim().slice(0, 80);
}

function isSkipDomain(url) {
    const lower = url.toLowerCase();
    return SKIP_DOMAINS.some(d => lower.includes(d));
}

async function searchWebsite(zai, name, city) {
    const query = `${name} ${city}`;
    try {
        const results = await zai.functions.invoke('web_search', { query, num: 5 });
        for (const item of results) {
            const url = item.url || '';
            if (!url.startsWith('http')) continue;
            if (isSkipDomain(url) || isSkipDomain(item.host_name || '')) continue;
            return url;
        }
        // Try quoted query if all filtered
        const query2 = `"${name}" ${city}`;
        const results2 = await zai.functions.invoke('web_search', { query: query2, num: 3 });
        for (const item of results2) {
            const url = item.url || '';
            if (!url.startsWith('http')) continue;
            if (isSkipDomain(url) || isSkipDomain(item.host_name || '')) continue;
            return url;
        }
    } catch (e) {
        // ignore
    }
    return null;
}

function httpGet(url, timeout = 8000) {
    return new Promise((resolve) => {
        const proto = url.startsWith('https') ? https : http;
        const req = proto.get(url, {
            headers: {
                'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64; rv:109.0) Gecko/20100101 Firefox/115.0',
                'Accept': 'text/html,application/xhtml+xml,*/*',
                'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
            },
            timeout,
        }, (res) => {
            // Handle redirects
            if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
                const loc = res.headers.location;
                const newUrl = loc.startsWith('http') ? loc : new URL(loc, url).href;
                httpGet(newUrl, timeout).then(resolve);
                return;
            }
            const chunks = [];
            let size = 0;
            res.on('data', (chunk) => {
                size += chunk.length;
                if (size <= 500000) chunks.push(chunk);
            });
            res.on('end', () => {
                const data = Buffer.concat(chunks);
                const ct = res.headers['content-type'] || '';
                resolve({ data, ct, finalUrl: res.responseUrl || url, size });
            });
        });
        req.on('error', () => resolve(null));
        req.on('timeout', () => { req.destroy(); resolve(null); });
    });
}

function extractEmails(html) {
    const text = typeof html === 'string' ? html : html.toString('utf-8');
    const emails = new Set();
    
    const skipInEmail = ['.png','.jpg','.gif','.svg','.css','.js','example.com','email.com','domain.com','sentry','wixpress','googlemail'];
    const skipEmails = ['sage','wordpress','admin@localhost','test@','noreply@','example@','webmaster@','postmaster@','donotreply@'];
    
    const regex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
    let m;
    while ((m = regex.exec(text)) !== null) {
        const e = m[0].toLowerCase();
        if (!skipInEmail.some(x => e.includes(x))) emails.add(e);
    }
    
    const mailtoRegex = /mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/gi;
    while ((m = mailtoRegex.exec(text)) !== null) {
        emails.add(m[1].toLowerCase());
    }
    
    const atRegex = /[a-zA-Z0-9._%+-]+\s*[\(\[]at[\)\]]\s*[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi;
    while ((m = atRegex.exec(text)) !== null) {
        emails.add(m[0].toLowerCase().replace(/\s*[\(\[]at[\)\]]\s*/, '@'));
    }
    
    let filtered = [...emails].filter(e => !skipEmails.some(s => e.includes(s)));
    
    // Prefer specific types
    const pref1 = filtered.filter(e => /bewerb|hr|recruit|karrier|job|career/.test(e));
    if (pref1.length) return pref1[0];
    const pref2 = filtered.filter(e => /info|kontakt|contact/.test(e));
    if (pref2.length) return pref2[0];
    filtered.sort();
    return filtered[0] || '';
}

async function main() {
    const startIdx = parseInt(process.argv[2] || '0');
    const batchSize = parseInt(process.argv[3] || '25');
    
    const ts = new Date().toLocaleTimeString('de-DE', {hour12:false});
    console.log(`[${ts}] ${'='.repeat(50)}`);
    console.log(`[${ts}] Website Enrichment V5 | start=${startIdx} batch=${batchSize}`);
    
    // Load data
    const allResults = JSON.parse(readFileSync(RESULTS_FILE, 'utf-8'));
    const filtered = allResults.filter(r => TARGET_CITIES.includes(r.city));
    
    let enrichment = {};
    if (existsSync(ENRICH_FILE)) {
        enrichment = JSON.parse(readFileSync(ENRICH_FILE, 'utf-8'));
    }
    
    const need = filtered
        .filter(r => (r.name || '').trim() && !(r.name.trim() in enrichment))
        .sort((a, b) => (a.city || '').localeCompare(b.city || ''));
    
    console.log(`[${ts}] Need enrichment: ${need.length}`);
    
    const batch = need.slice(startIdx, startIdx + batchSize);
    if (!batch.length) {
        console.log('Nothing to process!');
        return;
    }
    
    // Initialize z-ai SDK
    const zai = await ZAI.create();
    
    let ok = 0, fail = 0, noWeb = 0;
    
    for (let i = 0; i < batch.length; i++) {
        const firm = batch[i];
        const name = (firm.name || '').trim();
        const city = firm.city || '';
        const safe = safeFilename(name);
        const now = new Date().toISOString().slice(0, 19);
        
        const timeStr = new Date().toLocaleTimeString('de-DE', {hour12:false});
        console.log(`[${timeStr}] [${i+1}/${batch.length}] ${name} (${city})`);
        
        try {
            // Step 1: Search for website
            const website = await searchWebsite(zai, name, city);
            
            if (!website) {
                noWeb++;
                enrichment[name] = { website:'', html_file:'', email:'', status:'no_website', at: now };
                writeFileSync(ENRICH_FILE, JSON.stringify(enrichment, null, 0));
                continue;
            }
            
            // Step 2: Download HTML
            const result = await httpGet(website);
            
            if (!result || !result.ct.includes('text/html')) {
                fail++;
                enrichment[name] = { website, html_file:'', email:'', status:'html_fail', at: now };
                writeFileSync(ENRICH_FILE, JSON.stringify(enrichment, null, 0));
                continue;
            }
            
            // Save HTML
            const htmlPath = join(HTML_DIR, `${safe}.html`);
            writeFileSync(htmlPath, result.data);
            
            // Extract email
            let email = extractEmails(result.data);
            
            // Try impressum if no email
            let impUrl = '';
            if (!email) {
                const baseUrl = (result.finalUrl || website).replace(/\/[^/]*$/, '');
                for (const path of ['/impressum', '/kontakt']) {
                    const fullUrl = baseUrl + path;
                    const impResult = await httpGet(fullUrl, 5000);
                    if (impResult && impResult.ct.includes('text/html')) {
                        const impEmail = extractEmails(impResult.data);
                        if (impEmail) {
                            email = impEmail;
                            impUrl = fullUrl;
                            const impPath = join(HTML_DIR, `${safe}_imp${path.replace(/\//g, '_')}.html`);
                            writeFileSync(impPath, impResult.data);
                            break;
                        }
                    }
                }
            }
            
            ok++;
            const timeStr2 = new Date().toLocaleTimeString('de-DE', {hour12:false});
            console.log(`[${timeStr2}]   OK: ${(result.finalUrl || website).slice(0,60)} | ${result.data.length}b | email=${email || '-'}`);
            enrichment[name] = {
                website: result.finalUrl || website,
                html_file: `website_html/${safe}.html`,
                email,
                impressum_url: impUrl,
                html_size: result.data.length,
                status: 'success',
                at: now,
            };
            writeFileSync(ENRICH_FILE, JSON.stringify(enrichment, null, 0));
            
        } catch (e) {
            const timeStr3 = new Date().toLocaleTimeString('de-DE', {hour12:false});
            console.log(`[${timeStr3}]   ERROR: ${String(e).slice(0, 80)}`);
            enrichment[name] = { website:'', html_file:'', email:'', status:'error', error: String(e).slice(0,100), at: now };
            writeFileSync(ENRICH_FILE, JSON.stringify(enrichment, null, 0));
        }
    }
    
    const timeStr4 = new Date().toLocaleTimeString('de-DE', {hour12:false});
    console.log(`\n[${timeStr4}] Done! OK:${ok} Fail:${fail} NoWeb:${noWeb} | Total enriched:${Object.keys(enrichment).length}`);
}

main().catch(e => { console.error(e); process.exit(1); });
