import ZAI from 'z-ai-web-dev-sdk';
import Database from 'better-sqlite3';
import { writeFileSync, mkdirSync, existsSync } from 'fs';
import { execSync } from 'child_process';
import { createHash } from 'crypto';
import { sanitize } from './sanitize_utils.js';

const DB_PATH = '/home/z/my-project/download/it_employers.db';
const HTML_DIR = '/home/z/my-project/download/it_employer_html';
const GIT_DIR = '/home/z/my-project';
const COMMIT_EVERY = 10;

// Initialize
mkdirSync(HTML_DIR, { recursive: true });
const zai = await ZAI.create();

// SQLite - use better-sqlite3 for sync access
let db;
try {
    db = new Database(DB_PATH);
} catch(e) {
    // Fallback to subprocess
    console.log('better-sqlite3 not available, using subprocess mode');
    db = null;
}

function sanitizeFilename(name, maxLen = 80) {
    let result = name.replace(/[äöüßÄÖÜ]/g, m => ({ä:'ae',ö:'oe',ü:'ue',ß:'ss',Ä:'Ae',Ö:'Oe',Ü:'Ue'}[m]));
    result = result.replace(/[^a-zA-Z0-9._-]/g, '_');
    result = result.replace(/_+/g, '_').replace(/^_|_$/g, '');
    return result.slice(0, maxLen);
}

function queryDb(sql, params = []) {
    if (db) return db.prepare(sql).all(...params);
    // Fallback
    const result = execSync(`python3 -c "import sqlite3,json; c=sqlite3.connect('${DB_PATH}'); print(json.dumps([list(r) for r in c.execute('${sql.replace(/'/g,"\\'")}',${JSON.stringify(params)}).fetchall()]))"`, {encoding:'utf8'});
    return JSON.parse(result);
}

function runDb(sql, params = []) {
    if (db) { db.prepare(sql).run(...params); return; }
    execSync(`python3 -c "import sqlite3; c=sqlite3.connect('${DB_PATH}'); c.execute('${sql.replace(/'/g,"\\'")}',${JSON.stringify(params)}); c.commit()"`);
}

async function searchWebsite(companyName, city) {
    try {
        const results = await zai.functions.invoke('web_search', {
            query: `${companyName} ${city} website`,
            num: 5
        });
        
        const skipDomains = ['facebook.com','linkedin.com','instagram.com','twitter.com','youtube.com',
            'wikipedia.org','gelbeseiten.de','dasoertliche.de','meinestadt.de','kununu.com',
            'yelp.de','google.com','xing.com','northdata.de','kompass.com','hotfrog.de',
            'cybo.com','cylex.de','dhd24.com','firmenwissen.de','unternehmensregister.de'];
        
        for (const r of results) {
            const urlLower = r.url.toLowerCase();
            if (!skipDomains.some(d => urlLower.includes(d))) {
                return r.url;
            }
        }
        // Fallback: first result even if from skip list
        if (results.length > 0) return results[0].url;
    } catch(e) {
        console.error(`  Search error: ${e.message}`);
    }
    return null;
}

async function downloadHtml(url, timeout = 8000) {
    try {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), timeout);
        const resp = await fetch(url, { 
            signal: controller.signal,
            headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' }
        });
        clearTimeout(timer);
        if (resp.ok) {
            const ct = resp.headers.get('content-type') || '';
            if (ct.includes('text/')) return await resp.text();
        }
    } catch(e) {}
    return null;
}

function saveHtml(content, companyName, suffix = '') {
    if (!content) return [null, null];
    const safeName = sanitizeFilename(companyName);
    const filename = suffix ? `${safeName}_${suffix}.html` : `${safeName}.html`;
    const filepath = `${HTML_DIR}/${filename}`;
    writeFileSync(filepath, content, 'utf8');
    const hash = createHash('sha256').update(content).digest('hex').slice(0, 16);
    return [filepath, hash];
}

function extractEmails(html) {
    const emails = html.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || [];
    const skip = ['example.com','sentry','webpack','noreply','wixpress','github','googleapis'];
    return [...new Set(emails.filter(e => !skip.some(s => e.toLowerCase().includes(s))))].slice(0, 3);
}

// MAIN
console.log('=== Website Enrichment via SDK ===\n');

// Get employers without websites
const rows = queryDb(`SELECT id, firmenname, ort, hat_it_jobs, it_relevanz_score 
    FROM employers WHERE website_url = '' AND hat_it_jobs = 1
    ORDER BY it_relevanz_score DESC`);
console.log(`IT-Arbeitgeber ohne Website: ${rows.length}`);

let processed = 0;
let found = 0;
let batchCommit = 0;

for (const row of rows) {
    const [empId, name, ort, hasIt, score] = row;
    processed++;
    
    // Search for website
    const website = await searchWebsite(name, ort);
    
    if (website) {
        found++;
        let htmlPath = null;
        let htmlHash = null;
        let impressumPath = null;
        let karrierePath = null;
        let impressumUrl = null;
        let karriereUrl = null;
        let email = '';
        
        // Download main page
        const mainHtml = await downloadHtml(website);
        if (mainHtml) {
            [htmlPath, htmlHash] = saveHtml(mainHtml, name);
            
            // Extract emails
            const emails = extractEmails(mainHtml);
            if (emails.length > 0) email = emails[0];
            
            // Find impressum/karriere links
            const lowerHtml = mainHtml.toLowerCase();
            // Simple regex to find links
            const linkRegex = /href=["']([^"']*(?:impressum|imprint|career|karriere|jobs|stellenangebote|bewerbung)[^"']*)["']/gi;
            let match;
            while ((match = linkRegex.exec(mainHtml)) !== null) {
                const href = match[1];
                if (!impressumUrl && /impressum|imprint/i.test(href)) {
                    impressumUrl = href.startsWith('/') ? new URL(href, website).href : href;
                }
                if (!karriereUrl && /karriere|career|jobs|stellenangebote|bewerbung/i.test(href)) {
                    karriereUrl = href.startsWith('/') ? new URL(href, website).href : href;
                }
            }
            
            // Download impressum
            if (impressumUrl) {
                const impHtml = await downloadHtml(impressumUrl, 6000);
                if (impHtml) {
                    [impressumPath] = saveHtml(impHtml, name, 'impressum');
                    const impEmails = extractEmails(impHtml);
                    if (impEmails.length > 0 && !email) email = impEmails[0];
                }
            }
            
            // Download karriere
            if (karriereUrl) {
                const karHtml = await downloadHtml(karriereUrl, 6000);
                if (karHtml) [karrierePath] = saveHtml(karHtml, name, 'karriere');
            }
        }
        
        // Update DB
        runDb(`UPDATE employers SET website_url = ?, website_html_path = ?, website_html_hash = ?,
            impressum_url = ?, impressum_html_path = ?, karriere_url = ?, karriere_html_path = ?,
            email = COALESCE(NULLIF(email, ''), ?), scrape_status = 'enriched', updated_at = CURRENT_TIMESTAMP
            WHERE id = ?`, 
            [website, htmlPath || '', htmlHash || '', impressumUrl || '', impressumPath || '',
             karriereUrl || '', karrierePath || '', email, empId]);
        
        console.log(`  [${processed}/${rows.length}] ✓IT Sc:${score} | ${name.slice(0,42).padEnd(42)} | ${website.slice(0,40)}`);
    } else {
        runDb(`UPDATE employers SET scrape_status = 'no_website', updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [empId]);
        console.log(`  [${processed}/${rows.length}]    Sc:${score} | ${name.slice(0,42).padEnd(42)} | -`);
    }
    
    batchCommit++;
    
    // Git commit every COMMIT_EVERY
    if (batchCommit >= COMMIT_EVERY) {
        try {
            execSync('git add -A', {cwd: GIT_DIR, timeout: 30000});
            const msg = `Website-Enrichment: ${processed}/${rows.length} | ${found} found`;
            execSync(`git commit -m "${msg}"`, {cwd: GIT_DIR, timeout: 30000});
            execSync('git push origin main', {cwd: GIT_DIR, timeout: 60000});
            console.log(`  >>> GIT PUSH (${found} Websites) <<<`);
        } catch(e) {}
        batchCommit = 0;
    }
    
    // Rate limit
    await new Promise(r => setTimeout(r, 500));
}

// Final commit
try {
    execSync('git add -A', {cwd: GIT_DIR, timeout: 30000});
    execSync(`git commit -m "Website-Enrichment FERTIG: ${found}/${rows.length}"`, {cwd: GIT_DIR, timeout: 30000});
    execSync('git push origin main', {cwd: GIT_DIR, timeout: 60000});
} catch(e) {}

console.log(`\nFERTIG! ${found}/${rows.length} Websites für IT-Arbeitgeber gefunden`);
