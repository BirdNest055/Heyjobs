import ZAI from 'z-ai-web-dev-sdk';
import { execSync } from 'child_process';
import { writeFileSync, mkdirSync } from 'fs';

const DB = '/home/z/my-project/download/it_employers.db';
const HTML = '/home/z/my-project/download/it_employer_html';
const GIT = '/home/z/my-project';

mkdirSync(HTML, { recursive: true });

const zai = await ZAI.create();

// Get IT employers without website from SQLite via Python
function query(sql) {
    const result = execSync(`python3 -c "import sqlite3,json;c=sqlite3.connect('${DB}');print(json.dumps([list(r) for r in c.execute(\`${sql}\`).fetchall()]))"`, {encoding:'utf8'});
    return JSON.parse(result);
}

function run(sql, params=[]) {
    execSync(`python3 -c "import sqlite3;c=sqlite3.connect('${DB}');c.execute(\`${sql}\`,${JSON.stringify(params)});c.commit()"`);
}

console.log('=== Batch SDK Enrichment ===\n');

const rows = query(`SELECT id, firmenname, ort, it_relevanz_score FROM employers WHERE website_url='' AND hat_it_jobs=1 ORDER BY it_relevanz_score DESC`);
console.log(`IT-Arbeitgeber ohne Website: ${rows.length}`);

let processed = 0, found = 0;

for (const [empId, name, ort, score] of rows) {
    processed++;
    
    let website = null;
    try {
        const results = await zai.functions.invoke('web_search', {
            query: `${name} ${ort} website`,
            num: 5
        });
        
        const skip = ['facebook.com','linkedin.com','instagram.com','twitter.com','youtube.com',
            'wikipedia.org','gelbeseiten.de','dasoertliche.de','kununu.com','google.com',
            'xing.com','northdata.de','firmenwissen.de','yelp.de','kompass.com'];
        
        for (const r of results) {
            if (!skip.some(d => r.url.toLowerCase().includes(d))) {
                website = r.url;
                break;
            }
        }
        if (!website && results.length > 0) website = results[0].url;
    } catch(e) {
        console.error(`  Search error for ${name}: ${e.message}`);
    }
    
    if (website) {
        found++;
        // Update DB with website
        run(`UPDATE employers SET website_url = ?, scrape_status = 'enriched', updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [website, empId]);
        console.log(`  [${processed}/${rows.length}] ✓ Sc:${score} | ${name.slice(0,38).padEnd(38)} | ${website.slice(0,40)}`);
    } else {
        run(`UPDATE employers SET scrape_status = 'no_website', updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [empId]);
        console.log(`  [${processed}/${rows.length}]   Sc:${score} | ${name.slice(0,38).padEnd(38)} | -`);
    }
    
    // Commit every 10
    if (processed % 10 === 0) {
        try {
            execSync('git add -A', {cwd: GIT, timeout: 30000});
            execSync(`git commit -m "SDK-Batch: ${processed}/${rows.length} | ${found} ws"`, {cwd: GIT, timeout: 30000});
            execSync('git push origin main', {cwd: GIT, timeout: 60000});
            console.log(`  >>> GIT PUSH (${found} Websites) <<<`);
        } catch(e) {}
    }
    
    // Rate limit
    await new Promise(r => setTimeout(r, 300));
}

// Final
try {
    execSync('git add -A', {cwd: GIT, timeout: 30000});
    execSync(`git commit -m "SDK-Batch FERTIG: ${found}/${rows.length}"`, {cwd: GIT, timeout: 30000});
    execSync('git push origin main', {cwd: GIT, timeout: 60000});
} catch(e) {}

console.log(`\nFERTIG! ${found}/${rows.length} Websites gefunden`);
