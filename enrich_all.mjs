import ZAI from 'z-ai-web-dev-sdk';
import { execSync } from 'child_process';
import { writeFileSync } from 'fs';

const DB = '/home/z/my-project/download/it_employers.db';
const GIT = '/home/z/my-project';

const zai = await ZAI.create();
console.log('=== Full SDK Enrichment ===');

// Get all IT employers without website
writeFileSync('/tmp/qall.py', "import sqlite3,json\nconn=sqlite3.connect('" + DB + "')\nprint(json.dumps([list(r) for r in conn.execute(\"SELECT id, firmenname, ort, it_relevanz_score FROM employers WHERE website_url='' AND hat_it_jobs=1 ORDER BY it_relevanz_score DESC\").fetchall()]))\nconn.close()");
const rows = JSON.parse(execSync('python3 /tmp/qall.py', {encoding:'utf8'}));
console.log('IT-Arbeitgeber ohne Website: ' + rows.length);

let processed = 0, found = 0;

for (const [empId, name, ort, score] of rows) {
    processed++;
    let website = null;
    try {
        const results = await zai.functions.invoke('web_search', {query: name + ' ' + ort + ' website', num: 5});
        const skip = ['facebook.com','linkedin.com','instagram.com','wikipedia.org','gelbeseiten.de','dasoertliche.de','kununu.com','google.com','northdata.de','firmenwissen.de','yelp.de'];
        for (const r of results) {
            if (!skip.some(d => r.url.toLowerCase().includes(d))) { website = r.url; break; }
        }
        if (!website && results.length > 0) website = results[0].url;
    } catch(e) {
        console.error('  Search error: ' + e.message);
    }
    
    if (website) {
        found++;
        writeFileSync('/tmp/u.py', "import sqlite3\nconn=sqlite3.connect('" + DB + "')\nconn.execute(\"UPDATE employers SET website_url = ?, scrape_status = 'enriched', updated_at = CURRENT_TIMESTAMP WHERE id = ?\", ['" + website.replace(/'/g,"''") + "', " + empId + "])\nconn.commit()\nconn.close()");
        execSync('python3 /tmp/u.py');
        console.log('  [' + processed + '/' + rows.length + '] + ' + name.slice(0,38).padEnd(38) + ' | ' + website.slice(0,40));
    } else {
        writeFileSync('/tmp/u.py', "import sqlite3\nconn=sqlite3.connect('" + DB + "')\nconn.execute(\"UPDATE employers SET scrape_status = 'no_website' WHERE id = ?\", [" + empId + "])\nconn.commit()\nconn.close()");
        execSync('python3 /tmp/u.py');
        console.log('  [' + processed + '/' + rows.length + ']   ' + name.slice(0,38).padEnd(38) + ' | -');
    }
    
    // Git commit every 10
    if (processed % 10 === 0) {
        try {
            execSync('git add -A', {cwd: GIT, timeout: 30000});
            execSync('git commit -m "SDK: ' + processed + '/' + rows.length + ' ' + found + 'ws"', {cwd: GIT, timeout: 30000});
            execSync('git push origin main', {cwd: GIT, timeout: 60000});
            console.log('  >>> GIT PUSH (' + found + ' Websites) <<<');
        } catch(e) {}
    }
    
    await new Promise(r => setTimeout(r, 300));
}

// Final
try {
    execSync('git add -A', {cwd: GIT, timeout: 30000});
    execSync('git commit -m "SDK FERTIG: ' + found + '/' + rows.length + '"', {cwd: GIT, timeout: 30000});
    execSync('git push origin main', {cwd: GIT, timeout: 60000});
} catch(e) {}

console.log('\nFERTIG! ' + found + '/' + rows.length + ' Websites gefunden');
