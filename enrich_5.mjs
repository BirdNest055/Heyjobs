import ZAI from 'z-ai-web-dev-sdk';
import { execSync } from 'child_process';
import { writeFileSync } from 'fs';

const DB = '/home/z/my-project/download/it_employers.db';
const GIT = '/home/z/my-project';

const zai = await ZAI.create();
console.log('SDK ready');

// Simple query
writeFileSync('/tmp/q.py', "import sqlite3,json\nconn=sqlite3.connect('" + DB + "')\nprint(json.dumps([list(r) for r in conn.execute(\"SELECT id, firmenname, ort, it_relevanz_score FROM employers WHERE website_url='' AND hat_it_jobs=1 ORDER BY it_relevanz_score DESC LIMIT 5\").fetchall()]))\nconn.close()");
const rows = JSON.parse(execSync('python3 /tmp/q.py', {encoding:'utf8'}));
console.log('Processing ' + rows.length + ' employers');

for (const [empId, name, ort, score] of rows) {
    console.log('Searching: ' + name);
    try {
        const results = await zai.functions.invoke('web_search', {query: name + ' ' + ort + ' website', num: 5});
        const skip = ['facebook.com','linkedin.com','wikipedia.org','gelbeseiten.de','kununu.com','google.com','northdata.de'];
        let website = null;
        for (const r of results) {
            if (!skip.some(d => r.url.toLowerCase().includes(d))) { website = r.url; break; }
        }
        if (!website && results.length > 0) website = results[0].url;
        
        if (website) {
            writeFileSync('/tmp/u.py', "import sqlite3\nconn=sqlite3.connect('" + DB + "')\nconn.execute(\"UPDATE employers SET website_url = ?, scrape_status = 'enriched' WHERE id = ?\", ['" + website.replace(/'/g,"''") + "', " + empId + "])\nconn.commit()\nconn.close()");
            execSync('python3 /tmp/u.py');
            console.log('  -> ' + website);
        } else {
            console.log('  -> not found');
        }
    } catch(e) {
        console.error('  Error: ' + e.message);
    }
    await new Promise(r => setTimeout(r, 500));
}

console.log('Done with this batch');
