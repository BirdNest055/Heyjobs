import ZAI from 'z-ai-web-dev-sdk';
import fs from 'fs';

async function main() {
    const zai = await ZAI.create();
    
    const data = JSON.parse(fs.readFileSync('/home/z/my-project/download/arbeitsagentur_employers.json', 'utf-8'));
    
    // Get unique company names (top 300)
    const seen = new Set();
    const companies = [];
    for (const emp of data) {
        const name = (emp.Name || '').trim();
        if (name && !seen.has(name)) {
            seen.add(name);
            companies.push(emp);
        }
        if (companies.length >= 300) break;
    }
    
    console.log(`Searching websites for ${companies.length} unique companies...`);
    
    let found = 0;
    for (let i = 0; i < companies.length; i++) {
        const emp = companies[i];
        const name = emp.Name;
        const stadt = emp.Stadt || '';
        
        try {
            const results = await zai.functions.invoke("web_search", {
                query: name + ' ' + stadt + ' Website Kontakt',
                num: 3
            });
            
            if (results && results.length > 0) {
                for (const r of results) {
                    const url = r.url || '';
                    const snippet = r.snippet || '';
                    
                    const skipDomains = ['arbeitsagentur', 'google', 'facebook', 'linkedin', 'xing', 'indeed', 'stepstone', 'kununu', 'glassdoor', 'wiki'];
                    const shouldSkip = skipDomains.some(d => url.toLowerCase().includes(d));
                    
                    if (url && !shouldSkip) {
                        for (const e of data) {
                            if (e.Name === name && !e.Website) {
                                e.Website = url;
                                found++;
                                break;
                            }
                        }
                        break;
                    }
                    
                    // Extract email from snippet
                    const emailMatch = snippet.match(/[\w.+-]+@[\w-]+\.[\w.-]+/);
                    if (emailMatch) {
                        for (const e of data) {
                            if (e.Name === name && !e['E-Mail']) {
                                e['E-Mail'] = emailMatch[0];
                                break;
                            }
                        }
                    }
                }
            }
        } catch (e) {
            // Skip errors
        }
        
        if ((i + 1) % 25 === 0) {
            console.log('  ' + (i+1) + '/' + companies.length + ' | Found ' + found + ' websites');
            fs.writeFileSync('/home/z/my-project/download/arbeitsagentur_employers.json', JSON.stringify(data, null, 2));
        }
        
        await new Promise(r => setTimeout(r, 250));
    }
    
    fs.writeFileSync('/home/z/my-project/download/arbeitsagentur_employers.json', JSON.stringify(data, null, 2));
    
    const withWeb = data.filter(e => e.Website).length;
    const withEmail = data.filter(e => e['E-Mail']).length;
    console.log('Done! Websites: ' + withWeb + ' | Emails: ' + withEmail + ' | Total: ' + data.length);
}

main().catch(console.error);
