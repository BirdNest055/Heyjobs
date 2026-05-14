const ZAI = require('z-ai-web-dev-sdk').default;
const fs = require('fs');

async function main() {
    const batchSize = parseInt(process.argv[2]) || 20;
    const delay = parseInt(process.argv[3]) || 3000;
    
    const results = JSON.parse(fs.readFileSync('/home/z/my-project/gmaps_erlangen_results.json', 'utf-8'));
    let enrichment = {};
    try { enrichment = JSON.parse(fs.readFileSync('/home/z/my-project/website_enrichment.json', 'utf-8')); } catch {}
    
    const TARGET = ['Bamberg', 'Erlangen', 'Nürnberg'];
    const filtered = results.filter(r => TARGET.includes(r.city));
    const need = filtered.filter(r => r.name && r.name.trim() && !enrichment[r.name.trim()]);
    
    const batch = need.slice(0, batchSize);
    console.log(`[BATCH] ${batch.length} firms (${need.length} total remaining)`);
    
    const zai = await ZAI.create();
    
    for (let i = 0; i < batch.length; i++) {
        const firm = batch[i];
        const name = firm.name.trim();
        const city = firm.city;
        const query = `${name} ${city} Webseite`;
        
        let website = '';
        let retries = 0;
        
        while (retries < 3) {
            try {
                const searchResult = await zai.functions.invoke('web_search', {query, num: 5});
                const skip = ['facebook.com','instagram.com','twitter.com','linkedin.com','youtube.com',
                    'wikipedia.org','maps.google','gelbeseiten.de','yelp.de','tripadvisor',
                    'firmenwissen.de','northdata','kompass','wlw.de','cylex.de','dasoertliche.de'];
                
                for (const item of (searchResult || [])) {
                    const url = item.url || '';
                    if (!skip.some(s => url.toLowerCase().includes(s))) { website = url; break; }
                }
                if (!website && searchResult && searchResult.length > 0) website = searchResult[0].url || '';
                
                enrichment[name] = { website, status: website ? 'search_found' : 'no_website', at: new Date().toISOString().slice(0, 19) };
                console.log(`[${i+1}/${batch.length}] ${name}: ${website || 'NOT FOUND'}`);
                fs.writeFileSync('/home/z/my-project/website_enrichment.json', JSON.stringify(enrichment, null, 1));
                break;
            } catch (e) {
                retries++;
                if (e.message && e.message.includes('429')) {
                    const wait = 30 * retries;
                    console.log(`[${i+1}/${batch.length}] ${name}: Rate limited, waiting ${wait}s...`);
                    await new Promise(r => setTimeout(r, wait * 1000));
                } else {
                    console.log(`[${i+1}/${batch.length}] ${name}: ERROR ${e.message}`);
                    enrichment[name] = {website:'', status:'search_error', at: new Date().toISOString().slice(0,19)};
                    fs.writeFileSync('/home/z/my-project/website_enrichment.json', JSON.stringify(enrichment, null, 1));
                    break;
                }
            }
        }
        await new Promise(r => setTimeout(r, delay));
    }
    console.log(`[DONE] Total enriched: ${Object.keys(enrichment).length}`);
}
main().catch(e => console.error('Fatal:', e.message));
