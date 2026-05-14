const ZAI = require('z-ai-web-dev-sdk').default;
const fs = require('fs');

async function main() {
    const startIdx = parseInt(process.argv[2]) || 0;
    const batchSize = parseInt(process.argv[3]) || 20;
    
    // Load data
    const results = JSON.parse(fs.readFileSync('/home/z/my-project/gmaps_erlangen_results.json', 'utf-8'));
    const enrichment = JSON.parse(fs.readFileSync('/home/z/my-project/website_enrichment.json', 'utf-8'));
    
    const TARGET = ['Bamberg', 'Erlangen', 'Nürnberg'];
    const filtered = results.filter(r => TARGET.includes(r.city));
    const need = filtered.filter(r => r.name && r.name.trim() && !enrichment[r.name.trim()]);
    
    const batch = need.slice(startIdx, startIdx + batchSize);
    console.log(`[BATCH] ${batch.length} firms to search (from ${startIdx}, ${need.length} total remaining)`);
    
    const zai = await ZAI.create();
    
    for (let i = 0; i < batch.length; i++) {
        const firm = batch[i];
        const name = firm.name.trim();
        const city = firm.city;
        const query = `${name} ${city} Webseite`;
        
        try {
            const searchResult = await zai.functions.invoke('web_search', {query, num: 5});
            
            let website = '';
            const skip = ['facebook.com','instagram.com','twitter.com','linkedin.com','youtube.com',
                'wikipedia.org','maps.google','gelbeseiten.de','yelp.de','tripadvisor',
                'firmenwissen.de','northdata','kompass','wlw.de','cylex.de'];
            
            for (const item of (searchResult || [])) {
                const url = item.url || '';
                if (!skip.some(s => url.toLowerCase().includes(s))) {
                    website = url;
                    break;
                }
            }
            if (!website && searchResult && searchResult.length > 0) {
                website = searchResult[0].url || '';
            }
            
            enrichment[name] = {
                website: website,
                search_snippet: (searchResult && searchResult[0]) ? searchResult[0].snippet : '',
                status: website ? 'search_found' : 'no_website',
                at: new Date().toISOString().slice(0, 19)
            };
            
            console.log(`[${i+1}/${batch.length}] ${name}: ${website || 'NOT FOUND'}`);
            
            // Save after each
            fs.writeFileSync('/home/z/my-project/website_enrichment.json', JSON.stringify(enrichment, null, 1));
            
            // Delay to avoid rate limits
            await new Promise(r => setTimeout(r, 1500));
            
        } catch (e) {
            console.error(`[${i+1}/${batch.length}] ${name}: ERROR ${e.message}`);
            enrichment[name] = {website:'', status:'search_error', at: new Date().toISOString().slice(0,19)};
            fs.writeFileSync('/home/z/my-project/website_enrichment.json', JSON.stringify(enrichment, null, 1));
            await new Promise(r => setTimeout(r, 3000));
        }
    }
    
    console.log(`[DONE] Batch complete. Total enriched: ${Object.keys(enrichment).length}`);
}

main().catch(e => console.error('Fatal:', e.message));
