#!/usr/bin/env node
/**
 * Step 1: Batch web search for employer websites
 * Uses z-ai-web-dev-sdk with proper rate limiting
 * Outputs results to a JSON file for Python to consume
 */

import ZAI from 'z-ai-web-dev-sdk';
import { readFileSync, writeFileSync, existsSync } from 'fs';

const INPUT_FILE = '/home/z/my-project/search_needed.json';
const OUTPUT_FILE = '/home/z/my-project/website_search_results.json';

const BATCH_DELAY = 3500;  // 3.5s between searches to avoid rate limit
const JOB_BOARD_DOMAINS = ['indeed', 'stepstone', 'kununu', 'glassdoor', 'xing.com', 'linkedin.com', 'arbeitsagentur', 'stellenanzeigen', 'jobware', 'monster.de', 'meinestadt', 'ausbildung', 'jobtuple', 'stellenonline', 'jobsuche', 'jobware.de'];

async function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function searchWithRetry(zai, query, retries = 2) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      return await zai.functions.invoke("web_search", { query, num: 5 });
    } catch (e) {
      if (e.message && e.message.includes('429')) {
        const wait = 5000 * (attempt + 1);
        console.log(`    Rate limited, waiting ${wait/1000}s...`);
        await sleep(wait);
        continue;
      }
      return null;
    }
  }
  return null;
}

async function main() {
  const startIdx = parseInt(process.argv[2] || '0');
  const count = parseInt(process.argv[3] || '50');
  
  console.log(`=== Website Search: ${startIdx} to ${startIdx + count - 1} ===`);
  
  // Load employers
  const employers = JSON.parse(readFileSync(INPUT_FILE, 'utf-8'));
  
  // Deduplicate
  const seen = new Map();
  const unique = [];
  for (const e of employers) {
    const key = e.name.toLowerCase().trim();
    if (!seen.has(key)) {
      seen.set(key, e);
      unique.push(e);
    }
  }
  console.log(`${employers.length} entries → ${unique.length} unique`);
  
  // Load existing results
  let results = {};
  if (existsSync(OUTPUT_FILE)) {
    results = JSON.parse(readFileSync(OUTPUT_FILE, 'utf-8'));
  }
  console.log(`Already have ${Object.keys(results).length} results`);
  
  // Get slice to process
  const slice = unique.slice(startIdx, startIdx + count);
  const pending = slice.filter(e => !results[e.name.toLowerCase().trim()]);
  console.log(`Processing ${pending.length} new employers`);
  
  const zai = await ZAI.create();
  let found = 0;
  let startTime = Date.now();
  
  for (let i = 0; i < pending.length; i++) {
    const e = pending[i];
    const query = e.stadt 
      ? `"${e.name}" ${e.stadt} official website`
      : `"${e.name}" Germany official website`;
    
    const searchResults = await searchWithRetry(zai, query);
    
    let website = '';
    let source = 'not_found';
    
    if (searchResults && searchResults.length > 0) {
      for (const sr of searchResults) {
        const url = sr.url || sr.host_name || '';
        if (url && !JOB_BOARD_DOMAINS.some(d => url.toLowerCase().includes(d))) {
          website = url.startsWith('http') ? url : `https://${url}`;
          source = 'web_search';
          break;
        }
      }
      if (!website && searchResults[0]) {
        const url = searchResults[0].url || searchResults[0].host_name || '';
        if (url) {
          website = url.startsWith('http') ? url : `https://${url}`;
          source = 'web_search_fallback';
        }
      }
    }
    
    const key = e.name.toLowerCase().trim();
    results[key] = {
      name: e.name,
      website: website,
      source: source,
      plz: e.plz,
      stadt: e.stadt,
    };
    
    if (website) found++;
    
    // Progress
    const elapsed = ((Date.now() - startTime) / 1000).toFixed(0);
    const rate = (i + 1) / (Date.now() - startTime) * 1000;
    console.log(`  [${i+1}/${pending.length}] ${website ? '✓' : '✗'} ${e.name.substring(0, 50).padEnd(50)} → ${website.substring(0, 45)}`);
    
    // Save every 10 results
    if ((i + 1) % 10 === 0) {
      writeFileSync(OUTPUT_FILE, JSON.stringify(results, null, 0));
      console.log(`  → Saved ${Object.keys(results).length} results`);
    }
    
    await sleep(BATCH_DELAY);
  }
  
  // Final save
  writeFileSync(OUTPUT_FILE, JSON.stringify(results, null, 0));
  
  const total = Object.keys(results).length;
  const totalWithWebsite = Object.values(results).filter(r => r.website).length;
  console.log(`\n=== DONE ===`);
  console.log(`Processed: ${pending.length} | Found: ${found} | Total: ${total} (${totalWithWebsite} with website)`);
  console.log(`Time: ${((Date.now() - startTime) / 1000).toFixed(0)}s`);
}

main().catch(console.error);
