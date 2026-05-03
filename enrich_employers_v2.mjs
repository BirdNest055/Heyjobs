#!/usr/bin/env node
/**
 * Employer Enrichment V2 - Optimized for E-Mail extraction
 * 
 * Strategy:
 * 1. Web search for employer website (if missing)
 * 2. Use web_reader to get impressum/kontakt page content
 * 3. Use LLM to extract contact/bewerbung emails from page content
 * 4. Works in batches with rate limiting
 * 
 * Rate limit handling: 429 errors trigger exponential backoff
 */

import ZAI from 'z-ai-web-dev-sdk';
import { readFileSync, writeFileSync, existsSync } from 'fs';

const RESULTS_FILE = '/home/z/my-project/enrichment_results_v2.json';
const INPUT_FILE = '/home/z/my-project/search_needed.json';

// Config
const BATCH_SIZE = 15;
const CONCURRENCY = 3;         // parallel per batch
const SEARCH_DELAY = 800;      // ms between API calls
const MAX_RETRIES = 3;
const BACKOFF_BASE = 3000;     // base ms for exponential backoff

async function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function callWithRetry(fn, maxRetries = MAX_RETRIES) {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return await fn();
    } catch (e) {
      if (e.message && e.message.includes('429')) {
        const backoff = BACKOFF_BASE * Math.pow(2, attempt);
        console.log(`  Rate limited, waiting ${backoff/1000}s...`);
        await sleep(backoff);
        continue;
      }
      throw e;
    }
  }
  return null;
}

async function searchWebsite(zai, name, stadt) {
  const query = stadt 
    ? `${name} ${stadt} official website`
    : `${name} Germany official website`;
  
  return callWithRetry(async () => {
    const result = await zai.functions.invoke("web_search", {
      query,
      num: 5
    });
    
    if (!result || result.length === 0) return null;
    
    // Filter out job boards
    const jobBoardDomains = ['indeed', 'stepstone', 'kununu', 'glassdoor', 'xing.com', 'linkedin.com', 'arbeitsagentur', 'stellenanzeigen', 'jobware', 'monster.de', 'meinestadt', 'ausbildung', 'jobtuple'];
    
    for (const r of result) {
      const url = r.url || '';
      const isJobBoard = jobBoardDomains.some(d => url.toLowerCase().includes(d));
      if (!isJobBoard && url) {
        return url.startsWith('http') ? url : `https://${url}`;
      }
    }
    
    // Fallback: return first result
    const first = result[0];
    const url = first.url || '';
    return url.startsWith('http') ? url : (url ? `https://${url}` : null);
  });
}

async function extractEmailsFromPage(zai, url) {
  return callWithRetry(async () => {
    const result = await zai.functions.invoke("web_reader", {
      url
    });
    
    if (!result) return [];
    
    const content = result.text || result.html || result.content || '';
    
    if (!content || content.length < 20) return [];
    
    // Use LLM to extract emails intelligently
    const prompt = `Extract ALL email addresses from the following webpage content of a German company. 
Focus especially on emails for: Bewerbung (applications), Karriere (careers), Kontakt (contact), Recruiting, HR.
Return ONLY a JSON array of email strings, nothing else. If no emails found, return [].

Content:
${content.substring(0, 6000)}`;

    try {
      const completion = await zai.chat.completions.create({
        messages: [
          { role: 'system', content: 'You extract email addresses from German company webpages. Return ONLY a JSON array of email strings, nothing else.' },
          { role: 'user', content: prompt }
        ],
        temperature: 0,
        max_tokens: 500
      });
      
      const response = completion.choices?.[0]?.message?.content || '[]';
      
      // Parse JSON from response
      const jsonMatch = response.match(/\[[\s\S]*?\]/);
      if (jsonMatch) {
        const emails = JSON.parse(jsonMatch[0]);
        return emails.filter(e => typeof e === 'string' && e.includes('@') && e.length > 5);
      }
    } catch (e) {
      // Fallback to regex if LLM fails
      const emailRegex = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
      const emails = [...content.matchAll(emailRegex)].map(m => m[0].toLowerCase());
      const unwanted = ['example.com', 'test.com', 'domain.com', 'mustermann', '.png', '.jpg'];
      return emails.filter(e => !unwanted.some(p => e.includes(p)));
    }
    
    return [];
  });
}

async function enrichEmployer(zai, employer) {
  const { name, stadt, current_website } = employer;
  const result = { name, website: current_website || '', emails: [], bewerbung_email: '', kontakt_email: '', source: '' };
  
  // Step 1: Find website if missing
  let website = current_website;
  if (!website) {
    website = await searchWebsite(zai, name, stadt);
    if (website) {
      result.website = website;
      result.source = 'web_search';
    }
    await sleep(SEARCH_DELAY);
  }
  
  if (!website) {
    result.source = 'not_found';
    return result;
  }
  
  // Normalize URL for contact pages
  let baseUrl = website.replace(/\/$/, '');
  try {
    const urlObj = new URL(baseUrl);
    baseUrl = `${urlObj.protocol}//${urlObj.hostname}`;
  } catch(e) {}
  
  // Step 2: Try to get emails from impressum (required by German law - always has contact info!)
  const contactUrls = [
    `${baseUrl}/impressum`,
    `${baseUrl}/kontakt`, 
    `${baseUrl}/Impressum`,
    `${baseUrl}/Kontakt`,
    website,  // main page as fallback
    `${baseUrl}/karriere`,
    `${baseUrl}/career`,
    `${baseUrl}/ueber-uns`,
  ];
  
  let allEmails = [];
  
  for (const contactUrl of contactUrls) {
    const emails = await extractEmailsFromPage(zai, contactUrl);
    if (emails.length > 0) {
      allEmails.push(...emails);
      result.scraped_from = contactUrl;
      break;  // Stop after first successful extraction
    }
    await sleep(SEARCH_DELAY);
  }
  
  // Deduplicate and categorize
  const uniqueEmails = [...new Set(allEmails.map(e => e.toLowerCase()))];
  
  // Find bewerbung-specific email
  const bewerbungKeywords = ['bewerbung', 'karriere', 'career', 'recruiting', 'jobs', 'hr', 'talent', 'personal'];
  const kontaktKeywords = ['kontakt', 'contact', 'info', 'impressum'];
  
  result.emails = uniqueEmails;
  result.bewerbung_email = uniqueEmails.find(e => bewerbungKeywords.some(k => e.includes(k))) || '';
  result.kontakt_email = uniqueEmails.find(e => kontaktKeywords.some(k => e.includes(k))) || '';
  
  // If no specialized email, first email is the general one
  if (!result.kontakt_email && uniqueEmails.length > 0) {
    result.kontakt_email = uniqueEmails[0];
  }
  
  return result;
}

function loadResults() {
  if (existsSync(RESULTS_FILE)) {
    return JSON.parse(readFileSync(RESULTS_FILE, 'utf-8'));
  }
  return [];
}

function saveResults(results) {
  writeFileSync(RESULTS_FILE, JSON.stringify(results, null, 0));
}

async function main() {
  const args = process.argv.slice(2);
  const startBatch = parseInt(args[0] || '0');
  const maxBatches = parseInt(args[1] || '5');
  
  console.log(`=== Employer Enrichment V2 ===`);
  console.log(`Starting from batch ${startBatch}, max ${maxBatches} batches`);
  
  // Load employers and deduplicate
  const employers = JSON.parse(readFileSync(INPUT_FILE, 'utf-8'));
  const seen = new Map();
  const uniqueEmployers = [];
  for (const e of employers) {
    const key = e.name.toLowerCase().trim();
    if (!seen.has(key)) {
      seen.set(key, e);
      uniqueEmployers.push(e);
    }
  }
  console.log(`${employers.length} entries → ${uniqueEmployers.length} unique names`);
  
  // Load existing results
  const existingResults = loadResults();
  const resultMap = new Map();
  for (const r of existingResults) {
    resultMap.set(r.name.toLowerCase().trim(), r);
  }
  console.log(`Already enriched: ${resultMap.size}`);
  
  const zai = await ZAI.create();
  
  // Create batches of unique employers not yet processed
  const pending = uniqueEmployers.filter(e => !resultMap.has(e.name.toLowerCase().trim()));
  console.log(`Pending: ${pending.length}`);
  
  const batches = [];
  for (let i = 0; i < pending.length; i += BATCH_SIZE) {
    batches.push(pending.slice(i, i + BATCH_SIZE));
  }
  
  const batchesToProcess = batches.slice(startBatch, startBatch + maxBatches);
  console.log(`Processing batches ${startBatch+1}-${startBatch+batchesToProcess.length} of ${batches.length} total`);
  
  let totalWebsites = 0;
  let totalEmails = 0;
  let globalStart = Date.now();
  
  for (let bi = 0; bi < batchesToProcess.length; bi++) {
    const batch = batchesToProcess[bi];
    const batchNum = startBatch + bi + 1;
    const batchStart = Date.now();
    
    console.log(`\n--- Batch ${batchNum} (${batch.length} employers) ---`);
    
    // Process with limited concurrency
    const batchResults = [];
    for (let i = 0; i < batch.length; i += CONCURRENCY) {
      const chunk = batch.slice(i, i + CONCURRENCY);
      const promises = chunk.map(e => 
        enrichEmployer(zai, e).catch(err => ({
          name: e.name,
          website: e.current_website || '',
          emails: [],
          bewerbung_email: '',
          kontakt_email: '',
          error: err.message
        }))
      );
      
      const chunkResults = await Promise.all(promises);
      batchResults.push(...chunkResults);
      
      if (i + CONCURRENCY < batch.length) {
        await sleep(SEARCH_DELAY);
      }
    }
    
    // Count successes
    let bw = 0, be = 0;
    for (const r of batchResults) {
      resultMap.set(r.name.toLowerCase().trim(), r);
      if (r.website) bw++;
      if (r.emails.length > 0) be++;
    }
    totalWebsites += bw;
    totalEmails += be;
    
    // Save after each batch
    saveResults([...resultMap.values()]);
    
    const elapsed = ((Date.now() - batchStart) / 1000).toFixed(1);
    const totalElapsed = ((Date.now() - globalStart) / 1000).toFixed(0);
    const progress = resultMap.size / uniqueEmployers.length * 100;
    
    console.log(`  Websites: ${bw}/${batch.length} | E-Mails: ${be}/${batch.length} | Time: ${elapsed}s`);
    console.log(`  Progress: ${resultMap.size}/${uniqueEmployers.length} (${progress.toFixed(1)}%) | Total: ${totalWebsites}W ${totalEmails}E | ${totalElapsed}s`);
    
    await sleep(500);
  }
  
  // Final stats
  const finalResults = [...resultMap.values()];
  const wCount = finalResults.filter(r => r.website).length;
  const eCount = finalResults.filter(r => r.emails.length > 0).length;
  
  console.log(`\n=== COMPLETE ===`);
  console.log(`Processed: ${resultMap.size}/${uniqueEmployers.length}`);
  console.log(`Websites: ${wCount} (${(wCount/resultMap.size*100).toFixed(1)}%)`);
  console.log(`E-Mails: ${eCount} (${(eCount/resultMap.size*100).toFixed(1)}%)`);
  console.log(`Saved: ${RESULTS_FILE}`);
}

main().catch(console.error);
