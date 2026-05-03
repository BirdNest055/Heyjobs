#!/usr/bin/env node
/**
 * Employer Enrichment Script
 * - Searches for employer websites using z-ai-web-dev-sdk web search
 * - Scrapes contact/impressum pages for email addresses
 * - Works in parallel batches with rate limiting
 * - Saves progress incrementally
 */

import ZAI from 'z-ai-web-dev-sdk';
import { readFileSync, writeFileSync, existsSync } from 'fs';

const PROGRESS_FILE = '/home/z/my-project/enrichment_progress.json';
const RESULTS_FILE = '/home/z/my-project/enrichment_results.json';
const INPUT_FILE = '/home/z/my-project/search_needed.json';

// Rate limiting config
const BATCH_SIZE = 20;           // employers per batch
const SEARCH_CONCURRENCY = 5;    // parallel searches
const SCRAPE_CONCURRENCY = 3;    // parallel scrapes
const SEARCH_DELAY = 200;        // ms between search batches
const SCRAPE_DELAY = 500;        // ms between scrape requests
const MAX_RETRIES = 2;

// Email regex - comprehensive
const EMAIL_REGEX = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;

// Keywords for Bewerbung/Kontakt pages
const CONTACT_KEYWORDS = ['kontakt', 'contact', 'impressum', 'legal', 'bewerbung', 'career', 'jobs', 'karriere', 'ueber-uns', 'about'];

async function sleep(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function searchEmployerWebsite(zai, employerName, stadt) {
  const query = stadt 
    ? `"${employerName}" ${stadt} website official`
    : `"${employerName}" website official`;
  
  try {
    const result = await zai.functions.invoke("web_search", {
      query: query,
      num: 5
    });
    
    if (result && result.length > 0) {
      // Filter for likely official websites (not job boards)
      const jobBoardDomains = ['indeed.com', 'stepstone.de', 'kununu.com', 'glassdoor', 'xing.com', 'linkedin.com', 'arbeitsagentur', 'stellenanzeigen', 'jobware', 'monster.de', 'meinestadt', 'ausbildung'];
      
      for (const r of result) {
        const url = r.url || r.host_name || '';
        const isJobBoard = jobBoardDomains.some(d => url.toLowerCase().includes(d));
        if (!isJobBoard && url) {
          return {
            url: url.startsWith('http') ? url : `https://${url}`,
            name: r.name || '',
            snippet: r.snippet || '',
            host: r.host_name || ''
          };
        }
      }
      // If all are job boards, return first non-job-board or first result
      const first = result[0];
      return {
        url: (first.url || '').startsWith('http') ? first.url : `https://${first.url}`,
        name: first.name || '',
        snippet: first.snippet || '',
        host: first.host_name || ''
      };
    }
  } catch (e) {
    // Silent fail for individual searches
  }
  return null;
}

async function scrapePageForEmails(zai, url) {
  try {
    const result = await zai.functions.invoke("web_reader", {
      url: url
    });
    
    if (!result) return [];
    
    const html = result.html || result.content || '';
    const text = result.text || html;
    
    // Find all emails
    const emails = [...text.matchAll(EMAIL_REGEX)].map(m => m[0].toLowerCase());
    
    // Filter out generic/unwanted emails
    const unwantedPatterns = [
      'example.com', 'test.com', 'domain.com', 'email.de',
      'mustermann', 'noreply', 'no-reply', 'mailer-daemon',
      'postmaster', 'webmaster@', 'admin@', 'root@',
      '.png', '.jpg', '.gif', '.svg', '.css', '.js'
    ];
    
    const filtered = emails.filter(e => {
      return !unwantedPatterns.some(p => e.includes(p)) && e.length > 5;
    });
    
    // Prioritize bewerbung/karriere/kontakt emails
    const prioritized = [];
    const other = [];
    const priorityKeywords = ['bewerbung', 'karriere', 'career', 'recruiting', 'jobs', 'kontakt', 'contact', 'hr', 'talent'];
    
    for (const email of [...new Set(filtered)]) {
      if (priorityKeywords.some(k => email.includes(k))) {
        prioritized.push(email);
      } else {
        other.push(email);
      }
    }
    
    return [...prioritized, ...other];
  } catch (e) {
    return [];
  }
}

async function enrichSingleEmployer(zai, employer) {
  const { name, stadt, current_website, current_email } = employer;
  const result = { name, website: current_website || '', emails: [], scraped_pages: [] };
  
  // Step 1: Find website if missing
  let website = current_website;
  if (!website) {
    const searchResult = await searchEmployerWebsite(zai, name, stadt);
    if (searchResult) {
      website = searchResult.url;
      result.website = website;
      result.search_snippet = searchResult.snippet;
    }
    await sleep(SEARCH_DELAY);
  }
  
  // Step 2: Scrape for emails if website found
  if (website && !current_email) {
    // Normalize website URL
    let baseUrl = website.replace(/\/$/, '');
    // Remove trailing paths like /kontakt for the base
    try {
      const urlObj = new URL(baseUrl);
      baseUrl = `${urlObj.protocol}//${urlObj.hostname}`;
    } catch(e) {}
    
    // Try main page first
    const mainEmails = await scrapePageForEmails(zai, website);
    result.emails.push(...mainEmails);
    result.scraped_pages.push(website);
    await sleep(SCRAPE_DELAY);
    
    // If no emails found, try contact/impressum pages
    if (result.emails.length === 0) {
      const contactPaths = ['/kontakt', '/impressum', '/karriere', '/career', '/contact'];
      
      for (const path of contactPaths) {
        try {
          const contactUrl = `${baseUrl}${path}`;
          const contactEmails = await scrapePageForEmails(zai, contactUrl);
          result.emails.push(...contactEmails);
          result.scraped_pages.push(contactUrl);
          await sleep(SCRAPE_DELAY);
          
          if (result.emails.length > 0) break; // Stop if we found emails
        } catch(e) {
          continue;
        }
      }
    }
    
    // Deduplicate emails
    result.emails = [...new Set(result.emails)];
  }
  
  return result;
}

// Process batch with limited concurrency
async function processBatch(zai, batch, batchNum) {
  const results = [];
  
  for (let i = 0; i < batch.length; i += SEARCH_CONCURRENCY) {
    const chunk = batch.slice(i, i + SEARCH_CONCURRENCY);
    const promises = chunk.map(employer => 
      enrichSingleEmployer(zai, employer).catch(e => ({
        name: employer.name,
        website: employer.current_website || '',
        emails: [],
        error: e.message
      }))
    );
    
    const chunkResults = await Promise.all(promises);
    results.push(...chunkResults);
    
    if (i + SEARCH_CONCURRENCY < batch.length) {
      await sleep(SEARCH_DELAY);
    }
  }
  
  return results;
}

function loadProgress() {
  if (existsSync(PROGRESS_FILE)) {
    return JSON.parse(readFileSync(PROGRESS_FILE, 'utf-8'));
  }
  return { completedIndices: [], results: {} };
}

function saveProgress(progress) {
  writeFileSync(PROGRESS_FILE, JSON.stringify(progress, null, 0));
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
  const maxBatches = parseInt(args[1] || '999');
  
  console.log(`Starting enrichment from batch ${startBatch}, max ${maxBatches} batches`);
  
  // Load employers
  const employers = JSON.parse(readFileSync(INPUT_FILE, 'utf-8'));
  
  // Deduplicate by name - only search each unique name once
  const seen = new Map();
  const uniqueEmployers = [];
  for (const e of employers) {
    const key = e.name.toLowerCase().trim();
    if (!seen.has(key)) {
      seen.set(key, e);
      uniqueEmployers.push(e);
    }
  }
  console.log(`Total: ${employers.length} entries, ${uniqueEmployers.length} unique names`);
  
  // Load progress
  const progress = loadProgress();
  const existingResults = loadResults();
  const resultMap = new Map();
  for (const r of existingResults) {
    resultMap.set(r.name.toLowerCase().trim(), r);
  }
  
  // Create zai instance
  const zai = await ZAI.create();
  
  // Split into batches
  const batches = [];
  for (let i = 0; i < uniqueEmployers.length; i += BATCH_SIZE) {
    batches.push(uniqueEmployers.slice(i, i + BATCH_SIZE));
  }
  
  console.log(`Total batches: ${batches.length}, processing ${startBatch} to ${Math.min(startBatch + maxBatches - 1, batches.length - 1)}`);
  
  let totalFound = 0;
  let totalEmails = 0;
  
  for (let b = startBatch; b < Math.min(startBatch + maxBatches, batches.length); b++) {
    const batch = batches[b];
    
    // Skip already completed
    const pendingBatch = batch.filter(e => !resultMap.has(e.name.toLowerCase().trim()));
    if (pendingBatch.length === 0) {
      console.log(`Batch ${b+1}/${batches.length}: Skipped (all done)`);
      continue;
    }
    
    console.log(`\n--- Batch ${b+1}/${batches.length} (${pendingBatch.length} employers) ---`);
    
    const batchResults = await processBatch(zai, pendingBatch, b);
    
    // Save results
    let batchWebsites = 0;
    let batchEmails = 0;
    for (const r of batchResults) {
      resultMap.set(r.name.toLowerCase().trim(), r);
      if (r.website) batchWebsites++;
      if (r.emails.length > 0) batchEmails++;
    }
    
    totalFound += batchWebsites;
    totalEmails += batchEmails;
    
    // Save after each batch
    saveResults([...resultMap.values()]);
    
    console.log(`  Websites gefunden: ${batchWebsites}/${pendingBatch.length}`);
    console.log(`  E-Mails gefunden: ${batchEmails}/${pendingBatch.length}`);
    console.log(`  Gesamt Fortschritt: ${resultMap.size}/${uniqueEmployers.length} (${(resultMap.size/uniqueEmployers.length*100).toFixed(1)}%)`);
    console.log(`  Gesamt Websites: ${totalFound}, Gesamt E-Mails: ${totalEmails}`);
    
    // Brief pause between batches
    await sleep(300);
  }
  
  // Final save
  const finalResults = [...resultMap.values()];
  saveResults(finalResults);
  
  console.log(`\n=== FERTIG ===`);
  console.log(`Verarbeitet: ${resultMap.size}/${uniqueEmployers.length} einzigartige Arbeitgeber`);
  console.log(`Websites gefunden: ${finalResults.filter(r => r.website).length}`);
  console.log(`E-Mails gefunden: ${finalResults.filter(r => r.emails.length > 0).length}`);
  console.log(`Gespeichert in: ${RESULTS_FILE}`);
}

main().catch(console.error);
