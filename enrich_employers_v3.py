#!/usr/bin/env python3
"""
Employer Enrichment V3 - Python Edition
- Uses z-ai-web-dev-sdk web search ONLY for finding websites (rate-limited)
- Uses direct HTTP requests (requests/BeautifulSoup) for scraping emails
- Sequential processing with delays to respect rate limits
- Saves progress after every batch
"""

import json, time, re, os, sys
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin
from pathlib import Path

RESULTS_FILE = '/home/z/my-project/enrichment_results_v3.json'
INPUT_FILE = '/home/z/my-project/search_needed.json'

# Config
BATCH_SIZE = 25
SEARCH_DELAY = 4.0          # seconds between web search API calls
SCRAPE_DELAY = 0.8          # seconds between HTTP scrape requests
SCRAPE_TIMEOUT = 10         # seconds timeout for HTTP requests
MAX_CONTACT_URLS = 4        # max impressum/kontakt pages to try

EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')
UNWANTED_EMAILS = ['example.com', 'test.com', 'domain.com', 'mustermann', '.png', '.jpg', '.gif', '.svg', '.css', '.js', 'noreply', 'no-reply', 'mailer-daemon', 'postmaster']

JOB_BOARD_DOMAINS = ['indeed', 'stepstone', 'kununu', 'glassdoor', 'xing.com', 'linkedin.com', 'arbeitsagentur', 'stellenanzeigen', 'jobware', 'monster.de', 'meinestadt', 'ausbildung', 'jobtuple', 'stellenonline', 'jobsuche']

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
}

def is_valid_email(email):
    email = email.lower().strip()
    if len(email) < 6: return False
    return not any(u in email for u in UNWANTED_EMAILS)

def extract_emails_from_html(html, url=''):
    """Extract emails from HTML content using multiple strategies."""
    emails = set()
    
    # Strategy 1: Direct regex on full HTML
    for match in EMAIL_REGEX.finditer(html):
        email = match.group(0).lower().strip()
        if is_valid_email(email):
            emails.add(email)
    
    # Strategy 2: mailto: links
    mailto_pattern = re.compile(r'mailto:([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})', re.IGNORECASE)
    for match in mailto_pattern.finditer(html):
        email = match.group(1).lower().strip()
        if is_valid_email(email):
            emails.add(email)
    
    # Strategy 3: Common JS obfuscation patterns
    # [name, domain, tld] style
    js_patterns = [
        r"['\"]([a-zA-Z0-9._%+\-]+)['\"]\s*\+\s*['\"]@['\"]\s*\+\s*['\"]([a-zA-Z0-9.\-]+)['\"]",
        r"([a-zA-Z0-9._%+\-]+)\s*\(\s*at\s*\)\s*([a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})",
        r"([a-zA-Z0-9._%+\-]+)\s*\[\s*at\s*\]\s*([a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})",
        r"([a-zA-Z0-9._%+\-]+)\s*@\s*([a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})",
    ]
    
    for pattern in js_patterns:
        for match in re.finditer(pattern, html, re.IGNORECASE):
            try:
                email = f"{match.group(1)}@{match.group(2)}".lower().strip()
                if is_valid_email(email):
                    emails.add(email)
            except:
                pass
    
    # Strategy 4: HTML entity decoding
    try:
        soup = BeautifulSoup(html, 'html.parser')
        text = soup.get_text(separator=' ')
        for match in EMAIL_REGEX.finditer(text):
            email = match.group(0).lower().strip()
            if is_valid_email(email):
                emails.add(email)
    except:
        pass
    
    return list(emails)

def categorize_email(email):
    """Categorize email by purpose."""
    email = email.lower()
    bewerbung_keywords = ['bewerbung', 'karriere', 'career', 'recruiting', 'jobs', 'hr', 'talent', 'personal', 'bewerb']
    kontakt_keywords = ['kontakt', 'contact', 'info', 'impressum', 'office', 'mail']
    
    if any(k in email for k in bewerbung_keywords):
        return 'bewerbung'
    elif any(k in email for k in kontakt_keywords):
        return 'kontakt'
    return 'general'

def scrape_url(url):
    """Scrape a URL and return (html, emails)."""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=SCRAPE_TIMEOUT, allow_redirects=True, verify=False)
        if resp.status_code == 200:
            html = resp.text
            emails = extract_emails_from_html(html, url)
            return html, emails
    except Exception as e:
        pass
    return '', []

def get_contact_urls(base_url, html=''):
    """Generate likely contact/impressum URLs."""
    try:
        parsed = urlparse(base_url)
        domain_base = f"{parsed.scheme}://{parsed.hostname}"
    except:
        domain_base = base_url.rstrip('/')
    
    urls = [
        f"{domain_base}/impressum",
        f"{domain_base}/kontakt",
        f"{domain_base}/Impressum",
        f"{domain_base}/Kontakt",
        f"{domain_base}/de/impressum",
        f"{domain_base}/de/kontakt",
        f"{domain_base}/karriere",
        f"{domain_base}/career",
    ]
    
    # Also try to find contact links in the HTML
    if html:
        try:
            soup = BeautifulSoup(html, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href'].lower()
                link_text = a.get_text().lower()
                if any(k in href or k in link_text for k in ['impressum', 'kontakt', 'contact', 'legal', 'datenschutz']):
                    full_url = urljoin(base_url, a['href'])
                    if full_url not in urls:
                        urls.insert(0, full_url)  # Prioritize links found on the page
        except:
            pass
    
    return urls[:MAX_CONTACT_URLS + 2]  # Limit total

def enrich_employer_sequential(employer, zai=None):
    """Enrich a single employer - sequential with delays."""
    name = employer.get('name', '').strip()
    stadt = employer.get('stadt', '').strip()
    current_website = employer.get('current_website', '').strip()
    
    result = {
        'name': name,
        'website': current_website,
        'emails': [],
        'bewerbung_email': '',
        'kontakt_email': '',
        'scraped_from': [],
        'source': ''
    }
    
    # Step 1: Find website if missing (using API)
    if not current_website and zai:
        query = f'"{name}" {stadt} official website' if stadt else f'"{name}" Germany official website'
        try:
            search_results = zai.functions.invoke("web_search", {"query": query, "num": 5})
            if search_results:
                for sr in search_results:
                    url = sr.get('url', '') or sr.get('host_name', '')
                    if url and not any(d in url.lower() for d in JOB_BOARD_DOMAINS):
                        if not url.startswith('http'):
                            url = f"https://{url}"
                        result['website'] = url
                        result['source'] = 'web_search'
                        break
                else:
                    if search_results:
                        first = search_results[0]
                        url = first.get('url', '') or first.get('host_name', '')
                        if url:
                            if not url.startswith('http'):
                                url = f"https://{url}"
                            result['website'] = url
                            result['source'] = 'web_search_fallback'
        except Exception as e:
            result['source'] = f'search_error: {str(e)[:50]}'
        
        time.sleep(SEARCH_DELAY)
    
    # Step 2: Scrape for emails (direct HTTP - no API)
    website = result['website']
    if website:
        all_emails = []
        
        # First try main page
        html, emails = scrape_url(website)
        all_emails.extend(emails)
        result['scraped_from'].append(website)
        time.sleep(SCRAPE_DELAY)
        
        # If no emails on main page, try impressum/kontakt
        if not all_emails and html:
            contact_urls = get_contact_urls(website, html)
            for contact_url in contact_urls:
                if contact_url == website:
                    continue  # Already scraped
                _, emails = scrape_url(contact_url)
                all_emails.extend(emails)
                result['scraped_from'].append(contact_url)
                if emails:
                    break  # Stop after finding emails
                time.sleep(SCRAPE_DELAY)
        elif not all_emails:
            # No HTML from main page, try common contact pages anyway
            contact_urls = get_contact_urls(website)
            for contact_url in contact_urls[:3]:
                _, emails = scrape_url(contact_url)
                all_emails.extend(emails)
                result['scraped_from'].append(contact_url)
                if emails:
                    break
                time.sleep(SCRAPE_DELAY)
        
        # Deduplicate and categorize
        unique_emails = list(dict.fromkeys([e.lower() for e in all_emails if is_valid_email(e)]))
        result['emails'] = unique_emails
        
        # Categorize
        for email in unique_emails:
            cat = categorize_email(email)
            if cat == 'bewerbung' and not result['bewerbung_email']:
                result['bewerbung_email'] = email
            elif cat == 'kontakt' and not result['kontakt_email']:
                result['kontakt_email'] = email
        
        # Fallback: first email as kontakt
        if not result['kontakt_email'] and unique_emails:
            result['kontakt_email'] = unique_emails[0]
    
    return result

def load_results():
    if os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, 'r') as f:
            return json.load(f)
    return []

def save_results(results):
    with open(RESULTS_FILE, 'w') as f:
        json.dump(results, f, ensure_ascii=False)

def main():
    start_batch = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    max_batches = int(sys.argv[2]) if len(sys.argv) > 2 else 5
    
    print(f"=== Employer Enrichment V3 (Python) ===")
    print(f"Batch {start_batch}, max {max_batches} batches")
    
    # Load employers
    with open(INPUT_FILE, 'r') as f:
        employers = json.load(f)
    
    # Deduplicate by name
    seen = {}
    unique = []
    for e in employers:
        key = e['name'].lower().strip()
        if key not in seen:
            seen[key] = e
            unique.append(e)
    
    print(f"{len(employers)} entries → {len(unique)} unique names")
    
    # Load existing results
    existing = load_results()
    result_map = {r['name'].lower().strip(): r for r in existing}
    print(f"Already enriched: {len(result_map)}")
    
    # Filter to only those not yet done
    pending = [e for e in unique if e['name'].lower().strip() not in result_map]
    print(f"Pending: {len(pending)}")
    
    # Create batches
    batches = [pending[i:i+BATCH_SIZE] for i in range(0, len(pending), BATCH_SIZE)]
    to_process = batches[start_batch:start_batch + max_batches]
    print(f"Processing batches {start_batch+1}-{start_batch+len(to_process)} of {len(batches)}")
    
    # Init z-ai for web search
    # We'll use the SDK via subprocess to avoid import issues
    import subprocess
    
    total_w = 0
    total_e = 0
    start_time = time.time()
    
    for bi, batch in enumerate(to_process):
        batch_num = start_batch + bi + 1
        batch_start = time.time()
        print(f"\n--- Batch {batch_num} ({len(batch)} employers) ---")
        
        for ei, employer in enumerate(batch):
            name = employer['name']
            # Use z-ai CLI for web search
            result = enrich_employer_cli(employer)
            
            key = name.lower().strip()
            result_map[key] = result
            
            if result.get('website'):
                total_w += 1
            if result.get('emails'):
                total_e += 1
            
            # Print progress
            status = "W" if result.get('website') else "-"
            status += "E" if result.get('emails') else "-"
            email_preview = result['emails'][0] if result.get('emails') else ''
            print(f"  [{ei+1}/{len(batch)}] {status} {name[:45]:<45} | {result.get('website', '')[:40]} | {email_preview}")
        
        # Save after each batch
        save_results(list(result_map.values()))
        
        elapsed = time.time() - batch_start
        total_elapsed = time.time() - start_time
        progress = len(result_map) / len(unique) * 100
        print(f"  Batch done in {elapsed:.0f}s | Progress: {len(result_map)}/{len(unique)} ({progress:.1f}%) | Total: {total_w}W {total_e}E | {total_elapsed:.0f}s total")
    
    # Final stats
    final = list(result_map.values())
    w = sum(1 for r in final if r.get('website'))
    e = sum(1 for r in final if r.get('emails'))
    print(f"\n=== COMPLETE ===")
    print(f"Processed: {len(result_map)}/{len(unique)}")
    print(f"Websites: {w} ({w/len(result_map)*100:.1f}%)")
    print(f"E-Mails: {e} ({e/len(result_map)*100:.1f}%)")

def enrich_employer_cli(employer):
    """Enrich employer using CLI tools for web search."""
    name = employer.get('name', '').strip()
    stadt = employer.get('stadt', '').strip()
    current_website = employer.get('current_website', '').strip()
    
    result = {
        'name': name,
        'website': current_website,
        'emails': [],
        'bewerbung_email': '',
        'kontakt_email': '',
        'scraped_from': [],
        'source': ''
    }
    
    # Step 1: Find website if missing
    if not current_website:
        query = f'"{name}" {stadt} official website' if stadt else f'"{name}" Germany official website'
        try:
            # Use z-ai-web-dev-sdk via Node.js subprocess
            search_script = f'''
import ZAI from 'z-ai-web-dev-sdk';
async function main() {{
  const zai = await ZAI.create();
  const result = await zai.functions.invoke("web_search", {{query: "{query.replace('"', '\\"')}", num: 5}});
  console.log(JSON.stringify(result));
}}
main().catch(e => console.error(e.message));
'''
            tmp_script = '/tmp/search_employer.mjs'
            with open(tmp_script, 'w') as f:
                f.write(search_script)
            
            proc = subprocess.run(['node', tmp_script], capture_output=True, text=True, timeout=30)
            if proc.returncode == 0 and proc.stdout.strip():
                search_results = json.loads(proc.stdout.strip())
                if search_results:
                    for sr in search_results:
                        url = sr.get('url', '') or sr.get('host_name', '')
                        if url and not any(d in url.lower() for d in JOB_BOARD_DOMAINS):
                            if not url.startswith('http'):
                                url = f"https://{url}"
                            result['website'] = url
                            result['source'] = 'web_search'
                            break
                    else:
                        if search_results:
                            first = search_results[0]
                            url = first.get('url', '') or first.get('host_name', '')
                            if url:
                                if not url.startswith('http'):
                                    url = f"https://{url}"
                                result['website'] = url
                                result['source'] = 'web_search_fallback'
        except Exception as e:
            result['source'] = f'search_error: {str(e)[:50]}'
        
        time.sleep(SEARCH_DELAY)
    
    # Step 2: Scrape for emails (direct HTTP)
    website = result['website']
    if website:
        all_emails = []
        
        # Try main page
        html, emails = scrape_url(website)
        all_emails.extend(emails)
        result['scraped_from'].append(website)
        time.sleep(SCRAPE_DELAY)
        
        # If no emails, try impressum/kontakt
        if not all_emails:
            contact_urls = get_contact_urls(website, html)
            for contact_url in contact_urls[:4]:
                if contact_url == website:
                    continue
                _, emails = scrape_url(contact_url)
                all_emails.extend(emails)
                result['scraped_from'].append(contact_url)
                if emails:
                    break
                time.sleep(SCRAPE_DELAY)
        
        # Deduplicate and categorize
        unique_emails = list(dict.fromkeys([e.lower() for e in all_emails if is_valid_email(e)]))
        result['emails'] = unique_emails
        
        for email in unique_emails:
            cat = categorize_email(email)
            if cat == 'bewerbung' and not result['bewerbung_email']:
                result['bewerbung_email'] = email
            elif cat == 'kontakt' and not result['kontakt_email']:
                result['kontakt_email'] = email
        
        if not result['kontakt_email'] and unique_emails:
            result['kontakt_email'] = unique_emails[0]
    
    return result

if __name__ == '__main__':
    # Suppress SSL warnings
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
    main()
