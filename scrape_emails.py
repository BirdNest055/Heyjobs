#!/usr/bin/env python3
"""
Step 2: Scrape emails from employer websites
Uses direct HTTP requests - no API rate limits!
Reads website URLs from search results and scrapes impressum/kontakt pages
"""

import json, time, re, os, sys
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SEARCH_RESULTS = '/home/z/my-project/website_search_results.json'
EMAIL_RESULTS = '/home/z/my-project/email_results.json'

SCRAPE_DELAY = 0.5
SCRAPE_TIMEOUT = 8
MAX_CONTACT_URLS = 5

EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')
UNWANTED = ['example.com', 'test.com', 'domain.com', 'mustermann', '.png', '.jpg', '.gif', '.svg', '.css', '.js', 'noreply', 'no-reply', 'mailer-daemon', 'postmaster', 'sentry', 'wixpress', 'googleapis', 'cloudfront', 'amazonaws']

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
}

def is_valid_email(email):
    email = email.lower().strip()
    if len(email) < 6: return False
    return not any(u in email for u in UNWANTED)

def extract_emails(html):
    """Multi-strategy email extraction."""
    emails = set()
    
    # 1. Direct regex
    for m in EMAIL_REGEX.finditer(html):
        e = m.group(0).lower().strip()
        if is_valid_email(e): emails.add(e)
    
    # 2. mailto: links
    for m in re.finditer(r'mailto:([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})', html, re.I):
        e = m.group(1).lower().strip()
        if is_valid_email(e): emails.add(e)
    
    # 3. (at) / [at] obfuscation
    for m in re.finditer(r'([a-zA-Z0-9._%+\-]+)\s*[\(\[]\s*at\s*[\)\]]\s*([a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})', html, re.I):
        try:
            e = f"{m.group(1)}@{m.group(2)}".lower().strip()
            if is_valid_email(e): emails.add(e)
        except: pass
    
    # 4. BeautifulSoup text extraction
    try:
        soup = BeautifulSoup(html, 'html.parser')
        # Remove script/style tags
        for tag in soup(['script', 'style']):
            tag.decompose()
        text = soup.get_text(separator=' ')
        for m in EMAIL_REGEX.finditer(text):
            e = m.group(0).lower().strip()
            if is_valid_email(e): emails.add(e)
    except: pass
    
    return list(emails)

def categorize_email(email):
    email = email.lower()
    bewerbung_kw = ['bewerbung', 'karriere', 'career', 'recruiting', 'jobs', 'hr', 'talent', 'personal', 'bewerb']
    kontakt_kw = ['kontakt', 'contact', 'info', 'impressum', 'office', 'mail', 'service']
    if any(k in email for k in bewerbung_kw): return 'bewerbung'
    if any(k in email for k in kontakt_kw): return 'kontakt'
    return 'general'

def get_contact_urls(base_url, html=''):
    try:
        p = urlparse(base_url)
        domain_base = f"{p.scheme}://{p.hostname}"
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
        f"{domain_base}/about",
        f"{domain_base}/ueber-uns",
    ]
    
    # Find links in HTML
    if html:
        try:
            soup = BeautifulSoup(html, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href'].lower()
                text = a.get_text().lower()
                if any(k in href or k in text for k in ['impressum', 'kontakt', 'contact', 'legal']):
                    full = urljoin(base_url, a['href'])
                    if full not in urls:
                        urls.insert(0, full)
        except: pass
    
    return urls[:MAX_CONTACT_URLS + 2]

def scrape_url(url):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=SCRAPE_TIMEOUT, allow_redirects=True, verify=False)
        if resp.status_code == 200:
            return resp.text, extract_emails(resp.text)
    except: pass
    return '', []

def main():
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    count = int(sys.argv[2]) if len(sys.argv) > 2 else 100
    
    print(f"=== Email Scraper ===")
    print(f"Processing {start} to {start+count-1}")
    
    # Load search results
    with open(SEARCH_RESULTS, 'r') as f:
        search_data = json.load(f)
    
    employers_with_website = {k: v for k, v in search_data.items() if v.get('website')}
    print(f"Employers with website: {len(employers_with_website)}")
    
    # Load existing email results
    email_results = {}
    if os.path.exists(EMAIL_RESULTS):
        with open(EMAIL_RESULTS, 'r') as f:
            email_results = json.load(f)
    print(f"Already scraped: {len(email_results)}")
    
    # Get slice of employers to scrape
    items = list(employers_with_website.items())
    slice_items = items[start:start+count]
    pending = [(k, v) for k, v in slice_items if k not in email_results]
    print(f"Pending: {len(pending)}")
    
    total_emails = 0
    start_time = time.time()
    
    for i, (key, emp) in enumerate(pending):
        website = emp['website']
        name = emp['name']
        
        all_emails = []
        scraped_from = []
        
        # Scrape main page first
        html, emails = scrape_url(website)
        all_emails.extend(emails)
        scraped_from.append(website)
        time.sleep(SCRAPE_DELAY)
        
        # If no emails, try impressum/kontakt
        if not all_emails:
            contact_urls = get_contact_urls(website, html)
            for curl in contact_urls[:4]:
                if curl == website: continue
                _, cemails = scrape_url(curl)
                all_emails.extend(cemails)
                scraped_from.append(curl)
                if cemails: break
                time.sleep(SCRAPE_DELAY)
        
        # Deduplicate
        unique = list(dict.fromkeys([e.lower() for e in all_emails if is_valid_email(e)]))
        
        # Categorize
        bewerbung = next((e for e in unique if categorize_email(e) == 'bewerbung'), '')
        kontakt = next((e for e in unique if categorize_email(e) == 'kontakt'), '')
        if not kontakt and unique:
            kontakt = unique[0]
        
        email_results[key] = {
            'name': name,
            'website': website,
            'emails': unique,
            'bewerbung_email': bewerbung,
            'kontakt_email': kontakt,
            'scraped_from': scraped_from,
        }
        
        total_emails += len(unique)
        
        status = f"{len(unique)}E" if unique else "0E"
        preview = unique[0] if unique else ''
        print(f"  [{i+1}/{len(pending)}] {status:>3} {name[:50]:<50} | {preview}")
        
        # Save every 25
        if (i + 1) % 25 == 0:
            with open(EMAIL_RESULTS, 'w') as f:
                json.dump(email_results, f, ensure_ascii=False)
            elapsed = time.time() - start_time
            print(f"  → Saved {len(email_results)} results | {elapsed:.0f}s elapsed")
    
    # Final save
    with open(EMAIL_RESULTS, 'w') as f:
        json.dump(email_results, f, ensure_ascii=False)
    
    with_emails = sum(1 for v in email_results.values() if v.get('emails'))
    print(f"\n=== DONE ===")
    print(f"Scraped: {len(pending)} | With emails: {with_emails} | Total emails: {total_emails}")
    print(f"Total results: {len(email_results)}")

if __name__ == '__main__':
    main()
