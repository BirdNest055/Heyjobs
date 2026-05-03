#!/usr/bin/env python3
"""
Re-scrape employers that have websites but no emails.
Uses more aggressive scraping - tries more contact page variants.
"""

import json, time, re, os, sys
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

RESULTS_FILE = '/home/z/my-project/website_search_results.json'
RESCRAPE_FILE = '/home/z/my-project/rescrape_emails.json'

SCRAPE_DELAY = 0.3
SCRAPE_TIMEOUT = 10
MAX_URLS = 8

EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')
UNWANTED = ['example.com', 'test.com', 'domain.com', 'mustermann', '.png', '.jpg', '.gif', '.svg', '.css', '.js', 'noreply', 'no-reply', 'mailer-daemon', 'postmaster', 'sentry', 'wixpress', 'googleapis', 'cloudfront', 'amazonaws', 'gravatar', 'wordpress', 'schema.org', 'w3.org', '2x.webp', '3x.webp', 'email.protected', 'beispiel.de', 'nutzer@', 'cookielaw', 'sentry.io']

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
    emails = set()
    for m in EMAIL_REGEX.finditer(html):
        e = m.group(0).lower().strip()
        if is_valid_email(e): emails.add(e)
    for m in re.finditer(r'mailto:([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})', html, re.I):
        e = m.group(1).lower().strip()
        if is_valid_email(e): emails.add(e)
    for m in re.finditer(r'([a-zA-Z0-9._%+\-]+)\s*[\(\[]\s*at\s*[\)\]]\s*([a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})', html, re.I):
        try:
            e = f"{m.group(1)}@{m.group(2)}".lower().strip()
            if is_valid_email(e): emails.add(e)
        except: pass
    try:
        soup = BeautifulSoup(html, 'html.parser')
        for tag in soup(['script', 'style', 'noscript']):
            tag.decompose()
        text = soup.get_text(separator=' ')
        for m in EMAIL_REGEX.finditer(text):
            e = m.group(0).lower().strip()
            if is_valid_email(e): emails.add(e)
    except: pass
    return list(emails)

def categorize_email(email):
    email = email.lower()
    if any(k in email for k in ['bewerbung', 'karriere', 'career', 'recruiting', 'jobs', 'hr', 'talent', 'personal']): return 'bewerbung'
    if any(k in email for k in ['kontakt', 'contact', 'info', 'impressum', 'office']): return 'kontakt'
    return 'general'

def scrape_url(url):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=SCRAPE_TIMEOUT, allow_redirects=True, verify=False)
        if resp.status_code == 200:
            return resp.text, extract_emails(resp.text)
    except: pass
    return '', []

def get_all_contact_urls(base_url, html=''):
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
        f"{domain_base}/de/karriere",
        f"{domain_base}/karriere",
        f"{domain_base}/career",
        f"{domain_base}/about",
        f"{domain_base}/ueber-uns",
        f"{domain_base}/contact",
        f"{domain_base}/legal",
        f"{domain_base}/datenschutz",
    ]
    
    if html:
        try:
            soup = BeautifulSoup(html, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href'].lower()
                text = a.get_text().lower()
                if any(k in href or k in text for k in ['impressum', 'kontakt', 'contact', 'legal', 'datenschutz', 'karriere', 'career']):
                    full = urljoin(base_url, a['href'])
                    if full not in urls:
                        urls.insert(0, full)
        except: pass
    
    return urls[:MAX_URLS]

def main():
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    count = int(sys.argv[2]) if len(sys.argv) > 2 else 30
    
    print(f"=== Re-Scrape Email V2 ===")
    
    with open(RESCRAPE_FILE, 'r') as f:
        rescrape_list = json.load(f)
    
    with open(RESULTS_FILE, 'r') as f:
        results = json.load(f)
    
    slice_items = rescrape_list[start:start+count]
    print(f"Processing {len(slice_items)} employers")
    
    found = 0
    for i, emp in enumerate(slice_items):
        name = emp.get('name', '')
        website = emp.get('website', '')
        key = name.lower().strip()
        
        all_emails = []
        scraped_from = []
        
        # Try main page
        html, emails = scrape_url(website)
        all_emails.extend(emails)
        scraped_from.append(website)
        time.sleep(SCRAPE_DELAY)
        
        # Try all contact pages
        if not all_emails:
            contact_urls = get_all_contact_urls(website, html)
            for curl in contact_urls[:6]:
                if curl == website: continue
                _, cemails = scrape_url(curl)
                all_emails.extend(cemails)
                scraped_from.append(curl)
                if cemails: break
                time.sleep(SCRAPE_DELAY)
        
        if all_emails:
            found += 1
            unique = list(dict.fromkeys([e.lower() for e in all_emails if is_valid_email(e)]))
            
            bewerbung = next((e for e in unique if categorize_email(e) == 'bewerbung'), '')
            kontakt = next((e for e in unique if categorize_email(e) == 'kontakt'), '')
            if not kontakt and unique: kontakt = unique[0]
            
            # Update results
            if key in results:
                results[key]['emails'] = unique
                results[key]['bewerbung_email'] = bewerbung
                results[key]['kontakt_email'] = kontakt
                results[key]['scraped_from'] = scraped_from
            
            print(f"  [{i+1}/{len(slice_items)}] ✓ {name[:45]:<45} | {unique[0] if unique else ''}")
        else:
            print(f"  [{i+1}/{len(slice_items)}] ✗ {name[:45]:<45} | no email found")
        
        # Save every 10
        if (i + 1) % 10 == 0:
            with open(RESULTS_FILE, 'w') as f:
                json.dump(results, f, ensure_ascii=False)
    
    with open(RESULTS_FILE, 'w') as f:
        json.dump(results, f, ensure_ascii=False)
    
    print(f"\n=== DONE: Found {found}/{len(slice_items)} emails ===")

if __name__ == '__main__':
    main()
