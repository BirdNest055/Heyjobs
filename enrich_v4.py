#!/usr/bin/env python3
"""
Employer Enrichment V4 - No API needed!
- Uses DuckDuckGo HTML search (via requests) for finding websites
- Uses direct HTTP scraping for email extraction from impressum/kontakt
- No z-ai-web-dev-sdk, no rate limits!
"""

import json, time, re, os, sys, random
import requests
from bs4 import BeautifulSoup
from urllib.parse import urlparse, urljoin, quote_plus
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

SEARCH_RESULTS = '/home/z/my-project/website_search_results.json'
EMAIL_RESULTS = '/home/z/my-project/email_results.json'
INPUT_FILE = '/home/z/my-project/search_needed.json'

# Config
DDG_DELAY = 1.5           # seconds between DDG searches
SCRAPE_DELAY = 0.3         # seconds between HTTP scrapes
SCRAPE_TIMEOUT = 8
MAX_CONTACT_URLS = 5

EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')
UNWANTED = ['example.com', 'test.com', 'domain.com', 'mustermann', '.png', '.jpg', '.gif', '.svg', '.css', '.js', 'noreply', 'no-reply', 'mailer-daemon', 'postmaster', 'sentry', 'wixpress', 'googleapis', 'cloudfront', 'amazonaws', 'gravatar', 'wordpress', 'schema.org', 'w3.org', '2x.webp', '3x.webp', 'email.protected', 'e-mail', 'example@', 'your.', 'sentry.io', 'cookielaw']

JOB_BOARDS = ['indeed', 'stepstone', 'kununu', 'glassdoor', 'xing.com', 'linkedin.com', 'arbeitsagentur', 'stellenanzeigen', 'jobware', 'monster.de', 'meinestadt', 'ausbildung', 'jobtuple', 'stellenonline', 'jobsuche', 'jobtraffic', 'stellenwerk']

USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:120.0) Gecko/20100101 Firefox/120.0',
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.1 Safari/605.1.15',
]

def get_headers():
    return {
        'User-Agent': random.choice(USER_AGENTS),
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate',
    }

def is_valid_email(email):
    email = email.lower().strip()
    if len(email) < 6: return False
    return not any(u in email for u in UNWANTED)

def is_job_board(url):
    return any(d in url.lower() for d in JOB_BOARDS)

def search_ddg(query, num_results=8):
    """Search DuckDuckGo HTML version - no API needed!"""
    results = []
    try:
        url = f"https://html.duckduckgo.com/html/?q={quote_plus(query)}"
        resp = requests.get(url, headers=get_headers(), timeout=15, verify=False)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'html.parser')
            for result in soup.find_all('div', class_='result'):
                title_elem = result.find('a', class_='result__a')
                snippet_elem = result.find('a', class_='result__snippet')
                if title_elem:
                    href = title_elem.get('href', '')
                    # DDG redirects through their own URL, extract the actual URL
                    if 'uddg=' in href:
                        actual_url = href.split('uddg=')[1].split('&')[0]
                        from urllib.parse import unquote
                        actual_url = unquote(actual_url)
                    else:
                        actual_url = href
                    
                    results.append({
                        'url': actual_url,
                        'title': title_elem.get_text(strip=True),
                        'snippet': snippet_elem.get_text(strip=True) if snippet_elem else '',
                    })
    except Exception as e:
        pass  # Silent fail
    return results

def search_google_light(query, num_results=5):
    """Fallback: Google search via requests"""
    results = []
    try:
        url = f"https://www.google.com/search?q={quote_plus(query)}&num={num_results}&hl=de"
        headers = get_headers()
        headers['Accept'] = 'text/html'
        resp = requests.get(url, headers=headers, timeout=10, verify=False)
        if resp.status_code == 200:
            soup = BeautifulSoup(resp.text, 'html.parser')
            for g in soup.find_all('div'):
                links = g.find_all('a', href=True)
                for link in links:
                    href = link['href']
                    if href.startswith('/url?q='):
                        actual = href.split('/url?q=')[1].split('&')[0]
                        from urllib.parse import unquote
                        actual = unquote(actual)
                        if actual.startswith('http') and 'google' not in actual:
                            results.append({
                                'url': actual,
                                'title': link.get_text(strip=True),
                            })
    except:
        pass
    return results

def find_website(name, stadt=''):
    """Find employer website using search engines."""
    queries = []
    if stadt:
        queries.append(f"{name} {stadt} website")
    queries.append(f"{name} Germany official site")
    queries.append(f'"{name}" homepage')
    
    for query in queries:
        results = search_ddg(query)
        
        if results:
            # Try to find non-job-board result
            for r in results:
                url = r.get('url', '')
                if url and url.startswith('http') and not is_job_board(url):
                    return url, 'ddg'
            
            # Fallback to first result
            if results[0].get('url', '').startswith('http'):
                return results[0]['url'], 'ddg_fallback'
        
        time.sleep(1)
    
    return '', 'not_found'

def extract_emails(html):
    """Multi-strategy email extraction from HTML."""
    emails = set()
    
    # 1. Direct regex on raw HTML
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
    
    # 4. Clean text extraction
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
    bewerbung_kw = ['bewerbung', 'karriere', 'career', 'recruiting', 'jobs', 'hr', 'talent', 'personal', 'bewerb']
    kontakt_kw = ['kontakt', 'contact', 'info', 'impressum', 'office', 'mail', 'service']
    if any(k in email for k in bewerbung_kw): return 'bewerbung'
    if any(k in email for k in kontakt_kw): return 'kontakt'
    return 'general'

def get_contact_urls(base_url, html=''):
    """Generate impressum/kontakt URLs (required by German law!)."""
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
                if any(k in href or k in text for k in ['impressum', 'kontakt', 'contact', 'legal', 'datenschutz']):
                    full = urljoin(base_url, a['href'])
                    if full not in urls:
                        urls.insert(0, full)
        except: pass
    
    return urls[:MAX_CONTACT_URLS + 2]

def scrape_url(url):
    """Scrape a URL and return (html, emails)."""
    try:
        resp = requests.get(url, headers=get_headers(), timeout=SCRAPE_TIMEOUT, allow_redirects=True, verify=False)
        if resp.status_code == 200:
            return resp.text, extract_emails(resp.text)
    except: pass
    return '', []

def enrich_employer(employer, skip_website_search=False):
    """Full enrichment: find website + extract emails."""
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
        'source': 'existing' if current_website else 'not_found'
    }
    
    # Step 1: Find website if missing
    website = current_website
    if not website and not skip_website_search:
        website, source = find_website(name, stadt)
        result['website'] = website
        result['source'] = source
        time.sleep(DDG_DELAY)
    
    # Step 2: Scrape for emails
    if website:
        all_emails = []
        
        # Main page
        html, emails = scrape_url(website)
        all_emails.extend(emails)
        result['scraped_from'].append(website)
        time.sleep(SCRAPE_DELAY)
        
        # Impressum/Kontakt if no emails found
        if not all_emails:
            contact_urls = get_contact_urls(website, html)
            for curl in contact_urls[:4]:
                if curl == website: continue
                _, cemails = scrape_url(curl)
                all_emails.extend(cemails)
                result['scraped_from'].append(curl)
                if cemails: break
                time.sleep(SCRAPE_DELAY)
        
        # Deduplicate and categorize
        unique = list(dict.fromkeys([e.lower() for e in all_emails if is_valid_email(e)]))
        result['emails'] = unique
        
        bewerbung = next((e for e in unique if categorize_email(e) == 'bewerbung'), '')
        kontakt = next((e for e in unique if categorize_email(e) == 'kontakt'), '')
        if not kontakt and unique:
            kontakt = unique[0]
        
        result['bewerbung_email'] = bewerbung
        result['kontakt_email'] = kontakt
    
    return result

def main():
    start = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    count = int(sys.argv[2]) if len(sys.argv) > 2 else 50
    
    print(f"=== Employer Enrichment V4 (DuckDuckGo + HTTP Scraping) ===")
    print(f"Processing {start} to {start+count-1}")
    
    # Load employers
    with open(INPUT_FILE, 'r') as f:
        employers = json.load(f)
    
    # Deduplicate
    seen = {}
    unique = []
    for e in employers:
        key = e['name'].lower().strip()
        if key not in seen:
            seen[key] = e
            unique.append(e)
    
    print(f"{len(employers)} entries → {len(unique)} unique")
    
    # Load existing results
    results = {}
    if os.path.exists(SEARCH_RESULTS):
        with open(SEARCH_RESULTS, 'r') as f:
            results = json.load(f)
    
    # Count how many already have emails
    with_emails = sum(1 for v in results.values() if v.get('emails') and len(v['emails']) > 0)
    with_websites = sum(1 for v in results.values() if v.get('website'))
    print(f"Existing: {len(results)} results ({with_websites} websites, {with_emails} with emails)")
    
    # Get pending
    pending = [e for e in unique if e['name'].lower().strip() not in results]
    slice_items = pending[start:start+count]
    print(f"Pending total: {len(pending)}, processing: {len(slice_items)}")
    
    total_w = 0
    total_e = 0
    start_time = time.time()
    
    for i, employer in enumerate(slice_items):
        result = enrich_employer(employer)
        
        key = employer['name'].lower().strip()
        results[key] = result
        
        if result.get('website'): total_w += 1
        if result.get('emails'): total_e += 1
        
        # Progress
        elapsed = time.time() - start_time
        status = "W" if result.get('website') else "-"
        status += "E" if result.get('emails') else "-"
        email_preview = result['emails'][0] if result.get('emails') else ''
        website = result.get('website', '')[:40]
        print(f"  [{i+1}/{len(slice_items)}] {status} {employer['name'][:45]:<45} | {website:<40} | {email_preview}")
        
        # Save every 10
        if (i + 1) % 10 == 0:
            with open(SEARCH_RESULTS, 'w') as f:
                json.dump(results, f, ensure_ascii=False)
            print(f"  → Saved {len(results)} | {total_w}W {total_e}E | {elapsed:.0f}s")
    
    # Final save
    with open(SEARCH_RESULTS, 'w') as f:
        json.dump(results, f, ensure_ascii=False)
    
    w = sum(1 for v in results.values() if v.get('website'))
    e = sum(1 for v in results.values() if v.get('emails') and len(v['emails']) > 0)
    print(f"\n=== DONE ===")
    print(f"Total results: {len(results)} | Websites: {w} ({w/len(results)*100:.1f}%) | E-Mails: {e} ({e/len(results)*100:.1f}%)")

if __name__ == '__main__':
    main()
