#!/usr/bin/env python3
"""
Find websites for companies in Bamberg/Erlangen/Nürnberg and save their HTML.
- Uses DuckDuckGo HTML search to find official websites
- Downloads and saves HTML for each found website
- Progress tracking with JSON
"""
import json, re, time, os, sys, hashlib, traceback
from datetime import datetime
import urllib.request
import urllib.parse
import urllib.error

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
ENRICH_FILE = '/home/z/my-project/website_enrichment.json'
HTML_DIR = '/home/z/my-project/download/website_html'
LOG_FILE = '/home/z/my-project/enrich_website_log.txt'
BATCH_SIZE = 50
DELAY = 1.5  # seconds between requests

TARGET_CITIES = ['Bamberg', 'Erlangen', 'Nürnberg']

def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, 'a') as f:
        f.write(line + '\n')

def safe_filename(name):
    """Create safe filename from company name."""
    name = re.sub(r'[^\w\s\-.]', '', name)
    name = re.sub(r'\s+', '_', name.strip())
    return name[:80]

def search_duckduckgo(query):
    """Search DuckDuckGo HTML for a query, return first result URL."""
    url = f'https://html.duckduckgo.com/html/?q={urllib.parse.quote(query)}'
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5'
    })
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode('utf-8', errors='replace')
        
        # Extract result links from DuckDuckGo HTML
        # Pattern: <a rel="nofollow" class="result__a" href="//duckduckgo.com/l/?uddg=ENCODED_URL&amp;rut=...">
        links = re.findall(r'uddg=([^&"]+)', html)
        for link in links[:5]:
            decoded = urllib.parse.unquote(link)
            # Skip social media, directories, etc.
            skip_domains = ['facebook.com', 'instagram.com', 'twitter.com', 'linkedin.com', 
                          'youtube.com', 'wikipedia.org', 'maps.google', 'gelbeseiten.de',
                          'yelp.de', 'tripadvisor', 'kundeu.com', 'mojomox', 'unternehmensregister.de',
                          'firmenwissen.de', 'impressum', 'northdata', 'kompass', 'wlw.de',
                          'hotfrog', 'cybo.com', 'de.lusha', 'dascleverle', 'cylex.de',
                          'opening-hours', 'indofolio', 'bizdb', 'preventum']
            if not any(s in decoded.lower() for s in skip_domains):
                # Prefer .de domains
                if decoded.startswith('http'):
                    return decoded
        
        # Fallback: any link
        for link in links[:3]:
            decoded = urllib.parse.unquote(link)
            if decoded.startswith('http'):
                return decoded
        
        return None
    except Exception as e:
        log(f"  Search error: {str(e)[:60]}")
        return None

def download_html(url, filepath, timeout=15):
    """Download HTML from URL and save to file."""
    req = urllib.request.Request(url, headers={
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
    })
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            # Follow redirects
            final_url = resp.url or url
            content = resp.read()
            # Only save if it's HTML
            content_type = resp.headers.get('Content-Type', '')
            if 'text/html' in content_type or 'application/xhtml' in content_type:
                with open(filepath, 'wb') as f:
                    f.write(content)
                return final_url, len(content)
            else:
                return final_url, 0
    except Exception as e:
        return None, 0

def extract_email_from_html(html_bytes):
    """Extract email addresses from HTML content."""
    try:
        text = html_bytes.decode('utf-8', errors='replace')
    except:
        text = html_bytes.decode('latin-1', errors='replace')
    
    emails = set()
    # Standard email pattern
    for m in re.finditer(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text):
        email = m.group(0).lower()
        if not any(x in email for x in ['.png', '.jpg', '.gif', '.svg', '.css', '.js', 'example.com', 'email.com', 'domain.com']):
            emails.add(email)
    
    # mailto: links
    for m in re.finditer(r'mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', text, re.I):
        emails.add(m.group(1).lower())
    
    # (at) obfuscation
    for m in re.finditer(r'[a-zA-Z0-9._%+-]+\s*[\(\[]at[\)\]]\s*[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', text, re.I):
        email = m.group(0).lower()
        email = re.sub(r'\s*[\(\[]at[\)\]]\s*', '@', email)
        emails.add(email)
    
    # Skip common non-business emails
    skip = ['sage', 'wordpress', 'admin@localhost', 'test@', 'noreply@', 'example@']
    emails = {e for e in emails if not any(s in e for s in skip)}
    
    # Prefer bewerbung/human resources emails
    preferred = [e for e in emails if any(k in e for k in ['bewerb', 'hr', 'recruit', 'karrier', 'job', 'career'])]
    if preferred:
        return preferred[0]
    
    # Prefer info/kontakt
    preferred2 = [e for e in emails if any(k in e for k in ['info', 'kontakt', 'contact'])]
    if preferred2:
        return preferred2[0]
    
    # Return first valid email
    valid = [e for e in emails if not e.startswith('webmaster@') and not e.startswith('postmaster@')]
    if valid:
        return sorted(valid)[0]
    
    return ''

def try_impressum(base_url, html_dir, safe_name):
    """Try to find Impressum page and extract email."""
    impressum_paths = ['/impressum', '/Impressum', '/imprint', '/kontakt', '/Kontakt', '/contact', '/kontakt.html', '/impressum.html']
    
    # Normalize base URL
    base = base_url.rstrip('/')
    
    for path in impressum_paths:
        url = base + path
        filepath = os.path.join(html_dir, f"{safe_name}_impressum{os.path.splitext(path)[1] if '.' in path else '.html'}")
        try:
            req = urllib.request.Request(url, headers={
                'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36',
                'Accept': 'text/html',
                'Accept-Language': 'de-DE,de;q=0.9',
            })
            with urllib.request.urlopen(req, timeout=8) as resp:
                content = resp.read()
                if len(content) > 200:
                    with open(filepath, 'wb') as f:
                        f.write(content)
                    email = extract_email_from_html(content)
                    if email:
                        return email, url
        except:
            pass
    
    return '', ''

def main():
    start_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    
    log("=" * 60)
    log("Website Enrichment + HTML Download - Bamberg/Erlangen/Nürnberg")
    log("=" * 60)
    
    # Load results
    with open(RESULTS_FILE) as f:
        all_results = json.load(f)
    
    # Filter target cities
    filtered = [r for r in all_results if r.get('city') in TARGET_CITIES]
    log(f"Total firms in Bamberg/Erlangen/Nürnberg: {len(filtered)}")
    
    # Load enrichment progress
    enrichment = {}
    if os.path.exists(ENRICH_FILE):
        with open(ENRICH_FILE) as f:
            enrichment = json.load(f)
    log(f"Already enriched: {len(enrichment)}")
    
    # Find firms that need enrichment
    need_enrichment = []
    for r in filtered:
        name = r.get('name', '').strip()
        if not name or name in enrichment:
            continue
        need_enrichment.append(r)
    
    log(f"Need enrichment: {len(need_enrichment)}")
    
    if not need_enrichment:
        log("All done!")
        return
    
    # Process in batches
    batch = need_enrichment[start_idx:start_idx + BATCH_SIZE]
    log(f"Processing batch: {start_idx} to {start_idx + len(batch) - 1}")
    
    found_website = 0
    found_email = 0
    saved_html = 0
    
    for i, firm in enumerate(batch):
        name = firm.get('name', '').strip()
        city = firm.get('city', '')
        phone = firm.get('phone', '')
        
        log(f"\n[{i+1}/{len(batch)}] {name} ({city})")
        
        # Search query
        query = f'{name} {city}'
        log(f"  Searching: {query}")
        
        # Step 1: Search for website
        website_url = search_duckduckgo(query)
        
        if not website_url:
            # Try simpler query
            query2 = f'"{name}" {city} Webseite'
            log(f"  Retry: {query2}")
            website_url = search_duckduckgo(query2)
        
        if not website_url:
            log(f"  No website found")
            enrichment[name] = {
                'website': '', 'html_file': '', 'email': '',
                'impressum_url': '', 'impressum_file': '',
                'searched_at': datetime.now().isoformat()[:19],
                'status': 'no_website'
            }
            with open(ENRICH_FILE, 'w') as f:
                json.dump(enrichment, f, ensure_ascii=False, indent=1)
            time.sleep(DELAY)
            continue
        
        log(f"  Found: {website_url}")
        
        # Step 2: Download HTML
        safe_name = safe_filename(name)
        html_filepath = os.path.join(HTML_DIR, f"{safe_name}.html")
        rel_path = f"website_html/{safe_name}.html"
        
        final_url, html_size = download_html(website_url, html_filepath)
        
        if html_size > 0:
            saved_html += 1
            log(f"  HTML saved: {html_size} bytes")
            
            # Step 3: Extract email from main page
            with open(html_filepath, 'rb') as f:
                html_content = f.read()
            email = extract_email_from_html(html_content)
            
            # Step 4: Try Impressum if no email found
            impressum_url = ''
            impressum_file = ''
            if not email:
                base = final_url or website_url
                email, impressum_url = try_impressum(base, HTML_DIR, safe_name)
                if impressum_url:
                    impressum_file = f"website_html/{safe_name}_impressum.html"
            
            if email:
                found_email += 1
                log(f"  Email: {email}")
            
            found_website += 1
            enrichment[name] = {
                'website': final_url or website_url,
                'html_file': rel_path,
                'email': email,
                'impressum_url': impressum_url,
                'impressum_file': impressum_file,
                'html_size': html_size,
                'searched_at': datetime.now().isoformat()[:19],
                'status': 'success'
            }
        else:
            log(f"  HTML download failed or non-HTML content")
            found_website += 1
            enrichment[name] = {
                'website': website_url,
                'html_file': '',
                'email': '',
                'impressum_url': '',
                'impressum_file': '',
                'searched_at': datetime.now().isoformat()[:19],
                'status': 'html_failed'
            }
        
        # Save progress
        with open(ENRICH_FILE, 'w') as f:
            json.dump(enrichment, f, ensure_ascii=False, indent=1)
        
        time.sleep(DELAY)
    
    log(f"\n{'='*60}")
    log(f"Batch complete!")
    log(f"  Websites found: {found_website}/{len(batch)}")
    log(f"  HTML saved: {saved_html}")
    log(f"  Emails found: {found_email}")
    log(f"  Total enriched: {len(enrichment)}")
    log(f"  Remaining: {len(need_enrichment) - start_idx - len(batch)}")

if __name__ == '__main__':
    main()
