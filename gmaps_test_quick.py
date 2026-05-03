#!/usr/bin/env python3
"""Quick test of Google Maps scraping with 1 town, 2 search terms."""
import time, json, re, os, sys
from datetime import datetime
from playwright.sync_api import sync_playwright

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'

def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    print(f"[{ts}] {msg}", flush=True)

def setup_browser():
    p = sync_playwright().start()
    browser = p.chromium.launch(
        headless=False,
        args=['--no-sandbox', '--disable-blink-features=AutomationControlled', '--disable-dev-shm-usage']
    )
    context = browser.new_context(
        viewport={'width': 1920, 'height': 1080},
        user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        locale='de-DE'
    )
    context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        window.chrome = {runtime: {}};
    """)
    return p, browser, context

def handle_cookies(page):
    for _ in range(5):
        try:
            btn = page.query_selector('button[aria-label*="Alle ablehnen"]')
            if not btn:
                btn = page.query_selector('button[aria-label*="Reject all"]')
            if btn:
                btn.click()
                time.sleep(1)
                return True
        except:
            pass
        time.sleep(0.5)
    return False

def search_google_maps(page, query):
    url = f'https://www.google.com/maps/search/{query.replace(" ", "+")}'
    log(f"  Navigating to: {url[:80]}...")
    
    try:
        page.goto(url, timeout=25000, wait_until='domcontentloaded')
    except Exception as e:
        log(f"  Page load error: {str(e)[:80]}")
        return []
    
    time.sleep(3)
    handle_cookies(page)
    time.sleep(1)
    
    # Scroll to load more results
    for i in range(5):
        try:
            page.evaluate('document.querySelector("div[role=feed]")?.scrollBy(0, 800)')
            time.sleep(1.2)
        except:
            break
    
    time.sleep(1)
    
    # Extract results
    results = page.evaluate("""() => {
        const businesses = [];
        const seen = new Set();
        const links = document.querySelectorAll('a[href*="/maps/place/"]');
        
        links.forEach(link => {
            try {
                const href = link.getAttribute('href') || '';
                const match = href.match(/\\/maps\\/place\\/([^/]+)/);
                if (!match) return;
                const slug = match[1];
                if (seen.has(slug)) return;
                seen.add(slug);
                
                const ariaLabel = link.getAttribute('aria-label') || '';
                let card = link;
                for (let i = 0; i < 10; i++) {
                    if (card.parentElement) card = card.parentElement;
                    const t = card.innerText || '';
                    if (t.length > 40) break;
                }
                const text = card.innerText || '';
                
                businesses.push({
                    name: ariaLabel.substring(0, 150),
                    href: href,
                    text: text.substring(0, 800),
                });
            } catch(e) {}
        });
        return businesses;
    }""")
    
    return results if results else []

def parse_business(raw, town_name):
    text = raw.get('text', '')
    name = raw.get('name', '')
    href = raw.get('href', '')
    
    rating = ''
    rating_match = re.search(r'(\d[,.]\d)', text[:50])
    if rating_match:
        rating = rating_match.group(1).replace(',', '.')
    
    phone = ''
    phone_match = re.search(r'(\+49[\s/\-\d]{7,})', text)
    if phone_match:
        phone = phone_match.group(1).strip()
    
    address = ''
    addr_match = re.search(r'([A-ZÄÖÜ][a-zäöüß]+(?:straße|str\.|str|weg|platz|allee|ring|gasse|damm|ufer|berg|feld|garten|hof|graben|markt)\s*\.?\s*\d+[a-zA-Z]?)', text)
    if addr_match:
        address = addr_match.group(1).strip()
    
    plz = ''
    plz_match = re.search(r'\b(\d{5})\s+[A-ZÄÖÜ]', text)
    if plz_match:
        plz = plz_match.group(1)
    
    has_website = bool(re.search(r'(Website|Webseite|website|Besuchen)', text, re.I))
    
    lat, lon = '', ''
    coords_match = re.search(r'!3d([\d.]+)!4d([\d.]+)', href)
    if coords_match:
        lat = coords_match.group(1)
        lon = coords_match.group(2)
    
    return {
        'name': name.strip(),
        'rating': rating,
        'address': address,
        'plz': plz,
        'city': town_name,
        'phone': phone,
        'has_website': has_website,
        'lat': lat,
        'lon': lon,
        'href': href,
        'search_town': town_name,
        'source': 'google_maps',
    }

def main():
    log("=== Quick Test: Google Maps Erlangen ===")
    
    os.system('pkill Xvfb 2>/dev/null')
    os.system('Xvfb :99 -screen 0 1920x1080x24 &')
    time.sleep(2)
    os.environ['DISPLAY'] = ':99'
    
    p, browser, context = setup_browser()
    page = context.new_page()
    
    all_results = []
    seen = set()
    
    searches = [
        ("Firma in Erlangen", "Erlangen"),
        ("IT Firma in Erlangen", "Erlangen"),
        ("Handwerk in Erlangen", "Erlangen"),
    ]
    
    try:
        for query, town in searches:
            log(f"\n--- {query} ---")
            raw = search_google_maps(page, query)
            log(f"  Got {len(raw)} raw results")
            
            new = 0
            for r in raw:
                parsed = parse_business(r, town)
                key = parsed['name'].lower().strip()
                if key and len(key) > 2 and key not in seen:
                    all_results.append(parsed)
                    seen.add(key)
                    new += 1
            
            log(f"  {new} new unique | Total: {len(all_results)}")
            time.sleep(3)
    
    finally:
        browser.close()
        p.stop()
        os.system('pkill Xvfb 2>/dev/null')
    
    # Save
    with open(RESULTS_FILE, 'w') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=1)
    
    log(f"\n=== Results: {len(all_results)} unique businesses ===")
    for r in all_results[:10]:
        log(f"  {r['name']} | Tel: {r['phone']} | Web: {r['has_website']} | PLZ: {r['plz']}")

if __name__ == '__main__':
    main()
