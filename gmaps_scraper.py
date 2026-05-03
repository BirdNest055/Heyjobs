#!/usr/bin/env python3
"""
Google Maps Business Scraper V2
- Uses Playwright with Xvfb
- Searches for businesses by category + city
- Scrolls to load more results
- Extracts: Name, Rating, Category, Address, Phone, Website link
- Clicks into detail pages for more info (email, full address, etc.)
"""

import time, json, re, os, sys
from playwright.sync_api import sync_playwright

RESULTS_FILE = '/home/z/my-project/gmaps_results.json'
BATCH_SIZE = 20

# German cities to search in (top IT cities)
CITIES = [
    'München', 'Berlin', 'Hamburg', 'Frankfurt', 'Stuttgart',
    'Düsseldorf', 'Köln', 'Hannover', 'Nürnberg', 'Dresden',
    'Leipzig', 'Darmstadt', 'Bonn', 'Essen', 'Bremen',
    'Dortmund', 'Mannheim', 'Karlsruhe', 'Wiesbaden', 'Augsburg',
]

SEARCH_TERMS = [
    'IT Firma', 'Softwareentwicklung', 'IT Dienstleister', 
    'Systemhaus', 'IT Beratung', 'EDV',
]

def setup_browser():
    p = sync_playwright().start()
    browser = p.chromium.launch(
        headless=False,
        args=['--no-sandbox', '--disable-blink-features=AutomationControlled']
    )
    context = browser.new_context(
        viewport={'width': 1920, 'height': 1080},
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        locale='de-DE'
    )
    context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        window.chrome = {runtime: {}};
    """)
    return p, browser, context

def handle_cookies(page):
    """Reject Google cookies."""
    for _ in range(3):
        btn = page.query_selector('button[aria-label*="Alle ablehnen"]') or page.query_selector('button[aria-label*="Reject all"]')
        if btn:
            btn.click()
            time.sleep(1)
            return True
        time.sleep(1)
    return False

def search_maps(page, query):
    """Search Google Maps and return results."""
    url = f'https://www.google.com/maps/search/{query.replace(" ", "+")}'
    page.goto(url, timeout=30000)
    time.sleep(3)
    handle_cookies(page)
    time.sleep(2)
    
    # Scroll to load more results
    feed = page.query_selector('div[role="feed"]')
    if feed:
        for _ in range(5):
            page.evaluate('document.querySelector("div[role=feed]")?.scrollBy(0, 800)')
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
                
                // Get parent card text
                let card = link;
                for (let i = 0; i < 8; i++) {
                    if (card.parentElement) card = card.parentElement;
                    const t = card.innerText || '';
                    if (t.length > 30) break;
                }
                const text = card.innerText || '';
                
                businesses.push({
                    name: ariaLabel.substring(0, 120),
                    href: href,
                    text: text.substring(0, 500),
                });
            } catch(e) {}
        });
        
        return businesses;
    }""")
    
    return results

def parse_business(raw):
    """Parse raw Google Maps text into structured data."""
    text = raw.get('text', '')
    name = raw.get('name', '')
    href = raw.get('href', '')
    
    # Extract rating (e.g. "4,9" or "4.9")
    rating_match = re.search(r'(\d[,.]\d)', text)
    rating = rating_match.group(1) if rating_match else ''
    
    # Extract phone (German format)
    phone_match = re.search(r'(\+49[\s/\-\d]{8,})', text)
    phone = phone_match.group(1).strip() if phone_match else ''
    
    # Extract address (German street patterns)
    addr_match = re.search(r'([A-ZÄÖÜ][a-zäöüß]+(?:straße|str\.|weg|platz|allee|ring|gasse|damm|ufer)\s*\d+[a-zA-Z]?)', text)
    address = addr_match.group(1) if addr_match else ''
    
    # Check for website indicator
    has_website = 'Website' in text or 'Webseite' in text or 'website' in text.lower()
    
    # Extract category
    cat_match = re.search(r'(IT-Berater|Softwareentwicklung|Computerservice|IT-Dienstleister|EDV|Systemhaus|Telekommunikation|Webdesign|Software|IT-Service|IT-Sicherheit|Cloud|Computersupport und -dienste|Informatik)', text, re.I)
    category = cat_match.group(1) if cat_match else ''
    
    # Extract PLZ
    plz_match = re.search(r'\b(\d{5})\b', text)
    plz = plz_match.group(1) if plz_match else ''
    
    return {
        'name': name,
        'rating': rating,
        'category': category,
        'address': address,
        'plz': plz,
        'phone': phone,
        'has_website': has_website,
        'href': href,
        'source': 'google_maps',
    }

def main():
    start_city = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    num_cities = int(sys.argv[2]) if len(sys.argv) > 2 else 3
    
    print(f"=== Google Maps Scraper V2 ===")
    print(f"Cities: {start_city} to {start_city + num_cities - 1}")
    
    # Load existing results
    all_results = []
    if os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, 'r') as f:
            all_results = json.load(f)
    
    seen_names = set(r['name'].lower().strip() for r in all_results if r.get('name'))
    print(f"Existing: {len(all_results)}, unique: {len(seen_names)}")
    
    # Setup Xvfb
    os.system('export DISPLAY=:99; Xvfb :99 -screen 0 1920x1080x24 &')
    time.sleep(2)
    
    p, browser, context = setup_browser()
    page = context.new_page()
    
    cities_to_process = CITIES[start_city:start_city + num_cities]
    total_new = 0
    
    try:
        for city in cities_to_process:
            for term in SEARCH_TERMS:
                query = f'{term} {city}'
                print(f"\n--- Searching: {query} ---")
                
                try:
                    raw_results = search_maps(page, query)
                    new_count = 0
                    
                    for raw in raw_results:
                        parsed = parse_business(raw)
                        name_key = parsed['name'].lower().strip()
                        
                        if name_key and name_key not in seen_names and len(name_key) > 2:
                            all_results.append(parsed)
                            seen_names.add(name_key)
                            new_count += 1
                    
                    total_new += new_count
                    print(f"  Found {len(raw_results)} results, {new_count} new | Total: {len(all_results)}")
                    
                    # Save after each search
                    with open(RESULTS_FILE, 'w') as f:
                        json.dump(all_results, f, ensure_ascii=False)
                    
                    time.sleep(3)  # Be polite
                
                except Exception as e:
                    print(f"  Error: {str(e)[:100]}")
                    time.sleep(5)
                    continue
    
    finally:
        browser.close()
        p.stop()
        os.system('pkill Xvfb 2>/dev/null')
    
    # Final save
    with open(RESULTS_FILE, 'w') as f:
        json.dump(all_results, f, ensure_ascii=False)
    
    print(f"\n=== DONE ===")
    print(f"Total: {len(all_results)} businesses from {len(cities_to_process)} cities")
    print(f"New in this run: {total_new}")
    with_website = sum(1 for r in all_results if r.get('has_website'))
    with_phone = sum(1 for r in all_results if r.get('phone'))
    print(f"With website: {with_website}, With phone: {with_phone}")

if __name__ == '__main__':
    main()
