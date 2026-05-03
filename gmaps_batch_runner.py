#!/usr/bin/env python3
"""
Google Maps Erlangen 50km - Optimized Batch Runner
- Runs in small batches to avoid memory issues
- Properly cleans up browser between towns
- Resume capability via progress file
"""

import time, json, re, os, sys, math, traceback, gc, subprocess
from datetime import datetime
from playwright.sync_api import sync_playwright

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
PROGRESS_FILE = '/home/z/my-project/gmaps_erlangen_progress.json'
LOG_FILE = '/home/z/my-project/gmaps_erlangen_log.txt'

ERLANGEN_LAT = 49.5969
ERLANGEN_LON = 11.0043

def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, 'a') as f:
        f.write(line + '\n')

# Search terms by town tier
FULL_TERMS = [
    "Firma", "Unternehmen", "GmbH",
    "IT Firma", "Software", "Softwareentwicklung",
    "Handwerk", "Industrie", "Handel", "Dienstleistung",
    "Gastronomie", "Baufirma", "Kfz Werkstatt",
    "Steuerberater", "Rechtsanwalt",
    "Versicherung", "Immobilien",
    "Marketing Agentur", "Logistik",
    "Elektro", "Metallbau", "Maschinenbau",
    "Kunststoff", "Druckerei",
    "Apotheke", "Arztpraxis",
    "Friseur", "Bäckerei",
]

MEDIUM_TERMS = [
    "Firma", "Unternehmen", "GmbH",
    "Handwerk", "Dienstleistung",
    "Gastronomie", "Baufirma",
    "IT", "Steuerberater",
    "Kfz", "Elektro",
    "Friseur", "Bäckerei",
]

MINIMAL_TERMS = [
    "Firma", "Unternehmen", "Gewerbe",
    "Handwerk", "Dienstleistung",
]

def get_towns():
    with open('/home/z/my-project/gmaps_erlangen_towns.json', 'r') as f:
        return json.load(f)

def load_results():
    if os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, 'r') as f:
            return json.load(f)
    return []

def save_results(results):
    with open(RESULTS_FILE, 'w') as f:
        json.dump(results, f, ensure_ascii=False)

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'r') as f:
            return json.load(f)
    return {'completed_searches': [], 'stats': {}}

def save_progress(progress):
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, ensure_ascii=False)

def search_and_extract(page, query, max_scrolls=6):
    """Search Google Maps and extract business listings."""
    url = f'https://www.google.com/maps/search/{query.replace(" ", "+")}'
    
    try:
        page.goto(url, timeout=25000, wait_until='domcontentloaded')
    except Exception as e:
        log(f"  Page load error: {str(e)[:80]}")
        return []
    
    time.sleep(2.5)
    
    # Handle cookies
    try:
        btn = page.query_selector('button[aria-label*="Alle ablehnen"]')
        if not btn:
            btn = page.query_selector('button[aria-label*="Reject all"]')
        if btn:
            btn.click()
            time.sleep(1)
    except:
        pass
    
    # Check for CAPTCHA
    page_url = page.url
    if "sorry" in page_url or "captcha" in page_url.lower():
        log("  ⚠️ CAPTCHA detected! Waiting 60s...")
        time.sleep(60)
        return "CAPTCHA"
    
    # Scroll to load more results
    for i in range(max_scrolls):
        try:
            page.evaluate('document.querySelector("div[role=feed]")?.scrollBy(0, 1000)')
            time.sleep(1.0)
        except:
            break
    
    time.sleep(0.5)
    
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
    m = re.search(r'(\d[,.]\d)', text[:50])
    if m:
        rating = m.group(1).replace(',', '.')
    
    phone = ''
    m = re.search(r'(\+49[\s/\-\d]{7,})', text)
    if m:
        phone = m.group(1).strip()
    
    address = ''
    m = re.search(r'([A-ZÄÖÜ][a-zäöüß]+(?:straße|str\.|str|weg|platz|allee|ring|gasse|damm|ufer|berg|feld|garten|hof|graben|markt)\s*\.?\s*\d+[a-zA-Z]?)', text)
    if m:
        address = m.group(1).strip()
    
    plz = ''
    m = re.search(r'\b(9[0-5]\d{3})\b', text)
    if m:
        plz = m.group(1)
    
    has_website = bool(re.search(r'(Website|Webseite|website|Besuchen)', text, re.I))
    
    lat, lon = '', ''
    m = re.search(r'!3d([\d.]+)!4d([\d.]+)', href)
    if m:
        lat = m.group(1)
        lon = m.group(2)
    
    # Extract category from text
    category = ''
    lines = text.split('\n')
    cat_patterns = [
        'IT-Berater', 'Softwareentwicklung', 'Computerservice', 'IT-Dienstleister',
        'EDV', 'Systemhaus', 'Webdesign', 'IT-Service', 'IT-Sicherheit',
        'Restaurant', 'Café', 'Bäckerei', 'Metzgerei', 'Friseur', 'Apotheke',
        'Steuerberater', 'Rechtsanwalt', 'Notar', 'Versicherung', 'Immobilien',
        'Bauunternehmen', 'Elektroinstallateur', 'Maschinenbau', 'Metallbau',
        'Kfz-Werkstatt', 'Autohaus', 'Logistik', 'Spedition', 'Handel',
        'Marketing', 'Werbeagentur', 'Schreiner', 'Maler', 'Dachdecker',
        'Beratung', 'Dienstleistung', 'Gastronomie', 'Hotel', 'Arzt',
        'Zahnarzt', 'Bank', 'Architekt', 'Ingenieurbüro', 'Gartenbau',
        'Optiker', 'Reisebüro', 'Fahrschule', 'Sanitär', 'Heizung',
        'Druckerei', 'Kunststoff', 'Handwerker',
    ]
    for line in lines[1:5]:
        for cp in cat_patterns:
            if cp.lower() in line.lower():
                category = cp
                break
        if category:
            break
    
    return {
        'name': name.strip(),
        'rating': rating,
        'category': category,
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
        'scraped_at': datetime.now().isoformat()[:19],
    }

def main():
    start_town = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    num_towns = int(sys.argv[2]) if len(sys.argv) > 2 else 5
    
    log("=" * 50)
    log(f"Erlangen 50km Scraper - Towns {start_town} to {start_town + num_towns - 1}")
    log("=" * 50)
    
    towns = get_towns()
    all_results = load_results()
    progress = load_progress()
    
    seen_keys = set(r.get('name', '').lower().strip() for r in all_results if r.get('name'))
    log(f"Existing: {len(all_results)} unique businesses")
    
    # Ensure Xvfb is running
    result = subprocess.run(['pgrep', '-f', 'Xvfb'], capture_output=True)
    if result.returncode != 0:
        os.system('Xvfb :99 -screen 0 1920x1080x24 &')
        time.sleep(2)
    os.environ['DISPLAY'] = ':99'
    
    p = sync_playwright().start()
    browser = p.chromium.launch(
        headless=False,
        args=['--no-sandbox', '--disable-blink-features=AutomationControlled',
              '--disable-dev-shm-usage', '--disable-gpu',
              '--disable-software-rasterizer',
              '--single-process']  # Reduce memory
    )
    context = browser.new_context(
        viewport={'width': 1280, 'height': 800},  # Smaller viewport = less memory
        user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        locale='de-DE'
    )
    context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        window.chrome = {runtime: {}};
    """)
    page = context.new_page()
    
    end_town = min(start_town + num_towns, len(towns))
    total_new = 0
    search_count = 0
    
    try:
        for ti in range(start_town, end_town):
            town = towns[ti]
            town_name = town['name']
            dist = town['distance_km']
            tier = town['population_tier']
            
            if tier <= 2:
                terms = FULL_TERMS
            elif tier == 3:
                terms = MEDIUM_TERMS
            else:
                terms = MINIMAL_TERMS
            
            log(f"\n--- Town {ti+1}/{len(towns)}: {town_name} ({dist}km, tier={tier}, {len(terms)} terms) ---")
            
            for term in terms:
                search_key = f"{town_name}|{term}"
                if search_key in progress['completed_searches']:
                    continue
                
                query = f'{term} in {town_name}'
                search_count += 1
                
                try:
                    raw = search_and_extract(page, query)
                    
                    if raw == "CAPTCHA":
                        log("  ⚠️ CAPTCHA! Saving and stopping.")
                        break
                    
                    new_count = 0
                    for r in raw:
                        parsed = parse_business(r, town_name)
                        key = parsed['name'].lower().strip()
                        if key and len(key) > 2 and key not in seen_keys:
                            all_results.append(parsed)
                            seen_keys.add(key)
                            new_count += 1
                    
                    total_new += new_count
                    log(f"  '{term}': {len(raw)} listings, {new_count} new | Total: {len(all_results)}")
                    
                    progress['completed_searches'].append(search_key)
                    save_results(all_results)
                    save_progress(progress)
                    
                    time.sleep(2.5)
                
                except Exception as e:
                    log(f"  ❌ Error: {str(e)[:100]}")
                    # Try to recover browser
                    try:
                        page.evaluate('1+1')
                    except:
                        log("  Restarting browser...")
                        try: browser.close()
                        except: pass
                        browser = p.chromium.launch(
                            headless=False,
                            args=['--no-sandbox', '--disable-blink-features=AutomationControlled',
                                  '--disable-dev-shm-usage', '--single-process']
                        )
                        context = browser.new_context(
                            viewport={'width': 1280, 'height': 800},
                            user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36',
                            locale='de-DE'
                        )
                        context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined});")
                        page = context.new_page()
                    time.sleep(3)
                    continue
            
            # Force GC between towns
            gc.collect()
            log(f"  Town done. Total: {len(all_results)}")
    
    except KeyboardInterrupt:
        log("Interrupted!")
    
    except Exception as e:
        log(f"Fatal: {str(e)[:200]}")
        traceback.print_exc()
    
    finally:
        try: browser.close()
        except: pass
        try: p.stop()
        except: pass
    
    save_results(all_results)
    save_progress(progress)
    
    log(f"\nDone! {search_count} searches, {total_new} new, Total: {len(all_results)}")
    with_phone = sum(1 for r in all_results if r.get('phone'))
    with_website = sum(1 for r in all_results if r.get('has_website'))
    log(f"Phone: {with_phone}, Website: {with_website}")

if __name__ == '__main__':
    main()
