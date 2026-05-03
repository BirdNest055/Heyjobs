#!/usr/bin/env python3
"""
Google Maps Erlangen 50km Scraper - Playwright + Xvfb
- Searches for ALL businesses in towns within 50km of Erlangen
- Uses multiple search terms per location for maximum coverage
- Extensive scrolling to load all available results
- Deduplicates across all searches
- Saves progress incrementally
"""

import time, json, re, os, sys, math, traceback
from datetime import datetime
from playwright.sync_api import sync_playwright

# ============ CONFIGURATION ============
RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
PROGRESS_FILE = '/home/z/my-project/gmaps_erlangen_progress.json'
LOG_FILE = '/home/z/my-project/gmaps_erlangen_log.txt'

ERLANGEN_LAT = 49.5969
ERLANGEN_LON = 11.0043
RADIUS_KM = 50

SCROLL_ITERATIONS = 8  # Number of scroll cycles per search
DELAY_BETWEEN_SEARCHES = 3  # Seconds between searches
DELAY_BETWEEN_SCROLLS = 1.5  # Seconds between scrolls

# ============ LOGGING ============
def log(msg):
    ts = datetime.now().strftime('%H:%M:%S')
    line = f"[{ts}] {msg}"
    print(line, flush=True)
    with open(LOG_FILE, 'a') as f:
        f.write(line + '\n')

# ============ SEARCH TERMS ============
# Full list for large/medium cities
FULL_SEARCH_TERMS = [
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

# Medium list for small towns
MEDIUM_SEARCH_TERMS = [
    "Firma", "Unternehmen", "GmbH",
    "Handwerk", "Dienstleistung",
    "Gastronomie", "Baufirma",
    "IT", "Steuerberater",
    "Kfz", "Elektro",
    "Friseur", "Bäckerei",
]

# Minimal list for tiny villages
MINIMAL_SEARCH_TERMS = [
    "Firma", "Unternehmen", "Gewerbe",
    "Handwerk", "Dienstleistung",
]

# ============ TOWN LIST ============
def get_towns():
    """Get pre-computed town list."""
    with open('/home/z/my-project/gmaps_erlangen_towns.json', 'r') as f:
        return json.load(f)

# ============ PLAYWRIGHT SETUP ============
def setup_browser():
    p = sync_playwright().start()
    browser = p.chromium.launch(
        headless=False,
        args=[
            '--no-sandbox',
            '--disable-blink-features=AutomationControlled',
            '--disable-dev-shm-usage',
            '--window-size=1920,1080'
        ]
    )
    context = browser.new_context(
        viewport={'width': 1920, 'height': 1080},
        user_agent='Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        locale='de-DE',
        geolocation={'latitude': ERLANGEN_LAT, 'longitude': ERLANGEN_LON},
        permissions=['geolocation']
    )
    context.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        window.chrome = {runtime: {}};
    """)
    return p, browser, context

def handle_cookies(page):
    """Reject Google cookies dialog."""
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

# ============ GOOGLE MAPS SEARCH ============
def search_google_maps(page, query, max_scrolls=SCROLL_ITERATIONS):
    """Search Google Maps and extract business listings."""
    url = f'https://www.google.com/maps/search/{query.replace(" ", "+")}'
    
    try:
        page.goto(url, timeout=30000, wait_until='domcontentloaded')
    except Exception as e:
        log(f"  Page load error: {str(e)[:80]}")
        return []
    
    time.sleep(3)
    handle_cookies(page)
    time.sleep(1)
    
    # Check for CAPTCHA or blocking
    if "sorry" in page.url or "captcha" in page.url.lower():
        log("  ⚠️ CAPTCHA/Block detected! Waiting 30s...")
        time.sleep(30)
        return []
    
    # Scroll to load more results
    feed = page.query_selector('div[role="feed"]')
    if feed:
        for i in range(max_scrolls):
            try:
                page.evaluate('document.querySelector("div[role=feed]")?.scrollBy(0, 1000)')
                time.sleep(DELAY_BETWEEN_SCROLLS)
            except:
                break
    
    # Wait a moment for all results to render
    time.sleep(1)
    
    # Extract business listings from the page
    results = page.evaluate("""() => {
        const businesses = [];
        const seen = new Set();
        
        // Find all business listing links
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
                
                // Get the parent container with full text
                let card = link;
                for (let i = 0; i < 10; i++) {
                    if (card.parentElement) card = card.parentElement;
                    const t = card.innerText || '';
                    if (t.length > 40) break;
                }
                const text = card.innerText || '';
                
                // Try to extract structured data from the card
                const lines = text.split('\\n').filter(l => l.trim());
                
                businesses.push({
                    name: ariaLabel.substring(0, 150),
                    href: href,
                    text: text.substring(0, 800),
                    lines: lines.slice(0, 12),
                });
            } catch(e) {}
        });
        
        return businesses;
    }""")
    
    return results if results else []

# ============ PARSE BUSINESS DATA ============
def parse_business(raw, town_name):
    """Parse raw Google Maps listing into structured data."""
    text = raw.get('text', '')
    lines = raw.get('lines', [])
    name = raw.get('name', '')
    href = raw.get('href', '')
    
    # Extract rating (e.g. "4,9" or "4.9")
    rating = ''
    rating_match = re.search(r'(\d[,.]\d)', text[:50])  # Rating usually near beginning
    if rating_match:
        rating = rating_match.group(1).replace(',', '.')
    
    # Extract category (usually second line after name)
    category = ''
    category_patterns = [
        r'(IT-Berater|Softwareentwicklung|Computerservice|IT-Dienstleister|EDV|Systemhaus|'
        r'Telekommunikation|Webdesign|Software|IT-Service|IT-Sicherheit|Cloud|'
        r'Computersupport und -dienste|Informatik|Internetdienstleister|'
        r'Restaurant|Café|Bäckerei|Metzgerei|Friseur|Apotheke|Arzt|'
        r'Steuerberater|Rechtsanwalt|Notar|Versicherung|Immobilien|'
        r'Bauunternehmen|Elektroinstallateur|Sanitär|Heizung|'
        r'Maschinenbau|Metallbau|Kunststoff|Druckerei|'
        r'Kfz-Werkstatt|Autohaus|Tankstelle|'
        r'Logistik|Spedition|Handel|Großhandel|'
        r'Marketing|Werbeagentur|Design|'
        r'Handwerker|Schreiner|Maler|Dachdecker|'
        r'Beratung|Dienstleistung|Gastronomie|Hotel|Pension|'
        r'Schule|Kindergarten|Bildung|'
        r'Fitness|Sport|Gesundheit|Physiotherapie|Zahnarzt|'
        r'Bank|Finanzdienstleistung|Vermögensberatung|'
        r'Architekt|Ingenieurbüro|Vermessung|'
        r'Gartenbau|Florist|Landwirtschaft|'
        r'Optiker|Hörgeräte|Sanitätshaus|'
        r'Reisebüro|Ticketservice|'
        r'Fahrschule|Detektei|Sicherheit|'
        r'Gemeinde|Stadtverwaltung|Behörde)'
    ]
    for line in lines[:5]:
        cat_match = re.search(category_patterns[0], line, re.I)
        if cat_match:
            category = cat_match.group(1)
            break
    
    # If no specific category found, try to extract from text
    if not category:
        # Look for "·" separated category
        for line in lines[1:4]:
            if '·' in line and len(line) < 80:
                parts = line.split('·')
                for part in parts:
                    part = part.strip()
                    if part and not re.match(r'^[\d,.]+$', part) and len(part) > 3:
                        category = part
                        break
                if category:
                    break
    
    # Extract phone (German format)
    phone = ''
    phone_match = re.search(r'(\+49[\s/\-\d]{7,}|\d{3,4}[\s/\-]\d{5,8})', text)
    if phone_match:
        phone = phone_match.group(1).strip()
    
    # Extract address (German patterns)
    address = ''
    addr_match = re.search(r'([A-ZÄÖÜ][a-zäöüß]+(?:straße|str\.|straße|str|weg|platz|allee|ring|gasse|damm|ufer|berg|feld|garten|hof|graben|steg|markt|kirch|schloss|park|see|tal|bach)\s*\.?\s*\d+[a-zA-Z]?(?:\s*[-/]\s*\d+[a-zA-Z]?)?)', text)
    if addr_match:
        address = addr_match.group(1).strip()
    
    # Extract PLZ (5-digit German postal code)
    plz = ''
    plz_match = re.search(r'\b(\d{5})\s+[A-ZÄÖÜ]', text)
    if plz_match:
        plz = plz_match.group(1)
    if not plz:
        plz_match = re.search(r'\b(9[0-5]\d{3})\b', text)
        if plz_match:
            plz = plz_match.group(1)
    
    # Extract city from PLZ area or text
    city = town_name
    city_match = re.search(r'\b\d{5}\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s+[A-ZÄÖÜ][a-zäöüß]+)?)', text)
    if city_match:
        found_city = city_match.group(1).strip()
        if len(found_city) > 2:
            city = found_city
    
    # Check for website indicator
    has_website = bool(re.search(r'(Website|Webseite|website|Besuchen|Visit)', text, re.I))
    
    # Extract website URL if present
    website = ''
    website_match = re.search(r'(https?://[^\s<"\']+\.(de|com|eu|net|org|info)[^\s<"\']*)', text)
    if website_match:
        website = website_match.group(1).rstrip('.,;:')
    
    # Extract coordinates from href
    lat, lon = '', ''
    coords_match = re.search(r'!3d([\d.]+)!4d([\d.]+)', href)
    if coords_match:
        lat = coords_match.group(1)
        lon = coords_match.group(2)
    
    return {
        'name': name.strip(),
        'rating': rating,
        'category': category,
        'address': address,
        'plz': plz,
        'city': city,
        'phone': phone,
        'has_website': has_website,
        'website': website,
        'lat': lat,
        'lon': lon,
        'href': href,
        'search_town': town_name,
        'source': 'google_maps',
        'scraped_at': datetime.now().isoformat()[:19],
    }

# ============ PROGRESS MANAGEMENT ============
def load_progress():
    """Load progress from file."""
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, 'r') as f:
            return json.load(f)
    return {'completed_searches': [], 'last_town_idx': 0, 'last_term_idx': 0}

def save_progress(progress):
    """Save progress to file."""
    with open(PROGRESS_FILE, 'w') as f:
        json.dump(progress, f, ensure_ascii=False)

def load_results():
    """Load existing results."""
    if os.path.exists(RESULTS_FILE):
        with open(RESULTS_FILE, 'r') as f:
            return json.load(f)
    return []

def save_results(results):
    """Save results to file."""
    with open(RESULTS_FILE, 'w') as f:
        json.dump(results, f, ensure_ascii=False, indent=1)

# ============ MAIN ============
def main():
    start_town_idx = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    num_towns = int(sys.argv[2]) if len(sys.argv) > 2 else 999
    
    log("=" * 60)
    log("Google Maps Erlangen 50km Scraper")
    log("=" * 60)
    
    # Load towns and results
    towns = get_towns()
    all_results = load_results()
    progress = load_progress()
    
    # Build dedup set
    seen_keys = set()
    for r in all_results:
        key = r.get('name', '').lower().strip()
        if key:
            seen_keys.add(key)
    
    log(f"Towns: {len(towns)}, Existing results: {len(all_results)}, Unique: {len(seen_keys)}")
    
    # Setup Xvfb
    os.system('pkill Xvfb 2>/dev/null')
    os.system('Xvfb :99 -screen 0 1920x1080x24 &')
    time.sleep(2)
    os.environ['DISPLAY'] = ':99'
    
    p, browser, context = setup_browser()
    page = context.new_page()
    
    # Process towns
    end_town_idx = min(start_town_idx + num_towns, len(towns))
    total_new = 0
    search_count = 0
    
    try:
        for town_idx in range(start_town_idx, end_town_idx):
            town = towns[town_idx]
            town_name = town['name']
            dist = town['distance_km']
            tier = town['population_tier']
            
            # Select search terms based on town size
            if tier <= 2:  # Large/medium city
                terms = FULL_SEARCH_TERMS
            elif tier == 3:  # Small town
                terms = MEDIUM_SEARCH_TERMS
            else:  # Tiny village
                terms = MINIMAL_SEARCH_TERMS
            
            log(f"\n--- Town {town_idx+1}/{len(towns)}: {town_name} ({dist}km, tier={tier}, {len(terms)} terms) ---")
            
            for term_idx, term in enumerate(terms):
                query = f'{term} in {town_name}'
                search_key = f"{town_name}|{term}"
                
                # Skip already completed searches
                if search_key in progress['completed_searches']:
                    continue
                
                log(f"  Searching: '{query}'")
                search_count += 1
                
                try:
                    raw_results = search_google_maps(page, query)
                    new_count = 0
                    
                    for raw in raw_results:
                        parsed = parse_business(raw, town_name)
                        name_key = parsed['name'].lower().strip()
                        
                        # Dedup by name
                        if name_key and len(name_key) > 2 and name_key not in seen_keys:
                            all_results.append(parsed)
                            seen_keys.add(name_key)
                            new_count += 1
                    
                    total_new += new_count
                    log(f"  Found {len(raw_results)} listings, {new_count} new | Total unique: {len(all_results)}")
                    
                    # Mark search as completed
                    progress['completed_searches'].append(search_key)
                    
                    # Save after each search
                    save_results(all_results)
                    save_progress(progress)
                    
                    time.sleep(DELAY_BETWEEN_SEARCHES)
                
                except Exception as e:
                    log(f"  ❌ Error: {str(e)[:120]}")
                    traceback.print_exc()
                    
                    # Check if browser is still alive
                    try:
                        page.evaluate('1+1')
                    except:
                        log("  Browser crashed! Restarting...")
                        try:
                            browser.close()
                        except:
                            pass
                        try:
                            p.stop()
                        except:
                            pass
                        time.sleep(5)
                        p, browser, context = setup_browser()
                        page = context.new_page()
                    
                    time.sleep(5)
                    continue
            
            # Summary after each town
            log(f"  Town '{town_name}' done. Total unique businesses: {len(all_results)}")
    
    except KeyboardInterrupt:
        log("\n⚠️ Interrupted by user! Saving progress...")
    
    except Exception as e:
        log(f"\n❌ Fatal error: {str(e)[:200]}")
        traceback.print_exc()
    
    finally:
        try:
            browser.close()
        except:
            pass
        try:
            p.stop()
        except:
            pass
        os.system('pkill Xvfb 2>/dev/null')
    
    # Final save
    save_results(all_results)
    save_progress(progress)
    
    # Print summary
    log("\n" + "=" * 60)
    log("SCRAPING COMPLETE")
    log("=" * 60)
    log(f"Searches performed: {search_count}")
    log(f"Total unique businesses: {len(all_results)}")
    log(f"New in this run: {total_new}")
    
    with_website = sum(1 for r in all_results if r.get('has_website'))
    with_phone = sum(1 for r in all_results if r.get('phone'))
    with_plz = sum(1 for r in all_results if r.get('plz'))
    with_address = sum(1 for r in all_results if r.get('address'))
    with_category = sum(1 for r in all_results if r.get('category'))
    
    log(f"With website: {with_website} ({100*with_website/len(all_results):.0f}%)")
    log(f"With phone: {with_phone} ({100*with_phone/len(all_results):.0f}%)")
    log(f"With PLZ: {with_plz} ({100*with_plz/len(all_results):.0f}%)")
    log(f"With address: {with_address} ({100*with_address/len(all_results):.0f}%)")
    log(f"With category: {with_category} ({100*with_category/len(all_results):.0f}%)")
    
    # Top categories
    cat_counts = {}
    for r in all_results:
        cat = r.get('category', 'Unbekannt')
        cat_counts[cat] = cat_counts.get(cat, 0) + 1
    log(f"\nTop categories:")
    for cat, count in sorted(cat_counts.items(), key=lambda x: -x[1])[:15]:
        log(f"  {cat}: {count}")

if __name__ == '__main__':
    main()
