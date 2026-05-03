#!/usr/bin/env python3
"""
Arbeitsagentur Job Scraper using Playwright with virtual display.
Scrapes employer data from: https://www.arbeitsagentur.de/jobsuche/suche?angebotsart=1&branche=11
Processes in batches and saves progress regularly.
"""

import os
import re
import json
import time
import random
import hashlib
from datetime import datetime

os.environ['DISPLAY'] = ':99'

from playwright.sync_api import sync_playwright

# ========== Configuration ==========
BASE_URL = "https://www.arbeitsagentur.de/jobsuche/suche?angebotsart=1&branche=11"
OUTPUT_FILE = "/home/z/my-project/download/arbeitsagentur_employers.json"
EXCEL_FILE = "/home/z/my-project/download/Arbeitgeber_Arbeitsagentur_Bewerbungskontakte.xlsx"
MAX_PAGES = 200  # Maximum pages to scrape
BATCH_SIZE = 5   # Pages per batch before saving progress
PAGE_TIMEOUT = 30000  # 30 seconds timeout

# ========== State ==========
all_employers = []
seen_ids = set()
current_page = 1


def load_progress():
    """Load previously saved employers from JSON."""
    global all_employers, seen_ids
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
                all_employers = json.load(f)
            for emp in all_employers:
                eid = emp.get('Arbeitgeber_ID', '')
                if eid:
                    seen_ids.add(eid)
            print(f"[PROGRESS] Loaded {len(all_employers)} existing employers from progress file.")
        except Exception as e:
            print(f"[WARN] Could not load progress: {e}")
            all_employers = []
            seen_ids = set()


def save_progress():
    """Save current employers to JSON."""
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_employers, f, ensure_ascii=False, indent=2)
    print(f"[PROGRESS] Saved {len(all_employers)} employers to {OUTPUT_FILE}")


def generate_employer_id(plz, name):
    """Generate unique employer ID: DE-PLZ-NNNN."""
    clean_plz = re.sub(r'[^0-9]', '', str(plz))[:5] if plz else '00000'
    if len(clean_plz) < 5:
        clean_plz = clean_plz.zfill(5)
    hash_input = f"{name}_{plz}".lower().strip()
    hash_num = int(hashlib.md5(hash_input.encode()).hexdigest()[:4], 16) % 10000
    return f"DE-{clean_plz}-{hash_num:04d}"


def extract_employer_data_from_card(card_element, page):
    """Extract employer information from a job listing card."""
    employer = {}
    try:
        # Try to get the employer name - various selectors that might work
        name_selectors = [
            '[data-testid="employer-name"]',
            '.ba-jliste-arbeitgeber',
            '.jobliste-arbeitgeber',
            'span.angebot-arbeitgeber',
            '.ergebnis-arbeitgeber',
            '[class*="arbeitgeber"]',
            '[class*="employer"]',
            'button[title]',
        ]
        
        employer_name = None
        for sel in name_selectors:
            try:
                elem = card_element.query_selector(sel)
                if elem:
                    text = elem.inner_text().strip()
                    if text and len(text) > 1:
                        employer_name = text
                        break
            except:
                continue
        
        # Fallback: try getting all text and parsing
        if not employer_name:
            try:
                all_text = card_element.inner_text()
                lines = [l.strip() for l in all_text.split('\n') if l.strip()]
                # Usually employer name is one of the first lines
                for line in lines[:5]:
                    if len(line) > 2 and not line[0].isdigit() and '€' not in line and 'Stelle' not in line:
                        employer_name = line
                        break
            except:
                pass
        
        employer['Name'] = employer_name or 'Unbekannt'
        
        # Try to get location/address
        location_selectors = [
            '[data-testid="location"]',
            '.ba-jliste-ort',
            '[class*="ort"]',
            '[class*="location"]',
            '[class*="adresse"]',
        ]
        
        location = None
        for sel in location_selectors:
            try:
                elem = card_element.query_selector(sel)
                if elem:
                    location = elem.inner_text().strip()
                    break
            except:
                continue
        
        if not location:
            try:
                all_text = card_element.inner_text()
                # Look for PLZ pattern (5 digits)
                plz_match = re.search(r'\b(\d{5})\s+([A-ZÄÖÜ][a-zäöüß]+(?:\s[A-ZÄÖÜ][a-zäöüß]+)*)', all_text)
                if plz_match:
                    location = f"{plz_match.group(1)} {plz_match.group(2)}"
            except:
                pass
        
        employer['Ort_raw'] = location or ''
        
        # Parse PLZ and Stadt from location
        if location:
            plz_match = re.search(r'(\d{5})', location)
            stadt_match = re.search(r'\d{5}\s+(.+)', location)
            employer['PLZ'] = plz_match.group(1) if plz_match else ''
            employer['Stadt'] = stadt_match.group(1).strip() if stadt_match else location
        else:
            employer['PLZ'] = ''
            employer['Stadt'] = ''
        
        # Try to get the job title (to understand the industry)
        title_selectors = [
            '[data-testid="job-title"]',
            '.ba-jliste-titel',
            'h2', 'h3', 'h4',
            '[class*="titel"]',
            'a[href*="/jobdetail"]',
        ]
        
        job_title = None
        for sel in title_selectors:
            try:
                elem = card_element.query_selector(sel)
                if elem:
                    text = elem.inner_text().strip()
                    if text and len(text) > 3:
                        job_title = text
                        break
            except:
                continue
        
        employer['Job_Titel'] = job_title or ''
        
        # Get link to detail page
        try:
            link = card_element.query_selector('a[href*="jobdetail"]')
            if link:
                employer['Detail_URL'] = link.get_attribute('href') or ''
            else:
                employer['Detail_URL'] = ''
        except:
            employer['Detail_URL'] = ''
        
        # Determine Bundesland from PLZ
        employer['Bundesland'] = get_bundesland(employer.get('PLZ', ''))
        
        # Set defaults for fields we'll fill later
        employer['Straße'] = ''
        employer['Telefon'] = ''
        employer['E-Mail'] = ''
        employer['Website'] = ''
        employer['Branche'] = 'Gastronomie/Hotel'  # Branch 11
        employer['Quelle'] = 'Arbeitsagentur'
        employer['Bewerbungsstatus'] = 'Offen'
        
        # Generate ID
        employer['Arbeitgeber_ID'] = generate_employer_id(
            employer.get('PLZ', ''), 
            employer.get('Name', '')
        )
        
        return employer
        
    except Exception as e:
        print(f"[ERROR] extracting card data: {e}")
        return None


def get_bundesland(plz):
    """Determine Bundesland from PLZ (simplified)."""
    if not plz or not plz.isdigit():
        return ''
    
    plz_int = int(plz)
    
    if 1000 <= plz_int <= 1999: return 'Sachsen'
    if 2000 <= plz_int <= 2999: return 'Sachsen'
    if 3000 <= plz_int <= 3999: return 'Brandenburg'
    if 4000 <= plz_int <= 4999: return 'Sachsen'
    if 5000 <= plz_int <= 5999: return 'Nordrhein-Westfalen'
    if 6000 <= plz_int <= 6999: return 'Sachsen-Anhalt'
    if 7000 <= plz_int <= 7999: return 'Baden-Württemberg'
    if 8000 <= plz_int <= 8999: return 'Baden-Württemberg'
    if 9000 <= plz_int <= 9999: return 'Bayern'
    if 10000 <= plz_int <= 10999: return 'Berlin'
    if 11000 <= plz_int <= 11999: return 'Berlin'
    if 12000 <= plz_int <= 12999: return 'Berlin'
    if 13000 <= plz_int <= 13999: return 'Brandenburg'
    if 14000 <= plz_int <= 14999: return 'Brandenburg'
    if 15000 <= plz_int <= 15999: return 'Brandenburg'
    if 16000 <= plz_int <= 16999: return 'Brandenburg'
    if 17000 <= plz_int <= 17999: return 'Brandenburg'
    if 18000 <= plz_int <= 18999: return 'Mecklenburg-Vorpommern'
    if 19000 <= plz_int <= 19999: return 'Mecklenburg-Vorpommern'
    if 20000 <= plz_int <= 21999: return 'Hamburg'
    if 22000 <= plz_int <= 22999: return 'Hamburg'
    if 23000 <= plz_int <= 24999: return 'Schleswig-Holstein'
    if 25000 <= plz_int <= 25999: return 'Niedersachsen'
    if 26000 <= plz_int <= 26999: return 'Niedersachsen'
    if 27000 <= plz_int <= 27999: return 'Bremen'
    if 28000 <= plz_int <= 28999: return 'Bremen'
    if 29000 <= plz_int <= 29999: return 'Niedersachsen'
    if 30000 <= plz_int <= 31999: return 'Niedersachsen'
    if 32000 <= plz_int <= 33999: return 'Nordrhein-Westfalen'
    if 34000 <= plz_int <= 35999: return 'Hessen'
    if 36000 <= plz_int <= 36999: return 'Thüringen'
    if 37000 <= plz_int <= 38999: return 'Niedersachsen'
    if 39000 <= plz_int <= 39999: return 'Sachsen-Anhalt'
    if 40000 <= plz_int <= 42999: return 'Nordrhein-Westfalen'
    if 43000 <= plz_int <= 45999: return 'Nordrhein-Westfalen'
    if 46000 <= plz_int <= 48999: return 'Nordrhein-Westfalen'
    if 49000 <= plz_int <= 49999: return 'Nordrhein-Westfalen'
    if 50000 <= plz_int <= 53599: return 'Nordrhein-Westfalen'
    if 53600 <= plz_int <= 53999: return 'Rheinland-Pfalz'
    if 54000 <= plz_int <= 57599: return 'Nordrhein-Westfalen'
    if 57600 <= plz_int <= 59999: return 'Rheinland-Pfalz'
    if 60000 <= plz_int <= 61999: return 'Hessen'
    if 63000 <= plz_int <= 63999: return 'Hessen'
    if 64000 <= plz_int <= 65999: return 'Hessen'
    if 66000 <= plz_int <= 66999: return 'Saarland'
    if 67000 <= plz_int <= 67999: return 'Rheinland-Pfalz'
    if 68000 <= plz_int <= 69999: return 'Rheinland-Pfalz'
    if 70000 <= plz_int <= 71999: return 'Baden-Württemberg'
    if 72000 <= plz_int <= 72999: return 'Baden-Württemberg'
    if 73000 <= plz_int <= 74999: return 'Baden-Württemberg'
    if 75000 <= plz_int <= 76999: return 'Baden-Württemberg'
    if 77000 <= plz_int <= 78999: return 'Baden-Württemberg'
    if 79000 <= plz_int <= 79999: return 'Baden-Württemberg'
    if 80000 <= plz_int <= 81999: return 'Bayern'
    if 83000 <= plz_int <= 83999: return 'Bayern'
    if 84000 <= plz_int <= 84999: return 'Bayern'
    if 85000 <= plz_int <= 85999: return 'Bayern'
    if 86000 <= plz_int <= 86999: return 'Bayern'
    if 87000 <= plz_int <= 87999: return 'Bayern'
    if 88000 <= plz_int <= 88999: return 'Bayern'
    if 89000 <= plz_int <= 89999: return 'Bayern'
    if 90000 <= plz_int <= 91999: return 'Bayern'
    if 92000 <= plz_int <= 92999: return 'Bayern'
    if 93000 <= plz_int <= 94999: return 'Bayern'
    if 95000 <= plz_int <= 95999: return 'Bayern'
    if 96000 <= plz_int <= 96999: return 'Bayern'
    if 97000 <= plz_int <= 97999: return 'Bayern'
    if 98000 <= plz_int <= 98999: return 'Bayern'
    if 99000 <= plz_int <= 99999: return 'Thüringen'
    
    return 'Deutschland'


def click_cookie_consent(page):
    """Try to accept cookies/consent banner."""
    consent_selectors = [
        'button#cc-accept',
        'button[data-testid="cookie-accept"]',
        'button[aria-label*="akzeptieren"]',
        'button[aria-label*="Akzeptieren"]',
        'button[aria-label*="accept"]',
        '#ba-cookie-accept',
        'button.buttongrp-item-acceptall',
        'a[data-testid="cookie-accept"]',
        'button:has-text("Alle akzeptieren")',
        'button:has-text("Akzeptieren")',
        'button:has-text("Alle auswählen")',
        '#btn-cookie-accept',
    ]
    
    for sel in consent_selectors:
        try:
            btn = page.query_selector(sel)
            if btn and btn.is_visible():
                btn.click()
                print(f"[COOKIE] Clicked consent: {sel}")
                page.wait_for_timeout(1500)
                return True
        except:
            continue
    print("[COOKIE] No consent banner found or already dismissed.")
    return False


def scrape_list_page(page, page_num):
    """Scrape employers from a listing page."""
    employers = []
    
    try:
        # Wait for job listings to load
        print(f"[PAGE {page_num}] Waiting for listings to load...")
        
        # Try multiple selectors for job cards
        card_selectors = [
            '[data-testid="job-entry"]',
            '.ba-jliste-eintrag',
            '.job-list-entry',
            '.ergebnis-liste-item',
            'article[class*="ergebnis"]',
            'div[class*="job-entry"]',
            '.search-result-item',
            'li[class*="ergebnis"]',
            '[class*="Suchergebnis"]',
            '.jobliste-item',
        ]
        
        cards = []
        for sel in card_selectors:
            try:
                page.wait_for_selector(sel, timeout=8000)
                cards = page.query_selector_all(sel)
                if cards:
                    print(f"[PAGE {page_num}] Found {len(cards)} cards with selector: {sel}")
                    break
            except:
                continue
        
        # If no specific card selectors work, try to find listing elements
        if not cards:
            print(f"[PAGE {page_num}] Trying generic selectors...")
            # Try getting all clickable job entries
            try:
                # Wait a bit more for content to load
                page.wait_for_timeout(3000)
                
                # Try to find any container with job info
                links = page.query_selector_all('a[href*="jobdetail"]')
                if links:
                    print(f"[PAGE {page_num}] Found {len(links)} job detail links")
                    for link in links:
                        try:
                            # Get the parent container
                            parent = link.evaluate_handle('el => el.closest("div, li, article, section") || el.parentElement')
                            if parent:
                                emp = extract_employer_data_from_card(parent.as_element(), page)
                                if emp and emp.get('Name') != 'Unbekannt':
                                    employers.append(emp)
                        except:
                            continue
            except Exception as e:
                print(f"[PAGE {page_num}] Generic selector error: {e}")
        
        # Extract from cards
        for card in cards:
            emp = extract_employer_data_from_card(card, page)
            if emp:
                employers.append(emp)
        
        # If still no data, dump page content for analysis
        if not cards and not employers:
            print(f"[PAGE {page_num}] No cards found. Dumping page structure...")
            body_text = page.inner_text('body')
            # Save first 3000 chars for debugging
            debug_file = f"/home/z/my-project/download/debug_page_{page_num}.txt"
            with open(debug_file, 'w', encoding='utf-8') as f:
                f.write(body_text[:5000])
            print(f"[PAGE {page_num}] Debug content saved to {debug_file}")
        
        print(f"[PAGE {page_num}] Extracted {len(employers)} employers")
        return employers
        
    except Exception as e:
        print(f"[PAGE {page_num}] Error scraping page: {e}")
        return []


def get_detail_info(page, employer):
    """Visit detail page to get more employer info."""
    detail_url = employer.get('Detail_URL', '')
    if not detail_url:
        return employer
    
    try:
        # Build full URL if needed
        if detail_url.startswith('/'):
            detail_url = f"https://www.arbeitsagentur.de{detail_url}"
        
        page.goto(detail_url, timeout=20000, wait_until='domcontentloaded')
        page.wait_for_timeout(2000)
        
        # Try to extract contact info
        text = page.inner_text('body')
        
        # Email
        email_match = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', text)
        if email_match:
            employer['E-Mail'] = email_match.group(0)
        
        # Phone
        phone_match = re.search(r'(?:Tel\.?|Telefon|Telefax)[:\s]*([\d\s/\-+()]{8,})', text, re.IGNORECASE)
        if phone_match:
            employer['Telefon'] = phone_match.group(1).strip()
        
        # Street address
        street_match = re.search(r'([A-ZÄÖÜ][a-zäöüß]+(?:straße|weg|platz|gasse|allee|ring|damm|park|ufer|hof)\.?\s*\d*\w*)', text)
        if street_match:
            employer['Straße'] = street_match.group(1).strip()
        
        # Website
        website_match = re.search(r'(?:www\.|https?://)([\w.-]+\.[a-z]{2,})', text)
        if website_match:
            website = website_match.group(0)
            if not website.startswith('http'):
                website = f"https://{website}"
            if 'arbeitsagentur' not in website.lower():
                employer['Website'] = website
        
        return employer
        
    except Exception as e:
        print(f"[DETAIL] Error getting detail for {employer.get('Name', '?')}: {e}")
        return employer


def go_to_page(page, page_num):
    """Navigate to a specific page of results."""
    try:
        # Try clicking pagination button
        pagination_selectors = [
            f'button:has-text("{page_num}")',
            f'a:has-text("{page_num}")',
            f'[data-testid="page-{page_num}"]',
            f'li[data-page="{page_num}"] a',
            '.pagination button:not([disabled])',
        ]
        
        # Try "next" button approach
        next_selectors = [
            'button[aria-label*="nächste"]',
            'button[aria-label*="Weiter"]',
            'button[aria-label*="next"]',
            'a[aria-label*="nächste"]',
            '[data-testid="pagination-next"]',
            '.pagination-next',
            'button.buttongrp-item-next',
        ]
        
        if page_num == 1:
            return True
        
        # First try direct page number click
        for sel in pagination_selectors:
            try:
                btn = page.query_selector(sel)
                if btn and btn.is_visible():
                    btn.click()
                    page.wait_for_timeout(2000)
                    return True
            except:
                continue
        
        # Then try next button
        for sel in next_selectors:
            try:
                btn = page.query_selector(sel)
                if btn and btn.is_visible():
                    btn.click()
                    page.wait_for_timeout(2000)
                    return True
            except:
                continue
        
        # Last resort: try URL manipulation
        current_url = page.url
        if 'seite=' in current_url:
            new_url = re.sub(r'seite=\d+', f'seite={page_num}', current_url)
        else:
            separator = '&' if '?' in current_url else '?'
            new_url = f"{current_url}{separator}seite={page_num}"
        
        page.goto(new_url, timeout=PAGE_TIMEOUT, wait_until='domcontentloaded')
        page.wait_for_timeout(2000)
        return True
        
    except Exception as e:
        print(f"[NAV] Error going to page {page_num}: {e}")
        return False


def main():
    global all_employers, seen_ids, current_page
    
    print("=" * 60)
    print("ARBEITSAGENTUR SCRAPER - Playwright + Xvfb")
    print("=" * 60)
    print(f"URL: {BASE_URL}")
    print(f"Start time: {datetime.now().isoformat()}")
    print()
    
    # Load any existing progress
    load_progress()
    
    with sync_playwright() as p:
        # Launch browser with virtual display
        browser = p.chromium.launch(
            headless=True,
            args=[
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-gpu',
                '--window-size=1920,1080',
            ]
        )
        
        context = browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
            locale='de-DE',
            timezone_id='Europe/Berlin',
        )
        
        page = context.new_page()
        
        # Navigate to the job search page
        print("[INIT] Navigating to Arbeitsagentur...")
        try:
            page.goto(BASE_URL, timeout=60000, wait_until='networkidle')
            print(f"[INIT] Page loaded. URL: {page.url}")
        except Exception as e:
            print(f"[INIT] Navigation warning: {e}")
            page.wait_for_timeout(5000)
            print(f"[INIT] Current URL after wait: {page.url}")
        
        # Handle cookie consent
        page.wait_for_timeout(3000)
        click_cookie_consent(page)
        page.wait_for_timeout(2000)
        
        # Save initial page for debugging
        try:
            initial_html = page.content()
            debug_file = "/home/z/my-project/download/debug_initial_page.html"
            os.makedirs(os.path.dirname(debug_file), exist_ok=True)
            with open(debug_file, 'w', encoding='utf-8') as f:
                f.write(initial_html[:50000])
            print(f"[INIT] Initial page HTML saved ({len(initial_html)} chars)")
        except Exception as e:
            print(f"[INIT] Could not save debug HTML: {e}")
        
        # Also save a screenshot
        try:
            page.screenshot(path="/home/z/my-project/download/debug_initial_screenshot.png")
            print("[INIT] Screenshot saved")
        except Exception as e:
            print(f"[INIT] Screenshot error: {e}")
        
        # Try to understand page structure
        try:
            body_text = page.inner_text('body')
            print(f"[INIT] Page body text length: {len(body_text)} chars")
            # Print first 1000 chars to understand structure
            print(f"[INIT] Page preview:\n{body_text[:1000]}")
            print("-" * 40)
        except Exception as e:
            print(f"[INIT] Could not read body text: {e}")
        
        # Check if there's a results count
        try:
            results_text = page.inner_text('body')
            count_match = re.search(r'(\d[\d.]*)\s*(?:Ergebnisse|Treffer|Stellenangebote|Stellen)', results_text)
            if count_match:
                total_results = count_match.group(1).replace('.', '')
                print(f"[INIT] Total results found: {total_results}")
        except:
            pass
        
        # Start scraping pages
        batch_count = 0
        consecutive_empty_pages = 0
        
        for page_num in range(1, MAX_PAGES + 1):
            print(f"\n{'='*40}")
            print(f"[BATCH] Processing page {page_num}/{MAX_PAGES}")
            print(f"{'='*40}")
            
            if page_num > 1:
                success = go_to_page(page, page_num)
                if not success:
                    print(f"[PAGE {page_num}] Navigation failed, stopping.")
                    break
                page.wait_for_timeout(2000)
            
            # Scrape current page
            new_employers = scrape_list_page(page, page_num)
            
            # Add new unique employers
            added = 0
            for emp in new_employers:
                eid = emp.get('Arbeitgeber_ID', '')
                if eid and eid not in seen_ids:
                    seen_ids.add(eid)
                    all_employers.append(emp)
                    added += 1
            
            print(f"[PAGE {page_num}] Added {added} new employers (total: {len(all_employers)})")
            
            if added == 0:
                consecutive_empty_pages += 1
            else:
                consecutive_empty_pages = 0
            
            # Stop if 3 consecutive empty pages
            if consecutive_empty_pages >= 3:
                print("[STOP] 3 consecutive empty pages, stopping scraper.")
                break
            
            # Save progress every BATCH_SIZE pages
            batch_count += 1
            if batch_count >= BATCH_SIZE:
                save_progress()
                print(f"[UPDATE] === BATCH PROGRESS: {len(all_employers)} employers collected so far ===")
                batch_count = 0
            
            # Random delay to be respectful
            delay = random.uniform(2, 5)
            time.sleep(delay)
        
        # Final save
        save_progress()
        print(f"\n{'='*60}")
        print(f"SCRAPING COMPLETE")
        print(f"Total employers collected: {len(all_employers)}")
        print(f"End time: {datetime.now().isoformat()}")
        print(f"{'='*60}")
        
        # Now try to get more details for employers missing contact info
        employers_need_details = [e for e in all_employers if e.get('Detail_URL') and not e.get('E-Mail') and not e.get('Telefon')]
        print(f"\n[DETAILS] {len(employers_need_details)} employers need detail page scraping...")
        
        detail_count = 0
        for emp in employers_need_details[:100]:  # Limit to first 100
            detail_count += 1
            emp = get_detail_info(page, emp)
            if detail_count % 10 == 0:
                print(f"[DETAILS] Processed {detail_count}/{min(len(employers_need_details), 100)} detail pages")
                save_progress()
            time.sleep(random.uniform(1, 3))
        
        # Final save with detail info
        save_progress()
        
        browser.close()
    
    print(f"\n[DONE] Final count: {len(all_employers)} employers")


if __name__ == '__main__':
    main()
