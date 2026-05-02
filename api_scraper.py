#!/usr/bin/env python3
"""
Arbeitsagentur REST API Scraper - Direct API access (much faster than browser).
API: https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v6/jobs
API Key: jobboerse-jobsuche
"""

import os
import re
import json
import time
import random
import hashlib
import requests
from datetime import datetime, timezone

# ========== Configuration ==========
API_BASE = "https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v6/jobs"
API_KEY = "jobboerse-jobsuche"
DETAIL_API = "https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v2/jobdetails"
OUTPUT_FILE = "/home/z/my-project/download/arbeitsagentur_employers.json"
PAGE_SIZE = 25  # API default page size
MAX_PAGES = 350  # 350 * 25 = 8750 (covers 8737 jobs)
BATCH_SIZE = 10  # Pages per batch before saving

HEADERS = {
    'Accept': 'application/json, text/plain, */*',
    'X-API-Key': API_KEY,
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36',
    'Accept-Language': 'de-DE,de;q=0.9,en;q=0.5',
    'Origin': 'https://www.arbeitsagentur.de',
    'Referer': 'https://www.arbeitsagentur.de/',
}

# ========== State ==========
all_employers = []
seen_hashes = set()  # Deduplicate by firma name + PLZ


def load_progress():
    """Load previously saved employers from JSON."""
    global all_employers, seen_hashes
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
                all_employers = json.load(f)
            for emp in all_employers:
                h = emp.get('_dedup_hash', '')
                if h:
                    seen_hashes.add(h)
            print(f"[PROGRESS] Loaded {len(all_employers)} existing employers from progress file.")
        except Exception as e:
            print(f"[WARN] Could not load progress: {e}")
            all_employers = []
            seen_hashes = set()


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


def get_bundesland(region):
    """Map API region enum to Bundesland name."""
    mapping = {
        'BADEN_WUERTTEMBERG': 'Baden-Württemberg',
        'BAYERN': 'Bayern',
        'BERLIN': 'Berlin',
        'BRANDENBURG': 'Brandenburg',
        'BREMEN': 'Bremen',
        'HAMBURG': 'Hamburg',
        'HESSEN': 'Hessen',
        'MECKLENBURG_VORPOMMERN': 'Mecklenburg-Vorpommern',
        'NIEDERSACHSEN': 'Niedersachsen',
        'NORDRHEIN_WESTFALEN': 'Nordrhein-Westfalen',
        'RHEINLAND_PFALZ': 'Rheinland-Pfalz',
        'SAARLAND': 'Saarland',
        'SACHSEN': 'Sachsen',
        'SACHSEN_ANHALT': 'Sachsen-Anhalt',
        'SCHLESWIG_HOLSTEIN': 'Schleswig-Holstein',
        'THUERINGEN': 'Thüringen',
        'NIEDERSACHSEN': 'Niedersachsen',
    }
    return mapping.get(region, region.replace('_', ' ').title() if region else '')


def process_job_entry(job):
    """Process a single job entry from the API and extract employer info."""
    firma = job.get('firma', '').strip()
    if not firma:
        return None
    
    # Get the first location
    locations = job.get('stellenlokationen', [])
    if not locations:
        return None
    
    loc = locations[0]
    adresse = loc.get('adresse', {})
    
    strasse = adresse.get('strasse', '')
    hausnummer = adresse.get('hausnummer', '')
    plz = adresse.get('plz', '')
    ort = adresse.get('ort', '')
    region = adresse.get('region', '')
    ortsteil = adresse.get('ortsteil', '')
    
    # Dedup hash
    dedup_key = f"{firma.lower().strip()}_{plz}"
    dedup_hash = hashlib.md5(dedup_key.encode()).hexdigest()
    
    if dedup_hash in seen_hashes:
        return None  # Skip duplicate
    
    seen_hashes.add(dedup_hash)
    
    # Build full address
    full_strasse = f"{strasse} {hausnummer}".strip() if strasse else ''
    
    # Bundesland
    bundesland = get_bundesland(region)
    
    # Job title for industry context
    job_title = job.get('stellenangebotsTitel', '')
    hauptberuf = job.get('hauptberuf', '')
    
    # Reference number
    refnr = job.get('referenznummer', '')
    
    # Employer hash (for logo lookup)
    ag_hash = job.get('arbeitgeberKundennummerHash', '')
    
    # Homeoffice
    homeoffice = job.get('homeofficemoeglich', False)
    
    # Contract type
    vertragsdauer = job.get('vertragsdauer', '')
    
    # Employment type
    vollzeit = job.get('arbeitszeitVollzeit', False)
    teilzeit = job.get('arbeitszeitTeilzeitVormittag', False) or job.get('arbeitszeitTeilzeitNachmittag', False)
    
    employer = {
        'Arbeitgeber_ID': generate_employer_id(plz, firma),
        'Name': firma,
        'Straße': full_strasse,
        'PLZ': plz,
        'Stadt': ort,
        'Bundesland': bundesland,
        'Telefon': '',  # Need detail API
        'E-Mail': '',   # Need detail API
        'Website': '',  # Need detail API
        'Branche': 'IT, Computer, Telekommunikation',  # Branch 11
        'Quelle': 'Arbeitsagentur',
        'Bewerbungsstatus': 'Offen',
        # Extra fields for enrichment
        '_referenznummer': refnr,
        '_ag_hash': ag_hash,
        '_job_titel': job_title,
        '_hauptberuf': hauptberuf,
        '_homeoffice': homeoffice,
        '_vertragsdauer': vertragsdauer,
        '_vollzeit': vollzeit,
        '_teilzeit': teilzeit,
        '_ortsteil': ortsteil,
        '_dedup_hash': dedup_hash,
    }
    
    # If there are multiple locations, add them as additional employers
    # (but only the first one is returned, others are tracked in extra)
    if len(locations) > 1:
        employer['_weitere_standorte'] = len(locations)
    
    return employer


def fetch_list_page(page_num, updated_before=None):
    """Fetch a single page of job listings from the API."""
    if updated_before is None:
        updated_before = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S.000Z')
    
    params = {
        'angebotsart': 1,
        'branche': 11,
        'page': page_num,
        'size': PAGE_SIZE,
        'aktualisiertVor': updated_before,
        'pav': 'false',
    }
    
    try:
        response = requests.get(API_BASE, headers=HEADERS, params=params, timeout=30)
        
        if response.status_code == 200:
            data = response.json()
            return data
        elif response.status_code == 429:
            print(f"[RATE LIMIT] Page {page_num}: Rate limited, waiting 30s...")
            time.sleep(30)
            return fetch_list_page(page_num, updated_before)
        else:
            print(f"[ERROR] Page {page_num}: HTTP {response.status_code}")
            print(f"  Response: {response.text[:300]}")
            return None
            
    except requests.exceptions.Timeout:
        print(f"[TIMEOUT] Page {page_num}: Request timed out")
        return None
    except Exception as e:
        print(f"[ERROR] Page {page_num}: {e}")
        return None


def fetch_job_details(refnr):
    """Fetch job detail page to get contact info (phone, email, etc.)."""
    params = {'refnr': refnr}
    
    try:
        response = requests.get(DETAIL_API, headers=HEADERS, params=params, timeout=20)
        
        if response.status_code == 200:
            return response.json()
        else:
            return None
            
    except:
        return None


def extract_contact_from_details(detail_data):
    """Extract contact info from job detail API response."""
    contact = {}
    
    if not detail_data:
        return contact
    
    # Navigate the detail structure
    firma_details = detail_data.get('firmaDetails', {})
    ansprechpartner = firma_details.get('ansprechpartner', [])
    
    if ansprechpartner:
        for person in ansprechpartner:
            if person.get('email'):
                contact['E-Mail'] = person['email']
            if person.get('telefon'):
                contact['Telefon'] = person['telefon']
            if person.get('vorname') or person.get('nachname'):
                contact['Ansprechpartner'] = f"{person.get('vorname', '')} {person.get('nachname', '')}".strip()
    
    # Employer website
    url = firma_details.get('url', '')
    if url:
        contact['Website'] = url
    
    # Employer address (more detailed)
    adresse = firma_details.get('adresse', {})
    if adresse:
        if adresse.get('strasse') and not contact.get('Straße'):
            strasse = adresse.get('strasse', '')
            hausnr = adresse.get('hausnummer', '')
            contact['Straße'] = f"{strasse} {hausnr}".strip()
    
    return contact


def main():
    print("=" * 60)
    print("ARBEITSAGENTUR REST API SCRAPER")
    print("=" * 60)
    print(f"API: {API_BASE}")
    print(f"Start time: {datetime.now().isoformat()}")
    print()
    
    # Load any existing progress
    load_progress()
    
    # Phase 1: Collect all job listings (employer names + addresses)
    print("\n" + "=" * 60)
    print("PHASE 1: Collecting job listings from API")
    print("=" * 60)
    
    total_jobs = 0
    total_new = 0
    consecutive_empty = 0
    last_page_processed = 0
    
    for page_num in range(1, MAX_PAGES + 1):
        print(f"\n[PAGE {page_num}] Fetching...", end=" ", flush=True)
        
        data = fetch_list_page(page_num)
        
        if data is None:
            consecutive_empty += 1
            if consecutive_empty >= 5:
                print(f"\n[STOP] {consecutive_empty} consecutive errors, stopping.")
                break
            continue
        
        # Parse results
        ergebnisliste = data.get('ergebnisliste', [])
        
        if not ergebnisliste:
            print(f"Empty page (no results)")
            consecutive_empty += 1
            if consecutive_empty >= 3:
                print("[STOP] No more results, stopping.")
                break
            continue
        
        consecutive_empty = 0
        total_jobs += len(ergebnisliste)
        
        new_count = 0
        for job in ergebnisliste:
            emp = process_job_entry(job)
            if emp:
                all_employers.append(emp)
                new_count += 1
                total_new += 1
        
        print(f"Jobs: {len(ergebnisliste)}, New employers: {new_count}, Total: {len(all_employers)}")
        last_page_processed = page_num
        
        # Save progress every BATCH_SIZE pages
        if page_num % BATCH_SIZE == 0:
            save_progress()
            print(f"  === BATCH UPDATE: {len(all_employers)} employers after page {page_num} ===")
        
        # Small delay to be respectful
        time.sleep(random.uniform(0.5, 1.5))
    
    # Save after Phase 1
    save_progress()
    print(f"\n{'='*60}")
    print(f"PHASE 1 COMPLETE")
    print(f"Jobs processed: {total_jobs}")
    print(f"Unique employers: {len(all_employers)}")
    print(f"Pages scraped: {last_page_processed}")
    print(f"{'='*60}")
    
    # Phase 2: Enrich with contact details from detail API
    print("\n" + "=" * 60)
    print("PHASE 2: Enriching employer contact details")
    print("=" * 60)
    
    employers_need_details = [e for e in all_employers if e.get('_referenznummer') and not e.get('E-Mail')]
    print(f"Employers needing detail enrichment: {len(employers_need_details)}")
    
    enriched_count = 0
    email_count = 0
    phone_count = 0
    website_count = 0
    
    # Process in batches of 50 for progress updates
    for i, emp in enumerate(employers_need_details):
        refnr = emp.get('_referenznummer', '')
        if not refnr:
            continue
        
        detail = fetch_job_details(refnr)
        contact = extract_contact_from_details(detail)
        
        if contact.get('E-Mail'):
            emp['E-Mail'] = contact['E-Mail']
            email_count += 1
        if contact.get('Telefon'):
            emp['Telefon'] = contact['Telefon']
            phone_count += 1
        if contact.get('Website'):
            emp['Website'] = contact['Website']
            website_count += 1
        if contact.get('Straße') and not emp.get('Straße'):
            emp['Straße'] = contact['Straße']
        if contact.get('Ansprechpartner'):
            emp['_ansprechpartner'] = contact['Ansprechpartner']
        
        enriched_count += 1
        
        if enriched_count % 50 == 0:
            save_progress()
            print(f"  [DETAILS] {enriched_count}/{len(employers_need_details)} processed | Emails: {email_count} | Phones: {phone_count} | Websites: {website_count}")
        
        time.sleep(random.uniform(0.3, 0.8))
    
    # Final save
    save_progress()
    
    print(f"\n{'='*60}")
    print(f"SCRAPING COMPLETE")
    print(f"Total unique employers: {len(all_employers)}")
    print(f"With email: {email_count}")
    print(f"With phone: {phone_count}")
    print(f"With website: {website_count}")
    print(f"End time: {datetime.now().isoformat()}")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
