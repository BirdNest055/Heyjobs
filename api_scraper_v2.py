#!/usr/bin/env python3
"""
Arbeitsagentur REST API Scraper - Phase 1 continuation + Phase 2 enrichment
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
PAGE_SIZE = 25
BATCH_SIZE = 20

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
seen_hashes = set()


def load_progress():
    global all_employers, seen_hashes
    if os.path.exists(OUTPUT_FILE):
        try:
            with open(OUTPUT_FILE, 'r', encoding='utf-8') as f:
                all_employers = json.load(f)
            for emp in all_employers:
                h = emp.get('_dedup_hash', '')
                if h:
                    seen_hashes.add(h)
            print(f"[PROGRESS] Loaded {len(all_employers)} existing employers.")
        except Exception as e:
            print(f"[WARN] Could not load progress: {e}")
            all_employers = []
            seen_hashes = set()


def save_progress():
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_employers, f, ensure_ascii=False, indent=2)


def generate_employer_id(plz, name):
    clean_plz = re.sub(r'[^0-9]', '', str(plz))[:5] if plz else '00000'
    if len(clean_plz) < 5:
        clean_plz = clean_plz.zfill(5)
    hash_input = f"{name}_{plz}".lower().strip()
    hash_num = int(hashlib.md5(hash_input.encode()).hexdigest()[:4], 16) % 10000
    return f"DE-{clean_plz}-{hash_num:04d}"


def get_bundesland(region):
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
    }
    return mapping.get(region, region.replace('_', ' ').title() if region else '')


def process_job_entry(job):
    firma = job.get('firma', '').strip()
    if not firma:
        return None
    
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
    
    dedup_key = f"{firma.lower().strip()}_{plz}"
    dedup_hash = hashlib.md5(dedup_key.encode()).hexdigest()
    
    if dedup_hash in seen_hashes:
        return None
    
    seen_hashes.add(dedup_hash)
    
    full_strasse = f"{strasse} {hausnummer}".strip() if strasse else ''
    bundesland = get_bundesland(region)
    job_title = job.get('stellenangebotsTitel', '')
    hauptberuf = job.get('hauptberuf', '')
    refnr = job.get('referenznummer', '')
    ag_hash = job.get('arbeitgeberKundennummerHash', '')
    homeoffice = job.get('homeofficemoeglich', False)
    vertragsdauer = job.get('vertragsdauer', '')
    vollzeit = job.get('arbeitszeitVollzeit', False)
    teilzeit = job.get('arbeitszeitTeilzeitVormittag', False) or job.get('arbeitszeitTeilzeitNachmittag', False)
    
    return {
        'Arbeitgeber_ID': generate_employer_id(plz, firma),
        'Name': firma,
        'Straße': full_strasse,
        'PLZ': plz,
        'Stadt': ort,
        'Bundesland': bundesland,
        'Telefon': '',
        'E-Mail': '',
        'Website': '',
        'Branche': 'IT, Computer, Telekommunikation',
        'Quelle': 'Arbeitsagentur',
        'Bewerbungsstatus': 'Offen',
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


def fetch_list_page(page_num, updated_before=None):
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
            return response.json()
        elif response.status_code == 429:
            print(f"[RATE LIMIT] Waiting 30s...")
            time.sleep(30)
            return fetch_list_page(page_num, updated_before)
        else:
            print(f"[ERROR] HTTP {response.status_code}")
            return None
    except requests.exceptions.Timeout:
        print(f"[TIMEOUT]")
        return None
    except Exception as e:
        print(f"[ERROR] {e}")
        return None


def fetch_job_details(refnr):
    params = {'refnr': refnr}
    try:
        response = requests.get(DETAIL_API, headers=HEADERS, params=params, timeout=20)
        if response.status_code == 200:
            return response.json()
    except:
        pass
    return None


def extract_contact_from_details(detail_data):
    contact = {}
    if not detail_data:
        return contact
    
    # Try multiple paths for contact info
    firma_details = detail_data.get('firmaDetails', {})
    
    # Ansprechpartner
    ansprechpartner = firma_details.get('ansprechpartner', [])
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
    
    # Also check bewerbung
    bewerbung = detail_data.get('bewerbung', {})
    if bewerbung:
        if bewerbung.get('email') and not contact.get('E-Mail'):
            contact['E-Mail'] = bewerbung['email']
        if bewerbung.get('url') and not contact.get('Website'):
            contact['Website'] = bewerbung['url']
        if bewerbung.get('telefon') and not contact.get('Telefon'):
            contact['Telefon'] = bewerbung['telefon']
    
    # Employer address
    adresse = firma_details.get('adresse', {})
    if adresse and adresse.get('strasse'):
        strasse = adresse.get('strasse', '')
        hausnr = adresse.get('hausnummer', '')
        contact['Straße'] = f"{strasse} {hausnr}".strip()
        if adresse.get('plz'):
            contact['PLZ'] = adresse.get('plz')
        if adresse.get('ort'):
            contact['Stadt'] = adresse.get('ort')
    
    return contact


def main():
    print("=" * 60)
    print("ARBEITSAGENTUR REST API SCRAPER - CONTINUATION")
    print("=" * 60)
    
    load_progress()
    start_count = len(all_employers)
    
    # Phase 1: Continue collecting listings
    print(f"\nStarting with {start_count} employers. Continuing Phase 1...")
    
    total_jobs = 0
    consecutive_empty = 0
    
    for page_num in range(1, 351):
        data = fetch_list_page(page_num)
        
        if data is None:
            consecutive_empty += 1
            if consecutive_empty >= 5:
                print(f"[STOP] Too many errors.")
                break
            continue
        
        ergebnisliste = data.get('ergebnisliste', [])
        
        if not ergebnisliste:
            consecutive_empty += 1
            if consecutive_empty >= 3:
                print(f"[STOP] No more results at page {page_num}.")
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
        
        if page_num % 5 == 0:
            print(f"[PAGE {page_num}] +{new_count} new | Total: {len(all_employers)} employers | Jobs seen: {total_jobs}")
        
        if page_num % BATCH_SIZE == 0:
            save_progress()
            print(f"  === BATCH: {len(all_employers)} employers saved ===")
        
        time.sleep(random.uniform(0.3, 1.0))
    
    save_progress()
    new_from_phase1 = len(all_employers) - start_count
    print(f"\nPhase 1 done: {new_from_phase1} new employers. Total: {len(all_employers)}")
    
    # Phase 2: Enrich with contact details
    print(f"\n{'='*60}")
    print("PHASE 2: Enriching contact details")
    print(f"{'='*60}")
    
    employers_need_details = [e for e in all_employers if e.get('_referenznummer') and not e.get('E-Mail')]
    print(f"Employers needing details: {len(employers_need_details)}")
    
    enriched = 0
    emails_found = 0
    phones_found = 0
    websites_found = 0
    
    for i, emp in enumerate(employers_need_details):
        refnr = emp.get('_referenznummer', '')
        if not refnr:
            continue
        
        detail = fetch_job_details(refnr)
        contact = extract_contact_from_details(detail)
        
        updated = False
        if contact.get('E-Mail'):
            emp['E-Mail'] = contact['E-Mail']
            emails_found += 1
            updated = True
        if contact.get('Telefon'):
            emp['Telefon'] = contact['Telefon']
            phones_found += 1
            updated = True
        if contact.get('Website'):
            emp['Website'] = contact['Website']
            websites_found += 1
            updated = True
        if contact.get('Straße') and not emp.get('Straße'):
            emp['Straße'] = contact['Straße']
            updated = True
        if contact.get('Ansprechpartner'):
            emp['_ansprechpartner'] = contact['Ansprechpartner']
        
        enriched += 1
        
        if enriched % 50 == 0:
            save_progress()
            print(f"  [DETAILS] {enriched}/{len(employers_need_details)} | Emails: {emails_found} | Phones: {phones_found} | Websites: {websites_found}")
        
        time.sleep(random.uniform(0.2, 0.5))
    
    save_progress()
    
    # Final stats
    with_email = sum(1 for e in all_employers if e.get('E-Mail'))
    with_phone = sum(1 for e in all_employers if e.get('Telefon'))
    with_website = sum(1 for e in all_employers if e.get('Website'))
    with_strasse = sum(1 for e in all_employers if e.get('Straße'))
    
    print(f"\n{'='*60}")
    print(f"COMPLETE!")
    print(f"Total unique employers: {len(all_employers)}")
    print(f"With email: {with_email}")
    print(f"With phone: {with_phone}")
    print(f"With website: {with_website}")
    print(f"With address: {with_strasse}")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
