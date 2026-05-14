#!/usr/bin/env python3
"""Single-employer scraper that runs in a subprocess."""
import os, sys, json, time, traceback
os.environ['DISPLAY'] = ':99'
os.environ['PYTHONUNBUFFERED'] = '1'

from junior_jobs_scraper_v4 import *

def main():
    # Read employer data from stdin
    employer_json = sys.stdin.read()
    employer = json.loads(employer_json)
    
    print(f"SCRAPING: {employer['name'][:40]}", flush=True)
    
    excel = JuniorJobsExcel(EXCEL_OUT)
    scraper = CloakScraper(excel)
    
    try:
        scraper.start_browser()
        jobs = scraper.scrape_employer(employer)
        excel.save()
        print(f"RESULT: {len(jobs)} jobs", flush=True)
    except Exception as e:
        print(f"ERROR: {str(e)[:80]}", flush=True)
        traceback.print_exc()
    finally:
        scraper.stop_browser()

if __name__ == '__main__':
    main()
