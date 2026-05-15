#!/usr/bin/env python3
"""
Continuous runner for job_application_scraper_v5.py
Processes companies in batches, restarting fresh each batch.
"""
import json
import os
import subprocess
import sys
import time

PROGRESS_FILE = "/home/z/my-project/job_scraper_v5_progress.json"
BATCH_SIZE = 3
MAX_BATCHES = 300
TOTAL_COMPANIES = 1024

os.environ["DISPLAY"] = ":99"

for i in range(MAX_BATCHES):
    # Check progress
    try:
        p = json.load(open(PROGRESS_FILE))
        processed = len(p["processed"])
        jobs = len(p["jobs"])
        print(f"\n=== BATCH {i+1} | Progress: {processed} companies, {jobs} jobs ===", flush=True)
        if processed >= TOTAL_COMPANIES:
            print("All companies processed!")
            break
    except:
        print(f"\n=== BATCH {i+1} ===", flush=True)

    # Run one batch
    result = subprocess.run(
        [sys.executable, "-u", "job_application_scraper_v5.py", str(BATCH_SIZE)],
        timeout=300,
        cwd="/home/z/my-project"
    )
    print(f"Exit code: {result.returncode}", flush=True)

    # Brief pause
    time.sleep(2)

print("\nScraper runner finished.")
