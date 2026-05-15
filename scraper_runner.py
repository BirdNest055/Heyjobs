#!/usr/bin/env python3
"""
Scraper runner - runs v4 scraper with auto-restart on crash.
Resumes from saved progress.
"""
import subprocess
import sys
import time
import os

os.environ["DISPLAY"] = ":99"

MAX_RESTARTS = 20
WAIT_BETWEEN = 5

for attempt in range(MAX_RESTARTS):
    print(f"\n{'='*40}")
    print(f"ATTEMPT {attempt+1}/{MAX_RESTARTS}")
    print(f"{'='*40}")
    
    result = subprocess.run(
        [sys.executable, "-u", "job_application_scraper_v4.py"],
        timeout=600,  # 10 min max per attempt
        capture_output=False,
    )
    
    print(f"Exit code: {result.returncode}")
    
    if result.returncode == 0:
        print("Scraper completed successfully!")
        break
    
    print(f"Crashed with exit code {result.returncode}, restarting in {WAIT_BETWEEN}s...")
    time.sleep(WAIT_BETWEEN)
    
    # Check progress
    import json
    try:
        p = json.load(open("job_scraper_v4_progress.json"))
        print(f"  Progress: {len(p['processed'])} companies, {len(p['jobs'])} jobs")
    except:
        print("  No progress file yet")

print("\nRunner finished.")
