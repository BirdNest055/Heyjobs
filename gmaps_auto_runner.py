#!/usr/bin/env python3
"""
Auto-run Google Maps scraper for all 170 towns in batches.
Runs one batch at a time, waits for completion, then starts next.
"""
import subprocess, sys, time, json, os

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
PROGRESS_FILE = '/home/z/my-project/gmaps_erlangen_progress.json'
LOG_FILE = '/home/z/my-project/gmaps_erlangen_log.txt'

BATCH_SIZE = 5  # towns per batch

def get_stats():
    try:
        results = json.load(open(RESULTS_FILE))
        progress = json.load(open(PROGRESS_FILE))
        return len(results), len(progress['completed_searches'])
    except:
        return 0, 0

def main():
    start_from = int(sys.argv[1]) if len(sys.argv) > 1 else 15
    end_at = int(sys.argv[2]) if len(sys.argv) > 2 else 170
    
    print(f"=== Auto Runner: Towns {start_from} to {end_at} ===", flush=True)
    
    for town_start in range(start_from, end_at, BATCH_SIZE):
        results_count, searches_done = get_stats()
        print(f"\n--- Batch starting at town {town_start} (Current: {results_count} businesses, {searches_done} searches) ---", flush=True)
        
        # Run the batch
        cmd = ['python3', '/home/z/my-project/gmaps_lightweight.py', str(town_start), str(BATCH_SIZE)]
        result = subprocess.run(cmd, capture_output=False, timeout=900)  # 15 min timeout per batch
        
        if result.returncode != 0:
            print(f"Batch exited with code {result.returncode}", flush=True)
        
        # Brief pause
        time.sleep(5)
        
        # Show stats
        results_count, searches_done = get_stats()
        print(f"--- After batch: {results_count} businesses, {searches_done} searches ---", flush=True)
    
    # Final stats
    results_count, searches_done = get_stats()
    print(f"\n=== ALL DONE! {results_count} businesses collected ===", flush=True)

if __name__ == '__main__':
    main()
