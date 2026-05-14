#!/bin/bash
cd /home/z/my-project

for CITY in "Ebermannstadt" "Roth" "Altdorf bei Nürnberg" "Zirndorf" "Strullendorf" "Höchstadt an der Aisch"; do
    echo "=== Processing $CITY ==="
    pkill -9 Xvfb 2>/dev/null
    pkill -9 chromium 2>/dev/null
    pkill -9 chrome 2>/dev/null
    rm -f /tmp/.X99-lock
    sleep 3
    
    timeout 540 python3 gmaps_one_city.py "$CITY" 2>&1 | tail -5
    echo "=== $CITY done ==="
    sleep 5
done

echo "ALL CITIES COMPLETE!"
