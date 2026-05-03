#!/usr/bin/env python3
"""
Google Maps Scraper - Erlangen 50km Radius
- Comprehensive list of ALL towns/villages within 50km of Erlangen
- Multiple search terms per location to maximize coverage
- Extensive scrolling to load all results
- Extracts: Name, Rating, Category, Address, PLZ, Phone, Website, Coordinates
- Deduplicates across all searches
"""

import time, json, re, os, sys, math
from datetime import datetime

# Erlangen center coordinates
ERLANGEN_LAT = 49.5969
ERLANGEN_LON = 11.0043
RADIUS_KM = 50

RESULTS_FILE = '/home/z/my-project/gmaps_erlangen_results.json'
PROGRESS_FILE = '/home/z/my-project/gmaps_erlangen_progress.json'

# Haversine distance calculation
def haversine(lat1, lon1, lat2, lon2):
    R = 6371  # Earth radius in km
    dlat = math.radians(lat2 - lat1)
    dlon = math.radians(lon2 - lon1)
    a = math.sin(dlat/2)**2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlon/2)**2
    return R * 2 * math.asin(math.sqrt(a))

# Comprehensive list of towns within 50km of Erlangen
# Format: (name, lat, lon, approximate_population_tier)
# Population tier: 1=large city (>100k), 2=medium (20-100k), 3=small (5-20k), 4=tiny (<5k)
ALL_TOWNS = [
    # === Major Cities ===
    ("Nürnberg", 49.4520, 11.0769, 1),
    ("Erlangen", 49.5969, 11.0043, 1),
    ("Fürth", 49.4775, 10.9886, 2),
    ("Bamberg", 49.8989, 10.8978, 2),
    ("Bayreuth", 49.9428, 11.5780, 2),  # ~62km, outside but near border
    ("Schwabach", 49.3308, 11.0225, 3),
    ("Roth", 49.2482, 11.0914, 3),
    ("Neumarkt in der Oberpfalz", 49.2753, 11.4610, 3),
    
    # === Medium Towns (5-20k) ===
    ("Herzogenaurach", 49.5682, 10.8835, 3),
    ("Fürth", 49.4775, 10.9886, 2),
    ("Forchheim", 49.7186, 11.0587, 3),
    ("Höchstadt an der Aisch", 49.7044, 10.8212, 3),
    ("Lauf an der Pegnitz", 49.5119, 11.2829, 3),
    ("Hersbruck", 49.5028, 11.4299, 3),
    ("Ebermannstadt", 49.7805, 11.1830, 3),
    ("Neustadt an der Aisch", 49.5779, 10.5923, 3),  # ~42km
    ("Eckental", 49.5681, 11.2001, 3),
    ("Weißenburg in Bayern", 49.0317, 10.9717, 3),  # ~50km
    ("Treuchtlingen", 48.9566, 10.9076, 3),  # ~55km might be slightly outside
    ("Beilngries", 49.0330, 11.4680, 3),  # ~48km
    ("Berching", 49.1264, 11.4327, 3),  # ~44km
    ("Dietfurt an der Altmühl", 49.0328, 11.5846, 3),  # ~52km borderline
    
    # === Small Towns (2-5k) near Erlangen ===
    ("Uttenreuth", 49.5872, 11.0467, 4),
    ("Buckenhof", 49.5983, 11.0233, 4),
    ("Spardorf", 49.5833, 11.0167, 4),
    ("Marloffstein", 49.5850, 11.0433, 4),
    ("Möhrendorf", 49.6133, 10.9950, 4),
    ("Heßdorf", 49.6294, 10.8703, 4),
    ("Hemhofen", 49.6389, 10.8264, 4),
    ("Adelsdorf", 49.6775, 10.8939, 4),
    ("Kalchreuth", 49.5428, 11.1708, 4),
    ("Heroldsberg", 49.5358, 11.1586, 4),
    ("Schwaig", 49.5272, 11.2114, 4),
    ("Eckental", 49.5681, 11.2001, 4),
    ("Igelsdorf", 49.5750, 11.1750, 4),
    ("Eschenau", 49.5550, 11.1883, 4),
    ("Oberschöllenbach", 49.6264, 11.0083, 4),
    ("Obermichelbach", 49.5144, 10.9583, 4),
    ("Tuchenbach", 49.5142, 10.9414, 4),
    ("Veitsbronn", 49.5069, 10.9461, 4),
    ("Seukendorf", 49.4950, 10.9281, 4),
    ("Cadolzburg", 49.4642, 10.8314, 4),
    ("Großhabersdorf", 49.4167, 10.8333, 4),
    ("Dietenhofen", 49.4036, 10.6631, 4),
    ("Oberreichenbach", 49.4125, 10.5950, 4),
    ("Bruckberg", 49.4278, 10.5531, 4),  # ~38km
    ("Weisendorf", 49.5900, 10.8467, 4),
    ("Herzogenaurach", 49.5682, 10.8835, 4),
    ("Vach", 49.5136, 10.9431, 4),
    ("Stadeln", 49.4850, 10.9500, 4),
    ("Ronhof", 49.4817, 10.9678, 4),
    ("Stein", 49.4103, 11.0017, 4),
    ("Roßtal", 49.3875, 10.9378, 4),
    ("Rohr", 49.3172, 11.0331, 4),
    ("Schwabach", 49.3308, 11.0225, 4),
    ("Feucht", 49.3764, 11.2072, 4),
    ("Wendelstein", 49.3511, 11.2775, 4),
    ("Leinburg", 49.3714, 11.3408, 4),
    ("Altdorf bei Nürnberg", 49.3897, 11.3533, 3),
    ("Pfeffenhausen", 49.3850, 11.3650, 4),
    
    # === Forchheim area ===
    ("Forchheim", 49.7186, 11.0587, 3),
    ("Eggolsheim", 49.7419, 11.0472, 4),
    ("Gosberg", 49.7164, 11.1183, 4),
    ("Kunreuth", 49.6358, 11.1275, 4),
    ("Igensdorf", 49.6264, 11.1736, 4),
    ("Gräfenberg", 49.6453, 11.2442, 4),
    ("Hiltpoltstein", 49.6542, 11.3175, 4),
    ("Betzenstein", 49.6950, 11.4039, 4),
    ("Plech", 49.6528, 11.4661, 4),
    ("Velden", 49.6164, 11.4919, 4),
    ("Happurg", 49.5233, 11.4408, 4),
    ("Pommelsbrunn", 49.4842, 11.5164, 4),
    ("Lauterhofen", 49.3622, 11.5775, 4),  # ~46km
    
    # === Hersbruck/Lauf area ===
    ("Hersbruck", 49.5028, 11.4299, 3),
    ("Lauf an der Pegnitz", 49.5119, 11.2829, 3),
    ("Ottensoos", 49.5017, 11.3447, 4),
    ("Reichenschwand", 49.5122, 11.3258, 4),
    ("Engelthal", 49.4733, 11.3700, 4),
    ("Offenhausen", 49.4419, 11.3714, 4),
    ("Kürnberg", 49.4194, 11.2647, 4),
    ("Leinburg", 49.3714, 11.3408, 4),
    
    # === Bamberg area ===
    ("Bamberg", 49.8989, 10.8978, 2),
    ("Hirschaid", 49.8117, 10.9517, 4),
    ("Strullendorf", 49.8625, 10.9375, 4),
    ("Hallstadt", 49.9328, 10.8511, 4),
    ("Gundelsheim", 49.9153, 10.8550, 4),
    ("Memmelsdorf", 49.8936, 10.9375, 4),
    ("Lichtenfels", 50.1458, 11.0653, 3),  # too far ~52km
    ("Zapfendorf", 49.9050, 10.8167, 4),
    ("Burgebrach", 49.8564, 10.7156, 4),
    ("Waischenfeld", 49.8247, 11.2119, 4),
    ("Hollfeld", 49.8961, 11.2625, 4),  # might be ~50km
    ("Aufseß", 49.8614, 11.2764, 4),
    
    # === Höchstadt area ===
    ("Höchstadt an der Aisch", 49.7044, 10.8212, 3),
    ("Gremsdorf", 49.6825, 10.7819, 4),
    ("Mühlhausen", 49.6917, 10.7833, 4),
    ("Pommersfelden", 49.7281, 10.8172, 4),
    ("Scheßlitz", 49.9108, 10.7861, 4),  # borderline
    ("Reundorf", 49.6922, 10.7889, 4),
    
    # === Neustadt/Aisch area ===
    ("Neustadt an der Aisch", 49.5779, 10.5923, 3),
    ("Ippesheim", 49.5633, 10.5667, 4),
    ("Sugenheim", 49.5853, 10.5425, 4),
    ("Markt Bibart", 49.6078, 10.4569, 4),  # ~42km
    ("Scheinfeld", 49.6561, 10.5758, 4),
    ("Langenfeld", 49.5233, 10.5750, 4),
    ("Baudenbach", 49.6033, 10.6139, 4),
    ("Diespeck", 49.5714, 10.6192, 4),
    ("Gerhardshofen", 49.5864, 10.6567, 4),
    ("Emskirchen", 49.5461, 10.7169, 4),
    ("Willhermsdorf", 49.4711, 10.6844, 4),
    
    # === West of Erlangen ===
    ("Emskirchen", 49.5461, 10.7169, 4),
    ("Oberreichenbach", 49.4125, 10.5950, 4),
    ("Puschendorf", 49.5281, 10.8664, 4),
    ("Tennenlohe", 49.5636, 11.0028, 4),
    ("Frauenaurach", 49.5911, 10.9456, 4),
    ("Kosbach", 49.5900, 10.9500, 4),
    ("Büchenbach", 49.5889, 10.9736, 4),
    ("Alterlangen", 49.5850, 10.9850, 4),
    ("Bruck", 49.5800, 10.9700, 4),
    
    # === Nürnberg surrounding areas ===
    ("Nürnberg", 49.4520, 11.0769, 1),
    ("Zirndorf", 49.4389, 10.9547, 4),
    ("Oberasbach", 49.4253, 10.9647, 4),
    ("Stein bei Nürnberg", 49.4103, 11.0017, 4),
    ("Furth im Wald", 49.4553, 10.9608, 4),  # actually Fürth area
    ("Stadeln", 49.4850, 10.9500, 4),
    ("Großgründlach", 49.5031, 11.0175, 4),
    ("Kraftshof", 49.5064, 11.0050, 4),
    ("Neunhof", 49.5111, 11.0078, 4),
    ("Schniegling", 49.4714, 10.9950, 4),
    ("Doos", 49.4450, 10.9778, 4),
    ("Schweinau", 49.4344, 11.0372, 4),
    ("Gostenhof", 49.4528, 11.0494, 4),
    ("Lederergasse", 49.4494, 11.0703, 4),
    ("Wöhrd", 49.4583, 11.0883, 4),
    ("Johannis", 49.4581, 11.0603, 4),
    ("Erlenstegen", 49.4764, 11.1142, 4),
    ("Mögeldorf", 49.4633, 11.1231, 4),
    ("Laufamholz", 49.4758, 11.1461, 4),
    ("Schmausenbuck", 49.4456, 11.1400, 4),
    ("Röthenbach bei Schweinau", 49.4231, 11.0378, 4),
    ("Gebersdorf", 49.4153, 11.0164, 4),
    ("Großreuth bei Schweinau", 49.4294, 11.0244, 4),
    ("Kleinreuth bei Schweinau", 49.4314, 11.0300, 4),
    ("Altenfurt", 49.4194, 11.1111, 4),
    ("Gibitzenhof", 49.4256, 11.0758, 4),
    ("Hasenbuck", 49.4272, 11.0911, 4),
    ("Rangierbahnhof-Siedlung", 49.4142, 11.0756, 4),
    ("Katzwang", 49.3811, 11.0294, 4),
    ("Kornburg", 49.3772, 11.0786, 4),
    ("Worzeldorf", 49.3747, 11.0575, 4),
    ("Neunhof v. Wald", 49.3942, 11.0892, 4),
    
    # === Roth/Schwabach area ===
    ("Roth", 49.2482, 11.0914, 3),
    ("Schwabach", 49.3308, 11.0225, 3),
    ("Georgensgmünd", 49.2956, 11.0156, 4),
    ("Spalt", 49.2133, 11.0314, 4),
    ("Wendelstein", 49.3511, 11.2775, 4),
    ("Rednitzhembach", 49.3069, 11.1372, 4),
    ("Büchenbach bei Roth", 49.2764, 11.0822, 4),
    ("Kammerstein", 49.2739, 10.9806, 4),
    ("Rohr bei Schwabach", 49.3172, 11.0331, 4),
    ("Hilpoltstein", 49.1908, 11.1825, 4),  # ~47km
    ("Heideck", 49.1633, 11.2575, 4),  # ~49km
    ("Thalmässing", 49.1842, 11.2317, 4),  # ~49km
    
    # === Feucht/Altdorf area ===
    ("Feucht", 49.3764, 11.2072, 4),
    ("Altdorf bei Nürnberg", 49.3897, 11.3533, 3),
    ("Leinburg", 49.3714, 11.3408, 4),
    ("Schwarzenbruck", 49.3900, 11.2414, 4),
    ("Burgthann", 49.3636, 11.3083, 4),
    ("Winkelhaid", 49.3819, 11.2811, 4),
    ("Fischbach bei Nürnberg", 49.4114, 11.1978, 4),
    ("Rückersdorf bei Nürnberg", 49.4603, 11.2331, 4),
    
    # === Neumarkt/Upper Palatinate ===
    ("Neumarkt in der Oberpfalz", 49.2753, 11.4610, 3),
    ("Freystadt", 49.1933, 11.3297, 4),
    ("Berching", 49.1264, 11.4327, 3),
    ("Beilngries", 49.0330, 11.4680, 3),
    ("Deining", 49.2700, 11.5231, 4),
    ("Sengenthal", 49.2428, 11.5228, 4),
    ("Mühlhausen", 49.2294, 11.4106, 4),
    ("Pelchenhofen", 49.2672, 11.4158, 4),
    ("Pölling", 49.2583, 11.4567, 4),
    ("Lupburg", 49.1967, 11.7028, 4),  # might be >50km
    
    # === Weißenburg area (borderline 50km) ===
    ("Weißenburg in Bayern", 49.0317, 10.9717, 3),
    ("Ellingen", 49.0497, 10.9725, 4),
    ("Pleinfeld", 49.1161, 10.9906, 4),
    ("Pappenheim", 48.9317, 10.9658, 4),
    ("Solnhofen", 48.8936, 11.0000, 4),  # ~57km borderline
    ("Treuchtlingen", 48.9566, 10.9076, 3),  # ~55km
    ("Ettenstatt", 49.0625, 10.9344, 4),
    ("Alesheim", 49.0531, 10.9017, 4),
    
    # === Additional small villages around Erlangen ===
    ("Effeltrich", 49.6283, 11.0733, 4),
    ("Poxdorf bei Erlangen", 49.6394, 11.0508, 4),
    ("Langensendelbach", 49.6378, 11.0633, 4),
    ("Neudrossenfeld", 49.9931, 11.5719, 4),  # Bayreuth area, too far
    ("Hetzles", 49.6192, 11.1078, 4),
    ("Kleinsendelbach", 49.6333, 11.0750, 4),
    ("Dietenhofen", 49.4036, 10.6631, 4),
    ("Rügland", 49.3664, 10.6389, 4),
    ("Flachslanden", 49.3767, 10.5864, 4),
    ("Oberdachstetten", 49.3592, 10.5253, 4),
    ("Lehrberg", 49.3397, 10.5617, 4),  # ~44km
    ("Ansbach", 49.3003, 10.5722, 2),  # ~55km borderline
    ("Lichtenau", 49.1417, 10.6403, 4),  # ~50km borderline
    
    # === More Nürnberg neighborhoods and suburbs ===
    ("Kraftshof", 49.5064, 11.0050, 4),
    ("Neunhof", 49.5111, 11.0078, 4),
    ("Boxdorf", 49.4647, 10.9025, 4),
    ("Vogelherd", 49.4528, 10.8972, 4),
    ("Obermichelbach", 49.5144, 10.9583, 4),
    ("Untermichelbach", 49.5078, 10.9519, 4),
    ("Siegelsdorf", 49.4689, 10.9253, 4),
    ("Ammerndorf", 49.4478, 10.9092, 4),
    ("Roßtal", 49.3875, 10.9378, 4),
    ("Buch Schwabach", 49.3539, 10.9033, 4),
    ("Mittelreichenbach", 49.3592, 10.8894, 4),
    ("Wolkersdorf", 49.5394, 11.0594, 4),
    ("Röttenbach", 49.5856, 10.8789, 4),
    ("Herzogenaurach", 49.5682, 10.8835, 4),
    ("Hammerbach", 49.5625, 10.8683, 4),
    ("Hammersbach", 49.5600, 10.8700, 4),
]

def filter_towns_by_radius(towns, radius_km=RADIUS_KM):
    """Filter towns within the radius."""
    filtered = []
    for name, lat, lon, tier in towns:
        dist = haversine(ERLANGEN_LAT, ERLANGEN_LON, lat, lon)
        if dist <= radius_km + 5:  # Small buffer for edge cases
            filtered.append({
                'name': name,
                'lat': lat,
                'lon': lon,
                'population_tier': tier,
                'distance_km': round(dist, 1)
            })
    # Sort by distance
    filtered.sort(key=lambda x: x['distance_km'])
    return filtered

# Search terms to maximize coverage of ALL businesses
SEARCH_TERMS = [
    # General business terms
    "Firma",
    "Unternehmen",
    "GmbH",
    "AG",  # Aktiengesellschaft
    
    # Industry-specific
    "IT Firma",
    "Software",
    "Handwerk",
    "Industrie",
    "Handel",
    "Dienstleistung",
    "Gastronomie",
    "Baufirma",
    "Kfz Werkstatt",
    "Recht Steuer Beratung",
    "Gesundheit",
    "Bildung",
    "Versicherung",
    "Finanzdienstleistung",
    "Immobilien",
    "Marketing Agentur",
    "Logistik",
    "Elektro Firma",
    "Metallbau",
    "Maschinenbau",
    "Kunststoff",
    "Druckerei",
    "Friseur",
    "Apotheke",
    "Arzt",
    "Anwalt",
    "Steuerberater",
]

# Reduced search terms for tiny villages (to avoid redundant results)
SMALL_TOWN_SEARCH_TERMS = [
    "Firma",
    "Unternehmen", 
    "Gewerbe",
    "Handwerk",
    "Dienstleistung",
]

def main():
    # Filter and sort towns
    towns = filter_towns_by_radius(ALL_TOWNS)
    
    # Remove duplicates by name
    seen_names = set()
    unique_towns = []
    for t in towns:
        key = t['name'].lower().strip()
        if key not in seen_names:
            seen_names.add(key)
            unique_towns.append(t)
    towns = unique_towns
    
    # Sort: start from Erlangen center, spiral outward
    towns.sort(key=lambda x: x['distance_km'])
    
    print(f"=== Google Maps Erlangen 50km Scraper ===")
    print(f"Towns within {RADIUS_KM}km of Erlangen: {len(towns)}")
    print(f"Closest: {towns[0]['name']} ({towns[0]['distance_km']}km)")
    print(f"Farthest: {towns[-1]['name']} ({towns[-1]['distance_km']}km)")
    
    # Save town list
    with open('/home/z/my-project/gmaps_erlangen_towns.json', 'w') as f:
        json.dump(towns, f, ensure_ascii=False, indent=2)
    print(f"Town list saved to gmaps_erlangen_towns.json")
    
    # Print stats
    tier_counts = {}
    for t in towns:
        tier = t['population_tier']
        tier_counts[tier] = tier_counts.get(tier, 0) + 1
    
    print(f"\nPopulation distribution:")
    print(f"  Large cities (>100k): {tier_counts.get(1, 0)}")
    print(f"  Medium towns (20-100k): {tier_counts.get(2, 0)}")
    print(f"  Small towns (5-20k): {tier_counts.get(3, 0)}")
    print(f"  Tiny villages (<5k): {tier_counts.get(4, 0)}")
    
    # Estimate total searches
    total_searches = 0
    for t in towns:
        if t['population_tier'] <= 2:
            total_searches += len(SEARCH_TERMS)
        elif t['population_tier'] == 3:
            total_searches += len(SMALL_TOWN_SEARCH_TERMS) + 5  # subset
        else:
            total_searches += len(SMALL_TOWN_SEARCH_TERMS)
    
    print(f"\nEstimated total searches: {total_searches}")
    print(f"Estimated time: {total_searches * 15 / 60 / 60:.1f} hours (at ~15s per search)")
    print(f"\nReady to start scraping. Run with: python3 gmaps_erlangen_run.py")

if __name__ == '__main__':
    main()
