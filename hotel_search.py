import os
import json
import time
import math
import requests
import pandas as pd
import yaml
from dotenv import load_dotenv

# =============================================================
# LOAD API KEY from .env file
# =============================================================
load_dotenv()
API_KEY = os.getenv("GOOGLE_PLACES_API_KEY")

# =============================================================
# LOAD CONFIG from config.yaml
# =============================================================
with open("config.yaml", encoding="utf-8") as f:
    CONFIG = yaml.safe_load(f)

REGION_KEY = CONFIG["active_region"]
REGION = CONFIG["regions"][REGION_KEY]
REGION_NAME = REGION["name"]
BOUNDS = REGION["bounds"]

SETTINGS = CONFIG["hotel_search"]
SEARCH_RADIUS_KM = SETTINGS["radius_km"]
MIN_RADIUS_KM = SETTINGS["min_radius_km"]
SAVE_EVERY = SETTINGS["save_every"]
SLEEP_BETWEEN_REQUESTS = SETTINGS["sleep_between_requests"]
REQUEST_TIMEOUT = SETTINGS["request_timeout"]
SATURATION_LIMIT = SETTINGS["saturation_limit"]

# =============================================================
# GOOGLE PLACES API (New) - URL & Headers
# FieldMask: only the 5 fields we need + id
# (id is required for dedup, no extra tier cost)
# =============================================================
SEARCH_URL = "https://places.googleapis.com/v1/places:searchText"

HEADERS = {
    "Content-Type": "application/json",
    "X-Goog-Api-Key": API_KEY,
    "X-Goog-FieldMask": (
        "places.id,"
        "places.displayName,"
        "places.formattedAddress,"
        "places.internationalPhoneNumber,"
        "places.websiteUri,"
        "places.rating"
    )
}

# =============================================================
# OUTPUT FILES (auto-named by region)
# =============================================================
os.makedirs("output", exist_ok=True)
PROGRESS_FILE = f"output/{REGION_KEY}_progress.xlsx"
STATE_FILE = f"output/{REGION_KEY}_state.json"
OUTPUT_FILE = f"output/{REGION_KEY}_hotels.xlsx"

# =============================================================
# QUERIES: English + Greek + common accommodation types
# Greek queries catch listings with Greek-only names
# =============================================================
QUERIES = [
    "hotel", "boutique hotel", "apartment hotel", "aparthotel",
    "apartment", "villa", "resort", "studios",
    "rooms to let", "guesthouse", "pension", "bed and breakfast",
    "bungalows", "suites",
    "ξενοδοχείο", "ενοικιαζόμενα δωμάτια", "βίλα", "ξενώνας",
]

# =============================================================
# COST: Text Search with phone/website/rating = Advanced tier
# Price per 1000 calls (USD) - update if Google changes pricing
# =============================================================
COST_PER_1000_CALLS = 35.0

# API call counter (for cost tracking)
api_calls = 0


# =============================================================
# FUNCTION: Create grid of points covering the region
# Each point is spaced sqrt(2)*radius apart so the rectangular
# restrictions overlap slightly, leaving no gaps
# =============================================================
def create_grid(bounds, radius_km):
    # 1 degree latitude ≈ 111 km
    # 1 degree longitude ≈ 111 * cos(lat) km
    step_km = radius_km * 1.9  # slight overlap (not exactly 2x)
    lat_step = step_km / 111
    avg_lat = (bounds["south"] + bounds["north"]) / 2
    lng_step = step_km / (111 * math.cos(math.radians(avg_lat)))

    points = []
    lat = bounds["south"]
    while lat <= bounds["north"]:
        lng = bounds["west"]
        while lng <= bounds["east"]:
            points.append((round(lat, 4), round(lng, 4)))
            lng += lng_step
        lat += lat_step

    return points


# =============================================================
# FUNCTION: Convert (center, radius) → rectangular restriction
# searchText accepts locationRestriction ONLY as a rectangle.
# Restriction (unlike bias) strictly clips results to the area
# → all 20 slots are used for local results only.
# =============================================================
def make_rectangle(lat, lng, radius_km):
    dlat = radius_km / 111
    dlng = radius_km / (111 * math.cos(math.radians(lat)))
    return {
        "rectangle": {
            "low":  {"latitude": lat - dlat, "longitude": lng - dlng},
            "high": {"latitude": lat + dlat, "longitude": lng + dlng},
        }
    }


# =============================================================
# FUNCTION: Single API call with retry for 429 AND network errors
# Returns list of places or None on permanent failure
# =============================================================
def api_search(query, lat, lng, radius_km):
    global api_calls

    body = {
        "textQuery": query,
        "maxResultCount": 20,
        "locationRestriction": make_rectangle(lat, lng, radius_km),
    }

    for attempt in range(3):
        try:
            response = requests.post(
                SEARCH_URL, headers=HEADERS, json=body, timeout=REQUEST_TIMEOUT
            )
            api_calls += 1
        except requests.RequestException as e:
            # Network error (timeout, DNS, etc.) → short wait and retry
            print(f"    ⚠️ Network error ({type(e).__name__}), retry {attempt + 1}/3...")
            time.sleep(5)
            continue

        if response.status_code == 429:
            # Rate limited → longer wait and retry
            print(f"    ⏳ Rate limited, waiting 30s (attempt {attempt + 1}/3)...")
            time.sleep(30)
            continue

        if response.status_code != 200:
            print(f"    Error {response.status_code}: {response.text[:100]}")
            return None

        return response.json().get("places", [])

    # All 3 attempts failed
    print(f"    ✗ Failed after 3 attempts")
    return None


# =============================================================
# FUNCTION: Convert place object to flat dict with 5 columns
# =============================================================
def parse_place(place):
    return {
        "place_id": place.get("id", ""),  # for dedup only, not exported
        "name": place.get("displayName", {}).get("text", ""),
        "address": place.get("formattedAddress", ""),
        "phone": place.get("internationalPhoneNumber", ""),
        "website": place.get("websiteUri", ""),
        "rating": place.get("rating", ""),
    }


# =============================================================
# FUNCTION: Search a single point with ALL queries
# + ADAPTIVE SUBDIVISION:
#   If any query returns 20/20 results, the area is saturated
#   (more results exist but are hidden). Split the point into
#   4 sub-points with half the radius and search recursively.
#   This ensures no properties are missed in dense areas.
# =============================================================
def search_point(lat, lng, radius_km, all_hotels, unique_ids, depth=0):
    indent = "    " * (depth + 1)
    saturated = False
    empty_queries = 0

    for qi, query in enumerate(QUERIES):
        places = api_search(query, lat, lng, radius_km)
        time.sleep(SLEEP_BETWEEN_REQUESTS)

        if places is None:
            continue  # Call failed — move to next query

        # Sea point detection: skip only if the first 2 queries both
        # return 0 (coastal villages may have only studios/villas)
        if len(places) == 0:
            empty_queries += 1
            if qi <= 1 and empty_queries == 2:
                print(f"{indent}Likely sea point, skipping")
                return
            continue

        # Dedup: keep only new properties by place_id
        new_count = 0
        for place in places:
            pid = place.get("id", "")
            if pid and pid not in unique_ids:
                unique_ids.add(pid)
                all_hotels.append(parse_place(place))
                new_count += 1

        # Saturation check
        if len(places) >= SATURATION_LIMIT:
            saturated = True

        print(f"{indent}'{query}': {len(places)} found, {new_count} new | Total: {len(all_hotels)}")

    # RECURSION: if saturated and above minimum radius,
    # split into 4 sub-points (quadrants) with half the radius
    if saturated and radius_km / 2 >= MIN_RADIUS_KM:
        half = radius_km / 2
        dlat = half / 111
        dlng = half / (111 * math.cos(math.radians(lat)))
        print(f"{indent}🔍 Saturated! Subdividing into 4 sub-points (radius {half}km)")
        for sub_lat, sub_lng in [
            (lat - dlat, lng - dlng), (lat - dlat, lng + dlng),
            (lat + dlat, lng - dlng), (lat + dlat, lng + dlng),
        ]:
            search_point(round(sub_lat, 4), round(sub_lng, 4), half,
                         all_hotels, unique_ids, depth + 1)


# =============================================================
# FUNCTIONS: Resume system
# Saves (a) hotels to Excel and (b) completed grid points to
# JSON. If the script crashes, it restarts from where it left
# off — no duplicate API costs.
# =============================================================
def save_progress(all_hotels, done_points):
    if all_hotels:
        pd.DataFrame(all_hotels).to_excel(PROGRESS_FILE, index=False)
    with open(STATE_FILE, "w", encoding="utf-8") as f:
        json.dump({"done_points": sorted(done_points)}, f)


def load_progress():
    all_hotels, unique_ids, done_points = [], set(), set()
    if os.path.exists(PROGRESS_FILE) and os.path.exists(STATE_FILE):
        df = pd.read_excel(PROGRESS_FILE)
        all_hotels = df.to_dict("records")
        unique_ids = set(df["place_id"].dropna().astype(str))
        with open(STATE_FILE, encoding="utf-8") as f:
            done_points = set(json.load(f).get("done_points", []))
        print(f"↻ Resuming: {len(all_hotels)} properties, {len(done_points)} points already done")
    return all_hotels, unique_ids, done_points


# =============================================================
# MAIN
# =============================================================
def main():
    print("=" * 60)
    print(f"HOTEL SEARCH — {REGION_NAME.upper()} (GRID SEARCH)")
    print("=" * 60)

    if not API_KEY:
        print("✗ GOOGLE_PLACES_API_KEY not found in .env")
        return

    # =============================================================
    # STEP 1: Create grid + load any previous progress
    # =============================================================
    grid_points = create_grid(BOUNDS, SEARCH_RADIUS_KM)
    print(f"Grid: {len(grid_points)} points (radius {SEARCH_RADIUS_KM}km, "
          f"{len(QUERIES)} queries/point)")

    all_hotels, unique_ids, done_points = load_progress()
    start_time = time.time()

    # =============================================================
    # STEP 2: Search each point (with subdivision where needed)
    # =============================================================
    for i, (lat, lng) in enumerate(grid_points, 1):
        point_key = f"{lat},{lng}"

        # Skip points completed in a previous run
        if point_key in done_points:
            continue

        print(f"\n[{i}/{len(grid_points)}] Point ({lat}, {lng})")
        search_point(lat, lng, SEARCH_RADIUS_KM, all_hotels, unique_ids)
        done_points.add(point_key)

        # Auto-save progress every N points
        if i % SAVE_EVERY == 0:
            save_progress(all_hotels, done_points)
            est_cost = api_calls / 1000 * COST_PER_1000_CALLS
            print(f"    💾 Saved ({len(all_hotels)} hotels, {api_calls} calls, ~${est_cost:.2f})")

    # Final progress save (safety net before export)
    save_progress(all_hotels, done_points)

    # =============================================================
    # STEP 3: Filter, sort, export to 5 columns
    # =============================================================
    print(f"\n{'=' * 60}")
    print(f"Total unique properties (raw): {len(all_hotels)}")

    df = pd.DataFrame(all_hotels)

    # Filter: Greece only (English + Greek addresses)
    df = df[df['address'].str.contains('Greece|Ελλάδα', case=False, na=False)]
    print(f"After country filter: {len(df)}")

    # Sort by name + export only the 5 required columns
    df = df.sort_values("name").reset_index(drop=True)
    df_export = df[["name", "address", "phone", "website", "rating"]]
    df_export.to_excel(OUTPUT_FILE, index=False, sheet_name="Hotels")

    # =============================================================
    # STATS + cost estimate
    # =============================================================
    elapsed = time.time() - start_time
    est_cost = api_calls / 1000 * COST_PER_1000_CALLS

    print(f"\n{'=' * 60}")
    print(f"Saved to '{OUTPUT_FILE}'")
    print(f"Total properties: {len(df_export)}")
    print(f"With phone: {(df_export['phone'].astype(str).str.strip() != '').sum()}")
    print(f"With website: {(df_export['website'].astype(str).str.strip() != '').sum()}")
    print(f"API calls: {api_calls} (~${est_cost:.2f} at ${COST_PER_1000_CALLS}/1000)")
    print(f"Time: {elapsed/60:.1f} minutes")

    # Clean up progress files after successful completion
    for f in (PROGRESS_FILE, STATE_FILE):
        if os.path.exists(f):
            os.remove(f)
    print("Progress files removed.")


if __name__ == "__main__":
    main()