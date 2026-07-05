# 🏨 Hotel Scraping Pipeline

A pipeline for discovering Greek accommodation properties and extracting contact information for lead generation.

## Workflow

```
Step 1: hotel_search.py             →  Find properties via Google Places API (grid search)
Step 2: email_search.py             →  Extract emails from property websites
Step 3: email_search_playwright.py  →  (coming soon) Dynamic sites with JS-rendered emails
```

## Setup

```bash
git clone https://github.com/Dimitris-Gtn/Hotel_Scrapping.git
cd Hotel_Scraping
pip install -r requirements.txt
cp .env.example .env
# Add your Google Places API key to .env
```

## Usage

Set the target region in `config.yaml`:
```yaml
active_region: santorini
```

Run:
```bash
python hotel_search.py      # Step 1
python email_search.py      # Step 2
```

Results are saved to the `output/` directory.

## Adding a new region

In `config.yaml`:
```yaml
regions:
  mykonos:
    name: "Mykonos"
    bounds:
      south: 37.38
      north: 37.50
      west: 25.22
      east: 25.42
```

## API costs

- **Step 1**: Google Places API (Text Search, Advanced tier) ~$35/1000 calls
- **Step 2**: Free — direct HTTP requests to property websites

## Notes

- Both scripts support **resume** — if interrupted, they continue where they left off
- Rate limits are intentional — they protect your API key
- Excel output files are excluded from the repo (`.gitignore`)
