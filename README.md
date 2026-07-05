# 🏨 Hotel Scraping Pipeline

A pipeline for discovering Greek accommodation properties and extracting their publicly available contact information using the Google Places API.

## Workflow

```
Step 1: hotel_search.py        →  Find properties via Google Places API (grid search)
Step 2: email_search.py        →  Extract emails from property websites
```

## Setup

```bash
git clone https://github.com/Dimitris-Gtn/Hotel_Scraping.git
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

## Responsible use

This tool collects only **publicly available business information** (names, phone numbers, websites, emails displayed on public web pages). It does not access private data, bypass authentication, or scrape personal information. Users are expected to comply with applicable data protection regulations (GDPR) and respect `robots.txt` directives.

## Notes

- Both scripts support **resume** — if interrupted, they continue where they left off
- Rate limits are intentional — they protect your API key and respect target servers
- Excel output files are excluded from the repo (`.gitignore`)
