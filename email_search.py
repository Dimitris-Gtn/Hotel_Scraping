import re
import time
import os
import json
import requests
import pandas as pd
import yaml
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse, urljoin, quote

# =============================================================
# LOAD CONFIG from config.yaml
# =============================================================
with open("config.yaml", encoding="utf-8") as f:
    CONFIG = yaml.safe_load(f)

REGION_KEY = CONFIG["active_region"]
REGION_NAME = CONFIG["regions"][REGION_KEY]["name"]

SETTINGS = CONFIG["email_search"]
SAVE_EVERY = SETTINGS["save_every"]
MAX_WORKERS = SETTINGS["max_workers"]
TIMEOUT = SETTINGS["timeout"]
MAX_PAGES_PER_SITE = SETTINGS["max_pages_per_site"]

# =============================================================
# FILE PATHS (auto-named by region)
# =============================================================
os.makedirs("output", exist_ok=True)
INPUT_FILE = f"output/{REGION_KEY}_hotels.xlsx"
OUTPUT_FILE = f"output/{REGION_KEY}_with_emails.xlsx"
CLEAN_FILE = f"output/{REGION_KEY}_with_emails_clean.xlsx"

# =============================================================
# SKIP DOMAINS - OTAs, Social Media, Review sites
# =============================================================
SKIP_DOMAINS = [
    # OTAs
    'bluepillow', 'booking.com', 'expedia', 'trivago',
    'hotels.com', 'airbnb', 'agoda', 'freecancellations', 'hotelscheck-in',
    'us.despegar',
    # Social media
    'facebook.com', 'instagram.com', 'youtube.com', 'tiktok.com',
    'twitter.com', 'x.com', 'linkedin.com', 'pinterest.com',
    # Review sites
    'tripadvisor', 'google.com/maps'
]

# =============================================================
# JUNK EMAIL PATTERNS
# =============================================================
JUNK_PATTERNS = [
    'wixpress', 'sentry', 'googleapis', 'webpack', 'schema.org',
    'paradeigma', 'example', 'test@', 'info@example', 'email@',
    'your@', '@your', 'name@', '@email', 'domain.com', 'yoursite',
    '.png', '.jpg', '.jpeg', '.svg', '.webp', '.gif',
] + SKIP_DOMAINS

# =============================================================
# VALID TLDs - filters junk like user@2x.min
# =============================================================
VALID_TLDS = {
    'gr', 'com', 'net', 'org', 'eu', 'info', 'de', 'fr', 'it', 'uk',
    'co.uk', 'nl', 'es', 'ru', 'travel', 'hotel', 'gmail.com', 'io',
    'me', 'biz', 'online', 'site', 'cy', 'at', 'ch', 'be', 'se', 'dk',
}

# =============================================================
# CONTACT PAGES - targeted paths (servers are usually
# case-insensitive or redirect automatically)
# =============================================================
CONTACT_PATHS = [
    "/contact", "/contact-us", "/contacts",
    "/epikoinonia", "/" + quote("επικοινωνία"),
    "/about",
]

# Keywords for detecting contact links on the homepage
CONTACT_KEYWORDS = ['contact', 'επικοινων', 'epikoinon', 'kontakt', 'about']

# =============================================================
# HTTP Session - Reuses connections (faster)
# Full User-Agent (some sites block partial UA strings)
# =============================================================
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/120.0.0.0 Safari/537.36",
    "Accept-Language": "el-GR,el;q=0.9,en;q=0.8",
})


# =============================================================
# HELPERS: domains
# =============================================================
def get_domain(url):
    """Returns domain without www (e.g. 'myhotel.gr')."""
    try:
        netloc = urlparse(url).netloc.lower()
        return netloc[4:] if netloc.startswith('www.') else netloc
    except Exception:
        return ""


def same_domain(url_a, url_b):
    return get_domain(url_a) == get_domain(url_b) and get_domain(url_a) != ""


# =============================================================
# Cloudflare email decoder
# =============================================================
def decode_cloudflare_email(encoded):
    key = int(encoded[:2], 16)
    decoded = ""
    for i in range(2, len(encoded), 2):
        decoded += chr(int(encoded[i:i+2], 16) ^ key)
    return decoded


# =============================================================
# HELPERS: email cleaning & validation
# =============================================================
def is_valid_email(email):
    email = email.strip().lower()
    if any(j in email for j in JUNK_PATTERNS):
        return False
    # Basic structure check
    if not re.fullmatch(r'[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}', email):
        return False
    # TLD check - filters user@2x.min, package@1.2.3, etc.
    tld = email.rsplit('.', 1)[-1]
    domain = email.split('@')[-1]
    if tld not in VALID_TLDS and domain not in VALID_TLDS:
        return False
    return True


def clean_emails(emails):
    """Filter, lowercase, deduplicate, max 3."""
    clean = [e.strip().lower() for e in emails if is_valid_email(e)]
    return list(dict.fromkeys(clean))[:3]


# =============================================================
# Obfuscated emails: info [at] hotel [dot] gr → info@hotel.gr
# =============================================================
OBFUSCATED_RE = re.compile(
    r'([a-zA-Z0-9._%+\-]+)\s*[\[\(\{]?\s*(?:at|@)\s*[\]\)\}]?\s*'
    r'([a-zA-Z0-9\-]+)\s*[\[\(\{]?\s*(?:dot|\.)\s*[\]\)\}]?\s*([a-zA-Z]{2,})',
    re.IGNORECASE
)


def extract_obfuscated_emails(text):
    emails = []
    for m in OBFUSCATED_RE.finditer(text):
        candidate = f"{m.group(1)}@{m.group(2)}.{m.group(3)}"
        emails.append(candidate)
    return emails


# =============================================================
# JSON-LD (schema.org): hotels often declare their official
# email in <script type="application/ld+json">
# =============================================================
def extract_jsonld_emails(soup):
    emails = []

    def walk(obj):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k.lower() == 'email' and isinstance(v, str):
                    emails.append(v.replace('mailto:', ''))
                else:
                    walk(v)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    for script in soup.find_all('script', type='application/ld+json'):
        try:
            walk(json.loads(script.string or ""))
        except Exception:
            continue
    return emails


# =============================================================
# FUNCTION: Extract emails from HTML
# Strategies in order of reliability:
#   0. Cloudflare-protected  1. JSON-LD  2. mailto
#   3. regex  4. obfuscated ([at]/[dot])
# =============================================================
def extract_emails_from_html(html):
    soup = BeautifulSoup(html, 'html.parser')

    # STRATEGY 0: Cloudflare protected (span + href)
    cf_emails = []
    for span in soup.find_all('span', class_='__cf_email__'):
        encoded = span.get('data-cfemail', '')
        if encoded:
            try:
                cf_emails.append(decode_cloudflare_email(encoded))
            except Exception:
                pass
    for a in soup.find_all('a', href=True):
        if '/cdn-cgi/l/email-protection#' in a['href']:
            encoded = a['href'].split('#')[-1]
            if encoded:
                try:
                    cf_emails.append(decode_cloudflare_email(encoded))
                except Exception:
                    pass
    result = clean_emails(cf_emails)
    if result:
        return result

    # STRATEGY 1: JSON-LD / schema.org (almost always the official email)
    result = clean_emails(extract_jsonld_emails(soup))
    if result:
        return result

    # STRATEGY 2: mailto links
    mailto = [
        a['href'].replace('mailto:', '').split('?')[0]
        for a in soup.find_all('a', href=True)
        if a['href'].lower().startswith('mailto:')
    ]
    result = clean_emails(mailto)
    if result:
        return result

    # STRATEGY 3: regex on VISIBLE text first (less junk),
    # then on full HTML
    visible_text = soup.get_text(" ")
    for source in (visible_text, html):
        found = re.findall(
            r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}', source
        )
        result = clean_emails(found)
        if result:
            return result

    # STRATEGY 4: obfuscated - info [at] hotel [dot] gr
    result = clean_emails(extract_obfuscated_emails(visible_text))
    if result:
        return result

    return []


# =============================================================
# FUNCTION: Fetch page with redirect domain check
# Returns (html, final_url) or (None, None)
# =============================================================
def fetch_page(url, base_url=None):
    try:
        response = SESSION.get(url, timeout=TIMEOUT, allow_redirects=True)
        if response.status_code != 200:
            return None, None
        # If redirect went to a different domain (parked/OTA), ignore it
        if base_url and not same_domain(response.url, base_url):
            return None, None
        return response.text, str(response.url)
    except Exception:
        return None, None


# =============================================================
# FUNCTION: URL variants - tries https/http and www/no-www
# if the original URL fails (common in Maps data)
# =============================================================
def url_variants(base_url):
    parsed = urlparse(base_url)
    domain = get_domain(base_url)
    if not domain:
        return [base_url]
    path = parsed.path.rstrip("/")
    variants = [
        f"https://{domain}{path}",
        f"https://www.{domain}{path}",
        f"http://{domain}{path}",
    ]
    # Original first, then alternatives (no duplicates)
    return list(dict.fromkeys([base_url] + variants))


# =============================================================
# FUNCTION: Find contact URLs from sitemap.xml
# =============================================================
def contact_urls_from_sitemap(base_url):
    html, _ = fetch_page(base_url + "/sitemap.xml", base_url)
    if not html:
        return []
    urls = re.findall(r'<loc>\s*(.*?)\s*</loc>', html)
    hits = [
        u for u in urls
        if any(kw in u.lower() for kw in CONTACT_KEYWORDS)
        and same_domain(u, base_url)
    ]
    return hits[:3]


# =============================================================
# FUNCTION: Find contact links on the homepage
# =============================================================
def contact_urls_from_homepage(html, base_url):
    soup = BeautifulSoup(html, 'html.parser')
    hits = []
    for a in soup.find_all('a', href=True):
        href = a['href']
        text = a.get_text(" ").lower()
        if any(kw in text or kw in href.lower() for kw in CONTACT_KEYWORDS):
            full = urljoin(base_url + "/", href)
            if same_domain(full, base_url):
                hits.append(full.split('#')[0])
    return list(dict.fromkeys(hits))[:3]


# =============================================================
# FUNCTION: Search for email on a website
# Flow: variants → homepage → page links → sitemap → hardcoded paths
# Stops as soon as an email is found. Max MAX_PAGES_PER_SITE pages.
# =============================================================
def find_email(base_url):
    base_url = base_url.rstrip("/")

    # 1. Find which URL variant works
    html, working_url = None, None
    for variant in url_variants(base_url):
        html, final_url = fetch_page(variant, variant)
        if html:
            working_url = final_url.rstrip("/")
            # Keep the root of the final URL as base
            p = urlparse(working_url)
            working_url = f"{p.scheme}://{p.netloc}"
            break
    if not html:
        return ""

    # 2. Homepage
    emails = extract_emails_from_html(html)
    if emails:
        return ", ".join(emails)

    # 3. Gather candidate contact pages in order of reliability
    candidates = (
        contact_urls_from_homepage(html, working_url)
        + contact_urls_from_sitemap(working_url)
        + [working_url + path for path in CONTACT_PATHS]
    )
    candidates = list(dict.fromkeys(candidates))

    # 4. Try each candidate (with limit)
    tried = 0
    for url in candidates:
        if tried >= MAX_PAGES_PER_SITE:
            break
        page_html, _ = fetch_page(url, working_url)
        tried += 1
        if not page_html:
            continue
        emails = extract_emails_from_html(page_html)
        if emails:
            return ", ".join(emails)

    return ""


# =============================================================
# FUNCTION: Threading wrapper — searches a single site
# =============================================================
def process_site(idx, url):
    try:
        email = find_email(url)
        return (idx, email)
    except Exception:
        return (idx, "")


# =============================================================
# FUNCTION: Safe save (won't crash if Excel file is open)
# =============================================================
def safe_save(df, path):
    try:
        df.to_excel(path, index=False)
        return path
    except PermissionError:
        backup = path.replace(".xlsx", "_backup.xlsx")
        df.to_excel(backup, index=False)
        print(f"    ⚠️ '{path}' is open/locked — saved to '{backup}'")
        return backup


# =============================================================
# LOAD DATA - Resume if output already exists
# =============================================================
if os.path.exists(OUTPUT_FILE):
    df = pd.read_excel(OUTPUT_FILE)
    print(f"Resuming from existing file: {OUTPUT_FILE}")
else:
    df = pd.read_excel(INPUT_FILE)
    print(f"New search from: {INPUT_FILE}")

df.columns = df.columns.str.lower()

if 'email' not in df.columns:
    df['email'] = ""

# =============================================================
# FILTER: Greece only
# =============================================================
df = df[df['address'].str.contains('Greece|Ελλάδα', case=False, na=False)].reset_index(drop=True)
print(f"Total properties (Greece): {len(df)}")

# =============================================================
# FILTER: Has website, not OTA, no email found yet
# =============================================================
has_website = df[df['website'].notna()].copy()
print(f"With website: {len(has_website)}")

has_website = has_website[
    ~has_website['website'].str.lower().str.contains('|'.join(SKIP_DOMAINS), na=False)
]
print(f"With website (excluding OTA/Social): {len(has_website)}")

has_website = has_website[
    has_website['email'].isna() | (has_website['email'].astype(str).str.strip() == '')
]
print(f"Without email (to search): {len(has_website)}")

# =============================================================
# MAIN LOOP: Parallel email extraction
# =============================================================
found = 0
total = len(has_website)
tasks = [(idx, has_website.loc[idx, 'website'].rstrip("/")) for idx in has_website.index]

print(f"\nStarting search ({MAX_WORKERS} parallel threads)...")
start_time = time.time()

with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
    futures = {
        executor.submit(process_site, idx, url): (idx, url)
        for idx, url in tasks
    }

    completed = 0
    for future in as_completed(futures):
        idx, url = futures[future]
        completed += 1

        try:
            result_idx, email = future.result()
        except Exception:
            result_idx, email = idx, ""

        if email:
            df.loc[result_idx, 'email'] = email
            found += 1
            print(f"  [{completed}/{total}] ✓ {url[:50]} → {email[:40]}")
        else:
            print(f"  [{completed}/{total}] ✗ {url[:50]}")

        # Auto-save progress
        if completed % SAVE_EVERY == 0:
            safe_save(df, OUTPUT_FILE)
            elapsed = time.time() - start_time
            rate = completed / elapsed * 60
            print(f"    💾 Saved ({completed}/{total}, {found} emails, {rate:.0f} sites/min)")

# =============================================================
# FINAL SAVE
# =============================================================
elapsed = time.time() - start_time
safe_save(df, OUTPUT_FILE)

# =============================================================
# CLEAN EXPORT
# =============================================================
df['has_own_website'] = df['website'].apply(
    lambda w: 'No' if pd.isna(w) or any(s in str(w).lower() for s in SKIP_DOMAINS) else 'Yes'
)

clean_cols = ["name", "phone", "website", "email", "has_own_website"]
df_clean = df[clean_cols].sort_values("name").reset_index(drop=True)
df_clean.to_excel(CLEAN_FILE, index=False, sheet_name="Hotels")

# =============================================================
# STATS
# =============================================================
ota_count = (df['has_own_website'] == 'No').sum()

print(f"\n{'=' * 60}")
print(f"Found {found} new emails from {total} websites")
print(f"OTA/Social websites: {ota_count}")
print(f"Total properties: {len(df)}")
print(f"Time: {elapsed:.0f}s ({elapsed/60:.1f} minutes)")
print(f"Speed: {total/elapsed*60:.0f} sites/minute")
print(f"Full file: '{OUTPUT_FILE}'")
print(f"Clean export: '{CLEAN_FILE}'")