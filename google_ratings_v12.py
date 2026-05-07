"""
google_ratings_v12.py
─────────────────────
Reads 100data.xlsx (output of dental_scraper.py).
For each unique practice website, fetches the Google Maps star rating
and review count.

Search strategy (tried in order):
  1. Google Places API  — by website domain       (fast, free, no bot risk)
  2. Google Places API  — by practice name + city (fallback when domain search misses)
  3. Google Maps via Playwright — by domain        (browser fallback)
  4. Google Maps via Playwright — by name + city  (last resort)

HOW TO GET A FREE GOOGLE PLACES API KEY
  1. Go to https://console.cloud.google.com/
  2. Create a project  →  Enable "Places API"
  3. Credentials  →  Create API Key
  4. Paste the key into GOOGLE_PLACES_API_KEY below
  Free tier: $200/month credit  ≈  11,000 Text Search calls — 100 practices costs ~$1.70

Bot protection (for Playwright fallback):
  - Visible browser + persistent profile  (reuses cookies across runs)
  - Stealth JS patches  (navigator.webdriver hidden)
  - Random 4–9 s delay between requests
  - Auto-pause on CAPTCHA — solve in browser, press Enter to resume
"""

import os
import re
import json
import time
import random
import logging
import warnings
from urllib.parse import quote_plus, urlparse

import requests
import openpyxl
from bs4 import BeautifulSoup
from openpyxl.styles import Font

warnings.filterwarnings("ignore", message="Unverified HTTPS request")

try:
    from playwright.sync_api import sync_playwright
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Configuration ──────────────────────────────────────────────────────────────
INPUT_FILE   = "100data.xlsx"
OUTPUT_FILE  = "100data.xlsx"

DELAY_MIN    = 6.0        # minimum seconds between Google requests
DELAY_MAX    = 13.0       # maximum (random jitter)
TIMEOUT      = 15
PW_TIMEOUT   = 35_000     # ms
PW_USER_DATA = "/tmp/pw_google_profile"  # persistent session — keeps cookies across runs

# After every BATCH_SIZE practices, pause for BATCH_PAUSE seconds
# (helps avoid Google detecting a pattern of rapid automated queries)
BATCH_SIZE   = 10
BATCH_PAUSE  = 45         # seconds

# ── Google Places API key (optional — leave "" to use Playwright only) ─────────
GOOGLE_PLACES_API_KEY = ""   # paste key here if you have credits

# ── Column indices in 100data.xlsx (1-based) ────────────────────────────────
# Row 1 = title banner, Row 2 = column headers, Row 3+ = data
C_INDEX    = 1
C_PRACTICE = 2
C_DOCTOR   = 3
C_ADDRESS  = 4
C_CITY     = 5
C_STATE    = 6
C_ZIP      = 7
C_WEBSITE  = 8    # "Practice Website"
C_GOOGLE_R = 42   # "Google Reviews Ranking"
C_GOOGLE_N = 43   # "Total # of Google Reviews"
HDR_ROW    = 2    # row containing column headers
DATA_START = 3    # first data row

_USER_AGENTS = [
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36 Edg/122.0.0.0",
]

HEADERS = {
    "User-Agent": _USER_AGENTS[0],
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://www.google.com/",
}

PLACEHOLDER = {"Not Found", "See Website", "See Profile", "Blocked", "ERROR", "", "N/A"}


# ── Helpers ────────────────────────────────────────────────────────────────────

def get_domain(website: str) -> str:
    """'https://www.example-dental.com/about' → 'example-dental.com'"""
    if not website or str(website).strip() in ("", "None", "nan"):
        return ""
    if not website.startswith("http"):
        website = "https://" + website
    host = urlparse(website).netloc.lower()
    return re.sub(r"^www\d*\.", "", host)


def _human_delay(pw_page=None):
    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
    if pw_page:
        _jitter_mouse(pw_page)


def _is_captcha(html: str) -> bool:
    lc = html.lower()
    return (
        "recaptcha" in lc
        or "captcha" in lc
        or "before you continue" in lc
        or "i'm not a robot" in lc
        or "detected unusual traffic" in lc
        or "verify you're a human" in lc
    )


def _wait_captcha(pw_page):
    log.warning("  ⚠  CAPTCHA detected — solve it in the browser window, then press Enter here…")
    input("  [Press Enter after solving CAPTCHA] ")
    pw_page.wait_for_timeout(2500)


# ── Stealth ────────────────────────────────────────────────────────────────────

_STEALTH_JS = """
(() => {
  // Hide webdriver flag
  Object.defineProperty(navigator, 'webdriver',  {get: () => undefined});
  // Fake plugin list
  Object.defineProperty(navigator, 'plugins',    {get: () => [1,2,3,4,5]});
  // Realistic language + platform
  Object.defineProperty(navigator, 'languages',  {get: () => ['en-US', 'en']});
  Object.defineProperty(navigator, 'platform',   {get: () => 'MacIntel'});
  Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
  // Chrome runtime object (missing in raw Chromium)
  window.chrome = { runtime: {}, loadTimes: function(){}, csi: function(){} };
  // Permissions API — make notifications return the real permission
  const origQuery = window.navigator.permissions.query;
  window.navigator.permissions.query = (p) =>
    p.name === 'notifications'
      ? Promise.resolve({state: Notification.permission})
      : origQuery(p);
  // Remove automation-related properties from window
  delete window.__playwright;
  delete window.__pwInitScripts;
})();
"""

def apply_stealth(ctx_or_page):
    try:
        ctx_or_page.add_init_script(_STEALTH_JS)
    except Exception:
        pass


def _jitter_mouse(pw_page):
    """Move the mouse to a random position to simulate human presence."""
    try:
        x = random.randint(200, 900)
        y = random.randint(150, 600)
        pw_page.mouse.move(x, y, steps=random.randint(5, 15))
    except Exception:
        pass


def _human_scroll(pw_page):
    """Small random scroll to simulate reading the page."""
    try:
        pw_page.evaluate(
            f"window.scrollBy(0, {random.randint(80, 300)})"
        )
        time.sleep(random.uniform(0.3, 0.8))
    except Exception:
        pass


def _prewarm_browser(pw_page):
    """
    Visit Google homepage and wait a few seconds before starting searches.
    Establishes a real-looking session with cookies before any Maps queries.
    """
    try:
        log.info("  Pre-warming browser session on Google…")
        pw_page.goto("https://www.google.com/?hl=en&gl=us", timeout=PW_TIMEOUT, wait_until="domcontentloaded")
        time.sleep(random.uniform(4, 7))
        _jitter_mouse(pw_page)
        _human_scroll(pw_page)
        time.sleep(random.uniform(2, 4))
        log.info("  Browser session ready.")
    except Exception as e:
        log.debug(f"  Pre-warm failed (non-fatal): {e}")


# ── Rating extraction ──────────────────────────────────────────────────────────

def extract_rating(html: str) -> tuple:
    """Return (rating_str, count_str) from Google search/Maps HTML. Empty strings if not found."""
    soup = BeautifulSoup(html, "lxml")

    # ── Method 1: JSON-LD aggregateRating ──────────────────────────────────────
    for script in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(script.string or "")
            if not isinstance(data, list):
                data = [data]
            for item in data:
                ar = item.get("aggregateRating") or {}
                rv = str(ar.get("ratingValue", "")).strip()
                rc = str(ar.get("reviewCount", "") or ar.get("ratingCount", "")).strip()
                if rv and re.match(r"^[1-5](\.\d)?$", rv):
                    return rv, rc
        except Exception:
            pass

    # ── Method 2: aria-label "X stars" + separate "N reviews" aria-labels ────
    found_rating, found_count = "", ""
    for tag in soup.find_all(attrs={"aria-label": True}):
        lbl = tag.get("aria-label", "")
        # Rating label: "4.8 stars" / "Rated 4.8 out of 5"
        if re.search(r"[1-5]\.\d", lbl) and re.search(r"star|rated", lbl, re.I):
            rm = re.search(r"([1-5]\.\d)", lbl)
            if rm and not found_rating:
                found_rating = rm.group(1)
                # Count sometimes in the same label
                cm = re.search(r"([\d,]+)\s*reviews?", lbl, re.I)
                if cm:
                    found_count = cm.group(1).replace(",", "")
        # Separate count label: "439 reviews" / "439 Google reviews"
        if not found_count and re.search(r"review", lbl, re.I):
            cm = re.search(r"([\d,]+)\s*reviews?", lbl, re.I)
            if cm:
                found_count = cm.group(1).replace(",", "")
    if found_rating:
        return found_rating, found_count

    # ── Method 3: data-attrid (Google knowledge panel) ────────────────────────
    for tag in soup.find_all(attrs={"data-attrid": True}):
        if "rating" in tag.get("data-attrid", ""):
            txt = tag.get_text(" ", strip=True)
            rm  = re.search(r"([1-5]\.\d)", txt)
            if rm:
                cm = re.search(r"\((\d[\d,]+)\)", txt)
                if not cm:
                    cm = re.search(r"([\d,]+)\s*reviews?", txt, re.I)
                return rm.group(1), (cm.group(1).replace(",", "") if cm else "")

    # ── Method 4: visible text patterns ───────────────────────────────────────
    text = soup.get_text(" ", strip=True)
    rm = re.search(r"\b([1-5]\.\d)\s*(?:out of 5|stars?|\()", text, re.I)
    if rm:
        g_rating = rm.group(1)
        for pat in (
            r"\(" + re.escape(g_rating) + r"[^\)]*\)\s*([\d,]+)\s*(?:Google reviews?|reviews?)",
            r"\((\d[\d,]+)\)\s*(?:Google reviews?|reviews?)",
            r"([\d,]+)\s*(?:Google reviews?|reviews?)",
            r"\((\d[\d,]+)\)",          # bare (439) — Maps result cards, no "reviews" text
        ):
            cm = re.search(pat, text, re.I)
            if cm:
                count = cm.group(1).replace(",", "")
                if g_rating == "1.0" and int(count or "0") <= 1:
                    break
                return g_rating, count
        if g_rating != "1.0":
            return g_rating, ""

    # ── Method 5: "4.8 (342)" — allow newlines between rating and count ────────
    # [^\n]{0,30} was the old pattern; [\s\S]{0,60}? handles multiline layout
    rm = re.search(r"\b([1-5]\.\d)\b[\s\S]{0,60}?\((\d[\d,]+)\)", text)
    if rm:
        return rm.group(1), rm.group(2).replace(",", "")

    # ── Method 6: any (NNN) near any rating anywhere in the text ───────────────
    for rm_r in re.finditer(r"\b([1-5]\.\d)\b", text):
        window = text[rm_r.start(): rm_r.start() + 150]
        cm = re.search(r"\((\d[\d,]+)\)", window)
        if cm:
            candidate = int(cm.group(1).replace(",", ""))
            if candidate >= 1:
                return rm_r.group(1), cm.group(1).replace(",", "")

    return "", ""


def extract_rating_maps_pw(pw_page) -> tuple:
    """Extract rating + count from an already-loaded Google Maps business panel."""
    try:
        pw_page.wait_for_selector('[role="main"]', timeout=8000)
        pw_page.wait_for_timeout(2000)
    except Exception:
        pass

    rating, count = "", ""

    # ── Step 1: get rating from aria-label; also check if count is in the same label ─
    for sel in ('[aria-label*="stars"]', '[aria-label*="Star"]', 'span.ceNzKf', 'span.MW4etd'):
        try:
            for el in pw_page.query_selector_all(sel):
                lbl = el.get_attribute("aria-label") or ""
                rm  = re.search(r"([1-5]\.\d)", lbl)
                if rm:
                    rating = rm.group(1)
                    # Count sometimes lives in the same aria-label
                    cm = re.search(r"([\d,]+)\s*reviews?", lbl, re.I)
                    if cm:
                        count = cm.group(1).replace(",", "")
                    break
            if rating:
                break
        except Exception:
            pass

    # ── Step 2: if count still missing, try dedicated review-count elements ──────
    if rating and not count:

        # 2a: element with aria-label mentioning "review"
        for sel in ('[aria-label*="review"]', '[aria-label*="Review"]'):
            try:
                for el in pw_page.query_selector_all(sel):
                    lbl = el.get_attribute("aria-label") or ""
                    cm  = re.search(r"([\d,]+)\s*reviews?", lbl, re.I)
                    if cm:
                        count = cm.group(1).replace(",", "")
                        break
                if count:
                    break
            except Exception:
                pass

        # 2b: rating/review button (jsaction="pane.rating…")
        if not count:
            for btn_sel in (
                'button[jsaction*="pane.rating"]',
                'button[jsaction*="review"]',
                'button[data-value*="review"]',
            ):
                try:
                    btn = pw_page.query_selector(btn_sel)
                    if btn:
                        btn_text = btn.inner_text() or ""
                        cm = re.search(r"([\d,]+)", btn_text)
                        if cm:
                            count = cm.group(1).replace(",", "")
                            break
                except Exception:
                    pass

        # 2c: text_content() of main panel — includes CSS-hidden nodes inner_text() misses
        if not count:
            try:
                panel = pw_page.query_selector('[role="main"]')
                if panel:
                    pt = panel.text_content() or ""
                    cm = re.search(r"\((\d[\d,]*)\)", pt)
                    if not cm:
                        cm = re.search(r"([\d,]+)\s*reviews?", pt, re.I)
                    if cm:
                        count = cm.group(1).replace(",", "")
            except Exception:
                pass

        # 2d: JavaScript extraction — scans every span/button for "(NNN)" or "NNN reviews"
        if not count:
            try:
                count = pw_page.evaluate("""
                    () => {
                        // Pass 1: element whose full text is exactly "(NNN)" or "NNN reviews"
                        for (const el of document.querySelectorAll('span, button, a')) {
                            const t = (el.textContent || '').trim();
                            let m = t.match(/^\\((\\d[\\d,]*)\\)$/);
                            if (m) return m[1].replace(/,/g,'');
                            m = t.match(/^(\\d[\\d,]*)\\s+reviews?$/i);
                            if (m) return m[1].replace(/,/g,'');
                        }
                        // Pass 2: any element containing "NNN reviews" somewhere
                        for (const el of document.querySelectorAll('span, button, a, div')) {
                            const t = (el.textContent || '').trim();
                            const m = t.match(/(\\d[\\d,]+)\\s+reviews?/i);
                            if (m && t.length < 80) return m[1].replace(/,/g,'');
                        }
                        // Pass 3: full page body text
                        const body = document.body.innerText || '';
                        const m = body.match(/(\\d[\\d,]+)\\s+reviews?/i);
                        return m ? m[1].replace(/,/g,'') : '';
                    }
                """) or ""
            except Exception:
                pass

    # ── Step 3: if still no rating, fall back to full-page HTML parse ────────────
    if not rating:
        return extract_rating(pw_page.content())

    return rating, count


# ── Search functions ───────────────────────────────────────────────────────────

def _maps_navigate_and_extract(pw_page, query: str, domain: str = "") -> tuple:
    """
    Navigate Google Maps with query, try to identify the right listing,
    click it, and return (rating, count).
    """
    url = f"https://www.google.com/maps/search/{quote_plus(query)}?hl=en&gl=us"
    _human_delay(pw_page)
    pw_page.goto(url, timeout=PW_TIMEOUT, wait_until="domcontentloaded")
    pw_page.wait_for_timeout(random.randint(2500, 4500))
    _human_scroll(pw_page)
    html = pw_page.content()

    if _is_captcha(html):
        _wait_captcha(pw_page)
        html = pw_page.content()

    # Single-result direct panel
    if "/maps/place/" in pw_page.url:
        return extract_rating_maps_pw(pw_page)

    # List of results — try to find one matching domain, else take the first
    try:
        links = pw_page.query_selector_all('a[href*="/maps/place/"]')
        target = None
        if domain:
            for link in links[:6]:
                try:
                    card_text = link.evaluate(
                        "el => el.closest('[role=\"listitem\"]')?.innerText || ''"
                    )
                    if domain in card_text.lower():
                        target = link
                        break
                except Exception:
                    pass
        if target is None and links:
            target = links[0]
        if target:
            target.click()
            pw_page.wait_for_timeout(3000)
            return extract_rating_maps_pw(pw_page)
    except Exception as e:
        log.debug(f"  Maps list parsing error: {e}")

    return extract_rating(html)


def search_maps_by_website(domain: str, practice_name: str,
                            city: str, state: str, pw_page) -> tuple:
    """Method 1 — Google Maps search using the website domain."""
    query = f"{domain} dental {city} {state}"
    log.info(f"  [Maps/domain] {query}")
    try:
        return _maps_navigate_and_extract(pw_page, query, domain)
    except Exception as e:
        log.debug(f"  Maps/domain error: {e}")
        return "", ""


def search_google_by_website(domain: str, pw_page) -> tuple:
    """Method 2 — Google Web search for the domain → knowledge panel rating."""
    query = f'"{domain}" dentist'
    url   = f"https://www.google.com/search?q={quote_plus(query)}&hl=en&gl=us"
    log.info(f"  [Google/domain] {query}")
    try:
        _human_delay(pw_page)
        pw_page.goto(url, timeout=PW_TIMEOUT, wait_until="domcontentloaded")
        pw_page.wait_for_timeout(random.randint(2000, 3500))
        _human_scroll(pw_page)
        html = pw_page.content()
        if _is_captcha(html):
            _wait_captcha(pw_page)
            html = pw_page.content()
        return extract_rating(html)
    except Exception as e:
        log.debug(f"  Google/domain error: {e}")
        return "", ""


def search_maps_by_name(practice_name: str, doctor_name: str,
                        address: str, city: str, state: str,
                        zip_code: str, pw_page) -> tuple:
    """Method 3 — Google Maps search using practice/doctor name + location."""
    clean = re.sub(r'\b(LLC|PA|PLLC|DDS|DMD|Inc\.?|Corp\.?)\b', '', practice_name, flags=re.I).strip(', ')
    doc   = re.sub(r'^Dr\.?\s+', '', doctor_name or '', flags=re.I)
    doc   = re.sub(r',?\s*(DDS|DMD|MD).*$', '', doc, flags=re.I).strip()

    if doc and address and city:
        q = f"{doc} dentist {address} {city} {state}"
    elif clean and city:
        q = f"{clean} dentist {city} {state}"
    else:
        q = f"{practice_name} dentist {city} {state}"

    log.info(f"  [Maps/name] {q}")
    try:
        return _maps_navigate_and_extract(pw_page, q)
    except Exception as e:
        log.debug(f"  Maps/name error: {e}")
        return "", ""


def _places_text_search(query: str) -> tuple:
    """
    Call Google Places Text Search API.
    Returns (rating, review_count) or ("", "").
    Costs ~$0.017 per call — covered by the free $200/month credit.
    """
    if not GOOGLE_PLACES_API_KEY:
        return "", ""
    url = "https://maps.googleapis.com/maps/api/place/textsearch/json"
    params = {
        "query":  query,
        "key":    GOOGLE_PLACES_API_KEY,
        "region": "us",
        "type":   "dentist",
    }
    try:
        r = requests.get(url, params=params, timeout=15)
        if r.status_code == 200:
            data    = r.json()
            status  = data.get("status", "")
            if status not in ("OK", "ZERO_RESULTS"):
                log.warning(f"  Places API status: {status}")
            results = data.get("results", [])
            if results:
                first  = results[0]
                rating = str(first.get("rating", "")).strip()
                # user_ratings_total can be 0 (valid) — use explicit None check
                raw_count = first.get("user_ratings_total")
                count = str(raw_count) if raw_count is not None else ""
                if rating:
                    return rating, count
    except Exception as e:
        log.debug(f"  Places API error: {e}")
    return "", ""


def places_by_website(domain: str, city: str, state: str) -> tuple:
    """Method 1 — Places API search using the website domain."""
    if not GOOGLE_PLACES_API_KEY or not domain:
        return "", ""
    query = f"{domain} dental {city} {state}"
    log.info(f"  [Places/domain] {query}")
    return _places_text_search(query)


def places_by_name(practice_name: str, city: str, state: str, zip_code: str) -> tuple:
    """Method 2 — Places API search using practice name + location."""
    if not GOOGLE_PLACES_API_KEY:
        return "", ""
    clean = re.sub(r'\b(LLC|PA|PLLC|DDS|DMD|Inc\.?|Corp\.?)\b', '', practice_name, flags=re.I).strip(', ')
    query = f"{clean} dentist {city} {state} {zip_code}".strip()
    log.info(f"  [Places/name] {query}")
    return _places_text_search(query)


# ── Write helper ───────────────────────────────────────────────────────────────

def _write_results(ws, practices: dict, results: dict):
    """Write rating + count back into every row for each practice (multi-doctor rows)."""
    for col, hdr in ((C_GOOGLE_R, "Google Reviews Ranking"),
                     (C_GOOGLE_N, "Total # of Google Reviews")):
        if not ws.cell(HDR_ROW, col).value:
            ws.cell(HDR_ROW, col).value = hdr
            ws.cell(HDR_ROW, col).font  = Font(bold=True)

    for key, p in practices.items():
        if key not in results:
            continue
        rating, count = results[key]
        for r in p["rows"]:
            ws.cell(r, C_GOOGLE_R).value = rating
            ws.cell(r, C_GOOGLE_N).value = count


# ── API-only mode (no Playwright) ─────────────────────────────────────────────

def _run_api_only(practices: dict, wb, ws):
    """Use Google Places API only — no browser required."""
    results = {}
    total   = len(practices)
    for i, (key, p) in enumerate(practices.items(), 1):
        log.info(f"[{i}/{total}]  {p['practice']}  ({p['domain']})")
        if p["google_r_cur"] not in PLACEHOLDER:
            log.info(f"  Already done: {p['google_r_cur']} ★")
            results[key] = (p["google_r_cur"], p["google_n_cur"])
            continue
        rating, count = places_by_website(p["domain"], p["city"], p["state"])
        if not rating:
            rating, count = places_by_name(p["practice"], p["city"], p["state"], p["zip"])
        results[key] = (rating or "Not Found", count or "")
        log.info(f"  → {rating or 'Not Found'} ★  ({count})")
        time.sleep(0.2)   # Places API rate-limit buffer
        if i % 10 == 0:
            _write_results(ws, practices, results)
            wb.save(OUTPUT_FILE)
            log.info(f"  💾 Progress saved after {i} practices")
    _write_results(ws, practices, results)
    wb.save(OUTPUT_FILE)
    found = sum(1 for r, c in results.values() if r not in ("Not Found", ""))
    log.info(f"\n✅ Done — {found}/{len(results)} ratings found  →  {OUTPUT_FILE}")


# ── Main ──────────────────────────────────────────────────────────────────────

def run():
    if not PLAYWRIGHT_AVAILABLE and not GOOGLE_PLACES_API_KEY:
        log.error("Playwright not installed. Run:  pip install playwright && playwright install chromium")
        log.error("Or set GOOGLE_PLACES_API_KEY for API-based lookup (no browser needed).")
        return

    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb.active

    # Build unique-website dict  →  key = domain (or raw website if domain extraction fails)
    practices = {}
    for r in range(DATA_START, ws.max_row + 1):
        if ws.cell(r, C_INDEX).value is None:
            continue
        website = str(ws.cell(r, C_WEBSITE).value or "").strip()
        if not website or website in ("None", "nan", ""):
            continue
        domain = get_domain(website)
        key    = domain or website

        if key not in practices:
            practices[key] = {
                "practice":     str(ws.cell(r, C_PRACTICE).value or "").strip(),
                "doctor":       str(ws.cell(r, C_DOCTOR).value   or "").strip(),
                "address":      str(ws.cell(r, C_ADDRESS).value  or "").strip(),
                "city":         str(ws.cell(r, C_CITY).value     or "").strip(),
                "state":        str(ws.cell(r, C_STATE).value    or "").strip(),
                "zip":          str(ws.cell(r, C_ZIP).value      or "").strip(),
                "website":      website,
                "domain":       domain,
                "google_r_cur": str(ws.cell(r, C_GOOGLE_R).value or "").strip(),
                "google_n_cur": str(ws.cell(r, C_GOOGLE_N).value or "").strip(),
                "rows":         [r],
            }
        else:
            practices[key]["rows"].append(r)

    log.info(f"Loaded {ws.max_row - DATA_START + 1} rows | {len(practices)} unique websites")

    # ── Launch Playwright (visible — required for manual CAPTCHA solving) ──────
    # Skip if Places API key is set AND Playwright is unavailable
    if not PLAYWRIGHT_AVAILABLE and GOOGLE_PLACES_API_KEY:
        log.info("Running in Places API-only mode (no browser).")
        _run_api_only(practices, wb, ws)
        return

    log.info("Launching Playwright (visible browser)…")
    # Clear old profile if it might have cached non-English Google preferences
    _prefs_file = os.path.join(PW_USER_DATA, "Default", "Preferences")
    if os.path.exists(_prefs_file):
        try:
            import json as _json
            with open(_prefs_file, encoding="utf-8") as _pf:
                _prefs = _json.load(_pf)
            _lang = (_prefs.get("intl", {}) or {}).get("accept_languages", "en-US")
            if "en" not in _lang.lower():
                import shutil
                shutil.rmtree(PW_USER_DATA, ignore_errors=True)
                log.info("  Cleared cached non-English browser profile.")
        except Exception:
            pass
    os.makedirs(PW_USER_DATA, exist_ok=True)
    _pw = sync_playwright().__enter__()
    pw_ctx = _pw.chromium.launch_persistent_context(
        user_data_dir=PW_USER_DATA,
        headless=False,
        slow_mo=150,
        ignore_https_errors=True,
        locale="en-US",                        # force English UI
        timezone_id="America/New_York",        # appear as US East Coast
        geolocation={"latitude": 40.7128, "longitude": -74.0060},  # New York City
        permissions=["geolocation"],
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-infobars",
            "--lang=en-US",                    # Chrome language flag
        ],
        user_agent=HEADERS["User-Agent"],
    )
    # Apply stealth to every page opened from this context
    apply_stealth(pw_ctx)
    pw_page = pw_ctx.new_page()
    pw_page.set_extra_http_headers({
        "Accept-Language": "en-US,en;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    })

    # Pre-warm: visit Google.com first to build a real cookie session
    _prewarm_browser(pw_page)

    results = {}

    try:
        total = len(practices)
        for i, (key, p) in enumerate(practices.items(), 1):
            log.info(f"\n[{i}/{total}]  {p['practice']}  ({p['domain']})")

            # Skip if already filled
            if p["google_r_cur"] not in PLACEHOLDER:
                log.info(f"  Already done: {p['google_r_cur']} ★ ({p['google_n_cur']})")
                results[key] = (p["google_r_cur"], p["google_n_cur"])
                continue

            rating, count = "", ""

            # ── Method 1: Google Places API by website domain (fast + free) ────
            if GOOGLE_PLACES_API_KEY and p["domain"]:
                rating, count = places_by_website(p["domain"], p["city"], p["state"])
                if rating:
                    log.info(f"  ✓ Places/domain: {rating} ★  ({count} reviews)")

            # ── Method 2: Google Places API by practice name ───────────────────
            if not rating and GOOGLE_PLACES_API_KEY:
                rating, count = places_by_name(p["practice"], p["city"], p["state"], p["zip"])
                if rating:
                    log.info(f"  ✓ Places/name: {rating} ★  ({count} reviews)")

            # ── Method 3: Google Maps via Playwright by domain ─────────────────
            if not rating and p["domain"]:
                rating, count = search_maps_by_website(
                    p["domain"], p["practice"], p["city"], p["state"], pw_page
                )
                if rating:
                    log.info(f"  ✓ Maps/domain: {rating} ★  ({count} reviews)")

            # ── Method 4: Google Maps via Playwright by name + address ─────────
            if not rating:
                rating, count = search_maps_by_name(
                    p["practice"], p["doctor"],
                    p["address"], p["city"], p["state"], p["zip"],
                    pw_page,
                )
                if rating:
                    log.info(f"  ✓ Maps/name: {rating} ★  ({count} reviews)")

            if not rating:
                log.info("  ✗ Not found after all methods")

            results[key] = (rating or "Not Found", count or "")

            # Periodic save + batch pause every BATCH_SIZE practices
            if i % BATCH_SIZE == 0:
                _write_results(ws, practices, results)
                wb.save(OUTPUT_FILE)
                log.info(f"  💾 Progress saved after {i} practices")
                if i < total:
                    log.info(f"  ⏸  Batch pause {BATCH_PAUSE}s to avoid rate-limiting…")
                    time.sleep(BATCH_PAUSE)

    finally:
        try:
            pw_ctx.close()
            _pw.__exit__(None, None, None)
        except Exception:
            pass

    _write_results(ws, practices, results)
    wb.save(OUTPUT_FILE)

    found = sum(1 for r, c in results.values() if r not in ("Not Found", ""))
    log.info(f"\n✅ Done — {found}/{len(results)} ratings found")
    log.info(f"   Output: {OUTPUT_FILE}")


if __name__ == "__main__":
    run()
