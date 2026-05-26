#!/usr/bin/env python3
"""
bypass_scraper.py
=================
Re-scrape dental practices that were blocked by bot protection in the main scraper.

BYPASS STRATEGIES (tried in order per request):
  1. curl_cffi — rotates through 6 Chrome/Safari TLS fingerprint profiles
  2. curl_cffi + proxy rotation — same profiles but through a proxy
  3. Playwright + stealth JS — hides automation markers, no proxy
  4. Playwright + stealth JS + proxy — last resort

USAGE
-----
  # Re-scrape specific URLs:
  python bypass_scraper.py --urls https://site1.com https://site2.com

  # Auto-detect bot-blocked rows in a batch output Excel and re-scrape:
  python bypass_scraper.py --batch batch_05_deduped.xlsx

  # With a proxy list:
  python bypass_scraper.py --batch batch_05_deduped.xlsx --proxies proxies.txt

  # Custom output file:
  python bypass_scraper.py --batch batch_05_deduped.xlsx --output retried.xlsx

PROXIES FILE (proxies.txt) — one per line:
  http://host:port
  http://user:pass@host:port
  socks5://host:port

INSTALL
-------
  pip install curl_cffi playwright openpyxl requests beautifulsoup4 lxml
  playwright install chromium
"""

import argparse
import random
import sys
import time
import re
import os
from urllib.parse import urlparse

import warnings
warnings.filterwarnings("ignore")

# ── Dependencies ──────────────────────────────────────────────────────────────

try:
    import openpyxl
except ImportError:
    sys.exit("pip install openpyxl")

try:
    import requests as std_requests
except ImportError:
    sys.exit("pip install requests")

try:
    from curl_cffi import requests as cffi_requests
    _CFFI_OK = True
except ImportError:
    _CFFI_OK = False
    print("WARNING: curl_cffi not installed — proxy+TLS bypass disabled")
    print("         pip install curl_cffi")

try:
    from playwright.sync_api import sync_playwright
    _PW_OK = True
except ImportError:
    _PW_OK = False
    print("WARNING: playwright not installed — JS-site bypass disabled")
    print("         pip install playwright && playwright install chromium")

# ── Import dental_scraper functions ──────────────────────────────────────────

try:
    import dental_scraper as ds
except ImportError:
    sys.exit("dental_scraper.py must be in the same folder")

# ── Stealth JS (injected into every Playwright page) ─────────────────────────

_STEALTH_JS = """
(function(){
  Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
  window.chrome = {
    app:{ isInstalled:false, getDetails:function(){}, getIsInstalled:function(){}, runningState:function(){} },
    csi:function(){}, loadTimes:function(){},
    runtime:{ OnInstalledReason:{}, PlatformArch:{}, PlatformNaclArch:{}, PlatformOs:{}, RequestUpdateCheckStatus:{} }
  };
  const fp=[
    {name:'Chrome PDF Plugin',  filename:'internal-pdf-viewer',              description:'Portable Document Format'},
    {name:'Chrome PDF Viewer',  filename:'mhjfbmdgcfjbbpaeojofohoefgiehjai', description:''},
    {name:'Native Client',      filename:'internal-nacl-plugin',              description:''},
  ];
  Object.defineProperty(navigator,'plugins',{get:()=>{
    const a=fp.map(p=>Object.assign(Object.create(Plugin.prototype),p));
    Object.setPrototypeOf(a,PluginArray.prototype); return a;
  }});
  Object.defineProperty(navigator,'languages',{get:()=>['en-US','en']});
  try{
    const oq=window.navigator.permissions.query.bind(navigator.permissions);
    window.navigator.permissions.query=(p)=>
      p.name==='notifications'?Promise.resolve({state:Notification.permission}):oq(p);
  }catch(e){}
  try{
    const gp=WebGLRenderingContext.prototype.getParameter;
    WebGLRenderingContext.prototype.getParameter=function(p){
      if(p===37445)return 'Intel Inc.';
      if(p===37446)return 'Intel Iris OpenGL Engine';
      return gp.call(this,p);
    };
  }catch(e){}
  delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
  delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
  delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
})();
"""

# curl_cffi browser profiles to rotate through
_CFFI_PROFILES = ["chrome136", "chrome124", "chrome133a", "chrome110", "safari260", "safari17_2"]

# Markers that indicate a bot-challenge page was returned (even at HTTP 200)
_CHALLENGE_MARKERS = (
    "sgcaptcha", "robot challenge", "just a moment", "checking your browser",
    "access denied", "please enable cookies", "enable javascript and cookies",
)


def _is_challenge_html(r):
    """Return True if the response is a bot-challenge page regardless of status code."""
    if r.status_code in (202, 403):
        return True
    url_lower = r.url.lower()
    if "sgcaptcha" in url_lower or "captcha" in url_lower or "challenge" in url_lower:
        return True
    snippet = r.text[:2000].lower()
    return any(m in snippet for m in _CHALLENGE_MARKERS)

# ── Proxy pool ────────────────────────────────────────────────────────────────

_proxy_pool: list = []
_proxy_fail:  set  = set()   # proxies that failed — skip for a while


def load_proxies(path):
    global _proxy_pool
    if not path or not os.path.exists(path):
        return
    with open(path) as f:
        _proxy_pool = [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
    print(f"Loaded {len(_proxy_pool)} proxies from {path}")


def _next_proxy():
    """Return a random working proxy or None."""
    available = [p for p in _proxy_pool if p not in _proxy_fail]
    return random.choice(available) if available else None


# ── Enhanced safe_get (monkey-patched over ds.safe_get) ───────────────────────

_orig_safe_get = None


def _bypass_safe_get(url, retries=2):
    """
    Drop-in replacement for ds.safe_get that tries:
      1. curl_cffi with rotating browser profiles (no proxy)
      2. curl_cffi with proxy rotation
      3. Original safe_get as final fallback
    """
    if not url or str(url).strip() in ("", "N/A", "None"):
        return None
    if not url.startswith("http"):
        url = "https://" + url.lstrip("/")

    # ── Strategy 1: curl_cffi, no proxy ──────────────────────────────────────
    # Do NOT pass ds.HEADERS — curl_cffi sets headers that match the impersonated
    # browser profile; a mismatched User-Agent breaks the TLS fingerprint and gets blocked.
    if _CFFI_OK:
        profiles = random.sample(_CFFI_PROFILES, len(_CFFI_PROFILES))
        for profile in profiles:
            try:
                sess = cffi_requests.Session(impersonate=profile)
                r = sess.get(url, timeout=15, verify=False, allow_redirects=True)
                if r.status_code == 200 and not _is_challenge_html(r):
                    ds.log.info(f"   [bypass] curl_cffi/{profile} OK: {url}")
                    return r
                # 202 = Sucuri JS challenge; 403 = IP block — move to proxy
                break
            except Exception:
                continue

    # ── Strategy 2: curl_cffi + proxy ────────────────────────────────────────
    if _CFFI_OK and _proxy_pool:
        available = [p for p in _proxy_pool if p not in _proxy_fail]
        random.shuffle(available)
        for proxy in available[:6]:
            profile = random.choice(_CFFI_PROFILES[:3])
            try:
                sess = cffi_requests.Session(impersonate=profile)
                r = sess.get(url, timeout=25, verify=False, allow_redirects=True,
                             proxies={"http": proxy, "https": proxy})
                if r.status_code == 200 and not _is_challenge_html(r):
                    ds.log.info(f"   [bypass] curl_cffi/{profile}+proxy OK: {url}")
                    return r
                if r.status_code in (407, 403):
                    _proxy_fail.add(proxy)
            except Exception:
                _proxy_fail.add(proxy)
                continue

    # ── Strategy 3: original safe_get ────────────────────────────────────────
    return _orig_safe_get(url, retries=retries)


# ── Playwright with stealth + optional proxy ──────────────────────────────────

def _pw_proxy_dict(proxy_url):
    """Convert http://user:pass@host:port to Playwright proxy dict."""
    p = urlparse(proxy_url)
    d = {"server": f"{p.scheme}://{p.hostname}:{p.port}"}
    if p.username:
        d["username"] = p.username
    if p.password:
        d["password"] = p.password
    return d


def _make_pw_context(pw, proxy_url=None):
    kwargs = dict(
        user_data_dir="",
        headless=True,
        ignore_https_errors=True,
        args=[
            "--disable-blink-features=AutomationControlled",
            "--disable-infobars",
            "--window-size=1920,1080",
            "--no-first-run",
        ],
        user_agent=ds.HEADERS["User-Agent"],
    )
    if proxy_url:
        kwargs["proxy"] = _pw_proxy_dict(proxy_url)
    ctx = pw.chromium.launch_persistent_context(**kwargs)
    ctx.add_init_script(script=_STEALTH_JS)
    return ctx


def _pw_get_html(url, proxy_url=None):
    """Fetch a URL using Playwright with stealth. Returns HTML string or None."""
    if not _PW_OK:
        return None
    try:
        with sync_playwright() as pw:
            ctx  = _make_pw_context(pw, proxy_url)
            page = ctx.new_page()
            page.set_extra_http_headers({"Accept-Language": "en-US,en;q=0.9"})
            page.goto(url, timeout=25000, wait_until="domcontentloaded")
            page.wait_for_timeout(5000)
            # Wait for Cloudflare JS challenge to auto-resolve
            for _ in range(3):
                title = page.title().lower()
                if "just a moment" in title or "checking your" in title:
                    page.wait_for_timeout(5000)
                else:
                    break
            html = page.content()
            ctx.close()
            return html if len(html) > 2000 else None
    except Exception as e:
        ds.log.warning(f"   [bypass] Playwright failed for {url}: {e}")
        return None


# ── Detect bot-blocked rows in a batch Excel output ───────────────────────────

# Batch output column positions (1-based), matching write_output col_headers
_B_INDEX    = 1
_B_NAME     = 2
_B_DOCTOR   = 3
_B_CITY     = 5
_B_STATE    = 6
_B_ZIP      = 7
_B_WEBSITE  = 8
_B_EMAIL    = 9


def _is_blocked_row(ws, r):
    """Return True if this row likely failed due to bot protection."""
    web = str(ws.cell(r, _B_WEBSITE).value or "").strip()
    if not web or web.lower() in ("not found", "n/a", "none", ""):
        return False   # no website — not a bot-block issue
    doc   = str(ws.cell(r, _B_DOCTOR).value  or "").strip().lower()
    email = str(ws.cell(r, _B_EMAIL).value   or "").strip().lower()
    # If doctor name AND email are both missing/not-found, likely bot-blocked
    doc_missing   = doc   in ("not found", "", "none", "n/a")
    email_missing = email in ("not found", "", "none", "n/a")
    return doc_missing and email_missing


def load_blocked_from_batch(path):
    """
    Read a batch output Excel, return list of row dicts for bot-blocked practices.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    for r in range(3, ws.max_row + 1):   # row 1 = group header, row 2 = col headers
        if not _is_blocked_row(ws, r):
            continue
        rows.append({
            "Index":         ws.cell(r, _B_INDEX).value,
            "Practice Name": ws.cell(r, _B_NAME).value  or "",
            "Website":       ws.cell(r, _B_WEBSITE).value or "",
            "City":          ws.cell(r, _B_CITY).value   or "",
            "State":         ws.cell(r, _B_STATE).value  or "",
            "Zip":           ws.cell(r, _B_ZIP).value    or "",
        })
    wb.close()
    return rows


# ── Scrape one practice with all bypass strategies ────────────────────────────

def scrape_with_bypass(row, pw_page=None):
    """
    Monkey-patch ds.safe_get, run the full scraper pipeline, restore original.
    Playwright page uses stealth JS already applied at context level.
    """
    global _orig_safe_get
    _orig_safe_get = ds.safe_get
    ds.safe_get = _bypass_safe_get
    try:
        result = ds.scrape_practice(row, pw_page=pw_page)
    finally:
        ds.safe_get = _orig_safe_get
    return result


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(
        description="Re-scrape bot-blocked dental practice websites.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    grp = p.add_mutually_exclusive_group(required=True)
    grp.add_argument("--urls",  nargs="+", metavar="URL",
                     help="One or more website URLs to scrape")
    grp.add_argument("--batch", metavar="BATCH.xlsx",
                     help="Batch output Excel — auto-detects bot-blocked rows")
    p.add_argument("--proxies", metavar="proxies.txt",
                   help="Proxy list file (one proxy per line)")
    p.add_argument("--output", "-o", default="bypass_output.xlsx",
                   help="Output Excel file (default: bypass_output.xlsx)")
    args = p.parse_args()

    load_proxies(args.proxies)

    # ── Build practice rows ───────────────────────────────────────────────────
    if args.urls:
        rows = []
        for i, url in enumerate(args.urls, 1):
            if not url.startswith("http"):
                url = "https://" + url
            rows.append({
                "Index": i, "Practice Name": url,
                "Website": url, "City": "", "State": "", "Zip": "",
            })
    else:
        rows = load_blocked_from_batch(args.batch)
        if not rows:
            print("No bot-blocked rows found in the batch file.")
            print("Criteria: website present but Doctor Name + Email both missing.")
            sys.exit(0)
        print(f"Found {len(rows)} likely bot-blocked rows in {args.batch}")

    print(f"\nScraping {len(rows)} practice(s) with bypass strategies…")
    if _proxy_pool:
        print(f"Proxy pool: {len(_proxy_pool)} proxies")
    print()

    # ── Set up Playwright (one browser for all practices) ─────────────────────
    pw_ctx  = None
    pw_page = None
    _pw_mgr = None

    if _PW_OK:
        proxy_url = _next_proxy()
        _pw_mgr  = sync_playwright().__enter__()
        pw_ctx   = _make_pw_context(_pw_mgr, proxy_url)
        pw_page  = pw_ctx.new_page()
        pw_page.set_extra_http_headers({"Accept-Language": "en-US,en;q=0.9"})
        print(f"Playwright ready{' (+proxy)' if proxy_url else ''}\n")

    all_results = []
    try:
        for i, row in enumerate(rows, 1):
            name = row.get("Practice Name") or row.get("Website")
            print(f"[{i}/{len(rows)}] {name}")
            result = scrape_with_bypass(row, pw_page=pw_page)
            all_results.append((row, result))
            # Save checkpoint every 5 practices
            if i % 5 == 0 or i == len(rows):
                ds.write_output(all_results, args.output)
                print(f"  ✓ Checkpoint saved → {args.output}")
            time.sleep(1)
    finally:
        if pw_ctx:  pw_ctx.close()
        if _pw_mgr:
            try: _pw_mgr.__exit__(None, None, None)
            except Exception: pass

    ds.write_output(all_results, args.output)
    print(f"\nDone. {len(all_results)} practices saved → {args.output}")


if __name__ == "__main__":
    main()
