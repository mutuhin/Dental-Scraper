"""
scrape_rows.py
──────────────
Scrape specific row numbers from the master xlsx without touching dental_scraper.py.

Usage:
    python scrape_rows.py 609 613 616 635 678
    python scrape_rows.py 609 613 --output my_output.xlsx
    python scrape_rows.py 609 --website "609=https://crestviewsmiles.com/html/index.html"

--website ROWNUM=URL  : override the website for a specific row
--proxy-rotate ROWNUM : enable multi-country proxy rotation for a specific row

Row numbers are 1-based (same as batch definitions: batch 7 = rows 601-700).
"""

import argparse
import logging
import os
import re
import sys
import time
from urllib.parse import urlparse

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

import dental_scraper as ds

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

# ── Proxy rotation helpers ────────────────────────────────────────────────────
# Oxylabs supports country-specific residential IPs via the username prefix.
# Each country code fetches from a different IP pool.
_ROTATION_COUNTRIES = ["US", "GB", "CA", "AU", "DE", "FR", "NL", "JP", "SG", "BR"]


def _rotated_proxy(base_proxy: str, attempt: int) -> str:
    """Return a proxy URL with a different country and fresh session per attempt."""
    if not base_proxy:
        return base_proxy
    p = urlparse(base_proxy)
    user = p.username or ""
    country = _ROTATION_COUNTRIES[attempt % len(_ROTATION_COUNTRIES)]
    # Replace -cc-XX country code (e.g. -cc-US → -cc-GB)
    new_user = re.sub(r"-cc-[A-Z]{2}", f"-cc-{country}", user)
    if new_user == user:
        # No country code pattern found — append one
        new_user = f"{user}-cc-{country}"
    # Remove any existing session lock so we always get a fresh IP
    new_user = re.sub(r"-sessid-[^:@]+", "", new_user)
    pwd  = p.password or ""
    host = p.hostname or ""
    port = p.port or 7777
    log.info(f"   Proxy rotation attempt {attempt+1}: country={country}  user={new_user[:40]}…")
    return f"http://{new_user}:{pwd}@{host}:{port}"


def _scrape_with_rotation(practice: dict, pw_page, max_attempts: int = 8) -> dict:
    """
    Scrape a practice with proxy rotation: try up to max_attempts times,
    each time using a different country's residential IP pool.
    Returns the best result found (first one that has more than just a skip_reason).
    """
    base_proxy = ds._BYPASS_PROXY
    if not base_proxy:
        log.warning("   No proxy configured — cannot rotate. Attempting without proxy.")
        return ds.scrape_practice(practice, pw_page=pw_page)

    best_result = None
    for attempt in range(max_attempts):
        rotated = _rotated_proxy(base_proxy, attempt)
        ds._BYPASS_PROXY = rotated  # swap in the rotated proxy for this attempt

        log.info(f"   Rotation attempt {attempt + 1}/{max_attempts}…")
        try:
            result = ds.scrape_practice(practice, pw_page=pw_page)
        except Exception as e:
            log.warning(f"   Attempt {attempt + 1} raised exception: {e}")
            result = dict(ds.EMPTY_SCRAPED)
        finally:
            ds._BYPASS_PROXY = base_proxy  # always restore original

        skip = result.get("skip_reason", "")
        doctors = result.get("scraped_doctor_names", "Not Found")
        # Consider success if we got past bot-protection
        if not skip or ("Bot Protection" not in skip and "403" not in skip):
            log.info(f"   Rotation attempt {attempt + 1} succeeded (skip_reason={skip!r})")
            return result

        log.warning(f"   Attempt {attempt + 1} still blocked ({skip}) — rotating IP…")
        best_result = result
        time.sleep(3)  # brief pause before next attempt

    log.warning("   All rotation attempts blocked. Returning best partial result.")
    return best_result or dict(ds.EMPTY_SCRAPED)


def main():
    parser = argparse.ArgumentParser(
        description="Scrape specific row numbers from the master xlsx."
    )
    parser.add_argument(
        "rows", nargs="+", type=int,
        help="1-based row numbers to scrape (e.g. 609 613 616 635 678)",
    )
    parser.add_argument("--input", "-i", default=None,
                        help="Path to input xlsx (default: 6000 Data COMPLETE.xlsx)")
    parser.add_argument("--output", "-o", default=None,
                        help="Output xlsx path (default: rows_<N1>_<N2>_....xlsx)")
    parser.add_argument(
        "--website", "-w", action="append", default=[], metavar="ROW=URL",
        help="Override website for a row (e.g. --website 609=https://crestviewsmiles.com)",
    )
    parser.add_argument(
        "--proxy-rotate", action="append", default=[], type=int, metavar="ROW",
        help="Enable proxy rotation for this row number (e.g. --proxy-rotate 613)",
    )
    args = parser.parse_args()

    # ── Parse website overrides: "609=https://..." → {609: "https://..."} ─────
    website_overrides: dict = {}
    for spec in args.website:
        if "=" not in spec:
            log.error(f"--website must be ROW=URL, got: {spec!r}")
            sys.exit(1)
        row_str, url = spec.split("=", 1)
        website_overrides[int(row_str.strip())] = url.strip()

    rotate_rows = set(args.proxy_rotate)

    # ── Resolve input file ────────────────────────────────────────────────────
    if args.input:
        input_file = args.input
    else:
        here = os.path.dirname(os.path.abspath(__file__))
        input_file = os.path.join(here, "6000 Data COMPLETE.xlsx")
    if not os.path.exists(input_file):
        log.error(f"Input file not found: {input_file}")
        sys.exit(1)

    # ── Resolve output file ───────────────────────────────────────────────────
    if args.output:
        output_file = args.output
    else:
        tag = "_".join(str(r) for r in sorted(args.rows))
        output_file = f"rows_{tag}.xlsx"

    # ── Load and filter practices ─────────────────────────────────────────────
    log.info(f"Reading input: {input_file}")
    all_practices = ds.read_practices(input_file, start_idx=0, end_idx=None)
    total = len(all_practices)
    log.info(f"Total practices in file: {total}")

    target_rows = sorted(set(args.rows))
    selected = []
    for r in target_rows:
        idx = r - 1
        if idx < 0 or idx >= total:
            log.warning(f"Row {r} out of range (file has {total} practices) — skipping")
            continue
        practice = dict(all_practices[idx])  # copy so we can override website safely
        if r in website_overrides:
            old = practice.get("Website", "")
            practice["Website"] = website_overrides[r]
            log.info(f"Row {r}: website overridden  {old}  →  {website_overrides[r]}")
        selected.append((r, practice))

    if not selected:
        log.error("No valid rows to scrape.")
        sys.exit(1)

    log.info(f"Scraping {len(selected)} practices: rows {[r for r, _ in selected]}")
    if rotate_rows:
        log.info(f"Proxy rotation enabled for rows: {sorted(rotate_rows)}")

    # ── Proxy status ──────────────────────────────────────────────────────────
    proxy_url = ds._BYPASS_PROXY
    if proxy_url:
        log.info(f"Proxy available: {proxy_url.split('@')[-1]}")
    else:
        log.info("No proxy configured (set OXYLABS_USER/OXYLABS_PASS or proxies.txt)")

    # ── Launch Playwright ─────────────────────────────────────────────────────
    pw_context = None
    pw_page    = None

    if ds.USE_PLAYWRIGHT and ds.PLAYWRIGHT_AVAILABLE:
        try:
            from playwright.sync_api import sync_playwright
            _pw = sync_playwright().__enter__()
            pw_context = _pw.chromium.launch_persistent_context(
                user_data_dir="",
                headless=True,
                ignore_https_errors=True,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--disable-infobars",
                    "--window-size=1920,1080",
                    "--no-first-run",
                    "--disable-extensions",
                ],
                user_agent=ds.HEADERS["User-Agent"],
            )
            pw_context.add_init_script(script=_STEALTH_JS)
            pw_page = pw_context.new_page()
            pw_page.set_extra_http_headers({"Accept-Language": "en-US,en;q=0.9"})
            log.info("Playwright browser launched.")
        except Exception as e:
            log.warning(f"Playwright launch failed ({e}) — falling back to requests only")
            pw_context = pw_page = None

    # ── Scrape each practice ──────────────────────────────────────────────────
    all_results = []
    for row_num, practice in selected:
        site = practice.get("Website") or ""
        name = practice.get("Practice Name") or ""
        log.info(f"── Row {row_num}: {name}  ({site})")

        if not site or str(site).strip().lower() in ("", "not found", "n/a"):
            log.warning(f"   Row {row_num}: no website — skipping")
            all_results.append((practice, dict(ds.EMPTY_SCRAPED)))
            continue

        try:
            if row_num in rotate_rows:
                log.info(f"   Using proxy rotation for row {row_num}…")
                result = _scrape_with_rotation(practice, pw_page=pw_page)
            else:
                result = ds.scrape_practice(practice, pw_page=pw_page)
        except Exception as e:
            log.error(f"   Row {row_num} error: {e}", exc_info=True)
            result = dict(ds.EMPTY_SCRAPED)

        # Reopen Playwright page if it was closed/crashed
        if pw_context and pw_page and pw_page.is_closed():
            try:
                pw_page = pw_context.new_page()
                pw_page.set_extra_http_headers({"Accept-Language": "en-US,en;q=0.9"})
            except Exception:
                pw_page = None

        all_results.append((practice, result))
        time.sleep(1)

    # ── Clean up Playwright ───────────────────────────────────────────────────
    if pw_context:
        try:
            pw_context.close()
        except Exception:
            pass

    # ── Write output ──────────────────────────────────────────────────────────
    ds.write_output(all_results, output_file)
    log.info(f"Saved → {output_file}")

    # ── Summary ───────────────────────────────────────────────────────────────
    print("\n── Results Summary ──────────────────────────────────────────")
    for (practice, result), (row_num, _) in zip(all_results, selected):
        pname = practice.get("Practice Name", "")
        docs  = result.get("scraped_doctor_names", "Not Found")
        spec  = result.get("specialty", "")
        skip  = result.get("skip_reason", "")
        cerec = result.get("cerec", "")
        invis = result.get("invisalign", 0)
        impl  = result.get("implants", 0)
        status = f"BLOCKED: {skip}" if skip and "Bot" in skip else "OK"
        print(f"  Row {row_num:>4}  {pname[:34]:<35}  [{status}]")
        print(f"         doctors: {str(docs)[:70]}")
        if spec:
            print(f"         specialty: {spec[:70]}")
        tech_parts = []
        if cerec: tech_parts.append(f"CEREC={cerec}")
        if invis: tech_parts.append(f"Invisalign={invis}")
        if impl:  tech_parts.append(f"Implants={impl}")
        if tech_parts:
            print(f"         tech/svcs: {', '.join(tech_parts)}")
        print()


if __name__ == "__main__":
    main()
