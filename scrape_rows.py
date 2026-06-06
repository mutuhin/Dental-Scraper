"""
scrape_rows.py
──────────────
Scrape specific row numbers from the master xlsx without touching dental_scraper.py.

Usage:
    python scrape_rows.py 609 613 616 635 678
    python scrape_rows.py 609 613 --output my_output.xlsx
    python scrape_rows.py 609 613 --input "6000 Data COMPLETE.xlsx"

Row numbers are 1-based (same numbering as the spreadsheet / batch definitions).
Batch 7 = rows 601-700, so row 609 is the 9th practice in batch 7.

Proxy:
    Set OXYLABS_USER / OXYLABS_PASS env vars, or place credentials in proxies.txt.
    The proxy is loaded automatically by dental_scraper._BYPASS_PROXY.

Output:
    rows_<N1>_<N2>_...xlsx  (or whatever --output specifies)
"""

import argparse
import logging
import os
import sys
import time

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── Import dental_scraper (read-only — we never modify it) ────────────────────
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


def main():
    parser = argparse.ArgumentParser(description="Scrape specific row numbers from the master xlsx.")
    parser.add_argument("rows", nargs="+", type=int,
                        help="1-based row numbers to scrape (e.g. 609 613 616 635 678)")
    parser.add_argument("--input", "-i", default=None,
                        help="Path to input xlsx (default: auto-detect alongside this script)")
    parser.add_argument("--output", "-o", default=None,
                        help="Output xlsx path (default: rows_<N1>_<N2>_....xlsx)")
    args = parser.parse_args()

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

    # Row numbers are 1-based; list index is row_number - 1
    target_rows = sorted(set(args.rows))
    selected = []
    for r in target_rows:
        idx = r - 1
        if idx < 0 or idx >= total:
            log.warning(f"Row {r} out of range (file has {total} practices) — skipping")
            continue
        selected.append((r, all_practices[idx]))

    if not selected:
        log.error("No valid rows to scrape.")
        sys.exit(1)

    log.info(f"Scraping {len(selected)} practices: rows {[r for r, _ in selected]}")

    # ── Proxy status ──────────────────────────────────────────────────────────
    proxy_url = ds._BYPASS_PROXY
    if proxy_url:
        masked = proxy_url.split("@")[-1]  # show only host:port
        log.info(f"Proxy available: {masked}")
    else:
        log.info("No proxy configured (set OXYLABS_USER/OXYLABS_PASS or proxies.txt)")

    # ── Launch Playwright (reused across all practices) ───────────────────────
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
            result = ds.scrape_practice(practice, pw_page=pw_page)
        except Exception as e:
            log.error(f"   Row {row_num} error: {e}")
            result = dict(ds.EMPTY_SCRAPED)
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
        name = practice.get("Practice Name", "")
        docs = result.get("scraped_doctor_names", "Not Found")
        spec = result.get("specialty", "")
        print(f"  Row {row_num:>4}  {name[:35]:<36}  doctors: {docs[:60]}")
        if spec:
            print(f"          {'':>4}  {'':36}  specialty: {spec[:60]}")


if __name__ == "__main__":
    main()
