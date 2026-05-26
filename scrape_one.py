#!/usr/bin/env python3
"""
scrape_one.py — Scrape a single practice website and save to Excel.

Usage:
    python scrape_one.py "https://dentalkidds.com/about-us/"
    python scrape_one.py "https://dentalkidds.com/about-us/" --name "Dental Kidds" --output result.xlsx
"""
import argparse
import sys
from urllib.parse import urlparse

from dental_scraper import (
    scrape_practice, write_output,
    USE_PLAYWRIGHT, PLAYWRIGHT_AVAILABLE,
    HEADERS,
)

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

try:
    from playwright.sync_api import sync_playwright
    _PW_AVAIL = True
except ImportError:
    _PW_AVAIL = False


def _homepage(url):
    """Strip path/query so scraper always starts from the root domain."""
    p = urlparse(url)
    return f"{p.scheme}://{p.netloc}/"


def main():
    p = argparse.ArgumentParser()
    p.add_argument("url",            help="Practice website URL (any page — homepage is used automatically)")
    p.add_argument("--name",  "-n",  default="", help="Practice name (optional)")
    p.add_argument("--output", "-o", default="scrape_one_output.xlsx")
    args = p.parse_args()

    home = _homepage(args.url)
    # If user gave a sub-page URL (e.g. /about-us/), pass it as a hint so the
    # scraper treats it as the team/doctor page even on JS-rendered sites.
    extra_team = args.url if args.url.rstrip("/") != home.rstrip("/") else ""
    row = {
        "Index":         1,
        "Practice Name": args.name or home,
        "Website":       home,
        "City":          "",
        "State":         "",
        "Zip":           "",
        "_hint_team_url": extra_team,   # consumed by scrape_practice if present
    }
    if extra_team:
        print(f"Using homepage : {home}")
        print(f"Team page hint : {extra_team}")
    else:
        print(f"Using homepage: {home}")

    if USE_PLAYWRIGHT and PLAYWRIGHT_AVAILABLE and _PW_AVAIL:
        print("Launching Playwright browser…")
        with sync_playwright() as pw:
            pw_ctx  = pw.chromium.launch_persistent_context(
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
                user_agent=HEADERS["User-Agent"],
            )
            pw_ctx.add_init_script(script=_STEALTH_JS)
            pw_page = pw_ctx.new_page()
            pw_page.set_extra_http_headers({"Accept-Language": "en-US,en;q=0.9"})
            print(f"Scraping: {home}")
            result = scrape_practice(row, pw_page=pw_page)
            pw_ctx.close()
    else:
        print(f"Scraping: {home} (requests only — install playwright for JS sites)")
        result = scrape_practice(row, pw_page=None)

    write_output([(row, result)], args.output)
    print(f"\nSaved → {args.output}")

    print("\n── Key Results ──────────────────────────────────")
    for field in ["scraped_doctor_names", "hygienists", "email",
                  "associations", "specialty",
                  "cerec", "cbct", "lasers", "ai", "intraoral",
                  "invisalign", "implants", "veneers", "sedation",
                  "facebook_url", "instagram_url", "google_rating",
                  "google_reviews"]:
        val = result.get(field, "")
        if val and str(val).strip() not in ("Not Found", "N/A", "0", ""):
            print(f"  {field:<26} {val}")


if __name__ == "__main__":
    main()
