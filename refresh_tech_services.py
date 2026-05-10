"""
refresh_tech_services.py
========================
Re-extracts Technology in Practice, Services (# of Mentions), and Testimonials
from the page_cache/ built by dental_scraper.py, then patches a batch xlsx file.

Rules when writing back:
  • Tech fields  (CEREC/CBCT/Laser/AI/Intraoral) : keep "X" if EITHER old or new is "X"
  • Service counts (numeric)                      : keep MAX(old, new)
  • Dental Plan                                   : keep "Mentioned" if EITHER has it
  • Testimonials                                  : keep MAX(old, new)

Nothing is ever reduced. Only improves.

Usage:
    python3 refresh_tech_services.py <batch_deduped.xlsx>
    python3 refresh_tech_services.py <batch_deduped.xlsx> --cache-dir /path/to/page_cache

Outputs:
    <input>_refreshed.xlsx   — patched xlsx
    <input>_comparison.xlsx  — before/after report (changed cells highlighted)
"""

import os, sys, re, shutil, glob, logging
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dental_scraper as ds

logging.basicConfig(level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

# ── Column positions (1-based), matching dental_scraper.py ───────────────────
C_IDX   = 1;  C_NAME  = 2
C_CEREC = 23; C_CBCT  = 24; C_LASER = 25; C_AI   = 26; C_INTRA = 27
C_INV   = 28  # col 29 = InvTier (skip)
C_CLEAR = 30; C_VEN   = 31; C_IMPL  = 32; C_SMILE = 33; C_WHITE = 34
C_SED   = 35; C_HOL   = 36; C_PLAN  = 37; C_CANC  = 38
C_TESTI = 46

TECH_COLS = {C_CEREC, C_CBCT, C_LASER, C_AI, C_INTRA}

# (xlsx_col, result_key, display_label)
FIELD_MAP = [
    (C_CEREC, "cerec",           "CEREC"),
    (C_CBCT,  "cbct",            "CBCT"),
    (C_LASER, "lasers",          "Lasers"),
    (C_AI,    "ai",              "AI"),
    (C_INTRA, "intraoral",       "Intraoral"),
    (C_INV,   "invisalign",      "Invisalign"),
    (C_CLEAR, "clear_aligners",  "Clear Aligners"),
    (C_VEN,   "veneers",         "Veneers"),
    (C_IMPL,  "implants",        "Implants"),
    (C_SMILE, "smile_makeovers", "Smile Makeovers"),
    (C_WHITE, "whitening",       "Whitening"),
    (C_SED,   "sedation",        "Sedation"),
    (C_HOL,   "holistic",        "Holistic"),
    (C_PLAN,  "dental_plan",     "Dental Plan"),
    (C_CANC,  "cancer_screening","Cancer Screening"),
    (C_TESTI, "testimonials",    "Testimonials"),
]

DATA_START = 3

_TEST_RE = re.compile(
    r"(testimonial|review|quote|patient.story|patient.review|"
    r"feedback|client.say|what.people|slider.item|carousel.item|"
    r"swiper.slide|slick.slide|rating.block|star.review)", re.I
)


# ── Cache loader ──────────────────────────────────────────────────────────────

def _find_cache_folder(idx: int, cache_dir: str):
    prefix = f"{idx:03d}_"
    try:
        for name in sorted(os.listdir(cache_dir)):
            if name.startswith(prefix) and os.path.isdir(os.path.join(cache_dir, name)):
                return os.path.join(cache_dir, name)
    except FileNotFoundError:
        pass
    return None


def _load_pages(folder: str) -> list:
    """
    Load all cached HTML pages for a practice.
    Reads every .html file in the folder (not just manifest entries) to ensure
    nothing is missed.
    Returns list of (page_type, url, html_text).
    """
    import json

    # Build url map from manifest
    url_map = {}
    manifest_path = os.path.join(folder, "manifest.json")
    if os.path.exists(manifest_path):
        try:
            m = json.load(open(manifest_path, encoding="utf-8"))
            for ptype, info in m.get("pages", {}).items():
                url_map[info.get("file", "")] = (ptype, info.get("url", ""))
        except Exception:
            pass

    pages = []
    seen_files = set()
    for fpath in sorted(glob.glob(os.path.join(folder, "*.html"))):
        fname = os.path.basename(fpath)
        if fname in seen_files:
            continue
        seen_files.add(fname)
        try:
            html = open(fpath, encoding="utf-8", errors="replace").read()
        except Exception:
            continue
        if len(html) < 200:
            continue
        ptype, url = url_map.get(fname, (fname.replace(".html", ""), ""))
        pages.append((ptype, url, html))

    return pages


# ── Extraction ────────────────────────────────────────────────────────────────

def _extract(pages: list) -> dict:
    """
    Extract tech / services / testimonials from cached pages.
    - Body text (nav stripped): cap=5 per page
    - Full text + meta/JSON-LD/URL signals: cap=3 per page
    - Takes MAX of body and full counts so nothing is missed
    """
    all_text = ""
    per_page = []   # (body_text, full_text_augmented)
    all_soups = []

    for _ptype, url, html in pages:
        ft  = ds.extract_text(html)
        bt  = ds.extract_body_text(html)
        aug = ds.extract_augmented_text(html, url)
        ft_aug = ft + " " + aug
        all_text += " " + ft_aug
        per_page.append((bt, ft_aug))
        all_soups.append(BeautifulSoup(html, "lxml"))

    if not all_text.strip():
        return {}

    # ── Technology ────────────────────────────────────────────────────────
    tf = set()
    for kw, tn in ds.TECH_KEYWORDS.items():
        if kw in all_text:
            tf.add(tn)
    # AI needs word-boundary to avoid false matches (e.g. "said", "await")
    if "AI" not in tf and re.search(r"\bai\b", all_text, re.I):
        tf.add("AI")

    # ── Services (per-page, body primary / full+augmented fallback) ───────
    svc_b = dict.fromkeys(set(ds.SERVICE_KEYWORDS.values()), 0)
    svc_f = dict.fromkeys(set(ds.SERVICE_KEYWORDS.values()), 0)
    seen_urls: set = set()
    for i, (bt, ft_aug) in enumerate(per_page):
        url_key = pages[i][1] or str(i)
        if url_key in seen_urls:
            continue
        seen_urls.add(url_key)
        for kw, cat in ds.SERVICE_KEYWORDS.items():
            svc_b[cat] += ds.count_keyword_capped(bt, kw, cap=5)
            svc_f[cat] += ds.count_keyword_capped(ft_aug, kw, cap=3)
    # Take MAX of body and full — ensures meta/JSON-LD/URL mentions aren't lost
    svc = {cat: max(svc_b[cat], svc_f[cat]) for cat in svc_b}

    # ── Testimonials ──────────────────────────────────────────────────────
    seen_t, tt = set(), 0
    for soup in all_soups:
        for blk in soup.find_all(
            ["div", "section", "article", "blockquote", "li"], class_=_TEST_RE
        ):
            k = blk.get_text(separator=" ", strip=True)[:80]
            if k and k not in seen_t:
                seen_t.add(k); tt += 1
        for blk in soup.find_all(lambda t: any(
            _TEST_RE.search(str(v))
            for k, v in t.attrs.items()
            if k.startswith("data-") and isinstance(v, str)
        )):
            k = blk.get_text(separator=" ", strip=True)[:80]
            if k and k not in seen_t:
                seen_t.add(k); tt += 1
    if tt == 0:
        for soup in all_soups:
            for bq in soup.find_all("blockquote"):
                k = bq.get_text(separator=" ", strip=True)[:80]
                if k and k not in seen_t:
                    seen_t.add(k); tt += 1

    return {
        C_CEREC: "X" if "CEREC"              in tf else "",
        C_CBCT:  "X" if "CBCT"               in tf else "",
        C_LASER: "X" if "Lasers"             in tf else "",
        C_AI:    "X" if "AI"                 in tf else "",
        C_INTRA: "X" if "Intraoral Scanners" in tf else "",
        C_INV:   svc.get("Invisalign",        0),
        C_CLEAR: svc.get("Clear Aligners",    0),
        C_VEN:   svc.get("Veneers",           0),
        C_IMPL:  svc.get("Implants",          0),
        C_SMILE: svc.get("Smile Makeovers",   0),
        C_WHITE: svc.get("Teeth Whitening",   0),
        C_SED:   svc.get("Sedation Dentistry",0),
        C_HOL:   svc.get("Holistic Dentistry",0),
        C_PLAN:  "Mentioned" if svc.get("Dental Plan", 0) > 0 else "",
        C_CANC:  svc.get("Cancer Screening",  0),
        C_TESTI: str(tt),
    }


# ── Merge logic — never reduce ────────────────────────────────────────────────

def _merge(col: int, old, new):
    """
    Combine old (xlsx) value with new (cache-extracted) value.
    Never reduces: returns the better of the two.
    """
    if col in TECH_COLS:
        # Keep "X" if either side has it
        return "X" if (str(old or "").strip() == "X" or str(new or "").strip() == "X") else ""

    if col == C_PLAN:
        return "Mentioned" if (
            str(old or "").strip() == "Mentioned" or
            str(new or "").strip() == "Mentioned"
        ) else ""

    # Numeric fields — take max
    try:
        old_n = int(str(old or 0).replace(",", ""))
    except (ValueError, TypeError):
        old_n = 0
    try:
        new_n = int(str(new or 0).replace(",", ""))
    except (ValueError, TypeError):
        new_n = 0
    return max(old_n, new_n)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _str(v) -> str:
    if v is None or str(v).strip() in ("", "None", "Not Found", "ERROR"):
        return ""
    return str(v).strip()


# ── Main ──────────────────────────────────────────────────────────────────────

def refresh(input_xlsx: str, cache_dir: str = "page_cache"):
    base, ext = os.path.splitext(input_xlsx)
    out_path  = base + "_refreshed" + ext
    comp_path = base + "_comparison.xlsx"

    shutil.copy2(input_xlsx, out_path)
    log.info("Input  : %s", input_xlsx)
    log.info("Output : %s", out_path)
    log.info("Cache  : %s", cache_dir)

    # ── Read xlsx ──────────────────────────────────────────────────────────
    wb = openpyxl.load_workbook(out_path, data_only=True)
    ws = wb.active

    idx_rows   = {}   # idx -> [row_num, ...]
    idx_before = {}   # idx -> {col: old_value}
    idx_name   = {}   # idx -> practice name

    for row in ws.iter_rows(min_row=DATA_START, values_only=False):
        raw = row[C_IDX - 1].value
        if raw is None:
            continue
        try:
            idx = int(raw)
        except (TypeError, ValueError):
            continue
        rn = row[0].row
        idx_rows.setdefault(idx, []).append(rn)
        if idx not in idx_before:
            idx_before[idx] = {col: ws.cell(rn, col).value for col, _, _ in FIELD_MAP}
            idx_name[idx]   = str(row[C_NAME - 1].value or "")

    wb.close()
    log.info("Practices in xlsx: %d", len(idx_rows))

    # ── Extract from cache ─────────────────────────────────────────────────
    merged_vals = {}   # idx -> {col: merged_value}
    n_found = n_missing = 0

    for idx in sorted(idx_rows):
        folder = _find_cache_folder(idx, cache_dir)
        if not folder:
            n_missing += 1
            continue

        pages = _load_pages(folder)
        if not pages:
            n_missing += 1
            continue

        extracted = _extract(pages)
        if not extracted:
            n_missing += 1
            continue

        n_found += 1
        old = idx_before[idx]

        # Merge: take best of old and new
        merged = {}
        for col, _, _ in FIELD_MAP:
            merged[col] = _merge(col, old.get(col), extracted.get(col))
        merged_vals[idx] = merged

        log.info("  [%03d] %-30s  CEREC=%-2s  Inv=%-3s  Impl=%-3s  Testi=%s  pages=%d",
                 idx, idx_name[idx][:30],
                 merged[C_CEREC] or "-",
                 merged[C_INV],
                 merged[C_IMPL],
                 merged[C_TESTI],
                 len(pages))

    log.info("Cache found: %d / %d  (no cache: %d)", n_found, len(idx_rows), n_missing)

    if not merged_vals:
        log.warning("No cache data — nothing to patch.")
        return out_path, None

    # ── Write merged values ────────────────────────────────────────────────
    wb2 = openpyxl.load_workbook(out_path, data_only=True)
    ws2 = wb2.active
    updates = 0

    for idx, vals in merged_vals.items():
        for rn in idx_rows.get(idx, []):
            for col, v in vals.items():
                ws2.cell(rn, col).value = v
                updates += 1

    wb2.save(out_path)
    log.info("Wrote %d cell updates → %s", updates, out_path)

    # ── Comparison report ──────────────────────────────────────────────────
    _write_comparison(comp_path, idx_rows, idx_name, idx_before, merged_vals)
    return out_path, comp_path


# ── Comparison report ─────────────────────────────────────────────────────────

def _write_comparison(path, idx_rows, idx_name, before, after):
    from openpyxl import Workbook

    wb  = Workbook()
    ws  = wb.active
    ws.title = "Tech & Services Comparison"

    YELLOW = PatternFill("solid", fgColor="FFFACD")
    GREEN  = PatternFill("solid", fgColor="C6EFCE")
    GREY   = PatternFill("solid", fgColor="E8E8E8")
    BLUE   = PatternFill("solid", fgColor="DDEEFF")
    HDR    = Font(bold=True, size=9)
    DATA   = Font(size=9)
    ctr    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lft    = Alignment(horizontal="left",   vertical="center")

    labels = [lbl for _, _, lbl in FIELD_MAP]
    n      = len(FIELD_MAP)

    # Row 1: group labels
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    ws.cell(1, 1, "Practice").font = HDR; ws.cell(1, 1).alignment = ctr

    ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=3 + n)
    c = ws.cell(1, 4, "BEFORE"); c.font = HDR; c.fill = GREY; c.alignment = ctr

    ws.merge_cells(start_row=1, start_column=4 + n, end_row=1, end_column=3 + 2 * n)
    c = ws.cell(1, 4 + n, "AFTER (refreshed)"); c.font = HDR; c.fill = BLUE; c.alignment = ctr

    # Row 2: column headers
    for c_idx, h in enumerate(["Index", "Practice Name", "# Improved"] + labels + labels, 1):
        cell = ws.cell(2, c_idx, h)
        cell.font = HDR; cell.alignment = ctr
        if 4 <= c_idx <= 3 + n:
            cell.fill = GREY
        elif c_idx >= 4 + n:
            cell.fill = BLUE

    r = 3
    n_improved_total = 0

    for idx in sorted(idx_rows):
        if idx not in after:
            continue

        b = before.get(idx, {})
        a = after[idx]

        bv = [_str(b.get(col, "")) for col, _, _ in FIELD_MAP]
        av = [_str(a.get(col, "")) for col, _, _ in FIELD_MAP]
        improved = [av[i] != bv[i] and av[i] not in ("", "0") for i in range(n)]
        n_imp = sum(improved)

        if n_imp > 0:
            n_improved_total += 1

        row_data = [idx, idx_name.get(idx, ""), n_imp] + bv + av
        for c_idx, v in enumerate(row_data, 1):
            cell = ws.cell(r, c_idx, v)
            cell.font = DATA
            cell.alignment = lft if c_idx == 2 else ctr
            fi = c_idx - 4
            if 4 <= c_idx <= 3 + n and improved[fi]:
                cell.fill = YELLOW
            elif c_idx >= 4 + n and improved[c_idx - 4 - n]:
                cell.fill = GREEN

        r += 1

    # Widths
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 10
    for ci in range(4, 4 + 2 * n):
        ws.column_dimensions[get_column_letter(ci)].width = 11

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 32
    ws.freeze_panes = "D3"

    log.info("Comparison: %d practices, %d with improvements → %s",
             r - 3, n_improved_total, path)
    wb.save(path)


# ── Entry point ───────────────────────────────────────────────────────────────

def main():
    args      = sys.argv[1:]
    cache_dir = "page_cache"
    xlsx_file = None
    i = 0
    while i < len(args):
        if args[i] == "--cache-dir" and i + 1 < len(args):
            cache_dir = args[i + 1]; i += 2
        elif xlsx_file is None:
            xlsx_file = args[i]; i += 1
        else:
            i += 1

    if not xlsx_file:
        print("Usage: python3 refresh_tech_services.py <batch.xlsx> [--cache-dir dir]")
        sys.exit(1)
    if not os.path.exists(xlsx_file):
        log.error("File not found: %s", xlsx_file)
        sys.exit(1)

    refresh(xlsx_file, cache_dir)


if __name__ == "__main__":
    main()
