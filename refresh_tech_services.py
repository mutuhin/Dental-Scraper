"""
refresh_tech_services.py
========================
Re-extracts Technology in Practice, Services (# of Mentions), and Testimonials
from the page_cache/ built by dental_scraper.py, then patches a batch xlsx file.

Always overwrites the 16 target fields — does not skip non-blank values.

Usage:
    python3 refresh_tech_services.py <batch_deduped.xlsx>
    python3 refresh_tech_services.py <batch_deduped.xlsx> --cache-dir /path/to/page_cache

Outputs:
    <input>_refreshed.xlsx   — patched xlsx
    <input>_comparison.xlsx  — before/after report (changed cells highlighted)
"""

import os, sys, shutil, logging
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import dental_scraper as ds
import reprocess as rp

logging.basicConfig(level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger(__name__)

# ── Column positions (1-based), matching dental_scraper.py write_output ──────
C_IDX   = 1;  C_NAME  = 2
C_CEREC = 23; C_CBCT  = 24; C_LASER = 25; C_AI   = 26; C_INTRA = 27
C_INV   = 28  # col 29 = InvTier (skip)
C_CLEAR = 30; C_VEN   = 31; C_IMPL  = 32; C_SMILE = 33; C_WHITE = 34
C_SED   = 35; C_HOL   = 36; C_PLAN  = 37; C_CANC  = 38
C_TESTI = 46

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


# ── Helpers ──────────────────────────────────────────────────────────────────

def _find_cache_folder(idx: int, cache_dir: str):
    prefix = f"{idx:03d}_"
    try:
        for name in sorted(os.listdir(cache_dir)):
            if name.startswith(prefix) and os.path.isdir(os.path.join(cache_dir, name)):
                return os.path.join(cache_dir, name)
    except FileNotFoundError:
        pass
    return None


def _str(v) -> str:
    if v is None or str(v).strip() in ("", "None", "Not Found", "ERROR"):
        return ""
    return str(v).strip()


# ── Main logic ────────────────────────────────────────────────────────────────

def refresh(input_xlsx: str, cache_dir: str = "page_cache"):
    base, ext = os.path.splitext(input_xlsx)
    out_path  = base + "_refreshed" + ext
    comp_path = base + "_comparison.xlsx"

    shutil.copy2(input_xlsx, out_path)
    log.info("Input  : %s", input_xlsx)
    log.info("Output : %s", out_path)
    log.info("Cache  : %s", cache_dir)

    # ── Read xlsx structure ────────────────────────────────────────────────
    wb = openpyxl.load_workbook(out_path, data_only=True)
    ws = wb.active

    idx_rows   = {}   # idx -> [row_num, ...]
    idx_before = {}   # idx -> {col: old_value}
    idx_name   = {}   # idx -> practice name string

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

    # ── Re-extract from cache ──────────────────────────────────────────────
    new_vals = {}   # idx -> {col: new_value}
    n_found = n_missing = 0

    for idx in sorted(idx_rows):
        folder = _find_cache_folder(idx, cache_dir)
        if not folder:
            n_missing += 1
            continue

        manifest, pages, cached_result = rp.load_cache(folder)
        if not pages:
            n_missing += 1
            continue

        result  = rp.reextract(pages, cached_result)
        n_found += 1

        new_vals[idx] = {
            C_CEREC: result.get("cerec",           ""),
            C_CBCT:  result.get("cbct",            ""),
            C_LASER: result.get("lasers",          ""),
            C_AI:    result.get("ai",              ""),
            C_INTRA: result.get("intraoral",       ""),
            C_INV:   result.get("invisalign",      0),
            C_CLEAR: result.get("clear_aligners",  0),
            C_VEN:   result.get("veneers",         0),
            C_IMPL:  result.get("implants",        0),
            C_SMILE: result.get("smile_makeovers", 0),
            C_WHITE: result.get("whitening",       0),
            C_SED:   result.get("sedation",        0),
            C_HOL:   result.get("holistic",        0),
            C_PLAN:  result.get("dental_plan",     ""),
            C_CANC:  result.get("cancer_screening",0),
            C_TESTI: result.get("testimonials",    "0"),
        }
        log.info("  [%03d] %-30s  CEREC=%-2s  Inv=%-3s  Impl=%-3s  Testi=%s",
                 idx, idx_name[idx][:30],
                 result.get("cerec","") or "-",
                 result.get("invisalign", 0),
                 result.get("implants", 0),
                 result.get("testimonials", "0"))

    log.info("Cache found: %d / %d  (no cache: %d)", n_found, len(idx_rows), n_missing)

    if not new_vals:
        log.warning("No cache data found — nothing to patch.")
        return out_path, None

    # ── Write patched values to xlsx ───────────────────────────────────────
    wb2 = openpyxl.load_workbook(out_path, data_only=True)
    ws2 = wb2.active
    updates = 0

    for idx, vals in new_vals.items():
        for rn in idx_rows.get(idx, []):
            for col, new_v in vals.items():
                ws2.cell(rn, col).value = new_v
                updates += 1

    wb2.save(out_path)
    log.info("Wrote %d cell updates → %s", updates, out_path)

    # ── Comparison report ──────────────────────────────────────────────────
    _write_comparison(comp_path, idx_rows, idx_name, idx_before, new_vals)
    return out_path, comp_path


# ── Comparison report ─────────────────────────────────────────────────────────

def _write_comparison(path, idx_rows, idx_name, before, after):
    from openpyxl import Workbook

    wb  = Workbook()
    ws  = wb.active
    ws.title = "Tech & Services Comparison"

    YELLOW = PatternFill("solid", fgColor="FFFACD")   # before changed
    GREEN  = PatternFill("solid", fgColor="C6EFCE")   # after changed
    GREY   = PatternFill("solid", fgColor="F2F2F2")   # before header bg
    BLUE   = PatternFill("solid", fgColor="DDEEFF")   # after header bg
    HDR_FONT  = Font(bold=True, size=9)
    DATA_FONT = Font(size=9)
    ctr = Alignment(horizontal="center", vertical="center", wrap_text=True)
    lft = Alignment(horizontal="left",   vertical="center")

    labels = [lbl for _, _, lbl in FIELD_MAP]
    n      = len(FIELD_MAP)

    # Row 1: group headers
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    ws.cell(1, 1, "Practice").font      = HDR_FONT
    ws.cell(1, 1).alignment             = ctr

    ws.merge_cells(start_row=1, start_column=4, end_row=1, end_column=3 + n)
    ws.cell(1, 4, "BEFORE (batch-N-deduped)").font = HDR_FONT
    ws.cell(1, 4).fill      = GREY
    ws.cell(1, 4).alignment = ctr

    ws.merge_cells(start_row=1, start_column=4 + n, end_row=1, end_column=3 + 2 * n)
    ws.cell(1, 4 + n, "AFTER (refreshed)").font = HDR_FONT
    ws.cell(1, 4 + n).fill      = BLUE
    ws.cell(1, 4 + n).alignment = ctr

    # Row 2: column headers
    for c, h in enumerate(["Index", "Practice Name", "# Changed"] + labels + labels, 1):
        cell = ws.cell(2, c, h)
        cell.font      = HDR_FONT
        cell.alignment = ctr
        if 4 <= c <= 3 + n:
            cell.fill = GREY
        elif c >= 4 + n:
            cell.fill = BLUE

    r = 3
    n_changed_practices = 0

    for idx in sorted(idx_rows):
        if idx not in after:
            continue

        b = before.get(idx, {})
        a = after[idx]

        before_vals = [_str(b.get(col, "")) for col, _, _ in FIELD_MAP]
        after_vals  = [_str(a.get(col, "")) for col, _, _ in FIELD_MAP]
        changed     = [bv != av for bv, av in zip(before_vals, after_vals)]
        n_ch        = sum(changed)

        if n_ch > 0:
            n_changed_practices += 1

        row_data = [idx, idx_name.get(idx, ""), n_ch] + before_vals + after_vals
        for c, v in enumerate(row_data, 1):
            cell = ws.cell(r, c, v)
            cell.font      = DATA_FONT
            cell.alignment = lft if c == 2 else ctr
            if 4 <= c <= 3 + n and changed[c - 4]:
                cell.fill = YELLOW
            elif c >= 4 + n and changed[c - 4 - n]:
                cell.fill = GREEN

        r += 1

    # Column widths
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 9
    for ci in range(4, 4 + 2 * n):
        ws.column_dimensions[get_column_letter(ci)].width = 11

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 32
    ws.freeze_panes = "D3"

    log.info("Comparison: %d practices re-extracted, %d had changes → %s",
             r - 3, n_changed_practices, path)
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
        print("Usage: python3 refresh_tech_services.py <batch_deduped.xlsx> [--cache-dir dir]")
        sys.exit(1)

    if not os.path.exists(xlsx_file):
        log.error("File not found: %s", xlsx_file)
        sys.exit(1)

    refresh(xlsx_file, cache_dir)


if __name__ == "__main__":
    main()
