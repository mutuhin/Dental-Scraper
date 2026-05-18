#!/usr/bin/env python3
"""
qa_check.py
───────────
Reads a scraped batch Excel and produces a colour-coded QA report showing
exactly which fields are captured, missing, or possibly incomplete — grouped
by practice (Index).

Output sheets:
  1. QA Results      – one row per practice, every field colour-coded
  2. Needs Re-scrape – practices with critical/important fields missing
  3. Field Stats     – fill-rate % per column across all practices

Usage:
    python3 qa_check.py batch_1_deduped.xlsx
    python3 qa_check.py batch_1_deduped.xlsx --out my_qa.xlsx
"""

import sys, os, re, argparse, collections
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Constants ─────────────────────────────────────────────────────────────────
EMPTY_VALUES = {
    None, "", "Not Found", "N/A", "nan", "None", "none",
    "ERROR", "N/A – Not Offered", "0",
}

# Fields and their importance tier
# critical  → must have; drives FAIL status
# important → should have; drives REVIEW status
# optional  → nice to have; no status impact
# tech/service → checked as a group (expect at least 1 populated if site accessible)
QA_FIELDS = [
    # col_header (must match exactly)              tier
    ("Doctor Name",                    "critical"),
    ("Practice Email",                 "critical"),
    ("Google Reviews Ranking",         "critical"),
    ("Total # of Google Reviews",      "critical"),
    ("# of Hygienists",                "important"),
    ("Facebook URL",                   "important"),
    ("Associations / Memberships",     "important"),
    ("Doctor Specialty",               "important"),
    ("Instagram URL",                  "optional"),
    ("TikTok URL",                     "optional"),
    ("Yelp Rating",                    "optional"),
    ("Testimonials (Number of)",       "optional"),
    ("# of Locations",                 "optional"),
    # Technology (checked as a group)
    ("CEREC (Same Day Crowns)",        "tech"),
    ("CBCT (3D Imaging)",              "tech"),
    ("Lasers",                         "tech"),
    ("AI",                             "tech"),
    ("Intraoral Scanners",             "tech"),
    # Services (checked as a group)
    ("Invisalign (Mentions)",          "service"),
    ("Clear Aligners",                 "service"),
    ("Veneers",                        "service"),
    ("Implants",                       "service"),
    ("Smile Makeovers",                "service"),
    ("Teeth Whitening",                "service"),
    ("Sedation Dentistry",             "service"),
    ("Holistic Dentistry",             "service"),
    ("Dental Plan (Membership Plan)",  "service"),
    ("Cancer Screening",               "service"),
]

# Colour fills
F_GREEN  = PatternFill("solid", fgColor="C6EFCE")   # captured
F_YELLOW = PatternFill("solid", fgColor="FFEB9C")   # partial / group has ≥1
F_RED    = PatternFill("solid", fgColor="FFC7CE")   # missing
F_GREY   = PatternFill("solid", fgColor="EDEDED")   # N/A / not applicable
F_BLUE   = PatternFill("solid", fgColor="BDD7EE")   # header
F_DKBLUE = PatternFill("solid", fgColor="1F4E79")
F_DKRED  = PatternFill("solid", fgColor="9C0006")
F_DKGRN  = PatternFill("solid", fgColor="375623")
F_DKORG  = PatternFill("solid", fgColor="833C00")

FONT_HDR  = Font(name="Arial", bold=True, size=9, color="FFFFFF")
FONT_DATA = Font(name="Arial", size=9)
FONT_BOLD = Font(name="Arial", bold=True, size=9)
FONT_RED  = Font(name="Arial", bold=True, size=9, color="9C0006")
FONT_GRN  = Font(name="Arial", bold=True, size=9, color="375623")
ALIGN_CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LFT = Alignment(horizontal="left",   vertical="center", wrap_text=True)
THIN = Side(style="thin", color="CCCCCC")
BDR  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _is_empty(val) -> bool:
    return str(val).strip() in EMPTY_VALUES


def _cell(ws, r, c, value="", font=FONT_DATA, fill=None, align=ALIGN_CTR):
    cell = ws.cell(r, c, value)
    cell.font   = font
    cell.border = BDR
    cell.alignment = align
    if fill:
        cell.fill = fill
    return cell


# ── Load input ────────────────────────────────────────────────────────────────

def load_batch(path: str):
    """
    Returns:
        headers: list of column names (row 2 of the file)
        practices: dict {index -> {col_name -> best_value}}
        practice_names: dict {index -> Practice Name}
        row_counts: dict {index -> number of doctor rows}
    """
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb.active

    # Find header row (row 2 in our format)
    headers = []
    for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
        non_null = [v for v in row if v]
        if "Index" in non_null or "Doctor Name" in non_null:
            headers = [str(v).strip() if v else "" for v in row]
            break

    if not headers:
        raise ValueError("Could not find header row with 'Index' column")

    col_idx = {h: i for i, h in enumerate(headers) if h}

    practices     = collections.defaultdict(dict)   # index → {field → best value}
    practice_names = {}
    row_counts    = collections.Counter()

    for row in ws.iter_rows(min_row=3, values_only=True):
        raw_idx = row[col_idx.get("Index", 0)]
        try:
            idx = int(float(str(raw_idx)))
        except Exception:
            continue

        row_counts[idx] += 1
        if idx not in practice_names:
            practice_names[idx] = str(row[col_idx.get("Practice Name", 1)] or "")

        for field, _ in QA_FIELDS:
            ci = col_idx.get(field)
            if ci is None:
                continue
            val = row[ci]
            # Keep the best (non-empty) value seen across doctor rows for this index
            existing = practices[idx].get(field)
            if not _is_empty(val):
                practices[idx][field] = val
            elif existing is None:
                practices[idx][field] = val

    wb.close()
    return headers, dict(practices), practice_names, dict(row_counts)


# ── Scoring ───────────────────────────────────────────────────────────────────

def score_practice(data: dict) -> dict:
    """
    Returns a dict with:
        status        PASS | REVIEW | FAIL
        critical_miss list of missing critical fields
        important_miss list of missing important fields
        tech_any      bool — at least one tech field populated
        service_any   bool — at least one service field populated
        pct           overall fill % (critical + important only)
    """
    critical_miss  = []
    important_miss = []
    tech_vals      = []
    service_vals   = []

    for field, tier in QA_FIELDS:
        val = data.get(field)
        empty = _is_empty(val)
        if tier == "critical":
            if empty:
                critical_miss.append(field)
        elif tier == "important":
            if empty:
                important_miss.append(field)
        elif tier == "tech":
            tech_vals.append(not empty)
        elif tier == "service":
            service_vals.append(not empty)

    tech_any    = any(tech_vals)
    service_any = any(service_vals)

    total_ci = sum(1 for _, t in QA_FIELDS if t in ("critical", "important"))
    filled   = total_ci - len(critical_miss) - len(important_miss)
    pct      = round(filled / total_ci * 100) if total_ci else 0

    if critical_miss:
        status = "FAIL"
    elif len(important_miss) >= 2 or not tech_any or not service_any:
        status = "REVIEW"
    else:
        status = "PASS"

    return {
        "status":         status,
        "critical_miss":  critical_miss,
        "important_miss": important_miss,
        "tech_any":       tech_any,
        "service_any":    service_any,
        "pct":            pct,
    }


# ── Sheet 1: QA Results ───────────────────────────────────────────────────────

def write_qa_results(wb: Workbook, practices: dict, practice_names: dict,
                     row_counts: dict):
    ws = wb.create_sheet("QA Results")

    qa_cols = [f for f, _ in QA_FIELDS]
    tier_of = {f: t for f, t in QA_FIELDS}

    # Group headers row 1
    group_spans = [
        ("Practice",   ["Doctor Name", "Practice Email", "# of Hygienists"],        F_DKBLUE),
        ("Social",     ["Facebook URL", "Instagram URL", "TikTok URL"],              F_DKGRN),
        ("Tech",       [f for f, t in QA_FIELDS if t == "tech"],                    PatternFill("solid", fgColor="7030A0")),
        ("Services",   [f for f, t in QA_FIELDS if t == "service"],                 F_DKORG),
        ("Doctor",     ["Associations / Memberships", "Doctor Specialty"],           PatternFill("solid", fgColor="595959")),
        ("Reviews",    ["Google Reviews Ranking", "Total # of Google Reviews",
                        "Yelp Rating", "Testimonials (Number of)", "# of Locations"], F_DKBLUE),
    ]

    # Fixed columns: Index, Practice Name, # Doctors, Status, Score
    fixed = ["Index", "Practice Name", "# Doctor Rows", "Status", "Score %"]
    total_cols = len(fixed) + len(qa_cols)

    # Row 1 — group spans
    _cell(ws, 1, 1, "Practice Info", font=FONT_HDR, fill=F_DKBLUE)
    col_offset = len(fixed) + 1
    for grp_label, fields, fill in group_spans:
        count = sum(1 for f in fields if f in qa_cols)
        if count == 0:
            continue
        start_c = col_offset
        ws.merge_cells(start_row=1, start_column=start_c,
                       end_row=1,   end_column=start_c + count - 1)
        _cell(ws, 1, start_c, grp_label, font=FONT_HDR, fill=fill)
        col_offset += count
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1,   end_column=len(fixed))

    # Row 2 — column headers
    for c, h in enumerate(fixed, 1):
        _cell(ws, 2, c, h, font=FONT_HDR, fill=F_DKBLUE)
    for c, h in enumerate(qa_cols, len(fixed) + 1):
        tier = tier_of.get(h, "")
        fill_map = {
            "critical": F_DKBLUE, "important": F_DKBLUE,
            "tech": PatternFill("solid", fgColor="7030A0"),
            "service": F_DKORG, "optional": PatternFill("solid", fgColor="595959"),
        }
        _cell(ws, 2, c, h, font=FONT_HDR, fill=fill_map.get(tier, F_DKBLUE))

    # Freeze header rows
    ws.freeze_panes = "A3"

    # Data rows
    counts = {"PASS": 0, "REVIEW": 0, "FAIL": 0}
    field_filled  = collections.Counter()
    field_total   = collections.Counter()

    sorted_idxs = sorted(practices.keys())
    for row_num, idx in enumerate(sorted_idxs, 3):
        data   = practices[idx]
        name   = practice_names.get(idx, "")
        ndocs  = row_counts.get(idx, 1)
        score  = score_practice(data)
        status = score["status"]
        counts[status] += 1

        rf = PatternFill("solid", fgColor="FAFAFA") if row_num % 2 == 0 else None

        status_fill = {"PASS": F_GREEN, "REVIEW": F_YELLOW, "FAIL": F_RED}[status]
        status_font = {"PASS": FONT_GRN, "REVIEW": FONT_BOLD, "FAIL": FONT_RED}[status]

        _cell(ws, row_num, 1, idx,    font=FONT_DATA, fill=rf, align=ALIGN_CTR)
        _cell(ws, row_num, 2, name,   font=FONT_DATA, fill=rf, align=ALIGN_LFT)
        _cell(ws, row_num, 3, ndocs,  font=FONT_DATA, fill=rf, align=ALIGN_CTR)
        _cell(ws, row_num, 4, status, font=status_font, fill=status_fill, align=ALIGN_CTR)
        _cell(ws, row_num, 5, f"{score['pct']}%", font=FONT_DATA, fill=rf, align=ALIGN_CTR)

        for c, field in enumerate(qa_cols, len(fixed) + 1):
            val  = data.get(field)
            tier = tier_of.get(field, "")
            empty = _is_empty(val)
            field_total[field] += 1

            if not empty:
                cell_fill = F_GREEN
                disp = "✓"
                field_filled[field] += 1
            else:
                # Tech/service: yellow if the GROUP has at least one
                if tier == "tech" and score["tech_any"]:
                    cell_fill = F_YELLOW
                    disp = "–"
                elif tier == "service" and score["service_any"]:
                    cell_fill = F_YELLOW
                    disp = "–"
                else:
                    cell_fill = F_RED
                    disp = "✗"

            _cell(ws, row_num, c, disp, font=FONT_DATA, fill=cell_fill, align=ALIGN_CTR)

    # Summary row at top (row 3 offset — insert after headers)
    sum_row = ws.max_row + 2
    _cell(ws, sum_row, 1, "TOTALS", font=FONT_HDR, fill=F_DKBLUE, align=ALIGN_CTR)
    _cell(ws, sum_row, 2,
          f"PASS {counts['PASS']}  |  REVIEW {counts['REVIEW']}  |  FAIL {counts['FAIL']}",
          font=FONT_HDR, fill=F_DKBLUE, align=ALIGN_LFT)
    _cell(ws, sum_row, 3, len(sorted_idxs), font=FONT_HDR, fill=F_DKBLUE)
    for c, field in enumerate(qa_cols, len(fixed) + 1):
        total = field_total.get(field, 1)
        filled = field_filled.get(field, 0)
        pct = round(filled / total * 100) if total else 0
        fill = F_GREEN if pct >= 80 else (F_YELLOW if pct >= 50 else F_RED)
        _cell(ws, sum_row, c, f"{pct}%", font=FONT_HDR, fill=fill, align=ALIGN_CTR)

    # Column widths
    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 9
    ws.column_dimensions["E"].width = 8
    for c in range(len(fixed) + 1, len(fixed) + len(qa_cols) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 10

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 45

    return counts, field_filled, field_total


# ── Sheet 2: Needs Re-scrape ──────────────────────────────────────────────────

def write_needs_rescrape(wb: Workbook, practices: dict, practice_names: dict):
    ws = wb.create_sheet("Needs Re-scrape")

    headers = ["Index", "Practice Name", "Status", "Score %",
               "Missing Critical", "Missing Important",
               "Tech Group", "Service Group"]
    for c, h in enumerate(headers, 1):
        _cell(ws, 1, c, h, font=FONT_HDR, fill=F_DKBLUE)

    ws.freeze_panes = "A2"

    rows_written = 0
    for idx in sorted(practices.keys()):
        data  = practices[idx]
        score = score_practice(data)
        if score["status"] == "PASS":
            continue

        rows_written += 1
        r = rows_written + 1
        rf = PatternFill("solid", fgColor="FAFAFA") if r % 2 == 0 else None
        status = score["status"]
        sf = {"REVIEW": F_YELLOW, "FAIL": F_RED}[status]
        font = {"REVIEW": FONT_BOLD, "FAIL": FONT_RED}[status]

        _cell(ws, r, 1, idx,    font=FONT_DATA, fill=rf)
        _cell(ws, r, 2, practice_names.get(idx, ""), font=FONT_DATA, fill=rf, align=ALIGN_LFT)
        _cell(ws, r, 3, status, font=font, fill=sf)
        _cell(ws, r, 4, f"{score['pct']}%", font=FONT_DATA, fill=rf)
        _cell(ws, r, 5, ", ".join(score["critical_miss"])  or "—", font=FONT_RED if score["critical_miss"] else FONT_DATA,  fill=rf, align=ALIGN_LFT)
        _cell(ws, r, 6, ", ".join(score["important_miss"]) or "—", font=FONT_DATA, fill=rf, align=ALIGN_LFT)
        _cell(ws, r, 7, "✓" if score["tech_any"]     else "✗ None found",
              font=FONT_GRN if score["tech_any"]     else FONT_RED,
              fill=F_GREEN  if score["tech_any"]     else F_RED)
        _cell(ws, r, 8, "✓" if score["service_any"]  else "✗ None found",
              font=FONT_GRN if score["service_any"]  else FONT_RED,
              fill=F_GREEN  if score["service_any"]  else F_RED)

    if rows_written == 0:
        _cell(ws, 2, 1, "✅ All practices passed QA!", font=FONT_GRN, fill=F_GREEN, align=ALIGN_LFT)

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 9
    ws.column_dimensions["D"].width = 8
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 40
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 15


# ── Sheet 3: Field Stats ──────────────────────────────────────────────────────

def write_field_stats(wb: Workbook, field_filled: dict, field_total: dict):
    ws = wb.create_sheet("Field Stats")

    headers = ["Field", "Tier", "Filled", "Total", "Fill Rate", "Status"]
    for c, h in enumerate(headers, 1):
        _cell(ws, 1, c, h, font=FONT_HDR, fill=F_DKBLUE)

    ws.freeze_panes = "A2"

    tier_of = {f: t for f, t in QA_FIELDS}

    for r, (field, tier) in enumerate(QA_FIELDS, 2):
        total  = field_total.get(field, 0)
        filled = field_filled.get(field, 0)
        pct    = round(filled / total * 100) if total else 0
        rf = PatternFill("solid", fgColor="FAFAFA") if r % 2 == 0 else None

        if pct >= 80:
            status, sf = "Good", F_GREEN
        elif pct >= 50:
            status, sf = "Review", F_YELLOW
        else:
            status, sf = "Low", F_RED

        _cell(ws, r, 1, field,  font=FONT_DATA, fill=rf, align=ALIGN_LFT)
        _cell(ws, r, 2, tier,   font=FONT_DATA, fill=rf)
        _cell(ws, r, 3, filled, font=FONT_DATA, fill=rf)
        _cell(ws, r, 4, total,  font=FONT_DATA, fill=rf)
        _cell(ws, r, 5, f"{pct}%", font=FONT_BOLD, fill=sf)
        _cell(ws, r, 6, status,    font=FONT_DATA, fill=sf)

    ws.column_dimensions["A"].width = 35
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 9
    ws.column_dimensions["D"].width = 9
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 10


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input",  help="Scraped batch Excel file")
    parser.add_argument("--out",  help="Output QA report path (default: qa_<input>)")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"ERROR: File not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    out_path = args.out or ("qa_" + os.path.basename(args.input))

    print(f"Loading: {args.input}")
    headers, practices, practice_names, row_counts = load_batch(args.input)
    print(f"Found {len(practices)} unique practices")

    wb = Workbook()
    wb.remove(wb.active)   # remove default sheet

    counts, field_filled, field_total = write_qa_results(wb, practices, practice_names, row_counts)
    write_needs_rescrape(wb, practices, practice_names)
    write_field_stats(wb, field_filled, field_total)

    wb.save(out_path)

    total = len(practices)
    print(f"\n{'='*50}")
    print(f"QA SUMMARY — {total} practices")
    print(f"  ✅ PASS   : {counts['PASS']}  ({round(counts['PASS']/total*100)}%)")
    print(f"  ⚠️  REVIEW : {counts['REVIEW']}  ({round(counts['REVIEW']/total*100)}%)")
    print(f"  ❌ FAIL   : {counts['FAIL']}  ({round(counts['FAIL']/total*100)}%)")
    print(f"\nLowest fill-rate fields:")
    stats = [(f, field_filled.get(f,0), field_total.get(f,1)) for f, _ in QA_FIELDS]
    stats.sort(key=lambda x: x[1]/x[2])
    for field, filled, total_f in stats[:5]:
        print(f"  {field}: {round(filled/total_f*100)}%")
    print(f"\nSaved → {out_path}")


if __name__ == "__main__":
    main()
