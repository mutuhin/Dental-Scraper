"""
run_batch.py
─────────────
Wrapper for dental_scraper.py that accepts a batch number from the command line.
Used by GitHub Actions to run scraping in chunks of 100 rows.

Usage:
    python run_batch.py <batch_number> [start_row]

Examples:
    python run_batch.py 1       →  rows   1-100  →  batch_01_rows1_100.xlsx
    python run_batch.py 1 95    →  rows  95-100  →  batch_01_rows1_100.xlsx  (resume)
    python run_batch.py 2       →  rows 101-200  →  batch_02_rows101_200.xlsx
"""

import sys

BATCH_SIZE = 100
INPUT_FILE = "6000 Data COMPLETE.xlsx"

def main():
    batch = int(sys.argv[1]) if len(sys.argv) > 1 else 1

    batch_start = (batch - 1) * BATCH_SIZE + 1   # 1-based first row of this batch
    batch_end   = batch * BATCH_SIZE              # 1-based last row

    # Optional start_row lets you resume mid-batch (e.g. "python run_batch.py 1 95")
    if len(sys.argv) > 2 and sys.argv[2].strip():
        start_row = int(sys.argv[2])
        if start_row < batch_start or start_row > batch_end:
            print(f"start_row {start_row} is outside batch {batch} range ({batch_start}-{batch_end})")
            sys.exit(1)
    else:
        start_row = batch_start

    start_idx = start_row - 1        # 0-based
    end_idx   = batch_end            # exclusive upper bound (0-based)

    output_file = f"batch_{batch:02d}_rows{batch_start}_{batch_end}.xlsx"
    print(f"Batch {batch}: scraping rows {start_row}–{batch_end}  →  {output_file}")

    import dental_scraper as ds
    ds.INPUT_FILE  = INPUT_FILE
    ds.OUTPUT_FILE = output_file
    ds.START_IDX   = start_idx
    ds.END_IDX     = end_idx

    ds.main()

if __name__ == "__main__":
    main()
