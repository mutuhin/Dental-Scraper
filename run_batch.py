"""
run_batch.py
─────────────
Wrapper for dental_scraper.py that accepts a batch number from the command line.
Used by GitHub Actions to run scraping in chunks of 100 rows.

Usage:
    python run_batch.py <batch_number>

Examples:
    python run_batch.py 1   →  rows   1-100  →  batch_01_rows1_100.xlsx
    python run_batch.py 2   →  rows 101-200  →  batch_02_rows101_200.xlsx
    python run_batch.py 3   →  rows 201-300  →  batch_03_rows201_300.xlsx
"""

import sys
import os

BATCH_SIZE   = 100
INPUT_FILE   = "6000 Data COMPLETE.xlsx"   # must be in the same folder

def main():
    batch = int(sys.argv[1]) if len(sys.argv) > 1 else 1

    start_idx = (batch - 1) * BATCH_SIZE          # 0-based
    end_idx   = batch * BATCH_SIZE                 # exclusive

    output_file = (
        f"batch_{batch:02d}_rows{start_idx + 1}_{end_idx}.xlsx"
    )

    print(f"Batch {batch}: rows {start_idx + 1}–{end_idx}  →  {output_file}")

    # Patch module-level constants BEFORE calling main() so it picks them up
    import dental_scraper as ds
    ds.INPUT_FILE   = INPUT_FILE
    ds.OUTPUT_FILE  = output_file
    ds.START_IDX    = start_idx
    ds.END_IDX      = end_idx

    ds.main()

if __name__ == "__main__":
    main()
