#!/usr/bin/env python3
"""
ExtraDataCleaner — cleaner.py
==============================
Command-line interface for comprehensive CSV / Excel formatting cleanup.

Usage examples
──────────────
# Clean a CSV (writes  my_data_clean.csv  next to the original)
    python cleaner.py data/my_data.csv

# Clean an Excel file, all sheets → separate CSVs
    python cleaner.py data/report.xlsx

# Clean a specific sheet only
    python cleaner.py data/report.xlsx --sheet "Sales Q1"

# Keep original column capitalisation (no snake_case)
    python cleaner.py data/my_data.csv --no-snake-case

# Write output to a specific directory
    python cleaner.py data/my_data.csv --output-dir cleaned/

# Write as Excel instead of CSV
    python cleaner.py data/my_data.xlsx --format xlsx

# Force a specific header row (0-based)
    python cleaner.py messy.csv --header-row 3

# Dry-run: print cleaning report without writing files
    python cleaner.py data/my_data.csv --dry-run

# Keep hidden Excel rows/columns
    python cleaner.py report.xlsx --keep-hidden

All options
──────────────────────────────────────────────────────────────────────────────
  --output-dir PATH     Directory for output files  (default: same as input)
  --output-suffix STR   Suffix appended before extension (default: _clean)
  --format {csv,xlsx}   Output format (default: csv)
  --sheet NAME          Excel: process only this sheet (default: all)
  --header-row N        Force header at row N (0-based); default: auto-detect
  --no-snake-case       Keep original column capitalisation
  --no-null-unify       Do not replace null-like strings with NaN
  --no-numeric-coerce   Do not attempt to parse numeric columns
  --no-bool-coerce      Do not normalise boolean-like columns
  --keep-empty-rows     Do not drop fully-empty rows
  --keep-empty-cols     Do not drop fully-empty columns
  --keep-hidden         Include hidden Excel rows/columns
  --dry-run             Print report only; do not write any files
  --encoding ENC        Force output encoding for CSV (default: utf-8-sig)
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from core import DataCleaner

import pandas as pd


# ── writer ────────────────────────────────────────────────────────────────────

def _write_output(
    frames:   dict[str, pd.DataFrame],
    src_path: Path,
    out_dir:  Path,
    suffix:   str,
    fmt:      str,
    encoding: str,
) -> list[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    written: list[Path] = []

    if fmt == "xlsx":
        # All sheets → single Excel file
        stem    = src_path.stem + suffix
        out_path = out_dir / (stem + ".xlsx")
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            for sheet_name, df in frames.items():
                df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
        written.append(out_path)
    else:
        # Each sheet → separate CSV
        for sheet_name, df in frames.items():
            safe_sheet = sheet_name.replace("/", "-").replace("\\", "-")
            if len(frames) == 1 and sheet_name == "Sheet1":
                stem = src_path.stem + suffix
            else:
                stem = f"{src_path.stem}_{safe_sheet}{suffix}"
            out_path = out_dir / (stem + ".csv")
            df.to_csv(out_path, index=False, encoding=encoding)
            written.append(out_path)

    return written


# ── CLI ───────────────────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="cleaner",
        description="ExtraDataCleaner — comprehensive CSV/Excel formatter",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    p.add_argument("input", nargs="+", help="Input file(s)")
    p.add_argument("--output-dir",       metavar="PATH",   default=None)
    p.add_argument("--output-suffix",    metavar="STR",    default="_clean")
    p.add_argument("--format",           choices=["csv", "xlsx"], default="csv")
    p.add_argument("--sheet",            metavar="NAME",   default=None)
    p.add_argument("--header-row",       metavar="N",      type=int, default=None)
    p.add_argument("--encoding",         metavar="ENC",    default="utf-8-sig",
                   help="Output CSV encoding (default utf-8-sig adds BOM for Excel compat)")
    p.add_argument("--no-snake-case",    action="store_true")
    p.add_argument("--no-null-unify",    action="store_true")
    p.add_argument("--no-numeric-coerce",action="store_true")
    p.add_argument("--no-bool-coerce",   action="store_true")
    p.add_argument("--keep-empty-rows",  action="store_true")
    p.add_argument("--keep-empty-cols",  action="store_true")
    p.add_argument("--keep-hidden",      action="store_true")
    p.add_argument("--dry-run",          action="store_true")
    return p


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)

    cleaner = DataCleaner(
        snake_case       = not args.no_snake_case,
        unify_nulls      = not args.no_null_unify,
        numeric_coerce   = not args.no_numeric_coerce,
        bool_coerce      = not args.no_bool_coerce,
        drop_empty_rows  = not args.keep_empty_rows,
        drop_empty_cols  = not args.keep_empty_cols,
        skip_hidden      = not args.keep_hidden,
        sheet            = args.sheet,
        header_row       = args.header_row,
    )

    exit_code = 0

    for input_str in args.input:
        src = Path(input_str)
        if not src.exists():
            print(f"[ERROR] File not found: {src}", file=sys.stderr)
            exit_code = 1
            continue

        print(f"\n{'═'*60}")
        print(f"  Input : {src}")

        try:
            frames = cleaner.clean_file(src)
        except Exception as exc:
            print(f"[ERROR] Could not process {src}: {exc}", file=sys.stderr)
            exit_code = 1
            continue

        # Print report
        print("\n── Cleaning report ──────────────────────────────────────")
        print(cleaner.get_report())

        if args.dry_run:
            print("\n[dry-run] No files written.")
            continue

        out_dir = Path(args.output_dir) if args.output_dir else src.parent
        written = _write_output(frames, src, out_dir, args.output_suffix, args.format, args.encoding)

        print("\n── Output ────────────────────────────────────────────────")
        for p in written:
            print(f"  ✓  {p}")

    return exit_code


if __name__ == "__main__":
    # Launch GUI if no arguments are given (e.g. double-clicked)
    if len(sys.argv) == 1:
        from gui import main as gui_main
        gui_main()
    else:
        sys.exit(main())
