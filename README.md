# ExtraDataCleaner

A standalone Windows application for comprehensive CSV and Excel file formatting. Drop in a a buggered up spreadsheet, get back a clean one with on button, or customise your output. 
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![Python](https://img.shields.io/badge/python-3.10%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)

---

## What It Does

ExtraDataCleaner fixes every common formatting problem found in real-world CSV and Excel files so they load cleanly into pandas, Power BI, or any other tool without any further manual cleanup.

| Problem | Fix Applied |
|---|---|
| BOM characters, mixed encodings | Auto-detected, normalised to UTF-8 |
| CRLF / mixed line endings | Normalised |
| Title or metadata rows above the real header | Auto-skipped |
| Wrong delimiter (`;` `\t` `\|` instead of `,`) | Auto-detected, ignores commas inside quoted fields |
| Leading / trailing whitespace in cells | Stripped |
| Stray wrapping quotes around cell values | Removed |
| Null variants — `N/A`, `none`, `null`, `-`, `#N/A`, `nan` … | Unified to real `NaN` |
| Excel errors — `#VALUE!`, `#REF!`, `#DIV/0!` … | Replaced with `NaN` |
| Currency symbols — `$` `€` `£` `¥` … | Stripped |
| Thousands separators — `1,234` and EU-style `1.234` | Removed |
| Decimal comma — `1.234,56` | Converted to `1234.56` |
| Accounting negatives — `(500.00)` | Converted to `-500.0` |
| Percentages — `87.5%` | Converted to `0.875` float |
| Numbers stored as text | Coerced to `int64` / `float64` |
| Boolean strings — `yes/no/TRUE/Y/1/false` | Coerced to Python `bool` |
| Column names with spaces, special characters, duplicates | Cleaned to `snake_case` |
| Fully empty rows and columns | Dropped |
| Excel merged cells | Forward-filled |
| Hidden Excel rows / columns | Dropped (optional) |
| Multi-sheet Excel files | All sheets processed, output as separate CSVs or a single `.xlsx` |

---

## Running the Application

### Option A — Double-click the .exe (no setup required)

Download `ExtraDataCleaner.exe` and double-click it. You should be sweet.

### Option B — Run from source

```bash
pip install -r requirements.txt
python cleaner.py
```

### Option C — CLI usage

```bash
# Clean a CSV
python cleaner.py data/sales.csv

# Clean an Excel file (all sheets → separate CSVs)
python cleaner.py data/report.xlsx

# Specific sheet only
python cleaner.py data/report.xlsx --sheet "Sales Q1"

# Write output to a specific folder
python cleaner.py data/sales.csv --output-dir cleaned/

# Output as Excel instead of CSV
python cleaner.py data/report.xlsx --format xlsx

# Dry run — print report without writing any files
python cleaner.py data/sales.csv --dry-run

# Force header at row 3 (0-based)
python cleaner.py messy.csv --header-row 3

# Keep original column capitalisation (no snake_case)
python cleaner.py data/sales.csv --no-snake-case
```

**All CLI flags:**

| Flag | Default | Description |
|---|---|---|
| `--output-dir PATH` | Same as input | Directory for output files |
| `--output-suffix STR` | `_clean` | Suffix added before extension |
| `--format {csv,xlsx}` | `csv` | Output format |
| `--sheet NAME` | All sheets | Excel: process this sheet only |
| `--header-row N` | Auto-detect | Force header at row N (0-based) |
| `--no-snake-case` | — | Keep original column capitalisation |
| `--no-null-unify` | — | Don't replace null-like strings with NaN |
| `--no-numeric-coerce` | — | Don't parse numeric columns |
| `--no-bool-coerce` | — | Don't normalise boolean columns |
| `--keep-empty-rows` | — | Don't drop fully-empty rows |
| `--keep-empty-cols` | — | Don't drop fully-empty columns |
| `--keep-hidden` | — | Include hidden Excel rows/columns |
| `--dry-run` | — | Report only, no files written |
| `--encoding ENC` | `utf-8-sig` | Output CSV encoding |

---

## Building the .exe

> Only needed if you want to compile it yourself. Pre-built releases are available on the [Releases](../../releases) page.

**Requirements:** Python 3.10+ installed on Windows with "Add Python to PATH" checked.

1. Open a terminal inside the `ExtraDataCleaner` folder
2. Run:

```bat
build_exe.bat
```

The script will:
- Install all required packages including PyInstaller and Pillow
- Generate `icon.ico` (multi-resolution brush icon)
- Compile everything into a single `ExtraDataCleaner.exe`
- Offer to clean up the build folders

Build time is roughly 1–3 minutes on first run.

---

## Project Structure

```
ExtraDataCleaner/
├── cleaner.py              # CLI entry point — launches GUI if no args given
├── core.py                 # DataCleaner class — all cleaning logic lives here
├── gui.py                  # Windows classic-themed tkinter GUI
├── make_icon.py            # Build-time brush icon generator (Pillow + pure-Python fallback)
├── ExtraDataCleaner.spec   # PyInstaller build configuration
├── build_exe.bat           # One-click Windows .exe builder
└── requirements.txt        # Runtime dependencies
```

### core.py — DataCleaner

The `DataCleaner` class can be used directly in your own scripts:

```python
from core import DataCleaner

cleaner = DataCleaner(
    snake_case=True,        # Convert column names to snake_case
    unify_nulls=True,       # Replace null-like strings with NaN
    numeric_coerce=True,    # Parse numbers stored as text
    bool_coerce=True,       # Normalise yes/no/TRUE etc. to bool
    drop_empty_rows=True,
    drop_empty_cols=True,
    skip_hidden=True,       # Ignore hidden Excel rows/cols
)

frames = cleaner.clean_file("my_data.xlsx")  # returns {sheet_name: DataFrame}
df = frames["Sheet1"]

print(cleaner.get_report())  # Human-readable log of every change made
```

---

## Dependencies

| Package | Purpose |
|---|---|
| `pandas` | DataFrame operations |
| `numpy` | Numeric type handling |
| `openpyxl` | Excel read/write |
| `chardet` | Encoding detection |
| `python-dateutil` | Date parsing |
| `pillow` *(build only)* | Icon generation |
| `pyinstaller` *(build only)* | .exe compilation |

