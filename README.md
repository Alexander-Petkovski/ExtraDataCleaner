# ExtraDataCleaner

> **One-click cleanup for messy CSV and Excel files and no extra work required.**

---

## The Problem

If you've ever received a spreadsheet from a client, an external system, or a colleague, you already know: **the data is never clean.**, it's probably buggered.

You open it and find things like:

- Currency symbols stuck inside number columns (`$1,234.56` instead of `1234.56`)
- A mix of ways to say "nothing" , some cells say `N/A`, others say `null`, `none`, `-`, `missing`, or just nothing at all
- Numbers that look like numbers but Excel treats as text, so your formulas break
- Columns named `First Name` when your system expects `first_name`
- Negative numbers written as `(500.00)` instead of `-500.00`
- Percentages written as `75%` instead of `0.75`
- Hidden rows, merged cells, and junk title rows sitting above the real data
- Files saved in the wrong encoding so special characters come out as `â€™` garbage

Before you can actually use the data, go to analyse it, import it, run reports on it, **you have to fix all of this manually.** That can take hours, and it has to be done every single time a new file arrives.

**ExtraDataCleaner fixes all of it automatically, in seconds.**

---

## What It Does

Drop in a CSV or Excel file, click **Run Cleaner**, and get back a clean, consistent, analysis-ready file. No formulas. No macros. No Python knowledge needed.

It handles over 30 categories of data problems:

| Problem | What ExtraDataCleaner does |
|---|---|
| `$1,234.56` / `£99.00` / `€1.499,99` | Strips symbols, fixes decimal format → `1234.56` |
| `(500.00)` accounting negatives | Converts to proper `-500.00` |
| `75%` percentages | Converts to numeric `75.0` |
| `N/A`, `null`, `none`, `missing`, `-`, `--` | Unified to a single blank (NaN) |
| `yes` / `Y` / `1` / `True` / `on` | Normalised to `True` / `False` |
| `#VALUE!`, `#REF!`, `#DIV/0!` | Replaced with blank |
| `First Name`, `Total Sales ($)` | Cleaned to `first_name`, `total_sales` |
| Numbers stored as text | Converted to real numeric values |
| Merged cells in Excel | Forward-filled so every row has a value |
| Hidden rows / columns | Detected and removed |
| Junk title rows above the real header | Auto-detected and skipped |
| BOM characters, wrong encoding, mixed line endings | Normalised to clean UTF-8 |
| Fully blank rows or columns | Removed automatically |

It also includes a **Sort** feature — reorder rows by any column (alphabetically, numerically, by date, by frequency, or natural alphanumeric order) with the sorted result permanently written to the output file.

---

## Download

**→ [Download ExtraDataCleaner.exe](../../releases/latest)** — double-click to run, no installation needed.

Runs on Windows. No Python. No setup.

---

## How to Use

### The easy way (GUI)

1. Double-click `ExtraDataCleaner.exe`
2. Click **Browse** and select your CSV or Excel file(s)
3. Choose your output options (same folder, different folder, CSV or Excel output)
4. Optionally add sort rules
5. Click **Run Cleaner**

The cleaned file is saved next to your original with `_clean` added to the filename (e.g. `sales_data_clean.csv`).

### The power way (command line)

```bash
# Clean a single file
ExtraDataCleaner.exe data/sales.csv

# Clean and sort by last name A→Z
ExtraDataCleaner.exe data/customers.csv --sort-by "last_name:asc:natural"

# Output as Excel instead of CSV
ExtraDataCleaner.exe data/report.xlsx --format xlsx

# Preview what would be cleaned without writing any files
ExtraDataCleaner.exe data/sales.csv --dry-run
```

### All command-line options

| Option | Description |
|---|---|
| `--output-dir PATH` | Where to save the cleaned file (default: same folder as input) |
| `--output-suffix STR` | Text added before the extension (default: `_clean`) |
| `--format csv\|xlsx` | Output as CSV or Excel (default: csv) |
| `--sheet NAME` | Excel only: process one specific sheet |
| `--header-row N` | Force the header to a specific row number (default: auto-detect) |
| `--no-snake-case` | Keep original column names as-is |
| `--no-null-unify` | Don't replace null-like strings with blank |
| `--no-numeric-coerce` | Don't convert text to numbers |
| `--no-bool-coerce` | Don't normalise yes/no/true/false columns |
| `--keep-empty-rows` | Don't remove blank rows |
| `--keep-empty-cols` | Don't remove blank columns |
| `--keep-hidden` | Include hidden Excel rows and columns |
| `--dry-run` | Print a report of what would change, without writing anything |
| `--sort-by SPEC` | Sort after cleaning — format: `"column:asc\|desc:method"` |

**Sort methods:** `natural` (A1 < A2 < A10) · `alpha` (A–Z) · `numeric` (0→9) · `date` (oldest→newest) · `length` (shortest→longest) · `frequency` (most common first)

---

## Building from Source

Requires Python 3.10+ and Windows.

```
build_exe.bat
```

Finds Python automatically, installs all dependencies, and produces a standalone `ExtraDataCleaner.exe`. The resulting exe runs on any Windows machine with no Python required.

**Dependencies:**

| Package | Why it's needed |
|---|---|
| `pandas` | Core data loading, transformation and export |
| `numpy` | Numeric operations |
| `openpyxl` | Reading and writing Excel files |
| `chardet` | Auto-detects file encoding |
| `python-dateutil` | Date parsing |
| `natsort` | Natural sort order (Item 2 before Item 10) |

---

## Project Structure

```
ExtraDataCleaner/
├── cleaner.py        # Command-line entry point
├── core.py           # All cleaning logic (DataCleaner class)
├── gui.py            # Desktop GUI (Windows classic theme)
├── make_icon.py      # Generates the app icon at build time
├── build_exe.bat     # One-click .exe builder for Windows
├── ExtraDataCleaner.spec  # PyInstaller packaging config
└── requirements.txt  # Python dependencies
```

---

## Companion Tool

**ExtraDataGenerator** → [github.com/Alexander-Petkovski/ExtraDataGenerator](https://github.com/Alexander-Petkovski/ExtraDataGenerator)

Generates intentionally messy CSV/Excel test files with the exact problems ExtraDataCleaner fixes for useful for testing, demos, and regression checks.

---

## License

MIT
