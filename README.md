# rcpa2csv

Convert RCPA Excel files into tab-delimited text files for SNAP2SNOMED import.

The script scans an input directory for Excel files, extracts values from one column, normalizes/splits values, and writes one output file per source workbook.

## What It Produces

For each input Excel file, the script writes a `.txt` file with 2 columns:

- `code`: sequential integer (`1..n`)
- `display`: extracted specimen text value

Rules applied while building output:

- Values are split on `;`
- Leading/trailing whitespace is trimmed
- Empty values are removed
- Duplicates are removed
- Final rows are sorted by `display`
- `Serum;Plasma` (or `Plasma;Serum`) is treated as a single value

## Requirements

- Python 3.13+
- pandas

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run with defaults:

```bash
python main.py
```

Run with explicit paths/settings:

```bash
python main.py \
  --indir /path/to/input \
  --outdir /path/to/output \
  --sheet_name "RCPA SPIA Requesting_Mar 2026" \
  --column_name "Specimen"
```

## CLI Arguments

- `-i, --indir`: input directory containing Excel files
- `-o, --outdir`: output directory for generated `.txt` files
- `-s, --sheet_name`: sheet name to read from each workbook
- `-c, --column_name`: column to extract values from

Default values in the current script:

- `indir`: `$HOME/data/rcpa/in`
- `outdir`: `$HOME/data/rcpa/out`
- `sheet_name`: `RCPA SPIA Requesting_Mar 2026`
- `column_name`: `Specimen`

## Notes

- Temporary Excel lock files (starting with `~$`) are ignored.
- Only these file extensions are processed: `.xls`, `.xlsx`, `.xlsm`, `.xlsb`.
- Logs are written under `$HOME/data/ucum/logs`.
