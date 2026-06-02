---
name: testing-find-duplicates
description: Test the find_duplicates.py script that detects duplicate Word files in numbered subdirectories. Use when verifying duplicate detection, date parsing, or HTML report generation.
---

# Testing find_duplicates.py

## What It Does

`find_duplicates.py` scans a root directory for subdirectories starting with a digit, finds Word files (.docx/.doc) with the same base name (ignoring date suffixes), and generates an HTML report (`dubbele_bestanden.html`) with color-coded entries (green = newest, red = older).

## How to Run

```bash
cd /home/ubuntu/repos/Word-instructie-helper
. .venv/bin/activate
python find_duplicates.py <path-to-root-directory>
```

Output: `<path-to-root-directory>/dubbele_bestanden.html`

## Test Setup

No credentials or services needed — this is a standalone CLI tool. Create temp directories with dummy `.docx`/`.doc` files using `touch`.

Directory names must start with a digit (e.g., `1 Handleidingen`, `2 Procedures`) to be scanned.

## Supported Date Formats in Filenames

The script recognizes three date patterns at the end of the filename stem:
- `YYYY-MM-DD` (also `YYYY_MM_DD`, `YYYY.MM.DD`)
- `DD-MM-YYYY` (also `DD_MM_YYYY`, `DD.MM.YYYY`)
- `YYYYMMDD` (8 compact digits)

## Key Test Scenarios

1. **Happy path**: Multiple files with same base name + different dates → duplicates detected, newest marked green
2. **Different date formats**: DD-MM-YYYY and YYYYMMDD are parsed correctly
3. **Non-numbered dirs**: Directories without a leading digit are ignored
4. **Unique files**: Files with no duplicate base name don't appear in report
5. **Mixed extensions**: `.doc` and `.docx` with same base name are grouped together
6. **Temp files**: `~$...` prefixed files (Word temp files) are excluded
7. **Error handling**: No arguments → usage + exit 1; invalid path → error + exit 1
8. **No duplicates found**: Message printed, no HTML file generated

## Testing Method

All testing is shell-based (no GUI recording needed). Create temp dirs, run the script, then validate:
- stdout messages (duplicate count or "Geen dubbele bestanden gevonden.")
- Exit codes
- HTML content: grep for `class="nieuwste"` and `class="oudere"` to verify correct color assignment
- Absence checks: grep for files that should NOT appear (unique files, temp files)

## Devin Secrets Needed

None — this tool requires no API keys or credentials.
