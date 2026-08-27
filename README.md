# MOWR Website Date Updater

This tool reads topic names from the Word report (`ওয়েবসাইট_হালনাগাদ_তথ্য_আপডেটেড.docx`), checks each page on [https://mowr.gov.bd/](https://mowr.gov.bd/), and updates the **সর্বশেষ হালনাগাদের তারিখ** column with the latest date from the website.

## What it does

1. Reads numbered rows from the document (for example: `১.১`, `১.২`, `১০.১.১`)
2. Extracts the title after the number (for example: `ইতিহাস ও কার্যাবলী`)
3. Finds the matching link on the MOWR homepage using the anchor `title` attribute
4. Opens that page and reads the date from:

   `div.content-update-block > p`

   Example text:

   `কনটেন্টটি শেষ হাল-নাগাদ করা হয়েছে: রবিবার, ২৬ জুলাই, ২০২৬ এ ১১:৪২ AM`

5. Saves the date in Bengali `dd/mm/yyyy` format (for example: `২৬/০৭/২০২৬`) into the document table
6. Formats the entire **সর্বশেষ হালনাগাদের তারিখ** column as **Nikosh**, **14 pt**, and **centered** — including rows that keep their existing date

## Requirements

- Python 3.8 or newer
- Internet connection

## Setup

Open PowerShell or Command Prompt in this folder and run:

```powershell
cd "e:\ওয়েবসাইটের তথ্য হালনাগাদকরণ প্রতিবেদন"
pip install -r requirements.txt
```

## How to run

### 1. Test first (recommended)

This checks a few rows without changing the document:

```powershell
python update_mowr_dates.py --dry-run --limit 5
```

### 2. Update the original document

Close the `.docx` file in Word or Cursor first, then run:

```powershell
python update_mowr_dates.py
```

The script will:

- Create a backup file: `ওয়েবসাইট_হালনাগাদ_তথ্য_আপডেটেড_backup.docx`
- Update: `ওয়েবসাইট_হালনাগাদ_তথ্য_আপডেটেড.docx`

### 3. Save to a new file

Use this if the original file is still open:

```powershell
python update_mowr_dates.py --output updated_output.docx
```

## All command options

| Option | Description |
|--------|-------------|
| `--doc FILE` | Input Word file. Default: `ওয়েবসাইট_হালনাগাদ_তথ্য_আপডেটেড.docx` |
| `--output FILE` | Output Word file. If not given, the input file is overwritten |
| `--dry-run` | Show results only. Do not save changes |
| `--limit N` | Process only the first `N` numbered topics |
| `--delay SECONDS` | Wait time between website requests. Default: `0.4` |

### Examples

```powershell
# Use a different input file
python update_mowr_dates.py --doc "my_report.docx"

# Save result to another file
python update_mowr_dates.py --output "report_updated.docx"

# Test first 10 rows only
python update_mowr_dates.py --dry-run --limit 10

# Full run with slower requests
python update_mowr_dates.py --delay 1
```

## Files in this folder

| File | Purpose |
|------|---------|
| `update_mowr_dates.py` | Main script |
| `requirements.txt` | Python packages |
| `ওয়েবসাইট_হালনাগাদ_তথ্য_আপডেটেড.docx` | Source report |
| `ওয়েবসাইট_হালনাগাদ_তথ্য_আপডেটেড_backup.docx` | Automatic backup |

## Notes

- If the original `.docx` is open, saving may fail. Close it first, or use `--output`.
- Some rows may not update because:
  - The page has no `content-update-block`
  - The link points to an external website
  - No matching menu item was found on mowr.gov.bd
  - The website did not respond in time
- The script prints progress in the terminal and lists any rows that could not be updated.

## Troubleshooting

**`Permission denied` when saving**

Close the Word file and run again, or use:

```powershell
python update_mowr_dates.py --output updated_output.docx
```

**`ModuleNotFoundError`**

Install dependencies again:

```powershell
pip install -r requirements.txt
```

**SSL or connection errors**

Make sure you are connected to the internet and try again later.
