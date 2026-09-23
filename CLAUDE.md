# CLAUDE.md

Single-script tool (`download_form5.py`) that bulk-downloads UPLB AMIS **Form 5** (Certificate of Registration) PDFs from the enrolled-student list and logs each one to an Excel tracking sheet.

## How it works

1. Attaches via CDP (`http://127.0.0.1:9222`) to an Edge window the user already opened and logged into. The script never logs in itself.
2. Opens the enrolled-student list (`URL`), waits for the user to press ENTER, then clicks each `Form 5` button in turn.
3. Captures the `form5.php` response (same tab or popup) and parses the PDF text in memory with `pypdf`.
4. Saves the PDF as `downloaded_form5/<DEGREE>_<LAST, FIRST MIDDLE>.pdf` and adds a row to `downloaded_form5/form5_tracking.xlsx`.
5. Duplicates are skipped twice: **before clicking** (a name from an existing filename appears in the row text) and **after download** (degree+name key already exists). Excel rows are deduped by student number, falling back to name.

## Setup (Windows)

```
pip install playwright pypdf openpyxl
python -m playwright install chromium
```

Start Edge with remote debugging. Close all Edge windows first (`taskkill /F /IM msedge.exe`):

```
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="%USERPROFILE%\edge-amis"
```

If Edge is somewhere else, try `C:\Program Files\Microsoft\Edge\Application\msedge.exe` or `%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe`. To check it worked, open `http://127.0.0.1:9222/json/version`.

## Commands

```
python download_form5.py                  # download new Form 5s + log to Excel
python download_form5.py --debug          # also print each PDF's extracted text
python download_form5.py --rebuild-excel  # re-parse existing PDFs and rewrite the sheet (no browser)
```

Use `--rebuild-excel` after changing any parsing rule, because PDFs that are already downloaded are skipped and their rows will not update on their own. Notes typed into the sheet are kept (matched by PDF filename). Close the workbook in Excel first; if it's open, the script waits and asks you to close it.

## Form 5 PDF text layout (what the parser relies on)

Text as `pypdf` extracts it (verified on real forms):

| Field | Raw text | Extracted by |
|---|---|---|
| Registration status (watermark) | first line, letter-spaced: `O F F I C I A L L Y  R E G I S T E R E D` or `B I L L I N G` | `extract_registration_status` |
| Student number | `STUDENT NO. 202310846` (9 digits, no dash) | `STUDENT_NO_RE` |
| Name | `NAME: AALA, JOVIANNE XYRA DEDORO` | `NAME_RE` |
| Degree | `COLLEGE PROGRAM TERM & SY` followed by `CEAT BSCE` (2nd token) | `PROGRAM_RE`, falling back to the list-row text |
| Scholarship | `SCHOLARSHIP /\nPRIVILEGES\n<value>\nChange of Matriculation`, e.g. `RA 10931 FREE\nTUITION`, `PD80`, or empty | `extract_scholarship` |

Gotchas:
- **Do not search the whole PDF for scholarship codes.** Every Form 5 contains "…to avail Free Tuition and Other School Fees", so a whole-document search tags everyone as RA 10931. Only look inside the SCHOLARSHIP / PRIVILEGES box.
- The watermark is letter-spaced, so tag keys are compared against **whitespace-stripped, uppercased** text (`squeeze`). Keys in `WATERMARK_TAGS` / `SCHOLARSHIP_TAGS` must have **no spaces**.
- In `SCHOLARSHIP_TAGS`, list longer codes before shorter codes they contain (`FDS` before `FD`). Unknown box values are written as-is, and an empty box becomes `NE`.
- The `REGISTRATION STATUS:` label on the form is always blank. The watermark is the only source of that status.

## Conventions

- Configuration lives in the constants at the top of `download_form5.py` (URL, timeouts, tag lists). Tune those instead of hardcoding values inside functions.
- Parsing functions are pure (text in, value out), so you can test them without a browser:
  ```
  python -c "import download_form5 as d; print(d.parse_form5(d.pdf_text('downloaded_form5/<file>.pdf')))"
  ```
- The filename format `<DEGREE>_<NAME>.pdf` is also the dedup key. Changing it causes existing PDFs to be downloaded again.
- `downloaded_form5/` and the Edge profiles contain student personal data and are git-ignored. Never commit them.
