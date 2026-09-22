import asyncio
import re
from datetime import datetime
from pathlib import Path

from playwright.async_api import async_playwright
from pypdf import PdfReader
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURATION
# ============================================================

URL = (
    "https://amis-rapid.uplb.edu.ph/"
    "rapid-tools/ocs/enrolled_student_summary_list.php"
    "?orderby=alast_name"
)

OUTPUT_FOLDER = Path("downloaded_form5")
OUTPUT_FOLDER.mkdir(exist_ok=True)

EXCEL_PATH = OUTPUT_FOLDER / "form5_tracking.xlsx"
SHEET_NAME = "Form 5 Tracking"

# Your current term
TERM_ID = "1261"

# How long to wait for the form5.php response before giving up (seconds)
RESPONSE_TIMEOUT = 8.0

# Small pacing delay between students (seconds). Lower = faster,
# but too low can overload the AMIS server or trigger rate limiting.
INTER_STUDENT_DELAY = 0.15

# ------------------------------------------------------------
# Scholarship / privilege tags, and registration status tags.
#
# Form 5 stamps some headers (e.g. "OFFICIALLY REGISTERED") using
# individually-spaced letters -- "O F F I C I A L L Y  R E G I S T E R E D" --
# which breaks normal regex matching. To handle this, we compare against
# a "squeezed" version of the PDF text with ALL whitespace stripped out
# and uppercased, so spacing style never matters. That means each tag
# below is written as a plain no-space substring (not a regex).
#
# Each entry is (no-space substring to look for, friendly label written
# to Excel). Checked top to bottom; first match wins.
#
# !!! ADJUST THESE to match the exact wording your Form 5 uses !!!
# The scholarship/privilege block sits bottom-right on the form --
# paste that section's raw text (printed under "PDF TEXT:" when the
# script runs) and I can tighten this list to your real codes.
# ------------------------------------------------------------
SCHOLARSHIP_TAGS = [
    ("RA10931", "RA 10931 - Free Tuition"),
    ("FREETUITION", "RA 10931 - Free Tuition"),
    ("PD577", "PD 577"),
    ("PD1567", "PD 1567"),
    ("PD80", "PD 80"),
    ("STFAP", "STFAP"),
    ("SCHOLAR", "Scholar (unspecified)"),
]
DEFAULT_SCHOLARSHIP_LABEL = "Without"

REGISTRATION_TAGS = [
    ("NOTOFFICIALLYREGISTERED", "Not Registered"),
    ("OFFICIALLYREGISTERED", "Officially Registered"),
    ("FORBILLING", "For Billing"),
    ("BILLINGSTATEMENT", "For Billing"),
]
DEFAULT_REGISTRATION_LABEL = "Unknown"


def squeeze(text: str) -> str:
    """Strip ALL whitespace and uppercase -- makes matching immune to
    letter-spaced headers like 'O F F I C I A L L Y  R E G I S T E R E D'."""
    return re.sub(r'\s+', '', text or '').upper()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_filename(name: str) -> str:
    """Make the student name safe for a Windows filename."""
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    name = re.sub(r'\s+', ' ', name)
    return name.strip()


def normalize_text(s: str) -> str:
    """Lowercase + collapse whitespace, for fuzzy matching."""
    return re.sub(r'\s+', ' ', s or "").strip().lower()


DEGREE_PATTERNS = [
    r'\bBS[A-Z]{2,10}_[A-Z0-9]+\b',
    r'\bBS[A-Z]{2,10}\b',
    r'\bBA[A-Z]{2,10}_[A-Z0-9]+\b',
    r'\bBA[A-Z]{2,10}\b',
    r'\bB[A-Z]{2,10}_[A-Z0-9]+\b',
    r'\bB[A-Z]{2,10}\b',
    r'\bMS[A-Z]{2,10}_[A-Z0-9]+\b',
    r'\bMS[A-Z]{2,10}\b',
    r'\bMA[A-Z]{2,10}_[A-Z0-9]+\b',
    r'\bMA[A-Z]{2,10}\b',
]

NAME_PATTERNS = [
    r'Student Name\s*:\s*(.+)',
    r'Name of Student\s*:\s*(.+)',
    r'Name\s*:\s*(.+)',
    r'Student\s*:\s*(.+)',
]

# UP student numbers are typically "YYYY-NNNNN" (e.g. 2021-01234).
# !!! ADJUST if your student numbers look different !!!
STUDENT_NO_PATTERNS = [
    r'Student\s*No\.?\s*:?\s*(\d{4}-\d{4,6})',
    r'Student\s*Number\s*:?\s*(\d{4}-\d{4,6})',
    r'\b(\d{4}-\d{5})\b',
]


def extract_degree(row_text: str) -> str | None:
    """Pull a degree/program code out of a table row's text."""
    if not row_text:
        return None

    text = re.sub(r'\s+', ' ', row_text).strip()

    for pattern in DEGREE_PATTERNS:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return clean_filename(match.group(0))

    return None


def get_pdf_text(pdf_path: Path) -> str:
    """Extract all text from a PDF (used for name, student no, tags)."""
    try:
        reader = PdfReader(str(pdf_path))
        full_text = "\n".join(page.extract_text() or "" for page in reader.pages)

        print("\nPDF TEXT:")
        print("-" * 60)
        print(full_text[:3000])
        print("-" * 60)

        return full_text

    except Exception as e:
        print(f"ERROR extracting PDF text: {e}")
        return ""


def extract_student_name(full_text: str) -> str | None:
    """Extract the student's name from the PDF's text."""
    for pattern in NAME_PATTERNS:
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            name = match.group(1).split("\n")[0]
            name = clean_filename(name)
            if name:
                return name
    return None


def extract_student_number(full_text: str) -> str | None:
    """Extract the student number from the PDF's text."""
    for pattern in STUDENT_NO_PATTERNS:
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
    return None


def extract_tagged_field(full_text: str, tag_list, default_label: str) -> str:
    """Generic first-match tag lookup (used for registration + scholarship).
    Matches against whitespace-stripped text so letter-spaced stamps like
    'O F F I C I A L L Y  R E G I S T E R E D' still match."""
    squeezed = squeeze(full_text)
    for substring, label in tag_list:
        if substring in squeezed:
            return label
    return default_label


def unique_filename(folder: Path, name: str) -> Path:
    """Prevent overwriting existing files (within-run collisions only)."""
    path = folder / f"{name}.pdf"
    counter = 2
    while path.exists():
        path = folder / f"{name} ({counter}).pdf"
        counter += 1
    return path


def make_key(degree: str | None, student_name: str) -> str:
    """Canonical dedup key for a student's Form 5."""
    base = f"{degree}_{student_name}" if degree else student_name
    return normalize_text(base)


def build_existing_index(folder: Path) -> dict:
    """
    Scan already-downloaded PDFs so we can:
      (a) skip re-downloading a student we already have (post-check), and
      (b) recognize a student's row BEFORE clicking, by checking whether
          their name (pulled from an existing filename) shows up in the
          row text (pre-check) -- this lets us skip the click entirely
          on repeat runs.
    """
    keys = set()
    names = set()

    for f in folder.glob("*.pdf"):
        if f.name.startswith("_temporary_") or f.name.startswith("UNKNOWN_"):
            continue

        stem = f.stem
        stem = re.sub(r'\s*\(\d+\)$', '', stem)  # strip " (2)" style suffix
        keys.add(stem.lower())

        parts = stem.split('_', 1)
        name_part = parts[1] if len(parts) == 2 else stem
        norm = normalize_text(name_part)
        if norm:
            names.add(norm)

    return {"keys": keys, "names": names}


# ============================================================
# EXCEL TRACKING SHEET
# ============================================================

EXCEL_HEADERS = [
    "Date Processed",
    "Student Number",
    "Student Name",
    "Degree/Program",
    "Registration Status",
    "Scholarship Privilege",
    "PDF Filename",
    "Notes",
]


def init_excel(path: Path):
    """Open the tracking workbook, creating it with a header row if new."""
    if path.exists():
        wb = openpyxl.load_workbook(path)
        if SHEET_NAME not in wb.sheetnames:
            ws = wb.create_sheet(SHEET_NAME)
            _write_excel_header(ws)
        else:
            ws = wb[SHEET_NAME]
        return wb, ws

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    _write_excel_header(ws)
    return wb, ws


def _write_excel_header(ws):
    header_fill = PatternFill(start_color="8D1436", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Arial", bold=True, color="FFFFFF")

    for col_idx, header in enumerate(EXCEL_HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(EXCEL_HEADERS))}1"

    widths = [16, 16, 30, 18, 20, 24, 34, 30]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def append_excel_row(ws, row_values: list):
    row_idx = ws.max_row + 1
    for col_idx, value in enumerate(row_values, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.font = Font(name="Arial")


def build_excel_existing_keys(ws) -> set:
    """Avoid duplicate Excel rows across runs (by student number, else name)."""
    keys = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        student_no, student_name = row[1], row[2]
        key = normalize_text(student_no) if student_no else normalize_text(student_name)
        if key:
            keys.add(key)
    return keys


async def close_extra_tabs(context, main_page):
    """
    Close all tabs except the main AMIS page.
    Re-check several times because AMIS may open PDF tabs asynchronously.
    """
    for attempt in range(5):
        await asyncio.sleep(0.3)

        pages = list(context.pages)

        extra_pages = [
            p for p in pages
            if p != main_page
        ]

        if not extra_pages:
            print("No extra tabs found.")
            return

        print(f"Found {len(extra_pages)} extra tab(s). Closing...")

        for extra_page in extra_pages:
            try:
                print(f"  Closing: {extra_page.url}")
                await extra_page.close()
            except Exception as e:
                print(f"  Could not close tab: {e}")

    print("Finished checking extra tabs.")

# ============================================================
# MAIN
# ============================================================

async def main():

    async with async_playwright() as p:

        print("=" * 70)
        print("[+] FORM THUNDER // AMIS DOWNLOADER (optimized)")
        print("[+] SYSTEM STATUS: ONLINE")
        print("=" * 70)

        # ----------------------------------------------------
        # Connect to your already-open, already-logged-in browser
        # ----------------------------------------------------

        browser = await p.chromium.connect_over_cdp("http://127.0.0.1:9222")
        context = browser.contexts[0]
        page = context.pages[0]

        # ----------------------------------------------------
        # Open AMIS
        # ----------------------------------------------------

        print("\nOpening AMIS...")
        await page.goto(URL, wait_until="domcontentloaded")
        print("AMIS page opened.")

        print("\nIf AMIS requires login, log in now.")
        input("Press ENTER after the student list is visible...")

        # ----------------------------------------------------
        # Build the duplicate index + open the tracking workbook
        # ----------------------------------------------------

        existing_index = build_existing_index(OUTPUT_FOLDER)
        print(
            f"\nFound {len(existing_index['keys'])} existing Form 5 PDF(s) "
            f"already in '{OUTPUT_FOLDER}'. These will be skipped."
        )

        excel_wb, excel_ws = init_excel(EXCEL_PATH)
        excel_existing_keys = build_excel_existing_keys(excel_ws)
        print(f"Tracking sheet: {EXCEL_PATH} ({len(excel_existing_keys)} row(s) already logged)")

        # ----------------------------------------------------
        # Find Form 5 buttons
        # ----------------------------------------------------

        view_buttons = page.get_by_text("Form 5", exact=True)
        count = await view_buttons.count()

        print(f"\nFound {count} Form 5 buttons.")

        if count == 0:
            print("\nNo Form 5 buttons were found.")
            print("Inspect the page HTML and adjust the Form 5 selector.")
            excel_wb.save(EXCEL_PATH)
            await browser.close()
            return

        skipped_early = 0
        skipped_duplicate = 0
        downloaded = 0
        failed = 0

        # ----------------------------------------------------
        # Process each student
        # ----------------------------------------------------

        for index in range(count):

            print("\n" + "=" * 70)
            print(f"PROCESSING STUDENT {index + 1} / {count}")
            print("=" * 70)

            view_buttons = page.get_by_text("Form 5", exact=True)
            current_count = await view_buttons.count()
            if index >= current_count:
                print("No more Form 5 buttons.")
                break

            button = view_buttons.nth(index)
            await button.scroll_into_view_if_needed()

            # ------------------------------------------------
            # Grab the row text (used for dedup and degree)
            # ------------------------------------------------

            try:
                row = button.locator("xpath=ancestor::tr")
                row_text = await row.inner_text() if await row.count() else ""
            except Exception:
                row_text = ""

            # ------------------------------------------------
            # PRE-CLICK DUPLICATE CHECK
            # ------------------------------------------------

            row_norm = normalize_text(row_text)
            matched_existing_name = next(
                (nm for nm in existing_index["names"] if nm and nm in row_norm),
                None,
            )

            if matched_existing_name:
                print(f"\nSKIP (already downloaded): row matches '{matched_existing_name}'")
                skipped_early += 1
                continue

            # ------------------------------------------------
            # Intercept the form5 response (event-based, not polling)
            # Also watch for the case where Form 5 opens in a new tab.
            # ------------------------------------------------

            form5_response = None
            form5_event = asyncio.Event()

            def handle_response(response):
                nonlocal form5_response
                if "form5.php" in response.url and not form5_event.is_set():
                    form5_response = response
                    form5_event.set()
                    print(f"\nFORM 5 DETECTED:\n{response.url}")

            def handle_new_page(new_page):
                new_page.on("response", handle_response)

            page.on("response", handle_response)
            context.on("page", handle_new_page)

            existing_pages = list(context.pages)

            print("\nClicking Form 5...")
            try:
                await button.click(timeout=10000)
            except Exception as e:
                print(f"Could not click Form 5: {e}")
                page.remove_listener("response", handle_response)
                context.remove_listener("page", handle_new_page)
                failed += 1
                continue

            try:
                await asyncio.wait_for(form5_event.wait(), timeout=RESPONSE_TIMEOUT)
            except asyncio.TimeoutError:
                pass

            page.remove_listener("response", handle_response)
            context.remove_listener("page", handle_new_page)

            if not form5_response:
                print("\nWARNING: form5.php was not detected.")
                failed += 1
                for opened_page in context.pages:
                    if p != opened_page:
                        try:
                            await p.close()
                        except Exception:
                            pass
                continue

            # ------------------------------------------------
            # Get PDF bytes
            # ------------------------------------------------

            try:
                pdf_bytes = await form5_response.body()
            except Exception as e:
                print(f"Could not read PDF: {e}")
                failed += 1
                continue

            temporary_file = OUTPUT_FOLDER / f"_temporary_{index + 1}.pdf"
            temporary_file.write_bytes(pdf_bytes)

            # ------------------------------------------------
            # Extract everything from the PDF text in one pass
            # ------------------------------------------------

            full_text = get_pdf_text(temporary_file)
            student_name = extract_student_name(full_text)
            degree = extract_degree(row_text)
            student_number = extract_student_number(full_text)
            registration_status = extract_tagged_field(
                full_text, REGISTRATION_TAGS, DEFAULT_REGISTRATION_LABEL
            )
            scholarship = extract_tagged_field(
                full_text, SCHOLARSHIP_TAGS, DEFAULT_SCHOLARSHIP_LABEL
            )

            print(f"Degree detected: {degree}" if degree else "Degree could not be detected.")
            print(f"Student No: {student_number or 'UNKNOWN'}")
            print(f"Registration status: {registration_status}")
            print(f"Scholarship privilege: {scholarship}")

            if not student_name:
                unknown_file = OUTPUT_FOLDER / f"UNKNOWN_{index + 1}.pdf"
                temporary_file.rename(unknown_file)
                print("\nWARNING: Could not determine student name.")
                print(f"Saved as: {unknown_file}")
                failed += 1

                append_excel_row(excel_ws, [
                    datetime.now().strftime("%Y-%m-%d %H:%M"),
                    student_number or "",
                    "UNKNOWN",
                    degree or "",
                    registration_status,
                    scholarship,
                    unknown_file.name,
                    "Name could not be parsed from PDF - please check manually",
                ])
                excel_wb.save(EXCEL_PATH)

            else:
                key = make_key(degree, student_name)

                if key in existing_index["keys"]:
                    temporary_file.unlink(missing_ok=True)
                    print(f"\nSKIP (duplicate found after download): {student_name}")
                    skipped_duplicate += 1

                else:
                    filename = f"{degree}_{student_name}" if degree else student_name
                    final_file = unique_filename(OUTPUT_FOLDER, filename)
                    temporary_file.rename(final_file)

                    existing_index["keys"].add(key)
                    existing_index["names"].add(normalize_text(student_name))

                    downloaded += 1
                    print("\nSUCCESS!")
                    print(f"Degree: {degree or 'UNKNOWN'}")
                    print(f"Student: {student_name}")
                    print(f"Saved: {final_file}")

                    excel_key = normalize_text(student_number) if student_number else normalize_text(student_name)
                    if excel_key not in excel_existing_keys:
                        append_excel_row(excel_ws, [
                            datetime.now().strftime("%Y-%m-%d %H:%M"),
                            student_number or "",
                            student_name,
                            degree or "",
                            registration_status,
                            scholarship,
                            final_file.name,
                            "",
                        ])
                        excel_wb.save(EXCEL_PATH)
                        excel_existing_keys.add(excel_key)

            # ------------------------------------------------
            # Close any PDF tab that Form 5 opened
            # ------------------------------------------------

            # print("\nChecking for extra PDF tabs...")

            # for opened_page in list(context.pages):
            #     # Never close our main AMIS page
            #     if opened_page == page:
            #         continue

            #     try:
            #         print(f"Closing extra tab: {opened_page.url}")
            #         await opened_page.close()
            #     except Exception as e:
            #         print(f"Could not close tab: {e}")

            # await page.wait_for_timeout(
            #     int(INTER_STUDENT_DELAY * 1000)
            # )

            # ------------------------------------------------
            # Close any PDF tabs that Form 5 opened
            # ------------------------------------------------

            await close_extra_tabs(context, page)

            await page.wait_for_timeout(
                int(INTER_STUDENT_DELAY * 1000)
            )

        excel_wb.save(EXCEL_PATH)
        

        print("\n" + "=" * 70)
        print("FINISHED")
        print("=" * 70)
        print(f"Downloaded new:        {downloaded}")
        print(f"Skipped (pre-click):   {skipped_early}")
        print(f"Skipped (post-check):  {skipped_duplicate}")
        print(f"Failed:                {failed}")
        print(f"\nPDF files are in:\n{OUTPUT_FOLDER.absolute()}")
        print(f"Tracking sheet:\n{EXCEL_PATH.absolute()}")

        input("\nPress ENTER to close the browser...")
        await browser.close()


# ============================================================
# START
# ============================================================

if __name__ == "__main__":
    asyncio.run(main())