"""Main download loop: walk the enrolled-student list and save each Form 5."""

from .browser import click_and_capture_form5, close_extra_tabs
from .config import CDP_ENDPOINT, EXCEL_PATH, INTER_STUDENT_DELAY, OUTPUT_FOLDER, URL
from .excel import append_excel_row, build_excel_existing_keys, excel_key, init_excel, save_excel
from .parsing import parse_form5, pdf_text
from .storage import build_existing_index, make_key, unique_filename
from .text_utils import normalize_text


async def download_all(debug: bool = False):
    from playwright.async_api import async_playwright

    async with async_playwright() as p:

        print("=" * 70)
        print("[+] FORM THUNDER // AMIS DOWNLOADER")
        print("=" * 70)

        # Connect to the already-open, already-logged-in Edge (see CLAUDE.md)
        browser = await p.chromium.connect_over_cdp(CDP_ENDPOINT)
        context = browser.contexts[0]
        page = context.pages[0]

        print("\nOpening AMIS...")
        await page.goto(URL, wait_until="domcontentloaded")
        print("\nIf AMIS requires login, log in now.")
        input("Press ENTER after the student list is visible...")

        existing_index = build_existing_index(OUTPUT_FOLDER)
        print(f"\nFound {len(existing_index['keys'])} existing Form 5 PDF(s) "
              f"in '{OUTPUT_FOLDER}'. These will be skipped.")

        excel_wb, excel_ws = init_excel(EXCEL_PATH)
        excel_existing_keys = build_excel_existing_keys(excel_ws)
        print(f"Tracking sheet: {EXCEL_PATH} ({len(excel_existing_keys)} row(s) already logged)")

        view_buttons = page.get_by_text("Form 5", exact=True)
        count = await view_buttons.count()
        print(f"\nFound {count} Form 5 buttons.")

        if count == 0:
            print("No Form 5 buttons were found. Inspect the page HTML and adjust the selector.")
            await browser.close()
            return

        skipped_early = skipped_duplicate = downloaded = failed = 0

        for index in range(count):
            print(f"\n--- STUDENT {index + 1} / {count} ---")

            view_buttons = page.get_by_text("Form 5", exact=True)
            if index >= await view_buttons.count():
                print("No more Form 5 buttons.")
                break

            button = view_buttons.nth(index)
            await button.scroll_into_view_if_needed()

            try:
                row = button.locator("xpath=ancestor::tr")
                row_text = await row.inner_text() if await row.count() else ""
            except Exception:
                row_text = ""

            # Pre-click duplicate check: skip without opening the PDF
            row_norm = normalize_text(row_text)
            matched = next((nm for nm in existing_index["names"] if nm in row_norm), None)
            if matched:
                print(f"SKIP (already downloaded): {matched}")
                skipped_early += 1
                continue

            form5_response = await click_and_capture_form5(page, context, button)

            pdf_bytes = b""
            if form5_response:
                try:
                    pdf_bytes = await form5_response.body()
                except Exception as e:
                    print(f"Could not read PDF: {e}")

            if not pdf_bytes.startswith(b"%PDF"):
                print("WARNING: form5.php was not detected or did not return a PDF.")
                failed += 1
                await close_extra_tabs(context, page)
                continue

            text = pdf_text(pdf_bytes)
            if debug:
                print("\nPDF TEXT:\n" + "-" * 60 + f"\n{text[:3000]}\n" + "-" * 60)

            info = parse_form5(text, row_text)
            print(f"Student No: {info.student_number or 'UNKNOWN'} | Degree: {info.degree or 'UNKNOWN'} | "
                  f"Status: {info.registration_status} | Scholarship: {info.scholarship}")

            if not info.name:
                unknown_file = unique_filename(OUTPUT_FOLDER, f"UNKNOWN_{index + 1}")
                unknown_file.write_bytes(pdf_bytes)
                print(f"WARNING: Could not determine student name. Saved as: {unknown_file}")
                failed += 1
                append_excel_row(excel_ws, info, unknown_file.name,
                                 "Name could not be parsed from PDF - please check manually")
                save_excel(excel_wb, EXCEL_PATH)

            elif make_key(info.degree, info.name) in existing_index["keys"]:
                print(f"SKIP (duplicate found after download): {info.name}")
                skipped_duplicate += 1

            else:
                filename = f"{info.degree}_{info.name}" if info.degree else info.name
                final_file = unique_filename(OUTPUT_FOLDER, filename)
                final_file.write_bytes(pdf_bytes)

                existing_index["keys"].add(make_key(info.degree, info.name))
                existing_index["names"].add(normalize_text(info.name))
                downloaded += 1
                print(f"SAVED: {final_file.name}")

                key = excel_key(info.student_number, info.name)
                if key not in excel_existing_keys:
                    append_excel_row(excel_ws, info, final_file.name)
                    save_excel(excel_wb, EXCEL_PATH)
                    excel_existing_keys.add(key)

            await close_extra_tabs(context, page)
            await page.wait_for_timeout(int(INTER_STUDENT_DELAY * 1000))

        save_excel(excel_wb, EXCEL_PATH)

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
