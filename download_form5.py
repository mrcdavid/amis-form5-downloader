import asyncio
import re
from pathlib import Path

from playwright.async_api import async_playwright
from pypdf import PdfReader


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

# Your current term
TERM_ID = "1261"


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_filename(name: str) -> str:
    """
    Make the student name safe for a Windows filename.
    """

    # Windows-invalid characters
    name = re.sub(r'[<>:"/\\|?*]', '', name)

    # Remove excessive whitespace
    name = re.sub(r'\s+', ' ', name)

    return name.strip()


def extract_student_name(pdf_path: Path) -> str | None:
    """
    Extract the student's name from the downloaded PDF.
    """

    try:

        reader = PdfReader(str(pdf_path))

        full_text = ""

        for page in reader.pages:

            text = page.extract_text()

            if text:
                full_text += text + "\n"

        print("\nPDF TEXT:")
        print("-" * 60)
        print(full_text[:3000])
        print("-" * 60)

        # ----------------------------------------------------
        # Try several common name formats
        # ----------------------------------------------------

        patterns = [

            r'Student Name\s*:\s*(.+)',
            r'Name of Student\s*:\s*(.+)',
            r'Name\s*:\s*(.+)',
            r'Student\s*:\s*(.+)',

        ]

        for pattern in patterns:

            match = re.search(
                pattern,
                full_text,
                re.IGNORECASE
            )

            if match:

                name = match.group(1)

                # Only take the first line
                name = name.split("\n")[0]

                name = clean_filename(name)

                if name:
                    return name

        return None

    except Exception as e:

        print(
            f"ERROR extracting PDF text: {e}"
        )

        return None


def unique_filename(folder: Path, name: str) -> Path:
    """
    Prevent overwriting existing files.
    """

    path = folder / f"{name}.pdf"

    counter = 2

    while path.exists():

        path = folder / f"{name} ({counter}).pdf"

        counter += 1

    return path


# ============================================================
# MAIN
# ============================================================

async def main():

    async with async_playwright() as p:

        print("=" * 70)
        print("AMIS FORM 5 DOWNLOADER")
        print("=" * 70)

        # ----------------------------------------------------
        # Launch browser
        # ----------------------------------------------------

        #browser = await p.chromium.launch(
        #    headless=False
        #)
        browser = await p.chromium.launch(
            headless=False,
            channel="msedge"
        )

        context = await browser.new_context(
            accept_downloads=True
        )

        page = await context.new_page()

        # ----------------------------------------------------
        # Open AMIS
        # ----------------------------------------------------

        print("\nOpening AMIS...")

        await page.goto(
            URL,
            wait_until="domcontentloaded"
        )

        print("AMIS page opened.")

        # ----------------------------------------------------
        # Allow manual login
        # ----------------------------------------------------

        print("\nIf AMIS requires login, log in now.")

        input(
            "Press ENTER after the student list is visible..."
        )

        # ----------------------------------------------------
        # Find View buttons
        # ----------------------------------------------------

        # We search for links/buttons containing "View".
        #
        # This is intentionally broad because the exact
        # HTML structure of the AMIS page may differ.
        # ----------------------------------------------------

        view_buttons = page.get_by_text(
            "View",
            exact=True
        )

        count = await view_buttons.count()

        print(
            f"\nFound {count} View buttons."
        )

        if count == 0:

            print(
                "\nNo View buttons were found."
            )

            print(
                "Inspect the page HTML and adjust the "
                "View selector."
            )

            await browser.close()

            return

        # ----------------------------------------------------
        # Process each student
        # ----------------------------------------------------

        for index in range(count):

            print("\n")
            print("=" * 70)
            print(
                f"PROCESSING STUDENT "
                f"{index + 1} / {count}"
            )
            print("=" * 70)

            # Re-query because the page may change
            view_buttons = page.get_by_text(
                "View",
                exact=True
            )

            if index >= await view_buttons.count():

                print(
                    "No more View buttons."
                )

                break

            button = view_buttons.nth(index)

            # ------------------------------------------------
            # Scroll to button
            # ------------------------------------------------

            await button.scroll_into_view_if_needed()

            # ------------------------------------------------
            # Determine the student's row
            #
            # This is useful later for obtaining the name.
            # ------------------------------------------------

            try:

                row = button.locator(
                    "xpath=ancestor::tr"
                )

                if await row.count():

                    row_text = await row.inner_text()

                    print(
                        f"\nStudent row:\n{row_text}"
                    )

                else:

                    row_text = ""

            except:

                row_text = ""

            # ------------------------------------------------
            # Intercept the form5 request
            # ------------------------------------------------

            form5_response = None

            async def handle_response(response):

                nonlocal form5_response

                if "form5.php" in response.url:

                    form5_response = response

                    print(
                        "\nFORM 5 DETECTED:"
                    )

                    print(response.url)

            page.on(
                "response",
                handle_response
            )

            # ------------------------------------------------
            # Click View
            # ------------------------------------------------

            print(
                "\nClicking View..."
            )

            try:

                await button.click(
                    timeout=10000
                )

            except Exception as e:

                print(
                    f"Could not click View: {e}"
                )

                page.remove_listener(
                    "response",
                    handle_response
                )

                continue

            # ------------------------------------------------
            # Wait for form5.php
            # ------------------------------------------------

            for _ in range(20):

                if form5_response:

                    break

                await page.wait_for_timeout(
                    500
                )

            # ------------------------------------------------
            # Remove response listener
            # ------------------------------------------------

            page.remove_listener(
                "response",
                handle_response
            )

            # ------------------------------------------------
            # Check result
            # ------------------------------------------------

            if not form5_response:

                print(
                    "\nWARNING:"
                )

                print(
                    "form5.php was not detected."
                )

                continue

            # ------------------------------------------------
            # Get PDF bytes
            # ------------------------------------------------

            try:

                pdf_bytes = await form5_response.body()

            except Exception as e:

                print(
                    f"Could not read PDF: {e}"
                )

                continue

            # ------------------------------------------------
            # Save temporary PDF
            # ------------------------------------------------

            temporary_file = (
                OUTPUT_FOLDER
                / f"_temporary_{index + 1}.pdf"
            )

            temporary_file.write_bytes(
                pdf_bytes
            )

            print(
                f"\nTemporary PDF saved:"
            )

            print(
                temporary_file
            )

            # ------------------------------------------------
            # Extract name from PDF
            # ------------------------------------------------

            student_name = extract_student_name(
                temporary_file
            )

            # ------------------------------------------------
            # If PDF extraction fails, attempt to obtain
            # name from the table row.
            # ------------------------------------------------

            if not student_name and row_text:

                print(
                    "\nPDF name extraction failed."
                )

                print(
                    "You can customize row parsing here."
                )

            # ------------------------------------------------
            # Rename
            # ------------------------------------------------

            if student_name:

                final_file = unique_filename(
                    OUTPUT_FOLDER,
                    student_name
                )

                temporary_file.rename(
                    final_file
                )

                print(
                    "\nSUCCESS!"
                )

                print(
                    f"Student: {student_name}"
                )

                print(
                    f"Saved: {final_file}"
                )

            else:

                unknown_file = (
                    OUTPUT_FOLDER
                    / f"UNKNOWN_{index + 1}.pdf"
                )

                temporary_file.rename(
                    unknown_file
                )

                print(
                    "\nWARNING:"
                )

                print(
                    "Could not determine student name."
                )

                print(
                    f"Saved as: {unknown_file}"
                )

            # ------------------------------------------------
            # Wait before next student
            # ------------------------------------------------

            await page.wait_for_timeout(
                500
            )

        # ----------------------------------------------------
        # Finished
        # ----------------------------------------------------

        print("\n")
        print("=" * 70)
        print("FINISHED")
        print("=" * 70)

        print(
            f"\nPDF files are in:\n"
            f"{OUTPUT_FOLDER.absolute()}"
        )

        input(
            "\nPress ENTER to close the browser..."
        )

        await browser.close()


# ============================================================
# START
# ============================================================

if __name__ == "__main__":

    asyncio.run(main())