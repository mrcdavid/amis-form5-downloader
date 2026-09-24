"""Playwright helpers for clicking Form 5 buttons and capturing the PDF."""

import asyncio

from .config import RESPONSE_TIMEOUT


async def close_extra_tabs(context, main_page):
    """Close all tabs except the main AMIS page. Re-check a few times
    because AMIS may open PDF tabs asynchronously."""
    for _ in range(5):
        await asyncio.sleep(0.3)
        extra_pages = [p for p in context.pages if p != main_page]
        if not extra_pages:
            return
        for extra_page in extra_pages:
            try:
                await extra_page.close()
            except Exception as e:
                print(f"  Could not close tab: {e}")


async def click_and_capture_form5(page, context, button):
    """Click a Form 5 button and return the form5.php response (or None).
    Works whether AMIS loads the PDF in the same tab or a new one."""
    form5_response = None
    form5_event = asyncio.Event()

    def handle_response(response):
        nonlocal form5_response
        if "form5.php" in response.url and not form5_event.is_set():
            form5_response = response
            form5_event.set()

    def handle_new_page(new_page):
        new_page.on("response", handle_response)

    page.on("response", handle_response)
    context.on("page", handle_new_page)
    try:
        await button.click(timeout=10000)
        await asyncio.wait_for(form5_event.wait(), timeout=RESPONSE_TIMEOUT)
    except asyncio.TimeoutError:
        pass
    except Exception as e:
        print(f"Could not click Form 5: {e}")
    finally:
        page.remove_listener("response", handle_response)
        context.remove_listener("page", handle_new_page)

    return form5_response
