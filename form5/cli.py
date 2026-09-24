import argparse
import asyncio

from .downloader import download_all
from .rebuild import rebuild_excel


def main():
    parser = argparse.ArgumentParser(description="Download UPLB AMIS Form 5 PDFs and log them to Excel.")
    parser.add_argument("--rebuild-excel", action="store_true",
                        help="Re-parse the PDFs already in the output folder and rewrite the tracking sheet "
                             "(no browser needed).")
    parser.add_argument("--debug", action="store_true", help="Print the extracted PDF text for each student.")
    args = parser.parse_args()

    if args.rebuild_excel:
        rebuild_excel()
    else:
        asyncio.run(download_all(debug=args.debug))
