"""Entry point kept for `python download_form5.py`. The code lives in the form5/ package."""

from form5.cli import main
from form5.parsing import parse_form5, pdf_text  # noqa: F401  (for quick one-liner tests)

if __name__ == "__main__":
    main()
