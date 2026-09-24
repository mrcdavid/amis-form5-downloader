"""Pure Form 5 PDF parsing: text in, values out (no browser, no Excel)."""

import io
import re

from pypdf import PdfReader

from .config import (
    DEFAULT_REGISTRATION_LABEL,
    DEFAULT_SCHOLARSHIP_LABEL,
    SCHOLARSHIP_TAGS,
    WATERMARK_TAGS,
)
from .models import Form5Info
from .text_utils import clean_filename, squeeze


NAME_RE = re.compile(r"\bNAME\s*:\s*(.+)", re.IGNORECASE)

# Form 5 prints e.g. "STUDENT NO. 202310846" (no dash). Also accept
# the dashed "2023-10846" style just in case.
STUDENT_NO_RE = re.compile(r"STUDENT\s*NO\.?\s*:?\s*(\d{4}-?\d{5})", re.IGNORECASE)

# "COLLEGE PROGRAM TERM & SY\nCEAT BSCE" -> program is the 2nd token
PROGRAM_RE = re.compile(r"COLLEGE\s+PROGRAM\s+TERM\s*&\s*SY\s+(\S+)\s+(\S+)", re.IGNORECASE)

SCHOLARSHIP_BLOCK_RE = re.compile(
    r"SCHOLARSHIP\s*/\s*PRIVILEGES\s*(.*?)\s*(?:Change of Matriculation|Deposit Fee|Date Generated|$)",
    re.IGNORECASE | re.DOTALL,
)

# Fallback for the degree when the PDF layout is unexpected: look in
# the student-list row text. Case-sensitive so surnames don't match.
ROW_DEGREE_RE = re.compile(r"\b(?:BS|BA|MS|MA|PHD)[A-Z]{1,10}(?:_[A-Z0-9]+)?\b")


def pdf_text(source) -> str:
    """Extract all text from a PDF (a path or raw bytes)."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    try:
        reader = PdfReader(source)
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as e:
        print(f"ERROR extracting PDF text: {e}")
        return ""


def extract_student_name(text: str) -> str | None:
    match = NAME_RE.search(text)
    if not match:
        return None
    return clean_filename(match.group(1).split("\n")[0]) or None


def extract_student_number(text: str) -> str | None:
    match = STUDENT_NO_RE.search(text)
    return match.group(1) if match else None


def extract_degree(text: str, row_text: str = "") -> str | None:
    match = PROGRAM_RE.search(text)
    if match:
        return clean_filename(match.group(2))
    match = ROW_DEGREE_RE.search(row_text or "")
    return clean_filename(match.group(0)) if match else None


def extract_registration_status(text: str) -> str:
    """Read the watermark (OFFICIALLY REGISTERED / BILLING)."""
    # Normally the watermark is everything before the form header.
    header_pos = text.upper().find("UP FORM 5")
    if header_pos > 0:
        region = squeeze(text[:header_pos])
        for key, label in WATERMARK_TAGS:
            if key in region:
                return label

    # Fallback: find the letter-spaced stamp anywhere. Requiring
    # whitespace between every letter means ordinary words never match.
    for key, label in WATERMARK_TAGS:
        if re.search(r"\s+".join(key), text, re.IGNORECASE):
            return label

    return DEFAULT_REGISTRATION_LABEL


def extract_scholarship(text: str) -> str:
    """Read the value in the SCHOLARSHIP / PRIVILEGES box."""
    match = SCHOLARSHIP_BLOCK_RE.search(text)
    if not match:
        return DEFAULT_SCHOLARSHIP_LABEL

    raw = re.sub(r"\s+", " ", match.group(1)).strip()
    if not raw:
        return DEFAULT_SCHOLARSHIP_LABEL

    squeezed = squeeze(raw)
    for key, label in SCHOLARSHIP_TAGS:
        if key in squeezed:
            return label

    # Unknown code: keep the raw value (if it looks like a short code)
    # so it can be added to SCHOLARSHIP_TAGS later.
    return raw if len(raw) <= 40 else DEFAULT_SCHOLARSHIP_LABEL


def parse_form5(text: str, row_text: str = "") -> Form5Info:
    return Form5Info(
        name=extract_student_name(text),
        student_number=extract_student_number(text),
        degree=extract_degree(text, row_text),
        registration_status=extract_registration_status(text),
        scholarship=extract_scholarship(text),
    )
