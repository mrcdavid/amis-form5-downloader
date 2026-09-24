import re


def squeeze(text: str) -> str:
    """Strip ALL whitespace and uppercase, so letter-spaced stamps match."""
    return re.sub(r"\s+", "", text or "").upper()


def clean_filename(name: str) -> str:
    """Make the student name safe for a Windows filename."""
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    name = re.sub(r"\s+", " ", name)
    return name.strip()


def normalize_text(s: str) -> str:
    """Lowercase + collapse whitespace, for fuzzy matching."""
    return re.sub(r"\s+", " ", s or "").strip().lower()
