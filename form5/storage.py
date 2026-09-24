"""Output files and duplicate detection."""

import re
from pathlib import Path

from .text_utils import normalize_text


def unique_filename(folder: Path, name: str) -> Path:
    """Prevent overwriting existing files."""
    path = folder / f"{name}.pdf"
    counter = 2
    while path.exists():
        path = folder / f"{name} ({counter}).pdf"
        counter += 1
    return path


def make_key(degree: str | None, student_name: str) -> str:
    """Canonical dedup key for a student's Form 5 (matches the filename)."""
    base = f"{degree}_{student_name}" if degree else student_name
    return normalize_text(base)


def build_existing_index(folder: Path) -> dict:
    """
    Scan already-downloaded PDFs so we can:
      (a) skip re-downloading a student we already have (post-check), and
      (b) recognize a student's row BEFORE clicking, by checking whether
          their name (pulled from an existing filename) shows up in the
          row text (pre-check).
    """
    keys = set()
    names = set()

    for f in folder.glob("*.pdf"):
        if f.name.startswith(("_temporary_", "UNKNOWN_")):
            continue

        stem = re.sub(r"\s*\(\d+\)$", "", f.stem)  # strip " (2)" style suffix
        keys.add(normalize_text(stem))

        parts = stem.split("_", 1)
        norm = normalize_text(parts[1] if len(parts) == 2 else stem)
        if norm:
            names.add(norm)

    return {"keys": keys, "names": names}
