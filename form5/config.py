"""All tunable settings. Edit values here instead of inside functions."""

from pathlib import Path


URL = (
    "https://amis-rapid.uplb.edu.ph/"
    "rapid-tools/ocs/enrolled_student_summary_list.php"
    "?orderby=alast_name"
)

CDP_ENDPOINT = "http://127.0.0.1:9222"

OUTPUT_FOLDER = Path("downloaded_form5")
OUTPUT_FOLDER.mkdir(exist_ok=True)

EXCEL_PATH = OUTPUT_FOLDER / "form5_tracking.xlsx"
SHEET_NAME = "Form 5 Tracking"

# How long to wait for the form5.php response before giving up (seconds)
RESPONSE_TIMEOUT = 8.0

# Small pacing delay between students (seconds). Lower = faster,
# but too low can overload the AMIS server or trigger rate limiting.
INTER_STUDENT_DELAY = 0.15

# ------------------------------------------------------------
# Registration status = the watermark stamped across the PDF.
#
# The watermark is extracted as letter-spaced text, e.g.
# "O F F I C I A L L Y  R E G I S T E R E D", and appears before the
# "UP FORM 5." header. Keys below are written WITHOUT spaces because
# they are compared against whitespace-stripped ("squeezed") text.
# Checked top to bottom; first match wins.
# ------------------------------------------------------------
WATERMARK_TAGS = [
    ("OFFICIALLYREGISTERED", "Officially Registered"),
    ("BILLING", "Billing"),
]
DEFAULT_REGISTRATION_LABEL = "Unknown"

# ------------------------------------------------------------
# Scholarship / privilege = the value printed in the
# "SCHOLARSHIP / PRIVILEGES" box (bottom-right of the form),
# e.g. "RA 10931 FREE TUITION" or "PD80".
#
# Only that box is searched -- NOT the whole PDF, because every
# Form 5 contains the sentence "...to avail Free Tuition and Other
# School Fees", which would otherwise tag everyone as RA 10931.
#
# Keys are no-space, uppercase. Longer codes must come before
# shorter codes they contain (FDS before FD). If the box has a value
# that matches none of these, the raw value is written instead.
# ------------------------------------------------------------
SCHOLARSHIP_TAGS = [
    ("RA10931", "RA 10931 - Free Tuition"),
    ("FREETUITION", "RA 10931 - Free Tuition"),
    ("FDS", "FDS"),
    ("FD", "FD"),
    ("TFE", "TFE"),
    ("PD33", "PD 33"),
    ("PD60", "PD 60"),
    ("PD80", "PD 80"),
]
DEFAULT_SCHOLARSHIP_LABEL = "NE"


# Tracking sheet columns
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
EXCEL_WIDTHS = [16, 16, 30, 18, 22, 24, 34, 30]
