"""Parser tests on synthetic Form 5 text (fake names only -- never real student data)."""

from form5.parsing import (
    extract_degree,
    extract_registration_status,
    extract_scholarship,
    extract_student_name,
    extract_student_number,
)

BODY = (
    "UP FORM 5. CERTIFICATE OF REGISTRATION\n"
    "STUDENT NO. 202399999\n"
    "NAME: DELA CRUZ, JUAN SANTOS\n"
    "COLLEGE PROGRAM TERM & SY\nCEAT BSCE 1st 2026-2027\n"
    "I agree ... to avail Free Tuition and Other School Fees\n"
)


def form(watermark="", scholarship=""):
    return (f"{watermark}\n{BODY}"
            f"SCHOLARSHIP /\nPRIVILEGES\n{scholarship}\nChange of Matriculation\n")


def test_watermark_officially_registered():
    text = form("O F F I C I A L L Y  R E G I S T E R E D")
    assert extract_registration_status(text) == "Officially Registered"


def test_watermark_billing():
    assert extract_registration_status(form("B I L L I N G")) == "Billing"


def test_watermark_missing():
    assert extract_registration_status(form()) == "Unknown"


def test_student_number_name_degree():
    text = form()
    assert extract_student_number(text) == "202399999"
    assert extract_student_name(text) == "DELA CRUZ, JUAN SANTOS"
    assert extract_degree(text) == "BSCE"


def test_degree_falls_back_to_row_text():
    assert extract_degree("no program here", "DELA CRUZ JUAN BSCS 2023") == "BSCS"


def test_scholarship_ra10931():
    assert extract_scholarship(form(scholarship="RA 10931 FREE\nTUITION")) == "RA 10931 - Free Tuition"


def test_scholarship_pd80():
    assert extract_scholarship(form(scholarship="PD80")) == "PD 80"


def test_scholarship_fds_before_fd():
    assert extract_scholarship(form(scholarship="FDS")) == "FDS"


def test_empty_box_is_ne_despite_free_tuition_sentence():
    assert extract_scholarship(form()) == "NE"


def test_unknown_code_kept_raw():
    assert extract_scholarship(form(scholarship="XYZ 12")) == "XYZ 12"
