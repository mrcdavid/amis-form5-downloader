from dataclasses import dataclass


@dataclass
class Form5Info:
    name: str | None
    student_number: str | None
    degree: str | None
    registration_status: str
    scholarship: str
