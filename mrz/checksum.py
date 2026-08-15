from __future__ import annotations

from .models import ChecksumReport

WEIGHTS = (7, 3, 1)


def character_value(character: str) -> int:
    """Return the ICAO MRZ value for one character."""
    if len(character) != 1:
        raise ValueError("MRZ character must have length 1")
    if character == "<":
        return 0
    if "0" <= character <= "9":
        return ord(character) - ord("0")
    if "A" <= character <= "Z":
        return ord(character) - ord("A") + 10
    raise ValueError(f"Invalid MRZ character: {character!r}")


def calculate_check_digit(value: str) -> str:
    total = 0
    for index, character in enumerate(value):
        total += character_value(character) * WEIGHTS[index % len(WEIGHTS)]
    return str(total % 10)


def is_valid_check_digit(value: str, check_digit: str) -> bool:
    return (
        len(check_digit) == 1
        and check_digit.isdigit()
        and all(character in "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<" for character in value)
        and calculate_check_digit(value) == check_digit
    )


def validate_td3_checksums(line_2: str) -> ChecksumReport:
    """Validate the five TD3 checksums on a normalized 44-char second line."""
    if len(line_2) != 44:
        return ChecksumReport(False, False, False, False, False)

    passport_number = is_valid_check_digit(line_2[0:9], line_2[9])
    date_of_birth = is_valid_check_digit(line_2[13:19], line_2[19])
    expiry_date = is_valid_check_digit(line_2[21:27], line_2[27])
    personal_number = is_valid_check_digit(line_2[28:42], line_2[42])

    composite_input = line_2[0:10] + line_2[13:20] + line_2[21:43]
    composite = is_valid_check_digit(composite_input, line_2[43])

    return ChecksumReport(
        passport_number=passport_number,
        date_of_birth=date_of_birth,
        expiry_date=expiry_date,
        personal_number=personal_number,
        composite=composite,
    )
