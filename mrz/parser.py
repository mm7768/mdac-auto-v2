from __future__ import annotations

from datetime import date, datetime

from .checksum import validate_td3_checksums
from .models import MRZResult
from .normalize import MRZ_CHARSET, clean_candidate_line, valid_mrz_characters


class TD3Parser:
    """Strict parser for the two-line, 44-character TD3 passport MRZ."""

    def __init__(self, today: date | None = None):
        self.today = today or date.today()

    def parse(self, line_1: str, line_2: str) -> MRZResult | None:
        line_1 = clean_candidate_line(line_1)
        line_2 = clean_candidate_line(line_2)
        if len(line_1) != 44 or len(line_2) != 44:
            return None
        if not valid_mrz_characters(line_1) or not valid_mrz_characters(line_2):
            return None
        if line_1[0:2] != "P<":
            return None
        if not all(character in "ABCDEFGHIJKLMNOPQRSTUVWXYZ<" for character in line_1[2:5]):
            return None
        if not all(character in "ABCDEFGHIJKLMNOPQRSTUVWXYZ<" for character in line_2[10:13]):
            return None

        surname, given_names = self._parse_name(line_1[5:44])
        passport_number = line_2[0:9].rstrip("<")
        nationality = line_2[10:13]
        date_of_birth = self._parse_date(line_2[13:19], kind="birth")
        expiry_date = self._parse_date(line_2[21:27], kind="expiry")
        sex = line_2[20]
        personal_number = line_2[28:42].rstrip("<")
        checksum = validate_td3_checksums(line_2)

        result = MRZResult(
            passport_number=passport_number,
            surname=surname,
            given_names=given_names,
            nationality=nationality,
            date_of_birth=date_of_birth or "",
            sex=sex,
            expiry_date=expiry_date or "",
            personal_number=personal_number,
            mrz_line_1=line_1,
            mrz_line_2=line_2,
            checksum_valid=checksum.all_valid,
            checksum_details=checksum.as_dict(),
        )
        result.rejection_reason = self._rejection_reason(result)
        return result

    @staticmethod
    def _parse_name(value: str) -> tuple[str, str]:
        name_field = value.rstrip("<")
        if "<<" in name_field:
            surname_raw, given_raw = name_field.split("<<", 1)
        else:
            surname_raw, given_raw = name_field, ""
        surname = " ".join(surname_raw.replace("<", " ").split())
        given_names = " ".join(given_raw.replace("<", " ").split())
        return surname, given_names

    def _parse_date(self, value: str, kind: str) -> str | None:
        if len(value) != 6 or not value.isdigit():
            return None
        try:
            year = int(value[0:2])
            month = int(value[2:4])
            day = int(value[4:6])
            if kind == "birth":
                pivot = self.today.year % 100
                full_year = 2000 + year if year <= pivot else 1900 + year
            else:
                full_year = 2000 + year
            parsed = datetime(full_year, month, day).date()
        except ValueError:
            return None
        if kind == "birth" and parsed > self.today:
            return None
        if kind == "expiry" and parsed.year < 2000:
            return None
        return parsed.isoformat()

    @staticmethod
    def _rejection_reason(result: MRZResult) -> str:
        missing = []
        for field_name, value in (
            ("passport_number", result.passport_number),
            ("surname", result.surname),
            ("nationality", result.nationality),
            ("date_of_birth", result.date_of_birth),
            ("expiry_date", result.expiry_date),
            ("sex", result.sex if result.sex in {"M", "F"} else ""),
        ):
            if not value:
                missing.append(field_name)
        if missing:
            return "invalid_fields:" + ",".join(missing)
        if not result.checksum_valid:
            failed = [key for key, value in result.checksum_details.items() if not value]
            return "checksum_failed:" + ",".join(failed)
        return ""


def parse_td3(line_1: str, line_2: str, today: date | None = None) -> MRZResult | None:
    return TD3Parser(today=today).parse(line_1, line_2)
