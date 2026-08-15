from __future__ import annotations

from datetime import datetime

from .engine import MRZEngine


class MRZParser:
    """Backward-compatible facade for the existing main_console.py callers."""

    def __init__(self, engine: MRZEngine | None = None):
        self.engine = engine or MRZEngine()
        self.last_result = None
        self.last_error = ""

    def parse_image(self, img_bytes):
        try:
            result = self.engine.parse(img_bytes)
        except Exception as exc:
            self.last_result = None
            self.last_error = f"MRZ OCR 运行失败: {exc}"
            return False, self.last_error

        self.last_result = result
        if result is None:
            self.last_error = "未生成 MRZ 候选"
            return False, self.last_error
        if not result.accepted:
            checksum_detail = ",".join(
                key for key, value in result.checksum_details.items() if not value
            )
            detail = f"；失败校验: {checksum_detail}" if checksum_detail else ""
            self.last_error = (
                f"MRZ 未达到自动接收门槛（confidence={result.confidence:.2f}, "
                f"reason={result.rejection_reason or 'unknown'}{detail}）"
            )
            return False, self.last_error

        return True, self._legacy_result(result)

    @staticmethod
    def _legacy_result(result):
        date_of_birth = datetime.fromisoformat(result.date_of_birth)
        expiry_date = datetime.fromisoformat(result.expiry_date)
        sex_text = "男" if result.sex == "M" else "女"
        full_name = f"{result.surname} {result.given_names}".strip()
        legacy = {
            # Existing callers rely on these names.
            "name": full_name,
            "passport": result.passport_number,
            "nationality": result.nationality,
            "dob": date_of_birth,
            "sex_text": sex_text,
            "passport_exp": expiry_date,
            # Preserve the new structured result for diagnostics and future callers.
            "passport_number": result.passport_number,
            "surname": result.surname,
            "given_names": result.given_names,
            "date_of_birth": result.date_of_birth,
            "sex": result.sex,
            "expiry_date": result.expiry_date,
            "personal_number": result.personal_number,
            "mrz_line_1": result.mrz_line_1,
            "mrz_line_2": result.mrz_line_2,
            "confidence": round(result.confidence, 2),
            "checksum_valid": result.checksum_valid,
        }
        return legacy
