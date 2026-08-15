from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class OCRLine:
    """一条 OCR 文本及其可选的几何和识别置信度。"""

    text: str
    score: float = 0.0
    box: tuple[tuple[float, float], ...] | None = None
    variant: str = ""

    @property
    def center_y(self) -> float | None:
        if not self.box:
            return None
        return sum(point[1] for point in self.box) / len(self.box)

    @property
    def left_x(self) -> float | None:
        if not self.box:
            return None
        return min(point[0] for point in self.box)


@dataclass(frozen=True)
class ChecksumReport:
    passport_number: bool
    date_of_birth: bool
    expiry_date: bool
    personal_number: bool
    composite: bool

    @property
    def all_valid(self) -> bool:
        return all(
            (
                self.passport_number,
                self.date_of_birth,
                self.expiry_date,
                self.personal_number,
                self.composite,
            )
        )

    def as_dict(self) -> dict[str, bool]:
        return {
            "passport_number": self.passport_number,
            "date_of_birth": self.date_of_birth,
            "expiry_date": self.expiry_date,
            "personal_number": self.personal_number,
            "composite": self.composite,
        }


@dataclass
class MRZResult:
    """新 MRZ 引擎的结构化结果。

    required_fields() 返回用户约定的核心字段；诊断字段用于解释拒绝或纠错原因。
    """

    passport_number: str = ""
    surname: str = ""
    given_names: str = ""
    nationality: str = ""
    date_of_birth: str = ""
    sex: str = ""
    expiry_date: str = ""
    personal_number: str = ""
    mrz_line_1: str = ""
    mrz_line_2: str = ""
    confidence: float = 0.0
    checksum_valid: bool = False
    checksum_details: dict[str, bool] = field(default_factory=dict)
    corrections: list[str] = field(default_factory=list)
    ocr_confidence: float = 0.0
    consensus_count: int = 1
    preprocess_variant: str = ""
    accepted: bool = False
    rejection_reason: str = ""

    def required_fields(self) -> dict[str, Any]:
        return {
            "passport_number": self.passport_number,
            "surname": self.surname,
            "given_names": self.given_names,
            "nationality": self.nationality,
            "date_of_birth": self.date_of_birth,
            "sex": self.sex,
            "expiry_date": self.expiry_date,
            "personal_number": self.personal_number,
            "mrz_line_1": self.mrz_line_1,
            "mrz_line_2": self.mrz_line_2,
            "confidence": round(float(self.confidence), 2),
            "checksum_valid": bool(self.checksum_valid),
        }

    def as_dict(self) -> dict[str, Any]:
        data = self.required_fields()
        data.update(
            {
                "checksum_details": dict(self.checksum_details),
                "corrections": list(self.corrections),
                "ocr_confidence": round(float(self.ocr_confidence), 4),
                "consensus_count": self.consensus_count,
                "preprocess_variant": self.preprocess_variant,
                "accepted": self.accepted,
                "rejection_reason": self.rejection_reason,
            }
        )
        return data
