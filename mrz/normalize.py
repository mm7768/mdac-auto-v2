from __future__ import annotations

import re
import unicodedata

MRZ_CHARSET = set("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<")

_SEPARATOR_MAP = str.maketrans(
    {
        "‹": "<",
        "›": "<",
        "«": "<",
        "»": "<",
        "﹤": "<",
        "＜": "<",
        "人": "<",
        "|": "<",
        "¦": "<",
    }
)


def normalize_text(text: str) -> str:
    """Normalize without applying field-specific O/0-style substitutions."""
    if not text:
        return ""
    normalized = unicodedata.normalize("NFKC", str(text)).upper().translate(_SEPARATOR_MAP)
    normalized = re.sub(r"\s+", "", normalized)
    normalized = normalized.replace("—", "-").replace("–", "-")
    return normalized


def valid_mrz_characters(text: str) -> bool:
    return bool(text) and all(character in MRZ_CHARSET for character in text)


def split_ocr_text(text: str) -> list[str]:
    normalized = unicodedata.normalize("NFKC", str(text or "")).upper().translate(_SEPARATOR_MAP)
    lines = []
    for raw_line in re.split(r"[\r\n]+", normalized):
        line = normalize_text(raw_line)
        if line:
            lines.append(line)
    return lines


def clean_candidate_line(text: str) -> str:
    """Keep only MRZ-like characters while preserving unknown chars for rejection."""
    normalized = normalize_text(text)
    # OCR engines sometimes emit a literal apostrophe around a line; these are not
    # MRZ data and can be safely removed before strict character validation.
    normalized = normalized.strip("'`\"")
    return normalized
