from __future__ import annotations

from dataclasses import dataclass
from itertools import combinations
from typing import Iterable

from .corrector import CorrectionCandidate, generate_correction_candidates
from .models import MRZResult, OCRLine
from .normalize import clean_candidate_line, normalize_text
from .ocr import OCRBackend, PaddleOCRBackend
from .parser import TD3Parser
from .preprocess import PreprocessedImage, generate_preprocessed_variants


@dataclass
class _ScoredCandidate:
    result: MRZResult
    ocr_score: float
    variant: str


class MRZEngine:
    """High-confidence TD3 MRZ OCR pipeline.

    The engine is intentionally independent from Excel, Telegram and Playwright.
    """

    def __init__(
        self,
        ocr_backend: OCRBackend | None = None,
        min_confidence: float = 90.0,
        max_variants: int = 8,
        max_correction_candidates: int = 128,
        max_correction_edits: int = 3,
    ):
        self.ocr_backend = ocr_backend or PaddleOCRBackend()
        self.parser = TD3Parser()
        self.min_confidence = float(min_confidence)
        self.max_variants = max(1, int(max_variants))
        self.max_correction_candidates = max(1, int(max_correction_candidates))
        self.max_correction_edits = max(0, int(max_correction_edits))

    def parse(self, image) -> MRZResult | None:
        variants = generate_preprocessed_variants(image, max_variants=self.max_variants)
        if not variants:
            return None

        raw_candidates: list[_ScoredCandidate] = []
        for variant in variants:
            lines = self.ocr_backend.recognize(variant.image, variant.name)
            for line_1, line_2, ocr_score in self._line_pairs(lines):
                raw_candidates.extend(self._parse_pair(line_1, line_2, ocr_score, variant))

        if not raw_candidates:
            return MRZResult(rejection_reason="no_td3_candidate")
        return self._rank_and_finalize(raw_candidates)

    def _parse_pair(
        self,
        line_1: str,
        line_2: str,
        ocr_score: float,
        variant: PreprocessedImage,
    ) -> list[_ScoredCandidate]:
        line_1 = self._fit_line(line_1)
        line_2 = self._fit_line(line_2)
        if line_1 is None or line_2 is None:
            return []

        result = self.parser.parse(line_1, line_2)
        if result is None:
            return []
        direct = _ScoredCandidate(result=result, ocr_score=ocr_score, variant=variant.name)
        if result.checksum_valid and not result.rejection_reason:
            return [direct]

        corrected: list[_ScoredCandidate] = [direct]
        corrections = generate_correction_candidates(
            line_1,
            line_2,
            max_candidates=self.max_correction_candidates,
            max_edits=self.max_correction_edits,
        )
        for candidate in corrections:
            repaired = self.parser.parse(candidate.line_1, candidate.line_2)
            if repaired is None or not repaired.checksum_valid or repaired.rejection_reason:
                continue
            repaired.corrections = list(candidate.corrections)
            repaired.preprocess_variant = variant.name
            corrected.append(_ScoredCandidate(result=repaired, ocr_score=ocr_score, variant=variant.name))
        return corrected

    @staticmethod
    def _fit_line(line: str) -> str | None:
        line = clean_candidate_line(line)
        if len(line) == 44:
            return line
        # OCR frequently drops only trailing filler characters. Controlled padding
        # is allowed only when most of a full TD3 line is already present.
        if 40 <= len(line) < 44:
            return line.ljust(44, "<")
        return None

    def _line_pairs(self, lines: list[OCRLine]) -> Iterable[tuple[str, str, float]]:
        grouped = self._group_lines(lines)
        if len(grouped) < 2:
            return []

        line_values = []
        for text, score in grouped:
            normalized = normalize_text(text)
            if len(normalized) >= 35:
                line_values.append((normalized, max(0.0, min(1.0, score))))
        if len(line_values) < 2:
            return []

        first_lines = [item for item in line_values if item[0].startswith("P") and "<" in item[0]]
        second_lines = [item for item in line_values if item not in first_lines and ("<" in item[0] or sum(char.isdigit() for char in item[0]) >= 6)]
        if not first_lines:
            first_lines = sorted(line_values, key=lambda item: ("P" in item[0], len(item[0])), reverse=True)[:3]
        if not second_lines:
            second_lines = line_values

        pairs = []
        for first, second in combinations(line_values, 2):
            candidates = ((first, second), (second, first))
            for left, right in candidates:
                if not (left[0].startswith("P") or left[0].startswith("P<")):
                    continue
                if left[0] == right[0]:
                    continue
                pairs.append((left[0], right[0], (left[1] + right[1]) / 2.0))
        if not pairs:
            for first in first_lines:
                for second in second_lines:
                    if first[0] != second[0]:
                        pairs.append((first[0], second[0], (first[1] + second[1]) / 2.0))
        return pairs[:8]

    @staticmethod
    def _group_lines(lines: list[OCRLine]) -> list[tuple[str, float]]:
        if not lines:
            return []
        expanded: list[OCRLine] = []
        for line in lines:
            pieces = [piece for piece in line.text.replace("\r", "\n").split("\n") if piece.strip()]
            if len(pieces) == 1:
                expanded.append(line)
            else:
                expanded.extend(OCRLine(text=piece, score=line.score, box=line.box, variant=line.variant) for piece in pieces)

        with_positions = [line for line in expanded if line.center_y is not None]
        if len(with_positions) < 2:
            return [(line.text, line.score) for line in expanded]

        median_height = 0.0
        heights = []
        for line in with_positions:
            if line.box:
                ys = [point[1] for point in line.box]
                heights.append(max(1.0, max(ys) - min(ys)))
        if heights:
            median_height = sorted(heights)[len(heights) // 2]
        # 同一行的文字框中心应落在一个字高内；使用过大的阈值会把
        # TD3 的上下两行合并成一行，导致后续无法解析。
        threshold = max(8.0, median_height * 0.9)

        clusters: list[list[OCRLine]] = []
        for line in sorted(with_positions, key=lambda item: item.center_y or 0.0):
            if not clusters:
                clusters.append([line])
                continue
            previous_y = sum(item.center_y or 0.0 for item in clusters[-1]) / len(clusters[-1])
            if abs((line.center_y or 0.0) - previous_y) <= threshold:
                clusters[-1].append(line)
            else:
                clusters.append([line])

        grouped = []
        for cluster in clusters:
            ordered = sorted(cluster, key=lambda item: item.left_x if item.left_x is not None else 0.0)
            grouped.append(("".join(item.text for item in ordered), sum(item.score for item in ordered) / len(ordered)))
        return grouped

    def _rank_and_finalize(self, candidates: list[_ScoredCandidate]) -> MRZResult:
        grouped: dict[tuple[str, str], list[_ScoredCandidate]] = {}
        for candidate in candidates:
            key = (candidate.result.mrz_line_1, candidate.result.mrz_line_2)
            grouped.setdefault(key, []).append(candidate)

        ranked: list[MRZResult] = []
        for values in grouped.values():
            best = max(values, key=lambda item: self._base_score(item.result, item.ocr_score))
            result = best.result
            result.consensus_count = len({value.variant for value in values})
            result.preprocess_variant = ",".join(sorted({value.variant for value in values}))
            result.ocr_confidence = max(value.ocr_score for value in values)
            result.confidence = self._final_score(result, result.ocr_confidence)
            result.accepted = bool(
                result.checksum_valid
                and not result.rejection_reason
                and result.confidence >= self.min_confidence
            )
            if not result.accepted and not result.rejection_reason:
                result.rejection_reason = f"confidence_below_threshold:{result.confidence:.2f}"
            ranked.append(result)

        ranked.sort(key=lambda item: (item.accepted, item.confidence, item.consensus_count), reverse=True)
        return ranked[0]

    @staticmethod
    def _base_score(result: MRZResult, ocr_score: float) -> float:
        structure = 20.0
        checksum = 35.0 * (sum(result.checksum_details.values()) / 5.0 if result.checksum_details else 0.0)
        ocr = 20.0 * max(0.0, min(1.0, ocr_score))
        fields = 10.0 if not result.rejection_reason.startswith("invalid_fields") else 0.0
        return structure + checksum + ocr + fields

    @classmethod
    def _final_score(cls, result: MRZResult, ocr_score: float) -> float:
        structure = 20.0
        checksum = 35.0 if result.checksum_valid else 35.0 * (
            sum(result.checksum_details.values()) / 5.0 if result.checksum_details else 0.0
        )
        ocr = 20.0 * max(0.0, min(1.0, ocr_score))
        consensus = 15.0 * min(1.0, result.consensus_count / 2.0)
        fields = 10.0 if not result.rejection_reason.startswith("invalid_fields") else 0.0
        return round(structure + checksum + ocr + consensus + fields, 2)
