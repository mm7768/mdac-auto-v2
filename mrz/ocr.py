from __future__ import annotations

import json
from collections.abc import Iterable as IterableABC
from dataclasses import dataclass
from typing import Any, Protocol

import numpy as np

from .models import OCRLine


class OCRBackend(Protocol):
    def recognize(self, image: np.ndarray, variant: str = "") -> list[OCRLine]:
        ...


@dataclass
class PaddleOCRBackend:
    """Lazy PaddleOCR adapter supporting both predict() and legacy ocr()."""

    engine: Any = None

    def _get_engine(self) -> Any:
        if self.engine is not None:
            return self.engine
        try:
            from paddleocr import PaddleOCR
        except ImportError as exc:  # pragma: no cover - depends on deployment
            raise RuntimeError("未安装 PaddleOCR，请先安装 paddleocr 与 paddlepaddle") from exc

        try:
            self.engine = PaddleOCR(use_angle_cls=True, lang="en", enable_mkldnn=False)
        except (TypeError, ValueError):
            self.engine = PaddleOCR(lang="en", enable_mkldnn=False)
        return self.engine

    def recognize(self, image: np.ndarray, variant: str = "") -> list[OCRLine]:
        engine = self._get_engine()
        array = np.asarray(image)
        try:
            raw = engine.predict(array) if hasattr(engine, "predict") else engine.ocr(array, cls=True)
        except (AttributeError, TypeError):
            raw = engine.ocr(array, cls=True)
        return extract_ocr_lines(raw, variant=variant)


def _to_box(value: Any) -> tuple[tuple[float, float], ...] | None:
    if value is None:
        return None
    try:
        array = np.asarray(value, dtype=float)
        if array.size == 4:
            x1, y1, x2, y2 = array.reshape(-1).tolist()
            return ((x1, y1), (x2, y1), (x2, y2), (x1, y2))
        if array.ndim == 2 and array.shape[1] >= 2:
            return tuple((float(row[0]), float(row[1])) for row in array[:, :2])
    except (TypeError, ValueError):
        return None
    return None


def _as_float(value: Any, default: float = 0.5) -> float:
    try:
        value = float(value)
        return min(1.0, max(0.0, value))
    except (TypeError, ValueError):
        return default


def _dict_lines(item: dict[str, Any], variant: str) -> list[OCRLine]:
    texts = item.get("rec_texts")
    if texts is None:
        texts = item.get("texts")
    if texts is None:
        texts = item.get("text")
    if isinstance(texts, str):
        texts = [texts]
    if not isinstance(texts, (list, tuple, np.ndarray)):
        return []
    scores = item.get("rec_scores")
    if scores is None:
        scores = item.get("scores")
    if scores is None:
        scores = []
    boxes = item.get("rec_boxes")
    if boxes is None:
        boxes = item.get("boxes")
    if boxes is None:
        boxes = item.get("dt_polys")
    if boxes is None:
        boxes = []
    if not isinstance(scores, (list, tuple, np.ndarray)):
        scores = [scores] * len(texts)
    if not isinstance(boxes, (list, tuple, np.ndarray)):
        boxes = [None] * len(texts)
    output = []
    for index, text in enumerate(texts):
        if not isinstance(text, str) or not text.strip():
            continue
        score = scores[index] if index < len(scores) else 0.5
        box = boxes[index] if index < len(boxes) else None
        output.append(OCRLine(text=text.strip(), score=_as_float(score), box=_to_box(box), variant=variant))
    return output


def _object_to_dict(item: Any) -> dict[str, Any] | None:
    if isinstance(item, dict):
        return item
    data = getattr(item, "json", None)
    if callable(data):
        try:
            data = data()
        except Exception:
            data = None
    if isinstance(data, dict):
        return data
    if isinstance(data, str):
        try:
            parsed = json.loads(data)
            return parsed if isinstance(parsed, dict) else None
        except json.JSONDecodeError:
            return None
    return None


def _legacy_lines(item: Any, variant: str) -> list[OCRLine]:
    if not isinstance(item, (list, tuple)):
        return []
    output = []
    for entry in item:
        if not isinstance(entry, (list, tuple)) or len(entry) < 2:
            continue
        box = _to_box(entry[0])
        payload = entry[1]
        if isinstance(payload, (list, tuple)) and payload:
            text = payload[0]
            score = payload[1] if len(payload) > 1 else 0.5
        else:
            text, score = payload, 0.5
        if isinstance(text, str) and text.strip():
            output.append(OCRLine(text=text.strip(), score=_as_float(score), box=box, variant=variant))
    return output


def extract_ocr_lines(raw: Any, variant: str = "") -> list[OCRLine]:
    """Extract OCRLine values from PaddleOCR 2.x/3.x and test doubles."""
    if raw is None:
        return []
    if isinstance(raw, dict):
        direct = _dict_lines(raw, variant)
        if direct:
            return direct
        nested = []
        for value in raw.values():
            nested.extend(extract_ocr_lines(value, variant))
        return nested
    object_data = _object_to_dict(raw)
    if object_data is not None:
        return extract_ocr_lines(object_data, variant)
    if isinstance(raw, (list, tuple)):
        legacy = _legacy_lines(raw, variant)
        if legacy:
            return legacy
        output = []
        for item in raw:
            output.extend(extract_ocr_lines(item, variant))
        return output
    if isinstance(raw, IterableABC) and not isinstance(raw, (str, bytes, bytearray)):
        output = []
        for item in raw:
            output.extend(extract_ocr_lines(item, variant))
        return output
    return []
