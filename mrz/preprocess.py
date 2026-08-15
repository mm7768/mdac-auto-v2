from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Iterable

import numpy as np
from PIL import Image

try:
    import cv2
except ImportError:  # pragma: no cover - runtime fallback for minimal installations
    cv2 = None


@dataclass(frozen=True)
class PreprocessedImage:
    name: str
    image: np.ndarray
    roi: tuple[int, int, int, int]
    angle: float = 0.0


def decode_image(image: bytes | bytearray | BytesIO | Image.Image | np.ndarray) -> np.ndarray:
    if isinstance(image, Image.Image):
        return np.array(image.convert("RGB"))
    if isinstance(image, np.ndarray):
        if image.ndim == 2:
            return image
        if image.shape[-1] == 4:
            return image[..., :3]
        return image
    if hasattr(image, "getvalue"):
        image = image.getvalue()
    if not isinstance(image, (bytes, bytearray)):
        raise TypeError("image must be bytes, BytesIO, PIL.Image, or numpy.ndarray")

    raw = np.frombuffer(image, dtype=np.uint8)
    if cv2 is not None:
        decoded = cv2.imdecode(raw, cv2.IMREAD_COLOR)
        if decoded is not None:
            return cv2.cvtColor(decoded, cv2.COLOR_BGR2RGB)
    with Image.open(BytesIO(bytes(image))) as pil_image:
        return np.array(pil_image.convert("RGB"))


def _gray(image: np.ndarray) -> np.ndarray:
    if image.ndim == 2:
        return image
    if cv2 is not None:
        return cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    return np.asarray(Image.fromarray(image).convert("L"))


def _resize(image: np.ndarray, scale: float) -> np.ndarray:
    if scale == 1.0:
        return image.copy()
    height, width = image.shape[:2]
    new_size = (max(1, int(round(width * scale))), max(1, int(round(height * scale))))
    if cv2 is not None:
        interpolation = cv2.INTER_CUBIC if scale >= 1 else cv2.INTER_AREA
        return cv2.resize(image, new_size, interpolation=interpolation)
    pil_image = Image.fromarray(image)
    return np.array(pil_image.resize(new_size, Image.Resampling.LANCZOS))


def _contrast(gray: np.ndarray) -> np.ndarray:
    if cv2 is not None:
        return cv2.normalize(gray, None, 0, 255, cv2.NORM_MINMAX)
    low, high = float(gray.min()), float(gray.max())
    if high <= low:
        return gray.copy()
    return ((gray.astype(np.float32) - low) * 255.0 / (high - low)).astype(np.uint8)


def _clahe(gray: np.ndarray) -> np.ndarray:
    if cv2 is None:
        return _contrast(gray)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    return clahe.apply(gray)


def _sharpen(gray: np.ndarray) -> np.ndarray:
    if cv2 is None:
        return gray
    kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]], dtype=np.float32)
    return cv2.filter2D(gray, -1, kernel)


def _threshold_otsu(gray: np.ndarray) -> np.ndarray:
    if cv2 is None:
        threshold = float(gray.mean())
        return (gray >= threshold).astype(np.uint8) * 255
    _, output = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return output


def _threshold_adaptive(gray: np.ndarray) -> np.ndarray:
    if cv2 is None:
        return _threshold_otsu(gray)
    block_size = max(11, min(51, (min(gray.shape[:2]) // 20) | 1))
    return cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        block_size,
        7,
    )


def _rotate(image: np.ndarray, angle: float) -> np.ndarray:
    if abs(angle) < 0.01 or cv2 is None:
        return image.copy()
    height, width = image.shape[:2]
    center = (width / 2.0, height / 2.0)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    return cv2.warpAffine(
        image,
        matrix,
        (width, height),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_REPLICATE,
    )


def _horizontal_score(gray: np.ndarray) -> float:
    if gray.size == 0:
        return 0.0
    binary = gray < np.percentile(gray, 45)
    row_density = binary.mean(axis=1)
    if len(row_density) < 2:
        return 0.0
    return float(np.mean(np.sort(row_density)[-min(12, len(row_density)):]))


def auto_deskew(gray: np.ndarray, max_angle: int = 8) -> tuple[np.ndarray, float]:
    if cv2 is None or gray.size == 0:
        return gray.copy(), 0.0
    best_image = gray
    best_angle = 0.0
    best_score = _horizontal_score(gray)
    for angle in range(-max_angle, max_angle + 1):
        if angle == 0:
            continue
        candidate = _rotate(gray, float(angle))
        score = _horizontal_score(candidate)
        if score > best_score:
            best_image, best_angle, best_score = candidate, float(angle), score
    return best_image, best_angle


def _roi_candidates(image: np.ndarray) -> Iterable[tuple[str, np.ndarray, tuple[int, int, int, int]]]:
    height, width = image.shape[:2]
    if height < 2 or width < 2:
        return

    seen: set[tuple[int, int, int, int]] = set()
    for name, fraction in (("full", 1.0), ("bottom35", 0.35), ("bottom45", 0.45), ("bottom55", 0.55)):
        top = 0 if fraction == 1.0 else max(0, int(height * (1.0 - fraction)))
        roi = (0, top, width, height)
        if roi in seen:
            continue
        seen.add(roi)
        x1, y1, x2, y2 = roi
        yield name, image[y1:y2, x1:x2], roi

    # A projection-based crop often removes the visual zone while retaining both MRZ rows.
    gray = _gray(image)
    bottom_top = int(height * 0.35)
    bottom = gray[bottom_top:]
    if bottom.size:
        darkness = (bottom < np.percentile(bottom, 45)).mean(axis=1)
        active = darkness > max(0.04, float(np.percentile(darkness, 65)))
        if active.any():
            indexes = np.flatnonzero(active)
            y1 = max(bottom_top, bottom_top + int(indexes[0]) - max(8, height // 100))
            y2 = min(height, bottom_top + int(indexes[-1]) + max(8, height // 100))
            if y2 - y1 >= max(20, int(height * 0.08)):
                roi = (0, y1, width, y2)
                if roi not in seen:
                    yield "projection", image[y1:y2, 0:width], roi


def generate_preprocessed_variants(
    image: bytes | bytearray | BytesIO | Image.Image | np.ndarray,
    max_variants: int = 8,
) -> list[PreprocessedImage]:
    """Generate bounded, deliberately diverse OCR inputs."""
    source = decode_image(image)
    groups: list[list[PreprocessedImage]] = []
    for roi_name, roi_image, roi in _roi_candidates(source):
        gray = _gray(roi_image)
        deskewed, angle = auto_deskew(gray)
        base = _resize(roi_image, 3.0)
        gray3 = _resize(deskewed, 3.0)
        gray3 = _contrast(gray3)
        candidates = [
            (f"{roi_name}_raw3x", base),
            (f"{roi_name}_gray_contrast3x", gray3),
            (f"{roi_name}_clahe3x", _clahe(gray3)),
            (f"{roi_name}_otsu3x", _threshold_otsu(gray3)),
            (f"{roi_name}_adaptive3x", _threshold_adaptive(gray3)),
            (f"{roi_name}_sharpen2x", _resize(_sharpen(_contrast(deskewed)), 2.0)),
        ]
        groups.append(
            [
                PreprocessedImage(name=variant_name, image=variant_image, roi=roi, angle=angle)
                for variant_name, variant_image in candidates
            ]
        )

    # Prefer all enhancement variants of the first/full ROI, then add a second
    # ROI. This keeps the default budget useful for both crop diversity and
    # threshold diversity instead of exhausting it on one ROI only.
    variants: list[PreprocessedImage] = []
    if groups:
        variants.extend(groups[0][:max_variants])
        for group in groups[1:]:
            if len(variants) >= max_variants:
                break
            variants.extend(group[: max_variants - len(variants)])
    return variants[:max_variants]
