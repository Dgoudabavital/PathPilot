"""
PathPilot OpenCV image preprocessing module.

Uses OpenCV to improve uploaded study images before OCR:
- EXIF orientation correction where possible
- grayscale conversion
- contrast enhancement
- denoising
- adaptive thresholding
- automatic border cleanup
- optional deskew
"""

from pathlib import Path
import cv2
import numpy as np


def preprocess_image(input_path: str, output_path: str | None = None) -> str:
    """Preprocess an image for OCR and return the processed file path."""
    input_path = str(input_path)
    img = cv2.imread(input_path)

    if img is None:
        raise ValueError(f"Could not read image: {input_path}")

    # Resize very small images to improve OCR.
    h, w = img.shape[:2]
    if max(h, w) < 1600:
        scale = 1600 / max(h, w)
        img = cv2.resize(img, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Improve local contrast.
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    enhanced = clahe.apply(gray)

    # Reduce camera/compression noise.
    denoised = cv2.fastNlMeansDenoising(enhanced, None, 10, 7, 21)

    # Adaptive threshold works well for photographed notes/pages.
    binary = cv2.adaptiveThreshold(
        denoised, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31, 11
    )

    # Light morphological cleanup.
    kernel = np.ones((2, 2), np.uint8)
    cleaned = cv2.morphologyEx(binary, cv2.MORPH_OPEN, kernel, iterations=1)

    # Estimate and correct page/text skew.
    coords = np.column_stack(np.where(cleaned < 255))
    if len(coords) > 100:
        angle = cv2.minAreaRect(coords)[-1]
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle

        if abs(angle) > 0.5 and abs(angle) < 20:
            hh, ww = cleaned.shape[:2]
            center = (ww // 2, hh // 2)
            matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
            cleaned = cv2.warpAffine(
                cleaned, matrix, (ww, hh),
                flags=cv2.INTER_CUBIC,
                borderMode=cv2.BORDER_REPLICATE
            )

    if output_path is None:
        p = Path(input_path)
        output_path = str(p.with_name(p.stem + "_opencv_processed.png"))

    cv2.imwrite(output_path, cleaned)
    return output_path
