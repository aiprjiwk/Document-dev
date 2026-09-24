import cv2
import numpy as np
from PIL import Image
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("AlignmentService")

def order_points(pts: np.ndarray) -> np.ndarray:
    """
    Orders 4 points as: top-left, top-right, bottom-right, bottom-left.
    """
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]

    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]
    return rect

def align_document_page(image_input, target_width: int = 2100, target_height: int = 2970) -> Image.Image:
    """
    Aligns and deskews a scanned document page using OpenCV perspective transformation (Homography).
    Normalizes output to target dimensions (e.g. 2100x2970 pixels for 300 DPI A4).
    """
    if isinstance(image_input, Image.Image):
        img_np = np.array(image_input.convert('RGB'))
    else:
        img_np = image_input

    gray = cv2.cvtColor(img_np, cv2.COLOR_RGB2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(blurred, 50, 200)

    # Find contours
    contours, _ = cv2.findContours(edged.copy(), cv2.RETR_LIST, cv2.CHAIN_APPROX_SIMPLE)
    contours = sorted(contours, key=cv2.contourArea, reverse=True)[:5]

    screen_cnt = None
    for c in contours:
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, 0.02 * peri, True)
        if len(approx) == 4:
            screen_cnt = approx
            break

    if screen_cnt is not None and cv2.contourArea(screen_cnt) > (img_np.shape[0] * img_np.shape[1] * 0.2):
        pts = screen_cnt.reshape(4, 2)
        rect = order_points(pts)
        dst = np.array([
            [0, 0],
            [target_width - 1, 0],
            [target_width - 1, target_height - 1],
            [0, target_height - 1]
        ], dtype="float32")

        matrix = cv2.getPerspectiveTransform(rect, dst)
        warped = cv2.warpPerspective(img_np, matrix, (target_width, target_height))
        logger.info("Successfully aligned document using 4-point Perspective Transform.")
        return Image.fromarray(warped)
    else:
        # Fallback: Resize to target width/height
        logger.info("No distinct 4-corner document contour found. Resizing image to target bounds.")
        aligned_pil = Image.fromarray(img_np).resize((target_width, target_height), Image.Resampling.LANCZOS)
        return aligned_pil
