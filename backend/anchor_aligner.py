import cv2
import numpy as np
from PIL import Image
import logging

try:
    import winocr
    WINOCR_AVAILABLE = True
except Exception:
    WINOCR_AVAILABLE = False

logger = logging.getLogger("AnchorAligner")

# Mapping rules between printed labels and target Excel fields
LABEL_TARGET_MAP = [
    {"label": "format-no", "field": "FormatNo", "sheet": "F1", "cell": "F5"},
    {"label": "dimension", "field": "Dimension", "sheet": "F1", "cell": "E6"},
    {"label": "description", "field": "Description", "sheet": "F1", "cell": "E7"},
    {"label": "speed", "field": "Speed", "sheet": "F1", "cell": "E9"},
    {"label": "höhe faltschachtelkette", "field": "CartonChainHeight", "sheet": "F1", "cell": "H22"},
    {"label": "formatteile links", "field": "FormatPartsLeftRight", "sheet": "F1", "cell": "H24"},
    {"label": "saugerarm", "field": "SuctionArm", "sheet": "F1", "cell": "H25"},
    {"label": "anzahl sauger", "field": "SuctionCount", "sheet": "F1", "cell": "H26"},
    {"label": "gegensauger", "field": "CounterSuction_Z", "sheet": "F1", "cell": "H30"},
    {"label": "riegelöffner", "field": "LatchOpener_LY", "sheet": "F1", "cell": "H33"},
    {"label": "aufrichtweichen", "field": "ErectionGuides_LOY", "sheet": "F1", "cell": "H36"},
]

def auto_detect_field_boxes(page_img: Image.Image) -> list:
    """
    Dynamically scans the PDF page image for printed label text anchors using WinOCR & OpenCV,
    and returns a clean, perfectly aligned mapping list with Top/Left/Height/Width percentages.
    """
    w, h = page_img.size
    mapping_list = []
    
    if WINOCR_AVAILABLE:
        try:
            ocr_res = winocr.recognize_pil_sync(page_img)
            words = [w_info for line in ocr_res.get('lines', []) for w_info in line.get('words', [])]
            
            for item in LABEL_TARGET_MAP:
                lbl_query = item["label"].lower()
                matched_word = None
                
                for w_info in words:
                    w_text = w_info.get("text", "").lower()
                    if len(w_text) >= 3 and (w_text in lbl_query or lbl_query in w_text):
                        matched_word = w_info
                        break
                        
                if matched_word:
                    rect = matched_word["bounding_rect"]
                    bx, by, bw, bh = rect["x"], rect["y"], rect["width"], rect["height"]
                    
                    # Compute value cell box to the right
                    # If header row, value is at bx + bw + 20
                    # For main table rows, column H value is located at x ~ 76% to 80% of page
                    if bx < w * 0.5:
                        val_left_pct = ((bx + bw + 20) / w) * 100
                        val_width_pct = 25.0
                    else:
                        val_left_pct = 76.5
                        val_width_pct = 15.0
                        
                    val_top_pct = max(0.5, (by - 5) / h * 100)
                    val_height_pct = 2.2
                    
                    mapping_list.append({
                        "Field Name": item["field"],
                        "Sheet": item["sheet"],
                        "Excel Cell": item["cell"],
                        "PDF Page": 1,
                        "Top (%)": round(val_top_pct, 1),
                        "Left (%)": round(val_left_pct, 1),
                        "Height (%)": round(val_height_pct, 1),
                        "Width (%)": round(val_width_pct, 1)
                    })
        except Exception as e:
            logger.warning(f"Dynamic WinOCR anchor detection failed: {e}")

    # Fallback to calibrated default positions if anchor detection returns empty
    if not mapping_list:
        mapping_list = [
            {"Field Name": "FormatNo", "Sheet": "F1", "Excel Cell": "F5", "PDF Page": 1, "Top (%)": 19.3, "Left (%)": 57.2, "Height (%)": 2.0, "Width (%)": 20.0},
            {"Field Name": "Dimension", "Sheet": "F1", "Excel Cell": "E6", "PDF Page": 1, "Top (%)": 21.3, "Left (%)": 57.2, "Height (%)": 2.0, "Width (%)": 35.0},
            {"Field Name": "Description", "Sheet": "F1", "Excel Cell": "E7", "PDF Page": 1, "Top (%)": 23.3, "Left (%)": 57.2, "Height (%)": 2.0, "Width (%)": 35.0},
            {"Field Name": "Speed", "Sheet": "F1", "Excel Cell": "E9", "PDF Page": 1, "Top (%)": 26.9, "Left (%)": 54.9, "Height (%)": 2.2, "Width (%)": 12.0},
            {"Field Name": "CartonChainHeight", "Sheet": "F1", "Excel Cell": "H22", "PDF Page": 1, "Top (%)": 52.4, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "FormatPartsLeftRight", "Sheet": "F1", "Excel Cell": "H24", "PDF Page": 1, "Top (%)": 56.3, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "SuctionArm", "Sheet": "F1", "Excel Cell": "H25", "PDF Page": 1, "Top (%)": 58.1, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "CounterSuction_Z", "Sheet": "F1", "Excel Cell": "H30", "PDF Page": 1, "Top (%)": 63.6, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "CounterSuction_Y", "Sheet": "F1", "Excel Cell": "H31", "PDF Page": 1, "Top (%)": 66.5, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "LatchOpener_LY", "Sheet": "F1", "Excel Cell": "H33", "PDF Page": 1, "Top (%)": 71.2, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "LatchOpener_LX", "Sheet": "F1", "Excel Cell": "H34", "PDF Page": 1, "Top (%)": 73.0, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "Erection_LOY", "Sheet": "F1", "Excel Cell": "H36", "Top (%)": 80.6, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "Erection_LUY", "Sheet": "F1", "Excel Cell": "H37", "Top (%)": 83.0, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0}
        ]
        
    return mapping_list
