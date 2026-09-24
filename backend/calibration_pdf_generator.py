import fitz  # PyMuPDF
import io
import datetime
import logging

logger = logging.getLogger("CalibrationPDFGenerator")

def generate_calibration_sheet_pdf(user_name: str = "Operator") -> bytes:
    """
    Generates a printable A4 PDF Calibration Sheet with a 6x6 grid (36 boxes for 0-9 and A-Z).
    Includes 4 corner fiducial markers for precision homography alignment when scanned back.
    """
    # A4 Dimensions in points (72 DPI): 595 x 842 pt
    doc = fitz.open()
    page = doc.new_page(width=595, height=842)
    
    # Colors
    black = (0, 0, 0)
    blue = (0.1, 0.3, 0.8)
    gray = (0.5, 0.5, 0.5)
    light_blue = (0.9, 0.95, 1.0)
    
    # 1. Corner Fiducial Markers (Black 20x20 pt squares)
    page.draw_rect(fitz.Rect(20, 20, 40, 40), color=black, fill=black)
    page.draw_rect(fitz.Rect(555, 20, 575, 40), color=black, fill=black)
    page.draw_rect(fitz.Rect(20, 802, 40, 822), color=black, fill=black)
    page.draw_rect(fitz.Rect(555, 802, 575, 822), color=black, fill=black)
    
    # 2. Header Title & Instructions
    page.insert_text(fitz.Point(60, 40), "HANDWRITING CALIBRATION SHEET", fontsize=18, color=blue)
    page.insert_text(fitz.Point(60, 55), "36-Character Profile Training Grid (0-9 & A-Z)", fontsize=10, color=gray)
    
    date_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    page.insert_text(fitz.Point(400, 40), f"User Name: {user_name}", fontsize=11, color=black)
    page.insert_text(fitz.Point(400, 55), f"Date: {date_str}", fontsize=9, color=gray)
    
    page.draw_line(fitz.Point(50, 65), fitz.Point(545, 65), color=blue, width=1)
    
    instruction = "Instruction: Write each indicated character clearly inside the corresponding blue box below."
    page.insert_text(fitz.Point(50, 80), instruction, fontsize=9, color=black)
    
    # 3. 6x6 Grid Configuration
    chars = [str(i) for i in range(10)] + [chr(i) for i in range(ord('A'), ord('Z') + 1)] # 36 chars
    
    grid_left = 60
    grid_top = 100
    box_width = 75
    box_height = 105
    padding_x = 7
    padding_y = 12
    
    rows = 6
    cols = 6
    
    idx = 0
    for r in range(rows):
        for c in range(cols):
            if idx >= len(chars):
                break
            char_label = chars[idx]
            
            x0 = grid_left + c * (box_width + padding_x)
            y0 = grid_top + r * (box_height + padding_y)
            x1 = x0 + box_width
            y1 = y0 + box_height
            
            # Draw Outer Cell Boundary
            page.draw_rect(fitz.Rect(x0, y0, x1, y1), color=gray, fill=light_blue, width=0.5)
            
            # Character Tag Header
            tag_rect = fitz.Rect(x0, y0, x1, y0 + 20)
            page.draw_rect(tag_rect, color=blue, fill=blue)
            page.insert_text(fitz.Point(x0 + (box_width / 2) - 4, y0 + 14), char_label, fontsize=12, color=(1, 1, 1))
            
            # Writing Box Inside
            write_rect = fitz.Rect(x0 + 5, y0 + 25, x1 - 5, y1 - 5)
            page.draw_rect(write_rect, color=blue, width=1)
            
            idx += 1
            
    # Footer Note
    page.insert_text(fitz.Point(140, 815), "Scan at 300 DPI and upload to the Application for Calibration Training.", fontsize=8, color=gray)
    
    pdf_bytes = doc.tobytes()
    doc.close()
    logger.info(f"Generated printable Calibration Sheet PDF for user '{user_name}'.")
    return pdf_bytes
