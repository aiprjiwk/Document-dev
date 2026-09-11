import os
import re
import io
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy
import logging
import pandas as pd

# Set up logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(name)s: %(message)s")
logger = logging.getLogger("CalibrationService")

def get_project_code_from_filename(filename: str) -> str:
    """
    Extracts the project code (numeric prefix before underscore) from the filename.
    Example: "55782_IWK Calibration certificates_2026-01-26.xlsx" -> "55782"
    """
    base_name = os.path.splitext(filename)[0]
    token = base_name.split("_")[0]
    match = re.match(r"^\d+", token)
    if match:
        return match.group(0)
    return token

def find_image_file(base_name: str) -> str:
    picture_dir = "My Picture"
    os.makedirs(picture_dir, exist_ok=True)
    formats = [".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp"]
    for fmt in formats:
        target_lower = f"{base_name}{fmt}".lower()
        if os.path.exists(picture_dir):
            for f in os.listdir(picture_dir):
                if f.lower() == target_lower:
                    return os.path.join(picture_dir, f)
    return None

def insert_image_safely(ws, img_path: str, cell_coord: str) -> bool:
    try:
        from openpyxl.drawing.image import Image as OpenpyxlImage
        from PIL import Image as PILImage
        with PILImage.open(img_path) as pil_img:
            width, height = pil_img.size
            
        # Target bounds: cell height is 90 (approx 120 pixels), cell width is 15 (approx 112 pixels)
        # We constrain width to 100, height to 110
        max_w, max_h = 100, 110
        ratio = min(max_w / width, max_h / height)
        new_w = int(width * ratio)
        new_h = int(height * ratio)
        
        img = OpenpyxlImage(img_path)
        img.width = new_w
        img.height = new_h
        
        ws.add_image(img, cell_coord)
        return True
    except Exception as e:
        logger.warning(f"Could not insert image {img_path} at {cell_coord}: {e}")
        return False

def process_calibration_sheet(excel_bytes: bytes, filename: str) -> bytes:
    """
    Translates and executes the Excel layout transformations of the VBA Macro 'RunBothSubs' & 'InsertImagesWithSplitAndHeaders12'.
    
    Processing Steps:
      1. Capture original column B values (Z) for rows where column P is "CALIBRATION".
      2. Bold rows where A=2 above rows with "CALIBRATION" in P.
      3. Delete rows where Font.Bold = False and P is not "CALIBRATION" (header is kept).
      4. Split/Duplicate rows based on Repeat Count in Column E.
      5. Shift, copy, and map original columns to the new layout format:
         - A: Running numbers 1..n for non-bold B rows (bold B rows are empty).
         - B: Original Column B.
         - C: New empty column.
         - D: Original Column C.
         - E: Original Column G.
         - F: Original Column H.
         - G: Original Column D.
         - H: Picture (Column where images are dynamically inserted).
         - I: Original Column R ("Certificate No.").
         - J: Generated Serial No. (XXXXX-Y-Z-PC) where XXXXX is project code, Y is nearest B bold label,
              Z is the captured original column B value, and PC is the occurrence counter within group.
      6. Apply standard formatting, clear column J and A values for bold B rows.
      7. Configure column widths and return final bytes.
    """
    logger.info(f"Starting Calibration Sheet processing for file: {filename}")
    
    try:
        # Load workbook
        wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
        ws = wb.active
        
        # 1. Capture Z before run (AA is column 27, B is column 2, P is column 16)
        z_values = {}
        for r in range(2, ws.max_row + 1):
            p_val = str(ws.cell(row=r, column=16).value or "").strip()
            if p_val == "CALIBRATION":
                z_values[r] = str(ws.cell(row=r, column=2).value or "").strip()
        
        logger.info(f"Captured {len(z_values)} original Z values from column B.")
        
        # 2. Bold rows where A=2 above rows with "CALIBRATION" in P
        bold_font = Font(bold=True)
        bold_count = 0
        for r in range(2, ws.max_row + 1):
            p_val = str(ws.cell(row=r, column=16).value or "").strip()
            a_val = ws.cell(row=r, column=1)
            
            is_a_2 = False
            a_val_val = a_val.value
            try:
                if a_val_val is not None and int(float(a_val_val)) == 2:
                    is_a_2 = True
            except ValueError:
                pass
                
            if p_val == "CALIBRATION" and not is_a_2:
                # Look upwards for the nearest A=2 row
                for j in range(r - 1, 0, -1):
                    aj_val = ws.cell(row=j, column=1).value
                    is_aj_2 = False
                    try:
                        if aj_val is not None and int(float(aj_val)) == 2:
                            is_aj_2 = True
                    except ValueError:
                        pass
                    if is_aj_2:
                        for col in range(1, ws.max_column + 1):
                            ws.cell(row=j, column=col).font = bold_font
                        bold_count += 1
                        break
        
        logger.info(f"Bolded {bold_count} header rows matching A=2 criteria.")
        
        # 3. Read raw row data before any deletions/shifts to prevent index misalignment
        raw_rows = []
        for r in range(1, ws.max_row + 1):
            row_cells = []
            for c in range(1, max(ws.max_column + 1, 27)):
                cell = ws.cell(row=r, column=c)
                row_cells.append({
                    'value': cell.value,
                    'font': copy(cell.font) if cell.font else None,
                    'fill': copy(cell.fill) if cell.fill else None,
                    'alignment': copy(cell.alignment) if cell.alignment else None,
                    'border': copy(cell.border) if cell.border else None,
                    'number_format': cell.number_format
                })
            raw_rows.append((r, row_cells))
            
        # 4. Filter kept rows (Delete rows where Font.Bold = False and P is not "CALIBRATION")
        # Row 1 (header) is index 0, always kept
        kept_rows = [raw_rows[0]]
        for idx in range(1, len(raw_rows)):
            r_num, row_cells = raw_rows[idx]
            
            is_row_bold = False
            col1_font = row_cells[0]['font']
            if col1_font and col1_font.bold:
                is_row_bold = True
                
            p_val = str(row_cells[15]['value'] or "").strip()
            
            if is_row_bold or p_val == "CALIBRATION":
                kept_rows.append(raw_rows[idx])
                
        logger.info(f"Filtered rows. Kept {len(kept_rows)} out of {len(raw_rows)} rows.")
        
        # 5. Row splitting based on Comp. Qty (Column G, index 6 of raw row cells)
        split_kept_rows = [kept_rows[0]]
        for idx in range(1, len(kept_rows)):
            r_num, row_cells = kept_rows[idx]
            
            # Check if this is a header row (A == 2 in original raw sheet)
            orig_a_val = row_cells[0]['value']
            is_header_row = False
            try:
                if orig_a_val is not None and int(float(orig_a_val)) == 2:
                    is_header_row = True
            except ValueError:
                pass
                
            if is_header_row:
                split_kept_rows.append(kept_rows[idx])
            else:
                # Column G is index 6 (0-based) for "Comp. Qty"
                qty_val = row_cells[6]['value']
                qty = 1
                try:
                    if qty_val is not None:
                        qty = int(float(qty_val))
                except ValueError:
                    pass
                    
                if qty > 1:
                    for j in range(qty):
                        # Create copy
                        new_row_cells = [copy(c) for c in row_cells]
                        # Set Column G (index 6) to 1
                        new_row_cells[6]['value'] = 1
                        split_kept_rows.append((r_num, new_row_cells))
                else:
                    split_kept_rows.append(kept_rows[idx])
                    
        logger.info(f"Split rows. Expanded to {len(split_kept_rows)} rows.")
        
        # 6. Build final data structures mapped to new columns in a new workbook to prevent corruption
        project_code = get_project_code_from_filename(filename)
        
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = ws.title
        
        # Write headers
        headers = ["No.", "Component number", "Comp. Qty", "Component unit", "Object description", "Size/dimensions", "Internal processes TH", "Picture", "Certificate No.", "Serial No.", "Sticker attach."]
        for c_idx, h_name in enumerate(headers, 1):
            new_ws.cell(row=1, column=c_idx, value=h_name)
            
        # Style headers using formatting from raw row 1
        _, raw_header_cells = raw_rows[0]
        for c_idx in range(1, min(len(raw_header_cells) + 1, 12)):
            src_idx = c_idx - 1
            if src_idx < len(raw_header_cells):
                new_ws.cell(row=1, column=c_idx).font = copy(raw_header_cells[src_idx]['font'])
                new_ws.cell(row=1, column=c_idx).fill = copy(raw_header_cells[src_idx]['fill'])
                new_ws.cell(row=1, column=c_idx).alignment = copy(raw_header_cells[src_idx]['alignment'])
                new_ws.cell(row=1, column=c_idx).border = copy(raw_header_cells[src_idx]['border'])
                
        new_ws.cell(row=1, column=8, value="Picture")
        new_ws.cell(row=1, column=9, value="Certificate No.")
        new_ws.cell(row=1, column=10, value="Serial No.")
        new_ws.cell(row=1, column=11, value="Sticker attach.")
        
        y_header = ""
        prefix = ""
        dict_counts = {}
        pc_dict = {}
        run_no = 1
        
        for dest_row_idx, (orig_row_num, row_cells) in enumerate(split_kept_rows[1:], 2):
            # Check bold status of Column B cell
            b_cell = row_cells[1]
            is_b_bold = b_cell['font'].bold if b_cell['font'] else False
            b_val = str(b_cell['value'] or "").strip()
            
            if is_b_bold and b_val != "":
                y_header = str(row_cells[2]['value'] or "").strip()
                pc_dict = {} # Reset counter dict for this group
                prefix = str(row_cells[2]['value'] or "").strip() # Prefix is drawing number in bold row
                
            # Column mapping values:
            col_A_val = None
            if not is_b_bold:
                col_A_val = run_no
                run_no += 1
                
            col_B_val = row_cells[2]['value']   # Component number (Original Column C)
            col_C_val = row_cells[6]['value']   # Comp. Qty (Original Column G)
            col_D_val = row_cells[7]['value']   # Component unit (Original Column H)
            col_E_val = row_cells[3]['value']   # Object description (Original Column D)
            col_F_val = row_cells[4]['value']   # Size/dimensions (Original Column E)
            col_G_val = row_cells[15]['value']  # Internal processes TH (Original Column P)
            col_H_val = None                    # Picture Column
            col_I_val = None                    # Certificate No. (generated later)
            
            # Generate serial
            col_J_val = ""
            orig_a_val = row_cells[0]['value']
            is_orig_a_numeric = False
            try:
                if orig_a_val is not None and float(orig_a_val) >= 1:
                    is_orig_a_numeric = True
            except ValueError:
                pass
                
            if is_orig_a_numeric:
                # Find nearest bold value above if empty
                if y_header == "":
                    for prev_idx in range(dest_row_idx - 1, 1, -1):
                        _, prev_row_cells = split_kept_rows[prev_idx - 1]
                        prev_b_cell = prev_row_cells[1]
                        if prev_b_cell['font'] and prev_b_cell['font'].bold and str(prev_b_cell['value'] or "").strip() != "":
                            y_header = str(prev_b_cell['value'] or "").strip()
                            break
                            
                z_val_raw = str(z_values.get(orig_row_num, "")).strip()
                z_val = str(int(z_val_raw)) if z_val_raw.isdigit() else z_val_raw
                
                pc = 0
                is_b_numeric = False
                try:
                    if b_cell['value'] is not None:
                        float(b_cell['value'])
                        is_b_numeric = True
                except ValueError:
                    pass
                    
                if is_b_numeric and not is_b_bold:
                    b_str = str(b_cell['value']).strip()
                    pc = pc_dict.get(b_str, 0) + 1
                    pc_dict[b_str] = pc
                    
                if pc > 0:
                    col_J_val = f"{project_code}-{y_header}-{z_val}-{pc}"
                else:
                    col_J_val = f"{project_code}-{y_header}-{z_val}-"
                    
            if is_b_bold:
                col_J_val = None
                
            col_K_val = row_cells[14]['value']  # Sticker attach. (Original Column O)
            
            # Write row values
            row_values = [col_A_val, col_B_val, col_C_val, col_D_val, col_E_val, col_F_val, col_G_val, col_H_val, col_I_val, col_J_val, col_K_val]
            for col_idx, val in enumerate(row_values, 1):
                new_ws.cell(row=dest_row_idx, column=col_idx, value=val)
                
            # Formatting transfer
            # A: non-bold, transfer alignments and borders
            new_ws.cell(row=dest_row_idx, column=1).font = Font(bold=False)
            new_ws.cell(row=dest_row_idx, column=1).alignment = copy(row_cells[0]['alignment'])
            new_ws.cell(row=dest_row_idx, column=1).border = copy(row_cells[0]['border'])
            new_ws.cell(row=dest_row_idx, column=1).fill = copy(row_cells[0]['fill'])
            
            # B: Component number -> style from original Column C (index 2)
            new_ws.cell(row=dest_row_idx, column=2).font = copy(row_cells[2]['font'])
            new_ws.cell(row=dest_row_idx, column=2).fill = copy(row_cells[2]['fill'])
            new_ws.cell(row=dest_row_idx, column=2).alignment = copy(row_cells[2]['alignment'])
            new_ws.cell(row=dest_row_idx, column=2).border = copy(row_cells[2]['border'])
            
            # C: Comp. Qty -> style from original Column G (index 6)
            new_ws.cell(row=dest_row_idx, column=3).font = copy(row_cells[6]['font'])
            new_ws.cell(row=dest_row_idx, column=3).fill = copy(row_cells[6]['fill'])
            new_ws.cell(row=dest_row_idx, column=3).alignment = copy(row_cells[6]['alignment'])
            new_ws.cell(row=dest_row_idx, column=3).border = copy(row_cells[6]['border'])
            
            # D: Component unit -> style from original Column H (index 7)
            new_ws.cell(row=dest_row_idx, column=4).font = copy(row_cells[7]['font'])
            new_ws.cell(row=dest_row_idx, column=4).fill = copy(row_cells[7]['fill'])
            new_ws.cell(row=dest_row_idx, column=4).alignment = copy(row_cells[7]['alignment'])
            new_ws.cell(row=dest_row_idx, column=4).border = copy(row_cells[7]['border'])
            
            # E: Object description -> style from original Column D (index 3)
            new_ws.cell(row=dest_row_idx, column=5).font = copy(row_cells[3]['font'])
            new_ws.cell(row=dest_row_idx, column=5).fill = copy(row_cells[3]['fill'])
            new_ws.cell(row=dest_row_idx, column=5).alignment = copy(row_cells[3]['alignment'])
            new_ws.cell(row=dest_row_idx, column=5).border = copy(row_cells[3]['border'])
            
            # F: Size/dimensions -> style from original Column E (index 4)
            new_ws.cell(row=dest_row_idx, column=6).font = copy(row_cells[4]['font'])
            new_ws.cell(row=dest_row_idx, column=6).fill = copy(row_cells[4]['fill'])
            new_ws.cell(row=dest_row_idx, column=6).alignment = copy(row_cells[4]['alignment'])
            new_ws.cell(row=dest_row_idx, column=6).border = copy(row_cells[4]['border'])
            
            # G: Internal processes TH -> style from original Column P (index 15)
            new_ws.cell(row=dest_row_idx, column=7).font = copy(row_cells[15]['font'])
            new_ws.cell(row=dest_row_idx, column=7).fill = copy(row_cells[15]['fill'])
            new_ws.cell(row=dest_row_idx, column=7).alignment = copy(row_cells[15]['alignment'])
            new_ws.cell(row=dest_row_idx, column=7).border = copy(row_cells[15]['border'])
            
            # H: Picture Column formatting (height = 90)
            new_ws.row_dimensions[dest_row_idx].height = 90
            new_ws.cell(row=dest_row_idx, column=8).border = copy(row_cells[15]['border'])
            
            # Dynamic Image Lookup & Insertion into Column H (8)
            if not is_b_bold and prefix != "":
                item_no = str(col_B_val or "").strip()
                if item_no != "":
                    base_name = f"{prefix}@{item_no}"
                    if base_name in dict_counts:
                        dict_counts[base_name] += 1
                        file_name = f"{base_name}_{dict_counts[base_name]}"
                    else:
                        dict_counts[base_name] = 0
                        file_name = base_name
                        
                    img_path = find_image_file(file_name)
                    if img_path:
                        insert_image_safely(new_ws, img_path, f"H{dest_row_idx}")
                    else:
                        # Write red error text in cell
                        red_font = Font(color="FF0000", size=9)
                        new_ws.cell(row=dest_row_idx, column=8, value=f"File Not Found: {file_name}").font = red_font
                        new_ws.cell(row=dest_row_idx, column=8).alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            
            # I: Certificate No. -> style from original Column P (index 15)
            new_ws.cell(row=dest_row_idx, column=9).font = copy(row_cells[15]['font'])
            new_ws.cell(row=dest_row_idx, column=9).fill = copy(row_cells[15]['fill'])
            new_ws.cell(row=dest_row_idx, column=9).alignment = copy(row_cells[15]['alignment'])
            new_ws.cell(row=dest_row_idx, column=9).border = copy(row_cells[15]['border'])
            
            # J: Serial No.
            new_ws.cell(row=dest_row_idx, column=10).font = Font(bold=False)
            new_ws.cell(row=dest_row_idx, column=10).alignment = Alignment(horizontal="left")
            
            # K: Sticker attach. -> style from original Column P (index 15)
            new_ws.cell(row=dest_row_idx, column=11).font = copy(row_cells[15]['font'])
            new_ws.cell(row=dest_row_idx, column=11).fill = copy(row_cells[15]['fill'])
            new_ws.cell(row=dest_row_idx, column=11).alignment = copy(row_cells[15]['alignment'])
            new_ws.cell(row=dest_row_idx, column=11).border = copy(row_cells[15]['border'])
            
        # Set column widths as per VBA
        widths = {
            "A": 6, "B": 15, "C": 8, "D": 8, "E": 25, "F": 20, "G": 20, "H": 20, "I": 25, "J": 25, "K": 25
        }
        for col_letter, width in widths.items():
            new_ws.column_dimensions[col_letter].width = width
            
        logger.info("Successfully compiled processed calibration sheet.")
        
        # Save output
        out_buf = io.BytesIO()
        new_wb.save(out_buf)
        return out_buf.getvalue()
        
    except Exception as e:
        logger.error(f"Error processing calibration Excel sheet: {e}")
        raise ValueError(f"Failed to process calibration sheet: {str(e)}")

def write_df_to_excel(df: "pd.DataFrame") -> bytes:
    """
    Writes the edited calibration DataFrame back into a styled Excel workbook bytes.
    Preserves bold header formatting for rows where column A is empty or null,
    inserts images dynamically into Column H, and configures standard column widths and formats.
    """
    import io
    import pandas as pd
    logger.info("Converting updated DataFrame back to styled Excel workbook...")
    try:
        # Pre-calculate image paths using the original DataFrame row values
        image_paths_by_row = []
        prefix = ""
        dict_counts = {}
        for idx, row in df.iterrows():
            cell_A_val = row.iloc[0]
            is_row_header = cell_A_val is None or str(cell_A_val).strip() in ["", "None", "nan", "<NA>"]
            
            img_path = None
            err_msg = None
            
            if is_row_header:
                orig_d_val = row.iloc[1] # Column index 1 is 'Component number'
                prefix = str(orig_d_val or "").strip()
            else:
                orig_d_val = row.iloc[1] # Column index 1 is 'Component number' (item_no)
                item_no = str(orig_d_val or "").strip()
                if prefix != "" and item_no != "":
                    base_name = f"{prefix}@{item_no}"
                    if base_name in dict_counts:
                        dict_counts[base_name] += 1
                        file_name = f"{base_name}_{dict_counts[base_name]}"
                    else:
                        dict_counts[base_name] = 0
                        file_name = base_name
                        
                    img_path = find_image_file(file_name)
                    if not img_path:
                        err_msg = f"File Not Found: {file_name}"
                        
            image_paths_by_row.append((img_path, err_msg))
            
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = "Raw Data"
        
        # Write columns (headers)
        headers = df.columns.tolist()
        # Clean unnamed header columns back to empty strings
        clean_headers = ["" if "unnamed" in str(h).lower() else str(h) for h in headers]
        new_ws.append(clean_headers)
        
        # Write values
        for idx, row in df.iterrows():
            new_ws.append(row.tolist())
            
        # Bold header styling
        bold_font = Font(bold=True)
        normal_font = Font(bold=False)
        
        # Apply formatting
        for r_idx in range(2, new_ws.max_row + 1):
            cell_A_val = new_ws.cell(row=r_idx, column=1).value
            
            # Row header rows (where A is empty/null) should be bold
            is_row_header = False
            if cell_A_val is None or str(cell_A_val).strip() in ["", "None", "nan", "<NA>"]:
                is_row_header = True
                
            if is_row_header:
                for col_idx in range(1, new_ws.max_column + 1):
                    new_ws.cell(row=r_idx, column=col_idx).font = bold_font
            else:
                for col_idx in range(1, new_ws.max_column + 1):
                    new_ws.cell(row=r_idx, column=col_idx).font = normal_font
                
                # Image column height adjustment
                new_ws.row_dimensions[r_idx].height = 90
                
                # Clear cell text in Column H (8) to hold image
                new_ws.cell(row=r_idx, column=8, value=None)
                
                # Retrieve pre-calculated image details
                row_list_idx = r_idx - 2
                if row_list_idx < len(image_paths_by_row):
                    img_path, err_msg = image_paths_by_row[row_list_idx]
                    if img_path:
                        insert_image_safely(new_ws, img_path, f"H{r_idx}")
                    elif err_msg:
                        red_font = Font(color="FF0000", size=9)
                        new_ws.cell(row=r_idx, column=8, value=err_msg).font = red_font
                        new_ws.cell(row=r_idx, column=8).alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
                    
            # Serial No. column styling (left-aligned, normal font weight)
            new_ws.cell(row=r_idx, column=10).font = normal_font
            new_ws.cell(row=r_idx, column=10).alignment = Alignment(horizontal="left")
            
        # Set column widths as per VBA
        widths = {
            "A": 6, "B": 15, "C": 8, "D": 8, "E": 25, "F": 20, "G": 20, "H": 20, "I": 25, "J": 25
        }
        for col_letter, width in widths.items():
            new_ws.column_dimensions[col_letter].width = width
            
        out_buf = io.BytesIO()
        new_wb.save(out_buf)
        return out_buf.getvalue()
    except Exception as e:
        logger.error(f"Error compiling DataFrame to Excel: {e}")
        raise ValueError(f"Failed to generate Excel from DataFrame: {str(e)}")

def apply_header_image_with_com(file_bytes: bytes, logo_path: str, width: float = 57.6) -> bytes:
    """
    Applies the Left Header image (logo IWK.png) using Excel COM automation (pywin32)
    because openpyxl does not support embedding header/footer images.
    """
    import tempfile
    import win32com.client
    import pythoncom
    import os
    
    pythoncom.CoInitialize()
    temp_in = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    temp_in.write(file_bytes)
    temp_in.close()
    
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        
        abs_logo_path = os.path.abspath(logo_path)
        abs_excel_path = os.path.abspath(temp_in.name)
        
        wb = xl.Workbooks.Open(abs_excel_path)
        ws = wb.Worksheets(1)
        
        # Apply header image, dimensions, and placeholder string
        ws.PageSetup.LeftHeaderPicture.Filename = abs_logo_path
        ws.PageSetup.LeftHeaderPicture.Width = width
        ws.PageSetup.LeftHeader = "&G"
        
        wb.Save()
        wb.Close()
        xl.Quit()
        
        with open(temp_in.name, "rb") as f:
            result_bytes = f.read()
            
        os.unlink(temp_in.name)
        return result_bytes
    except Exception as e:
        logger.warning(f"COM Automation failed to set header image: {e}")
        if os.path.exists(temp_in.name):
            os.unlink(temp_in.name)
        return file_bytes
    finally:
        pythoncom.CoUninitialize()

def set_final_layout(df: pd.DataFrame, filename: str, layout_type: str = "standard") -> bytes:
    """
    Applies the final print layout styling (equivalent to 'SetHeaderFooter2' / 'SetHeaderFooter_approval' VBA macro).
    
    Processing Steps:
      1. Drops helper column 'Original D' (index 6).
      2. If approval layout, inserts approval columns H and I (with unicode checkboxes) and shifts Certificate No. & Serial No. to J and K.
      3. Renames headers and sets Left/Center/Right header and center/right footer text.
      4. Standardizes column widths for columns A to I.
      5. Bolds header row (row 1).
      6. Applies a thin continuous border to all cells in the data range.
      7. Dynamic Image insertion in Column G (7).
    """
    import io
    import pandas as pd
    logger.info(f"Applying final layout transformations (layout_type={layout_type})...")
    try:
        # Create folder 'My Picture' if it doesn't exist
        picture_dir = "My Picture"
        os.makedirs(picture_dir, exist_ok=True)
        
        # Pre-calculate image paths using the original DataFrame row values
        image_paths_by_row = []
        prefix = ""
        dict_counts = {}
        for idx, row in df.iterrows():
            cell_A_val = row.iloc[0]
            is_row_header = cell_A_val is None or str(cell_A_val).strip() in ["", "None", "nan", "<NA>"]
            
            img_path = None
            err_msg = None
            
            if is_row_header:
                orig_d_val = row.iloc[1] # Column index 1 is 'Component number'
                prefix = str(orig_d_val or "").strip()
            else:
                orig_d_val = row.iloc[1] # Column index 1 is 'Component number' (item_no)
                item_no = str(orig_d_val or "").strip()
                if prefix != "" and item_no != "":
                    base_name = f"{prefix}@{item_no}"
                    if base_name in dict_counts:
                        dict_counts[base_name] += 1
                        file_name = f"{base_name}_{dict_counts[base_name]}"
                    else:
                        dict_counts[base_name] = 0
                        file_name = base_name
                        
                    img_path = find_image_file(file_name)
                    if not img_path:
                        err_msg = f"File Not Found: {file_name}"
                        
            image_paths_by_row.append((img_path, err_msg))
            
        # 1. Structure the columns based on layout_type
        col_list = df.columns.tolist()
        if layout_type == "approval":
            # For approval layout, H and I are new columns. Cert No. and Serial No. are shifted to J and K
            new_rows = []
            for idx, row in df.iterrows():
                cell_A_val = row.iloc[0]
                is_row_header = cell_A_val is None or str(cell_A_val).strip() in ["", "None", "nan", "<NA>"]
                
                iwk_val = "" if is_row_header else "☐ Yes    ☐ No"
                supp_val = "" if is_row_header else "☐ Yes    ☐ No"
                
                new_rows.append([
                    row.iloc[0],  # No.
                    row.iloc[1],  # Component number
                    row.iloc[2],  # Comp. Qty
                    row.iloc[3],  # Component unit
                    row.iloc[4],  # Object description
                    row.iloc[5],  # Size/dimensions
                    None,         # Picture Column (G)
                    iwk_val,      # IWK production check (H)
                    supp_val,     # Supplier Check (I)
                    row.iloc[8],  # Certificate No. (J)
                    row.iloc[9]   # Serial No. (K)
                ])
            headers = ["No.", "Component number", "Comp. Qty", "Component unit", "Object description", "Size/dimensions", "Picture", "IWK production check", "Supplier Check", "Certificate No.", "Serial No."]
            final_df = pd.DataFrame(new_rows, columns=headers)
            max_print_col = 9
            max_total_col = 11
        else:
            # Standard Layout
            kept_cols = []
            for idx in [0, 1, 2, 3, 4, 5, 7, 8, 9]:
                if idx < len(col_list):
                    kept_cols.append(col_list[idx])
            final_df = df[kept_cols].copy()
            max_print_col = 9
            max_total_col = 9
            
        # Build workbook
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = "Final Layout"
        
        # Write columns
        headers = final_df.columns.tolist()
        clean_headers = ["" if "unnamed" in str(h).lower() else str(h) for h in headers]
        new_ws.append(clean_headers)
        
        # Write body rows
        for idx, row in final_df.iterrows():
            new_ws.append(row.tolist())
            
        # 2. Page Setup Configuration
        new_ws.page_setup.orientation = new_ws.ORIENTATION_LANDSCAPE
        new_ws.page_setup.fitToWidth = 1
        new_ws.page_setup.fitToHeight = 0
        new_ws.sheet_properties.pageSetUpPr.fitToPage = True
        new_ws.print_area = f"A1:I{new_ws.max_row}"
        
        # Header/Footer
        project_code = get_project_code_from_filename(filename)
        new_ws.HeaderFooter.oddHeader.center.text = "LIST OF CALIBRATION CERTIFICATE"
        new_ws.HeaderFooter.oddHeader.right.text = f"{project_code} IWK XX X"
        if layout_type == "approval":
            new_ws.HeaderFooter.oddFooter.center.text = "\n\nPage &P of &N"
            new_ws.HeaderFooter.oddFooter.right.text = "&12PM sign/date ____________________________ &12Production sign/date ____________________________&12Supplier sign/date ____________________________"
        else:
            new_ws.HeaderFooter.oddFooter.center.text = "Page &P of &N"
        
        # Left Header Image
        logo_path = os.path.join(picture_dir, "logo IWK.png")
        if os.path.exists(logo_path):
            try:
                new_ws.HeaderFooter.oddHeader.left.text = "&G"
                new_ws.HeaderFooter.oddHeader.left.image = logo_path
                logger.info("Logo image successfully embedded in Left Header.")
            except Exception as e:
                logger.warning(f"Could not set left header image: {e}")
                new_ws.HeaderFooter.oddHeader.left.text = "[Logo]"
        else:
            logger.info("Logo image not found in 'My Picture/logo IWK.png'. Skipping image embedding.")
            new_ws.HeaderFooter.oddHeader.left.text = ""
            
        # 3. Apply cell borders and bold header row
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        bold_font = Font(bold=True)
        normal_font = Font(bold=False)
        
        for r_idx in range(1, new_ws.max_row + 1):
            # Bold row 1 (header)
            if r_idx == 1:
                for col_idx in range(1, max_total_col + 1):
                    new_ws.cell(row=r_idx, column=col_idx).font = bold_font
            else:
                cell_A_val = new_ws.cell(row=r_idx, column=1).value
                # If this is a group header row (A is empty/null), set bold
                is_row_header = False
                if cell_A_val is None or str(cell_A_val).strip() in ["", "None", "nan", "<NA>"]:
                    is_row_header = True
                    
                if is_row_header:
                    for col_idx in range(1, max_total_col + 1):
                        new_ws.cell(row=r_idx, column=col_idx).font = bold_font
                else:
                    for col_idx in range(1, max_total_col + 1):
                        new_ws.cell(row=r_idx, column=col_idx).font = normal_font
                    
                    if layout_type == "approval":
                        new_ws.cell(row=r_idx, column=8).alignment = Alignment(horizontal="center", vertical="center")
                        new_ws.cell(row=r_idx, column=9).alignment = Alignment(horizontal="center", vertical="center")
                    
                    # Set row height to 90 for image cell
                    new_ws.row_dimensions[r_idx].height = 90
                    
                    # Clear cell text in Column G (7) to hold image
                    new_ws.cell(row=r_idx, column=7, value=None)
                    
                    # Re-insert image dynamically from pre-calculated list
                    row_list_idx = r_idx - 2
                    if row_list_idx < len(image_paths_by_row):
                        img_path, err_msg = image_paths_by_row[row_list_idx]
                        if img_path:
                            insert_image_safely(new_ws, img_path, f"G{r_idx}")
                        elif err_msg:
                            red_font = Font(color="FF0000", size=9)
                            new_ws.cell(row=r_idx, column=7, value=err_msg).font = red_font
                            new_ws.cell(row=r_idx, column=7).alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            
            # Apply thin borders to cols A to max_total_col
            for col_idx in range(1, max_total_col + 1):
                new_ws.cell(row=r_idx, column=col_idx).border = thin_border
                
        # 4. Configure Column Widths:
        if layout_type == "approval":
            widths = {
                "A": 3,
                "B": 15,
                "C": 15,
                "D": 12,
                "E": 25,
                "F": 20,
                "G": 15,
                "H": 14,
                "I": 14,
                "J": 15,
                "K": 15
            }
        else:
            widths = {
                "A": 3,
                "B": 15,
                "C": 15,
                "D": 12,
                "E": 25,
                "F": 20,
                "G": 15,
                "H": 12,
                "I": 15
            }
        for col_letter, width in widths.items():
            new_ws.column_dimensions[col_letter].width = width
            
        logger.info("Successfully applied final page setup and formatting.")
        
        out_buf = io.BytesIO()
        new_wb.save(out_buf)
        excel_bytes = out_buf.getvalue()
        
        # Inject LeftHeaderPicture if logo file exists
        logo_path = os.path.join(picture_dir, "logo IWK.png")
        if os.path.exists(logo_path):
            try:
                excel_bytes = apply_header_image_with_com(excel_bytes, logo_path)
                logger.info("Successfully embedded header image using COM.")
            except Exception as com_err:
                logger.warning(f"Failed to apply logo image via COM: {com_err}")
                
        return excel_bytes
        
    except Exception as e:
        logger.error(f"Error applying final layout settings: {e}")
        raise ValueError(f"Failed to set final layout settings: {str(e)}")

# ---------------------------------------------------------------------------
# Calibration Image Library & Maintenance Functions (ระบบบำรุงรักษาภาพเครื่องมือ)
# ---------------------------------------------------------------------------
import zipfile
from PIL import Image as PILImage
import datetime

PICTURE_DIR = os.path.abspath("My Picture")

def list_calibration_images(search_query: str = "") -> pd.DataFrame:
    """
    Returns a DataFrame containing detailed metadata of all images in 'My Picture/'.
    Columns: Filename, Drawing No, Item No, Dimensions, Size (KB), Format, Last Modified, Path
    """
    os.makedirs(PICTURE_DIR, exist_ok=True)
    valid_exts = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif", ".webp"}
    
    records = []
    for f in sorted(os.listdir(PICTURE_DIR)):
        ext = os.path.splitext(f)[1].lower()
        if ext not in valid_exts or f.startswith('.') or f.startswith('~$'):
            continue
            
        fp = os.path.join(PICTURE_DIR, f)
        if not os.path.isfile(fp):
            continue
            
        # Extract Drawing No (prefix) and Item No
        stem = os.path.splitext(f)[0]
        drawing_no = ""
        item_no = ""
        if "@" in stem:
            parts = stem.split("@", 1)
            drawing_no = parts[0].strip()
            item_no = parts[1].strip()
        else:
            drawing_no = stem
            
        # Search filter
        if search_query:
            sq = search_query.strip().lower()
            if sq not in f.lower() and sq not in drawing_no.lower() and sq not in item_no.lower():
                continue
                
        # File info
        sz = round(os.path.getsize(fp) / 1024.0, 1)
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(fp)).strftime('%Y-%m-%d %H:%M:%S')
        
        # Dimensions
        dim_str = "-"
        try:
            with PILImage.open(fp) as img:
                w, h = img.size
                dim_str = f"{w}x{h}"
        except Exception:
            pass
            
        records.append({
            "Filename": f,
            "Drawing No": drawing_no,
            "Item No": item_no,
            "Dimensions": dim_str,
            "Size (KB)": sz,
            "Format": ext.replace('.', '').upper(),
            "Last Modified": mtime,
            "Path": fp
        })
        
    return pd.DataFrame(records)

def save_calibration_image(image_bytes: bytes, filename: str, overwrite: bool = True) -> str:
    """
    Saves an uploaded image directly to 'My Picture/' folder.
    """
    os.makedirs(PICTURE_DIR, exist_ok=True)
    clean_filename = os.path.basename(filename.strip())
    if not clean_filename:
        raise ValueError("Filename cannot be empty.")
        
    target_path = os.path.join(PICTURE_DIR, clean_filename)
    if not overwrite and os.path.exists(target_path):
        raise ValueError(f"Image '{clean_filename}' already exists in library.")
        
    with open(target_path, "wb") as f:
        f.write(image_bytes)
        
    return clean_filename

def save_structured_calibration_image(image_bytes: bytes, drawing_no: str, item_no: str, index: str = "", ext: str = ".png") -> str:
    """
    Constructs {drawing_no}@{item_no}{_index}{ext} and saves to 'My Picture/'.
    """
    d_clean = str(drawing_no).strip()
    i_clean = str(item_no).strip()
    idx_clean = str(index).strip()
    
    if not d_clean or not i_clean:
        raise ValueError("Both Drawing No and Item No are required.")
        
    if not ext.startswith("."):
        ext = f".{ext}"
        
    if idx_clean:
        filename = f"{d_clean}@{i_clean}_{idx_clean}{ext}"
    else:
        filename = f"{d_clean}@{i_clean}{ext}"
        
    return save_calibration_image(image_bytes, filename, overwrite=True)

def delete_calibration_images(filenames: list) -> list:
    """
    Deletes specified image files from 'My Picture/'.
    """
    deleted = []
    for fn in filenames:
        fp = os.path.join(PICTURE_DIR, os.path.basename(fn))
        if os.path.exists(fp):
            try:
                os.remove(fp)
                deleted.append(fn)
            except Exception as e:
                logger.warning(f"Failed to delete {fp}: {e}")
    return deleted

def get_calibration_image_bytes(filename: str) -> bytes:
    """
    Reads and returns image bytes for display / download.
    """
    fp = os.path.join(PICTURE_DIR, os.path.basename(filename))
    if os.path.exists(fp):
        with open(fp, "rb") as f:
            return f.read()
    return None

def create_images_zip_bundle(filenames: list = None) -> tuple:
    """
    Bundles all or selected images into a downloadable ZIP archive.
    Returns (zip_filename, zip_bytes).
    """
    os.makedirs(PICTURE_DIR, exist_ok=True)
    out_io = io.BytesIO()
    
    if filenames is None:
        target_files = [f for f in os.listdir(PICTURE_DIR) if os.path.isfile(os.path.join(PICTURE_DIR, f)) and not f.startswith('.')]
    else:
        target_files = [os.path.basename(f) for f in filenames if os.path.exists(os.path.join(PICTURE_DIR, os.path.basename(f)))]
        
    with zipfile.ZipFile(out_io, "w", zipfile.ZIP_DEFLATED) as zf:
        for fn in target_files:
            fp = os.path.join(PICTURE_DIR, fn)
            zf.write(fp, arcname=fn)
            
    out_io.seek(0)
    zip_bytes = out_io.getvalue()
    zip_name = f"Calibration_Images_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    return zip_name, zip_bytes

def audit_excel_images(excel_bytes: bytes) -> dict:
    """
    Scans a raw CSP2 Excel sheet, extracts all required {prefix}@{item_no} image lookups,
    and returns a breakdown of Found vs Missing images.
    """
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    
    bold_font = Font(bold=True)
    # 1. Bold rows where A=2
    for r in range(2, ws.max_row + 1):
        p_val = str(ws.cell(row=r, column=16).value or "").strip()
        if p_val == "CALIBRATION":
            for j in range(r - 1, 1, -1):
                aj_val = ws.cell(row=j, column=1).value
                is_aj_2 = False
                try:
                    if aj_val is not None and int(float(aj_val)) == 2:
                        is_aj_2 = True
                except ValueError:
                    pass
                if is_aj_2:
                    for col in range(1, ws.max_column + 1):
                        ws.cell(row=j, column=col).font = bold_font
                    break

    # 2. Collect filtered rows
    records = []
    prefix = ""
    dict_counts = {}
    
    for r in range(2, ws.max_row + 1):
        cell_a = ws.cell(row=r, column=1)
        cell_b = ws.cell(row=r, column=2)
        cell_c = ws.cell(row=r, column=3)
        cell_d = ws.cell(row=r, column=4)
        cell_g = ws.cell(row=r, column=7)
        cell_p = ws.cell(row=r, column=16)
        
        is_bold = (cell_b.font and cell_b.font.bold) or (cell_a.font and cell_a.font.bold)
        p_val = str(cell_p.value or "").strip()
        
        if is_bold and cell_b.value:
            prefix = str(cell_c.value or "").strip()
            dict_counts = {}
            continue
            
        if p_val == "CALIBRATION" and not is_bold:
            item_no = str(cell_c.value or "").strip()
            desc = str(cell_d.value or "").strip()
            qty = 1
            try:
                if cell_g.value is not None:
                    qty = int(float(cell_g.value))
            except ValueError:
                pass
                
            for q_i in range(max(1, qty)):
                base_name = f"{prefix}@{item_no}"
                if base_name in dict_counts:
                    dict_counts[base_name] += 1
                    lookup_name = f"{base_name}_{dict_counts[base_name]}"
                else:
                    dict_counts[base_name] = 0
                    lookup_name = base_name
                    
                img_path = find_image_file(lookup_name)
                records.append({
                    "Row": r,
                    "Drawing No (Prefix)": prefix,
                    "Item No": item_no,
                    "Description": desc,
                    "Expected Filename": f"{lookup_name}.png",
                    "Status": "✅ Found" if img_path else "❌ Missing",
                    "Image Path": img_path
                })
                
    df = pd.DataFrame(records)
    found_count = sum(1 for r in records if r["Status"] == "✅ Found")
    missing_count = sum(1 for r in records if r["Status"] == "❌ Missing")
    
    return {
        "records_df": df,
        "total_required": len(records),
        "total_found": found_count,
        "total_missing": missing_count
    }
