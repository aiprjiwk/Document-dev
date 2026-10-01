import os
import re
import glob
import zipfile
import io
import shutil
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from typing import List, Dict, Tuple, Optional
from backend.models import get_db_session, SupplierOEMComponent

def parse_oem_bom_excel(file_path_or_buffer, require_d500_filter: bool = True) -> pd.DataFrame:
    """
    Parses OEM BOM Excel file based on the OEM App specification (bomProcessor.js):
    1. Inspects headers for Z Hersteller (A), Z Materialkurztext EN (B), Z Materialkurztext DE (C), 
       Component number (D), Z D500 (G)
    2. Filters rows where Column G (Z D500) contains '7' (if require_d500_filter=True)
    3. Deduplicates rows by Part No. (Component number)
    4. Sorts alphabetically by Supplier name
    5. Assigns Reg. (first letter of Supplier name upon first occurrence)
    """
    df = pd.read_excel(file_path_or_buffer)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Identify key column names
    col_supplier = next((c for c in df.columns if 'hersteller' in c.lower() and 'nr' not in c.lower()), 'Z Hersteller')
    col_desc_en = next((c for c in df.columns if 'materialkurztext en' in c.lower() or ('en' in c.lower() and 'text' in c.lower())), 'Z Materialkurztext EN')
    col_desc_de = next((c for c in df.columns if 'materialkurztext de' in c.lower() or ('de' in c.lower() and 'text' in c.lower())), 'Z Materialkurztext DE')
    col_part_no = next((c for c in df.columns if 'component number' in c.lower() or 'part' in c.lower() or 'sach' in c.lower()), 'Component number')
    col_d500 = next((c for c in df.columns if 'd500' in c.lower() or c.lower() == 'g'), 'Z D500')
    col_mfg_no = next((c for c in df.columns if 'herstellernr' in c.lower()), 'Z Herstellernr.')
    col_size = next((c for c in df.columns if 'size' in c.lower() or 'dimension' in c.lower()), 'Size/dimensions')

    filtered_rows = []
    seen_part_nos = set()

    for idx, row in df.iterrows():
        val_g = str(row.get(col_d500, '') if col_d500 in df.columns else '').strip()
        
        # Filter condition: Column G contains '7'
        if not require_d500_filter or '7' in val_g or not col_d500 in df.columns:
            supplier = str(row.get(col_supplier, '') if col_supplier in df.columns else '').strip()
            desc_en = str(row.get(col_desc_en, '') if col_desc_en in df.columns else '').strip()
            desc_de = str(row.get(col_desc_de, '') if col_desc_de in df.columns else '').strip()
            part_no = str(row.get(col_part_no, '') if col_part_no in df.columns else '').strip()
            mfg_no = str(row.get(col_mfg_no, '') if col_mfg_no in df.columns else '').strip()
            size = str(row.get(col_size, '') if col_size in df.columns else '').strip()

            # Handle NaN / None values
            if supplier.lower() == 'nan': supplier = ''
            if desc_en.lower() == 'nan': desc_en = ''
            if desc_de.lower() == 'nan': desc_de = ''
            if part_no.lower() == 'nan': part_no = ''
            if mfg_no.lower() == 'nan': mfg_no = ''
            if size.lower() == 'nan': size = ''

            if part_no and part_no not in seen_part_nos:
                seen_part_nos.add(part_no)
                filtered_rows.append({
                    'manufacturer': supplier,
                    'material_desc_en': desc_en,
                    'material_desc_de': desc_de,
                    'component_number': part_no,
                    'manufacturer_part_no': mfg_no,
                    'size_dimensions': size,
                    'd500': val_g
                })

    # Sort alphabetically by Supplier
    filtered_rows.sort(key=lambda x: x['manufacturer'].lower())

    # Assign Reg.
    current_supplier = ''
    for row in filtered_rows:
        sup = row['manufacturer']
        if sup and sup.upper() != current_supplier.upper():
            current_supplier = sup
            match = re.search(r'[A-Za-z]', sup)
            row['reg'] = match.group(0).upper() if match else sup[0].upper()
        else:
            row['reg'] = ''

    return pd.DataFrame(filtered_rows)

def clean_unwanted_files(target_dir: str) -> int:
    """
    Recursively deletes Thumbs.db and temporary OS files in target directory.
    """
    removed_count = 0
    if not os.path.exists(target_dir):
        return 0
        
    for root, dirs, files in os.walk(target_dir):
        for f in files:
            if f.lower() in ['thumbs.db', '.ds_store', 'desktop.ini'] or f.startswith('~$'):
                full_p = os.path.join(root, f)
                try:
                    os.remove(full_p)
                    removed_count += 1
                except Exception:
                    pass
    return removed_count

def batch_rename_replace_at(target_dir: str) -> Tuple[int, List[Dict[str, str]]]:
    """
    Recursively scans target_dir and renames files and folders replacing '@' with '_'.
    Skips dependency and system directories (e.g., node_modules, .git, venv).
    Returns (renamed_count, list_of_renames).
    """
    if not os.path.exists(target_dir):
        return 0, []
        
    renamed_items = []
    
    # 1. Rename files first (bottom-up to avoid path collision)
    for root, dirs, files in os.walk(target_dir, topdown=False):
        if 'node_modules' in root or '.git' in root or '__pycache__' in root:
            continue
        for f in files:
            if '@' in f:
                old_p = os.path.join(root, f)
                new_f = f.replace('@', '_')
                new_p = os.path.join(root, new_f)
                try:
                    if not os.path.exists(new_p):
                        os.rename(old_p, new_p)
                        renamed_items.append({
                            'type': 'file',
                            'old_name': f,
                            'new_name': new_f,
                            'folder': os.path.relpath(root, target_dir)
                        })
                    else:
                        # If destination already exists, safely remove old duplicate
                        os.remove(old_p)
                        renamed_items.append({
                            'type': 'file (dedup)',
                            'old_name': f,
                            'new_name': new_f,
                            'folder': os.path.relpath(root, target_dir)
                        })
                except Exception:
                    pass

    # 2. Rename directories bottom-up
    for root, dirs, files in os.walk(target_dir, topdown=False):
        if 'node_modules' in root or '.git' in root:
            continue
        for d in dirs:
            if '@' in d and 'node_modules' not in d:
                old_d = os.path.join(root, d)
                new_d = d.replace('@', '_')
                new_p = os.path.join(root, new_d)
                try:
                    if not os.path.exists(new_p):
                        os.rename(old_d, new_p)
                        renamed_items.append({
                            'type': 'folder',
                            'old_name': d,
                            'new_name': new_d,
                            'folder': os.path.relpath(root, target_dir)
                        })
                except Exception:
                    pass

    return len(renamed_items), renamed_items

def scan_and_match_oem_pdfs(oem_dir: str, df_bom: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, int]]:
    """
    Scans oem_dir for PDF/FDF files and matches them against BOM records.
    """
    clean_unwanted_files(oem_dir)
    
    pdf_files = []
    if os.path.exists(oem_dir):
        for root, _, files in os.walk(oem_dir):
            for f in files:
                if f.lower().endswith('.pdf') or f.lower().endswith('.fdf'):
                    pdf_files.append({
                        'filename': f,
                        'full_path': os.path.join(root, f),
                        'folder': os.path.basename(root)
                    })
                    
    matched_records = []
    matched_count = 0
    missing_count = 0
    
    for idx, row in df_bom.iterrows():
        mfg = str(row.get('manufacturer', '')).strip().lower()
        part_no = str(row.get('component_number', '')).strip()
        mfg_no = str(row.get('manufacturer_part_no', '')).strip()
        
        matched_pdf = None
        matched_full_path = ""
        matched_folder = ""
        
        if pdf_files:
            clean_part = part_no.lstrip('0') if part_no.isdigit() else part_no
            clean_mfg = mfg_no.lstrip('0') if mfg_no.isdigit() else mfg_no
            
            for pdf in pdf_files:
                fname = pdf['filename']
                fname_lower = fname.lower()
                
                # Check if Part No. or Manufacturer Part No. is in filename
                part_match = (part_no and len(part_no) > 2 and part_no.lower() in fname_lower)
                clean_part_match = (clean_part and len(clean_part) > 2 and clean_part.lower() in fname_lower)
                mfg_part_match = (mfg_no and len(mfg_no) > 2 and mfg_no.lower() in fname_lower)
                clean_mfg_match = (clean_mfg and len(clean_mfg) > 2 and clean_mfg.lower() in fname_lower)
                
                if part_match or clean_part_match or mfg_part_match or clean_mfg_match:
                    matched_pdf = fname
                    matched_full_path = pdf['full_path']
                    matched_folder = pdf['folder']
                    break
                    
        status = "Matched" if matched_pdf else "Missing"
        if matched_pdf:
            matched_count += 1
        else:
            missing_count += 1
            
        row_dict = row.to_dict()
        row_dict['pdf_filename'] = matched_pdf or ''
        row_dict['pdf_full_path'] = matched_full_path
        row_dict['pdf_folder'] = matched_folder
        row_dict['pdf_status'] = status
        matched_records.append(row_dict)
        
    res_df = pd.DataFrame(matched_records)
    stats = {
        'total': len(res_df),
        'matched': matched_count,
        'missing': missing_count,
        'total_pdfs_found': len(pdf_files)
    }
    return res_df, stats

def find_oem_template_path(custom_dir: Optional[str] = None) -> Optional[str]:
    """
    Finds the XXXXX_Template_Overview sub-supplier documentation.xlsx template file.
    """
    candidates = []
    if custom_dir:
        candidates.append(os.path.join(custom_dir, "XXXXX_Template_Overview sub-supplier documentation.xlsx"))
        candidates.append(os.path.join(custom_dir, "..", "XXXXX_Template_Overview sub-supplier documentation.xlsx"))
    
    PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    candidates.extend([
        os.path.join(PROJECT_ROOT, "OEM", "XXXXX_Template_Overview sub-supplier documentation.xlsx"),
        os.path.join(PROJECT_ROOT, "XXXXX_Template_Overview sub-supplier documentation.xlsx"),
    ])
    for p in candidates:
        if os.path.exists(p):
            return os.path.abspath(p)
    return None

def inject_vml_header_logo(raw_xlsx_bytes: bytes, template_path: Optional[str] = None) -> bytes:
    """
    Injects VML header drawing and logo image (image1.jpeg) from template into output xlsx zip archive.
    Ensures that the IWK logo in the header is preserved and visible upon printing and viewing.
    """
    tpl_path = template_path or find_oem_template_path()
    if not tpl_path or not os.path.exists(tpl_path):
        return raw_xlsx_bytes

    try:
        with zipfile.ZipFile(tpl_path, 'r') as tz:
            if 'xl/drawings/vmlDrawing1.vml' not in tz.namelist() or 'xl/media/image1.jpeg' not in tz.namelist():
                return raw_xlsx_bytes
            vml_drawing = tz.read('xl/drawings/vmlDrawing1.vml')
            vml_drawing_rels = tz.read('xl/drawings/_rels/vmlDrawing1.vml.rels')
            media_img = tz.read('xl/media/image1.jpeg')

        final_buf = io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(raw_xlsx_bytes), 'r') as gz, zipfile.ZipFile(final_buf, 'w', compression=zipfile.ZIP_DEFLATED) as oz:
            has_sheet1_rels = False
            for item in gz.infolist():
                data = gz.read(item.filename)
                if item.filename == '[Content_Types].xml':
                    xml_str = data.decode('utf-8')
                    if 'vmlDrawing' not in xml_str:
                        add_types = '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Default Extension="jpeg" ContentType="image/jpeg"/>'
                        xml_str = xml_str.replace('</Types>', add_types + '</Types>')
                    data = xml_str.encode('utf-8')
                elif item.filename == 'xl/worksheets/sheet1.xml':
                    xml_str = data.decode('utf-8')
                    if 'xmlns:r=' not in xml_str:
                        xml_str = xml_str.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ')
                    if 'legacyDrawingHF' not in xml_str:
                        xml_str = xml_str.replace('</worksheet>', '<legacyDrawingHF r:id="rId2"/></worksheet>')
                    data = xml_str.encode('utf-8')
                elif item.filename == 'xl/worksheets/_rels/sheet1.xml.rels':
                    has_sheet1_rels = True
                    xml_str = data.decode('utf-8')
                    if 'vmlDrawing1.vml' not in xml_str:
                        rel_tag = '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
                        xml_str = xml_str.replace('</Relationships>', rel_tag + '</Relationships>')
                    data = xml_str.encode('utf-8')
                oz.writestr(item, data)

            if not has_sheet1_rels:
                rels_xml = (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
                    '</Relationships>'
                )
                oz.writestr('xl/worksheets/_rels/sheet1.xml.rels', rels_xml.encode('utf-8'))

            oz.writestr('xl/drawings/vmlDrawing1.vml', vml_drawing)
            oz.writestr('xl/drawings/_rels/vmlDrawing1.vml.rels', vml_drawing_rels)
            oz.writestr('xl/media/image1.jpeg', media_img)

        return final_buf.getvalue()
    except Exception:
        return raw_xlsx_bytes

def generate_overview_excel_buffer(
    df_bom: pd.DataFrame, 
    machine_type: str = "5XXXX IWK TZC", 
    template_path: Optional[str] = None
) -> bytes:
    """
    Generates populated Overview Sub-Supplier Documentation Excel buffer based on
    the exact template: XXXXX_Template_Overview sub-supplier documentation.xlsx
    Configures landscape, A4, fit to 1 page width, repeat header, balanced column widths,
    merging for supplier names and Reg. letter groups, cross-page split, and restored header logo.
    """
    from openpyxl.worksheet.pagebreak import Break
    
    tpl_path = template_path or find_oem_template_path()
    if tpl_path and os.path.exists(tpl_path):
        wb = openpyxl.load_workbook(tpl_path)
        # Remove extra sheets
        for name in ['Sheet2', 'Sheet3']:
            if name in wb.sheetnames:
                del wb[name]
        ws = wb.worksheets[0]
        # Clear old merged cells from row 2 onwards
        for m in list(ws.merged_cells.ranges):
            if m.min_row >= 2:
                ws.merged_cells.remove(m)
        # Clear existing sample rows from 2 onwards
        while ws.max_row > 1:
            ws.delete_rows(2, ws.max_row)
    else:
        # Fallback if template file is missing
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        headers = ["Supplier\nHersteller", "Designation", "Bezeichnung", "Part No.\nSach.Nr.", "Reg.\nReg.\n"]
        ws.append(headers)

    # 1. Update oddHeader right section with machine type
    clean_machine = (machine_type or "5XXXX IWK TZC").strip()
    try:
        ws.HeaderFooter.oddHeader.right.text = clean_machine
    except Exception:
        pass

    # 2. Font, Border, Alignment Definitions matching template
    font_regular = Font(name='Arial', size=10, bold=False)
    font_bold = Font(name='Arial', size=10, bold=True)
    align_left = Alignment(horizontal='left', vertical='center')
    align_left_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_center = Alignment(horizontal='center', vertical='center')

    thin_side = Side(style='thin', color='000000')
    med_side = Side(style='medium', color='000000')

    # Ensure header row 1 has thick borders all around
    for c in range(1, 6):
        c_hdr = ws.cell(1, c)
        c_hdr.border = Border(left=med_side, right=med_side, top=med_side, bottom=med_side)

    records = df_bom.to_dict('records')
    last_row = max(len(records) + 1, 2)

    # 3. Balanced column widths fitting strictly within A4 Landscape 1-page width (~126 total units)
    col_widths = {
        'A': 28.0,  # Supplier
        'B': 36.0,  # Designation (reduced from 47.0)
        'C': 36.0,  # Bezeichnung (reduced from 47.0)
        'D': 18.0,  # Part No.
        'E': 8.0    # Reg.
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

    # Remove any extra column dimensions from template (like F, G)
    for col_extra in ['F', 'G', 'H']:
        if col_extra in ws.column_dimensions:
            del ws.column_dimensions[col_extra]

    # Populate data in columns B, C, D for every row
    for idx, r in enumerate(records, start=2):
        ws.row_dimensions[idx].height = 18.0
        
        # Col B: Designation
        c2 = ws.cell(idx, 2, str(r.get('material_desc_en', '')).strip())
        c2.font = font_regular
        c2.alignment = align_left
        
        # Col C: Bezeichnung
        c3 = ws.cell(idx, 3, str(r.get('material_desc_de', '')).strip())
        c3.font = font_regular
        c3.alignment = align_left
        
        # Col D: Part No. (Text format '@' to display green corner marker in Excel)
        c4 = ws.cell(idx, 4, str(r.get('component_number', '')).strip())
        c4.font = font_regular
        c4.number_format = '@'
        c4.alignment = align_center

    # 4. Smart Page Breaks definition (~25 data rows per page)
    page_breaks = [26, 51, 76, 101, 126, 151, 176, 201]

    # Group 1: Supplier (Column A) - Group by identical supplier names
    sup_groups = []
    curr_sup = None
    start_idx = 2
    for i, r in enumerate(records):
        row_num = i + 2
        sup = str(r.get('manufacturer', '')).strip()
        if sup != curr_sup:
            if curr_sup is not None:
                sup_groups.append((curr_sup, start_idx, row_num - 1))
            curr_sup = sup
            start_idx = row_num
    if curr_sup is not None:
        sup_groups.append((curr_sup, start_idx, len(records) + 1))

    # Split supplier groups across page boundaries so supplier name is retained on new page
    sup_segments = []
    for sup, s_row, e_row in sup_groups:
        cur_s = s_row
        while cur_s <= e_row:
            next_pb = None
            for pb in page_breaks:
                if cur_s <= pb < e_row:
                    next_pb = pb
                    break
            if next_pb is not None:
                sup_segments.append((sup, cur_s, next_pb))
                cur_s = next_pb + 1
            else:
                sup_segments.append((sup, cur_s, e_row))
                break

    # Format and merge Supplier (Col A) and format columns B, C, D
    # Vertical borders along columns are medium.
    # Horizontal borders between different suppliers are medium.
    # Horizontal borders between rows inside the same supplier are thin.
    for sup, s_row, e_row in sup_segments:
        ws.cell(s_row, 1, sup)
        if s_row < e_row:
            ws.merge_cells(start_row=s_row, end_row=e_row, start_column=1, end_column=1)

        for r in range(s_row, e_row + 1):
            is_top = (r == s_row)
            is_bot = (r == e_row)
            top_b = med_side if is_top else None
            bot_b = med_side if is_bot else None

            # Col 1: Supplier
            c1 = ws.cell(r, 1)
            c1.font = font_bold
            c1.alignment = align_left_wrap
            c1.border = Border(left=med_side, right=med_side, top=top_b, bottom=bot_b)

            # Col 2, 3, 4: Designation, Bezeichnung, Part No.
            top_data = med_side if is_top else thin_side
            bot_data = med_side if is_bot else thin_side
            for col_idx in [2, 3, 4]:
                c_data = ws.cell(r, col_idx)
                c_data.border = Border(left=med_side, right=med_side, top=top_data, bottom=bot_data)

    # Group 2: Reg. (Column E) - Group by starting letter (Letter Group)
    reg_groups = []
    curr_letter = None
    start_reg_idx = 2
    for i, r in enumerate(records):
        row_num = i + 2
        sup = str(r.get('manufacturer', '')).strip()
        m = re.search(r'[A-Za-z]', sup)
        letter = m.group(0).upper() if m else (sup[:1].upper() if sup else '')
        if letter != curr_letter:
            if curr_letter is not None:
                reg_groups.append((curr_letter, start_reg_idx, row_num - 1))
            curr_letter = letter
            start_reg_idx = row_num
    if curr_letter is not None:
        reg_groups.append((curr_letter, start_reg_idx, len(records) + 1))

    # Split Reg. letter groups across page boundaries
    reg_segments = []
    for letter, s_row, e_row in reg_groups:
        cur_s = s_row
        while cur_s <= e_row:
            next_pb = None
            for pb in page_breaks:
                if cur_s <= pb < e_row:
                    next_pb = pb
                    break
            if next_pb is not None:
                reg_segments.append((letter, cur_s, next_pb))
                cur_s = next_pb + 1
            else:
                reg_segments.append((letter, cur_s, e_row))
                break

    # Format and merge Reg. (Col E)
    for letter, s_row, e_row in reg_segments:
        ws.cell(s_row, 5, letter)
        if s_row < e_row:
            ws.merge_cells(start_row=s_row, end_row=e_row, start_column=5, end_column=5)

        for r in range(s_row, e_row + 1):
            c5 = ws.cell(r, 5)
            c5.font = font_bold
            c5.alignment = align_center
            is_top = (r == s_row)
            is_bot = (r == e_row)
            top_b = med_side if is_top else None
            bot_b = med_side if is_bot else None
            c5.border = Border(left=med_side, right=med_side, top=top_b, bottom=bot_b)

    # Bottom border of the very last row
    for col_i in range(1, 6):
        c = ws.cell(last_row, col_i)
        c.border = Border(left=c.border.left, right=c.border.right, top=c.border.top, bottom=med_side)

    # 5. Configure Ready-to-use Page Setup for Printing & PDF Export (Landscape, A4, 1-page width, Centered)
    try:
        ws.page_setup.scale = None
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.print_title_rows = '1:1'
        ws.print_area = f'A1:E{last_row}'
        ws.print_options.horizontalCentered = True

        # Configure manual page breaks
        ws.row_breaks.brk.clear()
        for pb in page_breaks:
            if pb < last_row:
                ws.row_breaks.append(Break(id=pb))
    except Exception:
        pass

    raw_buffer = io.BytesIO()
    wb.save(raw_buffer)
    raw_bytes = raw_buffer.getvalue()

    # 6. Inject restored VML header drawing & logo image into output archive
    final_bytes = inject_vml_header_logo(raw_bytes, tpl_path)
    return final_bytes

def build_documentation_zip_package(
    oem_dir: str, 
    df_matched: pd.DataFrame, 
    machine_type: str = "5XXXX IWK TZC",
    template_path: Optional[str] = None
) -> bytes:
    """
    Builds complete ZIP package containing Overview Excel + organized PDFs by Supplier folders
    (fileMatcher.js equivalent).
    """
    prefix = re.match(r'^([A-Za-z0-9_-]+)', machine_type)
    prefix_str = prefix.group(1) if prefix else "5XXXX"
    excel_filename = f"{prefix_str}_Overview sub-supplier documentation.xlsx"
    
    excel_buffer = generate_overview_excel_buffer(df_matched, machine_type, template_path)
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', compression=zipfile.ZIP_STORED) as zf:
        # Add Overview Excel at root
        zf.writestr(excel_filename, excel_buffer)
        
        # Scan repo PDFs
        repo_pdfs = []
        if os.path.exists(oem_dir):
            for root, _, files in os.walk(oem_dir):
                for f in files:
                    if f.lower().endswith('.pdf') or f.lower().endswith('.fdf'):
                        repo_pdfs.append({
                            'filename': f,
                            'full_path': os.path.join(root, f)
                        })
                        
        added_paths = set()
        for _, row in df_matched.iterrows():
            part_no = str(row.get('component_number', '')).strip()
            sup_folder = re.sub(r'[\\/:*?"<>|]', '_', str(row.get('manufacturer', 'UNASSIGNED'))).strip() or "UNASSIGNED"
            
            if part_no:
                for pdf in repo_pdfs:
                    if part_no.lower() in pdf['filename'].lower():
                        clean_fname = pdf['filename'].replace('@', '_')
                        zip_path = f"{sup_folder}/{clean_fname}"
                        if zip_path not in added_paths:
                            added_paths.add(zip_path)
                            with open(pdf['full_path'], 'rb') as pf:
                                zf.writestr(zip_path, pf.read())
                                
    return zip_buffer.getvalue()

def sync_oem_to_sqlite(df_records: pd.DataFrame) -> int:
    """
    Syncs processed OEM BOM records into SQLite database.
    """
    session = get_db_session()
    saved_count = 0
    try:
        session.query(SupplierOEMComponent).delete()
        
        for _, row in df_records.iterrows():
            item = SupplierOEMComponent(
                manufacturer=str(row.get('manufacturer', '')),
                material_desc_en=str(row.get('material_desc_en', '')),
                material_desc_de=str(row.get('material_desc_de', '')),
                component_number=str(row.get('component_number', '')),
                size_dimensions=str(row.get('size_dimensions', '')),
                manufacturer_part_no=str(row.get('manufacturer_part_no', '')),
                d500=str(row.get('d500', '')),
                pdf_filename=str(row.get('pdf_filename', '')),
                pdf_status=str(row.get('pdf_status', 'Missing'))
            )
            session.add(item)
            saved_count += 1
        session.commit()
    except Exception as e:
        session.rollback()
        raise e
    finally:
        session.close()
    return saved_count

def generate_missing_purchasing_report(df_records: pd.DataFrame) -> pd.DataFrame:
    """
    Filters missing items for Procurement / Purchasing email request.
    """
    missing_df = df_records[df_records['pdf_status'] == 'Missing'].copy()
    req_cols = ['manufacturer', 'manufacturer_part_no', 'component_number', 'material_desc_en', 'size_dimensions']
    present_cols = [c for c in req_cols if c in missing_df.columns]
    missing_clean = missing_df[present_cols].drop_duplicates()
    return missing_clean

def get_oem_repo_dir(base_dir: str) -> str:
    """
    Returns the primary folder where supplier directories reside.
    Checks 'DataBase_Supplier' first, then 'New folder', then base_dir.
    """
    candidate_db = os.path.join(base_dir, "DataBase_Supplier")
    if os.path.exists(candidate_db) and os.path.isdir(candidate_db):
        return candidate_db
    candidate_nf = os.path.join(base_dir, "New folder")
    if os.path.exists(candidate_nf) and os.path.isdir(candidate_nf):
        return candidate_nf
    return base_dir

def list_oem_folders(base_dir: str) -> List[str]:
    """
    Returns a sorted list of subfolder names in the repository.
    """
    repo = get_oem_repo_dir(base_dir)
    if not os.path.exists(repo):
        return []
    folders = [
        d for d in os.listdir(repo) 
        if os.path.isdir(os.path.join(repo, d)) and not d.startswith('.') and d.lower() != 'node_modules'
    ]
    return sorted(folders, key=lambda s: s.lower())

def create_oem_folder(base_dir: str, folder_name: str) -> Tuple[bool, str]:
    """
    Creates a new folder for a supplier in the repository.
    """
    clean_name = re.sub(r'[\\/:*?"<>|]', '_', folder_name).replace('@', '_').strip()
    if not clean_name:
        return False, "Invalid folder name or folder name cannot be empty."
    repo = get_oem_repo_dir(base_dir)
    target_path = os.path.join(repo, clean_name)
    if os.path.exists(target_path):
        return False, f"Folder '{clean_name}' already exists."
    try:
        os.makedirs(target_path, exist_ok=True)
        return True, f"Created folder '{clean_name}' successfully."
    except Exception as e:
        return False, f"Error creating folder: {e}"

def delete_oem_folder(base_dir: str, folder_name: str) -> Tuple[bool, str]:
    """
    Deletes a folder and all its contents safely.
    """
    repo = get_oem_repo_dir(base_dir)
    target_path = os.path.join(repo, folder_name)
    if not os.path.exists(target_path):
        return False, f"Folder '{folder_name}' not found."
    try:
        shutil.rmtree(target_path)
        return True, f"Deleted folder '{folder_name}' and all contents successfully."
    except Exception as e:
        return False, f"Error deleting folder: {e}"

def list_oem_files_in_folder(base_dir: str, folder_name: str) -> List[Dict[str, str]]:
    """
    Lists files inside a specified supplier folder with size and modified timestamp.
    """
    repo = get_oem_repo_dir(base_dir)
    folder_path = os.path.join(repo, folder_name)
    if not os.path.exists(folder_path):
        return []
    import datetime
    file_list = []
    for root, _, files in os.walk(folder_path):
        for f in files:
            full_p = os.path.join(root, f)
            try:
                stat = os.stat(full_p)
                size_kb = round(stat.st_size / 1024, 1)
                mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
                rel_path = os.path.relpath(full_p, folder_path)
                file_list.append({
                    'filename': f,
                    'relative_path': rel_path,
                    'full_path': full_p,
                    'size_kb': f"{size_kb:,} KB",
                    'modified': mtime
                })
            except Exception:
                pass
    return sorted(file_list, key=lambda x: x['filename'].lower())

def save_uploaded_oem_file(base_dir: str, folder_name: str, file_name: str, file_bytes: bytes) -> Tuple[bool, str]:
    """
    Saves uploaded file into specified folder, replacing '@' with '_'.
    """
    clean_name = file_name.replace('@', '_').strip()
    repo = get_oem_repo_dir(base_dir)
    folder_path = os.path.join(repo, folder_name)
    os.makedirs(folder_path, exist_ok=True)
    target_path = os.path.join(folder_path, clean_name)
    try:
        with open(target_path, 'wb') as f:
            f.write(file_bytes)
        return True, f"Saved file '{clean_name}' successfully."
    except Exception as e:
        return False, f"Error saving file: {e}"

def delete_oem_file(full_file_path: str) -> Tuple[bool, str]:
    """
    Deletes a specific file.
    """
    if not os.path.exists(full_file_path):
        return False, "File not found."
    try:
        fname = os.path.basename(full_file_path)
        os.remove(full_file_path)
        return True, f"Deleted file '{fname}' successfully."
    except Exception as e:
        return False, f"Failed to delete file: {e}"
