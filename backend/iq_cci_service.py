import os
import io
import re
import datetime
import pandas as pd
import openpyxl
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

DEFAULT_TEMPLATE_PATH = os.path.abspath(r"C:\Users\cpreephim\Desktop\App Team\IQOQDQ\IQ_CCI\XXXXX_05_IQ_Control Components_20XX-XX-XX_en.docx")
DEFAULT_EXCEL_PATH = os.path.abspath(r"C:\Users\cpreephim\Desktop\App Team\IQOQDQ\IQ_CCI\5XXXX_Part list.xlsx")

def parse_iq_cci_excel(excel_source):
    """
    Parses Parts List Excel for IQ CCI:
    1. Finds sheet starting with 'IQOQ list' (case-insensitive).
    2. Extracts Model, Order No, Serial No, Customer from top rows.
    3. Extracts Column B (CIS.), Column F (Designation), Column H (Supplier).
    """
    if isinstance(excel_source, (str, os.PathLike)):
        wb = openpyxl.load_workbook(excel_source, data_only=True)
    elif hasattr(excel_source, 'read'):
        wb = openpyxl.load_workbook(excel_source, data_only=True)
    else:
        wb = openpyxl.load_workbook(excel_source, data_only=True)

    # Find target sheet starting with 'IQOQ list'
    target_sheet = None
    for sn in wb.sheetnames:
        if sn.strip().lower().startswith("iqoq list"):
            target_sheet = sn
            break
    if not target_sheet:
        target_sheet = wb.sheetnames[0]

    ws = wb[target_sheet]

    # Extract metadata from top rows (Rows 1-5)
    meta = {
        "model": "",
        "order_no": "",
        "serial_no": "",
        "customer": "",
        "sheet_name": target_sheet
    }

    for r in range(1, min(7, ws.max_row + 1)):
        for c in range(1, min(10, ws.max_column + 1)):
            v = str(ws.cell(r, c).value or '').strip()
            if 'model' in v.lower():
                val_next = str(ws.cell(r, c + 1).value or '').strip().lstrip(':').strip()
                if val_next: meta["model"] = val_next
            elif 'order no' in v.lower():
                val_next = str(ws.cell(r, c + 1).value or '').strip().lstrip(':').strip()
                if val_next: meta["order_no"] = val_next
            elif 'serial no' in v.lower():
                val_next = str(ws.cell(r, c + 1).value or '').strip().lstrip(':').strip()
                if val_next: meta["serial_no"] = val_next
            elif 'customer' in v.lower():
                val_next = str(ws.cell(r, c + 1).value or '').strip().lstrip(':').strip()
                if val_next: meta["customer"] = val_next

    # Find header row (usually Row 6)
    header_row_idx = 6
    for r in range(1, min(10, ws.max_row + 1)):
        row_txt = " ".join([str(ws.cell(r, c).value or '') for c in range(1, min(10, ws.max_column + 1))]).lower()
        if 'cis' in row_txt and 'designation' in row_txt:
            header_row_idx = r
            break

    records = []
    item_counter = 1

    for r in range(header_row_idx + 1, ws.max_row + 1):
        part_no = str(ws.cell(r, 1).value or '').strip()
        cis = str(ws.cell(r, 2).value or '').strip()
        loc = str(ws.cell(r, 3).value or '').strip()
        qty = str(ws.cell(r, 4).value or '').strip()
        unit = str(ws.cell(r, 5).value or '').strip()
        desig = str(ws.cell(r, 6).value or '').strip()
        desc = str(ws.cell(r, 7).value or '').strip()
        supplier = str(ws.cell(r, 8).value or '').strip()

        # Skip rows if CIS and Designation are both empty
        if not cis and not desig:
            continue

        records.append({
            "Test-point": f"{item_counter}.",
            "CIS": cis,
            "Function": desig,
            "Manufacturer": supplier,
            "Location": loc,
            "Qty": qty,
            "Description": desc
        })
        item_counter += 1

    preview_df = pd.DataFrame(records)

    return {
        "metadata": meta,
        "sheet_name": target_sheet,
        "total_items": len(records),
        "records": records,
        "preview_df": preview_df
    }

from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE

def generate_iq_cci_word(
    excel_source,
    word_template_path=None,
    row_height_pt=22.0,
    **kwargs
):
    """
    Generates IQ Control Components & Instruments (CCI) Word Protocol:
    1. Parses target 'IQOQ list*' sheet in Excel for Cols B (CIS), F (Designation), H (Supplier).
    2. Opens Word Master Template and locates Table 6 (Section 3 Test protocol).
    3. Clears placeholder rows and populates all component records sequentially.
    4. Enforces row height fixed at 22.0 pt and vertical alignment matching the master template.
    5. Sets downloaded file name to the exact same name as the master template.
    6. Formats cells and returns binary doc_bytes for direct in-memory download.
    """
    parsed = parse_iq_cci_excel(excel_source)
    records = parsed["records"]
    meta = parsed["metadata"]

    if not records:
        raise ValueError("No valid component records found in sheet. Please verify Column B (CIS) and Column F (Designation).")

    template_to_use = word_template_path if (word_template_path and os.path.exists(word_template_path)) else DEFAULT_TEMPLATE_PATH
    if not os.path.exists(template_to_use):
        raise FileNotFoundError(f"Master template file not found: {template_to_use}")

    doc = docx.Document(template_to_use)

    # Locate Section 3 Test Protocol table (Table 6)
    # Header specifically has c0='Test-point' and c1='CIS' (without Tag no)
    target_table = None
    for tbl in doc.tables:
        if len(tbl.columns) == 8 and len(tbl.rows) > 0:
            c0 = tbl.rows[0].cells[0].text.strip().lower()
            c1 = tbl.rows[0].cells[1].text.strip().lower()
            if "test-point" in c0 and c1 == "cis":
                target_table = tbl
                break

    if target_table is None:
        target_table = doc.tables[5] if len(doc.tables) >= 6 else doc.tables[-2]

    # Clear placeholder rows from Row 1 onwards
    while len(target_table.rows) > 1:
        tr = target_table.rows[1]._tr
        target_table._tbl.remove(tr)

    # Populate all component records
    for item in records:
        row = target_table.add_row()
        row.height = Pt(row_height_pt)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

        cells = row.cells
        cells[0].text = item["Test-point"]
        cells[1].text = item["CIS"]
        cells[2].text = item["Function"]
        cells[3].text = item["Manufacturer"]
        cells[4].text = "" # Comment
        cells[5].text = "" # Pass / Fail
        cells[6].text = "" # Date
        cells[7].text = "" # Initials

        # Font and paragraph formatting
        for c_idx in range(8):
            cell = cells[c_idx]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(1.5)
                p.paragraph_format.space_after = Pt(1.5)
                p.paragraph_format.line_spacing = 1.0
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(9)
                if c_idx == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # Save to in-memory buffer
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    doc_bytes = doc_io.getvalue()

    # Use exact same file name as the master template
    file_name = os.path.basename(template_to_use)

    return {
        "success": True,
        "file_name": file_name,
        "doc_bytes": doc_bytes,
        "total_items": len(records),
        "sheet_name": parsed["sheet_name"],
        "metadata": meta,
        "preview_df": parsed["preview_df"]
    }
