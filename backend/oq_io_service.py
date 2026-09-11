import os
import io
import re
import copy
import datetime
import pandas as pd
import openpyxl
import xlrd
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE

DEFAULT_TEMPLATE_PATH = os.path.abspath(r"C:\Users\cpreephim\Desktop\App Team\IQOQDQ\IO list\XXXXX_03_OQ_IO Tests_20XX-XX-XX_en.docx")
DEFAULT_SAMPLE_PATH = os.path.abspath(r"C:\Users\cpreephim\Desktop\App Team\IQOQDQ\IO list\5XXXX_IOList_1.xls")

def parse_and_compute_iolist(file_source):
    """
    Parses IO List from ELCAD (*.xls, *.xlsx), filters blanks,
    sorts by Column E (cross ref.), and applies OQ IO test formulas.
    """
    raw_rows = []
    
    # Check if file_source is bytes/uploaded file or path
    if isinstance(file_source, (str, os.PathLike)):
        ext = os.path.splitext(str(file_source))[1].lower()
        if ext == '.xls':
            wb = xlrd.open_workbook(file_source)
            ws = wb.sheets()[0]
            for r in range(ws.nrows):
                row_vals = [str(ws.cell_value(r, c)).strip() for c in range(min(5, ws.ncols))]
                raw_rows.append(row_vals)
        else:
            wb = openpyxl.load_workbook(file_source, data_only=True)
            ws = wb.sheets[0] if hasattr(wb, 'sheets') else wb[wb.sheetnames[0]]
            for r in range(1, ws.max_row + 1):
                row_vals = [str(ws.cell(r, c).value or '').strip() for c in range(1, min(6, ws.max_column + 1))]
                raw_rows.append(row_vals)
    elif hasattr(file_source, 'read'):
        # Uploaded file buffer
        file_bytes = file_source.read()
        file_source.seek(0)
        # Try xlrd first
        try:
            wb = xlrd.open_workbook(file_contents=file_bytes)
            ws = wb.sheets()[0]
            for r in range(ws.nrows):
                row_vals = [str(ws.cell_value(r, c)).strip() for c in range(min(5, ws.ncols))]
                raw_rows.append(row_vals)
        except Exception:
            wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            ws = wb[wb.sheetnames[0]]
            for r in range(1, ws.max_row + 1):
                row_vals = [str(ws.cell(r, c).value or '').strip() for c in range(1, min(6, ws.max_column + 1))]
                raw_rows.append(row_vals)
    else:
        raise ValueError("Unsupported file source format.")

    # Filter out header and blank / separator rows
    filtered_rows = []
    header_found = False

    for row_vals in raw_rows:
        # Pad row to at least 5 elements
        while len(row_vals) < 5:
            row_vals.append("")

        # Normalize trailing '.0' from floats
        for i in range(len(row_vals)):
            if row_vals[i].endswith('.0') and row_vals[i][:-2].isdigit():
                row_vals[i] = row_vals[i][:-2]

        if not header_found:
            row_text = " ".join(row_vals).lower()
            if 'module' in row_text and ('address' in row_text or 'sybolic' in row_text):
                header_found = True
            continue

        # Extract columns A..E
        module, sybolic, address, comment, cross_ref = row_vals[0], row_vals[1], row_vals[2], row_vals[3], row_vals[4]
        
        # Skip empty / separator rows
        if not module and not address and not cross_ref:
            continue

        filtered_rows.append({
            "module": module,
            "sybolic": sybolic,
            "address": address,
            "comment": comment,
            "cross_ref": cross_ref
        })

    df = pd.DataFrame(filtered_rows)
    if df.empty:
        raise ValueError("No valid IO rows found in the provided file.")

    # Sort on Column E (cross_ref)
    def sort_key(cr):
        clean_cr = re.sub(r'[-]', '', str(cr)).strip()
        if '/' in clean_cr:
            parts = clean_cr.split('/', 1)
            p1 = int(parts[0]) if parts[0].isdigit() else 99999
            try:
                p2 = float(parts[1])
            except ValueError:
                p2 = 999.0
            return (p1, p2, clean_cr)
        return (99999, 999.0, clean_cr)

    df['sort_key'] = df['cross_ref'].apply(sort_key)
    df = df.sort_values(by='sort_key').reset_index(drop=True)
    df = df.drop(columns=['sort_key'])

    # Compute OQ IO Formulas
    records = []
    t1_count = 0
    t2_count = 0

    for idx, row in df.iterrows():
        a = row['module']
        b = row['sybolic']
        c = row['address']
        d = row['comment']
        e = row['cross_ref']

        # G: Operand prefix (IX / QX / IW / etc.)
        g = c.split('_')[0] if '_' in c else c

        # H (Address):
        a_str = str(a).strip()
        b_str = str(b).strip()
        if 'T' in a_str.upper() or a_str.upper().startswith('T'):
            h = f".007{a_str}:{b_str}"
        elif b_str.upper().startswith('Q'):
            h = f".140{a_str}:{b_str}"
        else:
            h = f".004{a_str}:{b_str}"

        # I (Description):
        i_desc = str(d).strip()

        # J (Page of Wiring Diagram):
        val = str(e).replace('-', '').strip()
        if '/' in val:
            lp, rp = val.split('/', 1)
            lp = lp.strip()
            rp = rp.strip()
            if lp.startswith('0') and lp.isdigit():
                trimmed_left = f"{int(lp):03d}"
            else:
                trimmed_left = lp
            j_page = f".{trimmed_left}/{rp}"
        else:
            j_page = f".{val}"

        # K (Test):
        if g.upper() == 'QX':
            k_test = 'T1'
            t1_count += 1
        else:
            k_test = 'T2'
            t2_count += 1

        records.append({
            "Test Point": f"{idx + 1}.",
            "Address": h,
            "Description": i_desc,
            "Page of Wiring Diagram": j_page,
            "Test": k_test,
            "Signal change observed": "",
            "Pass / Fail: Dev-No.": "",
            "Date": "",
            "Initials": ""
        })

    preview_df = pd.DataFrame(records)

    return {
        "records": records,
        "preview_df": preview_df,
        "total_items": len(records),
        "t1_count": t1_count,
        "t2_count": t2_count
    }

def generate_oq_io_word(
    excel_source,
    word_template_path=None,
    row_height_pt=22.0,
    **kwargs
):
    """
    Generates OQ IO Tests Word Protocol:
    1. Parses and computes formulas for all IO points.
    2. Opens Word Master Template and locates Table 5 (Section 3 Test Protocol).
    3. Clears placeholder rows and populates all IO test records sequentially.
    4. Enforces row height fixed at 22.0 pt and vertical alignment.
    5. Returns binary doc_bytes and exact master template file name.
    """
    parsed = parse_and_compute_iolist(excel_source)
    records = parsed["records"]

    template_to_use = word_template_path if (word_template_path and os.path.exists(word_template_path)) else DEFAULT_TEMPLATE_PATH
    if not os.path.exists(template_to_use):
        raise FileNotFoundError(f"Master template file not found: {template_to_use}")

    doc = docx.Document(template_to_use)

    # Locate Section 3 Test Protocol table (Table 5)
    target_table = None
    for tbl in doc.tables:
        if len(tbl.rows) >= 2:
            h0_txt = " ".join([c.text.strip().lower() for c in tbl.rows[0].cells])
            h1_txt = " ".join([c.text.strip().lower() for c in tbl.rows[1].cells])
            if "test" in h0_txt and "address" in h0_txt and "y / n" in h1_txt:
                target_table = tbl
                break

    if target_table is None:
        target_table = doc.tables[4] if len(doc.tables) >= 5 else doc.tables[-1]

    # Keep template row XML (Row 2) as blueprint to preserve exact 9-column gridSpan layout
    template_tr = copy.deepcopy(target_table.rows[2]._tr) if len(target_table.rows) > 2 else None

    # Clear placeholder rows from Row 2 onwards (preserve header rows 0 and 1)
    while len(target_table.rows) > 2:
        tr = target_table.rows[2]._tr
        target_table._tbl.remove(tr)

    # Populate all IO test records using template row blueprint
    for item in records:
        if template_tr is not None:
            new_tr = copy.deepcopy(template_tr)
            target_table._tbl.append(new_tr)
            row = target_table.rows[-1]
            row.height = Pt(row_height_pt)
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

            tcs = new_tr.findall(docx.oxml.ns.qn('w:tc'))
            vals = [
                item["Test Point"],
                item["Address"],
                item["Description"],
                item["Page of Wiring Diagram"],
                item["Test"],
                "", "", "", ""
            ]
            for idx, (tc, val) in enumerate(zip(tcs, vals)):
                # Clear existing paragraphs inside tc except the first one
                ps = tc.findall(docx.oxml.ns.qn('w:p'))
                first_p = ps[0]
                for extra_p in ps[1:]:
                    tc.remove(extra_p)
                for r in first_p.findall(docx.oxml.ns.qn('w:r')):
                    first_p.remove(r)

                p_obj = docx.text.paragraph.Paragraph(first_p, row)
                p_obj.paragraph_format.space_before = Pt(1.5)
                p_obj.paragraph_format.space_after = Pt(1.5)
                p_obj.paragraph_format.line_spacing = 1.0

                if val:
                    run = p_obj.add_run(val)
                    run.font.name = "Arial"
                    run.font.size = Pt(9)

                if idx in [0, 4]:
                    p_obj.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p_obj.alignment = WD_ALIGN_PARAGRAPH.LEFT
        else:
            row = target_table.add_row()
            row.height = Pt(row_height_pt)
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
            cells = row.cells
            cells[0].text = item["Test Point"]
            cells[1].text = item["Address"]
            cells[2].text = item["Description"]
            cells[3].text = item["Page of Wiring Diagram"]
            cells[4].text = item["Test"]
            for c_i in range(5, len(cells)):
                cells[c_i].text = ""

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
        "total_items": parsed["total_items"],
        "t1_count": parsed["t1_count"],
        "t2_count": parsed["t2_count"],
        "preview_df": parsed["preview_df"]
    }
