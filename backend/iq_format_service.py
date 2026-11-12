import os
import io
import re
import datetime
import tempfile
import copy
import pandas as pd
import docx
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

def find_default_iq_format_template() -> str:
    """
    Finds the active Word master template for IQ Format Parts from the project directory.
    Searches IQOQDQ/IQ_Format, IQOQDQ/Data IQOQDQ, and 0_Data/IQ.
    """
    search_dirs = [
        os.path.join(PROJECT_ROOT, "IQOQDQ", "IQ_Format"),
        os.path.join(PROJECT_ROOT, "IQOQDQ", "Data IQOQDQ"),
        os.path.join(PROJECT_ROOT, "0_Data", "IQ"),
    ]
    for d in search_dirs:
        if os.path.exists(d):
            files = [f for f in os.listdir(d) if not f.startswith("~$") and f.lower().endswith(('.docx', '.doc'))]
            for f in files:
                f_lower = f.lower()
                if "08_iq_format" in f_lower or "format parts" in f_lower:
                    return os.path.join(d, f)
            if files:
                return os.path.join(d, files[0])
    return os.path.join(PROJECT_ROOT, "IQOQDQ", "IQ_Format", "XXXXX_08_IQ_Format Parts_2025-01-02_en.docx")

DEFAULT_TEMPLATE_PATH = find_default_iq_format_template()
FALLBACK_TEMPLATE_PATH = DEFAULT_TEMPLATE_PATH

def clean_quantity(raw_qty):
    """
    Cleans quantity value and removes decimal part for whole numbers (e.g. 16.0 -> 16, 1.0 -> 1).
    """
    if pd.isna(raw_qty) or raw_qty is None:
        return '1'
    try:
        float_val = float(raw_qty)
        if float_val.is_integer():
            return str(int(float_val))
        return f"{float_val:g}"
    except (ValueError, TypeError):
        s = str(raw_qty).strip()
        if s.endswith('.0'):
            s = s[:-2]
        return s if s and s != 'nan' else '1'

def format_fn_codes(fmt_str):
    """
    Formats format codes (Funktion Zeile1).
    If multiple Fn codes are present (e.g. 'F2 F5 F6' or 'F3 F7 F8'),
    adds comma separators so it becomes 'F2, F5, F6', 'F3, F7, F8', etc.
    """
    if pd.isna(fmt_str) or not fmt_str:
        return ''
    s = str(fmt_str).strip()
    f_codes = re.findall(r'F\d+', s, flags=re.IGNORECASE)
    if len(f_codes) > 1:
        return ", ".join(code.upper() for code in f_codes)
    elif len(f_codes) == 1 and s.upper() == f_codes[0].upper():
        return f_codes[0].upper()
    return s

def parse_and_clean_format_excel(excel_source):
    """
    Parses Excel BOM and applies the required rules:
    1. Exclude non-format columns (E-Q).
    2. Exclude Size/dimensions (Col C) and Component number (Col A).
    3. Skip header row.
    4. Exclude rows where Format Code (Col D / Funktion Zeile1) is blank.
    5. Clean Quantity (remove decimal places like 16.0 -> 16).
    6. Format multiple Fn codes with commas (e.g. F2 F5 F6 -> F2, F5, F6).
    """
    if isinstance(excel_source, (str, os.PathLike)):
        df = pd.read_excel(excel_source, header=None)
    else:
        df = pd.read_excel(excel_source, header=None)

    data_rows = df.iloc[1:].copy()
    
    # Filter rows where col 3 (Funktion Zeile1 / Used for format) is valid
    valid_rows = data_rows[
        data_rows[3].notna() & 
        (data_rows[3].astype(str).str.strip() != '') & 
        (data_rows[3].astype(str).str.strip() != 'nan')
    ].copy()

    groups = {}
    flat_records = []
    all_fn_tags = set()
    
    item_counter = 1
    for idx, r in valid_rows.iterrows():
        desc = str(r[1]).strip() if pd.notna(r[1]) else 'Unknown'
        fmt_raw = str(r[3]).strip() if pd.notna(r[3]) else ''
        fmt = format_fn_codes(fmt_raw)
        
        for tag in re.findall(r'F\d+', fmt, flags=re.IGNORECASE):
            all_fn_tags.add(tag.upper())
            
        raw_qty = r[17] if len(r) > 17 else '1'
        qty = clean_quantity(raw_qty)
            
        if desc not in groups:
            groups[desc] = []
        groups[desc].append((desc, fmt, qty))
        
        flat_records.append({
            'Item No.': f'{item_counter}.',
            'Category / Group': desc,
            'Designation': desc,
            'Used for Format': fmt,
            'Quantity (PCS)': qty
        })
        item_counter += 1

    preview_df = pd.DataFrame(flat_records)
    fn_tags_sorted = sorted(list(all_fn_tags), key=lambda x: int(re.sub(r'\D', '', x)) if re.sub(r'\D', '', x).isdigit() else 999)
    return groups, preview_df, fn_tags_sorted

def set_row_repeat_header(row):
    """Sets w:tblHeader on table row so header repeats across pages."""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn('w:tblHeader')) is None:
        trPr.append(OxmlElement('w:tblHeader'))

def set_row_cant_split(row):
    """Sets w:cantSplit on table row to prevent awkward page splits."""
    trPr = row._tr.get_or_add_trPr()
    if trPr.find(qn('w:cantSplit')) is None:
        trPr.append(OxmlElement('w:cantSplit'))

def convert_legacy_doc_to_docx(doc_path: str) -> str:
    """
    Converts a legacy binary .doc file to temporary .docx via Word COM.
    Used only if a user uploads a legacy .doc file.
    """
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0
    temp_docx = os.path.join(tempfile.gettempdir(), f"conv_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.docx")
    try:
        doc = word.Documents.Open(os.path.abspath(doc_path), ReadOnly=True, Visible=False)
        doc.SaveAs2(temp_docx, FileFormat=16) # 16 = wdFormatXMLDocument (.docx)
        doc.Close(False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return temp_docx

def generate_iq_format_word(
    excel_source,
    word_template_path=None,
    word_template_source=None,
    edited_df=None,
    metadata=None,
    **kwargs
):
    """
    Directly opens the Master Word template from IQ_Format directory,
    modifies Table 3 (Completeness of Format Parts), updates Quantity (Set) -> Quantity (PCS),
    formats columns to exact widths and fonts, removes sample dummy rows,
    populates actual format parts grouped by category, preserves all original template layout,
    and returns document bytes and preview metadata for direct download.
    """
    # 1. Parse data (either from edited_df or raw Excel source)
    if edited_df is not None and not edited_df.empty:
        groups = {}
        flat_records = []
        item_counter = 1
        all_fn_tags = set()
        for idx, row in edited_df.iterrows():
            desc = str(row.get('Designation', row.get('Category / Group', 'Unknown'))).strip()
            fmt_raw = str(row.get('Used for Format', '')).strip()
            fmt = format_fn_codes(fmt_raw)
            for tag in re.findall(r'F\d+', fmt, flags=re.IGNORECASE):
                all_fn_tags.add(tag.upper())
            qty = clean_quantity(row.get('Quantity (PCS)', row.get('Quantity (PC)', '1')))
            if desc not in groups:
                groups[desc] = []
            groups[desc].append((desc, fmt, qty))
            flat_records.append({
                'Item No.': f'{item_counter}.',
                'Category / Group': desc,
                'Designation': desc,
                'Used for Format': fmt,
                'Quantity (PCS)': qty
            })
            item_counter += 1
        preview_df = pd.DataFrame(flat_records)
        fn_tags_sorted = sorted(list(all_fn_tags), key=lambda x: int(re.sub(r'\D', '', x)) if re.sub(r'\D', '', x).isdigit() else 999)
    else:
        groups, preview_df, fn_tags_sorted = parse_and_clean_format_excel(excel_source)

    if not groups:
        raise ValueError('No valid format parts found in Excel file. Please ensure Funktion Zeile1 (Col D) contains format data.')

    # 2. Resolve template path
    chosen_tmpl = word_template_path or word_template_source or kwargs.get('word_template_path') or kwargs.get('word_template_source')
    if chosen_tmpl and os.path.exists(os.path.abspath(str(chosen_tmpl))):
        template_to_use = os.path.abspath(str(chosen_tmpl))
    else:
        template_to_use = find_default_iq_format_template()

    if not os.path.exists(template_to_use):
        raise FileNotFoundError(f'Word master template not found for IQ Format Parts: {template_to_use}')

    # If the template is binary .doc, convert to temporary .docx first
    temp_converted_docx = None
    if template_to_use.lower().endswith('.doc'):
        temp_converted_docx = convert_legacy_doc_to_docx(template_to_use)
        active_docx_path = temp_converted_docx
    else:
        active_docx_path = template_to_use

    today_str = datetime.datetime.now().strftime('%Y-%m-%d')
    tmpl_base = os.path.splitext(os.path.basename(template_to_use))[0]

    # Format output filename
    if '20xx-xx-xx' in tmpl_base.lower() or 'yyyy-mm-dd' in tmpl_base.lower():
        download_file_name = re.sub(r'20XX-XX-XX|YYYY-MM-DD', today_str, tmpl_base, flags=re.IGNORECASE) + '.docx'
    elif re.search(r'\d{4}-\d{2}-\d{2}', tmpl_base):
        download_file_name = re.sub(r'\d{4}-\d{2}-\d{2}', today_str, tmpl_base) + '.docx'
    else:
        download_file_name = f'{tmpl_base}_{today_str}.docx'

    order_no = ""
    if metadata and isinstance(metadata, dict):
        order_no = str(metadata.get('order_no', '')).strip()
    if order_no and order_no.upper() != 'XXXXX':
        download_file_name = re.sub(r'^XXXXX', order_no, download_file_name, flags=re.IGNORECASE)

    # 3. Open document using python-docx (fast, stable, memory-efficient)
    doc = docx.Document(active_docx_path)

    # 4. Find Format Parts table dynamically
    target_tbl = None
    for tbl in doc.tables:
        for r in tbl.rows[:3]:
            row_txt = ' '.join(c.text.lower() for c in r.cells)
            if 'designation' in row_txt and ('format' in row_txt or 'quantity' in row_txt):
                target_tbl = tbl
                break
        if target_tbl:
            break

    # Fallback if header matching wasn't found: inspect tables by column count
    if not target_tbl:
        for idx in [2, 3, 1, 0]:
            if idx < len(doc.tables) and len(doc.tables[idx].columns) >= 7:
                target_tbl = doc.tables[idx]
                break

    if not target_tbl:
        raise ValueError(f'Completeness of Format Parts table not found in template. Available tables: {len(doc.tables)}')

    # 5. Update header: "Quantity (Set)" -> "Quantity\n(PCS)"
    if len(target_tbl.rows) > 1:
        for cell in target_tbl.rows[1].cells:
            c_text_l = cell.text.lower()
            if 'set' in c_text_l or 'quantity' in c_text_l or 'pc' in c_text_l:
                cell.text = 'Quantity\n(PCS)'
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if p.runs:
                    p.runs[0].font.name = 'Arial'
                    p.runs[0].font.size = Pt(9)
                    p.runs[0].font.bold = True
                break

    # Set repeating headers on row 0 and row 1
    for r in target_tbl.rows[:2]:
        set_row_repeat_header(r)
        set_row_cant_split(r)

    # 6. Capture XML row templates for section header (row 2) and detail row (row 3)
    if len(target_tbl.rows) >= 4:
        hdr_template_xml = copy.deepcopy(target_tbl.rows[2]._tr)
        detail_template_xml = copy.deepcopy(target_tbl.rows[3]._tr)
    elif len(target_tbl.rows) >= 3:
        hdr_template_xml = copy.deepcopy(target_tbl.rows[2]._tr)
        detail_template_xml = copy.deepcopy(target_tbl.rows[2]._tr)
    else:
        hdr_template_xml = None
        detail_template_xml = None

    # Remove all sample dummy rows from row 2 onwards
    while len(target_tbl.rows) > 2:
        tr = target_tbl.rows[-1]._tr
        tr.getparent().remove(tr)

    # Standard column widths (matching original template Table 3)
    TEMPLATE_WIDTHS = [Inches(0.5), Inches(1.75), Inches(1.0), Inches(0.6875), Inches(0.75), Inches(0.8125), Inches(0.8)]

    # 7. Populate sections and detail items
    item_counter = 1
    for grp_name, items in groups.items():
        # A. Section Header Row
        if hdr_template_xml is not None:
            new_tr = copy.deepcopy(hdr_template_xml)
            target_tbl._tbl.append(new_tr)
            new_row = target_tbl.rows[-1]
            new_row.cells[0].text = grp_name
            p = new_row.cells[0].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if p.runs:
                p.runs[0].font.name = 'Arial'
                p.runs[0].font.size = Pt(9.5)
                p.runs[0].font.bold = True
        else:
            new_row = target_tbl.add_row()
            merged_cell = new_row.cells[0]
            for c_idx in range(1, len(new_row.cells)):
                merged_cell.merge(new_row.cells[c_idx])
            merged_cell.text = grp_name
            p = merged_cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            if p.runs:
                p.runs[0].font.name = 'Arial'
                p.runs[0].font.size = Pt(9.5)
                p.runs[0].font.bold = True

        set_row_cant_split(new_row)

        # B. Detail Rows
        for item in items:
            if detail_template_xml is not None:
                det_tr = copy.deepcopy(detail_template_xml)
                target_tbl._tbl.append(det_tr)
                row = target_tbl.rows[-1]
            else:
                row = target_tbl.add_row()

            set_row_cant_split(row)

            vals = [
                (f'{item_counter}.', WD_ALIGN_PARAGRAPH.CENTER, False),
                (item[0], WD_ALIGN_PARAGRAPH.LEFT, False),
                (item[1], WD_ALIGN_PARAGRAPH.CENTER, False),
                (item[2], WD_ALIGN_PARAGRAPH.CENTER, False),
                ('', WD_ALIGN_PARAGRAPH.CENTER, False),
                ('', WD_ALIGN_PARAGRAPH.CENTER, False),
                ('', WD_ALIGN_PARAGRAPH.CENTER, False),
            ]

            for k, (val, align, bold) in enumerate(vals):
                if k < len(row.cells):
                    c = row.cells[k]
                    c.text = val
                    if k < len(TEMPLATE_WIDTHS):
                        c.width = TEMPLATE_WIDTHS[k]
                    p = c.paragraphs[0]
                    p.alignment = align
                    if p.runs:
                        p.runs[0].font.name = 'Arial'
                        p.runs[0].font.size = Pt(9)
                        p.runs[0].font.bold = bold

            item_counter += 1

    # Save to binary in-memory stream
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    doc_bytes = doc_io.getvalue()

    # Clean up converted temporary docx if one was created
    if temp_converted_docx and os.path.exists(temp_converted_docx):
        try:
            os.remove(temp_converted_docx)
        except Exception:
            pass

    return {
        'success': True,
        'total_groups': len(groups),
        'total_items': len(preview_df),
        'file_name': download_file_name,
        'doc_bytes': doc_bytes,
        'preview_df': preview_df,
        'groups': groups,
        'fn_tags': fn_tags_sorted,
        'metadata': metadata or {}
    }
