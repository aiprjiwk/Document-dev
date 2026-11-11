"""
fat_generator.py
Logic for reading Excel, comparing against Master Rules, and generating the Word FAT Protocol document.
"""

import os
import re
import copy
import json
import openpyxl
import docx

MASTER_FILE = os.path.join(os.path.dirname(__file__), 'master_rules.json')
CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'config.json')
DEFAULT_TEMPLATE = os.path.join(os.path.dirname(__file__), 'templates', '5XXXX_Template IWK XX FAT Protocol_20XX-MM-DD_en.docx')

def get_template_path():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
                custom_tpl = cfg.get('template_path')
                if custom_tpl and os.path.exists(custom_tpl):
                    return custom_tpl
        except Exception:
            pass
    return DEFAULT_TEMPLATE

def set_template_path(tpl_path):
    cfg = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}
    cfg['template_path'] = tpl_path
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

def reset_template_path():
    cfg = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
        except Exception:
            cfg = {}
    cfg.pop('template_path', None)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

TEMPLATE_FILE = DEFAULT_TEMPLATE

def format_l1_heading(sheet_name, template_sections=None):
    """
    Formats the Heading 1 title according to the standard:
    'Test result sheets IWK xx'
    e.g.
    'TZ' -> 'Test result sheets IWK TZ'
    'PF 34-1' -> 'Test result sheets IWK 34-1'
    '34-1' -> 'Test result sheets IWK 34-1'
    'TZC' -> 'Test result sheets IWK TZC'
    """
    clean = str(sheet_name or '').strip()
    clean = re.sub(r'^(?:test\s+result\s+sheets\s+)', '', clean, flags=re.IGNORECASE).strip()
    clean = re.sub(r'^IWK\s+', '', clean, flags=re.IGNORECASE).strip()
    
    clean_sub = re.sub(r'^(?:PF|FP)\s*', '', clean, flags=re.IGNORECASE).strip()
    if template_sections:
        for s in template_sections:
            if clean_sub.lower() in s['title'].lower():
                clean = clean_sub
                break
    else:
        if re.match(r'^(?:PF|FP)\s*\d', clean, flags=re.IGNORECASE):
            clean = clean_sub
            
    return f"Test result sheets IWK {clean}"

STANDARD_H3_LIST = [
    "Nameplate",
    "Basic machine",
    "IWK FP XX standard version",
    "Performance",
    "Machine marking",
    "Product contact parts",
    "Tube position",
    "Print registration",
    "Tube blow out",
    "Fill level monitoring",
    "Hopper lid",
    "Hopper heating system",
    "Product infeed pump",
    "Hopper",
    "Agitator",
    "Pressure cylinder",
    "Pressure hopper",
    "Filling station",
    "Blow-off system",
    "Dosing cylinder",
    "Rotary valve",
    "Mounting suspension for filling pipe",
    "Seals",
    "Filling pipe",
    "Change of product contact format parts",
    "Folding station for metal tubes",
    "Hot air station",
    "Press and embossing station",
    "Pressing and embossing tools for metal tubes",
    "Embossing digits for metal tubes",
    "Pressing and embossing tools for plastic and laminate tubes",
    "Embossing digits for plastic and laminate tubes",
    "Trimming or Punching station",
    "Extraction system tube cuttings",
    "Inkjet printing",
    "Camera control",
    "Laser marking",
    "Options HMI hardware",
    "Options HMI software",
    "Discharge station",
    "Faulty tube ejection",
    "Good tube discharge",
    "Size parts",
    "Control",
    "Electrical equipment",
    "Pressure gauge",
    "Options control cabinet",
    "Visual / audible signals",
    "Wear & tear parts",
    "Machine cover/ Guarding",
    "Leakage testing station",
    "Combination",
    "Bulkware station"
]

def load_master_rules():
    if os.path.exists(MASTER_FILE):
        with open(MASTER_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return []

def save_master_rules(rules):
    with open(MASTER_FILE, 'w', encoding='utf-8') as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)

def parse_master_rules_excel(file_or_bytes):
    """
    Parses an uploaded Excel file for Master Rules maintenance.
    Automatically detects header row, handles 5-column, 7-column, or custom formats,
    ignores timestamps and extra columns (like German Kurztext, Menge),
    and extracts: pos_no, kurztext, level_1, level_2, level_3.
    """
    import pandas as pd
    try:
        df_raw = pd.read_excel(file_or_bytes, header=None)
    except Exception as e:
        raise ValueError(f"Failed to read Excel file: {str(e)}")

    if df_raw.empty:
        return []

    # 1. Identify header row by scanning first 25 rows
    hdr_idx = None
    for idx, row in df_raw.head(25).iterrows():
        row_str = ' '.join([str(c).lower() for c in row if pd.notna(c)])
        if any(p in row_str for p in ['pos', 'posno', 'item', 'ลำดับ']) and any(k in row_str for k in ['kurztext', 'รายการ', 'ระดับ', 'level', 'description']):
            hdr_idx = idx
            break

    if hdr_idx is not None:
        header_vals = [str(c).strip() if pd.notna(c) else f'Col_{i}' for i, c in enumerate(df_raw.iloc[hdr_idx])]
        df = df_raw.iloc[hdr_idx + 1:].copy()
        df.columns = header_vals
    else:
        header_vals = [str(c).strip() if pd.notna(c) else f'Col_{i}' for i, c in enumerate(df_raw.iloc[0])]
        df = df_raw.iloc[1:].copy()
        df.columns = header_vals

    cols = list(df.columns)
    col_map = {'pos': None, 'kt': None, 'l1': None, 'l2': None, 'l3': None}

    # Match Level 1, 2, 3 and PosNo by header names
    for c in cols:
        cl = str(c).strip().lower()
        if not col_map['pos'] and any(k in cl for k in ['posno', 'pos no', 'pos.', 'pos', 'ลำดับ', 'item']):
            col_map['pos'] = c
        elif not col_map['l1'] and any(k in cl for k in ['ระดับ 1', 'ระดับ1', 'level 1', 'level1', 'heading 1', 'h1']):
            col_map['l1'] = c
        elif not col_map['l2'] and any(k in cl for k in ['ระดับ 2', 'ระดับ2', 'level 2', 'level2', 'heading 2', 'h2']):
            col_map['l2'] = c
        elif not col_map['l3'] and any(k in cl for k in ['ระดับ 3', 'ระดับ3', 'level 3', 'level3', 'หมวดหมู่', 'category', 'heading 3', 'h3']):
            col_map['l3'] = c

    # Match Kurztext / Test item
    for c in cols:
        cl = str(c).strip().lower()
        if 'kurztext/ข้อมูล' in cl or 'รายการทดสอบ' in cl:
            col_map['kt'] = c
            break
    if not col_map['kt']:
        for c in cols:
            cl = str(c).strip().lower()
            if 'kurztext' in cl and 'engl' not in cl:
                col_map['kt'] = c
                break
    if not col_map['kt']:
        for c in cols:
            cl = str(c).strip().lower()
            if any(k in cl for k in ['kurztext', 'description', 'test item', 'รายการ']):
                col_map['kt'] = c
                break

    # Fallback if some columns not found by header name:
    # Check if 7 columns (IWK export format: PosNo, Kurztext, Englischer Kurztext, Menge, Level 1, Level 2, Level 3)
    if len(cols) >= 7 and not (col_map['l1'] and col_map['l2'] and col_map['l3']):
        if not col_map['pos']: col_map['pos'] = cols[0]
        if not col_map['kt']: col_map['kt'] = cols[1]
        if not col_map['l1']: col_map['l1'] = cols[4]
        if not col_map['l2']: col_map['l2'] = cols[5]
        if not col_map['l3']: col_map['l3'] = cols[6]
    elif len(cols) >= 5 and not (col_map['l1'] and col_map['l2'] and col_map['l3']):
        if not col_map['pos']: col_map['pos'] = cols[0]
        if not col_map['kt']: col_map['kt'] = cols[1]
        if not col_map['l1']: col_map['l1'] = cols[2]
        if not col_map['l2']: col_map['l2'] = cols[3]
        if not col_map['l3']: col_map['l3'] = cols[4]

    def clean_cell_val(v):
        if pd.isna(v):
            return ""
        if isinstance(v, float) and v.is_integer():
            s = str(int(v)).strip()
        else:
            s = str(v).strip()
        if s.lower() in ["none", "nan", "null", "<na>"]:
            return ""
        return s

    imported_items = []
    for _, row in df.iterrows():
        pos = clean_cell_val(row[col_map['pos']]) if col_map['pos'] else ""
        kt = clean_cell_val(row[col_map['kt']]) if col_map['kt'] else ""
        l1 = clean_cell_val(row[col_map['l1']]) if col_map['l1'] else ""
        l2 = clean_cell_val(row[col_map['l2']]) if col_map['l2'] else ""
        l3 = clean_cell_val(row[col_map['l3']]) if col_map['l3'] else ""

        # Skip repeated header row
        if pos.lower() in ['posno', 'pos', 'ลำดับ', 'item'] or kt.lower() in ['kurztext', 'รายการทดสอบ', 'kurztext/ข้อมูล']:
            continue

        if pos or kt or l1 or l2 or l3:
            imported_items.append({
                "pos_no": pos,
                "kurztext": kt,
                "level_1": l1,
                "level_2": l2,
                "level_3": l3
            })

    return imported_items

def normalize_text(text):
    if not text:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip().lower()

def clean_val(val):
    if val is None:
        return ""
    s = str(val).strip()
    if s.lower() == 'none':
        return ""
    return s

def match_rule(kurztext, pos_no, rules):
    norm_kt = normalize_text(kurztext)
    norm_pos = clean_val(pos_no)
    
    # 1. Exact pos_no match if valid
    if norm_pos and norm_pos.lower() not in ['0', '', 'none']:
        for r in rules:
            r_pos = clean_val(r.get('pos_no', ''))
            if r_pos and r_pos == norm_pos:
                return r
                
    # 2. Exact kurztext match
    for r in rules:
        r_kt = clean_val(r.get('kurztext', ''))
        if r_kt and normalize_text(r_kt) == norm_kt:
            return r
            
    # 3. Fuzzy/Substring match
    for r in rules:
        r_kt = clean_val(r.get('kurztext', ''))
        master_kt = normalize_text(r_kt)
        if master_kt and (master_kt in norm_kt or norm_kt in master_kt):
            return r
            
    return None

def parse_excel_file(filepath):
    """
    Parses an uploaded Excel workbook.
    Extracts all sheets, each sheet representing a machine / Level 1.
    Compares against master_rules.
    """
    rules = load_master_rules()
    wb = openpyxl.load_workbook(filepath, data_only=True)
    
    result = {
        'sheets': wb.sheetnames,
        'items_by_sheet': {}
    }
    
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        items = []
        
        # Read header row (row 2 typically in sample)
        header_row_idx = 1
        rows = list(sheet.iter_rows(values_only=True))
        
        # Find which row has PosNo or Kurztext
        for idx, row in enumerate(rows[:5]):
            row_str = ' '.join([str(c or '').lower() for c in row])
            if 'posno' in row_str or 'kurztext' in row_str or 'pos' in row_str:
                header_row_idx = idx + 1
                break
                
        sheet_l1 = format_l1_heading(sheet_name)
        data_rows = rows[header_row_idx:]
        
        for r_i, row in enumerate(data_rows):
            if not any(row):
                continue
            pos_no = str(row[0] or '').strip() if len(row) > 0 and row[0] is not None else ''
            kurztext = str(row[1] or '').strip() if len(row) > 1 and row[1] is not None else ''
            kurztext_de = str(row[2] or '').strip() if len(row) > 2 and row[2] is not None else ''
            menge = str(row[3] or '1').strip() if len(row) > 3 and row[3] is not None else '1'
            
            # Excel might already have Level 1, 2, 3 in columns E, F, G
            excel_l1 = str(row[4] or '').strip() if len(row) > 4 and row[4] is not None else ''
            excel_l2 = str(row[5] or '').strip() if len(row) > 5 and row[5] is not None else ''
            excel_l3 = str(row[6] or '').strip() if len(row) > 6 and row[6] is not None else ''
            
            if not kurztext and not pos_no:
                continue
                
            matched = match_rule(kurztext, pos_no, rules)
            
            if matched:
                m_l2 = clean_val(matched.get('level_2'))
                m_l3 = clean_val(matched.get('level_3'))
                l1 = excel_l1 or sheet_l1
                l2 = m_l2 or excel_l2
                l3 = m_l3 or excel_l3
                is_new = False
                source = 'master'
            else:
                l1 = excel_l1 or sheet_l1
                l2 = excel_l2
                l3 = excel_l3
                is_new = True if (not l2 and not l3) else False
                source = 'excel' if (l2 or l3) else 'new'
                
            items.append({
                'row_idx': r_i + 1,
                'pos_no': pos_no,
                'kurztext': kurztext,
                'kurztext_de': kurztext_de,
                'menge': menge,
                'level_1': l1,
                'level_2': l2,
                'level_3': l3,
                'is_new': is_new,
                'source': source
            })
            
        result['items_by_sheet'][sheet_name] = items

    result['sheets'] = list(result['items_by_sheet'].keys())
    return result

def unwrap_sdts(docx_doc):
    """
    Unwraps all w:sdt (Structured Document Tags) in body so their inner paragraphs and tables
    are directly reachable and editable in docx python.
    """
    for sdt in docx_doc._body._element.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdt'):
        parent = sdt.getparent()
        content = sdt.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sdtContent')
        if content is not None:
            idx = parent.index(sdt)
            for child in list(content):
                parent.insert(idx, child)
                idx += 1
            parent.remove(sdt)
        else:
            parent.remove(sdt)

def clean_cell_0_formatting(cell, num_str, col0_width=950):
    """
    Removes automatic numbering (w:numPr) from cell 0 to prevent double numbering (e.g. 1. 2.1.1)
    and expands column width to prevent text wrap on short columns.
    """
    tc = cell._tc
    # Remove numPr
    for num_pr in tc.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr'):
        num_pr.getparent().remove(num_pr)
        
    # Change style to Normal to prevent numbering inheritance
    pStyle = tc.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
    if pStyle is not None:
        pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'Normal')
        
    # Set explicit cell width
    tcW = tc.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcW')
    if tcW is not None:
        tcW.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(col0_width))
        
    # Set text cleanly
    for p in cell.paragraphs[1:]:
        tc.remove(p._p)
    cell.paragraphs[0].text = str(num_str)

def ensure_table_borders(table):
    """
    Ensures standard single borders on the table properties (tblPr) so borders never disappear.
    """
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    tblPr = table._tbl.tblPr
    borders = tblPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders')
    if borders is None:
        ns = nsdecls("w")
        borders_xml = parse_xml(
            f'<w:tblBorders {ns}>\n'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n'
            f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n'
            f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n'
            f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n'
            f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n'
            f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>\n'
            f'</w:tblBorders>'
        )
        tblPr.append(borders_xml)

def format_table_pagination(table):
    """
    Sets keepWithNext, cantSplit, and tblHeader so tables break cleanly across pages
    and never orphan headers at the bottom of a page.
    """
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    ns = nsdecls("w")
    for i, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        if trPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}cantSplit') is None:
            trPr.append(parse_xml(f'<w:cantSplit {ns}/>'))
        if i == 0:
            if trPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblHeader') is None:
                trPr.append(parse_xml(f'<w:tblHeader {ns}/>'))

def adjust_table_widths(table, col0_width=950):
    """
    Expands column 0 width in tblGrid and header to prevent number text wrapping.
    """
    grid = table._tbl.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblGrid')
    if grid is not None and len(grid) >= 2:
        old_w0 = int(grid[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', '647'))
        diff = col0_width - old_w0
        grid[0].set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(col0_width))
        old_w1 = int(grid[1].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', '2971'))
        grid[1].set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(max(1000, old_w1 - diff)))
    if len(table.rows) > 0:
        tc0_w = table.rows[0].cells[0]._tc.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tcW')
        if tc0_w is not None:
            tc0_w.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', str(col0_width))

def fill_table_row(table, row_index, num_str, func_text, checkbox_template_tr=None, col0_width=950, is_new=False):
    """
    Fills row_index in table with num_str and func_text.
    Maintains checkboxes in column 2 and empty remark in column 3.
    If is_new is True, highlights the text in yellow so user can easily spot it.
    """
    if row_index < len(table.rows):
        row = table.rows[row_index]
    else:
        # Clone row from template
        if checkbox_template_tr is not None:
            new_tr = copy.deepcopy(checkbox_template_tr)
        else:
            new_tr = copy.deepcopy(table.rows[0]._tr)
        table._tbl.append(new_tr)
        row = table.rows[row_index]
        
    # Cell 0: No. (Clean formatting without numPr)
    clean_cell_0_formatting(row.cells[0], num_str, col0_width)
    
    # Cell 1: Function / Installation / Document
    c1 = row.cells[1]
    for p in c1.paragraphs[1:]:
        c1._tc.remove(p._p)
    p0 = c1.paragraphs[0]
    p0.text = ""
    run = p0.add_run(str(func_text))
    if is_new:
        run.font.highlight_color = docx.enum.text.WD_COLOR_INDEX.YELLOW
    
    # Cell 2: Checkbox (preserve existing XML or add checkbox if empty)
    c2 = row.cells[2]
    if "Yes" not in c2.text:
        c2.paragraphs[0].text = "☐\tYes\n☐\tNo"
        
    # Cell 3: Remark
    c3 = row.cells[3]
    for p in c3.paragraphs[1:]:
        c3._tc.remove(p._p)
    c3.paragraphs[0].text = ""

def clean_text(t):
    return ' '.join((t or '').split()).strip()

def match_h3_name(item_l3, h3_title):
    """
    Fuzzy matches item's level 3 with Word template Heading 3.
    e.g. 'Basic machine' <-> 'Basic machine'
    'Electrical' <-> 'Electrical equipment'
    """
    n_item = normalize_text(item_l3)
    n_title = normalize_text(h3_title)
    if not n_item or not n_title or n_item in ['no', 'none']:
        return False
    if n_item == n_title:
        return True
    if n_item in n_title or n_title in n_item:
        return True
    # Special abbreviation matches
    if 'electrical' in n_item and 'electrical' in n_title:
        return True
    return False

def build_fat_document(meta, items_by_sheet, output_path):
    """
    Builds the final FAT Protocol Word document using template.
    meta: {
        'project': '...',
        'order_no': '...',
        'customer': '...',
        'version': '1.0'
    }
    items_by_sheet: {
        'TZC': [ { 'level_1', 'level_2', 'level_3', 'kurztext', 'pos_no' }, ... ]
    }
    """
    active_template = get_template_path()
    if not os.path.exists(active_template):
        active_template = DEFAULT_TEMPLATE
    doc = docx.Document(active_template)
    unwrap_sdts(doc)
    
    # 1. Fill Table 0 (Project, Order no, Customer, Version)
    if len(doc.tables) > 0:
        t0 = doc.tables[0]
        if len(t0.rows) >= 4:
            t0.rows[0].cells[1].text = meta.get('project', '')
            t0.rows[1].cells[1].text = meta.get('order_no', '')
            t0.rows[2].cells[1].text = meta.get('customer', '')
            t0.rows[3].cells[1].text = meta.get('version', '1.0')

    # Also update Version in section headers
    for sec in doc.sections:
        for tbl in sec.header.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    if 'Version' in cell.text:
                        for p in cell.paragraphs:
                            if 'Version' in p.text:
                                p.text = f"Version {meta.get('version', '1.0')}"

    # Get sample checkbox row and sample elements for dynamic H3 creation BEFORE removing template sections!
    checkbox_tr = None
    sample_tbl_elem = None
    sample_h3_elem = None
    sample_spacer_elem = None

    if len(doc.tables) > 2 and len(doc.tables[2].rows) > 1:
        checkbox_tr = copy.deepcopy(doc.tables[2].rows[1]._tr)

    for p in doc.paragraphs:
        if p.style and p.style.name == 'Heading 3':
            sample_h3_elem = copy.deepcopy(p._p)
            break

    for t in doc.tables:
        if len(t.columns) == 4 and len(t.rows) > 0 and 'Function' in t.rows[0].cells[1].text:
            sample_tbl_elem = copy.deepcopy(t._tbl)
            break

    for p in doc.paragraphs:
        if p.style and p.style.name == 'Normal' and not p.text.strip():
            sample_spacer_elem = copy.deepcopy(p._p)
            break

    body = doc._body._element

    # Discover all machine-specific template sections between first "test result sheets" and "test run"
    template_sections = []
    current_sec = None
    all_sec_elements = []
    sec2_start = None

    for i, el in enumerate(body):
        if el.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(el, doc)
            st = p.style.name if p.style else ''
            p_txt = p.text.strip()
            if st == 'Heading 1':
                if 'test result sheets' in p_txt.lower():
                    if sec2_start is None:
                        sec2_start = i
                    if current_sec is not None:
                        template_sections.append(current_sec)
                    current_sec = {'title': p_txt, 'elements': [el]}
                    all_sec_elements.append(el)
                    continue
                elif 'test run' in p_txt.lower() and current_sec is not None:
                    template_sections.append(current_sec)
                    current_sec = None
                    break
        if current_sec is not None:
            current_sec['elements'].append(el)
            all_sec_elements.append(el)

    if current_sec is not None and current_sec not in template_sections:
        template_sections.append(current_sec)

    if not template_sections or sec2_start is None:
        raise ValueError("Could not find Chapter 'Test result sheets' in template.")

    # Remove all original template Section elements from body so we only insert the types present in Excel
    for el in all_sec_elements:
        body.remove(el)

    def find_best_template_section(l1_title, sections):
        norm = l1_title.lower()
        keywords = ['tzc', 'tzf', 'tz', '34-1', 'fp 34-1', 'fp34-1', 'fp 18', 'fp 46', 'fp']
        for kw in keywords:
            if kw in norm:
                for s in sections:
                    if kw in s['title'].lower():
                        return s
        for s in sections:
            s_words = [w for w in s['title'].lower().split() if len(w) > 2 and w not in ['test', 'result', 'sheets', 'iwk']]
            if any(w in norm for w in s_words):
                return s
        return sections[0]

    # Locate test_runs_el as the dynamic anchor point
    test_runs_el = None
    for el in body:
        if el.tag.endswith('p'):
            p = docx.text.paragraph.Paragraph(el, doc)
            if p.style and p.style.name == 'Heading 1' and 'test run' in p.text.lower():
                test_runs_el = el
                break

    sheet_keys = list(items_by_sheet.keys())
    item_counter = 1

    for sheet_idx, sheet_name in enumerate(sheet_keys):
        # Exclude items marked with 'No' (user specifies in template directly)
        sheet_items = [
            it for it in items_by_sheet[sheet_name]
            if str(it.get('level_2') or '').strip().lower() != 'no'
            and str(it.get('level_3') or '').strip().lower() != 'no'
        ]
        
        l1_heading_title = format_l1_heading(sheet_name, template_sections)
        matched_sec = find_best_template_section(l1_heading_title, template_sections)

        # Clone fresh Section elements from the matched machine template
        cloned_elements = [copy.deepcopy(el) for el in matched_sec['elements']]
        
        for el in cloned_elements:
            body.insert(body.index(test_runs_el), el)
            
        # Group items by Level 2
        docs_items = [it for it in sheet_items if it.get('level_2', '').strip().lower() == 'documents']
        safety_items = [it for it in sheet_items if it.get('level_2', '').strip().lower() == 'safety']
        mach_items = [it for it in sheet_items if 'machine details' in it.get('level_2', '').strip().lower() or (it.get('level_2', '').strip().lower() not in ['documents', 'safety'] and it.get('level_3'))]
        
        # Build list of all H3 blocks in this cloned section
        # Each H3 block: (h3_paragraph_obj, list_of_elements_in_block, table_obj_or_None)
        h3_blocks = []
        curr_h3_p = None
        curr_block = []
        
        for el in cloned_elements:
            if el.tag.endswith('p'):
                p = docx.text.paragraph.Paragraph(el, doc)
                st = p.style.name if p.style else ''
                txt = clean_text(p.text)
                
                # FIX 1: Set Heading 1 and Heading 2 titles WITHOUT hardcoded numbers
                # Let Word's Heading style auto-number them so there is no duplicate number!
                if st == 'Heading 1' and 'Test result sheets' in txt:
                    p.text = l1_heading_title
                elif st == 'Heading 2':
                    if 'Documents' in txt:
                        p.text = "Documents"
                    elif 'Safety' in txt:
                        p.text = "Safety"
                    elif 'Machine details' in txt:
                        p.text = "Machine details"
                elif st == 'Heading 3':
                    if curr_h3_p is not None:
                        tbl = None
                        for b_el in curr_block:
                            if b_el.tag.endswith('tbl'):
                                tbl = docx.table.Table(b_el, doc)
                                break
                        h3_blocks.append((curr_h3_p, curr_block, tbl))
                    curr_h3_p = p
                    curr_block = [el]
                    continue
                    
            if curr_h3_p is not None:
                curr_block.append(el)
                
        # Append last H3 block
        if curr_h3_p is not None:
            tbl = None
            for b_el in curr_block:
                if b_el.tag.endswith('tbl'):
                    tbl = docx.table.Table(b_el, doc)
                    break
            h3_blocks.append((curr_h3_p, curr_block, tbl))

        # ---------------------------------------------
        # POPULATE 2.1 DOCUMENTS TABLE
        # ---------------------------------------------
        docs_table = None
        for el in cloned_elements:
            if el.tag.endswith('tbl'):
                t_cand = docx.table.Table(el, doc)
                if len(t_cand.rows) > 0 and 'Machine layout' in ' '.join([c.text for c in t_cand.rows[1].cells]):
                    docs_table = t_cand
                    break
        if docs_table is None and len(cloned_elements) > 2:
            # fallback: first table in cloned elements
            for el in cloned_elements:
                if el.tag.endswith('tbl'):
                    docs_table = docx.table.Table(el, doc)
                    break

        if docs_table is not None:
            ensure_table_borders(docs_table)
            adjust_table_widths(docs_table, 950)
            while len(docs_table.rows) > 1 and not clean_text(docs_table.rows[-1].cells[1].text):
                docs_table._tbl.remove(docs_table.rows[-1]._tr)
            for r in docs_table.rows[1:]:
                clean_cell_0_formatting(r.cells[0], str(item_counter), 950)
                item_counter += 1
            for item in docs_items:
                func_str = item.get('kurztext', '')
                fill_table_row(docs_table, len(docs_table.rows), str(item_counter), func_str, checkbox_tr, 950, is_new=item.get('is_new', False))
                item_counter += 1
            format_table_pagination(docs_table)

        # ---------------------------------------------
        # POPULATE 2.2 SAFETY TABLE
        # ---------------------------------------------
        safety_table = None
        for el in cloned_elements:
            if el.tag.endswith('tbl'):
                t_cand = docx.table.Table(el, doc)
                if len(t_cand.rows) > 1 and 'protective guarding' in ' '.join([c.text for c in t_cand.rows[1].cells]).lower():
                    safety_table = t_cand
                    break
        if safety_table is not None:
            ensure_table_borders(safety_table)
            adjust_table_widths(safety_table, 950)
            while len(safety_table.rows) > 1 and not clean_text(safety_table.rows[-1].cells[1].text):
                safety_table._tbl.remove(safety_table.rows[-1]._tr)
            for r in safety_table.rows[1:]:
                clean_cell_0_formatting(r.cells[0], str(item_counter), 950)
                item_counter += 1
            for item in safety_items:
                func_str = item.get('kurztext', '')
                fill_table_row(safety_table, len(safety_table.rows), str(item_counter), func_str, checkbox_tr, 950, is_new=item.get('is_new', False))
                item_counter += 1
            format_table_pagination(safety_table)

        # ---------------------------------------------
        # POPULATE 2.3 MACHINE DETAILS (H3 BLOCKS)
        # ---------------------------------------------
        elements_to_delete = []
        handled_items = []
        handled_h3_names = set()
        
        def table_has_predefined_content(tbl):
            if tbl is None or len(tbl.rows) <= 1:
                return False
            for r in tbl.rows[1:]:
                # Sub-header rows like Optical, Acoustic
                if len(r.cells) >= 4 and r.cells[0].text.strip() and r.cells[0].text.strip() == r.cells[1].text.strip():
                    return True
                c1 = clean_text(r.cells[1].text) if len(r.cells) > 1 else ''
                c2 = clean_text(r.cells[2].text).replace('☐', '').replace('Yes', '').replace('No', '').strip() if len(r.cells) > 2 else ''
                c3 = clean_text(r.cells[3].text) if len(r.cells) > 3 else ''
                if c1 or c2 or c3:
                    return True
            return False

        for h3_p, block_elements, tbl in h3_blocks:
            h3_title = clean_text(h3_p.text)
            
            # If title is empty, remove it completely
            if not h3_title:
                elements_to_delete.extend(block_elements)
                continue
                
            handled_h3_names.add(h3_title.lower())
            matched_items = [it for it in mach_items if match_h3_name(it.get('level_3', ''), h3_title)]
            has_default_content = (tbl is not None and table_has_predefined_content(tbl)) or (tbl is None)
            
            # Rule: Retain sections that already have predefined content in template + sections with data from Excel
            if not matched_items and not has_default_content:
                # Category has NO data in Excel AND no predefined content in template: Remove it!
                elements_to_delete.extend(block_elements)
            else:
                handled_items.extend(matched_items)
                if tbl is not None:
                    ensure_table_borders(tbl)
                    adjust_table_widths(tbl, 950)
                    
                    if not matched_items and has_default_content:
                        # Keep template content as is, but renumber rows cleanly (1, 2, 3...)
                        for r_idx, r in enumerate(tbl.rows[1:], 1):
                            # Check if row is a subheader like Optical / Acoustic
                            is_sub_header = (len(r.cells) >= 4 and r.cells[0].text.strip() and r.cells[0].text.strip() == r.cells[1].text.strip())
                            has_checkbox = any('Yes' in c.text for c in r.cells)
                            
                            if has_checkbox and not is_sub_header:
                                # For vertically merged cell 0 (e.g. Green rows in Visual signals), only assign number once
                                is_duplicate_green_row = (h3_title == 'Visual / audible signals' and r_idx == 4)
                                if is_duplicate_green_row:
                                    clean_cell_0_formatting(r.cells[0], str(item_counter - 1), 950)
                                else:
                                    clean_cell_0_formatting(r.cells[0], str(item_counter), 950)
                                    item_counter += 1
                            elif not has_checkbox:
                                # Remove any numPr from subheader rows
                                for np in r.cells[0]._tc.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr'):
                                    np.getparent().remove(np)
                    else:
                        # Items exist from Excel
                        if not has_default_content:
                            # Clear empty placeholder rows
                            while len(tbl.rows) > 1:
                                tbl._tbl.remove(tbl.rows[-1]._tr)
                        else:
                            # Keep predefined rows and renumber them, then append Excel items
                            for r in tbl.rows[1:]:
                                is_sub_header = (len(r.cells) >= 4 and r.cells[0].text.strip() and r.cells[0].text.strip() == r.cells[1].text.strip())
                                has_checkbox = any('Yes' in c.text for c in r.cells)
                                if has_checkbox and not is_sub_header:
                                    clean_cell_0_formatting(r.cells[0], str(item_counter), 950)
                                    item_counter += 1
                                    
                        # Populate rows with items from Excel
                        for item in matched_items:
                            func_str = item.get('kurztext', '')
                            fill_table_row(tbl, len(tbl.rows), str(item_counter), func_str, checkbox_tr, 950, is_new=item.get('is_new', False))
                            item_counter += 1

                    format_table_pagination(tbl)

        # Delete all unused H3 blocks
        for del_el in elements_to_delete:
            if del_el in body:
                body.remove(del_el)

        # -------------------------------------------------------------
        # DYNAMIC ADDITION OF MISSING LEVEL 3 CATEGORIES FROM EXCEL
        # -------------------------------------------------------------
        unhandled_items = [it for it in mach_items if it not in handled_items]
        new_l3_groups = {}
        for it in unhandled_items:
            l3_name = clean_text(it.get('level_3', ''))
            if l3_name and l3_name.lower() not in ['none', 'nan', '', 'no']:
                if any(match_h3_name(l3_name, h_name) for h_name in handled_h3_names):
                    continue
                new_l3_groups.setdefault(l3_name, []).append(it)

        def l3_sort_key(name):
            for idx, std in enumerate(STANDARD_H3_LIST):
                if match_h3_name(name, std):
                    return idx
            return 1000

        sorted_new_l3 = sorted(new_l3_groups.keys(), key=l3_sort_key)

        for cat_name in sorted_new_l3:
            cat_items = new_l3_groups[cat_name]

            # 1. Heading 3 paragraph
            if sample_h3_elem is not None:
                new_h3_p = docx.text.paragraph.Paragraph(copy.deepcopy(sample_h3_elem), doc)
            else:
                new_h3_p = doc.add_paragraph(style='Heading 3')
            new_h3_p.text = cat_name
            # Ensure Heading 3 stays with table so it doesn't break onto separate pages
            new_h3_p.paragraph_format.keep_with_next = True
            # Remove any hardcoded numPr so Word auto-numbers seamlessly (e.g. 2.3.2, 3.3.3)
            for np in new_h3_p._p.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr'):
                np.getparent().remove(np)
            body.insert(body.index(test_runs_el), new_h3_p._p)

            # 2. Cloned 4-column Table with guaranteed borders
            if sample_tbl_elem is not None:
                new_tbl = docx.table.Table(copy.deepcopy(sample_tbl_elem), doc)
                while len(new_tbl.rows) > 1:
                    new_tbl._tbl.remove(new_tbl.rows[-1]._tr)
            else:
                new_tbl = doc.add_table(rows=1, cols=4)
                
            new_tbl.rows[0].cells[0].text = 'No.'
            new_tbl.rows[0].cells[1].text = 'Function / Installation /  Document'
            new_tbl.rows[0].cells[2].text = 'Accepted ?'
            new_tbl.rows[0].cells[3].text = 'Remark / Deviation'
            ensure_table_borders(new_tbl)
            adjust_table_widths(new_tbl, 950)

            for it in cat_items:
                fill_table_row(new_tbl, len(new_tbl.rows), str(item_counter), it.get('kurztext', ''), checkbox_tr, 950, is_new=it.get('is_new', False))
                item_counter += 1

            format_table_pagination(new_tbl)
            body.insert(body.index(test_runs_el), new_tbl._tbl)

            # 3. Spacer paragraph
            if sample_spacer_elem is not None:
                sp_p = docx.text.paragraph.Paragraph(copy.deepcopy(sample_spacer_elem), doc)
                sp_p.text = ''
                body.insert(body.index(test_runs_el), sp_p._p)

    # Force Word to auto-update Table of Contents and fields when opening the document
    try:
        w_update_fields = docx.oxml.OxmlElement('w:updateFields')
        w_update_fields.set(docx.oxml.ns.qn('w:val'), 'true')
        doc.settings.element.append(w_update_fields)
    except Exception:
        pass

    # Save output document
    doc.save(output_path)
    return output_path
