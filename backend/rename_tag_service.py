import os
import io
import re
import datetime
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import docx

def parse_user_date(date_str):
    """
    Parses various date format inputs (e.g. '22-Sep-26', '15-Sep-2026', '2026-09-15', '15/09/2026').
    Returns (iso_date_str '2026-09-22', doc_date_str '22-Sep-26').
    """
    if not date_str:
        now = datetime.datetime.now()
        return now.strftime("%Y-%m-%d"), now.strftime("%d-%b-%y")

    date_str = str(date_str).strip()

    formats = [
        "%d-%b-%Y", "%d-%b-%y", "%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y",
        "%d.%m.%Y", "%d.%m.%y", "%d %b %Y", "%d %b %y", "%B %d, %Y"
    ]

    parsed_dt = None
    for fmt in formats:
        try:
            parsed_dt = datetime.datetime.strptime(date_str, fmt)
            break
        except Exception:
            pass

    if not parsed_dt:
        m = re.search(r'(\d{1,2})[-/\s.]([A-Za-z]{3,9})[-/\s.](\d{2,4})', date_str)
        if m:
            day, month_str, year = m.groups()
            if len(year) == 4:
                year = year[2:]
            try:
                parsed_dt = datetime.datetime.strptime(f"{day}-{month_str[:3]}-{year}", "%d-%b-%y")
            except Exception:
                pass

    if not parsed_dt:
        now = datetime.datetime.now()
        return now.strftime("%Y-%m-%d"), date_str

    iso_str = parsed_dt.strftime("%Y-%m-%d")
    doc_str = parsed_dt.strftime("%d-%b-%y")
    return iso_str, doc_str

def convert_doc_to_docx_bytes(doc_bytes, temp_prefix="temp_doc"):
    """
    Converts legacy Word .doc bytes to modern .docx bytes using win32com (MS Word COM automation).
    """
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()

        temp_dir = os.path.abspath("scratch")
        os.makedirs(temp_dir, exist_ok=True)

        doc_path = os.path.join(temp_dir, f"{temp_prefix}_{os.getpid()}_{int(datetime.datetime.now().timestamp() * 1000)}.doc")
        docx_path = doc_path + "x"

        with open(doc_path, "wb") as f:
            f.write(doc_bytes)

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(doc_path)
        doc.SaveAs2(docx_path, FileFormat=16) # 16 = wdFormatXMLDocument (.docx)
        doc.Close()
        word.Quit()

        if os.path.exists(docx_path):
            with open(docx_path, "rb") as f:
                docx_bytes = f.read()
        else:
            docx_bytes = doc_bytes

        try:
            if os.path.exists(doc_path):
                os.remove(doc_path)
            if os.path.exists(docx_path):
                os.remove(docx_path)
        except Exception:
            pass

        return docx_bytes
    except Exception:
        return doc_bytes

def update_document_history_table(doc, doc_date_str, update_all_rows=False):
    """
    Scans tables in Word document for Document History table (where header row contains 'VERSION' and 'DATE'),
    and updates the Date cell in data rows to doc_date_str (dd-MMM-yy format).
    """
    for tbl in doc.tables:
        if not tbl.rows:
            continue
        col_ver = -1
        col_date = -1
        for idx, cell in enumerate(tbl.rows[0].cells):
            c_text = cell.text.strip().upper()
            if "VERSION" in c_text:
                col_ver = idx
            if "DATE" in c_text:
                col_date = idx

        if col_ver >= 0 and col_date >= 0:
            if len(tbl.rows) >= 2:
                if update_all_rows:
                    for r_idx in range(1, len(tbl.rows)):
                        tbl.rows[r_idx].cells[col_date].text = doc_date_str
                else:
                    tbl.rows[1].cells[col_date].text = doc_date_str
            return True
    return False

def update_docx_custom_properties(docx_bytes, properties_dict):
    """
    Updates Custom Document Properties in docProps/custom.xml,
    forces Word field updates via word/settings.xml (<w:updateFields w:val="true"/>),
    and updates cached XML text for all DOCPROPERTY fields in headers, footers, and body.
    """
    try:
        in_io = io.BytesIO(docx_bytes)
        out_io = io.BytesIO()

        ns_custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        ns_vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
        ns_w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

        ET.register_namespace('', ns_custom)
        ET.register_namespace('vt', ns_vt)
        ET.register_namespace('w', ns_w)

        custom_xml_path = 'docProps/custom.xml'

        def _update_xml_docproperty_fields(xml_content, props):
            try:
                root = ET.fromstring(xml_content)
                
                # Simple fields
                for fld in root.findall(f'.//{{{ns_w}}}fldSimple'):
                    instr = fld.get(f'{{{ns_w}}}instr', '')
                    for prop_name, prop_val in props.items():
                        if 'DOCPROPERTY' in instr and prop_name.lower() in instr.lower():
                            for t in fld.findall(f'.//{{{ns_w}}}t'):
                                t.text = str(prop_val)

                # Complex fields
                for p in root.findall(f'.//{{{ns_w}}}p'):
                    current_prop_val = None
                    in_separate = False

                    for child in list(p):
                        fldChar = child.find(f'.//{{{ns_w}}}fldChar')
                        if fldChar is not None:
                            cType = fldChar.get(f'{{{ns_w}}}fldCharType')
                            if cType == 'begin':
                                current_prop_val = None
                                in_separate = False
                            elif cType == 'separate':
                                in_separate = True
                            elif cType == 'end':
                                current_prop_val = None
                                in_separate = False

                        instrText = child.find(f'.//{{{ns_w}}}instrText')
                        if instrText is None and child.tag == f'{{{ns_w}}}instrText':
                            instrText = child

                        if instrText is not None and instrText.text and 'DOCPROPERTY' in instrText.text:
                            txt = instrText.text
                            for prop_name, prop_val in props.items():
                                if prop_name.lower() in txt.lower():
                                    current_prop_val = prop_val
                                    break

                        if in_separate and current_prop_val is not None:
                            for t in child.findall(f'.//{{{ns_w}}}t'):
                                t.text = str(current_prop_val)

                return ET.tostring(root, encoding='utf-8', xml_declaration=True)
            except Exception:
                return xml_content

        with zipfile.ZipFile(in_io, 'r') as zin, zipfile.ZipFile(out_io, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)

                # 1. docProps/custom.xml
                if item.filename == custom_xml_path:
                    try:
                        root = ET.fromstring(content)
                        existing_props = {}
                        max_pid = 1
                        for child in root.findall(f'{{{ns_custom}}}property'):
                            pname = child.get('name')
                            pid_str = child.get('pid', '1')
                            try:
                                pid = int(pid_str)
                                if pid > max_pid:
                                    max_pid = pid
                            except Exception:
                                pass
                            if pname:
                                existing_props[pname] = child

                        for key, val in properties_dict.items():
                            val_str = str(val or '')
                            if key in existing_props:
                                p_elem = existing_props[key]
                                vt_child = p_elem.find(f'{{{ns_vt}}}lpwstr')
                                if vt_child is not None:
                                    vt_child.text = val_str
                                else:
                                    for c in list(p_elem):
                                        p_elem.remove(c)
                                    new_vt = ET.SubElement(p_elem, f'{{{ns_vt}}}lpwstr')
                                    new_vt.text = val_str
                            else:
                                max_pid += 1
                                p_elem = ET.SubElement(root, f'{{{ns_custom}}}property', {
                                    'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                                    'pid': str(max_pid),
                                    'name': key
                                })
                                vt_child = ET.SubElement(p_elem, f'{{{ns_vt}}}lpwstr')
                                vt_child.text = val_str

                        content = ET.tostring(root, encoding='utf-8', xml_declaration=True)
                    except Exception:
                        pass

                # 2. word/settings.xml -> enable updateFields
                elif item.filename == 'word/settings.xml':
                    try:
                        s = content.decode('utf-8', errors='ignore')
                        if '<w:updateFields' in s:
                            s = re.sub(r'<w:updateFields[^>]*/>', '<w:updateFields w:val="true"/>', s)
                        else:
                            s = s.replace('<w:settings ', '<w:settings><w:updateFields w:val="true"/> ', 1)
                        content = s.encode('utf-8')
                    except Exception:
                        pass

                # 3. word/*.xml (document, headers, footers) -> update DOCPROPERTY cached field texts
                elif item.filename.startswith('word/') and item.filename.endswith('.xml'):
                    content = _update_xml_docproperty_fields(content, properties_dict)

                zout.writestr(item, content)

        out_io.seek(0)
        return out_io.getvalue()
    except Exception:
        return docx_bytes

def replace_text_in_doc(doc, replacements):
    """
    Replaces text tokens in paragraphs, tables, headers, and footers.
    replacements = {'XXXXX': '56021', '5XXXX': '56021', 'DD-MMM-YYYY': '15-Sep-2026', ...}
    """
    def _replace_in_p(p):
        full_text = p.text
        if not full_text:
            return
        needs_replace = any(k in full_text for k in replacements)
        if not needs_replace:
            return

        for k, v in replacements.items():
            if k in full_text:
                full_text = full_text.replace(k, v)

        if len(p.runs) > 0:
            p.runs[0].text = full_text
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.text = full_text

    for p in doc.paragraphs:
        _replace_in_p(p)

    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_p(p)

    for section in doc.sections:
        if section.header:
            for p in section.header.paragraphs:
                _replace_in_p(p)
            for tbl in section.header.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            _replace_in_p(p)
        if section.footer:
            for p in section.footer.paragraphs:
                _replace_in_p(p)
            for tbl in section.footer.tables:
                for row in tbl.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            _replace_in_p(p)

def compute_renamed_filename(orig_filename, order_val, iso_date_str):
    name, ext = os.path.splitext(orig_filename)

    # Force extension to .docx if .doc (since converted to docx)
    if ext.lower() == '.doc':
        ext = '.docx'

    # Replace XXXXX or 5XXXX or 5XXXX_ or XXXXX_ with Order value
    if order_val:
        name = re.sub(r'5?XXXXX?', str(order_val), name, flags=re.IGNORECASE)

    # Replace date pattern YYYY-MM-DD, 2019-MM-DD, 20XX-XX-XX, 2024-02-06, etc.
    if iso_date_str:
        name = re.sub(r'20[0-9X]{2}-[a-zA-Z0-9]{2}-[a-zA-Z0-9]{2}', iso_date_str, name, flags=re.IGNORECASE)
        name = re.sub(r'YYYY-MM-DD', iso_date_str, name, flags=re.IGNORECASE)
        name = re.sub(r'DD-MMM-YYYY', iso_date_str, name, flags=re.IGNORECASE)

    return name + ext

def process_single_docx(
    docx_bytes,
    orig_filename,
    copyright_val,
    version_val,
    machine_val,
    order_val,
    baunummer_val,
    date_val,
    bezeichnung_val="Cartoning machine, Tube filling machine, Filling platform"
):
    """
    Processes a single Word document:
    1. Converts legacy .doc to .docx if needed.
    2. Updates Custom Document Properties (Copyright, Version, Machine, Order, Baunummer, Bezeichnung).
    3. Replaces text tokens (DD-MMM-YYYY, XXXXX, 5XXXX) in document body, tables, headers, footers.
    4. Renames the file according to Order and Date (ISO format).
    Returns dict containing processed bytes and new file name.
    """
    iso_date_str, doc_date_str = parse_user_date(date_val)

    # 0. Check if legacy .doc file or not zip format -> Convert to .docx using win32com if needed
    is_legacy_doc = orig_filename.lower().endswith('.doc') or not docx_bytes.startswith(b'PK\x03\x04')
    if is_legacy_doc:
        docx_bytes = convert_doc_to_docx_bytes(docx_bytes, temp_prefix="doc_conv")

    # 1. Update XML Custom Properties
    custom_props = {
        'Copyright': copyright_val or '',
        'Version': version_val or '',
        'Machine': machine_val or '',
        'Order': order_val or '',
        'Baunummer': baunummer_val or '',
        'Bezeichnung': bezeichnung_val or ''
    }
    step1_bytes = update_docx_custom_properties(docx_bytes, custom_props)

    # 2. Update Document Text / Tables / Headers
    try:
        doc_io = io.BytesIO(step1_bytes)
        doc = docx.Document(doc_io)

        replacements = {
            'DD-MMM-YYYY': doc_date_str,
            'DD-Mmm-YYYY': doc_date_str,
            'DD.MM.YYYY': doc_date_str,
            '20XX-XX-XX': iso_date_str,
        }
        if order_val:
            replacements['XXXXX'] = str(order_val)
            replacements['5XXXX'] = str(order_val)

        # Update Document History table date cell(s) to dd-MMM-yy format
        update_document_history_table(doc, doc_date_str, update_all_rows=False)

        replace_text_in_doc(doc, replacements)

        step2_io = io.BytesIO()
        doc.save(step2_io)
        step2_io.seek(0)
        processed_bytes = step2_io.getvalue()
    except Exception:
        processed_bytes = step1_bytes

    # 3. Compute New Filename
    new_filename = compute_renamed_filename(orig_filename, order_val, iso_date_str)

    return {
        "original_filename": orig_filename,
        "new_filename": new_filename,
        "processed_bytes": processed_bytes,
        "iso_date": iso_date_str,
        "doc_date": doc_date_str,
        "order": order_val
    }

def get_workspace_word_files(*args, **kwargs):
    """
    Scans specified folder location for all Word template files (.docx, .doc).
    Supports both relative and absolute folder paths.
    Usage:
      get_workspace_word_files(base_dir="IQOQDQ", recursive=True)
      get_workspace_word_files("IQOQDQ", True)
      get_workspace_word_files("IQOQDQ")
      get_workspace_word_files()
    """
    base_dir = "IQOQDQ"
    recursive = True

    if len(args) > 0 and args[0] is not None:
        base_dir = str(args[0])
    elif "base_dir" in kwargs and kwargs["base_dir"] is not None:
        base_dir = str(kwargs["base_dir"])
    elif "target_folder_path" in kwargs and kwargs["target_folder_path"] is not None:
        base_dir = str(kwargs["target_folder_path"])

    if len(args) > 1:
        recursive = bool(args[1])
    elif "recursive" in kwargs:
        recursive = bool(kwargs["recursive"])

    if not base_dir or not base_dir.strip():
        base_dir = "IQOQDQ"

    base_abs = os.path.abspath(base_dir.strip())
    found_files = []
    if not os.path.exists(base_abs) or not os.path.isdir(base_abs):
        return found_files

    if recursive:
        for root, dirs, files in os.walk(base_abs):
            for f in files:
                if not f.startswith("~$") and f.lower().endswith(('.doc', '.docx')):
                    rel_p = os.path.relpath(os.path.join(root, f), base_abs)
                    found_files.append({
                        "rel_path": rel_p,
                        "filename": f,
                        "full_path": os.path.join(root, f)
                    })
    else:
        for f in os.listdir(base_abs):
            full_p = os.path.join(base_abs, f)
            if os.path.isfile(full_p) and not f.startswith("~$") and f.lower().endswith(('.doc', '.docx')):
                found_files.append({
                    "rel_path": f,
                    "filename": f,
                    "full_path": full_p
                })

    return found_files
