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

def sanitize_footer_filename_in_docx_bytes(docx_bytes, target_filename_no_ext):
    """
    Sanitizes docx_bytes by replacing any temporary file names (e.g. in_17804_1790129252259,
    out_17804_..., norm_..., temp_...) in header/footer/document XML with target_filename_no_ext.
    """
    if not docx_bytes or not docx_bytes.startswith(b'PK\x03\x04') or not target_filename_no_ext:
        return docx_bytes

    try:
        in_io = io.BytesIO(docx_bytes)
        out_io = io.BytesIO()

        temp_name_pattern = re.compile(
            r'\b(?:in|out|norm|temp|doc_conv|work)_[0-9]+_[0-9]+(?:_out)?(?:\.docx?)?\b',
            re.IGNORECASE
        )

        with zipfile.ZipFile(in_io, 'r') as zin, zipfile.ZipFile(out_io, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                if item.filename.startswith('word/') or item.filename.startswith('docProps/'):
                    try:
                        text = content.decode('utf-8', errors='ignore')
                        if temp_name_pattern.search(text):
                            text = temp_name_pattern.sub(target_filename_no_ext, text)
                            content = text.encode('utf-8')
                    except Exception:
                        pass
                zout.writestr(item, content)

        out_io.seek(0)
        return out_io.getvalue()
    except Exception:
        return docx_bytes

def normalize_docx_bytes_via_word_saveas(docx_bytes, target_filename=None):
    """
    Opens docx_bytes in MS Word in background and performs a native Word SaveAs (wdFormatXMLDocument = 16)
    to normalize XML structures, repair any minor schema discrepancies, and output 100% official Microsoft Word .docx bytes.
    Uses target_filename if provided to ensure Word fields like { FILENAME } populate with the actual target filename.
    """
    if not docx_bytes:
        return docx_bytes

    try:
        import win32com.client
        import pythoncom
        import shutil
        pythoncom.CoInitialize()

        temp_dir = os.path.abspath("scratch")
        os.makedirs(temp_dir, exist_ok=True)

        target_base = os.path.splitext(target_filename)[0] if target_filename else f"doc_{os.getpid()}_{int(datetime.datetime.now().timestamp()*1000)}"
        sub_dir = os.path.join(temp_dir, f"norm_{os.getpid()}_{int(datetime.datetime.now().timestamp()*1000)}")
        os.makedirs(sub_dir, exist_ok=True)

        temp_out_path = os.path.join(sub_dir, f"{target_base}.docx")

        with open(temp_out_path, "wb") as f:
            f.write(docx_bytes)

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone

        doc = word.Documents.Open(temp_out_path)
        doc.SaveAs2(temp_out_path, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
        try:
            doc.Fields.Update()
            for sec in doc.Sections:
                for hdr in sec.Headers:
                    hdr.Range.Fields.Update()
                for ftr in sec.Footers:
                    ftr.Range.Fields.Update()
        except Exception:
            pass

        doc.Save()
        doc.Close(False)
        word.Quit()

        if os.path.exists(temp_out_path):
            with open(temp_out_path, "rb") as f:
                norm_bytes = f.read()
        else:
            norm_bytes = docx_bytes

        try:
            shutil.rmtree(sub_dir, ignore_errors=True)
        except Exception:
            pass

        if target_filename:
            norm_bytes = sanitize_footer_filename_in_docx_bytes(norm_bytes, target_base)

        return norm_bytes
    except Exception:
        return docx_bytes

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
    and forces Word field updates via word/settings.xml (<w:updateFields w:val="true"/>).
    Only modifies docProps/custom.xml and word/settings.xml (does NOT alter word/document.xml namespaces).
    """
    try:
        in_io = io.BytesIO(docx_bytes)
        out_io = io.BytesIO()

        ns_custom = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
        ns_vt = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

        ET.register_namespace('', ns_custom)
        ET.register_namespace('vt', ns_vt)

        custom_xml_path = 'docProps/custom.xml'

        # Synonym map so we don't create duplicate properties in XML
        synonym_map = {
            'Serial no.': ['Baunummer', 'Serial no.', 'Serial Number', 'SerialNo'],
            'Designation': ['Bezeichnung', 'Designation', 'Description'],
            'Baunummer': ['Baunummer', 'Serial no.', 'Serial Number', 'SerialNo'],
            'Bezeichnung': ['Bezeichnung', 'Designation', 'Description']
        }

        with zipfile.ZipFile(in_io, 'r') as zin, zipfile.ZipFile(out_io, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)

                # 1. docProps/custom.xml
                if item.filename == custom_xml_path:
                    try:
                        xml_str = content.decode('utf-8', errors='ignore')
                        root = ET.fromstring(xml_str)
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
                            
                            # Find matching target property element in existing XML
                            target_prop_elem = None
                            if key in existing_props:
                                target_prop_elem = existing_props[key]
                            elif key in synonym_map:
                                for syn in synonym_map[key]:
                                    if syn in existing_props:
                                        target_prop_elem = existing_props[syn]
                                        break

                            if target_prop_elem is not None:
                                vt_child = target_prop_elem.find(f'{{{ns_vt}}}lpwstr')
                                if vt_child is not None:
                                    vt_child.text = val_str
                                else:
                                    for c in list(target_prop_elem):
                                        target_prop_elem.remove(c)
                                    new_vt = ET.SubElement(target_prop_elem, f'{{{ns_vt}}}lpwstr')
                                    new_vt.text = val_str
                            else:
                                is_synonym_already_added = False
                                if key in synonym_map:
                                    for syn in synonym_map[key]:
                                        if syn in existing_props:
                                            is_synonym_already_added = True
                                            break
                                if not is_synonym_already_added:
                                    max_pid += 1
                                    p_elem = ET.SubElement(root, f'{{{ns_custom}}}property', {
                                        'fmtid': '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}',
                                        'pid': str(max_pid),
                                        'name': key
                                    })
                                    new_vt = ET.SubElement(p_elem, f'{{{ns_vt}}}lpwstr')
                                    new_vt.text = val_str
                                    existing_props[key] = p_elem

                        raw_xml_bytes = ET.tostring(root, encoding='utf-8')
                        xml_header = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                        content = xml_header + raw_xml_bytes
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

    # Replace any date pattern YYYY-DD-MM, YYYY-MM-DD, YYYY-XX-XX, 20XX-XX-XX, 20XX-MM-DD, 2024-02-06, 2019-MM-DD, etc.
    if iso_date_str:
        name = re.sub(r'(?:20\d{2}|20XX|YYYY)-(?:MM|DD|XX|\d{2})-(?:MM|DD|XX|\d{2})', iso_date_str, name, flags=re.IGNORECASE)
        name = re.sub(r'(?:DD|\d{2})-(?:MMM|Mmm|MM|XX|\d{2})-(?:YYYY|YY|20XX|\d{4})', iso_date_str, name, flags=re.IGNORECASE)

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
    Processes a single Word document cleanly in memory, then performs a Word SaveAs (.docx) pass.
    """
    iso_date_str, doc_date_str = parse_user_date(date_val)
    new_filename = compute_renamed_filename(orig_filename, order_val, iso_date_str)
    new_filename_no_ext = os.path.splitext(new_filename)[0]

    # Check if legacy .doc file -> convert to docx bytes
    is_legacy_doc = orig_filename.lower().endswith('.doc') or not docx_bytes.startswith(b'PK\x03\x04')
    if is_legacy_doc:
        docx_bytes = convert_doc_to_docx_bytes(docx_bytes, temp_prefix="doc_conv")

    # Update XML Custom Properties cleanly in memory
    custom_props = {
        'Copyright': copyright_val or '',
        'Version': version_val or '',
        'Machine': machine_val or '',
        'Order': order_val or '',
        'Baunummer': baunummer_val or '',
        'Serial no.': baunummer_val or '',
        'Bezeichnung': bezeichnung_val or '',
        'Designation': bezeichnung_val or ''
    }
    step1_bytes = update_docx_custom_properties(docx_bytes, custom_props)

    # Update Document Text / Tables / Headers cleanly in memory
    try:
        doc_io = io.BytesIO(step1_bytes)
        doc = docx.Document(doc_io)

        replacements = {
            'DD-MMM-YYYY': doc_date_str,
            'DD-Mmm-YYYY': doc_date_str,
            'DD.MM.YYYY': doc_date_str,
            '20XX-XX-XX': iso_date_str,
            'YYYY-DD-MM': iso_date_str,
            'YYYY-MM-DD': iso_date_str,
            'YYYY-XX-XX': iso_date_str,
            '20XX-MM-DD': iso_date_str,
            '20XX-DD-MM': iso_date_str,
        }
        if order_val:
            replacements['XXXXX'] = str(order_val)
            replacements['5XXXX'] = str(order_val)

        orig_fn_no_ext = os.path.splitext(orig_filename)[0]
        if orig_fn_no_ext and new_filename_no_ext and orig_fn_no_ext != new_filename_no_ext:
            replacements[orig_fn_no_ext] = new_filename_no_ext

        update_document_history_table(doc, doc_date_str, update_all_rows=False)
        replace_text_in_doc(doc, replacements)

        step2_io = io.BytesIO()
        doc.save(step2_io)
        step2_io.seek(0)
        processed_bytes = step2_io.getvalue()
    except Exception:
        processed_bytes = step1_bytes

    # Word SaveAs (.docx) normalization pass
    final_docx_bytes = normalize_docx_bytes_via_word_saveas(processed_bytes, target_filename=new_filename)
    final_docx_bytes = sanitize_footer_filename_in_docx_bytes(final_docx_bytes, new_filename_no_ext)

    return {
        "original_filename": orig_filename,
        "new_filename": new_filename,
        "processed_bytes": final_docx_bytes,
        "iso_date": iso_date_str,
        "doc_date": doc_date_str,
        "order": order_val
    }

def process_batch_word_files(
    file_items,
    copyright_val,
    version_val,
    machine_val,
    order_val,
    baunummer_val,
    date_val,
    bezeichnung_val="Cartoning machine, Tube filling machine, Filling platform"
):
    """
    High-speed batch processing of Word documents using 1 single persistent MS Word COM instance.
    Processes 50+ files in seconds with zero startup/shutdown overhead.
    Ensures footers/headers and FILENAME fields display the target renamed filename cleanly.
    """
    iso_date_str, doc_date_str = parse_user_date(date_val)
    results = []

    # 1. High-speed single-instance MS Word COM batch processor
    try:
        import win32com.client
        import pythoncom
        import shutil
        pythoncom.CoInitialize()

        temp_dir = os.path.abspath("scratch")
        os.makedirs(temp_dir, exist_ok=True)

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone
        try:
            word.ScreenUpdating = False
        except Exception:
            pass

        try:
            for idx, item in enumerate(file_items):
                orig_filename = item["filename"]
                docx_bytes = item["bytes"]

                new_fn = compute_renamed_filename(orig_filename, order_val, iso_date_str)
                new_fn_no_ext = os.path.splitext(new_fn)[0]
                orig_fn_no_ext = os.path.splitext(orig_filename)[0]

                is_doc = orig_filename.lower().endswith('.doc')
                timestamp_str = f"{os.getpid()}_{int(datetime.datetime.now().timestamp()*1000)}_{idx}"

                sub_temp_dir = os.path.join(temp_dir, f"work_{timestamp_str}")
                os.makedirs(sub_temp_dir, exist_ok=True)

                temp_in_path = os.path.join(sub_temp_dir, orig_filename)
                temp_out_path = os.path.join(sub_temp_dir, new_fn)

                with open(temp_in_path, "wb") as f:
                    f.write(docx_bytes)

                try:
                    doc = word.Documents.Open(temp_in_path)

                    props_to_set = {
                        'Copyright': str(copyright_val or ''),
                        'Version': str(version_val or ''),
                        'Machine': str(machine_val or ''),
                        'Order': str(order_val or ''),
                        'Baunummer': str(baunummer_val or ''),
                        'Serial no.': str(baunummer_val or ''),
                        'Bezeichnung': str(bezeichnung_val or ''),
                        'Designation': str(bezeichnung_val or '')
                    }

                    for prop_name, prop_val in props_to_set.items():
                        val_s = str(prop_val)
                        try:
                            doc.CustomDocumentProperties(prop_name).Value = val_s
                        except Exception:
                            try:
                                doc.CustomDocumentProperties.Add(prop_name, False, 1, val_s)
                            except Exception:
                                pass

                    replacements = {
                        'DD-MMM-YYYY': doc_date_str,
                        'DD-Mmm-YYYY': doc_date_str,
                        'DD.MM.YYYY': doc_date_str,
                        '20XX-XX-XX': iso_date_str,
                        'YYYY-DD-MM': iso_date_str,
                        'YYYY-MM-DD': iso_date_str,
                        'YYYY-XX-XX': iso_date_str,
                        '20XX-MM-DD': iso_date_str,
                        '20XX-DD-MM': iso_date_str,
                    }
                    if order_val:
                        replacements['XXXXX'] = str(order_val)
                        replacements['5XXXX'] = str(order_val)
                    if orig_fn_no_ext and new_fn_no_ext and orig_fn_no_ext != new_fn_no_ext:
                        replacements[orig_fn_no_ext] = new_fn_no_ext

                    for search_txt, replace_txt in replacements.items():
                        try:
                            for story in doc.StoryRanges:
                                find_obj = story.Find
                                find_obj.ClearFormatting()
                                find_obj.Replacement.ClearFormatting()
                                find_obj.Execute(FindText=search_txt, ReplaceWith=replace_txt, Replace=2, Forward=True, Wrap=1)
                        except Exception:
                            pass

                    try:
                        for tbl in doc.Tables:
                            if tbl.Rows.Count >= 2:
                                is_hist_table = False
                                col_date_idx = -1
                                for c_i in range(1, tbl.Columns.Count + 1):
                                    c_text = tbl.Cell(1, c_i).Range.Text.upper()
                                    if "VERSION" in c_text or "DATE" in c_text:
                                        is_hist_table = True
                                    if "DATE" in c_text:
                                        col_date_idx = c_i
                                if is_hist_table and col_date_idx > 0:
                                    tbl.Cell(2, col_date_idx).Range.Text = doc_date_str
                    except Exception:
                        pass

                    # Save to temp_out_path (named new_fn) FIRST so document name in Word becomes new_fn
                    doc.SaveAs2(temp_out_path, FileFormat=16)

                    # Update fields so FILENAME field evaluates to new_fn_no_ext
                    try:
                        doc.Fields.Update()
                        for sec in doc.Sections:
                            for hdr in sec.Headers:
                                hdr.Range.Fields.Update()
                            for ftr in sec.Footers:
                                ftr.Range.Fields.Update()
                    except Exception:
                        pass

                    doc.Save()
                    doc.Close(False)

                    if os.path.exists(temp_out_path):
                        with open(temp_out_path, "rb") as f:
                            p_bytes = f.read()
                    else:
                        p_bytes = docx_bytes

                    p_bytes = sanitize_footer_filename_in_docx_bytes(p_bytes, new_fn_no_ext)

                    results.append({
                        "original_filename": orig_filename,
                        "new_filename": new_fn,
                        "processed_bytes": p_bytes,
                        "source_path": item.get("source_path"),
                        "iso_date": iso_date_str,
                        "doc_date": doc_date_str,
                        "order": order_val
                    })
                except Exception:
                    # Single file fallback
                    res_single = process_single_docx(docx_bytes, orig_filename, copyright_val, version_val, machine_val, order_val, baunummer_val, date_val, bezeichnung_val)
                    res_single["source_path"] = item.get("source_path")
                    results.append(res_single)
                finally:
                    try:
                        shutil.rmtree(sub_temp_dir, ignore_errors=True)
                    except Exception:
                        pass
        finally:
            try:
                word.Quit()
            except Exception:
                pass

        if len(results) == len(file_items):
            return results
    except Exception:
        pass

    # 2. In-memory fast fallback if win32com is unavailable
    results = []
    for item in file_items:
        res = process_single_docx(
            docx_bytes=item["bytes"],
            orig_filename=item["filename"],
            copyright_val=copyright_val,
            version_val=version_val,
            machine_val=machine_val,
            order_val=order_val,
            baunummer_val=baunummer_val,
            date_val=date_val,
            bezeichnung_val=bezeichnung_val
        )
        res["source_path"] = item.get("source_path")
        results.append(res)
    return results

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
