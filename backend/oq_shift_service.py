import os
import io
import re
import copy
import shutil
import datetime
import pandas as pd
import openpyxl
import docx
from docx.shared import Pt
from docx.enum.text import WD_BREAK

DEFAULT_TEMPLATE_PATH = os.path.abspath(r"IQOQDQ/OQ_Shift/XXXXX_11_OQ_Shift Register_2025-01-02_en.doc")
DEFAULT_SAMPLE_EXCEL_PATH = os.path.abspath(r"IQOQDQ/OQ_Shift/5XXXX-ShiftRegisterInfo.xlsx")
GMP_SHIFT_BASE = os.path.abspath(r"IQOQDQ/OQ_Shift/GMP_Shift Register")
DOCX_CACHE_BASE = os.path.abspath(r"IQOQDQ/OQ_Shift/.docx_cache")

def get_available_machine_types():
    """
    Returns list of available machine types inside GMP_Shift Register directory ONLY.
    """
    if not os.path.exists(GMP_SHIFT_BASE):
        return []
    types = [d for d in os.listdir(GMP_SHIFT_BASE) if os.path.isdir(os.path.join(GMP_SHIFT_BASE, d)) and not d.startswith('.')]
    return sorted(types)

def create_machine_type(machine_name):
    """
    Creates a new machine type directory structure under GMP_Shift and .docx_cache.
    Initializes default Performer sign-off template if available.
    """
    clean_name = re.sub(r'[^A-Za-z0-9_\-\s]', '', str(machine_name).strip())
    if not clean_name:
        raise ValueError("Machine Type name cannot be empty and must contain valid characters (A-Z, 0-9, _, -, space).")
    
    src_mach_dir = os.path.join(GMP_SHIFT_BASE, clean_name)
    if os.path.exists(src_mach_dir):
        raise ValueError(f"Machine Type '{clean_name}' already exists.")
        
    src_dir = os.path.join(src_mach_dir, "EN")
    cache_dir = os.path.join(DOCX_CACHE_BASE, clean_name, "EN")
    
    os.makedirs(src_dir, exist_ok=True)
    os.makedirs(cache_dir, exist_ok=True)
    
    # Copy default Performer template from Alarm or FP if available
    alarm_perf_doc = os.path.abspath(r"IQOQDQ/OQ Alarm/GMP_Alarme/FP/EN/Performer.doc")
    alarm_perf_docx = os.path.abspath(r"IQOQDQ/OQ Alarm/.docx_cache/FP/EN/Performer.docx")
    
    if os.path.exists(alarm_perf_doc):
        shutil.copy2(alarm_perf_doc, os.path.join(src_dir, "Performer.doc"))
    if os.path.exists(alarm_perf_docx):
        shutil.copy2(alarm_perf_docx, os.path.join(src_dir, "Performer.docx"))
        shutil.copy2(alarm_perf_docx, os.path.join(cache_dir, "Performer.docx"))
        
    return clean_name

def get_performer_table_xml(machine_type="FP", lang="EN"):
    """
    Locates and returns the Performer sign-off table XML.
    Uses multi-level fallback across shift cache, alarm cache, and standard templates.
    """
    cache_dir = os.path.join(DOCX_CACHE_BASE, machine_type, lang)
    src_dir = os.path.join(GMP_SHIFT_BASE, machine_type, lang)
    os.makedirs(cache_dir, exist_ok=True)
    performer_cache_path = os.path.join(cache_dir, "Performer.docx")
    
    # 1. Check local cache
    if os.path.exists(performer_cache_path):
        try:
            p_doc = docx.Document(performer_cache_path)
            if p_doc.tables:
                return p_doc.tables[0]._tbl
        except Exception:
            pass

    # 2. Check local source folder
    src_docx = os.path.join(src_dir, "Performer.docx")
    if os.path.exists(src_docx):
        try:
            shutil.copy2(src_docx, performer_cache_path)
            p_doc = docx.Document(performer_cache_path)
            if p_doc.tables:
                return p_doc.tables[0]._tbl
        except Exception:
            pass

    # 3. Search fallbacks from Alarm Performer templates
    alarm_fallbacks = [
        os.path.abspath(r"IQOQDQ/OQ Alarm/.docx_cache/FP/EN/Performer.docx"),
        os.path.abspath(r"IQOQDQ/OQ Alarm/GMP_Alarme/FP/EN/Performer.docx"),
    ]
    for fb_path in alarm_fallbacks:
        if os.path.exists(fb_path):
            try:
                shutil.copy2(fb_path, performer_cache_path)
                if os.path.exists(src_dir):
                    shutil.copy2(fb_path, os.path.join(src_dir, "Performer.docx"))
                p_doc = docx.Document(performer_cache_path)
                if p_doc.tables:
                    return p_doc.tables[0]._tbl
            except Exception:
                pass

    return None

def delete_machine_type(machine_name):
    """
    Safely deletes a machine type directory from GMP_Shift and .docx_cache.
    """
    clean_name = re.sub(r'[^A-Za-z0-9_\-]', '', str(machine_name).strip()).upper()
    if not clean_name:
        raise ValueError("Invalid machine type name.")
        
    src_mach_dir = os.path.join(GMP_SHIFT_BASE, clean_name)
    cache_mach_dir = os.path.join(DOCX_CACHE_BASE, clean_name)
    
    if os.path.exists(src_mach_dir):
        shutil.rmtree(src_mach_dir, ignore_errors=True)
    if os.path.exists(cache_mach_dir):
        shutil.rmtree(cache_mach_dir, ignore_errors=True)
        
    return True

def list_machine_files_detailed(machine_type, lang="EN"):
    """
    Returns a pandas DataFrame of all template files in GMP_Shift/{machine_type}/{lang}/.
    Columns: Filename, File Size (KB), Format, Last Modified, Cached .docx
    """
    src_dir = os.path.join(GMP_SHIFT_BASE, machine_type, lang)
    cache_dir = os.path.join(DOCX_CACHE_BASE, machine_type, lang)
    
    if not os.path.exists(src_dir):
        return pd.DataFrame(columns=["Filename", "Size (KB)", "Format", "Last Modified", "Cached .docx"])
        
    records = []
    for f in sorted(os.listdir(src_dir)):
        if f.startswith('.') or f.startswith('~$'):
            continue
        fp = os.path.join(src_dir, f)
        if not os.path.isfile(fp):
            continue
            
        sz = os.path.getsize(fp) / 1024.0
        mtime = datetime.datetime.fromtimestamp(os.path.getmtime(fp)).strftime('%Y-%m-%d %H:%M:%S')
        ext = os.path.splitext(f)[1].lower()
        
        # Check cache
        cache_name = os.path.splitext(f)[0] + ".docx"
        is_cached = os.path.exists(os.path.join(cache_dir, cache_name))
        
        records.append({
            "Filename": f,
            "Size (KB)": round(sz, 1),
            "Format": ext.replace('.', '').upper(),
            "Last Modified": mtime,
            "Cached .docx": "✅ Ready" if is_cached else "⏳ Pending"
        })
        
    return pd.DataFrame(records)

def upload_machine_files(machine_type, uploaded_files, lang="EN"):
    """
    Saves uploaded files (.doc, .docx) into GMP_Shift/{machine_type}/{lang}/ and syncs .docx_cache.
    uploaded_files: list of Streamlit UploadedFile objects or tuples (filename, bytes).
    """
    src_dir = os.path.join(GMP_SHIFT_BASE, machine_type, lang)
    cache_dir = os.path.join(DOCX_CACHE_BASE, machine_type, lang)
    os.makedirs(src_dir, exist_ok=True)
    os.makedirs(cache_dir, exist_ok=True)
    
    saved_files = []
    for uf in uploaded_files:
        fn = getattr(uf, 'name', None) or (uf[0] if isinstance(uf, (tuple, list)) else None)
        if not fn:
            continue
            
        content = uf.getvalue() if hasattr(uf, 'getvalue') else (uf.read() if hasattr(uf, 'read') else (uf[1] if isinstance(uf, (tuple, list)) else None))
        if content is None:
            continue
            
        if not fn.lower().endswith(('.doc', '.docx', '.txt')):
            continue
            
        target_src = os.path.join(src_dir, fn)
        with open(target_src, "wb") as f:
            f.write(content)
            
        # If it's a docx, save directly to cache too
        if fn.lower().endswith('.docx'):
            target_cache = os.path.join(cache_dir, fn)
            with open(target_cache, "wb") as f:
                f.write(content)
                
        saved_files.append(fn)
        
    # Trigger cache sync / conversion for any .doc files
    ensure_docx_cache(machine_type, lang)
    return saved_files

def delete_machine_files(machine_type, filenames, lang="EN"):
    """
    Deletes specified files from GMP_Shift/{machine_type}/{lang}/ and their cache counterparts.
    """
    src_dir = os.path.join(GMP_SHIFT_BASE, machine_type, lang)
    cache_dir = os.path.join(DOCX_CACHE_BASE, machine_type, lang)
    
    deleted = []
    for fn in filenames:
        src_fp = os.path.join(src_dir, fn)
        if os.path.exists(src_fp):
            try:
                os.remove(src_fp)
                deleted.append(fn)
            except Exception:
                pass
                
        # Also remove cache
        stem = os.path.splitext(fn)[0]
        cache_fp = os.path.join(cache_dir, stem + ".docx")
        if os.path.exists(cache_fp):
            try:
                os.remove(cache_fp)
            except Exception:
                pass
                
    return deleted

def get_machine_file_bytes(machine_type, filename, lang="EN"):
    """
    Reads and returns file bytes for download.
    """
    src_fp = os.path.join(GMP_SHIFT_BASE, machine_type, lang, filename)
    if os.path.exists(src_fp):
        with open(src_fp, "rb") as f:
            return f.read()
    cache_fp = os.path.join(DOCX_CACHE_BASE, machine_type, lang, filename)
    if os.path.exists(cache_fp):
        with open(cache_fp, "rb") as f:
            return f.read()
    return None

def get_available_shift_templates(machine_type="FP", lang="EN"):
    """
    Returns list of available shift register template files for a given machine type and language.
    """
    cache_dir = os.path.join(DOCX_CACHE_BASE, machine_type, lang)
    if os.path.exists(cache_dir):
        files = [f for f in os.listdir(cache_dir) if f.lower().endswith('.docx') and not f.lower().startswith('performer')]
        return sorted(files)
    src_dir = os.path.join(GMP_SHIFT_BASE, machine_type, lang)
    if os.path.exists(src_dir):
        files = [f for f in os.listdir(src_dir) if f.lower().endswith(('.doc', '.docx')) and not f.lower().startswith('performer')]
        return sorted(files)
    return []

def ensure_docx_cache(machine_type="FP", lang="EN"):
    """
    Ensures .docx versions of .doc files exist in DOCX_CACHE_BASE for fast merging.
    Also synchronizes any .docx files from GMP_Shift and ensures Performer.docx is present.
    """
    src_dir = os.path.join(GMP_SHIFT_BASE, machine_type, lang)
    dst_dir = os.path.join(DOCX_CACHE_BASE, machine_type, lang)
    if not os.path.exists(src_dir):
        return dst_dir

    os.makedirs(dst_dir, exist_ok=True)
    
    # 1. Sync any existing .docx files from source to cache if missing or newer
    for f in os.listdir(src_dir):
        if f.lower().endswith('.docx'):
            src_fp = os.path.join(src_dir, f)
            dst_fp = os.path.join(dst_dir, f)
            if not os.path.exists(dst_fp) or os.path.getmtime(src_fp) > os.path.getmtime(dst_fp):
                try:
                    shutil.copy2(src_fp, dst_fp)
                except Exception:
                    pass

    # 2. Check if any .doc needs conversion
    needs_conversion = []
    for f in os.listdir(src_dir):
        if f.lower().endswith('.doc'):
            src_fp = os.path.join(src_dir, f)
            dst_fp = os.path.join(dst_dir, os.path.splitext(f)[0] + '.docx')
            if not os.path.exists(dst_fp) or os.path.getmtime(src_fp) > os.path.getmtime(dst_fp):
                needs_conversion.append((src_fp, dst_fp))

    if needs_conversion:
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            for src_fp, dst_fp in needs_conversion:
                try:
                    doc = word.Documents.Open(src_fp, False, True)
                    doc.SaveAs2(dst_fp, 16) # 16 = wdFormatXMLDocument
                    doc.Close(False)
                except Exception:
                    pass
            word.Quit()
        except Exception:
            pass

    # 3. Ensure Performer.docx is synchronized
    get_performer_table_xml(machine_type, lang)

    return dst_dir

def parse_and_match_shifts(excel_source, machine_type="FP", lang="EN"):
    """
    Parses Shift Register Excel, extracts Column A (Variable) and Column E (Description/Text),
    and looks up matching shift templates in GMP_Shift/{machine_type}/{lang}/.
    """
    # Load Excel
    if isinstance(excel_source, (str, os.PathLike)):
        wb = openpyxl.load_workbook(excel_source, data_only=True)
    elif hasattr(excel_source, 'read'):
        file_bytes = excel_source.read()
        excel_source.seek(0)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    else:
        wb = openpyxl.load_workbook(excel_source, data_only=True)

    sheet_name = wb.sheetnames[0]
    ws = wb[sheet_name]

    # Locate headers
    col_var_idx = 1
    col_en_idx = 5
    header_row = 1

    for r in range(1, min(6, ws.max_row + 1)):
        for c in range(1, min(10, ws.max_column + 1)):
            v = str(ws.cell(r, c).value or '').strip().lower()
            if 'variable' in v or 'item' in v or 'shift' in v or 'code' in v:
                col_var_idx = c
                header_row = r
            elif 'description' in v or 'english' in v or 'text' in v or 'name' in v:
                col_en_idx = c

    # Ensure docx cache
    cache_dir = ensure_docx_cache(machine_type, lang)
    available_docx = {}
    if os.path.exists(cache_dir):
        for f in os.listdir(cache_dir):
            if f.lower().endswith('.docx'):
                stem = os.path.splitext(f)[0].lower()
                available_docx[stem] = os.path.join(cache_dir, f)

    matched_records = []
    all_records = []
    item_counter = 1

    for r in range(header_row + 1, ws.max_row + 1):
        var_val = str(ws.cell(r, col_var_idx).value or '').strip()
        text_en = str(ws.cell(r, col_en_idx).value or '').strip()

        if not var_val:
            continue

        # Extract code (e.g. MX_... or SR_... or filename stem)
        m = re.search(r'([A-Za-z0-9_]+)', var_val, re.IGNORECASE)
        shift_code = m.group(1) if m else var_val

        # Match against available shift docx
        matched_file = available_docx.get(shift_code.lower())

        if matched_file:
            rec = {
                "No.": item_counter,
                "Variable": var_val,
                "Shift Code": shift_code,
                "Description": text_en,
                "File Path": matched_file,
                "Status": "MATCHED"
            }
            matched_records.append(rec)
            all_records.append(rec)
            item_counter += 1
        else:
            rec = {
                "No.": "-",
                "Variable": var_val,
                "Shift Code": shift_code,
                "Description": text_en,
                "File Path": None,
                "Status": "NOT_FOUND"
            }
            all_records.append(rec)

    matched_df = pd.DataFrame(matched_records)
    full_df = pd.DataFrame(all_records)

    return {
        "matched_records": matched_records,
        "all_records": all_records,
        "matched_df": matched_df,
        "full_df": full_df,
        "total_matched": len(matched_records),
        "total_scanned": len(all_records)
    }

def generate_matched_shifts_excel(matched_records, base_name="5XXXX"):
    """
    Generates summary Excel containing Column A (No.), Column B (Variable / Shift Code), and Column C (Description).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matched Shift Register"

    # Header
    headers = ["No.", "Variable / Shift Code", "Description"]
    ws.append(headers)

    # Header styling
    header_fill = openpyxl.styles.PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = openpyxl.styles.Font(color="FFFFFF", bold=True, name="Arial", size=10)
    center_align = openpyxl.styles.Alignment(horizontal="center", vertical="center")
    left_align = openpyxl.styles.Alignment(horizontal="left", vertical="center")

    for c_idx in range(1, 4):
        cell = ws.cell(1, c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align if c_idx == 1 else left_align

    # Rows
    for idx, r in enumerate(matched_records, 1):
        ws.append([idx, r["Variable"], r["Description"]])
        ws.cell(idx + 1, 1).alignment = center_align
        ws.cell(idx + 1, 2).alignment = left_align
        ws.cell(idx + 1, 3).alignment = left_align

    # Column widths
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 42
    ws.column_dimensions['C'].width = 65

    doc_io = io.BytesIO()
    wb.save(doc_io)
    doc_io.seek(0)

    excel_filename = f"{base_name}_Matched_Shift_Register_Summary.xlsx"
    return doc_io.getvalue(), excel_filename

def generate_oq_shift_word(
    excel_source,
    machine_type="FP",
    word_template_path=None,
    lang="EN",
    **kwargs
):
    """
    Generates OQ Shift Register Word Protocol:
    1. Matches shift register items against GMP_Shift/{machine_type}/{lang}/.
    2. Opens Master Template and locates 'InsertTests' or 'XXXX' in Section 3.
    3. Merges shift register tables and Performer tables consecutively.
    4. Generates summary Excel of matched shift items.
    5. Returns doc_bytes, summary_excel_bytes, and result metadata.
    """
    parsed = parse_and_match_shifts(excel_source, machine_type, lang)
    matched_records = parsed["matched_records"]

    if not matched_records:
        raise ValueError(f"No matching shift register templates found in {machine_type} for the provided Excel variables.")

    template_to_use = word_template_path if (word_template_path and os.path.exists(word_template_path)) else DEFAULT_TEMPLATE_PATH
    if not os.path.exists(template_to_use):
        raise FileNotFoundError(f"Master template file not found: {template_to_use}")

    doc = docx.Document(template_to_use)

    # Locate placeholder paragraph in Section 3 Test Protocol
    target_p = None
    for p in doc.paragraphs:
        txt = p.text.strip().lower()
        if "inserttests" in txt or "xxxx" in txt:
            target_p = p
            break

    if target_p is None:
        # Fallback: search after Heading '3 Test Protocol'
        for idx, p in enumerate(doc.paragraphs):
            if "test protocol" in p.text.lower() and idx + 1 < len(doc.paragraphs):
                target_p = doc.paragraphs[idx + 1]
                break

    if target_p is None:
        raise ValueError("Could not locate 'InsertTests' or 'XXXX' placeholder under Section 3 in Word template.")

    # Load Performer sign-off table with multi-level fallback
    performer_tbl_xml = get_performer_table_xml(machine_type, lang)

    # Insert all matched shift register items
    for idx, item in enumerate(matched_records):
        shift_file = item["File Path"]
        if not shift_file or not os.path.exists(shift_file):
            continue

        shift_doc = docx.Document(shift_file)
        if not shift_doc.tables:
            continue

        # Clone and insert shift table
        shift_tbl_elem = copy.deepcopy(shift_doc.tables[0]._tbl)
        target_p._p.addprevious(shift_tbl_elem)

        # Insert spacing paragraph
        p_space = doc.add_paragraph()
        p_space.paragraph_format.space_before = Pt(4)
        p_space.paragraph_format.space_after = Pt(4)
        target_p._p.addprevious(p_space._p)

        # Insert Performer table
        if performer_tbl_xml is not None:
            perf_elem = copy.deepcopy(performer_tbl_xml)
            target_p._p.addprevious(perf_elem)

        # Insert page break between test items (except after the last one)
        if idx < len(matched_records) - 1:
            p_break = doc.add_paragraph()
            p_break.add_run().add_break(WD_BREAK.PAGE)
            target_p._p.addprevious(p_break._p)

    # Remove the placeholder paragraph
    target_p._p.getparent().remove(target_p._p)

    # Save Word document in memory
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    doc_bytes = doc_io.getvalue()

    file_name = os.path.basename(template_to_use)

    # Generate summary Excel
    base_name = os.path.splitext(file_name)[0].split('_')[0]
    if base_name == "XXXXX" or not base_name:
        base_name = "5XXXX"
    excel_bytes, excel_name = generate_matched_shifts_excel(matched_records, base_name)

    return {
        "success": True,
        "file_name": file_name,
        "doc_bytes": doc_bytes,
        "summary_excel_bytes": excel_bytes,
        "summary_excel_name": excel_name,
        "total_matched": parsed["total_matched"],
        "total_scanned": parsed["total_scanned"],
        "matched_df": parsed["matched_df"],
        "full_df": parsed["full_df"],
        "machine_type": machine_type
    }
