import os
import io
import re
import datetime
import zipfile
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sqlite3

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DB_PATH = os.path.join(BASE_DIR, "database", "ocr_system.db")
MACHINE_LIB_DIR = os.path.join(BASE_DIR, "IQOQDQ", "OQ_HMI", "Machine type")

def _find_default_word_template():
    hmi_dir = os.path.join(BASE_DIR, "IQOQDQ", "OQ_HMI")
    if os.path.exists(hmi_dir):
        files = [f for f in os.listdir(hmi_dir) if f.lower().endswith(('.docx', '.docm')) and not f.startswith('~$') and not f.startswith('Generated')]
        for f in files:
            if 'XXXXX' in f or 'NavPar' in f:
                return os.path.join(hmi_dir, f)
        if files:
            return os.path.join(hmi_dir, files[0])
    return os.path.join(hmi_dir, "XXXXX_04_OQ_HMI-V3_NavPar_2025-01-02_en.docx")

DEFAULT_MASTER_TEMPLATE = _find_default_word_template()
DEFAULT_SAMPLE_TEMPLATE = DEFAULT_MASTER_TEMPLATE

def _find_default_excel():
    hmi_dir = os.path.join(BASE_DIR, "IQOQDQ", "OQ_HMI")
    if os.path.exists(hmi_dir):
        files = [f for f in os.listdir(hmi_dir) if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$')]
        if files:
            for f in files:
                if 'text' in f.lower():
                    return os.path.join(hmi_dir, f)
            return os.path.join(hmi_dir, files[0])
    return os.path.join(hmi_dir, "56001-Texts.xlsx")

DEFAULT_EXCEL_PATH = _find_default_excel()
DEFAULT_SCREENSHOT_DIR = os.path.join(BASE_DIR, "IQOQDQ", "OQ_HMI", "Screen short")

def init_hmi_maintenance_db():
    """Initializes oq_hmi_template_maintenance SQLite table if not exists."""
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    os.makedirs(MACHINE_LIB_DIR, exist_ok=True)
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS oq_hmi_template_maintenance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            machine_type TEXT NOT NULL,
            visu_type TEXT NOT NULL DEFAULT 'IPC',
            function_code TEXT,
            screen_name TEXT NOT NULL,
            tab_index INTEGER DEFAULT 0,
            expected_header_title TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()

    # Seed default Machine Types if empty
    cursor.execute("SELECT COUNT(*) FROM oq_hmi_template_maintenance")
    count = cursor.fetchone()[0]
    if count == 0:
        seed_data = [
            ('Standard', 'IPC', '', 'Production', 0, r'\Operation\Production'),
            ('Standard', 'IPC', '', 'Setup', 0, r'\Operation\Setup'),
            ('Standard', 'IPC', '', 'Reference run', 0, r'\Operation\Reference run'),
            ('Standard', 'IPC', '', 'KPI-Values', 0, r'\Statistics\KPI-Values'),
            ('Standard', 'IPC', '', 'Batch report', 0, r'\Statistics\Batch report\Data'),
            ('Standard', 'IPC', 'V0160', 'Laminar flow', 1, r'\Operation\Laminar flow\Laminar flow'),
            ('Standard', 'IPC', 'V0200', 'Tube infeed', 1, r'\Operation\Tube infeed\Tube infeed'),
            ('Standard', 'IPC', 'V0252', 'Tube print registration', 1, r'\Operation\Tube print registration\Tube print registration'),
            ('Standard', 'IPC', 'V0270', 'Tube blow out', 1, r'\Operation\Tube blow out\Tube blow out'),
            ('Standard', 'IPC', 'V0300', 'Tube filling - Process values', 1, r'\Operation\Tube filling\Process values'),
            ('Standard', 'IPC', 'V0300', 'Tube filling - Dosing', 3, r'\Operation\Tube filling\Dosing'),
            ('Standard', 'IPC', 'V0300', 'Tube filling - Lifter', 4, r'\Operation\Tube filling\Lifter'),
            ('Standard', 'IPC', 'V0409', 'Product hopper', 1, r'\Operation\Product hopper\Monitoring, fill level'),
            ('Standard', 'IPC', 'V0980', 'Mechanical adjustment - Basic settings', 1, r'\Operation\Mechanical adjustment\Basic settings'),
            ('Standard', 'IPC', 'V0980', 'Mechanical adjustment - Positions', 2, r'\Operation\Mechanical adjustment\Positions'),
            ('FP', 'IPC', 'V0300', 'Tube filling', 1, r'\Operation\Tube filling\Process values'),
            ('FP', 'IPC', 'V0980', 'Mechanical adjustment', 1, r'\Operation\Mechanical adjustment\Basic settings'),
        ]
        cursor.executemany("""
            INSERT INTO oq_hmi_template_maintenance 
            (machine_type, visu_type, function_code, screen_name, tab_index, expected_header_title)
            VALUES (?, ?, ?, ?, ?, ?)
        """, seed_data)
        conn.commit()

    conn.close()

init_hmi_maintenance_db()

def get_available_hmi_machine_types() -> list[str]:
    """Retrieves available Machine Types from directory and SQLite database."""
    init_hmi_maintenance_db()
    machines = set()
    
    # Check directory
    if os.path.exists(MACHINE_LIB_DIR):
        for item in os.listdir(MACHINE_LIB_DIR):
            item_p = os.path.join(MACHINE_LIB_DIR, item)
            if os.path.isdir(item_p) and not item.startswith('.'):
                machines.add(item)

    # Check DB
    try:
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT DISTINCT machine_type FROM oq_hmi_template_maintenance WHERE machine_type IS NOT NULL AND machine_type != '' AND machine_type != 'Standard'")
        for row in cursor.fetchall():
            machines.add(row[0])
        conn.close()
    except Exception:
        pass

    if not machines:
        machines.add('FP')

    return sorted(list(machines))

def create_hmi_machine_type(machine_name: str) -> str:
    """Creates a new Machine Type library directory and default database entries."""
    clean_name = re.sub(r'[^a-zA-Z0-9_\-]', '', machine_name).strip().upper()
    if not clean_name:
        raise ValueError("Invalid machine type name.")

    mach_dir = os.path.join(MACHINE_LIB_DIR, clean_name)
    os.makedirs(mach_dir, exist_ok=True)

    # Add default seed entry if not exists in DB
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM oq_hmi_template_maintenance WHERE machine_type = ?", (clean_name,))
    if cursor.fetchone()[0] == 0:
        seed = [
            (clean_name, 'IPC', 'V0100', 'Product sensing', 1, r'\Operation\Product sensing\Product sensing'),
            (clean_name, 'IPC', 'V0980', 'Mechanical adjustment', 1, r'\Operation\Mechanical adjustment\Basic settings'),
        ]
        cursor.executemany("""
            INSERT INTO oq_hmi_template_maintenance 
            (machine_type, visu_type, function_code, screen_name, tab_index, expected_header_title)
            VALUES (?, ?, ?, ?, ?, ?)
        """, seed)
        conn.commit()
    conn.close()

    return clean_name

def delete_hmi_machine_type(machine_name: str):
    """Deletes a Machine Type directory and SQLite records."""
    clean_name = machine_name.strip().upper()
    mach_dir = os.path.join(MACHINE_LIB_DIR, clean_name)
    if os.path.exists(mach_dir):
        import shutil
        shutil.rmtree(mach_dir, ignore_errors=True)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM oq_hmi_template_maintenance WHERE machine_type = ?", (clean_name,))
    conn.commit()
    conn.close()

def list_hmi_machine_files_detailed(machine_type: str) -> pd.DataFrame:
    """
    Returns a pandas DataFrame of all template files in IQOQDQ/OQ_HMI/Machine type/{machine_type}/.
    Columns: Filename, Size (KB), Format, Last Modified
    """
    clean_mach = machine_type.strip() if machine_type else "Standard"
    src_dir = os.path.join(MACHINE_LIB_DIR, clean_mach)
    
    if not os.path.exists(src_dir):
        return pd.DataFrame(columns=["Filename", "Size (KB)", "Format", "Last Modified"])
        
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
        
        records.append({
            "Filename": f,
            "Size (KB)": round(sz, 1),
            "Format": ext.replace('.', '').upper(),
            "Last Modified": mtime
        })
        
    return pd.DataFrame(records)

def upload_hmi_machine_files(machine_type: str, uploaded_files) -> list[str]:
    """
    Saves uploaded files (.doc, .docx) into IQOQDQ/OQ_HMI/Machine type/{machine_type}/.
    """
    clean_mach = machine_type.strip() if machine_type else "Standard"
    src_dir = os.path.join(MACHINE_LIB_DIR, clean_mach)
    os.makedirs(src_dir, exist_ok=True)
    
    saved_files = []
    for uf in uploaded_files:
        fn = getattr(uf, 'name', None) or (uf[0] if isinstance(uf, (tuple, list)) else None)
        if not fn:
            continue
            
        content = uf.getvalue() if hasattr(uf, 'getvalue') else (uf.read() if hasattr(uf, 'read') else (uf[1] if isinstance(uf, (tuple, list)) else None))
        if content is None:
            continue
            
        if not fn.lower().endswith(('.doc', '.docx', '.docm', '.txt')):
            continue
            
        target_src = os.path.join(src_dir, fn)
        with open(target_src, "wb") as f:
            f.write(content)
                
        saved_files.append(fn)
        
    return saved_files

def delete_hmi_machine_files(machine_type: str, filenames: list[str]) -> list[str]:
    """
    Deletes specified files from IQOQDQ/OQ_HMI/Machine type/{machine_type}/.
    """
    clean_mach = machine_type.strip() if machine_type else "Standard"
    src_dir = os.path.join(MACHINE_LIB_DIR, clean_mach)
    
    deleted = []
    for fn in filenames:
        src_fp = os.path.join(src_dir, fn)
        if os.path.exists(src_fp):
            try:
                os.remove(src_fp)
                deleted.append(fn)
            except Exception:
                pass
                
    return deleted

def get_hmi_machine_file_bytes(machine_type: str, filename: str) -> bytes | None:
    """
    Reads and returns file bytes for download.
    """
    clean_mach = machine_type.strip() if machine_type else "Standard"
    src_fp = os.path.join(MACHINE_LIB_DIR, clean_mach, filename)
    if os.path.exists(src_fp):
        with open(src_fp, "rb") as f:
            return f.read()
    return None

def get_hmi_maintenance_records(machine_type: str = "Standard", visu_type: str = "IPC") -> pd.DataFrame:
    """Gets maintenance records for given machine type and visu type."""
    init_hmi_maintenance_db()
    
    clean_mach = machine_type.strip() if machine_type else "Standard"
    mach_dir = os.path.join(MACHINE_LIB_DIR, clean_mach)
    
    if clean_mach and clean_mach != "Standard" and os.path.isdir(mach_dir):
        docx_files = sorted([
            f for f in os.listdir(mach_dir)
            if f.lower().endswith(('.docx', '.docm')) and not f.startswith('~$')
        ])
        if docx_files:
            records = []
            for idx, fname in enumerate(docx_files, 1):
                fpath = os.path.join(mach_dir, fname)
                func_match = re.search(r'V\d{4}', fname)
                func_code = func_match.group(0) if func_match else ""
                
                tab_match = re.search(r'Tab_(\d+)', fname, re.IGNORECASE)
                tab_index = int(tab_match.group(1)) if tab_match else 0
                
                screen_name = os.path.splitext(fname)[0]
                
                expected_title = screen_name
                try:
                    doc = docx.Document(fpath)
                    if doc.tables and len(doc.tables[0].rows) > 0 and len(doc.tables[0].rows[0].cells) > 0:
                        cell_txt = doc.tables[0].rows[0].cells[0].text.strip()
                        if cell_txt:
                            expected_title = re.sub(r'[\r\n]+', ' ', cell_txt).strip()
                except Exception:
                    pass

                mtime = os.path.getmtime(fpath)
                created_str = datetime.datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")

                records.append({
                    "id": idx,
                    "machine_type": clean_mach,
                    "visu_type": visu_type,
                    "function_code": func_code,
                    "screen_name": screen_name,
                    "tab_index": tab_index,
                    "expected_header_title": expected_title,
                    "created_at": created_str
                })
            
            conn = sqlite3.connect(DB_PATH)
            sql_df = pd.read_sql_query("""
                SELECT id, machine_type, visu_type, function_code, screen_name, tab_index, expected_header_title, created_at
                FROM oq_hmi_template_maintenance
                WHERE machine_type = ? AND visu_type = ?
            """, conn, params=(clean_mach, visu_type))
            conn.close()

            dir_df = pd.DataFrame(records)
            if not sql_df.empty:
                combined = pd.concat([dir_df, sql_df], ignore_index=True)
                combined.drop_duplicates(subset=["screen_name"], keep="first", inplace=True)
                combined["id"] = range(1, len(combined) + 1)
                return combined
            return dir_df

    # Fallback SQLite query for Standard or machine_type without directory sub-files
    conn = sqlite3.connect(DB_PATH)
    query = """
        SELECT id, machine_type, visu_type, function_code, screen_name, tab_index, expected_header_title, created_at
        FROM oq_hmi_template_maintenance
        WHERE (machine_type = ? OR machine_type = 'Standard')
          AND visu_type = ?
        ORDER BY machine_type DESC, function_code ASC, id ASC
    """
    df = pd.read_sql_query(query, conn, params=(clean_mach, visu_type))
    conn.close()
    return df

def add_hmi_maintenance_record(machine_type: str, visu_type: str, function_code: str, screen_name: str, tab_index: int, expected_header_title: str):
    """Adds a new maintenance record to SQLite."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO oq_hmi_template_maintenance 
        (machine_type, visu_type, function_code, screen_name, tab_index, expected_header_title)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (machine_type.upper(), visu_type, function_code, screen_name, int(tab_index), expected_header_title))
    conn.commit()
    conn.close()

def delete_hmi_maintenance_record(record_id: int):
    """Deletes a maintenance record by ID."""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM oq_hmi_template_maintenance WHERE id = ?", (int(record_id),))
    conn.commit()
    conn.close()

def find_matching_sub_docx(fname: str, machine_type: str) -> str | None:
    """
    Finds matching .docx sub-file in Machine Type directory for given screenshot image filename.
    """
    clean_mach = machine_type.strip() if machine_type else ""
    mach_dir = os.path.join(MACHINE_LIB_DIR, clean_mach)
    if not clean_mach or not os.path.isdir(mach_dir):
        return None

    fname_base = os.path.splitext(os.path.basename(fname))[0]
    norm_img_base = re.sub(r'[^a-zA-Z0-9]', '', fname_base).lower()
    norm_img_no_tab0 = re.sub(r'tab0+$', '', norm_img_base)

    docx_files = [f for f in os.listdir(mach_dir) if f.lower().endswith(('.docx', '.docm')) and not f.startswith('~$')]
    
    # 1. Check exact or normalized filename matches
    for f in docx_files:
        f_base = os.path.splitext(f)[0]
        norm_f = re.sub(r'[^a-zA-Z0-9]', '', f_base).lower()
        norm_f_no_tab0 = re.sub(r'tab0+$', '', norm_f)

        if norm_img_base == norm_f or fname_base.lower() == f_base.lower() or norm_img_no_tab0 == norm_f or norm_img_base == norm_f_no_tab0:
            return os.path.join(mach_dir, f)

    # 2. Check function code & tab index match
    func_m = re.search(r'V\d{4}', fname_base)
    func_code = func_m.group(0).lower() if func_m else ''
    tab_m = re.search(r'Tab_(\d+)', fname_base, re.IGNORECASE)
    tab_num = tab_m.group(1) if tab_m else '0'

    if func_code:
        for f in docx_files:
            f_base = os.path.splitext(f)[0]
            norm_f = re.sub(r'[^a-zA-Z0-9]', '', f_base).lower()
            if func_code in norm_f:
                if f'tab{tab_num}' in norm_f or tab_num in ('0', '00'):
                    return os.path.join(mach_dir, f)

    return None

def perform_existing_template_recheck(image_ocr_map: list[dict], machine_type: str = "Standard", visu_type: str = "IPC") -> tuple[pd.DataFrame, dict]:
    """
    Compares uploaded screenshot images against sub-docx templates in Machine Type directory.
    Returns comparison dataframe and metrics dict.
    """
    recheck_records = []
    maintained_count = 0
    missing_maint_count = 0
    duplicate_count = 0

    seen_sub_docx = {}

    for idx, item in enumerate(image_ocr_map, 1):
        fname = item['fname']
        t_title = item['title']
        func_match = re.search(r'V\d{4}', fname)
        func_code = func_match.group(0) if func_match else ''

        # Match using sub-docx file in Machine Type directory
        matched_docx = find_matching_sub_docx(fname, machine_type)

        if matched_docx and os.path.exists(matched_docx):
            if matched_docx in seen_sub_docx:
                status = "⚡ Skipped (Duplicate)"
                action = f"Duplicate screen view (covered by '{seen_sub_docx[matched_docx]}')"
                duplicate_count += 1
            else:
                seen_sub_docx[matched_docx] = fname
                status = "✅ Maintained & Matched"
                action = "None (Ready for Word Generation)"
                maintained_count += 1
                try:
                    doc_sub = docx.Document(matched_docx)
                    if doc_sub.tables and len(doc_sub.tables[0].rows) > 0 and len(doc_sub.tables[0].rows[0].cells) > 0:
                        cell_txt = doc_sub.tables[0].rows[0].cells[0].text.strip()
                        if cell_txt:
                            t_title = re.sub(r'[\r\n]+', ' ', cell_txt).strip()
                except Exception:
                    pass
        else:
            status = "⚠️ Unmaintained / Missing in Library"
            action = f"Add '{os.path.splitext(fname)[0]}.docx' to Machine Type library"
            missing_maint_count += 1

        recheck_records.append({
            "No.": idx,
            "Image File Name": fname,
            "Function Code": func_code if func_code else "N/A",
            "Header / Screen Title": t_title,
            "Template Maintenance Status": status,
            "Recommended Action": action
        })

    recheck_df = pd.DataFrame(recheck_records)
    metrics = {
        "total_evaluated": len(image_ocr_map),
        "maintained_count": maintained_count,
        "duplicate_count": duplicate_count,
        "missing_maint_count": missing_maint_count
    }
    return recheck_df, metrics

def parse_hmi_texts_excel(excel_source) -> tuple[pd.DataFrame, dict]:
    """
    Reads HMI Texts Excel (e.g. 56041-Texts.xlsx), extracts GroupPath, Name, Description,
    and multilingual text columns (English, German, Chinese).
    Returns (cleaned_dataframe, summary_dict).
    """
    if isinstance(excel_source, (str, os.PathLike)):
        df_raw = pd.read_excel(excel_source)
    elif hasattr(excel_source, 'read'):
        file_bytes = excel_source.read()
        excel_source.seek(0)
        df_raw = pd.read_excel(io.BytesIO(file_bytes))
    else:
        raise ValueError("Unsupported Excel file source.")

    if df_raw.empty:
        raise ValueError("Provided Excel file is empty.")

    # Check language row header (Row 0)
    lang_map = {}
    if pd.isna(df_raw.iloc[0]['GroupPath']):
        row0 = df_raw.iloc[0]
        for col in ['L33796', 'L32777', 'L32775']:
            if col in row0 and pd.notna(row0[col]):
                lang_map[col] = str(row0[col]).strip()
        df_data = df_raw.iloc[1:].copy().reset_index(drop=True)
    else:
        df_data = df_raw.copy()

    english_col = lang_map.get('L32777', 'IWKEnglish')
    german_col = lang_map.get('L32775', 'IWKGerman')
    chinese_col = lang_map.get('L33796', 'IWKChinesisch')

    df_data.rename(columns={
        'L32777': english_col,
        'L32775': german_col,
        'L33796': chinese_col
    }, inplace=True)

    # Clean text values
    for col in df_data.columns:
        df_data[col] = df_data[col].apply(lambda x: str(x).strip() if pd.notna(x) and str(x).strip() != 'nan' else '')

    categories = {}
    for gp in df_data['GroupPath'].unique():
        if gp:
            cat = gp.split('.')[0]
            categories[cat] = categories.get(cat, 0) + 1

    summary = {
        "total_rows": len(df_data),
        "english_col": english_col,
        "german_col": german_col,
        "chinese_col": chinese_col,
        "categories": categories
    }

    return df_data, summary


def get_screenshot_images(image_source=None) -> list[dict]:
    """
    Retrieves available screenshot image files from directory or uploaded ZIP.
    Returns list of dicts: [{'name': filename, 'path': path_or_bytes, 'is_bytes': bool}]
    """
    image_list = []

    if image_source is None or (isinstance(image_source, str) and not os.path.exists(image_source)):
        image_source = DEFAULT_SCREENSHOT_DIR

    if isinstance(image_source, (str, os.PathLike)) and os.path.isdir(image_source):
        for root, dirs, files in os.walk(image_source):
            for fname in sorted(files):
                if fname.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                    fpath = os.path.join(root, fname)
                    image_list.append({
                        'name': fname,
                        'path': fpath,
                        'is_bytes': False
                    })
    elif hasattr(image_source, 'read'):
        # ZIP file uploaded
        file_bytes = image_source.read()
        image_source.seek(0)
        try:
            with zipfile.ZipFile(io.BytesIO(file_bytes), 'r') as z:
                for fname in sorted(z.namelist()):
                    if fname.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')) and not fname.startswith('__MACOSX'):
                        img_bytes = z.read(fname)
                        base_name = os.path.basename(fname)
                        image_list.append({
                            'name': base_name,
                            'path': img_bytes,
                            'is_bytes': True
                        })
        except zipfile.BadZipFile:
            pass
    elif isinstance(image_source, list):
        # List of Streamlit UploadedFile objects or image dicts
        for uploaded_file in image_source:
            if isinstance(uploaded_file, dict) and 'name' in uploaded_file:
                image_list.append(uploaded_file)
            elif hasattr(uploaded_file, 'name') and uploaded_file.name.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp')):
                img_bytes = uploaded_file.read()
                uploaded_file.seek(0)
                image_list.append({
                    'name': uploaded_file.name,
                    'path': img_bytes,
                    'is_bytes': True
                })

    return image_list


def perform_ocr_on_image(img_info: dict) -> dict:
    """
    Performs OCR on a screenshot image using winocr (Windows Native OCR) or pytesseract fallback.
    Extracts text, lines, and screen tab title.
    """
    extracted_text = ""
    lines = []
    try:
        from PIL import Image
        import winocr

        if img_info.get('is_bytes'):
            img_obj = Image.open(io.BytesIO(img_info['path']))
        else:
            img_obj = Image.open(img_info['path'])

        res = winocr.recognize_pil_sync(img_obj, lang='en')
        extracted_text = res.get('text', '')
        lines = [line['text'].strip() for line in res.get('lines', []) if line['text'].strip()]
    except Exception:
        try:
            import pytesseract
            from PIL import Image
            if img_info.get('is_bytes'):
                img_obj = Image.open(io.BytesIO(img_info['path']))
            else:
                img_obj = Image.open(img_info['path'])
            extracted_text = pytesseract.image_to_string(img_obj)
            lines = [l.strip() for l in extracted_text.split('\n') if l.strip()]
        except Exception:
            extracted_text = ""
            lines = []

    fname = img_info['name']
    base_name = os.path.splitext(fname)[0]

    # Clean fallback title from filename
    words = re.sub(r'([a-z0-9])([A-Z])', r'\1 \2', base_name)
    words = re.sub(r'[^a-zA-Z0-9]', ' ', words)
    skip_words = {'view', 'sub', 'navigation', 'k', 'v0000', 'v0023', 'v0042', 'v0043', 'v0051', 'v0100', 'v0101', 'v0103', 'v0106', 'v0151', 'v0170', 'v0172', 'v0181', 'v0201', 'v0303', 'v0536', 'v0700', 'v0760', 'v0900', 'v1536', 'v2536', 'v3536', '00', '01', '02', '03', '04', '05', 'tab', 'production', 'kpivalues', 'recipe', 'history', 'diagnosis', 'service', 'more', 'settings'}
    clean_words = [w for w in words.split() if w.lower() not in skip_words]
    fn_title = ' '.join(clean_words).strip() if clean_words else base_name

    # Filter standard HMI top bar status words to locate screen title
    ignore_set = {'fault', 'latest message', 'power supply', 'speed [cyc/min]', 'format', 'user', 'administrator', 'pieces', '0 / 0', '100', 'f 1', 'no.', 'time', 'date', 'duration', 'y', 'o.', 'data', 'group', 'usergroup', 'administrator-group', 'users'}
    candidates = []
    for line in lines[:15]:
        l_lower = line.lower()
        if l_lower not in ignore_set and not l_lower.startswith('latest') and not 'power supply' in l_lower and not l_lower.isdigit() and len(line) > 2:
            candidates.append(line)

    best_title = ''
    for cand in candidates:
        cand_clean = re.sub(r'[^a-zA-Z0-9]', '', cand).lower()
        fn_clean = re.sub(r'[^a-zA-Z0-9]', '', fn_title).lower()
        if cand_clean and fn_clean and (cand_clean in fn_clean or fn_clean in cand_clean or any(w.lower() in cand_clean for w in fn_title.split() if len(w) > 3)):
            best_title = cand
            break

    if not best_title and candidates:
        best_title = candidates[0]

    # Special overrides for specific sub-nav tabs & sub-pages
    if 'runtaskbutton' in base_name.lower() and 'refrun' not in base_name.lower():
        derived_title = "Production"
    elif 'setuptaskbutton' in base_name.lower() and 'refrun' not in base_name.lower():
        derived_title = "Setup"
    elif base_name == 'View_KPIValuesKSubNavigationView_BatchReportView':
        derived_title = "Batch report/Data"
    elif base_name == 'View_KPIValuesKSubNavigationView_EfficiencyView':
        derived_title = "Batch report/Times"
    elif 'StatisticCounterView_MonitoringCounterView' in base_name:
        derived_title = "Current counters/Monitoring counter"
    elif 'StatisticCounterView_ProductionCounterView' in base_name:
        derived_title = "Current counters/Production counter"
    elif 'StatisticCounterView_RejectsHistory' in base_name:
        derived_title = "Current counters/Rejects history"
    elif base_name in ('View_RecipeSubNavigationView_FormatView', 'View_RecipeSubNavigationViewRecipeView'):
        derived_title = "Format"
    elif base_name == 'View_HistorySubNavigationView_HistoryMessagesView':
        derived_title = "Logbook"
    elif base_name == 'View_HistorySubNavigationView_StatisticalMessagesView':
        derived_title = "Message history"
    elif base_name == 'View_HistorySubNavigationView_DiagnosisLoggingView':
        derived_title = "Message statistics"
    elif base_name == 'View_ServiceSubNavigationView_ServiceBackupView':
        derived_title = "Data backup"
    elif base_name in ('View_MoreView_SetupStripTransferOverviewView', 'View_MoreView_SetupStripTransferOverviewView_Tab_0'):
        derived_title = "Strip transfer/Process values"
    else:
        derived_title = best_title if best_title else fn_title

    return {
        'extracted_text': extracted_text,
        'derived_title': derived_title,
        'clean_ocr': re.sub(r'[^a-zA-Z0-9]', '', extracted_text).lower()
    }


def generate_image_audit_excel(all_images: list[dict], inserted_map: dict) -> bytes:
    """
    Generates an Excel report listing all screenshot image names,
    indicating whether each image was inserted into the Word template or not.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "HMI Image Insertion Report"
    ws.views.sheetView[0].showGridLines = True

    # Styling
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    inserted_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    inserted_font = Font(name="Segoe UI", size=10, color="276A3C", bold=True)
    not_inserted_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    not_inserted_font = Font(name="Segoe UI", size=10, color="C65911", bold=True)
    recheck_verified_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    recheck_verified_font = Font(name="Segoe UI", size=10, color="276A3C", bold=True)
    recheck_aligned_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    recheck_aligned_font = Font(name="Segoe UI", size=10, color="1F4E78", bold=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Title Block
    ws.merge_cells("A1:F1")
    ws["A1"] = "OQ HMI Screenshot Image Insertion & OCR Verification Audit Report"
    ws["A1"].font = Font(name="Segoe UI", size=14, bold=True, color="1F4E78")
    ws["A1"].alignment = Alignment(vertical="center")

    ws["A2"] = f"Generated Date: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws["A2"].font = Font(name="Segoe UI", size=10, italic=True, color="595959")

    # Table Headers
    headers = ["No.", "Image File Name", "Insertion Status", "Matched Screen / Table", "OCR Recheck Status", "Remarks"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 5
    inserted_count = 0
    not_inserted_count = 0

    for idx, img_info in enumerate(all_images, 1):
        fname = img_info['name']
        match_info = inserted_map.get(fname)

        is_inserted = match_info is not None and not match_info.get('is_duplicate', False)
        is_duplicate = match_info is not None and match_info.get('is_duplicate', False)

        if is_inserted:
            inserted_count += 1
            status_str = "Inserted"
            target_screen = match_info.get('table_title', 'Matched Table')
            recheck_status = match_info.get('recheck_status', '✅ Verified (OCR Match)')
            remarks = match_info.get('remarks', f"Inserted into Table {match_info.get('table_idx', '')}")
        elif is_duplicate:
            not_inserted_count += 1
            status_str = "Skipped (Duplicate)"
            target_screen = match_info.get('table_title', 'N / A')
            recheck_status = "Skipped (Duplicate)"
            remarks = match_info.get('remarks', f"Duplicate screen view skipped (already covered by '{match_info.get('primary_fname', '')}')")
        else:
            not_inserted_count += 1
            status_str = "Not Inserted"
            target_screen = "N / A"
            recheck_status = "N / A"
            remarks = "No matching screen header found in Word Master Template"

        ws.cell(row=row_idx, column=1, value=idx).alignment = Alignment(horizontal="center")
        ws.cell(row=row_idx, column=2, value=fname).alignment = Alignment(horizontal="left")
        
        status_cell = ws.cell(row=row_idx, column=3, value=status_str)
        status_cell.alignment = Alignment(horizontal="center")
        if is_inserted:
            status_cell.fill = inserted_fill
            status_cell.font = inserted_font
        else:
            status_cell.fill = not_inserted_fill
            status_cell.font = not_inserted_font

        ws.cell(row=row_idx, column=4, value=target_screen).alignment = Alignment(horizontal="left")

        recheck_cell = ws.cell(row=row_idx, column=5, value=recheck_status)
        recheck_cell.alignment = Alignment(horizontal="center")
        if "Auto-Aligned" in recheck_status:
            recheck_cell.fill = recheck_aligned_fill
            recheck_cell.font = recheck_aligned_font
        elif "Verified" in recheck_status:
            recheck_cell.fill = recheck_verified_fill
            recheck_cell.font = recheck_verified_font
        else:
            recheck_cell.fill = not_inserted_fill
            recheck_cell.font = not_inserted_font

        ws.cell(row=row_idx, column=6, value=remarks).alignment = Alignment(horizontal="left")

        for c in range(1, 7):
            ws.cell(row=row_idx, column=c).border = thin_border

        row_idx += 1

    # Summary Block at Bottom
    row_idx += 1
    ws.cell(row=row_idx, column=2, value="Total Images Evaluated:").font = Font(bold=True)
    ws.cell(row=row_idx, column=3, value=len(all_images)).font = Font(bold=True)
    row_idx += 1
    ws.cell(row=row_idx, column=2, value="Total Images Inserted:").font = Font(color="276A3C", bold=True)
    ws.cell(row=row_idx, column=3, value=inserted_count).font = Font(color="276A3C", bold=True)
    row_idx += 1
    ws.cell(row=row_idx, column=2, value="Total Images Not Inserted:").font = Font(color="C65911", bold=True)
    ws.cell(row=row_idx, column=3, value=not_inserted_count).font = Font(color="C65911", bold=True)

    # Adjust Column Widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 45
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 45
    ws.column_dimensions['E'].width = 25
    ws.column_dimensions['F'].width = 50

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()



def get_reference_standard_header(fname, ocr_res):
    base = os.path.splitext(fname)[0]

    # 1. OPERATION (Main)
    if base.startswith('View_ProductionSubNavigation') or base.startswith('View_RunTaskButton') or 'SetupRefRunView' in base or ('RunTaskButton' in base and 'View_Settings' not in base):
        if 'runtaskbutton' in base.lower() and 'refrun' not in base.lower():
            return r'\Operation\Production'
        elif 'setuptaskbutton' in base.lower() and 'refrun' not in base.lower():
            return r'\Operation\Setup'
        elif 'refrun' in base.lower():
            return r'\Operation\Reference run'
        return r'\Operation\Setup'

    # 2. STATISTICS
    elif base.startswith('View_KPIValuesKSubNavigation'):
        if 'StatisticCounterView_ProductionCounter' in base or ('ProductionCounter' in base and 'Performance' not in base):
            return r'\Statistics\Current counters\Production'
        elif 'MonitoringCounterView' in base:
            return r'\Statistics\Current counters\Monitoring'
        elif 'RejectsHistory' in base or 'RejectReasons' in base:
            return r'\Statistics\Current counters\Reject reasons'
        elif 'BatchReportView' in base:
            return r'\Statistics\Batch report\Data'
        elif 'EfficiencyView' in base:
            return r'\Statistics\Batch report\Times'
        elif base.endswith('KPIValuesView') or base.endswith('KPIValues'):
            return r'\Statistics\KPI-Values'
        elif 'Performance' in base and ('Counter' in base or 'ProductionCounter' in base):
            return r'\Statistics\Performance\Counter'
        elif 'Performance' in base and ('State' in base or 'MachineState' in base):
            return r'\Statistics\Performance\State'
        elif 'RejectsView' in base or 'FaultsRejects' in base:
            return r'\Statistics\Faults / rejects\Rejects'
        elif 'FaultsView' in base:
            return r'\Statistics\Faults / rejects\Faults'
        
        # OCR Fallback if filename is non-standard
        txt = ocr_res.get('extracted_text', '') if ocr_res else ''
        txt_lower = txt.lower()

        if 'performance' in txt_lower and 'state' in txt_lower:
            return r'\Statistics\Performance\State'
        elif 'performance' in txt_lower and 'counter' in txt_lower and 'current' not in txt_lower:
            return r'\Statistics\Performance\Counter'
        elif 'fault' in txt_lower or 'reject' in txt_lower:
            if 'fault' in txt_lower and 'reject' not in txt_lower:
                return r'\Statistics\Faults / rejects\Faults'
            return r'\Statistics\Faults / rejects\Rejects'
        elif 'kpi' in txt_lower or 'oee' in txt_lower:
            return r'\Statistics\KPI-Values'
        return r'\Statistics\Batch report\Data'

    # 3. FORMATS
    elif base.startswith('View_RecipeSubNavigation'):
        if 'FormatCompare' in base:
            return r'\Formats\Format compare'
        elif 'FormatBackup' in base:
            return r'\Formats\Format backup'
        return r'\Formats\Format'

    # 4. HISTORY
    elif base.startswith('View_HistorySubNavigation'):
        if 'HistoryMessagesView' in base or 'HistoryMessage' in base:
            return r'\History\Message history'
        elif 'StatisticalMessagesView' in base or 'StatisticalMessage' in base:
            return r'\History\Message statistics'
        elif 'DiagnosisLoggingView' in base or 'Logbook' in base:
            return r'\History\Logbook'
        return r'\History\Message history'

    # 5. DIAGNOSIS
    elif base.startswith('View_DiagnosisSubNavigation'):
        if 'ShiftRegister' in base or 'Shiftregister' in base:
            return r'\Diagnosis\Shiftregister-View'
        return r'\Diagnosis\PLC diagnosis'

    # 6. SERVICE
    elif base.startswith('View_ServiceSubNavigation'):
        if 'UserAdministrationView_User' in base:
            return r'\Service\User administration\User'
        elif 'UserAdministrationView_Group' in base:
            return r'\Service\User administration\Group'
        elif 'UserAdministrationView_Common' in base:
            return r'\Service\User administration\General'
        elif 'ServiceBackup' in base:
            return r'\Service\Data backup'
        elif 'ServiceFileExport' in base:
            return r'\Service\Export files'
        elif 'ServiceArchiving' in base:
            return r'\Service\Scheduled archiving'
        elif 'SystemSettings_Misc' in base:
            return r'\Service\System settings\General'
        elif 'SystemSettings_Net' in base:
            return r'\Service\System settings\Network'
        elif 'SystemSettings_Time' in base:
            return r'\Service\System settings\Time'
        elif 'TextAdministration' in base:
            return r'\Service\Texts'
        elif 'UserAdministration' in base:
            return r'\Service\User administration\User'
        return r'\Service\General'

    # 7. (...) MORE
    elif base.startswith('View_More'):
        if base == 'View_MoreView' or base == 'View_More':
            return r'\(...)'
        elif 'SingleJog' in base or 'SJog' in base:
            if 'TZ_' in base:
                return r'\(...)\TZ - jogging single drive'
            return r'\(...)\Jogging single drive'
        elif 'RefRun' in base:
            if 'TZ_' in base:
                return r'\(...)\TZ - reference run'
            return r'\(...)\Reference run'
        elif 'SetupAbsEnc' in base:
            return r'\(...)\Setup single drive'
        elif 'SetupStripTransferOverview' in base:
            if 'Tab_1_StripInfeed' in base:
                unit = base.split('StripInfeed')[-1]
                return f'\\(...)\\Strip transfer\\Strip infeed-Unit {unit}'
            elif 'Tab_2_StripTransfer' in base:
                unit = base.split('StripTransfer')[-1]
                return f'\\(...)\\Strip transfer\\Strip transfer-Unit {unit}'
            elif 'Tab_3_StripTransferSlide' in base:
                unit = base.split('StripTransferSlide')[-1]
                return f'\\(...)\\Strip transfer\\Strip transfer, slide-Unit {unit}'
            elif 'Tab_1' in base:
                return r'\(...)\Strip transfer\Strip infeed'
            elif 'Tab_2' in base:
                return r'\(...)\Strip transfer\Strip transfer'
            elif 'Tab_3' in base:
                return r'\(...)\Strip transfer\Strip transfer, slide'
            else:
                return r'\(...)\Strip transfer\Process values'
        elif 'SetupHeightAdjust' in base:
            return r'\(...)\Height adjustment'
        elif 'SetupTubePrintReg' in base:
            return r'\(...)\Setup tube print registration'
        elif 'SetupTubeFill' in base:
            return r'\(...)\Setup tube filling'
        elif 'SetupPackTags' in base:
            if 'Tab_1' in base:
                return r'\(...)\PackTags\Messages'
            elif 'Tab_2' in base:
                return r'\(...)\PackTags\Unit mode / State'
            elif 'Tab_3' in base:
                return r'\(...)\PackTags\Parameter'
            elif 'Tab_4' in base:
                return r'\(...)\PackTags\Other'
            elif 'Tab_5' in base:
                return r'\(...)\PackTags\Production counter'
            elif 'Tab_6' in base:
                return r'\(...)\PackTags\Product data'
            return r'\(...)\PackTags\Machine speed'
        elif 'SetupVersionInfo' in base:
            return r'\(...)\Version display'
        return r'\(...)'

    # 8. OPERATION (SETTINGS)
    elif base.startswith('View_Settings') or (base.startswith('View_SetupTaskButton') and 'RefRun' not in base):
        base_clean = base.split('.')[0]
        # Normalize Tab_0 / Tab_00 to base view name so Tab_0 yields identical header title as non-Tab_0
        base_norm = re.sub(r'_Tab_00?(_.*)?$', '', base_clean)
        has_sub_parts = bool(re.search(r'Tab_\d+_.+', base_clean))

        if not has_sub_parts and ('V0160' in base_norm or 'V1015' in base_norm or 'LaminarFlow' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Laminar flow device\Laminar flow device' if ('V1015' in base_norm or 'TZ' in base_norm) else r'\Operation\Laminar flow\Laminar flow'
            return r'\Operation\Laminar flow device\Settings' if ('V1015' in base_norm or 'TZ' in base_norm) else r'\Operation\Laminar flow\Settings'
        elif not has_sub_parts and ('V0200_00_TubeInfeedView' in base_norm or 'TubeInfeed' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Tube infeed\Tube infeed'
            elif 'Tab_2' in base:
                return r'\Operation\Tube infeed\Monitoring, tube position'
            return r'\Operation\Tube infeed\Settings'
        elif not has_sub_parts and ('V0252_00_TubePrintRegView' in base_norm or 'TubePrintReg' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Tube print registration\Tube print registration'
            return r'\Operation\Tube print registration\Settings'
        elif not has_sub_parts and ('V0270_00_TubeBlowOutView' in base_norm or 'TubeBlowOut' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Tube blow out\Tube blow out'
            return r'\Operation\Tube blow out\Settings'
        elif not has_sub_parts and ('V0300_00_TubeFillAView' in base_norm or 'TubeFill' in base_norm):
            if 'Tab_1' in base or 'Process' in base or 'Special' in base:
                return r'\Operation\Tube filling\Process values'
            elif 'Tab_2' in base or 'Weight' in base:
                return r'\Operation\Tube filling\Weight settings'
            elif 'Tab_3' in base or 'Dosing' in base:
                return r'\Operation\Tube filling\Dosing'
            elif 'Tab_4' in base or 'Lifter' in base:
                return r'\Operation\Tube filling\Lifter'
            elif 'Tab_5' in base or 'Shut' in base:
                return r'\Operation\Tube filling\Shut-off'
            elif 'Tab_6' in base or 'Blow' in base:
                return r'\Operation\Tube filling\Blow-off filling'
            return r'\Operation\Tube filling\Settings'
        elif not has_sub_parts and ('V0409' in base_norm or 'ProdHopper' in base_norm):
            if 'Tab_0' in base_clean or 'Tab_00' in base_clean or not re.search(r'Tab_\d+', base_clean):
                return r'\Operation\Product hopper\Settings'
            
            txt_l = (ocr_res.get('extracted_text', '') if ocr_res else '').lower()
            if 'target fill level' in txt_l or 'fill level' in txt_l or 'pump off' in txt_l or 'pump on' in txt_l or 'max. supply' in txt_l or 'min. supply' in txt_l:
                return r'\Operation\Product hopper\Monitoring, fill level'
            elif 'gas purging' in txt_l or 'purging interval' in txt_l:
                return r'\Operation\Product hopper\Gas purging'
            elif 'current temperature' in txt_l or 'regulator output' in txt_l:
                return r'\Operation\Product hopper\Process values'
            elif 'target value, speed' in txt_l or '[1 /min]' in txt_l:
                return r'\Operation\Product hopper\Agitator'
            elif 'calibration offset' in txt_l or 'target value, temperature' in txt_l or 'temperature sensor' in txt_l:
                return r'\Operation\Product hopper\Heater'

            if 'Tab_1' in base:
                return r'\Operation\Product hopper\Monitoring, fill level'
            elif 'Tab_2' in base:
                return r'\Operation\Product hopper\Gas purging'
            elif 'Tab_3' in base:
                return r'\Operation\Product hopper\Agitator'
            elif 'Tab_4' in base:
                return r'\Operation\Product hopper\Heater'
            return r'\Operation\Product hopper\Settings'
        elif not has_sub_parts and ('V0530' in base_norm or 'FoldView' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Folding station\Folding station'
            return r'\Operation\Folding station\Folding station'
        elif not has_sub_parts and ('V0610' in base_norm or 'GoodTubeDisch' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Good tube discharge\Good tube discharge'
            return r'\Operation\Good tube discharge\Settings'
        elif not has_sub_parts and ('V0620' in base_norm or 'BadTubeDisch' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Faulty tube ejection\Faulty tube ejection'
            return r'\Operation\Faulty tube ejection\Settings'
        elif not has_sub_parts and ('V0690' in base_norm or 'TubeDischConv' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Tube discharge conveyor\Tube discharge conveyor'
            return r'\Operation\Tube discharge conveyor\Settings'
        elif not has_sub_parts and 'V0900' in base_norm:
            if 'Tab_1' in base:
                return r'\Operation\Basic settings\Tube dimensions'
            elif 'Tab_2' in base:
                return r'\Operation\Basic settings\Speed'
            elif 'Tab_3' in base:
                return r'\Operation\Basic settings\Reference offset values'
            return r'\Operation\Basic settings\Settings'
        elif not has_sub_parts and ('V1100' in base_norm or 'BoxChain' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Box chain\Box chain'
            return r'\Operation\Box chain\Settings'
        elif not has_sub_parts and ('V1140' in base_norm or 'BoxInfeed' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Box infeed\Box infeed'
            return r'\Operation\Box infeed\Settings'
        elif not has_sub_parts and ('V1200' in base_norm or 'SwivelArm' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Swivel arm\Swivel arm'
            return r'\Operation\Swivel arm\Settings'
        elif not has_sub_parts and ('V1300' in base_norm or 'TubeTransConv' in base_norm):
            if 'Tab_1' in base:
                return r'\Operation\Tube transport conveyor\Tube transport conveyor'
            return r'\Operation\Tube transport conveyor\Settings'
        elif not has_sub_parts and 'V1900' in base_norm:
            if 'Tab_1' in base:
                return r'\Operation\Basic settings\Speed'
            return r'\Operation\Basic settings\Settings'
        elif not has_sub_parts and 'V0000_00_PowerMonitoringView' in base_norm:
            return r'\Operation\Basic settings\Settings'
        elif not has_sub_parts and 'V0023_00_InsertionView' in base_norm:
            return r'\Operation\Inserting\Settings'
        elif not has_sub_parts and 'V0042_00_CartonErectionView' in base_norm:
            return r'\Operation\Carton sensing\Settings'
        elif not has_sub_parts and 'V0043_00_CartonEjectionView' in base_norm:
            return r'\Operation\Carton ejection\Settings'
        elif not has_sub_parts and 'V0051_00_GluingView' in base_norm:
            if 'GluingFront' in base_norm:
                return r'\Operation\Gluing\Gluing front'
            elif 'GluingRear' in base_norm:
                return r'\Operation\Gluing\Gluing rear'
            return r'\Operation\Gluing\Settings'
        elif not has_sub_parts and 'V0100_00_ProductSensingView' in base_norm:
            return r'\Operation\Product sensing\Product sensing'
        elif not has_sub_parts and 'V0101_01_CartonSensingView' in base_norm:
            return r'\Operation\Carton sensing\Settings'
        elif not has_sub_parts and 'V0103_00_MonDustFlapsView' in base_norm:
            return r'\Operation\Monitoring, dust flap\Settings'
        elif not has_sub_parts and 'V0106_00_MonProductPresenceView' in base_norm:
            return r'\Operation\Monitoring, product presence in carton\Settings'
        elif not has_sub_parts and 'V0151_00_MonCartonCodeView' in base_norm:
            return r'\Operation\Monitoring, carton code\Settings'
        elif not has_sub_parts and 'V0170_00_MonCartonLabelView' in base_norm:
            return r'\Operation\Monitoring, carton labeling\Settings'
        elif not has_sub_parts and 'V0172_00_MonProductLabelView' in base_norm:
            return r'\Operation\Monitoring, product labeling\Settings'
        elif not has_sub_parts and 'V0181_01_CoverTrackView' in base_norm:
            return r'\Operation\Cover rail\Settings'
        elif not has_sub_parts and 'V0201_00_PrinterView' in base_norm:
            return r'\Operation\Carton Labeling\Printer'
        elif not has_sub_parts and 'V0303_00_DischargeConvView' in base_norm:
            return r'\Operation\Discharging conveyor\Settings'
        elif not has_sub_parts and 'V0536_01_StripTransferView' in base_norm:
            return r'\Operation\Strip transfer, unit A\Settings'
        elif not has_sub_parts and 'V0700_00_UpstrMachineView' in base_norm:
            return r'\Operation\Upstream machine\Settings'
        elif not has_sub_parts and 'V0760_00_DownstrMachineView' in base_norm:
            return r'\Operation\Downstream machine\Settings'
        elif not has_sub_parts and 'V0980_00_MechAdjustView' in base_norm:
            if 'Tab_1' in base:
                return r'\Operation\Mechanical adjustment\Basic settings'
            elif 'Tab_2' in base:
                return r'\Operation\Mechanical adjustment\Positions'
            elif 'Tab_3' in base:
                return r'\Operation\Mechanical adjustment\Reference run'
            elif 'Tab_4' in base:
                return r'\Operation\Mechanical adjustment\Tube infeed magazine'
            return r'\Operation\Mechanical adjustment\Settings'
        elif not has_sub_parts and 'V1536_01_StripTransferView' in base_norm:
            return r'\Operation\Strip transfer, unit B\Settings'
        elif not has_sub_parts and 'V2536_01_StripTransferView' in base_norm:
            return r'\Operation\Strip transfer, unit C\Settings'
        elif not has_sub_parts and 'V3536_01_StripTransferView' in base_norm:
            return r'\Operation\Strip transfer, unit D\Settings'
        # Dynamic fallback for any custom or new station code VXXXX
        match = re.search(r'V\d{4}_\d{2}_([A-Za-z0-9]+)View', base_norm)
        if match:
            raw_station = match.group(1)
            clean_map = {
                'BadTubeDisch': 'Faulty tube ejection',
                'GoodTubeDisch': 'Good tube discharge',
                'TubeDischConv': 'Tube discharge conveyor',
                'TubeTransConv': 'Tube transport conveyor',
                'ProdHopperA': 'Product hopper, component A',
                'ProdHopperB': 'Product hopper, component B',
                'BoxChain': 'Box chain',
                'BoxInfeed': 'Box infeed',
                'SwivelArm': 'Swivel arm',
                'LaminarFlow': 'Laminar flow device',
                'MechAdjust': 'Mechanical adjustment',
                'UpstrMachine': 'Upstream machine',
                'DownstrMachine': 'Downstream machine',
                'GasPurge': 'Gas Purge',
                'TubeBlowOut': 'Tube blow out',
                'TubeInfeed': 'Tube infeed',
                'TubePrintReg': 'Tube print registration',
                'TubeFill': 'Tube filling',
                'TubeFillA': 'Tube filling',
                'TubeFillB': 'Tube filling',
                'Fold': 'Folding station',
                'IPCBulkDisch': 'IPC bulk discharge',
                'LaserLabel': 'Laser labeling',
                'TZF': 'Tube infeed magazine',
            }
            def clean_subview_name(sv):
                sv_map = {
                    'NozzleA': 'Nozzles A',
                    'TunnelA': 'Tunnel A',
                    'TunnelB': 'Tunnel B',
                    'FillA': 'Filling station',
                    'SpecialFunctions': 'Special functions',
                    'Timing': 'Timing',
                    'Positions': 'Positions',
                    'GluingFront': 'Gluing front',
                    'GluingRear': 'Gluing rear',
                    'AdjustPos_FillingSystem': 'Positions\\Filling system',
                    'AdjustPos_FoldingStation': 'Positions\\Folding station',
                    'AdjustPos_MonTubeEdge': 'Positions\\Monitoring tube edge',
                    'AdjustPos_MonTubePosition': 'Positions\\Monitoring tube position',
                    'AdjustPos_Press': 'Positions\\Press',
                    'AdjustPos_TubeBlowOut': 'Positions\\Tube blow out',
                    'AdjustPos_TubeDischarge': 'Positions\\Tube discharge',
                    'AdjustPos_TubeInfeed': 'Positions\\Tube infeed',
                    'AdjustPos_TubeLabeling': 'Positions\\Tube labeling',
                    'AdjustPos_TubePrintReg': 'Positions\\Tube print registration',
                    'TZF_AdjustPos_TransConv': 'Positions\\Transfer conveyor',
                    'TZF_AdjustPos_TubeStorageUnit': 'Positions\\Tube storage unit',
                    'TZF_Formatparts': 'Format parts',
                    'DosingA_SpecialFunctions': 'Dosing\\Special functions',
                    'DosingA_Timing': 'Dosing\\Timing',
                    'Lifter_Positions': 'Lifter\\Positions',
                    'Lifter_SpecialFunctions': 'Lifter\\Special functions',
                    'Lifter_Timing': 'Lifter\\Timing'
                }
                if sv in sv_map:
                    return sv_map[sv]
                return re.sub(r'([a-z0-9])([A-Z])', r'\1 \2', sv).strip()

            station_clean = clean_map.get(raw_station, re.sub(r'([a-z0-9])([A-Z])', r'\1 \2', raw_station).strip())
            def format_part(p):
                if p in clean_map:
                    return clean_map[p]
                return clean_subview_name(p)

            subtab = 'Settings'
            if 'Tab_0' in base_clean or 'Tab_00' in base_clean or not re.search(r'Tab_\d+', base_clean):
                subtab = 'Settings'
            else:
                t_match = re.search(r'Tab_(\d+)(?:_(.*))?$', base_clean)
                if t_match:
                    t_idx = t_match.group(1)
                    t_sub = t_match.group(2)
                    if t_idx in ('0', '00'):
                        subtab = 'Settings'
                    else:
                        if t_sub:
                            skip_tokens = {'view', raw_station.lower(), raw_station.lower() + 'a', raw_station.lower() + 'b'}
                            sub_parts = [p for p in t_sub.split('_') if p and p.lower() not in skip_tokens]
                            if sub_parts:
                                full_sub_key = '_'.join(sub_parts)
                                clean_sv = clean_subview_name(full_sub_key)
                                if clean_sv != full_sub_key:
                                    subtab = clean_sv
                                elif len(sub_parts) == 1:
                                    clean_sv = format_part(sub_parts[0])
                                    subtab = f"{station_clean}\\{clean_sv}"
                                else:
                                    p1 = format_part(sub_parts[0])
                                    rest = "-".join([format_part(p) for p in sub_parts[1:]])
                                    subtab = f"{station_clean}\\{p1}-{rest}"
                            else:
                                subtab = f"Tab {t_idx}"
                        else:
                            subtab = f"Tab {t_idx}"
                else:
                    subtab = station_clean
            return f'\\Operation\\{station_clean}\\{subtab}'
        return r'\Operation\Settings'

    # Fallback to OCR title
    sub = ocr_res.get('derived_title') if ocr_res else base
    return f'\\Operation\\{sub}'


def get_category_and_sort_key(fname):
    """
    Returns ((cat_idx, sub_order, name_no_ext), cat_name, name_no_ext) for sorting screen images/templates by 7 Categories:
    1: OPERATION, 2: STATISTICS, 3: FORMATS, 4: HISTORY, 5: DIAGNOSIS, 6: SERVICE, 7: OPERATION (พิเศษ)
    """
    name_no_ext = os.path.splitext(fname)[0]
    if name_no_ext.startswith('View_ProductionSubNavigation') or name_no_ext.startswith('View_RunTaskButton') or 'SetupRefRunView' in name_no_ext or ('RunTaskButton' in name_no_ext and 'View_Settings' not in name_no_ext):
        if 'runtaskbutton' in name_no_ext.lower() and 'refrun' not in name_no_ext.lower():
            sub_order = 1
        elif 'setuptaskbutton' in name_no_ext.lower() and 'refrun' not in name_no_ext.lower():
            sub_order = 2
        elif 'refrun' in name_no_ext.lower():
            sub_order = 3
        else:
            sub_order = 1
        return (1, sub_order, name_no_ext), "OPERATION", name_no_ext
    elif name_no_ext.startswith('View_KPIValuesKSubNavigation'):
        if 'StatisticCounterView_ProductionCounter' in name_no_ext or ('ProductionCounter' in name_no_ext and 'Performance' not in name_no_ext):
            sub_order = 1
        elif 'MonitoringCounterView' in name_no_ext:
            sub_order = 2
        elif 'RejectsHistory' in name_no_ext or 'RejectReasons' in name_no_ext:
            sub_order = 3
        elif 'BatchReportView' in name_no_ext:
            sub_order = 4
        elif 'EfficiencyView' in name_no_ext:
            sub_order = 5
        elif name_no_ext.endswith('KPIValuesView') or name_no_ext.endswith('KPIValues'):
            sub_order = 6
        elif 'Performance' in name_no_ext and ('Counter' in name_no_ext or 'ProductionCounter' in name_no_ext):
            sub_order = 7
        elif 'Performance' in name_no_ext and ('State' in name_no_ext or 'MachineState' in name_no_ext):
            sub_order = 8
        elif 'RejectsView' in name_no_ext or 'FaultsRejects' in name_no_ext:
            sub_order = 9
        elif 'FaultsView' in name_no_ext:
            sub_order = 10
        else:
            sub_order = 99
        return (2, sub_order, name_no_ext), "STATISTICS", name_no_ext
    elif name_no_ext.startswith('View_RecipeSubNavigation'):
        if 'FormatCompare' in name_no_ext:
            sub_order = 2
        elif 'FormatBackup' in name_no_ext:
            sub_order = 3
        else:
            sub_order = 1
        return (3, sub_order, name_no_ext), "FORMATS", name_no_ext
    elif name_no_ext.startswith('View_HistorySubNavigation'):
        if 'HistoryMessagesView' in name_no_ext or 'HistoryMessage' in name_no_ext:
            sub_order = 1
        elif 'StatisticalMessagesView' in name_no_ext or 'StatisticalMessage' in name_no_ext:
            sub_order = 2
        elif 'DiagnosisLoggingView' in name_no_ext or 'Logbook' in name_no_ext:
            sub_order = 3
        else:
            sub_order = 1
        return (4, sub_order, name_no_ext), "HISTORY", name_no_ext
    elif name_no_ext.startswith('View_DiagnosisSubNavigation'):
        if 'ShiftRegister' in name_no_ext or 'Shiftregister' in name_no_ext:
            sub_order = 2
        else:
            sub_order = 1
        return (5, sub_order, name_no_ext), "DIAGNOSIS", name_no_ext
    elif name_no_ext.startswith('View_ServiceSubNavigation'):
        if 'UserAdministrationView_User' in name_no_ext:
            sub_order = 1
        elif 'UserAdministrationView_Group' in name_no_ext:
            sub_order = 2
        elif 'UserAdministrationView_Common' in name_no_ext:
            sub_order = 3
        elif 'ServiceBackup' in name_no_ext:
            sub_order = 4
        elif 'ServiceFileExport' in name_no_ext:
            sub_order = 5
        elif 'ServiceArchiving' in name_no_ext:
            sub_order = 6
        elif 'SystemSettings_Misc' in name_no_ext:
            sub_order = 7
        elif 'SystemSettings_Time' in name_no_ext:
            sub_order = 8
        elif 'SystemSettings_Net' in name_no_ext:
            sub_order = 9
        elif 'TextAdministration' in name_no_ext:
            sub_order = 10
        else:
            sub_order = 99
        return (6, sub_order, name_no_ext), "SERVICE", name_no_ext
    elif name_no_ext.startswith('View_More'):
        if name_no_ext == 'View_MoreView' or name_no_ext == 'View_More':
            sub_order = 0
        elif 'SingleJog' in name_no_ext or 'SJog' in name_no_ext:
            sub_order = 1 if 'TZ_' not in name_no_ext else 15
        elif 'RefRun' in name_no_ext:
            sub_order = 16
        elif 'SetupAbsEnc' in name_no_ext:
            sub_order = 2
        elif 'SetupStripTransferOverview' in name_no_ext:
            if 'Tab_1_StripInfeed' in name_no_ext:
                unit_char = name_no_ext.split('StripInfeed')[-1]
                unit_num = ord(unit_char.upper()) - ord('A') + 1 if unit_char and unit_char.isalpha() else 0
                sub_order = 10 + unit_num
            elif 'Tab_2_StripTransfer' in name_no_ext:
                unit_char = name_no_ext.split('StripTransfer')[-1]
                unit_num = ord(unit_char.upper()) - ord('A') + 1 if unit_char and unit_char.isalpha() else 0
                sub_order = 20 + unit_num
            elif 'Tab_3_StripTransferSlide' in name_no_ext:
                unit_char = name_no_ext.split('StripTransferSlide')[-1]
                unit_num = ord(unit_char.upper()) - ord('A') + 1 if unit_char and unit_char.isalpha() else 0
                sub_order = 30 + unit_num
            elif 'Tab_1' in name_no_ext:
                sub_order = 10
            elif 'Tab_2' in name_no_ext:
                sub_order = 20
            elif 'Tab_3' in name_no_ext:
                sub_order = 30
            else:
                sub_order = 3
        elif 'SetupHeightAdjust' in name_no_ext:
            sub_order = 40
        elif 'SetupTubePrintReg' in name_no_ext:
            sub_order = 41
        elif 'SetupTubeFill' in name_no_ext:
            sub_order = 42
        elif 'SetupPackTags' in name_no_ext:
            if 'Tab_1' in name_no_ext:
                sub_order = 44
            elif 'Tab_2' in name_no_ext:
                sub_order = 45
            elif 'Tab_3' in name_no_ext:
                sub_order = 46
            elif 'Tab_4' in name_no_ext:
                sub_order = 47
            elif 'Tab_5' in name_no_ext:
                sub_order = 48
            elif 'Tab_6' in name_no_ext:
                sub_order = 49
            else:
                sub_order = 43
        elif 'SetupVersionInfo' in name_no_ext:
            sub_order = 90
        else:
            sub_order = 99
        return (7, sub_order, name_no_ext), "(...)", name_no_ext
    elif name_no_ext.startswith('View_Settings') or name_no_ext.startswith('View_SetupTaskButton'):
        match = re.search(r'V(\d{4})', name_no_ext)
        v_num = int(match.group(1)) if match else 9999
        
        tab_match = re.search(r'Tab_(\d+)', name_no_ext)
        tab_num = int(tab_match.group(1)) if tab_match else 0
        
        sub_order = v_num * 100 + tab_num
        return (8, sub_order, name_no_ext), "OPERATION_SPECIAL", name_no_ext
    else:
        return (8, 99999, name_no_ext), "OPERATION_SPECIAL", name_no_ext


def generate_oq_hmi_word(
    excel_source=None,
    word_template_path=None,
    image_source=None,
    order_no=None,
    serial_no=None,
    machine_type=None,
    visu_type="IPC",
    custom_image_order=None
) -> dict:
    """
    Executes OQ HMI protocol generation:
    1. Parses HMI Texts Excel.
    2. Loads Word Master Template.
    3. Dynamically generates Screen View test tables from Excel & Images.
    4. Matches & physically inserts screenshot images into screen test tables.
    5. Generates Image Audit Excel Report.
    Returns dict with file_name, doc_bytes, excel_audit_bytes, metrics, and preview_df.
    """
    import copy

    # 1. Load Excel
    if excel_source is None:
        excel_source = DEFAULT_EXCEL_PATH
    df_hmi, hmi_summary = parse_hmi_texts_excel(excel_source)

    # 2. Get Screenshots
    image_list = get_screenshot_images(image_source)

    # 3. Load Word Master Template
    if word_template_path and os.path.exists(word_template_path):
        template_file = word_template_path
    elif os.path.exists(DEFAULT_MASTER_TEMPLATE):
        template_file = DEFAULT_MASTER_TEMPLATE
    elif os.path.exists(DEFAULT_SAMPLE_TEMPLATE):
        template_file = DEFAULT_SAMPLE_TEMPLATE
    else:
        raise FileNotFoundError("OQ HMI Word Master Template not found.")

    doc = docx.Document(template_file)

    # Preserve Page 1 Master Template Header Metadata 100% untouched (managed by user's dedicated tool)

    # Step 1: Learn Screen Definitions & Parameter Terms from Excel dataset
    excel_terms = []
    if not df_hmi.empty:
        target_rows = df_hmi[df_hmi['GroupPath'] == 'Application.DynamicTexts'] if 'GroupPath' in df_hmi.columns and not df_hmi[df_hmi['GroupPath'] == 'Application.DynamicTexts'].empty else df_hmi

        for _, row in target_rows.iterrows():
            name_val = str(row.get('Name', '')).strip()
            desc_val = str(row.get('Description', '')).strip()
            en_val = str(row.get('IWKEnglish', row.get('L32777', ''))).strip()

            clean_n = re.sub(r'[^a-zA-Z0-9]', '', name_val).replace('idx', '').lower()
            clean_e = re.sub(r'[^a-zA-Z0-9]', '', en_val).lower()

            if clean_n or clean_e or en_val:
                excel_terms.append({
                    'raw_name': name_val,
                    'en_desc': en_val if en_val else name_val,
                    'clean_name': clean_n,
                    'clean_en': clean_e
                })

    # Step 2: Pre-match Images to Excel Learned Terms
    image_excel_map = {}
    for img_info in image_list:
        fname = img_info['name']
        clean_img = re.sub(r'[^a-zA-Z0-9]', '', os.path.splitext(fname)[0]).lower()

        matched_term = None
        for term in excel_terms:
            cn = term['clean_name']
            ce = term['clean_en']
            if (cn and len(cn) > 3 and cn in clean_img) or (ce and len(ce) > 3 and ce in clean_img) or (clean_img and len(clean_img) > 5 and (clean_img in cn or clean_img in ce)):
                matched_term = term
                break

        image_excel_map[fname] = {
            'img_info': img_info,
            'clean_img': clean_img,
            'excel_term': matched_term
        }

    # Helper to map filename & OCR result to exact Reference Standard Table Header Title (based on 56041_04_OQ_HMI)
    # Sort image_list by custom_image_order if provided, or by Category Order (1 to 8) and Filename
    if custom_image_order and isinstance(custom_image_order, list):
        order_map = {fn: i for i, fn in enumerate(custom_image_order)}
        image_list = sorted(image_list, key=lambda x: order_map.get(x['name'], 9999))
    else:
        image_list = sorted(image_list, key=lambda x: get_category_and_sort_key(x['name']))

    inserted_map = {}
    total_screens = 0
    total_parameters = 0

    # Step 1: Perform OCR & Extract Titles from All Images
    image_ocr_map = []
    for img_info in image_list:
        fname = img_info['name']
        sort_order, category, name_no_ext = get_category_and_sort_key(fname)
        ocr_res = perform_ocr_on_image(img_info)
        header_title = get_reference_standard_header(fname, ocr_res)
        image_ocr_map.append({
            'fname': fname,
            'img_info': img_info,
            'ocr_res': ocr_res,
            'category': category,
            'title': header_title
        })

    inserted_map = {}
    total_screens = 0
    total_parameters = 0

    # Step 1.5: Deduplicate images by unique matching sub-docx file
    unique_image_ocr_map = []
    seen_sub_docx = {}

    for item in image_ocr_map:
        t = item['title']
        fname = item['fname']
        matched_sub_docx = find_matching_sub_docx(fname, machine_type)

        if matched_sub_docx:
            if matched_sub_docx in seen_sub_docx:
                primary_fname = seen_sub_docx[matched_sub_docx]
                inserted_map[fname] = {
                    'is_duplicate': True,
                    'primary_fname': primary_fname,
                    'table_title': t,
                    'recheck_status': 'Skipped (Duplicate)',
                    'remarks': f"Duplicate screen view skipped (already covered by '{primary_fname}')"
                }
            else:
                seen_sub_docx[matched_sub_docx] = fname
                unique_image_ocr_map.append(item)
        else:
            # Images without matching sub-docx in Machine Type folder are kept in map so they get marked as Not Inserted
            unique_image_ocr_map.append(item)

    image_ocr_map = unique_image_ocr_map

    # Step 1.6: Automated Post-Generation OCR Recheck (Self-Validation Engine)
    for item in image_ocr_map:
        fname = item['fname']
        header_title = item['title']
        ocr_res = item['ocr_res']
        txt = ocr_res.get('extracted_text', '')
        lines = [l.strip() for l in txt.split('\n') if l.strip()]

        recheck_status = "✅ Verified (OCR Match)"
        recheck_log = f"Header '{header_title}' verified against screenshot OCR."

        # Guard: Main Overview page View_MoreView.jpg must always remain '\(...)'
        if fname in ('View_MoreView.jpg', 'View_More.jpg'):
            item['recheck_status'] = "✅ Verified (Main Overview)"
            item['recheck_log'] = "Main MORE category overview screen verified."
            continue

        # 1. Check for Category 7 (...) MORE screens alignment
        if item.get('category') == '(...)':
            exact_title = None
            for l in lines[:20]:
                l_clean = l.strip()
                l_lower = l_clean.lower()

                # Skip common status/alarm bar messages (e.g. 'Reference run is required')
                if 'reference run is required' in l_lower or 'reference run required' in l_lower:
                    continue

                if 'height adjustment' in l_lower:
                    exact_title = r'\(...)\Height adjustment'
                    break
                elif 'pack tags' in l_lower or 'packtags' in l_lower:
                    if 'Tab_1' in fname or 'messages' in l_lower:
                        exact_title = r'\(...)\PackTags\Messages'
                    elif 'Tab_2' in fname or 'unit mode' in l_lower:
                        exact_title = r'\(...)\PackTags\Unit mode / State'
                    elif 'Tab_3' in fname or 'parameter' in l_lower:
                        exact_title = r'\(...)\PackTags\Parameter'
                    elif 'Tab_4' in fname or 'other' in l_lower:
                        exact_title = r'\(...)\PackTags\Other'
                    elif 'Tab_5' in fname or 'production counter' in l_lower:
                        exact_title = r'\(...)\PackTags\Production counter'
                    elif 'Tab_6' in fname or 'product data' in l_lower:
                        exact_title = r'\(...)\PackTags\Product data'
                    else:
                        exact_title = r'\(...)\PackTags\Machine speed'
                    break
                elif 'strip transfer' in l_lower:
                    if 'Tab_1' in fname or 'strip infeed' in l_lower:
                        exact_title = r'\(...)\Strip transfer\Strip infeed'
                    elif 'Tab_2' in fname or ('strip transfer' in l_lower and 'slide' not in l_lower):
                        exact_title = r'\(...)\Strip transfer\Strip transfer'
                    elif 'Tab_3' in fname or 'slide' in l_lower:
                        exact_title = r'\(...)\Strip transfer\Strip transfer, slide'
                    else:
                        exact_title = r'\(...)\Strip transfer\Process values'
                    break
                elif 'tz - jogging single drive' in l_lower or 'tz — jogging single drive' in l_lower:
                    exact_title = r'\(...)\TZ - jogging single drive'
                    break
                elif 'jogging single drive' in l_lower and 'tz' not in l_lower:
                    exact_title = r'\(...)\Jogging single drive'
                    break
                elif 'tz - reference run' in l_lower or 'tz — reference run' in l_lower:
                    exact_title = r'\(...)\TZ - reference run'
                    break
                elif 'reference run' in l_lower and 'tz' not in l_lower and 'required' not in l_lower:
                    exact_title = r'\(...)\Reference run'
                    break
                elif 'setup single drive' in l_lower or 'absolute encoder' in l_lower:
                    exact_title = r'\(...)\Setup single drive'
                    break
                elif 'setup tube filling' in l_lower:
                    exact_title = r'\(...)\Setup tube filling'
                    break
                elif 'setup tube print registration' in l_lower or 'tube print registration' in l_lower:
                    exact_title = r'\(...)\Setup tube print registration'
                    break
                elif 'version display' in l_lower or 'version info' in l_lower:
                    exact_title = r'\(...)\Version display'
                    break

            if exact_title and item['title'] != exact_title:
                recheck_status = "⚡ Auto-Aligned by OCR Recheck"
                recheck_log = f"Auto-adjusted title from '{header_title}' to '{exact_title}' based on exact screen OCR match."
                item['title'] = exact_title

        # 2. Check for Category 8 (OPERATION Special / Settings) screen alignment
        if item.get('category') == 'OPERATION' and (fname.startswith('View_Settings') or fname.startswith('View_SetupTaskButton')):
            # Universal alias alignment map (e.g. remove component A if screen only says 'Product hopper')
            clean_station_map = [
                ('Product hopper, component A', 'Product hopper', 'component a'),
                ('Product hopper, component B', 'Product hopper, component B', 'component b'),
                ('Bad tube discharge', 'Faulty tube ejection', 'bad tube discharge'),
                ('TZF', 'Tube infeed magazine', 'tube infeed magazine'),
            ]

            for orig_phrase, clean_phrase, check_ocr_kw in clean_station_map:
                if orig_phrase.lower() in item['title'].lower() and check_ocr_kw not in txt.lower():
                    old_t = item['title']
                    pattern = re.compile(re.escape(orig_phrase), re.IGNORECASE)
                    new_t = pattern.sub(clean_phrase, item['title'])
                    if new_t != old_t:
                        recheck_status = "⚡ Auto-Aligned by OCR Recheck"
                        recheck_log = f"Auto-adjusted Category 8 title from '{old_t}' to '{new_t}' based on screen OCR verification."
                        item['title'] = new_t

            # Product Hopper Subtab Dynamic Alignment via parameter OCR
            if 'product hopper' in item['title'].lower():
                full_txt = txt.lower()
                target_sub = None
                if 'target fill level' in full_txt or 'fill level' in full_txt or 'pump off' in full_txt or 'pump on' in full_txt or 'max. supply' in full_txt or 'min. supply' in full_txt:
                    target_sub = 'Monitoring, fill level'
                elif 'gas purging' in full_txt or 'purging interval' in full_txt:
                    target_sub = 'Gas purging'
                elif 'current temperature' in full_txt or 'regulator output' in full_txt:
                    target_sub = 'Process values'
                elif 'target value, speed' in full_txt or '[1 /min]' in full_txt:
                    target_sub = 'Agitator'
                elif 'calibration offset' in full_txt or 'target value, temperature' in full_txt or 'temperature sensor' in full_txt:
                    target_sub = 'Heater'
                
                if target_sub and '\\settings' not in item['title'].lower():
                    new_t = f"\\Operation\\Product hopper\\{target_sub}"
                    if new_t != item['title']:
                        recheck_status = "⚡ Auto-Aligned by OCR Recheck"
                        recheck_log = f"Auto-adjusted Product hopper title from '{item['title']}' to '{new_t}' based on screen parameter OCR verification."
                        item['title'] = new_t

            # B. Universal No-Settings-Tab Auto-Alignment (Works dynamically for ANY station screen)
            if '\\settings' in item['title'].lower() and 'settings' not in txt.lower():
                ignore_kw = {'ready', 'latest message', 'operation', 'speed [cyc/min]', 'pieces', 'format', 'user', 'administrator', 'close', 'transmit', 'reset', 'machine position', 'type', 'min', 'value', 'max', 'unit', 'info', 'required in the format'}
                
                detected_station = None
                for l in lines[:15]:
                    l_clean = l.strip()
                    l_lower = l_clean.lower()
                    if l_lower not in ignore_kw and not any(l_lower.startswith(k) for k in ['ready', 'latest', 'speed', 'pieces', 'format', 'user', 'close', 'transmit', '8/8/']) and len(l_clean) > 2:
                        detected_station = l_clean
                        break
                
                if not detected_station:
                    parts = [p for p in item['title'].split('\\') if p]
                    detected_station = parts[-2] if len(parts) >= 2 else parts[0]

                new_t = f"\\Operation\\{detected_station}\\{detected_station}"
                if new_t != item['title']:
                    recheck_status = "⚡ Auto-Aligned by OCR Recheck"
                    recheck_log = f"Detected no 'Settings' tab on screenshot; auto-aligned title dynamically from '{item['title']}' to '{new_t}' based on screen OCR."
                    item['title'] = new_t

        item['recheck_status'] = recheck_status
        item['recheck_log'] = recheck_log

    # Path A: Blank Master Template (Cloning sub-docx templates from Machine Type folder for each image)
    if len(doc.tables) <= 10:
        template_tbl = doc.tables[4] if len(doc.tables) > 4 else doc.tables[-2]
        template_xml = template_tbl._tbl
        signoff_tbl = doc.tables[-1]

        # Clean Heading 2 'Masks...' paragraph if it contains German template instructions
        for p in doc.paragraphs:
            if 'masks' in p.text.lower() and ('unten stehende' in p.text.lower() or 'tabelle kopieren' in p.text.lower()):
                p.text = "Masks"
                try:
                    p.style = "Heading 2"
                except Exception:
                    pass

        # Locate heading '3.3 Masks' or 'XXXX' placeholder paragraph as insertion reference point
        target_p = None

        # 1. Search for explicit 'XXXX' placeholder paragraph
        for p in doc.paragraphs:
            if 'xxxx' in p.text.lower():
                target_p = p
                break

        # 2. Search for Heading paragraph containing 'masks' or '3.2' heading
        if not target_p:
            for p in doc.paragraphs:
                cleaned = re.sub(r'[^a-z0-9]', '', p.text.lower())
                if cleaned in ['masks', '32masks'] or (p.style.name.startswith('Heading') and 'masks' in p.text.lower()):
                    target_p = p
                    break

        # 3. Fallback: Search for paragraph starting with '3.2' or exact 'masks'
        if not target_p:
            for p in doc.paragraphs:
                t = p.text.strip().lower()
                if t.startswith('3.2') or t == 'masks':
                    target_p = p
                    break

        current_insert_ref = target_p._p if target_p else signoff_tbl._tbl
        remove_target_p = target_p if (target_p and 'xxxx' in target_p.text.lower()) else None

        # Build sub-docx lookup map from selected Machine Type library folder
        mach_dir = os.path.join(MACHINE_LIB_DIR, machine_type.strip()) if machine_type else None
        sub_docx_map = {}
        if mach_dir and os.path.isdir(mach_dir):
            for f in os.listdir(mach_dir):
                if f.lower().endswith(('.docx', '.docm')) and not f.startswith('~$'):
                    f_base = os.path.splitext(f)[0]
                    norm_k = re.sub(r'[^a-zA-Z0-9]', '', f_base).lower()
                    sub_docx_map[norm_k] = os.path.join(mach_dir, f)
                    sub_docx_map[f_base.lower()] = os.path.join(mach_dir, f)

        for idx, item in enumerate(image_ocr_map, 1):
            img_info = item['img_info']
            fname = item['fname']

            # Match sub-docx using unified helper function
            matched_sub_docx = find_matching_sub_docx(fname, machine_type)

            # Rule: If no sub-docx template file exists in Machine Type folder, DO NOT INSERT this image
            if not matched_sub_docx or not os.path.exists(matched_sub_docx):
                continue

            total_screens += 1

            try:
                sub_doc = docx.Document(matched_sub_docx)
                if sub_doc.tables and len(sub_doc.tables[0].rows) > 0:
                    new_tbl_xml = copy.deepcopy(sub_doc.tables[0]._tbl)
                else:
                    new_tbl_xml = copy.deepcopy(template_xml)
            except Exception:
                new_tbl_xml = copy.deepcopy(template_xml)

            # For the FIRST inserted table (total_screens == 1), override pageBreakBefore so table starts directly under Heading 2 '3.3 Masks'
            if total_screens == 1:
                for p_elem in new_tbl_xml.xpath('.//w:p'):
                    pPr = p_elem.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
                    if pPr is None:
                        pPr = docx.oxml.OxmlElement('w:pPr')
                        p_elem.insert(0, pPr)
                    for pbb in pPr.findall('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pageBreakBefore'):
                        pPr.remove(pbb)
                    pbb_override = docx.oxml.OxmlElement('w:pageBreakBefore')
                    pbb_override.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0')
                    pPr.append(pbb_override)

            # Insert right after current_insert_ref (starts under 3.3 Masks)
            current_insert_ref.addnext(new_tbl_xml)
            current_insert_ref = new_tbl_xml

            new_tbl = docx.table.Table(new_tbl_xml, doc)

            # Ensure header row (Row 0 Cell 0) uses style 'Path' and has no bullet 'numPr' override so automatic list numbers format properly (e.g. 108. -> ...)
            try:
                hdr_cell = new_tbl.rows[0].cells[0]
                for p_hdr in hdr_cell.paragraphs:
                    p_hdr.style = "Path"
                    pPr = p_hdr._p.get_or_add_pPr()
                    numPr = pPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr')
                    if numPr is not None:
                        pPr.remove(numPr)
            except Exception:
                pass

            # Insert new screenshot picture into Row 2 Cell 1 (Row 1 Index 0)
            try:
                img_cell = new_tbl.rows[1].cells[0] if len(new_tbl.rows) > 1 else new_tbl.rows[0].cells[0]
                img_cell.text = ""
                p = img_cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)

                # Set cell padding (left/right margins) to 0 for full edge-to-edge screenshot width
                try:
                    tcPr = img_cell._tc.get_or_add_tcPr()
                    tcMar = docx.oxml.OxmlElement('w:tcMar')
                    for side in ['left', 'right']:
                        m = docx.oxml.OxmlElement(f'w:{side}')
                        m.set(docx.oxml.ns.qn('w:w'), '0')
                        m.set(docx.oxml.ns.qn('w:type'), 'dxa')
                        tcMar.append(m)
                    tcPr.append(tcMar)
                except Exception:
                    pass

                target_width = img_cell.width if (img_cell.width and img_cell.width > Inches(4.0)) else Inches(6.25)
                run = p.add_run()

                if img_info.get('is_bytes'):
                    img_stream = io.BytesIO(img_info['path'])
                    run.add_picture(img_stream, width=target_width)
                else:
                    run.add_picture(img_info['path'], width=target_width)

                tbl_header_title = new_tbl.rows[0].cells[0].text.strip() if new_tbl.rows else item['title']
                inserted_map[fname] = {
                    'table_idx': idx,
                    'table_title': tbl_header_title,
                    'recheck_status': '✅ Sub-Template Matched',
                    'remarks': f"Copied sub-template '{os.path.basename(matched_sub_docx)}' & inserted screenshot"
                }
            except Exception as ex:
                pass

            # Add Page Break after each screen view section so next screen starts on a new page
            p_pb = doc.add_paragraph()
            p_pb.add_run().add_break(docx.enum.text.WD_BREAK.PAGE)
            current_insert_ref.addnext(p_pb._p)
            current_insert_ref = p_pb._p

        # Remove original XXXX placeholder paragraph if it existed
        if remove_target_p:
            try:
                remove_target_p._p.getparent().remove(remove_target_p._p)
            except Exception:
                pass

        # Keep original template tables intact (e.g. 4.1 Comments table)
        pass

    # Path B: Multi-table Master Document Template
    else:
        for tbl_idx, tbl in enumerate(doc.tables):
            if not tbl.rows:
                continue

            header_text = tbl.rows[0].cells[0].text.strip()
            if not header_text or header_text.startswith('Machine type') or header_text.startswith('Version') or header_text.startswith('Acceptance'):
                continue

            total_screens += 1
            total_parameters += max(0, len(tbl.rows) - 4)

            norm_hdr = re.sub(r'[^a-zA-Z0-9]', '', header_text).lower()

            matched_img_info = None
            matched_fname = None
            matched_item = None

            # 1. Primary: Match using OCR-derived Titles from image_ocr_map (Highest Accuracy)
            for item in image_ocr_map:
                fname = item['fname']
                if fname in inserted_map:
                    continue

                norm_t = re.sub(r'[^a-zA-Z0-9]', '', item['title']).lower()
                norm_fn = re.sub(r'[^a-zA-Z0-9]', '', os.path.splitext(fname)[0]).lower()

                # Clean kw
                clean_kw = re.sub(r'View|SubNavigation|SettingsView|V\d{4}_\d{2}_|_Tab_\d+', '', os.path.splitext(fname)[0])
                kw = re.sub(r'[^a-zA-Z0-9]', '', clean_kw).lower()

                # Check Tab index
                t_fn = re.search(r'Tab_(\d+)', fname)
                t_fn_num = t_fn.group(1) if t_fn else '0'
                
                t_hdr = re.search(r'(\d+)$', norm_hdr)
                t_hdr_num = t_hdr.group(1) if t_hdr else '0'

                if norm_t and (norm_t == norm_hdr or norm_t in norm_hdr or norm_hdr in norm_t):
                    matched_fname = fname
                    matched_img_info = item['img_info']
                    matched_item = item
                    break
                elif len(kw) >= 3 and kw in norm_hdr and (t_fn_num == t_hdr_num or (t_fn_num in ('0', '00') and t_hdr_num in ('0', '00'))):
                    matched_fname = fname
                    matched_img_info = item['img_info']
                    matched_item = item
                    break
                elif norm_fn in norm_hdr or norm_hdr in norm_fn:
                    matched_fname = fname
                    matched_img_info = item['img_info']
                    matched_item = item
                    break

            # 2. Fallback: Match using image_excel_map
            if not matched_fname:
                for fname, map_info in image_excel_map.items():
                    if fname in inserted_map:
                        continue

                    c_img = map_info['clean_img']
                    ex_term = map_info['excel_term']

                    if ex_term:
                        ce = ex_term['clean_en']
                        cn = ex_term['clean_name']
                        if (ce and len(ce) > 3 and (ce in norm_hdr or norm_hdr in ce)) or (cn and len(cn) > 3 and (cn in norm_hdr or norm_hdr in cn)):
                            matched_fname = fname
                            matched_img_info = map_info['img_info']
                            break

                    if c_img in norm_hdr or norm_hdr in c_img or any(part in c_img for part in norm_hdr.split('\\') if len(part) > 4):
                        matched_fname = fname
                        matched_img_info = map_info['img_info']
                        break

            if matched_img_info and matched_fname:
                try:
                    img_cell = tbl.rows[1].cells[0]
                    img_cell.text = ""
                    p = img_cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)

                    # Set cell padding (left/right margins) to 0 for full edge-to-edge screenshot width
                    try:
                        tcPr = img_cell._tc.get_or_add_tcPr()
                        tcMar = docx.oxml.OxmlElement('w:tcMar')
                        for side in ['left', 'right']:
                            m = docx.oxml.OxmlElement(f'w:{side}')
                            m.set(docx.oxml.ns.qn('w:w'), '0')
                            m.set(docx.oxml.ns.qn('w:type'), 'dxa')
                            tcMar.append(m)
                        tcPr.append(tcMar)
                    except Exception:
                        pass

                    target_width = img_cell.width if (img_cell.width and img_cell.width > Inches(4.0)) else Inches(6.25)
                    run = p.add_run()

                    if matched_img_info.get('is_bytes'):
                        img_stream = io.BytesIO(matched_img_info['path'])
                        run.add_picture(img_stream, width=target_width)
                    else:
                        run.add_picture(matched_img_info['path'], width=target_width)

                    inserted_map[matched_fname] = {
                        'table_idx': tbl_idx + 1,
                        'table_title': header_text,
                        'recheck_status': matched_item.get('recheck_status', '✅ Verified (OCR Match)') if matched_item else '✅ Inserted',
                        'remarks': f"Inserted into Table {tbl_idx + 1}"
                    }
                except Exception as ex:
                    print(f"Error inserting image {matched_fname}: {ex}")

    # Generate Image Audit Excel
    excel_audit_bytes = generate_image_audit_excel(image_list, inserted_map)

    # Output file name
    today_str = datetime.datetime.now().strftime('%Y-%m-%d')
    ord_str = order_no if order_no else "XXXXX"
    output_filename = f"{ord_str}_04_OQ_HMI-V3_NavPar_en_{today_str}_en.docx"

    doc_output = io.BytesIO()
    doc.save(doc_output)
    doc_output.seek(0)
    doc_bytes = doc_output.getvalue()

    # Build Preview DataFrame for UI
    preview_records = []
    inserted_count = 0
    duplicate_count = 0
    for idx, img_info in enumerate(image_list, 1):
        fname = img_info['name']
        match_info = inserted_map.get(fname)
        is_dup = match_info is not None and match_info.get('is_duplicate', False)
        is_ins = match_info is not None and not is_dup

        if is_ins:
            status = '✅ Inserted'
            inserted_count += 1
        elif is_dup:
            status = '⚡ Skipped (Duplicate)'
            duplicate_count += 1
        else:
            status = '⚠️ Not Inserted'

        preview_records.append({
            'No.': idx,
            'Image File Name': fname,
            'Insertion Status': status,
            'Matched Screen / Protocol Section': match_info['table_title'] if match_info else 'N/A'
        })

    preview_df = pd.DataFrame(preview_records)

    return {
        "file_name": output_filename,
        "doc_bytes": doc_bytes,
        "excel_audit_name": f"HMI_Image_Insertion_Report_{ord_str}_{today_str}.xlsx",
        "excel_audit_bytes": excel_audit_bytes,
        "total_screens": total_screens,
        "total_parameters": total_parameters,
        "total_images": len(image_list),
        "inserted_images_count": inserted_count,
        "duplicate_images_count": duplicate_count,
        "not_inserted_images_count": len(image_list) - (inserted_count + duplicate_count),
        "inserted_map": inserted_map,
        "preview_df": preview_df,
        "hmi_texts_df": df_hmi
    }


def generate_checklist_xlsm(export_rows: list[dict]) -> tuple[bytes, str, str]:
    """
    Populates pre-built Sequence_Checklist_Template.xlsm using Excel COM.
    Pre-ticks Categories 1..7 (Operation - MORE) with sequence numbers 1..N and soft green row highlighting (RGB: 217, 234, 211).
    Leaves Operation (พิเศษ) categories unticked for manual user selection.
    Returns (bytes_content, filename, mime_type).
    """
    tmpl_path = os.path.join(BASE_DIR, "IQOQDQ", "OQ_HMI", "Sequence_Checklist_Template.xlsm")
    
    if os.path.exists(tmpl_path):
        try:
            import win32com.client
            import tempfile

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp:
                tmp_out_path = os.path.abspath(tmp.name)

            xl = win32com.client.Dispatch('Excel.Application')
            xl.Visible = False
            xl.DisplayAlerts = False
            
            try:
                wb = xl.Workbooks.Open(os.path.abspath(tmpl_path))
                ws = wb.Worksheets(1)
                total_rows = len(export_rows)

                current_seq = 1

                # Populate data starting from row 2
                for idx, row in enumerate(export_rows, start=2):
                    fn = row.get('File Name', '')
                    sort_key, cat_name, _ = get_category_and_sort_key(fn)
                    is_standard_cat = (sort_key[0] <= 7) or (cat_name in ["OPERATION", "STATISTICS", "FORMATS", "HISTORY", "DIAGNOSIS", "SERVICE", "(...)"])

                    ws.Cells(idx, 4).Value = fn
                    ws.Cells(idx, 5).Value = row.get('Function Code', '')
                    ws.Cells(idx, 6).Value = row.get('Header Title', '')
                    ws.Cells(idx, 7).Value = row.get('Status', '')

                    cell_range = ws.Range(ws.Cells(idx, 1), ws.Cells(idx, 7))

                    if is_standard_cat:
                        ws.Cells(idx, 3).Value = current_seq
                        current_seq += 1
                        cell_range.Interior.Color = 13890265 # RGB(217, 234, 211) - Soft Green
                    else:
                        ws.Cells(idx, 3).Value = ""
                        cell_range.Interior.ColorIndex = -4142 # xlNone

                # Toggle Checkbox value & visibility
                try:
                    for chk in ws.CheckBoxes():
                        row_num = chk.TopLeftCell.Row
                        if row_num <= total_rows + 1:
                            chk.Visible = True
                            fn = export_rows[row_num - 2].get('File Name', '')
                            sort_key, cat_name, _ = get_category_and_sort_key(fn)
                            is_std = (sort_key[0] <= 7) or (cat_name in ["OPERATION", "STATISTICS", "FORMATS", "HISTORY", "DIAGNOSIS", "SERVICE", "(...)"])
                            chk.Value = 1 if is_std else -4146 # 1 = xlOn, -4146 = xlOff
                        else:
                            chk.Visible = False
                except Exception:
                    pass

                ws.Columns("D:G").AutoFit()
                wb.SaveAs(tmp_out_path, 52) # 52 = xlOpenXMLWorkbookMacroEnabled
                wb.Close(False)

                with open(tmp_out_path, "rb") as f:
                    excel_bytes = f.read()

                return excel_bytes, "Sequence_Checklist_Tool.xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12"
            finally:
                xl.Quit()
                try:
                    if os.path.exists(tmp_out_path):
                        os.remove(tmp_out_path)
                except Exception:
                    pass
        except Exception as ex_com:
            pass

    # Fallback to standard openpyxl .xlsx
    df_export = pd.DataFrame(export_rows)
    cols = ["Checkbox", "Select", "ลำดับที่เลือก", "File Name", "Function Code", "Header Title", "Status"]
    for c in cols:
        if c not in df_export.columns:
            df_export[c] = ""
    df_export = df_export[cols]

    xl_buf = io.BytesIO()
    with pd.ExcelWriter(xl_buf, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Sequence_Control')
    return xl_buf.getvalue(), "Sequence_Checklist_Tool.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def parse_checklist_excel(uploaded_file, all_valid_filenames: list[str]) -> tuple[list[str], int]:
    """
    Parses uploaded Excel checklist file and returns (reordered_filenames, selected_count).
    Sorts files strictly by 'ลำดับที่เลือก' (Selected Sequence) Column C if ticked/numbered (1, 2, 3... n).
    """
    if hasattr(uploaded_file, 'seek'):
        try:
            uploaded_file.seek(0)
        except Exception:
            pass

    df = pd.read_excel(uploaded_file)
    
    # Identify Filename column
    fn_col = None
    for col in df.columns:
        c_lower = str(col).strip().lower()
        if any(k in c_lower for k in ['file name', 'filename', 'file', 'fname', 'screen', 'image', 'ชื่อไฟล์']):
            fn_col = col
            break
    if not fn_col:
        for col in df.columns:
            matches = [str(val).strip() for val in df[col].dropna() if any(str(val).strip().endswith(ext) for ext in ['.jpg', '.png', '.jpeg'])]
            if matches:
                fn_col = col
                break
    if not fn_col:
        fn_col = df.columns[3] if len(df.columns) > 3 else df.columns[0]

    # Identify Sequence / ลำดับที่เลือก column (Column C)
    seq_col = None
    for col in df.columns:
        c_lower = str(col).strip().lower()
        if any(k in c_lower for k in ['ลำดับที่เลือก', 'sequence', 'seq', 'order', 'no.', 'no']):
            seq_col = col
            break
    if not seq_col:
        seq_col = df.columns[2] if len(df.columns) > 2 else None

    valid_set = set(all_valid_filenames)
    selected_items = []
    unselected_items = []

    for idx, row in df.iterrows():
        fname = str(row[fn_col]).strip() if pd.notna(row[fn_col]) else None
        if fname and fname in valid_set:
            num_val = None
            if seq_col and pd.notna(row[seq_col]):
                try:
                    val_str = str(row[seq_col]).strip()
                    val_num = float(val_str)
                    if val_num > 0:
                        num_val = val_num
                except ValueError:
                    pass
            if num_val is not None:
                selected_items.append((num_val, idx, fname))
            else:
                unselected_items.append((idx, fname))

    # Auto-resolve duplicate sequence numbers if present in uploaded Excel (e.g. from older VBA macros)
    num_counts = {}
    for item in selected_items:
        num_counts[item[0]] = num_counts.get(item[0], 0) + 1

    if any(c > 1 for c in num_counts.values()):
        items_by_row = sorted(selected_items, key=lambda x: x[1])
        adjusted_items = []
        seen_nums = set()
        current_max = 0
        for num, idx, fn in items_by_row:
            if num in seen_nums:
                new_num = max(current_max + 1, max(seen_nums) + 1)
                adjusted_items.append((new_num, idx, fn))
                seen_nums.add(new_num)
                current_max = max(current_max, new_num)
            else:
                adjusted_items.append((num, idx, fn))
                seen_nums.add(num)
                current_max = max(current_max, num)
        selected_items = adjusted_items

    # Sort selected items strictly by sequence number ascending, breaking ties with original row index
    selected_items.sort(key=lambda x: (x[0], x[1]))
    final_order = [item[2] for item in selected_items]
    seen = set(final_order)

    # Append unselected items in their original row order
    for _, f in unselected_items:
        if f not in seen:
            final_order.append(f)
            seen.add(f)

    # Append any remaining files missing entirely from Excel
    for f in all_valid_filenames:
        if f not in seen:
            final_order.append(f)
            seen.add(f)

    return final_order, len(selected_items)


