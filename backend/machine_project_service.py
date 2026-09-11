"""
Machine Project & Function Database Service
Manages master machine functions (Name, Main Function, Picture, Note, Machine Type), 
project machine lists (Project 5xxxxx, Customer, Machine Type),
dynamic Excel generation using openpyxl with embedded images, and historical filtering.
"""

import os
import io
import zipfile
import re
import sqlite3
import datetime
import logging
from typing import List, Dict, Tuple, Any, Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as OpenPyXLImage
from PIL import Image as PILImage

logger = logging.getLogger(__name__)

DB_PATH = "database/ocr_system.db"
BASE_DATA_DIR = "database/functions_data"
IMAGES_DIR = os.path.join(BASE_DATA_DIR, "images")
EXPORTS_DIR = os.path.join(BASE_DATA_DIR, "exports")

# Standard Machine Types defined by IWK Specifications
STANDARD_MACHINE_TYPES = [
    "Any",
    "TFS 10",
    "TFS 15",
    "TFS 25",
    "TFS 30",
    "TFS 30-3",
    "TFS E",
    "FP 8",
    "FP 10",
    "FP 18",
    "FP 34",
    "FP 46-2",
    "FP 34-2",
    "TZ",
    "TZF",
    "TZS",
    "TZC",
    "TZM",
    "CPC-APC",
    "VI 5",
    "VI 5X",
    "VI10",
    "HC 5",
    "VC 5",
    "CH 4",
    "SC4",
    "SC5",
    "SI5-SI6",
    "CABLI",
    "Overhaul"
]

def get_db_connection():
    os.makedirs("database", exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_machine_database():
    """Create folders and required tables in SQLite if they do not exist."""
    os.makedirs(IMAGES_DIR, exist_ok=True)
    os.makedirs(EXPORTS_DIR, exist_ok=True)

    conn = get_db_connection()
    cursor = conn.cursor()

    # 1. Master Machine Functions
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS machine_master_functions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        function_name TEXT NOT NULL,
        main_function TEXT DEFAULT '',
        image_path TEXT,
        note TEXT,
        category TEXT DEFAULT 'General',
        machine_type TEXT DEFAULT 'Any',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Schema migration: check columns in machine_master_functions
    cursor.execute("PRAGMA table_info(machine_master_functions)")
    cols = [col[1] for col in cursor.fetchall()]
    if "machine_type" not in cols:
        try:
            cursor.execute("ALTER TABLE machine_master_functions ADD COLUMN machine_type TEXT DEFAULT 'Any'")
            conn.commit()
        except Exception:
            pass
    if "main_function" not in cols:
        try:
            cursor.execute("ALTER TABLE machine_master_functions ADD COLUMN main_function TEXT DEFAULT ''")
            conn.commit()
        except Exception:
            pass

    # 2. Machine Project Records
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS machine_project_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_no TEXT NOT NULL,
        customer TEXT NOT NULL,
        machine_type TEXT NOT NULL,
        created_by TEXT DEFAULT 'Operator',
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        excel_filename TEXT,
        notes TEXT
    )
    """)

    # 3. Machine Project Items (Selected functions per project)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS machine_project_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        project_id INTEGER NOT NULL,
        function_id INTEGER,
        function_name TEXT NOT NULL,
        main_function TEXT DEFAULT '',
        image_path TEXT,
        note TEXT,
        category TEXT,
        machine_type TEXT DEFAULT 'Any',
        is_selected INTEGER DEFAULT 1,
        FOREIGN KEY (project_id) REFERENCES machine_project_records (id) ON DELETE CASCADE
    )
    """)

    # Schema migration for machine_project_items
    cursor.execute("PRAGMA table_info(machine_project_items)")
    item_cols = [col[1] for col in cursor.fetchall()]
    if "machine_type" not in item_cols:
        try:
            cursor.execute("ALTER TABLE machine_project_items ADD COLUMN machine_type TEXT DEFAULT 'Any'")
            conn.commit()
        except Exception:
            pass
    if "main_function" not in item_cols:
        try:
            cursor.execute("ALTER TABLE machine_project_items ADD COLUMN main_function TEXT DEFAULT ''")
            conn.commit()
        except Exception:
            pass

    # Seed or backfill main_function for existing entries if empty
    cursor.execute("SELECT COUNT(*) FROM machine_master_functions")
    count = cursor.fetchone()[0]
    if count == 0:
        default_functions = [
            ("Infeed Conveyor", "Infeed Station", None, "Transports incoming empty cartons/bottles to the forming/transfer station with minimum level detection.", "Any"),
            ("Carton Erection Unit", "Forming Station", None, "Extracts and erects cartons using reverse gear box and suction arm system.", "VI 5"),
            ("Leaflet Cross Transport", "Leaflet Station", None, "Pulls, folds, and inserts package leaflets alongside products.", "Any"),
            ("Product Inserter / Tube Inserter", "Insertion Station", None, "Pushes products into open cartons using push finger and mouth piece assembly.", "VI 5"),
            ("Gluing System / Hot Melt Unit", "Sealing Station", None, "Applies hotmelt adhesive onto top and bottom flaps with temperature control.", "Any"),
            ("Tuck-in Closure Bottom/Top", "Closure Station", None, "Folds and tucks carton closure flaps on both upper and lower carton sides.", "VI 5"),
            ("Embossing & Code Reader", "Quality & Inspection", None, "Embosses lot/expiry code and inspects 2D Datamatrix/barcodes for verification.", "Any"),
            ("Discharge Pusher & Reject Station", "Discharge Station", None, "Transfers compliant cartons to outfeed conveyor and diverts defective cartons.", "Any"),
            ("Safety Guarding & Emergency Circuit", "Safety & Monitoring", None, "Safety switches on all protective doors with interconnected emergency stop circuit.", "Any")
        ]
        for name, main_f, img, note, mtype in default_functions:
            cursor.execute(
                "INSERT INTO machine_master_functions (function_name, main_function, image_path, note, category, machine_type) VALUES (?, ?, ?, ?, ?, ?)",
                (name, main_f, img, note, "General", mtype)
            )
        conn.commit()
    else:
        # Backfill main_function from category if empty
        cursor.execute("UPDATE machine_master_functions SET main_function = category WHERE (main_function IS NULL OR main_function = '') AND (category IS NOT NULL AND category != 'General')")
        conn.commit()

    conn.close()

# Initialize DB on module load
init_machine_database()

# ---------------------------------------------------------------------------
# Master Functions Management (CRUD)
# ---------------------------------------------------------------------------
def add_master_function(
    function_name: str,
    main_function: str = "",
    image_bytes: Optional[bytes] = None,
    image_filename: str = "",
    note: str = "",
    category: str = "General",
    machine_type: str = "Any"
) -> Tuple[bool, str]:
    """Add a new master function with optional picture, note, main function, and machine type."""
    if not function_name.strip():
        return False, "Function name cannot be empty."

    try:
        mtype_clean = machine_type.strip() if machine_type else "Any"
        main_f_clean = main_function.strip()
        image_path = None
        if image_bytes and image_filename:
            safe_type = mtype_clean.replace(" ", "_").replace("/", "_")
            safe_name = f"{safe_type}_func_{int(datetime.datetime.now().timestamp())}_{image_filename.replace(' ', '_')}"
            dest = os.path.join(IMAGES_DIR, safe_name)
            with open(dest, "wb") as f:
                f.write(image_bytes)
            image_path = dest

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute(
            "INSERT INTO machine_master_functions (function_name, main_function, image_path, note, category, machine_type) VALUES (?, ?, ?, ?, ?, ?)",
            (function_name.strip(), main_f_clean, image_path, note.strip(), category.strip(), mtype_clean)
        )
        conn.commit()
        conn.close()
        return True, f"Function '{function_name}' ({mtype_clean}) added successfully."
    except Exception as e:
        logger.error(f"Error adding master function: {e}")
        return False, f"Failed to add function: {str(e)}"

def update_master_function(
    func_id: int,
    function_name: str,
    main_function: str = "",
    image_bytes: Optional[bytes] = None,
    image_filename: str = "",
    note: str = "",
    category: str = "General",
    machine_type: str = "Any",
    keep_existing_image: bool = True
) -> Tuple[bool, str]:
    """Update an existing master function."""
    if not function_name.strip():
        return False, "Function name cannot be empty."

    try:
        mtype_clean = machine_type.strip() if machine_type else "Any"
        main_f_clean = main_function.strip()
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT image_path FROM machine_master_functions WHERE id = ?", (func_id,))
        row = cursor.fetchone()
        if not row:
            conn.close()
            return False, "Function not found."

        current_image_path = row["image_path"]
        image_path = current_image_path if keep_existing_image else None

        if image_bytes and image_filename:
            safe_type = mtype_clean.replace(" ", "_").replace("/", "_")
            safe_name = f"{safe_type}_func_{int(datetime.datetime.now().timestamp())}_{image_filename.replace(' ', '_')}"
            dest = os.path.join(IMAGES_DIR, safe_name)
            with open(dest, "wb") as f:
                f.write(image_bytes)
            image_path = dest

        cursor.execute("""
            UPDATE machine_master_functions
            SET function_name = ?, main_function = ?, image_path = ?, note = ?, category = ?, machine_type = ?, updated_at = CURRENT_TIMESTAMP
            WHERE id = ?
        """, (function_name.strip(), main_f_clean, image_path, note.strip(), category.strip(), mtype_clean, func_id))
        conn.commit()
        conn.close()
        return True, f"Function '{function_name}' updated successfully."
    except Exception as e:
        logger.error(f"Error updating master function: {e}")
        return False, f"Failed to update function: {str(e)}"

def delete_master_function(func_id: int) -> Tuple[bool, str]:
    """Delete a master function."""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT function_name, image_path FROM machine_master_functions WHERE id = ?", (func_id,))
        row = cursor.fetchone()
        if not row:
            conn.close()
            return False, "Function not found."

        name = row["function_name"]
        img_path = row["image_path"]
        if img_path and os.path.exists(img_path):
            try:
                os.remove(img_path)
            except Exception:
                pass

        cursor.execute("DELETE FROM machine_master_functions WHERE id = ?", (func_id,))
        conn.commit()
        conn.close()
        return True, f"Function '{name}' deleted successfully."
    except Exception as e:
        logger.error(f"Error deleting master function: {e}")
        return False, f"Failed to delete function: {str(e)}"

def get_master_functions(search: str = "", category: str = "", machine_type: str = "", main_function: str = "") -> List[Dict[str, Any]]:
    """Retrieve list of all master functions with optional search, machine_type, and main_function filter."""
    conn = get_db_connection()
    cursor = conn.cursor()

    query = "SELECT * FROM machine_master_functions WHERE 1=1"
    params = []

    if search.strip():
        query += " AND (function_name LIKE ? OR note LIKE ? OR main_function LIKE ?)"
        params.extend([f"%{search.strip()}%", f"%{search.strip()}%", f"%{search.strip()}%"])

    if main_function.strip() and main_function != "All":
        query += " AND main_function = ?"
        params.append(main_function.strip())

    if category.strip() and category != "All":
        query += " AND category = ?"
        params.append(category.strip())

    if machine_type.strip() and machine_type not in ["All", "Any"]:
        query += " AND (machine_type = ? OR machine_type = 'Any' OR machine_type = 'All Types' OR machine_type IS NULL OR machine_type = '')"
        params.append(machine_type.strip())

    query += " ORDER BY id ASC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    results = [dict(r) for r in rows]
    conn.close()
    return results

def get_pictures_by_machine_type(machine_type: str = "Any") -> List[Dict[str, Any]]:
    """Retrieve all pictures and functions associated with a specific machine type."""
    conn = get_db_connection()
    cursor = conn.cursor()
    if machine_type in ["All", "Any", ""]:
        cursor.execute("SELECT id, function_name, main_function, category, machine_type, image_path, note FROM machine_master_functions WHERE image_path IS NOT NULL AND image_path != ''")
    else:
        cursor.execute("""
            SELECT id, function_name, main_function, category, machine_type, image_path, note 
            FROM machine_master_functions 
            WHERE image_path IS NOT NULL AND image_path != '' 
            AND (machine_type = ? OR machine_type = 'Any' OR machine_type = 'All Types')
        """, (machine_type.strip(),))
    rows = cursor.fetchall()
    results = [dict(r) for r in rows]
    conn.close()
    return results

# ---------------------------------------------------------------------------
# Project Machine List & Excel Export
# ---------------------------------------------------------------------------
def generate_project_excel(
    project_no: str,
    customer: str,
    machine_type: str,
    selected_functions: List[Dict[str, Any]],
    created_at: str,
    created_by: str = "Operator",
    notes: str = ""
) -> bytes:
    """
    Generate a professional Excel document (.xlsx) using openpyxl
    with embedded function pictures, project details, notes, and metadata.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Project_{project_no}"
    ws.views.sheetView[0].showGridLines = True

    # Styling Palettes
    header_fill = PatternFill(start_color="0B1A30", end_color="0B1A30", fill_type="solid")
    header_font = Font(name="Segoe UI", size=14, bold=True, color="FFFFFF")
    
    subheader_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    subheader_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")

    th_fill = PatternFill(start_color="0078D4", end_color="0078D4", fill_type="solid")
    th_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

    bold_font = Font(name="Segoe UI", size=10, bold=True)
    normal_font = Font(name="Segoe UI", size=10)

    thin_border_side = Side(border_style="thin", color="CCCCCC")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center")

    # 1. Main Header Title
    ws.merge_cells("A1:G2")
    title_cell = ws["A1"]
    title_cell.value = "IWK MACHINE CONFIGURATION & FUNCTION LIST"
    title_cell.font = header_font
    title_cell.fill = header_fill
    title_cell.alignment = align_center

    # 2. Project Information Block
    ws.merge_cells("A3:G3")
    sec1_cell = ws["A3"]
    sec1_cell.value = "📋 PROJECT SPECIFICATION SUMMARY"
    sec1_cell.font = subheader_font
    sec1_cell.fill = subheader_fill
    sec1_cell.alignment = align_left

    meta_rows = [
        ("Project No:", project_no, "Machine Type:", machine_type),
        ("Customer Name:", customer, "Created By:", created_by),
        ("Export Date / Time:", created_at, "Total Functions Selected:", len(selected_functions)),
    ]

    for row_idx, (k1, v1, k2, v2) in enumerate(meta_rows, start=4):
        ws.cell(row=row_idx, column=1, value=k1).font = bold_font
        ws.cell(row=row_idx, column=2, value=str(v1)).font = normal_font
        ws.cell(row=row_idx, column=4, value=k2).font = bold_font
        ws.cell(row=row_idx, column=5, value=str(v2)).font = normal_font
        
        for c in range(1, 8):
            ws.cell(row=row_idx, column=c).border = thin_border

    if notes.strip():
        r_note = 7
        ws.cell(row=r_note, column=1, value="Project Notes:").font = bold_font
        ws.merge_cells(start_row=r_note, start_column=2, end_row=r_note, end_column=7)
        ws.cell(row=r_note, column=2, value=notes).font = normal_font
        for c in range(1, 8):
            ws.cell(row=r_note, column=c).border = thin_border
        table_start_row = 9
    else:
        table_start_row = 8

    # 3. Selected Functions Table Header
    headers = ["Item", "Main Function", "Function Name", "Machine Type", "Picture", "Function Note / Description", "Status"]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=table_start_row, column=col_idx, value=h)
        cell.font = th_font
        cell.fill = th_fill
        cell.alignment = align_center
        cell.border = thin_border
    ws.row_dimensions[table_start_row].height = 26

    # 4. Table Data Rows with Embedded Pictures
    current_row = table_start_row + 1
    for idx, func in enumerate(selected_functions, start=1):
        img_path = func.get("image_path")
        has_pic = bool(img_path and os.path.exists(img_path))
        m_type = func.get("machine_type") or "Any"
        main_f = func.get("main_function") or "-"
        
        c_item = ws.cell(row=current_row, column=1, value=idx)
        c_item.alignment = align_center
        c_item.font = normal_font
        c_item.border = thin_border

        c_main = ws.cell(row=current_row, column=2, value=main_f)
        c_main.alignment = align_center
        c_main.font = normal_font
        c_main.border = thin_border

        c_name = ws.cell(row=current_row, column=3, value=func.get("function_name", ""))
        c_name.alignment = align_left
        c_name.font = bold_font
        c_name.border = thin_border

        c_mtype = ws.cell(row=current_row, column=4, value=m_type)
        c_mtype.alignment = align_center
        c_mtype.font = normal_font
        c_mtype.border = thin_border

        # Column 5: Picture Function (4.2)
        c_pic = ws.cell(row=current_row, column=5)
        c_pic.alignment = align_center
        c_pic.font = normal_font
        c_pic.border = thin_border

        if has_pic:
            try:
                pil_img = PILImage.open(img_path)
                # Resize image proportionally to fit inside the cell
                pil_img.thumbnail((120, 58), PILImage.Resampling.LANCZOS)
                temp_img_io = io.BytesIO()
                pil_img.convert("RGB").save(temp_img_io, format="PNG")
                temp_img_io.seek(0)

                xl_img = OpenPyXLImage(temp_img_io)
                xl_img.width = pil_img.width
                xl_img.height = pil_img.height

                # Embed image in column E cell
                ws.add_image(xl_img, f"E{current_row}")
                c_pic.value = ""
                ws.row_dimensions[current_row].height = 52
            except Exception as e:
                logger.warning(f"Could not embed picture for function {func.get('function_name')}: {e}")
                c_pic.value = "🖼️ (Attached)"
                ws.row_dimensions[current_row].height = 26
        else:
            c_pic.value = "(No Picture)"
            ws.row_dimensions[current_row].height = 26

        c_note = ws.cell(row=current_row, column=6, value=func.get("note", ""))
        c_note.alignment = align_left
        c_note.font = normal_font
        c_note.border = thin_border

        c_status = ws.cell(row=current_row, column=7, value="Configured")
        c_status.alignment = align_center
        c_status.font = normal_font
        c_status.border = thin_border

        current_row += 1

    # Auto-adjust column widths for Sheet 1
    column_widths = {1: 8, 2: 24, 3: 36, 4: 16, 5: 24, 6: 45, 7: 14}
    for col_idx, width in column_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    # =========================================================================
    # SHEET 2: Complete Function Matrix for the Machine Type (Green Highlights)
    # =========================================================================
    safe_mtype = re.sub(r'[^a-zA-Z0-9_\-]', '_', machine_type.strip()) or "Machine"
    sheet2_title = f"All_{safe_mtype[:15]}_Matrix"
    ws2 = wb.create_sheet(title=sheet2_title)
    ws2.views.sheetView[0].showGridLines = True

    # Palette for Sheet 2
    matrix_th_fill = PatternFill(start_color="0E5A8A", end_color="0E5A8A", fill_type="solid")
    matrix_th_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")

    # Green Highlights for Selected Items
    row_green_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    status_selected_fill = PatternFill(start_color="C3E6CB", end_color="C3E6CB", fill_type="solid")
    status_selected_font = Font(name="Segoe UI", size=10, bold=True, color="155724")

    # Neutral for Non-selected Items
    status_unsel_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    status_unsel_font = Font(name="Segoe UI", size=10, color="6C757D")

    # 1. Sheet 2 Main Header
    ws2.merge_cells("A1:G2")
    t2 = ws2["A1"]
    t2.value = f"IWK COMPLETE FUNCTION MATRIX — TYPE: {machine_type.upper()}"
    t2.font = header_font
    t2.fill = header_fill
    t2.alignment = align_center

    # 2. Sheet 2 Information Block
    ws2.merge_cells("A3:G3")
    sec2 = ws2["A3"]
    sec2.value = "📑 COMPLETE FUNCTIONS MASTER LIST (🟢 GREEN = SELECTED IN PROJECT | ⚪ WHITE = OPTIONAL)"
    sec2.font = subheader_font
    sec2.fill = subheader_fill
    sec2.alignment = align_left

    # Retrieve all available functions for this machine type
    all_type_funcs = get_master_functions(machine_type=machine_type)
    selected_func_names = {f.get("function_name", "").strip() for f in selected_functions}
    selected_func_ids = {f.get("id") for f in selected_functions if f.get("id") is not None} | {f.get("function_id") for f in selected_functions if f.get("function_id") is not None}

    meta_rows2 = [
        ("Target Machine Type:", machine_type, "Project No / Scope:", f"{project_no} ({customer})"),
        ("Total Functions Scope:", len(all_type_funcs), "Selected for Project:", f"{len(selected_functions)} / {len(all_type_funcs)} Included"),
        ("Matrix Export Date:", created_at, "Selection Status:", f"🟢 {len(selected_functions)} Included | ⚪ {len(all_type_funcs) - len(selected_functions)} Optional")
    ]

    for row_idx, (k1, v1, k2, v2) in enumerate(meta_rows2, start=4):
        ws2.cell(row=row_idx, column=1, value=k1).font = bold_font
        ws2.cell(row=row_idx, column=2, value=str(v1)).font = normal_font
        ws2.cell(row=row_idx, column=4, value=k2).font = bold_font
        ws2.cell(row=row_idx, column=5, value=str(v2)).font = normal_font
        for c in range(1, 8):
            ws2.cell(row=row_idx, column=c).border = thin_border

    table2_start_row = 8

    # 3. Sheet 2 Table Header
    headers2 = ["Item", "Main Function", "Function Name", "Machine Type", "Picture", "Function Note / Description", "Project Status"]
    for col_idx, h in enumerate(headers2, start=1):
        cell = ws2.cell(row=table2_start_row, column=col_idx, value=h)
        cell.font = matrix_th_font
        cell.fill = matrix_th_fill
        cell.alignment = align_center
        cell.border = thin_border
    ws2.row_dimensions[table2_start_row].height = 26

    # 4. Sheet 2 Data Rows (Highlighted in Green if selected)
    curr_r2 = table2_start_row + 1
    for idx, func in enumerate(all_type_funcs, start=1):
        fname = func.get("function_name", "").strip()
        fid = func.get("id")
        is_selected = (fid in selected_func_ids) or (fname in selected_func_names)

        img_path = func.get("image_path")
        has_pic = bool(img_path and os.path.exists(img_path))
        m_type = func.get("machine_type") or "Any"
        main_f = func.get("main_function") or "-"

        # Item Number
        c_item = ws2.cell(row=curr_r2, column=1, value=idx)
        c_item.alignment = align_center
        c_item.font = bold_font if is_selected else normal_font
        c_item.border = thin_border

        # Main Function
        c_main = ws2.cell(row=curr_r2, column=2, value=main_f)
        c_main.alignment = align_center
        c_main.font = normal_font
        c_main.border = thin_border

        # Function Name
        c_name = ws2.cell(row=curr_r2, column=3, value=fname)
        c_name.alignment = align_left
        c_name.font = bold_font
        c_name.border = thin_border

        # Machine Type
        c_mtype = ws2.cell(row=curr_r2, column=4, value=m_type)
        c_mtype.alignment = align_center
        c_mtype.font = normal_font
        c_mtype.border = thin_border

        # Picture Function
        c_pic = ws2.cell(row=curr_r2, column=5)
        c_pic.alignment = align_center
        c_pic.font = normal_font
        c_pic.border = thin_border

        if has_pic:
            try:
                pil_img = PILImage.open(img_path)
                pil_img.thumbnail((120, 58), PILImage.Resampling.LANCZOS)
                temp_img_io = io.BytesIO()
                pil_img.convert("RGB").save(temp_img_io, format="PNG")
                temp_img_io.seek(0)

                xl_img = OpenPyXLImage(temp_img_io)
                xl_img.width = pil_img.width
                xl_img.height = pil_img.height

                ws2.add_image(xl_img, f"E{curr_r2}")
                c_pic.value = ""
                ws2.row_dimensions[curr_r2].height = 52
            except Exception:
                c_pic.value = "🖼️ (Attached)"
                ws2.row_dimensions[curr_r2].height = 26
        else:
            c_pic.value = "(No Picture)"
            ws2.row_dimensions[curr_r2].height = 26

        # Note / Description
        c_note = ws2.cell(row=curr_r2, column=6, value=func.get("note", ""))
        c_note.alignment = align_left
        c_note.font = normal_font
        c_note.border = thin_border

        # Project Status (Green for selected, Neutral for optional)
        c_status = ws2.cell(row=curr_r2, column=7)
        c_status.alignment = align_center
        c_status.border = thin_border

        if is_selected:
            # Highlight entire row in soft green
            for c_idx in range(1, 7):
                ws2.cell(row=curr_r2, column=c_idx).fill = row_green_fill
            
            c_status.value = "✅ SELECTED (Included)"
            c_status.fill = status_selected_fill
            c_status.font = status_selected_font
        else:
            c_status.value = "⚪ Optional (Not Included)"
            c_status.fill = status_unsel_fill
            c_status.font = status_unsel_font

        curr_r2 += 1

    column_widths2 = {1: 8, 2: 24, 3: 36, 4: 16, 5: 24, 6: 45, 7: 24}
    for col_idx, width in column_widths2.items():
        ws2.column_dimensions[get_column_letter(col_idx)].width = width

    # Save to BytesIO
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

def save_project_record(
    project_no: str,
    customer: str,
    machine_type: str,
    selected_function_ids: List[int],
    created_by: str = "Operator",
    notes: str = ""
) -> Tuple[bool, str, Optional[int], Optional[bytes]]:
    """
    Save the project machine record and its selected functions to SQLite,
    generate the Excel file, and store it for instant and future downloads.
    """
    if not project_no.strip():
        return False, "Project No cannot be empty (e.g. 5xxxxx).", None, None
    if not customer.strip():
        return False, "Customer name cannot be empty.", None, None
    if not machine_type.strip():
        return False, "Machine Type cannot be empty.", None, None
    if not selected_function_ids:
        return False, "Please select at least one function for this machine.", None, None

    try:
        now_str = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        safe_proj = project_no.strip().replace(" ", "_")
        excel_filename = f"Project_{safe_proj}_{customer.strip().replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        conn = get_db_connection()
        cursor = conn.cursor()

        # 1. Insert Project Header
        cursor.execute("""
            INSERT INTO machine_project_records (project_no, customer, machine_type, created_by, created_at, excel_filename, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (project_no.strip(), customer.strip(), machine_type.strip(), created_by, now_str, excel_filename, notes.strip()))
        project_id = cursor.lastrowid

        # 2. Fetch Selected Master Functions & Insert Project Items
        placeholders = ",".join("?" for _ in selected_function_ids)
        cursor.execute(f"SELECT * FROM machine_master_functions WHERE id IN ({placeholders})", selected_function_ids)
        selected_rows = [dict(r) for r in cursor.fetchall()]

        for item in selected_rows:
            cursor.execute("""
                INSERT INTO machine_project_items (project_id, function_id, function_name, main_function, image_path, note, category, machine_type, is_selected)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
            """, (project_id, item["id"], item["function_name"], item.get("main_function", ""), item["image_path"], item["note"], item["category"], item.get("machine_type", "Any")))

        conn.commit()
        conn.close()

        # 3. Generate Excel file
        excel_bytes = generate_project_excel(
            project_no=project_no.strip(),
            customer=customer.strip(),
            machine_type=machine_type.strip(),
            selected_functions=selected_rows,
            created_at=now_str,
            created_by=created_by,
            notes=notes.strip()
        )

        # 4. Save Excel file to exports folder
        excel_dest = os.path.join(EXPORTS_DIR, excel_filename)
        with open(excel_dest, "wb") as f:
            f.write(excel_bytes)

        return True, f"Project '{project_no}' ({customer}) processed and saved successfully!", project_id, excel_bytes
    except Exception as e:
        logger.error(f"Error saving project record: {e}")
        return False, f"Failed to process project: {str(e)}", None, None

def get_project_records(
    filter_project: str = "",
    filter_customer: str = "",
    filter_type: str = ""
) -> List[Dict[str, Any]]:
    """Retrieve saved project machine records with filtering capabilities."""
    conn = get_db_connection()
    cursor = conn.cursor()

    query = """
        SELECT r.*, 
               (SELECT COUNT(*) FROM machine_project_items i WHERE i.project_id = r.id) as total_functions
        FROM machine_project_records r
        WHERE 1=1
    """
    params = []

    if filter_project.strip():
        query += " AND r.project_no LIKE ?"
        params.append(f"%{filter_project.strip()}%")

    if filter_customer.strip():
        query += " AND r.customer LIKE ?"
        params.append(f"%{filter_customer.strip()}%")

    if filter_type.strip() and filter_type not in ["All", "Any"]:
        query += " AND r.machine_type = ?"
        params.append(filter_type.strip())

    query += " ORDER BY r.id DESC"
    cursor.execute(query, params)
    rows = cursor.fetchall()
    results = [dict(r) for r in rows]
    conn.close()
    return results

def get_project_details(project_id: int) -> Optional[Dict[str, Any]]:
    """Get full details of a project including its selected function items."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM machine_project_records WHERE id = ?", (project_id,))
    rec = cursor.fetchone()
    if not rec:
        conn.close()
        return None

    project_data = dict(rec)
    cursor.execute("SELECT * FROM machine_project_items WHERE project_id = ? ORDER BY id ASC", (project_id,))
    items = [dict(r) for r in cursor.fetchall()]
    project_data["items"] = items
    conn.close()
    return project_data

def get_project_excel_bytes(project_id: int) -> Optional[Tuple[str, bytes]]:
    """Fetch existing Excel file for a project or regenerate it if missing."""
    details = get_project_details(project_id)
    if not details:
        return None

    filename = details.get("excel_filename", f"Project_{details['project_no']}.xlsx")
    file_path = os.path.join(EXPORTS_DIR, filename)

    if os.path.exists(file_path):
        with open(file_path, "rb") as f:
            return filename, f.read()
    else:
        # Regenerate on the fly
        excel_bytes = generate_project_excel(
            project_no=details["project_no"],
            customer=details["customer"],
            machine_type=details["machine_type"],
            selected_functions=details["items"],
            created_at=details["created_at"],
            created_by=details.get("created_by", "Operator"),
            notes=details.get("notes", "")
        )
        return filename, excel_bytes

def generate_project_zip_bundle(project_id: int) -> Optional[Tuple[str, bytes]]:
    """
    Generate a complete ZIP package containing:
    1. The project Excel document (.xlsx) with embedded pictures
    2. An 'images/' folder with all high-resolution pictures for each function
    """
    details = get_project_details(project_id)
    if not details:
        return None

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        # 1. Add Excel file
        excel_bytes = generate_project_excel(
            project_no=details["project_no"],
            customer=details["customer"],
            machine_type=details["machine_type"],
            selected_functions=details["items"],
            created_at=details["created_at"],
            created_by=details.get("created_by", "Operator"),
            notes=details.get("notes", "")
        )
        safe_proj = details["project_no"].replace(" ", "_")
        safe_cust = details["customer"].replace(" ", "_")
        excel_name = f"Project_{safe_proj}_{safe_cust}.xlsx"
        zip_file.writestr(excel_name, excel_bytes)

        # 2. Add high-res function images into images/ subfolder
        for idx, item in enumerate(details["items"], start=1):
            img_path = item.get("image_path")
            if img_path and os.path.exists(img_path):
                ext = os.path.splitext(img_path)[1] or ".png"
                clean_name = re.sub(r'[^a-zA-Z0-9_\-]', '_', item["function_name"])
                zip_img_name = f"images/{idx:02d}_{clean_name}{ext}"
                try:
                    with open(img_path, "rb") as f:
                        zip_file.writestr(zip_img_name, f.read())
                except Exception as e:
                    logger.warning(f"Failed to add image to zip: {e}")

    zip_buffer.seek(0)
    zip_filename = f"Project_{safe_proj}_{safe_cust}_Package.zip"
    return zip_filename, zip_buffer.getvalue()

def delete_project_record(project_id: int) -> Tuple[bool, str]:
    """Delete a project record and its associated items and Excel file."""
    try:
        details = get_project_details(project_id)
        if not details:
            return False, "Project record not found."

        filename = details.get("excel_filename")
        if filename:
            file_path = os.path.join(EXPORTS_DIR, filename)
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except Exception:
                    pass

        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM machine_project_items WHERE project_id = ?", (project_id,))
        cursor.execute("DELETE FROM machine_project_records WHERE id = ?", (project_id,))
        conn.commit()
        conn.close()
        return True, f"Project '{details['project_no']}' deleted successfully."
    except Exception as e:
        logger.error(f"Error deleting project record: {e}")
        return False, f"Failed to delete project: {str(e)}"
