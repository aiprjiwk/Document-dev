import os
import sqlite3
import json
import uuid
import datetime
import difflib
import re
import numpy as np
from PIL import Image

DB_DIR = "database"
DB_FILE = os.path.join(DB_DIR, "ocr_system.db")
TRAINING_DIR = os.path.join("training", "corrected_samples")
LOG_FILE = "ocr_processing_log.csv"

# Ensure directories exist
os.makedirs(DB_DIR, exist_ok=True)
os.makedirs(TRAINING_DIR, exist_ok=True)

def get_db_connection():
    """
    Establishes connection to the local SQLite database.
    """
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    """
    Initializes the database schema if the tables do not exist.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. document_master
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS document_master (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT NOT NULL,
        uploaded_at TEXT NOT NULL,
        status TEXT NOT NULL
    )
    """)
    
    # 2. extracted_fields
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS extracted_fields (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        document_id INTEGER NOT NULL,
        row_num INTEGER NOT NULL,
        de_label TEXT,
        en_label TEXT,
        original_value TEXT,
        extracted_value TEXT,
        confidence REAL,
        approved INTEGER DEFAULT 0,
        page_num INTEGER DEFAULT 0,
        bbox_json TEXT,
        FOREIGN KEY (document_id) REFERENCES document_master (id) ON DELETE CASCADE
    )
    """)
    
    # 3. MappingTable (Allows dynamic cell-based Excel mapping)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS MappingTable (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        FieldName TEXT UNIQUE,
        SheetName TEXT NOT NULL,
        CellAddress TEXT NOT NULL
    )
    """)
    
    # 4. TrainingData (Human corrections log for models learning feedback loop)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS TrainingData (
        id TEXT PRIMARY KEY,
        DocumentID INTEGER,
        FieldName TEXT,
        OCRValue TEXT,
        CorrectedValue TEXT,
        Confidence REAL,
        ReviewedBy TEXT,
        ReviewedDate TEXT,
        Status TEXT
    )
    """)
    
    # 5. ExportHistory (Audits file exports)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS ExportHistory (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        DocumentID INTEGER,
        TemplateName TEXT,
        OutputFileName TEXT,
        ExportedBy TEXT,
        ExportedDate TEXT,
        ExportStatus TEXT
    )
    """)
    
    # 6. handwriting_learning (Visual hashes lookup)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS handwriting_learning (
        id TEXT PRIMARY KEY,
        original_ocr_text TEXT,
        corrected_text TEXT,
        anchor_label TEXT,
        confidence REAL,
        image_hash TEXT,
        image_path TEXT,
        timestamp TEXT
    )
    """)
    
    # Seed default mapping rows if empty
    # Upgrade migration: if old template seeds exist (using 'Master' sheet name), clear to re-seed
    cursor.execute("SELECT COUNT(*) as cnt FROM MappingTable WHERE SheetName = 'Master'")
    old_row = cursor.fetchone()
    if old_row['cnt'] > 0:
        cursor.execute("DELETE FROM MappingTable")
        
    cursor.execute("SELECT COUNT(*) as cnt FROM MappingTable")
    row = cursor.fetchone()
    if row['cnt'] == 0:
        default_mappings = [
            ("Auftrags-Nr. / Order-no.", "Adjustment Table", "B3"),
            ("Bau-Nr. / Serial-no.", "Adjustment Table", "B4"),
            ("Kunde / Customer", "Adjustment Table", "B19"),
            ("Case discharge", "Adjustment Table", "F12"),
            ("Case turning unit", "Adjustment Table", "F15"),
            ("Inserting pusher", "Adjustment Table", "F21"),
            ("Stacking pusher", "Adjustment Table", "F25"),
        ]
        for f, s, c in default_mappings:
            try:
                cursor.execute(
                    "INSERT OR IGNORE INTO MappingTable (FieldName, SheetName, CellAddress) VALUES (?, ?, ?)",
                    (f, s, c)
                )
            except sqlite3.IntegrityError:
                pass
                
    conn.commit()
    conn.close()

# Initialize schemas on import
init_db()

# ---------------------------------------------------------------------------
# MappingTable CRUD Operations
# ---------------------------------------------------------------------------

def get_all_mappings() -> list:
    """
    Returns all mapping rows from MappingTable.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM MappingTable ORDER BY id ASC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_mapping(field_name: str, sheet_name: str, cell_address: str) -> bool:
    """
    Adds a new cell mapping to MappingTable.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    success = False
    try:
        cursor.execute(
            "INSERT INTO MappingTable (FieldName, SheetName, CellAddress) VALUES (?, ?, ?)",
            (field_name.strip(), sheet_name.strip(), cell_address.strip().upper())
        )
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        pass
    conn.close()
    return success

def update_mapping(mapping_id: int, sheet_name: str, cell_address: str):
    """
    Updates the target sheet name or cell address of a mapping.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE MappingTable SET SheetName = ?, CellAddress = ? WHERE id = ?",
        (sheet_name.strip(), cell_address.strip().upper(), mapping_id)
    )
    conn.commit()
    conn.close()

def delete_mapping(mapping_id: int):
    """
    Deletes a mapping row.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM MappingTable WHERE id = ?", (mapping_id,))
    conn.commit()
    conn.close()

def get_mapping_by_label(de_label: str, en_label: str) -> dict:
    """
    Looks up custom cell mapping settings in MappingTable.
    Fuzzy matches against labels.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM MappingTable")
    mappings = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    clean_de = re.sub(r'[^a-zA-Z0-9ก-๙]', '', de_label).lower()
    clean_en = re.sub(r'[^a-zA-Z0-9ก-๙]', '', en_label).lower()
    
    for m in mappings:
        m_field_clean = re.sub(r'[^a-zA-Z0-9ก-๙]', '', m['FieldName']).lower()
        # Direct check
        if (clean_de and clean_de in m_field_clean) or (clean_en and clean_en in m_field_clean) or (m_field_clean in clean_de) or (m_field_clean in clean_en):
            return m
            
    return None

# ---------------------------------------------------------------------------
# TrainingData Logging Operations
# ---------------------------------------------------------------------------

def log_training_data(doc_id: int, field_name: str, ocr_value: str, corrected_value: str, confidence: float, status="Corrected", user_name="User"):
    """
    Saves verified human correction data to the TrainingData table.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    new_id = str(uuid.uuid4())
    now_str = datetime.date.today().isoformat()
    cursor.execute("""
    INSERT INTO TrainingData (id, DocumentID, FieldName, OCRValue, CorrectedValue, Confidence, ReviewedBy, ReviewedDate, Status)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (new_id, doc_id, field_name, ocr_value, corrected_value, confidence, user_name, now_str, status))
    conn.commit()
    conn.close()

def get_all_training_data() -> list:
    """
    Returns all correction feedback datasets.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM TrainingData ORDER BY ReviewedDate DESC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

# ---------------------------------------------------------------------------
# ExportHistory Logging Operations
# ---------------------------------------------------------------------------

def log_export_history(doc_id: int, template_name: str, output_filename: str, status="Success", user_name="User"):
    """
    Records file compile/export transactions.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    now_str = datetime.datetime.now().isoformat()
    cursor.execute("""
    INSERT INTO ExportHistory (DocumentID, TemplateName, OutputFileName, ExportedBy, ExportedDate, ExportStatus)
    VALUES (?, ?, ?, ?, ?, ?)
    """, (doc_id, template_name, output_filename, user_name, now_str, status))
    conn.commit()
    conn.close()

def get_export_history() -> list:
    """
    Retrieves complete export history records.
    """
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM ExportHistory ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

# ---------------------------------------------------------------------------
# Document & Field Database Operations (Modified)
# ---------------------------------------------------------------------------

def insert_document(filename: str, status="Extracted") -> int:
    conn = get_db_connection()
    cursor = conn.cursor()
    now_str = datetime.datetime.now().isoformat()
    cursor.execute(
        "INSERT INTO document_master (filename, uploaded_at, status) VALUES (?, ?, ?)",
        (filename, now_str, status)
    )
    doc_id = cursor.lastrowid
    conn.commit()
    conn.close()
    return doc_id

def update_document_status(doc_id: int, status: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE document_master SET status = ? WHERE id = ?", (status, doc_id))
    conn.commit()
    conn.close()

def insert_extracted_fields(doc_id: int, mappings: list):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM extracted_fields WHERE document_id = ?", (doc_id,))
    for m in mappings:
        bbox_str = None
        if m.get('anchor_rect') or m.get('value_rect'):
            bbox_str = json.dumps({
                'anchor_rect': m.get('anchor_rect'),
                'value_rect': m.get('value_rect'),
                'is_manual_crop': m.get('is_manual_crop', False),
                'is_blank': m.get('is_blank', False)
            })
        cursor.execute("""
        INSERT INTO extracted_fields 
        (document_id, row_num, de_label, en_label, original_value, extracted_value, confidence, approved, page_num, bbox_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            doc_id,
            m['row_num'],
            m['de_label'],
            m['en_label'],
            m['original_value'],
            m['value'],
            m['confidence'],
            1 if m.get('approved', False) else 0,
            m.get('page', 0),
            bbox_str
        ))
    conn.commit()
    conn.close()

def get_all_documents() -> list:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM document_master ORDER BY id DESC")
    rows = cursor.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_document_fields(doc_id: int) -> list:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM extracted_fields WHERE document_id = ? ORDER BY row_num ASC", (doc_id,))
    rows = cursor.fetchall()
    conn.close()
    
    fields = []
    for r in rows:
        d = dict(r)
        bbox = {}
        if d['bbox_json']:
            try:
                bbox = json.loads(d['bbox_json'])
            except Exception:
                pass
        
        fields.append({
            'id': d['id'],
            'row_num': d['row_num'],
            'de_label': d['de_label'],
            'en_label': d['en_label'],
            'original_value': d['original_value'],
            'value': d['extracted_value'],
            'confidence': d['confidence'],
            'approved': bool(d['approved']),
            'page': d['page_num'],
            'anchor_rect': bbox.get('anchor_rect'),
            'value_rect': bbox.get('value_rect'),
            'is_manual_crop': bbox.get('is_manual_crop', False),
            'is_blank': bbox.get('is_blank', False)
        })
    return fields

def update_field_value(field_id: int, val: str, approved: bool, confidence: float):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
    UPDATE extracted_fields 
    SET extracted_value = ?, approved = ?, confidence = ?
    WHERE id = ?
    """, (val, 1 if approved else 0, confidence, field_id))
    conn.commit()
    conn.close()

# ---------------------------------------------------------------------------
# Audit Logger (CSV transaction logger)
# ---------------------------------------------------------------------------

def log_ocr_transaction(file_name: str, page_num: int, extracted_text: str, corrected_text: str, confidence_score: float):
    import csv
    file_exists = os.path.exists(LOG_FILE)
    try:
        with open(LOG_FILE, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["Timestamp", "File Name", "Page Number", "Extracted Text", "Corrected Text", "Confidence Score"])
            writer.writerow([
                datetime.datetime.now().isoformat(),
                file_name,
                page_num,
                extracted_text,
                corrected_text,
                confidence_score
            ])
    except Exception as e:
        print(f"Logging transaction failed: {e}")

# ---------------------------------------------------------------------------
# Handwriting Learning Module (Average Hash + String Similarity)
# ---------------------------------------------------------------------------

def calculate_ahash(image: Image.Image) -> str:
    # Convert to grayscale and resize to 8x8
    img = image.convert('L').resize((8, 8), Image.Resampling.LANCZOS)
    pixels = np.array(img)
    avg = pixels.mean()
    diff = pixels > avg
    hash_str = "".join(["1" if b else "0" for b in diff.flatten()])
    return f"{int(hash_str, 2):016x}"

def hamming_distance(hash1: str, hash2: str) -> int:
    h1_bin = bin(int(hash1, 16))[2:].zfill(64)
    h2_bin = bin(int(hash2, 16))[2:].zfill(64)
    return sum(c1 != c2 for c1, c2 in zip(h1_bin, h2_bin))

def save_correction(original_ocr: str, corrected_text: str, crop_img: Image.Image, anchor_label: str = None, confidence: float = 0.0):
    if not corrected_text or original_ocr.strip() == corrected_text.strip():
        return
        
    image_hash = None
    image_path = None
    if crop_img is not None:
        try:
            image_hash = calculate_ahash(crop_img)
            image_filename = f"{image_hash}.png"
            image_path = os.path.join(TRAINING_DIR, image_filename)
            crop_img.save(image_path)
        except Exception as e:
            print(f"Failed to save corrected sample: {e}")
            
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Check if duplicate correction exists
    cursor.execute("""
    SELECT id FROM handwriting_learning 
    WHERE original_ocr_text = ? AND (anchor_label = ? OR (anchor_label IS NULL AND ? IS NULL))
    """, (original_ocr, anchor_label, anchor_label))
    row = cursor.fetchone()
    
    now_str = datetime.datetime.now().isoformat()
    if row:
        cursor.execute("""
        UPDATE handwriting_learning 
        SET corrected_text = ?, image_hash = ?, image_path = ?, timestamp = ?
        WHERE id = ?
        """, (corrected_text, image_hash, image_path, now_str, row['id']))
    else:
        new_id = str(uuid.uuid4())
        cursor.execute("""
        INSERT INTO handwriting_learning (id, original_ocr_text, corrected_text, anchor_label, confidence, image_hash, image_path, timestamp)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (new_id, original_ocr, corrected_text, anchor_label, confidence, image_hash, image_path, now_str))
        
    conn.commit()
    conn.close()

def find_auto_correction(raw_text: str, crop_img: Image.Image = None, anchor_label: str = None) -> dict:
    raw_text_clean = raw_text.strip().lower()
    
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM handwriting_learning")
    learning_rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    if not learning_rows:
        return None
        
    # 1. Visual Similarity Check
    if crop_img is not None:
        try:
            curr_hash = calculate_ahash(crop_img)
            best_dist = 64
            best_match = None
            
            for row in learning_rows:
                entry_hash = row.get('image_hash')
                if entry_hash:
                    dist = hamming_distance(curr_hash, entry_hash)
                    if dist < best_dist:
                        best_dist = dist
                        best_match = row
                        
            if best_match and best_dist <= 8:
                return {
                    'corrected_text': best_match['corrected_text'],
                    'match_type': 'visual',
                    'similarity_score': round(100 * (1.0 - best_dist / 64.0)),
                    'match_reason': f"Visual handwriting match (Hamming: {best_dist})"
                }
        except Exception as e:
            print(f"Visual handwriting match failed: {e}")
            
    # 2. String Fuzzy Similarity Check
    if not raw_text_clean:
        return None
        
    best_ratio = 0.0
    best_text_match = None
    
    for row in learning_rows:
        db_ocr_clean = (row['original_ocr_text'] or "").strip().lower()
        if not db_ocr_clean:
            continue
            
        ratio = difflib.SequenceMatcher(None, raw_text_clean, db_ocr_clean).ratio()
        
        # Contextual label boost
        if anchor_label and row.get('anchor_label'):
            if anchor_label.strip().lower() == row['anchor_label'].strip().lower():
                ratio += 0.08
                
        if ratio > best_ratio:
            best_ratio = ratio
            best_text_match = row
            
    if best_text_match and best_ratio >= 0.85:
        final_score = min(1.0, best_ratio)
        return {
            'corrected_text': best_text_match['corrected_text'],
            'match_type': 'fuzzy_text',
            'similarity_score': round(final_score * 100),
            'match_reason': f"Fuzzy OCR match (similarity: {int(final_score * 100)}%)"
        }
        
    return None
