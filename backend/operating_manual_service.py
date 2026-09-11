"""
Operating Manual Service Module
Handles storage, retrieval, metadata indexing, and search for machine operating manuals.
"""

import os
import io
import json
import logging
import datetime
from typing import List, Dict, Tuple, Any, Optional

logger = logging.getLogger(__name__)

DEFAULT_MANUAL_DIR = "Operating manual"
METADATA_FILE = "manuals_meta.json"

def ensure_manual_dir(folder_path: str = DEFAULT_MANUAL_DIR) -> str:
    """Ensure the manual storage directory exists."""
    os.makedirs(folder_path, exist_ok=True)
    return folder_path

def _get_metadata_path(folder_path: str = DEFAULT_MANUAL_DIR) -> str:
    return os.path.join(folder_path, METADATA_FILE)

def load_manual_metadata(folder_path: str = DEFAULT_MANUAL_DIR) -> Dict[str, Any]:
    """Load metadata registry for manuals."""
    meta_path = _get_metadata_path(folder_path)
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Error loading manual metadata: {e}")
    return {}

def save_manual_metadata(metadata: Dict[str, Any], folder_path: str = DEFAULT_MANUAL_DIR) -> bool:
    """Save metadata registry to JSON file."""
    meta_path = _get_metadata_path(folder_path)
    try:
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        logger.error(f"Error saving manual metadata: {e}")
        return False

def list_manual_files(folder_path: str = DEFAULT_MANUAL_DIR) -> List[Dict[str, Any]]:
    """
    List all operating manuals with their metadata.
    """
    ensure_manual_dir(folder_path)
    meta_db = load_manual_metadata(folder_path)
    results = []
    
    if not os.path.exists(folder_path):
        return results

    for fname in sorted(os.listdir(folder_path)):
        if fname == METADATA_FILE or fname.startswith("."):
            continue
        
        full_path = os.path.join(folder_path, fname)
        if os.path.isfile(full_path):
            stat = os.stat(full_path)
            size_kb = round(stat.st_size / 1024, 2)
            mod_time = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
            ext = os.path.splitext(fname)[1].lower().replace(".", "").upper()
            
            # Match with metadata registry if available
            doc_meta = meta_db.get(fname, {})
            title = doc_meta.get("title", fname)
            category = doc_meta.get("category", "General")
            machine_model = doc_meta.get("machine_model", "-")
            version = doc_meta.get("version", "1.0")
            uploaded_by = doc_meta.get("uploaded_by", "System")
            
            results.append({
                "filename": fname,
                "filepath": full_path,
                "title": title,
                "category": category,
                "machine_model": machine_model,
                "version": version,
                "format": ext,
                "size_kb": size_kb,
                "modified": mod_time,
                "uploaded_by": uploaded_by,
                "display": f"{fname} ({ext}, {size_kb} KB)"
            })
            
    return results

def save_manual_document(
    file_bytes: bytes,
    filename: str,
    title: str = "",
    category: str = "General",
    machine_model: str = "",
    version: str = "1.0",
    uploaded_by: str = "Operator",
    folder_path: str = DEFAULT_MANUAL_DIR
) -> Tuple[bool, str]:
    """
    Save a new manual document and update its metadata registry.
    """
    try:
        ensure_manual_dir(folder_path)
        dest_path = os.path.join(folder_path, filename)
        
        with open(dest_path, "wb") as f:
            f.write(file_bytes)
            
        meta_db = load_manual_metadata(folder_path)
        meta_db[filename] = {
            "title": title if title.strip() else filename,
            "category": category,
            "machine_model": machine_model,
            "version": version,
            "uploaded_by": uploaded_by,
            "uploaded_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        save_manual_metadata(meta_db, folder_path)
        return True, f"Manual '{filename}' uploaded successfully."
    except Exception as e:
        logger.error(f"Failed to save manual document {filename}: {e}")
        return False, f"Failed to upload manual: {str(e)}"

def delete_manual_document(filename: str, folder_path: str = DEFAULT_MANUAL_DIR) -> Tuple[bool, str]:
    """
    Delete a manual document and remove its entry from metadata registry.
    """
    try:
        full_path = os.path.join(folder_path, filename)
        if os.path.exists(full_path):
            os.remove(full_path)
            
        meta_db = load_manual_metadata(folder_path)
        if filename in meta_db:
            del meta_db[filename]
            save_manual_metadata(meta_db, folder_path)
            
        return True, f"Manual '{filename}' deleted successfully."
    except Exception as e:
        logger.error(f"Failed to delete manual {filename}: {e}")
        return False, f"Failed to delete manual: {str(e)}"
