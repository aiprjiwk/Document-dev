import os
import json
import cv2
import numpy as np
from PIL import Image
import fitz
import logging
from backend.models import CalibrationProfile, get_db_session
from backend.alignment_service import align_document_page

logger = logging.getLogger("CalibrationMatching")

def compute_image_embedding(crop_img: Image.Image) -> np.ndarray:
    """
    Computes a normalized feature embedding vector for a character crop image.
    Uses normalized grayscale intensity distribution + Sobel gradient directional histograms.
    """
    img_gray = crop_img.convert('L').resize((32, 32), Image.Resampling.LANCZOS)
    arr = np.array(img_gray, dtype=np.float32) / 255.0
    
    # Invert so ink is 1 and background is 0
    arr = 1.0 - arr
    
    # Compute Sobel Gradients
    gx = cv2.Sobel(arr, cv2.CV_32F, 1, 0, ksize=3)
    gy = cv2.Sobel(arr, cv2.CV_32F, 0, 1, ksize=3)
    magnitude = np.sqrt(gx**2 + gy**2)
    
    # Flatten and concatenate pixel intensities + gradient magnitude
    flat_pixels = arr.flatten()
    flat_grads = magnitude.flatten()
    feature_vec = np.concatenate([flat_pixels, flat_grads])
    
    # L2 Normalize
    norm = np.linalg.norm(feature_vec)
    if norm > 0:
        feature_vec = feature_vec / norm
        
    return feature_vec

def cosine_similarity(v1: np.ndarray, v2: np.ndarray) -> float:
    """Computes cosine similarity between two 1D vectors."""
    dot = np.dot(v1, v2)
    norm1 = np.linalg.norm(v1)
    norm2 = np.linalg.norm(v2)
    if norm1 == 0 or norm2 == 0:
        return 0.0
    return float(dot / (norm1 * norm2))

def process_scanned_calibration_sheet(pdf_or_img_bytes: bytes, user_name: str) -> dict:
    """
    Parses a scanned 36-box calibration sheet PDF/image, extracts char crops,
    computes embeddings, and stores them in SQLite.
    """
    session = get_db_session()
    try:
        # Load image
        try:
            doc = fitz.open(stream=pdf_or_img_bytes, filetype="pdf")
            page = doc.load_page(0)
            pix = page.get_pixmap(dpi=300)
            img_pil = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            doc.close()
        except Exception:
            img_pil = Image.open(io.BytesIO(pdf_or_img_bytes)).convert("RGB")
            
        # Align document to standard 2100x2970 dimensions
        aligned_img = align_document_page(img_pil, target_width=2100, target_height=2970)
        
        # Grid layout matching generate_calibration_sheet_pdf
        # Relative coordinates in percentage (0 to 100)
        grid_left_pct = 10.0
        grid_top_pct = 12.0
        box_w_pct = 12.5
        box_h_pct = 12.5
        pad_x_pct = 1.2
        pad_y_pct = 1.4
        
        chars = [str(i) for i in range(10)] + [chr(i) for i in range(ord('A'), ord('Z') + 1)] # 36 chars
        
        save_dir = os.path.join("dataset_images", "calibration", user_name)
        os.makedirs(save_dir, exist_ok=True)
        
        # Clear previous embeddings for user_name
        session.query(CalibrationProfile).filter(CalibrationProfile.user_name == user_name).delete()
        
        saved_count = 0
        img_w, img_h = aligned_img.size
        
        idx = 0
        for r in range(6):
            for c in range(6):
                if idx >= len(chars):
                    break
                char_code = chars[idx]
                
                left = (grid_left_pct + c * (box_w_pct + pad_x_pct)) / 100.0 * img_w
                top = (grid_top_pct + r * (box_h_pct + pad_y_pct)) / 100.0 * img_h
                width = box_w_pct / 100.0 * img_w
                height = box_h_pct / 100.0 * img_h
                
                # Focus crop on the writing box (skip character label header)
                crop_box = (
                    int(left + width * 0.1),
                    int(top + height * 0.3),
                    int(left + width * 0.9),
                    int(top + height * 0.95)
                )
                crop_img = aligned_img.crop(crop_box)
                
                # Save crop image
                crop_file = os.path.join(save_dir, f"{char_code}.png")
                crop_img.save(crop_file)
                
                # Compute embedding
                embedding = compute_image_embedding(crop_img)
                emb_json = json.dumps(embedding.tolist())
                
                # Save to database
                profile_entry = CalibrationProfile(
                    user_name=user_name,
                    char_code=char_code,
                    image_path=crop_file,
                    feature_vector_json=emb_json
                )
                session.add(profile_entry)
                saved_count += 1
                idx += 1
                
        session.commit()
        logger.info(f"Processed calibration sheet for '{user_name}'. Saved {saved_count} character embeddings.")
        return {"status": "success", "user_name": user_name, "chars_saved": saved_count}
    except Exception as e:
        session.rollback()
        logger.error(f"Error processing calibration sheet: {e}")
        raise e
    finally:
        session.close()

def classify_with_calibration(crop_img: Image.Image, user_name: str, primary_text: str, primary_conf: float) -> tuple:
    """
    Uses Cosine Similarity comparison against user's calibration profile embeddings.
    Returns (suggested_text, adjusted_confidence, top_profile_matches)
    """
    session = get_db_session()
    try:
        profiles = session.query(CalibrationProfile).filter(CalibrationProfile.user_name == user_name).all()
        if not profiles:
            return primary_text, primary_conf, []
            
        crop_emb = compute_image_embedding(crop_img)
        
        sim_scores = []
        for p in profiles:
            vec = np.array(json.loads(p.feature_vector_json), dtype=np.float32)
            sim = cosine_similarity(crop_emb, vec)
            sim_scores.append((p.char_code, sim, p.image_path))
            
        sim_scores.sort(key=lambda x: x[1], reverse=True)
        top_char, top_score, _ = sim_scores[0]
        
        # If primary OCR confidence is low (< 0.75) or ambiguous, check calibration profile match
        if primary_conf < 0.75 or primary_text.strip() == "":
            if top_score > 0.65:
                logger.info(f"Calibration profile matched '{top_char}' with Cosine Similarity {top_score:.2f} (overriding '{primary_text}')")
                return top_char, round(top_score, 2), sim_scores[:3]
                
        return primary_text, primary_conf, sim_scores[:3]
    except Exception as e:
        logger.warning(f"Classification matching error: {e}")
        return primary_text, primary_conf, []
    finally:
        session.close()
