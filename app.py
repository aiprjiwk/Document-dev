import streamlit as st
from streamlit_option_menu import option_menu
import fitz  # PyMuPDF — used for both PDF→image conversion AND text extraction (no Tesseract needed)
from pypdf import PdfReader, PdfWriter
import pandas as pd
from PIL import Image
import re
import io
import zipfile
import os
import importlib
import datetime

# Try to import zxingcpp (recommended as it has no external DLL requirements on Windows)
try:
    import zxingcpp
    ZXING_AVAILABLE = True
except Exception:
    zxingcpp = None
    ZXING_AVAILABLE = False

# Try to import pyzbar — if DLLs are missing on Windows, disable barcode detection gracefully
try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    PYZBAR_AVAILABLE = True
except Exception:
    pyzbar_decode = None
    PYZBAR_AVAILABLE = False

# Try to import winocr for native Windows OCR fallback
try:
    from winocr import recognize_pil_sync as winocr_recognize
    WINOCR_AVAILABLE = True
except Exception:
    winocr_recognize = None
    WINOCR_AVAILABLE = False

# Try to import surya-ocr
try:
    from surya.ocr import run_ocr
    from surya.model.detection.model import load_model as load_det_model, load_processor as load_det_processor
    from surya.model.recognition.model import load_model as load_rec_model
    from surya.model.recognition.processor import load_processor as load_rec_processor
    SURYA_AVAILABLE = True
except Exception:
    SURYA_AVAILABLE = False

# Import custom backend services for Handwriting Form-to-Excel OCR
try:
    from backend.alignment_service import align_document_page
    from backend.calibration_pdf_generator import generate_calibration_sheet_pdf
    from backend.calibration_matching import process_scanned_calibration_sheet, classify_with_calibration
    from backend.excel_service import ExcelExportService
    from backend.models import OCRData, CalibrationProfile, get_db_session, init_db
    init_db()
except Exception as e:
    print(f"Warning: Could not import backend services: {e}")

# ---------------------------------------------------------------------------
# PDF → Image conversion using PyMuPDF (no Poppler or Tesseract required)
# ---------------------------------------------------------------------------
def convert_pdf_to_images(pdf_bytes, dpi=300):
    """Convert PDF bytes into a list of PIL Images using PyMuPDF."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        images = []
        zoom = dpi / 72  # 72 is the default PDF DPI
        mat = fitz.Matrix(zoom, zoom)
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(img)
        doc.close()
        return images
    except Exception as e:
        raise Exception(f"Failed to convert PDF to images using PyMuPDF. Error: {e}")

# ---------------------------------------------------------------------------
# Text extraction (no Tesseract required, uses native Windows OCR fallback)
# ---------------------------------------------------------------------------
def extract_text_pymupdf(pdf_bytes, page_index):
    """Extract embedded text from a PDF page using PyMuPDF (fast, no OCR needed)."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        page = doc.load_page(page_index)
        text = page.get_text("text")
        doc.close()
        return text
    except Exception as e:
        return ""

@st.cache_resource
def load_surya_models():
    """Load and cache Surya OCR models."""
    if not SURYA_AVAILABLE:
        return None, None, None, None
    try:
        det_processor = load_det_processor()
        det_model = load_det_model()
        rec_model = load_rec_model()
        rec_processor = load_rec_processor()
        return det_model, det_processor, rec_model, rec_processor
    except Exception as e:
        st.error(f"Failed to load Surya models: {e}")
        return None, None, None, None

def run_surya_ocr(image):
    """Run Surya OCR on an image and return the combined text."""
    det_model, det_processor, rec_model, rec_processor = load_surya_models()
    if not det_model:
        return ""
    try:
        predictions = run_ocr(
            [image], 
            ["en", "th"], 
            det_model, 
            det_processor, 
            rec_model, 
            rec_processor
        )
        text_lines = []
        for page_result in predictions:
            for text_line in page_result.text_lines:
                text_lines.append(text_line.text)
        return "\n".join(text_lines)
    except Exception as e:
        return ""

def extract_text_for_page(pdf_bytes, page_index, page_image, ocr_engine="Windows Media OCR (Native)"):
    """Extract text from the page. Try PyMuPDF first, fallback to selected OCR if empty/insufficient."""
    # 1. Try PyMuPDF (fast)
    text = extract_text_pymupdf(pdf_bytes, page_index)
    
    # 2. Check if we found PO and Material. If not, perform OCR
    has_po = extract_po_number(text) is not None
    has_material = extract_material_number(text) is not None
    
    if (not has_po or not has_material) and page_image:
        if ocr_engine == "Surya AI OCR (Deep Learning)" and SURYA_AVAILABLE:
            try:
                ocr_text = run_surya_ocr(page_image)
                if ocr_text:
                    text = text + "\n" + ocr_text
            except Exception:
                pass
        elif ocr_engine == "Windows Media OCR (Native)" and WINOCR_AVAILABLE:
            try:
                # Run native Windows OCR on the page image
                ocr_result = winocr_recognize(page_image)
                ocr_text = ocr_result.get('text', '')
                if ocr_text:
                    # Merge or use the OCR text
                    text = text + "\n" + ocr_text
            except Exception:
                pass
            
    return text

# ---------------------------------------------------------------------------
# Barcode utility - check if a barcode is a 1D linear format
# ---------------------------------------------------------------------------
def is_valid_1d_barcode(format_name):
    """Check if a barcode format is a 1D linear format to avoid matching 2D matrix/QR on contents pages."""
    allowed = {
        'CODE128', 'CODE39', 'CODE93', 'CODABAR', 
        'EAN13', 'EAN8', 'ITF', 'UPCA', 'UPCE', 
        'I25', 'ITF14'
    }
    return format_name.upper() in allowed

# ---------------------------------------------------------------------------
# Barcode detection using zxingcpp (primary) and pyzbar (fallback)
# ---------------------------------------------------------------------------
def detect_barcode(img):
    """Detect barcode in the image, prioritizing the top 35%.
    Returns the barcode value if found, otherwise None.
    """
    width, height = img.size
    top_area = img.crop((0, 0, width, int(height * 0.35)))
    
    # 1. Try modern zxingcpp first
    if ZXING_AVAILABLE:
        try:
            barcodes = zxingcpp.read_barcodes(top_area)
            if not barcodes:
                barcodes = zxingcpp.read_barcodes(img)
            # Filter for 1D linear barcodes only
            valid_barcodes = [b for b in barcodes if is_valid_1d_barcode(b.format.name)]
            if valid_barcodes:
                return valid_barcodes[0].text
        except Exception:
            pass

    # 2. Fallback to pyzbar
    if PYZBAR_AVAILABLE:
        try:
            barcodes = pyzbar_decode(top_area)
            if not barcodes:
                barcodes = pyzbar_decode(img)
            # Filter for 1D linear barcodes only
            valid_barcodes = [b for b in barcodes if is_valid_1d_barcode(b.type)]
            if valid_barcodes:
                return valid_barcodes[0].data.decode('utf-8')
        except Exception:
            pass

    return None

# ---------------------------------------------------------------------------
# RG / Quality Inspection page detection
# ---------------------------------------------------------------------------
def is_rg_page(ocr_text, barcode_value, split_mode):
    """Determine if a page is a separator page based on the selected split mode."""
    if split_mode == "Split by Barcode only (แยกด้วยบาร์โค้ดเท่านั้น)":
        return barcode_value is not None

    elif split_mode == "Split by Keywords only (แยกด้วยคำสำคัญเท่านั้น)":
        keywords = [
            "Quality Inspection",
            "PURCHASE - INFORMATION",
            "QS-INFORMATION",
            "WAREHOUSE INFORMATION",
            "Goods receipt date",
            "Inspection lot"
        ]
        text_upper = ocr_text.upper()
        for kw in keywords:
            if kw.upper() in text_upper:
                return True
        return False

    else:  # "Split by Keywords or Barcode (คำสำคัญ หรือ บาร์โค้ด)"
        keywords = [
            "Quality Inspection",
            "PURCHASE - INFORMATION",
            "QS-INFORMATION",
            "WAREHOUSE INFORMATION",
            "Goods receipt date",
            "Inspection lot"
        ]
        text_upper = ocr_text.upper()
        for kw in keywords:
            if kw.upper() in text_upper:
                return True

        # If a barcode is detected and the page contains common document header keywords
        if barcode_value and ("MATERIAL" in text_upper or "DOCUMENT" in text_upper or "INSPECTION" in text_upper):
            return True

        return False

# ---------------------------------------------------------------------------
# Material number extraction
# ---------------------------------------------------------------------------
def extract_material_number(ocr_text):
    """Extract material number using regex with smart fallbacks."""
    # Matches Material, Materail, Materia, Mat., etc. and digits
    match = re.search(r'Materi[al1e\s]*\s*[-:;.]*\s*(\d+)', ocr_text, re.IGNORECASE)
    if not match:
        # Fallback to any standalone 7-8 digit number (common Material number formats)
        match = re.search(r'\b(\d{7,8})\b', ocr_text)
    if match:
        return match.group(1)
    return None

# ---------------------------------------------------------------------------
# PO number extraction
# ---------------------------------------------------------------------------
def extract_po_number(ocr_text):
    """Extract PO number using regex with smart fallbacks."""
    # Matches PO, P.O., P0, Purchase Order, and the number
    match = re.search(r'(?:PO|P\.O\.|P0|Purchase\s*Order)\s*[-:;.]*\s*([\d/-]+)', ocr_text, re.IGNORECASE)
    if not match:
        # Fallback: PO number is typically 10 digits followed by a slash and 5 digits (e.g. 4300217445/00160)
        match = re.search(r'\b(\d{10}/\d{5})\b', ocr_text)
    if match:
        return match.group(1)
    return None

# ---------------------------------------------------------------------------
# Filename sanitization
# ---------------------------------------------------------------------------
def sanitize_filename_part(name):
    """Replace invalid Windows filename characters with underscores or dashes."""
    if not name:
        return ""
    # Replace slashes with dashes to preserve PO/barcode separation format
    name = name.replace('/', '-').replace('\\', '-')
    return re.sub(r'[?*:"<>|]', '_', name).strip()

# ---------------------------------------------------------------------------
# ZIP creation
# ---------------------------------------------------------------------------
def create_zip(split_files):
    """Create a ZIP file in memory containing all split PDFs."""
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for filename, pdf_data in split_files:
            zip_file.writestr(filename, pdf_data)
    return zip_buffer.getvalue()

# ---------------------------------------------------------------------------
# Excel export
# ---------------------------------------------------------------------------
def to_excel(df):
    """Convert DataFrame to Excel format in memory."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Log')
    return output.getvalue()

# --------------------------------------------------------
def render_home_page():
    import sqlite3
    import datetime
    
    # Page Header with greeting based on local time
    current_hour = datetime.datetime.now().hour
    if current_hour < 12:
        greeting = "🌅 Good Morning"
    elif current_hour < 18:
        greeting = "☀️ Good Afternoon"
    else:
        greeting = "🌙 Good Evening"
        
    st.title(f"{greeting}! Welcome to IWK Document Portal")
    st.subheader("Smart Document Automation & Verification Engine")
    
    # Query database for metrics
    db_path = "database/ocr_system.db"
    total_docs = 0
    total_fields = 0
    recent_docs = []
    
    try:
        if os.path.exists(db_path):
            conn = sqlite3.connect(db_path)
            cursor = conn.cursor()
            
            # Count docs
            cursor.execute("SELECT COUNT(*) FROM document_master")
            total_docs = cursor.fetchone()[0]
            
            # Count fields
            cursor.execute("SELECT COUNT(*) FROM extracted_fields")
            total_fields = cursor.fetchone()[0]
            
            # Get latest 5 docs
            cursor.execute("SELECT filename, uploaded_at, status FROM document_master ORDER BY uploaded_at DESC LIMIT 5")
            recent_docs = cursor.fetchall()
            
            conn.close()
    except Exception as db_err:
        logger.warning(f"Failed to query database for dashboard: {db_err}")
        
    # Count images in database
    pic_dir = "My Picture"
    total_images = 0
    try:
        if os.path.exists(pic_dir):
            total_images = len([f for f in os.listdir(pic_dir) if os.path.isfile(os.path.join(pic_dir, f))])
    except Exception:
        pass

    # Style definitions for modern dashboard cards
    st.markdown("""
    <style>
        .kpi-card {
            background-color: #0b1a30;
            border: 1px solid #1e3a5f;
            padding: 20px;
            border-radius: 12px;
            text-align: center;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            transition: all 0.2s ease-in-out;
        }
        .kpi-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(0, 120, 212, 0.15);
            border-color: #0078d4;
        }
        .kpi-title {
            color: #94a3b8;
            font-size: 0.9rem;
            text-transform: uppercase;
            font-weight: 600;
            margin-bottom: 8px;
        }
        .kpi-value {
            color: #ffffff;
            font-size: 2.2rem;
            font-weight: 700;
            margin-bottom: 4px;
        }
        .kpi-icon {
            font-size: 1.8rem;
            margin-bottom: 10px;
        }
    </style>
    """, unsafe_allow_html=True)

    # 📊 KPI Overview Section
    st.markdown("### 📊 Portal Insights")
    kpi_col1, kpi_col2, kpi_col3 = st.columns(3)
    
    with kpi_col1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-icon">📁</div>
            <div class="kpi-title">Documents Processed</div>
            <div class="kpi-value">{total_docs}</div>
            <div style="color: #10b981; font-size: 0.8rem;">⚡ Real-time database count</div>
        </div>
        """, unsafe_allow_html=True)
        
    with kpi_col2:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-icon">📝</div>
            <div class="kpi-title">OCR Extracted Fields</div>
            <div class="kpi-value">{total_fields}</div>
            <div style="color: #10b981; font-size: 0.8rem;">⚙️ Successfully parsed</div>
        </div>
        """, unsafe_allow_html=True)
        
    with kpi_col3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-icon">🖼️</div>
            <div class="kpi-title">Reference Database Images</div>
            <div class="kpi-value">{total_images}</div>
            <div style="color: #60a5fa; font-size: 0.8rem;">📂 Image search library</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<br>", unsafe_allow_html=True)

    # ⚙️ Quick Navigation & Workflow Section
    col_nav, col_status = st.columns([2, 1])
    
    with col_nav:
        st.markdown("### ⚙️ Quick Access Tools")
        
        shortcut_col1, shortcut_col2 = st.columns(2)
        with shortcut_col1:
            with st.container(border=True):
                st.markdown("##### 📜 Calibration Certificate")
                st.write("Format reports, split rows, insert reference images, and generate Excel print layouts (Standard/Approval).")
                st.write("**Quick Start:** Select *Document Tool Center* ➡️ *Calibration Certificate* in the sidebar.")
                
        with shortcut_col2:
            with st.container(border=True):
                st.markdown("##### 🤖 OCR & AI QC")
                st.write("Upload PDF documents, extract PO/Material barcodes, and verify OCR output before exporting to SQLite database.")
                st.write("**Quick Start:** Select *Document Tool Center* ➡️ *OCR & AI* in the sidebar.")
                
    with col_status:
        st.markdown("### 🛟 Engine Status")
        
        # Display detection engine status compactly
        if ZXING_AVAILABLE:
            st.success("🟢 **Barcode Engine (zxing)**\n\nReady (Scans 1D linear barcodes)")
        else:
            st.error("🔴 **Barcode Engine**\n\nNot found in environment")
            
        if WINOCR_AVAILABLE:
            st.success("🟢 **Windows OCR (Native)**\n\nReady (High-speed recognition)")
        else:
            st.warning("🟡 **Windows OCR**\n\nUnavailable on this system")
            
    st.markdown("---")
    
    # ⏳ Recent Activity Table
    st.markdown("### ⏳ Recently Processed Documents")
    if recent_docs:
        # Build pandas DataFrame for nice rendering
        df_recent = pd.DataFrame(recent_docs, columns=["Filename", "Uploaded At", "Status"])
        # Format Timestamp nicely
        df_recent["Uploaded At"] = pd.to_datetime(df_recent["Uploaded At"]).dt.strftime("%Y-%m-%d %H:%M:%S")
        st.dataframe(df_recent, use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ No documents processed yet. Upload a PDF using the OCR tools to get started.")

    st.markdown("---")
    
    # 🗺️ System Workflow Goals
    st.markdown("### 🗺️ System Workflow Lifecycle")
    st.markdown("""
    This application utilizes the standard document automation pipeline:
    1. **PDF Processing**: Upload files and detect separating barcodes.
    2. **OCR Engine**: Windows OCR and Surya AI parse key values and coordinates.
    3. **Human Review**: Interactive review interface for data corrections.
    4. **SQLite Storage**: Store template definitions and extracted values.
    5. **Mapping Sheet**: Dynamic lookups match item data with reference values.
    6. **Excel Export**: Generate perfectly styled Excel workbooks (standard & approval layouts).
    """)

    st.markdown("---")
    st.markdown("### 📚 Document Tool Center - Full Working Mechanisms Manual")
    st.write("สรุปกลไกการทำงานอย่างละเอียดของทุกโมดูลภายใต้ Document Tool Center (สามารถอ่านหรือดาวน์โหลดเอกสารคู่มือฉบับเต็มได้ที่นี่)")

    guide_path = os.path.join("database", "Document_Tool_Center_Mechanisms_Guide.md")
    if os.path.exists(guide_path):
        with open(guide_path, "r", encoding="utf-8") as f:
            guide_text = f.read()

        d_col1, d_col2 = st.columns(2)
        with d_col1:
            st.download_button(
                label="📥 ดาวน์โหลดคู่มือกลไกการทำงาน (.md)",
                data=guide_text.encode("utf-8"),
                file_name="Document_Tool_Center_Mechanisms_Guide.md",
                mime="text/markdown",
                use_container_width=True,
                type="primary",
                key="dashboard_download_guide_md"
            )
        with d_col2:
            st.download_button(
                label="📄 ดาวน์โหลดคู่มือกลไกการทำงาน (.txt)",
                data=guide_text.encode("utf-8"),
                file_name="Document_Tool_Center_Mechanisms_Guide.txt",
                mime="text/plain",
                use_container_width=True,
                key="dashboard_download_guide_txt"
            )

        with st.expander("📖 คลิกเพื่ออ่านสรุปกลไกการทำงานฉบับเต็มบน Dashboard", expanded=False):
            st.markdown(guide_text)

def render_ocr_certificate_page():
    st.title("📑 OCR Certificate Tool")
    st.subheader("Split PDF documents using Barcodes & OCR")

    # Sidebar Configs
    st.sidebar.header("Split Configuration")
    split_mode = st.sidebar.selectbox(
        "Split Condition",
        options=[
            "Split by Keywords or Barcode",
            "Split by Barcode only",
            "Split by Keywords only"
        ],
        index=0,
        help="Select how the app detects the start of a new document set."
    )

    # Status banners
    st.sidebar.markdown("---")
    st.sidebar.header("Engine Status")
    if ZXING_AVAILABLE:
        st.sidebar.success("✅ Barcode engine (zxing-cpp) is loaded and ready.")
    elif PYZBAR_AVAILABLE:
        st.sidebar.info("ℹ️ Barcode engine (pyzbar) is loaded.")
    else:
        st.sidebar.warning("⚠️ Barcode splitting is disabled.")

    if WINOCR_AVAILABLE:
        st.sidebar.success("✅ Native Windows OCR engine is active.")
    else:
        st.sidebar.warning("⚠️ Native Windows OCR engine is unavailable.")

    # OCR Engine defaults to Windows Media OCR (Native) in the background (hidden from sidebar UI as requested)
    ocr_engine = "Windows Media OCR (Native)"

    st.info("ℹ️ Using PyMuPDF for PDF processing — no Tesseract or Poppler required.")

    # Main UI
    uploaded_file = st.file_uploader("Upload PDF Document", type=["pdf"])

    if uploaded_file:
        if st.button("Start OCR / Split"):
            pdf_bytes = uploaded_file.read()

            progress_text = st.empty()
            progress_bar = st.progress(0)

            # Step 1: Convert PDF to images (for barcode detection)
            progress_text.text("Converting PDF to images...")
            try:
                images = convert_pdf_to_images(pdf_bytes, dpi=300)
            except Exception as e:
                st.error(str(e))
                return

            total_pages = len(images)
            document_sets = []
            current_set = None

            # Step 2, 3, 4: Extract text, detect RG pages, extract Material number
            for i, img in enumerate(images):
                progress_text.text(f"Processing page {i+1} of {total_pages}...")
                progress_bar.progress((i + 1) / total_pages)

                # Extract text (using PyMuPDF with selected OCR fallback)
                text = extract_text_for_page(pdf_bytes, i, img, ocr_engine=ocr_engine)

                # Detect barcode
                barcode = detect_barcode(img)

                if is_rg_page(text, barcode, split_mode):
                    if current_set:
                        current_set['end_page'] = i - 1
                        document_sets.append(current_set)

                    material_no = extract_material_number(text)
                    po_no = extract_po_number(text)
                    current_set = {
                        'start_page': i,
                        'material_no': material_no,
                        'po_no': po_no,
                        'barcode': barcode,
                        'raw_text': text,
                        'notes': 'New Set Detected'
                    }
                elif current_set is None:
                    current_set = {
                        'start_page': i,
                        'material_no': None,
                        'po_no': None,
                        'barcode': barcode,
                        'raw_text': text,
                        'notes': 'No start marker detected initially'
                    }

            if current_set:
                current_set['end_page'] = total_pages - 1
                document_sets.append(current_set)

            if not document_sets:
                st.error("No pages processed.")
                return

            progress_text.text("Splitting PDF...")

            # Handle file names (X_Y_Z where X=barcode, Y=PO, Z=Material) and duplicate counts
            filename_counts = {}
            for idx, doc_set in enumerate(document_sets):
                mat_no = doc_set['material_no'] or "NoMaterial"
                barcode_val = doc_set['barcode'] or "NoBarcode"
                po_val = doc_set.get('po_no') or "NoPO"

                # Sanitize each part to be safe for filenames
                x = sanitize_filename_part(barcode_val)
                y = sanitize_filename_part(po_val)
                z = sanitize_filename_part(mat_no)

                base_name = f"{x}_{y}_{z}"

                filename_counts[base_name] = filename_counts.get(base_name, 0) + 1
                doc_set['base_name'] = base_name

            base_seen = {}
            for doc_set in document_sets:
                base_name = doc_set['base_name']
                if filename_counts[base_name] > 1:
                    base_seen[base_name] = base_seen.get(base_name, 0) + 1
                    filename = f"{base_name}_Set_{base_seen[base_name]:03d}.pdf"
                else:
                    filename = f"{base_name}.pdf"
                doc_set['filename'] = filename

            # Step 5 & 6: Split original PDF using pypdf (preserves original quality)
            pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
            split_files = []

            for doc_set in document_sets:
                start = doc_set['start_page']
                end = doc_set['end_page']

                pdf_writer = PdfWriter()
                for page_num in range(start, end + 1):
                    pdf_writer.add_page(pdf_reader.pages[page_num])

                out_pdf = io.BytesIO()
                pdf_writer.write(out_pdf)
                split_files.append((doc_set['filename'], out_pdf.getvalue()))
                doc_set['status'] = 'Success'

            # -----------------------------------------------------------------
            # Save files directly to O:\...\2027\Test Folder
            # -----------------------------------------------------------------
            target_dir = r"O:\090 Documentation\030_Docu_internal\600_KPI\030_Sustainable Development( Value added)\Improvement\Doc team\2027\Test Folder"
            save_success = True
            save_error_msg = ""
            try:
                if not os.path.exists(target_dir):
                    os.makedirs(target_dir, exist_ok=True)
                
                # Save split PDFs
                for filename, pdf_data in split_files:
                    file_path = os.path.join(target_dir, filename)
                    with open(file_path, "wb") as f:
                        f.write(pdf_data)
            except Exception as e:
                save_success = False
                save_error_msg = str(e)
            # -----------------------------------------------------------------

            progress_text.text("Done! ✅")
            progress_bar.progress(1.0)

            # Step 7 & 8: Generate ZIP and Log
            zip_data = create_zip(split_files)

            df_data = []
            for idx, doc_set in enumerate(document_sets):
                df_data.append({
                    "Set No.": idx + 1,
                    "Start Page": doc_set['start_page'] + 1,
                    "End Page": doc_set['end_page'] + 1,
                    "Barcode Value": doc_set['barcode'] or "Not Found",
                    "PO No.": doc_set.get('po_no') or "Not Found",
                    "Material No.": doc_set['material_no'] or "Not Found",
                    "Output Filename": doc_set['filename'],
                    "Status": doc_set['status'],
                    "Notes": doc_set['notes']
                })
            df = pd.DataFrame(df_data)

            st.subheader("Preview Result")
            st.dataframe(df)

            # Show success/error message
            if save_success:
                st.success(f"📂 Saved all split PDFs and log files directly to:\n`{target_dir}`")
            else:
                st.error(f"❌ Failed to save files directly to target folder. Error: {save_error_msg}")

            # Download Buttons
            col1, col2, col3 = st.columns(3)
            with col1:
                st.download_button(
                    label="⬇️ Download ZIP (Split PDFs)",
                    data=zip_data,
                    file_name="split_pdfs.zip",
                    mime="application/zip"
                )
            with col2:
                csv = df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="⬇️ Download Log (CSV)",
                    data=csv,
                    file_name="split_log.csv",
                    mime="text/csv"
                )
            with col3:
                excel_data = to_excel(df)
                st.download_button(
                    label="⬇️ Download Log (Excel)",
                    data=excel_data,
                    file_name="split_log.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.markdown("---")
            st.subheader("⬇️ Download Individual Sets")
            
            st.download_button(
                label="📦 Download All as ZIP",
                data=zip_data,
                file_name="split_pdfs.zip",
                mime="application/zip",
                key="download_all_zip_section"
            )
            st.write("")
            
            # Show a grid of individual files
            for idx, doc_set in enumerate(document_sets):
                filename = doc_set['filename']
                pdf_data = split_files[idx][1] # Get PDF bytes
                start_pg = doc_set['start_page'] + 1
                end_pg = doc_set['end_page'] + 1
                barcode_val = doc_set['barcode'] or "No Barcode"
                mat_no = doc_set['material_no'] or "No Material"
                po_val = doc_set.get('po_no') or "No PO"
                
                # Render each set in an expander card
                with st.expander(f"📦 Set {idx+1}: {filename} (Pages {start_pg} - {end_pg})"):
                    col_info, col_dl = st.columns([3, 1])
                    with col_info:
                        st.markdown(f"**Barcode (X):** `{barcode_val}`")
                        st.markdown(f"**PO (Y):** `{po_val}`")
                        st.markdown(f"**Material (Z):** `{mat_no}`")
                        st.markdown(f"**Pages:** `{start_pg}` to `{end_pg}` ({end_pg - start_pg + 1} pages)")
                    with col_dl:
                        st.download_button(
                            label="⬇️ Download PDF",
                            data=pdf_data,
                            file_name=filename,
                            mime="application/pdf",
                            key=f"dl_single_{idx}"
                        )
 
            # Debug section
            st.markdown("---")
            with st.expander("🔍 Debug: View Extracted Text for Separator Pages"):
                for idx, doc_set in enumerate(document_sets):
                    start_pg = doc_set['start_page'] + 1
                    raw_text = doc_set.get('raw_text', 'No text extracted')
                    st.write(f"📂 **Set {idx+1} (Page {start_pg}):**")
                    st.code(raw_text)
 
# ---------------------------------------------------------------------------
# OCR Adjustment Change - Handwriting OCR to Excel Mapping
# ---------------------------------------------------------------------------
from PIL import ImageDraw

DEFAULT_ADJ_MAPPING = [
    # PAGE 1 (Column I, Rows 5 to 20)
    {"Field Name": "P1 - Format-no.", "Excel Cell": "I5", "PDF Page": 1, "Top (%)": 26.0, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Description", "Excel Cell": "I6", "PDF Page": 1, "Top (%)": 28.8, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Dimensions", "Excel Cell": "I7", "PDF Page": 1, "Top (%)": 31.6, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Speed", "Excel Cell": "I8", "PDF Page": 1, "Top (%)": 34.4, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Case size [mm]", "Excel Cell": "I9", "PDF Page": 1, "Top (%)": 37.2, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - product size [mm]", "Excel Cell": "I10", "PDF Page": 1, "Top (%)": 40.0, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Robot tool", "Excel Cell": "I12", "PDF Page": 1, "Top (%)": 45.6, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Suction cups", "Excel Cell": "I13", "PDF Page": 1, "Top (%)": 48.4, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Carton stacker", "Excel Cell": "I14", "PDF Page": 1, "Top (%)": 51.2, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Carton cross pusher", "Excel Cell": "I15", "PDF Page": 1, "Top (%)": 54.0, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Insertion product pusher", "Excel Cell": "I16", "PDF Page": 1, "Top (%)": 56.8, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Carton back pusher", "Excel Cell": "I17", "PDF Page": 1, "Top (%)": 59.6, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Mouth piece", "Excel Cell": "I18", "PDF Page": 1, "Top (%)": 62.4, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Case erection", "Excel Cell": "I19", "PDF Page": 1, "Top (%)": 65.2, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    {"Field Name": "P1 - Case Transport", "Excel Cell": "I20", "PDF Page": 1, "Top (%)": 68.0, "Left (%)": 70.0, "Height (%)": 2.8, "Width (%)": 25.0},
    
    # PAGE 2 (Column I, Rows 30 to 57)
    {"Field Name": "P2 - Partition mag HZ", "Excel Cell": "I30", "PDF Page": 2, "Top (%)": 19.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Partition mag VZ", "Excel Cell": "I31", "PDF Page": 2, "Top (%)": 22.3, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Width adj HZ", "Excel Cell": "I32", "PDF Page": 2, "Top (%)": 25.1, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Width adj VZ", "Excel Cell": "I33", "PDF Page": 2, "Top (%)": 27.9, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Partition mag OY", "Excel Cell": "I34", "PDF Page": 2, "Top (%)": 30.7, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail LY (6)", "Excel Cell": "I35", "PDF Page": 2, "Top (%)": 33.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail RY (6)", "Excel Cell": "I36", "PDF Page": 2, "Top (%)": 36.3, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail LY (7)", "Excel Cell": "I37", "PDF Page": 2, "Top (%)": 39.1, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail RY (7)", "Excel Cell": "I38", "PDF Page": 2, "Top (%)": 41.9, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail LY (8)", "Excel Cell": "I39", "PDF Page": 2, "Top (%)": 44.7, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail RY (8)", "Excel Cell": "I40", "PDF Page": 2, "Top (%)": 47.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail LZ (Conveyor 1)", "Excel Cell": "I44", "PDF Page": 2, "Top (%)": 54.0, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail RZ (Conveyor 1)", "Excel Cell": "I45", "PDF Page": 2, "Top (%)": 56.8, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail LZ (Conveyor 2)", "Excel Cell": "I46", "PDF Page": 2, "Top (%)": 59.6, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail RZ (Conveyor 2)", "Excel Cell": "I47", "PDF Page": 2, "Top (%)": 62.4, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail top LY", "Excel Cell": "I48", "PDF Page": 2, "Top (%)": 65.2, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail top LZ", "Excel Cell": "I49", "PDF Page": 2, "Top (%)": 68.0, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail top RY", "Excel Cell": "I50", "PDF Page": 2, "Top (%)": 70.8, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Rail top RZ", "Excel Cell": "I51", "PDF Page": 2, "Top (%)": 73.6, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Inserting pusher X", "Excel Cell": "I55", "PDF Page": 2, "Top (%)": 80.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Stacking blocker Y", "Excel Cell": "I56", "PDF Page": 2, "Top (%)": 83.3, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P2 - Stacking pusher Z", "Excel Cell": "I57", "PDF Page": 2, "Top (%)": 86.1, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    
    # PAGE 3 (Column I, Rows 61 to 90)
    {"Field Name": "P3 - Width adj Z", "Excel Cell": "I61", "PDF Page": 3, "Top (%)": 19.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Taping VY", "Excel Cell": "I62", "PDF Page": 3, "Top (%)": 22.3, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Taping HY", "Excel Cell": "I63", "PDF Page": 3, "Top (%)": 25.1, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Rail VY", "Excel Cell": "I64", "PDF Page": 3, "Top (%)": 27.9, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Rail VZ", "Excel Cell": "I65", "PDF Page": 3, "Top (%)": 30.7, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Rail HY", "Excel Cell": "I66", "PDF Page": 3, "Top (%)": 33.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Rail HZ", "Excel Cell": "I67", "PDF Page": 3, "Top (%)": 36.3, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Guide LVX", "Excel Cell": "I68", "PDF Page": 3, "Top (%)": 39.1, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Guide RVX", "Excel Cell": "I69", "PDF Page": 3, "Top (%)": 41.9, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Guide LHX", "Excel Cell": "I70", "PDF Page": 3, "Top (%)": 44.7, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Guide RHX", "Excel Cell": "I71", "PDF Page": 3, "Top (%)": 47.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Case transport OY", "Excel Cell": "I75", "PDF Page": 3, "Top (%)": 54.0, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Case finger OZ", "Excel Cell": "I76", "PDF Page": 3, "Top (%)": 56.8, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Case width UZ", "Excel Cell": "I77", "PDF Page": 3, "Top (%)": 59.6, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Dust flap OX", "Excel Cell": "I78", "PDF Page": 3, "Top (%)": 62.4, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Case erection Z", "Excel Cell": "I82", "PDF Page": 3, "Top (%)": 69.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Width adj Z (2)", "Excel Cell": "I86", "PDF Page": 3, "Top (%)": 76.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
    {"Field Name": "P3 - Width adj Z (3)", "Excel Cell": "I90", "PDF Page": 3, "Top (%)": 83.5, "Left (%)": 80.0, "Height (%)": 2.6, "Width (%)": 15.0},
]

def crop_by_percent(img, top_pct, left_pct, height_pct, width_pct):
    w, h = img.size
    x1 = int(left_pct * w / 100)
    y1 = int(top_pct * h / 100)
    x2 = int((left_pct + width_pct) * w / 100)
    y2 = int((top_pct + height_pct) * h / 100)
    return img.crop((x1, y1, x2, y2))

def draw_mapping_boxes(image, mapping_df):
    draw_img = image.copy()
    draw = ImageDraw.Draw(draw_img)
    w, h = draw_img.size
    for idx, row in mapping_df.iterrows():
        try:
            top_pct = float(row["Top (%)"])
            left_pct = float(row["Left (%)"])
            h_pct = float(row["Height (%)"])
            w_pct = float(row["Width (%)"])
            name = str(row["Field Name"])
            cell = str(row["Excel Cell"])
            
            x1 = int(left_pct * w / 100)
            y1 = int(top_pct * h / 100)
            x2 = int((left_pct + w_pct) * w / 100)
            y2 = int((top_pct + h_pct) * h / 100)
            
            # Draw red rectangle and label
            draw.rectangle([x1, y1, x2, y2], outline="red", width=3)
            draw.text((x1, max(0, y1 - 15)), f"{name} ({cell})", fill="red")
        except Exception:
            pass
    return draw_img

def run_crop_ocr(crop_img, ocr_engine="Windows Media OCR (Native)"):
    # Upscale crop_img for better handwriting OCR
    w, h = crop_img.size
    upscaled = crop_img.resize((w * 3, h * 3), Image.Resampling.LANCZOS)
    
    if ocr_engine == "Surya AI OCR (Deep Learning)" and SURYA_AVAILABLE:
        try:
            txt = run_surya_ocr(upscaled)
            return txt.strip(), 0.85
        except Exception:
            pass
            
    if WINOCR_AVAILABLE:
        try:
            res = winocr_recognize(upscaled)
            txt = res.get('text', '').strip()
            return txt, 0.90
        except Exception:
            pass
            
    # Mock/simulated value if no engine works
    return "0.0", 0.50

def run_qwen_ocr(crop_img):
    import base64
    import requests
    import io
    
    # Encode PIL Image to base64
    buffered = io.BytesIO()
    crop_img.convert('RGB').save(buffered, format="JPEG")
    base64_image = base64.b64encode(buffered.getvalue()).decode('utf-8')
    
    vllm_api_url = "http://10.52.65.63:8000/v1/chat/completions"
    headers = {
        "Content-Type": "application/json",
        "Authorization": "Bearer EMPTY"
    }
    
    payload = {
        "model": "nvidia/Qwen3.6-35B-A3B-NVFP4",
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "Can you read the text in this image and return only the text as plain text without any other words?"
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}"
                        }
                    }
                ]
            }
        ],
        "max_tokens": 1024,
        "temperature": 0.2
    }
    
    try:
        response = requests.post(vllm_api_url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        resp_json = response.json()
        extracted_text = resp_json["choices"][0]["message"]["content"].strip()
        return extracted_text, 0.95
    except Exception as e:
        return f"Error: {str(e)}", 0.0

def write_values_to_excel(template_bytes, results):
    import io
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(template_bytes))
    ws = wb.active
    
    for field_name, item in results.items():
        cell_addr = item["Excel Cell"]
        val = item["Corrected Value"]
        
        # Handle sheet prefix e.g. Sheet1!C4
        if "!" in cell_addr:
            sheet_name, cell_loc = cell_addr.split("!", 1)
            if sheet_name in wb.sheetnames:
                ws_target = wb[sheet_name]
            else:
                ws_target = ws
        else:
            cell_loc = cell_addr
            ws_target = ws
            
        # Write numeric if possible
        try:
            if val.isdigit():
                val_to_write = int(val)
            else:
                val_to_write = float(val)
        except ValueError:
            val_to_write = val
            
        # Write value safely, handling MergedCell object (which has read-only value attribute)
        try:
            ws_target[cell_loc] = val_to_write
        except AttributeError:
            # If MergedCell, find the top-left cell of the merged range
            import openpyxl.utils
            from openpyxl.utils.cell import range_boundaries
            
            target_coord = ws_target[cell_loc].coordinate
            for merged_range in list(ws_target.merged_cells.ranges):
                min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
                cell_row, cell_col = openpyxl.utils.coordinate_to_tuple(target_coord)
                if min_col <= cell_col <= max_col and min_row <= cell_row <= max_row:
                    top_left_coord = openpyxl.utils.get_column_letter(min_col) + str(min_row)
                    ws_target[top_left_coord] = val_to_write
                    break
        
    out_buf = io.BytesIO()
    wb.save(out_buf)
    return out_buf.getvalue()

def render_iwk_certificate_page():
    st.title("🏆 IWK Certificate Tool")
    
    st.subheader("Import BOM & Select Type")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        bom_file = st.file_uploader("Upload BOM (Excel)", type=["xlsx", "xls", "csv"])
        bom_format = st.radio(
            "BOM Input Format:",
            ["🔄 Original BOM (CSP2)", "📝 Manual / Transformed BOM"],
            horizontal=True,
            key="bom_format_select"
        )
        is_manual_bom = (bom_format == "📝 Manual / Transformed BOM")
        
        if is_manual_bom and bom_file:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🔄 Refresh table", use_container_width=True, help="Reload data from the uploaded Excel file"):
                st.session_state.current_bom_id = None
                st.rerun()
        
    with col2:
        cert_type = st.radio("Select Certificate Type", ["WAZ and FAD", "OZ", "SZ", "OMP"])
        
    with col3:
        uploaded_files = st.file_uploader("Upload Certificates (PDFs)", type=["pdf"], accept_multiple_files=True)
        
    if "processed_zip" not in st.session_state:
        st.session_state.processed_zip = None
    if "processed_status_bytes" not in st.session_state:
        st.session_state.processed_status_bytes = None
    if "processed_final_pdfs" not in st.session_state:
        st.session_state.processed_final_pdfs = None
    if "processed_cert_type" not in st.session_state:
        st.session_state.processed_cert_type = None
    if "missing_files" not in st.session_state:
        st.session_state.missing_files = False
        
    with col4:
        st.write("Actions & Status")
        process_btn = st.button("Process Files", type="primary", use_container_width=True)
        status_placeholder = st.empty()
        action_container = st.container()
    
    # Initialize session state for status matching sets and active bom dataframe
    if "matched_set" not in st.session_state:
        st.session_state.matched_set = None
    if "missing_set" not in st.session_state:
        st.session_state.missing_set = None
    if "active_bom_df" not in st.session_state:
        st.session_state.active_bom_df = None
    if "current_bom_id" not in st.session_state:
        st.session_state.current_bom_id = None

    # Display Excel Preview Window based on selected Certificate Type
    if bom_file:
        try:
            import sys
            import importlib
            import backend.iwk_cert_service
            importlib.reload(backend.iwk_cert_service)
            from backend.iwk_cert_service import parse_and_transform_bom, style_bom_dataframe
            
            bom_id = f"{bom_file.name}_{cert_type}_{bom_format}"
            if st.session_state.current_bom_id != bom_id:
                success, bom_records, df_preview, _ = parse_and_transform_bom(
                    bom_file.getvalue(), cert_type, is_manual=is_manual_bom
                )
                if success:
                    st.session_state.current_bom_id = bom_id
                    st.session_state.active_bom_df = df_preview
                    st.session_state.bom_records = bom_records

            df_to_use = st.session_state.get("active_bom_df")
            if df_to_use is not None and not df_to_use.empty:
                with st.expander(f"📊 Preview Processed BOM ({cert_type})", expanded=True):
                    # Check for forced mode override from Process Files click
                    current_mode = "🎨 Visual Status View (Colors & Bold)"
                    if st.session_state.get("forced_mode"):
                        current_mode = st.session_state.pop("forced_mode")

                    default_idx = 0 if current_mode == "🎨 Visual Status View (Colors & Bold)" else 1

                    display_mode = st.radio(
                        "📌 Select Display Mode:",
                        ["🎨 Visual Status View (Colors & Bold)", "✏️ Editable Table Mode"],
                        index=default_idx,
                        horizontal=True,
                        key=f"mode_select_{cert_type}"
                    )
                    if display_mode == "🎨 Visual Status View (Colors & Bold)":
                        styled_df = style_bom_dataframe(
                            df_to_use,
                            cert_type,
                            matched_files_set=st.session_state.matched_set,
                            missing_files_set=st.session_state.missing_set
                        )
                        st.dataframe(styled_df, use_container_width=True)
                    else:
                        from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode
                        
                        col_names = df_to_use.columns.tolist()
                        colA_name = col_names[0] if len(col_names) > 0 else ""
                        colC_name = col_names[2] if len(col_names) > 2 else ""
                        colJ_name = col_names[9] if len(col_names) > 9 else ""
                        colK_name = col_names[10] if len(col_names) > 10 else ""
                        colL_name = col_names[11] if len(col_names) > 11 else ""
                        
                        matched_list = list(st.session_state.matched_set) if st.session_state.matched_set else []
                        missing_list = list(st.session_state.missing_set) if st.session_state.missing_set else []
                        
                        jscode = f"""
                        function(params) {{
                            var styles = {{}};
                            
                            var colA = params.data['{colA_name}'] ? String(params.data['{colA_name}']).trim() : "";
                            if (colA.endsWith(".0")) colA = colA.slice(0, -2);
                            
                            if (colA === "2") {{
                                styles['font-weight'] = 'bold';
                            }}
                            
                            var is_red = false;
                            var cert_type = '{cert_type}';
                            
                            var valJ = params.data['{colJ_name}'] ? String(params.data['{colJ_name}']).trim() : "";
                            var valK = params.data['{colK_name}'] ? String(params.data['{colK_name}']).trim() : "";
                            var valL = params.data['{colL_name}'] ? String(params.data['{colL_name}']).trim() : "";
                            
                            if (cert_type === 'WAZ and FAD') {{
                                if (valL === "" || valL === "None") is_red = true;
                            }} else if (cert_type === 'OZ') {{
                                if (valJ === "" || valJ === "None" || valK === "" || valK === "None" || valL === "" || valL === "None") is_red = true;
                            }} else if (cert_type === 'SZ' || cert_type === 'OMP') {{
                                if (valK === "" || valK === "None" || valL === "" || valL === "None") is_red = true;
                            }}
                            
                            if (is_red) {{
                                styles['background-color'] = '#ff4d4d';
                                styles['color'] = 'white';
                            }}
                            
                            return styles;
                        }}
                        """
                        
                        cell_jscode = f"""
                        function(params) {{
                            var styles = {{}};
                            var val = params.value ? String(params.value).trim() : "";
                            if (val.endsWith(".0")) val = val.slice(0, -2);
                            var val_clean = val.toLowerCase();
                            
                            var matched = {matched_list};
                            var missing = {missing_list};
                            
                            if (matched.includes(val_clean)) {{
                                styles['background-color'] = '#4caf50';
                                styles['color'] = 'white';
                                styles['font-weight'] = 'bold';
                            }} else if (missing.includes(val_clean)) {{
                                styles['background-color'] = '#ffca28';
                                styles['color'] = 'black';
                                styles['font-weight'] = 'bold';
                            }}
                            return styles;
                        }}
                        """
                        
                        df_to_aggrid = df_to_use.copy()
                        df_to_aggrid['_row_id'] = range(len(df_to_aggrid))
                        
                        # --- Custom Toolbar ---
                        st.markdown("<br>", unsafe_allow_html=True)
                        tb_col1, tb_col2, tb_col3 = st.columns([6, 3, 2])
                        with tb_col2:
                            search_text = st.text_input("Search", key=f"tb_search_{cert_type}", label_visibility="collapsed", placeholder="🔍 Search table...")
                        with tb_col3:
                            csv_data = df_to_use.to_csv(index=False).encode('utf-8')
                            st.download_button(label="📥 Download CSV", data=csv_data, file_name="edited_bom.csv", mime="text/csv", key=f"tb_dl_{cert_type}", use_container_width=True)
                        # ----------------------
                        
                        gb = GridOptionsBuilder.from_dataframe(df_to_aggrid)
                        gb.configure_default_column(editable=True, sortable=False, filter=True)
                        gb.configure_column('_row_id', hide=True)
                        if colC_name:
                            gb.configure_column(colC_name, cellStyle=JsCode(cell_jscode))
                        gb.configure_grid_options(
                            getRowStyle=JsCode(jscode),
                            quickFilterText=search_text
                        )
                        gb.configure_selection('multiple', use_checkbox=True)
                        grid_options = gb.build()
                        
                        grid_response = AgGrid(
                            df_to_aggrid,
                            gridOptions=grid_options,
                            update_mode=GridUpdateMode.MODEL_CHANGED,
                            allow_unsafe_jscode=True,
                            theme="streamlit",
                            key=f"bom_aggrid_{cert_type}"
                        )
                        
                        updated_df = pd.DataFrame(grid_response['data'])
                        if '_row_id' in updated_df.columns:
                            updated_df = updated_df.drop(columns=['_row_id'])
                        
                        if st.session_state.active_bom_df is None or not updated_df.equals(st.session_state.active_bom_df):
                            st.session_state.active_bom_df = updated_df
                            st.rerun()
                        
                        st.markdown("**✏️ Table Actions (Add / Remove)**")
                        col_btn1, col_btn2, col_btn3 = st.columns(3)
                        
                        selected_rows = grid_response.get("selected_rows", [])
                        sel_ids = []
                        if isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:
                            if '_row_id' in selected_rows.columns:
                                sel_ids = selected_rows['_row_id'].tolist()
                        elif isinstance(selected_rows, list) and selected_rows:
                            sel_ids = [s['_row_id'] for s in selected_rows if '_row_id' in s]
                            
                        with col_btn1:
                            if st.button("⬆️ Insert Row Above", use_container_width=True, key=f"add_row_above_{cert_type}"):
                                new_row = pd.DataFrame([[None]*len(updated_df.columns)], columns=updated_df.columns)
                                if sel_ids:
                                    df_from_grid = pd.DataFrame(grid_response['data'])
                                    if '_row_id' in df_from_grid.columns:
                                        target_idx = df_from_grid[df_from_grid['_row_id'] == sel_ids[0]].index[0]
                                        df1 = updated_df.iloc[:target_idx]
                                        df2 = updated_df.iloc[target_idx:]
                                        st.session_state.active_bom_df = pd.concat([df1, new_row, df2], ignore_index=True)
                                        st.rerun()
                                else:
                                    st.session_state.active_bom_df = pd.concat([new_row, updated_df], ignore_index=True)
                                    st.rerun()
                        with col_btn2:
                            if st.button("⬇️ Insert Row Below", use_container_width=True, key=f"add_row_below_{cert_type}"):
                                new_row = pd.DataFrame([[None]*len(updated_df.columns)], columns=updated_df.columns)
                                if sel_ids:
                                    df_from_grid = pd.DataFrame(grid_response['data'])
                                    if '_row_id' in df_from_grid.columns:
                                        target_idx = df_from_grid[df_from_grid['_row_id'] == sel_ids[-1]].index[0]
                                        df1 = updated_df.iloc[:target_idx+1]
                                        df2 = updated_df.iloc[target_idx+1:]
                                        st.session_state.active_bom_df = pd.concat([df1, new_row, df2], ignore_index=True)
                                        st.rerun()
                                else:
                                    st.session_state.active_bom_df = pd.concat([updated_df, new_row], ignore_index=True)
                                    st.rerun()
                        with col_btn3:
                            if st.button("🗑️ Delete Selected Rows", use_container_width=True, key=f"del_row_{cert_type}"):
                                if sel_ids:
                                    df_filtered = pd.DataFrame(grid_response['data'])
                                    df_filtered = df_filtered[~df_filtered['_row_id'].isin(sel_ids)]
                                    if '_row_id' in df_filtered.columns:
                                        df_filtered = df_filtered.drop(columns=['_row_id'])
                                    st.session_state.active_bom_df = df_filtered.reset_index(drop=True)
                                    st.rerun()
                                else:
                                    st.warning("⚠️ Please select rows using checkboxes first.")
        except Exception as preview_err:
            st.warning(f"Unable to preview BOM: {str(preview_err)}")
            
    st.markdown("---")
    
    # 4 & 5. Add Text and Merge logic is triggered here
    if process_btn:
        if not bom_file:
            status_placeholder.error("Please upload BOM.")
            return
        if not uploaded_files:
            status_placeholder.error("Please upload PDFs.")
            return
            
        with st.spinner("Processing files..."):
            try:
                import sys
                import importlib
                import backend.iwk_cert_service
                importlib.reload(backend.iwk_cert_service)
                from backend.iwk_cert_service import parse_and_transform_bom, process_iwk_certificates, separate_and_match_files
                
                success, raw_bom_records, df_preview, excel_status_bytes = parse_and_transform_bom(
                    bom_file.getvalue(), cert_type, is_manual=is_manual_bom
                )
                if not success:
                    status_placeholder.error("BOM Error")
                    return
                    
                bom_records = raw_bom_records
                if st.session_state.get("active_bom_df") is not None and not st.session_state.active_bom_df.empty:
                    active_df = st.session_state.active_bom_df
                    bom_records = []
                    current_group = "Unknown_Group"
                    for idx, row in active_df.iterrows():
                        colA = str(row.iloc[0]).strip() if len(row) > 0 and pd.notna(row.iloc[0]) else ""
                        if colA.endswith(".0"): colA = colA[:-2]
                        colC = str(row.iloc[2]).strip() if len(row) > 2 and pd.notna(row.iloc[2]) else ""
                        if colC.endswith(".0"): colC = colC[:-2]
                        colM = str(row.iloc[12]).strip() if len(row) > 12 and pd.notna(row.iloc[12]) else ""
                        if colA == "2":
                            current_group = colM if colM else f"Group_{idx}"
                            bom_records.append({"File Name": "", "Group": current_group})
                        else:
                            if colC:
                                bom_records.append({"File Name": f"{colC}.pdf", "Group": current_group})
                    
                zip_bytes, multi_sheet_excel_bytes, final_pdfs = process_iwk_certificates(
                    uploaded_files, bom_records, cert_type, bom_file_bytes=bom_file.getvalue(), is_manual=is_manual_bom
                )
                
                # Update session state sets for live preview highlighting
                _, matched_files, missing_files, _ = separate_and_match_files(uploaded_files, bom_records)
                st.session_state.matched_set = {os.path.splitext(f)[0].lower() for f in matched_files.keys()}
                st.session_state.missing_set = {str(item.get("File Name", "")).replace(".pdf", "").lower() for item in missing_files}
                
                # Automatically switch to Visual Status View so colors show up immediately without modifying widget key
                st.session_state.forced_mode = "🎨 Visual Status View (Colors & Bold)"
                
                st.session_state.processed_zip = zip_bytes
                st.session_state.processed_status_bytes = multi_sheet_excel_bytes
                st.session_state.processed_final_pdfs = final_pdfs
                st.session_state.processed_cert_type = cert_type
                st.session_state.missing_files = False
                
                if final_pdfs:
                    import fitz
                    merged_all = fitz.open()
                    for pdf_name, pdf_bytes in final_pdfs.items():
                        try:
                            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
                            merged_all.insert_pdf(doc)
                        except Exception:
                            pass
                    st.session_state.processed_merged_pdf = merged_all.write()
                    merged_all.close()
                else:
                    st.session_state.processed_merged_pdf = None
                
                st.rerun()
                
            except Exception as e:
                status_placeholder.error(f"Error: {str(e)}")
                return
                
    if st.session_state.processed_zip:
        if len(st.session_state.processed_zip) < 100:
            status_placeholder.error("⚠️ ไม่พบไฟล์ PDF ที่ตรงกับ Part No. ใน BOM เลยครับ ทำให้ไม่มีไฟล์ใน ZIP")
        elif st.session_state.missing_files:
            status_placeholder.warning("⚠️ Done (some missing)")
        else:
            status_placeholder.success("✅ Complete!")
            
        with action_container:
            if len(st.session_state.processed_zip) >= 100:
                st.download_button(
                    label="⬇️ Download (ZIP)",
                    data=st.session_state.processed_zip,
                    file_name=f"IWK_{st.session_state.processed_cert_type}_Output.zip",
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )
                if st.session_state.get("processed_merged_pdf"):
                    st.download_button(
                        label="🖨️ Download Merged PDF",
                        data=st.session_state.processed_merged_pdf,
                        file_name=f"IWK_{st.session_state.processed_cert_type}_Merged.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
            
            if st.session_state.processed_status_bytes:
                st.download_button(
                    label="⬇️ Download Excel Status",
                    data=st.session_state.processed_status_bytes,
                    file_name=f"Status_{st.session_state.processed_cert_type}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="secondary",
                    use_container_width=True
                )
            
            if st.session_state.processed_cert_type == "WAZ and FAD":
                if st.button("💾 Save to Database", use_container_width=True):
                    target_dir = r"O:\090 Documentation\010 Templates Tech Doc\010_Templates_Documents\020 Plans+Add. Documents\060 Protocols-Certificates\00 Database Material certificate\Material certificate"
                    
                    try:
                        os.makedirs(target_dir, exist_ok=True)
                        saved_count = 0
                        
                        for fname, f_bytes in st.session_state.processed_final_pdfs.items():
                            base_name, ext = os.path.splitext(fname)
                            out_path = os.path.join(target_dir, fname)
                            counter = 1
                            
                            while os.path.exists(out_path):
                                out_path = os.path.join(target_dir, f"{base_name} ({counter}){ext}")
                                counter += 1
                                
                            with open(out_path, "wb") as f:
                                f.write(f_bytes)
                            saved_count += 1
                            
                        st.success(f"Saved {saved_count} PDFs to Database!")
                    except Exception as db_e:
                        st.error(f"Failed to save to database: {str(db_e)}")

def run_ollama_ocr(img, model_name="glm-ocr"):
    import base64
    import io
    import requests
    
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode("utf-8")
    
    url = "http://localhost:11434/api/chat"
    payload = {
        "model": model_name,
        "messages": [
            {
                "role": "user",
                "content": "Identify the columns 'Including', 'drawings', 'RevFix', 'Vs', 'Stat' in the image. Return all table rows exactly, with columns separated by two or more spaces.",
                "images": [img_str]
            }
        ],
        "stream": False
    }
    try:
        response = requests.post(url, json=payload, timeout=45)
        if response.status_code == 200:
            content = response.json().get("message", {}).get("content", "")
            return content
    except Exception:
        pass
    return ""

def extract_drawing_status_from_pdf(pdf_bytes, ocr_engine="Windows Media OCR (Native)", ollama_model="glm-ocr"):
    import fitz
    import pandas as pd
    import re
    
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    machine_name = "Unknown"
    all_table_rows = []
    
    # 1. Find Machine Name by collapsing text and stopping before BOM
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        collapsed_text = re.sub(r'\s+', ' ', text)
        
        m = re.search(r'(?i)m\s*a\s*c\s*h\s*i\s*n\s*e\s*:\s*(.*?)(?:\s+b\s*o\s*m\s*:|$)', collapsed_text)
        if m:
            val = m.group(1).strip()
            # Clean up spacing if it looks like spaced characters (e.g. "T / 0 5 6 0 2 1 X")
            if len(re.findall(r'\s', val)) > len(val) / 3:
                val = re.sub(r'\s+', '', val)
            else:
                val = val.strip()
            # Final cleanup of any trailing labels or garbage
            val = re.sub(r'(?i)\b(bom|status|date|page).*', '', val).strip()
            machine_name = val
            break
            
    # 2. Extract Tables using coordinate-based word clustering & binning
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        page_width = page.rect.width
        
        words = page.get_text("words")
        if not words:
            continue
            
        # Group words into lines by clustering y0 coordinates (within 8 pixels)
        sorted_by_y = sorted(words, key=lambda w: w[1])
        lines = []
        current_line = []
        current_y = None
        
        for w in sorted_by_y:
            y0 = w[1]
            if current_y is None:
                current_y = y0
                current_line.append(w)
            elif abs(y0 - current_y) < 8:
                current_line.append(w)
            else:
                lines.append(current_line)
                current_line = [w]
                current_y = y0
        if current_line:
            lines.append(current_line)
            
        for line_words in lines:
            # Sort words of the line by x0 coordinate
            line_words = sorted(line_words, key=lambda w: w[0])
            
            # Merge characters that are very close to each other (handles spaced OCR and vector fonts)
            merged_words = []
            for w in line_words:
                if not merged_words:
                    merged_words.append(list(w))
                else:
                    last = merged_words[-1]
                    gap = w[0] - last[2]
                    if gap < 5:
                        # Close characters: merge directly
                        last[4] += w[4]
                        last[2] = w[2]
                    elif gap < 14:
                        # Word spacing: merge with a space
                        last[4] += " " + w[4]
                        last[2] = w[2]
                    else:
                        # Column gap: keep as separate word
                        merged_words.append(list(w))
                        
            # Bin words into columns based on relative x0 coordinate
            cols = {
                "including": [],
                "drawings": [],
                "revfix": [],
                "vs": [],
                "stat": []
            }
            
            for w in merged_words:
                x0, y0, x1, y1, text = w[:5]
                rel_x = x0 / page_width
                
                if rel_x < 0.12:
                    cols["including"].append(text)
                elif 0.12 <= rel_x < 0.42:
                    cols["drawings"].append(text)
                elif 0.58 <= rel_x < 0.66:
                    cols["revfix"].append(text)
                elif 0.79 <= rel_x < 0.84:
                    cols["vs"].append(text)
                elif 0.84 <= rel_x < 0.92:
                    cols["stat"].append(text)
                    
            row_dict = {
                "Including": "".join(cols["including"]).strip(),
                "drawings": " ".join(cols["drawings"]).strip(),
                "RevFix": "".join(cols["revfix"]).strip(),
                "Vs": "".join(cols["vs"]).strip(),
                "Stat": "".join(cols["stat"]).strip()
            }
            
            # Normalize column spacing
            row_dict["Including"] = re.sub(r'\s+', '', row_dict["Including"])
            row_dict["RevFix"] = re.sub(r'\s+', '', row_dict["RevFix"])
            row_dict["Vs"] = re.sub(r'\s+', '', row_dict["Vs"])
            row_dict["Stat"] = re.sub(r'\s+', '', row_dict["Stat"])
            
            # Validation: Include row if "Including" is a number and "Stat" is not empty
            if row_dict["Including"].isdigit() and len(row_dict["Including"]) >= 5 and row_dict["Stat"]:
                all_table_rows.append(row_dict)
                
    # 3. Fallback for scanned PDF (OCR based)
    total_text = ""
    for page_num in range(len(doc)):
        total_text += doc.load_page(page_num).get_text("text")
        
    if len(total_text.strip()) < 50 or not all_table_rows:
        try:
            images = convert_pdf_to_images(pdf_bytes, dpi=150)
            for img in images:
                ocr_text = ""
                if ocr_engine == "GLM-OCR (Ollama Local)":
                    ocr_text = run_ollama_ocr(img, ollama_model)
                elif ocr_engine == "Windows Media OCR (Native)" and WINOCR_AVAILABLE:
                    try:
                        res = winocr_recognize(img)
                        ocr_text = res.get('text', '')
                    except Exception:
                        pass
                elif ocr_engine == "Surya AI OCR (Deep Learning)" and SURYA_AVAILABLE:
                    try:
                        ocr_text = run_surya_ocr(img)
                    except Exception:
                        pass
                        
                if ocr_text:
                    lines = ocr_text.split('\n')
                    for line in lines:
                        parts = [p.strip() for p in re.split(r'\s{2,}', line.strip()) if p.strip()]
                        if len(parts) >= 5:
                            row_dict = {
                                "Including": re.sub(r'\s+', '', parts[0]),
                                "drawings": parts[1],
                                "RevFix": re.sub(r'\s+', '', parts[2]),
                                "Vs": re.sub(r'\s+', '', parts[3]),
                                "Stat": re.sub(r'\s+', '', parts[4])
                            }
                            if row_dict["Including"].isdigit() and len(row_dict["Including"]) >= 5 and row_dict["Stat"]:
                                all_table_rows.append(row_dict)
                    if machine_name == "Unknown":
                        m = re.search(r'(?i)(?:Machine|M\s*a\s*c\s*h\s*i\s*n\s*e)\s*:\s*([A-Za-z0-9_\-\.\/\\\s]+)', ocr_text)
                        if not m:
                            m = re.search(r'(?i)(?:Machine|M\s*a\s*c\s*h\s*i\s*n\s*e)\s*:\s*([^\n\r]+)', ocr_text)
                        if m:
                            machine_name = m.group(1).strip()
        except Exception:
            pass
            
    # Final cleanup and header removal
    cleaned_rows = []
    for r in all_table_rows:
        inc = str(r["Including"]).strip().lower()
        drw = str(r["drawings"]).strip().lower()
        if "including" in inc or "drawings" in drw or "revfix" in inc or "vs" in inc or "stat" in inc:
            continue
        if not any(r.values()):
            continue
        cleaned_rows.append(r)
        
    doc.close()
    return machine_name, cleaned_rows

def extract_drawing_status_from_excel(excel_bytes):
    import io
    import openpyxl
    import re
    
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws = wb.active
    
    machine_name = "Unknown"
    all_table_rows = []
    
    # 1. Fetch Machine Name from B4 (Row 4, Column 2) as primary, fallback to scanning
    machine_val_b4 = ws.cell(row=4, column=2).value
    if machine_val_b4:
        machine_name = str(machine_val_b4).strip()
    else:
        for r_idx in range(1, min(ws.max_row + 1, 100)):
            row_cells = [ws.cell(row=r_idx, column=c_idx).value for c_idx in range(1, min(ws.max_column + 1, 50))]
            for cell_val in row_cells:
                if cell_val and isinstance(cell_val, str):
                    m = re.search(r'(?i)machine\s*:\s*(.*)', cell_val)
                    if m:
                        machine_name = m.group(1).strip()
                        break
                    elif cell_val.strip().lower() == "machine":
                        cell_idx = row_cells.index(cell_val)
                        if cell_idx + 1 < len(row_cells) and row_cells[cell_idx + 1]:
                            machine_name = str(row_cells[cell_idx + 1]).strip()
                            break
            if machine_name != "Unknown":
                break
            
    # 2. Find header row and column mappings (case-insensitive)
    header_row_idx = None
    col_mapping = {}
    
    for r_idx in range(1, min(ws.max_row + 1, 200)):
        row_cells = [ws.cell(row=r_idx, column=c_idx).value for c_idx in range(1, min(ws.max_column + 1, 50))]
        
        including_col = None
        for c_idx, val in enumerate(row_cells, start=1):
            if val and isinstance(val, str) and "including" in val.lower():
                including_col = c_idx
                break
                
        if including_col is not None:
            header_row_idx = r_idx
            col_mapping["including"] = including_col
            # The next column is drawing description
            col_mapping["drawings"] = including_col + 1
            
            for c_idx, val in enumerate(row_cells, start=1):
                if val and isinstance(val, str):
                    val_clean = val.strip().lower()
                    if val_clean == "revfix":
                        col_mapping["revfix"] = c_idx
                    elif val_clean == "vs":
                        col_mapping["vs"] = c_idx
                    elif val_clean == "stat":
                        col_mapping["stat"] = c_idx + 1  # Actual status values are in Column R (column Q + 1)
            break
            
    # 3. Fallback standard columns if header row not found
    if header_row_idx is None:
        col_mapping = {
            "including": 1,
            "drawings": 2,
            "revfix": 9,
            "vs": 15,
            "stat": 18  # Column R
        }
        header_row_idx = 33
        
    # Ensure standard fallbacks
    if "including" not in col_mapping: col_mapping["including"] = 1
    if "drawings" not in col_mapping: col_mapping["drawings"] = 2
    if "revfix" not in col_mapping: col_mapping["revfix"] = 9
    if "vs" not in col_mapping: col_mapping["vs"] = 15
    if "stat" not in col_mapping: col_mapping["stat"] = 18

    # 4. Extract rows
    start_row = header_row_idx + 1
    for r_idx in range(start_row, ws.max_row + 1):
        inc_val = ws.cell(row=r_idx, column=col_mapping["including"]).value
        dwg_val = ws.cell(row=r_idx, column=col_mapping["drawings"]).value
        rev_val = ws.cell(row=r_idx, column=col_mapping["revfix"]).value
        vs_val = ws.cell(row=r_idx, column=col_mapping["vs"]).value
        stat_val = ws.cell(row=r_idx, column=col_mapping["stat"]).value
        
        inc_str = str(inc_val).strip() if inc_val is not None else ""
        dwg_str = str(dwg_val).strip() if dwg_val is not None else ""
        rev_str = str(rev_val).strip() if rev_val is not None else ""
        vs_str = str(vs_val).strip() if vs_val is not None else ""
        stat_str = str(stat_val).strip() if stat_val is not None else ""
        
        inc_clean = re.sub(r'\s+', '', inc_str)
        rev_clean = re.sub(r'\s+', '', rev_str)
        vs_clean = re.sub(r'\s+', '', vs_str)
        stat_clean = re.sub(r'\s+', '', stat_str)
        
        if inc_clean.isdigit() and len(inc_clean) >= 5:
            row_dict = {
                "Including": inc_clean,
                "drawings": dwg_str,
                "RevFix": rev_clean,
                "Vs": vs_clean,
                "Stat": stat_clean
            }
            all_table_rows.append(row_dict)
            
    if machine_name == "Unknown":
        if ws.title != "Sheet" and ws.title != "Sheet1":
            machine_name = ws.title
            
    return machine_name, all_table_rows

def generate_dwg_verification_excel(original_rows, filtered_rows, machine_name):
    import io
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
    bold_font = Font(bold=True)
    yellow_fill = PatternFill(start_color="FFE599", end_color="FFE599", fill_type="solid")
    orange_fill = PatternFill(start_color="F9CB9C", end_color="F9CB9C", fill_type="solid")
    
    headers = ["Including", "drawings", "RevFix", "Vs", "Stat"]
    
    # Sheet 1: Original
    ws_orig = wb.create_sheet(title="Original")
    ws_orig.cell(row=1, column=1, value=f"Machine: {machine_name}").font = bold_font
    ws_orig.append([])
    ws_orig.append(headers)
    ws_orig.row_dimensions[3].font = bold_font
    
    for r in original_rows:
        ws_orig.append([r.get("Including", ""), r.get("drawings", ""), r.get("RevFix", ""), r.get("Vs", ""), r.get("Stat", "")])
        vs = str(r.get("Vs", "")).strip()
        stat = str(r.get("Stat", "")).strip().upper()
        
        last_row = ws_orig.max_row
        if vs == "" or vs.lower() == "none" or vs.lower() == "nan":
            for col_idx in range(1, 6):
                ws_orig.cell(row=last_row, column=col_idx).fill = orange_fill
        elif stat != "FR":
            for col_idx in range(1, 6):
                ws_orig.cell(row=last_row, column=col_idx).fill = yellow_fill
                
    # Sheet 2: Filtered
    ws_filt = wb.create_sheet(title="Filtered")
    ws_filt.cell(row=1, column=1, value=f"Machine: {machine_name}").font = bold_font
    ws_filt.append([])
    ws_filt.append(headers)
    ws_filt.row_dimensions[3].font = bold_font
    
    for r in filtered_rows:
        ws_filt.append([r.get("Including", ""), r.get("drawings", ""), r.get("RevFix", ""), r.get("Vs", ""), r.get("Stat", "")])
        vs = str(r.get("Vs", "")).strip()
        stat = str(r.get("Stat", "")).strip().upper()
        
        last_row = ws_filt.max_row
        if vs == "" or vs.lower() == "none" or vs.lower() == "nan":
            for col_idx in range(1, 6):
                ws_filt.cell(row=last_row, column=col_idx).fill = orange_fill
        elif stat != "FR":
            for col_idx in range(1, 6):
                ws_filt.cell(row=last_row, column=col_idx).fill = yellow_fill
                
    # Auto-fit columns
    for ws in [ws_orig, ws_filt]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def render_etk_verification_page():
    st.title("🔍 ETK Verification Tool")
    st.markdown("Automate PROJECT Plan vs CSP2 / CSPB BOM comparison and report generation.")
    
    # Initialize session state for PDF extraction
    if "pdf_det_completed" not in st.session_state:
        st.session_state.pdf_det_completed = False
    if "pdf_det_raw_rows" not in st.session_state:
        st.session_state.pdf_det_raw_rows = []
    if "pdf_det_filtered_rows" not in st.session_state:
        st.session_state.pdf_det_filtered_rows = []
    if "pdf_det_machine_name" not in st.session_state:
        st.session_state.pdf_det_machine_name = "Unknown"
    if "pdf_det_excel_bytes" not in st.session_state:
        st.session_state.pdf_det_excel_bytes = None
    if "pdf_det_filename" not in st.session_state:
        st.session_state.pdf_det_filename = ""
        
    with st.container(border=True):
        st.markdown("### 📥 Select Comparison Mode & Upload Files")
        
        # Select comparison type horizontally
        etk_mode = st.radio(
            "Select BOM Comparison Type",
            ["CSP2 BOM", "CSPB BOM"],
            key="etk_mode_select",
            horizontal=True
        )
        
        # 2-column layout for symmetric file upload boxes
        col1, col2 = st.columns(2)
        with col1:
            project_file = st.file_uploader("Insert #1: Select PROJECT file", type=["xlsx", "xlsm", "xls"], key="etk_project_file")
        with col2:
            bom_file = st.file_uploader(f"Insert #2: Select {etk_mode} file", type=["xlsx", "xlsm", "xls"], key="etk_bom_file")
            
        # Centered action button
        col_btn1, col_btn2, col_btn3 = st.columns([1, 2, 1])
        with col_btn2:
            compare_btn = st.button("🚀 Compare & Verify ETK", type="primary", use_container_width=True)
    
    if compare_btn:
        if not project_file:
            st.error("Please upload Insert #1 (PROJECT file).")
            return
        if not bom_file:
            st.error(f"Please upload Insert #2 ({etk_mode} file).")
            return
            
        with st.spinner("Comparing files and generating ETK verification report..."):
            try:
                import importlib
                import backend.etk_verification_service
                importlib.reload(backend.etk_verification_service)
                from backend.etk_verification_service import compare_etk_verification
                
                success, stats_or_msg, preview_df, excel_bytes = compare_etk_verification(
                    project_file.getvalue(), bom_file.getvalue(), mode=etk_mode
                )
                
                if not success:
                    st.error(f"ETK Verification Failed: {stats_or_msg}")
                    return
                    
                st.session_state.etk_stats = stats_or_msg
                st.session_state.etk_preview_df = preview_df
                st.session_state.etk_excel_bytes = excel_bytes
                st.session_state.etk_mode_run = etk_mode
                
            except Exception as e:
                st.error(f"Error during ETK Verification: {str(e)}")
 
    if "etk_preview_df" in st.session_state and st.session_state.etk_preview_df is not None:
        stats = st.session_state.get("etk_stats", {})
        st.success("✅ ETK Verification Completed Successfully!")
        
        # Display KPI metrics
        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Project Plan Groups", stats.get("Project Rows", 0))
        m2.metric(f"{stats.get('Mode', 'BOM')} Groups", stats.get("BOM Rows", 0))
        m3.metric("Matched Groups", stats.get("Matched Rows", 0))
        m4.metric("DIFF: Project Only", stats.get("Diff Project Only", 0))
        m5.metric(f"DIFF: {stats.get('Mode', 'BOM').split()[0]} Only", stats.get("Diff BOM Only", 0))
        
        st.markdown("### 📊 Live Comparison Preview")
        
        def highlight_status(row):
            if row.get("Status") == "Matched":
                return ['background-color: #C6EFCE; color: #006100; font-weight: bold;'] * len(row)
            else:
                return ['background-color: #FFC9CE; color: #9C0006;'] * len(row)
                
        styled_df = st.session_state.etk_preview_df.style.apply(highlight_status, axis=1)
        st.dataframe(styled_df, use_container_width=True)
        
        st.download_button(
            label=f"⬇️ Download ETK Comparison Excel Report ({st.session_state.get('etk_mode_run')})",
            data=st.session_state.etk_excel_bytes,
            file_name=f"Project_vs_{st.session_state.get('etk_mode_run').replace(' ', '_')}_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # ---------------------------------------------------------------------------
    # Drawing Status Detection Section
    # ---------------------------------------------------------------------------
    st.markdown("---")
    
    with st.container(border=True):
        st.markdown("### 📊 Drawing Status 'FR' Detection")
        
        pdf_file = st.file_uploader(
            "Upload Drawing Status Excel file (ZETK2)",
            type=["xlsx", "xlsm", "xls"],
            key="etk_pdf_det_file"
        )
        
        if not pdf_file:
            # Reset detection state when file is removed
            st.session_state.pdf_det_completed = False
            st.session_state.pdf_det_raw_rows = []
            st.session_state.pdf_det_filtered_rows = []
            st.session_state.pdf_det_machine_name = "Unknown"
            st.session_state.pdf_det_excel_bytes = None
            st.session_state.pdf_det_filename = ""
            
        col_pdf_btn1, col_pdf_btn2, col_pdf_btn3 = st.columns([1, 2, 1])
        with col_pdf_btn2:
            detect_btn = st.button("🔍 Process Excel File", type="primary", use_container_width=True, key="etk_pdf_detect_btn")
            
    if detect_btn:
        if not pdf_file:
            st.error("Please upload an Excel file first.")
        else:
            with st.spinner("Extracting drawing status from Excel..."):
                try:
                    # Run extraction
                    machine_val, rows = extract_drawing_status_from_excel(
                        pdf_file.getvalue()
                    )
                    
                    # Filter rows where Stat != 'FR' (case-insensitive)
                    filtered_rows = [
                        r for r in rows
                        if str(r.get("Stat", "")).strip().upper() != "FR"
                    ]
                    
                    # Generate Excel bytes
                    excel_data = generate_dwg_verification_excel(rows, filtered_rows, machine_val)
                    
                    st.session_state.pdf_det_raw_rows = rows
                    st.session_state.pdf_det_filtered_rows = filtered_rows
                    st.session_state.pdf_det_machine_name = machine_val
                    st.session_state.pdf_det_excel_bytes = excel_data
                    st.session_state.pdf_det_filename = pdf_file.name
                    st.session_state.pdf_det_completed = True
                    
                except Exception as e:
                    st.error(f"Error during Excel extraction: {str(e)}")
                    
    # Render Review and Download section if completed
    if st.session_state.get("pdf_det_completed"):
        st.markdown("---")
        st.success("✅ Drawing Status Extracted Successfully!")
        
        # Display Machine Name
        st.markdown(f"### ⚙️ Machine: **{st.session_state.pdf_det_machine_name}**")
        
        # Table display
        st.markdown("### 📊 Live Drawing Status Review")
        
        def style_dwg_dataframe(df, is_filtered_view=False):
            def highlight_row(row):
                vs = str(row.get("Vs", "")).strip()
                stat = str(row.get("Stat", "")).strip().upper()
                
                # 1. Vs is empty or blank -> Orange
                if vs == "" or vs.lower() == "none" or vs.lower() == "nan":
                    return ["background-color: #ff9f1c; color: #000000; font-weight: bold;"] * len(row)
                # 2. Stat != 'FR' (or blank status) -> Yellow
                if is_filtered_view or stat != "FR":
                    return ["background-color: #ffd166; color: #000000;"] * len(row)
                return [""] * len(row)
            return df.style.apply(highlight_row, axis=1)
            
        preview_df = pd.DataFrame(st.session_state.pdf_det_filtered_rows)
        if not preview_df.empty:
            # Ensure columns are in the correct order
            columns_order = ["Including", "drawings", "RevFix", "Vs", "Stat"]
            cols_to_show = [c for c in columns_order if c in preview_df.columns]
            preview_df = preview_df[cols_to_show]
            
            st.dataframe(preview_df, use_container_width=True)
        else:
            st.info("No rows matching the filter (Stat != 'FR') were found.")
            
        # Download Excel button
        base_name, _ = os.path.splitext(st.session_state.pdf_det_filename)
        st.download_button(
            label="⬇️ Download Drawing Status Excel Report",
            data=st.session_state.pdf_det_excel_bytes,
            file_name=f"{base_name}_Drawing_Status.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="download_pdf_det_excel"
        )
        
        # Debug Expander
        with st.expander("🔍 Debug: View Raw Extracted Rows"):
            st.write("All extracted rows before filtering:")
            raw_df = pd.DataFrame(st.session_state.pdf_det_raw_rows)
            if not raw_df.empty:
                columns_order = ["Including", "drawings", "RevFix", "Vs", "Stat"]
                cols_to_show = [c for c in columns_order if c in raw_df.columns]
                raw_df = raw_df[cols_to_show]
                
                styled_raw_df = style_dwg_dataframe(raw_df, is_filtered_view=False)
                st.dataframe(styled_raw_df, use_container_width=True)
            else:
                st.dataframe(raw_df)

def render_ocr_adjustment_page():
    st.title("✍️ OCR Adjustment Change")
    st.subheader("Extract Handwritten Values & Populate Excel Templates")
    
    # Initialize session state for OCR Adjustment
    if "adj_mapping" not in st.session_state:
        st.session_state.adj_mapping = pd.DataFrame(DEFAULT_ADJ_MAPPING)
    if "adj_excel_bytes" not in st.session_state:
        st.session_state.adj_excel_bytes = None
    if "adj_excel_filename" not in st.session_state:
        st.session_state.adj_excel_filename = None
    if "adj_pdf_bytes" not in st.session_state:
        st.session_state.adj_pdf_bytes = None
    if "adj_pdf_filename" not in st.session_state:
        st.session_state.adj_pdf_filename = None
    if "adj_pdf_images" not in st.session_state:
        st.session_state.adj_pdf_images = None
    if "adj_ocr_results" not in st.session_state:
        st.session_state.adj_ocr_results = {}
        
    # Render Workflow Ribbon Tab
    sub_tab1, sub_tab2, sub_tab3, sub_tab4, sub_tab5, sub_tab6, sub_tab7 = st.tabs([
        "📋 Overview",
        "⚙️ Excel Mapping",
        "📂 Upload Template",
        "📄 Upload PDF",
        "⚡ OCR Processing",
        "✍️ Review & Correction",
        "📥 Export"
    ])
    
    with sub_tab1:
        st.markdown("### 📊 Workflow Overview")
        st.write("This module automates the extraction of handwritten numbers from scanned PDF certificates and writes them directly into your formatted Excel templates.")
        
        # Display Status checklist
        st.markdown("#### 📂 Upload Status")
        col_st1, col_st2 = st.columns(2)
        with col_st1:
            if st.session_state.adj_excel_bytes:
                st.success(f"✅ Excel Template Uploaded: `{st.session_state.adj_excel_filename}`")
            else:
                st.warning("⚠️ Excel Template: Not Uploaded yet")
        with col_st2:
            if st.session_state.adj_pdf_bytes:
                st.success(f"✅ PDF Scanned Form Uploaded: `{st.session_state.adj_pdf_filename}`")
            else:
                st.warning("⚠️ PDF Scanned Form: Not Uploaded yet")
                
        st.markdown("""
        #### 🔄 Steps to Complete:
        1. **Excel Mapping**: Verify the coordinate mapping between the PDF fields and target Excel cells.
        2. **Upload Template**: Upload your master `.xlsx` workbook.
        3. **Upload PDF**: Upload the scanned form.
        4. **OCR Processing**: Run high-accuracy deep learning OCR on cropped fields.
        5. **Review & Correction**: Audit detected handwriting crops and correct OCR values.
        6. **Export**: Export the final Excel workbook filled with your corrected values.
        """)
        
    with sub_tab2:
        st.markdown("### ⚙️ Template & Cell Mapping Configuration")
        st.write("Map each target Excel cell to the specific coordinates of the scanned PDF form. You can edit the values directly in the table below.")
        
        # Editable DataFrame
        edited_df = st.data_editor(
            st.session_state.adj_mapping,
            num_rows="dynamic",
            use_container_width=True
        )
        st.session_state.adj_mapping = edited_df
        
        # Visual Bounding Box Preview
        if st.session_state.adj_pdf_images:
            st.markdown("#### 👁️ Bounding Box Visual Overlay")
            st.write("Check if red boxes align correctly with the handwritten cells on your form.")
            try:
                preview_img = draw_mapping_boxes(st.session_state.adj_pdf_images[0], edited_df)
                st.image(preview_img, use_container_width=True)
            except Exception as e:
                st.error(f"Failed to generate preview: {e}")
        else:
            st.info("ℹ️ Upload a scanned PDF form in the **Upload PDF** tab to see a visual box overlay preview.")
            
    with sub_tab3:
        st.markdown("### 📂 Upload Excel Template")
        excel_file = st.file_uploader("Choose Excel Template file (.xlsx)", type=["xlsx"])
        if excel_file:
            st.session_state.adj_excel_bytes = excel_file.read()
            st.session_state.adj_excel_filename = excel_file.name
            st.success(f"Successfully loaded master Excel template: `{excel_file.name}`")
            
    with sub_tab4:
        st.markdown("### 📄 Upload Scanned PDF Form")
        pdf_file = st.file_uploader("Choose Scanned PDF file (.pdf)", type=["pdf"])
        if pdf_file:
            st.session_state.adj_pdf_bytes = pdf_file.read()
            st.session_state.adj_pdf_filename = pdf_file.name
            # Convert PDF to PIL Images
            try:
                st.session_state.adj_pdf_images = convert_pdf_to_images(st.session_state.adj_pdf_bytes, dpi=150)
                st.success(f"Successfully loaded and rendered `{pdf_file.name}` ({len(st.session_state.adj_pdf_images)} pages)")
            except Exception as e:
                st.error(f"Failed to convert PDF: {e}")
                
    with sub_tab5:
        st.markdown("### ⚡ Run OCR Processing")
        st.write("Process the cropped target cells using the deep learning OCR engine to detect handwritten values.")
        
        # Choose Engine
        ocr_engine = st.selectbox(
            "Select Handwriting Recognition Engine",
            options=["Surya AI OCR (Deep Learning)", "Windows Media OCR (Native)"] if SURYA_AVAILABLE else ["Windows Media OCR (Native)"]
        )
        
        if not st.session_state.adj_pdf_images:
            st.warning("⚠️ Please upload a scanned PDF form first.")
        else:
            if st.button("Start Handwriting OCR Extraction", type="primary"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                results = {}
                mapping_df = st.session_state.adj_mapping
                total_fields = len(mapping_df)
                
                for idx, row in mapping_df.iterrows():
                    name = row["Field Name"]
                    cell = row["Excel Cell"]
                    page_num = int(row["PDF Page"]) - 1
                    top = float(row["Top (%)"])
                    left = float(row["Left (%)"])
                    h = float(row["Height (%)"])
                    w = float(row["Width (%)"])
                    
                    status_text.text(f"Extracting field: '{name}' on page {page_num+1}...")
                    
                    # Get correct page image
                    if page_num < len(st.session_state.adj_pdf_images):
                        page_img = st.session_state.adj_pdf_images[page_num]
                        crop_img = crop_by_percent(page_img, top, left, h, w)
                        
                        # Run OCR
                        text, conf = run_crop_ocr(crop_img, ocr_engine)
                        results[name] = {
                            "Field Name": name,
                            "Excel Cell": cell,
                            "Detected Value": text,
                            "Corrected Value": text,
                            "Confidence": conf,
                            "Crop Image": crop_img
                        }
                    progress_bar.progress((idx + 1) / total_fields)
                    
                st.session_state.adj_ocr_results = results
                status_text.text("OCR processing completed successfully! ✅")
                st.success("Extraction done. Proceed to **Review & Correction** tab to audit values.")
                
    with sub_tab6:
        st.markdown("### ✍️ Review & Correct Extracted Handwriting")
        st.write("Audit detected values alongside cropped pictures of the handwritten fields.")
        
        if not st.session_state.adj_ocr_results:
            st.info("ℹ️ No OCR results yet. Please run the OCR processing task in the **OCR Processing** tab.")
        else:
            for name, item in list(st.session_state.adj_ocr_results.items()):
                with st.container(border=True):
                    col_info, col_crop, col_input = st.columns([2, 2, 3])
                    with col_info:
                        st.markdown(f"**Field:** `{name}`")
                        st.markdown(f"**Excel Cell:** `{item['Excel Cell']}`")
                        conf_pct = int(item['Confidence'] * 100)
                        st.markdown(f"**OCR Value:** `{item['Detected Value']}`")
                        st.markdown(f"**Confidence:** `{conf_pct}%`")
                        st.progress(item['Confidence'])
                    with col_crop:
                        st.markdown("**Handwriting Crop:**")
                        st.image(item["Crop Image"], use_container_width=True)
                    with col_input:
                        corrected_val = st.text_input(
                            f"Confirm / Correct Value for '{name}'",
                            value=item["Corrected Value"],
                            key=f"corr_{name}"
                        )
                        st.session_state.adj_ocr_results[name]["Corrected Value"] = corrected_val
                        
    with sub_tab7:
        st.markdown("### 📥 Export Formatted Workbook")
        st.write("Review completed data writeback sheet and download the completed Excel file.")
        
        if not st.session_state.adj_excel_bytes:
            st.warning("⚠️ Master Excel template is missing. Please upload it in the **Upload Template** tab.")
        elif not st.session_state.adj_ocr_results:
            st.warning("⚠️ No corrected values available. Please run OCR and verify corrections first.")
        else:
            if st.button("Generate & Write to Excel Workbook", type="primary"):
                try:
                    final_excel = write_values_to_excel(
                        st.session_state.adj_excel_bytes,
                        st.session_state.adj_ocr_results
                    )
                    st.success("Successfully populated all values to Excel! 🎉")
                    st.download_button(
                        label="⬇️ Download Completed Excel File",
                        data=final_excel,
                        file_name="completed_inspection_form.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Failed to generate Excel sheet: {e}")

# ---------------------------------------------------------------------------
# Advanced OCR Adjustment (Qwen vLLM)
# ---------------------------------------------------------------------------
def render_advanced_ocr_adjustment_page():
    st.title("🌟 Handwriting Form-to-Excel OCR with Custom Calibration")
    st.subheader("Perspective Alignment, Vision-LLM OCR, Few-Shot Embedding Calibration & Multi-Sheet Excel Writeback")
    
    # Initialize session state
    if "qwen_adj_mapping" not in st.session_state:
        # Default dynamic mapping matching sample 57511 Excel sheet F1
        st.session_state.qwen_adj_mapping = pd.DataFrame([
            {"Field Name": "FormatNo", "Sheet": "F1", "Excel Cell": "F5", "PDF Page": 1, "Top (%)": 19.3, "Left (%)": 57.2, "Height (%)": 2.0, "Width (%)": 20.0},
            {"Field Name": "Dimension", "Sheet": "F1", "Excel Cell": "E6", "PDF Page": 1, "Top (%)": 21.3, "Left (%)": 57.2, "Height (%)": 2.0, "Width (%)": 35.0},
            {"Field Name": "Description", "Sheet": "F1", "Excel Cell": "E7", "PDF Page": 1, "Top (%)": 23.3, "Left (%)": 57.2, "Height (%)": 2.0, "Width (%)": 35.0},
            {"Field Name": "Speed", "Sheet": "F1", "Excel Cell": "E9", "PDF Page": 1, "Top (%)": 26.9, "Left (%)": 54.9, "Height (%)": 2.2, "Width (%)": 12.0},
            {"Field Name": "CartonChainHeight", "Sheet": "F1", "Excel Cell": "H22", "PDF Page": 1, "Top (%)": 52.4, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "FormatPartsLeftRight", "Sheet": "F1", "Excel Cell": "H24", "PDF Page": 1, "Top (%)": 56.3, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "SuctionArm", "Sheet": "F1", "Excel Cell": "H25", "PDF Page": 1, "Top (%)": 58.1, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "CounterSuction_Z", "Sheet": "F1", "Excel Cell": "H30", "PDF Page": 1, "Top (%)": 63.6, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "CounterSuction_Y", "Sheet": "F1", "Excel Cell": "H31", "PDF Page": 1, "Top (%)": 66.5, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "LatchOpener_LY", "Sheet": "F1", "Excel Cell": "H33", "PDF Page": 1, "Top (%)": 71.2, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "LatchOpener_LX", "Sheet": "F1", "Excel Cell": "H34", "PDF Page": 1, "Top (%)": 73.0, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "Erection_LOY", "Sheet": "F1", "Excel Cell": "H36", "PDF Page": 1, "Top (%)": 80.6, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0},
            {"Field Name": "Erection_LUY", "Sheet": "F1", "Excel Cell": "H37", "PDF Page": 1, "Top (%)": 83.0, "Left (%)": 76.5, "Height (%)": 2.0, "Width (%)": 15.0}
        ])
    if "qwen_adj_excel_bytes" not in st.session_state:
        st.session_state.qwen_adj_excel_bytes = None
    if "qwen_adj_excel_filename" not in st.session_state:
        st.session_state.qwen_adj_excel_filename = None
    if "qwen_adj_pdf_bytes" not in st.session_state:
        st.session_state.qwen_adj_pdf_bytes = None
    if "qwen_adj_pdf_filename" not in st.session_state:
        st.session_state.qwen_adj_pdf_filename = None
    if "qwen_adj_pdf_images" not in st.session_state:
        st.session_state.qwen_adj_pdf_images = None
    if "qwen_adj_ocr_results" not in st.session_state:
        st.session_state.qwen_adj_ocr_results = {}
    if "calibration_user_name" not in st.session_state:
        st.session_state.calibration_user_name = "Operator_1"
        
    # Tabs
    tab_over, tab_calib, tab_map, tab_up, tab_proc, tab_review, tab_export = st.tabs([
        "📋 Overview",
        "📜 36-Box Calibration",
        "⚙️ Cell Mapping",
        "📂 Upload Files",
        "⚡ OCR & Calibration",
        "✍️ Review UX & Rapid-Fire",
        "📥 Multi-Sheet Export"
    ])
    
    with tab_over:
        st.markdown("### 📊 Project Overview & Pipeline")
        st.markdown("""
        **Workflow Rule (Strict Enforcement)**:
        `PDF -> OCR -> Human Review -> SQLite -> Mapping -> Excel Export`
        
        * **Module 1**: OpenCV Perspective Alignment & Homography Deskewing.
        * **Module 2**: Handwriting Recognition (Qwen 3.6 35B / Fallback) + 36-Box Calibration Embeddings & Cosine Similarity Matcher.
        * **Module 3**: Verification UX with **Split-Screen**, **Rapid-Fire Keyboard Mode (< 75% Confidence)**, and **Active Learning Profile Tuning**.
        * **Module 4**: Dynamic Multi-Sheet `openpyxl` Writeback without formula or formatting loss.
        """)
        
        st.markdown("#### 📂 Active System Status")
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.session_state.qwen_adj_excel_bytes:
                st.success(f"✅ Excel Template: `{st.session_state.qwen_adj_excel_filename}`")
            else:
                st.info("ℹ️ Excel Template: Not Uploaded")
        with c2:
            if st.session_state.qwen_adj_pdf_bytes:
                st.success(f"✅ PDF Form: `{st.session_state.qwen_adj_pdf_filename}`")
            else:
                st.info("ℹ️ PDF Form: Not Uploaded")
        with c3:
            session = get_db_session()
            profile_count = session.query(CalibrationProfile.user_name).distinct().count()
            session.close()
            st.success(f"👤 Calibration Profiles in DB: `{profile_count}`")
            
    with tab_calib:
        st.markdown("### 📜 Personalized Handwriting Calibration Sheet (36-Box Grid)")
        st.write("Generate, print, fill, and train user-specific handwriting profiles for 0-9 and A-Z.")
        
        col_gen, col_upload = st.columns(2)
        with col_gen:
            st.markdown("#### 1️⃣ Print Calibration Template")
            user_input_name = st.text_input("Operator / User Profile Name:", value=st.session_state.calibration_user_name)
            st.session_state.calibration_user_name = user_input_name
            
            pdf_bytes = generate_calibration_sheet_pdf(user_input_name)
            st.download_button(
                label="🖨️ Download Printable 36-Box PDF Calibration Sheet",
                data=pdf_bytes,
                file_name=f"handwriting_calibration_sheet_{user_input_name}.pdf",
                mime="application/pdf",
                use_container_width=True
            )
            st.caption("Print out this A4 form, fill out the 36 boxes (0-9, A-Z) with pen, scan at 300 DPI, and upload on the right.")
            
        with col_upload:
            st.markdown("#### 2️⃣ Upload Scanned Calibration Sheet")
            calib_file = st.file_uploader("Upload Scanned Calibration Sheet (.pdf / .png / .jpg)", type=["pdf", "png", "jpg", "jpeg"], key="calib_up")
            if calib_file:
                if st.button("⚡ Process & Train Calibration Profile", type="primary", use_container_width=True):
                    with st.spinner("Extracting 36 character crops and computing feature embeddings..."):
                        res = process_scanned_calibration_sheet(calib_file.read(), user_input_name)
                        st.success(f"🎉 Successfully trained profile for '{user_input_name}'! Saved {res['chars_saved']} character embeddings to SQLite.")
                        
    with tab_map:
        st.markdown("### ⚙️ Template & Cell Mapping Configuration")
        st.write("Configure target Excel sheets and cell addresses dynamically. Move boxes freely on the canvas with mouse or auto-align.")
        
        col_map_b1, col_map_b2 = st.columns([2, 1])
        with col_map_b2:
            if st.session_state.qwen_adj_pdf_images:
                if st.button("🪄 Auto-Snap Boxes to Form Labels", type="primary", use_container_width=True):
                    from backend.anchor_aligner import auto_detect_field_boxes
                    auto_boxes = auto_detect_field_boxes(st.session_state.qwen_adj_pdf_images[0])
                    st.session_state.qwen_adj_mapping = pd.DataFrame(auto_boxes)
                    st.toast("Successfully snapped bounding boxes to document labels!", icon="✨")
                    st.rerun()
                    
        edited_df = st.data_editor(
            st.session_state.qwen_adj_mapping,
            num_rows="dynamic",
            use_container_width=True
        )
        st.session_state.qwen_adj_mapping = edited_df
        
        if st.session_state.qwen_adj_pdf_images:
            st.markdown("#### 🖱️ Direct Multi-Box Mouse Adjuster (Fabric.js Interactive Canvas)")
            st.caption("Click and drag ANY red box directly on the document image below with your mouse to move or resize it freely!")
            
            from backend.interactive_canvas import render_fabric_canvas
            import streamlit.components.v1 as components
            
            page_img = st.session_state.qwen_adj_pdf_images[0]
            fabric_html = render_fabric_canvas(page_img, edited_df)
            components.html(fabric_html, height=720, scrolling=True)

            st.markdown("#### 🎯 Single-Box Precision Mouse Cropper")
            st.caption("Select a specific field below for fine-grained handle adjustments.")
            
            field_names = edited_df["Field Name"].tolist() if "Field Name" in edited_df else []
            if field_names:
                sel_field = st.selectbox("Choose Field to Move Box:", field_names)
                field_rows = edited_df[edited_df["Field Name"] == sel_field]
                if not field_rows.empty:
                    f_row = field_rows.iloc[0]
                    f_idx = field_rows.index[0]
                    
                    w_img, h_img = page_img.size
                    
                    top_pct = float(f_row["Top (%)"])
                    left_pct = float(f_row["Left (%)"])
                    h_pct = float(f_row["Height (%)"])
                    w_pct = float(f_row["Width (%)"])
                    
                    x1 = int(left_pct * w_img / 100)
                    y1 = int(top_pct * h_img / 100)
                    x2 = int((left_pct + w_pct) * w_img / 100)
                    y2 = int((top_pct + h_pct) * h_img / 100)
                    
                    from streamlit_cropper import st_cropper
                    cropper_box = st_cropper(
                        page_img.convert("RGB"),
                        realtime_update=True,
                        box_color='#EF4444',
                        aspect_ratio=None,
                        default_coords=(x1, y1, x2, y2),
                        return_type='box',
                        key=f"cropper_{sel_field}"
                    )
                    
                    if cropper_box:
                        new_left_pct = round((cropper_box['left'] / w_img) * 100, 1)
                        new_top_pct = round((cropper_box['top'] / h_img) * 100, 1)
                        new_width_pct = round((cropper_box['width'] / w_img) * 100, 1)
                        new_height_pct = round((cropper_box['height'] / h_img) * 100, 1)
                        
                        col_c1, col_c2 = st.columns([1, 1])
                        with col_c1:
                            st.info(f"New Position: Left `{new_left_pct}%`, Top `{new_top_pct}%`, Width `{new_width_pct}%`, Height `{new_height_pct}%`")
                        with col_c2:
                            if st.button(f"💾 Save Moved Box for '{sel_field}'", type="primary", use_container_width=True):
                                st.session_state.qwen_adj_mapping.at[f_idx, "Top (%)"] = new_top_pct
                                st.session_state.qwen_adj_mapping.at[f_idx, "Left (%)"] = new_left_pct
                                st.session_state.qwen_adj_mapping.at[f_idx, "Height (%)"] = new_height_pct
                                st.session_state.qwen_adj_mapping.at[f_idx, "Width (%)"] = new_width_pct
                                st.toast(f"Saved new position for '{sel_field}'!", icon="🎉")
                                st.rerun()

            st.markdown("#### 👁️ Bounding Box Visual Overlay Preview")
            try:
                preview_img = draw_mapping_boxes(st.session_state.qwen_adj_pdf_images[0], st.session_state.qwen_adj_mapping)
                st.image(preview_img, use_container_width=True)
            except Exception as e:
                st.error(f"Failed to generate box preview: {e}")

                
    with tab_up:
        st.markdown("### 📂 Upload Master Files")
        col_ex, col_pdf = st.columns(2)
        with col_ex:
            st.markdown("#### Master Excel Template (.xlsx)")
            excel_file = st.file_uploader("Choose Excel Template file", type=["xlsx"], key="qwen_excel_up_v2")
            if excel_file:
                st.session_state.qwen_adj_excel_bytes = excel_file.read()
                st.session_state.qwen_adj_excel_filename = excel_file.name
                st.success(f"Loaded Excel template: `{excel_file.name}`")
        with col_pdf:
            st.markdown("#### Scanned Form PDF (.pdf)")
            pdf_file = st.file_uploader("Choose Scanned PDF Form", type=["pdf"], key="qwen_pdf_up_v2")
            if pdf_file:
                st.session_state.qwen_adj_pdf_bytes = pdf_file.read()
                st.session_state.qwen_adj_pdf_filename = pdf_file.name
                try:
                    imgs = convert_pdf_to_images(st.session_state.qwen_adj_pdf_bytes, dpi=150)
                    st.session_state.qwen_adj_pdf_images = imgs
                    st.success(f"Rendered {len(imgs)} page(s) from `{pdf_file.name}`")
                except Exception as e:
                    st.error(f"Failed to convert PDF: {e}")
                    
    with tab_proc:
        st.markdown("### ⚡ Run Alignment, OCR & Calibration Processing")
        
        if not st.session_state.qwen_adj_pdf_images:
            st.warning("⚠️ Please upload a scanned PDF form in the **Upload Files** tab first.")
        else:
            col_opt1, col_opt2 = st.columns(2)
            with col_opt1:
                use_alignment = st.checkbox("Enable OpenCV Homography Perspective Alignment (Deskew)", value=True)
            with col_opt2:
                session = get_db_session()
                available_profiles = [p[0] for p in session.query(CalibrationProfile.user_name).distinct().all()]
                session.close()
                if not available_profiles:
                    available_profiles = ["Default"]
                selected_profile = st.selectbox("Select Active Calibration Profile:", available_profiles)
                
            if st.button("🚀 Start Form Extraction & Calibration Matching", type="primary", use_container_width=True):
                progress_bar = st.progress(0)
                status_text = st.empty()
                results = {}
                mapping_df = st.session_state.qwen_adj_mapping
                total_fields = len(mapping_df)
                
                # Database session for audit trail
                session = get_db_session()
                doc_id = st.session_state.qwen_adj_pdf_filename or "DOC_001"
                
                for idx, row in mapping_df.iterrows():
                    name = str(row["Field Name"])
                    sheet = str(row.get("Sheet", "F1"))
                    cell = str(row["Excel Cell"])
                    cell_full = f"{sheet}!{cell}"
                    
                    page_num = int(row["PDF Page"]) - 1
                    top = float(row["Top (%)"])
                    left = float(row["Left (%)"])
                    h = float(row["Height (%)"])
                    w = float(row["Width (%)"])
                    
                    status_text.text(f"Processing Field [{idx+1}/{total_fields}]: '{name}' ({cell_full})...")
                    
                    if page_num < len(st.session_state.qwen_adj_pdf_images):
                        raw_page_img = st.session_state.qwen_adj_pdf_images[page_num]
                        
                        # Apply OpenCV Homography Perspective Alignment if enabled
                        if use_alignment:
                            page_img = align_document_page(raw_page_img, target_width=raw_page_img.width, target_height=raw_page_img.height)
                        else:
                            page_img = raw_page_img
                            
                        crop_img = crop_by_percent(page_img, top, left, h, w)
                        
                        # Primary OCR (Qwen / Fallback)
                        text, conf = run_qwen_ocr(crop_img)
                        
                        # Few-Shot Cosine Similarity Calibration Matching
                        final_text, final_conf, top_matches = classify_with_calibration(
                            crop_img, selected_profile, text, conf
                        )
                        
                        results[name] = {
                            "Field Name": name,
                            "Sheet": sheet,
                            "Excel Cell": cell,
                            "Cell Full": cell_full,
                            "Detected Value": text,
                            "Corrected Value": final_text,
                            "Confidence": final_conf,
                            "Crop Image": crop_img,
                            "Top Matches": top_matches
                        }
                        
                        # Record audit trail in SQLite ocr_data table
                        db_record = OCRData(
                            document_id=doc_id,
                            field_name=name,
                            sheet_name=sheet,
                            excel_cell=cell,
                            ocr_value=text,
                            corrected_value=final_text,
                            confidence=final_conf,
                            engine_used="Qwen3.6-vLLM + CosineCalibration",
                            status="Reviewed" if final_conf >= 0.75 else "Pending"
                        )
                        session.add(db_record)
                        
                    progress_bar.progress((idx + 1) / total_fields)
                    
                session.commit()
                session.close()
                
                st.session_state.qwen_adj_ocr_results = results
                status_text.text("Extraction and Cosine Similarity Calibration completed! ✅")
                st.success("Proceed to **Review UX & Rapid-Fire** tab to verify results.")
                
    with tab_review:
        st.markdown("### ✍️ Human Verification & Rapid-Fire UX Workflow")
        
        if not st.session_state.qwen_adj_ocr_results:
            st.info("ℹ️ No OCR extraction results available yet. Run processing in Tab 5.")
        else:
            mode = st.radio("Select Verification Mode:", ["1️⃣ Split-Screen View", "2️⃣ Rapid-Fire Keyboard Mode (< 75% Confidence)", "3️⃣ Profile-Tuning Active Learning"], horizontal=True)
            
            if mode == "1️⃣ Split-Screen View":
                col_left, col_right = st.columns([1, 1])
                with col_left:
                    st.markdown("#### 📄 Document Visualizer")
                    if st.session_state.qwen_adj_pdf_images:
                        preview_img = draw_mapping_boxes(st.session_state.qwen_adj_pdf_images[0], st.session_state.qwen_adj_mapping)
                        st.image(preview_img, use_container_width=True)
                with col_right:
                    st.markdown("#### 📝 Verification Data Table")
                    for name, item in list(st.session_state.qwen_adj_ocr_results.items()):
                        conf = item["Confidence"]
                        status_color = "🟢" if conf >= 0.75 else "🔴"
                        with st.container(border=True):
                            c_snippet, c_details = st.columns([1, 2])
                            with c_snippet:
                                st.image(item["Crop Image"], caption=f"{item['Field Name']} ({item['Excel Cell']})", use_container_width=True)
                            with c_details:
                                st.markdown(f"**{name}** (`{item['Cell Full']}`) {status_color} Conf: `{conf*100:.0f}%`")
                                new_val = st.text_input(
                                    f"Value:",
                                    value=item["Corrected Value"],
                                    key=f"split_corr_{name}"
                                )
                                st.session_state.qwen_adj_ocr_results[name]["Corrected Value"] = new_val
                                
            elif mode == "2️⃣ Rapid-Fire Keyboard Mode (< 75% Confidence)":
                low_conf_items = {k: v for k, v in st.session_state.qwen_adj_ocr_results.items() if v["Confidence"] < 0.75}
                if not low_conf_items:
                    st.success("🎉 All fields have high confidence (≥ 75%)! No rapid-fire corrections needed.")
                else:
                    st.warning(f"⚠️ `{len(low_conf_items)}` field(s) require rapid keyboard review.")
                    st.caption("Press Enter in each text box to quickly confirm and move to the next field.")
                    
                    for name, item in low_conf_items.items():
                        with st.container(border=True):
                            c_img, c_in = st.columns([1, 2])
                            with c_img:
                                st.image(item["Crop Image"], caption=f"Field: {name}", use_container_width=True)
                            with c_in:
                                st.markdown(f"Target: `{item['Cell Full']}` | Confidence: **🔴 {item['Confidence']*100:.0f}%**")
                                corrected = st.text_input(
                                    f"Rapid Confirm '{name}'",
                                    value=item["Corrected Value"],
                                    key=f"rf_corr_{name}"
                                )
                                st.session_state.qwen_adj_ocr_results[name]["Corrected Value"] = corrected
                                
            elif mode == "3️⃣ Profile-Tuning Active Learning":
                st.markdown("#### 🧠 Profile Tuning & Character Embedding Suggestions")
                st.write("Compare extracted crop snippets against stored calibration profile samples.")
                
                for name, item in list(st.session_state.qwen_adj_ocr_results.items()):
                    with st.container(border=True):
                        c1, c2, c3 = st.columns([1, 1, 2])
                        with c1:
                            st.markdown("**Extracted Form Crop:**")
                            st.image(item["Crop Image"], use_container_width=True)
                        with c2:
                            st.markdown("**Profile Matches:**")
                            matches = item.get("Top Matches", [])
                            if matches:
                                for m_char, m_score, m_path in matches:
                                    st.write(f"- Match **'{m_char}'** (Sim: `{m_score:.2f}`)")
                            else:
                                st.caption("No profile matches found.")
                        with c3:
                            st.markdown(f"**Field:** `{name}`")
                            corrected = st.text_input(f"Corrected Value:", value=item["Corrected Value"], key=f"tune_corr_{name}")
                            st.session_state.qwen_adj_ocr_results[name]["Corrected Value"] = corrected
                            
                            if st.button(f"➕ Add sample to '{st.session_state.calibration_user_name}' Profile", key=f"active_learn_{name}"):
                                # Save active learning sample to DB
                                session = get_db_session()
                                emb = compute_image_embedding(item["Crop Image"])
                                p_entry = CalibrationProfile(
                                    user_name=st.session_state.calibration_user_name,
                                    char_code=corrected.strip().upper()[:1] if corrected.strip() else "0",
                                    feature_vector_json=json.dumps(emb.tolist())
                                )
                                session.add(p_entry)
                                session.commit()
                                session.close()
                                st.toast(f"Saved '{corrected}' to calibration profile!", icon="✨")
                                
    with tab_export:
        st.markdown("### 📥 Multi-Sheet Excel Export Integration")
        if not st.session_state.qwen_adj_excel_bytes:
            st.warning("⚠️ Master Excel template missing. Upload in Tab 4.")
        elif not st.session_state.qwen_adj_ocr_results:
            st.warning("⚠️ No extraction results available.")
        else:
            if st.button("🚀 Generate Multi-Sheet Excel File", type="primary", use_container_width=True):
                try:
                    exporter = ExcelExportService()
                    final_excel = exporter.export_excel(
                        st.session_state.qwen_adj_excel_bytes,
                        st.session_state.qwen_adj_ocr_results
                    )
                    
                    st.success("Successfully populated all multi-sheet cells using openpyxl! 🎉")
                    st.download_button(
                        label="⬇️ Download Completed Excel File (.xlsx)",
                        data=final_excel,
                        file_name="completed_adjustment_chart_SAT.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Failed to export Excel workbook: {e}")

# ---------------------------------------------------------------------------
# Calibration Certificate Processing Tool


# ---------------------------------------------------------------------------
# Calibration Certificate Processing Tool
# ---------------------------------------------------------------------------
def render_calibration_certificate_page():
    st.title("📜 Calibration Certificate Processing Tool")
    st.subheader("Automate calibration sheet transformation, serial number generation & image library maintenance")
    
    import backend.calibration_service as calib_svc
    try:
        importlib.reload(calib_svc)
    except Exception:
        pass
        
    tab_calib_process, tab_calib_maint = st.tabs([
        "⚡ 1. Process Calibration Sheet",
        "🖼️ 2. Calibration Image Library & Maintenance"
    ])
    
    # ---------------------------------------------------------------------------
    # TAB 1: Process Calibration Sheet
    # ---------------------------------------------------------------------------
    with tab_calib_process:
        if "active_calib_df" not in st.session_state:
            st.session_state.active_calib_df = None
        
        st.markdown("""
        <div style="background-color: rgba(0, 120, 212, 0.05); padding: 20px; border-radius: 10px; margin-bottom: 25px; border-left: 5px solid #0078d4;">
            <h4 style="margin-top: 0; color: #0078d4;">📋 Process Description</h4>
            <p style="color: var(--text-color); opacity: 0.9; font-size: 1rem; margin-bottom: 0;">
                Upload your calibration Excel sheet to run the automated formatting and structural compilation process. 
                The system will automatically capture values, clean and group sections, build sequential numbering, and format columns.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("---")
        st.markdown("### 🚀 Step 1: Process Calibration")
        st.write("Upload your raw Excel sheet to execute row splitting, serial number generation, and automated image lookups.")

        if "calib_msg_success" in st.session_state:
            st.success(st.session_state.pop("calib_msg_success"))

        col_left, col_right = st.columns(2)
        
        with col_left:
            uploaded_file = st.file_uploader("Upload Raw Calibration Excel Sheet (CSP2)", type=["xlsx"], key="calib_excel_uploader")
            process_btn = st.button(
                "🚀 Process Calibration", 
                type="primary", 
                use_container_width=True, 
                key="calib_process_btn", 
                disabled=(uploaded_file is None)
            )
            
        with col_right:
            missing_uploader_key = f"calib_missing_uploader_{st.session_state.get('calib_missing_uploader_ver', 0)}"
            uploaded_yellow_imgs = st.file_uploader(
                "Upload Missing Images (Drag & Drop here)", 
                type=["png", "jpg", "jpeg", "bmp", "gif", "tiff", "tif", "webp"], 
                accept_multiple_files=True, 
                key=missing_uploader_key
            )
            has_files_or_df = bool(uploaded_yellow_imgs) or (st.session_state.get("active_calib_df") is not None)
            process_missing_btn = st.button(
                "⚡ Process Upload Missing", 
                type="secondary", 
                use_container_width=True, 
                key="calib_process_missing_btn", 
                disabled=(not has_files_or_df)
            )
            
        if uploaded_file and process_btn:
            file_bytes = uploaded_file.read()
            with st.spinner("Processing calibration sheet and running layout transformations..."):
                try:
                    processed_bytes = calib_svc.process_calibration_sheet(file_bytes, uploaded_file.name)
                    
                    st.session_state.calib_raw_filename = uploaded_file.name
                    st.session_state.calib_processed_bytes = processed_bytes
                    st.session_state.calib_processed_filename = f"Processed_{uploaded_file.name}"
                    
                    # Convert to pandas DataFrame for interactive editing
                    import io
                    st.session_state.active_calib_df = pd.read_excel(io.BytesIO(processed_bytes))
                    st.session_state["calib_msg_success"] = "✅ Calibration Sheet Processed Successfully!"
                    st.rerun()
                except Exception as e:
                    st.error(f"Failed to process sheet: {str(e)}")
                    
        if process_missing_btn:
            with st.spinner("Processing missing images and updating calibration data..."):
                saved_count = 0
                if uploaded_yellow_imgs:
                    for uploaded_img in uploaded_yellow_imgs:
                        try:
                            calib_svc.save_calibration_image(uploaded_img.getvalue(), uploaded_img.name, overwrite=True)
                            saved_count += 1
                        except Exception as save_err:
                            st.error(f"Error saving image {uploaded_img.name}: {save_err}")
                
                df_to_use = st.session_state.get("active_calib_df")
                updated_count = 0
                if df_to_use is not None:
                    prefix = ""
                    dict_counts = {}
                    for r_idx in range(len(df_to_use)):
                        row = df_to_use.iloc[r_idx]
                        cell_A_val = row.iloc[0]
                        is_row_header = cell_A_val is None or str(cell_A_val).strip() in ["", "None", "nan", "<NA>"]
                        
                        if is_row_header:
                            prefix = str(row.iloc[1] or "").strip()
                        else:
                            pic_val = str(row.iloc[7] or "")
                            if "File Not Found" in pic_val:
                                item_no = str(row.iloc[1] or "").strip()
                                base_name = f"{prefix}@{item_no}"
                                
                                if base_name in dict_counts:
                                    dict_counts[base_name] += 1
                                    file_name = f"{base_name}_{dict_counts[base_name]}"
                                else:
                                    dict_counts[base_name] = 0
                                    file_name = base_name
                                    
                                img_path = calib_svc.find_image_file(file_name)
                                if img_path:
                                    df_to_use.iloc[r_idx, 7] = "" # Clear error text!
                                    updated_count += 1
                                    
                    if updated_count > 0:
                        try:
                            from backend.calibration_service import write_df_to_excel
                            st.session_state.calib_processed_bytes = write_df_to_excel(df_to_use)
                            st.session_state.active_calib_df = df_to_use
                        except Exception as compile_err:
                            st.error(f"Error compiling Excel: {compile_err}")
                            
                # Reset uploader version so the dropzone is cleared for subsequent drops
                st.session_state["calib_missing_uploader_ver"] = st.session_state.get("calib_missing_uploader_ver", 0) + 1
                
                msg_parts = []
                if saved_count > 0:
                    msg_parts.append(f"Saved {saved_count} image file(s) successfully")
                if updated_count > 0:
                    msg_parts.append(f"Updated {updated_count} table row(s) successfully")
                
                if msg_parts:
                    st.session_state["calib_msg_success"] = f"✅ Complete! ({ ', '.join(msg_parts) })"
                elif saved_count == 0 and updated_count == 0:
                    st.session_state["calib_msg_success"] = "✅ Complete! (No pending missing images found)"
                
                st.rerun()
                        
        # If processed dataframe exists, render preview and action toolbar
        df_to_use = st.session_state.active_calib_df
        if df_to_use is not None and not df_to_use.empty:
            with st.expander("📊 Preview Processed Calibration Certificate", expanded=True):
                display_mode = st.radio(
                    "📌 Select Display Mode:",
                    ["🎨 Visual Status View (Colors & Bold)", "✏️ Editable Table Mode"],
                    index=0,
                    horizontal=True,
                    key="calib_mode_select"
                )
                
                if display_mode == "🎨 Visual Status View (Colors & Bold)":
                    def style_calib_dataframe(df):
                        def highlight_rows(row):
                            val_a = row.iloc[0]
                            is_header = False
                            if val_a is None or val_a != val_a or str(val_a).strip() in ["", "None", "nan", "<NA>"]:
                                is_header = True
                            if is_header:
                                return ['font-weight: bold; background-color: rgba(0, 120, 212, 0.05);'] * len(row)
                            
                            # Highlight row in yellow if image is missing
                            pic_val = str(row.iloc[7] or "")
                            if "File Not Found" in pic_val:
                                return ['background-color: #FFF2CC; color: #000000;'] * len(row)
                                
                            return [''] * len(row)
                        return df.style.apply(highlight_rows, axis=1)
                        
                    styled_df = style_calib_dataframe(df_to_use)
                    st.dataframe(styled_df, use_container_width=True)
                else:
                    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode, GridUpdateMode
                    
                    col_names = df_to_use.columns.tolist()
                    colA_name = col_names[0] if len(col_names) > 0 else ""
                    
                    jscode = f"""
                    function(params) {{
                        var styles = {{}};
                        var colA = params.data['{colA_name}'];
                        if (colA === null || colA === undefined || String(colA).trim() === "" || String(colA).trim() === "None") {{
                            styles['font-weight'] = 'bold';
                            styles['background-color'] = 'rgba(0, 120, 212, 0.05)';
                        }}
                        
                        var pic = params.data['Picture'];
                        if (pic && String(pic).indexOf('File Not Found') !== -1) {{
                            styles['background-color'] = '#FFF2CC';
                            styles['color'] = '#000000';
                        }}
                        return styles;
                    }}
                    """
                    
                    df_to_aggrid = df_to_use.copy()
                    df_to_aggrid['_row_id'] = range(len(df_to_aggrid))
                    
                    # Create a container to render the toolbar above the grid
                    toolbar_container = st.container()
                    
                    gb = GridOptionsBuilder.from_dataframe(df_to_aggrid)
                    gb.configure_default_column(editable=True, sortable=False, filter=True)
                    gb.configure_column('_row_id', hide=True)
                    gb.configure_grid_options(
                        getRowStyle=JsCode(jscode),
                        quickFilterText=""  # Will be controlled by search input inside container
                    )
                    gb.configure_selection('multiple', use_checkbox=True)
                    grid_options = gb.build()
                    
                    grid_response = AgGrid(
                        df_to_aggrid,
                        gridOptions=grid_options,
                        update_mode=GridUpdateMode.MODEL_CHANGED,
                        allow_unsafe_jscode=True,
                        theme="streamlit",
                        key="calib_aggrid"
                    )
                    
                    updated_df = pd.DataFrame(grid_response['data'])
                    if '_row_id' in updated_df.columns:
                        updated_df = updated_df.drop(columns=['_row_id'])
                    
                    if st.session_state.active_calib_df is None or not updated_df.equals(st.session_state.active_calib_df):
                        st.session_state.active_calib_df = updated_df
                        st.rerun()
                        
                    # Get selected rows
                    selected_rows = grid_response.get("selected_rows", [])
                    sel_ids = []
                    if isinstance(selected_rows, pd.DataFrame) and not selected_rows.empty:
                        if '_row_id' in selected_rows.columns:
                            sel_ids = selected_rows['_row_id'].tolist()
                    elif isinstance(selected_rows, list) and selected_rows:
                        sel_ids = [s['_row_id'] for s in selected_rows if '_row_id' in s]
                        
                    # Render toolbar content above the grid
                    with toolbar_container:
                        tb_col1, tb_col2, tb_col3, tb_col4 = st.columns([3, 2.2, 2.2, 2.2])
                        with tb_col1:
                            search_text = st.text_input("Search", key="calib_tb_search", label_visibility="collapsed", placeholder="🔍 Search table...")
                        with tb_col2:
                            if st.button("⬆️ Insert Above", use_container_width=True, key="calib_add_row_above"):
                                new_row = pd.DataFrame([[None]*len(updated_df.columns)], columns=updated_df.columns)
                                if sel_ids:
                                    df_from_grid = pd.DataFrame(grid_response['data'])
                                    if '_row_id' in df_from_grid.columns:
                                        target_idx = df_from_grid[df_from_grid['_row_id'] == sel_ids[0]].index[0]
                                        df1 = updated_df.iloc[:target_idx]
                                        df2 = updated_df.iloc[target_idx:]
                                        st.session_state.active_calib_df = pd.concat([df1, new_row, df2], ignore_index=True)
                                        st.rerun()
                                else:
                                    st.session_state.active_calib_df = pd.concat([new_row, updated_df], ignore_index=True)
                                    st.rerun()
                        with tb_col3:
                            if st.button("⬇️ Insert Below", use_container_width=True, key="calib_add_row_below"):
                                new_row = pd.DataFrame([[None]*len(updated_df.columns)], columns=updated_df.columns)
                                if sel_ids:
                                    df_from_grid = pd.DataFrame(grid_response['data'])
                                    if '_row_id' in df_from_grid.columns:
                                        target_idx = df_from_grid[df_from_grid['_row_id'] == sel_ids[-1]].index[0]
                                        df1 = updated_df.iloc[:target_idx+1]
                                        df2 = updated_df.iloc[target_idx+1:]
                                        st.session_state.active_calib_df = pd.concat([df1, new_row, df2], ignore_index=True)
                                        st.rerun()
                                else:
                                    st.session_state.active_calib_df = pd.concat([updated_df, new_row], ignore_index=True)
                                    st.rerun()
                        with tb_col4:
                            if st.button("🗑️ Delete Selected", use_container_width=True, key="calib_del_row"):
                                if sel_ids:
                                    df_filtered = pd.DataFrame(grid_response['data'])
                                    df_filtered = df_filtered[~df_filtered['_row_id'].isin(sel_ids)]
                                    if '_row_id' in df_filtered.columns:
                                        df_filtered = df_filtered.drop(columns=['_row_id'])
                                    st.session_state.active_calib_df = df_filtered.reset_index(drop=True)
                                    st.rerun()
                                else:
                                    st.warning("⚠️ Select rows first.")
                                    
                    # Update grid options quickFilterText from search input
                    if search_text:
                        gb.configure_grid_options(quickFilterText=search_text)
                        grid_options = gb.build()

            # Compile DataFrame back to styled Excel workbook bytes
            try:
                from backend.calibration_service import write_df_to_excel
                st.session_state.calib_processed_bytes = write_df_to_excel(st.session_state.active_calib_df)
            except Exception as compile_err:
                st.error(f"Error compiling Excel: {compile_err}")

            # Show 3 balanced metrics as requested
            expected_serials = df_to_use['No.'].dropna().astype(str).str.strip().str.len().gt(0).sum() if 'No.' in df_to_use.columns else 0
            generated_serials = df_to_use['Serial No.'].dropna().astype(str).str.strip().str.len().gt(0).sum() if 'Serial No.' in df_to_use.columns else 0
            missing_serials = max(0, expected_serials - generated_serials)
            
            m1, m2, m3 = st.columns(3)
            m1.metric("SN:Total", generated_serials)
            m2.metric("SN: Missing", missing_serials)
            m3.metric("SN: Generated ", expected_serials)
            
            raw_filename_active = uploaded_file.name if uploaded_file else st.session_state.get("calib_raw_filename", "Calibration_Sheet.xlsx")
            proc_filename_active = st.session_state.get("calib_processed_filename", f"Processed_{raw_filename_active}")
            
            st.download_button(
                label="📥 Download Processed Excel Workbook",
                data=st.session_state.calib_processed_bytes,
                file_name=proc_filename_active,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="calib_download"
            )

            st.markdown("---")
            st.markdown("### 📐 Step 2: Final Layout Settings")
            st.write("Apply landscape print layout, insert the logo header, adjust column widths, add thin borders, and drop helper columns.")
            
            if "calib_final_bytes" not in st.session_state:
                st.session_state.calib_final_bytes = None
            if "calib_final_filename" not in st.session_state:
                st.session_state.calib_final_filename = None
                
            col_final_btn, _ = st.columns([1, 3])
            with col_final_btn:
                run_final_layout = st.button("📐 Set Final Layout", type="primary", use_container_width=True, key="calib_run_final")
                
            if run_final_layout:
                with st.spinner("Applying final layout transformations, borders, page setup, and header logo..."):
                    try:
                        fn_to_use = uploaded_file.name if uploaded_file else st.session_state.get("calib_raw_filename", "Calibration_Sheet.xlsx")
                        final_bytes = calib_svc.set_final_layout(st.session_state.active_calib_df, fn_to_use, layout_type="standard")
                        st.session_state.calib_final_bytes = final_bytes
                        st.session_state.calib_final_filename = f"Final_{fn_to_use}"
                        st.success("✅ Final Layout Applied Successfully!")
                    except Exception as fe:
                        st.error(f"Failed to apply final layout: {str(fe)}")
                        
            # If final bytes exist, display preview of columns A-I and download final button
            if st.session_state.calib_final_bytes is not None:
                try:
                    import io
                    df_final = pd.read_excel(io.BytesIO(st.session_state.calib_final_bytes))
                    st.markdown("#### 👁️ Final Printed Sheet Preview")
                    st.dataframe(df_final, use_container_width=True)
                    
                    st.download_button(
                        label="📥 Download Final Calibration Layout Excel",
                        data=st.session_state.calib_final_bytes,
                        file_name=st.session_state.calib_final_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="calib_final_download"
                    )
                except Exception as preview_err:
                    st.error(f"Error loading final preview: {preview_err}")

            st.markdown("---")
            st.markdown("### 📝 Step 3: Set Layout Approval")
            st.write("Apply approval print layout: insert columns for IWK/Supplier check, insert checkboxes, and set approval signature footer.")
            
            if "calib_approval_bytes" not in st.session_state:
                st.session_state.calib_approval_bytes = None
            if "calib_approval_filename" not in st.session_state:
                st.session_state.calib_approval_filename = None
                
            col_approval_btn, _ = st.columns([1, 3])
            with col_approval_btn:
                run_approval_layout = st.button("📝 Set Approval Layout", type="primary", use_container_width=True, key="calib_run_approval")
                
            if run_approval_layout:
                with st.spinner("Applying approval layout transformations, borders, checkboxes, and signatures..."):
                    try:
                        fn_to_use = uploaded_file.name if uploaded_file else st.session_state.get("calib_raw_filename", "Calibration_Sheet.xlsx")
                        approval_bytes = calib_svc.set_final_layout(st.session_state.active_calib_df, fn_to_use, layout_type="approval")
                        st.session_state.calib_approval_bytes = approval_bytes
                        st.session_state.calib_approval_filename = f"Approval_{fn_to_use}"
                        st.success("✅ Approval Layout Applied Successfully!")
                    except Exception as fe:
                        st.error(f"Failed to apply approval layout: {str(fe)}")
                        
            # If approval bytes exist, display preview of columns A-I and download approval button
            if st.session_state.calib_approval_bytes is not None:
                try:
                    import io
                    df_approval = pd.read_excel(io.BytesIO(st.session_state.calib_approval_bytes))
                    st.markdown("#### 👁️ Approval Sheet Preview")
                    st.dataframe(df_approval, use_container_width=True)
                    
                    st.download_button(
                        label="📥 Download Approval Layout Excel",
                        data=st.session_state.calib_approval_bytes,
                        file_name=st.session_state.calib_approval_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="calib_approval_download"
                    )
                except Exception as preview_err:
                    st.error(f"Error loading final preview: {preview_err}")

    # ---------------------------------------------------------------------------
    # TAB 2: Calibration Image Library & Maintenance System (ระบบบำรุงรักษาภาพเครื่องมือ)
    # ---------------------------------------------------------------------------
    with tab_calib_maint:
        st.markdown("### 🖼️ Calibration Image Library & Maintenance System")
        st.caption("Centralized workspace to browse, search, upload, replace, audit, and backup instrument calibration images stored in the `My Picture/` repository.")
        
        # Load all images metadata
        df_all_imgs = calib_svc.list_calibration_images()
        
        # Metrics Bar
        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.metric("🖼️ Total Library Images", f"{len(df_all_imgs)} Files")
        unique_drawings = df_all_imgs["Drawing No"].nunique() if not df_all_imgs.empty else 0
        m_col2.metric("📐 Unique Drawings/Assemblies", f"{unique_drawings} Drawings")
        tot_mb = (df_all_imgs["Size (KB)"].sum() / 1024.0) if not df_all_imgs.empty else 0.0
        m_col3.metric("💾 Total Storage Size", f"{tot_mb:.2f} MB")
        m_col4.metric("📁 Repository Folder", "My Picture/")
        
        st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
        
        subtab_exp, subtab_upload_missing, subtab_backup = st.tabs([
            "🔍 1. Gallery & Image Explorer",
            "📤 2. Upload Missing Images",
            "📦 3. Library Backup & Export"
        ])
        
        # 1. Gallery & Image Explorer
        with subtab_exp:
            st.markdown("##### 🔍 Browse & Search Instrument Pictures")
            
            f_col1, f_col2, f_col3 = st.columns([3, 1.5, 1.5])
            with f_col1:
                search_kw = st.text_input("Search Image Library", placeholder="🔍 Search by Drawing No, Item No, or Filename...", key="calib_img_search_kw")
            with f_col2:
                all_formats = ["All"] + sorted(df_all_imgs["Format"].unique().tolist()) if not df_all_imgs.empty else ["All"]
                sel_format = st.selectbox("Filter Format", all_formats, key="calib_img_fmt_filter")
            with f_col3:
                st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
                view_mode = st.radio("View Mode", ["🖼️ Card / Grid View", "📋 Table View"], horizontal=True, key="calib_img_view_mode", label_visibility="collapsed")
                
            filtered_df = calib_svc.list_calibration_images(search_query=search_kw)
            if sel_format != "All" and not filtered_df.empty:
                filtered_df = filtered_df[filtered_df["Format"] == sel_format]
                
            st.caption(f"Showing **{len(filtered_df)}** of **{len(df_all_imgs)}** images:")
            
            if filtered_df.empty:
                st.info("No calibration images found matching your search criteria.")
            elif view_mode == "📋 Table View":
                st.dataframe(filtered_df, use_container_width=True, hide_index=True)
                
                # Management Actions for Table View
                act_c1, act_c2 = st.columns(2)
                with act_c1:
                    st.markdown("###### 🗑️ Delete Image(s)")
                    del_selected = st.multiselect("Select image(s) to remove", filtered_df["Filename"].tolist(), key="calib_tbl_del_sel")
                    if del_selected:
                        if st.button(f"🗑️ Delete {len(del_selected)} Selected Image(s)", type="secondary", key="calib_tbl_btn_del"):
                            del_res = calib_svc.delete_calibration_images(del_selected)
                            st.success(f"Deleted {len(del_res)} image(s).")
                            st.rerun()
                with act_c2:
                    st.markdown("###### 📥 Download Single Image")
                    dl_selected = st.selectbox("Select image to download", filtered_df["Filename"].tolist(), key="calib_tbl_dl_sel")
                    if dl_selected:
                        img_b = calib_svc.get_calibration_image_bytes(dl_selected)
                        if img_b:
                            st.download_button(f"📥 Download `{dl_selected}`", data=img_b, file_name=dl_selected, use_container_width=True, key="calib_tbl_btn_dl")
            else:
                # Card / Grid View (Paginated to 24 images per page for snappy rendering)
                cards_per_page = 24
                total_pages = max(1, (len(filtered_df) + cards_per_page - 1) // cards_per_page)
                
                page_col1, page_col2 = st.columns([2, 4])
                with page_col1:
                    cur_page = st.number_input("Page", min_value=1, max_value=total_pages, value=1, step=1, key="calib_gallery_page")
                with page_col2:
                    st.caption(f"Page **{cur_page}** of **{total_pages}** ({len(filtered_df)} items)")
                    
                start_idx = (cur_page - 1) * cards_per_page
                end_idx = min(start_idx + cards_per_page, len(filtered_df))
                page_items = filtered_df.iloc[start_idx:end_idx]
                
                # Render 4 items per row
                row_cols = st.columns(4)
                for idx, (_, row) in enumerate(page_items.iterrows()):
                    col_idx = idx % 4
                    with row_cols[col_idx]:
                        with st.container(border=True):
                            img_p = row["Path"]
                            if os.path.exists(img_p):
                                try:
                                    st.image(img_p, use_container_width=True)
                                except Exception:
                                    st.caption("🖼️ Preview unavailable")
                            
                            st.markdown(f"**`{row['Filename']}`**")
                            st.caption(f"📐 Drawing: `{row['Drawing No']}` | Item: `{row['Item No']}`")
                            st.caption(f"📏 `{row['Dimensions']}` • `{row['Size (KB)']} KB`")
                            
                            card_b1, card_b2 = st.columns(2)
                            with card_b1:
                                img_b = calib_svc.get_calibration_image_bytes(row["Filename"])
                                if img_b:
                                    st.download_button("📥", data=img_b, file_name=row["Filename"], key=f"c_dl_{row['Filename']}", help=f"Download {row['Filename']}", use_container_width=True)
                            with card_b2:
                                with st.popover("🗑️", use_container_width=True):
                                    st.markdown(f"Delete `{row['Filename']}`?")
                                    if st.button("Confirm", type="secondary", key=f"c_del_{row['Filename']}", use_container_width=True):
                                        calib_svc.delete_calibration_images([row["Filename"]])
                                        st.rerun()

        # 2. Upload Missing Images
        with subtab_upload_missing:
            st.markdown("##### 📤 Upload Missing Images (Drag & Drop)")
            st.caption("Drop one or multiple missing calibration images (*.png, *.jpg, *.jpeg, *.bmp, *.webp). Once processed, images will be saved directly into 'My Picture/' and synchronized with the active calibration sheet.")
            
            if "calib_maint_msg_success" in st.session_state:
                st.success(st.session_state.pop("calib_maint_msg_success"))
                
            maint_uploader_key = f"calib_maint_missing_uploader_{st.session_state.get('calib_maint_uploader_ver', 0)}"
            maint_missing_files = st.file_uploader(
                "Upload Missing Images (Drag & Drop here)", 
                type=["png", "jpg", "jpeg", "bmp", "gif", "tiff", "tif", "webp"], 
                accept_multiple_files=True, 
                key=maint_uploader_key
            )
            
            if st.button("⚡ Process Upload Missing", type="primary", key="calib_maint_process_missing_btn", disabled=(not maint_missing_files)):
                with st.spinner("Saving images and updating library..."):
                    saved_count = 0
                    for uf in maint_missing_files:
                        try:
                            calib_svc.save_calibration_image(uf.getvalue(), uf.name, overwrite=True)
                            saved_count += 1
                        except Exception as ex:
                            st.error(f"Error saving {uf.name}: {ex}")
                    
                    # Sync with active calibration dataframe if present
                    updated_count = 0
                    df_to_use = st.session_state.get("active_calib_df")
                    if df_to_use is not None:
                        prefix = ""
                        dict_counts = {}
                        for r_idx in range(len(df_to_use)):
                            row = df_to_use.iloc[r_idx]
                            cell_A_val = row.iloc[0]
                            is_row_header = cell_A_val is None or str(cell_A_val).strip() in ["", "None", "nan", "<NA>"]
                            if is_row_header:
                                prefix = str(row.iloc[1] or "").strip()
                            else:
                                pic_val = str(row.iloc[7] or "")
                                if "File Not Found" in pic_val:
                                    item_no = str(row.iloc[1] or "").strip()
                                    base_name = f"{prefix}@{item_no}"
                                    if base_name in dict_counts:
                                        dict_counts[base_name] += 1
                                        file_name = f"{base_name}_{dict_counts[base_name]}"
                                    else:
                                        dict_counts[base_name] = 0
                                        file_name = base_name
                                    img_path = calib_svc.find_image_file(file_name)
                                    if img_path:
                                        df_to_use.iloc[r_idx, 7] = ""
                                        updated_count += 1
                        if updated_count > 0:
                            from backend.calibration_service import write_df_to_excel
                            st.session_state.calib_processed_bytes = write_df_to_excel(df_to_use)
                            st.session_state.active_calib_df = df_to_use
                            
                    msg_parts = []
                    if saved_count > 0:
                        msg_parts.append(f"บันทึกไฟล์ภาพ {saved_count} รูปเรียบร้อยแล้ว")
                    if updated_count > 0:
                        msg_parts.append(f"อัปเดตข้อมูลตาราง {updated_count} แถวสำเร็จ")
                        
                    if msg_parts:
                        st.session_state["calib_maint_msg_success"] = f"✅ สมบูรณ์แล้ว! ({', '.join(msg_parts)})"
                    else:
                        st.session_state["calib_maint_msg_success"] = "✅ สมบูรณ์แล้ว!"
                        
                    st.session_state["calib_maint_uploader_ver"] = st.session_state.get("calib_maint_uploader_ver", 0) + 1
                    st.rerun()

        # 3. Library Backup & Export
        with subtab_backup:
            st.markdown("##### 📦 Export Entire Calibration Image Library")
            st.caption("Download a complete ZIP package containing all calibration images for backup, archiving, or offline bulk editing.")
            
            st.write(f"The archive currently contains **{len(df_all_imgs)}** images ({tot_mb:.2f} MB).")
            
            if st.button("📦 Generate Downloadable ZIP Archive", type="primary", key="calib_btn_gen_zip"):
                with st.spinner("Compressing image library into ZIP package..."):
                    zip_name, zip_bytes = calib_svc.create_images_zip_bundle()
                    st.download_button(
                        label=f"📥 Download `{zip_name}` ({len(zip_bytes)/1024/1024:.2f} MB)",
                        data=zip_bytes,
                        file_name=zip_name,
                        mime="application/zip",
                        type="primary",
                        use_container_width=True,
                        key="calib_btn_dl_zip"
                    )

# ---------------------------------------------------------------------------
# Quotation Verification & Certificate No. Mapping Tools
# ---------------------------------------------------------------------------
def render_quotation_verification_page():
    st.title("💼 Calibration Quotation Verification Tool")
    st.subheader("Cross-reference and verify vendor calibration quotation PDFs against your calibration certificate list")
    
    import backend.quotation_verification_service as q_svc
    try:
        importlib.reload(q_svc)
    except Exception:
        pass
        
    st.markdown("""
    <div style="background-color: rgba(0, 120, 212, 0.05); padding: 20px; border-radius: 10px; margin-bottom: 25px; border-left: 5px solid #0078d4;">
        <h4 style="margin-top: 0; color: #0078d4;">📋 Process Description</h4>
        <p style="color: var(--text-color); opacity: 0.9; font-size: 1rem; margin-bottom: 0;">
            Upload a vendor calibration quotation PDF (e.g., Thai Heart Calibration or other supplier quote) and cross-reference it 
            against the Serial Numbers and Component lists from your Calibration Certificate processing sheet (Step 2/3 or uploaded Excel).
            The system automatically identifies <b>Matched Items</b>, <b>Missing Items in Quotation</b>, and <b>Extra Items</b>.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Input Selection (Two Columns)
    col_pdf, col_calib = st.columns(2)
    
    with col_pdf:
        st.markdown("##### 📄 1. Vendor Quotation PDF")
        uploaded_quote_pdf = st.file_uploader(
            "Upload Quotation PDF (*.pdf)", 
            type=["pdf"], 
            key="quote_ver_pdf_uploader",
            help="Upload vendor quotation containing itemized S/N and component numbers"
        )
        
        quote_parsed_data = None
        if uploaded_quote_pdf:
            try:
                quote_bytes = uploaded_quote_pdf.getvalue()
                quote_parsed_data = q_svc.extract_quotation_data(quote_bytes)
                st.caption(
                    f"🏷️ **Quotation No:** `{quote_parsed_data['quotation_no']}` | "
                    f"📅 **Date:** `{quote_parsed_data['date']}` | "
                    f"📦 **Items Found:** `{quote_parsed_data['total_items']}`"
                )
            except Exception as q_err:
                st.error(f"Error parsing Quotation PDF: {q_err}")

    with col_calib:
        st.markdown("##### 📊 2. Calibration Certificate List")
        
        # Check active session state
        has_active_calib = (
            st.session_state.get("active_calib_df") is not None and not st.session_state.get("active_calib_df").empty
        ) or (st.session_state.get("calib_final_bytes") is not None)
        
        calib_source_options = []
        if has_active_calib:
            calib_source_options.append("⚡ Use Active Calibration Data from Step 1 / 2 / 3 (Auto-Detected)")
        calib_source_options.extend([
            "📂 Upload File (Excel *.xlsx, *.xls or Converted PDF *.pdf)"
        ])
        
        calib_src_choice = st.selectbox(
            "Select Calibration List Source",
            calib_source_options,
            key="quote_ver_calib_src_select"
        )
        
        calib_extracted_df = pd.DataFrame()
        if "Auto-Detected" in calib_src_choice:
            if st.session_state.get("calib_final_bytes") is not None:
                calib_extracted_df = q_svc.extract_calibration_list_data(st.session_state.calib_final_bytes)
            elif st.session_state.get("active_calib_df") is not None:
                calib_extracted_df = q_svc.extract_calibration_list_data(st.session_state.active_calib_df)
            st.caption(f"✅ Loaded **{len(calib_extracted_df)}** valid serial item(s) from current session.")
        else:
            uploaded_calib_file = st.file_uploader(
                "Upload Calibration Sheet or Converted PDF (*.xlsx, *.xls, *.pdf)",
                type=["xlsx", "xls", "pdf"],
                key="quote_ver_calib_file_uploader",
                help="Upload either an Excel file (*.xlsx, *.xls) or a converted Calibration Certificate PDF (*.pdf)"
            )
            if uploaded_calib_file:
                try:
                    calib_extracted_df = q_svc.extract_calibration_list_data(uploaded_calib_file.getvalue(), filename=uploaded_calib_file.name)
                    st.caption(f"✅ Loaded **{len(calib_extracted_df)}** item(s) from `{uploaded_calib_file.name}`.")
                except Exception as ce:
                    st.error(f"Error loading calibration file: {ce}")

    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
    
    # Action Button
    can_verify = (quote_parsed_data is not None) and (calib_extracted_df is not None and not calib_extracted_df.empty)
    
    col_btn, _ = st.columns([1.5, 2.5])
    with col_btn:
        run_verify = st.button(
            "⚡ Verify & Compare Quotation",
            type="primary",
            use_container_width=True,
            key="quote_ver_run_btn",
            disabled=(not can_verify)
        )
        
    if run_verify and can_verify:
        with st.spinner("Cross-referencing quotation serial numbers against calibration list..."):
            try:
                comp_result = q_svc.compare_quotation_and_calibration(quote_parsed_data, calib_extracted_df)
                st.session_state["quote_ver_last_result"] = comp_result
                st.success("🎉 **Quotation Verification Completed!** Results are displayed below.")
            except Exception as v_err:
                st.error(f"Error during verification: {v_err}")

    # Display Verification Results
    if "quote_ver_last_result" in st.session_state and st.session_state["quote_ver_last_result"]:
        res = st.session_state["quote_ver_last_result"]
        q_meta = res.get("quotation_meta", {})
        
        st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
        st.markdown("### 📊 Verification Summary Dashboard")
        
        # Meta Card
        with st.container(border=True):
            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.markdown(f"**Quotation No:** `{q_meta.get('quotation_no', 'N/A')}`")
            mc2.markdown(f"**Date:** `{q_meta.get('date', 'N/A')}`")
            mc3.markdown(f"**Customer:** `{q_meta.get('customer', 'N/A')}`")
            mc4.markdown(f"**Project Ref:** `{q_meta.get('project_ref', 'N/A')}`")
        
        # Metric KPIs
        extra_count = res.get("extra_in_quote_count", res.get("extra_in_quotation_count", 0))
        m_col1, m_col2, m_col3, m_col4, m_col5, m_col6 = st.columns(6)
        m_col1.metric("📋 Calibration Items", f"{res['total_calibration_items']}")
        m_col2.metric("📄 Quoted Items", f"{res['total_quotation_items']}")
        m_col3.metric("🟢 Matched", f"{res['matched_count']}")
        m_col4.metric("🟡 Missing in Quote", f"{res['missing_in_quotation_count']}")
        m_col5.metric("⚠️ Extra in Quote", f"{extra_count}")
        m_col6.metric("🎯 Match Rate", f"{res['match_rate_pct']}%")
        
        # Match Rate Progress Bar
        st.progress(min(1.0, res['match_rate_pct'] / 100.0), text=f"Overall Quotation Match Compliance: {res['match_rate_pct']}%")
        
        # Download Excel Report Button
        report_bytes = q_svc.generate_verification_report_excel(res)
        report_fn = f"Quotation_Verification_{q_meta.get('quotation_no', 'Report').replace('/', '_')}.xlsx"
        
        st.download_button(
            label=f"📥 Download Excel Verification Report ({report_fn})",
            data=report_bytes,
            file_name=report_fn,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            key="quote_ver_dl_report_btn"
        )
        
        st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
        
        # Detailed Item Breakdown Tabs
        df_comp = res["comparison_df"]
        
        tab_all, tab_matched, tab_missing, tab_extra, tab_raw = st.tabs([
            f"📋 All Items ({len(df_comp)})",
            f"🟢 Matched Items ({res['matched_count']})",
            f"🟡 Missing in Quotation ({res['missing_in_quotation_count']})",
            f"⚠️ Extra in Quotation ({extra_count})",
            "🔍 Raw Quotation Data"
        ])
        
        def render_styled_comp_table(sub_df):
            if sub_df.empty:
                st.info("No items in this category.")
                return
            display_cols = [
                "No.", "Status", "Serial No.", "List Comp No", "Quotation Comp No",
                "List Desc", "Quotation Desc", "List Dimensions", "Quotation Range / Specs", "Notes"
            ]
            cols_to_show = [c for c in display_cols if c in sub_df.columns]
            
            def highlight_status(row):
                status_v = str(row.get("Status", ""))
                # If matched -> Soft green background with crisp readable text
                if "Matched" in status_v and "Discrepancy" not in status_v and "Missing" not in status_v:
                    return ['background-color: rgba(40, 167, 69, 0.18); font-weight: 500;'] * len(row)
                # If not matched -> Soft yellow background with crisp readable text
                else:
                    return ['background-color: rgba(255, 193, 7, 0.30); font-weight: 500;'] * len(row)
                
            column_config = {
                "No.": st.column_config.NumberColumn("No.", width="small"),
                "Status": st.column_config.TextColumn("Status", width="medium"),
                "Serial No.": st.column_config.TextColumn("Serial No.", width="large"),
                "List Comp No": st.column_config.TextColumn("List Comp No", width="small"),
                "Quotation Comp No": st.column_config.TextColumn("Quotation Comp No", width="small"),
                "List Desc": st.column_config.TextColumn("List Description", width="medium"),
                "Quotation Desc": st.column_config.TextColumn("Quotation Description", width="medium"),
                "List Dimensions": st.column_config.TextColumn("List Dimensions", width="large"),
                "Quotation Range / Specs": st.column_config.TextColumn("Quotation Specs", width="medium"),
                "Notes": st.column_config.TextColumn("Verification Notes", width="large"),
            }
            
            styled = sub_df[cols_to_show].style.apply(highlight_status, axis=1)
            st.dataframe(styled, column_config=column_config, use_container_width=True, hide_index=True)
            
        with tab_all:
            st.markdown("##### 📋 Complete Verification Item Breakdown (Ordered by Calibration List)")
            search_item = st.text_input("🔍 Filter items by Serial No or Component No...", key="quote_ver_search_all")
            filtered_all = df_comp
            if search_item:
                filtered_all = df_comp[
                    df_comp["Serial No."].astype(str).str.contains(search_item, case=False, na=False) |
                    df_comp["List Comp No"].astype(str).str.contains(search_item, case=False, na=False) |
                    df_comp["Quotation Comp No"].astype(str).str.contains(search_item, case=False, na=False)
                ]
            render_styled_comp_table(filtered_all)
            
        with tab_matched:
            st.markdown("##### 🟢 Successfully Matched Items (Present in Both Quotation & Calibration List)")
            matched_df = df_comp[df_comp["Status"].str.contains("Matched", na=False) & ~df_comp["Status"].str.contains("Discrepancy", na=False)]
            render_styled_comp_table(matched_df)
            
        with tab_missing:
            st.markdown("##### 🟡 Missing / Discrepancy Items (Items Missing from Quotation or with Discrepancies)")
            missing_df = df_comp[df_comp["Status"].str.contains("Missing", na=False) | df_comp["Status"].str.contains("Discrepancy", na=False)]
            if not missing_df.empty:
                st.warning(f"⚠️ **Attention Required:** Found {len(missing_df)} unmatched item(s) or discrepancies. Please notify the vendor to revise the quotation.")
            render_styled_comp_table(missing_df)
            
        with tab_extra:
            st.markdown("##### ⚠️ Extra in Quotation (Quoted by Vendor but Not in Calibration List)")
            extra_df = df_comp[df_comp["Status"].str.contains("Extra", na=False)]
            if not extra_df.empty:
                st.info(f"ℹ️ Found {len(extra_df)} extra serial number(s) on the quotation.")
            render_styled_comp_table(extra_df)
            
        with tab_raw:
            st.markdown("##### 🔍 Extracted Quotation PDF Content")
            st.markdown("###### Structured Line Items Extracted from PDF:")
            st.dataframe(q_meta.get("items_df", pd.DataFrame()), use_container_width=True, hide_index=True)
            with st.expander("📄 View Full Raw Text Extracted from PDF", expanded=False):
                st.text(q_meta.get("raw_text", ""))

def render_certificate_mapping_page():
    st.title("🔢 Certificate No. Mapping Tool")
    st.subheader("Automated matching of scanned PDF Calibration Certificates with Excel (Step 2) Serial Numbers")

    st.markdown("""
    <div style="background-color: rgba(0, 120, 212, 0.05); padding: 20px; border-radius: 10px; margin-bottom: 25px; border-left: 5px solid #0078d4;">
        <h4 style="margin-top: 0; color: #0078d4;">📋 Instructions & Workflow</h4>
        <ol style="color: var(--text-color); opacity: 0.9; font-size: 0.95rem; margin-bottom: 0; padding-left: 20px;">
            <li>Upload the Excel file downloaded from Step 2 of <b>📜 Calibration Certificate Processing Tool</b> (e.g. <code>Final_56021_20260910.xlsx</code>).</li>
            <li>Upload one or multiple scanned Calibration Certificate PDF files.</li>
            <li>Click <b>⚡ Process & Map Certificate Numbers</b>.</li>
            <li>The system will OCR each PDF, extract <b>Serial No.</b> & <b>Certificate No.</b>, update the Excel cells dynamically, rename PDFs according to sequence prefix & Certificate No. (e.g. <code>001_PI-1407003_26.pdf</code>), and package the output into a downloadable ZIP.</li>
        </ol>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader("1. Upload Step 2 Excel File", type=["xlsx", "xlsm"], key="cert_mapping_excel")
    with col2:
        pdf_files = st.file_uploader("2. Upload Calibration PDF Certificates (Multiple)", type=["pdf"], accept_multiple_files=True, key="cert_mapping_pdfs")

    if st.button("⚡ Process & Map Certificate Numbers", type="primary", use_container_width=True):
        if not excel_file:
            st.error("⚠️ Please upload the Step 2 Excel file first.")
            return
        if not pdf_files:
            st.error("⚠️ Please upload at least one Calibration PDF Certificate.")
            return

        with st.spinner("Processing PDFs and mapping Certificate Numbers to Excel..."):
            try:
                from backend.cert_no_mapping_service import process_certificate_no_mapping

                pdf_list = [
                    {"filename": pdf.name, "bytes": pdf.getvalue()}
                    for pdf in pdf_files
                ]

                result = process_certificate_no_mapping(
                    excel_file.getvalue(),
                    excel_file.name,
                    pdf_list
                )

                st.success("✅ Certificate No. Mapping completed successfully!")

                # Key Performance Metrics
                m_col1, m_col2, m_col3 = st.columns(3)
                m_col1.metric("Total PDFs Processed", result["total_pdfs"])
                m_col2.metric("Matched & Renamed PDFs", result["matched_pdfs"])
                m_col3.metric("Excel Rows Updated", result["rows_updated"])

                st.markdown("### 📊 Extraction & Mapping Summary")
                st.dataframe(result["summary_df"], use_container_width=True)

                st.markdown("### 📥 Download Processed Results")
                dl_col1, dl_col2 = st.columns(2)

                with dl_col1:
                    st.download_button(
                        label="📄 Download Updated Excel File",
                        data=result["updated_excel_bytes"],
                        file_name=result["updated_excel_filename"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

                with dl_col2:
                    st.download_button(
                        label="📦 Download Complete ZIP Package (Excel + PDFs)",
                        data=result["zip_bytes"],
                        file_name=result["zip_filename"],
                        mime="application/zip",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"❌ An error occurred during mapping: {e}")



# ---------------------------------------------------------------------------
# Placeholder for coming soon features
# ---------------------------------------------------------------------------
def render_fault_assistance_suite_page():
    tab1, tab2 = st.tabs(["⚙️ Fault Assistance Processing Tool", "✏️ Fault HTML Maintenance"])
    with tab1:
        render_fault_assistance_page()
    with tab2:
        render_fault_html_maintenance_page()

def render_fault_assistance_page():
    st.title("🔧 Fault Assistance Processing Tool")
    st.subheader("Process and package fault assistance HTML files and assets")
    
    st.info("💡 **Instruction:** Upload your Excel file and specify the project details to copy HTML files from the **Fault assistance** folder, automatically update the project name inside the HTML files, and download the results as a .Zip file and status Excel report.")
    
    col1, col2 = st.columns([3, 2])
    with col1:
        excel_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls", "xlsm"], key="fault_excel_uploader")
    with col2:
        customer_name = st.text_input("Customer Name", placeholder="e.g., 56061_Henkel Maribor SI_VI5X", key="fault_customer_name")
        
    process_btn = st.button("🚀 Process & Package", type="primary", use_container_width=True)
    
    if process_btn:
        if not excel_file:
            st.error("Please upload the Excel file first.")
            return
        if not customer_name.strip():
            st.error("Please specify the Customer Name.")
            return
            
        with st.spinner("Copying and updating HTML files..."):
            import backend.fault_assistance_service
            from backend.fault_assistance_service import process_fault_assistance_files
            
            success, msg, status_df, zip_bytes, excel_status_bytes, project_number, final_zip_bytes = process_fault_assistance_files(
                excel_file.getvalue(), customer_name
            )
            
            if success:
                st.success(msg)
                
                # Action Buttons
                st.markdown("### 📥 Download Results")
                dl_col1, dl_col2, dl_col3 = st.columns(3)
                with dl_col1:
                    st.download_button(
                        label="📦 Download Final Package (.ZIP)",
                        data=final_zip_bytes,
                        file_name=f"Final_{project_number}.zip",
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )
                with dl_col2:
                    st.download_button(
                        label="⬇️ Download Output Files (.ZIP)",
                        data=zip_bytes,
                        file_name=f"Fault_Assistance_{project_number}.zip",
                        mime="application/zip",
                        use_container_width=True
                    )
                with dl_col3:
                    st.download_button(
                        label="⬇️ Download Status Excel",
                        data=excel_status_bytes,
                        file_name=f"Status_Excel_{project_number}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                
                # Show status table
                st.markdown("### 📊 Status Preview Table")
                
                # Highlight rows based on status (yellow for NOT FOUND, red for FAILED)
                def highlight_status_rows(row):
                    status = row["Status"]
                    if status == "NOT FOUND":
                        return ["background-color: rgba(255, 193, 7, 0.25);"] * len(row)
                    elif "FAILED" in str(status):
                        return ["background-color: rgba(244, 67, 54, 0.25);"] * len(row)
                    return [""] * len(row)
                
                styled_df = status_df.style.apply(highlight_status_rows, axis=1)
                st.dataframe(styled_df, use_container_width=True)
            else:
                st.error(msg)

def render_fault_html_maintenance_page():
    st.title("🔧 Fault Assistance HTML Maintenance")
    st.subheader("Manage, preview, and visually edit fault manuals.")
    
    import backend.fault_maintenance_service as maintenance_service
    src_folder = "Fault assistance"
    
    # 1. Action Expanders
    col_act1, col_act2 = st.columns(2)
    with col_act1:
        with st.expander("➕ Create New Manual"):
            new_code = st.text_input("New Fault Code (Fehlern)", placeholder="e.g. 21000216", key="new_fault_code")
            new_text = st.text_input("New Alarm Message (Alarmtext)", placeholder="e.g. Protective guarding 17 is open", key="new_fault_text")
            create_btn = st.button("Create File", use_container_width=True)
            if create_btn:
                if not new_code.strip() or not new_text.strip():
                    st.error("Please specify both code and text.")
                else:
                    success, msg = maintenance_service.create_new_fault_file(src_folder, new_code, new_text)
                    if success:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)
                        
    with col_act2:
        with st.expander("❌ Delete Selected Manual"):
            st.warning("This will permanently delete the currently selected manual file.")
            delete_confirm = st.checkbox("I confirm that I want to delete the selected file", key="delete_confirm")
            
    # 2. Load and Select Files
    files_list = maintenance_service.list_html_files(src_folder)
    if not files_list:
        st.warning("No HTML manuals found in 'Fault assistance' directory.")
        return
        
    display_options = [f["display"] for f in files_list]
    selected_display = st.selectbox("Select Fault Manual to Edit/Preview", display_options, key="select_fault_manual")
    
    selected_file_idx = display_options.index(selected_display)
    selected_file_info = files_list[selected_file_idx]
    selected_filename = selected_file_info["filename"]
    selected_file_path = os.path.join(src_folder, selected_filename)
    
    # Handle Delete action
    if col_act2.button("Delete File", use_container_width=True, disabled=not delete_confirm):
        success, msg = maintenance_service.delete_fault_file(src_folder, selected_filename)
        if success:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)
            
    # 3. Read and Parse HTML
    parsed_data = maintenance_service.read_and_parse_html(selected_file_path)
    if not parsed_data["success"]:
        st.error(f"Error parsing file: {parsed_data.get('error')}")
        return
        
    # Initialize session state cache for editing
    if (st.session_state.get("current_edit_file") != selected_file_path) or ("edit_blocks" not in st.session_state):
        import time
        st.session_state["current_edit_file"] = selected_file_path
        heading, blocks = maintenance_service.parse_st4_to_blocks(parsed_data["raw_html"])
        # Add a unique stable "id" to each block to prevent key clashing during move operations
        for i, b in enumerate(blocks):
            b["id"] = f"b_{i}_{int(time.time() * 1000)}"
        st.session_state["edit_heading_val"] = heading
        st.session_state["edit_blocks"] = blocks
        
    # 4. Preview and Edit Layout Columns
    col_preview, col_editor = st.columns([1, 1])
    
    with col_preview:
        st.markdown("### 👁️ Live Preview")
        preview_html = maintenance_service.get_html_preview_with_styles(selected_file_path, src_folder)
        st.components.v1.html(preview_html, height=600, scrolling=True)
        
    with col_editor:
        st.markdown("### ✏️ Document Editor")
        edit_mode = st.radio("Editor Mode", ["Visual Form", "Raw HTML"], horizontal=True, key="editor_mode_radio")
        
        if edit_mode == "Visual Form":
            st.markdown('<div id="editor-marker" style="display:none;"></div>', unsafe_allow_html=True)
            st.info("💡 **Instructions:** Edit the fields below. Use ▲/▼ buttons to reorder blocks. Use Remove Block to delete blocks. Use dropdown at the bottom to add new blocks.")
            col_h1, col_h2 = st.columns([4, 1])
            with col_h1:
                heading = st.text_input("Heading / Alarm Message", value=st.session_state["edit_heading_val"], key=f"edit_heading_{selected_filename}")
            with col_h2:
                st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
                if st.button("🔄 Reload File", use_container_width=True, help="Discard unsaved changes and reload this file from disk."):
                    if "current_edit_file" in st.session_state:
                        del st.session_state["current_edit_file"]
                    st.rerun()
            
            blocks = st.session_state["edit_blocks"]
            
            # Synchronize session state widget values back to blocks dynamically on every rerun
            for b in blocks:
                b_id = b.get("id")
                if not b_id:
                    continue
                if b["type"] == "SubHeading":
                    if f"subheading_{selected_filename}_{b_id}" in st.session_state:
                        b["text"] = st.session_state[f"subheading_{selected_filename}_{b_id}"]
                elif b["type"] == "Paragraph":
                    if f"paragraph_{selected_filename}_{b_id}" in st.session_state:
                        b["text"] = st.session_state[f"paragraph_{selected_filename}_{b_id}"]
                    if f"paragraph_icon_{selected_filename}_{b_id}" in st.session_state:
                        b["icon"] = st.session_state[f"paragraph_icon_{selected_filename}_{b_id}"]
                elif b["type"] == "List":
                    if f"list_{selected_filename}_{b_id}" in st.session_state:
                        t_val = st.session_state[f"list_{selected_filename}_{b_id}"]
                        b["items"] = [line.strip() for line in t_val.split("\n") if line.strip()]
                elif b["type"] == "Table":
                    for r_i in range(len(b.get("rows", []))):
                        if f"table_cause_{selected_filename}_{b_id}_{r_i}" in st.session_state:
                            b["rows"][r_i]["cause"] = st.session_state[f"table_cause_{selected_filename}_{b_id}_{r_i}"]
                        if f"table_elim_{selected_filename}_{b_id}_{r_i}" in st.session_state:
                            b["rows"][r_i]["elimination"] = st.session_state[f"table_elim_{selected_filename}_{b_id}_{r_i}"]
                elif b["type"] == "Safety":
                    if f"safety_type_{selected_filename}_{b_id}" in st.session_state:
                        b["warn_type"] = st.session_state[f"safety_type_{selected_filename}_{b_id}"]
                    if f"safety_word_{selected_filename}_{b_id}" in st.session_state:
                        b["word"] = st.session_state[f"safety_word_{selected_filename}_{b_id}"]
                    if f"safety_cause_{selected_filename}_{b_id}" in st.session_state:
                        b["cause"] = st.session_state[f"safety_cause_{selected_filename}_{b_id}"]
                    if f"safety_icon_{selected_filename}_{b_id}" in st.session_state:
                        b["icon"] = st.session_state[f"safety_icon_{selected_filename}_{b_id}"]
                    if f"safety_cons_{selected_filename}_{b_id}" in st.session_state:
                        s_cons = st.session_state[f"safety_cons_{selected_filename}_{b_id}"]
                        b["consequences"] = [line.strip() for line in s_cons.split("\n") if line.strip()]
            st.markdown("#### Table Content / Document Blocks")
            
            new_blocks = []
            for idx, block in enumerate(blocks):
                # Auto-heal block IDs if missing from older session state cache
                if "id" not in block:
                    import time
                    block["id"] = f"b_{idx}_{int(time.time() * 1000)}"
                b_type = block["type"]
                b_id = block["id"]
                
                remove_b = False
                move_up = False
                move_down = False
                
                with st.container(border=True):
                    col_b1, col_up, col_down, col_b2 = st.columns([5, 1, 1, 2])
                    with col_b1:
                        if b_type == "Paragraph":
                            col_title, col_icon_sel = st.columns([2, 3])
                            with col_title:
                                st.markdown(f"**Block {idx + 1}: Paragraph**")
                            with col_icon_sel:
                                icon_options = {
                                    "": "None (Normal)",
                                    "action": "👉 Action (Hand)",
                                    "touch": "➔ Touch (Arrow)"
                                }
                                current_icon = block.get("icon", "")
                                icon_keys = list(icon_options.keys())
                                default_idx = icon_keys.index(current_icon) if current_icon in icon_keys else 0
                                
                                p_icon = st.selectbox(
                                    "Icon / Symbol",
                                    options=icon_keys,
                                    format_func=lambda x: icon_options[x],
                                    index=default_idx,
                                    key=f"paragraph_icon_{selected_filename}_{b_id}",
                                    label_visibility="collapsed"
                                )
                        else:
                            st.markdown(f"**Block {idx + 1}: {b_type}**")
                    with col_up:
                        move_up = st.button("▲", key=f"move_up_{selected_filename}_{b_id}", disabled=(idx == 0), use_container_width=True)
                    with col_down:
                        move_down = st.button("▼", key=f"move_down_{selected_filename}_{b_id}", disabled=(idx == len(blocks) - 1), use_container_width=True)
                    with col_b2:
                        remove_b = st.checkbox("Remove Block", key=f"remove_block_{selected_filename}_{b_id}")
                        
                    # Sync and Swap Logic for Move Up
                    if move_up:
                        synced_blocks = []
                        for i, b in enumerate(blocks):
                            b_t = b["type"]
                            if b_t == "SubHeading":
                                t_val = st.session_state.get(f"subheading_{selected_filename}_{b['id']}", b.get("text", ""))
                                synced_blocks.append({"id": b["id"], "type": "SubHeading", "text": t_val})
                            elif b_t == "Paragraph":
                                t_val = st.session_state.get(f"paragraph_{selected_filename}_{b['id']}", b.get("text", ""))
                                p_icon = st.session_state.get(f"paragraph_icon_{selected_filename}_{b['id']}", b.get("icon", ""))
                                synced_blocks.append({"id": b["id"], "type": "Paragraph", "text": t_val, "icon": p_icon})
                            elif b_t == "List":
                                t_val = st.session_state.get(f"list_{selected_filename}_{b['id']}", "\n".join(b.get("items", [])))
                                itms = [line.strip() for line in t_val.split("\n") if line.strip()]
                                synced_blocks.append({"id": b["id"], "type": "List", "items": itms})
                            elif b_t == "Table":
                                t_rows_synced = []
                                for r_i in range(len(b.get("rows", []))):
                                    if not st.session_state.get(f"table_remove_{selected_filename}_{b['id']}_{r_i}", False):
                                        c_v = st.session_state.get(f"table_cause_{selected_filename}_{b['id']}_{r_i}", b["rows"][r_i].get("cause", ""))
                                        e_v = st.session_state.get(f"table_elim_{selected_filename}_{b['id']}_{r_i}", b["rows"][r_i].get("elimination", ""))
                                        t_rows_synced.append({"cause": c_v, "elimination": e_v})
                                synced_blocks.append({"id": b["id"], "type": "Table", "rows": t_rows_synced})
                            elif b_t == "Safety":
                                w_t = st.session_state.get(f"safety_type_{selected_filename}_{b['id']}", b.get("warn_type", "caution"))
                                w_w = st.session_state.get(f"safety_word_{selected_filename}_{b['id']}", b.get("word", "CAUTION"))
                                s_c = st.session_state.get(f"safety_cause_{selected_filename}_{b['id']}", b.get("cause", ""))
                                s_cons = st.session_state.get(f"safety_cons_{selected_filename}_{b['id']}", "\n".join(b.get("consequences", [])))
                                cons_l = [line.strip() for line in s_cons.split("\n") if line.strip()]
                                icon_v = st.session_state.get(f"safety_icon_{selected_filename}_{b['id']}", "general_general")
                                synced_blocks.append({
                                    "id": b["id"],
                                    "type": "Safety",
                                    "warn_type": w_t,
                                    "word": w_w,
                                    "cause": s_c,
                                    "consequences": cons_l,
                                    "icon": icon_v
                                })
                            elif b_t == "Image":
                                s_val = st.session_state.get(f"image_src_{selected_filename}_{b['id']}", b.get("src", ""))
                                synced_blocks.append({"id": b["id"], "type": "Image", "src": s_val})
                        synced_blocks[idx], synced_blocks[idx-1] = synced_blocks[idx-1], synced_blocks[idx]
                        st.session_state["edit_blocks"] = synced_blocks
                        st.rerun()
                        
                    # Sync and Swap Logic for Move Down
                    if move_down:
                        synced_blocks = []
                        for i, b in enumerate(blocks):
                            b_t = b["type"]
                            if b_t == "SubHeading":
                                t_val = st.session_state.get(f"subheading_{selected_filename}_{b['id']}", b.get("text", ""))
                                synced_blocks.append({"id": b["id"], "type": "SubHeading", "text": t_val})
                            elif b_t == "Paragraph":
                                t_val = st.session_state.get(f"paragraph_{selected_filename}_{b['id']}", b.get("text", ""))
                                p_icon = st.session_state.get(f"paragraph_icon_{selected_filename}_{b['id']}", b.get("icon", ""))
                                synced_blocks.append({"id": b["id"], "type": "Paragraph", "text": t_val, "icon": p_icon})
                            elif b_t == "List":
                                t_val = st.session_state.get(f"list_{selected_filename}_{b['id']}", "\n".join(b.get("items", [])))
                                itms = [line.strip() for line in t_val.split("\n") if line.strip()]
                                synced_blocks.append({"id": b["id"], "type": "List", "items": itms})
                            elif b_t == "Table":
                                t_rows_synced = []
                                for r_i in range(len(b.get("rows", []))):
                                    if not st.session_state.get(f"table_remove_{selected_filename}_{b['id']}_{r_i}", False):
                                        c_v = st.session_state.get(f"table_cause_{selected_filename}_{b['id']}_{r_i}", b["rows"][r_i].get("cause", ""))
                                        e_v = st.session_state.get(f"table_elim_{selected_filename}_{b['id']}_{r_i}", b["rows"][r_i].get("elimination", ""))
                                        t_rows_synced.append({"cause": c_v, "elimination": e_v})
                                synced_blocks.append({"id": b["id"], "type": "Table", "rows": t_rows_synced})
                            elif b_t == "Safety":
                                w_t = st.session_state.get(f"safety_type_{selected_filename}_{b['id']}", b.get("warn_type", "caution"))
                                w_w = st.session_state.get(f"safety_word_{selected_filename}_{b['id']}", b.get("word", "CAUTION"))
                                s_c = st.session_state.get(f"safety_cause_{selected_filename}_{b['id']}", b.get("cause", ""))
                                s_cons = st.session_state.get(f"safety_cons_{selected_filename}_{b['id']}", "\n".join(b.get("consequences", [])))
                                cons_l = [line.strip() for line in s_cons.split("\n") if line.strip()]
                                icon_v = st.session_state.get(f"safety_icon_{selected_filename}_{b['id']}", "general_general")
                                synced_blocks.append({
                                    "id": b["id"],
                                    "type": "Safety",
                                    "warn_type": w_t,
                                    "word": w_w,
                                    "cause": s_c,
                                    "consequences": cons_l,
                                    "icon": icon_v
                                })
                            elif b_t == "Image":
                                s_val = st.session_state.get(f"image_src_{selected_filename}_{b['id']}", b.get("src", ""))
                                synced_blocks.append({"id": b["id"], "type": "Image", "src": s_val})
                            elif b_t == "Image":
                                s_val = st.session_state.get(f"image_src_{selected_filename}_{b['id']}", b.get("src", ""))
                                synced_blocks.append({"id": b["id"], "type": "Image", "src": s_val})
                        synced_blocks[idx], synced_blocks[idx+1] = synced_blocks[idx+1], synced_blocks[idx]
                        st.session_state["edit_blocks"] = synced_blocks
                        st.rerun()

                    if not remove_b:
                        if b_type == "SubHeading":
                            text_val = st.text_input("SubHeading Text", value=block.get("text", ""), key=f"subheading_{selected_filename}_{b_id}")
                            new_blocks.append({"id": b_id, "type": "SubHeading", "text": text_val})
                        elif b_type == "Paragraph":
                            text_val = st.text_area("Paragraph Text", value=block.get("text", ""), key=f"paragraph_{selected_filename}_{b_id}", height=100)
                            p_icon = st.session_state.get(f"paragraph_icon_{selected_filename}_{b_id}", block.get("icon", ""))
                            new_blocks.append({"id": b_id, "type": "Paragraph", "text": text_val, "icon": p_icon})
                        elif b_type == "List":
                            joined_items = "\n".join(block.get("items", []))
                            text_val = st.text_area("List Items (one per line)", value=joined_items, key=f"list_{selected_filename}_{b_id}")
                            items_list = [line.strip() for line in text_val.split("\n") if line.strip()]
                            new_blocks.append({"id": b_id, "type": "List", "items": items_list})
                        elif b_type == "Table":
                            t_rows = block.get("rows", [])
                            new_t_rows = []
                            
                            for r_idx, row in enumerate(t_rows):
                                st.markdown(f"**Table Row {r_idx + 1}**")
                                col_r1, col_r2 = st.columns(2)
                                with col_r1:
                                    c_text = st.text_area(f"Cause {r_idx + 1}", value=row.get("cause", ""), key=f"table_cause_{selected_filename}_{b_id}_{r_idx}")
                                with col_r2:
                                    e_text = st.text_area(f"Elimination {r_idx + 1}", value=row.get("elimination", ""), key=f"table_elim_{selected_filename}_{b_id}_{r_idx}")
                                
                                r_remove = st.checkbox(f"Remove Row {r_idx + 1}", key=f"table_remove_{selected_filename}_{b_id}_{r_idx}")
                                if not r_remove:
                                    new_t_rows.append({"cause": c_text, "elimination": e_text})
                                st.markdown("---")
                                
                            # Add Table Row option inside Table block
                            add_table_row = st.button("➕ Add Table Row", key=f"add_table_row_btn_{b_id}")
                            if add_table_row:
                                # Sync current rows
                                synced_rows = []
                                for i in range(len(t_rows)):
                                    c_v = st.session_state.get(f"table_cause_{selected_filename}_{b_id}_{i}", "")
                                    e_v = st.session_state.get(f"table_elim_{selected_filename}_{b_id}_{i}", "")
                                    synced_rows.append({"cause": c_v, "elimination": e_v})
                                synced_rows.append({"cause": "", "elimination": ""})
                                block["rows"] = synced_rows
                                st.session_state["edit_blocks"] = blocks
                                st.rerun()
                                
                            new_blocks.append({"id": b_id, "type": "Table", "rows": new_t_rows})
                        elif b_type == "Safety":
                            st.markdown(f"⚠️ **Safety Warning Box ({block.get('warn_type', 'caution').upper()})**")
                            col_s1, col_s2 = st.columns(2)
                            with col_s1:
                                w_t_val = st.selectbox("Warning Type", ["caution", "warning", "danger", "notice"], index=["caution", "warning", "danger", "notice"].index(block.get("warn_type", "caution")), key=f"safety_type_{selected_filename}_{b_id}")
                                word_val = st.text_input("Warning Word", value=block.get("word", "CAUTION"), key=f"safety_word_{selected_filename}_{b_id}")
                            with col_s2:
                                cause_val = st.text_input("Safety Cause", value=block.get("cause", ""), key=f"safety_cause_{selected_filename}_{b_id}")
                                
                                # Safety Icon Selector
                                available_icons = ["general_general", "heat_heat", "Maschinenschaden_Maschinenschaden", "notice_notice"]
                                default_icon = block.get("icon", "general_general")
                                if default_icon not in available_icons:
                                    default_icon = "general_general"
                                icon_val = st.selectbox("Safety Icon", available_icons, index=available_icons.index(default_icon), key=f"safety_icon_{selected_filename}_{b_id}")
                                
                                joined_cons = "\n".join(block.get("consequences", []))
                                cons_val = st.text_area("Consequences (one per line)", value=joined_cons, key=f"safety_cons_{selected_filename}_{b_id}")
                                cons_list = [line.strip() for line in cons_val.split("\n") if line.strip()]
                            
                            new_blocks.append({
                                "id": b_id,
                                "type": "Safety",
                                "warn_type": w_t_val,
                                "word": word_val,
                                "cause": cause_val,
                                "consequences": cons_list,
                                "icon": icon_val
                            })
                        elif b_type == "Image":
                            img_src = block.get("src", "")
                            new_src = st.text_input("Image Source Path", value=img_src, key=f"image_src_{selected_filename}_{b_id}")
                            
                            # Render image preview safely
                            if os.path.exists(new_src):
                                st.image(new_src, caption="Image Preview", use_container_width=True)
                            else:
                                fallback_src = os.path.join(os.path.dirname(selected_file_path), new_src)
                                if not os.path.exists(fallback_src) and os.path.exists(os.path.basename(new_src)):
                                    fallback_src = os.path.basename(new_src)
                                
                                if os.path.exists(fallback_src):
                                    st.image(fallback_src, caption="Image Preview", use_container_width=True)
                                else:
                                    st.warning(f"Image file not found: {new_src}")
                            new_blocks.append({"id": b_id, "type": "Image", "src": new_src})
                            
            # Add New Block option
            st.markdown("#### ➕ Add New Block")
            col_add1, col_add2 = st.columns([2, 1])
            with col_add1:
                new_block_type = st.selectbox("Block Type", ["Paragraph", "SubHeading", "List", "Table", "Safety", "Image"], key="new_block_type_select")
            with col_add2:
                add_block_btn = st.button("Add Block", use_container_width=True)
                
            if add_block_btn:
                # Sync current state of existing blocks
                synced_blocks = []
                for i, b in enumerate(blocks):
                    if not st.session_state.get(f"remove_block_{selected_filename}_{b['id']}", False):
                        if b["type"] == "SubHeading":
                            t_val = st.session_state.get(f"subheading_{selected_filename}_{b['id']}", "")
                            synced_blocks.append({"id": b["id"], "type": "SubHeading", "text": t_val})
                        elif b["type"] == "Paragraph":
                            t_val = st.session_state.get(f"paragraph_{selected_filename}_{b['id']}", "")
                            p_icon = st.session_state.get(f"paragraph_icon_{selected_filename}_{b['id']}", b.get("icon", ""))
                            synced_blocks.append({"id": b["id"], "type": "Paragraph", "text": t_val, "icon": p_icon})
                        elif b["type"] == "List":
                            t_val = st.session_state.get(f"list_{selected_filename}_{b['id']}", "")
                            itms = [line.strip() for line in t_val.split("\n") if line.strip()]
                            synced_blocks.append({"id": b["id"], "type": "List", "items": itms})
                        elif b["type"] == "Table":
                            t_rows_synced = []
                            for r_i in range(len(b["rows"])):
                                if not st.session_state.get(f"table_remove_{selected_filename}_{b['id']}_{r_i}", False):
                                    c_v = st.session_state.get(f"table_cause_{selected_filename}_{b['id']}_{r_i}", "")
                                    e_v = st.session_state.get(f"table_elim_{selected_filename}_{b['id']}_{r_i}", "")
                                    t_rows_synced.append({"cause": c_v, "elimination": e_v})
                            synced_blocks.append({"id": b["id"], "type": "Table", "rows": t_rows_synced})
                        elif b["type"] == "Safety":
                            w_t = st.session_state.get(f"safety_type_{selected_filename}_{b['id']}", "caution")
                            w_w = st.session_state.get(f"safety_word_{selected_filename}_{b['id']}", "CAUTION")
                            s_c = st.session_state.get(f"safety_cause_{selected_filename}_{b['id']}", "")
                            s_cons = st.session_state.get(f"safety_cons_{selected_filename}_{b['id']}", "")
                            cons_l = [line.strip() for line in s_cons.split("\n") if line.strip()]
                            icon_v = st.session_state.get(f"safety_icon_{selected_filename}_{b['id']}", "general_general")
                            synced_blocks.append({
                                "id": b["id"],
                                "type": "Safety",
                                "warn_type": w_t,
                                "word": w_w,
                                "cause": s_c,
                                "consequences": cons_l,
                                "icon": icon_v
                            })
                        elif b["type"] == "Image":
                            s_val = st.session_state.get(f"image_src_{selected_filename}_{b['id']}", b.get("src", ""))
                            synced_blocks.append({"id": b["id"], "type": "Image", "src": s_val})
                
                import time
                # Append the new block
                if new_block_type == "SubHeading":
                    synced_blocks.append({"id": f"b_new_{int(time.time() * 1000)}", "type": "SubHeading", "text": ""})
                elif new_block_type == "Paragraph":
                    synced_blocks.append({"id": f"b_new_{int(time.time() * 1000)}", "type": "Paragraph", "text": "", "icon": ""})

                elif new_block_type == "List":
                    synced_blocks.append({"id": f"b_new_{int(time.time() * 1000)}", "type": "List", "items": []})
                elif new_block_type == "Table":
                    synced_blocks.append({"id": f"b_new_{int(time.time() * 1000)}", "type": "Table", "rows": [{"cause": "", "elimination": ""}]})
                elif new_block_type == "Safety":
                    synced_blocks.append({
                        "id": f"b_new_{int(time.time() * 1000)}",
                        "type": "Safety",
                        "warn_type": "caution",
                        "word": "CAUTION",
                        "cause": "",
                        "consequences": [],
                        "icon": "general_general"
                    })
                elif new_block_type == "Image":
                    synced_blocks.append({"id": f"b_new_{int(time.time() * 1000)}", "type": "Image", "src": "My Picture/logo IWK.png"})
                    
                st.session_state["edit_blocks"] = synced_blocks
                st.rerun()
                
            # Save button
            save_btn = st.button("💾 Save Changes (Form)", type="primary", use_container_width=True)
            if save_btn:
                # Gather final blocks from synced inputs
                final_blocks = []
                for i, b in enumerate(blocks):
                    if not st.session_state.get(f"remove_block_{selected_filename}_{b['id']}", False):
                        if b["type"] == "SubHeading":
                            t_val = st.session_state.get(f"subheading_{selected_filename}_{b['id']}", "")
                            final_blocks.append({"type": "SubHeading", "text": t_val})
                        elif b["type"] == "Paragraph":
                            t_val = st.session_state.get(f"paragraph_{selected_filename}_{b['id']}", "")
                            p_icon = st.session_state.get(f"paragraph_icon_{selected_filename}_{b['id']}", b.get("icon", ""))
                            final_blocks.append({"type": "Paragraph", "text": t_val, "icon": p_icon})
                        elif b["type"] == "List":
                            t_val = st.session_state.get(f"list_{selected_filename}_{b['id']}", "")
                            itms = [line.strip() for line in t_val.split("\n") if line.strip()]
                            final_blocks.append({"type": "List", "items": itms})
                        elif b["type"] == "Table":
                            t_rows_synced = []
                            for r_i in range(len(b["rows"])):
                                if not st.session_state.get(f"table_remove_{selected_filename}_{b['id']}_{r_i}", False):
                                    c_v = st.session_state.get(f"table_cause_{selected_filename}_{b['id']}_{r_i}", "")
                                    e_v = st.session_state.get(f"table_elim_{selected_filename}_{b['id']}_{r_i}", "")
                                    t_rows_synced.append({"cause": c_v, "elimination": e_v})
                            final_blocks.append({"type": "Table", "rows": t_rows_synced})
                        elif b["type"] == "Safety":
                            w_t = st.session_state.get(f"safety_type_{selected_filename}_{b['id']}", "caution")
                            w_w = st.session_state.get(f"safety_word_{selected_filename}_{b['id']}", "CAUTION")
                            s_c = st.session_state.get(f"safety_cause_{selected_filename}_{b['id']}", "")
                            s_cons = st.session_state.get(f"safety_cons_{selected_filename}_{b['id']}", "")
                            cons_l = [line.strip() for line in s_cons.split("\n") if line.strip()]
                            icon_v = st.session_state.get(f"safety_icon_{selected_filename}_{b['id']}", "general_general")
                            final_blocks.append({
                                "type": "Safety",
                                "warn_type": w_t,
                                "word": w_w,
                                "cause": s_c,
                                "consequences": cons_l,
                                "icon": icon_v
                            })
                        elif b["type"] == "Image":
                            s_val = st.session_state.get(f"image_src_{selected_filename}_{b['id']}", "")
                            final_blocks.append({"type": "Image", "src": s_val})
                            
                success, msg = maintenance_service.save_html_blocks(selected_file_path, heading, final_blocks)
                if success:
                    st.success(msg)
                    del st.session_state["current_edit_file"]
                    st.rerun()
                else:
                    st.error(msg)
                    
            pass
                    
        else:  # Raw HTML mode
            # Render HTML search bar widget above the textarea with class-based selectors (immune to ID stripping)
            st.markdown('''
            <div class="html-search-bar" style="display: flex; gap: 8px; align-items: center; margin-bottom: 8px; background-color: #f8fafc; padding: 6px 12px; border: 1px solid #e2e8f0; border-radius: 8px;">
                <span style="font-size: 14px; color: #475569; font-weight: 600;">🔍 Find:</span>
                <input type="text" class="html-search-input" placeholder="Type query to search..." style="flex-grow: 1; padding: 6px 10px; font-size: 14px; border: 1px solid #cbd5e1; border-radius: 6px; outline: none; transition: all 0.15s; background-color: white; color: #1e293b;" onfocus="this.style.borderColor='#3b82f6';this.style.boxShadow='0 0 0 1px #3b82f6';" onblur="this.style.borderColor='#cbd5e1';this.style.boxShadow='none';" />
                <button class="html-search-prev" type="button" style="padding: 6px 12px; font-size: 13px; font-weight: 600; border: 1px solid #cbd5e1; border-radius: 6px; background-color: white; cursor: pointer; color: #475569; transition: all 0.1s;" onmouseover="this.style.backgroundColor='#f1f5f9'" onmouseout="this.style.backgroundColor='white'">◀ Prev</button>
                <button class="html-search-next" type="button" style="padding: 6px 12px; font-size: 13px; font-weight: 600; border: 1px solid #cbd5e1; border-radius: 6px; background-color: white; cursor: pointer; color: #475569; transition: all 0.1s;" onmouseover="this.style.backgroundColor='#f1f5f9'" onmouseout="this.style.backgroundColor='white'">Next ▶</button>
                <span class="html-search-status" style="font-size: 13px; font-weight: 600; color: #64748b; min-width: 60px; text-align: center;">0/0</span>
                <span class="html-search-debug" style="font-size: 12px; color: #10b981; font-weight: 600; margin-left: 8px;"></span>
            </div>
            <div class="html-search-context" style="font-size: 13px; color: #475569; margin-top: -4px; margin-bottom: 8px; background-color: #f1f5f9; padding: 6px 12px; border: 1px solid #e2e8f0; border-radius: 8px; display: none; font-family: monospace; white-space: pre-wrap; word-break: break-all;"></div>
            ''', unsafe_allow_html=True)
            
            raw_html = st.text_area("Raw HTML Code", value=parsed_data["raw_html"], height=450, key=f"edit_raw_html_{selected_filename}")
            save_raw_btn = st.button("💾 Save Changes (Raw HTML)", type="primary", use_container_width=True)
            if save_raw_btn:
                success, msg = maintenance_service.save_raw_html_file(selected_file_path, raw_html)
                if success:
                    st.success(msg)
                    st.rerun()
                else:
                    st.error(msg)
            
            st.markdown("---")
            with st.expander("ℹ️ Reference: Bold and Background Formatting", expanded=True):
                st.image("My Picture/Bold and background.png", caption="Bold and Background formatting reference guide")

        # Inject parent window script controller (loaded in both modes)
        st.components.v1.html(r'''
        <script>
        (function() {
            const doc = window.parent.document;
            
            // Force recreation of script tag to prevent stale cached versions from executing
            const oldScript = doc.getElementById('wysiwyg-parent-script');
            if (oldScript) {
                oldScript.remove();
            }
            
            const script = doc.createElement('script');
            script.id = 'wysiwyg-parent-script';
            script.innerHTML = `
                (function() {
                    const doc = document;
                    
                    // Inject custom styles
                    let style = doc.getElementById('wysiwyg-custom-styles');
                    if (!style) {
                        style = doc.createElement('style');
                        style.id = 'wysiwyg-custom-styles';
                        doc.head.appendChild(style);
                    }
                    style.innerHTML = " .wysiwyg-editor-container code { background-color: #f1f5f9 !important; border: 1px solid #cbd5e1 !important; border-radius: 4px !important; padding: 2px 6px !important; font-family: monospace !important; color: #e11d48 !important; font-weight: 600 !important; display: inline-block !important; } textarea::selection { background-color: #fef08a !important; color: #0c0a09 !important; } textarea::-moz-selection { background-color: #fef08a !important; color: #0c0a09 !important; } ";
                    
                    function htmlToMarkdown(html) {
                        if (!html) return '';
                        const temp = doc.createElement('div');
                        temp.innerHTML = html;
                        
                        function processNode(node) {
                            if (node.nodeType === Node.TEXT_NODE) {
                                return node.textContent;
                            }
                            if (node.nodeType === Node.ELEMENT_NODE) {
                                let inner = '';
                                node.childNodes.forEach(child => {
                                    inner += processNode(child);
                                });
                                
                                const tag = node.tagName.toLowerCase();
                                const isBold = tag === 'b' || tag === 'strong' || 
                                               (node.style && (node.style.fontWeight === 'bold' || node.style.fontWeight === '700' || parseInt(node.style.fontWeight) >= 600)) ||
                                               (node.classList && (node.classList.contains('Bold') || node.classList.contains('bold')));
                                
                                const isItalic = tag === 'i' || tag === 'em' || 
                                                 (node.style && node.style.fontStyle === 'italic') ||
                                                 (node.classList && (node.classList.contains('Italic') || node.classList.contains('italic')));
                                
                                const isCode = tag === 'code' || tag === 'pre' || 
                                               (node.style && node.style.fontFamily && node.style.fontFamily.includes('monospace')) ||
                                               (node.classList && (node.classList.contains('Code') || node.classList.contains('code')));

                                if (isBold) {
                                    return "**" + inner + "**";
                                }
                                if (isItalic) {
                                    return "*" + inner + "*";
                                }
                                if (isCode) {
                                    return "{" + inner + "}";
                                }
                                if (tag === 'br') {
                                    return '\\\\n';
                                }
                                if (tag === 'div' || tag === 'p') {
                                    return inner + '\\\\n';
                                }
                                return inner;
                            }
                            return '';
                        }
                        return processNode(temp).trim();
                    }

                    function markdownToHtml(md) {
                        if (!md) return '';
                        let html = md;
                        html = html.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                        html = html.replace(/\\\\{([^{}]+)\\\\}/g, '<code>$1</code>');
                        html = html.replace(/\\\\*\\\\*([^*]+)\\\\*\\\\*/g, '<strong>$1</strong>');
                        html = html.replace(/\\\\*([^*]+)\\\\/g, '<em>$1</em>');
                        html = html.replace(/\\\\n/g, '<br>');
                        return html;
                    }
                    
                    if (!doc.getAttribute('data-wysiwyg-sync-attached')) {
                        doc.setAttribute('data-wysiwyg-sync-attached', 'true');
                        doc.addEventListener('mousedown', (e) => {
                            const btn = e.target.closest('button');
                            if (btn) {
                                const activeEditors = doc.querySelectorAll('.wysiwyg-editor-container [contenteditable="true"]');
                                activeEditors.forEach(editor => {
                                    const textarea = editor.targetTextarea;
                                    if (textarea) {
                                        const md = htmlToMarkdown(editor.innerHTML);
                                        if (textarea.value !== md) {
                                            let setSuccess = false;
                                            try {
                                                const proto = HTMLTextAreaElement.prototype;
                                                const desc = Object.getOwnPropertyDescriptor(proto, 'value');
                                                if (desc && desc.set) {
                                                    desc.set.call(textarea, md);
                                                    setSuccess = true;
                                                }
                                            } catch (err) {
                                                console.error('Error setting React value bypass:', err);
                                            }
                                            if (!setSuccess) {
                                                textarea.value = md;
                                            }
                                            textarea.dispatchEvent(new Event('input', { bubbles: true }));
                                            textarea.dispatchEvent(new Event('change', { bubbles: true }));
                                            textarea.dispatchEvent(new Event('blur', { bubbles: true }));
                                        }
                                    });
                                }
                            }, true);
                        }

                        function setupSearch() {
                            const searchInput = doc.querySelector('.html-search-input');
                            if (!searchInput) return;
                            if (searchInput.getAttribute('data-search-attached') === 'true') return;
                            searchInput.setAttribute('data-search-attached', 'true');
                            
                            const btnPrev = doc.querySelector('.html-search-prev');
                            const btnNext = doc.querySelector('.html-search-next');
                            const statusSpan = doc.querySelector('.html-search-status');
                            const debugSpan = doc.querySelector('.html-search-debug');
                            
                            if (debugSpan) {
                                debugSpan.textContent = 'Ready';
                                debugSpan.style.color = '#10b981';
                            }
                            
                            let matches = [];
                            let currentMatchIndex = -1;
                            
                            function getTextArea() {
                                // Search for visible textareas inside the same column (which has max value length)
                                const col = searchInput.closest('div[data-testid="column"]') || 
                                            searchInput.closest('div[class*="stColumn"]') || 
                                            searchInput.parentElement;
                                if (col) {
                                    const textareas = Array.from(col.querySelectorAll('textarea'));
                                    if (textareas.length > 0) {
                                        let best = textareas[0];
                                        let maxLength = best.value ? best.value.length : 0;
                                        for (let t of textareas) {
                                            const len = t.value ? t.value.length : 0;
                                            if (len > maxLength) {
                                                maxLength = len;
                                                best = t;
                                            }
                                        }
                                        return best;
                                    }
                                }
                                return doc.querySelector('textarea');
                            }
                            
                            function performSearch() {
                                const textarea = getTextArea();
                                if (debugSpan) {
                                    if (!textarea) {
                                        debugSpan.textContent = '[Err: Editor not matched]';
                                        debugSpan.style.color = '#ef4444';
                                    } else {
                                        debugSpan.textContent = '';
                                    }
                                }
                                if (!textarea) return;
                                
                                const text = textarea.value;
                                const query = searchInput.value;
                                
                                matches = [];
                                currentMatchIndex = -1;
                                
                                if (!query) {
                                    statusSpan.textContent = '0/0';
                                    const contextDiv = doc.querySelector('.html-search-context');
                                    if (contextDiv) contextDiv.style.display = 'none';
                                    return;
                                }
                                
                                let pos = text.toLowerCase().indexOf(query.toLowerCase());
                                while (pos !== -1) {
                                    matches.push({ start: pos, end: pos + query.length });
                                    pos = text.toLowerCase().indexOf(query.toLowerCase(), pos + 1);
                                }
                                
                                if (matches.length > 0) {
                                    currentMatchIndex = 0;
                                    highlightMatch(true); // Keep focus on search input when typing
                                } else {
                                    statusSpan.textContent = '0/0';
                                    const contextDiv = doc.querySelector('.html-search-context');
                                    if (contextDiv) contextDiv.style.display = 'none';
                                }
                            }
                            
                            function highlightMatch(keepInputFocus) {
                                const textarea = getTextArea();
                                if (!textarea || currentMatchIndex === -1 || matches.length === 0) return;
                                
                                const match = matches[currentMatchIndex];
                                if (!keepInputFocus) {
                                    textarea.focus();
                                }
                                textarea.setSelectionRange(match.start, match.end);
                                
                                const lines = textarea.value.split(String.fromCharCode(10));
                                
                                let lineNum = 1;
                                let charCount = 0;
                                let matchLineText = '';
                                
                                for (let i = 0; i < lines.length; i++) {
                                    const lineLen = lines[i].length + 1; // +1 for newline character
                                    if (match.start >= charCount && match.start < charCount + lineLen) {
                                        lineNum = i + 1;
                                        matchLineText = lines[i];
                                        break;
                                    }
                                    charCount += lineLen;
                                }
                                
                                const lineHeight = 20; 
                                textarea.scrollTop = (lineNum - 5) * lineHeight;
                                
                                statusSpan.textContent = (currentMatchIndex + 1) + '/' + matches.length;
                                
                                const contextDiv = doc.querySelector('.html-search-context');
                                if (contextDiv && matchLineText) {
                                    contextDiv.style.display = 'block';
                                    const safeText = matchLineText.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                                    contextDiv.innerHTML = '<strong>Line ' + lineNum + ':</strong> ' + safeText;
                                }
                            }
                            
                            searchInput.oninput = performSearch;
                            
                            btnNext.onclick = function(e) {
                                e.preventDefault();
                                e.stopPropagation();
                                if (matches.length === 0) return;
                                currentMatchIndex = (currentMatchIndex + 1) % matches.length;
                                highlightMatch(false); // Focus to textarea on click to show highlight
                            };
                            
                            btnPrev.onclick = function(e) {
                                e.preventDefault();
                                e.stopPropagation();
                                if (matches.length === 0) return;
                                currentMatchIndex = (currentMatchIndex - 1 + matches.length) % matches.length;
                                highlightMatch(false); // Focus to textarea on click to show highlight
                            };
                            
                            searchInput.onkeydown = function(e) {
                                if (e.key === 'Enter') {
                                    e.preventDefault();
                                    if (matches.length === 0) return;
                                    currentMatchIndex = (currentMatchIndex + 1) % matches.length;
                                    highlightMatch(true); // Keep focus on input for keyboard navigation cycling
                                }
                            };
                            
                        }

                        function setupWYSIWYG() {
                            const marker = doc.getElementById('editor-marker');
                            if (!marker) return;
                            
                            const editorContainer = marker.closest('div[data-testid=\"stVerticalBlock\"]') || 
                                                  marker.closest('div[class*=\"stVerticalBlock\"]') || 
                                                  marker.parentElement;
                            if (!editorContainer) return;
                            
                            const textareas = editorContainer.querySelectorAll('textarea');
                            textareas.forEach(textarea => {
                                try {
                                    if (textarea.getAttribute('data-has-wysiwyg') === 'true') return;
                                    textarea.setAttribute('data-has-wysiwyg', 'true');
                                    
                                    textarea.style.display = 'none';
                                    const basewebWrapper = textarea.closest('div[data-baseweb=\"textarea\"]');
                                    if (basewebWrapper) {
                                        basewebWrapper.style.display = 'none';
                                    }
                                    
                                    const container = doc.createElement('div');
                                    container.className = 'wysiwyg-editor-container';
                                    container.style.border = '1px solid #cbd5e1';
                                    container.style.borderRadius = '8px';
                                    container.style.padding = '12px';
                                    container.style.backgroundColor = '#ffffff';
                                    container.style.marginBottom = '12px';
                                    container.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
                                    container.style.position = 'relative';
                                    
                                    const stBlock = textarea.closest('div[data-testid=\"stVerticalBlock\"]') || 
                                                    textarea.closest('div[class*=\"stVerticalBlock\"]');
                                    const meta = stBlock ? stBlock.querySelector('.wysiwyg-metadata') : null;
                                    const idx = meta ? meta.getAttribute('data-idx') : '';
                                    const type = meta ? meta.getAttribute('data-type') : '';
                                    
                                    const fallback = stBlock ? stBlock.querySelector('.wysiwyg-fallback-title') : null;
                                    if (fallback) fallback.style.display = 'none';
                                    
                                    const floatToolbar = doc.createElement('div');
                                    floatToolbar.className = 'wysiwyg-floating-toolbar';
                                    floatToolbar.style.position = 'absolute';
                                    floatToolbar.style.display = 'none';
                                    floatToolbar.style.backgroundColor = '#ffffff';
                                    floatToolbar.style.border = '1px solid #cbd5e1';
                                    floatToolbar.style.borderRadius = '8px';
                                    floatToolbar.style.padding = '4px 6px';
                                    floatToolbar.style.gap = '4px';
                                    floatToolbar.style.boxShadow = '0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05)';
                                    floatToolbar.style.zIndex = '999';
                                    floatToolbar.style.flexDirection = 'row';
                                    
                                    function createFloatBtn(label, cmd, customFunc) {
                                        const btn = doc.createElement('button');
                                        btn.type = 'button';
                                        btn.innerHTML = label;
                                        btn.style.padding = '4px 10px';
                                        btn.style.fontSize = '12px';
                                        btn.style.fontWeight = '600';
                                        btn.style.backgroundColor = 'transparent';
                                        btn.style.border = 'none';
                                        btn.style.borderRadius = '4px';
                                        btn.style.color = '#334155';
                                        btn.style.cursor = 'pointer';
                                        btn.style.transition = 'all 0.1s';
                                        btn.style.display = 'inline-flex';
                                        btn.style.alignItems = 'center';
                                        btn.style.justifyContent = 'center';
                                        
                                        btn.onmouseover = () => btn.style.backgroundColor = '#f1f5f9';
                                        btn.onmouseout = () => btn.style.backgroundColor = 'transparent';
                                        
                                        btn.onclick = (e) => {
                                            e.preventDefault();
                                            e.stopPropagation();
                                            editor.focus();
                                            if (customFunc) {
                                                customFunc();
                                            } else {
                                                doc.execCommand(cmd, false, null);
                                            }
                                            syncToTextarea();
                                            floatToolbar.style.display = 'none';
                                        };
                                        return btn;
                                    }
                                    
                                    floatToolbar.appendChild(createFloatBtn('<span style=\"font-weight:bold;font-family:Georgia,serif;font-size:16px;color:#1e293b;\">B</span>', 'bold'));
                                    floatToolbar.appendChild(createFloatBtn('<span style=\"font-style:italic;font-family:Georgia,serif;font-size:16px;color:#1e293b;\">I</span>', 'italic'));
                                    floatToolbar.appendChild(createFloatBtn('<span style=\"font-family:monospace;font-weight:bold;font-size:15px;color:#e11d48;\">{ }</span>', null, () => {
                                        const selection = doc.getSelection();
                                        if (selection && selection.rangeCount > 0 && !selection.isCollapsed) {
                                            const range = selection.getRangeAt(0);
                                            const code = doc.createElement('code');
                                            code.textContent = selection.toString();
                                            range.deleteContents();
                                            range.insertNode(code);
                                        }
                                    }));
                                    container.appendChild(floatToolbar);
                                    
                                    function createBtn(label, cmd, customFunc) {
                                        const btn = doc.createElement('button');
                                        btn.type = 'button';
                                        btn.innerHTML = label;
                                        btn.style.padding = '4px 10px';
                                        btn.style.backgroundColor = 'transparent';
                                        btn.style.border = 'none';
                                        btn.style.borderRadius = '4px';
                                        btn.style.cursor = 'pointer';
                                        btn.style.transition = 'all 0.1s';
                                        btn.style.display = 'inline-flex';
                                        btn.style.alignItems = 'center';
                                        btn.style.justifyContent = 'center';
                                        
                                        btn.onmouseover = () => {
                                            btn.style.backgroundColor = '#f1f5f9';
                                        };
                                        btn.onmouseout = () => {
                                            btn.style.backgroundColor = 'transparent';
                                        };
                                        
                                        btn.onclick = (e) => {
                                            e.preventDefault();
                                            e.stopPropagation();
                                            editor.focus();
                                            
                                            if (customFunc) {
                                                customFunc();
                                            } else {
                                                doc.execCommand(cmd, false, null);
                                            }
                                            syncToTextarea();
                                        };
                                        return btn;
                                    }
                                    
                                    const headerRow = doc.createElement('div');
                                    headerRow.style.display = 'flex';
                                    headerRow.style.justifyContent = 'space-between';
                                    headerRow.style.alignItems = 'center';
                                    headerRow.style.marginBottom = '10px';
                                    headerRow.style.borderBottom = '1px solid #e2e8f0';
                                    headerRow.style.paddingBottom = '8px';
                                    
                                    const titleDiv = doc.createElement('div');
                                    titleDiv.style.fontSize = '14px';
                                    titleDiv.style.color = '#1e293b';
                                    
                                    const toolbar = doc.createElement('div');
                                    toolbar.style.display = 'flex';
                                    toolbar.style.gap = '6px';
                                    
                                    const boldBtn = createBtn('<span style=\"font-weight:bold;font-family:Georgia,serif;font-size:16px;color:#1e293b;\">B</span>', 'bold');
                                    const italicBtn = createBtn('<span style=\"font-style:italic;font-family:Georgia,serif;font-size:16px;color:#1e293b;\">I</span>', 'italic');
                                    const codeBtn = createBtn('<span style=\"font-family:monospace;font-weight:bold;font-size:15px;color:#e11d48;\">{ }</span>', null, () => {
                                        const selection = doc.getSelection();
                                        if (selection && selection.rangeCount > 0 && !selection.isCollapsed) {
                                            const range = selection.getRangeAt(0);
                                            const code = doc.createElement('code');
                                            code.textContent = selection.toString();
                                            range.deleteContents();
                                            range.insertNode(code);
                                        }
                                    });
                                    
                                    toolbar.appendChild(boldBtn);
                                    toolbar.appendChild(italicBtn);
                                    toolbar.appendChild(codeBtn);
                                    
                                    if (idx && type) {
                                        titleDiv.innerHTML = "<strong>Block " + idx + ": " + type + "</strong>";
                                        headerRow.appendChild(titleDiv);
                                        headerRow.appendChild(toolbar);
                                        container.appendChild(headerRow);
                                    } else {
                                        const plainToolbar = doc.createElement('div');
                                        plainToolbar.style.display = 'flex';
                                        plainToolbar.style.gap = '6px';
                                        plainToolbar.style.marginBottom = '6px';
                                        plainToolbar.appendChild(createBtn('<span style=\"font-weight:bold;font-family:Georgia,serif;font-size:16px;color:#1e293b;\">B</span>', 'bold'));
                                        plainToolbar.appendChild(createBtn('<span style=\"font-style:italic;font-family:Georgia,serif;font-size:16px;color:#1e293b;\">I</span>', 'italic'));
                                        plainToolbar.appendChild(createBtn('<span style=\"font-family:monospace;font-weight:bold;font-size:15px;color:#e11d48;\">{ }</span>', null, () => {
                                            const selection = doc.getSelection();
                                            if (selection && selection.rangeCount > 0 && !selection.isCollapsed) {
                                                const range = selection.getRangeAt(0);
                                                const code = doc.createElement('code');
                                                code.textContent = selection.toString();
                                                range.deleteContents();
                                                range.insertNode(code);
                                            }
                                        }));
                                        container.appendChild(plainToolbar);
                                    }
                                    
                                    const editor = doc.createElement('div');
                                    editor.contentEditable = 'true';
                                    editor.style.minHeight = '100px';
                                    editor.style.outline = 'none';
                                    editor.style.fontSize = '14px';
                                    editor.style.color = '#1e293b';
                                    editor.style.lineHeight = '1.5';
                                    editor.style.fontFamily = 'inherit';
                                    let initialValue = textarea.value || '';
                                    let prevVal = "";
                                    let limit = 0;
                                    while (initialValue !== prevVal && (initialValue.includes('<') || initialValue.includes('&lt;')) && limit < 10) {
                                        prevVal = initialValue;
                                        const tempClean = doc.createElement('div');
                                        tempClean.innerHTML = initialValue;
                                        initialValue = tempClean.textContent || tempClean.innerText || initialValue;
                                        limit++;
                                    }
                                    editor.innerHTML = markdownToHtml(initialValue);
                                    editor.targetTextarea = textarea;
                                    
                                    function syncToTextarea() {
                                        if (!editor.isConnected) return;
                                        const md = htmlToMarkdown(editor.innerHTML);
                                        if (textarea.value !== md) {
                                            let setSuccess = false;
                                            try {
                                                const proto = HTMLTextAreaElement.prototype;
                                                const desc = Object.getOwnPropertyDescriptor(proto, 'value');
                                                if (desc && desc.set) {
                                                    desc.set.call(textarea, md);
                                                    setSuccess = true;
                                                }
                                            } catch (err) {
                                                console.error('Error setting React value bypass:', err);
                                            }
                                            if (!setSuccess) {
                                                textarea.value = md;
                                            }
                                            textarea.dispatchEvent(new Event('input', { bubbles: true }));
                                            textarea.dispatchEvent(new Event('change', { bubbles: true }));
                                            textarea.dispatchEvent(new Event('blur', { bubbles: true }));
                                        }
                                    }
                                    
                                    editor.oninput = syncToTextarea;
                                    
                                    editor.onmouseup = editor.onkeyup = (e) => {
                                        const selection = doc.getSelection();
                                        if (selection && selection.rangeCount > 0 && !selection.isCollapsed && selection.toString().trim().length > 0) {
                                            const range = selection.getRangeAt(0);
                                            const rect = range.getBoundingClientRect();
                                            const containerRect = container.getBoundingClientRect();
                                            
                                            floatToolbar.style.display = 'flex';
                                            floatToolbar.style.top = (rect.top - containerRect.top - 38) + 'px';
                                            floatToolbar.style.left = (rect.left - containerRect.left + (rect.width / 2) - 45) + 'px';
                                        } else {
                                            floatToolbar.style.display = 'none';
                                        }
                                    };
                                    
                                    doc.addEventListener('mousedown', (e) => {
                                        if (floatToolbar.style.display === 'flex') {
                                            if (!floatToolbar.contains(e.target) && !editor.contains(e.target)) {
                                                floatToolbar.style.display = 'none';
                                            }
                                        }
                                    });
                                    
                                    editor.onkeydown = (e) => {
                                        if (e.ctrlKey && (e.key === '`' || (e.shiftKey && e.key.toLowerCase() === 'c'))) {
                                            e.preventDefault();
                                            const selection = doc.getSelection();
                                            if (selection && selection.rangeCount > 0 && !selection.isCollapsed) {
                                                const range = selection.getRangeAt(0);
                                                const code = doc.createElement('code');
                                                code.textContent = selection.toString();
                                                range.deleteContents();
                                                range.insertNode(code);
                                                syncToTextarea();
                                            }
                                        }
                                    };
                                    
                                    editor.onfocus = () => {
                                        container.style.borderColor = '#3b82f6';
                                        container.style.boxShadow = '0 0 0 1px #3b82f6';
                                    };
                                    editor.onblur = () => {
                                        container.style.borderColor = '#cbd5e1';
                                        container.style.boxShadow = '0 1px 2px rgba(0,0,0,0.05)';
                                        syncToTextarea();
                                    };
                                    
                                    container.appendChild(editor);
                                    
                                    if (basewebWrapper) {
                                        basewebWrapper.parentNode.insertBefore(container, basewebWrapper);
                                    } else {
                                        textarea.parentNode.insertBefore(container, textarea);
                                    }
                                } catch (err) {
                                    console.error('Error setting up rich editor for individual textarea:', err);
                                }
                            });
                        }
                        
                        function setupFullscreenButton() {
                            const textareas = doc.querySelectorAll('textarea');
                            let rawTextarea = null;
                            for (let t of textareas) {
                                if (t.value && (t.value.includes('<!DOCTYPE') || t.value.includes('<html') || t.value.includes('<div'))) {
                                    rawTextarea = t;
                                    break;
                                }
                            }
                            if (!rawTextarea) return;
                            
                            const stTextArea = rawTextarea.closest('[data-testid="stTextArea"]');
                            if (!stTextArea) return;
                            
                            const col = rawTextarea.closest('div[data-testid="column"]') || 
                                        rawTextarea.closest('div[class*="stColumn"]') || 
                                        rawTextarea.parentElement;
                            if (!col) return;
                            
                            let btnFullscreen = stTextArea.querySelector('.html-editor-fullscreen-btn');
                            if (!btnFullscreen) {
                                btnFullscreen = doc.createElement('button');
                                btnFullscreen.className = 'html-editor-fullscreen-btn';
                                btnFullscreen.type = 'button';
                                btnFullscreen.innerHTML = '&#x26F6;'; // ⛶
                                btnFullscreen.style.cssText = 'position: absolute; top: 0px; right: 0px; width: 34px; height: 34px; border-radius: 4px; border: 1px solid #cbd5e1; background-color: #ffffff; cursor: pointer; display: flex; align-items: center; justify-content: center; box-shadow: 0 1px 2px rgba(0,0,0,0.05); font-size: 16px; color: #475569; z-index: 999; transition: all 0.15s;';
                                btnFullscreen.onmouseover = () => { btnFullscreen.style.backgroundColor = '#f1f5f9'; };
                                btnFullscreen.onmouseout = () => { btnFullscreen.style.backgroundColor = '#ffffff'; };
                                
                                stTextArea.style.position = 'relative';
                                stTextArea.appendChild(btnFullscreen);
                                
                                btnFullscreen.onclick = function(e) {
                                    e.preventDefault();
                                    e.stopPropagation();
                                    const isFs = col.classList.contains('raw-html-fullscreen-active');
                                    if (!isFs) {
                                        col.classList.add('raw-html-fullscreen-active');
                                        col.style.position = 'fixed';
                                        col.style.top = '5vh';
                                        col.style.left = '5vw';
                                        col.style.width = '90vw';
                                        col.style.height = '90vh';
                                        col.style.zIndex = '999999';
                                        col.style.backgroundColor = '#ffffff';
                                        col.style.padding = '32px';
                                        col.style.borderRadius = '12px';
                                        col.style.boxShadow = '0 10px 30px rgba(0,0,0,0.15)';
                                        col.style.border = '1px solid #cbd5e1';
                                        col.style.overflowY = 'auto';
                                        
                                        btnFullscreen.innerHTML = '❌';
                                        btnFullscreen.style.backgroundColor = '#fecaca';
                                        btnFullscreen.style.color = '#b91c1c';
                                        btnFullscreen.style.position = 'fixed';
                                        btnFullscreen.style.top = '7vh';
                                        btnFullscreen.style.right = '7vw';
                                        if (rawTextarea) {
                                            rawTextarea.style.height = 'calc(90vh - 200px)';
                                        }
                                    } else {
                                        col.classList.remove('raw-html-fullscreen-active');
                                        col.style.position = '';
                                        col.style.top = '';
                                        col.style.left = '';
                                        col.style.width = '';
                                        col.style.height = '';
                                        col.style.zIndex = '';
                                        col.style.backgroundColor = '';
                                        col.style.padding = '';
                                        col.style.borderRadius = '';
                                        col.style.boxShadow = '';
                                        col.style.border = '';
                                        col.style.overflowY = '';
                                        
                                        btnFullscreen.innerHTML = '&#x26F6;'; // ⛶
                                        btnFullscreen.style.backgroundColor = '#ffffff';
                                        btnFullscreen.style.color = '#475569';
                                        btnFullscreen.style.position = 'absolute';
                                        btnFullscreen.style.top = '0px';
                                        btnFullscreen.style.right = '0px';
                                        if (rawTextarea) {
                                            rawTextarea.style.height = '450px';
                                        }
                                    }
                                };
                            }
                        }
                        
                        function runLoop() {
                            setupFullscreenButton();
                            setupSearch();
                            const marker = doc.getElementById('editor-marker');
                            if (marker) {
                                setupWYSIWYG();
                            } else {
                                const containers = doc.querySelectorAll('.wysiwyg-editor-container');
                                if (containers.length > 0) {
                                    containers.forEach(c => c.remove());
                                    const textareas = doc.querySelectorAll('textarea');
                                    textareas.forEach(t => {
                                        t.style.display = '';
                                        t.removeAttribute('data-has-wysiwyg');
                                        const basewebWrapper = t.closest('div[data-baseweb=\"textarea\"]');
                                        if (basewebWrapper) {
                                            basewebWrapper.style.display = '';
                                        }
                                    });
                                }
                            }
                        }
                        
                        if (window.wysiwygInterval) clearInterval(window.wysiwygInterval);
                        window.wysiwygInterval = setInterval(runLoop, 500);
                        runLoop();
                    })();
                `;
                doc.head.appendChild(script);
            }
        })();
        </script>
        ''', height=5)


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

def render_operating_manual_page():
    st.title("⚙️ Machine Configuration System")
    st.subheader("User-Managed Machine Database, Project Configuration, Multi-Sheet Excel Export & History")
    
    import importlib
    import backend.machine_project_service as machine_svc
    try:
        importlib.reload(machine_svc)
    except Exception:
        pass
    import backend.operating_manual_service as manual_service
    try:
        importlib.reload(manual_service)
    except Exception:
        pass
    
    tab1, tab2, tab3 = st.tabs([
        "📋 Create Project Machine List",
        "🗂️ Project List & History",
        "⚙️ Master Function Maintenance"
    ])
    
    # ---------------------------------------------------------------------------
    # TAB 1: Create Project Machine List (Selection & Process)
    # ---------------------------------------------------------------------------
    with tab1:
        st.markdown("### 📋 Create New Machine Configuration List")
        st.info("💡 **Instructions:** Enter Project details, select required functions via checkboxes, and click **Process** to save and download the configured Excel list.")
        
        # Project Information Form
        with st.container(border=True):
            st.markdown("#### 1️⃣ Project Specification (Header Information)")
            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                proj_no = st.text_input("Project No.", placeholder="e.g. 56061, 57791", key="cfg_proj_no")
            with col_p2:
                customer = st.text_input("Customer Name", placeholder="e.g. Henkel Maribor, Bayer, XX", key="cfg_customer")
            with col_p3:
                mtype_presets = [t for t in STANDARD_MACHINE_TYPES if t != "Any"] + ["Other (Custom)..."]
                mtype_choice = st.selectbox("Machine Type / Series", mtype_presets, key="cfg_machine_type_choice")
                if mtype_choice == "Other (Custom)...":
                    machine_type = st.text_input("Specify Custom Machine Type", placeholder="e.g. SpecialSeries", key="cfg_machine_type_custom")
                else:
                    machine_type = mtype_choice
                
            proj_notes = st.text_area("Project Notes / Scope (Optional)", placeholder="Enter any specific commissioning or configuration notes...", key="cfg_notes", height=70)
            
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Function Selection List
        with st.container(border=True):
            st.markdown("#### 2️⃣ Machine Functions Selection (Check Blocks)")
            
            # Master functions retrieval & search
            col_fsearch, col_fmain, col_ftype, col_factions = st.columns([3, 2, 2, 2])
            with col_fsearch:
                f_search = st.text_input("🔍 Search Functions / Notes", placeholder="Filter by keyword...", key="cfg_f_search")
            with col_fmain:
                all_master_funcs = machine_svc.get_master_functions()
                main_funcs_set = sorted(list(set([f["main_function"] for f in all_master_funcs if f.get("main_function")])))
                main_filter_list = ["All"] + main_funcs_set
                f_main = st.selectbox("Main Function Filter", main_filter_list, key="cfg_f_main")
            with col_ftype:
                f_type = st.selectbox("Type machine Filter", STANDARD_MACHINE_TYPES, key="cfg_f_type")
            with col_factions:
                st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("Select All", use_container_width=True, key="cfg_select_all"):
                        st.session_state["cfg_selected_funcs"] = [f["id"] for f in all_master_funcs]
                        for f in all_master_funcs:
                            st.session_state[f"func_chk_tree_{f['id']}"] = True
                            st.session_state[f"func_chk_flat_{f['id']}"] = True
                        st.rerun()
                with col_btn2:
                    if st.button("Clear All", use_container_width=True, key="cfg_clear_all"):
                        st.session_state["cfg_selected_funcs"] = []
                        for f in all_master_funcs:
                            st.session_state[f"func_chk_tree_{f['id']}"] = False
                            st.session_state[f"func_chk_flat_{f['id']}"] = False
                        st.rerun()
                        
            # Filter functions list
            filtered_funcs = machine_svc.get_master_functions(search=f_search, machine_type=f_type, main_function=f_main)
            
            if "cfg_selected_funcs" not in st.session_state:
                st.session_state["cfg_selected_funcs"] = []
                
            def on_tree_chk_change(fid):
                k = f"func_chk_tree_{fid}"
                is_on = st.session_state.get(k, False)
                if is_on and fid not in st.session_state["cfg_selected_funcs"]:
                    st.session_state["cfg_selected_funcs"].append(fid)
                elif not is_on and fid in st.session_state["cfg_selected_funcs"]:
                    st.session_state["cfg_selected_funcs"].remove(fid)
                st.session_state[f"func_chk_flat_{fid}"] = is_on

            def on_flat_chk_change(fid):
                k = f"func_chk_flat_{fid}"
                is_on = st.session_state.get(k, False)
                if is_on and fid not in st.session_state["cfg_selected_funcs"]:
                    st.session_state["cfg_selected_funcs"].append(fid)
                elif not is_on and fid in st.session_state["cfg_selected_funcs"]:
                    st.session_state["cfg_selected_funcs"].remove(fid)
                st.session_state[f"func_chk_tree_{fid}"] = is_on
                
            # View selector (Navigation Tree vs Flat List)
            col_hdr1, col_hdr2 = st.columns([3, 2])
            with col_hdr1:
                st.markdown(f"**Available Functions ({len(filtered_funcs)} items) — Selected: `{len(st.session_state['cfg_selected_funcs'])}` / `{len(all_master_funcs)}`**")
            with col_hdr2:
                view_mode = st.radio(
                    "View Mode",
                    ["🌳 Navigation Tree", "📄 Flat List"],
                    horizontal=True,
                    label_visibility="collapsed",
                    key="cfg_view_mode"
                )

            if not filtered_funcs:
                st.warning("No master functions found matching your filter criteria. Please check your search or filter.")
            elif view_mode == "🌳 Navigation Tree":
                # Group by Main Function Hierarchy
                from collections import OrderedDict
                grouped_tree = OrderedDict()
                for func in filtered_funcs:
                    mf = func.get("main_function") or "Other / General Functions"
                    if mf not in grouped_tree:
                        grouped_tree[mf] = []
                    grouped_tree[mf].append(func)

                for group_idx, (grp_name, grp_items) in enumerate(grouped_tree.items(), start=1):
                    grp_sel_count = sum(1 for it in grp_items if it["id"] in st.session_state["cfg_selected_funcs"])
                    grp_total = len(grp_items)
                    
                    # Status Indicator
                    if grp_sel_count == grp_total:
                        status_icon = "🟢"
                    elif grp_sel_count > 0:
                        status_icon = "🟡"
                    else:
                        status_icon = "⚪"
                        
                    exp_title = f"{status_icon} 📂 {group_idx}. {grp_name} — ({grp_total} Functions • Selected: {grp_sel_count}/{grp_total})"
                    
                    # Default expanded if filtered or group has selections
                    is_expanded = bool(f_search.strip() or f_main != "All" or grp_sel_count > 0 or len(grouped_tree) <= 3)
                    
                    with st.expander(exp_title, expanded=is_expanded):
                        # Group Header Quick Actions
                        col_ga1, col_ga2, col_ga3 = st.columns([6, 2, 2])
                        with col_ga1:
                            st.caption(f"📁 Station Group: **{grp_name}** | **{grp_total}** Functions")
                        with col_ga2:
                            if st.button("Select Group", key=f"tree_sel_all_{group_idx}_{grp_name[:12]}", use_container_width=True):
                                for it in grp_items:
                                    fid = it["id"]
                                    if fid not in st.session_state["cfg_selected_funcs"]:
                                        st.session_state["cfg_selected_funcs"].append(fid)
                                    st.session_state[f"func_chk_tree_{fid}"] = True
                                    st.session_state[f"func_chk_flat_{fid}"] = True
                                st.rerun()
                        with col_ga3:
                            if st.button("Clear Group", key=f"tree_clr_all_{group_idx}_{grp_name[:12]}", use_container_width=True):
                                for it in grp_items:
                                    fid = it["id"]
                                    if fid in st.session_state["cfg_selected_funcs"]:
                                        st.session_state["cfg_selected_funcs"].remove(fid)
                                    st.session_state[f"func_chk_tree_{fid}"] = False
                                    st.session_state[f"func_chk_flat_{fid}"] = False
                                st.rerun()
                        
                        st.markdown("<hr style='margin: 4px 0 10px 0;'>", unsafe_allow_html=True)
                        
                        # Render function cards inside tree group
                        for func in grp_items:
                            fid = func["id"]
                            t_key = f"func_chk_tree_{fid}"
                            if t_key not in st.session_state:
                                st.session_state[t_key] = (fid in st.session_state["cfg_selected_funcs"])
                            
                            with st.container(border=True):
                                fc1, fc2, fc3 = st.columns([1, 2, 4])
                                
                                with fc1:
                                    img_path = func.get("image_path")
                                    if img_path and os.path.exists(img_path):
                                        try:
                                            st.image(img_path, width=110)
                                        except Exception:
                                            st.markdown("🖼️ *(Image preview)*")
                                    else:
                                        st.markdown("""
                                        <div style="background-color: #0b1a30; border: 1px dashed #334155; border-radius: 8px; height: 80px; display: flex; align-items: center; justify-content: center; color: #64748b; font-size: 0.8rem; text-align: center; padding: 5px;">
                                            ⚙️ No Image
                                        </div>
                                        """, unsafe_allow_html=True)
                                        
                                with fc2:
                                    st.checkbox(
                                        f"**{func['function_name']}**",
                                        key=t_key,
                                        on_change=on_tree_chk_change,
                                        args=(fid,)
                                    )
                                    st.caption(f"📌 Main Function: `{func.get('main_function') or 'General'}` | ⚙️ Type: `{func.get('machine_type', 'Any')}`")
                                    
                                with fc3:
                                    note_text = func.get("note", "").strip()
                                    if note_text:
                                        st.markdown(f"📝 **Note:** {note_text}")
                                    else:
                                        st.caption("📝 No specific note recorded for this function.")
            else:
                # Flat List View
                for func in filtered_funcs:
                    fid = func["id"]
                    fl_key = f"func_chk_flat_{fid}"
                    if fl_key not in st.session_state:
                        st.session_state[fl_key] = (fid in st.session_state["cfg_selected_funcs"])
                    
                    with st.container(border=True):
                        fc1, fc2, fc3 = st.columns([1, 2, 4])
                        
                        with fc1:
                            img_path = func.get("image_path")
                            if img_path and os.path.exists(img_path):
                                try:
                                    st.image(img_path, width=110)
                                except Exception:
                                    st.markdown("🖼️ *(Image preview)*")
                            else:
                                st.markdown("""
                                <div style="background-color: #0b1a30; border: 1px dashed #334155; border-radius: 8px; height: 80px; display: flex; align-items: center; justify-content: center; color: #64748b; font-size: 0.8rem; text-align: center; padding: 5px;">
                                    ⚙️ No Image
                                </div>
                                """, unsafe_allow_html=True)
                                
                        with fc2:
                            st.checkbox(
                                f"**{func['function_name']}**",
                                key=fl_key,
                                on_change=on_flat_chk_change,
                                args=(fid,)
                            )
                            st.caption(f"📌 Main Function: `{func.get('main_function') or 'General'}` | ⚙️ Type: `{func.get('machine_type', 'Any')}`")
                            
                        with fc3:
                            note_text = func.get("note", "").strip()
                            if note_text:
                                st.markdown(f"📝 **Note:** {note_text}")
                            else:
                                st.caption("📝 No specific note recorded for this function.")
                                
        st.markdown("<br>", unsafe_allow_html=True)
        
        # 3️⃣ Process Button (5.3)
        col_proc1, col_proc2 = st.columns([2, 1])
        with col_proc1:
            st.markdown(f"**Ready to Process:** Project: `{proj_no or 'Not specified'}` | Customer: `{customer or 'Not specified'}` | Type: `{machine_type or 'Not specified'}` | Selected: **{len(st.session_state.get('cfg_selected_funcs', []))}** functions")
        with col_proc2:
            process_btn = st.button("🚀 Process & Save Machine List", type="primary", use_container_width=True, key="cfg_process_btn")
            
        if process_btn:
            if not proj_no.strip():
                st.error("❌ Please specify the Project No. (e.g. 5xxxxx)")
            elif not customer.strip():
                st.error("❌ Please specify the Customer Name.")
            elif not machine_type.strip():
                st.error("❌ Please specify the Machine Type.")
            elif not st.session_state.get("cfg_selected_funcs"):
                st.error("❌ Please select at least one function for this machine.")
            else:
                uploader_name = st.session_state.get("username", "Operator")
                with st.spinner("Processing configuration and generating Excel file..."):
                    success, msg, project_id, excel_bytes = machine_svc.save_project_record(
                        project_no=proj_no,
                        customer=customer,
                        machine_type=machine_type,
                        selected_function_ids=st.session_state["cfg_selected_funcs"],
                        created_by=uploader_name,
                        notes=proj_notes
                    )
                    
                    if success:
                        st.success(f"✅ {msg}")
                        st.session_state["last_processed_proj_id"] = project_id
                        st.session_state["last_processed_excel_bytes"] = excel_bytes
                        st.session_state["last_processed_proj_no"] = proj_no
                        st.session_state["last_processed_customer"] = customer
                    else:
                        st.error(f"❌ {msg}")
                        
        # Direct Download for recently processed project (5.3)
        if "last_processed_excel_bytes" in st.session_state and st.session_state["last_processed_excel_bytes"]:
            st.markdown("### 📥 Download Configuration Files")
            st.info("🖼️ **Note:** The downloaded Excel file (.xlsx) contains embedded pictures for each selected function in Sheet 1, and a complete machine function matrix in Sheet 2 with green selection highlights. You can also download the ZIP Package for original high-resolution picture files.")
            col_dl1, col_dl2 = st.columns(2)
            with col_dl1:
                st.download_button(
                    label=f"📥 Download Project Excel: {st.session_state.get('last_processed_proj_no', '')} (.xlsx)",
                    data=st.session_state["last_processed_excel_bytes"],
                    file_name=f"Project_{st.session_state.get('last_processed_proj_no', 'Machine')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                    key="dl_fresh_excel_btn"
                )
            with col_dl2:
                pid = st.session_state.get("last_processed_proj_id")
                if pid:
                    zip_res = machine_svc.generate_project_zip_bundle(pid)
                    if zip_res:
                        zname, zdata = zip_res
                        st.download_button(
                            label=f"📦 Download Package (.zip) [Excel + Image Folder]",
                            data=zdata,
                            file_name=zname,
                            mime="application/zip",
                            type="secondary",
                            use_container_width=True,
                            key="dl_fresh_zip_btn"
                        )

    # ---------------------------------------------------------------------------
    # TAB 2: Project List & History (Filter & Retrospective View)
    # ---------------------------------------------------------------------------
    with tab2:
        st.markdown("### 🗂️ Machine Project List & Historical Records")
        st.info("💡 **History & Filter:** View saved machine project lists, filter by Project No., Customer, or Machine Type, and download the Excel file. This data can be consumed by other tools in the future.")
        
        # Filter Bar
        col_flt1, col_flt2, col_flt3, col_flt4 = st.columns([2, 2, 2, 1])
        with col_flt1:
            hist_proj_filter = st.text_input("🔍 Filter Project", placeholder="e.g. 56061", key="hist_filter_proj")
        with col_flt2:
            hist_cust_filter = st.text_input("🔍 Filter Customer", placeholder="e.g. Henkel", key="hist_filter_cust")
        with col_flt3:
            all_records_for_types = machine_svc.get_project_records()
            hist_type_filter = st.selectbox("Machine Type", STANDARD_MACHINE_TYPES, key="hist_filter_type")
        with col_flt4:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            if st.button("🔄 Refresh", use_container_width=True, key="hist_refresh_btn"):
                st.rerun()
                
        # Query Projects
        projects = machine_svc.get_project_records(
            filter_project=hist_proj_filter,
            filter_customer=hist_cust_filter,
            filter_type=hist_type_filter
        )
        
        if not projects:
            st.warning("No project records found matching the filter criteria.")
        else:
            st.markdown(f"**Found {len(projects)} Project Record(s):**")
            
            # Display Table
            summary_df = pd.DataFrame([{
                "ID": p["id"],
                "Project No": p["project_no"],
                "Customer": p["customer"],
                "Machine Type": p["machine_type"],
                "Total Functions": p["total_functions"],
                "Created At": p["created_at"],
                "Created By": p["created_by"],
                "Excel File": p.get("excel_filename", "-")
            } for p in projects])
            
            st.dataframe(summary_df, use_container_width=True)
            
            # Quick Batch Delete Manager
            with st.expander("🗑️ Delete Test or Erroneous Project Records", expanded=False):
                st.markdown("Select project records you wish to delete:")
                del_col1, del_col2 = st.columns([4, 1])
                with del_col1:
                    proj_to_delete = st.multiselect(
                        "Select Projects to Delete",
                        options=[p["id"] for p in projects],
                        format_func=lambda pid: next((f"ID {p['id']} - Project #{p['project_no']} ({p['customer']}, {p['machine_type']}) [{p['created_at']}]" for p in projects if p["id"] == pid), str(pid)),
                        key="multi_del_projects"
                    )
                with del_col2:
                    st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
                    if st.button("🗑️ Delete Selected", type="primary", use_container_width=True, key="btn_confirm_multi_del", disabled=len(proj_to_delete)==0):
                        del_count = 0
                        for pid in proj_to_delete:
                            ok, _ = machine_svc.delete_project_record(pid)
                            if ok:
                                del_count += 1
                        st.success(f"Deleted {del_count} project record(s) successfully!")
                        st.rerun()
            
            # Select Project to View Details & Download
            st.markdown("---")
            st.markdown("#### 🔍 Project Detail Inspector & Download")
            
            proj_options = {f"Project #{p['project_no']} - {p['customer']} ({p['machine_type']}) [Created: {p['created_at']}]": p["id"] for p in projects}
            selected_proj_label = st.selectbox("Select Project Record to Inspect", list(proj_options.keys()), key="hist_select_proj")
            
            selected_pid = proj_options[selected_proj_label]
            proj_details = machine_svc.get_project_details(selected_pid)
            
            if proj_details:
                col_dinfo1, col_dinfo2, col_ddl = st.columns([3, 3, 2])
                with col_dinfo1:
                    st.markdown(f"**Project No:** `{proj_details['project_no']}` | **Customer:** `{proj_details['customer']}`")
                    st.caption(f"Created: {proj_details['created_at']} by {proj_details.get('created_by', 'Operator')}")
                with col_dinfo2:
                    st.markdown(f"**Machine Type:** `{proj_details['machine_type']}` | **Functions Count:** `{len(proj_details['items'])}`")
                    if proj_details.get("notes"):
                        st.caption(f"Notes: {proj_details['notes']}")
                with col_ddl:
                    # Download Excel file & ZIP Package
                    excel_res = machine_svc.get_project_excel_bytes(selected_pid)
                    if excel_res:
                        fn, xdata = excel_res
                        st.download_button(
                            label="📥 Download Excel (.xlsx)",
                            data=xdata,
                            file_name=fn,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True,
                            key=f"hist_dl_{selected_pid}"
                        )
                    zip_res = machine_svc.generate_project_zip_bundle(selected_pid)
                    if zip_res:
                        zname, zdata = zip_res
                        st.download_button(
                            label="📦 Download ZIP Package",
                            data=zdata,
                            file_name=zname,
                            mime="application/zip",
                            type="secondary",
                            use_container_width=True,
                            key=f"hist_zip_{selected_pid}"
                        )
                    
                    # Delete Record Popover
                    with st.popover("🗑️ Delete Record", use_container_width=True):
                        st.warning(f"Delete Project **#{proj_details['project_no']}** ({proj_details['customer']})?")
                        st.caption("This will permanently remove this project record and its generated files.")
                        if st.button("🔴 Confirm Delete", type="primary", key=f"del_proj_btn_{selected_pid}", use_container_width=True):
                            d_ok, d_msg = machine_svc.delete_project_record(selected_pid)
                            if d_ok:
                                st.success(d_msg)
                                st.rerun()
                            else:
                                st.error(d_msg)
                        
                # Show items in this project
                with st.expander("👁️ View Configured Functions & Details", expanded=True):
                    for item in proj_details["items"]:
                        with st.container(border=True):
                            ic1, ic2, ic3 = st.columns([1, 3, 4])
                            with ic1:
                                ipath = item.get("image_path")
                                if ipath and os.path.exists(ipath):
                                    try:
                                        st.image(ipath, width=100)
                                    except Exception:
                                        st.caption("🖼️ Image Attached")
                                else:
                                    st.caption("No image")
                            with ic2:
                                st.markdown(f"**{item['function_name']}**")
                                st.caption(f"📌 Main Function: `{item.get('main_function') or 'General'}` | Type machine: `{item.get('machine_type', 'Any')}`")
                            with ic3:
                                st.markdown(f"📝 {item.get('note', 'No note specified.')}")

    # ---------------------------------------------------------------------------
    # TAB 3: Master Function Maintenance (4.1 Function, 4.2 Picture, 4.3 Note, Type machine)
    # ---------------------------------------------------------------------------
    with tab3:
        st.markdown("### ⚙️ Master Function Maintenance")
        st.info("💡 **Maintain Master Functions:** Manage function names, main function categories, machine types, upload function pictures, and add technical notes. Changes made here will be available for all future project lists.")
        
        # 1. Add New Master Function
        with st.expander("➕ Add New Master Machine Function", expanded=False):
            with st.form("add_master_func_form", clear_on_submit=True):
                # Left Column: Function Name & Type machine | Right Column: Upload Picture & Main Function Name
                col_left, col_right = st.columns(2)
                with col_left:
                    new_func_name = st.text_input("Function Name *", placeholder="e.g. Main flap guide bottom")
                    new_mtype_choice = st.selectbox("Type machine", STANDARD_MACHINE_TYPES + ["Other (Custom)..."], key="add_mtype_choice")
                    if new_mtype_choice == "Other (Custom)...":
                        new_func_mtype = st.text_input("Enter Custom Type", placeholder="e.g. CustomSeries", key="add_mtype_custom")
                    else:
                        new_func_mtype = new_mtype_choice
                with col_right:
                    new_func_img = st.file_uploader("Upload Picture Function", type=["png", "jpg", "jpeg"], key="new_func_img_uploader")
                    new_main_func = st.text_input("Main Function Name", placeholder="e.g. Infeed Station, Forming Unit, Inserter...", key="add_main_func")
                    
                # Note (Full Width)
                new_func_note = st.text_area("Function Note / Description", placeholder="Describe function operation, safety points, or specifications...", height=80)
                
                submitted = st.form_submit_button("💾 Save New Function", type="primary", use_container_width=True)
                if submitted:
                    if not new_func_name.strip():
                        st.error("Function Name is required.")
                    else:
                        img_bytes = new_func_img.getvalue() if new_func_img else None
                        img_fname = new_func_img.name if new_func_img else ""
                        a_ok, a_msg = machine_svc.add_master_function(
                            function_name=new_func_name,
                            main_function=new_main_func,
                            image_bytes=img_bytes,
                            image_filename=img_fname,
                            note=new_func_note,
                            category="General",
                            machine_type=new_func_mtype
                        )
                        if a_ok:
                            st.success(a_msg)
                            st.rerun()
                        else:
                            st.error(a_msg)
                            
        # 2. Existing Master Functions Gallery / Table
        st.markdown("#### 📚 Current Master Functions & Picture Library")
        
        # Library Filter Bar
        col_flt_s, col_flt_m, col_flt_t, col_flt_v = st.columns([3, 2, 2, 2])
        with col_flt_s:
            search_m = st.text_input("🔍 Search Function / Notes", placeholder="Search by keyword...", key="maint_search_f")
        with col_flt_m:
            all_m_funcs = machine_svc.get_master_functions()
            main_f_opts = ["All"] + sorted(list(set([f["main_function"] for f in all_m_funcs if f.get("main_function")])))
            main_m = st.selectbox("Filter Main Function", main_f_opts, key="maint_main_f")
        with col_flt_t:
            type_m = st.selectbox("Filter Type machine", STANDARD_MACHINE_TYPES, key="maint_type_f")
        with col_flt_v:
            st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
            maint_view = st.radio(
                "Maint View Mode",
                ["🌳 Tree View", "📄 Flat View"],
                horizontal=True,
                label_visibility="collapsed",
                key="maint_view_mode"
            )
            
        all_funcs = machine_svc.get_master_functions(search=search_m, machine_type=type_m, main_function=main_m)
        
        if not all_funcs:
            st.info("No master functions found matching your filter.")
        elif maint_view == "🌳 Tree View":
            # Grouped Tree View
            from collections import OrderedDict
            maint_tree = OrderedDict()
            for f in all_funcs:
                mf = f.get("main_function") or "Other / General Functions"
                if mf not in maint_tree:
                    maint_tree[mf] = []
                maint_tree[mf].append(f)
                
            for g_idx, (g_name, g_items) in enumerate(maint_tree.items(), start=1):
                exp_title = f"📂 {g_idx}. {g_name} ({len(g_items)} functions)"
                with st.expander(exp_title, expanded=True):
                    for f in g_items:
                        fid = f["id"]
                        with st.container(border=True):
                            mc1, mc2, mc3, mc4 = st.columns([2, 4, 4, 2])
                            with mc1:
                                ip = f.get("image_path")
                                if ip and os.path.exists(ip):
                                    try:
                                        st.image(ip, width=120)
                                    except Exception:
                                        st.caption("🖼️ Image Attached")
                                else:
                                    st.markdown("""
                                    <div style="background-color: #0b1a30; border: 1px dashed #334155; border-radius: 8px; height: 75px; display: flex; align-items: center; justify-content: center; color: #64748b; font-size: 0.75rem;">
                                        No Picture
                                    </div>
                                    """, unsafe_allow_html=True)
                            with mc2:
                                st.markdown(f"**{f['function_name']}**")
                                st.caption(f"📌 Main Function: `{f.get('main_function') or 'General'}` | ⚙️ Type: `{f.get('machine_type', 'Any')}` | ID: `{fid}`")
                            with mc3:
                                st.markdown(f"📝 **Note:** {f.get('note', 'No note.')}")
                            with mc4:
                                with st.popover("✏️ Edit / 🗑️ Delete"):
                                    st.markdown(f"**Manage Function #{fid}**")
                                    col_pe1, col_pe2 = st.columns(2)
                                    with col_pe1:
                                        edit_name = st.text_input("Function Name", value=f["function_name"], key=f"edit_t_name_{fid}")
                                        edit_main = st.text_input("Main Function Name", value=f.get("main_function", ""), key=f"edit_t_main_{fid}")
                                        current_mtype = f.get("machine_type", "Any") or "Any"
                                        type_opts = list(STANDARD_MACHINE_TYPES)
                                        if current_mtype not in type_opts and current_mtype:
                                            type_opts.append(current_mtype)
                                        type_opts.append("Other (Custom)...")
                                        
                                        mtype_idx = type_opts.index(current_mtype) if current_mtype in type_opts else 0
                                        edit_mtype_choice = st.selectbox("Type machine", type_opts, index=mtype_idx, key=f"edit_t_mtype_sel_{fid}")
                                        if edit_mtype_choice == "Other (Custom)...":
                                            edit_mtype = st.text_input("Custom Type Name", value=current_mtype, key=f"edit_t_mtype_cust_{fid}")
                                        else:
                                            edit_mtype = edit_mtype_choice
                                    with col_pe2:
                                        edit_img = st.file_uploader("Replace Picture", type=["png", "jpg", "jpeg"], key=f"edit_t_img_{fid}")
                                    
                                    edit_note = st.text_area("Note", value=f.get("note", ""), key=f"edit_t_note_{fid}")
                                    
                                    col_save_e, col_del_e = st.columns(2)
                                    with col_save_e:
                                        if st.button("Save", type="primary", key=f"btn_t_save_e_{fid}", use_container_width=True):
                                            img_b = edit_img.getvalue() if edit_img else None
                                            img_fn = edit_img.name if edit_img else ""
                                            u_ok, u_msg = machine_svc.update_master_function(
                                                func_id=fid,
                                                function_name=edit_name,
                                                main_function=edit_main,
                                                image_bytes=img_b,
                                                image_filename=img_fn,
                                                note=edit_note,
                                                category="General",
                                                machine_type=edit_mtype,
                                                keep_existing_image=(img_b is None)
                                            )
                                            if u_ok:
                                                st.success(u_msg)
                                                st.rerun()
                                            else:
                                                st.error(u_msg)
                                    with col_del_e:
                                        if st.button("Delete", type="secondary", key=f"btn_t_del_e_{fid}", use_container_width=True):
                                            d_ok, d_msg = machine_svc.delete_master_function(fid)
                                            if d_ok:
                                                st.success(d_msg)
                                                st.rerun()
                                            else:
                                                st.error(d_msg)
        else:
            # Flat View
            for f in all_funcs:
                fid = f["id"]
                with st.container(border=True):
                    mc1, mc2, mc3, mc4 = st.columns([2, 4, 4, 2])
                    with mc1:
                        # Picture Function (4.2)
                        ip = f.get("image_path")
                        if ip and os.path.exists(ip):
                            try:
                                st.image(ip, width=120)
                            except Exception:
                                st.caption("🖼️ Image Attached")
                        else:
                            st.markdown("""
                            <div style="background-color: #0b1a30; border: 1px dashed #334155; border-radius: 8px; height: 75px; display: flex; align-items: center; justify-content: center; color: #64748b; font-size: 0.75rem;">
                                No Picture
                            </div>
                            """, unsafe_allow_html=True)
                    with mc2:
                        st.markdown(f"**{f['function_name']}**")
                        st.caption(f"📌 Main Function: `{f.get('main_function') or 'General'}` | ⚙️ Type: `{f.get('machine_type', 'Any')}` | ID: `{fid}`")
                    with mc3:
                        st.markdown(f"📝 **Note:** {f.get('note', 'No note.')}")
                    with mc4:
                        # Action Buttons (Edit / Delete)
                        with st.popover("✏️ Edit / 🗑️ Delete"):
                            st.markdown(f"**Manage Function #{fid}**")
                            col_pe1, col_pe2 = st.columns(2)
                            with col_pe1:
                                edit_name = st.text_input("Function Name", value=f["function_name"], key=f"edit_name_{fid}")
                                edit_main = st.text_input("Main Function Name", value=f.get("main_function", ""), key=f"edit_main_{fid}")
                                current_mtype = f.get("machine_type", "Any") or "Any"
                                type_opts = list(STANDARD_MACHINE_TYPES)
                                if current_mtype not in type_opts and current_mtype:
                                    type_opts.append(current_mtype)
                                type_opts.append("Other (Custom)...")
                                
                                mtype_idx = type_opts.index(current_mtype) if current_mtype in type_opts else 0
                                edit_mtype_choice = st.selectbox("Type machine", type_opts, index=mtype_idx, key=f"edit_mtype_sel_{fid}")
                                if edit_mtype_choice == "Other (Custom)...":
                                    edit_mtype = st.text_input("Custom Type Name", value=current_mtype, key=f"edit_mtype_cust_{fid}")
                                else:
                                    edit_mtype = edit_mtype_choice
                            with col_pe2:
                                edit_img = st.file_uploader("Replace Picture", type=["png", "jpg", "jpeg"], key=f"edit_img_{fid}")
                            
                            edit_note = st.text_area("Note", value=f.get("note", ""), key=f"edit_note_{fid}")
                            
                            col_save_e, col_del_e = st.columns(2)
                            with col_save_e:
                                if st.button("Save", type="primary", key=f"btn_save_e_{fid}", use_container_width=True):
                                    img_b = edit_img.getvalue() if edit_img else None
                                    img_fn = edit_img.name if edit_img else ""
                                    u_ok, u_msg = machine_svc.update_master_function(
                                        func_id=fid,
                                        function_name=edit_name,
                                        main_function=edit_main,
                                        image_bytes=img_b,
                                        image_filename=img_fn,
                                        note=edit_note,
                                        category="General",
                                        machine_type=edit_mtype,
                                        keep_existing_image=(img_b is None)
                                    )
                                    if u_ok:
                                        st.success(u_msg)
                                        st.rerun()
                                    else:
                                        st.error(u_msg)
                            with col_del_e:
                                if st.button("Delete", type="secondary", key=f"btn_del_e_{fid}", use_container_width=True):
                                    d_ok, d_msg = machine_svc.delete_master_function(fid)
                                    if d_ok:
                                        st.success(d_msg)
                                        st.rerun()
                                    else:
                                        st.error(d_msg)

def render_iqoqdq_page(sub_section: str = "📐 DQ Stage"):
    st.title("📋 IQOQDQ Qualification System")
    st.subheader(f"Equipment Validation & Qualification Protocol Workspace — **{sub_section}**")
    
    # Render direct view based on semi-sub menu selection
    if "DQ" in sub_section:
        st.markdown("### 📐 DQ — Design Qualification")
        st.caption("Equipment design specification review, User Requirement Specification (URS), technical drawings, and design qualification protocol generation.")
        
        st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
        
        with st.container(border=True):
            st.markdown("### 🛠️ To be Dev. (To be Developed)")
            st.markdown("""
            This module is reserved for future Design Qualification (DQ) workspace enhancements.
            
            * **Module Purpose:** User Requirement Specification (URS), Functional Design Specification (FDS), technical drawings verification, and DQ Protocol Document Generation.
            * **Development Status:** To be Developed by user.
            """)
            st.info("💡 **Notice:** All content in this section is currently set to To be Dev. and will be updated in a future release.")

    elif "IQ" in sub_section:
        st.markdown("### 🔧 IQ — Installation Qualification")
        st.caption("Equipment installation inspection, utility connections, safety components, and format verification.")
        
        tab_iq_inst, tab_iq_cci, tab_iq_fmt = st.tabs([
            "🔧 IQ Installation",
            "🛡️ IQ CCI",
            "📦 IQ Format"
        ])
        
        with tab_iq_inst:
            st.markdown("#### 🔧 IQ Installation Protocol Generator")
            st.caption("Generate project-specific IQ Installation qualification documents by matching configured functions against Word template test points.")
            
            import backend.iq_installation_service as iq_inst_svc
            try:
                importlib.reload(iq_inst_svc)
            except Exception:
                pass
            
            subtab_gen, subtab_tags = st.tabs([
                "⚡ 1. Generate IQ Installation Protocol",
                "🏷️ 2. Tag Mapping Management (บำรุงรักษา Tag)"
            ])
            
            # Sub-Tab 1: Generation Workflow
            with subtab_gen:
                st.markdown("##### 🎛️ 1. Select Target Machine Protocol(s)")
                st.caption("Select which machine protocols apply to this project. Unselected sections will be completely removed.")
                
                proto_col1, proto_col2, proto_col3 = st.columns(3)
                with proto_col1:
                    sel_cartoner = st.checkbox("3 Test Protocol (Cartoner)", value=False, key="iq_inst_sel_cartoner")
                with proto_col2:
                    sel_tubefiller = st.checkbox("4 Test Protocol (Tube Filler)", value=False, key="iq_inst_sel_tubefiller")
                with proto_col3:
                    sel_fp10 = st.checkbox("5 Test Protocol (Tube Filler IWK FP 10)", value=True, key="iq_inst_sel_fp10")
                
                selected_protocols = []
                if sel_cartoner:
                    selected_protocols.append("Cartoner")
                if sel_tubefiller:
                    selected_protocols.append("Tube Filler")
                if sel_fp10:
                    selected_protocols.append("Tube Filler FP 10")
                
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                
                col_excel_inst, col_word_inst = st.columns(2)
                
                with col_excel_inst:
                    st.markdown("##### 📊 2. Excel Project Configuration")
                    sample_excel_inst = os.path.abspath(r"IQOQDQ/IQ_Installation/Project_XXXXX.xlsx")
                    if not os.path.exists(sample_excel_inst):
                        sample_excel_inst = os.path.abspath(r"IQOQDQ/Data IQOQDQ/Project_XXXXX.xlsx")
                    
                    sample_excel_bytes_inst = b""
                    if os.path.exists(sample_excel_inst):
                        with open(sample_excel_inst, "rb") as sf:
                            sample_excel_bytes_inst = sf.read()
                    
                    uploaded_excel_inst = st.file_uploader(
                        "Upload Project Excel File (*.xlsx, *.xls)",
                        type=["xlsx", "xls"],
                        key="iq_inst_upload_excel",
                        help="Upload Project_XXXXX.xlsx containing Column C Function Names"
                    )
                    
                    if uploaded_excel_inst is not None:
                        excel_input_inst = uploaded_excel_inst
                        st.caption(f"📂 Active File: `{uploaded_excel_inst.name}`")
                    elif os.path.exists(sample_excel_inst):
                        excel_input_inst = sample_excel_inst
                        st.caption(f"💡 *Defaulting to sample:* `{os.path.basename(sample_excel_inst)}`")
                    else:
                        excel_input_inst = None
                    
                    if sample_excel_bytes_inst:
                        st.download_button(
                            label="📥 Download Sample Excel (Project_XXXXX.xlsx)",
                            data=sample_excel_bytes_inst,
                            file_name="Project_XXXXX.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="iq_inst_dl_sample_excel",
                            use_container_width=True
                        )
                
                with col_word_inst:
                    st.markdown("##### 📄 3. Word Master Template")
                    tmpl_dir_inst = os.path.abspath(r"IQOQDQ/IQ_Installation")
                    tmpl_files_inst = [f for f in os.listdir(tmpl_dir_inst) if not f.startswith("~$") and f.lower().endswith(('.doc', '.docx'))] if os.path.exists(tmpl_dir_inst) else []
                    word_tmpl_path_inst = None
                    if tmpl_files_inst:
                        default_tmpl_idx_inst = 0
                        for idx, fn in enumerate(tmpl_files_inst):
                            if "02_iq_installation" in fn.lower():
                                default_tmpl_idx_inst = idx
                                break
                        selected_tmpl_inst = st.selectbox("Select Word Master Template", tmpl_files_inst, index=default_tmpl_idx_inst, key="iq_inst_tmpl_sel")
                        word_tmpl_path_inst = os.path.abspath(os.path.join(tmpl_dir_inst, selected_tmpl_inst))
                        st.caption(f"📄 Master Template: `{selected_tmpl_inst}` *(from IQ_Installation)*")
                        
                        # Download current template button
                        with open(word_tmpl_path_inst, "rb") as tf:
                            tmpl_bytes = tf.read()
                        st.download_button(
                            label=f"📥 Download Current Master Template ({selected_tmpl_inst})",
                            data=tmpl_bytes,
                            file_name=selected_tmpl_inst,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="iq_inst_dl_current_tmpl",
                            use_container_width=True
                        )
                    else:
                        uploaded_tmpl_inst = st.file_uploader("Upload Word Master Template (*.docx, *.doc)", type=["docx", "doc"], key="iq_inst_upload_word")
                        if uploaded_tmpl_inst:
                            temp_tmpl_p = os.path.abspath(os.path.join(tmpl_dir_inst, uploaded_tmpl_inst.name))
                            with open(temp_tmpl_p, "wb") as f:
                                f.write(uploaded_tmpl_inst.getvalue())
                            word_tmpl_path_inst = temp_tmpl_p

                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                
                # Action Button
                btn_gen_col1, btn_gen_col2 = st.columns([1.2, 2])
                with btn_gen_col1:
                    process_btn_inst = st.button("⚡ Generate IQ Installation Document", type="primary", use_container_width=True, key="iq_inst_process_btn")
                
                if process_btn_inst:
                    if not selected_protocols:
                        st.error("❌ Please select at least one Test Protocol.")
                    elif not excel_input_inst:
                        st.error("❌ Please select or upload an Excel Project file.")
                    else:
                        with st.spinner("Analyzing functions, pruning untagged rows, and generating document..."):
                            try:
                                result_inst = iq_inst_svc.generate_iq_installation_word(
                                    excel_source=excel_input_inst,
                                    selected_protocols=selected_protocols,
                                    word_template_path=word_tmpl_path_inst
                                )
                                st.session_state["iq_inst_last_result"] = result_inst
                                st.success(f"🎉 **Word Document Ready!** Click the button below to download `{result_inst['file_name']}` directly to your Downloads folder.")
                            except Exception as ex:
                                st.error(f"❌ Error generating IQ Installation document: {str(ex)}")
                
                # Display Results
                if "iq_inst_last_result" in st.session_state and st.session_state["iq_inst_last_result"]:
                    res_inst = st.session_state["iq_inst_last_result"]
                    
                    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("📋 Protocols", ", ".join(res_inst["selected_protocols"]))
                    m2.metric("✅ Test Points Kept", f"{res_inst['total_kept']} Items")
                    m3.metric("✂️ Test Points Pruned", f"{res_inst['total_pruned']} Items")
                    m4.metric("📄 Ready for Download", res_inst["file_name"])
                    
                    st.download_button(
                        label=f"📥 Download {res_inst['file_name']}",
                        data=res_inst["doc_bytes"],
                        file_name=res_inst["file_name"],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        key="iq_inst_dl_btn",
                        use_container_width=True
                    )
                    
                    with st.expander("👀 View Test Point Decisions (Kept vs Pruned Breakdown)", expanded=True):
                        st.dataframe(res_inst["breakdown_df"], use_container_width=True, hide_index=True)

            # Sub-Tab 2: Tag Mapping Management
            with subtab_tags:
                st.markdown("##### 🏷️ จัดการและกำหนด Tag Mapping ด้วยตนเอง (Interactive Tag Mapper)")
                st.info("💡 **วิธีจับคู่ Tag:** คุณสามารถเลือกหรือย้าย Function Name จาก Excel Column C มาใส่ในข้อความ Word แต่ละข้อได้โดยตรง หากข้อความใดผูกไว้หลายฟังก์ชัน **ขอเพียงใน Excel มีตรงกัน 1 ข้อขึ้นไป ระบบจะคงแถวใน Word ไว้เสมอครับ**")
                
                curr_mapping = iq_inst_svc.load_tag_mapping()
                protocols_dict = curr_mapping.get("protocols", {})
                
                # Load available functions from Excel to use as selectable chips
                sample_excel_inst = os.path.abspath(r"IQOQDQ/IQ_Installation/Project_XXXXX.xlsx")
                if not os.path.exists(sample_excel_inst):
                    sample_excel_inst = os.path.abspath(r"IQOQDQ/Data IQOQDQ/Project_XXXXX.xlsx")
                
                avail_excel_funcs = ["Hold"]
                if os.path.exists(sample_excel_inst):
                    try:
                        p_excel = iq_inst_svc.parse_project_excel(sample_excel_inst)
                        avail_excel_funcs += sorted(list(p_excel["functions_df"]["function_name"].dropna().unique()))
                    except Exception:
                        pass
                
                # Declare Drag & Drop custom component
                tag_drag_drop_comp_dir = os.path.abspath(r"backend/tag_drag_drop_component")
                tag_drag_drop_component = st.components.v1.declare_component("tag_drag_drop_board", path=tag_drag_drop_comp_dir)

                col_proto_sel, col_mode_sel = st.columns([1.5, 1.8])
                with col_proto_sel:
                    selected_edit_proto = st.radio(
                        "เลือก Protocol ที่ต้องการจับคู่ Tag",
                        ["Cartoner", "Tube Filler", "Tube Filler FP 10"],
                        horizontal=True,
                        key="iq_inst_edit_proto_sel"
                    )
                with col_mode_sel:
                    tag_view_mode = st.radio(
                        "โหมดการแก้ไข",
                        ["🖱️ ลากและวาง Tag ด้วยเมาส์ (Drag & Drop)", "🎯 ตัวเลือก Multi-Select", "📝 ตารางรวม"],
                        horizontal=True,
                        key="iq_inst_tag_view_mode"
                    )
                
                proto_tags = protocols_dict.get(selected_edit_proto, {}).get("tags", [])
                
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                
                if tag_view_mode == "🖱️ ลากและวาง Tag ด้วยเมาส์ (Drag & Drop)":
                    st.caption("🖱️ **วิธีใช้งาน:** ลากชิป Function จาก 'คลัง Function' ด้านบนลงในกล่องข้อความ Word หรือลากย้ายข้ามข้อความไปมาได้อย่างอิสระ แล้วกดปุ่มบันทึกด้านบน")
                    dnd_res = tag_drag_drop_component(
                        protocolName=selected_edit_proto,
                        tags=proto_tags,
                        allExcelFuncs=avail_excel_funcs,
                        key=f"iq_inst_dnd_{selected_edit_proto}"
                    )
                    if dnd_res and isinstance(dnd_res, dict) and dnd_res.get("action") == "save_mappings":
                        up_proto = dnd_res.get("protocolName")
                        up_tags = dnd_res.get("tags")
                        if up_proto and up_tags is not None:
                            curr_mapping["protocols"][up_proto]["tags"] = up_tags
                            iq_inst_svc.save_tag_mapping(curr_mapping)
                            st.toast(f"✅ บันทึก Tag Mappings สำหรับ {up_proto} สำเร็จแล้ว!", icon="💾")
                            st.rerun()

                elif tag_view_mode == "🎯 ตัวเลือก Multi-Select":
                    col_top_save, col_top_clear = st.columns([1.5, 3])
                    with col_top_save:
                        save_all_visual = st.button("💾 บันทึกการจับคู่ Tag ทั้งหมด (Save All Mappings)", type="primary", key="iq_inst_save_visual_btn", use_container_width=True)
                    with col_top_clear:
                        if st.button("🧹 ล้าง Tag ทั้งหมดของ Protocol นี้เพื่อจับคู่ใหม่ (Clear All)", key="iq_inst_clear_visual_btn"):
                            for t in proto_tags:
                                t["excel_keywords"] = []
                            curr_mapping["protocols"][selected_edit_proto]["tags"] = proto_tags
                            iq_inst_svc.save_tag_mapping(curr_mapping)
                            st.rerun()

                    # Group tags by category
                    tags_by_cat = {}
                    for t in proto_tags:
                        cat = t.get("category", "General")
                        if cat not in tags_by_cat:
                            tags_by_cat[cat] = []
                        tags_by_cat[cat].append(t)
                    
                    updated_proto_tags = []
                    
                    for cat_name, cat_tags in tags_by_cat.items():
                        st.markdown(f"#### 📁 หมวดหมู่: `{cat_name}`")
                        for t in cat_tags:
                            t_id = t.get("id", "")
                            t_text = t.get("test_point_text", "")
                            t_keywords = t.get("excel_keywords", [])
                            t_mand = t.get("is_mandatory", False)
                            
                            # Merge available options with current keywords so everything is selectable
                            all_options = list(dict.fromkeys(["Hold"] + t_keywords + avail_excel_funcs))
                            
                            with st.container():
                                c_title, c_picker, c_mand = st.columns([1.8, 3, 0.8])
                                with c_title:
                                    st.markdown(f"**📄 ข้อความใน Word:**")
                                    st.markdown(f"`{t_text}`")
                                    is_hold_selected = any(k.strip().lower() == "hold" for k in t_keywords)
                                    if is_hold_selected:
                                        st.caption("🟠 **สถานะ: Hold** *(พักไว้สำหรับอนาคต - ลบแถวออก)*")
                                    elif t_keywords:
                                        st.caption(f"🟢 ผูกไว้ **{len(t_keywords)}** ฟังก์ชัน")
                                    else:
                                        st.caption("⚪ *ยังไม่ได้ผูก Tag (จะถูกลบแถวออก)*")
                                
                                with c_picker:
                                    selected_funcs = st.multiselect(
                                        f"เลือก/ลบ Function Name จาก Excel (สำหรับ: {t_text})",
                                        options=all_options,
                                        default=t_keywords,
                                        key=f"tag_pick_{selected_edit_proto}_{t_id}",
                                        label_visibility="collapsed",
                                        placeholder="🔍 เลือก Function จาก Excel หรือเลือก 'Hold' เพื่อพักไว้..."
                                    )
                                
                                with c_mand:
                                    is_mandatory_val = st.checkbox(
                                        "คงไว้เสมอ (Mandatory)",
                                        value=t_mand,
                                        key=f"tag_mand_{selected_edit_proto}_{t_id}"
                                    )
                                
                                updated_proto_tags.append({
                                    "id": t_id,
                                    "category": cat_name,
                                    "test_point_text": t_text,
                                    "excel_keywords": selected_funcs,
                                    "is_mandatory": is_mandatory_val
                                })
                                st.markdown("<hr style='margin: 8px 0; border-color: #f0f2f6;'>", unsafe_allow_html=True)
                    
                    if save_all_visual:
                        curr_mapping["protocols"][selected_edit_proto]["tags"] = updated_proto_tags
                        iq_inst_svc.save_tag_mapping(curr_mapping)
                        st.success(f"🎉 บันทึกการจับคู่ Tag สำหรับ **{selected_edit_proto}** สำเร็จแล้ว!")
                        st.rerun()

                else:
                    # Data Table Editor Mode
                    table_rows = []
                    for t in proto_tags:
                        table_rows.append({
                            "ID": t.get("id", ""),
                            "Category": t.get("category", ""),
                            "Word Test Point Text": t.get("test_point_text", ""),
                            "Excel Function Keywords (คั่นด้วยเครื่องหมาย ,)": ", ".join(t.get("excel_keywords", [])),
                            "Mandatory (คงไว้เสมอ)": t.get("is_mandatory", False)
                        })
                    
                    tag_df = pd.DataFrame(table_rows)
                    
                    st.markdown("###### 📝 ตาราง Tag Mapping:")
                    edited_df = st.data_editor(
                        tag_df,
                        use_container_width=True,
                        num_rows="dynamic",
                        key=f"iq_inst_tag_editor_{selected_edit_proto}"
                    )
                    
                    if st.button("💾 บันทึกตาราง Tag Mappings (Save to JSON)", type="primary", key="iq_inst_save_table_tags_btn"):
                        updated_tags = []
                        for _, r in edited_df.iterrows():
                            kw_list = [k.strip() for k in str(r["Excel Function Keywords (คั่นด้วยเครื่องหมาย ,)"]).split(",") if k.strip()]
                            updated_tags.append({
                                "id": str(r["ID"]).strip(),
                                "category": str(r["Category"]).strip(),
                                "test_point_text": str(r["Word Test Point Text"]).strip(),
                                "excel_keywords": kw_list,
                                "is_mandatory": bool(r["Mandatory (คงไว้เสมอ)"])
                            })
                        
                        curr_mapping["protocols"][selected_edit_proto]["tags"] = updated_tags
                        iq_inst_svc.save_tag_mapping(curr_mapping)
                        st.success(f"✅ บันทึก Tag Mappings สำหรับ **{selected_edit_proto}** เรียบร้อยแล้ว!")
            
        with tab_iq_cci:
            st.markdown("#### 🛡️ IQ Control Components & Instruments (IQ CCI) Protocol Generator")
            st.caption("Parse Parts List Excel ('IQOQ list*'), extract CIS (Col B), Function/Designation (Col F), Manufacturer/Supplier (Col H), and populate Table 6 in the official Word template.")
            
            import backend.iq_cci_service as iq_cci_svc
            try:
                importlib.reload(iq_cci_svc)
            except Exception:
                pass
            
            col_excel_cci, col_word_cci = st.columns(2)
            
            with col_excel_cci:
                st.markdown("##### 📊 1. Excel Parts List Data Source")
                sample_excel_cci = os.path.abspath(r"IQOQDQ/IQ_CCI/5XXXX_Part list.xlsx")
                sample_excel_bytes_cci = b""
                if os.path.exists(sample_excel_cci):
                    with open(sample_excel_cci, "rb") as sf:
                        sample_excel_bytes_cci = sf.read()
                
                uploaded_excel_cci = st.file_uploader(
                    "Upload Custom Parts List Excel (*.xlsx, *.xls)",
                    type=["xlsx", "xls"],
                    key="iq_cci_upload_excel",
                    help="Upload Excel file containing sheet starting with 'IQOQ list' (Cols B, F, H)"
                )
                
                if uploaded_excel_cci is not None:
                    excel_input_cci = uploaded_excel_cci
                    st.caption(f"📂 Active File: `{uploaded_excel_cci.name}`")
                elif os.path.exists(sample_excel_cci):
                    excel_input_cci = sample_excel_cci
                    st.caption(f"💡 *Defaulting to sample:* `{os.path.basename(sample_excel_cci)}`")
                else:
                    excel_input_cci = None
                
                if sample_excel_bytes_cci:
                    st.download_button(
                        label="📥 Download Sample Excel (5XXXX_Part list.xlsx)",
                        data=sample_excel_bytes_cci,
                        file_name="5XXXX_Part list.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="iq_cci_dl_sample_excel",
                        use_container_width=True
                    )
            
            with col_word_cci:
                st.markdown("##### 📄 2. Word Master Template")
                tmpl_dir_cci = os.path.abspath(r"IQOQDQ/IQ_CCI")
                os.makedirs(tmpl_dir_cci, exist_ok=True)
                
                tmpl_files_cci = [f for f in os.listdir(tmpl_dir_cci) if not f.startswith("~$") and f.lower().endswith(('.doc', '.docx'))]
                word_tmpl_path_cci = None
                
                if tmpl_files_cci:
                    default_tmpl_idx_cci = 0
                    for idx, fn in enumerate(tmpl_files_cci):
                        if "05_iq_control" in fn.lower():
                            default_tmpl_idx_cci = idx
                            break
                    selected_tmpl_cci = st.selectbox("Select Word Master Template", tmpl_files_cci, index=default_tmpl_idx_cci, key="iq_cci_tmpl_sel")
                    word_tmpl_path_cci = os.path.abspath(os.path.join(tmpl_dir_cci, selected_tmpl_cci))
                    st.caption(f"📄 Master Template: `{selected_tmpl_cci}` *(from IQ_CCI)*")
                    
                    # 1. Download current active template
                    if os.path.exists(word_tmpl_path_cci):
                        with open(word_tmpl_path_cci, "rb") as tf:
                            tmpl_bytes_cci = tf.read()
                        is_doc_ext_cci = selected_tmpl_cci.lower().endswith('.doc')
                        mime_cci = "application/msword" if is_doc_ext_cci else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        st.download_button(
                            label=f"📥 Download Current Master Template ({selected_tmpl_cci})",
                            data=tmpl_bytes_cci,
                            file_name=selected_tmpl_cci,
                            mime=mime_cci,
                            key="iq_cci_dl_current_tmpl",
                            use_container_width=True
                        )
                
                # 2. Upload new template to IQ_CCI directory
                uploaded_tmpl_cci = st.file_uploader(
                    "📤 Upload New Master Template (*.docx, *.doc)",
                    type=["docx", "doc"],
                    key="iq_cci_upload_word",
                    help="Upload a Word master template to save into IQ_CCI directory"
                )
                if uploaded_tmpl_cci is not None:
                    temp_tmpl_p = os.path.abspath(os.path.join(tmpl_dir_cci, uploaded_tmpl_cci.name))
                    with open(temp_tmpl_p, "wb") as f:
                        f.write(uploaded_tmpl_cci.getvalue())
                    word_tmpl_path_cci = temp_tmpl_p
                    st.success(f"✅ Saved new template `{uploaded_tmpl_cci.name}` to IQ_CCI directory successfully!")
                    st.rerun()

            st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
            
            # Action Button
            btn_gen_col1, btn_gen_col2 = st.columns([1.2, 2])
            with btn_gen_col1:
                process_btn_cci = st.button("⚡ Populate & Generate IQ CCI Word Document", type="primary", use_container_width=True, key="iq_cci_process_btn")
            
            if process_btn_cci:
                if not excel_input_cci:
                    st.error("❌ Please select or upload an Excel Parts List file.")
                else:
                    with st.spinner("Parsing 'IQOQ list*' sheet and populating Table 6 in Word document..."):
                        try:
                            result_cci = iq_cci_svc.generate_iq_cci_word(
                                excel_source=excel_input_cci,
                                word_template_path=word_tmpl_path_cci,
                                row_height_pt=22.0
                            )
                            st.session_state["iq_cci_last_result"] = result_cci
                            st.success(f"🎉 **Word Document Ready!** Generated `{result_cci['file_name']}` with {result_cci['total_items']} control components.")
                        except Exception as ex:
                            st.error(f"❌ Error generating IQ CCI document: {str(ex)}")
            
            # Display Results
            if "iq_cci_last_result" in st.session_state and st.session_state["iq_cci_last_result"]:
                res_cci = st.session_state["iq_cci_last_result"]
                
                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("📄 Target Sheet", res_cci.get("sheet_name", "IQOQ list"))
                m2.metric("📋 Total Components", f"{res_cci['total_items']} Items")
                m3.metric("🏷️ Order No.", res_cci["metadata"].get("order_no", "-"))
                m4.metric("⚙️ Model", res_cci["metadata"].get("model", "-"))
                
                st.download_button(
                    label=f"📥 Download {res_cci['file_name']}",
                    data=res_cci["doc_bytes"],
                    file_name=res_cci["file_name"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    key="iq_cci_dl_btn",
                    use_container_width=True
                )
                
                with st.expander("👀 View Extracted Control Components (Preview Data)", expanded=True):
                    st.dataframe(res_cci["preview_df"], use_container_width=True, hide_index=True)
            
        with tab_iq_fmt:
            st.markdown("#### 📦 IQ Format Parts Protocol Generator")
            st.caption("Automatically parse Excel BOM data, apply column/row filtering rules, group by component designation, and populate the official IWK Word document.")
            
            import backend.iq_format_service as iq_fmt_svc
            try:
                importlib.reload(iq_fmt_svc)
            except Exception:
                pass
            
            # File Selection Layout
            col_excel, col_word = st.columns(2)
            
            with col_excel:
                st.markdown("##### 📊 1. Excel Data Source")
                
                sample_excel_path = os.path.abspath(r"IQOQDQ/IQ_Format/EXPORT_20260901_151220.XLSX")
                if not os.path.exists(sample_excel_path):
                    sample_excel_path = os.path.abspath(r"IQOQDQ/Data IQOQDQ/EXPORT_20260901_151220.XLSX")
                
                sample_excel_bytes = b""
                if os.path.exists(sample_excel_path):
                    with open(sample_excel_path, "rb") as sf:
                        sample_excel_bytes = sf.read()
                
                uploaded_excel = st.file_uploader(
                    "Upload Custom Excel File (CSP2/FORMAT1)",
                    type=["xlsx", "xls"],
                    key="iq_fmt_upload_excel",
                    help="Upload your Excel BOM containing format columns (Funktion Zeile1)"
                )
                
                if uploaded_excel is not None:
                    excel_input = uploaded_excel
                    st.caption(f"📂 Active File: `{uploaded_excel.name}`")
                elif os.path.exists(sample_excel_path):
                    excel_input = sample_excel_path
                    st.caption(f"💡 *Defaulting to sample:* `{os.path.basename(sample_excel_path)}`")
                else:
                    excel_input = None
                
                if sample_excel_bytes:
                    st.download_button(
                        label="📥 Download Sample Excel (EXPORT_20260901_151220.XLSX)",
                        data=sample_excel_bytes,
                        file_name="EXPORT_20260901_151220.XLSX",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="iq_fmt_dl_sample_excel",
                        use_container_width=True
                    )
            
            with col_word:
                st.markdown("##### 📄 2. Word Master Template")
                tmpl_dir = os.path.abspath(r"IQOQDQ/IQ_Format")
                if not os.path.exists(tmpl_dir) or not [f for f in os.listdir(tmpl_dir) if not f.startswith("~$") and f.lower().endswith(('.doc', '.docx'))]:
                    tmpl_dir = os.path.abspath(r"IQOQDQ/Data IQOQDQ")
                
                tmpl_files = [f for f in os.listdir(tmpl_dir) if not f.startswith("~$") and f.lower().endswith(('.doc', '.docx'))] if os.path.exists(tmpl_dir) else []
                word_tmpl_path = None
                if tmpl_files:
                    default_tmpl_idx = 0
                    for idx, fn in enumerate(tmpl_files):
                        if "5xxxx_08_iq_format" in fn.lower():
                            default_tmpl_idx = idx
                            break
                    selected_tmpl = st.selectbox("Select Word Template (.doc / .docx)", tmpl_files, index=default_tmpl_idx, key="iq_fmt_tmpl_sel")
                    word_tmpl_path = os.path.abspath(os.path.join(tmpl_dir, selected_tmpl))
                    st.caption(f"📄 Master Template: `{selected_tmpl}` *(from IQ_Format)*")
                    
                    # Download current active template
                    if os.path.exists(word_tmpl_path):
                        with open(word_tmpl_path, "rb") as tf:
                            tmpl_bytes_fmt = tf.read()
                        is_doc_ext_fmt = selected_tmpl.lower().endswith('.doc')
                        mime_fmt = "application/msword" if is_doc_ext_fmt else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        st.download_button(
                            label=f"📥 Download Current Master Template ({selected_tmpl})",
                            data=tmpl_bytes_fmt,
                            file_name=selected_tmpl,
                            mime=mime_fmt,
                            key="iq_fmt_dl_current_tmpl",
                            use_container_width=True
                        )

                # Upload new template
                uploaded_tmpl = st.file_uploader("📤 Upload New Master Template (*.docx, *.doc)", type=["doc", "docx"], key="iq_fmt_upload_word")
                if uploaded_tmpl:
                    temp_tmpl_path = os.path.abspath(os.path.join(tmpl_dir, uploaded_tmpl.name))
                    with open(temp_tmpl_path, "wb") as f:
                        f.write(uploaded_tmpl.getvalue())
                    word_tmpl_path = temp_tmpl_path
                    st.success(f"✅ Saved new template `{uploaded_tmpl.name}` to IQ_Format directory successfully!")
                    st.rerun()

            st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
            
            # 3. Pre-generation Content Review & Data Verification
            st.markdown("##### 📝 3. Pre-download Content Preview & Data Verification")
            st.caption("ตรวจสอบและแก้ไขข้อมูลเนื้อหา Format Parts ก่อนทำการสร้างและดาวน์โหลดเอกสาร Word")
            
            # Load Data Preview before downloading
            parsed_df = None
            fn_tags = []
            if excel_input:
                try:
                    init_groups, parsed_df, fn_tags = iq_fmt_svc.parse_and_clean_format_excel(excel_input)
                except Exception as parse_err:
                    st.warning(f"⚠️ Unable to parse Excel preview: {str(parse_err)}")
                    parsed_df = None
                    fn_tags = []
            
            if parsed_df is not None and not parsed_df.empty:
                st.markdown("<div style='background-color: #eef6ff; padding: 12px 16px; border-radius: 8px; border-left: 4px solid #1e88e5; margin: 10px 0;'>", unsafe_allow_html=True)
                st.markdown("💡 **Format Codes (Fn) Formatting Note:** หากรายการใดมี Format Code หลายตัว (เช่น `F2 F5 F6`) ระบบจะทำการจัดฟอร์แมตคั่นด้วยเครื่องหมายจุลภาคให้อัตโนมัติเป็น `F2, F5, F6` ก่อนนำเข้าสู่เอกสาร Word", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)
                
                # Tags & Summary metrics before download
                s_col1, s_col2, s_col3 = st.columns([1, 1, 2])
                s_col1.metric("📂 Total Component Groups", f"{len(parsed_df['Category / Group'].unique())} Groups")
                s_col2.metric("📋 Total Format Parts", f"{len(parsed_df)} Items")
                with s_col3:
                    st.markdown("**🏷️ Format Tags Identified:**")
                    if fn_tags:
                        tags_html = " ".join([f"<span style='background-color:#0d6efd; color:white; padding:3px 8px; border-radius:12px; font-size:12px; margin-right:4px; display:inline-block;'>{t}</span>" for t in fn_tags])
                        st.markdown(tags_html, unsafe_allow_html=True)
                    else:
                        st.caption("No specific Fn tags found")
                
                with st.expander("🔍 Extracted & Grouped Format Parts Data (Data Editor / Preview)", expanded=True):
                    edited_format_df = st.data_editor(
                        parsed_df,
                        use_container_width=True,
                        hide_index=True,
                        num_rows="dynamic",
                        key="iq_fmt_data_editor"
                    )
            else:
                edited_format_df = None
                st.info("ℹ️ Upload or select an Excel Data Source to preview and verify format parts content.")
            
            st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
            
            # Action Button
            btn_col1, btn_col2 = st.columns([1.5, 2])
            with btn_col1:
                process_btn = st.button("⚡ Populate & Generate Word Document", type="primary", use_container_width=True, key="iq_fmt_process_btn")
                
            if process_btn:
                if not excel_input:
                    st.error("❌ Please select or upload an Excel data file first.")
                else:
                    with st.spinner("Opening original Word template, populating Table 4 with formatted Fn data, and preparing download..."):
                        try:
                            result = iq_fmt_svc.generate_iq_format_word(
                                excel_source=excel_input,
                                word_template_path=word_tmpl_path,
                                edited_df=edited_format_df
                            )
                            st.session_state["iq_fmt_last_result"] = result
                            st.success(f"🎉 **Word Document Ready!** Generated `{result['file_name']}`. Review details and download below.")
                        except Exception as ex:
                            st.error(f"❌ Error generating document: {str(ex)}")
            
            # Display Generated Results & Download Button
            if "iq_fmt_last_result" in st.session_state and st.session_state["iq_fmt_last_result"]:
                res = st.session_state["iq_fmt_last_result"]
                
                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                st.markdown("##### 📥 Ready for Download")
                # Metrics Row
                m_col1, m_col2, m_col3 = st.columns(3)
                m_col1.metric("📂 Component Groups", f"{res['total_groups']} Groups")
                m_col2.metric("📋 Format Items", f"{res['total_items']} Items")
                m_col3.metric("📄 File Name", res["file_name"])
                
                # Download Button
                is_doc = res["file_name"].lower().endswith(".doc")
                mime_type = "application/msword" if is_doc else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                
                st.download_button(
                    label=f"📥 Download {res['file_name']} (Original Layout)",
                    data=res["doc_bytes"],
                    file_name=res["file_name"],
                    mime=mime_type,
                    type="primary",
                    key="iq_fmt_dl_btn",
                    use_container_width=True
                )


        
    elif "OQ" in sub_section:
        st.markdown("### ⚡ OQ — Operational Qualification")
        st.caption("Operational test protocols, IO verification, safety interlocks, HMI checks, alarm limits, and shift register tracking.")
        
        tab_oq_io, tab_oq_hmi, tab_oq_alarm, tab_oq_shift = st.tabs([
            "📋 OQ IO List",
            "🖥️ OQ HMI",
            "🚨 OQ Alarm",
            "🔄 OQ Shift Register"
        ])
        
        with tab_oq_io:
            st.markdown("#### 📋 OQ IO List Protocol Generator")
            st.caption("Automatically parse ELCAD IO Lists, filter blanks, sort by Column E (cross ref.), compute Address / Description / Page / Test formulas, and populate Table 5 in the official Word template.")
            
            import backend.oq_io_service as oq_io_svc
            try:
                importlib.reload(oq_io_svc)
            except Exception:
                pass
            
            col_excel_io, col_word_io = st.columns(2)
            
            with col_excel_io:
                st.markdown("##### 📊 1. ELCAD IO List Data Source")
                io_dir = os.path.abspath(r"IQOQDQ/IO list")
                sample_io_files = [f for f in os.listdir(io_dir) if f.lower().startswith("5xxxx_iolist") and f.lower().endswith(('.xls', '.xlsx'))] if os.path.exists(io_dir) else []
                
                selected_sample_io = None
                if sample_io_files:
                    selected_sample_io = st.selectbox("Select Sample IO List", sample_io_files, index=0, key="oq_io_sample_sel")
                    sample_io_path = os.path.join(io_dir, selected_sample_io)
                else:
                    sample_io_path = os.path.join(io_dir, "5XXXX_IOList_1.xls")
                
                sample_io_bytes = b""
                if os.path.exists(sample_io_path):
                    with open(sample_io_path, "rb") as sf:
                        sample_io_bytes = sf.read()
                
                uploaded_excel_io = st.file_uploader(
                    "Upload Custom IO List Excel (*.xls, *.xlsx)",
                    type=["xls", "xlsx"],
                    key="oq_io_upload_excel",
                    help="Upload ELCAD IO list containing columns A..E (module, sybolic, address, comment, cross ref.)"
                )
                
                if uploaded_excel_io is not None:
                    excel_input_io = uploaded_excel_io
                    st.caption(f"📂 Active File: `{uploaded_excel_io.name}`")
                elif os.path.exists(sample_io_path):
                    excel_input_io = sample_io_path
                    st.caption(f"💡 *Defaulting to sample:* `{os.path.basename(sample_io_path)}`")
                else:
                    excel_input_io = None
                
                if sample_io_bytes and selected_sample_io:
                    st.download_button(
                        label=f"📥 Download Sample Excel ({selected_sample_io})",
                        data=sample_io_bytes,
                        file_name=selected_sample_io,
                        mime="application/vnd.ms-excel" if selected_sample_io.endswith('.xls') else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="oq_io_dl_sample_excel",
                        use_container_width=True
                    )
            
            with col_word_io:
                st.markdown("##### 📄 2. Word Master Template")
                tmpl_dir_io = os.path.abspath(r"IQOQDQ/IO list")
                os.makedirs(tmpl_dir_io, exist_ok=True)
                tmpl_files_io = [f for f in os.listdir(tmpl_dir_io) if not f.startswith("~$") and f.lower().endswith(('.doc', '.docx'))] if os.path.exists(tmpl_dir_io) else []
                word_tmpl_path_io = None
                if tmpl_files_io:
                    default_tmpl_idx_io = 0
                    for idx, fn in enumerate(tmpl_files_io):
                        if "03_oq_io" in fn.lower():
                            default_tmpl_idx_io = idx
                            break
                    selected_tmpl_io = st.selectbox("Select Word Master Template", tmpl_files_io, index=default_tmpl_idx_io, key="oq_io_tmpl_sel")
                    word_tmpl_path_io = os.path.abspath(os.path.join(tmpl_dir_io, selected_tmpl_io))
                    st.caption(f"📄 Master Template: `{selected_tmpl_io}` *(from IO list)*")
                    
                    if os.path.exists(word_tmpl_path_io):
                        with open(word_tmpl_path_io, "rb") as tf:
                            tmpl_bytes_io = tf.read()
                        is_doc_ext_io = selected_tmpl_io.lower().endswith('.doc')
                        mime_io = "application/msword" if is_doc_ext_io else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        st.download_button(
                            label=f"📥 Download Current Master Template ({selected_tmpl_io})",
                            data=tmpl_bytes_io,
                            file_name=selected_tmpl_io,
                            mime=mime_io,
                            key="oq_io_dl_current_tmpl",
                            use_container_width=True
                        )
                
                uploaded_tmpl_io = st.file_uploader(
                    "📤 Upload New Master Template (*.docx, *.doc)",
                    type=["docx", "doc"],
                    key="oq_io_upload_word",
                    help="Upload a Word master template to save into IO list directory"
                )
                if uploaded_tmpl_io is not None:
                    temp_tmpl_p = os.path.abspath(os.path.join(tmpl_dir_io, uploaded_tmpl_io.name))
                    with open(temp_tmpl_p, "wb") as f:
                        f.write(uploaded_tmpl_io.getvalue())
                    word_tmpl_path_io = temp_tmpl_p
                    st.success(f"✅ Saved new template `{uploaded_tmpl_io.name}` to IO list directory successfully!")
                    st.rerun()

            st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
            
            # Action Button
            btn_gen_col1, btn_gen_col2 = st.columns([1.2, 2])
            with btn_gen_col1:
                process_btn_io = st.button("⚡ Populate & Generate OQ IO List Word Document", type="primary", use_container_width=True, key="oq_io_process_btn")
            
            if process_btn_io:
                if not excel_input_io:
                    st.error("❌ Please select or upload an IO List Excel file.")
                else:
                    with st.spinner("Sorting by Column E, calculating formulas, and populating Table 5 in Word document..."):
                        try:
                            result_io = oq_io_svc.generate_oq_io_word(
                                excel_source=excel_input_io,
                                word_template_path=word_tmpl_path_io,
                                row_height_pt=22.0
                            )
                            st.session_state["oq_io_last_result"] = result_io
                            st.success(f"🎉 **Word Document Ready!** Generated `{result_io['file_name']}` with {result_io['total_items']} IO points.")
                        except Exception as ex:
                            st.error(f"❌ Error generating OQ IO List document: {str(ex)}")
            
            # Display Results
            if "oq_io_last_result" in st.session_state and st.session_state["oq_io_last_result"]:
                res_io = st.session_state["oq_io_last_result"]
                
                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("📋 Total IO Points", f"{res_io['total_items']} Items")
                m2.metric("⚡ T1 (Outputs)", f"{res_io['t1_count']} Items")
                m3.metric("🔍 T2 (Inputs)", f"{res_io['t2_count']} Items")
                m4.metric("📄 Ready for Download", res_io['file_name'])
                
                st.download_button(
                    label=f"📥 Download {res_io['file_name']}",
                    data=res_io["doc_bytes"],
                    file_name=res_io["file_name"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    key="oq_io_dl_btn",
                    use_container_width=True
                )
                
                with st.expander("👀 View Computed & Sorted IO List Records (Preview Data)", expanded=True):
                    st.dataframe(res_io["preview_df"], use_container_width=True, hide_index=True)
            

        with tab_oq_hmi:
            import importlib
            import backend.oq_hmi_service as oq_hmi_svc
            try:
                importlib.reload(oq_hmi_svc)
            except Exception:
                pass

            tab_hmi_gen, tab_hmi_maint = st.tabs([
                "⚡ OQ HMI Protocol Generator",
                "⚙️ Machine Type & Template Maintenance"
            ])

            with tab_hmi_gen:
                st.markdown("#### 🖥️ OQ HMI Protocol Generator (OCR Image System)")
                st.caption("Automatically run Image OCR on screenshot images, match templates by Machine & Visu Type, insert page breaks, and replace `XXXX` placeholder under `3.3 Masks` in Word master template.")

                # Configuration Row: Select Machine Type & Screen/Visu Type
                cfg_hmi_col1, cfg_hmi_col2 = st.columns(2)
                with cfg_hmi_col1:
                    available_hmi_machines = oq_hmi_svc.get_available_hmi_machine_types()
                    default_hmi_m_idx = available_hmi_machines.index("FP") if "FP" in available_hmi_machines else 0
                    selected_hmi_machine = st.selectbox(
                        "🏭 Select Machine Type",
                        available_hmi_machines,
                        index=default_hmi_m_idx,
                        key="oq_hmi_machine_sel",
                        help="Target machine type folder (e.g. FP, SC_SI, TZC, FC, TL, Standard)"
                    )
                with cfg_hmi_col2:
                    hmi_screen_type = st.radio(
                        "🖥️ Screen / Visu Type",
                        ["🖥️ IPC (Industrial PC)", "📱 Magilis"],
                        horizontal=True,
                        key="oq_hmi_screen_type",
                        help="Select Visu type"
                    )

                visu_clean = "Magilis" if "Magilis" in hmi_screen_type else "IPC"
                if "Magilis" in hmi_screen_type:
                    st.info("💡 **Magilis Visu Note:** Magilis visu template integration active. System maps IPC and Magilis screen layouts dynamically.")

                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)

                col_img_hmi, col_tmpl_hmi = st.columns(2)

                with col_img_hmi:
                    st.markdown("##### 🖼️ 1. Insert Screenshots / Images")
                    sample_img_dir = os.path.abspath(r"IQOQDQ/OQ_HMI/Screen short")
                    sample_img_count = len([f for f in os.listdir(sample_img_dir) if f.lower().endswith(('.jpg', '.jpeg', '.png'))]) if os.path.exists(sample_img_dir) else 0

                    uploaded_imgs_hmi = st.file_uploader(
                        "Upload Screenshot Images (*.jpg, *.png) or ZIP",
                        type=["jpg", "jpeg", "png", "zip"],
                        accept_multiple_files=True,
                        key="oq_hmi_upload_images",
                        help="Upload screenshot images or a ZIP archive containing HMI screen view images"
                    )

                    if uploaded_imgs_hmi:
                        image_input_hmi = uploaded_imgs_hmi if len(uploaded_imgs_hmi) > 1 or not uploaded_imgs_hmi[0].name.endswith('.zip') else uploaded_imgs_hmi[0]
                        st.caption(f"🖼️ Active Images: `{len(uploaded_imgs_hmi)} file(s) uploaded`")
                    elif os.path.exists(sample_img_dir):
                        image_input_hmi = sample_img_dir
                        st.caption(f"💡 *Defaulting to sample folder:* `Screen short` ({sample_img_count} images)")
                    else:
                        image_input_hmi = None

                with col_tmpl_hmi:
                    st.markdown("##### 📄 2. Word Master Template")
                    tmpl_dir_hmi = os.path.abspath(r"IQOQDQ/OQ_HMI")
                    os.makedirs(tmpl_dir_hmi, exist_ok=True)
                    tmpl_files_hmi = [f for f in os.listdir(tmpl_dir_hmi) if not f.startswith("~$") and f.lower().endswith(('.docx', '.docm', '.doc'))] if os.path.exists(tmpl_dir_hmi) else []

                    word_tmpl_path_hmi = None
                    if tmpl_files_hmi:
                        default_tmpl_idx_hmi = 0
                        for idx, fn in enumerate(tmpl_files_hmi):
                            if "xxxx" in fn.lower():
                                default_tmpl_idx_hmi = idx
                                break
                        selected_tmpl_hmi = st.selectbox("Select Word Master Template", tmpl_files_hmi, index=default_tmpl_idx_hmi, key="oq_hmi_tmpl_sel")
                        word_tmpl_path_hmi = os.path.abspath(os.path.join(tmpl_dir_hmi, selected_tmpl_hmi))
                        st.caption(f"📄 Master Template: `{selected_tmpl_hmi}` *(Target: Replacing XXXX under 3.3 Masks)*")
                        
                        if os.path.exists(word_tmpl_path_hmi):
                            with open(word_tmpl_path_hmi, "rb") as tf:
                                tmpl_bytes_hmi = tf.read()
                            is_doc_ext_hmi = selected_tmpl_hmi.lower().endswith('.doc')
                            mime_hmi = "application/msword" if is_doc_ext_hmi else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            st.download_button(
                                label=f"📥 Download Current Master Template ({selected_tmpl_hmi})",
                                data=tmpl_bytes_hmi,
                                file_name=selected_tmpl_hmi,
                                mime=mime_hmi,
                                key="oq_hmi_dl_current_tmpl",
                                use_container_width=True
                            )

                    uploaded_tmpl_hmi = st.file_uploader(
                        "📤 Upload New Master Template (*.docx, *.docm, *.doc)",
                        type=["docx", "docm", "doc"],
                        key="oq_hmi_upload_word",
                        help="Upload a Word master template to save into OQ_HMI directory"
                    )
                    if uploaded_tmpl_hmi is not None:
                        temp_tmpl_p = os.path.abspath(os.path.join(tmpl_dir_hmi, uploaded_tmpl_hmi.name))
                        with open(temp_tmpl_p, "wb") as f:
                            f.write(uploaded_tmpl_hmi.getvalue())
                        word_tmpl_path_hmi = temp_tmpl_p
                        st.success(f"✅ Saved new template `{uploaded_tmpl_hmi.name}` to OQ_HMI directory successfully!")
                        st.rerun()

                excel_input_hmi = None

                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)

                # Requirement 4 & 5 & 6: Interactive Existing Recheck Window & Screenshot Sequence Control
                if image_input_hmi:
                    eval_list = oq_hmi_svc.get_screenshot_images(image_input_hmi)
                    eval_list_sorted = sorted(eval_list, key=lambda x: oq_hmi_svc.get_category_and_sort_key(x['name']))
                    all_fnames = [img['name'] for img in eval_list_sorted]
                    
                    # Initialize or validate session state image order (Category 1 Operation first)
                    if "oq_hmi_custom_image_order" not in st.session_state or set(st.session_state["oq_hmi_custom_image_order"]) != set(all_fnames):
                        st.session_state["oq_hmi_custom_image_order"] = all_fnames.copy()

                    # Window 1: Existing Recheck with Template Maintenance
                    with st.expander("🖼️ 1. Insert Screenshots / Existing Recheck with Template", expanded=True):
                        st.markdown("##### 🖼️ 1. Existing Recheck with Template Maintenance")
                        st.caption(f"Evaluating uploaded screenshots against maintained template library for **{selected_hmi_machine} ({visu_clean})**...")

                        try:
                            if "oq_hmi_custom_image_order" not in st.session_state or not st.session_state["oq_hmi_custom_image_order"]:
                                st.session_state["oq_hmi_custom_image_order"] = [img['name'] for img in sorted(eval_list, key=lambda x: oq_hmi_svc.get_category_and_sort_key(x['name']))]

                            current_order = st.session_state["oq_hmi_custom_image_order"]

                            eval_list_ordered = [next(img for img in eval_list if img['name'] == fn) for fn in current_order if any(img['name'] == fn for img in eval_list)]

                            ocr_map_preview = []
                            for img_info in eval_list_ordered:
                                fn = img_info['name']
                                ocr_res = oq_hmi_svc.perform_ocr_on_image(img_info)
                                h_title = ocr_res.get('derived_title', fn)
                                ocr_map_preview.append({'fname': fn, 'title': h_title})

                            recheck_df, recheck_metrics = oq_hmi_svc.perform_existing_template_recheck(
                                ocr_map_preview, machine_type=selected_hmi_machine, visu_type=visu_clean
                            )

                            rc_m1, rc_m2, rc_m3 = st.columns(3)
                            rc_m1.metric("📋 Total Evaluated", f"{recheck_metrics['total_evaluated']} Images")
                            rc_m2.metric("✅ Maintained & Matched", f"{recheck_metrics['maintained_count']} Screens")
                            rc_m3.metric("⚠️ Unmaintained / Missing", f"{recheck_metrics['missing_maint_count']} Screens")

                            st.dataframe(recheck_df, use_container_width=True, hide_index=True)
                            if recheck_metrics['missing_maint_count'] > 0:
                                st.warning(f"💡 **Maintenance Alert:** {recheck_metrics['missing_maint_count']} screen(s) are missing from the template library for `{selected_hmi_machine}`. Switch to the '⚙️ Machine Type & Template Maintenance' tab to add them.")
                        except Exception as ex_rc:
                            st.info("💡 Upload screenshots to run live Existing Recheck against Template Library.")

                    # Window 2: Sequence Reordering via Excel Checklist Tool
                    with st.expander("📤 2. นำเข้าไฟล์ Excel Checklist ลำดับใหม่ (Sequence Control)", expanded=True):
                        st.markdown("##### 📤 2. นำเข้าไฟล์ Excel Checklist ลำดับใหม่")
                        st.caption("💡 **วิธีใช้ Checklist Tool:** 1. กดดาวน์โหลดไฟล์ Excel (.xlsm) 2. เปิดไฟล์ ติ๊ก Checkbox เพื่อรับหมายเลขลำดับ (คอลัมน์ C) 3. อัปโหลดไฟล์กลับเพื่อจัดลำดับภาพสำหรับการ insert ในแม่แบบ Word")

                        try:
                            current_order = st.session_state.get("oq_hmi_custom_image_order", all_fnames)
                            eval_list_ordered = [next(img for img in eval_list if img['name'] == fn) for fn in current_order if any(img['name'] == fn for img in eval_list)]
                            ocr_map_preview = []
                            for img_info in eval_list_ordered:
                                fn = img_info['name']
                                ocr_res = oq_hmi_svc.perform_ocr_on_image(img_info)
                                h_title = ocr_res.get('derived_title', fn)
                                ocr_map_preview.append({'fname': fn, 'title': h_title})

                            xl_col1, xl_col2 = st.columns(2)
                            
                            with xl_col1:
                                st.markdown("###### 📥 1. ดาวน์โหลดไฟล์ Excel Checklist Tool")
                                export_rows = []
                                for idx, fn in enumerate(current_order, start=1):
                                    info = next((item for item in ocr_map_preview if item['fname'] == fn), {})
                                    title = info.get('title', fn)
                                    status_str = "N/A"
                                    m = re.search(r'V\d{4}', fn)
                                    fc_code = m.group(0) if m else "General"
                                    export_rows.append({
                                        'Checkbox': '',
                                        'Select': '',
                                        'ลำดับที่เลือก': '',
                                        'File Name': fn,
                                        'Function Code': fc_code,
                                        'Header Title': title,
                                        'Status': status_str
                                    })
                                
                                xl_bytes, xl_fname, xl_mime = oq_hmi_svc.generate_checklist_xlsm(export_rows)

                                st.download_button(
                                    label=f"📥 ดาวน์โหลด Checklist Tool ({xl_fname})",
                                    data=xl_bytes,
                                    file_name=xl_fname,
                                    mime=xl_mime,
                                    use_container_width=True,
                                    key="oq_hmi_btn_dl_seq_excel"
                                )
                                st.caption("✨ *ไฟล์ Excel มี Form Control Checkbox & VBA Macro สำหรับรันหมายเลข 1, 2, 3... n ตามลำดับการติ๊ก*")

                            with xl_col2:
                                st.markdown("###### 📤 2. นำเข้าไฟล์ Excel Checklist ลำดับใหม่")
                                uploaded_seq_xl = st.file_uploader(
                                    "อัปโหลดไฟล์ Excel Checklist ที่เลือกและติ๊กแล้ว (*.xlsm, *.xlsx, *.xls)",
                                    type=["xlsm", "xlsx", "xls"],
                                    key="oq_hmi_uploader_seq_excel"
                                )
                                if uploaded_seq_xl:
                                    try:
                                        xl_bytes = uploaded_seq_xl.getvalue()
                                        final_xl_order, sel_count = oq_hmi_svc.parse_checklist_excel(io.BytesIO(xl_bytes), all_fnames)
                                        if final_xl_order:
                                            # Automatically update custom image order in session state without requiring confirm buttons
                                            if st.session_state.get("oq_hmi_custom_image_order") != final_xl_order:
                                                st.session_state["oq_hmi_custom_image_order"] = list(final_xl_order)
                                                st.rerun()

                                            st.success(f"📋 **ผลการอ่าน Checklist:** อ่านลำดับตามคอลัมน์ C สำเร็จ พบภาพที่เลือกไว้ `{sel_count}` ภาพ (พร้อมใช้งานเรียบร้อยแล้ว)")
                                    except Exception as ex_xl:
                                        st.error(f"❌ ไม่สามารถอ่านไฟล์ Excel ได้: {str(ex_xl)}")

                            # Status Bar
                            st.markdown("---")
                            st.info("💡 **สถานะลำดับภาพ:** ระบบนำลำดับภาพตามหมายเลขใน คอลัมน์ C จากไฟล์ Excel มาใช้งานอัตโนมัติสำหรับการ Generate OQ HMI Word และ Audit Log")
                        except Exception as ex_xl_outer:
                            st.info("💡 Upload screenshots to manage sequence control via Excel Checklist.")

                    # Window 3: Image Insertion Audit Log (Inserted vs Not Inserted)
                    with st.expander("🖼️ 3. Image Insertion Audit Log (Inserted vs Not Inserted)", expanded=True):
                        st.markdown("##### 🖼️ 3. Image Insertion Audit Log (Inserted vs Not Inserted)")
                        st.caption("Detailed breakdown of screenshot image files evaluated, showing insertion status (Inserted vs Skipped vs Not Inserted).")

                        try:
                            if "oq_hmi_last_result" in st.session_state and st.session_state["oq_hmi_last_result"].get("preview_df") is not None:
                                st.dataframe(st.session_state["oq_hmi_last_result"]["preview_df"], use_container_width=True, hide_index=True)
                            elif 'recheck_df' in locals() and not recheck_df.empty:
                                audit_records = []
                                for idx, row in recheck_df.iterrows():
                                    fn = row['Image File Name']
                                    status_maint = str(row['Template Maintenance Status'])
                                    t_title = row['Header / Screen Title']
                                    
                                    if "Maintained & Matched" in status_maint:
                                        ins_status = "✅ Inserted"
                                    elif "Skipped" in status_maint:
                                        ins_status = "⚡ Skipped (Duplicate)"
                                    else:
                                        ins_status = "⚠️ Not Inserted"
                                        
                                    audit_records.append({
                                        "No.": idx + 1,
                                        "Image File Name": fn,
                                        "Insertion Status": ins_status,
                                        "Matched Screen / Protocol Section": t_title if ins_status != "⚠️ Not Inserted" else "N/A"
                                    })
                                audit_preview_df = pd.DataFrame(audit_records)
                                st.dataframe(audit_preview_df, use_container_width=True, hide_index=True)
                            else:
                                st.info("💡 Upload screenshots to view live Image Insertion Audit Log.")
                        except Exception as ex_audit:
                            st.info("💡 Upload screenshots to view live Image Insertion Audit Log.")

                # Action Button (Generate Word)
                btn_hmi_col1, btn_hmi_col2 = st.columns([1.5, 2])
                with btn_hmi_col1:
                    process_btn_hmi = st.button("⚡ Generate OQ HMI Word", type="primary", use_container_width=True, key="oq_hmi_process_btn")

                if process_btn_hmi:
                    if not image_input_hmi:
                        st.error("❌ Please select or upload screenshot images first.")
                    else:
                        with st.spinner(f"นำข้อมูลจาก 🔍 Existing Recheck & Sequence Control มาเปิดไฟล์ย่อยใน '{selected_hmi_machine}', แทรกรูปภาพ และประกอบเข้ากับแม่แบบ Word หลัก..."):
                            try:
                                # Ensure custom_order is derived from uploaded Excel sequence bytes if present
                                uploaded_xl_file = st.session_state.get("oq_hmi_uploader_seq_excel")
                                if uploaded_xl_file is not None:
                                    try:
                                        xl_bytes = uploaded_xl_file.getvalue()
                                        xl_order, _ = oq_hmi_svc.parse_checklist_excel(io.BytesIO(xl_bytes), all_fnames)
                                        if xl_order:
                                            st.session_state["oq_hmi_custom_image_order"] = list(xl_order)
                                    except Exception as ex_xl_gen:
                                        print(f"Warning: Could not re-parse uploaded checklist Excel on generation: {ex_xl_gen}")

                                custom_order = st.session_state.get("oq_hmi_custom_image_order", None)

                                result_hmi = oq_hmi_svc.generate_oq_hmi_word(
                                    excel_source=excel_input_hmi,
                                    word_template_path=word_tmpl_path_hmi,
                                    image_source=image_input_hmi,
                                    machine_type=selected_hmi_machine,
                                    visu_type=visu_clean,
                                    custom_image_order=custom_order
                                )
                                st.session_state["oq_hmi_last_result"] = result_hmi
                                st.success(f"🎉 **Image OCR & Generation Complete!** Successfully inserted {result_hmi['inserted_images_count']} / {result_hmi['total_images']} images into `{result_hmi['file_name']}`.")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"❌ Error during Image OCR processing: {str(ex)}")

                # Display Results & Downloads
                if "oq_hmi_last_result" in st.session_state and st.session_state["oq_hmi_last_result"]:
                    res_hmi = st.session_state["oq_hmi_last_result"]

                    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)

                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("🖥️ Screen Views", f"{res_hmi['total_screens']} Screens")
                    m2.metric("📋 Parameter Items", f"{res_hmi['total_parameters']} Items")
                    m3.metric("🖼️ Inserted Images", f"{res_hmi['inserted_images_count']} / {res_hmi['total_images']}")

                    dup_cnt = res_hmi.get('duplicate_images_count', 0)
                    not_ins = res_hmi.get('not_inserted_images_count', 0)
                    if dup_cnt > 0 and not_ins == 0:
                        m4.metric("⚡ Skipped Duplicates", f"{dup_cnt} Images")
                    else:
                        m4.metric("⚠️ Uninserted Images", f"{not_ins} Images")

                    dl_col1, dl_col2 = st.columns(2)

                    with dl_col1:
                        st.download_button(
                            label=f"📥 Download Populated Word Master Template ({res_hmi['file_name']})",
                            data=res_hmi["doc_bytes"],
                            file_name=res_hmi["file_name"],
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                            key="oq_hmi_dl_word_btn",
                            use_container_width=True,
                            help="Download the completed OQ HMI Word document template populated with test data"
                        )

                    with dl_col2:
                        st.download_button(
                            label=f"📊 Download Image Insertion Audit Excel Report ({res_hmi['excel_audit_name']})",
                            data=res_hmi["excel_audit_bytes"],
                            file_name=res_hmi["excel_audit_name"],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="secondary",
                            key="oq_hmi_dl_excel_audit_btn",
                            use_container_width=True,
                            help="Download Excel report listing all screenshot image names and their insertion status (Inserted vs Not Inserted)"
                        )

                    st.markdown("##### 🖼️ Image Insertion Audit Log (Inserted vs Not Inserted)")
                    st.caption("Detailed breakdown of screenshot image files evaluated, showing which images were inserted into the template and which were not.")
                    st.dataframe(res_hmi["preview_df"], use_container_width=True, hide_index=True)

            # Requirement 7: Tab 2 - Machine Type & Template Maintenance
            with tab_hmi_maint:
                st.markdown("#### ⚙️ Machine Type & Template Maintenance")
                st.caption("Manage machine types, Word master templates, and maintain expected screen headers for each machine and Visu type (IPC / Magilis).")

                maint_hmi_machines = oq_hmi_svc.get_available_hmi_machine_types()

                # Header Bar
                mt_col1, mt_col2, mt_col3 = st.columns([2, 1.2, 1.2])

                with mt_col1:
                    sel_hmi_maint_mach = st.selectbox(
                        "🏭 Active Machine Type to Maintain",
                        maint_hmi_machines,
                        index=0 if maint_hmi_machines else None,
                        key="oq_hmi_maint_active_mach"
                    )

                with mt_col2:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    with st.popover("➕ Add New Machine", use_container_width=True):
                        st.markdown("##### ➕ Create New Machine Type")
                        new_hmi_mach_name = st.text_input("Machine Type Code", placeholder="e.g. FP, TZC, FC, SC_SI", key="oq_hmi_new_mach_name")
                        if st.button("Create & Initialize", type="primary", use_container_width=True, key="oq_hmi_btn_create_mach"):
                            if not new_hmi_mach_name.strip():
                                st.error("❌ Please enter a valid Machine Type code.")
                            else:
                                try:
                                    created_m = oq_hmi_svc.create_hmi_machine_type(new_hmi_mach_name)
                                    st.success(f"🎉 Machine Type `{created_m}` created successfully!")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Error: {str(ex)}")

                with mt_col3:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    with st.popover("🗑️ Delete Machine", use_container_width=True):
                        st.markdown(f"##### ⚠️ Delete Machine Type: `{sel_hmi_maint_mach}`")
                        st.warning(f"This will delete template files and maintenance records for `{sel_hmi_maint_mach}`.")
                        confirm_del_hmi_mach = st.checkbox(f"Yes, delete '{sel_hmi_maint_mach}'", key="oq_hmi_del_mach_chk")
                        if st.button("Delete Machine Type", type="secondary", disabled=not confirm_del_hmi_mach, use_container_width=True, key="oq_hmi_btn_del_mach"):
                            try:
                                oq_hmi_svc.delete_hmi_machine_type(sel_hmi_maint_mach)
                                st.success(f"🗑️ Machine Type `{sel_hmi_maint_mach}` deleted.")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"❌ Error: {str(ex)}")

                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)

                if sel_hmi_maint_mach:
                    # 1. Upload / Dropzone Card (Matching OQ Alarm UI)
                    st.markdown(f"##### 📤 Drop / Upload Word Templates to `{sel_hmi_maint_mach}`")
                    maint_upload_files = st.file_uploader(
                        f"Drop or Select Word Files (*.docx, *.doc) for {sel_hmi_maint_mach}",
                        type=["docx", "doc", "docm"],
                        accept_multiple_files=True,
                        key="oq_hmi_maint_upload_dropzone",
                        help="Upload master screen sub-templates for this machine"
                    )

                    if maint_upload_files:
                        up_c1, up_c2 = st.columns([1.5, 2])
                        with up_c1:
                            if st.button(f"💾 Save & Sync {len(maint_upload_files)} File(s) into {sel_hmi_maint_mach}", type="primary", use_container_width=True, key="oq_hmi_btn_save_uploaded_files"):
                                with st.spinner(f"Saving and synchronizing template files for {sel_hmi_maint_mach}..."):
                                    saved = oq_hmi_svc.upload_hmi_machine_files(sel_hmi_maint_mach, maint_upload_files)
                                    st.success(f"🎉 Successfully saved {len(saved)} file(s) into `{sel_hmi_maint_mach}`!")
                                    st.rerun()

                    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)

                    # 2. File Table and Metrics (Matching OQ Alarm UI)
                    df_maint_files = oq_hmi_svc.list_hmi_machine_files_detailed(sel_hmi_maint_mach)

                    f_m1, f_m2, f_m3 = st.columns(3)
                    f_m1.metric("📂 Total Templates", f"{len(df_maint_files)} Files")
                    tot_kb = df_maint_files["Size (KB)"].sum() if not df_maint_files.empty else 0
                    f_m2.metric("💾 Total Library Size", f"{tot_kb:.1f} KB")
                    f_m3.metric("🏷️ Selected Machine", sel_hmi_maint_mach)

                    st.markdown(f"##### 📋 Existing Template Files in `{sel_hmi_maint_mach}`")
                    if not df_maint_files.empty:
                        st.dataframe(df_maint_files, use_container_width=True, hide_index=True)

                        # File Management Actions (Matching OQ Alarm UI)
                        f_act_c1, f_act_c2 = st.columns(2)

                        with f_act_c1:
                            st.markdown("###### 🗑️ Delete Template File(s)")
                            files_to_remove = st.multiselect(
                                "Select file(s) to remove from library",
                                df_maint_files["Filename"].tolist(),
                                key="oq_hmi_maint_remove_sel",
                                help="Select one or more files to delete"
                            )
                            if files_to_remove:
                                if st.button(f"🗑️ Delete {len(files_to_remove)} Selected File(s)", type="secondary", key="oq_hmi_maint_btn_remove_files"):
                                    deleted_list = oq_hmi_svc.delete_hmi_machine_files(sel_hmi_maint_mach, files_to_remove)
                                    st.success(f"🗑️ Successfully deleted {len(deleted_list)} file(s).")
                                    st.rerun()

                        with f_act_c2:
                            st.markdown("###### 📥 Download Template File")
                            file_to_inspect = st.selectbox(
                                "Select file to download & inspect",
                                df_maint_files["Filename"].tolist(),
                                key="oq_hmi_maint_inspect_sel"
                            )
                            if file_to_inspect:
                                inspect_bytes = oq_hmi_svc.get_hmi_machine_file_bytes(sel_hmi_maint_mach, file_to_inspect)
                                if inspect_bytes:
                                    is_doc_ext = file_to_inspect.lower().endswith('.doc')
                                    mime_ext = "application/msword" if is_doc_ext else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    st.download_button(
                                        label=f"📥 Download `{file_to_inspect}`",
                                        data=inspect_bytes,
                                        file_name=file_to_inspect,
                                        mime=mime_ext,
                                        key="oq_hmi_maint_btn_dl_single",
                                        use_container_width=True
                                    )
                    else:
                        st.info(f"💡 No template files found in `{sel_hmi_maint_mach}`. Use the uploader above to drop and save Word templates.")

            
        with tab_oq_alarm:
            import importlib
            import backend.oq_alarm_service as oq_alarm_svc
            try:
                importlib.reload(oq_alarm_svc)
            except Exception:
                pass
            
            tab_alarm_gen, tab_alarm_maint = st.tabs([
                "⚡ Protocol Generator",
                "⚙️ Machine Type & Template Maintenance"
            ])
            
            with tab_alarm_gen:
                st.markdown("#### 🚨 OQ Alarm Protocol Generator")
                st.caption("Automatically match Alarm Excel variables (e.g. `5XXXX-AlarmInfo.xlsx`) against machine-specific alarm templates (`MX_*`), assemble test protocols into the Word master template, and generate matched alarms summary reports.")
                
                # Machine & Screen Configuration Row
                cfg_col1, cfg_col2 = st.columns(2)
                with cfg_col1:
                    available_machines = oq_alarm_svc.get_available_machine_types()
                    default_m_idx = available_machines.index("FP") if "FP" in available_machines else 0
                    selected_machine = st.selectbox(
                        "🏭 Select Machine Type",
                        available_machines,
                        index=default_m_idx,
                        key="oq_alarm_machine_sel",
                        help="Target machine type folder in GMP_Alarme (e.g. FP, SC_SI, TZC, FC, TL)"
                    )
                with cfg_col2:
                    screen_type = st.radio(
                        "🖥️ Screen / Visu Type",
                        ["🖥️ IPC (Industrial PC)", "📱 Magilis"],
                        horizontal=True,
                        key="oq_alarm_screen_type",
                        help="Magilis file formats will be integrated in future releases"
                    )
                
                if "Magilis" in screen_type:
                    st.info("💡 **Magilis Visu Note:** Magilis file structure parser is scheduled for upcoming release. Current processing uses standard IPC alarm variable mapping.")
                
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                
                col_excel_alarm, col_word_alarm = st.columns(2)
                
                with col_excel_alarm:
                    st.markdown("##### 📊 1. Alarm Information Excel")
                    
                    uploaded_excel_alarm = st.file_uploader(
                        "Upload Custom Alarm Excel (*.xlsx, *.xls)",
                        type=["xlsx", "xls"],
                        key="oq_alarm_upload_excel",
                        help="Upload Excel containing alarm variables in Column A (e.g. FP:.VC_MX_006_MS_FunctionRequired)"
                    )
                    
                    alarm_dir = os.path.abspath(r"IQOQDQ/OQ Alarm")
                    sample_alarm_path = os.path.join(alarm_dir, "5XXXX-AlarmInfo.xlsx")
                    
                    if uploaded_excel_alarm is not None:
                        excel_input_alarm = uploaded_excel_alarm
                        st.caption(f"📂 Active File: `{uploaded_excel_alarm.name}`")
                    elif os.path.exists(sample_alarm_path):
                        excel_input_alarm = sample_alarm_path
                        st.caption(f"💡 *Defaulting to sample:* `{os.path.basename(sample_alarm_path)}`")
                    else:
                        excel_input_alarm = None
                    
                    if os.path.exists(sample_alarm_path):
                        with open(sample_alarm_path, "rb") as sf:
                            sample_alarm_bytes = sf.read()
                        st.download_button(
                            label=f"📥 Download Sample Excel ({os.path.basename(sample_alarm_path)})",
                            data=sample_alarm_bytes,
                            file_name=os.path.basename(sample_alarm_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="oq_alarm_dl_sample_excel",
                            use_container_width=True
                        )
                
                with col_word_alarm:
                    st.markdown("##### 📄 2. Word Master Template")
                    tmpl_dir_alarm = os.path.abspath(r"IQOQDQ/OQ Alarm")
                    os.makedirs(tmpl_dir_alarm, exist_ok=True)
                    tmpl_files_alarm = [f for f in os.listdir(tmpl_dir_alarm) if not f.startswith("~$") and not f.startswith(".") and f.lower().endswith(('.doc', '.docx'))] if os.path.exists(tmpl_dir_alarm) else []
                    word_tmpl_path_alarm = None
                    if tmpl_files_alarm:
                        default_tmpl_idx_alarm = 0
                        for idx, fn in enumerate(tmpl_files_alarm):
                            if "10_oq_alarms" in fn.lower() or "alarms_gmp" in fn.lower():
                                default_tmpl_idx_alarm = idx
                                break
                        selected_tmpl_alarm = st.selectbox("Select Word Master Template", tmpl_files_alarm, index=default_tmpl_idx_alarm, key="oq_alarm_tmpl_sel")
                        word_tmpl_path_alarm = os.path.abspath(os.path.join(tmpl_dir_alarm, selected_tmpl_alarm))
                        st.caption(f"📄 Master Template: `{selected_tmpl_alarm}` *(from OQ Alarm)*")
                        
                        if os.path.exists(word_tmpl_path_alarm):
                            with open(word_tmpl_path_alarm, "rb") as tf:
                                tmpl_bytes_alarm = tf.read()
                            is_doc_ext_alarm = selected_tmpl_alarm.lower().endswith('.doc')
                            mime_alarm = "application/msword" if is_doc_ext_alarm else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            st.download_button(
                                label=f"📥 Download Current Master Template ({selected_tmpl_alarm})",
                                data=tmpl_bytes_alarm,
                                file_name=selected_tmpl_alarm,
                                mime=mime_alarm,
                                key="oq_alarm_dl_current_tmpl",
                                use_container_width=True
                            )

                    uploaded_tmpl_alarm = st.file_uploader(
                        "📤 Upload New Master Template (*.docx, *.doc)",
                        type=["docx", "doc"],
                        key="oq_alarm_upload_word",
                        help="Upload a Word master template to save into OQ Alarm directory"
                    )
                    if uploaded_tmpl_alarm is not None:
                        temp_tmpl_p = os.path.abspath(os.path.join(tmpl_dir_alarm, uploaded_tmpl_alarm.name))
                        with open(temp_tmpl_p, "wb") as f:
                            f.write(uploaded_tmpl_alarm.getvalue())
                        word_tmpl_path_alarm = temp_tmpl_p
                        st.success(f"✅ Saved new template `{uploaded_tmpl_alarm.name}` to OQ Alarm directory successfully!")
                        st.rerun()

                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                
                # Action Button
                btn_gen_col1, btn_gen_col2 = st.columns([1.2, 2])
                with btn_gen_col1:
                    process_btn_alarm = st.button("⚡ Populate & Generate OQ Alarm Word Document", type="primary", use_container_width=True, key="oq_alarm_process_btn")
                
                if process_btn_alarm:
                    if not excel_input_alarm:
                        st.error("❌ Please select or upload an Alarm Excel file.")
                    else:
                        with st.spinner(f"Matching alarm variables against '{selected_machine}' library and assembling Section 3 test protocols..."):
                            try:
                                result_alarm = oq_alarm_svc.generate_oq_alarm_word(
                                    excel_source=excel_input_alarm,
                                    machine_type=selected_machine,
                                    word_template_path=word_tmpl_path_alarm
                                )
                                st.session_state["oq_alarm_last_result"] = result_alarm
                                st.success(f"🎉 **Word Document Ready!** Successfully matched {result_alarm['total_matched']} alarms and assembled `{result_alarm['file_name']}`.")
                            except Exception as ex:
                                st.error(f"❌ Error generating OQ Alarm document: {str(ex)}")
                
                # Display Results
                if "oq_alarm_last_result" in st.session_state and st.session_state["oq_alarm_last_result"]:
                    res_alarm = st.session_state["oq_alarm_last_result"]
                    
                    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("📋 Total Scanned", f"{res_alarm['total_scanned']} Variables")
                    m2.metric("✅ Matched Alarms", f"{res_alarm['total_matched']} Alarms")
                    skipped_count = res_alarm['total_scanned'] - res_alarm['total_matched']
                    m3.metric("⏭️ Skipped / Not Found", f"{skipped_count} Rows")
                    m4.metric("🏷️ Machine Type", res_alarm.get('machine_type', selected_machine))
                    
                    # Dual Download Buttons
                    dl_col1, dl_col2 = st.columns(2)
                    with dl_col1:
                        st.download_button(
                            label=f"📥 Download Word Protocol ({res_alarm['file_name']})",
                            data=res_alarm["doc_bytes"],
                            file_name=res_alarm["file_name"],
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                            key="oq_alarm_dl_doc_btn",
                            use_container_width=True
                        )
                    with dl_col2:
                        st.download_button(
                            label=f"📊 Download Matched Alarms Summary (Excel)",
                            data=res_alarm["summary_excel_bytes"],
                            file_name=res_alarm["summary_excel_name"],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="oq_alarm_dl_excel_btn",
                            use_container_width=True
                        )
                    
                    # Interactive Previews
                    prev_tab1, prev_tab2 = st.tabs([
                        f"✅ Matched Alarms ({res_alarm['total_matched']})",
                        f"🔍 All Scanned Rows ({res_alarm['total_scanned']})"
                    ])
                    with prev_tab1:
                        if not res_alarm["matched_df"].empty:
                            st.dataframe(res_alarm["matched_df"], use_container_width=True, hide_index=True)
                        else:
                            st.warning("No matched alarms found.")
                    with prev_tab2:
                        if not res_alarm["full_df"].empty:
                            st.dataframe(res_alarm["full_df"], use_container_width=True, hide_index=True)
            
            with tab_alarm_maint:
                st.markdown("#### ⚙️ Machine Type & Alarm Template Library Maintenance")
                st.caption("Create new Machine Types, drop / upload Word alarm templates (*.docx, *.doc), and delete obsolete files to keep each machine library up-to-date.")
                
                maint_machines = oq_alarm_svc.get_available_machine_types()
                
                # Machine Type Management Header Bar
                maint_top_col1, maint_top_col2, maint_top_col3 = st.columns([2, 1.2, 1.2])
                
                with maint_top_col1:
                    sel_maint_mach = st.selectbox(
                        "🏭 Active Machine Type to Maintain",
                        maint_machines,
                        index=0 if maint_machines else None,
                        key="oq_maint_active_mach"
                    )
                
                with maint_top_col2:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    with st.popover("➕ Add New Machine", use_container_width=True):
                        st.markdown("##### ➕ Create New Machine Type")
                        new_mach_input = st.text_input("Machine Type Code", placeholder="e.g. TZF, SC_DUO, KARTON", key="oq_new_mach_name")
                        if st.button("Create & Initialize", type="primary", use_container_width=True, key="oq_btn_create_mach"):
                            if not new_mach_input.strip():
                                st.error("❌ Please enter a valid Machine Type name.")
                            else:
                                try:
                                    created_name = oq_alarm_svc.create_machine_type(new_mach_input)
                                    st.success(f"🎉 Machine Type `{created_name}` created and initialized successfully!")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Error creating machine type: {str(ex)}")
                
                with maint_top_col3:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    with st.popover("🗑️ Delete Machine", use_container_width=True):
                        st.markdown(f"##### ⚠️ Delete Machine Type: `{sel_maint_mach}`")
                        st.warning(f"This will delete all template files and cache for `{sel_maint_mach}`.")
                        confirm_del_mach = st.checkbox(f"Yes, delete '{sel_maint_mach}'", key="oq_del_mach_chk")
                        if st.button("Delete Machine Type", type="secondary", disabled=not confirm_del_mach, use_container_width=True, key="oq_btn_del_mach"):
                            try:
                                oq_alarm_svc.delete_machine_type(sel_maint_mach)
                                st.success(f"🗑️ Machine Type `{sel_maint_mach}` has been deleted.")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"❌ Error deleting machine type: {str(ex)}")
                
                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                
                if sel_maint_mach:
                    # Upload / Dropzone Card
                    st.markdown(f"##### 📤 Drop / Upload Word Templates to `{sel_maint_mach}`")
                    maint_upload_files = st.file_uploader(
                        f"Drop or Select Word Files (*.docx, *.doc) for {sel_maint_mach}",
                        type=["docx", "doc"],
                        accept_multiple_files=True,
                        key="oq_maint_upload_dropzone",
                        help="Upload MX_*.doc or MX_*.docx alarm test templates"
                    )
                    
                    if maint_upload_files:
                        up_c1, up_c2 = st.columns([1.5, 2])
                        with up_c1:
                            if st.button(f"💾 Save & Sync {len(maint_upload_files)} File(s) into {sel_maint_mach}", type="primary", use_container_width=True, key="oq_btn_save_uploaded_files"):
                                with st.spinner(f"Saving and synchronizing template files for {sel_maint_mach}..."):
                                    saved = oq_alarm_svc.upload_machine_files(sel_maint_mach, maint_upload_files)
                                    st.success(f"🎉 Successfully saved and synchronized {len(saved)} file(s) into `{sel_maint_mach}`!")
                                    st.rerun()
                    
                    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                    
                    # File Table and Metrics
                    df_maint_files = oq_alarm_svc.list_machine_files_detailed(sel_maint_mach)
                    
                    f_m1, f_m2, f_m3 = st.columns(3)
                    f_m1.metric("📂 Total Templates", f"{len(df_maint_files)} Files")
                    tot_kb = df_maint_files["Size (KB)"].sum() if not df_maint_files.empty else 0
                    f_m2.metric("💾 Total Library Size", f"{tot_kb:.1f} KB")
                    f_m3.metric("🏷️ Selected Machine", sel_maint_mach)
                    
                    st.markdown(f"##### 📋 Existing Template Files in `{sel_maint_mach}`")
                    if not df_maint_files.empty:
                        st.dataframe(df_maint_files, use_container_width=True, hide_index=True)
                        
                        # File Management Actions
                        f_act_c1, f_act_c2 = st.columns(2)
                        
                        with f_act_c1:
                            st.markdown("###### 🗑️ Delete Template File(s)")
                            files_to_remove = st.multiselect(
                                "Select file(s) to remove from library",
                                df_maint_files["Filename"].tolist(),
                                key="oq_maint_remove_sel",
                                help="Select one or more files to delete"
                            )
                            if files_to_remove:
                                if st.button(f"🗑️ Delete {len(files_to_remove)} Selected File(s)", type="secondary", key="oq_maint_btn_remove_files"):
                                    deleted_list = oq_alarm_svc.delete_machine_files(sel_maint_mach, files_to_remove)
                                    st.success(f"🗑️ Successfully deleted {len(deleted_list)} file(s).")
                                    st.rerun()
                        
                        with f_act_c2:
                            st.markdown("###### 📥 Download Template File")
                            file_to_inspect = st.selectbox(
                                "Select file to download & inspect",
                                df_maint_files["Filename"].tolist(),
                                key="oq_maint_inspect_sel"
                            )
                            if file_to_inspect:
                                inspect_bytes = oq_alarm_svc.get_machine_file_bytes(sel_maint_mach, file_to_inspect)
                                if inspect_bytes:
                                    is_doc_ext = file_to_inspect.lower().endswith('.doc')
                                    mime_ext = "application/msword" if is_doc_ext else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    st.download_button(
                                        label=f"📥 Download `{file_to_inspect}`",
                                        data=inspect_bytes,
                                        file_name=file_to_inspect,
                                        mime=mime_ext,
                                        key="oq_maint_btn_dl_single",
                                        use_container_width=True
                                    )
                    else:
                        st.info(f"💡 No template files found in `{sel_maint_mach}`. Use the uploader above to drop and save Word templates.")
            
        with tab_oq_shift:
            import importlib
            import backend.oq_shift_service as oq_shift_svc
            try:
                importlib.reload(oq_shift_svc)
            except Exception:
                pass
            
            tab_shift_gen, tab_shift_maint = st.tabs([
                "⚡ Protocol Generator",
                "⚙️ Machine Type & Template Maintenance"
            ])
            
            with tab_shift_gen:
                st.markdown("#### 🔄 OQ Shift Register Protocol Generator")
                st.caption("Automatically match Shift Register Excel variables against machine-specific shift templates, assemble test protocols into the Word master template, and generate matched shift register summary reports.")
                
                # Machine & Screen Configuration Row
                cfg_shift_col1, cfg_shift_col2 = st.columns(2)
                with cfg_shift_col1:
                    available_shift_machines = oq_shift_svc.get_available_machine_types()
                    selected_shift_machine = None
                    if available_shift_machines:
                        default_shift_m_idx = 0
                        selected_shift_machine = st.selectbox(
                            "🏭 Select Machine Type",
                            available_shift_machines,
                            index=default_shift_m_idx,
                            key="oq_shift_machine_sel",
                            help="Target machine type folder in GMP_Shift Register"
                        )
                    else:
                        st.selectbox(
                            "🏭 Select Machine Type",
                            ["No machine types found"],
                            index=0,
                            disabled=True,
                            key="oq_shift_machine_sel_empty",
                            help="No machine type directories found in GMP_Shift Register. Use Maintenance tab to add one."
                        )
                        st.caption("💡 *No machine types found in `IQOQDQ/OQ_Shift/GMP_Shift Register`. Switch to '⚙️ Machine Type & Template Maintenance' tab to create one.*")
                with cfg_shift_col2:
                    shift_screen_type = st.radio(
                        "🖥️ Screen / Visu Type",
                        ["🖥️ IPC (Industrial PC)", "📱 Magilis"],
                        horizontal=True,
                        key="oq_shift_screen_type",
                        help="Magilis file formats will be integrated in future releases"
                    )
                
                if "Magilis" in shift_screen_type:
                    st.info("💡 **Magilis Visu Note:** Magilis file structure parser is scheduled for upcoming release. Current processing uses standard IPC shift register variable mapping.")
                
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
                
                col_excel_shift, col_word_shift = st.columns(2)
                
                with col_excel_shift:
                    st.markdown("##### 📊 1. Shift Register Information Excel")
                    
                    uploaded_excel_shift = st.file_uploader(
                        "Upload Custom Shift Register Excel (*.xlsx, *.xls)",
                        type=["xlsx", "xls"],
                        key="oq_shift_upload_excel",
                        help="Upload Excel containing shift register variables/codes in Column A"
                    )
                    
                    shift_dir = os.path.abspath(r"IQOQDQ/OQ_Shift")
                    os.makedirs(shift_dir, exist_ok=True)
                    sample_shift_path = os.path.join(shift_dir, "5XXXX-ShiftRegisterInfo.xlsx")
                    
                    if uploaded_excel_shift is not None:
                        excel_input_shift = uploaded_excel_shift
                        st.caption(f"📂 Active File: `{uploaded_excel_shift.name}`")
                    elif os.path.exists(sample_shift_path):
                        excel_input_shift = sample_shift_path
                        st.caption(f"💡 *Defaulting to sample:* `{os.path.basename(sample_shift_path)}`")
                    else:
                        excel_input_shift = None
                    
                    if os.path.exists(sample_shift_path):
                        with open(sample_shift_path, "rb") as sf:
                            sample_shift_bytes = sf.read()
                        st.download_button(
                            label=f"📥 Download Sample Excel ({os.path.basename(sample_shift_path)})",
                            data=sample_shift_bytes,
                            file_name=os.path.basename(sample_shift_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="oq_shift_dl_sample_excel",
                            use_container_width=True
                        )
                
                with col_word_shift:
                    st.markdown("##### 📄 2. Word Master Template")
                    tmpl_dir_shift = os.path.abspath(r"IQOQDQ/OQ_Shift")
                    os.makedirs(tmpl_dir_shift, exist_ok=True)
                    tmpl_files_shift = [f for f in os.listdir(tmpl_dir_shift) if not f.startswith("~$") and not f.startswith(".") and f.lower().endswith(('.doc', '.docx'))] if os.path.exists(tmpl_dir_shift) else []
                    word_tmpl_path_shift = None
                    if tmpl_files_shift:
                        default_tmpl_idx_shift = 0
                        for idx, fn in enumerate(tmpl_files_shift):
                            if "11_oq_shift" in fn.lower() or "shift" in fn.lower():
                                default_tmpl_idx_shift = idx
                                break
                        selected_tmpl_shift = st.selectbox("Select Word Master Template", tmpl_files_shift, index=default_tmpl_idx_shift, key="oq_shift_tmpl_sel")
                        word_tmpl_path_shift = os.path.abspath(os.path.join(tmpl_dir_shift, selected_tmpl_shift))
                        st.caption(f"📄 Master Template: `{selected_tmpl_shift}` *(from OQ_Shift)*")
                        
                        if os.path.exists(word_tmpl_path_shift):
                            with open(word_tmpl_path_shift, "rb") as tf:
                                tmpl_bytes_shift = tf.read()
                            is_doc_ext_shift = selected_tmpl_shift.lower().endswith('.doc')
                            mime_shift = "application/msword" if is_doc_ext_shift else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            st.download_button(
                                label=f"📥 Download Current Master Template ({selected_tmpl_shift})",
                                data=tmpl_bytes_shift,
                                file_name=selected_tmpl_shift,
                                mime=mime_shift,
                                key="oq_shift_dl_current_tmpl",
                                use_container_width=True
                            )

                    uploaded_tmpl_shift = st.file_uploader(
                        "📤 Upload New Master Template (*.docx, *.doc)",
                        type=["docx", "doc"],
                        key="oq_shift_upload_word",
                        help="Upload a Word master template to save into OQ_Shift directory"
                    )
                    if uploaded_tmpl_shift is not None:
                        temp_tmpl_p = os.path.abspath(os.path.join(tmpl_dir_shift, uploaded_tmpl_shift.name))
                        with open(temp_tmpl_p, "wb") as f:
                            f.write(uploaded_tmpl_shift.getvalue())
                        word_tmpl_path_shift = temp_tmpl_p
                        st.success(f"✅ Saved new template `{uploaded_tmpl_shift.name}` to OQ_Shift directory successfully!")
                        st.rerun()

                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                
                # Action Button
                btn_gen_shift1, btn_gen_shift2 = st.columns([1.2, 2])
                with btn_gen_shift1:
                    process_btn_shift = st.button("⚡ Populate & Generate OQ Shift Register Word Document", type="primary", use_container_width=True, key="oq_shift_process_btn")
                
                if process_btn_shift:
                    if not excel_input_shift:
                        st.error("❌ Please select or upload a Shift Register Excel file.")
                    else:
                        with st.spinner(f"Matching shift register variables against '{selected_shift_machine}' library and assembling Section 3 test protocols..."):
                            try:
                                result_shift = oq_shift_svc.generate_oq_shift_word(
                                    excel_source=excel_input_shift,
                                    machine_type=selected_shift_machine,
                                    word_template_path=word_tmpl_path_shift
                                )
                                st.session_state["oq_shift_last_result"] = result_shift
                                st.success(f"🎉 **Word Document Ready!** Successfully matched {result_shift['total_matched']} shift items and assembled `{result_shift['file_name']}`.")
                            except Exception as ex:
                                st.error(f"❌ Error generating OQ Shift Register document: {str(ex)}")
                
                # Display Results
                if "oq_shift_last_result" in st.session_state and st.session_state["oq_shift_last_result"]:
                    res_shift = st.session_state["oq_shift_last_result"]
                    
                    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                    m1, m2, m3, m4 = st.columns(4)
                    m1.metric("📋 Total Scanned", f"{res_shift['total_scanned']} Variables")
                    m2.metric("✅ Matched Items", f"{res_shift['total_matched']} Items")
                    skipped_count = res_shift['total_scanned'] - res_shift['total_matched']
                    m3.metric("⏭️ Skipped / Not Found", f"{skipped_count} Rows")
                    m4.metric("🏷️ Machine Type", res_shift.get('machine_type', selected_shift_machine))
                    
                    # Dual Download Buttons
                    dl_col1, dl_col2 = st.columns(2)
                    with dl_col1:
                        st.download_button(
                            label=f"📥 Download Word Protocol ({res_shift['file_name']})",
                            data=res_shift["doc_bytes"],
                            file_name=res_shift["file_name"],
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary",
                            key="oq_shift_dl_doc_btn",
                            use_container_width=True
                        )
                    with dl_col2:
                        st.download_button(
                            label=f"📊 Download Matched Shift Register Summary (Excel)",
                            data=res_shift["summary_excel_bytes"],
                            file_name=res_shift["summary_excel_name"],
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="oq_shift_dl_excel_btn",
                            use_container_width=True
                        )
                    
                    # Interactive Previews
                    prev_tab1, prev_tab2 = st.tabs([
                        f"✅ Matched Items ({res_shift['total_matched']})",
                        f"🔍 All Scanned Rows ({res_shift['total_scanned']})"
                    ])
                    with prev_tab1:
                        if not res_shift["matched_df"].empty:
                            st.dataframe(res_shift["matched_df"], use_container_width=True, hide_index=True)
                        else:
                            st.warning("No matched shift register items found.")
                    with prev_tab2:
                        if not res_shift["full_df"].empty:
                            st.dataframe(res_shift["full_df"], use_container_width=True, hide_index=True)
            
            with tab_shift_maint:
                st.markdown("#### ⚙️ Machine Type & Shift Register Template Library Maintenance")
                st.caption("Create new Machine Types, drop / upload Word shift templates (*.docx, *.doc), and delete obsolete files to keep each machine library up-to-date.")
                
                maint_shift_machines = oq_shift_svc.get_available_machine_types()
                
                # Machine Type Management Header Bar
                maint_top_col1, maint_top_col2, maint_top_col3 = st.columns([2, 1.2, 1.2])
                
                with maint_top_col1:
                    if maint_shift_machines:
                        sel_shift_maint_mach = st.selectbox(
                            "🏭 Active Machine Type to Maintain",
                            maint_shift_machines,
                            index=0,
                            key="oq_shift_maint_active_mach"
                        )
                    else:
                        sel_shift_maint_mach = None
                        st.selectbox(
                            "🏭 Active Machine Type to Maintain",
                            ["No machine types created"],
                            index=0,
                            disabled=True,
                            key="oq_shift_maint_active_mach_empty"
                        )
                
                with maint_top_col2:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    with st.popover("➕ Add New Machine", use_container_width=True):
                        st.markdown("##### ➕ Create New Machine Type")
                        new_shift_mach_input = st.text_input("Machine Type Code", placeholder="e.g. FP, SC_SI, TZC, FC, TL", key="oq_shift_new_mach_name")
                        if st.button("Create & Initialize", type="primary", use_container_width=True, key="oq_shift_btn_create_mach"):
                            if not new_shift_mach_input.strip():
                                st.error("❌ Please enter a valid Machine Type name.")
                            else:
                                try:
                                    created_name = oq_shift_svc.create_machine_type(new_shift_mach_input)
                                    st.success(f"🎉 Machine Type `{created_name}` created and initialized successfully!")
                                    st.rerun()
                                except Exception as ex:
                                    st.error(f"❌ Error creating machine type: {str(ex)}")
                
                with maint_top_col3:
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    with st.popover("🗑️ Delete Machine", use_container_width=True):
                        st.markdown(f"##### ⚠️ Delete Machine Type: `{sel_shift_maint_mach}`")
                        st.warning(f"This will delete all template files and cache for `{sel_shift_maint_mach}`.")
                        confirm_del_shift_mach = st.checkbox(f"Yes, delete '{sel_shift_maint_mach}'", key="oq_shift_del_mach_chk")
                        if st.button("Delete Machine Type", type="secondary", disabled=not confirm_del_shift_mach, use_container_width=True, key="oq_shift_btn_del_mach"):
                            try:
                                oq_shift_svc.delete_machine_type(sel_shift_maint_mach)
                                st.success(f"🗑️ Machine Type `{sel_shift_maint_mach}` has been deleted.")
                                st.rerun()
                            except Exception as ex:
                                st.error(f"❌ Error deleting machine type: {str(ex)}")
                
                st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                
                if sel_shift_maint_mach:
                    # Upload / Dropzone Card
                    st.markdown(f"##### 📤 Drop / Upload Word Templates to `{sel_shift_maint_mach}`")
                    maint_shift_upload_files = st.file_uploader(
                        f"Drop or Select Word Files (*.docx, *.doc) for {sel_shift_maint_mach}",
                        type=["docx", "doc"],
                        accept_multiple_files=True,
                        key="oq_shift_maint_upload_dropzone",
                        help="Upload test templates (e.g. SR_*.doc, MX_*.docx)"
                    )
                    
                    if maint_shift_upload_files:
                        up_c1, up_c2 = st.columns([1.5, 2])
                        with up_c1:
                            if st.button(f"💾 Save & Sync {len(maint_shift_upload_files)} File(s) into {sel_shift_maint_mach}", type="primary", use_container_width=True, key="oq_shift_btn_save_uploaded_files"):
                                with st.spinner(f"Saving and synchronizing template files for {sel_shift_maint_mach}..."):
                                    saved = oq_shift_svc.upload_machine_files(sel_shift_maint_mach, maint_shift_upload_files)
                                    st.success(f"🎉 Successfully saved and synchronized {len(saved)} file(s) into `{sel_shift_maint_mach}`!")
                                    st.rerun()
                    
                    st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)
                    
                    # File Table and Metrics
                    df_maint_shift_files = oq_shift_svc.list_machine_files_detailed(sel_shift_maint_mach)
                    
                    f_m1, f_m2, f_m3 = st.columns(3)
                    f_m1.metric("📂 Total Templates", f"{len(df_maint_shift_files)} Files")
                    tot_kb = df_maint_shift_files["Size (KB)"].sum() if not df_maint_shift_files.empty else 0
                    f_m2.metric("💾 Total Library Size", f"{tot_kb:.1f} KB")
                    f_m3.metric("🏷️ Selected Machine", sel_shift_maint_mach)
                    
                    st.markdown(f"##### 📋 Existing Template Files in `{sel_shift_maint_mach}`")
                    if not df_maint_shift_files.empty:
                        st.dataframe(df_maint_shift_files, use_container_width=True, hide_index=True)
                        
                        # File Management Actions
                        f_act_c1, f_act_c2 = st.columns(2)
                        
                        with f_act_c1:
                            st.markdown("###### 🗑️ Delete Template File(s)")
                            shift_files_to_remove = st.multiselect(
                                "Select file(s) to remove from library",
                                df_maint_shift_files["Filename"].tolist(),
                                key="oq_shift_maint_remove_sel",
                                help="Select one or more files to delete"
                            )
                            if shift_files_to_remove:
                                if st.button(f"🗑️ Delete {len(shift_files_to_remove)} Selected File(s)", type="secondary", key="oq_shift_maint_btn_remove_files"):
                                    deleted_list = oq_shift_svc.delete_machine_files(sel_shift_maint_mach, shift_files_to_remove)
                                    st.success(f"🗑️ Successfully deleted {len(deleted_list)} file(s).")
                                    st.rerun()
                        
                        with f_act_c2:
                            st.markdown("###### 📥 Download Template File")
                            shift_file_to_inspect = st.selectbox(
                                "Select file to download & inspect",
                                df_maint_shift_files["Filename"].tolist(),
                                key="oq_shift_maint_inspect_sel"
                            )
                            if shift_file_to_inspect:
                                inspect_bytes = oq_shift_svc.get_machine_file_bytes(sel_shift_maint_mach, shift_file_to_inspect)
                                if inspect_bytes:
                                    is_doc_ext = shift_file_to_inspect.lower().endswith('.doc')
                                    mime_ext = "application/msword" if is_doc_ext else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                    st.download_button(
                                        label=f"📥 Download `{shift_file_to_inspect}`",
                                        data=inspect_bytes,
                                        file_name=shift_file_to_inspect,
                                        mime=mime_ext,
                                        key="oq_shift_maint_btn_dl_single",
                                        use_container_width=True
                                    )
                    else:
                        st.info(f"💡 No template files found in `{sel_shift_maint_mach}`. Use the uploader above to drop and save Word templates.")
        
    elif "Rename" in sub_section or "Tag" in sub_section:
        st.markdown("### 🏷️ Rename Tag & Custom Document Properties Tool")
        st.caption("Batch update Custom Document Properties (Copyright, Version, Machine, Order, Serial no., Designation), Document History dates, and automated file renaming.")
        
        import backend.rename_tag_service as rename_tag_svc
        try:
            rename_tag_svc = importlib.reload(rename_tag_svc)
        except Exception:
            pass

        # Step 1: Input Word Files
        st.markdown("##### 📁 1. Target Word Document File(s) Selection")
        st.caption("Upload Word (.docx, .doc) files directly to update properties, Document History dates, and file names.")

        uploaded_word_files = st.file_uploader(
            "Upload Target Word Files (*.docx, *.doc)",
            type=["docx", "doc"],
            accept_multiple_files=True,
            key="rename_tag_file_uploader",
            help="Insert one or multiple Word document files to update properties and file names"
        )
        selected_files_to_process = []
        if uploaded_word_files:
            for uf in uploaded_word_files:
                selected_files_to_process.append({
                    "filename": uf.name,
                    "bytes": uf.getvalue(),
                    "source_path": None
                })
            st.caption(f"✅ Loaded `{len(uploaded_word_files)}` file(s) for processing.")

        st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)

        # Step 2: 7 Data Input Fields
        st.markdown("##### 📝 2. Enter Document Properties & Date (7 Input Fields)")
        st.caption("Fill in the 6 Custom Document Properties and the 7th Date parameter to update across all files.")

        default_today_str = datetime.datetime.now().strftime("%d-%b-%y")

        col_p1, col_p2, col_p3 = st.columns(3)
        with col_p1:
            val_copyright = st.text_input(
                "1. Copyright",
                value="© Copyright by IWK (Thailand) Limited 2026",
                key="rt_val_copyright",
                help="Custom property 'Copyright'"
            )
            val_order = st.text_input(
                "4. Order (Order No.)",
                value="56021",
                key="rt_val_order",
                help="Custom property 'Order' & replaces XXXXX in filenames"
            )

        with col_p2:
            val_version = st.text_input(
                "2. Version",
                value="01",
                key="rt_val_version",
                help="Custom property 'Version'"
            )
            val_baunummer = st.text_input(
                "5. Serial no.",
                value="XXX",
                key="rt_val_baunummer",
                help="Custom property 'Serial no.' (Baunummer)"
            )

        with col_p3:
            val_machine = st.text_input(
                "3. Machine",
                value="IWK XX",
                key="rt_val_machine",
                help="Custom property 'Machine'"
            )
            val_bezeichnung = st.text_input(
                "6. Designation",
                value="Cartoning machine, Tube filling machine, Filling platform",
                key="rt_val_bezeichnung",
                help="Custom property 'Designation' (Bezeichnung)"
            )

        val_date = st.text_input(
            "7. Date (Document History & Filename)",
            value=default_today_str,
            key="rt_val_date",
            help="Input 7: Date format (e.g. 22-Sep-26). Replaces DD-MMM-YYYY in Document History and date in filename."
        )

        # Compute ISO date & doc date preview
        iso_date_preview, doc_date_preview = rename_tag_svc.parse_user_date(val_date)
        st.caption(f"🗓️ **Parsed Date Preview:** Document History Date: `{doc_date_preview}` | Filename Date (ISO): `{iso_date_preview}`")

        st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)

        # Step 3: Filename Renaming Preview
        st.markdown("##### 🔍 3. Filename & Renaming Preview")
        if selected_files_to_process:
            preview_rows = []
            for fitem in selected_files_to_process:
                new_fn = rename_tag_svc.compute_renamed_filename(fitem["filename"], val_order, iso_date_preview)
                preview_rows.append({
                    "Original File Name": fitem["filename"],
                    "Renamed File Name": new_fn,
                    "Order Tag": val_order,
                    "Date Tag (Document History)": doc_date_preview,
                    "Date Tag (Filename)": iso_date_preview
                })
            st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
        else:
            st.info("💡 Please select or upload at least one Word file to preview renaming.")

        st.markdown("<hr style='margin: 15px 0;'>", unsafe_allow_html=True)

        # Step 4: Action Button Process
        col_proc1, col_proc2 = st.columns([1.5, 2])
        with col_proc1:
            btn_process_rt = st.button("⚡ Process & Update Word Files", type="primary", use_container_width=True, key="rt_process_btn")

        if btn_process_rt:
            if not selected_files_to_process:
                st.error("❌ Please insert or select at least one Word file.")
            elif not val_order:
                st.error("❌ Please enter the Order number.")
            else:
                with st.spinner("Updating Custom Properties, Document History text, and renaming files..."):
                    results_list = rename_tag_svc.process_batch_word_files(
                        file_items=selected_files_to_process,
                        copyright_val=val_copyright,
                        version_val=val_version,
                        machine_val=val_machine,
                        order_val=val_order,
                        baunummer_val=val_baunummer,
                        date_val=val_date,
                        bezeichnung_val=val_bezeichnung
                    )

                    st.session_state["rename_tag_last_results"] = results_list
                    st.success(f"🎉 **Successfully Processed {len(results_list)} File(s)!** Custom Document Properties and File Names updated.")

        # Step 5: Save & Download Results
        if "rename_tag_last_results" in st.session_state and st.session_state["rename_tag_last_results"]:
            res_items = st.session_state["rename_tag_last_results"]
            st.markdown("##### 📥 5. Processed Files & Download")

            can_overwrite = any(item.get("source_path") for item in res_items)
            if can_overwrite:
                if st.button("💾 Overwrite Workspace Files Directly (บันทึกทับไฟล์เดิมในดิสก์)", type="secondary", key="rt_overwrite_disk_btn"):
                    overwritten_count = 0
                    for item in res_items:
                        src_p = item.get("source_path")
                        if src_p and os.path.exists(src_p):
                            target_dir = os.path.dirname(src_p)
                            new_path = os.path.join(target_dir, item["new_filename"])
                            with open(new_path, "wb") as out_f:
                                out_f.write(item["processed_bytes"])
                            if new_path != src_p and os.path.exists(src_p):
                                try:
                                    os.remove(src_p)
                                except Exception:
                                    pass
                            overwritten_count += 1
                    st.toast(f"✅ บันทึกทับไฟล์สำเร็จเรียบร้อย {overwritten_count} ไฟล์!", icon="💾")
                    st.rerun()

            if len(res_items) > 1:
                zip_io = io.BytesIO()
                with zipfile.ZipFile(zip_io, "w", compression=zipfile.ZIP_DEFLATED) as zf:
                    for item in res_items:
                        zf.writestr(item["new_filename"], item["processed_bytes"])
                zip_io.seek(0)

                st.download_button(
                    label=f"📦 Download All Processed Files as ZIP ({len(res_items)} Files)",
                    data=zip_io.getvalue(),
                    file_name=f"{val_order}_Processed_Word_Files.zip",
                    mime="application/zip",
                    type="primary",
                    key="rt_dl_zip_btn",
                    use_container_width=True
                )
                st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)

            for idx, item in enumerate(res_items):
                col_res1, col_res2 = st.columns([3, 2])
                with col_res1:
                    st.markdown(f"📄 **Original:** `{item['original_filename']}` → **New:** `{item['new_filename']}`")
                with col_res2:
                    st.download_button(
                        label=f"📥 Download {item['new_filename']}",
                        data=item["processed_bytes"],
                        file_name=item["new_filename"],
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"rt_dl_single_{idx}",
                        use_container_width=True
                    )
        
    else:
        tab_dq, tab_iq, tab_oq, tab_rt = st.tabs(["📐 DQ", "🔧 IQ", "⚡ OQ", "🏷️ Rename tag"])
        with tab_dq:
            st.markdown("### 📐 DQ — Design Qualification")
            st.info("💡 Ready for DQ steps.")
        with tab_iq:
            st.markdown("### 🔧 IQ — Installation Qualification")
            st.info("💡 Ready for IQ steps.")
        with tab_oq:
            st.markdown("### ⚡ OQ — Operational Qualification")
            st.info("💡 Ready for OQ steps.")
        with tab_rt:
            st.markdown("### 🏷️ Rename Tag")
            st.info("💡 Ready for Rename tag steps.")



def render_placeholder_page(page_name):
    st.title(f"{page_name}")
    st.subheader("Feature Under Development")
    
    st.markdown(f"""
    <div style="background-color: var(--secondary-background-color, rgba(255, 255, 255, 0.05)); padding: 40px; border-radius: 12px; border: 1px solid rgba(128, 128, 128, 0.2); box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1); text-align: center; margin-top: 20px;">
        <div style="font-size: 3rem; margin-bottom: 15px;">🚀</div>
        <h3 style="color: var(--text-color); margin-top: 0;">{page_name} feature is coming soon</h3>
        <p style="color: var(--text-color); opacity: 0.8; font-size: 1.1rem; max-width: 500px; margin: 0 auto 25px auto;">
            This module is part of the 2026 Enterprise Portal roadmap. The system is preparing to integrate this module with the core databases.
        </p>
        <span style="background-color: rgba(0, 120, 212, 0.15); color: #0078d4; padding: 8px 16px; border-radius: 20px; font-weight: 600; font-size: 0.9rem;">
            Status: Coming Soon (Fluent UI 2026)
        </span>
    </div>
    """, unsafe_allow_html=True)

def render_login_page():
    # Inject CSS specific to the login page to style standard Streamlit containers and inputs
    st.markdown("""
    <style>
        /* Target the root app to apply light gray background */
        .stApp {
            background-color: #f8fafc !important;
        }
        
        /* Hide sidebar on login page */
        [data-testid="stSidebar"] {
            display: none !important;
        }
        
        /* Center and style the login card container (pure white card) */
        div[data-testid="stVerticalBlockBorder"] {
            max-width: 440px;
            margin: 80px auto !important;
            padding: 40px 35px !important;
            background-color: #ffffff !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 16px !important;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.05) !important;
        }
        
        /* Welcome header typography */
        .login-header-title {
            font-size: 1.85rem;
            font-weight: 700;
            color: #0f172a !important;
            margin-bottom: 2px;
            text-align: left;
        }
        
        .login-header-subtitle {
            font-size: 0.9rem;
            color: #64748b;
            margin-bottom: 25px;
            text-align: left;
        }
        
        /* Text input formatting (light theme) */
        .stTextInput label {
            color: #0f172a !important;
            font-size: 0.85rem !important;
            font-weight: 600 !important;
            margin-bottom: 6px !important;
        }
        
        .stTextInput input {
            background-color: #ffffff !important;
            color: #0f172a !important;
            border: 1px solid #cbd5e1 !important;
            border-radius: 8px !important;
            padding: 10px 14px !important;
            font-size: 0.95rem !important;
            height: 42px !important;
        }
        
        .stTextInput input:focus {
            border-color: #2563eb !important;
            box-shadow: 0 0 0 1px #2563eb !important;
            background-color: #ffffff !important;
            color: #0f172a !important;
        }
        
        /* Selectbox label and container formatting */
        .stSelectbox label {
            color: #0f172a !important;
            font-size: 0.85rem !important;
            font-weight: 600 !important;
        }
        
        .stSelectbox div[data-baseweb="select"] {
            background-color: #ffffff !important;
            border: 1px solid #cbd5e1 !important;
            border-radius: 8px !important;
            color: #0f172a !important;
        }
        
        /* Checkbox formatting */
        .stCheckbox label span {
            color: #475569 !important;
            font-size: 0.85rem !important;
        }
        
        /* Blue checkbox tick box */
        .stCheckbox [data-testid="stCheckbox"] {
            background-color: #2563eb !important;
        }
        
        /* Pill-shaped blue button */
        div.stFormSubmitButton > button {
            background-color: #0066fe !important;
            color: #ffffff !important;
            border: none !important;
            padding: 12px 24px !important;
            font-weight: 600 !important;
            border-radius: 24px !important; /* Pill shape! */
            width: 100% !important;
            box-shadow: 0 4px 10px rgba(0, 102, 254, 0.15) !important;
            height: 44px !important;
            margin-top: 20px !important;
            transition: all 0.2s !important;
        }
        
        div.stFormSubmitButton > button:hover {
            background-color: #0052cc !important;
            box-shadow: 0 6px 15px rgba(0, 102, 254, 0.25) !important;
        }
        
        /* Forgot password link style */
        .forgot-link a {
            color: #2563eb !important;
            text-decoration: none !important;
            font-size: 0.85rem !important;
            font-weight: 500 !important;
        }
        .forgot-link a:hover {
            text-decoration: underline !important;
        }
        
        /* Footer navigation link switch style override */
        .switch-btn-container button {
            background-color: transparent !important;
            color: #2563eb !important;
            border: 1px solid #2563eb !important;
            font-weight: 600 !important;
            font-size: 0.85rem !important;
            border-radius: 8px !important;
            transition: all 0.2s !important;
        }
        
        .switch-btn-container button:hover {
            background-color: rgba(37, 99, 235, 0.05) !important;
            color: #1d4ed8 !important;
        }
    </style>
    """, unsafe_allow_html=True)
    
    if "login_view" not in st.session_state:
        st.session_state["login_view"] = "signin"
        
    with st.container(border=True):
        if st.session_state["login_view"] == "signin":
            st.markdown('<div class="login-header-title">Welcome Back</div>', unsafe_allow_html=True)
            st.markdown('<div class="login-header-subtitle">Sign in to your account</div>', unsafe_allow_html=True)
            
            with st.form("signin_form", clear_on_submit=False):
                username = st.text_input("Username", placeholder="Enter your username", label_visibility="visible")
                password = st.text_input("Password", type="password", placeholder="••••••••", label_visibility="visible")
                
                # Checkbox inside form
                col_rem, col_forgot = st.columns([1, 1])
                with col_rem:
                    remember = st.checkbox("Remember me", key="login_remember_me")
                with col_forgot:
                    st.markdown('<div class="forgot-link" style="text-align: right; padding-top: 3px;"><a href="#">Forgot password?</a></div>', unsafe_allow_html=True)
                    
                submit_btn = st.form_submit_button("Sign In")
                
                if submit_btn:
                    import backend.auth_service as auth_service
                    success, user, msg = auth_service.verify_login(username, password)
                    if success:
                        st.session_state["logged_in"] = True
                        st.session_state["user_id"] = user["id"]
                        st.session_state["username"] = user["username"]
                        st.session_state["user_role"] = user["role"]
                        st.session_state["user_dept"] = user["department"]
                        st.success("✅ Login successful!")
                        st.rerun()
                    else:
                        st.error(f"❌ {msg}")
                        
            # Link to Sign Up
            st.markdown("<div style='margin-top: 20px;'></div>", unsafe_allow_html=True)
            col_left, col_right = st.columns([5, 3])
            with col_left:
                st.markdown('<div style="padding-top: 8px;"><span style="color: #64748b; font-size: 0.85rem;">Don\'t have an account?</span></div>', unsafe_allow_html=True)
            with col_right:
                st.markdown('<div class="switch-btn-container">', unsafe_allow_html=True)
                if st.button("Sign Up", key="link_to_signup", use_container_width=True):
                    st.session_state["login_view"] = "signup"
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
                    
        else: # signup view
            st.markdown('<div class="login-header-title">Create Account</div>', unsafe_allow_html=True)
            st.markdown('<div class="login-header-subtitle">Register a new portal account</div>', unsafe_allow_html=True)
            
            with st.form("signup_form", clear_on_submit=False):
                reg_username = st.text_input("Username", placeholder="Choose a username")
                reg_password = st.text_input("Password", type="password", placeholder="Create secure password")
                reg_confirm = st.text_input("Confirm Password", type="password", placeholder="Repeat password")
                reg_dept = st.text_input("Department / Section", placeholder="e.g. Doc Team, Quality Control")
                reg_role = st.selectbox("Access Level Requested", options=["Operator Department (Level 3)", "Admin Department (Level 2)"])
                
                submit_reg_btn = st.form_submit_button("Register Account")
                
                if submit_reg_btn:
                    if reg_password != reg_confirm:
                        st.error("❌ Passwords do not match.")
                    elif not reg_username.strip() or not reg_password or not reg_dept.strip():
                        st.error("❌ Please fill in all fields.")
                    else:
                        role_code = "admin_department" if "Admin" in reg_role else "operator_department"
                        import backend.auth_service as auth_service
                        success, msg = auth_service.register_user(reg_username, reg_password, role_code, reg_dept)
                        if success:
                            st.success(f"✅ {msg}")
                        else:
                            st.error(f"❌ {msg}")
                            
            st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
            col_left, col_right = st.columns([5, 3])
            with col_left:
                st.markdown('<div style="padding-top: 8px;"><span style="color: #64748b; font-size: 0.85rem;">Already have an account?</span></div>', unsafe_allow_html=True)
            with col_right:
                st.markdown('<div class="switch-btn-container">', unsafe_allow_html=True)
                if st.button("Sign In", key="link_to_signin", use_container_width=True):
                    st.session_state["login_view"] = "signin"
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)

def render_user_management_page():
    st.title("👥 User Authorization Console")
    st.subheader("Manage user registrations and roles (Admin Master Only)")
    
    import backend.auth_service as auth_service
    
    users = auth_service.get_all_users()
    
    # Split into Pending and All Users
    pending_users = [u for u in users if u["is_approved"] == 0]
    active_users = [u for u in users if u["is_approved"] == 1]
    suspended_users = [u for u in users if u["is_approved"] == -1]
    
    st.markdown("### ⏳ Pending Approval Request")
    if not pending_users:
        st.success("No pending approval requests.")
    else:
        for u in pending_users:
            with st.container(border=True):
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.write(f"**Username:** `{u['username']}` | **Role:** `{u['role']}` | **Department:** `{u['department']}`")
                    st.caption(f"Registered at: {u['created_at']}")
                with col2:
                    if st.button("✅ Approve", key=f"app_{u['id']}", use_container_width=True):
                        success, msg = auth_service.update_user_status(u["id"], 1, st.session_state["username"])
                        if success:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                with col3:
                    if st.button("❌ Reject / Suspend", key=f"rej_{u['id']}", use_container_width=True):
                        success, msg = auth_service.update_user_status(u["id"], -1, st.session_state["username"])
                        if success:
                            st.success(msg)
                            st.rerun()
                        else:
                            st.error(msg)
                            
    st.markdown("---")
    st.markdown("### 📋 Users Directory")
    tab_active, tab_suspended = st.tabs([f"🟢 Active Users ({len(active_users)})", f"🔴 Suspended Users ({len(suspended_users)})"])
    
    with tab_active:
        if not active_users:
            st.write("No active users.")
        else:
            for u in active_users:
                with st.container(border=True):
                    col1, col2, col3 = st.columns([3, 1, 1])
                    with col1:
                        st.write(f"**Username:** `{u['username']}` | **Role:** `{u['role']}` | **Department:** `{u['department']}`")
                        st.caption(f"Approved by: {u['approved_by']} at {u['approved_at']}")
                    with col2:
                        # Suspend option
                        if u["role"] != "admin_master":
                            if st.button("🛑 Suspend", key=f"sus_{u['id']}", use_container_width=True):
                                success, msg = auth_service.update_user_status(u["id"], -1, st.session_state["username"])
                                if success:
                                    st.success(msg)
                                    st.rerun()
                                else:
                                    st.error(msg)
                        else:
                            st.write("🛡️ Master Account")
                    with col3:
                        # Delete option
                        if u["role"] != "admin_master":
                            if st.button("🗑️ Delete", key=f"del_{u['id']}", use_container_width=True):
                                success, msg = auth_service.delete_user(u["id"])
                                if success:
                                    st.success(msg)
                                    st.rerun()
                                else:
                                    st.error(msg)
                                    
    with tab_suspended:
        if not suspended_users:
            st.write("No suspended users.")
        else:
            for u in suspended_users:
                with st.container(border=True):
                    col1, col2, col3 = st.columns([3, 1, 1])
                    with col1:
                        st.write(f"**Username:** `{u['username']}` | **Role:** `{u['role']}` | **Department:** `{u['department']}`")
                        st.caption(f"Suspended by: {u['approved_by']} at {u['approved_at']}")
                    with col2:
                        if st.button("🟢 Activate", key=f"act_{u['id']}", use_container_width=True):
                            success, msg = auth_service.update_user_status(u["id"], 1, st.session_state["username"])
                            if success:
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)
                    with col3:
                        if st.button("🗑️ Delete", key=f"del_sus_{u['id']}", use_container_width=True):
                            success, msg = auth_service.delete_user(u["id"])
                            if success:
                                st.success(msg)
                                st.rerun()
                            else:
                                st.error(msg)

def main():
    st.set_page_config(page_title="PDF OCR Splitter & Tools", layout="wide")
    
    # Initialize authentication DB in background (preserved for future development)
    try:
        import backend.auth_service as auth_service
        auth_service.initialize_auth_db()
    except Exception:
        pass
    
    # Authentication temporarily bypassed: auto-assign Administrator profile
    if not st.session_state.get("logged_in"):
        st.session_state["logged_in"] = True
        st.session_state["user_id"] = 1
        st.session_state["username"] = "Administrator"
        st.session_state["user_role"] = "admin_master"
        st.session_state["user_dept"] = "IT"
        
    user_role = st.session_state.get("user_role")
    username = st.session_state.get("username")
    user_dept = st.session_state.get("user_dept")
    
    # Inject Microsoft Fabric Theme CSS (Dark navy blue gradient, Fluent UI 2026)
    st.markdown("""
    <style>
        /* Import Google Font Inter */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
        
        /* Apply Inter globally */
        html, body, [class*="css"], .stApp {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
        }
        
        /* Sidebar Microsoft Fabric Style (Dark Navy Blue Gradient for both Light & Dark Mode) */
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #091930 0%, #112d55 100%) !important;
            color: #ffffff !important;
            border-right: 1px solid #1e3a5f !important;
        }
        
        /* Force iframe and option menu containers in sidebar to dark navy */
        section[data-testid="stSidebar"] iframe {
            background-color: #091930 !important;
            border-radius: 8px !important;
        }
        
        /* Sidebar Typography */
        section[data-testid="stSidebar"] h1, 
        section[data-testid="stSidebar"] h2, 
        section[data-testid="stSidebar"] h3, 
        section[data-testid="stSidebar"] h4, 
        section[data-testid="stSidebar"] h5, 
        section[data-testid="stSidebar"] h6, 
        section[data-testid="stSidebar"] span, 
        section[data-testid="stSidebar"] p, 
        section[data-testid="stSidebar"] label {
            color: #ffffff !important;
        }
        
        /* Style headers in Main content dynamically for Light & Dark mode */
        h1, h2, h3, h4, h5, h6 {
            color: var(--text-color) !important;
            font-weight: 700 !important;
        }
        
        /* Adaptive text contrast for labels & markdown */
        .stRadio label, .stFileUploader label, div[data-testid="stMarkdownContainer"] p {
            color: var(--text-color) !important;
        }
        
        /* Style Cards/Containers in Main Workspace adaptively */
        div[data-testid="stExpander"], div[data-testid="stVerticalBlockBorder"] {
            background-color: var(--secondary-background-color, rgba(255, 255, 255, 0.05)) !important;
            border: 1px solid rgba(128, 128, 128, 0.2) !important;
            border-radius: 12px !important;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1) !important;
            margin-bottom: 12px !important;
            transition: all 0.2s ease !important;
        }
        div[data-testid="stExpander"]:hover, div[data-testid="stVerticalBlockBorder"]:hover {
            border-color: #0078d4 !important;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1) !important;
        }
        div[data-testid="stVerticalBlockBorder"] {
            padding: 24px !important;
        }
        
        /* Style buttons to look professional and crisp */
        button {
            border-radius: 8px !important;
            transition: all 0.2s ease !important;
        }
        
        /* Primary Button (e.g. Start Split) */
        button[kind="primary"] {
            background-color: #0078d4 !important;
            color: white !important;
            border: none !important;
            padding: 10px 24px !important;
            font-weight: 600 !important;
            box-shadow: 0 4px 6px rgba(0, 120, 212, 0.15) !important;
        }
        button[kind="primary"]:hover {
            background-color: #006cbe !important;
            transform: translateY(-1px) !important;
            box-shadow: 0 6px 12px rgba(0, 120, 212, 0.2) !important;
        }
        
        /* Styling for Logout Button in sidebar to ensure high contrast and readability */
        div[data-testid="stSidebar"] div[data-testid="stButton"] button {
            background-color: #000000 !important;
            color: #ffffff !important;
            border: 1px solid #334155 !important;
            font-weight: 600 !important;
            font-size: 0.95rem !important;
            padding: 8px 16px !important;
            transition: all 0.2s ease-in-out !important;
        }
        div[data-testid="stSidebar"] div[data-testid="stButton"] button:hover {
            background-color: #dc2626 !important;
            color: #ffffff !important;
            box-shadow: 0 4px 10px rgba(220, 38, 38, 0.25) !important;
            transform: translateY(-1px) !important;
        }
    </style>
    """, unsafe_allow_html=True)
    
    st.sidebar.title("🏢 IWK Document")
    
    # Render user profile card in the sidebar
    st.sidebar.markdown(f"""
    <div style="background-color: rgba(255, 255, 255, 0.05); padding: 12px; border-radius: 8px; border-left: 4px solid #0078d4; margin-bottom: 15px;">
        <span style="color: #94a3b8; font-size: 0.8rem; text-transform: uppercase;">Logged in as:</span><br/>
        <strong style="color: #ffffff; font-size: 1.05rem;">{username}</strong><br/>
        <span style="color: #60a5fa; font-size: 0.85rem; font-weight: 500;">💼 {user_role.replace('_', ' ').title()}</span><br/>
        <span style="color: #94a3b8; font-size: 0.8rem;">🏢 Dept: {user_dept}</span>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar menu options (User Authorization bypassed, all tools unlocked)
    sidebar_options = ["Dashboard", "Document Tool Center", "Data Extraction", "Reports", "Workflow", "Administration", "Recycle Bin"]
    sidebar_icons = ["house", "folder2-open", "search", "graph-up", "gear", "shield-lock", "trash"]

    with st.sidebar:
        page = option_menu(
            menu_title="Select Section",
            options=sidebar_options,
            icons=sidebar_icons,
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "#091930!important"},
                "menu-title": {"font-weight": "700", "color": "#ffffff!important"},
                "nav-link": {"font-size": "14px", "text-align": "left", "margin": "3px 0px", "color": "#e2e8f0!important", "background-color": "#091930!important", "font-weight": "500"},
                "nav-link-selected": {"background-color": "#0078d4!important", "color": "#ffffff!important", "font-weight": "600", "border-left": "4px solid #60a5fa"}
            }
        )
        
        selected_tool = None
        if page == "Document Tool Center":
            dtc_tools = ["OCR & AI (from QC)", "Advanced OCR Adjustment", "IWK Certificate", "ETK Verification", "Calibration Certificate", "Fault Assistance", "Machine Configuration System", "IQOQDQ"]
            dtc_icons = ["robot", "stars", "award", "check2-all", "patch-check", "wrench", "gear", "clipboard-check"]
                
            st.markdown("<hr style='margin: 10px 0; border-color: #1e3a5f;'>", unsafe_allow_html=True)
            selected_tool = option_menu(
                menu_title="Available Tools",
                options=dtc_tools,
                icons=dtc_icons,
                menu_icon="tools",
                default_index=0,
                styles={
                    "container": {"padding": "0!important", "background-color": "#091930!important", "border": "none"},
                    "menu-title": {"font-weight": "700", "color": "#ffffff!important"},
                    "nav-link": {"font-size": "13px", "text-align": "left", "margin":"2px 0px", "padding-left": "25px", "color": "#e2e8f0!important", "background-color": "#091930!important", "font-weight": "500"},
                    "nav-link-selected": {"background-color": "#0078d4!important", "color": "#ffffff!important", "font-weight": "600", "border-left": "4px solid #60a5fa"}
                }
            )
            
            selected_calib_sub = None
            if selected_tool == "Calibration Certificate":
                st.markdown("""
                <div style='margin-left: 15px; margin-top: 6px; margin-bottom: 2px; font-size: 0.75rem; font-weight: 700; color: #60a5fa; letter-spacing: 0.5px;'>
                    ↳ Calibration Modules
                </div>
                """, unsafe_allow_html=True)
                selected_calib_sub = option_menu(
                    menu_title=None,
                    options=["📜 Calibration Processing", "💼 Quotation Verification", "🔢 Certificate No. Mapping"],
                    icons=["file-earmark-spreadsheet", "check2-square", "card-list"],
                    menu_icon="cast",
                    default_index=0,
                    key="calib_semi_submenu_nav",
                    styles={
                        "container": {"padding": "0 0 0 12px!important", "background-color": "#061224!important", "border": "none"},
                        "nav-link": {"font-size": "12px", "text-align": "left", "margin":"2px 0px", "padding-left": "20px", "color": "#cbd5e1!important", "background-color": "#0a192f!important", "border-radius": "6px"},
                        "nav-link-selected": {"background-color": "#0284c7!important", "color": "#ffffff!important", "font-weight": "700", "border-left": "3px solid #38bdf8"}
                    }
                )

            selected_iqoqdq_sub = None
            if selected_tool == "IQOQDQ":
                st.markdown("""
                <div style='margin-left: 15px; margin-top: 6px; margin-bottom: 2px; font-size: 0.75rem; font-weight: 700; color: #60a5fa; letter-spacing: 0.5px;'>
                    ↳ Qualification Stages
                </div>
                """, unsafe_allow_html=True)
                selected_iqoqdq_sub = option_menu(
                    menu_title=None,
                    options=["📐 DQ", "🔧 IQ", "⚡ OQ", "🏷️ Rename tag"],
                    icons=None,
                    menu_icon="cast",
                    default_index=0,
                    key="iqoqdq_semi_submenu_nav",
                    styles={
                        "container": {"padding": "0 0 0 12px!important", "background-color": "#061224!important", "border": "none"},
                        "nav-link": {"font-size": "12px", "text-align": "left", "margin":"2px 0px", "padding-left": "20px", "color": "#cbd5e1!important", "background-color": "#0a192f!important", "border-radius": "6px"},
                        "nav-link-selected": {"background-color": "#0284c7!important", "color": "#ffffff!important", "font-weight": "700", "border-left": "3px solid #38bdf8"}
                    }
                )
            
        # Session Reset Button
        st.markdown("<hr style='margin: 10px 0; border-color: #1e3a5f;'>", unsafe_allow_html=True)
        if st.button("🔄 Reset / Clear Session", use_container_width=True, key="sidebar_logout_btn"):
            st.session_state.clear()
            st.rerun()
    
    # Sidebar Collapse Button at the Bottom
    st.sidebar.markdown("""
    <div style="margin-top: 50px; padding-top: 15px; border-top: 1px solid #1e3a5f; text-align: center;">
        <button id="collapse-btn" style="background-color: transparent; border: 1px solid #334155; color: #94a3b8; border-radius: 8px; padding: 10px 16px; font-size: 0.85rem; cursor: pointer; width: 100%; transition: all 0.2s;" onclick="
            const btn = window.parent.document.querySelector('button[data-testid=\\'collapsedControl\']');
            if (btn) btn.click();
        ">
            🗂️ Collapse Navigation
        </button>
    </div>
    """, unsafe_allow_html=True)
    
    # Page Routing & Protection (All pages unlocked)
    if page == "Dashboard":
        render_home_page()
    elif page == "Document Tool Center":
        if selected_tool == "OCR & AI (from QC)":
            render_ocr_certificate_page()
        elif selected_tool == "Advanced OCR Adjustment":
            render_advanced_ocr_adjustment_page()
        elif selected_tool == "IWK Certificate":
            render_iwk_certificate_page()
        elif selected_tool == "ETK Verification":
            render_etk_verification_page()
        elif selected_tool == "Calibration Certificate":
            sub_mod = selected_calib_sub or "📜 Calibration Processing"
            if "Calibration Processing" in sub_mod:
                render_calibration_certificate_page()
            elif "Quotation Verification" in sub_mod:
                render_quotation_verification_page()
            elif "Certificate No. Mapping" in sub_mod:
                render_certificate_mapping_page()
        elif selected_tool == "Fault Assistance":
            render_fault_assistance_suite_page()
        elif selected_tool in ["Machine Configuration System", "Operating Manual"]:
            render_operating_manual_page()
        elif selected_tool == "IQOQDQ":
            render_iqoqdq_page(sub_section=selected_iqoqdq_sub or "📐 DQ")
        elif selected_tool:
            st.title(f"🛠️ {selected_tool}")
            st.info(f"You have selected the **{selected_tool}** from the Document Tool Center. Development for this tool is in progress.")
        else:
            render_placeholder_page(page)
    elif page in ["Data Extraction", "Reports", "Workflow", "Administration", "Recycle Bin"]:
        render_placeholder_page(page)
    else:
        render_placeholder_page(page)

if __name__ == "__main__":
    main()
