import sys
import os
# Append parent directory of frontend/app.py to python search path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import io
import re
import difflib
import pandas as pd
import streamlit as st
from PIL import Image

# Import backend/frontend modular services
from backend import ocr_service
from backend import mapping_service
from backend import excel_service
from backend import audit_service
from frontend import visualizer
import utils

# ---------------------------------------------------------------------------
# Set Streamlit Page Configuration
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Production OCR & Excel Mapping System",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom premium styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700;800&family=Sarabun:wght@300;400;600;700&family=JetBrains+Mono:wght@500&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', 'Sarabun', sans-serif;
    }
    
    .premium-title {
        background: linear-gradient(90deg, #10b981 0%, #3b82f6 50%, #8b5cf6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 2.3rem;
        font-weight: 800;
        margin-bottom: 2px;
        letter-spacing: -0.03em;
        text-transform: uppercase;
    }
    
    .premium-subtitle {
        color: #64748b;
        font-size: 0.95rem;
        font-weight: 500;
        margin-bottom: 25px;
        letter-spacing: 0.05em;
        text-transform: uppercase;
    }
    
    .glass-card {
        background: rgba(255, 255, 255, 0.7);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid rgba(226, 232, 240, 0.8);
        border-radius: 16px;
        padding: 24px;
        margin-bottom: 20px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.04), 0 4px 6px -2px rgba(0, 0, 0, 0.01);
    }
    
    .status-badge {
        padding: 4px 8px;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 600;
        display: inline-block;
    }
    .badge-success { background-color: #dcfce7; color: #15803d; }
    .badge-warning { background-color: #fef9c3; color: #a16207; }
    .badge-error { background-color: #fee2e2; color: #b91c1c; }
    .badge-neutral { background-color: #f1f5f9; color: #475569; }
    
    .column-title {
        font-size: 1.25rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 12px;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 8px;
    }
</style>
""", unsafe_allow_html=True)

# Helper for cropping region
def crop_match_region(img, anchor_rect, value_rect, padding=12):
    """
    Crops image around label anchor and value.
    """
    if not img:
        return None
        
    ax, ay, aw, ah = anchor_rect.get('x', 0), anchor_rect.get('y', 0), anchor_rect.get('width', 0), anchor_rect.get('height', 0)
    
    if value_rect:
        vx, vy, vw, vh = value_rect.get('x', 0), value_rect.get('y', 0), value_rect.get('width', 0), value_rect.get('height', 0)
        x1 = min(ax, vx)
        y1 = min(ay, vy)
        x2 = max(ax + aw, vx + vw)
        y2 = max(ay + ah, vy + vh)
    else:
        x1, y1, x2, y2 = ax, ay, ax + aw, ay + ah
        
    w, h = img.size
    x1 = max(0, int(x1 - padding))
    y1 = max(0, int(y1 - padding))
    x2 = min(w, int(x2 + padding))
    y2 = min(h, int(y2 + padding))
    
    return img.crop((x1, y1, x2, y2))

def main():
    st.markdown('<div class="premium-title">🤖 Production OCR & Excel Mapping System</div>', unsafe_allow_html=True)
    st.markdown('<div class="premium-subtitle">Azure Document Intelligence, SQLite Review Database & Interactive Verification</div>', unsafe_allow_html=True)
    
    # Initialize session state
    if "ocr_completed" not in st.session_state:
        st.session_state.ocr_completed = False
        st.session_state.excel_bytes = None
        st.session_state.excel_filename = ""
        st.session_state.uploaded_pdfs = {}
        st.session_state.pdf_images = {}
        st.session_state.ocr_results = {}
        st.session_state.mappings = []
        st.session_state.active_pdf = None
        st.session_state.active_adjust_row = None
        st.session_state.loaded_doc_id = None
        
    # Initialize SQLite database schema
    audit_service.init_db()
    
    # -----------------------------------------------------------------------
    # Sidebar Settings
    # -----------------------------------------------------------------------
    st.sidebar.markdown("### 🔑 AZURE CREDENTIALS")
    azure_endpoint = st.sidebar.text_input(
        "Azure Endpoint",
        value=st.session_state.get("azure_endpoint", ""),
        placeholder="https://<resource>.cognitiveservices.azure.com/",
        type="password"
    )
    azure_key = st.sidebar.text_input(
        "Azure API Key",
        value=st.session_state.get("azure_key", ""),
        placeholder="Enter Azure Document Intelligence key...",
        type="password"
    )
    st.session_state["azure_endpoint"] = azure_endpoint
    st.session_state["azure_key"] = azure_key
    
    st.sidebar.markdown("### ⚙️ SYSTEM SETTINGS")
    ocr_engine_choice = st.sidebar.selectbox(
        "OCR Engine",
        options=["Azure Document Intelligence (prebuilt-layout)", "Windows Media OCR (Instant)", "Tesseract OCR (Custom Path)", "TrOCR (Handwriting)"]
    )
    
    ocr_lang_choice = st.sidebar.selectbox(
        "OCR Language",
        options=["German + English (deu+eng)", "English (eng)"],
        index=0
    )
    lang_code = "deu+eng" if "German" in ocr_lang_choice else "eng"
    
    tesseract_path = ""
    if "Tesseract" in ocr_engine_choice:
        tesseract_path = st.sidebar.text_input(
            "Tesseract Binary Path",
            value=r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        )
        
    poppler_path = st.sidebar.text_input(
        "Poppler Bin Path (Optional)",
        value=""
    )
    
    st.sidebar.markdown("### 🔍 OCR PREPROCESSING")
    do_grayscale = st.sidebar.checkbox("Grayscale", value=True)
    do_denoise = st.sidebar.checkbox("Noise Removal", value=True)
    do_contrast = st.sidebar.checkbox("Contrast Enhancement", value=True)
    
    st.sidebar.markdown("### 📏 NEIGHBORHOOD MATCHING")
    search_limit = st.sidebar.slider("Search Limit (px)", 100, 1200, 800, 50)
    search_direction = st.sidebar.selectbox("Search Direction", ["Right", "Below"])
    
    # SQLite Document History / Archive
    st.sidebar.markdown("### 📁 REVIEW DATABASE ARCHIVE")
    db_docs = audit_service.get_all_documents()
    if db_docs:
        doc_names = [f"#{d['id']} - {d['filename']} ({d['status']})" for d in db_docs]
        selected_doc_str = st.sidebar.selectbox(
            "Load verified document:",
            options=["-- Upload New PDF --"] + doc_names,
            key="archive_selector"
        )
        
        if selected_doc_str != "-- Upload New PDF --":
            doc_id = int(selected_doc_str.split(" - ")[0].replace("#", ""))
            doc_info = next(d for d in db_docs if d['id'] == doc_id)
            
            if st.session_state.loaded_doc_id != doc_id:
                db_fields = audit_service.get_document_fields(doc_id)
                st.session_state.mappings = db_fields
                st.session_state.active_pdf = doc_info['filename']
                st.session_state.ocr_completed = True
                st.session_state.loaded_doc_id = doc_id
                st.session_state.active_adjust_row = None
                st.rerun()
                
    # -----------------------------------------------------------------------
    # Main Tabs Dashboard Layout
    # -----------------------------------------------------------------------
    tab_review, tab_mapping, tab_audit = st.tabs([
        "🔍 Verification & Review (การตรวจสอบข้อมูล)",
        "🗺️ Mapping Table Rules (การตั้งค่าพิกัด Excel Cell)",
        "📜 Export & Audit Log (ประวัติการส่งออกและระบบเรียนรู้)"
    ])
    
    # --- Tab 2: Mapping Configuration ---
    with tab_mapping:
        st.subheader("🗺️ Excel Cell Target Mapping Configuration Table")
        st.write("Configure custom target cell coordinates in the Excel templates. Unmapped labels default to the writeback column offsets.")
        
        # CRUD Form
        col_form, col_view = st.columns([1, 2])
        
        with col_form:
            st.markdown("##### Add or Edit Mapping Rule")
            all_maps = audit_service.get_all_mappings()
            map_dict = {f"{m['FieldName']} ({m['SheetName']}!{m['CellAddress']})": m for m in all_maps}
            
            mode = st.radio("Mode", ["Add New Mapping Rule", "Edit/Delete Existing Rule"], horizontal=True)
            
            if mode == "Add New Mapping Rule":
                new_f = st.text_input("FieldName (e.g. Auftrags-Nr. / Order-no.)")
                new_s = st.text_input("SheetName", value="Master")
                new_c = st.text_input("CellAddress", placeholder="E.g. F12, B3")
                
                if st.button("➕ Save New Rule", use_container_width=True):
                    if new_f and new_s and new_c:
                        success = audit_service.add_mapping(new_f, new_s, new_c)
                        if success:
                            st.success(f"Mapping for '{new_f}' created successfully!")
                            st.rerun()
                        else:
                            st.error("Error: FieldName must be unique.")
                    else:
                        st.error("Please fill in all fields.")
            else:
                if map_dict:
                    selected_key = st.selectbox("Select mapping rule:", options=list(map_dict.keys()))
                    selected_map = map_dict[selected_key]
                    
                    edit_s = st.text_input("Edit SheetName", value=selected_map['SheetName'])
                    edit_c = st.text_input("Edit CellAddress", value=selected_map['CellAddress'])
                    
                    col_edit_btn, col_del_btn = st.columns(2)
                    with col_edit_btn:
                        if st.button("💾 Update Rule", use_container_width=True):
                            audit_service.update_mapping(selected_map['id'], edit_s, edit_c)
                            st.success("Mapping updated successfully!")
                            st.rerun()
                    with col_del_btn:
                        if st.button("🗑️ Delete Rule", use_container_width=True, type="primary"):
                            audit_service.delete_mapping(selected_map['id'])
                            st.warning("Mapping rule deleted!")
                            st.rerun()
                else:
                    st.info("No mapping rules available to edit.")
                    
        with col_view:
            st.markdown("##### Active Mapping Table Settings")
            maps = audit_service.get_all_mappings()
            if maps:
                st.dataframe(pd.DataFrame(maps), use_container_width=True, hide_index=True)
            else:
                st.info("MappingTable is currently empty.")
                
    # --- Tab 3: Export & Audit Log ---
    with tab_audit:
        st.subheader("📜 System Audit Trail & Export History")
        
        tab_history, tab_feedback = st.tabs([
            "📂 File Export History Logs",
            "✨ Human Correction Feedback Datasets (TrainingData)"
        ])
        
        with tab_history:
            history = audit_service.get_export_history()
            if history:
                st.dataframe(pd.DataFrame(history), use_container_width=True, hide_index=True)
            else:
                st.info("No export history logged yet.")
                
        with tab_feedback:
            feedback = audit_service.get_all_training_data()
            if feedback:
                st.write("Corrected handwriting dataset collected from human verification modifications:")
                st.dataframe(pd.DataFrame(feedback), use_container_width=True, hide_index=True)
            else:
                st.info("No corrected handwriting feedback logged yet.")
                
    # --- Tab 1: Verification & Review ---
    with tab_review:
        # File Dropzones UI
        st.markdown("### 📂 Load Template & Scan Files")
        
        col_xl, col_pdf = st.columns(2)
        with col_xl:
            excel_file = st.file_uploader(
                "1. Upload Excel Template (.xlsx)",
                type=["xlsx"],
                key="excel_uploader"
            )
        with col_pdf:
            pdf_files = st.file_uploader(
                "2. Upload Scanned PDF(s)",
                type=["pdf"],
                accept_multiple_files=True,
                key="pdf_uploader"
            )
            
        if excel_file:
            st.session_state.excel_bytes = excel_file.getvalue()
            st.session_state.excel_filename = excel_file.name
            
        if pdf_files:
            current_names = [f.name for f in pdf_files]
            existing_names = list(st.session_state.uploaded_pdfs.keys())
            if sorted(current_names) != sorted(existing_names):
                st.session_state.uploaded_pdfs = {f.name: f.getvalue() for f in pdf_files}
                st.session_state.ocr_completed = False
                st.session_state.mappings = []
                st.session_state.loaded_doc_id = None
            
        # -------------------------------------------------------------------
        # OCR Mapping Loop Execution
        # -------------------------------------------------------------------
        if st.session_state.excel_bytes and (st.session_state.uploaded_pdfs or st.session_state.loaded_doc_id):
            # If OCR completed, show review dashboard
            if st.session_state.ocr_completed:
                active_pdf_name = st.session_state.active_pdf
                
                # Check if PDF bytes exist (archived doc loaded)
                pdf_bytes_missing = active_pdf_name not in st.session_state.uploaded_pdfs or len(st.session_state.uploaded_pdfs[active_pdf_name]) == 0
                if pdf_bytes_missing:
                    st.warning(f"⚠️ PDF File '{active_pdf_name}' is not uploaded. Bounding highlights and crops are disabled. You can still modify text values below. Upload the matching PDF to restore crops.")
                    
                num_pages = 1
                pages = []
                if not pdf_bytes_missing:
                    if active_pdf_name in st.session_state.pdf_images:
                        pages = st.session_state.pdf_images[active_pdf_name]
                        num_pages = len(pages)
                        
                active_page = st.number_input(
                    f"📄 Page (1 - {num_pages}):",
                    min_value=1,
                    max_value=num_pages,
                    value=1,
                    step=1,
                    key="review_page_num_input"
                )
                page_idx = active_page - 1
                
                # Fetch mappings for this page
                page_mappings = []
                page_mappings_indices = []
                for idx, m in enumerate(st.session_state.mappings):
                    if m['page'] == page_idx:
                        page_mappings.append(m)
                        page_mappings_indices.append(idx)
                        
                panel_left, panel_right = st.columns([1, 1])
                
                # --- LEFT PANEL: PDF Page Preview ---
                with panel_left:
                    st.markdown('<div class="column-title">👁️ Scanned Document Preview</div>', unsafe_allow_html=True)
                    
                    zoom_factor = st.slider("Zoom Factor (%)", 50, 150, 100, 10) / 100.0
                    
                    # Active Row Selector for quick crop
                    all_rows_options = ["None (View Mode)"]
                    for pm in page_mappings:
                        de_lbl = pm['de_label'] or "(Empty)"
                        all_rows_options.append(f"Row {pm['row_num']}: {de_lbl[:30]}")
                        
                    current_sel = "None (View Mode)"
                    if st.session_state.active_adjust_row is not None:
                        act_m = st.session_state.mappings[st.session_state.active_adjust_row]
                        de_lbl = act_m['de_label'] or "(Empty)"
                        current_sel = f"Row {act_m['row_num']}: {de_lbl[:30]}"
                        
                    selected_row_str = st.selectbox(
                        "🎯 Active Row to Adjust with Mouse (เลือกแถวเพื่อใช้เมาส์ลากปรับกรอบ):",
                        options=all_rows_options,
                        index=all_rows_options.index(current_sel) if current_sel in all_rows_options else 0,
                        key="left_panel_row_selector"
                    )
                    
                    if selected_row_str == "None (View Mode)":
                        if st.session_state.active_adjust_row is not None:
                            st.session_state.active_adjust_row = None
                            st.rerun()
                    else:
                        r_num = int(selected_row_str.split(":")[0].replace("Row ", ""))
                        new_active_idx = None
                        for idx, pm in enumerate(st.session_state.mappings):
                            if pm['page'] == page_idx and pm['row_num'] == r_num:
                                new_active_idx = idx
                                break
                        if st.session_state.active_adjust_row != new_active_idx:
                            st.session_state.active_adjust_row = new_active_idx
                            st.rerun()
                            
                    if not pdf_bytes_missing:
                        base_img = pages[page_idx]
                        
                        if st.session_state.active_adjust_row is not None:
                            active_idx = st.session_state.active_adjust_row
                            active_m = st.session_state.mappings[active_idx]
                            
                            st.markdown(f"""
                            <div style="background-color: #fee2e2; border-left: 4px solid #ef4444; padding: 10px; border-radius: 8px; margin-bottom: 10px;">
                                <strong>🎯 Active Mouse Adjuster: Row {active_m['row_num']}</strong><br/>
                                <span style="font-size: 0.85rem;">Drag/resize the red box on the canvas below to select the handwritten target area.</span>
                            </div>
                            """, unsafe_allow_html=True)
                            
                            from streamlit_cropper import st_cropper
                            cropper_img = base_img.convert("RGB")
                            
                            current_rect = active_m.get('value_rect') or {
                                'x': int(base_img.size[0] * 0.7),
                                'y': int(base_img.size[1] * (active_m['row_num'] / 100)),
                                'width': 120,
                                'height': 40
                            }
                            x, y, w, h = int(current_rect['x']), int(current_rect['y']), int(current_rect['width']), int(current_rect['height'])
                            default_coords = (x, y, x + w, y + h)
                            
                            box_coords = st_cropper(
                                cropper_img,
                                realtime_update=True,
                                box_color='#EF4444',
                                aspect_ratio=None,
                                default_coords=default_coords,
                                return_type='box',
                                key=f"cropper_canvas_{active_idx}"
                            )
                            
                            if box_coords:
                                active_m['value_rect'] = {
                                    'x': box_coords['left'],
                                    'y': box_coords['top'],
                                    'width': box_coords['width'],
                                    'height': box_coords['height']
                                }
                                active_m['is_manual_crop'] = True
                                active_m['is_blank'] = False
                                
                            col_save, col_cancel = st.columns(2)
                            with col_save:
                                if st.button("💾 Save & Run OCR", key=f"save_mouse_crop_{active_idx}", use_container_width=True, type="primary"):
                                    with st.spinner("Analyzing handwriting..."):
                                        manual_crop = crop_match_region(pages[page_idx], active_m['anchor_rect'] or active_m['value_rect'], active_m['value_rect'])
                                        
                                        if "TrOCR" in ocr_engine_choice:
                                            new_text = ocr_service.run_trocr_on_crop(manual_crop)
                                            active_m['confidence'] = 0.88
                                        elif "Tesseract" in ocr_engine_choice:
                                            tess_res = ocr_service.perform_ocr(manual_crop, "Tesseract OCR", tesseract_path, lang_code)
                                            new_text = tess_res[0]['text'] if tess_res else ""
                                            active_m['confidence'] = tess_res[0]['confidence'] if tess_res else 0.50
                                        else:
                                            win_res = ocr_service.perform_ocr(manual_crop, "Windows Media OCR")
                                            new_text = " ".join([b['text'] for b in win_res]) if win_res else ""
                                            active_m['confidence'] = 0.85
                                            
                                        active_m['value'] = new_text
                                        st.session_state.mappings[active_idx]['value'] = new_text
                                        st.session_state.active_adjust_row = None
                                        st.rerun()
                            with col_cancel:
                                if st.button("❌ Cancel", key=f"cancel_mouse_crop_{active_idx}", use_container_width=True):
                                    st.session_state.active_adjust_row = None
                                    st.rerun()
                        else:
                            annotated_toggle = st.radio(
                                "Display Type",
                                options=["Original Scanned Page", "Highlighted OCR Detections"],
                                horizontal=True,
                                key="toggle_original_highlight_preview"
                            )
                            
                            vis_m = []
                            for pm in page_mappings:
                                if pm.get('anchor_rect') or pm.get('value_rect'):
                                    vis_m.append({
                                        'cell': f"Row {pm['row_num']}",
                                        'anchor_rect': pm.get('anchor_rect'),
                                        'value_rect': pm.get('value_rect'),
                                        'is_manual_crop': pm.get('is_manual_crop', False)
                                    })
                                    
                            page_ocr_data = st.session_state.ocr_results.get(active_pdf_name, {}).get(page_idx, [])
                            if "Highlighted" in annotated_toggle:
                                display_img = visualizer.draw_visual_highlights(base_img, page_ocr_data, vis_m, draw_gray_layout=True)
                            else:
                                display_img = visualizer.draw_visual_highlights(base_img, page_ocr_data, vis_m, draw_gray_layout=False)
                                
                            if zoom_factor != 1.0:
                                w, h = display_img.size
                                display_img = display_img.resize((int(w * zoom_factor), int(h * zoom_factor)), Image.Resampling.LANCZOS)
                                
                            st.image(display_img, use_container_width=True, caption=f"{active_pdf_name} - Page {active_page}")
                    else:
                        st.info("Upload matching PDF scan file to activate page highlights visualizer.")
                        
                # --- RIGHT PANEL: Editable Extracted Form ---
                with panel_right:
                    st.markdown('<div class="column-title">✏️ Extracted Data & Excel Verification</div>', unsafe_allow_html=True)
                    
                    if not page_mappings:
                        st.info("No matching mapped fields found on this page.")
                    else:
                        with st.container(height=750):
                            for order_idx, (global_idx, m) in enumerate(zip(page_mappings_indices, page_mappings)):
                                de_lbl = m['de_label'] or "(Empty)"
                                en_lbl = m['en_label'] or "(Empty)"
                                
                                st.markdown(f"""
                                <div style="background: rgba(16, 185, 129, 0.04); padding: 8px 12px; border-radius: 8px; margin-bottom: 5px; border-left: 4px solid #10b981;">
                                    <strong>Row {m['row_num']}</strong> | DE: <em>"{de_lbl}"</em> | EN: <em>"{en_lbl}"</em>
                                </div>
                                """, unsafe_allow_html=True)
                                
                                col_crop, col_edit = st.columns([1, 2])
                                
                                with col_crop:
                                    if m.get('is_blank'):
                                        st.markdown('<div style="background-color: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 6px; padding: 25px 5px; text-align: center; color: #64748b; font-size: 0.8rem;">Empty cell (Blank)</div>', unsafe_allow_html=True)
                                    elif not pdf_bytes_missing and m.get('anchor_rect') and m.get('value_rect'):
                                        try:
                                            crop_img = crop_match_region(pages[page_idx], m['anchor_rect'], m['value_rect'])
                                            st.image(crop_img, use_container_width=True, caption="Paper Crop Snippet")
                                        except Exception:
                                            st.caption("Crop error")
                                    else:
                                        st.caption("No snippet crop available")
                                        
                                with col_edit:
                                    if m.get('is_auto_corrected'):
                                        st.markdown(f'<span class="status-badge badge-success">✨ Auto-Corrected: {m.get("correction_reason")}</span>', unsafe_allow_html=True)
                                        
                                    val_display = m['value']
                                    if val_display == "" or val_display == "Missing Data":
                                        val_display = "Missing Data"
                                        st.markdown('<span class="status-badge badge-error">Missing Data</span>', unsafe_allow_html=True)
                                        
                                    new_val = st.text_input(
                                        f"Value (Row {m['row_num']}):",
                                        value=m['value'],
                                        key=f"field_{global_idx}",
                                        label_visibility="collapsed"
                                    )
                                    st.session_state.mappings[global_idx]['value'] = new_val
                                    
                                    if new_val != m['original_value'] and m['original_value']:
                                        st.caption(f"Original OCR Text: `'{m['original_value']}'`")
                                        
                                    conf = m['confidence']
                                    
                                    # Form 90%-70% Confidence UI Badges
                                    if conf >= 0.90:
                                        conf_class = "badge-success"
                                        conf_lbl = f"High ({int(conf*100)}%) - Auto Approved"
                                    elif conf >= 0.70:
                                        conf_class = "badge-warning"
                                        conf_lbl = f"Medium ({int(conf*100)}%) - Needs Review"
                                    else:
                                        conf_class = "badge-error"
                                        conf_lbl = f"Low ({int(conf*100)}%) - Force Human Check"
                                        
                                    st.markdown(f"Confidence: <span class=\"status-badge {conf_class}\">{conf_lbl}</span>", unsafe_allow_html=True)
                                    if conf < 0.70:
                                        st.markdown('<span class="status-badge badge-error" style="color: white; background-color: #ef4444;">⚠️ Requires Review & Human Edit (บังคับตรวจสอบ)</span>', unsafe_allow_html=True)
                                        
                                    # Auto Approve logic for >= 90%
                                    default_approved = m['approved']
                                    if conf >= 0.90 and not m.get('is_manual_crop', False):
                                        default_approved = True
                                        
                                    approved = st.checkbox(
                                        "Approve Value",
                                        value=default_approved,
                                        key=f"appr_{global_idx}"
                                    )
                                    st.session_state.mappings[global_idx]['approved'] = approved
                                    
                                with st.expander("⚙️ Adjust OCR Area (ปรับแต่งพื้นที่ OCR)"):
                                    if not pdf_bytes_missing:
                                        if st.button("🖱️ Draw Crop Area with Mouse (ลากครอบด้วยเมาส์)", key=f"btn_mouse_{global_idx}", use_container_width=True):
                                            st.session_state.active_adjust_row = global_idx
                                            st.rerun()
                                    else:
                                        st.caption("Upload PDF file to enable mouse crop adjustments.")
                                        
                                    if not pdf_bytes_missing:
                                        w_img, h_img = pages[page_idx].size
                                        current_rect = m.get('value_rect') or {
                                            'x': int(w_img * 0.7),
                                            'y': int(h_img * (m['row_num'] / 100)),
                                            'width': 120,
                                            'height': 40
                                        }
                                        
                                        new_x = st.slider("X Position", 0, w_img, int(current_rect['x']), key=f"x_sld_{global_idx}")
                                        new_y = st.slider("Y Position", 0, h_img, int(current_rect['y']), key=f"y_sld_{global_idx}")
                                        new_w = st.slider("Width", 10, 600, int(current_rect['width']), key=f"w_sld_{global_idx}")
                                        new_h = st.slider("Height", 10, 300, int(current_rect['height']), key=f"h_sld_{global_idx}")
                                        
                                        m['value_rect'] = {'x': new_x, 'y': new_y, 'width': new_w, 'height': new_h}
                                        m['is_manual_crop'] = True
                                        m['is_blank'] = False
                                        
                                        manual_crop = crop_match_region(pages[page_idx], m['anchor_rect'] or m['value_rect'], m['value_rect'])
                                        if st.button("🔄 Rescan Adjusted Area", key=f"rescan_{global_idx}"):
                                            with st.spinner("Running OCR..."):
                                                if "TrOCR" in ocr_engine_choice:
                                                    new_text = ocr_service.run_trocr_on_crop(manual_crop)
                                                    m['confidence'] = 0.88
                                                elif "Tesseract" in ocr_engine_choice:
                                                    tess_res = ocr_service.perform_ocr(manual_crop, "Tesseract OCR", tesseract_path, lang_code)
                                                    new_text = tess_res[0]['text'] if tess_res else ""
                                                    m['confidence'] = tess_res[0]['confidence'] if tess_res else 0.50
                                                else:
                                                    win_res = ocr_service.perform_ocr(manual_crop, "Windows Media OCR")
                                                    new_text = " ".join([b['text'] for b in win_res]) if win_res else ""
                                                    m['confidence'] = 0.85
                                                    
                                                m['value'] = new_text
                                                st.session_state.mappings[global_idx]['value'] = new_text
                                                st.rerun()
                                                
                                st.markdown("<hr style='margin: 8px 0; border: 0.5px solid #f1f5f9;'/>", unsafe_allow_html=True)
                                
                        # Save & Apply Corrections
                        if st.button("💾 Apply Corrections & Save to Database", use_container_width=True, type="primary"):
                            with st.spinner("Saving verification database..."):
                                for global_idx in page_mappings_indices:
                                    entry = st.session_state.mappings[global_idx]
                                    original = entry['original_value']
                                    corrected = entry['value']
                                    label = entry['de_label'] or entry['en_label']
                                    conf = entry['confidence']
                                    
                                    # Update SQLite review database
                                    audit_service.update_field_value(
                                        entry['id'],
                                        corrected,
                                        entry['approved'],
                                        conf
                                    )
                                    
                                    crop_img = None
                                    if not pdf_bytes_missing and entry.get('anchor_rect') and entry.get('value_rect'):
                                        crop_img = crop_match_region(pages[page_idx], entry['anchor_rect'], entry['value_rect'])
                                        
                                    # Record correction trail in TrainingData table
                                    audit_service.log_training_data(
                                        doc_id=st.session_state.loaded_doc_id,
                                        field_name=label,
                                        ocr_value=original,
                                        corrected_value=corrected,
                                        confidence=conf,
                                        status="Corrected" if corrected != original else "Approved"
                                    )
                                    
                                    # Also train visual patterns learning
                                    if corrected.strip() != original.strip() and corrected.strip() != "":
                                        audit_service.save_correction(
                                            original_ocr=original,
                                            corrected_text=corrected,
                                            crop_img=crop_img,
                                            anchor_label=label,
                                            confidence=conf
                                        )
                                        
                                    audit_service.log_ocr_transaction(
                                        file_name=entry['pdf_name'],
                                        page_num=entry['page'] + 1,
                                        extracted_text=original,
                                        corrected_text=corrected,
                                        confidence_score=conf
                                    )
                                    
                                audit_service.update_document_status(st.session_state.loaded_doc_id, "Reviewed")
                                
                            st.success("✅ Verification database saved and training trail recorded!")
                            st.rerun()
                            
                # Export Compile Section
                st.markdown("### 💾 Export & Compile Outputs")
                st.markdown('<div class="glass-card">', unsafe_allow_html=True)
                
                total_fields = len(st.session_state.mappings)
                approved_fields = sum(1 for m in st.session_state.mappings if m['approved'])
                low_conf_fields = sum(1 for m in st.session_state.mappings if m['confidence'] < 0.70 and not m['approved'])
                
                st.markdown(f"""
                - **Total Extracted Fields**: `{total_fields}`
                - **Approved / Verified Fields**: `{approved_fields} / {total_fields}`
                - **Pending Low-Confidence Fields (< 70% Review Required)**: `{low_conf_fields}`
                """)
                
                if low_conf_fields > 0:
                    st.warning(f"⚠️ Warning: There are still {low_conf_fields} unverified low-confidence fields (< 70%) that require review before exporting.")
                    
                if st.button("📥 Compile Final Excel Workbook", use_container_width=True, type="primary"):
                    with st.spinner("Writing outputs to Excel template..."):
                        final_excel_bytes = excel_service.write_values_to_excel(
                            st.session_state.excel_bytes,
                            st.session_state.mappings,
                            target_col
                        )
                        
                        # Find Order No and Serial No dynamically to construct filename
                        order_no = "OrderUnknown"
                        serial_no = "SerialUnknown"
                        for m in st.session_state.mappings:
                            lbl_de = (m.get('de_label') or '').lower()
                            lbl_en = (m.get('en_label') or '').lower()
                            if 'auftrag' in lbl_de or 'order' in lbl_en:
                                order_no = str(m['value']).strip() or "OrderBlank"
                            if 'bau-nr' in lbl_de or 'serial' in lbl_en:
                                serial_no = str(m['value']).strip() or "SerialBlank"
                                
                        output_filename = f"{order_no}_{serial_no}_Adjustment_Table_Completed.xlsx"
                        
                        # Ensure /Output/ folder exists
                        os.makedirs("Output", exist_ok=True)
                        output_filepath = os.path.join("Output", output_filename)
                        
                        # Save excel workbook to Output/ folder
                        with open(output_filepath, "wb") as f_out:
                            f_out.write(final_excel_bytes)
                            
                        # Log export history in SQLite
                        audit_service.log_export_history(
                            st.session_state.loaded_doc_id,
                            st.session_state.excel_filename,
                            output_filename,
                            "Success"
                        )
                        
                        audit_service.update_document_status(st.session_state.loaded_doc_id, "Exported")
                        
                        st.success(f"✅ Excel compiled successfully and saved to Output folder: `{output_filepath}`")
                        
                        st.download_button(
                            label="📥 Download Mapped Excel File (.xlsx)",
                            data=final_excel_bytes,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                st.markdown('</div>', unsafe_allow_html=True)
                
        else:
            st.markdown('<div class="glass-card" style="text-align: center; padding: 50px 20px;">', unsafe_allow_html=True)
            st.write("### 📂 Welcome to the OCR Adjustment Review Database System")
            st.write("To begin, upload your **Excel Template** and **Scanned PDF(s)**, or load an archived document from the database history list in the sidebar.")
            st.markdown('</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
