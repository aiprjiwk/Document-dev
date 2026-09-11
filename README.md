# 📊 Smart OCR & Excel Document Processor
(ระบบตรวจจับตัวเขียนและกรอกข้อมูลลงเซลล์ Excel อัตโนมัติ พร้อมระบบเรียนรู้การแก้ไขข้อมูล)

A modular, production-ready local application for scanning multiple handwritten PDF documents using advanced OCR, mapping data fields into Excel templates, enabling human-in-the-loop verification, and learning from manual corrections to improve handwriting recognition over time.

This version is specifically customized for scanned **"Adjustment Chart Master"** documents, mapping German and English bilingual rows to target Excel columns.

---

## 🛠️ Requirements & Installation Guide

### 1. Python Environment Setup
1. Download Python 3.10+ (Recommended: Python 3.11) from [python.org](https://www.python.org/downloads/).
   * ⚠️ **Important during installation**: Check the box **"Add python.exe to PATH"**.
2. Open your terminal in this directory and install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

### 2. Tesseract OCR Setup (For Tesseract engine)
To use Tesseract OCR, you must install the Tesseract binary:
1. Download the Windows installer from [UB-Mannheim Tesseract Wiki](https://github.com/UB-Mannheim/tesseract/wiki).
2. Install it. By default, it installs to `C:\Program Files\Tesseract-OCR\tesseract.exe`.
3. If installed in a custom directory, you can specify this path directly in the application's sidebar.
4. For German + English support, ensure **German (deu)** and **English (eng)** language data are selected during installation.

### 3. Poppler Setup (For pdf2image pdf conversion)
The application converts PDFs to images. It uses `pdf2image` (requires Poppler) and automatically falls back to `PyMuPDF` (does not require any external binaries).
To configure Poppler (optional for performance):
1. Download Poppler for Windows (e.g., from [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases)).
2. Extract the ZIP file and input the path to the `bin` directory of Poppler in the sidebar field (e.g., `C:\poppler\bin`).

---

## 🚀 How to Run the Application

You can launch the application instantly by running `run.bat` or executing the following command in your terminal:
```bash
python -m streamlit run app.py --server.port 8080 --server.address 0.0.0.0
```
Then, navigate to `http://localhost:8080` in your web browser.

---

## 💡 Key Custom Features

1. **Bilingual Anchor Matching**:
   * Reads German labels from **Column B** and English labels from **Column D** on the `"Master"` sheet of the template Excel file.
   * Matches these labels to spatial text blocks in the scanned PDF page.
2. **Neighborhood Value Mapping & Custom Writeback**:
   * Scans the neighborhood of the detected labels for handwriting values.
   * Prompts you in the sidebar to choose the target value column (e.g. Column E, F, G, H, or I) to populate.
3. **Focused Handwriting OCR & Blank Crop Bypass**:
   * Crops the isolated neighborhood value zone (where the handwriting is expected).
   * Analyzes the crop using OpenCV. If the zone has no pen/ink strokes, it skips the OCR entirely (leaving it blank) to prevent hallucination.
   * If strokes are detected, runs handwriting OCR (TrOCR, EasyOCR, or Windows Media OCR).
4. **Interactive Dual-Panel Form**:
   * **LEFT Panel**: Zoomable scanned PDF page visualizer highlighting mapped sections.
   * **RIGHT Panel**: Row-by-row editor aligning German/English labels with cropped visual snippets from the PDF, editable text fields, and confidence scores (red highlight if confidence < 60% for required review).
5. **Corrections & Learning Feedback**:
   * Saving manual edits logs the transaction and saves the visual average hash (`aHash`) and image snippet in `dataset_images/`.
   * Automatically auto-corrects matching handwriting patterns in future runs, indicating it with a `✨ Auto-Corrected` badge.

---

## 📁 File Structure & Modules

* **`app.py`**: Streamlit interface, split-screen panels, and workflow orchestration.
* **`ocr_engine.py`**: Preprocessing (grayscale, bilateral denoising, CLAHE, Otsu thresholding) and OCR engine wrappers (Windows Media, Tesseract, TrOCR).
* **`learning_module.py`**: User corrections database, Average Hash (`aHash`) calculation, and fuzzy text similarity mapping.
* **`excel_mapper.py`**: Specialized `"Master"` worksheet parsing (German Col B / English Col D) and target column cell writeback.
* **`utils.py`**: Fallback PDF-to-image conversion, CSV logging, and ZIP packaging.
* **`requirements.txt`**: Package dependencies.
* **`run.bat`**: Starter batch file.
