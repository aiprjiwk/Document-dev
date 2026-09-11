# PROJECT REQUIREMENTS

## PROJECT OVERVIEW
### PROJECT NAME
OCR Adjustment Table Automation

### PURPOSE
The application shall read handwritten values from scanned PDF files, allow human verification, store OCR and corrected values, populate an Excel template, and export a completed Excel file.

The application is intended for the Documentation Department workflow.

---

## BUSINESS PROCESS
### Current Process
1. User receives scanned PDF from machine commissioning.
2. Handwritten values exist on the PDF.
3. User manually reads values.
4. User manually enters values into Adjustment Table Excel.
5. User exports completed Excel.

### Target Process
1. Upload Scan PDF.
2. OCR extracts handwritten values.
3. Human reviews and corrects values.
4. Approved values stored in database.
5. System fills Excel template automatically.
6. Excel exported automatically.

---

## SYSTEM ARCHITECTURE
* **Frontend**: React
* **Backend**: FastAPI
* **Database**: SQLite
* **OCR Engine**: Azure Document Intelligence
* **Excel Engine**: openpyxl

---

## DATABASE DESIGN
### TABLE: `ocr_data`
Tracks OCR results, user corrections, and confidence levels.
* `id`: INTEGER PRIMARY KEY AUTOINCREMENT
* `document_id`: TEXT / VARCHAR
* `field_name`: TEXT / VARCHAR
* `ocr_value`: TEXT / VARCHAR
* `corrected_value`: TEXT / VARCHAR
* `confidence`: REAL / FLOAT
* `reviewed_by`: TEXT / VARCHAR
* `reviewed_date`: TEXT / DATETIME
* `status`: TEXT / VARCHAR (e.g. `"Pending"`, `"Reviewed"`, `"Exported"`)

---

## EXCEL MAPPING DESIGN
### Mapping Sheet Example
| FieldName | Cell |
| :--- | :--- |
| OrderNo | B3 |
| SerialNo | B4 |
| CaseDischarge | F12 |
| CaseTurning | F15 |
| Pusher | F21 |

*Note: Target cell values are handwritten on the document.*

---

## APPLICATION MODULES

### MODULE 1: OCR Service
**Responsibilities**:
* Receive PDF files.
* Call Azure Document Intelligence layout analysis.
* Return extracted text fields.
* Return extraction confidence scores.

### MODULE 2: Human Verification
**Responsibilities**:
* Display original scanned PDF pages.
* Display OCR results side-by-side.
* Allow manual corrections and overrides.
* Save approved cell data.

### MODULE 3: Database Service
**Responsibilities**:
* Store raw OCR values.
* Store human-corrected values.
* Store confidence scores.
* Keep detailed audit trail log.

### MODULE 4: Mapping Service
**Responsibilities**:
* Read Excel Mapping Sheets.
* Load cell locations.
* Match FieldName with designated Target Cell.

### MODULE 5: Excel Export Service
**Responsibilities**:
* Load Excel templates.
* Populate spreadsheet cells dynamically based on Mapping Table.
* Export and save completed files.

---

## CODING RULES
1. **NEVER** hardcode cell addresses.
2. Always read the Mapping Sheet dynamically.
3. Use modular service-oriented architecture.
4. Use Python classes.
5. Use type hints.
6. Add comprehensive logging.
7. Add robust exception handling.
8. Separate frontend and backend.

---

## TASK LIST
* **Task 1**: Create SQLite database structure (Generate `models.py` only. Do not create frontend, OCR, or Excel module).
* **Task 2**: Create OCR Service (Use FastAPI, read PDF, call OCR engine, return JSON. Do not modify `models.py`).
* **Task 3**: Create Mapping Service (Read Mapping sheet from Excel, return dictionary).
* **Task 4**: Create Excel Export Service (Use openpyxl, populate template by mapping).
