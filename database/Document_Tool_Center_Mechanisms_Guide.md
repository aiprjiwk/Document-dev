# 📚 คู่มือและสรุปกลไกการทำงานอย่างละเอียด: Document Tool Center (App Team System)

เอกสารฉบับนี้สรุปโครงสร้าง สถาปัตยกรรม และกลไกการทำงานอย่างละเอียดของทุกโมดูลใน **Document Tool Center** รวมถึงระบบ **📜 Calibration Certificate**, ชุดประมวลผล **📋 IQOQDQ Qualification Suite** และเครื่องมือ **🏷️ Rename Tag Tool**

---

## 📑 รายการโมดูลทั้งหมดใน Document Tool Center

1. [📜 Calibration Certificate Processing Suite (เครื่องมือจัดการใบรับรองการสอบเทียบ)](#1--calibration-certificate-processing-suite)
   - 1.1 📜 Calibration Processing Tool (โหมดประมวลผลแบบแปลนมาตรฐาน & Approval)
   - 1.2 🔢 Certificate No. Mapping Tool (ระบบจับคู่รหัสใบสอบเทียบจาก PDF กับ Excel)
   - 1.3 💼 Quotation Verification Tool (ระบบตรวจสอบความถูกต้องกับใบเสนอราคา)
   - 1.4 🖼️ Calibration Image Maintenance & Audit (ระบบตรวจเช็คและบำรุงรักษาคลังภาพเครื่องมือ)
2. [🤖 OCR & AI (from QC)](#2--ocr--ai-from-qc)
3. [🌟 Advanced OCR Adjustment](#3--advanced-ocr-adjustment)
4. [🏆 IWK Certificate](#4--iwk-certificate)
5. [🔍 ETK Verification](#5--etk-verification)
6. [🔧 Fault Assistance Suite](#6--fault-assistance-suite)
7. [📘 Machine Configuration System & Operating Manual](#7--machine-configuration-system--operating-manual)
8. [📋 IQOQDQ Qualification Suite (ระบบสร้างเอกสารรับรองคุณภาพเครื่องจักร)](#8--iqoqdq-qualification-suite)
   - 8.1 📐 DQ — Design Qualification Workspace
   - 8.2 🔧 IQ — Installation Qualification Workspace (IQ Installation, IQ CCI, IQ Format)
   - 8.3 ⚡ OQ — Operational Qualification Workspace (OQ IO List, OQ HMI, OQ Alarm, OQ Shift Register)
9. [🏷️ Rename Tag & Custom Document Properties Tool](#9--rename-tag--custom-document-properties-tool)

---

## 1. 📜 Calibration Certificate Processing Suite

ระบบ **Calibration Certificate** เป็นโมดูลหลักสำหรับประมวลผล ตรวจสอบ จับคู่ และจัดรูปแบบตารางใบรับรองการสอบเทียบเครื่องมือวัดของโรงงาน โดยแบ่งออกเป็น 4 ฟังก์ชันหลัก ดังนี้:

### 1.1 📜 Calibration Processing Tool (การประมวลผลตาราง Calibration)
* **วัตถุประสงค์**: ถอดรหัส VBA Macro เดิม แปลงตารางข้อมูลดิบ (Raw Data) ให้อยู่ในรูปแบบ Final Layout พร้อมดึงภาพเครื่องมือวัดและฝัง Header Logo อัตโนมัติ
* **ขั้นตอนและกลไกการทำงานอย่างละเอียด**:
  1. **Project Code Extraction**: อ่านชื่อไฟล์นำเข้า (เช่น `55782_IWK Calibration certificates.xlsx`) และสกัดรหัสโครงการ (เช่น `55782`)
  2. **Z-Value Capture**: สแกนตารางดั้งเดิมเพื่อดักจับค่า Z จาก Column B สำหรับแถวที่มีประเภทงานเป็น `CALIBRATION` ใน Column P
  3. **Row Filtering & Bold Formatting**:
     - ปรับฟอนต์ตัวหนา (Bold) สำหรับแถวหมวดหลัก (Column A == 2)
     - ลบแถวที่ไม่ใช่รายการ Calibration ออกจากตาราง
  4. **Row Splitting & Duplication**: แตกแถวซ้ำตามจำนวน Repeat Count ที่ระบุไว้ใน Column E
  5. **Serial No. Generation Formula**: สร้างรหัส Serial No. อัตโนมัติในรูปแบบ `XXXXX-Y-Z-PC`:
     - `XXXXX`: รหัสโครงการ (เช่น `56021`)
     - `Y`: ชื่อหมวดอุปกรณ์ดั้งเดิมจากแถวตัวหนาล่าสุด (Column B)
     - `Z`: ค่า Z ดั้งเดิมที่ดักจับได้
     - `PC`: ลำดับตัวนับซ้ำของอุปกรณ์ในกลุ่มนั้น
  6. **Dynamic Image Lookup & Insertion**:
     - ค้นหาภาพเครื่องมือวัดในคลังโฟลเดอร์ `My Picture/` อัตโนมัติ โดยจับคู่ตามชื่ออุปกรณ์ (`{prefix}@{item_no}`) รองรับนามสกุล `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`
     - คำนวณอัตราส่วนย่อขนาดภาพ (Constrain Bounds) สูงสุดไม่เกิน ความกว้าง 100px และความสูง 110px
     - กำหนดความสูงของแถวใน Excel เป็น 90pt และแทรกรูปภาพลงใน Column `Picture`
  7. **Approval Layout Option**:
     - หากเลือกโหมด Approval Layout ระบบจะแทรกคอลัมน์ `IWK production check` และ `Supplier Check` พร้อมสัญลักษณ์กล่องติ๊ก `[ ]`
     - แทรกบรรทัดลงนามท้ายกระดาษ (Footer) สำหรับ PM, Production และ Supplier
  8. **Header Logo & Page Setup**:
     - ปรับหน้ากระดาษเป็นแนวนอน (Landscape Fit-to-Page)
     - ฝังภาพ **IWK Logo** (`My Picture/logo IWK.png`) ลงในตำแหน่ง Top Left Header ผ่าน openpyxl และ COM Automation (pywin32)

---

### 1.2 🔢 Certificate No. Mapping Tool (ระบบจับคู่รหัสใบสอบเทียบ)
* **วัตถุประสงค์**: สแกนอ่านไฟล์ PDF ใบสอบเทียบของซัพพลายเออร์ ดึงรหัส Certificate No. และจับคู่ลงตาราง Excel อัตโนมัติ
* **ขั้นตอนและกลไกการทำงานอย่างละเอียด**:
  1. **OCR PDF Scanning**:
     - รองรับการอัปโหลดไฟล์สแกน PDF หลายไฟล์พร้อมกัน
     - ใช้ **PyMuPDF** ดึงข้อความดิจิทัล หากเป็นไฟล์สแกนรูปภาพ จะรัน **WinOCR** แปลงภาพเป็นข้อความด้วยความละเอียดสูง 200 DPI
  2. **Data Extraction (Regex)**:
     - สกัดรหัส **Serial No.** (เช่น `56021-2532979-40-1`) และรหัส **Certificate No.** (เช่น `PI-1407003/26`) จากเนื้อหาใน PDF
  3. **Dynamic Column Lookup & Value Writeback**:
     - ค้นหาตำแหน่งคอลัมน์ `"Serial No."` และ `"Certificate No."` ในไฟล์ Excel (Step 2) แบบ Dynamic
     - เติมรหัส `Certificate No.` ลงใน Excel ด้วยตัวอักษร **สีดำสนิท (`#000000`) ตัวปกติ (Non-Bold / Normal weight)**
  4. **Unmatched Entire Row Yellow Highlighting**:
     - หากรายการใดใน Excel ไม่มีไฟล์ PDF มาจับคู่ ระบบจะ **ระบายสีเหลือง (`#FFFF00`) ทั้งแถว** (ตั้งแต่ Column A ถึง Column I) เพื่อเตือนผู้ตรวจสอบ
  5. **Renamed PDF Sequence Prefix**:
     - ดึงลำดับรายการ 3 หลักจาก Column A (`No.`) มาตั้งชื่อไฟล์ PDF ใหม่ เช่น `001_PI-1407003_26.pdf`
  6. **Unmatched PDF Exclusion**:
     - คัดกรองไฟล์ PDF ที่ไม่ตรงออกจาก ZIP Package (บรรจุเฉพาะไฟล์ Excel และไฟล์ PDF ที่จับคู่สำเร็จเท่านั้น)
  7. **Header Logo Scaling**:
     - ฝัง **IWK Logo** ย่อขนาดลง 20% (ความกว้าง 57.6 pt) ใน Left Header ของไฟล์ Excel

---

### 1.3 💼 Quotation Verification Tool (การตรวจสอบใบเสนอราคา)
* **วัตถุประสงค์**: เปรียบเทียบข้อมูลอุปกรณ์ในใบเสนอราคากับตารางสอบเทียบจริง
* **ขั้นตอนและกลไกการทำงานอย่างละเอียด**:
  1. เปรียบเทียบรหัสอุปกรณ์, Part Number, สเปกความหนา และยี่ห้อ
  2. สรุปผลการตรวจสอบรายการที่ตรงกันและรายการที่มีความคลาดเคลื่อนเพื่ออนุมัติการสั่งซื้อ

---

### 1.4 🖼️ Calibration Image Maintenance & Audit (ระบบจัดการคลังภาพเครื่องมือ)
* **วัตถุประสงค์**: บริหารจัดการ อัปเดต และตรวจสอบความครบถ้วนของรูปภาพเครื่องมือวัดในคลัง `My Picture/`
* **ขั้นตอนและกลไกการทำงานอย่างละเอียด**:
  1. **Image Audit**: สแกนตาราง Excel เพื่อตรวจสอบหาภาพเครื่องมือวัดที่พบ (Found) และภาพที่ยังขาดอยู่ (Missing)
  2. **Image Upload & Save**: อัปโหลดภาพเครื่องมือวัดใหม่ลงโฟลเดอร์ `My Picture/` พร้อมตั้งชื่อตามโครงสร้างมาตรฐาน
  3. **Bundle Package**: รวบรวมภาพในคลังแพ็กเกจเป็นไฟล์ ZIP ดาวน์โหลด

---

## 2. 🤖 OCR & AI (from QC)
* **วัตถุประสงค์**: ตรวจจับ บาร์โค้ด และสกัดข้อความจากเอกสารสแกน PDF
* **กลไกการทำงาน**:
  1. แยกหน้า PDF และสแกนหาบาร์โค้ด 1D/2D ด้วย `zxingcpp` / `pyzbar`
  2. หากไม่พบบาร์โค้ด จะประมวลผล OCR ด้วย **Windows Media OCR (WinOCR)** หรือ **Surya OCR**
  3. แสดงหน้าจอเปรียบเทียบภาพกับตารางแก้ไขข้อความ (Human-in-the-loop)
  4. บันทึกค่า Average Hash (`aHash`) และภาพตัดย่อยลง `dataset_images/` เพื่อเรียนรู้คำอ่านลายมืออัตโนมัติ
  5. บันทึกผลลัพธ์ลงฐานข้อมูล SQLite (`database/ocr_system.db`)

---

## 3. 🌟 Advanced OCR Adjustment
* **วัตถุประสงค์**: ถอดรหัสตาราง Adjustment Chart Master ภาษาเยอรมัน-อังกฤษ
* **กลไกการทำงาน**:
  1. อ่านป้ายกำกับภาษาเยอรมัน (Col B) และภาษาอังกฤษ (Col D) ใน Sheet `"Master"`
  2. ค้นหาพิกัดคำสำคัญบน PDF และสแกนพื้นที่ข้างเคียง (Neighborhood Zone)
  3. ใช้ OpenCV กรองจังหวะหมึก หากพื้นที่ว่างเปล่าจะข้าม OCR ทันทีเพื่อป้องกัน Hallucination
  4. ประมวลผลลายมือเฉพาะจุดด้วย TrOCR / Windows OCR

---

## 4. 🏆 IWK Certificate
* **วัตถุประสงค์**: แปลงโครงสร้าง BOM สำหรับใบรับรองชิ้นส่วนสัมผัสผลิตภัณฑ์ (Product Contact Parts)
* **กลไกการทำงาน**:
  1. จัดกลุ่มประเภทใบรับรอง: `WAZ and FAD`, `OZ`, `SZ`, `OMP`
  2. ตัดแถวและคอลัมน์ที่ไม่เกี่ยวข้อง เช่น "Electrical parts" และย้ายคอลัมน์ K ไป Q
  3. ปรับฟอนต์ตัวหนา และใส่สีไฮไลท์ (แดง = ไม่ผ่าน Macro, เขียว = พบ PDF, เหลือง = ขาด PDF)

---

## 5. 🔍 ETK Verification
* **วัตถุประสงค์**: ตรวจสอบรายการอะไหล่และชิ้นส่วนเครื่องจักรกับรหัส ETK Catalog
* **กลไกการทำงาน**:
  1. จับคู่รหัส Component Number และ Drawing Number ระหว่างตาราง BOM กับ ETK Catalog
  2. คำนวณตรวจสอบผลรวมจำนวนชิ้นส่วนอุปกรณ์ (Quantity Verification)
  3. ออกรายงานสรุปรายการที่ถูกต้องและรายการที่มีความคลาดเคลื่อน

---

## 6. 🔧 Fault Assistance Suite
* **วัตถุประสงค์**: จัดการและแพ็กเกจไฟล์คู่มือแก้ไขข้อผิดพลาด (Fault HTML Maintenance)
* **กลไกการทำงาน**:
  1. จับคู่รหัส Fault Code จาก Excel กับไฟล์ HTML ในโฟลเดอร์ `Fault assistance/`
  2. แทนที่รหัสโครงการและชื่อลูกค้าภายในไฟล์ HTML อัตโนมัติ
  3. รวบรวมไฟล์ HTML รูปภาพประกอบ และรายงานสรุป Excel แพ็กเกจลงไฟล์ ZIP

---

## 7. 📘 Machine Configuration System & Operating Manual
* **วัตถุประสงค์**: รวบรวม จัดหมวดหมู่ และส่งมอบคู่มือการใช้งานเครื่องจักร
* **กลไกการทำงาน**:
  1. ดึงไฟล์ Datasheet และคู่มือทางเทคนิคตามรุ่นเครื่องจักร
  2. ตรวจสอบความสมบูรณ์ของโครงสร้างเอกสารเพื่อสร้างไฟล์แพ็กเกจส่งมอบลูกค้า

---

## 8. 📋 IQOQDQ Qualification Suite

ระบบสร้างและบริหารจัดการเอกสารรับรองคุณภาพเครื่องจักร (Qualification System) แบ่งตามระยะ Qualification (DQ, IQ, OQ) พร้อมรองรับการดาวน์โหลดและอัปโหลดแม่แบบ Word Master Template ภาษาอังกฤษในทุกโมดูล:

### 8.1 📐 DQ — Design Qualification Workspace
* **วัตถุประสงค์**: พื้นที่ทำงานสำหรับสร้างและจัดการเอกสารรับรองการออกแบบเครื่องจักร (Design Qualification Protocol & Report)
* **การจัดการแม่แบบ**: จัดเก็บไฟล์แม่แบบ Word ภายใต้โฟลเดอร์ `IQOQDQ/DQ_Design/` พร้อมปุ่มดาวน์โหลดและอัปโหลดแม่แบบใหม่

### 8.2 🔧 IQ — Installation Qualification Workspace
ประกอบด้วย 3 โมดูลย่อยหลักสำหรับประมวลผลการติดตั้งเครื่องจักร:
1. **🔧 IQ Installation Protocol Generator**:
   - อ่านข้อมูลจากไฟล์ Excel Parts List (`IQOQDQ/IQ_Installation/`)
   - กรองและเติมข้อมูลชิ้นส่วนลงใน Table 5 ของไฟล์แม่แบบ Word `02_iq_installation`
2. **🛡️ IQ CCI (Control Components & Instruments)**:
   - อ่านข้อมูลจาก Sheet ที่ขึ้นต้นด้วย `IQOQ list*` (Columns B, F, H) ในไฟล์ Excel Parts List
   - ประมวลผลและเติมข้อมูลอุปกรณ์ควบคุมลงใน Table 6 ของไฟล์แม่แบบ Word `05_iq_control` (`IQOQDQ/IQ_CCI/`)
3. **📦 IQ Format Parts Protocol Generator**:
   - ประมวลผลข้อมูล BOM จากไฟล์ Excel (`EXPORT_*.XLSX`)
   - กรองแถว/คอลัมน์ จัดกลุ่มตาม Component Designation และเติมลงในไฟล์แม่แบบ Word (`IQOQDQ/IQ_Format/`)

### 8.3 ⚡ OQ — Operational Qualification Workspace
ประกอบด้วย 4 โมดูลย่อยสำหรับประมวลผลการสอบเทียบและทดสอบการทำงานของเครื่องจักร:
1. **📋 OQ IO List Protocol Generator**:
   - ถอดรหัส ELCAD IO List Excel (`IQOQDQ/IO list/`)
   - กรองค่าว่าง เรียงลำดับตาม Column E (Cross Reference) คำนวณสูตร Address / Description / Page / Test
   - เติมข้อมูลจุด IO ทั้งหมดลงใน Table 5 ของไฟล์แม่แบบ Word `03_oq_io`
2. **🖥️ OQ HMI Protocol Generator (OCR Image System)**:
   - **Interactive Existing Recheck & Sequence Control**: สแกนภาพจับคู่กับคลัง Expected Header Titles ประจำรุ่นเครื่องจักร (เช่น `FP`, `SC 5`) และประเภท Visu (`IPC` / `Magilis`)
   - **Excel Checklist Tool**: ส่งออกไฟล์ `.xlsm` พร้อม VBA Macro และ Checkbox ให้ผู้ใช้รันลำดับ 1, 2, 3... ก่อนอัปโหลดกลับเพื่อจัดลำดับภาพ
   - **Word Document Population**: ประมวลผล Image OCR แทรกรูปภาพและขึ้นหน้าใหม่ (Page Break) อัตโนมัติ เพื่อแทนที่ข้อความมาร์กเกอร์ `XXXX` ใต้หัวข้อ `3.3 Masks` ในไฟล์แม่แบบ Word `XXXXX_04_OQ_HMI` (`IQOQDQ/OQ_HMI/`)
   - **Audit Report Export**: สร้างรายงานสรุป Excel ตรวจสอบการแทรกภาพ (Inserted vs Not Inserted)
3. **🚨 OQ Alarm Protocol Generator**:
   - อ่านตัวแปร Alarm จากไฟล์ Excel (เช่น `5XXXX-AlarmInfo.xlsx`)
   - ค้นหาและจับคู่แม่แบบแบบทดสอบ Alarm (`MX_*`) ในคลังประจำเครื่องจักร (`IQOQDQ/OQ Alarm/GMP_Alarme/{Machine_Type}/EN/`)
   - ประกอบตารางทดสอบ Alarm และตารางลงนาม Performer Sign-off ลงใน Section 3 ของไฟล์แม่แบบ Word `XXXXX_10_OQ_Alarms`
   - สร้างไฟล์รายงานสรุป Matched Alarms Excel อัตโนมัติ
4. **🔄 OQ Shift Register Protocol Generator**:
   - อ่านตัวแปร Shift Register จากไฟล์ Excel (เช่น `5XXXX-ShiftRegisterInfo.xlsx`)
   - ค้นหาและจับคู่แม่แบบ Shift Register ในคลังประจำเครื่องจักร (`IQOQDQ/OQ_Shift/GMP_Shift Register/{Machine_Type}/EN/`)
   - ประกอบตารางทดสอบ Shift Register และตารางลงนาม Performer Sign-off ลงใน Section 3 ของไฟล์แม่แบบ Word `XXXXX_11_OQ_Shift Register`
   - สร้างไฟล์รายงานสรุป Matched Shift Register Excel อัตโนมัติ

*หมายเหตุ: ทุกโมดูลย่อยใน IQOQDQ มาพร้อมแท็บ **⚙️ Machine Type & Template Maintenance** สำหรับสร้าง/ลบเครื่องจักร อัปโหลด และจัดการไฟล์แม่แบบในคลังอย่างสะดวก*

---

## 9. 🏷️ Rename Tag & Custom Document Properties Tool

* **วัตถุประสงค์**: เครื่องมือประมวลผลไฟล์เอกสาร Word (`.docx`, `.doc`) แบบกลุ่ม (Batch Processing) เพื่ออัปเดตข้อมูล Custom Document Properties, วันที่ในตารางประวัติเอกสาร และเปลี่ยนชื่อไฟล์ตามมาตรฐานองค์กร IWK
* **ขั้นตอนและกลไกการทำงานอย่างละเอียด**:
  1. **Batch Custom Document Properties Writeback**:
     - สแกนองค์ประกอบ XML ใน `docProps/custom.xml` ของไฟล์เอกสาร Word
     - เขียนอัปเดต 6 Custom Properties สำคัญ ได้แก่:
       - `Copyright`: ค่าเริ่มต้น `© Copyright by IWK (Thailand) Limited 2026`
       - `Version`: ค่าเริ่มต้น `0.1`
       - `Machine`: ค่าเริ่มต้น `IWK XX`
       - `Order`: รหัสคำสั่งซื้อ 5 หลัก (เช่น `56021`)
       - `Baunummer`: หมายเลขเครื่อง (เช่น `56021`)
       - `Bezeichnung`: ชื่อประเภทเอกสาร Qualification Protocol
  2. **Document History Table Date Synchronization**:
     - สแกนตารางประวัติการแก้ไขเอกสาร (Document History Table) ภายในไฟล์ Word
     - แปลงและอัปเดตวันที่ทั้งหมดให้อยู่ในฟอร์แมตมาตรฐาน `DD-MMM-YYYY` (เช่น `02-Jan-2025`)
  3. **DOCPROPERTY Dynamic XML Field Update**:
     - เปิดใช้งานสวิตช์ `<w:updateFields w:val="true"/>` ใน `word/settings.xml` เพื่อสั่งให้ Word อัปเดตฟิลด์อัตโนมัติเมื่อเปิดไฟล์
     - สแกนและอัปเดตข้อความในแท็ก `<w:t>` สำหรับทุกฟิลด์ `DOCPROPERTY` ที่ใช้อ้างอิงใน Header, Footer, Table และ Paragraphs ทั้งหมด
  4. **Standardized Automated File Renaming**:
     - เปลี่ยนชื่อไฟล์เอกสาร Word อัตโนมัติ โดยแทนที่สัญลักษณ์มาร์กเกอร์ `XXXXX` ด้วย Order No. 5 หลัก และปรับรูปแบบวันที่เป็น ISO Date `YYYY-MM-DD`
  5. **ZIP Package Export**:
     - รวบรวมไฟล์ Word ที่ได้รับการอัปเดตข้อมูลและเปลี่ยนชื่อเรียบร้อยแล้วทั้งหมด แพ็กเกจเป็นไฟล์ ZIP ดาวน์โหลดในคราวเดียว

---
*เอกสารนี้ได้รับการอัปเดตล่าสุดสำหรับระบบ IWK Document Portal*
