# FAT Protocol Generator & Maintenance Tool

Automated tool for generating **FAT Protocol (.docx)** documents directly from **Excel (.xlsx)** calculation sheets, complete with a central category maintenance system (Master Rules).

---

## 🌟 Key Features
1. **Excel Import (Multi-sheet support):**
   - Each sheet is mapped into a **Level 1** section (e.g. `2 Test result sheets IWK TZC`, `3 Test result sheets ...`)
   - Text from **Column B (Kurztext / Information)** is placed into the `Function / Installation / Document` column in Word
2. **Project Metadata:**
   - `Project:`
   - `Order no.:`
   - `Customer:`
   - `Version:`
   - Automatically populates the document header table (Table 0) and page headers
3. **Category Verification & Master Rules System:**
   - Verifies PosNo or Kurztext against Master Rules
   - New items not found in the database are flagged as `⭐ New Item` and highlighted for review
   - Users can edit and reassign categories (Level 1, Level 2, Level 3) directly in the interactive table
   - Dedicated **🛠️ Master Rules Maintenance** tab to add, edit, or delete rules anytime
4. **Table Formatting & Hierarchical Structure in Word:**
   - **Level 2:** `Documents` and `Safety` (Safety is always included per specifications)
   - **Level 3:** Sub-categories under `Machine details` (empty categories without data are automatically pruned)
   - Automatically numbers items in the `No.` column (e.g., `2.1.1`, `2.1.2`, `1`, `2`, ...)
   - Retains original checkbox styling `☐ Yes ☐ No` exactly per the template
5. **Download:**
   - Click **🚀 Generate FAT Protocol (Word)** to build and download the final `.docx` file immediately

---

## 🚀 How to Run Standalone
Double-click:
```bat
start.bat
```
The application will launch in your browser at `http://localhost:8503`.

---

## 📁 File Structure
- `app.py`: Standalone Streamlit Web UI
- `fat_generator.py`: Processing engine for Excel reading, category mapping, and Word (.docx) generation
- `master_rules.json`: Central category rules database
- `start.bat`: One-click startup script
- `templates/`: FAT Protocol Word templates
