import os
import io
import re
import json
import datetime
import tempfile
import pandas as pd
import openpyxl
import docx

DEFAULT_TEMPLATE_PATH = os.path.abspath(r"C:\Users\cpreephim\Desktop\App Team\IQOQDQ\IQ_Installation\XXXXX_02_IQ_Installation_20XX-XX-XX_en.docx")
DEFAULT_TAG_MAPPING_PATH = os.path.abspath(r"C:\Users\cpreephim\Desktop\App Team\IQOQDQ\IQ_Installation\iq_installation_tag_mapping.json")

def load_tag_mapping():
    """
    Loads tag mapping configuration from JSON file.
    """
    if os.path.exists(DEFAULT_TAG_MAPPING_PATH):
        try:
            with open(DEFAULT_TAG_MAPPING_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_tag_mapping(tag_data):
    """
    Saves tag mapping configuration to JSON file.
    """
    os.makedirs(os.path.dirname(DEFAULT_TAG_MAPPING_PATH), exist_ok=True)
    with open(DEFAULT_TAG_MAPPING_PATH, "w", encoding="utf-8") as f:
        json.dump(tag_data, f, ensure_ascii=False, indent=2)
    return True

def normalize_text(text):
    if not text:
        return ""
    # Replace hyphens, underscores, slashes, tabs, brackets with space
    t = re.sub(r'[-_/\t\(\)\[\]\,]', ' ', str(text).lower())
    return re.sub(r'\s+', ' ', t).strip()

def is_keyword_match(keyword, func_name):
    kw_norm = normalize_text(keyword)
    func_norm = normalize_text(func_name)
    if not kw_norm or not func_norm:
        return False
    if kw_norm in func_norm or func_norm in kw_norm:
        return True
    # Word subset check: if all significant words of keyword appear in func_name
    kw_words = [w for w in kw_norm.split() if len(w) >= 3 and w not in ["the", "and", "for", "with", "option", "optional"]]
    if kw_words and all(w in func_norm for w in kw_words):
        return True
    return False

def parse_project_excel(excel_source):
    """
    Parses Project_XXXXX.xlsx file.
    Extracts Project Metadata and Column C Function Names.
    """
    if isinstance(excel_source, (str, os.PathLike)):
        wb = openpyxl.load_workbook(excel_source, data_only=True)
    elif hasattr(excel_source, 'read'):
        wb = openpyxl.load_workbook(excel_source, data_only=True)
    else:
        wb = openpyxl.load_workbook(excel_source, data_only=True)

    # Prefer project specific sheet or first sheet with data
    sheet_name = None
    for sn in wb.sheetnames:
        if 'project' in sn.lower():
            sheet_name = sn
            break
    if not sheet_name:
        sheet_name = wb.sheetnames[0]

    ws = wb[sheet_name]

    # Extract Project Metadata from top rows
    meta = {
        "project_no": "",
        "customer": "",
        "machine_type": "",
        "export_date": "",
        "sheet_name": sheet_name
    }

    for r in range(1, min(10, ws.max_row + 1)):
        for c in range(1, min(6, ws.max_column + 1)):
            val = str(ws.cell(r, c).value or '').strip()
            if 'project n' in val.lower() or 'project no' in val.lower():
                next_val = str(ws.cell(r, c + 1).value or '').strip()
                if next_val:
                    meta["project_no"] = next_val
            elif 'customer' in val.lower() or 'custome' in val.lower():
                next_val = str(ws.cell(r, c + 1).value or '').strip()
                if next_val:
                    meta["customer"] = next_val
            elif 'machine type' in val.lower():
                next_val = str(ws.cell(r, c + 1).value or '').strip()
                if next_val:
                    meta["machine_type"] = next_val
            elif 'export d' in val.lower() or 'date' in val.lower():
                next_val = str(ws.cell(r, c + 1).value or '').strip()
                if next_val:
                    meta["export_date"] = next_val

    # Find Header row (usually contains 'Function Name' or 'Main Function')
    header_row_idx = 8
    for r in range(1, min(15, ws.max_row + 1)):
        row_str = " ".join([str(ws.cell(r, c).value or '') for c in range(1, ws.max_column + 1)]).lower()
        if 'function name' in row_str or 'main function' in row_str:
            header_row_idx = r
            break

    functions_list = []
    configured_function_names = set()

    for r in range(header_row_idx + 1, ws.max_row + 1):
        item_val = str(ws.cell(r, 1).value or '').strip()
        main_func = str(ws.cell(r, 2).value or '').strip()
        func_name = str(ws.cell(r, 3).value or '').strip()
        mtype = str(ws.cell(r, 4).value or '').strip()
        note = str(ws.cell(r, 6).value or '').strip()
        status = str(ws.cell(r, 7).value or '').strip()

        if not func_name and not main_func:
            continue

        if func_name:
            configured_function_names.add(func_name.strip().lower())

        functions_list.append({
            "item": item_val,
            "main_function": main_func,
            "function_name": func_name,
            "machine_type": mtype,
            "description": note,
            "status": status if status else "Configured"
        })

    functions_df = pd.DataFrame(functions_list)

    return {
        "metadata": meta,
        "functions_df": functions_df,
        "configured_functions": configured_function_names,
        "total_functions": len(functions_list)
    }

def generate_iq_installation_word(
    excel_source,
    selected_protocols=None,
    word_template_path=None,
    custom_tag_mapping=None,
    **kwargs
):
    """
    Generates IQ Installation Word Protocol:
    1. Removes unselected Test Protocol sections & tables.
    2. Evaluates tagged test points vs configured Excel functions.
    3. Prunes untagged/missing function rows.
    4. Re-numbers test points sequentially (1., 2., 3. ...).
    5. Returns in-memory bytes for direct download.
    """
    if selected_protocols is None or len(selected_protocols) == 0:
        selected_protocols = ["Tube Filler FP 10"]

    parsed_excel = parse_project_excel(excel_source)
    configured_funcs = parsed_excel["configured_functions"]

    tag_mapping = custom_tag_mapping if custom_tag_mapping else load_tag_mapping()
    protocols_cfg = tag_mapping.get("protocols", {})

    template_to_use = word_template_path if (word_template_path and os.path.exists(word_template_path)) else DEFAULT_TEMPLATE_PATH
    if not os.path.exists(template_to_use):
        raise FileNotFoundError(f"Master template file not found: {template_to_use}")

    doc = docx.Document(template_to_use)

    # Protocol mapping to Table in doc
    # Table 3 (index 2): Cartoner
    # Table 4 (index 3): Tube Filler
    # Table 5 (index 4): Tube Filler FP 10
    protocol_info = [
        {"name": "Cartoner", "heading_prefix": "3 Test Protocol (Cartoner)", "table_idx": 2},
        {"name": "Tube Filler", "heading_prefix": "4 Test Protocol (Tube Filler)", "table_idx": 3},
        {"name": "Tube Filler FP 10", "heading_prefix": "5 Test Protocol (Tube Filler IWK FP 10)", "table_idx": 4}
    ]

    breakdown = []
    total_kept = 0
    total_pruned = 0

    # 1. Process each Protocol Table
    for p_info in protocol_info:
        p_name = p_info["name"]
        t_idx = p_info["table_idx"]

        if t_idx >= len(doc.tables):
            continue

        tbl = doc.tables[t_idx]
        p_cfg = protocols_cfg.get(p_name, {})
        p_tags = p_cfg.get("tags", [])

        # Create lookup by test point text
        tags_lookup = {}
        for t in p_tags:
            tags_lookup[t["test_point_text"].strip().lower()] = t

        if p_name not in selected_protocols:
            # Delete unselected protocol table rows
            for row in list(tbl.rows):
                tr = row._tr
                tbl._tbl.remove(tr)
            
            # Find and remove heading paragraph
            for p in doc.paragraphs:
                if p_info["heading_prefix"].lower() in p.text.lower() or f"test protocol ({p_name.lower()})" in p.text.lower():
                    p.text = ""
            continue

        # For Selected Protocol: Filter rows
        rows_to_delete = []
        item_rows = []

        for r_idx, row in enumerate(tbl.rows):
            cells_text = [c.text.strip() for c in row.cells]
            
            # Skip header row (Row 0)
            if r_idx == 0:
                continue

            # Check if row is a Category Header (all cells have same text or cells 1-5 are empty)
            first_cell = cells_text[0]
            second_cell = cells_text[1] if len(cells_text) > 1 else ""
            
            is_cat_header = (all(c == first_cell for c in cells_text) and first_cell != "") or (first_cell != "" and second_cell == "")

            if is_cat_header:
                continue

            # Standard Info Rows (Model, Serial no, Order no, Voltage, Frequency, Control voltage)
            tp_text = second_cell if second_cell else first_cell
            if not tp_text or any(tp_text.startswith(prefix) for prefix in ["Model", "Serial no", "Order no", "Voltage", "Frequency", "Control voltage"]):
                continue

            # It's a test point row!
            clean_tp = tp_text.strip().lower()
            tag_entry = tags_lookup.get(clean_tp)

            # Check matching tag
            if tag_entry:
                is_mand = tag_entry.get("is_mandatory", False)
                excel_keywords = tag_entry.get("excel_keywords", [])
                
                # Check if explicitly set to Hold
                is_hold_tag = any(k.strip().lower() == "hold" for k in excel_keywords)
                if is_hold_tag and not is_mand:
                    total_pruned += 1
                    breakdown.append({
                        "protocol": p_name,
                        "test_point": tp_text,
                        "status": "PRUNED (HOLD)",
                        "reason": "Tagged as 'Hold' (พักไว้สำหรับอนาคต / ลบแถวออก)"
                    })
                    rows_to_delete.append(row)
                    continue

                if is_mand:
                    total_kept += 1
                    breakdown.append({
                        "protocol": p_name,
                        "test_point": tp_text,
                        "status": "KEPT",
                        "reason": "Mandatory Test Point (คงไว้เสมอ)"
                    })
                    item_rows.append(row)
                    continue

                if not excel_keywords:
                    total_pruned += 1
                    breakdown.append({
                        "protocol": p_name,
                        "test_point": tp_text,
                        "status": "PRUNED (UNMAPPED)",
                        "reason": "No functions mapped / ลบแถวออก"
                    })
                    rows_to_delete.append(row)
                    continue

                # Check if any keyword matches any configured function
                is_matched = False
                matched_func_name = ""
                for kw in excel_keywords:
                    kw_clean = kw.strip().lower()
                    if kw_clean == "hold":
                        continue
                    for conf_f in configured_funcs:
                        if is_keyword_match(kw_clean, conf_f):
                            is_matched = True
                            matched_func_name = conf_f
                            break
                    if is_matched:
                        break

                if is_matched:
                    total_kept += 1
                    breakdown.append({
                        "protocol": p_name,
                        "test_point": tp_text,
                        "status": "KEPT",
                        "reason": f"Matched Excel Function: '{matched_func_name}'"
                    })
                    item_rows.append(row)
                else:
                    total_pruned += 1
                    breakdown.append({
                        "protocol": p_name,
                        "test_point": tp_text,
                        "status": "PRUNED (DELETED)",
                        "reason": f"Not found in project functions (Keywords: {', '.join(excel_keywords[:2])})"
                    })
                    rows_to_delete.append(row)
            else:
                # Untagged rows are mandatory/standard items -> ALWAYS KEPT
                total_kept += 1
                breakdown.append({
                    "protocol": p_name,
                    "test_point": tp_text,
                    "status": "KEPT",
                    "reason": "Mandatory / Untagged Standard Test Point"
                })
                item_rows.append(row)

        # Execute row deletions
        for r in rows_to_delete:
            tr = r._tr
            tbl._tbl.remove(tr)

        # Sequential Re-Numbering of Test Points (1., 2., 3. ...)
        counter = 1
        for row in tbl.rows[1:]:
            cells_text = [c.text.strip() for c in row.cells]
            first_cell = cells_text[0]
            second_cell = cells_text[1] if len(cells_text) > 1 else ""
            
            is_cat_header = (all(c == first_cell for c in cells_text) and first_cell != "") or (first_cell != "" and second_cell == "")
            if is_cat_header:
                continue

            tp_text = second_cell if second_cell else first_cell
            if any(tp_text.startswith(prefix) for prefix in ["Model", "Serial no", "Order no", "Voltage", "Frequency", "Control voltage"]):
                continue

            # Set Test Point number
            if len(row.cells) > 0:
                row.cells[0].text = f"{counter}."
                if row.cells[0].paragraphs:
                    row.cells[0].paragraphs[0].alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                counter += 1

    # Save to BytesIO in memory
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    doc_bytes = doc_io.getvalue()

    today_str = datetime.datetime.now().strftime("%Y-%m-%d")
    proto_slug = "_".join(selected_protocols).replace(" ", "_").replace("(", "").replace(")", "")
    project_no = parsed_excel["metadata"].get("project_no", "5XXXX")
    out_filename = f"{project_no}_02_IQ_Installation_{proto_slug}_{today_str}_en.docx"

    return {
        "success": True,
        "file_name": out_filename,
        "doc_bytes": doc_bytes,
        "total_kept": total_kept,
        "total_pruned": total_pruned,
        "selected_protocols": selected_protocols,
        "breakdown": breakdown,
        "breakdown_df": pd.DataFrame(breakdown),
        "excel_metadata": parsed_excel["metadata"]
    }
