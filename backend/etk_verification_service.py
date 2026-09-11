import io
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

PROJECT_SHEET_NAMES = [
    "CABLIblue", "TZC", "TZ201", "TZ204", "TZ104", "TZ202", "TZ101", "TZ102",
    "FP series", "SC5", "FP8", "TZF-TZ", "TZF", "TZ", "HC5", "CPC", "APC",
    "VI5", "VC5", "VIX", "VI10", "VI-VC", "SC4", "CH4", "SI6"
]

def is_valid_group_code(s) -> bool:
    """
    Validates Group Code according to VBA rules:
    - Excludes 800xxxxx (starts with 800 + 5 digits)
    - True if 7 numeric digits
    - True if >7 digits, first 7 are numeric, and suffix is alphanumeric
    """
    if s is None:
        return False
    val_str = str(s).strip()
    if not val_str:
        return False
        
    # Exclude 800xxxxx (e.g. 80055843)
    if len(val_str) == 8 and val_str.startswith("800") and val_str[3:].isdigit():
        return False
        
    # Numeric 7 digits
    if len(val_str) == 7 and val_str.isdigit():
        return True
        
    # 7 digits + alphanumeric suffix (e.g. A02 / AXX)
    if len(val_str) > 7 and val_str[:7].isdigit():
        suf = val_str[7:]
        return suf.isalnum()
        
    return False

def extract_project_plan(file_bytes):
    """
    Extracts Project Plan groups and function groups from File #1.
    Scans for matching sheet name in PROJECT_SHEET_NAMES.
    Reads from row 13 onwards (Col B for Group, Col E for Function groups).
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    target_ws = None
    
    # Search for sheet matching names (case-insensitive)
    sheet_name_map = {name.lower(): name for name in wb.sheetnames}
    for expected_name in PROJECT_SHEET_NAMES:
        if expected_name.lower() in sheet_name_map:
            target_ws = wb[sheet_name_map[expected_name.lower()]]
            break
            
    if target_ws is None:
        return False, f"Project sheet not found. Expected one of: {', '.join(PROJECT_SHEET_NAMES)}", None, None

    project_rows = []
    dict_proj = {}
    
    for r in range(13, target_ws.max_row + 1):
        grp = str(target_ws.cell(row=r, column=2).value or "").strip()
        fn = str(target_ws.cell(row=r, column=5).value or "").strip()
        
        if grp.endswith(".0"):
            grp = grp[:-2]
            
        if is_valid_group_code(grp):
            if fn and not fn.startswith("***"):
                project_rows.append({"Group": grp, "Function groups": fn})
                if grp not in dict_proj:
                    dict_proj[grp] = fn

    return True, "Success", project_rows, dict_proj

def extract_csp2_bom(file_bytes):
    """
    Extracts CSP2 BOM groups from File #2 (Sheet 1).
    Reads Explosion Level (Col A) and Component Number / Group (Col C).
    Applies filtering: Col A ends with 1 or 2, Col C has no '*'.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    
    csp2_rows = []
    dict_csp2 = {}
    dict_csp2_f = {}
    
    for r in range(2, ws.max_row + 1):
        a_val = str(ws.cell(row=r, column=1).value or "").strip()
        d_val = str(ws.cell(row=r, column=3).value or "").strip()
        
        if d_val.endswith(".0"):
            d_val = d_val[:-2]
            
        reason = ""
        if not a_val:
            reason = "Excluded from CSP2 (Explosion level blank)"
        else:
            last_ch = a_val[-1]
            if last_ch not in ["1", "2"]:
                reason = "Excluded from CSP2 (Explosion level not 1 or 2)"
                
        if not reason:
            if "*" in d_val:
                reason = 'Excluded from CSP2 (Column C contains "*")'
                
        if reason:
            if is_valid_group_code(d_val):
                dict_csp2_f[d_val] = reason
        else:
            if is_valid_group_code(d_val):
                csp2_rows.append({"Group": d_val})
                dict_csp2[d_val] = True

    return csp2_rows, dict_csp2, dict_csp2_f

def extract_cspb_bom(file_bytes):
    """
    Extracts CSPB BOM groups from File #2 (Sheet 1).
    Reads Usage Flag (Col A) and Group (Col B).
    Applies filtering: Col A != 'F'.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active
    
    cspb_rows = []
    dict_cspb = {}
    dict_cspb_f = {}
    
    for r in range(2, ws.max_row + 1):
        usage_flag = str(ws.cell(row=r, column=1).value or "").strip().upper()
        grp = str(ws.cell(row=r, column=2).value or "").strip()
        
        if grp.endswith(".0"):
            grp = grp[:-2]
            
        if usage_flag == "F":
            if is_valid_group_code(grp):
                dict_cspb_f[grp] = "Excluded from CSPB (Column A = F)"
        else:
            if is_valid_group_code(grp):
                cspb_rows.append({"Group": grp})
                dict_cspb[grp] = True

    return cspb_rows, dict_cspb, dict_cspb_f

def compare_etk_verification(project_file_bytes, bom_file_bytes, mode="CSP2 BOM"):
    """
    Executes comparison between PROJECT Plan File and CSP2/CSPB BOM File.
    Generates summary statistics, UI comparison dataframe, and 3-Sheet Excel report.
    """
    # 1. Extract Project Plan
    success, msg, project_rows, dict_proj = extract_project_plan(project_file_bytes)
    if not success:
        return False, msg, None, None
        
    # 2. Extract BOM based on mode
    bom_sheet_name = "CSP2" if mode == "CSP2 BOM" else "CSPB_Extract"
    merge_sheet_name = "ProjectPlan&CSP2" if mode == "CSP2 BOM" else "Project&CSPB"
    bom_label = "CSP2 BOM" if mode == "CSP2 BOM" else "CSPB BOM "
    
    if mode == "CSP2 BOM":
        bom_rows, dict_bom, dict_bom_f = extract_csp2_bom(bom_file_bytes)
    else:
        bom_rows, dict_bom, dict_bom_f = extract_cspb_bom(bom_file_bytes)
        
    # 3. Create comparison records
    matched_count = 0
    merged_rows = []
    
    for item in project_rows:
        grp = item["Group"]
        fn = item["Function groups"]
        
        is_matched = grp in dict_bom
        note = ""
        if is_matched:
            matched_count += 1
            bom_grp_val = grp
        else:
            bom_grp_val = ""
            if grp in dict_bom_f:
                note = dict_bom_f[grp]
                
        merged_rows.append({
            "BOM Group": bom_grp_val,
            "Project Plan": grp,
            "Function groups": fn,
            "Status": "Matched" if is_matched else "Unmatched",
            "Note": note
        })

    # 4. DIFF: Project only
    diff_project_only = []
    for grp, fn in dict_proj.items():
        if grp not in dict_bom:
            note = dict_bom_f.get(grp, f"Not found in {mode.split()[0]} (after filter)")
            diff_project_only.append({
                "Group": grp,
                "Function groups": fn,
                "Note": note
            })

    # 5. DIFF: BOM only
    diff_bom_only = []
    for grp in dict_bom.keys():
        if grp not in dict_proj:
            diff_bom_only.append({
                "Group": grp,
                "Note": "Not found in Project"
            })

    summary_stats = {
        "Mode": mode,
        "Project Rows": len(project_rows),
        "BOM Rows": len(bom_rows),
        "Matched Rows": matched_count,
        "Diff Project Only": len(diff_project_only),
        "Diff BOM Only": len(diff_bom_only)
    }

    # Build multi-sheet Excel output using openpyxl
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet
    
    bold_font = Font(bold=True)
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # RGB(198, 239, 206)
    red_fill = PatternFill(start_color="FFC9CE", end_color="FFC9CE", fill_type="solid") # RGB(255, 199, 206)
    
    green_font = Font(color="006100")
    red_font = Font(color="9C0006")
    
    # Sheet 1: Project Plan / From_Project
    sheet1_name = "Project plan" if mode == "CSP2 BOM" else "From_Project"
    ws1 = wb.create_sheet(title=sheet1_name)
    ws1.append(["Group", "Function groups"])
    ws1.row_dimensions[1].font = bold_font
    for item in project_rows:
        ws1.append([item["Group"], item["Function groups"]])
        ws1.cell(row=ws1.max_row, column=1).number_format = "@"

    # Sheet 2: CSP2 / CSPB_Extract
    ws2 = wb.create_sheet(title=bom_sheet_name)
    header_ws2 = "Group (from CSP2 BOM)" if mode == "CSP2 BOM" else "Group"
    ws2.append([header_ws2])
    ws2.row_dimensions[1].font = bold_font
    for item in bom_rows:
        ws2.append([item["Group"]])
        ws2.cell(row=ws2.max_row, column=1).number_format = "@"

    # Sheet 3: ProjectPlan&CSP2 / Project&CSPB
    ws3 = wb.create_sheet(title=merge_sheet_name)
    headers_ws3 = [
        bom_label, "Project Plan", "Function groups",
        "DIFF: Project only", "Function groups",
        f"DIFF: {mode.split()[0]} only", "Note"
    ]
    ws3.append(headers_ws3)
    for col_idx in range(1, 8):
        ws3.cell(row=1, column=col_idx).font = bold_font

    # Populate Project Plan rows into A:C
    for idx, item in enumerate(project_rows, start=2):
        grp = item["Group"]
        fn = item["Function groups"]
        
        ws3.cell(row=idx, column=2, value=grp).number_format = "@"
        ws3.cell(row=idx, column=3, value=fn)
        
        if grp in dict_bom:
            ws3.cell(row=idx, column=1, value=grp).number_format = "@"
            for col in range(1, 4):
                ws3.cell(row=idx, column=col).fill = green_fill
                ws3.cell(row=idx, column=col).font = green_font
        else:
            for col in range(2, 4):
                ws3.cell(row=idx, column=col).fill = red_fill
                ws3.cell(row=idx, column=col).font = red_font
            if grp in dict_bom_f:
                ws3.cell(row=idx, column=7, value=dict_bom_f[grp])

    # Populate DIFF: Project only into D:E and G
    for idx, diff_item in enumerate(diff_project_only, start=2):
        ws3.cell(row=idx, column=4, value=diff_item["Group"]).number_format = "@"
        ws3.cell(row=idx, column=5, value=diff_item["Function groups"])
        if not ws3.cell(row=idx, column=7).value:
            ws3.cell(row=idx, column=7, value=diff_item["Note"])

    # Populate DIFF: BOM only into F and G
    for idx, diff_item in enumerate(diff_bom_only, start=2):
        ws3.cell(row=idx, column=6, value=diff_item["Group"]).number_format = "@"
        if not ws3.cell(row=idx, column=7).value:
            ws3.cell(row=idx, column=7, value=diff_item["Note"])

    # Auto-fit columns
    for ws in [ws1, ws2, ws3]:
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    output = io.BytesIO()
    wb.save(output)
    excel_bytes = output.getvalue()
    
    # Prepare preview dataframe for Streamlit
    preview_df = pd.DataFrame(merged_rows)

    return True, summary_stats, preview_df, excel_bytes
