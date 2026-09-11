import os
import io
import re
import datetime
import tempfile
import pandas as pd
import pythoncom
import win32com.client

DEFAULT_TEMPLATE_PATH = os.path.abspath(r"C:\Users\cpreephim\Desktop\App Team\IQOQDQ\IQ_Format\5XXXX_08_IQ_Format Parts_20XX-XX-XX_en.doc")
FALLBACK_TEMPLATE_PATH = os.path.abspath(r"C:\Users\cpreephim\Desktop\App Team\IQOQDQ\Data IQOQDQ\5XXXX_08_IQ_Format Parts_20XX-XX-XX_en.doc")

def clean_quantity(raw_qty):
    """
    Cleans quantity value and removes decimal part for whole numbers (e.g. 16.0 -> 16, 1.0 -> 1).
    """
    if pd.isna(raw_qty) or raw_qty is None:
        return '1'
    try:
        float_val = float(raw_qty)
        if float_val.is_integer():
            return str(int(float_val))
        return f"{float_val:g}"
    except (ValueError, TypeError):
        s = str(raw_qty).strip()
        if s.endswith('.0'):
            s = s[:-2]
        return s if s and s != 'nan' else '1'

def parse_and_clean_format_excel(excel_source):
    """
    Parses Excel BOM and applies the 4 required rules:
    1. Exclude non-format columns (E-Q).
    2. Exclude Size/dimensions (Col C) and Component number (Col A).
    3. Skip header row.
    4. Exclude rows where Format Code (Col D / Funktion Zeile1) is blank.
    5. Clean Quantity (remove decimal places like 16.0 -> 16).
    """
    if isinstance(excel_source, (str, os.PathLike)):
        df = pd.read_excel(excel_source, header=None)
    else:
        df = pd.read_excel(excel_source, header=None)

    data_rows = df.iloc[1:].copy()
    
    # Filter rows where col 3 (Funktion Zeile1 / Used for format) is valid
    valid_rows = data_rows[
        data_rows[3].notna() & 
        (data_rows[3].astype(str).str.strip() != '') & 
        (data_rows[3].astype(str).str.strip() != 'nan')
    ].copy()

    groups = {}
    flat_records = []
    
    item_counter = 1
    for idx, r in valid_rows.iterrows():
        desc = str(r[1]).strip() if pd.notna(r[1]) else 'Unknown'
        fmt = str(r[3]).strip() if pd.notna(r[3]) else ''
        raw_qty = r[17] if len(r) > 17 else '1'
        qty = clean_quantity(raw_qty)
            
        if desc not in groups:
            groups[desc] = []
        groups[desc].append((desc, fmt, qty))
        
        flat_records.append({
            'Item No.': f'{item_counter}.',
            'Category / Group': desc,
            'Designation': desc,
            'Used for Format': fmt,
            'Quantity (PC)': qty
        })
        item_counter += 1

    preview_df = pd.DataFrame(flat_records)
    return groups, preview_df

def generate_iq_format_word(
    excel_source,
    word_template_path=None,
    word_template_source=None,
    **kwargs
):
    """
    Directly opens the Master Word template (.doc) from IQ_Format directory,
    modifies Table 4 only, updates Quantity (Set) -> Quantity (PC), formats columns to exact widths,
    formats quantities as integers without decimals, preserves all original template metadata,
    and generates the file in memory for direct download.
    """
    groups, preview_df = parse_and_clean_format_excel(excel_source)
    if not groups:
        raise ValueError('No valid format parts found in Excel file. Please ensure Funktion Zeile1 (Col D) contains format data.')

    chosen_tmpl = word_template_path or word_template_source or kwargs.get('word_template_path') or kwargs.get('word_template_source')
    if chosen_tmpl and os.path.exists(os.path.abspath(str(chosen_tmpl))):
        template_to_use = os.path.abspath(str(chosen_tmpl))
    elif os.path.exists(DEFAULT_TEMPLATE_PATH):
        template_to_use = DEFAULT_TEMPLATE_PATH
    elif os.path.exists(FALLBACK_TEMPLATE_PATH):
        template_to_use = FALLBACK_TEMPLATE_PATH
    else:
        raise FileNotFoundError(f'Word master template not found in IQ_Format or Data IQOQDQ.')

    today_str = datetime.datetime.now().strftime('%Y-%m-%d')
    tmpl_base = os.path.splitext(os.path.basename(template_to_use))[0]
    if '20xx-xx-xx' in tmpl_base.lower():
        download_file_name = re.sub(r'20XX-XX-XX', today_str, tmpl_base, flags=re.IGNORECASE) + '.doc'
    else:
        download_file_name = f'{tmpl_base}_{today_str}.doc'

    # Use temporary file to hold the COM output before loading into memory
    temp_dir = tempfile.gettempdir()
    temp_out_path = os.path.join(temp_dir, f"temp_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S_%f')}.doc")

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        doc = word.Documents.Open(
            FileName=template_to_use,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
            Visible=False
        )

        if doc.Tables.Count < 4:
            raise ValueError(f'Table 4 not found in template. Found {doc.Tables.Count} tables.')

        tbl = doc.Tables(4)

        # Update Quantity header from (Set) to (PC)
        for i in range(1, min(tbl.Range.Cells.Count + 1, 9)):
            try:
                c = tbl.Range.Cells(i)
                if 'set' in c.Range.Text.lower():
                    c.Range.Text = "Quantity\r(PC)"
                    c.Range.Font.Name = "Arial"
                    c.Range.Font.Size = 9
                    c.Range.Font.Bold = 1
                    c.Range.ParagraphFormat.Alignment = 1 # Center
            except Exception:
                pass

        # Clear existing sample rows from row 3 onwards
        if tbl.Rows.Count >= 3 or tbl.Range.Cells.Count > 8:
            try:
                start_pos = tbl.Cell(3, 1).Range.Start
                end_pos = tbl.Range.Cells(tbl.Range.Cells.Count).Range.End
                clear_rng = doc.Range(start_pos, end_pos)
                clear_rng.Select()
                word.Selection.Rows.Delete()
            except Exception:
                pass

        # Exact column widths matching original template Table 4
        TEMPLATE_WIDTHS = [36.0, 126.0, 72.0, 49.5, 54.0, 58.5, 57.6]

        item_counter = 1
        for grp_name, items in groups.items():
            # 1. Section Header Row (Left-aligned as in template)
            last_count = tbl.Range.Cells.Count
            tbl.Rows.Add()
            h_start = last_count + 1
            h_end = tbl.Range.Cells.Count
            if (h_end - h_start + 1) == 7:
                tbl.Range.Cells(h_start).Merge(tbl.Range.Cells(h_end))
                
            h_cell = tbl.Range.Cells(h_start)
            h_cell.Range.Text = grp_name
            h_cell.Range.Font.Bold = 1
            h_cell.Range.Font.Name = 'Arial'
            h_cell.Range.Font.Size = 9.5
            h_cell.Range.ParagraphFormat.Alignment = 0 # Left aligned

            # 2. Detail Rows
            is_first_detail = True
            for item in items:
                last_count = tbl.Range.Cells.Count
                tbl.Rows.Add()
                if is_first_detail:
                    try:
                        tbl.Range.Cells(last_count + 1).Split(1, 7)
                    except Exception:
                        pass
                    is_first_detail = False
                    
                start_idx = last_count + 1
                tbl.Range.Cells(start_idx + 0).Range.Text = f'{item_counter}.'
                tbl.Range.Cells(start_idx + 1).Range.Text = item[0]
                tbl.Range.Cells(start_idx + 2).Range.Text = item[1]
                tbl.Range.Cells(start_idx + 3).Range.Text = item[2]
                tbl.Range.Cells(start_idx + 4).Range.Text = ''
                tbl.Range.Cells(start_idx + 5).Range.Text = ''
                tbl.Range.Cells(start_idx + 6).Range.Text = ''
                
                # Alignments:
                tbl.Range.Cells(start_idx + 0).Range.ParagraphFormat.Alignment = 1 # Center
                tbl.Range.Cells(start_idx + 1).Range.ParagraphFormat.Alignment = 0 # Left
                tbl.Range.Cells(start_idx + 2).Range.ParagraphFormat.Alignment = 1 # Center
                tbl.Range.Cells(start_idx + 3).Range.ParagraphFormat.Alignment = 1 # Center
                
                # Set exact column widths matching master template
                for k in range(7):
                    c_cell = tbl.Range.Cells(start_idx + k)
                    c_cell.Width = TEMPLATE_WIDTHS[k]
                    c_rng = c_cell.Range
                    c_rng.Font.Name = 'Arial'
                    c_rng.Font.Size = 9
                    c_rng.Font.Bold = 0
                    
                item_counter += 1

        doc.SaveAs(temp_out_path)
        doc.Close(False)

    finally:
        try:
            word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    # Read binary bytes into memory for direct browser download
    with open(temp_out_path, 'rb') as f:
        doc_bytes = f.read()

    # Clean up temp file immediately so nothing is stored on disk
    try:
        if os.path.exists(temp_out_path):
            os.remove(temp_out_path)
    except Exception:
        pass

    return {
        'success': True,
        'total_groups': len(groups),
        'total_items': len(preview_df),
        'file_name': download_file_name,
        'doc_bytes': doc_bytes,
        'preview_df': preview_df,
        'groups': groups
    }
