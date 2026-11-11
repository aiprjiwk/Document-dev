import os
import sys
import io
import json
import pandas as pd
import streamlit as st

def render_fat_protocol_page():
    st.title("📋 FAT Protocol Generator & Master Maintenance")
    st.markdown("Automated **FAT Protocol (.docx)** generation tool from Excel with a central category maintenance system (Master Rules).")
    st.markdown("---")

    # Add FAT Protocal tool directory to sys.path
    app_team_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    fat_tool_dir = os.path.join(app_team_root, "FAT Protocal tool")
    if fat_tool_dir not in sys.path:
        sys.path.insert(0, fat_tool_dir)

    try:
        import fat_generator as fg
    except Exception as e:
        st.error(f"❌ Failed to load fat_generator module: {e}")
        return

    tab1, tab2 = st.tabs(["📄 Import & Generate Protocol", "🛠️ Master Rules Maintenance"])

    # ========================================================
    # TAB 1: IMPORT EXCEL & GENERATE
    # ========================================================
    with tab1:
        st.subheader("📄 Generate FAT Protocol from Excel")
        st.markdown("Upload Excel calculation files to verify categorization, review mapped tables, and generate Word FAT Protocol (.docx) documents.")

        # Project Metadata Card
        with st.container(border=True):
            st.markdown("##### ⚙️ Project Information & Document Header Metadata")
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            with col_m1:
                project = st.text_input("Project:", value="IWK TZF / IWK TZC", key="fat_proj_input")
            with col_m2:
                order_no = st.text_input("Order no.:", value="55990-06767", key="fat_ord_input")
            with col_m3:
                customer = st.text_input("Customer:", value="Customer Name", key="fat_cust_input")
            with col_m4:
                version = st.text_input("Version:", value="1.0", key="fat_ver_input")

            auto_save_new = st.checkbox("💾 Automatically save newly mapped categories to Master Rules on Generate", value=True, key="fat_autosave_chk")

        uploaded_file = st.file_uploader("📂 Select Excel (.xlsx) file to Import", type=["xlsx"], key="fat_excel_uploader")

        # Save to temp directory and parse
        if uploaded_file is not None:
            temp_dir = os.path.join(fat_tool_dir, "temp")
            os.makedirs(temp_dir, exist_ok=True)
            temp_excel_path = os.path.join(temp_dir, uploaded_file.name)
            with open(temp_excel_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            if "fat_last_uploaded" not in st.session_state or st.session_state["fat_last_uploaded"] != uploaded_file.name:
                st.session_state["fat_parsed_data"] = fg.parse_excel_file(temp_excel_path)
                st.session_state["fat_last_uploaded"] = uploaded_file.name

        if "fat_parsed_data" in st.session_state and st.session_state["fat_parsed_data"]:
            parsed = st.session_state["fat_parsed_data"]
            sheets = parsed.get("sheets", [])

            st.success(f"Detected {len(sheets)} Sheet(s): **{', '.join(sheets)}**")

            # Collect Level 2 and Level 3 options dynamically from Master Rules + Parsed Data
            m_rules = fg.load_master_rules()
            m_l2_list = [r.get("level_2", "") for r in m_rules if r.get("level_2")]
            m_l3_list = [r.get("level_3", "") for r in m_rules if r.get("level_3")]
            for s_items in parsed.get("items_by_sheet", {}).values():
                for it in s_items:
                    if it.get("level_2"): m_l2_list.append(it["level_2"])
                    if it.get("level_3"): m_l3_list.append(it["level_3"])

            existing_l2 = sorted(list(set(x.strip() for x in m_l2_list if x.strip() and x.strip() not in ["None", "No"])))
            level_2_options = ["Documents", "Safety", "Machine details", "No", "None"] + [x for x in existing_l2 if x not in ["Documents", "Safety", "Machine details"]]

            existing_l3 = sorted(list(set(x.strip() for x in m_l3_list if x.strip() and x.strip() not in ["None", "No"])))
            all_h3_options = ["None", "No"] + fg.STANDARD_H3_LIST + [x for x in existing_l3 if x not in fg.STANDARD_H3_LIST]

            updated_items_by_sheet = {}

            for sheet_name in sheets:
                items = parsed["items_by_sheet"].get(sheet_name, [])
                l1_title = fg.format_l1_heading(sheet_name)
                st.subheader(f"📑 Sheet: {sheet_name} (Level 1: {l1_title})")

                if not items:
                    st.info("No items found in this sheet.")
                    continue

                # Metric Counters
                skip_count = sum(1 for it in items if str(it.get("level_2", "")).strip().lower() == "no" or str(it.get("level_3", "")).strip().lower() == "no")
                new_count = sum(1 for it in items if it.get("is_new") and str(it.get("level_2", "")).strip().lower() != "no" and str(it.get("level_3", "")).strip().lower() != "no")
                mapped_count = len(items) - new_count - skip_count

                c1, c2, c3, c4 = st.columns([1, 1, 1, 1.5])
                c1.metric("Total Items", len(items))
                c2.metric("Matched Master", mapped_count)
                c3.metric("🚫 Skipped by Rule (No)", skip_count)
                c4.metric("🟡 New Items", new_count)

                # Build table rows for data_editor
                table_rows = []
                for idx, it in enumerate(items):
                    l1_val = it.get("level_1") or l1_title
                    l2_val = it.get("level_2", "") or "None"
                    l3_val = it.get("level_3", "") or "None"
                    is_no = (str(l2_val).strip().lower() == "no" or str(l3_val).strip().lower() == "no")
                    status = "🚫 Skipped (No)" if is_no else ("🟡 New Item (Yellow Highlight)" if it.get("is_new") else "✅ Master")
                    table_rows.append({
                        "Status": status,
                        "PosNo": it.get("pos_no", "") or "None",
                        "Kurztext (Function / Installation)": it.get("kurztext", "") or "None",
                        "Level 1": l1_val,
                        "Level 2": l2_val,
                        "Level 3 / Category": l3_val,
                    })

                df = pd.DataFrame(table_rows)

                if new_count > 0:
                    st.warning(f"⚠️ Found **{new_count}** new item(s) not in Master Database. Please review or select 'Level 2' and 'Level 3 / Category' in the table below. When generated, these items will be **highlighted in yellow** in the Word file for easy identification.")

                edited_df = st.data_editor(
                    df,
                    column_config={
                        "Status": st.column_config.TextColumn("Status", disabled=True),
                        "PosNo": st.column_config.TextColumn("PosNo", width="small"),
                        "Kurztext (Function / Installation)": st.column_config.TextColumn("Kurztext (Function / Installation)", width="large"),
                        "Level 1": st.column_config.TextColumn("Level 1", width="medium"),
                        "Level 2": st.column_config.SelectboxColumn("Level 2", options=level_2_options, width="medium"),
                        "Level 3 / Category": st.column_config.SelectboxColumn("Level 3 / Category", options=all_h3_options, width="large"),
                    },
                    disabled=["Status", "PosNo", "Kurztext (Function / Installation)"],
                    use_container_width=True,
                    key=f"fat_editor_{sheet_name}"
                )

                # Reconstruct updated items
                clean_none = lambda v: "" if str(v).strip().lower() in ["none", "nan", ""] else str(v).strip()
                final_sheet_items = []
                for i, row in edited_df.iterrows():
                    orig_item = items[i]
                    final_sheet_items.append({
                        "pos_no": clean_none(row["PosNo"]),
                        "kurztext": clean_none(row["Kurztext (Function / Installation)"]),
                        "kurztext_de": orig_item.get("kurztext_de", ""),
                        "menge": orig_item.get("menge", "1"),
                        "level_1": clean_none(row["Level 1"]),
                        "level_2": clean_none(row["Level 2"]),
                        "level_3": clean_none(row["Level 3 / Category"]),
                        "is_new": orig_item.get("is_new", False)
                    })
                updated_items_by_sheet[sheet_name] = final_sheet_items

            st.markdown("---")

            col_gen1, col_gen2 = st.columns([1.5, 3])
            with col_gen1:
                gen_btn = st.button("🚀 Generate FAT Protocol (Word)", type="primary", use_container_width=True, key="fat_gen_doc_btn")

            if gen_btn:
                with st.spinner("Generating Word FAT Protocol document..."):
                    # Auto save new rules if checked
                    if auto_save_new:
                        rules = fg.load_master_rules()
                        existing_keys = {r.get("pos_no"): r for r in rules if r.get("pos_no")}
                        for sheet_name, s_items in updated_items_by_sheet.items():
                            for item in s_items:
                                pos = str(item.get("pos_no") or "").strip()
                                if pos and pos in existing_keys:
                                    existing_keys[pos]["level_1"] = item.get("level_1", "")
                                    existing_keys[pos]["level_2"] = item.get("level_2", "")
                                    existing_keys[pos]["level_3"] = item.get("level_3", "")
                                elif pos or item.get("kurztext"):
                                    rules.append({
                                        "pos_no": pos,
                                        "kurztext": item.get("kurztext", ""),
                                        "level_1": item.get("level_1", ""),
                                        "level_2": item.get("level_2", ""),
                                        "level_3": item.get("level_3", "")
                                    })
                        fg.save_master_rules(rules)

                    meta_data = {
                        "project": project,
                        "order_no": order_no,
                        "customer": customer,
                        "version": version
                    }
                    out_dir = os.path.join(fat_tool_dir, "output")
                    os.makedirs(out_dir, exist_ok=True)
                    out_filename = f"{order_no}_FAT_Protocol_v{version}.docx".replace(" ", "_")
                    output_docx_path = os.path.join(out_dir, out_filename)

                    fg.build_fat_document(meta_data, updated_items_by_sheet, output_docx_path)

                    with open(output_docx_path, "rb") as f:
                        doc_bytes = f.read()

                    st.session_state["fat_generated_bytes"] = doc_bytes
                    st.session_state["fat_generated_filename"] = out_filename
                    st.success("✅ FAT Protocol document generated successfully!")

            if "fat_generated_bytes" in st.session_state:
                st.download_button(
                    label=f"📥 Download Word Document: {st.session_state['fat_generated_filename']}",
                    data=st.session_state["fat_generated_bytes"],
                    file_name=st.session_state["fat_generated_filename"],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    key="fat_download_result_btn"
                )

        else:
            st.info("👆 Please upload an Excel (.xlsx) file to start.")

    # ========================================================
    # TAB 2: MASTER RULES MAINTENANCE
    # ========================================================
    with tab2:
        st.subheader("🛠️ Master Rules Maintenance")
        st.markdown("Manage central system rules: Define how PosNo or test item descriptions are assigned to Level 1, Level 2, and Level 3 categories.")

        rules = fg.load_master_rules()

        # ----------------- TEMPLATE MANAGEMENT BAR -----------------
        st.markdown("#### 📄 Word Document Template Management (FAT Protocol Template)")
        active_tpl_path = fg.get_template_path()
        is_custom_tpl = (os.path.abspath(active_tpl_path) != os.path.abspath(fg.DEFAULT_TEMPLATE))
        current_tpl_name = os.path.basename(active_tpl_path)

        col_tp1, col_tp2, col_tp3 = st.columns([2.5, 1, 1])
        with col_tp1:
            tpl_status_badge = "🔵 System Default Template" if not is_custom_tpl else "🟢 Custom Template"
            st.info(f"**Active Template:** `{current_tpl_name}` ({tpl_status_badge})")
        with col_tp2:
            if os.path.exists(active_tpl_path):
                with open(active_tpl_path, "rb") as f_tpl:
                    tpl_bytes = f_tpl.read()
                st.download_button(
                    label="📥 Download Current Template",
                    data=tpl_bytes,
                    file_name=current_tpl_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="fat_dl_tpl_btn"
                )
        with col_tp3:
            if is_custom_tpl:
                if st.button("🗑️ Reset to Default Template", help="Revert back to standard system default template", use_container_width=True, key="fat_reset_tpl_btn"):
                    fg.reset_template_path()
                    st.success("✅ Reset to default template successfully!")
                    st.rerun()

        with st.expander("🔄 Update / Change Word Template (.docx)", expanded=False):
            st.caption("Upload a new Word Template (.docx) version for future document generation")
            new_tpl_upload = st.file_uploader("Select Word Template file (.docx)", type=["docx"], key="fat_upload_new_docx_template")
            if new_tpl_upload is not None:
                if st.button("🚀 Confirm Template Update", type="primary", key="fat_confirm_new_tpl_btn"):
                    templates_dir = os.path.join(fat_tool_dir, "templates")
                    os.makedirs(templates_dir, exist_ok=True)
                    saved_tpl_path = os.path.join(templates_dir, new_tpl_upload.name)
                    with open(saved_tpl_path, "wb") as f_out:
                        f_out.write(new_tpl_upload.getbuffer())
                    fg.set_template_path(saved_tpl_path)
                    st.success(f"✅ Updated template to '{new_tpl_upload.name}' successfully!")
                    st.rerun()

        st.divider()

        # ----------------- IMPORT & EXPORT EXCEL BAR -----------------
        st.markdown("#### 📥 Import & 📤 Export Master Rules via Excel")

        export_rules = fg.load_master_rules()
        df_export = pd.DataFrame(export_rules) if export_rules else pd.DataFrame(columns=["pos_no", "kurztext", "level_1", "level_2", "level_3"])
        df_export = df_export.rename(columns={
            "pos_no": "PosNo",
            "kurztext": "Test Item / Description",
            "level_1": "Level 1",
            "level_2": "Level 2",
            "level_3": "Level 3"
        })
        for c in ["PosNo", "Test Item / Description", "Level 1", "Level 2", "Level 3"]:
            if c in df_export.columns:
                df_export[c] = df_export[c].apply(lambda x: "None" if (pd.isna(x) or not str(x).strip()) else str(x).strip())

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Master_Rules")
        excel_bytes = excel_buffer.getvalue()

        col_card1, col_card2 = st.columns(2)
        with col_card1:
            with st.container(border=True):
                st.markdown("##### 📥 Import Data (Excel)")
                st.caption("Upload an Excel file to batch add or update Master Rules (5 columns)")
                import_master_file = st.file_uploader("Select Excel file (.xlsx)", type=["xlsx"], key="fat_upload_master_excel")
                import_mode = st.radio("Save Mode:", ["Update & Append", "Replace All"], horizontal=True, key="fat_import_mode_radio")

                if import_master_file is not None:
                    if st.button("🚀 Confirm Excel Import to Master Rules", type="primary", use_container_width=True, key="fat_confirm_master_import_btn"):
                        try:
                            imported_items = fg.parse_master_rules_excel(import_master_file)

                            if import_mode == "Replace All":
                                final_rules = imported_items
                            else:
                                current_rules = fg.load_master_rules()
                                existing_map = {r.get("pos_no"): r for r in current_rules if r.get("pos_no")}
                                for it in imported_items:
                                    p = it["pos_no"]
                                    if p and p in existing_map:
                                        if it["kurztext"]: existing_map[p]["kurztext"] = it["kurztext"]
                                        if it["level_1"]: existing_map[p]["level_1"] = it["level_1"]
                                        if it["level_2"]: existing_map[p]["level_2"] = it["level_2"]
                                        if it["level_3"]: existing_map[p]["level_3"] = it["level_3"]
                                    else:
                                        current_rules.append(it)
                                final_rules = current_rules

                            fg.save_master_rules(final_rules)
                            st.success(f"✅ Imported {len(imported_items)} items successfully! Current total: {len(final_rules)} rules.")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Error importing file: {str(e)}")

        with col_card2:
            with st.container(border=True):
                st.markdown("##### 📤 Export Data (Excel)")
                st.caption("Download all Master Rules for backup or external editing in Excel")
                st.write(f"Total rules to export: **{len(df_export)} items**")
                st.download_button(
                    label="📤 Download Excel File (Export Master Rules)",
                    data=excel_bytes,
                    file_name="Master_Rules_Export.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="secondary",
                    key="fat_dl_master_rules_btn"
                )

        st.divider()

        # ----------------- DATA TABLE & ACTION BUTTONS -----------------
        rules = fg.load_master_rules()
        df_rules = pd.DataFrame(rules) if rules else pd.DataFrame(columns=["pos_no", "kurztext", "level_1", "level_2", "level_3"])
        df_rules = df_rules.rename(columns={
            "pos_no": "PosNo",
            "kurztext": "Kurztext / Test Description",
            "level_1": "Level 1",
            "level_2": "Level 2",
            "level_3": "Level 3 / Category"
        })

        st.write(f"Current Master Rules in System: **{len(df_rules)} items** (Enter or select **None** for unused fields)")

        df_rules.insert(0, "🗑️ Delete", False)

        for col in ["PosNo", "Kurztext / Test Description", "Level 1", "Level 2", "Level 3 / Category"]:
            if col in df_rules.columns:
                df_rules[col] = df_rules[col].apply(lambda x: "None" if (pd.isna(x) or str(x).strip() in ["", "None", "nan", "NoneType"]) else str(x).strip())

        existing_l2 = [x for x in df_rules["Level 2"].unique() if x and x not in ["None", "No"]]
        standard_l2 = ["Documents", "Safety", "Machine details"]
        level_2_opts = ["None", "No"] + standard_l2 + [x for x in existing_l2 if x not in standard_l2]

        existing_l3 = [x for x in df_rules["Level 3 / Category"].unique() if x and x not in ["None", "No"]]
        h3_opts = ["None", "No"] + fg.STANDARD_H3_LIST + [x for x in existing_l3 if x not in fg.STANDARD_H3_LIST]

        st.info("💡 **Direct Editing Supported:** Double-click any cell in the table below to edit, or click **➕ Add New Row** to insert a new blank row at the top (Set any level to **No** to skip inserting that item into the document). Click **💾 Save Master Rules** when finished.")

        c_btn1, c_btn2, c_btn3, c_sp = st.columns([1.5, 1.4, 1.8, 2.3])
        with c_btn1:
            save_btn = st.button("💾 Save Master Rules", type="primary", use_container_width=True, key="fat_save_rules_btn")
        with c_btn2:
            add_row_btn = st.button("➕ Add New Row", use_container_width=True, key="fat_add_row_btn")
        with c_btn3:
            del_selected_btn = st.button("🗑️ Delete Selected Rows", use_container_width=True, key="fat_del_selected_btn")

        edited_rules_df = st.data_editor(
            df_rules,
            num_rows="dynamic",
            column_config={
                "🗑️ Delete": st.column_config.CheckboxColumn("🗑️ Delete", help="Check rows to delete, then click 'Delete Selected Rows' or Save", default=False, width="small"),
                "PosNo": st.column_config.TextColumn("PosNo", width="small", help="PosNo code or None if unused"),
                "Kurztext / Test Description": st.column_config.TextColumn("Kurztext / Test Description", width="large", help="Test description text or None if unused"),
                "Level 1": st.column_config.TextColumn("Level 1", width="medium", help="e.g. IWK TZC or None to use sheet name"),
                "Level 2": st.column_config.SelectboxColumn("Level 2", options=level_2_opts, width="medium", help="Select Documents, Safety, Machine details or None"),
                "Level 3 / Category": st.column_config.SelectboxColumn("Level 3 / Category", options=h3_opts, width="large", help="Select sub-category or None"),
            },
            use_container_width=True,
            key="fat_master_rules_data_editor"
        )

        def process_and_save(df_to_save, filter_deleted=True):
            new_rules = []
            is_empty_or_none = lambda v: not v or str(v).strip().lower() in ["none", "nan", ""]
            for _, r in df_to_save.iterrows():
                if filter_deleted and r.get("🗑️ Delete", False):
                    continue
                pos = str(r.get("PosNo") or "").strip()
                kt = str(r.get("Kurztext / Test Description") or "").strip()
                l1 = str(r.get("Level 1") or "").strip()
                l2 = str(r.get("Level 2") or "").strip()
                l3 = str(r.get("Level 3 / Category") or "").strip()

                if is_empty_or_none(pos) and is_empty_or_none(kt) and is_empty_or_none(l1) and is_empty_or_none(l2) and is_empty_or_none(l3):
                    continue

                new_rules.append({
                    "pos_no": "" if is_empty_or_none(pos) else pos,
                    "kurztext": "" if is_empty_or_none(kt) else kt,
                    "level_1": "" if is_empty_or_none(l1) else l1,
                    "level_2": "" if is_empty_or_none(l2) else l2,
                    "level_3": "" if is_empty_or_none(l3) else l3
                })
            fg.save_master_rules(new_rules)
            return len(new_rules)

        if add_row_btn:
            new_rules = []
            is_empty_or_none = lambda v: not v or str(v).strip().lower() in ["none", "nan", ""]
            for _, r in edited_rules_df.iterrows():
                pos = str(r.get("PosNo") or "").strip()
                kt = str(r.get("Kurztext / Test Description") or "").strip()
                l1 = str(r.get("Level 1") or "").strip()
                l2 = str(r.get("Level 2") or "").strip()
                l3 = str(r.get("Level 3 / Category") or "").strip()
                if is_empty_or_none(pos) and is_empty_or_none(kt) and is_empty_or_none(l1) and is_empty_or_none(l2) and is_empty_or_none(l3):
                    continue
                new_rules.append({
                    "pos_no": "" if is_empty_or_none(pos) else pos,
                    "kurztext": "" if is_empty_or_none(kt) else kt,
                    "level_1": "" if is_empty_or_none(l1) else l1,
                    "level_2": "" if is_empty_or_none(l2) else l2,
                    "level_3": "" if is_empty_or_none(l3) else l3
                })
            new_rules.insert(0, {
                "pos_no": "",
                "kurztext": "",
                "level_1": "",
                "level_2": "",
                "level_3": ""
            })
            fg.save_master_rules(new_rules)
            st.success("✅ Added a new blank row at the top! You can enter data and click save.")
            st.rerun()

        if del_selected_btn:
            selected_count = sum(1 for _, r in edited_rules_df.iterrows() if r.get("🗑️ Delete", False))
            if selected_count == 0:
                st.warning("⚠️ Please check rows to delete using the '🗑️ Delete' checkbox first.")
            else:
                remaining = process_and_save(edited_rules_df, filter_deleted=True)
                st.success(f"🗑️ Deleted {selected_count} row(s) successfully! ({remaining} rules remaining)")
                st.rerun()

        if save_btn:
            saved_count = process_and_save(edited_rules_df, filter_deleted=True)
            st.success(f"✅ Saved Master Rules successfully ({saved_count} items)")
            st.rerun()
