def save_all_assignments_to_sheet():
    if not gc:
        st.warning("Δεν είναι δυνατή η αποθήκευση λόγω σφάλματος σύνδεσης API.")
        return
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Assignments")
        rows = [["Project", "Item_ID", "Task_Name", "Assigned_User", "Assigned_Date", "Status_Done", "Task_Type"]]

        # Δημιουργία χάρτη αντιστοίχισης Item_ID -> Project Name
        item_to_project = {}
        if procurement_df is not None and not procurement_df.empty:
            for _, r in procurement_df.iterrows():
                item_to_project[str(r["ID"])] = str(r["Project"])

        for u_key, t_list in st.session_state.get("tasks_store", {}).items():
            item_id = u_key.split("_")[0]
            proj_name = item_to_project.get(item_id, "-")
            for t in t_list:
                if t.get("task") and t.get("task") != "- Επιλογή Εργασίας -":
                    rows.append([
                        proj_name, item_id, t.get("task"), t.get("user"), str(t.get("date")), str(t.get("done")), "ITEM"
                    ])

        for p_key, p_dict in st.session_state.get("project_tasks_store", {}).items():
            proj_name = p_key.replace("proj_", "")
            if isinstance(p_dict, dict):
                for t_name, p_data in p_dict.items():
                    if isinstance(p_data, dict) and p_data.get("active", False):
                        rows.append([
                            proj_name, "-", t_name, p_data.get("user"), str(p_data.get("date")), str(p_data.get("done")), "PROJECT"
                        ])

        sheet.clear()
        sheet.update(range_name="A1", values=rows)
    except Exception as e:
        st.error(f"Σφάλμα κατά την αποθήκευση στο Google Sheet: {e}")
