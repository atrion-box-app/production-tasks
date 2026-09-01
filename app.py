import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="Production Tasks App", layout="wide")

st.title("🏭 Σύστημα Διαχείρισης Παραγωγής & Tasks")

# Google Sheet URLs
PROC_SHEET_ID = "1QhTd58vuulaC_73sgbjuwG5MVxT6c1c_-MbhypGx0fA"
PROC_GID = "1639392743"
PROC_CSV_URL = f"https://docs.google.com/spreadsheets/d/{PROC_SHEET_ID}/export?format=csv&gid={PROC_GID}"

MY_SHEET_ID = "1rps5ha4wyo8DQ3zwUTqS5BSNMrJPatvqdh8M0iMHVEg"
TIMES_GID = "2126316973"
TEAM_GID = "1303086311"

TIMES_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={TIMES_GID}"
TEAM_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={TEAM_GID}"

WEEKDAYS_GREEK = {
    0: "Δευτέρα",
    1: "Τρίτη",
    2: "Τετάρτη",
    3: "Πέμπτη",
    4: "Παρασκευή",
    5: "Σάββατο",
    6: "Κυριακή"
}

@st.cache_data(ttl=10)
def load_all_data():
    # 1. Φόρτωση Procurement
    try:
        df_proc_raw = pd.read_csv(PROC_CSV_URL, header=None)
        indices = [18, 3, 17, 1, 2, 5, 6, 4, 10, 12]
        df_proc = df_proc_raw.iloc[1:, indices].copy()
        df_proc.columns = [
            "ID", "Ημερομηνία Παράδοσης", "Είδος Δώρου", "Project", 
            "Ποσότητα", "Προμηθευτής", "Υλικό / Προϊόν", 
            "Αναμενόμενη Ημ. Παραλαβής", "Αναμενόμενη Ποσότητα Παραλαβής", "Status Procurement"
        ]
        df_proc = df_proc.fillna("-")
    except Exception:
        df_proc = pd.DataFrame()

    # 2. Φόρτωση Πρότυπων Χρόνων
    tasks_dict = {}
    try:
        df_times_raw = pd.read_csv(TIMES_CSV_URL, header=None)
        for col_idx in range(len(df_times_raw.columns) - 1):
            for row_idx in range(len(df_times_raw)):
                task_name = str(df_times_raw.iloc[row_idx, col_idx]).strip()
                time_val_raw = df_times_raw.iloc[row_idx, col_idx + 1]
                
                if task_name and task_name.lower() not in ["nan", "none", "τύπος εργασίας / υλικό"] and not task_name.startswith("TASK"):
                    try:
                        time_val = float(str(time_val_raw).replace(',', '.'))
                        if time_val >= 0:
                            tasks_dict[task_name] = time_val
                    except ValueError:
                        pass
    except Exception:
        tasks_dict = {"Έλεγχος (εύκολο)": 1.0, "Συναρμολόγηση": 2.0, "Συσκευασία": 1.5}

    # 3. Φόρτωση Ομάδας & Διαθεσιμότητας
    team_members = ["Βαγγέλης Μ.", "Βαγγέλης JR.", "Εποχικός 1", "Εποχικός 2", "Ana", "Alex"]
    availability_dict = {day: {m: 6.0 for m in team_members} for day in WEEKDAYS_GREEK.values()}

    try:
        df_team_raw = pd.read_csv(TEAM_CSV_URL)
        df_team_raw.columns = [str(c).strip() for c in df_team_raw.columns]
        
        ignore_cols = ["Ημέρα", "Σύνολο διαθέσιμων ωρών", "Unnamed: 0"]
        found_members = [c for c in df_team_raw.columns if c and c not in ignore_cols and "Unnamed" not in c]
        
        if found_members:
            team_members = found_members
            
        for _, row in df_team_raw.iterrows():
            day_name = str(row.iloc[0]).strip()
            if day_name in WEEKDAYS_GREEK.values():
                if day_name not in availability_dict:
                    availability_dict[day_name] = {}
                for member in team_members:
                    if member in df_team_raw.columns:
                        try:
                            val = float(str(row[member]).replace(',', '.'))
                            availability_dict[day_name][member] = val
                        except:
                            availability_dict[day_name][member] = 6.0
    except Exception:
        pass

    return df_proc, tasks_dict, team_members, availability_dict

procurement_df, tasks_database, team_database, availability_database = load_all_data()

# Κεντρική Αποθήκη Δεδομένων στη Μνήμη (Session State)
if "tasks_store" not in st.session_state:
    st.session_state["tasks_store"] = {}

if "project_tasks_store" not in st.session_state:
    st.session_state["project_tasks_store"] = {}

FIXED_PROJECT_TASKS = [
    "Σύνθεση (κουτί)",
    "Σύνθεση (πουγκί / τσάντα)",
    "Σύνθεση (χειροποίητο)",
    "Φωτογράφιση",
    "Τοποθέτηση σε χαρτοκιβώτια"
]

option = st.sidebar.selectbox("Επιλογή Οθόνης", ["Καρτέλα Project", "Ημερήσιο Πλάνο Παραγωγής", "Πρότυπα Χρόνων & Ομάδα"])

if option == "Καρτέλα Project":
    st.header("📋 Διαχείριση Παραγωγής & Αναθέσεις ανά Project")
    
    if not procurement_df.empty:
        projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
        selected_project = st.selectbox("Επιλέξτε Project:", projects_list)
        
        filtered_df = procurement_df[procurement_df["Project"] == selected_project].copy()
        
        project_main_qty = 1
        for _, r in filtered_df.iterrows():
            if str(r["Ποσότητα"]).isdigit():
                project_main_qty = max(project_main_qty, int(r["Ποσότητα"]))
                
        total_project_hours = 0.0
        completed_project_hours = 0.0
        total_tasks_count = 0
        completed_tasks_count = 0

        task_options = ["- Επιλογή Εργασίας -"] + sorted(list(tasks_database.keys()))
        team_options = ["- Χωρίς Ανάθεση -"] + team_database

        # --- 1. ΕΡΓΑΣΙΕΣ ΑΝΑ ΥΛΙΚΟ (ITEM LEVEL TASKS) ---
        st.subheader(f"📦 Εργασίες ανά Υλικό ({len(filtered_df)} Υλικά)")

        for idx, row in filtered_df.iterrows():
            item_id = str(row["ID"])
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            status = row["Status Procurement"]
            
            if item_id not in st.session_state["tasks_store"]:
                st.session_state["tasks_store"][item_id] = [
                    {"done": False, "task": "- Επιλογή Εργασίας -", "user": "- Χωρίς Ανάθεση -", "date": date.today()}
                ]
            
            item_tasks = st.session_state["tasks_store"][item_id]
            card_title = f"🆔 {item_id} | {material} — (Ποσότητα: {qty} τμχ) | Status: {status}"
            
            with st.expander(card_title):
                col_info, col_tasks = st.columns([1, 2.5])
                
                with col_info:
                    st.markdown("**ℹ️ Στοιχεία Procurement**")
                    st.write(f"• **Είδος Δώρου:** {row['Είδος Δώρου']}")
                    st.write(f"• **Προμηθευτής:** {row['Προμηθευτής']}")
                    st.write(f"• **Ημ. Παράδοσης:** {row['Ημερομηνία Παράδοσης']}")
                    st.write(f"• **Αναμ. Παραλαβή:** {row['Αναμενόμενη Ημ. Παραλαβής']}")
                    st.divider()

                with col_tasks:
                    st.markdown("**⚙️ Εργασίες Προετοιμασίας Υλικού**")
                    
                    if len(item_tasks) == 0:
                        st.info("Δεν έχουν οριστεί εργασίες για αυτό το υλικό.")

                    for t_idx, t_data in enumerate(list(item_tasks)):
                        c_check, c_task, c_user, c_date, c_time, c_del = st.columns([0.08, 0.32, 0.23, 0.18, 0.11, 0.08])
                        
                        is_done = c_check.checkbox("", value=t_data["done"], key=f"chk_{item_id}_{t_idx}")
                        
                        task_idx = task_options.index(t_data["task"]) if t_data["task"] in task_options else 0
                        selected_task = c_task.selectbox(
                            "Εργασία", task_options, index=task_idx, 
                            key=f"task_{item_id}_{t_idx}", label_visibility="collapsed"
                        )
                        
                        user_idx = team_options.index(t_data["user"]) if t_data["user"] in team_options else 0
                        assigned_user = c_user.selectbox(
                            "Ανάθεση", team_options, index=user_idx, 
                            key=f"user_{item_id}_{t_idx}", label_visibility="collapsed"
                        )
                        
                        assign_date = c_date.date_input(
                            "Ημερομηνία", value=t_data["date"], format="DD/MM/YYYY", 
                            key=f"date_{item_id}_{t_idx}", label_visibility="collapsed"
                        )

                        st.session_state["tasks_store"][item_id][t_idx] = {
                            "done": is_done,
                            "task": selected_task,
                            "user": assigned_user,
                            "date": assign_date
                        }
                        
                        auto_time = tasks_database.get(selected_task, 0.0)
                        
                        if selected_task != "- Επιλογή Εργασίας -":
                            task_hours = (auto_time * qty) / 60
                            total_project_hours += task_hours
                            total_tasks_count += 1
                            
                            if is_done:
                                completed_project_hours += task_hours
                                completed_tasks_count += 1
                                c_time.markdown("✅ **Done**")
                            else:
                                c_time.metric("λ/τμχ", f"{auto_time}λ")
                        else:
                            c_time.caption("0.0λ")
                        
                        if c_del.button("🗑️", key=f"del_{item_id}_{t_idx}"):
                            st.session_state["tasks_store"][item_id].pop(t_idx)
                            st.rerun()

                    col_btn1, col_btn2, _ = st.columns([0.35, 0.35, 0.3])
                    if col_btn1.button("➕ Προσθήκη Εργασίας Υλικού", key=f"add_btn_{item_id}"):
                        st.session_state["tasks_store"][item_id].append(
                            {"done": False, "task": "- Επιλογή Εργασίας -", "user": "- Χωρίς Ανάθεση -", "date": date.today()}
                        )
                        st.rerun()
                    
                    if len(item_tasks) > 0:
                        if col_btn2.button("➖ Αφαίρεση Εργασίας", key=f"rem_btn_{item_id}"):
                            st.session_state["tasks_store"][item_id].pop()
                            st.rerun()

        st.divider()

        # --- 2. ΓΕΝΙΚΕΣ ΕΡΓΑΣΙΕΣ PROJECT (ΣΤΟ ΚΑΤΩ ΜΕΡΟΣ) ---
        st.markdown("### 🛠️ Γενικές Εργασίες Project (Σύνθεση, Συσκευασία & Box)")
        
        proj_key = f"proj_{selected_project}"
        
        if proj_key not in st.session_state["project_tasks_store"] or not isinstance(st.session_state["project_tasks_store"][proj_key], dict):
            st.session_state["project_tasks_store"][proj_key] = {
                t_name: {"active": False, "done": False, "user": "- Χωρίς Ανάθεση -", "date": date.today()}
                for t_name in FIXED_PROJECT_TASKS
            }
        
        proj_tasks_dict = st.session_state["project_tasks_store"][proj_key]
        
        with st.expander(f"📦 5 Σταθερές Γενικές Εργασίες για το Project: {selected_project}", expanded=True):
            st.caption("Ενεργοποιήστε [✓] τις εργασίες που απαιτούνται για το συγκεκριμένο project, αναθέστε σε άτομο και ορίστε ημερομηνία.")
            
            for task_name in FIXED_PROJECT_TASKS:
                t_data = proj_tasks_dict.get(task_name, {"active": False, "done": False, "user": "- Χωρίς Ανάθεση -", "date": date.today()})
                
                c_active, c_name, c_done, c_user, c_date, c_time = st.columns([0.06, 0.30, 0.10, 0.22, 0.18, 0.14])
                
                is_active = c_active.checkbox("", value=t_data["active"], key=f"pact_{proj_key}_{task_name}")
                c_name.markdown(f"**{task_name}**" if is_active else f"<span style='color:gray;'>{task_name}</span>", unsafe_allow_html=True)
                
                is_done = False
                if is_active:
                    is_done = c_done.checkbox("Done", value=t_data["done"], key=f"pdone_{proj_key}_{task_name}")
                else:
                    c_done.caption("—")
                    
                user_idx = team_options.index(t_data["user"]) if t_data["user"] in team_options else 0
                assigned_user = c_user.selectbox(
                    "Ανάθεση", team_options, index=user_idx, 
                    key=f"puser_{proj_key}_{task_name}", label_visibility="collapsed", disabled=not is_active
                )
                
                assign_date = c_date.date_input(
                    "Ημερομηνία", value=t_data["date"], format="DD/MM/YYYY", 
                    key=f"pdate_{proj_key}_{task_name}", label_visibility="collapsed", disabled=not is_active
                )

                st.session_state["project_tasks_store"][proj_key][task_name] = {
                    "active": is_active,
                    "done": is_done,
                    "user": assigned_user,
                    "date": assign_date
                }
                
                auto_time = tasks_database.get(task_name, 0.0)
                
                if is_active:
                    task_hours = (auto_time * project_main_qty) / 60
                    total_project_hours += task_hours
                    total_tasks_count += 1
                    
                    if is_done:
                        completed_project_hours += task_hours
                        completed_tasks_count += 1
                        c_time.markdown("✅ **Done**")
                    else:
                        c_time.metric("λ/σετ", f"{auto_time}λ")
                else:
                    c_time.caption("Ανενεργό")

        # --- 3. ΠΡΟΟΔΟΣ PROJECT & ΣΥΝΟΛΙΚΟΙ ΥΠΟΛΟΓΙΣΜΟΙ ---
        st.divider()
        progress_pct = int((completed_tasks_count / total_tasks_count) * 100) if total_tasks_count > 0 else 0
        remaining_hours = round(total_project_hours - completed_project_hours, 1)

        st.markdown(f"### 📊 Πρόοδος Παραγωγής Project {selected_project}: **{progress_pct}%**")
        st.progress(progress_pct / 100)

        m1, m2, m3 = st.columns(3)
        m1.metric("Συνολικές Ώρες (Γενικές + Υλικών)", f"{round(total_project_hours, 1)} Ώρες")
        m2.metric("Ώρες που Ολοκληρώθηκαν", f"{round(completed_project_hours, 1)} Ώρες")
        m3.metric("Υπολειπόμενες Ώρες", f"{remaining_hours} Ώρες", delta=f"-{remaining_hours}h" if remaining_hours > 0 else "Έτοιμο!")

elif option == "Ημερήσιο Πλάνο Παραγωγής":
    st.header("🗓️ Συγκεντρωτικό Πλάνο Παραγωγής & Έλεγχος Διαθεσιμότητας")
    
    target_date = st.date_input("Επιλέξτε Ημερομηνία Πλάνου:", value=date.today(), format="DD/MM/YYYY")
    greek_day_name = WEEKDAYS_GREEK.get(target_date.weekday(), "Δευτέρα")
    st.caption(f"Ημέρα εβδομάδας: **{greek_day_name}**")
    st.divider()

    daily_tasks = []
    
    for p_key, p_tasks_dict in st.session_state["project_tasks_store"].items():
        if isinstance(p_tasks_dict, dict):
            proj_name = p_key.replace("proj_", "")
            
            proj_qty = 1
            if not procurement_df.empty:
                p_items = procurement_df[procurement_df["Project"] == proj_name]
                for _, r in p_items.iterrows():
                    if str(r["Ποσότητα"]).isdigit():
                        proj_qty = max(proj_qty, int(r["Ποσότητα"]))

            for task_name, p_data in p_tasks_dict.items():
                if isinstance(p_data, dict) and p_data.get("active", False) and p_data.get("date") == target_date:
                    t_user = p_data.get("user", "- Χωρίς Ανάθεση -")
                    t_done = p_data.get("done", False)
                    auto_time = tasks_database.get(task_name, 0.0)
                    hours = (auto_time * proj_qty) / 60
                    
                    daily_tasks.append({
                        "ID": "Project Task",
                        "Project": proj_name,
                        "Υλικό": "Γενική Σύνθεση / Box",
                        "Ποσότητα": proj_qty,
                        "Εργασία": task_name,
                        "Υπεύθυνος": t_user,
                        "Ώρες": round(hours, 2),
                        "Κατάσταση": "✅ Ολοκληρώθηκε" if t_done else "⏳ Σε Εκκρεμότητα"
                    })

    if not procurement_df.empty:
        for idx, row in procurement_df.iterrows():
            item_id = str(row["ID"])
            project_name = row["Project"]
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            
            item_tasks = st.session_state["tasks_store"].get(item_id, [])
            
            for t_data in item_tasks:
                t_task = t_data["task"]
                t_user = t_data["user"]
                t_date = t_data["date"]
                t_done = t_data["done"]
                
                if t_task != "- Επιλογή Εργασίας -" and t_date == target_date:
                    auto_time = tasks_database.get(t_task, 0.0)
                    hours = (auto_time * qty) / 60
                    
                    daily_tasks.append({
                        "ID": item_id,
                        "Project": project_name,
                        "Υλικό": material,
                        "Ποσότητα": qty,
                        "Εργασία": t_task,
                        "Υπεύθυνος": t_user,
                        "Ώρες": round(hours, 2),
                        "Κατάσταση": "✅ Ολοκληρώθηκε" if t_done else "⏳ Σε Εκκρεμότητα"
                    })

    if daily_tasks:
        daily_df = pd.DataFrame(daily_tasks)
        
        st.subheader(f"📌 Εργασίες για τις {target_date.strftime('%d/%m/%Y')} ({len(daily_df)} Tasks)")
        
        st.markdown("#### 👥 Φόρτος Εργασίας & Διαθεσιμότητα Ομάδας")
        
        day_availability = availability_database.get(greek_day_name, {})
        team_summary = daily_df.groupby("Υπεύθυνος")["Ώρες"].sum().reset_index()
        
        cols = st.columns(max(len(team_summary), 1))
        for i, r in team_summary.iterrows():
            member_name = r["Υπεύθυνος"]
            assigned_hrs = round(r["Ώρες"], 2)
            
            if member_name != "- Χωρίς Ανάθεση -":
                max_hrs = day_availability.get(member_name, 6.0)
                delta_hrs = round(assigned_hrs - max_hrs, 2)
                
                if delta_hrs > 0:
                    cols[i].metric(
                        f"⚠️ {member_name}", 
                        f"{assigned_hrs} / {max_hrs}h", 
                        delta=f"+{delta_hrs}h Υπερκάλυψη", 
                        delta_color="inverse"
                    )
                else:
                    cols[i].metric(
                        f"🟢 {member_name}", 
                        f"{assigned_hrs} / {max_hrs}h", 
                        delta=f"{delta_hrs}h Διαθέσιμο", 
                        delta_color="normal"
                    )
            else:
                cols[i].metric(f"❓ {member_name}", f"{assigned_hrs} Ώρες")
            
        st.divider()
        st.markdown("#### 📋 Αναλυτικός Πίνακας Εργασιών")
        
        column_config = {
            col: st.column_config.Column(alignment="center") for col in daily_df.columns
        }
        st.dataframe(daily_df, use_container_width=True, hide_index=True, column_config=column_config)
    else:
        st.info(f"Δεν έχουν προγραμματιστεί εργασίες για τις {target_date.strftime('%d/%m/%Y')}.")

elif option == "Πρότυπα Χρόνων & Ομάδα":
    st.header("📊 Βάση Δεδομένων Χρόνων & Ομάδας")
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f"⏱️ Πρότυπα Χρόνων ({len(tasks_database)} Εργασίες)")
        st.json(tasks_database)
    with col_b:
        st.subheader("👥 Ομάδα Παραγωγής & Ημερήσια Όρια Ώρων")
        st.write(availability_database)
