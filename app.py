import streamlit as st
import pandas as pd
from datetime import date, timedelta

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

WEEKDAYS_SHORT_GREEK = {
    0: "Δευ",
    1: "Τρι",
    2: "Τετ",
    3: "Πεμ",
    4: "Παρ",
    5: "Σαβ",
    6: "Κυρ"
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

option = st.sidebar.selectbox("Επιλογή Οθόνης", [
    "Dashboard Παραγωγής",
    "Καρτέλα Project", 
    "Ημερήσιο Πλάνο Παραγωγής", 
    "Ημερήσιο Πρόγραμμα (Ανά Τεχνίτη)", 
    "Εβδομαδιαίο Projection",
    "Πρότυπα Χρόνων & Ομάδα"
])

if option == "Dashboard Παραγωγής":
    st.header("📈 Dashboard & Επισκόπηση Παραγωγής")
    
    if not procurement_df.empty:
        projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
        
        dashboard_data = []
        tot_all_hours = 0.0
        tot_done_hours = 0.0
        tot_tasks_count = 0
        tot_done_tasks = 0
        
        for p_name in projects_list:
            filtered_p = procurement_df[procurement_df["Project"] == p_name]
            
            p_main_qty = 1
            for _, r in filtered_p.iterrows():
                if str(r["Ποσότητα"]).isdigit():
                    p_main_qty = max(p_main_qty, int(r["Ποσότητα"]))

            p_total_hrs = 0.0
            p_done_hrs = 0.0
            p_tasks_cnt = 0
            p_done_cnt = 0
            
            for idx, r in filtered_p.iterrows():
                item_id = str(r["ID"])
                u_key = f"{item_id}_{idx}"
                qty = int(r["Ποσότητα"]) if str(r["Ποσότητα"]).isdigit() else 1
                
                item_tasks = st.session_state["tasks_store"].get(u_key, [])
                for t in item_tasks:
                    if t["task"] != "- Επιλογή Εργασίας -":
                        auto_t = tasks_database.get(t["task"], 0.0)
                        hrs = (auto_t * qty) / 60
                        p_total_hrs += hrs
                        p_tasks_cnt += 1
                        if t["done"]:
                            p_done_hrs += hrs
                            p_done_cnt += 1

            p_key = f"proj_{p_name}"
            p_tasks_dict = st.session_state["project_tasks_store"].get(p_key, {})
            if isinstance(p_tasks_dict, dict):
                for t_name, p_data in p_tasks_dict.items():
                    if isinstance(p_data, dict) and p_data.get("active", False):
                        auto_t = tasks_database.get(t_name, 0.0)
                        hrs = (auto_t * p_main_qty) / 60
                        p_total_hrs += hrs
                        p_tasks_cnt += 1
                        if p_data.get("done", False):
                            p_done_hrs += hrs
                            p_done_cnt += 1

            tot_all_hours += p_total_hrs
            tot_done_hours += p_done_hrs
            tot_tasks_count += p_tasks_cnt
            tot_done_tasks += p_done_cnt
            
            p_progress = int((p_done_cnt / p_tasks_cnt) * 100) if p_tasks_cnt > 0 else 0
            
            dashboard_data.append({
                "Project": p_name,
                "Υλικά": len(filtered_p),
                "Σύνολο Tasks": p_tasks_cnt,
                "Ολοκληρωμένα Tasks": p_done_cnt,
                "Συνολικές Ώρες": round(p_total_hrs, 1),
                "Υπολειπόμενες Ώρες": round(p_total_hrs - p_done_hrs, 1),
                "Πρόοδος (%)": f"{p_progress}%"
            })

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Ενεργά Projects", len(projects_list))
        c2.metric("Συνολικές Ώρες Παραγωγής", f"{round(tot_all_hours, 1)}h")
        
        overall_pct = int((tot_done_tasks / tot_tasks_count) * 100) if tot_tasks_count > 0 else 0
        c3.metric("Συνολική Πρόοδος Παραγωγής", f"{overall_pct}%")
        c4.metric("Εκκρεμή Tasks", tot_tasks_count - tot_done_tasks)

        st.divider()
        st.subheader("📊 Κατάσταση & Πρόοδος ανά Project")
        
        if dashboard_data:
            dash_df = pd.DataFrame(dashboard_data)
            st.dataframe(dash_df, use_container_width=True, hide_index=True)

elif option == "Καρτέλα Project":
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

        st.subheader(f"📦 Εργασίες ανά Υλικό ({len(filtered_df)} Υλικά)")

        for idx, row in filtered_df.iterrows():
            item_id = str(row["ID"])
            unique_item_key = f"{item_id}_{idx}"
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            status = row["Status Procurement"]
            
            if unique_item_key not in st.session_state["tasks_store"]:
                st.session_state["tasks_store"][unique_item_key] = [
                    {"done": False, "task": "- Επιλογή Εργασίας -", "user": "- Χωρίς Ανάθεση -", "date": date.today()}
                ]
            
            item_tasks = st.session_state["tasks_store"][unique_item_key]
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
                        
                        is_done = c_check.checkbox("", value=t_data["done"], key=f"chk_{unique_item_key}_{t_idx}")
                        
                        task_idx = task_options.index(t_data["task"]) if t_data["task"] in task_options else 0
                        selected_task = c_task.selectbox(
                            "Εργασία", task_options, index=task_idx, 
                            key=f"task_{unique_item_key}_{t_idx}", label_visibility="collapsed"
                        )
                        
                        user_idx = team_options.index(t_data["user"]) if t_data["user"] in team_options else 0
                        assigned_user = c_user.selectbox(
                            "Ανάθεση", team_options, index=user_idx, 
                            key=f"user_{unique_item_key}_{t_idx}", label_visibility="collapsed"
                        )
                        
                        assign_date = c_date.date_input(
                            "Ημερομηνία", value=t_data["date"], format="DD/MM/YYYY", 
                            key=f"date_{unique_item_key}_{t_idx}", label_visibility="collapsed"
                        )

                        st.session_state["tasks_store"][unique_item_key][t_idx] = {
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
                        
                        if c_del.button("🗑️", key=f"del_{unique_item_key}_{t_idx}"):
                            st.session_state["tasks_store"][unique_item_key].pop(t_idx)
                            st.rerun()

                    col_btn1, col_btn2, _ = st.columns([0.35, 0.35, 0.3])
                    if col_btn1.button("➕ Προσθήκη Εργασίας Υλικού", key=f"add_btn_{unique_item_key}"):
                        st.session_state["tasks_store"][unique_item_key].append(
                            {"done": False, "task": "- Επιλογή Εργασίας -", "user": "- Χωρίς Ανάθεση -", "date": date.today()}
                        )
                        st.rerun()
                    
                    if len(item_tasks) > 0:
                        if col_btn2.button("➖ Αφαίρεση Εργασίας", key=f"rem_btn_{unique_item_key}"):
                            st.session_state["tasks_store"][unique_item_key].pop()
                            st.rerun()

        st.divider()

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
            unique_item_key = f"{item_id}_{idx}"
            project_name = row["Project"]
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            
            item_tasks = st.session_state["tasks_store"].get(unique_item_key, [])
            
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

elif option == "Ημερήσιο Πρόγραμμα (Ανά Τεχνίτη)":
    st.header("👤 Ημερήσιο Πρόγραμμα Εργασιών ανά Τεχνίτη")
    
    c_date, c_user = st.columns([1, 1])
    target_date = c_date.date_input("Ημερομηνία:", value=date.today(), format="DD/MM/YYYY")
    selected_member = c_user.selectbox("Επιλέξτε Τεχνίτη / Εργαζόμενο:", team_database)
    
    st.divider()

    worker_tasks = []

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
                if isinstance(p_data, dict) and p_data.get("active", False):
                    if p_data.get("user") == selected_member and p_data.get("date") == target_date:
                        auto_time = tasks_database.get(task_name, 0.0)
                        hours = (auto_time * proj_qty) / 60
                        worker_tasks.append({
                            "type": "project",
                            "key": f"p_{p_key}_{task_name}",
                            "project": proj_name,
                            "item": "Γενική Σύνθεση / Box",
                            "qty": proj_qty,
                            "task": task_name,
                            "hours": round(hours, 2),
                            "done": p_data.get("done", False),
                            "status_proc": "READY"
                        })

    if not procurement_df.empty:
        for idx, row in procurement_df.iterrows():
            item_id = str(row["ID"])
            unique_item_key = f"{item_id}_{idx}"
            project_name = row["Project"]
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            proc_status = row["Status Procurement"]
            
            item_tasks = st.session_state["tasks_store"].get(unique_item_key, [])
            
            for t_data in item_tasks:
                if t_data.get("user") == selected_member and t_data.get("date") == target_date:
                    t_task = t_data.get("task")
                    if t_task != "- Επιλογή Εργασίας -":
                        auto_time = tasks_database.get(t_task, 0.0)
                        hours = (auto_time * qty) / 60
                        worker_tasks.append({
                            "type": "item",
                            "key": f"i_{unique_item_key}_{t_task}",
                            "project": project_name,
                            "item": f"[{item_id}] {material}",
                            "qty": qty,
                            "task": t_task,
                            "hours": round(hours, 2),
                            "done": t_data.get("done", False),
                            "status_proc": proc_status
                        })

    st.subheader(f"📋 Πρόγραμμα για τον/την {selected_member} — {target_date.strftime('%d/%m/%Y')}")

    if worker_tasks:
        total_w_hours = sum(t["hours"] for t in worker_tasks)
        st.info(f"💡 Συνολικός εκτιμώμενος χρόνος εργασίας: **{round(total_w_hours, 1)} Ώρες** ({len(worker_tasks)} Tasks)")
        
        st.divider()

        for w_idx, wt in enumerate(worker_tasks):
            col_c, col_proj, col_mat, col_task, col_qty, col_h, col_proc = st.columns([0.08, 0.20, 0.28, 0.22, 0.08, 0.10, 0.14])
            
            col_c.write(f"#{w_idx+1}")
            col_proj.markdown(f"**{wt['project']}**")
            col_mat.write(wt['item'])
            col_task.markdown(f"`{wt['task']}`")
            col_qty.write(f"{wt['qty']} τμχ")
            col_h.caption(f"{wt['hours']}h")
            
            if wt['status_proc'] in ["OK STOCK", "RECEIVED", "READY"]:
                col_proc.success(wt['status_proc'])
            else:
                col_proc.warning(wt['status_proc'])
    else:
        st.success(f"🎉 Δεν έχουν ανατεθεί εργασίες στον/στην {selected_member} για τις {target_date.strftime('%d/%m/%Y')}.")

elif option == "Εβδομαδιαίο Projection":
    st.header("📆 Πρόβλεψη Φόρτου Εργασίας (Projection)")
    
    start_monday = date.today() - timedelta(days=date.today().weekday())
    
    col_w_choice, col_range_choice, col_weekend = st.columns([1, 1, 1])
    
    week_choice = col_w_choice.selectbox("ΕΝΑΡΞΗ ΠΡΟΒΟΛΗΣ:", [
        "Τρέχουσα Εβδομάδα",
        "Επόμενη Εβδομάδα (+1)",
        "Μεθεπόμενη Εβδομάδα (+2)",
        "Προσαρμοσμένη Ημερομηνία"
    ])
    
    range_weeks = col_range_choice.selectbox("ΕΥΡΟΣ ΠΡΟΒΟΛΗΣ:", [
        "1 Εβδομάδα",
        "2 Εβδομάδες",
        "4 Εβδομάδες / Μήνας"
    ])
    
    include_weekends = col_weekend.checkbox("📅 Συμπερίληψη Σαββατοκύριακων (ΣΚ)", value=False)
    
    num_weeks = 1
    if "2 Εβδομάδες" in range_weeks:
        num_weeks = 2
    elif "4 Εβδομάδες" in range_weeks:
        num_weeks = 4

    if week_choice == "Τρέχουσα Εβδομάδα":
        sel_start = start_monday
    elif week_choice == "Επόμενη Εβδομάδα (+1)":
        sel_start = start_monday + timedelta(days=7)
    elif week_choice == "Μεθεπόμενη Εβδομάδα (+2)":
        sel_start = start_monday + timedelta(days=14)
    else:
        sel_start = st.date_input("Επιλέξτε Δευτέρα Εναρξης:", value=start_monday, format="DD/MM/YYYY")
    
    days_per_week = 7 if include_weekends else 5
    
    # Δημιουργία δομής ημερών ανά εβδομάδα
    weeks_days_list = []
    all_flat_days = []
    
    for w in range(num_weeks):
        w_monday = sel_start + timedelta(days=w*7)
        w_days = [w_monday + timedelta(days=i) for i in range(days_per_week)]
        weeks_days_list.append((w+1, w_monday, w_days))
        all_flat_days.extend(w_days)

    # Υπολογισμοί για τα KPI Metrics συνολικά
    total_assigned_range = 0.0
    total_available_range = 0.0
    overbooked_days_count = 0

    for d in all_flat_days:
        g_day = WEEKDAYS_GREEK.get(d.weekday(), "Δευτέρα")
        day_avail = availability_database.get(g_day, {})
        day_max = sum(day_avail.get(m, 6.0) for m in team_database)
        total_available_range += day_max

    # Υπολογισμός ωρών ανά εργαζόμενο και ανά ημέρα
    # Matrix: {w_idx: {m: {day_col_title: hours}}}
    weeks_matrices = []
    
    for w_num, w_monday, w_days in weeks_days_list:
        w_matrix = {m: {f"{WEEKDAYS_SHORT_GREEK[d.weekday()]} {d.strftime('%d/%m')}": 0.0 for d in w_days} for m in team_database}
        d_totals = {f"{WEEKDAYS_SHORT_GREEK[d.weekday()]} {d.strftime('%d/%m')}": 0.0 for d in w_days}
        
        # 1. Project Level Tasks
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
                    if isinstance(p_data, dict) and p_data.get("active", False):
                        t_date = p_data.get("date")
                        t_user = p_data.get("user")
                        if t_user in w_matrix and t_date in w_days:
                            auto_time = tasks_database.get(task_name, 0.0)
                            hrs = (auto_time * proj_qty) / 60
                            col_str = f"{WEEKDAYS_SHORT_GREEK[t_date.weekday()]} {t_date.strftime('%d/%m')}"
                            w_matrix[t_user][col_str] += hrs
                            d_totals[col_str] += hrs

        # 2. Item Level Tasks
        if not procurement_df.empty:
            for idx, row in procurement_df.iterrows():
                item_id = str(row["ID"])
                unique_item_key = f"{item_id}_{idx}"
                qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
                
                item_tasks = st.session_state["tasks_store"].get(unique_item_key, [])
                for t_data in item_tasks:
                    t_task = t_data.get("task")
                    t_user = t_data.get("user")
                    t_date = t_data.get("date")
                    
                    if t_task != "- Επιλογή Εργασίας -" and t_user in w_matrix and t_date in w_days:
                        auto_time = tasks_database.get(t_task, 0.0)
                        hrs = (auto_time * qty) / 60
                        col_str = f"{WEEKDAYS_SHORT_GREEK[t_date.weekday()]} {t_date.strftime('%d/%m')}"
                        w_matrix[t_user][col_str] += hrs
                        d_totals[col_str] += hrs

        w_assigned_tot = sum(d_totals.values())
        total_assigned_range += w_assigned_tot
        
        for d in w_days:
            col_str = f"{WEEKDAYS_SHORT_GREEK[d.weekday()]} {d.strftime('%d/%m')}"
            g_day = WEEKDAYS_GREEK.get(d.weekday(), "Δευτέρα")
            day_avail = availability_database.get(g_day, {})
            day_max = sum(day_avail.get(m, 6.0) for m in team_database)
            if d_totals[col_str] > day_max:
                overbooked_days_count += 1
                
        weeks_matrices.append((w_num, w_monday, w_days, w_matrix))

    load_ratio = int((total_assigned_range / total_available_range) * 100) if total_available_range > 0 else 0

    st.divider()
    kc1, kc2, kc3, kc4 = st.columns(4)
    kc1.metric("Προγραμματισμένες Ώρες", f"{round(total_assigned_range, 1)}h")
    kc2.metric(f"Διαθέσιμες Ώρες ({num_weeks} εβδ.)", f"{round(total_available_range, 1)}h")
    
    if overbooked_days_count > 0:
        kc3.metric("Overbooked Ημέρες", f"⚠️ {overbooked_days_count} / {len(all_flat_days)}", delta_color="inverse")
    else:
        kc3.metric("Overbooked Ημέρες", f"🟢 0 / {len(all_flat_days)}")
        
    kc4.metric("Πληρότητα Περιόδου", f"{load_ratio}%", delta=f"{load_ratio - 100}%" if load_ratio > 100 else "Εντός Ορίων")

    st.divider()

    # Προβολή Πινάκων — Ο ένας κάτω από τον άλλο
    for w_num, w_monday, w_days, w_matrix in weeks_matrices:
        w_sunday = w_days[-1]
        st.subheader(f"📅 Εβδομάδα {w_num}: {w_monday.strftime('%d/%m/%Y')} έως {w_sunday.strftime('%d/%m/%Y')}")
        
        proj_df = pd.DataFrame(w_matrix).T
        proj_df = proj_df.round(1)
        proj_df["Σύνολο Εβδομάδας (h)"] = proj_df.sum(axis=1)

        st.dataframe(proj_df, use_container_width=True)
        st.markdown("<br>", unsafe_allow_html=True)

elif option == "Πρότυπα Χρόνων & Ομάδα":
    st.header("📊 Βάση Δεδομένων Χρόνων & Ομάδας")
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f"⏱️ Πρότυπα Χρόνων ({len(tasks_database)} Εργασίες)")
        st.json(tasks_database)
    with col_b:
        st.subheader("👥 Ομάδα Παραγωγής & Ημερήσια Όρια Ώρων")
        st.write(availability_database)
