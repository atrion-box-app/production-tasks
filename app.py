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

@st.cache_data(ttl=10)
def load_all_data():
    # 1. Φόρτωση Procurement
    df_proc_raw = pd.read_csv(PROC_CSV_URL, header=None)
    indices = [18, 3, 17, 1, 2, 5, 6, 4, 10, 12]
    df_proc = df_proc_raw.iloc[1:, indices].copy()
    df_proc.columns = [
        "ID", "Ημερομηνία Παράδοσης", "Είδος Δώρου", "Project", 
        "Ποσότητα", "Προμηθευτής", "Υλικό / Προϊόν", 
        "Αναμενόμενη Ημ. Παραλαβής", "Αναμενόμενη Ποσότητα Παραλαβής", "Status Procurement"
    ]
    df_proc = df_proc.fillna("-")

    # 2. Φόρτωση ΟΛΩΝ των Προτύπων Χρόνων (Σάρωση όλων των στηλών)
    tasks_dict = {}
    try:
        df_times_raw = pd.read_csv(TIMES_CSV_URL, header=None)
        
        # Σαρώνουμε ανά ζεύγη στηλών (0-1, 3-4, 7-8 κτλ.)
        for col_idx in range(len(df_times_raw.columns) - 1):
            for row_idx in range(len(df_times_raw)):
                task_name = str(df_times_raw.iloc[row_idx, col_idx]).strip()
                time_val_raw = df_times_raw.iloc[row_idx, col_idx + 1]
                
                # Έλεγχος αν είναι έγκυρο όνομα εργασίας και αριθμητικός χρόνος
                if task_name and task_name.lower() not in ["nan", "none", "τύπος εργασίας / υλικό"] and not task_name.startswith("TASK"):
                    try:
                        time_val = float(str(time_val_raw).replace(',', '.'))
                        if time_val >= 0:
                            tasks_dict[task_name] = time_val
                    except ValueError:
                        pass
    except Exception as e:
        tasks_dict = {"Έλεγχος (εύκολο)": 1.0, "Συναρμολόγηση": 2.0, "Συσκευασία": 1.5}

    # 3. Φόρτωση Ομάδας / Προσωπικού
    try:
        df_team_raw = pd.read_csv(TEAM_CSV_URL)
        team_members = [c.strip() for c in df_team_raw.columns if c and "Unnamed" not in c and c not in ["Ημέρα", "Σύνολο διαθέσιμων ωρών"]]
        if not team_members:
            team_members = ["Βαγγέλης Μ.", "Βαγγέλης JR.", "Εποχικός 1", "Εποχικός 2", "Ana", "Alex"]
    except Exception as e:
        team_members = ["Βαγγέλης Μ.", "Βαγγέλης JR.", "Εποχικός 1", "Εποχικός 2", "Ana", "Alex"]

    return df_proc, tasks_dict, team_members

procurement_df, tasks_database, team_database = load_all_data()

option = st.sidebar.selectbox("Επιλογή Οθόνης", ["Καρτέλα Project", "Πρότυπα Χρόνων & Ομάδα"])

if option == "Καρτέλα Project":
    st.header("📋 Διαχείριση Παραγωγής & Αναθέσεις ανά Project")
    
    projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
    selected_project = st.selectbox("Επιλέξτε Project:", projects_list)
    
    filtered_df = procurement_df[procurement_df["Project"] == selected_project].copy()
    
    st.subheader(f"📦 Υλικά & Tasks για το {selected_project} ({len(filtered_df)} Υλικά)")
    
    total_project_hours = 0.0
    task_options = ["- Επιλογή Εργασίας -"] + sorted(list(tasks_database.keys()))
    team_options = ["- Χωρίς Ανάθεση -"] + team_database

    for idx, row in filtered_df.iterrows():
        item_id = row["ID"]
        material = row["Υλικό / Προϊόν"]
        qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
        status = row["Status Procurement"]
        
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
                st.markdown("**⚙️ Στάδια Παραγωγής & Αυτόματοι Χρόνοι**")
                
                for t_num in [1, 2]:
                    st.caption(f"**Στάδιο {t_num}**")
                    c_check, c_task, c_user, c_date, c_time = st.columns([0.1, 0.35, 0.25, 0.20, 0.10])
                    
                    # Checkbox
                    done = c_check.checkbox("", key=f"chk_{item_id}_{t_num}")
                    
                    # Dropdown Εργασιών
                    selected_task = c_task.selectbox(
                        "Εργασία", task_options, 
                        key=f"task_{item_id}_{t_num}", 
                        label_visibility="collapsed"
                    )
                    
                    auto_time = tasks_database.get(selected_task, 0.0)
                    
                    # Dropdown Ομάδας
                    assigned_user = c_user.selectbox(
                        "Ανάθεση", team_options, 
                        key=f"user_{item_id}_{t_num}", 
                        label_visibility="collapsed"
                    )
                    
                    # Ημερομηνία Ανάθεσης
                    assign_date = c_date.date_input(
                        "Ημερομηνία", value=date.today(), 
                        format="DD/MM/YYYY",
                        key=f"date_{item_id}_{t_num}", 
                        label_visibility="collapsed"
                    )
                    
                    # Εμφάνιση Αυτόματου Χρόνου
                    c_time.metric("λ/τμχ", f"{auto_time}λ")
                    
                    task_hours = (auto_time * qty) / 60
                    total_project_hours += task_hours

    st.divider()
    st.success(f"🎯 **Συνολικός Χρόνος Παραγωγής για το Project {selected_project}: {round(total_project_hours, 1)} Ώρες**")

elif option == "Πρότυπα Χρόνων & Ομάδα":
    st.header("📊 Βάση Δεδομένων Χρόνων & Ομάδας")
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f"⏱️ Πρότυπα Χρόνων ({len(tasks_database)} Εργασίες)")
        st.json(tasks_database)
    with col_b:
        st.subheader("👥 Ομάδα Παραγωγής")
        st.write(team_database)
