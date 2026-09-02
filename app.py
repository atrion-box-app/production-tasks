import streamlit as st
import pandas as pd
import gspread
import json
import base64
import hashlib
import time
from google.oauth2.service_account import Credentials
from datetime import date, datetime, timedelta
from streamlit_option_menu import option_menu
import xlsxwriter
from io import BytesIO
import re

# --- ΡΥΘΜΙΣΕΙΣ ΣΕΛΙΔΑΣ ---
st.set_page_config(
    page_title="Production Tasks App",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- CUSTOM CSS ---
st.markdown("""
<style>
    @media (max-width: 768px) {
        .stColumns { flex-direction: column !important; }
        .stButton button { width: 100% !important; }
        .stSelectbox, .stDateInput { margin-bottom: 10px; }
    }
</style>
""", unsafe_allow_html=True)

# --- AUTHENTICATION SYSTEM ---
def init_auth():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "username" not in st.session_state:
        st.session_state.username = None
    if "login_attempts" not in st.session_state:
        st.session_state.login_attempts = 0
    if "last_login_attempt" not in st.session_state:
        st.session_state.last_login_attempt = None

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    valid_users = {
        "admin": hash_password("admin123"),
        "manager": hash_password("manager123"),
        "operator": hash_password("operator123")
    }
    
    if st.session_state.login_attempts >= 5:
        last_attempt = st.session_state.last_login_attempt
        if last_attempt and (datetime.now() - last_attempt).seconds < 300:
            st.error("🔒 Too many failed attempts. Please wait 5 minutes.")
            return False
    
    if username in valid_users and valid_users[username] == hash_password(password):
        st.session_state.authenticated = True
        st.session_state.username = username
        st.session_state.login_attempts = 0
        return True
    
    st.session_state.login_attempts += 1
    st.session_state.last_login_attempt = datetime.now()
    return False

def login_form():
    with st.container():
        st.markdown("""
        <div style="max-width:400px;margin:100px auto;padding:40px;border-radius:10px;box-shadow:0 4px 6px rgba(0,0,0,0.1);background:white;">
            <h2 style="text-align:center;color:#1e88e5;">🏭 Production Tasks</h2>
            <p style="text-align:center;color:#666;">Please login to continue</p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("👤 Username", placeholder="Enter your username")
            password = st.text_input("🔑 Password", type="password", placeholder="Enter your password")
            submitted = st.form_submit_button("🚀 Login", use_container_width=True)
            
            if submitted:
                if verify_user(username, password):
                    st.success("✅ Login successful!")
                    st.rerun()
                else:
                    st.error("❌ Invalid username or password")

def logout():
    st.session_state.authenticated = False
    st.session_state.username = None
    st.rerun()

# --- GOOGLE SHEETS CONNECTION ---
@st.cache_resource
def get_gspread_client():
    if "gcp_service_account" not in st.secrets:
        return None, "Το [gcp_service_account] δεν βρέθηκε στα Secrets του Streamlit."
    
    try:
        sec = st.secrets["gcp_service_account"]
        if "b64_json" in sec:
            decoded_bytes = base64.b64decode(sec["b64_json"])
            creds_dict = json.loads(decoded_bytes.decode("utf-8"))
        elif "json_str" in sec:
            creds_dict = json.loads(sec["json_str"])
        else:
            creds_dict = dict(sec)
            if "private_key" in creds_dict:
                creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n").strip()
        
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(credentials)
        return client, None
    except Exception as e:
        return None, str(e)

# --- CONSTANTS ---
PROC_SHEET_ID = "1QhTd58vuulaC_73sgbjuwG5MVxT6c1c_-MbhypGx0fA"
PROC_GID = "1639392743"
PROC_CSV_URL = f"https://docs.google.com/spreadsheets/d/{PROC_SHEET_ID}/export?format=csv&gid={PROC_GID}"

MY_SHEET_ID = "1rps5ha4wyo8DQ3zwUTqS5BSNMrJPatvqdh8M0iMHVEg"
TIMES_GID = "2126316973"
TEAM_GID = "1303086311"

TIMES_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={TIMES_GID}"
TEAM_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={TEAM_GID}"

WEEKDAYS_GREEK = {
    0: "Δευτέρα", 1: "Τρίτη", 2: "Τετάρτη", 3: "Πέμπτη", 
    4: "Παρασκευή", 5: "Σάββατο", 6: "Κυριακή"
}
WEEKDAYS_SHORT_GREEK = {
    0: "Δευ", 1: "Τρι", 2: "Τετ", 3: "Πεμ", 
    4: "Παρ", 5: "Σαβ", 6: "Κυρ"
}

FIXED_PROJECT_TASKS = [
    "Σύνθεση (κουτί)",
    "Σύνθεση (πουγκί / τσάντα)",
    "Σύνθεση (χειροποίητο)",
    "Φωτογράφιση",
    "Τοποθέτηση σε χαρτοκιβώτια"
]

# --- DATA LOADING ---
@st.cache_data(ttl=60, show_spinner=False)
def load_all_data(version=0):
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

# --- ASSIGNMENTS MANAGEMENT ---
@st.cache_data(ttl=30)
def load_assignments_from_sheet():
    assignments_item = {}
    assignments_proj = {}
    
    gc, _ = get_gspread_client()
    if not gc:
        return assignments_item, assignments_proj
    
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Assignments")
        records = sheet.get_all_records()
        
        for r in records:
            p_name = str(r.get("Project", ""))
            item_id = str(r.get("Item_ID", ""))
            task_name = str(r.get("Task_Name", ""))
            user = str(r.get("Assigned_User", "- Χωρίς Ανάθεση -"))
            assign_date_str = str(r.get("Assigned_Date", ""))
            done = True if str(r.get("Status_Done", "")).upper() in ["TRUE", "1", "YES"] else False
            task_type = str(r.get("Task_Type", ""))

            try:
                assign_date = datetime.strptime(assign_date_str, "%Y-%m-%d").date()
            except Exception:
                assign_date = date.today()

            if task_type == "PROJECT":
                p_key = f"proj_{p_name}"
                if p_key not in assignments_proj:
                    assignments_proj[p_key] = {}
                assignments_proj[p_key][task_name] = {
                    "active": True,
                    "done": done,
                    "user": user,
                    "date": assign_date
                }
            else:
                if item_id not in assignments_item:
                    assignments_item[item_id] = []
                assignments_item[item_id].append({
                    "done": done,
                    "task": task_name,
                    "user": user,
                    "date": assign_date
                })
    except Exception:
        pass
    
    return assignments_item, assignments_proj

def save_all_assignments_to_sheet():
    gc, _ = get_gspread_client()
    if not gc:
        st.warning("⚠️ Δεν είναι δυνατή η αποθήκευση λόγω σφάλματος σύνδεσης API.")
        return False
    
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Assignments")
        rows = [["Project", "Item_ID", "Task_Name", "Assigned_User", "Assigned_Date", "Status_Done", "Task_Type"]]

        item_to_project = {}
        if 'procurement_df' in st.session_state and st.session_state.procurement_df is not None:
            for _, r in st.session_state.procurement_df.iterrows():
                item_to_project[str(r["ID"])] = str(r["Project"])

        for u_key, t_list in st.session_state.get("tasks_store", {}).items():
            item_id = u_key.split("_")[0]
            proj_name = item_to_project.get(item_id, "-")
            for t in t_list:
                if t.get("task") and t.get("task") != "- Επιλογή Εργασίας -":
                    rows.append([
                        proj_name, item_id, t.get("task"), t.get("user"), 
                        str(t.get("date")), str(t.get("done")), "ITEM"
                    ])

        for p_key, p_dict in st.session_state.get("project_tasks_store", {}).items():
            proj_name = p_key.replace("proj_", "")
            if isinstance(p_dict, dict):
                for t_name, p_data in p_dict.items():
                    if isinstance(p_data, dict) and p_data.get("active", False):
                        rows.append([
                            proj_name, "-", t_name, p_data.get("user"), 
                            str(p_data.get("date")), str(p_data.get("done")), "PROJECT"
                        ])

        sheet.clear()
        sheet.update(range_name="A1", values=rows)
        st.session_state.last_save = datetime.now()
        return True
        
    except Exception as e:
        st.error(f"❌ Σφάλμα κατά την αποθήκευση: {e}")
        return False

# --- AUDIT LOG ---
def add_to_audit_log(action, details):
    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []
    
    st.session_state.audit_log.append({
        "timestamp": datetime.now(),
        "user": st.session_state.get("username", "unknown"),
        "action": action,
        "details": details
    })
    
    if len(st.session_state.audit_log) > 1000:
        st.session_state.audit_log = st.session_state.audit_log[-1000:]

# --- EXPORT FUNCTIONS ---
def generate_printable_html(title, date_str, df_data):
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{title}</title>
        <style>
            * {{ box-sizing: border-box; }}
            body {{ font-family: Arial, sans-serif; margin: 20px; color: #333; }}
            h2 {{ color: #1e88e5; border-bottom: 2px solid #1e88e5; padding-bottom: 5px; }}
            .date {{ font-size: 14px; color: #666; margin-bottom: 20px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; }}
            th, td {{ border: 1px solid #ddd; padding: 10px; text-align: left; font-size: 13px; }}
            th {{ background-color: #f2f2f2; font-weight: bold; }}
            .done {{ color: green; font-weight: bold; }}
            .pending {{ color: #d32f2f; font-weight: bold; }}
            @media print {{ button {{ display: none; }} }}
        </style>
    </head>
    <body>
        <h2>🏭 {title}</h2>
        <div class="date">📅 Ημερομηνία: <b>{date_str}</b> | Σύνολο: {len(df_data)}</div>
        <table>
            <thead><tr>{"".join([f"<th>{col}</th>" for col in df_data.columns])}</tr></thead>
            <tbody>
    """
    for _, row in df_data.iterrows():
        html += "<tr>"
        for col in df_data.columns:
            val = str(row[col])
            if val == "ΝΑΙ":
                val_str = '<span class="done">✅ Ολοκληρώθηκε</span>'
            elif val == "ΟΧΙ":
                val_str = '<span class="pending">⏳ Εκκρεμεί</span>'
            else:
                val_str = val
            html += f"<td>{val_str}</td>"
        html += "</tr>"
    html += """
            </tbody>
        </table>
        <br><button onclick="window.print()" style="padding:10px 20px;background:#1e88e5;color:white;border:none;border-radius:5px;cursor:pointer;">🖨️ Εκτύπωση</button>
    </body>
    </html>
    """
    return html

def export_to_excel(df, title):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Report', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Report']
        header_format = workbook.add_format({'bold': True, 'text_wrap': True, 'valign': 'top', 'fg_color': '#1e88e5', 'font_color': 'white', 'border': 1})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        for i, col in enumerate(df.columns):
            max_width = max(df[col].astype(str).str.len().max(), len(col)) + 2
            worksheet.set_column(i, i, min(max_width, 50))
    return output.getvalue()

# --- NOTIFICATIONS ---
def check_notifications():
    notifications = []
    if 'procurement_df' in st.session_state and not st.session_state.procurement_df.empty:
        pending = st.session_state.procurement_df[
            ~st.session_state.procurement_df["Status Procurement"].isin(["OK STOCK", "RECEIVED", "READY"])
        ]
        if not pending.empty:
            notifications.append(f"⚠️ {len(pending)} υλικά σε εκκρεμότητα procurement")
    return notifications

# --- TOGGLE FUNCTIONS ---
def toggle_project_task(p_key, task_name, chk_key):
    st.session_state["project_tasks_store"][p_key][task_name]["done"] = st.session_state[chk_key]
    save_all_assignments_to_sheet()

def toggle_item_task(u_key, t_idx, chk_key):
    st.session_state["tasks_store"][u_key][t_idx]["done"] = st.session_state[chk_key]
    save_all_assignments_to_sheet()

def update_item_field(u_key, t_idx, field, widget_key):
    st.session_state["tasks_store"][u_key][t_idx][field] = st.session_state[widget_key]
    save_all_assignments_to_sheet()

def update_proj_field(p_key, task_name, field, widget_key):
    st.session_state["project_tasks_store"][p_key][task_name][field] = st.session_state[widget_key]
    save_all_assignments_to_sheet()

# --- RENDER FUNCTIONS ---
def render_dashboard(procurement_df, tasks_database, team_database, availability_database):
    st.header("📈 Dashboard & Επισκόπηση Παραγωγής")
    
    if procurement_df.empty:
        st.warning("⚠️ No procurement data available.")
        return
    
    projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
    dashboard_data = []
    tot_all_hours = 0.0
    tot_done_hours = 0.0
    tot_tasks_count = 0
    tot_done_tasks = 0
    project_hours = {}
    project_progress = {}
    
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
        project_hours[p_name] = p_total_hrs
        project_progress[p_name] = p_progress
        dashboard_data.append({
            "Project": p_name,
            "Υλικά": len(filtered_p),
            "Σύνολο Tasks": p_tasks_cnt,
            "Ολοκληρωμένα": p_done_cnt,
            "Συνολικές Ώρες": round(p_total_hrs, 1),
            "Υπολειπόμενες": round(p_total_hrs - p_done_hrs, 1),
            "Πρόοδος": f"{p_progress}%"
        })

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ενεργά Projects", len(projects_list))
    c2.metric("Συνολικές Ώρες", f"{round(tot_all_hours, 1)}h")
    overall_pct = int((tot_done_tasks / tot_tasks_count) * 100) if tot_tasks_count > 0 else 0
    c3.metric("Συνολική Πρόοδος", f"{overall_pct}%")
    c4.metric("Εκκρεμή Tasks", tot_tasks_count - tot_done_tasks)

    st.divider()
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📊 Ώρες ανά Project")
        if project_hours:
            chart_data = pd.DataFrame({"Project": list(project_hours.keys()), "Ώρες": list(project_hours.values())})
            st.bar_chart(chart_data, x="Project", y="Ώρες", use_container_width=True)
    
    with col2:
        st.subheader("📈 Πρόοδος ανά Project")
        if project_progress:
            chart_data = pd.DataFrame({"Project": list(project_progress.keys()), "Πρόοδος (%)": list(project_progress.values())})
            st.bar_chart(chart_data, x="Project", y="Πρόοδος (%)", use_container_width=True)

    st.divider()
    if dashboard_data:
        dash_df = pd.DataFrame(dashboard_data)
        st.dataframe(dash_df, use_container_width=True, hide_index=True)
        
        col_exp1, col_exp2 = st.columns(2)
        with col_exp1:
            csv = dash_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(label="📊 Εξαγωγή CSV", data=csv, file_name=f"Dashboard_{date.today().strftime('%Y-%m-%d')}.csv", mime="text/csv", use_container_width=True)
        with col_exp2:
            excel_data = export_to_excel(dash_df, "Dashboard")
            st.download_button(label="📄 Εξαγωγή Excel", data=excel_data, file_name=f"Dashboard_{date.today().strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

def render_project(procurement_df, tasks_database, team_database, availability_database):
    st.header("📋 Διαχείριση Παραγωγής & Αναθέσεις ανά Project")
    
    if procurement_df.empty:
        st.warning("⚠️ No procurement data available.")
        return
    
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
                if status not in ["OK STOCK", "RECEIVED", "READY"]:
                    st.warning(f"⚠️ Εκκρεμότητα Procurement: {status}")
                else:
                    st.success(f"✅ Υλικό Διαθέσιμο: {status}")

            with col_tasks:
                st.markdown("**⚙️ Εργασίες Προετοιμασίας Υλικού**")
                for t_idx, t_data in enumerate(list(item_tasks)):
                    c_check, c_task, c_user, c_date, c_time, c_del = st.columns([0.08, 0.32, 0.23, 0.18, 0.11, 0.08])
                    
                    chk_k = f"proj_chk_{unique_item_key}_{t_idx}"
                    is_done = c_check.checkbox("", value=t_data["done"], key=chk_k, on_change=toggle_item_task, args=(unique_item_key, t_idx, chk_k))
                    
                    task_idx = task_options.index(t_data["task"]) if t_data["task"] in task_options else 0
                    task_k = f"proj_task_{unique_item_key}_{t_idx}"
                    selected_task = c_task.selectbox("Εργασία", task_options, index=task_idx, key=task_k, label_visibility="collapsed", on_change=update_item_field, args=(unique_item_key, t_idx, "task", task_k))
                    
                    user_idx = team_options.index(t_data["user"]) if t_data["user"] in team_options else 0
                    user_k = f"proj_user_{unique_item_key}_{t_idx}"
                    assigned_user = c_user.selectbox("Ανάθεση", team_options, index=user_idx, key=user_k, label_visibility="collapsed", on_change=update_item_field, args=(unique_item_key, t_idx, "user", user_k))
                    
                    date_k = f"proj_date_{unique_item_key}_{t_idx}"
                    assign_date = c_date.date_input("Ημερομηνία", value=t_data["date"], format="DD/MM/YYYY", key=date_k, label_visibility="collapsed", on_change=update_item_field, args=(unique_item_key, t_idx, "date", date_k))

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
                        save_all_assignments_to_sheet()
                        st.rerun()

                col_btn1, col_btn2, _ = st.columns([0.35, 0.35, 0.3])
                if col_btn1.button("➕ Προσθήκη Εργασίας Υλικού", key=f"add_btn_{unique_item_key}"):
                    st.session_state["tasks_store"][unique_item_key].append({"done": False, "task": "- Επιλογή Εργασίας -", "user": "- Χωρίς Ανάθεση -", "date": date.today()})
                    save_all_assignments_to_sheet()
                    st.rerun()
                if len(item_tasks) > 0 and col_btn2.button("➖ Αφαίρεση Εργασίας", key=f"rem_btn_{unique_item_key}"):
                    st.session_state["tasks_store"][unique_item_key].pop()
                    save_all_assignments_to_sheet()
                    st.rerun()

    st.divider()
    st.markdown("### 🛠️ Γενικές Εργασίες Project")
    proj_key = f"proj_{selected_project}"
    if proj_key not in st.session_state["project_tasks_store"] or not isinstance(st.session_state["project_tasks_store"][proj_key], dict):
        st.session_state["project_tasks_store"][proj_key] = {t_name: {"active": False, "done": False, "user": "- Χωρίς Ανάθεση -", "date": date.today()} for t_name in FIXED_PROJECT_TASKS}
    
    proj_tasks_dict = st.session_state["project_tasks_store"][proj_key]
    with st.expander(f"📦 5 Σταθερές Γενικές Εργασίες για το Project: {selected_project}", expanded=True):
        for task_name in FIXED_PROJECT_TASKS:
            t_data = proj_tasks_dict.get(task_name, {"active": False, "done": False, "user": "- Χωρίς Ανάθεση -", "date": date.today()})
            c_active, c_name, c_done, c_user, c_date, c_time = st.columns([0.06, 0.30, 0.10, 0.22, 0.18, 0.14])
            pact_k = f"pact_{proj_key}_{task_name}"
            is_active = c_active.checkbox("", value=t_data["active"], key=pact_k, on_change=update_proj_field, args=(proj_key, task_name, "active", pact_k))
            c_name.markdown(f"**{task_name}**" if is_active else f"<span style='color:gray;'>{task_name}</span>", unsafe_allow_html=True)
            if is_active:
                pdone_k = f"proj_pdone_{proj_key}_{task_name}"
                is_done = c_done.checkbox("Done", value=t_data["done"], key=pdone_k, on_change=toggle_project_task, args=(proj_key, task_name, pdone_k))
            else:
                c_done.caption("—")
            user_idx = team_options.index(t_data["user"]) if t_data["user"] in team_options else 0
            puser_k = f"puser_{proj_key}_{task_name}"
            assigned_user = c_user.selectbox("Ανάθεση", team_options, index=user_idx, key=puser_k, label_visibility="collapsed", disabled=not is_active, on_change=update_proj_field, args=(proj_key, task_name, "user", puser_k))
            pdate_k = f"pdate_{proj_key}_{task_name}"
            assign_date = c_date.date_input("Ημερομηνία", value=t_data["date"], format="DD/MM/YYYY", key=pdate_k, label_visibility="collapsed", disabled=not is_active, on_change=update_proj_field, args=(proj_key, task_name, "date", pdate_k))
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
    m1.metric("Συνολικές Ώρες", f"{round(total_project_hours, 1)} Ώρες")
    m2.metric("Ώρες που Ολοκληρώθηκαν", f"{round(completed_project_hours, 1)} Ώρες")
    m3.metric("Υπολειπόμενες Ώρες", f"{remaining_hours} Ώρες", delta=f"-{remaining_hours}h" if remaining_hours > 0 else "Έτοιμο!")
    if st.button("💾 Αποθήκευση Αλλαγών στο Google Sheet", use_container_width=True):
        if save_all_assignments_to_sheet():
            st.success("✅ Όλες οι αναθέσεις αποθηκεύτηκαν επιτυχώς!")
        else:
            st.error("❌ Σφάλμα κατά την αποθήκευση")

def render_daily_plan(procurement_df, tasks_database, team_database, availability_database):
    st.header("🗓️ Συγκεντρωτικό Πλάνο Παραγωγής")
    col_d, col_fp, col_fu, col_fs = st.columns([1, 1, 1, 1])
    target_date = col_d.date_input("Ημερομηνία Πλάνου:", value=date.today(), format="DD/MM/YYYY")
    greek_day_name = WEEKDAYS_GREEK.get(target_date.weekday(), "Δευτέρα")
    st.caption(f"Ημέρα εβδομάδας: **{greek_day_name}**")

    daily_tasks_raw = []
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
                    daily_tasks_raw.append({"type": "project", "p_key": p_key, "task_name": task_name, "Project": proj_name, "Υλικό": "Γενική Σύνθεση / Box", "Ποσότητα": proj_qty, "Εργασία": task_name, "Υπεύθυνος": t_user, "Ώρες": round(hours, 2), "done": t_done, "status_proc": "READY"})

    if not procurement_df.empty:
        for idx, row in procurement_df.iterrows():
            item_id = str(row["ID"])
            unique_item_key = f"{item_id}_{idx}"
            project_name = row["Project"]
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            proc_status = row["Status Procurement"]
            item_tasks = st.session_state["tasks_store"].get(unique_item_key, [])
            for t_idx, t_data in enumerate(item_tasks):
                t_task = t_data["task"]
                t_user = t_data["user"]
                t_date = t_data["date"]
                t_done = t_data["done"]
                if t_task != "- Επιλογή Εργασίας -" and t_date == target_date:
                    auto_time = tasks_database.get(t_task, 0.0)
                    hours = (auto_time * qty) / 60
                    daily_tasks_raw.append({"type": "item", "u_key": unique_item_key, "t_idx": t_idx, "Project": project_name, "Υλικό": material, "Ποσότητα": qty, "Εργασία": t_task, "Υπεύθυνος": t_user, "Ώρες": round(hours, 2), "done": t_done, "status_proc": proc_status})

    available_projects = ["Όλα τα Projects"] + sorted(list(set(d["Project"] for d in daily_tasks_raw))) if daily_tasks_raw else ["Όλα τα Projects"]
    available_users = ["Όλοι οι Τεχνίτες"] + sorted(list(set(d["Υπεύθυνος"] for d in daily_tasks_raw))) if daily_tasks_raw else ["Όλοι οι Τεχνίτες"]
    available_statuses = ["Όλα τα Status"] + sorted(list(set(d["status_proc"] for d in daily_tasks_raw))) if daily_tasks_raw else ["Όλα τα Status"]

    selected_filter_proj = col_fp.selectbox("🔍 Φίλτρο Project:", available_projects)
    selected_filter_user = col_fu.selectbox("👤 Φίλτρο Τεχνίτη:", available_users)
    selected_filter_status = col_fs.selectbox("📦 Φίλτρο Procurement:", available_statuses)

    daily_tasks = [d for d in daily_tasks_raw if (selected_filter_proj == "Όλα τα Projects" or d["Project"] == selected_filter_proj) and (selected_filter_user == "Όλοι οι Τεχνίτες" or d["Υπεύθυνος"] == selected_filter_user) and (selected_filter_status == "Όλα τα Status" or d["status_proc"] == selected_filter_status)]

    st.divider()
    if daily_tasks:
        export_list = [{"Project": dt["Project"], "Υλικό / Είδος": dt["Υλικό"], "Ποσότητα": dt["Ποσότητα"], "Εργασία": dt["Εργασία"], "Υπεύθυνος": dt["Υπεύθυνος"], "Ώρες": dt["Ώρες"], "Status Procurement": dt["status_proc"], "Ολοκληρώθηκε": "ΝΑΙ" if dt["done"] else "ΟΧΙ"} for dt in daily_tasks]
        export_df = pd.DataFrame(export_list)
        csv_data = export_df.to_csv(index=False).encode('utf-8-sig')
        col_head, col_exp_csv, col_exp_pdf, col_exp_excel = st.columns([0.4, 0.2, 0.2, 0.2])
        col_head.subheader(f"📌 Εργασίες για τις {target_date.strftime('%d/%m/%Y')} ({len(daily_tasks)} Tasks)")
        col_exp_csv.download_button(label="📊 CSV", data=csv_data, file_name=f"Daily_Plan_{target_date.strftime('%Y-%m-%d')}.csv", mime="text/csv", use_container_width=True)
        printable_html = generate_printable_html("Ημερήσιο Πλάνο Παραγωγής", target_date.strftime('%d/%m/%Y'), export_df)
        col_exp_pdf.download_button(label="📄 PDF", data=printable_html, file_name=f"Daily_Plan_{target_date.strftime('%Y-%m-%d')}.html", mime="text/html", use_container_width=True)
        excel_data = export_to_excel(export_df, "Daily Plan")
        col_exp_excel.download_button(label="📊 Excel", data=excel_data, file_name=f"Daily_Plan_{target_date.strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        pending_proc = [dt for dt in daily_tasks if dt["status_proc"] not in ["OK STOCK", "RECEIVED", "READY"]]
        if pending_proc:
            st.warning(f"⚠️ **Προσοχή:** Υπάρχουν **{len(pending_proc)} tasks** με υλικά σε εκκρεμότητα!")

        st.markdown("#### 👥 Φόρτος Εργασίας & Διαθεσιμότητα Ομάδας")
        day_availability = availability_database.get(greek_day_name, {})
        user_hours = {}
        for d in daily_tasks:
            u = d["Υπεύθυνος"]
            user_hours[u] = user_hours.get(u, 0.0) + d["Ώρες"]
        cols = st.columns(max(len(user_hours), 1))
        for i, (member_name, assigned_hrs) in enumerate(user_hours.items()):
            assigned_hrs = round(assigned_hrs, 2)
            if member_name != "- Χωρίς Ανάθεση -":
                max_hrs = day_availability.get(member_name, 6.0)
                delta_hrs = round(assigned_hrs - max_hrs, 2)
                if delta_hrs > 0:
                    cols[i].metric(f"⚠️ {member_name}", f"{assigned_hrs} / {max_hrs}h", delta=f"+{delta_hrs}h Υπερκάλυψη", delta_color="inverse")
                else:
                    cols[i].metric(f"🟢 {member_name}", f"{assigned_hrs} / {max_hrs}h", delta=f"{delta_hrs}h Διαθέσιμο", delta_color="normal")
            else:
                cols[i].metric(f"❓ {member_name}", f"{assigned_hrs} Ώρες")
            
        st.divider()
        st.markdown("#### 📋 Διαδραστική Λίστα Εργασιών")
        for d_idx, dt in enumerate(daily_tasks):
            col_chk, col_p, col_mat, col_tsk, col_user, col_hrs, col_st = st.columns([0.08, 0.20, 0.26, 0.20, 0.14, 0.07, 0.12])
            if dt["type"] == "project":
                chk_k = f"plan_pdone_{dt['p_key']}_{dt['task_name']}_{d_idx}"
                is_done = col_chk.checkbox("Done", value=dt["done"], key=chk_k, on_change=toggle_project_task, args=(dt['p_key'], dt['task_name'], chk_k))
            else:
                chk_k = f"plan_idone_{dt['u_key']}_{dt['t_idx']}_{d_idx}"
                is_done = col_chk.checkbox("Done", value=dt["done"], key=chk_k, on_change=toggle_item_task, args=(dt['u_key'], dt['t_idx'], chk_k))
            col_p.markdown(f"**{dt['Project']}**")
            col_mat.caption(f"{dt['Υλικό']} ({dt['Ποσότητα']} τμχ)")
            col_tsk.markdown(f"~~{dt['Εργασία']}~~ ✅" if is_done else f"**{dt['Εργασία']}**")
            col_user.write(dt['Υπεύθυνος'])
            col_hrs.write(f"{dt['Ώρες']}h")
            col_st.success(f"✅ {dt['status_proc']}") if dt['status_proc'] in ["OK STOCK", "RECEIVED", "READY"] else col_st.error(f"⚠️ {dt['status_proc']}")
    else:
        st.info(f"Δεν βρέθηκαν εργασίες για τις {target_date.strftime('%d/%m/%Y')} με τα συγκεκριμένα φίλτρα.")

def render_technician(procurement_df, tasks_database, team_database, availability_database):
    st.header("👤 Ημερήσιο Πρόγραμμα Εργασιών ανά Τεχνίτη")
    c_date, c_user = st.columns([1, 1])
    target_date = c_date.date_input("Ημερομηνία:", value=date.today(), format="DD/MM/YYYY", key="tech_date")
    selected_member = c_user.selectbox("Επιλέξτε Τεχνίτη:", team_database)
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
                if isinstance(p_data, dict) and p_data.get("active", False) and p_data.get("user") == selected_member and p_data.get("date") == target_date:
                    auto_time = tasks_database.get(task_name, 0.0)
                    hours = (auto_time * proj_qty) / 60
                    worker_tasks.append({"type": "project", "p_key": p_key, "task_name": task_name, "project": proj_name, "item": "Γενική Σύνθεση / Box", "qty": proj_qty, "task": task_name, "hours": round(hours, 2), "done": p_data.get("done", False), "status_proc": "READY"})

    if not procurement_df.empty:
        for idx, row in procurement_df.iterrows():
            item_id = str(row["ID"])
            unique_item_key = f"{item_id}_{idx}"
            project_name = row["Project"]
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            proc_status = row["Status Procurement"]
            item_tasks = st.session_state["tasks_store"].get(unique_item_key, [])
            for t_idx, t_data in enumerate(item_tasks):
                if t_data.get("user") == selected_member and t_data.get("date") == target_date:
                    t_task = t_data.get("task")
                    if t_task != "- Επιλογή Εργασίας -":
                        auto_time = tasks_database.get(t_task, 0.0)
                        hours = (auto_time * qty) / 60
                        worker_tasks.append({"type": "item", "u_key": unique_item_key, "t_idx": t_idx, "project": project_name, "item": f"[{item_id}] {material}", "qty": qty, "task": t_task, "hours": round(hours, 2), "done": t_data.get("done", False), "status_proc": proc_status})

    if worker_tasks:
        w_export_list = [{"Project": wt["project"], "Υλικό / Είδος": wt["item"], "Ποσότητα": wt["qty"], "Εργασία": wt["task"], "Ώρες": wt["hours"], "Status Procurement": wt["status_proc"], "Ολοκληρώθηκε": "ΝΑΙ" if wt["done"] else "ΟΧΙ"} for wt in worker_tasks]
        w_export_df = pd.DataFrame(w_export_list)
        w_csv_data = w_export_df.to_csv(index=False).encode('utf-8-sig')
        col_w_head, col_w_csv, col_w_pdf, col_w_excel = st.columns([0.4, 0.2, 0.2, 0.2])
        col_w_head.subheader(f"📋 Πρόγραμμα για τον/την {selected_member} — {target_date.strftime('%d/%m/%Y')}")
        col_w_csv.download_button(label="📊 CSV", data=w_csv_data, file_name=f"Schedule_{selected_member.replace(' ', '_')}_{target_date.strftime('%Y-%m-%d')}.csv", mime="text/csv", use_container_width=True)
        w_printable_html = generate_printable_html(f"Πρόγραμμα Τεχνίτη: {selected_member}", target_date.strftime('%d/%m/%Y'), w_export_df)
        col_w_pdf.download_button(label="📄 PDF", data=w_printable_html, file_name=f"Schedule_{selected_member.replace(' ', '_')}_{target_date.strftime('%Y-%m-%d')}.html", mime="text/html", use_container_width=True)
        excel_data = export_to_excel(w_export_df, f"Schedule {selected_member}")
        col_w_excel.download_button(label="📊 Excel", data=excel_data, file_name=f"Schedule_{selected_member.replace(' ', '_')}_{target_date.strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        total_w_hours = sum(t["hours"] for t in worker_tasks)
        st.info(f"💡 Συνολικός εκτιμώμενος χρόνος: **{round(total_w_hours, 1)} Ώρες** ({len(worker_tasks)} Tasks)")
        pending_w_proc = [wt for wt in worker_tasks if wt["status_proc"] not in ["OK STOCK", "RECEIVED", "READY"]]
        if pending_w_proc:
            st.warning(f"⚠️ Ο/Η {selected_member} έχει **{len(pending_w_proc)} tasks** με υλικά σε εκκρεμότητα.")
    else:
        st.success(f"🎉 Δεν έχουν ανατεθεί εργασίες στον/στην {selected_member} για τις {target_date.strftime('%d/%m/%Y')}.")

def render_projection(procurement_df, tasks_database, team_database, availability_database):
    st.header("📆 Πρόβλεψη Φόρτου Εργασίας (Projection)")
    start_monday = date.today() - timedelta(days=date.today().weekday())
    col_w_choice, col_range_choice, col_weekend = st.columns([1, 1, 1])
    week_choice = col_w_choice.selectbox("ΕΝΑΡΞΗ ΠΡΟΒΟΛΗΣ:", ["Τρέχουσα Εβδομάδα", "Επόμενη Εβδομάδα (+1)", "Μεθεπόμενη Εβδομάδα (+2)", "Προσαρμοσμένη Ημερομηνία"])
    range_weeks = col_range_choice.selectbox("ΕΥΡΟΣ ΠΡΟΒΟΛΗΣ:", ["1 Εβδομάδα", "2 Εβδομάδες", "4 Εβδομάδες / Μήνας"])
    include_weekends = col_weekend.checkbox("📅 Συμπερίληψη Σαββατοκύριακων", value=False, key="proj_wknd")
    num_weeks = 1 if "1 Εβδομάδα" in range_weeks else 2 if "2 Εβδομάδες" in range_weeks else 4
    if week_choice == "Τρέχουσα Εβδομάδα":
        sel_start = start_monday
    elif week_choice == "Επόμενη Εβδομάδα (+1)":
        sel_start = start_monday + timedelta(days=7)
    elif week_choice == "Μεθεπόμενη Εβδομάδα (+2)":
        sel_start = start_monday + timedelta(days=14)
    else:
        sel_start = st.date_input("Επιλέξτε Δευτέρα Εναρξης:", value=start_monday, format="DD/MM/YYYY", key="proj_start")
    days_per_week = 7 if include_weekends else 5
    weeks_days_list = []
    all_flat_days = []
    for w in range(num_weeks):
        w_monday = sel_start + timedelta(days=w*7)
        w_days = [w_monday + timedelta(days=i) for i in range(days_per_week)]
        weeks_days_list.append((w+1, w_monday, w_days))
        all_flat_days.extend(w_days)

    total_assigned_range = 0.0
    total_available_range = 0.0
    overbooked_days_count = 0
    for d in all_flat_days:
        g_day = WEEKDAYS_GREEK.get(d.weekday(), "Δευτέρα")
        day_avail = availability_database.get(g_day, {})
        day_max = sum(day_avail.get(m, 6.0) for m in team_database)
        total_available_range += day_max

    weeks_matrices = []
    for w_num, w_monday, w_days in weeks_days_list:
        w_matrix = {m: {f"{WEEKDAYS_SHORT_GREEK[d.weekday()]} {d.strftime('%d/%m')}": 0.0 for d in w_days} for m in team_database}
        d_totals = {f"{WEEKDAYS_SHORT_GREEK[d.weekday()]} {d.strftime('%d/%m')}": 0.0 for d in w_days}
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
    kc3.metric("Overbooked Ημέρες", f"⚠️ {overbooked_days_count} / {len(all_flat_days)}" if overbooked_days_count > 0 else f"🟢 0 / {len(all_flat_days)}")
    kc4.metric("Πληρότητα Περιόδου", f"{load_ratio}%", delta=f"{load_ratio - 100}%" if load_ratio > 100 else "Εντός Ορίων")

    st.divider()
    def highlight_total_row(row):
        if row.name == "Σύνολο Ημέρας (h)":
            return ["background-color: #2b303a; font-weight: bold; color: #00e676; border-top: 2px solid #00e676;"] * len(row)
        return [""] * len(row)

    for w_num, w_monday, w_days, w_matrix in weeks_matrices:
        w_sunday = w_days[-1]
        st.subheader(f"📅 Εβδομάδα {w_num}: {w_monday.strftime('%d/%m/%Y')} έως {w_sunday.strftime('%d/%m/%Y')}")
        proj_df = pd.DataFrame(w_matrix).T
        proj_df = proj_df.round(1)
        proj_df["Σύνολο (h)"] = proj_df.sum(axis=1)
        total_row = proj_df.sum(axis=0).round(1)
        total_row.name = "Σύνολο Ημέρας (h)"
        proj_df = pd.concat([proj_df, pd.DataFrame(total_row).T])
        styled_df = proj_df.style.apply(highlight_total_row, axis=1)
        st.dataframe(styled_df, use_container_width=True)

def render_daily_report(procurement_df, tasks_database, team_database, availability_database):
    st.header("📝 Ημερήσιος Απολογισμός Παραγωγής")
    rep_date = st.date_input("Επιλέξτε Ημερομηνία:", value=date.today(), format="DD/MM/YYYY", key="rep_date_input")
    st.divider()
    rep_completed = []
    rep_pending = []

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
                if isinstance(p_data, dict) and p_data.get("active", False) and p_data.get("date") == rep_date:
                    auto_time = tasks_database.get(task_name, 0.0)
                    hrs = round((auto_time * proj_qty) / 60, 2)
                    item_info = {"Project": proj_name, "Εργασία": task_name, "Υλικό / Είδος": "Γενική Σύνθεση / Box", "Υπεύθυνος": p_data.get("user", "-"), "Ώρες": hrs}
                    if p_data.get("done", False):
                        rep_completed.append(item_info)
                    else:
                        rep_pending.append(item_info)

    if not procurement_df.empty:
        for idx, row in procurement_df.iterrows():
            item_id = str(row["ID"])
            unique_item_key = f"{item_id}_{idx}"
            project_name = row["Project"]
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            item_tasks = st.session_state["tasks_store"].get(unique_item_key, [])
            for t_data in item_tasks:
                if t_data.get("task") != "- Επιλογή Εργασίας -" and t_data.get("date") == rep_date:
                    auto_time = tasks_database.get(t_data["task"], 0.0)
                    hrs = round((auto_time * qty) / 60, 2)
                    item_info = {"Project": project_name, "Εργασία": t_data["task"], "Υλικό / Είδος": f"[{item_id}] {material}", "Υπεύθυνος": t_data.get("user", "-"), "Ώρες": hrs}
                    if t_data.get("done", False):
                        rep_completed.append(item_info)
                    else:
                        rep_pending.append(item_info)

    rc1, rc2, rc3 = st.columns(3)
    tot_done_hrs = sum(x["Ώρες"] for x in rep_completed)
    tot_pend_hrs = sum(x["Ώρες"] for x in rep_pending)
    rc1.metric("Ολοκληρωμένα Tasks", len(rep_completed), delta=f"{round(tot_done_hrs, 1)}h")
    rc2.metric("Εκκρεμή Tasks", len(rep_pending), delta=f"-{round(tot_pend_hrs, 1)}h", delta_color="inverse")
    completion_rate = int((len(rep_completed) / (len(rep_completed) + len(rep_pending))) * 100) if (len(rep_completed) + len(rep_pending)) > 0 else 100
    rc3.metric("Ποσοστό Ολοκλήρωσης", f"{completion_rate}%")

    st.divider()
    st.subheader("✅ Ολοκληρωμένες Εργασίες")
    if rep_completed:
        st.dataframe(pd.DataFrame(rep_completed), use_container_width=True, hide_index=True)
    else:
        st.info("Δεν υπάρχουν ολοκληρωμένες εργασίες για αυτή την ημερομηνία.")

    st.divider()
    st.subheader("⏳ Εκκρεμότητες")
    if rep_pending:
        st.dataframe(pd.DataFrame(rep_pending), use_container_width=True, hide_index=True)
    else:
        st.success("🎉 Όλες οι εργασίες έχουν ολοκληρωθεί!")

def render_database(tasks_database, team_database, availability_database):
    st.header("📊 Βάση Δεδομένων")
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f"⏱️ Πρότυπα Χρόνων ({len(tasks_database)} Εργασίες)")
        tasks_df = pd.DataFrame(list(tasks_database.items()), columns=["Εργασία", "Χρόνος (λεπτά)"])
        st.dataframe(tasks_df, use_container_width=True, hide_index=True)
    with col_b:
        st.subheader("👥 Ομάδα & Όρια Ώρων")
        avail_data = [{"Ημέρα": day, "Τεχνίτης": member, "Ώρες": hours} for day, members in availability_database.items() for member, hours in members.items()]
        st.dataframe(pd.DataFrame(avail_data), use_container_width=True, hide_index=True)

def render_settings():
    st.header("⚙️ Ρυθμίσεις")
    st.subheader("🔐 Ασφάλεια")
    st.info("🔒 Οι ρυθμίσεις ασφαλείας διαχειρίζονται μέσω των Streamlit Secrets")
    st.markdown("**Users:** admin/admin123, manager/manager123, operator/operator123")
    
    st.subheader("💾 Αποθήκευση")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("💾 Αποθήκευση", use_container_width=True):
            if save_all_assignments_to_sheet():
                st.success("✅ Αποθηκεύτηκε!")
    with col2:
        if st.button("🔄 Επαναφόρτωση", use_container_width=True):
            st.cache_data.clear()
            st.session_state.data_version = st.session_state.get("data_version", 0) + 1
            st.rerun()
    
    st.subheader("📋 Audit Log")
    if "audit_log" in st.session_state and st.session_state.audit_log:
        st.dataframe(pd.DataFrame(st.session_state.audit_log[-50:]), use_container_width=True, hide_index=True)
    else:
        st.info("Δεν υπάρχουν καταχωρήσεις.")

# --- MAIN ---
def main():
    init_auth()
    if not st.session_state.authenticated:
        login_form()
        return
    
    version = st.session_state.get("data_version", 0)
    procurement_df, tasks_database, team_database, availability_database = load_all_data(version)
    st.session_state.procurement_df = procurement_df
    st.session_state.availability_database = availability_database
    
    sheet_item_assignments, sheet_proj_assignments = load_assignments_from_sheet()
    
    if "tasks_store" not in st.session_state:
        st.session_state["tasks_store"] = {}
    if procurement_df is not None and not procurement_df.empty:
        for idx, row in procurement_df.iterrows():
            item_id = str(row["ID"])
            u_key = f"{item_id}_{idx}"
            if u_key not in st.session_state["tasks_store"]:
                st.session_state["tasks_store"][u_key] = sheet_item_assignments.get(item_id, [])
    
    if "project_tasks_store" not in st.session_state:
        st.session_state["project_tasks_store"] = sheet_proj_assignments
    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []
    if "last_save" not in st.session_state:
        st.session_state.last_save = datetime.now()

    # Sidebar
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/000000/factory.png", width=80)
        st.markdown(f"### 🏭 Production Tasks")
        st.markdown(f"👋 Welcome, **{st.session_state.username}**!")
        
        notifications = check_notifications()
        if notifications:
            with st.expander(f"🔔 Notifications ({len(notifications)})", expanded=True):
                for notif in notifications:
                    st.warning(notif)
        st.divider()
        
        selected = option_menu(
            menu_title="Navigation",
            options=["📈 Dashboard", "📋 Project", "🗓️ Daily Plan", "👤 Technician", "📆 Projection", "📝 Daily Report", "📊 Database", "⚙️ Settings"],
            icons=["bar-chart", "list-task", "calendar", "person", "graph-up", "clipboard", "database", "gear"],
            menu_icon="menu-button",
            default_index=0,
            styles={"container": {"padding": "0!important"}, "icon": {"font-size": "20px"}, "nav-link": {"font-size": "15px", "text-align": "left", "margin": "0px"}, "nav-link-selected": {"background-color": "#1e88e5"}}
        )
        st.divider()
        total_tasks = sum(len(tasks) for tasks in st.session_state.get("tasks_store", {}).values())
        st.metric("Total Tasks", total_tasks)
        st.metric("Active Projects", len(procurement_df["Project"].unique()) if not procurement_df.empty else 0)
        st.divider()
        if st.button("🚪 Logout", use_container_width=True):
            logout()

    # Auto-save
    if (datetime.now() - st.session_state.last_save).seconds > 300:
        if save_all_assignments_to_sheet():
            st.session_state.last_save = datetime.now()

    # Tabs
    if selected == "📈 Dashboard":
        render_dashboard(procurement_df, tasks_database, team_database, availability_database)
    elif selected == "📋 Project":
        render_project(procurement_df, tasks_database, team_database, availability_database)
    elif selected == "🗓️ Daily Plan":
        render_daily_plan(procurement_df, tasks_database, team_database, availability_database)
    elif selected == "👤 Technician":
        render_technician(procurement_df, tasks_database, team_database, availability_database)
    elif selected == "📆 Projection":
        render_projection(procurement_df, tasks_database, team_database, availability_database)
    elif selected == "📝 Daily Report":
        render_daily_report(procurement_df, tasks_database, team_database, availability_database)
    elif selected == "📊 Database":
        render_database(tasks_database, team_database, availability_database)
    elif selected == "⚙️ Settings":
        render_settings()

if __name__ == "__main__":
    main()
