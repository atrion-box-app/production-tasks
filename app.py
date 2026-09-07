import streamlit as st
import pandas as pd
import gspread
import json
import base64
import hashlib
import time
from google.oauth2.service_account import Credentials
from datetime import date, datetime, timedelta
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
    if "page" not in st.session_state:
        st.session_state.page = "Dashboard"

def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    valid_users = {
        "admin": hash_password("admin123"),
        "manager": hash_password("manager123"),
        "operator": hash_password("operator123"),
        "maria@atrionartgifts.com": hash_password("atrionmaria")
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

# --- IMPORTANT: Add the GID for "Incoming Projects_List" sheet ---
# YOU NEED TO FIND THIS GID FROM YOUR SHEET URL
INCOMING_GID = "1362920506"  # <--- CHANGE THIS TO YOUR ACTUAL GID!

TIMES_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={TIMES_GID}"
TEAM_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={TEAM_GID}"
INCOMING_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={INCOMING_GID}"

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
        # Load procurement materials
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

    # --- NEW: Load Incoming Projects List ---
    df_incoming = pd.DataFrame()
    try:
        df_incoming_raw = pd.read_csv(INCOMING_CSV_URL)
        # Try to find the shipping status column
        # Based on your screenshot, column L is "Shipping Status"
        # We need to find which column that is in the CSV
        if not df_incoming_raw.empty:
            # Look for columns that might contain "Shipping" or "Status"
            shipping_col = None
            for col in df_incoming_raw.columns:
                if "shipping" in str(col).lower() or "status" in str(col).lower():
                    shipping_col = col
                    break
            
            # If we found a shipping status column, use it
            if shipping_col:
                # Keep only Project Name and Shipping Status
                # Project name is usually in column B (index 1)
                project_col = df_incoming_raw.columns[1] if len(df_incoming_raw.columns) > 1 else df_incoming_raw.columns[0]
                df_incoming = df_incoming_raw[[project_col, shipping_col]].copy()
                df_incoming.columns = ["Project", "Shipping Status"]
                df_incoming = df_incoming.dropna(subset=["Project"])
                df_incoming["Shipping Status"] = df_incoming["Shipping Status"].fillna("").str.strip()
    except Exception as e:
        st.warning(f"⚠️ Could not load Incoming Projects List: {e}")

    # Load tasks database
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

    # Load team and availability
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

    return df_proc, tasks_dict, team_members, availability_dict, df_incoming

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

# --- PROJECT CARDS FUNCTIONS ---
def get_project_details(project_name, procurement_df, tasks_database, incoming_df):
    items = procurement_df[procurement_df["Project"] == project_name].copy() if not procurement_df.empty else pd.DataFrame()
    
    total_hours = 0
    completed_hours = 0
    total_tasks = 0
    completed_tasks = 0
    materials_count = len(items)
    
    p_key = f"proj_{project_name}"
    project_tasks = st.session_state.get("project_tasks_store", {}).get(p_key, {})
    
    for idx, row in items.iterrows():
        item_id = str(row["ID"])
        u_key = f"{item_id}_{idx}"
        qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
        
        item_tasks = st.session_state.get("tasks_store", {}).get(u_key, [])
        for t in item_tasks:
            if t.get("task") and t["task"] != "- Επιλογή Εργασίας -":
                auto_t = tasks_database.get(t["task"], 0.0)
                hrs = (auto_t * qty) / 60
                total_hours += hrs
                total_tasks += 1
                if t.get("done", False):
                    completed_hours += hrs
                    completed_tasks += 1
    
    main_qty = 1
    if not items.empty:
        for _, r in items.iterrows():
            if str(r["Ποσότητα"]).isdigit():
                main_qty = max(main_qty, int(r["Ποσότητα"]))
    
    if isinstance(project_tasks, dict):
        for task_name, p_data in project_tasks.items():
            if isinstance(p_data, dict) and p_data.get("active", False):
                auto_t = tasks_database.get(task_name, 0.0)
                hrs = (auto_t * main_qty) / 60
                total_hours += hrs
                total_tasks += 1
                if p_data.get("done", False):
                    completed_hours += hrs
                    completed_tasks += 1
    
    progress = int((completed_tasks / total_tasks) * 100) if total_tasks > 0 else 0
    
    # --- CHECK IF PROJECT IS SHIPPED (from Incoming Projects List) ---
    is_shipped = False
    if not incoming_df.empty:
        # Find the project in incoming list
        project_row = incoming_df[incoming_df["Project"].str.strip().str.upper() == project_name.strip().upper()]
        if not project_row.empty:
            shipping_status = str(project_row.iloc[0]["Shipping Status"]).strip().upper()
            if shipping_status == "OK SHIPPED":
                is_shipped = True
    
    # Project is active if:
    # 1. It has at least one material
    # 2. AND it's NOT shipped
    is_active = materials_count > 0 and not is_shipped
    
    if progress == 100 and total_tasks > 0:
        status = "Ολοκληρώθηκε"
        status_class = "completed"
    elif progress > 0:
        status = "Σε Εξέλιξη"
        status_class = "in-progress"
    else:
        status = "Αναμονή"
        status_class = "pending"
    
    materials_list = []
    for _, item in items.iterrows():
        materials_list.append({
            "id": item['ID'],
            "name": item['Υλικό / Προϊόν'],
            "qty": item['Ποσότητα'],
            "status": item['Status Procurement']
        })
    
    return {
        "name": project_name,
        "items": items,
        "materials_list": materials_list,
        "total_hours": round(total_hours, 1),
        "completed_hours": round(completed_hours, 1),
        "remaining_hours": round(total_hours - completed_hours, 1),
        "total_tasks": total_tasks,
        "completed_tasks": completed_tasks,
        "progress": progress,
        "materials_count": materials_count,
        "status": status,
        "status_class": status_class,
        "is_active": is_active,
        "is_shipped": is_shipped,
        "project_tasks": project_tasks
    }

# --- RENDER FUNCTIONS ---
def render_dashboard(procurement_df, tasks_database, team_database, availability_database, incoming_df):
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

def render_project_cards(procurement_df, tasks_database, team_database, availability_database, incoming_df):
    """Render projects as expandable cards with active filter"""
    st.header("📇 Project Cards")
    
    if procurement_df.empty:
        st.warning("⚠️ No projects available.")
        return
    
    projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
    
    # --- FILTERS SECTION ---
    col_search, col_filter, col_active = st.columns([2, 1, 1])
    with col_search:
        search_term = st.text_input("🔍 Αναζήτηση Project:", placeholder="Πληκτρολόγησε το όνομα του project...")
    with col_filter:
        status_filter = st.selectbox("📌 Φίλτρο Κατάστασης:", ["Όλα", "Σε Εξέλιξη", "Ολοκληρώθηκε", "Αναμονή"])
    with col_active:
        show_active_only = st.checkbox("✅ Μόνο Ενεργά Projects", value=True, help="Εμφάνιση μόνο projects που ΔΕΝ έχουν φύγει (Shipping Status ≠ OK Shipped)")
    
    # Get all project details with active status
    all_projects = []
    for p_name in projects_list:
        if search_term and search_term.lower() not in p_name.lower():
            continue
        
        proj_data = get_project_details(p_name, procurement_df, tasks_database, incoming_df)
        
        if status_filter != "Όλα" and proj_data["status"] != status_filter:
            continue
        
        # Filter active projects only
        if show_active_only and not proj_data['is_active']:
            continue
        
        all_projects.append(proj_data)
    
    if not all_projects:
        if show_active_only:
            st.info("📌 Δεν υπάρχουν ενεργά projects. Απενεργοποίησε το 'Μόνο Ενεργά Projects' για να δεις όλα τα projects.")
        else:
            st.info("Δεν βρέθηκαν projects που να ταιριάζουν με τα κριτήρια αναζήτησης.")
        return
    
    # Show count of active vs total
    active_count = sum(1 for p in all_projects if p.get('is_active', False))
    total_projects = len(all_projects)
    st.caption(f"📊 Εμφανίζονται {total_projects} projects ({active_count} ενεργά)")
    
    # Display cards as expandable cards
    for i, proj_data in enumerate(all_projects):
        # Status emoji
        if proj_data['status'] == "Ολοκληρώθηκε":
            status_emoji = "✅"
        elif proj_data['status'] == "Σε Εξέλιξη":
            status_emoji = "🔄"
        else:
            status_emoji = "⏳"
        
        # Add active indicator
        if proj_data.get('is_shipped', False):
            active_indicator = "🔴 ΦΥΓΕ"
        elif proj_data.get('is_active', False):
            active_indicator = "🟢 ΕΝΕΡΓΟ"
        else:
            active_indicator = "⚪ ΑΝΕΝΕΡΓΟ"
        
        # Progress bar color based on status
        progress_color = "#2e7d32" if proj_data['progress'] == 100 else "#1e88e5" if proj_data['progress'] > 0 else "#ff9800"
        
        # Create expander
        with st.expander(
            f"{active_indicator} {status_emoji} 📦 {proj_data['name']}  |  {proj_data['progress']}%  |  {proj_data['status']}  |  {proj_data['total_tasks']} Tasks",
            expanded=False
        ):
            # Summary metrics in a row
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Συνολικές Ώρες", f"{proj_data['total_hours']}h")
            col2.metric("Ολοκληρωμένες Ώρες", f"{proj_data['completed_hours']}h")
            col3.metric("Υπολειπόμενες Ώρες", f"{proj_data['remaining_hours']}h")
            col4.metric("Πρόοδος", f"{proj_data['progress']}%")
            
            # Progress bar
            st.markdown(f"""
            <div style="background:#e0e0e0;border-radius:8px;height:20px;overflow:hidden;margin:10px 0;">
                <div style="background:{progress_color};height:100%;width:{proj_data['progress']}%;border-radius:8px;transition:width 0.5s ease;display:flex;align-items:center;justify-content:center;color:white;font-size:12px;font-weight:bold;">
                    {proj_data['progress']}%
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            st.divider()
            
            # --- TASKS PER MATERIAL ---
            st.subheader("⚙️ Tasks ανά Υλικό")
            
            # Get all items with their tasks
            items_with_tasks = []
            for idx, row in proj_data['items'].iterrows():
                item_id = str(row["ID"])
                u_key = f"{item_id}_{idx}"
                material = row["Υλικό / Προϊόν"]
                qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
                
                item_tasks = st.session_state.get("tasks_store", {}).get(u_key, [])
                
                if item_tasks:
                    for t in item_tasks:
                        if t.get("task") and t["task"] != "- Επιλογή Εργασίας -":
                            auto_time = tasks_database.get(t["task"], 0.0)
                            hrs = (auto_time * qty) / 60
                            items_with_tasks.append({
                                "ID": item_id,
                                "Υλικό": material[:40] + "..." if len(material) > 40 else material,
                                "Ποσότητα": qty,
                                "Εργασία": t["task"],
                                "Υπεύθυνος": t.get("user", "-"),
                                "Ημερομηνία": t.get("date", ""),
                                "Κατάσταση": "✅ Ολοκληρώθηκε" if t.get("done", False) else "⏳ Εκκρεμεί",
                                "Ώρες": round(hrs, 2)
                            })
            
            if items_with_tasks:
                tasks_df = pd.DataFrame(items_with_tasks)
                st.dataframe(tasks_df, use_container_width=True, hide_index=True)
            else:
                st.info("📌 Δεν έχουν οριστεί εργασίες για τα υλικά αυτού του project")
            
            st.divider()
            
            # Materials
            st.subheader(f"📋 Υλικά & Είδη ({proj_data['materials_count']})")
            if proj_data['materials_list']:
                materials_df = pd.DataFrame(proj_data['materials_list'])
                st.dataframe(materials_df, use_container_width=True, hide_index=True)
            else:
                st.info("Δεν βρέθηκαν υλικά")
            
            st.divider()
            
            # Project tasks
            st.subheader("🏗️ Γενικές Εργασίες Project")
            if isinstance(proj_data['project_tasks'], dict):
                tasks_data = []
                for task_name, p_data in proj_data['project_tasks'].items():
                    if isinstance(p_data, dict):
                        tasks_data.append({
                            "Εργασία": task_name,
                            "Κατάσταση": "✅ Ολοκληρώθηκε" if p_data.get("done", False) else "⏳ Εκκρεμεί",
                            "Υπεύθυνος": p_data.get("user", "-"),
                            "Ημερομηνία": p_data.get("date", ""),
                            "Ενεργό": "ΝΑΙ" if p_data.get("active", False) else "ΟΧΙ"
                        })
                if tasks_data:
                    tasks_df = pd.DataFrame(tasks_data)
                    st.dataframe(tasks_df, use_container_width=True, hide_index=True)
                else:
                    st.info("Δεν έχουν οριστεί γενικές εργασίες")
            else:
                st.info("Δεν έχουν οριστεί γενικές εργασίες")
            
            # Progress summary
            st.divider()
            st.caption(f"📊 Πρόοδος: {proj_data['completed_tasks']} από {proj_data['total_tasks']} tasks ολοκληρώθηκαν")
    
    # Export buttons
    st.divider()
    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        export_data = []
        for p in all_projects:
            export_data.append({
                "Project": p["name"],
                "Status": p["status"],
                "Ενεργό": "ΝΑΙ" if p.get('is_active', False) else "ΟΧΙ",
                "Έχει Φύγει": "ΝΑΙ" if p.get('is_shipped', False) else "ΟΧΙ",
                "Υλικά": p["materials_count"],
                "Σύνολο Tasks": p["total_tasks"],
                "Ολοκληρωμένα": p["completed_tasks"],
                "Πρόοδος": f"{p['progress']}%",
                "Συνολικές Ώρες": p["total_hours"],
                "Υπολειπόμενες Ώρες": p["remaining_hours"]
            })
        export_df = pd.DataFrame(export_data)
        csv = export_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(label="📊 Εξαγωγή Projects CSV", data=csv, file_name=f"Projects_{date.today().strftime('%Y-%m-%d')}.csv", mime="text/csv", use_container_width=True)
    with col_exp2:
        export_data = []
        for p in all_projects:
            export_data.append({
                "Project": p["name"],
                "Status": p["status"],
                "Ενεργό": "ΝΑΙ" if p.get('is_active', False) else "ΟΧΙ",
                "Έχει Φύγει": "ΝΑΙ" if p.get('is_shipped', False) else "ΟΧΙ",
                "Υλικά": p["materials_count"],
                "Σύνολο Tasks": p["total_tasks"],
                "Ολοκληρωμένα": p["completed_tasks"],
                "Πρόοδος": f"{p['progress']}%",
                "Συνολικές Ώρες": p["total_hours"],
                "Υπολειπόμενες Ώρες": p["remaining_hours"]
            })
        export_df = pd.DataFrame(export_data)
        excel_data = export_to_excel(export_df, "Projects")
        st.download_button(label="📄 Εξαγωγή Projects Excel", data=excel_data, file_name=f"Projects_{date.today().strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

# --- OTHER RENDER FUNCTIONS (keep the same as before) ---
# ... (rest of the code remains the same)

# --- MAIN ---
def main():
    init_auth()
    if not st.session_state.authenticated:
        login_form()
        return
    
    version = st.session_state.get("data_version", 0)
    procurement_df, tasks_database, team_database, availability_database, incoming_df = load_all_data(version)
    st.session_state.procurement_df = procurement_df
    st.session_state.availability_database = availability_database
    st.session_state.incoming_df = incoming_df
    
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
        
        # Simple navigation using st.radio
        st.markdown("### 📋 Navigation")
        page = st.radio(
            "Select Page",
            ["📈 Dashboard", "📇 Project Cards", "📋 Project", "🗓️ Daily Plan", "👤 Technician", "📆 Projection", "📝 Daily Report", "📊 Database", "⚙️ Settings"],
            index=0,
            label_visibility="collapsed"
        )
        st.session_state.page = page
        
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

    # Tabs - using session state page
    page = st.session_state.page
    
    if page == "📈 Dashboard":
        render_dashboard(procurement_df, tasks_database, team_database, availability_database, incoming_df)
    elif page == "📇 Project Cards":
        render_project_cards(procurement_df, tasks_database, team_database, availability_database, incoming_df)
    elif page == "📋 Project":
        render_project(procurement_df, tasks_database, team_database, availability_database)
    elif page == "🗓️ Daily Plan":
        render_daily_plan(procurement_df, tasks_database, team_database, availability_database)
    elif page == "👤 Technician":
        render_technician(procurement_df, tasks_database, team_database, availability_database)
    elif page == "📆 Projection":
        render_projection(procurement_df, tasks_database, team_database, availability_database)
    elif page == "📝 Daily Report":
        render_daily_report(procurement_df, tasks_database, team_database, availability_database)
    elif page == "📊 Database":
        render_database(tasks_database, team_database, availability_database)
    elif page == "⚙️ Settings":
        render_settings()

if __name__ == "__main__":
    main()
