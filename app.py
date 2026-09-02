import streamlit as st
import pandas as pd
import gspread
import json
import base64
import hashlib
import time
from google.oauth2.service_account import Credentials
from datetime import date, datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
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
    /* Mobile friendly */
    @media (max-width: 768px) {
        .stColumns { flex-direction: column !important; }
        .stButton button { width: 100% !important; }
        .stSelectbox, .stDateInput { margin-bottom: 10px; }
    }
    /* Tooltips */
    .tooltip {
        position: relative;
        display: inline-block;
        cursor: help;
    }
    .tooltip .tooltiptext {
        visibility: hidden;
        width: 200px;
        background-color: #555;
        color: #fff;
        text-align: center;
        border-radius: 6px;
        padding: 5px;
        position: absolute;
        z-index: 1;
        bottom: 125%;
        left: 50%;
        margin-left: -100px;
        opacity: 0;
        transition: opacity 0.3s;
        font-size: 12px;
    }
    .tooltip:hover .tooltiptext {
        visibility: visible;
        opacity: 1;
    }
    /* Notification badges */
    .notification-badge {
        background-color: #ff4444;
        color: white;
        border-radius: 50%;
        padding: 2px 8px;
        font-size: 12px;
        margin-left: 5px;
    }
    /* Progress animation */
    .progress-animated {
        transition: width 0.5s ease-in-out;
    }
    /* Better cards */
    .metric-card {
        background: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        margin: 5px 0;
    }
</style>
""", unsafe_allow_html=True)

# --- AUTHENTICATION SYSTEM ---
def init_auth():
    """Initialize authentication session state"""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "username" not in st.session_state:
        st.session_state.username = None
    if "login_attempts" not in st.session_state:
        st.session_state.login_attempts = 0
    if "last_login_attempt" not in st.session_state:
        st.session_state.last_login_attempt = None

def hash_password(password):
    """Hash password for security"""
    return hashlib.sha256(password.encode()).hexdigest()

def verify_user(username, password):
    """Verify user credentials (store in secrets for production)"""
    # In production, store hashed passwords in a secure database
    # This is a simple demo with hardcoded users
    valid_users = {
        "admin": hash_password("admin123"),
        "manager": hash_password("manager123"),
        "operator": hash_password("operator123")
    }
    
    # Rate limiting - prevent brute force
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
    """Display login form"""
    with st.container():
        st.markdown("""
        <div style="
            max-width: 400px; 
            margin: 100px auto; 
            padding: 40px; 
            border-radius: 10px; 
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            background: white;
        ">
            <h2 style="text-align: center; color: #1e88e5;">🏭 Production Tasks</h2>
            <p style="text-align: center; color: #666;">Please login to continue</p>
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
                    if st.session_state.login_attempts >= 3:
                        st.warning(f"⚠️ {5 - st.session_state.login_attempts} attempts remaining before lockout")

def logout():
    """Logout user"""
    st.session_state.authenticated = False
    st.session_state.username = None
    st.rerun()

# --- GOOGLE SHEETS CONNECTION WITH RETRY ---
@st.cache_resource
def get_gspread_client():
    """Get Google Sheets client with retry logic"""
    if "gcp_service_account" not in st.secrets:
        return None, "Το [gcp_service_account] δεν βρέθηκε στα Secrets του Streamlit."
    
    max_retries = 3
    retry_delay = 1
    
    for attempt in range(max_retries):
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
            if attempt < max_retries - 1:
                time.sleep(retry_delay * (attempt + 1))
                continue
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

# --- DATA LOADING WITH CACHE AND VERSIONING ---
@st.cache_data(ttl=60, show_spinner=False)
def load_all_data(version=0):
    """Load all data with caching and versioning"""
    try:
        # Load procurement data
        df_proc_raw = pd.read_csv(PROC_CSV_URL, header=None)
        indices = [18, 3, 17, 1, 2, 5, 6, 4, 10, 12]
        df_proc = df_proc_raw.iloc[1:, indices].copy()
        df_proc.columns = [
            "ID", "Ημερομηνία Παράδοσης", "Είδος Δώρου", "Project", 
            "Ποσότητα", "Προμηθευτής", "Υλικό / Προϊόν", 
            "Αναμενόμενη Ημ. Παραλαβής", "Αναμενόμενη Ποσότητα Παραλαβής", "Status Procurement"
        ]
        df_proc = df_proc.fillna("-")
    except Exception as e:
        st.warning(f"⚠️ Could not load procurement data: {e}")
        df_proc = pd.DataFrame()

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
    except Exception as e:
        st.warning(f"⚠️ Could not load tasks database: {e}")
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
    except Exception as e:
        st.warning(f"⚠️ Could not load team data: {e}")

    return df_proc, tasks_dict, team_members, availability_dict

# --- ASSIGNMENTS MANAGEMENT ---
@st.cache_data(ttl=30)
def load_assignments_from_sheet():
    """Load assignments with caching"""
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
    except Exception as e:
        st.warning(f"⚠️ Could not load assignments: {e}")
    
    return assignments_item, assignments_proj

def save_all_assignments_to_sheet():
    """Save assignments with retry logic and backup"""
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

        # Item tasks
        for u_key, t_list in st.session_state.get("tasks_store", {}).items():
            item_id = u_key.split("_")[0]
            proj_name = item_to_project.get(item_id, "-")
            for t in t_list:
                if t.get("task") and t.get("task") != "- Επιλογή Εργασίας -":
                    rows.append([
                        proj_name, item_id, t.get("task"), t.get("user"), 
                        str(t.get("date")), str(t.get("done")), "ITEM"
                    ])

        # Project tasks
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
        
        # Update last save timestamp
        st.session_state.last_save = datetime.now()
        return True
        
    except Exception as e:
        st.error(f"❌ Σφάλμα κατά την αποθήκευση: {e}")
        return False

# --- AUDIT LOG ---
def add_to_audit_log(action, details):
    """Add entry to audit log"""
    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []
    
    st.session_state.audit_log.append({
        "timestamp": datetime.now(),
        "user": st.session_state.get("username", "unknown"),
        "action": action,
        "details": details
    })
    
    # Keep only last 1000 entries
    if len(st.session_state.audit_log) > 1000:
        st.session_state.audit_log = st.session_state.audit_log[-1000:]

# --- EXPORT FUNCTIONS ---
def generate_printable_html(title, date_str, df_data):
    """Generate printable HTML with better formatting"""
    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{title}</title>
        <style>
            * {{ box-sizing: border-box; }}
            body {{ 
                font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Arial, sans-serif; 
                margin: 20px; 
                color: #333; 
                line-height: 1.6;
            }}
            .container {{ max-width: 1200px; margin: 0 auto; }}
            h2 {{ 
                color: #1e88e5; 
                border-bottom: 3px solid #1e88e5; 
                padding-bottom: 10px; 
                margin-bottom: 20px;
            }}
            .header-info {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                flex-wrap: wrap;
                margin-bottom: 20px;
                padding: 15px;
                background: #f5f5f5;
                border-radius: 8px;
            }}
            .date {{ font-size: 14px; color: #666; }}
            .summary {{ 
                background: #e3f2fd; 
                padding: 10px 20px; 
                border-radius: 5px;
                font-weight: bold;
            }}
            table {{ 
                width: 100%; 
                border-collapse: collapse; 
                margin-top: 15px;
                font-size: 13px;
            }}
            th, td {{ 
                border: 1px solid #ddd; 
                padding: 10px; 
                text-align: left; 
            }}
            th {{ 
                background-color: #f2f2f2; 
                font-weight: bold; 
                color: #111; 
                position: sticky;
                top: 0;
            }}
            tr:nth-child(even) {{ background-color: #f9f9f9; }}
            .done {{ color: #2e7d32; font-weight: bold; }}
            .pending {{ color: #c62828; font-weight: bold; }}
            .status-ok {{ color: #2e7d32; }}
            .status-pending {{ color: #ed6c02; }}
            .status-error {{ color: #c62828; }}
            @media print {{
                .no-print {{ display: none; }}
                body {{ margin: 10px; }}
                th {{ background-color: #e0e0e0 !important; }}
                .header-info {{ background: #f5f5f5 !important; }}
            }}
            @media (max-width: 768px) {{
                table { font-size: 11px; }
                th, td {{ padding: 6px; }}
                .header-info {{ flex-direction: column; align-items: flex-start; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>🏭 {title}</h2>
            <div class="header-info">
                <div class="date">📅 Ημερομηνία: <b>{date_str}</b></div>
                <div class="summary">📊 Σύνολο Εργασιών: {len(df_data)}</div>
            </div>
            <table>
                <thead>
                    <tr>
                        {"".join([f"<th>{col}</th>" for col in df_data.columns])}
                    </tr>
                </thead>
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
            elif "OK STOCK" in val or "RECEIVED" in val or "READY" in val:
                val_str = f'<span class="status-ok">✅ {val}</span>'
            elif "ORDERED" in val or "PENDING" in val:
                val_str = f'<span class="status-pending">⏳ {val}</span>'
            else:
                val_str = val
            html += f"<td>{val_str}</td>"
        html += "</tr>"
    
    html += f"""
                </tbody>
            </table>
            <br>
            <div class="no-print" style="text-align: center; margin-top: 20px;">
                <button onclick="window.print()" style="
                    padding:12px 30px; 
                    background:#1e88e5; 
                    color:white; 
                    border:none; 
                    border-radius:5px; 
                    cursor:pointer;
                    font-size:16px;
                ">
                    🖨️ Εκτύπωση / Αποθήκευση σε PDF
                </button>
            </div>
        </div>
    </body>
    </html>
    """
    return html

def export_to_excel(df, title):
    """Export DataFrame to Excel with formatting"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Report', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Report']
        
        # Add formats
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#1e88e5',
            'font_color': 'white',
            'border': 1
        })
        
        # Write headers with formatting
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Auto-adjust columns
        for i, col in enumerate(df.columns):
            max_width = max(df[col].astype(str).str.len().max(), len(col)) + 2
            worksheet.set_column(i, i, min(max_width, 50))
    
    return output.getvalue()

# --- NOTIFICATIONS ---
def check_notifications():
    """Check for notifications and return count"""
    notifications = []
    
    # Check for pending procurement items
    if 'procurement_df' in st.session_state and not st.session_state.procurement_df.empty:
        pending = st.session_state.procurement_df[
            ~st.session_state.procurement_df["Status Procurement"].isin(["OK STOCK", "RECEIVED", "READY"])
        ]
        if not pending.empty:
            notifications.append(f"⚠️ {len(pending)} υλικά σε εκκρεμότητα procurement")
    
    # Check for overbooked days
    if 'availability_database' in st.session_state:
        today = date.today()
        greek_day = WEEKDAYS_GREEK.get(today.weekday(), "Δευτέρα")
        day_avail = st.session_state.availability_database.get(greek_day, {})
        
        # Calculate total assigned for today
        total_assigned = 0
        for _, t_list in st.session_state.get("tasks_store", {}).items():
            for t in t_list:
                if t.get("date") == today and t.get("task") != "- Επιλογή Εργασίας -":
                    total_assigned += 1
        
        max_capacity = sum(day_avail.values()) * 60 / 60  # Convert to hours
        if total_assigned > max_capacity:
            notifications.append(f"⚠️ Overbooked today: {total_assigned} tasks vs {max_capacity:.0f} capacity")
    
    return notifications

# --- TOGGLE FUNCTIONS ---
def toggle_project_task(p_key, task_name, chk_key):
    st.session_state["project_tasks_store"][p_key][task_name]["done"] = st.session_state[chk_key]
    save_all_assignments_to_sheet()
    add_to_audit_log("toggle_project_task", f"{p_key} - {task_name} - {st.session_state[chk_key]}")

def toggle_item_task(u_key, t_idx, chk_key):
    st.session_state["tasks_store"][u_key][t_idx]["done"] = st.session_state[chk_key]
    save_all_assignments_to_sheet()
    add_to_audit_log("toggle_item_task", f"{u_key} - {t_idx} - {st.session_state[chk_key]}")

def update_item_field(u_key, t_idx, field, widget_key):
    st.session_state["tasks_store"][u_key][t_idx][field] = st.session_state[widget_key]
    save_all_assignments_to_sheet()
    add_to_audit_log("update_item_field", f"{u_key} - {field}: {st.session_state[widget_key]}")

def update_proj_field(p_key, task_name, field, widget_key):
    st.session_state["project_tasks_store"][p_key][task_name][field] = st.session_state[widget_key]
    save_all_assignments_to_sheet()
    add_to_audit_log("update_proj_field", f"{p_key} - {task_name} - {field}: {st.session_state[widget_key]}")

# --- RENDER FUNCTIONS ---
def render_dashboard(procurement_df, tasks_database, team_database, availability_database):
    """Render Dashboard tab"""
    st.header("📈 Dashboard & Επισκόπηση Παραγωγής")
    
    if procurement_df.empty:
        st.warning("⚠️ No procurement data available.")
        return
    
    projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
    
    # Calculate metrics
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
        
        # Item tasks
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

        # Project tasks
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

    # Metrics row
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Ενεργά Projects", len(projects_list))
    c2.metric("Συνολικές Ώρες", f"{round(tot_all_hours, 1)}h")
    
    overall_pct = int((tot_done_tasks / tot_tasks_count) * 100) if tot_tasks_count > 0 else 0
    c3.metric("Συνολική Πρόοδος", f"{overall_pct}%")
    c4.metric("Εκκρεμή Tasks", tot_tasks_count - tot_done_tasks)

    st.divider()
    
    # Charts row
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📊 Ώρες ανά Project")
        if project_hours:
            fig = px.bar(
                x=list(project_hours.keys()),
                y=list(project_hours.values()),
                title="Συνολικές Ώρες ανά Project",
                labels={"x": "Project", "y": "Ώρες"}
            )
            fig.update_layout(
                xaxis_tickangle=-45,
                height=400,
                showlegend=False
            )
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.subheader("📈 Πρόοδος ανά Project")
        if project_progress:
            fig = px.bar(
                x=list(project_progress.keys()),
                y=list(project_progress.values()),
                title="Ποσοστό Ολοκλήρωσης ανά Project",
                labels={"x": "Project", "y": "Πρόοδος (%)"},
                color=list(project_progress.values()),
                color_continuous_scale="RdYlGn",
                range_color=[0, 100]
            )
            fig.update_layout(
                xaxis_tickangle=-45,
                height=400,
                showlegend=False
            )
            st.plotly_chart(fig, use_container_width=True)

    st.divider()
    
    # Data table
    st.subheader("📋 Λεπτομερής Κατάσταση Projects")
    if dashboard_data:
        dash_df = pd.DataFrame(dashboard_data)
        st.dataframe(dash_df, use_container_width=True, hide_index=True)
    
    # Add export buttons
    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        if dashboard_data:
            dash_df = pd.DataFrame(dashboard_data)
            csv = dash_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📊 Εξαγωγή CSV",
                data=csv,
                file_name=f"Dashboard_{date.today().strftime('%Y-%m-%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )
    with col_exp2:
        if dashboard_data:
            dash_df = pd.DataFrame(dashboard_data)
            excel_data = export_to_excel(dash_df, "Dashboard")
            st.download_button(
                label="📄 Εξαγωγή Excel",
                data=excel_data,
                file_name=f"Dashboard_{date.today().strftime('%Y-%m-%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

def render_project(procurement_df, tasks_database, team_database, availability_database):
    """Render Project tab"""
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

    for
