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
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    div[data-testid="stMetricValue"] { font-size: 1.2rem; }
    div[data-testid="stMetricLabel"] { font-size: 0.75rem; }
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
        st.session_state.page = "📦 Projects"

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
            st.error("Too many failed attempts. Please wait 5 minutes.")
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
    st.markdown("<div style='height:20vh;'></div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("""
        <div style="text-align:center;margin-bottom:30px;">
            <div style="font-size:20px;font-weight:600;color:#555;">Production Tasks</div>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="", label_visibility="collapsed")
            password = st.text_input("Password", type="password", placeholder="Password", label_visibility="collapsed")
            submitted = st.form_submit_button("Login", use_container_width=True)
            
            if submitted:
                if verify_user(username, password):
                    st.rerun()
                else:
                    st.error("Λάθος στοιχεία")

def logout():
    total_tasks_in_memory = sum(len(tasks) for tasks in st.session_state.get("tasks_store", {}).values())
    total_proj_tasks_in_memory = sum(
        len([t for t in p_dict.values() if isinstance(t, dict) and t.get("active", False)])
        for p_dict in st.session_state.get("project_tasks_store", {}).values()
        if isinstance(p_dict, dict)
    )
    
    if total_tasks_in_memory > 0 or total_proj_tasks_in_memory > 0:
        save_all_assignments_to_sheet()
    
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
INCOMING_GID = "1362920506"
PROC_GID = "1639392743"

INCOMING_CSV_URL = f"https://docs.google.com/spreadsheets/d/{PROC_SHEET_ID}/export?format=csv&gid={INCOMING_GID}"
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
    "Φωτογράφιση"
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

    df_incoming = pd.DataFrame()
    try:
        df_incoming_raw = pd.read_csv(INCOMING_CSV_URL, header=None)
        
        if not df_incoming_raw.empty:
            df_incoming = df_incoming_raw.iloc[1:, [1, 11]].copy()
            df_incoming.columns = ["Project", "Shipping Status"]
            df_incoming = df_incoming.dropna(subset=["Project"])
            df_incoming["Project"] = df_incoming["Project"].astype(str).str.strip()
            df_incoming["Shipping Status"] = df_incoming["Shipping Status"].astype(str).str.strip()
            df_incoming = df_incoming[df_incoming["Project"] != ""]
            df_incoming = df_incoming[df_incoming["Project"] != "nan"]
    except Exception as e:
        st.warning(f"Could not load Incoming Projects List: {e}")

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
            user_raw = str(r.get("Assigned_User", "- Χωρίς Ανάθεση -"))
            user = user_raw
            users_list = [u.strip() for u in user_raw.split(",") if u.strip() and u.strip() != "- Χωρίς Ανάθεση -"]
            
            assign_date_str = str(r.get("Assigned_Date", r.get("Assigned_Done", "")))
            
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
                    "user": users_list[0] if users_list else "- Χωρίς Ανάθεση -",
                    "users": users_list,
                    "date": assign_date
                }
            else:
                if item_id not in assignments_item:
                    assignments_item[item_id] = []
                
                is_duplicate = False
                for existing in assignments_item[item_id]:
                    if (existing["task"] == task_name and 
                        existing["user"] == user and 
                        str(existing["date"]) == str(assign_date)):
                        is_duplicate = True
                        break
                
                if not is_duplicate:
                    assignments_item[item_id].append({
                        "done": done,
                        "task": task_name,
                        "user": users_list[0] if users_list else "- Χωρίς Ανάθεση -",
                        "users": users_list,
                        "date": assign_date
                    })
                    
    except Exception as e:
        st.warning(f"Could not load assignments: {e}")
    
    return assignments_item, assignments_proj

def save_all_assignments_to_sheet():
    gc, _ = get_gspread_client()
    if not gc:
        st.warning("Δεν είναι δυνατή η αποθήκευση λόγω σφάλματος σύνδεσης API.")
        return False
    
    try:
        total_tasks = sum(len(tasks) for tasks in st.session_state.get("tasks_store", {}).values())
        total_proj = sum(
            len([t for t in p_dict.values() if isinstance(t, dict) and t.get("active", False)])
            for p_dict in st.session_state.get("project_tasks_store", {}).values()
            if isinstance(p_dict, dict)
        )
        
        if total_tasks == 0 and total_proj == 0:
            try:
                existing_sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Assignments")
                existing_data = existing_sheet.get_all_values()
                if len(existing_data) > 1:
                    st.warning("⚠️ Αποτροπή διαγραφής: Η μνήμη είναι κενή αλλά το Sheet έχει δεδομένα.")
                    return False
            except:
                pass
        
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
                    task_users = t.get("users", [])
                    if not task_users:
                        task_users = [t.get("user", "- Χωρίς Ανάθεση -")]
                    
                    users_str = ", ".join([str(u) for u in task_users if u and u != "- Χωρίς Ανάθεση -"])
                    if not users_str:
                        users_str = "- Χωρίς Ανάθεση -"
                    
                    rows.append([
                        proj_name, item_id, t.get("task"), users_str, 
                        str(t.get("date")), str(t.get("done")), "ITEM"
                    ])

        for p_key, p_dict in st.session_state.get("project_tasks_store", {}).items():
            proj_name = p_key.replace("proj_", "")
            if isinstance(p_dict, dict):
                for t_name, p_data in p_dict.items():
                    if isinstance(p_data, dict) and p_data.get("active", False):
                        task_users = p_data.get("users", [])
                        if not task_users:
                            task_users = [p_data.get("user", "- Χωρίς Ανάθεση -")]
                        
                        users_str = ", ".join([str(u) for u in task_users if u and u != "- Χωρίς Ανάθεση -"])
                        if not users_str:
                            users_str = "- Χωρίς Ανάθεση -"
                        
                        rows.append([
                            proj_name, "-", t_name, users_str, 
                            str(p_data.get("date")), str(p_data.get("done")), "PROJECT"
                        ])

        sheet.clear()
        sheet.update(range_name="A1", values=rows)
        st.session_state.last_save = datetime.now()
        return True
        
    except Exception as e:
        st.error(f"Σφάλμα κατά την αποθήκευση: {e}")
        return False

# --- EXPORT FUNCTIONS ---
def generate_printable_html(title, date_str, df_data):
    """Δημιουργεί ένα βελτιωμένο HTML report για εκτύπωση."""
    
    # --- ΥΠΟΛΟΓΙΣΜΟΙ ΓΙΑ ΣΥΝΟΨΗ ---
    total_rows = len(df_data)
    
    # Προσπαθούμε να βρούμε στήλες με συγκεκριμένα ονόματα
    completed_count = 0
    pending_count = 0
    total_hours = 0.0
    
    if "Ολοκληρώθηκε" in df_data.columns:
        completed_count = len(df_data[df_data["Ολοκληρώθηκε"] == "ΝΑΙ"])
        pending_count = len(df_data[df_data["Ολοκληρώθηκε"] == "ΟΧΙ"])
    elif "Status" in df_data.columns:
        completed_count = len(df_data[df_data["Status"].astype(str).str.contains("✅", na=False)])
        pending_count = len(df_data[df_data["Status"].astype(str).str.contains("⏳", na=False)])
    
    if "Ώρες" in df_data.columns:
        try:
            total_hours = round(df_data["Ώρες"].astype(float).sum(), 1)
        except:
            pass
    
    # --- ΔΗΜΙΟΥΡΓΙΑ ΠΙΝΑΚΑ ---
    table_rows = ""
    for idx, row in df_data.iterrows():
        row_class = ""
        table_rows += "<tr>"
        for col in df_data.columns:
            val = str(row[col])
            cell_class = ""
            
            # Χρωματισμός για Status columns
            if val == "ΝΑΙ" or "✅" in val:
                cell_class = ' class="status-done"'
            elif val == "ΟΧΙ" or "⏳" in val:
                cell_class = ' class="status-pending"'
            elif val == "OK Shipped" or "🔴" in val:
                cell_class = ' class="status-shipped"'
            elif val == "ΕΝΕΡΓΟ" or "🟢" in val:
                cell_class = ' class="status-active"'
            elif "⚠️" in val:
                cell_class = ' class="status-warning"'
            
            table_rows += f"<td{cell_class}>{val}</td>"
        table_rows += "</tr>"
    
    # --- SUMMARY CARDS ---
    summary_html = ""
    if completed_count > 0 or pending_count > 0:
        summary_html = f"""
        <div class="summary-cards">
            <div class="summary-card">
                <div class="summary-value">{total_rows}</div>
                <div class="summary-label">Σύνολο</div>
            </div>
            <div class="summary-card success">
                <div class="summary-value">{completed_count}</div>
                <div class="summary-label">✅ Ολοκληρωμένα</div>
            </div>
            <div class="summary-card warning">
                <div class="summary-value">{pending_count}</div>
                <div class="summary-label">⏳ Εκκρεμή</div>
            </div>
            <div class="summary-card info">
                <div class="summary-value">{total_hours}h</div>
                <div class="summary-label">⏱️ Συνολικές Ώρες</div>
            </div>
        </div>
        """
    
    # --- HTML ---
    html = f"""
    <!DOCTYPE html>
    <html lang="el">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{title}</title>
        <style>
            * {{ box-sizing: border-box; margin: 0; padding: 0; }}
            body {{ 
                font-family: 'Segoe UI', Arial, sans-serif; 
                margin: 0; 
                padding: 30px; 
                color: #2c3e50; 
                background: #f5f7fa;
                line-height: 1.5;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background: white;
                padding: 40px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                border-radius: 8px;
            }}
            .header {{
                border-bottom: 3px solid #1e88e5;
                padding-bottom: 20px;
                margin-bottom: 30px;
            }}
            .header-top {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 15px;
            }}
            .company-name {{
                font-size: 24px;
                font-weight: 700;
                color: #1e88e5;
                letter-spacing: 0.5px;
            }}
            .report-badge {{
                background: #1e88e5;
                color: white;
                padding: 6px 14px;
                border-radius: 20px;
                font-size: 12px;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 1px;
            }}
            h1 {{
                font-size: 28px;
                color: #2c3e50;
                font-weight: 600;
                margin-bottom: 8px;
            }}
            .meta {{
                display: flex;
                gap: 25px;
                font-size: 13px;
                color: #7f8c8d;
                margin-top: 10px;
            }}
            .meta span {{
                display: flex;
                align-items: center;
                gap: 6px;
            }}
            .summary-cards {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
                gap: 15px;
                margin: 30px 0;
            }}
            .summary-card {{
                padding: 18px;
                border-radius: 8px;
                background: #f8f9fa;
                border-left: 4px solid #95a5a6;
                text-align: center;
            }}
            .summary-card.success {{
                background: #e8f5e9;
                border-left-color: #2e7d32;
            }}
            .summary-card.warning {{
                background: #fff8e1;
                border-left-color: #f57c00;
            }}
            .summary-card.info {{
                background: #e3f2fd;
                border-left-color: #1565c0;
            }}
            .summary-value {{
                font-size: 26px;
                font-weight: 700;
                color: #2c3e50;
                margin-bottom: 4px;
            }}
            .summary-card.success .summary-value {{ color: #2e7d32; }}
            .summary-card.warning .summary-value {{ color: #f57c00; }}
            .summary-card.info .summary-value {{ color: #1565c0; }}
            .summary-label {{
                font-size: 12px;
                color: #7f8c8d;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                margin-top: 20px;
                font-size: 13px;
                border-radius: 6px;
                overflow: hidden;
            }}
            thead {{
                background: #1e88e5;
                color: white;
            }}
            th {{
                padding: 12px 15px;
                text-align: left;
                font-weight: 600;
                font-size: 12px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            td {{
                padding: 10px 15px;
                border-bottom: 1px solid #ecf0f1;
                color: #34495e;
            }}
            tr:nth-child(even) {{ background: #f8f9fa; }}
            tr:hover {{ background: #e3f2fd; }}
            .status-done {{
                color: #2e7d32;
                font-weight: 700;
                background: #e8f5e9;
            }}
            .status-pending {{
                color: #d32f2f;
                font-weight: 700;
                background: #ffebee;
            }}
            .status-shipped {{
                color: #c62828;
                font-weight: 700;
            }}
            .status-active {{
                color: #2e7d32;
                font-weight: 700;
            }}
            .status-warning {{
                color: #f57c00;
                font-weight: 700;
            }}
            .footer {{
                margin-top: 40px;
                padding-top: 20px;
                border-top: 1px solid #ecf0f1;
                display: flex;
                justify-content: space-between;
                font-size: 11px;
                color: #95a5a6;
            }}
            .btn {{
                display: inline-block;
                padding: 12px 24px;
                background: #1e88e5;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 600;
                cursor: pointer;
                text-decoration: none;
                margin-top: 20px;
                transition: background 0.2s;
            }}
            .btn:hover {{ background: #1565c0; }}
            .btn-secondary {{
                background: #78909c;
                margin-left: 10px;
            }}
            .btn-secondary:hover {{ background: #546e7a; }}
            @media print {{
                body {{ background: white; padding: 0; }}
                .container {{ box-shadow: none; padding: 20px; }}
                .btn, .btn-secondary {{ display: none; }}
                table {{ font-size: 11px; }}
                th, td {{ padding: 6px 10px; }}
                .summary-cards {{ break-inside: avoid; }}
                tr {{ break-inside: avoid; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div class="header-top">
                    <div class="company-name">🏭 Production Tasks</div>
                    <div class="report-badge">Report</div>
                </div>
                <h1>{title}</h1>
                <div class="meta">
                    <span>📅 <b>{date_str}</b></span>
                    <span>📊 <b>{total_rows}</b> εγγραφές</span>
                    <span>🕐 Δημιουργήθηκε: <b>{datetime.now().strftime('%d/%m/%Y %H:%M')}</b></span>
                </div>
            </div>
            
            {summary_html}
            
            <table>
                <thead>
                    <tr>
                        {"".join([f"<th>{col}</th>" for col in df_data.columns])}
                    </tr>
                </thead>
                <tbody>
                    {table_rows}
                </tbody>
            </table>
            
            <div class="footer">
                <div>© Production Tasks App</div>
                <div>Σελίδα 1</div>
            </div>
            
            <div style="text-align:center;">
                <button onclick="window.print()" class="btn">🖨️ Εκτύπωση</button>
                <button onclick="window.close()" class="btn btn-secondary">✕ Κλείσιμο</button>
            </div>
        </div>
    </body>
    </html>
    """
    return html

def generate_printable_html_two_sections(title, date_str, pending_df, completed_df):
    """Δημιουργεί HTML report με 2 ενότητες: Εκκρεμή + Ολοκληρωμένα."""
    
    pending_count = len(pending_df)
    completed_count = len(completed_df)
    total_count = pending_count + completed_count
    
    # --- ΠΙΝΑΚΑΣ ΕΚΚΡΕΜΩΝ ---
    pending_table = ""
    if pending_count > 0:
        pending_rows = ""
        for _, row in pending_df.iterrows():
            pending_rows += "<tr>"
            for col in pending_df.columns:
                pending_rows += f"<td>{str(row[col])}</td>"
            pending_rows += "</tr>"
        
        pending_table = f"""
        <div class="section">
            <h2 class="section-title pending-title">⏳ ΕΚΚΡΕΜΗ ({pending_count})</h2>
            <table>
                <thead>
                    <tr>{"".join([f"<th>{col}</th>" for col in pending_df.columns])}</tr>
                </thead>
                <tbody>
                    {pending_rows}
                </tbody>
            </table>
        </div>
        """
    else:
        pending_table = """
        <div class="section">
            <h2 class="section-title success-title">🎉 Όλες οι εργασίες έχουν ολοκληρωθεί!</h2>
            <p style="text-align:center;font-size:16px;color:#2e7d32;padding:20px;">Δεν υπάρχουν εκκρεμείς εργασίες.</p>
        </div>
        """
    
    # --- ΠΙΝΑΚΑΣ ΟΛΟΚΛΗΡΩΜΕΝΩΝ ---
    completed_table = ""
    if completed_count > 0:
        completed_rows = ""
        for _, row in completed_df.iterrows():
            completed_rows += "<tr>"
            for col in completed_df.columns:
                completed_rows += f"<td>{str(row[col])}</td>"
            completed_rows += "</tr>"
        
        completed_table = f"""
        <div class="section">
            <h2 class="section-title completed-title">✅ ΟΛΟΚΛΗΡΩΜΕΝΑ ({completed_count})</h2>
            <table>
                <thead>
                    <tr>{"".join([f"<th>{col}</th>" for col in completed_df.columns])}</tr>
                </thead>
                <tbody>
                    {completed_rows}
                </tbody>
            </table>
        </div>
        """
    
    html = f"""
    <!DOCTYPE html>
    <html lang="el">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{title}</title>
        <style>
            * {{ box-sizing: border-box; margin: 0; padding: 0; }}
            body {{ 
                font-family: 'Segoe UI', Arial, sans-serif; 
                margin: 0; 
                padding: 30px; 
                color: #2c3e50; 
                background: #f5f7fa;
                line-height: 1.5;
            }}
            .container {{
                max-width: 1200px;
                margin: 0 auto;
                background: white;
                padding: 40px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                border-radius: 8px;
            }}
            .header {{
                border-bottom: 3px solid #1e88e5;
                padding-bottom: 20px;
                margin-bottom: 30px;
            }}
            .header-top {{
                display: flex;
                justify-content: space-between;
                align-items: center;
                margin-bottom: 15px;
            }}
            .company-name {{
                font-size: 24px;
                font-weight: 700;
                color: #1e88e5;
                letter-spacing: 0.5px;
            }}
            .report-badge {{
                background: #1e88e5;
                color: white;
                padding: 6px 14px;
                border-radius: 20px;
                font-size: 12px;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 1px;
            }}
            h1 {{
                font-size: 28px;
                color: #2c3e50;
                font-weight: 600;
                margin-bottom: 8px;
            }}
            .meta {{
                display: flex;
                gap: 25px;
                font-size: 13px;
                color: #7f8c8d;
                margin-top: 10px;
            }}
            .meta span {{
                display: flex;
                align-items: center;
                gap: 6px;
            }}
            .summary-cards {{
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
                gap: 15px;
                margin: 30px 0;
            }}
            .summary-card {{
                padding: 18px;
                border-radius: 8px;
                background: #f8f9fa;
                border-left: 4px solid #95a5a6;
                text-align: center;
            }}
            .summary-card.success {{
                background: #e8f5e9;
                border-left-color: #2e7d32;
            }}
            .summary-card.warning {{
                background: #fff8e1;
                border-left-color: #f57c00;
            }}
            .summary-card.info {{
                background: #e3f2fd;
                border-left-color: #1565c0;
            }}
            .summary-value {{
                font-size: 26px;
                font-weight: 700;
                color: #2c3e50;
                margin-bottom: 4px;
            }}
            .summary-card.success .summary-value {{ color: #2e7d32; }}
            .summary-card.warning .summary-value {{ color: #f57c00; }}
            .summary-card.info .summary-value {{ color: #1565c0; }}
            .summary-label {{
                font-size: 12px;
                color: #7f8c8d;
                font-weight: 600;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            .section {{
                margin-bottom: 40px;
            }}
            .section-title {{
                font-size: 20px;
                padding: 12px 18px;
                border-radius: 6px;
                margin-bottom: 15px;
                font-weight: 700;
                letter-spacing: 0.5px;
            }}
            .pending-title {{
                background: #fff8e1;
                color: #f57c00;
                border-left: 5px solid #f57c00;
            }}
            .completed-title {{
                background: #e8f5e9;
                color: #2e7d32;
                border-left: 5px solid #2e7d32;
            }}
            .success-title {{
                background: #e8f5e9;
                color: #2e7d32;
                border-left: 5px solid #2e7d32;
                text-align: center;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                font-size: 13px;
                border-radius: 6px;
                overflow: hidden;
            }}
            thead {{
                background: #1e88e5;
                color: white;
            }}
            th {{
                padding: 12px 15px;
                text-align: left;
                font-weight: 600;
                font-size: 12px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            td {{
                padding: 10px 15px;
                border-bottom: 1px solid #ecf0f1;
                color: #34495e;
            }}
            tr:nth-child(even) {{ background: #f8f9fa; }}
            tr:hover {{ background: #e3f2fd; }}
            .footer {{
                margin-top: 40px;
                padding-top: 20px;
                border-top: 1px solid #ecf0f1;
                display: flex;
                justify-content: space-between;
                font-size: 11px;
                color: #95a5a6;
            }}
            .btn {{
                display: inline-block;
                padding: 12px 24px;
                background: #1e88e5;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 600;
                cursor: pointer;
                text-decoration: none;
                margin-top: 20px;
                transition: background 0.2s;
            }}
            .btn:hover {{ background: #1565c0; }}
            .btn-secondary {{
                background: #78909c;
                margin-left: 10px;
            }}
            .btn-secondary:hover {{ background: #546e7a; }}
            @media print {{
                body {{ background: white; padding: 0; }}
                .container {{ box-shadow: none; padding: 20px; }}
                .btn, .btn-secondary {{ display: none; }}
                table {{ font-size: 11px; }}
                th, td {{ padding: 6px 10px; }}
                .summary-cards {{ break-inside: avoid; }}
                tr {{ break-inside: avoid; }}
                .section {{ break-inside: avoid; }}
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div class="header-top">
                    <div class="company-name">🏭 Production Tasks</div>
                    <div class="report-badge">Report</div>
                </div>
                <h1>{title}</h1>
                <div class="meta">
                    <span>📅 <b>{date_str}</b></span>
                    <span>📊 <b>{total_count}</b> εργασίες</span>
                    <span>🕐 <b>{datetime.now().strftime('%d/%m/%Y %H:%M')}</b></span>
                </div>
            </div>
            
            <div class="summary-cards">
                <div class="summary-card">
                    <div class="summary-value">{total_count}</div>
                    <div class="summary-label">Σύνολο</div>
                </div>
                <div class="summary-card warning">
                    <div class="summary-value">{pending_count}</div>
                    <div class="summary-label">⏳ Εκκρεμή</div>
                </div>
                <div class="summary-card success">
                    <div class="summary-value">{completed_count}</div>
                    <div class="summary-label">✅ Ολοκληρωμένα</div>
                </div>
            </div>
            
            {pending_table}
            
            {completed_table}
            
            <div class="footer">
                <div>© Production Tasks App</div>
                <div>Σελίδα 1</div>
            </div>
            
            <div style="text-align:center;">
                <button onclick="window.print()" class="btn">🖨️ Εκτύπωση</button>
                <button onclick="window.close()" class="btn btn-secondary">✕ Κλείσιμο</button>
            </div>
        </div>
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
            notifications.append(f"{len(pending)} υλικά σε εκκρεμότητα procurement")
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

def update_item_users(u_key, t_idx, widget_key):
    selected = st.session_state[widget_key]
    st.session_state["tasks_store"][u_key][t_idx]["users"] = list(selected)
    if selected:
        st.session_state["tasks_store"][u_key][t_idx]["user"] = selected[0]
    else:
        st.session_state["tasks_store"][u_key][t_idx]["user"] = "- Χωρίς Ανάθεση -"
    save_all_assignments_to_sheet()

def update_proj_field(p_key, task_name, field, widget_key):
    st.session_state["project_tasks_store"][p_key][task_name][field] = st.session_state[widget_key]
    save_all_assignments_to_sheet()

def update_proj_users(p_key, task_name, widget_key):
    selected = st.session_state[widget_key]
    st.session_state["project_tasks_store"][p_key][task_name]["users"] = list(selected)
    if selected:
        st.session_state["project_tasks_store"][p_key][task_name]["user"] = selected[0]
    else:
        st.session_state["project_tasks_store"][p_key][task_name]["user"] = "- Χωρίς Ανάθεση -"
    save_all_assignments_to_sheet()

# --- CHECK USER OVERLOAD ---
def check_user_overload(target_date, availability_database):
    """
    Επιστρέφει dict με {user: {'assigned': X, 'max': Y, 'over': Z}} 
    για όλους τους χρήστες που έχουν ΥΠΕΡΒΑΣΗ για τη συγκεκριμένη ημερομηνία.
    """
    if not target_date:
        return {}
    
    greek_day = WEEKDAYS_GREEK.get(target_date.weekday(), "Δευτέρα")
    day_availability = availability_database.get(greek_day, {})
    tasks_db = st.session_state.get("_tasks_db_cache", {})
    procurement_df = st.session_state.get("procurement_df", pd.DataFrame())
    
    user_hours = {}
    
    # Item tasks
    for u_key, task_list in st.session_state.get("tasks_store", {}).items():
        for t in task_list:
            if t.get("task") and t["task"] != "- Επιλογή Εργασίας -" and t.get("date") == target_date:
                t_users = t.get("users", [])
                if not t_users:
                    t_users = [t.get("user", "- Χωρίς Ανάθεση -")]
                num_users = max(len(t_users), 1)
                
                # Παίρνουμε qty από το αντίστοιχο item
                qty = 1
                try:
                    parts = u_key.split("_")
                    item_id = parts[0]
                    idx_part = int(parts[1]) if len(parts) > 1 else 0
                    if not procurement_df.empty and idx_part < len(procurement_df):
                        row = procurement_df.iloc[idx_part]
                        if str(row["ID"]) == item_id:
                            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
                except:
                    pass
                
                auto_time = tasks_db.get(t["task"], 0.0)
                hrs_per_user = (auto_time * qty / 60) / num_users if auto_time else 0
                for user in t_users:
                    if user and user != "- Χωρίς Ανάθεση -":
                        user_hours[user] = user_hours.get(user, 0.0) + hrs_per_user
    
    # Project tasks
    for p_key, p_dict in st.session_state.get("project_tasks_store", {}).items():
        if isinstance(p_dict, dict):
            proj_name = p_key.replace("proj_", "")
            proj_qty = 1
            if not procurement_df.empty:
                p_items = procurement_df[procurement_df["Project"] == proj_name]
                for _, r in p_items.iterrows():
                    if str(r["Ποσότητα"]).isdigit():
                        proj_qty = max(proj_qty, int(r["Ποσότητα"]))
            
            for t_name, p_data in p_dict.items():
                if isinstance(p_data, dict) and p_data.get("active", False) and p_data.get("date") == target_date:
                    t_users = p_data.get("users", [])
                    if not t_users:
                        t_users = [p_data.get("user", "- Χωρίς Ανάθεση -")]
                    num_users = max(len(t_users), 1)
                    auto_time = tasks_db.get(t_name, 0.0)
                    hrs_per_user = (auto_time * proj_qty / 60) / num_users if auto_time else 0
                    for user in t_users:
                        if user and user != "- Χωρίς Ανάθεση -":
                            user_hours[user] = user_hours.get(user, 0.0) + hrs_per_user
    
    overloads = {}
    for user, hrs in user_hours.items():
        max_hrs = day_availability.get(user, 6.0)
        if hrs > max_hrs:
            overloads[user] = {
                "assigned": round(hrs, 2),
                "max": round(max_hrs, 2),
                "over": round(hrs - max_hrs, 2)
            }
    
    return overloads


def render_overload_warnings(target_dates, availability_database):
    """
    Εμφανίζει warning banners για ΥΠΕΡΒΑΣΕΙΣ ωρών.
    """
    if not target_dates:
        return
    
    all_warnings = {}
    for check_date in sorted(target_dates):
        overloads = check_user_overload(check_date, availability_database)
        for user, info in overloads.items():
            if check_date not in all_warnings:
                all_warnings[check_date] = []
            all_warnings[check_date].append((user, info))
    
    for check_date in sorted(all_warnings.keys()):
        for user, info in all_warnings[check_date]:
            st.warning(
                f"⚠️ **Υπέρβαση ωρών** — **{user}** στις **{check_date.strftime('%d/%m/%Y')}**: "
                f"**{info['assigned']}h** (όριο: {info['max']}h) — **+{info['over']}h**"
            )

# --- PROJECT DETAILS FUNCTION ---
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
    
    is_shipped = False
    if not incoming_df.empty and "Project" in incoming_df.columns and "Shipping Status" in incoming_df.columns:
        project_name_clean = str(project_name).strip().upper()
        
        for _, inc_row in incoming_df.iterrows():
            inc_project = str(inc_row["Project"]).strip().upper()
            if inc_project == project_name_clean:
                shipping_status = str(inc_row["Shipping Status"]).strip().upper()
                if "OK SHIPPED" in shipping_status or "OK SHIP" in shipping_status or shipping_status == "SHIPPED":
                    is_shipped = True
                break
    
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

# --- RENDER PROJECTS (MAIN PAGE) ---
def render_projects(procurement_df, tasks_database, team_database, availability_database, incoming_df):
    st.header("Projects")
    
    if procurement_df.empty:
        st.warning("No procurement data available.")
        return
    
    if "selected_project_drill" not in st.session_state:
        st.session_state.selected_project_drill = None
    if "selected_material_expand" not in st.session_state:
        st.session_state.selected_material_expand = None
    
    # =========================================================
    # MODE 2: DETAILS VIEW (Drill-Down)
    # =========================================================
    if st.session_state.selected_project_drill is not None:
        selected_project = st.session_state.selected_project_drill
        
        col_back, _ = st.columns([1, 4])
        with col_back:
            if st.button("← Πίσω στα Projects", use_container_width=True):
                st.session_state.selected_project_drill = None
                st.session_state.selected_material_expand = None
                st.rerun()
        
        st.divider()
        
        proj_details = get_project_details(selected_project, procurement_df, tasks_database, incoming_df)
        
        col_title, col_status = st.columns([3, 1])
        with col_title:
            st.markdown(f"## {selected_project}")
        with col_status:
            if proj_details['is_shipped']:
                st.markdown('<div style="background:#ffebee;color:#c62828;padding:8px 16px;border-radius:10px;text-align:center;font-weight:700;font-size:13px;">SHIPPED</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div style="background:#e8f5e9;color:#2e7d32;padding:8px 16px;border-radius:10px;text-align:center;font-weight:700;font-size:13px;">ΕΝΕΡΓΟ</div>', unsafe_allow_html=True)
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ώρες", f"{proj_details['total_hours']}h")
        m2.metric("Πρόοδος", f"{proj_details['progress']}%")
        m3.metric("Υλικά", proj_details['materials_count'])
        m4.metric("Tasks", f"{proj_details['completed_tasks']}/{proj_details['total_tasks']}")
        
        st.progress(proj_details['progress'] / 100)
        st.caption(f"**Κατάσταση:** {proj_details['status']}")
        
        # --- ΕΛΕΓΧΟΣ ΥΠΕΡΒΑΣΗΣ ΩΡΩΝ ---
        filtered_df_check = procurement_df[procurement_df["Project"] == selected_project].copy()
        all_dates_check = set()
        for idx_chk, row_chk in filtered_df_check.iterrows():
            item_id_chk = str(row_chk["ID"])
            u_key_chk = f"{item_id_chk}_{idx_chk}"
            for t_chk in st.session_state.get("tasks_store", {}).get(u_key_chk, []):
                if t_chk.get("task") and t_chk["task"] != "- Επιλογή Εργασίας -" and t_chk.get("date"):
                    all_dates_check.add(t_chk["date"])
        proj_key_chk = f"proj_{selected_project}"
        for t_name_chk, p_data_chk in st.session_state.get("project_tasks_store", {}).get(proj_key_chk, {}).items():
            if isinstance(p_data_chk, dict) and p_data_chk.get("active", False) and p_data_chk.get("date"):
                all_dates_check.add(p_data_chk["date"])
        render_overload_warnings(all_dates_check, availability_database)
        
        st.divider()
        
        tab1, tab2 = st.tabs(["Υλικά", "Γενικές Εργασίες"])
        
        # ==================== TAB 1: ΥΛΙΚΑ ====================
        with tab1:
            filtered_df = procurement_df[procurement_df["Project"] == selected_project].copy()
            
            col_search, col_status_filter, col_clear = st.columns([2, 1, 1])
            with col_search:
                search_term = st.text_input("Αναζήτηση υλικού:", placeholder="ID ή όνομα...", key="proj_mat_search")
            with col_status_filter:
                status_filter = st.selectbox("Status:", ["Όλα", "Έτοιμα", "Σε Εκκρεμότητα"], key="proj_mat_filter")
            with col_clear:
                st.write("")
                st.write("")
                if st.button("Καθαρισμός", use_container_width=True):
                    st.session_state.selected_material_expand = None
                    st.rerun()
            
            if search_term:
                filtered_df = filtered_df[
                    filtered_df["ID"].astype(str).str.contains(search_term, case=False, na=False) |
                    filtered_df["Υλικό / Προϊόν"].astype(str).str.contains(search_term, case=False, na=False)
                ]
            
            if status_filter == "Έτοιμα":
                filtered_df = filtered_df[filtered_df["Status Procurement"].isin(["OK STOCK", "RECEIVED", "READY"])]
            elif status_filter == "Σε Εκκρεμότητα":
                filtered_df = filtered_df[~filtered_df["Status Procurement"].isin(["OK STOCK", "RECEIVED", "READY"])]
            
            st.caption(f"Εμφανίζονται **{len(filtered_df)}** υλικά")
            
            if filtered_df.empty:
                st.info("Δεν βρέθηκαν υλικά με τα συγκεκριμένα φίλτρα.")
            else:
                for idx, row in filtered_df.iterrows():
                    item_id = str(row["ID"])
                    unique_key = f"{item_id}_{idx}"
                    material = str(row["Υλικό / Προϊόν"])
                    qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
                    status = str(row["Status Procurement"])
                    
                    item_tasks = st.session_state.get("tasks_store", {}).get(unique_key, [])
                    task_count = sum(1 for t in item_tasks if t.get("task") and t["task"] != "- Επιλογή Εργασίας -")
                    done_count = sum(1 for t in item_tasks if t.get("task") and t["task"] != "- Επιλογή Εργασίας -" and t.get("done", False))
                    
                    is_open = st.session_state.selected_material_expand == unique_key
                    
                    if task_count > 0:
                        progress_pct = int((done_count / task_count) * 100)
                        if progress_pct == 100:
                            bar_color = "#2e7d32"
                        elif progress_pct > 0:
                            bar_color = "#1e88e5"
                        else:
                            bar_color = "#9e9e9e"
                        progress_html = f'<div style="display:inline-flex;align-items:center;gap:6px;"><div style="width:50px;height:6px;background:#e0e0e0;border-radius:3px;overflow:hidden;"><div style="width:{progress_pct}%;height:100%;background:{bar_color};"></div></div><span style="font-size:11px;color:#555;">{done_count}/{task_count}</span></div>'
                    else:
                        progress_html = '<span style="font-size:11px;color:#999;">0/0</span>'

                    row_due_date = str(row["Αναμενόμενη Ημ. Παραλαβής"]).strip() if "Αναμενόμενη Ημ. Παραλαβής" in row else ""
                    if row_due_date in ["-", "nan", ""]:
                        row_due_date = ""

                    due_date_html = ""
                    if row_due_date:
                        due_date_html = f'<span style="background:#e3f2fd;color:#1565c0;padding:3px 8px;border-radius:6px;font-size:11px;font-weight:600;margin-left:6px;">📅 {row_due_date}</span>'
                    
                    if status in ["OK STOCK", "RECEIVED", "READY"]:
                        status_bg = "background:#e8f5e9;color:#2e7d32;"
                    else:
                        status_bg = "background:#fff8e1;color:#f57c00;"

                    col_info, col_btn = st.columns([8, 1])
                    
                    with col_info:
                        html_line = (
                            f'<div style="display:flex;align-items:center;gap:12px;padding:10px 15px;background:#ffffff;border:1px solid #e0e0e0;border-radius:6px;font-size:13px;">'
                            f'<span style="font-weight:700;color:#333;min-width:70px;">{item_id}</span>'
                            f'<span style="color:#333;flex:1;">{material[:50]}{"..." if len(material)>50 else ""}</span>'
                            f'<span style="color:#666;min-width:60px;">{qty} τμχ</span>'
                            f'{due_date_html}'
                            f'<span style="padding:3px 8px;border-radius:6px;font-size:11px;font-weight:600;{status_bg}">{status}</span>'
                            f'{progress_html}'
                            f'</div>'
                        )
                        st.markdown(html_line, unsafe_allow_html=True)
                    
                    with col_btn:
                        btn_icon = "▼" if is_open else "▶"
                        if st.button(btn_icon, key=f"mat_row_{unique_key}", use_container_width=True):
                            if is_open:
                                st.session_state.selected_material_expand = None
                            else:
                                st.session_state.selected_material_expand = unique_key
                            st.rerun()
                    
                    if is_open:
                        if status not in ["OK STOCK", "RECEIVED", "READY"]:
                            st.warning(f"Εκκρεμότητα Procurement: {status}")
                        else:
                            st.success(f"Υλικό Διαθέσιμο: {status}")
                        
                        due_date = str(row["Αναμενόμενη Ημ. Παραλαβής"]) if "Αναμενόμενη Ημ. Παραλαβής" in row else "-"
                        expected_qty = str(row["Αναμενόμενη Ποσότητα Παραλαβής"]) if "Αναμενόμενη Ποσότητα Παραλαβής" in row else "-"
                        
                        info_parts = []
                        if due_date and due_date != "-" and due_date != "nan":
                            info_parts.append(f"📅 Αναμενόμενη Παραλαβή: **{due_date}**")
                        if expected_qty and expected_qty != "-" and expected_qty != "nan":
                            info_parts.append(f"📦 Αναμενόμενη Ποσότητα: **{expected_qty}**")
                        
                        if info_parts:
                            st.info(" | ".join(info_parts))
                        
                        if unique_key not in st.session_state["tasks_store"]:
                            st.session_state["tasks_store"][unique_key] = []
                        
                        item_tasks = st.session_state["tasks_store"][unique_key]
                        task_options = ["- Επιλογή Εργασίας -"] + sorted(list(tasks_database.keys()))
                        team_options = ["- Χωρίς Ανάθεση -"] + team_database
                        
                        if not item_tasks:
                            st.info("Δεν έχουν οριστεί εργασίες. Πάτησε **Προσθήκη Εργασίας** παρακάτω.")
                        
                        for t_idx, t_data in enumerate(list(item_tasks)):
                            with st.container(border=True):
                                c_done, c_task = st.columns([0.1, 0.9])
                                
                                chk_k = f"mat_chk_{unique_key}_{t_idx}"
                                c_done.checkbox("Done", value=t_data["done"], key=chk_k, on_change=toggle_item_task, args=(unique_key, t_idx, chk_k))
                                
                                task_idx = task_options.index(t_data["task"]) if t_data["task"] in task_options else 0
                                task_k = f"mat_task_{unique_key}_{t_idx}"
                                c_task.selectbox("Εργασία", task_options, index=task_idx, key=task_k, on_change=update_item_field, args=(unique_key, t_idx, "task", task_k))
                                
                                c_user, c_date, c_time, c_del = st.columns([0.4, 0.3, 0.2, 0.1])
                                
                                current_users = t_data.get("users", [])
                                if not isinstance(current_users, list):
                                    current_users = [current_users] if current_users else []
                                
                                valid_users = [u for u in current_users if u in team_database]
                                
                                user_k = f"mat_users_{unique_key}_{t_idx}"
                                c_user.multiselect(
                                    "Υπεύθυνοι",
                                    options=team_database,
                                    default=valid_users,
                                    key=user_k,
                                    on_change=update_item_users,
                                    args=(unique_key, t_idx, user_k)
                                )
                                
                                date_k = f"mat_date_{unique_key}_{t_idx}"
                                c_date.date_input("Ημερομηνία", value=t_data["date"], format="DD/MM/YYYY", key=date_k, on_change=update_item_field, args=(unique_key, t_idx, "date", date_k))
                                
                                auto_time = tasks_database.get(t_data["task"], 0.0)
                                if t_data["task"] != "- Επιλογή Εργασίας -":
                                    num_users = max(len(t_data.get("users", [])), 1)
                                    task_hours_total = (auto_time * qty) / 60
                                    task_hours_per_user = task_hours_total / num_users
                                    if num_users > 1:
                                        c_time.metric("Ώρες/άτομο", f"{round(task_hours_per_user, 2)}h")
                                    else:
                                        c_time.metric("Ώρες", f"{round(task_hours_per_user, 2)}h")
                                else:
                                    c_time.caption("—")
                                
                                if c_del.button("🗑️", key=f"mat_del_{unique_key}_{t_idx}"):
                                    st.session_state["tasks_store"][unique_key].pop(t_idx)
                                    save_all_assignments_to_sheet()
                                    st.rerun()
                        
                        col_add, col_rem, col_close = st.columns([1, 1, 1])
                        if col_add.button("Προσθήκη Εργασίας", key=f"mat_add_{unique_key}", use_container_width=True):
                            auto_date = date.today()
                            try:
                                due_date_str = str(row["Αναμενόμενη Ημ. Παραλαβής"]).strip()
                                if due_date_str and due_date_str not in ["-", "nan", ""]:
                                    for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"]:
                                        try:
                                            auto_date = datetime.strptime(due_date_str.split(" ")[0], fmt.split(" ")[0]).date()
                                            break
                                        except:
                                            continue
                            except:
                                pass
                            
                            st.session_state["tasks_store"][unique_key].append({
                                "done": False, 
                                "task": "- Επιλογή Εργασίας -", 
                                "user": "- Χωρίς Ανάθεση -", 
                                "users": [], 
                                "date": auto_date
                            })
                            save_all_assignments_to_sheet()
                            st.rerun()
                        if len(item_tasks) > 0 and col_rem.button("Αφαίρεση", key=f"mat_rem_{unique_key}", use_container_width=True):
                            st.session_state["tasks_store"][unique_key].pop()
                            save_all_assignments_to_sheet()
                            st.rerun()
                        if col_close.button("Κλείσιμο", key=f"mat_close_{unique_key}", use_container_width=True):
                            st.session_state.selected_material_expand = None
                            st.rerun()
        
        # ==================== TAB 2: ΓΕΝΙΚΕΣ ΕΡΓΑΣΙΕΣ ====================
        with tab2:
            st.markdown(f"### Γενικές Εργασίες — {selected_project}")
            st.caption("Ενεργοποίησε τις εργασίες που χρειάζονται και όρισε υπεύθυνο + ημερομηνία.")
            
            proj_key = f"proj_{selected_project}"
            if proj_key not in st.session_state["project_tasks_store"] or not isinstance(st.session_state["project_tasks_store"][proj_key], dict):
                st.session_state["project_tasks_store"][proj_key] = {t_name: {"active": False, "done": False, "user": "- Χωρίς Ανάθεση -", "users": [], "date": date.today()} for t_name in FIXED_PROJECT_TASKS}
            
            proj_tasks_dict = st.session_state["project_tasks_store"][proj_key]
            
            project_main_qty = 1
            for _, r in filtered_df.iterrows():
                if str(r["Ποσότητα"]).isdigit():
                    project_main_qty = max(project_main_qty, int(r["Ποσότητα"]))
            
            for task_name in FIXED_PROJECT_TASKS:
                t_data = proj_tasks_dict.get(task_name, {"active": False, "done": False, "user": "- Χωρίς Ανάθεση -", "users": [], "date": date.today()})
                
                with st.container(border=True):
                    c_active, c_name, c_done = st.columns([0.08, 0.62, 0.30])
                    pact_k = f"pact_proj_{proj_key}_{task_name}"
                    is_active = c_active.checkbox("", value=t_data["active"], key=pact_k, on_change=update_proj_field, args=(proj_key, task_name, "active", pact_k))
                    c_name.markdown(f"**{task_name}**" if is_active else f"<span style='color:gray;'>{task_name}</span>", unsafe_allow_html=True)
                    
                    if is_active:
                        pdone_k = f"proj_pdone_proj_{proj_key}_{task_name}"
                        is_done = c_done.checkbox("Done", value=t_data["done"], key=pdone_k, on_change=toggle_project_task, args=(proj_key, task_name, pdone_k))
                    else:
                        c_done.caption("—")
                    
                    if is_active:
                        c_user, c_date, c_time = st.columns([0.4, 0.4, 0.2])
                        current_users = t_data.get("users", [])
                        if not isinstance(current_users, list):
                            current_users = [current_users] if current_users else []
                        valid_users = [u for u in current_users if u in team_database]
                        
                        puser_k = f"puser_proj_{proj_key}_{task_name}"
                        c_user.multiselect(
                            "Υπεύθυνοι",
                            options=team_database,
                            default=valid_users,
                            key=puser_k,
                            on_change=update_proj_users,
                            args=(proj_key, task_name, puser_k)
                        )
                        
                        pdate_k = f"pdate_proj_{proj_key}_{task_name}"
                        c_date.date_input("Ημερομηνία", value=t_data["date"], format="DD/MM/YYYY", key=pdate_k, on_change=update_proj_field, args=(proj_key, task_name, "date", pdate_k))
                        
                        auto_time = tasks_database.get(task_name, 0.0)
                        num_users = max(len(t_data.get("users", [])), 1)
                        task_hours_total = (auto_time * project_main_qty) / 60
                        task_hours_per_user = task_hours_total / num_users
                        if num_users > 1:
                            c_time.metric("Ώρες/άτομο", f"{round(task_hours_per_user, 2)}h")
                        else:
                            c_time.metric("Ώρες", f"{round(task_hours_per_user, 2)}h")
        
        st.divider()
        col_save, _ = st.columns([1, 3])
        if col_save.button("Αποθήκευση Αλλαγών", use_container_width=True, type="primary", key="save_project"):
            if save_all_assignments_to_sheet():
                st.success("Όλες οι αναθέσεις αποθηκεύτηκαν.")
            else:
                st.error("Σφάλμα κατά την αποθήκευση")
        
        return
    
    # =========================================================
    # MODE 1: GRID VIEW (default)
    # =========================================================
    
    col_check, col_info = st.columns([1, 3])
    with col_check:
        show_all = st.checkbox(
            "Εμφάνιση Όλων των Projects (μαζί με OK SHIPPED)", 
            value=False,
            help="Αν τσεκαριστεί, εμφανίζονται ΟΛΑ τα projects."
        )
    
    all_projects = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
    
    if show_all:
        projects_to_show = all_projects
        section_title = "Όλα τα Projects"
    else:
        projects_to_show = []
        for p_name in all_projects:
            proj_data = get_project_details(p_name, procurement_df, tasks_database, incoming_df)
            if proj_data['is_active']:
                projects_to_show.append(p_name)
        section_title = "Ενεργά Projects"
    
    dashboard_data = []
    tot_all_hours = 0.0
    tot_tasks_count = 0
    tot_done_tasks = 0
    projects_details_cache = {}
    
    for p_name in projects_to_show:
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
        tot_tasks_count += p_tasks_cnt
        tot_done_tasks += p_done_cnt
        p_progress = int((p_done_cnt / p_tasks_cnt) * 100) if p_tasks_cnt > 0 else 0
        
        proj_details = get_project_details(p_name, procurement_df, tasks_database, incoming_df)
        projects_details_cache[p_name] = proj_details
        is_shipped = proj_details.get('is_shipped', False)
        
        dashboard_data.append({
            "Project": p_name,
            "Υλικά": len(filtered_p),
            "Σύνολο Tasks": p_tasks_cnt,
            "Ολοκληρωμένα": p_done_cnt,
            "Συνολικές Ώρες": round(p_total_hrs, 1),
            "Υπολειπόμενες": round(p_total_hrs - p_done_hrs, 1),
            "Πρόοδος": f"{p_progress}%",
            "Κατάσταση": "OK Shipped" if is_shipped else "ΕΝΕΡΓΟ"
        })

    st.markdown(f"### {section_title}")
    st.caption(f"Εμφανίζονται **{len(projects_to_show)}** projects")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Projects" if show_all else "Ενεργά", len(projects_to_show))
    c2.metric("Συνολικές Ώρες", f"{round(tot_all_hours, 1)}h")
    
    overall_pct = int((tot_done_tasks / tot_tasks_count) * 100) if tot_tasks_count > 0 else 0
    c3.metric("Συνολική Πρόοδος", f"{overall_pct}%")
    c4.metric("Εκκρεμή Tasks", tot_tasks_count - tot_done_tasks)

    st.divider()
    
    # --- ΕΛΕΓΧΟΣ ΥΠΕΡΒΑΣΗΣ ΩΡΩΝ ΓΙΑ ΟΛΑ ΤΑ ΕΝΕΡΓΑ PROJECTS ---
    all_dates_dashboard = set()
    for p_name_chk in projects_to_show:
        filtered_p_chk = procurement_df[procurement_df["Project"] == p_name_chk]
        for idx_chk, r_chk in filtered_p_chk.iterrows():
            u_key_chk = f"{str(r_chk['ID'])}_{idx_chk}"
            for t_chk in st.session_state.get("tasks_store", {}).get(u_key_chk, []):
                if t_chk.get("task") and t_chk["task"] != "- Επιλογή Εργασίας -" and t_chk.get("date"):
                    all_dates_dashboard.add(t_chk["date"])
        proj_key_chk = f"proj_{p_name_chk}"
        for t_name_chk, p_data_chk in st.session_state.get("project_tasks_store", {}).get(proj_key_chk, {}).items():
            if isinstance(p_data_chk, dict) and p_data_chk.get("active", False) and p_data_chk.get("date"):
                all_dates_dashboard.add(p_data_chk["date"])
    render_overload_warnings(all_dates_dashboard, availability_database)
    
    st.subheader(section_title)
    st.caption("Πάτησε σε μια κάρτα για να δεις τις λεπτομέρειες")
    
    if not projects_to_show:
        st.info("Δεν βρέθηκαν projects.")
    else:
        st.markdown("""
        <style>
            .compact-card {
                background: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                padding: 10px;
                margin-bottom: 8px;
                box-shadow: 0 1px 2px rgba(0,0,0,0.04);
                transition: transform 0.15s ease, box-shadow 0.15s ease;
                height: 100%;
            }
            .compact-card:hover {
                transform: translateY(-2px);
                box-shadow: 0 3px 8px rgba(0,0,0,0.1);
            }
            .compact-card.active {
                border-left: 3px solid #2e7d32;
            }
            .compact-card.shipped {
                border-left: 3px solid #c62828;
                opacity: 0.7;
            }
            .compact-title {
                font-size: 12px;
                font-weight: 700;
                color: #1a1a1a;
                margin-bottom: 4px;
                line-height: 1.2;
                word-wrap: break-word;
                min-height: 30px;
            }
            .compact-status {
                font-size: 9px;
                font-weight: 600;
                padding: 2px 6px;
                border-radius: 8px;
                display: inline-block;
                margin-bottom: 6px;
            }
            .compact-progress-bg {
                background: #e0e0e0;
                border-radius: 4px;
                height: 5px;
                overflow: hidden;
                margin: 6px 0;
            }
            .compact-progress-fill {
                height: 100%;
                border-radius: 4px;
            }
            .compact-meta {
                font-size: 10px;
                color: #666;
                margin-top: 4px;
                display: flex;
                justify-content: space-between;
            }
        </style>
        """, unsafe_allow_html=True)
        
        num_cols = 5
        for i in range(0, len(projects_to_show), num_cols):
            cols = st.columns(num_cols)
            batch = projects_to_show[i:i+num_cols]
            for j, p_name in enumerate(batch):
                proj_data = projects_details_cache[p_name]
                
                with cols[j]:
                    is_shipped = proj_data.get('is_shipped', False)
                    if is_shipped:
                        card_class = "compact-card shipped"
                        status_html = '<span class="compact-status" style="background:#ffebee;color:#c62828;">SHIPPED</span>'
                    else:
                        card_class = "compact-card active"
                        status_html = '<span class="compact-status" style="background:#e8f5e9;color:#2e7d32;">ΕΝΕΡΓΟ</span>'
                    
                    progress = proj_data['progress']
                    if progress == 100:
                        bar_color = "#2e7d32"
                    elif progress > 50:
                        bar_color = "#1e88e5"
                    elif progress > 0:
                        bar_color = "#ff9800"
                    else:
                        bar_color = "#9e9e9e"
                    
                    total_h = proj_data['total_hours']
                    materials = proj_data['materials_count']
                    done_tasks = proj_data['completed_tasks']
                    total_tasks = proj_data['total_tasks']
                    
                    display_name = p_name[:22] + "..." if len(p_name) > 22 else p_name
                    
                    card_html = f"""
                    <div class="{card_class}">
                        <div class="compact-title">{display_name}</div>
                        {status_html}
                        <div class="compact-progress-bg">
                            <div class="compact-progress-fill" style="width:{progress}%;background:{bar_color};"></div>
                        </div>
                        <div style="text-align:center;font-size:11px;font-weight:600;color:#333;">{progress}%</div>
                        <div class="compact-meta">
                            <span>{total_h}h</span>
                            <span>{materials} υλικά</span>
                            <span>{done_tasks}/{total_tasks}</span>
                        </div>
                    </div>
                    """
                    st.markdown(card_html, unsafe_allow_html=True)
                    
                    if st.button("Άνοιγμα", key=f"drill_{p_name}", use_container_width=True):
                        st.session_state.selected_project_drill = p_name
                        st.session_state.selected_material_expand = None
                        st.rerun()
        
        st.divider()
        col_exp1, col_exp2 = st.columns(2)
        dash_df = pd.DataFrame(dashboard_data)
        with col_exp1:
            csv = dash_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="Εξαγωγή CSV", 
                data=csv, 
                file_name=f"Projects_{'All' if show_all else 'Active'}_{date.today().strftime('%Y-%m-%d')}.csv", 
                mime="text/csv", 
                use_container_width=True
            )
        with col_exp2:
            excel_data = export_to_excel(dash_df, "Projects")
            st.download_button(
                label="Εξαγωγή Excel", 
                data=excel_data, 
                file_name=f"Projects_{'All' if show_all else 'Active'}_{date.today().strftime('%Y-%m-%d')}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True
            )

# --- RENDER MASTER VIEW ---
def render_master_view(procurement_df, tasks_database, team_database, availability_database, incoming_df):
    st.header("Master View")
    st.caption("Όλες οι εργασίες από όλα τα ενεργά projects σε ένα σημείο")
    
    if procurement_df.empty:
        st.warning("No procurement data available.")
        return
    
    all_tasks = []
    
    col_chk, col_info = st.columns([1, 3])
    with col_chk:
        show_all_projects = st.checkbox(
            "Εμφάνιση Όλων των Projects (μαζί με OK Shipped)",
            value=False,
            key="mv_show_all",
            help="Αν τσεκαριστεί, εμφανίζονται εργασίες από ΟΛΑ τα projects. Αν όχι, μόνο από τα ενεργά."
        )
    
    if show_all_projects:
        active_projects_set = set(p for p in procurement_df["Project"].unique() if p != "-")
    else:
        active_projects_set = set()
        for p_name in procurement_df["Project"].unique():
            if p_name != "-":
                proj_data = get_project_details(p_name, procurement_df, tasks_database, incoming_df)
                if proj_data['is_active']:
                    active_projects_set.add(p_name)
    
    for idx, row in procurement_df.iterrows():
        item_id = str(row["ID"])
        unique_key = f"{item_id}_{idx}"
        project_name = str(row["Project"])

        if project_name not in active_projects_set:
            continue
            
        material = str(row["Υλικό / Προϊόν"])
        qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
        status_proc = str(row["Status Procurement"])
        
        item_tasks = st.session_state.get("tasks_store", {}).get(unique_key, [])
        for t_idx, t in enumerate(item_tasks):
            if t.get("task") and t["task"] != "- Επιλογή Εργασίας -":
                auto_time = tasks_database.get(t["task"], 0.0)
                hrs = (auto_time * qty) / 60
                t_users = t.get("users", [])
                if not t_users:
                    t_users = [t.get("user", "- Χωρίς Ανάθεση -")]
                
                all_tasks.append({
                    "type": "item",
                    "u_key": unique_key,
                    "t_idx": t_idx,
                    "Project": project_name,
                    "Υλικό": material,
                    "ID": item_id,
                    "Ποσότητα": qty,
                    "Εργασία": t["task"],
                    "Υπεύθυνος": ", ".join(t_users),
                    "Ημερομηνία": t.get("date", date.today()),
                    "Status": "✅" if t.get("done", False) else "⏳",
                    "done": t.get("done", False),
                    "Ώρες": round(hrs, 2),
                    "status_proc": status_proc
                })
    
    for p_key, p_tasks_dict in st.session_state["project_tasks_store"].items():
        if isinstance(p_tasks_dict, dict):
            proj_name = p_key.replace("proj_", "")

            if proj_name not in active_projects_set:
                continue
                
            proj_qty = 1
            p_items = procurement_df[procurement_df["Project"] == proj_name]
            for _, r in p_items.iterrows():
                if str(r["Ποσότητα"]).isdigit():
                    proj_qty = max(proj_qty, int(r["Ποσότητα"]))
            for task_name, p_data in p_tasks_dict.items():
                if isinstance(p_data, dict) and p_data.get("active", False):
                    auto_time = tasks_database.get(task_name, 0.0)
                    hrs = (auto_time * proj_qty) / 60
                    t_users = p_data.get("users", [])
                    if not t_users:
                        t_users = [p_data.get("user", "- Χωρίς Ανάθεση -")]
                    
                    all_tasks.append({
                        "type": "project",
                        "p_key": p_key,
                        "task_name": task_name,
                        "Project": proj_name,
                        "Υλικό": "Γενική Σύνθεση",
                        "ID": "-",
                        "Ποσότητα": proj_qty,
                        "Εργασία": task_name,
                        "Υπεύθυνος": ", ".join(t_users),
                        "Ημερομηνία": p_data.get("date", date.today()),
                        "Status": "✅" if p_data.get("done", False) else "⏳",
                        "done": p_data.get("done", False),
                        "Ώρες": round(hrs, 2),
                        "status_proc": "READY"
                    })
    
    if not all_tasks:
        st.info("Δεν υπάρχουν εργασίες. Πήγαινε στο Projects για να προσθέσεις.")
        return
    
    st.divider()
    col_f1, col_f2, col_f3, col_f4 = st.columns([1, 1, 1, 2])
    with col_f1:
        status_filter = st.selectbox("Status:", ["Όλα", "⏳ Εκκρεμείς", "✅ Ολοκληρωμένες"], key="mv_status")
    with col_f2:
        available_projects = ["Όλα τα Projects"] + sorted(list(set(t["Project"] for t in all_tasks)))
        proj_filter = st.selectbox("Project:", available_projects, key="mv_proj")
    with col_f3:
        available_users = ["Όλοι οι Τεχνίτες"] + sorted(list(set(t["Υπεύθυνος"] for t in all_tasks)))
        user_filter = st.selectbox("Υπεύθυνος:", available_users, key="mv_user")
    with col_f4:
        search_term = st.text_input("Αναζήτηση:", placeholder="Project, υλικό, εργασία...", key="mv_search")
    
    filtered_tasks = all_tasks.copy()
    
    if status_filter == "⏳ Εκκρεμείς":
        filtered_tasks = [t for t in filtered_tasks if not t["done"]]
    elif status_filter == "✅ Ολοκληρωμένες":
        filtered_tasks = [t for t in filtered_tasks if t["done"]]
    
    if proj_filter != "Όλα τα Projects":
        filtered_tasks = [t for t in filtered_tasks if t["Project"] == proj_filter]
    
    if user_filter != "Όλοι οι Τεχνίτες":
        filtered_tasks = [t for t in filtered_tasks if t["Υπεύθυνος"] == user_filter]
    
    if search_term:
        search_lower = search_term.lower()
        filtered_tasks = [t for t in filtered_tasks if 
            search_lower in str(t["Project"]).lower() or
            search_lower in str(t["Υλικό"]).lower() or
            search_lower in str(t["Εργασία"]).lower()
        ]
    
    filtered_tasks = sorted(filtered_tasks, key=lambda x: (x["done"], x["Ημερομηνία"]))
    
    # --- ΕΛΕΓΧΟΣ ΥΠΕΡΒΑΣΗΣ ΩΡΩΝ ---
    all_dates_mv = set(t["Ημερομηνία"] for t in filtered_tasks if t.get("Ημερομηνία"))
    render_overload_warnings(all_dates_mv, availability_database)
    
    st.divider()
    total = len(filtered_tasks)
    pending = sum(1 for t in filtered_tasks if not t["done"])
    completed = sum(1 for t in filtered_tasks if t["done"])
    
    m1, m2, m3 = st.columns(3)
    m1.metric("Σύνολο", total)
    m2.metric("Εκκρεμείς", pending)
    m3.metric("Ολοκληρωμένες", completed)
    
    st.divider()
    
    st.markdown("### Bulk Actions")
    col_b1, col_b2, col_b3 = st.columns([1, 1, 2])
    with col_b1:
        bulk_date = st.date_input("Νέα ημερομηνία:", value=date.today(), format="DD/MM/YYYY", key="mv_bulk_date")
    with col_b2:
        st.write("")
        st.write("")
        if st.button("Εφαρμογή σε επιλεγμένες", use_container_width=True):
            changed = 0
            for t in filtered_tasks:
                chk_key = f"mv_chk_{t.get('u_key', t.get('p_key'))}_{t.get('t_idx', t.get('task_name'))}"
                if st.session_state.get(chk_key, False):
                    if t["type"] == "item":
                        st.session_state["tasks_store"][t["u_key"]][t["t_idx"]]["date"] = bulk_date
                    else:
                        st.session_state["project_tasks_store"][t["p_key"]][t["task_name"]]["date"] = bulk_date
                    changed += 1
            if changed > 0:
                save_all_assignments_to_sheet()
                st.success(f"Άλλαξε ημερομηνία σε {changed} εργασίες!")
                st.rerun()
            else:
                st.warning("Δεν επιλέχθηκε καμία εργασία.")
    
    with col_b3:
        st.write("")
        st.write("")
        st.caption("Επίλεξε εργασίες με τα checkboxes και πάτησε το κουμπί για να αλλάξεις την ημερομηνία τους.")
    
    st.divider()
    
    if not filtered_tasks:
        st.info("Δεν βρέθηκαν εργασίες με τα συγκεκριμένα φίλτρα.")
        return
    
    st.markdown(f"### Εργασίες ({len(filtered_tasks)})")
    
    h0, h1, h2, h3, h4, h5, h6, h7 = st.columns([0.4, 1.5, 2, 1.8, 2.5, 1.3, 0.7, 0.5])
    h0.markdown("**✓**")
    h1.markdown("**Project**")
    h2.markdown("**Υλικό**")
    h3.markdown("**Εργασία**")
    h4.markdown("**Υπεύθυνος**")
    h5.markdown("**Ημερομηνία**")
    h6.markdown("**Status**")
    h7.markdown("**🗑️**")
    
    for i, t in enumerate(filtered_tasks):
        row_id = f"mv_row_{i}_{t.get('u_key', t.get('p_key'))}_{t.get('t_idx', t.get('task_name'))}"
        
        c0, c1, c2, c3, c4, c5, c6, c7 = st.columns([0.4, 1.5, 2, 1.8, 2.5, 1.3, 0.7, 0.5])
        
        chk_key = f"mv_chk_{t.get('u_key', t.get('p_key'))}_{t.get('t_idx', t.get('task_name'))}"
        c0.checkbox("", key=chk_key)
        
        c1.caption(f"**{t['Project'][:18]}{'...' if len(t['Project'])>18 else ''}**")
        c2.caption(f"{t['Υλικό'][:25]}{'...' if len(t['Υλικό'])>25 else ''}")
        c3.caption(t['Εργασία'][:20])
        c4.caption(t['Υπεύθυνος'])
        
        date_key = f"mv_date_{t.get('u_key', t.get('p_key'))}_{t.get('t_idx', t.get('task_name'))}"
        new_date = c5.date_input("", value=t['Ημερομηνία'], format="DD/MM/YYYY", key=date_key, label_visibility="collapsed")
        if new_date != t['Ημερομηνία']:
            if t["type"] == "item":
                st.session_state["tasks_store"][t["u_key"]][t["t_idx"]]["date"] = new_date
            else:
                st.session_state["project_tasks_store"][t["p_key"]][t["task_name"]]["date"] = new_date
            save_all_assignments_to_sheet()
            st.rerun()
        
        c6.markdown(t['Status'])
        
        if c7.button("🗑️", key=f"mv_del_{t.get('u_key', t.get('p_key'))}_{t.get('t_idx', t.get('task_name'))}"):
            if t["type"] == "item":
                st.session_state["tasks_store"][t["u_key"]].pop(t["t_idx"])
            else:
                del st.session_state["project_tasks_store"][t["p_key"]][t["task_name"]]
            save_all_assignments_to_sheet()
            st.rerun()
    
    st.divider()
    export_data = [{
        "Project": t["Project"],
        "Υλικό": t["Υλικό"],
        "Εργασία": t["Εργασία"],
        "Υπεύθυνος": t["Υπεύθυνος"],
        "Ημερομηνία": str(t["Ημερομηνία"]),
        "Status": "Ολοκληρώθηκε" if t["done"] else "Εκκρεμεί",
        "Ώρες": t["Ώρες"],
        "Status Procurement": t["status_proc"]
    } for t in filtered_tasks]
    export_df = pd.DataFrame(export_data)
    csv_data = export_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button("Εξαγωγή CSV", data=csv_data, file_name=f"Master_View_{date.today().strftime('%Y-%m-%d')}.csv", mime="text/csv", use_container_width=True)


# --- RENDER DAILY PLAN ---
def render_daily_plan(procurement_df, tasks_database, team_database, availability_database):
    st.header("Συγκεντρωτικό Πλάνο Παραγωγής")
    col_d, col_fp, col_fu, col_fs = st.columns([1, 1, 1, 1])
    target_date = col_d.date_input("Ημερομηνία Πλάνου:", value=date.today(), format="DD/MM/YYYY")
    greek_day_name = WEEKDAYS_GREEK.get(target_date.weekday(), "Δευτέρα")
    st.caption(f"Ημέρα εβδομάδας: **{greek_day_name}**")

    # --- ΕΛΕΓΧΟΣ ΥΠΕΡΒΑΣΗΣ ΩΡΩΝ ---
    render_overload_warnings({target_date}, availability_database)

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
                    t_users = p_data.get("users", [])
                    if not t_users:
                        t_users = [p_data.get("user", "- Χωρίς Ανάθεση -")]
                    t_done = p_data.get("done", False)
                    auto_time = tasks_database.get(task_name, 0.0)
                    num_users = max(len(t_users), 1)
                    hours_total = (auto_time * proj_qty) / 60
                    hours_per_user = hours_total / num_users
                    
                    daily_tasks_raw.append({
                        "type": "project",
                        "p_key": p_key,
                        "task_name": task_name,
                        "Project": proj_name,
                        "Υλικό": "Γενική Σύνθεση / Box",
                        "Ποσότητα": proj_qty,
                        "Εργασία": task_name,
                        "Υπεύθυνος": ", ".join(t_users),
                        "Υπεύθυνος_list": t_users,
                        "Ώρες": round(hours_total, 2),
                        "Ώρες/άτομο": round(hours_per_user, 2),
                        "num_users": num_users,
                        "done": t_done,
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
            for t_idx, t_data in enumerate(item_tasks):
                t_task = t_data["task"]
                t_users = t_data.get("users", [])
                if not t_users:
                    t_users = [t_data.get("user", "- Χωρίς Ανάθεση -")]
                t_date = t_data["date"]
                t_done = t_data["done"]
                if t_task != "- Επιλογή Εργασίας -" and t_date == target_date:
                    auto_time = tasks_database.get(t_task, 0.0)
                    num_users = max(len(t_users), 1)
                    hours_total = (auto_time * qty) / 60
                    hours_per_user = hours_total / num_users
                    
                    daily_tasks_raw.append({
                        "type": "item",
                        "u_key": unique_item_key,
                        "t_idx": t_idx,
                        "Project": project_name,
                        "Υλικό": material,
                        "Ποσότητα": qty,
                        "Εργασία": t_task,
                        "Υπεύθυνος": ", ".join(t_users),
                        "Υπεύθυνος_list": t_users,
                        "Ώρες": round(hours_total, 2),
                        "Ώρες/άτομο": round(hours_per_user, 2),
                        "num_users": num_users,
                        "done": t_done,
                        "status_proc": proc_status
                    })

    available_projects = ["Όλα τα Projects"] + sorted(list(set(d["Project"] for d in daily_tasks_raw))) if daily_tasks_raw else ["Όλα τα Projects"]
    available_users = ["Όλοι οι Τεχνίτες"]
    if daily_tasks_raw:
        all_users = set()
        for d in daily_tasks_raw:
            for u in d["Υπεύθυνος_list"]:
                all_users.add(u)
        available_users += sorted(list(all_users))
    else:
        available_users = ["Όλοι οι Τεχνίτες"]
    
    available_statuses = ["Όλα τα Status"] + sorted(list(set(d["status_proc"] for d in daily_tasks_raw))) if daily_tasks_raw else ["Όλα τα Status"]

    selected_filter_proj = col_fp.selectbox("Φίλτρο Project:", available_projects)
    selected_filter_user = col_fu.selectbox("Φίλτρο Τεχνίτη:", available_users)
    selected_filter_status = col_fs.selectbox("Φίλτρο Procurement:", available_statuses)

    daily_tasks = []
    for d in daily_tasks_raw:
        if selected_filter_proj != "Όλα τα Projects" and d["Project"] != selected_filter_proj:
            continue
        if selected_filter_user != "Όλοι οι Τεχνίτες" and selected_filter_user not in d["Υπεύθυνος_list"]:
            continue
        if selected_filter_status != "Όλα τα Status" and d["status_proc"] != selected_filter_status:
            continue
        daily_tasks.append(d)

    st.divider()
    if daily_tasks:
        export_list = [{"Project": dt["Project"], "Υλικό / Είδος": dt["Υλικό"], "Ποσότητα": dt["Ποσότητα"], "Εργασία": dt["Εργασία"], "Υπεύθυνος": dt["Υπεύθυνος"], "Ώρες": dt["Ώρες"], "Status Procurement": dt["status_proc"], "Ολοκληρώθηκε": "ΝΑΙ" if dt["done"] else "ΟΧΙ"} for dt in daily_tasks]
        export_df = pd.DataFrame(export_list)
        csv_data = export_df.to_csv(index=False).encode('utf-8-sig')
        col_head, col_exp_csv, col_exp_pdf, col_exp_excel = st.columns([0.4, 0.2, 0.2, 0.2])
        col_head.subheader(f"Εργασίες για τις {target_date.strftime('%d/%m/%Y')} ({len(daily_tasks)} Tasks)")
        col_exp_csv.download_button(label="CSV", data=csv_data, file_name=f"Daily_Plan_{target_date.strftime('%Y-%m-%d')}.csv", mime="text/csv", use_container_width=True)
        printable_html = generate_printable_html("Ημερήσιο Πλάνο Παραγωγής", target_date.strftime('%d/%m/%Y'), export_df)
        col_exp_pdf.download_button(label="PDF", data=printable_html, file_name=f"Daily_Plan_{target_date.strftime('%Y-%m-%d')}.html", mime="text/html", use_container_width=True)
        excel_data = export_to_excel(export_df, "Daily Plan")
        col_exp_excel.download_button(label="Excel", data=excel_data, file_name=f"Daily_Plan_{target_date.strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        pending_proc = [dt for dt in daily_tasks if dt["status_proc"] not in ["OK STOCK", "RECEIVED", "READY"]]
        if pending_proc:
            st.warning(f"**Προσοχή:** Υπάρχουν **{len(pending_proc)} tasks** με υλικά σε εκκρεμότητα!")

        st.markdown("#### Φόρτος Εργασίας & Διαθεσιμότητα Ομάδας")
        day_availability = availability_database.get(greek_day_name, {})
        user_hours = {}
        for d in daily_tasks:
            for u in d["Υπεύθυνος_list"]:
                user_hours[u] = user_hours.get(u, 0.0) + d["Ώρες/άτομο"]
        
        cols = st.columns(max(len(user_hours), 1))
        for i, (member_name, assigned_hrs) in enumerate(user_hours.items()):
            assigned_hrs = round(assigned_hrs, 2)
            if member_name != "- Χωρίς Ανάθεση -":
                max_hrs = day_availability.get(member_name, 6.0)
                delta_hrs = round(assigned_hrs - max_hrs, 2)
                if delta_hrs > 0:
                    cols[i].metric(f"{member_name}", f"{assigned_hrs} / {max_hrs}h", delta=f"+{delta_hrs}h Υπερκάλυψη", delta_color="inverse")
                else:
                    cols[i].metric(f"{member_name}", f"{assigned_hrs} / {max_hrs}h", delta=f"{delta_hrs}h Διαθέσιμο", delta_color="normal")
            else:
                cols[i].metric(f"{member_name}", f"{assigned_hrs} Ώρες")
            
        st.divider()
        st.markdown("#### Διαδραστική Λίστα Εργασιών")
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
            col_tsk.markdown(f"~~{dt['Εργασία']}~~" if is_done else f"**{dt['Εργασία']}**")
            col_user.caption(dt['Υπεύθυνος'])
            
            if dt['num_users'] > 1:
                col_hrs.caption(f"{dt['Ώρες/άτομο']}h/άτομο")
            else:
                col_hrs.caption(f"{dt['Ώρες']}h")
            
            if dt['status_proc'] in ["OK STOCK", "RECEIVED", "READY"]:
                col_st.success(f"{dt['status_proc']}")
            else:
                col_st.error(f"{dt['status_proc']}")
    else:
        st.info(f"Δεν βρέθηκαν εργασίες για τις {target_date.strftime('%d/%m/%Y')} με τα συγκεκριμένα φίλτρα.")


# --- RENDER TECHNICIAN ---
def render_technician(procurement_df, tasks_database, team_database, availability_database):
    st.header("Ημερήσιο Πρόγραμμα ανά Τεχνίτη")
    c_date, c_user = st.columns([1, 1])
    target_date = c_date.date_input("Ημερομηνία:", value=date.today(), format="DD/MM/YYYY", key="tech_date")
    selected_member = c_user.selectbox("Επιλέξτε Τεχνίτη:", team_database)
    st.divider()

    # --- ΕΛΕΓΧΟΣ ΥΠΕΡΒΑΣΗΣ ΩΡΩΝ ---
    render_overload_warnings({target_date}, availability_database)

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
                if isinstance(p_data, dict) and p_data.get("active", False) and p_data.get("date") == target_date:
                    t_users = p_data.get("users", [])
                    if not t_users:
                        t_users = [p_data.get("user", "- Χωρίς Ανάθεση -")]
                    if selected_member in t_users:
                        auto_time = tasks_database.get(task_name, 0.0)
                        num_users = max(len(t_users), 1)
                        hours = (auto_time * proj_qty) / 60 / num_users
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
                t_users = t_data.get("users", [])
                if not t_users:
                    t_users = [t_data.get("user", "- Χωρίς Ανάθεση -")]
                if selected_member in t_users and t_data.get("date") == target_date:
                    t_task = t_data.get("task")
                    if t_task != "- Επιλογή Εργασίας -":
                        auto_time = tasks_database.get(t_task, 0.0)
                        num_users = max(len(t_users), 1)
                        hours = (auto_time * qty) / 60 / num_users
                        worker_tasks.append({"type": "item", "u_key": unique_item_key, "t_idx": t_idx, "project": project_name, "item": f"[{item_id}] {material}", "qty": qty, "task": t_task, "hours": round(hours, 2), "done": t_data.get("done", False), "status_proc": proc_status})

    if worker_tasks:
        w_export_list = [{"Project": wt["project"], "Υλικό / Είδος": wt["item"], "Ποσότητα": wt["qty"], "Εργασία": wt["task"], "Ώρες": wt["hours"], "Status Procurement": wt["status_proc"], "Ολοκληρώθηκε": "ΝΑΙ" if wt["done"] else "ΟΧΙ"} for wt in worker_tasks]
        w_export_df = pd.DataFrame(w_export_list)
        w_csv_data = w_export_df.to_csv(index=False).encode('utf-8-sig')
        col_w_head, col_w_csv, col_w_pdf, col_w_excel = st.columns([0.4, 0.2, 0.2, 0.2])
        col_w_head.subheader(f"Πρόγραμμα για {selected_member} — {target_date.strftime('%d/%m/%Y')}")
        col_w_csv.download_button(label="CSV", data=w_csv_data, file_name=f"Schedule_{selected_member.replace(' ', '_')}_{target_date.strftime('%Y-%m-%d')}.csv", mime="text/csv", use_container_width=True)
        # --- ΔΗΜΙΟΥΡΓΙΑ 2 DATAFRAMES (ΕΚΚΡΕΜΗ + ΟΛΟΚΛΗΡΩΜΕΝΑ) ---
        pending_worker_tasks = [wt for wt in worker_tasks if not wt["done"]]
        completed_worker_tasks = [wt for wt in worker_tasks if wt["done"]]
        
        if pending_worker_tasks:
            pending_export_list = [{"Project": wt["project"], "Υλικό / Είδος": wt["item"], "Ποσότητα": wt["qty"], "Εργασία": wt["task"], "Ώρες": wt["hours"], "Status Procurement": wt["status_proc"]} for wt in pending_worker_tasks]
            pending_export_df = pd.DataFrame(pending_export_list)
        else:
            pending_export_df = pd.DataFrame(columns=["Project", "Υλικό / Είδος", "Ποσότητα", "Εργασία", "Ώρες", "Status Procurement"])
        
        if completed_worker_tasks:
            completed_export_list = [{"Project": wt["project"], "Υλικό / Είδος": wt["item"], "Ποσότητα": wt["qty"], "Εργασία": wt["task"], "Ώρες": wt["hours"], "Status Procurement": wt["status_proc"]} for wt in completed_worker_tasks]
            completed_export_df = pd.DataFrame(completed_export_list)
        else:
            completed_export_df = pd.DataFrame(columns=["Project", "Υλικό / Είδος", "Ποσότητα", "Εργασία", "Ώρες", "Status Procurement"])
        
        w_printable_html = generate_printable_html_two_sections(f"Πρόγραμμα Τεχνίτη: {selected_member}", target_date.strftime('%d/%m/%Y'), pending_export_df, completed_export_df)
        col_w_pdf.download_button(label="PDF", data=w_printable_html, file_name=f"Schedule_{selected_member.replace(' ', '_')}_{target_date.strftime('%Y-%m-%d')}.html", mime="text/html", use_container_width=True)
        excel_data = export_to_excel(w_export_df, f"Schedule {selected_member}")
        col_w_excel.download_button(label="Excel", data=excel_data, file_name=f"Schedule_{selected_member.replace(' ', '_')}_{target_date.strftime('%Y-%m-%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        total_w_hours = sum(t["hours"] for t in worker_tasks)
        st.info(f"Συνολικός εκτιμώμενος χρόνος: **{round(total_w_hours, 1)} Ώρες** ({len(worker_tasks)} Tasks)")
        pending_w_proc = [wt for wt in worker_tasks if wt["status_proc"] not in ["OK STOCK", "RECEIVED", "READY"]]
        if pending_w_proc:
            st.warning(f"Ο/Η {selected_member} έχει **{len(pending_w_proc)} tasks** με υλικά σε εκκρεμότητα.")

        st.divider()
        for w_idx, wt in enumerate(worker_tasks):
            col_c, col_proj, col_mat, col_task, col_qty, col_h, col_proc = st.columns([0.10, 0.20, 0.26, 0.20, 0.08, 0.08, 0.12])
            if wt["type"] == "project":
                chk_k = f"tech_pdone_{wt['p_key']}_{wt['task_name']}_{w_idx}"
                is_done = col_c.checkbox("Done", value=wt["done"], key=chk_k, on_change=toggle_project_task, args=(wt['p_key'], wt['task_name'], chk_k))
            else:
                chk_k = f"tech_idone_{wt['u_key']}_{wt['t_idx']}_{w_idx}"
                is_done = col_c.checkbox("Done", value=wt["done"], key=chk_k, on_change=toggle_item_task, args=(wt['u_key'], wt['t_idx'], chk_k))
            col_proj.markdown(f"**{wt['project']}**")
            col_mat.write(wt['item'])
            col_task.markdown(f"~~{wt['task']}~~" if is_done else f"`{wt['task']}`")
            col_qty.write(f"{wt['qty']} τμχ")
            col_h.caption(f"{wt['hours']}h")
            
            if wt['status_proc'] in ["OK STOCK", "RECEIVED", "READY"]:
                col_proc.success(f"{wt['status_proc']}")
            else:
                col_proc.error(f"{wt['status_proc']}")
    else:
        st.success(f"Δεν έχουν ανατεθεί εργασίες στον/στην {selected_member} για τις {target_date.strftime('%d/%m/%Y')}.")


# --- RENDER PROJECTION ---
def render_projection(procurement_df, tasks_database, team_database, availability_database):
    st.header("Πρόβλεψη Φόρτου Εργασίας")
    start_monday = date.today() - timedelta(days=date.today().weekday())
    col_w_choice, col_range_choice, col_weekend = st.columns([1, 1, 1])
    week_choice = col_w_choice.selectbox("ΕΝΑΡΞΗ ΠΡΟΒΟΛΗΣ:", ["Τρέχουσα Εβδομάδα", "Επόμενη Εβδομάδα (+1)", "Μεθεπόμενη Εβδομάδα (+2)", "Προσαρμοσμένη Ημερομηνία"])
    range_weeks = col_range_choice.selectbox("ΕΥΡΟΣ ΠΡΟΒΟΛΗΣ:", ["1 Εβδομάδα", "2 Εβδομάδες", "4 Εβδομάδες / Μήνας"])
    include_weekends = col_weekend.checkbox("Συμπερίληψη Σαββατοκύριακων", value=False, key="proj_wknd")
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
                        t_users = p_data.get("users", [])
                        if not t_users:
                            t_users = [p_data.get("user", "- Χωρίς Ανάθεση -")]
                        if t_date in w_days:
                            auto_time = tasks_database.get(task_name, 0.0)
                            num_users = max(len(t_users), 1)
                            hrs_per_user = (auto_time * proj_qty) / 60 / num_users
                            col_str = f"{WEEKDAYS_SHORT_GREEK[t_date.weekday()]} {t_date.strftime('%d/%m')}"
                            for user in t_users:
                                if user in w_matrix:
                                    w_matrix[user][col_str] += hrs_per_user
                            d_totals[col_str] += hrs_per_user * num_users

        if not procurement_df.empty:
            for idx, row in procurement_df.iterrows():
                item_id = str(row["ID"])
                unique_item_key = f"{item_id}_{idx}"
                qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
                item_tasks = st.session_state["tasks_store"].get(unique_item_key, [])
                for t_data in item_tasks:
                    t_task = t_data.get("task")
                    t_users = t_data.get("users", [])
                    if not t_users:
                        t_users = [t_data.get("user", "- Χωρίς Ανάθεση -")]
                    t_date = t_data.get("date")
                    if t_task != "- Επιλογή Εργασίας -" and t_date in w_days:
                        auto_time = tasks_database.get(t_task, 0.0)
                        num_users = max(len(t_users), 1)
                        hrs_per_user = (auto_time * qty) / 60 / num_users
                        col_str = f"{WEEKDAYS_SHORT_GREEK[t_date.weekday()]} {t_date.strftime('%d/%m')}"
                        for user in t_users:
                            if user in w_matrix:
                                w_matrix[user][col_str] += hrs_per_user
                        d_totals[col_str] += hrs_per_user * num_users

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
    kc3.metric("Overbooked Ημέρες", f"{overbooked_days_count} / {len(all_flat_days)}")
    kc4.metric("Πληρότητα Περιόδου", f"{load_ratio}%", delta=f"{load_ratio - 100}%" if load_ratio > 100 else "Εντός Ορίων")

    st.divider()
    def highlight_total_row(row):
        if row.name == "Σύνολο Ημέρας (h)":
            return ["background-color: #2b303a; font-weight: bold; color: #00e676; border-top: 2px solid #00e676;"] * len(row)
        return [""] * len(row)

    for w_num, w_monday, w_days, w_matrix in weeks_matrices:
        w_sunday = w_days[-1]
        st.subheader(f"Εβδομάδα {w_num}: {w_monday.strftime('%d/%m/%Y')} έως {w_sunday.strftime('%d/%m/%Y')}")
        proj_df = pd.DataFrame(w_matrix).T
        proj_df = proj_df.round(1)
        proj_df["Σύνολο (h)"] = proj_df.sum(axis=1)
        total_row = proj_df.sum(axis=0).round(1)
        total_row.name = "Σύνολο Ημέρας (h)"
        proj_df = pd.concat([proj_df, pd.DataFrame(total_row).T])
        styled_df = proj_df.style.apply(highlight_total_row, axis=1)
        st.dataframe(styled_df, use_container_width=True)


# --- RENDER DAILY REPORT ---
def render_daily_report(procurement_df, tasks_database, team_database, availability_database):
    st.header("Ημερήσιος Απολογισμός Παραγωγής")
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
                    t_users = p_data.get("users", [])
                    if not t_users:
                        t_users = [p_data.get("user", "- Χωρίς Ανάθεση -")]
                    num_users = max(len(t_users), 1)
                    hrs_total = round((auto_time * proj_qty) / 60, 2)
                    hrs_per_user = round(hrs_total / num_users, 2)
                    
                    item_info = {
                        "Project": proj_name,
                        "Εργασία": task_name,
                        "Υλικό / Είδος": "Γενική Σύνθεση / Box",
                        "Υπεύθυνος": ", ".join(t_users),
                        "Ώρες": hrs_total,
                        "Ώρες/άτομο": hrs_per_user,
                        "num_users": num_users
                    }
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
                    t_users = t_data.get("users", [])
                    if not t_users:
                        t_users = [t_data.get("user", "- Χωρίς Ανάθεση -")]
                    num_users = max(len(t_users), 1)
                    hrs_total = round((auto_time * qty) / 60, 2)
                    hrs_per_user = round(hrs_total / num_users, 2)
                    
                    item_info = {
                        "Project": project_name,
                        "Εργασία": t_data["task"],
                        "Υλικό / Είδος": f"[{item_id}] {material}",
                        "Υπεύθυνος": ", ".join(t_users),
                        "Ώρες": hrs_total,
                        "Ώρες/άτομο": hrs_per_user,
                        "num_users": num_users
                    }
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
    st.subheader("Ολοκληρωμένες Εργασίες")
    if rep_completed:
        display_data = []
        for item in rep_completed:
            display_data.append({
                "Project": item["Project"],
                "Εργασία": item["Εργασία"],
                "Υλικό / Είδος": item["Υλικό / Είδος"],
                "Υπεύθυνος": item["Υπεύθυνος"],
                "Ώρες": item["Ώρες"],
                "Ώρες/άτομο": item["Ώρες/άτομο"] if item["num_users"] > 1 else ""
            })
        st.dataframe(pd.DataFrame(display_data), use_container_width=True, hide_index=True)
    else:
        st.info("Δεν υπάρχουν ολοκληρωμένες εργασίες για αυτή την ημερομηνία.")

    st.divider()
    st.subheader("Εκκρεμότητες")
    if rep_pending:
        display_data = []
        for item in rep_pending:
            display_data.append({
                "Project": item["Project"],
                "Εργασία": item["Εργασία"],
                "Υλικό / Είδος": item["Υλικό / Είδος"],
                "Υπεύθυνος": item["Υπεύθυνος"],
                "Ώρες": item["Ώρες"],
                "Ώρες/άτομο": item["Ώρες/άτομο"] if item["num_users"] > 1 else ""
            })
        st.dataframe(pd.DataFrame(display_data), use_container_width=True, hide_index=True)
    else:
        st.success("Όλες οι εργασίες έχουν ολοκληρωθεί!")


# --- RENDER DATABASE ---
def render_database(tasks_database, team_database, availability_database):
    st.header("Βάση Δεδομένων")
    col_a, col_b = st.columns(2)
    with col_a:
        st.subheader(f"Πρότυπα Χρόνων ({len(tasks_database)} Εργασίες)")
        tasks_df = pd.DataFrame(list(tasks_database.items()), columns=["Εργασία", "Χρόνος (λεπτά)"])
        st.dataframe(tasks_df, use_container_width=True, hide_index=True)
    with col_b:
        st.subheader("Ομάδα & Όρια Ωρών")
        avail_data = [{"Ημέρα": day, "Τεχνίτης": member, "Ώρες": hours} for day, members in availability_database.items() for member, hours in members.items()]
        st.dataframe(pd.DataFrame(avail_data), use_container_width=True, hide_index=True)


# --- RENDER SETTINGS ---
def render_settings():
    st.header("Ρυθμίσεις")
    st.subheader("Ασφάλεια")
    st.info("Οι ρυθμίσεις ασφαλείας διαχειρίζονται μέσω των Streamlit Secrets")
    st.markdown("**Users:** admin/admin123, manager/manager123, operator/operator123, maria@atrionartgifts.com/atrionmaria")
    
    st.subheader("Αποθήκευση")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Αποθήκευση", use_container_width=True):
            if save_all_assignments_to_sheet():
                st.success("Αποθηκεύτηκε!")
    with col2:
        if st.button("Επαναφόρτωση", use_container_width=True):
            st.cache_data.clear()
            st.session_state.data_version = st.session_state.get("data_version", 0) + 1
            st.rerun()
    
    st.subheader("Audit Log")
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
    procurement_df, tasks_database, team_database, availability_database, incoming_df = load_all_data(version)
    st.session_state.procurement_df = procurement_df
    st.session_state.availability_database = availability_database
    st.session_state.incoming_df = incoming_df
    st.session_state._tasks_db_cache = tasks_database
    
    sheet_item_assignments, sheet_proj_assignments = load_assignments_from_sheet()
    
    if "tasks_store" not in st.session_state:
        st.session_state["tasks_store"] = {}
    else:
        for u_key, task_list in st.session_state["tasks_store"].items():
            unique_tasks = []
            seen = set()
            for t in task_list:
                signature = (t.get("task", ""), t.get("user", ""), str(t.get("date", "")))
                if signature not in seen:
                    seen.add(signature)
                    unique_tasks.append(t)
            st.session_state["tasks_store"][u_key] = unique_tasks
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

    with st.sidebar:
        st.markdown("""
        <div style="text-align:left;padding:5px 0 10px 0;">
            <div style="font-size:18px;font-weight:600;color:#fafafa;">Production Tasks</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown(f"👋 Welcome, **{st.session_state.username}**!")
        
        st.divider()
        
        notifications = check_notifications()
        if notifications:
            st.markdown(f"### 🔔 Ειδοποιήσεις ({len(notifications)})")
            for notif in notifications:
                st.warning(notif)
            st.divider()
        
        st.markdown("### 📋 Navigation")
        pages_list = ["📦 Projects", "📋 Master View", "🗓️ Daily Plan", "👤 Technician", "📆 Projection", "📝 Daily Report", "📊 Database", "⚙️ Settings"]
        current_page = st.session_state.get("page", "📦 Projects")
        if current_page not in pages_list:
            current_page = "📦 Projects"
        page_index = pages_list.index(current_page)
        
        page = st.radio(
            "Select Page",
            pages_list,
            index=page_index,
            label_visibility="collapsed",
            key="sidebar_page_radio"
        )
        st.session_state.page = page
        
        st.divider()
        
        total_tasks = 0
        for u_key, task_list in st.session_state.get("tasks_store", {}).items():
            for t in task_list:
                if t.get("task") and t["task"] != "- Επιλογή Εργασίας -":
                    total_tasks += 1
        for p_key, p_dict in st.session_state.get("project_tasks_store", {}).items():
            if isinstance(p_dict, dict):
                for t_name, p_data in p_dict.items():
                    if isinstance(p_data, dict) and p_data.get("active", False):
                        total_tasks += 1
        st.metric("Total Tasks", total_tasks)
        
        active_count = 0
        if not procurement_df.empty:
            for p_name in procurement_df["Project"].unique():
                if p_name != "-":
                    proj_data = get_project_details(p_name, procurement_df, tasks_database, incoming_df)
                    if proj_data['is_active']:
                        active_count += 1
        st.metric("Active Projects", active_count)
        
        st.divider()
        
        if "last_save" in st.session_state:
            seconds_ago = int((datetime.now() - st.session_state.last_save).seconds)
            if seconds_ago < 60:
                st.caption(f"Αποθηκεύτηκε πριν {seconds_ago}δ")
            else:
                minutes_ago = seconds_ago // 60
                st.caption(f"Αποθηκεύτηκε πριν {minutes_ago}λ")
        
        if st.button("Logout", use_container_width=True):
            logout()

    if (datetime.now() - st.session_state.last_save).seconds > 120:
        total_tasks_in_memory = sum(len(tasks) for tasks in st.session_state.get("tasks_store", {}).values())
        total_proj_tasks_in_memory = sum(
            len([t for t in p_dict.values() if isinstance(t, dict) and t.get("active", False)])
            for p_dict in st.session_state.get("project_tasks_store", {}).values()
            if isinstance(p_dict, dict)
        )
        
        if total_tasks_in_memory > 0 or total_proj_tasks_in_memory > 0:
            if save_all_assignments_to_sheet():
                st.session_state.last_save = datetime.now()

    page = st.session_state.page
    
    if page == "📦 Projects":
        render_projects(procurement_df, tasks_database, team_database, availability_database, incoming_df)
    elif page == "📋 Master View":
        render_master_view(procurement_df, tasks_database, team_database, availability_database, incoming_df)
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
