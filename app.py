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
import calendar

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
    if st.session_state.get("_has_unsaved_changes", False):
        st.warning("⚠️ Έχεις μη αποθηκευμένες αλλαγές! Πάτησε 'Αποθήκευση' πριν βγεις.")
        if st.button("Logout χωρίς αποθήκευση"):
            st.session_state.authenticated = False
            st.session_state.username = None
            st.session_state["_has_unsaved_changes"] = False
            st.rerun()
        if st.button("Ακύρωση"):
            st.rerun()
        return
    
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
TEAM_MEMBERS_GID = "1035290737"
AVAILABILITY_WEEKLY_GID = "1303086311"
AVAILABILITY_OVERRIDES_GID = "1438091482"
PARTIAL_HISTORY_GID = "601876035"
PARTIAL_HISTORY_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={PARTIAL_HISTORY_GID}"

TIMES_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={TIMES_GID}"
TEAM_MEMBERS_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={TEAM_MEMBERS_GID}"
AVAILABILITY_WEEKLY_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={AVAILABILITY_WEEKLY_GID}"
AVAILABILITY_OVERRIDES_CSV_URL = f"https://docs.google.com/spreadsheets/d/{MY_SHEET_ID}/export?format=csv&gid={AVAILABILITY_OVERRIDES_GID}"

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

# --- LOCATIONS ---
LOCATION_MAIN = "Κεντρικό"
LOCATION_WAREHOUSE = "Παράρτημα"
LOCATION_ALL = "Όλα"
WAREHOUSE_KEYWORD = "ΠΑΡΑΡΤΗΜΑ"

def detect_location(project_name):
    if not project_name:
        return LOCATION_MAIN
    name_upper = str(project_name).strip().upper()
    if WAREHOUSE_KEYWORD.upper() in name_upper:
        return LOCATION_WAREHOUSE
    return LOCATION_MAIN

def get_team_members_by_location(team_members_list, location_filter):
    if location_filter == LOCATION_ALL:
        return [m for m in team_members_list if m.get("active", True)]
    return [
        m for m in team_members_list
        if m.get("active", True) and m.get("location", LOCATION_MAIN) == location_filter
    ]

def compute_location_progress(procurement_df, tasks_database, location):
    proj_df = procurement_df[
        procurement_df["Project"].apply(lambda p: detect_location(p) == location)
    ] if not procurement_df.empty else pd.DataFrame()
    
    if proj_df.empty:
        return 0, 0, 0
    
    total_tasks = 0
    done_tasks = 0
    
    for idx, row in proj_df.iterrows():
        item_id = str(row["ID"])
        u_key = f"{item_id}_{idx}"
        item_tasks = st.session_state.get("tasks_store", {}).get(u_key, [])
        for t in item_tasks:
            if t.get("task") and t["task"] != "- Επιλογή Εργασίας -":
                total_tasks += 1
                if t.get("done", False):
                    done_tasks += 1
    
    for p_key, p_tasks_dict in st.session_state.get("project_tasks_store", {}).items():
        if isinstance(p_tasks_dict, dict):
            proj_name = p_key.replace("proj_", "")
            if detect_location(proj_name) != location:
                continue
            for task_name, p_data in p_tasks_dict.items():
                if isinstance(p_data, dict) and p_data.get("active", False):
                    total_tasks += 1
                    if p_data.get("done", False):
                        done_tasks += 1
    
    progress = int((done_tasks / total_tasks) * 100) if total_tasks > 0 else 0
    return done_tasks, total_tasks, progress

def get_location_display(project_name):
    loc = detect_location(project_name)
    if loc == LOCATION_WAREHOUSE:
        return loc, "#ff9800"
    return loc, "#1e88e5"

def get_partial_percentage(task):
    if not task.get("is_partial", False):
        return None
    done = task.get("done_qty", 0)
    total = task.get("total_qty", 0)
    if total == 0:
        return None
    return int((done / total) * 100)

def is_partial_complete(task):
    if not task.get("is_partial", False):
        return False
    done = task.get("done_qty", 0)
    total = task.get("total_qty", 0)
    return done >= total

def get_all_partial_tasks():
    result = []
    for u_key, task_list in st.session_state.get("tasks_store", {}).items():
        for t_idx, t in enumerate(task_list):
            if t.get("is_partial", False) and not is_partial_complete(t):
                item_id = u_key.split("_")[0]
                proj_name = "-"
                if 'procurement_df' in st.session_state:
                    for _, r in st.session_state.procurement_df.iterrows():
                        if str(r["ID"]) == item_id:
                            proj_name = str(r["Project"])
                            break
                result.append({
                    "type": "item",
                    "project": proj_name,
                    "task_name": t.get("task", "?"),
                    "item_id": item_id,
                    "u_key": u_key,
                    "t_idx": t_idx,
                    "done_qty": t.get("done_qty", 0),
                    "total_qty": t.get("total_qty", 0),
                    "reason": t.get("partial_reason", ""),
                    "date": t.get("date", date.today()),
                })
    
    for p_key, p_dict in st.session_state.get("project_tasks_store", {}).items():
        if isinstance(p_dict, dict):
            proj_name = p_key.replace("proj_", "")
            for task_name, p_data in p_dict.items():
                if isinstance(p_data, dict) and p_data.get("is_partial", False) and not is_partial_complete(p_data):
                    result.append({
                        "type": "project",
                        "project": proj_name,
                        "task_name": task_name,
                        "item_id": "-",
                        "p_key": p_key,
                        "task_name_key": task_name,
                        "done_qty": p_data.get("done_qty", 0),
                        "total_qty": p_data.get("total_qty", 0),
                        "reason": p_data.get("partial_reason", ""),
                        "date": p_data.get("date", date.today()),
                    })
    return result

def save_partial_to_history(project, task_name, item_id, done_qty, total_qty, reason):
    gc, _ = get_gspread_client()
    if not gc:
        return False
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Partial_History")
        row = [
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            project,
            task_name,
            item_id,
            done_qty,
            total_qty,
            reason,
        ]
        sheet.append_row(row)
        return True
    except Exception as e:
        st.warning(f"Could not save partial to history: {e}")
        return False

def load_dismissed_warnings():
    gc, _ = get_gspread_client()
    if not gc:
        return set()
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Dismissed_Warnings")
        records = sheet.get_all_records()
        dismissed = set()
        for r in records:
            date_raw = str(r.get("Date", "")).strip()
            user = str(r.get("User", "")).strip()
            if not date_raw or not user:
                continue
            
            date_part = date_raw.split(" ")[0]
            date_part = date_part.replace("-", "/")
            
            parts = date_part.split("/")
            if len(parts) == 3:
                if len(parts[0]) == 4: 
                    final_date = f"{parts[2]}/{parts[1]}/{parts[0]}"
                else:
                    final_date = f"{parts[0]}/{parts[1]}/{parts[2]}"
            else:
                final_date = date_part
                
            dismissed.add(f"{final_date}_{user}")
        return dismissed
    except Exception as e:
        st.error(f"⚠️ ΣΦΑΛΜΑ κατά την ανάγνωση του Dismissed_Warnings: {e}")
        return set()

def dismiss_warning(date_str, user):
    gc, _ = get_gspread_client()
    if not gc:
        return False
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Dismissed_Warnings")
        sheet.append_row([date_str, user, datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
        return True
    except Exception as e:
        st.error(f"Σφάλμα: {e}")
        return False

def update_partial_status_item(u_key, t_idx, done_qty, total_qty, reason):
    if u_key not in st.session_state.get("tasks_store", {}):
        return False
    if t_idx >= len(st.session_state["tasks_store"][u_key]):
        return False
    
    task = st.session_state["tasks_store"][u_key][t_idx]
    task["done_qty"] = done_qty
    task["total_qty"] = total_qty
    task["partial_reason"] = reason
    task["is_partial"] = done_qty < total_qty
    task["done"] = (done_qty >= total_qty)
    mark_unsaved_changes()
    return True

def update_partial_status_project(p_key, task_name, done_qty, total_qty, reason):
    if p_key not in st.session_state.get("project_tasks_store", {}):
        return False
    if task_name not in st.session_state["project_tasks_store"][p_key]:
        return False
    
    task = st.session_state["project_tasks_store"][p_key][task_name]
    task["done_qty"] = done_qty
    task["total_qty"] = total_qty
    task["partial_reason"] = reason
    task["is_partial"] = done_qty < total_qty
    task["done"] = (done_qty >= total_qty)
    mark_unsaved_changes()
    return True

@st.cache_data(ttl=60, show_spinner=False)
def load_partial_history(version=0):
    try:
        df = pd.read_csv(PARTIAL_HISTORY_CSV_URL)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()

# --- HELPERS ---
def clean_project_name(name):
    if name is None:
        return ""
    s = str(name).strip()
    s = re.sub(r'^\*+\s*', '', s)
    s = re.sub(r'\s*\*+$', '', s)
    s = s.replace('**', '')
    return s.strip()

def strip_all_fields(value):
    if value is None:
        return ""
    if not isinstance(value, str):
        return value
    s = value.replace('\xa0', ' ').replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def make_checkbox_key(task):
    if task["type"] == "item":
        return f"mv_chk_item_{task['u_key']}_{task['t_idx']}"
    else:
        return f"mv_chk_proj_{task['p_key']}_{task['task_name']}"

def mark_unsaved_changes():
    st.session_state["_has_unsaved_changes"] = True

def parse_date_safe(date_str):
    s = str(date_str).strip()
    if s in ["-", "nan", "", "None", "NaT"]:
        return None
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"]:
        try:
            return datetime.strptime(s.split(" ")[0], fmt.split(" ")[0]).date()
        except:
            continue
    return None

def parse_date_ambiguous(date_str, prefer_eu=True):
    s = str(date_str).strip()
    if s in ["-", "nan", "", "None", "NaT"]:
        return None
    if prefer_eu:
        try:
            return datetime.strptime(s.split(" ")[0], "%d/%m/%Y").date()
        except:
            pass
    for fmt in ["%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"]:
        try:
            return datetime.strptime(s.split(" ")[0], fmt).date()
        except:
            continue
    return None

def get_shipping_info(project_name, incoming_df):
    if incoming_df is None or incoming_df.empty:
        return False, ""
    if "Project" not in incoming_df.columns:
        return False, ""
    
    project_name_clean = str(project_name).strip().upper()
    
    for _, inc_row in incoming_df.iterrows():
        inc_project = str(inc_row["Project"]).strip().upper()
        if inc_project == project_name_clean:
            shipping_status = str(inc_row.get("Shipping Status", "")).strip()
            shipping_date = str(inc_row.get("Shipping Date", "")).strip()
            
            status_clean = re.sub(r'\s+', ' ', shipping_status.upper()).strip()
            is_shipped = (status_clean == "OK SHIPPED")
            
            if shipping_date in ["-", "nan", "", "None", "NaT", "NAT"]:
                shipping_date = ""
            
            return is_shipped, shipping_date
    return False, ""

def get_availability_for_date(target_date, availability_weekly, availability_overrides, team_members_list):
    if target_date is None:
        return {}
    
    greek_day = WEEKDAYS_GREEK.get(target_date.weekday(), "Δευτέρα")
    day_avail = availability_weekly.get(greek_day, {})
    
    result = {}
    for member in team_members_list:
        if not isinstance(member, dict):
            result[member] = day_avail.get(member, 0.0)
            continue
        
        name = member.get("name", "")
        if not name:
            continue
        
        if not member.get("active", True):
            continue
        
        date_from = member.get("start_date")
        date_to = member.get("end_date")
        
        if date_from and target_date < date_from:
            result[name] = 0.0
            continue
        
        if date_to and target_date > date_to:
            result[name] = 0.0
            continue
        
        result[name] = day_avail.get(name, 0.0)
    
    for override in availability_overrides:
        ov_name = override.get("name", "").strip()
        ov_from = override.get("date_from")
        ov_to = override.get("date_to")
        ov_hours = override.get("hours", 0.0)
        
        if not ov_name or ov_from is None:
            continue
        
        if ov_to is None:
            ov_to = ov_from
        
        if ov_from <= target_date <= ov_to:
            if ov_name in result:
                result[ov_name] = ov_hours
    
    return result

def get_team_options_for_date(target_date, team_members_list, availability_database, availability_overrides,
                              current_location=None):
    if target_date is None:
        target_date = date.today()
    if isinstance(target_date, str):
        target_date = parse_date_safe(target_date)
    if target_date is None:
        target_date = date.today()
    
    day_avail = get_availability_for_date(
        target_date, availability_database, availability_overrides, team_members_list
    )
    
    assigned_hours_per_user = {}
    tasks_db = st.session_state.get("_tasks_db_cache", {})
    procurement_df = st.session_state.get("procurement_df", pd.DataFrame())
    
    for u_key, task_list in st.session_state.get("tasks_store", {}).items():
        for t in task_list:
            if t.get("task") and t["task"] != "- Επιλογή Εργασίας -" and t.get("date") == target_date:
                t_users = t.get("users", [])
                if not t_users:
                    t_users = [t.get("user", "- Χωρίς Ανάθεση -")]
                num_users = max(len(t_users), 1)
                
                qty = 1
                try:
                    item_id = u_key.split("_")[0]
                    if not procurement_df.empty:
                        match = procurement_df[procurement_df["ID"].astype(str) == str(item_id)]
                        if not match.empty:
                            qty_str = str(match.iloc[0]["Ποσότητα"])
                            if qty_str.isdigit():
                                qty = int(qty_str)
                except:
                    pass
                
                auto_time = tasks_db.get(t["task"], 0.0)
                eff_qty = t.get("task_qty", qty)
                try:
                    eff_qty = float(eff_qty) if eff_qty not in [None, ""] else float(qty)
                except:
                    eff_qty = float(qty)
                hrs_per_user = (auto_time * eff_qty / 60) / num_users if auto_time else 0
                for user in t_users:
                    if user and user != "- Χωρίς Ανάθεση -":
                        assigned_hours_per_user[user] = assigned_hours_per_user.get(user, 0.0) + hrs_per_user

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
                    eff_qty = p_data.get("task_qty", proj_qty)
                    try:
                        eff_qty = float(eff_qty) if eff_qty not in [None, ""] else float(proj_qty)
                    except:
                        eff_qty = float(proj_qty)
                    hrs_per_user = (auto_time * eff_qty / 60) / num_users if auto_time else 0
                    for user in t_users:
                        if user and user != "- Χωρίς Ανάθεση -":
                            assigned_hours_per_user[user] = assigned_hours_per_user.get(user, 0.0) + hrs_per_user
    
    display_options = []
    display_to_value = {}
    value_to_display = {}
    
    for member in team_members_list:
        if not member.get("active", True):
            continue
        
        name = member.get("name", "")
        mtype = member.get("type", "")
        m_loc = member.get("location", LOCATION_MAIN)
        
        if not name:
            continue
        
        type_suffix = f" ({mtype})" if mtype and mtype not in ["—", "Άγνωστος"] else ""
        location_suffix = ""
        if current_location and current_location != LOCATION_ALL:
            if m_loc != current_location:
                location_suffix = f" ⚠️ Άλλο Location"
            else:
                location_suffix = f" — {m_loc}"
        else:
            location_suffix = f" — {m_loc}"
        
        has_override = False
        override_reason = ""
        for ov in availability_overrides:
            if ov.get("name") == name:
                ov_from = ov.get("date_from")
                ov_to = ov.get("date_to") or ov_from
                if ov_from and ov_to and ov_from <= target_date <= ov_to:
                    has_override = True
                    override_reason = ov.get("reason", "")
                    break
        
        hrs = day_avail.get(name, 0.0)
        if has_override and hrs == 0:
            reason_str = f" ⚠️ {override_reason}" if override_reason else " ⚠️ Άδεια"
            display = f"{name}{type_suffix}{location_suffix}{reason_str}"
        else:
            display = f"{name}{type_suffix}{location_suffix}"
        
        display_options.append(display)
        display_to_value[display] = name
        value_to_display[name] = display

    st.session_state["assigned_hours_cache"] = assigned_hours_per_user
    return display_options, display_to_value, value_to_display

def clean_users_to_display(clean_list, value_to_display):
    result = []
    for name in clean_list:
        if name in value_to_display:
            result.append(value_to_display[name])
        else:
            result.append(name)
    return result

def update_item_users_mapped(u_key, t_idx, widget_key, display_to_value):
    selected_display = st.session_state[widget_key]
    
    selected_clean = []
    for d in selected_display:
        clean_name = display_to_value.get(d, d)
        if clean_name not in selected_clean:
            selected_clean.append(clean_name)
            
    st.session_state["tasks_store"][u_key][t_idx]["users"] = selected_clean
    if selected_clean:
        st.session_state["tasks_store"][u_key][t_idx]["user"] = selected_clean[0]
    else:
        st.session_state["tasks_store"][u_key][t_idx]["user"] = "- Χωρίς Ανάθεση -"
    mark_unsaved_changes()

def update_proj_users_mapped(p_key, task_name, widget_key, display_to_value):
    selected_display = st.session_state[widget_key]
    
    selected_clean = []
    for d in selected_display:
        clean_name = display_to_value.get(d, d)
        if clean_name not in selected_clean:
            selected_clean.append(clean_name)

    st.session_state["project_tasks_store"][p_key][task_name]["users"] = selected_clean
    if selected_clean:
        st.session_state["project_tasks_store"][p_key][task_name]["user"] = selected_clean[0]
    else:
        st.session_state["project_tasks_store"][p_key][task_name]["user"] = "- Χωρίς Ανάθεση -"
    mark_unsaved_changes()

def apply_bulk_assign_users(tasks, mode, target_user, display_to_value=None):
    changed = 0
    details = []
    
    for t in tasks:
        old_users = t.get("users", [])
        if not isinstance(old_users, list):
            old_users = [old_users] if old_users else []
        
        old_users = [u for u in old_users if u and u != "- Χωρίς Ανάθεση -"]
        
        if mode == "Replace":
            new_users = [target_user]
        elif mode == "Add":
            if target_user in old_users:
                continue
            new_users = old_users + [target_user]
        elif mode == "Remove":
            if target_user not in old_users:
                continue
            new_users = [u for u in old_users if u != target_user]
        else:
            continue
        
        t["users"] = new_users
        if new_users:
            t["user"] = new_users[0]
        else:
            t["user"] = "- Χωρίς Ανάθεση -"
        
        changed += 1
        old_str = ", ".join(old_users) if old_users else "(κανείς)"
        new_str = ", ".join(new_users) if new_users else "(κανείς)"
        details.append(f"{t.get('task', '?')}: {old_str} → {new_str}")
    
    return changed, details

# ============================================================
# === ΝΕΕΣ ΣΥΝΑΡΤΗΣΕΙΣ ΓΙΑ ΤΟ ΠΛΑΝΟ ΕΒΔΟΜΑΔΑΣ ==============
# ============================================================

def get_unassigned_ready_tasks(location_filter):
    """
    Επιστρέφει λίστα με tasks που είναι unassigned ή έχουν παρελθούσα ημερομηνία,
    φιλτραρισμένα ανά location, με flag 'ready' από Procurement status.
    """
    READY_STATES = {"OK STOCK", "RECEIVED", "READY"}
    result = []
    procurement_df = st.session_state.get("procurement_df", pd.DataFrame())
    tasks_db = st.session_state.get("_tasks_db_cache", {})
    today = date.today()

    for u_key, tlist in st.session_state.get("tasks_store", {}).items():
        # Βρες το row του procurement για αυτό το u_key
        try:
            parts = u_key.rsplit("_", 1)
            item_id = parts[0]
            idx_part = int(parts[1])
            if procurement_df.empty or idx_part >= len(procurement_df):
                continue
            row = procurement_df.iloc[idx_part]
            if str(row["ID"]) != item_id:
                continue
        except Exception:
            continue

        project_name = str(row["Project"])
        if location_filter != LOCATION_ALL:
            if detect_location(project_name) != location_filter:
                continue
        if project_name == "-":
            continue

        material = str(row["Υλικό / Προϊόν"])
        status_proc = str(row["Status Procurement"])
        ready = status_proc in READY_STATES
        qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1

        for t_idx, t in enumerate(tlist):
            if not t.get("task") or t["task"] == "- Επιλογή Εργασίας -":
                continue

            users = t.get("users", []) or []
            users_clean = [u for u in users if u and u != "- Χωρίς Ανάθεση -"]
            is_unassigned = len(users_clean) == 0

            task_date = t.get("date")
            is_past = isinstance(task_date, date) and task_date < today

            if not (is_unassigned or is_past):
                continue

            auto_time = tasks_db.get(t["task"], 0.0)
            eff_qty = t.get("task_qty", qty)
            try:
                eff_qty = float(eff_qty) if eff_qty not in [None, ""] else float(qty)
            except Exception:
                eff_qty = float(qty)
            hours = (auto_time * eff_qty) / 60

            due_raw = str(row.get("Αναμενόμενη Ημ. Παραλαβής", "-")).strip()
            due_date = parse_date_safe(due_raw) if due_raw not in ["-", "", "nan"] else None

            result.append({
                "u_key": u_key,
                "t_idx": t_idx,
                "uid": t.get("uid"),
                "project": project_name,
                "material": material,
                "task": t["task"],
                "hours": round(hours, 2),
                "qty": qty,
                "due_date": due_date,
                "ready": ready,
                "status_proc": status_proc,
                "task_date": task_date,
                "is_past": is_past,
            })

    result.sort(key=lambda x: (not x["ready"], x["due_date"] or date(9999, 12, 31)))
    return result


def build_week_plan(week_start, location_filter):
    """
    Επιστρέφει DataFrame με μία γραμμή ανά (άτομο, ημέρα):
      name, location, date, weekday, available_h, assigned_h, tasks, overload
    """
    team_list = st.session_state.get("team_members_list", [])
    avail_weekly = st.session_state.get("availability_weekly", {})
    avail_overrides = st.session_state.get("availability_overrides", [])
    tasks_db = st.session_state.get("_tasks_db_cache", {})
    procurement_df = st.session_state.get("procurement_df", pd.DataFrame())

    members = [m for m in team_list
               if m.get("active", True)
               and (location_filter == LOCATION_ALL
                    or m.get("location", LOCATION_MAIN) == location_filter)]

    days = [week_start + timedelta(days=i) for i in range(5)]  # Δευ-Παρ

    rows = []
    for m in members:
        for d in days:
            avail_map = get_availability_for_date(
                d, avail_weekly, avail_overrides, team_list)
            rows.append({
                "name": m["name"],
                "location": m.get("location", LOCATION_MAIN),
                "date": d,
                "weekday": WEEKDAYS_SHORT_GREEK.get(d.weekday(), ""),
                "available_h": round(avail_map.get(m["name"], 0.0), 2),
                "assigned_h": 0.0,
                "tasks": [],
                "overload": False,
            })

    row_idx = {(r["name"], r["date"]): i for i, r in enumerate(rows)}

    # Item tasks
    for u_key, tlist in st.session_state.get("tasks_store", {}).items():
        qty = 1
        try:
            parts = u_key.rsplit("_", 1)
            idx_part = int(parts[1])
            if not procurement_df.empty and idx_part < len(procurement_df):
                q = str(procurement_df.iloc[idx_part]["Ποσότητα"])
                if q.isdigit():
                    qty = int(q)
        except Exception:
            pass

        for t in tlist:
            if not t.get("task") or t["task"] == "- Επιλογή Εργασίας -":
                continue
            td = t.get("date")
            if not isinstance(td, date):
                continue
            users = t.get("users", []) or []
            users_clean = [u for u in users if u and u != "- Χωρίς Ανάθεση -"]
            if not users_clean:
                continue
            num_users = len(users_clean)
            auto_time = tasks_db.get(t["task"], 0.0)
            eff_qty = t.get("task_qty", qty)
            try:
                eff_qty = float(eff_qty) if eff_qty not in [None, ""] else float(qty)
            except Exception:
                eff_qty = float(qty)
            hpu = (auto_time * eff_qty / 60) / num_users if auto_time else 0.0

            for u in users_clean:
                key = (u, td)
                if key in row_idx:
                    i = row_idx[key]
                    rows[i]["assigned_h"] = round(rows[i]["assigned_h"] + hpu, 2)
                    rows[i]["tasks"].append(f"{t['task'][:18]} ({hpu:.1f}h)")

    # Project-level tasks (PROJ_ keys μέσα στο tasks_store)
    for p_key, tlist in st.session_state.get("tasks_store", {}).items():
        if not p_key.startswith("PROJ_"):
            continue
        proj_name = p_key.replace("PROJ_", "")
        if location_filter != LOCATION_ALL:
            if detect_location(proj_name) != location_filter:
                continue
        main_qty = 1
        if not procurement_df.empty:
            for _, r in procurement_df[procurement_df["Project"] == proj_name].iterrows():
                if str(r["Ποσότητα"]).isdigit():
                    main_qty = max(main_qty, int(r["Ποσότητα"]))
        for t in tlist:
            if not t.get("task"):
                continue
            td = t.get("date")
            if not isinstance(td, date):
                continue
            users = t.get("users", []) or []
            users_clean = [u for u in users if u and u != "- Χωρίς Ανάθεση -"]
            if not users_clean:
                continue
            num_users = len(users_clean)
            auto_time = tasks_db.get(t["task"], 0.0)
            eff_qty = t.get("task_qty", main_qty)
            try:
                eff_qty = float(eff_qty) if eff_qty not in [None, ""] else float(main_qty)
            except Exception:
                eff_qty = float(main_qty)
            hpu = (auto_time * eff_qty / 60) / num_users if auto_time else 0.0
            for u in users_clean:
                key = (u, td)
                if key in row_idx:
                    i = row_idx[key]
                    rows[i]["assigned_h"] = round(rows[i]["assigned_h"] + hpu, 2)
                    rows[i]["tasks"].append(f"{t['task'][:18]} ({hpu:.1f}h)")

    for r in rows:
        r["overload"] = r["assigned_h"] > r["available_h"] and r["available_h"] > 0

    return pd.DataFrame(rows)


def assign_tasks_bulk(task_refs, target_user, target_date):
    """
    task_refs: λίστα από (u_key, t_idx)
    Αναθέτει όλα τα tasks στον target_user για target_date.
    """
    changed = 0
    for u_key, t_idx in task_refs:
        if u_key not in st.session_state.get("tasks_store", {}):
            continue
        if t_idx >= len(st.session_state["tasks_store"][u_key]):
            continue
        t = st.session_state["tasks_store"][u_key][t_idx]
        t["users"] = [target_user]
        t["user"] = target_user
        t["date"] = target_date
        changed += 1
    if changed > 0:
        mark_unsaved_changes()
    return changed

# ============================================================
# === ΤΕΛΟΣ ΝΕΩΝ ΣΥΝΑΡΤΗΣΕΩΝ ==============================
# ============================================================

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
        if "Project" in df_proc.columns:
            df_proc["Project"] = df_proc["Project"].apply(clean_project_name)
        for col in df_proc.select_dtypes(include=['object']).columns:
            df_proc[col] = df_proc[col].apply(strip_all_fields)
    except Exception:
        df_proc = pd.DataFrame()

    df_incoming = pd.DataFrame()
    try:
        df_incoming_raw = pd.read_csv(INCOMING_CSV_URL, header=None)
        if not df_incoming_raw.empty:
            df_incoming = df_incoming_raw.iloc[1:, [1, 8, 11]].copy()
            df_incoming.columns = ["Project", "Shipping Date", "Shipping Status"]
            df_incoming = df_incoming.dropna(subset=["Project"])
            df_incoming["Project"] = df_incoming["Project"].astype(str).str.strip()
            df_incoming["Shipping Date"] = df_incoming["Shipping Date"].astype(str).str.strip()
            df_incoming["Shipping Status"] = df_incoming["Shipping Status"].astype(str).str.strip()
            df_incoming = df_incoming[df_incoming["Project"] != ""]
            df_incoming = df_incoming[df_incoming["Project"] != "nan"]
            df_incoming["Project"] = df_incoming["Project"].apply(clean_project_name)
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

    team_members_list = []
    team_member_names = []
    
    try:
        df_members = pd.read_csv(TEAM_MEMBERS_CSV_URL)
        df_members.columns = [str(c).strip() for c in df_members.columns]
        
        for _, row in df_members.iterrows():
            name = strip_all_fields(str(row.get("Όνομα", "")))
            if not name or name.lower() in ["nan", "none"]:
                continue
            
            member_type = strip_all_fields(str(row.get("Τύπος", "Άγνωστος")))
            active_raw = strip_all_fields(str(row.get("Ενεργός", "ΝΑΙ"))).upper()
            is_active = active_raw in ["ΝΑΙ", "YES", "TRUE", "1", "✅"]
            
            start_str = strip_all_fields(str(row.get("Ημ. Έναρξης", "")))
            end_str = strip_all_fields(str(row.get("Ημ. Λήξης", "")))
            
            start_date = parse_date_ambiguous(start_str, prefer_eu=True) if start_str else None
            end_date = parse_date_ambiguous(end_str, prefer_eu=True) if end_str else None
            
            location = strip_all_fields(str(row.get("Location", LOCATION_MAIN)))
            if not location or location.lower() in ["nan", "none", ""]:
                location = LOCATION_MAIN
            
            team_members_list.append({
                "name": name,
                "type": member_type,
                "active": is_active,
                "location": location,
                "start_date": start_date,
                "end_date": end_date,
            })
            
            if is_active:
                team_member_names.append(name)
                
    except Exception as e:
        st.warning(f"Could not load Team_Members: {e}")
    
    if not team_member_names:
        team_member_names = ["Βαγγέλης Μ.", "Βαγγέλης JR.", "Μαρία"]
        team_members_list = [{"name": n, "type": "Μόνιμος", "active": True, "location": LOCATION_MAIN} for n in team_member_names]

    # 5. Availability Weekly
    availability_dict = {}
    try:
        df_week = pd.read_csv(AVAILABILITY_WEEKLY_CSV_URL)
        df_week.columns = [str(c).strip() for c in df_week.columns]
        
        ignore_cols = ["Ημέρα", "Σύνολο", "ΣΥΝΟΛΟ", "Σύνολο διαθέσιμων ωρών", "Unnamed: 0"]
        member_cols = [c for c in df_week.columns if c and c not in ignore_cols and "Unnamed" not in c]
        
        col_to_clean = {}
        for col in member_cols:
            clean_name = col.split("(")[0].strip()
            col_to_clean[col] = clean_name
        
        for _, row in df_week.iterrows():
            day_name = str(row.iloc[0]).strip()
            if day_name in WEEKDAYS_GREEK.values():
                availability_dict[day_name] = {}
                for member_col in member_cols:
                    try:
                        val = float(str(row[member_col]).replace(',', '.'))
                        if pd.isna(val):
                            val = 0.0
                    except:
                        val = 0.0
                    clean_name = col_to_clean.get(member_col, member_col)
                    availability_dict[day_name][clean_name] = val
    except Exception as e:
        st.warning(f"Could not load Availability_Weekly: {e}")
        for day in WEEKDAYS_GREEK.values():
            availability_dict[day] = {m: 6.0 for m in team_member_names}
    
    for day in WEEKDAYS_GREEK.values():
        if day not in availability_dict:
            availability_dict[day] = {m: 0.0 for m in team_member_names}

    availability_overrides = []
    
    try:
        df_ov = pd.read_csv(AVAILABILITY_OVERRIDES_CSV_URL)
        df_ov.columns = [str(c).strip() for c in df_ov.columns]
        
        for _, row in df_ov.iterrows():
            name = strip_all_fields(str(row.get("Όνομα", "")))
            if not name or name.lower() in ["nan", "none"]:
                continue
            
            date_from_str = strip_all_fields(str(row.get("Ημ. Από", "")))
            date_to_str = strip_all_fields(str(row.get("Ημ. Έως", "")))
            
            date_from = parse_date_ambiguous(date_from_str, prefer_eu=True)
            date_to = parse_date_ambiguous(date_to_str, prefer_eu=True)
            
            if date_from is None:
                continue
            if date_to is None:
                date_to = date_from
            
            try:
                hours = float(str(row.get("Ώρες Override", 0)).replace(',', '.'))
            except:
                hours = 0.0
            
            reason = strip_all_fields(str(row.get("Αιτία", "")))
            
            availability_overrides.append({
                "name": name,
                "date_from": date_from,
                "date_to": date_to,
                "hours": hours,
                "reason": reason,
            })
    except Exception as e:
        pass

    # Εισαγωγή εικονικού υλικού για Γενικές Εργασίες
    if not df_proc.empty:
        projects_list = df_proc["Project"].unique()
        for proj in projects_list:
            if proj == "-":
                continue
            proj_items = df_proc[df_proc["Project"] == proj]
            main_qty = 1
            for _, r in proj_items.iterrows():
                if str(r["Ποσότητα"]).isdigit():
                    main_qty = max(main_qty, int(r["Ποσότητα"]))
            
            virtual_row = {
                "ID": f"VIRTUAL_{proj}",
                "Ημερομηνία Παράδοσης": "-",
                "Είδος Δώρου": "Γενική Σύνθεση",
                "Project": proj,
                "Ποσότητα": main_qty,
                "Προμηθευτής": "-",
                "Υλικό / Προϊόν": "Γενική Σύνθεση / Box",
                "Αναμενόμενη Ημ. Παραλαβής": "-",
                "Αναμενόμενη Ποσότητα Παραλαβής": "-",
                "Status Procurement": "READY",
                "is_virtual": True
            }
            df_proc = pd.concat([df_proc, pd.DataFrame([virtual_row])], ignore_index=True)

    return df_proc, tasks_dict, team_member_names, availability_dict, df_incoming, team_members_list, availability_overrides


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
            p_name = clean_project_name(strip_all_fields(r.get("Project", "")))
            
            raw_id = str(r.get("Item_ID", "")).strip()
            if raw_id.endswith(".0"):
                raw_id = raw_id[:-2]
            item_id = raw_id
            
            task_name = strip_all_fields(str(r.get("Task_Name", "")))
            user_raw = strip_all_fields(str(r.get("Assigned_User", "- Χωρίς Ανάθεση -")))
            user = user_raw
            users_list = []
            for u in user_raw.split(","):
                u_clean = strip_all_fields(u)
                if u_clean and u_clean != "- Χωρίς Ανάθεση -":
                    u_name = u_clean.split(" — ")[0].split(" | ")[0].strip()
                    if u_name:
                        users_list.append(u_name)
            
            assign_date_str = str(r.get("Assigned_Date", r.get("Assigned_Done", ""))).strip()
            done = True if str(r.get("Status_Done", "")).upper() in ["TRUE", "1", "YES"] else False
            task_type = strip_all_fields(str(r.get("Task_Type", "")))

            assign_date = None
            for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"]:
                try:
                    assign_date = datetime.strptime(assign_date_str.split(" ")[0], fmt.split(" ")[0]).date()
                    break
                except:
                    continue
            
            if assign_date is None:
                assign_date = date.today()

            done_qty = r.get("Done_Qty", "")
            total_qty = r.get("Total_Qty", "")
            partial_reason = r.get("Partial_Reason", "")
            
            is_partial = False
            try:
                if done_qty and total_qty:
                    done_qty_val = int(float(str(done_qty).replace(',', '.')))
                    total_qty_val = int(float(str(total_qty).replace(',', '.')))
                    if done_qty_val < total_qty_val:
                        is_partial = True
            except:
                pass

            if task_type == "PROJECT":
                p_store_key = f"PROJ_{p_name}"
                if p_store_key not in assignments_item:
                    assignments_item[p_store_key] = []
                
                assignments_item[p_store_key].append({
                    "uid": f"task_{p_name}_{task_name}_{assign_date_str}_{int(time.time() * 1000) % 1000000}",
                    "project": p_name,
                    "done": done,
                    "task": task_name,
                    "user": users_list[0] if users_list else "- Χωρίς Ανάθεση -",
                    "users": users_list,
                    "date": assign_date,
                    "is_partial": is_partial,
                    "done_qty": done_qty,
                    "total_qty": total_qty,
                    "task_qty": total_qty if total_qty else "",
                    "partial_reason": partial_reason,
                    "progress_logs": [],
                })
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
                        "uid": f"task_{item_id}_{task_name}_{assign_date_str}_{int(time.time() * 1000) % 1000000}",
                        "material": strip_all_fields(str(r.get("Material", ""))),
                        "project": p_name,
                        "done": done,
                        "task": task_name,
                        "user": users_list[0] if users_list else "- Χωρίς Ανάθεση -",
                        "users": users_list,
                        "date": assign_date,
                        "is_partial": is_partial,
                        "done_qty": done_qty,
                        "total_qty": total_qty,
                        "partial_reason": partial_reason,
                        "progress_logs": [],
                    })
                    
    except Exception as e:
        st.warning(f"Could not load assignments: {e}")
    
    return assignments_item, assignments_proj

# === ΤΕΛΟΣ ΜΕΡΟΥΣ 1 ===

def auto_rollover_partial_tasks():
    today = date.today()
    rolled_over_count = 0
    
    for u_key, task_list in st.session_state.get("tasks_store", {}).items():
        for t in task_list:
            if t.get("is_partial", False) and not is_partial_complete(t):
                task_date = t.get("date")
                if task_date and isinstance(task_date, date) and task_date < today:
                    t["date"] = today
                    rolled_over_count += 1
                    
    for p_key, p_dict in st.session_state.get("project_tasks_store", {}).items():
        if isinstance(p_dict, dict):
            for task_name, p_data in p_dict.items():
                if isinstance(p_data, dict) and p_data.get("is_partial", False) and not is_partial_complete(p_data):
                    task_date = p_data.get("date")
                    if task_date and isinstance(task_date, date) and task_date < today:
                        p_data["date"] = today
                        rolled_over_count += 1
                        
    if rolled_over_count > 0:
        mark_unsaved_changes()
    return rolled_over_count

def normalize_item_id(v):
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s

def _norm_txt(v):
    return re.sub(r'\s+', ' ', str(v).strip())

def save_all_assignments_to_sheet():
    gc, _ = get_gspread_client()
    if not gc:
        st.warning("Δεν είναι δυνατή η αποθήκευση λόγω σφάλματος σύνδεσης API.")
        return False
    
    try:
        rolled = auto_rollover_partial_tasks()
        if rolled > 0:
            st.toast(f"🔄 {rolled} partial tasks μεταφέρθηκαν αυτόματα στη σημερινή ημέρα!")
        
        if 'procurement_df' in st.session_state and not st.session_state.procurement_df.empty:
            valid_u_keys = set()
            for idx, row in st.session_state.procurement_df.iterrows():
                item_id = str(row["ID"])
                valid_u_keys.add(f"{item_id}_{idx}")
            
            stale_keys = [k for k in list(st.session_state.get("tasks_store", {}).keys()) if k not in valid_u_keys]
            for k in stale_keys:
                del st.session_state["tasks_store"][k]
    
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

        if 'procurement_df' in st.session_state and not st.session_state.procurement_df.empty:
            latest_u_key_per_item = {}
            for idx, row in st.session_state.procurement_df.iterrows():
                item_id = str(row["ID"])
                u_key = f"{item_id}_{idx}"
                latest_u_key_per_item[item_id] = u_key
            
            cleaned_store = {}
            for u_key, tasks in st.session_state.get("tasks_store", {}).items():
                parts = u_key.rsplit("_", 1)
                item_id_key = parts[0] if len(parts) == 2 else u_key
                if item_id_key in latest_u_key_per_item:
                    if u_key == latest_u_key_per_item[item_id_key]:
                        cleaned_store[u_key] = tasks
            st.session_state["tasks_store"] = cleaned_store

        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Assignments")
        rows = [["Project", "Item_ID", "Task_Name", "Assigned_User", "Assigned_Date", "Status_Done", "Task_Type",
                 "Done_Qty", "Total_Qty", "Partial_Reason", "Material"]]

        item_to_project = {}
        if 'procurement_df' in st.session_state and st.session_state.procurement_df is not None:
            for _, r in st.session_state.procurement_df.iterrows():
                item_to_project[str(r["ID"])] = str(r["Project"])

        for u_key, t_list in st.session_state.get("tasks_store", {}).items():
            item_id = strip_all_fields(u_key.split("_")[0])
            proj_name = strip_all_fields(item_to_project.get(item_id, "-"))
            mat_name = ""
            try:
                _ix = int(u_key.rsplit("_", 1)[1])
                _r = st.session_state.procurement_df.loc[_ix]
                mat_name = strip_all_fields(str(_r["Υλικό / Προϊόν"]))
                proj_name = strip_all_fields(str(_r["Project"]))
            except Exception:
                pass            
            for t in t_list:
                if t.get("task") and t.get("task") != "- Επιλογή Εργασίας -":
                    task_users = t.get("users", [])
                    if not task_users:
                        task_users = [t.get("user", "- Χωρίς Ανάθεση -")]
                    
                    users_str = ", ".join([
                        strip_all_fields(u) for u in task_users
                        if strip_all_fields(u) and strip_all_fields(u) != "- Χωρίς Ανάθεση -"
                    ])
                    if not users_str:
                        users_str = "- Χωρίς Ανάθεση -"

                    partial_done = t.get("done_qty", "")
                    partial_total = t.get("total_qty", "")
                    partial_reason = t.get("partial_reason", "")
                    
                    rows.append([
                        proj_name, item_id, strip_all_fields(t.get("task")), 
                        users_str, str(t.get("date")), str(t.get("done")), "ITEM",
                        partial_done, partial_total, partial_reason, mat_name
                    ])
                    
        # Αποθήκευση Γενικών Εργασιών (PROJ)
        for p_key, p_tasks_list in st.session_state.get("tasks_store", {}).items():
            if p_key.startswith("PROJ_"):
                proj_name = p_key.replace("PROJ_", "")
                if not proj_name:
                    continue
                for t in p_tasks_list:
                    if t.get("task"):
                        task_users = t.get("users", [])
                        if not task_users:
                            task_users = [t.get("user", "- Χωρίς Ανάθεση -")]
                        users_str = ", ".join([strip_all_fields(u) for u in task_users if strip_all_fields(u) and strip_all_fields(u) != "- Χωρίς Ανάθεση -"])
                        if not users_str:
                            users_str = "- Χωρίς Ανάθεση -"

                        rows.append([
                            proj_name, "-", strip_all_fields(t.get("task")),
                            users_str, str(t.get("date")), str(t.get("done")), "PROJECT",
                            t.get("done_qty", ""), t.get("total_qty", ""), t.get("partial_reason", ""), ""
                        ])

        # --- ΔΙΑΤΗΡΗΣΗ ORPHAN ITEM ROWS ---
        try:
            existing_vals = sheet.get_all_values()
            valid_ids_now = {
                normalize_item_id(r["ID"])
                for _, r in st.session_state.procurement_df.iterrows()
            }
            relinked = st.session_state.get("_relinked_rows", set())
            for er in existing_vals[1:]:
                if len(er) < 7 or not any(str(c).strip() for c in er):
                    continue
                if strip_all_fields(str(er[6])).upper() == "PROJECT":
                    continue
                if normalize_item_id(strip_all_fields(str(er[1]))) in valid_ids_now:
                    continue
                sig = (_norm_txt(er[0]), normalize_item_id(er[1]), _norm_txt(er[2]), str(er[4]))
                if sig in relinked:
                    continue
                rows.append(list(er) + [""] * (11 - len(er)))
        except Exception as e:
            st.error(f"Η αποθήκευση ακυρώθηκε, δεν διαβάστηκε το sheet για έλεγχο orphans: {e}")
            return False
        # ----------------------------------        
        
        seen = set()
        deduped_rows = [rows[0]]
        for r in rows[1:]:
            sig = tuple(str(x) for x in r)
            if sig not in seen:
                seen.add(sig)
                deduped_rows.append(r)
        removed_count = len(rows) - len(deduped_rows)
        rows = deduped_rows

        sheet.clear()
        time.sleep(0.5)
        sheet.update(range_name="A1", values=rows)
        time.sleep(0.5)
        st.session_state.last_save = datetime.now()
        
        st.session_state["_has_unsaved_changes"] = False
        load_assignments_from_sheet.clear()
        
        if "audit_log" not in st.session_state:
            st.session_state.audit_log = []
        log_entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "user": st.session_state.get("username", "-"),
            "action": "save",
            "item_tasks": total_tasks,
            "project_tasks": total_proj,
            "duplicates_removed": removed_count
        }
        st.session_state.audit_log.append(log_entry)
        st.session_state.audit_log = st.session_state.audit_log[-200:]
        
        return True
        
    except Exception as e:
        st.error(f"Σφάλμα κατά την αποθήκευση: {e}")
        return False

def cleanup_duplicates_in_sheet():
    gc, err = get_gspread_client()
    if not gc:
        return False, 0, 0, f"Σφάλμα σύνδεσης: {err}"
    
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Assignments")
        all_values = sheet.get_all_values()
        
        if len(all_values) <= 1:
            return True, 0, 0, "Το sheet είναι ήδη κενό (μόνο header)."
        
        header = all_values[0]
        seen = set()
        clean = [header]
        for row in all_values[1:]:
            if not any(str(c).strip() for c in row):
                continue
            normalized = []
            for i, c in enumerate(row):
                if i in (4, 5, 6):
                    normalized.append(str(c).strip())
                else:
                    normalized.append(strip_all_fields(str(c)))
            sig = tuple(normalized)
            if sig not in seen:
                seen.add(sig)
                clean.append(row)
        
        removed = len(all_values) - len(clean)
        total_before = len(all_values) - 1
        total_after = len(clean) - 1
        
        if removed == 0:
            return True, 0, total_before, "Δεν βρέθηκαν διπλότυπα."
        
        sheet.clear()
        time.sleep(0.5)
        sheet.update(range_name="A1", values=clean)
        time.sleep(0.5)
        
        load_assignments_from_sheet.clear()
        
        if "audit_log" not in st.session_state:
            st.session_state.audit_log = []
        st.session_state.audit_log.append({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "user": st.session_state.get("username", "-"),
            "action": "cleanup_duplicates",
            "removed": removed
        })
        
        return True, removed, total_before, f"Καθαρίστηκαν {removed} διπλότυπα (από {total_before} → {total_after} γραμμές)."
    
    except Exception as e:
        return False, 0, 0, f"Σφάλμα: {e}"

def cleanup_orphans_in_sheet(procurement_df):
    gc, err = get_gspread_client()
    if not gc:
        return False, 0, 0, f"Σφάλμα σύνδεσης: {err}"
    
    if procurement_df is None or procurement_df.empty:
        return False, 0, 0, "Δεν υπάρχουν δεδομένα procurement."
    
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Assignments")
        all_values = sheet.get_all_values()
        
        if len(all_values) <= 1:
            return True, 0, 0, "Το sheet είναι ήδη κενό (μόνο header)."
        
        header = all_values[0]
        valid_item_ids = set(str(r["ID"]) for _, r in procurement_df.iterrows())
        valid_projects = set(str(p) for p in procurement_df["Project"].unique() if p != "-")
        
        clean_rows = [header]
        removed = 0
        removed_list = []
        
        for row in all_values[1:]:
            if not any(str(c).strip() for c in row):
                continue
            
            if len(row) < 7:
                removed += 1
                continue
            
            project = strip_all_fields(str(row[0]))
            item_id = strip_all_fields(str(row[1]))
            task_type = strip_all_fields(str(row[6])).upper()
            
            is_orphan = False
            
            if task_type == "PROJECT":
                if project not in valid_projects:
                    is_orphan = True
            else:
                if item_id not in valid_item_ids:
                    is_orphan = True
            
            if is_orphan:
                removed += 1
                removed_list.append(f"{project} — {row[2][:40]}")
            else:
                clean_rows.append(row)
        
        if removed == 0:
            return True, 0, 0, "Δεν βρέθηκαν orphan tasks."
        
        sheet.clear()
        time.sleep(0.5)
        sheet.update(range_name="A1", values=clean_rows)
        time.sleep(0.5)
        
        load_assignments_from_sheet.clear()
        
        if "audit_log" not in st.session_state:
            st.session_state.audit_log = []
        st.session_state.audit_log.append({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "user": st.session_state.get("username", "-"),
            "action": "cleanup_orphans",
            "removed": removed,
            "examples": removed_list[:5]
        })
        
        return True, removed, len(all_values) - 1, f"Καθαρίστηκαν {removed} orphan tasks (από {len(all_values) - 1} γραμμές)"
    
    except Exception as e:
        return False, 0, 0, f"Σφάλμα: {e}"

def strip_spaces_in_sheet():
    gc, err = get_gspread_client()
    if not gc:
        return False, 0, f"Σφάλμα σύνδεσης: {err}"
    
    try:
        sheet = gc.open_by_key(MY_SHEET_ID).worksheet("Assignments")
        all_values = sheet.get_all_values()
        
        if len(all_values) <= 1:
            return True, 0, "Το sheet είναι κενό."
        
        header = all_values[0]
        clean_rows = [header]
        changed = 0
        
        for row in all_values[1:]:
            if not any(str(c).strip() for c in row):
                continue
            
            cleaned_row = []
            for i, cell in enumerate(row):
                if i == 4:
                    cleaned_row.append(str(cell))
                else:
                    cleaned_row.append(strip_all_fields(str(cell)))
            
            if cleaned_row != row:
                changed += 1
            clean_rows.append(cleaned_row)
        
        if changed == 0:
            return True, 0, "Δεν βρέθηκαν κενά για καθαρισμό."
        
        sheet.clear()
        time.sleep(0.5)
        sheet.update(range_name="A1", values=clean_rows)
        time.sleep(0.5)
        
        load_assignments_from_sheet.clear()
        
        if "audit_log" not in st.session_state:
            st.session_state.audit_log = []
        st.session_state.audit_log.append({
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "user": st.session_state.get("username", "-"),
            "action": "strip_spaces",
            "changed": changed
        })
        
        return True, changed, f"Καθαρίστηκαν κενά σε {changed} γραμμές."
    
    except Exception as e:
        return False, 0, f"Σφάλμα: {e}"

# --- EXPORT FUNCTIONS ---
def generate_printable_html(title, date_str, df_data):
    total_rows = len(df_data)
    
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
    
    table_rows = ""
    for idx, row in df_data.iterrows():
        table_rows += "<tr>"
        for col in df_data.columns:
            val = str(row[col])
            cell_class = ""
            
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
    
    html = f"""<!DOCTYPE html>
    <html lang="el">
    <head>
        <meta charset="utf-8">
        <title>{title}</title>
        <style>
            body {{ font-family: 'Segoe UI', Arial, sans-serif; padding: 30px; background: #f5f7fa; }}
            .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 40px; border-radius: 8px; }}
            h1 {{ font-size: 28px; color: #2c3e50; margin-bottom: 20px; }}
            .summary-cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 15px; margin: 30px 0; }}
            .summary-card {{ padding: 18px; border-radius: 8px; background: #f8f9fa; text-align: center; }}
            .summary-card.success {{ background: #e8f5e9; }}
            .summary-card.warning {{ background: #fff8e1; }}
            .summary-card.info {{ background: #e3f2fd; }}
            .summary-value {{ font-size: 26px; font-weight: 700; }}
            .summary-label {{ font-size: 12px; color: #7f8c8d; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th {{ background: #1e88e5; color: white; padding: 12px 15px; text-align: left; }}
            td {{ padding: 10px 15px; border-bottom: 1px solid #ecf0f1; }}
            .status-done {{ color: #2e7d32; font-weight: 700; }}
            .status-pending {{ color: #d32f2f; font-weight: 700; }}
            .btn {{ padding: 12px 24px; background: #1e88e5; color: white; border: none; border-radius: 6px; cursor: pointer; margin: 20px 5px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>{title}</h1>
            <p>📅 {date_str} &nbsp;•&nbsp; 📊 {total_rows} εγγραφές</p>
            {summary_html}
            <table>
                <thead><tr>{"".join([f"<th>{col}</th>" for col in df_data.columns])}</tr></thead>
                <tbody>{table_rows}</tbody>
            </table>
            <div style="text-align:center;">
                <button onclick="window.print()" class="btn">🖨️ Εκτύπωση</button>
                <button onclick="window.close()" class="btn">✕ Κλείσιμο</button>
            </div>
        </div>
    </body>
    </html>"""
    return html

def generate_printable_html_two_sections(title, date_str, pending_df, completed_df):
    pending_count = len(pending_df)
    completed_count = len(completed_df)
    total_count = pending_count + completed_count
    
    if pending_count > 0:
        columns = list(pending_df.columns)
    elif completed_count > 0:
        columns = list(completed_df.columns)
    else:
        columns = ["Project", "Υλικό / Είδος", "Ποσότητα", "Εργασία", "Ώρες", "Status Procurement"]
    
    col_widths = {"Project": 13, "Υλικό / Είδος": 30, "Ποσότητα": 8, "Εργασία": 20, "Ώρες": 7, "Status Procurement": 17}
    DEFAULT_W = 15
    weights = [col_widths.get(c, DEFAULT_W) for c in columns]
    weight_sum = sum(weights) if sum(weights) > 0 else 1
    data_total = 95.0
    
    colgroup_html = '<colgroup><col style="width:5%;">'
    for w in weights:
        pct = (w / weight_sum) * data_total
        colgroup_html += f'<col style="width:{pct:.2f}%;">'
    colgroup_html += '</colgroup>'
    
    header_cells = '<th class="checkbox-col">✓</th>'
    for col in columns:
        header_cells += f'<th>{col}</th>'
    
    body_rows = ""
    body_rows += f'<tr class="section-row section-row-pending"><td colspan="{len(columns) + 1}">Εκκρεμή ({pending_count})</td></tr>'
    
    if pending_count > 0:
        for _, row in pending_df.iterrows():
            body_rows += '<tr class="row-pending">'
            body_rows += '<td class="checkbox-cell">☐</td>'
            for col in columns:
                body_rows += f'<td>{str(row[col])}</td>'
            body_rows += '</tr>'
    else:
        body_rows += f'<tr class="empty-row"><td colspan="{len(columns) + 1}">Δεν υπάρχουν εκκρεμείς εργασίες.</td></tr>'
    
    body_rows += f'<tr class="section-row section-row-completed"><td colspan="{len(columns) + 1}">Ολοκληρωμένα ({completed_count})</td></tr>'
    
    if completed_count > 0:
        for _, row in completed_df.iterrows():
            body_rows += '<tr class="row-completed">'
            body_rows += '<td class="checkbox-cell empty-check">—</td>'
            for col in columns:
                body_rows += f'<td>{str(row[col])}</td>'
            body_rows += '</tr>'
    else:
        body_rows += f'<tr class="empty-row"><td colspan="{len(columns) + 1}">Δεν υπάρχουν ολοκληρωμένες εργασίες.</td></tr>'
    
    html = f"""<!DOCTYPE html>
    <html lang="el">
    <head>
        <meta charset="utf-8">
        <title>{title}</title>
        <style>
            body {{ font-family: Arial, sans-serif; padding: 40px; }}
            .container {{ max-width: 1100px; margin: 0 auto; }}
            h1 {{ font-size: 20px; margin-bottom: 6px; }}
            table.report-table {{ width: 100%; border-collapse: collapse; font-size: 12px; table-layout: fixed; }}
            th {{ padding: 10px 8px; text-align: left; font-size: 10px; background: #f5f5f5; border-bottom: 2px solid #ddd; }}
            td {{ padding: 10px 8px; border-bottom: 1px solid #eee; font-size: 11.5px; word-wrap: break-word; }}
            .section-row-pending td {{ background: #fff3d6; color: #b8860b; font-weight: 700; text-transform: uppercase; padding: 8px 10px; }}
            .section-row-completed td {{ background: #e6f4ea; color: #3d7a4e; font-weight: 700; text-transform: uppercase; padding: 8px 10px; }}
            .row-pending td {{ background: #fffdf7; }}
            .row-completed td {{ background: #fafdfb; }}
            .empty-row td {{ padding: 15px; text-align: center; color: #999; font-style: italic; }}
            th.checkbox-col, td.checkbox-cell {{ text-align: center; }}
            .actions {{ margin-top: 30px; text-align: center; }}
            .btn {{ padding: 10px 24px; background: #333; color: white; border: none; border-radius: 4px; cursor: pointer; margin: 0 5px; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>{title}</h1>
            <p>{date_str} &nbsp;•&nbsp; {total_count} εργασίες</p>
            <table class="report-table">
                {colgroup_html}
                <thead><tr>{header_cells}</tr></thead>
                <tbody>{body_rows}</tbody>
            </table>
            <div class="actions">
                <button onclick="window.print()" class="btn">Εκτύπωση</button>
                <button onclick="window.close()" class="btn">Κλείσιμο</button>
            </div>
        </div>
    </body>
    </html>"""
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

# --- NOTIFICATIONS & TOGGLES ---
def check_notifications():
    notifications = []
    if 'procurement_df' in st.session_state and not st.session_state.procurement_df.empty:
        pending = st.session_state.procurement_df[
            ~st.session_state.procurement_df["Status Procurement"].isin(["OK STOCK", "RECEIVED", "READY"])
        ]
        if not pending.empty:
            notifications.append(f"{len(pending)} υλικά σε εκκρεμότητα procurement")
    return notifications

def toggle_project_task(p_key, task_name, chk_key):
    st.session_state["project_tasks_store"][p_key][task_name]["done"] = st.session_state[chk_key]
    mark_unsaved_changes()

def toggle_item_task(u_key, t_idx, chk_key):
    st.session_state["tasks_store"][u_key][t_idx]["done"] = st.session_state[chk_key]
    mark_unsaved_changes()

def ensure_task_uids():
    for u_key, task_list in st.session_state.get("tasks_store", {}).items():
        for i, t in enumerate(task_list):
            if not t.get("uid"):
                t["uid"] = f"task_{u_key}_{i}_{int(time.time() * 1000) % 1000000}"
            if "task_qty" not in t:
                try:
                    t["task_qty"] = int(t.get("total_qty", 1))
                except:
                    t["task_qty"] = 1

def delete_item_task_by_uid(u_key, task_uid):
    if u_key not in st.session_state.get("tasks_store", {}):
        return False
    tasks = st.session_state["tasks_store"][u_key]
    for i, t in enumerate(tasks):
        if t.get("uid") == task_uid:
            tasks.pop(i)
            mark_unsaved_changes()
            return True
    return False

def delete_project_task_by_name(p_key, task_name):
    if p_key in st.session_state.get("project_tasks_store", {}):
        if task_name in st.session_state["project_tasks_store"][p_key]:
            del st.session_state["project_tasks_store"][p_key][task_name]
            mark_unsaved_changes()
            return True
    return False

def update_item_field(u_key, t_idx, field, widget_key):
    st.session_state["tasks_store"][u_key][t_idx][field] = st.session_state[widget_key]
    mark_unsaved_changes()

def update_proj_field(p_key, task_name, field, widget_key):
    if p_key in st.session_state.get("project_tasks_store", {}):
        if task_name in st.session_state["project_tasks_store"][p_key]:
            task = st.session_state["project_tasks_store"][p_key][task_name]
            
            if field == "task_qty":
                task["task_qty"] = st.session_state[widget_key]
                task["total_qty"] = st.session_state[widget_key]
            else:
                task[field] = st.session_state[widget_key]
                
            mark_unsaved_changes()

# --- CHECK USER OVERLOAD ---
def check_user_overload(target_date, availability_database, availability_overrides=None, team_members_list=None):
    if not target_date:
        return {}
    
    if availability_overrides is None:
        availability_overrides = st.session_state.get("availability_overrides", [])
    if team_members_list is None:
        team_members_list = st.session_state.get("team_members_list", [])
    
    day_availability = get_availability_for_date(
        target_date, availability_database, availability_overrides, team_members_list
    )
    
    tasks_db = st.session_state.get("_tasks_db_cache", {})
    procurement_df = st.session_state.get("procurement_df", pd.DataFrame())
    
    user_hours = {}
    
    # 1. Item Tasks
    for u_key, task_list in st.session_state.get("tasks_store", {}).items():
        for t in task_list:
            if t.get("task") and t["task"] != "- Επιλογή Εργασίας -" and t.get("date") == target_date:
                t_users = t.get("users", [])
                if not t_users:
                    t_users = [t.get("user", "- Χωρίς Ανάθεση -")]
                num_users = max(len(t_users), 1)
                
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
                eff_qty = t.get("task_qty", qty)
                
                try:
                    eff_qty = float(eff_qty) if eff_qty not in [None, ""] else float(qty)
                except:
                    eff_qty = float(qty)
                    
                hrs_per_user = (auto_time * eff_qty / 60) / num_users if auto_time else 0
                
                for user in t_users:
                    if user and user != "- Χωρίς Ανάθεση -":
                        user_hours[user] = user_hours.get(user, 0.0) + hrs_per_user
    
    # 2. Project Tasks
    for p_key, p_tasks_list in st.session_state.get("tasks_store", {}).items():
        if p_key.startswith("PROJ_"):
            for t_data in p_tasks_list:
                if t_data.get("task") and t_data.get("date") == target_date:
                    t_users = t_data.get("users", [])
                    if not t_users:
                        t_users = [t_data.get("user", "- Χωρίς Ανάθεση -")]
                    num_users = max(len(t_users), 1)
                    auto_time = tasks_db.get(t_data["task"], 0.0)
                    eff_qty = t_data.get("task_qty", 1)
                    
                    try:
                        eff_qty = float(eff_qty) if eff_qty not in [None, ""] else 1.0
                    except:
                        eff_qty = 1.0
                        
                    hrs_per_user = (auto_time * eff_qty / 60) / num_users if auto_time else 0
                    
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
    if not target_dates:
        return
    
    availability_overrides = st.session_state.get("availability_overrides", [])
    team_members_list = st.session_state.get("team_members_list", [])
    
    dismissed_warnings = load_dismissed_warnings()
    
    all_warnings = {}
    for check_date in sorted(target_dates):
        overloads = check_user_overload(check_date, availability_database, availability_overrides, team_members_list)
        for user, info in overloads.items():
            date_str = check_date.strftime("%d/%m/%Y")
            warn_key = f"{date_str}_{user}"
            
            if warn_key in dismissed_warnings:
                continue
                
            if check_date not in all_warnings:
                all_warnings[check_date] = []
            all_warnings[check_date].append((user, info, warn_key))
    
    if not all_warnings:
        return

    for check_date in sorted(all_warnings.keys()):
        for user, info, warn_key in all_warnings[check_date]:
            col_msg, col_btn = st.columns([10, 1])
            
            with col_msg:
                st.markdown(
                    f'<div style="background: #b71c1c; color: white; padding: 12px 16px; '
                    f'border-radius: 8px; margin-bottom: 8px; border-left: 5px solid #ff5252;">'
                    f'<b>⚠️ Υπέρβαση ωρών — {user}</b> στις <b>{check_date.strftime("%d/%m/%Y")}</b>: '
                    f'<b>{info["assigned"]}h</b> (όριο: {info["max"]}h) — <b>+{info["over"]}h</b>'
                    f'</div>',
                    unsafe_allow_html=True
                )
            
            with col_btn:
                if st.button("✕", key=f"dismiss_{warn_key}", help="Διαγραφή ειδοποίησης"):
                    date_str = check_date.strftime("%d/%m/%Y")
                    if dismiss_warning(date_str, user):
                        st.rerun()

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
                eff_qty = t.get("task_qty", qty)
                try:
                    eff_qty = float(eff_qty) if eff_qty not in [None, ""] else float(qty)
                except:
                    eff_qty = float(qty)
                hrs = (auto_t * eff_qty) / 60
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
                eff_qty = p_data.get("task_qty", main_qty)
                try:
                    eff_qty = float(eff_qty) if eff_qty not in [None, ""] else float(main_qty)
                except:
                    eff_qty = float(main_qty)
                hrs = (auto_t * eff_qty) / 60
                total_hours += hrs
                total_tasks += 1
                if p_data.get("done", False):
                    completed_hours += hrs
                    completed_tasks += 1
    
    progress = int((completed_tasks / total_tasks) * 100) if total_tasks > 0 else 0
    is_shipped, shipping_date_str = get_shipping_info(project_name, incoming_df)
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
        "shipping_date": shipping_date_str,
        "project_tasks": project_tasks
    }

# === ΤΕΛΟΣ ΜΕΡΟΥΣ 2 ===

def get_expected_qty(row, fallback_qty):
    try:
        expected = str(row.get("Αναμενόμενη Ποσότητα Παραλαβής", "")).strip()
        if expected and expected not in ["-", "nan", ""]:
            val = float(expected.replace(',', '.'))
            if val > 0:
                return int(val)
    except Exception:
        pass
    return fallback_qty


def render_partial_dialog(u_key, t_idx, task_type, expected_qty, current_done, current_reason, project_name, task_name):
    state_k = f"_partial_dialog_{task_type}_{u_key}_{t_idx}"
    if not st.session_state.get(state_k, False):
        return
    
    with st.container(border=True):
        st.markdown(f"**🔄 Μερική Ολοκλήρωση: {task_name}**")
        col_done, col_total, col_reason = st.columns([1, 1, 2])
        
        done_val = col_done.number_input(
            "Ολοκληρωμένα τμχ:",
            min_value=0,
            max_value=int(expected_qty),
            value=int(current_done) if current_done else 0,
            step=1,
            key=f"_pd_done_{task_type}_{u_key}_{t_idx}",
        )
        total_val = col_total.number_input(
            "Σύνολο τμχ:",
            min_value=1,
            value=int(expected_qty),
            step=1,
            key=f"_pd_total_{task_type}_{u_key}_{t_idx}",
        )
        reason_val = col_reason.text_input(
            "Αιτία (προαιρετικό):",
            value=current_reason or "",
            key=f"_pd_reason_{task_type}_{u_key}_{t_idx}",
        )
        
        col_save, col_cancel = st.columns([1, 1])
        with col_save:
            if st.button("💾 Αποθήκευση", key=f"_pd_save_{task_type}_{u_key}_{t_idx}",
                         type="primary", use_container_width=True):
                if task_type == "item":
                    update_partial_status_item(u_key, t_idx, done_val, total_val, reason_val)
                else:
                    update_partial_status_project(u_key, task_name, done_val, total_val, reason_val)
                
                save_partial_to_history(
                    project_name, task_name,
                    u_key.split("_")[0] if task_type == "item" else "-",
                    done_val, total_val, reason_val
                )
                
                st.session_state[state_k] = False
                st.rerun()
        with col_cancel:
            if st.button("✕ Ακύρωση", key=f"_pd_cancel_{task_type}_{u_key}_{t_idx}",
                         use_container_width=True):
                st.session_state[state_k] = False
                st.rerun()

# --- RENDER PROJECTS (ΔΙΟΡΘΩΜΕΝΟ) ---
def render_projects(procurement_df, tasks_database, team_database, availability_database, incoming_df):
    st.header("Projects")

    dashboard_data = []
    projects_details_cache = {}
    tot_all_hours = 0.0
    tot_tasks_count = 0
    tot_done_tasks = 0

    team_members_list = st.session_state.get("team_members_list", [])
    availability_overrides = st.session_state.get("availability_overrides", [])
    current_location = st.session_state.get("_location_filter", LOCATION_ALL)
    
    if current_location == LOCATION_ALL:
        procurement_filtered = procurement_df.copy()
    else:
        procurement_filtered = procurement_df[
            procurement_df["Project"].apply(lambda p: detect_location(p) == current_location)
        ].copy()
    
    if procurement_filtered.empty:
        st.info(f"Δεν υπάρχουν projects για location: **{current_location}**")
        return
    
    procurement_df = procurement_filtered
    
    if "selected_project_drill" not in st.session_state:
        st.session_state.selected_project_drill = None
    if "selected_material_expand" not in st.session_state:
        st.session_state.selected_material_expand = None

    # --- Drill-down προβολή ---
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
            if proj_details.get("shipping_date"):
                st.caption(f"📦 Αποστολή: **{proj_details['shipping_date']}**")
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
        tab1 = st.tabs(["Υλικά"])[0]        
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
                        
                        if not item_tasks:
                            st.info("Δεν έχουν οριστεί εργασίες. Πάτησε **Προσθήκη Εργασίας** παρακάτω.")
                        
                        for t_idx, t_data in enumerate(list(item_tasks)):
                            task_uid = t_data.get("uid", f"task_{t_idx}")
                            
                            with st.container(border=True):
                                c_done, c_task, c_partial = st.columns([0.1, 0.8, 0.1])
                                
                                chk_k = f"mat_chk_{unique_key}_{task_uid}"
                                c_done.checkbox("Done", value=t_data["done"], key=chk_k, on_change=toggle_item_task, args=(unique_key, t_idx, chk_k))
                                
                                is_partial = t_data.get("is_partial", False)
                                partial_icon = "🔄" if is_partial else "⏳"
                                partial_help = "Επεξεργασία partial" if is_partial else "Δήλωση μερικής ολοκλήρωσης"
                                if c_partial.button(partial_icon, key=f"mat_partial_{unique_key}_{task_uid}", 
                                                   help=partial_help, use_container_width=True):
                                    state_k = f"_partial_dialog_item_{unique_key}_{t_idx}"
                                    st.session_state[state_k] = not st.session_state.get(state_k, False)
                                    st.rerun()
                                
                                task_idx = task_options.index(t_data["task"]) if t_data["task"] in task_options else 0
                                task_k = f"mat_task_{unique_key}_{task_uid}"
                                c_task.selectbox("Εργασία", task_options, index=task_idx, key=task_k, on_change=update_item_field, args=(unique_key, t_idx, "task", task_k))
                               
                                if t_data.get("is_partial", False):
                                    done_q = t_data.get("done_qty", 0)
                                    total_q = t_data.get("total_qty", 0)
                                    
                                    try:
                                        done_q = int(float(str(done_q).replace(',', '.'))) if done_q != "" else 0
                                    except:
                                        done_q = 0
                                    try:
                                        total_q = int(float(str(total_q).replace(',', '.'))) if total_q != "" else 0
                                    except:
                                        total_q = 0
                                    
                                    if total_q > 0:
                                        pct = int((done_q / total_q) * 100)
                                        reason_text = f' · 📝 {t_data.get("partial_reason", "")[:40]}' if t_data.get("partial_reason") else ''
                                        st.markdown(
                                            f'<div style="background:#263238;padding:6px 10px;'
                                            f'border-radius:5px;margin:5px 0;font-size:11px;border-left:3px solid #ff9800;">'
                                            f'<span style="color:#ffb74d !important;font-weight:600;">'
                                            f'⏳ {done_q}/{total_q} ({pct}%){reason_text}'
                                            f'</span>'
                                            f'</div>',
                                            unsafe_allow_html=True
                                        )
                                        st.progress(pct / 100)
                                
                                if t_data.get("task") and t_data["task"] != "- Επιλογή Εργασίας -":
                                    render_partial_dialog(
                                        unique_key, t_idx, "item",
                                        get_expected_qty(row, qty), t_data.get("done_qty", 0),
                                        t_data.get("partial_reason", ""),
                                        selected_project, t_data["task"]
                                    )
                                
                                c_user, c_date, c_time, c_del = st.columns([0.4, 0.3, 0.2, 0.1])
                                
                                current_users = t_data.get("users", [])
                                if not isinstance(current_users, list):
                                    current_users = [current_users] if current_users else []
                                
                                task_date = t_data.get("date", date.today())
                                
                                display_opts, disp_to_val, val_to_disp = get_team_options_for_date(
                                    task_date, team_members_list, availability_database, availability_overrides,
                                    current_location=current_location
                                )
                                
                                valid_users_clean = [u for u in current_users if u in team_database]
                                current_display = clean_users_to_display(valid_users_clean, val_to_disp)
                                
                                user_k = f"mat_users_{unique_key}_{task_uid}"
                                
                                current_display = [val_to_disp.get(u, u) for u in valid_users_clean if u in val_to_disp]
                                
                                c_user.multiselect(
                                    "Υπεύθυνοι",
                                    options=display_opts,
                                    default=current_display,
                                    key=user_k,
                                    on_change=update_item_users_mapped,
                                    args=(unique_key, t_idx, user_k, disp_to_val)
                                )
                                
                                with c_user.expander("📊 Δες Διαθεσιμότητα Ομάδας", expanded=False):
                                    current_task_date = st.session_state.get(f"mat_date_{unique_key}_{task_uid}", task_date)
                                    if current_task_date is None:
                                        current_task_date = task_date
                                    
                                    day_avail_map = get_availability_for_date(current_task_date, availability_database, availability_overrides, team_members_list)
                                    
                                    temp_assigned = {}
                                    
                                    for uk, tlist in st.session_state.get("tasks_store", {}).items():
                                        for ti, tt in enumerate(tlist):
                                            if tt.get("task") and tt["task"] != "- Επιλογή Εργασίας -" and tt.get("date") == current_task_date:
                                                tut = st.session_state.get(f"mat_users_{uk}_{tt.get('uid')}", tt.get("users", []))
                                                if not tut: tut = [tt.get("user", "- Χωρίς Ανάθεση -")]
                                                nu = max(len(tut), 1)
                                                
                                                q = 1
                                                try:
                                                    iid = uk.split("_")[0]
                                                    if not procurement_df.empty:
                                                        match = procurement_df[procurement_df["ID"].astype(str) == str(iid)]
                                                        if not match.empty:
                                                            q_str = str(match.iloc[0]["Ποσότητα"])
                                                            if q_str.isdigit(): q = int(q_str)
                                                except: pass
                                                
                                                at = tasks_database.get(tt["task"], 0.0)
                                                eff_q = tt.get("task_qty", q)
                                                try:
                                                    eff_q = float(eff_q) if eff_q not in [None, ""] else float(q)
                                                except:
                                                    eff_q = float(q)
                                                hpu = (at * eff_q / 60) / nu if at else 0
                                                for uu in tut:
                                                    if uu and uu != "- Χωρίς Ανάθεση -":
                                                        temp_assigned[uu] = temp_assigned.get(uu, 0.0) + hpu
                                    
                                    for pk, pd in st.session_state.get("project_tasks_store", {}).items():
                                        if isinstance(pd, dict):
                                            pn = pk.replace("proj_", "")
                                            pq = 1
                                            if not procurement_df.empty:
                                                pi = procurement_df[procurement_df["Project"] == pn]
                                                for _, pr in pi.iterrows():
                                                    if str(pr["Ποσότητα"]).isdigit():
                                                        pq = max(pq, int(pr["Ποσότητα"]))
                                            
                                            for tn, pdata in pd.items():
                                                if isinstance(pdata, dict) and pdata.get("active", False) and pdata.get("date") == current_task_date:
                                                    tut_p = st.session_state.get(f"puser_proj_{pk}_{tn}", pdata.get("users", []))
                                                    if not tut_p: tut_p = [pdata.get("user", "- Χωρίς Ανάθεση -")]
                                                    nu_p = max(len(tut_p), 1)
                                                    at_p = tasks_database.get(tn, 0.0)
                                                    eff_q_p = pdata.get("task_qty", pq)
                                                    try:
                                                        eff_q_p = float(eff_q_p) if eff_q_p not in [None, ""] else float(pq)
                                                    except:
                                                        eff_q_p = float(pq)
                                                    hpu_p = (at_p * eff_q_p / 60) / nu_p if at_p else 0
                                                    for uu in tut_p:
                                                        if uu and uu != "- Χωρίς Ανάθεση -":
                                                            temp_assigned[uu] = temp_assigned.get(uu, 0.0) + hpu_p

                                    for d_opt in display_opts:
                                        clean_n = disp_to_val.get(d_opt)
                                        if not clean_n: continue
                                        
                                        assigned_h = temp_assigned.get(clean_n, 0.0)
                                        avail_h = day_avail_map.get(clean_n, 0.0)
                                        
                                        if assigned_h > avail_h:
                                            h_color = "#ff5252"
                                            h_icon = "🔴"
                                        elif assigned_h > 0:
                                            h_color = "#69f0ae"
                                            h_icon = "🟢"
                                        else:
                                            h_color = "#b0bec5"
                                            h_icon = "⚪"
                                            
                                        st.markdown(
                                            f'<div style="display:flex; justify-content:space-between; font-size:12px; margin-bottom:4px;">'
                                            f'<span>{d_opt}</span>'
                                            f'<span style="color:{h_color}; font-weight:bold;">{h_icon} {assigned_h:.1f}h / {avail_h:.1f}h</span>'
                                            f'</div>',
                                            unsafe_allow_html=True
                                        )
                                
                                date_k = f"mat_date_{unique_key}_{task_uid}"
                                c_date.date_input("Ημερομηνία", value=t_data["date"], format="DD/MM/YYYY", key=date_k, on_change=update_item_field, args=(unique_key, t_idx, "date", date_k))
                                
                                auto_time = tasks_database.get(t_data["task"], 0.0)
                                if t_data["task"] != "- Επιλογή Εργασίας -":
                                    num_users = max(len(t_data.get("users", [])), 1)
                                    effective_qty = t_data.get("task_qty", qty)
                                    try:
                                        effective_qty = float(effective_qty) if effective_qty not in [None, ""] else float(qty)
                                    except:
                                        effective_qty = float(qty)
                                    task_hours_total = (auto_time * effective_qty) / 60
                                    task_hours_per_user = task_hours_total / num_users
                                    if num_users > 1:
                                        c_time.metric("Ώρες/άτομο", f"{round(task_hours_per_user, 2)}h")
                                    else:
                                        c_time.metric("Ώρες", f"{round(task_hours_per_user, 2)}h")
                                else:
                                    c_time.caption("—")
                                
                                if c_del.button("🗑️", key=f"mat_del_{unique_key}_{t_idx}"):
                                    if unique_key in st.session_state["tasks_store"]:
                                        if t_idx < len(st.session_state["tasks_store"][unique_key]):
                                            st.session_state["tasks_store"][unique_key].pop(t_idx)
                                            mark_unsaved_changes()
                                            st.rerun()
                        
                        col_add, col_rem, col_close = st.columns([1, 1, 1])
                        if col_add.button("Προσθήκη Εργασίας", key=f"mat_add_{unique_key}", use_container_width=True):
                            st.session_state[f"_add_task_form_{unique_key}"] = True

                        if st.session_state.get(f"_add_task_form_{unique_key}", False):
                            with st.container(border=True):
                                st.markdown("**➕ Νέα Εργασία**")
                                col_new_task, col_new_qty = st.columns([3, 1])
                                
                                task_options_new = ["- Επιλογή Εργασίας -"] + sorted(list(tasks_database.keys()))
                                new_task_name = col_new_task.selectbox("Εργασία:", task_options_new, key=f"_new_task_name_{unique_key}")
                                
                                new_task_qty = col_new_qty.number_input("Ποσότητα:", min_value=1, value=int(qty), step=1, key=f"_new_task_qty_{unique_key}")
                                
                                col_save_new, col_cancel_new = st.columns([1, 1])
                                with col_save_new:
                                    if st.button("💾 Δημιουργία", key=f"_save_new_task_{unique_key}", type="primary", use_container_width=True):
                                        if new_task_name == "- Επιλογή Εργασίας -":
                                            st.warning("⚠️ Διάλεξε εργασία.")
                                        else:
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
                                            
                                            new_uid = f"task_{datetime.now().timestamp()}"
                                            st.session_state["tasks_store"][unique_key].append({
                                                "uid": new_uid,
                                                "done": False, "task": new_task_name, 
                                                "user": "- Χωρίς Ανάθεση -", "users": [], "date": auto_date,
                                                "is_partial": False, "done_qty": 0, "total_qty": new_task_qty, 
                                                "task_qty": new_task_qty,
                                                "partial_reason": "", "progress_logs": []
                                            })
                                            st.session_state[f"_add_task_form_{unique_key}"] = False
                                            mark_unsaved_changes()
                                            st.rerun()
                                with col_cancel_new:
                                    if st.button("✕ Ακύρωση", key=f"_cancel_new_task_{unique_key}", use_container_width=True):
                                        st.session_state[f"_add_task_form_{unique_key}"] = False
                                        st.rerun()

                        if len(item_tasks) > 0 and col_rem.button("Αφαίρεση", key=f"mat_rem_{unique_key}", use_container_width=True):
                            st.session_state["tasks_store"][unique_key].pop()
                            mark_unsaved_changes()
                            st.rerun()
                        if col_close.button("Κλείσιμο", key=f"mat_close_{unique_key}", use_container_width=True):
                            st.session_state.selected_material_expand = None
                            st.rerun()
        
        st.divider()
        col_save, _ = st.columns([1, 3])
        if col_save.button("💾 Αποθήκευση Αλλαγών", use_container_width=True, type="primary", key=f"save_project_{selected_project}"):
            if save_all_assignments_to_sheet():
                st.success("✅ Όλες οι αναθέσεις αποθηκεύτηκαν.")
                st.session_state["_has_unsaved_changes"] = False
                st.rerun()
            else:
                st.error("❌ Σφάλμα κατά την αποθήκευση")
        
        return
    
    col_check, col_info = st.columns([1, 3])
    with col_check:
        show_all = st.checkbox(
            "Εμφάνιση Όλων των Projects (μαζί με OK SHIPPED)", 
            value=False,
            help="Αν τσεκαριστεί, εμφανίζονται ΟΛΑ τα projects."
        )
    
    all_projects_raw = [p for p in procurement_df["Project"].unique().tolist() if p != "-"]
    
    def get_sort_date(p_name):
        is_shipped, ship_date_str = get_shipping_info(p_name, incoming_df)
        if ship_date_str:
            parsed = parse_date_safe(ship_date_str)
            if parsed:
                return parsed
        return date(9999, 12, 31)

    all_projects = sorted(all_projects_raw, key=get_sort_date)
    
    # === CACHE για get_project_details ===
    if "_proj_details_cache" not in st.session_state:
        st.session_state["_proj_details_cache"] = {}
    if st.session_state.get("_proj_cache_ts") != date.today():
        st.session_state["_proj_details_cache"] = {}
        st.session_state["_proj_cache_ts"] = date.today()
    
    def get_proj_details_cached(p_name):
        cache = st.session_state["_proj_details_cache"]
        if p_name in cache:
            return cache[p_name]
        details = get_project_details(p_name, procurement_df, tasks_database, incoming_df)
        cache[p_name] = details
        return details
    # ====================================
    
    if show_all:
        projects_to_show = all_projects
        section_title = "Όλα τα Projects"
    else:
        projects_to_show = []
        for p_name in all_projects:
            proj_data = get_proj_details_cached(p_name)
            if proj_data['is_active']:
                projects_to_show.append(p_name)
        section_title = "Ενεργά Projects"

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
            
            item_tasks = st.session_state.get("tasks_store", {}).get(u_key, [])
            for t in item_tasks:
                if t.get("task") and t["task"] != "- Επιλογή Εργασίας -":
                    auto_t = tasks_database.get(t["task"], 0.0)
                    eff_q = t.get("task_qty", qty)
                    try:
                        eff_q = float(eff_q) if eff_q not in [None, ""] else float(qty)
                    except:
                        eff_q = float(qty)
                    hrs = (auto_t * eff_q) / 60
                    p_total_hrs += hrs
                    p_tasks_cnt += 1
                    if t.get("done", False):
                        p_done_hrs += hrs
                        p_done_cnt += 1

        p_key = f"proj_{p_name}"
        p_tasks_dict = st.session_state.get("project_tasks_store", {}).get(p_key, {})
        if isinstance(p_tasks_dict, dict):
            for t_name, p_data in p_tasks_dict.items():
                if isinstance(p_data, dict) and p_data.get("active", False):
                    auto_t = tasks_database.get(t_name, 0.0)
                    eff_q = p_data.get("task_qty", p_main_qty)
                    try:
                        eff_q = float(eff_q) if eff_q not in [None, ""] else float(p_main_qty)
                    except:
                        eff_q = float(p_main_qty)
                    hrs = (auto_t * eff_q) / 60
                    p_total_hrs += hrs
                    p_tasks_cnt += 1
                    if p_data.get("done", False):
                        p_done_hrs += hrs
                        p_done_cnt += 1

        tot_all_hours += p_total_hrs
        tot_tasks_count += p_tasks_cnt
        tot_done_tasks += p_done_cnt
        p_progress = int((p_done_cnt / p_tasks_cnt) * 100) if p_tasks_cnt > 0 else 0
        
        proj_details = get_proj_details_cached(p_name)
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
            "Shipping Date": proj_details.get("shipping_date", ""),
            "Κατάσταση": "OK Shipped" if is_shipped else "ΕΝΕΡΓΟ"
        })

    # === ΔΙΟΡΘΩΣΗ: Απλό DataFrame αντί locals() ===
    dash_df = pd.DataFrame(dashboard_data) if dashboard_data else pd.DataFrame()
    # ==================================================

    st.markdown(f"### {section_title}")
    st.caption(f"Εμφανίζονται **{len(projects_to_show)}** projects")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Projects" if show_all else "Ενεργά", len(projects_to_show))
    c2.metric("Συνολικές Ώρες", f"{round(tot_all_hours, 1)}h")
    
    overall_pct = int((tot_done_tasks / tot_tasks_count) * 100) if tot_tasks_count > 0 else 0
    c3.metric("Συνολική Πρόοδος", f"{overall_pct}%")
    c4.metric("Εκκρεμή Tasks", tot_tasks_count - tot_done_tasks)

    st.divider()
    
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
            .compact-card { background: #ffffff; border: 1px solid #e0e0e0; border-radius: 8px; padding: 10px; margin-bottom: 8px; }
            .compact-card.active { border-left: 3px solid #2e7d32; }
            .compact-card.shipped { border-left: 3px solid #c62828; opacity: 0.7; }
            .compact-title { font-size: 12px; font-weight: 700; color: #1a1a1a; margin-bottom: 4px; min-height: 30px; }
            .compact-status { font-size: 9px; font-weight: 600; padding: 2px 6px; border-radius: 8px; display: inline-block; margin-bottom: 6px; }
            .compact-progress-bg { background: #e0e0e0; border-radius: 4px; height: 5px; overflow: hidden; margin: 6px 0; }
            .compact-progress-fill { height: 100%; border-radius: 4px; }
            .compact-meta { font-size: 10px; color: #666; margin-top: 4px; display: flex; justify-content: space-between; }
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
                    
                    shipping_date_html = ""
                    if proj_data.get("shipping_date"):
                        shipping_date_html = f'<div style="font-size:9px;color:#c62828;margin-top:3px;">📦 {proj_data["shipping_date"]}</div>'
                    
                    tooltip_lines = [f"📦 {p_name}", "━━━━━━━━━━━━━━━━"]
                    tooltip_lines.append(f"🚚 Αποστολή: {proj_data.get('shipping_date', '-') or '-'}")
                    
                    completed_h = proj_data.get('completed_hours', 0.0)
                    total_h_check = proj_data.get('total_hours', total_h)
                    tooltip_lines.append(f"📊 Πρόοδος: {progress}% ({done_tasks}/{total_tasks} tasks)")
                    
                    if 0 < completed_h < 0.1:
                        tooltip_lines.append(f"⏱️ Ώρες: {completed_h:.2f}h / {total_h_check}h")
                    else:
                        tooltip_lines.append(f"⏱️ Ώρες: {completed_h}h / {total_h_check}h")
                    
                    tooltip_lines.append(f"📦 Υλικά: {materials}")
                    tooltip_lines.append(f"✅ Tasks: {done_tasks}/{total_tasks}")
                    
                    pending_tasks = []
                    filtered_p_items = procurement_df[procurement_df["Project"] == p_name]
                    for idx_t, r_t in filtered_p_items.iterrows():
                        item_id_t = str(r_t["ID"])
                        u_key_t = f"{item_id_t}_{idx_t}"
                        qty_t = int(r_t["Ποσότητα"]) if str(r_t["Ποσότητα"]).isdigit() else 1
                        for tt in st.session_state["tasks_store"].get(u_key_t, []):
                            if tt.get("task") and tt["task"] != "- Επιλογή Εργασίας -" and not tt.get("done", False):
                                auto_t = tasks_database.get(tt["task"], 0.0)
                                eff_q = tt.get("task_qty", qty_t)
                                try:
                                    eff_q = float(eff_q) if eff_q not in [None, ""] else float(qty_t)
                                except:
                                    eff_q = float(qty_t)
                                hrs_t = (auto_t * eff_q) / 60
                                pending_tasks.append({"name": tt["task"], "hours": hrs_t})
                    
                    p_key_t = f"proj_{p_name}"
                    p_tasks_dict_t = st.session_state["project_tasks_store"].get(p_key_t, {})
                    if isinstance(p_tasks_dict_t, dict):
                        p_main_qty_t = 1
                        for _, r_t in filtered_p_items.iterrows():
                            if str(r_t["Ποσότητα"]).isdigit():
                                p_main_qty_t = max(p_main_qty_t, int(r_t["Ποσότητα"]))
                        for task_name_t, p_data_t in p_tasks_dict_t.items():
                            if isinstance(p_data_t, dict) and p_data_t.get("active", False) and not p_data_t.get("done", False):
                                auto_t = tasks_database.get(task_name_t, 0.0)
                                eff_q = p_data_t.get("task_qty", p_main_qty_t)
                                try:
                                    eff_q = float(eff_q) if eff_q not in [None, ""] else float(p_main_qty_t)
                                except:
                                    eff_q = float(p_main_qty_t)
                                hrs_t = (auto_t * eff_q) / 60
                                pending_tasks.append({"name": task_name_t, "hours": hrs_t})
                    
                    task_agg = {}
                    for pt in pending_tasks:
                        n = pt["name"]
                        if n not in task_agg:
                            task_agg[n] = {"count": 0, "hours": 0.0}
                        task_agg[n]["count"] += 1
                        task_agg[n]["hours"] += pt["hours"]
                    
                    sorted_pending = sorted(task_agg.items(), key=lambda x: -x[1]["hours"])[:3]
                    
                    if sorted_pending:
                        tooltip_lines.append("")
                        tooltip_lines.append(f"🔧 Εκκρεμή Tasks (top {len(sorted_pending)}):")
                        for task_n, agg in sorted_pending:
                            tooltip_lines.append(f"  • {task_n} — {round(agg['hours'], 1)}h ({agg['count']}x)")
                    
                    all_users_in_proj = set()
                    for idx_u, r_u in filtered_p_items.iterrows():
                        item_id_u = str(r_u["ID"])
                        u_key_u = f"{item_id_u}_{idx_u}"
                        for tt in st.session_state["tasks_store"].get(u_key_u, []):
                            if tt.get("task") and tt["task"] != "- Επιλογή Εργασίας -":
                                for uu in tt.get("users", []):
                                    if uu and uu != "- Χωρίς Ανάθεση -":
                                        all_users_in_proj.add(uu)
                    
                    if isinstance(p_tasks_dict_t, dict):
                        for task_name_t, p_data_t in p_tasks_dict_t.items():
                            if isinstance(p_data_t, dict) and p_data_t.get("active", False):
                                for uu in p_data_t.get("users", []):
                                    if uu and uu != "- Χωρίς Ανάθεση -":
                                        all_users_in_proj.add(uu)
                    
                    if all_users_in_proj:
                        tooltip_lines.append("")
                        tooltip_lines.append(f"👥 Υπεύθυνοι ({len(all_users_in_proj)}):")
                        for uu in sorted(all_users_in_proj):
                            tooltip_lines.append(f"  • {uu}")
                    
                    tooltip_text = "&#10;".join(tooltip_lines).replace('"', '&quot;')
                    
                    card_html = (
                        f'<div class="{card_class}" title="{tooltip_text}" style="cursor:help;">'
                        f'<div class="compact-title">{display_name}</div>'
                        f'{status_html}'
                        f'{shipping_date_html}'
                        f'<div class="compact-progress-bg">'
                        f'<div class="compact-progress-fill" style="width:{progress}%;background:{bar_color};"></div>'
                        f'</div>'
                        f'<div style="text-align:center;font-size:11px;font-weight:600;color:#333;">{progress}%</div>'
                        f'<div class="compact-meta">'
                        f'<span>{total_h}h</span>'
                        f'<span>{materials} υλικά</span>'
                        f'<span>{done_tasks}/{total_tasks}</span>'
                        f'</div>'
                        f'</div>'
                    )
                    st.markdown(card_html, unsafe_allow_html=True)
                    
                    if st.button("Άνοιγμα", key=f"drill_{p_name}", use_container_width=True):
                        st.session_state.selected_project_drill = p_name
                        st.session_state.selected_material_expand = None
                        st.rerun()
        
        st.divider()
        col_exp1, col_exp2 = st.columns(2)
        
        with col_exp1:
            csv = dash_df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="Εξαγωγή CSV", data=csv, 
                file_name=f"Projects_{'All' if show_all else 'Active'}_{date.today().strftime('%Y-%m-%d')}.csv", 
                mime="text/csv", use_container_width=True
            )
        with col_exp2:
            excel_data = export_to_excel(dash_df, "Projects")
            st.download_button(
                label="Εξαγωγή Excel", data=excel_data, 
                file_name=f"Projects_{'All' if show_all else 'Active'}_{date.today().strftime('%Y-%m-%d')}.xlsx", 
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                use_container_width=True
            )


# --- RENDER MASTER VIEW (ΔΕΝ ΚΑΛΕΙΤΑΙ ΠΙΑ - ΚΡΑΤΙΕΤΑΙ ΓΙΑ ROLLBACK) ---
def render_master_view(procurement_df, tasks_database, team_database, availability_database, incoming_df):
    st.header("Master View")
    st.caption("⚠️ Αυτή η σελίδα έχει αντικατασταθεί από το '📅 Πλάνο Εβδομάδας'")
    st.info("Ο παλιός κώδικας του Master View έχει αφαιρεθεί για καθαριότητα. "
            "Αν χρειαστείς rollback, χρησιμοποίησε το backup σου.")


# --- RENDER WEEK PLANNER (ΝΕΟ - αντικαθιστά το Master View) ---
def render_week_planner(procurement_df, tasks_database, team_database,
                        availability_database, incoming_df):
    st.header("📅 Πλάνο Εβδομάδας")
    st.caption("Σχεδίασε την εβδομάδα της ομάδας — Κεντρικό & Παράρτημα σε μία οθόνη")

    # ---------- Φίλτρα ----------
    col_loc, col_week, col_btn = st.columns([1, 2, 1])
    with col_loc:
        location_filter = st.selectbox(
            "Location",
            [LOCATION_ALL, LOCATION_MAIN, LOCATION_WAREHOUSE],
            index=0,
            key="wp_location",
        )
    with col_week:
        today = date.today()
        default_week = today - timedelta(days=today.weekday())
        week_start = st.date_input(
            "Εβδομάδα (επέλεξε Δευτέρα)",
            value=default_week,
            format="DD/MM/YYYY",
            key="wp_week_start",
        )
        if isinstance(week_start, date):
            week_start = week_start - timedelta(days=week_start.weekday())
    with col_btn:
        st.write("")
        st.write("")
        if st.button("🔄 Ανανέωση", use_container_width=True, key="wp_refresh"):
            load_assignments_from_sheet.clear()
            st.rerun()

    week_end = week_start + timedelta(days=4)
    st.markdown(f"### Εβδομάδα: **{week_start.strftime('%d/%m/%Y')} – {week_end.strftime('%d/%m/%Y')}**")

    # ---------- KPIs ----------
    plan_df = build_week_plan(week_start, location_filter)
    unassigned = get_unassigned_ready_tasks(location_filter)

    total_avail = plan_df["available_h"].sum() if not plan_df.empty else 0
    total_assigned = plan_df["assigned_h"].sum() if not plan_df.empty else 0
    coverage = int((total_assigned / total_avail) * 100) if total_avail > 0 else 0
    overload_count = int(plan_df["overload"].sum()) if not plan_df.empty else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Διαθέσιμες Ώρες", f"{round(total_avail, 1)}h")
    k2.metric("Assigned Ώρες", f"{round(total_assigned, 1)}h")
    k3.metric("Κάλυψη", f"{coverage}%")
    k4.metric("⚠️ Overloads", overload_count)

    st.divider()

    # ---------- Πίνακας εβδομάδας ----------
    st.subheader("Πλάνο ανά άτομο")
    if plan_df.empty:
        st.info("Δεν υπάρχουν μέλη για το επιλεγμένο location.")
    else:
        pivot = plan_df.pivot_table(index="name", columns="date", values="assigned_h", aggfunc="sum")
        avail_pivot = plan_df.pivot_table(index="name", columns="date", values="available_h", aggfunc="sum")

        display_df = pd.DataFrame(index=pivot.index)
        for d in sorted(pivot.columns):
            col_label = f"{WEEKDAYS_SHORT_GREEK.get(d.weekday(), '')} {d.strftime('%d/%m')}"
            display_df[col_label] = [
                f"{pivot.loc[name, d]:.1f}h / {avail_pivot.loc[name, d]:.1f}h"
                for name in pivot.index
            ]
        display_df["Σύνολο"] = [
            f"{plan_df[plan_df['name'] == n]['assigned_h'].sum():.1f}h / "
            f"{plan_df[plan_df['name'] == n]['available_h'].sum():.1f}h"
            for n in display_df.index
        ]

        def color_cell(val):
            try:
                a, v = val.split(" / ")
                a = float(a.replace("h", ""))
                v = float(v.replace("h", ""))
                if v == 0:
                    return "background-color:#eeeeee;color:#999"
                if a > v:
                    return "background-color:#ffcdd2;color:#b71c1c;font-weight:600"
                if a > 0:
                    return "background-color:#c8e6c9;color:#1b5e20"
                return "background-color:#fffde7;color:#827717"
            except Exception:
                return ""

        try:
            styled = display_df.style.map(color_cell)
            st.dataframe(styled, use_container_width=True)
        except AttributeError:
            # Παλιά έκδοση Streamlit - χωρίς styling
            st.dataframe(display_df, use_container_width=True)

    st.divider()

    # ---------- Εκκρεμή tasks ----------
    st.subheader(f"🔴 Εκκρεμή Tasks ({len(unassigned)})")
    st.caption("Tasks χωρίς ανάθεση ή με παρελθούσα ημερομηνία. "
               "✅ = υλικό έτοιμο (Procurement), ⏳ = εκκρεμεί παραλαβή.")

    if not unassigned:
        st.success("✅ Δεν υπάρχουν εκκρεμή tasks για αυτό το location.")
        return

    preview = pd.DataFrame([{
        "Ready": "✅" if x["ready"] else "⏳",
        "Project": x["project"][:30],
        "Υλικό": x["material"][:35],
        "Εργασία": x["task"][:28],
        "Ώρες": x["hours"],
        "Due": x["due_date"].strftime("%d/%m") if x["due_date"] else "-",
        "Status Proc": x["status_proc"][:18],
    } for x in unassigned])
    st.dataframe(preview, use_container_width=True, hide_index=True)

    st.divider()

    # ---------- Bulk assign ----------
    st.subheader("📌 Ανάθεση σε άτομο")
    st.caption("Επίλεξε tasks, άτομο και ημερομηνία → Ανάθεση.")

    task_options = []
    task_lookup = {}
    for x in unassigned:
        label = (f"[{'✅' if x['ready'] else '⏳'}] {x['project'][:20]} — "
                 f"{x['task'][:20]} — {x['hours']}h — "
                 f"due {x['due_date'].strftime('%d/%m') if x['due_date'] else '-'}")
        task_options.append(label)
        task_lookup[label] = (x["u_key"], x["t_idx"])

    selected_labels = st.multiselect(
        "Επιλογή tasks",
        options=task_options,
        default=[],
        key="wp_selected_tasks",
    )

    col_user, col_date, col_btn2 = st.columns([2, 1, 1])
    with col_user:
        members_for_loc = [
            m for m in st.session_state.get("team_members_list", [])
            if m.get("active", True)
            and (location_filter == LOCATION_ALL
                 or m.get("location", LOCATION_MAIN) == location_filter)
        ]
        member_names = [m["name"] for m in members_for_loc]
        target_user = st.selectbox(
            "Άτομο",
            options=member_names if member_names else ["-"],
            key="wp_target_user",
        )
    with col_date:
        target_date = st.date_input(
            "Ημερομηνία ανάθεσης",
            value=week_start,
            format="DD/MM/YYYY",
            key="wp_target_date",
        )
    with col_btn2:
        st.write("")
        st.write("")
        do_assign = st.button(
            "✅ Ανάθεση", type="primary", use_container_width=True, key="wp_assign_btn"
        )

    if do_assign:
        if not selected_labels:
            st.warning("⚠️ Διάλεξε τουλάχιστον ένα task.")
        elif target_user == "-":
            st.warning("⚠️ Διάλεξε άτομο.")
        else:
            refs = [task_lookup[lbl] for lbl in selected_labels]
            changed = assign_tasks_bulk(refs, target_user, target_date)
            st.success(f"✅ Ανατέθηκαν {changed} tasks στον/στην {target_user} "
                       f"για {target_date.strftime('%d/%m/%Y')}.")
            st.session_state["_has_unsaved_changes"] = True
            st.rerun()

    # ---------- Save & Export ----------
    st.divider()
    col_save, col_export = st.columns(2)
    with col_save:
        if st.button("💾 Αποθήκευση στο Google Sheet", type="primary",
                     use_container_width=True, key="wp_save"):
            if save_all_assignments_to_sheet():
                st.success("✅ Αποθηκεύτηκε.")
                st.session_state["_has_unsaved_changes"] = False
                st.rerun()
            else:
                st.error("❌ Σφάλμα αποθήκευσης.")

    with col_export:
        if not plan_df.empty:
            export_df = plan_df[["name", "location", "date", "weekday",
                                 "available_h", "assigned_h", "overload"]].copy()
            export_df["date"] = export_df["date"].astype(str)
            csv = export_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Export CSV Πλάνου",
                data=csv,
                file_name=f"plan_{week_start.strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True,
                key="wp_export",
            )


# ============================================================
# === MAIN BLOCK =============================================
# ============================================================

init_auth()

if not st.session_state.authenticated:
    login_form()
else:
    # ---------- SIDEBAR ----------
    with st.sidebar:
        st.markdown(f"### 👤 {st.session_state.username}")
        st.divider()
        
        st.markdown("**Location Filter**")
        location_choice = st.radio(
            "Επιλογή τοποθεσίας:",
            options=[LOCATION_ALL, LOCATION_MAIN, LOCATION_WAREHOUSE],
            index=0,
            key="sidebar_location",
            label_visibility="collapsed",
        )
        st.session_state["_location_filter"] = location_choice
        
        st.divider()
        
        st.markdown("**Navigation**")
        pages = [
            "📦 Projects",
            "📅 Πλάνο Εβδομάδας",
        ]
        
        current_page = st.session_state.get("page", "📦 Projects")
        if current_page not in pages:
            current_page = pages[0]
        
        selected_page = st.radio(
            "Σελίδα:",
            options=pages,
            index=pages.index(current_page),
            key="sidebar_nav",
            label_visibility="collapsed",
        )
        st.session_state.page = selected_page
        
        st.divider()
        
        if st.session_state.get("_has_unsaved_changes", False):
            st.warning("⚠️ Μη αποθηκευμένες αλλαγές")
            if st.button("💾 Αποθήκευση τώρα", use_container_width=True, type="primary"):
                if save_all_assignments_to_sheet():
                    st.success("✅ Αποθηκεύτηκε")
                    st.rerun()
        
        st.divider()
        
        if st.button("🚪 Logout", use_container_width=True):
            logout()
        
        st.caption("v2.0 — Πλάνο Εβδομάδας")
    
    # ---------- LOAD DATA ----------
    with st.spinner("Φόρτωση δεδομένων..."):
        procurement_df, tasks_dict, team_member_names, availability_dict, df_incoming, team_members_list, availability_overrides = load_all_data()
        
        st.session_state["procurement_df"] = procurement_df
        st.session_state["_tasks_db_cache"] = tasks_dict
        st.session_state["team_member_names"] = team_member_names
        st.session_state["availability_weekly"] = availability_dict
        st.session_state["incoming_df"] = df_incoming
        st.session_state["team_members_list"] = team_members_list
        st.session_state["availability_overrides"] = availability_overrides
        
        if "tasks_store" not in st.session_state or not st.session_state["tasks_store"]:
            assignments_item, assignments_proj = load_assignments_from_sheet()
            st.session_state["tasks_store"] = assignments_item
            st.session_state["project_tasks_store"] = assignments_proj
        
        if "project_tasks_store" not in st.session_state:
            st.session_state["project_tasks_store"] = {}
        
        ensure_task_uids()
    
    # ---------- DATA CHECK ----------
    if procurement_df.empty:
        st.error("⚠️ Δεν φορτώθηκαν δεδομένα από το Procurement sheet.")
        st.info("Έλεγξε τη σύνδεση με το Google Sheets ή τα permissions του service account.")
        st.stop()
    
    # ---------- ROUTING ----------
    if st.session_state.page == "📦 Projects":
        render_projects(
            st.session_state["procurement_df"],
            st.session_state["_tasks_db_cache"],
            st.session_state["team_member_names"],
            st.session_state["availability_weekly"],
            st.session_state["incoming_df"],
        )
    elif st.session_state.page == "📅 Πλάνο Εβδομάδας":
        render_week_planner(
            st.session_state["procurement_df"],
            st.session_state["_tasks_db_cache"],
            st.session_state["team_member_names"],
            st.session_state["availability_weekly"],
            st.session_state["incoming_df"],
        )
    else:
        st.error(f"Άγνωστη σελίδα: {st.session_state.page}")



