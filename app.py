import streamlit as st
import pandas as pd

st.set_page_config(page_title="Production Tasks App", layout="wide")

st.title("🏭 Σύστημα Διαχείρισης Παραγωγής & Tasks")

# Google Sheet URL του Procurement (CSV export link)
SHEET_ID = "1QhTd58vuulaC_73sgbjuwG5MVxT6c1c_-MbhypGx0fA"
GID = "1639392743"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

# Συνάρτηση για ζωντανή ανάγνωση των δεδομένων
@st.cache_data(ttl=10)  # Ανανέωση δεδομένων κάθε 10 δευτερόλεπτα
def load_procurement_data():
    try:
        df = pd.read_csv(CSV_URL)
        return df
    except Exception as e:
        st.error(f"Σφάλμα κατά τη σύνδεση με το Google Sheet του Procurement: {e}")
        return pd.DataFrame()

# Φόρτωση Δεδομένων
procurement_df = load_procurement_data()

# Πλευρικό μενού (Sidebar)
option = st.sidebar.selectbox(
    "Επιλογή Οθόνης",
    ["Καρτέλα Project", "Όλα τα Υλικά Παραγωγής", "Ημερολόγιο Διαθεσιμότητας"]
)

if not procurement_df.empty:
    
    if option == "Καρτέλα Project":
        st.header("📋 Προβολή & Διαχείριση ανά Project")
        
        # Καθαρισμός και εύρεση μοναδικών Projects
        # Προσαρμόστε το όνομα της στήλης αν στο αρχείο λέγεται 'Project'
        project_column = [col for col in procurement_df.columns if 'Project' in col or 'project' in col]
        
        if project_column:
            proj_col_name = project_column[0]
            projects_list = procurement_df[proj_col_name].dropna().unique().tolist()
            
            selected_project = st.selectbox("Επιλέξτε Project:", sorted(projects_list))
            
            st.subheader(f"Υλικά Procurement για το Project: {selected_project}")
            
            # Φιλτράρισμα υλικών μόνο για το επιλεγμένο Project
            filtered_df = procurement_df[procurement_df[proj_col_name] == selected_project]
            
            st.dataframe(filtered_df, use_container_width=True)
        else:
            st.warning("Δεν βρέθηκε στήλη με όνομα 'Project' στο αρχείο.")
            st.dataframe(procurement_df.head(), use_container_width=True)

    elif option == "Όλα τα Υλικά Παραγωγής":
        st.header("📦 Ζωντανή Λίστα Υλικών από Procurement")
        st.dataframe(procurement_df, use_container_width=True)

    elif option == "Ημερολόγιο Διαθεσιμότητας":
        st.header("🗓️ Ημερήσιο Πλάνο & Διαθεσιμότητα Ομάδας")
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Άτομα Βάρδιας", "4 άτομα")
        col2.metric("Συνολικά Υλικά Procurement", len(procurement_df))
        col3.metric("Ενεργά Projects", len(procurement_df[project_column[0]].unique()) if project_column else 0)

else:
    st.info("Βεβαιωθείτε ότι το αρχείο του Procurement έχει ρυθμιστεί σε 'Anyone with the link can view'.")
