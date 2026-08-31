import streamlit as st
import pandas as pd

st.set_page_config(page_title="Production Tasks App", layout="wide")

st.title("🏭 Σύστημα Διαχείρισης Παραγωγής & Tasks")

# Google Sheet URL του Procurement (CSV export link)
SHEET_ID = "1QhTd58vuulaC_73sgbjuwG5MVxT6c1c_-MbhypGx0fA"
GID = "1639392743"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

@st.cache_data(ttl=10)
def load_procurement_data():
    try:
        # Διαβάζουμε το CSV
        df = pd.read_csv(CSV_URL)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Σφάλμα κατά τη σύνδεση με το Google Sheet του Procurement: {e}")
        return pd.DataFrame()

procurement_df = load_procurement_data()

# Πλευρικό μενού (Sidebar)
option = st.sidebar.selectbox(
    "Επιλογή Οθόνης",
    ["Καρτέλα Project", "Όλα τα Υλικά Παραγωγής", "Ημερολόγιο Διαθεσιμότητας"]
)

if not procurement_df.empty:
    
    # Χαρτογράφηση των 10 στηλών με βάση τα πραγματικά ονόματα του Procurement Sheet
    # ID (S), Ημ. Παράδοσης (D), Είδος Δώρου (R), Project (B), Ποσότητα (C), Προμηθευτής (F), Υλικό (G), Αναμ. Ημ. Παραλαβής (E), Αναμ. Ποσότητα (K), Status (M)
    
    column_mapping = {}
    for col in procurement_df.columns:
        c_lower = col.lower()
        if c_lower in ['id', 'sku']:
            column_mapping['ID'] = col
        elif 'project\'s due date' in c_lower or 'ημερομηνία παράδοσης' in c_lower:
            column_mapping['Ημερομηνία Παράδοσης'] = col
        elif 'type of gift' in c_lower or 'είδος δώρου' in c_lower:
            column_mapping['Είδος Δώρου'] = col
        elif c_lower == 'project':
            column_mapping['Project'] = col
        elif 'project\'s q' in c_lower or 'quantity order' in c_lower or 'ποσότητα' in c_lower:
            column_mapping['Ποσότητα'] = col
        elif 'suppliers' in c_lower or 'προμηθευτής' in c_lower:
            column_mapping['Προμηθευτής'] = col
        elif 'description' in c_lower or 'υλικό' in c_lower:
            column_mapping['Υλικό / Προϊόν'] = col
        elif 'order\'s due date' in c_lower or 'received date' in c_lower or 'αναμενόμενη ημερομηνία' in c_lower:
            column_mapping['Αναμενόμενη Ημ. Παραλαβής'] = col
        elif 'quantity stock' in c_lower or 'αναμενόμενη ποσότητα' in c_lower:
            column_mapping['Αναμενόμενη Ποσότητα'] = col
        elif 'status' in c_lower:
            column_mapping['Status Procurement'] = col

    # Ανάκτηση των διαθέσιμων στηλών
    selected_cols_raw = list(column_mapping.values())
    
    if option == "Καρτέλα Project":
        st.header("📋 Προβολή & Διαχείριση ανά Project")
        
        project_col_raw = column_mapping.get('Project', 'Project')
        
        if project_col_raw in procurement_df.columns:
            projects_list = sorted(procurement_df[project_col_raw].dropna().unique().tolist())
            selected_project = st.selectbox("Επιλέξτε Project:", projects_list)
            
            st.subheader(f"Υλικά Procurement για το Project: {selected_project}")
            
            # Φιλτράρισμα και μετονομασία στηλών στα ελληνικά
            filtered_df = procurement_df[procurement_df[project_col_raw] == selected_project][selected_cols_raw].copy()
            
            # Αντίστροφη μετονομασία για όμορφη εμφανιση
            rename_dict = {v: k for k, v in column_mapping.items()}
            filtered_df = filtered_df.rename(columns=rename_dict)
            
            # Κεντράρισμα όλων των στηλών
            column_config = {
                col: st.column_config.Column(alignment="center") for col in filtered_df.columns
            }
            
            st.dataframe(
                filtered_df, 
                use_container_width=True,
                column_config=column_config,
                hide_index=True
            )
            
        else:
            st.warning("Δεν βρέθηκε η στήλη 'Project' στο αρχείο.")

    elif option == "Όλα τα Υλικά Παραγωγής":
        st.header("📦 Ζωντανή Λίστα Υλικών Procurement (10 Στήλες)")
        
        display_df = procurement_df[selected_cols_raw].copy()
        rename_dict = {v: k for k, v in column_mapping.items()}
        display_df = display_df.rename(columns=rename_dict)
        
        column_config = {
            col: st.column_config.Column(alignment="center") for col in display_df.columns
        }
        
        st.dataframe(
            display_df, 
            use_container_width=True,
            column_config=column_config,
            hide_index=True
        )

    elif option == "Ημερολόγιο Διαθεσιμότητας":
        st.header("🗓️ Ημερήσιο Πλάνο Παραγωγής")
        st.info("Εδώ θα συνδέσουμε το Ημερολόγιο Συνθέσεων!")

else:
    st.info("Βεβαιωθείτε ότι το αρχείο του Procurement έχει ρυθμιστεί σε 'Anyone with the link can view'.")
