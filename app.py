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
        df = pd.read_csv(CSV_URL)
        # Καθαρισμός ονομάτων στηλών από κενά
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
    
    # Συνάρτηση αναζήτησης στηλών
    def get_column(df, possible_names):
        for name in possible_names:
            matches = [col for col in df.columns if name.lower() in col.lower()]
            if matches:
                return matches[0]
        return None

    # Εύρεση των 10 στηλών Procurement
    col_id = get_column(procurement_df, ['id', 'sku'])
    col_date = get_column(procurement_df, ["project's due date", 'ημερομηνία παράδοσης', 'due date'])
    col_gift = get_column(procurement_df, ['είδος δώρου', 'gift', 'type'])
    col_project = get_column(procurement_df, ['project'])
    col_qty = get_column(procurement_df, ['quantity order', 'ποσότητα', 'project\'s q'])
    col_supplier = get_column(procurement_df, ['suppliers', 'προμηθευτής'])
    col_material = get_column(procurement_df, ['description', 'υλικό', 'προϊόν'])
    col_expected_date = get_column(procurement_df, ['order\'s due date', 'αναμενόμενη ημερομηνία παράδοσης'])
    col_expected_qty = get_column(procurement_df, ['quantity stock', 'αναμενόμενη ποσότητα', 'received'])
    col_status = get_column(procurement_df, ['status', 'status procurement'])

    # Συγκέντρωση των στηλών
    selected_cols = [c for c in [col_id, col_date, col_gift, col_project, col_qty, col_supplier, col_material, col_expected_date, col_expected_qty, col_status] if c is not None]

    if option == "Καρτέλα Project":
        st.header("📋 Προβολή & Διαχείριση ανά Project")
        
        if col_project:
            projects_list = sorted(procurement_df[col_project].dropna().unique().tolist())
            selected_project = st.selectbox("Επιλέξτε Project:", projects_list)
            
            st.subheader(f"Υλικά Procurement για το Project: {selected_project}")
            
            # Φιλτράρισμα υλικών για το επιλεγμένο Project
            filtered_df = procurement_df[procurement_df[col_project] == selected_project][selected_cols].copy()
            
            # Εμφάνιση καθαρού πίνακα με τις 10 στήλες Procurement
            st.dataframe(filtered_df, use_container_width=True)
            
        else:
            st.warning("Δεν βρέθηκε η στήλη 'Project' στο αρχείο.")

    elif option == "Όλα τα Υλικά Παραγωγής":
        st.header("📦 Ζωντανή Λίστα Υλικών Procurement (10 Στήλες)")
        st.dataframe(procurement_df[selected_cols], use_container_width=True)

    elif option == "Ημερολόγιο Διαθεσιμότητας":
        st.header("🗓️ Ημερήσιο Πλάνο Παραγωγής")
        st.info("Εδώ θα συνδέσουμε το Ημερολόγιο Συνθέσεων!")

else:
    st.info("Βεβαιωθείτε ότι το αρχείο του Procurement έχει ρυθμιστεί σε 'Anyone with the link can view'.")
