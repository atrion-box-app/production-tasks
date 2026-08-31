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
        # Διαβάζουμε το CSV χωρίς επικεφαλίδες πρώτα για να έχουμε ακριβείς θέσεις στηλών A=0, B=1, C=2 κτλ.
        df_raw = pd.read_csv(CSV_URL, header=None)
        
        # Οι στήλες που ζητήθηκαν βάσει των γραμμάτων του Excel:
        # B=1 (Project), C=2 (Ποσότητα), D=3 (Ημ. Παράδοσης), E=4 (Αναμ. Ημ. Παραλαβής), 
        # F=5 (Προμηθευτής), G=6 (Υλικό), K=10 (Αναμ. Ποσότητα), M=12 (Status Procurement), 
        # R=17 (Είδος Δώρου), S=18 (ID)
        indices = [18, 3, 17, 1, 2, 5, 6, 4, 10, 12]
        
        # Παίρνουμε μόνο αυτές τις στήλες
        df_selected = df_raw.iloc[1:, indices].copy() # Παραλείπουμε την πρώτη γραμμή τίτλων
        
        # Ορίζουμε τους δικούς μας καθαρούς τίτλους στα ελληνικά
        df_selected.columns = [
            "ID", 
            "Ημερομηνία Παράδοσης", 
            "Είδος Δώρου", 
            "Project", 
            "Ποσότητα", 
            "Προμηθευτής", 
            "Υλικό / Προϊόν", 
            "Αναμενόμενη Ημ. Παραλαβής", 
            "Αναμενόμενη Ποσότητα Παραλαβής", 
            "Status Procurement"
        ]
        
        # Καθαρισμός κενών τιμών
        df_selected = df_selected.fillna("-")
        return df_selected
        
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

    if option == "Καρτέλα Project":
        st.header("📋 Προβολή & Διαχείριση ανά Project")
        
        projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
        selected_project = st.selectbox("Επιλέξτε Project:", projects_list)
        
        st.subheader(f"Υλικά Procurement για το Project: {selected_project}")
        
        # Φιλτράρισμα υλικών για το επιλεγμένο Project
        filtered_df = procurement_df[procurement_df["Project"] == selected_project].copy()
        
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

    elif option == "Όλα τα Υλικά Παραγωγής":
        st.header("📦 Ζωντανή Λίστα Υλικών Procurement (Όλες οι 10 Στήλες)")
        
        column_config = {
            col: st.column_config.Column(alignment="center") for col in procurement_df.columns
        }
        
        st.dataframe(
            procurement_df, 
            use_container_width=True,
            column_config=column_config,
            hide_index=True
        )

    elif option == "Ημερολόγιο Διαθεσιμότητας":
        st.header("🗓️ Ημερήσιο Πλάνο Παραγωγής")
        st.info("Εδώ θα συνδέσουμε το Ημερολόγιο Συνθέσεων!")

else:
    st.info("Βεβαιωθείτε ότι το αρχείο του Procurement έχει ρυθμιστεί σε 'Anyone with the link can view'.")
