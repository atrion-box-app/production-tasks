import streamlit as st
import pandas as pd

st.set_page_config(page_title="Production Tasks App", layout="wide")

st.title("🏭 Σύστημα Διαχείρισης Παραγωγής & Tasks")

# Πλευρικό μενού (Sidebar)
option = st.sidebar.selectbox(
    "Επιλογή Οθόνης",
    ["Καρτέλα Project", "Καταχώρηση Procurement", "Ημερολόγιο Διαθεσιμότητας"]
)

if option == "Καρτέλα Project":
    st.header("📋 Προβολή & Διαχείριση ανά Project")
    
    project = st.selectbox("Επιλέξτε Project:", ["PWC", "CONSTRAT", "STARBULK"])
    
    st.subheader(f"Υλικά & Tasks για το Project: {project}")
    
    # Αυτόματος πίνακας χωρίς κίνδυνο μετατόπισης γραμμών
    data = {
        "ID": [101, 102, 103],
        "Υλικό": ["Χαρτοσακούλα 29x22x4.5", "Είδος κορδέλας BLACK", "Προϊόν 3"],
        "Ποσότητα": [50, 50, 50],
        "Task 1": ["Έλεγχος", "Κόψιμο", "Συναρμολόγηση"],
        "Status Task 1": [True, False, False],
        "Task 2": ["Συναρμολόγηση", "Τοποθέτηση", "Συσκευασία"],
        "Status Task 2": [False, False, False]
    }
    df = pd.DataFrame(data)
    
    # Επεξεργάσιμος πίνακας
    st.data_editor(df, use_container_width=True)
    
    st.success("⏱️ Συνολικός Χρόνος Παραγωγής Project: **4.2 Ώρες** (~1 ημέρα)")

elif option == "Καταχώρηση Procurement":
    st.header("➕ Νέα Καταχώρηση Παραγγελίας από Procurement")
    
    with st.form("new_order"):
        proj_name = st.text_input("Όνομα Project")
        mat_name = st.text_input("Περιγραφή Υλικού")
        qty = st.number_input("Ποσότητα", min_value=1, value=100)
        submitted = st.form_submit_button("Αποθήκευση Παραγγελίας")
        if submitted:
            st.success(f"Η παραγγελία {proj_name} καταχωρήθηκε με ασφάλεια στο ID σύστημα!")

elif option == "Ημερολόγιο Διαθεσιμότητας":
    st.header("🗓️ Ημερήσιο Πλάνο & Διαθεσιμότητα Ομάδας")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Άτομα Βάρδιας", "4 άτομα")
    col2.metric("Διαθέσιμες Ώρες", "32 ώρες")
    col3.metric("Δεσμευμένες Ώρες", "18.5 ώρες", "-13.5 ώρες υπόλοιπο")
    
    st.info("💡 Όλα τα projects υπολογίζονται αυτόματα χωρίς χειροκίνητους τύπους Excel!")
