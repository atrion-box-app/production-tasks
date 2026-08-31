import streamlit as st
import pandas as pd
from datetime import date

st.set_page_config(page_title="Production Tasks App", layout="wide")

st.title("🏭 Σύστημα Διαχείρισης Παραγωγής & Tasks")

# Google Sheet URL του Procurement (CSV export link)
SHEET_ID = "1QhTd58vuulaC_73sgbjuwG5MVxT6c1c_-MbhypGx0fA"
GID = "1639392743"
CSV_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={GID}"

@st.cache_data(ttl=10)
def load_procurement_data():
    try:
        df_raw = pd.read_csv(CSV_URL, header=None)
        indices = [18, 3, 17, 1, 2, 5, 6, 4, 10, 12]
        df_selected = df_raw.iloc[1:, indices].copy()
        
        df_selected.columns = [
            "ID", "Ημερομηνία Παράδοσης", "Είδος Δώρου", "Project", 
            "Ποσότητα", "Προμηθευτής", "Υλικό / Προϊόν", 
            "Αναμενόμενη Ημ. Παραλαβής", "Αναμενόμενη Ποσότητα Παραλαβής", "Status Procurement"
        ]
        return df_selected.fillna("-")
    except Exception as e:
        st.error(f"Σφάλμα κατά τη σύνδεση: {e}")
        return pd.DataFrame()

procurement_df = load_procurement_data()

# Λίστα Προσωπικού Παραγωγής (Μπορείτε να προσθέσετε/αλλάξετε ονόματα)
TEAM_MEMBERS = ["- Χωρίς Ανάθεση -", "Γιώργος", "Βασίλης", "Μαρία", "Κώστας", "Ελένη", "Νίκος"]

# Λίστα Διαθέσιμων Σταδίων Παραγωγής
AVAILABLE_TASKS = ["Έλεγχος & Διαλογή", "Συναρμολόγηση / Δέσιμο", "Χάραξη / Εκτύπωση", "Συσκευασία & Box", "Τελικός Έλεγχος Quality"]

option = st.sidebar.selectbox("Επιλογή Οθόνης", ["Καρτέλα Project", "Όλα τα Υλικά Παραγωγής"])

if not procurement_df.empty:
    if option == "Καρτέλα Project":
        st.header("📋 Διαχείριση Παραγωγής & Αναθέσεις ανά Project")
        
        projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
        selected_project = st.selectbox("Επιλέξτε Project:", projects_list)
        
        filtered_df = procurement_df[procurement_df["Project"] == selected_project].copy()
        
        st.subheader(f"📦 Υλικά & Tasks για το {selected_project} ({len(filtered_df)} Υλικά)")
        
        total_project_hours = 0.0
        
        # Εμφάνιση κάθε Υλικού σε διαδραστική Κάρτα (Accordion)
        for idx, row in filtered_df.iterrows():
            item_id = row["ID"]
            material = row["Υλικό / Προϊόν"]
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            status = row["Status Procurement"]
            
            card_title = f"🆔 {item_id} | {material} — (Ποσότητα: {qty} τμχ) | Status: {status}"
            
            with st.expander(card_title):
                col_info, col_tasks = st.columns([1, 2.5])
                
                with col_info:
                    st.markdown("**ℹ️ Στοιχεία Procurement**")
                    st.write(f"• **Είδος Δώρου:** {row['Είδος Δώρου']}")
                    st.write(f"• **Προμηθευτής:** {row['Προμηθευτής']}")
                    st.write(f"• **Ημ. Παράδοσης:** {row['Ημερομηνία Παράδοσης']}")
                    st.write(f"• **Αναμ. Παραλαβή:** {row['Αναμενόμενη Ημ. Παραλαβής']}")
                    st.divider()

                with col_tasks:
                    st.markdown("**⚙️ Αναθέσεις & Στάδια Παραγωγής**")
                    
                    # Ορίζουμε 2 δυναμικά στάδια παραγωγής ανά υλικό
                    for t_num in [1, 2]:
                        st.caption(f"**Στάδιο {t_num}**")
                        c_check, c_task, c_user, c_date, c_time = st.columns([0.1, 0.3, 0.25, 0.23, 0.17])
                        
                        # Checkbox Ολοκλήρωσης
                        done = c_check.checkbox("", key=f"chk_{item_id}_{t_num}")
                        
                        # Επιλογή Σταδίου
                        task_name = c_task.selectbox(
                            "Στάδιο", AVAILABLE_TASKS, 
                            index=0 if t_num==1 else 1, 
                            key=f"task_{item_id}_{t_num}", 
                            label_visibility="collapsed"
                        )
                        
                        # Επιλογή Ατόμου
                        assigned_user = c_user.selectbox(
                            "Ανάθεση", TEAM_MEMBERS, 
                            index=1 if t_num==1 else 2, 
                            key=f"user_{item_id}_{t_num}", 
                            label_visibility="collapsed"
                        )
                        
                        # Επιλογή Ημερομηνίας Ανάθεσης
                        assign_date = c_date.date_input(
                            "Ημερομηνία", value=date.today(), 
                            key=f"date_{item_id}_{t_num}", 
                            label_visibility="collapsed"
                        )
                        
                        # Χρόνος σε λεπτά ανά τεμάχιο
                        mins = c_time.number_input(
                            "λεπτά/τμχ", min_value=0.0, value=2.0 if t_num==1 else 5.0, step=0.5,
                            key=f"time_{item_id}_{t_num}", 
                            label_visibility="collapsed"
                        )
                        
                        # Υπολογισμός ωρών
                        task_hours = (mins * qty) / 60
                        total_project_hours += task_hours

        st.divider()
        st.success(f"🎯 **Συνολικός Χρόνος Παραγωγής για το Project {selected_project}: {round(total_project_hours, 1)} Ώρες**")

    elif option == "Όλα τα Υλικά Παραγωγής":
        st.header("📦 Λίστα Υλικών")
        st.dataframe(procurement_df, use_container_width=True, hide_index=True)
