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
        df_raw = pd.read_csv(CSV_URL, header=None)
        # B=1, C=2, D=3, E=4, F=5, G=6, K=10, M=12, R=17, S=18
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

# Πρότυπα Χρόνων (Templates) ανά Είδος Δώρου
TASK_TEMPLATES = {
    "POUCH": [
        {"Task": "1. Έλεγχος & Διαλογή", "Χρόνος_λεπτά": 2, "Υπεύθυνος": "Γιώργος"},
        {"Task": "2. Συναρμολόγηση / Δέσιμο", "Χρόνος_λεπτά": 5, "Υπεύθυνος": "Βασίλης"},
        {"Task": "3. Συσκευασία & Box", "Χρόνος_λεπτά": 1.5, "Υπεύθυνος": "Μαρία"}
    ],
    "GIFT/AWARD": [
        {"Task": "1. Έλεγχος Υλικών", "Χρόνος_λεπτά": 3, "Υπεύθυνος": "Γιώργος"},
        {"Task": "2. Χάραξη / Κοπή", "Χρόνος_λεπτά": 10, "Υπεύθυνος": "Κώστας"},
        {"Task": "3. Τελικός Έλεγχος Quality", "Χρόνος_λεπτά": 2, "Υπεύθυνος": "Ελένη"}
    ]
}
DEFAULT_TASKS = [
    {"Task": "1. Γενικός Έλεγχος", "Χρόνος_λεπτά": 2, "Υπεύθυνος": "Γραμμή 1"},
    {"Task": "2. Σύνθεση & Συσκευασία", "Χρόνος_λεπτά": 5, "Υπεύθυνος": "Γραμμή 2"}
]

option = st.sidebar.selectbox("Επιλογή Οθόνης", ["Καρτέλα Project", "Όλα τα Υλικά Παραγωγής"])

if not procurement_df.empty:
    if option == "Καρτέλα Project":
        st.header("📋 Διαχείριση Παραγωγής ανά Project")
        
        projects_list = sorted([p for p in procurement_df["Project"].unique().tolist() if p != "-"])
        selected_project = st.selectbox("Επιλέξτε Project:", projects_list)
        
        filtered_df = procurement_df[procurement_df["Project"] == selected_project].copy()
        
        st.subheader(f"📦 Υλικά & Tasks για το {selected_project} ({len(filtered_df)} Υλικά)")
        
        total_project_hours = 0.0
        
        # Εμφάνιση κάθε Υλικού σε διαδραστική Κάρτα (Accordion)
        for idx, row in filtered_df.iterrows():
            item_id = row["ID"]
            material = row["Υλικό / Προϊόν"]
            gift_type = str(row["Είδος Δώρου"]).upper()
            qty = int(row["Ποσότητα"]) if str(row["Ποσότητα"]).isdigit() else 1
            status = row["Status Procurement"]
            
            # Επιλογή προτύπου tasks
            tasks = TASK_TEMPLATES.get(gift_type, DEFAULT_TASKS)
            
            # Υπολογισμός χρόνου για αυτό το υλικό
            unit_time = sum([t["Χρόνος_λεπτά"] for t in tasks])
            item_total_hours = round((unit_time * qty) / 60, 2)
            total_project_hours += item_total_hours
            
            # Τίτλος Κάρτας Υλικού
            card_title = f"🆔 {item_id} | {material} — (Ποσότητα: {qty} τμχ) | Status: {status}"
            
            with st.expander(card_title):
                col_info, col_tasks = st.columns([1, 2])
                
                with col_info:
                    st.markdown("**ℹ️ Στοιχεία Procurement**")
                    st.write(f"• **Είδος Δώρου:** {gift_type}")
                    st.write(f"• **Προμηθευτής:** {row['Προμηθευτής']}")
                    st.write(f"• **Ημ. Παράδοσης:** {row['Ημερομηνία Παράδοσης']}")
                    st.write(f"• **Αναμ. Παραλαβή:** {row['Αναμενόμενη Ημ. Παραλαβής']}")
                    st.metric("Εκτιμώμενος Χρόνος", f"{item_total_hours} Ώρες")

                with col_tasks:
                    st.markdown("**⚙️ Στάδια Παραγωγής & Checkboxes**")
                    for t_idx, t in enumerate(tasks):
                        c1, c2, c3, c4 = st.columns([0.1, 0.4, 0.25, 0.25])
                        is_done = c1.checkbox("", key=f"chk_{item_id}_{t_idx}")
                        c2.write(f"**{t['Task']}**")
                        c3.write(f"⏱️ {t['Χρόνος_λεπτά']} λ/τμχ")
                        c4.write(f"👤 {t['Υπεύθυνος']}")
        
        st.divider()
        st.success(f"🎯 **Συνολικός Χρόνος Παραγωγής για το Project {selected_project}: {round(total_project_hours, 1)} Ώρες**")

    elif option == "Όλα τα Υλικά Παραγωγής":
        st.header("📦 Λίστα Υλικών")
        st.dataframe(procurement_df, use_container_width=True, hide_index=True)
