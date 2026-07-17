import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Delivery Planning Tool", layout="wide")

st.title("📦 Delivery Report & Planning Tool")

# Print styling
st.markdown("""
    <style>
        @media print {
            .stDataFrame {
                font-size: 12px;
            }
        }
    </style>
""", unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Ανέβασε ένα ή περισσότερα Excel αρχεία (Πρωινή Διαλογή)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    df_list = []

    for file in uploaded_files:
        temp_df = pd.read_excel(file)
        df_list.append(temp_df)

    df = pd.concat(df_list, ignore_index=True)

    st.success(f"Φορτώθηκαν {len(uploaded_files)} αρχεία")
    st.info(f"Συνολικές γραμμές αποστολών: {df.shape[0]}")

    required_columns = ["CneeAdd1", "Cnee", "PostalCd", "ShptWt"]

    for col in required_columns:
        if col not in df.columns:
            st.error(f"Η στήλη '{col}' δεν βρέθηκε στο Excel.")
            st.stop()

    # Cleaning
    df = df.dropna(subset=["CneeAdd1"])
    df["CneeAdd1"] = df["CneeAdd1"].astype(str).str.strip()

    df["PostalCd"] = (
        df["PostalCd"]
        .fillna("")
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    df["Postal_Group"] = df["PostalCd"].apply(
        lambda x: x if x.startswith("5") else "Διάφοροι"
    )

    df["ShptWt"] = pd.to_numeric(df["ShptWt"], errors="coerce").fillna(0)

    # Unique stops
    unique_stops = df.drop_duplicates(subset=["CneeAdd1"]).copy()

    # KPIs
    st.header("📊 Σύνοψη")

    col1, col2 = st.columns(2)
    col1.metric("Συνολικές Αποστολές", df.shape[0])
    col2.metric("Μοναδικές Στάσεις", unique_stops.shape[0])

    # Shipments per Address
    st.header("📍 Σύνολο Αποστολών ανά Διεύθυνση")

    address_counts = (
        df.groupby("CneeAdd1")
        .size()
        .reset_index(name="Σύνολο Αποστολών")
        .sort_values(by="Σύνολο Αποστολών", ascending=False)
    )

    st.dataframe(address_counts, width="stretch")

    # Group Deliveries >=5
    st.header("🚚 Ομαδικές Παραδόσεις (>=5 ίδια διεύθυνση & παραλήπτης)")

    grouped = (
        df.groupby(["Cnee", "CneeAdd1", "Postal_Group"])
        .agg(
            Ποσότητα=("CneeAdd1","size"),
            Συνολικά_Κιλά=("ShptWt","sum")
        )
        .reset_index()
    )

    grouped["Συνολικά_Κιλά"]=grouped["Συνολικά_Κιλά"].round(2)

    over5 = grouped[grouped["Ποσότητα"] >= 5] \
        .sort_values(by=["Ποσότητα","Συνολικά_Κιλά"], ascending=False)

    st.dataframe(over5, width="stretch")

    st.info(f"Σύνολο ομαδικών παραδόσεων: {over5.shape[0]}")

    # =====================================
    # Stops per Postal Code (PRINT TABLE)
    # =====================================

    st.header("📮 Στάσεις ανά Ταχυδρομικό Κώδικα")

    postal_summary = (
        unique_stops.groupby("Postal_Group")
        .size()
        .reset_index(name="Στάσεις")
        .sort_values(by="Postal_Group")
        .reset_index(drop=True)
    )

    # Δημιουργία πίνακα 4 στηλών για εκτύπωση
    rows = []
    half = (len(postal_summary) + 1) // 2

    for i in range(half):
        left_postal = postal_summary.iloc[i]["Postal_Group"]
        left_stops = postal_summary.iloc[i]["Στάσεις"]

        if i + half < len(postal_summary):
            right_postal = postal_summary.iloc[i + half]["Postal_Group"]
            right_stops = postal_summary.iloc[i + half]["Στάσεις"]
        else:
            right_postal = ""
            right_stops = ""

        rows.append([left_postal, left_stops, right_postal, right_stops])

    print_table = pd.DataFrame(
        rows,
        columns=["Postal Group", "Στάσεις", "Postal Group ", "Στάσεις "]
    )

    st.dataframe(print_table, width="stretch")

    # =====================================
    # Postal Code Filter
    # =====================================

    st.subheader("🔎 Επιλογή Ταχυδρομικών Κωδικών")

    postal_options = sorted(postal_summary["Postal_Group"])

    selected_codes = st.multiselect(
        "Επίλεξε έναν ή περισσότερους ΤΚ",
        options=postal_options
    )

    if selected_codes:

        filtered = unique_stops[
            unique_stops["Postal_Group"].isin(selected_codes)
        ]

        tk_summary = (
            filtered.groupby("Postal_Group")
            .size()
            .reset_index(name="Σύνολο Στάσεων")
            .sort_values(by="Postal_Group")
        )

        st.subheader("📮 Σύνολο Στάσεων ανά ΤΚ")
        st.dataframe(tk_summary, width="stretch")

        st.info(f"Συνολικές στάσεις επιλογής: {filtered.shape[0]}")

        with st.expander("📋 Δες τις διευθύνσεις"):
            st.dataframe(
                filtered[["Postal_Group", "Cnee", "CneeAdd1"]],
                width="stretch"
            )

    # =====================================
    # EXPORT EXCEL REPORT
    # =====================================

    st.header("📥 Export Excel Report")

    total_stops = unique_stops.shape[0]
    total_shipments = df.shape[0]

    output = BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:

        over5.to_excel(
            writer,
            sheet_name='Group Deliveries',
            index=False,
            startrow=4
        )

        worksheet = writer.sheets['Group Deliveries']
        worksheet.cell(row=1, column=1).value = "Συνολικές Αποστολές"
        worksheet.cell(row=1, column=2).value = total_shipments
        worksheet.cell(row=2, column=1).value = "Μοναδικές Στάσεις"
        worksheet.cell(row=2, column=2).value = total_stops

        postal_summary.to_excel(
            writer,
            sheet_name='Stops per TK',
            index=False
        )

    excel_data = output.getvalue()

    # 👉 ΗΜΕΡΟΜΗΝΙΑ ΣΤΟ ΟΝΟΜΑ
    today = datetime.now().strftime("%d-%m-%Y")
    filename = f"delivery_report_{today}.xlsx"

    st.download_button(
        label="📥 Κατέβασε Excel Report",
        data=excel_data,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("⬆ Ανέβασε Excel αρχεία για να ξεκινήσει το report.")
