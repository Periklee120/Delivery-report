import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Delivery Planning Tool",
    layout="wide"
)

st.title("📦 Delivery Report & Planning Tool")

# ============================================================
# PRINT STYLING
# ============================================================

st.markdown("""
    <style>
        @media print {
            .stDataFrame {
                font-size: 12px;
            }
        }
    </style>
""", unsafe_allow_html=True)


# ============================================================
# UPLOAD EXCEL FILES
# ============================================================

uploaded_files = st.file_uploader(
    "Ανέβασε ένα ή περισσότερα Excel αρχεία (Πρωινή Διαλογή)",
    type=["xlsx"],
    accept_multiple_files=True
)


# ============================================================
# MAIN PROCESS
# ============================================================

if uploaded_files:

    df_list = []

    required_columns = [
        "CneeAdd1",
        "Cnee",
        "PostalCd",
        "ShptWt"
    ]

    # Μετρητές για ενημέρωση χρήστη
    total_sheets = 0
    accepted_sheets = 0
    ignored_sheets = []

    # ========================================================
    # READ ALL EXCEL FILES + ALL TABS
    # ========================================================

    for file in uploaded_files:

        try:
            # Διαβάζει ΟΛΑ τα tabs του Excel, αναγκάζοντας όλες τις στήλες να διαβαστούν ως STRING
            # για να μην μπερδεύονται τα αριθμητικά timestamps (π.χ. 46274) με τις ημερομηνίες.
            excel_sheets = pd.read_excel(
                file,
                sheet_name=None,
                dtype=str
            )

            for sheet_name, temp_df in excel_sheets.items():

                total_sheets += 1

                # Έλεγχος αν το συγκεκριμένο tab έχει
                # όλες τις απαραίτητες στήλες
                if all(
                    col in temp_df.columns
                    for col in required_columns
                ):

                    df_list.append(temp_df)
                    accepted_sheets += 1

                else:
                    ignored_sheets.append(
                        f"{file.name} → {sheet_name}"
                    )

        except Exception as e:

            st.error(
                f"❌ Πρόβλημα κατά την ανάγνωση του αρχείου "
                f"'{file.name}': {e}"
            )

    # ========================================================
    # CHECK IF ANY VALID SHEETS WERE FOUND
    # ========================================================

    if not df_list:

        st.error(
            "❌ Δεν βρέθηκε κανένα tab με τις απαραίτητες στήλες:"
        )

        st.write(
            ", ".join(required_columns)
        )

        if ignored_sheets:

            st.warning("Τα tabs που αγνοήθηκαν:")

            for sheet in ignored_sheets:
                st.write(f"• {sheet}")

        st.stop()


    # ========================================================
    # COMBINE ALL VALID SHEETS
    # ========================================================

    df = pd.concat(
        df_list,
        ignore_index=True
    )


    # ========================================================
    # UPLOAD SUMMARY
    # ========================================================

    st.success(
        f"✅ Φορτώθηκαν {len(uploaded_files)} Excel αρχεία"
    )

    st.info(
        f"📑 Ελέγχθηκαν {total_sheets} tabs — "
        f"χρησιμοποιήθηκαν {accepted_sheets}"
    )

    if ignored_sheets:

        with st.expander(
            "ℹ️ Tabs που αγνοήθηκαν"
        ):

            for sheet in ignored_sheets:
                st.write(f"• {sheet}")


    st.info(
        f"📦 Συνολικές γραμμές αποστολών: {df.shape[0]}"
    )


    # ========================================================
    # CLEANING & UNIFORMITY (Διορθωμένο για απόλυτη ταύτιση)
    # ========================================================

    df = df.dropna(
        subset=["CneeAdd1"]
    )

    # 1. Μετατροπή σε κείμενο και αφαίρεση κενών στην αρχή/τέλος
    df["CneeAdd1"] = (
        df["CneeAdd1"]
        .astype(str)
        .str.strip()
    )

    # 2. Μετατροπή σε ΚΕΦΑΛΑΙΑ για να μην υπάρχει θέμα με πεζά/κεφαλαία
    df["CneeAdd1"] = df["CneeAdd1"].str.upper()

    # 3. Αφαίρεση ελληνικών τόνων για απόλυτη ταύτιση διευθύνσεων
    tonos_map = {
        "Ά": "Α", "Έ": "Ε", "Ή": "Η", "Ί": "Ι", "Ό": "Ο", "Ύ": "Υ", "Ώ": "Ω",
        "Ϊ": "Ι", "Ϋ": "Υ"
    }
    for t, r in tonos_map.items():
        df["CneeAdd1"] = df["CneeAdd1"].str.replace(t, r, regex=False)

    # 4. Αντικατάσταση πολλαπλών κενών ανάμεσα στις λέξεις με ένα μόνο κενό
    df["CneeAdd1"] = df["CneeAdd1"].str.replace(r"\s+", " ", regex=True)


    # ========================================================
    # POSTAL CODE CLEANING
    # ========================================================

    df["PostalCd"] = (
        df["PostalCd"]
        .fillna("")
        .astype(str)
        .str.replace(
            ".0",
            "",
            regex=False
        )
        .str.strip()
    )


    # ========================================================
    # POSTAL GROUP
    # ========================================================

    df["Postal_Group"] = df["PostalCd"].apply(
        lambda x:
            x if x.startswith("5")
            else "Διάφοροι"
    )


    # ========================================================
    # WEIGHT
    # ========================================================

    df["ShptWt"] = pd.to_numeric(
        df["ShptWt"],
        errors="coerce"
    ).fillna(0)


    # ========================================================
    # UNIQUE STOPS
    # ========================================================

    unique_stops = (
        df
        .drop_duplicates(
            subset=["CneeAdd1"]
        )
        .copy()
    )


    # ========================================================
    # KPIs
    # ========================================================

    st.header("📊 Σύνοψη")

    col1, col2 = st.columns(2)

    col1.metric(
        "Συνολικές Αποστολές",
        df.shape[0]
    )

    col2.metric(
        "Μοναδικές Στάσεις",
        unique_stops.shape[0]
    )


    # ========================================================
    # SHIPMENTS PER ADDRESS
    # ========================================================

    st.header(
        "📍 Σύνολο Αποστολών ανά Διεύθυνση"
    )

    address_counts = (
        df.groupby("CneeAdd1")
        .size()
        .reset_index(
            name="Σύνολο Αποστολών"
        )
        .sort_values(
            by="Σύνολο Αποστολών",
            ascending=False
        )
    )

    st.dataframe(
        address_counts,
        use_container_width=True
    )


    # ========================================================
    # GROUP DELIVERIES (Δυναμικό φίλτρο με Slider)
    # ========================================================

    st.header("🚚 Ομαδικές Παραδόσεις")
    
    # Slider για να επιλέγεις εσύ το όριο παραδόσεων στην οθόνη
    min_deliveries = st.slider(
        "Ελάχιστος αριθμός αποστολών στην ίδια διεύθυνση για να θεωρηθεί ομαδική:", 
        min_value=2, 
        max_value=10, 
        value=5
    )

    grouped = (
        df
        .groupby(
            [
                "CneeAdd1",
                "Postal_Group"
            ]
        )
        .agg(
            Ποσότητα=(
                "CneeAdd1",
                "size"
            ),
            Συνολικά_Κιλά=(
                "ShptWt",
                "sum"
            )
        )
        .reset_index()
    )

    grouped["Συνολικά_Κιλά"] = (
        grouped["Συνολικά_Κιλά"]
        .round(2)
    )

    over_limit = (
        grouped[
            grouped["Ποσότητα"] >= min_deliveries
        ]
        .sort_values(
            by=[
                "Ποσότητα",
                "Συνολικά_Κιλά"
            ],
            ascending=False
        )
    )

    st.dataframe(
        over_limit,
        use_container_width=True
    )

    st.info(
        f"Σύνολο ομαδικών παραδόσεων (με όριο >= {min_deliveries}): "
        f"{over_limit.shape[0]}"
    )


    # ========================================================
    # STOPS PER POSTAL CODE
    # ========================================================

    st.header(
        "📮 Στάσεις ανά Ταχυδρομικό Κώδικα"
    )

    postal_summary = (
        unique_stops
        .groupby("Postal_Group")
        .size()
        .reset_index(
            name="Στάσεις"
        )
        .sort_values(
            by="Postal_Group"
        )
        .reset_index(
            drop=True
        )
    )


    # ========================================================
    # PRINT TABLE - 4 COLUMNS
    # ========================================================

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

        rows.append({
            "ΤΚ (Α)": left_postal,
            "Στάσεις (Α)": left_stops,
            "ΤΚ (Β)": right_postal,
            "Στάσεις (Β)": right_stops
        })

    print_df = pd.DataFrame(rows)
    st.dataframe(print_df, use_container_width=True)


    # ========================================================
    # EXPORT TO EXCEL FUNCTIONALITY
    # ========================================================
    st.header("💾 Εξαγωγή Αποτελεσμάτων")
    
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
