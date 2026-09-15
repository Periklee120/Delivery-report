import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime


# ============================================================
# PAGE CONFIGURATION
# ============================================================

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
# FILE UPLOADER
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

    # ========================================================
    # REQUIRED COLUMNS
    # ========================================================

    required_columns = [
        "CneeAdd1",
        "Cnee",
        "PostalCd",
        "ShptWt"
    ]

    total_sheets = 0
    accepted_sheets = 0
    ignored_sheets = []


    # ========================================================
    # READ ALL EXCEL FILES + ALL TABS
    # ========================================================

    for file in uploaded_files:

        try:

            # Διαβάζει ΟΛΑ τα tabs
            excel_sheets = pd.read_excel(
                file,
                sheet_name=None
            )

            for sheet_name, temp_df in excel_sheets.items():

                total_sheets += 1

                # ------------------------------------------------
                # Έλεγχος απαραίτητων στηλών
                # ------------------------------------------------

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
                f"❌ Πρόβλημα κατά την ανάγνωση του "
                f"αρχείου '{file.name}': {e}"
            )


    # ========================================================
    # CHECK VALID SHEETS
    # ========================================================

    if not df_list:

        st.error(
            "❌ Δεν βρέθηκε κανένα tab με τις απαραίτητες στήλες:"
        )

        st.write(
            ", ".join(required_columns)
        )

        if ignored_sheets:

            st.warning(
                "Τα tabs που αγνοήθηκαν:"
            )

            for sheet in ignored_sheets:
                st.write(f"• {sheet}")

        st.stop()


    # ========================================================
    # COMBINE ALL VALID TABS
    # ========================================================

    df = pd.concat(
        df_list,
        ignore_index=True
    )


    # ========================================================
    # FILE / TAB SUMMARY
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


    # ========================================================
    # REMOVE ROWS WITHOUT ADDRESS
    # ========================================================

    df = df.dropna(
        subset=["CneeAdd1"]
    ).copy()


    # ========================================================
    # CLEAN ADDRESS
    # ========================================================

    df["CneeAdd1"] = (
        df["CneeAdd1"]
        .astype(str)
        .str.strip()
    )


    # ========================================================
    # CLEAN RECIPIENT
    # ========================================================

    df["Cnee"] = (
        df["Cnee"]
        .fillna("")
        .astype(str)
        .str.strip()
    )


    # ========================================================
    # CLEAN POSTAL CODE
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
    # CLEAN WEIGHT
    # ========================================================

    df["ShptWt"] = pd.to_numeric(
        df["ShptWt"],
        errors="coerce"
    ).fillna(0)


    # ========================================================
    # TOTAL SHIPMENTS
    # ========================================================

    total_shipments = df.shape[0]


    # ========================================================
    # UNIQUE STOPS
    # ========================================================
    #
    # ΣΗΜΑΝΤΙΚΟ:
    #
    # Μία στάση =
    #
    #       CneeAdd1 + Cnee
    #
    # Ίδια διεύθυνση + ίδιος παραλήπτης
    #       -> 1 στάση
    #
    # Ίδια διεύθυνση + διαφορετικός παραλήπτης
    #       -> διαφορετικές στάσεις
    #
    # Το BulkLookup ΔΕΝ χρησιμοποιείται εδώ.
    #
    # ========================================================

    unique_stops = (
        df
        .drop_duplicates(
            subset=[
                "CneeAdd1",
                "Cnee"
            ]
        )
        .copy()
    )


    # ========================================================
    # TOTAL STOPS
    # ========================================================

    total_stops = unique_stops.shape[0]


    # ========================================================
    # SUMMARY
    # ========================================================

    st.header("📊 Σύνοψη")

    col1, col2 = st.columns(2)

    col1.metric(
        "Συνολικές Αποστολές",
        total_shipments
    )

    col2.metric(
        "Μοναδικές Στάσεις",
        total_stops
    )


    # ========================================================
    # SHIPMENTS PER ADDRESS
    # ========================================================

    st.header(
        "📍 Σύνολο Αποστολών ανά Διεύθυνση"
    )

    address_counts = (
        df
        .groupby("CneeAdd1")
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
    # GROUP DELIVERIES >= 5
    # ========================================================
    #
    # Εδώ ΔΕΝ χρησιμοποιούμε unique_stops.
    #
    # Μετράμε τα πραγματικά δέματα από το αρχικό df.
    #
    # Ίδιος παραλήπτης + ίδια διεύθυνση:
    # 5 ή περισσότερα δέματα = ομαδική παράδοση.
    #
    # Διαφορετικά BulkLookup = διαφορετικά δέματα,
    # αλλά παραμένουν μία ομαδική παράδοση εφόσον
    # Cnee + CneeAdd1 είναι ίδια.
    #
    # ========================================================

    st.header(
        "🚚 Ομαδικές Παραδόσεις "
        "(>=5 ίδια διεύθυνση & παραλήπτης)"
    )


    grouped = (
        df
        .groupby(
            [
                "Cnee",
                "CneeAdd1",
                "Postal_Group"
            ],
            dropna=False
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


    # Στρογγυλοποίηση κιλών

    grouped["Συνολικά_Κιλά"] = (
        grouped["Συνολικά_Κιλά"]
        .round(2)
    )


    # Μόνο ομάδες με 5+ δέματα

    over5 = (
        grouped[
            grouped["Ποσότητα"] >= 5
        ]
        .sort_values(
            by=[
                "Ποσότητα",
                "Συνολικά_Κιλά"
            ],
            ascending=False
        )
        .reset_index(drop=True)
    )


    st.dataframe(
        over5,
        use_container_width=True
    )


    st.info(
        f"Σύνολο ομαδικών παραδόσεων: "
        f"{over5.shape[0]}"
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

    half = (
        len(postal_summary) + 1
    ) // 2


    for i in range(half):

        # LEFT

        left_postal = (
            postal_summary.iloc[i][
                "Postal_Group"
            ]
        )

        left_stops = (
            postal_summary.iloc[i][
                "Στάσεις"
            ]
        )


        # RIGHT

        if i + half < len(postal_summary):

            right_postal = (
                postal_summary.iloc[
                    i + half
                ][
                    "Postal_Group"
                ]
            )

            right_stops = (
                postal_summary.iloc[
                    i + half
                ][
                    "Στάσεις"
                ]
            )

        else:

            right_postal = ""
            right_stops = ""


        rows.append(
            [
                left_postal,
                left_stops,
                right_postal,
                right_stops
            ]
        )


    print_table = pd.DataFrame(
        rows,
        columns=[
            "Postal Group",
            "Στάσεις",
            "Postal Group ",
            "Στάσεις "
        ]
    )


    st.dataframe(
        print_table,
        use_container_width=True
    )


    # ========================================================
    # POSTAL CODE FILTER
    # ========================================================

    st.subheader(
        "🔎 Επιλογή Ταχυδρομικών Κωδικών"
    )


    postal_options = sorted(
        postal_summary[
            "Postal_Group"
        ].tolist()
    )


    selected_codes = st.multiselect(
        "Επίλεξε έναν ή περισσότερους ΤΚ",
        options=postal_options
    )


    if selected_codes:

        filtered = unique_stops[
            unique_stops[
                "Postal_Group"
            ].isin(selected_codes)
        ].copy()


        # --------------------------------------------
        # TK SUMMARY
        # --------------------------------------------

        tk_summary = (
            filtered
            .groupby("Postal_Group")
            .size()
            .reset_index(
                name="Σύνολο Στάσεων"
            )
            .sort_values(
                by="Postal_Group"
            )
        )


        st.subheader(
            "📮 Σύνολο Στάσεων ανά ΤΚ"
        )


        st.dataframe(
            tk_summary,
            use_container_width=True
        )


        st.info(
            f"Συνολικές στάσεις επιλογής: "
            f"{filtered.shape[0]}"
        )


        # --------------------------------------------
        # ADDRESS LIST
        # --------------------------------------------

        with st.expander(
            "📋 Δες τις διευθύνσεις"
        ):

            st.dataframe(
                filtered[
                    [
                        "Postal_Group",
                        "Cnee",
                        "CneeAdd1"
                    ]
                ],
                use_container_width=True
            )


    # ========================================================
    # ROUTE MAP
    # ========================================================

    route_map = {}


    def add(route_name, codes):

        for code in codes:

            route_map[
                str(code)
            ] = route_name


    # ========================================================
    # ROUTE 1
    # ========================================================

    add(
        "ΚΑΛΑΜΑΡΙΑ, ΝΤΕΠΩ, ΧΑΡΙΛΑΟΥ, ΑΝΑΛΗΨΗ, ΤΟΥΜΠΑ",
        [
            54248,
            54249,
            54250,
            54351,
            54352,
            54453,
            54454,
            55534,
            54646,
            54655,
            54638,
            54639,
            54641,
            54642,
            54643,
            54644,
            54645,
            55131,
            55132,
            55133,
            55134,
            55135
        ]
    )


    # ========================================================
    # ROUTE 2
    # ========================================================

    add(
        "ΚΕΝΤΡΟ, ΑΝΩ ΠΟΛΗ, 40ΕΚΚΛΗΣΙΕΣ, ΤΡΙΑΝΔΡΙΑ",
        [
            54621,
            54622,
            54623,
            54624,
            54625,
            54626,
            54630,
            54631,
            54632,
            54633,
            54634,
            54635,
            54636,
            55337
        ]
    )


    # ========================================================
    # ROUTE 3
    # ========================================================

    add(
        "ΠΥΛΑΙΑ, ΠΑΝΟΡΑΜΑ, ΘΕΡΜΗ",
        [
            55535,
            55536,
            55236,
            57001
        ]
    )


    # ========================================================
    # ROUTE 4
    # ========================================================

    add(
        "ΕΥΟΣΜΟΣ, ΣΤΑΥΡΟΥΠΟΛΗ, ΣΥΚΙΕΣ, ΝΕΑΠΟΛΗ, "
        "ΑΓΙΟΣ ΠΑΥΛΟΣ, ΠΟΛΙΧΝΗ, ΩΡΑΙΟΚΑΣΤΡΟ",
        [
            55438,
            56224,
            56225,
            56226,
            56238,
            56430,
            56431,
            56437,
            56532,
            56533,
            56625,
            56626,
            56727,
            56728,
            57013
        ]
    )


    # ========================================================
    # ROUTE 5
    # ========================================================

    add(
        "ΓΙΑΝΝΙΤΣΩΝ, ΑΜΠΕΛΟΚΗΠΟΙ, ΜΕΝΕΜΕΝΗ",
        [
            54627,
            54628,
            56121,
            56122,
            56123
        ]
    )


    # ========================================================
    # ROUTE 6
    # ========================================================

    add(
        "ΚΟΡΔΕΛΙΟ, ΚΑΛΟΧΩΡΙ",
        [
            56334,
            57009
        ]
    )


    # ========================================================
    # ROUTE 7
    # ========================================================

    add(
        "ΣΙΝΔΟΣ, ΙΩΝΙΑ, ΔΙΑΒΑΤΑ",
        [
            57008,
            57022,
            57400,
            54500
        ]
    )


    # ========================================================
    # ROUTE 8
    # ========================================================

    add(
        "ΧΩΡΙΑ, ΕΥΚΑΡΠΙΑ",
        [
            57018,
            57200,
            56429
        ]
    )


    # ========================================================
    # ASSIGN ROUTE
    # ========================================================

    unique_stops["Route"] = (
        unique_stops[
            "Postal_Group"
        ]
        .map(route_map)
        .fillna("ΛΟΙΠΑ")
    )


    # ========================================================
    # ROUTES SUMMARY
    # ========================================================

    routes_summary = (
        unique_stops
        .groupby("Route")
        .size()
        .reset_index(
            name="Στάσεις"
        )
        .sort_values(
            "Route"
        )
        .reset_index(drop=True)
    )


    # ========================================================
    # EXPORT EXCEL REPORT
    # ========================================================

    st.header(
        "📥 Export Excel Report"
    )


    output = BytesIO()


    # ΣΗΜΑΝΤΙΚΟ:
    # Εδώ χρησιμοποιούμε "as writer"

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:


        # ====================================================
        # GROUP DELIVERIES
        # ====================================================

        over5.to_excel(
            writer,
            sheet_name="Group Deliveries",
            index=False,
            startrow=4
        )


        worksheet = writer.sheets[
            "Group Deliveries"
        ]


        worksheet.cell(
            row=1,
            column=1
        ).value = (
            "Συνολικές Αποστολές"
        )


        worksheet.cell(
            row=1,
            column=2
        ).value = total_shipments


        worksheet.cell(
            row=2,
            column=1
        ).value = (
            "Μοναδικές Στάσεις"
        )


        worksheet.cell(
            row=2,
            column=2
        ).value = total_stops


        # ====================================================
        # STOPS PER TK
        # ====================================================

        postal_summary.to_excel(
            writer,
            sheet_name="Stops per TK",
            index=False
        )


        # ====================================================
        # ROUTES
        # ====================================================

        routes_summary.to_excel(
            writer,
            sheet_name="Routes",
            index=False
        )


    # ========================================================
    # EXCEL DATA
    # ========================================================

    excel_data = output.getvalue()


    # ========================================================
    # FILE NAME
    # ========================================================

    today = datetime.now().strftime(
        "%d-%m-%Y"
    )


    filename = (
        f"delivery_report_{today}.xlsx"
    )


    # ========================================================
    # DOWNLOAD
    # ========================================================

    st.download_button(
        label="📥 Κατέβασε Excel Report",
        data=excel_data,
        file_name=filename,
        mime=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


# ============================================================
# NO FILE UPLOADED
# ============================================================

else:

    st.info(
        "⬆ Ανέβασε Excel αρχεία για να ξεκινήσει το report."
    )
