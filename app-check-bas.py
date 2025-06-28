import streamlit as st
import pandas as pd
import io
import csv

st.title("Leverblok Updater - Artikelnummer Match")

st.header("Upload Files")
excel_file = st.file_uploader("Upload 'check Bas.xlsx'", type=['xlsx'])
csv_file = st.file_uploader("Upload 'artikel-data-export-17-06-2025-12_29.csv'", type=['csv'])

if excel_file and csv_file:
    # Read files
    df_check = pd.read_excel(excel_file)
    df_source = pd.read_csv(csv_file, delimiter=';', dtype=str, quoting=csv.QUOTE_MINIMAL)

    # Ensure correct columns (assumes 'Artikelnummer' is column D in both, and 'Leverblok' is column O in Excel, BD in CSV)
    check_art_col = df_check.columns[3]        # D
    check_kleurnr_col = df_check.columns[5]    # F
    check_lev_col = df_check.columns[14]       # O
    source_art_col = df_source.columns[3]      # D
    source_kleurnr_col = df_source.columns[52] # BA
    source_lev_col = df_source.columns[55]     # BD

    # Standardize Artikelnummer: remove trailing '000' in Excel version
    df_check['Artikelnummer_match'] = df_check[check_art_col].astype(str).str.replace(r'000$', '', regex=True)
    df_source[source_art_col] = df_source[source_art_col].astype(str)

    # Build a lookup for the source data on Artikelnummer
    source_lookup = df_source.set_index(source_art_col)

    # Helper: Get Leverblok value based on Artikelnummer and Kleurnummer
    def get_leverblok_value(artnr, kleurnr):
        matches = df_source[
            (df_source[source_art_col] == artnr) &
            (df_source[source_kleurnr_col] == kleurnr)
        ]
        if not matches.empty:
            return matches.iloc[0][source_lev_col]
        return None

    # Build updated Leverblok column
    updated_leverblok = []
    for idx, row in df_check.iterrows():
        artnr = row['Artikelnummer_match']
        kleurnr = row[check_kleurnr_col]
        leverblok_val = get_leverblok_value(artnr, kleurnr)
        if leverblok_val is not None:
            updated_leverblok.append(leverblok_val)
        else:
            updated_leverblok.append(row[check_lev_col])  # keep original if no match

    df_check[check_lev_col] = updated_leverblok
    df_check.drop(columns=['Artikelnummer_match'], inplace=True)

    # Output to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_check.to_excel(writer, index=False)
    st.success("Matching and updating done! Download the updated Excel file below:")
    st.download_button("Download Updated Excel", data=output.getvalue(), file_name="check Bas updated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Please upload both files to continue.")
