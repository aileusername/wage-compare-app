import streamlit as st
import pandas as pd
import os
from tempfile import TemporaryDirectory

from wage_compare import extract_data, compare_variants, apply_excel_styling, get_rev_label

st.set_page_config(page_title="Wage Comparator", layout="centered")
st.title("📊 Wage Determination Comparator")

st.write("Upload exactly two wage .txt files (e.g. r0 vs r1) to compare.")

uploaded_files = st.file_uploader("Choose 2 wage files", type="txt", accept_multiple_files=True)

if uploaded_files and len(uploaded_files) == 2:
    with TemporaryDirectory() as temp_dir:
        # Save uploaded files locally
        file_paths = []
        for f in uploaded_files:
            path = os.path.join(temp_dir, f.name)
            with open(path, "wb") as out:
                out.write(f.read())
            file_paths.append(path)

        file_old, file_new = sorted(file_paths, key=os.path.getmtime)

        rev1 = get_rev_label(file_old, "Version_1")
        rev2 = get_rev_label(file_new, "Version_2")

        df1 = extract_data(file_old)
        df2 = extract_data(file_new)
        df_diff = compare_variants(df1, df2, rev1, rev2)

        output_path = os.path.join(temp_dir, "wage_comparison.xlsx")
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df1.to_excel(writer, index=False, sheet_name=rev1)
            df2.to_excel(writer, index=False, sheet_name=rev2)
            if not df_diff.empty:
                df_diff.to_excel(writer, index=False, sheet_name="Changes")

        apply_excel_styling(output_path)

        st.success("✅ Comparison complete. Download the Excel file below:")
        with open(output_path, "rb") as f:
            st.download_button("📥 Download Excel", f, file_name="wage_comparison.xlsx")
else:
    st.info("Drag in exactly two .txt files to begin comparison.")