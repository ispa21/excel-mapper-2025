import streamlit as st
import pandas as pd
from difflib import get_close_matches
import io

# Page setup
st.set_page_config(page_title="Excel Mapper 2025", layout="wide")
st.title("📊 Generalized Excel Column Mapper")

st.markdown("Welcome! This tool helps you map columns from a **Source Excel file** to a **Target Template** and export the mapped data in your desired format.")

# --- Step 1: File Upload ---
st.markdown("## 📂 Step 1: Upload Files")

col1, col2 = st.columns(2)

with col1:
    source_file = st.file_uploader("🔁 Upload **Source Excel** (Raw Data)", type=["xlsx"], key="source")

with col2:
    target_file = st.file_uploader("📥 Upload **Target Template Excel** (Format You Want)", type=["xlsx"], key="target")

if source_file and target_file:
    source_df = pd.read_excel(source_file)
    target_df = pd.read_excel(target_file)

    st.toast("✅ Files loaded successfully!")

    # --- Step 2: Data Preview ---
    st.markdown("## 👀 Step 2: Preview Files")

    preview_col1, preview_col2 = st.columns(2)
    with preview_col1:
        st.markdown("### 🧾 Source Data")
        st.dataframe(source_df.head(30), use_container_width=True)

    with preview_col2:
        st.markdown("### 📋 Target Format")
        st.dataframe(target_df.head(30), use_container_width=True)

    # --- Step 3: Mapping Section ---
    st.markdown("## 🧩 Step 3: Map Columns")
    st.markdown("Map each **Target Column** to the appropriate **Source Column** using the dropdowns below.")

    source_cols = list(source_df.columns)
    target_cols = list(target_df.columns)
    column_mapping = {}

    for tgt_col in target_cols:
        suggested = get_close_matches(tgt_col, source_cols, n=1, cutoff=0.3)
        default_idx = source_cols.index(suggested[0]) + 1 if suggested else 0

        selected = st.selectbox(
            f"🔗 Map for Target Column: `{tgt_col}`",
            options=["-- Ignore --"] + source_cols,
            index=default_idx,
            key=f"map_{tgt_col}"
        )
        column_mapping[tgt_col] = None if selected == "-- Ignore --" else selected

    # --- Step 4: Generate Output ---
    st.markdown("## 🚀 Step 4: Generate Mapped Output")

    if st.button("🎉 Generate Excel File"):
        output_rows = []
        for _, row in source_df.iterrows():
            mapped_row = {}
            for tgt_col in target_cols:
                src_col = column_mapping[tgt_col]
                mapped_row[tgt_col] = row[src_col] if src_col in source_df.columns else None
            output_rows.append(mapped_row)

        output_df = pd.DataFrame(output_rows, columns=target_cols)

        # Save to memory
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            output_df.to_excel(writer, index=False)
        buffer.seek(0)

        st.success("✅ Mapped Excel file is ready!")
        st.download_button(
            label="📥 Download Mapped Excel",
            data=buffer,
            file_name="Mapped_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload both the Source and Target Excel files to proceed.")
