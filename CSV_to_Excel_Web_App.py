import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="CSV to Excel Converter", layout="centered")

st.title("📄 CSV to Excel Converter")
st.write("Upload your CSV files and download Excel versions. Leading zeros and formatting are preserved.")

uploaded_files = st.file_uploader("Upload CSV file(s)", type=["csv"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        try:
            df = pd.read_csv(uploaded_file, dtype=str, encoding='utf-8', keep_default_na=False)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            output.seek(0)

            st.download_button(
                label=f"Download {uploaded_file.name.replace('.csv', '.xlsx')}",
                data=output,
                file_name=uploaded_file.name.replace('.csv', '.xlsx'),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Error processing {uploaded_file.name}: {e}")
