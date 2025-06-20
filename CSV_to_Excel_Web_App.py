import streamlit as st
import pandas as pd
import io
import zipfile
import csv
from datetime import datetime
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(
    page_title="CSV to Excel Pro Converter",
    page_icon="📊",
    layout="centered"
)

# Logo and header
st.markdown("""
    <div style='text-align: center;'>
        <img src='https://upload.wikimedia.org/wikipedia/commons/thumb/a/ae/Excel_icon_%282019%29.svg/120px-Excel_icon_%282019%29.svg.png' width='60'/>
        <h2 style='margin-bottom: 5px;'>CSV to Excel Pro Converter</h2>
        <p style='font-size: 15px;'>Smart upload, preview, Excel conversion, and ZIP download</p>
    </div>
""", unsafe_allow_html=True)

# Helper: detect delimiter
def detect_delimiter(file):
    sample = file.read(2048).decode("utf-8", errors="ignore")
    file.seek(0)
    sniffer = csv.Sniffer()
    try:
        return sniffer.sniff(sample).delimiter
    except Exception:
        return ","  # default fallback

# Upload
uploaded_files = st.file_uploader(
    "Upload CSV file(s)",
    type=["csv"],
    accept_multiple_files=True,
    help="Supports comma, semicolon, or tab-delimited CSVs."
)

password = st.text_input("🔐 Enter access password to convert files:", type="password")
if password and password != "1234":
    st.warning("❌ Incorrect password. Please try again.")
    st.stop()

if uploaded_files and password == "1234":
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zipf:
        for uploaded_file in uploaded_files:
            st.markdown(f"### 📂 {uploaded_file.name}")
            if uploaded_file.size > 5 * 1024 * 1024:
                st.warning("⚠️ File too large (max 5MB). Skipping.")
                continue

            try:
                delimiter = detect_delimiter(uploaded_file)
                df = pd.read_csv(uploaded_file, dtype=str, encoding='utf-8', delimiter=delimiter, keep_default_na=False)

                if df.empty:
                    st.warning("⚠️ File is empty. Skipping.")
                    continue

                st.dataframe(df.head())

                # Data summary
                st.markdown("**📊 Quick Analysis:**")
                col_count = len(df.columns)
                row_count = len(df)
                nulls = df.isnull().sum().sum()
                st.markdown(f"- Rows: `{row_count}`  Columns: `{col_count}`  Null values: `{nulls}`")

                # Chart preview
                if col_count <= 20:
                    fig, ax = plt.subplots(figsize=(8, 3))
                    sns.heatmap(df.isnull(), cbar=False, yticklabels=False, cmap="YlGnBu", ax=ax)
                    st.pyplot(fig)

                output = io.BytesIO()
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                excel_filename = uploaded_file.name.replace('.csv', f'_{timestamp}.xlsx')

                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                output.seek(0)

                st.download_button(
                    label=f"⬇️ Download {excel_filename}",
                    data=output,
                    file_name=excel_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                # Add to zip for bulk download
                zipf.writestr(excel_filename, output.read())

            except Exception as e:
                st.error(f"❌ Error processing {uploaded_file.name}: {e}")

    zip_buffer.seek(0)
    if len(zipf.namelist()) > 1:
        st.download_button(
            label="⬇️ Download All as ZIP",
            data=zip_buffer,
            file_name=f"converted_excels_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip"
        )
