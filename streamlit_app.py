import streamlit as st
from generate_contracts import generate_contracts
import os
import zipfile

st.set_page_config(page_title="Contract Generator", layout="centered")

st.title("📄 Faculty Contract Generator")

excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
output_dir = "generated_contracts"

if excel_file:
    with open("uploaded_data.xlsx", "wb") as f:
        f.write(excel_file.read())

    st.success("✅ File uploaded successfully.")

    if st.button("Generate Contracts"):
        result = generate_contracts("uploaded_data.xlsx", output_dir)
        st.success(result)

        # Zip files for download
        zip_path = "contracts.zip"
        with zipfile.ZipFile(zip_path, "w") as zipf:
            for root, _, files in os.walk(output_dir):
                for file in files:
                    zipf.write(os.path.join(root, file), arcname=file)

        with open(zip_path, "rb") as f:
            st.download_button("📦 Download All Contracts", f, file_name="contracts.zip")
