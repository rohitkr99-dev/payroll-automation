import streamlit as st

st.title("Payroll Automation System")

uploaded_file = st.file_uploader(
    "Upload attendance file",
    type=["xls", "xlsx"]
)

if uploaded_file:
    st.success("File uploaded successfully")
