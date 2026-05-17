import streamlit as st
import pdfplumber
import re

st.title("COC Automation")

uploaded_file = st.file_uploader("Upload COT PDF", type="pdf")

if uploaded_file:

    text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"

    st.subheader("Extracted Text")
    st.text(text[:3000])

    # Extract fields
    part_no = re.search(r'Part No\.\s*(\S+)', text)
    part_name = re.search(r'Part Name\s*(.+)', text)
    rev = re.search(r'Drawing Rev\s*(\S+)', text)
    qty = re.search(r'Lot Qty\s*(\d+)', text)

    st.subheader("Extracted Data")

    st.write("Part No:", part_no.group(1) if part_no else "Not found")
    st.write("Description:", part_name.group(1) if part_name else "Not found")
    st.write("REV:", rev.group(1) if rev else "Not found")
    st.write("QTY:", qty.group(1) if qty else "Not found")
