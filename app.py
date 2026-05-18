import streamlit as st
import pdfplumber
import re
from docx import Document
from io import BytesIO
from datetime import datetime

st.title("COC Automation")

uploaded_file = st.file_uploader("Upload COT PDF", type="pdf")

if uploaded_file:
    text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"

    # Extract fields
    date_match = re.search(r'Date[:\s]*([0-9]{1,2}[-/][A-Za-z0-9]{1,3}[-/][0-9]{2,4})', text)
    customer_match = re.search(r'Customer\s*(.+)', text)
    po_match = re.search(r'PO Number[:\s]*([A-Za-z0-9\-]+)', text)

    part_no = re.search(r'Part No\.\s*(\S+)', text)
    part_name = re.search(r'Part Name\s*(.+)', text)
    rev = re.search(r'Drawing Rev\s*(\S+)', text)
    qty = re.search(r'Lot Qty\s*(\d+)', text)

    extracted_date = date_match.group(1) if date_match else datetime.today().strftime("%d %b %Y")
    extracted_customer = customer_match.group(1).strip() if customer_match else ""
    extracted_po = po_match.group(1).strip() if po_match else ""

    extracted_part_no = part_no.group(1).strip() if part_no else ""
    extracted_description = part_name.group(1).strip() if part_name else ""
    extracted_rev = rev.group(1).strip() if rev else ""
    extracted_qty = qty.group(1).strip() if qty else ""

    st.subheader("Extracted Data")

    st.write("Date:", extracted_date)
    st.write("Customer:", extracted_customer)
    st.write("PO Number:", extracted_po)
    st.write("Part No:", extracted_part_no)
    st.write("Description:", extracted_description)
    st.write("REV:", extracted_rev)
    st.write("QTY:", extracted_qty)

    if st.button("Generate COC"):
        doc = Document("COC Format.docx")

        # Replace text in paragraphs
        for para in doc.paragraphs:
            if para.text.strip().startswith("Date:"):
                para.text = f"Date: {extracted_date}"

            if para.text.strip().startswith("Customer:"):
                para.text = f"Customer: {extracted_customer}"

            if para.text.strip().startswith("PO Number:"):
                para.text = f"PO Number: {extracted_po}"

        # Fill item table
        table = doc.tables[0]

        # Use first empty row if it exists, otherwise add row
        if len(table.rows) > 1:
            row_cells = table.rows[1].cells
        else:
            row_cells = table.add_row().cells

        row_cells[0].text = "1"
        row_cells[1].text = extracted_part_no
        row_cells[2].text = extracted_description
        row_cells[3].text = extracted_rev
        row_cells[4].text = extracted_qty

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        file_name = f"COC_{extracted_part_no or 'Generated'}.docx"

        st.download_button(
            label="Download COC DOCX",
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
