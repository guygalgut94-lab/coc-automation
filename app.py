import streamlit as st
import pdfplumber
import re
from docx import Document
from io import BytesIO
from datetime import datetime

st.title("COC Automation")

uploaded_files = st.file_uploader(
    "Upload COT PDF files",
    type="pdf",
    accept_multiple_files=True
)

def extract_text_from_pdf(uploaded_file):
    text = ""
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                text += extracted + "\n"
    return text

def extract_item_data(text, file_name):
    customer_match = re.search(r'Customer\s*(.+?)\s*KSW Project ID', text)
    po_match = re.search(r'PO Number[:\s]*([A-Za-z0-9\-_]+)', text)
    part_no_match = re.search(r'Part No\.\s*(\S+)', text)
    desc_match = re.search(r'Part Name\s*(.*?)\s+Material', text)
    rev_match = re.search(r'Drawing Rev\s*(\d+)', text)
    qty_match = re.search(r'Lot Qty\s*(\d+)', text)

    return {
        "file_name": file_name,
        "customer": customer_match.group(1).strip() if customer_match else "",
        "po": po_match.group(1).strip() if po_match else "",
        "part_no": part_no_match.group(1).strip() if part_no_match else "",
        "description": desc_match.group(1).strip() if desc_match else "",
        "rev": rev_match.group(1).strip() if rev_match else "",
        "qty": qty_match.group(1).strip() if qty_match else ""
    }

if uploaded_files:
    items = []

    for uploaded_file in uploaded_files:
        text = extract_text_from_pdf(uploaded_file)
        item = extract_item_data(text, uploaded_file.name)
        items.append(item)

    extracted_date = datetime.today().strftime("%d %b %Y")

    first_customer = items[0]["customer"] if items else ""
    first_po = items[0]["po"] if items else ""

    st.subheader("Extracted Parts")

    st.dataframe([
        {
            "#": index + 1,
            "File Name": item["file_name"],
            "Customer": item["customer"],
            "PO Number": item["po"],
            "Part No": item["part_no"],
            "Description": item["description"],
            "REV": item["rev"],
            "QTY": item["qty"]
        }
        for index, item in enumerate(items)
    ])

    st.write("COC Date:", extracted_date)
    st.write("Customer:", first_customer)
    st.write("PO Number:", first_po)

    if st.button("Generate COC"):
        doc = Document("COC Format.docx")

        # Table 1 = top header table
        top_table = doc.tables[0]

        for row in top_table.rows:
            for cell in row.cells:
                if "Date:" in cell.text:
                    cell.text = f"Date: {extracted_date}"
                elif "Manufacturer:" in cell.text:
                    cell.text = "Manufacturer: Bconduct HK"
                elif "Customer:" in cell.text:
                    cell.text = f"Customer: {first_customer}"
                elif "PO Number:" in cell.text:
                    cell.text = f"PO Number: {first_po}"

        # Table 2 = parts table
        parts_table = doc.tables[1]

        # Remove existing empty rows under the header
        while len(parts_table.rows) > 1:
            row = parts_table.rows[1]
            row._element.getparent().remove(row._element)

        # Add one row per uploaded COT PDF
        for index, item in enumerate(items, start=1):
            row_cells = parts_table.add_row().cells
            row_cells[0].text = str(index)
            row_cells[1].text = item["part_no"]
            row_cells[2].text = item["description"]
            row_cells[3].text = item["rev"]
            row_cells[4].text = item["qty"]

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        file_name = f"COC_{first_po or first_customer or 'Generated'}.docx"

        st.download_button(
            label="Download COC DOCX",
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
