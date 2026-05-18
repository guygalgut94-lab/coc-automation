import streamlit as st
import pdfplumber
import re
from docx import Document
from docx.shared import Inches
from io import BytesIO
from datetime import datetime
import os

st.title("COC Automation")

uploaded_files = st.file_uploader(
    "Upload COT PDF files",
    type="pdf",
    accept_multiple_files=True
)

worker = st.selectbox(
    "Certified by",
    ["Guy", "Ilan" ]
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
    po_match = re.search(r'Customer ID\s*([A-Za-z0-9\/\-_]+)', text)
    part_no_match = re.search(r'Part No\.\s*(\S+)', text)
    desc_match = re.search(r'Part Name\s*(.*?)\s+Material', text)
    rev_match = re.search(r'Drawing Rev\s*(\d+)', text)
    qty_match = re.search(r'Lot Qty\s*(\d+)', text)

    return {
        "file_name": file_name,
        "po": po_match.group(1).strip() if po_match else "",
        "part_no": part_no_match.group(1).strip() if part_no_match else "",
        "description": desc_match.group(1).strip() if desc_match else "",
        "rev": rev_match.group(1).strip() if rev_match else "",
        "qty": qty_match.group(1).strip() if qty_match else ""
    }

def fill_value_after_label(table, label, value):
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if label in cell.text:
                if i + 1 < len(row.cells):
                    row.cells[i + 1].text = value
                else:
                    cell.text = f"{label} {value}"
                return

def insert_signature(doc, worker_name):
    signature_path = f"signatures/{worker_name}.png"

    if not os.path.exists(signature_path):
        st.warning(f"Signature image not found: {signature_path}")
        return

    for para in doc.paragraphs:
        if "Certified by:" in para.text:
            para.text = "Certified by: "
            run = para.add_run()
            run.add_picture(signature_path, width=Inches(1.2))
            return

    st.warning("Could not find 'Certified by:' in the COC template.")

if uploaded_files:
    items = []

    for uploaded_file in uploaded_files:
        text = extract_text_from_pdf(uploaded_file)
        item = extract_item_data(text, uploaded_file.name)
        items.append(item)

    coc_date = datetime.today().strftime("%d %b %Y")
    default_po = items[0]["po"] if items else ""

    st.subheader("Header Information")

    customer_name = st.text_input("Customer Name", value="KSW")
    po_number = st.text_input("PO Number", value=default_po)

    st.subheader("Extracted Parts")

    st.dataframe([
        {
            "#": index + 1,
            "File Name": item["file_name"],
            "Part No": item["part_no"],
            "Description": item["description"],
            "REV": item["rev"],
            "QTY": item["qty"]
        }
        for index, item in enumerate(items)
    ])

    if st.button("Generate COC"):
        doc = Document("COC Format.docx")

        top_table = doc.tables[0]
        parts_table = doc.tables[1]

        fill_value_after_label(top_table, "Date:", coc_date)
        fill_value_after_label(top_table, "Manufacturer:", "Bconduct HK")
        fill_value_after_label(top_table, "Customer:", customer_name)
        fill_value_after_label(top_table, "PO Number:", po_number)

        while len(parts_table.rows) > 1:
            row = parts_table.rows[1]
            row._element.getparent().remove(row._element)

        for index, item in enumerate(items, start=1):
            row_cells = parts_table.add_row().cells
            row_cells[0].text = str(index)
            row_cells[1].text = item["part_no"]
            row_cells[2].text = item["description"]
            row_cells[3].text = item["rev"]
            row_cells[4].text = item["qty"]

        insert_signature(doc, worker)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        file_name = f"COC_{po_number or customer_name or 'Generated'}.docx"

        st.download_button(
            label="Download COC DOCX",
            data=buffer,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
