import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.title("Document Data Extractor")

# 1. Allow PDF and Excel file types in the uploader
uploaded_files = st.file_uploader(
    "Upload Invoice PDFs or Excel Files",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True,
)


def process_pdf(uploaded_file):
    """Extract structured data from uploaded PDF file."""
    file_bytes = uploaded_file.read()
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    # Extract Header Fields from Page 1
    page1_text = doc[0].get_text("text")

    inv_match = re.search(r"Invoice No\.\s*([A-Z0-9/-]+)", page1_text)
    invoice_no = inv_match.group(1).strip() if inv_match else None

    po_match = re.search(
        r"Buyer'?s Order No\.\s*(\d+)", page1_text, re.IGNORECASE
    )
    po_number = po_match.group(1).strip() if po_match else None

    dest_match = re.search(r"Destination\s*([A-Za-z]+)", page1_text)
    destination = dest_match.group(1).strip() if dest_match else "Bangalore"

    # Extract Line Items
    file_items = []
    for page in doc:
        text = page.get_text("text")
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if line.startswith("FG-PURPLLE-"):
                desc = line.strip()
                j = i + 1
                while j < len(lines) and not re.match(r"^\d{8}|\d+%", lines[j]):
                    if lines[j].strip() and not lines[j].startswith("3303"):
                        desc += " " + lines[j].strip()
                    j += 1

                file_items.append(
                    {
                        "File Name": uploaded_file.name,
                        "Invoice Number": invoice_no,
                        "PO Number": po_number,
                        "Destination": destination,
                        "Description of Goods": desc,
                    }
                )

    for idx, item in enumerate(file_items, start=1):
        item["SI No"] = idx

    return file_items


def process_excel(uploaded_file):
    """Read data directly from uploaded Excel file."""
    # Reads the first sheet into a DataFrame
    df_excel = pd.read_excel(uploaded_file)
    df_excel["File Name"] = uploaded_file.name
    return df_excel.to_dict(orient="records")


# 2. Main Logic to Handle Uploads
if uploaded_files:
    combined_data = []

    for file in uploaded_files:
        if file.name.endswith(".pdf"):
            pdf_records = process_pdf(file)
            combined_data.extend(pdf_records)
        elif file.name.endswith((".xlsx", ".xls")):
            excel_records = process_excel(file)
            combined_data.extend(excel_records)

    if combined_data:
        df_result = pd.DataFrame(combined_data)
        st.success(
            f"Successfully processed {len(uploaded_files)} file(s)![cite: 1, 2]"
        )
        st.dataframe(df_result)
    else:
        st.warning("No valid data found in uploaded files.")
