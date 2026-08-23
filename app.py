import io
import re
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Invoice Data & SKU Extractor", layout="wide")
st.title("Invoice Extractor & Master SKU Mapper")

# --- SIDEBAR: MASTER SKU FILE UPLOAD ---
st.sidebar.header("Master SKU Mapping")
master_sku_file = st.sidebar.file_uploader(
    "Upload Master SKU File",
    type=["xlsx", "xls", "csv"],
    help="Upload your master list to map descriptions or codes to Master SKUs.",
)

master_sku_df = None
if master_sku_file:
    try:
        if master_sku_file.name.endswith(".csv"):
            master_sku_df = pd.read_csv(master_sku_file)
        else:
            master_sku_df = pd.read_excel(master_sku_file)
        st.sidebar.success(
            f"Loaded {len(master_sku_df)} Master SKU records!"
        )
        with st.sidebar.expander("Preview Master SKU Data"):
            st.dataframe(master_sku_df.head())
    except Exception as e:
        st.sidebar.error(f"Error loading Master SKU file: {e}")

# --- MAIN SECTION: INVOICE / EXCEL FILE UPLOAD ---
st.header("Upload Invoices / Excel Files")
uploaded_files = st.file_uploader(
    "Upload PDF Invoices or Data Spreadsheets",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True,
)


def process_pdf(uploaded_file):
    """Extract structured data from uploaded PDF invoice."""
    file_bytes = uploaded_file.read()
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    # Extract Header Fields from Page 1
    page1_text = doc[0].get_text("text")

    inv_match = re.search(r"Invoice No\.\s*([A-Z0-9/-]+)", page1_text)
    invoice_no = inv_match.group(1).strip() if inv_match else None

    # Capture 10-digit PO number starting with 4 or general buyer order digits
    po_match = re.search(
        r"Buyer[’'s\s]*Order\s*No\.?[^\n]*\n\s*(\d{8,12})",
        page1_text,
        re.IGNORECASE,
    )
    if not po_match:
        po_match = re.search(r"\b(4\d{9})\b", page1_text)
    po_number = (
        po_match.group(1).strip()
        if po_match
        else (
            re.search(
                r"Buyer'?s Order No\.\s*(\d+)", page1_text, re.IGNORECASE
            ).group(1)
            if re.search(
                r"Buyer'?s Order No\.\s*(\d+)", page1_text, re.IGNORECASE
            )
            else None
        )
    )

    dest_match = re.search(r"Destination\s*([A-Za-z]+)", page1_text)
    destination = dest_match.group(1).strip() if dest_match else "Bangalore"

    # Extract Line Items across all pages (excluding e-Way bill pages)
    file_items = []
    for page in doc:
        text = page.get_text("text")

        # Skip standalone e-Way bill page
        if "1. e-Way Bill Details" in text or "FORM GST EWB-01" in text:
            continue

        lines = text.split("\n")
        for i, line in enumerate(lines):
            if line.startswith("FG-PURPLLE-") or "FG-PURPLLE" in line:
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
    df_excel = pd.read_excel(uploaded_file)
    df_excel["File Name"] = uploaded_file.name
    return df_excel.to_dict(orient="records")


# --- DATA PROCESSING & DISPLAY ---
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

        # Merge with Master SKU mapping if uploaded
        if master_sku_df is not None:
            st.subheader("Map Master SKU")
            col1, col2 = st.columns(2)
            with col1:
                inv_col = st.selectbox(
                    "Select Invoice Column to Match",
                    options=df_result.columns,
                    index=(
                        list(df_result.columns).index("Description of Goods")
                        if "Description of Goods" in df_result.columns
                        else 0
                    ),
                )
            with col2:
                sku_col = st.selectbox(
                    "Select Master SKU File Column to Match On",
                    options=master_sku_df.columns,
                )

            # Perform Left Join to keep all extracted rows
            df_result = df_result.merge(
                master_sku_df,
                left_on=inv_col,
                right_on=sku_col,
                how="left",
            )

        st.success(f"Processed {len(uploaded_files)} file(s) successfully!")
        st.dataframe(df_result, use_container_width=True)

        # Download Result as Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="Extracted_Data")
        buffer.seek(0)

        st.download_button(
            label="Download Result Excel",
            data=buffer,
            file_name="Extracted_Invoice_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("No line items extracted from the uploaded file(s).")
