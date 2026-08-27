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
    help="Upload your master list to auto-map EAN, SKU Code, SKU Name, etc.",
)

master_sku_df = None
if master_sku_file:
    try:
        if master_sku_file.name.endswith(".csv"):
            master_sku_df = pd.read_csv(master_sku_file)
        else:
            master_sku_df = pd.read_excel(master_sku_file)

        # Ensure all columns are converted to clean strings for merging
        for col in master_sku_df.columns:
            master_sku_df[col] = master_sku_df[col].astype(str).str.strip()

        st.sidebar.success(
            f"Loaded {len(master_sku_df)} Master SKU records!"
        )
        with st.sidebar.expander("Preview Master SKU Data"):
            st.sidebar.dataframe(master_sku_df.head())
    except Exception as e:
        st.sidebar.error(f"Error loading Master SKU file: {e}")

# --- MAIN SECTION: INVOICE / EXCEL FILE UPLOAD ---
st.header("Upload Invoices / Excel Files")
uploaded_files = st.file_uploader(
    "Upload PDF Invoices or Data Spreadsheets",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True,
)


def clean_description(text):
    """Strip out numbers, amounts, PCS, box counts, and serial numbers leaving pure SKU description."""
    # Remove item numbers at start (e.g., '1 ', '2 ')
    text = re.sub(r"^\d+\s+", "", text)
    # Remove amounts, rates, quantities (e.g., '1,40,833.44', 'PCS', '177.82')
    text = re.sub(
        r"[\d,]+\.\d{2,4}\s*(?:PCS)?|\bPCS\b|[\d,]+\.\d+|\bBOX\b",
        "",
        text,
        flags=re.IGNORECASE,
    )
    # Remove extra whitespace
    return " ".join(text.split())


def process_pdf(uploaded_file):
    """Extract structured data from uploaded PDF invoice."""
    file_bytes = uploaded_file.read()
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    # 1. Robust Invoice Number Extraction from Page 1
    page1_text = doc[0].get_text("text")

    inv_match = re.search(r"ADF/\d{4}-\d{2}/\d+", page1_text)
    if not inv_match:
        inv_match = re.search(
            r"Invoice\s*No\.?\s*[:\n\s]*([A-Z0-9/\-]+)",
            page1_text,
            re.IGNORECASE,
        )
    invoice_no = inv_match.group(0).strip() if inv_match else ""

    # 2. Extract 10-digit PO number
    po_match = re.search(
        r"Buyer[’'s\s]*Order\s*No\.?[^\n]*\n\s*(\d{8,12})",
        page1_text,
        re.IGNORECASE,
    )
    if not po_match:
        po_match = re.search(r"\b(4\d{9})\b", page1_text)
    po_number = po_match.group(1).strip() if po_match else ""

    # 3. Extract Destination
    dest_match = re.search(r"Destination\s*[:\n\s]*([A-Za-z]+)", page1_text)
    destination = dest_match.group(1).strip() if dest_match else "Bangalore"

    # 4. Extract Line Items across all pages
    file_items = []
    for page in doc:
        text = page.get_text("text")

        # Skip standalone e-Way bill pages
        if "1. e-Way Bill Details" in text or "FORM GST EWB-01" in text:
            continue

        lines = text.split("\n")
        for i, line in enumerate(lines):
            if "FG-PURPLLE" in line:
                desc = line.strip()
                j = i + 1
                while j < len(lines) and not re.match(r"^\d{8}|\d+%", lines[j]):
                    if (
                        lines[j].strip()
                        and not lines[j].startswith("3303")
                        and not "ROUNDING OFF" in lines[j]
                    ):
                        desc += " " + lines[j].strip()
                    j += 1

                cleaned_desc = clean_description(desc)

                file_items.append(
                    {
                        "File Name": uploaded_file.name,
                        "Invoice Number": invoice_no,
                        "PO Number": po_number,
                        "Destination": destination,
                        "Description of Goods": cleaned_desc,
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

        # Automatic merge with Master SKU file
        if master_sku_df is not None and not master_sku_df.empty:
            inv_col = "Description of Goods"
            sku_col = master_sku_df.columns[0]  # Match against first column

            # Normalize values for exact matching
            df_result["join_key"] = (
                df_result[inv_col].astype(str).str.upper().str.strip()
            )
            master_sku_temp = master_sku_df.copy()
            master_sku_temp["join_key"] = (
                master_sku_temp[sku_col].astype(str).str.upper().str.strip()
            )

            # Left Join master SKU columns onto extracted result
            df_result = df_result.merge(
                master_sku_temp, on="join_key", how="left"
            )
            df_result.drop(columns=["join_key"], inplace=True)

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
