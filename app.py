from difflib import get_close_matches
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
    "Upload EAN_SKU File",
    type=["xlsx", "xls", "csv"],
    help="Upload master file with columns like 'SKU Name', 'SKU Code', 'EAN'",
)

master_sku_df = None
if master_sku_file:
    try:
        if master_sku_file.name.endswith(".csv"):
            master_sku_df = pd.read_csv(master_sku_file)
        else:
            master_sku_df = pd.read_excel(master_sku_file)

        # Standardize column names to string
        master_sku_df.columns = [
            str(c).strip() for c in master_sku_df.columns
        ]
        st.sidebar.success(
            f"Loaded {len(master_sku_df)} Master SKU records!"
        )
        with st.sidebar.expander("Preview Master SKU Data"):
            st.sidebar.dataframe(master_sku_df.head())
    except Exception as e:
        st.sidebar.error(f"Error loading Master SKU file: {e}")

# --- MAIN SECTION: INVOICE UPLOAD ---
st.header("Upload Invoices")
uploaded_files = st.file_uploader(
    "Upload PDF Invoices", type=["pdf"], accept_multiple_files=True
)


def clean_text_for_matching(text):
    """Normalize string by upper-casing, removing hyphens, and stripping extra spaces."""
    if not isinstance(text, str):
        return ""
    # Convert to uppercase and replace hyphens/symbols with space
    text = re.sub(r"[-_:\s]+", " ", text.upper())
    return text.strip()


def extract_sku_name(raw_line):
    """Extract clean SKU name without quantity/amount noise."""
    # Capture string starting at FG-PURPLLE up to quantity markers (e.g. X 36 or prices)
    match = re.search(r"(FG-PURPLLE[^\d\n]+(?:\s*X\s*\d+)?)", raw_line)
    if match:
        return match.group(1).strip()

    # Fallback cleaning
    cleaned = re.sub(r"^\d+\s+", "", raw_line)
    cleaned = re.sub(
        r"[\d,]+\.\d{2,4}\s*(?:PCS)?|\bPCS\b|[\d,]+\.\d+|\bBOX\b",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    return " ".join(cleaned.split())


def process_pdf(uploaded_file):
    """Extract Invoice No, PO, Destination, and SKU Name using PyMuPDF (OCR/Text parser)."""
    file_bytes = uploaded_file.read()
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    # 1. OCR / Text Extraction for Header Fields from Page 1
    page1_text = doc[0].get_text("text")

    # Invoice Number
    inv_match = re.search(r"ADF/\d{4}-\d{2}/\d+", page1_text)
    if not inv_match:
        inv_match = re.search(
            r"Invoice\s*No\.?\s*[:\n\s]*([A-Z0-9/\-]+)",
            page1_text,
            re.IGNORECASE,
        )
    invoice_no = inv_match.group(0).strip() if inv_match else ""

    # PO Number
    po_match = re.search(r"\b(4\d{9})\b", page1_text)
    if not po_match:
        po_match = re.search(
            r"Buyer[’'s\s]*Order\s*No\.?[^\n]*\n\s*(\d{8,12})",
            page1_text,
            re.IGNORECASE,
        )
    po_number = po_match.group(1).strip() if po_match else ""

    # Destination
    dest_match = re.search(r"Destination\s*[:\n\s]*([A-Za-z]+)", page1_text)
    destination = dest_match.group(1).strip() if dest_match else ""

    # 2. Extract SKU Name per line item across pages
    file_items = []
    for page in doc:
        text = page.get_text("text")

        # Skip e-Way bill pages
        if "1. e-Way Bill Details" in text or "FORM GST EWB-01" in text:
            continue

        lines = text.split("\n")
        for i, line in enumerate(lines):
            if "FG-PURPLLE" in line:
                raw_sku = line.strip()
                j = i + 1
                while j < len(lines) and not re.match(r"^\d{8}|\d+%", lines[j]):
                    if (
                        lines[j].strip()
                        and not lines[j].startswith("3303")
                        and "ROUNDING OFF" not in lines[j]
                    ):
                        raw_sku += " " + lines[j].strip()
                    j += 1

                sku_name = extract_sku_name(raw_sku)

                file_items.append(
                    {
                        "File Name": uploaded_file.name,
                        "Invoice Number": invoice_no,
                        "PO Number": po_number,
                        "Destination": destination,
                        "Description of Goods": sku_name,
                    }
                )

    for idx, item in enumerate(file_items, start=1):
        item["SI No"] = idx

    return file_items


def xlookup_sku_data(sku_query, master_df):
    """XLOOKUP logic: Attempts exact lookup, then normalized lookup, then fuzzy matching."""
    if master_df is None or master_df.empty:
        return {}

    # Identify SKU Name/Description column in Master file
    master_sku_col = None
    for col in master_df.columns:
        if any(
            k in col.lower() for k in ["sku name", "description", "sku_name", "item"]
        ):
            master_sku_col = col
            break
    if not master_sku_col:
        master_sku_col = master_df.columns[0]

    # Pre-clean query
    query_clean = clean_text_for_matching(sku_query)

    # Build lookup dictionary with normalized keys
    master_keys = [
        clean_text_for_matching(str(val))
        for val in master_df[master_sku_col].values
    ]

    # 1. Direct / Exact Normalized Match
    if query_clean in master_keys:
        idx = master_keys.index(query_clean)
        return master_df.iloc[idx].to_dict()

    # 2. Fuzzy Match (XLOOKUP fallback)
    matches = get_close_matches(query_clean, master_keys, n=1, cutoff=0.6)
    if matches:
        best_match = matches[0]
        idx = master_keys.index(best_match)
        return master_df.iloc[idx].to_dict()

    return {}


# --- MAIN EXECUTION ---
if uploaded_files:
    combined_data = []

    for file in uploaded_files:
        pdf_records = process_pdf(file)
        combined_data.extend(pdf_records)

    if combined_data:
        df_result = pd.DataFrame(combined_data)

        # 2. XLOOKUP Extraction for EAN, SKU Code, SKU Name
        if master_sku_df is not None:
            extracted_records = []
            for _, row in df_result.iterrows():
                sku_query = row["Description of Goods"]
                master_match = xlookup_sku_data(sku_query, master_sku_df)

                # Merge master file fields (EAN, SKU Code, SKU Name)
                row_dict = row.to_dict()
                for col in master_sku_df.columns:
                    row_dict[col] = master_match.get(col, None)
                extracted_records.append(row_dict)

            df_result = pd.DataFrame(extracted_records)

        st.success(f"Processed {len(uploaded_files)} file(s) successfully!")
        st.dataframe(df_result, use_container_width=True)

        # Excel Export
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
