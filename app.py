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
    text = re.sub(r"[-_:\s]+", " ", text.upper())
    return text.strip()


def extract_sku_name(raw_text):
    """Clean multi-line raw item text into a clean SKU description."""
    if not raw_text:
        return ""

    # Remove item numbers at start (e.g., '1 ', '2 ') or HSN codes ('33030050 ')
    cleaned = re.sub(r"^\d+\s+", "", raw_text)
    cleaned = re.sub(r"^\d{8}\s+", "", cleaned)

    # Clean out prices, quantities, PCS, rates, percentages
    cleaned = re.sub(
        r"[\d,]+\.\d{2,4}\s*(?:PCS)?|\bPCS\b|[\d,]+\.\d+|\bBOX\b",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    # Flatten wrapped hyphens (e.g. 'AURA -SPARKLING' -> 'AURA-SPARKLING')
    cleaned = re.sub(r"(\b[A-Z0-9]+)\s*-\s*([A-Z0-9]+\b)", r"\1-\2", cleaned)

    # Capture complete product name starting at FG-PURPLLE
    match = re.search(
        r"(FG-PURPLLE-[A-Z0-9\-\s&\.\/]+?(?:\s*X\s*\d+)?)",
        cleaned,
        re.IGNORECASE,
    )
    if match:
        return " ".join(match.group(1).split())

    return " ".join(cleaned.split())


def process_pdf(uploaded_file):
    """Extract Header fields and line item descriptions across PDF pages."""
    file_bytes = uploaded_file.read()
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    page1_text = doc[0].get_text("text")

    # 1. Invoice Number
    inv_match = re.search(r"ADF/\d{4}-\d{2}/\d+", page1_text)
    if not inv_match:
        inv_match = re.search(
            r"Invoice\s*No\.?\s*[:\n\s]*([A-Z0-9/\-]+)",
            page1_text,
            re.IGNORECASE,
        )
    invoice_no = inv_match.group(0).strip() if inv_match else ""

    # 2. PO Number
    po_match = re.search(r"\b(4\d{9})\b", page1_text)
    if not po_match:
        po_match = re.search(
            r"Buyer[’'s\s]*Order\s*No\.?[^\n]*\n\s*(\d{8,12})",
            page1_text,
            re.IGNORECASE,
        )
    po_number = po_match.group(1).strip() if po_match else ""

    # 3. Destination
    dest_match = re.search(r"Destination\s*[:\n\s]*([A-Za-z]+)", page1_text)
    destination = dest_match.group(1).strip() if dest_match else ""

    # 4. Extract Line Items across pages using text blocks
    file_items = []

    for page in doc:
        text = page.get_text("text")

        # Skip e-Way bill summary pages
        if "1. e-Way Bill Details" in text or "FORM GST EWB-01" in text:
            continue

        # Use blocks to preserve line wrapping context within table cells
        blocks = page.get_text("blocks")
        full_page_str = ""
        for b in blocks:
            full_page_str += b[4] + "\n"

        lines = [l.strip() for l in full_page_str.split("\n") if l.strip()]

        i = 0
        while i < len(lines):
            line = lines[i]

            if "FG-PURPLLE" in line:
                accumulated_block = line
                j = i + 1

                # Read ahead and append next wrapped lines if present
                while j < len(lines):
                    next_l = lines[j]

                    # Stop if we hit HSN Code (33030050), Next Serial Item, or Totals block
                    if (
                        re.match(r"^\d{8}\b", next_l)
                        or re.match(r"^\d+\s+FG-PURPLLE", next_l)
                        or any(
                            stop_kw in next_l
                            for stop_kw in [
                                "I.G.S.T.",
                                "Total",
                                "Amount Chargeable",
                                "ROUNDING OFF",
                            ]
                        )
                    ):
                        break

                    # Append wrapped line text
                    if accumulated_block.endswith("-") or next_l.startswith(
                        "-"
                    ):
                        accumulated_block += next_l
                    else:
                        accumulated_block += " " + next_l

                    j += 1

                    # Stop when pack count indicator is found (e.g. 'X 36')
                    if re.search(r"X\s*\d+$", next_l, re.IGNORECASE):
                        break

                cleaned_sku = extract_sku_name(accumulated_block)

                if cleaned_sku and not cleaned_sku.startswith("3303"):
                    file_items.append(
                        {
                            "File Name": uploaded_file.name,
                            "Invoice Number": invoice_no,
                            "PO Number": po_number,
                            "Destination": destination,
                            "Description of Goods": cleaned_sku,
                        }
                    )
                i = j
            else:
                i += 1

    # Remove duplicates if same item was captured twice from tax summary tables
    unique_items = []
    seen = set()
    for item in file_items:
        key = (item["Invoice Number"], item["Description of Goods"])
        if key not in seen:
            seen.add(key)
            unique_items.append(item)

    for idx, item in enumerate(unique_items, start=1):
        item["SI No"] = idx

    return unique_items


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

    query_clean = clean_text_for_matching(sku_query)

    master_keys = [
        clean_text_for_matching(str(val))
        for val in master_df[master_sku_col].values
    ]

    # 1. Direct Exact Normalized Match
    if query_clean in master_keys:
        idx = master_keys.index(query_clean)
        return master_df.iloc[idx].to_dict()

    # 2. Fuzzy Match (XLOOKUP fallback)
    matches = get_close_matches(query_clean, master_keys, n=1, cutoff=0.45)
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

        # XLOOKUP extraction for EAN, SKU Code, SKU Name from Master SKU File
        if master_sku_df is not None:
            extracted_records = []
            for _, row in df_result.iterrows():
                sku_query = row["Description of Goods"]
                master_match = xlookup_sku_data(sku_query, master_sku_df)

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
