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


def extract_full_sku_from_block(block_text):
    """Clean and extract full product description without table noise."""
    # Strip item numbers/HSN numbers at start
    cleaned = re.sub(r"^\d+\s+", "", block_text)
    cleaned = re.sub(r"^\d{8}\s+", "", cleaned)

    # Remove price, rate, quantity noise
    cleaned = re.sub(
        r"[\d,]+\.\d{2,4}\s*(?:PCS)?|\bPCS\b|[\d,]+\.\d+|\bBOX\b",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    # Capture complete SKU name up to pack indicator (e.g. X 36 / X 48)
    match = re.search(
        r"(FG-PURPLLE[^\n]+?(?:\s*X\s*\d+)?)", cleaned, re.IGNORECASE
    )
    if match:
        return " ".join(match.group(1).split())

    return " ".join(cleaned.split())


def process_pdf(uploaded_file):
    """Extract Header fields and assemble full multi-line SKU text using word coordinates."""
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

    # 4. Extract Line Items across pages using Coordinate-Based Word Grouping
    file_items = []

    for page in doc:
        text = page.get_text("text")

        # Skip e-Way bill pages
        if "1. e-Way Bill Details" in text or "FORM GST EWB-01" in text:
            continue

        # Extract words: (x0, y0, x1, y1, word, block_no, line_no, word_no)
        words = page.get_text("words")
        if not words:
            continue

        # Group words by vertical line position (y0 rounded within 3 points)
        lines_dict = {}
        for w in words:
            y_key = round(w[1] / 3) * 3
            lines_dict.setdefault(y_key, []).append(w)

        sorted_y_keys = sorted(lines_dict.keys())

        # Combine grouped words into clean horizontal lines
        page_lines = []
        for y_key in sorted_y_keys:
            line_words = sorted(lines_dict[y_key], key=lambda x: x[0])
            line_str = " ".join([w[4] for w in line_words])
            page_lines.append(line_str)

        # Iterate over line-grouped text
        idx = 0
        while idx < len(page_lines):
            line = page_lines[idx]

            if "FG-PURPLLE" in line:
                full_sku_line = line
                next_idx = idx + 1

                # Continue stitching lines until hitting totals or next item
                while next_idx < len(page_lines):
                    next_l = page_lines[next_idx]
                    if (
                        re.match(r"^\d+\s+FG-PURPLLE", next_l)
                        or re.match(r"^\d{8}\b", next_l)
                        or any(
                            k in next_l
                            for k in [
                                "I.G.S.T.",
                                "Total",
                                "Amount Chargeable",
                                "OUTPUT",
                                "ROUNDING OFF",
                            ]
                        )
                    ):
                        break

                    full_sku_line += " " + next_l
                    next_idx += 1

                    # Stop stitching when hitting end pack marker (e.g. X 36, X 48)
                    if re.search(r"\bX\s*\d+\b", next_l, re.IGNORECASE):
                        break

                cleaned_sku = extract_full_sku_from_block(full_sku_line)

                # Ignore HSN/summary rows
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
                idx = next_idx
            else:
                idx += 1

    for item_idx, item in enumerate(file_items, start=1):
        item["SI No"] = item_idx

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

    query_clean = clean_text_for_matching(sku_query)

    master_keys = [
        clean_text_for_matching(str(val))
        for val in master_df[master_sku_col].values
    ]

    # 1. Exact Match
    if query_clean in master_keys:
        match_idx = master_keys.index(query_clean)
        return master_df.iloc[match_idx].to_dict()

    # 2. Fuzzy Match Fallback
    matches = get_close_matches(query_clean, master_keys, n=1, cutoff=0.45)
    if matches:
        best_match = matches[0]
        match_idx = master_keys.index(best_match)
        return master_df.iloc[match_idx].to_dict()

    return {}


# --- MAIN EXECUTION ---
if uploaded_files:
    combined_data = []

    for file in uploaded_files:
        pdf_records = process_pdf(file)
        combined_data.extend(pdf_records)

    if combined_data:
        df_result = pd.DataFrame(combined_data)

        # Perform XLOOKUP matching against Master SKU file
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
