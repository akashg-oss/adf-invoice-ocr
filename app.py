import io
import re
import fitz  # PyMuPDF
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Invoice Extractor & Master SKU Mapper", layout="wide")
st.title("Invoice Extractor & Master SKU Mapper")

# --- SIDEBAR: MASTER SKU FILE UPLOAD ---
st.sidebar.header("Master SKU Mapping")
master_sku_file = st.sidebar.file_uploader(
    "Upload EAN Master File",
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

        master_sku_df.columns = [str(c).strip() for c in master_sku_df.columns]
        st.sidebar.success(f"Loaded {len(master_sku_df)} Master SKU records!")
        with st.sidebar.expander("Preview Master SKU Data"):
            st.sidebar.dataframe(master_sku_df.head())
    except Exception as e:
        st.sidebar.error(f"Error loading Master SKU file: {e}")

# --- MAIN SECTION: INVOICE UPLOAD ---
st.header("Upload Invoices")
uploaded_files = st.file_uploader(
    "Upload PDF Invoices", type=["pdf"], accept_multiple_files=True
)


def extract_search_key(text):
    """Extracts text starting from volume (e.g., 100ML, 50ML, 20ML) up to product keywords.
    Example: '1 FG-PURPLLE-PER-100ML-AURA-LOVE STRUCK DELIGHT X 36' -> '100ML AURA LOVE STRUCK DELIGHT'
    """
    if not isinstance(text, str):
        return ""

    match = re.search(r"(\d+\s*ML.*)", text, re.IGNORECASE)
    if not match:
        return ""

    extracted = match.group(1)

    # Trim trailing pack sizes like 'X 36', 'X36', or '-X-36'
    extracted = re.sub(
        r"[\s\-_]+X[\s\-_]*\d+.*$", "", extracted, flags=re.IGNORECASE
    )

    # Clean punctuation and normalize spacing
    clean_key = re.sub(r"[\-_]+", " ", extracted).strip().upper()
    return " ".join(clean_key.split())


def extract_line_items_from_fitz(doc):
    """Extract line items robustly across pages using PyMuPDF block parsing."""
    items = []

    for page in doc:
        page_text = page.get_text("text")

        # Skip e-Way bill summary sections
        if "1. e-Way Bill Details" in page_text or "CEWB No." in page_text:
            continue

        # Extract blocks: b[4] contains full text content of a block box
        blocks = page.get_text("blocks")
        lines = []
        for b in blocks:
            for l in b[4].split("\n"):
                if l.strip():
                    lines.append(l.strip())

        i = 0
        while i < len(lines):
            line = lines[i]
            if "FG-PURPLLE" in line:
                combined_desc = line
                j = i + 1
                while j < len(lines):
                    next_l = lines[j]

                    # Stop if encountering HSN code, next serial number, or invoice totals
                    if (
                        re.match(r"^\d{8}\b", next_l)
                        or re.match(r"^\d+\s+FG-PURPLLE", next_l)
                        or any(
                            k in next_l
                            for k in [
                                "I.G.S.T.",
                                "Total",
                                "Amount Chargeable",
                                "ROUNDING OFF",
                            ]
                        )
                    ):
                        break

                    # Join multi-line split text
                    if combined_desc.endswith("-") or next_l.startswith("-"):
                        combined_desc += next_l
                    else:
                        combined_desc += " " + next_l

                    j += 1
                    if "X " in next_l or re.search(r"X\s*\d+$", next_l):
                        break

                items.append(combined_desc)
                i = j
            else:
                i += 1

    return items


def process_pdf_file(uploaded_file):
    """Processes header attributes and line item extractions per PDF."""
    file_bytes = uploaded_file.read()
    doc = fitz.open(stream=file_bytes, filetype="pdf")

    page1_text = doc[0].get_text("text")

    # Header Extraction
    inv_match = re.search(r"ADF/\d{4}-\d{2}/\d+", page1_text)
    if not inv_match:
        inv_match = re.search(
            r"Invoice\s*No\.?\s*[:\n\s]*([A-Z0-9/\-]+)", page1_text, re.IGNORECASE
        )
    invoice_no = inv_match.group(0).strip() if inv_match else ""

    po_match = re.search(r"\b(4\d{9})\b", page1_text)
    if not po_match:
        po_match = re.search(
            r"Buyer[’'s\s]*Order\s*No\.?[^\n]*\n\s*(\d{8,12})",
            page1_text,
            re.IGNORECASE,
        )
    po_number = po_match.group(1).strip() if po_match else ""

    dest_match = re.search(r"Destination\s*[:\n\s]*([A-Za-z]+)", page1_text)
    destination = dest_match.group(1).strip() if dest_match else ""

    raw_descriptions = extract_line_items_from_fitz(doc)
    doc.close()

    records = []
    for desc in raw_descriptions:
        records.append(
            {
                "File Name": uploaded_file.name,
                "Invoice Number": invoice_no,
                "PO Number": po_number,
                "Destination": destination,
                "Description of Goods": desc,
                "Extracted Search Key": extract_search_key(desc),
            }
        )

    # Deduplicate repeated tax summary rows
    seen = set()
    unique_records = []
    for r in records:
        key = (r["Invoice Number"], r["Description of Goods"])
        if key not in seen:
            seen.add(key)
            unique_records.append(r)

    for idx, item in enumerate(unique_records, start=1):
        item["SI No"] = idx

    return unique_records


# --- MAIN PIPELINE & EXCEL DOWNLOAD ---
if uploaded_files:
    combined_records = []

    for pdf_file in uploaded_files:
        pdf_data = process_pdf_file(pdf_file)
        combined_records.extend(pdf_data)

    if combined_records:
        df_result = pd.DataFrame(combined_records)

        # Merge with EAN Master File
        if master_sku_df is not None:
            sku_col = None
            for c in master_sku_df.columns:
                if any(k in c.lower() for k in ["sku name", "description", "item"]):
                    sku_col = c
                    break
            if not sku_col:
                sku_col = master_sku_df.columns[0]

            master_sku_df["Master_Search_Key"] = master_sku_df[sku_col].apply(
                extract_search_key
            )

            df_result = pd.merge(
                df_result,
                master_sku_df,
                left_on="Extracted Search Key",
                right_on="Master_Search_Key",
                how="left",
            )
            df_result.drop(
                columns=["Master_Search_Key"], inplace=True, errors="ignore"
            )

        st.success(f"Successfully processed {len(uploaded_files)} PDF file(s)!")
        st.dataframe(df_result, use_container_width=True)

        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="Extracted_Data")
        excel_buffer.seek(0)

        st.download_button(
            label="Download Result Excel",
            data=excel_buffer,
            file_name="Extracted_Invoice_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("No line items extracted from the uploaded PDF file(s).")
