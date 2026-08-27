import io
import re
import fitz  # PyMuPDF
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer
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


def extract_search_key(text):
    """Extracts text starting from volume (e.g., 100ML, 50ML, 20ML) up to product keywords.
    Example: '1 FG-PURPLLE-PER-100ML-AURA-LOVE STRUCK DELIGHT X 36' -> '100ML AURA LOVE STRUCK DELIGHT'
    """
    if not isinstance(text, str):
        return ""

    # Search for volume start (e.g., 100ML, 50 ML, 20ML)
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


def extract_pdf_line_items(file_bytes):
    """Parses wrapped multi-line descriptions from invoice tables using vertical coordinates."""
    file_items = []
    
    # Process PDF layout coordinates to solve column text wrapping
    for page_layout in extract_pages(io.BytesIO(file_bytes)):
        text_elements = []
        for element in page_layout:
            if isinstance(element, LTTextContainer):
                for text_line in element:
                    t = text_line.get_text().strip()
                    if t:
                        # Store bounding boxes: (x0, y0, text)
                        text_elements.append((text_line.bbox[0], text_line.bbox[1], t))

        # Filter elements inside table description column coordinates (x0 between 60 & 250)
        col_items = [item for item in text_elements if 60 <= item[0] <= 250 and item[1] > 100]
        col_items.sort(key=lambda x: x[1], reverse=True)  # Sort top to bottom (y-descending)

        # Merge wrapped lines belonging to the same product row
        i = 0
        while i < len(col_items):
            text = col_items[i][2]
            if "FG-PURPLLE" in text or re.search(r"^\d+\s+FG-PURPLLE", text):
                combined_desc = text
                j = i + 1
                while j < len(col_items):
                    next_y = col_items[j][1]
                    next_text = col_items[j][2]

                    # Stop if reaching next item index or HSN code boundary
                    if re.match(r"^\d+\s+FG-PURPLLE", next_text) or next_text.startswith("3303"):
                        break

                    combined_desc += " " + next_text
                    j += 1
                    if "X " in next_text or re.search(r"X\s*\d+", next_text):
                        break

                file_items.append(combined_desc)
                i = j
            else:
                i += 1

    return file_items


def process_pdf_file(uploaded_file):
    """Extracts metadata and line items from an individual PDF."""
    file_bytes = uploaded_file.read()

    # Extract Page 1 Header information using PyMuPDF
    doc = fitz.open(stream=file_bytes, filetype="pdf")
    page1_text = doc[0].get_text("text")

    # Invoice Number
    inv_match = re.search(r"ADF/\d{4}-\d{2}/\d+", page1_text)
    if not inv_match:
        inv_match = re.search(r"Invoice\s*No\.?\s*[:\n\s]*([A-Z0-9/\-]+)", page1_text, re.IGNORECASE)
    invoice_no = inv_match.group(0).strip() if inv_match else ""

    # PO Number
    po_match = re.search(r"\b(4\d{9})\b", page1_text)
    if not po_match:
        po_match = re.search(r"Buyer[’'s\s]*Order\s*No\.?[^\n]*\n\s*(\d{8,12})", page1_text, re.IGNORECASE)
    po_number = po_match.group(1).strip() if po_match else ""

    # Destination
    dest_match = re.search(r"Destination\s*[:\n\s]*([A-Za-z]+)", page1_text)
    destination = dest_match.group(1).strip() if dest_match else ""

    doc.close()

    # Extract Descriptions
    raw_descriptions = extract_pdf_line_items(file_bytes)

    records = []
    for desc in raw_descriptions:
        records.append({
            "File Name": uploaded_file.name,
            "Invoice Number": invoice_no,
            "PO Number": po_number,
            "Destination": destination,
            "Description of Goods": desc,
            "Extracted Search Key": extract_search_key(desc)
        })

    # Deduplicate repeated items from PDF tax summaries
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


# --- PROCESSING & EXCEL DOWNLOAD ---
if uploaded_files:
    combined_records = []

    for pdf_file in uploaded_files:
        pdf_data = process_pdf_file(pdf_file)
        combined_records.extend(pdf_data)

    if combined_records:
        df_result = pd.DataFrame(combined_records)

        # Merge with EAN Master File if uploaded
        if master_sku_df is not None:
            # Identify SKU Name column in Master
            sku_col = None
            for c in master_sku_df.columns:
                if any(k in c.lower() for k in ["sku name", "description", "item"]):
                    sku_col = c
                    break
            if not sku_col:
                sku_col = master_sku_df.columns[0]

            master_sku_df["Master_Search_Key"] = master_sku_df[sku_col].apply(extract_search_key)

            # Left join extracted search key with master file search key
            df_result = pd.merge(
                df_result,
                master_sku_df,
                left_on="Extracted Search Key",
                right_on="Master_Search_Key",
                how="left"
            )

            # Clean up redundant key columns
            df_result.drop(columns=["Master_Search_Key"], inplace=True, errors="ignore")

        st.success(f"Successfully processed {len(uploaded_files)} PDF file(s)!")
        st.dataframe(df_result, use_container_width=True)

        # Generate Excel Stream for Download
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="Extracted_Invoice_Data")
        excel_buffer.seek(0)

        st.download_button(
            label="Download Result Excel",
            data=excel_buffer,
            file_name="Extracted_Invoice_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("No line items extracted from the uploaded PDF file(s).")
