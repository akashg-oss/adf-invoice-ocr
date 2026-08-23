import io
import re
import logging
import pandas as pd
import streamlit as st
import fitz  # PyMuPDF
from pypdf import PdfReader
import pytesseract
from PIL import Image

st.set_page_config(page_title="Perfume Invoice & Master SKU Extractor", layout="wide")
logging.basicConfig(level=logging.INFO)


# --- MASTER SKU & EAN MATCHING ENGINE ---
def match_master_sku(extracted_text, size_val, master_df):
    """Matches extracted line item text and size against the Master Excel File."""
    if master_df is None or master_df.empty:
        return "", ""

    desc_clean = re.sub(r'[^A-Z0-9]', ' ', str(extracted_text).upper())
    size_clean = re.sub(r'[^A-Z0-9]', '', str(size_val).upper())

    best_match = ("", "")
    best_score = 0

    for _, row in master_df.iterrows():
        m_sku = str(row.get('SKU Code', '')).strip()
        m_ean = str(row.get('EAN', '')).strip()
        m_name = str(row.get('SKU Name', '')).upper()

        # Size check
        m_size = ""
        if "20ML" in m_name or "MINI" in m_name:
            m_size = "20ML"
        elif "50ML" in m_name:
            m_size = "50ML"
        elif "100ML" in m_name:
            m_size = "100ML"

        if size_clean and m_size and size_clean != m_size:
            continue

        # Exact SKU Code or EAN match in text
        if (m_sku and m_sku.upper() in desc_clean) or (m_ean and m_ean in desc_clean):
            return m_sku, m_ean

        # Extract keywords for fuzzy match
        words = set(re.findall(r'[A-Z]{3,}', m_name)) - {'FACES', 'CANADA', 'EAU', 'PARFUM', 'MINI'}
        score = sum(1 for w in words if w in desc_clean)

        if "DAWN" in words and ("DAWN" in desc_clean or "DOWN" in desc_clean):
            score += 1

        if score > best_score:
            best_score = score
            best_match = (m_sku, m_ean)

    if best_match[0] and best_score >= 1:
        return best_match

    return "", ""


# --- ACCURATE INVOICE PARSER ---
def extract_perfume_invoice_data(pdf_bytes, file_name, master_df):
    """Extracts Tax Invoice line items while filtering out e-Way Bill pages."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    records = []
    sl_no = 1

    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # 1. Extract words with bounding box coordinates to preserve spatial row layout
        words = page.get_text("words")
        if not words:
            continue

        # Group words vertically into physical lines
        lines = {}
        for w in words:
            y_group = round(w[1] / 3.0) * 3.0
            lines.setdefault(y_group, []).append(w)

        sorted_y = sorted(lines.keys())
        page_lines = []
        for y in sorted_y:
            line_words = sorted(lines[y], key=lambda x: x[0])
            line_str = " ".join([w[4] for w in line_words])
            page_lines.append(line_str)

        page_text = "\n".join(page_lines)

        # 2. FILTER OUT E-WAY BILL PAGES (Prevents duplicate item counts)
        if "e-Way Bill System" in page_text or "FORM GST EWB-01" in page_text:
            continue

        # Header Metadata Extraction
        inv_match = re.search(r'ADF/\d{4}-\d{2}/\d+', page_text)
        invoice_number = inv_match.group(0) if inv_match else ""

        po_match = re.search(r'Buyer[\'’s\s]*Order\s*No\.?\s*(\d+)', page_text, re.IGNORECASE)
        po_number = po_match.group(1) if po_match else ""

        dest_match = re.search(r'Destination\s*[:\n\s]*([A-Za-z]+)', page_text, re.IGNORECASE)
        destination = dest_match.group(1) if dest_match else ""

        # Extract item descriptions starting with FG-PURPLLE
        raw_items = page_text.split('\n')
        desc_chunks = []
        curr_chunk = ""

        for line in raw_items:
            if re.match(r'^\d+\s+FG-PURPLLE', line) or 'FG-PURPLLE' in line:
                if curr_chunk:
                    desc_chunks.append(curr_chunk)
                curr_chunk = line
            elif curr_chunk and any(k in line for k in ['STAY', 'MINI', '-20ML-', '-50ML-', '-100ML-', 'UNTIL', 'AFTER', 'TILL', 'X ']):
                curr_chunk += " " + line

        if curr_chunk:
            desc_chunks.append(curr_chunk)

        # Match table quantity and unit price: [Billed Qty] PCS [Unit Rate] PCS [Amount]
        row_pattern = re.compile(
            r'([\d,]+\.\d{4})\s*PCS\s+([\d,]+\.\d{2})\s*(?:PCS)?\s+([\d,]+\.\d{2})',
            re.IGNORECASE
        )
        matches = list(row_pattern.finditer(page_text))

        for idx, match in enumerate(matches):
            # Number of Units (Billed PCS)
            num_units_raw = match.group(1).replace(',', '')
            num_units = str(int(float(num_units_raw)))

            # Unit Price / Rate
            unit_price = match.group(2).replace(',', '')

            item_str = desc_chunks[idx] if idx < len(desc_chunks) else ""

            # Pack Size (e.g., X 48 or X 36)
            pack_m = re.search(r'X\s*(\d+)', item_str, re.IGNORECASE)
            pack_size = pack_m.group(1) if pack_m else "1"

            # Size (e.g., 20ML, 50ML, 100ML)
            size_m = re.search(r'(\d+\s*ML)', item_str, re.IGNORECASE)
            size = size_m.group(1).replace(" ", "").upper() if size_m else ""

            # Master SKU & EAN Lookup
            sku_code, ean_code = match_master_sku(item_str, size, master_df)

            records.append({
                "File Name": file_name,
                "Sl No": sl_no,
                "Invoice Number": invoice_number,
                "PO Number": po_number,
                "Destination": destination,
                "SKU Code": sku_code,
                "EAN Code": ean_code,
                "Size": size,
                "Pack Size": pack_size,
                "Unit Price": unit_price,
                "Number of Units": num_units
            })
            sl_no += 1

    return records


# --- STREAMLIT USER INTERFACE ---
st.title("Perfume Invoice Data Extractor")
st.write("Upload your **Master SKU Excel file (`EAN_SKU.xlsx`)** and **PDF Invoices** below.")

col1, col2 = st.columns([1, 1])

with col1:
    master_file = st.file_uploader(
        "1. Upload Master SKU Excel File",
        type=["xlsx", "xls"],
        key="master_file"
    )

with col2:
    uploaded_pdfs = st.file_uploader(
        "2. Upload Perfume Invoice PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_files"
    )

if uploaded_pdfs:
    master_df = None
    if master_file is not None:
        try:
            master_df = pd.read_excel(master_file)
            st.sidebar.success(f"Loaded Master File with {len(master_df)} SKUs.")
        except Exception as e:
            st.sidebar.error(f"Error reading Master Excel File: {e}")
    else:
        st.warning("No Master SKU File uploaded. 'SKU Code' and 'EAN Code' lookup will be empty.")

    all_data = []
    with st.spinner("Processing PDF invoices..."):
        for pdf_file in uploaded_pdfs:
            bytes_data = pdf_file.read()
            records = extract_perfume_invoice_data(bytes_data, pdf_file.name, master_df)
            all_data.extend(records)

    if all_data:
        df = pd.DataFrame(all_data)

        # Drop empty rows and sanitize
        df = df.dropna(how='all')

        columns_order = [
            "File Name", "Sl No", "Invoice Number", "PO Number", "Destination",
            "SKU Code", "EAN Code", "Size", "Pack Size", "Unit Price", "Number of Units"
        ]
        df = df.reindex(columns=columns_order)

        st.success(f"Successfully extracted {len(df)} item rows from {len(uploaded_pdfs)} invoice(s)!")

        # Display Data Grid
        st.dataframe(df, use_container_width=True)

        # Download CSV option
        csv_data = df.to_csv(index=False, lineterminator='\n').encode('utf-8')
        st.download_button(
            label="Download CSV (.csv)",
            data=csv_data,
            file_name="Extracted_Perfume_Invoices.csv",
            mime="text/csv"
        )

        # Download Excel option
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Summary')
        excel_data = output_buffer.getvalue()

        st.download_button(
            label="Download Excel (.xlsx)",
            data=excel_data,
            file_name="Extracted_Perfume_Invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Could not extract any matching invoice items from the uploaded PDF(s).")
