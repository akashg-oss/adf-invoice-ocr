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


# --- STRATEGY: Robust PDF Text Extraction ---
def extract_text_robust(pdf_bytes):
    """Extracts line-sorted text from PDF bytes using PyMuPDF, pypdf, and OCR fallbacks."""
    text_content = []

    # 1. PyMuPDF with vertical position sorting
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page_num in range(len(doc)):
            page = doc[page_num]
            # Sort text blocks top-to-bottom, left-to-right
            blocks = page.get_text("blocks", flags=fitz.TEXT_PRESERVE_WHITESPACE)
            blocks.sort(key=lambda b: (b[1], b[0]))
            
            page_text = "\n".join([b[4] for b in blocks if b[4].strip()])
            if not page_text.strip():
                page_text = _ocr_page(page)

            text_content.append(f"--- Page {page_num + 1} ---\n{page_text}")

        if any(t.strip() for t in text_content):
            return "\n".join(text_content)
    except Exception as e:
        logging.warning(f"PyMuPDF failed: {e}. Falling back to pypdf...")

    # 2. pypdf Fallback
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text_content = []
        for i, page in enumerate(reader.pages):
            extracted = page.extract_text() or ""
            text_content.append(f"--- Page {i + 1} ---\n{extracted}")

        if any(t.strip() for t in text_content):
            return "\n".join(text_content)
    except Exception as e:
        logging.warning(f"pypdf failed: {e}")

    return ""


def _ocr_page(page):
    """Fallback OCR method for scanned or image-only pages using Tesseract."""
    try:
        pix = page.get_pixmap(dpi=150)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return pytesseract.image_to_string(img)
    except Exception as e:
        logging.error(f"OCR failure: {e}")
        return ""


# --- MASTER SKU & EAN MATCHING ENGINE ---
def match_master_sku(extracted_text, size_val, master_df):
    """
    Matches extracted PDF line text and size against the Master Excel File.
    Returns tuple: (Master SKU Code, Master EAN Code)
    """
    if master_df is None or master_df.empty:
        return "", ""

    desc_clean = re.sub(r'[^A-Z0-9]', ' ', str(extracted_text).upper())
    size_clean = re.sub(r'[^A-Z0-9]', '', str(size_val).upper())  # e.g., "20ML"

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

        # Handle specific common typos (e.g., TILL DOWN vs TILL DAWN)
        if "DAWN" in words and ("DAWN" in desc_clean or "DOWN" in desc_clean):
            score += 1

        if score > best_score:
            best_score = score
            best_match = (m_sku, m_ean)

    if best_match[0] and best_score >= 1:
        return best_match

    return "", ""


# --- INVOICE PARSER ---
def extract_perfume_invoice_data(pdf_bytes, file_name, master_df):
    """Parses perfume invoice PDF and maps extracted line items to Master Excel data."""
    text = extract_text_robust(pdf_bytes)
    if not text.strip():
        return []

    # Filter out e-Way Bill pages to prevent duplicate/invalid parsing
    pages = text.split("--- Page ")
    tax_invoice_pages = [p for p in pages if "TAX INVOICE" in p and "e-Way Bill" not in p]
    invoice_text = "\n".join(tax_invoice_pages) if tax_invoice_pages else text

    # Header Metadata Extraction
    inv_match = re.search(r'ADF/\d{4}-\d{2}/\d+', invoice_text)
    invoice_number = inv_match.group(0) if inv_match else ""

    po_match = re.search(r'Buyer[\'’s\s]*Order\s*No\.?\s*(\d+)', invoice_text, re.IGNORECASE)
    po_number = po_match.group(1) if po_match else ""

    dest_match = re.search(r'Destination\s*[:\n\s]*([A-Za-z]+)', invoice_text, re.IGNORECASE)
    destination = dest_match.group(1) if dest_match else ""

    # Stitch multi-line line items
    raw_lines = [l.strip() for l in invoice_text.split('\n') if l.strip()]
    structured_items = []
    current_item = ""

    for line in raw_lines:
        # Detect start of new item row
        if re.match(r'^\d+\s+FG-PURPLLE', line) or 'FG-PURPLLE' in line:
            if current_item:
                structured_items.append(current_item)
            current_item = line
        elif current_item and any(k in line for k in ['STAY', '-20ML-', '-50ML-', '-100ML-', 'UNTIL', 'AFTER', 'TILL', 'X ']):
            current_item += " " + line
        elif current_item and not any(k in line for k in ['Total', 'Amount', 'IGST', 'Subtotal', 'Taxable']):
            if re.search(r'X\s*\d+', line):
                current_item += " " + line
                structured_items.append(current_item)
                current_item = ""

    if current_item:
        structured_items.append(current_item)

    records = []
    sl_no = 1

    for item_str in structured_items:
        # Pack Size (e.g., X 48 or X 36)
        pack_m = re.search(r'X\s*(\d+)', item_str, re.IGNORECASE)
        pack_size = pack_m.group(1) if pack_m else "1"

        # Size (e.g., 20ML, 50ML, 100ML)
        size_m = re.search(r'(\d+\s*ML)', item_str, re.IGNORECASE)
        size = size_m.group(1).replace(" ", "").upper() if size_m else ""

        # Unit Price / Rate
        prices = re.findall(r'\b\d+\.\d{2}\b', item_str)
        unit_price = prices[-2] if len(prices) >= 2 else (prices[0] if prices else "")

        # Quantity / Number of Units
        pcs_match = re.search(r'([\d,]+\.\d{4})\s*PCS', item_str, re.IGNORECASE)
        if pcs_match:
            num_units = str(int(float(pcs_match.group(1).replace(',', ''))))
        else:
            qty_match = re.search(r'(\d+)\s*(?:Pcs|Nos|Units|Qty|PCS|NOS)', item_str, re.IGNORECASE)
            num_units = qty_match.group(1) if qty_match else ""

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

        # Drop any empty/NaN rows and sanitize
        df = df.dropna(how='all')

        columns_order = [
            "File Name", "Sl No", "Invoice Number", "PO Number", "Destination",
            "SKU Code", "EAN Code", "Size", "Pack Size", "Unit Price", "Number of Units"
        ]
        df = df.reindex(columns=columns_order)

        st.success(f"Successfully extracted {len(df)} item rows from {len(uploaded_pdfs)} invoice(s)!")

        # Display Data Grid
        st.dataframe(df, use_container_width=True)

        # Download CSV option (without empty lines)
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
