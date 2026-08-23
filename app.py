import io
import re
import logging
import pandas as pd
import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from pypdf import PdfReader
import pytesseract
from PIL import Image

st.set_page_config(page_title="Perfume Invoice Data Extractor", layout="wide")

# Configure logging
logging.basicConfig(level=logging.INFO)


def extract_text_robust(pdf_bytes):
    """
    Extracts text from PDF bytes using PyMuPDF, pypdf, pdfplumber, and OCR fallbacks.
    """
    text_content = []

    # --- STRATEGY 1: PyMuPDF ---
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page_num in range(len(doc)):
            page = doc[page_num]
            extracted = page.get_text("text")

            if not extracted.strip():
                extracted = _ocr_page(page)

            text_content.append(f"--- Page {page_num + 1} ---\n{extracted}")

        if any(t.strip() for t in text_content):
            return "\n".join(text_content)
    except Exception as e:
        logging.warning(f"PyMuPDF failed: {e}. Falling back to pypdf...")

    # --- STRATEGY 2: pypdf ---
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        text_content = []
        for i, page in enumerate(reader.pages):
            extracted = page.extract_text() or ""
            text_content.append(f"--- Page {i + 1} ---\n{extracted}")

        if any(t.strip() for t in text_content):
            return "\n".join(text_content)
    except Exception as e:
        logging.warning(f"pypdf failed: {e}. Falling back to pdfplumber...")

    # --- STRATEGY 3: pdfplumber ---
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            text_content = []
            for i, page in enumerate(pdf.pages):
                extracted = page.extract_text() or ""
                text_content.append(f"--- Page {i + 1} ---\n{extracted}")
            return "\n".join(text_content)
    except Exception as e:
        logging.error(f"All extraction backends failed: {e}")
        return ""


def _ocr_page(page):
    """Fallback OCR method for scanned/image pages using Tesseract."""
    try:
        pix = page.get_pixmap(dpi=150)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return pytesseract.image_to_string(img)
    except Exception as e:
        logging.error(f"OCR failure: {e}")
        return ""


def extract_perfume_invoice_data(pdf_bytes, file_name):
    """
    Parses perfume invoice PDF content to extract metadata and line items.
    """
    text = extract_text_robust(pdf_bytes)
    if not text.strip():
        return []

    # 1. Invoice Number Extraction
    inv_match = re.search(r'ADF/\d{4}-\d{2}/\d+', text)
    invoice_number = inv_match.group(0) if inv_match else ""

    # 2. PO Number Extraction
    po_match = re.search(r'Buyer[\'’s\s]*Order\s*No\.?\s*\n?\s*(\d+)', text, re.IGNORECASE)
    po_number = po_match.group(1) if po_match else ""

    # 3. Destination Extraction
    dest_match = re.search(r'Destination\s*\n\s*([A-Za-z]+)', text, re.IGNORECASE)
    destination = dest_match.group(1) if dest_match else ""

    # 4. Extract Line Items
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    fg_indices = [idx for idx, line in enumerate(lines) if 'FG-PURPLLE' in line]

    raw_descriptions = []
    for idx in fg_indices:
        line = lines[idx]
        desc = line
        j = idx + 1
        while j < len(lines):
            nxt = lines[j]
            if any(k in nxt for k in ['STAY', 'ML-', 'OUD', 'AMBER', 'BLOOM', 'SUGAR', 'AFTER', 'DUSK', 'SUNSET', 'TILL', 'DOWN']) or nxt.startswith('-') or 'X ' in nxt:
                if not re.match(r'^\d+\s+FG-PURPLLE', nxt) and not nxt.startswith('33030050'):
                    desc += " " + nxt
                    j += 1
                else:
                    break
            else:
                break

        clean_desc = re.sub(r'^\d+\s+', '', desc).strip()
        clean_desc = re.sub(r'\s+', ' ', clean_desc)
        raw_descriptions.append(clean_desc)

    unique_descriptions = []
    for desc in raw_descriptions:
        if desc not in unique_descriptions and len(desc) > 15:
            unique_descriptions.append(desc)

    if not unique_descriptions:
        fragments = [l for l in lines if 'FG-PURPLLE' in l or '-50ML-' in l or '-20ML-' in l]
        combined = []
        curr = ""
        for frag in fragments:
            if 'FG-PURPLLE' in frag:
                curr = re.sub(r'^\d+\s+', '', frag)
            elif curr:
                curr += " " + frag
                combined.append(re.sub(r'\s+', ' ', curr))
                curr = ""
        unique_descriptions = list(dict.fromkeys(combined))

    records = []
    for idx, desc in enumerate(unique_descriptions, start=1):
        pack_m = re.search(r'X\s*(\d+)', desc, re.IGNORECASE)
        pack_size = pack_m.group(1) if pack_m else "1"

        size_m = re.search(r'(\d+\s*ML)', desc, re.IGNORECASE)
        size = size_m.group(1).replace(" ", "") if size_m else ""

        frag_m = re.search(r'STAY[ -]?([A-Z\s]+?)(?=\s*X\s*\d+|$)', desc)
        fragrance = "STAY " + frag_m.group(1).strip() if frag_m else desc

        sku_code = desc.split()[0] if desc else ""

        records.append({
            "File Name": file_name,
            "Sl No": idx,
            "Invoice Number": invoice_number,
            "PO Number": po_number,
            "Destination": destination,
            "SKU Code": sku_code,
            "Fragrance Name": fragrance,
            "Size": size,
            "Pack Size": pack_size,
            "Full Description": desc
        })

    return records


# Streamlit UI
st.title("Perfume PDF Invoice Extractor")
st.write("Upload your PDF invoices below to extract items into structured table data.")

uploaded_files = st.file_uploader(
    "Choose PDF Invoices", 
    type=["pdf"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    with st.spinner("Processing PDF invoices..."):
        for uploaded_file in uploaded_files:
            bytes_data = uploaded_file.read()
            records = extract_perfume_invoice_data(bytes_data, uploaded_file.name)
            all_data.extend(records)

    if all_data:
        df = pd.DataFrame(all_data)
        st.success(f"Successfully processed {len(uploaded_files)} PDF file(s)!")
        
        # Display data grid
        st.dataframe(df, use_container_width=True)

        # Convert to Excel stream for download
        output_buffer = io.BytesIO()
        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Summary')
        excel_data = output_buffer.getvalue()

        st.download_button(
            label="Download Data as Excel (.xlsx)",
            data=excel_data,
            file_name="Extracted_Perfume_Invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Could not extract any matching invoice records from the uploaded PDF(s).")
