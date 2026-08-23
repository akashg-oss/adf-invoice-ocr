import os
import re
import io
import logging
import pandas as pd
import fitz  # PyMuPDF
import pdfplumber
from pypdf import PdfReader
import pytesseract
from PIL import Image

# Configure logging to catch PDF repair warnings
logging.basicConfig(level=logging.INFO)


def extract_text_robust(pdf_path_or_bytes):
    """
    Extracts text from a PDF, resilient to structural corruption,
    encoding anomalies, missing font maps, and scanned/image-only pages.
    """
    text_content = []

    # --- STRATEGY 1: PyMuPDF (Fastest, repairs minor header/xref corruption) ---
    try:
        doc = fitz.open(stream=pdf_path_or_bytes, filetype="pdf") if isinstance(pdf_path_or_bytes, bytes) else fitz.open(pdf_path_or_bytes)
        for page_num in range(len(doc)):
            page = doc[page_num]
            extracted = page.get_text("text")

            # Fallback to OCR if page has no selectable text
            if not extracted.strip():
                extracted = _ocr_page(page)

            text_content.append(f"--- Page {page_num + 1} ---\n{extracted}")

        if any(t.strip() for t in text_content):
            return "\n".join(text_content)
    except Exception as e:
        logging.warning(f"PyMuPDF failed on file: {e}. Falling back to pypdf...")

    # --- STRATEGY 2: pypdf (Handles strict specification bugs & stream defects) ---
    try:
        reader = PdfReader(pdf_path_or_bytes if isinstance(pdf_path_or_bytes, str) else io.BytesIO(pdf_path_or_bytes))
        text_content = []
        for i, page in enumerate(reader.pages):
            extracted = page.extract_text() or ""
            text_content.append(f"--- Page {i + 1} ---\n{extracted}")

        if any(t.strip() for t in text_content):
            return "\n".join(text_content)
    except Exception as e:
        logging.warning(f"pypdf failed: {e}. Falling back to pdfplumber...")

    # --- STRATEGY 3: pdfplumber (Handles layout anomalies & complex encodings) ---
    try:
        with pdfplumber.open(pdf_path_or_bytes if isinstance(pdf_path_or_bytes, str) else io.BytesIO(pdf_path_or_bytes)) as pdf:
            text_content = []
            for i, page in enumerate(pdf.pages):
                extracted = page.extract_text() or ""
                text_content.append(f"--- Page {i + 1} ---\n{extracted}")
            return "\n".join(text_content)
    except Exception as e:
        logging.error(f"All extraction backends failed: {e}")
        raise RuntimeError("PDF is severely damaged or unreadable.") from e


def _ocr_page(page):
    """Fallback OCR method for scanned/image pages using Tesseract."""
    try:
        pix = page.get_pixmap(dpi=150)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return pytesseract.image_to_string(img)
    except Exception as e:
        logging.error(f"OCR failure on page: {e}")
        return ""


def extract_perfume_invoice_data(pdf_path):
    """
    Parses perfume invoice PDF content to extract metadata and line items.
    """
    text = extract_text_robust(pdf_path)
    if not text.strip():
        return []

    # 1. Invoice Number Extraction
    inv_match = re.search(r'ADF/\d{4}-\d{2}/\d+', text)
    invoice_number = inv_match.group(0) if inv_match else ""

    # 2. PO / Buyer Order Number Extraction
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
        # Stitch multi-line product descriptions
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

    # Deduplicate items in order while filtering incomplete fragments
    unique_descriptions = []
    for desc in raw_descriptions:
        if desc not in unique_descriptions and len(desc) > 15:
            unique_descriptions.append(desc)

    # Fallback logic for split-row invoice tables
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

    # Parse individual item columns
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


def batch_process_invoices(directory_path, output_excel="Perfume_Invoices_Extracted.xlsx"):
    """
    Processes all PDFs in a folder and saves the consolidated data into an Excel spreadsheet.
    """
    all_records = []
    pdf_files = [f for f in os.listdir(directory_path) if f.lower().endswith('.pdf')]

    logging.info(f"Found {len(pdf_files)} PDF file(s) to process.")

    for file_name in pdf_files:
        full_path = os.path.join(directory_path, file_name)
        logging.info(f"Processing invoice: {file_name}")
        records = extract_perfume_invoice_data(full_path)
        all_records.extend(records)

    if not all_records:
        logging.warning("No records extracted.")
        return None

    df = pd.DataFrame(all_records)
    df.to_excel(output_excel, index=False)
    logging.info(f"Successfully exported data to {output_excel}")
    return df


if __name__ == "__main__":
    # Specify folder directory containing your perfume PDFs
    process_directory = "." 
    result_df = batch_process_invoices(process_directory)
    if result_df is not None:
        print(result_df)
