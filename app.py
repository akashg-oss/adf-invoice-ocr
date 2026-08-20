import re
import difflib
import pandas as pd
import streamlit as st
import pdfplumber

st.set_page_config(page_title="Invoice SKU & EAN Mapper", layout="wide")

# Function to clean and tokenize product strings
def clean_tokens(text):
    s = str(text).upper()
    s = re.sub(r"[^A-Z0-9]", " ", s)
    stop_words = {
        "FG", "PURPLLE", "PER", "STAY", "EAU", "DE", "PARFUM", "MINI", 
        "FACES", "CANADA", "33030050", "PCS", "BOX", "X", "20ML", "50ML", "100ML"
    }
    return [t for t in s.split() if t not in stop_words]

# Extract pack size identifier
def extract_size(text):
    m = re.search(r"(\d+\s*ML)", str(text), re.IGNORECASE)
    return m.group(1).upper().replace(" ", "") if m else ""

# Fuzzy SKU Matcher against EAN_SKU.xlsx
def match_sku(raw_item_desc, master_df):
    if master_df is None or master_df.empty:
        return "", "", raw_item_desc, 0.0

    vol = extract_size(raw_item_desc)
    subset = master_df.copy()

    # Pre-filter master catalog by volumetric pack size
    if vol == "20ML":
        f = subset['SKU Name'].str.contains("20ml|mini", case=False, na=False)
        if f.any(): subset = subset[f]
    elif vol == "50ML":
        f = subset['SKU Name'].str.contains("50ml", case=False, na=False) & \
            ~subset['SKU Name'].str.contains("20ml|mini", case=False, na=False)
        if f.any(): subset = subset[f]
    elif vol == "100ML":
        f = subset['SKU Name'].str.contains("100ml", case=False, na=False)
        if f.any(): subset = subset[f]

    inv_tokens = clean_tokens(raw_item_desc)
    inv_str = " ".join(inv_tokens)
    
    best_row = None
    best_score = -1.0

    for idx, row in subset.iterrows():
        m_tokens = clean_tokens(row['SKU Name'])
        m_str = " ".join(m_tokens)
        
        # Sequence matching & token overlap score
        seq_ratio = difflib.SequenceMatcher(None, inv_str, m_str).ratio()
        token_scores = [
            max([difflib.SequenceMatcher(None, it, mt).ratio() for mt in m_tokens] or [0])
            for it in inv_tokens
        ]
        avg_token_score = sum(token_scores) / len(token_scores) if token_scores else 0
        total_score = (seq_ratio * 0.4) + (avg_token_score * 0.6)

        if total_score > best_score:
            best_score = total_score
            best_row = row

    if best_row is not None and best_score >= 0.35:
        return best_row['SKU Code'], str(best_row['EAN']), best_row['SKU Name'], round(best_score, 2)

    return "", "", raw_item_desc, 0.0

# Extract multi-line item blocks from invoice text
def parse_invoice_text(text, invoice_num, master_df):
    items = []
    # Pattern to capture multi-line product descriptions before HSN code
    line_pattern = re.compile(
        r"(?:(?<=\n)|\A)\s*(\d+)\s+(FG-[\s\S]+?)\s+(?:33030050|\d{8})\s+[\d,]+\.\d+\s*BOX\s+([\d,]+(?:\.\d+)?)\s*PCS\s+([\d,]+\.\d+)\s*PCS\s+([\d,]+\.\d+)",
        re.IGNORECASE
    )

    for m in line_pattern.finditer(text):
        sl_no = m.group(1)
        raw_desc = m.group(2)
        qty = float(m.group(3).replace(",", ""))
        rate = float(m.group(4).replace(",", ""))
        amount = float(m.group(5).replace(",", ""))

        # Standardize whitespace and trim trail pack quantity
        clean_desc = re.sub(r"\s+", " ", raw_desc).strip()
        clean_desc = re.sub(r"\s+X\s+\d+$", "", clean_desc, flags=re.IGNORECASE).strip()

        sku_code, ean, sku_name, match_conf = match_sku(clean_desc, master_df)

        items.append({
            "Sl No": sl_no,
            "SKU Code": sku_code,
            "EAN": ean,
            "Name of FG": sku_name,
            "Invoice Raw Desc": clean_desc,
            "Invoice Number": invoice_num,
            "Quantity Dispatched (PCS)": qty,
            "Price of FG (₹/PCS)": rate,
            "Amount (₹)": amount,
            "Match Score": match_conf
        })
    return items

# Streamlit User Interface
st.title("📦 Invoice SKU & EAN Automated Parser")

col1, col2 = st.columns(2)
with col1:
    master_file = st.file_uploader("Upload EAN_SKU.xlsx Catalog", type=["xlsx", "xls"])
with col2:
    pdf_files = st.file_uploader("Upload Invoice PDFs", type=["pdf"], accept_multiple_files=True)

master_df = None
if master_file:
    master_df = pd.read_excel(master_file)
    st.success(f"Loaded Master Catalog: {len(master_df)} SKUs")

if pdf_files and master_df is not None:
    all_extracted = []
    for pdf in pdf_files:
        inv_num_match = re.search(r"(\d{4}\b|[A-Z0-9/-]{10,})", pdf.name)
        invoice_num = pdf.name.replace(".pdf", "")

        with pdfplumber.open(pdf) as pdf_doc:
            full_text = "\n".join([page.extract_text() or "" for page in pdf_doc.pages])
            
            # Extract header invoice number if present
            inv_header = re.search(r"Invoice\s*No\.?\s*[:\-]?\s*([A-Z0-9/-]+)", full_text, re.I)
            if inv_header:
                invoice_num = inv_header.group(1)

            parsed_items = parse_invoice_text(full_text, invoice_num, master_df)
            all_extracted.extend(parsed_items)

    if all_extracted:
        df_results = pd.DataFrame(all_extracted)
        st.success(f"Processed {len(pdf_files)} invoice(s), {len(df_results)} line items extracted.")
        
        # Display extracted table with auto-matched SKU & EAN
        st.dataframe(df_results, use_container_width=True)

        # Export Options
        csv_data = df_results.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Download Extracted Data (CSV)", csv_data, "extracted_invoices.csv", "text/csv")
