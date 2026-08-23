import pandas as pd
import pdfplumber
import re
import io
import streamlit as st

def normalize_text(text):
    """Normalize string values to ensure matching succeeds."""
    if pd.isna(text):
        return ""
    # Remove decimal representation for codes (e.g., 1234.0 -> 1234)
    text_str = str(text).strip()
    if text_str.endswith(".0"):
        text_str = text_str[:-2]
    # Remove non-alphanumeric characters for clean comparison
    return re.sub(r'[^A-Za-z0-9]', '', text_str).upper()

def process_invoices(master_excel_file, pdf_files):
    # 1. Load and prepare Master SKU dataset
    df_master = pd.read_excel(master_excel_file)
    
    # Standardize column headers to lowercase
    df_master.columns = [str(col).strip().lower() for col in df_master.columns]
    
    # Identify target matching columns (e.g., 'ean' or 'sku')
    match_col = next((col for col in ['ean', 'sku', 'barcode', 'item_code'] if col in df_master.columns), df_master.columns[0])
    
    # Store normalized master codes in a set for O(1) lookup
    master_codes = set(df_master[match_col].apply(normalize_text))
    
    extracted_records = []

    # 2. Iterate through all uploaded PDFs
    for pdf_file in pdf_files:
        with pdfplumber.open(pdf_file) as pdf:
            # Iterate through ALL pages instead of restricting to page 1
            for page_num, page in enumerate(pdf.pages, start=1):
                
                # Extract structured tables if available
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # Clean and check each cell for a master code match
                        row_cleaned = [str(cell).strip() if cell else "" for cell in row]
                        row_normalized = [normalize_text(cell) for cell in row_cleaned]
                        
                        if any(code in master_codes for code in row_normalized if code):
                            extracted_records.append({
                                "Source PDF": pdf_file.name,
                                "Page": page_num,
                                "Extracted Data": " | ".join(row_cleaned)
                            })
                
                # Fallback: Extract raw text line-by-line if table parsing misses items
                text = page.extract_text()
                if text:
                    for line in text.split("\n"):
                        tokens = [normalize_text(tok) for tok in line.split()]
                        if any(tok in master_codes for tok in tokens if tok):
                            extracted_records.append({
                                "Source PDF": pdf_file.name,
                                "Page": page_num,
                                "Extracted Data": line.strip()
                            })

    return pd.DataFrame(extracted_records).drop_duplicates()

# --- Streamlit UI Integration ---
st.title("Perfume Invoice Data Extractor")

excel_file = st.file_uploader("1. Upload Master SKU Excel File", type=["xlsx", "xls"])
pdf_files = st.file_uploader("2. Upload Perfume Invoice PDFs", type=["pdf"], accept_multiple_files=True)

if st.button("Process Invoices"):
    if excel_file and pdf_files:
        results_df = process_invoices(excel_file, pdf_files)
        
        if not results_df.empty:
            st.success(f"Successfully extracted {len(results_df)} line items!")
            st.dataframe(results_df)
        else:
            st.error("Could not extract any matching invoice items. Verify that EAN/SKU values in the Excel file match the PDF text.")
    else:
        st.warning("Please upload both the Master Excel file and PDF invoices.")
