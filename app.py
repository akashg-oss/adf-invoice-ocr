import io
import re
import pandas as pd
import streamlit as st
import pdfplumber
from rapidfuzz import fuzz, process

OUTPUT_COLUMNS = [
    "SKU Code",
    "EAN",
    "Name of FG",
    "Invoice Number",
    "City / Destination",
    "PO Number",
    "Quantity Dispatched (PCS)",
    "Price of FG (₹/PCS)"
]

CITY_MAP = {
    "BANGLORE": "Bangalore", "BANGALORE": "Bangalore", "BENGALURU": "Bangalore",
    "GURGAON": "Gurgaon", "GURUGRAM": "Gurgaon", "HARYANA": "Gurgaon",
    "KOLKATA": "Kolkata", "HOWRAH": "Howrah", "THANE": "Thane",
    "MUMBAI": "Mumbai", "DELHI": "Delhi", "NEW DELHI": "Delhi",
    "HYDERABAD": "Hyderabad", "CHENNAI": "Chennai", "PUNE": "Pune",
    "AHMEDABAD": "Ahmedabad", "NOIDA": "Noida", "BHIWANDI": "Bhiwandi"
}

def norm(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", str(x).upper())).strip()

def extract_variant_tokens(s):
    # Extracts distinct variant words (e.g., OUD, AMBER, BLOOM, SUGAR)
    s = norm(s)
    ignore = {"FG", "PURPLLE", "PER", "20ML", "STAY", "33030050", "PCS", "BOX", "X", "48", "TILL", "AFTER", "UNTIL"}
    return {w for w in s.split() if w not in ignore}

def load_master(f):
    df = pd.read_excel(f, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    n = {norm(c): c for c in df.columns}
    
    def col(names):
        for x in names:
            if norm(x) in n: return n[norm(x)]
        for c in df.columns:
            if any(norm(x) in norm(c) or norm(c) in norm(x) for x in names): return c

    ec = col(["EAN", "EAN Code", "EAN No"])
    sc = col(["SKU Code", "SKU"])
    nc = col(["SKU Name", "Name", "Product Name", "Name of FG", "FG Name", "Description"])
    
    if not all([ec, sc, nc]): 
        raise ValueError("Master Excel must contain EAN, SKU Code, and SKU Name columns.")
        
    out = df[[ec, sc, nc]].copy()
    out.columns = ["EAN", "SKU Code", "SKU Name"]
    for c in out.columns: 
        out[c] = out[c].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    return out[(out.EAN != "") & (out["SKU Code"] != "") & (out["SKU Name"] != "")].drop_duplicates().reset_index(drop=True)

def extract_text(data):
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        if not pdf.pages: return ""
        text = pdf.pages[0].extract_text() or ""
        
    if len(re.sub(r"\s+", "", text)) < 100:
        try:
            import fitz, pytesseract
            from PIL import Image
            doc = fitz.open(stream=data, filetype="pdf")
            if len(doc) > 0:
                text = pytesseract.image_to_string(
                    Image.open(io.BytesIO(doc[0].get_pixmap(matrix=fitz.Matrix(2,2), alpha=False).tobytes("png")))
                )
        except Exception: 
            pass
    return text

def parse(text):
    inv = ""
    m_inv = re.search(r"\b(ADF/\d{4}\s*-\s*\d{2}/\d+)\b", text, re.I)
    if m_inv:
        inv = re.sub(r"\s+", "", m_inv.group(1))
    else:
        m_inv = re.search(r"Invoice\s*(?:No|Number|#)?\.?\s*[:\-]?\s*([A-Z0-9\/\-_\s]{5,20})", text, re.I)
        if m_inv and m_inv.group(1).strip().upper() not in ["TAX", "INVOICE", "TAX INVOICE"]:
            inv = re.sub(r"\s+", "", m_inv.group(1))

    po = ""
    m_po = re.search(r"Buyer[’']?s\s+Order\s+No\.?\s*[:\-]?\s*([0-9A-Z\/\-_]{5,})", text, re.I)
    if not m_po:
        m_po = re.search(r"PO\s*(?:No|Number)?\.?\s*[:\-]?\s*([0-9A-Z\/\-_]{5,})", text, re.I)
    if m_po and m_po.group(1).upper() not in ["DATED", "DATE"]:
        po = m_po.group(1).strip()

    city = ""
    m_dest = re.search(r"Destination\s*\n?\s*([A-Za-z0-9\s\,\-]+)", text, re.I)
    raw_dest = norm(m_dest.group(1)) if m_dest else norm(text)
    for k, v in CITY_MAP.items():
        if k in raw_dest:
            city = v
            break

    items = []
    # Extract line item block using Sl No anchors
    pat = re.compile(
        r"(?:(?<=\n)|\A)\s*\d+\s+(FG-[\s\S]+?)\s+(?:33030050|\d{8})\s+[\d,]+\.\d+\s*BOX\s+([\d,]+(?:\.\d+)?)\s*PCS\s+([\d,]+\.\d+)\s*PCS",
        re.I
    )
    
    for m in pat.finditer(text):
        raw_desc = m.group(1)
        clean_desc = re.sub(r"\s+", " ", raw_desc).strip()
        clean_desc = re.sub(r"\s+X\s+\d+$", "", clean_desc, flags=re.I).strip()
        try:
            qty = int(round(float(m.group(2).replace(",", ""))))
            rate = float(m.group(3).replace(",", ""))
            items.append((clean_desc, qty, rate))
        except ValueError:
            continue

    return inv, po, city, items

def match(s, master):
    inv_norm = norm(s)
    inv_tokens = extract_variant_tokens(s)
    master_names = [norm(x) for x in master["SKU Name"]]

    # 1. Direct exact normalized match
    if inv_norm in master_names:
        idx = master_names.index(inv_norm)
        return master.iloc[idx], 100

    # 2. Strict variant token overlap match (prevents tie-breaks on index 0)
    best_idx = None
    best_score = -1

    for idx, m_name in enumerate(master_names):
        m_tokens = extract_variant_tokens(m_name)
        # Check if key variant terms (e.g. OUD, AMBER, BLOOM, SUGAR) overlap
        token_overlap = len(inv_tokens.intersection(m_tokens))
        
        # Combine token_sort_ratio with token overlap weight
        sort_score = fuzz.token_sort_ratio(inv_norm, m_name)
        composite_score = sort_score + (token_overlap * 20)

        if composite_score > best_score:
            best_score = composite_score
            best_idx = idx

    if best_idx is not None and best_score >= 40:
        return master.iloc[best_idx], min(100, best_score)

    return None, 0

# Streamlit Interface
st.set_page_config(page_title="ADF Invoice OCR → Standard Excel", layout="wide")
st.title("ADF Invoice OCR → Standard Excel")

mf = st.file_uploader("1. Upload EAN ↔ SKU Master Excel", type=["xlsx", "xls"])
pf = st.file_uploader("2. Upload ADF Invoice PDFs", type=["pdf"], accept_multiple_files=True)

if mf and pf and st.button("🚀 Process Invoices", type="primary"):
    try: 
        master = load_master(mf)
    except Exception as e: 
        st.error(str(e))
        st.stop()
        
    rows, issues = [], []
    
    for f in pf:
        inv, po, city, items = parse(extract_text(f.read()))
        
        if not inv: issues.append([f.name, "Invoice number not detected"])
        if not po: issues.append([f.name, "PO number not detected"])
        if not city: issues.append([f.name, "City / Destination not detected"])
        if not items: issues.append([f.name, "No line items detected on Page 1"])
        
        for s, q, r in items:
            m, score = match(s, master)
            if m is None:
                issues.append([f.name, f"Product mapping not found: {s}"])
                sku, ean, fg_name = "", "", s
            else:
                sku, ean, fg_name = m["SKU Code"], m["EAN"], m["SKU Name"]
            rows.append([sku, ean, fg_name, inv, city, po, q, r])
            
    result = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    if result.empty: 
        st.error("No line items extracted.")
        st.stop()
        
    result["Quantity Dispatched (PCS)"] = pd.to_numeric(result["Quantity Dispatched (PCS)"]).astype("Int64")
    result["Price of FG (₹/PCS)"] = pd.to_numeric(result["Price of FG (₹/PCS)"])
    
    st.success(f"Processed {len(pf)} invoice(s), {len(result)} line items extracted.")
    st.dataframe(result, use_container_width=True)
    
    if issues: 
        st.warning("Mapping Warnings:")
        st.dataframe(pd.DataFrame(issues, columns=["File", "Issue"]), use_container_width=True)
