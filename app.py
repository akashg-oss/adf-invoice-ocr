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
    "GURGAON": "Gurgaon", "GURUGRAM": "Gurgaon", "HARYANA": "Gurgaon",
    "BANGALORE": "Bangalore", "BANGLORE": "Bangalore", "BENGALURU": "Bangalore",
    "KOLKATA": "Kolkata", "HOWRAH": "Howrah", "THANE": "Thane",
    "MUMBAI": "Mumbai", "DELHI": "Delhi", "NEW DELHI": "Delhi",
    "HYDERABAD": "Hyderabad", "CHENNAI": "Chennai", "PUNE": "Pune",
    "AHMEDABAD": "Ahmedabad", "NOIDA": "Noida", "BHIWANDI": "Bhiwandi"
}

def norm(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", str(x).upper())).strip()

def clean_sku_str(s):
    s = norm(s)
    # Strip common invoice noise words to expose core product variant & volume
    s = re.sub(r"\b(FG|PURPLLE|PER|STAY|33030050|PCS|BOX|X)\b", " ", s)
    return re.sub(r"\s+", " ", s).strip()

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
    return out[(out.EAN != "") & (out["SKU Code"] != "") & (out["SKU Name"] != "")].drop_duplicates()

def extract_text(data):
    # Strictly extract text from Page 1 (index 0) only
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        if not pdf.pages:
            return ""
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
    # 1. Invoice Number
    inv = ""
    m_inv = re.search(r"\b(ADF/\d{4}-\d{2}/\d+)\b", text, re.I)
    if not m_inv:
        m_inv = re.search(r"Invoice\s*(?:No|Number|#)?\.?\s*[:\-]?\s*([A-Z0-9\/\-_]{5,})", text, re.I)
    if m_inv and m_inv.group(1).upper() not in ["TAX", "INVOICE", "TAX INVOICE"]:
        inv = m_inv.group(1).strip()

    # 2. PO Number
    po = ""
    m_po = re.search(r"Buyer[’']?s\s+Order\s+No\.?\s*[:\-]?\s*([0-9A-Z\/\-_]{5,})", text, re.I)
    if not m_po:
        m_po = re.search(r"PO\s*(?:No|Number)?\.?\s*[:\-]?\s*([0-9A-Z\/\-_]{5,})", text, re.I)
    if m_po and m_po.group(1).upper() not in ["DATED", "DATE"]:
        po = m_po.group(1).strip()

    # 3. City / Destination Extraction
    city = ""
    m_dest = re.search(r"Destination\s*[:\-]?\s*([A-Za-z0-9\s\,\-]+)", text, re.I)
    if not m_dest:
        m_dest = re.search(r"Ship\s*To\s*[:\-]?\s*([A-Za-z0-9\s\,\-]+)", text, re.I)
    if not m_dest:
        m_dest = re.search(r"Consignee\s*\(Ship\s*to\)\s*[:\-]?\s*([A-Za-z0-9\s\,\-]+)", text, re.I)

    raw_dest = norm(m_dest.group(1)) if m_dest else norm(text)
    
    # Map raw destination string against known cities
    for k, v in CITY_MAP.items():
        if k in raw_dest:
            city = v
            break
            
    if not city and m_dest:
        clean_d = re.sub(r"\(SHIP TO\)|DATED|INVOICE", "", raw_dest).strip()
        if clean_d and len(clean_d) > 2:
            city = clean_d.title()

    # 4. Line Items Extraction (Page 1 Only)
    items = []
    lines = text.split("\n")
    
    for line in lines:
        line_clean = line.strip()
        if not line_clean:
            continue

        # Skip Box summary/total rows
        if re.search(r"\bBOX\b", line_clean, re.I) and not re.search(r"\bPCS\b", line_clean, re.I):
            continue

        # Item pattern: captures product name, PCS quantity, and unit price
        m = re.search(
            r"(?:33030050\s+)?(FG-[A-Z0-9\s\-\/\.\&]+?)(?:\s+X|\s+33030050)?\s+.*?([\d,]+(?:\.\d+)?)\s*PCS\s+([\d,]+\.\d+)", 
            line_clean, 
            re.I
        )
        if m:
            fg_desc = m.group(1).strip()
            fg_desc = re.sub(r"\s+X$", "", fg_desc, flags=re.I).strip()
            fg_desc = re.sub(r"^\d+\s+", "", fg_desc).strip()
            fg_desc = re.sub(r"^33030050\s+", "", fg_desc).strip()
            
            try:
                qty = int(round(float(m.group(2).replace(",", ""))))
                rate = float(m.group(3).replace(",", ""))
                items.append((fg_desc, qty, rate))
            except ValueError:
                continue

    return inv, po, city, items

def match(s, master):
    clean_s = clean_sku_str(s)
    if not clean_s:
        clean_s = norm(s)
        
    master_sku_names = master["SKU Name"].tolist()
    master_clean = [clean_sku_str(x) for x in master_sku_names]

    # 1. Direct exact clean match
    if clean_s in master_clean:
        idx = master_clean.index(clean_s)
        return master.iloc[idx], 100

    # 2. High-precision fuzzy token match on cleaned names
    res = process.extractOne(clean_s, master_clean, scorer=fuzz.token_set_ratio)
    if res and res[1] >= 60:
        return master.iloc[res[2]], res[1]

    # 3. Fallback fuzzy match on raw master names
    res_raw = process.extractOne(norm(s), [norm(x) for x in master_sku_names], scorer=fuzz.token_set_ratio)
    if res_raw and res_raw[1] >= 60:
        return master.iloc[res_raw[2]], res_raw[1]

    return None, 0

def make_excel(df):
    b = io.BytesIO()
    tmp = df.assign(_value=df["Quantity Dispatched (PCS)"] * df["Price of FG (₹/PCS)"])
    summ = tmp.groupby(["Invoice Number", "City / Destination", "PO Number"], as_index=False).agg(
        **{
            "Total Quantity Dispatched (PCS)": ("Quantity Dispatched (PCS)", "sum"),
            "Taxable FG Value (₹)": ("_value", "sum")
        }
    )
    with pd.ExcelWriter(b, engine="openpyxl") as w:
        df[OUTPUT_COLUMNS].to_excel(w, index=False, sheet_name="Invoice Details")
        summ.to_excel(w, index=False, sheet_name="Summary")
    return b.getvalue()

# Streamlit UI
st.set_page_config(page_title="ADF Invoice OCR → Standard Excel", layout="wide")
st.title("ADF Invoice OCR → Standard Excel")
st.caption("Upload EAN/SKU master and ADF invoices. Extracts line items, maps SKU/EAN, and detects destination city.")

mf = st.file_uploader("1. Upload EAN ↔ SKU Master Excel", type=["xlsx", "xls"])
pf = st.file_uploader("2. Upload ADF Invoice PDFs", type=["pdf"], accept_multiple_files=True)

if mf and pf and st.button("🚀 Process Invoices", type="primary"):
    try: 
        master = load_master(mf)
    except Exception as e: 
        st.error(str(e))
        st.stop()
        
    rows = []
    issues = []
    
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
                sku = ean = ""
                fg_name = s
            else:
                sku, ean, fg_name = m["SKU Code"], m["EAN"], m["SKU Name"]
                if score < 85: 
                    issues.append([f.name, f"Product match {score}%: {s} → {fg_name}"])
            rows.append([sku, ean, fg_name, inv, city, po, q, r])
            
    result = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    if result.empty: 
        st.error("No line items extracted.")
        st.stop()
        
    result["Quantity Dispatched (PCS)"] = pd.to_numeric(result["Quantity Dispatched (PCS)"]).astype("Int64")
    result["Price of FG (₹/PCS)"] = pd.to_numeric(result["Price of FG (₹/PCS)"])
    
    st.success(f"Processed {len(pf)} invoice(s), {len(result)} line items.")
    st.dataframe(result, use_container_width=True)
    st.download_button(
        "📥 Download Standardized Excel", 
        make_excel(result), 
        "ADF_Invoice_Dispatch_Details.xlsx", 
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    if issues: 
        st.warning("Some items need review.")
        st.dataframe(pd.DataFrame(issues, columns=["File", "Issue"]), use_container_width=True)
    else: 
        st.success("✅ All invoices mapped successfully.")
