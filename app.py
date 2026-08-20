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

def norm(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+", " ", re.sub(r"[^A-Z0-9]+", " ", str(x).upper())).strip()

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

def first(patterns, text):
    for p in patterns:
        m = re.search(p, text, re.I)
        if m: return m.group(1).strip()
    return ""

def parse(text):
    # 1. Invoice Number (exclusively matches invoice pattern, ignoring static headers like "TAX")
    inv = first([
        r"\b(ADF/\d{4}-\d{2}/\d+)\b",
        r"Invoice\s*(?:No|Number|#)?\.?\s*[:\-]?\s*([A-Z0-9\/\-_]{5,})"
    ], text)
    if inv.upper() in ["TAX", "INVOICE", "TAX INVOICE"]:
        inv = ""

    # 2. PO Number (extracts numeric/alphanumeric PO, ignoring labels like "Dated")
    po = first([
        r"Buyer[’']?s\s+Order\s+No\.?\s*[:\-]?\s*([0-9A-Z\/\-_]{5,})",
        r"PO\s*(?:No|Number)?\.?\s*[:\-]?\s*([0-9A-Z\/\-_]{5,})"
    ], text)
    if po.upper() in ["DATED", "DATE"]:
        po = ""

    # 3. City / Destination (ignores static label "(Ship to)")
    dest = first([
        r"Destination\s*[:\-]?\s*([A-Za-z\s]{3,})",
        r"Ship\s*To\s*[:\-]?\s*([A-Za-z\s]{3,})"
    ], text)
    if "(SHIP TO)" in dest.upper() or dest.upper() in ["SHIP TO", "(SHIP TO)", "DATED"]:
        dest = ""

    d = norm(dest)
    if any(k in d for k in ["GURGAON", "GURUGRAM", "HARYANA"]): dest = "Gurgaon"
    elif any(k in d for k in ["BANGLORE", "BANGALORE"]): dest = "Bangalore"
    elif d in ("KOLKATA", "HOWRAH", "THANE", "MUMBAI", "DELHI"): dest = d.title()

    # 4. Extract Line Items (Page 1 Only)
    items = []
    lines = text.split("\n")
    
    for line in lines:
        line_clean = line.strip()
        if not line_clean:
            continue

        # Skip summary BOX lines (e.g. "1 FG-PURPLLE-PER-20ML 33030050 88.00 BOX")
        if re.search(r"\bBOX\b", line_clean, re.I) and not re.search(r"\bPCS\b", line_clean, re.I):
            continue
        if re.match(r"^\d+\s+FG-", line_clean, re.I) and "BOX" in line_clean.upper():
            continue

        # Primary pattern: Captures detailed item rows containing PCS quantities
        m = re.search(
            r"(?:33030050\s+)?(FG-[A-Z0-9\s\-\/\.\&]+?)(?:\s+X|\s+33030050)?\s+.*?([\d,]+(?:\.\d+)?)\s*PCS\s+([\d,]+\.\d+)", 
            line_clean, 
            re.I
        )
        if m:
            fg_desc = m.group(1).strip()
            
            # Cleanup leading numbers, HSN codes, and trailing 'X'
            fg_desc = re.sub(r"\s+X$", "", fg_desc, flags=re.I).strip()
            fg_desc = re.sub(r"^\d+\s+", "", fg_desc).strip()
            fg_desc = re.sub(r"^33030050\s+", "", fg_desc).strip()
            
            try:
                qty = int(round(float(m.group(2).replace(",", ""))))
                rate = float(m.group(3).replace(",", ""))
                items.append((fg_desc, qty, rate))
            except ValueError:
                continue

    # Secondary pattern fallback if strict FG prefix is absent on page 1
    if not items:
        for line in lines:
            line_clean = line.strip()
            if not line_clean or ("BOX" in line_clean.upper() and "PCS" not in line_clean.upper()):
                continue
            m = re.search(
                r"([A-Z0-9\s\-\/\.\&]{5,})\s+([\d,]+(?:\.\d+)?)\s*PCS\s+([\d,]+\.\d+)", 
                line_clean, 
                re.I
            )
            if m:
                fg_desc = m.group(1).strip()
                if any(h in fg_desc.upper() for h in ["DESCRIPTION", "TOTAL", "SUBTOTAL", "INVOICE", "TAXABLE", "AMOUNT"]):
                    continue
                try:
                    qty = int(round(float(m.group(2).replace(",", ""))))
                    rate = float(m.group(3).replace(",", ""))
                    items.append((fg_desc, qty, rate))
                except ValueError:
                    continue

    return inv, po, dest, items

def match(s, master):
    clean_s = norm(s)
    norm_master_names = [norm(x) for x in master["SKU Name"]]

    if clean_s in norm_master_names:
        idx = norm_master_names.index(clean_s)
        return master.iloc[idx], 100

    res = process.extractOne(clean_s, norm_master_names, scorer=fuzz.token_set_ratio)
    if res and res[1] >= 70:
        return master.iloc[res[2]], res[1]

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
st.caption("Upload your EAN / SKU master and ADF invoices. Reads line items strictly from Page 1.")

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
                sku = ean = name = ""
            else:
                sku, ean, name = m["SKU Code"], m["EAN"], m["SKU Name"]
                if score < 90: 
                    issues.append([f.name, f"Product match {score}%: {s} → {name}"])
            rows.append([sku, ean, name or s, inv, city, po, q, r])
            
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
