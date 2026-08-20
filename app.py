import io, re
import pandas as pd
import streamlit as st
import pdfplumber
from rapidfuzz import fuzz, process

OUTPUT_COLUMNS = ["SKU Code","EAN","Name of FG","Invoice Number","City / Destination","PO Number","Quantity Dispatched (PCS)","Price of FG (₹/PCS)"]

def norm(x):
    if pd.isna(x): return ""
    return re.sub(r"\s+"," ",re.sub(r"[^A-Z0-9]+"," ",str(x).upper())).strip()

def load_master(f):
    df=pd.read_excel(f,dtype=str)
    df.columns=[str(c).strip() for c in df.columns]
    n={norm(c):c for c in df.columns}
    def col(names):
        for x in names:
            if norm(x) in n: return n[norm(x)]
        for c in df.columns:
            if any(norm(x) in norm(c) or norm(c) in norm(x) for x in names): return c
    ec,sc,nc=col(["EAN","EAN Code","EAN No"]),col(["SKU Code","SKU"]),col(["SKU Name","Name","Product Name","Name of FG","FG Name","Description"])
    if not all([ec,sc,nc]): raise ValueError("Master Excel must contain EAN, SKU Code and SKU Name columns.")
    out=df[[ec,sc,nc]].copy(); out.columns=["EAN","SKU Code","SKU Name"]
    for c in out.columns: out[c]=out[c].fillna("").astype(str).str.replace(r"\.0$","",regex=True).str.strip()
    return out[(out.EAN!="")&(out["SKU Code"]!="")&(out["SKU Name"]!="")].drop_duplicates()

def extract_text(data):
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text="\n".join(p.extract_text() or "" for p in pdf.pages)
    if len(re.sub(r"\s+","",text))<250:
        try:
            import fitz,pytesseract
            from PIL import Image
            doc=fitz.open(stream=data,filetype="pdf")
            text+="\n"+"\n".join(pytesseract.image_to_string(Image.open(io.BytesIO(p.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False).tobytes("png")))) for p in doc)
        except Exception: pass
    return text

def first(patterns,text):
    for p in patterns:
        m=re.search(p,text,re.I)
        if m:return m.group(1).strip()
    return ""

def parse(text):
    inv=first([r"\b(ADF/\d{4}-\d{2}/\d+)\b"],text)
    po=first([r"Buyer[’']?s\s+Order\s+No\.?\s*\n?\s*([0-9]{6,})",r"PO\s*(?:No|Number)?\.?\s*[:\-]?\s*([0-9]{6,})"],text)
    dest=first([r"Destination\s*\n?\s*([^\n]+)"],text)
    d=norm(dest)
    if d=="HARYANA" and re.search(r"GURGAON|GURUGRAM",text,re.I): dest="Gurgaon"
    elif d in ("BANGLORE","BANGALORE"): dest="Bangalore"
    elif d in ("KOLKATA","HOWRAH","THANE"): dest=d.title()
    pat=re.compile(r"(FG-PURPLLE-PER-\d+ML-[A-Z0-9][A-Z0-9 \-]*?)\s+X\s+\d+\s+33030050\s+[\d,]+\.\d+\s+BOX\s+([\d,]+\.\d+)\s+PCS\s+([\d,]+\.\d+)\s+PCS\s+([\d,]+\.\d+)",re.I)
    items=[(re.sub(r"\s+"," ",m.group(1)).strip(),int(round(float(m.group(2).replace(",","")))),float(m.group(3).replace(",",""))) for m in pat.finditer(text)]
    return inv,po,dest,items

phrases=["OUD TILL DOWN","AMBER UNTIL SUNSET","BLOOM AFTER DARK","SUGAR AFTER DUSK","VANILLA PAST MIDNIGHT","WHITE MOON LIGHT","SPARKLING ECSTASY","AURA SOFT SERENITY","AURA SILENT FIRE","AURA ROMANTIC DAYDREAMS","AURA LOVESTRUCK DELIGHT","AURA WHIMSICAL WILD"]
def key(s):
    s=norm(s); vol=re.search(r"\b(20ML|50ML|100ML)\b",s)
    p=next((x for x in phrases if x in s),"")
    return norm(p+" "+(vol.group(1) if vol else "")) if p else s

def match(s,master):
    keys=[key(x) for x in master["SKU Name"]]; k=key(s)
    exact=[i for i,x in enumerate(keys) if x==k]
    if exact:return master.iloc[exact[0]],100
    r=process.extractOne(k,keys,scorer=fuzz.token_set_ratio)
    if r and r[1]>=80:return master.iloc[r[2]],r[1]
    return None,0

def make_excel(df):
    b=io.BytesIO(); tmp=df.assign(_value=df["Quantity Dispatched (PCS)"]*df["Price of FG (₹/PCS)"])
    summ=tmp.groupby(["Invoice Number","City / Destination","PO Number"],as_index=False).agg(**{"Total Quantity Dispatched (PCS)":("Quantity Dispatched (PCS)","sum"),"Taxable FG Value (₹)":("_value","sum")})
    with pd.ExcelWriter(b,engine="openpyxl") as w:
        df[OUTPUT_COLUMNS].to_excel(w,index=False,sheet_name="Invoice Details"); summ.to_excel(w,index=False,sheet_name="Summary")
    return b.getvalue()

st.set_page_config(page_title="ADF Invoice OCR → Standard Excel",layout="wide")
st.title("ADF Invoice OCR → Standard Excel")
st.caption("Upload your EAN / SKU master and ADF invoices. The app creates one fixed Excel format.")
st.info("Master format: EAN | SKU Code | SKU Name")
mf=st.file_uploader("1. Upload EAN ↔ SKU Master Excel",type=["xlsx","xls"])
pf=st.file_uploader("2. Upload ADF Invoice PDFs",type=["pdf"],accept_multiple_files=True)

if mf and pf and st.button("🚀 Process Invoices",type="primary"):
    try: master=load_master(mf)
    except Exception as e: st.error(str(e)); st.stop()
    rows=[]; issues=[]
    for f in pf:
        inv,po,city,items=parse(extract_text(f.read()))
        if not inv: issues.append([f.name,"Invoice number not detected"])
        if not po: issues.append([f.name,"PO number not detected"])
        if not city: issues.append([f.name,"City / Destination not detected"])
        if not items: issues.append([f.name,"No line items detected"])
        for s,q,r in items:
            m,score=match(s,master)
            if m is None:
                issues.append([f.name,"Product mapping not found: "+s]); sku=ean=name=""
            else:
                sku,ean,name=m["SKU Code"],m["EAN"],m["SKU Name"]
                if score<95: issues.append([f.name,f"Product match {score}%: {s} → {name}"])
            rows.append([sku,ean,name,inv,city,po,q,r])
    result=pd.DataFrame(rows,columns=OUTPUT_COLUMNS)
    if result.empty: st.error("No line items extracted."); st.stop()
    result["Quantity Dispatched (PCS)"]=pd.to_numeric(result["Quantity Dispatched (PCS)"]).astype("Int64")
    result["Price of FG (₹/PCS)"]=pd.to_numeric(result["Price of FG (₹/PCS)"])
    st.success(f"Processed {len(pf)} invoice(s), {len(result)} line items.")
    st.dataframe(result,use_container_width=True)
    st.download_button("📥 Download Standardized Excel",make_excel(result),"ADF_Invoice_Dispatch_Details.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if issues: st.warning("Some items need review."); st.dataframe(pd.DataFrame(issues,columns=["File","Issue"]),use_container_width=True)
    else: st.success("✅ All invoices mapped successfully.")
