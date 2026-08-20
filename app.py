import io, re
import pandas as pd
import streamlit as st
import pdfplumber

OUTPUT_COLUMNS = ["SKU Code","EAN","Name of FG","Invoice Number","City / Destination","PO Number","Quantity Dispatched (PCS)","Price of FG (₹/PCS)"]

def norm(s): return re.sub(r"\s+"," ",str(s or "")).strip()

def load_mapping(uploaded):
    df=pd.read_excel(uploaded)
    cols={str(c).strip().lower():c for c in df.columns}
    ean_col=next((cols[k] for k in cols if "ean" in k),None)
    name_col=next((cols[k] for k in cols if "sku name" in k or k=="name"),None)
    if not ean_col or not name_col: raise ValueError("Mapping Excel must contain EAN and SKU Name columns.")
    out=df[[ean_col,name_col]].copy(); out.columns=["EAN","Name of FG"]
    out["EAN"]=out["EAN"].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
    return out

def extract_text(data):
    text=""
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for p in pdf.pages: text += "\n"+(p.extract_text() or "")
    if len(re.sub(r"\s+","",text))<200:
        try:
            import fitz, pytesseract
            from PIL import Image
            doc=fitz.open(stream=data,filetype="pdf")
            for p in doc:
                pix=p.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
                text += "\n"+pytesseract.image_to_string(Image.open(io.BytesIO(pix.tobytes("png"))))
        except Exception: pass
    return text

def parse_invoice(text):
    inv=re.search(r"(ADF/\d{4}-\d{2}/\d+)",text,re.I)
    po=re.search(r"Buyer[’']s Order No\.\s*\n?\s*([0-9]{6,})",text,re.I)
    dest=re.search(r"Destination\s*\n?\s*([^\n]+)",text,re.I)
    invoice=inv.group(1) if inv else ""
    po_no=po.group(1) if po else ""
    destination=norm(dest.group(1)) if dest else ""
    d=destination.lower()
    if d=="haryana": destination="Gurgaon"
    elif d in ("banglore","bangalore"): destination="Bangalore"
    elif d=="kolkata": destination="Kolkata"
    elif d=="howrah": destination="Howrah"
    elif d=="thane": destination="Thane"
    pattern=re.compile(r"(FG-PURPLLE-PER-\d+ML-[A-Z0-9 -]+?)\s+X\s+(\d+)\s+33030050\s+([\d,]+\.\d+)\s+BOX\s+([\d,]+\.\d+)\s+PCS\s+([\d,]+\.\d+)\s+PCS\s+([\d,]+\.\d+)",re.I)
    lines=[]
    for m in pattern.finditer(text):
        sku=norm(m.group(1)); pcs=int(round(float(m.group(4).replace(",","")))); rate=float(m.group(5).replace(",",""))
        lines.append((sku,pcs,rate))
    return invoice,po_no,destination,lines

def map_sku(sku,mapping):
    s=sku.upper()
    best=None; score=-1
    for _,r in mapping.iterrows():
        n=str(r["Name of FG"]).upper(); sc=sum(1 for t in ["20ML","50ML","OUD TILL DOWN","AMBER UNTIL SUNSET","BLOOM AFTER DARK","SUGAR AFTER DUSK"] if t in s and t in n)
        if sc>score: score=sc; best=r
    if best is not None and score>=2: return str(best["EAN"]),str(best["Name of FG"])
    return "",""

st.set_page_config(page_title="ADF Invoice OCR → Excel",layout="wide")
st.title("ADF Invoice OCR → Standard Excel")
st.write("Upload the EAN/SKU master and ADF invoice PDFs. The app always produces the same Excel columns.")

mapping_file=st.file_uploader("1. EAN ↔ SKU master Excel",type=["xlsx","xls"])
pdf_files=st.file_uploader("2. Invoice PDFs",type=["pdf"],accept_multiple_files=True)

if mapping_file and pdf_files and st.button("Process invoices",type="primary"):
    try: mapping=load_mapping(mapping_file)
    except Exception as e: st.error(str(e)); st.stop()
    rows=[]; issues=[]
    for f in pdf_files:
        text=extract_text(f.read()); invoice,po,city,lines=parse_invoice(text)
        if not invoice: issues.append([f.name,"Invoice number not detected"])
        if not lines: issues.append([f.name,"No line items detected"])
        for sku,qty,rate in lines:
            ean,name=map_sku(sku,mapping)
            if not ean: issues.append([f.name,"EAN mapping not found: "+sku])
            rows.append([sku,ean,name,invoice,city,po,qty,rate])
    result=pd.DataFrame(rows,columns=OUTPUT_COLUMNS)
    if result.empty: st.error("No invoice line items extracted."); st.stop()
    result["EAN"]=result["EAN"].astype(str).str.replace(r"\.0$","",regex=True)
    result["Quantity Dispatched (PCS)"]=pd.to_numeric(result["Quantity Dispatched (PCS)"],errors="coerce").astype("Int64")
    result["Price of FG (₹/PCS)"]=pd.to_numeric(result["Price of FG (₹/PCS)"],errors="coerce")
    tmp=result.assign(_value=result["Quantity Dispatched (PCS)"]*result["Price of FG (₹/PCS)"])
    summary=tmp.groupby(["Invoice Number","City / Destination","PO Number"],dropna=False,as_index=False).agg(**{"Total Quantity Dispatched (PCS)":("Quantity Dispatched (PCS)","sum"),"Taxable FG Value (₹)":("_value","sum")})
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as writer:
        result.to_excel(writer,index=False,sheet_name="Invoice Details"); summary.to_excel(writer,index=False,sheet_name="Summary")
    st.success(f"Processed {len(pdf_files)} invoice(s) and {len(result)} line items.")
    st.dataframe(result,use_container_width=True)
    st.download_button("Download standardized Excel",out.getvalue(),"ADF_Invoice_Dispatch_Details.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if issues:
        st.warning("Some items need review."); st.dataframe(pd.DataFrame(issues,columns=["File","Issue"]),use_container_width=True)
