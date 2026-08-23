import difflib
import io
import re
import pandas as pd
from pdfminer.high_level import extract_text
import streamlit as st

st.set_page_config(
    page_title="ADF Invoice OCR & SKU Matcher", layout="wide"
)
st.title("📄 ADF Invoice OCR & SKU Matcher")


def parse_and_match_invoices(pdf_file, df_master):
    # Use pdfminer to read corrupted/invalid PDF streams without throwing PdfReadError
    pdf_bytes = io.BytesIO(pdf_file.read())
    full_text = extract_text(pdf_bytes) or ""

    # 1. Extract Header Details
    inv_match = re.search(r"ADF/\d{4}-\d{2}/\d+", full_text, re.IGNORECASE)
    inv_num = inv_match.group(0) if inv_match else "UNKNOWN"

    po_match = re.search(
        r"Buyer(?:'|’|s|\s)*Order\s*No\.?\s*\n?\s*([A-Z0-9/-]+)",
        full_text,
        re.IGNORECASE,
    )
    po_number = po_match.group(1).strip() if po_match else "UNKNOWN"

    dest_match = re.search(
        r"Destination\s*\n?\s*([A-Za-z0-9\s.\-]+?)(?=\n[A-Z][a-z]|\n\n|\nMotor|\nTerms|\Z)",
        full_text,
        re.IGNORECASE,
    )
    destination = dest_match.group(1).strip() if dest_match else "UNKNOWN"

    # Helper Functions
    def clean_tokens(text):
        s = re.sub(r"[^A-Z0-9]", " ", str(text).upper())
        stop_words = {
            "FG",
            "PURPLLE",
            "PER",
            "STAY",
            "EAU",
            "DE",
            "PARFUM",
            "MINI",
            "FACES",
            "CANADA",
            "33030050",
            "PCS",
            "BOX",
            "X",
            "20ML",
            "50ML",
            "100ML",
        }
        return [t for t in s.split() if t not in stop_words]

    def extract_size(text):
        m = re.search(r"(\d+\s*ML)", str(text), re.IGNORECASE)
        return m.group(1).upper().replace(" ", "") if m else ""

    def match_sku(raw_item_desc):
        vol = extract_size(raw_item_desc)
        subset = df_master.copy()

        if vol == "20ML":
            f = subset["SKU Name"].str.contains(
                "20ml|mini", case=False, na=False
            )
            if f.any():
                subset = subset[f]
        elif vol == "50ML":
            f = subset["SKU Name"].str.contains(
                "50ml", case=False, na=False
            ) & ~subset["SKU Name"].str.contains(
                "20ml|mini", case=False, na=False
            )
            if f.any():
                subset = subset[f]

        inv_tokens = clean_tokens(raw_item_desc)
        inv_str = " ".join(inv_tokens)

        best_row, best_score = None, -1.0
        for idx, row in subset.iterrows():
            m_tokens = clean_tokens(row["SKU Name"])
            m_str = " ".join(m_tokens)
            seq_ratio = difflib.SequenceMatcher(None, inv_str, m_str).ratio()
            token_scores = [
                max(
                    [
                        difflib.SequenceMatcher(None, it, mt).ratio()
                        for mt in m_tokens
                    ]
                    or [0]
                )
                for it in inv_tokens
            ]
            avg_token_score = (
                sum(token_scores) / len(token_scores) if token_scores else 0
            )
            total_score = (seq_ratio * 0.4) + (avg_token_score * 0.6)
            if total_score > best_score:
                best_score = total_score
                best_row = row

        if best_row is not None and best_score >= 0.35:
            return (
                best_row["SKU Code"],
                str(best_row["EAN"]),
                best_row["SKU Name"],
                round(best_score, 2),
            )
        return "", "", raw_item_desc, 0.0

    # 2. Extract Line Items
    items = []
    lines = [line.strip() for line in full_text.split("\n") if line.strip()]

    i = 0
    while i < len(lines):
        line = lines[i]
        sl_match = re.match(r"^(\d+)\s+(FG-PURPLLE-[A-Z0-9-]+)", line)
        if sl_match:
            sl_no = sl_match.group(1)
            desc_parts = [sl_match.group(2)]

            i += 1
            while i < len(lines) and not re.search(
                r"\d{8}", lines[i]
            ):  # HSN 33030050
                if lines[i]:
                    desc_parts.append(lines[i])
                i += 1

            raw_desc = " ".join(desc_parts)
            clean_desc = re.sub(
                r"\s+X\s+\d+$", "", raw_desc, flags=re.IGNORECASE
            ).strip()

            # Find item metrics across nearby lines
            qty_pcs, rate, amount = 0.0, 0.0, 0.0
            search_window = " ".join(lines[i : i + 10])

            amt_match = re.search(r"([\d,]+\.\d{2})", search_window)
            pcs_match = re.search(
                r"([\d,]+(?:\.\d+)?)\s*PCS", search_window, re.IGNORECASE
            )

            if amt_match:
                amount = float(amt_match.group(1).replace(",", ""))
            if pcs_match:
                qty_pcs = float(pcs_match.group(1).replace(",", ""))

            sku_code, ean, sku_name, score = match_sku(clean_desc)

            items.append({
                "Sl No": int(sl_no),
                "SKU Code": sku_code,
                "EAN": ean,
                "Name of FG": sku_name,
                "Invoice Raw Desc": clean_desc,
                "Invoice Number": inv_num,
                "PO Number": po_number,
                "Destination": destination,
                "Quantity Dispatched (PCS)": int(qty_pcs),
                "Amount (₹)": amount,
                "Match Score": score,
            })
        i += 1

    return pd.DataFrame(items)


# Streamlit Interface
st.sidebar.header("📁 File Uploads")
uploaded_excel = st.sidebar.file_uploader(
    "Upload Master SKU Excel", type=["xlsx", "xls"]
)
uploaded_pdfs = st.sidebar.file_uploader(
    "Upload Invoice PDFs", type=["pdf"], accept_multiple_files=True
)

if uploaded_excel and uploaded_pdfs:
    df_master = pd.read_excel(uploaded_excel)

    if st.button("Process Invoices"):
        all_dfs = []
        for pdf_file in uploaded_pdfs:
            df_res = parse_and_match_invoices(pdf_file, df_master)
            all_dfs.append(df_res)

        if all_dfs:
            final_df = pd.concat(all_dfs, ignore_index=True)
            st.success(
                f"Successfully extracted {len(final_df)} items across {len(uploaded_pdfs)} invoice(s)."
            )
            st.dataframe(final_df, use_container_width=True)

            csv_data = final_df.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="📥 Download Extracted Results CSV",
                data=csv_data,
                file_name="extracted_invoice_matches.csv",
                mime="text/csv",
            )
elif not uploaded_excel:
    st.info("👈 Please upload the Master SKU Excel file in the sidebar.")
elif not uploaded_pdfs:
    st.info("👈 Please upload Invoice PDFs in the sidebar.")
