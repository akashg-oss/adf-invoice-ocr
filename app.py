import difflib
import re
import pandas as pd
import pdfplumber  # Replace pypdf import


def parse_and_match_invoices(pdf_path, master_excel_path):
    df_master = pd.read_excel(master_excel_path)

    # Extract text using pdfplumber
    full_text = ""
    pages_text = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text() or ""
            pages_text.append(t)
            full_text += t + "\n"

    # 1. Extract Invoice Number
    inv_match = re.search(r"ADF/\d{4}-\d{2}/\d+", full_text)
    inv_num = inv_match.group(0) if inv_match else "UNKNOWN"

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

    # 2. Line Items Extraction
    items = []
    p1_lines = pages_text[0].split("\n")

    i = 0
    while i < len(p1_lines):
        line = p1_lines[i].strip()
        sl_match = re.match(r"^(\d+)\s+(FG-PURPLLE-[A-Z0-9-]+)", line)
        if sl_match:
            sl_no = sl_match.group(1)
            desc_parts = [sl_match.group(2)]

            i += 1
            while i < len(p1_lines) and "PCS" not in p1_lines[i]:
                if p1_lines[i].strip():
                    desc_parts.append(p1_lines[i].strip())
                i += 1

            raw_desc = " ".join(desc_parts)
            clean_desc = re.sub(
                r"\s+X\s+\d+$", "", raw_desc, flags=re.IGNORECASE
            ).strip()

            num_line = p1_lines[i] if i < len(p1_lines) else ""
            m_vals = re.search(
                r"([\d,]+\.\d+)PCS(\d+\.\d{2})([\d,]+\.\d+)\s*PCS", num_line
            )

            if m_vals:
                amount = float(m_vals.group(1).replace(",", ""))
                rate = float(m_vals.group(2))
                qty_pcs = float(m_vals.group(3).replace(",", ""))
            else:
                amount, rate, qty_pcs = 0.0, 0.0, 0.0

            sku_code, ean, sku_name, score = match_sku(clean_desc)

            items.append({
                "Sl No": int(sl_no),
                "SKU Code": sku_code,
                "EAN": ean,
                "Name of FG": sku_name,
                "Invoice Raw Desc": clean_desc,
                "Invoice Number": inv_num,
                "Quantity Dispatched (PCS)": int(qty_pcs),
                "Price of FG (₹/PCS)": rate,
                "Amount (₹)": amount,
                "Match Score": score,
            })
        i += 1

    return pd.DataFrame(items)
