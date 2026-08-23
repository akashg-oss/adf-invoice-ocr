import re
import fitz  # PyMuPDF
import pandas as pd


def extract_pdf_data(pdf_paths):
    all_data = []

    for pdf_path in pdf_paths:
        doc = fitz.open(pdf_path)

        # 1. Extract Header Fields from Page 1
        page1_text = doc[0].get_text("text")

        # Extract Invoice Number
        inv_match = re.search(r"Invoice No\.\s*([A-Z0-9/-]+)", page1_text)
        invoice_no = inv_match.group(1).strip() if inv_match else None

        # Extract Buyer's Order No. / PO Number
        po_match = re.search(
            r"Buyer'?s Order No\.\s*(\d+)", page1_text, re.IGNORECASE
        )
        po_number = po_match.group(1).strip() if po_match else None

        # Extract Destination
        dest_match = re.search(r"Destination\s*([A-Za-z]+)", page1_text)
        destination = dest_match.group(1).strip() if dest_match else "Bangalore"

        # 2. Extract All Line Items (Iterate through all pages in document)
        file_items = []
        for page in doc:
            text = page.get_text("text")
            # Match Description and Quantity lines
            # Looks for item code pattern (FG-PURPLLE...)
            lines = text.split("\n")
            for i, line in enumerate(lines):
                if line.startswith("FG-PURPLLE-"):
                    # Combine multi-line item descriptions if split across lines
                    desc = line.strip()
                    j = i + 1
                    while j < len(lines) and not re.match(
                        r"^\d{8}|\d+%", lines[j]
                    ):
                        if lines[j].strip() and not lines[j].startswith("3303"):
                            desc += " " + lines[j].strip()
                        j += 1

                    file_items.append(
                        {
                            "File Name": pdf_path.split("/")[-1],
                            "Invoice Number": invoice_no,
                            "PO Number": po_number,
                            "Destination": destination,
                            "Description of Goods": desc,
                        }
                    )

        # Assign Serial Numbers per invoice
        for idx, item in enumerate(file_items, start=1):
            item["SI No"] = idx
            all_data.append(item)

    df = pd.DataFrame(all_data)

    # Reorder columns as expected
    columns_order = [
        "File Name",
        "SI No",
        "Invoice Number",
        "PO Number",
        "Destination",
        "Description of Goods",
    ]
    return df[columns_order]


# List of PDF files to process
pdf_files = ["3206 - MANASH LIFESTYLE - SV.pdf", "3207 - MANASH LIFESTYLE - SV.pdf"]

# Run Extraction
df_result = extract_pdf_data(pdf_files)
print(df_result.to_string())
