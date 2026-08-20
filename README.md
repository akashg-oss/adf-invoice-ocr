# ADF Invoice OCR → Standard Excel

Streamlit application for turning ADF/Aroma De France invoice PDFs into a fixed Excel format.

Output columns:
SKU Code | EAN | Name of FG | Invoice Number | City / Destination | PO Number | Quantity Dispatched (PCS) | Price of FG (₹/PCS)

## Local setup
python -m venv .venv
pip install -r requirements.txt
streamlit run app.py

For scanned PDFs, install Tesseract OCR locally.

## Streamlit Community Cloud
1. Create a GitHub repository.
2. Upload app.py, requirements.txt, packages.txt and README.md.
3. In Streamlit Community Cloud, create an app from the repository and select app.py.
4. Deploy.

The parser is tailored to the ADF invoice format in the supplied sample invoices. If ADF changes its invoice layout, the parsing rules should be updated.
