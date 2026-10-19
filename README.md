# PDF Invoices -> Excel Extractor (Python)
Extracts data from PDF invoices and exports a clean Excel file.

**Features**
- Reads all PDFs in a folder (pdfplumber), extracts header fields + line-item tables
- Validates each invoice: sum of lines must equal Grand Total
- Skips broken files and logs them in an Errors sheet
- Output sheets: Invoices, Line Items, By Customer

**Run**
```
pip install pdfplumber pandas openpyxl reportlab
python make_sample_invoices.py        # creates 30 fake invoices (optional)
python pdf_to_excel.py sample_invoices output/invoices_extracted.xlsx
```
Result on sample: 30 files, 107 line items, 30/30 totals validated.
Note: the regex patterns in `HEADER_PATTERNS` must be adapted to your own invoice layout.
