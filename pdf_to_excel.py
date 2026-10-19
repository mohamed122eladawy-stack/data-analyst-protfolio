"""
PDF Invoices -> Excel extractor
Usage:  python pdf_to_excel.py <folder_with_pdfs> <output.xlsx>
Reads every PDF in the folder, pulls invoice header fields and line-item
tables, validates totals, and writes a clean multi-sheet Excel file.
"""
import re
import sys
from pathlib import Path

import pandas as pd
import pdfplumber

HEADER_PATTERNS = {
    "invoice_no": r"Invoice No:\s*(\S+)",
    "date": r"Date:\s*(\d{4}-\d{2}-\d{2})",
    "customer": r"Customer:\s*(.+)",
    "grand_total": r"Grand Total:\s*([\d,]+\.\d{2})",
}


def to_number(text):
    return float(str(text).replace(",", "").strip())


def extract_invoice(pdf_path):
    """Return (header_dict, list_of_line_items) for one PDF."""
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        tables = [t for page in pdf.pages for t in page.extract_tables()]

    header = {"file": pdf_path.name}
    for key, pattern in HEADER_PATTERNS.items():
        m = re.search(pattern, text)
        header[key] = m.group(1).strip() if m else None

    items = []
    for table in tables:
        for row in table[1:]:  # skip the header row
            if row and len(row) >= 4 and row[0]:
                items.append({
                    "invoice_no": header["invoice_no"],
                    "item": row[0].strip(),
                    "qty": int(row[1]),
                    "unit_price": to_number(row[2]),
                    "line_total": to_number(row[3]),
                })
    return header, items


def main(folder, output):
    headers, all_items, errors = [], [], []
    pdfs = sorted(Path(folder).glob("*.pdf"))
    for pdf in pdfs:
        try:
            h, items = extract_invoice(pdf)
            calc = sum(i["line_total"] for i in items)
            h["calculated_total"] = round(calc, 2)
            h["total_matches"] = abs(calc - to_number(h["grand_total"])) < 0.01
            h["grand_total"] = to_number(h["grand_total"])
            headers.append(h)
            all_items.extend(items)
        except Exception as e:  # keep going if one file is broken
            errors.append({"file": pdf.name, "error": str(e)})

    inv = pd.DataFrame(headers)
    inv["date"] = pd.to_datetime(inv["date"])
    inv = inv.sort_values("date")
    items_df = pd.DataFrame(all_items)
    by_customer = (inv.groupby("customer")["grand_total"]
                   .agg(invoices="count", total_sales="sum").reset_index()
                   .sort_values("total_sales", ascending=False))

    with pd.ExcelWriter(output, engine="openpyxl") as xw:
        inv.to_excel(xw, sheet_name="Invoices", index=False)
        items_df.to_excel(xw, sheet_name="Line Items", index=False)
        by_customer.to_excel(xw, sheet_name="By Customer", index=False)
        if errors:
            pd.DataFrame(errors).to_excel(xw, sheet_name="Errors", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col) + 3
                ws.column_dimensions[col[0].column_letter].width = min(width, 40)

    print(f"Processed: {len(pdfs)} files | OK: {len(headers)} | Errors: {len(errors)}")
    print(f"Totals validated: {int(inv['total_matches'].sum())}/{len(inv)}")
    print(f"Saved -> {output}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("Usage: python pdf_to_excel.py <pdf_folder> <output.xlsx>")
    main(sys.argv[1], sys.argv[2])
