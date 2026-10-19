"""Generate 30 sample invoice PDFs (fake data) to test the extractor."""
import random
from datetime import date, timedelta
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer
from reportlab.lib.styles import getSampleStyleSheet

random.seed(7)
CUSTOMERS = ["Nile Trading Co.", "Delta Supplies", "Alex Electronics", "Cairo Foods", "Giza Stationery", "Sinai Hardware"]
PRODUCTS = [("Laptop", 18500), ("Mouse", 250), ("Keyboard", 450), ("Monitor", 4200),
            ("Printer", 3800), ("Router", 900), ("USB Cable", 60), ("Webcam", 700)]

styles = getSampleStyleSheet()
for i in range(1, 31):
    inv_no = f"INV-2026-{i:04d}"
    d = date(2026, 1, 1) + timedelta(days=random.randint(0, 270))
    cust = random.choice(CUSTOMERS)
    items = random.sample(PRODUCTS, random.randint(2, 5))
    rows = [["Item", "Qty", "Unit Price", "Total"]]
    grand = 0
    for name, price in items:
        q = random.randint(1, 10)
        rows.append([name, str(q), f"{price:,.2f}", f"{q*price:,.2f}"])
        grand += q * price
    doc = SimpleDocTemplate(f"sample_invoices/{inv_no}.pdf", pagesize=A4)
    t = Table(rows, colWidths=[200, 60, 100, 100])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                           ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                           ("ALIGN", (1, 0), (-1, -1), "RIGHT")]))
    doc.build([Paragraph("INVOICE", styles["Title"]),
               Paragraph(f"Invoice No: {inv_no}", styles["Normal"]),
               Paragraph(f"Date: {d.isoformat()}", styles["Normal"]),
               Paragraph(f"Customer: {cust}", styles["Normal"]),
               Spacer(1, 20), t, Spacer(1, 15),
               Paragraph(f"Grand Total: {grand:,.2f} EGP", styles["Heading3"])])
print("Created 30 invoices")
