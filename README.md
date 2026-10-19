# Customer Data Cleaning & Segmentation (Python + Excel)
Cleans a messy customer file (460 rows, fictional) and segments customers.

**Problems handled**
- Extra spaces / mixed name casing, city spelling variants (Cario, Alex, ...)
- Phone numbers in many formats -> standard 01XXXXXXXXX, invalid ones flagged
- Invalid emails, mixed date formats, money stored as text ("EGP 5,000")
- Duplicate customers removed (29 found), kept in a separate sheet for audit

**Output (Customers_Cleaned.xlsx):** Quality Report, Cleaned Data, Segments (VIP / Regular / Low value), By City, Removed Duplicates, Raw Data.

**Run**
```
python make_messy_data.py
python clean_customers.py customers_raw.csv Customers_Cleaned.xlsx
```
