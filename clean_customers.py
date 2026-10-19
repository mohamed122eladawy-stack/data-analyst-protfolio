"""
Customer data cleaning & segmentation
Usage: python clean_customers.py customers_raw.csv Customers_Cleaned.xlsx
Steps: trim/standardize text -> normalize phones (Egypt) -> validate emails ->
fix city spelling -> parse dates & money -> remove duplicates -> segment ->
export cleaned data + data-quality report.
"""
import re
import sys

import pandas as pd

CITY_MAP = {"cairo": "Cairo", "cario": "Cairo", "al qahira": "Cairo", "giza": "Giza", "gizah": "Giza",
            "alexandria": "Alexandria", "alex": "Alexandria", "mansoura": "Mansoura",
            "el mansoura": "Mansoura", "aswan": "Aswan", "tanta": "Tanta"}
EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+\.[\w.]+$")


def clean_phone(p):
    if pd.isna(p):
        return None
    d = re.sub(r"\D", "", str(p))
    if d.startswith("20") and len(d) == 12:
        d = "0" + d[2:]
    return d if re.fullmatch(r"01[0125]\d{8}", d) else None


def clean_email(e):
    if pd.isna(e):
        return None
    e = str(e).strip().lower()
    return e if EMAIL_RE.match(e) else None


def clean_money(v):
    if pd.isna(v):
        return None
    s = re.sub(r"[^\d.]", "", str(v))
    return float(s) if s else None


def parse_date(s):
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return pd.to_datetime(s, format=fmt)
        except (ValueError, TypeError):
            continue
    return pd.NaT


def segment(row):
    if pd.isna(row["total_spend"]):
        return "Unknown"
    if row["total_spend"] >= 15000:
        return "VIP"
    if row["total_spend"] >= 5000:
        return "Regular"
    return "Low value"


def main(src, out):
    raw = pd.read_csv(src, dtype=str)
    df = raw.copy()
    df["name"] = df["name"].str.strip().str.replace(r"\s+", " ", regex=True).str.title()
    df["email"] = df["email"].map(clean_email)
    df["phone"] = df["phone"].map(clean_phone)
    df["city"] = df["city"].str.strip().str.lower().map(CITY_MAP)
    df["signup_date"] = df["signup_date"].map(parse_date)
    df["total_spend"] = df["total_spend"].map(clean_money)
    df["orders"] = pd.to_numeric(df["orders"], errors="coerce")

    # quality issues BEFORE removing duplicates
    issues = {
        "Total rows (raw)": len(raw),
        "Invalid/missing emails": int(df["email"].isna().sum()),
        "Invalid/missing phones": int(df["phone"].isna().sum()),
        "Missing city": int(df["city"].isna().sum()),
        "Missing total_spend": int(df["total_spend"].isna().sum()),
        "Missing orders": int(df["orders"].isna().sum()),
    }
    # duplicates: same name + same signup date + same city
    key = ["name", "signup_date", "city", "phone"]
    dup_mask = df.duplicated(subset=key, keep="first") & df["phone"].notna()
    issues["Duplicate rows removed"] = int(dup_mask.sum())
    removed = df[dup_mask]
    df = df[~dup_mask].copy()
    issues["Total rows (clean)"] = len(df)

    df["segment"] = df.apply(segment, axis=1)
    df["avg_order_value"] = (df["total_spend"] / df["orders"]).round(2)
    df["signup_date"] = df["signup_date"].dt.date

    report = pd.DataFrame(list(issues.items()), columns=["Check", "Count"])
    seg = (df.groupby("segment").agg(customers=("customer_id", "count"),
                                     total_spend=("total_spend", "sum"),
                                     avg_spend=("total_spend", "mean")).round(0).reset_index())
    by_city = (df.groupby("city", dropna=False).agg(customers=("customer_id", "count"),
                                                    total_spend=("total_spend", "sum")).reset_index())

    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        report.to_excel(xw, sheet_name="Quality Report", index=False)
        df.to_excel(xw, sheet_name="Cleaned Data", index=False)
        seg.to_excel(xw, sheet_name="Segments", index=False)
        by_city.to_excel(xw, sheet_name="By City", index=False)
        removed.to_excel(xw, sheet_name="Removed Duplicates", index=False)
        raw.to_excel(xw, sheet_name="Raw Data", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(max(len(str(c.value or "")) for c in col) + 3, 38)
    print(report.to_string(index=False))
    print(f"Saved -> {out}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit("Usage: python clean_customers.py <raw.csv> <output.xlsx>")
    main(sys.argv[1], sys.argv[2])
