"""Generate a deliberately messy customer file (fake data) for the cleaning project."""
import random
import pandas as pd

random.seed(11)
first = ["Ahmed", "Mohamed", "Mona", "Sara", "Omar", "Hassan", "Nour", "Youssef", "Salma", "Karim", "Laila", "Tarek"]
last = ["Hassan", "Ali", "Mostafa", "Ibrahim", "Saleh", "Adel", "Farouk", "Nabil", "Samir", "Khaled"]
cities = ["Cairo", "Giza", "Alexandria", "Mansoura", "Aswan", "Tanta"]
city_variants = {"Cairo": ["cairo", "CAIRO", "Cario", "Al Qahira"], "Giza": ["giza", "GIZA", "Gizah"],
                 "Alexandria": ["alexandria", "Alex", "ALEXANDRIA"], "Mansoura": ["mansoura", "El Mansoura"],
                 "Aswan": ["aswan", "ASWAN"], "Tanta": ["tanta", "TANTA"]}

def phone():
    n = "01" + random.choice("0125") + "".join(random.choice("0123456789") for _ in range(8))
    return random.choice([n, f"+2{n}", f"{n[:4]} {n[4:7]} {n[7:]}", f"{n[:4]}-{n[4:]}", n[:-2], ""])

def date():
    y, m, d = 2025, random.randint(1, 12), random.randint(1, 28)
    return random.choice([f"{y}-{m:02d}-{d:02d}", f"{d:02d}/{m:02d}/{y}", f"{d}-{m}-{y}"])

rows = []
for i in range(1, 421):
    f, l, c = random.choice(first), random.choice(last), random.choice(cities)
    name = random.choice([f"{f} {l}", f"  {f.lower()} {l.upper()} ", f"{f}  {l}"])
    email = f"{f}.{l}{random.randint(1,99)}@{random.choice(['gmail.com','yahoo.com','outlook.com'])}"
    email = random.choice([email, email.upper(), email.replace("@", ""), "", f" {email} "])
    city = random.choice([c, c, random.choice(city_variants[c]), None])
    spend = random.randint(200, 25000)
    spend = random.choice([str(spend), f"{spend:,}", f"{spend} EGP", f"EGP {spend:,}", ""])
    rows.append({"customer_id": 1000 + i, "name": name, "email": email, "phone": phone(),
                 "city": city, "signup_date": date(), "orders": random.choice([random.randint(1, 30), None]),
                 "total_spend": spend})
df = pd.DataFrame(rows)
dups = df.sample(40, random_state=3).copy()          # exact-ish duplicates with new IDs
dups["customer_id"] = range(2000, 2040)
dups["name"] = dups["name"].str.upper()
df = pd.concat([df, dups]).sample(frac=1, random_state=5).reset_index(drop=True)
df.to_csv("customers_raw.csv", index=False)
print("Raw file:", df.shape)
