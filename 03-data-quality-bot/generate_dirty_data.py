"""
generate_dirty_data.py
─────────────────────────────────────────────────────────────
Generates a realistic but intentionally dirty CSV dataset
to demonstrate the Data Quality Bot's detection capabilities.

Issues planted:
  - Missing values in critical columns
  - Duplicate rows
  - Negative values in numeric columns
  - Invalid email formats
  - Out-of-range ages
  - Inconsistent text casing
  - Invalid date formats
  - Outlier revenue values
─────────────────────────────────────────────────────────────
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

random.seed(77)
np.random.seed(77)

FIRST   = ["Alice","Bob","Carol","David","Eva","Frank","Grace","Henry",
           "Iris","Jake","Karen","Leo","Mia","Nathan","Olivia","Paul"]
LAST    = ["Smith","Jones","Brown","Taylor","Wilson","Davis","Clark","Hall"]
REGIONS = ["North","South","East","West"]
PRODUCTS= ["Laptop","Monitor","Keyboard","Mouse","Headset"]

def rand_email(name):
    return f"{name.lower().replace(' ','.')}@example.com"

rows = []
start = datetime(2023, 1, 1)

for i in range(1, 301):
    name    = f"{random.choice(FIRST)} {random.choice(LAST)}"
    email   = rand_email(name)
    age     = random.randint(22, 60)
    region  = random.choice(REGIONS)
    product = random.choice(PRODUCTS)
    qty     = random.randint(1, 10)
    revenue = round(random.uniform(50, 2000), 2)
    date    = (start + timedelta(days=random.randint(0, 364))).strftime("%Y-%m-%d")
    status  = random.choice(["Completed","Pending","Returned"])

    rows.append([i, name, email, age, region, product, qty, revenue, date, status])

df = pd.DataFrame(rows, columns=[
    "customer_id","name","email","age","region",
    "product","quantity","revenue","order_date","status"
])

# ── Plant Issues ──────────────────────────────────────────────

# 1. Missing values in critical columns
for idx in random.sample(range(300), 18):
    col = random.choice(["name","email","revenue","region"])
    df.at[idx, col] = np.nan

# 2. Duplicate rows (copy 10 rows)
dupes = df.sample(10, random_state=1)
df = pd.concat([df, dupes], ignore_index=True)

# 3. Negative revenue
for idx in random.sample(range(300), 8):
    df.at[idx, "revenue"] = round(random.uniform(-500, -10), 2)

# 4. Invalid emails
bad_emails = ["notanemail","@nodomain","user@@double.com","missingat.com",""]
for idx in random.sample(range(300), 12):
    df.at[idx, "email"] = random.choice(bad_emails)

# 5. Out-of-range ages
for idx in random.sample(range(300), 6):
    df.at[idx, "age"] = random.choice([-5, 0, 150, 999])

# 6. Inconsistent text casing
for idx in random.sample(range(300), 15):
    df.at[idx, "region"] = random.choice(["NORTH","south","EAST","west","North "])

# 7. Invalid date formats
bad_dates = ["13/45/2023","2023-99-01","not-a-date","01-01-23",""]
for idx in random.sample(range(300), 10):
    df.at[idx, "order_date"] = random.choice(bad_dates)

# 8. Revenue outliers (extreme values)
for idx in random.sample(range(300), 5):
    df.at[idx, "revenue"] = round(random.uniform(50000, 200000), 2)

df.to_csv("dirty_data.csv", index=False)
print(f"✅  dirty_data.csv created — {len(df)} rows with intentional quality issues")
