"""
generate_data.py
────────────────────────────────────────────────────────
Generates a 2-year sales dataset for the BI dashboard.
Run once to create dashboard_data.csv
────────────────────────────────────────────────────────
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

random.seed(55)
np.random.seed(55)

PRODUCTS   = ["Laptop", "Monitor", "Keyboard", "Mouse", "Headset", "Webcam", "Docking Station", "USB Hub"]
CATEGORIES = {"Laptop":"Electronics","Monitor":"Electronics","Headset":"Electronics","Webcam":"Electronics",
               "Keyboard":"Accessories","Mouse":"Accessories","USB Hub":"Accessories","Docking Station":"Accessories"}
REGIONS    = ["North", "South", "East", "West"]
SEGMENTS   = ["Premium", "Standard", "Budget"]
REPS       = ["Alice Monroe","Bob Chen","Carol Singh","David Kim","Eva Patel","Frank Torres","Grace Lee","Henry Zhang"]
STATUSES   = ["Completed","Completed","Completed","Pending","Returned"]

PRICE = {"Laptop":1200,"Monitor":350,"Keyboard":80,"Mouse":45,
         "Headset":120,"Webcam":90,"Docking Station":200,"USB Hub":35}

rows = []
start = datetime(2024, 1, 1)

for i in range(1200):
    product  = random.choice(PRODUCTS)
    date     = start + timedelta(days=random.randint(0, 364))
    qty      = random.randint(1, 12)
    price    = PRICE[product] * random.uniform(0.9, 1.1)
    discount = round(random.uniform(0, 0.2), 2)
    revenue  = round(qty * price * (1 - discount), 2)
    status   = random.choice(STATUSES)
    rows.append({
        "Order_ID":   f"ORD-{20000+i}",
        "Date":       date.strftime("%Y-%m-%d"),
        "Product":    product,
        "Category":   CATEGORIES[product],
        "Region":     random.choice(REGIONS),
        "Segment":    random.choice(SEGMENTS),
        "Sales_Rep":  random.choice(REPS),
        "Quantity":   qty,
        "Unit_Price": round(price, 2),
        "Discount":   discount,
        "Revenue":    revenue,
        "Status":     status
    })

df = pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)
df.to_csv("dashboard_data.csv", index=False)
print(f"✅  dashboard_data.csv — {len(df)} rows")
