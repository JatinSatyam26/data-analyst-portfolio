"""
generate_data.py
Generates a realistic sample sales dataset for demonstration purposes.
Run this once to create sample_data.csv
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

random.seed(42)
np.random.seed(42)

PRODUCTS = ["Laptop", "Monitor", "Keyboard", "Mouse", "Headset", "Webcam", "Docking Station", "USB Hub"]
REGIONS  = ["North", "South", "East", "West"]
REPS     = ["Alice Monroe", "Bob Chen", "Carol Singh", "David Kim", "Eva Patel",
            "Frank Torres", "Grace Lee", "Henry Zhang"]

PRODUCT_PRICE = {
    "Laptop": 1200, "Monitor": 350, "Keyboard": 80, "Mouse": 45,
    "Headset": 120, "Webcam": 90, "Docking Station": 200, "USB Hub": 35
}

rows = []
start = datetime(2024, 1, 1)

for _ in range(600):
    product  = random.choice(PRODUCTS)
    region   = random.choice(REGIONS)
    rep      = random.choice(REPS)
    date     = start + timedelta(days=random.randint(0, 364))
    qty      = random.randint(1, 15)
    price    = PRODUCT_PRICE[product] * round(random.uniform(0.90, 1.10), 2)
    discount = round(random.uniform(0, 0.20), 2)
    revenue  = round(qty * price * (1 - discount), 2)
    status   = random.choices(["Completed", "Pending", "Returned"],
                               weights=[80, 12, 8])[0]

    rows.append({
        "Order_ID":      f"ORD-{random.randint(10000,99999)}",
        "Date":          date.strftime("%Y-%m-%d"),
        "Product":       product,
        "Region":        region,
        "Sales_Rep":     rep,
        "Quantity":      qty,
        "Unit_Price":    round(price, 2),
        "Discount":      discount,
        "Revenue":       revenue,
        "Order_Status":  status
    })

df = pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)
df.to_csv("sample_data.csv", index=False)
print(f"✅  sample_data.csv created — {len(df)} rows")
