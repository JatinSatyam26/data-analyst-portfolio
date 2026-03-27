"""
generate_data.py
────────────────────────────────────────────────────────
Generates a realistic daily sales CSV dataset.
Run once to create sales_data.csv
────────────────────────────────────────────────────────
"""

import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

random.seed(21)
np.random.seed(21)

PRODUCTS  = ["Laptop", "Monitor", "Keyboard", "Mouse", "Headset", "Webcam", "Docking Station"]
REGIONS   = ["North", "South", "East", "West"]
REPS      = ["Alice Monroe", "Bob Chen", "Carol Singh", "David Kim", "Eva Patel", "Frank Torres"]
STATUSES  = ["Completed", "Completed", "Completed", "Pending", "Returned"]

PRICE_MAP = {
    "Laptop": 1200, "Monitor": 350, "Keyboard": 80,
    "Mouse": 45, "Headset": 120, "Webcam": 90, "Docking Station": 200
}

rows = []
start = datetime.today() - timedelta(days=30)

for i in range(400):
    product = random.choice(PRODUCTS)
    date    = start + timedelta(days=random.randint(0, 29))
    qty     = random.randint(1, 10)
    price   = PRICE_MAP[product] * random.uniform(0.92, 1.08)
    revenue = round(qty * price * random.uniform(0.85, 1.0), 2)
    rows.append({
        "Order_ID":    f"ORD-{10000 + i}",
        "Date":        date.strftime("%Y-%m-%d"),
        "Product":     product,
        "Region":      random.choice(REGIONS),
        "Sales_Rep":   random.choice(REPS),
        "Quantity":    qty,
        "Revenue":     revenue,
        "Status":      random.choice(STATUSES)
    })

df = pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)
df.to_csv("sales_data.csv", index=False)
print(f"✅  sales_data.csv created — {len(df)} rows")
