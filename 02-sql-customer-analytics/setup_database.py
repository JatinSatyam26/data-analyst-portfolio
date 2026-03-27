"""
setup_database.py
─────────────────────────────────────────────────────────────
Creates and populates the SQLite database: customer_analytics.db

Tables:
  - customers   : customer profiles with segments and regions
  - products    : product catalogue with categories and pricing
  - orders      : order header (customer, date, status)
  - order_items : line items per order (product, qty, price)
─────────────────────────────────────────────────────────────
"""

import sqlite3
import random
from datetime import datetime, timedelta

random.seed(99)

DB_PATH = "customer_analytics.db"

# ── Schema ────────────────────────────────────────────────────
SCHEMA = """
CREATE TABLE IF NOT EXISTS customers (
    customer_id   INTEGER PRIMARY KEY,
    name          TEXT    NOT NULL,
    email         TEXT    UNIQUE,
    region        TEXT,
    segment       TEXT,
    signup_date   TEXT,
    age           INTEGER
);

CREATE TABLE IF NOT EXISTS products (
    product_id    INTEGER PRIMARY KEY,
    product_name  TEXT    NOT NULL,
    category      TEXT,
    unit_price    REAL,
    cost_price    REAL
);

CREATE TABLE IF NOT EXISTS orders (
    order_id      INTEGER PRIMARY KEY,
    customer_id   INTEGER REFERENCES customers(customer_id),
    order_date    TEXT,
    status        TEXT,
    shipping_city TEXT
);

CREATE TABLE IF NOT EXISTS order_items (
    item_id       INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id      INTEGER REFERENCES orders(order_id),
    product_id    INTEGER REFERENCES products(product_id),
    quantity      INTEGER,
    unit_price    REAL,
    discount      REAL
);
"""

# ── Seed Data ─────────────────────────────────────────────────
FIRST_NAMES = ["Aiden","Priya","Carlos","Sofia","James","Mei","Omar","Nina",
               "Raj","Elena","Marcus","Zara","Tyler","Amara","Kevin","Leila",
               "Sam","Yuki","Andre","Fatima","Noah","Chloe","Diego","Aisha"]
LAST_NAMES  = ["Johnson","Patel","Garcia","Kim","Williams","Chen","Ali","Müller",
               "Brown","Singh","Davis","Ahmed","Wilson","Nguyen","Martinez","Lee"]
REGIONS     = ["North","South","East","West","Central"]
SEGMENTS    = ["Premium","Standard","Budget"]
CITIES      = ["Phoenix","Austin","Chicago","Seattle","Miami","Denver","Boston","Atlanta"]

PRODUCTS = [
    (1,  "Laptop Pro 15",       "Electronics",   1299.99,  780.00),
    (2,  "Wireless Mouse",      "Accessories",     45.99,   12.00),
    (3,  "Mechanical Keyboard", "Accessories",     89.99,   28.00),
    (4,  "4K Monitor",          "Electronics",    399.99,  210.00),
    (5,  "USB-C Hub",           "Accessories",     59.99,   18.00),
    (6,  "Webcam HD",           "Electronics",     79.99,   25.00),
    (7,  "Desk Lamp LED",       "Office",          34.99,    9.00),
    (8,  "Ergonomic Chair",     "Furniture",      349.99,  140.00),
    (9,  "Standing Desk",       "Furniture",      599.99,  260.00),
    (10, "Noise Cancelling Headphones", "Electronics", 249.99, 90.00),
    (11, "Laptop Stand",        "Accessories",     39.99,   11.00),
    (12, "External SSD 1TB",    "Storage",        109.99,   45.00),
    (13, "Smart Whiteboard",    "Office",         199.99,   80.00),
    (14, "Cable Management Kit","Accessories",     19.99,    5.00),
    (15, "Portable Charger",    "Electronics",     49.99,   15.00),
]

def generate_customers(n=200):
    rows = []
    start = datetime(2021, 1, 1)
    for i in range(1, n + 1):
        name    = f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"
        email   = f"user{i}@example.com"
        region  = random.choice(REGIONS)
        segment = random.choices(SEGMENTS, weights=[20, 55, 25])[0]
        signup  = (start + timedelta(days=random.randint(0, 900))).strftime("%Y-%m-%d")
        age     = random.randint(22, 65)
        rows.append((i, name, email, region, segment, signup, age))
    return rows

def generate_orders(n=800):
    rows = []
    start = datetime(2023, 1, 1)
    statuses = ["Completed", "Completed", "Completed", "Pending", "Returned", "Cancelled"]
    for i in range(1, n + 1):
        cust_id = random.randint(1, 200)
        date    = (start + timedelta(days=random.randint(0, 364))).strftime("%Y-%m-%d")
        status  = random.choice(statuses)
        city    = random.choice(CITIES)
        rows.append((i, cust_id, date, status, city))
    return rows

def generate_order_items(orders):
    rows = []
    for order_id, _, _, status, _ in orders:
        n_items = random.randint(1, 4)
        products_chosen = random.sample(PRODUCTS, n_items)
        for prod in products_chosen:
            qty      = random.randint(1, 5)
            price    = prod[3] * random.uniform(0.95, 1.05)
            discount = round(random.uniform(0, 0.15), 2)
            rows.append((order_id, prod[0], qty, round(price, 2), discount))
    return rows

def build():
    conn = sqlite3.connect(DB_PATH)
    cur  = conn.cursor()
    cur.executescript(SCHEMA)

    customers  = generate_customers(200)
    orders     = generate_orders(800)
    items      = generate_order_items(orders)

    cur.executemany("INSERT OR IGNORE INTO customers VALUES (?,?,?,?,?,?,?)", customers)
    cur.executemany("INSERT OR IGNORE INTO products  VALUES (?,?,?,?,?)",     PRODUCTS)
    cur.executemany("INSERT OR IGNORE INTO orders    VALUES (?,?,?,?,?)",     orders)
    cur.executemany(
        "INSERT INTO order_items (order_id,product_id,quantity,unit_price,discount) VALUES (?,?,?,?,?)",
        items
    )

    conn.commit()

    print("✅  Database built successfully!")
    print(f"    Customers  : {cur.execute('SELECT COUNT(*) FROM customers').fetchone()[0]}")
    print(f"    Products   : {cur.execute('SELECT COUNT(*) FROM products').fetchone()[0]}")
    print(f"    Orders     : {cur.execute('SELECT COUNT(*) FROM orders').fetchone()[0]}")
    print(f"    Order Items: {cur.execute('SELECT COUNT(*) FROM order_items').fetchone()[0]}")
    conn.close()

if __name__ == "__main__":
    build()
