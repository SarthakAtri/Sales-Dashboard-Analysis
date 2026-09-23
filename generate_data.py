import random
from datetime import date, timedelta
import csv

random.seed(42)

regions = ["North", "South", "East", "West"]

products = {
    "Electronics": [
        ("Wireless Earbuds", 1800, 3200),
        ("Smartwatch", 2500, 6500),
        ("Bluetooth Speaker", 1200, 2800),
        ("Power Bank 20000mAh", 900, 1800),
        ("USB-C Charger 65W", 700, 1400),
    ],
    "Clothing": [
        ("Men's Cotton T-Shirt", 350, 700),
        ("Women's Kurti", 500, 1100),
        ("Denim Jeans", 900, 1800),
        ("Winter Jacket", 1500, 3200),
        ("Running Shoes", 1800, 3800),
    ],
    "Home & Kitchen": [
        ("Non-Stick Cookware Set", 1200, 2600),
        ("Electric Kettle", 600, 1300),
        ("LED Table Lamp", 400, 900),
        ("Air Purifier", 4500, 9000),
        ("Vacuum Cleaner", 3200, 6800),
    ],
    "Sports": [
        ("Yoga Mat", 400, 900),
        ("Dumbbell Set 10kg", 1500, 2800),
        ("Cricket Bat", 900, 2200),
        ("Football", 500, 1200),
        ("Cycling Helmet", 700, 1600),
    ],
}

start = date(2024, 1, 1)
end = date(2024, 12, 31)
delta_days = (end - start).days

rows = []
num_transactions = 900

# region demand weighting so charts show meaningful differences
region_weight = {"North": 1.15, "South": 1.0, "East": 0.85, "West": 1.05}
# add a seasonal boost for certain months per category to make the monthly trend interesting
def seasonal_multiplier(month, category):
    if category == "Clothing" and month in (10, 11, 12):
        return 1.4
    if category == "Electronics" and month in (11, 12):
        return 1.5
    if category == "Sports" and month in (1, 6, 7):
        return 1.25
    if category == "Home & Kitchen" and month in (3, 4):
        return 1.2
    return 1.0

for _ in range(num_transactions):
    txn_date = start + timedelta(days=random.randint(0, delta_days))
    category = random.choice(list(products.keys()))
    product_name, price_low, price_high = random.choice(products[category])
    region = random.choices(regions, weights=[region_weight[r] for r in regions])[0]

    mult = seasonal_multiplier(txn_date.month, category) * region_weight[region]
    base_units = random.randint(1, 8)
    units = max(1, round(base_units * mult))
    unit_price = round(random.uniform(price_low, price_high), 2)

    rows.append({
        "Date": txn_date.strftime("%Y-%m-%d"),
        "Region": region,
        "Category": category,
        "Product": product_name,
        "Units": units,
        "UnitPrice": unit_price,
    })

rows.sort(key=lambda r: r["Date"])

with open("/home/claude/sales_dashboard/sales_raw.csv", "w", newline="") as f:
    writer = csv.DictWriter(f, fieldnames=["Date", "Region", "Category", "Product", "Units", "UnitPrice"])
    writer.writeheader()
    writer.writerows(rows)

print(f"Generated {len(rows)} transactions")
print("Unique products:", sum(len(v) for v in products.values()))
