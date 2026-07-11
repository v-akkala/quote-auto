from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "quote_app"))

from generator import COUNTER_PATH, ProductItem, TestItem, generate_quote

if COUNTER_PATH.exists():
    COUNTER_PATH.unlink()

sample_data = {
    "quote_date": "7-11-2026",
    "email_date": "7-11-2026",
    "product_name": "Inverter",
    "dim1": "100",
    "dim2": "200",
    "dim3": "300",
    "dim_unit": "mm",
    "weight": "12",
    "weight_unit": "kg",
    "customer": "Test Customer",
    "designation": "Engineer",
    "company": "Example Company",
    "tel_no": "1234567890",
    "email": "test@example.com",
    "address": "123 Test Street",
}

tests = [
    TestItem("Vibration", "Sample vibration requirement", "1", "1", "10000"),
    TestItem("High Temperature", "55 C for 24 hours", "2", "1", "20000"),
]

products = [
    ProductItem("prod 1", "1", "1", "1", "mm", "1", "kg"),
    ProductItem("prod 2", "2", "2", "2", "mm", "2", "kg"),
]

output = generate_quote(sample_data, tests, products)
print(output)

if COUNTER_PATH.exists():
    COUNTER_PATH.unlink()
