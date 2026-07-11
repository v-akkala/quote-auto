from __future__ import annotations

import json
import re
import shutil
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.cell_range import CellRange

ROOT_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = ROOT_DIR / "templates" / "Template Quote.xlsx"
OUTPUT_DIR = ROOT_DIR / "outputs"
DATA_DIR = ROOT_DIR / "data"
COUNTER_PATH = DATA_DIR / "quote_counter.json"

QUOTE_PREFIX = "SQ2627"
STARTING_QUOTE_NUMBER = 1
SAC_CODE = 998346
TEST_START_ROW = 14
TOTAL_START_ROW = 15
TEST_MERGES = ((2, 3), (4, 30), (31, 33), (35, 40), (41, 45), (46, 50))


@dataclass
class TestItem:
    name: str
    requirements: str
    qty: str = "1"
    batch: str = "1"
    total_cost: str = ""


@dataclass
class ProductItem:
    name: str
    dim1: str = ""
    dim2: str = ""
    dim3: str = ""
    dim_unit: str = "mm"
    weight: str = ""
    weight_unit: str = "kg"


def today_string() -> str:
    now = date.today()
    return f"{now.month}-{now.day}-{now.year}"


def next_quote_ref() -> str:
    return f"{QUOTE_PREFIX}{_read_counter():04d}"


def _read_counter() -> int:
    if not COUNTER_PATH.exists():
        return STARTING_QUOTE_NUMBER
    try:
        data = json.loads(COUNTER_PATH.read_text(encoding="utf-8"))
        value = int(data.get("next_number", STARTING_QUOTE_NUMBER))
        return max(value, STARTING_QUOTE_NUMBER)
    except (ValueError, TypeError, json.JSONDecodeError):
        return STARTING_QUOTE_NUMBER


def _write_counter(next_number: int) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    COUNTER_PATH.write_text(
        json.dumps({"next_number": next_number}, indent=2),
        encoding="utf-8",
    )


def reserve_next_quote_ref() -> str:
    number = _read_counter()
    _write_counter(number + 1)
    return f"{QUOTE_PREFIX}{number:04d}"


def parse_date_text(value: str) -> str:
    value = value.strip()
    if not value:
        return today_string()

    for fmt in ("%m-%d-%Y", "%m/%d/%Y", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            parsed = datetime.strptime(value, fmt).date()
            return f"{parsed.month}-{parsed.day}-{parsed.year}"
        except ValueError:
            pass
    return value


def coerce_number(value: str) -> Any:
    text = str(value).strip()
    if not text:
        return ""
    if text.startswith("="):
        return text
    try:
        number = float(text.replace(",", ""))
    except ValueError:
        return text
    if number.is_integer():
        return int(number)
    return number


def format_product_details(data: dict[str, str], products: list[ProductItem] | None = None) -> str:
    product_items = products or [
        ProductItem(
            name=data.get("product_name", "").strip(),
            dim1=data.get("dim1", "").strip(),
            dim2=data.get("dim2", "").strip(),
            dim3=data.get("dim3", "").strip(),
            dim_unit=data.get("dim_unit", "mm").strip() or "mm",
            weight=data.get("weight", "").strip(),
            weight_unit=data.get("weight_unit", "kg").strip() or "kg",
        )
    ]

    blocks = [_format_one_product(product) for product in product_items if product.name.strip()]
    if not blocks:
        return "Product Name: \nDimensions: \nWeight: "
    return "\n\n".join(blocks)


def _format_one_product(product: ProductItem) -> str:
    dimensions = "x".join(part.strip() for part in (product.dim1, product.dim2, product.dim3) if part.strip())
    if dimensions:
        dimensions = f"{dimensions} {(product.dim_unit or 'mm').strip()}"

    weight = product.weight.strip()
    weight_text = f"{weight} {(product.weight_unit or 'kg').strip()}" if weight else ""
    return (
        f"Product Name: {product.name.strip()}\n"
        f"Dimensions: {dimensions}\n"
        f"Weight: {weight_text}"
    )


def format_customer_details(data: dict[str, str]) -> str:
    return "\n".join(
        [
            f"Customer: {data.get('customer', '').strip()}",
            f"Designation: {data.get('designation', '').strip()}",
            f"Company: {data.get('company', '').strip()}",
            f"Tel No.: {data.get('tel_no', '').strip()}",
            f"E-mail: {data.get('email', '').strip()}",
            f"Address: {data.get('address', '').strip()}",
            "",
        ]
    )


def generate_quote(data: dict[str, str], tests: list[TestItem], products: list[ProductItem] | None = None) -> Path:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
    if not tests:
        raise ValueError("Add at least one test before generating a quote.")

    OUTPUT_DIR.mkdir(exist_ok=True)
    quote_ref = reserve_next_quote_ref()
    output_path = OUTPUT_DIR / f"{quote_ref}.xlsx"
    shutil.copy2(TEMPLATE_PATH, output_path)

    workbook = load_workbook(output_path)
    worksheet = workbook.active
    worksheet.title = quote_ref

    worksheet["AJ3"] = quote_ref
    worksheet["AJ4"] = "Dated"
    worksheet["AR4"] = parse_date_text(data.get("quote_date", today_string()))
    product_details = format_product_details(data, products)
    worksheet["L6"] = product_details
    _fit_product_rows(worksheet, product_details)
    worksheet["AJ6"] = format_customer_details(data)
    worksheet["X10"] = parse_date_text(data.get("email_date", today_string()))
    worksheet["AI13"] = "Qty"
    worksheet["AO13"] = "Batch"

    _prepare_test_rows(worksheet, len(tests))
    _write_tests(worksheet, tests)
    _write_totals(worksheet, tests)

    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(output_path)
    return output_path


def _fit_product_rows(worksheet, product_details: str) -> None:
    line_count = max(3, product_details.count("\n") + 1)
    total_height = min(409, max(78, line_count * 15))
    worksheet.row_dimensions[6].height = total_height / 2
    worksheet.row_dimensions[7].height = total_height / 2


def _prepare_test_rows(worksheet, test_count: int) -> None:
    rows_to_insert = max(test_count - 1, 0)
    if rows_to_insert:
        shifted_merges = _remove_merges_for_insert(worksheet, TOTAL_START_ROW, rows_to_insert)
        worksheet.insert_rows(TOTAL_START_ROW, rows_to_insert)
        for merge_range in shifted_merges:
            worksheet.merge_cells(str(merge_range))

    for offset in range(test_count):
        row = TEST_START_ROW + offset
        _copy_row_style(worksheet, TEST_START_ROW, row)
        _ensure_test_row_merges(worksheet, row)
        worksheet.row_dimensions[row].height = max(worksheet.row_dimensions[row].height or 48, 48)


def _remove_merges_for_insert(worksheet, insert_at: int, amount: int) -> list[CellRange]:
    shifted = []
    for merge_range in list(worksheet.merged_cells.ranges):
        cell_range = CellRange(str(merge_range))
        if cell_range.min_row >= insert_at:
            worksheet.unmerge_cells(str(cell_range))
            cell_range.shift(row_shift=amount, col_shift=0)
            shifted.append(cell_range)
    return shifted


def _copy_row_style(worksheet, source_row: int, target_row: int) -> None:
    if source_row == target_row:
        return
    worksheet.row_dimensions[target_row].height = worksheet.row_dimensions[source_row].height
    for col in range(1, worksheet.max_column + 1):
        source = worksheet.cell(source_row, col)
        target = worksheet.cell(target_row, col)
        if source.has_style:
            target._style = copy(source._style)
        if source.number_format:
            target.number_format = source.number_format
        if source.alignment:
            target.alignment = copy(source.alignment)
        if source.font:
            target.font = copy(source.font)
        if source.fill:
            target.fill = copy(source.fill)
        if source.border:
            target.border = copy(source.border)
        if source.protection:
            target.protection = copy(source.protection)


def _ensure_test_row_merges(worksheet, row: int) -> None:
    for start_col, end_col in TEST_MERGES:
        start = f"{get_column_letter(start_col)}{row}"
        end = f"{get_column_letter(end_col)}{row}"
        cell_range = f"{start}:{end}"
        if cell_range not in {str(rng) for rng in worksheet.merged_cells.ranges}:
            worksheet.merge_cells(cell_range)


def _write_tests(worksheet, tests: list[TestItem]) -> None:
    for index, test in enumerate(tests, start=1):
        row = TEST_START_ROW + index - 1
        worksheet[f"B{row}"] = index
        worksheet[f"D{row}"] = (
            f"Type of Test: {test.name.strip()}\n"
            f"Test Method/Requirement:\n{test.requirements.strip()}"
        )
        worksheet[f"AE{row}"] = SAC_CODE
        worksheet[f"AI{row}"] = coerce_number(test.qty)
        worksheet[f"AO{row}"] = coerce_number(test.batch)
        worksheet[f"AT{row}"] = coerce_number(test.total_cost)

        line_count = worksheet[f"D{row}"].value.count("\n") + 1
        worksheet.row_dimensions[row].height = max(48, min(180, 16 * line_count))


def _write_totals(worksheet, tests: list[TestItem]) -> None:
    last_test_row = TEST_START_ROW + len(tests) - 1
    amount_row = last_test_row + 1
    total_row = amount_row
    cgst_row = total_row + 1
    sgst_row = total_row + 2
    igst_row = total_row + 3
    grand_total_row = total_row + 4

    worksheet[f"B{amount_row}"] = "Amount in Words"
    worksheet[f"H{amount_row}"] = amount_in_words(_estimated_grand_total(tests))
    worksheet[f"AE{total_row}"] = "Total"
    worksheet[f"AN{total_row}"] = f"=SUM(AT{TEST_START_ROW}:AX{last_test_row})"
    worksheet[f"AE{cgst_row}"] = "CGST 9%"
    worksheet[f"AN{cgst_row}"] = f"=9%*AN{total_row}"
    worksheet[f"AE{sgst_row}"] = "SGST 9%"
    worksheet[f"AN{sgst_row}"] = f"=9%*AN{total_row}"
    worksheet[f"AE{igst_row}"] = "IGST 18%"
    worksheet[f"AN{igst_row}"] = 0
    worksheet[f"AE{grand_total_row}"] = "Grand Total"
    worksheet[f"AN{grand_total_row}"] = f"=SUM(AN{total_row}:AX{igst_row})"


def _estimated_grand_total(tests: list[TestItem]) -> float | None:
    total = 0.0
    found_any = False
    for test in tests:
        value = str(test.total_cost).strip().replace(",", "")
        if not value or value.startswith("="):
            continue
        try:
            total += float(value)
            found_any = True
        except ValueError:
            pass
    if not found_any:
        return None
    return round(total * 1.18)


ONES = [
    "",
    "one",
    "two",
    "three",
    "four",
    "five",
    "six",
    "seven",
    "eight",
    "nine",
    "ten",
    "eleven",
    "twelve",
    "thirteen",
    "fourteen",
    "fifteen",
    "sixteen",
    "seventeen",
    "eighteen",
    "nineteen",
]
TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]


def amount_in_words(amount: float | None) -> str:
    if amount is None:
        return ""
    number = int(round(amount))
    if number <= 0:
        return "zero rupees only"
    words = _indian_number_words(number)
    return f"{words} rupees only".capitalize()


def _under_thousand(value: int) -> str:
    parts = []
    hundreds, remainder = divmod(value, 100)
    if hundreds:
        parts.append(f"{ONES[hundreds]} hundred")
    if remainder:
        if remainder < 20:
            parts.append(ONES[remainder])
        else:
            tens, ones = divmod(remainder, 10)
            parts.append(TENS[tens] + (f" {ONES[ones]}" if ones else ""))
    return " ".join(parts)


def _indian_number_words(value: int) -> str:
    crore, value = divmod(value, 10_000_000)
    lakh, value = divmod(value, 100_000)
    thousand, value = divmod(value, 1_000)
    parts = []
    if crore:
        parts.append(f"{_under_thousand(crore)} crore")
    if lakh:
        parts.append(f"{_under_thousand(lakh)} lakh")
    if thousand:
        parts.append(f"{_under_thousand(thousand)} thousand")
    if value:
        parts.append(_under_thousand(value))
    return re.sub(r"\s+", " ", " ".join(parts)).strip()
