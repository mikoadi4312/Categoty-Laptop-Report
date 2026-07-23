import unittest
from datetime import date
from pathlib import Path

import pandas as pd

from etl import upload_stock


class PrepareStockTest(unittest.TestCase):
    def test_brand_extraction_and_stock_statuses(self) -> None:
        source = pd.DataFrame(
            [
                self.stock_row("Macbook Air 13 M5", "1-New", 2),
                self.stock_row("Laptop Acer Test", "3-Show", 3),
                self.stock_row("Unidentified notebook", "5-Error (New)", 4),
            ]
        )
        original_reader = upload_stock.read_input_file
        upload_stock.read_input_file = lambda _path: source.copy()
        try:
            result = upload_stock.prepare_stock(Path("stock.xlsx"), date(2026, 7, 23))
        finally:
            upload_stock.read_input_file = original_reader

        by_brand = result.set_index("brand")
        self.assertEqual(int(by_brand.loc["Apple", "new_stock"]), 2)
        self.assertEqual(int(by_brand.loc["Acer", "demo_units"]), 3)
        self.assertEqual(int(by_brand.loc["UNKNOWN", "new_stock"]), 4)

    def test_stock_date_is_detected_from_filename(self) -> None:
        fallback = date(2026, 7, 1)
        detected = upload_stock.stock_date_from_filename(
            "GeneralInventory_20260723_160024.xlsx",
            fallback,
        )
        self.assertEqual(detected, date(2026, 7, 23))
        self.assertEqual(upload_stock.stock_date_from_filename("GeneralInventory.xlsx", fallback), fallback)

    @staticmethod
    def stock_row(product_name: str, status: str, quantity: int) -> dict:
        return {
            "ID Store": 1,
            "Store Name": "Test Store",
            "Main Category": "1581 - INDO - Laptop",
            "Sub Category": "5652 - INDO - Laptop",
            "ID Model": "MODEL-1",
            "Product code": "SKU-1",
            "Product name": product_name,
            "Inventory Status": status,
            "Quantity_1": quantity,
            "QUANTITYEX": 0,
        }


if __name__ == "__main__":
    unittest.main()
