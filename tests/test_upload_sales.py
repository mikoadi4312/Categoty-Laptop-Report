import unittest
from pathlib import Path

import pandas as pd

from etl import upload_sales


class PrepareSalesTest(unittest.TestCase):
    def test_quantity_counts_transaction_rows(self) -> None:
        source = pd.DataFrame(
            [
                self.sales_row(quantity=2, revenue=200),
                self.sales_row(quantity=0, revenue=0, output="Warranty change"),
                self.sales_row(quantity=-1, revenue=-100),
            ]
        )
        original_reader = upload_sales.read_input_file
        upload_sales.read_input_file = lambda _path: source.copy()
        try:
            result = upload_sales.prepare_sales(Path("sales.xlsx"))
        finally:
            upload_sales.read_input_file = original_reader

        self.assertEqual(len(result), 1)
        self.assertEqual(int(result.iloc[0]["day_qty"]), 3)
        self.assertEqual(float(result.iloc[0]["day_revenue"]), 100.0)

    @staticmethod
    def sales_row(quantity: int, revenue: float, output: str = "Store sale") -> dict:
        return {
            "DATE": "23/07/2026",
            "IDStore": 1,
            "Store Name": "Test Store",
            "MainCategory": "Laptop",
            "SubCategory": "Laptop",
            "Product code": "SKU-1",
            "Product name": "Laptop Asus Test",
            "Brand name": "ASUS",
            "InventoryStatus": "New",
            "TypeofOutput": output,
            "Quantity_1": quantity,
            "REVENUE": revenue,
        }


if __name__ == "__main__":
    unittest.main()
