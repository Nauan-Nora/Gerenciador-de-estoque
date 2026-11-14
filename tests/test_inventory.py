import sys
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "app"))

from database import ProductRepository
from models import Product, StockMovementError


class FakeCursor:
    def __init__(self, row=None):
        self._row = row

    def fetchone(self):
        return self._row


class FakeConnection:
    def __init__(self, quantity=10):
        self.quantity = quantity
        self.calls = []

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def execute(self, query, params=()):
        query_text = str(query)
        self.calls.append((query_text, params))
        if "SELECT code, name, quantity, purchase_price, sale_price" in query_text:
            return FakeCursor(
                {
                    "code": "P-1",
                    "name": "Produto",
                    "quantity": self.quantity,
                    "purchase_price": Decimal("2.00"),
                    "sale_price": Decimal("5.00"),
                }
            )
        if "INSERT INTO products" in query_text:
            return FakeCursor({"id": 41})
        return FakeCursor()


class InventoryMovementTests(unittest.TestCase):
    @patch("database.psycopg.connect")
    def test_receipt_recalculates_weighted_average_cost(self, connect):
        connection = FakeConnection(quantity=10)
        connect.return_value = connection

        ProductRepository("postgresql://test").register_movement(
            7, "ENTRADA", 10, purchase_unit_cost=Decimal("4.00")
        )

        self.assertEqual(connection.calls[1][1], (20, Decimal("3.00"), 7))
        self.assertEqual(connection.calls[2][1][5], Decimal("4.00"))

    @patch("database.psycopg.connect")
    def test_sale_records_price_snapshot(self, connect):
        connection = FakeConnection(quantity=10)
        connect.return_value = connection

        ProductRepository("postgresql://test").register_movement(7, "SAIDA", 3)

        self.assertEqual(connection.calls[1][1], (7, Decimal("2.00"), 7))
        self.assertEqual(connection.calls[2][1][6], Decimal("5.00"))

    @patch("database.psycopg.connect")
    def test_loss_is_not_recorded_as_a_sale(self, connect):
        connection = FakeConnection(quantity=10)
        connect.return_value = connection

        ProductRepository("postgresql://test").register_movement(7, "PERDA", 1)

        self.assertIsNone(connection.calls[2][1][6])

    @patch("database.psycopg.connect")
    def test_new_product_initial_quantity_creates_opening_movement(self, connect):
        connection = FakeConnection()
        connect.return_value = connection

        ProductRepository("postgresql://test").add_product(
            Product(
                code="P-2",
                name="Caderno",
                purchase_price=Decimal("3.00"),
                quantity=4,
                sale_price=Decimal("6.00"),
            )
        )

        self.assertEqual(connection.calls[1][1][0], 41)
        self.assertEqual(connection.calls[1][1][4], Decimal("3.00"))

    @patch("database.psycopg.connect")
    def test_outflow_cannot_make_stock_negative(self, connect):
        connection = FakeConnection(quantity=1)
        connect.return_value = connection

        with self.assertRaises(StockMovementError):
            ProductRepository("postgresql://test").register_movement(7, "SAIDA", 2)

        self.assertEqual(len(connection.calls), 1)


if __name__ == "__main__":
    unittest.main()
