from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from typing import Literal


@dataclass(frozen=True)
class Product:
    code: str
    name: str
    purchase_price: Decimal
    quantity: int
    sale_price: Decimal
    id: int | None = None
    stock_days: Decimal | None = None
    abc_class: str = "C"


MovementType = Literal["ENTRADA", "SAIDA", "PERDA"]


@dataclass(frozen=True)
class StockMovement:
    id: int
    product_id: int | None
    product_code: str
    product_name: str
    movement_type: MovementType
    quantity: int
    notes: str
    created_at: datetime
    unit_cost: Decimal | None = None
    unit_sale_price: Decimal | None = None


class StockMovementError(ValueError):
    pass
