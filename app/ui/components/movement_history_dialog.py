from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QHeaderView,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from models import StockMovement
from ui.formatting import format_currency


class MovementHistoryDialog(QDialog):
    def __init__(
        self,
        movements: list[StockMovement],
        product_name: str | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle(
            f"Histórico de estoque — {product_name}" if product_name else "Histórico de estoque"
        )
        self.resize(850, 420)

        self.table = QTableWidget(len(movements), 8)
        self.table.setHorizontalHeaderLabels(
            [
                "Data e hora",
                "Código",
                "Produto",
                "Tipo",
                "Quantidade",
                "Custo unitário",
                "Venda unitária",
                "Observação",
            ]
        )
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.Stretch)

        for row, movement in enumerate(movements):
            values = (
                movement.created_at.strftime("%d/%m/%Y %H:%M"),
                movement.product_code,
                movement.product_name,
                {
                    "ENTRADA": "Entrada",
                    "SAIDA": "Venda",
                    "PERDA": "Perda/uso",
                }[movement.movement_type],
                str(movement.quantity),
                (
                    format_currency(movement.unit_cost)
                    if movement.unit_cost is not None
                    else "—"
                ),
                (
                    format_currency(movement.unit_sale_price)
                    if movement.unit_sale_price is not None
                    else "—"
                ),
                movement.notes,
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column in (4, 5, 6):
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                self.table.setItem(row, column, item)

        layout = QVBoxLayout(self)
        layout.addWidget(self.table)
