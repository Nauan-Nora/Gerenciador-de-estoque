from decimal import Decimal

from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QSpinBox,
    QVBoxLayout,
)

from models import MovementType, Product


class StockMovementDialog(QDialog):
    def __init__(self, product: Product, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Movimentar estoque")
        self.setMinimumWidth(390)

        self.movement_type_input = QComboBox()
        self.movement_type_input.addItem("Entrada", "ENTRADA")
        self.movement_type_input.addItem("Saída por venda", "SAIDA")
        self.movement_type_input.addItem("Saída por perda/uso", "PERDA")
        self.quantity_input = QSpinBox()
        self.quantity_input.setRange(1, 2_147_483_647)
        self.quantity_input.setValue(1)
        self.purchase_cost_input = QDoubleSpinBox()
        self.purchase_cost_input.setRange(0, 9_999_999_999.99)
        self.purchase_cost_input.setDecimals(2)
        self.purchase_cost_input.setValue(float(product.purchase_price))
        self.purchase_cost_label = QLabel("Custo unitário (R$):")
        self.purchase_cost_input.setVisible(False)
        self.purchase_cost_label.setVisible(False)
        self.notes_input = QLineEdit()
        self.notes_input.setMaxLength(500)

        product_label = QLabel(
            f"{product.name} ({product.code}) — estoque atual: {product.quantity}"
        )
        product_label.setWordWrap(True)

        form = QFormLayout()
        form.addRow("Produto:", product_label)
        form.addRow("Movimentação:", self.movement_type_input)
        form.addRow("Quantidade:", self.quantity_input)
        form.addRow(self.purchase_cost_label, self.purchase_cost_input)
        form.addRow("Observação:", self.notes_input)
        self.movement_type_input.currentIndexChanged.connect(self._update_cost_visibility)
        self._update_cost_visibility()

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    @property
    def movement_type(self) -> MovementType:
        return self.movement_type_input.currentData()

    @property
    def quantity(self) -> int:
        return self.quantity_input.value()

    @property
    def notes(self) -> str:
        return self.notes_input.text().strip()

    @property
    def purchase_unit_cost(self) -> Decimal | None:
        if self.movement_type != "ENTRADA":
            return None
        return Decimal(str(self.purchase_cost_input.value()))

    def _update_cost_visibility(self):
        is_entry = self.movement_type == "ENTRADA"
        self.purchase_cost_label.setVisible(is_entry)
        self.purchase_cost_input.setVisible(is_entry)
