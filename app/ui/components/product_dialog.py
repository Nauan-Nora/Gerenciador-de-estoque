from decimal import Decimal

from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QFormLayout,
    QLineEdit,
    QMessageBox,
    QSpinBox,
    QVBoxLayout,
)

from models import Product


class ProductDialog(QDialog):
    def __init__(self, product: Product | None = None, parent=None):
        super().__init__(parent)
        self._existing_product = product
        self.product: Product | None = None
        self.setWindowTitle("Editar produto" if product else "Novo produto")
        self.setMinimumWidth(390)

        self.code_input = QLineEdit(product.code if product else "")
        self.code_input.setMaxLength(50)
        self.name_input = QLineEdit(product.name if product else "")
        self.name_input.setMaxLength(150)

        self.purchase_price_input = self._price_input(
            product.purchase_price if product else Decimal("0.00")
        )
        self.purchase_price_input.setEnabled(product is None)
        self.quantity_input = QSpinBox()
        self.quantity_input.setRange(0, 2_147_483_647)
        self.quantity_input.setValue(product.quantity if product else 0)
        self.quantity_input.setEnabled(product is None)
        self.sale_price_input = self._price_input(
            product.sale_price if product else Decimal("0.00")
        )

        form = QFormLayout()
        form.addRow("Código:", self.code_input)
        form.addRow("Nome do produto:", self.name_input)
        purchase_price_label = (
            "Custo médio (alterado por entrada):"
            if product is not None
            else "Preço de custo (R$):"
        )
        form.addRow(purchase_price_label, self.purchase_price_input)
        quantity_label = (
            "Quantidade inicial:"
            if product is None
            else "Estoque atual (use movimentação):"
        )
        form.addRow(quantity_label, self.quantity_input)
        form.addRow("Preço de venda (R$):", self.sale_price_input)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._save)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

    @staticmethod
    def _price_input(value: Decimal):
        widget = QDoubleSpinBox()
        widget.setRange(0, 9_999_999_999.99)
        widget.setDecimals(2)
        widget.setSingleStep(1)
        widget.setValue(float(value))
        return widget

    def _save(self):
        code = self.code_input.text().strip()
        name = self.name_input.text().strip()
        if not code or not name:
            QMessageBox.warning(self, "Dados obrigatórios", "Informe o código e o nome.")
            return

        self.product = Product(
            id=self._existing_product.id if self._existing_product else None,
            code=code,
            name=name,
            purchase_price=Decimal(str(self.purchase_price_input.value())),
            quantity=self.quantity_input.value(),
            sale_price=Decimal(str(self.sale_price_input.value())),
        )
        self.accept()
