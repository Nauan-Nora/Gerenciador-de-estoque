import sys
from decimal import Decimal

import psycopg
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from database import DatabaseConfigurationError, ProductRepository
from models import Product, StockMovementError
from ui.components.database_setup_dialog import DatabaseSetupDialog
from ui.components.movement_history_dialog import MovementHistoryDialog
from ui.components.product_dialog import ProductDialog
from ui.components.stat_card import StatCard
from ui.components.stock_movement_dialog import StockMovementDialog
from ui.formatting import format_currency
from ui.styles.loader import aplicar_tema


class MainWindow(QMainWindow):
    def __init__(self, repository: ProductRepository):
        super().__init__()
        self._repository = repository
        self._tema = "claro"
        self._products: list[Product] = []

        self.setWindowTitle("Gerenciador de Estoque")
        self.setMinimumSize(1180, 700)
        self._build_ui()
        self.refresh_products()

    def _build_ui(self):
        central = QWidget()
        root = QVBoxLayout(central)
        root.setContentsMargins(28, 24, 28, 24)
        root.setSpacing(18)

        header = QHBoxLayout()
        title_block = QVBoxLayout()
        title = QLabel("Gerenciador de Estoque")
        title.setObjectName("PageTitle")
        subtitle = QLabel("Acompanhe seus produtos e níveis de estoque.")
        subtitle.setObjectName("PageSubtitle")
        title_block.addWidget(title)
        title_block.addWidget(subtitle)
        header.addLayout(title_block)
        header.addStretch()
        database_button = QPushButton("Configurar banco")
        database_button.setObjectName("SecondaryButton")
        database_button.clicked.connect(self.configure_database)
        header.addWidget(database_button)
        self.theme_button = QPushButton("Tema escuro")
        self.theme_button.setObjectName("SecondaryButton")
        self.theme_button.clicked.connect(self.toggle_theme)
        header.addWidget(self.theme_button)
        root.addLayout(header)

        cards = QHBoxLayout()
        cards.setSpacing(14)
        self.total_card = StatCard("Produtos cadastrados", "0")
        self.low_stock_card = StatCard("Estoque baixo (até 5)", "0")
        self.inventory_card = StatCard("Valor em estoque", "R$ 0,00")
        self.sales_value_card = StatCard("Potencial de venda", "R$ 0,00")
        self.profit_card = StatCard("Lucro potencial", "R$ 0,00")
        self.revenue_card = StatCard("Vendas nos últimos 30 dias", "R$ 0,00")
        cards.addWidget(self.total_card)
        cards.addWidget(self.low_stock_card)
        cards.addWidget(self.inventory_card)
        cards.addWidget(self.sales_value_card)
        cards.addWidget(self.profit_card)
        cards.addWidget(self.revenue_card)
        root.addLayout(cards)

        scanner_row = QHBoxLayout()
        scanner_label = QLabel("Bipagem rápida")
        self.scanner_input = QLineEdit()
        self.scanner_input.setPlaceholderText(
            "Leia ou digite o código de barras e pressione Enter"
        )
        self.scanner_input.returnPressed.connect(self.handle_scan)
        scanner_row.addWidget(scanner_label)
        scanner_row.addWidget(self.scanner_input, 1)
        root.addLayout(scanner_row)

        controls = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Pesquisar por código ou nome")
        self.search_input.returnPressed.connect(self.refresh_products)
        controls.addWidget(self.search_input, 1)

        search_button = QPushButton("Pesquisar")
        search_button.clicked.connect(self.refresh_products)
        controls.addWidget(search_button)

        add_button = QPushButton("Novo produto")
        add_button.clicked.connect(self.add_product)
        controls.addWidget(add_button)
        root.addLayout(controls)

        self.table = QTableWidget(0, 8)
        self.table.setHorizontalHeaderLabels(
            [
                "Código",
                "Produto",
                "Custo médio",
                "Estoque",
                "Preço de venda",
                "Margem",
                "ABC",
                "Dias de estoque",
            ]
        )
        self.table.setObjectName("ProductTable")
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.verticalHeader().setVisible(False)
        header_view = self.table.horizontalHeader()
        header_view.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header_view.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header_view.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header_view.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header_view.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header_view.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header_view.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        header_view.setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)
        self.table.doubleClicked.connect(self.edit_product)
        root.addWidget(self.table, 1)

        actions = QHBoxLayout()
        self.selection_label = QLabel("Selecione um produto para editar ou excluir.")
        self.selection_label.setObjectName("PageSubtitle")
        actions.addWidget(self.selection_label)
        actions.addStretch()
        movement_button = QPushButton("Movimentar estoque")
        movement_button.clicked.connect(self.move_stock)
        actions.addWidget(movement_button)
        history_button = QPushButton("Histórico")
        history_button.setObjectName("SecondaryButton")
        history_button.clicked.connect(self.show_movement_history)
        actions.addWidget(history_button)
        edit_button = QPushButton("Editar")
        edit_button.setObjectName("SecondaryButton")
        edit_button.clicked.connect(self.edit_product)
        actions.addWidget(edit_button)
        delete_button = QPushButton("Excluir")
        delete_button.setObjectName("DangerButton")
        delete_button.clicked.connect(self.delete_product)
        actions.addWidget(delete_button)
        root.addLayout(actions)

        self.setCentralWidget(central)
        self.statusBar().showMessage("Conectado ao PostgreSQL")
        self.scanner_input.setFocus()

    def refresh_products(self):
        try:
            products = self._repository.list_products(self.search_input.text().strip())
            summary = self._repository.get_summary()
        except psycopg.Error as error:
            QMessageBox.critical(
                self,
                "Erro no banco de dados",
                f"Não foi possível carregar os produtos.\n\n{error}",
            )
            return

        self._products = products
        self.table.setRowCount(len(products))
        for row, product in enumerate(products):
            values = (
                product.code,
                product.name,
                format_currency(product.purchase_price),
                str(product.quantity),
                format_currency(product.sale_price),
                (
                    f"{(product.sale_price - product.purchase_price) * Decimal('100') / product.sale_price:.1f}%"
                    if product.sale_price > 0
                    else "—"
                ),
                product.abc_class,
                (
                    f"{product.stock_days:.1f} dias"
                    if product.stock_days is not None
                    else "Sem histórico"
                ),
            )
            for column, value in enumerate(values):
                item = QTableWidgetItem(value)
                if column in (2, 3, 4, 5, 6, 7):
                    item.setTextAlignment(
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                    )
                if column == 0:
                    item.setData(Qt.ItemDataRole.UserRole, product.id)
                self.table.setItem(row, column, item)

        self.total_card.set_valor(str(summary["total_products"]))
        self.low_stock_card.set_valor(str(summary["low_stock"]))
        self.inventory_card.set_valor(format_currency(summary["inventory_value"]))
        self.sales_value_card.set_valor(format_currency(summary["potential_sales"]))
        self.profit_card.set_valor(format_currency(summary["potential_profit"]))
        self.revenue_card.set_valor(format_currency(summary["revenue_30_days"]))
        self.statusBar().showMessage(
            f"{len(products)} produto(s) exibido(s) | "
            f"Margem potencial: {summary['margin_percent']:.1f}%"
        )

    def _selected_product(self) -> Product | None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self._products):
            return None
        return self._products[row]

    def add_product(self):
        dialog = ProductDialog(parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if dialog.product is None:
            return
        try:
            self._repository.add_product(dialog.product)
        except psycopg.Error as error:
            QMessageBox.critical(
                self,
                "Erro ao cadastrar",
                f"Não foi possível cadastrar o produto.\n\n{error}",
            )
            return
        self.refresh_products()

    def edit_product(self):
        product = self._selected_product()
        if product is None or product.id is None:
            QMessageBox.information(self, "Produto", "Selecione um produto para editar.")
            return
        dialog = ProductDialog(product=product, parent=self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        if dialog.product is None:
            return
        try:
            self._repository.update_product(product.id, dialog.product)
        except psycopg.Error as error:
            QMessageBox.critical(
                self,
                "Erro ao atualizar",
                f"Não foi possível atualizar o produto.\n\n{error}",
            )
            return
        self.refresh_products()

    def delete_product(self):
        product = self._selected_product()
        if product is None or product.id is None:
            QMessageBox.information(self, "Produto", "Selecione um produto para excluir.")
            return
        answer = QMessageBox.question(
            self,
            "Excluir produto",
            f"Deseja realmente excluir “{product.name}”?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if answer != QMessageBox.StandardButton.Yes:
            return
        try:
            self._repository.delete_product(product.id)
        except psycopg.Error as error:
            QMessageBox.critical(
                self,
                "Erro ao excluir",
                f"Não foi possível excluir o produto.\n\n{error}",
            )
            return
        self.refresh_products()

    def move_stock(self):
        product = self._selected_product()
        if product is None or product.id is None:
            QMessageBox.information(
                self, "Movimentação", "Selecione um produto para movimentar o estoque."
            )
            return
        self._move_stock_for(product)

    def handle_scan(self):
        code = self.scanner_input.text().strip()
        self.scanner_input.clear()
        if not code:
            return
        try:
            product = self._repository.get_product_by_code(code)
        except psycopg.Error as error:
            QMessageBox.critical(
                self,
                "Erro no banco de dados",
                f"Não foi possível localizar o produto.\n\n{error}",
            )
            self.scanner_input.setFocus()
            return
        if product is None:
            QMessageBox.information(
                self,
                "Produto não encontrado",
                f"Não existe produto cadastrado com o código “{code}”.",
            )
            self.scanner_input.setFocus()
            return
        self._move_stock_for(product)
        self.scanner_input.setFocus()

    def _move_stock_for(self, product: Product):
        if product.id is None:
            return
        dialog = StockMovementDialog(product, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            self._repository.register_movement(
                product.id,
                dialog.movement_type,
                dialog.quantity,
                dialog.notes,
                dialog.purchase_unit_cost,
            )
        except StockMovementError as error:
            QMessageBox.warning(self, "Movimentação não realizada", str(error))
            return
        except psycopg.Error as error:
            QMessageBox.critical(
                self,
                "Erro no banco de dados",
                f"Não foi possível registrar a movimentação.\n\n{error}",
            )
            return
        self.refresh_products()
        self.scanner_input.setFocus()

    def show_movement_history(self):
        product = self._selected_product()
        try:
            movements = self._repository.list_movements(
                product.id if product and product.id is not None else None
            )
        except psycopg.Error as error:
            QMessageBox.critical(
                self,
                "Erro no banco de dados",
                f"Não foi possível carregar o histórico.\n\n{error}",
            )
            return
        MovementHistoryDialog(
            movements,
            product.name if product else None,
            self,
        ).exec()

    def configure_database(self):
        dialog = DatabaseSetupDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted or dialog.repository is None:
            return
        self._repository = dialog.repository
        self.refresh_products()

    def toggle_theme(self):
        self._tema = "escuro" if self._tema == "claro" else "claro"
        aplicar_tema(QApplication.instance(), self._tema)
        self.theme_button.setText("Tema claro" if self._tema == "escuro" else "Tema escuro")


def main():
    app = QApplication(sys.argv)
    aplicar_tema(app, "claro")
    try:
        repository = ProductRepository.from_environment()
    except DatabaseConfigurationError:
        repository = None

    if repository is None:
        setup_dialog = DatabaseSetupDialog()
        if setup_dialog.exec() != QDialog.DialogCode.Accepted:
            return 0
        repository = setup_dialog.repository
    else:
        try:
            repository.initialize()
        except psycopg.Error as error:
            answer = QMessageBox.question(
                None,
                "Falha na conexão",
                "Não foi possível conectar ou preparar o PostgreSQL.\n\n"
                f"Detalhes: {error}\n\nDeseja abrir as configurações?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.Yes,
            )
            if answer != QMessageBox.StandardButton.Yes:
                return 1
            setup_dialog = DatabaseSetupDialog()
            if setup_dialog.exec() != QDialog.DialogCode.Accepted:
                return 0
            repository = setup_dialog.repository

    if repository is None:
        return 0

    window = MainWindow(repository)
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
