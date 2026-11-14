from PySide6.QtWidgets import QFrame, QVBoxLayout, QLabel


class StatCard(QFrame):
    def __init__(self, titulo: str, valor: str = "0", parent=None):
        super().__init__(parent)
        self.setObjectName("StatCard")

        self._titulo = QLabel(titulo)
        self._titulo.setObjectName("StatTitle")
        self._valor = QLabel(valor)
        self._valor.setObjectName("StatValue")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(8)
        layout.addWidget(self._titulo)
        layout.addWidget(self._valor)
        self.setMinimumHeight(98)

    def set_valor(self, valor: str):
        self._valor.setText(valor)