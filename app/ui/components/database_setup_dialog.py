from urllib.parse import quote

import psycopg
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QSpinBox,
    QVBoxLayout,
)

from database import ProductRepository


class DatabaseSetupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.repository: ProductRepository | None = None
        self.setWindowTitle("Configurar PostgreSQL")
        self.setMinimumWidth(420)

        explanation = QLabel(
            "Informe os dados de acesso ao PostgreSQL. "
            "Se o banco ainda não existir, o aplicativo tentará criá-lo."
        )
        explanation.setWordWrap(True)

        self.host_input = QLineEdit("localhost")
        self.port_input = QSpinBox()
        self.port_input.setRange(1, 65535)
        self.port_input.setValue(5432)
        self.database_input = QLineEdit("gerenciador_estoque")
        self.user_input = QLineEdit("postgres")
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)

        form = QFormLayout()
        form.addRow("Servidor:", self.host_input)
        form.addRow("Porta:", self.port_input)
        form.addRow("Banco:", self.database_input)
        form.addRow("Usuário:", self.user_input)
        form.addRow("Senha:", self.password_input)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save
            | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Save).setText("Conectar e salvar")
        buttons.accepted.connect(self._connect_and_save)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addWidget(explanation)
        layout.addLayout(form)
        layout.addWidget(buttons)

    def _database_url(self) -> str:
        host = self.host_input.text().strip()
        database = self.database_input.text().strip()
        username = self.user_input.text().strip()
        if not host or not database or not username:
            raise ValueError("Servidor, banco e usuário são obrigatórios.")
        if ":" in host and not host.startswith("["):
            host = f"[{host}]"

        user_part = quote(username, safe="")
        password_part = quote(self.password_input.text(), safe="")
        database_part = quote(database, safe="")
        return (
            f"postgresql://{user_part}:{password_part}@{host}:"
            f"{self.port_input.value()}/{database_part}"
        )

    def _connect_and_save(self):
        try:
            database_url = self._database_url()
        except ValueError as error:
            QMessageBox.warning(self, "Configuração inválida", str(error))
            return

        repository = ProductRepository(database_url)
        try:
            repository.initialize()
            repository.save_database_url(database_url)
        except psycopg.Error as error:
            QMessageBox.critical(
                self,
                "Falha na conexão",
                "Não foi possível conectar ou preparar o PostgreSQL. "
                "Confira servidor, porta, credenciais e permissões.\n\n"
                f"Detalhes: {error}",
            )
            return
        except OSError as error:
            QMessageBox.critical(
                self,
                "Falha ao salvar configuração",
                f"O banco foi preparado, mas não foi possível salvar o arquivo .env.\n\n{error}",
            )
            return

        self.repository = repository
        self.accept()
