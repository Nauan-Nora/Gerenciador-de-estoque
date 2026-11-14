from pathlib import Path
from string import Template

from PySide6.QtWidgets import QApplication

from .themes import TEMA_CLARO, TEMA_ESCURO

QSS_PATH = Path(__file__).parent / "theme.qss"
TEMAS = {"claro": TEMA_CLARO, "escuro": TEMA_ESCURO}

tema_atual = TEMA_CLARO  # acessível para código Python (tabelas, gráficos)


def aplicar_tema(app: QApplication, nome: str = "claro"):
    global tema_atual
    tema_atual = TEMAS[nome]
    template = Template(QSS_PATH.read_text(encoding="utf-8"))
    app.setStyleSheet(template.safe_substitute(tema_atual))