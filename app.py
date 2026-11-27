"""
Ponto de entrada principal do Gerador de Certificados ANM (DIVCAM / SGP).
Configura renderização em software antes da inicialização do Qt para máxima
compatibilidade em estações corporativas e sessões remotas (RDP/VDI).
"""

import os
import sys
from pathlib import Path

# Deve ser definido antes da criação do QApplication
os.environ.setdefault("QT_OPENGL", "software")

from PySide6.QtGui import QFont, QIcon
from PySide6.QtWidgets import QApplication

from gui.main_window import MainWindow
from gui.theme import aplicar_tema_anm


def _resolver_caminho_recurso(relativo: str) -> Path:
    if getattr(sys, "frozen", False):
        base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    else:
        base = Path(__file__).resolve().parent
    return base / relativo


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Gerador de Certificados ANM")
    app.setOrganizationName("ANM - DIVCAM/SGP")
    app.setFont(QFont("Segoe UI", 10))

    aplicar_tema_anm(app)

    icone_path = _resolver_caminho_recurso("assets/anm.ico")
    if icone_path.exists():
        app.setWindowIcon(QIcon(str(icone_path)))

    janela = MainWindow()
    janela.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
