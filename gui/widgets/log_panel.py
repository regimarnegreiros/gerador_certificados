"""
Painel compacto com barra de progresso, botão cancelar e registro recolhível de logs.
"""

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
)


_PREFIXOS_NIVEL = {
    "info": "[INFO]",
    "sucesso": "[OK]",
    "aviso": "[AVISO]",
    "erro": "[ERRO]",
}


class LogPanelWidget(QFrame):
    """Barra de status/progresso com log expansível para acompanhamento de tarefas."""

    cancelar_solicitado = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setProperty("class", "Card")
        self.setObjectName("LogPanel")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(6)

        barra_top = QHBoxLayout()
        barra_top.setSpacing(8)

        self.lbl_status = QLabel("Pronto.")
        self.lbl_status.setStyleSheet("font-weight: 600;")

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedWidth(220)

        self.btn_cancelar = QPushButton("Cancelar")
        self.btn_cancelar.setObjectName("DangerButton")
        self.btn_cancelar.setVisible(False)
        self.btn_cancelar.clicked.connect(self.cancelar_solicitado.emit)

        self.btn_toggle_log = QPushButton("Exibir detalhes")
        self.btn_toggle_log.setCheckable(True)
        self.btn_toggle_log.toggled.connect(self._alternar_logs)

        barra_top.addWidget(self.lbl_status, 1)
        barra_top.addWidget(self.progress_bar)
        barra_top.addWidget(self.btn_cancelar)
        barra_top.addWidget(self.btn_toggle_log)
        layout.addLayout(barra_top)

        self.txt_logs = QPlainTextEdit()
        self.txt_logs.setReadOnly(True)
        self.txt_logs.setMaximumHeight(130)
        self.txt_logs.setVisible(False)
        self.txt_logs.setStyleSheet("font-family: Consolas, 'Courier New', monospace; font-size: 8.5pt;")
        layout.addWidget(self.txt_logs)

    def _alternar_logs(self, visivel: bool) -> None:
        self.txt_logs.setVisible(visivel)
        self.btn_toggle_log.setText("Ocultar detalhes" if visivel else "Exibir detalhes")

    def iniciar_tarefa(self, mensagem: str, cancelavel: bool = True) -> None:
        self.lbl_status.setText(mensagem)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.btn_cancelar.setEnabled(True)
        self.btn_cancelar.setVisible(cancelavel)

    def atualizar_progresso(self, atual: int, total: int, descricao: str) -> None:
        if total > 0:
            pct = max(0, min(100, int((atual / total) * 100)))
            self.progress_bar.setRange(0, 100)
            self.progress_bar.setValue(pct)
        else:
            self.progress_bar.setRange(0, 0)
        if descricao:
            self.lbl_status.setText(descricao)

    def registrar_log(self, nivel: str, mensagem: str) -> None:
        prefixo = _PREFIXOS_NIVEL.get(nivel.lower(), "[LOG]")
        self.txt_logs.appendPlainText(f"{prefixo} {mensagem}")

    def finalizar_tarefa(self, mensagem: str, erro: bool = False) -> None:
        self.lbl_status.setText(mensagem)
        self.btn_cancelar.setVisible(False)
        self.progress_bar.setRange(0, 100)
        if erro:
            self.progress_bar.setVisible(False)
            if not self.txt_logs.isVisible():
                self.btn_toggle_log.setChecked(True)
        else:
            self.progress_bar.setValue(100)

    def limpar_logs(self) -> None:
        self.txt_logs.clear()
        self.lbl_status.setText("Pronto.")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        self.btn_cancelar.setVisible(False)
