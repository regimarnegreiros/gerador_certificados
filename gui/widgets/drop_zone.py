"""
Componente reutilizável de Arrastar e Soltar arquivos (com clique para seletor nativo).
"""

from pathlib import Path
from typing import Iterable, List, Optional, Tuple

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QDragEnterEvent, QDragLeaveEvent, QDropEvent
from PySide6.QtWidgets import QFileDialog, QFrame, QHBoxLayout, QLabel, QPushButton, QVBoxLayout


class DropZoneWidget(QFrame):
    """Área de arrastar-e-soltar arquivos que também abre seletor ao clicar."""

    arquivos_selecionados = Signal(list)
    limpar_solicitado = Signal()

    def __init__(
        self,
        titulo: str,
        extensoes: Iterable[str],
        dica: str = "Arraste o arquivo aqui ou clique para selecionar",
        multiplos: bool = False,
        compacto: bool = False,
        permitir_limpar: bool = False,
        parent=None,
    ):
        super().__init__(parent)
        self.setObjectName("DropZone")
        self.setAcceptDrops(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

        self._titulo_base = titulo
        self._dica_base = dica
        self._extensoes: Tuple[str, ...] = tuple(e.lower() if e.startswith(".") else f".{e.lower()}" for e in extensoes)
        self._multiplos = multiplos
        self._compacto = compacto

        self.setProperty("state", "empty")
        self.setProperty("dragging", "false")

        self._layout = QVBoxLayout(self) if compacto else QHBoxLayout(self)
        margem = 8 if compacto else 12
        self._layout.setContentsMargins(margem, margem, margem, margem)
        self._layout.setSpacing(6 if compacto else 10)

        textos_box = QVBoxLayout()
        textos_box.setSpacing(2)

        self.lbl_titulo = QLabel(titulo)
        self.lbl_titulo.setStyleSheet("font-weight: 700; background: transparent; border: none;")
        self.lbl_titulo.setWordWrap(True)

        self.lbl_status = QLabel(dica)
        self.lbl_status.setObjectName("DropStatus")
        self.lbl_status.setProperty("class", "MutedText")
        self.lbl_status.setStyleSheet("color: #5A6B7F; font-size: 8.5pt; background: transparent; border: none;")
        self.lbl_status.setWordWrap(True)

        textos_box.addWidget(self.lbl_titulo)
        textos_box.addWidget(self.lbl_status)

        if compacto:
            self._layout.addLayout(textos_box)
        else:
            self._layout.addLayout(textos_box, 1)

        acoes_box = QHBoxLayout()
        acoes_box.setSpacing(6)

        self.btn_selecionar = QPushButton("Selecionar..." if not compacto else "Alterar...")
        self.btn_selecionar.clicked.connect(self.abrir_seletor)
        acoes_box.addWidget(self.btn_selecionar)

        self.btn_limpar: Optional[QPushButton] = None
        if permitir_limpar:
            self.btn_limpar = QPushButton("Remover")
            self.btn_limpar.setObjectName("DangerButton")
            self.btn_limpar.setVisible(False)
            self.btn_limpar.clicked.connect(self.limpar_solicitado.emit)
            acoes_box.addWidget(self.btn_limpar)

        self._layout.addLayout(acoes_box)

    def _atualizar_estilo(self) -> None:
        self.style().unpolish(self)
        self.style().polish(self)

    def definir_carregado(self, texto_status: str) -> None:
        """Atualiza visual para estado carregado (verde) com resumo do arquivo."""
        self.setProperty("state", "loaded")
        self.lbl_status.setText(texto_status)
        self.lbl_status.setStyleSheet("color: #168821; font-size: 8.5pt; font-weight: 600; background: transparent; border: none;")
        if self.btn_limpar is not None:
            self.btn_limpar.setVisible(True)
        self._atualizar_estilo()

    def definir_pendente(self, texto_status: Optional[str] = None) -> None:
        """Restaura visual para estado pendente."""
        self.setProperty("state", "empty")
        self.lbl_status.setText(texto_status or self._dica_base)
        self.lbl_status.setStyleSheet("color: #5A6B7F; font-size: 8.5pt; background: transparent; border: none;")
        if self.btn_limpar is not None:
            self.btn_limpar.setVisible(False)
        self._atualizar_estilo()

    def _filtrar_caminhos_validos(self, urls) -> List[Path]:
        validos: List[Path] = []
        for url in urls:
            if not url.isLocalFile():
                continue
            p = Path(url.toLocalFile())
            if p.is_file() and not p.name.startswith("~$") and p.suffix.lower() in self._extensoes:
                validos.append(p)
        if not self._multiplos and len(validos) > 1:
            return [validos[0]]
        return validos

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if event.mimeData().hasUrls():
            validos = self._filtrar_caminhos_validos(event.mimeData().urls())
            if validos:
                self.setProperty("dragging", "true")
                self._atualizar_estilo()
                event.acceptProposedAction()
                return
        event.ignore()

    def dragLeaveEvent(self, event: QDragLeaveEvent) -> None:
        self.setProperty("dragging", "false")
        self._atualizar_estilo()
        super().dragLeaveEvent(event)

    def dropEvent(self, event: QDropEvent) -> None:
        self.setProperty("dragging", "false")
        self._atualizar_estilo()
        if event.mimeData().hasUrls():
            validos = self._filtrar_caminhos_validos(event.mimeData().urls())
            if validos:
                event.acceptProposedAction()
                self.arquivos_selecionados.emit([str(p) for p in validos])
                return
        event.ignore()

    def mousePressEvent(self, event) -> None:
        if event.button() == Qt.MouseButton.LeftButton:
            self.abrir_seletor()
        super().mousePressEvent(event)

    def abrir_seletor(self) -> None:
        ext_filtro = " ".join(f"*{e}" for e in self._extensoes)
        filtro = f"Arquivos suportados ({ext_filtro})"
        if self._multiplos:
            caminhos, _ = QFileDialog.getOpenFileNames(self, self._titulo_base, "", filtro)
            if caminhos:
                self.arquivos_selecionados.emit(caminhos)
        else:
            caminho, _ = QFileDialog.getOpenFileName(self, self._titulo_base, "", filtro)
            if caminho:
                self.arquivos_selecionados.emit([caminho])
