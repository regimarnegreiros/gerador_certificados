"""
Helpers compartilhados de abertura e exportação segura de arquivos gerados pela GUI.
"""

import os
from pathlib import Path
import subprocess
from typing import Optional

from PySide6.QtWidgets import QFileDialog, QMessageBox, QWidget

from core.caminhos import copiar_arquivo_com_fallback, existe_com_fallback
from core.validator import listar_arquivos_validos


def localizar_primeiro_arquivo(pasta: Path, extensoes: tuple) -> Optional[Path]:
    """Retorna o primeiro arquivo válido encontrado na pasta (ignorando temporários ~$)."""
    encontrados = listar_arquivos_validos(pasta, extensoes)
    return encontrados[0] if encontrados else None


def abrir_no_sistema(parent: QWidget, caminho: Path) -> None:
    """Abre um arquivo ou pasta no programa padrão do sistema operacional."""
    alvo = Path(caminho)
    if not existe_com_fallback(alvo):
        QMessageBox.warning(parent, "Arquivo não encontrado", f"O caminho não existe:\n{alvo}")
        return
    try:
        if os.name == "nt":
            os.startfile(str(alvo))  # type: ignore[attr-defined]
        else:
            subprocess.Popen(["xdg-open", str(alvo)])
    except Exception as exc:
        QMessageBox.warning(parent, "Erro ao abrir", f"Não foi possível abrir '{alvo.name}':\n{exc}")


def salvar_copia_como(parent: QWidget, origem: Path, filtro: str = "Todos os arquivos (*.*)") -> Optional[Path]:
    """Permite ao usuário salvar uma cópia do arquivo gerado fora da pasta interna do app."""
    origem_path = Path(origem)
    if not existe_com_fallback(origem_path):
        QMessageBox.warning(parent, "Arquivo não encontrado", f"Arquivo de origem não encontrado:\n{origem_path.name}")
        return None

    destino_str, _ = QFileDialog.getSaveFileName(
        parent,
        "Salvar cópia como...",
        origem_path.name,
        filtro,
    )
    if not destino_str:
        return None

    try:
        destino = copiar_arquivo_com_fallback(origem_path, Path(destino_str))
        QMessageBox.information(parent, "Cópia salva", f"Arquivo salvo com sucesso em:\n{destino}")
        return destino
    except Exception as exc:
        QMessageBox.critical(parent, "Erro ao salvar cópia", str(exc))
        return None
