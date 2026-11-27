"""
Workers assíncronos em QThread com inicialização segura de COM (pythoncom.CoInitialize)
para automação de Microsoft PowerPoint e Outlook sem travar a interface gráfica.
"""

import threading
from typing import Any, Callable, Dict

from PySide6.QtCore import QThread, Signal


class BaseTaskWorker(QThread):
    """
    Thread genérica de execução em segundo plano com suporte a:
    - Inicialização e limpeza COM por thread (pythoncom)
    - Sinais thread-safe para logs, progresso, status por item, sucesso e erro
    - Cancelamento cooperativo via threading.Event
    """

    log_emitido = Signal(str, str)
    progresso_atualizado = Signal(int, int, str)
    item_status_emitido = Signal(dict)
    concluido = Signal(object)
    erro_ocorrido = Signal(str)

    def __init__(
        self,
        funcao: Callable[..., Any],
        kwargs: Dict[str, Any] = None,
        usar_com: bool = True,
        suporta_cancelamento: bool = True,
        suporta_progresso: bool = True,
        suporta_log: bool = True,
        suporta_item_status: bool = False,
        parent=None,
    ):
        super().__init__(parent)
        self._funcao = funcao
        self._kwargs = dict(kwargs or {})
        self._usar_com = usar_com
        self._suporta_cancelamento = suporta_cancelamento
        self._suporta_progresso = suporta_progresso
        self._suporta_log = suporta_log
        self._suporta_item_status = suporta_item_status
        self.cancel_event = threading.Event()

    def solicitar_cancelamento(self) -> None:
        self.cancel_event.set()

    def run(self) -> None:
        pythoncom_mod = None
        com_inicializado = False
        if self._usar_com:
            try:
                import pythoncom

                pythoncom_mod = pythoncom
                pythoncom_mod.CoInitialize()
                com_inicializado = True
            except Exception:
                pass

        try:
            call_kwargs = dict(self._kwargs)
            if self._suporta_log:
                call_kwargs["on_log"] = lambda nivel, msg: self.log_emitido.emit(nivel, str(msg))
            if self._suporta_progresso:
                call_kwargs["on_progress"] = lambda a, t, d: self.progresso_atualizado.emit(int(a), int(t), str(d))
            if self._suporta_cancelamento:
                call_kwargs["cancel_event"] = self.cancel_event
            if self._suporta_item_status:
                call_kwargs["on_item_status"] = lambda reg: self.item_status_emitido.emit(dict(reg))

            resultado = self._funcao(**call_kwargs)
            self.concluido.emit(resultado)
        except Exception as exc:
            self.erro_ocorrido.emit(str(exc))
        finally:
            if com_inicializado and pythoncom_mod is not None:
                try:
                    pythoncom_mod.CoUninitialize()
                except Exception:
                    pass
