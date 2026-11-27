"""
Helpers leves para emissão de eventos (logs, progresso e cancelamento cooperativo)
desacoplados da camada de interface gráfica.
"""

from typing import Any, Callable, Optional


LogCallback = Optional[Callable[[str, str], None]]
ProgressCallback = Optional[Callable[[int, int, str], None]]


def emitir_log(on_log: LogCallback, nivel: str, mensagem: str) -> None:
    """Envia mensagem estruturada para callback de log, quando fornecido."""
    if on_log is not None:
        on_log(nivel, mensagem)


def emitir_progresso(
    on_progress: ProgressCallback,
    atual: int,
    total: int,
    descricao: str,
) -> None:
    """Envia atualização de progresso (atual/total) para callback, quando fornecido."""
    if on_progress is not None:
        on_progress(atual, total, descricao)


def cancelado(cancel_event: Optional[Any]) -> bool:
    """Verifica de forma segura se um cancelamento foi solicitado (ex: threading.Event)."""
    return bool(cancel_event is not None and callable(getattr(cancel_event, "is_set", None)) and cancel_event.is_set())
