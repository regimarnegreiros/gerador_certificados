"""
Ponto de entrada principal do Gerador de Certificados ANM (DIVCAM / SGP).
Configura renderização em software antes da inicialização do Qt para máxima
compatibilidade em estações corporativas e sessões remotas (RDP/VDI).
"""

import argparse
import os
import sys
import time
from pathlib import Path

# Deve ser definido antes da criação do QApplication
os.environ.setdefault("QT_OPENGL", "software")

from PySide6.QtCore import QTimer
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


def _executar_smoke_test(app: QApplication, janela: MainWindow, t_inicio: float) -> int:
    """Valida abertura da GUI, drag-and-drop e geração sem descompactar em %TEMP%."""
    import json
    from PySide6.QtCore import QMimeData, QPointF, Qt, QUrl
    from PySide6.QtGui import QDropEvent
    from PySide6.QtWidgets import QMessageBox

    resultado = {
        "frozen": bool(getattr(sys, "frozen", False)),
        "tabs": janela.tabs.count(),
        "raiz_dados": str(janela.raiz_dados),
        "startup_ms": None,
        "relatorio_gerado": False,
        "meipass": getattr(sys, "_MEIPASS", ""),
        "erro": "",
    }

    def _gravar_e_sair(codigo: int = 0) -> None:
        (janela.raiz_dados / "smoke_result.json").write_text(
            json.dumps(resultado, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(json.dumps(resultado, ensure_ascii=False))
        janela.close()
        app.quit()
        app.setProperty("_smoke_rc", codigo)

    def _capturar_alerta(_parent, titulo, texto, *_args, **_kwargs):
        resultado["erro"] = f"{titulo}: {texto}"
        return QMessageBox.StandardButton.Ok

    QMessageBox.warning = _capturar_alerta  # type: ignore[method-assign]
    QMessageBox.critical = _capturar_alerta  # type: ignore[method-assign]

    def _iniciar_apos_abertura() -> None:
        resultado["startup_ms"] = int((time.perf_counter() - t_inicio) * 1000)
        pasta_in = janela.raiz_dados / "_smoke_input"
        if not pasta_in.exists():
            _gravar_e_sair(0)
            return

        try:
            def _drop(widget, arquivos):
                mime = QMimeData()
                mime.setUrls([QUrl.fromLocalFile(str(p)) for p in arquivos])
                widget.dropEvent(
                    QDropEvent(
                        QPointF(8, 8),
                        Qt.DropAction.CopyAction,
                        mime,
                        Qt.MouseButton.LeftButton,
                        Qt.KeyboardModifier.NoModifier,
                    )
                )

            _drop(janela.drop_servidores, [pasta_in / "servidores.xlsx"])
            if (pasta_in / "modelo.pptx").exists():
                _drop(janela.drop_modelo, [pasta_in / "modelo.pptx"])
            janela.acao_novo_curso("Curso Smoke EXE")
            _drop(janela.tab_arquivos.drop_convidados, [pasta_in / "convocados.csv"])
            _drop(janela.tab_arquivos.drop_presencas, [pasta_in / "dia_1.csv"])

            def _ao_terminar_worker() -> None:
                from core.app_paths import obter_pasta_saida_curso
                from gui.file_actions import localizar_primeiro_arquivo

                if janela.curso_atual_pasta is not None:
                    rel = localizar_primeiro_arquivo(
                        obter_pasta_saida_curso(janela.curso_atual_pasta) / "relatorio_presencas",
                        (".xlsx",),
                    )
                    resultado["relatorio_gerado"] = rel is not None and rel.exists()
                _gravar_e_sair(0 if resultado["relatorio_gerado"] else 1)

            janela.tab_relatorio.acao_gerar_relatorio()
            if janela._worker_ativo is not None:
                janela._worker_ativo.finished.connect(_ao_terminar_worker)
            else:
                _gravar_e_sair(1)
        except Exception as exc:
            resultado["erro"] = str(exc)
            _gravar_e_sair(1)

    QTimer.singleShot(50, _iniciar_apos_abertura)
    rc = app.exec()
    return int(app.property("_smoke_rc") or rc)


def main(argv=None) -> int:
    t_inicio = time.perf_counter()
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--raiz-dados", default=None)
    parser.add_argument("--smoke-test", action="store_true")
    args, qt_args = parser.parse_known_args(argv)

    app = QApplication([sys.argv[0], *qt_args])
    app.setApplicationName("Gerador de Certificados ANM")
    app.setOrganizationName("ANM - DIVCAM/SGP")
    app.setFont(QFont("Segoe UI", 10))

    aplicar_tema_anm(app)

    icone_path = _resolver_caminho_recurso("assets/anm.ico")
    if icone_path.exists():
        app.setWindowIcon(QIcon(str(icone_path)))

    janela = MainWindow(raiz_dados=args.raiz_dados)
    janela.show()

    if args.smoke_test:
        return _executar_smoke_test(app, janela, t_inicio)
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
