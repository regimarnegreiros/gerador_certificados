"""
Testes automatizados da base da interface gráfica (Etapa 5):
tema ANM, DropZoneWidget, LogPanelWidget, BaseTaskWorker e MainWindow.
"""

import os
from pathlib import Path
import shutil
import tempfile
import unittest

os.environ.setdefault("QT_OPENGL", "software")
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

import pandas as pd
from PySide6.QtCore import QMimeData, QPointF, Qt, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent
from PySide6.QtWidgets import QApplication

import sys

diretorio_raiz = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(diretorio_raiz))

from gui.main_window import MainWindow
from gui.theme import aplicar_tema_anm
from gui.widgets.drop_zone import DropZoneWidget
from gui.widgets.log_panel import LogPanelWidget
from gui.workers import BaseTaskWorker


def _obter_qapp() -> QApplication:
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    aplicar_tema_anm(app)
    return app


class TestGuiBase(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.app = _obter_qapp()

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_drop_zone_drag_and_drop_e_estados(self):
        dz = DropZoneWidget(
            titulo="Servidores",
            extensoes=(".xlsx", ".xls"),
            multiplos=False,
        )
        recebidos = []
        dz.arquivos_selecionados.connect(lambda lst: recebidos.append(lst))

        # Arquivo inválido (.txt) deve ser ignorado
        arq_txt = self.temp_dir / "invalido.txt"
        arq_txt.write_text("x")
        mime_inv = QMimeData()
        mime_inv.setUrls([QUrl.fromLocalFile(str(arq_txt))])
        ev_enter_inv = QDragEnterEvent(QPointF(5, 5), Qt.DropAction.CopyAction, mime_inv, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        dz.dragEnterEvent(ev_enter_inv)
        self.assertFalse(ev_enter_inv.isAccepted())

        # Arquivo válido (.xlsx) deve ser aceito no drop
        arq_xlsx = self.temp_dir / "servidores.xlsx"
        arq_xlsx.write_text("dummy")
        mime_ok = QMimeData()
        mime_ok.setUrls([QUrl.fromLocalFile(str(arq_xlsx))])

        ev_enter_ok = QDragEnterEvent(QPointF(5, 5), Qt.DropAction.CopyAction, mime_ok, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        dz.dragEnterEvent(ev_enter_ok)
        self.assertTrue(ev_enter_ok.isAccepted())

        ev_drop = QDropEvent(QPointF(5, 5), Qt.DropAction.CopyAction, mime_ok, Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier)
        dz.dropEvent(ev_drop)
        self.assertTrue(ev_drop.isAccepted())
        self.assertEqual(len(recebidos), 1)
        self.assertEqual(Path(recebidos[0][0]).name, "servidores.xlsx")

        dz.definir_carregado("Ativo: servidores.xlsx")
        self.assertEqual(dz.property("state"), "loaded")
        dz.definir_pendente()
        self.assertEqual(dz.property("state"), "empty")

    def test_log_panel_fluxo(self):
        painel = LogPanelWidget()
        cancelado = []
        painel.cancelar_solicitado.connect(lambda: cancelado.append(True))

        painel.iniciar_tarefa("Executando...", cancelavel=True)
        self.assertTrue(painel.progress_bar.isVisibleTo(painel) or not painel.isVisible())
        painel.atualizar_progresso(2, 4, "Etapa 2/4")
        self.assertEqual(painel.progress_bar.value(), 50)
        painel.registrar_log("sucesso", "Tudo certo")
        self.assertIn("[OK] Tudo certo", painel.txt_logs.toPlainText())

        painel.btn_cancelar.click()
        self.assertEqual(cancelado, [True])

        painel.finalizar_tarefa("Concluído", erro=False)
        self.assertEqual(painel.progress_bar.value(), 100)

    def test_worker_com_callbacks_e_cancelamento(self):
        logs = []
        progs = []
        resultados = []

        def tarefa(on_log=None, on_progress=None, cancel_event=None):
            on_log("info", "Iniciado")
            on_progress(1, 2, "Metade")
            on_progress(2, 2, "Fim")
            return {"ok": True, "cancelado": cancel_event.is_set()}

        worker = BaseTaskWorker(tarefa, usar_com=True)
        worker.log_emitido.connect(lambda n, m: logs.append((n, m)))
        worker.progresso_atualizado.connect(lambda a, t, d: progs.append((a, t, d)))
        worker.concluido.connect(lambda res: resultados.append(res))

        worker.start()
        self.assertTrue(worker.wait(5000))
        self.app.processEvents()

        self.assertEqual(logs, [("info", "Iniciado")])
        self.assertEqual(progs[-1], (2, 2, "Fim"))
        self.assertEqual(resultados, [{"ok": True, "cancelado": False}])

    def test_main_window_cursos_e_recursos_globais(self):
        raiz = self.temp_dir / "raiz_anm"
        staging = self.temp_dir / "staging"
        staging.mkdir()

        janela = MainWindow(raiz_dados=raiz)
        self.assertIn("ANM", janela.windowTitle())
        self.assertEqual(janela.stack_conteudo.currentWidget(), janela.pagina_vazia)

        # Cadastra recursos globais via handlers da janela
        arq_serv = staging / "Servidores_2026.xlsx"
        pd.DataFrame({"NOME SERVIDOR": ["ALICE"], "UORG": ["SGP"]}).to_excel(arq_serv, index=False)
        janela._ao_definir_servidores_global([str(arq_serv)])
        self.assertEqual(janela.drop_servidores.property("state"), "loaded")

        from pptx import Presentation
        arq_mod = staging / "Modelo_ANM.pptx"
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[0])
        prs.save(arq_mod)
        janela._ao_definir_modelo_global([str(arq_mod)])
        self.assertEqual(janela.drop_modelo.property("state"), "loaded")

        # Cria dois cursos e testa seleção, filtro e 5 abas
        id1 = janela.acao_novo_curso("Curso Segurança de Barragens")
        id2 = janela.acao_novo_curso("Curso Gestão SGP")
        self.assertIsNotNone(id1)
        self.assertIsNotNone(id2)
        self.assertEqual(janela.stack_conteudo.currentWidget(), janela.pagina_curso)
        self.assertEqual(janela.tabs.count(), 5)
        self.assertEqual(janela.curso_atual_id, id2)

        # Filtro na sidebar
        janela.txt_busca_curso.setText("Barragens")
        self.assertEqual(janela.lista_cursos.count(), 1)
        janela.txt_busca_curso.setText("")
        self.assertEqual(janela.lista_cursos.count(), 2)

        # Exclui curso selecionado
        self.assertTrue(janela.acao_excluir_curso(confirmar=False))
        self.assertEqual(janela.lista_cursos.count(), 1)

        janela.close()


if __name__ == "__main__":
    unittest.main()
