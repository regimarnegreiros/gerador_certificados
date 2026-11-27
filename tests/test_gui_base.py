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

    def test_tab_arquivos_e_tab_config_autopreenchimento_e_edicao(self):
        from core.cursos_repo import obter_curso

        raiz = self.temp_dir / "raiz_etapa6"
        staging = self.temp_dir / "staging_etapa6"
        staging.mkdir()

        janela = MainWindow(raiz_dados=raiz)
        id_curso = janela.acao_novo_curso("Capacitação Licitações ANM")
        self.assertIsNotNone(id_curso)

        # 1. Importa lista de convidados e 2 listas de presença na Aba 1
        arq_conv = staging / "convocados_licitacoes.csv"
        arq_conv.write_text("Relatório\nData\nNome\nALICE <alice@anm.gov.br>\n", encoding="utf-8")
        janela.tab_arquivos._ao_importar_convidados([str(arq_conv)])
        self.assertEqual(janela.tab_arquivos.drop_convidados.property("state"), "loaded")

        arq_d1 = staging / "dia_1.csv"
        arq_d1.write_text(
            "1. Resumo\n"
            "Hora de início\t8/10/26, 8:00:00 AM\n"
            "Hora de término\t8/10/26, 12:00:00 PM\n"
            "Duração da reunião\t4h 0m 0s\n"
            "Tempo médio de participação\t3h 30m 0s\n"
            "2. Participantes\nNome\tEmail\tDuração da Reunião\n",
            encoding="utf-16",
        )
        arq_d2 = staging / "dia_2.csv"
        arq_d2.write_text(
            "1. Resumo\n"
            "Hora de início\t8/11/26, 2:00:00 PM\n"
            "Hora de término\t8/11/26, 6:00:00 PM\n"
            "Duração da reunião\t4h 0m 0s\n"
            "Tempo médio de participação\t4h 30m 0s\n"
            "2. Participantes\nNome\tEmail\tDuração da Reunião\n",
            encoding="utf-16",
        )
        janela.tab_arquivos._ao_importar_presencas([str(arq_d2), str(arq_d1)])

        # 2. Tabela de encontros ordenada cronologicamente com horários extraídos do 1. Resumo
        tbl = janela.tab_arquivos.tabela_encontros
        self.assertEqual(tbl.rowCount(), 2)
        self.assertEqual(tbl.item(0, 1).text(), "dia_1.csv")
        self.assertEqual(tbl.item(0, 4).text(), "08:00")
        self.assertEqual(tbl.item(0, 5).text(), "12:00")
        self.assertEqual(tbl.item(1, 1).text(), "dia_2.csv")
        self.assertEqual(tbl.item(1, 4).text(), "14:00")
        self.assertEqual(tbl.item(1, 5).text(), "18:00")

        # 3. Aba 2 recebeu automaticamente soma de tempos médios (8h) e datas (10/08 a 11/08)
        self.assertEqual(janela.tab_config.spin_carga_horaria.value(), 8.0)
        self.assertEqual(janela.tab_config.txt_data_inicio.text(), "10/08/2026")
        self.assertEqual(janela.tab_config.txt_data_fim.text(), "11/08/2026")
        self.assertEqual(janela.tab_config.txt_data_emissao.text(), "11/08/2026")

        # 4. Edição individual de horário na célula da tabela persiste no curso.json
        tbl.item(0, 4).setText("08:30")
        _, cfg_pos_h = obter_curso(id_curso, raiz=raiz)
        self.assertEqual(cfg_pos_h.encontros[0].horario_inicio, "08:30")

        # 5. Aplicar horário para todos
        janela.tab_arquivos.txt_todos_ini.setText("09:00")
        janela.tab_arquivos.txt_todos_fim.setText("12:00")
        janela.tab_arquivos._ao_aplicar_todos()
        self.assertEqual(tbl.item(1, 4).text(), "09:00")
        self.assertEqual(tbl.item(1, 5).text(), "12:00")

        # 6. Salvar custo/e-mails ignorados na Aba 2 NÃO trava campos auto-preenchidos
        janela.tab_config.spin_custo_total.setValue(4500.0)
        janela.tab_config.txt_ignorar_emails.setPlainText("apoio@anm.gov.br")
        self.assertTrue(janela.tab_config.salvar_configuracoes())
        _, cfg_salvo = obter_curso(id_curso, raiz=raiz)
        self.assertEqual(cfg_salvo.custo_total_em_reais, 4500.0)
        self.assertEqual(cfg_salvo.ignorar_emails, ["apoio@anm.gov.br"])
        self.assertEqual(cfg_salvo.campos_editados_manualmente, [])

        # 7. Alterar carga horária manualmente marca campo e Recalcular pelos CSVs restaura
        janela.tab_config.spin_carga_horaria.setValue(12.0)
        self.assertTrue(janela.tab_config.salvar_configuracoes())
        _, cfg_manual = obter_curso(id_curso, raiz=raiz)
        self.assertIn("carga_horaria", cfg_manual.campos_editados_manualmente)

        janela.tab_config._ao_recalcular()
        self.assertEqual(janela.tab_config.spin_carga_horaria.value(), 8.0)

        # 8. Remover uma lista atualiza tabela e campos derivados
        janela.tab_arquivos._ao_remover_presenca("dia_1.csv")
        self.assertEqual(tbl.rowCount(), 1)
        self.assertEqual(janela.tab_config.spin_carga_horaria.value(), 4.5)
        self.assertEqual(janela.tab_config.txt_data_inicio.text(), "11/08/2026")

        janela.close()


if __name__ == "__main__":
    unittest.main()
