"""
Janela principal do Gerador de Certificados ANM (DIVCAM / SGP).
Contém o cabeçalho institucional, a sidebar fixa (Recursos Globais + Lista de Cursos)
e a área central com o fluxo em 5 abas.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStackedWidget,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from core.app_paths import garantir_estrutura_raiz, obter_caminho_app_config
from core.config import CursoConfig, carregar_app_config, salvar_app_config
from core.cursos_repo import (
    criar_curso,
    definir_modelo_global,
    definir_servidores_global,
    excluir_curso,
    listar_cursos,
    obter_curso,
    obter_status_recursos_globais,
)
from gui.tabs.tab_arquivos import TabArquivosWidget
from gui.tabs.tab_certificados import TabCertificadosWidget
from gui.tabs.tab_config import TabConfigWidget
from gui.tabs.tab_emails import TabEmailsWidget
from gui.tabs.tab_relatorio import TabRelatorioWidget
from gui.widgets.drop_zone import DropZoneWidget
from gui.widgets.log_panel import LogPanelWidget
from gui.workers import BaseTaskWorker


class MainWindow(QMainWindow):
    """Janela principal do aplicativo desktop ANM DIVCAM/SGP."""

    def __init__(self, raiz_dados: Optional[Union[str, Path]] = None, parent=None):
        super().__init__(parent)
        self.raiz_dados = garantir_estrutura_raiz(raiz_dados)
        self.curso_atual_id: Optional[str] = None
        self.curso_atual_pasta: Optional[Path] = None
        self.curso_atual_cfg: Optional[CursoConfig] = None
        self._cursos_cache: List[Dict[str, Any]] = []
        self._worker_ativo: Optional[BaseTaskWorker] = None

        self.setWindowTitle("Gerador de Certificados — ANM (DIVCAM / SGP)")
        self.resize(1180, 740)
        self.setMinimumSize(980, 640)

        self._construir_ui()
        self._atualizar_recursos_globais()
        self._carregar_lista_cursos(selecionar_ultimo=True)

    # ------------------------------------------------------------------
    # CONSTRUÇÃO DA INTERFACE
    # ------------------------------------------------------------------
    def _construir_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)

        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        root_layout.addWidget(self._criar_header())

        corpo = QHBoxLayout()
        corpo.setContentsMargins(0, 0, 0, 0)
        corpo.setSpacing(0)

        self.sidebar = self._criar_sidebar()
        corpo.addWidget(self.sidebar)
        corpo.addWidget(self._criar_area_conteudo(), 1)

        root_layout.addLayout(corpo, 1)

    def _criar_header(self) -> QFrame:
        header = QFrame()
        header.setObjectName("HeaderBar")
        header.setFixedHeight(62)

        layout = QHBoxLayout(header)
        layout.setContentsMargins(18, 8, 18, 8)

        titulos = QVBoxLayout()
        titulos.setSpacing(1)
        lbl_titulo = QLabel("Gerador de Certificados — ANM")
        lbl_titulo.setObjectName("HeaderTitle")
        lbl_sub = QLabel("DIVCAM • Divisão de Capacitação e Movimentação  |  SGP — Superintendência de Gestão de Pessoas")
        lbl_sub.setObjectName("HeaderSubtitle")
        titulos.addWidget(lbl_titulo)
        titulos.addWidget(lbl_sub)

        layout.addLayout(titulos, 1)
        return header

    def _criar_sidebar(self) -> QFrame:
        sidebar = QFrame()
        sidebar.setObjectName("SidebarPanel")
        sidebar.setFixedWidth(300)

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(10)

        # 1. Recursos Globais
        lbl_globais = QLabel("RECURSOS GLOBAIS (ANM)")
        lbl_globais.setObjectName("SidebarSectionTitle")
        layout.addWidget(lbl_globais)

        self.drop_servidores = DropZoneWidget(
            titulo="Lista de Servidores (.xlsx)",
            extensoes=(".xlsx", ".xls"),
            dica="Arraste ou clique para cadastrar",
            multiplos=False,
            compacto=True,
        )
        self.drop_servidores.arquivos_selecionados.connect(self._ao_definir_servidores_global)
        layout.addWidget(self.drop_servidores)

        self.drop_modelo = DropZoneWidget(
            titulo="Modelo de Certificado (.pptx)",
            extensoes=(".pptx", ".pptm"),
            dica="Arraste ou clique para cadastrar",
            multiplos=False,
            compacto=True,
        )
        self.drop_modelo.arquivos_selecionados.connect(self._ao_definir_modelo_global)
        layout.addWidget(self.drop_modelo)

        separador = QFrame()
        separador.setFrameShape(QFrame.Shape.HLine)
        separador.setStyleSheet("color: #D6DEE8;")
        layout.addWidget(separador)

        # 2. Meus Cursos
        topo_cursos = QHBoxLayout()
        lbl_cursos = QLabel("MEUS CURSOS")
        lbl_cursos.setObjectName("SidebarSectionTitle")
        self.btn_novo_curso = QPushButton("+ Novo Curso")
        self.btn_novo_curso.setObjectName("PrimaryButton")
        self.btn_novo_curso.clicked.connect(self.acao_novo_curso)
        topo_cursos.addWidget(lbl_cursos, 1)
        topo_cursos.addWidget(self.btn_novo_curso)
        layout.addLayout(topo_cursos)

        self.txt_busca_curso = QLineEdit()
        self.txt_busca_curso.setPlaceholderText("Filtrar cursos pelo nome...")
        self.txt_busca_curso.textChanged.connect(self._filtrar_lista_cursos)
        layout.addWidget(self.txt_busca_curso)

        self.lista_cursos = QListWidget()
        self.lista_cursos.setObjectName("CourseList")
        self.lista_cursos.currentItemChanged.connect(self._ao_selecionar_item_curso)
        layout.addWidget(self.lista_cursos, 1)

        self.btn_excluir_curso = QPushButton("Excluir curso selecionado")
        self.btn_excluir_curso.setObjectName("DangerButton")
        self.btn_excluir_curso.setEnabled(False)
        self.btn_excluir_curso.clicked.connect(self.acao_excluir_curso)
        layout.addWidget(self.btn_excluir_curso)

        return sidebar

    def _criar_area_conteudo(self) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(16, 14, 16, 12)
        layout.setSpacing(10)

        self.log_panel = LogPanelWidget()
        self.log_panel.cancelar_solicitado.connect(self._solicitar_cancelamento_worker)

        self.stack_conteudo = QStackedWidget()
        self.pagina_vazia = self._criar_pagina_vazia()
        self.pagina_curso = self._criar_pagina_curso()

        self.stack_conteudo.addWidget(self.pagina_vazia)
        self.stack_conteudo.addWidget(self.pagina_curso)

        layout.addWidget(self.stack_conteudo, 1)
        layout.addWidget(self.log_panel)

        return container

    def _criar_pagina_vazia(self) -> QWidget:
        pag = QWidget()
        layout = QVBoxLayout(pag)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(10)

        card = QFrame()
        card.setProperty("class", "Card")
        card.setMaximumWidth(520)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(24, 24, 24, 24)
        card_layout.setSpacing(10)

        titulo = QLabel("Bem-vindo ao Gerador de Certificados ANM")
        titulo.setProperty("class", "SectionHeader")
        titulo.setAlignment(Qt.AlignmentFlag.AlignCenter)

        desc = QLabel(
            "1. Na barra lateral à esquerda, verifique se a Lista de Servidores (.xlsx) e o Modelo (.pptx) estão cadastrados.\n"
            "2. Clique em \"+ Novo Curso\" para criar um treinamento ou selecione um curso existente na lista."
        )
        desc.setWordWrap(True)
        desc.setProperty("class", "MutedText")

        btn = QPushButton("+ Criar Primeiro Curso")
        btn.setObjectName("PrimaryButton")
        btn.clicked.connect(self.acao_novo_curso)

        card_layout.addWidget(titulo)
        card_layout.addWidget(desc)
        card_layout.addWidget(btn)
        layout.addWidget(card)
        return pag

    def _criar_pagina_curso(self) -> QWidget:
        pag = QWidget()
        layout = QVBoxLayout(pag)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Cabeçalho do curso selecionado
        cab_card = QFrame()
        cab_card.setProperty("class", "Card")
        cab_layout = QHBoxLayout(cab_card)
        cab_layout.setContentsMargins(14, 10, 14, 10)

        info_box = QVBoxLayout()
        info_box.setSpacing(2)
        self.lbl_curso_titulo = QLabel("Curso")
        self.lbl_curso_titulo.setStyleSheet("font-size: 12pt; font-weight: 700; color: #071D41;")
        self.lbl_curso_meta = QLabel("")
        self.lbl_curso_meta.setProperty("class", "MutedText")
        info_box.addWidget(self.lbl_curso_titulo)
        info_box.addWidget(self.lbl_curso_meta)
        cab_layout.addLayout(info_box, 1)

        layout.addWidget(cab_card)

        # Abas principais (5 etapas)
        self.tabs = QTabWidget()
        self.tab_arquivos = TabArquivosWidget()
        self.tab_arquivos.curso_atualizado.connect(self._ao_curso_modificado_na_aba)
        self.tab_arquivos.log_solicitado.connect(self.log_panel.registrar_log)
        self.tab_arquivos.avancar_solicitado.connect(lambda: self.tabs.setCurrentIndex(1))

        self.tab_config = TabConfigWidget()
        self.tab_config.curso_atualizado.connect(self._ao_curso_modificado_na_aba)
        self.tab_config.log_solicitado.connect(self.log_panel.registrar_log)
        self.tab_config.avancar_solicitado.connect(lambda: self.tabs.setCurrentIndex(2))

        self.tab_relatorio = TabRelatorioWidget(self._executar_worker, self.raiz_dados)
        self.tab_relatorio.curso_atualizado.connect(self._ao_curso_modificado_na_aba)
        self.tab_relatorio.relatorio_atualizado.connect(self._ao_relatorio_atualizado)
        self.tab_relatorio.log_solicitado.connect(self.log_panel.registrar_log)
        self.tab_relatorio.avancar_solicitado.connect(lambda: self.tabs.setCurrentIndex(3))

        self.tab_certificados = TabCertificadosWidget(self._executar_worker, self.raiz_dados)
        self.tab_certificados.certificados_atualizados.connect(self._ao_certificados_atualizados)
        self.tab_certificados.log_solicitado.connect(self.log_panel.registrar_log)
        self.tab_certificados.avancar_solicitado.connect(lambda: self.tabs.setCurrentIndex(4))

        self.tab_emails = TabEmailsWidget(self._executar_worker)
        self.tab_emails.log_solicitado.connect(self.log_panel.registrar_log)

        self.tabs.addTab(self.tab_arquivos, "1. Arquivos e Horários")
        self.tabs.addTab(self.tab_config, "2. Dados e E-mail")
        self.tabs.addTab(self.tab_relatorio, "3. Relatório e Justificativas")
        self.tabs.addTab(self.tab_certificados, "4. Certificados")
        self.tabs.addTab(self.tab_emails, "5. Envio de E-mails")

        layout.addWidget(self.tabs, 1)
        return pag

    # ------------------------------------------------------------------
    # RECURSOS GLOBAIS
    # ------------------------------------------------------------------
    def _atualizar_recursos_globais(self) -> None:
        status = obter_status_recursos_globais(raiz=self.raiz_dados)
        if status["servidores_disponivel"]:
            self.drop_servidores.definir_carregado(f"Ativo: {status['servidores_nome']}")
        else:
            self.drop_servidores.definir_pendente("Pendente: arraste servidores (.xlsx)")

        if status["modelo_disponivel"]:
            self.drop_modelo.definir_carregado(f"Ativo: {status['modelo_nome']}")
        else:
            self.drop_modelo.definir_pendente("Pendente: arraste modelo (.pptx/.pptm)")

    def _ao_definir_servidores_global(self, caminhos: List[str]) -> None:
        if not caminhos:
            return
        try:
            definir_servidores_global(caminhos[0], raiz=self.raiz_dados)
            self._atualizar_recursos_globais()
            if self.curso_atual_pasta is not None:
                self.tab_relatorio.atualizar_visao_relatorio()
                self.tab_certificados.atualizar_dados()
            self.log_panel.registrar_log("sucesso", f"Lista global de servidores atualizada: {Path(caminhos[0]).name}")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao importar servidores", str(exc))

    def _ao_definir_modelo_global(self, caminhos: List[str]) -> None:
        if not caminhos:
            return
        try:
            definir_modelo_global(caminhos[0], raiz=self.raiz_dados)
            self._atualizar_recursos_globais()
            if self.curso_atual_pasta is not None:
                self.tab_relatorio.atualizar_visao_relatorio()
                self.tab_certificados.atualizar_dados()
            self.log_panel.registrar_log("sucesso", f"Modelo global de certificado atualizado: {Path(caminhos[0]).name}")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao importar modelo", str(exc))

    # ------------------------------------------------------------------
    # GESTÃO DE CURSOS NA SIDEBAR
    # ------------------------------------------------------------------
    def _formatar_texto_item_curso(self, c: Dict[str, Any]) -> str:
        badges = []
        if c.get("qtd_certificados", 0) > 0:
            badges.append("Certificados OK")
        elif c.get("tem_relatorio"):
            badges.append("Relatório OK")
        elif c.get("qtd_presencas", 0) > 0:
            badges.append(f"{c['qtd_presencas']} lista(s)")
        else:
            badges.append("Novo")

        periodo = ""
        if c.get("data_inicio"):
            periodo = f" • {c['data_inicio']}"
        return f"{c['nome_do_curso']}\n[{badges[0]}]{periodo}"

    def _carregar_lista_cursos(self, selecionar_ultimo: bool = False, id_preferido: Optional[str] = None) -> None:
        self._cursos_cache = listar_cursos(raiz=self.raiz_dados)
        alvo_id = id_preferido
        if not alvo_id and selecionar_ultimo:
            app_cfg = carregar_app_config(obter_caminho_app_config(self.raiz_dados))
            alvo_id = app_cfg.ultimo_curso_id or None

        self._filtrar_lista_cursos(self.txt_busca_curso.text(), id_selecionar=alvo_id)

    def _popular_lista_cursos_widget(self, texto: str = "", id_selecionar: Optional[str] = None) -> Optional[QListWidgetItem]:
        termo = (texto or "").strip().lower()
        alvo = id_selecionar or self.curso_atual_id

        self.lista_cursos.blockSignals(True)
        self.lista_cursos.clear()

        item_para_selecionar: Optional[QListWidgetItem] = None
        for c in self._cursos_cache:
            if termo and termo not in c["nome_do_curso"].lower():
                continue
            item = QListWidgetItem(self._formatar_texto_item_curso(c))
            item.setData(Qt.ItemDataRole.UserRole, c["id_curso"])
            self.lista_cursos.addItem(item)
            if alvo and c["id_curso"] == alvo:
                item_para_selecionar = item

        self.lista_cursos.blockSignals(False)
        return item_para_selecionar

    def _filtrar_lista_cursos(self, texto: str = "", id_selecionar: Optional[str] = None) -> None:
        item_para_selecionar = self._popular_lista_cursos_widget(texto=texto, id_selecionar=id_selecionar)

        if item_para_selecionar is not None:
            self.lista_cursos.setCurrentItem(item_para_selecionar)
            self._carregar_curso_na_tela(item_para_selecionar.data(Qt.ItemDataRole.UserRole))
        elif self.lista_cursos.count() > 0 and self.curso_atual_id is None:
            primeiro = self.lista_cursos.item(0)
            self.lista_cursos.setCurrentItem(primeiro)
            self._carregar_curso_na_tela(primeiro.data(Qt.ItemDataRole.UserRole))
        elif self.lista_cursos.count() == 0:
            self._limpar_selecao_curso()

    def _ao_selecionar_item_curso(self, atual: Optional[QListWidgetItem], _anterior: Optional[QListWidgetItem]) -> None:
        if atual is None:
            self._limpar_selecao_curso()
            return
        id_curso = atual.data(Qt.ItemDataRole.UserRole)
        self._carregar_curso_na_tela(id_curso)

    def _limpar_selecao_curso(self) -> None:
        self.curso_atual_id = None
        self.curso_atual_pasta = None
        self.curso_atual_cfg = None
        self.btn_excluir_curso.setEnabled(False)
        self.stack_conteudo.setCurrentWidget(self.pagina_vazia)

    def _carregar_curso_na_tela(self, id_curso: str) -> None:
        try:
            pasta, cfg = obter_curso(id_curso, raiz=self.raiz_dados)
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao abrir curso", str(exc))
            self._limpar_selecao_curso()
            return

        self.curso_atual_id = id_curso
        self.curso_atual_pasta = pasta
        self.curso_atual_cfg = cfg
        self.btn_excluir_curso.setEnabled(True)

        app_cfg_path = obter_caminho_app_config(self.raiz_dados)
        app_cfg = carregar_app_config(app_cfg_path)
        if app_cfg.ultimo_curso_id != id_curso:
            app_cfg.ultimo_curso_id = id_curso
            salvar_app_config(app_cfg_path, app_cfg)

        self._atualizar_cabecalho_curso()
        self.tab_arquivos.carregar_curso(pasta, cfg)
        self.tab_config.carregar_curso(pasta, cfg)
        self.tab_relatorio.carregar_curso(pasta, cfg)
        self.tab_certificados.carregar_curso(pasta, cfg)
        self.tab_emails.carregar_curso(pasta, cfg)
        self.stack_conteudo.setCurrentWidget(self.pagina_curso)

    def _ao_curso_modificado_na_aba(self, cfg_atualizado: CursoConfig) -> None:
        """Sincroniza cabeçalho, abas irmãs e badges da sidebar após edição em uma aba."""
        if self.curso_atual_pasta is None:
            return
        self.curso_atual_cfg = cfg_atualizado
        self._atualizar_cabecalho_curso()
        sender = self.sender()
        if sender is not self.tab_arquivos:
            self.tab_arquivos.carregar_curso(self.curso_atual_pasta, cfg_atualizado)
        if sender is not self.tab_config:
            self.tab_config.carregar_curso(self.curso_atual_pasta, cfg_atualizado)
        if sender is not self.tab_relatorio:
            self.tab_relatorio.carregar_curso(self.curso_atual_pasta, cfg_atualizado)
        if sender is not self.tab_certificados:
            self.tab_certificados.carregar_curso(self.curso_atual_pasta, cfg_atualizado)
        if sender is not self.tab_emails:
            self.tab_emails.carregar_curso(self.curso_atual_pasta, cfg_atualizado)
        self._atualizar_sidebar_sem_recarregar_abas()

    def _ao_relatorio_atualizado(self) -> None:
        """Atualiza abas dependentes e badge na barra lateral após gerar relatório Excel."""
        self.tab_certificados.atualizar_dados()
        self.tab_emails.atualizar_dados()
        self._atualizar_sidebar_sem_recarregar_abas()

    def _ao_certificados_atualizados(self) -> None:
        """Atualiza aba de e-mails e badge na barra lateral após gerar/juntar certificados."""
        self.tab_emails.atualizar_dados()
        self._atualizar_sidebar_sem_recarregar_abas()

    # ------------------------------------------------------------------
    # EXECUÇÃO ASSÍNCRONA DE TAREFAS (WORKERS QTHREAD)
    # ------------------------------------------------------------------
    def _solicitar_cancelamento_worker(self) -> None:
        if self._worker_ativo is not None and self._worker_ativo.isRunning():
            self._worker_ativo.solicitar_cancelamento()
            self.log_panel.btn_cancelar.setEnabled(False)
            self.log_panel.lbl_status.setText("Cancelamento solicitado — aguardando conclusão da etapa atual...")

    def _executar_worker(
        self,
        worker: BaseTaskWorker,
        titulo_inicio: str,
        titulo_fim: str,
        cancelavel: bool = True,
        ao_concluir=None,
    ) -> None:
        """Executa um BaseTaskWorker mantendo a UI responsiva e conectada ao LogPanelWidget."""
        if self._worker_ativo is not None and self._worker_ativo.isRunning():
            QMessageBox.information(
                self,
                "Tarefa em andamento",
                "Aguarde a conclusão ou cancele a tarefa atual antes de iniciar outra operação.",
            )
            return

        self._worker_ativo = worker
        self.sidebar.setEnabled(False)
        self.tabs.setEnabled(False)
        self.log_panel.iniciar_tarefa(titulo_inicio, cancelavel=cancelavel)

        estado: Dict[str, Any] = {"resultado": None, "erro": None}

        worker.log_emitido.connect(self.log_panel.registrar_log)
        worker.progresso_atualizado.connect(self.log_panel.atualizar_progresso)
        worker.concluido.connect(lambda res: estado.__setitem__("resultado", res))
        worker.erro_ocorrido.connect(lambda msg: estado.__setitem__("erro", msg))

        def _finalizar_worker() -> None:
            if worker.isRunning():
                worker.wait(1000)
            self._worker_ativo = None
            self.sidebar.setEnabled(True)
            self.tabs.setEnabled(True)

            if estado["erro"] is not None:
                msg_erro = str(estado["erro"])
                self.log_panel.registrar_log("erro", msg_erro)
                self.log_panel.finalizar_tarefa("Falha na execução da tarefa.", erro=True)
                QMessageBox.critical(self, "Erro na execução", msg_erro)
            else:
                if ao_concluir is not None:
                    try:
                        ao_concluir(estado["resultado"])
                    except Exception as exc:
                        self.log_panel.registrar_log("erro", f"Erro ao atualizar interface: {exc}")
                msg_fim = "Operação cancelada pelo usuário." if worker.cancel_event.is_set() else titulo_fim
                self.log_panel.finalizar_tarefa(msg_fim, erro=False)

            worker.deleteLater()

        worker.finished.connect(_finalizar_worker)
        worker.start()

    def _atualizar_sidebar_sem_recarregar_abas(self) -> None:
        self._cursos_cache = listar_cursos(raiz=self.raiz_dados)
        item_sel = self._popular_lista_cursos_widget(texto=self.txt_busca_curso.text(), id_selecionar=self.curso_atual_id)
        if item_sel is not None:
            self.lista_cursos.blockSignals(True)
            self.lista_cursos.setCurrentItem(item_sel)
            self.lista_cursos.blockSignals(False)

    def _atualizar_cabecalho_curso(self) -> None:
        if self.curso_atual_cfg is None:
            return
        cfg = self.curso_atual_cfg
        self.lbl_curso_titulo.setText(cfg.nome_do_curso)

        partes_meta = [f"ID: {cfg.id_curso}"]
        if cfg.carga_horaria > 0:
            ch_txt = int(cfg.carga_horaria) if float(cfg.carga_horaria).is_integer() else cfg.carga_horaria
            partes_meta.append(f"CH: {ch_txt}h")
        if cfg.data_inicio:
            if cfg.data_fim and cfg.data_fim != cfg.data_inicio:
                partes_meta.append(f"Período: {cfg.data_inicio} a {cfg.data_fim}")
            else:
                partes_meta.append(f"Data: {cfg.data_inicio}")
        partes_meta.append(f"Encontros: {len(cfg.encontros)}")
        self.lbl_curso_meta.setText("  |  ".join(partes_meta))

    def acao_novo_curso(self, nome_predefinido: Optional[str] = None) -> Optional[str]:
        """Cria um novo curso (solicitando o nome por diálogo quando não fornecido)."""
        if isinstance(nome_predefinido, bool):
            nome_predefinido = None

        if nome_predefinido is None:
            nome, ok = QInputDialog.getText(
                self,
                "Novo Curso — ANM",
                "Digite o nome oficial do curso (como aparecerá no certificado):",
            )
            if not ok:
                return None
            nome_curso = nome.strip()
        else:
            nome_curso = str(nome_predefinido).strip()

        if not nome_curso:
            QMessageBox.warning(self, "Nome inválido", "O nome do curso não pode ficar vazio.")
            return None

        try:
            _, cfg = criar_curso(nome_curso, raiz=self.raiz_dados)
            self._carregar_lista_cursos(id_preferido=cfg.id_curso)
            self.log_panel.registrar_log("sucesso", f"Curso criado: {cfg.nome_do_curso} ({cfg.id_curso})")
            return cfg.id_curso
        except Exception as exc:
            QMessageBox.critical(self, "Erro ao criar curso", str(exc))
            return None

    def acao_excluir_curso(self, confirmar: bool = True) -> bool:
        """Exclui o curso atualmente selecionado após confirmação."""
        if self.curso_atual_pasta is None or self.curso_atual_cfg is None:
            return False

        nome = self.curso_atual_cfg.nome_do_curso
        if confirmar:
            resp = QMessageBox.question(
                self,
                "Confirmar exclusão",
                f"Deseja realmente excluir o curso '{nome}' e todos os seus arquivos gerados?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if resp != QMessageBox.StandardButton.Yes:
                return False

        try:
            excluir_curso(self.curso_atual_pasta)
            self._limpar_selecao_curso()
            self._carregar_lista_cursos(selecionar_ultimo=False)
            self.log_panel.registrar_log("info", f"Curso excluído: {nome}")
            return True
        except Exception as exc:
            QMessageBox.critical(self, "Erro ao excluir curso", str(exc))
            return False

    def closeEvent(self, event) -> None:
        if self._worker_ativo is not None and self._worker_ativo.isRunning():
            resp = QMessageBox.question(
                self,
                "Tarefa em andamento",
                "Há uma operação em andamento. Deseja cancelar e fechar o aplicativo?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if resp != QMessageBox.StandardButton.Yes:
                event.ignore()
                return
            self._worker_ativo.solicitar_cancelamento()
            self._worker_ativo.wait(3000)
        super().closeEvent(event)
