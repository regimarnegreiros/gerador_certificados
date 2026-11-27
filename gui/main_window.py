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
from gui.widgets.drop_zone import DropZoneWidget
from gui.widgets.log_panel import LogPanelWidget


class MainWindow(QMainWindow):
    """Janela principal do aplicativo desktop ANM DIVCAM/SGP."""

    def __init__(self, raiz_dados: Optional[Union[str, Path]] = None, parent=None):
        super().__init__(parent)
        self.raiz_dados = garantir_estrutura_raiz(raiz_dados)
        self.curso_atual_id: Optional[str] = None
        self.curso_atual_pasta: Optional[Path] = None
        self.curso_atual_cfg: Optional[CursoConfig] = None
        self._cursos_cache: List[Dict[str, Any]] = []

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

        corpo.addWidget(self._criar_sidebar())
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

        self.stack_conteudo = QStackedWidget()
        self.pagina_vazia = self._criar_pagina_vazia()
        self.pagina_curso = self._criar_pagina_curso()

        self.stack_conteudo.addWidget(self.pagina_vazia)
        self.stack_conteudo.addWidget(self.pagina_curso)

        layout.addWidget(self.stack_conteudo, 1)

        self.log_panel = LogPanelWidget()
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
        self.tab_arquivos = self._criar_aba_placeholder(
            "1. Arquivos e Horários",
            "Aqui serão importados os arquivos de convidados (.csv) e listas de presença (.csv/.tsv), com edição de horários por encontro. (Etapa 6)",
        )
        self.tab_config = self._criar_aba_placeholder(
            "2. Dados e E-mail",
            "Aqui ficarão as variáveis pré-preenchidas automaticamente a partir do resumo dos CSVs e o texto do e-mail. (Etapa 6)",
        )
        self.tab_relatorio = self._criar_aba_placeholder(
            "3. Relatório e Justificativas",
            "Aqui será gerado o relatório Excel e exibida a matriz interativa para justificar faltas pela interface. (Etapa 7)",
        )
        self.tab_certificados = self._criar_aba_placeholder(
            "4. Certificados",
            "Aqui será feita a conferência final de aptos e geração/junção de certificados em PowerPoint e PDF. (Etapa 7)",
        )
        self.tab_emails = self._criar_aba_placeholder(
            "5. Envio de E-mails",
            "Aqui serão realizados o teste no Outlook e o disparo em lote para os participantes aptos. (Etapa 7)",
        )

        self.tabs.addTab(self.tab_arquivos, "1. Arquivos e Horários")
        self.tabs.addTab(self.tab_config, "2. Dados e E-mail")
        self.tabs.addTab(self.tab_relatorio, "3. Relatório e Justificativas")
        self.tabs.addTab(self.tab_certificados, "4. Certificados")
        self.tabs.addTab(self.tab_emails, "5. Envio de E-mails")

        layout.addWidget(self.tabs, 1)
        return pag

    def _criar_aba_placeholder(self, titulo: str, descricao: str) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        lbl_t = QLabel(titulo)
        lbl_t.setProperty("class", "SectionHeader")
        lbl_d = QLabel(descricao)
        lbl_d.setWordWrap(True)
        lbl_d.setProperty("class", "MutedText")
        layout.addWidget(lbl_t)
        layout.addWidget(lbl_d)
        return w

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
            self.log_panel.registrar_log("sucesso", f"Lista global de servidores atualizada: {Path(caminhos[0]).name}")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao importar servidores", str(exc))

    def _ao_definir_modelo_global(self, caminhos: List[str]) -> None:
        if not caminhos:
            return
        try:
            definir_modelo_global(caminhos[0], raiz=self.raiz_dados)
            self._atualizar_recursos_globais()
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

    def _filtrar_lista_cursos(self, texto: str = "", id_selecionar: Optional[str] = None) -> None:
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
        self.stack_conteudo.setCurrentWidget(self.pagina_curso)

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
