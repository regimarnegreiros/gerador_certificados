"""
Aba 1: Importação de arquivos de entrada (arrastar-e-soltar) e configuração
individual de Horário de Início e Horário de Fim por lista de presença.
"""

from pathlib import Path
from typing import List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.app_paths import obter_pasta_entrada_curso
from core.config import CursoConfig
from core.cursos_repo import (
    aplicar_horario_para_todos_encontros,
    atualizar_horario_encontro,
    importar_convidados,
    importar_listas_presenca,
    remover_convidados,
    remover_lista_presenca,
    sincronizar_encontros_e_autopreencher,
)
from core.presenca import segundos_em_formato_hh_mm_ss
from core.validator import listar_arquivos_validos
from gui.widgets.drop_zone import DropZoneWidget


class TabArquivosWidget(QWidget):
    """Aba 1: Convidados, listas de presença e tabela editável de horários por encontro."""

    curso_atualizado = Signal(object)
    log_solicitado = Signal(str, str)
    avancar_solicitado = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pasta_curso: Optional[Path] = None
        self._cfg: Optional[CursoConfig] = None
        self._preenchendo_tabela = False
        self._construir_ui()

    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        # 1. Áreas de Arrastar e Soltar
        card_entradas = QFrame()
        card_entradas.setProperty("class", "Card")
        ent_layout = QVBoxLayout(card_entradas)
        ent_layout.setContentsMargins(14, 12, 14, 12)
        ent_layout.setSpacing(10)

        lbl_tit = QLabel("Arquivos de Entrada do Curso")
        lbl_tit.setProperty("class", "SectionHeader")
        lbl_sub = QLabel(
            "Arraste os arquivos do Microsoft Teams abaixo. Ao inserir as listas de presença, "
            "a carga horária, as datas e os horários iniciais/finais serão pré-preenchidos automaticamente."
        )
        lbl_sub.setProperty("class", "MutedText")
        lbl_sub.setWordWrap(True)
        ent_layout.addWidget(lbl_tit)
        ent_layout.addWidget(lbl_sub)

        drops_row = QHBoxLayout()
        drops_row.setSpacing(10)

        self.drop_convidados = DropZoneWidget(
            titulo="Lista de Convidados (.csv)",
            extensoes=(".csv",),
            dica="Arraste o CSV de convocados do Teams ou clique aqui",
            multiplos=False,
            permitir_limpar=True,
        )
        self.drop_convidados.arquivos_selecionados.connect(self._ao_importar_convidados)
        self.drop_convidados.limpar_solicitado.connect(self._ao_remover_convidados)

        self.drop_presencas = DropZoneWidget(
            titulo="Listas de Presença (.csv / .tsv)",
            extensoes=(".csv", ".tsv", ".txt"),
            dica="Arraste uma ou várias listas de encontros do Teams",
            multiplos=True,
            permitir_limpar=False,
        )
        self.drop_presencas.arquivos_selecionados.connect(self._ao_importar_presencas)

        drops_row.addWidget(self.drop_convidados, 1)
        drops_row.addWidget(self.drop_presencas, 1)
        ent_layout.addLayout(drops_row)
        layout.addWidget(card_entradas)

        # 2. Tabela de encontros e horários individuais
        card_encontros = QFrame()
        card_encontros.setProperty("class", "Card")
        enc_layout = QVBoxLayout(card_encontros)
        enc_layout.setContentsMargins(14, 12, 14, 12)
        enc_layout.setSpacing(8)

        topo_enc = QHBoxLayout()
        lbl_enc = QLabel("Encontros Detectados e Horários Individuais")
        lbl_enc.setProperty("class", "SectionHeader")
        topo_enc.addWidget(lbl_enc, 1)

        self.txt_todos_ini = QLineEdit()
        self.txt_todos_ini.setPlaceholderText("Início (ex: 08:00)")
        self.txt_todos_ini.setFixedWidth(120)
        self.txt_todos_fim = QLineEdit()
        self.txt_todos_fim.setPlaceholderText("Fim (ex: 12:00)")
        self.txt_todos_fim.setFixedWidth(120)

        self.btn_aplicar_todos = QPushButton("Aplicar horário a todos")
        self.btn_aplicar_todos.clicked.connect(self._ao_aplicar_todos)

        self.btn_recalcular = QPushButton("Recalcular pelos CSVs")
        self.btn_recalcular.clicked.connect(self._ao_recalcular_resumos)

        topo_enc.addWidget(self.txt_todos_ini)
        topo_enc.addWidget(self.txt_todos_fim)
        topo_enc.addWidget(self.btn_aplicar_todos)
        topo_enc.addWidget(self.btn_recalcular)
        enc_layout.addLayout(topo_enc)

        lbl_dica_h = QLabel(
            "Os horários abaixo são pré-preenchidos a partir de 'Hora de início' e 'Hora de término' da Seção 1 (Resumo). "
            "Clique diretamente nas células de horário (HH:MM) para ajustar encontros em turnos diferentes."
        )
        lbl_dica_h.setProperty("class", "MutedText")
        lbl_dica_h.setWordWrap(True)
        enc_layout.addWidget(lbl_dica_h)

        self.tabela_encontros = QTableWidget(0, 7)
        self.tabela_encontros.setHorizontalHeaderLabels([
            "Encontro",
            "Arquivo",
            "Data",
            "Tempo Médio",
            "Horário Início",
            "Horário Fim",
            "Ação",
        ])
        self.tabela_encontros.setAlternatingRowColors(True)
        self.tabela_encontros.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabela_encontros.verticalHeader().setVisible(False)
        self.tabela_encontros.itemChanged.connect(self._ao_editar_celula_horario)

        header = self.tabela_encontros.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)

        enc_layout.addWidget(self.tabela_encontros, 1)
        layout.addWidget(card_encontros, 1)

        # Rodapé de navegação
        rodape = QHBoxLayout()
        rodape.addStretch(1)
        self.btn_avancar = QPushButton("Avançar para Dados e E-mail →")
        self.btn_avancar.setObjectName("PrimaryButton")
        self.btn_avancar.clicked.connect(self.avancar_solicitado.emit)
        rodape.addWidget(self.btn_avancar)
        layout.addLayout(rodape)

    def carregar_curso(self, pasta_curso: Path, cfg: CursoConfig) -> None:
        self._pasta_curso = pasta_curso
        self._cfg = cfg

        # Estado da lista de convidados
        pasta_ent = obter_pasta_entrada_curso(pasta_curso)
        csvs_conv = listar_arquivos_validos(pasta_ent, (".csv",))
        if csvs_conv:
            nome_exibir = cfg.convidados_nome_original or csvs_conv[0].name
            self.drop_convidados.definir_carregado(f"Importado: {nome_exibir}")
        else:
            self.drop_convidados.definir_pendente()

        # Estado das listas de presença
        qtd = len(cfg.encontros)
        if qtd > 0:
            self.drop_presencas.definir_carregado(f"{qtd} lista(s) de presença carregada(s) (arraste mais para incluir)")
        else:
            self.drop_presencas.definir_pendente()

        self.btn_aplicar_todos.setEnabled(qtd > 0)
        self.btn_recalcular.setEnabled(qtd > 0)
        self._preencher_tabela_encontros()

    def _preencher_tabela_encontros(self) -> None:
        self._preenchendo_tabela = True
        try:
            self.tabela_encontros.setRowCount(0)
            if self._cfg is None:
                return

            for idx, enc in enumerate(self._cfg.encontros):
                row = self.tabela_encontros.rowCount()
                self.tabela_encontros.insertRow(row)

                item_enc = QTableWidgetItem(f"Encontro {idx + 1}")
                item_arq = QTableWidgetItem(enc.arquivo)
                item_data = QTableWidgetItem(enc.data_detectada or "-")
                tm_txt = segundos_em_formato_hh_mm_ss(enc.tempo_medio_segundos) if enc.tempo_medio_segundos > 0 else "-"
                item_tm = QTableWidgetItem(tm_txt)

                for item in (item_enc, item_arq, item_data, item_tm):
                    item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
                    if item is not item_arq:
                        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

                item_ini = QTableWidgetItem(enc.horario_inicio)
                item_ini.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item_ini.setData(Qt.ItemDataRole.UserRole, enc.horario_inicio)

                item_fim = QTableWidgetItem(enc.horario_fim)
                item_fim.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                item_fim.setData(Qt.ItemDataRole.UserRole, enc.horario_fim)

                self.tabela_encontros.setItem(row, 0, item_enc)
                self.tabela_encontros.setItem(row, 1, item_arq)
                self.tabela_encontros.setItem(row, 2, item_data)
                self.tabela_encontros.setItem(row, 3, item_tm)
                self.tabela_encontros.setItem(row, 4, item_ini)
                self.tabela_encontros.setItem(row, 5, item_fim)

                btn_rem = QPushButton("Remover")
                btn_rem.setObjectName("DangerButton")
                btn_rem.clicked.connect(lambda _chk=False, arq_nome=enc.arquivo: self._ao_remover_presenca(arq_nome))
                self.tabela_encontros.setCellWidget(row, 6, btn_rem)
        finally:
            self._preenchendo_tabela = False

    def _ao_importar_convidados(self, caminhos: List[str]) -> None:
        if self._pasta_curso is None or not caminhos:
            return
        try:
            cfg_novo = importar_convidados(self._pasta_curso, caminhos[0])
            self.carregar_curso(self._pasta_curso, cfg_novo)
            self.curso_atualizado.emit(cfg_novo)
            self.log_solicitado.emit("sucesso", f"Lista de convidados importada: {Path(caminhos[0]).name}")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao importar convidados", str(exc))

    def _ao_remover_convidados(self) -> None:
        if self._pasta_curso is None:
            return
        try:
            cfg_novo = remover_convidados(self._pasta_curso)
            self.carregar_curso(self._pasta_curso, cfg_novo)
            self.curso_atualizado.emit(cfg_novo)
            self.log_solicitado.emit("info", "Lista de convidados removida.")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao remover convidados", str(exc))

    def _ao_importar_presencas(self, caminhos: List[str]) -> None:
        if self._pasta_curso is None or not caminhos:
            return
        try:
            cfg_novo = importar_listas_presenca(self._pasta_curso, caminhos)
            self.carregar_curso(self._pasta_curso, cfg_novo)
            self.curso_atualizado.emit(cfg_novo)
            self.log_solicitado.emit(
                "sucesso",
                f"{len(caminhos)} lista(s) de presença importada(s). Dados do curso pré-preenchidos.",
            )
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao importar presenças", str(exc))

    def _ao_remover_presenca(self, nome_arquivo: str) -> None:
        if self._pasta_curso is None:
            return
        try:
            cfg_novo = remover_lista_presenca(self._pasta_curso, nome_arquivo)
            self.carregar_curso(self._pasta_curso, cfg_novo)
            self.curso_atualizado.emit(cfg_novo)
            self.log_solicitado.emit("info", f"Lista removida: {nome_arquivo}")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao remover lista", str(exc))

    def _ao_editar_celula_horario(self, item: QTableWidgetItem) -> None:
        if self._preenchendo_tabela or self._pasta_curso is None or item.column() not in (4, 5):
            return

        row = item.row()
        arq_item = self.tabela_encontros.item(row, 1)
        ini_item = self.tabela_encontros.item(row, 4)
        fim_item = self.tabela_encontros.item(row, 5)
        if not arq_item or not ini_item or not fim_item:
            return

        nome_arq = arq_item.text().strip()
        h_ini = ini_item.text().strip()
        h_fim = fim_item.text().strip()

        try:
            cfg_novo = atualizar_horario_encontro(self._pasta_curso, nome_arq, h_ini, h_fim)
            self._cfg = cfg_novo
            # Atualiza valores normalizados e backup em UserRole
            enc_atual = next((e for e in cfg_novo.encontros if e.arquivo == nome_arq), None)
            if enc_atual is not None:
                self._preenchendo_tabela = True
                try:
                    ini_item.setText(enc_atual.horario_inicio)
                    ini_item.setData(Qt.ItemDataRole.UserRole, enc_atual.horario_inicio)
                    fim_item.setText(enc_atual.horario_fim)
                    fim_item.setData(Qt.ItemDataRole.UserRole, enc_atual.horario_fim)
                finally:
                    self._preenchendo_tabela = False
            self.curso_atualizado.emit(cfg_novo)
            self.log_solicitado.emit("info", f"Horário atualizado para '{nome_arq}': {h_ini or '-'} às {h_fim or '-'}")
        except Exception as exc:
            QMessageBox.warning(self, "Horário inválido", str(exc))
            self._preenchendo_tabela = True
            try:
                anterior = item.data(Qt.ItemDataRole.UserRole) or ""
                item.setText(anterior)
            finally:
                self._preenchendo_tabela = False

    def _ao_aplicar_todos(self) -> None:
        if self._pasta_curso is None:
            return
        h_ini = self.txt_todos_ini.text().strip()
        h_fim = self.txt_todos_fim.text().strip()
        try:
            cfg_novo = aplicar_horario_para_todos_encontros(self._pasta_curso, h_ini, h_fim)
            self.carregar_curso(self._pasta_curso, cfg_novo)
            self.curso_atualizado.emit(cfg_novo)
            self.log_solicitado.emit("sucesso", f"Horário ({h_ini or '-'} às {h_fim or '-'}) aplicado a todos os encontros.")
        except Exception as exc:
            QMessageBox.warning(self, "Horário inválido", str(exc))

    def _ao_recalcular_resumos(self) -> None:
        if self._pasta_curso is None:
            return
        try:
            cfg_novo = sincronizar_encontros_e_autopreencher(self._pasta_curso, forcar_recalculo=True)
            self.carregar_curso(self._pasta_curso, cfg_novo)
            self.curso_atualizado.emit(cfg_novo)
            self.log_solicitado.emit("sucesso", "Datas, carga horária e horários recalculados a partir dos CSVs.")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao recalcular", str(exc))
