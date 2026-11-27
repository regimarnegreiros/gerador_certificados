"""
Aba 3: Geração do Relatório de Presenças Excel e matriz interativa para
revisão de ausências e aplicação de justificativas pela interface.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
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

from core.app_paths import obter_pasta_saida_curso
from core.caminhos import esvaziar_pasta, reduzir_nome_curso_para_arquivo, sanitizar_nome_arquivo
from core.config import ConfigValidationError, CursoConfig, format_time_str, obter_configuracao, parse_time
from core.cursos_repo import (
    calcular_tempo_sugerido_encontro,
    remover_justificativa,
    salvar_justificativa,
)
from core.presenca import (
    carregar_matriz_presencas,
    gerar_relatorio,
    segundos_em_formato_digital,
)
from core.validator import ValidationError, validar_estrutura_curso
from gui.file_actions import abrir_no_sistema, localizar_primeiro_arquivo, salvar_copia_como
from gui.workers import BaseTaskWorker


class _DialogJustificativa(QDialog):
    """Diálogo simples para informar tempo abonado e motivo opcional de uma falta."""

    def __init__(
        self,
        nome_participante: str,
        titulo_encontro: str,
        tempo_sugerido_seg: int,
        motivo_atual: str = "",
        ja_justificado: bool = False,
        parent=None,
    ):
        super().__init__(parent)
        self.setWindowTitle("Justificar Falta" if not ja_justificado else "Editar / Remover Justificativa")
        self.setMinimumWidth(420)
        self.remover_solicitado = False

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        lbl_info = QLabel(f"Participante: <b>{nome_participante}</b><br>Encontro: <b>{titulo_encontro}</b>")
        lbl_info.setWordWrap(True)
        layout.addWidget(lbl_info)

        form = QFormLayout()
        self.txt_tempo = QLineEdit(segundos_em_formato_digital(max(0, tempo_sugerido_seg)))
        self.txt_tempo.setPlaceholderText("HH:MM:SS (ex: 04:00:00)")
        self.txt_motivo = QLineEdit(motivo_atual)
        self.txt_motivo.setPlaceholderText("Opcional (ex: Atestado médico, missão institucional)")

        form.addRow("Tempo a abonar (HH:MM:SS):", self.txt_tempo)
        form.addRow("Motivo / Observação:", self.txt_motivo)
        layout.addLayout(form)

        botoes_box = QHBoxLayout()
        if ja_justificado:
            btn_rem = QPushButton("Remover Justificativa")
            btn_rem.setObjectName("DangerButton")
            btn_rem.clicked.connect(self._ao_remover)
            botoes_box.addWidget(btn_rem)

        botoes_box.addStretch(1)
        box_ok_cancel = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        box_ok_cancel.button(QDialogButtonBox.StandardButton.Ok).setText("Aplicar e Recalcular")
        box_ok_cancel.button(QDialogButtonBox.StandardButton.Cancel).setText("Cancelar")
        box_ok_cancel.accepted.connect(self._ao_confirmar)
        box_ok_cancel.rejected.connect(self.reject)
        botoes_box.addWidget(box_ok_cancel)
        layout.addLayout(botoes_box)

    def _ao_remover(self) -> None:
        self.remover_solicitado = True
        self.accept()

    def _ao_confirmar(self) -> None:
        try:
            segs = self.obter_tempo_segundos()
            if segs <= 0:
                raise ValueError("Informe um tempo maior que zero.")
            self.accept()
        except Exception as exc:
            QMessageBox.warning(self, "Tempo inválido", str(exc))

    def obter_tempo_segundos(self) -> int:
        t = parse_time(self.txt_tempo.text().strip())
        return t.hour * 3600 + t.minute * 60 + t.second

    def obter_motivo(self) -> str:
        return self.txt_motivo.text().strip()


class TabRelatorioWidget(QWidget):
    """Aba 3: Geração da planilha Excel e matriz interativa de justificativas."""

    relatorio_atualizado = Signal()
    curso_atualizado = Signal(object)
    log_solicitado = Signal(str, str)
    avancar_solicitado = Signal()

    def __init__(self, executar_worker_fn, raiz_dados: Path, parent=None):
        super().__init__(parent)
        self._executar_worker = executar_worker_fn
        self._raiz_dados = raiz_dados
        self._pasta_curso: Optional[Path] = None
        self._cfg: Optional[CursoConfig] = None
        self._caminho_relatorio: Optional[Path] = None
        self._matriz: Dict[str, Any] = {"titulos_encontros": [], "participantes": []}
        self._construir_ui()

    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # Card 1: Ações de geração e resumo
        card_top = QFrame()
        card_top.setProperty("class", "Card")
        top_v = QVBoxLayout(card_top)
        top_v.setContentsMargins(14, 12, 14, 12)
        top_v.setSpacing(8)

        linha_acoes = QHBoxLayout()
        lbl_t = QLabel("Apuração de Presenças e Planilha Excel")
        lbl_t.setProperty("class", "SectionHeader")

        self.btn_gerar = QPushButton("Validar e Gerar Relatório de Presenças")
        self.btn_gerar.setObjectName("PrimaryButton")
        self.btn_gerar.clicked.connect(self.acao_gerar_relatorio)

        self.btn_abrir_excel = QPushButton("Abrir no Excel")
        self.btn_abrir_excel.clicked.connect(self._ao_abrir_excel)
        self.btn_salvar_copia = QPushButton("Salvar Cópia em...")
        self.btn_salvar_copia.clicked.connect(self._ao_salvar_copia)

        linha_acoes.addWidget(lbl_t, 1)
        linha_acoes.addWidget(self.btn_gerar)
        linha_acoes.addWidget(self.btn_abrir_excel)
        linha_acoes.addWidget(self.btn_salvar_copia)
        top_v.addLayout(linha_acoes)

        self.lbl_resumo = QLabel("Nenhum relatório gerado ainda para este curso.")
        self.lbl_resumo.setProperty("class", "MutedText")
        top_v.addWidget(self.lbl_resumo)
        layout.addWidget(card_top)

        # Card 2: Matriz Interativa e Justificativas
        card_matriz = QFrame()
        card_matriz.setProperty("class", "Card")
        mat_v = QVBoxLayout(card_matriz)
        mat_v.setContentsMargins(14, 12, 14, 12)
        mat_v.setSpacing(8)

        topo_m = QHBoxLayout()
        lbl_m = QLabel("Revisão de Presenças e Justificativas de Faltas")
        lbl_m.setProperty("class", "SectionHeader")

        self.txt_busca = QLineEdit()
        self.txt_busca.setPlaceholderText("Buscar participante...")
        self.txt_busca.setFixedWidth(210)
        self.txt_busca.textChanged.connect(self._renderizar_matriz)

        self.cmb_filtro = QComboBox()
        self.cmb_filtro.addItems([
            "Todos os participantes",
            "Com faltas (Ausente)",
            "Apenas Não Aptos",
            "Com Justificativa",
        ])
        self.cmb_filtro.currentIndexChanged.connect(self._renderizar_matriz)

        self.btn_justificar = QPushButton("Justificar / Editar Falta...")
        self.btn_justificar.clicked.connect(self._ao_clicar_justificar)

        topo_m.addWidget(lbl_m, 1)
        topo_m.addWidget(self.txt_busca)
        topo_m.addWidget(self.cmb_filtro)
        topo_m.addWidget(self.btn_justificar)
        mat_v.addLayout(topo_m)

        lbl_dica = QLabel(
            "Dica: dê duplo clique em qualquer célula marcada como 'Ausente' ou 'Justificado' (ou selecione o encontro e clique em 'Justificar / Editar Falta...') "
            "para abonar o tempo do encontro e recalcular automaticamente a planilha."
        )
        lbl_dica.setProperty("class", "MutedText")
        lbl_dica.setWordWrap(True)
        mat_v.addWidget(lbl_dica)

        self.tabela_matriz = QTableWidget(0, 0)
        self.tabela_matriz.setAlternatingRowColors(True)
        self.tabela_matriz.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabela_matriz.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.tabela_matriz.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.tabela_matriz.verticalHeader().setVisible(False)
        self.tabela_matriz.cellDoubleClicked.connect(self._ao_duplo_clique_celula)
        mat_v.addWidget(self.tabela_matriz, 1)

        layout.addWidget(card_matriz, 1)

        rodape = QHBoxLayout()
        rodape.addStretch(1)
        self.btn_avancar = QPushButton("Avançar para Certificados →")
        self.btn_avancar.setObjectName("PrimaryButton")
        self.btn_avancar.clicked.connect(self.avancar_solicitado.emit)
        rodape.addWidget(self.btn_avancar)
        layout.addLayout(rodape)

    def carregar_curso(self, pasta_curso: Path, cfg: CursoConfig) -> None:
        self._pasta_curso = pasta_curso
        self._cfg = cfg
        self.atualizar_visao_relatorio()

    def atualizar_visao_relatorio(self) -> None:
        if self._pasta_curso is None:
            return
        pasta_rel = obter_pasta_saida_curso(self._pasta_curso) / "relatorio_presencas"
        self._caminho_relatorio = localizar_primeiro_arquivo(pasta_rel, (".xlsx",))

        tem_rel = self._caminho_relatorio is not None
        self.btn_abrir_excel.setEnabled(tem_rel)
        self.btn_salvar_copia.setEnabled(tem_rel)
        self.btn_justificar.setEnabled(tem_rel)

        if not tem_rel:
            self.lbl_resumo.setText("Nenhum relatório gerado ainda. Clique em 'Validar e Gerar Relatório de Presenças'.")
            self._matriz = {"titulos_encontros": [], "participantes": []}
            self._renderizar_matriz()
            return

        try:
            self._matriz = carregar_matriz_presencas(self._caminho_relatorio)
        except Exception as exc:
            self.lbl_resumo.setText(f"Erro ao ler relatório '{self._caminho_relatorio.name}': {exc}")
            self._matriz = {"titulos_encontros": [], "participantes": []}

        parts = self._matriz.get("participantes", [])
        total = len(parts)
        aptos = sum(1 for p in parts if str(p.get("apto_ch", "")).upper() == "APTO")
        nao_aptos = total - aptos
        qtd_just = len(self._cfg.justificativas) if self._cfg else 0

        self.lbl_resumo.setText(
            f"Arquivo: {self._caminho_relatorio.name}  |  "
            f"Convocados: {total}  |  Aptos: {aptos}  |  Não aptos: {nao_aptos}  |  Justificativas aplicadas: {qtd_just}"
        )
        self._renderizar_matriz()

    def _passa_filtro(self, p: Dict[str, Any]) -> bool:
        termo = self.txt_busca.text().strip().lower()
        if termo:
            alvo = f"{p.get('nome', '')} {p.get('email', '')}".lower()
            if termo not in alvo:
                return False

        modo = self.cmb_filtro.currentIndex()
        encontros = p.get("encontros", [])
        if modo == 1:  # Com faltas (Ausente)
            return any(e.get("status", "").upper() == "AUSENTE" for e in encontros)
        if modo == 2:  # Apenas Não Aptos
            return str(p.get("apto_ch", "")).upper() != "APTO"
        if modo == 3:  # Com Justificativa
            return any(e.get("status", "").upper() == "JUSTIFICADO" for e in encontros)
        return True

    def _renderizar_matriz(self) -> None:
        titulos = self._matriz.get("titulos_encontros", [])
        participantes = [p for p in self._matriz.get("participantes", []) if self._passa_filtro(p)]

        colunas = ["Nome", "E-mail"] + titulos + ["Presenças", "CH Total", "% CH", "Situação"]
        self.tabela_matriz.clear()
        self.tabela_matriz.setColumnCount(len(colunas))
        self.tabela_matriz.setHorizontalHeaderLabels(colunas)
        self.tabela_matriz.setRowCount(0)

        header = self.tabela_matriz.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        for col_idx in range(2, len(colunas)):
            header.setSectionResizeMode(col_idx, QHeaderView.ResizeMode.ResizeToContents)

        fonte_negrito = QFont()
        fonte_negrito.setBold(True)
        fonte_italico = QFont()
        fonte_italico.setBold(True)
        fonte_italico.setItalic(True)

        for p in participantes:
            row = self.tabela_matriz.rowCount()
            self.tabela_matriz.insertRow(row)

            item_nome = QTableWidgetItem(p["nome"])
            item_nome.setData(Qt.ItemDataRole.UserRole, p)
            item_email = QTableWidgetItem(p["email"])
            self.tabela_matriz.setItem(row, 0, item_nome)
            self.tabela_matriz.setItem(row, 1, item_email)

            encontros = p.get("encontros", [])
            for idx_e in range(len(titulos)):
                col_c = 2 + idx_e
                if idx_e < len(encontros):
                    enc_info = encontros[idx_e]
                    status = enc_info.get("status", "Ausente")
                    dur = enc_info.get("duracao", "00:00:00")
                    item_e = QTableWidgetItem(f"{status} ({dur})")
                    item_e.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    item_e.setData(Qt.ItemDataRole.UserRole, idx_e)

                    st_u = status.upper()
                    if st_u == "AUSENTE":
                        item_e.setForeground(QColor("#C62828"))
                        item_e.setBackground(QColor("#FFEBEE"))
                        item_e.setFont(fonte_negrito)
                    elif st_u == "JUSTIFICADO":
                        item_e.setForeground(QColor("#9C6500"))
                        item_e.setBackground(QColor("#FEF3C7"))
                        item_e.setFont(fonte_italico)
                    else:
                        item_e.setForeground(QColor("#168821"))
                else:
                    item_e = QTableWidgetItem("-")
                    item_e.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.tabela_matriz.setItem(row, col_c, item_e)

            base_col = 2 + len(titulos)
            item_pres = QTableWidgetItem(str(p.get("presencas", 0)))
            item_ch = QTableWidgetItem(str(p.get("ch_total", "0:00:00")))
            item_pct = QTableWidgetItem(f"{float(p.get('pct_ch', 0.0)) * 100:.1f}%")
            apto_txt = str(p.get("apto_ch", "NÃO APTO"))
            item_apto = QTableWidgetItem(apto_txt)

            for it in (item_pres, item_ch, item_pct, item_apto):
                it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            item_apto.setFont(fonte_negrito)
            if apto_txt.upper() == "APTO":
                item_apto.setForeground(QColor("#168821"))
                item_apto.setBackground(QColor("#E8F5E9"))
            else:
                item_apto.setForeground(QColor("#C62828"))
                item_apto.setBackground(QColor("#FFEBEE"))

            self.tabela_matriz.setItem(row, base_col, item_pres)
            self.tabela_matriz.setItem(row, base_col + 1, item_ch)
            self.tabela_matriz.setItem(row, base_col + 2, item_pct)
            self.tabela_matriz.setItem(row, base_col + 3, item_apto)

    def _preparar_parametros_relatorio(self) -> Optional[Dict[str, Any]]:
        if self._pasta_curso is None:
            return None

        try:
            auditoria = validar_estrutura_curso(self._pasta_curso, raiz_dados=self._raiz_dados)
        except ValidationError as exc:
            QMessageBox.warning(self, "Estrutura ou arquivos pendentes", "\n".join(exc.erros))
            return None

        if auditoria.get("avisos"):
            msg_avisos = "\n".join(auditoria["avisos"])
            resp = QMessageBox.question(
                self,
                "Avisos detectados",
                f"{msg_avisos}\n\nDeseja continuar mesmo assim?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if resp != QMessageBox.StandardButton.Yes:
                return None

        try:
            config = obter_configuracao(self._pasta_curso)
        except ConfigValidationError as exc:
            QMessageBox.warning(
                self,
                "Configurações inválidas",
                "Corrija os campos na aba '2. Dados e E-mail':\n\n" + "\n".join(exc.errors),
            )
            return None

        arquivos = auditoria["arquivos"]
        nome_seguro_curso = sanitizar_nome_arquivo(config["NOME_DO_CURSO"])
        pasta_saida_rel = obter_pasta_saida_curso(self._pasta_curso) / "relatorio_presencas"
        nome_arq_rel = f"Relatório de Presença - {nome_seguro_curso}.xlsx"
        if len(str(pasta_saida_rel / nome_arq_rel)) > 210 or len(nome_seguro_curso) > 40:
            curso_curto = reduzir_nome_curso_para_arquivo(config["NOME_DO_CURSO"], max_chars=35)
            nome_arq_rel = f"Relatório de Presença - {curso_curto}.xlsx"

        caminho_saida = pasta_saida_rel / nome_arq_rel
        return {
            "auditoria": auditoria,
            "config": config,
            "pasta_saida_rel": pasta_saida_rel,
            "caminho_saida": caminho_saida,
        }

    def acao_gerar_relatorio(self) -> None:
        """Valida entradas/configurações e dispara geração do relatório Excel em thread."""
        params = self._preparar_parametros_relatorio()
        if params is None:
            return

        def _tarefa_relatorio(on_log=None, on_progress=None):
            esvaziar_pasta(params["pasta_saida_rel"], (".xlsx",))
            arquivos = params["auditoria"]["arquivos"]
            config = params["config"]
            return gerar_relatorio(
                caminho_servidores=arquivos["servidores"],
                caminho_convocados=arquivos["convidados"],
                arquivos_presenca=arquivos["presencas"],
                caminho_saida_excel=params["caminho_saida"],
                carga_horaria_horas=config["CARGA_HORARIA"],
                tempo_comparecimento_minutos=config["TEMPO_MINIMO_PRESENCA_MINUTOS"],
                percentual_minimo_carga=config["PERCENTUAL_MINIMO_CARGA"],
                percentual_minimo_encontros=config["PERCENTUAL_MINIMO_ENCONTROS"],
                config=config,
                on_log=on_log,
                on_progress=on_progress,
            )

        worker = BaseTaskWorker(
            _tarefa_relatorio,
            usar_com=False,
            suporta_cancelamento=False,
            suporta_progresso=True,
            suporta_log=True,
        )

        def _ao_concluir(_df):
            self.atualizar_visao_relatorio()
            self.relatorio_atualizado.emit()

        self._executar_worker(
            worker,
            titulo_inicio="Gerando relatório de presenças Excel...",
            titulo_fim="Relatório de presenças gerado com sucesso!",
            cancelavel=False,
            ao_concluir=_ao_concluir,
        )

    def _ao_abrir_excel(self) -> None:
        if self._caminho_relatorio is not None:
            abrir_no_sistema(self, self._caminho_relatorio)

    def _ao_salvar_copia(self) -> None:
        if self._caminho_relatorio is not None:
            salvar_copia_como(self, self._caminho_relatorio, "Planilha Excel (*.xlsx)")

    def _ao_clicar_justificar(self) -> None:
        row = self.tabela_matriz.currentRow()
        col = self.tabela_matriz.currentColumn()
        if row < 0:
            QMessageBox.information(self, "Selecione um participante", "Selecione o encontro de um participante na tabela.")
            return
        titulos = self._matriz.get("titulos_encontros", [])
        if col < 2 or col >= 2 + len(titulos):
            # Se clicou no nome, tenta localizar o primeiro encontro Ausente/Justificado da linha
            item_nome = self.tabela_matriz.item(row, 0)
            p = item_nome.data(Qt.ItemDataRole.UserRole) if item_nome else None
            idx_alvo = None
            if p:
                for idx_e, enc in enumerate(p.get("encontros", [])):
                    if enc.get("status", "").upper() in ("AUSENTE", "JUSTIFICADO"):
                        idx_alvo = idx_e
                        break
            if idx_alvo is None:
                QMessageBox.information(self, "Selecione o encontro", "Clique diretamente na coluna do encontro que deseja justificar.")
                return
            col = 2 + idx_alvo
        self._ao_duplo_clique_celula(row, col)

    def _ao_duplo_clique_celula(self, row: int, col: int) -> None:
        if self._pasta_curso is None or self._cfg is None:
            return
        titulos = self._matriz.get("titulos_encontros", [])
        if col < 2 or col >= 2 + len(titulos):
            return

        item_nome = self.tabela_matriz.item(row, 0)
        if item_nome is None:
            return
        p = item_nome.data(Qt.ItemDataRole.UserRole)
        idx_enc = col - 2
        if not p or idx_enc >= len(self._cfg.encontros):
            return

        enc_cfg = self._cfg.encontros[idx_enc]
        enc_status = p["encontros"][idx_enc] if idx_enc < len(p.get("encontros", [])) else {}
        status_atual = str(enc_status.get("status", "Ausente")).strip().upper()

        if status_atual == "PRESENTE":
            QMessageBox.information(
                self,
                "Encontro já validado",
                f"{p['nome']} já consta como 'Presente' em '{titulos[idx_enc]}'.",
            )
            return

        ja_just = status_atual == "JUSTIFICADO"
        just_existente = next(
            (
                j for j in self._cfg.justificativas
                if j.arquivo_encontro == enc_cfg.arquivo
                and (
                    (p.get("email") and j.email.lower() == p["email"].lower())
                    or j.nome.upper() == p["nome"].upper()
                )
            ),
            None,
        )

        tempo_sugerido = (
            just_existente.tempo_abonado_segundos
            if just_existente is not None
            else calcular_tempo_sugerido_encontro(self._cfg, enc_cfg.arquivo)
        )

        dlg = _DialogJustificativa(
            nome_participante=p["nome"],
            titulo_encontro=titulos[idx_enc],
            tempo_sugerido_seg=tempo_sugerido,
            motivo_atual=just_existente.motivo if just_existente else "",
            ja_justificado=ja_just,
            parent=self,
        )
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        if dlg.remover_solicitado:
            cfg_novo = remover_justificativa(
                self._pasta_curso,
                nome=p["nome"],
                email=p.get("email", ""),
                arquivo_encontro=enc_cfg.arquivo,
            )
            self.log_solicitado.emit("info", f"Justificativa removida: {p['nome']} ({enc_cfg.arquivo})")
        else:
            cfg_novo = salvar_justificativa(
                self._pasta_curso,
                nome=p["nome"],
                email=p.get("email", ""),
                arquivo_encontro=enc_cfg.arquivo,
                tempo_abonado_segundos=dlg.obter_tempo_segundos(),
                motivo=dlg.obter_motivo(),
            )
            self.log_solicitado.emit("sucesso", f"Falta justificada: {p['nome']} ({enc_cfg.arquivo})")

        self._cfg = cfg_novo
        self.curso_atualizado.emit(cfg_novo)
        self.acao_gerar_relatorio()
