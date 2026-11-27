"""
Aba 4: Conferência de participantes aptos e geração/junção dos certificados
individuais e consolidados (PowerPoint COM e PDF).
"""

from pathlib import Path
from typing import Optional

import pandas as pd
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QButtonGroup,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from core.app_paths import localizar_modelo_global, obter_pasta_saida_curso
from core.caminhos import (
    LIMITE_CARACTERES_DIRETORIO_CURSO,
    esvaziar_pasta,
    existe_com_fallback,
    obter_nome_arquivo_junto,
    verificar_profundidade_caminho,
)
from core.certificados import (
    carregar_relatorio_presenca,
    filtrar_aptos,
    gerar_todos_certificados,
    juntar_certificados_pdf,
)
from core.config import ConfigValidationError, CursoConfig, obter_configuracao
from core.validator import listar_arquivos_validos
from gui.file_actions import abrir_no_sistema, localizar_primeiro_arquivo, salvar_copia_como
from gui.workers import BaseTaskWorker


class TabCertificadosWidget(QWidget):
    """Aba 4: Prévia de aptos e emissão de certificados em PDF/PowerPoint."""

    certificados_atualizados = Signal()
    log_solicitado = Signal(str, str)
    avancar_solicitado = Signal()

    def __init__(self, executar_worker_fn, raiz_dados: Path, parent=None):
        super().__init__(parent)
        self._executar_worker = executar_worker_fn
        self._raiz_dados = raiz_dados
        self._pasta_curso: Optional[Path] = None
        self._cfg: Optional[CursoConfig] = None
        self._df_relatorio: Optional[pd.DataFrame] = None
        self._df_aptos: pd.DataFrame = pd.DataFrame()
        self._df_nao_aptos: pd.DataFrame = pd.DataFrame()
        self._caminho_pdf_junto: Optional[Path] = None
        self._caminho_ppt_junto: Optional[Path] = None
        self._construir_ui()

    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # Card 1: Opções de Geração
        card_op = QFrame()
        card_op.setProperty("class", "Card")
        op_v = QVBoxLayout(card_op)
        op_v.setContentsMargins(14, 12, 14, 12)
        op_v.setSpacing(6)

        lbl_op = QLabel("Geração e Consolidação de Certificados (PowerPoint + PDF)")
        lbl_op.setProperty("class", "SectionHeader")
        op_v.addWidget(lbl_op)

        self.grupo_modo = QButtonGroup(self)
        self.rb_completo = QRadioButton(
            "Gerar todos os certificados + PowerPoint consolidado + PDF único consolidado (Recomendado)"
        )
        self.rb_completo.setChecked(True)
        self.rb_individuais = QRadioButton("Gerar apenas os PDFs individuais (sem consolidar)")
        self.rb_juntar = QRadioButton("Apenas juntar os PDFs já existentes em um único PDF")

        for rb in (self.rb_completo, self.rb_individuais, self.rb_juntar):
            self.grupo_modo.addButton(rb)
            op_v.addWidget(rb)

        barra_exec = QHBoxLayout()
        self.lbl_resumo_aptos = QLabel("Gere o relatório de presenças na Aba 3 para visualizar os aptos.")
        self.lbl_resumo_aptos.setProperty("class", "MutedText")

        self.btn_executar = QPushButton("Gerar Certificados")
        self.btn_executar.setObjectName("PrimaryButton")
        self.btn_executar.clicked.connect(self.acao_executar)

        barra_exec.addWidget(self.lbl_resumo_aptos, 1)
        barra_exec.addWidget(self.btn_executar)
        op_v.addLayout(barra_exec)
        layout.addWidget(card_op)

        # Card 2: Prévia de Participantes
        card_tab = QFrame()
        card_tab.setProperty("class", "Card")
        tab_v = QVBoxLayout(card_tab)
        tab_v.setContentsMargins(14, 12, 14, 12)
        tab_v.setSpacing(8)

        topo_tab = QHBoxLayout()
        lbl_prev = QLabel("Conferência de Participantes para Emissão")
        lbl_prev.setProperty("class", "SectionHeader")

        self.cmb_filtro = QComboBox()
        self.cmb_filtro.addItems(["Somente Aptos", "Somente Não Aptos", "Todos"])
        self.cmb_filtro.currentIndexChanged.connect(self._renderizar_tabela)

        topo_tab.addWidget(lbl_prev, 1)
        topo_tab.addWidget(self.cmb_filtro)
        tab_v.addLayout(topo_tab)

        self.tabela = QTableWidget(0, 6)
        self.tabela.setHorizontalHeaderLabels(["Nome", "UORG", "E-mail", "Presenças", "CH Total", "Situação"])
        self.tabela.setAlternatingRowColors(True)
        self.tabela.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabela.verticalHeader().setVisible(False)
        hdr = self.tabela.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        for c in range(1, 6):
            hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)
        tab_v.addWidget(self.tabela, 1)
        layout.addWidget(card_tab, 1)

        # Card 3: Saídas Geradas e Exportação
        card_saida = QFrame()
        card_saida.setProperty("class", "Card")
        sai_h = QHBoxLayout(card_saida)
        sai_h.setContentsMargins(14, 10, 14, 10)
        sai_h.setSpacing(8)

        self.lbl_saidas = QLabel("Nenhum certificado gerado ainda.")
        self.lbl_saidas.setProperty("class", "MutedText")

        self.btn_abrir_pdf = QPushButton("Abrir PDF Único")
        self.btn_abrir_pdf.clicked.connect(lambda: self._abrir(self._caminho_pdf_junto))
        self.btn_salvar_pdf = QPushButton("Salvar PDF...")
        self.btn_salvar_pdf.clicked.connect(lambda: self._salvar(self._caminho_pdf_junto, "Arquivo PDF (*.pdf)"))

        self.btn_abrir_ppt = QPushButton("Abrir PPTX Único")
        self.btn_abrir_ppt.clicked.connect(lambda: self._abrir(self._caminho_ppt_junto))
        self.btn_salvar_ppt = QPushButton("Salvar PPTX...")
        self.btn_salvar_ppt.clicked.connect(lambda: self._salvar(self._caminho_ppt_junto, "PowerPoint (*.pptx *.pptm)"))

        self.btn_avancar = QPushButton("Avançar para E-mails →")
        self.btn_avancar.setObjectName("PrimaryButton")
        self.btn_avancar.clicked.connect(self.avancar_solicitado.emit)

        sai_h.addWidget(self.lbl_saidas, 1)
        sai_h.addWidget(self.btn_abrir_pdf)
        sai_h.addWidget(self.btn_salvar_pdf)
        sai_h.addWidget(self.btn_abrir_ppt)
        sai_h.addWidget(self.btn_salvar_ppt)
        sai_h.addWidget(self.btn_avancar)
        layout.addWidget(card_saida)

    def carregar_curso(self, pasta_curso: Path, cfg: CursoConfig) -> None:
        self._pasta_curso = pasta_curso
        self._cfg = cfg
        self.atualizar_dados()

    def atualizar_dados(self) -> None:
        if self._pasta_curso is None:
            return

        pasta_saida = obter_pasta_saida_curso(self._pasta_curso)
        caminho_rel = localizar_primeiro_arquivo(pasta_saida / "relatorio_presencas", (".xlsx",))

        if caminho_rel is None:
            self._df_relatorio = None
            self._df_aptos = pd.DataFrame()
            self._df_nao_aptos = pd.DataFrame()
            self.lbl_resumo_aptos.setText("Nenhum relatório encontrado. Gere o relatório na Aba 3 primeiro.")
        else:
            try:
                self._df_relatorio = carregar_relatorio_presenca(caminho_rel)
                self._df_aptos, self._df_nao_aptos = filtrar_aptos(self._df_relatorio)
                self.lbl_resumo_aptos.setText(
                    f"Convocados: {len(self._df_relatorio)}  |  "
                    f"Aptos para certificação: {len(self._df_aptos)}  |  "
                    f"Não aptos: {len(self._df_nao_aptos)}"
                )
            except Exception as exc:
                self._df_relatorio = None
                self._df_aptos = pd.DataFrame()
                self._df_nao_aptos = pd.DataFrame()
                self.lbl_resumo_aptos.setText(f"Erro ao carregar relatório: {exc}")

        self._renderizar_tabela()
        self._atualizar_saidas_existentes()

    def _atualizar_saidas_existentes(self) -> None:
        if self._pasta_curso is None:
            return
        pasta_saida = obter_pasta_saida_curso(self._pasta_curso)
        pdfs_ind = listar_arquivos_validos(pasta_saida / "certificados_pdf", (".pdf",))
        self._caminho_pdf_junto = localizar_primeiro_arquivo(pasta_saida / "certificados_juntos_pdf", (".pdf",))
        self._caminho_ppt_junto = localizar_primeiro_arquivo(pasta_saida / "certificados_juntos_ppt", (".pptx", ".pptm"))

        self.btn_abrir_pdf.setEnabled(self._caminho_pdf_junto is not None)
        self.btn_salvar_pdf.setEnabled(self._caminho_pdf_junto is not None)
        self.btn_abrir_ppt.setEnabled(self._caminho_ppt_junto is not None)
        self.btn_salvar_ppt.setEnabled(self._caminho_ppt_junto is not None)

        self.lbl_saidas.setText(
            f"PDFs individuais: {len(pdfs_ind)}  |  "
            f"PDF único: {'Sim' if self._caminho_pdf_junto else 'Não'}  |  "
            f"PPTX único: {'Sim' if self._caminho_ppt_junto else 'Não'}"
        )

    def _renderizar_tabela(self) -> None:
        self.tabela.setRowCount(0)
        if self._df_relatorio is None:
            return

        idx_f = self.cmb_filtro.currentIndex()
        if idx_f == 0:
            df = self._df_aptos
        elif idx_f == 1:
            df = self._df_nao_aptos
        else:
            df = self._df_relatorio

        negrito = QFont()
        negrito.setBold(True)

        for _, row_s in df.iterrows():
            r = self.tabela.rowCount()
            self.tabela.insertRow(r)
            apto = str(row_s.get("Apto por CH", "")).strip()
            vals = [
                str(row_s.get("Nome", "")),
                str(row_s.get("UORG", "") if pd.notna(row_s.get("UORG")) else ""),
                str(row_s.get("Email", "") if pd.notna(row_s.get("Email")) else ""),
                str(row_s.get("Presenças", "") if pd.notna(row_s.get("Presenças")) else ""),
                str(row_s.get("CH Total", "") if pd.notna(row_s.get("CH Total")) else ""),
                apto,
            ]
            for c_idx, txt in enumerate(vals):
                it = QTableWidgetItem(txt)
                if c_idx >= 3:
                    it.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if c_idx == 5:
                    it.setFont(negrito)
                    if apto.upper() == "APTO":
                        it.setForeground(QColor("#168821"))
                        it.setBackground(QColor("#E8F5E9"))
                    else:
                        it.setForeground(QColor("#C62828"))
                        it.setBackground(QColor("#FFEBEE"))
                self.tabela.setItem(r, c_idx, it)

    def _abrir(self, caminho: Optional[Path]) -> None:
        if caminho is not None:
            abrir_no_sistema(self, caminho)

    def _salvar(self, caminho: Optional[Path], filtro: str) -> None:
        if caminho is not None:
            salvar_copia_como(self, caminho, filtro)

    def acao_executar(self) -> None:
        if self._pasta_curso is None:
            return

        eh_profundo, _, msg_prof = verificar_profundidade_caminho(self._pasta_curso, LIMITE_CARACTERES_DIRETORIO_CURSO)
        if eh_profundo:
            QMessageBox.warning(self, "Caminho longo detectado", msg_prof)
            return

        pasta_saida = obter_pasta_saida_curso(self._pasta_curso)
        pasta_pdf = pasta_saida / "certificados_pdf"
        pasta_juntos_pdf = pasta_saida / "certificados_juntos_pdf"
        pasta_juntos_ppt = pasta_saida / "certificados_juntos_ppt"
        pasta_logs = pasta_saida / "logs"

        # Modo 3: Apenas juntar PDFs existentes
        if self.rb_juntar.isChecked():
            nome_curso = self._cfg.nome_do_curso if self._cfg else self._pasta_curso.name
            nome_arq = obter_nome_arquivo_junto(nome_curso, pasta_juntos_pdf, ".pdf")
            caminho_junto = pasta_juntos_pdf / nome_arq

            def _tarefa_juntar(on_log=None, on_progress=None, cancel_event=None):
                esvaziar_pasta(pasta_juntos_pdf, (".pdf",))
                return juntar_certificados_pdf(
                    pasta_pdf,
                    caminho_junto,
                    on_log=on_log,
                    on_progress=on_progress,
                    cancel_event=cancel_event,
                )

            worker = BaseTaskWorker(_tarefa_juntar, usar_com=False)

            def _fim_juntar(_res):
                self._atualizar_saidas_existentes()
                self.certificados_atualizados.emit()

            self._executar_worker(
                worker,
                titulo_inicio="Unificando certificados PDF...",
                titulo_fim="PDF consolidado gerado com sucesso!",
                cancelavel=True,
                ao_concluir=_fim_juntar,
            )
            return

        # Modos 1 e 2: Geração via PowerPoint COM
        if self._df_relatorio is None or self._df_aptos.empty:
            QMessageBox.information(
                self,
                "Nenhum participante apto",
                "Não há participantes aptos no relatório de presença para emitir certificados.",
            )
            return

        modelo_ppt = localizar_modelo_global(self._raiz_dados)
        if modelo_ppt is None or not existe_com_fallback(modelo_ppt):
            QMessageBox.warning(
                self,
                "Modelo não cadastrado",
                "Cadastre o modelo de certificado (.pptx/.pptm) nos Recursos Globais na barra lateral.",
            )
            return

        try:
            config = obter_configuracao(self._pasta_curso)
        except ConfigValidationError as exc:
            QMessageBox.warning(self, "Configurações inválidas", "\n".join(exc.errors))
            return

        consolidar = self.rb_completo.isChecked()
        nome_curso = config.get("NOME_DO_CURSO", self._pasta_curso.name)
        ext_modelo = modelo_ppt.suffix.lower()

        caminho_ppt_junto = (
            pasta_juntos_ppt / obter_nome_arquivo_junto(nome_curso, pasta_juntos_ppt, ext_modelo)
            if consolidar else None
        )
        caminho_pdf_junto = (
            pasta_juntos_pdf / obter_nome_arquivo_junto(nome_curso, pasta_juntos_pdf, ".pdf")
            if consolidar else None
        )

        def _tarefa_certificados(on_log=None, on_progress=None, cancel_event=None):
            esvaziar_pasta(pasta_juntos_pdf, (".pdf",))
            esvaziar_pasta(pasta_juntos_ppt, (".pptx", ".pptm"))
            esvaziar_pasta(pasta_pdf, (".pdf",))

            df_res = gerar_todos_certificados(
                df_aptos=self._df_aptos,
                modelo_powerpoint=modelo_ppt,
                pasta_ppt=None,
                pasta_pdf=pasta_pdf,
                pasta_logs=pasta_logs,
                config=config,
                caminho_ppt_junto=caminho_ppt_junto,
                salvar_ppt_individual=False,
                on_log=on_log,
                on_progress=on_progress,
                cancel_event=cancel_event,
            )

            if consolidar and caminho_pdf_junto is not None and not (cancel_event and cancel_event.is_set()):
                juntar_certificados_pdf(
                    pasta_pdf,
                    caminho_pdf_junto,
                    on_log=on_log,
                    on_progress=on_progress,
                    cancel_event=cancel_event,
                )

            return df_res

        worker = BaseTaskWorker(_tarefa_certificados, usar_com=True)

        def _fim_certs(_res):
            self._atualizar_saidas_existentes()
            self.certificados_atualizados.emit()

        self._executar_worker(
            worker,
            titulo_inicio=f"Gerando certificados para {len(self._df_aptos)} participante(s)...",
            titulo_fim="Geração de certificados concluída!",
            cancelavel=True,
            ao_concluir=_fim_certs,
        )
