"""
Aba 5: Conferência por rascunho no Microsoft Outlook e disparo em massa
com tabela de status em tempo real e exportação do log de envio.
"""

from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QColor, QFont
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
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
from core.caminhos import existe_com_fallback
from core.certificados import carregar_relatorio_presenca, filtrar_aptos
from core.config import ConfigValidationError, CursoConfig, obter_configuracao
from core.emails import enviar_email_teste, enviar_todos_emails, localizar_certificado_pdf
from core.validator import listar_arquivos_validos
from gui.file_actions import abrir_no_sistema, localizar_primeiro_arquivo, salvar_copia_como
from gui.workers import BaseTaskWorker


class TabEmailsWidget(QWidget):
    """Aba 5: Teste de rascunho e disparo em lote via Microsoft Outlook."""

    log_solicitado = Signal(str, str)

    def __init__(self, executar_worker_fn, parent=None):
        super().__init__(parent)
        self._executar_worker = executar_worker_fn
        self._pasta_curso: Optional[Path] = None
        self._cfg: Optional[CursoConfig] = None
        self._df_aptos: pd.DataFrame = pd.DataFrame()
        self._ultimo_log: Optional[Path] = None
        self._construir_ui()

    def _construir_ui(self) -> None:
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(10)

        # Card 1: Teste de conferência no Outlook
        card_teste = QFrame()
        card_teste.setProperty("class", "Card")
        tst_v = QVBoxLayout(card_teste)
        tst_v.setContentsMargins(14, 12, 14, 12)
        tst_v.setSpacing(8)

        lbl_tst = QLabel("1. Conferência Prévia (Abrir Rascunho no Outlook)")
        lbl_tst.setProperty("class", "SectionHeader")
        tst_v.addWidget(lbl_tst)

        linha_tst = QHBoxLayout()
        self.cmb_participante_teste = QComboBox()
        self.cmb_participante_teste.setMinimumWidth(300)

        self.txt_email_manual = QLineEdit()
        self.txt_email_manual.setPlaceholderText("Ou digite um e-mail alternativo de teste (opcional)")

        self.btn_teste_outlook = QPushButton("Abrir Rascunho de Teste no Outlook")
        self.btn_teste_outlook.clicked.connect(self.acao_teste_outlook)

        linha_tst.addWidget(self.cmb_participante_teste, 2)
        linha_tst.addWidget(self.txt_email_manual, 2)
        linha_tst.addWidget(self.btn_teste_outlook)
        tst_v.addLayout(linha_tst)
        layout.addWidget(card_teste)

        # Card 2: Disparo em massa e auditoria pré-envio
        card_envio = QFrame()
        card_envio.setProperty("class", "Card")
        env_v = QVBoxLayout(card_envio)
        env_v.setContentsMargins(14, 12, 14, 12)
        env_v.setSpacing(8)

        topo_env = QHBoxLayout()
        lbl_env = QLabel("2. Disparo em Massa via Microsoft Outlook")
        lbl_env.setProperty("class", "SectionHeader")

        self.btn_enviar_todos = QPushButton("Enviar para Todos os Aptos")
        self.btn_enviar_todos.setObjectName("PrimaryButton")
        self.btn_enviar_todos.clicked.connect(self.acao_enviar_todos)

        topo_env.addWidget(lbl_env, 1)
        topo_env.addWidget(self.btn_enviar_todos)
        env_v.addLayout(topo_env)

        self.lbl_auditoria_envio = QLabel("Gere o relatório e os certificados para habilitar o envio.")
        self.lbl_auditoria_envio.setProperty("class", "MutedText")
        self.lbl_auditoria_envio.setWordWrap(True)
        env_v.addWidget(self.lbl_auditoria_envio)

        self.tabela_envio = QTableWidget(0, 4)
        self.tabela_envio.setHorizontalHeaderLabels(["Destinatário", "E-mail", "Status", "Detalhe"])
        self.tabela_envio.setAlternatingRowColors(True)
        self.tabela_envio.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabela_envio.verticalHeader().setVisible(False)
        hdr = self.tabela_envio.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        env_v.addWidget(self.tabela_envio, 1)

        rodape_log = QHBoxLayout()
        self.lbl_log = QLabel("Nenhum log de envio registrado.")
        self.lbl_log.setProperty("class", "MutedText")
        self.btn_abrir_log = QPushButton("Abrir Último Log (.xlsx)")
        self.btn_abrir_log.clicked.connect(self._ao_abrir_log)
        self.btn_salvar_log = QPushButton("Salvar Log em...")
        self.btn_salvar_log.clicked.connect(self._ao_salvar_log)
        rodape_log.addWidget(self.lbl_log, 1)
        rodape_log.addWidget(self.btn_abrir_log)
        rodape_log.addWidget(self.btn_salvar_log)
        env_v.addLayout(rodape_log)

        layout.addWidget(card_envio, 1)

    def carregar_curso(self, pasta_curso: Path, cfg: CursoConfig) -> None:
        self._pasta_curso = pasta_curso
        self._cfg = cfg
        self.tabela_envio.setRowCount(0)
        self.atualizar_dados()

    def atualizar_dados(self) -> None:
        if self._pasta_curso is None:
            return

        pasta_saida = obter_pasta_saida_curso(self._pasta_curso)
        pasta_pdf = pasta_saida / "certificados_pdf"
        pasta_logs = pasta_saida / "logs"

        caminho_rel = localizar_primeiro_arquivo(pasta_saida / "relatorio_presencas", (".xlsx",))
        self.cmb_participante_teste.clear()

        if caminho_rel is None:
            self._df_aptos = pd.DataFrame()
            self.lbl_auditoria_envio.setText("Relatório de presença não encontrado.")
            self.btn_teste_outlook.setEnabled(False)
            self.btn_enviar_todos.setEnabled(False)
        else:
            try:
                df_rel = carregar_relatorio_presenca(caminho_rel)
                self._df_aptos, _ = filtrar_aptos(df_rel)
            except Exception as exc:
                self._df_aptos = pd.DataFrame()
                self.lbl_auditoria_envio.setText(f"Erro ao ler relatório: {exc}")

            prontos = 0
            sem_email = 0
            sem_pdf = 0

            for idx, (_, row_s) in enumerate(self._df_aptos.iterrows()):
                nome = str(row_s.get("Nome", "")).strip()
                email = str(row_s.get("Email", "")).strip() if pd.notna(row_s.get("Email")) else ""
                pdf = localizar_certificado_pdf(pasta_pdf, nome)
                tem_pdf = existe_com_fallback(pdf)

                if not email:
                    sem_email += 1
                elif not tem_pdf:
                    sem_pdf += 1
                else:
                    prontos += 1

                self.cmb_participante_teste.addItem(
                    f"{nome} ({email or 'sem e-mail'})",
                    userData={"idx": idx, "nome": nome, "email": email},
                )

            total = len(self._df_aptos)
            self.lbl_auditoria_envio.setText(
                f"Aptos: {total}  |  Prontos para envio: {prontos}  |  "
                f"Sem e-mail: {sem_email}  |  Sem PDF gerado: {sem_pdf}"
            )
            pode_enviar = total > 0 and prontos > 0
            self.btn_teste_outlook.setEnabled(total > 0)
            self.btn_enviar_todos.setEnabled(pode_enviar)

        logs_envio = sorted(
            [f for f in listar_arquivos_validos(pasta_logs, (".xlsx",)) if f.name.startswith("log_envio_")],
            reverse=True,
        )
        self._ultimo_log = logs_envio[0] if logs_envio else None
        self.btn_abrir_log.setEnabled(self._ultimo_log is not None)
        self.btn_salvar_log.setEnabled(self._ultimo_log is not None)
        self.lbl_log.setText(
            f"Último log: {self._ultimo_log.name}" if self._ultimo_log else "Nenhum log de envio registrado."
        )

    def _adicionar_linha_status(self, reg: Dict[str, Any]) -> None:
        row = self.tabela_envio.rowCount()
        self.tabela_envio.insertRow(row)

        status = str(reg.get("Status", ""))
        item_nome = QTableWidgetItem(str(reg.get("Nome", "")))
        item_email = QTableWidgetItem(str(reg.get("Email", "")))
        item_status = QTableWidgetItem(status)
        item_erro = QTableWidgetItem(str(reg.get("Erro", "")))

        item_status.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
        negrito = QFont()
        negrito.setBold(True)
        item_status.setFont(negrito)

        if status == "ENVIADO":
            item_status.setForeground(QColor("#168821"))
            item_status.setBackground(QColor("#E8F5E9"))
        else:
            item_status.setForeground(QColor("#C62828"))
            item_status.setBackground(QColor("#FFEBEE"))

        self.tabela_envio.setItem(row, 0, item_nome)
        self.tabela_envio.setItem(row, 1, item_email)
        self.tabela_envio.setItem(row, 2, item_status)
        self.tabela_envio.setItem(row, 3, item_erro)
        self.tabela_envio.scrollToBottom()

    def acao_teste_outlook(self) -> None:
        if self._pasta_curso is None or self._df_aptos.empty:
            return
        dados_sel = self.cmb_participante_teste.currentData()
        if not dados_sel:
            return

        nome = dados_sel["nome"]
        email_dest = self.txt_email_manual.text().strip() or dados_sel["email"]
        if not email_dest:
            QMessageBox.warning(self, "E-mail ausente", "Selecione um participante com e-mail ou informe um endereço no campo de teste.")
            return

        pasta_pdf = obter_pasta_saida_curso(self._pasta_curso) / "certificados_pdf"
        caminho_pdf = localizar_certificado_pdf(pasta_pdf, nome)
        if not existe_com_fallback(caminho_pdf):
            QMessageBox.warning(self, "PDF não encontrado", f"Gere o certificado de '{nome}' antes de testar o envio.")
            return

        try:
            config = obter_configuracao(self._pasta_curso)
        except ConfigValidationError as exc:
            QMessageBox.warning(self, "Configurações inválidas", "\n".join(exc.errors))
            return

        worker = BaseTaskWorker(
            enviar_email_teste,
            kwargs={
                "caminho_pdf": caminho_pdf,
                "nome": nome,
                "email_destinatario": email_dest,
                "config": config,
                "abrir_no_outlook": True,
            },
            usar_com=True,
            suporta_cancelamento=False,
            suporta_progresso=False,
            suporta_log=True,
        )

        self._executar_worker(
            worker,
            titulo_inicio=f"Abrindo rascunho de teste no Outlook para {nome}...",
            titulo_fim="Rascunho exibido no Microsoft Outlook!",
            cancelavel=False,
        )

    def acao_enviar_todos(self) -> None:
        if self._pasta_curso is None or self._df_aptos.empty:
            return

        try:
            config = obter_configuracao(self._pasta_curso)
        except ConfigValidationError as exc:
            QMessageBox.warning(self, "Configurações inválidas", "\n".join(exc.errors))
            return

        total = len(self._df_aptos)
        resp = QMessageBox.question(
            self,
            "Confirmar envio em massa",
            f"Serão disparados e-mails individuais pelo Microsoft Outlook para até {total} participante(s) apto(s).\n\n"
            "Deseja iniciar o envio definitivo agora?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if resp != QMessageBox.StandardButton.Yes:
            return

        self.tabela_envio.setRowCount(0)
        pasta_saida = obter_pasta_saida_curso(self._pasta_curso)

        worker = BaseTaskWorker(
            enviar_todos_emails,
            kwargs={
                "df_aptos": self._df_aptos,
                "pasta_pdf": pasta_saida / "certificados_pdf",
                "pasta_logs": pasta_saida / "logs",
                "config": config,
            },
            usar_com=True,
            suporta_cancelamento=True,
            suporta_progresso=True,
            suporta_log=True,
            suporta_item_status=True,
        )
        worker.item_status_emitido.connect(self._adicionar_linha_status)

        def _fim_envio(_df_envio):
            self.atualizar_dados()

        self._executar_worker(
            worker,
            titulo_inicio=f"Enviando certificados via Outlook para {total} participante(s)...",
            titulo_fim="Disparo de e-mails concluído!",
            cancelavel=True,
            ao_concluir=_fim_envio,
        )

    def _ao_abrir_log(self) -> None:
        if self._ultimo_log is not None:
            abrir_no_sistema(self, self._ultimo_log)

    def _ao_salvar_log(self) -> None:
        if self._ultimo_log is not None:
            salvar_copia_como(self, self._ultimo_log, "Planilha Excel (*.xlsx)")
