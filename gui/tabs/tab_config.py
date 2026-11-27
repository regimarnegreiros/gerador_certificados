"""
Aba 2: Revisão e edição das variáveis pré-preenchidas do curso,
critérios de aprovação, e-mails ignorados e mensagem do e-mail.
"""

from pathlib import Path
from typing import Dict, Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QDialog,
    QDoubleSpinBox,
    QFormLayout,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from core.config import (
    ConfigValidationError,
    CursoConfig,
    DEFAULT_ASSUNTO_EMAIL,
    DEFAULT_LOCAL_EMISSAO,
    normalizar_lista_emails,
    validar_curso_config,
)
from core.cursos_repo import salvar_curso, sincronizar_encontros_e_autopreencher
from core.emails import criar_corpo_email


class TabConfigWidget(QWidget):
    """Aba 2: Formulário estruturado para curso.json com pré-preenchimento automático."""

    curso_atualizado = Signal(object)
    log_solicitado = Signal(str, str)
    avancar_solicitado = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._pasta_curso: Optional[Path] = None
        self._cfg: Optional[CursoConfig] = None
        self._valores_carregados: Dict[str, object] = {}
        self._construir_ui()

    def _construir_ui(self) -> None:
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        conteudo = QWidget()
        layout = QVBoxLayout(conteudo)
        layout.setContentsMargins(4, 4, 8, 4)
        layout.setSpacing(12)

        # Card 1: Dados Gerais e Emissão
        card_gerais = QFrame()
        card_gerais.setProperty("class", "Card")
        gerais_v = QVBoxLayout(card_gerais)
        gerais_v.setContentsMargins(14, 12, 14, 12)
        gerais_v.setSpacing(8)

        topo_g = QHBoxLayout()
        lbl_g = QLabel("Dados Gerais e Emissão do Certificado")
        lbl_g.setProperty("class", "SectionHeader")
        self.btn_recalcular_vars = QPushButton("Recalcular pelos CSVs")
        self.btn_recalcular_vars.clicked.connect(self._ao_recalcular)
        topo_g.addWidget(lbl_g, 1)
        topo_g.addWidget(self.btn_recalcular_vars)
        gerais_v.addLayout(topo_g)

        form_g = QFormLayout()
        form_g.setSpacing(8)

        self.txt_nome_curso = QLineEdit()
        self.spin_carga_horaria = QDoubleSpinBox()
        self.spin_carga_horaria.setRange(0.0, 2000.0)
        self.spin_carga_horaria.setDecimals(2)
        self.spin_carga_horaria.setSingleStep(1.0)
        self.spin_carga_horaria.setSuffix(" h")

        self.spin_custo_total = QDoubleSpinBox()
        self.spin_custo_total.setRange(0.0, 99999999.99)
        self.spin_custo_total.setDecimals(2)
        self.spin_custo_total.setPrefix("R$ ")

        datas_row = QHBoxLayout()
        self.txt_data_inicio = QLineEdit()
        self.txt_data_inicio.setPlaceholderText("DD/MM/AAAA")
        self.txt_data_fim = QLineEdit()
        self.txt_data_fim.setPlaceholderText("DD/MM/AAAA (ou vazio se 1 dia)")
        self.txt_data_emissao = QLineEdit()
        self.txt_data_emissao.setPlaceholderText("DD/MM/AAAA")
        datas_row.addWidget(QLabel("Início:"))
        datas_row.addWidget(self.txt_data_inicio)
        datas_row.addWidget(QLabel("Fim:"))
        datas_row.addWidget(self.txt_data_fim)
        datas_row.addWidget(QLabel("Emissão:"))
        datas_row.addWidget(self.txt_data_emissao)

        self.txt_local_emissao = QLineEdit()

        form_g.addRow("Nome oficial do curso:", self.txt_nome_curso)

        linha_ch_custo = QHBoxLayout()
        linha_ch_custo.addWidget(self.spin_carga_horaria)
        linha_ch_custo.addSpacing(14)
        linha_ch_custo.addWidget(QLabel("Custo total do curso:"))
        linha_ch_custo.addWidget(self.spin_custo_total)
        linha_ch_custo.addStretch(1)
        form_g.addRow("Carga horária total:", linha_ch_custo)

        form_g.addRow("Datas (DD/MM/AAAA):", datas_row)
        form_g.addRow("Local de emissão:", self.txt_local_emissao)
        gerais_v.addLayout(form_g)
        layout.addWidget(card_gerais)

        # Card 2: Critérios e E-mails Ignorados
        card_criterios = QFrame()
        card_criterios.setProperty("class", "Card")
        crit_v = QVBoxLayout(card_criterios)
        crit_v.setContentsMargins(14, 12, 14, 12)
        crit_v.setSpacing(8)

        lbl_c = QLabel("Critérios de Aprovação e E-mails Ignorados")
        lbl_c.setProperty("class", "SectionHeader")
        crit_v.addWidget(lbl_c)

        criterios_row = QHBoxLayout()
        self.spin_pct_carga = QSpinBox()
        self.spin_pct_carga.setRange(0, 100)
        self.spin_pct_carga.setSuffix("%")

        self.spin_pct_encontros = QSpinBox()
        self.spin_pct_encontros.setRange(0, 100)
        self.spin_pct_encontros.setSuffix("%")

        self.spin_tempo_min = QSpinBox()
        self.spin_tempo_min.setRange(0, 600)
        self.spin_tempo_min.setSuffix(" min")

        criterios_row.addWidget(QLabel("Carga mínima:"))
        criterios_row.addWidget(self.spin_pct_carga)
        criterios_row.addSpacing(10)
        criterios_row.addWidget(QLabel("Frequência mínima:"))
        criterios_row.addWidget(self.spin_pct_encontros)
        criterios_row.addSpacing(10)
        criterios_row.addWidget(QLabel("Permanência mínima/encontro:"))
        criterios_row.addWidget(self.spin_tempo_min)
        criterios_row.addStretch(1)
        crit_v.addLayout(criterios_row)

        lbl_ign = QLabel("E-mails ignorados na apuração (organizadores, instrutores, apoio técnico — um por linha):")
        lbl_ign.setProperty("class", "MutedText")
        self.txt_ignorar_emails = QPlainTextEdit()
        self.txt_ignorar_emails.setMaximumHeight(85)
        self.txt_ignorar_emails.setPlaceholderText("organizador1@anm.gov.br\norganizador2@anm.gov.br")
        crit_v.addWidget(lbl_ign)
        crit_v.addWidget(self.txt_ignorar_emails)
        layout.addWidget(card_criterios)

        # Card 3: Comunicação e E-mail
        card_email = QFrame()
        card_email.setProperty("class", "Card")
        em_v = QVBoxLayout(card_email)
        em_v.setContentsMargins(14, 12, 14, 12)
        em_v.setSpacing(8)

        topo_em = QHBoxLayout()
        lbl_em = QLabel("Assunto e Mensagem do E-mail (Outlook)")
        lbl_em.setProperty("class", "SectionHeader")
        btn_tag_nome = QPushButton("+ Inserir {nome}")
        btn_tag_nome.clicked.connect(lambda: self._inserir_marcador("{nome}"))
        btn_tag_curso = QPushButton("+ Inserir {CURSO}")
        btn_tag_curso.clicked.connect(lambda: self._inserir_marcador("{CURSO}"))
        btn_preview = QPushButton("Pré-visualizar E-mail")
        btn_preview.clicked.connect(self._ao_previsualizar_email)
        topo_em.addWidget(lbl_em, 1)
        topo_em.addWidget(btn_tag_nome)
        topo_em.addWidget(btn_tag_curso)
        topo_em.addWidget(btn_preview)
        em_v.addLayout(topo_em)

        form_em = QFormLayout()
        self.txt_assunto_email = QLineEdit()
        self.txt_corpo_email = QPlainTextEdit()
        self.txt_corpo_email.setMinimumHeight(145)
        form_em.addRow("Assunto:", self.txt_assunto_email)
        form_em.addRow("Corpo:", self.txt_corpo_email)
        em_v.addLayout(form_em)
        layout.addWidget(card_email)

        layout.addStretch(1)
        scroll.setWidget(conteudo)
        root.addWidget(scroll, 1)

        # Rodapé de ações
        rodape = QHBoxLayout()
        rodape.addStretch(1)
        self.btn_salvar = QPushButton("Salvar Configurações")
        self.btn_salvar.setObjectName("PrimaryButton")
        self.btn_salvar.clicked.connect(self.salvar_configuracoes)

        self.btn_avancar = QPushButton("Avançar para Relatório →")
        self.btn_avancar.clicked.connect(self._salvar_e_avancar)

        rodape.addWidget(self.btn_salvar)
        rodape.addWidget(self.btn_avancar)
        root.addLayout(rodape)

    def carregar_curso(self, pasta_curso: Path, cfg: CursoConfig) -> None:
        self._pasta_curso = pasta_curso
        self._cfg = cfg

        self.txt_nome_curso.setText(cfg.nome_do_curso)
        self.spin_carga_horaria.setValue(float(cfg.carga_horaria or 0.0))
        self.spin_custo_total.setValue(float(cfg.custo_total_em_reais or 0.0))
        self.txt_data_inicio.setText(cfg.data_inicio)
        self.txt_data_fim.setText(cfg.data_fim)
        self.txt_data_emissao.setText(cfg.data_emissao)
        self.txt_local_emissao.setText(cfg.local_emissao or DEFAULT_LOCAL_EMISSAO)

        self.spin_pct_carga.setValue(int(round(float(cfg.percentual_minimo_carga) * 100)))
        self.spin_pct_encontros.setValue(int(round(float(cfg.percentual_minimo_encontros) * 100)))
        self.spin_tempo_min.setValue(int(cfg.tempo_minimo_presenca_minutos))
        self.txt_ignorar_emails.setPlainText("\n".join(cfg.ignorar_emails))

        self.txt_assunto_email.setText(cfg.assunto_email or DEFAULT_ASSUNTO_EMAIL)
        self.txt_corpo_email.setPlainText(cfg.corpo_email)

        self._valores_carregados = {
            "carga_horaria": round(float(cfg.carga_horaria or 0.0), 2),
            "data_inicio": cfg.data_inicio.strip(),
            "data_fim": cfg.data_fim.strip(),
            "data_emissao": cfg.data_emissao.strip(),
        }

    def _inserir_marcador(self, marcador: str) -> None:
        self.txt_corpo_email.insertPlainText(marcador)
        self.txt_corpo_email.setFocus()

    def _ao_previsualizar_email(self) -> None:
        nome_curso = self.txt_nome_curso.text().strip() or "Curso Exemplo"
        assunto = (self.txt_assunto_email.text().strip() or DEFAULT_ASSUNTO_EMAIL)
        assunto_render = assunto.replace("{CURSO}", nome_curso).replace("{nome}", "Maria da Silva")
        corpo_render = criar_corpo_email(
            nome="Maria da Silva",
            template_corpo=self.txt_corpo_email.toPlainText(),
            nome_curso=nome_curso,
        )

        dlg = QDialog(self)
        dlg.setWindowTitle("Pré-visualização do E-mail")
        dlg.resize(560, 360)
        v = QVBoxLayout(dlg)
        lbl_ass = QLabel(f"Assunto: {assunto_render}")
        lbl_ass.setStyleSheet("font-weight: 700; color: #071D41;")
        lbl_ass.setWordWrap(True)
        txt = QPlainTextEdit(corpo_render)
        txt.setReadOnly(True)
        btn_ok = QPushButton("Fechar")
        btn_ok.clicked.connect(dlg.accept)
        v.addWidget(lbl_ass)
        v.addWidget(txt, 1)
        v.addWidget(btn_ok, 0, Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def _ao_recalcular(self) -> None:
        if self._pasta_curso is None:
            return
        try:
            cfg_novo = sincronizar_encontros_e_autopreencher(self._pasta_curso, forcar_recalculo=True)
            self.carregar_curso(self._pasta_curso, cfg_novo)
            self.curso_atualizado.emit(cfg_novo)
            self.log_solicitado.emit("sucesso", "Variáveis recalculadas a partir das listas de presença.")
        except Exception as exc:
            QMessageBox.warning(self, "Erro ao recalcular", str(exc))

    def _construir_config_do_formulario(self) -> CursoConfig:
        base = CursoConfig.from_dict(self._cfg.to_dict()) if self._cfg is not None else CursoConfig()

        ch_val = round(float(self.spin_carga_horaria.value()), 2)
        d_ini = self.txt_data_inicio.text().strip()
        d_fim = self.txt_data_fim.text().strip()
        d_emi = self.txt_data_emissao.text().strip()

        editados = set(base.campos_editados_manualmente)
        if abs(ch_val - float(self._valores_carregados.get("carga_horaria", 0.0))) > 1e-6:
            editados.add("carga_horaria")
        if d_ini != self._valores_carregados.get("data_inicio", ""):
            editados.add("data_inicio")
        if d_fim != self._valores_carregados.get("data_fim", ""):
            editados.add("data_fim")
        if d_emi != self._valores_carregados.get("data_emissao", ""):
            editados.add("data_emissao")

        base.nome_do_curso = self.txt_nome_curso.text().strip()
        base.carga_horaria = int(ch_val) if ch_val.is_integer() else ch_val
        base.custo_total_em_reais = round(float(self.spin_custo_total.value()), 2)
        base.data_inicio = d_ini
        base.data_fim = d_fim
        base.data_emissao = d_emi
        base.local_emissao = self.txt_local_emissao.text().strip() or DEFAULT_LOCAL_EMISSAO
        base.percentual_minimo_carga = round(self.spin_pct_carga.value() / 100.0, 4)
        base.percentual_minimo_encontros = round(self.spin_pct_encontros.value() / 100.0, 4)
        base.tempo_minimo_presenca_minutos = int(self.spin_tempo_min.value())
        base.ignorar_emails = normalizar_lista_emails(self.txt_ignorar_emails.toPlainText())
        base.assunto_email = self.txt_assunto_email.text().strip() or DEFAULT_ASSUNTO_EMAIL
        base.corpo_email = self.txt_corpo_email.toPlainText()
        base.campos_editados_manualmente = sorted(editados)
        return base

    def tem_alteracoes_nao_salvas(self) -> bool:
        """Retorna True se o formulário difere do último CursoConfig carregado/salvo."""
        if self._cfg is None or self._pasta_curso is None:
            return False
        atual = self._construir_config_do_formulario()
        ref = self._cfg
        return any([
            atual.nome_do_curso != ref.nome_do_curso,
            abs(float(atual.carga_horaria) - float(ref.carga_horaria or 0.0)) > 1e-6,
            abs(float(atual.custo_total_em_reais) - float(ref.custo_total_em_reais or 0.0)) > 1e-6,
            atual.data_inicio != ref.data_inicio,
            atual.data_fim != ref.data_fim,
            atual.data_emissao != ref.data_emissao,
            atual.local_emissao != (ref.local_emissao or DEFAULT_LOCAL_EMISSAO),
            abs(float(atual.percentual_minimo_carga) - float(ref.percentual_minimo_carga)) > 1e-6,
            abs(float(atual.percentual_minimo_encontros) - float(ref.percentual_minimo_encontros)) > 1e-6,
            int(atual.tempo_minimo_presenca_minutos) != int(ref.tempo_minimo_presenca_minutos),
            atual.ignorar_emails != list(ref.ignorar_emails),
            atual.assunto_email != (ref.assunto_email or DEFAULT_ASSUNTO_EMAIL),
            atual.corpo_email != ref.corpo_email,
        ])

    def salvar_configuracoes(self, exibir_alerta_erro: bool = True) -> bool:
        """Valida o formulário atual e grava em curso.json."""
        if self._pasta_curso is None or self._cfg is None:
            return False

        candidato = self._construir_config_do_formulario()
        try:
            validar_curso_config(candidato)
        except ConfigValidationError as exc:
            if exibir_alerta_erro:
                QMessageBox.warning(self, "Configurações inválidas", "\n".join(exc.errors))
            return False

        salvar_curso(self._pasta_curso, candidato)
        self.carregar_curso(self._pasta_curso, candidato)
        self.curso_atualizado.emit(candidato)
        self.log_solicitado.emit("sucesso", f"Configurações do curso '{candidato.nome_do_curso}' salvas.")
        return True

    def _salvar_e_avancar(self) -> None:
        if self.salvar_configuracoes():
            self.avancar_solicitado.emit()
