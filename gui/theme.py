"""
Tema visual institucional da Agência Nacional de Mineração (ANM) - DIVCAM/SGP.
Estilo QSS plano e leve (sem animações, sombras ou transparências pesadas)
visando alto desempenho e legibilidade em estações corporativas.
"""

from PySide6.QtWidgets import QApplication


COR_AZUL_MARINHO = "#071D41"
COR_AZUL_EXCEL = "#1F4E78"
COR_AZUL_PRIMARIO = "#1351B4"
COR_AZUL_PRIMARIO_HOVER = "#0C3D8A"
COR_AMBAR_ANM = "#D97706"
COR_OURO_GOV = "#FFD000"
COR_FUNDO_GERAL = "#F4F6F9"
COR_SUPERFICIE = "#FFFFFF"
COR_BORDA = "#D6DEE8"
COR_TEXTO = "#1E293B"
COR_TEXTO_SUAVE = "#5A6B7F"
COR_SUCESSO = "#168821"
COR_SUCESSO_FUNDO = "#E8F5E9"
COR_AVISO = "#B45309"
COR_AVISO_FUNDO = "#FEF3C7"
COR_ERRO = "#C62828"
COR_ERRO_FUNDO = "#FFEBEE"


QSS_ANM = f"""
QWidget {{
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 10pt;
    color: {COR_TEXTO};
}}

QMainWindow, QDialog {{
    background-color: {COR_FUNDO_GERAL};
}}

/* Cabeçalho institucional */
#HeaderBar {{
    background-color: {COR_AZUL_MARINHO};
    border-bottom: 3px solid {COR_OURO_GOV};
}}
#HeaderTitle {{
    color: #FFFFFF;
    font-size: 13pt;
    font-weight: 700;
}}
#HeaderSubtitle {{
    color: #CBD5E1;
    font-size: 9pt;
}}

/* Sidebar */
#SidebarPanel {{
    background-color: #FFFFFF;
    border-right: 1px solid {COR_BORDA};
}}
#SidebarSectionTitle {{
    color: {COR_AZUL_MARINHO};
    font-size: 9.5pt;
    font-weight: 700;
    padding-top: 4px;
}}

/* Lista de cursos na sidebar */
QListWidget#CourseList {{
    background-color: #FFFFFF;
    border: 1px solid {COR_BORDA};
    border-radius: 4px;
    outline: none;
}}
QListWidget#CourseList::item {{
    padding: 8px 10px;
    border-bottom: 1px solid #EEF2F6;
}}
QListWidget#CourseList::item:selected {{
    background-color: #E8F0FE;
    color: {COR_AZUL_MARINHO};
    border-left: 3px solid {COR_AZUL_PRIMARIO};
    font-weight: 600;
}}
QListWidget#CourseList::item:hover:!selected {{
    background-color: #F8FAFC;
}}

/* Cards e painéis */
QFrame.Card {{
    background-color: {COR_SUPERFICIE};
    border: 1px solid {COR_BORDA};
    border-radius: 6px;
}}
QLabel.SectionHeader {{
    color: {COR_AZUL_EXCEL};
    font-size: 11pt;
    font-weight: 700;
}}
QLabel.MutedText {{
    color: {COR_TEXTO_SUAVE};
    font-size: 9pt;
}}

/* Inputs */
QLineEdit, QPlainTextEdit, QTextEdit, QSpinBox, QDoubleSpinBox, QComboBox {{
    background-color: #FFFFFF;
    border: 1px solid {COR_BORDA};
    border-radius: 4px;
    padding: 5px 8px;
    selection-background-color: {COR_AZUL_PRIMARIO};
}}
QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {{
    border: 1px solid {COR_AZUL_PRIMARIO};
}}

/* Botões */
QPushButton {{
    background-color: #FFFFFF;
    border: 1px solid {COR_BORDA};
    border-radius: 4px;
    padding: 6px 12px;
    font-weight: 600;
    color: {COR_TEXTO};
}}
QPushButton:hover {{
    background-color: #F1F5F9;
    border-color: #B6C4D4;
}}
QPushButton:disabled {{
    color: #94A3B8;
    background-color: #F8FAFC;
}}
QPushButton#PrimaryButton {{
    background-color: {COR_AZUL_PRIMARIO};
    border: 1px solid {COR_AZUL_PRIMARIO};
    color: #FFFFFF;
}}
QPushButton#PrimaryButton:hover {{
    background-color: {COR_AZUL_PRIMARIO_HOVER};
}}
QPushButton#PrimaryButton:disabled {{
    background-color: #93B4E8;
    border-color: #93B4E8;
}}
QPushButton#DangerButton {{
    background-color: #FFFFFF;
    border: 1px solid #FCA5A5;
    color: {COR_ERRO};
}}
QPushButton#DangerButton:hover {{
    background-color: {COR_ERRO_FUNDO};
}}

/* Abas principais */
QTabWidget::pane {{
    border: 1px solid {COR_BORDA};
    background-color: {COR_SUPERFICIE};
    border-radius: 4px;
}}
QTabBar::tab {{
    background-color: #E9EFF6;
    color: {COR_TEXTO_SUAVE};
    padding: 8px 14px;
    margin-right: 2px;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    border: 1px solid {COR_BORDA};
    border-bottom: none;
    font-weight: 600;
}}
QTabBar::tab:selected {{
    background-color: {COR_SUPERFICIE};
    color: {COR_AZUL_PRIMARIO};
    border-top: 3px solid {COR_AZUL_PRIMARIO};
}}
QTabBar::tab:hover:!selected {{
    background-color: #DDE7F3;
}}

/* Tabelas */
QTableWidget, QTableView {{
    background-color: #FFFFFF;
    alternate-background-color: #F8FAFC;
    border: 1px solid {COR_BORDA};
    gridline-color: #E2E8F0;
    selection-background-color: #DCE8FA;
    selection-color: {COR_AZUL_MARINHO};
}}
QHeaderView::section {{
    background-color: {COR_AZUL_EXCEL};
    color: #FFFFFF;
    padding: 6px 8px;
    border: none;
    border-right: 1px solid #2C6291;
    font-weight: 600;
}}

/* DropZone */
QFrame#DropZone {{
    background-color: #F8FAFC;
    border: 1px dashed #94A3B8;
    border-radius: 6px;
}}
QFrame#DropZone[state="loaded"] {{
    background-color: {COR_SUCESSO_FUNDO};
    border: 1px solid {COR_SUCESSO};
}}
QFrame#DropZone[dragging="true"] {{
    background-color: #E8F0FE;
    border: 2px dashed {COR_AZUL_PRIMARIO};
}}

/* Barra de progresso */
QProgressBar {{
    border: 1px solid {COR_BORDA};
    border-radius: 4px;
    background-color: #E2E8F0;
    text-align: center;
    height: 18px;
    font-size: 8.5pt;
    font-weight: 600;
}}
QProgressBar::chunk {{
    background-color: {COR_AZUL_PRIMARIO};
    border-radius: 3px;
}}
"""


def aplicar_tema_anm(app: QApplication) -> None:
    """Aplica o estilo plano institucional ANM/DIVCAM na aplicação Qt."""
    app.setStyle("Fusion")
    app.setStyleSheet(QSS_ANM)
