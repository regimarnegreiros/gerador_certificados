"""
Módulo para validação e auditoria de integridade da estrutura de pastas e arquivos do curso.
"""

from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any
import os
from core.caminhos import (
    garantir_diretorio,
    existe_com_fallback,
    listar_diretorio_com_fallback,
    caminho_estendido,
    verificar_profundidade_caminho,
    LIMITE_CARACTERES_DIRETORIO_CURSO,
)


PASTAS_OBRIGATORIAS_ENTRADA = [
    "lista_convidados",
    "listas_de_presenca",
    "lista_servidores",
    "modelo_certificado",
]

SUBPASTAS_SAIDA = [
    "relatorio_presencas",
    "certificados_pdf",
    "certificados_juntos_pdf",
    "certificados_juntos_ppt",
    "logs",
]


class ValidationError(Exception):
    """Exceção levantada quando a estrutura ou arquivos estão incorretos."""
    def __init__(self, erros: List[str], avisos: Optional[List[str]] = None):
        super().__init__("\n".join(erros))
        self.erros = erros
        self.avisos = avisos or []


def listar_arquivos_validos(pasta: Path, extensoes: Tuple[str, ...]) -> List[Path]:
    """Retorna arquivos com as extensões desejadas, ignorando temporários do Office (~$)."""
    if not existe_com_fallback(pasta):
        return []
    arquivos = []
    for item in listar_diretorio_com_fallback(pasta):
        if item.is_file():
            # Ignora arquivos temporários do Office e ocultos
            if item.name.startswith("~$") or item.name.startswith("."):
                continue
            if item.suffix.lower() in [ext.lower() for ext in extensoes]:
                arquivos.append(item)
    return sorted(arquivos)


def detectar_arquivos_temporarios(pasta_raiz: Path) -> List[Path]:
    """Detecta arquivos temporários do Office (~$*.xlsx, ~$*.pptx, etc.) abertos pelo usuário."""
    temporarios = []
    try:
        arvore = list(os.walk(pasta_raiz))
    except Exception:
        arvore = list(os.walk(caminho_estendido(pasta_raiz)))

    for raiz, _, arquivos in arvore:
        for arq in arquivos:
            if arq.startswith("~$"):
                temporarios.append(Path(raiz) / arq)
    return temporarios


def validar_estrutura_curso(diretorio_curso: Path) -> Dict[str, Any]:
    """
    Executa a auditoria completa da pasta do curso.
    Garante que:
    1. O diretório do curso e variaveis.txt existem.
    2. As pastas de entrada obrigatórias existem e não foram renomeadas.
    3. As pastas de saída existem (cria se necessário).
    4. Cada pasta de entrada possui os arquivos exigidos nos formatos corretos.
    5. Nenhum arquivo essencial está aberto no Office (arquivos ~$*).
    6. Identifica arquivos estranhos ou indevidos.
    """
    erros: List[str] = []
    avisos: List[str] = []

    if not existe_com_fallback(diretorio_curso):
        raise ValidationError([f"[ERRO] O diretório do curso não foi encontrado: {diretorio_curso}"])

    # 0. Checagem de profundidade do caminho do curso
    eh_profundo, tam, msg_orientacao = verificar_profundidade_caminho(diretorio_curso, LIMITE_CARACTERES_DIRETORIO_CURSO)
    if eh_profundo:
        erros.append(msg_orientacao)
        raise ValidationError(erros, avisos)

    # 1. Checa variaveis.txt
    caminho_variaveis = diretorio_curso / "variaveis.txt"
    if not existe_com_fallback(caminho_variaveis):
        erros.append("[ERRO] O arquivo 'variaveis.txt' não foi encontrado na pasta do curso.")

    # 2. Checa pastas de entrada
    for pasta_nome in PASTAS_OBRIGATORIAS_ENTRADA:
        p = diretorio_curso / pasta_nome
        if not existe_com_fallback(p):
            erros.append(f"[ERRO] A pasta obrigatória '{pasta_nome}' não foi encontrada (pode ter sido renomeada ou apagada).")

    # 3. Garante estrutura de saída
    pasta_saida = diretorio_curso / "saida"
    garantir_diretorio(pasta_saida)
    for sub in SUBPASTAS_SAIDA:
        garantir_diretorio(pasta_saida / sub)

    # 4. Alerta sobre arquivos temporários do Office
    temporarios = detectar_arquivos_temporarios(diretorio_curso)
    if temporarios:
        for temp in temporarios:
            avisos.append(
                f"[AVISO] ATENÇÃO: O arquivo temporário '{temp.name}' foi detectado.\n"
                f"   Isso geralmente significa que a planilha ou apresentação está aberta no Excel/PowerPoint.\n"
                f"   Por favor, salve e feche todos os arquivos do Office antes de prosseguir."
            )

    # Se já há erros estruturais graves de pastas inexistentes, interrompe antes de checar arquivos internos
    if erros:
        raise ValidationError(erros, avisos)

    # 5. Validação detalhada por pasta de entrada
    arquivos_encontrados: Dict[str, Any] = {}

    # 5.1 Lista de convidados (deve ter 1 arquivo CSV)
    pasta_convidados = diretorio_curso / "lista_convidados"
    csvs_convidados = listar_arquivos_validos(pasta_convidados, (".csv",))
    outros_convidados = [f for f in listar_diretorio_com_fallback(pasta_convidados) if f.is_file() and not f.name.startswith("~$") and f.suffix.lower() != ".csv"]

    if len(csvs_convidados) == 0:
        erros.append("[ERRO] Pasta 'lista_convidados': Nenhum arquivo .csv encontrado. Insira o relatório de convidados do Teams.")
    elif len(csvs_convidados) > 1:
        avisos.append(f"[AVISO] Pasta 'lista_convidados': Mais de um CSV encontrado. Será utilizado '{csvs_convidados[0].name}'.")
    if outros_convidados:
        nomes_outros = ", ".join([f.name for f in outros_convidados])
        avisos.append(f"[AVISO] Pasta 'lista_convidados': Arquivos indevidos ou em formato incorreto encontrados: {nomes_outros}")

    if csvs_convidados:
        arquivos_encontrados["convidados"] = csvs_convidados[0]

    # 5.2 Lista de servidores (deve ter 1 arquivo Excel .xlsx)
    pasta_servidores = diretorio_curso / "lista_servidores"
    xlsxs_servidores = listar_arquivos_validos(pasta_servidores, (".xlsx", ".xls"))
    outros_servidores = [f for f in listar_diretorio_com_fallback(pasta_servidores) if f.is_file() and not f.name.startswith("~$") and f.suffix.lower() not in (".xlsx", ".xls")]

    if len(xlsxs_servidores) == 0:
        erros.append("[ERRO] Pasta 'lista_servidores': Nenhuma planilha Excel (.xlsx) encontrada. Insira a lista de servidores da ANM.")
    elif len(xlsxs_servidores) > 1:
        avisos.append(f"[AVISO] Pasta 'lista_servidores': Mais de uma planilha encontrada. Será utilizada '{xlsxs_servidores[0].name}'.")
    if outros_servidores:
        nomes_outros = ", ".join([f.name for f in outros_servidores])
        avisos.append(f"[AVISO] Pasta 'lista_servidores': Arquivos indevidos encontrados: {nomes_outros}")

    if xlsxs_servidores:
        arquivos_encontrados["servidores"] = xlsxs_servidores[0]

    # 5.3 Listas de presença (deve ter ao menos 1 arquivo .csv ou .tsv)
    pasta_presencas = diretorio_curso / "listas_de_presenca"
    arquivos_presencas = listar_arquivos_validos(pasta_presencas, (".csv", ".tsv", ".txt"))
    if len(arquivos_presencas) == 0:
        erros.append("[ERRO] Pasta 'listas_de_presenca': Nenhuma lista de presença (.csv/.tsv) encontrada.")
    else:
        from core.presenca import ordenar_arquivos_presenca
        arquivos_presencas = ordenar_arquivos_presenca(arquivos_presencas)
    arquivos_encontrados["presencas"] = arquivos_presencas

    # 5.4 Modelo de certificado (deve ter 1 arquivo .pptx ou .pptm)
    pasta_modelo = diretorio_curso / "modelo_certificado"
    modelos_ppt = listar_arquivos_validos(pasta_modelo, (".pptx", ".pptm"))
    outros_modelos = [f for f in listar_diretorio_com_fallback(pasta_modelo) if f.is_file() and not f.name.startswith("~$") and f.suffix.lower() not in (".pptx", ".pptm")]

    if len(modelos_ppt) == 0:
        erros.append("[ERRO] Pasta 'modelo_certificado': Nenhum modelo PowerPoint (.pptx ou .pptm) encontrado.")
    elif len(modelos_ppt) > 1:
        avisos.append(f"[AVISO] Pasta 'modelo_certificado': Mais de um modelo encontrado. Será utilizado '{modelos_ppt[0].name}'.")
    if outros_modelos:
        nomes_outros = ", ".join([f.name for f in outros_modelos])
        avisos.append(f"[AVISO] Pasta 'modelo_certificado': Arquivos indevidos encontrados: {nomes_outros}")

    if modelos_ppt:
        arquivos_encontrados["modelo_ppt"] = modelos_ppt[0]

    if erros:
        raise ValidationError(erros, avisos)

    return {
        "diretorio_curso": diretorio_curso,
        "arquivos": arquivos_encontrados,
        "avisos": avisos,
    }
