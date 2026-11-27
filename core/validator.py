"""
Módulo para validação e auditoria de integridade da estrutura interna do curso
e dos recursos globais compartilhados.
"""

import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from core.app_paths import (
    SUBPASTAS_SAIDA_CURSO,
    localizar_modelo_global,
    obter_caminho_curso_json,
    obter_caminho_servidores_global,
    obter_pasta_entrada_curso,
    obter_pasta_global,
    obter_pasta_presencas_curso,
    obter_pasta_saida_curso,
    obter_raiz_dados,
)
from core.caminhos import (
    LIMITE_CARACTERES_DIRETORIO_CURSO,
    caminho_estendido,
    existe_com_fallback,
    garantir_diretorio,
    listar_diretorio_com_fallback,
    verificar_profundidade_caminho,
)


SUBPASTAS_SAIDA = SUBPASTAS_SAIDA_CURSO


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
    ext_lower = [ext.lower() for ext in extensoes]
    arquivos = []
    for item in listar_diretorio_com_fallback(pasta):
        if item.is_file():
            if item.name.startswith("~$") or item.name.startswith("."):
                continue
            if item.suffix.lower() in ext_lower:
                arquivos.append(item)
    return sorted(arquivos)


def detectar_arquivos_temporarios(pasta_raiz: Path) -> List[Path]:
    """Detecta arquivos temporários do Office (~$*.xlsx, ~$*.pptx, etc.) abertos pelo usuário."""
    if not existe_com_fallback(pasta_raiz):
        return []
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


def _resolver_raiz_para_validacao(
    diretorio_curso: Path,
    raiz_dados: Optional[Path] = None,
) -> Path:
    if raiz_dados is not None:
        return Path(raiz_dados)
    if diretorio_curso.parent.name == "cursos":
        return diretorio_curso.parent.parent
    return obter_raiz_dados()


def validar_estrutura_curso(
    diretorio_curso: Path,
    raiz_dados: Optional[Path] = None,
) -> Dict[str, Any]:
    """
    Executa a auditoria completa do curso e dos recursos globais:
    1. O diretório do curso e curso.json existem (dentro do limite de caminho).
    2. Os recursos globais (servidores.xlsx e modelo_certificado.pptx/.pptm) estão disponíveis.
    3. As entradas do curso (convidados .csv e >=1 lista de presença) estão presentes.
    4. As subpastas de saída existem.
    5. Identifica arquivos temporários (~$) abertos no Office.
    """
    diretorio_curso = Path(diretorio_curso)
    erros: List[str] = []
    avisos: List[str] = []

    if not existe_com_fallback(diretorio_curso):
        raise ValidationError([f"[ERRO] O diretório do curso não foi encontrado: {diretorio_curso}"])

    eh_profundo, _, msg_orientacao = verificar_profundidade_caminho(
        diretorio_curso,
        LIMITE_CARACTERES_DIRETORIO_CURSO,
    )
    if eh_profundo:
        raise ValidationError([msg_orientacao], avisos)

    raiz_path = _resolver_raiz_para_validacao(diretorio_curso, raiz_dados)

    # 1. Verifica curso.json
    caminho_json = obter_caminho_curso_json(diretorio_curso)
    if not existe_com_fallback(caminho_json):
        erros.append("[ERRO] O arquivo 'curso.json' não foi encontrado na pasta do curso.")

    # 2. Verifica pastas de entrada do curso
    pasta_entrada = obter_pasta_entrada_curso(diretorio_curso)
    pasta_presencas = obter_pasta_presencas_curso(diretorio_curso)
    if not existe_com_fallback(pasta_entrada):
        erros.append("[ERRO] A pasta 'entrada' do curso não foi encontrada.")
    if not existe_com_fallback(pasta_presencas):
        erros.append("[ERRO] A pasta 'entrada/presencas' do curso não foi encontrada.")

    # 3. Garante estrutura de saída
    pasta_saida = obter_pasta_saida_curso(diretorio_curso)
    garantir_diretorio(pasta_saida)
    for sub in SUBPASTAS_SAIDA_CURSO:
        garantir_diretorio(pasta_saida / sub)

    # 4. Verifica arquivos temporários do Office no curso e nos recursos globais
    temporarios = detectar_arquivos_temporarios(diretorio_curso)
    pasta_global = obter_pasta_global(raiz_path)
    temporarios.extend(detectar_arquivos_temporarios(pasta_global))
    for temp in temporarios:
        avisos.append(
            f"[AVISO] O arquivo temporário '{temp.name}' foi detectado. "
            "Feche o Excel ou PowerPoint antes de continuar para evitar falhas de leitura/gravação."
        )

    if erros:
        raise ValidationError(erros, avisos)

    arquivos_encontrados: Dict[str, Any] = {}

    # 5.1 Recursos globais
    caminho_servidores = obter_caminho_servidores_global(raiz_path)
    if not existe_com_fallback(caminho_servidores):
        erros.append("[ERRO] A lista global de servidores (servidores.xlsx) ainda não foi cadastrada na interface.")
    else:
        arquivos_encontrados["servidores"] = caminho_servidores

    caminho_modelo = localizar_modelo_global(raiz_path)
    if caminho_modelo is None:
        erros.append("[ERRO] O modelo global de certificado (.pptx ou .pptm) ainda não foi cadastrado na interface.")
    else:
        arquivos_encontrados["modelo_ppt"] = caminho_modelo

    # 5.2 Lista de convidados (.csv em entrada/)
    csvs_convidados = listar_arquivos_validos(pasta_entrada, (".csv",))
    if not csvs_convidados:
        erros.append("[ERRO] Lista de convidados (.csv) não foi importada para este curso.")
    else:
        arquivos_encontrados["convidados"] = csvs_convidados[0]

    # 5.3 Listas de presença (.csv, .tsv, .txt em entrada/presencas/)
    arquivos_presencas = listar_arquivos_validos(pasta_presencas, (".csv", ".tsv", ".txt"))
    if not arquivos_presencas:
        erros.append("[ERRO] Nenhuma lista de presença (.csv/.tsv) foi importada para este curso.")
    else:
        from core.presenca import ordenar_arquivos_presenca

        arquivos_presencas = ordenar_arquivos_presenca(arquivos_presencas)
    arquivos_encontrados["presencas"] = arquivos_presencas

    if erros:
        raise ValidationError(erros, avisos)

    return {
        "diretorio_curso": diretorio_curso,
        "raiz_dados": raiz_path,
        "arquivos": arquivos_encontrados,
        "avisos": avisos,
    }
