"""
Módulo central de diretórios gerenciados pelo aplicativo (Gerador de Certificados ANM).
Define a pasta raiz em caminho curto (C:\\ANM_Certificados\\ com fallback em %LOCALAPPDATA%),
os recursos globais compartilhados e a estrutura isolada de cada curso.
"""

import os
from pathlib import Path
from typing import Optional, Union

from core.caminhos import existe_com_fallback, garantir_diretorio


NOME_PASTA_APP = "ANM_Certificados"
CAMINHO_RAIZ_PADRAO = Path("C:/ANM_Certificados")
ENV_RAIZ_CUSTOM = "ANM_CERTIFICADOS_HOME"

SUBPASTAS_SAIDA_CURSO = [
    "relatorio_presencas",
    "certificados_pdf",
    "certificados_juntos_pdf",
    "certificados_juntos_ppt",
    "logs",
]


def diretorio_gravavel(pasta: Path) -> bool:
    """Verifica se é possível criar e gravar arquivos no diretório informado."""
    try:
        garantir_diretorio(pasta)
        probe = pasta / ".anm_write_probe"
        with open(probe, "w", encoding="utf-8") as f:
            f.write("ok")
        probe.unlink(missing_ok=True)
        return True
    except Exception:
        return False


def obter_raiz_dados(raiz_custom: Optional[Union[str, Path]] = None) -> Path:
    """
    Resolve o diretório raiz onde o aplicativo armazena configurações,
    recursos globais e cursos.
    Prioridade:
    1. Parâmetro explícito `raiz_custom`
    2. Variável de ambiente `ANM_CERTIFICADOS_HOME`
    3. `C:\\ANM_Certificados` (caminho curto fora de AppData, ideal para TrendAI Vision One)
    4. Fallback em `%LOCALAPPDATA%\\ANM_Certificados`
    """
    if raiz_custom is not None:
        return Path(raiz_custom)

    env_path = os.environ.get(ENV_RAIZ_CUSTOM, "").strip()
    if env_path:
        return Path(env_path)

    if os.name == "nt":
        if diretorio_gravavel(CAMINHO_RAIZ_PADRAO):
            return CAMINHO_RAIZ_PADRAO
        local_app = os.environ.get("LOCALAPPDATA", "").strip()
        base = Path(local_app) if local_app else Path.home()
        return base / NOME_PASTA_APP

    return Path.home() / NOME_PASTA_APP


def obter_caminho_app_config(raiz: Path) -> Path:
    """Retorna o caminho do arquivo app_config.json na raiz."""
    return Path(raiz) / "app_config.json"


def obter_pasta_global(raiz: Path) -> Path:
    """Retorna a pasta de recursos globais compartilhados entre cursos."""
    return Path(raiz) / "global"


def obter_pasta_cursos(raiz: Path) -> Path:
    """Retorna a pasta onde ficam os diretórios individuais de cursos."""
    return Path(raiz) / "cursos"


def obter_caminho_servidores_global(raiz: Path) -> Path:
    """Retorna o caminho esperado da planilha global de servidores."""
    return obter_pasta_global(raiz) / "servidores.xlsx"


def localizar_modelo_global(raiz: Path) -> Optional[Path]:
    """Localiza o modelo de certificado global (.pptx ou .pptm), se existir."""
    pasta_global = obter_pasta_global(raiz)
    for ext in (".pptx", ".pptm"):
        candidato = pasta_global / f"modelo_certificado{ext}"
        if existe_com_fallback(candidato) and candidato.is_file():
            return candidato
    return None


def garantir_estrutura_raiz(raiz: Optional[Union[str, Path]] = None) -> Path:
    """Garante que a raiz de dados e as subpastas global/ e cursos/ existam."""
    raiz_path = obter_raiz_dados(raiz)
    garantir_diretorio(raiz_path)
    garantir_diretorio(obter_pasta_global(raiz_path))
    garantir_diretorio(obter_pasta_cursos(raiz_path))
    return raiz_path


def obter_caminho_curso_json(diretorio_curso: Path) -> Path:
    """Retorna o caminho do arquivo curso.json dentro da pasta do curso."""
    return Path(diretorio_curso) / "curso.json"


def obter_pasta_entrada_curso(diretorio_curso: Path) -> Path:
    """Retorna a pasta entrada/ do curso."""
    return Path(diretorio_curso) / "entrada"


def obter_pasta_presencas_curso(diretorio_curso: Path) -> Path:
    """Retorna a pasta entrada/presencas/ do curso."""
    return obter_pasta_entrada_curso(diretorio_curso) / "presencas"


def obter_pasta_saida_curso(diretorio_curso: Path) -> Path:
    """Retorna a pasta saida/ do curso."""
    return Path(diretorio_curso) / "saida"


def garantir_estrutura_curso(diretorio_curso: Path) -> Path:
    """Cria todas as subpastas internas necessárias dentro de um curso."""
    curso_path = Path(diretorio_curso)
    garantir_diretorio(curso_path)
    garantir_diretorio(obter_pasta_entrada_curso(curso_path))
    garantir_diretorio(obter_pasta_presencas_curso(curso_path))
    pasta_saida = obter_pasta_saida_curso(curso_path)
    garantir_diretorio(pasta_saida)
    for sub in SUBPASTAS_SAIDA_CURSO:
        garantir_diretorio(pasta_saida / sub)
    return curso_path
