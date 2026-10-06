"""
Módulo central para manipulação robusta de caminhos de arquivos no Windows.
Resolve limitações de MAX_PATH (260 caracteres) através de caminhos estendidos (\\\\?\\)
e oferece helpers seguros com fallback automático para leitura, gravação e cópia.
"""

import os
import sys
import re
import shutil
from pathlib import Path
from typing import Union, List, Optional, Tuple, Any
from contextlib import contextmanager


# Limites recomendados de caminho para estabilidade no Windows e Office (PowerPoint/Excel)
LIMITE_CARACTERES_DIRETORIO_CURSO = 130
LIMITE_CARACTERES_CAMINHO_FINAL = 230


def verificar_profundidade_caminho(
    caminho: Union[str, Path],
    limite: int = LIMITE_CARACTERES_DIRETORIO_CURSO,
) -> Tuple[bool, int, str]:
    """
    Verifica se um caminho é excessivamente longo ou profundo.
    Retorna uma tupla (eh_profundo, tamanho, mensagem_orientacao).
    Caso não seja profundo, mensagem_orientacao é vazia ("").
    """
    if caminho is None:
        return False, 0, ""

    try:
        caminho_res = Path(caminho).resolve()
        caminho_str = str(caminho_res)
    except Exception:
        caminho_str = str(caminho)

    tamanho = len(caminho_str)
    if tamanho > limite:
        msg = obter_mensagem_orientacao_caminho_profundo(caminho_str, tamanho=tamanho, limite=limite)
        return True, tamanho, msg

    return False, tamanho, ""


def obter_mensagem_orientacao_caminho_profundo(
    caminho: Union[str, Path],
    detalhe: str = "",
    tamanho: Optional[int] = None,
    limite: Optional[int] = None,
) -> str:
    """
    Gera texto de orientação amigável para o usuário mover a pasta
    do curso para um diretório menos profundo.
    """
    caminho_str = str(caminho)
    tam = tamanho if tamanho is not None else len(caminho_str)
    info_limite = f" (limite recomendado: {limite} caracteres)" if limite else ""

    linhas = [
        "ATENÇÃO: DIRETÓRIO MUITO PROFUNDO OU CAMINHO MUITO LONGO",
        f"O caminho atual possui {tam} caracteres{info_limite}:",
        f"  '{caminho_str}'",
        "",
        "O Windows e os programas do Microsoft Office (PowerPoint e Excel) possuem",
        "limitações rígidas no tamanho dos caminhos de arquivos. Em pastas muito",
        "profundas, a geração de certificados e relatórios pode falhar ao salvar.",
        "",
        "💡 ORIENTAÇÃO:",
        "Por favor, mova a pasta do curso para um diretório menos profundo,",
        "mais próximo da raiz do seu disco (por exemplo, na unidade C:).",
        "",
        "Exemplos de locais recomendados:",
        "  • C:\\Cursos\\",
        "  • C:\\Treinamentos\\",
        "  • C:\\ANM\\",
    ]
    if detalhe:
        linhas.append(f"\nDetalhe adicional: {detalhe}")

    return "\n".join(linhas)


def esvaziar_pasta(pasta: Union[str, Path], extensoes: Optional[Tuple[str, ...]] = None) -> int:
    """
    Remove todos os arquivos de um diretório especificado antes de gerar novos arquivos.
    Evita que arquivos antigos de execuções anteriores permaneçam e causem inconsistências
    (por exemplo, juntar certificados de participantes que foram removidos do relatório).

    Se 'extensoes' for fornecido, remove apenas arquivos com as extensões indicadas.
    Caso contrário, remove todos os arquivos da pasta.
    Levanta PermissionError com mensagem clara se algum arquivo estiver aberto em outro programa.
    Retorna a quantidade de arquivos removidos.
    """
    p = Path(pasta)
    if not existe_com_fallback(p) or not p.is_dir():
        return 0

    removidos = 0
    itens = listar_diretorio_com_fallback(p)
    for item in itens:
        if item.is_file():
            if extensoes is None or item.suffix.lower() in [ext.lower() for ext in extensoes]:
                try:
                    item.unlink()
                    removidos += 1
                except PermissionError:
                    raise PermissionError(
                        f"Não foi possível remover o arquivo anterior '{item.name}' da pasta '{p.name}'.\n"
                        f"O arquivo pode estar aberto no Excel, PowerPoint ou visualizador de PDF.\n"
                        f"Por favor, feche todos os programas que estejam utilizando arquivos desta pasta e tente novamente."
                    )
                except Exception:
                    try:
                        Path(caminho_estendido(item)).unlink()
                        removidos += 1
                    except Exception:
                        pass
    return removidos



def caminho_estendido(caminho: Union[str, Path, None]) -> str:
    """
    Retorna o caminho no formato estendido do Windows (prefixo \\\\?\\),
    desativando a restrição de MAX_PATH (260 caracteres) e permitindo
    caminhos de até 32.767 caracteres no subsistema Win32.
    Em sistemas não-Windows, retorna o caminho absoluto normal resolvido.
    """
    if caminho is None:
        return ""

    caminho_str = str(caminho).strip()
    if not caminho_str:
        return ""

    if os.name != "nt":
        try:
            return str(Path(caminho_str).resolve())
        except Exception:
            return os.path.abspath(caminho_str)

    # Se já possui prefixo estendido, apenas normaliza barras invertidas
    if caminho_str.startswith("\\\\?\\") or caminho_str.startswith("//?/"):
        return "\\\\?\\" + caminho_str[4:].replace("/", "\\")

    try:
        abs_str = str(Path(caminho_str).resolve())
    except Exception:
        abs_str = os.path.abspath(caminho_str)

    abs_str = abs_str.replace("/", "\\")

    # Tratamento para caminhos de rede UNC (\\\\servidor\\compartilhamento\\...)
    if abs_str.startswith("\\\\"):
        return "\\\\?\\UNC\\" + abs_str[2:]

    return "\\\\?\\" + abs_str


def caminho_estendido_path(caminho: Union[str, Path]) -> Path:
    """Retorna um objeto Path apontando para o caminho estendido."""
    return Path(caminho_estendido(caminho))


def obter_caminho_curto_windows(caminho: Union[str, Path]) -> str:
    """
    Retorna o caminho no formato 8.3 do Windows (GetShortPathNameW),
    permitindo que programas com limite de 255 caracteres (como a API COM do PowerPoint)
    acessem pastas em caminhos longos sem falhas.
    """
    if os.name != "nt":
        return str(caminho)
    try:
        import ctypes
        buffer = ctypes.create_unicode_buffer(1024)
        caminho_str = str(Path(caminho).resolve())
        res = ctypes.windll.kernel32.GetShortPathNameW(caminho_str, buffer, 1024)
        if res > 0:
            return buffer.value
    except Exception:
        pass
    return str(caminho)


def garantir_diretorio(caminho: Union[str, Path]) -> Path:
    """
    Cria todos os diretórios do caminho especificado (equivalente a mkdir -p).
    Possui fallback automático para caminho estendido caso ocorra estouro de MAX_PATH.
    """
    p = Path(caminho)
    try:
        p.mkdir(parents=True, exist_ok=True)
        return p
    except (FileNotFoundError, OSError):
        ext_p = Path(caminho_estendido(p))
        ext_p.mkdir(parents=True, exist_ok=True)
        return ext_p


@contextmanager
def abrir_arquivo(caminho: Union[str, Path], mode: str = "r", **kwargs):
    """
    Context manager seguro para abertura de arquivos (open).
    Tenta primeiramente o caminho padrão; caso ocorra FileNotFoundError ou OSError
    provocado por caminho longo no Windows, faz fallback automático para o caminho estendido (\\\\?\\).
    """
    f = None
    try:
        f = open(caminho, mode=mode, **kwargs)
        yield f
    except (FileNotFoundError, OSError):
        if f is not None and not f.closed:
            try:
                f.close()
            except Exception:
                pass
        ext = caminho_estendido(caminho)
        f = open(ext, mode=mode, **kwargs)
        yield f
    finally:
        if f is not None and not f.closed:
            try:
                f.close()
            except Exception:
                pass


def copiar_arquivo_com_fallback(origem: Union[str, Path], destino: Union[str, Path]) -> Path:
    """
    Copia um arquivo com shutil.copy2.
    Caso ocorra erro de caminho longo no Windows, tenta novamente utilizando caminho estendido.
    """
    try:
        res = shutil.copy2(origem, destino)
        return Path(res)
    except (FileNotFoundError, OSError):
        ext_origem = caminho_estendido(origem)
        ext_destino = caminho_estendido(destino)
        res = shutil.copy2(ext_origem, ext_destino)
        return Path(res)


def existe_com_fallback(caminho: Union[str, Path]) -> bool:
    """
    Verifica se um arquivo ou diretório existe.
    Possui fallback para o formato estendido caso o caminho exceda MAX_PATH.
    """
    try:
        if Path(caminho).exists():
            return True
    except Exception:
        pass

    try:
        return os.path.exists(caminho_estendido(caminho))
    except Exception:
        return False


def listar_diretorio_com_fallback(pasta: Union[str, Path]) -> List[Path]:
    """
    Lista os itens de um diretório com tolerância a caminhos longos.
    """
    p = Path(pasta)
    try:
        return list(p.iterdir())
    except (FileNotFoundError, OSError):
        ext_p = Path(caminho_estendido(pasta))
        return list(ext_p.iterdir())


def ler_excel_com_fallback(caminho: Union[str, Path], **kwargs):
    """
    Lê uma planilha Excel utilizando pandas com fallback automático para caminho estendido.
    Importa pandas sob demanda para não onerar scripts que não utilizam DataFrames.
    """
    import pandas as pd

    try:
        return pd.read_excel(caminho, **kwargs)
    except (FileNotFoundError, OSError):
        ext = caminho_estendido(caminho)
        return pd.read_excel(ext, **kwargs)


def sanitizar_nome_arquivo(nome: Any) -> str:
    """
    Higieniza o nome para ser um nome de arquivo válido no Windows,
    substituindo separadores comuns (: / \\ |) por hífens espaçados,
    removendo caracteres proibidos (* ? " < > : e controle) e normalizando espaços.
    NÃO trunca o nome para garantir que nomes completos sejam preservados integralmente.
    """
    texto = str(nome)
    # Substitui separadores comuns por hífen espaçado
    texto = re.sub(r"\s*[:/\\|]+\s*", " - ", texto)
    # Remove caracteres inválidos no Windows: < > " ? * : e caracteres de controle (0-31 e 127)
    texto = re.sub(r'[*?"<>:\x00-\x1f\x7f]', "", texto)
    # Colapsa múltiplos hífens consecutivos (ex: " - - " -> " - ")
    texto = re.sub(r"(?:\s*-\s*)+", " - ", texto)
    # Colapsa múltiplos espaços em um único
    texto = re.sub(r"\s+", " ", texto)
    # Remove espaços, hífens, underscores ou pontos nas pontas
    texto = texto.strip(" .-_")
    return texto if texto else "arquivo"


def encurtar_caminho_se_necessario(caminho: Path, max_len: int = 245) -> Path:
    """
    Mantido para compatibilidade. Caso seja estritamente necessário encurtar um caminho
    (por exemplo em sistemas legados), encurta o stem mantendo a extensão.
    """
    caminho = Path(caminho)
    try:
        caminho_str = str(caminho.resolve() if caminho.is_absolute() else caminho)
    except Exception:
        caminho_str = str(caminho)

    if len(caminho_str) <= max_len:
        return caminho

    pasta = caminho.parent
    sufixo = caminho.suffix
    try:
        tamanho_pasta = len(str(pasta.resolve() if pasta.is_absolute() else pasta))
    except Exception:
        tamanho_pasta = len(str(pasta))

    espaco_restante = max_len - tamanho_pasta - 1 - len(sufixo)
    if espaco_restante >= 10:
        novo_stem = caminho.stem[:espaco_restante].rstrip(" .-_")
        return pasta / f"{novo_stem}{sufixo}"
    else:
        return pasta / f"arquivo{sufixo}"


def reduzir_nome_curso_para_arquivo(nome_curso: str, max_chars: int = 30) -> str:
    """
    Reduz o nome de um curso para ser usado em nomes de arquivos seguros e curtos,
    removendo preposições e palavras de ligação e limitando o tamanho total.
    Exemplo:
        'Hidrologia Aplicada à Segurança de Barragens e Pilhas de Mineração'
        -> 'Hidrologia_Aplicada_Seguranca'
    """
    if not nome_curso:
        return "curso"
    texto = sanitizar_nome_arquivo(nome_curso)
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", texto)
    sem_acento = "".join(c for c in nfkd if not unicodedata.combining(c))
    palavras = re.findall(r"[A-Za-z0-9]+", sem_acento)
    stop_words = {"de", "da", "do", "das", "dos", "e", "a", "o", "as", "os", "em", "para", "com", "por", "no", "na", "nos", "nas", "ao", "aos"}
    palavras_filtradas = [p for p in palavras if p.lower() not in stop_words]
    if not palavras_filtradas:
        palavras_filtradas = palavras

    resultado = []
    tamanho_atual = 0
    for p in palavras_filtradas:
        tam_novo = tamanho_atual + (1 if resultado else 0) + len(p)
        if tam_novo <= max_chars or not resultado:
            resultado.append(p)
            tamanho_atual = tam_novo
        else:
            break

    res_str = "_".join(resultado)
    return res_str[:max_chars].rstrip("_") if res_str else "curso"


def obter_nome_arquivo_junto(
    nome_curso: str,
    pasta_destino: Union[str, Path],
    extensao: str = ".pdf",
    max_caminho_total: int = 215,
) -> str:
    """
    Gera um nome de arquivo consolidado encurtado dinamicamente para garantir que
    o caminho completo final nunca ultrapasse o limite seguro do Office/Windows (230 caracteres).
    """
    pasta_str = str(Path(pasta_destino).resolve())
    tam_pasta = len(pasta_str) + 1  # inclui o separador de pasta
    ext = extensao if extensao.startswith(".") else f".{extensao}"

    # Espaço disponível para o nome do arquivo respeitando o teto seguro
    espaco_disponivel = max(15, max_caminho_total - tam_pasta)

    prefixo = "cert_"
    sufixo = f"_todos_juntos{ext}"
    tamanho_fixo = len(prefixo) + len(sufixo)

    espaco_para_curso = espaco_disponivel - tamanho_fixo
    if espaco_para_curso >= 8:
        max_curso = min(30, espaco_para_curso)
        curso_curto = reduzir_nome_curso_para_arquivo(nome_curso, max_chars=max_curso)
        nome_arq = f"{prefixo}{curso_curto}{sufixo}"
    else:
        # Se a pasta for extremamente profunda, usa o menor nome possível
        nome_arq = f"cert_todos_juntos{ext}"

    return nome_arq
