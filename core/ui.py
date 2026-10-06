"""
Módulo de Interface de Linha de Comando (CLI) Amigável - ANM.
Fornece suporte a cores ANSI, limpeza de tela, cabeçalhos padronizados,
mensagens de status imediatas e pausas amigáveis.
Compatível com todos os terminais e encodings do Windows (UTF-8, CP1252, CP850).
"""

import os
import sys

# ============================================================================
# CÓDIGOS ANSI DE CORES E ESTILOS
# ============================================================================
RESET = "\033[0m"
BOLD = "\033[1m"
DIM = "\033[2m"
ITALIC = "\033[3m"
UNDERLINE = "\033[4m"

# Cores de primeiro plano (Texto)
FG_BLACK = "\033[30m"
FG_RED = "\033[91m"
FG_GREEN = "\033[92m"
FG_YELLOW = "\033[93m"
FG_BLUE = "\033[94m"
FG_MAGENTA = "\033[95m"
FG_CYAN = "\033[96m"
FG_WHITE = "\033[97m"

# Cores de fundo
BG_DARK = "\033[40m"
BG_BLUE = "\033[44m"
BG_CYAN = "\033[46m"


def inicializar_terminal():
    """
    Configura codificação UTF-8 no stdout/stderr e habilita
    o processamento de sequências virtuais de terminal (ANSI) no Windows.
    """
    if sys.stdout and hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
    if sys.stderr and hasattr(sys.stderr, "reconfigure"):
        try:
            sys.stderr.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    if os.name == "nt":
        # Ativa o modo ANSI através da chamada ao shell do Windows
        os.system("")
        # Garante a ativação via API do Windows (ENABLE_VIRTUAL_TERMINAL_PROCESSING)
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            hStdOut = kernel32.GetStdHandle(-11)
            mode = ctypes.c_ulong()
            kernel32.GetConsoleMode(hStdOut, ctypes.byref(mode))
            mode.value |= 0x0004  # ENABLE_VIRTUAL_TERMINAL_PROCESSING
            kernel32.SetConsoleMode(hStdOut, mode)
        except Exception:
            pass


def seguro_print(texto: str = "", **kwargs):
    """
    Imprime texto garantindo que qualquer caractere não suportado
    pelo encoding atual do terminal seja tratado sem lançar exceções.
    """
    try:
        print(texto, **kwargs)
    except UnicodeEncodeError:
        try:
            enc = getattr(sys.stdout, "encoding", None) or "ascii"
            texto_limpo = texto.encode(enc, errors="replace").decode(enc)
            print(texto_limpo, **kwargs)
        except Exception:
            # Fallback definitivo em ASCII
            print(texto.encode("ascii", errors="replace").decode("ascii"), **kwargs)


def limpar_tela():
    """Limpa o console de forma compatível com Windows e Unix."""
    os.system("cls" if os.name == "nt" else "clear")


# ============================================================================
# FORMATAÇÃO DE TEXTOS COM CORES
# ============================================================================
def verde(texto: str) -> str:
    return f"{FG_GREEN}{texto}{RESET}"


def vermelho(texto: str) -> str:
    return f"{FG_RED}{texto}{RESET}"


def amarelo(texto: str) -> str:
    return f"{FG_YELLOW}{texto}{RESET}"


def ciano(texto: str) -> str:
    return f"{FG_CYAN}{texto}{RESET}"


def azul(texto: str) -> str:
    return f"{FG_BLUE}{texto}{RESET}"


def negrito(texto: str) -> str:
    return f"{BOLD}{texto}{RESET}"


def dim(texto: str) -> str:
    return f"{DIM}{texto}{RESET}"


# ============================================================================
# COMPONENTES VISUAIS DA INTERFACE
# ============================================================================
def exibir_mensagem_aguarde(texto: str = "Processando... Aguarde um momento!"):
    """Exibe mensagem de processamento em destaque imediato com descarga do buffer."""
    seguro_print(f"\n  {FG_YELLOW}{BOLD}⏳ {texto}{RESET}\n", flush=True)


def exibir_cabecalho(titulo: str, subtitulo: str = None, largura: int = 66):
    """Exibe cabeçalho elegante e centralizado com bordas estilizadas."""
    # Testa se o terminal suporta bordas unicode duplas
    pode_unicode = True
    try:
        "╔═╗║╚╝".encode(getattr(sys.stdout, "encoding", "ascii") or "ascii")
    except (UnicodeEncodeError, LookupError):
        pode_unicode = False

    if pode_unicode:
        c_tl, c_tr, c_bl, c_br, c_h, c_v = "╔", "╗", "╚", "╝", "═", "║"
    else:
        c_tl, c_tr, c_bl, c_br, c_h, c_v = "+", "+", "+", "+", "=", "|"

    borda_superior = f"{FG_CYAN}{c_tl}{c_h * (largura - 2)}{c_tr}{RESET}"
    borda_inferior = f"{FG_CYAN}{c_bl}{c_h * (largura - 2)}{c_br}{RESET}"

    seguro_print(borda_superior)

    # Linha do título centralizado
    espaco_util = largura - 4
    tit_trunc = titulo[:espaco_util]
    margem_esq = (espaco_util - len(tit_trunc)) // 2
    margem_dir = espaco_util - len(tit_trunc) - margem_esq
    seguro_print(f"{FG_CYAN}{c_v} {BOLD}{FG_WHITE}{' ' * margem_esq}{tit_trunc}{' ' * margem_dir}{RESET}{FG_CYAN} {c_v}{RESET}")

    if subtitulo:
        sub_trunc = subtitulo[:espaco_util]
        margem_sub_esq = (espaco_util - len(sub_trunc)) // 2
        margem_sub_dir = espaco_util - len(sub_trunc) - margem_sub_esq
        seguro_print(f"{FG_CYAN}{c_v} {DIM}{' ' * margem_sub_esq}{sub_trunc}{' ' * margem_sub_dir}{RESET}{FG_CYAN} {c_v}{RESET}")

    seguro_print(borda_inferior)
    seguro_print()


def exibir_sucesso(texto: str):
    """Exibe mensagem de sucesso com badge verde."""
    seguro_print(f"  {FG_GREEN}{BOLD}[OK]{RESET} {texto}")


def exibir_aviso(texto: str):
    """Exibe mensagem de aviso com badge amarelo."""
    seguro_print(f"  {FG_YELLOW}{BOLD}[AVISO]{RESET} {texto}")


def exibir_erro(texto: str, detalhe: str = None, dica: str = None):
    """Exibe mensagem de erro com badge vermelho e sugestão opcional."""
    seguro_print(f"  {FG_RED}{BOLD}[ERRO]{RESET} {texto}")
    if detalhe:
        seguro_print(f"         {FG_RED}{detalhe}{RESET}")
    if dica:
        seguro_print(f"         {FG_CYAN}Dica: {dica}{RESET}")


def exibir_info(texto: str):
    """Exibe mensagem informativa com badge ciano."""
    seguro_print(f"  {FG_CYAN}[INFO]{RESET} {texto}")


def exibir_passo(numero: int, total: int, descricao: str):
    """Exibe marcador de etapa do processo."""
    seguro_print(f"\n{FG_CYAN}{BOLD}[{numero}/{total}] {descricao}{RESET}")


def formatar_opcao(tecla: str, descricao: str) -> str:
    """Retorna uma linha de opção formatada com badge para menus."""
    return f"  {FG_CYAN}{BOLD}[{tecla}]{RESET}  {descricao}"


def pausar_e_continuar(mensagem: str = "Pressione [ENTER] para continuar..."):
    """Exibe uma solicitação amigável de continuação."""
    seguro_print(f"\n{DIM}{'-' * 66}{RESET}")
    try:
        input(f"  {FG_YELLOW}{BOLD}▶ {mensagem}{RESET} ")
    except (KeyboardInterrupt, EOFError):
        seguro_print()


def pausar_e_sair(codigo: int = 0, mensagem: str = "Pressione [ENTER] para fechar esta janela..."):
    """Exibe uma solicitação amigável antes de encerrar o programa."""
    seguro_print(f"\n{DIM}{'=' * 66}{RESET}")
    try:
        input(f"  {FG_YELLOW}{BOLD}▶ {mensagem}{RESET} ")
    except (KeyboardInterrupt, EOFError):
        seguro_print()
    sys.exit(codigo)


def confirmar_acao(pergunta: str, padrao: bool = True) -> bool:
    """Solicita confirmação [S/N] amigável e tolerante."""
    sugestao = "[S/n]" if padrao else "[s/N]"
    try:
        resposta = input(f"  {FG_YELLOW}{pergunta} {sugestao}:{RESET} ").strip().lower()
    except (KeyboardInterrupt, EOFError):
        return False

    if not resposta:
        return padrao
    return resposta in ("s", "sim", "y", "yes")


# Inicializa o terminal imediatamente quando o módulo é carregado
inicializar_terminal()
