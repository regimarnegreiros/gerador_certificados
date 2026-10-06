"""
Módulo para leitura, parsing e validação do arquivo variaveis.txt.
"""

from pathlib import Path
from datetime import datetime
import re
from typing import Dict, Any, List, Tuple
from core.caminhos import abrir_arquivo, existe_com_fallback


class ConfigValidationError(Exception):
    """Exceção levantada quando há erro de validação nas variáveis do curso."""
    def __init__(self, errors: List[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


def parse_date(date_str: str) -> datetime:
    """Valida e converte string de data DD/MM/AAAA para objeto datetime."""
    clean = date_str.strip()
    return datetime.strptime(clean, "%d/%m/%Y")


def parse_percentage(val_str: str) -> float:
    """Converte valores como '0.75', '75', '75%' para float entre 0 e 1."""
    clean = str(val_str).strip().replace("%", "").replace(",", ".")
    val = float(clean)
    if val > 1.0:
        val = val / 100.0
    if not (0.0 <= val <= 1.0):
        raise ValueError(f"Percentual deve estar entre 0 e 1 (ou 0% e 100%), recebido: {val_str}")
    return val


def carregar_variaveis(caminho_arquivo: Path) -> Dict[str, Any]:
    """
    Lê e extrai os campos do arquivo variaveis.txt de forma robusta e tolerante.
    """
    if not existe_com_fallback(caminho_arquivo):
        raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

    with abrir_arquivo(caminho_arquivo, "r", encoding="utf-8-sig") as f:
        linhas = f.readlines()

    config_raw: Dict[str, Any] = {}
    corpo_email_linhas: List[str] = []
    lendo_corpo_email = False
    ignorar_emails_linhas: List[str] = []
    lendo_ignorar_emails = False

    tags_inicio_ignorar = {
        "[INICIO_IGNORAR_EMAILS]",
        "[IGNORAR_EMAILS]",
        "[INICIO_EMAILS_IGNORADOS]",
        "[EMAILS_IGNORADOS]",
        "[IGNORAR EMAILS]",
        "[EMAILS A IGNORAR]",
        "[INICIO_IGNORAR_EMAIL]",
        "[IGNORAR_EMAIL]",
    }
    tags_fim_ignorar = {
        "[FIM_IGNORAR_EMAILS]",
        "[FIM_EMAILS_IGNORADOS]",
        "[FIM_IGNORAR]",
        "[FIM_EMAILS]",
        "[FIM_IGNORAR_EMAIL]",
    }

    for num_linha, linha_original in enumerate(linhas, start=1):
        linha = linha_original.strip()

        # Comentários fora de blocos multilinhas são ignorados (evita que citações em comentários ativem leitura)
        if not lendo_corpo_email and not lendo_ignorar_emails and linha.startswith("#"):
            continue

        # Início / Fim do corpo do e-mail (tags exatas, sem considerar comentários)
        if not lendo_ignorar_emails:
            if linha.upper() in ("[INICIO_CORPO_EMAIL]", "[CORPO_EMAIL]") or (
                linha.upper().startswith("[INICIO_CORPO_EMAIL]") and not linha.startswith("#")
            ):
                lendo_corpo_email = True
                continue
            if linha.upper() in ("[FIM_CORPO_EMAIL]",) or (
                linha.upper().startswith("[FIM_CORPO_EMAIL]") and not linha.startswith("#")
            ):
                lendo_corpo_email = False
                continue

        # Início / Fim da sessão de e-mails ignorados
        if not lendo_corpo_email:
            if any(
                linha.upper() == tag or (linha.upper().startswith(tag) and not linha.startswith("#"))
                for tag in tags_inicio_ignorar
            ):
                lendo_ignorar_emails = True
                continue
            if any(
                linha.upper() == tag or (linha.upper().startswith(tag) and not linha.startswith("#"))
                for tag in tags_fim_ignorar
            ):
                lendo_ignorar_emails = False
                continue

        if lendo_corpo_email:
            corpo_email_linhas.append(linha_original)
            continue

        if lendo_ignorar_emails:
            ignorar_emails_linhas.append(linha_original)
            continue

        # Ignora linhas em branco e comentários
        if not linha or linha.startswith("#"):
            continue

        # Linhas no formato CHAVE = VALOR
        if "=" in linha:
            partes = linha.split("=", 1)
            chave = partes[0].strip().upper()
            valor = partes[1].strip()

            # Remove aspas extras ao redor do valor, caso existam
            if (valor.startswith('"') and valor.endswith('"')) or (valor.startswith("'") and valor.endswith("'")):
                valor = valor[1:-1].strip()

            config_raw[chave] = valor

    # Junta o corpo do e-mail
    if corpo_email_linhas:
        # Remove eventuais linhas iniciais de comentário separador acidental (ex: # --- ou # ===)
        while corpo_email_linhas and corpo_email_linhas[0].strip().startswith("#"):
            corpo_email_linhas.pop(0)
        corpo = "".join(corpo_email_linhas).strip()
    else:
        # Tenta pegar se veio como variável comum CORPO_EMAIL
        corpo = config_raw.get("CORPO_EMAIL", "")

    config_raw["CORPO_EMAIL"] = corpo

    # Extrai e-mails da sessão de ignorar
    emails_ignorados: List[str] = []
    for l in ignorar_emails_linhas:
        l_limpa = l.strip()
        if not l_limpa or l_limpa.startswith("#"):
            continue
        partes = [p.strip() for p in re.split(r"[,;]", l_limpa) if p.strip()]
        for p in partes:
            m = re.search(r"<([^>]+)>", p)
            em = m.group(1).strip() if m else p
            if em and em not in emails_ignorados:
                emails_ignorados.append(em)

    # Se também veio como variável avulsa IGNORAR_EMAILS = a@b.com, c@d.com
    raw_var_emails = config_raw.get("IGNORAR_EMAILS") or config_raw.get("EMAILS_IGNORADOS") or ""
    if raw_var_emails:
        partes = [p.strip() for p in re.split(r"[,;\n]", str(raw_var_emails)) if p.strip()]
        for p in partes:
            m = re.search(r"<([^>]+)>", p)
            em = m.group(1).strip() if m else p
            if em and em not in emails_ignorados:
                emails_ignorados.append(em)

    config_raw["IGNORAR_EMAILS"] = emails_ignorados
    config_raw["EMAILS_IGNORADOS"] = emails_ignorados
    return config_raw


def validar_variaveis(config_raw: Dict[str, str]) -> Dict[str, Any]:
    """
    Valida as regras de negócio e tipos das variáveis carregadas.
    Retorna dicionário tipado e higienizado ou levanta ConfigValidationError.
    """
    erros: List[str] = []
    config: Dict[str, Any] = {}

    # 1. NOME_DO_CURSO
    nome_curso = config_raw.get("NOME_DO_CURSO", "").strip()
    if not nome_curso:
        erros.append("[ERRO] O campo 'NOME_DO_CURSO' é obrigatório e não foi informado.")
    config["NOME_DO_CURSO"] = nome_curso

    # 2. CARGA_HORARIA
    carga_str = config_raw.get("CARGA_HORARIA", "").strip()
    if not carga_str:
        erros.append("[ERRO] O campo 'CARGA_HORARIA' é obrigatório.")
    else:
        try:
            carga = float(carga_str.replace(",", "."))
            if carga <= 0:
                erros.append(f"[ERRO] 'CARGA_HORARIA' deve ser maior que zero (recebido: {carga_str}).")
            config["CARGA_HORARIA"] = int(carga) if carga.is_integer() else carga
        except ValueError:
            erros.append(f"[ERRO] 'CARGA_HORARIA' deve ser um número válido (ex: 16 ou 20), recebido: '{carga_str}'.")

    # 2.1. CUSTO_TOTAL_EM_REAIS
    custo_str = (
        config_raw.get("CUSTO_TOTAL_EM_REAIS")
        or config_raw.get("CUSTO_TOTAL_REAIS")
        or config_raw.get("CUSTO_TOTAL")
        or ""
    ).strip()
    if not custo_str:
        config["CUSTO_TOTAL_EM_REAIS"] = 0.0
    else:
        limpo = custo_str.replace("R$", "").replace("r$", "").strip()
        if "," in limpo and "." in limpo:
            limpo = limpo.replace(".", "").replace(",", ".")
        elif "," in limpo:
            limpo = limpo.replace(",", ".")
        try:
            val_custo = float(limpo)
            if val_custo < 0:
                erros.append("[ERRO] 'CUSTO_TOTAL_EM_REAIS' não pode ser negativo.")
            config["CUSTO_TOTAL_EM_REAIS"] = val_custo
        except ValueError:
            erros.append(f"[ERRO] 'CUSTO_TOTAL_EM_REAIS' deve ser um número válido (ex: 5000.00 ou 5000,00), recebido: '{custo_str}'.")

    # 3. DATA_INICIO
    data_inicio_str = config_raw.get("DATA_INICIO", "").strip()
    dt_inicio = None
    if not data_inicio_str:
        erros.append("[ERRO] O campo 'DATA_INICIO' é obrigatório no formato DD/MM/AAAA.")
    else:
        try:
            dt_inicio = parse_date(data_inicio_str)
            config["DATA_INICIO"] = data_inicio_str
            config["DT_INICIO"] = dt_inicio
        except Exception:
            erros.append(f"[ERRO] 'DATA_INICIO' inválida: '{data_inicio_str}'. Use o formato DD/MM/AAAA (ex: 17/08/2026).")

    # 4. DATA_FIM (opcional se curso de 1 dia)
    data_fim_str = config_raw.get("DATA_FIM", "").strip()
    dt_fim = None
    if data_fim_str:
        try:
            dt_fim = parse_date(data_fim_str)
            config["DATA_FIM"] = data_fim_str
            config["DT_FIM"] = dt_fim

            if dt_inicio and dt_fim < dt_inicio:
                erros.append(f"[ERRO] 'DATA_FIM' ({data_fim_str}) não pode ser anterior a 'DATA_INICIO' ({data_inicio_str}).")
        except Exception:
            erros.append(f"[ERRO] 'DATA_FIM' inválida: '{data_fim_str}'. Use o formato DD/MM/AAAA ou deixe em branco.")
    else:
        config["DATA_FIM"] = ""
        config["DT_FIM"] = None

    # 5. LOCAL_EMISSAO
    local_emissao = config_raw.get("LOCAL_EMISSAO", "").strip()
    if not local_emissao:
        local_emissao = "Brasília — DF"
    config["LOCAL_EMISSAO"] = local_emissao

    # 6. DATA_EMISSAO
    data_emissao_str = config_raw.get("DATA_EMISSAO", "").strip()
    if not data_emissao_str:
        erros.append("[ERRO] O campo 'DATA_EMISSAO' é obrigatório no formato DD/MM/AAAA.")
    else:
        try:
            config["DATA_EMISSAO"] = data_emissao_str
            config["DT_EMISSAO"] = parse_date(data_emissao_str)
        except Exception:
            erros.append(f"[ERRO] 'DATA_EMISSAO' inválida: '{data_emissao_str}'. Use o formato DD/MM/AAAA.")

    # 7. PERCENTUAL_MINIMO_CARGA
    pct_carga_str = config_raw.get("PERCENTUAL_MINIMO_CARGA", "0.75").strip()
    try:
        config["PERCENTUAL_MINIMO_CARGA"] = parse_percentage(pct_carga_str)
    except Exception as e:
        erros.append(f"[ERRO] 'PERCENTUAL_MINIMO_CARGA': {e}")

    # 8. PERCENTUAL_MINIMO_ENCONTROS / PERCENTUAL_MINIMO_DIAS
    pct_encontros_str = config_raw.get("PERCENTUAL_MINIMO_ENCONTROS")
    if pct_encontros_str is None or not pct_encontros_str.strip():
        pct_encontros_str = config_raw.get("PERCENTUAL_MINIMO_DIAS", "0.75")
    pct_encontros_str = pct_encontros_str.strip()
    try:
        val_pct = parse_percentage(pct_encontros_str)
        config["PERCENTUAL_MINIMO_ENCONTROS"] = val_pct
        config["PERCENTUAL_MINIMO_DIAS"] = val_pct
    except Exception as e:
        erros.append(f"[ERRO] 'PERCENTUAL_MINIMO_ENCONTROS': {e}")

    # 9. TEMPO_MINIMO_PRESENCA_MINUTOS
    tempo_min_str = config_raw.get("TEMPO_MINIMO_PRESENCA_MINUTOS", "30").strip()
    try:
        tempo_min = int(tempo_min_str)
        if tempo_min < 0:
            erros.append("[ERRO] 'TEMPO_MINIMO_PRESENCA_MINUTOS' não pode ser negativo.")
        config["TEMPO_MINIMO_PRESENCA_MINUTOS"] = tempo_min
    except ValueError:
        erros.append(f"[ERRO] 'TEMPO_MINIMO_PRESENCA_MINUTOS' deve ser um número inteiro (ex: 30), recebido: '{tempo_min_str}'.")

    # 10. ASSUNTO_EMAIL
    assunto = config_raw.get("ASSUNTO_EMAIL", f"Certificado - Curso {nome_curso}").strip()
    config["ASSUNTO_EMAIL"] = assunto

    # 11. CORPO_EMAIL
    corpo = config_raw.get("CORPO_EMAIL", "").strip()
    if not corpo:
        erros.append("[ERRO] O corpo do e-mail está vazio no arquivo 'variaveis.txt'.")
    config["CORPO_EMAIL"] = corpo

    # 12. IGNORAR_EMAILS / EMAILS_IGNORADOS
    ignorar_raw = config_raw.get("IGNORAR_EMAILS", config_raw.get("EMAILS_IGNORADOS", []))
    if isinstance(ignorar_raw, str):
        emails_ign = [p.strip() for p in re.split(r"[,;\n]", ignorar_raw) if p.strip()]
    elif isinstance(ignorar_raw, (list, tuple, set)):
        emails_ign = [str(p).strip() for p in ignorar_raw if str(p).strip()]
    else:
        emails_ign = []

    emails_finais = []
    for em in emails_ign:
        m = re.search(r"<([^>]+)>", em)
        em_limpo = m.group(1).strip() if m else em
        if em_limpo and em_limpo not in emails_finais:
            emails_finais.append(em_limpo)

    config["IGNORAR_EMAILS"] = emails_finais
    config["EMAILS_IGNORADOS"] = emails_finais

    if erros:
        raise ConfigValidationError(erros)

    return config


def obter_configuracao(diretorio_curso: Path) -> Dict[str, Any]:
    """Carrega e valida as configurações de um curso a partir do diretório."""
    caminho_variaveis = diretorio_curso / "variaveis.txt"
    raw = carregar_variaveis(caminho_variaveis)
    return validar_variaveis(raw)
