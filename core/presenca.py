"""
Módulo para processamento de listas de presença do Microsoft Teams,
cruzamento com dados de servidores ANM e geração do relatório Excel formatado.
Extraído e adaptado diretamente de presensa_teams.ipynb.
"""

from pathlib import Path
import unicodedata
import os
import re
import math
import csv
import io
from datetime import datetime, timedelta, date, time
import pandas as pd
from openpyxl.styles import (
    PatternFill,
    Font,
    Alignment,
    Border,
    Side
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule
import zipfile
import xml.etree.ElementTree as ET
from typing import Tuple, List, Dict, Any, Optional, Set, Collection, Union, Callable
from core.config import format_time_str, parse_time
from core.eventos import emitir_log as _emitir_log, emitir_progresso as _emitir_progresso
from core.caminhos import (
    caminho_estendido,
    caminho_estendido_path,
    garantir_diretorio,
    abrir_arquivo,
    copiar_arquivo_com_fallback,
    existe_com_fallback,
    listar_diretorio_com_fallback,
    ler_excel_com_fallback,
    sanitizar_nome_arquivo,
    encurtar_caminho_se_necessario,
)


def tempo_em_segundos(tempo: Any) -> int:
    """Converte strings de tempo (HH:MM:SS ou 1h 30m 20s) em segundos."""
    if pd.isna(tempo):
        return 0

    tempo = str(tempo).strip()

    # Formato HH:MM:SS ou MM:SS
    if ":" in tempo:
        partes = tempo.split(":")
        if len(partes) == 3:
            horas, minutos, segundos = map(int, partes)
            return horas * 3600 + minutos * 60 + segundos
        elif len(partes) == 2:
            minutos, segundos = map(int, partes)
            return minutos * 60 + segundos

    # Formato 1h 30m 20s
    valores = {
        unidade: int(valor)
        for valor, unidade in re.findall(r"(\d+)([hms])", tempo)
    }

    return (
        valores.get("h", 0) * 3600 +
        valores.get("m", 0) * 60 +
        valores.get("s", 0)
    )


def segundos_em_tempo(segundos: int) -> str:
    """Formata segundos em 'Xh Ym Zs'."""
    horas = segundos // 3600
    segundos %= 3600
    minutos = segundos // 60
    segundos %= 60

    partes = []
    if horas:
        partes.append(f"{horas}h")
    if minutos:
        partes.append(f"{minutos}m")
    if segundos:
        partes.append(f"{segundos}s")

    return " ".join(partes) if partes else "0s"


def segundos_em_formato_digital(segundos: int) -> str:
    """Formata segundos em 'H:MM:SS'."""
    horas = segundos // 3600
    minutos = (segundos % 3600) // 60
    segundos = segundos % 60
    return f"{horas}:{minutos:02d}:{segundos:02d}"


def segundos_em_formato_hh_mm_ss(segundos: int) -> str:
    """Formata segundos em 'HH:MM:SS' (ex: 3665 -> '01:01:05', 0 -> '00:00:00')."""
    if segundos is None or segundos < 0:
        return "00:00:00"
    horas = segundos // 3600
    minutos = (segundos % 3600) // 60
    segs = segundos % 60
    return f"{horas:02d}:{minutos:02d}:{segs:02d}"


def normalizar_texto(texto: Any) -> str:
    """Normaliza texto removendo acentos, apóstrofos e espaços duplicados em maiúsculas."""
    if not isinstance(texto, str):
        return ""
    texto = texto.strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.replace("'", " ")
    texto = " ".join(texto.split())
    return texto


def normalizar_email(email: Any) -> str:
    """Normaliza e-mail removendo tags <...>, espaços e convertendo para minúsculas."""
    if email is None or pd.isna(email):
        return ""
    txt = str(email).strip().lower()
    m = re.search(r"<([^>]+)>", txt)
    if m:
        return m.group(1).strip().lower()
    return txt


def filtrar_emails_ignorados(df: pd.DataFrame, emails_ignorados: Optional[Collection[str]] = None) -> pd.DataFrame:
    """
    Filtra e remove registros de participantes cujos e-mails constam na lista de ignorados.
    Verifica tanto colunas de e-mail quanto a coluna Nome (caso contenha Nome <email>).
    """
    if not emails_ignorados or df is None or df.empty:
        return df

    emails_set = {normalizar_email(e) for e in emails_ignorados if normalizar_email(e)}
    if not emails_set:
        return df

    col_email = None
    for c in df.columns:
        c_upper = str(c).strip().upper()
        if c_upper in ["EMAIL", "E-MAIL", "CORREIO", "E-MAIL ADDRESS", "EMAIL ADDRESS"]:
            col_email = c
            break
    if col_email is None:
        for c in df.columns:
            c_upper = str(c).strip().upper()
            if "EMAIL" in c_upper or "CORREIO" in c_upper:
                col_email = c
                break

    mascara = pd.Series(True, index=df.index)
    if col_email is not None:
        serie_emails = df[col_email].apply(normalizar_email)
        mascara = mascara & ~serie_emails.isin(emails_set)

    if "Nome" in df.columns:
        serie_nome_email = df["Nome"].apply(normalizar_email)
        mascara = mascara & ~serie_nome_email.isin(emails_set)

    return df[mascara].copy().reset_index(drop=True)


def converter_para_inteiro(val: Any) -> int:
    """Converte valor para inteiro de forma segura, tratando NaN, nulos e strings vazias."""
    if pd.isna(val) or val is None:
        return 0
    if isinstance(val, (int, float)):
        return int(val) if not math.isnan(val) else 0
    val_str = str(val).strip()
    if not val_str:
        return 0
    try:
        return int(float(val_str.replace(",", ".")))
    except (ValueError, TypeError):
        return 0


def extrair_datetime_resumo(texto: Any) -> Optional[datetime]:
    """
    Extrai objeto datetime a partir de strings como '9/14/26,8:01:54 AM', '14/09/2026' ou '2026-09-14'.
    """
    if pd.isna(texto) or not texto:
        return None
    texto_str = str(texto).strip().strip('"\'')
    if not texto_str:
        return None

    texto_str = re.sub(r"\ba\s*\.\s*m\s*\.?", "AM", texto_str, flags=re.IGNORECASE)
    texto_str = re.sub(r"\bp\s*\.\s*m\s*\.?", "PM", texto_str, flags=re.IGNORECASE)

    formatos = [
        "%m/%d/%y,%I:%M:%S %p",
        "%m/%d/%y, %I:%M:%S %p",
        "%m/%d/%Y,%I:%M:%S %p",
        "%m/%d/%Y, %I:%M:%S %p",
        "%d/%m/%Y,%H:%M:%S",
        "%d/%m/%Y, %H:%M:%S",
        "%d/%m/%y,%H:%M:%S",
        "%d/%m/%y, %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%m/%d/%y",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y-%m-%d",
    ]
    for fmt in formatos:
        try:
            return datetime.strptime(texto_str, fmt)
        except ValueError:
            pass

    m = re.match(
        r"^(\d{1,2})/(\d{1,2})/(\d{2,4})(?:[,\s]+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*(AM|PM)?)?",
        texto_str,
        re.IGNORECASE,
    )
    if m:
        p1, p2, p3 = int(m.group(1)), int(m.group(2)), int(m.group(3))
        ano = p3 if p3 > 99 else 2000 + p3
        h = int(m.group(4)) if m.group(4) else 0
        mi = int(m.group(5)) if m.group(5) else 0
        s = int(m.group(6)) if m.group(6) else 0
        ampm = m.group(7)
        if ampm:
            if ampm.upper() == "PM" and h < 12:
                h += 12
            elif ampm.upper() == "AM" and h == 12:
                h = 0

        if p1 > 12:
            dia, mes = p1, p2
        elif p2 > 12:
            mes, dia = p1, p2
        else:
            if ampm:
                mes, dia = p1, p2
            else:
                dia, mes = p1, p2
        try:
            return datetime(ano, mes, dia, h, mi, s)
        except ValueError:
            pass

    return None


def extrair_data_resumo(texto: Any) -> str:
    """
    Extrai e padroniza a data a partir de strings como '9/14/26,8:01:54 AM' para 'DD/MM/AAAA'.
    """
    dt = extrair_datetime_resumo(texto)
    if dt:
        return dt.strftime("%d/%m/%Y")
    if pd.isna(texto) or not texto:
        return ""
    texto_str = str(texto).strip().strip('"\'')
    return texto_str.split(",")[0].strip() if texto_str else ""


def _normalizar_horario(valor: Any) -> Optional[time]:
    """Converte datetime, time ou string ('HH:MM', 'HH:MM:SS', 'AM/PM') em datetime.time."""
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.time()
    if isinstance(valor, time):
        return valor
    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return None
    return parse_time(texto)


def parse_datetime_atividade(texto: Any) -> Optional[datetime]:
    """
    Converte data/hora da seção '3. Atividades em Reunião' no formato
    MM/DD/AA, HH:MM:SS AM/PM (ex.: '8/02/26, 8:01:06 AM' -> 02/08/2026 08:01:06),
    tratando mês/dia/hora sem zero à esquerda e 12h AM/PM (12:xx AM = 00:xx, 12:xx PM = 12:xx).
    """
    if texto is None or pd.isna(texto):
        return None
    if isinstance(texto, datetime):
        return texto
    if isinstance(texto, pd.Timestamp):
        return texto.to_pydatetime()

    texto_str = str(texto).strip().strip('"\'')
    if not texto_str or texto_str.lower() == "nan":
        return None

    texto_norm = re.sub(r"\ba\s*\.\s*m\s*\.?", "AM", texto_str, flags=re.IGNORECASE)
    texto_norm = re.sub(r"\bp\s*\.\s*m\s*\.?", "PM", texto_norm, flags=re.IGNORECASE)

    m = re.match(
        r"^(\d{1,2})/(\d{1,2})/(\d{2,4})[,\s]+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?(?:\s*(AM|PM))?$",
        texto_norm,
        re.IGNORECASE,
    )
    if m:
        p1, p2, p3 = int(m.group(1)), int(m.group(2)), int(m.group(3))
        ano = p3 if p3 > 99 else 2000 + p3
        h = int(m.group(4))
        mi = int(m.group(5))
        s = int(m.group(6)) if m.group(6) is not None else 0
        ampm = m.group(7)

        if ampm:
            if not (1 <= h <= 12):
                return None
            ampm_u = ampm.upper()
            if ampm_u == "AM" and h == 12:
                h = 0
            elif ampm_u == "PM" and h < 12:
                h += 12
        elif not (0 <= h <= 23):
            return None

        if not (0 <= mi <= 59) or not (0 <= s <= 59):
            return None

        # Prioriza MM/DD/AA (fallback DD/MM/AA somente se p1 > 12)
        if 1 <= p1 <= 12:
            mes, dia = p1, p2
        elif p1 > 12 and 1 <= p2 <= 12:
            dia, mes = p1, p2
        else:
            return None

        try:
            return datetime(ano, mes, dia, h, mi, s)
        except ValueError:
            return None

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.strptime(texto_norm, fmt)
        except ValueError:
            pass

    return None


def _eh_marcador_secao3(linha: str) -> bool:
    """Verifica se a linha representa o cabeçalho da seção '3. Atividades em Reunião'."""
    limpa = linha.strip().strip('"\'')
    if not limpa:
        return False
    primeira_celula = re.split(r"[\t,;]", limpa, maxsplit=1)[0].strip().strip('"\'')
    norm = normalizar_texto(primeira_celula)
    return (
        norm.startswith("3. ATIVIDADES")
        or norm.startswith("3 ATIVIDADES")
        or "ATIVIDADES EM REUNI" in norm
        or norm.startswith("3. IN-MEETING")
        or "IN-MEETING ACTIVITIES" in norm
    )


def _identificar_colunas_atividades(colunas: Any) -> Dict[str, Any]:
    """Identifica as colunas da seção '3. Atividades em Reunião' tolerando variações de acento/caixa."""
    mapa: Dict[str, Any] = {
        "nome": None,
        "entrada": None,
        "saida": None,
        "duracao": None,
        "email": None,
        "funcao": None,
    }
    norm_list = [(col, normalizar_texto(str(col))) for col in colunas]

    # 1. Prioriza correspondências exatas
    for col, c_norm in norm_list:
        if mapa["nome"] is None and c_norm in ("NOME", "NAME"):
            mapa["nome"] = col
        elif mapa["entrada"] is None and c_norm in ("HORARIO DE ENTRADA", "HORA DE ENTRADA", "JOIN TIME"):
            mapa["entrada"] = col
        elif mapa["saida"] is None and c_norm in ("HORARIO DE SAIDA", "HORA DE SAIDA", "LEAVE TIME"):
            mapa["saida"] = col
        elif mapa["duracao"] is None and c_norm in ("DURACAO", "DURACAO NA REUNIAO", "DURACAO DA REUNIAO", "DURATION"):
            mapa["duracao"] = col
        elif mapa["email"] is None and c_norm in ("EMAIL", "E-MAIL", "CORREIO", "EMAIL ADDRESS", "E-MAIL ADDRESS"):
            mapa["email"] = col
        elif mapa["funcao"] is None and c_norm in ("FUNCAO", "PAPEL", "ROLE"):
            mapa["funcao"] = col

    # 2. Fallbacks parciais
    for col, c_norm in norm_list:
        if mapa["nome"] is None and (c_norm.startswith("NOME ") or "NOME" in c_norm):
            mapa["nome"] = col
        elif mapa["entrada"] is None and ("ENTRADA" in c_norm or "JOIN" in c_norm):
            mapa["entrada"] = col
        elif mapa["saida"] is None and ("SAIDA" in c_norm or "LEAVE" in c_norm):
            mapa["saida"] = col
        elif mapa["duracao"] is None and ("DURA" in c_norm or "DURATION" in c_norm):
            mapa["duracao"] = col
        elif mapa["email"] is None and (("EMAIL" in c_norm and "ID" not in c_norm) or "CORREIO" in c_norm):
            mapa["email"] = col
        elif mapa["funcao"] is None and ("FUNCAO" in c_norm or "ROLE" in c_norm):
            mapa["funcao"] = col

    return mapa


def ler_atividades_reuniao(caminho: Path) -> Optional[pd.DataFrame]:
    """
    Lê a seção '3. Atividades em Reunião' da lista de presença do Teams.
    Retorna DataFrame com colunas padronizadas ['Nome', 'Horário de entrada',
    'Horário de Saída', 'Duração', 'Email', 'Função'] ou None se a seção/tabela
    não existir no arquivo.
    """
    encodings = ["utf-16", "utf-8-sig", "utf-8", "latin1"]

    for enc in encodings:
        try:
            err_mode = "replace" if enc == "latin1" else "strict"
            with abrir_arquivo(caminho, mode="r", encoding=enc, errors=err_mode) as f:
                linhas_arquivo = f.readlines()
        except Exception:
            continue

        idx_secao3 = None
        for idx, linha in enumerate(linhas_arquivo):
            if _eh_marcador_secao3(linha):
                idx_secao3 = idx
                break

        if idx_secao3 is None:
            continue

        proximo = idx_secao3 + 1
        while proximo < len(linhas_arquivo) and not linhas_arquivo[proximo].strip():
            proximo += 1

        if proximo >= len(linhas_arquivo):
            return None

        linhas_secao: List[str] = []
        for linha in linhas_arquivo[proximo:]:
            limpa = linha.strip().strip('"\'')
            primeira_cel = re.split(r"[\t,;]", limpa, maxsplit=1)[0].strip().strip('"\'') if limpa else ""
            if re.match(r"^[4-9]\.\s+\S", primeira_cel):
                break
            linhas_secao.append(linha)

        conteudo_secao = "".join(linhas_secao)
        if not conteudo_secao.strip():
            return None

        for sep in ["\t", ",", ";"]:
            try:
                candidato = pd.read_csv(io.StringIO(conteudo_secao), sep=sep, dtype=str)
            except Exception:
                continue

            mapa = _identificar_colunas_atividades(candidato.columns)
            if mapa["nome"] is not None and mapa["entrada"] is not None and mapa["saida"] is not None:
                df_padrao = pd.DataFrame({
                    "Nome": candidato[mapa["nome"]],
                    "Horário de entrada": candidato[mapa["entrada"]],
                    "Horário de Saída": candidato[mapa["saida"]],
                    "Duração": candidato[mapa["duracao"]] if mapa["duracao"] is not None else "",
                    "Email": candidato[mapa["email"]] if mapa["email"] is not None else "",
                    "Função": candidato[mapa["funcao"]] if mapa["funcao"] is not None else "",
                })
                mascara_valida = (
                    df_padrao["Nome"].fillna("").astype(str).str.strip().ne("")
                    | df_padrao["Email"].fillna("").astype(str).str.strip().ne("")
                )
                return df_padrao[mascara_valida].copy().reset_index(drop=True)

        # Encontrou o marcador da Seção 3, mas não há tabela de atividades válida
        return None

    return None


def resolver_data_encontro_consistente(
    resumo: Dict[str, Any],
    df_atividades: Optional[pd.DataFrame],
) -> Optional[date]:
    """
    Retorna a data da Seção 1 (Resumo) quando ela coincide com alguma data presente
    na Seção 3 (Atividades em Reunião). Caso contrário, retorna None para usar a data
    de cada linha da Seção 3.
    """
    dt_res = resumo.get("datetime") if resumo else None
    if df_atividades is None or df_atividades.empty or "Horário de entrada" not in df_atividades.columns:
        return dt_res.date() if isinstance(dt_res, datetime) else None

    datas_s3: Set[date] = set()
    for val in df_atividades["Horário de entrada"]:
        dt = parse_datetime_atividade(val)
        if dt is not None:
            datas_s3.add(dt.date())

    if isinstance(dt_res, datetime) and datas_s3 and dt_res.date() in datas_s3:
        return dt_res.date()

    return None


def calcular_duracao_recortada_segundos(
    dt_entrada: Optional[datetime],
    dt_saida: Optional[datetime],
    horario_inicio: Optional[Any] = None,
    horario_fim: Optional[Any] = None,
    data_referencia: Optional[date] = None,
) -> int:
    """
    Calcula a duração válida (em segundos) da interseção entre [dt_entrada, dt_saida]
    e a janela [horario_inicio, horario_fim].
    Se nenhum horário limite estiver configurado, retorna a duração total da linha.
    """
    if dt_entrada is None or dt_saida is None or dt_saida <= dt_entrada:
        return 0

    tm_inicio = _normalizar_horario(horario_inicio)
    tm_fim = _normalizar_horario(horario_fim)

    if tm_inicio is None and tm_fim is None:
        return int((dt_saida - dt_entrada).total_seconds())

    def _intersecao_para_data(ref_date: date) -> int:
        janela_inicio = datetime.combine(ref_date, tm_inicio) if tm_inicio is not None else None
        janela_fim = datetime.combine(ref_date, tm_fim) if tm_fim is not None else None
        if janela_inicio is not None and janela_fim is not None and janela_fim <= janela_inicio:
            return 0
        inicio_efetivo = max(dt_entrada, janela_inicio) if janela_inicio is not None else dt_entrada
        fim_efetivo = min(dt_saida, janela_fim) if janela_fim is not None else dt_saida
        if fim_efetivo <= inicio_efetivo:
            return 0
        return int((fim_efetivo - inicio_efetivo).total_seconds())

    if data_referencia is not None:
        return _intersecao_para_data(data_referencia)

    dias_span = (dt_saida.date() - dt_entrada.date()).days
    if dias_span == 0 or dias_span > 31:
        return _intersecao_para_data(dt_entrada.date())

    total = 0
    for desloc in range(dias_span + 1):
        total += _intersecao_para_data(dt_entrada.date() + timedelta(days=desloc))
    return total


def processar_atividades_reuniao(
    df_atividades: pd.DataFrame,
    horario_inicio: Optional[Any] = None,
    horario_fim: Optional[Any] = None,
    data_referencia: Optional[date] = None,
) -> pd.DataFrame:
    """
    Processa a tabela da seção '3. Atividades em Reunião', aplicando o recorte pela janela
    [horario_inicio, horario_fim] e somando o tempo válido de linhas repetidas da mesma
    pessoa (agrupando por Email normalizado; se vazio, por Nome normalizado).
    Retorna DataFrame no padrão ['Nome', 'Email', 'Duração da Reunião'].
    """
    colunas_saida = ["Nome", "Email", "Duração da Reunião"]
    if df_atividades is None or df_atividades.empty:
        return pd.DataFrame(columns=colunas_saida)

    mapa = _identificar_colunas_atividades(df_atividades.columns)
    col_nome = mapa["nome"]
    col_entrada = mapa["entrada"]
    col_saida = mapa["saida"]
    col_duracao = mapa["duracao"]
    col_email = mapa["email"]

    tm_inicio = _normalizar_horario(horario_inicio)
    tm_fim = _normalizar_horario(horario_fim)

    agrupados: Dict[Tuple[str, str], Dict[str, Any]] = {}

    for _, row in df_atividades.iterrows():
        nome_norm = normalizar_texto(row.get(col_nome, "")) if col_nome is not None else ""
        email_norm = normalizar_email(row.get(col_email, "")) if col_email is not None else ""

        if not nome_norm and not email_norm:
            continue

        dt_entrada = parse_datetime_atividade(row.get(col_entrada)) if col_entrada is not None else None
        dt_saida = parse_datetime_atividade(row.get(col_saida)) if col_saida is not None else None

        if dt_entrada is not None and dt_saida is not None:
            segundos = calcular_duracao_recortada_segundos(
                dt_entrada=dt_entrada,
                dt_saida=dt_saida,
                horario_inicio=tm_inicio,
                horario_fim=tm_fim,
                data_referencia=data_referencia,
            )
        elif tm_inicio is None and tm_fim is None and col_duracao is not None:
            segundos = tempo_em_segundos(row.get(col_duracao))
        else:
            segundos = 0

        chave = ("EMAIL", email_norm) if email_norm else ("NOME", nome_norm)
        if chave not in agrupados:
            agrupados[chave] = {
                "Nome": nome_norm or email_norm.upper(),
                "Email": email_norm,
                "segundos": 0,
            }
        elif not agrupados[chave]["Nome"] and nome_norm:
            agrupados[chave]["Nome"] = nome_norm

        agrupados[chave]["segundos"] += segundos

    if not agrupados:
        return pd.DataFrame(columns=colunas_saida)

    registros = [
        {
            "Nome": dados["Nome"],
            "Email": dados["Email"],
            "Duração da Reunião": segundos_em_formato_hh_mm_ss(dados["segundos"]),
        }
        for dados in agrupados.values()
    ]
    return pd.DataFrame(registros, columns=colunas_saida)


def ler_resumo_reuniao(caminho: Path) -> Dict[str, Any]:
    """
    Lê a seção '1. Resumo' da lista de presença do Microsoft Teams.
    Extrai:
    - 'data': data formatada como DD/MM/AAAA
    - 'datetime': objeto datetime de início da reunião (se disponível)
    - 'datetime_fim': objeto datetime de término da reunião (se disponível)
    - 'hora_inicio_raw': texto bruto do campo 'Hora de início'
    - 'hora_termino_raw': texto bruto do campo 'Hora de término'
    - 'horario_inicio': horário formatado HH:MM (24h) de início
    - 'horario_fim': horário formatado HH:MM (24h) de término
    - 'duracao_str': duração da reunião (ex: '4h 2m 32s')
    - 'duracao_segundos': duração da reunião em segundos
    - 'tempo_medio_str': tempo médio de participação (ex: '3h 15m 10s')
    - 'tempo_medio_segundos': tempo médio de participação em segundos
    """
    resumo = {
        "data": "",
        "datetime": None,
        "datetime_fim": None,
        "hora_inicio_raw": "",
        "hora_termino_raw": "",
        "horario_inicio": "",
        "horario_fim": "",
        "duracao_str": "",
        "duracao_segundos": 0,
        "tempo_medio_str": "",
        "tempo_medio_segundos": 0,
        "tempo_medio_participacao_str": "",
        "tempo_medio_participacao_segundos": 0,
    }
    encodings = ["utf-16", "utf-8-sig", "utf-8", "latin1"]
    linhas_resumo = []

    for enc in encodings:
        try:
            with abrir_arquivo(caminho, mode="r", encoding=enc, errors="replace") as f:
                linhas = f.readlines()
                for linha in linhas:
                    limpa = linha.strip().strip('"\'').lower()
                    if limpa.startswith("2. participantes") or "2. participantes" in limpa:
                        break
                    linhas_resumo.append(linha)
                if linhas_resumo:
                    break
        except Exception:
            continue

    for linha in linhas_resumo:
        linha_str = linha.strip()
        if not linha_str:
            continue

        partes = None
        for sep in ["\t", ";", ","]:
            try:
                reader = csv.reader(io.StringIO(linha_str), delimiter=sep)
                row = next(reader)
                if len(row) >= 2 and row[0].strip():
                    partes = [row[0].strip(), row[1].strip()]
                    break
            except Exception:
                continue

        if not partes:
            for sep in ["\t", ";"]:
                if sep in linha_str:
                    bits = linha_str.split(sep, 1)
                    partes = [bits[0].strip().strip('"\''), bits[1].strip().strip('"\'')]
                    break

        if not partes or len(partes) < 2:
            continue

        chave_norm = normalizar_texto(partes[0])
        valor = partes[1].strip().strip('"\'')

        if "HORA DE INICIO" in chave_norm or "START TIME" in chave_norm:
            if not resumo["hora_inicio_raw"]:
                resumo["hora_inicio_raw"] = valor
            dt = extrair_datetime_resumo(valor)
            if dt:
                if resumo["datetime"] is None:
                    resumo["datetime"] = dt
                if not resumo["data"]:
                    resumo["data"] = dt.strftime("%d/%m/%Y")
                if not resumo["horario_inicio"]:
                    resumo["horario_inicio"] = format_time_str(dt.time())
            elif not resumo["data"]:
                resumo["data"] = extrair_data_resumo(valor)
        elif "HORA DE TERMINO" in chave_norm or "HORA DO TERMINO" in chave_norm or "HORA DE FIM" in chave_norm or "HORA DO FIM" in chave_norm:
            if not resumo["hora_termino_raw"]:
                resumo["hora_termino_raw"] = valor
            dt_fim = extrair_datetime_resumo(valor)
            if dt_fim:
                if resumo["datetime_fim"] is None:
                    resumo["datetime_fim"] = dt_fim
                if not resumo["horario_fim"]:
                    resumo["horario_fim"] = format_time_str(dt_fim.time())
        elif "TEMPO MEDIO DE PARTICIPACAO" in chave_norm or "TEMPO MEDIO" in chave_norm or "AVERAGE ATTENDANCE TIME" in chave_norm or "AVERAGE PARTICIPATION TIME" in chave_norm:
            if not resumo["tempo_medio_str"]:
                segs = tempo_em_segundos(valor)
                resumo["tempo_medio_segundos"] = segs
                val_str = valor if valor else (segundos_em_tempo(segs) if segs > 0 else "")
                resumo["tempo_medio_str"] = val_str
                resumo["tempo_medio_participacao_segundos"] = segs
                resumo["tempo_medio_participacao_str"] = val_str
        elif "DURACAO DA REUNIAO" in chave_norm or "DURACAO" in chave_norm or "MEETING DURATION" in chave_norm:
            if not resumo["duracao_str"]:
                segs = tempo_em_segundos(valor)
                resumo["duracao_segundos"] = segs
                resumo["duracao_str"] = valor if valor else (segundos_em_tempo(segs) if segs > 0 else "")

    return resumo


def _chave_natural(texto: str) -> List[Any]:
    """Chave para ordenação natural alfanumérica (ex: 'dia_2' antes de 'dia_10')."""
    return [int(c) if c.isdigit() else c.lower() for c in re.split(r"(\d+)", str(texto))]


def obter_chave_ordenacao_arquivo(caminho: Path) -> Tuple[datetime, List[Any]]:
    """
    Retorna uma tupla (datetime, natural_sort_key) para ordenação cronológica precisa dos encontros.
    Prioridades:
    1. Seção 1 (Resumo da Reunião): 'Hora de início'
    2. Data encontrada no nome do arquivo (YYYY-MM-DD, DD-MM-YYYY, etc.)
    3. Número do dia/encontro no nome do arquivo (ex: dia_1, encontro_2)
    4. Data encontrada dentro das primeiras linhas do arquivo
    5. Fallback com desempate por ordenação natural do nome.
    """
    caminho = Path(caminho)

    # 1. Tenta ler resumo da reunião
    try:
        resumo = ler_resumo_reuniao(caminho)
        if resumo.get("datetime"):
            return (resumo["datetime"], _chave_natural(caminho.name))
        if resumo.get("data"):
            try:
                dt = datetime.strptime(resumo["data"], "%d/%m/%Y")
                return (dt, _chave_natural(caminho.name))
            except Exception:
                pass
    except Exception:
        pass

    nome = caminho.stem

    # 2. Data no nome do arquivo: YYYY-MM-DD ou YYYYMMDD
    m_ano = re.search(r"(\d{4})[-_]?(\d{2})[-_]?(\d{2})", nome)
    if m_ano:
        try:
            ano, mes, dia = int(m_ano.group(1)), int(m_ano.group(2)), int(m_ano.group(3))
            return (datetime(ano, mes, dia), _chave_natural(caminho.name))
        except ValueError:
            pass

    # Data no nome do arquivo: DD-MM-YYYY ou DD_MM_YYYY ou DD.MM.YYYY
    m_dia = re.search(r"(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})", nome)
    if m_dia:
        try:
            dia, mes, ano = int(m_dia.group(1)), int(m_dia.group(2)), int(m_dia.group(3))
            return (datetime(ano, mes, dia), _chave_natural(caminho.name))
        except ValueError:
            pass

    # 3. Número no nome do arquivo (ex: dia_1, encontro_2, aula 3)
    m_num = re.search(r"(?:dia|encontro|aula|reuni[aã]o|modulo|m[oó]dulo)[_\s-]*(\d+)", nome, re.IGNORECASE)
    if m_num:
        num = int(m_num.group(1))
        return (datetime(1970, 1, 1) + timedelta(days=num), _chave_natural(caminho.name))

    # Número isolado no nome (ex: "1.csv", "02 - Presença.csv")
    m_num_iso = re.search(r"^\s*(\d+)", nome)
    if m_num_iso:
        num = int(m_num_iso.group(1))
        return (datetime(1970, 1, 1) + timedelta(days=num), _chave_natural(caminho.name))

    # 4. Tenta encontrar qualquer data nas primeiras 30 linhas do arquivo
    encodings = ["utf-16", "utf-8-sig", "utf-8", "latin1"]
    for enc in encodings:
        try:
            with abrir_arquivo(caminho, mode="r", encoding=enc, errors="replace") as f:
                for _ in range(30):
                    linha = f.readline()
                    if not linha:
                        break
                    dt = extrair_datetime_resumo(linha)
                    if dt:
                        return (dt, _chave_natural(caminho.name))
            break
        except Exception:
            continue

    # 5. Fallback final: tenta ler data na Seção 3 antes de retornar datetime máximo
    try:
        df_s3 = ler_atividades_reuniao(caminho)
        if df_s3 is not None and not df_s3.empty and "Horário de entrada" in df_s3.columns:
            for val in df_s3["Horário de entrada"]:
                dt = parse_datetime_atividade(val)
                if dt is not None:
                    return (dt, _chave_natural(caminho.name))
    except Exception:
        pass

    return (datetime(9999, 12, 31), _chave_natural(caminho.name))


def ordenar_arquivos_presenca(arquivos: List[Path]) -> List[Path]:
    """Ordena uma lista de arquivos de presença estritamente por data do encontro."""
    if not arquivos:
        return []
    return sorted(arquivos, key=obter_chave_ordenacao_arquivo)



def identificar_colunas_metricas(colunas: List[str]) -> Dict[str, List[str]]:
    """
    Identifica colunas de engajamento e compromisso no Teams:
    - Câmera Ligada
    - Desativar Mudo (Microfone)
    - Levantar as Mãos
    - Reações (Aplauso, Gargalhada, Curtir, Amor, Surpreso, etc.)
    """
    cols_camera = []
    cols_microfone = []
    cols_maos = []
    cols_reacoes = []

    for col in colunas:
        c_norm = normalizar_texto(col)

        # Câmera Ligada
        if "CAMERA LIGADA" in c_norm or ("CAMERA" in c_norm and "COMPROMISSO" in c_norm):
            cols_camera.append(col)
            continue

        # Desativar Mudo (Microfone)
        if "DESATIVAR MUDO" in c_norm or ("MUDO" in c_norm and "COMPROMISSO" in c_norm):
            cols_microfone.append(col)
            continue

        # Levantar as Mãos
        if ("LEVANTAR" in c_norm and "MAO" in c_norm) or ("MAOS" in c_norm and "COMPROMISSO" in c_norm):
            cols_maos.append(col)
            continue

        # Reações
        if (
            "ENVOLVIMENTO: REACAO" in c_norm
            or "REACAO -" in c_norm
            or any(r in c_norm for r in ["APLAUSO", "GARGALHADA", "CURTIR", "AMOR", "SURPRESO"])
            and ("ENVOLVIMENTO" in c_norm or "REACAO" in c_norm)
        ):
            cols_reacoes.append(col)
            continue

    return {
        "camera": cols_camera,
        "microfone": cols_microfone,
        "maos": cols_maos,
        "reacoes": cols_reacoes,
    }


def processar_interacoes_participantes(df_presenca: pd.DataFrame) -> pd.DataFrame:
    """
    Extrai as métricas de interação de cada participante em um único encontro.
    Calcula:
    - camera_qtd: quantidade total de vezes que ligou a câmera
    - camera_ligou: 1 se ligou pelo menos uma vez, 0 caso contrário
    - microfone_qtd: quantidade total de vezes que desativou o mudo
    - microfone_ligou: 1 se desativou mudo pelo menos uma vez, 0 caso contrário
    - maos_qtd: total de vezes que levantou a mão
    - reacoes_qtd: total de reações enviadas
    """
    cols_map = identificar_colunas_metricas(list(df_presenca.columns))

    col_nome = "Nome"
    col_email = "Email" if "Email" in df_presenca.columns else None
    if col_email is None:
        for c in df_presenca.columns:
            if "EMAIL" in str(c).upper() or "CORREIO" in str(c).upper():
                col_email = c
                break

    registros = []
    for _, row in df_presenca.iterrows():
        nome = normalizar_texto(row.get(col_nome, ""))
        if not nome:
            continue
        email = str(row.get(col_email, "")).strip() if col_email else ""
        if email.lower() == "nan":
            email = ""

        cam_val = sum(converter_para_inteiro(row.get(c)) for c in cols_map["camera"])
        mic_val = sum(converter_para_inteiro(row.get(c)) for c in cols_map["microfone"])
        mao_val = sum(converter_para_inteiro(row.get(c)) for c in cols_map["maos"])
        reac_val = sum(converter_para_inteiro(row.get(c)) for c in cols_map["reacoes"])

        registros.append({
            "Nome": nome,
            "Email": email,
            "camera_qtd": cam_val,
            "camera_ligou": 1 if cam_val > 0 else 0,
            "microfone_qtd": mic_val,
            "microfone_ligou": 1 if mic_val > 0 else 0,
            "maos_qtd": mao_val,
            "reacoes_qtd": reac_val,
        })

    df_res = pd.DataFrame(registros)
    if df_res.empty:
        return pd.DataFrame(columns=[
            "Nome", "Email", "camera_qtd", "camera_ligou",
            "microfone_qtd", "microfone_ligou", "maos_qtd", "reacoes_qtd"
        ])

    df_agrupado = df_res.groupby(["Nome", "Email"], as_index=False).agg({
        "camera_qtd": "sum",
        "camera_ligou": lambda s: 1 if (s > 0).any() else 0,
        "microfone_qtd": "sum",
        "microfone_ligou": lambda s: 1 if (s > 0).any() else 0,
        "maos_qtd": "sum",
        "reacoes_qtd": "sum",
    })
    return df_agrupado


def ler_presenca(caminho: Path) -> pd.DataFrame:
    """
    Lê a lista de presença do Teams pulando os cabeçalhos de resumo
    e recortando na seção '3. Atividades em Reunião'.
    Tolerante a diferentes codificações (utf-16, utf-8, latin1) e caminhos longos no Windows.
    """
    encodings = ["utf-16", "utf-8-sig", "utf-8", "latin1"]
    df = None

    for enc in encodings:
        try:
            # Localiza a linha com '2. Participantes' para definir o início dos dados
            skiprows = None
            err_mode = "replace" if enc == "latin1" else "strict"
            with abrir_arquivo(caminho, mode="r", encoding=enc, errors=err_mode) as f:
                linhas_arquivo = f.readlines()
                for idx, linha in enumerate(linhas_arquivo):
                    linha_limpa = linha.strip().strip('"\'').lower()
                    if linha_limpa.startswith("2. participantes") or "2. participantes" in linha_limpa:
                        # Pula até a primeira linha com conteúdo logo abaixo de '2. Participantes'
                        proximo = idx + 1
                        while proximo < len(linhas_arquivo) and not linhas_arquivo[proximo].strip():
                            proximo += 1
                        skiprows = proximo
                        break

            if skiprows is not None:
                linhas_secao2 = []
                for l in linhas_arquivo[skiprows:]:
                    if _eh_marcador_secao3(l):
                        break
                    linhas_secao2.append(l)
                conteudo_secao2 = "".join(linhas_secao2)
                if conteudo_secao2.strip():
                    for sep in ["\t", ",", ";"]:
                        try:
                            candidato = pd.read_csv(io.StringIO(conteudo_secao2), sep=sep)
                            if "Nome" in candidato.columns:
                                df = candidato
                                break
                        except Exception:
                            continue
                if df is not None and "Nome" in df.columns:
                    break
            else:
                # Fallback: tenta com o padrão fixo de 9 linhas caso não haja o marcador '2. Participantes'
                for sep in ["\t", ",", ";"]:
                    try:
                        candidato = pd.read_csv(caminho, skiprows=9, encoding=enc, sep=sep)
                        if "Nome" in candidato.columns:
                            df = candidato
                            break
                    except (FileNotFoundError, OSError):
                        try:
                            candidato = pd.read_csv(caminho_estendido(caminho), skiprows=9, encoding=enc, sep=sep)
                            if "Nome" in candidato.columns:
                                df = candidato
                                break
                        except Exception:
                            continue
                    except Exception:
                        continue
                if df is not None and "Nome" in df.columns:
                    break
        except Exception:
            continue

    if df is None or "Nome" not in df.columns:
        # Tenta sem skiprows
        for enc in encodings:
            try:
                candidato = pd.read_csv(caminho, encoding=enc, sep=None, engine="python")
                if "Nome" in candidato.columns:
                    df = candidato
                    break
            except (FileNotFoundError, OSError):
                try:
                    candidato = pd.read_csv(caminho_estendido(caminho), encoding=enc, sep=None, engine="python")
                    if "Nome" in candidato.columns:
                        df = candidato
                        break
                except Exception:
                    continue
            except Exception:
                continue

    if df is None or "Nome" not in df.columns:
        raise ValueError(f"Não foi possível identificar a coluna 'Nome' no arquivo de presença: {Path(caminho).name}")

    inicio_segunda_tabela = df[
        df["Nome"].astype(str).str.strip().str.startswith("3. Atividades em Reuni")
    ].index

    if len(inicio_segunda_tabela) > 0:
        df = df.iloc[:inicio_segunda_tabela[0]]

    # Limpa linhas vazias
    df = df.dropna(subset=["Nome"])
    return df


def localizar_coluna_servidor(colunas: Any) -> Any:
    """Localiza a coluna do nome do servidor, priorizando correspondência exata."""
    for col in colunas:
        nome = str(col).strip().upper()
        if nome in ["SERVIDOR", "NOME", "NOME SERVIDOR", "NOME DO SERVIDOR"]:
            return col

    for col in colunas:
        nome = str(col).strip().upper()
        if "SERVIDOR" in nome or "NOME" in nome:
            return col

    return None


def localizar_coluna_uorg(colunas: Any) -> Any:
    """
    Localiza a coluna correspondente a UORG na base de servidores.
    Prioriza correspondência exata para a coluna chamada exatamente 'UORG'
    (ignorando maiúsculas/minúsculas e espaços em branco nas extremidades),
    evitando selecionar colunas auxiliares de código (ex: 'COD UORG' ou 'CÓD UORG').
    """
    # 1. Prioridade absoluta: nome exatamente 'UORG'
    for col in colunas:
        if str(col).strip().upper() == "UORG":
            return col

    # 2. Busca exata por termos equivalentes conhecidos
    for col in colunas:
        nome = str(col).strip().upper()
        if nome in ["LOTACAO", "LOTAÇÃO", "UNIDADE"]:
            return col

    # 3. Busca parcial, descartando explicitamente códigos/IDs (ex: COD UORG, CÓDIGO)
    for col in colunas:
        nome = str(col).strip().upper()
        if any(cod in nome for cod in ["COD", "CÓD", "ID"]):
            continue
        if "UORG" in nome or "LOTACAO" in nome or "LOTAÇÃO" in nome or "UNIDADE" in nome:
            return col

    return None


def juntar_servidor_presenca(servidores: pd.DataFrame, presenca: pd.DataFrame) -> pd.DataFrame:
    """Cruza a presença com os servidores da ANM para obter UORG."""
    servidores = servidores.copy()
    presenca = presenca.copy()

    # Localiza coluna de servidor
    col_servidor = localizar_coluna_servidor(servidores.columns)
    if col_servidor is None:
        col_servidor = servidores.columns[0]

    # Localiza coluna UORG garantindo prioridade para o nome exato "UORG"
    col_uorg = localizar_coluna_uorg(servidores.columns)
    if col_uorg is None:
        col_uorg = servidores.columns[1] if len(servidores.columns) > 1 else col_servidor

    servidores["NOME_NORM"] = servidores[col_servidor].apply(normalizar_texto)
    presenca["Nome"] = presenca["Nome"].apply(normalizar_texto)
    presenca["NOME_NORM"] = presenca["Nome"]

    uorg_map = servidores.drop_duplicates(subset=["NOME_NORM"])[["NOME_NORM", col_uorg]].rename(columns={col_uorg: "UORG"})

    presentes = presenca.merge(
        uorg_map,
        on="NOME_NORM",
        how="left"
    )

    col_duracao = "Duração da Reunião" if "Duração da Reunião" in presentes.columns else None
    if col_duracao is None:
        for c in presentes.columns:
            if "DURA" in str(c).upper() or "PERMANENCIA" in str(c).upper() or "TEMPO" in str(c).upper():
                col_duracao = c
                break
    if col_duracao is None:
        col_duracao = "Duração da Reunião"
        presentes[col_duracao] = "00:00:00"

    col_email = "Email" if "Email" in presentes.columns else None
    if col_email is None:
        for c in presentes.columns:
            if "EMAIL" in str(c).upper() or "CORREIO" in str(c).upper():
                col_email = c
                break
    if col_email is None:
        col_email = "Email"
        presentes[col_email] = ""

    presentes = presentes.rename(columns={col_duracao: "Duração da Reunião", col_email: "Email"})
    return presentes[["Nome", "UORG", "Email", "Duração da Reunião"]]


def ler_convocados(caminho: Path) -> pd.DataFrame:
    """Lê a lista de convocados exportada do Teams (formato 'Nome <email>')."""
    encodings = ["utf-8-sig", "utf-8", "latin1", "utf-16"]
    df = None

    for enc in encodings:
        try:
            with abrir_arquivo(caminho, mode="r", encoding=enc, errors="replace") as f:
                df = pd.read_csv(f, skiprows=2)
            if len(df.columns) > 0:
                break
        except Exception:
            continue

    if df is None:
        for enc in encodings:
            try:
                with abrir_arquivo(caminho, mode="r", encoding=enc, errors="replace") as f:
                    df = pd.read_csv(f)
                if len(df.columns) > 0:
                    break
            except Exception:
                continue

    col_nome = df.columns[0]
    convocados = pd.DataFrame()

    # Extrai Nome e Email
    extraidos = df[col_nome].astype(str).str.extract(r"^(.*?)\s*<([^>]+)>$")
    if extraidos[0].notna().any():
        convocados["Nome"] = extraidos[0].str.strip().apply(normalizar_texto)
        convocados["Email"] = extraidos[1].str.strip()
    else:
        # Se já estiver em colunas separadas
        convocados["Nome"] = df[col_nome].astype(str).str.strip().apply(normalizar_texto)
        if len(df.columns) > 1:
            convocados["Email"] = df[df.columns[1]].astype(str).str.strip()
        else:
            convocados["Email"] = ""

    convocados = convocados.dropna(subset=["Nome"])
    convocados = convocados[convocados["Nome"] != ""]
    return convocados


def marcar_comparecimento(resultado: pd.DataFrame, tempo_comparecimento_minutos: int) -> pd.DataFrame:
    """Marca comparecimento como Presente ou Ausente com base no tempo mínimo."""
    tempo_minimo_segundos = tempo_comparecimento_minutos * 60
    resultado = resultado.copy()

    resultado["Compareceu"] = resultado["Duração da Reunião"].apply(
        lambda tempo: (
            "Presente"
            if (pd.notna(tempo) and tempo_em_segundos(tempo) >= tempo_minimo_segundos)
            else "Ausente"
        )
    )
    return resultado


def criar_dados_do_encontro(resultado: pd.DataFrame, numero_encontro: int) -> pd.DataFrame:
    """Formata colunas do encontro para consolidação."""
    encontro = resultado[["Nome", "Email", "Duração da Reunião", "Compareceu"]].copy()

    encontro["Duração da Reunião"] = encontro["Duração da Reunião"].apply(
        lambda tempo: (
            segundos_em_formato_digital(tempo_em_segundos(tempo))
            if pd.notna(tempo)
            else "00:00:00"
        )
    )

    encontro = encontro.rename(
        columns={
            "Duração da Reunião": f"Encontro {numero_encontro}",
            "Compareceu": f"Comparecimento {numero_encontro}",
        }
    )
    return encontro


def calcular_presenca_por_uorg(resultado: pd.DataFrame) -> pd.DataFrame:
    """Calcula quantidade de servidores presentes por UORG no encontro."""
    presentes = resultado[resultado["Compareceu"] == "Presente"]
    if "UORG" in presentes.columns:
        contagem = presentes.groupby("UORG").size().reset_index(name="PRESENTES")
    else:
        contagem = pd.DataFrame(columns=["UORG", "PRESENTES"])
    return contagem


def processar_encontro(
    arquivo: Path,
    servidores: pd.DataFrame,
    numero_encontro: int,
    tempo_comparecimento_minutos: int,
    emails_ignorados: Optional[Collection[str]] = None,
    horario_inicio: Optional[Any] = None,
    horario_fim: Optional[Any] = None,
):
    """Processa um arquivo individual de presença de um encontro."""
    dados_atividades = ler_atividades_reuniao(arquivo)
    if dados_atividades is not None:
        if emails_ignorados:
            dados_atividades = filtrar_emails_ignorados(dados_atividades, emails_ignorados)
        resumo = ler_resumo_reuniao(arquivo)
        data_ref = resolver_data_encontro_consistente(resumo, dados_atividades)
        dados = processar_atividades_reuniao(
            dados_atividades,
            horario_inicio=horario_inicio,
            horario_fim=horario_fim,
            data_referencia=data_ref,
        )
    else:
        dados = ler_presenca(arquivo)
        if emails_ignorados:
            dados = filtrar_emails_ignorados(dados, emails_ignorados)

    resultado = juntar_servidor_presenca(servidores, dados)
    resultado = marcar_comparecimento(resultado, tempo_comparecimento_minutos)
    encontro = criar_dados_do_encontro(resultado, numero_encontro)
    por_uorg = calcular_presenca_por_uorg(resultado)
    return resultado, encontro, por_uorg


def preparar_convocados(
    caminho_convocados: Path,
    servidores: pd.DataFrame,
    emails_ignorados: Optional[Collection[str]] = None,
) -> pd.DataFrame:
    """Carrega convocados e anexa o UORG a partir da base de servidores."""
    convocados = ler_convocados(caminho_convocados)
    if emails_ignorados:
        convocados = filtrar_emails_ignorados(convocados, emails_ignorados)
    servidores = servidores.copy()

    col_servidor = localizar_coluna_servidor(servidores.columns)
    if col_servidor is None:
        col_servidor = servidores.columns[0]

    col_uorg = localizar_coluna_uorg(servidores.columns)
    if col_uorg is None:
        col_uorg = servidores.columns[1] if len(servidores.columns) > 1 else col_servidor

    servidores["NOME_NORM"] = servidores[col_servidor].apply(normalizar_texto)
    convocados["NOME_NORM"] = convocados["Nome"].apply(normalizar_texto)

    uorg_servidores = servidores[["NOME_NORM", col_uorg]].drop_duplicates(subset=["NOME_NORM"]).rename(columns={col_uorg: "UORG"})
    convocados = convocados.merge(uorg_servidores, on="NOME_NORM", how="left").drop(columns=["NOME_NORM"])
    return convocados


def _indexar_participantes_por_email_e_nome(df: pd.DataFrame) -> Tuple[Dict[str, pd.Series], Dict[str, pd.Series], Dict[str, pd.Series]]:
    """Indexa registros de participantes por e-mail normalizado e nome normalizado."""
    por_email: Dict[str, pd.Series] = {}
    por_nome_geral: Dict[str, pd.Series] = {}
    por_nome_sem_email: Dict[str, pd.Series] = {}

    if df is None or df.empty:
        return por_email, por_nome_geral, por_nome_sem_email

    for _, row in df.iterrows():
        email_k = normalizar_email(row.get("Email", "")) if "Email" in df.columns else ""
        nome_k = normalizar_texto(row.get("Nome", "")) if "Nome" in df.columns else ""
        if email_k and email_k not in por_email:
            por_email[email_k] = row
        if nome_k and nome_k not in por_nome_geral:
            por_nome_geral[nome_k] = row
        if not email_k and nome_k and nome_k not in por_nome_sem_email:
            por_nome_sem_email[nome_k] = row

    return por_email, por_nome_geral, por_nome_sem_email


def _buscar_registro_participante(
    nome: Any,
    email: Any,
    por_email: Dict[str, pd.Series],
    por_nome_geral: Dict[str, pd.Series],
    por_nome_sem_email: Dict[str, pd.Series],
) -> Optional[pd.Series]:
    """Localiza participante priorizando Email normalizado e usando Nome como fallback quando Email estiver vazio."""
    email_k = normalizar_email(email)
    nome_k = normalizar_texto(nome)

    if email_k and email_k in por_email:
        return por_email[email_k]
    if not email_k and nome_k and nome_k in por_nome_geral:
        return por_nome_geral[nome_k]
    if email_k and nome_k and nome_k in por_nome_sem_email:
        return por_nome_sem_email[nome_k]
    return None


def calcular_total_presencas(resultados: pd.DataFrame) -> pd.DataFrame:
    """Calcula quantidade total de encontros em que o participante esteve 'Presente'."""
    if resultados is None or resultados.empty:
        return pd.DataFrame(columns=["Nome", "Email", "Presenças"])

    presentes = resultados[resultados["Compareceu"] == "Presente"]
    if presentes.empty:
        return pd.DataFrame(columns=["Nome", "Email", "Presenças"])

    contagem: Dict[Tuple[str, str], Dict[str, Any]] = {}
    for _, row in presentes.iterrows():
        nome_raw = row.get("Nome", "")
        email_raw = row.get("Email", "") if "Email" in presentes.columns else ""
        email_k = normalizar_email(email_raw)
        nome_k = normalizar_texto(nome_raw)
        chave = ("EMAIL", email_k) if email_k else ("NOME", nome_k)
        if chave not in contagem:
            contagem[chave] = {
                "Nome": nome_raw,
                "Email": email_k if email_k else "",
                "Presenças": 0,
            }
        contagem[chave]["Presenças"] += 1

    return pd.DataFrame(list(contagem.values()), columns=["Nome", "Email", "Presenças"])


def adicionar_total_presencas(convocados: pd.DataFrame, resultados: pd.DataFrame) -> pd.DataFrame:
    """Adiciona a contagem de presenças aos convocados (cruzando por Email e fallback por Nome)."""
    relatorio = convocados.copy()
    totais = calcular_total_presencas(resultados)
    por_email, por_nome_geral, por_nome_sem_email = _indexar_participantes_por_email_e_nome(totais)

    presencas_list: List[int] = []
    for _, row in relatorio.iterrows():
        email_k = normalizar_email(row.get("Email", "")) if "Email" in relatorio.columns else ""
        nome_k = normalizar_texto(row.get("Nome", ""))
        if email_k:
            total_p = 0
            if email_k in por_email:
                total_p += int(por_email[email_k]["Presenças"])
            if nome_k and nome_k in por_nome_sem_email:
                total_p += int(por_nome_sem_email[nome_k]["Presenças"])
            presencas_list.append(total_p)
        else:
            encontrado = por_nome_geral.get(nome_k) if nome_k else None
            presencas_list.append(int(encontrado["Presenças"]) if encontrado is not None else 0)

    relatorio["Presenças"] = presencas_list
    return relatorio


def adicionar_encontros_presenca(relatorio: pd.DataFrame, encontros_presenca: List[pd.DataFrame]) -> pd.DataFrame:
    """Adiciona colunas dos encontros (Encontro X e Comparecimento X) ao relatório."""
    relatorio = relatorio.copy()
    for encontro in encontros_presenca:
        if encontro is None or encontro.empty:
            for col in encontro.columns if encontro is not None else []:
                if col not in ("Nome", "Email") and col not in relatorio.columns:
                    relatorio[col] = None
            continue

        colunas_dados = [c for c in encontro.columns if c not in ("Nome", "Email")]
        por_email, por_nome_geral, por_nome_sem_email = _indexar_participantes_por_email_e_nome(encontro)

        valores_por_col: Dict[str, List[Any]] = {c: [] for c in colunas_dados}
        for _, row in relatorio.iterrows():
            encontrado = _buscar_registro_participante(
                row.get("Nome", ""),
                row.get("Email", "") if "Email" in relatorio.columns else "",
                por_email,
                por_nome_geral,
                por_nome_sem_email,
            )
            for c in colunas_dados:
                valores_por_col[c].append(encontrado[c] if encontrado is not None else None)

        for c in colunas_dados:
            relatorio[c] = valores_por_col[c]

    qtd_encontros = len(encontros_presenca)
    for num in range(1, qtd_encontros + 1):
        col_encontro = f"Encontro {num}"
        if col_encontro in relatorio.columns:
            relatorio[col_encontro] = relatorio[col_encontro].fillna("00:00:00")

        col_comp = f"Comparecimento {num}"
        if col_comp in relatorio.columns:
            relatorio[col_comp] = relatorio[col_comp].fillna("Ausente")

    return relatorio


def aplicar_justificativas(
    relatorio: pd.DataFrame,
    arquivos_presenca: List[Path],
    justificativas: Optional[List[Dict[str, Any]]] = None,
) -> pd.DataFrame:
    """
    Aplica justificativas de falta registradas no curso:
    para cada justificativa cujo participante esteja 'Ausente' no encontro correspondente,
    define Comparecimento N = 'Justificado', abona o tempo em Encontro N e recalcula 'Presenças'.
    """
    if relatorio is None or relatorio.empty or not justificativas:
        return relatorio

    relatorio = relatorio.copy()
    mapa_arquivo_para_num = {Path(arq).name: idx for idx, arq in enumerate(arquivos_presenca, start=1)}

    por_email: Dict[str, int] = {}
    por_nome: Dict[str, int] = {}
    for idx_row, row in relatorio.iterrows():
        email_k = normalizar_email(row.get("Email", "")) if "Email" in relatorio.columns else ""
        nome_k = normalizar_texto(row.get("Nome", ""))
        if email_k and email_k not in por_email:
            por_email[email_k] = idx_row
        if nome_k and nome_k not in por_nome:
            por_nome[nome_k] = idx_row

    for just in justificativas:
        arq_nome = str(just.get("arquivo_encontro", "")).strip()
        num_encontro = mapa_arquivo_para_num.get(arq_nome)
        if num_encontro is None:
            continue

        col_enc = f"Encontro {num_encontro}"
        col_comp = f"Comparecimento {num_encontro}"
        if col_enc not in relatorio.columns or col_comp not in relatorio.columns:
            continue

        email_j = normalizar_email(just.get("email", ""))
        nome_j = normalizar_texto(just.get("nome", ""))
        idx_alvo = por_email.get(email_j) if email_j else None
        if idx_alvo is None and nome_j:
            idx_alvo = por_nome.get(nome_j)
        if idx_alvo is None:
            continue

        status_atual = str(relatorio.at[idx_alvo, col_comp]).strip().upper()
        if status_atual == "PRESENTE":
            continue

        tempo_abonado = max(0, int(just.get("tempo_abonado_segundos", 0) or 0))
        tempo_atual_seg = tempo_em_segundos(relatorio.at[idx_alvo, col_enc])
        tempo_final_seg = max(tempo_atual_seg, tempo_abonado)

        relatorio.at[idx_alvo, col_comp] = "Justificado"
        relatorio.at[idx_alvo, col_enc] = segundos_em_formato_digital(tempo_final_seg)

    # Recalcula 'Presenças' de forma idêntica à fórmula Excel (=COUNTIFS(..., "<>Ausente"))
    qtd_encontros = len(arquivos_presenca)
    novas_presencas: List[int] = []
    for _, row in relatorio.iterrows():
        total_p = 0
        for num in range(1, qtd_encontros + 1):
            col_comp = f"Comparecimento {num}"
            if col_comp in relatorio.columns:
                val = str(row.get(col_comp, "")).strip()
                if val and val.upper() != "AUSENTE":
                    total_p += 1
        novas_presencas.append(total_p)
    relatorio["Presenças"] = novas_presencas

    return relatorio


def calcular_carga_participada(linha: pd.Series, quantidade_encontros: int) -> int:
    """Soma tempo total participado em todos os encontros em segundos."""
    total = 0
    for num in range(1, quantidade_encontros + 1):
        tempo = linha.get(f"Encontro {num}")
        if pd.notna(tempo):
            total += tempo_em_segundos(tempo)
    return total


def adicionar_carga_horaria(relatorio: pd.DataFrame, quantidade_encontros: int, carga_horaria_horas: float) -> Tuple[pd.DataFrame, pd.Series]:
    """Calcula carga horária participada e percentual em relação à carga total."""
    relatorio = relatorio.copy()
    carga_segundos = relatorio.apply(lambda linha: calcular_carga_participada(linha, quantidade_encontros), axis=1)

    relatorio["CH Total"] = carga_segundos.apply(segundos_em_formato_digital)

    total_segundos_esperado = int(carga_horaria_horas * 3600)
    if total_segundos_esperado > 0:
        relatorio["% da CH Total"] = carga_segundos / total_segundos_esperado
    else:
        relatorio["% da CH Total"] = 0.0

    return relatorio, carga_segundos


def calcular_encontros_minimos(quantidade_encontros: int, percentual_minimo: float) -> int:
    """
    Calcula quantidade mínima de encontros de presença exigidos.
    Usa arredondamento para cima (math.ceil) para garantir que, caso haja
    exigência de presença (percentual > 0), seja necessário pelo menos 1 encontro.
    Evita que cursos de 1 encontro com 75% resultem em 0 encontros mínimos (o que aprovava todos).
    """
    if quantidade_encontros <= 0 or percentual_minimo <= 0:
        return 0
    encontros = math.ceil(round(quantidade_encontros * percentual_minimo, 6))
    return max(1, min(quantidade_encontros, encontros))


def adicionar_aptidao(
    relatorio: pd.DataFrame,
    carga_segundos: pd.Series,
    quantidade_encontros: int = 0,
    carga_horaria_horas: float = 0.0,
    percentual_minimo_carga: float = 0.75,
    percentual_minimo_encontros: float = 0.75,
) -> pd.DataFrame:
    """Calcula critérios de aprovação / aptidão."""
    relatorio = relatorio.copy()

    carga_total_seg = int(carga_horaria_horas * 3600)
    carga_minima_seg = int(carga_total_seg * percentual_minimo_carga)
    encontros_minimos = calcular_encontros_minimos(quantidade_encontros, percentual_minimo_encontros)

    status_presenca = (relatorio["Presenças"] >= encontros_minimos).map({True: "APTO", False: "NÃO APTO"})
    relatorio["Apto por CH"] = (carga_segundos >= carga_minima_seg).map({True: "APTO", False: "NÃO APTO"})
    relatorio["Apto por Presença"] = status_presenca
    relatorio["Apto por Encontros"] = status_presenca
    relatorio["Apto em Ambos"] = (
        (relatorio["Apto por CH"] == "APTO") & (status_presenca == "APTO")
    ).map({True: "APTO", False: "NÃO APTO"})

    return relatorio


def criar_aba_presencas(relatorio: pd.DataFrame, quantidade_encontros: int) -> pd.DataFrame:
    """
    Cria DataFrame para a aba 'Presenças'.
    Contém apenas a coluna 'Nome' e as colunas 'Encontro n' e 'Comparecimento n'.
    """
    colunas_presencas = ["Nome"]
    for num in range(1, quantidade_encontros + 1):
        col_encontro = f"Encontro {num}"
        col_comp = f"Comparecimento {num}"
        if col_encontro in relatorio.columns:
            colunas_presencas.append(col_encontro)
        if col_comp in relatorio.columns:
            colunas_presencas.append(col_comp)

    colunas_existentes = [c for c in colunas_presencas if c in relatorio.columns]
    return relatorio[colunas_existentes].copy()


def criar_aba_interacoes_participante(
    relatorio: pd.DataFrame,
    lista_interacoes_encontros: List[pd.DataFrame],
    lista_resultados_encontros: List[pd.DataFrame],
) -> pd.DataFrame:
    """
    Cria a Aba 1: 'Interações por participante'.
    Colunas:
    - Participante
    - Qtd Presenças
    - Câmera
    - % Câmera
    - Microfone
    - % Microfone
    - Mãos levantadas
    - Total Reações
    """
    col_nome = "Nome"
    col_email = "Email" if "Email" in relatorio.columns else None

    # Base de participantes a partir do relatório consolidado
    df_base = relatorio[[col_nome]].copy()
    if col_email:
        df_base["Email"] = relatorio[col_email]

    if "Presenças" in relatorio.columns:
        df_base["Presenças"] = relatorio["Presenças"]
    else:
        df_base["Presenças"] = 0

    mapas_encontros = [
        (
            _indexar_participantes_por_email_e_nome(df_int),
            _indexar_participantes_por_email_e_nome(df_res),
        )
        for df_int, df_res in zip(lista_interacoes_encontros, lista_resultados_encontros)
    ]

    dados_participantes = []

    for _, part in df_base.iterrows():
        nome = normalizar_texto(part[col_nome])
        email = part["Email"] if col_email else ""
        qtd_presencas = int(part["Presenças"])

        presencas_com_camera = 0
        presencas_com_microfone = 0
        total_maos = 0
        total_reacoes = 0

        for mapas_int, mapas_res in mapas_encontros:
            # Verifica comparecimento no encontro
            row_res = _buscar_registro_participante(nome, email, *mapas_res)
            compareceu = bool(row_res is not None and row_res.get("Compareceu") == "Presente")

            # Localiza dados de interação do participante no encontro
            row_int = _buscar_registro_participante(nome, email, *mapas_int)
            if row_int is not None:
                cam_ligou = int(row_int.get("camera_ligou", 0))
                mic_ligou = int(row_int.get("microfone_ligou", 0))
                maos = int(row_int.get("maos_qtd", 0))
                reac = int(row_int.get("reacoes_qtd", 0))

                total_maos += maos
                total_reacoes += reac

                # Se compareceu ao encontro e ligou a câmera, contabiliza 1 presença com câmera
                if compareceu and cam_ligou >= 1:
                    presencas_com_camera += 1
                elif cam_ligou >= 1 and qtd_presencas > 0:
                    presencas_com_camera += 1

                # Se compareceu ao encontro e desativou mudo, contabiliza 1 presença com microfone
                if compareceu and mic_ligou >= 1:
                    presencas_com_microfone += 1
                elif mic_ligou >= 1 and qtd_presencas > 0:
                    presencas_com_microfone += 1

        if qtd_presencas > 0:
            presencas_com_camera = min(presencas_com_camera, qtd_presencas)
            presencas_com_microfone = min(presencas_com_microfone, qtd_presencas)
            cobertura_cam = presencas_com_camera / qtd_presencas
            cobertura_mic = presencas_com_microfone / qtd_presencas
        else:
            presencas_com_camera = 0
            presencas_com_microfone = 0
            cobertura_cam = 0.0
            cobertura_mic = 0.0

        dados_participantes.append({
            "Participante": nome,
            "Qtd Presenças": qtd_presencas,
            "Câmera": presencas_com_camera,
            "% Câmera": cobertura_cam,
            "Microfone": presencas_com_microfone,
            "% Microfone": cobertura_mic,
            "Mãos levantadas": total_maos,
            "Total Reações": total_reacoes,
        })

    df_aba1 = pd.DataFrame(dados_participantes)
    if "Participante" in df_aba1.columns:
        df_aba1 = df_aba1.sort_values(
            by="Participante",
            ascending=True,
            key=lambda col: col.astype(str).str.strip().str.lower()
        ).reset_index(drop=True)
    return df_aba1


def _calcular_tempo_medio_encontro_seg(
    resumo: Dict[str, Any],
    df_res: pd.DataFrame,
    config: Optional[Dict[str, Any]] = None,
) -> int:
    """Calcula o tempo médio de participação do encontro (priorizando tempos recortados quando aplicável)."""
    if resumo.get("tempo_medio_recortado_segundos") is not None:
        return int(resumo["tempo_medio_recortado_segundos"])
    if resumo.get("tempo_medio_recortado_calculado"):
        return int(resumo.get("tempo_medio_segundos", 0) or 0)

    recorte_ativo = bool(resumo.get("recorte_aplicado")) or bool(
        config
        and (
            config.get("TM_INICIO") is not None
            or config.get("TM_FIM") is not None
            or str(config.get("HORARIO_INICIO", "")).strip()
            or str(config.get("HORARIO_FIM", "")).strip()
        )
    )
    if recorte_ativo and df_res is not None and "Duração da Reunião" in df_res.columns:
        tempos_presentes = [
            tempo_em_segundos(t) for t in df_res["Duração da Reunião"]
            if tempo_em_segundos(t) > 0
        ]
        return int(round(sum(tempos_presentes) / len(tempos_presentes))) if tempos_presentes else 0

    seg_tm = resumo.get("tempo_medio_segundos", 0)
    if seg_tm == 0:
        seg_tm_raw = resumo.get("tempo_medio_str", "")
        if seg_tm_raw:
            seg_tm = tempo_em_segundos(seg_tm_raw)
    if seg_tm == 0 and df_res is not None and "Duração da Reunião" in df_res.columns:
        tempos = [tempo_em_segundos(t) for t in df_res["Duração da Reunião"] if tempo_em_segundos(t) > 0]
        seg_tm = int(round(sum(tempos) / len(tempos))) if tempos else 0
    return int(seg_tm)


def criar_aba_interacoes_encontro(
    total_convocados: int,
    lista_resumos: List[Dict[str, Any]],
    lista_interacoes_encontros: List[pd.DataFrame],
    lista_resultados_encontros: List[pd.DataFrame],
    config: Optional[Dict[str, Any]] = None,
) -> pd.DataFrame:
    """
    Cria a Aba 2: 'Interações por encontro'.
    Colunas:
    - ID
    - Data
    - CH Realizada
    - Tempo Médio
    - Presentes
    - Ligou a câmera
    - Ligou o microfone
    - Levantou a mão
    - Reações
    """
    linhas = []
    data_inicio_cfg = config.get("DATA_INICIO", "") if config else ""

    for idx, (resumo, df_int, df_res) in enumerate(
        zip(lista_resumos, lista_interacoes_encontros, lista_resultados_encontros),
        start=1
    ):
        data_enc = resumo.get("data") or data_inicio_cfg or ""

        # Duração da reunião em formato digital 00:00:00
        dur_seg = resumo.get("duracao_segundos", 0)
        if dur_seg == 0:
            dur_str_raw = resumo.get("duracao_str", "")
            if dur_str_raw:
                dur_seg = tempo_em_segundos(dur_str_raw)
        if dur_seg == 0 and "Duração da Reunião" in df_res.columns:
            dur_seg = max([tempo_em_segundos(t) for t in df_res["Duração da Reunião"]], default=0)

        dur_str = segundos_em_formato_hh_mm_ss(dur_seg)

        # Tempo médio de participação em formato digital 00:00:00
        tempo_medio_seg = _calcular_tempo_medio_encontro_seg(resumo, df_res, config=config)

        tempo_medio_str = segundos_em_formato_hh_mm_ss(tempo_medio_seg)

        # Presentes em formato 'X de Y' para eliminar aviso de heurística de data do Excel
        presentes_count = int((df_res["Compareceu"] == "Presente").sum()) if "Compareceu" in df_res.columns else 0
        txt_presentes = f"{presentes_count} de {total_convocados}"

        # Ligou a câmera e Ligou o microfone (usuários presentes naquele dia)
        if "Compareceu" in df_res.columns and not df_int.empty:
            df_pres_dia = df_res[df_res["Compareceu"] == "Presente"]
            emails_pres = (
                {normalizar_email(e) for e in df_pres_dia["Email"] if normalizar_email(e)}
                if "Email" in df_pres_dia.columns
                else set()
            )
            nomes_pres = {normalizar_texto(n) for n in df_pres_dia["Nome"] if normalizar_texto(n)}

            def _esta_presente(row_int):
                em = normalizar_email(row_int.get("Email", ""))
                if em and em in emails_pres:
                    return True
                nm = normalizar_texto(row_int.get("Nome", ""))
                return nm in nomes_pres

            df_int_presentes = df_int[df_int.apply(_esta_presente, axis=1)]
            camera_count = int((df_int_presentes["camera_ligou"] >= 1).sum()) if not df_int_presentes.empty else 0
            microfone_count = int((df_int_presentes["microfone_ligou"] >= 1).sum()) if not df_int_presentes.empty else 0
        else:
            camera_count = int((df_int["camera_ligou"] >= 1).sum()) if not df_int.empty else 0
            microfone_count = int((df_int["microfone_ligou"] >= 1).sum()) if not df_int.empty else 0

        txt_camera = f"{camera_count} de {presentes_count}"
        txt_microfone = f"{microfone_count} de {presentes_count}"

        # Mãos levantadas no encontro
        maos_count = int(df_int["maos_qtd"].sum()) if not df_int.empty else 0

        # Reações no encontro
        reacoes_count = int(df_int["reacoes_qtd"].sum()) if not df_int.empty else 0

        linhas.append({
            "ID": idx,
            "Data": data_enc,
            "CH Realizada": dur_str,
            "Tempo Médio": tempo_medio_str,
            "Presentes": txt_presentes,
            "Ligou a câmera": txt_camera,
            "Ligou o microfone": txt_microfone,
            "Levantou a mão": maos_count,
            "Reações": reacoes_count,
        })

    def _parse_data_para_sort(d_str: str) -> datetime:
        if not d_str:
            return datetime(9999, 12, 31)
        try:
            return datetime.strptime(d_str, "%d/%m/%Y")
        except Exception:
            return datetime(9999, 12, 31)

    linhas.sort(key=lambda r: _parse_data_para_sort(r.get("Data", "")))
    for idx_novo, r in enumerate(linhas, start=1):
        r["ID"] = idx_novo

    return pd.DataFrame(linhas)


def criar_aba_indicadores(
    total_convidados: int,
    relatorio: pd.DataFrame,
    df_interacoes_participante: pd.DataFrame,
    df_interacoes_encontro: pd.DataFrame,
    lista_resumos: List[Dict[str, Any]],
    lista_resultados_encontros: List[pd.DataFrame],
) -> pd.DataFrame:
    """
    Cria a Aba 3: 'Indicadores'.
    Colunas: 'Indicador', 'Resultado'
    Linhas:
    - Convidados a participar
    - Participantes totais
    - Participantes certificados
    - Carga horária realizada
    - Tempo médio de participação
    - Taxa de certificação
    - Participações com Câmera
    - Participações com Microfone
    - Total de reações
    - Total de mãos levantadas
    """
    # 1. Convidados a participar
    convidados_val = total_convidados

    # 2. Participantes totais (presentes em pelo menos 1 encontro)
    if "Presenças" in relatorio.columns:
        participantes_totais = int((relatorio["Presenças"] >= 1).sum())
    elif "Qtd Presenças" in df_interacoes_participante.columns:
        participantes_totais = int((df_interacoes_participante["Qtd Presenças"] >= 1).sum())
    elif "Quantidade de presenças" in df_interacoes_participante.columns:
        participantes_totais = int((df_interacoes_participante["Quantidade de presenças"] >= 1).sum())
    else:
        participantes_totais = 0

    # 3. Participantes certificados (por CH)
    if "Apto por CH" in relatorio.columns:
        participantes_certificados = int((relatorio["Apto por CH"] == "APTO").sum())
    elif "Apto ao Certificado" in relatorio.columns:
        participantes_certificados = int((relatorio["Apto ao Certificado"].astype(str).str.upper().isin(["SIM", "APTO"])).sum())
    else:
        participantes_certificados = 0

    # 4. Carga horária realizada (soma duração de reunião de todos os encontros no formato 00:00:00)
    total_segundos_encontros = 0
    total_segundos_tempo_medio = 0
    for idx, (resumo, df_res) in enumerate(zip(lista_resumos, lista_resultados_encontros)):
        seg = resumo.get("duracao_segundos", 0)
        if seg == 0 and "Duração da Reunião" in df_res.columns:
            seg = max([tempo_em_segundos(t) for t in df_res["Duração da Reunião"]], default=0)
        total_segundos_encontros += seg

        seg_tm = _calcular_tempo_medio_encontro_seg(resumo, df_res)
        total_segundos_tempo_medio += seg_tm

    carga_realizada_str = segundos_em_formato_hh_mm_ss(total_segundos_encontros)
    tempo_medio_realizado_str = segundos_em_formato_hh_mm_ss(total_segundos_tempo_medio)

    # 5. Taxa de certificação (certificados / convidados)
    taxa_certificacao = (
        participantes_certificados / convidados_val
        if convidados_val > 0 else 0.0
    )

    # 6. Participações com Câmera (% de pessoas que ligaram câmera ao menos 1 vez sobre Participantes Totais)
    col_cam = (
        "Câmera" if "Câmera" in df_interacoes_participante.columns
        else ("Presenças com câmera" if "Presenças com câmera" in df_interacoes_participante.columns else None)
    )
    pessoas_camera = int((df_interacoes_participante[col_cam] >= 1).sum()) if col_cam else 0
    participacoes_camera = (
        pessoas_camera / participantes_totais
        if participantes_totais > 0 else 0.0
    )

    # 7. Participações com Microfone (% de pessoas que ligaram microfone ao menos 1 vez sobre Participantes Totais)
    col_mic = (
        "Microfone" if "Microfone" in df_interacoes_participante.columns
        else ("Presenças com microfone" if "Presenças com microfone" in df_interacoes_participante.columns else None)
    )
    pessoas_microfone = int((df_interacoes_participante[col_mic] >= 1).sum()) if col_mic else 0
    participacoes_microfone = (
        pessoas_microfone / participantes_totais
        if participantes_totais > 0 else 0.0
    )

    # 8. Total de reações (soma de todos os encontros e participantes)
    col_reac = (
        "Total Reações" if "Total Reações" in df_interacoes_participante.columns
        else ("Quantidade de reações" if "Quantidade de reações" in df_interacoes_participante.columns else None)
    )
    total_reacoes = int(df_interacoes_participante[col_reac].sum()) if col_reac else 0

    # 9. Total de mãos levantadas (soma de todos os encontros e participantes)
    col_maos = "Mãos levantadas" if "Mãos levantadas" in df_interacoes_participante.columns else None
    total_maos = int(df_interacoes_participante[col_maos].sum()) if col_maos else 0

    linhas = [
        {"Indicador": "Convidados a participar", "Resultado": convidados_val},
        {"Indicador": "Participantes totais", "Resultado": participantes_totais},
        {"Indicador": "Participantes certificados", "Resultado": participantes_certificados},
        {"Indicador": "Carga horária realizada", "Resultado": carga_realizada_str},
        {"Indicador": "Tempo médio de participação", "Resultado": tempo_medio_realizado_str},
        {"Indicador": "Taxa de certificação", "Resultado": taxa_certificacao},
        {"Indicador": "Participações com Câmera", "Resultado": participacoes_camera},
        {"Indicador": "Participações com Microfone", "Resultado": participacoes_microfone},
        {"Indicador": "Total de reações", "Resultado": total_reacoes},
        {"Indicador": "Total de mãos levantadas", "Resultado": total_maos},
    ]

    return pd.DataFrame(linhas)


def criar_aba_custos(
    total_convidados: int,
    relatorio: pd.DataFrame,
    quantidade_encontros: int,
    carga_horaria_horas: float,
    lista_resumos: List[Dict[str, Any]],
    lista_resultados_encontros: List[pd.DataFrame],
    config: Optional[Dict[str, Any]] = None,
) -> pd.DataFrame:
    """
    Cria a Aba 'Custos'.
    Colunas: 'Item', 'Valor'
    Linhas:
    - Convidados a participar
    - Participantes totais
    - Participantes certificados
    - Custo total
    - Quantidade de encontros
    - Carga horária
    - Carga horária realizada
    - Tempo médio de participação realizado
    - Custo por convidado
    - Custo por participante
    - Custo por certificado
    - Custo por encontro
    - Custo por hora prevista
    - Custo por hora realizada
    - Custo por hora de participação média
    """
    convidados_val = int(total_convidados)

    # Participantes totais (presentes em pelo menos 1 encontro)
    if "Presenças" in relatorio.columns:
        participantes_totais = int((relatorio["Presenças"] >= 1).sum())
    elif "Qtd Presenças" in relatorio.columns:
        participantes_totais = int((relatorio["Qtd Presenças"] >= 1).sum())
    elif "Quantidade de presenças" in relatorio.columns:
        participantes_totais = int((relatorio["Quantidade de presenças"] >= 1).sum())
    else:
        participantes_totais = 0

    # Participantes certificados
    if "Apto por CH" in relatorio.columns:
        participantes_certificados = int((relatorio["Apto por CH"] == "APTO").sum())
    elif "Apto ao Certificado" in relatorio.columns:
        participantes_certificados = int((relatorio["Apto ao Certificado"].astype(str).str.upper().isin(["SIM", "APTO"])).sum())
    else:
        participantes_certificados = 0

    # Custo total
    if config and ("CUSTO_TOTAL_EM_REAIS" in config or "CUSTO_TOTAL_REAIS" in config or "CUSTO_TOTAL" in config):
        custo_raw = config.get("CUSTO_TOTAL_EM_REAIS", config.get("CUSTO_TOTAL_REAIS", config.get("CUSTO_TOTAL", 0.0)))
        try:
            if isinstance(custo_raw, str):
                limpo = custo_raw.replace("R$", "").replace("r$", "").strip()
                if "," in limpo and "." in limpo:
                    limpo = limpo.replace(".", "").replace(",", ".")
                elif "," in limpo:
                    limpo = limpo.replace(",", ".")
                custo_total = float(limpo)
            else:
                custo_total = float(custo_raw)
        except (ValueError, TypeError):
            custo_total = 0.0
    else:
        custo_total = 0.0

    qtd_encontros = int(quantidade_encontros)

    # Carga horária prevista (formato HH:MM:SS)
    try:
        ch_segundos = int(float(carga_horaria_horas or 0.0) * 3600)
    except (ValueError, TypeError):
        ch_segundos = 0
    carga_horaria_str = segundos_em_formato_hh_mm_ss(ch_segundos)

    # Carga horária realizada (soma duração real dos encontros no formato HH:MM:SS)
    total_segundos_encontros = 0
    total_segundos_tempo_medio = 0
    for resumo, df_res in zip(lista_resumos, lista_resultados_encontros):
        seg = resumo.get("duracao_segundos", 0)
        if seg == 0 and "Duração da Reunião" in df_res.columns:
            seg = max([tempo_em_segundos(t) for t in df_res["Duração da Reunião"]], default=0)
        total_segundos_encontros += seg

        seg_tm = _calcular_tempo_medio_encontro_seg(resumo, df_res, config=config)
        total_segundos_tempo_medio += seg_tm

    carga_realizada_str = segundos_em_formato_hh_mm_ss(total_segundos_encontros)
    tempo_medio_realizado_str = segundos_em_formato_hh_mm_ss(total_segundos_tempo_medio)

    # Métricas derivadas de custo
    custo_por_convidado = (custo_total / convidados_val) if convidados_val > 0 else 0.0
    custo_por_participante = (custo_total / participantes_totais) if participantes_totais > 0 else 0.0
    custo_por_certificado = (custo_total / participantes_certificados) if participantes_certificados > 0 else 0.0
    custo_por_encontro = (custo_total / qtd_encontros) if qtd_encontros > 0 else 0.0

    # Custo por hora prevista: Custo total / Carga Horária Prevista
    horas_previstas = float(carga_horaria_horas or 0.0)
    if horas_previstas > 0:
        custo_por_hora_prevista = custo_total / horas_previstas
    else:
        custo_por_hora_prevista = 0.0

    # Custo por hora realizada: Custo total / Carga Horária Realizada (com fallback para prevista)
    horas_realizadas = total_segundos_encontros / 3600.0
    if horas_realizadas > 0:
        custo_por_hora_realizada = custo_total / horas_realizadas
    elif horas_previstas > 0:
        custo_por_hora_realizada = custo_por_hora_prevista
    else:
        custo_por_hora_realizada = 0.0

    # Custo por hora de participação média: Custo total / Tempo Médio Acumulado
    horas_tempo_medio = total_segundos_tempo_medio / 3600.0
    if horas_tempo_medio > 0:
        custo_por_hora_participacao = custo_total / horas_tempo_medio
    else:
        custo_por_hora_participacao = 0.0

    linhas = [
        {"Item": "Convidados a participar", "Valor": convidados_val},
        {"Item": "Participantes totais", "Valor": participantes_totais},
        {"Item": "Participantes certificados", "Valor": participantes_certificados},
        {"Item": "Custo total", "Valor": custo_total},
        {"Item": "Quantidade de encontros", "Valor": qtd_encontros},
        {"Item": "Carga horária", "Valor": carga_horaria_str},
        {"Item": "Carga horária realizada", "Valor": carga_realizada_str},
        {"Item": "Tempo médio de participação realizado", "Valor": tempo_medio_realizado_str},
        {"Item": "Custo por convidado", "Valor": custo_por_convidado},
        {"Item": "Custo por participante", "Valor": custo_por_participante},
        {"Item": "Custo por certificado", "Valor": custo_por_certificado},
        {"Item": "Custo por encontro", "Valor": custo_por_encontro},
        {"Item": "Custo por hora prevista", "Valor": custo_por_hora_prevista},
        {"Item": "Custo por hora realizada", "Valor": custo_por_hora_realizada},
        {"Item": "Custo por hora de participação média", "Valor": custo_por_hora_participacao},
    ]

    return pd.DataFrame(linhas)


def formatar_valor_configuracao(chave: str, valor: Any, config: Dict[str, Any]) -> str:
    """Formata valores das variáveis para exibição amigável e não técnica."""
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        if chave == "DATA_FIM":
            dt_inicio = config.get("DATA_INICIO", "")
            return f"{dt_inicio} (Curso de 1 dia)" if dt_inicio else "Curso de 1 dia"
        if chave == "LOCAL_EMISSAO":
            return "Brasília — DF"
        if chave in ("HORARIO_INICIO", "HORARIO_FIM"):
            return "Sem recorte"
        return "-"

    val_str = str(valor).strip()

    if chave in ("HORARIO_INICIO", "HORARIO_FIM"):
        if hasattr(valor, "strftime"):
            return valor.strftime("%H:%M:%S") if valor.second else valor.strftime("%H:%M")
        return val_str

    if chave == "CARGA_HORARIA":
        try:
            num = float(val_str.replace(",", "."))
            if num.is_integer():
                return f"{int(num)} horas" if int(num) != 1 else "1 hora"
            return f"{num} horas"
        except (ValueError, TypeError):
            return f"{val_str} horas"

    elif chave in ("PERCENTUAL_MINIMO_CARGA", "PERCENTUAL_MINIMO_ENCONTROS"):
        try:
            pct_num = float(str(val_str).replace("%", "").replace(",", "."))
            if pct_num > 1.0:
                pct_num = pct_num / 100.0
            pct_100 = pct_num * 100.0
            if pct_100.is_integer():
                return f"{int(pct_100)}%"
            return f"{pct_100:.1f}%"
        except (ValueError, TypeError):
            return f"{val_str}%"

    elif chave == "TEMPO_MINIMO_PRESENCA_MINUTOS":
        try:
            minutos = int(float(str(val_str).replace(",", ".")))
            return f"{minutos} minutos" if minutos != 1 else "1 minuto"
        except (ValueError, TypeError):
            return f"{val_str} minutos"

    elif chave == "DATA_FIM":
        if not val_str or val_str == config.get("DATA_INICIO"):
            dt_inicio = config.get("DATA_INICIO", "")
            return f"{dt_inicio} (Curso de 1 dia)" if dt_inicio else "Curso de 1 dia"
        return val_str

    elif chave == "ASSUNTO_EMAIL":
        nome_curso = config.get("NOME_DO_CURSO", "")
        return val_str.replace("{CURSO}", nome_curso).replace("{curso}", nome_curso)

    elif chave in ("CUSTO_TOTAL_EM_REAIS", "CUSTO_TOTAL_REAIS", "CUSTO_TOTAL"):
        try:
            val_limpo = str(val_str).replace("R$", "").replace("r$", "").strip()
            if "," in val_limpo and "." in val_limpo:
                val_limpo = val_limpo.replace(".", "").replace(",", ".")
            elif "," in val_limpo:
                val_limpo = val_limpo.replace(",", ".")
            val_f = float(val_limpo)
            return f"R$ {val_f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return f"R$ {val_str}"

    return val_str


def _resumir_horarios_encontros(config: Dict[str, Any]) -> Tuple[str, str, str, str]:
    """
    Retorna (val_inicio, val_fim, obs_inicio, obs_fim) para a aba 'Informações do Curso'.
    Se todos os encontros compartilham a mesma janela ou se há apenas horário global, exibe o valor único.
    Se os encontros possuem horários distintos, indica 'Variável por encontro' e detalha nas observações.
    """
    obs_padrao_ini = "Horário inicial considerado para contabilizar a permanência em cada encontro."
    obs_padrao_fim = "Horário final considerado para contabilizar a permanência em cada encontro."

    encontros = config.get("ENCONTROS") or []
    pares = [(str(e.get("horario_inicio", "")).strip(), str(e.get("horario_fim", "")).strip()) for e in encontros]
    pares_preenchidos = [(i, f) for i, f in pares if i or f]

    if not pares_preenchidos:
        return (
            formatar_valor_configuracao("HORARIO_INICIO", config.get("HORARIO_INICIO", ""), config),
            formatar_valor_configuracao("HORARIO_FIM", config.get("HORARIO_FIM", ""), config),
            obs_padrao_ini,
            obs_padrao_fim,
        )

    unicos = set(pares)
    if len(unicos) == 1:
        ini_u, fim_u = next(iter(unicos))
        return (
            formatar_valor_configuracao("HORARIO_INICIO", ini_u, config),
            formatar_valor_configuracao("HORARIO_FIM", fim_u, config),
            obs_padrao_ini,
            obs_padrao_fim,
        )

    detalhes_ini = " | ".join(
        f"Enc. {idx}: {ini or '-'}" for idx, (ini, _) in enumerate(pares, start=1)
    )
    detalhes_fim = " | ".join(
        f"Enc. {idx}: {fim or '-'}" for idx, (_, fim) in enumerate(pares, start=1)
    )
    return (
        "Variável por encontro",
        "Variável por encontro",
        f"Por encontro: {detalhes_ini}",
        f"Por encontro: {detalhes_fim}",
    )


def criar_aba_informacoes_curso(config: Optional[Dict[str, Any]] = None) -> pd.DataFrame:
    """
    Cria DataFrame para a aba 'Informações do Curso'.
    Contém 3 colunas amigáveis:
    - 'Item': Nome descritivo da informação do curso
    - 'Configuração': Valor formatado de modo acessível e não técnico (ex: '75%', '30 minutos')
    - 'Observação': Breve explicação sobre a finalidade da informação
    """
    if not config:
        config = {}

    pct_encontros = config.get("PERCENTUAL_MINIMO_ENCONTROS", 0.75)
    try:
        pct_encontros_num = float(str(pct_encontros).replace("%", "").replace(",", "."))
        if pct_encontros_num > 1.0:
            pct_encontros_num = pct_encontros_num / 100.0
    except (ValueError, TypeError):
        pct_encontros_num = 0.75

    pct_carga = config.get("PERCENTUAL_MINIMO_CARGA", 0.75)
    try:
        pct_carga_num = float(str(pct_carga).replace("%", "").replace(",", "."))
        if pct_carga_num > 1.0:
            pct_carga_num = pct_carga_num / 100.0
    except (ValueError, TypeError):
        pct_carga_num = 0.75

    data_emissao = config.get("DATA_EMISSAO")
    if not data_emissao:
        data_emissao = config.get("DATA_FIM") or config.get("DATA_INICIO") or ""

    val_h_ini, val_h_fim, obs_h_ini, obs_h_fim = _resumir_horarios_encontros(config)

    linhas = [
        {
            "Item": "Nome do Curso",
            "Configuração": formatar_valor_configuracao("NOME_DO_CURSO", config.get("NOME_DO_CURSO", ""), config),
            "Observação": "Nome oficial da capacitação que será impresso no certificado.",
        },
        {
            "Item": "Carga Horária",
            "Configuração": formatar_valor_configuracao("CARGA_HORARIA", config.get("CARGA_HORARIA", ""), config),
            "Observação": "Carga horária total prevista para o curso, utilizada para o cálculo de frequência.",
        },
        {
            "Item": "Data de Início",
            "Configuração": formatar_valor_configuracao("DATA_INICIO", config.get("DATA_INICIO", ""), config),
            "Observação": "Data do primeiro encontro ou início das atividades do treinamento.",
        },
        {
            "Item": "Data de Fim",
            "Configuração": formatar_valor_configuracao("DATA_FIM", config.get("DATA_FIM", ""), config),
            "Observação": "Data de encerramento das atividades.",
        },
        {
            "Item": "Local de Emissão",
            "Configuração": formatar_valor_configuracao("LOCAL_EMISSAO", config.get("LOCAL_EMISSAO", "Brasília — DF"), config),
            "Observação": "Cidade e UF do órgão emissor exibidas no certificado.",
        },
        {
            "Item": "Data de Emissão",
            "Configuração": formatar_valor_configuracao("DATA_EMISSAO", data_emissao, config),
            "Observação": "Data oficial de registro e expedição impressa no certificado.",
        },
        {
            "Item": "Percentual Mínimo de Carga Horária",
            "Configuração": pct_carga_num,
            "Observação": "Percentual mínimo da carga horária exigido para que o participante seja aprovado.",
        },
        {
            "Item": "Percentual Mínimo de Presença em Encontros",
            "Configuração": pct_encontros_num,
            "Observação": "Percentual mínimo de encontros/aulas em que o participante deve comparecer.",
        },
        {
            "Item": "Tempo Mínimo de Permanência na Reunião",
            "Configuração": formatar_valor_configuracao("TEMPO_MINIMO_PRESENCA_MINUTOS", config.get("TEMPO_MINIMO_PRESENCA_MINUTOS", 30), config),
            "Observação": "Tempo mínimo exigido em cada reunião para validação da presença no encontro.",
        },
        {
            "Item": "Horário de Início",
            "Configuração": val_h_ini,
            "Observação": obs_h_ini,
        },
        {
            "Item": "Horário de Fim",
            "Configuração": val_h_fim,
            "Observação": obs_h_fim,
        },
    ]

    return pd.DataFrame(linhas)


def organizar_colunas(relatorio: pd.DataFrame, quantidade_encontros: int = 0) -> pd.DataFrame:
    """
    Ordena as colunas da aba principal 'Certificação' de forma limpa e objetiva.
    Remove as colunas 'Encontro n' e 'Comparecimento n' (que agora ficam na aba 'Presenças').
    """
    colunas_principais = ["Nome", "UORG", "Email"]
    colunas_finais = [
        "Presenças",
        "CH Total",
        "% da CH Total",
        "Apto por CH",
        "Apto por Presença" if "Apto por Presença" in relatorio.columns else "Apto por Encontros",
        "Apto em Ambos",
    ]

    colunas_existentes = [c for c in colunas_principais + colunas_finais if c in relatorio.columns]
    return relatorio[colunas_existentes].copy()


def criar_estilos() -> Dict[str, Any]:
    """Estilos visuais do Excel idênticos aos definidos no notebook."""
    return {
        "cabecalho_fill": PatternFill(fill_type="solid", fgColor="1F4E78"),
        "cabecalho_font": Font(color="FFFFFF", bold=True),
        "linha_par": PatternFill(fill_type="solid", fgColor="F2F2F2"),
        "linha_impar": PatternFill(fill_type="solid", fgColor="FFFFFF"),
        "apto_fill": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "apto_font": Font(color="006100", bold=True),
        "nao_apto_fill": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
        "nao_apto_font": Font(color="9C0006", bold=True),
        "borda_branca": Border(
            left=Side(style="thin", color="FFFFFF"),
            right=Side(style="thin", color="FFFFFF"),
            top=Side(style="thin", color="FFFFFF"),
            bottom=Side(style="thin", color="FFFFFF"),
        ),
    }


def alinhar_colunas(planilha):
    """Aplica alinhamentos e formatos de números (como %) nas células."""
    for celula in planilha[1]:
        celula.alignment = Alignment(horizontal="center", vertical="center")

    for num_linha in range(2, planilha.max_row + 1):
        for celula in planilha[num_linha]:
            nome_col = str(planilha.cell(row=1, column=celula.column).value)
            if (
                nome_col in (
                    "Presenças",
                    "Qtd Presenças",
                    "Quantidade de presenças",
                    "Câmera",
                    "Presenças com câmera",
                    "Microfone",
                    "Presenças com microfone",
                    "Mãos levantadas",
                    "Total Reações",
                    "Quantidade de reações",
                    "ID",
                    "Data",
                    "Carga horária",
                    "Carga horária realizada",
                    "CH Realizada",
                    "Tempo médio de participação",
                    "Tempo Médio",
                    "Presentes",
                    "Ligou a câmera",
                    "Ligou o microfone",
                    "Levantou a mão",
                    "Reações",
                    "CH Total",
                )
                or nome_col.startswith("Encontro ")
                or nome_col.startswith("Comparecimento")
                or nome_col.startswith("Apto")
            ):
                celula.alignment = Alignment(horizontal="center", vertical="center")
            elif nome_col == "% da CH Total":
                celula.alignment = Alignment(horizontal="center", vertical="center")
                celula.number_format = "0.00%"
            elif nome_col in ("% Câmera", "% Microfone", "Cobertura da câmera", "Cobertura do microfone"):
                celula.alignment = Alignment(horizontal="center", vertical="center")
                if isinstance(celula.value, (int, float)):
                    if (celula.value * 100).is_integer():
                        celula.number_format = "0%"
                    else:
                        celula.number_format = "0.0%"
                else:
                    try:
                        v = float(str(celula.value).replace("%", "").strip().replace(",", "."))
                        if v > 1.0:
                            v = v / 100.0
                        celula.value = v
                        celula.number_format = "0.0%"
                    except Exception:
                        pass
            elif nome_col in ("Configuração", "Valor", "Resultado"):
                celula.alignment = Alignment(horizontal="center", vertical="center")
            elif nome_col in ("Observação", "Indicador", "Item"):
                celula.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                celula.alignment = Alignment(horizontal="left", vertical="center")


def formatar_linhas(planilha, estilos):
    """Aplica cores, zebras e bordas."""
    colunas_apto = set()
    colunas_comparecimento = {}

    for celula in planilha[1]:
        nome = str(celula.value)
        if nome.startswith("Apto"):
            colunas_apto.add(celula.column)
        elif nome.startswith("Comparecimento "):
            num_encontro = nome.replace("Comparecimento ", "")
            for outra in planilha[1]:
                if outra.value in (f"Encontro {num_encontro}", f"Dia {num_encontro}"):
                    colunas_comparecimento[celula.column] = outra.column
                    break

    for num_linha, linha in enumerate(planilha.iter_rows(min_row=2), start=2):
        cor_linha = estilos["linha_par"] if num_linha % 2 == 0 else estilos["linha_impar"]
        for celula in linha:
            celula.fill = cor_linha
            if celula.column not in colunas_apto:
                continue

            celula.border = estilos["borda_branca"]
            if celula.value == "APTO":
                celula.fill = estilos["apto_fill"]
                celula.font = estilos["apto_font"]
            elif celula.value == "NÃO APTO":
                celula.fill = estilos["nao_apto_fill"]
                celula.font = estilos["nao_apto_font"]

    for num_linha in range(2, planilha.max_row + 1):
        for col_comp, col_encontro in colunas_comparecimento.items():
            cel_comp = planilha.cell(row=num_linha, column=col_comp)
            cel_encontro = planilha.cell(row=num_linha, column=col_encontro)
            if str(cel_comp.value).strip().upper() == "AUSENTE":
                cel_comp.font = Font(bold=True)
                cel_encontro.font = Font(bold=True)


def formatar_aba_informacoes(planilha, estilos):
    """Aplica formatação específica e refinada para a aba 'Informações do Curso'."""
    borda_celula = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    planilha.row_dimensions[1].height = 28

    for num_linha in range(2, planilha.max_row + 1):
        planilha.row_dimensions[num_linha].height = 24
        cor_fundo = estilos["linha_par"] if num_linha % 2 == 0 else estilos["linha_impar"]

        cel_item = planilha.cell(row=num_linha, column=1)
        cel_conf = planilha.cell(row=num_linha, column=2)
        cel_obs = planilha.cell(row=num_linha, column=3)

        for cel in (cel_item, cel_conf, cel_obs):
            cel.fill = cor_fundo
            cel.border = borda_celula

        cel_item.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
        cel_item.alignment = Alignment(horizontal="left", vertical="center")

        item_nome = str(cel_item.value or "")

        # Formata porcentagens como número nativo do Excel com formatação percentual ("0%").
        # Isso elimina completamente o aviso do Excel:
        # "O número nesta célula é formatado como texto ou preenchido por um apóstrofo".
        if "Percentual" in item_nome:
            try:
                val_limpo = str(cel_conf.value or "").replace("%", "").strip().replace(",", ".")
                num_pct = float(val_limpo)
                if num_pct > 1.0:
                    num_pct = num_pct / 100.0
                cel_conf.value = num_pct
                if (num_pct * 100).is_integer():
                    cel_conf.number_format = "0%"
                else:
                    cel_conf.number_format = "0.0%"
            except (ValueError, TypeError):
                pass
        elif "Carga Horária" in item_nome:
            try:
                val_limpo = str(cel_conf.value or "").replace("horas", "").replace("hora", "").strip().replace(",", ".")
                num_ch = float(val_limpo)
                cel_conf.value = int(num_ch) if num_ch.is_integer() else num_ch
                if num_ch == 1:
                    cel_conf.number_format = '0" hora"'
                elif num_ch.is_integer():
                    cel_conf.number_format = '0" horas"'
                else:
                    cel_conf.number_format = '0.0" horas"'
            except (ValueError, TypeError):
                pass

        cel_conf.alignment = Alignment(horizontal="center", vertical="center")
        cel_conf.font = Font(name="Calibri", size=11, bold=False, color="000000")

        cel_obs.font = Font(name="Calibri", size=10, italic=False, color="333333")
        cel_obs.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)


def formatar_aba_certificacao(planilha, estilos, workbook=None, mapa_cacheados=None):
    """
    Aplica formatação específica e refinada para a aba 'Certificação',
    injetando fórmulas dinâmicas e regras de formatação condicional para as colunas de Aptidão.
    """
    if workbook is None:
        return

    # 1. Identifica colunas pelo cabeçalho (linha 1)
    col_nome = "A"
    col_pres = None
    col_ch = None
    col_pct_ch = None
    col_apto_ch = None
    col_apto_pres = None
    col_apto_ambos = None

    for col_idx in range(1, planilha.max_column + 1):
        val_cab = str(planilha.cell(row=1, column=col_idx).value or "").strip()
        letra = get_column_letter(col_idx)
        if val_cab.lower() in ("nome", "participante", "servidor"):
            col_nome = letra
        elif val_cab == "Presenças":
            col_pres = letra
        elif val_cab == "CH Total":
            col_ch = letra
        elif val_cab == "% da CH Total":
            col_pct_ch = letra
        elif val_cab in ("Apto por CH", "Apto ao Certificado"):
            col_apto_ch = letra
        elif val_cab.startswith("Apto por Presença") or val_cab.startswith("Apto por Encontros"):
            col_apto_pres = letra
        elif val_cab == "Apto em Ambos":
            col_apto_ambos = letra

    # 2. Localiza parâmetros na aba 'Informações do Curso'
    r_ch_info = 3
    r_pct_ch_info = 8
    r_pct_enc_info = 9
    if "Informações do Curso" in workbook.sheetnames:
        ws_info = workbook["Informações do Curso"]
        for r in range(2, ws_info.max_row + 1):
            item_txt = str(ws_info.cell(row=r, column=1).value or "").strip()
            if item_txt == "Carga Horária":
                r_ch_info = r
            elif item_txt == "Percentual Mínimo de Carga Horária":
                r_pct_ch_info = r
            elif item_txt == "Percentual Mínimo de Presença em Encontros":
                r_pct_enc_info = r

    # 3. Localiza 'Quantidade de encontros' na aba 'Custos'
    r_enc_custos = 6
    if "Custos" in workbook.sheetnames:
        ws_custos = workbook["Custos"]
        for r in range(2, ws_custos.max_row + 1):
            item_txt = str(ws_custos.cell(row=r, column=1).value or "").strip()
            if item_txt == "Quantidade de encontros":
                r_enc_custos = r
                break

    # 4. Mapeia colunas de duração e comparecimento na aba 'Presenças'
    cols_comp_pres = []
    cols_dur_pres = []
    if "Presenças" in workbook.sheetnames:
        ws_pres = workbook["Presenças"]
        num_enc = (ws_pres.max_column - 1) // 2
        for idx in range(1, num_enc + 1):
            col_dur_idx = 2 * idx
            col_comp_idx = 2 * idx + 1
            if col_comp_idx <= ws_pres.max_column:
                cols_dur_pres.append(get_column_letter(col_dur_idx))
                cols_comp_pres.append(get_column_letter(col_comp_idx))

    # 5. Injeta fórmulas linha a linha para cada participante
    for num_linha in range(2, planilha.max_row + 1):
        # A. Presenças (conta presenças não 'Ausente' via COUNTIFS em intervalo contínuo)
        if col_pres and cols_comp_pres:
            cel_pres = planilha[f"{col_pres}{num_linha}"]
            if mapa_cacheados is not None and cel_pres.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_pres.coordinate] = cel_pres.value
            col_fim = cols_comp_pres[-1] if cols_comp_pres else cols_dur_pres[-1]
            cel_pres.value = f'=COUNTIFS(\'Presenças\'!B{num_linha}:{col_fim}{num_linha}, "?*", \'Presenças\'!B{num_linha}:{col_fim}{num_linha}, "<>Ausente")'
            cel_pres.number_format = "0"
            cel_pres.alignment = Alignment(horizontal="center", vertical="center")

        # B. CH Total (soma durações via SUMIF filtrando apenas colunas de Encontro no cabeçalho)
        # Isso pula de forma 100% segura as células de comparecimento/justificativa mesmo se contiverem números
        if col_ch and cols_dur_pres:
            cel_ch = planilha[f"{col_ch}{num_linha}"]
            if mapa_cacheados is not None and cel_ch.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_ch.coordinate] = cel_ch.value
            col_fim = cols_comp_pres[-1] if cols_comp_pres else cols_dur_pres[-1]
            cel_ch.value = f"=SUMIF('Presenças'!$B$1:${col_fim}$1, \"Encontro*\", 'Presenças'!B{num_linha}:{col_fim}{num_linha})"
            cel_ch.number_format = "[h]:mm:ss"
            cel_ch.alignment = Alignment(horizontal="center", vertical="center")

        # C. % da CH Total
        if col_pct_ch and col_ch:
            cel_pct = planilha[f"{col_pct_ch}{num_linha}"]
            if mapa_cacheados is not None and cel_pct.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_pct.coordinate] = cel_pct.value
            cel_pct.value = f"=IF('Informações do Curso'!$B${r_ch_info}>0, ({col_ch}{num_linha}*24)/'Informações do Curso'!$B${r_ch_info}, 0)"
            cel_pct.number_format = "0.0%"
            cel_pct.alignment = Alignment(horizontal="center", vertical="center")

        # D. Apto por CH
        if col_apto_ch and col_pct_ch:
            cel_apto_ch_cell = planilha[f"{col_apto_ch}{num_linha}"]
            if mapa_cacheados is not None and cel_apto_ch_cell.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_apto_ch_cell.coordinate] = cel_apto_ch_cell.value
            cel_apto_ch_cell.value = f'=IF({col_pct_ch}{num_linha}>=\'Informações do Curso\'!$B${r_pct_ch_info}, "APTO", "NÃO APTO")'
            cel_apto_ch_cell.alignment = Alignment(horizontal="center", vertical="center")

        # E. Apto por Presença
        if col_apto_pres and col_pres:
            cel_apto_pres_cell = planilha[f"{col_apto_pres}{num_linha}"]
            if mapa_cacheados is not None and cel_apto_pres_cell.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_apto_pres_cell.coordinate] = cel_apto_pres_cell.value
            cel_apto_pres_cell.value = (
                f"=IF({col_pres}{num_linha}>=MAX(1, ROUNDUP('Custos'!$B${r_enc_custos} * 'Informações do Curso'!$B${r_pct_enc_info}, 0)), "
                f'"APTO", "NÃO APTO")'
            )
            cel_apto_pres_cell.alignment = Alignment(horizontal="center", vertical="center")

        # F. Apto em Ambos
        if col_apto_ambos and col_apto_ch and col_apto_pres:
            cel_ambos = planilha[f"{col_apto_ambos}{num_linha}"]
            if mapa_cacheados is not None and cel_ambos.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_ambos.coordinate] = cel_ambos.value
            cel_ambos.value = f'=IF(AND({col_apto_ch}{num_linha}="APTO", {col_apto_pres}{num_linha}="APTO"), "APTO", "NÃO APTO")'
            cel_ambos.alignment = Alignment(horizontal="center", vertical="center")

    # 6. Formatação Condicional para colunas de Aptidão
    if col_apto_ch and col_apto_ambos and planilha.max_row >= 2:
        rule_apto = CellIsRule(
            operator="equal",
            formula=['"APTO"'],
            stopIfTrue=True,
            fill=estilos["apto_fill"],
            font=estilos["apto_font"],
        )
        rule_nao_apto = CellIsRule(
            operator="equal",
            formula=['"NÃO APTO"'],
            stopIfTrue=True,
            fill=estilos["nao_apto_fill"],
            font=estilos["nao_apto_font"],
        )
        faixa = f"{col_apto_ch}2:{col_apto_ambos}{planilha.max_row}"
        planilha.conditional_formatting.add(faixa, rule_apto)
        planilha.conditional_formatting.add(faixa, rule_nao_apto)


def formatar_aba_interacoes_participante(planilha, estilos, workbook=None, mapa_cacheados=None):
    """
    Aplica formatação específica e refinada para a aba 'Interações por participante',
    injetando fórmulas dinâmicas para 'Qtd Presenças', '% Câmera' e '% Microfone'.
    """
    if workbook is None:
        return

    # 1. Identifica colunas pelo cabeçalho
    col_pres = None
    col_cam = None
    col_pct_cam = None
    col_mic = None
    col_pct_mic = None

    for col_idx in range(1, planilha.max_column + 1):
        val_cab = str(planilha.cell(row=1, column=col_idx).value or "").strip()
        letra = get_column_letter(col_idx)
        if val_cab in ("Qtd Presenças", "Quantidade de presenças", "Presenças"):
            col_pres = letra
        elif val_cab in ("Câmera", "Presenças com câmera"):
            col_cam = letra
        elif val_cab in ("% Câmera", "Percentual Câmera"):
            col_pct_cam = letra
        elif val_cab in ("Microfone", "Presenças com microfone"):
            col_mic = letra
        elif val_cab in ("% Microfone", "Percentual Microfone"):
            col_pct_mic = letra

    # 2. Localiza coluna de presenças na aba 'Certificação'
    col_pres_cert = "D"
    if "Certificação" in workbook.sheetnames:
        ws_cert = workbook["Certificação"]
        for col_idx in range(1, ws_cert.max_column + 1):
            if str(ws_cert.cell(row=1, column=col_idx).value or "").strip() == "Presenças":
                col_pres_cert = get_column_letter(col_idx)
                break

    # 3. Injeta fórmulas
    for num_linha in range(2, planilha.max_row + 1):
        # A. Qtd Presenças vinculada à aba Certificação
        if col_pres:
            cel_pres = planilha[f"{col_pres}{num_linha}"]
            if mapa_cacheados is not None and cel_pres.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_pres.coordinate] = cel_pres.value
            cel_pres.value = f"='Certificação'!{col_pres_cert}{num_linha}"
            cel_pres.number_format = "0"
            cel_pres.alignment = Alignment(horizontal="center", vertical="center")

        # B. % Câmera
        if col_pct_cam and col_cam and col_pres:
            cel_pct_cam = planilha[f"{col_pct_cam}{num_linha}"]
            if mapa_cacheados is not None and cel_pct_cam.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_pct_cam.coordinate] = cel_pct_cam.value
            cel_pct_cam.value = f"=IF({col_pres}{num_linha}>0, {col_cam}{num_linha}/{col_pres}{num_linha}, 0)"
            cel_pct_cam.number_format = "0.0%"
            cel_pct_cam.alignment = Alignment(horizontal="center", vertical="center")

        # C. % Microfone
        if col_pct_mic and col_mic and col_pres:
            cel_pct_mic = planilha[f"{col_pct_mic}{num_linha}"]
            if mapa_cacheados is not None and cel_pct_mic.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_pct_mic.coordinate] = cel_pct_mic.value
            cel_pct_mic.value = f"=IF({col_pres}{num_linha}>0, {col_mic}{num_linha}/{col_pres}{num_linha}, 0)"
            cel_pct_mic.number_format = "0.0%"
            cel_pct_mic.alignment = Alignment(horizontal="center", vertical="center")


def formatar_aba_interacoes_encontro(planilha, estilos, workbook=None, mapa_cacheados=None):
    """
    Aplica formatação específica e refinada para a aba 'Interações por encontro',
    injetando fórmulas dinâmicas para 'Tempo Médio' e 'Presentes' (formato 'x de y').
    """
    if workbook is None:
        return

    # 1. Identifica colunas
    col_tm = None
    col_presentes = None

    for col_idx in range(1, planilha.max_column + 1):
        val_cab = str(planilha.cell(row=1, column=col_idx).value or "").strip()
        letra = get_column_letter(col_idx)
        if val_cab in ("Tempo Médio", "Tempo médio"):
            col_tm = letra
        elif val_cab == "Presentes":
            col_presentes = letra

    # 2. Localiza linha 'Convidados a participar' na aba 'Indicadores'
    r_conv = 2
    if "Indicadores" in workbook.sheetnames:
        ws_ind = workbook["Indicadores"]
        for r in range(2, ws_ind.max_row + 1):
            if str(ws_ind.cell(row=r, column=1).value or "").strip() == "Convidados a participar":
                r_conv = r
                break

    # 3. Injeta fórmulas para cada encontro
    # Na aba 'Interações por encontro', a linha r = idx + 1 corresponde ao Encontro idx
    for num_linha in range(2, planilha.max_row + 1):
        idx_encontro = num_linha - 1
        col_dur_letter = get_column_letter(2 * idx_encontro)
        col_comp_letter = get_column_letter(2 * idx_encontro + 1)

        # Tempo Médio
        if col_tm:
            cel_tm = planilha[f"{col_tm}{num_linha}"]
            if mapa_cacheados is not None and cel_tm.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_tm.coordinate] = cel_tm.value
            cel_tm.value = f"=IFERROR(AVERAGEIF('Presenças'!{col_dur_letter}:{col_dur_letter}, \">0\"), 0)"
            cel_tm.number_format = "hh:mm:ss"
            cel_tm.alignment = Alignment(horizontal="center", vertical="center")

        # Presentes ('x de y')
        if col_presentes:
            cel_pres = planilha[f"{col_presentes}{num_linha}"]
            if mapa_cacheados is not None and cel_pres.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_pres.coordinate] = cel_pres.value
            cel_pres.value = (
                f"=COUNTIFS('Presenças'!{col_comp_letter}:{col_comp_letter}, \"<>Ausente\", "
                f"'Presenças'!{col_comp_letter}:{col_comp_letter}, \"<>\") & \" de \" & Indicadores!$B${r_conv}"
            )
            cel_pres.alignment = Alignment(horizontal="center", vertical="center")


def formatar_aba_indicadores(planilha, estilos, workbook=None, mapa_cacheados=None):
    """Aplica formatação específica e refinada para a aba 'Indicadores', injetando fórmulas dinâmicas."""
    borda_celula = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    planilha.row_dimensions[1].height = 28

    # Mapeia linhas pelos nomes dos indicadores
    mapa_linhas = {}
    for num_linha in range(2, planilha.max_row + 1):
        ind_nome = str(planilha.cell(row=num_linha, column=1).value or "").strip()
        mapa_linhas[ind_nome] = num_linha

    r_conv = mapa_linhas.get("Convidados a participar", 2)
    r_tot = mapa_linhas.get("Participantes totais", 3)
    r_cert = mapa_linhas.get("Participantes certificados", 4)
    r_tm = mapa_linhas.get("Tempo médio de participação", 6)
    r_taxa = mapa_linhas.get("Taxa de certificação", 7)
    r_cam = mapa_linhas.get("Participações com Câmera", 8)
    r_mic = mapa_linhas.get("Participações com Microfone", 9)
    r_reac = mapa_linhas.get("Total de reações", 10)
    r_maos = mapa_linhas.get("Total de mãos levantadas", 11)

    # Injeção de fórmulas dinâmicas caso workbook fornecido
    if workbook is not None:
        # 1. Certificação: Convidados a participar, Participantes totais, Participantes certificados e Tempo médio
        if "Certificação" in workbook.sheetnames:
            ws_cert = workbook["Certificação"]

            col_nome = "A"
            col_pres = None
            col_ch = None
            col_apto = None
            for col_idx in range(1, ws_cert.max_column + 1):
                val_cab = str(ws_cert.cell(row=1, column=col_idx).value or "").strip()
                if val_cab.lower() in ("nome", "participante", "servidor"):
                    col_nome = get_column_letter(col_idx)
                elif val_cab == "Presenças":
                    col_pres = get_column_letter(col_idx)
                elif val_cab == "CH Total":
                    col_ch = get_column_letter(col_idx)
                elif val_cab in ("Apto por CH", "Apto ao Certificado"):
                    col_apto = get_column_letter(col_idx)

            if "Convidados a participar" in mapa_linhas:
                cel = planilha.cell(row=r_conv, column=2)
                if mapa_cacheados is not None and cel.value is not None:
                    mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
                cel.value = f"=MAX(0, COUNTA('Certificação'!{col_nome}:{col_nome}) - 1)"
            if col_pres and "Participantes totais" in mapa_linhas:
                cel = planilha.cell(row=r_tot, column=2)
                if mapa_cacheados is not None and cel.value is not None:
                    mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
                cel.value = f'=COUNTIF(\'Certificação\'!{col_pres}:{col_pres}, ">=1")'
            if col_apto and "Participantes certificados" in mapa_linhas:
                cel = planilha.cell(row=r_cert, column=2)
                if mapa_cacheados is not None and cel.value is not None:
                    mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
                cel.value = f'=COUNTIF(\'Certificação\'!{col_apto}:{col_apto}, "APTO")'
            if col_ch and "Tempo médio de participação" in mapa_linhas:
                cel = planilha.cell(row=r_tm, column=2)
                if mapa_cacheados is not None and cel.value is not None:
                    mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
                cel.value = f'=IFERROR(AVERAGEIF(\'Certificação\'!{col_ch}:{col_ch}, ">0"), 0)'

        # 2. Taxa de certificação: Certificados (B4) / Convidados (B2)
        if "Taxa de certificação" in mapa_linhas and "Convidados a participar" in mapa_linhas and "Participantes certificados" in mapa_linhas:
            cel = planilha.cell(row=r_taxa, column=2)
            if mapa_cacheados is not None and cel.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
            cel.value = f"=IF(B{r_conv}>0, B{r_cert}/B{r_conv}, 0)"

        # 3. Interações por participante: Câmera, Microfone, Reações, Mãos
        if "Interações por participante" in workbook.sheetnames:
            ws_part = workbook["Interações por participante"]

            col_cam = None
            col_mic = None
            col_reac = None
            col_maos = None

            for col_idx in range(1, ws_part.max_column + 1):
                val_cab = str(ws_part.cell(row=1, column=col_idx).value or "").strip()
                if val_cab in ("Câmera", "Presenças com câmera"):
                    col_cam = get_column_letter(col_idx)
                elif val_cab in ("Microfone", "Presenças com microfone"):
                    col_mic = get_column_letter(col_idx)
                elif val_cab in ("Total Reações", "Quantidade de reações"):
                    col_reac = get_column_letter(col_idx)
                elif val_cab in ("Mãos levantadas", "Mãos"):
                    col_maos = get_column_letter(col_idx)

            if col_cam and "Participações com Câmera" in mapa_linhas and "Participantes totais" in mapa_linhas:
                cel = planilha.cell(row=r_cam, column=2)
                if mapa_cacheados is not None and cel.value is not None:
                    mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
                cel.value = f'=IF(B{r_tot}>0, COUNTIF(\'Interações por participante\'!{col_cam}:{col_cam}, ">=1")/B{r_tot}, 0)'
            if col_mic and "Participações com Microfone" in mapa_linhas and "Participantes totais" in mapa_linhas:
                cel = planilha.cell(row=r_mic, column=2)
                if mapa_cacheados is not None and cel.value is not None:
                    mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
                cel.value = f'=IF(B{r_tot}>0, COUNTIF(\'Interações por participante\'!{col_mic}:{col_mic}, ">=1")/B{r_tot}, 0)'
            if col_reac and "Total de reações" in mapa_linhas:
                cel = planilha.cell(row=r_reac, column=2)
                if mapa_cacheados is not None and cel.value is not None:
                    mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
                cel.value = f"=SUM('Interações por participante'!{col_reac}:{col_reac})"
            if col_maos and "Total de mãos levantadas" in mapa_linhas:
                cel = planilha.cell(row=r_maos, column=2)
                if mapa_cacheados is not None and cel.value is not None:
                    mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
                cel.value = f"=SUM('Interações por participante'!{col_maos}:{col_maos})"

    for num_linha in range(2, planilha.max_row + 1):
        planilha.row_dimensions[num_linha].height = 24
        cor_fundo = estilos["linha_par"] if num_linha % 2 == 0 else estilos["linha_impar"]

        cel_ind = planilha.cell(row=num_linha, column=1)
        cel_res = planilha.cell(row=num_linha, column=2)

        for cel in (cel_ind, cel_res):
            cel.fill = cor_fundo
            cel.border = borda_celula

        cel_ind.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
        cel_ind.alignment = Alignment(horizontal="left", vertical="center")

        ind_nome = str(cel_ind.value or "")

        # Formata porcentagens como número nativo do Excel com formatação percentual ("0%" ou "0.0%")
        if any(termo in ind_nome for termo in ["Taxa", "Participações", "Porcentagem", "%"]):
            if not str(cel_res.value or "").startswith("="):
                try:
                    val_limpo = str(cel_res.value or "").replace("%", "").strip().replace(",", ".")
                    num_pct = float(val_limpo)
                    if num_pct > 1.0:
                        num_pct = num_pct / 100.0
                    cel_res.value = num_pct
                except (ValueError, TypeError):
                    pass
            cel_res.number_format = "0.0%"
            cel_res.alignment = Alignment(horizontal="center", vertical="center")
            cel_res.font = Font(name="Calibri", size=11, bold=True, color="000000")
        elif "Tempo médio" in ind_nome:
            cel_res.number_format = "hh:mm:ss"
            cel_res.alignment = Alignment(horizontal="center", vertical="center")
            cel_res.font = Font(name="Calibri", size=11, bold=False, color="000000")
        else:
            cel_res.alignment = Alignment(horizontal="center", vertical="center")
            cel_res.font = Font(name="Calibri", size=11, bold=False, color="000000")


def formatar_aba_custos(planilha, estilos, workbook=None, mapa_cacheados=None):
    """Aplica formatação específica e refinada para a aba 'Custos', injetando fórmulas dinâmicas."""
    borda_celula = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )
    planilha.row_dimensions[1].height = 28

    # Mapeia linhas pelos nomes dos itens
    mapa_linhas = {}
    for num_linha in range(2, planilha.max_row + 1):
        item_nome = str(planilha.cell(row=num_linha, column=1).value or "").strip()
        mapa_linhas[item_nome] = num_linha

    r_conv = mapa_linhas.get("Convidados a participar", 2)
    r_tot = mapa_linhas.get("Participantes totais", 3)
    r_cert = mapa_linhas.get("Participantes certificados", 4)
    r_custo = mapa_linhas.get("Custo total", 5)
    r_encontros = mapa_linhas.get("Quantidade de encontros", 6)
    r_ch_prev = mapa_linhas.get("Carga horária", 7)
    r_ch_real = mapa_linhas.get("Carga horária realizada", 8)
    r_tm_real = mapa_linhas.get("Tempo médio de participação realizado", mapa_linhas.get("Tempo médio de participação", 9))

    # Se workbook fornecido e contém a aba Indicadores, vincula convidados, totais, certificados e tempo médio
    if workbook is not None and "Indicadores" in workbook.sheetnames:
        ws_ind = workbook["Indicadores"]
        mapa_ind_linhas = {}
        for r in range(2, ws_ind.max_row + 1):
            ind_nome = str(ws_ind.cell(row=r, column=1).value or "").strip()
            mapa_ind_linhas[ind_nome] = r

        r_ind_conv = mapa_ind_linhas.get("Convidados a participar", 2)
        r_ind_tot = mapa_ind_linhas.get("Participantes totais", 3)
        r_ind_cert = mapa_ind_linhas.get("Participantes certificados", 4)
        r_ind_tm = mapa_ind_linhas.get("Tempo médio de participação", 6)

        if "Convidados a participar" in mapa_linhas:
            cel = planilha.cell(row=r_conv, column=2)
            if mapa_cacheados is not None and cel.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
            cel.value = f"=Indicadores!B{r_ind_conv}"
        if "Participantes totais" in mapa_linhas:
            cel = planilha.cell(row=r_tot, column=2)
            if mapa_cacheados is not None and cel.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
            cel.value = f"=Indicadores!B{r_ind_tot}"
        if "Participantes certificados" in mapa_linhas:
            cel = planilha.cell(row=r_cert, column=2)
            if mapa_cacheados is not None and cel.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
            cel.value = f"=Indicadores!B{r_ind_cert}"
        if "Tempo médio de participação realizado" in mapa_linhas:
            cel = planilha.cell(row=r_tm_real, column=2)
            if mapa_cacheados is not None and cel.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
            cel.value = f"=Indicadores!B{r_ind_tm}"
    elif workbook is not None and "Certificação" in workbook.sheetnames:
        ws_cert = workbook["Certificação"]
        col_nome = "A"
        for col_idx in range(1, ws_cert.max_column + 1):
            val_cab = str(ws_cert.cell(row=1, column=col_idx).value or "").strip()
            if val_cab.lower() in ("nome", "participante", "servidor"):
                col_nome = get_column_letter(col_idx)
                break
        if "Convidados a participar" in mapa_linhas:
            cel = planilha.cell(row=r_conv, column=2)
            if mapa_cacheados is not None and cel.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
            cel.value = f"=MAX(0, COUNTA('Certificação'!{col_nome}:{col_nome}) - 1)"

    # Se workbook fornecido e contém a aba Interações por encontro, vincula Quantidade de encontros
    if workbook is not None and "Interações por encontro" in workbook.sheetnames:
        ws_enc = workbook["Interações por encontro"]
        col_id = "A"
        for col_idx in range(1, ws_enc.max_column + 1):
            val_cab = str(ws_enc.cell(row=1, column=col_idx).value or "").strip()
            if val_cab.lower() in ("id", "encontro", "data"):
                col_id = get_column_letter(col_idx)
                break
        if "Quantidade de encontros" in mapa_linhas:
            cel = planilha.cell(row=r_encontros, column=2)
            if mapa_cacheados is not None and cel.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
            cel.value = f"=MAX(0, COUNTA('Interações por encontro'!{col_id}:{col_id}) - 1)"

    # Se workbook fornecido e contém a aba Informações do Curso, vincula Carga horária
    if workbook is not None and "Informações do Curso" in workbook.sheetnames:
        ws_info = workbook["Informações do Curso"]
        r_ch_info = 3
        for r in range(2, ws_info.max_row + 1):
            if str(ws_info.cell(row=r, column=1).value or "").strip() == "Carga Horária":
                r_ch_info = r
                break
        if "Carga horária" in mapa_linhas:
            cel = planilha.cell(row=r_ch_prev, column=2)
            if mapa_cacheados is not None and cel.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel.coordinate] = cel.value
            cel.value = f"='Informações do Curso'!$B${r_ch_info}/24"

    # Fórmulas de custos derivados
    formulas_custos = {
        "Custo por convidado": f"=IF(B{r_conv}>0, B{r_custo}/B{r_conv}, 0)",
        "Custo por participante": f"=IF(B{r_tot}>0, B{r_custo}/B{r_tot}, 0)",
        "Custo por certificado": f"=IF(B{r_cert}>0, B{r_custo}/B{r_cert}, 0)",
        "Custo por encontro": f"=IF(B{r_encontros}>0, B{r_custo}/B{r_encontros}, 0)",
        "Custo por hora prevista": f"=IF(B{r_ch_prev}*24>0, B{r_custo}/(B{r_ch_prev}*24), 0)",
        "Custo por hora realizada": f"=IF(B{r_ch_real}*24>0, B{r_custo}/(B{r_ch_real}*24), IF(B{r_ch_prev}*24>0, B{r_custo}/(B{r_ch_prev}*24), 0))",
        "Custo por hora de participação média": f"=IF(B{r_tm_real}*24>0, B{r_custo}/(B{r_tm_real}*24), 0)",
    }

    for num_linha in range(2, planilha.max_row + 1):
        planilha.row_dimensions[num_linha].height = 24
        cor_fundo = estilos["linha_par"] if num_linha % 2 == 0 else estilos["linha_impar"]

        cel_item = planilha.cell(row=num_linha, column=1)
        cel_val = planilha.cell(row=num_linha, column=2)

        for cel in (cel_item, cel_val):
            cel.fill = cor_fundo
            cel.border = borda_celula

        cel_item.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
        cel_item.alignment = Alignment(horizontal="left", vertical="center")

        item_nome = str(cel_item.value or "").strip()

        # Injeta fórmula caso seja métrica derivada
        if item_nome in formulas_custos:
            if mapa_cacheados is not None and cel_val.value is not None:
                mapa_cacheados.setdefault(planilha.title, {})[cel_val.coordinate] = cel_val.value
            cel_val.value = formulas_custos[item_nome]

        # Formata custos como moeda brasileira nativa do Excel ("R$" #,##0.00)
        if "Custo" in item_nome:
            if not str(cel_val.value or "").startswith("="):
                try:
                    val_limpo = str(cel_val.value or "").replace("R$", "").replace("r$", "").strip().replace(",", ".")
                    num_moeda = float(val_limpo)
                    cel_val.value = num_moeda
                except (ValueError, TypeError):
                    pass
            cel_val.number_format = '"R$" #,##0.00'
            cel_val.alignment = Alignment(horizontal="center", vertical="center")
            cel_val.font = Font(name="Calibri", size=11, bold=(item_nome == "Custo total"), color="000000")
        elif "Carga horária" in item_nome or "Tempo médio" in item_nome:
            cel_val.number_format = "hh:mm:ss"
            cel_val.alignment = Alignment(horizontal="center", vertical="center")
            cel_val.font = Font(name="Calibri", size=11, bold=False, color="000000")
        else:
            cel_val.alignment = Alignment(horizontal="center", vertical="center")
            cel_val.font = Font(name="Calibri", size=11, bold=False, color="000000")


def formatar_aba_presencas(
    planilha,
    workbook,
    estilos: Dict[str, Any],
    datas_encontros: Optional[List[str]] = None,
):
    """
    Aplica formatação específica e refinada para a aba 'Presenças':
    - Mescla os títulos 'Encontro n' e 'Comparecimento n' em um título único por encontro contendo ID e data (ex: 'Encontro 1 — 14/09/2026').
    - Centraliza as colunas de duração e comparecimento.
    - Aplica negrito nas faltas ('Ausente') e tempo zerado ('00:00:00').
    - Converte durações em tempo nativo do Excel ('hh:mm:ss').
    - Define larguras balanceadas para as colunas.
    """
    if datas_encontros is None and "Interações por encontro" in workbook.sheetnames:
        sheet_int = workbook["Interações por encontro"]
        col_data_idx = None
        for col_idx in range(1, sheet_int.max_column + 1):
            if str(sheet_int.cell(row=1, column=col_idx).value).strip() == "Data":
                col_data_idx = col_idx
                break
        if col_data_idx is not None:
            datas_encontros = []
            for r in range(2, sheet_int.max_row + 1):
                val = sheet_int.cell(row=r, column=col_data_idx).value
                datas_encontros.append(str(val).strip() if val is not None else "")

    num_encontros = (planilha.max_column - 1) // 2
    if num_encontros < 1:
        return

    # 1. Mescla os cabeçalhos dos encontros com ID e data
    for idx in range(1, num_encontros + 1):
        col_inicio = 2 * idx
        col_fim = 2 * idx + 1
        if col_fim > planilha.max_column:
            break

        data_str = ""
        if datas_encontros and idx - 1 < len(datas_encontros):
            data_str = str(datas_encontros[idx - 1] or "").strip()

        titulo = f"Encontro {idx} — {data_str}" if data_str else f"Encontro {idx}"

        planilha.cell(row=1, column=col_inicio, value=titulo)
        planilha.cell(row=1, column=col_fim, value=None)

        faixa = f"{get_column_letter(col_inicio)}1:{get_column_letter(col_fim)}1"
        ja_mesclado = any(str(r) == faixa for r in planilha.merged_cells.ranges)
        if not ja_mesclado:
            planilha.merge_cells(start_row=1, start_column=col_inicio, end_row=1, end_column=col_fim)

    # 2. Estilização do cabeçalho
    for col_idx in range(1, planilha.max_column + 1):
        cel = planilha.cell(row=1, column=col_idx)
        cel.fill = estilos["cabecalho_fill"]
        cel.font = estilos["cabecalho_font"]
        cel.alignment = Alignment(horizontal="center", vertical="center")

    # 3. Formatação das linhas de dados (zebra, alinhamento e destaque de 'Ausente')
    for num_linha in range(2, planilha.max_row + 1):
        cor_linha = estilos["linha_par"] if num_linha % 2 == 0 else estilos["linha_impar"]

        for col_idx in range(1, planilha.max_column + 1):
            cel = planilha.cell(row=num_linha, column=col_idx)
            cel.fill = cor_linha
            if col_idx == 1:
                cel.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cel.alignment = Alignment(horizontal="center", vertical="center")

        for idx in range(1, num_encontros + 1):
            col_dur = 2 * idx
            col_comp = 2 * idx + 1
            if col_comp <= planilha.max_column:
                cel_dur = planilha.cell(row=num_linha, column=col_dur)
                cel_comp = planilha.cell(row=num_linha, column=col_comp)

                if cel_dur.value is not None:
                    if isinstance(cel_dur.value, str) and ":" in cel_dur.value:
                        seg = tempo_em_segundos(cel_dur.value)
                        cel_dur.value = seg / 86400.0
                    cel_dur.number_format = "hh:mm:ss"

                status_comp_u = str(cel_comp.value).strip().upper()
                if status_comp_u == "AUSENTE":
                    cel_comp.font = Font(name="Calibri", size=11, bold=True, color="000000")
                    cel_dur.font = Font(name="Calibri", size=11, bold=True, color="000000")
                elif status_comp_u == "JUSTIFICADO":
                    cel_comp.font = Font(name="Calibri", size=11, italic=True, bold=True, color="9C6500")
                    cel_dur.font = Font(name="Calibri", size=11, italic=True, color="9C6500")
                else:
                    cel_comp.font = Font(name="Calibri", size=11, bold=False, color="000000")
                    cel_dur.font = Font(name="Calibri", size=11, bold=False, color="000000")

    # 4. Largura das colunas ajustada proporcionalmente ao tamanho do título mesclado
    maior_nome = max((len(str(c.value)) for c in planilha["A"] if c.value is not None), default=0)
    planilha.column_dimensions["A"].width = min(max(maior_nome + 3, 12), 50)

    for idx in range(1, num_encontros + 1):
        col_dur = 2 * idx
        col_comp = 2 * idx + 1
        if col_comp > planilha.max_column:
            break

        titulo = str(planilha.cell(row=1, column=col_dur).value or "")
        tam_titulo = len(titulo)

        def _tam_dur(val):
            if isinstance(val, (int, float)):
                return 8
            return len(str(val or ""))

        maior_dur = max(
            (_tam_dur(planilha.cell(row=r, column=col_dur).value) for r in range(2, planilha.max_row + 1)),
            default=8,
        )
        maior_comp = max(
            (len(str(planilha.cell(row=r, column=col_comp).value or "")) for r in range(2, planilha.max_row + 1)),
            default=8,
        )

        min_w_dur = max(maior_dur + 3, 10)
        min_w_comp = max(maior_comp + 3, 11)

        w_titulo_total = tam_titulo + 3
        w_total = max(w_titulo_total, min_w_dur + min_w_comp)

        w_dur = max(min_w_dur, math.ceil(w_total / 2))
        w_comp = max(min_w_comp, w_total - w_dur)

        letra_dur = get_column_letter(col_dur)
        letra_comp = get_column_letter(col_comp)

        planilha.column_dimensions[letra_dur].width = w_dur
        planilha.column_dimensions[letra_comp].width = w_comp


def ajustar_largura_colunas(planilha):
    """Ajusta automaticamente a largura de cada coluna."""
    if planilha.title == "Informações do Curso":
        for coluna in planilha.columns:
            letra = coluna[0].column_letter
            maior = 0
            for c in coluna:
                if c.value is None:
                    continue
                # Se for porcentagem numérica (ex: 0.75), calcula o tamanho baseado na representação formatada (ex: '75%')
                if isinstance(c.value, float) and "Percentual" in str(planilha.cell(row=c.row, column=1).value or ""):
                    tam = len(f"{int(c.value * 100)}%")
                elif isinstance(c.value, (int, float)) and "Carga Horária" in str(planilha.cell(row=c.row, column=1).value or ""):
                    tam = len(f"{int(c.value)} horas")
                else:
                    tam = len(str(c.value))
                if tam > maior:
                    maior = tam

            if letra == "A":
                planilha.column_dimensions[letra].width = max(maior + 4, 38)
            elif letra == "B":
                # Adapta dinamicamente ao maior valor (ex: nome do curso longo), evitando corte de texto
                planilha.column_dimensions[letra].width = max(maior + 6, 32)
            elif letra == "C":
                planilha.column_dimensions[letra].width = min(max(maior + 4, 50), 75)
        return

    if planilha.title in ("Indicadores", "Custos", "Presenças"):
        if planilha.title in ("Indicadores", "Custos"):
            maior = max((len(str(c.value)) for c in planilha["A"] if c.value is not None), default=30)
            planilha.column_dimensions["A"].width = max(maior + 4, 38)
            planilha.column_dimensions["B"].width = 24
        return

    for coluna in planilha.columns:
        maior = max((len(str(c.value)) for c in coluna if c.value is not None), default=0)
        letra = coluna[0].column_letter
        largura = min(maior + 3, 50)
        planilha.column_dimensions[letra].width = max(largura, 12)


def formatar_cabecalho(planilha, estilos):
    """Estiliza linha de cabeçalho."""
    for celula in planilha[1]:
        celula.fill = estilos["cabecalho_fill"]
        celula.font = estilos["cabecalho_font"]


def formatar_workbook(
    workbook,
    datas_encontros: Optional[List[str]] = None,
    mapa_valores_cacheados: Optional[Dict[str, Dict[str, Any]]] = None,
):
    """Aplica formatação a todas as abas da planilha na ordem correta de dependências."""
    estilos = criar_estilos()

    # 1. Aplica formatações gerais básicas (cabeçalho, linhas zebradas, alinhamento, larguras)
    for sheet in workbook.worksheets:
        formatar_cabecalho(sheet, estilos)
        formatar_linhas(sheet, estilos)
        alinhar_colunas(sheet)
        ajustar_largura_colunas(sheet)

    # 2. Formatações específicas por aba com vínculos e fórmulas dinâmicas
    if "Informações do Curso" in workbook.sheetnames:
        formatar_aba_informacoes(workbook["Informações do Curso"], estilos)

    if "Presenças" in workbook.sheetnames:
        formatar_aba_presencas(workbook["Presenças"], workbook, estilos, datas_encontros=datas_encontros)

    if "Certificação" in workbook.sheetnames:
        formatar_aba_certificacao(
            workbook["Certificação"],
            estilos,
            workbook=workbook,
            mapa_cacheados=mapa_valores_cacheados,
        )

    if "Interações por participante" in workbook.sheetnames:
        formatar_aba_interacoes_participante(
            workbook["Interações por participante"],
            estilos,
            workbook=workbook,
            mapa_cacheados=mapa_valores_cacheados,
        )

    if "Indicadores" in workbook.sheetnames:
        formatar_aba_indicadores(
            workbook["Indicadores"],
            estilos,
            workbook=workbook,
            mapa_cacheados=mapa_valores_cacheados,
        )

    if "Interações por encontro" in workbook.sheetnames:
        formatar_aba_interacoes_encontro(
            workbook["Interações por encontro"],
            estilos,
            workbook=workbook,
            mapa_cacheados=mapa_valores_cacheados,
        )

    if "Custos" in workbook.sheetnames:
        formatar_aba_custos(
            workbook["Custos"],
            estilos,
            workbook=workbook,
            mapa_cacheados=mapa_valores_cacheados,
        )


def injetar_valores_cacheados_xml(
    caminho_xlsx: Union[str, Path],
    mapa_celulas_valores: Dict[str, Dict[str, Any]],
):
    """
    Injeta valores pré-calculados pelo Python nas tags <v> do arquivo Excel (.xlsx),
    garantindo que bibliotecas que realizam leitura sem avaliar fórmulas (como pandas
    e openpyxl em modo data_only) leiam os resultados corretos sem quebrar o pipeline,
    ao mesmo tempo em que preserva as fórmulas dinâmicas para recálculo no Microsoft Excel.
    """
    if not mapa_celulas_valores:
        return

    try:
        with abrir_arquivo(caminho_xlsx, "rb") as f:
            conteudo = f.read()

        buf_in = io.BytesIO(conteudo)
        buf_out = io.BytesIO()

        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        ET.register_namespace("", ns)

        with zipfile.ZipFile(buf_in, "r") as zin, zipfile.ZipFile(buf_out, "w") as zout:
            wb_xml = ET.fromstring(zin.read("xl/workbook.xml"))
            sheets = wb_xml.findall(f".//{{{ns}}}sheet")
            mapa_abas = {}
            for idx, s in enumerate(sheets, start=1):
                mapa_abas[s.get("name")] = f"xl/worksheets/sheet{idx}.xml"

            for item in zin.infolist():
                data = zin.read(item.filename)
                for nome_aba, arq_xml in mapa_abas.items():
                    if item.filename == arq_xml and nome_aba in mapa_celulas_valores:
                        celulas = mapa_celulas_valores[nome_aba]
                        root = ET.fromstring(data)
                        modificado = False
                        for c in root.findall(f".//{{{ns}}}c"):
                            coord = c.get("r")
                            if coord in celulas:
                                val = celulas[coord]
                                if val is not None:
                                    is_num = isinstance(val, (int, float))
                                    if is_num:
                                        if "t" in c.attrib:
                                            del c.attrib["t"]
                                    else:
                                        c.set("t", "str")
                                    v_el = c.find(f"{{{ns}}}v")
                                    if v_el is None:
                                        v_el = ET.SubElement(c, f"{{{ns}}}v")
                                    v_el.text = str(val)
                                    modificado = True
                        if modificado:
                            data = ET.tostring(root, xml_declaration=True, encoding="utf-8")
                zout.writestr(item, data)

        with abrir_arquivo(caminho_xlsx, "wb") as f:
            f.write(buf_out.getvalue())
    except Exception:
        pass


def exportar_excel(
    relatorio: pd.DataFrame,
    caminho_saida: Any,
    df_presencas: Optional[pd.DataFrame] = None,
    df_informacoes: Optional[pd.DataFrame] = None,
    df_interacoes_participante: Optional[pd.DataFrame] = None,
    df_interacoes_encontro: Optional[pd.DataFrame] = None,
    df_indicadores: Optional[pd.DataFrame] = None,
    df_custos: Optional[pd.DataFrame] = None,
    datas_encontros: Optional[List[str]] = None,
):
    """
    Salva planilha formatada contendo:
    - Aba principal 'Certificação'
    - Aba detalhada 'Presenças'
    - Aba 'Interações por participante'
    - Aba 'Interações por encontro'
    - Aba 'Indicadores'
    - Aba 'Custos'
    - Aba informativa 'Informações do Curso'
    """
    if isinstance(caminho_saida, pd.DataFrame) and df_presencas is not None:
        aba_presencas = caminho_saida
        caminho_alvo = Path(df_presencas)
    elif isinstance(caminho_saida, (str, Path)):
        caminho_alvo = Path(caminho_saida)
        aba_presencas = df_presencas
    else:
        caminho_alvo = Path(caminho_saida)
        aba_presencas = None

    if datas_encontros is None and df_interacoes_encontro is not None and "Data" in df_interacoes_encontro.columns:
        datas_encontros = [str(d).strip() if pd.notna(d) else "" for d in df_interacoes_encontro["Data"]]

    nome_seguro = sanitizar_nome_arquivo(caminho_alvo.stem) + caminho_alvo.suffix
    caminho_alvo = caminho_alvo.parent / nome_seguro
    garantir_diretorio(caminho_alvo.parent)

    def _gravar(caminho_final):
        mapa_valores_cacheados = {}
        with pd.ExcelWriter(caminho_final, engine="openpyxl") as writer:
            relatorio.to_excel(writer, sheet_name="Certificação", index=False)
            if aba_presencas is not None and not aba_presencas.empty:
                aba_presencas.to_excel(writer, sheet_name="Presenças", index=False)
            if df_interacoes_participante is not None and not df_interacoes_participante.empty:
                df_interacoes_participante.to_excel(writer, sheet_name="Interações por participante", index=False)
            if df_interacoes_encontro is not None and not df_interacoes_encontro.empty:
                df_interacoes_encontro.to_excel(writer, sheet_name="Interações por encontro", index=False)
            if df_indicadores is not None and not df_indicadores.empty:
                df_indicadores.to_excel(writer, sheet_name="Indicadores", index=False)
            if df_custos is not None and not df_custos.empty:
                df_custos.to_excel(writer, sheet_name="Custos", index=False)
            if df_informacoes is not None and not df_informacoes.empty:
                df_informacoes.to_excel(writer, sheet_name="Informações do Curso", index=False)
            formatar_workbook(
                writer.book,
                datas_encontros=datas_encontros,
                mapa_valores_cacheados=mapa_valores_cacheados,
            )

        # Injeta valores pré-calculados nas tags <v> do arquivo gerado
        injetar_valores_cacheados_xml(caminho_final, mapa_valores_cacheados)

    try:
        _gravar(caminho_alvo)
    except (FileNotFoundError, OSError):
        # Fallback para caminho estendido do Windows (\\?\) caso o caminho ultrapasse MAX_PATH
        caminho_ext = caminho_estendido(caminho_alvo)
        _gravar(caminho_ext)


def _resolver_janela_encontro(
    arquivo: Path,
    config: Optional[Dict[str, Any]],
    fallback_ini: Any,
    fallback_fim: Any,
) -> Tuple[Any, Any]:
    """Resolve (horario_inicio, horario_fim) específicos de um encontro ou usa o padrão global."""
    if config and config.get("ENCONTROS"):
        for enc in config["ENCONTROS"]:
            if str(enc.get("arquivo", "")).strip() == arquivo.name:
                ini = str(enc.get("horario_inicio", "")).strip() or None
                fim = str(enc.get("horario_fim", "")).strip() or None
                if ini is not None or fim is not None:
                    return ini, fim
                if fallback_ini is None and fallback_fim is None:
                    return None, None
                return fallback_ini, fallback_fim
    return fallback_ini, fallback_fim


def carregar_matriz_presencas(caminho_relatorio: Path) -> Dict[str, Any]:
    """
    Lê as abas 'Certificação' e 'Presenças' de um relatório Excel gerado
    para exibição interativa de faltas e justificativas na interface.
    """
    from openpyxl import load_workbook
    from core.certificados import carregar_relatorio_presenca

    if not existe_com_fallback(caminho_relatorio):
        raise FileNotFoundError(f"Relatório de presenças não encontrado: {caminho_relatorio}")

    df_cert = carregar_relatorio_presenca(caminho_relatorio)

    try:
        wb = load_workbook(caminho_relatorio, data_only=True, read_only=True)
    except (FileNotFoundError, OSError):
        wb = load_workbook(caminho_estendido(caminho_relatorio), data_only=True, read_only=True)

    titulos_encontros: List[str] = []
    mapa_presencas: Dict[str, List[Dict[str, str]]] = {}
    try:
        if "Presenças" in wb.sheetnames:
            ws = wb["Presenças"]
            rows = list(ws.iter_rows(values_only=True))
            if rows:
                cab = rows[0]
                num_enc = (len(cab) - 1) // 2
                for idx in range(1, num_enc + 1):
                    tit = str(cab[2 * idx - 1] or f"Encontro {idx}").strip()
                    titulos_encontros.append(tit)
                for row in rows[1:]:
                    nome = str(row[0] or "").strip()
                    if not nome:
                        continue
                    itens = []
                    for idx in range(1, num_enc + 1):
                        val_dur = row[2 * idx - 1] if 2 * idx - 1 < len(row) else ""
                        val_comp = row[2 * idx] if 2 * idx < len(row) else "Ausente"
                        if isinstance(val_dur, (int, float)):
                            dur_str = segundos_em_formato_digital(int(round(float(val_dur) * 86400)))
                        elif isinstance(val_dur, time):
                            dur_str = val_dur.strftime("%H:%M:%S")
                        elif isinstance(val_dur, timedelta):
                            dur_str = segundos_em_formato_digital(int(val_dur.total_seconds()))
                        else:
                            dur_str = str(val_dur or "00:00:00")
                        itens.append({
                            "numero": idx,
                            "titulo": titulos_encontros[idx - 1],
                            "duracao": dur_str,
                            "status": str(val_comp or "Ausente").strip(),
                        })
                    mapa_presencas[normalizar_texto(nome)] = itens
    finally:
        wb.close()

    linhas: List[Dict[str, Any]] = []
    for _, row in df_cert.iterrows():
        nome = str(row.get("Nome", "")).strip()
        email = str(row.get("Email", "")).strip() if pd.notna(row.get("Email")) else ""
        uorg = str(row.get("UORG", "")).strip() if pd.notna(row.get("UORG")) else ""
        pct_raw = row.get("% da CH Total", 0.0)
        try:
            pct_val = float(pct_raw) if pd.notna(pct_raw) else 0.0
        except (ValueError, TypeError):
            pct_val = 0.0

        linhas.append({
            "nome": nome,
            "email": email,
            "uorg": uorg,
            "presencas": int(row.get("Presenças", 0) or 0) if pd.notna(row.get("Presenças")) else 0,
            "ch_total": str(row.get("CH Total", "0:00:00") or "0:00:00"),
            "pct_ch": pct_val,
            "apto_ch": str(row.get("Apto por CH", "")).strip(),
            "apto_presenca": str(row.get("Apto por Presença", "")).strip(),
            "apto_ambos": str(row.get("Apto em Ambos", "")).strip(),
            "encontros": mapa_presencas.get(normalizar_texto(nome), []),
        })

    return {
        "titulos_encontros": titulos_encontros,
        "participantes": linhas,
    }


def gerar_relatorio(
    caminho_servidores: Path,
    caminho_convocados: Path,
    arquivos_presenca: List[Path],
    caminho_saida_excel: Path,
    carga_horaria_horas: float,
    tempo_comparecimento_minutos: int,
    percentual_minimo_carga: float = 0.75,
    percentual_minimo_encontros: float = 0.75,
    config: Optional[Dict[str, Any]] = None,
    on_log: Optional[Callable[[str, str], None]] = None,
    on_progress: Optional[Callable[[int, int, str], None]] = None,
) -> pd.DataFrame:
    """Executa o pipeline completo de geração de relatório de presenças."""
    caminho_saida_excel = Path(caminho_saida_excel)
    caminho_convocados = Path(caminho_convocados)
    caminho_servidores = Path(caminho_servidores)

    if config is None:
        candidatos_var = []
        if len(caminho_saida_excel.parents) > 2:
            candidatos_var.append(caminho_saida_excel.parents[2] / "curso.json")
        if len(caminho_convocados.parents) > 1:
            candidatos_var.append(caminho_convocados.parents[1] / "curso.json")
        candidatos_var.append(caminho_convocados.parent / "curso.json")
        candidatos_var.append(Path.cwd() / "curso.json")

        for cand in candidatos_var:
            if cand and existe_com_fallback(cand):
                try:
                    from core.config import obter_configuracao
                    config = obter_configuracao(cand.parent)
                    break
                except Exception:
                    try:
                        from core.config import carregar_curso_json, validar_curso_config
                        cfg_obj = carregar_curso_json(cand)
                        config = validar_curso_config(cfg_obj)
                        break
                    except Exception:
                        pass

        if config is None:
            nome_curso_estimado = caminho_saida_excel.stem.replace("Relatório de Presença - ", "").replace("Relatorio", "").strip()
            config = {
                "NOME_DO_CURSO": nome_curso_estimado or "Curso",
                "CARGA_HORARIA": carga_horaria_horas,
                "CUSTO_TOTAL_EM_REAIS": 0.0,
                "DATA_INICIO": "",
                "DATA_FIM": "",
                "LOCAL_EMISSAO": "Brasília — DF",
                "DATA_EMISSAO": "",
                "PERCENTUAL_MINIMO_CARGA": percentual_minimo_carga,
                "PERCENTUAL_MINIMO_ENCONTROS": percentual_minimo_encontros,
                "TEMPO_MINIMO_PRESENCA_MINUTOS": tempo_comparecimento_minutos,
                "HORARIO_INICIO": "",
                "HORARIO_FIM": "",
                "TM_INICIO": None,
                "TM_FIM": None,
                "IGNORAR_EMAILS": [],
                "ENCONTROS": [],
                "JUSTIFICATIVAS": [],
            }

    emails_ignorados = set()
    if config:
        raw_ign = config.get("IGNORAR_EMAILS") or []
        if isinstance(raw_ign, str):
            raw_ign = [p.strip() for p in re.split(r"[,;\n]", raw_ign) if p.strip()]
        emails_ignorados = {normalizar_email(e) for e in raw_ign if normalizar_email(e)}
        if emails_ignorados:
            _emitir_log(on_log, "info", f"E-mails configurados para ignorar: {len(emails_ignorados)}")

    horario_inicio_cfg = None
    horario_fim_cfg = None
    if config:
        horario_inicio_cfg = config.get("TM_INICIO") if config.get("TM_INICIO") is not None else config.get("HORARIO_INICIO")
        horario_fim_cfg = config.get("TM_FIM") if config.get("TM_FIM") is not None else config.get("HORARIO_FIM")
        if isinstance(horario_inicio_cfg, str) and not horario_inicio_cfg.strip():
            horario_inicio_cfg = None
        if isinstance(horario_fim_cfg, str) and not horario_fim_cfg.strip():
            horario_fim_cfg = None

    _emitir_log(on_log, "info", "Carregando lista global de servidores...")
    servidores = ler_excel_com_fallback(caminho_servidores)

    arquivos_presenca = ordenar_arquivos_presenca(arquivos_presenca)
    quantidade_encontros = len(arquivos_presenca)
    if quantidade_encontros == 0:
        raise ValueError("Nenhum arquivo de presença foi fornecido.")

    total_etapas = quantidade_encontros + 4
    resultados = []
    encontros_presenca = []
    lista_resumos = []
    lista_interacoes = []

    _emitir_log(on_log, "info", f"Processando {quantidade_encontros} lista(s) de presença...")
    for num_encontro, arq in enumerate(arquivos_presenca, start=1):
        _emitir_progresso(on_progress, num_encontro, total_etapas, f"Processando Encontro {num_encontro}: {arq.name}")
        _emitir_log(on_log, "info", f"[Encontro {num_encontro}] Processando: {arq.name}")

        h_ini_enc, h_fim_enc = _resolver_janela_encontro(arq, config, horario_inicio_cfg, horario_fim_cfg)
        recorte_enc_ativo = h_ini_enc is not None or h_fim_enc is not None

        resultado, encontro, _ = processar_encontro(
            arquivo=arq,
            servidores=servidores,
            numero_encontro=num_encontro,
            tempo_comparecimento_minutos=tempo_comparecimento_minutos,
            emails_ignorados=emails_ignorados,
            horario_inicio=h_ini_enc,
            horario_fim=h_fim_enc,
        )
        resumo = ler_resumo_reuniao(arq)
        dados_atividades = ler_atividades_reuniao(arq)
        resumo["usou_secao3"] = dados_atividades is not None
        resumo["recorte_aplicado"] = bool(recorte_enc_ativo and dados_atividades is not None)
        if not resumo.get("data") and dados_atividades is not None and not dados_atividades.empty:
            datas_s3 = [
                dt.date() for dt in (parse_datetime_atividade(v) for v in dados_atividades["Horário de entrada"])
                if dt is not None
            ]
            if datas_s3:
                resumo["data"] = min(datas_s3).strftime("%d/%m/%Y")

        try:
            dados_brutos = ler_presenca(arq)
            if emails_ignorados:
                dados_brutos = filtrar_emails_ignorados(dados_brutos, emails_ignorados)
            interacoes = processar_interacoes_participantes(dados_brutos)
        except ValueError:
            interacoes = pd.DataFrame(columns=[
                "Nome", "Email", "camera_qtd", "camera_ligou",
                "microfone_qtd", "microfone_ligou", "maos_qtd", "reacoes_qtd"
            ])

        resultados.append(resultado)
        encontros_presenca.append(encontro)
        lista_resumos.append(resumo)
        lista_interacoes.append(interacoes)

    todos_resultados = pd.concat(resultados, ignore_index=True)

    _emitir_progresso(on_progress, quantidade_encontros + 1, total_etapas, "Cruzando dados com a lista de convocados...")
    _emitir_log(on_log, "info", "Cruzando dados com a lista de convocados...")
    convocados = preparar_convocados(caminho_convocados, servidores, emails_ignorados=emails_ignorados)

    relatorio = adicionar_total_presencas(convocados, todos_resultados)
    relatorio = adicionar_encontros_presenca(relatorio, encontros_presenca)

    justificativas = config.get("JUSTIFICATIVAS", []) if config else []
    if justificativas:
        _emitir_log(on_log, "info", f"Aplicando {len(justificativas)} justificativa(s) registrada(s)...")
        relatorio = aplicar_justificativas(relatorio, arquivos_presenca, justificativas=justificativas)

    # Recalcula Tempo Médio dos encontros com recorte ativo a partir dos tempos recortados
    for idx_enc, resumo in enumerate(lista_resumos, start=1):
        if resumo.get("recorte_aplicado"):
            col_enc = f"Encontro {idx_enc}"
            if col_enc in relatorio.columns:
                tempos_validos = [tempo_em_segundos(v) for v in relatorio[col_enc] if tempo_em_segundos(v) > 0]
                tm_recortado = int(round(sum(tempos_validos) / len(tempos_validos))) if tempos_validos else 0
                resumo["tempo_medio_recortado_calculado"] = True
                resumo["tempo_medio_recortado_segundos"] = tm_recortado
                resumo["tempo_medio_segundos"] = tm_recortado
                resumo["tempo_medio_str"] = segundos_em_tempo(tm_recortado)
                resumo["tempo_medio_participacao_segundos"] = tm_recortado
                resumo["tempo_medio_participacao_str"] = segundos_em_tempo(tm_recortado)

    _emitir_progresso(on_progress, quantidade_encontros + 2, total_etapas, "Calculando carga horária e critérios de aptidão...")
    _emitir_log(on_log, "info", "Calculando carga horária e critérios de aptidão...")
    relatorio, carga_segundos = adicionar_carga_horaria(
        relatorio=relatorio,
        quantidade_encontros=quantidade_encontros,
        carga_horaria_horas=carga_horaria_horas,
    )

    relatorio = adicionar_aptidao(
        relatorio=relatorio,
        carga_segundos=carga_segundos,
        quantidade_encontros=quantidade_encontros,
        carga_horaria_horas=carga_horaria_horas,
        percentual_minimo_carga=percentual_minimo_carga,
        percentual_minimo_encontros=percentual_minimo_encontros,
    )

    _emitir_progresso(on_progress, quantidade_encontros + 3, total_etapas, "Organizando abas e fórmulas da planilha...")
    _emitir_log(on_log, "info", "Organizando abas e fórmulas da planilha...")
    if "Nome" in relatorio.columns:
        relatorio = relatorio.sort_values(
            by="Nome",
            ascending=True,
            key=lambda col: col.astype(str).str.strip().str.lower(),
        ).reset_index(drop=True)
    elif "Presenças" in relatorio.columns:
        relatorio = relatorio.sort_values(by="Presenças", ascending=False).reset_index(drop=True)

    df_presencas = criar_aba_presencas(relatorio, quantidade_encontros)
    df_certificacao = organizar_colunas(relatorio, quantidade_encontros)
    df_interacoes_participante = criar_aba_interacoes_participante(relatorio, lista_interacoes, resultados)
    df_interacoes_encontro = criar_aba_interacoes_encontro(len(convocados), lista_resumos, lista_interacoes, resultados, config=config)
    df_indicadores = criar_aba_indicadores(len(convocados), relatorio, df_interacoes_participante, df_interacoes_encontro, lista_resumos, resultados)
    df_custos = criar_aba_custos(
        total_convidados=len(convocados),
        relatorio=relatorio,
        quantidade_encontros=quantidade_encontros,
        carga_horaria_horas=carga_horaria_horas,
        lista_resumos=lista_resumos,
        lista_resultados_encontros=resultados,
        config=config,
    )
    df_informacoes = criar_aba_informacoes_curso(config)

    caminho_saida_excel = Path(caminho_saida_excel)
    nome_arquivo_seguro = sanitizar_nome_arquivo(caminho_saida_excel.stem) + caminho_saida_excel.suffix
    caminho_saida_excel = caminho_saida_excel.parent / nome_arquivo_seguro

    datas_encontros = []
    for arq, resumo in zip(arquivos_presenca, lista_resumos):
        dt_res = resumo.get("data")
        if not dt_res:
            try:
                dt, _ = obter_chave_ordenacao_arquivo(arq)
                if dt and dt.year not in (1970, 9999):
                    dt_res = dt.strftime("%d/%m/%Y")
            except Exception:
                pass
        if not dt_res and config and config.get("DATA_INICIO"):
            dt_res = str(config["DATA_INICIO"])
        datas_encontros.append(dt_res or "")

    _emitir_progresso(on_progress, total_etapas, total_etapas, f"Salvando planilha em {caminho_saida_excel.name}...")
    _emitir_log(on_log, "info", f"Formatando e salvando planilha em: {caminho_saida_excel.name}...")
    garantir_diretorio(caminho_saida_excel.parent)
    exportar_excel(
        df_certificacao,
        caminho_saida_excel,
        df_presencas=df_presencas,
        df_informacoes=df_informacoes,
        df_interacoes_participante=df_interacoes_participante,
        df_interacoes_encontro=df_interacoes_encontro,
        df_indicadores=df_indicadores,
        df_custos=df_custos,
        datas_encontros=datas_encontros,
    )

    total_aptos = int((df_certificacao["Apto por CH"] == "APTO").sum())
    total_participantes = len(df_certificacao)
    _emitir_log(
        on_log,
        "sucesso",
        f"Relatório gerado! Convocados: {total_participantes} | Aptos: {total_aptos}",
    )

    return df_certificacao