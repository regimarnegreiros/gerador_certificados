"""
Módulo de configuração estruturada em JSON versionado (curso.json e app_config.json),
validação de regras de negócio e migração automática de esquema entre versões.
"""

from dataclasses import asdict, dataclass, field
from datetime import datetime, time
import json
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional, Tuple, Union

from core.caminhos import abrir_arquivo, existe_com_fallback, garantir_diretorio


CURRENT_CURSO_SCHEMA_VERSION = 1
CURRENT_APP_SCHEMA_VERSION = 1

DEFAULT_LOCAL_EMISSAO = "Brasília — DF"
DEFAULT_ASSUNTO_EMAIL = "Certificado - Curso {CURSO}"
DEFAULT_CORPO_EMAIL = (
    "Prezada(o) {nome},\n\n"
    "Esta é uma mensagem automatizada. Não é necessário respondê-la.\n\n"
    "Encaminhamos, em anexo, o certificado referente à sua participação no curso “{CURSO}”.\n\n"
    "Agradecemos pela participação.\n\n"
    "Atenciosamente,\n\n"
    "Divisão de Capacitação e Movimentação – DIVCAM/SGP/ANM\n"
    "Superintendência de Gestão de Pessoas – SGP/ANM"
)


class ConfigValidationError(Exception):
    """Exceção levantada quando há erro de validação nas configurações do curso."""

    def __init__(self, errors: List[str]):
        super().__init__("\n".join(errors))
        self.errors = errors


def parse_date(date_str: str) -> datetime:
    """Valida e converte string de data DD/MM/AAAA para objeto datetime."""
    clean = str(date_str).strip()
    return datetime.strptime(clean, "%d/%m/%Y")


def parse_time(time_str: str) -> time:
    """Valida e converte string de horário ('HH:MM', 'HH:MM:SS' ou 12h 'AM/PM') para datetime.time."""
    clean = str(time_str).strip()
    m = re.match(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*(AM|PM))?$", clean, re.IGNORECASE)
    if not m:
        raise ValueError(f"Horário inválido: '{time_str}'. Use o formato HH:MM (ex: 08:00).")
    h, mi, s = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
    ampm = m.group(4)
    if ampm:
        if not (1 <= h <= 12):
            raise ValueError(f"Hora inválida no formato 12h: '{time_str}'.")
        if ampm.upper() == "AM" and h == 12:
            h = 0
        elif ampm.upper() == "PM" and h < 12:
            h += 12
    elif not (0 <= h <= 23):
        raise ValueError(f"Hora inválida: '{time_str}'.")
    if not (0 <= mi <= 59) or not (0 <= s <= 59):
        raise ValueError(f"Minutos ou segundos inválidos: '{time_str}'.")
    return time(h, mi, s)


def format_time_str(t: Optional[time]) -> str:
    """Formata datetime.time para 'HH:MM' (ou 'HH:MM:SS' se houver segundos)."""
    if t is None:
        return ""
    return t.strftime("%H:%M:%S") if t.second else t.strftime("%H:%M")


def parse_percentage(val: Any) -> float:
    """Converte valores como 0.75, 75, '75%' ou '0,75' para float entre 0 e 1."""
    if isinstance(val, (int, float)):
        num = float(val)
    else:
        clean = str(val).strip().replace("%", "").replace(",", ".")
        num = float(clean)
    if num > 1.0:
        num = num / 100.0
    if not (0.0 <= num <= 1.0):
        raise ValueError(f"Percentual deve estar entre 0 e 1 (ou 0% e 100%), recebido: {val}")
    return num


def parse_custo(val: Any) -> float:
    """Converte custo informado como número ou texto monetário em float."""
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        num = float(val)
    else:
        limpo = str(val).replace("R$", "").replace("r$", "").strip()
        if not limpo:
            return 0.0
        if "," in limpo and "." in limpo:
            limpo = limpo.replace(".", "").replace(",", ".")
        elif "," in limpo:
            limpo = limpo.replace(",", ".")
        num = float(limpo)
    return num


def normalizar_lista_emails(emails_raw: Any) -> List[str]:
    """Normaliza lista ou texto de e-mails para lista limpa sem duplicidades."""
    if not emails_raw:
        return []
    if isinstance(emails_raw, str):
        itens = [p.strip() for p in re.split(r"[,;\n]", emails_raw) if p.strip()]
    elif isinstance(emails_raw, (list, tuple, set)):
        itens = [str(p).strip() for p in emails_raw if str(p).strip()]
    else:
        return []

    resultado: List[str] = []
    vistos = set()
    for item in itens:
        m = re.search(r"<([^>]+)>", item)
        email = m.group(1).strip() if m else item
        chave = email.lower()
        if email and chave not in vistos:
            vistos.add(chave)
            resultado.append(email)
    return resultado


@dataclass
class EncontroConfig:
    """Configuração e metadados de uma lista de presença individual."""

    arquivo: str
    data_detectada: str = ""
    tempo_medio_segundos: int = 0
    duracao_segundos: int = 0
    horario_inicio: str = ""
    horario_fim: str = ""

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "EncontroConfig":
        return cls(
            arquivo=str(data.get("arquivo", "")).strip(),
            data_detectada=str(data.get("data_detectada", "")).strip(),
            tempo_medio_segundos=int(data.get("tempo_medio_segundos", 0) or 0),
            duracao_segundos=int(data.get("duracao_segundos", 0) or 0),
            horario_inicio=str(data.get("horario_inicio", "")).strip(),
            horario_fim=str(data.get("horario_fim", "")).strip(),
        )


@dataclass
class JustificativaConfig:
    """Registro de falta justificada para um participante em um encontro específico."""

    nome: str
    email: str
    arquivo_encontro: str
    tempo_abonado_segundos: int
    motivo: str = ""

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "JustificativaConfig":
        return cls(
            nome=str(data.get("nome", "")).strip(),
            email=str(data.get("email", "")).strip(),
            arquivo_encontro=str(data.get("arquivo_encontro", "")).strip(),
            tempo_abonado_segundos=max(0, int(data.get("tempo_abonado_segundos", 0) or 0)),
            motivo=str(data.get("motivo", "")).strip(),
        )


@dataclass
class CursoConfig:
    """Modelo completo persistido em curso.json."""

    schema_version: int = CURRENT_CURSO_SCHEMA_VERSION
    id_curso: str = ""
    nome_do_curso: str = ""
    carga_horaria: float = 0.0
    custo_total_em_reais: float = 0.0
    data_inicio: str = ""
    data_fim: str = ""
    local_emissao: str = DEFAULT_LOCAL_EMISSAO
    data_emissao: str = ""
    percentual_minimo_carga: float = 0.75
    percentual_minimo_encontros: float = 0.75
    tempo_minimo_presenca_minutos: int = 30
    ignorar_emails: List[str] = field(default_factory=list)
    assunto_email: str = DEFAULT_ASSUNTO_EMAIL
    corpo_email: str = DEFAULT_CORPO_EMAIL
    convidados_nome_original: str = ""
    encontros: List[EncontroConfig] = field(default_factory=list)
    justificativas: List[JustificativaConfig] = field(default_factory=list)
    campos_editados_manualmente: List[str] = field(default_factory=list)
    criado_em: str = ""
    atualizado_em: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "CursoConfig":
        encontros_raw = data.get("encontros") or []
        justificativas_raw = data.get("justificativas") or []
        campos_manuais = data.get("campos_editados_manualmente") or []

        return cls(
            schema_version=int(data.get("schema_version", CURRENT_CURSO_SCHEMA_VERSION) or CURRENT_CURSO_SCHEMA_VERSION),
            id_curso=str(data.get("id_curso", "")).strip(),
            nome_do_curso=str(data.get("nome_do_curso", "")).strip(),
            carga_horaria=float(data.get("carga_horaria", 0.0) or 0.0),
            custo_total_em_reais=float(data.get("custo_total_em_reais", 0.0) or 0.0),
            data_inicio=str(data.get("data_inicio", "")).strip(),
            data_fim=str(data.get("data_fim", "")).strip(),
            local_emissao=str(data.get("local_emissao", DEFAULT_LOCAL_EMISSAO)).strip() or DEFAULT_LOCAL_EMISSAO,
            data_emissao=str(data.get("data_emissao", "")).strip(),
            percentual_minimo_carga=float(data.get("percentual_minimo_carga", 0.75) if data.get("percentual_minimo_carga") is not None else 0.75),
            percentual_minimo_encontros=float(data.get("percentual_minimo_encontros", 0.75) if data.get("percentual_minimo_encontros") is not None else 0.75),
            tempo_minimo_presenca_minutos=int(data.get("tempo_minimo_presenca_minutos", 30) if data.get("tempo_minimo_presenca_minutos") is not None else 30),
            ignorar_emails=normalizar_lista_emails(data.get("ignorar_emails", [])),
            assunto_email=str(data.get("assunto_email", DEFAULT_ASSUNTO_EMAIL)).strip() or DEFAULT_ASSUNTO_EMAIL,
            corpo_email=str(data.get("corpo_email", DEFAULT_CORPO_EMAIL)),
            convidados_nome_original=str(data.get("convidados_nome_original", "")).strip(),
            encontros=[EncontroConfig.from_dict(e) for e in encontros_raw if isinstance(e, dict)],
            justificativas=[JustificativaConfig.from_dict(j) for j in justificativas_raw if isinstance(j, dict)],
            campos_editados_manualmente=[str(c).strip() for c in campos_manuais if str(c).strip()],
            criado_em=str(data.get("criado_em", "")).strip(),
            atualizado_em=str(data.get("atualizado_em", "")).strip(),
        )


@dataclass
class AppConfig:
    """Configurações globais persistidas em app_config.json."""

    schema_version: int = CURRENT_APP_SCHEMA_VERSION
    ultimo_curso_id: str = ""
    servidores_nome_original: str = ""
    modelo_nome_original: str = ""
    modelo_arquivo: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AppConfig":
        return cls(
            schema_version=int(data.get("schema_version", CURRENT_APP_SCHEMA_VERSION) or CURRENT_APP_SCHEMA_VERSION),
            ultimo_curso_id=str(data.get("ultimo_curso_id", "")).strip(),
            servidores_nome_original=str(data.get("servidores_nome_original", "")).strip(),
            modelo_nome_original=str(data.get("modelo_nome_original", "")).strip(),
            modelo_arquivo=str(data.get("modelo_arquivo", "")).strip(),
        )


MigrationFn = Callable[[Dict[str, Any]], Dict[str, Any]]
MIGRACOES_CURSO: Dict[int, MigrationFn] = {}
MIGRACOES_APP: Dict[int, MigrationFn] = {}


def _executar_migracoes(
    raw: Dict[str, Any],
    versao_alvo: int,
    registro_migracoes: Dict[int, MigrationFn],
) -> Tuple[Dict[str, Any], bool]:
    """Aplica migrações sequenciais v -> v+1 até atingir versao_alvo."""
    dados = dict(raw)
    versao_atual = int(dados.get("schema_version", 1) or 1)
    foi_migrado = "schema_version" not in dados or versao_atual < versao_alvo

    while versao_atual < versao_alvo:
        fn = registro_migracoes.get(versao_atual)
        if fn is not None:
            dados = fn(dados)
        versao_atual += 1
        dados["schema_version"] = versao_atual

    if "schema_version" not in dados:
        dados["schema_version"] = versao_alvo

    return dados, foi_migrado


def migrar_curso_dict(
    raw: Dict[str, Any],
    versao_alvo: int = CURRENT_CURSO_SCHEMA_VERSION,
    registro_migracoes: Optional[Dict[int, MigrationFn]] = None,
) -> Tuple[Dict[str, Any], bool]:
    """Migra um dicionário de curso.json para a versão de esquema atual."""
    reg = registro_migracoes if registro_migracoes is not None else MIGRACOES_CURSO
    return _executar_migracoes(raw, versao_alvo=versao_alvo, registro_migracoes=reg)


def migrar_app_dict(
    raw: Dict[str, Any],
    versao_alvo: int = CURRENT_APP_SCHEMA_VERSION,
    registro_migracoes: Optional[Dict[int, MigrationFn]] = None,
) -> Tuple[Dict[str, Any], bool]:
    """Migra um dicionário de app_config.json para a versão de esquema atual."""
    reg = registro_migracoes if registro_migracoes is not None else MIGRACOES_APP
    return _executar_migracoes(raw, versao_alvo=versao_alvo, registro_migracoes=reg)


def _ler_json(caminho: Path) -> Dict[str, Any]:
    if not existe_com_fallback(caminho):
        raise FileNotFoundError(f"Arquivo JSON não encontrado: {caminho}")
    with abrir_arquivo(caminho, "r", encoding="utf-8-sig") as f:
        dados = json.load(f)
    if not isinstance(dados, dict):
        raise ValueError(f"Estrutura inválida no arquivo JSON: {caminho}")
    return dados


def _gravar_json(caminho: Path, dados: Dict[str, Any]) -> Path:
    caminho = Path(caminho)
    garantir_diretorio(caminho.parent)
    with abrir_arquivo(caminho, "w", encoding="utf-8") as f:
        json.dump(dados, f, ensure_ascii=False, indent=2)
    return caminho


def carregar_curso_json(caminho_json: Path) -> CursoConfig:
    """Lê, migra se necessário e desserializa curso.json em CursoConfig."""
    raw = _ler_json(caminho_json)
    migrado, precisa_salvar = migrar_curso_dict(raw)
    cfg = CursoConfig.from_dict(migrado)
    if precisa_salvar:
        salvar_curso_json(caminho_json, cfg, atualizar_timestamp=False)
    return cfg


def salvar_curso_json(
    caminho_json: Path,
    config: Union[CursoConfig, Dict[str, Any]],
    atualizar_timestamp: bool = True,
) -> Path:
    """Grava CursoConfig (ou dict) em curso.json com schema_version atualizado."""
    cfg = config if isinstance(config, CursoConfig) else CursoConfig.from_dict(config)
    agora = datetime.now().isoformat(timespec="seconds")
    if not cfg.criado_em:
        cfg.criado_em = agora
    if atualizar_timestamp or not cfg.atualizado_em:
        cfg.atualizado_em = agora
    cfg.schema_version = CURRENT_CURSO_SCHEMA_VERSION
    return _gravar_json(caminho_json, cfg.to_dict())


def carregar_app_config(caminho_json: Path) -> AppConfig:
    """Lê app_config.json (retornando configuração padrão se ainda não existir)."""
    if not existe_com_fallback(caminho_json):
        return AppConfig()
    raw = _ler_json(caminho_json)
    migrado, precisa_salvar = migrar_app_dict(raw)
    cfg = AppConfig.from_dict(migrado)
    if precisa_salvar:
        salvar_app_config(caminho_json, cfg)
    return cfg


def salvar_app_config(caminho_json: Path, config: Union[AppConfig, Dict[str, Any]]) -> Path:
    """Grava AppConfig em app_config.json."""
    cfg = config if isinstance(config, AppConfig) else AppConfig.from_dict(config)
    cfg.schema_version = CURRENT_APP_SCHEMA_VERSION
    return _gravar_json(caminho_json, cfg.to_dict())


def _extrair_valor_chave(dados: Dict[str, Any], *chaves: str, padrao: Any = None) -> Any:
    for c in chaves:
        if c in dados and dados[c] is not None:
            return dados[c]
    return padrao


def validar_curso_config(config_input: Union[CursoConfig, Dict[str, Any]]) -> Dict[str, Any]:
    """
    Valida regras de negócio de um CursoConfig (ou dicionário equivalente).
    Retorna dicionário normalizado em MAIÚSCULAS consumido pelos módulos de
    relatório, certificados e e-mails, ou levanta ConfigValidationError.
    """
    raw = config_input.to_dict() if isinstance(config_input, CursoConfig) else dict(config_input)
    erros: List[str] = []
    config: Dict[str, Any] = {}

    # 1. Nome do curso
    nome_curso = str(_extrair_valor_chave(raw, "nome_do_curso", "NOME_DO_CURSO", padrao="")).strip()
    if not nome_curso:
        erros.append("[ERRO] O campo 'Nome do curso' é obrigatório.")
    config["NOME_DO_CURSO"] = nome_curso

    # 2. Carga horária
    carga_raw = _extrair_valor_chave(raw, "carga_horaria", "CARGA_HORARIA", padrao="")
    if carga_raw is None or str(carga_raw).strip() == "":
        erros.append("[ERRO] O campo 'Carga horária' é obrigatório.")
    else:
        try:
            carga = float(str(carga_raw).replace(",", "."))
            if carga <= 0:
                erros.append(f"[ERRO] 'Carga horária' deve ser maior que zero (recebido: {carga_raw}).")
            config["CARGA_HORARIA"] = int(carga) if carga.is_integer() else carga
        except ValueError:
            erros.append(f"[ERRO] 'Carga horária' deve ser um número válido (ex: 16 ou 20), recebido: '{carga_raw}'.")

    # 3. Custo total em reais
    custo_raw = _extrair_valor_chave(raw, "custo_total_em_reais", "CUSTO_TOTAL_EM_REAIS", padrao=0.0)
    try:
        val_custo = parse_custo(custo_raw)
        if val_custo < 0:
            erros.append("[ERRO] 'Custo total em reais' não pode ser negativo.")
        config["CUSTO_TOTAL_EM_REAIS"] = val_custo
    except ValueError:
        erros.append(f"[ERRO] 'Custo total em reais' deve ser um número válido, recebido: '{custo_raw}'.")

    # 4. Data de início
    data_inicio_str = str(_extrair_valor_chave(raw, "data_inicio", "DATA_INICIO", padrao="")).strip()
    dt_inicio = None
    if not data_inicio_str:
        erros.append("[ERRO] O campo 'Data de início' é obrigatório no formato DD/MM/AAAA.")
    else:
        try:
            dt_inicio = parse_date(data_inicio_str)
            config["DATA_INICIO"] = data_inicio_str
            config["DT_INICIO"] = dt_inicio
        except Exception:
            erros.append(f"[ERRO] 'Data de início' inválida: '{data_inicio_str}'. Use DD/MM/AAAA.")

    # 5. Data de término (opcional se curso de 1 dia)
    data_fim_str = str(_extrair_valor_chave(raw, "data_fim", "DATA_FIM", padrao="")).strip()
    if data_fim_str:
        try:
            dt_fim = parse_date(data_fim_str)
            config["DATA_FIM"] = data_fim_str
            config["DT_FIM"] = dt_fim
            if dt_inicio and dt_fim < dt_inicio:
                erros.append(f"[ERRO] 'Data de término' ({data_fim_str}) não pode ser anterior a 'Data de início' ({data_inicio_str}).")
        except Exception:
            erros.append(f"[ERRO] 'Data de término' inválida: '{data_fim_str}'. Use DD/MM/AAAA ou deixe em branco.")
    else:
        config["DATA_FIM"] = ""
        config["DT_FIM"] = None

    # 6. Local e data de emissão
    local_emissao = str(_extrair_valor_chave(raw, "local_emissao", "LOCAL_EMISSAO", padrao=DEFAULT_LOCAL_EMISSAO)).strip()
    config["LOCAL_EMISSAO"] = local_emissao or DEFAULT_LOCAL_EMISSAO

    data_emissao_str = str(_extrair_valor_chave(raw, "data_emissao", "DATA_EMISSAO", padrao="")).strip()
    if not data_emissao_str:
        erros.append("[ERRO] O campo 'Data de emissão' é obrigatório no formato DD/MM/AAAA.")
    else:
        try:
            config["DATA_EMISSAO"] = data_emissao_str
            config["DT_EMISSAO"] = parse_date(data_emissao_str)
        except Exception:
            erros.append(f"[ERRO] 'Data de emissão' inválida: '{data_emissao_str}'. Use DD/MM/AAAA.")

    # 7. Critérios de aprovação
    pct_carga_raw = _extrair_valor_chave(raw, "percentual_minimo_carga", "PERCENTUAL_MINIMO_CARGA", padrao=0.75)
    try:
        config["PERCENTUAL_MINIMO_CARGA"] = parse_percentage(pct_carga_raw)
    except Exception as e:
        erros.append(f"[ERRO] 'Percentual mínimo de carga': {e}")

    pct_enc_raw = _extrair_valor_chave(raw, "percentual_minimo_encontros", "PERCENTUAL_MINIMO_ENCONTROS", padrao=0.75)
    try:
        config["PERCENTUAL_MINIMO_ENCONTROS"] = parse_percentage(pct_enc_raw)
    except Exception as e:
        erros.append(f"[ERRO] 'Percentual mínimo de encontros': {e}")

    tempo_min_raw = _extrair_valor_chave(raw, "tempo_minimo_presenca_minutos", "TEMPO_MINIMO_PRESENCA_MINUTOS", padrao=30)
    try:
        tempo_min = int(tempo_min_raw)
        if tempo_min < 0:
            erros.append("[ERRO] 'Tempo mínimo de presença (minutos)' não pode ser negativo.")
        config["TEMPO_MINIMO_PRESENCA_MINUTOS"] = tempo_min
    except (ValueError, TypeError):
        erros.append(f"[ERRO] 'Tempo mínimo de presença (minutos)' deve ser inteiro, recebido: '{tempo_min_raw}'.")

    # 8. Horários globais opcionais + horários por encontro
    h_ini_str = str(_extrair_valor_chave(raw, "horario_inicio", "HORARIO_INICIO", padrao="")).strip()
    h_fim_str = str(_extrair_valor_chave(raw, "horario_fim", "HORARIO_FIM", padrao="")).strip()
    tm_ini = None
    tm_fim = None
    if h_ini_str:
        try:
            tm_ini = parse_time(h_ini_str)
            config["HORARIO_INICIO"] = format_time_str(tm_ini)
            config["TM_INICIO"] = tm_ini
        except Exception as e:
            erros.append(f"[ERRO] 'Horário de início' inválido ('{h_ini_str}'): {e}")
            config["HORARIO_INICIO"] = h_ini_str
            config["TM_INICIO"] = None
    else:
        config["HORARIO_INICIO"] = ""
        config["TM_INICIO"] = None

    if h_fim_str:
        try:
            tm_fim = parse_time(h_fim_str)
            config["HORARIO_FIM"] = format_time_str(tm_fim)
            config["TM_FIM"] = tm_fim
        except Exception as e:
            erros.append(f"[ERRO] 'Horário de fim' inválido ('{h_fim_str}'): {e}")
            config["HORARIO_FIM"] = h_fim_str
            config["TM_FIM"] = None
    else:
        config["HORARIO_FIM"] = ""
        config["TM_FIM"] = None

    if tm_ini and tm_fim and tm_fim <= tm_ini:
        erros.append(
            f"[ERRO] 'Horário de fim' ({format_time_str(tm_fim)}) deve ser posterior a 'Horário de início' ({format_time_str(tm_ini)})."
        )

    encontros_validados: List[Dict[str, Any]] = []
    for enc in raw.get("encontros") or []:
        enc_obj = enc if isinstance(enc, EncontroConfig) else EncontroConfig.from_dict(enc)
        e_ini_tm = None
        e_fim_tm = None
        e_ini_fmt = ""
        e_fim_fmt = ""

        if enc_obj.horario_inicio:
            try:
                e_ini_tm = parse_time(enc_obj.horario_inicio)
                e_ini_fmt = format_time_str(e_ini_tm)
            except Exception as e:
                erros.append(f"[ERRO] Encontro '{enc_obj.arquivo}': 'Horário de início' inválido ({e}).")
        if enc_obj.horario_fim:
            try:
                e_fim_tm = parse_time(enc_obj.horario_fim)
                e_fim_fmt = format_time_str(e_fim_tm)
            except Exception as e:
                erros.append(f"[ERRO] Encontro '{enc_obj.arquivo}': 'Horário de fim' inválido ({e}).")

        if e_ini_tm and e_fim_tm and e_fim_tm <= e_ini_tm:
            erros.append(
                f"[ERRO] Encontro '{enc_obj.arquivo}': 'Horário de fim' ({e_fim_fmt}) deve ser posterior ao início ({e_ini_fmt})."
            )

        encontros_validados.append({
            "arquivo": enc_obj.arquivo,
            "data_detectada": enc_obj.data_detectada,
            "tempo_medio_segundos": enc_obj.tempo_medio_segundos,
            "duracao_segundos": enc_obj.duracao_segundos,
            "horario_inicio": e_ini_fmt,
            "horario_fim": e_fim_fmt,
        })

    config["ENCONTROS"] = encontros_validados

    # 9. Justificativas
    justificativas_norm: List[Dict[str, Any]] = []
    for just in raw.get("justificativas") or []:
        j_obj = just if isinstance(just, JustificativaConfig) else JustificativaConfig.from_dict(just)
        justificativas_norm.append(asdict(j_obj))
    config["JUSTIFICATIVAS"] = justificativas_norm

    # 10. E-mail
    emails_ign = normalizar_lista_emails(_extrair_valor_chave(raw, "ignorar_emails", "IGNORAR_EMAILS", padrao=[]))
    config["IGNORAR_EMAILS"] = emails_ign

    assunto = str(_extrair_valor_chave(raw, "assunto_email", "ASSUNTO_EMAIL", padrao=DEFAULT_ASSUNTO_EMAIL)).strip()
    config["ASSUNTO_EMAIL"] = assunto or DEFAULT_ASSUNTO_EMAIL

    corpo = str(_extrair_valor_chave(raw, "corpo_email", "CORPO_EMAIL", padrao="")).strip()
    if not corpo:
        erros.append("[ERRO] O corpo do e-mail não pode estar vazio.")
    config["CORPO_EMAIL"] = corpo

    if erros:
        raise ConfigValidationError(erros)

    return config


def obter_configuracao(diretorio_curso: Path) -> Dict[str, Any]:
    """Carrega e valida as configurações do arquivo curso.json de um curso."""
    caminho_json = Path(diretorio_curso) / "curso.json"
    cfg = carregar_curso_json(caminho_json)
    return validar_curso_config(cfg)
