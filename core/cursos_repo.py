"""
Repositório central de cursos, recursos globais compartilhados, pré-leitura de resumos
para autopreenchimento e gerenciamento de justificativas de faltas.
"""

import csv
from datetime import datetime
import io
from pathlib import Path
import re
import shutil
from typing import Any, Dict, List, Optional, Tuple, Union

from core.app_paths import (
    garantir_estrutura_curso,
    garantir_estrutura_raiz,
    localizar_modelo_global,
    obter_caminho_app_config,
    obter_caminho_curso_json,
    obter_caminho_servidores_global,
    obter_pasta_cursos,
    obter_pasta_entrada_curso,
    obter_pasta_global,
    obter_pasta_presencas_curso,
    obter_pasta_saida_curso,
)
from core.caminhos import (
    abrir_arquivo,
    caminho_estendido,
    copiar_arquivo_com_fallback,
    existe_com_fallback,
    listar_diretorio_com_fallback,
    sanitizar_nome_arquivo,
)
from core.config import (
    CursoConfig,
    EncontroConfig,
    JustificativaConfig,
    carregar_app_config,
    carregar_curso_json,
    format_time_str,
    parse_time,
    salvar_app_config,
    salvar_curso_json,
)
from core.validator import listar_arquivos_validos


def _resolver_raiz(raiz: Optional[Union[str, Path]] = None) -> Path:
    return garantir_estrutura_raiz(raiz)


def _limpar_arquivo(caminho: Path) -> None:
    try:
        caminho.unlink()
    except PermissionError:
        raise PermissionError(
            f"Não foi possível remover '{caminho.name}'. "
            "Verifique se o arquivo não está aberto no Excel ou PowerPoint."
        )
    except FileNotFoundError:
        pass
    except OSError:
        try:
            Path(caminho_estendido(caminho)).unlink()
        except FileNotFoundError:
            pass


# ============================================================================
# RECURSOS GLOBAIS COMPARTILHADOS (SERVIDORES E MODELO DE CERTIFICADO)
# ============================================================================

def obter_status_recursos_globais(raiz: Optional[Union[str, Path]] = None) -> Dict[str, Any]:
    """Retorna o estado atual da planilha global de servidores e do modelo de certificado."""
    raiz_path = _resolver_raiz(raiz)
    app_cfg = carregar_app_config(obter_caminho_app_config(raiz_path))

    caminho_serv = obter_caminho_servidores_global(raiz_path)
    tem_serv = existe_com_fallback(caminho_serv) and caminho_serv.is_file()

    caminho_modelo = localizar_modelo_global(raiz_path)
    tem_modelo = caminho_modelo is not None

    return {
        "raiz": raiz_path,
        "servidores_disponivel": tem_serv,
        "servidores_caminho": caminho_serv if tem_serv else None,
        "servidores_nome": app_cfg.servidores_nome_original or (caminho_serv.name if tem_serv else ""),
        "modelo_disponivel": tem_modelo,
        "modelo_caminho": caminho_modelo,
        "modelo_nome": app_cfg.modelo_nome_original or (caminho_modelo.name if caminho_modelo else ""),
    }


def definir_servidores_global(
    origem: Union[str, Path],
    raiz: Optional[Union[str, Path]] = None,
) -> Path:
    """Copia a planilha global de servidores (.xlsx/.xls) para global/servidores.xlsx."""
    origem_path = Path(origem)
    if not existe_com_fallback(origem_path) or not origem_path.is_file():
        raise FileNotFoundError(f"Arquivo de servidores não encontrado: {origem_path}")
    if origem_path.name.startswith("~$"):
        raise ValueError("O arquivo selecionado é um arquivo temporário do Excel (~$). Feche o Excel e selecione o arquivo original.")
    if origem_path.suffix.lower() not in (".xlsx", ".xls"):
        raise ValueError("A lista de servidores deve estar no formato Excel (.xlsx ou .xls).")

    raiz_path = _resolver_raiz(raiz)
    destino = obter_caminho_servidores_global(raiz_path)
    copiar_arquivo_com_fallback(origem_path, destino)

    app_cfg = carregar_app_config(obter_caminho_app_config(raiz_path))
    app_cfg.servidores_nome_original = origem_path.name
    salvar_app_config(obter_caminho_app_config(raiz_path), app_cfg)
    return destino


def definir_modelo_global(
    origem: Union[str, Path],
    raiz: Optional[Union[str, Path]] = None,
) -> Path:
    """Copia o modelo global de certificado (.pptx/.pptm) para global/modelo_certificado.<ext>."""
    origem_path = Path(origem)
    if not existe_com_fallback(origem_path) or not origem_path.is_file():
        raise FileNotFoundError(f"Arquivo de modelo PowerPoint não encontrado: {origem_path}")
    if origem_path.name.startswith("~$"):
        raise ValueError("O arquivo selecionado é um arquivo temporário do PowerPoint (~$). Feche o PowerPoint e selecione o arquivo original.")

    ext = origem_path.suffix.lower()
    if ext not in (".pptx", ".pptm"):
        raise ValueError("O modelo de certificado deve ser uma apresentação PowerPoint (.pptx ou .pptm).")

    raiz_path = _resolver_raiz(raiz)
    pasta_global = obter_pasta_global(raiz_path)
    for ext_ant in (".pptx", ".pptm"):
        antigo = pasta_global / f"modelo_certificado{ext_ant}"
        if existe_com_fallback(antigo):
            _limpar_arquivo(antigo)

    destino = pasta_global / f"modelo_certificado{ext}"
    copiar_arquivo_com_fallback(origem_path, destino)

    app_cfg = carregar_app_config(obter_caminho_app_config(raiz_path))
    app_cfg.modelo_nome_original = origem_path.name
    app_cfg.modelo_arquivo = destino.name
    salvar_app_config(obter_caminho_app_config(raiz_path), app_cfg)
    return destino


def remover_recurso_global(tipo: str, raiz: Optional[Union[str, Path]] = None) -> None:
    """Remove um recurso global ('servidores' ou 'modelo')."""
    raiz_path = _resolver_raiz(raiz)
    app_cfg_path = obter_caminho_app_config(raiz_path)
    app_cfg = carregar_app_config(app_cfg_path)

    if tipo == "servidores":
        caminho = obter_caminho_servidores_global(raiz_path)
        if existe_com_fallback(caminho):
            _limpar_arquivo(caminho)
        app_cfg.servidores_nome_original = ""
    elif tipo == "modelo":
        modelo = localizar_modelo_global(raiz_path)
        if modelo is not None:
            _limpar_arquivo(modelo)
        app_cfg.modelo_nome_original = ""
        app_cfg.modelo_arquivo = ""
    else:
        raise ValueError(f"Tipo de recurso global inválido: {tipo}")

    salvar_app_config(app_cfg_path, app_cfg)


# ============================================================================
# CRUD DE CURSOS
# ============================================================================

def gerar_proximo_id_curso(raiz: Optional[Union[str, Path]] = None) -> str:
    """Gera o próximo identificador curto sequencial de curso (ex: c_2026_001)."""
    raiz_path = _resolver_raiz(raiz)
    pasta_cursos = obter_pasta_cursos(raiz_path)
    ano = datetime.now().year
    prefixo = f"c_{ano}_"

    maior_seq = 0
    if existe_com_fallback(pasta_cursos):
        for item in listar_diretorio_com_fallback(pasta_cursos):
            if item.is_dir() and item.name.startswith(prefixo):
                sufixo = item.name[len(prefixo):]
                if sufixo.isdigit():
                    maior_seq = max(maior_seq, int(sufixo))

    return f"{prefixo}{maior_seq + 1:03d}"


def criar_curso(
    nome_do_curso: str,
    raiz: Optional[Union[str, Path]] = None,
) -> Tuple[Path, CursoConfig]:
    """Cria a pasta isolada de um novo curso com estrutura completa e curso.json inicial."""
    nome_limpo = str(nome_do_curso or "").strip()
    if not nome_limpo:
        raise ValueError("O nome do curso não pode estar vazio.")

    raiz_path = _resolver_raiz(raiz)
    id_curso = gerar_proximo_id_curso(raiz_path)
    pasta_curso = obter_pasta_cursos(raiz_path) / id_curso
    garantir_estrutura_curso(pasta_curso)

    cfg = CursoConfig(
        id_curso=id_curso,
        nome_do_curso=nome_limpo,
    )
    salvar_curso_json(obter_caminho_curso_json(pasta_curso), cfg)

    app_cfg_path = obter_caminho_app_config(raiz_path)
    app_cfg = carregar_app_config(app_cfg_path)
    app_cfg.ultimo_curso_id = id_curso
    salvar_app_config(app_cfg_path, app_cfg)

    return pasta_curso, cfg


def obter_curso(
    id_curso: str,
    raiz: Optional[Union[str, Path]] = None,
) -> Tuple[Path, CursoConfig]:
    """Carrega um curso pelo seu ID."""
    raiz_path = _resolver_raiz(raiz)
    pasta_curso = obter_pasta_cursos(raiz_path) / id_curso
    if not existe_com_fallback(pasta_curso):
        raise FileNotFoundError(f"Curso não encontrado: {id_curso}")
    cfg = carregar_curso_json(obter_caminho_curso_json(pasta_curso))
    if not cfg.id_curso:
        cfg.id_curso = id_curso
    return pasta_curso, cfg


def salvar_curso(diretorio_curso: Path, config: CursoConfig) -> Path:
    """Persiste alterações de um CursoConfig em curso.json."""
    return salvar_curso_json(obter_caminho_curso_json(diretorio_curso), config)


def listar_cursos(raiz: Optional[Union[str, Path]] = None) -> List[Dict[str, Any]]:
    """Lista todos os cursos cadastrados com metadados e indicadores de progresso."""
    raiz_path = _resolver_raiz(raiz)
    pasta_cursos = obter_pasta_cursos(raiz_path)
    if not existe_com_fallback(pasta_cursos):
        return []

    cursos: List[Dict[str, Any]] = []
    for item in listar_diretorio_com_fallback(pasta_cursos):
        if not item.is_dir():
            continue
        caminho_json = obter_caminho_curso_json(item)
        if not existe_com_fallback(caminho_json):
            continue
        try:
            cfg = carregar_curso_json(caminho_json)
        except Exception:
            continue

        pasta_ent = obter_pasta_entrada_curso(item)
        pasta_pres = obter_pasta_presencas_curso(item)
        pasta_saida = obter_pasta_saida_curso(item)

        convidados = listar_arquivos_validos(pasta_ent, (".csv",))
        presencas = listar_arquivos_validos(pasta_pres, (".csv", ".tsv", ".txt"))
        relatorios = listar_arquivos_validos(pasta_saida / "relatorio_presencas", (".xlsx",))
        pdfs = listar_arquivos_validos(pasta_saida / "certificados_pdf", (".pdf",))

        cursos.append({
            "id_curso": cfg.id_curso or item.name,
            "nome_do_curso": cfg.nome_do_curso or item.name,
            "pasta": item,
            "data_inicio": cfg.data_inicio,
            "data_fim": cfg.data_fim,
            "carga_horaria": cfg.carga_horaria,
            "tem_convidados": len(convidados) > 0,
            "qtd_presencas": len(presencas),
            "tem_relatorio": len(relatorios) > 0,
            "qtd_certificados": len(pdfs),
            "atualizado_em": cfg.atualizado_em,
        })

    cursos.sort(key=lambda c: c.get("atualizado_em") or c.get("id_curso"), reverse=True)
    return cursos


def excluir_curso(diretorio_curso: Path) -> None:
    """Remove completamente a pasta do curso informado."""
    pasta = Path(diretorio_curso)
    if not existe_com_fallback(pasta):
        return
    try:
        shutil.rmtree(pasta)
    except PermissionError:
        raise PermissionError(
            f"Não foi possível excluir o curso '{pasta.name}'. "
            "Feche qualquer planilha, PDF ou PowerPoint aberto desta pasta e tente novamente."
        )
    except OSError:
        shutil.rmtree(caminho_estendido(pasta))


# ============================================================================
# LEITURA RÁPIDA DE RESUMO, AUTOPREENCHIMENTO E GESTÃO DE ENTRADAS
# ============================================================================

def _extrair_hora_termino_resumo(arquivo: Path) -> str:
    """Extrai o valor bruto do campo 'Hora de término' / 'Hora de fim' da Seção 1."""
    from core.presenca import normalizar_texto

    for enc in ("utf-16", "utf-8-sig", "utf-8", "latin1"):
        try:
            with abrir_arquivo(arquivo, mode="r", encoding=enc, errors="replace") as f:
                linhas = f.readlines()
        except Exception:
            continue

        if not linhas:
            continue

        for linha in linhas:
            limpa = linha.strip().strip('"\'')
            if not limpa:
                continue
            if limpa.lower().startswith("2. participantes") or "2. participantes" in limpa.lower():
                break

            partes = None
            for sep in ("\t", ";", ","):
                try:
                    row = next(csv.reader(io.StringIO(limpa), delimiter=sep))
                    if len(row) >= 2 and row[0].strip():
                        partes = [row[0].strip(), row[1].strip()]
                        break
                except Exception:
                    continue
            if not partes:
                continue

            chave_norm = normalizar_texto(partes[0])
            if "HORA DE TERMINO" in chave_norm or "HORA DO TERMINO" in chave_norm or "HORA DE FIM" in chave_norm or "HORA DO FIM" in chave_norm:
                return partes[1].strip().strip('"\'')
        break

    return ""


def _extrair_hora_de_texto_resumo(valor_raw: str) -> str:
    from core.presenca import extrair_datetime_resumo

    if not valor_raw:
        return ""
    dt = extrair_datetime_resumo(valor_raw)
    if dt is not None:
        return format_time_str(dt.time())
    m = re.search(r"(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?)", valor_raw, re.IGNORECASE)
    if m:
        try:
            return format_time_str(parse_time(m.group(1)))
        except Exception:
            pass
    return ""


def extrair_metadados_resumo_presenca(arquivo: Path) -> EncontroConfig:
    """
    Lê rapidamente a Seção '1. Resumo' de uma lista de presença do Teams
    para pré-preenchimento de data, tempo médio e horários de início/término.
    """
    from core.presenca import extrair_datetime_resumo, ler_resumo_reuniao

    resumo = ler_resumo_reuniao(arquivo)
    data_str = resumo.get("data", "") or ""

    hora_ini_raw = resumo.get("hora_inicio_raw", "") or ""
    if not data_str and hora_ini_raw:
        dt_ini = extrair_datetime_resumo(hora_ini_raw)
        if dt_ini is not None:
            data_str = dt_ini.strftime("%d/%m/%Y")

    horario_inicio = _extrair_hora_de_texto_resumo(hora_ini_raw)
    hora_fim_raw = resumo.get("hora_termino_raw", "") or _extrair_hora_termino_resumo(arquivo)
    horario_fim = _extrair_hora_de_texto_resumo(hora_fim_raw)

    return EncontroConfig(
        arquivo=arquivo.name,
        data_detectada=data_str,
        tempo_medio_segundos=int(resumo.get("tempo_medio_segundos", 0) or 0),
        duracao_segundos=int(resumo.get("duracao_segundos", 0) or 0),
        horario_inicio=horario_inicio,
        horario_fim=horario_fim,
    )


def _segundos_para_horas_arredondadas(total_segundos: int) -> float:
    if total_segundos <= 0:
        return 0.0
    horas = round(total_segundos / 3600.0, 2)
    return int(horas) if horas.is_integer() else horas


def sincronizar_encontros_e_autopreencher(
    diretorio_curso: Path,
    forcar_recalculo: bool = False,
) -> CursoConfig:
    """
    Reconcilia os arquivos em entrada/presencas/ com curso.json em ordem cronológica,
    pré-preenchendo carga horária, datas e horários de início/fim a partir do '1. Resumo'.
    """
    from core.presenca import ordenar_arquivos_presenca

    caminho_json = obter_caminho_curso_json(diretorio_curso)
    cfg = carregar_curso_json(caminho_json)
    pasta_pres = obter_pasta_presencas_curso(diretorio_curso)

    arquivos = listar_arquivos_validos(pasta_pres, (".csv", ".tsv", ".txt"))
    arquivos_ordenados = ordenar_arquivos_presenca(arquivos) if arquivos else []

    existentes_por_arq = {e.arquivo: e for e in cfg.encontros}
    novos_encontros: List[EncontroConfig] = []

    for arq in arquivos_ordenados:
        meta = extrair_metadados_resumo_presenca(arq)
        anterior = existentes_por_arq.get(arq.name)
        if anterior is not None and not forcar_recalculo:
            if anterior.horario_inicio:
                meta.horario_inicio = anterior.horario_inicio
            if anterior.horario_fim:
                meta.horario_fim = anterior.horario_fim
        novos_encontros.append(meta)

    cfg.encontros = novos_encontros

    if forcar_recalculo:
        cfg.campos_editados_manualmente = []

    editados = set(cfg.campos_editados_manualmente)

    if novos_encontros:
        total_medio_seg = sum(e.tempo_medio_segundos for e in novos_encontros)
        if "carga_horaria" not in editados and total_medio_seg > 0:
            cfg.carga_horaria = _segundos_para_horas_arredondadas(total_medio_seg)

        datas_validas = [e.data_detectada for e in novos_encontros if e.data_detectada]
        if datas_validas:
            if "data_inicio" not in editados:
                cfg.data_inicio = datas_validas[0]
            if "data_fim" not in editados:
                cfg.data_fim = datas_validas[-1]
            if "data_emissao" not in editados:
                cfg.data_emissao = datas_validas[-1]
    else:
        if "carga_horaria" not in editados:
            cfg.carga_horaria = 0.0
        if "data_inicio" not in editados:
            cfg.data_inicio = ""
        if "data_fim" not in editados:
            cfg.data_fim = ""
        if "data_emissao" not in editados:
            cfg.data_emissao = ""

    salvar_curso_json(caminho_json, cfg)
    return cfg


def importar_convidados(diretorio_curso: Path, origem: Union[str, Path]) -> CursoConfig:
    """Importa a lista de convidados (.csv) copiando para entrada/convidados.csv."""
    origem_path = Path(origem)
    if not existe_com_fallback(origem_path) or not origem_path.is_file():
        raise FileNotFoundError(f"Arquivo de convidados não encontrado: {origem_path}")
    if origem_path.name.startswith("~$"):
        raise ValueError("Arquivo temporário inválido (~$).")
    if origem_path.suffix.lower() != ".csv":
        raise ValueError("A lista de convidados deve estar no formato .csv.")

    garantir_estrutura_curso(diretorio_curso)
    pasta_ent = obter_pasta_entrada_curso(diretorio_curso)
    for antigo in listar_arquivos_validos(pasta_ent, (".csv",)):
        _limpar_arquivo(antigo)

    destino = pasta_ent / "convidados.csv"
    copiar_arquivo_com_fallback(origem_path, destino)

    caminho_json = obter_caminho_curso_json(diretorio_curso)
    cfg = carregar_curso_json(caminho_json)
    cfg.convidados_nome_original = origem_path.name
    salvar_curso_json(caminho_json, cfg)
    return cfg


def remover_convidados(diretorio_curso: Path) -> CursoConfig:
    """Remove a lista de convidados importada do curso."""
    pasta_ent = obter_pasta_entrada_curso(diretorio_curso)
    for arq in listar_arquivos_validos(pasta_ent, (".csv",)):
        _limpar_arquivo(arq)

    caminho_json = obter_caminho_curso_json(diretorio_curso)
    cfg = carregar_curso_json(caminho_json)
    cfg.convidados_nome_original = ""
    salvar_curso_json(caminho_json, cfg)
    return cfg


def importar_listas_presenca(
    diretorio_curso: Path,
    origens: List[Union[str, Path]],
    forcar_recalculo: bool = False,
) -> CursoConfig:
    """Copia uma ou mais listas de presença (.csv/.tsv/.txt) para entrada/presencas/ e autopreenche."""
    garantir_estrutura_curso(diretorio_curso)
    pasta_pres = obter_pasta_presencas_curso(diretorio_curso)

    for item in origens:
        origem_path = Path(item)
        if not existe_com_fallback(origem_path) or not origem_path.is_file():
            raise FileNotFoundError(f"Lista de presença não encontrada: {origem_path}")
        if origem_path.name.startswith("~$"):
            continue
        ext = origem_path.suffix.lower()
        if ext not in (".csv", ".tsv", ".txt"):
            raise ValueError(f"Formato inválido para lista de presença '{origem_path.name}'. Use .csv, .tsv ou .txt.")

        nome_limpo = sanitizar_nome_arquivo(origem_path.stem) + ext
        destino = pasta_pres / nome_limpo
        copiar_arquivo_com_fallback(origem_path, destino)

    return sincronizar_encontros_e_autopreencher(diretorio_curso, forcar_recalculo=forcar_recalculo)


def remover_lista_presenca(diretorio_curso: Path, nome_arquivo: str) -> CursoConfig:
    """Remove uma lista de presença e suas justificativas associadas, atualizando o curso."""
    pasta_pres = obter_pasta_presencas_curso(diretorio_curso)
    alvo = pasta_pres / nome_arquivo
    if existe_com_fallback(alvo):
        _limpar_arquivo(alvo)

    caminho_json = obter_caminho_curso_json(diretorio_curso)
    cfg = carregar_curso_json(caminho_json)
    cfg.justificativas = [j for j in cfg.justificativas if j.arquivo_encontro != nome_arquivo]
    salvar_curso_json(caminho_json, cfg)

    return sincronizar_encontros_e_autopreencher(diretorio_curso)


def atualizar_horario_encontro(
    diretorio_curso: Path,
    nome_arquivo: str,
    horario_inicio: str,
    horario_fim: str,
) -> CursoConfig:
    """Atualiza o horário de início e término de um encontro específico."""
    caminho_json = obter_caminho_curso_json(diretorio_curso)
    cfg = carregar_curso_json(caminho_json)

    ini_fmt = format_time_str(parse_time(horario_inicio)) if str(horario_inicio).strip() else ""
    fim_fmt = format_time_str(parse_time(horario_fim)) if str(horario_fim).strip() else ""
    if ini_fmt and fim_fmt and parse_time(fim_fmt) <= parse_time(ini_fmt):
        raise ValueError("O horário de término deve ser posterior ao horário de início.")

    encontrado = False
    for enc in cfg.encontros:
        if enc.arquivo == nome_arquivo:
            enc.horario_inicio = ini_fmt
            enc.horario_fim = fim_fmt
            encontrado = True
            break

    if not encontrado:
        raise FileNotFoundError(f"Encontro não encontrado no curso: {nome_arquivo}")

    salvar_curso_json(caminho_json, cfg)
    return cfg


def aplicar_horario_para_todos_encontros(
    diretorio_curso: Path,
    horario_inicio: str,
    horario_fim: str,
) -> CursoConfig:
    """Define o mesmo horário de início e fim para todas as listas de presença do curso."""
    caminho_json = obter_caminho_curso_json(diretorio_curso)
    cfg = carregar_curso_json(caminho_json)

    ini_fmt = format_time_str(parse_time(horario_inicio)) if str(horario_inicio).strip() else ""
    fim_fmt = format_time_str(parse_time(horario_fim)) if str(horario_fim).strip() else ""
    if ini_fmt and fim_fmt and parse_time(fim_fmt) <= parse_time(ini_fmt):
        raise ValueError("O horário de término deve ser posterior ao horário de início.")

    for enc in cfg.encontros:
        enc.horario_inicio = ini_fmt
        enc.horario_fim = fim_fmt

    salvar_curso_json(caminho_json, cfg)
    return cfg


# ============================================================================
# JUSTIFICATIVAS DE FALTAS
# ============================================================================

def calcular_tempo_sugerido_encontro(config: CursoConfig, arquivo_encontro: str) -> int:
    """
    Calcula o tempo padrão sugerido (em segundos) para abonar uma falta justificada:
    1. Diferença entre horario_fim e horario_inicio do encontro
    2. Duração da reunião lida de '1. Resumo'
    3. Carga horária total dividida pela quantidade de encontros
    """
    enc = next((e for e in config.encontros if e.arquivo == arquivo_encontro), None)
    if enc is not None:
        if enc.horario_inicio and enc.horario_fim:
            try:
                t_ini = parse_time(enc.horario_inicio)
                t_fim = parse_time(enc.horario_fim)
                dt_base = datetime(2000, 1, 1)
                diff = int((datetime.combine(dt_base.date(), t_fim) - datetime.combine(dt_base.date(), t_ini)).total_seconds())
                if diff > 0:
                    return diff
            except Exception:
                pass
        if enc.duracao_segundos > 0:
            return int(enc.duracao_segundos)

    qtd = len(config.encontros)
    if config.carga_horaria > 0 and qtd > 0:
        return int(round((float(config.carga_horaria) * 3600.0) / qtd))

    return 0


def _chave_participante_justificativa(nome: str, email: str, arquivo_encontro: str) -> Tuple[str, str]:
    id_part = email.strip().lower() if email and email.strip() else nome.strip().upper()
    return id_part, arquivo_encontro.strip()


def salvar_justificativa(
    diretorio_curso: Path,
    nome: str,
    email: str,
    arquivo_encontro: str,
    tempo_abonado_segundos: int,
    motivo: str = "",
) -> CursoConfig:
    """Adiciona ou atualiza uma justificativa de falta no curso.json."""
    if tempo_abonado_segundos < 0:
        raise ValueError("O tempo abonado não pode ser negativo.")

    caminho_json = obter_caminho_curso_json(diretorio_curso)
    cfg = carregar_curso_json(caminho_json)
    chave_alvo = _chave_participante_justificativa(nome, email, arquivo_encontro)

    nova = JustificativaConfig(
        nome=nome.strip(),
        email=email.strip(),
        arquivo_encontro=arquivo_encontro.strip(),
        tempo_abonado_segundos=int(tempo_abonado_segundos),
        motivo=motivo.strip(),
    )

    cfg.justificativas = [
        j for j in cfg.justificativas
        if _chave_participante_justificativa(j.nome, j.email, j.arquivo_encontro) != chave_alvo
    ]
    cfg.justificativas.append(nova)
    salvar_curso_json(caminho_json, cfg)
    return cfg


def remover_justificativa(
    diretorio_curso: Path,
    nome: str,
    email: str,
    arquivo_encontro: str,
) -> CursoConfig:
    """Remove uma justificativa registrada em curso.json."""
    caminho_json = obter_caminho_curso_json(diretorio_curso)
    cfg = carregar_curso_json(caminho_json)
    chave_alvo = _chave_participante_justificativa(nome, email, arquivo_encontro)

    cfg.justificativas = [
        j for j in cfg.justificativas
        if _chave_participante_justificativa(j.nome, j.email, j.arquivo_encontro) != chave_alvo
    ]
    salvar_curso_json(caminho_json, cfg)
    return cfg
