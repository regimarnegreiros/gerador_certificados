"""
Módulo para geração de certificados individuais em PPT/PDF via Microsoft PowerPoint
e consolidação em um único PDF (junção de certificados).
Extraído e adaptado diretamente de certificados.ipynb.
"""

from pathlib import Path
from datetime import datetime
import re
import os
import shutil
import pandas as pd
from typing import Dict, Any, List, Tuple, Optional
from pypdf import PdfWriter
import win32com.client as win32
from core.caminhos import (
    caminho_estendido,
    garantir_diretorio,
    abrir_arquivo,
    copiar_arquivo_com_fallback,
    existe_com_fallback,
    listar_diretorio_com_fallback,
    ler_excel_com_fallback,
    sanitizar_nome_arquivo,
    encurtar_caminho_se_necessario,
    LIMITE_CARACTERES_CAMINHO_FINAL,
    obter_mensagem_orientacao_caminho_profundo,
    esvaziar_pasta,
)
from core.ui import verde, vermelho, ciano, negrito, dim, seguro_print


MESES = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro",
}


def converter_data(data_val: Any) -> Optional[datetime]:
    """Converte string DD/MM/AAAA ou Timestamp em objeto datetime."""
    if pd.isna(data_val) or not data_val:
        return None
    if isinstance(data_val, (datetime, pd.Timestamp)):
        return pd.to_datetime(data_val).to_pydatetime()
    texto = str(data_val).strip()
    try:
        return datetime.strptime(texto, "%d/%m/%Y")
    except ValueError:
        try:
            return pd.to_datetime(texto).to_pydatetime()
        except Exception:
            return None


def data_por_extenso(dt: datetime, mes_maiusculo: bool = True) -> str:
    """Formata data no padrão oficial: '14 de setembro de 2026'."""
    mes = MESES.get(dt.month, "")
    if mes_maiusculo and mes:
        mes = mes.capitalize()
    return f"{dt.day} de {mes} de {dt.year}"


def periodo_por_extenso(data_inicio_str: str, data_fim_str: str) -> Tuple[str, str]:
    """
    Retorna ('no dia', '14 de setembro de 2026') para 1 dia
    ou ('no período de', '14 a 18 de setembro de 2026') para múltiplos dias.
    """
    inicio = converter_data(data_inicio_str)
    fim = converter_data(data_fim_str) if data_fim_str else None

    if inicio is None and fim is None:
        return "no período de", ""

    if fim is None or (inicio and fim and inicio.date() == fim.date()):
        dt = inicio if inicio else fim
        return "dia", f"{dt.day} de {MESES[dt.month]} de {dt.year}"

    if inicio and fim and inicio.year == fim.year and inicio.month == fim.month:
        periodo = f"{inicio.day} a {fim.day} de {MESES[inicio.month]} de {inicio.year}"
        return "período", periodo

    if inicio and fim and inicio.year == fim.year:
        periodo = f"{inicio.day} de {MESES[inicio.month]} a {fim.day} de {MESES[fim.month]} de {fim.year}"
        return "período", periodo

    periodo = f"{inicio.day} de {MESES[inicio.month]} de {inicio.year} a {fim.day} de {MESES[fim.month]} de {fim.year}"
    return "período", periodo


def nome_seguro_arquivo(nome: str) -> str:
    """Higieniza o nome para ser usado como nome de arquivo válido no Windows."""
    nome = re.sub(r'[\\/*?:"<>|\x00-\x1f\x7f]', "", str(nome))
    nome = re.sub(r"\s+", "_", nome.strip(" ."))
    return nome.upper()


def carregar_relatorio_presenca(caminho_relatorio: Path) -> pd.DataFrame:
    """Lê a aba 'Certificação' (ou 'Relatório Geral') da planilha de presenças."""
    if not existe_com_fallback(caminho_relatorio):
        raise FileNotFoundError(f"Arquivo de relatório de presença não encontrado: {caminho_relatorio}")

    try:
        try:
            df = ler_excel_com_fallback(caminho_relatorio, sheet_name="Certificação")
        except (ValueError, KeyError):
            try:
                df = ler_excel_com_fallback(caminho_relatorio, sheet_name="Relatório Geral")
            except (ValueError, KeyError):
                # Se não houver a aba nomeada, lê a primeira aba
                df = ler_excel_com_fallback(caminho_relatorio)
    except PermissionError:
        raise PermissionError(
            f"Permissão negada ao abrir '{Path(caminho_relatorio).name}'. "
            f"O arquivo pode estar aberto no Microsoft Excel. Feche o Excel e tente novamente."
        )

    colunas_necessarias = ["Nome", "Apto por CH"]
    faltantes = [c for c in colunas_necessarias if c not in df.columns]
    if faltantes:
        raise ValueError(f"O relatório não possui as colunas necessárias: {', '.join(faltantes)}")

    return df


def filtrar_aptos(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Separa participantes aptos e não aptos com base na coluna 'Apto por CH'."""
    condicao_apto = df["Apto por CH"].astype(str).str.strip().str.upper().eq("APTO")
    aptos = df[condicao_apto].copy()
    nao_aptos = df[~condicao_apto].copy()
    return aptos, nao_aptos


def criar_substituicoes(linha: pd.Series, config: Dict[str, Any]) -> Dict[str, str]:
    """Cria dicionário de marcadores do PowerPoint para o participante."""
    nome = str(linha["Nome"]).strip()
    dia_ou_periodo, periodo = periodo_por_extenso(
        config.get("DATA_INICIO", ""),
        config.get("DATA_FIM", "")
    )

    data_emissao_dt = converter_data(config.get("DATA_EMISSAO"))
    data_emissao_ext = data_por_extenso(data_emissao_dt, mes_maiusculo=True) if data_emissao_dt else ""

    return {
        "<<NOME>>": nome,
        "<<CURSO>>": str(config.get("NOME_DO_CURSO", "")).strip(),
        "<<DIA_OU_PERIODO>>": dia_ou_periodo,
        "<<PERIODO>>": periodo,
        "<<CARGA_HORARIA>>": str(config.get("CARGA_HORARIA", "")).strip(),
        "<<LOCAL_EMISSAO>>": str(config.get("LOCAL_EMISSAO", "")).strip(),
        "<<DATA_EMISSAO>>": data_emissao_ext,
    }


def substituir_texto_preservando_formatacao(text_range, substituicoes: Dict[str, str]):
    """
    Substitui marcadores caractere a caractere preservando fontes,
    tamanhos e estilos na caixa de texto do PowerPoint.
    """
    for marcador, valor in substituicoes.items():
        valor = str(valor)
        inicio_busca = 1
        while True:
            texto_atual = text_range.Text
            posicao = texto_atual.find(marcador, inicio_busca - 1)
            if posicao == -1:
                break

            inicio = posicao + 1
            tamanho = len(marcador)
            caracteres = text_range.Characters(inicio, tamanho)
            caracteres.Text = valor
            inicio_busca = inicio + len(valor)

            if valor == marcador:
                break


def substituir_textos_slide(slide, substituicoes: Dict[str, str]):
    """Varre caixas de texto, tabelas e formas agrupadas de um slide específico substituindo os marcadores."""
    for shape in slide.Shapes:
        try:
            if shape.HasTextFrame:
                if shape.TextFrame.HasText:
                    substituir_texto_preservando_formatacao(shape.TextFrame.TextRange, substituicoes)

            if shape.HasTable:
                tabela = shape.Table
                for linha in range(1, tabela.Rows.Count + 1):
                    for coluna in range(1, tabela.Columns.Count + 1):
                        celula = tabela.Cell(linha, coluna)
                        substituir_texto_preservando_formatacao(celula.Shape.TextFrame.TextRange, substituicoes)

            # Tratamento para formas agrupadas (msoGroup = 6)
            if getattr(shape, "Type", None) == 6:
                for item in shape.GroupItems:
                    if item.HasTextFrame and item.TextFrame.HasText:
                        substituir_texto_preservando_formatacao(item.TextFrame.TextRange, substituicoes)
        except Exception:
            pass


def substituir_textos_powerpoint(apresentacao, substituicoes: Dict[str, str]):
    """Varre todos os slides da apresentação substituindo os marcadores."""
    for slide in apresentacao.Slides:
        substituir_textos_slide(slide, substituicoes)


def gerar_pdf_pessoa(
    linha: pd.Series,
    modelo_powerpoint: Path,
    pasta_ppt: Optional[Path],
    pasta_pdf: Path,
    config: Dict[str, Any],
    pasta_staging: Optional[Path] = None,
    modelo_staging: Optional[Path] = None,
    powerpoint: Optional[Any] = None,
    salvar_ppt: bool = False,
) -> Path:
    """
    Abre o modelo no PowerPoint via COM, substitui tags e exporta para PDF.
    Gera diretamente nas pastas finais (sem arquivos temporários em %TEMP%),
    evitando bloqueios comportamentais de antivírus/EDR como Trend Micro Vision One.
    Reutiliza a instância ativa do PowerPoint quando fornecida, garantindo estabilidade total.
    Por padrão não salva PPT individual (salvar_ppt=False), poupando I/O em disco.
    """
    nome = str(linha["Nome"]).strip()
    nome_arq = nome_seguro_arquivo(nome)

    extensao_origem = modelo_powerpoint.suffix.lower()
    caminho_pdf_final = pasta_pdf / f"CERT_{nome_arq}.pdf"
    garantir_diretorio(pasta_pdf)
    str_pdf_final = str(caminho_pdf_final.resolve())

    caminho_ppt_final = None
    str_ppt_final = ""
    formato_ppt = 25 if extensao_origem == ".pptm" else 24
    if salvar_ppt and pasta_ppt:
        caminho_ppt_final = pasta_ppt / f"CERT_{nome_arq}{extensao_origem}"
        garantir_diretorio(pasta_ppt)
        str_ppt_final = str(caminho_ppt_final.resolve())

    # Se o caminho for excessivamente longo para o PowerPoint COM, orienta o usuário
    caminho_checar = str_ppt_final if (salvar_ppt and len(str_ppt_final) > LIMITE_CARACTERES_CAMINHO_FINAL) else str_pdf_final
    if len(caminho_checar) > LIMITE_CARACTERES_CAMINHO_FINAL:
        msg_erro = obter_mensagem_orientacao_caminho_profundo(
            caminho_checar,
            detalhe=f"O caminho do certificado para '{nome}' ultrapassa {LIMITE_CARACTERES_CAMINHO_FINAL} caracteres.",
            limite=LIMITE_CARACTERES_CAMINHO_FINAL,
        )
        raise ValueError(msg_erro)

    ppt_instancia = powerpoint
    powerpoint_criado_localmente = False
    apresentacao = None

    try:
        if ppt_instancia is None:
            ppt_instancia = win32.Dispatch("PowerPoint.Application")
            try:
                ppt_instancia.DisplayAlerts = 1  # ppAlertsNone (bloqueia caixas de diálogo modais)
            except Exception:
                pass
            powerpoint_criado_localmente = True

        # Abre o modelo diretamente como cópia independente (Untitled=True, ReadOnly=True) e sem janela (WithWindow=False)
        caminho_modelo_abrir = str(modelo_powerpoint.resolve())
        apresentacao = ppt_instancia.Presentations.Open(
            caminho_modelo_abrir,
            ReadOnly=True,
            Untitled=True,
            WithWindow=False,
        )

        substituicoes = criar_substituicoes(linha, config)
        substituir_textos_powerpoint(apresentacao, substituicoes)

        # Salva PPT individual apenas se explicitamente solicitado
        if salvar_ppt and caminho_ppt_final:
            apresentacao.SaveAs(str_ppt_final, formato_ppt)

        # Salva diretamente em PDF
        apresentacao.SaveAs(str_pdf_final, 32)

        try:
            apresentacao.Close()
        except Exception:
            pass
        apresentacao = None

        if not existe_com_fallback(caminho_pdf_final):
            raise FileNotFoundError(f"Falha ao gerar arquivo PDF para {nome}.")

        return caminho_pdf_final

    finally:
        if apresentacao is not None:
            try:
                apresentacao.Close()
            except Exception:
                pass
            apresentacao = None

        if powerpoint_criado_localmente and ppt_instancia is not None:
            try:
                ppt_instancia.Quit()
            except Exception:
                pass
            del ppt_instancia
            import gc
            gc.collect()


def gerar_powerpoint_consolidado(
    df_aptos: pd.DataFrame,
    modelo_powerpoint: Path,
    caminho_ppt_junto: Path,
    config: Dict[str, Any],
    powerpoint: Optional[Any] = None,
) -> Path:
    """
    Cria uma apresentação única no Microsoft PowerPoint contendo todos os certificados
    dos participantes aptos em sequência direta.
    Suporta modelos com 1 ou múltiplos slides (ex: frente e verso), duplicando
    os slides de forma 100% nativa sem uso da área de transferência (clipboard).
    """
    if df_aptos.empty:
        raise ValueError("Nenhum participante apto para gerar a apresentação consolidada.")

    caminho_ppt_junto = Path(caminho_ppt_junto)
    garantir_diretorio(caminho_ppt_junto.parent)

    extensao_origem = modelo_powerpoint.suffix.lower()
    formato_ppt = 25 if extensao_origem == ".pptm" else 24

    str_ppt_junto = str(caminho_ppt_junto.resolve())
    if len(str_ppt_junto) > LIMITE_CARACTERES_CAMINHO_FINAL:
        msg_erro = obter_mensagem_orientacao_caminho_profundo(
            str_ppt_junto,
            detalhe=f"O caminho do PowerPoint consolidado ultrapassa {LIMITE_CARACTERES_CAMINHO_FINAL} caracteres.",
            limite=LIMITE_CARACTERES_CAMINHO_FINAL,
        )
        raise ValueError(msg_erro)

    ppt_instancia = powerpoint
    powerpoint_criado_localmente = False
    apresentacao = None

    try:
        if ppt_instancia is None:
            ppt_instancia = win32.Dispatch("PowerPoint.Application")
            try:
                ppt_instancia.DisplayAlerts = 1
            except Exception:
                pass
            powerpoint_criado_localmente = True

        caminho_modelo_abrir = str(modelo_powerpoint.resolve())
        apresentacao = ppt_instancia.Presentations.Open(
            caminho_modelo_abrir,
            ReadOnly=True,
            Untitled=True,
            WithWindow=False,
        )

        num_slides_modelo = apresentacao.Slides.Count

        for _, linha in df_aptos.iterrows():
            substituicoes = criar_substituicoes(linha, config)
            for k in range(1, num_slides_modelo + 1):
                orig_slide = apresentacao.Slides(k)
                dup = orig_slide.Duplicate()
                novo_slide = dup.Item(1)
                novo_slide.MoveTo(apresentacao.Slides.Count)
                substituir_textos_slide(novo_slide, substituicoes)

        # Remove os slides originais de modelo
        for _ in range(num_slides_modelo):
            apresentacao.Slides(1).Delete()

        apresentacao.SaveAs(str_ppt_junto, formato_ppt)
        try:
            apresentacao.Close()
        except Exception:
            pass
        apresentacao = None

        if not existe_com_fallback(caminho_ppt_junto):
            raise FileNotFoundError(f"Falha ao gerar PowerPoint consolidado: {caminho_ppt_junto}")

        return caminho_ppt_junto

    finally:
        if apresentacao is not None:
            try:
                apresentacao.Close()
            except Exception:
                pass
            apresentacao = None

        if powerpoint_criado_localmente and ppt_instancia is not None:
            try:
                ppt_instancia.Quit()
            except Exception:
                pass
            del ppt_instancia
            import gc
            gc.collect()


def gerar_todos_certificados(
    df_aptos: pd.DataFrame,
    modelo_powerpoint: Path,
    pasta_ppt: Optional[Path],
    pasta_pdf: Path,
    pasta_logs: Path,
    config: Dict[str, Any],
    caminho_ppt_junto: Optional[Path] = None,
    salvar_ppt_individual: bool = False,
) -> pd.DataFrame:
    """
    Gera certificados para todos os participantes aptos diretamente para as pastas de destino.
    Gera PDFs individuais e, opcionalmente, um arquivo PowerPoint único com todos os certificados.
    Por padrão não cria arquivos PPT individuais.
    """
    if salvar_ppt_individual and pasta_ppt:
        garantir_diretorio(pasta_ppt)
    garantir_diretorio(pasta_pdf)
    garantir_diretorio(pasta_logs)

    # Limpeza preventiva: remove certificados anteriores para garantir que
    # participantes removidos do relatório não permaneçam nas pastas de saída
    removidos_pdf = esvaziar_pasta(pasta_pdf, (".pdf",))
    removidos_ppt = 0
    if pasta_ppt and existe_com_fallback(pasta_ppt):
        removidos_ppt = esvaziar_pasta(pasta_ppt, (".pptx", ".pptm"))
    if removidos_pdf > 0 or removidos_ppt > 0:
        seguro_print(f"  {dim(f'-> Limpeza prévia: {removidos_pdf} PDF(s) e {removidos_ppt} PPT(s) de execuções anteriores foram removidos.')}")

    total = len(df_aptos)
    resultados = []

    # Validação antecipada de caminhos para evitar falhas no PowerPoint
    caminhos_excessivos = []
    ext_origem = modelo_powerpoint.suffix.lower()
    for _, linha in df_aptos.iterrows():
        nome = str(linha["Nome"]).strip()
        nome_arq = nome_seguro_arquivo(nome)
        c_pdf = pasta_pdf / f"CERT_{nome_arq}.pdf"
        str_pdf = str(c_pdf.resolve())
        if len(str_pdf) > LIMITE_CARACTERES_CAMINHO_FINAL:
            caminhos_excessivos.append(str_pdf)
        if salvar_ppt_individual and pasta_ppt:
            c_ppt = pasta_ppt / f"CERT_{nome_arq}{ext_origem}"
            str_ppt = str(c_ppt.resolve())
            if len(str_ppt) > LIMITE_CARACTERES_CAMINHO_FINAL:
                caminhos_excessivos.append(str_ppt)

    if caminho_ppt_junto:
        str_ppt_junto = str(Path(caminho_ppt_junto).resolve())
        if len(str_ppt_junto) > LIMITE_CARACTERES_CAMINHO_FINAL:
            caminhos_excessivos.append(str_ppt_junto)

    if caminhos_excessivos:
        msg_orientacao = obter_mensagem_orientacao_caminho_profundo(
            caminhos_excessivos[0],
            detalhe=f"{len(caminhos_excessivos)} certificado(s) ultrapassam o limite seguro de {LIMITE_CARACTERES_CAMINHO_FINAL} caracteres.",
            limite=LIMITE_CARACTERES_CAMINHO_FINAL,
        )
        raise ValueError(msg_orientacao)

    seguro_print(f"\n  {ciano(negrito(f'GERAÇÃO DE CERTIFICADOS ({total} participantes aptos)'))}")
    seguro_print(f"  {dim('-' * 60)}")

    # Inicia o PowerPoint UMA ÚNICA VEZ antes do loop para garantir velocidade e estabilidade total
    powerpoint = None
    try:
        powerpoint = win32.Dispatch("PowerPoint.Application")
        try:
            powerpoint.DisplayAlerts = 1  # ppAlertsNone (impede travamentos com janelas modais)
        except Exception:
            pass
    except Exception as e:
        raise RuntimeError(f"Não foi possível inicializar o Microsoft PowerPoint: {e}")

    try:
        for i, (_, linha) in enumerate(df_aptos.iterrows(), start=1):
            nome = str(linha["Nome"]).strip()
            email = str(linha.get("Email", "")).strip() if pd.notna(linha.get("Email")) else ""
            seguro_print(f"  {ciano(f'[{i}/{total}]')} Gerando certificado para: {negrito(nome)}")

            try:
                caminho_pdf = gerar_pdf_pessoa(
                    linha=linha,
                    modelo_powerpoint=modelo_powerpoint,
                    pasta_ppt=pasta_ppt,
                    pasta_pdf=pasta_pdf,
                    config=config,
                    powerpoint=powerpoint,
                    salvar_ppt=salvar_ppt_individual,
                )
                resultados.append({
                    "Nome": nome,
                    "Email": email,
                    "PDF": str(caminho_pdf),
                    "Status": "OK",
                    "Erro": "",
                })
                seguro_print(f"     -> {verde('[OK] Sucesso:')} {caminho_pdf.name}")
            except Exception as e:
                seguro_print(f"     -> {vermelho('[ERRO] Falha:')} {e}")
                resultados.append({
                    "Nome": nome,
                    "Email": email,
                    "PDF": "",
                    "Status": "ERRO",
                    "Erro": str(e),
                })

        # Geração do PowerPoint único consolidado com todos os participantes
        if caminho_ppt_junto and len(df_aptos) > 0:
            seguro_print(f"\n  {ciano(negrito('GERAÇÃO DE POWERPOINT CONSOLIDADO (TODOS OS CERTIFICADOS)'))}")
            seguro_print(f"  {dim('-' * 60)}")
            seguro_print(f"  {ciano('->')} Consolidando todos os certificados em um único arquivo PowerPoint...")
            try:
                gerar_powerpoint_consolidado(
                    df_aptos=df_aptos,
                    modelo_powerpoint=modelo_powerpoint,
                    caminho_ppt_junto=caminho_ppt_junto,
                    config=config,
                    powerpoint=powerpoint,
                )
                seguro_print(f"     -> {verde('[OK] Sucesso:')} {Path(caminho_ppt_junto).name}")
            except Exception as e:
                seguro_print(f"     -> {vermelho('[ERRO] Falha ao gerar PPT consolidado:')} {e}")
                raise

    finally:
        if powerpoint is not None:
            try:
                powerpoint.Quit()
            except Exception:
                pass
            del powerpoint
            import gc
            gc.collect()

    df_resultados = pd.DataFrame(resultados)

    # Salva log de geração com fallback para caminho estendido
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arq_log = pasta_logs / f"log_certificados_{timestamp}.xlsx"
    try:
        df_resultados.to_excel(arq_log, index=False)
    except (FileNotFoundError, OSError):
        df_resultados.to_excel(caminho_estendido(arq_log), index=False)
    seguro_print(f"\n  {verde('[OK]')} {dim(f'Log de auditoria salvo em: {arq_log.name}')}")

    return df_resultados


def juntar_certificados_pdf(pasta_pdf: Path, caminho_saida_junto: Path) -> Path:
    """
    Junta todos os certificados PDF encontrados em pasta_pdf em um único arquivo PDF.
    Pode ser executado separadamente caso o usuário adicione certificados manuais depois.
    """
    caminho_saida_junto = Path(caminho_saida_junto)
    nome_seguro = sanitizar_nome_arquivo(caminho_saida_junto.stem) + caminho_saida_junto.suffix
    caminho_saida_junto = caminho_saida_junto.parent / nome_seguro
    garantir_diretorio(caminho_saida_junto.parent)

    arquivos_pdf = sorted([
        f for f in listar_diretorio_com_fallback(pasta_pdf)
        if f.is_file() and f.suffix.lower() == ".pdf" and not f.name.startswith("~$")
    ])

    if not arquivos_pdf:
        raise FileNotFoundError(f"Nenhum arquivo PDF encontrado na pasta '{pasta_pdf}'.")

    seguro_print(f"\n  {ciano(negrito(f'JUNÇÃO DE CERTIFICADOS EM PDF ÚNICO ({len(arquivos_pdf)} arquivos)'))}")
    seguro_print(f"  {dim('-' * 60)}")

    writer = PdfWriter()
    for arq in arquivos_pdf:
        seguro_print(f"     + {dim('Incluindo:')} {arq.name}")
        with abrir_arquivo(arq, "rb") as f_in:
            writer.append(f_in)

    with abrir_arquivo(caminho_saida_junto, "wb") as f_out:
        writer.write(f_out)

    seguro_print(f"\n  {verde('[OK] PDF consolidado gerado com sucesso!')}")
    seguro_print(f"     {dim(f'Arquivo: {caminho_saida_junto}')}")
    return caminho_saida_junto
