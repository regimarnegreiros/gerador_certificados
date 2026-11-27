"""
Módulo para envio de e-mails com certificados via Microsoft Outlook.
"""

from datetime import datetime
from pathlib import Path
import time
from typing import Any, Callable, Dict, Optional
import pandas as pd
import win32com.client as win32

from core.caminhos import (
    LIMITE_CARACTERES_CAMINHO_FINAL,
    caminho_estendido,
    existe_com_fallback,
    garantir_diretorio,
    obter_mensagem_orientacao_caminho_profundo,
)
from core.certificados import nome_seguro_arquivo
from core.eventos import cancelado, emitir_log, emitir_progresso


def conectar_outlook(on_log: Optional[Callable[[str, str], None]] = None):
    """Conecta à instância do Microsoft Outlook aberta ou inicia uma nova."""
    try:
        outlook = win32.GetActiveObject("Outlook.Application")
        emitir_log(on_log, "sucesso", "Conectado ao Microsoft Outlook em execução.")
        return outlook
    except Exception:
        try:
            outlook = win32.Dispatch("Outlook.Application")
            emitir_log(on_log, "sucesso", "Iniciado novo processo do Microsoft Outlook.")
            return outlook
        except Exception as e:
            raise RuntimeError(
                f"Não foi possível conectar ao Microsoft Outlook: {e}\n"
                "Verifique se o aplicativo Outlook está instalado e configurado no computador."
            )


def criar_corpo_email(nome: str, template_corpo: str, nome_curso: str) -> str:
    """Preenche marcadores {nome} e {CURSO} no texto do e-mail."""
    linhas = template_corpo.splitlines()
    while linhas and linhas[0].strip().startswith("#"):
        linhas.pop(0)
    texto = "\n".join(linhas)
    return texto.replace("{nome}", nome).replace("{CURSO}", nome_curso)


def criar_mensagem_outlook(
    outlook,
    nome: str,
    email: str,
    caminho_pdf: Path,
    assunto_template: str,
    corpo_template: str,
    nome_curso: str,
):
    """Cria e configura o item de e-mail no Outlook."""
    assunto = assunto_template.replace("{CURSO}", nome_curso).replace("{nome}", nome)
    corpo = criar_corpo_email(nome=nome, template_corpo=corpo_template, nome_curso=nome_curso)

    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.To = email
    mail.Subject = assunto
    mail.Body = corpo

    caminho_pdf_resolvido = Path(caminho_pdf).resolve()
    if not existe_com_fallback(caminho_pdf_resolvido):
        raise FileNotFoundError(f"Arquivo PDF do certificado não encontrado para anexar: {caminho_pdf_resolvido}")

    caminho_str = str(caminho_pdf_resolvido)
    if len(caminho_str) > LIMITE_CARACTERES_CAMINHO_FINAL:
        raise ValueError(
            obter_mensagem_orientacao_caminho_profundo(
                caminho_str,
                detalhe=f"Não foi possível anexar o certificado de '{nome}' pois o caminho ultrapassa {LIMITE_CARACTERES_CAMINHO_FINAL} caracteres.",
                limite=LIMITE_CARACTERES_CAMINHO_FINAL,
            )
        )

    try:
        mail.Attachments.Add(caminho_str)
    except Exception as e:
        raise RuntimeError(
            f"Erro ao anexar arquivo no Microsoft Outlook ({caminho_pdf_resolvido.name}): {e}\n"
            "Verifique se o arquivo não está bloqueado e se a pasta do curso não está muito profunda."
        )

    return mail


def localizar_certificado_pdf(pasta_pdf: Path, nome: str) -> Path:
    """Retorna o caminho esperado do certificado PDF do participante no padrão 'CERT_NOME.pdf'."""
    nome_limpo = nome_seguro_arquivo(nome)
    return Path(pasta_pdf) / f"CERT_{nome_limpo}.pdf"


def enviar_email_teste(
    caminho_pdf: Path,
    nome: str,
    email_destinatario: str,
    config: Dict[str, Any],
    abrir_no_outlook: bool = True,
    on_log: Optional[Callable[[str, str], None]] = None,
):
    """Gera um único e-mail de teste para validação do usuário."""
    outlook = conectar_outlook(on_log=on_log)
    nome_curso = config.get("NOME_DO_CURSO", "")
    assunto = config.get("ASSUNTO_EMAIL", f"Certificado - Curso {nome_curso}")
    corpo = config.get("CORPO_EMAIL", "")

    if not email_destinatario or not email_destinatario.strip():
        raise ValueError(f"O participante '{nome}' não possui e-mail cadastrado para teste.")

    msg = criar_mensagem_outlook(
        outlook=outlook,
        nome=nome,
        email=email_destinatario,
        caminho_pdf=caminho_pdf,
        assunto_template=assunto,
        corpo_template=corpo,
        nome_curso=nome_curso,
    )

    if abrir_no_outlook:
        emitir_log(on_log, "info", f"Abrindo rascunho de conferência no Outlook para {nome} ({email_destinatario})...")
        msg.Display()
        emitir_log(on_log, "sucesso", "Mensagem exibida na tela do Outlook para conferência.")
    else:
        msg.Send()
        emitir_log(on_log, "sucesso", f"E-mail de teste enviado para: {email_destinatario}")


def enviar_todos_emails(
    df_aptos: pd.DataFrame,
    pasta_pdf: Path,
    pasta_logs: Path,
    config: Dict[str, Any],
    on_log: Optional[Callable[[str, str], None]] = None,
    on_progress: Optional[Callable[[int, int, str], None]] = None,
    on_item_status: Optional[Callable[[Dict[str, Any]], None]] = None,
    cancel_event: Optional[Any] = None,
    intervalo_segundos: float = 1.0,
) -> pd.DataFrame:
    """
    Dispara os e-mails para todos os participantes aptos que possuem certificado PDF.
    Suporta callbacks de progresso, status por item e cancelamento cooperativo.
    """
    outlook = conectar_outlook(on_log=on_log)
    nome_curso = config.get("NOME_DO_CURSO", "")
    assunto = config.get("ASSUNTO_EMAIL", f"Certificado - Curso {nome_curso}")
    corpo = config.get("CORPO_EMAIL", "")

    total = len(df_aptos)
    resultados = []

    emitir_log(on_log, "info", f"Iniciando envio de e-mails via Outlook ({total} participante(s) apto(s))...")

    for i, (_, linha) in enumerate(df_aptos.iterrows(), start=1):
        if cancelado(cancel_event):
            emitir_log(on_log, "aviso", "Envio de e-mails cancelado pelo usuário.")
            break

        nome = str(linha["Nome"]).strip()
        email = str(linha.get("Email", "")).strip() if pd.notna(linha.get("Email")) else ""
        caminho_pdf = localizar_certificado_pdf(pasta_pdf, nome)

        emitir_progresso(on_progress, i, total, f"Enviando ({i}/{total}): {nome}")
        emitir_log(on_log, "info", f"[{i}/{total}] Destinatário: {nome} ({email or 'sem e-mail'})")

        if not email:
            registro = {
                "Nome": nome,
                "Email": "",
                "PDF": str(caminho_pdf) if existe_com_fallback(caminho_pdf) else "",
                "Status": "SEM_EMAIL",
                "Data": datetime.now(),
                "Erro": "Participante sem e-mail cadastrado",
            }
            resultados.append(registro)
            if on_item_status is not None:
                on_item_status(registro)
            emitir_log(on_log, "erro", f"{nome}: participante sem e-mail cadastrado.")
            continue

        if not existe_com_fallback(caminho_pdf):
            registro = {
                "Nome": nome,
                "Email": email,
                "PDF": "",
                "Status": "SEM_PDF",
                "Data": datetime.now(),
                "Erro": f"PDF não encontrado: {caminho_pdf.name}",
            }
            resultados.append(registro)
            if on_item_status is not None:
                on_item_status(registro)
            emitir_log(on_log, "erro", f"{nome}: certificado PDF não encontrado ({caminho_pdf.name}).")
            continue

        try:
            msg = criar_mensagem_outlook(
                outlook=outlook,
                nome=nome,
                email=email,
                caminho_pdf=caminho_pdf,
                assunto_template=assunto,
                corpo_template=corpo,
                nome_curso=nome_curso,
            )
            msg.Send()
            registro = {
                "Nome": nome,
                "Email": email,
                "PDF": str(caminho_pdf),
                "Status": "ENVIADO",
                "Data": datetime.now(),
                "Erro": "",
            }
            resultados.append(registro)
            if on_item_status is not None:
                on_item_status(registro)
            emitir_log(on_log, "sucesso", f"E-mail enviado com sucesso para {nome} ({email}).")

            if intervalo_segundos > 0:
                if cancel_event is not None and callable(getattr(cancel_event, "wait", None)):
                    if cancel_event.wait(intervalo_segundos):
                        emitir_log(on_log, "aviso", "Envio de e-mails cancelado pelo usuário.")
                        break
                else:
                    time.sleep(intervalo_segundos)

        except Exception as erro:
            registro = {
                "Nome": nome,
                "Email": email,
                "PDF": str(caminho_pdf),
                "Status": "ERRO",
                "Data": datetime.now(),
                "Erro": str(erro),
            }
            resultados.append(registro)
            if on_item_status is not None:
                on_item_status(registro)
            emitir_log(on_log, "erro", f"Falha no envio para {nome}: {erro}")

    df_envio = pd.DataFrame(resultados)

    garantir_diretorio(pasta_logs)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arq_log = pasta_logs / f"log_envio_{timestamp}.xlsx"
    try:
        df_envio.to_excel(arq_log, index=False)
    except (FileNotFoundError, OSError):
        df_envio.to_excel(caminho_estendido(arq_log), index=False)
    emitir_log(on_log, "info", f"Log de auditoria salvo em: {arq_log.name}")

    return df_envio
