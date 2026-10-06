"""
Módulo para envio de e-mails com certificados via Microsoft Outlook.
Extraído e adaptado diretamente de certificados.ipynb.
"""

from pathlib import Path
from datetime import datetime
import time
import os
import pandas as pd
from typing import Dict, Any, List, Optional
import win32com.client as win32
from core.caminhos import (
    caminho_estendido,
    garantir_diretorio,
    copiar_arquivo_com_fallback,
    existe_com_fallback,
    LIMITE_CARACTERES_CAMINHO_FINAL,
    obter_mensagem_orientacao_caminho_profundo,
)
from core.certificados import nome_seguro_arquivo
from core.ui import verde, vermelho, amarelo, ciano, negrito, dim, seguro_print


def conectar_outlook():
    """Conecta à instância do Microsoft Outlook aberta ou inicia uma nova."""
    try:
        outlook = win32.GetActiveObject("Outlook.Application")
        seguro_print(f"  {verde('[OK] Conectado ao Microsoft Outlook em execução.')}")
        return outlook
    except Exception:
        try:
            outlook = win32.Dispatch("Outlook.Application")
            seguro_print(f"  {verde('[OK] Iniciado novo processo do Microsoft Outlook.')}")
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
    texto = texto.replace("{nome}", nome).replace("{CURSO}", nome_curso)
    return texto


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

    # Anexa o certificado PDF diretamente pelo seu caminho (sem uso de arquivos temporários do sistema)
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


def enviar_email_teste(
    caminho_pdf: Path,
    nome: str,
    email_destinatario: str,
    config: Dict[str, Any],
    abrir_no_outlook: bool = True,
):
    """Gera um único e-mail de teste para validação do usuário."""
    outlook = conectar_outlook()
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
        seguro_print(f"\n  {dim('-> Abrindo rascunho de conferência no Outlook para')} {negrito(nome)} {dim(f'({email_destinatario})...')}")
        msg.Display()
        seguro_print(f"  {verde('[OK] Mensagem exibida na tela do Outlook!')} {dim('Verifique o anexo, assunto e texto.')}")
    else:
        msg.Send()
        seguro_print(f"  {verde('[OK] E-mail de teste enviado para:')} {email_destinatario}")


def localizar_certificado_pdf(pasta_pdf: Path, nome: str) -> Path:
    """
    Localiza o arquivo PDF do participante na pasta informada.
    Prioriza o novo padrão compacto ('CERT_NOME.pdf') e mantém fallback
    automático para o padrão legado ('CERTIFICADO_NOME.pdf').
    """
    pasta = Path(pasta_pdf)
    nome_limpo = nome_seguro_arquivo(nome)
    candidato_novo = pasta / f"CERT_{nome_limpo}.pdf"
    if existe_com_fallback(candidato_novo):
        return candidato_novo
    candidato_legado = pasta / f"CERTIFICADO_{nome_limpo}.pdf"
    if existe_com_fallback(candidato_legado):
        return candidato_legado
    return candidato_novo


def enviar_todos_emails(
    df_aptos: pd.DataFrame,
    pasta_pdf: Path,
    pasta_logs: Path,
    config: Dict[str, Any],
) -> pd.DataFrame:
    """Dispara os e-mails para todos os participantes aptos que possuem certificado PDF."""
    outlook = conectar_outlook()
    nome_curso = config.get("NOME_DO_CURSO", "")
    assunto = config.get("ASSUNTO_EMAIL", f"Certificado - Curso {nome_curso}")
    corpo = config.get("CORPO_EMAIL", "")

    total = len(df_aptos)
    resultados = []

    seguro_print(f"\n  {ciano(negrito(f'ENVIO DE E-MAILS VIA OUTLOOK ({total} participantes aptos)'))}")
    seguro_print(f"  {dim('-' * 60)}")

    for i, (_, linha) in enumerate(df_aptos.iterrows(), start=1):
        nome = str(linha["Nome"]).strip()
        email = str(linha.get("Email", "")).strip() if pd.notna(linha.get("Email")) else ""

        # Localiza o arquivo PDF do participante (suporta CERT_ e CERTIFICADO_)
        caminho_pdf = localizar_certificado_pdf(pasta_pdf, nome)

        seguro_print(f"\n  {ciano(f'[{i}/{total}]')} Destinatário: {negrito(nome)}")
        seguro_print(f"     • E-mail: {email if email else amarelo('(não informado)')}")

        if not email:
            seguro_print(f"     -> {vermelho('[ERRO] Participante sem e-mail cadastrado.')}")
            resultados.append({
                "Nome": nome,
                "Email": "",
                "PDF": str(caminho_pdf) if existe_com_fallback(caminho_pdf) else "",
                "Status": "SEM_EMAIL",
                "Data": datetime.now(),
                "Erro": "Participante sem e-mail cadastrado",
            })
            continue

        if not existe_com_fallback(caminho_pdf):
            seguro_print(f"     -> {vermelho(f'[ERRO] Certificado PDF não encontrado ({caminho_pdf.name}).')}")
            resultados.append({
                "Nome": nome,
                "Email": email,
                "PDF": "",
                "Status": "SEM_PDF",
                "Data": datetime.now(),
                "Erro": f"PDF não encontrado: {caminho_pdf.name}",
            })
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
            seguro_print(f"     -> {verde('[OK] E-mail enviado com sucesso!')}")

            resultados.append({
                "Nome": nome,
                "Email": email,
                "PDF": str(caminho_pdf),
                "Status": "ENVIADO",
                "Data": datetime.now(),
                "Erro": "",
            })

            # Pausa de segurança de 1 segundo entre envios
            time.sleep(1)

        except Exception as erro:
            seguro_print(f"     -> {vermelho(f'[ERRO] Falha no envio: {erro}')}")
            resultados.append({
                "Nome": nome,
                "Email": email,
                "PDF": str(caminho_pdf),
                "Status": "ERRO",
                "Data": datetime.now(),
                "Erro": str(erro),
            })

    df_envio = pd.DataFrame(resultados)

    # Salva log de envio com fallback para caminho estendido
    garantir_diretorio(pasta_logs)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arq_log = pasta_logs / f"log_envio_{timestamp}.xlsx"
    try:
        df_envio.to_excel(arq_log, index=False)
    except (FileNotFoundError, OSError):
        df_envio.to_excel(caminho_estendido(arq_log), index=False)
    seguro_print(f"\n  {verde('[OK]')} {dim(f'Log de auditoria salvo em: {arq_log.name}')}")

    return df_envio
