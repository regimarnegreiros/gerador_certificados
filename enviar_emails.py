"""
Script da Etapa 4: Envio de e-mails com os certificados anexados através do Microsoft Outlook.
Origem do executável 'enviar_emails.exe'.
Permite teste com abertura de rascunho e envio em lote com confirmação.
"""

import sys
import os

# Configuração imediata de terminal e feedback inicial instantâneo
if sys.stdout:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if sys.stderr:
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

if os.name == "nt":
    os.system("")
print("\n  \033[93m\033[1m⏳ Processando... Aguarde um momento!\033[0m\n", flush=True)

from pathlib import Path
import traceback

diretorio_atual = Path.cwd()
diretorio_raiz = Path(__file__).resolve().parent

if str(diretorio_raiz) not in sys.path:
    sys.path.insert(0, str(diretorio_raiz))
if str(diretorio_raiz.parent) not in sys.path:
    sys.path.insert(0, str(diretorio_raiz.parent))

from core.ui import (
    inicializar_terminal,
    limpar_tela,
    exibir_mensagem_aguarde,
    exibir_cabecalho,
    exibir_sucesso,
    exibir_aviso,
    exibir_erro,
    exibir_info,
    formatar_opcao,
    pausar_e_sair,
    pausar_e_continuar,
    verde,
    vermelho,
    amarelo,
    ciano,
    negrito,
    dim,
)
from core.validator import listar_arquivos_validos
from core.config import obter_configuracao
from core.caminhos import (
    existe_com_fallback,
    listar_diretorio_com_fallback,
    verificar_profundidade_caminho,
    LIMITE_CARACTERES_DIRETORIO_CURSO,
)
from core.certificados import carregar_relatorio_presenca, filtrar_aptos, nome_seguro_arquivo
from core.emails import enviar_email_teste, enviar_todos_emails, localizar_certificado_pdf

inicializar_terminal()


def encontrar_relatorio_presenca(diretorio_curso: Path) -> Path:
    pasta_relatorios = diretorio_curso / "saida" / "relatorio_presencas"
    if not existe_com_fallback(pasta_relatorios):
        raise FileNotFoundError(
            "A pasta 'saida/relatorio_presencas' não foi encontrada.\n"
            "Execute primeiramente a Etapa 2 ('gerar_relatorio.exe')."
        )

    arquivos = listar_arquivos_validos(pasta_relatorios, (".xlsx",))
    if not arquivos:
        raise FileNotFoundError("Nenhum relatório de presença (.xlsx) foi encontrado.")
    return arquivos[0]


def main():
    # Validação preventiva de profundidade do caminho
    eh_profundo, tam, msg_orientacao = verificar_profundidade_caminho(diretorio_atual, LIMITE_CARACTERES_DIRETORIO_CURSO)
    if eh_profundo:
        limpar_tela()
        exibir_cabecalho(
            "ETAPA 4: DISPARO DE E-MAILS COM CERTIFICADOS - ANM",
            f"Curso: {diretorio_atual.name}"
        )
        print()
        exibir_erro(
            "A pasta do curso está em um caminho excessivamente longo ou profundo:",
            detalhe=msg_orientacao,
            dica="Mova a pasta do curso para um diretório menos profundo (ex: C:\\Cursos\\) antes de continuar."
        )
        pausar_e_sair(1)

    while True:
        limpar_tela()
        exibir_cabecalho(
            "ETAPA 4: DISPARO DE E-MAILS COM CERTIFICADOS - ANM",
            f"Curso: {diretorio_atual.name}"
        )

        pasta_saida = diretorio_atual / "saida"
        pasta_pdf = pasta_saida / "certificados_pdf"
        pasta_logs = pasta_saida / "logs"

        tem_pdfs = existe_com_fallback(pasta_pdf) and any(
            f.is_file() and f.suffix.lower() == ".pdf"
            for f in listar_diretorio_com_fallback(pasta_pdf)
        )
        if not tem_pdfs:
            print()
            exibir_erro(
                "Nenhum certificado em PDF encontrado em 'saida/certificados_pdf/'.",
                dica="Execute a Etapa 3 ('gerar_certificados.exe') antes de enviar e-mails."
            )
            pausar_e_sair(1)

        try:
            config = obter_configuracao(diretorio_atual)
            caminho_relatorio = encontrar_relatorio_presenca(diretorio_atual)
            df_relatorio = carregar_relatorio_presenca(caminho_relatorio)
            aptos, _ = filtrar_aptos(df_relatorio)
        except PermissionError:
            print()
            exibir_erro(
                "Permissão negada ao abrir a planilha de relatório de presença.",
                detalhe=f"O arquivo '{caminho_relatorio.name}' está aberto no Microsoft Excel ou bloqueado pelo sistema.",
                dica="Feche a janela do Microsoft Excel e tente executar o programa novamente."
            )
            pausar_e_sair(1)
        except Exception as e:
            print()
            exibir_erro(f"Erro ao ler configurações ou relatório: {e}")
            pausar_e_sair(1)

        print(f"  {ciano(negrito('👥 AUDITORIA:'))} {verde(negrito(str(len(aptos))))} participante(s) apto(s) para receber certificado.\n")
        print("  Selecione o modo de disparo desejado:\n")
        print(formatar_opcao("1", "Modo de Teste: Abrir rascunho de conferência no Outlook para 1 pessoa"))
        print(formatar_opcao("2", "Modo de Disparo Geral: Enviar certificados para TODOS os participantes aptos"))
        print(formatar_opcao("0", "Cancelar e sair"))
        print()

        try:
            opcao = input(f"  {amarelo(negrito('Opção desejada [1, 2 ou 0]:'))} ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n  Operação cancelada.\n")
            pausar_e_sair(0)

        if opcao == "0":
            print("\n  Operação cancelada pelo usuário.")
            pausar_e_sair(0)

        if opcao in ("1", "2"):
            break
        else:
            exibir_aviso("Opção inválida. Escolha entre 1, 2 ou 0.")
            pausar_e_continuar()

    # ----------------------------------------------------
    # MODO 1: TESTE
    # ----------------------------------------------------
    if opcao == "1":
        limpar_tela()
        exibir_cabecalho(
            "MODO DE TESTE (CONFERÊNCIA NO OUTLOOK)",
            "Visualização prévia do certificado e mensagem"
        )

        lista_aptos = aptos.reset_index(drop=True)
        limite = min(5, len(lista_aptos))
        print("  Primeiros participantes aptos para escolha de teste:\n")
        for idx in range(limite):
            row = lista_aptos.iloc[idx]
            email = row.get("Email", "Sem e-mail cadastrado")
            print(f"  {ciano(negrito(f'[{idx + 1}]'))} {row['Nome']} {dim(f'({email})')}")
        print(formatar_opcao("M", "Digitar manualmente outro e-mail de destino para teste"))
        print(formatar_opcao("0", "Cancelar e voltar"))
        print()

        try:
            escolha = input(f"  {amarelo(negrito(f'Escolha [1 a {limite}], [M] ou [0]:'))} ").strip()
        except (KeyboardInterrupt, EOFError):
            return

        if escolha in ("0", ""):
            print("\n  Teste cancelado.")
            pausar_e_sair(0)

        if escolha.upper() == "M":
            email_teste = input(f"\n  {amarelo('Digite o e-mail de destino do teste:')} ").strip()
            if not email_teste:
                exibir_erro("E-mail inválido. Operação cancelada.")
                pausar_e_sair(1)
            linha_selecionada = lista_aptos.iloc[0]
            nome_teste = linha_selecionada["Nome"]
        else:
            try:
                num = int(escolha) - 1
                if 0 <= num < len(lista_aptos):
                    linha_selecionada = lista_aptos.iloc[num]
                    nome_teste = linha_selecionada["Nome"]
                    email_teste = str(linha_selecionada.get("Email", "")).strip()
                else:
                    exibir_erro("Número inválido.")
                    pausar_e_sair(1)
            except ValueError:
                exibir_erro("Opção inválida.")
                pausar_e_sair(1)

        # Localiza o PDF do participante (suporta CERT_ e CERTIFICADO_)
        caminho_pdf = localizar_certificado_pdf(pasta_pdf, nome_teste)

        if not existe_com_fallback(caminho_pdf):
            exibir_erro(f"O certificado PDF não foi encontrado: {caminho_pdf.name}")
            pausar_e_sair(1)

        exibir_mensagem_aguarde(f"Conectando ao Outlook e preparando e-mail de teste para '{nome_teste}'...")

        try:
            enviar_email_teste(
                caminho_pdf=caminho_pdf,
                nome=nome_teste,
                email_destinatario=email_teste,
                config=config,
                abrir_no_outlook=True,
            )
            limpar_tela()
            exibir_cabecalho("TESTE GERADO COM SUCESSO!", "Verifique a janela do Microsoft Outlook")
            exibir_sucesso(f"E-mail de teste exibido na tela para: {verde(nome_teste)}")
            exibir_info(f"Destinatário informado: {email_teste}")
            exibir_info(f"Certificado anexado:    {caminho_pdf.name}")

            print(f"\n  {ciano(negrito('📋 DICA:'))}")
            print("  Verifique no Outlook se o assunto, o texto e o certificado em anexo estão corretos.")
            print("  Quando estiver pronto, execute novamente o programa e escolha a Opção [2] para o envio geral.")

            pausar_e_sair(0, "Pressione [ENTER] para fechar esta janela...")
        except Exception as e:
            print()
            exibir_erro(
                f"Erro durante o teste no Outlook: {e}",
                dica="Certifique-se de que o Microsoft Outlook está aberto e configurado com sua conta de e-mail."
            )
            pausar_e_sair(1)

    # ----------------------------------------------------
    # MODO 2: DISPARO GERAL
    # ----------------------------------------------------
    elif opcao == "2":
        limpar_tela()
        exibir_cabecalho(
            "DISPARO EM MASSA DE E-MAILS",
            "Etapa 4 - Envio Definitivo via Outlook"
        )

        print(f"  {amarelo(negrito('⚠️  ATENÇÃO: DISPARO AUTOMÁTICO VIA MICROSOFT OUTLOOK'))}")
        print(f"     • Total de participantes que receberão o certificado: {verde(negrito(str(len(aptos))))}")
        print(f"     • Cada e-mail será enviado individualmente com o PDF correspondente em anexo.")
        print(f"     • Esta operação {vermelho(negrito('NÃO'))} poderá ser cancelada após o início do envio.\n")

        try:
            confirmacao = input(f"  Digite {verde(negrito('ENVIAR'))} (em maiúsculas) para confirmar ou [0] para cancelar: ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n  Disparo cancelado.")
            pausar_e_sair(0)

        if confirmacao.upper() != "ENVIAR":
            print(f"\n  {dim('Disparo cancelado pelo usuário.')}")
            pausar_e_sair(0)

        exibir_mensagem_aguarde(f"Iniciando envio dos certificados para {len(aptos)} participante(s)...")

        try:
            df_envio = enviar_todos_emails(
                df_aptos=aptos,
                pasta_pdf=pasta_pdf,
                pasta_logs=pasta_logs,
                config=config,
            )

            enviados = len(df_envio[df_envio["Status"] == "ENVIADO"])
            falhas = len(df_envio[df_envio["Status"] != "ENVIADO"])

            limpar_tela()
            exibir_cabecalho("DISPARO DE E-MAILS CONCLUÍDO!", "Etapa 4 Finalizada")

            print(f"  {ciano(negrito('📊 RESUMO DO PROCESSAMENTO:'))}")
            print(f"     • Total processado:       {len(df_envio)}")
            print(f"     • Enviados com sucesso:   {verde(negrito(str(enviados)))}")
            if falhas > 0:
                print(f"     • Falhas ou pendências:   {amarelo(str(falhas))}")

            print(f"\n  {ciano(negrito('📁 REGISTRO DE AUDITORIA:'))}")
            print(f"     O log detalhado de cada envio foi salvo na pasta {verde(str(pasta_logs))}.")

            pausar_e_sair(0, "Pressione [ENTER] para fechar esta janela...")
        except Exception as e:
            print()
            exibir_erro(
                f"Erro durante o disparo em massa de e-mails: {e}",
                dica="Verifique se o Microsoft Outlook não foi fechado ou se há alguma caixa de diálogo bloqueando o envio."
            )
            traceback.print_exc()
            pausar_e_sair(1)


if __name__ == "__main__":
    main()
