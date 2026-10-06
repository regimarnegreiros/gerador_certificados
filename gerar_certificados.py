"""
Script da Etapa 3: Geração dos certificados individuais (PPT/PDF) e junção em PDF único.
Origem do executável 'gerar_certificados.exe'.
Permite gerar certificados e também executar a junção de PDFs de forma independente.
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
    amarelo,
    ciano,
    negrito,
    dim,
)
from core.validator import listar_arquivos_validos
from core.config import obter_configuracao
from core.caminhos import (
    existe_com_fallback,
    garantir_diretorio,
    verificar_profundidade_caminho,
    LIMITE_CARACTERES_DIRETORIO_CURSO,
    esvaziar_pasta,
    obter_nome_arquivo_junto,
    reduzir_nome_curso_para_arquivo,
)
from core.certificados import (
    carregar_relatorio_presenca,
    filtrar_aptos,
    gerar_todos_certificados,
    gerar_powerpoint_consolidado,
    juntar_certificados_pdf,
    sanitizar_nome_arquivo,
)

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
        raise FileNotFoundError(
            "Nenhum relatório de presença (.xlsx) foi encontrado em 'saida/relatorio_presencas/'.\n"
            "Execute primeiramente a Etapa 2 ('gerar_relatorio.exe')."
        )
    return arquivos[0]


def encontrar_modelo_ppt(diretorio_curso: Path) -> Path:
    pasta_modelo = diretorio_curso / "modelo_certificado"
    if not existe_com_fallback(pasta_modelo):
        raise FileNotFoundError(
            "A pasta 'modelo_certificado' não foi encontrada."
        )
    modelos = listar_arquivos_validos(pasta_modelo, (".pptx", ".pptm"))
    if not modelos:
        raise FileNotFoundError(
            "Nenhum modelo PowerPoint (.pptx ou .pptm) foi encontrado na pasta 'modelo_certificado/'."
        )
    return modelos[0]


def main():
    # Validação preventiva de profundidade do caminho
    eh_profundo, tam, msg_orientacao = verificar_profundidade_caminho(diretorio_atual, LIMITE_CARACTERES_DIRETORIO_CURSO)
    if eh_profundo:
        limpar_tela()
        exibir_cabecalho(
            "ETAPA 3: GERADOR DE CERTIFICADOS - ANM",
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
            "ETAPA 3: GERADOR DE CERTIFICADOS - ANM",
            f"Curso: {diretorio_atual.name}"
        )

        pasta_saida = diretorio_atual / "saida"
        pasta_pdf = pasta_saida / "certificados_pdf"
        pasta_juntos_pdf = pasta_saida / "certificados_juntos_pdf"
        pasta_juntos_ppt = pasta_saida / "certificados_juntos_ppt"
        pasta_logs = pasta_saida / "logs"

        # Pastas legadas para limpeza
        pasta_ppt_legada = pasta_saida / "certificados_ppt"
        pasta_juntos_legada = pasta_saida / "certificados_juntos"

        print("  Selecione a operação desejada:\n")
        print(formatar_opcao("1", "Gerar novos certificados e consolidar (PDFs individuais, PDF único e PPT único) (Recomendado)"))
        print(formatar_opcao("2", "Apenas gerar certificados individuais em PDF (sem consolidar)"))
        print(formatar_opcao("3", "Apenas juntar certificados existentes em PDF (caso tenha adicionado PDFs)"))
        print(formatar_opcao("0", "Cancelar e sair"))
        print()

        try:
            opcao = input(f"  {amarelo(negrito('Opção desejada [1, 2, 3 ou 0]:'))} ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n  Operação cancelada.\n")
            pausar_e_sair(0)

        if opcao == "0":
            print("\n  Operação cancelada pelo usuário.")
            pausar_e_sair(0)

        if opcao in ("1", "2", "3"):
            break
        else:
            exibir_aviso("Opção inválida. Escolha entre 1, 2, 3 ou 0.")
            pausar_e_continuar()

    # ----------------------------------------------------
    # OPÇÃO 3: Apenas juntar PDFs existentes
    # ----------------------------------------------------
    if opcao == "3":
        limpar_tela()
        exibir_cabecalho("JUNÇÃO DE CERTIFICADOS EM PDF", "Consolidando arquivos existentes")
        exibir_mensagem_aguarde("Localizando certificados em PDF...")

        try:
            config = obter_configuracao(diretorio_atual)
            nome_curso = config.get("NOME_DO_CURSO", diretorio_atual.name)
        except Exception:
            nome_curso = diretorio_atual.name

        nome_arq_saida = obter_nome_arquivo_junto(nome_curso, pasta_juntos_pdf, ".pdf")
        caminho_saida_junto_pdf = pasta_juntos_pdf / nome_arq_saida

        try:
            # Limpa arquivos anteriores da pasta de PDFs consolidados
            esvaziar_pasta(pasta_juntos_pdf, (".pdf",))
            juntar_certificados_pdf(pasta_pdf, caminho_saida_junto_pdf)
            limpar_tela()
            exibir_cabecalho("JUNÇÃO CONCLUÍDA COM SUCESSO!", "PDF consolidado gerado")
            exibir_sucesso(f"Arquivo gerado: {caminho_saida_junto_pdf.name}")
            exibir_info(f"Caminho completo: {caminho_saida_junto_pdf}")
            pausar_e_sair(0)
        except Exception as e:
            print()
            exibir_erro(f"Erro ao juntar certificados: {e}")
            pausar_e_sair(1)

    # ----------------------------------------------------
    # OPÇÃO 1 OU 2: Gerar certificados
    # ----------------------------------------------------
    limpar_tela()
    exibir_cabecalho("GERANDO CERTIFICADOS", "Etapa 3 em andamento")
    exibir_mensagem_aguarde("Localizando relatório de presenças e modelo PowerPoint...")

    try:
        caminho_relatorio = encontrar_relatorio_presenca(diretorio_atual)
        caminho_modelo = encontrar_modelo_ppt(diretorio_atual)
        config = obter_configuracao(diretorio_atual)

        exibir_sucesso(f"Relatório de Presença: {ciano(caminho_relatorio.name)}")
        exibir_sucesso(f"Modelo PowerPoint:      {ciano(caminho_modelo.name)}")
    except FileNotFoundError as e:
        print()
        exibir_erro(str(e), dica="Certifique-se de executar 'gerar_relatorio.exe' antes desta etapa.")
        pausar_e_sair(1)
    except Exception as e:
        print()
        exibir_erro(f"Erro ao carregar arquivos: {e}")
        pausar_e_sair(1)

    try:
        df_relatorio = carregar_relatorio_presenca(caminho_relatorio)
        aptos, nao_aptos = filtrar_aptos(df_relatorio)
    except PermissionError:
        print()
        exibir_erro(
            "Permissão negada ao abrir a planilha de relatório de presença.",
            detalhe=f"O arquivo '{caminho_relatorio.name}' está aberto no Microsoft Excel ou bloqueado pelo sistema.",
            dica="Feche a janela do Microsoft Excel e execute o programa novamente."
        )
        pausar_e_sair(1)
    except Exception as e:
        print()
        exibir_erro(
            f"Erro ao carregar o relatório de presença: {e}",
            dica="Certifique-se de que a planilha gerada na Etapa 2 não está corrompida ou aberta em outro programa."
        )
        pausar_e_sair(1)

    print(f"\n  {ciano(negrito('👥 AUDITORIA DE PARTICIPANTES:'))}")
    print(f"     • Total de convocados no relatório: {len(df_relatorio)}")
    print(f"     • Participantes aptos (emitir):    {verde(negrito(str(len(aptos))))}")
    print(f"     • Não aptos (não emitir):          {amarelo(str(len(nao_aptos)))}")

    if len(aptos) == 0:
        print()
        exibir_aviso("Nenhum participante atingiu os critérios de aprovação no relatório de presença.")
        exibir_info("Nenhum certificado precisa ser emitido.")
        pausar_e_sair(0)

    try:
        exibir_mensagem_aguarde("Limpando arquivos antigos das pastas de saída...")
        esvaziar_pasta(pasta_juntos_pdf, (".pdf",))
        esvaziar_pasta(pasta_juntos_ppt, (".pptx", ".pptm"))
        esvaziar_pasta(pasta_pdf, (".pdf",))

        # Remove pastas legadas se existirem
        if existe_com_fallback(pasta_ppt_legada):
            esvaziar_pasta(pasta_ppt_legada)
            try:
                pasta_ppt_legada.rmdir()
            except Exception:
                pass
        if existe_com_fallback(pasta_juntos_legada):
            esvaziar_pasta(pasta_juntos_legada)
            try:
                pasta_juntos_legada.rmdir()
            except Exception:
                pass

        nome_curso = config.get("NOME_DO_CURSO", diretorio_atual.name)
        nome_curso_seguro = sanitizar_nome_arquivo(nome_curso).replace(" ", "_")
        extensao_origem = caminho_modelo.suffix.lower()

        caminho_ppt_junto = None
        caminho_saida_junto_pdf = None
        if opcao == "1":
            nome_arq_ppt = obter_nome_arquivo_junto(nome_curso, pasta_juntos_ppt, extensao_origem)
            nome_arq_pdf = obter_nome_arquivo_junto(nome_curso, pasta_juntos_pdf, ".pdf")
            caminho_ppt_junto = pasta_juntos_ppt / nome_arq_ppt
            caminho_saida_junto_pdf = pasta_juntos_pdf / nome_arq_pdf

        exibir_mensagem_aguarde(f"Iniciando PowerPoint e emitindo {len(aptos)} certificado(s)...")

        df_resultados = gerar_todos_certificados(
            df_aptos=aptos,
            modelo_powerpoint=caminho_modelo,
            pasta_ppt=None,
            pasta_pdf=pasta_pdf,
            pasta_logs=pasta_logs,
            config=config,
            caminho_ppt_junto=caminho_ppt_junto,
            salvar_ppt_individual=False,
        )

        # Se opção 1, executa a junção dos PDFs gerados
        if opcao == "1" and caminho_saida_junto_pdf is not None:
            exibir_mensagem_aguarde("Unificando certificados em um único arquivo PDF consolidado...")
            juntar_certificados_pdf(pasta_pdf, caminho_saida_junto_pdf)

        limpar_tela()
        exibir_cabecalho("ETAPA 3 CONCLUÍDA COM SUCESSO!", "Todos os certificados foram gerados")
        exibir_sucesso(f"{len(aptos)} certificado(s) gerado(s) com êxito!")

        print(f"\n  {ciano(negrito('📁 PASTAS DE SAÍDA:'))}")
        print(f"     • {negrito('PDFs individuais:')}         {verde(str(pasta_pdf))}")
        if opcao == "1":
            print(f"     • {negrito('PowerPoint consolidado:')}   {verde(str(caminho_ppt_junto))}")
            print(f"     • {negrito('PDF consolidado:')}          {verde(str(caminho_saida_junto_pdf))}")
        print(f"     • {negrito('Relatório de logs:')}        {dim(str(pasta_logs))}")

        print(f"\n  {ciano(negrito('▶ PRÓXIMO PASSO:'))}")
        print(f"  Execute {verde(negrito('enviar_emails.exe'))} para disparar os certificados via Microsoft Outlook.")

        pausar_e_sair(0, "Pressione [ENTER] para fechar esta janela...")

    except PermissionError as e:
        print()
        exibir_erro(
            f"Permissão negada ao salvar ou acessar arquivos: {e}",
            dica="Certifique-se de que nenhum arquivo PDF ou PowerPoint de saída está aberto em outro programa."
        )
        traceback.print_exc()
        pausar_e_sair(1)
    except Exception as e:
        print()
        exibir_erro(
            f"Ocorreu um erro durante a geração dos certificados no PowerPoint: {e}",
            dica="Certifique-se de que o Microsoft PowerPoint está instalado e não há janelas travadas."
        )
        traceback.print_exc()
        pausar_e_sair(1)


if __name__ == "__main__":
    main()
