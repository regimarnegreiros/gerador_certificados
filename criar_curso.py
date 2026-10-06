"""
Script principal: Criador e Gerenciador de Cursos - ANM.
Origem do executável 'criar_curso.exe'.

Funções:
1. Criar novo curso com a estrutura de diretórios, leia-me.txt, variaveis.txt e executáveis.
2. Identificar cursos já existentes e restaurar/recriar executáveis caso tenham sido apagados.
3. Limpar executáveis de cursos finalizados para economizar espaço em disco.
4. Totalmente autônomo (Standalone): embute internamente os executáveis das etapas e templates,
   podendo ser enviado individualmente para qualquer pessoa sem nenhuma pasta pré-existente.
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

import shutil
import re
import traceback
from pathlib import Path
from typing import List

# Diretórios de execução e de arquivos internos (quando empacotado pelo PyInstaller)
if getattr(sys, "frozen", False):
    diretorio_usuario = Path(sys.executable).resolve().parent
    diretorio_interno = Path(getattr(sys, "_MEIPASS", sys.executable)).resolve()
else:
    diretorio_usuario = Path(__file__).resolve().parent
    diretorio_interno = Path(__file__).resolve().parent

if str(diretorio_interno) not in sys.path:
    sys.path.insert(0, str(diretorio_interno))
if str(diretorio_usuario) not in sys.path:
    sys.path.insert(0, str(diretorio_usuario))

from core.ui import (
    inicializar_terminal,
    limpar_tela,
    exibir_mensagem_aguarde,
    exibir_cabecalho,
    exibir_sucesso,
    exibir_aviso,
    exibir_erro,
    exibir_info,
    exibir_passo,
    formatar_opcao,
    pausar_e_continuar,
    pausar_e_sair,
    confirmar_acao,
    verde,
    amarelo,
    ciano,
    negrito,
    dim,
)
from core.caminhos import (
    garantir_diretorio,
    abrir_arquivo,
    copiar_arquivo_com_fallback,
    existe_com_fallback,
    listar_diretorio_com_fallback,
    verificar_profundidade_caminho,
    LIMITE_CARACTERES_DIRETORIO_CURSO,
)

inicializar_terminal()

PASTAS_ENTRADA = [
    "lista_convidados",
    "listas_de_presenca",
    "lista_servidores",
    "modelo_certificado",
]

SUBPASTAS_SAIDA = [
    "relatorio_presencas",
    "certificados_pdf",
    "certificados_juntos_pdf",
    "certificados_juntos_ppt",
    "logs",
]

EXECUTAVEIS_ETAPAS = [
    "gerar_relatorio.exe",
    "gerar_certificados.exe",
    "enviar_emails.exe",
]

SCRIPTS_PYTHON_ETAPAS = [
    "gerar_relatorio.py",
    "gerar_certificados.py",
    "enviar_emails.py",
]


def formatar_nome_pasta(nome_curso: str) -> str:
    """Gera um nome de pasta amigável e compacto para o curso (ex: 'curso_reforma_tributaria')."""
    limpo = unicodedata_normalizar(nome_curso).lower()
    limpo = re.sub(r"[^a-z0-9]+", "_", limpo).strip("_")
    # Limita a 40 caracteres para evitar ultrapassar o limite de caminhos longos no Windows
    if len(limpo) > 40:
        limpo = limpo[:40].rstrip("_")
    return f"curso_{limpo}" if not limpo.startswith("curso_") else limpo


def unicodedata_normalizar(texto: str) -> str:
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def detectar_cursos_existentes(pasta_raiz: Path) -> List[Path]:
    """Identifica subdiretórios que representam cursos (possuem variaveis.txt ou saida/)."""
    cursos = []
    if not existe_com_fallback(pasta_raiz):
        return cursos

    for item in listar_diretorio_com_fallback(pasta_raiz):
        if item.is_dir() and not item.name.startswith(".") and item.name not in ("core", "templates", "scratch", "dist", "build", "_bin", "tests"):
            if existe_com_fallback(item / "variaveis.txt") or existe_com_fallback(item / "saida") or any(existe_com_fallback(item / p) for p in PASTAS_ENTRADA):
                cursos.append(item)
    return sorted(cursos, key=lambda p: p.name)


def copiar_executaveis_para_curso(pasta_destino: Path):
    """
    Copia os executáveis das etapas para o curso.
    Prioriza os arquivos embutidos internamente no próprio criar_curso.exe (sys._MEIPASS),
    tornando o criar_curso.exe 100% autônomo e portátil.
    """
    garantir_diretorio(pasta_destino)
    encontrou_algum = False

    # 1. Procura executáveis compilados (.exe) reais (> 0 bytes)
    for exe in EXECUTAVEIS_ETAPAS:
        locais = [
            diretorio_interno / "_bin" / exe,
            diretorio_interno / exe,
            diretorio_usuario / "_bin" / exe,
            diretorio_usuario / exe,
            diretorio_usuario / "dist" / exe,
        ]
        for origem in locais:
            if existe_com_fallback(origem) and origem.is_file() and origem.stat().st_size > 0:
                try:
                    copiar_arquivo_com_fallback(origem, pasta_destino / exe)
                    encontrou_algum = True
                except Exception as e:
                    print()
                    exibir_aviso(f"Aviso ao copiar '{exe}': {e}")
                break

    # 2. Se em ambiente Python de desenvolvimento (sem os .exes gerados ainda), copia os scripts .py correspondentes
    if not encontrou_algum:
        for py_script in SCRIPTS_PYTHON_ETAPAS:
            for d in (diretorio_interno, diretorio_usuario):
                origem_py = d / py_script
                if existe_com_fallback(origem_py):
                    try:
                        copiar_arquivo_com_fallback(origem_py, pasta_destino / py_script)
                        encontrou_algum = True
                    except Exception as e:
                        print()
                        exibir_aviso(f"Aviso ao copiar '{py_script}': {e}")
                    break


def carregar_template(nome_arquivo: str) -> str:
    """Carrega template interno do pacote ou do diretório do usuário com fallback embutido."""
    locais = [
        diretorio_interno / "templates" / nome_arquivo,
        diretorio_usuario / "templates" / nome_arquivo,
    ]
    for loc in locais:
        if existe_com_fallback(loc) and loc.is_file():
            try:
                with abrir_arquivo(loc, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception:
                pass
    return ""


def criar_novo_curso():
    limpar_tela()
    exibir_cabecalho("CRIAR NOVO CURSO - ANM", "Estrutura de Pastas e Configurações Automáticas")

    print(f"  Digite o nome oficial do curso abaixo.")
    print(f"  {dim('Exemplo: Reforma Tributária e Impactos na Mineração')}\n")

    try:
        nome_curso = input(f"  {amarelo(negrito('Nome do curso:'))} ").strip()
    except (KeyboardInterrupt, EOFError):
        return

    if not nome_curso:
        print()
        exibir_erro("O nome do curso não pode ser vazio.")
        pausar_e_continuar("Pressione [ENTER] para voltar ao menu...")
        return

    nome_curso = nome_curso.strip('"\'')

    nome_pasta_sugerida = formatar_nome_pasta(nome_curso)
    print(f"\n  Pasta sugerida: {ciano(negrito(nome_pasta_sugerida))}")
    print(f"  {dim('Pressione [ENTER] para aceitar ou digite um nome alternativo.')}")

    try:
        confirmacao_pasta = input(f"  {amarelo('Nome da pasta:')} ").strip()
    except (KeyboardInterrupt, EOFError):
        return

    if confirmacao_pasta:
        nome_pasta = re.sub(r'[*?"<>:|/\\]', "", confirmacao_pasta).strip(' .\'"')
        if not nome_pasta:
            nome_pasta = nome_pasta_sugerida
    else:
        nome_pasta = nome_pasta_sugerida

    pasta_curso = diretorio_usuario / nome_pasta

    # Validação de profundidade do caminho
    eh_profundo, tam, msg_orientacao = verificar_profundidade_caminho(pasta_curso, LIMITE_CARACTERES_DIRETORIO_CURSO)
    if eh_profundo:
        print()
        exibir_aviso(msg_orientacao)
        if not confirmar_acao("Deseja criar a pasta neste local mesmo assim?", padrao=False):
            print(f"\n  {dim('Criação cancelada. Mova o sistema para uma pasta menos profunda.')}")
            pausar_e_continuar("Pressione [ENTER] para voltar ao menu...")
            return

    if existe_com_fallback(pasta_curso):
        print()
        exibir_aviso(f"A pasta '{nome_pasta}' já existe no diretório.")
        if not confirmar_acao("Deseja criar/atualizar a estrutura dentro dela mesmo assim?", padrao=False):
            print(f"\n  {dim('Criação cancelada pelo usuário.')}")
            pausar_e_continuar("Pressione [ENTER] para voltar ao menu...")
            return

    exibir_mensagem_aguarde("Criando pastas e configurando arquivos do curso...")

    try:
        garantir_diretorio(pasta_curso)

        # 1. Cria pastas de entrada
        for p in PASTAS_ENTRADA:
            garantir_diretorio(pasta_curso / p)

        # 2. Cria pastas de saída
        pasta_saida = pasta_curso / "saida"
        garantir_diretorio(pasta_saida)
        for sub in SUBPASTAS_SAIDA:
            garantir_diretorio(pasta_saida / sub)

        # 3. Cria variaveis.txt preenchido com o nome do curso
        conteudo_var = carregar_template("variaveis_template.txt")
        if conteudo_var:
            conteudo_var = conteudo_var.replace("{NOME_DO_CURSO}", nome_curso)
        else:
            conteudo_var = (
                f"# ============================================================\n"
                f"# CONFIGURAÇÕES DO CURSO E CERTIFICADOS\n"
                f"# ============================================================\n\n"
                f"# ------------------------------------------------------------\n"
                f"# 1. DADOS GERAIS DO CURSO\n"
                f"# ------------------------------------------------------------\n"
                f"NOME_DO_CURSO = {nome_curso}\n"
                f"CARGA_HORARIA = 20\n"
                f"CUSTO_TOTAL_EM_REAIS = 0.00\n"
                f"DATA_INICIO = 17/08/2026\n"
                f"DATA_FIM = 21/08/2026\n\n"
                f"# ------------------------------------------------------------\n"
                f"# 2. INFORMAÇÕES DE EMISSÃO DO CERTIFICADO\n"
                f"# ------------------------------------------------------------\n"
                f"LOCAL_EMISSAO = Brasília — DF\n"
                f"DATA_EMISSAO = 21/08/2026\n\n"
                f"# ------------------------------------------------------------\n"
                f"# 3. CRITÉRIOS DE APROVAÇÃO E APURAÇÃO DE PRESENÇA\n"
                f"# ------------------------------------------------------------\n"
                f"PERCENTUAL_MINIMO_CARGA = 0.75\n"
                f"PERCENTUAL_MINIMO_ENCONTROS = 0.75\n"
                f"TEMPO_MINIMO_PRESENCA_MINUTOS = 30\n\n"
                f"# ------------------------------------------------------------\n"
                f"# 4. E-MAILS A IGNORAR NOS RELATÓRIOS\n"
                f"# ------------------------------------------------------------\n"
                f"[INICIO_IGNORAR_EMAILS]\n"
                f"organizador1@exemplo.gov.br\n"
                f"organizador2@exemplo.gov.br\n"
                f"organizador3@exemplo.gov.br\n"
                f"[FIM_IGNORAR_EMAILS]\n\n"
                f"# ------------------------------------------------------------\n"
                f"# 5. COMUNICAÇÃO E ENVIO POR E-MAIL\n"
                f"# ------------------------------------------------------------\n"
                f"ASSUNTO_EMAIL = Certificado - Curso {{CURSO}}\n\n"
                f"[INICIO_CORPO_EMAIL]\n"
                f"Prezada(o) {{nome}},\n\n"
                f"Esta é uma mensagem automatizada. Não é necessário respondê-la.\n\n"
                f"Encaminhamos, em anexo, o certificado referente à sua participação no curso “{{CURSO}}”.\n\n"
                f"Agradecemos pela participação.\n\n"
                f"Atenciosamente,\n\n"
                f"Superintendência de Gestão de Pessoas – SGP/ANM\n"
                f"[FIM_CORPO_EMAIL]\n"
            )

        with abrir_arquivo(pasta_curso / "variaveis.txt", "w", encoding="utf-8") as f:
            f.write(conteudo_var)

        # 4. Cria leia-me.txt
        conteudo_leiame = carregar_template("leia-me_template.txt")
        if not conteudo_leiame:
            conteudo_leiame = (
                "Instruções de uso do Gerador de Certificados ANM.\n"
                "1. Preencha 'variaveis.txt'\n"
                "2. Coloque os arquivos nas pastas correspondentes\n"
                "3. Execute 'gerar_relatorio.exe' -> 'gerar_certificados.exe' -> 'enviar_emails.exe'\n"
            )

        with abrir_arquivo(pasta_curso / "leia-me.txt", "w", encoding="utf-8") as f:
            f.write(conteudo_leiame)

        # 5. Copia os executáveis embutidos
        copiar_executaveis_para_curso(pasta_curso)

        limpar_tela()
        exibir_cabecalho("CURSO CRIADO COM SUCESSO!", "Siga os passos abaixo para dar continuidade")
        exibir_sucesso(f"Pasta criada: {verde(pasta_curso.name)}")
        exibir_info(f"Caminho completo: {pasta_curso}")

        print(f"\n  {ciano(negrito('📋 PRÓXIMOS PASSOS:'))}")
        print(f"  {ciano('1.')} Abra a pasta {verde(nome_pasta)}.")
        print(f"  {ciano('2.')} Abra o arquivo {amarelo('variaveis.txt')} e confira datas e carga horária.")
        print(f"  {ciano('3.')} Coloque seus arquivos nas pastas de entrada:")
        print(f"     • {negrito('lista_convidados/')}: planilha/lista de convocados (.csv)")
        print(f"     • {negrito('lista_servidores/')}: cadastro geral de servidores (.xlsx)")
        print(f"     • {negrito('listas_de_presenca/')}: relatórios de presença diários do Teams (.csv/.tsv)")
        print(f"     • {negrito('modelo_certificado/')}: apresentação do PowerPoint (.pptx ou .pptm)")
        print(f"  {ciano('4.')} Dê duplo clique em {verde('gerar_relatorio.exe')} para iniciar a validação.")

        pausar_e_continuar("Pressione [ENTER] para voltar ao menu principal...")

    except PermissionError as e:
        print()
        exibir_erro(
            f"Permissão negada ao criar pastas ou arquivos do curso: {e}",
            detalhe=f"Caminho: {pasta_curso}",
            dica="Verifique se a pasta não está aberta no Explorer ou bloqueada pelo OneDrive/Antivírus."
        )
        traceback.print_exc()
        pausar_e_continuar("Pressione [ENTER] para voltar ao menu principal...")
    except OSError as e:
        print()
        exibir_erro(
            f"Erro do sistema de arquivos ao criar o curso: {e}",
            detalhe=f"Caminho: {pasta_curso}",
            dica="Verifique se o caminho não é excessivamente longo para o Windows ou contém caracteres inválidos."
        )
        traceback.print_exc()
        pausar_e_continuar("Pressione [ENTER] para voltar ao menu principal...")
    except Exception as e:
        print()
        exibir_erro(f"Erro inesperado ao criar estrutura do curso: {e}")
        traceback.print_exc()
        pausar_e_continuar("Pressione [ENTER] para voltar ao menu principal...")


def restaurar_executaveis_cursos():
    limpar_tela()
    exibir_cabecalho("RESTAURAR EXECUTÁVEIS NOS CURSOS", "Recupere os executáveis de apoio apagados")

    exibir_mensagem_aguarde("Localizando cursos disponíveis...")
    cursos = detectar_cursos_existentes(diretorio_usuario)
    limpar_tela()
    exibir_cabecalho("RESTAURAR EXECUTÁVEIS NOS CURSOS", "Recupere os executáveis de apoio apagados")

    if not cursos:
        exibir_aviso("Nenhum curso existente foi encontrado neste diretório.")
        pausar_e_continuar("Pressione [ENTER] para voltar ao menu principal...")
        return

    print("  Selecione o curso desejado:\n")
    for idx, c in enumerate(cursos, start=1):
        print(formatar_opcao(str(idx), c.name))
    print(formatar_opcao("T", "Restaurar em TODOS os cursos acima"))
    print(formatar_opcao("0", "Voltar ao menu principal"))
    print()

    try:
        escolha = input(f"  {amarelo(negrito('Escolha uma opção:'))} ").strip().upper()
    except (KeyboardInterrupt, EOFError):
        return

    if escolha in ("0", ""):
        return

    cursos_alvo = []
    if escolha == "T":
        cursos_alvo = cursos
    else:
        try:
            num = int(escolha) - 1
            if 0 <= num < len(cursos):
                cursos_alvo = [cursos[num]]
            else:
                exibir_erro("Opção inválida.")
                pausar_e_continuar()
                return
        except ValueError:
            exibir_erro("Opção inválida.")
            pausar_e_continuar()
            return

    exibir_mensagem_aguarde("Copiando executáveis para o(s) curso(s) selecionado(s)...")

    for c in cursos_alvo:
        copiar_executaveis_para_curso(c)
        exibir_sucesso(f"Executáveis restaurados na pasta: '{c.name}'")

    pausar_e_continuar("Pressione [ENTER] para voltar ao menu principal...")


def limpar_executaveis_cursos():
    limpar_tela()
    exibir_cabecalho("LIMPAR EXECUTÁVEIS DE CURSOS", "Economiza espaço em disco mantendo seus relatórios intactos")

    exibir_mensagem_aguarde("Verificando cursos existentes...")
    cursos = detectar_cursos_existentes(diretorio_usuario)
    limpar_tela()
    exibir_cabecalho("LIMPAR EXECUTÁVEIS DE CURSOS", "Economiza espaço em disco mantendo seus relatórios intactos")

    if not cursos:
        exibir_aviso("Nenhum curso existente foi encontrado.")
        pausar_e_continuar("Pressione [ENTER] para voltar ao menu principal...")
        return

    print("  Cursos disponíveis para limpeza:\n")
    for idx, c in enumerate(cursos, start=1):
        exes = [f for f in listar_diretorio_com_fallback(c) if f.is_file() and f.suffix.lower() == ".exe"]
        status = verde(f"({len(exes)} executável(is))") if exes else dim("(Sem executáveis)")
        print(f"  {ciano(negrito(f'[{idx}]'))} {c.name} {status}")
    print(formatar_opcao("0", "Voltar ao menu principal"))
    print()

    try:
        escolha = input(f"  {amarelo(negrito('Escolha o curso para remover os executáveis [0 para voltar]:'))} ").strip()
    except (KeyboardInterrupt, EOFError):
        return

    if escolha in ("0", ""):
        return

    try:
        num = int(escolha) - 1
        if 0 <= num < len(cursos):
            curso_selecionado = cursos[num]
            exes = [f for f in listar_diretorio_com_fallback(curso_selecionado) if f.is_file() and f.suffix.lower() == ".exe"]
            if not exes:
                exibir_info(f"A pasta '{curso_selecionado.name}' já não possui nenhum executável.")
                pausar_e_continuar()
                return

            if not confirmar_acao(f"Deseja realmente remover os {len(exes)} executáveis de '{curso_selecionado.name}'?", padrao=False):
                print(f"\n  {dim('Operação cancelada.')}")
                pausar_e_continuar()
                return

            exibir_mensagem_aguarde("Removendo executáveis...")
            removidos = 0
            for exe in exes:
                try:
                    exe.unlink()
                    removidos += 1
                except Exception as e:
                    exibir_erro(f"Erro ao remover {exe.name}", detalhe=str(e))

            exibir_sucesso(f"{removidos} arquivo(s) executável(is) removido(s) com sucesso!")
            exibir_info("Todos os relatórios, certificados em PDF e dados permanecem 100% intactos.")
            exibir_info("Se precisar rodar novamente no futuro, basta usar a opção 'Restaurar executáveis'.")
            pausar_e_continuar("Pressione [ENTER] para voltar ao menu principal...")
        else:
            exibir_erro("Opção inválida.")
            pausar_e_continuar()
    except ValueError:
        exibir_erro("Opção inválida.")
        pausar_e_continuar()


def main():
    while True:
        limpar_tela()
        exibir_cabecalho(
            "SISTEMA GERADOR DE CERTIFICADOS - ANM",
            "Superintendência de Gestão de Pessoas (SGP)"
        )

        eh_profundo, tam, msg_orientacao = verificar_profundidade_caminho(diretorio_usuario, LIMITE_CARACTERES_DIRETORIO_CURSO)
        if eh_profundo:
            exibir_aviso(msg_orientacao)
            print()

        print("  Selecione a operação desejada:\n")
        print(formatar_opcao("1", "Criar novo curso (gerar estrutura de pastas e arquivos)"))
        print(formatar_opcao("2", "Restaurar / recriar executáveis em cursos existentes"))
        print(formatar_opcao("3", "Limpar executáveis de cursos finalizados (economizar espaço)"))
        print(formatar_opcao("0", "Sair do sistema"))
        print()

        try:
            opcao = input(f"  {amarelo(negrito('Opção desejada:'))} ").strip()
        except (KeyboardInterrupt, EOFError):
            print("\n\n  Encerrando o programa. Até logo!\n")
            break

        if opcao == "1":
            criar_novo_curso()
        elif opcao == "2":
            restaurar_executaveis_cursos()
        elif opcao == "3":
            limpar_executaveis_cursos()
        elif opcao == "0":
            print(f"\n  {verde('Encerrando o programa. Bom trabalho!')}\n")
            break
        else:
            exibir_aviso("Opção inválida. Digite 1, 2, 3 ou 0.")
            pausar_e_continuar()


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):
        print("\n\n  Encerrando o programa. Até logo!\n")
    except Exception as e:
        print()
        exibir_erro(f"Erro inesperado no sistema: {e}")
        traceback.print_exc()
        pausar_e_sair(1, "Pressione [ENTER] para fechar esta janela...")
