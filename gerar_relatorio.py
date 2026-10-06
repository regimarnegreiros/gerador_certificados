"""
Script da Etapa 2: Validação da estrutura de pastas, arquivos e geração do Relatório de Presenças.
Origem do executável 'gerar_relatorio.exe'.
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

# Permite importação dos módulos mesmo quando executado dentro de subpastas de cursos
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
    exibir_passo,
    pausar_e_sair,
    verde,
    amarelo,
    ciano,
    negrito,
    dim,
)
from core.validator import validar_estrutura_curso, ValidationError
from core.config import obter_configuracao, ConfigValidationError
from core.presenca import gerar_relatorio, sanitizar_nome_arquivo
from core.caminhos import (
    verificar_profundidade_caminho,
    LIMITE_CARACTERES_DIRETORIO_CURSO,
    esvaziar_pasta,
    reduzir_nome_curso_para_arquivo,
)

inicializar_terminal()


def main():
    limpar_tela()
    exibir_cabecalho(
        "ETAPA 2: VALIDAÇÃO E RELATÓRIO DE PRESENÇAS",
        f"Curso atual: {diretorio_atual.name}"
    )

    # Validação preventiva de profundidade do caminho
    eh_profundo, tam, msg_orientacao = verificar_profundidade_caminho(diretorio_atual, LIMITE_CARACTERES_DIRETORIO_CURSO)
    if eh_profundo:
        print()
        exibir_erro(
            "A pasta do curso está em um caminho excessivamente longo ou profundo:",
            detalhe=msg_orientacao,
            dica="Mova a pasta do curso para um diretório menos profundo (ex: C:\\Cursos\\) antes de continuar."
        )
        pausar_e_sair(1)

    # 1. Validação de Pastas e Arquivos
    exibir_passo(1, 3, "Verificando integridade das pastas e arquivos de entrada...")
    try:
        dados_auditoria = validar_estrutura_curso(diretorio_atual)
        exibir_sucesso("Estrutura de pastas e arquivos de entrada validada!")
        if dados_auditoria.get("avisos"):
            print()
            for aviso in dados_auditoria["avisos"]:
                exibir_aviso(aviso)
    except ValidationError as e:
        print()
        exibir_erro(
            "Foram encontrados problemas na estrutura de pastas ou arquivos:",
            dica="Certifique-se de que os arquivos necessários (.csv, .xlsx, .pptx) foram copiados para as respectivas pastas de entrada."
        )
        for erro in e.erros:
            print(f"     • {amarelo(erro)}")
        if e.avisos:
            print(f"\n  {dim('Avisos adicionais:')}")
            for av in e.avisos:
                print(f"     • {dim(av)}")
        pausar_e_sair(1, "Corrija os arquivos acima e pressione [ENTER] para sair...")

    # 2. Leitura e Validação do variaveis.txt
    exibir_passo(2, 3, "Lendo e validando configurações em 'variaveis.txt'...")
    try:
        config = obter_configuracao(diretorio_atual)
        exibir_sucesso("Arquivo 'variaveis.txt' carregado com sucesso:")
        print(f"     • {negrito('Curso:')}                {ciano(config['NOME_DO_CURSO'])}")
        print(f"     • {negrito('Carga Horária:')}        {config['CARGA_HORARIA']} horas")
        periodo_fim = config['DATA_FIM'] or config['DATA_INICIO']
        print(f"     • {negrito('Período:')}              {config['DATA_INICIO']} a {periodo_fim}")
        print(f"     • {negrito('Critério de Frequência:')} Carga Mínima: {int(config['PERCENTUAL_MINIMO_CARGA'] * 100)}% | Tempo por Encontro: {config['TEMPO_MINIMO_PRESENCA_MINUTOS']} min")
    except ConfigValidationError as e:
        print()
        exibir_erro(
            "Erros nas configurações de 'variaveis.txt':",
            dica="Abra o arquivo 'variaveis.txt' no Bloco de Notas para corrigir os campos indicados."
        )
        for erro in e.errors:
            print(f"     • {amarelo(erro)}")
        pausar_e_sair(1, "Corrija o arquivo e pressione [ENTER] para fechar...")
    except Exception as e:
        print()
        exibir_erro(f"Erro inesperado ao ler 'variaveis.txt': {e}")
        pausar_e_sair(1)

    # 3. Execução da Geração do Relatório
    exibir_passo(3, 3, "Processando listas de presença dos encontros e cruzando dados...")
    exibir_mensagem_aguarde("Cruzando dados dos participantes e gerando planilha Excel...")

    arquivos = dados_auditoria["arquivos"]
    caminho_servidores = arquivos["servidores"]
    caminho_convocados = arquivos["convidados"]
    arquivos_presenca = arquivos["presencas"]

    nome_seguro_curso = sanitizar_nome_arquivo(config["NOME_DO_CURSO"])
    pasta_saida_relatorio = diretorio_atual / "saida" / "relatorio_presencas"
    nome_arq_rel = f"Relatório de Presença - {nome_seguro_curso}.xlsx"
    if len(str(pasta_saida_relatorio / nome_arq_rel)) > 210 or len(nome_seguro_curso) > 40:
        curso_curto = reduzir_nome_curso_para_arquivo(config["NOME_DO_CURSO"], max_chars=35)
        nome_arq_rel = f"Relatório de Presença - {curso_curto}.xlsx"
    caminho_saida = pasta_saida_relatorio / nome_arq_rel

    try:
        esvaziar_pasta(pasta_saida_relatorio, (".xlsx",))
        df_final = gerar_relatorio(
            caminho_servidores=caminho_servidores,
            caminho_convocados=caminho_convocados,
            arquivos_presenca=arquivos_presenca,
            caminho_saida_excel=caminho_saida,
            carga_horaria_horas=config["CARGA_HORARIA"],
            tempo_comparecimento_minutos=config["TEMPO_MINIMO_PRESENCA_MINUTOS"],
            percentual_minimo_carga=config["PERCENTUAL_MINIMO_CARGA"],
            percentual_minimo_encontros=config.get("PERCENTUAL_MINIMO_ENCONTROS", config.get("PERCENTUAL_MINIMO_DIAS", 0.75)),
            config=config,
        )

        limpar_tela()
        exibir_cabecalho("RELATÓRIO GERADO COM SUCESSO!", "Etapa 2 Concluída")

        total_participantes = len(df_final) if df_final is not None else 0
        if df_final is not None and "Apto por CH" in df_final.columns:
            total_aptos = int(df_final["Apto por CH"].astype(str).str.strip().str.upper().eq("APTO").sum())
        elif df_final is not None and "Apto ao Certificado" in df_final.columns:
            total_aptos = int(df_final["Apto ao Certificado"].astype(str).str.strip().str.upper().isin(["SIM", "APTO"]).sum())
        elif df_final is not None:
            colunas_apto = [c for c in df_final.columns if "apto" in c.lower()]
            if colunas_apto:
                total_aptos = int(df_final[colunas_apto[0]].astype(str).str.strip().str.upper().isin(["APTO", "SIM"]).sum())
            else:
                total_aptos = 0
        else:
            total_aptos = 0

        print(f"  {ciano(negrito('📊 RESUMO DA AUDITORIA DE PRESENÇA:'))}")
        print(f"     • Total de participantes convocados: {negrito(str(total_participantes))}")
        print(f"     • Total de participantes aptos:      {verde(negrito(str(total_aptos)))}")
        if total_participantes > total_aptos:
            print(f"     • Não aptos / desistentes:           {amarelo(str(total_participantes - total_aptos))}")

        print(f"\n  {ciano(negrito('📁 ARQUIVO GERADO:'))}")
        print(f"     {verde(str(caminho_saida))}")
        print(f"\n  {dim('A planilha contém as abas: Certificação, Presenças, Interações por participante, Interações por encontro, Indicadores, Custos e Informações do Curso.')}")

        print(f"\n  {ciano(negrito('▶ PRÓXIMO PASSO:'))}")
        print(f"  Execute {verde(negrito('gerar_certificados.exe'))} para emitir os certificados dos participantes aptos.")

        pausar_e_sair(0, "Pressione [ENTER] para fechar esta janela...")
    except PermissionError:
        print()
        exibir_erro(
            "Permissão negada ao salvar a planilha de relatório.",
            detalhe=f"O arquivo '{caminho_saida.name}' já está aberto no Microsoft Excel.",
            dica="Feche a janela do Excel com a planilha aberta e execute o programa novamente."
        )
        pausar_e_sair(1)
    except FileNotFoundError as e:
        print()
        exibir_erro(
            f"Erro ao salvar arquivo de relatório: {e}",
            detalhe=f"Caminho: {caminho_saida}",
            dica="O caminho ultrapassou o limite do Windows (260 caracteres). Mova a pasta do curso para um local mais curto (ex: C:\\Cursos\\)."
        )
        traceback.print_exc()
        pausar_e_sair(1)
    except Exception as e:
        print()
        exibir_erro(f"Erro durante a geração do relatório: {e}")
        traceback.print_exc()
        pausar_e_sair(1)


if __name__ == "__main__":
    main()
