"""
Suite de testes automatizados para validação do Gerador de Certificados.
"""

import unittest
from pathlib import Path
import shutil
import tempfile
import pandas as pd
from datetime import datetime
from pypdf import PdfWriter

import sys
diretorio_raiz = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(diretorio_raiz))

from core.config import carregar_variaveis, validar_variaveis, ConfigValidationError
from core.validator import validar_estrutura_curso, ValidationError, detectar_arquivos_temporarios
from core.certificados import (
    periodo_por_extenso,
    data_por_extenso,
    converter_data,
    nome_seguro_arquivo,
    juntar_certificados_pdf,
    carregar_relatorio_presenca,
)
from core.emails import criar_corpo_email
from core.presenca import (
    tempo_em_segundos,
    segundos_em_tempo,
    segundos_em_formato_digital,
    segundos_em_formato_hh_mm_ss,
    normalizar_texto,
    sanitizar_nome_arquivo,
    localizar_coluna_uorg,
    juntar_servidor_presenca,
    calcular_dias_minimos,
    calcular_encontros_minimos,
    criar_dados_do_encontro,
    adicionar_encontros_presenca,
    criar_aba_presencas,
    organizar_colunas,
    adicionar_aptidao,
    ler_presenca,
    preparar_convocados,
    processar_encontro,
    encurtar_caminho_se_necessario,
    converter_para_inteiro,
    extrair_datetime_resumo,
    extrair_data_resumo,
    ler_resumo_reuniao,
    ordenar_arquivos_presenca,
    identificar_colunas_metricas,
    processar_interacoes_participantes,
    criar_aba_interacoes_participante,
    criar_aba_interacoes_encontro,
    criar_aba_indicadores,
    criar_aba_custos,
    formatar_aba_custos,
    normalizar_email,
    filtrar_emails_ignorados,
    exportar_excel,
    formatar_workbook,
)
from core.caminhos import (
    caminho_estendido,
    garantir_diretorio,
    abrir_arquivo,
    copiar_arquivo_com_fallback,
    existe_com_fallback,
    listar_diretorio_com_fallback,
    ler_excel_com_fallback,
    verificar_profundidade_caminho,
    obter_mensagem_orientacao_caminho_profundo,
    LIMITE_CARACTERES_DIRETORIO_CURSO,
    LIMITE_CARACTERES_CAMINHO_FINAL,
    esvaziar_pasta,
)


class TestGeradorCertificados(unittest.TestCase):

    def setUp(self):
        self.temp_dir = Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_normalizar_texto(self):
        self.assertEqual(normalizar_texto("  João D'Ávila  "), "JOAO D AVILA")
        self.assertEqual(normalizar_texto("Ana Carolina  Soares"), "ANA CAROLINA SOARES")

    def test_conversao_tempo(self):
        self.assertEqual(tempo_em_segundos("01:30:00"), 5400)
        self.assertEqual(tempo_em_segundos("45:15"), 2715)
        self.assertEqual(tempo_em_segundos("1h 20m 10s"), 4810)
        self.assertEqual(segundos_em_formato_digital(3665), "1:01:05")
        self.assertEqual(segundos_em_tempo(3665), "1h 1m 5s")

    def test_periodo_por_extenso(self):
        tipo, texto = periodo_por_extenso("24/08/2026", "27/08/2026")
        self.assertEqual(tipo, "período")
        self.assertEqual(texto, "24 a 27 de agosto de 2026")

        tipo_dia, texto_dia = periodo_por_extenso("24/08/2026", None)
        self.assertEqual(tipo_dia, "dia")
        self.assertEqual(texto_dia, "24 de agosto de 2026")

    def test_nome_seguro_arquivo(self):
        self.assertEqual(nome_seguro_arquivo('José da Silva / "Diretor"'), "JOSÉ_DA_SILVA_DIRETOR")

    def test_sanitizar_nome_arquivo(self):
        # Caso com dois-pontos (exatamente o do erro reportado)
        self.assertEqual(
            sanitizar_nome_arquivo("Ambientação TAM Módulo: Transparência e Ética: Corregedoria"),
            "Ambientação TAM Módulo - Transparência e Ética - Corregedoria",
        )
        # Casos com outros caracteres proibidos no Windows (/ \ * ? " < > |)
        self.assertEqual(sanitizar_nome_arquivo('Curso "Gestão" / Módulo 1'), "Curso Gestão - Módulo 1")
        self.assertEqual(sanitizar_nome_arquivo("Como Liderar? Princípios: Guia"), "Como Liderar Princípios - Guia")
        self.assertEqual(sanitizar_nome_arquivo("Área | Subárea <Principal>"), "Área - Subárea Principal")
        self.assertEqual(sanitizar_nome_arquivo("  Curso com pontos e espaços...  "), "Curso com pontos e espaços")
        self.assertEqual(sanitizar_nome_arquivo(":::"), "arquivo")
        self.assertEqual(sanitizar_nome_arquivo(""), "arquivo")

    def test_calcular_dias_minimos(self):
        # Curso de 1 encontro com 75% deve exigir 1 encontro de presença (não 0!)
        self.assertEqual(calcular_encontros_minimos(1, 0.75), 1)
        self.assertEqual(calcular_dias_minimos(1, 0.75), 1)
        # Curso de 2 encontros com 75% exige 2 encontros (1.5 arredonda para cima)
        self.assertEqual(calcular_encontros_minimos(2, 0.75), 2)
        # Curso de 4 encontros com 75% exige 3 encontros
        self.assertEqual(calcular_encontros_minimos(4, 0.75), 3)
        # Curso de 5 encontros com 75% exige 4 encontros (3.75 arredonda para cima)
        self.assertEqual(calcular_encontros_minimos(5, 0.75), 4)
        # Sem exigência de presença
        self.assertEqual(calcular_encontros_minimos(1, 0.0), 0)

    def test_aptidao_por_dias_curso_um_dia(self):
        # Simula relatório com 1 encontro de curso:
        # Pessoa A esteve presente (Presenças = 1)
        # Pessoa B esteve ausente (Presenças = 0)
        df_simulado = pd.DataFrame({
            "Nome": ["Pessoa A", "Pessoa B"],
            "Presenças": [1, 0],
        })
        carga_segundos = pd.Series([7200, 0])  # 2h e 0h
        resultado = adicionar_aptidao(
            relatorio=df_simulado,
            carga_segundos=carga_segundos,
            quantidade_dias=1,
            carga_horaria_horas=2.0,
            percentual_minimo_carga=0.75,
            percentual_minimo_dias=0.75,
        )

        apto_a = resultado[resultado["Nome"] == "Pessoa A"].iloc[0]
        apto_b = resultado[resultado["Nome"] == "Pessoa B"].iloc[0]

        # Pessoa A esteve presente no único encontro -> APTO
        self.assertEqual(apto_a["Apto por Presença"], "APTO")
        self.assertEqual(apto_a["Apto por Encontros"], "APTO")
        self.assertEqual(apto_a["Apto por Dias"], "APTO")
        self.assertEqual(apto_a["Apto em Ambos"], "APTO")

        # Pessoa B NÃO esteve presente -> NÃO APTO
        self.assertEqual(apto_b["Apto por Presença"], "NÃO APTO")
        self.assertEqual(apto_b["Apto por Encontros"], "NÃO APTO")
        self.assertEqual(apto_b["Apto por Dias"], "NÃO APTO")
        self.assertEqual(apto_b["Apto em Ambos"], "NÃO APTO")

    def test_colunas_encontros_e_organizar_colunas(self):
        df_res = pd.DataFrame({
            "Nome": ["Alice"],
            "Email": ["alice@anm.gov.br"],
            "Duração da Reunião": ["01:30:00"],
            "Compareceu": ["Presente"],
        })
        encontro1 = criar_dados_do_encontro(df_res, 1)
        self.assertIn("Encontro 1", encontro1.columns)
        self.assertIn("Comparecimento 1", encontro1.columns)

        convocados = pd.DataFrame({"Nome": ["Alice"], "UORG": ["SGP"], "Email": ["alice@anm.gov.br"], "Presenças": [1]})
        rel = adicionar_encontros_presenca(convocados, [encontro1])
        self.assertEqual(rel.iloc[0]["Encontro 1"], "1:30:00")
        self.assertEqual(rel.iloc[0]["Comparecimento 1"], "Presente")

        rel["CH Total"] = "01:30:00"
        rel["% da CH Total"] = 1.0
        rel["Apto por CH"] = "APTO"
        rel["Apto por Presença"] = "APTO"
        rel["Apto em Ambos"] = "APTO"

        organizado = organizar_colunas(rel, 1)
        self.assertEqual(
            list(organizado.columns),
            ["Nome", "UORG", "Email", "Presenças", "CH Total", "% da CH Total", "Apto por CH", "Apto por Presença", "Apto em Ambos"]
        )
        self.assertNotIn("Encontro 1", organizado.columns)
        self.assertNotIn("Comparecimento 1", organizado.columns)

        aba_presencas = criar_aba_presencas(rel, 1)
        self.assertEqual(
            list(aba_presencas.columns),
            ["Nome", "Encontro 1", "Comparecimento 1"]
        )

    def test_parser_variaveis_valido(self):
        caminho_var = self.temp_dir / "variaveis.txt"
        conteudo = """
        NOME_DO_CURSO = Curso Teste ANM
        CARGA_HORARIA = 16
        DATA_INICIO = 10/08/2026
        DATA_FIM = 14/08/2026
        LOCAL_EMISSAO = Brasília — DF
        DATA_EMISSAO = 14/08/2026
        PERCENTUAL_MINIMO_CARGA = 75%
        PERCENTUAL_MINIMO_ENCONTROS = 0.75
        TEMPO_MINIMO_PRESENCA_MINUTOS = 30
        ASSUNTO_EMAIL = Certificado - {CURSO}

        [INICIO_CORPO_EMAIL]
        Olá {nome}, segue certificado do curso {CURSO}.
        [FIM_CORPO_EMAIL]
        """
        caminho_var.write_text(conteudo, encoding="utf-8")
        raw = carregar_variaveis(caminho_var)
        config = validar_variaveis(raw)

        self.assertEqual(config["NOME_DO_CURSO"], "Curso Teste ANM")
        self.assertEqual(config["CARGA_HORARIA"], 16)
        self.assertEqual(config["PERCENTUAL_MINIMO_CARGA"], 0.75)
        self.assertEqual(config["PERCENTUAL_MINIMO_ENCONTROS"], 0.75)
        self.assertIn("Olá {nome}", config["CORPO_EMAIL"])

    def test_corpo_email_sem_comentarios_indesejados(self):
        # 1. Testa com o arquivo de template oficial (que possui explicações e divisórias # ---)
        caminho_template = diretorio_raiz / "templates" / "variaveis_template.txt"
        config_template = carregar_variaveis(caminho_template)
        corpo = config_template["CORPO_EMAIL"]
        self.assertFalse(corpo.startswith("#"))
        self.assertTrue(corpo.startswith("Prezada(o) {nome},"))

        # 2. Testa caso o usuário tenha deixado um comentário acidental dentro de [INICIO_CORPO_EMAIL]
        caminho_var = self.temp_dir / "variaveis_com_comentario.txt"
        conteudo = """
        NOME_DO_CURSO = Curso Teste
        CARGA_HORARIA = 10
        DATA_INICIO = 10/08/2026
        LOCAL_EMISSAO = Brasília
        DATA_EMISSAO = 10/08/2026
        [INICIO_CORPO_EMAIL]
        # ------------------------------------------------------------
        # CORPO DO E-MAIL
        Prezada(o) {nome}, seu certificado de {CURSO} está pronto.
        [FIM_CORPO_EMAIL]
        """
        caminho_var.write_text(conteudo, encoding="utf-8")
        raw = carregar_variaveis(caminho_var)
        self.assertTrue(raw["CORPO_EMAIL"].startswith("Prezada(o) {nome},"))
        self.assertNotIn("# ------------------------------------------------------------", raw["CORPO_EMAIL"])

        # 3. Testa função criar_corpo_email diretamente com comentário inicial
        texto_final = criar_corpo_email(
            nome="João",
            template_corpo="# ------------------------------------------------------------\nOlá {nome}, curso {CURSO}!",
            nome_curso="Direito",
        )
        self.assertEqual(texto_final, "Olá João, curso Direito!")

    def test_parser_variaveis_erros(self):
        raw_invalido = {
            "NOME_DO_CURSO": "",
            "CARGA_HORARIA": "abc",
            "DATA_INICIO": "data_errada",
        }
        with self.assertRaises(ConfigValidationError) as ctx:
            validar_variaveis(raw_invalido)
        erros = ctx.exception.errors
        self.assertTrue(any("NOME_DO_CURSO" in e for e in erros))
        self.assertTrue(any("CARGA_HORARIA" in e for e in erros))
        self.assertTrue(any("DATA_INICIO" in e for e in erros))

    def test_validador_pastas_faltantes(self):
        # Pasta vazia sem estrutura
        with self.assertRaises(ValidationError) as ctx:
            validar_estrutura_curso(self.temp_dir)
        self.assertTrue(any("variaveis.txt" in e for e in ctx.exception.erros))
        self.assertTrue(any("lista_convidados" in e for e in ctx.exception.erros))

    def test_validador_detecta_arquivos_office_temporarios(self):
        temp_office = self.temp_dir / "~$planilha.xlsx"
        temp_office.write_text("dummy")
        temps = detectar_arquivos_temporarios(self.temp_dir)
        self.assertEqual(len(temps), 1)
        self.assertEqual(temps[0].name, "~$planilha.xlsx")

    def test_juntar_pdfs_separado(self):
        pasta_pdf = self.temp_dir / "pdfs"
        pasta_pdf.mkdir()

        # Cria 2 PDFs mínimos válidos com pypdf
        for i in range(1, 3):
            writer = PdfWriter()
            writer.add_blank_page(width=200, height=200)
            with open(pasta_pdf / f"cert_{i}.pdf", "wb") as f:
                writer.write(f)

        saida_junto = self.temp_dir / "juntos" / "todos.pdf"
        res = juntar_certificados_pdf(pasta_pdf, saida_junto)
        self.assertTrue(res.exists())
        self.assertGreater(res.stat().st_size, 0)

    def test_juntar_pdfs_com_caracteres_especiais(self):
        pasta_pdf = self.temp_dir / "pdfs_esp"
        pasta_pdf.mkdir()

        for i in range(1, 3):
            writer = PdfWriter()
            writer.add_blank_page(width=200, height=200)
            with open(pasta_pdf / f"cert_{i}.pdf", "wb") as f:
                writer.write(f)

        saida_invalida = self.temp_dir / "juntos_esp" / "certificados_Ambientação: Transparência: Corregedoria.pdf"
        res = juntar_certificados_pdf(pasta_pdf, saida_invalida)
        self.assertTrue(res.exists())
        self.assertEqual(res.name, "certificados_Ambientação - Transparência - Corregedoria.pdf")

    def test_localizar_coluna_uorg(self):
        # Quando 'COD UORG' vem antes de 'UORG'
        colunas1 = ["COD UORG", "UORG", "NOME SERVIDOR"]
        self.assertEqual(localizar_coluna_uorg(colunas1), "UORG")

        # Quando case varia ou há espaços extras
        colunas2 = ["COD_UORG", "  uorg  ", "NOME"]
        self.assertEqual(localizar_coluna_uorg(colunas2), "  uorg  ")

        # Quando não tem UORG, mas tem LOTACAO
        colunas3 = ["COD UORG", "LOTACAO", "SERVIDOR"]
        self.assertEqual(localizar_coluna_uorg(colunas3), "LOTACAO")

    def test_juntar_servidor_com_cod_uorg_e_uorg(self):
        df_servidores = pd.DataFrame({
            "COD UORG": ["12345", "67890"],
            "UORG": ["SGP/ANM", "DITEC/ANM"],
            "NOME SERVIDOR": ["ALICE SILVA", "BRUNO SOUZA"],
        })
        df_presenca = pd.DataFrame({
            "Nome": ["Alice Silva", "Bruno Souza"],
            "Duração da Reunião": ["01:00:00", "02:00:00"],
            "Email": ["alice@anm.gov.br", "bruno@anm.gov.br"],
        })

        resultado = juntar_servidor_presenca(df_servidores, df_presenca)
        # Garante que pegou 'SGP/ANM' e não '12345'
        self.assertListEqual(resultado["UORG"].tolist(), ["SGP/ANM", "DITEC/ANM"])

    def test_ler_presenca_linhas_variaveis_antes_participantes(self):
        # Caso 1: Apenas 3 linhas antes de "2. Participantes" (menos que 9)
        linhas_curtas = [
            "1. Resumo da Reunião",
            "Total de participantes\t2",
            "2. Participantes",
            "Nome\tEmail\tDuração da Reunião",
            "Alice Silva\talice@anm.gov.br\t2h 00m",
            "Bruno Souza\tbruno@anm.gov.br\t1h 30m",
            "3. Atividades em Reunião",
            "Outro Dado\tValor",
        ]
        arq_curto = self.temp_dir / "presenca_curta.csv"
        arq_curto.write_text("\n".join(linhas_curtas), encoding="utf-16")

        df_curto = ler_presenca(arq_curto)
        self.assertEqual(len(df_curto), 2)
        self.assertListEqual(df_curto["Nome"].tolist(), ["Alice Silva", "Bruno Souza"])

        # Caso 2: 15 linhas antes de "2. Participantes" (mais que 9)
        linhas_longas = [f"Metadado {i}\tValor {i}" for i in range(1, 16)]
        linhas_longas.append("2. Participantes")
        linhas_longas.append("Nome\tEmail\tDuração da Reunião")
        linhas_longas.append("Carlos Pereira\tcarlos@anm.gov.br\t3h 00m")
        linhas_longas.append("Daniela Lima\tdaniela@anm.gov.br\t3h 00m")
        linhas_longas.append("3. Atividades em Reunião")
        linhas_longas.append("Info\tX")
        arq_longo = self.temp_dir / "presenca_longa.csv"
        arq_longo.write_text("\n".join(linhas_longas), encoding="utf-8")

        df_longo = ler_presenca(arq_longo)
        self.assertEqual(len(df_longo), 2)
        self.assertListEqual(df_longo["Nome"].tolist(), ["Carlos Pereira", "Daniela Lima"])

    def test_ler_presenca_formatos_variados(self):
        # Linha com aspas e espaços: "2. Participantes" e separador vírgula
        conteudo = (
            '"1. Resumo da Reunião"\n'
            '"Info","Val"\n'
            '"2. Participantes"\n'
            'Nome,Email,Duração da Reunião\n'
            'Eduardo Rocha,eduardo@anm.gov.br,2h\n'
        )
        arq_aspas = self.temp_dir / "presenca_aspas.csv"
        arq_aspas.write_text(conteudo, encoding="utf-8")

        df_aspas = ler_presenca(arq_aspas)
        self.assertEqual(len(df_aspas), 1)
        self.assertEqual(df_aspas.iloc[0]["Nome"], "Eduardo Rocha")

    def test_ler_presenca_fallback_sem_skiprows(self):
        # Arquivo CSV direto, sem cabeçalhos do Teams (começa direto na coluna Nome)
        conteudo = "Nome,Email,Duração da Reunião\nFernanda Costa,fernanda@anm.gov.br,1h\n"
        arq_direto = self.temp_dir / "presenca_direto.csv"
        arq_direto.write_text(conteudo, encoding="utf-8")

        df_direto = ler_presenca(arq_direto)
        self.assertEqual(len(df_direto), 1)
        self.assertEqual(df_direto.iloc[0]["Nome"], "Fernanda Costa")

    def test_carregar_relatorio_presenca_sucesso_e_abas(self):
        # Arquivo com aba 'Certificação'
        caminho = self.temp_dir / "relatorio_certificacao.xlsx"
        df_dados = pd.DataFrame({
            "Nome": ["Alice Silva", "Bruno Souza"],
            "Apto por CH": ["APTO", "NÃO APTO"],
        })
        with pd.ExcelWriter(caminho) as writer:
            df_dados.to_excel(writer, sheet_name="Certificação", index=False)

        carregado = carregar_relatorio_presenca(caminho)
        self.assertEqual(len(carregado), 2)
        self.assertIn("Apto por CH", carregado.columns)

        # Arquivo com aba legada 'Relatório Geral'
        caminho_legado = self.temp_dir / "relatorio_legado.xlsx"
        with pd.ExcelWriter(caminho_legado) as writer:
            df_dados.to_excel(writer, sheet_name="Relatório Geral", index=False)
        carregado_legado = carregar_relatorio_presenca(caminho_legado)
        self.assertEqual(len(carregado_legado), 2)

        # Arquivo sem a aba nomeada (fallback para a primeira aba)
        caminho_sem_aba = self.temp_dir / "relatorio_sem_aba.xlsx"
        df_dados.to_excel(caminho_sem_aba, sheet_name="OutraAba", index=False)
        carregado_fallback = carregar_relatorio_presenca(caminho_sem_aba)
        self.assertEqual(len(carregado_fallback), 2)

    def test_carregar_relatorio_presenca_colunas_faltantes(self):
        caminho = self.temp_dir / "relatorio_incompleto.xlsx"
        df_incompleto = pd.DataFrame({"Nome": ["Alice"]})
        df_incompleto.to_excel(caminho, index=False)
        with self.assertRaises(ValueError) as ctx:
            carregar_relatorio_presenca(caminho)
        self.assertIn("Apto por CH", str(ctx.exception))

    def test_encurtar_caminho_se_necessario(self):
        # Caminho curto não deve sofrer alteração
        caminho_curto = Path("C:/Cursos/curso1/saida/relatorio_presencas/Relatorio.xlsx")
        self.assertEqual(encurtar_caminho_se_necessario(caminho_curto, max_len=245), caminho_curto)

        # Caminho longo deve ser encurtado mantendo o sufixo e respeitando max_len
        pasta_longa = Path("C:/" + "a" * 150 + "/saida/relatorio_presencas")
        nome_longo = "Relatório de Presença - " + "X" * 80 + ".xlsx"
        caminho_longo = pasta_longa / nome_longo
        encurtado = encurtar_caminho_se_necessario(caminho_longo, max_len=220)
        self.assertLessEqual(len(str(encurtado)), 220)
        self.assertTrue(encurtado.name.endswith(".xlsx"))
        self.assertTrue(encurtado.name.startswith("Relatório de Presença - "))

    def test_caminho_estendido_formatacao(self):
        import os
        caminho_simples = "C:/pastateste/arquivo.txt"
        res = caminho_estendido(caminho_simples)
        if os.name == "nt":
            self.assertTrue(res.startswith("\\\\?\\"))
            self.assertNotIn("/", res[4:])
            # Idempotência
            self.assertEqual(caminho_estendido(res), res)

    def test_garantir_diretorio_e_abrir_arquivo(self):
        nova_pasta = self.temp_dir / "sub1" / "sub2" / "sub3"
        garantir_diretorio(nova_pasta)
        self.assertTrue(existe_com_fallback(nova_pasta))

        arq_teste = nova_pasta / "teste_escrita.txt"
        with abrir_arquivo(arq_teste, "w", encoding="utf-8") as f:
            f.write("Conteúdo gravado com sucesso!")

        self.assertTrue(existe_com_fallback(arq_teste))

        with abrir_arquivo(arq_teste, "r", encoding="utf-8") as f:
            conteudo = f.read()
        self.assertEqual(conteudo, "Conteúdo gravado com sucesso!")

    def test_copiar_arquivo_com_fallback(self):
        origem = self.temp_dir / "origem.txt"
        with abrir_arquivo(origem, "w", encoding="utf-8") as f:
            f.write("Arquivo original")

        destino = self.temp_dir / "copia" / "destino.txt"
        garantir_diretorio(destino.parent)
        copiar_arquivo_com_fallback(origem, destino)

        self.assertTrue(existe_com_fallback(destino))
        with abrir_arquivo(destino, "r", encoding="utf-8") as f:
            self.assertEqual(f.read(), "Arquivo original")

    def test_caminho_longo_estendido(self):
        # Cria estrutura de pastas aninhada para testar caminhos longos (>260 caracteres)
        pasta_profunda = self.temp_dir
        for i in range(5):
            pasta_profunda = pasta_profunda / f"pasta_com_nome_longo_nivel_{i}_{'x' * 25}"

        garantir_diretorio(pasta_profunda)
        self.assertTrue(existe_com_fallback(pasta_profunda))

        arq_profundo = pasta_profunda / "variaveis.txt"
        with abrir_arquivo(arq_profundo, "w", encoding="utf-8") as f:
            f.write("NOME_DO_CURSO = Teste Longo\nCARGA_HORARIA = 20\n")

        self.assertTrue(existe_com_fallback(arq_profundo))
        with abrir_arquivo(arq_profundo, "r", encoding="utf-8") as f:
            self.assertIn("NOME_DO_CURSO", f.read())

    def test_verificar_profundidade_caminho(self):
        caminho_curto = "C:/Cursos/curso_anm"
        eh_profundo, tam, msg = verificar_profundidade_caminho(caminho_curto, limite=130)
        self.assertFalse(eh_profundo)
        self.assertEqual(msg, "")

        caminho_longo = "C:/" + "subpasta_muito_profunda/" * 10
        eh_profundo_l, tam_l, msg_l = verificar_profundidade_caminho(caminho_longo, limite=130)
        self.assertTrue(eh_profundo_l)
        self.assertGreater(tam_l, 130)
        self.assertIn("ORIENTAÇÃO", msg_l)
        self.assertIn("C:\\Cursos\\", msg_l)

    def test_validador_bloqueia_pasta_profunda(self):
        # Cria uma pasta com caminho longo (>130 caracteres)
        sub_longa = self.temp_dir / ("a" * 80) / ("b" * 60)
        sub_longa.mkdir(parents=True, exist_ok=True)
        self.assertGreater(len(str(sub_longa.resolve())), LIMITE_CARACTERES_DIRETORIO_CURSO)

        with self.assertRaises(ValidationError) as ctx:
            validar_estrutura_curso(sub_longa)
        
        erro_str = str(ctx.exception)
        self.assertIn("DIRETÓRIO MUITO PROFUNDO", erro_str)
        self.assertIn("ORIENTAÇÃO", erro_str)
        self.assertIn("C:\\Cursos\\", erro_str)

    def test_ausencia_de_tempfile_e_temp_em_modulos_core(self):
        import inspect
        import core.certificados as cert_mod
        import core.emails as email_mod

        # Verifica se tempfile não foi importado pelos módulos
        self.assertNotIn("tempfile", dir(cert_mod), "core.certificados não deve importar tempfile")
        self.assertNotIn("tempfile", dir(email_mod), "core.emails não deve importar tempfile")

        # Verifica se o código-fonte de ambos não contém chamadas a gettempdir()
        codigo_cert = inspect.getsource(cert_mod)
        codigo_email = inspect.getsource(email_mod)

        self.assertNotIn("gettempdir", codigo_cert, "core.certificados não deve chamar gettempdir")
        self.assertNotIn("gettempdir", codigo_email, "core.emails não deve chamar gettempdir")
        self.assertNotIn("%temp%", codigo_email.lower(), "core.emails não deve referenciar %temp%")

    def test_gerar_todos_certificados_detecta_caminho_excessivo_preventivamente(self):
        from core.certificados import gerar_todos_certificados

        df_aptos = pd.DataFrame([{"Nome": "João da Silva", "Email": "joao@anm.gov.br"}])
        modelo = self.temp_dir / "modelo.pptx"
        modelo.touch()

        # Cria caminhos onde a pasta do PDF é excessivamente longa (> 230 chars ao somar o nome do arquivo)
        pasta_excessiva = self.temp_dir / ("sub_" * 50)
        
        with self.assertRaises(ValueError) as ctx:
            gerar_todos_certificados(
                df_aptos=df_aptos,
                modelo_powerpoint=modelo,
                pasta_ppt=pasta_excessiva / "ppt",
                pasta_pdf=pasta_excessiva / "pdf",
                pasta_logs=pasta_excessiva / "logs",
                config={"NOME_DO_CURSO": "Curso Teste"},
            )

        self.assertIn("ORIENTAÇÃO", str(ctx.exception))
        self.assertIn("C:\\Cursos\\", str(ctx.exception))

    def test_esvaziar_pasta_completo_e_filtrado(self):
        pasta = self.temp_dir / "teste_esvaziar"
        pasta.mkdir()

        (pasta / "cert_1.pdf").write_text("dummy pdf 1")
        (pasta / "cert_2.pdf").write_text("dummy pdf 2")
        (pasta / "apresentacao.pptx").write_text("dummy pptx")
        (pasta / "anotacoes.txt").write_text("dummy txt")

        # 1. Limpa apenas .pdf
        removidos_pdf = esvaziar_pasta(pasta, (".pdf",))
        self.assertEqual(removidos_pdf, 2)
        self.assertFalse((pasta / "cert_1.pdf").exists())
        self.assertFalse((pasta / "cert_2.pdf").exists())
        self.assertTrue((pasta / "apresentacao.pptx").exists())
        self.assertTrue((pasta / "anotacoes.txt").exists())

        # 2. Limpa tudo (sem filtro de extensões)
        removidos_todos = esvaziar_pasta(pasta)
        self.assertEqual(removidos_todos, 2)
        self.assertFalse((pasta / "apresentacao.pptx").exists())
        self.assertFalse((pasta / "anotacoes.txt").exists())
        self.assertEqual(len(list(pasta.iterdir())), 0)

    def test_esvaziar_pasta_previne_certificado_usuario_removido(self):
        # Simula pasta com certificados anteriores
        pasta_pdf = self.temp_dir / "saida_certificados_pdf"
        pasta_pdf.mkdir()

        # Certificado antigo de usuário que foi removido do relatório
        cert_antigo = pasta_pdf / "CERTIFICADO_USUARIO_REMOVIDO.pdf"
        cert_antigo.write_text("conteudo antigo")
        self.assertTrue(cert_antigo.exists())

        # Ao gerar novos arquivos, esvaziamos a pasta
        removidos = esvaziar_pasta(pasta_pdf, (".pdf",))
        self.assertEqual(removidos, 1)
        self.assertFalse(cert_antigo.exists())

        # Agora cria os certificados dos participantes atualmente aptos
        writer = PdfWriter()
        writer.add_blank_page(width=200, height=200)
        cert_atual = pasta_pdf / "CERTIFICADO_USUARIO_ATUAL.pdf"
        with open(cert_atual, "wb") as f:
            writer.write(f)

        # Junta os certificados e verifica que apenas o atual foi incluído
        saida_junto = self.temp_dir / "saida_juntos" / "todos.pdf"
        juntar_certificados_pdf(pasta_pdf, saida_junto)
        self.assertTrue(saida_junto.exists())
        
        # Garante que na pasta só existe o arquivo atual
        arquivos_na_pasta = [f.name for f in pasta_pdf.iterdir()]
        self.assertNotIn("CERTIFICADO_USUARIO_REMOVIDO.pdf", arquivos_na_pasta)
        self.assertIn("CERTIFICADO_USUARIO_ATUAL.pdf", arquivos_na_pasta)

    def test_validador_subpastas_saida_novas_pastas(self):
        from core.validator import SUBPASTAS_SAIDA
        self.assertIn("certificados_juntos_pdf", SUBPASTAS_SAIDA)
        self.assertIn("certificados_juntos_ppt", SUBPASTAS_SAIDA)
        self.assertIn("certificados_pdf", SUBPASTAS_SAIDA)
        self.assertNotIn("certificados_ppt", SUBPASTAS_SAIDA)
        self.assertNotIn("certificados_juntos", SUBPASTAS_SAIDA)

    def test_gerar_powerpoint_consolidado_multiplos_slides(self):
        from pptx import Presentation
        from pptx.util import Inches
        from core.certificados import gerar_powerpoint_consolidado

        # 1. Cria modelo com 2 slides (frente e verso)
        caminho_modelo = self.temp_dir / "modelo_teste.pptx"
        prs = Presentation()
        # Slide 1 (Frente)
        slide1 = prs.slides.add_slide(prs.slide_layouts[6])
        tx_box1 = slide1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        tx_box1.text_frame.text = "Certificado de <<NOME>> no curso <<CURSO>>"

        # Slide 2 (Verso / Ementa)
        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        tx_box2 = slide2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        tx_box2.text_frame.text = "Ementa de <<CURSO>>: Concluinte <<NOME>>"
        prs.save(caminho_modelo)

        # 2. DataFrame de aptos com 2 pessoas
        df_aptos = pd.DataFrame([
            {"Nome": "Alice Silva", "Email": "alice@anm.gov.br"},
            {"Nome": "Bruno Souza", "Email": "bruno@anm.gov.br"},
        ])

        config = {
            "NOME_DO_CURSO": "Segurança de Barragens",
            "CARGA_HORARIA": "20",
        }

        saida_ppt_junto = self.temp_dir / "saida" / "certificados_juntos_ppt" / "todos.pptx"
        res = gerar_powerpoint_consolidado(
            df_aptos=df_aptos,
            modelo_powerpoint=caminho_modelo,
            caminho_ppt_junto=saida_ppt_junto,
            config=config,
        )

        self.assertTrue(res.exists())
        self.assertGreater(res.stat().st_size, 0)

        # 3. Valida a apresentação consolidada gerada
        prs_resultado = Presentation(saida_ppt_junto)
        # 2 participantes x 2 slides = 4 slides
        self.assertEqual(len(prs_resultado.slides), 4)

        # Slide 1: Alice Frente
        texto_s1 = prs_resultado.slides[0].shapes[0].text_frame.text
        self.assertIn("Alice Silva", texto_s1)
        self.assertIn("Segurança de Barragens", texto_s1)

        # Slide 2: Alice Verso
        texto_s2 = prs_resultado.slides[1].shapes[0].text_frame.text
        self.assertIn("Alice Silva", texto_s2)

        # Slide 3: Bruno Frente
        texto_s3 = prs_resultado.slides[2].shapes[0].text_frame.text
        self.assertIn("Bruno Souza", texto_s3)
        self.assertIn("Segurança de Barragens", texto_s3)

        # Slide 4: Bruno Verso
        texto_s4 = prs_resultado.slides[3].shapes[0].text_frame.text
        self.assertIn("Bruno Souza", texto_s4)

    def test_gerar_todos_certificados_e2e_sem_ppt_individual(self):
        from pptx import Presentation
        from pptx.util import Inches
        from core.certificados import gerar_todos_certificados

        # Cria modelo com 1 slide
        caminho_modelo = self.temp_dir / "modelo_e2e.pptx"
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        tx_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
        tx_box.text_frame.text = "Certificado de <<NOME>> - <<CURSO>>"
        prs.save(caminho_modelo)

        df_aptos = pd.DataFrame([
            {"Nome": "Carla Dias", "Email": "carla@anm.gov.br"},
            {"Nome": "Daniel Lima", "Email": "daniel@anm.gov.br"},
        ])

        pasta_pdf = self.temp_dir / "saida" / "certificados_pdf"
        pasta_logs = self.temp_dir / "saida" / "logs"
        pasta_ppt = self.temp_dir / "saida" / "certificados_ppt"
        saida_ppt_junto = self.temp_dir / "saida" / "certificados_juntos_ppt" / "todos.pptx"

        config = {"NOME_DO_CURSO": "Fiscalização de Minas"}

        df_res = gerar_todos_certificados(
            df_aptos=df_aptos,
            modelo_powerpoint=caminho_modelo,
            pasta_ppt=pasta_ppt,
            pasta_pdf=pasta_pdf,
            pasta_logs=pasta_logs,
            config=config,
            caminho_ppt_junto=saida_ppt_junto,
            salvar_ppt_individual=False,
        )

        self.assertEqual(len(df_res), 2)

        # 1. Verifica que os PDFs individuais foram gerados com o prefixo compacto CERT_
        self.assertTrue((pasta_pdf / "CERT_CARLA_DIAS.pdf").exists())
        self.assertTrue((pasta_pdf / "CERT_DANIEL_LIMA.pdf").exists())

        # 2. Verifica que NÃO foram criados arquivos PPT individuais
        if pasta_ppt.exists():
            ppts_individuais = list(pasta_ppt.glob("*.pptx"))
            self.assertEqual(len(ppts_individuais), 0)

        # 3. Verifica que o PowerPoint consolidado foi gerado com os 2 participantes
        self.assertTrue(saida_ppt_junto.exists())
        prs_resultado = Presentation(saida_ppt_junto)
        self.assertEqual(len(prs_resultado.slides), 2)
        self.assertIn("Carla Dias", prs_resultado.slides[0].shapes[0].text_frame.text)
        self.assertIn("Daniel Lima", prs_resultado.slides[1].shapes[0].text_frame.text)

    def test_import_enviar_emails_e_simbolos(self):
        import enviar_emails
        self.assertTrue(hasattr(enviar_emails, "obter_configuracao"))
        self.assertTrue(callable(enviar_emails.obter_configuracao))
        self.assertTrue(hasattr(enviar_emails, "vermelho"))
        self.assertTrue(callable(enviar_emails.vermelho))
        self.assertTrue(hasattr(enviar_emails, "encontrar_relatorio_presenca"))

    def test_reduzir_nome_curso_para_arquivo(self):
        from core.caminhos import reduzir_nome_curso_para_arquivo
        # Remove preposições e respeita o tamanho máximo
        nome_longo = "Hidrologia Aplicada à Segurança de Barragens e Pilhas de Mineração"
        res = reduzir_nome_curso_para_arquivo(nome_longo, max_chars=30)
        self.assertLessEqual(len(res), 30)
        self.assertIn("Hidrologia", res)
        self.assertIn("Seguranca", res)
        self.assertNotIn("de", res.split("_"))

    def test_obter_nome_arquivo_junto_evita_estouro(self):
        from core.caminhos import obter_nome_arquivo_junto
        # Simula pasta com 140 caracteres
        pasta_simulada = Path("C:\\" + "a" * 137)
        nome_longo = "Hidrologia Aplicada à Segurança de Barragens e Pilhas de Mineração"
        nome_arq = obter_nome_arquivo_junto(nome_longo, pasta_simulada, ".pptm", max_caminho_total=215)
        caminho_total = str(pasta_simulada / nome_arq)
        self.assertLessEqual(len(caminho_total), 215)
        self.assertTrue(nome_arq.startswith("cert_"))
        self.assertTrue(nome_arq.endswith("_todos_juntos.pptm"))

    def test_localizar_certificado_pdf_novo_e_legado(self):
        from core.emails import localizar_certificado_pdf
        pasta_pdf = self.temp_dir / "teste_localizar"
        pasta_pdf.mkdir()

        # 1. Deve encontrar o novo padrão CERT_
        (pasta_pdf / "CERT_MARCOS_SOUZA.pdf").write_text("pdf novo")
        res_novo = localizar_certificado_pdf(pasta_pdf, "Marcos Souza")
        self.assertEqual(res_novo.name, "CERT_MARCOS_SOUZA.pdf")

        # 2. Deve encontrar o padrão legado CERTIFICADO_ caso exista
        (pasta_pdf / "CERTIFICADO_BEATRIZ_LIMA.pdf").write_text("pdf legado")
        res_legado = localizar_certificado_pdf(pasta_pdf, "Beatriz Lima")
        self.assertEqual(res_legado.name, "CERTIFICADO_BEATRIZ_LIMA.pdf")


    def test_formatar_valor_configuracao(self):
        from core.presenca import formatar_valor_configuracao

        config = {
            "NOME_DO_CURSO": "Segurança de Barragens",
            "DATA_INICIO": "17/08/2026",
            "DATA_FIM": "",
        }

        # Carga horária
        self.assertEqual(formatar_valor_configuracao("CARGA_HORARIA", 20, config), "20 horas")
        self.assertEqual(formatar_valor_configuracao("CARGA_HORARIA", 1, config), "1 hora")
        self.assertEqual(formatar_valor_configuracao("CARGA_HORARIA", "16,5", config), "16.5 horas")

        # Percentuais
        self.assertEqual(formatar_valor_configuracao("PERCENTUAL_MINIMO_CARGA", 0.75, config), "75%")
        self.assertEqual(formatar_valor_configuracao("PERCENTUAL_MINIMO_ENCONTROS", "75%", config), "75%")
        self.assertEqual(formatar_valor_configuracao("PERCENTUAL_MINIMO_CARGA", 0.8, config), "80%")

        # Tempo de permanência
        self.assertEqual(formatar_valor_configuracao("TEMPO_MINIMO_PRESENCA_MINUTOS", 30, config), "30 minutos")
        self.assertEqual(formatar_valor_configuracao("TEMPO_MINIMO_PRESENCA_MINUTOS", 1, config), "1 minuto")

        # Data de fim vazia (curso de 1 dia)
        self.assertIn("Curso de 1 dia", formatar_valor_configuracao("DATA_FIM", "", config))

        # Assunto do email com substituição de {CURSO}
        self.assertEqual(
            formatar_valor_configuracao("ASSUNTO_EMAIL", "Certificado - Curso {CURSO}", config),
            "Certificado - Curso Segurança de Barragens"
        )

    def test_criar_aba_informacoes_curso_estrutura(self):
        from core.presenca import criar_aba_informacoes_curso

        config = {
            "NOME_DO_CURSO": "Ambientação ANM",
            "CARGA_HORARIA": 24,
            "DATA_INICIO": "01/09/2026",
            "DATA_FIM": "05/09/2026",
            "LOCAL_EMISSAO": "Brasília — DF",
            "DATA_EMISSAO": "05/09/2026",
            "PERCENTUAL_MINIMO_CARGA": 0.75,
            "PERCENTUAL_MINIMO_ENCONTROS": 0.75,
            "TEMPO_MINIMO_PRESENCA_MINUTOS": 45,
        }

        df_info = criar_aba_informacoes_curso(config)

        # 3 colunas amigáveis: Item, Configuração, Observação
        self.assertEqual(list(df_info.columns), ["Item", "Configuração", "Observação"])
        self.assertEqual(len(df_info), 9)

        # Todas as observações devem estar preenchidas com explicações
        for obs in df_info["Observação"]:
            self.assertTrue(len(obs) > 10)

        # Valores e itens devem estar devidamente formatados e amigáveis (sem parênteses com tags)
        for item in df_info["Item"]:
            self.assertNotIn("(", item)
            self.assertNotIn(")", item)
            self.assertNotIn("EMAIL", item.upper())

        mapa = dict(zip(df_info["Item"], df_info["Configuração"]))
        self.assertEqual(mapa["Carga Horária"], "24 horas")
        self.assertEqual(mapa["Tempo Mínimo de Permanência na Reunião"], "45 minutos")
        self.assertIn(mapa["Percentual Mínimo de Carga Horária"], (0.75, "75%"))

    def test_largura_dinamica_e_porcentagem_nativa_excel(self):
        import openpyxl
        from core.presenca import criar_aba_informacoes_curso, formatar_workbook

        config = {
            "NOME_DO_CURSO": "Curso com Nome Extremamente Longo para Teste de Largura da Coluna de Configuração no Excel",
            "CARGA_HORARIA": 40,
            "PERCENTUAL_MINIMO_CARGA": 0.75,
            "PERCENTUAL_MINIMO_ENCONTROS": "75%",
            "TEMPO_MINIMO_PRESENCA_MINUTOS": 30,
        }

        df_info = criar_aba_informacoes_curso(config)
        saida_teste = self.temp_dir / "teste_largura.xlsx"

        with pd.ExcelWriter(saida_teste, engine="openpyxl") as writer:
            df_info.to_excel(writer, sheet_name="Informações do Curso", index=False)
            formatar_workbook(writer.book)

        wb = openpyxl.load_workbook(saida_teste)
        ws = wb["Informações do Curso"]

        # 1. Verifica se a largura da coluna B adaptou ao nome longo do curso (muito maior que 35)
        largura_b = ws.column_dimensions["B"].width
        self.assertGreater(largura_b, 70)

        # 2. Verifica se as células de porcentagem foram gravadas como número float com formato nativo 0%
        # evitando o erro de "número armazenado como texto"
        for row in range(2, ws.max_row + 1):
            item = str(ws.cell(row=row, column=1).value or "")
            if "Percentual" in item:
                cel_conf = ws.cell(row=row, column=2)
                self.assertIsInstance(cel_conf.value, float)
                self.assertEqual(cel_conf.value, 0.75)
                self.assertIn("%", cel_conf.number_format)
        # 3. Verifica se o nome do curso e demais configurações estão centralizados
        for row in range(2, ws.max_row + 1):
            cel_conf = ws.cell(row=row, column=2)
            self.assertEqual(cel_conf.alignment.horizontal, "center")

    def test_extrair_data_resumo_e_duracao(self):
        # 1. Testes de extração e conversão de datas da Seção 1 (Resumo)
        # Formato informado pelo usuário: "9/14/26,8:01:54 AM" -> "14/09/2026"
        self.assertEqual(extrair_data_resumo("9/14/26,8:01:54 AM"), "14/09/2026")
        self.assertEqual(extrair_data_resumo("9/14/26, 8:01:54 AM"), "14/09/2026")
        self.assertEqual(extrair_data_resumo("14/09/2026, 08:01:54"), "14/09/2026")
        self.assertEqual(extrair_data_resumo("14/09/2026"), "14/09/2026")
        self.assertEqual(extrair_data_resumo("2026-09-14 08:01:54"), "14/09/2026")
        self.assertEqual(extrair_data_resumo(""), "")
        self.assertEqual(extrair_data_resumo(None), "")

        # 2. Testes de conversão de tempo da Seção 1 (Resumo)
        # Formato informado pelo usuário: "4h 2m 32s"
        duracao_seg = tempo_em_segundos("4h 2m 32s")
        self.assertEqual(duracao_seg, 4 * 3600 + 2 * 60 + 32)
        self.assertEqual(segundos_em_tempo(duracao_seg), "4h 2m 32s")
        self.assertEqual(segundos_em_formato_hh_mm_ss(duracao_seg), "04:02:32")
        # Soma de durações de dois encontros (4h 2m 32s + 4h 2m 32s)
        total_seg = duracao_seg * 2
        self.assertEqual(segundos_em_tempo(total_seg), "8h 5m 4s")
        self.assertEqual(segundos_em_formato_hh_mm_ss(total_seg), "08:05:04")

    def test_ler_resumo_reuniao_formatos(self):
        # Caso 1: TSV com Hora de início e Duração da reunião
        conteudo_tsv = (
            "1. Resumo da Reunião\n"
            "Total de participantes\t10\n"
            "Hora de início\t9/14/26,8:01:54 AM\n"
            "Duração da reunião\t4h 2m 32s\n"
            "2. Participantes\n"
            "Nome\tEmail\n"
        )
        arq_tsv = self.temp_dir / "resumo_tsv.csv"
        arq_tsv.write_text(conteudo_tsv, encoding="utf-16")
        res_tsv = ler_resumo_reuniao(arq_tsv)
        self.assertEqual(res_tsv["data"], "14/09/2026")
        self.assertEqual(res_tsv["duracao_str"], "4h 2m 32s")
        self.assertEqual(res_tsv["duracao_segundos"], 14552)

        # Caso 2: CSV com aspas e vírgulas
        conteudo_csv = (
            '"1. Resumo da Reunião"\n'
            '"Total de participantes","5"\n'
            '"Hora de início","9/14/26,8:01:54 AM"\n'
            '"Duração da reunião","3h 15m 0s"\n'
            '"2. Participantes"\n'
        )
        arq_csv = self.temp_dir / "resumo_csv.csv"
        arq_csv.write_text(conteudo_csv, encoding="utf-8")
        res_csv = ler_resumo_reuniao(arq_csv)
        self.assertEqual(res_csv["data"], "14/09/2026")
        self.assertEqual(res_csv["duracao_str"], "3h 15m 0s")
        self.assertEqual(res_csv["duracao_segundos"], 3 * 3600 + 15 * 60)

        # Caso 3: Arquivo sem Seção 1 (Resumo ausente)
        arq_vazio = self.temp_dir / "resumo_vazio.csv"
        arq_vazio.write_text("Nome,Email\n", encoding="utf-8")
        res_vazio = ler_resumo_reuniao(arq_vazio)
        self.assertEqual(res_vazio["data"], "")
        self.assertEqual(res_vazio["duracao_str"], "")
        self.assertEqual(res_vazio["duracao_segundos"], 0)

    def test_ordenar_arquivos_presenca_cronologico(self):
        # 1. Arquivos com Seção 1 (Resumo) fora de ordem
        arq1 = self.temp_dir / "arquivo_b.csv"
        arq1.write_text("1. Resumo\nHora de início\t9/15/26,8:01:54 AM\nDuração da reunião\t1h\n2. Participantes\nNome\tEmail\n", encoding="utf-16")
        arq2 = self.temp_dir / "arquivo_a.csv"
        arq2.write_text("1. Resumo\nHora de início\t9/14/26,8:01:54 AM\nDuração da reunião\t1h\n2. Participantes\nNome\tEmail\n", encoding="utf-16")
        arq3 = self.temp_dir / "arquivo_c.csv"
        arq3.write_text("1. Resumo\nHora de início\t9/16/26,8:01:54 AM\nDuração da reunião\t1h\n2. Participantes\nNome\tEmail\n", encoding="utf-16")

        ordenados = ordenar_arquivos_presenca([arq3, arq1, arq2])
        self.assertEqual([f.name for f in ordenados], ["arquivo_a.csv", "arquivo_b.csv", "arquivo_c.csv"])

        # 2. Arquivos sem Seção 1, com datas nos nomes fora de ordem
        arq_d1 = self.temp_dir / "Reunião_2026-09-14.csv"
        arq_d2 = self.temp_dir / "Reunião_2026-09-15.csv"
        arq_d3 = self.temp_dir / "Reunião_2026-09-18.csv"
        for a in (arq_d1, arq_d2, arq_d3):
            a.write_text("Nome,Email\n", encoding="utf-8")

        ordenados_nome = ordenar_arquivos_presenca([arq_d3, arq_d1, arq_d2])
        self.assertEqual([f.name for f in ordenados_nome], ["Reunião_2026-09-14.csv", "Reunião_2026-09-15.csv", "Reunião_2026-09-18.csv"])

        # 3. Arquivos numerados (ordenação natural: dia_1, dia_2, dia_10)
        arq_num1 = self.temp_dir / "dia_1.csv"
        arq_num2 = self.temp_dir / "dia_2.csv"
        arq_num10 = self.temp_dir / "dia_10.csv"
        for a in (arq_num1, arq_num2, arq_num10):
            a.write_text("Nome,Email\n", encoding="utf-8")

        ordenados_num = ordenar_arquivos_presenca([arq_num10, arq_num2, arq_num1])
        self.assertEqual([f.name for f in ordenados_num], ["dia_1.csv", "dia_2.csv", "dia_10.csv"])

    def test_processar_interacoes_participantes_com_colunas_faltantes_e_nan(self):
        # Valida que colunas ausentes ou células com NaN/vazio são tratadas como 0
        # E que todas as reações são agrupadas/somadas corretamente
        df_dados = pd.DataFrame({
            "Nome": ["Alice Silva", "Bruno Souza", "Carlos Pereira"],
            "Email": ["alice@anm.gov.br", "bruno@anm.gov.br", "carlos@anm.gov.br"],
            "Compromisso: Câmera Ligada": [2, None, "1"],
            "Compromisso: Desativar Mudo": [3, "", None],
            "Compromisso: Levantar as mãos": [1, None, 4],
            "Envolvimento: Reação - Aplauso": [5, None, 1],
            "Envolvimento: Reação - Gargalhada": [2, None, None],
            "Envolvimento: Reação - Curtir": [1, 2, None],
            # As colunas 'Amor' e 'Surpreso' não existem no DF (simulando coluna inexistente)
        })

        df_int = processar_interacoes_participantes(df_dados)
        self.assertEqual(len(df_int), 3)

        alice = df_int[df_int["Nome"] == "ALICE SILVA"].iloc[0]
        self.assertEqual(alice["camera_qtd"], 2)
        self.assertEqual(alice["camera_ligou"], 1)
        self.assertEqual(alice["microfone_qtd"], 3)
        self.assertEqual(alice["microfone_ligou"], 1)
        self.assertEqual(alice["maos_qtd"], 1)
        self.assertEqual(alice["reacoes_qtd"], 8)  # 5 aplausos + 2 gargalhadas + 1 curtir

        bruno = df_int[df_int["Nome"] == "BRUNO SOUZA"].iloc[0]
        self.assertEqual(bruno["camera_qtd"], 0)
        self.assertEqual(bruno["camera_ligou"], 0)
        self.assertEqual(bruno["microfone_qtd"], 0)
        self.assertEqual(bruno["microfone_ligou"], 0)
        self.assertEqual(bruno["maos_qtd"], 0)
        self.assertEqual(bruno["reacoes_qtd"], 2)  # 2 curtir

        carlos = df_int[df_int["Nome"] == "CARLOS PEREIRA"].iloc[0]
        self.assertEqual(carlos["camera_qtd"], 1)
        self.assertEqual(carlos["camera_ligou"], 1)
        self.assertEqual(carlos["microfone_qtd"], 0)
        self.assertEqual(carlos["microfone_ligou"], 0)
        self.assertEqual(carlos["maos_qtd"], 4)
        self.assertEqual(carlos["reacoes_qtd"], 1)

    def test_construcao_e_calculos_das_tres_novas_abas(self):
        # 1. Base consolidada
        relatorio = pd.DataFrame({
            "Nome": ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA", "DANIEL LIMA"],
            "Email": ["alice@anm.gov.br", "bruno@anm.gov.br", "carlos@anm.gov.br", "daniel@anm.gov.br"],
            "Presenças": [2, 1, 1, 0],
            "Apto por CH": ["APTO", "NÃO APTO", "NÃO APTO", "NÃO APTO"],
        })

        # Encontro 1
        res_e1 = pd.DataFrame({
            "Nome": ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA", "DANIEL LIMA"],
            "Compareceu": ["Presente", "Presente", "Ausente", "Ausente"],
            "Duração da Reunião": ["04:00:00", "04:00:00", "00:00:00", "00:00:00"],
        })
        int_e1 = pd.DataFrame({
            "Nome": ["ALICE SILVA", "BRUNO SOUZA"],
            "Email": ["alice@anm.gov.br", "bruno@anm.gov.br"],
            "camera_qtd": [1, 0],
            "camera_ligou": [1, 0],
            "microfone_qtd": [2, 0],
            "microfone_ligou": [1, 0],
            "maos_qtd": [1, 0],
            "reacoes_qtd": [4, 0],
        })
        resumo_e1 = {"data": "14/09/2026", "duracao_str": "4h 2m 32s", "duracao_segundos": 14552}

        # Encontro 2
        res_e2 = pd.DataFrame({
            "Nome": ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA", "DANIEL LIMA"],
            "Compareceu": ["Presente", "Ausente", "Presente", "Ausente"],
            "Duração da Reunião": ["04:00:00", "00:00:00", "04:00:00", "00:00:00"],
        })
        int_e2 = pd.DataFrame({
            "Nome": ["ALICE SILVA", "CARLOS PEREIRA"],
            "Email": ["alice@anm.gov.br", "carlos@anm.gov.br"],
            "camera_qtd": [1, 1],
            "camera_ligou": [1, 1],
            "microfone_qtd": [1, 0],
            "microfone_ligou": [1, 0],
            "maos_qtd": [2, 0],
            "reacoes_qtd": [3, 1],
        })
        resumo_e2 = {"data": "15/09/2026", "duracao_str": "4h 2m 32s", "duracao_segundos": 14552}

        lista_int = [int_e1, int_e2]
        lista_res = [res_e1, res_e2]
        lista_resumos = [resumo_e1, resumo_e2]

        # Geração da Aba 1 (Interações por participante)
        df_aba1 = criar_aba_interacoes_participante(relatorio, lista_int, lista_res)
        self.assertEqual(
            list(df_aba1.columns),
            [
                "Participante",
                "Qtd Presenças",
                "Câmera",
                "% Câmera",
                "Microfone",
                "% Microfone",
                "Mãos levantadas",
                "Total Reações",
            ]
        )
        # Alice: 2 presenças, ligou câmera nos 2 (100%), ligou mic nos 2 (100%), 3 mãos, 7 reações
        alice_aba1 = df_aba1[df_aba1["Participante"] == "ALICE SILVA"].iloc[0]
        self.assertEqual(alice_aba1["Qtd Presenças"], 2)
        self.assertEqual(alice_aba1["Câmera"], 2)
        self.assertEqual(alice_aba1["% Câmera"], 1.0)
        self.assertEqual(alice_aba1["Microfone"], 2)
        self.assertEqual(alice_aba1["% Microfone"], 1.0)
        self.assertEqual(alice_aba1["Mãos levantadas"], 3)
        self.assertEqual(alice_aba1["Total Reações"], 7)

        # Daniel: 0 presenças -> cobertura 0.0 sem ZeroDivisionError
        daniel_aba1 = df_aba1[df_aba1["Participante"] == "DANIEL LIMA"].iloc[0]
        self.assertEqual(daniel_aba1["Qtd Presenças"], 0)
        self.assertEqual(daniel_aba1["Câmera"], 0)
        self.assertEqual(daniel_aba1["% Câmera"], 0.0)

        # Geração da Aba 2 (Interações por encontro)
        df_aba2 = criar_aba_interacoes_encontro(4, lista_resumos, lista_int, lista_res)
        self.assertEqual(
            list(df_aba2.columns),
            ["ID", "Data", "Carga horária", "Presentes", "Ligou a câmera", "Ligou o microfone", "Levantou a mão", "Reações"]
        )
        self.assertEqual(len(df_aba2), 2)
        e1 = df_aba2.iloc[0]
        self.assertEqual(e1["ID"], 1)
        self.assertEqual(e1["Data"], "14/09/2026")
        self.assertEqual(e1["Carga horária"], "04:02:32")
        self.assertEqual(e1["Presentes"], "2 de 4")
        self.assertEqual(e1["Ligou a câmera"], "1 de 2")
        self.assertEqual(e1["Ligou o microfone"], "1 de 2")
        self.assertEqual(e1["Levantou a mão"], 1)
        self.assertEqual(e1["Reações"], 4)

        # Geração da Aba 3 (Indicadores)
        df_aba3 = criar_aba_indicadores(4, relatorio, df_aba1, df_aba2, lista_resumos, lista_res)
        self.assertEqual(list(df_aba3.columns), ["Indicador", "Resultado"])
        mapa_ind = dict(zip(df_aba3["Indicador"], df_aba3["Resultado"]))
        self.assertEqual(mapa_ind["Convidados a participar"], 4)
        self.assertEqual(mapa_ind["Participantes totais"], 3)  # Alice, Bruno, Carlos
        self.assertEqual(mapa_ind["Participantes certificados"], 1)  # Alice
        self.assertEqual(mapa_ind["Carga horária realizada"], "08:05:04")
        self.assertEqual(mapa_ind["Taxa de certificação"], 1 / 4)
        self.assertEqual(mapa_ind["Participações com Câmera"], 2 / 3)  # Alice e Carlos
        self.assertEqual(mapa_ind["Participações com Microfone"], 1 / 3)  # Alice
        self.assertEqual(mapa_ind["Total de reações"], 8)
        self.assertEqual(mapa_ind["Total de mãos levantadas"], 3)

    def test_exportacao_e_formatacao_estilos_openpyxl(self):
        import openpyxl
        relatorio = pd.DataFrame({
            "Nome": ["ALICE SILVA"],
            "Presenças": [1],
            "Apto por CH": ["APTO"],
        })
        df_aba1 = pd.DataFrame({
            "Participante": ["ALICE SILVA"],
            "Qtd Presenças": [1],
            "Câmera": [1],
            "% Câmera": [1.0],
            "Microfone": [1],
            "% Microfone": [1.0],
            "Mãos levantadas": [2],
            "Total Reações": [5],
        })
        df_aba2 = pd.DataFrame({
            "ID": [1],
            "Data": ["14/09/2026"],
            "Carga horária": ["04:00:00"],
            "Presentes": ["1 de 1"],
            "Ligou a câmera": ["1 de 1"],
            "Ligou o microfone": ["1 de 1"],
            "Levantou a mão": [2],
            "Reações": [5],
        })
        df_aba3 = pd.DataFrame([
            {"Indicador": "Convidados a participar", "Resultado": 1},
            {"Indicador": "Taxa de certificação", "Resultado": 1.0},
        ])

        saida = self.temp_dir / "teste_formatacao_completa.xlsx"
        exportar_excel(
            relatorio=relatorio,
            caminho_saida=saida,
            df_interacoes_participante=df_aba1,
            df_interacoes_encontro=df_aba2,
            df_indicadores=df_aba3,
        )

        wb = openpyxl.load_workbook(saida)
        self.assertIn("Interações por participante", wb.sheetnames)
        self.assertIn("Interações por encontro", wb.sheetnames)
        self.assertIn("Indicadores", wb.sheetnames)

        # Valida formatação de porcentagem na Aba 1
        ws1 = wb["Interações por participante"]
        cel_cob_cam = ws1.cell(row=2, column=4)
        self.assertIn("%", cel_cob_cam.number_format)
        self.assertEqual(cel_cob_cam.value, 1.0)

        # Valida formatação de porcentagem na Aba 3
        ws3 = wb["Indicadores"]
        cel_taxa = ws3.cell(row=3, column=2)
        self.assertIn("%", cel_taxa.number_format)
        self.assertEqual(cel_taxa.value, 1.0)

    def test_criar_aba_custos_e_formatacao(self):
        """Testa geração da aba Custos e sua formatação no Excel."""
        import openpyxl

        relatorio = pd.DataFrame({
            "Nome": ["Alice", "Bruno", "Carlos", "Diana", "Eduardo"],
            "Presenças": [2, 1, 0, 2, 1],
            "Apto por CH": ["APTO", "APTO", "NÃO APTO", "APTO", "NÃO APTO"],
        })
        lista_resumos = [
            {"duracao_segundos": 7200},  # 2 horas
            {"duracao_segundos": 7200},  # 2 horas -> total 4 horas (14400s)
        ]
        lista_resultados = [
            pd.DataFrame({"Nome": ["Alice", "Bruno"]}),
            pd.DataFrame({"Nome": ["Alice", "Diana"]}),
        ]
        config = {
            "CUSTO_TOTAL_EM_REAIS": 4000.0,
            "CARGA_HORARIA": 10.0,
        }

        df_custos = criar_aba_custos(
            total_convidados=5,
            relatorio=relatorio,
            quantidade_encontros=2,
            carga_horaria_horas=10.0,
            lista_resumos=lista_resumos,
            lista_resultados_encontros=lista_resultados,
            config=config,
        )

        self.assertEqual(len(df_custos), 12)
        mapa = dict(zip(df_custos["Item"], df_custos["Valor"]))
        self.assertEqual(mapa["Convidados a participar"], 5)
        self.assertEqual(mapa["Participantes totais"], 4)  # Alice, Bruno, Diana, Eduardo
        self.assertEqual(mapa["Participantes certificados"], 3)  # Alice, Bruno, Diana
        self.assertEqual(mapa["Custo total"], 4000.0)
        self.assertEqual(mapa["Quantidade de encontros"], 2)
        self.assertEqual(mapa["Carga horária"], "10:00:00")
        self.assertEqual(mapa["Carga horária realizada"], "04:00:00")
        self.assertEqual(mapa["Custo por convidado"], 800.0)
        self.assertEqual(mapa["Custo por participante"], 1000.0)
        self.assertAlmostEqual(mapa["Custo por certificado"], 4000.0 / 3)
        self.assertEqual(mapa["Custo por encontro"], 2000.0)
        self.assertEqual(mapa["Custo por hora"], 1000.0)  # 4000 / 4h realizadas

        # Valida exportação e openpyxl
        saida = self.temp_dir / "teste_aba_custos.xlsx"
        exportar_excel(
            relatorio=relatorio,
            caminho_saida=saida,
            df_custos=df_custos,
        )

        wb = openpyxl.load_workbook(saida)
        self.assertIn("Custos", wb.sheetnames)
        ws = wb["Custos"]
        self.assertEqual(ws.column_dimensions["A"].width, 36)
        self.assertEqual(ws.column_dimensions["B"].width, 24)

        # Checa células de custo com formatação de moeda R$ #,##0.00
        for r in range(2, ws.max_row + 1):
            item = str(ws.cell(row=r, column=1).value)
            if "Custo" in item:
                self.assertIn("R$", ws.cell(row=r, column=2).number_format)
                self.assertIsInstance(ws.cell(row=r, column=2).value, (int, float))

    def test_carregar_variaveis_com_sessao_ignorar_emails(self):
        """Testa parsing da seção [INICIO_IGNORAR_EMAILS] em variaveis.txt."""
        conteudo = """
        NOME_DO_CURSO = Curso Teste
        CARGA_HORARIA = 20
        DATA_INICIO = 01/08/2026
        DATA_FIM = 05/08/2026
        DATA_EMISSAO = 05/08/2026

        [INICIO_IGNORAR_EMAILS]
        organizador1@exemplo.gov.br
        organizador2@exemplo.gov.br
        organizador3@exemplo.gov.br
        [FIM_IGNORAR_EMAILS]

        [INICIO_CORPO_EMAIL]
        Corpo teste
        [FIM_CORPO_EMAIL]
        """
        arq = self.temp_dir / "variaveis_ignorar.txt"
        arq.write_text(conteudo, encoding="utf-8")

        raw = carregar_variaveis(arq)
        self.assertIn("IGNORAR_EMAILS", raw)
        self.assertEqual(
            raw["IGNORAR_EMAILS"],
            [
                "organizador1@exemplo.gov.br",
                "organizador2@exemplo.gov.br",
                "organizador3@exemplo.gov.br",
            ]
        )

        cfg = validar_variaveis(raw)
        self.assertEqual(
            cfg["IGNORAR_EMAILS"],
            [
                "organizador1@exemplo.gov.br",
                "organizador2@exemplo.gov.br",
                "organizador3@exemplo.gov.br",
            ]
        )

    def test_filtrar_emails_ignorados(self):
        """Testa função filtrar_emails_ignorados com diferentes maiúsculas/minúsculas e formatos."""
        df = pd.DataFrame({
            "Nome": ["Alice", "Organizador Um", "Bruno", "Organizador Dois"],
            "Email": [
                "alice@anm.gov.br",
                "organizador1@exemplo.gov.br",
                "bruno@anm.gov.br",
                "organizador2@exemplo.gov.br",
            ],
            "Duracao": ["1h", "2h", "1h", "2h"],
        })
        ignorados = ["organizador1@exemplo.gov.br", "organizador2@exemplo.gov.br"]

        filtrado = filtrar_emails_ignorados(df, ignorados)
        self.assertEqual(len(filtrado), 2)
        self.assertEqual(list(filtrado["Nome"]), ["Alice", "Bruno"])
        self.assertEqual(list(filtrado["Email"]), ["alice@anm.gov.br", "bruno@anm.gov.br"])

    def test_preparar_convocados_com_emails_ignorados(self):
        """Testa se pessoas com e-mails ignorados são excluídas da lista de convocados."""
        csv_convocados = self.temp_dir / "convocados_teste.csv"
        csv_convocados.write_text(
            "Relatório\nData\nNome\n"
            "Alice Silva <alice@anm.gov.br>\n"
            "Organizador Um <organizador1@exemplo.gov.br>\n"
            "Bruno Souza <bruno@anm.gov.br>\n",
            encoding="utf-8"
        )
        servidores = pd.DataFrame({
            "NOME SERVIDOR": ["ALICE SILVA", "ORGANIZADOR UM", "BRUNO SOUZA"],
            "UORG": ["SGP", "DIR", "DITEC"],
        })
        ignorados = ["organizador1@exemplo.gov.br"]

        df_conv = preparar_convocados(csv_convocados, servidores, emails_ignorados=ignorados)
        self.assertEqual(len(df_conv), 2)
        self.assertNotIn("ORGANIZADOR UM", list(df_conv["Nome"]))
        self.assertNotIn("organizador1@exemplo.gov.br", list(df_conv["Email"]))

    def test_gerar_relatorio_com_emails_ignorados(self):
        """Testa se gerar_relatorio exclui integralmente os e-mails ignorados de todas as abas."""
        from core.presenca import gerar_relatorio
        import openpyxl

        # 1. Servidores
        caminho_servidores = self.temp_dir / "servidores_ign.xlsx"
        df_serv = pd.DataFrame({
            "NOME SERVIDOR": ["ALICE SILVA", "ORGANIZADOR UM", "BRUNO SOUZA"],
            "UORG": ["SGP", "DIR", "DITEC"],
            "EMAIL": ["alice@anm.gov.br", "organizador1@exemplo.gov.br", "bruno@anm.gov.br"],
        })
        df_serv.to_excel(caminho_servidores, index=False)

        # 2. Convocados
        caminho_conv = self.temp_dir / "convocados_ign.csv"
        caminho_conv.write_text(
            "Relatório de Convocados\nData\nNome\n"
            "ALICE SILVA <alice@anm.gov.br>\n"
            "ORGANIZADOR UM <organizador1@exemplo.gov.br>\n"
            "BRUNO SOUZA <bruno@anm.gov.br>\n",
            encoding="utf-8"
        )

        # 3. Presença Teams
        caminho_pres = self.temp_dir / "presenca_ign.csv"
        caminho_pres.write_text(
            "1. Resumo da Reunião\n"
            "Total de participantes\t3\n"
            "Hora de início\t9/14/26,8:01:54 AM\n"
            "Duração da reunião\t2h 0m 0s\n"
            "2. Participantes\n"
            "Nome\tEmail\tDuração da Reunião\tPapel\tCompromisso: Câmera Ligada\tCompromisso: Desativar Mudo\tCompromisso: Levantar as mãos\tEnvolvimento: Reação - Aplauso\n"
            "ALICE SILVA\talice@anm.gov.br\t2h 0m 0s\tParticipante\t1\t1\t1\t1\n"
            "ORGANIZADOR UM\torganizador1@exemplo.gov.br\t2h 0m 0s\tParticipante\t1\t1\t1\t1\n"
            "BRUNO SOUZA\tbruno@anm.gov.br\t2h 0m 0s\tParticipante\t0\t0\t0\t0\n"
            "3. Atividades em Reunião\n",
            encoding="utf-8"
        )

        caminho_saida = self.temp_dir / "Relatorio_Ignorados.xlsx"
        config = {
            "NOME_DO_CURSO": "Curso Teste Ignorados",
            "CARGA_HORARIA": 2.0,
            "CUSTO_TOTAL_EM_REAIS": 1000.0,
            "DATA_INICIO": "14/09/2026",
            "DATA_FIM": "14/09/2026",
            "DATA_EMISSAO": "14/09/2026",
            "PERCENTUAL_MINIMO_CARGA": 0.75,
            "PERCENTUAL_MINIMO_ENCONTROS": 0.75,
            "TEMPO_MINIMO_PRESENCA_MINUTOS": 30,
            "IGNORAR_EMAILS": ["organizador1@exemplo.gov.br"],
        }

        df_rel = gerar_relatorio(
            caminho_servidores=caminho_servidores,
            caminho_convocados=caminho_conv,
            arquivos_presenca=[caminho_pres],
            caminho_saida_excel=caminho_saida,
            carga_horaria_horas=2.0,
            tempo_comparecimento_minutos=30,
            config=config,
        )

        # Organizador Um não deve constar no relatório principal
        self.assertEqual(len(df_rel), 2)
        self.assertNotIn("ORGANIZADOR UM", list(df_rel["Nome"]))

        # Verifica as abas do Excel
        wb = openpyxl.load_workbook(caminho_saida)
        for sname in ["Certificação", "Presenças", "Interações por participante"]:
            ws = wb[sname]
            nomes_na_aba = [str(ws.cell(row=r, column=1).value) for r in range(2, ws.max_row + 1)]
            self.assertNotIn("ORGANIZADOR UM", nomes_na_aba)

        # Aba Custos: total_convidados deve ser 2 e não 3
        df_custos = pd.read_excel(caminho_saida, sheet_name="Custos")
        mapa_c = dict(zip(df_custos["Item"], df_custos["Valor"]))
        self.assertEqual(mapa_c["Convidados a participar"], 2)
        self.assertEqual(mapa_c["Custo por convidado"], 500.0)

    def test_formatar_aba_presencas_mesclagem_titulos(self):
        """Testa mesclagem de títulos 'Encontro n' e 'Comparecimento n' com ID e data na aba Presenças."""
        import openpyxl
        from core.presenca import formatar_aba_presencas, criar_estilos

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Presenças"
        ws.append(["Nome", "Encontro 1", "Comparecimento 1", "Encontro 2", "Comparecimento 2"])
        ws.append(["Alice Silva", "04:00:00", "Presente", "04:00:00", "Presente"])
        ws.append(["Carlos Pereira", "00:00:00", "Ausente", "04:00:00", "Presente"])

        estilos = criar_estilos()
        datas = ["14/09/2026", "15/09/2026"]
        formatar_aba_presencas(ws, wb, estilos, datas_encontros=datas)

        # Valida faixas mescladas
        merged = {str(r) for r in ws.merged_cells.ranges}
        self.assertIn("B1:C1", merged)
        self.assertIn("D1:E1", merged)

        # Valida novos títulos unificados com ID e data
        self.assertEqual(ws["B1"].value, "Encontro 1 — 14/09/2026")
        self.assertIsNone(ws["C1"].value)
        self.assertEqual(ws["D1"].value, "Encontro 2 — 15/09/2026")
        self.assertIsNone(ws["E1"].value)

        # Valida alinhamento e formatação de 'Ausente'
        self.assertEqual(ws["B2"].alignment.horizontal, "center")
        self.assertEqual(ws["C2"].alignment.horizontal, "center")
        self.assertTrue(ws["B3"].font.bold)
        self.assertTrue(ws["C3"].font.bold)
        self.assertFalse(ws["B2"].font.bold)

        # Valida largura das colunas
        self.assertGreaterEqual(ws.column_dimensions["A"].width, 30)
        self.assertEqual(ws.column_dimensions["B"].width, 17)
        self.assertEqual(ws.column_dimensions["C"].width, 17)


if __name__ == "__main__":
    unittest.main()



