"""
Teste de ponta a ponta (E2E) simulando um curso completo com dados sintéticos.
Valida:
1. Criação das pastas e variaveis.txt
2. Validação da estrutura
3. Processamento e geração do relatório de presenças Excel
"""

import unittest
from pathlib import Path
import shutil
import tempfile
import pandas as pd
from openpyxl import Workbook

import sys
diretorio_raiz = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(diretorio_raiz))

from core.validator import validar_estrutura_curso
from core.config import obter_configuracao
from core.presenca import gerar_relatorio


class TestE2EPipeline(unittest.TestCase):

    def setUp(self):
        self.temp_curso = Path(tempfile.mkdtemp()) / "curso_teste_integracao"
        self.temp_curso.mkdir(parents=True)

        # 1. Estrutura de pastas
        (self.temp_curso / "lista_convidados").mkdir()
        (self.temp_curso / "listas_de_presenca").mkdir()
        (self.temp_curso / "lista_servidores").mkdir()
        (self.temp_curso / "modelo_certificado").mkdir()

        # 2. variaveis.txt
        var_content = """
        NOME_DO_CURSO = Curso de Teste Unitario
        CARGA_HORARIA = 10
        CUSTO_TOTAL_EM_REAIS = 3000.00
        DATA_INICIO = 01/09/2026
        DATA_FIM = 02/09/2026
        LOCAL_EMISSAO = Brasília — DF
        DATA_EMISSAO = 02/09/2026
        PERCENTUAL_MINIMO_CARGA = 0.75
        PERCENTUAL_MINIMO_ENCONTROS = 0.75
        TEMPO_MINIMO_PRESENCA_MINUTOS = 15
        ASSUNTO_EMAIL = Certificado - {CURSO}

        [INICIO_CORPO_EMAIL]
        Olá {nome}, seu certificado está pronto.
        [FIM_CORPO_EMAIL]
        """
        (self.temp_curso / "variaveis.txt").write_text(var_content, encoding="utf-8")

        # 3. Planilha de Servidores Excel (.xlsx)
        df_servidores = pd.DataFrame({
            "NOME SERVIDOR": ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA"],
            "UORG": ["SGP/ANM", "DITEC/ANM", "SGP/ANM"],
            "EMAIL": ["alice@anm.gov.br", "bruno@anm.gov.br", "carlos@anm.gov.br"],
        })
        df_servidores.to_excel(self.temp_curso / "lista_servidores" / "servidores.xlsx", index=False)

        # 4. Lista de Convocados (.csv)
        # Formato do Teams: 2 linhas de cabeçalho, depois "Nome <email>"
        convocados_content = (
            "Relatório de Convocados\n"
            "Gerado em 01/09/2026\n"
            "Nome\n"
            "ALICE SILVA <alice@anm.gov.br>\n"
            "BRUNO SOUZA <bruno@anm.gov.br>\n"
            "CARLOS PEREIRA <carlos@anm.gov.br>\n"
        )
        (self.temp_curso / "lista_convidados" / "convocados.csv").write_text(convocados_content, encoding="utf-8")

        # 5. Listas de Presença (.csv diários - formato Teams)
        # Teams tem seção "1. Resumo", depois "2. Participantes" com interações, depois "3. Atividades em Reunião"
        def gerar_presenca_teams(participantes, data_inicio="9/14/26,8:01:54 AM", duracao="4h 0m 0s"):
            linhas = [
                "1. Resumo da Reunião",
                f"Total de participantes\t{len(participantes)}",
                f"Hora de início\t{data_inicio}",
                f"Duração da reunião\t{duracao}",
                "2. Participantes",
                "Nome\tEmail\tDuração da Reunião\tPapel\tCompromisso: Câmera Ligada\tCompromisso: Desativar Mudo\tCompromisso: Levantar as mãos\tEnvolvimento: Reação - Aplauso\tEnvolvimento: Reação - Curtir",
            ]
            for p in participantes:
                nome = p[0]
                email = p[1]
                dur = p[2]
                cam = p[3] if len(p) > 3 else ""
                mic = p[4] if len(p) > 4 else ""
                mao = p[5] if len(p) > 5 else ""
                aplauso = p[6] if len(p) > 6 else ""
                curtir = p[7] if len(p) > 7 else ""
                linhas.append(f"{nome}\t{email}\t{dur}\tParticipante\t{cam}\t{mic}\t{mao}\t{aplauso}\t{curtir}")
            linhas.append("3. Atividades em Reunião")
            linhas.append("Apresentador\tSim")
            return "\n".join(linhas)

        # Encontro 1: Alice 4h (cam=1, mic=2, mao=1, aplauso=3, curtir=2), Bruno 4h (sem interações), Carlos ausente
        dia1 = gerar_presenca_teams([
            ("Alice Silva", "alice@anm.gov.br", "4h 0m 0s", "1", "2", "1", "3", "2"),
            ("Bruno Souza", "bruno@anm.gov.br", "4h 0m 0s", "", "", "", "", ""),
        ], data_inicio="9/14/26,8:01:54 AM", duracao="4h 0m 0s")
        (self.temp_curso / "listas_de_presenca" / "dia_1.csv").write_text(dia1, encoding="utf-16")

        # Encontro 2: Alice 4h (cam=1, mic=1, mao=2, aplauso=1, curtir=1), Carlos 4h (cam=1), Bruno ausente
        dia2 = gerar_presenca_teams([
            ("Alice Silva", "alice@anm.gov.br", "4h 0m 0s", "1", "1", "2", "1", "1"),
            ("Carlos Pereira", "carlos@anm.gov.br", "4h 0m 0s", "1", "", "", "", ""),
        ], data_inicio="9/15/26,8:01:54 AM", duracao="4h 0m 0s")
        (self.temp_curso / "listas_de_presenca" / "dia_2.csv").write_text(dia2, encoding="utf-16")

        # 6. Modelo de PPTX mínimo
        from pptx import Presentation
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[0])
        prs.save(self.temp_curso / "modelo_certificado" / "modelo.pptx")

    def tearDown(self):
        shutil.rmtree(self.temp_curso.parent, ignore_errors=True)

    def test_pipeline_completo_relatorio(self):
        # 1. Valida estrutura
        auditoria = validar_estrutura_curso(self.temp_curso)
        self.assertIn("convidados", auditoria["arquivos"])
        self.assertIn("servidores", auditoria["arquivos"])
        self.assertIn("presencas", auditoria["arquivos"])
        self.assertEqual(len(auditoria["arquivos"]["presencas"]), 2)

        # 2. Carrega config
        config = obter_configuracao(self.temp_curso)
        self.assertEqual(config["NOME_DO_CURSO"], "Curso de Teste Unitario")

        # 3. Executa geração do relatório
        saida_excel = self.temp_curso / "saida" / "relatorio_presencas" / "Relatorio.xlsx"
        df_resultado = gerar_relatorio(
            caminho_servidores=auditoria["arquivos"]["servidores"],
            caminho_convocados=auditoria["arquivos"]["convidados"],
            arquivos_presenca=auditoria["arquivos"]["presencas"],
            caminho_saida_excel=saida_excel,
            carga_horaria_horas=config["CARGA_HORARIA"],
            tempo_comparecimento_minutos=config["TEMPO_MINIMO_PRESENCA_MINUTOS"],
            percentual_minimo_carga=config["PERCENTUAL_MINIMO_CARGA"],
            percentual_minimo_encontros=config["PERCENTUAL_MINIMO_ENCONTROS"],
        )

        self.assertTrue(saida_excel.exists())
        self.assertEqual(len(df_resultado), 3)

        # Valida que as colunas da aba Certificação NÃO contêm Encontro/Comparecimento
        self.assertNotIn("Encontro 1", df_resultado.columns)
        self.assertNotIn("Encontro 2", df_resultado.columns)
        self.assertNotIn("Comparecimento 1", df_resultado.columns)
        self.assertNotIn("Comparecimento 2", df_resultado.columns)
        self.assertIn("Presenças", df_resultado.columns)
        self.assertIn("Apto por Presença", df_resultado.columns)
        # Valida que as pessoas estão ordenadas por ordem alfabética
        self.assertEqual(list(df_resultado["Nome"]), ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA"])

        # Valida as 6 abas da planilha
        import openpyxl
        wb = openpyxl.load_workbook(saida_excel)
        self.assertEqual(
            wb.sheetnames,
            [
                "Certificação",
                "Presenças",
                "Interações por participante",
                "Interações por encontro",
                "Indicadores",
                "Custos",
                "Informações do Curso",
            ]
        )

        # Valida conteúdo da aba 'Presenças' (apenas Nome e colunas dos encontros)
        df_presencas_excel = pd.read_excel(saida_excel, sheet_name="Presenças")
        self.assertEqual(
            list(df_presencas_excel.columns),
            ["Nome", "Encontro 1", "Comparecimento 1", "Encontro 2", "Comparecimento 2"]
        )
        self.assertEqual(list(df_presencas_excel["Nome"]), ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA"])

        # Valida Aba 1: 'Interações por participante'
        df_int_part_excel = pd.read_excel(saida_excel, sheet_name="Interações por participante")
        self.assertEqual(
            list(df_int_part_excel.columns),
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
        alice_int = df_int_part_excel[df_int_part_excel["Participante"] == "ALICE SILVA"].iloc[0]
        self.assertEqual(alice_int["Qtd Presenças"], 2)
        self.assertEqual(alice_int["Câmera"], 2)
        self.assertAlmostEqual(alice_int["% Câmera"], 1.0)
        self.assertEqual(alice_int["Microfone"], 2)
        self.assertAlmostEqual(alice_int["% Microfone"], 1.0)
        self.assertEqual(alice_int["Mãos levantadas"], 3)
        self.assertEqual(alice_int["Total Reações"], 7)

        bruno_int = df_int_part_excel[df_int_part_excel["Participante"] == "BRUNO SOUZA"].iloc[0]
        self.assertEqual(bruno_int["Qtd Presenças"], 1)
        self.assertEqual(bruno_int["Câmera"], 0)
        self.assertAlmostEqual(bruno_int["% Câmera"], 0.0)
        self.assertEqual(bruno_int["Mãos levantadas"], 0)
        self.assertEqual(bruno_int["Total Reações"], 0)

        carlos_int = df_int_part_excel[df_int_part_excel["Participante"] == "CARLOS PEREIRA"].iloc[0]
        self.assertEqual(carlos_int["Qtd Presenças"], 1)
        self.assertEqual(carlos_int["Câmera"], 1)
        self.assertAlmostEqual(carlos_int["% Câmera"], 1.0)
        self.assertEqual(carlos_int["Microfone"], 0)

        # Valida Aba 2: 'Interações por encontro'
        df_int_enc_excel = pd.read_excel(saida_excel, sheet_name="Interações por encontro")
        self.assertEqual(
            list(df_int_enc_excel.columns),
            ["ID", "Data", "Carga horária", "Presentes", "Ligou a câmera", "Ligou o microfone", "Levantou a mão", "Reações"]
        )
        self.assertEqual(len(df_int_enc_excel), 2)
        enc1 = df_int_enc_excel.iloc[0]
        self.assertEqual(enc1["ID"], 1)
        self.assertEqual(enc1["Data"], "14/09/2026")
        self.assertEqual(enc1["Carga horária"], "04:00:00")
        self.assertEqual(enc1["Presentes"], "2 de 3")
        self.assertEqual(enc1["Ligou a câmera"], "1 de 2")
        self.assertEqual(enc1["Ligou o microfone"], "1 de 2")
        self.assertEqual(enc1["Levantou a mão"], 1)
        self.assertEqual(enc1["Reações"], 5)

        enc2 = df_int_enc_excel.iloc[1]
        self.assertEqual(enc2["ID"], 2)
        self.assertEqual(enc2["Data"], "15/09/2026")
        self.assertEqual(enc2["Presentes"], "2 de 3")
        self.assertEqual(enc2["Ligou a câmera"], "2 de 2")
        self.assertEqual(enc2["Ligou o microfone"], "1 de 2")
        self.assertEqual(enc2["Levantou a mão"], 2)
        self.assertEqual(enc2["Reações"], 2)

        # Valida Aba 3: 'Indicadores'
        df_ind_excel = pd.read_excel(saida_excel, sheet_name="Indicadores")
        self.assertEqual(list(df_ind_excel.columns), ["Indicador", "Resultado"])
        mapa_ind = dict(zip(df_ind_excel["Indicador"], df_ind_excel["Resultado"]))
        self.assertEqual(mapa_ind["Convidados a participar"], 3)
        self.assertEqual(mapa_ind["Participantes totais"], 3)
        self.assertEqual(mapa_ind["Participantes certificados"], 1)
        self.assertEqual(mapa_ind["Carga horária realizada"], "08:00:00")
        self.assertAlmostEqual(mapa_ind["Taxa de certificação"], 1 / 3)
        self.assertAlmostEqual(mapa_ind["Participações com Câmera"], 2 / 3)
        self.assertAlmostEqual(mapa_ind["Participações com Microfone"], 1 / 3)
        self.assertEqual(mapa_ind["Total de reações"], 7)
        self.assertEqual(mapa_ind["Total de mãos levantadas"], 3)

        # Valida Aba 4: 'Custos'
        df_custos_excel = pd.read_excel(saida_excel, sheet_name="Custos")
        self.assertEqual(list(df_custos_excel.columns), ["Item", "Valor"])
        self.assertEqual(len(df_custos_excel), 12)
        mapa_custos = dict(zip(df_custos_excel["Item"], df_custos_excel["Valor"]))
        self.assertEqual(mapa_custos["Convidados a participar"], 3)
        self.assertEqual(mapa_custos["Participantes totais"], 3)
        self.assertEqual(mapa_custos["Participantes certificados"], 1)
        self.assertEqual(mapa_custos["Custo total"], 3000.0)
        self.assertEqual(mapa_custos["Quantidade de encontros"], 2)
        self.assertEqual(mapa_custos["Carga horária"], "10:00:00")
        self.assertEqual(mapa_custos["Carga horária realizada"], "08:00:00")
        self.assertEqual(mapa_custos["Custo por convidado"], 1000.0)
        self.assertEqual(mapa_custos["Custo por participante"], 1000.0)
        self.assertEqual(mapa_custos["Custo por certificado"], 3000.0)
        self.assertEqual(mapa_custos["Custo por encontro"], 1500.0)
        self.assertEqual(mapa_custos["Custo por hora"], 375.0)

        ws_custos = wb["Custos"]
        # Verifica se células de custos possuem formatação de moeda
        for row in range(2, ws_custos.max_row + 1):
            item_txt = str(ws_custos.cell(row=row, column=1).value)
            if "Custo" in item_txt:
                self.assertIn("R$", ws_custos.cell(row=row, column=2).number_format)
                self.assertIsInstance(ws_custos.cell(row=row, column=2).value, (int, float))

        # Valida conteúdo da aba 'Informações do Curso'
        df_info_excel = pd.read_excel(saida_excel, sheet_name="Informações do Curso")
        self.assertEqual(list(df_info_excel.columns), ["Item", "Configuração", "Observação"])
        self.assertEqual(len(df_info_excel), 9)

        valores_por_item = dict(zip(df_info_excel["Item"], df_info_excel["Configuração"]))
        # Verifica se porcentagem está gravada como número nativo com formato percentual (0.75 / 75%)
        self.assertIn(valores_por_item["Percentual Mínimo de Carga Horária"], (0.75, "75%"))
        # Verifica no openpyxl se a célula tem tipo numérico float e formato de porcentagem (0%)
        wb = openpyxl.load_workbook(saida_excel)
        ws_info = wb["Informações do Curso"]
        for row in range(2, ws_info.max_row + 1):
            if "Percentual" in str(ws_info.cell(row=row, column=1).value):
                self.assertIn("%", ws_info.cell(row=row, column=2).number_format)
                self.assertIsInstance(ws_info.cell(row=row, column=2).value, float)
        # Verifica se tempo de presença está formatado (ex: 15 minutos configurados no teste)
        self.assertEqual(valores_por_item["Tempo Mínimo de Permanência na Reunião"], "15 minutos")
        # Verifica se carga horária está formatada (ex: 10 horas)
        self.assertEqual(valores_por_item["Carga Horária"], "10 horas")
        # Verifica se não há parênteses nem tags de email
        for item in df_info_excel["Item"]:
            self.assertNotIn("(", item)
            self.assertNotIn(")", item)
            self.assertNotIn("EMAIL", item.upper())

        # Alice participou 8h de 10h (80% >= 75%) e 2 de 2 encontros -> APTO
        alice = df_resultado[df_resultado["Nome"] == "ALICE SILVA"].iloc[0]
        self.assertEqual(alice["Presenças"], 2)
        self.assertEqual(alice["Apto por CH"], "APTO")

        # Bruno participou 4h de 10h (40% < 75%) -> NÃO APTO
        bruno = df_resultado[df_resultado["Nome"] == "BRUNO SOUZA"].iloc[0]
        self.assertEqual(bruno["Presenças"], 1)
        self.assertEqual(bruno["Apto por CH"], "NÃO APTO")

    def test_pipeline_relatorio_com_caracteres_especiais(self):
        """Testa se o gerador lida corretamente com dois-pontos e caracteres especiais no caminho do relatório."""
        auditoria = validar_estrutura_curso(self.temp_curso)
        config = obter_configuracao(self.temp_curso)

        # Caminho com dois-pontos (como no caso real do erro)
        saida_invalida = (
            self.temp_curso
            / "saida"
            / "relatorio_presencas"
            / "Relatório de Presença - Ambientação TAM Módulo: Transparência e Ética: Corregedoria.xlsx"
        )

        df_resultado = gerar_relatorio(
            caminho_servidores=auditoria["arquivos"]["servidores"],
            caminho_convocados=auditoria["arquivos"]["convidados"],
            arquivos_presenca=auditoria["arquivos"]["presencas"],
            caminho_saida_excel=saida_invalida,
            carga_horaria_horas=config["CARGA_HORARIA"],
            tempo_comparecimento_minutos=config["TEMPO_MINIMO_PRESENCA_MINUTOS"],
        )

        # O arquivo sanitizado esperado deve existir
        saida_esperada = (
            self.temp_curso
            / "saida"
            / "relatorio_presencas"
            / "Relatório de Presença - Ambientação TAM Módulo - Transparência e Ética - Corregedoria.xlsx"
        )
        self.assertTrue(saida_esperada.exists())
        self.assertEqual(len(df_resultado), 3)

    def test_ordenacao_cronologica_encontros_presencas_e_interacoes(self):
        # Valida que mesmo passando arquivos fora de ordem cronológica,
        # o relatório final ordena os encontros por data tanto na aba 'Presenças'
        # quanto na aba 'Interações por encontro'.
        auditoria = validar_estrutura_curso(self.temp_curso)
        config = obter_configuracao(self.temp_curso)
        saida_excel = self.temp_curso / "saida" / "relatorio_presencas" / "Relatorio_Ordem.xlsx"

        # Passa arquivos na ordem INVERSA (dia 2 antes do dia 1)
        arquivos_invertidos = list(reversed(auditoria["arquivos"]["presencas"]))

        gerar_relatorio(
            caminho_servidores=auditoria["arquivos"]["servidores"],
            caminho_convocados=auditoria["arquivos"]["convidados"],
            arquivos_presenca=arquivos_invertidos,
            caminho_saida_excel=saida_excel,
            carga_horaria_horas=config["CARGA_HORARIA"],
            tempo_comparecimento_minutos=config["TEMPO_MINIMO_PRESENCA_MINUTOS"],
            config=config,
        )

        # 1. Valida Aba 'Presenças': Encontro 1 deve ser 14/09 (dia 1) e Encontro 2 deve ser 15/09 (dia 2)
        # No dia 1, Carlos Pereira estava AUSENTE ("00:00:00") e no dia 2 estava PRESENTE ("04:00:00")
        df_pres = pd.read_excel(saida_excel, sheet_name="Presenças")
        carlos = df_pres[df_pres["Nome"] == "CARLOS PEREIRA"].iloc[0]
        self.assertEqual(carlos["Encontro 1"], "00:00:00")
        self.assertEqual(carlos["Comparecimento 1"], "Ausente")
        self.assertIn(carlos["Encontro 2"], ("4:00:00", "04:00:00"))
        self.assertEqual(carlos["Comparecimento 2"], "Presente")

        # 2. Valida Aba 'Interações por encontro': Linhas ordenadas por data cronológica
        df_int_enc = pd.read_excel(saida_excel, sheet_name="Interações por encontro")
        self.assertEqual(df_int_enc.iloc[0]["ID"], 1)
        self.assertEqual(df_int_enc.iloc[0]["Data"], "14/09/2026")
        self.assertEqual(df_int_enc.iloc[1]["ID"], 2)
        self.assertEqual(df_int_enc.iloc[1]["Data"], "15/09/2026")


if __name__ == "__main__":
    unittest.main()
