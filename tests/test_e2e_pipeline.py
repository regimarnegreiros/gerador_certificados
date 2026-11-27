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
from core.config import obter_configuracao, salvar_curso_json, CursoConfig
from core.cursos_repo import (
    criar_curso,
    definir_modelo_global,
    definir_servidores_global,
    importar_convidados,
    importar_listas_presenca,
)
from core.presenca import gerar_relatorio


class TestE2EPipeline(unittest.TestCase):

    def setUp(self):
        self.temp_raiz = Path(tempfile.mkdtemp())
        self.temp_staging = self.temp_raiz / "_staging"
        self.temp_staging.mkdir(parents=True)

        # 1. Cria curso pela camada de repositório (estrutura global/ + cursos/id/)
        self.temp_curso, cfg = criar_curso("Curso de Teste Unitario", raiz=self.temp_raiz)
        cfg.carga_horaria = 10
        cfg.custo_total_em_reais = 3000.00
        cfg.data_inicio = "01/09/2026"
        cfg.data_fim = "02/09/2026"
        cfg.data_emissao = "02/09/2026"
        cfg.tempo_minimo_presenca_minutos = 15
        cfg.assunto_email = "Certificado - {CURSO}"
        cfg.corpo_email = "Olá {nome}, seu certificado está pronto."
        cfg.campos_editados_manualmente = ["carga_horaria", "data_inicio", "data_fim", "data_emissao"]
        salvar_curso_json(self.temp_curso / "curso.json", cfg)

        # 2. Planilha global de Servidores Excel (.xlsx)
        df_servidores = pd.DataFrame({
            "NOME SERVIDOR": ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA"],
            "UORG": ["SGP/ANM", "DITEC/ANM", "SGP/ANM"],
            "EMAIL": ["alice@anm.gov.br", "bruno@anm.gov.br", "carlos@anm.gov.br"],
        })
        arq_serv = self.temp_staging / "servidores.xlsx"
        df_servidores.to_excel(arq_serv, index=False)
        definir_servidores_global(arq_serv, raiz=self.temp_raiz)

        # 3. Lista de Convocados (.csv)
        convocados_content = (
            "Relatório de Convocados\n"
            "Gerado em 01/09/2026\n"
            "Nome\n"
            "ALICE SILVA <alice@anm.gov.br>\n"
            "BRUNO SOUZA <bruno@anm.gov.br>\n"
            "CARLOS PEREIRA <carlos@anm.gov.br>\n"
        )
        arq_conv = self.temp_staging / "convocados.csv"
        arq_conv.write_text(convocados_content, encoding="utf-8")
        importar_convidados(self.temp_curso, arq_conv)

        # 4. Listas de Presença (.csv diários - formato Teams)
        def gerar_presenca_teams(participantes, data_inicio="9/14/26,8:01:54 AM", duracao="4h 0m 0s", tempo_medio="3h 0m 0s"):
            linhas = [
                "1. Resumo da Reunião",
                f"Total de participantes\t{len(participantes)}",
                f"Hora de início\t{data_inicio}",
                f"Duração da reunião\t{duracao}",
                f"Tempo médio de participação\t{tempo_medio}",
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
        arq_d1 = self.temp_staging / "dia_1.csv"
        arq_d1.write_text(dia1, encoding="utf-16")

        # Encontro 2: Alice 4h (cam=1, mic=1, mao=2, aplauso=1, curtir=1), Carlos 4h (cam=1), Bruno ausente
        dia2 = gerar_presenca_teams([
            ("Alice Silva", "alice@anm.gov.br", "4h 0m 0s", "1", "1", "2", "1", "1"),
            ("Carlos Pereira", "carlos@anm.gov.br", "4h 0m 0s", "1", "", "", "", ""),
        ], data_inicio="9/15/26,8:01:54 AM", duracao="4h 0m 0s")
        arq_d2 = self.temp_staging / "dia_2.csv"
        arq_d2.write_text(dia2, encoding="utf-16")

        importar_listas_presenca(self.temp_curso, [arq_d1, arq_d2])

        # 5. Modelo global de PPTX mínimo
        from pptx import Presentation
        prs = Presentation()
        prs.slides.add_slide(prs.slide_layouts[0])
        arq_modelo = self.temp_staging / "modelo.pptx"
        prs.save(arq_modelo)
        definir_modelo_global(arq_modelo, raiz=self.temp_raiz)

    def tearDown(self):
        shutil.rmtree(self.temp_raiz, ignore_errors=True)

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

        # Valida conteúdo da aba 'Presenças' com títulos mesclados (ID e data)
        df_presencas_excel = pd.read_excel(saida_excel, sheet_name="Presenças")
        self.assertEqual(
            list(df_presencas_excel.columns),
            ["Nome", "Encontro 1 — 14/09/2026", "Unnamed: 2", "Encontro 2 — 15/09/2026", "Unnamed: 4"]
        )
        self.assertEqual(list(df_presencas_excel["Nome"]), ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA"])

        # Valida que as células de cabeçalho dos encontros estão mescladas no Excel
        import openpyxl
        wb_check = openpyxl.load_workbook(saida_excel)
        ws_pres = wb_check["Presenças"]
        merged_faixas = {str(r) for r in ws_pres.merged_cells.ranges}
        self.assertIn("B1:C1", merged_faixas)
        self.assertIn("D1:E1", merged_faixas)
        self.assertEqual(ws_pres["B1"].value, "Encontro 1 — 14/09/2026")
        self.assertEqual(ws_pres["D1"].value, "Encontro 2 — 15/09/2026")

                # Valida células e fórmulas dinâmicas da Aba Certificação no openpyxl
        ws_cert = wb["Certificação"]
        self.assertTrue(str(ws_cert["D2"].value).startswith("=COUNTIFS('Presenças'!B2:"))
        self.assertTrue(str(ws_cert["E2"].value).startswith("=SUMIF('Presenças'!$B$1:"))
        self.assertTrue(str(ws_cert["F2"].value).startswith("=IF('Informações do Curso'!$B$"))
        self.assertTrue(str(ws_cert["G2"].value).startswith("=IF(F2>='Informações do Curso'!$B$"))
        self.assertTrue(str(ws_cert["H2"].value).startswith("=IF(D2>=MAX(1, ROUNDUP('Custos'!$B$"))
        self.assertTrue(str(ws_cert["I2"].value).startswith('=IF(AND(G2="APTO"'))

        # Valida Formatação Condicional nas colunas de aptidão da Certificação
        self.assertGreater(len(ws_cert.conditional_formatting), 0)

        # Valida fórmulas dinâmicas na Aba Interações por participante
        ws_part = wb["Interações por participante"]
        self.assertEqual(ws_part["B2"].value, "='Certificação'!D2")
        self.assertEqual(ws_part["D2"].value, "=IF(B2>0, C2/B2, 0)")
        self.assertEqual(ws_part["F2"].value, "=IF(B2>0, E2/B2, 0)")

        # Valida fórmulas dinâmicas na Aba Interações por encontro
        ws_enc = wb["Interações por encontro"]
        self.assertTrue(str(ws_enc["D2"].value).startswith("=IFERROR(AVERAGEIF('Presenças'!B:B"))
        self.assertTrue(str(ws_enc["E2"].value).startswith("=COUNTIFS('Presenças'!C:C"))

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
            ["ID", "Data", "CH Realizada", "Tempo Médio", "Presentes", "Ligou a câmera", "Ligou o microfone", "Levantou a mão", "Reações"]
        )
        self.assertEqual(len(df_int_enc_excel), 2)
        enc1 = df_int_enc_excel.iloc[0]
        self.assertEqual(enc1["ID"], 1)
        self.assertEqual(enc1["Data"], "14/09/2026")
        self.assertEqual(enc1["CH Realizada"], "04:00:00")
        self.assertEqual(enc1["Tempo Médio"], "03:00:00")
        self.assertEqual(enc1["Presentes"], "2 de 3")
        self.assertEqual(enc1["Ligou a câmera"], "1 de 2")
        self.assertEqual(enc1["Ligou o microfone"], "1 de 2")
        self.assertEqual(enc1["Levantou a mão"], 1)
        self.assertEqual(enc1["Reações"], 5)

        enc2 = df_int_enc_excel.iloc[1]
        self.assertEqual(enc2["ID"], 2)
        self.assertEqual(enc2["Data"], "15/09/2026")
        self.assertEqual(enc2["CH Realizada"], "04:00:00")
        self.assertEqual(enc2["Tempo Médio"], "03:00:00")
        self.assertEqual(enc2["Presentes"], "2 de 3")
        self.assertEqual(enc2["Ligou a câmera"], "2 de 2")
        self.assertEqual(enc2["Ligou o microfone"], "1 de 2")
        self.assertEqual(enc2["Levantou a mão"], 2)
        self.assertEqual(enc2["Reações"], 2)

        # Valida Aba 3: 'Indicadores'
        df_ind_excel = pd.read_excel(saida_excel, sheet_name="Indicadores")
        self.assertEqual(list(df_ind_excel.columns), ["Indicador", "Resultado"])
        mapa_ind = dict(zip(df_ind_excel["Indicador"], df_ind_excel["Resultado"]))
        self.assertEqual(mapa_ind["Carga horária realizada"], "08:00:00")
        self.assertEqual(mapa_ind["Tempo médio de participação"], "06:00:00")

        # Valida células e fórmulas dinâmicas da Aba Indicadores no openpyxl
        ws_ind = wb["Indicadores"]
        mapa_ind_wb = {str(ws_ind.cell(row=r, column=1).value): ws_ind.cell(row=r, column=2).value for r in range(2, ws_ind.max_row + 1)}
        self.assertEqual(mapa_ind_wb["Convidados a participar"], "=MAX(0, COUNTA('Certificação'!A:A) - 1)")
        self.assertTrue(str(mapa_ind_wb["Participantes totais"]).startswith("=COUNTIF("))
        self.assertTrue(str(mapa_ind_wb["Participantes certificados"]).startswith("=COUNTIF("))
        self.assertEqual(mapa_ind_wb["Carga horária realizada"], "08:00:00")
        self.assertTrue(str(mapa_ind_wb["Tempo médio de participação"]).startswith("=IFERROR(AVERAGEIF("))
        self.assertEqual(mapa_ind_wb["Taxa de certificação"], "=IF(B2>0, B4/B2, 0)")
        self.assertTrue(str(mapa_ind_wb["Participações com Câmera"]).startswith("=IF(B3>0, COUNTIF("))
        self.assertTrue(str(mapa_ind_wb["Participações com Microfone"]).startswith("=IF(B3>0, COUNTIF("))
        self.assertTrue(str(mapa_ind_wb["Total de reações"]).startswith("=SUM("))
        self.assertTrue(str(mapa_ind_wb["Total de mãos levantadas"]).startswith("=SUM("))

        # Formatação percentual nas linhas devidas
        for r in [7, 8, 9]:
            self.assertIn("%", ws_ind.cell(row=r, column=2).number_format)

        # Valida Aba 4: 'Custos'
        df_custos_excel = pd.read_excel(saida_excel, sheet_name="Custos")
        self.assertEqual(list(df_custos_excel.columns), ["Item", "Valor"])
        self.assertEqual(len(df_custos_excel), 15)
        mapa_custos = dict(zip(df_custos_excel["Item"], df_custos_excel["Valor"]))
        self.assertEqual(mapa_custos["Custo total"], 3000.0)
        self.assertEqual(mapa_custos["Carga horária"], "10:00:00")
        self.assertEqual(mapa_custos["Carga horária realizada"], "08:00:00")
        self.assertEqual(mapa_custos["Tempo médio de participação realizado"], "06:00:00")

        # Valida células e fórmulas dinâmicas da Aba Custos no openpyxl
        ws_custos = wb["Custos"]
        mapa_custos_wb = {str(ws_custos.cell(row=r, column=1).value): ws_custos.cell(row=r, column=2).value for r in range(2, ws_custos.max_row + 1)}
        self.assertEqual(mapa_custos_wb["Convidados a participar"], "=Indicadores!B2")
        self.assertEqual(mapa_custos_wb["Participantes totais"], "=Indicadores!B3")
        self.assertEqual(mapa_custos_wb["Participantes certificados"], "=Indicadores!B4")
        self.assertEqual(mapa_custos_wb["Custo total"], 3000.0)
        self.assertEqual(mapa_custos_wb["Quantidade de encontros"], "=MAX(0, COUNTA('Interações por encontro'!A:A) - 1)")
        self.assertTrue(str(mapa_custos_wb["Carga horária"]).startswith("='Informações do Curso'!$B$"))
        self.assertEqual(mapa_custos_wb["Carga horária realizada"], "08:00:00")
        self.assertEqual(mapa_custos_wb["Tempo médio de participação realizado"], "=Indicadores!B6")
        self.assertEqual(mapa_custos_wb["Custo por convidado"], "=IF(B2>0, B5/B2, 0)")
        self.assertEqual(mapa_custos_wb["Custo por participante"], "=IF(B3>0, B5/B3, 0)")
        self.assertEqual(mapa_custos_wb["Custo por certificado"], "=IF(B4>0, B5/B4, 0)")
        self.assertEqual(mapa_custos_wb["Custo por encontro"], "=IF(B6>0, B5/B6, 0)")
        self.assertEqual(mapa_custos_wb["Custo por hora prevista"], "=IF(B7*24>0, B5/(B7*24), 0)")
        self.assertEqual(mapa_custos_wb["Custo por hora realizada"], "=IF(B8*24>0, B5/(B8*24), IF(B7*24>0, B5/(B7*24), 0))")
        self.assertEqual(mapa_custos_wb["Custo por hora de participação média"], "=IF(B9*24>0, B5/(B9*24), 0)")

        # Verifica se células de custos possuem formatação de moeda
        for row in range(2, ws_custos.max_row + 1):
            item_txt = str(ws_custos.cell(row=row, column=1).value)
            cel = ws_custos.cell(row=row, column=2)
            if "Custo" in item_txt:
                self.assertIn("R$", cel.number_format)
                if item_txt == "Custo total":
                    self.assertIsInstance(cel.value, (int, float))
                else:
                    self.assertTrue(str(cel.value).startswith("=IF("))

        # Valida conteúdo da aba 'Informações do Curso'
        df_info_excel = pd.read_excel(saida_excel, sheet_name="Informações do Curso")
        self.assertEqual(list(df_info_excel.columns), ["Item", "Configuração", "Observação"])
        self.assertEqual(len(df_info_excel), 11)

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
        self.assertIn(valores_por_item["Carga Horária"], (10, "10 horas", 10.0))
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
        self.assertIn(str(carlos["Encontro 1 — 14/09/2026"]), ("00:00:00", "0:00:00"))
        self.assertEqual(carlos["Unnamed: 2"], "Ausente")
        self.assertIn(str(carlos["Encontro 2 — 15/09/2026"]), ("4:00:00", "04:00:00"))
        self.assertEqual(carlos["Unnamed: 4"], "Presente")

        # 2. Valida Aba 'Interações por encontro': Linhas ordenadas por data cronológica
        df_int_enc = pd.read_excel(saida_excel, sheet_name="Interações por encontro")
        self.assertEqual(df_int_enc.iloc[0]["ID"], 1)
        self.assertEqual(df_int_enc.iloc[0]["Data"], "14/09/2026")
        self.assertEqual(df_int_enc.iloc[1]["ID"], 2)
        self.assertEqual(df_int_enc.iloc[1]["Data"], "15/09/2026")

    def test_pipeline_secao3_recorte_horario_e_linhas_repetidas(self):
        """
        Valida fluxo E2E completo com HORARIO_INICIO e HORARIO_FIM configurados
        e presença apurada pela seção '3. Atividades em Reunião' (MM/DD/AA, HH:MM:SS AM/PM),
        incluindo múltiplas entradas/saídas, 12:xx PM, agrupamento por Email/Nome,
        e recálculo de Tempo Médio recortado.
        """
        raiz_sec3 = Path(tempfile.mkdtemp())
        try:
            pasta, cfg_sec3 = criar_curso("Curso Secao 3", raiz=raiz_sec3)
            cfg_sec3.carga_horaria = 4
            cfg_sec3.custo_total_em_reais = 1000.00
            cfg_sec3.data_inicio = "02/08/2026"
            cfg_sec3.data_fim = "02/08/2026"
            cfg_sec3.data_emissao = "02/08/2026"
            cfg_sec3.ignorar_emails = ["org@anm.gov.br"]
            cfg_sec3.corpo_email = "Olá {nome}, certificado {CURSO}."
            cfg_sec3.campos_editados_manualmente = ["carga_horaria", "data_inicio", "data_fim", "data_emissao"]

            arq_serv = raiz_sec3 / "serv_tmp.xlsx"
            pd.DataFrame({
                "NOME SERVIDOR": ["ALICE SILVA", "BRUNO SOUZA", "CARLOS PEREIRA", "ORGANIZADOR"],
                "UORG": ["SGP", "DITEC", "DIR", "SGP"],
                "EMAIL": ["alice@anm.gov.br", "bruno@anm.gov.br", "carlos@anm.gov.br", "org@anm.gov.br"],
            }).to_excel(arq_serv, index=False)
            definir_servidores_global(arq_serv, raiz=raiz_sec3)

            arq_conv = raiz_sec3 / "conv_tmp.csv"
            arq_conv.write_text(
                "Relatório\nData\nNome\n"
                "ALICE SILVA <alice@anm.gov.br>\n"
                "BRUNO SOUZA <bruno@anm.gov.br>\n"
                "CARLOS PEREIRA <carlos@anm.gov.br>\n"
                "ORGANIZADOR <org@anm.gov.br>\n",
                encoding="utf-8",
            )
            importar_convidados(pasta, arq_conv)

            # Encontro único em 8/02/26 (02/08/2026), janela 08:00 às 12:00:
            # - Alice: 7:45 AM - 12:30 PM -> recortado para 08:00 - 12:00 (4h) -> APTO
            # - Bruno: duas sessões (8:00 AM - 9:00 AM = 1h; 11:00 AM - 12:15 PM -> 11:00 - 12:00 = 1h),
            #          sem email na Seção 3 (agrupado por Nome) -> total 2h (< 75% de 4h = 3h) -> NÃO APTO por CH, Presente no encontro
            # - Carlos: 7:00 AM - 7:50 AM (totalmente fora) -> 0h -> Ausente
            # - Organizador: ignorado por IGNORAR_EMAILS
            arq_enc = raiz_sec3 / "encontro_1.csv"
            arq_enc.write_text(
                "1. Resumo da Reunião\n"
                "Total de participantes\t4\n"
                "Hora de início\t8/02/26, 7:30:00 AM\n"
                "Duração da reunião\t5h 0m 0s\n"
                "Tempo médio de participação\t4h 30m 0s\n"
                "2. Participantes\n"
                "Nome\tEmail\tDuração da Reunião\tCompromisso: Câmera Ligada\tCompromisso: Desativar Mudo\n"
                "Alice Silva\talice@anm.gov.br\t4h 45m 0s\t2\t1\n"
                "Bruno Souza\tbruno@anm.gov.br\t2h 15m 0s\t1\t0\n"
                "Carlos Pereira\tcarlos@anm.gov.br\t50m 0s\t0\t0\n"
                "Organizador\torg@anm.gov.br\t5h 0m 0s\t1\t1\n"
                "3. Atividades em Reunião\n"
                "Nome\tHorário de entrada\tHorário de Saída\tDuração\tEmail\tFunção\n"
                "Alice S.\t8/02/26, 7:45:00 AM\t8/02/26, 12:30:00 PM\t4h 45m 0s\tALICE@ANM.GOV.BR\tParticipante\n"
                "Bruno Souza\t8/02/26, 8:00:00 AM\t8/02/26, 9:00:00 AM\t1h 0m 0s\t\tParticipante\n"
                "Bruno Souza\t8/02/26, 11:00:00 AM\t8/02/26, 12:15:00 PM\t1h 15m 0s\t\tParticipante\n"
                "Carlos Pereira\t8/02/26, 7:00:00 AM\t8/02/26, 7:50:00 AM\t50m 0s\tcarlos@anm.gov.br\tParticipante\n"
                "Organizador\t8/02/26, 7:30:00 AM\t8/02/26, 12:30:00 PM\t5h 0m 0s\torg@anm.gov.br\tOrganizador\n",
                encoding="utf-16",
            )
            salvar_curso_json(pasta / "curso.json", cfg_sec3)
            importar_listas_presenca(pasta, [arq_enc])
            from core.cursos_repo import atualizar_horario_encontro
            atualizar_horario_encontro(pasta, "encontro_1.csv", "08:00", "12:00")

            from pptx import Presentation
            prs = Presentation()
            prs.slides.add_slide(prs.slide_layouts[0])
            arq_mod = raiz_sec3 / "modelo.pptx"
            prs.save(arq_mod)
            definir_modelo_global(arq_mod, raiz=raiz_sec3)

            auditoria = validar_estrutura_curso(pasta)
            config = obter_configuracao(pasta)
            self.assertEqual(config["ENCONTROS"][0]["horario_inicio"], "08:00")
            self.assertEqual(config["ENCONTROS"][0]["horario_fim"], "12:00")

            saida_excel = pasta / "saida" / "relatorio_presencas" / "Relatorio_Secao3.xlsx"
            df_cert = gerar_relatorio(
                caminho_servidores=auditoria["arquivos"]["servidores"],
                caminho_convocados=auditoria["arquivos"]["convidados"],
                arquivos_presenca=auditoria["arquivos"]["presencas"],
                caminho_saida_excel=saida_excel,
                carga_horaria_horas=config["CARGA_HORARIA"],
                tempo_comparecimento_minutos=config["TEMPO_MINIMO_PRESENCA_MINUTOS"],
                percentual_minimo_carga=config["PERCENTUAL_MINIMO_CARGA"],
                percentual_minimo_encontros=config["PERCENTUAL_MINIMO_ENCONTROS"],
                config=config,
            )

            self.assertEqual(len(df_cert), 3)
            self.assertNotIn("ORGANIZADOR", list(df_cert["Nome"]))

            alice = df_cert[df_cert["Nome"] == "ALICE SILVA"].iloc[0]
            self.assertEqual(alice["CH Total"], "4:00:00")
            self.assertEqual(alice["Apto por CH"], "APTO")
            self.assertEqual(alice["Apto em Ambos"], "APTO")

            bruno = df_cert[df_cert["Nome"] == "BRUNO SOUZA"].iloc[0]
            self.assertEqual(bruno["CH Total"], "2:00:00")
            self.assertEqual(bruno["Presenças"], 1)
            self.assertEqual(bruno["Apto por CH"], "NÃO APTO")

            carlos = df_cert[df_cert["Nome"] == "CARLOS PEREIRA"].iloc[0]
            self.assertEqual(carlos["CH Total"], "0:00:00")
            self.assertEqual(carlos["Presenças"], 0)
            self.assertEqual(carlos["Apto em Ambos"], "NÃO APTO")

            # Tempo Médio recalculado: média de (4h + 2h) = 3h -> 03:00:00 (não 04:30:00 da Seção 1)
            df_enc = pd.read_excel(saida_excel, sheet_name="Interações por encontro")
            self.assertEqual(df_enc.iloc[0]["Data"], "02/08/2026")
            self.assertEqual(df_enc.iloc[0]["Tempo Médio"], "03:00:00")
            self.assertEqual(df_enc.iloc[0]["Presentes"], "2 de 3")

            df_ind = pd.read_excel(saida_excel, sheet_name="Indicadores")
            mapa_ind = dict(zip(df_ind["Indicador"], df_ind["Resultado"]))
            self.assertEqual(mapa_ind["Tempo médio de participação"], "03:00:00")

            df_info = pd.read_excel(saida_excel, sheet_name="Informações do Curso")
            self.assertEqual(len(df_info), 11)
            mapa_info = dict(zip(df_info["Item"], df_info["Configuração"]))
            self.assertEqual(mapa_info["Horário de Início"], "08:00")
            self.assertEqual(mapa_info["Horário de Fim"], "12:00")

            # Verifica preservação das referências de fórmula ($B$3, $B$8, $B$9) e interações da Seção 2
            import openpyxl
            wb_check = openpyxl.load_workbook(saida_excel)
            ws_cert = wb_check["Certificação"]
            self.assertEqual(ws_cert["F2"].value, "=IF('Informações do Curso'!$B$3>0, (E2*24)/'Informações do Curso'!$B$3, 0)")
            self.assertEqual(ws_cert["G2"].value, '=IF(F2>=\'Informações do Curso\'!$B$8, "APTO", "NÃO APTO")')
            self.assertTrue(str(ws_cert["H2"].value).startswith("=IF(D2>=MAX(1, ROUNDUP('Custos'!$B$6 * 'Informações do Curso'!$B$9, 0))"))

            df_part = pd.read_excel(saida_excel, sheet_name="Interações por participante")
            alice_int = df_part[df_part["Participante"] == "ALICE SILVA"].iloc[0]
            self.assertEqual(alice_int["Câmera"], 1)
            self.assertEqual(alice_int["Microfone"], 1)

            df_custos = pd.read_excel(saida_excel, sheet_name="Custos")
            mapa_custos = dict(zip(df_custos["Item"], df_custos["Valor"]))
            self.assertEqual(mapa_custos["Tempo médio de participação realizado"], "03:00:00")
        finally:
            shutil.rmtree(raiz_sec3, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()