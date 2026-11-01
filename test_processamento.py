from decimal import Decimal
import unittest
from io import BytesIO

from docx import Document
from openpyxl import load_workbook

from processamento import gerar_peticao, gerar_planilha, nome_seguro, parsear_rubricas


class ProcessamentoTests(unittest.TestCase):
    def test_parseia_linhas_com_valor_brasileiro(self) -> None:
        rubricas = parsear_rubricas(
            "0001 VENCIMENTO BÁSICO 3.500,00\n"
            "12345 CONTRIBUICAO PREVIDENCIARIA -250,50\n"
            "texto de cabeçalho sem valor"
        )

        self.assertEqual(len(rubricas), 2)
        self.assertEqual(rubricas[0]["valor"], Decimal("3500.00"))
        self.assertEqual(rubricas[1]["valor"], Decimal("-250.50"))

    def test_planilha_contem_formula_de_total_e_rubricas(self) -> None:
        conteudo = gerar_planilha(
            [
                {
                    "competencia": "2021-01",
                    "arquivo": "01_2021.pdf",
                    "codigo": "0001",
                    "descricao": "VENCIMENTO BÁSICO",
                    "valor": Decimal("3500.00"),
                }
            ]
        )
        workbook = load_workbook(BytesIO(conteudo), data_only=False)

        self.assertEqual(workbook["Resumo"]["B2"].value, "=SUM(Rubricas!E2:E2)")
        self.assertEqual(workbook["Rubricas"]["E2"].value, 3500)

    def test_peticao_substitui_placeholders_do_modelo(self) -> None:
        modelo = Document()
        modelo.add_paragraph("{{NOME_CLIENTE}} - {{MATRICULA}}")
        modelo.add_paragraph("Total: {{VALOR_TOTAL_DIFERENCAS}}")
        modelo.add_paragraph("{{TABELA_CALCULOS}}")
        buffer = BytesIO()
        modelo.save(buffer)

        resultado = gerar_peticao(
            "Ana Silva",
            "9876",
            [
                {
                    "competencia": "2021-01",
                    "arquivo": "01_2021.pdf",
                    "codigo": "0001",
                    "descricao": "VENCIMENTO BÁSICO",
                    "valor": Decimal("3500.00"),
                }
            ],
            buffer.getvalue(),
        )
        documento = Document(BytesIO(resultado))

        self.assertIn("Ana Silva - 9876", documento.paragraphs[0].text)
        self.assertIn("R$ 3.500,00", documento.paragraphs[1].text)
        self.assertIn("VENCIMENTO BÁSICO", documento.paragraphs[2].text)

    def test_nome_seguro_remove_caracteres_invalidos(self) -> None:
        self.assertEqual(nome_seguro("João/PMBA"), "Joao_PMBA")


if __name__ == "__main__":
    unittest.main()
