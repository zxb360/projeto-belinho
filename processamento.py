from __future__ import annotations

import io
import re
import unicodedata
from decimal import Decimal
from typing import Any

from docx import Document
from docx.shared import Inches
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


PADRAO_RUBRICA = re.compile(
    r"^\s*(?P<codigo>\d{3,6})\s+"
    r"(?P<descricao>.*?)\s+"
    r"(?:R\$\s*)?(?P<valor>-?(?:\d{1,3}(?:\.\d{3})+|\d+),\d{2})\s*$",
    re.IGNORECASE,
)


def ler_pdf(conteudo: bytes) -> str:
    leitor = PdfReader(io.BytesIO(conteudo))
    if leitor.is_encrypted:
        raise ValueError("o arquivo está protegido por senha")
    texto = "\n".join(pagina.extract_text() or "" for pagina in leitor.pages)
    if not texto.strip():
        raise ValueError("não foi encontrado texto selecionável (PDF pode ser digitalizado)")
    return texto


def parsear_rubricas(texto: str) -> list[dict[str, Any]]:
    rubricas = []
    for linha in texto.splitlines():
        correspondencia = PADRAO_RUBRICA.match(linha.strip())
        if correspondencia:
            valor = Decimal(
                correspondencia.group("valor").replace(".", "").replace(",", ".")
            )
            rubricas.append(
                {
                    "codigo": correspondencia.group("codigo"),
                    "descricao": correspondencia.group("descricao").strip(),
                    "valor": valor,
                }
            )
    return rubricas


def nome_seguro(valor: str) -> str:
    normalizado = unicodedata.normalize("NFKD", valor)
    sem_acentos = "".join(
        caractere for caractere in normalizado if not unicodedata.combining(caractere)
    )
    nome = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", sem_acentos).strip(" .")
    if not nome:
        raise ValueError("o nome da pasta não pode ficar vazio")
    return nome


def gerar_planilha(linhas: list[dict[str, Any]]) -> bytes:
    pasta = Workbook()
    resumo = pasta.active
    resumo.title = "Resumo"
    resumo.append(["Resumo das rubricas reconhecidas"])
    ultima_linha = max(2, len(linhas) + 1)
    resumo.append(["Total identificado (R$)", f"=SUM(Rubricas!E2:E{ultima_linha})"])
    resumo["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    resumo["A1"].fill = PatternFill("solid", fgColor="17365D")
    resumo["B2"].number_format = '"R$" #,##0.00'
    resumo.column_dimensions["A"].width = 30
    resumo.column_dimensions["B"].width = 24

    planilha = pasta.create_sheet("Rubricas")
    planilha.append(["Competência", "PDF de origem", "Código", "Descrição", "Valor (R$)"])
    for linha in linhas:
        planilha.append(
            [
                linha["competencia"],
                linha["arquivo"],
                linha["codigo"],
                linha["descricao"],
                float(linha["valor"]),
            ]
        )
    for celula in planilha[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="17365D")
    for linha in planilha.iter_rows(min_row=2, min_col=5, max_col=5):
        linha[0].number_format = '"R$" #,##0.00'
    for coluna, largura in enumerate((16, 36, 12, 42, 18), start=1):
        planilha.column_dimensions[get_column_letter(coluna)].width = largura
    planilha.freeze_panes = "A2"
    planilha.auto_filter.ref = planilha.dimensions

    buffer = io.BytesIO()
    pasta.save(buffer)
    return buffer.getvalue()


def _formatar_moeda(valor: Decimal) -> str:
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _substituir_paragrafo(paragrafo: Any, substituicoes: dict[str, str]) -> None:
    texto = paragrafo.text
    novo_texto = texto
    for marcador, valor in substituicoes.items():
        novo_texto = novo_texto.replace(marcador, valor)
    if novo_texto != texto:
        if paragrafo.runs:
            paragrafo.runs[0].text = novo_texto
            for run in paragrafo.runs[1:]:
                run.text = ""
        else:
            paragrafo.add_run(novo_texto)


def gerar_peticao(
    nome_cliente: str,
    matricula: str,
    linhas: list[dict[str, Any]],
    modelo_bytes: bytes | None = None,
) -> bytes:
    total = sum((linha["valor"] for linha in linhas), Decimal("0"))
    tabela_texto = "\n".join(
        f"{linha['competencia']} | {linha['codigo']} - {linha['descricao']} | "
        f"{_formatar_moeda(linha['valor'])}"
        for linha in linhas
    ) or "Nenhuma rubrica financeira reconhecida."
    substituicoes = {
        "{{NOME_CLIENTE}}": nome_cliente,
        "{{MATRICULA}}": matricula,
        "{{VALOR_TOTAL_DIFERENCAS}}": _formatar_moeda(total),
        "{{TABELA_CALCULOS}}": tabela_texto,
    }

    if modelo_bytes:
        documento = Document(io.BytesIO(modelo_bytes))
        for paragrafo in documento.paragraphs:
            _substituir_paragrafo(paragrafo, substituicoes)
        for tabela in documento.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for paragrafo in celula.paragraphs:
                        _substituir_paragrafo(paragrafo, substituicoes)
    else:
        documento = Document()
        documento.add_heading("Petição inicial — minuta para revisão", 0)
        documento.add_paragraph(f"Parte autora: {nome_cliente}")
        documento.add_paragraph(f"Matrícula: {matricula}")
        documento.add_paragraph(
            "Valor total identificado nas rubricas: " + _formatar_moeda(total)
        )
        documento.add_paragraph(
            "Minuta de apoio. Revise fatos, fundamentos, documentos e cálculos antes "
            "de utilizar ou protocolar."
        )
        documento.add_heading("Rubricas reconhecidas", level=1)
        tabela = documento.add_table(rows=1, cols=4)
        tabela.style = "Light Shading Accent 1"
        cabecalhos = ("Competência", "Código", "Descrição", "Valor")
        for celula, titulo in zip(tabela.rows[0].cells, cabecalhos):
            celula.text = titulo
        for linha in linhas:
            celulas = tabela.add_row().cells
            valores = (
                linha["competencia"],
                linha["codigo"],
                linha["descricao"],
                _formatar_moeda(linha["valor"]),
            )
            for celula, valor in zip(celulas, valores):
                celula.text = str(valor)
        documento.sections[0].left_margin = Inches(1)
        documento.sections[0].right_margin = Inches(1)

    buffer = io.BytesIO()
    documento.save(buffer)
    return buffer.getvalue()
