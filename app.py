from __future__ import annotations

from pathlib import Path
from zipfile import BadZipFile

import pandas as pd
import streamlit as st

from processamento import (
    gerar_planilha,
    gerar_peticao,
    ler_pdf,
    nome_seguro,
    parsear_rubricas,
)
from docx.opc.exceptions import PackageNotFoundError
from pypdf.errors import PdfReadError


MESES = {
    1: "Janeiro",
    2: "Fevereiro",
    3: "Março",
    4: "Abril",
    5: "Maio",
    6: "Junho",
    7: "Julho",
    8: "Agosto",
    9: "Setembro",
    10: "Outubro",
    11: "Novembro",
    12: "Dezembro",
}


def periodo_do_nome(nome: str) -> tuple[str, str]:
    import re

    correspondencias = (
        re.search(r"(?<!\d)(20\d{2})[-_. ](0?[1-9]|1[0-2])(?!\d)", nome),
        re.search(r"(?<!\d)(0?[1-9]|1[0-2])[-_. ](20\d{2})(?!\d)", nome),
    )
    if correspondencias[0]:
        ano, mes = correspondencias[0].groups()
        return str(int(mes)), ano
    if correspondencias[1]:
        mes, ano = correspondencias[1].groups()
        return str(int(mes)), ano
    return "", ""


st.set_page_config(
    page_title="Holerites PMBA | Processador local",
    page_icon="📄",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { max-width: 1180px; padding-top: 2rem; }
    [data-testid="stMetric"] {
        background: #f5f7fb; border: 1px solid #e6eaf1;
        padding: 1rem; border-radius: 12px;
    }
    div[data-testid="stFileUploader"] {
        border: 1px dashed #9aa8bd; border-radius: 14px; padding: .5rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Processador de holerites")
st.caption("PMBA · processamento local · arquivos só são copiados ao gerar o pacote")

with st.container(border=True):
    st.markdown("#### 1 · Identificação do caso")
    dados_col1, dados_col2 = st.columns(2)
    with dados_col1:
        nome_cliente = st.text_input("Nome do policial", placeholder="Nome completo")
    with dados_col2:
        matricula = st.text_input("Matrícula", placeholder="Ex.: 12345678")

with st.container(border=True):
    st.markdown("#### 2 · Holerites em PDF")
    st.write("Arraste os arquivos para a área abaixo ou selecione vários PDFs.")
    arquivos = st.file_uploader(
        "Selecionar holerites",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed",
    )

documentos: list[dict] = []
falhas: list[str] = []
if arquivos:
    with st.spinner("Lendo o texto dos PDFs localmente..."):
        for indice, arquivo in enumerate(arquivos):
            try:
                texto = ler_pdf(arquivo.getvalue())
                rubricas = parsear_rubricas(texto)
                mes, ano = periodo_do_nome(arquivo.name)
                documentos.append(
                    {
                        "indice": indice,
                        "arquivo": arquivo.name,
                        "conteudo": arquivo.getvalue(),
                        "texto": texto,
                        "rubricas": rubricas,
                        "mes": mes,
                        "ano": ano,
                    }
                )
            except (OSError, ValueError, PdfReadError) as erro:
                falhas.append(f"{arquivo.name}: {erro}")

    for falha in falhas:
        st.error(f"Não foi possível ler o PDF — {falha}")

    if documentos:
        st.markdown("#### 3 · Confira as competências")
        st.caption(
            "Mês e ano são sugeridos pelo nome do arquivo. Corrija ou preencha "
            "os campos antes de gerar o pacote."
        )
        dados_periodos = pd.DataFrame(
            [
                {
                    "PDF": f"{doc['indice'] + 1} · {doc['arquivo']}",
                    "Mês (1–12)": doc["mes"],
                    "Ano": doc["ano"],
                }
                for doc in documentos
            ]
        )
        periodos_editados = st.data_editor(
            dados_periodos,
            hide_index=True,
            use_container_width=True,
            disabled=["PDF"],
            column_config={
                "Mês (1–12)": st.column_config.TextColumn(
                    "Mês (1–12)", help="Número do mês, de 1 a 12"
                ),
                "Ano": st.column_config.TextColumn(
                    "Ano", help="Ano com quatro dígitos"
                ),
            },
            key="periodos_holerites",
        )

        linhas: list[dict] = []
        for indice, doc in enumerate(documentos):
            try:
                mes = int(str(periodos_editados.iloc[indice]["Mês (1–12)"]).strip())
                ano = int(str(periodos_editados.iloc[indice]["Ano"]).strip())
                if mes not in MESES or ano < 1900 or ano > 2100:
                    raise ValueError
                doc["mes_num"] = mes
                doc["ano_num"] = ano
                doc["competencia"] = f"{ano}-{mes:02d}"
                for rubrica in doc["rubricas"]:
                    linhas.append(
                        {
                            "competencia": doc["competencia"],
                            "arquivo": doc["arquivo"],
                            **rubrica,
                        }
                    )
            except (TypeError, ValueError):
                st.warning(f"Informe uma competência válida para {doc['arquivo']}.")
                doc["competencia"] = ""

        if linhas:
            tabela = pd.DataFrame(linhas)
            st.markdown("#### Prévia das rubricas reconhecidas")
            st.dataframe(
                tabela.rename(
                    columns={
                        "competencia": "Competência",
                        "arquivo": "PDF de origem",
                        "codigo": "Código",
                        "descricao": "Descrição",
                        "valor": "Valor (R$)",
                    }
                ),
                hide_index=True,
                use_container_width=True,
            )
        else:
            st.info(
                "Nenhuma rubrica financeira foi reconhecida. Verifique se os PDFs "
                "contêm texto selecionável e se as linhas seguem o padrão "
                "código, descrição e valor."
            )

        st.markdown("#### 4 · Destino dos arquivos")
        destino = st.text_input(
            "Pasta de saída",
            value=str(Path.cwd() / "DADOS_PROCESSADOS"),
            help="Escolha uma pasta local, inclusive uma pasta sincronizada pelo Google Drive para Desktop.",
        )
        modelo = st.file_uploader(
            "Modelo de petição .docx (opcional)",
            type=["docx"],
            help=(
                "Placeholders reconhecidos: {{NOME_CLIENTE}}, {{MATRICULA}}, "
                "{{VALOR_TOTAL_DIFERENCAS}} e {{TABELA_CALCULOS}}."
            ),
            key="modelo_docx",
        )

        duplicadas = [
            d["competencia"]
            for d in documentos
            if d.get("competencia")
            and sum(x.get("competencia") == d["competencia"] for x in documentos) > 1
        ]
        if duplicadas:
            st.error(
                "Há mais de um PDF na mesma competência. Ajuste os períodos para "
                "evitar sobrescrever documentos."
            )

        pode_gerar = bool(nome_cliente.strip() and matricula.strip() and destino.strip())
        pode_gerar = pode_gerar and all(d.get("competencia") for d in documentos)
        pode_gerar = pode_gerar and not duplicadas and not falhas

        if st.button(
            "Gerar pacote do caso",
            type="primary",
            use_container_width=True,
            disabled=not pode_gerar,
        ):
            try:
                pasta_caso = (
                    Path(destino).expanduser()
                    / f"{nome_seguro(matricula)} - {nome_seguro(nome_cliente)}"
                )
                caminhos_pdf = [
                    pasta_caso / str(doc["ano_num"]) / f"{doc['mes_num']:02d}_{doc['ano_num']}.pdf"
                    for doc in documentos
                ]
                caminho_planilha = pasta_caso / "Planilha_Calculos_Consolidada.xlsx"
                caminho_peticao = pasta_caso / "Peticao_Inicial_Calculada.docx"
                caminhos_saida = [*caminhos_pdf, caminho_planilha, caminho_peticao]
                existentes = [str(caminho) for caminho in caminhos_saida if caminho.exists()]
                if existentes:
                    st.error(
                        "Já existem arquivos para estas competências. Escolha outra "
                        "pasta de saída para preservar os documentos anteriores."
                    )
                else:
                    planilha = gerar_planilha(linhas)
                    peticao = gerar_peticao(
                        nome_cliente=nome_cliente.strip(),
                        matricula=matricula.strip(),
                        linhas=linhas,
                        modelo_bytes=modelo.getvalue() if modelo else None,
                    )
                    for doc, caminho in zip(documentos, caminhos_pdf):
                        caminho.parent.mkdir(parents=True, exist_ok=True)
                        caminho.write_bytes(doc["conteudo"])

                    pasta_caso.mkdir(parents=True, exist_ok=True)
                    caminho_planilha.write_bytes(planilha)
                    caminho_peticao.write_bytes(peticao)

                    st.success(f"Pacote gerado em: {pasta_caso}")
                    st.info(
                        "Revise todos os dados e documentos. O protótipo soma as "
                        "rubricas reconhecidas, mas não calcula diferenças jurídicas "
                        "nem substitui conferência profissional."
                    )
                    resultado_col1, resultado_col2 = st.columns(2)
                    with resultado_col1:
                        st.download_button(
                            "Baixar planilha XLSX",
                            data=planilha,
                            file_name=caminho_planilha.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True,
                        )
                    with resultado_col2:
                        st.download_button(
                            "Baixar petição DOCX",
                            data=peticao,
                            file_name=caminho_peticao.name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True,
                        )
            except (OSError, ValueError, BadZipFile, PackageNotFoundError) as erro:
                st.error(f"Falha ao gerar o pacote: {erro}")

st.divider()
st.caption(
    "Protótipo sem OCR e sem envio de dados para Gemini ou outros serviços externos. "
    "PDFs digitalizados como imagem não terão texto extraído."
)
