import io
import math
import re

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process

# ──────────────────────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="RD — Aplicação de Fontes",
    layout="wide",
    page_icon="📄",
)

st.title("📄 RD — Aplicação de Fontes")
st.caption(
    "Cole o texto do RD e a lista de fontes. "
    "O app executa as 3 etapas automaticamente: "
    "extração de parágrafos → matching fuzzy de fontes → geração do Word com fontes aplicadas."
)
st.markdown("---")


# ──────────────────────────────────────────────────────────────
# HELPERS COMPARTILHADOS
# ──────────────────────────────────────────────────────────────

def ajustar_altura(ws, col=1, largura_coluna=100, altura_por_linha=20):
    """Ajusta altura das linhas baseado no tamanho do texto (replicado de rd_docx_to_excel_paragraphs.py)."""
    col_letter = get_column_letter(col)
    ws.column_dimensions[col_letter].width = largura_coluna
    for row_cells in ws.iter_rows(min_row=1, max_col=col, max_row=ws.max_row, min_col=col):
        for cell in row_cells:
            cell.alignment = Alignment(wrap_text=True)
            if cell.value:
                texto = str(cell.value)
                linhas = 0
                for paragrafo in texto.split("\n"):
                    linhas += math.ceil(len(paragrafo) / (largura_coluna * 1.2))
                ws.row_dimensions[cell.row].height = max(linhas * altura_por_linha, altura_por_linha)


# ──────────────────────────────────────────────────────────────
# ETAPA 1 — Extração de parágrafos (rd_docx_to_excel_paragraphs)
# ──────────────────────────────────────────────────────────────

def extrair_paragrafos(texto_rd: str) -> list[str]:
    """
    Recebe o texto bruto do RD e retorna lista de parágrafos válidos.
    Replica os filtros de rd_docx_to_excel_paragraphs.py:
      - Remove linhas vazias
      - Remove linhas que começam com Note: ou Source:
      - Remove linhas que começam com (número) ex: (1) ...
      - Remove linhas entre aspas curvas duplas "..." 
    """
    paragrafos = []
    for linha in texto_rd.split("\n"):
        linha = linha.strip()
        if not linha:
            continue
        if linha.startswith("Note:") or linha.startswith("Source:"):
            continue
        if re.match(r"^\(\d+\)", linha):
            continue
        if linha.startswith("\u201c") and linha.endswith("\u201d"):
            continue
        if linha.startswith('"') and linha.endswith('"'):
            continue
        paragrafos.append(linha)
    return paragrafos


def paragrafos_para_excel(paragrafos: list[str]) -> bytes:
    """
    Cria um Workbook openpyxl com os parágrafos na coluna A
    e retorna os bytes do arquivo .xlsx.
    Replica a lógica de word_to_excel() em rd_docx_to_excel_paragraphs.py.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Parágrafos"

    for i, texto in enumerate(paragrafos, start=1):
        cell = ws.cell(row=i, column=1, value=texto)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    ajustar_altura(ws, col=1, largura_coluna=100, altura_por_linha=20)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ──────────────────────────────────────────────────────────────
# ETAPA 2 — Fuzzy matching de fontes (rd_match_excel_file)
# ──────────────────────────────────────────────────────────────

def parse_fontes(texto_fontes: str) -> pd.DataFrame:
    """
    Converte o texto de fontes (uma por linha, formato trecho|URL)
    em um DataFrame com colunas [trecho, url].
    """
    linhas = []
    for linha in texto_fontes.strip().split("\n"):
        if "|" in linha:
            partes = linha.split("|", 1)
            trecho = partes[0].strip()
            url = partes[1].strip()
            if trecho and url:
                linhas.append({"trecho": trecho, "url": url})
    return pd.DataFrame(linhas)


def fuzzy_match(paragrafos: list[str], df_fontes: pd.DataFrame, threshold: int = 80) -> list[str]:
    """
    Para cada parágrafo, encontra a melhor correspondência fuzzy
    nos trechos de df_fontes.
    Replica a lógica de fuzzy_match_and_copy() em rd_match_excel_file.py.
    Retorna lista de URLs matched (ou "" se sem match).
    """
    if df_fontes.empty:
        return [""] * len(paragrafos)

    trechos = df_fontes["trecho"].tolist()
    urls_matched = []

    for paragrafo in paragrafos:
        resultado = process.extractOne(paragrafo, trechos, scorer=fuzz.partial_ratio)
        if resultado and resultado[1] >= threshold:
            idx = trechos.index(resultado[0])
            urls_matched.append(df_fontes.iloc[idx]["url"])
        else:
            urls_matched.append("")

    return urls_matched


def matched_para_excel(paragrafos: list[str], urls: list[str]) -> bytes:
    """
    Cria o Excel com col A = parágrafo, col B = URL matched.
    Replica a formatação de fuzzy_match_and_copy() em rd_match_excel_file.py.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Parágrafos"

    for i, (texto, url) in enumerate(zip(paragrafos, urls), start=1):
        cell_a = ws.cell(row=i, column=1, value=texto)
        cell_a.alignment = Alignment(wrap_text=True, vertical="top")
        cell_b = ws.cell(row=i, column=2, value=url)
        cell_b.alignment = Alignment(wrap_text=True, vertical="top")

    ws.column_dimensions[get_column_letter(1)].width = 100
    ws.column_dimensions[get_column_letter(2)].width = 80
    ajustar_altura(ws, col=1, largura_coluna=100, altura_por_linha=20)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ──────────────────────────────────────────────────────────────
# ETAPA 3 — Aplicação das fontes no Word (rd_apply_sources_to_file)
# ──────────────────────────────────────────────────────────────

def gerar_docx(paragrafos: list[str], urls: list[str]) -> bytes:
    """
    Gera um documento Word onde cada parágrafo é seguido,
    quando há URL correspondente, de um parágrafo em vermelho Pt(10)
    com a fonte.
    Replica a lógica de rd_apply_sources_to_file.py.
    """
    doc = Document()

    # Remove o parágrafo vazio padrão que o python-docx cria
    for element in doc.element.body:
        doc.element.body.remove(element)

    mapa_url = {p: u for p, u in zip(paragrafos, urls)}

    for paragrafo in paragrafos:
        # Adiciona o parágrafo original
        p = doc.add_paragraph()
        run = p.add_run(paragrafo)
        run.font.size = Pt(11)

        # Se há URL, adiciona parágrafo em vermelho logo após
        url = mapa_url.get(paragrafo, "")
        if url:
            p_fonte = doc.add_paragraph()
            run_fonte = p_fonte.add_run(url)
            run_fonte.font.color.rgb = RGBColor(255, 0, 0)
            run_fonte.font.size = Pt(10)
            p_fonte.paragraph_format.space_before = Pt(0)
            p_fonte.paragraph_format.space_after = Pt(12)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ──────────────────────────────────────────────────────────────
# PIPELINE PRINCIPAL
# ──────────────────────────────────────────────────────────────

def executar_pipeline(texto_rd: str, texto_fontes: str, threshold: int):
    """Executa as 3 etapas e retorna dict com resultados."""
    resultados = {}

    # ── Etapa 1 ──
    paragrafos = extrair_paragrafos(texto_rd)
    resultados["paragrafos"] = paragrafos
    resultados["n_paragrafos"] = len(paragrafos)
    resultados["xlsx_etapa1"] = paragrafos_para_excel(paragrafos)

    # ── Etapa 2 ──
    df_fontes = parse_fontes(texto_fontes)
    urls = fuzzy_match(paragrafos, df_fontes, threshold=threshold)
    resultados["urls"] = urls
    resultados["n_com_fonte"] = sum(1 for u in urls if u)
    resultados["n_sem_fonte"] = sum(1 for u in urls if not u)
    resultados["xlsx_matched"] = matched_para_excel(paragrafos, urls)

    # ── Etapa 3 ──
    resultados["docx"] = gerar_docx(paragrafos, urls)

    # DataFrame para preview
    resultados["df_preview"] = pd.DataFrame({
        "Parágrafo": paragrafos,
        "Fonte encontrada": urls,
    })

    return resultados


# ──────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────

col_inputs, col_outputs = st.columns([1, 1], gap="large")

with col_inputs:
    st.subheader("📥 Entradas")

    texto_rd = st.text_area(
        "Texto do RD",
        height=380,
        placeholder=(
            "Cole aqui o texto do RD, um parágrafo por linha.\n\n"
            "Linhas começando com Note:, Source: ou (1) serão ignoradas automaticamente."
        ),
        key="texto_rd",
    )

    texto_fontes = st.text_area(
        "Fontes",
        height=260,
        placeholder=(
            "Cole aqui as fontes, uma por linha.\n"
            "Formato: trecho do parágrafo | URL\n\n"
            "Exemplo:\n"
            "The fiscal deficit reached 3.2% of GDP | https://www.gov.br/nota1\n"
            "Inflation remained below target | https://www.bcb.gov.br/nota2"
        ),
        help="Cada linha deve conter o trecho do parágrafo e a URL separados por | (pipe).",
        key="texto_fontes",
    )

    threshold = st.slider(
        "Threshold de similaridade (fuzzy match)",
        min_value=50,
        max_value=100,
        value=80,
        step=5,
        help=(
            "Percentual mínimo de similaridade para considerar um match válido. "
            "Valores menores = mais matches (menos precisos). "
            "Valores maiores = menos matches (mais precisos)."
        ),
        key="threshold",
    )

    processar = st.button("▶️ Processar", type="primary", use_container_width=True)

# ── Execução ──
if processar:
    if not texto_rd.strip():
        st.error("❌ Cole o texto do RD antes de processar.")
    elif not texto_fontes.strip():
        st.warning("⚠️ Nenhuma fonte informada. O Word será gerado sem fontes aplicadas.")
        with st.spinner("Processando..."):
            st.session_state["resultados"] = executar_pipeline(texto_rd, "", threshold)
    else:
        with st.spinner("Processando as 3 etapas..."):
            st.session_state["resultados"] = executar_pipeline(texto_rd, texto_fontes, threshold)

# ── Outputs ──
with col_outputs:
    st.subheader("📤 Resultados")

    if "resultados" not in st.session_state:
        st.info("ℹ️ Preencha as entradas e clique em **▶️ Processar** para ver os resultados aqui.")
    else:
        r = st.session_state["resultados"]

        # Status das etapas
        st.success(f"✅ **Etapa 1** — {r['n_paragrafos']} parágrafos extraídos")
        if r["n_com_fonte"] > 0:
            st.success(f"✅ **Etapa 2** — {r['n_com_fonte']} parágrafos com fonte encontrada")
        else:
            st.warning("⚠️ **Etapa 2** — Nenhuma fonte encontrada. Tente reduzir o threshold.")
        st.success("✅ **Etapa 3** — Documento Word gerado")

        st.markdown("---")

        # Métricas
        m1, m2, m3 = st.columns(3)
        m1.metric("Total de parágrafos", r["n_paragrafos"])
        m2.metric("Com fonte", r["n_com_fonte"])
        m3.metric("Sem fonte", r["n_sem_fonte"])

        st.markdown("---")

        # Preview da tabela de matching
        st.subheader("🔍 Resultado do matching")
        st.dataframe(
            r["df_preview"],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Parágrafo": st.column_config.TextColumn("Parágrafo", width="large"),
                "Fonte encontrada": st.column_config.TextColumn("Fonte encontrada", width="medium"),
            },
        )

        st.markdown("---")

        # Botões de download
        st.subheader("⬇️ Downloads")
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "⬇️ Excel com fontes matched",
                data=r["xlsx_matched"],
                file_name="RD_sourcebook_matched.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_xlsx",
            )
        with d2:
            st.download_button(
                "⬇️ Word com fontes aplicadas",
                data=r["docx"],
                file_name="RD_com_fontes.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="dl_docx",
            )
