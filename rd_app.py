import io
import math
import re

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_COLOR_INDEX
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process

# ──────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="RD → 18-K | Aplicação de Fontes",
    layout="wide",
    page_icon="📄",
)

st.title("📄 RD → 18-K | Aplicação de Fontes")
st.caption(
    "Pipeline completo: extração de fontes do RD → Excel acumulado → aplicação no 18-K"
)
st.markdown("---")


# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────

def is_red_paragraph(paragraph) -> bool:
    """Retorna True se o parágrafo está em vermelho (fonte vermelha em todos os runs)."""
    runs_com_texto = [r for r in paragraph.runs if r.text.strip()]
    if not runs_com_texto:
        return False
    for run in runs_com_texto:
        try:
            rgb = run.font.color.rgb
            if rgb != RGBColor(0xFF, 0x00, 0x00):
                return False
        except Exception:
            return False
    return True


def is_paragrafo_valido(paragraph) -> bool:
    """Aplica os filtros de exclusão do script original."""
    txt = paragraph.text.strip()
    if not txt:
        return False
    if txt.startswith("Note:") or txt.startswith("Source:"):
        return False
    if re.match(r"^\(\d+\)", txt):
        return False
    if txt.startswith("\u201c") and txt.endswith("\u201d"):
        return False
    if txt.startswith('\u201c') and txt.endswith('\u201d'):
        return False
    return True


def ajustar_altura(ws, col=1, largura_coluna=100, altura_por_linha=20):
    """Ajusta altura das linhas proporcional ao tamanho do texto."""
    col_letter = get_column_letter(col)
    ws.column_dimensions[col_letter].width = largura_coluna
    for row_cells in ws.iter_rows(min_row=1, max_col=col, max_row=ws.max_row, min_col=col):
        for cell in row_cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cell.value:
                texto = str(cell.value)
                linhas = sum(
                    math.ceil(len(p) / (largura_coluna * 1.2)) or 1
                    for p in texto.split("\n")
                )
                ws.row_dimensions[cell.row].height = max(linhas * altura_por_linha, altura_por_linha)


# ──────────────────────────────────────────────────────────────
# ETAPA 1 — Extrair pares (parágrafo, fonte vermelha) do .docx
# ──────────────────────────────────────────────────────────────

def extrair_pares_docx(docx_bytes: bytes) -> pd.DataFrame:
    """
    Lê o .docx e extrai pares (parágrafo normal → fonte vermelha logo abaixo).
    Retorna DataFrame com colunas [paragrafo, fonte].
    """
    doc = Document(io.BytesIO(docx_bytes))
    paragrafos_doc = doc.paragraphs

    pares = []
    i = 0
    while i < len(paragrafos_doc):
        p = paragrafos_doc[i]

        # Pula parágrafos vermelhos isolados (sem parágrafo normal antes)
        if is_red_paragraph(p):
            i += 1
            continue

        # Verifica se é um parágrafo válido (não vazio, não heading, etc.)
        if not is_paragrafo_valido(p):
            i += 1
            continue

        texto = p.text.strip()
        fonte = ""

        # Verifica se o próximo parágrafo é vermelho (fonte deste parágrafo)
        if i + 1 < len(paragrafos_doc) and is_red_paragraph(paragrafos_doc[i + 1]):
            fonte = paragrafos_doc[i + 1].text.strip()
            i += 2  # pula o parágrafo de fonte
        else:
            i += 1

        pares.append({"paragrafo": texto, "fonte": fonte})

    return pd.DataFrame(pares, columns=["paragrafo", "fonte"])


# ──────────────────────────────────────────────────────────────
# ETAPA 2 — Exportar e acumular Excel
# ──────────────────────────────────────────────────────────────

def exportar_excel(df: pd.DataFrame) -> bytes:
    """Exporta DataFrame [paragrafo, fonte] para bytes de .xlsx formatado."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Fontes"

    # Cabeçalho
    ws.cell(row=1, column=1, value="Parágrafo").alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=1, column=2, value="Fonte").alignment = Alignment(wrap_text=True, vertical="top")

    for i, row in enumerate(df.itertuples(index=False), start=2):
        ws.cell(row=i, column=1, value=row.paragrafo).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=i, column=2, value=row.fonte).alignment = Alignment(wrap_text=True, vertical="top")

    ws.column_dimensions[get_column_letter(1)].width = 100
    ws.column_dimensions[get_column_letter(2)].width = 80
    ajustar_altura(ws, col=1, largura_coluna=100, altura_por_linha=20)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def acumular_excel(df_novo: pd.DataFrame, excel_bytes_acumulado: bytes | None) -> bytes:
    """
    Combina o novo DataFrame com o Excel acumulado existente.
    Remove duplicatas exatas em 'paragrafo' (mantém o mais recente = df_novo).
    """
    if excel_bytes_acumulado is not None:
        df_acumulado = pd.read_excel(
            io.BytesIO(excel_bytes_acumulado),
            header=0,
            engine="openpyxl"
        )
        df_acumulado.columns = ["paragrafo", "fonte"]
        df_acumulado = df_acumulado.astype(str).replace("nan", "")
        # Novo em cima, acumulado embaixo — drop_duplicates mantém o primeiro (novo)
        df_merged = pd.concat([df_novo, df_acumulado], ignore_index=True)
        df_merged = df_merged.drop_duplicates(subset=["paragrafo"], keep="first")
    else:
        df_merged = df_novo.copy()

    return exportar_excel(df_merged), df_merged


# ──────────────────────────────────────────────────────────────
# ETAPA 3 — Fuzzy match parágrafo do Word alvo × Excel acumulado
# ──────────────────────────────────────────────────────────────

def extrair_paragrafos_alvo(docx_bytes: bytes) -> list:
    """
    Extrai parágrafos do Word alvo (18-K / RD acumulado).
    Retorna lista de dicts com {texto, style_name, is_red}.
    Preserva TODOS os parágrafos (incluindo headings) para manter estrutura.
    """
    doc = Document(io.BytesIO(docx_bytes))
    resultado = []
    for p in doc.paragraphs:
        resultado.append({
            "texto": p.text,
            "style_name": p.style.name,
            "is_red": is_red_paragraph(p),
        })
    return resultado


def fuzzy_match_paragrafos(
    paragrafos: list,
    df_acumulado: pd.DataFrame,
    threshold: int = 80,
) -> dict:
    """
    Faz fuzzy match entre parágrafos do Word alvo e col 'paragrafo' do Excel acumulado.
    Retorna dict {texto_paragrafo: fonte_matched ou ""}.
    """
    if df_acumulado.empty:
        return {p["texto"]: "" for p in paragrafos}

    trechos = df_acumulado["paragrafo"].tolist()
    fontes  = df_acumulado["fonte"].tolist()

    mapa = {}
    for p in paragrafos:
        txt = p["texto"].strip()
        if not txt or p["is_red"] or not is_paragrafo_valido_texto(txt):
            mapa[txt] = ""
            continue
        resultado = process.extractOne(txt, trechos, scorer=fuzz.partial_ratio)
        if resultado and resultado[1] >= threshold:
            idx = trechos.index(resultado[0])
            mapa[txt] = fontes[idx] if fontes[idx] else ""
        else:
            mapa[txt] = ""
    return mapa


def is_paragrafo_valido_texto(txt: str) -> bool:
    """Versão texto-only dos filtros de exclusão."""
    if not txt:
        return False
    if txt.startswith("Note:") or txt.startswith("Source:"):
        return False
    if re.match(r"^\(\d+\)", txt):
        return False
    return True


# ──────────────────────────────────────────────────────────────
# ETAPA 4 — Gerar Word com fontes aplicadas
# ──────────────────────────────────────────────────────────────

def gerar_docx_com_fontes(
    docx_bytes_alvo: bytes,
    paragrafos_info: list,
    mapa_fontes: dict,
) -> bytes:
    """
    Reconstrói o Word alvo adicionando parágrafos vermelhos
    com a fonte logo após cada parágrafo que teve match.
    Preserva a estrutura e estilos do documento original.
    """
    doc_original = Document(io.BytesIO(docx_bytes_alvo))
    doc_novo = Document(io.BytesIO(docx_bytes_alvo))

    # Remove todos os parágrafos do doc_novo e reconstrói
    # (mais seguro do que modificar in-place)
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    from copy import deepcopy

    body = doc_novo.element.body
    # Remove todos os filhos do body exceto sectPr (configurações de página)
    children = list(body)
    for child in children:
        if child.tag != qn("w:sectPr"):
            body.remove(child)

    # Reconstrói inserindo os parágrafos originais + fontes
    sect_pr = body.find(qn("w:sectPr"))

    for p_orig in doc_original.paragraphs:
        txt = p_orig.text.strip()

        # Copia o parágrafo original via XML
        p_copy = deepcopy(p_orig._element)
        if sect_pr is not None:
            body.insert(list(body).index(sect_pr), p_copy)
        else:
            body.append(p_copy)

        # Se há fonte e o parágrafo é válido (não vermelho, não vazio)
        fonte = mapa_fontes.get(txt, "")
        if fonte and not is_red_paragraph(p_orig) and txt:
            # Cria parágrafo vermelho
            p_fonte = OxmlElement("w:p")
            pPr = OxmlElement("w:pPr")
            spacing = OxmlElement("w:spacing")
            spacing.set(qn("w:before"), "0")
            spacing.set(qn("w:after"), "170")  # ~Pt(12)
            pPr.append(spacing)
            p_fonte.append(pPr)

            r = OxmlElement("w:r")
            rPr = OxmlElement("w:rPr")
            color = OxmlElement("w:color")
            color.set(qn("w:val"), "FF0000")
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), "20")  # Pt(10) = 20 half-points
            rPr.append(color)
            rPr.append(sz)
            r.append(rPr)

            t = OxmlElement("w:t")
            t.text = fonte
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.append(t)
            p_fonte.append(r)

            if sect_pr is not None:
                body.insert(list(body).index(sect_pr), p_fonte)
            else:
                body.append(p_fonte)

    buf = io.BytesIO()
    doc_novo.save(buf)
    buf.seek(0)
    return buf.read()


# ──────────────────────────────────────────────────────────────
# UI — ABAS
# ──────────────────────────────────────────────────────────────

tab1, tab2, tab3 = st.tabs([
    "📥 Etapa 1 & 2 — Extrair e Acumular",
    "🔗 Etapa 3 & 4 — Aplicar no 18-K",
    "📊 Histórico Acumulado",
])


# ══════════════════════════════════════════════════════════════
# TAB 1
# ══════════════════════════════════════════════════════════════
with tab1:
    st.subheader("Etapa 1 — Extrair parágrafos e fontes do RD")
    st.caption(
        "Faça upload do documento RD (.docx). "
        "O app identifica automaticamente os parágrafos em vermelho "
        "como fontes do parágrafo anterior."
    )

    col_in1, col_out1 = st.columns([1, 1], gap="large")

    with col_in1:
        rd_file = st.file_uploader(
            "📄 Documento RD (.docx)",
            type=["docx"],
            key="rd_upload",
            help="O documento deve conter os parágrafos normais seguidos das fontes em vermelho.",
        )
        acum_file_1 = st.file_uploader(
            "📊 Excel acumulado — opcional (para combinar com fontes anteriores)",
            type=["xlsx"],
            key="acum_upload_1",
            help="Se informado, as novas fontes serão combinadas com as anteriores.",
        )

        btn_extrair = st.button("▶️ Extrair e Acumular", type="primary", use_container_width=True, key="btn_extrair")

    if btn_extrair:
        if rd_file is None:
            st.error("❌ Faça upload do documento RD antes de processar.")
        else:
            with st.spinner("Extraindo parágrafos e fontes..."):
                rd_bytes = rd_file.read()
                df_novo = extrair_pares_docx(rd_bytes)

                acum_bytes = acum_file_1.read() if acum_file_1 else None
                xlsx_acumulado_bytes, df_acumulado = acumular_excel(df_novo, acum_bytes)
                xlsx_novo_bytes = exportar_excel(df_novo)

                st.session_state["df_novo_tab1"]       = df_novo
                st.session_state["df_acumulado_tab1"]  = df_acumulado
                st.session_state["xlsx_novo_bytes"]    = xlsx_novo_bytes
                st.session_state["xlsx_acumulado_bytes"] = xlsx_acumulado_bytes

    with col_out1:
        if "df_novo_tab1" not in st.session_state:
            st.info("ℹ️ Faça upload do RD e clique em **▶️ Extrair e Acumular**.")
        else:
            df_n = st.session_state["df_novo_tab1"]
            df_a = st.session_state["df_acumulado_tab1"]

            n_com = (df_n["fonte"] != "").sum()
            n_sem = (df_n["fonte"] == "").sum()

            st.success(f"✅ **Etapa 1** — {len(df_n)} parágrafos extraídos")
            st.success(f"✅ **Etapa 2** — Excel acumulado gerado ({len(df_a)} pares no total)")

            m1, m2, m3 = st.columns(3)
            m1.metric("Parágrafos extraídos", len(df_n))
            m2.metric("Com fonte", int(n_com))
            m3.metric("Sem fonte", int(n_sem))

            st.markdown("#### 🔍 Preview — pares extraídos desta extração")
            st.dataframe(
                df_n,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "paragrafo": st.column_config.TextColumn("Parágrafo", width="large"),
                    "fonte":     st.column_config.TextColumn("Fonte", width="medium"),
                },
            )

            d1, d2 = st.columns(2)
            with d1:
                st.download_button(
                    "⬇️ Excel desta extração",
                    data=st.session_state["xlsx_novo_bytes"],
                    file_name="RD_fontes_nova_extracao.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_xlsx_novo",
                )
            with d2:
                st.download_button(
                    "⬇️ Excel acumulado atualizado",
                    data=st.session_state["xlsx_acumulado_bytes"],
                    file_name="RD_fontes_acumulado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_xlsx_acum",
                )


# ══════════════════════════════════════════════════════════════
# TAB 2
# ══════════════════════════════════════════════════════════════
with tab2:
    st.subheader("Etapa 3 & 4 — Aplicar fontes no 18-K")
    st.caption(
        "Faça upload do documento alvo (18-K ou RD acumulado) e do Excel acumulado de fontes. "
        "O app faz o matching fuzzy e insere as fontes em vermelho no documento."
    )

    col_in2, col_out2 = st.columns([1, 1], gap="large")

    with col_in2:
        alvo_file = st.file_uploader(
            "📄 Documento alvo (.docx) — 18-K ou RD acumulado",
            type=["docx"],
            key="alvo_upload",
        )
        acum_file_2 = st.file_uploader(
            "📊 Excel acumulado de fontes (.xlsx)",
            type=["xlsx"],
            key="acum_upload_2",
            help="Use o Excel acumulado gerado na Etapa 1 & 2.",
        )
        threshold = st.slider(
            "Threshold de similaridade (fuzzy match)",
            min_value=50,
            max_value=100,
            value=80,
            step=5,
            help="Percentual mínimo de similaridade para aceitar um match. Reduza se poucos matches forem encontrados.",
            key="threshold_tab2",
        )

        btn_aplicar = st.button("▶️ Aplicar Fontes", type="primary", use_container_width=True, key="btn_aplicar")

    if btn_aplicar:
        if alvo_file is None:
            st.error("❌ Faça upload do documento alvo.")
        elif acum_file_2 is None:
            st.error("❌ Faça upload do Excel acumulado de fontes.")
        else:
            with st.spinner("Fazendo matching e gerando documento..."):
                alvo_bytes = alvo_file.read()
                acum_bytes_2 = acum_file_2.read()

                # Carrega Excel acumulado
                df_acum2 = pd.read_excel(io.BytesIO(acum_bytes_2), header=0, engine="openpyxl")
                df_acum2.columns = ["paragrafo", "fonte"]
                df_acum2 = df_acum2.astype(str).replace("nan", "")

                # Etapa 3: extrai parágrafos do alvo e faz fuzzy match
                paragrafos_info = extrair_paragrafos_alvo(alvo_bytes)
                mapa_fontes = fuzzy_match_paragrafos(paragrafos_info, df_acum2, threshold)

                # Etapa 4: gera Word com fontes
                docx_final = gerar_docx_com_fontes(alvo_bytes, paragrafos_info, mapa_fontes)

                # Preview
                df_preview = pd.DataFrame([
                    {"Parágrafo": txt, "Fonte encontrada": fonte}
                    for txt, fonte in mapa_fontes.items()
                    if txt.strip()
                ])

                n_com2 = (df_preview["Fonte encontrada"] != "").sum()
                n_sem2 = (df_preview["Fonte encontrada"] == "").sum()

                st.session_state["docx_final"]   = docx_final
                st.session_state["df_preview_t2"] = df_preview
                st.session_state["n_com2"]        = int(n_com2)
                st.session_state["n_sem2"]        = int(n_sem2)
                st.session_state["n_total2"]      = len(df_preview)

    with col_out2:
        if "docx_final" not in st.session_state:
            st.info("ℹ️ Faça upload dos arquivos e clique em **▶️ Aplicar Fontes**.")
        else:
            st.success("✅ **Etapa 3 & 4** — Documento gerado com sucesso!")

            m1, m2, m3 = st.columns(3)
            m1.metric("Total de parágrafos", st.session_state["n_total2"])
            m2.metric("Com fonte aplicada", st.session_state["n_com2"])
            m3.metric("Sem fonte", st.session_state["n_sem2"])

            st.markdown("#### 🔍 Preview — resultado do matching")
            st.dataframe(
                st.session_state["df_preview_t2"],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Parágrafo":        st.column_config.TextColumn("Parágrafo", width="large"),
                    "Fonte encontrada": st.column_config.TextColumn("Fonte encontrada", width="medium"),
                },
            )

            st.download_button(
                "⬇️ Word com fontes aplicadas (RD_v1.docx)",
                data=st.session_state["docx_final"],
                file_name="RD_v1_com_fontes.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="dl_docx_final",
            )


# ══════════════════════════════════════════════════════════════
# TAB 3
# ══════════════════════════════════════════════════════════════
with tab3:
    st.subheader("📊 Histórico Acumulado de Fontes")
    st.caption("Visualize o conteúdo completo do Excel acumulado.")

    acum_file_3 = st.file_uploader(
        "📊 Excel acumulado (.xlsx)",
        type=["xlsx"],
        key="acum_upload_3",
    )

    if acum_file_3 is not None:
        df_hist = pd.read_excel(io.BytesIO(acum_file_3.read()), header=0, engine="openpyxl")
        df_hist.columns = ["paragrafo", "fonte"]
        df_hist = df_hist.astype(str).replace("nan", "")

        n_total_h = len(df_hist)
        n_com_h   = (df_hist["fonte"] != "").sum()
        n_sem_h   = (df_hist["fonte"] == "").sum()

        m1, m2, m3 = st.columns(3)
        m1.metric("Total de pares", n_total_h)
        m2.metric("Com fonte", int(n_com_h))
        m3.metric("Sem fonte", int(n_sem_h))

        st.markdown("---")
        st.dataframe(
            df_hist,
            use_container_width=True,
            hide_index=True,
            column_config={
                "paragrafo": st.column_config.TextColumn("Parágrafo", width="large"),
                "fonte":     st.column_config.TextColumn("Fonte", width="medium"),
            },
        )
    else:
        st.info("ℹ️ Faça upload do Excel acumulado para visualizar o histórico.")
