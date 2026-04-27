import io
import math
import re
from copy import deepcopy

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
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
st.caption("Pipeline completo: extração de fontes do RD → Excel acumulado → aplicação no 18-K")
st.markdown("---")


# ──────────────────────────────────────────────────────────────
# HELPERS DE DETECÇÃO DE COR VERMELHA
# ──────────────────────────────────────────────────────────────

def _run_is_red(run) -> bool:
    """
    Verifica se um run é vermelho usando 3 métodos:
    1) run.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)
    2) busca direta no XML do run por w:val="FF0000"
    3) regex no XML do run
    """
    # Método 1: API python-docx
    try:
        if run.font.color.rgb == RGBColor(0xFF, 0x00, 0x00):
            return True
    except Exception:
        pass

    # Método 2 e 3: XML direto
    try:
        xml = run._r.xml
        if re.search(r'w:val=["\'](?:FF0000|ff0000)["\']', xml):
            return True
        if "FF0000" in xml.upper():
            return True
    except Exception:
        pass

    return False


def is_red_paragraph(paragraph) -> bool:
    """
    Retorna True se o parágrafo contém texto vermelho.
    Usa ANY (qualquer run vermelho) para capturar parágrafos
    mistos e URLs em vermelho.
    Fallback: verifica XML do parágrafo inteiro.
    """
    runs_com_texto = [r for r in paragraph.runs if r.text.strip()]

    if not runs_com_texto:
        # Fallback: verifica XML do parágrafo
        try:
            xml = paragraph._p.xml
            return bool(re.search(r'w:val=["\'](?:FF0000|ff0000)["\']', xml))
        except Exception:
            return False

    return any(_run_is_red(r) for r in runs_com_texto)


def is_paragrafo_valido_texto(txt: str) -> bool:
    """Filtros de exclusão para parágrafos do Word."""
    if not txt:
        return False
    if txt.startswith("Note:") or txt.startswith("Source:"):
        return False
    if re.match(r"^\(\d+\)", txt):
        return False
    return True


# ──────────────────────────────────────────────────────────────
# HELPERS DE FORMATAÇÃO EXCEL
# ──────────────────────────────────────────────────────────────

def ajustar_altura(ws, col=1, largura_coluna=100, altura_por_linha=20):
    """Ajusta altura das linhas proporcional ao tamanho do texto."""
    col_letter = get_column_letter(col)
    ws.column_dimensions[col_letter].width = largura_coluna
    for row_cells in ws.iter_rows(
        min_row=1, max_col=col, max_row=ws.max_row, min_col=col
    ):
        for cell in row_cells:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cell.value:
                texto = str(cell.value)
                linhas = sum(
                    math.ceil(len(p) / (largura_coluna * 1.2)) or 1
                    for p in texto.split("\n")
                )
                ws.row_dimensions[cell.row].height = max(
                    linhas * altura_por_linha, altura_por_linha
                )


# ──────────────────────────────────────────────────────────────
# ETAPA 1 — Extrair pares (parágrafo → fontes vermelhas)
# ──────────────────────────────────────────────────────────────

def extrair_pares_docx(docx_bytes: bytes) -> pd.DataFrame:
    """
    Lê o .docx e extrai pares parágrafo → fontes vermelhas.
    Se houver múltiplas fontes vermelhas consecutivas sob o mesmo
    parágrafo, todas são reunidas na mesma célula separadas por '\\n'.
    """
    doc = Document(io.BytesIO(docx_bytes))
    paragrafos_doc = doc.paragraphs
    n = len(paragrafos_doc)
    pares = []
    i = 0

    while i < n:
        p = paragrafos_doc[i]

        # Pula parágrafos vermelhos isolados
        if is_red_paragraph(p):
            i += 1
            continue

        txt = p.text.strip()

        # Filtros de exclusão
        if not txt:
            i += 1
            continue
        if txt.startswith("Note:") or txt.startswith("Source:"):
            i += 1
            continue
        if re.match(r"^\(\d+\)", txt):
            i += 1
            continue
        if txt.startswith("\u201c") and txt.endswith("\u201d"):
            i += 1
            continue

        # Parágrafo válido — avança i
        i += 1

        # Coleta TODOS os parágrafos vermelhos consecutivos logo abaixo
        fontes_coletadas = []
        while i < n and is_red_paragraph(paragrafos_doc[i]):
            fonte_txt = paragrafos_doc[i].text.strip()
            if fonte_txt:
                fontes_coletadas.append(fonte_txt)
            i += 1

        fonte_final = "\n".join(fontes_coletadas)
        pares.append({"paragrafo": txt, "fonte": fonte_final})

    return pd.DataFrame(pares, columns=["paragrafo", "fonte"])


# ──────────────────────────────────────────────────────────────
# ETAPA 2 — Exportar e acumular Excel
# ──────────────────────────────────────────────────────────────

def exportar_excel(df: pd.DataFrame) -> bytes:
    """Exporta DataFrame [paragrafo, fonte] para bytes de .xlsx formatado."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Fontes"

    for ci, header in enumerate(["Parágrafo", "Fonte"], start=1):
        c = ws.cell(row=1, column=ci, value=header)
        c.alignment = Alignment(wrap_text=True, vertical="top")

    for i, row in enumerate(df.itertuples(index=False), start=2):
        ws.cell(row=i, column=1, value=row.paragrafo).alignment = Alignment(
            wrap_text=True, vertical="top"
        )
        ws.cell(row=i, column=2, value=row.fonte).alignment = Alignment(
            wrap_text=True, vertical="top"
        )

    ws.column_dimensions[get_column_letter(1)].width = 100
    ws.column_dimensions[get_column_letter(2)].width = 80
    ajustar_altura(ws, col=1, largura_coluna=100, altura_por_linha=20)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def acumular_excel(
    df_novo: pd.DataFrame, excel_bytes_acumulado: bytes | None
) -> tuple:
    """
    Combina novo DataFrame com Excel acumulado.
    Novo em cima → drop_duplicates mantém mais recente.
    """
    if excel_bytes_acumulado is not None:
        df_acumulado = pd.read_excel(
            io.BytesIO(excel_bytes_acumulado), header=0, engine="openpyxl"
        )
        df_acumulado.columns = ["paragrafo", "fonte"]
        df_acumulado = df_acumulado.astype(str).replace("nan", "")
        df_merged = pd.concat([df_novo, df_acumulado], ignore_index=True)
        df_merged = df_merged.drop_duplicates(subset=["paragrafo"], keep="first")
        df_merged = df_merged.reset_index(drop=True)
    else:
        df_merged = df_novo.copy()

    return exportar_excel(df_merged), df_merged


# ──────────────────────────────────────────────────────────────
# ETAPA 3 — Fuzzy match Word alvo × Excel acumulado
# ──────────────────────────────────────────────────────────────

def extrair_paragrafos_alvo(docx_bytes: bytes) -> list:
    """
    Extrai todos os parágrafos do Word alvo preservando estrutura.
    Retorna lista de dicts {texto, style_name, is_red}.
    """
    doc = Document(io.BytesIO(docx_bytes))
    return [
        {
            "texto": p.text,
            "style_name": p.style.name,
            "is_red": is_red_paragraph(p),
        }
        for p in doc.paragraphs
    ]


def fuzzy_match_paragrafos(
    paragrafos_info: list, df_acumulado: pd.DataFrame, threshold: int = 80
) -> dict:
    """
    Fuzzy match entre parágrafos do Word alvo e col 'paragrafo' do Excel acumulado.
    Retorna dict {texto: fonte_matched}.
    """
    if df_acumulado.empty:
        return {p["texto"]: "" for p in paragrafos_info}

    trechos = df_acumulado["paragrafo"].tolist()
    fontes = df_acumulado["fonte"].tolist()
    mapa = {}

    for p in paragrafos_info:
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


# ──────────────────────────────────────────────────────────────
# ETAPA 4 — Gerar Word com fontes aplicadas
# ──────────────────────────────────────────────────────────────

def gerar_docx_com_fontes(
    docx_bytes_alvo: bytes,
    paragrafos_info: list,
    mapa_fontes: dict,
) -> bytes:
    """
    Reconstrói o Word alvo inserindo parágrafos vermelhos
    com as fontes logo após cada parágrafo com match.
    Se a fonte tiver múltiplas linhas (separadas por \\n),
    cada linha vira um parágrafo vermelho separado.
    Preserva estrutura e estilos do documento original.
    """
    doc_original = Document(io.BytesIO(docx_bytes_alvo))
    doc_novo = Document(io.BytesIO(docx_bytes_alvo))

    body = doc_novo.element.body
    sect_pr = body.find(qn("w:sectPr"))

    # Remove todos os filhos exceto sectPr
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)

    def _inserir(elem):
        if sect_pr is not None:
            body.insert(list(body).index(sect_pr), elem)
        else:
            body.append(elem)

    def _criar_paragrafo_vermelho(texto_linha: str):
        """Cria um w:p com texto em vermelho Pt(10)."""
        p_el = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")
        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), "0")
        spacing.set(qn("w:after"), "170")
        pPr.append(spacing)
        p_el.append(pPr)

        r_el = OxmlElement("w:r")
        rPr = OxmlElement("w:rPr")
        color = OxmlElement("w:color")
        color.set(qn("w:val"), "FF0000")
        sz = OxmlElement("w:sz")
        sz.set(qn("w:val"), "20")  # 20 half-points = Pt(10)
        rPr.append(color)
        rPr.append(sz)
        r_el.append(rPr)

        t_el = OxmlElement("w:t")
        t_el.text = texto_linha
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r_el.append(t_el)
        p_el.append(r_el)
        return p_el

    for p_orig in doc_original.paragraphs:
        _inserir(deepcopy(p_orig._element))

        txt = p_orig.text.strip()
        fonte = mapa_fontes.get(txt, "")

        if fonte and not is_red_paragraph(p_orig) and txt:
            # Cada linha da fonte vira um parágrafo vermelho separado
            for linha in fonte.split("\n"):
                linha = linha.strip()
                if linha:
                    _inserir(_criar_paragrafo_vermelho(linha))

    buf = io.BytesIO()
    doc_novo.save(buf)
    buf.seek(0)
    return buf.read()


# ──────────────────────────────────────────────────────────────
# UI — 3 ABAS
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
    st.subheader("Etapa 1 & 2 — Extrair parágrafos e fontes do RD")
    st.caption(
        "Faça upload do RD (.docx). "
        "O app detecta automaticamente os parágrafos em vermelho como fontes. "
        "Se houver múltiplas fontes sob o mesmo parágrafo, todas são reunidas na mesma célula."
    )

    col_in1, col_out1 = st.columns([1, 1], gap="large")

    with col_in1:
        rd_file = st.file_uploader(
            "📄 Documento RD (.docx)",
            type=["docx"],
            key="rd_upload",
            help="Parágrafos normais seguidos de fontes em vermelho.",
        )
        acum_file_1 = st.file_uploader(
            "📊 Excel acumulado — opcional (combinar com fontes anteriores)",
            type=["xlsx"],
            key="acum_upload_1",
        )
        btn_extrair = st.button(
            "▶️ Extrair e Acumular",
            type="primary",
            use_container_width=True,
            key="btn_extrair",
        )

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

                st.session_state["df_novo_tab1"]          = df_novo
                st.session_state["df_acumulado_tab1"]     = df_acumulado
                st.session_state["xlsx_novo_bytes"]       = xlsx_novo_bytes
                st.session_state["xlsx_acumulado_bytes"]  = xlsx_acumulado_bytes

    with col_out1:
        if "df_novo_tab1" not in st.session_state:
            st.info("ℹ️ Faça upload do RD e clique em **▶️ Extrair e Acumular**.")
        else:
            df_n = st.session_state["df_novo_tab1"]
            df_a = st.session_state["df_acumulado_tab1"]
            n_com = int((df_n["fonte"] != "").sum())
            n_sem = int((df_n["fonte"] == "").sum())

            st.success(f"✅ Etapa 1 — {len(df_n)} parágrafos extraídos")
            st.success(f"✅ Etapa 2 — Excel acumulado gerado ({len(df_a)} pares no total)")

            m1, m2, m3 = st.columns(3)
            m1.metric("Parágrafos extraídos", len(df_n))
            m2.metric("Com fonte", n_com)
            m3.metric("Sem fonte", n_sem)

            st.markdown("#### 🔍 Preview — pares extraídos")
            st.dataframe(
                df_n,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "paragrafo": st.column_config.TextColumn("Parágrafo", width="large"),
                    "fonte":     st.column_config.TextColumn("Fonte(s)", width="medium"),
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
        "Faça upload do documento alvo (18-K ou RD acumulado) e do Excel acumulado. "
        "O app faz o matching fuzzy e insere as fontes em vermelho — "
        "uma por linha, caso haja múltiplas fontes para o mesmo parágrafo."
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
            help="Reduza se poucos matches forem encontrados.",
            key="threshold_tab2",
        )
        btn_aplicar = st.button(
            "▶️ Aplicar Fontes",
            type="primary",
            use_container_width=True,
            key="btn_aplicar",
        )

    if btn_aplicar:
        if alvo_file is None:
            st.error("❌ Faça upload do documento alvo.")
        elif acum_file_2 is None:
            st.error("❌ Faça upload do Excel acumulado de fontes.")
        else:
            with st.spinner("Fazendo matching e gerando documento..."):
                alvo_bytes  = alvo_file.read()
                acum_bytes2 = acum_file_2.read()

                df_acum2 = pd.read_excel(
                    io.BytesIO(acum_bytes2), header=0, engine="openpyxl"
                )
                df_acum2.columns = ["paragrafo", "fonte"]
                df_acum2 = df_acum2.astype(str).replace("nan", "")

                paragrafos_info = extrair_paragrafos_alvo(alvo_bytes)
                mapa_fontes     = fuzzy_match_paragrafos(paragrafos_info, df_acum2, threshold)
                docx_final      = gerar_docx_com_fontes(alvo_bytes, paragrafos_info, mapa_fontes)

                df_preview = pd.DataFrame([
                    {"Parágrafo": txt, "Fonte(s) encontrada(s)": fonte}
                    for txt, fonte in mapa_fontes.items()
                    if txt.strip()
                ])

                n_com2   = int((df_preview["Fonte(s) encontrada(s)"] != "").sum())
                n_sem2   = int((df_preview["Fonte(s) encontrada(s)"] == "").sum())
                n_total2 = len(df_preview)

                st.session_state["docx_final"]    = docx_final
                st.session_state["df_preview_t2"] = df_preview
                st.session_state["n_com2"]        = n_com2
                st.session_state["n_sem2"]        = n_sem2
                st.session_state["n_total2"]      = n_total2

    with col_out2:
        if "docx_final" not in st.session_state:
            st.info("ℹ️ Faça upload dos arquivos e clique em **▶️ Aplicar Fontes**.")
        else:
            st.success("✅ Documento gerado com sucesso!")

            m1, m2, m3 = st.columns(3)
            m1.metric("Total parágrafos", st.session_state["n_total2"])
            m2.metric("Com fonte aplicada", st.session_state["n_com2"])
            m3.metric("Sem fonte", st.session_state["n_sem2"])

            st.markdown("#### 🔍 Preview — resultado do matching")
            st.dataframe(
                st.session_state["df_preview_t2"],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Parágrafo":               st.column_config.TextColumn("Parágrafo", width="large"),
                    "Fonte(s) encontrada(s)":  st.column_config.TextColumn("Fonte(s)", width="medium"),
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
        df_hist = pd.read_excel(
            io.BytesIO(acum_file_3.read()), header=0, engine="openpyxl"
        )
        df_hist.columns = ["paragrafo", "fonte"]
        df_hist = df_hist.astype(str).replace("nan", "")

        n_total_h = len(df_hist)
        n_com_h   = int((df_hist["fonte"] != "").sum())
        n_sem_h   = int((df_hist["fonte"] == "").sum())

        m1, m2, m3 = st.columns(3)
        m1.metric("Total de pares", n_total_h)
        m2.metric("Com fonte", n_com_h)
        m3.metric("Sem fonte", n_sem_h)

        st.markdown("---")
        st.dataframe(
            df_hist,
            use_container_width=True,
            hide_index=True,
            column_config={
                "paragrafo": st.column_config.TextColumn("Parágrafo", width="large"),
                "fonte":     st.column_config.TextColumn("Fonte(s)", width="medium"),
            },
        )
    else:
        st.info("ℹ️ Faça upload do Excel acumulado para visualizar o histórico.")
