import io
import re
import math
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process

st.set_page_config(page_title="RD Pipeline", layout="wide")
st.title("RD Pipeline – Extração · Consolidação · Fontes · Word")

# ── regex ─────────────────────────────────────────────────────────────────────
URL_RE = re.compile(r"https?://[^\s\)\]\}\"\'>]+", re.IGNORECASE)

# ═════════════════════════════════════════════════════════════════════════════
# ── helpers gerais ────────────────────────────────────────────────────────────
# ═════════════════════════════════════════════════════════════════════════════

def extract_urls(text: str):
    return URL_RE.findall(text or "")


def is_source_line(text: str) -> bool:
    t = (text or "").strip().lower()
    if not t:
        return False
    if t.startswith("same from above"):
        return True
    if t.startswith("source:"):
        return True
    return bool(extract_urls(t))


def append_sources(data, idx: int, sources):
    if idx is None or not sources:
        return
    seen, ordered = set(), []
    for s in sources:
        s = (s or "").strip()
        if not s:
            continue
        if s.lower().startswith("source:"):
            s = s.split(":", 1)[1].strip()
        urls = extract_urls(s)
        if urls:
            for u in urls:
                if u not in seen:
                    seen.add(u); ordered.append(u)
        else:
            if s not in seen:
                seen.add(s); ordered.append(s)
    if not ordered:
        return
    existing = (data[idx][1] or "").strip()
    if existing:
        ex_lines = [ln.strip() for ln in existing.split("\n") if ln.strip()]
        for ln in ex_lines:
            if ln not in seen:
                seen.add(ln)
        data[idx][1] = existing + "\n" + "\n".join([s for s in ordered if s not in ex_lines])
    else:
        data[idx][1] = "\n".join(ordered)


def ajustar_altura(ws, cols=(1, 2), largura_coluna=100, altura_por_linha=15):
    for col in cols:
        ws.column_dimensions[get_column_letter(col)].width = largura_coluna
    for row in range(1, ws.max_row + 1):
        max_lines = 1
        for col in cols:
            cell = ws.cell(row=row, column=col)
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            if cell.value:
                texto = str(cell.value)
                lines = 0
                for part in texto.split("\n"):
                    if len(part) == 0:
                        lines += 1
                    else:
                        lines += max(1, math.ceil(len(part) / (largura_coluna * 1.2)))
                max_lines = max(max_lines, lines)
        ws.row_dimensions[row].height = max_lines * altura_por_linha


def _run_is_red(run) -> bool:
    try:
        c = run.font.color
        if c and c.rgb:
            return str(c.rgb).upper() == "FF0000"
    except Exception:
        pass
    return False


def is_red_paragraph(paragraph) -> bool:
    runs = [r for r in paragraph.runs if r.text.strip()]
    return bool(runs) and all(_run_is_red(r) for r in runs)


def clean_text_for_match(text: str) -> str:
    cleaned = URL_RE.sub("", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned.lower()


def insert_paragraph_after(paragraph, text=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


# ═════════════════════════════════════════════════════════════════════════════
# ── Etapa 1: docx → xlsx ─────────────────────────────────────────────────────
# ═════════════════════════════════════════════════════════════════════════════

def word_to_excel_bytes(docx_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
    data = []
    current_idx = None

    for paragraph in doc.paragraphs:
        raw  = paragraph.text
        text = (raw or "").strip()
        if not text:
            continue
        if paragraph.style and paragraph.style.name.startswith("Heading"):
            continue
        if paragraph._element.xpath("ancestor::w:tbl"):
            continue
        all_bold      = all(r.bold       for r in paragraph.runs if r.text.strip())
        all_italic    = all(r.italic     for r in paragraph.runs if r.text.strip())
        all_underlined= all(r.font.underline for r in paragraph.runs if r.text.strip())
        if all_bold or all_italic or all_underlined:
            continue
        if text.startswith("\u201c") and text.endswith("\u201d"):
            continue

        if is_source_line(text):
            if current_idx is None:
                data.append(["", ""])
                current_idx = len(data) - 1
            append_sources(data, current_idx, [text])
            continue

        if text.startswith("Note:") or re.match(r"^\(\d+\)", text):
            continue

        urls    = extract_urls(text)
        cleaned = URL_RE.sub("", text)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if cleaned:
            data.append([cleaned, ""])
            current_idx = len(data) - 1
            if urls:
                append_sources(data, current_idx, urls)
        else:
            if urls:
                if current_idx is None:
                    data.append(["", ""])
                    current_idx = len(data) - 1
                append_sources(data, current_idx, urls)

    wb = Workbook()
    ws = wb.active
    ws.title = "Parágrafos"
    for i, (t, s) in enumerate(data, start=1):
        ws.cell(row=i, column=1, value=t)
        ws.cell(row=i, column=2, value=s)
    ajustar_altura(ws, cols=(1, 2), largura_coluna=100, altura_por_linha=20)
    align = Alignment(wrap_text=True, vertical="top", horizontal="left")
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.alignment = align

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ═════════════════════════════════════════════════════════════════════════════
# ── Etapa 2: combinar excels ─────────────────────────────────────────────────
# ═════════════════════════════════════════════════════════════════════════════

def combinar_excels(bytes_novo: bytes, bytes_antigo: bytes) -> bytes:
    df_novo   = pd.read_excel(io.BytesIO(bytes_novo),   header=None, engine="openpyxl")
    df_antigo = pd.read_excel(io.BytesIO(bytes_antigo), header=None, engine="openpyxl")
    df_final  = pd.concat([df_antigo, df_novo], ignore_index=True)
    buf = io.BytesIO()
    df_final.to_excel(buf, index=False, header=False)
    buf.seek(0)
    return buf.read()


# ═════════════════════════════════════════════════════════════════════════════
# ── Etapa 3: aplicar fontes (fuzzy match) ─────────────────────────────────────
# ═════════════════════════════════════════════════════════════════════════════

def apply_sources_bytes(docx_bytes: bytes, excel_bytes: bytes, threshold: int = 80):
    """Aplica fontes no Word usando fuzzy matching (rapidfuzz partial_ratio)."""
    wb = load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active

    pares = []
    for row in ws.iter_rows(min_row=1, max_col=2, values_only=True):
        a, b = row[0], row[1]
        if not a or not b:
            continue
        chave = clean_text_for_match(str(a).strip())
        fonte = str(b).strip()
        if chave and fonte:
            pares.append((chave, fonte))

    doc = Document(io.BytesIO(docx_bytes))

    if not pares:
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.read(), pd.DataFrame(columns=["Parágrafo", "Fonte aplicada", "Score"])

    chaves     = [p[0] for p in pares]
    fontes_map = {p[0]: p[1] for p in pares}
    preview    = []

    for p in doc.paragraphs:
        txt = p.text.strip()
        if not txt or is_red_paragraph(p):
            continue
        texto_par = clean_text_for_match(txt)
        resultado = process.extractOne(texto_par, chaves, scorer=fuzz.partial_ratio)
        score     = resultado[1] if resultado else 0

        if resultado and score >= threshold:
            fonte_val = fontes_map[resultado[0]]
            linhas    = [ln.strip() for ln in fonte_val.split("\n") if ln.strip()]
            last_p    = p
            for linha in linhas:
                new_p  = insert_paragraph_after(last_p)
                run    = new_p.add_run(linha)
                run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
                run.font.size      = Pt(10)
                new_p.paragraph_format.space_before = Pt(12)
                last_p = new_p
            preview.append({"Parágrafo": txt, "Fonte aplicada": fonte_val, "Score": score})
        else:
            preview.append({"Parágrafo": txt, "Fonte aplicada": "", "Score": score})

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read(), pd.DataFrame(preview)


# ═════════════════════════════════════════════════════════════════════════════
# ── Etapa 4: excel → docx ────────────────────────────────────────────────────
# ═════════════════════════════════════════════════════════════════════════════

def excel_para_docx_bytes(excel_bytes: bytes) -> bytes:
    df  = pd.read_excel(io.BytesIO(excel_bytes), header=None, engine="openpyxl")
    doc = Document()

    title = doc.add_paragraph("", style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("Recent Developments – With Sources")

    for i in range(df.shape[0]):
        texto = str(df.iat[i, 0]).strip()
        fonte = str(df.iat[i, 1]).strip() if df.shape[1] > 1 else ""
        if not texto or texto.lower() in ("nan", ""):
            continue
        p_txt = doc.add_paragraph(texto)
        p_txt.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_txt.paragraph_format.space_after = Pt(3)
        if fonte and fonte.lower() not in ("nan", ""):
            p_src = doc.add_paragraph()
            r     = p_src.add_run(fonte)
            r.font.size       = Pt(10)
            r.font.color.rgb  = RGBColor(255, 0, 0)
            p_src.paragraph_format.space_after = Pt(12)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ═════════════════════════════════════════════════════════════════════════════
# ── UI ───────────────────────────────────────────────────────────────────────
# ═════════════════════════════════════════════════════════════════════════════

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1 · .docx → .xlsx",
    "2 · Combinar excels",
    "3 · Aplicar fontes",
    "4 · .xlsx → .docx",
    "5 · Visualizar excel",
])

# ── Tab 1 ──────────────────────────────────────────────────────────────────
with tab1:
    st.header("Etapa 1 – Extrair parágrafos do Word para Excel")
    st.write("Carrega um ficheiro **.docx** e gera um **.xlsx** com os parágrafos (col A) e as fontes/URLs detectadas (col B).")
    f1 = st.file_uploader("Ficheiro Word (.docx)", type=["docx"], key="t1_docx")
    if f1:
        with st.spinner("Extraindo parágrafos…"):
            out_bytes = word_to_excel_bytes(f1.read())
        st.success("Extração concluída!")
        st.download_button(
            "⬇️ Baixar Excel",
            data=out_bytes,
            file_name="RD_sourcebook_v0.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ── Tab 2 ──────────────────────────────────────────────────────────────────
with tab2:
    st.header("Etapa 2 – Combinar ficheiros Excel")
    st.write("Empilha o Excel **novo** abaixo do Excel **antigo** (acumulado), gerando um único ficheiro.")
    col_n, col_a = st.columns(2)
    with col_n:
        f2_novo   = st.file_uploader("Excel NOVO (.xlsx)", type=["xlsx"], key="t2_novo")
    with col_a:
        f2_antigo = st.file_uploader("Excel ANTIGO / acumulado (.xlsx)", type=["xlsx"], key="t2_antigo")
    if st.button("▶️ Combinar", type="primary", use_container_width=True, key="t2_btn"):
        if not f2_novo or not f2_antigo:
            st.error("❌ Carregue os dois ficheiros.")
        else:
            with st.spinner("Combinando…"):
                combined = combinar_excels(f2_novo.read(), f2_antigo.read())
            st.success("Combinação concluída!")
            st.download_button(
                "⬇️ Baixar Excel combinado",
                data=combined,
                file_name="RD_sourcebook_matched_sources.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="t2_dl",
            )

# ── Tab 3 ──────────────────────────────────────────────────────────────────
with tab3:
    st.header("Etapa 3 – Aplicar fontes no documento Word")
    st.write(
        "Carrega o Word alvo e o Excel acumulado com as fontes. "
        "O app usa **fuzzy matching** (rapidfuzz) para encontrar o parágrafo mais similar e "
        "insere a fonte em vermelho logo após."
    )

    col_w, col_x = st.columns(2)
    with col_w:
        f_word  = st.file_uploader("Ficheiro Word (.docx)", type=["docx"], key="t3_docx")
    with col_x:
        f_excel = st.file_uploader("Excel com fontes (.xlsx)", type=["xlsx"], key="t3_xlsx")

    threshold_val = st.slider(
        "Threshold de similaridade (fuzzy match)",
        min_value=50, max_value=100, value=80, step=5,
        key="t3_threshold",
        help="Score mínimo para aceitar um match. 100 = match quase exato.",
    )
    st.info(
        "ℹ️ **Threshold**: valores mais baixos → mais matches (menos precisos). "
        "Valores mais altos → menos matches (mais precisos). "
        "Score calculado com `partial_ratio` do rapidfuzz."
    )

    btn_t3 = st.button("▶️ Aplicar fontes", type="primary", use_container_width=True, key="t3_btn")

    if btn_t3:
        if not f_word or not f_excel:
            st.error("❌ Carregue o Word e o Excel antes de continuar.")
        else:
            with st.spinner("Fazendo matching e inserindo fontes…"):
                result_bytes, df_prev = apply_sources_bytes(
                    f_word.read(), f_excel.read(), threshold=threshold_val
                )
            st.session_state["t3_result"]  = result_bytes
            st.session_state["t3_preview"] = df_prev

    if "t3_result" in st.session_state:
        df_prev = st.session_state["t3_preview"]
        n_com   = int((df_prev["Fonte aplicada"] != "").sum())
        n_sem   = int((df_prev["Fonte aplicada"] == "").sum())
        m1, m2, m3 = st.columns(3)
        m1.metric("Parágrafos verificados", len(df_prev))
        m2.metric("Com fonte aplicada ✅", n_com)
        m3.metric("Sem match ❌", n_sem)

        st.markdown("#### 🔍 Resultado do matching")
        st.dataframe(
            df_prev,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Parágrafo":      st.column_config.TextColumn("Parágrafo",      width="large"),
                "Fonte aplicada": st.column_config.TextColumn("Fonte aplicada", width="medium"),
                "Score":          st.column_config.NumberColumn("Score", format="%d", width="small"),
            },
        )

        st.download_button(
            "⬇️ Baixar Word com fontes",
            data=st.session_state["t3_result"],
            file_name="RD_v1.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            key="t3_dl",
        )

# ── Tab 4 ──────────────────────────────────────────────────────────────────
with tab4:
    st.header("Etapa 4 – Gerar Word com fontes a partir do Excel")
    st.write("Carrega o Excel acumulado (col A = texto, col B = fonte) e gera um **.docx** com texto e fontes em vermelho.")
    f4 = st.file_uploader("Excel (.xlsx)", type=["xlsx"], key="t4_xlsx")
    if st.button("▶️ Gerar Word", type="primary", use_container_width=True, key="t4_btn"):
        if not f4:
            st.error("❌ Carregue o Excel.")
        else:
            with st.spinner("Gerando Word…"):
                docx_out = excel_para_docx_bytes(f4.read())
            st.success("Word gerado!")
            st.download_button(
                "⬇️ Baixar Word",
                data=docx_out,
                file_name="RD_with_sources.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="t4_dl",
            )

# ── Tab 5 ──────────────────────────────────────────────────────────────────
with tab5:
    st.header("Visualizar Excel")
    f5 = st.file_uploader("Excel (.xlsx)", type=["xlsx"], key="t5_xlsx")
    if f5:
        df5 = pd.read_excel(io.BytesIO(f5.read()), header=None, engine="openpyxl")
        df5.columns = [f"Col {i+1}" for i in range(df5.shape[1])]
        st.dataframe(df5, use_container_width=True, hide_index=True)
        st.caption(f"{len(df5)} linhas · {df5.shape[1]} colunas")
