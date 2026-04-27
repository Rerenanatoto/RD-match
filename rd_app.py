
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

st.set_page_config(page_title="RD – Gestão de Fontes", layout="wide")

# ============================================================
# Constantes
# ============================================================

URL_RE = re.compile(r"https?://[^\s\)\]\}\"\'>]+", re.IGNORECASE)

# ============================================================
# Utilidades comuns
# ============================================================

def extract_urls(text: str):
    """Extrai todas as URLs de um texto."""
    return URL_RE.findall(text or "")


def is_source_line(text: str) -> bool:
    """Detecta se um parágrafo é linha de fonte.

    Considera como fonte:
    - linhas que comecem por 'same from above'
    - linhas que comecem por 'source:'
    - linhas que contenham qualquer URL
    """
    t = (text or "").strip().lower()
    if not t:
        return False
    if t.startswith("same from above"):
        return True
    if t.startswith("source:"):
        return True
    return bool(extract_urls(t))


def append_sources(data, idx, sources):
    """Acrescenta fontes na coluna B (índice 1) de data[idx], sem duplicatas."""
    if idx is None or not sources:
        return
    seen = set()
    ordered = []
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
                    seen.add(u)
                    ordered.append(u)
        else:
            if s not in seen:
                seen.add(s)
                ordered.append(s)
    if not ordered:
        return
    existing = (data[idx][1] or "").strip()
    if existing:
        existing_lines = [ln.strip() for ln in existing.split("\n") if ln.strip()]
        for ln in existing_lines:
            seen.add(ln)
        new_lines = [s for s in ordered if s not in existing_lines]
        if new_lines:
            data[idx][1] = existing + "\n" + "\n".join(new_lines)
    else:
        data[idx][1] = "\n".join(ordered)


def ajustar_altura(ws, cols=(1, 2), largura_coluna=100, altura_por_linha=15):
    """Ajusta largura e altura de células com wrap."""
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


# ============================================================
# Etapa 1 – Extrair parágrafos + fontes do Word → Excel
# (fiel a rd_docx_to_excel_paragraphs.py – versão nova)
# ============================================================

def word_to_excel_bytes(docx_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
    data = []          # lista de [texto, fontes_str]
    current_idx = None

    for paragraph in doc.paragraphs:
        raw = (paragraph.text or "").strip()
        if not raw:
            continue

        # Ignora headings e parágrafos dentro de tabelas
        if paragraph.style and paragraph.style.name.startswith("Heading"):
            continue
        if paragraph._element.xpath("ancestor::w:tbl"):
            continue

        # Filtros de formatação
        runs_with_text = [r for r in paragraph.runs if r.text.strip()]
        if runs_with_text:
            all_bold       = all(r.bold for r in runs_with_text)
            all_italic     = all(r.italic for r in runs_with_text)
            all_underlined = all(r.font.underline for r in runs_with_text)
            if all_bold or all_italic or all_underlined:
                continue

        # Ignora citações entre aspas curvas
        if raw.startswith("\u201c") and raw.endswith("\u201d"):
            continue

        # Se é linha de fonte → acumula no parágrafo anterior
        if is_source_line(raw):
            if current_idx is None:
                data.append(["", ""])
                current_idx = len(data) - 1
            append_sources(data, current_idx, [raw])
            continue

        # Ignora Note: e parágrafos numéricos tipo (1)
        if raw.startswith("Note:") or re.match(r"^\(\d+\)", raw):
            continue

        # Extrai URLs inline e limpa o texto
        urls = extract_urls(raw)
        cleaned = URL_RE.sub("", raw)
        cleaned = re.sub(r"\s+", " ", cleaned).strip()

        if cleaned:
            data.append([cleaned, ""])
            current_idx = len(data) - 1
            if urls:
                append_sources(data, current_idx, urls)
        else:
            # Parágrafo inteiro era só URL → trata como fonte do anterior
            if urls:
                if current_idx is None:
                    data.append(["", ""])
                    current_idx = len(data) - 1
                append_sources(data, current_idx, urls)

    # Cria Excel
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


# ============================================================
# Etapa 2 – Combinar dois Excels (antigo + novo)
# (fiel a rd_match_excel_file.py – versão nova)
# ============================================================

def combinar_excels(bytes_antigo: bytes, bytes_novo: bytes) -> bytes:
    df_antigo = pd.read_excel(io.BytesIO(bytes_antigo), header=None, engine="openpyxl")
    df_novo   = pd.read_excel(io.BytesIO(bytes_novo),   header=None, engine="openpyxl")
    # Antigo primeiro, novo abaixo
    df_final = pd.concat([df_antigo, df_novo], ignore_index=True)
    buf = io.BytesIO()
    df_final.to_excel(buf, index=False, header=False, engine="openpyxl")
    buf.seek(0)
    return buf.read()


# ============================================================
# Etapa 3 – Aplicar fontes no Word
# (fiel a rd_apply_sources_to_file.py)
# ============================================================

def _run_is_red(run) -> bool:
    try:
        if run.font.color.rgb == RGBColor(0xFF, 0x00, 0x00):
            return True
    except Exception:
        pass
    try:
        if "FF0000" in run._r.xml.upper():
            return True
    except Exception:
        pass
    return False


def is_red_paragraph(paragraph) -> bool:
    runs_with_text = [r for r in paragraph.runs if r.text.strip()]
    if not runs_with_text:
        return False
    if any(_run_is_red(r) for r in runs_with_text):
        return True
    try:
        if "FF0000" in paragraph._p.xml.upper():
            return True
    except Exception:
        pass
    return False


def clean_text_for_match(text: str) -> str:
    cleaned = re.sub(r"https?://\S+", "", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    return cleaned


def insert_paragraph_after(paragraph, text=None):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        new_para.add_run(text)
    return new_para


def apply_sources_bytes(docx_bytes: bytes, excel_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(excel_bytes))
    ws = wb.active
    doc = Document(io.BytesIO(docx_bytes))

    for row in ws.iter_rows(min_row=1, max_col=2, values_only=True):
        val_a, val_b = row
        if not val_a or not val_b:
            continue
        chave = str(val_a).strip()
        fonte = str(val_b)

        for p in doc.paragraphs:
            if not p.text.strip():
                continue
            # Pula parágrafos já vermelhos (fontes já inseridas)
            if is_red_paragraph(p):
                continue
            texto_par   = clean_text_for_match(p.text)
            chave_limpa = clean_text_for_match(chave)
            if texto_par.lower() == chave_limpa.lower():
                new_p = insert_paragraph_after(p)
                run = new_p.add_run(fonte)
                run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
                run.font.size = Pt(10)
                new_p.paragraph_format.space_before = Pt(12)
                break

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ============================================================
# Etapa 4 – Criar Word com fontes a partir do Excel acumulado
# (fiel a rd_create_docx_with_sourcebook.py)
# ============================================================

def excel_para_docx_bytes(excel_bytes: bytes) -> bytes:
    df = pd.read_excel(io.BytesIO(excel_bytes), header=None, engine="openpyxl")

    doc = Document()

    # Título
    title = doc.add_paragraph("", style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run("Recent Developments – With Sources")

    for i in range(df.shape[0]):
        texto = str(df.iat[i, 0]).strip() if pd.notna(df.iat[i, 0]) else ""
        fonte = str(df.iat[i, 1]).strip() if df.shape[1] > 1 and pd.notna(df.iat[i, 1]) else ""

        if not texto:
            continue

        # Parágrafo principal
        p_texto = doc.add_paragraph(texto)
        p_texto.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_texto.paragraph_format.space_after = Pt(3)

        # Parágrafo de fonte (vermelho)
        p_fonte = doc.add_paragraph()
        r = p_fonte.add_run(fonte if fonte else "")
        r.font.size = Pt(10)
        r.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        p_fonte.paragraph_format.space_after = Pt(12)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


# ============================================================
# APP PRINCIPAL
# ============================================================

st.title("RD – Gestão de Fontes")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1 · Extrair parágrafos",
    "2 · Combinar Excels",
    "3 · Aplicar fontes no Word",
    "4 · Gerar Word com fontes",
    "5 · Visualizar Excel",
])

# ------------------------------------------------------------------
# TAB 1 – Extrair parágrafos do Word → Excel
# ------------------------------------------------------------------
with tab1:
    st.header("Etapa 1 – Extrair parágrafos e fontes do Word")
    st.write(
        "Carrega o arquivo Word do RD. O app extrai cada parágrafo (col A) "
        "junto com a URL/fonte associada (col B) no Excel."
    )
    file_docx = st.file_uploader("Ficheiro Word (.docx)", type=["docx"], key="t1_docx")
    if file_docx:
        with st.spinner("Extraindo parágrafos…"):
            xlsx_bytes = word_to_excel_bytes(file_docx.read())
        st.success("Extração concluída!")
        st.download_button(
            "⬇️ Baixar Excel extraído",
            data=xlsx_bytes,
            file_name="RD_sourcebook_v0.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ------------------------------------------------------------------
# TAB 2 – Combinar dois Excels
# ------------------------------------------------------------------
with tab2:
    st.header("Etapa 2 – Combinar Excel acumulado + Excel novo")
    st.write(
        "Carrega o Excel antigo (acumulado) e o Excel novo (recém extraído). "
        "O resultado terá o **antigo primeiro**, o novo abaixo."
    )
    col_a, col_b = st.columns(2)
    with col_a:
        f_antigo = st.file_uploader("Excel acumulado (antigo)", type=["xlsx"], key="t2_old")
    with col_b:
        f_novo = st.file_uploader("Excel novo", type=["xlsx"], key="t2_new")

    if f_antigo and f_novo:
        with st.spinner("Combinando…"):
            combined = combinar_excels(f_antigo.read(), f_novo.read())
        st.success("Combinação concluída!")
        st.download_button(
            "⬇️ Baixar Excel combinado",
            data=combined,
            file_name="RD_sourcebook_combinado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ------------------------------------------------------------------
# TAB 3 – Aplicar fontes no Word
# ------------------------------------------------------------------
with tab3:
    st.header("Etapa 3 – Aplicar fontes no documento Word")
    st.write(
        "Carrega o Word do RD (sem fontes) e o Excel acumulado com as fontes. "
        "O app insere cada fonte como parágrafo vermelho logo após o texto correspondente."
    )
    col_w, col_x = st.columns(2)
    with col_w:
        f_word = st.file_uploader("Ficheiro Word (.docx)", type=["docx"], key="t3_docx")
    with col_x:
        f_excel = st.file_uploader("Excel com fontes (.xlsx)", type=["xlsx"], key="t3_xlsx")

    if f_word and f_excel:
        with st.spinner("Aplicando fontes…"):
            result_bytes = apply_sources_bytes(f_word.read(), f_excel.read())
        st.success("Fontes aplicadas!")
        st.download_button(
            "⬇️ Baixar Word com fontes",
            data=result_bytes,
            file_name="RD_v1.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# ------------------------------------------------------------------
# TAB 4 – Gerar Word a partir do Excel acumulado
# ------------------------------------------------------------------
with tab4:
    st.header("Etapa 4 – Gerar Word completo a partir do Excel")
    st.write(
        "Carrega o Excel acumulado (col A = texto, col B = fontes). "
        "O app gera um Word com cada parágrafo seguido da respectiva fonte em vermelho."
    )
    f_xl4 = st.file_uploader("Excel acumulado (.xlsx)", type=["xlsx"], key="t4_xlsx")
    if f_xl4:
        with st.spinner("Gerando Word…"):
            docx_bytes = excel_para_docx_bytes(f_xl4.read())
        st.success("Documento gerado!")
        st.download_button(
            "⬇️ Baixar Word gerado",
            data=docx_bytes,
            file_name="RD_com_fontes.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# ------------------------------------------------------------------
# TAB 5 – Visualizar Excel
# ------------------------------------------------------------------
with tab5:
    st.header("Visualizar Excel")
    f_view = st.file_uploader("Excel (.xlsx)", type=["xlsx"], key="t5_xlsx")
    if f_view:
        df_view = pd.read_excel(io.BytesIO(f_view.read()), header=None, engine="openpyxl")
        df_view.columns = ["Parágrafo", "Fontes"] if df_view.shape[1] >= 2 else df_view.columns
        st.dataframe(df_view, use_container_width=True)
        st.caption(f"{len(df_view)} linhas")
