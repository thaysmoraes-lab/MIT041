# -*- coding: utf-8 -*-
"""
MIT041 - Automação do Levantamento Fiscal (TOTVS Protheus)
Cruza a TES Inteligente (F4M) com o cadastro de TES (SF4), incorpora
cBenef (F3K) e códigos de lançamento (CC7), e gera o documento MIT041.

Autor: Prime Flow  |  Frontend: Streamlit
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="MIT041 - Levantamento Fiscal", layout="wide")

# ----------------------------------------------------------------------
# Mapa de colunas (índice 0-based) conforme layout real das exportações
# ----------------------------------------------------------------------
F4M_COLS = {
    "chave": 0, "filial": 1, "descricao": 2, "tp_operacao": 3, "tes": 4,
    "estado": 5, "ncm": 15, "cfop": 16, "calc_icms": 17, "cred_icms": 18,
    "lf_icms": 19, "red_icms": 20, "cst_icms": 21, "calc_difal": 22,
    "calc_ipi": 23, "cred_ipi": 24, "lf_ipi": 25, "cod_trib_ipi": 26,
    "pis_cofins": 27, "cred_piscof": 28, "cst_pis": 29, "cst_cof": 30,
    "status": 31,
}
SF4_COLS = {
    "tes": 4, "tipo_tes": 5, "calc_icms": 6, "calc_ipi": 7, "cred_icms": 8,
    "cred_ipi": 9, "gera_dupl": 10, "atu_estoque": 11, "cfop": 12,
    "txt_padrao": 14, "finalidade": 15, "red_icms": 16, "lf_icms": 19,
    "lf_ipi": 20, "cst_icms": 45, "pis_cofins": 46, "cred_piscof": 47,
    "cod_trib_ipi": 80, "cst_pis": 233, "cst_cof": 234,
}
F3K_COLS = {
    "filial": 0, "produto": 1, "cfop": 2, "cbenef": 3, "valor": 4,
    "cst": 5, "cod_reflexo": 6, "grp_clientes": 7, "grp_produto": 8,
    "grp_fornec": 9, "obs_lanc_fis": 10, "cod_lanc": 11,
}
# CC7 (códigos de lançamento fiscal) — export sem linha de cabeçalho
CC7_COLS = {
    "filial": 0, "cod_lanc": 1, "tipo": 2, "codigo": 3,
    "descricao": 5, "indicador": 6, "data": 7, "valor": 8,
}

# Campos de cálculo comparados no cruzamento F4M x SF4
CAMPOS_COMPARA = [
    ("calc_icms", "Calcula ICMS"), ("cred_icms", "Credita ICMS"),
    ("red_icms", "% Red. ICMS"), ("cst_icms", "CST ICMS"),
    ("calc_ipi", "Calcula IPI"), ("cred_ipi", "Credita IPI"),
    ("cod_trib_ipi", "Cód.Trib.IPI"), ("pis_cofins", "PIS/COFINS"),
    ("cred_piscof", "Cred.PIS/COF"), ("cst_pis", "CST PIS"),
    ("cst_cof", "CST COFINS"),
]

# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------
def _norm(v):
    """Normaliza valor para comparação (trim, str, sem acento de Nao/Näo)."""
    if v is None:
        return ""
    s = str(v).strip()
    s = s.replace("ä", "a").replace("Ä", "A").replace("ã", "a")
    if s.replace(".", "", 1).replace("-", "", 1).isdigit():
        try:
            f = float(s)
            return str(int(f)) if f == int(f) else str(f)
        except ValueError:
            pass
    return s.lower()


def _read(file, sheet=None, cols=None):
    xls = pd.ExcelFile(file)
    sh = sheet if sheet in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sh, header=0, dtype=object)
    if cols:
        out = pd.DataFrame()
        for name, idx in cols.items():
            out[name] = df.iloc[:, idx] if idx < df.shape[1] else None
        return out
    return df


def _read_no_header(file, cols):
    """Leitura para exports sem linha de cabeçalho (ex.: CC7)."""
    xls = pd.ExcelFile(file)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, dtype=object)
    out = pd.DataFrame()
    for name, idx in cols.items():
        out[name] = df.iloc[:, idx] if idx < df.shape[1] else None
    return out


def cruzar(f4m, sf4):
    """Cruza cada regra da TES Inteligente com a TES da SF4."""
    sf4_idx = {}
    for _, r in sf4.iterrows():
        k = _norm(r["tes"])
        if k:
            sf4_idx[k] = r
    linhas = []
    for _, r in f4m.iterrows():
        tes = r["tes"]
        base = {
            "TES": tes, "CFOP": r["cfop"], "NCM": r["ncm"],
            "UF": r["estado"], "Descrição": r["descricao"],
        }
        srow = sf4_idx.get(_norm(tes))
        if srow is None:
            base["Status"] = "TES não encontrada na SF4"
            base["Divergências"] = "—"
            linhas.append(base)
            continue
        divs = []
        for key, label in CAMPOS_COMPARA:
            a, b = _norm(r.get(key)), _norm(srow.get(key))
            # Se o campo não veio populado na SF4, não compara (evita falso positivo)
            if b in ("", "nan", "none"):
                continue
            if a != b:
                divs.append(f"{label}: F4M='{r.get(key)}' x SF4='{srow.get(key)}'")
        base["Status"] = "OK" if not divs else f"Divergente ({len(divs)})"
        base["Divergências"] = " | ".join(divs) if divs else "—"
        linhas.append(base)
    return pd.DataFrame(linhas)


def enriquecer_f3k(f3k):
    if f3k is None:
        return None
    df = f3k.copy()
    df = df.rename(columns={
        "cfop": "CFOP", "cbenef": "cBenef", "cst": "CST",
        "cod_reflexo": "Cód.Reflexo", "cod_lanc": "Cód.Lançamento",
    })
    return df


def processar_cc7(cc7, f3k):
    """Formata a CC7 e marca quais códigos estão vinculados a algum cBenef da F3K."""
    if cc7 is None:
        return None
    df = cc7.copy()
    # cBenefs presentes na F3K (col 'cbenef' original)
    cbenefs = set()
    if f3k is not None:
        col = "cBenef" if "cBenef" in f3k.columns else "cbenef"
        cbenefs = {str(v).strip() for v in f3k[col].dropna()}
    df["Na F3K"] = df["codigo"].apply(
        lambda v: "Sim" if str(v).strip() in cbenefs else "Não"
    )
    df = df.rename(columns={
        "filial": "Filial", "cod_lanc": "Cód.Lançamento", "tipo": "Tipo",
        "codigo": "Código/cBenef", "descricao": "Descrição",
        "indicador": "Ind.", "valor": "Valor",
    })
    return df[["Filial", "Cód.Lançamento", "Tipo", "Código/cBenef",
               "Descrição", "Ind.", "Valor", "Na F3K"]]


# ----------------------------------------------------------------------
# Geração do documento MIT041 (.docx)
# ----------------------------------------------------------------------
def _clean(v):
    """Remove nan/None/'. .' para célula em branco."""
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() in ("nan", "none", "nat", ". .", ".."):
        return ""
    return s


def _set_cell(cell, text, bold=False, color=None, size=8):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(_clean(text))
    run.bold = bold
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor(*color)


def gerar_docx(meta, df_cruz, f3k, resumo, cc7=None):
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10)

    h = doc.add_heading("MIT041 - Especificação do Processo Fiscal (Levantamento)", level=0)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Ambientação
    doc.add_heading("1. Ambientação", level=1)
    amb = doc.add_table(rows=0, cols=4)
    amb.style = "Table Grid"
    pares = [
        ("Nome do cliente", meta.get("cliente"), "Código de cliente", meta.get("cod_cliente")),
        ("Nome do projeto", meta.get("projeto"), "Código do projeto", meta.get("cod_projeto")),
        ("Segmento cliente", meta.get("segmento"), "Unidade TOTVS", meta.get("unidade")),
        ("Data", meta.get("data"), "Proposta comercial", meta.get("proposta")),
        ("Gerente/Coordenador TOTVS", meta.get("ger_totvs"),
         "Gerente/Coordenador cliente", meta.get("ger_cliente")),
        ("Consultor TOTVS", meta.get("consultor"), "Filial base", meta.get("filial")),
        ("Módulo", "Fiscal (SIGAFIS)", "", ""),
    ]
    for lbl1, v1, lbl2, v2 in pares:
        cells = amb.add_row().cells
        _set_cell(cells[0], lbl1, bold=True, size=9)
        _set_cell(cells[1], v1, size=9)
        _set_cell(cells[2], lbl2, bold=True, size=9)
        _set_cell(cells[3], v2, size=9)

    # Resumo do cruzamento
    doc.add_heading("2. Resumo do Levantamento (TES Inteligente x SF4)", level=1)
    doc.add_paragraph(
        f"Total de amarrações analisadas: {resumo['total']}  |  "
        f"OK: {resumo['ok']}  |  Divergentes: {resumo['div']}  |  "
        f"TES não encontradas: {resumo['nf']}."
    )
    rt = doc.add_table(rows=1, cols=3)
    rt.style = "Table Grid"
    for i, c in enumerate(["Status", "Qtd", "%"]):
        _set_cell(rt.rows[0].cells[i], c, bold=True, size=9)
    for st_lbl, qtd in resumo["por_status"].items():
        pct = f"{(qtd / resumo['total'] * 100):.1f}%" if resumo["total"] else "0%"
        r = rt.add_row().cells
        _set_cell(r[0], st_lbl, size=9)
        _set_cell(r[1], qtd, size=9)
        _set_cell(r[2], pct, size=9)

    # Detalhe completo
    doc.add_heading("3. Detalhamento das Amarrações", level=1)
    cols = ["TES", "CFOP", "NCM", "UF", "Status", "Divergências"]
    dt = doc.add_table(rows=1, cols=len(cols))
    dt.style = "Table Grid"
    for i, c in enumerate(cols):
        _set_cell(dt.rows[0].cells[i], c, bold=True, size=8)
    for _, r in df_cruz.iterrows():
        cells = dt.add_row().cells
        status = str(r["Status"])
        color = (198, 40, 40) if ("Diverg" in status or "não" in status) else (46, 125, 50)
        for i, c in enumerate(cols):
            _set_cell(cells[i], r[c], color=color if c == "Status" else None, size=8)

    # cBenef / F3K
    doc.add_heading("4. Código de Benefício Fiscal (cBenef - F3K)", level=1)
    if f3k is not None and not f3k.empty:
        cols3 = ["CFOP", "cBenef", "CST", "Cód.Reflexo", "Cód.Lançamento"]
        cols3 = [c for c in cols3 if c in f3k.columns]
        ft = doc.add_table(rows=1, cols=len(cols3))
        ft.style = "Table Grid"
        for i, c in enumerate(cols3):
            _set_cell(ft.rows[0].cells[i], c, bold=True, size=8)
        for _, r in f3k.head(200).iterrows():
            cells = ft.add_row().cells
            for i, c in enumerate(cols3):
                _set_cell(cells[i], r[c], size=8)
    else:
        doc.add_paragraph("Nenhum registro de cBenef importado.")

    # CC7 - placeholder
    doc.add_heading("5. Códigos de Lançamento (CC7)", level=1)
    if cc7 is not None and not cc7.empty:
        vinc = (cc7["Na F3K"] == "Sim").sum()
        doc.add_paragraph(
            f"Total de códigos de lançamento: {len(cc7)}  |  "
            f"vinculados a cBenef da F3K: {vinc}."
        )
        cols5 = ["Cód.Lançamento", "Tipo", "Código/cBenef", "Descrição",
                 "Ind.", "Valor", "Na F3K"]
        ct = doc.add_table(rows=1, cols=len(cols5))
        ct.style = "Table Grid"
        for i, c in enumerate(cols5):
            _set_cell(ct.rows[0].cells[i], c, bold=True, size=8)
        for _, r in cc7.head(300).iterrows():
            cells = ct.add_row().cells
            cor = (46, 125, 50) if r["Na F3K"] == "Sim" else (198, 40, 40)
            for i, c in enumerate(cols5):
                _set_cell(cells[i], r[c],
                          color=cor if c == "Na F3K" else None, size=8)
    else:
        doc.add_paragraph("Nenhum código de lançamento (CC7) importado.")

    doc.add_heading("6. Aceite", level=1)
    doc.add_paragraph(
        "Confirmo que os processos e amarrações fiscais descritos neste documento "
        "refletem as necessidades da operação."
    )
    ac = doc.add_table(rows=2, cols=3)
    ac.style = "Table Grid"
    for i, c in enumerate(["Aprovado por", "Assinatura", "Data"]):
        _set_cell(ac.rows[0].cells[i], c, bold=True, size=9)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ----------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------
st.title("📋 MIT041 · Automação do Levantamento Fiscal")
st.caption("TOTVS Protheus · Cruzamento TES Inteligente (F4M) × Cadastro de TES (SF4) · cBenef (F3K) · CC7")

with st.sidebar:
    st.header("⚙️ Ambientação")
    cliente = st.text_input("Nome do cliente", "")
    cod_cliente = st.text_input("Código de cliente", "")
    projeto = st.text_input("Nome do projeto", "")
    cod_projeto = st.text_input("Código do projeto", "")
    segmento = st.text_input("Segmento cliente", "")
    unidade = st.text_input("Unidade TOTVS", "")
    data_ref = st.date_input("Data", datetime.today())
    proposta = st.text_input("Proposta comercial", "")
    ger_totvs = st.text_input("Gerente/Coordenador TOTVS", "")
    ger_cliente = st.text_input("Gerente/Coordenador cliente", "")
    consultor = st.text_input("Consultor TOTVS", "Thays Caroline")
    filial = st.text_input("Filial base", "")

st.subheader("1️⃣ Importação das tabelas")
c1, c2 = st.columns(2)
with c1:
    up_f4m = st.file_uploader("TES Inteligente — F4M (aba SFM)", type=["xlsx"], key="f4m")
    up_f3k = st.file_uploader("cBenef — F3K", type=["xlsx"], key="f3k")
with c2:
    up_sf4 = st.file_uploader("Cadastro de TES — SF4", type=["xlsx"], key="sf4")
    up_cc7 = st.file_uploader("Códigos de Lançamento — CC7", type=["xlsx"], key="cc7")

if up_cc7 and not (up_f4m and up_sf4):
    st.info("CC7 será processada junto com o cruzamento assim que F4M e SF4 forem enviadas.")

if up_f4m and up_sf4:
    try:
        f4m = _read(up_f4m, sheet="SFM", cols=F4M_COLS)
        sf4 = _read(up_sf4, sheet="SF4", cols=SF4_COLS)
        f3k = enriquecer_f3k(_read(up_f3k, cols=F3K_COLS)) if up_f3k else None
        cc7 = processar_cc7(_read_no_header(up_cc7, CC7_COLS), f3k) if up_cc7 else None
    except Exception as e:
        st.error(f"Erro ao ler as planilhas: {e}")
        st.stop()

    f4m = f4m[f4m["tes"].notna()]
    sf4 = sf4[sf4["tes"].notna()]

    n_bruto = len(f4m)
    dedup = st.checkbox(
        "Deduplicar amarrações iguais (TES+CFOP+NCM+UF)", value=True,
        help="A aba SFM repete a mesma regra por produto/filial. "
             "Recomendado para o relatório não ficar com milhares de linhas idênticas.",
    )
    if dedup:
        f4m = f4m.drop_duplicates(subset=["tes", "cfop", "ncm", "estado"])

    st.success(
        f"F4M: {len(f4m)} amarrações distintas (de {n_bruto} linhas) · "
        f"SF4: {len(sf4)} TES" +
        (f" · F3K: {len(f3k)} cBenef" if f3k is not None else "")
    )

    df_cruz = cruzar(f4m, sf4)

    total = len(df_cruz)
    por_status = df_cruz["Status"].apply(
        lambda s: "OK" if s == "OK" else ("Divergente" if "Diverg" in s else "TES não encontrada")
    ).value_counts().to_dict()
    resumo = {
        "total": total,
        "ok": por_status.get("OK", 0),
        "div": por_status.get("Divergente", 0),
        "nf": por_status.get("TES não encontrada", 0),
        "por_status": por_status,
    }

    st.subheader("2️⃣ Resumo do cruzamento")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Amarrações", total)
    m2.metric("✅ OK", resumo["ok"])
    m3.metric("⚠️ Divergentes", resumo["div"])
    m4.metric("❌ TES ausente", resumo["nf"])

    st.subheader("3️⃣ Detalhamento")
    filtro = st.multiselect(
        "Filtrar status",
        options=sorted(df_cruz["Status"].unique()),
        default=list(sorted(df_cruz["Status"].unique())),
    )
    st.dataframe(df_cruz[df_cruz["Status"].isin(filtro)], use_container_width=True, height=380)

    if f3k is not None:
        st.subheader("4️⃣ cBenef (F3K)")
        st.dataframe(f3k, use_container_width=True, height=240)

    if cc7 is not None:
        st.subheader("5️⃣ Códigos de Lançamento (CC7)")
        vinc = (cc7["Na F3K"] == "Sim").sum()
        st.caption(f"{len(cc7)} códigos · {vinc} vinculados a cBenef da F3K")
        st.dataframe(cc7, use_container_width=True, height=240)

    st.subheader("5️⃣ Gerar MIT041")
    if st.button("📄 Gerar documento .docx", type="primary"):
        meta = {
            "cliente": cliente, "cod_cliente": cod_cliente,
            "projeto": projeto, "cod_projeto": cod_projeto,
            "segmento": segmento, "unidade": unidade,
            "data": data_ref.strftime("%d/%m/%Y"), "proposta": proposta,
            "ger_totvs": ger_totvs, "ger_cliente": ger_cliente,
            "consultor": consultor, "filial": filial,
        }
        buf = gerar_docx(meta, df_cruz, f3k, resumo, cc7)
        nome = f"MIT041_Fiscal_{(cliente or 'cliente').replace(' ', '_')}.docx"
        st.download_button("⬇️ Baixar MIT041", buf, file_name=nome,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        csv = df_cruz.to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button("⬇️ Baixar cruzamento (CSV)", csv,
                           file_name="cruzamento_tes.csv", mime="text/csv")
else:
    st.info("Envie ao menos a **F4M (aba SFM)** e a **SF4** para iniciar o cruzamento.")
