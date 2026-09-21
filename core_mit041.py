# -*- coding: utf-8 -*-
"""core_mit041.py — motor de análise + preenchimento do modelo MIT041.

Não gera documento do zero: ABRE o modelo padrão TOTVS e preenche
os campos, preservando 100% do layout. Appenda as tabelas de dados
(SFM/F4M, SF4, F3K, CC7, SFT) ao final, antes do Aceite.
"""

import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor

# ----------------------------------------------------------------------
# Mapa de colunas (0-based) conforme layout real das exportações
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
    "lf_ipi": 20, "calc_difal": 23, "cst_icms": 45, "pis_cofins": 46,
    "cred_piscof": 47, "cod_trib_ipi": 80, "bs_icms_st": 40,
    "funrural": 202, "lf_ciap": 34, "trib_ciap": 154, "prot_go": 238,
    "cst_pis": 233, "cst_cof": 234,
}
F3K_COLS = {
    "filial": 0, "produto": 1, "cfop": 2, "cbenef": 3, "valor": 4,
    "cst": 5, "cod_reflexo": 6, "grp_clientes": 7, "grp_produto": 8,
    "grp_fornec": 9, "obs_lanc_fis": 10, "cod_lanc": 11,
}
CC7_COLS = {
    "filial": 0, "cod_lanc": 1, "tipo": 2, "codigo": 3,
    "descricao": 5, "indicador": 6, "data": 7, "valor": 8,
}
# SFT (Exceções Fiscais) — mapa provisório; ajustar quando o arquivo chegar
SFT_COLS = None


# ----------------------------------------------------------------------
# Leitura
# ----------------------------------------------------------------------
def _read(file, sheet=None, cols=None):
    xls = pd.ExcelFile(file)
    sh = sheet if (sheet and sheet in xls.sheet_names) else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sh, header=0, dtype=object)
    if cols:
        out = pd.DataFrame()
        for name, idx in cols.items():
            out[name] = df.iloc[:, idx] if idx < df.shape[1] else None
        return out
    return df


def _read_no_header(file, cols):
    xls = pd.ExcelFile(file)
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, dtype=object)
    out = pd.DataFrame()
    for name, idx in cols.items():
        out[name] = df.iloc[:, idx] if idx < df.shape[1] else None
    return out


def _sim(series):
    """True se alguma linha tem 'Sim' (tolerante a acento/caixa)."""
    return series.astype(str).str.strip().str.lower().isin(["sim", "s"]).any()


def _distintos(series, ignora=("", "nan", "none", "0")):
    vals = []
    for v in series.dropna():
        s = str(v).strip()
        if s.lower() not in ignora and s not in vals:
            vals.append(s)
    return vals


# ----------------------------------------------------------------------
# Análise: responde as perguntas do questionário a partir das tabelas
# ----------------------------------------------------------------------
def analisar(f4m=None, sf4=None, f3k=None, cc7=None):
    """Retorna dict {item: resposta} para as perguntas respondíveis."""
    r = {}

    if f4m is not None:
        if _sim(f4m["calc_icms"]):
            r["2.1"] = "Sim — há regras de TES com cálculo de ICMS."
        reds = [x for x in _distintos(f4m["red_icms"]) if x not in ("0", "0.0")]
        if reds:
            r["2.3"] = "Sim — há % de redução de ICMS distintos: " + ", ".join(reds[:8]) + "."
        else:
            r["2.3"] = "Não identificado nas TES analisadas."
        if _sim(f4m["calc_ipi"]):
            r["2.20"] = "Sim — há TES com cálculo de IPI."
        else:
            r["2.20"] = "Não — nenhuma TES com cálculo de IPI nas regras analisadas."
        csts = sorted(set(_distintos(f4m["cst_pis"]) + _distintos(f4m["cst_cof"])))
        if csts:
            r["2.15"] = "Sim — CSTs de PIS/COFINS identificados: " + ", ".join(csts[:12]) + "."

    if sf4 is not None:
        # ST
        st_flag = _sim(sf4["bs_icms_st"].apply(lambda v: "sim" if str(v).strip() not in ("", "nan", "None") else "")) \
            if "bs_icms_st" in sf4 else False
        csticm = _distintos(sf4["cst_icms"])
        tem_st = any(c in ("10", "30", "60", "70") for c in csticm)
        r["2.4"] = "Sim — há TES com Substituição Tributária (CST/base ST)." if tem_st \
            else "Não identificado nas TES analisadas."
        # DIFAL
        difal = _sim(sf4["calc_difal"]) if "calc_difal" in sf4 else False
        if difal:
            r["2.5"] = "Sim — há TES com cálculo de Diferencial de Alíquota."
        # FUNRURAL
        if "funrural" in sf4:
            fr = sf4["funrural"].astype(str).str.contains("Rural", case=False, na=False).any()
            r["2.7"] = "Sim — há TES com base FUNRURAL (Produto Rural)." if fr \
                else "Não identificado nas TES analisadas."
        # FECP
        if "prot_go" in sf4:
            fecp = pd.to_numeric(sf4["prot_go"], errors="coerce").fillna(0).gt(0).any()
            r["2.9"] = "Sim — há alíquota de fundo de combate à pobreza (FECP/PROT) > 0." if fecp \
                else "Não identificado nas TES analisadas."
        # CIAP
        ciap = False
        if "lf_ciap" in sf4:
            ciap = ciap or _sim(sf4["lf_ciap"])
        if "trib_ciap" in sf4:
            ciap = ciap or _sim(sf4["trib_ciap"])
        r["2.12"] = "Sim — há TES com controle de CIAP." if ciap \
            else "Não identificado nas TES analisadas."

    if f3k is not None:
        col = "cbenef" if "cbenef" in f3k.columns else "cBenef"
        cbenefs = _distintos(f3k[col])
        if cbenefs:
            r["2.2"] = ("Sim — códigos de benefício fiscal (cBenef) identificados na F3K: "
                        + ", ".join(cbenefs[:10])
                        + (" ..." if len(cbenefs) > 10 else "") + ".")

    # DIFAL reforçado pela CC7
    if cc7 is not None:
        desc = cc7["descricao"].astype(str).str.contains("DIFERENCIAL", case=False, na=False).any() \
            if "descricao" in cc7 else False
        if desc:
            r["2.5"] = ("Sim — há código de lançamento de Diferencial de Alíquota na CC7"
                        + (" e cálculo configurado nas TES." if r.get("2.5") else "."))

    return r


# ----------------------------------------------------------------------
# Preenchimento do modelo (sem alterar layout)
# ----------------------------------------------------------------------
# Ordem das perguntas do questionário -> item numerado
QUEST_ITENS = [
    ("1.1", "Qual será o Regime de Tributação"),
    ("1.2", "Como está organizado o departamento fiscal"),
    ("1.3", "Quais sistemas e software fiscal"),
    ("1.4", "Qual quantidade média de notas fiscais"),
    ("2.1", "Contribuinte do ICMS"),
    ("2.2", "Possui algum tare, benefício ou incentivo"),
    ("2.3", "Alíquotas diferenciadas de ICMS"),
    ("2.4", "Substituição Tributária"),
    ("2.5", "Diferencial de Alíquota"),
    ("2.6", "Nota fiscal de Importação"),
    ("2.7", "Produtor Rural"),
    ("2.8", "Zona Franca de Manaus"),
    ("2.9", "fundo de combate"),
    ("2.10", "Órgão Público"),
    ("2.11", "código de ajustes específico na apuração"),
    ("2.12", "Controle de CIAP"),
    ("2.13", "principais erros de apuração"),
    ("2.15", "Contribuinte de PIS COFINS"),
    ("2.16", "movimentos financeiros"),
    ("2.17", "Ativo fixo"),
    ("2.18", "diferimento ou crédito por decisão"),
    ("2.20", "Contribuinte do IPI"),
    ("2.21", "IPI Presumido"),
    ("2.22", "código de ajustes específico na apuração"),
    ("2.24", "Contribuinte do ISS"),
    ("2.25", "municípios para apuração de ISS"),
    ("2.26", "redução na base de Cálculo de ISS"),
    ("2.27", "retenção de ISS"),
    ("2.28", "cliente Órgão público"),
    ("2.29", "retenções de impostos"),
    ("2.30", "base de Cálculo do INSS"),
    ("2.31", "alíquota de IR diferenciada"),
    ("2.32", "retenção do CSRF"),
    ("2.34", "SPED FISCAL"),
    ("2.35", "bloco em específico"),
    ("2.36", "EFD CONTRIBUIÇÕES"),
    ("2.37", "REINF"),
    ("2.38", "evento específico da REINF"),
    ("2.39", "acessórias Federais"),
    ("2.40", "acessórias estaduais"),
    ("2.41", "filiais em outras UF"),
    ("2.42", "acessórias Municipais"),
]


def _set_resp(cell, texto):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(texto or "")
    run.font.size = Pt(9)


def preencher_modelo(modelo_path, meta, respostas_auto, respostas_manual,
                     f4m=None, sf4=None, f3k=None, cc7=None, sft=None):
    """Abre o modelo, preenche Ambientação + questionário, appenda tabelas."""
    doc = Document(modelo_path)

    # --- Ambientação (Tabela 0) ---
    amb = doc.tables[0]
    mapa_amb = {
        "Nome do cliente": meta.get("cliente"),
        "Código de cliente": meta.get("cod_cliente"),
        "Nome do projeto": meta.get("projeto"),
        "Código do projeto": meta.get("cod_projeto"),
        "Segmento cliente": meta.get("segmento"),
        "Unidade TOTVS": meta.get("unidade"),
        "Data": meta.get("data"),
        "Proposta comercial": meta.get("proposta"),
    }
    for row in amb.rows:
        for cell in row.cells:
            txt = cell.text.strip().rstrip(":")
            for k, v in mapa_amb.items():
                if txt == k and v:
                    # acrescenta o valor ao rótulo existente
                    cell.paragraphs[0].add_run(" " + str(v))

    # --- Questionário (Tabela 3) ---
    quest = doc.tables[3]
    todas = {**respostas_auto, **{k: v for k, v in respostas_manual.items() if v}}
    idx = 0  # ponteiro na lista QUEST_ITENS
    for row in quest.rows:
        cells = row.cells
        # pula cabeçalho e linhas de seção (col1==col2==col3)
        if cells[1].text.strip() == cells[-1].text.strip() and cells[1].text.strip():
            continue
        perg = cells[1].text.strip()
        if not perg or perg == "Item":
            continue
        # casa a pergunta com o item pela ordem
        while idx < len(QUEST_ITENS):
            item, chave = QUEST_ITENS[idx]
            if chave.lower() in perg.lower():
                if item in todas:
                    _set_resp(cells[-1], todas[item])
                idx += 1
                break
            idx += 1

    # --- Numera a coluna Item do questionário ---
    _numerar_itens(quest)

    # --- Appenda tabelas de dados antes do Aceite ---
    _append_dados(doc, f4m, sf4, f3k, cc7, sft)

    return doc


def _numerar_itens(quest):
    """Preenche a 1ª coluna com a numeração 1, 1.1... conforme o padrão."""
    itens_ordem = [i for i, _ in QUEST_ITENS]
    ptr = 0
    for row in quest.rows:
        cells = row.cells
        c0 = cells[0]
        secao = cells[1].text.strip() == cells[-1].text.strip() and cells[1].text.strip()
        perg = cells[1].text.strip()
        if perg == "Item" or not perg:
            continue
        if secao:
            continue
        if ptr < len(itens_ordem) and not c0.text.strip():
            _set_resp(c0, itens_ordem[ptr])
        ptr += 1


def _hdr(doc, texto):
    doc.add_heading(texto, level=2)


def _tabela(doc, df, cols, limite=300):
    t = doc.add_table(rows=1, cols=len(cols))
    t.style = "Table Grid"
    for i, c in enumerate(cols):
        cell = t.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(c)
        run.bold = True
        run.font.size = Pt(8)
    for _, r in df.head(limite).iterrows():
        cells = t.add_row().cells
        for i, c in enumerate(cols):
            v = r.get(c, "")
            s = "" if (v is None or str(v).strip().lower() in ("nan", "none")) else str(v)
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(s)
            run.font.size = Pt(8)


def _append_dados(doc, f4m, sf4, f3k, cc7, sft):
    doc.add_page_break()
    doc.add_heading("Anexo — Levantamento das Tabelas (carga)", level=1)

    if f4m is not None:
        _hdr(doc, "TES Inteligente (F4M / SFM)")
        d = f4m.rename(columns={"tes": "TES", "cfop": "CFOP", "ncm": "NCM",
                                "estado": "UF", "descricao": "Descrição"})
        _tabela(doc, d, ["TES", "CFOP", "NCM", "UF", "Descrição"])

    if sf4 is not None:
        _hdr(doc, "Cadastro de TES (SF4)")
        d = sf4.rename(columns={"tes": "TES", "tipo_tes": "Tipo", "cfop": "CFOP",
                                "finalidade": "Finalidade", "cst_icms": "CST ICMS"})
        _tabela(doc, d, ["TES", "Tipo", "CFOP", "CST ICMS", "Finalidade"])

    if f3k is not None:
        _hdr(doc, "cBenef (F3K)")
        _tabela(doc, f3k, ["CFOP", "cBenef", "CST", "Cód.Reflexo", "Cód.Lançamento"])

    if cc7 is not None:
        _hdr(doc, "Códigos de Lançamento (CC7)")
        _tabela(doc, cc7, ["Cód.Lançamento", "Tipo", "Código/cBenef",
                           "Descrição", "Ind.", "Valor", "Na F3K"])

    if sft is not None:
        _hdr(doc, "Exceções Fiscais (SFT)")
        _tabela(doc, sft, list(sft.columns))
