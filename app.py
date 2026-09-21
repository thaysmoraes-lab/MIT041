# -*- coding: utf-8 -*-
"""
MIT041 · Automação do Levantamento Fiscal (TOTVS Protheus)

Preenche o MODELO PADRÃO MIT041 (layout imutável) a partir das tabelas
SFM/F4M, SF4, F3K, CC7 e SFT: responde automaticamente as perguntas do
questionário que as tabelas permitem e disponibiliza campos em tela para
as demais. Appenda as tabelas de dados como anexo.

Frontend: Streamlit  |  Autor: Prime Flow
"""

import io
import os
from datetime import datetime

import pandas as pd
import streamlit as st

import core_mit041 as core

st.set_page_config(page_title="MIT041 · Levantamento Fiscal", layout="wide")

MODELO_PATH = os.path.join(os.path.dirname(__file__), "modelo", "MIT041_modelo.docx")

# Perguntas que NÃO são respondidas pelas tabelas -> campo em tela.
PERG_MANUAIS = [
    ("1.1", "Regime de Tributação (Real/Presumido/Simples)", "Cenário Geral"),
    ("1.2", "Como está organizado o departamento fiscal", "Cenário Geral"),
    ("1.3", "Sistemas/software fiscal utilizados", "Cenário Geral"),
    ("1.4", "Quantidade média de NF entrada/saída", "Cenário Geral"),
    ("2.6", "Emite NF de Importação?", "ICMS"),
    ("2.8", "Operações com Zona Franca de Manaus?", "ICMS"),
    ("2.10", "Vendem para Cliente Órgão Público?", "ICMS"),
    ("2.11", "Código de ajustes específico na apuração (ICMS)?", "ICMS"),
    ("2.13", "Principais erros de apuração hoje", "ICMS"),
    ("2.16", "PIS/COFINS sobre movimentos financeiros?", "PIS/COFINS"),
    ("2.17", "PIS/COFINS sobre Ativo Fixo?", "PIS/COFINS"),
    ("2.18", "Benefício/tare/diferimento/crédito judicial?", "PIS/COFINS"),
    ("2.21", "Benefício/IPI Presumido (Atacadista)?", "IPI"),
    ("2.22", "Código de ajustes específico na apuração (IPI)?", "IPI"),
    ("2.24", "Contribuinte do ISS?", "ISS"),
    ("2.25", "Lista de municípios para apuração de ISS", "ISS"),
    ("2.26", "Redução na base de cálculo de ISS?", "ISS"),
    ("2.27", "Modo de retenção de ISS (emissão/baixa)", "ISS"),
    ("2.28", "Serviço para Órgão Público? Particularidade?", "ISS"),
    ("2.29", "Há retenções de impostos? Quais?", "Retenções"),
    ("2.30", "Redução base INSS em serviços? % varia?", "Retenções"),
    ("2.31", "Alíquota de IR diferenciada? Quais?", "Retenções"),
    ("2.32", "Modo de retenção do CSRF (emissão/baixa)", "Retenções"),
    ("2.34", "Obrigados ao SPED FISCAL?", "Obrigações Acessórias"),
    ("2.35", "Algum bloco específico a destacar? (ex: K, B)", "Obrigações Acessórias"),
    ("2.36", "Obrigados à EFD CONTRIBUIÇÕES?", "Obrigações Acessórias"),
    ("2.37", "Obrigados à REINF?", "Obrigações Acessórias"),
    ("2.38", "Evento específico da REINF? (ex: CPRB)", "Obrigações Acessórias"),
    ("2.39", "Outras obrigações Federais? Quais?", "Obrigações Acessórias"),
    ("2.40", "Outras obrigações estaduais? Quais?", "Obrigações Acessórias"),
    ("2.41", "Considerando filiais em outras UFs", "Obrigações Acessórias"),
    ("2.42", "Obrigações Municipais? Quais?", "Obrigações Acessórias"),
]


def prep_f3k(up):
    raw = core._read(up, cols=core.F3K_COLS)
    disp = raw.rename(columns={
        "cfop": "CFOP", "cbenef": "cBenef", "cst": "CST",
        "cod_reflexo": "Cód.Reflexo", "cod_lanc": "Cód.Lançamento",
    })
    return raw, disp


def prep_cc7(up, f3k_raw):
    raw = core._read_no_header(up, core.CC7_COLS)
    cbset = set()
    if f3k_raw is not None:
        cbset = {str(v).strip() for v in f3k_raw["cbenef"].dropna()}
    disp = raw.rename(columns={
        "filial": "Filial", "cod_lanc": "Cód.Lançamento", "tipo": "Tipo",
        "codigo": "Código/cBenef", "descricao": "Descrição",
        "indicador": "Ind.", "valor": "Valor",
    })
    disp["Na F3K"] = disp["Código/cBenef"].apply(
        lambda v: "Sim" if str(v).strip() in cbset else "Não")
    return raw, disp[["Cód.Lançamento", "Tipo", "Código/cBenef",
                      "Descrição", "Ind.", "Valor", "Na F3K"]]


st.title("📋 MIT041 · Automação do Levantamento Fiscal")
st.caption("Preenche o modelo padrão TOTVS (layout imutável) · SFM/F4M · SF4 · F3K · CC7 · SFT")

if not os.path.exists(MODELO_PATH):
    st.error("Modelo `modelo/MIT041_modelo.docx` não encontrado no repositório.")
    st.stop()

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

meta = {
    "cliente": cliente, "cod_cliente": cod_cliente, "projeto": projeto,
    "cod_projeto": cod_projeto, "segmento": segmento, "unidade": unidade,
    "data": data_ref.strftime("%d/%m/%Y"), "proposta": proposta,
}

st.subheader("1️⃣ Importação das tabelas")
c1, c2, c3 = st.columns(3)
with c1:
    up_f4m = st.file_uploader("TES Inteligente — F4M (aba SFM)", type=["xlsx"], key="f4m")
    up_cc7 = st.file_uploader("Códigos de Lançamento — CC7", type=["xlsx"], key="cc7")
with c2:
    up_sf4 = st.file_uploader("Cadastro de TES — SF4", type=["xlsx"], key="sf4")
    up_sft = st.file_uploader("Exceções Fiscais — SFT", type=["xlsx"], key="sft")
with c3:
    up_f3k = st.file_uploader("cBenef — F3K", type=["xlsx"], key="f3k")

f4m = sf4 = f3k_raw = f3k = cc7_raw = cc7 = sft = None
try:
    if up_f4m:
        f4m = core._read(up_f4m, sheet="SFM", cols=core.F4M_COLS)
        f4m = f4m[f4m["tes"].notna()].drop_duplicates(subset=["tes", "cfop", "ncm", "estado"])
    if up_sf4:
        sf4 = core._read(up_sf4, sheet="SF4", cols=core.SF4_COLS)
        sf4 = sf4[sf4["tes"].notna()]
    if up_f3k:
        f3k_raw, f3k = prep_f3k(up_f3k)
    if up_cc7:
        cc7_raw, cc7 = prep_cc7(up_cc7, f3k_raw)
    if up_sft:
        sft = core._read(up_sft) if core.SFT_COLS is None else core._read(up_sft, cols=core.SFT_COLS)
except Exception as e:
    st.error(f"Erro ao ler as planilhas: {e}")
    st.stop()

if up_sft and core.SFT_COLS is None:
    st.info("SFT importada como anexo bruto. O mapeamento de colunas da SFT ainda não foi definido — envie o layout para eu vinculá-la às respostas.")

carregou = any(x is not None for x in [f4m, sf4, f3k, cc7])
if carregou:
    resumo = []
    if f4m is not None: resumo.append(f"F4M: {len(f4m)} amarrações")
    if sf4 is not None: resumo.append(f"SF4: {len(sf4)} TES")
    if f3k is not None: resumo.append(f"F3K: {len(f3k)} cBenef")
    if cc7 is not None: resumo.append(f"CC7: {len(cc7)} lançamentos")
    st.success(" · ".join(resumo))

respostas_auto = {}
if carregou:
    respostas_auto = core.analisar(f4m=f4m, sf4=sf4, f3k=f3k_raw, cc7=cc7_raw)
    st.subheader("2️⃣ Respostas automáticas (a partir das tabelas)")
    if respostas_auto:
        df_auto = pd.DataFrame(
            [{"Item": k, "Resposta": v} for k, v in sorted(respostas_auto.items())])
        st.dataframe(df_auto, use_container_width=True, hide_index=True)
    else:
        st.write("Nenhuma resposta automática gerada com as tabelas atuais.")

st.subheader("3️⃣ Perguntas a responder em tela")
st.caption("Não extraídas das tabelas. Preencha o que tiver; em branco fica em branco no documento.")
respostas_manual = {}
secao_atual = None
for item, rotulo, secao in PERG_MANUAIS:
    if secao != secao_atual:
        st.markdown(f"**{secao}**")
        secao_atual = secao
    respostas_manual[item] = st.text_input(f"{item} — {rotulo}", key=f"m_{item}")

if carregou:
    with st.expander("4️⃣ Tabelas de dados (anexo do documento)"):
        if f3k is not None:
            st.markdown("**cBenef (F3K)**")
            st.dataframe(f3k, use_container_width=True, height=200)
        if cc7 is not None:
            st.markdown("**Códigos de Lançamento (CC7)**")
            st.dataframe(cc7, use_container_width=True, height=200)
        if sft is not None:
            st.markdown("**Exceções Fiscais (SFT)**")
            st.dataframe(sft, use_container_width=True, height=200)

st.subheader("5️⃣ Gerar MIT041")
if not carregou:
    st.info("Envie ao menos a **F4M (aba SFM)** ou a **SF4** para habilitar a geração.")
else:
    if st.button("📄 Gerar MIT041 (modelo padrão preenchido)", type="primary"):
        try:
            doc = core.preencher_modelo(
                MODELO_PATH, meta, respostas_auto, respostas_manual,
                f4m=f4m, sf4=sf4, f3k=f3k, cc7=cc7, sft=sft)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            nome = f"MIT041_Fiscal_{(cliente or 'cliente').replace(' ', '_')}.docx"
            st.success("Documento gerado com o layout padrão preservado.")
            st.download_button(
                "⬇️ Baixar MIT041 (.docx)", buf, file_name=nome,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Erro ao gerar o documento: {e}")
