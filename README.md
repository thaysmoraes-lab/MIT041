# MIT041 · Automação do Levantamento Fiscal

App Streamlit que **preenche o modelo padrão MIT041** (layout TOTVS imutável)
a partir das tabelas do Protheus, respondendo automaticamente as perguntas do
questionário fiscal que os dados permitem e oferecendo campos em tela para as
demais.

## Como funciona
1. Abre `modelo/MIT041_modelo.docx` (não gera do zero — preserva 100% do layout).
2. Preenche a **Ambientação** e numera/responde o **questionário** (Tabela 3).
3. Appenda as tabelas de dados (F4M, SF4, F3K, CC7, SFT) como **Anexo**.

## Tabelas e mapeamento
| Upload | Tabela | Aba | Observação |
|---|---|---|---|
| Cadastros TES e TES Inteligente | F4M | `SFM` | dedup por TES+CFOP+NCM+UF |
| SF4 | SF4 | `SF4` | código da TES na col 4 |
| F3K | cBenef | 1ª aba | |
| CC7 | cód. lançamento | 1ª aba | **sem cabeçalho** (1ª linha já é dado) |
| SFT | exceções fiscais | 1ª aba | *mapeamento a definir* |

## Perguntas respondidas automaticamente
2.1 (ICMS), 2.2 (cBenef/F3K), 2.3 (red. ICMS), 2.4 (ST), 2.5 (DIFAL),
2.7 (FUNRURAL), 2.9 (FECP), 2.12 (CIAP), 2.15 (CST PIS/COFINS), 2.20 (IPI).
As demais têm campo em tela.

## Estrutura do repositório
```
app.py
core_mit041.py
requirements.txt
modelo/MIT041_modelo.docx   <-- OBRIGATÓRIO (o app lê daqui)
```

## Deploy (Streamlit Cloud)
1. Suba os 4 itens acima no GitHub (inclusive a pasta `modelo/`).
2. share.streamlit.io → New app → aponte para `app.py`.

## Local
```bash
pip install -r requirements.txt
streamlit run app.py
```
