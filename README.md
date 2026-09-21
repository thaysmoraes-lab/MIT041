# MIT041 · Automação do Levantamento Fiscal

App Streamlit que cruza a **TES Inteligente (F4M)** com o **cadastro de TES (SF4)**,
incorpora **cBenef (F3K)** e **códigos de lançamento (CC7)**, e gera o documento
**MIT041** (.docx) para o levantamento fiscal em implantações TOTVS Protheus.

## Lógica do cruzamento
- **Chave:** `F4M.TES (col 4)` ↔ `SF4.Cod. do Tipo (col 4)`
- Compara os campos de cálculo: Calcula/Credita ICMS, %Red ICMS, CST ICMS,
  Calcula/Credita IPI, Cód.Trib.IPI, PIS/COFINS, CST PIS/COFINS.
- **Status por amarração:** OK · Divergente (lista as diferenças) · TES não encontrada.

## Arquivos esperados
| Upload | Tabela | Aba |
|---|---|---|
| Cadastros TES e TES Inteligente | F4M | `SFM` |
| SF4 | SF4 | `SF4` |
| F3K | cBenef | 1ª aba |
| CC7 | cód. lançamento | 1ª aba (**sem cabeçalho** — 1ª linha já é dado) |

## Cruzamentos secundários
- **CC7 → F3K:** cada código de lançamento é marcado "Na F3K = Sim/Não"
  conforme o `Código/cBenef` (col 3) exista entre os cBenef da F3K.

## Deploy no Streamlit Cloud
1. Suba `app.py` e `requirements.txt` num repositório GitHub.
2. Em share.streamlit.io → New app → aponte para o repo e `app.py`.
3. Deploy.

## Rodar local
```bash
pip install -r requirements.txt
streamlit run app.py
```
