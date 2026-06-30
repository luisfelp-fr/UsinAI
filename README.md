# UsinAI

Coleção de aplicativos Streamlit organizados por área.

## 📂 Estrutura

| Pasta | Aplicativo | Descrição |
|-------|------------|-----------|
| [`Analisador RI/`](./Analisador%20RI/) | Analisador de Sensibilidade — Recuperação Industrial | Lê os indicadores diários, reconstrói a cadeia de cálculo da Recuperação Industrial e roda uma análise de sensibilidade, mostrando o % de influência de cada indicador. |
| [`OPS/`](./OPS/) | Analisador de Abordagens de Segurança (BBS) | Importa observações comportamentais de segurança, calcula a criticidade de cada abordagem e cruza com o local (Processo / Sub-processo / Área). Veja o [README da pasta](./OPS/README.md). |

## ▶️ Como executar

Cada pasta é autocontida (tem seu próprio `app.py` e `requirements.txt`):

```bash
# Analisador RI
pip install -r "Analisador RI/requirements.txt"
streamlit run "Analisador RI/app.py"

# OPS (BBS)
pip install -r "OPS/requirements.txt"
streamlit run "OPS/app.py"
```
