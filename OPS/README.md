# Analisador de Abordagens de Segurança (BBS)

Aplicativo em Streamlit que importa uma planilha de observações comportamentais
de segurança, calcula a **criticidade** de cada abordagem e a **cruza com o local**
(Processo / Sub-processo / Área). Nenhum nome de pessoa é exibido nos dashboards;
empresas parceiras são anonimizadas em códigos (Empresa 01, 02, ...).

## Instalação e execução

```bash
pip install -r requirements.txt
streamlit run app.py
```

Abra o endereço exibido (geralmente http://localhost:8501) e importe a planilha.
Aceita **Excel** (.xlsx, .xls, .xlsm) e **CSV** (detecta automaticamente encoding
Latin-1/UTF-8 e separador `;` ou `,`).

## Colunas reconhecidas

Processo, Sub-processo, Área, Tipo, Público, Parceiro/Empresa, Desvio,
Atividade Observada, Descrição da atividade, Descrição da abordagem, Data.
Os nomes podem variar (com/sem acento); o app faz o mapeamento automático.

## Como a criticidade é calculada (0 a 100)

| Fator | Peso | Lógica |
|-------|------|--------|
| Risco da atividade observada | 5 | Altura, espaço confinado, bloqueio de energia e eletricidade pesam mais |
| Gravidade do desvio | 4 | Trabalho sem ATC/ATE, linha de tiro, área de risco, EPI de altura são os mais graves |
| Palavras-chave no texto | 3 | queda, energizado, vapor, pressão, explosão, choque, asfixia, etc. |
| Comportamento inseguro | 3 | Marcação do tipo |

Faixas: Baixa (<40), Média (40–59), Alta (60–79), Crítica (≥80).
Os pesos e listas de palavras estão no topo do `app.py`, fáceis de ajustar.

## Dashboards

1. **Visão Geral** — KPIs, abordagens por processo, atividades observadas, faixas, famílias de desvio, evolução temporal.
2. **Análise de Criticidade** — ranking de atividades mais críticas, criticidade por público e por empresa (anonimizada).
3. **Local × Criticidade** — mapa de calor local × faixa e matriz de prioridade (volume × criticidade).
4. **Abordagens Críticas** — texto integral das abordagens de maior risco (sem nomes) + exportação CSV.
