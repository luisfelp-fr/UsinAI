# Espelho de Simulacao

Aplicativo Streamlit que importa um arquivo **XML de simulacao** de usina
(raiz `<Simulation>`, com as secoes `Correntes`, `Operacoes` e `Sequencia`) e
extrai os indicadores operacionais, reproduzindo a logica da planilha
`Modelo_Espelho`.

## O que ele faz

O XML descreve as **correntes** do processo (cada `<Corrente>` tem atributos como
`W`, `WVol`, `T`, `POL`, `AR`, `pot` e componentes com `brix`, `pureza`,
`fracaoMassica`). O app localiza cada corrente **pelo nome** (equivalente ao
`PROCV` da planilha), le o atributo desejado e aplica a transformacao de unidade
correspondente.

Duas abas reproduzem as abas da planilha:

- **Espelho SimDiaria** — tabela com ~70 indicadores agrupados por etapa
  (Cana/Qualidade, Extracao, Tratamento, Concentracao, Fabrica de Acucar,
  Producao de Etanol, Utilidades/Energia). Inclui exportacao em CSV.
- **Espelho Visual** — os mesmos numeros apresentados sobre um **fluxograma de
  processo** (Graphviz), da cana ao produto final e a energia exportada.

Correntes ausentes no XML aparecem como `-` (equivalente ao `#N/A` do `PROCV`).

## Como executar

```bash
pip install -r requirements.txt
streamlit run app.py
```

Depois, na barra lateral, envie o arquivo `.xml` da simulacao.

## Mapeamento de atributos (referencia)

| Rotulo na planilha | Atributo no XML | Nivel      | Transformacao |
|--------------------|-----------------|------------|---------------|
| W                  | `W`             | Corrente   | x24 (t/dia)   |
| Wvol               | `WVol`          | Corrente   | direto (m3/h) |
| T                  | `T`             | Corrente   | -273 (C)      |
| POL / ar           | `POL` / `AR`    | Corrente   | direto (%)    |
| pot                | `pot`           | Corrente   | MW (neg -> 0) |
| brix / pureza      | `brix`/`pureza` | Componente | x100 (%)      |
| fracaoMassica      | `fracaoMassica` | Componente | x100 (%)      |
