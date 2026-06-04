# -*- coding: utf-8 -*-
"""
Analisador de Abordagens de Seguranca (BBS - Observacao Comportamental)
-----------------------------------------------------------------------
Importa uma planilha (Excel ou CSV) com observacoes de seguranca em atividades
observadas, calcula um indice de criticidade por abordagem e cruza com o LOCAL
da abordagem (Processo / Sub-processo / Area).

Privacidade: nenhum nome de pessoa e exibido. Empresas/parceiros sao
anonimizados em codigos (Empresa 01, 02, ...).

Como executar:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import re
import sys
import unicodedata

import altair as alt
import numpy as np
import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Pagina
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Analisador de Abordagens de Seguranca",
                   page_icon="⛑️", layout="wide")

st.markdown("""
<style>
    /* Tema dark - base azul #004881 + acentos verdes #00983A / #8EC89A */
    .stApp {background-color:#0e1b2a;}                 /* fundo principal escuro */
    section[data-testid="stSidebar"] {background-color:#0a2540; border-right:1px solid #1d3a57;}
    .main .block-container {padding-top: 1.4rem; max-width: 1400px;}

    h1, h2, h3 {color:#8EC89A;}                        /* titulos em verde claro */
    .stApp p, .stApp span, .stApp label {color:#e6edf3;}
    [data-testid="stCaptionContainer"] p {color:#b0c4d8 !important; font-size:0.95rem;}

    div[data-testid="stMetric"] {
        background:#15273b; border:1px solid #25415c; border-radius:12px;
        padding:14px 16px; box-shadow:0 1px 4px rgba(0,0,0,.35);}
    div[data-testid="stMetricValue"] {font-size:1.7rem; color:#8EC89A;}

    /* Area de upload: card escuro, borda tracejada neutra harmonizada */
    section[data-testid="stFileUploaderDropzone"],
    div[data-testid="stFileUploader"] section {
        background:#15273b; border:1px dashed #3f5f7e; border-radius:12px;}

    .crit-card {border-left:6px solid; border-radius:10px; padding:14px 18px;
        margin-bottom:12px; background:#15273b; border:1px solid #25415c; color:#e6edf3;}
    .badge {display:inline-block; padding:2px 12px; border-radius:20px;
        font-size:.72rem; font-weight:700; color:#fff;}
    .stTabs [data-baseweb="tab-list"] {gap:4px;}
    .stTabs [aria-selected="true"] {color:#8EC89A;}
</style>
""", unsafe_allow_html=True)

FAIXAS = ["Baixa", "Media", "Alta", "Critica"]
CORES = {"Baixa": "#2e9e5b", "Media": "#e0a800", "Alta": "#e8590c", "Critica": "#c92a2a"}
# Cores de acento da marca (usadas nos graficos nativos do Streamlit)
VERDE = "#00983A"
VERDE_CLARO = "#8EC89A"

# ----------------------------------------------------------------------------
# Mapa de colunas (canonico -> variacoes possiveis no arquivo)
# ----------------------------------------------------------------------------
COLUNAS = {
    "numero": ["numero", "n", "id"],
    "data": ["data"],
    "unidade": ["unidade"],
    "area": ["area"],
    "processo": ["processo"],
    "subprocesso": ["sub-processo", "subprocesso", "sub processo"],
    "tipo": ["tipo"],
    "publico": ["publico"],
    "parceiro": ["parceiro"],
    "empresa": ["empresa"],
    "desvio": ["desvio"],
    "crencas": ["crencas"],
    "atividade_obs": ["atividade observada"],
    "desc_atividade": ["descricao da atividade"],
    "desc_abordagem": ["descricao da abordagem"],
    "observador": ["observador"],
    "colaborador": ["colaborador"],
    "lancado_por": ["lancado por"],
    "lancado_em": ["lancado em"],
}
NOMES_PESSOAS = ["observador", "colaborador", "lancado_por"]
SEP_MULTI = r",\s*<br\s*/?>|<br\s*/?>|\s*;\s*"


# ----------------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------------
def norm(s) -> str:
    s = str(s).strip().lower()
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def split_multi(cell):
    if pd.isna(cell):
        return []
    return [p.strip() for p in re.split(SEP_MULTI, str(cell)) if p and p.strip()]


def mapear_colunas(df):
    achado = {}
    norm_cols = {norm(c): c for c in df.columns}
    for canon, variacoes in COLUNAS.items():
        for v in variacoes:
            if norm(v) in norm_cols:
                achado[canon] = norm_cols[norm(v)]
                break
    return achado


# ----------------------------------------------------------------------------
# Leitura do arquivo
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def carregar(arquivo_bytes, nome):
    nome_l = nome.lower()
    if nome_l.endswith((".xlsx", ".xlsm")):
        df = pd.read_excel(io.BytesIO(arquivo_bytes))
    elif nome_l.endswith(".xls"):
        df = pd.read_excel(io.BytesIO(arquivo_bytes), engine="xlrd")
    else:
        df = None
        for enc in ("utf-8-sig", "latin-1", "cp1252"):
            for sep in (";", ",", "\t"):
                try:
                    tmp = pd.read_csv(io.BytesIO(arquivo_bytes), sep=sep, encoding=enc)
                    if tmp.shape[1] >= 3:
                        df = tmp
                        break
                except Exception:
                    continue
            if df is not None:
                break
        if df is None:
            raise ValueError("Nao foi possivel ler o CSV (encoding/separador).")
    df = df.dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


# ----------------------------------------------------------------------------
# Motor de criticidade
# ----------------------------------------------------------------------------
# Atividades criticas (peso 5): trabalho em altura, bloqueio de energias
# perigosas, equipamentos moveis, combate a incendios e icamentos.
ATIV_PESO = {
    "espaco confinado": 5, "trabalho em altura": 5, "bloqueio de energias perigosas": 5,
    "trabalho com eletricidade": 5, "equipamentos moveis": 5,
    "combate a incendios": 5, "combate a incendio": 5,
    "icamento de cargas": 5, "icamentos": 5, "icamento": 5,
    "trabalho a quente": 4, "escavacao/demolicao": 4,
    "manipulacao produtos quimicos": 4,
    "conducao veiculos": 2, "conducao de maquina ou equipamento": 2,
    "manutencao": 1, "atividades manuais": 1, "servicos de limpeza": 1,
    "atividade administrativa": 0,
}
DESVIO_CRIT = ["sem atc", "sem ate", "linha de tiro", "area de risco", "sem estar autorizado",
               "cinto trabalho em altura", "energias perigosas", "espaco confinado"]
DESVIO_ALTO = ["nao cumpre o procedimento", "nao adota as medidas", "macacao quimico",
               "protetor facial", "respirador", "ferramenta improvisada"]
KW_CRIT = ["queda", "altura", "energizad", "linha de tiro", "espaco confinado", "sem bloqueio",
           "choque", "vapor", "pressao", "explos", "incendio", "soterr", "amputac", "esmagam",
           "asfixia", "toxic", "inflamavel", " gas"]

P_ATIV, P_DESVIO, P_KW, P_INSEG = 5, 4, 3, 3
MAX_BRUTO = P_ATIV * 5 + P_DESVIO * 5 + P_KW * 3 + P_INSEG * 1  # 50


def score_atividade(cell):
    parts = [norm(p) for p in split_multi(cell)]
    return max([ATIV_PESO.get(p, 1) for p in parts] or [1])


def score_desvio(cell):
    t = norm(cell)
    if not t or t == "nan":
        return 0
    if any(k in t for k in DESVIO_CRIT):
        return 5
    if any(k in t for k in DESVIO_ALTO):
        return 3
    if "epi" in t:
        return 2
    if "suja ou desordenada" in t or "padroes" in t:
        return 1
    return 2


def score_texto(a, b):
    t = norm(a) + " " + norm(b)
    n = sum(1 for k in KW_CRIT if k in t)
    interromp = ("interromp" in t) or ("paralis" in t)
    return min(n, 3), interromp


def faixa_de(score):
    if score >= 80:
        return "Critica"
    if score >= 60:
        return "Alta"
    if score >= 40:
        return "Media"
    return "Baixa"


def processar(df_raw, cols):
    df = df_raw.copy()

    def get(c):
        return df[cols[c]] if c in cols else pd.Series([np.nan] * len(df))

    out = pd.DataFrame(index=df.index)
    out["Numero"] = get("numero")
    out["Area"] = get("area").fillna("Nao informado")
    out["Processo"] = get("processo").fillna("Nao informado")
    out["Subprocesso"] = get("subprocesso").fillna("Nao informado")
    out["Tipo"] = get("tipo").fillna("Nao informado")
    out["Publico"] = get("publico").fillna("Nao informado")
    out["Atividade"] = get("atividade_obs").fillna("Nao informado")
    out["Desvio"] = get("desvio")
    out["Crencas"] = get("crencas")
    out["DescAtividade"] = get("desc_atividade").fillna("")
    out["DescAbordagem"] = get("desc_abordagem").fillna("")

    # Data
    out["Data"] = pd.to_datetime(get("data"), dayfirst=True, errors="coerce")

    # Anonimizacao de empresa/parceiro -> codigos
    emp = get("empresa").fillna(get("parceiro"))
    codigos = {nome: f"Empresa {i+1:02d}"
               for i, nome in enumerate(sorted(emp.dropna().unique()))}
    out["EmpresaCod"] = emp.map(codigos).fillna("Quadro proprio")

    # Scores
    out["s_ativ"] = out["Atividade"].apply(score_atividade)
    out["s_desv"] = out["Desvio"].apply(score_desvio)
    tx = [score_texto(a, b) for a, b in zip(out["DescAtividade"], out["DescAbordagem"])]
    out["s_kw"] = [x[0] for x in tx]
    out["Interrompida"] = [x[1] for x in tx]
    out["s_inseg"] = (out["Tipo"].apply(norm) == "comportamento inseguro").astype(int)

    bruto = (out["s_ativ"] * P_ATIV + out["s_desv"] * P_DESVIO
             + out["s_kw"] * P_KW + out["s_inseg"] * P_INSEG)
    out["Criticidade"] = (bruto.clip(upper=MAX_BRUTO) / MAX_BRUTO * 100).round(1)
    out["Faixa"] = out["Criticidade"].apply(faixa_de)
    out["Faixa"] = pd.Categorical(out["Faixa"], categories=FAIXAS, ordered=True)
    return out


def explodir(series):
    return series.apply(split_multi).explode().str.strip()


def rankear(df, cat, val):
    """Prefixa a categoria com a posicao (01., 02., ...) ordenada por `val`
    decrescente. Como os graficos nativos ordenam o eixo alfabeticamente, isso
    faz o maior aparecer em primeiro (no topo, nas barras horizontais)."""
    d = df.sort_values(val, ascending=False).reset_index(drop=True)
    w = max(2, len(str(len(d))))
    d[cat] = [f"{i:0{w}d}. {v}" for i, v in enumerate(d[cat], 1)]
    return d


# ----------------------------------------------------------------------------
# Cabecalho
# ----------------------------------------------------------------------------
st.title("\U0001F9BA Analisador de Abordagens de Seguranca")
st.caption("Importe a planilha de observacoes comportamentais. O sistema calcula a "
           "criticidade de cada abordagem e a cruza com o local. Nenhum nome de "
           "pessoa e exibido.")

arquivo = st.file_uploader("Importar planilha (Excel ou CSV)",
                           type=["xlsx", "xls", "xlsm", "csv", "tsv"])

if arquivo is None:
    st.info("Aguardando o arquivo. A planilha deve conter colunas como Processo, "
            "Sub-processo, Tipo, Desvio, Atividade Observada e as descricoes da "
            "atividade e da abordagem.")
    st.stop()
    sys.exit(0)

try:
    df_raw = carregar(arquivo.getvalue(), arquivo.name)
except Exception as e:
    st.error(f"Falha ao ler o arquivo: {e}")
    st.stop()
    raise

cols = mapear_colunas(df_raw)
faltando = [c for c in ["processo", "atividade_obs", "desvio"] if c not in cols]
if faltando:
    st.warning("Colunas-chave nao encontradas: " + ", ".join(faltando) +
               ". O app segue, mas a analise pode ficar limitada.")

dados = processar(df_raw, cols)

# ----------------------------------------------------------------------------
# Filtros (sidebar)
# ----------------------------------------------------------------------------
st.sidebar.header("Filtros")
if dados["Data"].notna().any():
    dmin, dmax = dados["Data"].min(), dados["Data"].max()
    periodo = st.sidebar.date_input("Periodo", (dmin, dmax),
                                    min_value=dmin, max_value=dmax)
    if isinstance(periodo, (list, tuple)) and len(periodo) == 2:
        ini, fim = pd.to_datetime(periodo[0]), pd.to_datetime(periodo[1])
        dados = dados[(dados["Data"].isna()) |
                      ((dados["Data"] >= ini) & (dados["Data"] <= fim + pd.Timedelta(days=1)))]

proc_sel = st.sidebar.multiselect("Processo (local)", sorted(dados["Processo"].unique()))
if proc_sel:
    dados = dados[dados["Processo"].isin(proc_sel)]

ativ_opts = sorted(explodir(dados["Atividade"]).dropna().unique())
ativ_sel = st.sidebar.multiselect("Atividade observada", ativ_opts)
if ativ_sel:
    mask = dados["Atividade"].apply(lambda c: any(a in split_multi(c) for a in ativ_sel))
    dados = dados[mask]

faixa_sel = st.sidebar.multiselect("Faixa de criticidade", FAIXAS)
if faixa_sel:
    dados = dados[dados["Faixa"].isin(faixa_sel)]

st.sidebar.markdown("---")
st.sidebar.caption(f"{len(dados)} abordagens apos filtros")

if dados.empty:
    st.warning("Nenhum registro com os filtros atuais.")
    st.stop()

# ----------------------------------------------------------------------------
# Abas
# ----------------------------------------------------------------------------
aba1, aba2, aba3, aba4, aba5 = st.tabs(
    ["Visao Geral", "Analise de Criticidade", "Local x Criticidade",
     "Abordagens Criticas", "Areas de Sugestao de GEMBA"])

# ====================== ABA 1 - VISAO GERAL ======================
with aba1:
    total = len(dados)
    inseg = int((dados["Tipo"].apply(norm) == "comportamento inseguro").sum())
    crit_alta = int(dados["Faixa"].isin(["Alta", "Critica"]).sum())
    interr = int(dados["Interrompida"].sum())
    media = dados["Criticidade"].mean()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Abordagens", f"{total}")
    c2.metric("Comport. inseguro", f"{inseg}", f"{inseg/total*100:.0f}%")
    c3.metric("Alta + Critica", f"{crit_alta}", f"{crit_alta/total*100:.0f}%")
    c4.metric("Criticidade media", f"{media:.1f}")
    c5.metric("Atividades interrompidas", f"{interr}")

    st.markdown("###")
    g1, g2 = st.columns(2)

    with g1:
        st.subheader("Abordagens por processo (local)")
        proc = dados.groupby("Processo").size().reset_index(name="Qtd")
        proc = rankear(proc, "Processo", "Qtd")
        st.bar_chart(proc, x="Processo", y="Qtd", color=VERDE,
                     horizontal=True, height=380)

    with g2:
        st.subheader("Atividades observadas")
        a = explodir(dados["Atividade"]).value_counts().head(12).reset_index()
        a.columns = ["Atividade", "Qtd"]
        a = rankear(a, "Atividade", "Qtd")
        st.bar_chart(a, x="Atividade", y="Qtd", color=VERDE,
                     horizontal=True, height=380)

    g3, g4 = st.columns(2)
    with g3:
        st.subheader("Distribuicao por faixa de criticidade")
        f = dados["Faixa"].value_counts().reindex(FAIXAS).fillna(0).reset_index()
        f.columns = ["Faixa", "Qtd"]
        st.bar_chart(f, x="Faixa", y="Qtd", color=VERDE_CLARO, height=360)
    with g4:
        st.subheader("Familias de desvio")
        fam = (explodir(dados["Desvio"]).dropna()
               .apply(lambda x: x.split(">")[1].strip() if ">" in x else x))
        fam = fam.value_counts().head(10).reset_index()
        fam.columns = ["Familia", "Qtd"]
        fam = rankear(fam, "Familia", "Qtd")
        st.bar_chart(fam, x="Familia", y="Qtd", color=VERDE,
                     horizontal=True, height=360)

    if dados["Data"].notna().any():
        st.subheader("Evolucao temporal")
        serie = (dados.dropna(subset=["Data"])
                 .groupby(pd.Grouper(key="Data", freq="W"))
                 .agg(Qtd=("Numero", "size"), Criticidade=("Criticidade", "mean"))
                 .reset_index())
        st.line_chart(serie, x="Data", y="Criticidade", color=VERDE_CLARO,
                      height=300)

# ====================== ABA 2 - CRITICIDADE ======================
with aba2:
    st.subheader("Como a criticidade e calculada")
    st.markdown(
        "Cada abordagem recebe uma pontuacao de **0 a 100** combinando quatro fatores:\n"
        "- **Risco da atividade observada** (peso 5): trabalho em altura, espaco confinado, "
        "bloqueio de energias perigosas, eletricidade, equipamentos moveis, combate a "
        "incendios e icamentos pesam mais; atividades administrativas pesam menos.\n"
        "- **Gravidade do desvio** (peso 4): trabalho sem ATC/ATE, linha de tiro, area de risco e "
        "EPI de altura sao os mais graves.\n"
        "- **Palavras-chave criticas no texto** (peso 3): queda, energizado, vapor, pressao, "
        "explosao, choque, asfixia, entre outras.\n"
        "- **Comportamento inseguro** (peso 3)."
    )

    st.markdown("###")
    k1, k2 = st.columns([1.3, 1])
    with k1:
        st.subheader("Ranking de atividades mais criticas")
        st.caption("Atividades com 2+ observacoes. Passe o mouse para ver "
                   "criticidade media e n de abordagens.")
        rk = (dados.assign(A=dados["Atividade"].apply(split_multi)).explode("A")
              .dropna(subset=["A"])
              .groupby("A").agg(Criticidade=("Criticidade", "mean"),
                                Qtd=("Numero", "size")).reset_index())
        rk = rk[rk["Qtd"] >= 2]
        rk["Criticidade"] = rk["Criticidade"].round(1)
        rk = rk.rename(columns={"A": "Atividade", "Qtd": "n"})
        rk = rankear(rk, "Atividade", "Criticidade")
        st.bar_chart(rk, x="Atividade", y="Criticidade", color=VERDE,
                     horizontal=True, height=460)
    with k2:
        st.subheader("Criticidade media por publico")
        pub = (dados.groupby("Publico").agg(Criticidade=("Criticidade", "mean"),
                                            n=("Numero", "size")).reset_index())
        pub["Criticidade"] = pub["Criticidade"].round(1)
        pub = rankear(pub, "Publico", "Criticidade")
        st.bar_chart(pub, x="Publico", y="Criticidade", color=VERDE_CLARO,
                     horizontal=True, height=240)

        st.subheader("Empresas com maior criticidade")
        emp = (dados[dados["EmpresaCod"] != "Quadro proprio"]
               .groupby("EmpresaCod").agg(Criticidade=("Criticidade", "mean"),
                                          n=("Numero", "size")).reset_index())
        emp = emp[emp["n"] >= 2].sort_values("Criticidade", ascending=False).head(8)
        if not emp.empty:
            emp["Criticidade"] = emp["Criticidade"].round(1)
            emp = rankear(emp, "EmpresaCod", "Criticidade")
            st.bar_chart(emp, x="EmpresaCod", y="Criticidade", color=VERDE,
                         horizontal=True, height=240)
        else:
            st.caption("Sem empresas parceiras com 2+ registros no filtro atual.")

# ====================== ABA 3 - LOCAL x CRITICIDADE ======================
with aba3:
    st.subheader("Cruzamento: local da abordagem x criticidade")
    nivel = st.radio("Nivel de local", ["Processo", "Subprocesso", "Area"],
                     horizontal=True)

    # Mapa de calor (tons de verde): quantidade de abordagens por local x faixa
    st.caption("Mapa de calor: numero de abordagens por local e faixa, em tons de verde.")
    pivot = (dados.pivot_table(index=nivel, columns="Faixa", values="Numero",
                               aggfunc="count", fill_value=0, observed=False)
             .reindex(columns=FAIXAS, fill_value=0))
    pivot["_ord"] = dados.groupby(nivel)["Criticidade"].mean()
    pivot = pivot.sort_values("_ord", ascending=False).drop(columns="_ord")
    pivot = pivot[pivot.sum(axis=1) > 0]

    vmax = int(pivot.values.max()) if pivot.size else 0

    def cor_verde(v):
        # intensidade proporcional ao valor, sobre o card escuro (tons de verde)
        a = 0 if vmax == 0 else v / vmax
        return (f"background-color: rgba(0,152,58,{0.12 + 0.78 * a:.2f}); "
                f"color: {'#0e1b2a' if a > 0.55 else '#e6edf3'}")

    sty = pivot.style.applymap(cor_verde).format("{:d}")
    st.dataframe(sty, use_container_width=True)

    st.subheader("Matriz de prioridade: volume x criticidade")
    g = (dados.groupby(nivel).agg(Qtd=("Numero", "size"),
                                  Criticidade=("Criticidade", "mean"),
                                  Criticas=("Faixa", lambda s: s.isin(["Alta", "Critica"]).sum()))
         .reset_index())
    g = g[g["Qtd"] > 0]
    g["Criticidade"] = g["Criticidade"].round(1)
    st.scatter_chart(g, x="Qtd", y="Criticidade", size="Criticas", color=VERDE_CLARO,
                     x_label="Volume de abordagens", y_label="Criticidade media",
                     height=460)
    st.caption("Quadrante superior direito = locais que combinam alto volume e alta "
               "criticidade: prioridade maxima de atuacao. O tamanho do ponto indica "
               "quantas abordagens sao de faixa Alta/Critica. Criticidade media geral: "
               f"{dados['Criticidade'].mean():.1f}.")

    st.subheader("Tabela detalhada por local")
    tab = g.sort_values("Criticidade", ascending=False).rename(
        columns={nivel: "Local", "Qtd": "Abordagens",
                 "Criticidade": "Criticidade media", "Criticas": "Alta/Critica"})
    tab["Criticidade media"] = tab["Criticidade media"].round(1)
    st.dataframe(tab, use_container_width=True, hide_index=True)

# ====================== ABA 4 - ABORDAGENS CRITICAS ======================
with aba4:
    st.subheader("Abordagens mais criticas (texto integral, sem nomes)")
    top_n = st.slider("Quantas exibir", 5, 50, 15)
    crit = dados.sort_values("Criticidade", ascending=False).head(top_n)

    def limpar(t):
        return re.sub(r"<br\s*/?>", " ", str(t)).strip()

    for _, r in crit.iterrows():
        cor = CORES[str(r["Faixa"])]
        desvio_txt = limpar(r["Desvio"]) if pd.notna(r["Desvio"]) else "-"
        local = f"{r['Processo']} | {r['Subprocesso']}"
        st.markdown(f"""
<div class="crit-card" style="border-left-color:{cor}">
  <span class="badge" style="background:{cor}">{r['Faixa']} - {r['Criticidade']:.0f}</span>
  &nbsp;<b>{limpar(r['Atividade'])}</b>
  &nbsp;<span style="color:#9fb0bf">| Local: {local} | {r['Publico']} | {r['EmpresaCod']}</span>
  <div style="margin-top:8px"><b>Desvio:</b> {desvio_txt}</div>
  <div style="margin-top:4px"><b>Atividade:</b> {limpar(r['DescAtividade'])}</div>
  <div style="margin-top:4px"><b>Abordagem:</b> {limpar(r['DescAbordagem'])}</div>
</div>
""", unsafe_allow_html=True)

    # Exportacao (anonimizada)
    exp = crit[["Numero", "Area", "Processo", "Subprocesso", "Publico", "EmpresaCod",
                "Atividade", "Desvio", "Criticidade", "Faixa",
                "DescAtividade", "DescAbordagem"]].copy()
    csv = exp.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button("Baixar abordagens criticas (CSV, sem nomes)", csv,
                       file_name="abordagens_criticas.csv", mime="text/csv")

# ============= ABA 5 - AREAS DE SUGESTAO DE GEMBA =============
with aba5:
    st.subheader("Areas sugeridas para GEMBA (ida ao local)")
    st.caption("Esta aba prioriza cada combinacao Processo + Sub-processo que reune maior "
               "criticidade media, maior concentracao de abordagens Alta/Critica e "
               "mais comportamentos inseguros - onde uma verificacao presencial tende "
               "a trazer mais retorno.")

    cmin, ctop = st.columns(2)
    min_obs = cmin.slider("Minimo de abordagens", 1, 10, 2, key="gemba_min")
    top_n = ctop.slider("Locais a destacar", 3, 20, 5, key="gemba_top")

    base = dados.copy()
    base["_critico"] = base["Faixa"].isin(["Alta", "Critica"])
    base["_inseg"] = base["Tipo"].apply(norm) == "comportamento inseguro"

    g = (base.groupby(["Processo", "Subprocesso"]).agg(
            Abordagens=("Numero", "size"),
            Criticidade=("Criticidade", "mean"),
            Criticas=("_critico", "sum"),
            Inseguros=("_inseg", "sum"),
            Interrompidas=("Interrompida", "sum"),
         ).reset_index())
    g["Local"] = g["Processo"] + " | " + g["Subprocesso"]
    g = g[(g["Abordagens"] >= min_obs) & (g["Processo"] != "Nao informado")]

    if g.empty:
        st.info("Nenhum local atinge o minimo de abordagens selecionado.")
    else:
        g["%Critica"] = (g["Criticas"] / g["Abordagens"] * 100)
        g["%Inseguro"] = (g["Inseguros"] / g["Abordagens"] * 100)
        # Indice GEMBA (0-100): criticidade media + concentracao de
        # abordagens Alta/Critica + concentracao de comportamento inseguro.
        g["IndiceGEMBA"] = (0.5 * g["Criticidade"]
                            + 0.35 * g["%Critica"]
                            + 0.15 * g["%Inseguro"]).round(1)
        g = g.sort_values("IndiceGEMBA", ascending=False)
        destaque = g.head(top_n)

        st.subheader("Ranking de prioridade GEMBA (Indice 0-100)")
        chart = (alt.Chart(destaque).mark_bar(color=VERDE, cornerRadiusEnd=3)
                 .encode(
                     x=alt.X("IndiceGEMBA:Q",
                             scale=alt.Scale(domain=[0, 100]),
                             axis=alt.Axis(values=list(range(0, 101, 10)),
                                           title="Indice GEMBA (0-100)")),
                     y=alt.Y("Local:N", sort="-x", title=None,
                             scale=alt.Scale(paddingInner=0.35, paddingOuter=0.25),
                             axis=alt.Axis(labelLimit=360, labelPadding=10,
                                           labelFontSize=12, labelLineHeight=14)),
                     tooltip=["Local", "IndiceGEMBA", "Abordagens",
                              "Criticidade", "Criticas"])
                 .properties(height=max(340, 60 * len(destaque))))
        st.altair_chart(chart, use_container_width=True)

        st.subheader("Por que estes locais foram indicados")
        for _, r in destaque.iterrows():
            local = r["Local"]
            cor = CORES[faixa_de(r["IndiceGEMBA"])]
            sub = base[(base["Processo"] == r["Processo"])
                       & (base["Subprocesso"] == r["Subprocesso"])]

            ativ = explodir(sub["Atividade"]).dropna().value_counts().head(3)
            ativ_txt = ", ".join(f"{a} ({q})" for a, q in ativ.items()) or "-"
            fam = (explodir(sub["Desvio"]).dropna()
                   .apply(lambda x: x.split(">")[1].strip() if ">" in x else x))
            fam = fam.value_counts().head(3)
            fam_txt = ", ".join(f"{a} ({q})" for a, q in fam.items()) or "-"

            st.markdown(f"""
<div class="crit-card" style="border-left-color:{cor}">
  <span class="badge" style="background:{cor}">GEMBA {r['IndiceGEMBA']:.0f}</span>
  &nbsp;<b>{local}</b>
  &nbsp;<span style="color:#9fb0bf">| {int(r['Abordagens'])} abordagens |
  criticidade media {r['Criticidade']:.0f} |
  {int(r['Criticas'])} Alta/Critica ({r['%Critica']:.0f}%) |
  {int(r['Inseguros'])} comport. inseguro |
  {int(r['Interrompidas'])} interrompidas</span>
  <div style="margin-top:8px"><b>Atividades mais observadas:</b> {ativ_txt}</div>
  <div style="margin-top:4px"><b>Principais desvios:</b> {fam_txt}</div>
</div>
""", unsafe_allow_html=True)

        st.subheader("Tabela de priorizacao")
        tab = g.rename(columns={"Criticidade": "Criticidade media",
                                "Criticas": "Alta/Critica",
                                "Inseguros": "Comport. inseguro"})
        tab["Criticidade media"] = tab["Criticidade media"].round(1)
        tab["%Critica"] = tab["%Critica"].round(0)
        tab["%Inseguro"] = tab["%Inseguro"].round(0)
        tab = tab[["Local", "IndiceGEMBA", "Abordagens", "Criticidade media",
                   "Alta/Critica", "%Critica", "Comport. inseguro", "%Inseguro",
                   "Interrompidas"]]
        st.dataframe(tab, use_container_width=True, hide_index=True)

        csv_g = tab.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("Baixar priorizacao GEMBA (CSV)", csv_g,
                           file_name="areas_sugestao_gemba.csv", mime="text/csv")
