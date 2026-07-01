# -*- coding: utf-8 -*-
"""
Espelho de Simulacao - Usina
----------------------------
Importa um arquivo XML de simulacao (raiz <Simulation> com secoes Correntes /
Operacoes / Sequencia) e extrai indicadores operacionais reproduzindo a logica
da planilha "Modelo_Espelho":

  - Aba "Espelho SimDiaria": extracao robusta de ~70 indicadores. Cada indicador
    e um "PROCV por nome de corrente" - dado o nome de uma corrente, le um
    atributo (W, WVol, T, POL, AR, pot ...) do elemento <Corrente>, ou um
    atributo do <Componente> (brix, pureza, fracaoMassica), aplicando a
    transformacao de unidade correspondente.

  - Aba "Espelho Visual": os mesmos numeros apresentados sobre um fluxograma de
    processo (Graphviz), da cana ao produto e a energia.

Como executar:
    pip install -r requirements.txt
    streamlit run app.py
"""

import math
import xml.etree.ElementTree as ET

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------------
# Pagina + tema (mesma identidade visual dos demais apps do repo)
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Espelho de Simulacao",
                   page_icon="🪞", layout="wide")

st.markdown("""
<style>
    .stApp {background-color:#0e1b2a;}
    section[data-testid="stSidebar"] {background-color:#0a2540; border-right:1px solid #1d3a57;}
    .main .block-container {padding-top: 1.4rem; max-width: 1500px;}
    h1, h2, h3 {color:#8EC89A;}
    div[data-testid="stMetric"] {
        background:#15273b; border:1px solid #25415c; border-radius:12px;
        padding:14px 16px; box-shadow:0 1px 4px rgba(0,0,0,.35);}
    div[data-testid="stMetricValue"] {font-size:1.6rem; color:#8EC89A;}
    section[data-testid="stFileUploaderDropzone"],
    div[data-testid="stFileUploader"] section {
        background:#15273b; border:1px dashed #3f5f7e; border-radius:12px;}
    .stTabs [data-baseweb="tab-list"] {gap:4px;}
    .stTabs [aria-selected="true"] {color:#8EC89A;}
    .sec-title {color:#8EC89A; font-weight:700; font-size:1.02rem;
        margin:18px 0 4px 0; border-bottom:1px solid #25415c; padding-bottom:4px;}
</style>
""", unsafe_allow_html=True)

VERDE = "#00983A"
VERDE_CLARO = "#8EC89A"
AZUL = "#004881"

NaN = float("nan")


# ----------------------------------------------------------------------------
# Modelo: parse do XML e acesso por nome de corrente (replica o VLOOKUP)
# ----------------------------------------------------------------------------
def _to_float(v):
    if v is None:
        return NaN
    s = str(v).strip()
    if s == "":
        return NaN
    try:
        return float(s)
    except ValueError:
        # tolera decimal com virgula
        try:
            return float(s.replace(".", "").replace(",", "."))
        except ValueError:
            return NaN


class Sim:
    """Correntes indexadas por nome (nomes sao unicos no XML)."""

    def __init__(self, root):
        self.meta = dict(root.attrib)
        self.correntes = {}
        cs = root.find("Correntes")
        if cs is not None:
            for c in cs.findall("Corrente"):
                name = c.get("name")
                comps = [dict(cp.attrib) for cp in c.findall("Componente")]
                self.correntes[name] = {"attrs": dict(c.attrib), "comps": comps}

    # atributo de nivel-corrente
    def cor(self, name, attr):
        c = self.correntes.get(name)
        if not c:
            return NaN
        return _to_float(c["attrs"].get(attr))

    # atributo de nivel-componente: primeiro componente (padrao) ou nomeado
    def comp(self, name, attr, comp_name=None):
        c = self.correntes.get(name)
        if not c or not c["comps"]:
            return NaN
        if comp_name is None:
            return _to_float(c["comps"][0].get(attr))
        for cp in c["comps"]:
            if cp.get("name") == comp_name:
                return _to_float(cp.get(attr))
        return NaN

    # atalhos legiveis
    def W(self, n):     return self.cor(n, "W")
    def WVol(self, n):  return self.cor(n, "WVol")
    def T(self, n):     return self.cor(n, "T")
    def POL(self, n):   return self.cor(n, "POL")
    def AR(self, n):    return self.cor(n, "AR")
    def pot(self, n):   return self.cor(n, "pot")
    def brix(self, n):          return self.comp(n, "brix")
    def pureza(self, n, c=None):return self.comp(n, "pureza", c)
    def frac(self, n, c=None):  return self.comp(n, "fracaoMassica", c)


def _pos(x):
    """IF(x<0;0;x) - usado nas potencias das turbinas."""
    return 0.0 if (x == x and x < 0) else x


# ----------------------------------------------------------------------------
# Definicao dos indicadores (aba Espelho SimDiaria)
# Cada item: (secao, rotulo, funcao(sim)->valor). NaN => exibido como "-".
# Fielmente derivado das formulas da planilha.
# ----------------------------------------------------------------------------
def build_indicators(s: Sim):
    # parametros auxiliares (linhas 86-100 da planilha)
    emb_total = s.WVol("w_1_1") + s.WVol("w_1_2") + s.WVol("w_1_3")
    consumidores_utl = (s.W("ve_3") + s.W("ve_9_7") + s.W("vr22_3") +
                        s.W("vr13_10_1") + s.W("vr6_10_8") + s.W("i_13") +
                        s.W("ve_11_1"))
    vapor_processo = s.W("i_9_8") - consumidores_utl
    moagem = s.W("1_1") * 24.0
    consumo_esp = (vapor_processo * 1000.0 * 24.0 / moagem) if moagem else NaN

    tg5 = _pos(s.pot("e_10_22"))
    tg6 = _pos(s.pot("e_10_21"))
    tg7 = _pos(s.pot("e_19"))
    tgs_14 = _pos(s.pot("e_10_13"))
    pot_consumida = s.pot("e_10_17")
    exp_sm = tg5 + tgs_14 - pot_consumida
    exp_sme = tg6
    exp_smbio = s.pot("e_19")
    exp_total = exp_sm + exp_sme + exp_smbio

    etanol_anidro = s.WVol("ani_8_8") * 24.0 + s.WVol("20") * 24.0
    etanol_h2 = s.WVol("h2_8_4") * 24.0
    etanol_carb = s.WVol("hid_8_10") * 24.0

    rows = [
        # -- Qualidade da cana --
        ("CANA / QUALIDADE", "Moagem [ton/dia]",                 moagem),
        ("CANA / QUALIDADE", "Fibra [%]",                        s.frac("1_1", "Fibra") * 100.0),
        ("CANA / QUALIDADE", "Impureza Mineral [kg/tc]",         s.frac("1_1", "Terra") * 1000.0),
        ("CANA / QUALIDADE", "Pol [%]",                          s.POL("1_1")),
        ("CANA / QUALIDADE", "AR [%]",                           s.AR("1_1")),
        ("CANA / QUALIDADE", "Pureza do caldo [%]",              s.pureza("1_1", "Caldo") * 100.0),

        # -- Extracao --
        ("EXTRACAO", "Vazao de agua de embebicao total [m3/h]",  emb_total),
        ("EXTRACAO", "Temperatura de agua de embebicao moendas [C]", s.T("i_1_1") - 273.0),

        # -- Tratamento de caldo --
        ("TRATAMENTO DE CALDO", "Vazao agua embebicao filtros de lodo [m3/h]", s.WVol("w_6_3")),
        ("TRATAMENTO DE CALDO", "Vazao caldo misto fabrica [m3/h]",            s.WVol("1_34")),
        ("TRATAMENTO DE CALDO", "Temperatura do caldo (Entrada TRC) [C]",      s.T("1_34") - 273.0),
        ("TRATAMENTO DE CALDO", "Vazao caldo misto destilaria [m3/h]",         s.WVol("1_29")),
        ("TRATAMENTO DE CALDO", "Vazao de filtrado [m3/h]",                    s.WVol("6_48")),
        ("TRATAMENTO DE CALDO", "Vazao de clarificado fabrica - Thermol [m3/h]", s.WVol("2_39")),
        ("TRATAMENTO DE CALDO", "Vazao de clarificado fabrica - ENET [m3/h]",  s.WVol("2_41")),
        ("TRATAMENTO DE CALDO", "Temp. caldo misto destilaria - VV3 [C]",      s.T("6_3") - 273.0),
        ("TRATAMENTO DE CALDO", "Delta T Caldo Misto Regenerador Placas - CxC [C]", s.T("2_13") - s.T("1_34")),
        ("TRATAMENTO DE CALDO", "Temp. Saida Aquecedores Fabrica [C]",         s.T("2_23") - 273.0),
        ("TRATAMENTO DE CALDO", "Temp. Saida Aquecedores Destilaria [C]",      s.T("6_25") - 273.0),

        # -- Concentracao de caldo --
        ("CONCENTRACAO DE CALDO", "Temp. Saida Caldo Clarif. Thermol (CxV) ROBERT [C]", s.T("2_39") - 273.0),
        ("CONCENTRACAO DE CALDO", "Temp. Saida Caldo Clarif. Regenerador (CxV) ENET [C]", s.T("2_41") - 273.0),
        ("CONCENTRACAO DE CALDO", "Brix clarificado fabrica [%]",             s.brix("2_42") * 100.0),
        ("CONCENTRACAO DE CALDO", "Brix clarificado destilaria [%]",          s.brix("2_6") * 100.0),
        ("CONCENTRACAO DE CALDO", "Brix caldo filtrado [%]",                  s.brix("6_48") * 100.0),
        ("CONCENTRACAO DE CALDO", "Vazao de caldo pre evaporado [m3/h]",      s.WVol("3_6")),
        ("CONCENTRACAO DE CALDO", "Vazao xarope fabrica [m3/h]",              s.WVol("3_21")),
        ("CONCENTRACAO DE CALDO", "Brix Xarope [%]",                          s.brix("3_21") * 100.0),
        ("CONCENTRACAO DE CALDO", "Pureza Xarope [%]",                        s.pureza("3_21") * 100.0),

        # -- Fabrica de acucar --
        ("FABRICA DE ACUCAR", "Recuperacao Fabrica [%]",                      s.W("rec_fabrica") / 100.0),
        ("FABRICA DE ACUCAR", "Pureza mel A [%]",                             s.pureza("4_26") * 100.0),
        ("FABRICA DE ACUCAR", "Pureza mel B kont [%]",                        s.pureza("5_56") * 100.0),
        ("FABRICA DE ACUCAR", "Producao Acucar [sc/d]",                       s.W("sc_por_dia")),
        ("FABRICA DE ACUCAR", "Relacao producao de acucar [sc/tc]",           s.W("sc_por_tonelada")),

        # -- Producao de etanol --
        ("PRODUCAO DE ETANOL", "Vazao caldo clarificado - Destilaria [m3/h]", s.WVol("2_6")),
        ("PRODUCAO DE ETANOL", "Vazao de mel A p/ fermentacao [m3/h]",        s.WVol("4_26")),
        ("PRODUCAO DE ETANOL", "Vazao de mel B p/ fermentacao [m3/h]",        s.WVol("5_47")),
        ("PRODUCAO DE ETANOL", "Vazao de mel p/ composicao mosto [m3/h]",     s.WVol("5_56")),
        ("PRODUCAO DE ETANOL", "Vazao de agua p/ composicao do mosto [m3/h]", s.WVol("w_7_2")),
        ("PRODUCAO DE ETANOL", "Brix do Mosto [%]",                           s.brix("7_10") * 100.0),
        ("PRODUCAO DE ETANOL", "Vazao de mosto [m3/h]",                       s.WVol("7_10")),
        ("PRODUCAO DE ETANOL", "Variacao Estoque de Mel [m3/dia]",            abs(s.WVol("i_5_100") * 24.0)),
        ("PRODUCAO DE ETANOL", "Etanol Anidro ANP [m3/d]",                    etanol_anidro),
        ("PRODUCAO DE ETANOL", "Etanol Hidratado H2 [m3/d]",                  etanol_h2),
        ("PRODUCAO DE ETANOL", "Hidratado para Redestilo [m3/d]",             s.WVol("hid_3") * 24.0),
        ("PRODUCAO DE ETANOL", "Etanol Hidratado Carburante [m3/d]",          etanol_carb),
        ("PRODUCAO DE ETANOL", "Vazao Xarope para fermentacao [m3/h]",        s.WVol("3_18")),

        # -- Utilidades / energia --
        ("UTILIDADES", "Consumo especifico (s/ consumidores UTL) [kgv/tc]",   consumo_esp),
        ("UTILIDADES", "Vapor Processo (s/ consumo UTL) [t/h]",               vapor_processo),
        ("UTILIDADES", "PRODUCAO CBC-1 [t/h]",                                s.W("vd22_10")),
        ("UTILIDADES", "PRODUCAO CBC-2 [t/h]",                                s.W("vd22_12")),
        ("UTILIDADES", "PRODUCAO CBC-3 [t/h]",                                s.W("vd22_8")),
        ("UTILIDADES", "PRODUCAO CBC-4 [t/h]",                                s.W("vd22_9_2")),
        ("UTILIDADES", "Producao Caldeira 9 [t/h]",                           s.W("vd67_9_4")),
        ("UTILIDADES", "Producao Caldeira 10 [t/h]",                          s.W("vd100_9_2")),
        ("UTILIDADES", "Vazao condicionadora 67x55 kgf/cm2 [t/h]",            s.W("vr55_10_2")),
        ("UTILIDADES", "Vazao condicionadora 67x22 kgf/cm2 [t/h]",            s.W("vr22_10_1")),
        ("UTILIDADES", "Potencia TG7 [MWh]",                                  tg7),
        ("UTILIDADES", "Potencia TG6 [MWh]",                                  tg6),
        ("UTILIDADES", "Potencia TG5 [MWh]",                                  tg5),
        ("UTILIDADES", "Potencia TGs 1-4 [MWh]",                              tgs_14),
        ("UTILIDADES", "Potencia consumida [MWh/dia]",                        pot_consumida),
        ("UTILIDADES", "Exportacao SM [MW]",                                  exp_sm),
        ("UTILIDADES", "Exportacao SME [MW]",                                 exp_sme),
        ("UTILIDADES", "Exportacao SMBio [MW]",                               exp_smbio),
        ("UTILIDADES", "Exportacao TOTAL [MW]",                               exp_total),
        ("UTILIDADES", "Sobra de bagaco [t/dia]",                             s.W("9_3") * 24.0),
        ("UTILIDADES", "Consumo bagaco [t/dia]",                              s.W("9_5") * 24.0),
    ]
    return pd.DataFrame(rows, columns=["Secao", "Indicador", "Simulado"])


def ind_map(df):
    """Dicionario Indicador->valor para reuso na aba Visual."""
    return dict(zip(df["Indicador"], df["Simulado"]))


# ----------------------------------------------------------------------------
# Formatacao
# ----------------------------------------------------------------------------
def fmt(v, dec=2):
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "-"
    return f"{v:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ----------------------------------------------------------------------------
# Aba Visual - fluxograma Graphviz
# ----------------------------------------------------------------------------
def build_flow_dot(m):
    def g(label, dec=1):
        return fmt(m.get(label), dec)

    def node(nid, titulo, linhas, fill="#15273b"):
        body = "".join(
            f'<TR><TD ALIGN="LEFT"><FONT COLOR="#c9d6e3" POINT-SIZE="10">{k}</FONT></TD>'
            f'<TD ALIGN="RIGHT"><FONT COLOR="#8EC89A" POINT-SIZE="11"><B>{v}</B></FONT></TD></TR>'
            for k, v in linhas
        )
        lbl = (
            f'<<TABLE BORDER="0" CELLBORDER="0" CELLSPACING="2" CELLPADDING="2">'
            f'<TR><TD COLSPAN="2"><FONT COLOR="#8EC89A" POINT-SIZE="13"><B>{titulo}</B></FONT></TD></TR>'
            f'{body}</TABLE>>'
        )
        return f'  "{nid}" [label={lbl}, fillcolor="{fill}"];\n'

    dot = [
        'digraph espelho {',
        '  rankdir=LR; bgcolor="transparent"; splines=ortho; nodesep=0.5; ranksep=0.7;',
        '  node [shape=box, style="filled,rounded", color="#25415c", fontname="Helvetica"];',
        '  edge [color="#3f5f7e", penwidth=1.4, arrowsize=0.8];',
    ]

    dot.append(node("cana", "CANA / QUALIDADE", [
        ("Moagem [tc/dia]", g("Moagem [ton/dia]", 0)),
        ("Fibra [%]", g("Fibra [%]")),
        ("Imp. Mineral [kg/tc]", g("Impureza Mineral [kg/tc]")),
        ("Pol [%]", g("Pol [%]")),
        ("AR [%]", g("AR [%]", 2)),
        ("Pureza [%]", g("Pureza do caldo [%]")),
    ], fill="#123a2a"))

    dot.append(node("extr", "EXTRACAO / MOENDAS", [
        ("Agua embebicao [m3/h]", g("Vazao de agua de embebicao total [m3/h]")),
        ("Temp. embebicao [C]", g("Temperatura de agua de embebicao moendas [C]")),
        ("Sobra de bagaco [t/dia]", g("Sobra de bagaco [t/dia]", 0)),
    ]))

    dot.append(node("trat", "TRATAMENTO DE CALDO", [
        ("Caldo misto fabrica [m3/h]", g("Vazao caldo misto fabrica [m3/h]")),
        ("Caldo misto destil. [m3/h]", g("Vazao caldo misto destilaria [m3/h]")),
        ("Clarificado Thermol [m3/h]", g("Vazao de clarificado fabrica - Thermol [m3/h]")),
        ("Clarificado ENET [m3/h]", g("Vazao de clarificado fabrica - ENET [m3/h]")),
        ("Filtrado [m3/h]", g("Vazao de filtrado [m3/h]")),
    ]))

    dot.append(node("conc", "CONCENTRACAO", [
        ("Caldo pre evap. [m3/h]", g("Vazao de caldo pre evaporado [m3/h]")),
        ("Xarope fabrica [m3/h]", g("Vazao xarope fabrica [m3/h]")),
        ("Brix Xarope [%]", g("Brix Xarope [%]")),
        ("Pureza Xarope [%]", g("Pureza Xarope [%]")),
    ]))

    dot.append(node("mosto", "COMPOSICAO DO MOSTO", [
        ("Caldo clar. destil. [m3/h]", g("Vazao caldo clarificado - Destilaria [m3/h]")),
        ("Mel p/ mosto [m3/h]", g("Vazao de mel p/ composicao mosto [m3/h]")),
        ("Agua p/ mosto [m3/h]", g("Vazao de agua p/ composicao do mosto [m3/h]")),
        ("Xarope p/ ferm. [m3/h]", g("Vazao Xarope para fermentacao [m3/h]")),
        ("Vazao mosto [m3/h]", g("Vazao de mosto [m3/h]")),
    ]))

    dot.append(node("acucar", "FABRICA DE ACUCAR", [
        ("Producao Acucar [sc/d]", g("Producao Acucar [sc/d]", 0)),
        ("Relacao [sc/tc]", g("Relacao producao de acucar [sc/tc]", 2)),
        ("Recuperacao [%]", g("Recuperacao Fabrica [%]", 2)),
        ("Pureza mel A [%]", g("Pureza mel A [%]")),
        ("Pureza mel B [%]", g("Pureza mel B kont [%]")),
    ], fill="#123a2a"))

    dot.append(node("etanol", "DESTILARIA / ETANOL", [
        ("Etanol Anidro ANP [m3/d]", g("Etanol Anidro ANP [m3/d]", 0)),
        ("Etanol Hidratado H2 [m3/d]", g("Etanol Hidratado H2 [m3/d]", 0)),
        ("Etanol Carburante [m3/d]", g("Etanol Hidratado Carburante [m3/d]", 0)),
        ("Mel A ferm. [m3/h]", g("Vazao de mel A p/ fermentacao [m3/h]")),
        ("Mel B ferm. [m3/h]", g("Vazao de mel B p/ fermentacao [m3/h]")),
    ], fill="#123a2a"))

    dot.append(node("energia", "UTILIDADES / ENERGIA", [
        ("Vapor Processo [t/h]", g("Vapor Processo (s/ consumo UTL) [t/h]")),
        ("Consumo esp. [kgv/tc]", g("Consumo especifico (s/ consumidores UTL) [kgv/tc]")),
        ("Exportacao SM [MW]", g("Exportacao SM [MW]", 2)),
        ("Exportacao SME [MW]", g("Exportacao SME [MW]", 2)),
        ("Exportacao SMBio [MW]", g("Exportacao SMBio [MW]", 2)),
        ("Exportacao TOTAL [MW]", g("Exportacao TOTAL [MW]", 2)),
    ], fill="#1a2f1a"))

    dot += [
        '  cana -> extr -> trat -> conc;',
        '  conc -> mosto; conc -> acucar;',
        '  mosto -> etanol;',
        '  acucar -> energia; etanol -> energia; extr -> energia [style=dashed];',
        '}',
    ]
    return "".join(l if l.endswith("\n") else l + "\n" for l in dot)


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
st.title("🪞 Espelho de Simulacao")
st.caption("Importe o XML da simulacao e visualize os indicadores diarios e o fluxo de processo.")

with st.sidebar:
    st.header("Arquivo")
    up = st.file_uploader("Arquivo de simulacao (.xml)", type=["xml"])
    st.markdown("---")
    st.caption("Os indicadores reproduzem a planilha Modelo_Espelho "
               "(abas Espelho_SimDiaria e Espelho_Visual).")

if up is None:
    st.info("⬅️ Envie o arquivo **.xml** da simulacao na barra lateral para comecar.")
    st.stop()

try:
    root = ET.fromstring(up.getvalue())
except ET.ParseError as e:
    st.error(f"Nao foi possivel ler o XML: {e}")
    st.stop()

if root.tag != "Simulation":
    st.warning(f"A raiz do XML e <{root.tag}>, esperado <Simulation>. "
               "Tentando processar mesmo assim.")

sim = Sim(root)
df = build_indicators(sim)
m = ind_map(df)

c1, c2, c3 = st.columns(3)
c1.metric("Cenario", sim.meta.get("diagram_name", "-"))
c2.metric("Versao", sim.meta.get("version", "-"))
c3.metric("Correntes", f"{len(sim.correntes):,}".replace(",", "."))

tab1, tab2 = st.tabs(["📋 Espelho SimDiaria", "🔗 Espelho Visual"])

with tab1:
    st.subheader("Avaliacao de aderencia a simulacao - Simulado")
    for sec in df["Secao"].unique():
        st.markdown(f'<div class="sec-title">{sec}</div>', unsafe_allow_html=True)
        sub = df[df["Secao"] == sec].copy()
        sub["Simulado"] = sub["Simulado"].map(lambda v: fmt(v, 4))
        st.dataframe(sub[["Indicador", "Simulado"]], hide_index=True,
                     use_container_width=True)

    csv = df.assign(Simulado=df["Simulado"]).to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Baixar CSV", csv,
                       file_name=f"espelho_simdiaria_{sim.meta.get('diagram_name','sim')}.csv",
                       mime="text/csv")

with tab2:
    st.subheader("Espelho Visual - Fluxo de processo")
    st.graphviz_chart(build_flow_dot(m), use_container_width=True)
