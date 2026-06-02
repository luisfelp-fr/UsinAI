# -*- coding: utf-8 -*-
"""
=====================================================================
 Analisador de Sensibilidade — Recuperação Industrial
=====================================================================
Aplicativo (Streamlit) que:
  1. Lê uma planilha Excel com os indicadores (mesma estrutura da aba "DB"
     do arquivo original: colunas DATA | DESC_PERIODO | DSC_METRICA | VALOR),
     contendo dados de VÁRIOS dias.
  2. Deixa o usuário escolher o PERÍODO e a DATA de análise.
  3. Reconstrói toda a cadeia de cálculo da Recuperação Industrial para
     aquela data e roda uma análise de sensibilidade.
  4. Mostra um RANKING (tema escuro) com o % de influência de cada
     indicador sobre a Recuperação Industrial.

COMO RODAR:
    pip install streamlit openpyxl pandas
    streamlit run app.py

O modelo de cálculo abaixo é uma tradução fiel da COLUNA L da aba "Simulador"
e reproduz exatamente o valor da planilha original.
"""
import io
import datetime
import pandas as pd
import streamlit as st

# =====================================================================
#  MODELO (tradução da coluna L da aba "Simulador")
# =====================================================================
ROW_NAME = {
    11: 'CANA TOTAL',
    12: 'PC',
    13: 'AR Cana (%)',
    14: 'Caldo da Cana - Pol',
    20: 'ART Disponível - Mel Comprado',
    21: 'ART Disponível - Xarope Comprado',
    22: 'ART Disponível - Creme de Levedura Comprado',
    36: 'Açúcar Tipo VVHP',
    37: 'Açúcar Tipo VHP',
    39: 'Pol  - Açucar Produzido',
    45: 'Etanol Anidro',
    46: 'Etanol Hidratado',
    47: 'Etanol Hidratado Industrial',
    48: 'Etanol Anidro Europeu',
    49: 'Etanol de Segunda',
    50: 'Teor Alcoólico Etanol Anidro',
    51: 'Teor Alcoólico Etanol Hidratado ',
    52: 'Teor Alcoólico Etanol Hidratado Industrial',
    53: 'Teor Alcoólico Etanol Anidro Europeu',
    54: 'Teor Alcoólico Etanol de 2º',
    60: 'Levedura Inativa Seca 300 g/kg',
    61: 'Levedura Inativa Seca 350 g/kg',
    62: 'Levedura Inativa Seca 370 g/kg',
    63: 'Levedura Inativa Seca 390 g/kg',
    64: 'Levedura Inativa Seca 400 g/kg',
    69: 'Óleo Fúsel',
    76: 'Açúcar em Processo - Silo',
    79: 'Recuperação da Fábrica',
    80: 'Pureza do Xarope - Cálculo do Processo',
    81: 'Pureza do Mel Desviado - Cálculo do Processo',
    82: 'Pureza do Açúcar - Cálculo do Processo',
    85: 'Pol Clarificado Fábrica',
    86: 'Quantidade Clarificado Fábrica',
    88: 'Pol Pré-Evaporado Fábrica',
    89: 'Quantidade Pré-Evaporado Fábrica',
    91: 'Pol Xarope Fábrica',
    92: 'Quantidade Xarope Fábrica',
    94: 'Pol Massa "A" Fábrica',
    95: 'Quantidade Massa "A" Fábrica',
    97: 'Pol Massa "B" Fábrica',
    98: 'Quantidade Massa "B" Fábrica',
    100: 'Pol Massa "C" Fábrica',
    101: 'Quantidade Massa "C" Fábrica',
    103: 'Pol Mel A Pobre Diluído Fábrica',
    104: 'Quantidade Mel A Pobre Diluído Fábrica',
    106: 'Pol Mel B  Diluído Fábrica',
    107: 'Quantidade Mel B  Diluído Fábrica',
    109: 'Pol Mel A Rico + Magma C Fábrica',
    110: 'Quantidade Mel A Rico + Magma C Fábrica',
    112: 'Quantidade Mel A Rico FZ-1000 Fábrica',
    113: 'Pol Mel A Rico FZ-1000 Fábrica',
    115: 'Quantidade Mel A Pobre FZ-1000 Fábrica',
    116: 'Pol Mel A Pobre FZ-1000 Fábrica',
    118: 'Quantidade Mel B Kont 10/14 Fábrica',
    119: 'Pol Mel B Kont 10/14 Fábrica',
    121: 'Quantidade Mel Final Fábrica',
    122: 'Pol Mel Final Fábrica',
    124: 'Quantidade Magma B Fábrica',
    125: 'Pol Magma B Fábrica',
    127: 'Quantidade Açúcar Diluído Fábrica',
    128: 'Pol Açúcar Diluído Fábrica',
    130: 'Quantidade Caldo Filtrado Fábrica',
    131: 'Pol Caldo Filtrado Fábrica',
    133: 'Quantidade Lodo Fábrica',
    134: 'Pol Lodo Fábrica',
    135: 'AÇÚCAR EM PROCESSO TOTAL (Dia Anterior)',
    144: 'Volume em Processo - Dornas',
    145: 'ºGL Vinho em Processo - Dornas',
    147: 'Volume em Processo - Cubas',
    148: 'ºGL Vinho em Processo - Cubas',
    150: 'Volume em Processo - Volante',
    151: 'ºGL Vinho em Processo - Volante',
    153: 'Volume em Processo - Caixa de Vinho Bruto',
    154: 'ºGL Vinho em Processo - Caixa de Vinho Bruto',
    155: 'Etanol 100% - Caixa de Vinho Centrifugado',
    156: 'Tanque de Etanol Anidro - Peneira Nº 1 e 2',
    157: 'Tanque de Etanol Hidratado - Peneira Nº 1 e 2',
    161: 'Quantidade Total Mel Final Tanque Nº1',
    162: 'ART % Mel Final Tanque Nº1',
    164: 'Quantidade Total Mel Final Tanque Nº2',
    165: 'ART % Mel Final Tanque Nº2',
    167: 'Quantidade Total Mel Final Tanque Nº3',
    168: 'Volume Mel Final Tanque Nº3',
    169: 'ART % Mel Final Tanque Nº3',
    171: 'Quantidade Total Mel Final Tanque Nº4',
    172: 'Volume Mel Final Tanque Nº4',
    173: 'ART % Mel Final Tanque Nº4',
    175: 'Quantidade Total Mel Final Tanque Nº5',
    176: 'Volume Mel Final Tanque Nº5',
    177: 'ART % Mel Final Tanque Nº5',
    180: 'Volume Mel Final Tanque Nº6',
    181: 'Densidade Mel Final Tanque Nº6',
    182: 'ART % Mel Final Tanque Nº6',
    184: 'Quantidade Total Mel Final Tanque Nº7',
    185: 'ART % Mel Final Tanque Nº7',
    190: 'Pol Clarificado Fábrica',
    191: 'Pureza Clarificado Fábrica',
    192: 'Quantidade Clarificado Fábrica',
    194: 'Pol Pré-Evaporado Fábrica',
    195: 'Pureza Pré-Evaporado Fábrica',
    196: 'Quantidade Pré-Evaporado Fábrica',
    198: 'Pol Xarope Fábrica',
    199: 'Pureza Xarope Fábrica',
    200: 'Quantidade Xarope Fábrica',
    202: 'Pol Massa "A" Fábrica',
    203: 'Pureza Massa "A" Fábrica',
    204: 'Quantidade Massa "A" Fábrica',
    207: 'Pureza Massa "B" Fábrica',
    208: 'Quantidade Massa "B" Fábrica',
    210: 'Pol Massa "C" Fábrica',
    211: 'Pureza Massa "C" Fábrica',
    212: 'Quantidade Massa "C" Fábrica',
    214: 'Pol Mel A Pobre Diluído Fábrica',
    215: 'Pureza Mel A Pobre Diluído Fábrica',
    216: 'Quantidade Mel A Pobre Diluído Fábrica',
    218: 'Pol Mel B  Diluído Fábrica',
    219: 'Pureza Mel B  Diluído Fábrica',
    220: 'Quantidade Mel B  Diluído Fábrica',
    222: 'Pol Mel A Rico FZ-1000 Fábrica',
    223: 'Pureza Mel A Rico FZ-1000 Fábrica',
    224: 'Quantidade Mel A Rico FZ-1000 Fábrica',
    226: 'Pol Mel A Pobre FZ-1000 Fábrica',
    227: 'Pureza Mel A Pobre FZ-1000 Fábrica',
    228: 'Quantidade Mel A Pobre FZ-1000 Fábrica',
    230: 'Pol Mel B Kont 10/14 Fábrica',
    231: 'Pureza Mel B Kont 10/14 Fábrica',
    232: 'Quantidade Mel B Kont 10/14 Fábrica',
    235: 'Pureza Magma B Fábrica',
    236: 'Quantidade Magma B Fábrica',
    238: 'Pol Açúcar Diluído Fábrica',
    239: 'Pureza Açúcar Diluído Fábrica',
    240: 'Quantidade Açúcar Diluído Fábrica',
    242: 'Pol Caldo Filtrado Fábrica',
    243: 'Pureza Caldo Filtrado Fábrica',
    244: 'Quantidade Caldo Filtrado Fábrica',
    246: 'Pol Lodo Fábrica',
    247: 'Pureza Lodo Fábrica',
    248: 'Quantidade Lodo Fábrica',
    249: 'ETANOL EM PROCESSO TOTAL (Dia Anterior)',
    258: 'Volume em Processo - Dornas',
    259: 'Levedo Vinho Bruto',
    261: 'Volume em Processo - Cubas',
    262: 'Levedo Fermento Tratado',
    264: 'Volume em Processo - Volante',
    265: 'Levedo Dorna Volante',
    267: 'Volume em Processo - Caixa de Vinho Bruto',
    268: 'Levedo Vinho Bruto',
    271: 'Volume em Processo - Dornas',
    272: 'Levedo Vinho Bruto',
    274: 'Volume em Processo - Cubas',
    275: 'Levedo Fermento Tratado',
    277: 'Volume em Processo - Volante',
    278: 'Levedo Dorna Volante',
    280: 'Volume em Processo - Caixa de Vinho Bruto',
    281: 'Levedo Vinho Bruto',
    286: 'Mel Vendido - ART Recuperado',
    287: 'Xarope Vendido - ART Recuperado',
    288: 'Creme de Levedura Vendido - ART Recuperado (%)',
    295: 'Fibra Bagaço 6º Terno - Geral',
    296: 'Análise de Cana (Fibra Real) - Fibra',
    298: 'Pol do Bagaço',
    299: 'AR - Bagaço 6º Terno',
    301: 'Torta de Filtro Produzida',
    303: 'Pol da Torta',
    304: 'AR - Torta de Filtro Geral',
    307: 'Perda Vinhaça - ºGL',
    308: 'Perda Flegmaça - ºGL',
    309: 'Vinhaça + Flegmaça Produzida',
    313: 'Etanol Vinho Centrifugado',
    315: 'Levedo Vinho Bruto',
    316: 'Levedo Vinho Centrifugado',
    317: 'Creme - Teor de Levedo',
    318: 'Volume - Vinho Bruto (RF)',
    320: 'Acidez Mosto',
    321: 'Densidade Mosto',
    322: 'Acidez Vinho Centrifugado',
    323: 'Volume - Ácido Sulfúrico (RF)',
    324: 'Volume - Vinho Centrifugado (RF)',
    325: 'Massa de Mosto - RF',
    327: 'Glicerol Vinho Centrifugado',
    329: 'Levedura Produzida Base Seca - RF',
    331: 'ARRT Vinho Centrifugado',
    334: 'Levedura - ART Recuperado',
    335: 'Creme de Levedura Vendido - ART Recuperado',
    336: 'PERDAS DETERMINADAS EM ÁGUAS - TOTAL',
    338: 'ART - Água Lavagem de Esteiras',
    339: 'Volume Água Lavagem de Esteiras',
    341: 'ART - Água das Colunas Barométricas',
    342: 'Volume Água das Colunas Barométricas',
    344: 'ART - Água Residual Geral',
    345: 'Volume Água Residual Geral',
}

INPUT_ROWS = [11, 12, 13, 14, 20, 21, 22, 36, 37, 39, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 60, 61, 62, 63, 64, 69, 76, 79, 80, 81, 82, 85, 86, 88, 89, 91, 92, 94, 95, 97, 98, 100, 101, 103, 104, 106, 107, 109, 110, 112, 113, 115, 116, 118, 119, 121, 122, 124, 125, 127, 128, 130, 131, 133, 134, 135, 144, 145, 147, 148, 150, 151, 153, 154, 155, 156, 157, 161, 162, 164, 165, 167, 168, 169, 171, 172, 173, 175, 176, 177, 180, 181, 182, 184, 185, 190, 191, 192, 194, 195, 196, 198, 199, 200, 202, 203, 204, 207, 208, 210, 211, 212, 214, 215, 216, 218, 219, 220, 222, 223, 224, 226, 227, 228, 230, 231, 232, 235, 236, 238, 239, 240, 242, 243, 244, 246, 247, 248, 249, 258, 259, 261, 262, 264, 265, 267, 268, 271, 272, 274, 275, 277, 278, 280, 281, 286, 287, 288, 295, 296, 298, 299, 301, 303, 304, 307, 308, 309, 313, 315, 316, 317, 318, 320, 321, 322, 323, 324, 325, 327, 329, 331, 334, 335, 336, 338, 339, 341, 342, 344, 345]

ENTRADAS_BASE = {
    11: 49269.668,
    12: 11.88,
    13: 0.87,
    14: 12.47,
    20: 0.0,
    21: 0.0,
    22: 0.0,
    36: 0.0,
    37: 3697.9930000000004,
    39: 99.584999084483,
    45: 0.0,
    46: 900.0,
    47: 430.0,
    48: 0.0,
    49: 0.0,
    50: 0.0,
    51: 95.3185697283,
    52: 96.29499893188,
    53: 0.0,
    54: 0.0,
    60: 0.0,
    61: 0.0,
    62: 34.4,
    63: 0.0,
    64: 0.0,
    69: 0.0,
    76: 14.581771850586001,
    79: 65.398559275407,
    80: 82.0739974976,
    81: 60.69005751,
    82: 99.60988479584101,
    85: 10.953000068700002,
    86: 4556.41421254347,
    88: 16.857000351,
    89: 385.187513555829,
    91: 49.8190002441,
    92: 444.846501746945,
    94: 76.6549987793,
    95: 1081.16668522983,
    97: 65.8140029907,
    98: 1636.0253053814,
    100: 0.0,
    101: 0.0,
    103: 48.4239997864,
    104: 357.916781650592,
    106: 0.0,
    107: 0.0,
    109: 0.0,
    110: 0.0,
    112: 34.75918471715401,
    113: 58.736999511700006,
    115: 7.862297802908,
    116: 56.944999694799996,
    118: 0.0,
    119: 45.8930015564,
    121: 0.0,
    122: 0.0,
    124: 243.44852763766303,
    125: 85.0299987793,
    127: 4.46660080191,
    128: 5.0399999619,
    130: 42.24659016734601,
    131: 6.5170001984,
    133: 32.905558081112005,
    134: 6.5170001984,
    135: 2071.54038646129,
    144: 8348.7421875,
    145: 9.3999996185,
    147: 413.962951660156,
    148: 4.3000001907,
    150: 1277.31591796875,
    151: 10.6999998093,
    153: 27.076667785645004,
    154: 9.3999996185,
    155: 12.032175933792,
    156: 0.0,
    157: 0.0,
    161: 0.0,
    162: 0.0,
    164: 0.0,
    165: 0.0,
    167: 0.0,
    168: 0.0,
    169: 0.0,
    171: 0.0,
    172: 0.0,
    173: 0.0,
    175: 0.0,
    176: 0.0,
    177: 0.0,
    180: 1809.0,
    181: 1.4342649619690002,
    182: 64.1900024414,
    184: 0.0,
    185: 0.0,
    190: 10.953000068700002,
    191: 81.6780014038,
    192: 4556.41421254347,
    194: 16.857000351,
    195: 81.4349975586,
    196: 385.187513555829,
    198: 49.8190002441,
    199: 82.0739974976,
    200: 444.846501746945,
    202: 76.6549987793,
    203: 82.6019973755,
    204: 1081.16668522983,
    207: 71.15000152590001,
    208: 1636.0253053814,
    210: 0.0,
    211: 0.0,
    212: 0.0,
    214: 48.4239997864,
    215: 71.8460006714,
    216: 357.916781650592,
    218: 0.0,
    219: 0.0,
    220: 0.0,
    222: 58.736999511700006,
    223: 72.5149993896,
    224: 34.75918471715401,
    226: 56.944999694799996,
    227: 71.71900177,
    228: 7.862297802908,
    230: 45.8930015564,
    231: 53.208999633800005,
    232: 0.0,
    235: 91.3320007324,
    236: 243.44852763766303,
    238: 5.0399999619,
    239: 59.293998718299996,
    240: 4.46660080191,
    242: 6.5170001984,
    243: 80.7559967041,
    244: 42.24659016734601,
    246: 6.5170001984,
    247: 80.7559967041,
    248: 32.905558081112005,
    249: 2577.71494413967,
    258: 8348.7421875,
    259: 12.833333333333002,
    261: 413.962951660156,
    262: 25.666666666667002,
    264: 1277.31591796875,
    265: 1.1000000139000001,
    267: 27.076667785645004,
    268: 12.833333333333002,
    271: 8042.32666015625,
    272: 12.0,
    274: 227.17698669433602,
    275: 25.666666666667002,
    277: 1022.48046875,
    278: 1.29999999205,
    280: 26.639722824097003,
    281: 12.0,
    286: 0.0,
    287: 0.0,
    288: 0.0,
    295: 46.24488910629,
    296: 11.352166493733,
    298: 1.890579046288,
    299: 0.297053111372,
    301: 1913.91,
    303: 2.140666723267,
    304: 0.1499999985,
    307: 0.047166667267,
    308: 0.059166667367,
    309: 15861.4855957031,
    313: 8.816666603083002,
    315: 12.833333333333002,
    316: 0.166666669167,
    317: 63.333333333333,
    318: 21290.4921875,
    320: 2.0109999974329997,
    321: 1.0945504293750001,
    322: 2.5346666177,
    323: 9.323,
    324: 17021.1586624482,
    325: 13271.9111384919,
    327: 4.399999936417,
    329: 94.380806725449,
    331: 0.057833333167000005,
    334: 68.8,
    335: 0.0,
    336: 288.14380124833,
    338: 0.04965372308700001,
    339: 11439.22265625,
    341: 0.066474579875,
    342: 360000.0,
    344: 0.29634502812200003,
    345: 15355.529296875,
}

def _iferror(fn, fb):
    try:
        r = fn()
        return fb if (r != r) else r
    except Exception:
        return fb

def calcular(entradas):
    e = dict(ENTRADAS_BASE)
    if entradas: e.update(entradas)
    L = {}
    L[36] = e[36]  # Açúcar Tipo VVHP
    L[37] = e[37]  # Açúcar Tipo VHP
    L[35] = (L[36]+L[37])  # AÇÚCAR TOTAL
    L[39] = e[39]  # Pol  - Açucar Produzido
    L[38] = L[39]/0.95  # ART - Açúcar Produzido
    L[34] = L[35]*L[38]/100  # Açúcar - ART Recuperado
    L[69] = e[69]  # Óleo Fúsel
    L[68] = L[69]/1000*0.8*1.54428  # Óleo Fúsel - ART Recuperado
    L[76] = e[76]  # Açúcar em Processo - Silo
    L[97] = e[97]  # Pol Massa "B" Fábrica
    L[98] = e[98]  # Quantidade Massa "B" Fábrica
    L[96] = L[98]*L[97]/100  # Quantidade Pol Massa "B" Fábrica
    L[130] = e[130]  # Quantidade Caldo Filtrado Fábrica
    L[131] = e[131]  # Pol Caldo Filtrado Fábrica
    L[129] = _iferror(lambda: L[130]*L[131]/100, 0)  # Quantidade Pol Caldo Filtrado Fábrica
    L[100] = e[100]  # Pol Massa "C" Fábrica
    L[101] = e[101]  # Quantidade Massa "C" Fábrica
    L[99] = L[100]*L[101]/100  # Quantidade Pol Massa "C" Fábrica
    L[133] = e[133]  # Quantidade Lodo Fábrica
    L[134] = e[134]  # Pol Lodo Fábrica
    L[132] = _iferror(lambda: L[133]*L[134]/100, 0)  # Quantidade Pol Lodo Fábrica
    L[104] = e[104]  # Quantidade Mel A Pobre Diluído Fábrica
    L[103] = e[103]  # Pol Mel A Pobre Diluído Fábrica
    L[102] = L[103]*L[104]/100  # Quantidade Pol Mel A Pobre Diluído Fábrica
    L[106] = e[106]  # Pol Mel B  Diluído Fábrica
    L[107] = e[107]  # Quantidade Mel B  Diluído Fábrica
    L[105] = _iferror(lambda: L[106]*L[107]/100, 0)  # Quantidade Pol Mel B  Diluído Fábrica
    L[109] = e[109]  # Pol Mel A Rico + Magma C Fábrica
    L[110] = e[110]  # Quantidade Mel A Rico + Magma C Fábrica
    L[108] = _iferror(lambda: L[109]*L[110]/100, 0)  # Quantidade Pol Mel A Rico + Magma C Fábrica
    L[112] = e[112]  # Quantidade Mel A Rico FZ-1000 Fábrica
    L[113] = e[113]  # Pol Mel A Rico FZ-1000 Fábrica
    L[111] = _iferror(lambda: L[112]*L[113]/100, 0)  # Quantidade Pol Mel A Rico FZ-1000 Fábrica
    L[115] = e[115]  # Quantidade Mel A Pobre FZ-1000 Fábrica
    L[116] = e[116]  # Pol Mel A Pobre FZ-1000 Fábrica
    L[114] = _iferror(lambda: L[115]*L[116]/100, 0)  # Quantidade Pol Mel A Pobre FZ-1000 Fábrica
    L[85] = e[85]  # Pol Clarificado Fábrica
    L[86] = e[86]  # Quantidade Clarificado Fábrica
    L[84] = L[85]*L[86]/100  # Quantidade Pol Clarificado Fábrica
    L[118] = e[118]  # Quantidade Mel B Kont 10/14 Fábrica
    L[119] = e[119]  # Pol Mel B Kont 10/14 Fábrica
    L[117] = _iferror(lambda: L[118]*L[119]/100, 0)  # Quantidade Pol Mel B Kont 10/14 Fábrica
    L[88] = e[88]  # Pol Pré-Evaporado Fábrica
    L[89] = e[89]  # Quantidade Pré-Evaporado Fábrica
    L[87] = L[88]*L[89]/100  # Quantidade Pol Pré-Evaporado Fábrica
    L[121] = e[121]  # Quantidade Mel Final Fábrica
    L[122] = e[122]  # Pol Mel Final Fábrica
    L[120] = _iferror(lambda: L[121]*L[122]/100, 0)  # Quantidade Pol Mel Final Fábrica
    L[91] = e[91]  # Pol Xarope Fábrica
    L[92] = e[92]  # Quantidade Xarope Fábrica
    L[90] = L[91]*L[92]/100  # Quantidade Pol Xarope Fábrica
    L[124] = e[124]  # Quantidade Magma B Fábrica
    L[125] = e[125]  # Pol Magma B Fábrica
    L[123] = _iferror(lambda: L[124]*L[125]/100, 0)  # Quantidade Pol Magma B Fábrica
    L[94] = e[94]  # Pol Massa "A" Fábrica
    L[95] = e[95]  # Quantidade Massa "A" Fábrica
    L[93] = L[94]*L[95]/100  # Quantidade Pol Massa "A" Fábrica
    L[128] = e[128]  # Pol Açúcar Diluído Fábrica
    L[127] = e[127]  # Quantidade Açúcar Diluído Fábrica
    L[126] = _iferror(lambda: L[127]*L[128]/100, 0)  # Quantidade Pol Açúcar Diluído Fábrica
    L[83] = L[84]+L[87]+L[90]+L[93]+L[96]+L[99]+L[102]+L[105]+L[108]+L[111]+L[114]+L[117]+L[120]+L[123]+L[126]+L[129]+L[132]  # Quantidade Pol Fábrica de Açúcar - TOTAL
    L[80] = e[80]  # Pureza do Xarope - Cálculo do Processo
    L[81] = e[81]  # Pureza do Mel Desviado - Cálculo do Processo
    L[82] = e[82]  # Pureza do Açúcar - Cálculo do Processo
    L[78] = ((70) if (((L[80]==0) or (L[81]==0) or (L[82]==0))) else ((L[82]*(L[80]-L[81])/(L[80]*(L[82]-L[81]))*100)))  # Recuperação da Fábrica - Cálculo do Processo
    L[77] = L[78]*L[83]/100  # Açúcar em Processo - Processo Fábrica
    L[75] = L[76]+L[77]  # AÇÚCAR EM PROCESSO ATUAL
    L[135] = e[135]  # AÇÚCAR EM PROCESSO TOTAL (Dia Anterior)
    L[74] = L[75]-L[135]  # DIFERENÇA PROCESSO AÇÚCAR
    L[73] = L[74]*L[38]/100  # Processo Açúcar - ART
    L[45] = e[45]  # Etanol Anidro
    L[46] = e[46]  # Etanol Hidratado
    L[47] = e[47]  # Etanol Hidratado Industrial
    L[48] = e[48]  # Etanol Anidro Europeu
    L[49] = e[49]  # Etanol de Segunda
    L[50] = e[50]  # Teor Alcoólico Etanol Anidro
    L[51] = e[51]  # Teor Alcoólico Etanol Hidratado
    L[52] = e[52]  # Teor Alcoólico Etanol Hidratado Industrial
    L[53] = e[53]  # Teor Alcoólico Etanol Anidro Europeu
    L[54] = e[54]  # Teor Alcoólico Etanol de 2º
    L[44] = (L[45]*L[50]+L[46]*L[51]+L[47]*L[52]+L[48]*L[53]+L[49]*L[54])/100  # Etanol Total Produzido 100% (Balanço de ART)
    L[43] = L[44]*1.54428  # Etanol - ART Recuperado
    L[249] = e[249]  # ETANOL EM PROCESSO TOTAL (Dia Anterior)
    L[187] = L[78]  # Recuperação da Fábrica - Cálculo do Processo
    L[194] = e[194]  # Pol Pré-Evaporado Fábrica
    L[195] = e[195]  # Pureza Pré-Evaporado Fábrica
    L[196] = e[196]  # Quantidade Pré-Evaporado Fábrica
    L[193] = L[196]*(L[194]/100*1.0526+((9.8648-(0.1039*L[195]))/100))  # Quantidade ART Pré-Evaporado Fábrica
    L[226] = e[226]  # Pol Mel A Pobre FZ-1000 Fábrica
    L[227] = e[227]  # Pureza Mel A Pobre FZ-1000 Fábrica
    L[228] = e[228]  # Quantidade Mel A Pobre FZ-1000 Fábrica
    L[225] = L[228]*(L[226]/100*1.0526+((9.8648-(0.1039*L[227]))/100))  # Quantidade ART Mel A Pobre FZ-1000 Fábrica
    L[200] = e[200]  # Quantidade Xarope Fábrica
    L[198] = e[198]  # Pol Xarope Fábrica
    L[199] = e[199]  # Pureza Xarope Fábrica
    L[197] = L[200]*(L[198]/100*1.0526+((9.8648-(0.1039*L[199]))/100))  # Quantidade ART Xarope Fábrica
    L[232] = e[232]  # Quantidade Mel B Kont 10/14 Fábrica
    L[230] = e[230]  # Pol Mel B Kont 10/14 Fábrica
    L[231] = e[231]  # Pureza Mel B Kont 10/14 Fábrica
    L[229] = L[232]*(L[230]/100*1.0526+((9.8648-(0.1039*L[231]))/100))  # Quantidade ART Mel B Kont 10/14 Fábrica
    L[202] = e[202]  # Pol Massa "A" Fábrica
    L[203] = e[203]  # Pureza Massa "A" Fábrica
    L[204] = e[204]  # Quantidade Massa "A" Fábrica
    L[201] = L[204]*(L[202]/100*1.0526+((9.8648-(0.1039*L[203]))/100))  # Quantidade ART Massa "A" Fábrica
    L[234] = L[125]  # Pol Magma B Fábrica
    L[235] = e[235]  # Pureza Magma B Fábrica
    L[236] = e[236]  # Quantidade Magma B Fábrica
    L[233] = L[236]*(L[234]/100*1.0526+((9.8648-(0.1039*L[235]))/100))  # Quantidade ART Magma B Fábrica
    L[208] = e[208]  # Quantidade Massa "B" Fábrica
    L[206] = L[97]  # Pol Massa "B" Fábrica
    L[207] = e[207]  # Pureza Massa "B" Fábrica
    L[205] = L[208]*(L[206]/100*1.0526+((9.8648-(0.1039*L[207]))/100))  # Quantidade ART Massa "B" Fábrica
    L[240] = e[240]  # Quantidade Açúcar Diluído Fábrica
    L[238] = e[238]  # Pol Açúcar Diluído Fábrica
    L[239] = e[239]  # Pureza Açúcar Diluído Fábrica
    L[237] = L[240]*(L[238]/100*1.0526+((9.8648-(0.1039*L[239]))/100))  # Quantidade ART Açúcar Diluído Fábrica
    L[210] = e[210]  # Pol Massa "C" Fábrica
    L[211] = e[211]  # Pureza Massa "C" Fábrica
    L[212] = e[212]  # Quantidade Massa "C" Fábrica
    L[209] = L[212]*(L[210]/100*1.0526+((9.8648-(0.1039*L[211]))/100))  # Quantidade ART Massa "C" Fábrica
    L[242] = e[242]  # Pol Caldo Filtrado Fábrica
    L[243] = e[243]  # Pureza Caldo Filtrado Fábrica
    L[244] = e[244]  # Quantidade Caldo Filtrado Fábrica
    L[241] = L[244]*(L[242]/100*1.0526+((9.8648-(0.1039*L[243]))/100))  # Quantidade ART Caldo Filtrado Fábrica
    L[216] = e[216]  # Quantidade Mel A Pobre Diluído Fábrica
    L[214] = e[214]  # Pol Mel A Pobre Diluído Fábrica
    L[215] = e[215]  # Pureza Mel A Pobre Diluído Fábrica
    L[213] = L[216]*(L[214]/100*1.0526+((9.8648-(0.1039*L[215]))/100))  # Quantidade ART Mel A Pobre Diluído Fábrica
    L[248] = e[248]  # Quantidade Lodo Fábrica
    L[246] = e[246]  # Pol Lodo Fábrica
    L[247] = e[247]  # Pureza Lodo Fábrica
    L[245] = L[248]*(L[246]/100*1.0526+((9.8648-(0.1039*L[247]))/100))  # Quantidade ART Lodo Fábrica
    L[224] = e[224]  # Quantidade Mel A Rico FZ-1000 Fábrica
    L[222] = e[222]  # Pol Mel A Rico FZ-1000 Fábrica
    L[223] = e[223]  # Pureza Mel A Rico FZ-1000 Fábrica
    L[221] = L[224]*(L[222]/100*1.0526+((9.8648-(0.1039*L[223]))/100))  # Quantidade ART Mel A Rico FZ-1000 Fábrica
    L[218] = e[218]  # Pol Mel B  Diluído Fábrica
    L[219] = e[219]  # Pureza Mel B  Diluído Fábrica
    L[220] = e[220]  # Quantidade Mel B  Diluído Fábrica
    L[217] = L[220]*(L[218]/100*1.0526+((9.8648-(0.1039*L[219]))/100))  # Quantidade ART Mel B  Diluído Fábrica
    L[192] = e[192]  # Quantidade Clarificado Fábrica
    L[190] = e[190]  # Pol Clarificado Fábrica
    L[191] = e[191]  # Pureza Clarificado Fábrica
    L[189] = L[192]*(L[190]/100*1.0526+((9.8648-(0.1039*L[191]))/100))  # Quantidade ART Clarificado Fábrica
    L[188] = (L[189]+L[193]+L[197]+L[201]+L[205]+L[209]+L[213]+L[217]+L[221]+L[225]+L[229]+L[233]+L[237]+L[241]+L[245])  # Quantidade ART Fábrica de Açúcar - TOTAL
    L[186] = (L[188]*(1-L[187]/100)*89/100)*0.64755  # Etanol em Processo - Mel Desviado
    L[161] = e[161]  # Quantidade Total Mel Final Tanque Nº1
    L[162] = e[162]  # ART % Mel Final Tanque Nº1
    L[160] = _iferror(lambda: L[161]*L[162]/100, 0)  # Quantidade ART Mel Final Tanque Nº1
    L[164] = e[164]  # Quantidade Total Mel Final Tanque Nº2
    L[165] = e[165]  # ART % Mel Final Tanque Nº2
    L[163] = _iferror(lambda: L[164]*L[165]/100, 0)  # Quantidade ART Mel Final Tanque Nº2
    L[169] = e[169]  # ART % Mel Final Tanque Nº3
    L[167] = e[167]  # Quantidade Total Mel Final Tanque Nº3
    L[166] = _iferror(lambda: L[167]*L[169]/100, 0)  # Quantidade ART Mel Final Tanque Nº3
    L[171] = e[171]  # Quantidade Total Mel Final Tanque Nº4
    L[173] = e[173]  # ART % Mel Final Tanque Nº4
    L[170] = _iferror(lambda: L[171]*L[173]/100, 0)  # Quantidade ART Mel Final Tanque Nº4
    L[177] = e[177]  # ART % Mel Final Tanque Nº5
    L[175] = e[175]  # Quantidade Total Mel Final Tanque Nº5
    L[174] = _iferror(lambda: L[175]*L[177]/100, 0)  # Quantidade ART Mel Final Tanque Nº5
    L[180] = e[180]  # Volume Mel Final Tanque Nº6
    L[181] = e[181]  # Densidade Mel Final Tanque Nº6
    L[179] = L[180]*L[181]  # Quantidade Total Mel Final Tanque Nº6
    L[182] = e[182]  # ART % Mel Final Tanque Nº6
    L[178] = _iferror(lambda: L[179]*L[182]/100, 0)  # Quantidade ART Mel Final Tanque Nº6
    L[184] = e[184]  # Quantidade Total Mel Final Tanque Nº7
    L[185] = e[185]  # ART % Mel Final Tanque Nº7
    L[183] = _iferror(lambda: L[184]*L[185]/100, 0)  # Quantidade ART Mel Final Tanque Nº7
    L[159] = (L[160]+L[163]+L[166]+L[170]+L[174]+L[178]+L[183])  # Quantidade ART Mel Final Tanque - Geral
    L[158] = L[159]*0.89*0.64755  # Etanol em Processo - Tanques de Mel
    L[144] = e[144]  # Volume em Processo - Dornas
    L[145] = e[145]  # ºGL Vinho em Processo - Dornas
    L[143] = _iferror(lambda: L[144]*L[145]/100, 0)  # Etanol 100% - Dornas
    L[147] = e[147]  # Volume em Processo - Cubas
    L[148] = e[148]  # ºGL Vinho em Processo - Cubas
    L[146] = _iferror(lambda: L[147]*L[148]/100, 0)  # Etanol 100% - Cubas
    L[150] = e[150]  # Volume em Processo - Volante
    L[151] = e[151]  # ºGL Vinho em Processo - Volante
    L[149] = _iferror(lambda: L[150]*L[151]/100, 0)  # Etanol 100% - Volante
    L[153] = e[153]  # Volume em Processo - Caixa de Vinho Bruto
    L[154] = e[154]  # ºGL Vinho em Processo - Caixa de Vinho Bruto
    L[152] = _iferror(lambda: L[153]*L[154]/100, 0)  # Etanol 100% - Caixa de Vinho Bruto
    L[155] = e[155]  # Etanol 100% - Caixa de Vinho Centrifugado
    L[156] = e[156]  # Tanque de Etanol Anidro - Peneira Nº 1 e 2
    L[157] = e[157]  # Tanque de Etanol Hidratado - Peneira Nº 1 e 2
    L[142] = L[143]+L[146]+L[149]+L[152]+L[155]+L[156]+L[157]  # Etanol em Processo - Processo Etanol
    L[141] = (L[142]+L[158]+L[186])  # ETANOL EM PROCESSO ATUAL
    L[140] = L[141]-L[249]  # DIFERENÇA PROCESSO ETANOL
    L[139] = L[140]*1.54428  # Processo Etanol - ART
    L[274] = e[274]  # Volume em Processo - Cubas
    L[275] = e[275]  # Levedo Fermento Tratado
    L[273] = _iferror(lambda: L[274]*L[275]/100, 0)  # Levedo - Cubas
    L[277] = e[277]  # Volume em Processo - Volante
    L[278] = e[278]  # Levedo Dorna Volante
    L[276] = _iferror(lambda: L[277]*L[278]/100, 0)  # Levedo - Volante
    L[272] = e[272]  # Levedo Vinho Bruto
    L[271] = e[271]  # Volume em Processo - Dornas
    L[270] = _iferror(lambda: L[271]*L[272]/100, 0)  # Levedo - Dornas
    L[280] = e[280]  # Volume em Processo - Caixa de Vinho Bruto
    L[281] = e[281]  # Levedo Vinho Bruto
    L[279] = _iferror(lambda: L[280]*L[281]/100, 0)  # Levedo - Caixa de Vinho Bruto
    L[269] = (L[270]+L[273]+L[276]+L[279])
    L[258] = e[258]  # Volume em Processo - Dornas
    L[259] = e[259]  # Levedo Vinho Bruto
    L[257] = _iferror(lambda: L[258]*L[259]/100, 0)  # Levedo - Dornas
    L[267] = e[267]  # Volume em Processo - Caixa de Vinho Bruto
    L[268] = e[268]  # Levedo Vinho Bruto
    L[266] = _iferror(lambda: L[267]*L[268]/100, 0)  # Levedo - Caixa de Vinho Bruto
    L[261] = e[261]  # Volume em Processo - Cubas
    L[262] = e[262]  # Levedo Fermento Tratado
    L[260] = _iferror(lambda: L[261]*L[262]/100, 0)  # Levedo - Cubas
    L[264] = e[264]  # Volume em Processo - Volante
    L[265] = e[265]  # Levedo Dorna Volante
    L[263] = _iferror(lambda: L[264]*L[265]/100, 0)  # Levedo - Volante
    L[256] = (L[257]+L[260]+L[263]+L[266])  # Levedo em Processo - Processo Etanol
    L[255] = (L[256])
    L[254] = L[255]-L[269]
    L[253] = L[254]*2/3
    L[64] = e[64]  # Levedura Inativa Seca 400 g/kg
    L[60] = e[60]  # Levedura Inativa Seca 300 g/kg
    L[61] = e[61]  # Levedura Inativa Seca 350 g/kg
    L[62] = e[62]  # Levedura Inativa Seca 370 g/kg
    L[63] = e[63]  # Levedura Inativa Seca 390 g/kg
    L[59] = (L[60]+L[61]+L[62]+L[63]+L[64])
    L[58] = L[59]*2  # Levedura - ART Recuperado
    L[286] = e[286]  # Mel Vendido - ART Recuperado
    L[287] = e[287]  # Xarope Vendido - ART Recuperado
    L[285] = (L[286]+L[287])
    L[26] = L[34]+L[43]+L[58]+L[68]+L[73]+L[139]+L[285]+L[253]  # ART TOTAL RECUPERADO
    L[11] = e[11]  # CANA TOTAL
    L[12] = e[12]  # PC
    L[13] = e[13]  # AR Cana (%)
    L[14] = e[14]  # Caldo da Cana - Pol
    L[15] = L[12]*(1+(0.95*L[13]/L[14]))/0.95  # ART Cana
    L[10] = L[11]*L[15]/100  # ART Disponível - Cana
    L[20] = e[20]  # ART Disponível - Mel Comprado
    L[21] = e[21]  # ART Disponível - Xarope Comprado
    L[22] = e[22]  # ART Disponível - Creme de Levedura Comprado
    L[19] = L[20]+L[21]+L[22]
    L[6] = L[10]+L[19]  # ART DISPONÍVEL TOTAL
    L[2] = L[26]/L[6]*100  # Recuperação Industrial
    L[298] = e[298]  # Pol do Bagaço
    L[299] = e[299]  # AR - Bagaço 6º Terno
    L[297] = (L[298]/0.95)+L[299]  # ART - Bagaço
    L[296] = e[296]  # Análise de Cana (Fibra Real) - Fibra
    L[295] = e[295]  # Fibra Bagaço 6º Terno - Geral
    L[294] = L[11]*(L[296]/L[295])  # Bagaço Produzido
    L[293] = L[294]*L[297]/100  # Perda no Bagaço (t)
    L[301] = e[301]  # Torta de Filtro Produzida
    L[304] = e[304]  # AR - Torta de Filtro Geral
    L[303] = e[303]  # Pol da Torta
    L[302] = (L[303]/0.95)+L[304]  # ART - Torta de Filtro
    L[300] = L[301]*L[302]/100  # Perda na Torta de Filtro (t)
    L[336] = e[336]  # PERDAS DETERMINADAS EM ÁGUAS - TOTAL
    L[307] = e[307]  # Perda Vinhaça - ºGL
    L[308] = e[308]  # Perda Flegmaça - ºGL
    L[309] = e[309]  # Vinhaça + Flegmaça Produzida
    L[306] = ((((L[309]*6)/7)*L[307])+(((L[309]*1)/7)*L[308]))/L[309]*1.54428  # ART - Vinhaça + Flegmaça
    L[305] = L[309]*L[306]/100  # Perda na Vinhaça + Flegmaça (t)
    L[332] = L[141]-L[249]  # DIFERENÇA PROCESSO ETANOL
    L[333] = L[44]*1.54428  # Etanol - ART Recuperado
    L[334] = e[334]  # Levedura - ART Recuperado
    L[335] = e[335]  # Creme de Levedura Vendido - ART Recuperado
    L[329] = e[329]  # Levedura Produzida Base Seca - RF
    L[315] = e[315]  # Levedo Vinho Bruto
    L[316] = e[316]  # Levedo Vinho Centrifugado
    L[317] = e[317]  # Creme - Teor de Levedo
    L[318] = e[318]  # Volume - Vinho Bruto (RF)
    L[314] = L[318]*((L[317]-L[315])/(L[317]-L[316]))  # Volume - Vinho Centrifugado (RF)
    L[313] = e[313]  # Etanol Vinho Centrifugado
    L[328] = (L[329]/((L[314]*L[313]*0.01)*0.7893))  # PERDA DE FERMENTO (Kg/Kg Etanol)
    L[320] = e[320]  # Acidez Mosto
    L[321] = e[321]  # Densidade Mosto
    L[322] = e[322]  # Acidez Vinho Centrifugado
    L[323] = e[323]  # Volume - Ácido Sulfúrico (RF)
    L[324] = e[324]  # Volume - Vinho Centrifugado (RF)
    L[325] = e[325]  # Massa de Mosto - RF
    L[319] = (L[324]*L[322]*0.001)-(((L[325]/L[321])*L[320]*0.001)+L[323]*0.97*1.84)  # Massa de Ácido Produzido
    L[312] = (L[319]/(L[314]*L[313]*0.7893/100))  # PRODUÇÃO TOTAL ÁCIDOS (Kg/kg Etanol)
    L[331] = e[331]  # ARRT Vinho Centrifugado
    L[330] = L[331]/(L[313]*0.7893)  # PERDA DE AÇÚCAR (Kg/Kg Etanol)
    L[327] = e[327]  # Glicerol Vinho Centrifugado
    L[326] = L[327]/(L[313]*7.893)  # PRODUÇÃO DE GLICEROL (Kg/Kg Etanol)
    L[311] = 100/(1+(1.19*L[328])+(0.511*L[312])+(0.511*L[326])+(0.511*L[330]))  # Rendimento Fermentativo
    L[310] = ((((L[333]+(L[332]*1.54428)+L[305]))/(L[311]/100))-(L[333]+(L[332]*1.54428)+L[305]))-(L[334]+L[335])  # Perda na Fermentação (t)
    L[292] = (L[293]+L[300]+L[305]+L[310]+L[336])  # PERDAS DETERMINADAS (t)
    L[346] = L[6]-L[26]-L[292]  # PERDAS INDETERMINADAS (t)
    L[360] = L[346]/L[6]*100  # PERDAS INDETERMINADAS
    L[351] = L[292]/L[6]*100  # PERDAS DETERMINADAS
    L[350] = (L[351]+L[360])  # PERDAS TOTAIS (%)
    L[3] = L[350]  # PERDAS TOTAIS (%)
    L[9] = L[10]/L[6]*100
    L[18] = L[19]/L[6]*100
    L[25] = L[26]/L[6]*100
    L[30] = (L[73]+L[139])
    L[29] = L[30]/L[6]*100
    L[27] = L[29]/L[6]*100
    L[28] = (L[34]+L[43]+L[58]+L[68])
    L[33] = L[34]/L[6]*100
    L[42] = L[43]/L[6]*100
    L[57] = L[58]/L[6]*100
    L[67] = L[68]/L[6]*100
    L[72] = L[73]/L[6]*100
    L[79] = e[79]  # Recuperação da Fábrica
    L[138] = L[139]/L[6]*100
    L[168] = e[168]  # Volume Mel Final Tanque Nº3
    L[172] = e[172]  # Volume Mel Final Tanque Nº4
    L[176] = e[176]  # Volume Mel Final Tanque Nº5
    L[252] = L[253]/L[6]*100
    L[284] = L[285]/L[6]*100
    L[288] = e[288]  # Creme de Levedura Vendido - ART Recuperado (%)
    L[291] = (L[292]+L[346])  # PERDAS TOTAIS
    L[338] = e[338]  # ART - Água Lavagem de Esteiras
    L[339] = e[339]  # Volume Água Lavagem de Esteiras
    L[337] = L[339]*L[338]/100  # Perda - Água Lavagem de Esteiras (t)
    L[341] = e[341]  # ART - Água das Colunas Barométricas
    L[342] = e[342]  # Volume Água das Colunas Barométricas
    L[340] = L[342]*L[341]/100  # Perda - Água das Colunas Barométricas (t)
    L[344] = e[344]  # ART - Água Residual Geral
    L[345] = e[345]  # Volume Água Residual Geral
    L[343] = L[345]*L[344]/100  # Perda - Água Residual Geral (t)
    L[352] = L[293]/L[6]*100  # Perda no Bagaço
    L[353] = L[300]/L[6]*100  # Perda na Torta de Filtro
    L[354] = L[305]/L[6]*100  # Perda na Vinhaça + Flegmaça
    L[355] = L[310]/L[6]*100  # Perda na Fermentação
    L[357] = L[337]/L[6]*100  # Perda - Água Lavagem de Esteiras
    L[358] = L[340]/L[6]*100  # Perda nas Colunas Barométricas
    L[359] = L[343]/L[6]*100  # Perda na Água Residual Geral
    L[356] = (L[357]+L[358]+L[359])  # PERDAS DETERMINADAS EM ÁGUAS - TOTAL (%)
    return L

def rotulo(row):
    """Nome 'limpo' do indicador para exibição."""
    nm = ROW_NAME.get(row, f"Linha {row}")
    return str(nm).replace("\u200e", "").strip()


def recuperacao_industrial(entradas):
    return calcular(entradas)[2]


# =====================================================================
#  LEITURA DA PLANILHA (equivalente ao AVERAGEIFS por DATA/PERÍODO/MÉTRICA)
# =====================================================================
HDR = {
    "data": ["data", "date", "dt"],
    "periodo": ["desc_periodo", "periodo", "período", "desc periodo"],
    "metrica": ["dsc_metrica", "metrica", "métrica", "indicador", "desc_metrica"],
    "valor": ["valor", "value", "vlr"],
}


def _achar_coluna(cols, chaves):
    norm = {str(c).strip().lower(): c for c in cols}
    for k in chaves:
        if k in norm:
            return norm[k]
    for c in cols:                       # match parcial
        cl = str(c).strip().lower()
        if any(k in cl for k in chaves):
            return c
    return None


@st.cache_data(show_spinner=False)
def ler_planilha(file_bytes):
    """Lê o Excel e devolve um DataFrame longo: [data, periodo, metrica, valor]."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    melhor = None
    for nome in xls.sheet_names:
        df = xls.parse(nome)
        if df.empty:
            continue
        c_met = _achar_coluna(df.columns, HDR["metrica"])
        c_val = _achar_coluna(df.columns, HDR["valor"])
        c_dat = _achar_coluna(df.columns, HDR["data"])
        if c_met is not None and c_val is not None and c_dat is not None:
            c_per = _achar_coluna(df.columns, HDR["periodo"])
            out = pd.DataFrame({
                "data": pd.to_datetime(df[c_dat], errors="coerce"),
                "periodo": df[c_per].astype(str).str.strip() if c_per is not None else "—",
                "metrica": df[c_met].astype(str),
                "valor": pd.to_numeric(df[c_val], errors="coerce"),
            }).dropna(subset=["data"])
            out["metrica_norm"] = out["metrica"].str.strip().str.lower()
            if melhor is None or len(out) > len(melhor):
                melhor = out
    return melhor


def valores_para_data(df, data_sel, periodo_sel):
    """Para a data/período escolhidos, monta {linha_entrada: valor} (média, como AVERAGEIFS)."""
    sub = df[(df["data"].dt.date == data_sel)]
    if periodo_sel and periodo_sel != "—":
        sub = sub[sub["periodo"] == periodo_sel]
    medias = sub.groupby("metrica_norm")["valor"].mean().to_dict()
    entradas, faltando = {}, []
    for row in INPUT_ROWS:
        nome_norm = str(ROW_NAME[row]).strip().lower()
        if nome_norm in medias and pd.notna(medias[nome_norm]):
            entradas[row] = float(medias[nome_norm])
        else:
            entradas[row] = 0.0          # ausente -> 0 (idêntico ao IFERROR da planilha)
            faltando.append(rotulo(row))
    return entradas, faltando


# =====================================================================
#  ANÁLISE DE SENSIBILIDADE
# =====================================================================
def analisar(entradas, choque=0.01):
    y0 = recuperacao_industrial(entradas)
    res = []
    for row, x0 in entradas.items():
        if x0 == 0:
            dx, rel = choque, False
        else:
            dx, rel = x0 * choque, True
        cen = dict(entradas)
        cen[row] = x0 + dx
        delta = recuperacao_industrial(cen) - y0
        elast = (delta / y0) / choque if (rel and y0) else 0.0
        res.append({"linha": row, "indicador": rotulo(row), "valor_base": x0,
                    "delta_pp": delta, "elasticidade": elast})
    soma = sum(abs(r["elasticidade"]) for r in res) or 1.0
    for r in res:
        r["influencia"] = abs(r["elasticidade"]) / soma * 100
    res.sort(key=lambda r: abs(r["delta_pp"]), reverse=True)
    return y0, res


# =====================================================================
#  INTERFACE (tema escuro)
# =====================================================================
st.set_page_config(page_title="Sensibilidade — Recuperação Industrial",
                   page_icon="🏭", layout="wide")

CSS = """
<style>
:root{
  --bg:#0c0f16; --panel:#141925; --panel2:#1b2230; --line:#28304180;
  --txt:#e8edf6; --muted:#8b96ad; --accent:#3dd6c4;
}
.stApp{ background:radial-gradient(1200px 600px at 15% -10%, #18324a33, transparent),
        radial-gradient(1000px 500px at 110% 10%, #2a1d4a33, transparent), var(--bg); }
section[data-testid="stSidebar"]{ background:var(--panel); border-right:1px solid var(--line); }
h1,h2,h3,h4,label,p,span,div{ color:var(--txt); }
.block-container{ padding-top:2rem; }
.kpi{ background:linear-gradient(135deg,#10212e,#171a2b); border:1px solid var(--line);
      border-radius:18px; padding:22px 26px; }
.kpi .lbl{ color:var(--muted); font-size:.85rem; letter-spacing:.06em; text-transform:uppercase; }
.kpi .val{ font-size:3.0rem; font-weight:800; line-height:1;
           background:linear-gradient(90deg,#3dd6c4,#5aa6ff); -webkit-background-clip:text;
           -webkit-text-fill-color:transparent; }
.kpi .sub{ color:var(--muted); font-size:.9rem; margin-top:6px; }
.rowit{ display:flex; align-items:center; gap:14px; padding:9px 4px; border-bottom:1px solid var(--line); }
.rank{ width:30px; text-align:center; color:var(--muted); font-weight:700; font-variant-numeric:tabular-nums;}
.name{ width:300px; min-width:300px; font-size:.92rem; word-break:break-word; white-space:normal; line-height:1.35;}
.barwrap{ flex:1; background:#0a0d14; border-radius:8px; height:22px; overflow:hidden; border:1px solid var(--line);}
.bar{ height:100%; border-radius:8px 0 0 8px; }
.pct{ width:64px; text-align:right; font-weight:700; font-variant-numeric:tabular-nums;}
.tag{ font-size:.72rem; padding:2px 8px; border-radius:20px; margin-left:6px; }
.up{ background:#10362e; color:#48e0b6; border:1px solid #1c5e4f;}
.down{ background:#3a1620; color:#ff7b8a; border:1px solid #6e2734;}
.legend{ color:var(--muted); font-size:.82rem; }
[data-testid="stSelectbox"] > div > div,
[data-testid="stSelectbox"] > div > div > div{
  background:#000 !important;
  border:1px solid var(--line) !important;
  color:var(--txt) !important;
}
[data-testid="stSelectbox"] svg{ fill:var(--txt) !important; }
[data-baseweb="popover"] ul,
[data-baseweb="menu"]{ background:#000 !important; }
[data-baseweb="menu"] li{ background:#000 !important; color:var(--txt) !important; }
[data-baseweb="menu"] li:hover{ background:#111 !important; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

st.markdown("## 🏭 Recuperação Industrial — Analisador de Sensibilidade")
st.markdown("<p class='legend'>Carregue a planilha de indicadores, escolha a data e veja "
            "o ranking de influência de cada indicador na recuperação industrial.</p>",
            unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Configuração")
    up = st.file_uploader("Planilha de indicadores (.xlsx)", type=["xlsx", "xlsm"])
    st.caption("Estrutura esperada: colunas DATA · DESC_PERIODO · DSC_METRICA · VALOR")

if up is None:
    st.info("⬅️ Envie a planilha Excel com os indicadores para começar.")
    st.stop()

df = ler_planilha(up.getvalue()) if up is not None else None

if df is None or df.empty:
    st.error("Não encontrei colunas de DATA, MÉTRICA e VALOR na planilha. "
             "Verifique se há uma aba com esse formato (ex.: aba 'DB').")
else:
    with st.sidebar:
        periodos = sorted(p for p in df["periodo"].dropna().unique() if p)
        pad = "Diária" if "Diária" in periodos else (periodos[0] if periodos else "—")
        periodo_sel = st.selectbox("Período", periodos or ["—"],
                                   index=(periodos.index(pad) if pad in periodos else 0))
        datas = sorted(df[df["periodo"] == periodo_sel]["data"].dt.date.unique())
        if not datas:
            st.error("Sem datas para este período.")
        else:
            data_sel = st.selectbox("Data de análise", datas, index=len(datas) - 1,
                                    format_func=lambda d: d.strftime("%d/%m/%Y"))
            choque = st.slider("Choque aplicado a cada indicador (%)", 0.5, 10.0, 1.0, 0.5) / 100
            topn = st.slider("Mostrar top N indicadores", 5, 40, 15, 1)

    if datas:
        entradas, faltando = valores_para_data(df, data_sel, periodo_sel)
        y0, res = analisar(entradas, choque)

        # Recuperação Industrial direto da planilha (não recalculada)
        sub_ri = df[(df["data"].dt.date == data_sel)]
        if periodo_sel and periodo_sel != "—":
            sub_ri = sub_ri[sub_ri["periodo"] == periodo_sel]
        mask_ri = sub_ri["metrica_norm"] == "recuperação industrial"
        serie_ri = sub_ri.loc[mask_ri, "valor"].dropna()
        y_planilha = float(serie_ri.mean()) if not serie_ri.empty else None

        c1, c2, c3 = st.columns([1.4, 1, 1])
        with c1:
            val_exib = y_planilha if y_planilha is not None else y0
            sub_label = (f"{data_sel.strftime('%d/%m/%Y')} · {periodo_sel}"
                         if y_planilha is not None
                         else f"{data_sel.strftime('%d/%m/%Y')} · {periodo_sel} (recalculado)")
            st.markdown(f"<div class='kpi'><div class='lbl'>Recuperação Industrial</div>"
                        f"<div class='val'>{val_exib:.2f}%</div>"
                        f"<div class='sub'>{sub_label}</div></div>",
                        unsafe_allow_html=True)
        with c2:
            top = res[0]
            st.markdown(f"<div class='kpi'><div class='lbl'>Maior influência</div>"
                        f"<div class='val' style='font-size:1.5rem'>{top['indicador']}</div>"
                        f"<div class='sub'>{top['influencia']:.1f}% do impacto · "
                        f"{top['delta_pp']:+.3f} p.p.</div></div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div class='kpi'><div class='lbl'>Indicadores avaliados</div>"
                        f"<div class='val' style='font-size:1.5rem'>{len(res)}</div>"
                        f"<div class='sub'>choque de ±{choque*100:.1f}% cada</div></div>",
                        unsafe_allow_html=True)

        st.markdown("### 🏆 Ranking de influência")
        st.markdown("<p class='legend'>Largura da barra = participação no impacto total. "
                    "<span class='tag up'>↑ aumenta</span> ou "
                    "<span class='tag down'>↓ reduz</span> a recuperação ao crescer o indicador.</p>",
                    unsafe_allow_html=True)

        vmax = max((r["influencia"] for r in res[:topn]), default=1) or 1
        html = []
        for i, r in enumerate(res[:topn], 1):
            w = r["influencia"] / vmax * 100
            pos = r["delta_pp"] >= 0
            grad = ("linear-gradient(90deg,#1f8f7f,#3dd6c4)" if pos
                    else "linear-gradient(90deg,#a83246,#ff7b8a)")
            tag = ("<span class='tag up'>↑</span>" if pos else "<span class='tag down'>↓</span>")
            html.append(
                f"<div class='rowit'><div class='rank'>{i}</div>"
                f"<div class='name' title=\"{r['indicador']}\">{r['indicador']}{tag}</div>"
                f"<div class='barwrap'><div class='bar' style='width:{w:.1f}%;background:{grad}'></div></div>"
                f"<div class='pct'>{r['influencia']:.1f}%</div></div>")
        st.markdown("".join(html), unsafe_allow_html=True)

        tab = pd.DataFrame(res)[["indicador", "valor_base", "delta_pp", "elasticidade", "influencia"]]
        tab.columns = ["Indicador", "Valor base", "Δ Recup. (p.p.)", "Elasticidade", "Influência (%)"]
        with st.expander("📋 Ver tabela completa / exportar"):
            st.dataframe(tab, use_container_width=True, height=420)
            st.download_button("⬇️ Baixar ranking (CSV)", tab.to_csv(index=False).encode("utf-8"),
                               f"ranking_{data_sel}.csv", "text/csv")

        if faltando:
            with st.expander(f"⚠️ {len(faltando)} indicador(es) sem dado nesta data (tratados como 0)"):
                st.write(", ".join(faltando))

        # ── Ranking de variação dia anterior ─────────────────────────────────
        st.markdown("---")
        st.markdown("### 📅 Variação em relação ao dia anterior")

        datas_periodo = sorted(df[df["periodo"] == periodo_sel]["data"].dt.date.unique())
        idx_sel = datas_periodo.index(data_sel) if data_sel in datas_periodo else -1

        if idx_sel <= 0:
            st.info("Não há data anterior disponível para o período selecionado.")
        else:
            data_ant = datas_periodo[idx_sel - 1]
            entradas_ant, _ = valores_para_data(df, data_ant, periodo_sel)

            deltas_raw = []
            for row in INPUT_ROWS:
                nome = rotulo(row)
                v_atual = entradas.get(row, 0.0)
                v_ant   = entradas_ant.get(row, 0.0)
                diff    = v_atual - v_ant
                pct     = diff / abs(v_ant) * 100 if v_ant != 0 else 0.0
                deltas_raw.append({"linha": row, "indicador": nome,
                                   "valor_atual": v_atual, "valor_anterior": v_ant,
                                   "delta_abs": diff, "delta_pct": pct})

            seen = {}
            for r in deltas_raw:
                nome = r["indicador"]
                if nome not in seen or abs(r["delta_abs"]) > abs(seen[nome]["delta_abs"]):
                    seen[nome] = r
            deltas = list(seen.values())

            deltas.sort(key=lambda r: abs(r["delta_abs"]), reverse=True)

            st.markdown(
                f"<p class='legend'>Comparando <b>{data_sel.strftime('%d/%m/%Y')}</b> "
                f"com <b>{data_ant.strftime('%d/%m/%Y')}</b>. "
                f"Ordenado pela maior variação absoluta.</p>",
                unsafe_allow_html=True)

            vmax_d = max((abs(r["delta_abs"]) for r in deltas[:topn]), default=1) or 1
            html_d = []
            for i, r in enumerate(deltas[:topn], 1):
                w = abs(r["delta_abs"]) / vmax_d * 100
                pos = r["delta_abs"] >= 0
                grad = ("linear-gradient(90deg,#1f8f7f,#3dd6c4)" if pos
                        else "linear-gradient(90deg,#a83246,#ff7b8a)")
                tag = ("<span class='tag up'>↑</span>" if pos else "<span class='tag down'>↓</span>")
                sinal = "+" if pos else ""
                html_d.append(
                    f"<div class='rowit'>"
                    f"<div class='rank'>{i}</div>"
                    f"<div class='name' title=\"{r['indicador']}\">{r['indicador']}{tag}</div>"
                    f"<div class='barwrap'><div class='bar' style='width:{w:.1f}%;background:{grad}'></div></div>"
                    f"<div class='pct' style='width:90px'>{sinal}{r['delta_abs']:,.2f}</div>"
                    f"<div class='pct' style='width:72px;color:{'#48e0b6' if pos else '#ff7b8a'}'>"
                    f"{sinal}{r['delta_pct']:.1f}%</div>"
                    f"</div>")
            st.markdown("".join(html_d), unsafe_allow_html=True)

            tab_d = pd.DataFrame(deltas)[["indicador", "valor_anterior", "valor_atual", "delta_abs", "delta_pct"]]
            tab_d.columns = ["Indicador", f"Valor {data_ant.strftime('%d/%m/%Y')}",
                             f"Valor {data_sel.strftime('%d/%m/%Y')}", "Δ Absoluto", "Δ (%)"]
            with st.expander("📋 Ver tabela completa de variações / exportar"):
                st.dataframe(tab_d, use_container_width=True, height=420)
                st.download_button("⬇️ Baixar variações (CSV)",
                                   tab_d.to_csv(index=False).encode("utf-8"),
                                   f"variacao_{data_ant}_{data_sel}.csv", "text/csv")

            # ── Ranking de impacto na recuperação por variação dia anterior ──────
            st.markdown("---")
            st.markdown("### 📊 Ranking de Impacto na Recuperação Industrial pelas Variações")
            st.markdown(
                f"<p class='legend'>Para cada indicador, aplica-se isoladamente sua variação "
                f"(<b>{data_ant.strftime('%d/%m/%Y')}</b> → <b>{data_sel.strftime('%d/%m/%Y')}</b>) "
                f"mantendo os demais na base do dia anterior e mede-se o impacto em pontos percentuais "
                f"na Recuperação Industrial.</p>",
                unsafe_allow_html=True)

            y_ant_base = recuperacao_industrial(entradas_ant)
            y_atual_total = recuperacao_industrial(entradas)
            delta_recup_total = y_atual_total - y_ant_base

            _EXCLUIR_IMPACTO = {
                'Etanol Anidro', 'Etanol Hidratado', 'Etanol Hidratado Industrial',
                'Etanol Anidro Europeu', 'Etanol de Segunda',
                'Teor Alcoólico Etanol Anidro', 'Teor Alcoólico Etanol Hidratado ',
                'Teor Alcoólico Etanol Hidratado Industrial', 'Teor Alcoólico Etanol Anidro Europeu',
                'Teor Alcoólico Etanol de 2º',
                'Levedura Inativa Seca 300 g/kg', 'Levedura Inativa Seca 350 g/kg',
                'Levedura Inativa Seca 370 g/kg', 'Levedura Inativa Seca 390 g/kg',
                'Levedura Inativa Seca 400 g/kg',
                'PC', 'AR Cana (%)', 'Caldo da Cana - Pol', 'Óleo Fúsel',
                'CANA TOTAL', 'Açúcar Tipo VHP',
                'AÇÚCAR EM PROCESSO TOTAL (Dia Anterior)', 'ETANOL EM PROCESSO TOTAL (Dia Anterior)',
            }

            impactos_raw = []
            for row in INPUT_ROWS:
                nome = rotulo(row)
                if nome in _EXCLUIR_IMPACTO:
                    continue
                v_atual_i = entradas.get(row, 0.0)
                v_ant_i   = entradas_ant.get(row, 0.0)
                if v_atual_i == v_ant_i:
                    continue
                cenario = dict(entradas_ant)
                cenario[row] = v_atual_i
                y_mixed = recuperacao_industrial(cenario)
                impacto_pp = y_mixed - y_ant_base
                delta_i = v_atual_i - v_ant_i
                delta_pct_i = delta_i / abs(v_ant_i) * 100 if v_ant_i != 0 else 0.0
                impactos_raw.append({
                    "linha": row,
                    "indicador": nome,
                    "valor_anterior": v_ant_i,
                    "valor_atual": v_atual_i,
                    "delta_abs": delta_i,
                    "delta_pct": delta_pct_i,
                    "impacto_pp": impacto_pp,
                })

            seen_imp = {}
            for r in impactos_raw:
                nome = r["indicador"]
                if nome not in seen_imp or abs(r["impacto_pp"]) > abs(seen_imp[nome]["impacto_pp"]):
                    seen_imp[nome] = r
            impactos = list(seen_imp.values())
            impactos.sort(key=lambda r: abs(r["impacto_pp"]), reverse=True)

            col_kpi1, col_kpi2, col_kpi3 = st.columns(3)
            with col_kpi1:
                if impactos:
                    top1 = impactos[0]
                    if top1["indicador"] == "Quantidade Total Mel Final Tanque Nº5":
                        delta_abs_top = entradas.get(176, 0.0) - entradas_ant.get(176, 0.0)
                        lbl_variacao = "Volume Mel Final Tanque Nº5"
                    else:
                        delta_abs_top = top1["delta_abs"]
                        lbl_variacao = top1["indicador"]
                    sinal_tot = "+" if delta_abs_top >= 0 else ""
                    cor_tot = "#48e0b6" if delta_abs_top >= 0 else "#ff7b8a"
                    delta_abs_fmt = f"{delta_abs_top:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    st.markdown(
                        f"<div class='kpi'><div class='lbl'>Variação do indicador de maior impacto</div>"
                        f"<div class='val' style='font-size:2.2rem;background:none;"
                        f"-webkit-text-fill-color:{cor_tot}'>{sinal_tot}{delta_abs_fmt}</div>"
                        f"<div class='sub'>{lbl_variacao}<br>{data_ant.strftime('%d/%m/%Y')} → {data_sel.strftime('%d/%m/%Y')}</div></div>",
                        unsafe_allow_html=True)
            with col_kpi2:
                if impactos:
                    top_imp = impactos[0]
                    sinal_top = "+" if top_imp["impacto_pp"] >= 0 else ""
                    top_val_fmt = f"{top_imp['impacto_pp']:.2f}".replace(".", ",")
                    st.markdown(
                        f"<div class='kpi'><div class='lbl'>Maior impacto isolado</div>"
                        f"<div class='val' style='font-size:1.3rem'>{top_imp['indicador']}</div>"
                        f"<div class='sub'>{sinal_top}{top_val_fmt}% na recuperação</div></div>",
                        unsafe_allow_html=True)
            with col_kpi3:
                soma_top = sum(r["impacto_pp"] for r in impactos[:topn])
                rec_ajustada = y_atual_total + soma_top
                cor_aj = "#48e0b6" if rec_ajustada >= y_atual_total else "#ff7b8a"
                rec_aj_fmt = f"{rec_ajustada:.2f}".replace(".", ",")
                ri_fmt = f"{y_atual_total:.2f}".replace(".", ",")
                sinal_soma = "+" if soma_top >= 0 else ""
                soma_fmt = f"{soma_top:.2f}".replace(".", ",")
                st.markdown(
                    f"<div class='kpi'>"
                    f"<div class='lbl'>Recuperação ajustada</div>"
                    f"<div class='val' style='font-size:2.2rem;background:none;"
                    f"-webkit-text-fill-color:{cor_aj}'>{rec_aj_fmt}%</div>"
                    f"<div class='sub'>{ri_fmt}% base {sinal_soma}{soma_fmt}% (Δ top {topn})</div>"
                    f"</div>",
                    unsafe_allow_html=True)

            if impactos:
                vmax_imp = max((abs(r["impacto_pp"]) for r in impactos[:topn]), default=1) or 1
                html_imp = []
                for i, r in enumerate(impactos[:topn], 1):
                    w = abs(r["impacto_pp"]) / vmax_imp * 100
                    pos = r["impacto_pp"] >= 0
                    grad = ("linear-gradient(90deg,#1f8f7f,#3dd6c4)" if pos
                            else "linear-gradient(90deg,#a83246,#ff7b8a)")
                    tag = ("<span class='tag up'>↑</span>" if pos
                           else "<span class='tag down'>↓</span>")
                    sinal = "+" if pos else ""
                    cor_val = "#48e0b6" if pos else "#ff7b8a"
                    val_fmt = f"{r['impacto_pp']:.2f}".replace(".", ",")
                    html_imp.append(
                        f"<div class='rowit'>"
                        f"<div class='rank'>{i}</div>"
                        f"<div class='name' title=\"{r['indicador']}\">{r['indicador']}{tag}</div>"
                        f"<div class='barwrap'><div class='bar' style='width:{w:.1f}%;background:{grad}'></div></div>"
                        f"<div class='pct' style='width:120px;color:{cor_val}'>{sinal}{val_fmt}%</div>"
                        f"</div>")
                st.markdown("".join(html_imp), unsafe_allow_html=True)

                tab_imp = pd.DataFrame(impactos)[[
                    "indicador", "valor_anterior", "valor_atual",
                    "delta_abs", "delta_pct", "impacto_pp"]]
                tab_imp.columns = [
                    "Indicador",
                    f"Valor {data_ant.strftime('%d/%m/%Y')}",
                    f"Valor {data_sel.strftime('%d/%m/%Y')}",
                    "Δ Indicador (abs)", "Δ Indicador (%)",
                    "Impacto na Recuperação (%)"]
                with st.expander("📋 Ver tabela completa de impactos / exportar"):
                    st.dataframe(tab_imp, use_container_width=True, height=420)
                    st.download_button(
                        "⬇️ Baixar ranking de impactos (CSV)",
                        tab_imp.to_csv(index=False).encode("utf-8"),
                        f"impacto_recuperacao_{data_ant}_{data_sel}.csv",
                        "text/csv")
            else:
                st.info("Nenhuma variação encontrada entre as duas datas para calcular impactos.")