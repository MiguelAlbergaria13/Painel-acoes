# python -m pip install yfinance openpyxl pywin32
# Painel de ações europeias e americanas: preço atual, preço-alvo médio dos analistas, ranking pelo potencial,
# máximos/mínimos e distância ao máximo da última semana, mês, 3 meses, ano e 3 anos. Fonte: Yahoo Finance.
#
# Uso: duplo clique em "Atualizar_painel.bat" (ou no botão ▶ ATUALIZAR dentro do Painel_acoes.xlsx,
# que fica ativo depois de o script correr uma vez neste PC).
# Em cada folha: "Ordenar por" (lista na linha 3) + ↕ ORDENAR reordena logo, sem descarregar nada; a ordem escolhida
# também se aplica em cada ▶ ATUALIZAR.
# Pontuação 1-5 (médias 100/200, valor justo e analistas), preços em euros (bolsa europeia) e 4 níveis por título (compra razoável, compra máxima,
# venda razoável, venda mínima). Folha Procurar: procura títulos no Yahoo (nome/ticker, ou os mais significativos
# de uma região) e ➕ ACRESCENTAR MARCADOS junta-os à folha Acrescentados; entram no painel na próxima ▶ ATUALIZAR.
# O script atualiza o _Painel_acoes.xlsx (mesmo nome que o .py) na mesma pasta e abre-o. Se estiver aberto no Excel, fecha-o primeiro sozinho.
import json
import shutil
import os
import re
import subprocess
import sys
import tempfile
import time
import traceback
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

PASTA = os.path.dirname(os.path.abspath(__file__))
# o Excel tem o mesmo nome que este script (_Painel_acoes.py -> _Painel_acoes.xlsx); o .bat é o _Atualizar_painel.bat da pasta
BASE = os.path.splitext(os.path.basename(os.path.abspath(__file__)))[0]
FICHEIRO = os.path.join(PASTA, BASE + ".xlsx")


def _bat():
    """O .bat do painel: de preferência o que começa como este script (com "_" se o script tiver "_")."""
    import glob
    import re
    pre = "_" if BASE.startswith("_") else ""
    r = [g for g in glob.glob(os.path.join(PASTA, "*.bat")) if re.search(r"atualizar.*painel", os.path.basename(g), re.I)]
    r.sort(key=lambda g: (not os.path.basename(g).startswith(pre) if pre else os.path.basename(g).startswith("_"), g))
    return r[0] if r else os.path.join(PASTA, pre + "Atualizar_painel.bat")


BAT = _bat()

# (nome habitual, código habitual, ticker no Yahoo, bolsa) — acrescenta ou tira ações à vontade
ACOES = [
    ("EDP", "EDP", "EDP.LS", "Euronext Lisboa"),
    ("EDP Renováveis", "EDPR", "EDPR.LS", "Euronext Lisboa"),
    ("Galp Energia", "GALP", "GALP.LS", "Euronext Lisboa"),
    ("Jerónimo Martins", "JMT", "JMT.LS", "Euronext Lisboa"),
    ("Millennium BCP", "BCP", "BCP.LS", "Euronext Lisboa"),
    ("NOS", "NOS", "NOS.LS", "Euronext Lisboa"),
    ("The Navigator Company", "NVG", "NVG.LS", "Euronext Lisboa"),
    ("Sonae", "SON", "SON.LS", "Euronext Lisboa"),
    ("ASML", "ASML", "ASML.AS", "Euronext Amesterdão"),
    ("ING Groep", "INGA", "INGA.AS", "Euronext Amesterdão"),
    ("Prosus", "PRX", "PRX.AS", "Euronext Amesterdão"),
    ("Adyen", "ADYEN", "ADYEN.AS", "Euronext Amesterdão"),
    ("Ahold Delhaize", "AD", "AD.AS", "Euronext Amesterdão"),
    ("Heineken", "HEIA", "HEIA.AS", "Euronext Amesterdão"),
    ("Philips", "PHIA", "PHIA.AS", "Euronext Amesterdão"),
    ("Wolters Kluwer", "WKL", "WKL.AS", "Euronext Amesterdão"),
    ("Anheuser-Busch InBev", "ABI", "ABI.BR", "Euronext Bruxelas"),
    ("LVMH", "MC", "MC.PA", "Euronext Paris"),
    ("TotalEnergies", "TTE", "TTE.PA", "Euronext Paris"),
    ("Hermès", "RMS", "RMS.PA", "Euronext Paris"),
    ("L'Oréal", "OR", "OR.PA", "Euronext Paris"),
    ("Schneider Electric", "SU", "SU.PA", "Euronext Paris"),
    ("Sanofi", "SAN", "SAN.PA", "Euronext Paris"),
    ("Airbus", "AIR", "AIR.PA", "Euronext Paris"),
    ("Air Liquide", "AI", "AI.PA", "Euronext Paris"),
    ("BNP Paribas", "BNP", "BNP.PA", "Euronext Paris"),
    ("Safran", "SAF", "SAF.PA", "Euronext Paris"),
    ("EssilorLuxottica", "EL", "EL.PA", "Euronext Paris"),
    ("AXA", "CS", "CS.PA", "Euronext Paris"),
    ("Vinci", "DG", "DG.PA", "Euronext Paris"),
    ("Danone", "BN", "BN.PA", "Euronext Paris"),
    ("Kering", "KER", "KER.PA", "Euronext Paris"),
    ("Pernod Ricard", "RI", "RI.PA", "Euronext Paris"),
    ("Saint-Gobain", "SGO", "SGO.PA", "Euronext Paris"),
    ("SAP", "SAP", "SAP.DE", "Xetra"),
    ("Siemens", "SIE", "SIE.DE", "Xetra"),
    ("Allianz", "ALV", "ALV.DE", "Xetra"),
    ("Deutsche Telekom", "DTE", "DTE.DE", "Xetra"),
    ("Munich Re", "MUV2", "MUV2.DE", "Xetra"),
    ("Mercedes-Benz Group", "MBG", "MBG.DE", "Xetra"),
    ("BMW", "BMW", "BMW.DE", "Xetra"),
    ("Volkswagen (pref.)", "VOW3", "VOW3.DE", "Xetra"),
    ("BASF", "BAS", "BAS.DE", "Xetra"),
    ("Bayer", "BAYN", "BAYN.DE", "Xetra"),
    ("Deutsche Börse", "DB1", "DB1.DE", "Xetra"),
    ("Infineon", "IFX", "IFX.DE", "Xetra"),
    ("Adidas", "ADS", "ADS.DE", "Xetra"),
    ("Deutsche Bank", "DBK", "DBK.DE", "Xetra"),
    ("Rheinmetall", "RHM", "RHM.DE", "Xetra"),
    ("Siemens Energy", "ENR", "ENR.DE", "Xetra"),
    ("Banco Santander", "SAN", "SAN.MC", "Bolsa de Madrid"),
    ("BBVA", "BBVA", "BBVA.MC", "Bolsa de Madrid"),
    ("Iberdrola", "IBE", "IBE.MC", "Bolsa de Madrid"),
    ("Inditex", "ITX", "ITX.MC", "Bolsa de Madrid"),
    ("Telefónica", "TEF", "TEF.MC", "Bolsa de Madrid"),
    ("UniCredit", "UCG", "UCG.MI", "Borsa Italiana"),
    ("Intesa Sanpaolo", "ISP", "ISP.MI", "Borsa Italiana"),
    ("Enel", "ENEL", "ENEL.MI", "Borsa Italiana"),
    ("Eni", "ENI", "ENI.MI", "Borsa Italiana"),
    ("Ferrari", "RACE", "RACE.MI", "Borsa Italiana"),
    ("Stellantis", "STLAM", "STLAM.MI", "Borsa Italiana"),
    ("Generali", "G", "G.MI", "Borsa Italiana"),
    ("Shell", "SHEL", "SHEL.L", "Londres"),
    ("AstraZeneca", "AZN", "AZN.L", "Londres"),
    ("HSBC", "HSBA", "HSBA.L", "Londres"),
    ("Unilever", "ULVR", "ULVR.L", "Londres"),
    ("BP", "BP", "BP.L", "Londres"),
    ("Rio Tinto", "RIO", "RIO.L", "Londres"),
    ("GSK", "GSK", "GSK.L", "Londres"),
    ("Barclays", "BARC", "BARC.L", "Londres"),
    ("Lloyds Banking Group", "LLOY", "LLOY.L", "Londres"),
    ("Rolls-Royce", "RR", "RR.L", "Londres"),
    ("BAE Systems", "BA", "BA.L", "Londres"),
    ("Diageo", "DGE", "DGE.L", "Londres"),
    ("British American Tobacco", "BATS", "BATS.L", "Londres"),
    ("RELX", "REL", "REL.L", "Londres"),
    ("London Stock Exchange Group", "LSEG", "LSEG.L", "Londres"),
    ("Glencore", "GLEN", "GLEN.L", "Londres"),
    ("Nestlé", "NESN", "NESN.SW", "SIX Suíça"),
    ("Roche", "ROP", "ROP.SW", "SIX Suíça"),   # desde 17/03/2026 (antes ROG)
    ("Novartis", "NOVN", "NOVN.SW", "SIX Suíça"),
    ("UBS", "UBSG", "UBSG.SW", "SIX Suíça"),
    ("ABB", "ABBN", "ABBN.SW", "SIX Suíça"),
    ("Zurich Insurance", "ZURN", "ZURN.SW", "SIX Suíça"),
    ("Richemont", "CFR", "CFR.SW", "SIX Suíça"),
    ("Novo Nordisk", "NOVO B", "NOVO-B.CO", "Copenhaga"),
    ("Equinor", "EQNR", "EQNR.OL", "Oslo"),
    ("Volvo", "VOLV B", "VOLV-B.ST", "Estocolmo"),
    ("Atlas Copco", "ATCO A", "ATCO-A.ST", "Estocolmo"),
    ("Ericsson", "ERIC B", "ERIC-B.ST", "Estocolmo"),
    ("Nokia", "NOKIA", "NOKIA.HE", "Helsínquia"),
    ("Nordea", "NDA FI", "NDA-FI.HE", "Helsínquia"),
    ("Apple", "AAPL", "AAPL", "Nasdaq"),
    ("Microsoft", "MSFT", "MSFT", "Nasdaq"),
    ("NVIDIA", "NVDA", "NVDA", "Nasdaq"),
    ("Amazon", "AMZN", "AMZN", "Nasdaq"),
    ("Alphabet (Google)", "GOOGL", "GOOGL", "Nasdaq"),
    ("Meta Platforms", "META", "META", "Nasdaq"),
    ("Tesla", "TSLA", "TSLA", "Nasdaq"),
    ("Broadcom", "AVGO", "AVGO", "Nasdaq"),
    ("Berkshire Hathaway (B)", "BRK.B", "BRK-B", "NYSE"),
    ("JPMorgan Chase", "JPM", "JPM", "NYSE"),
    ("Visa", "V", "V", "NYSE"),
    ("Mastercard", "MA", "MA", "NYSE"),
    ("Eli Lilly", "LLY", "LLY", "NYSE"),
    ("UnitedHealth", "UNH", "UNH", "NYSE"),
    ("Exxon Mobil", "XOM", "XOM", "NYSE"),
    ("Chevron", "CVX", "CVX", "NYSE"),
    ("Johnson & Johnson", "JNJ", "JNJ", "NYSE"),
    ("Procter & Gamble", "PG", "PG", "NYSE"),
    ("Walmart", "WMT", "WMT", "NYSE"),
    ("Costco", "COST", "COST", "Nasdaq"),
    ("Home Depot", "HD", "HD", "NYSE"),
    ("Oracle", "ORCL", "ORCL", "NYSE"),
    ("Netflix", "NFLX", "NFLX", "Nasdaq"),
    ("AMD", "AMD", "AMD", "Nasdaq"),
    ("Bank of America", "BAC", "BAC", "NYSE"),
    ("Coca-Cola", "KO", "KO", "NYSE"),
    ("PepsiCo", "PEP", "PEP", "Nasdaq"),
    ("Merck & Co.", "MRK", "MRK", "NYSE"),
    ("AbbVie", "ABBV", "ABBV", "NYSE"),
    ("Salesforce", "CRM", "CRM", "NYSE"),
    ("Adobe", "ADBE", "ADBE", "Nasdaq"),
    ("Cisco", "CSCO", "CSCO", "Nasdaq"),
    ("Intel", "INTC", "INTC", "Nasdaq"),
    ("Qualcomm", "QCOM", "QCOM", "Nasdaq"),
    ("Palantir", "PLTR", "PLTR", "Nasdaq"),
    ("Uber", "UBER", "UBER", "NYSE"),
    ("Boeing", "BA", "BA", "NYSE"),
    ("Walt Disney", "DIS", "DIS", "NYSE"),
    ("Nike", "NKE", "NKE", "NYSE"),
    ("McDonald's", "MCD", "MCD", "NYSE"),
    ("IBM", "IBM", "IBM", "NYSE"),
    ("Pfizer", "PFE", "PFE", "NYSE"),
    ("Micron", "MU", "MU", "Nasdaq"),
    ("Goldman Sachs", "GS", "GS", "NYSE"),
    ("Caterpillar", "CAT", "CAT", "NYSE"),
    ("Wells Fargo", "WFC", "WFC", "NYSE"),
    ("Taiwan Semiconductor (ADR)", "TSM", "TSM", "NYSE"),
]

# ETFs UCITS (os que se compram na Europa, p. ex. na Trading 212) sobre os mesmos mercados.
# Portugal não tem hoje nenhum ETF do PSI. Mesmo formato da lista de ações: acrescenta ou tira à vontade.
ETFS = [
    # ações (um por mercado, sem repetidos; de preferência a cotar em euros)
    ("iShares STOXX Europe 600 (DE)", "EXSA", "EXSA.DE", "Xetra"),
    ("iShares Core EURO STOXX 50 (DE)", "EXW1", "EXW1.DE", "Xetra"),
    ("VanEck AEX (Países Baixos)", "TDT", "TDT.AS", "Euronext Amesterdão"),
    ("Amundi CAC 40 (França)", "CAC", "CAC.PA", "Euronext Paris"),
    ("iShares Core DAX (Alemanha)", "EXS1", "EXS1.DE", "Xetra"),
    ("Amundi IBEX 35 (Espanha)", "LYXIB", "LYXIB.MC", "Bolsa de Madrid"),
    ("iShares FTSE MIB (Itália)", "IMIB", "IMIB.MI", "Borsa Italiana"),
    ("iShares Core FTSE 100 (Reino Unido)", "ISF", "ISF.L", "Londres"),
    ("iShares SMI (Suíça)", "CSSMI", "CSSMI.SW", "SIX Suíça"),
    ("XACT OMXS30 (Suécia)", "XACT OMXS30", "XACT-OMXS30.ST", "Estocolmo"),
    ("iShares Core S&P 500 (Acc)", "SXR8", "SXR8.DE", "Xetra"),
    ("iShares Nasdaq 100 (Acc)", "SXRV", "SXRV.DE", "Xetra"),
    ("iShares Core MSCI EM IMI (emergentes)", "EMIM", "EMIM.AS", "Euronext Amesterdão"),
    ("Vanguard FTSE All-World (Acc)", "VWCE", "VWCE.DE", "Xetra"),
    # não são ações: nas crises a pontuação passa para os que estão a subir (testado: com obrigações e ouro na lista a
    # queda máxima ficou em −21% em vez de −34%)
    ("iShares Core Euro Government Bond (obrigações do Estado zona euro)", "IEGA", "IEGA.AS", "Euronext Amesterdão"),
    ("iShares $ Treasury Bond 20+yr (obrigações EUA longo prazo)", "IDTL", "IDTL.L", "Londres"),
    ("Xetra-Gold (ouro)", "4GLD", "4GLD.DE", "Xetra"),
]

# períodos analisados: (nome curto, dias de calendário)
PERIODOS = [("semana", 7), ("mês", 30), ("3 meses", 91), ("ano", 365), ("3 anos", 1095)]
NP = len(PERIODOS)

# Preços de compra (σd = σ diário = σ mensal / √21; tendência = subida dos últimos 20 dias em σ, entre −2 e +2),
# SEMPRE abaixo do preço atual (uma ordem limite acima do preço é executada logo):
#   compra máxima (a ordem limite principal) = preço × exp(σd × (−0,75 + 0,25 × tendência)): ~0,75 σd abaixo (≈ −1,7%);
#     em subida franca só ~0,25 σd abaixo (senão nunca é executada), em queda franca ~1,25 σd abaixo;
#   compra razoável (trigger / ordem de oportunidade) = preço × exp(−1 σ mensal) (≈ −10%): só numa correção a sério.
# Teste em 38 títulos, 2016–2026, dias com sinal de compra: −0,75 σd ajustado é executado ~58% das vezes em 3 dias e
# ~71% em 7 (o anterior, −1 σd: 49% e 63%); quando não é executado, comprar logo ao preço do momento;
# −1 σ mensal só ~6% em 5 dias, ~25% em 20 e ~45% em 60 (quando não é executado, o título sobe em média ~4% num mês).
TETO_BASE, TETO_TEND = -0.75, 0.25
RAZOAVEL_SIGMAS = 1.0

# Colunas da tabela, por ordem: (chave, cabeçalho, largura, formato, grupo)
#   formato: "txt", "preco" ou "pct";  grupo: decide a cor do cabeçalho
COLS = [("sel", "Atualizar (x)", 3, "txt", "sel"), ("det", "Detido (x)", 4, "txt", "det"),
        ("qtd", "Nº de ações detidas", 7, "qtd", "det"), ("pcompra", "Preço médio de compra (por ação)", 8, "preco", "det"),
        ("valor_eur", "Valor detido (€)", 12, "eur", "det"), ("lp_eur", "Lucro / prejuízo (€)", 8, "eur", "det"),
        ("lp_pct", "Lucro / prejuízo (%)", 7, "pct", "det"), ("corr", "Corretora", 8, "txt", "det"),
        ("nome", "Empresa / ETF", 28, "txt", "id"),
        ("tipo", "Tipo", 5, "txt", "id"), ("codigo", "Código", 7, "txt", "id"),
        ("ticker", "Ticker Yahoo", 10, "txt", "id"), ("bolsa", "Bolsa", 18, "txt", "id"),
        ("moeda", "Moeda", 7, "txt", "id"),
        ("horario", "Horário da bolsa (hora de Lisboa)", 11, "txt", "id"),
        ("preco_hora", "Preço de (sessão, hora de Lisboa)", 16, "txt", "id"), ("preco", "Preço atual", 12, "preco", "id"),
        ("desde_ref", "Da data de referência até hoje (%)", 11, "pct", "teste"),
        ("nota_geral", "Nota global: nota geral + nível de preço (2 : 1, folha Ponderações)", 9, "nota", "global"),
        ("decisao", "Decisão (pela nota global: nota geral + preço)", 14, "txt", "global"),
        ("nota", "Nota geral: negócio, tendência e momento, sem o preço (pesos na folha Ponderações)", 9, "nota", "global"),
        ("nivel_preco", "Nível de preço (5 muito barata … 1 muito cara)", 19, "nivel", "global"),
        ("beta", "Beta: oscilação face ao mercado (S&P 500; 1 = igual, >1 amplifica, <1 amortece)", 6, "num1", "global"),
        ("pont", "Pontuação do sinal (5 = compra forte … 1 = venda)", 10, "pont", "conselho"),
        ("etf_acoes", "Nota ponderada das ações do ETF", 9, "nota", "global"),
        ("etf_desc", "Descrição do ETF", 40, "txt", "id"),
        ("etf_comp", "Composição: maiores posições (peso; nota geral)", 50, "txt", "id"),
        ("qual", "Qualidade do negócio (0-5)", 8, "nota", "conselho"),
        ("saude_n", "Saúde financeira (0-5)", 8, "nota", "conselho"),
        ("sinal", "Médias 100/200: compra / venda próxima ou por inversão", 18, "txt", "conselho"),
        ("saude", "Saúde financeira: detalhe (ALERTA = nota geral no máx. 2,5)", 16, "txt", "conselho"),
        ("lp_txt", "Longo prazo: subida média por ano (desde sempre; 10 anos; 5 anos; 3 anos; 1 ano)", 15, "txt", "conselho"),
        ("div", "Dividendo anual por ação", 7, "preco", "fund"),
        ("div_yield", "Dividend yield (%)", 7, "pct", "fund"),
        ("payout", "Payout ratio (% do lucro)", 7, "pct", "fund"),
        ("tend", "Médias 100/200: tendência desde", 13, "txt", "conselho"),
        ("pe", "P/E (preço / lucro dos últimos 12 meses)", 10, "num1", "conselho"),
        ("compra", "Compra razoável (trigger, ≈ −1σ ≈ −10%)", 13, "preco", "sugerido"),
        ("compra_pct", "Compra razoável / atual (%)", 10, "pct", "sugerido"),
        ("teto", "Compra máxima (ordem limite)", 13, "preco", "sugerido"),
        ("teto_pct", "Compra máxima / atual (%)", 10, "pct", "sugerido"),
        ("venda_alvo", "Venda razoável (objetivo de ganho)", 13, "preco", "venda"),
        ("venda_alvo_pct", "Venda razoável / atual (%)", 10, "pct", "venda"),
        ("venda_min", "Venda mínima aconselhável (rede de segurança)", 13, "preco", "venda"),
        ("venda_min_pct", "Venda mínima / atual (%)", 10, "pct", "venda"),
        ("venda_min_compra", "Venda mínima face à tua compra (%): o que ganhas (+) ou perdes (−) se vender na rede", 11, "pct", "venda"),
        ("coment", "Comentários (compra / venda próxima, o que falta para subir de nível)", 55, "txt", "conselho"),
        ("mm100c", "Média 100 dias", 12, "preco", "conselho"),
        ("gap5", "Média 100 / média 200 há 5 dias (%)", 11, "pct", "conselho"),
        ("mm200c", "Média 200 dias", 13, "preco", "conselho"),
        ("justo", "Valor justo (lucros × P/E habitual)", 14, "preco", "justo"),
        ("justo_pct", "Valor justo / atual (%)", 12, "pct", "justo"),
        ("alvo", "Preço-alvo médio (analistas)", 15, "preco", "id"),
        ("pot", "Alvo / atual (%)", 12, "pct", "id"),
        ("rec", "Recomendação", 13, "txt", "id"),
        ("var1d", "Variação último dia (%)", 11, "pct", "mov"), ("var1m", "Variação último mês (%)", 13, "pct", "mov"),
        ("valor_tx", "Valor transacionado/dia, média 3 meses (M €)", 14, "num1", "vol"),
        ("volume", "Volume/dia, média 3 meses (nº de títulos)", 13, "int", "vol"),
        ("cap", "Capitalização (mil M €)", 12, "num1", "vol"),
        ("sigma", "Volatilidade mensal σ (%)", 11, "pct", "teste"),
        ("t_m200", "Último ano: seguir o sinal (%)", 12, "pct", "teste"),
        ("t_hold", "Último ano: manter (%)", 13, "pct", "teste")]
for i, (p, _) in enumerate(PERIODOS):   # posição atual face ao máximo + oscilação em torno da média
    COLS += [(f"dist{i}", f"Distância ao máx. {p} (%)", 13, "pct", f"osc{i}"),
             (f"cima{i}", f"Oscilação acima da média {p} (%)", 13, "pct", f"osc{i}"),
             (f"baixo{i}", f"Oscilação abaixo da média {p} (%)", 13, "pct", f"osc{i}")]
COLS += [("nanal", "Nº de analistas", 10, "txt", "id")]
for i, (p, _) in enumerate(PERIODOS):
    COLS += [(f"max{i}", f"Máx. {p}", 13, "preco", f"px{i}"), (f"min{i}", f"Mín. {p}", 13, "preco", f"px{i}"),
             (f"med{i}", f"Média {p}", 13, "preco", f"px{i}")]
COLS += [("notas", "Notas", 30, "txt", "id")]
COLS += [("ppo", "PPO 12/26 (%)", 10, "pct", "ind"), ("ppo_sinal", "PPO: linha de sinal 9 (%)", 13, "pct", "ind"),
         ("ppo_hist", "PPO: histograma (%)", 13, "pct", "ind"),
         ("ppo_dh", "Histograma: variação em 5 dias (%)", 11, "pct", "ind"),
         ("rsi", "RSI 14 dias", 8, "num1", "ind"), ("mm200", "Média 200 dias (indicador)", 11, "preco", "ind"),
         ("dist200", "Preço / média 200 dias (%)", 13, "pct", "ind"),
         ("mom12", "Momentum 12 meses, sem o último (%)", 12, "pct", "ind"),
         ("tend20", "Tendência 20 dias (subida em σ, −2 a +2)", 11, "num1", "ind"),
         ("lp", "Longo prazo: subida média por ano (%)", 11, "pct", "ind"),
         ("lp10", "Longo prazo: subida média por ano nos últimos 10 anos (%)", 11, "pct", "ind"),
         ("so_caro", "Cara (1): só aviso, decisão a azul", 5, "geral", "ind"),
         ("pe_fwd", "P/E futuro (lucros previstos para os próximos 12 meses)", 8, "num1", "ind"),
         ("pe_hab", "PER habitual do próprio título (mediana de 3 anos, usado no valor justo)", 8, "num1", "ind"),
         ("etf_caro", "ETF: parte das ações conhecidas que estão caras (%)", 8, "pct", "ind"),
         ("nota_sc", "Nota geral sem o travão de caro", 6, "nota", "ind"),
         ("lp5", "Longo prazo: subida média por ano nos últimos 5 anos (%)", 11, "pct", "ind"),
         ("val_sempre", "Valorização média por ano desde sempre (%)", 9, "pct", "ind"),
         ("val10", "Valorização média por ano em 10 anos (%)", 9, "pct", "ind"),
         ("val5", "Valorização média por ano em 5 anos (%)", 9, "pct", "ind"),
         ("val3", "Valorização média por ano em 3 anos (%)", 9, "pct", "ind"),
         ("val1", "Valorização em 1 ano (%)", 9, "pct", "ind"),
         ("inc1a", "Inclinação da reta de 1 ano, cotações diárias (% por ano)", 9, "pct", "ind"),
         ("r2_1a", "Regularidade da reta de 1 ano (R², 0 a 1)", 7, "num1", "ind"),
         ("inc6m", "Inclinação da reta de 6 meses, cotações diárias (% por ano)", 9, "pct", "ind"),
         ("r2_6m", "Regularidade da reta de 6 meses (R², 0 a 1)", 7, "num1", "ind"),
         ("mm50", "Média 50 dias (entrada: o preço acima dela = estabilizou)", 10, "preco", "ind"),
         ("arranque", "Início de subida depois de 120 dias parada (1 = sim)", 7, "int", "ind"),
         ("sub60", "Subida forte e regular nos últimos 60 dias (1 = sim)", 7, "int", "ind"),
         ("max1a", "Máximo de 1 ano (para a recuperação)", 10, "preco", "ind"),
         ("inc60", "Reta dos últimos 60 dias (subida por ano)", 8, "pct", "ind"),
         ("rec_a", "Dias de bolsa desde o último sinal de recuperação (regra das ações)", 7, "int", "ind"),
         ("rec_e", "Dias de bolsa desde o último sinal de recuperação (regra dos ETFs)", 7, "int", "ind"),
         ("vol1a", "Volatilidade do último ano (desvio-padrão anual das variações diárias) (%)", 9, "pct", "ind"),
         ("rec_media", "Recomendação média dos analistas (1 compra forte … 5 venda)", 7, "num1", "ind"),
         ("fator_tend", "Fator de tendência (multiplica a nota geral)", 7, "num1", "ind"),
         ("mom6", "Subida 6 meses (%)", 11, "pct", "ind"), ("ref6m", "Preço há 6 meses", 11, "preco", "ind"),
         ("roic", "ROIC (EBIT após impostos / capital investido) (%)", 12, "pct", "fund"),
         ("gpa", "Lucro bruto / ativo total (%)", 11, "pct", "fund"),
         ("fcf_lucro", "FCF / lucro líquido (×)", 10, "num1", "fund"),
         ("fcf_yield", "FCF yield (FCF / capitalização) (%)", 11, "pct", "fund"),
         ("ev_ebitda", "EV / EBITDA (×)", 10, "num1", "fund"),
         ("divl_ebitda", "Dívida líquida / EBITDA (×)", 11, "num1", "fund"),
         ("ebit_juros", "EBIT / juros (cobertura) (×)", 11, "num1", "fund")]
# preços: ao lado de cada um, o mesmo valor noutra moeda (letra 9, entre parênteses): em dólares se a cotação for em
# euros, em euros nas outras moedas
ALT_KEYS = ("preco", "compra", "teto", "venda_alvo", "venda_min", "mm100c", "mm200c", "justo", "alvo", "div")
NEGRITO = ("preco", "compra", "teto", "venda_alvo", "venda_min")
_cols = []
for _c in COLS:
    _cols.append(_c)
    if _c[0] in ALT_KEYS:
        _cols.append((_c[0] + "_alt", _c[1] + " (noutra moeda)", 8, "alt", _c[4]))
COLS = _cols + [("moeda_alt", "Moeda alternativa", 6, "txt", "ind"),
                ("fx_alt", "Câmbio: moeda da cotação -> alternativa", 9, "geral", "ind"),
                ("desat", "Não atualizado na última atualização parcial (1)", 5, "geral", "ind")]
OCULTAS = ("moeda_alt", "fx_alt", "desat", "corr", "lp10", "lp5", "val_sempre", "val10", "val5", "val3", "val1", "inc1a", "r2_1a", "inc6m", "r2_6m", "mm50", "arranque", "sub60", "max1a", "inc60", "rec_a", "rec_e", "vol1a", "rec_media", "fator_tend", "so_caro", "nota_sc", "pe_fwd", "etf_caro", "pe_hab",
           "compra", "compra_alt", "compra_pct")     # compra razoável (trigger): já não se usa; fica só a compra máxima   # corr: corretora (escondida por agora; os valores ficam)
# larguras mais estreitas e letra 9 em algumas colunas
_LARG = {"preco_hora": 10, "pont": 7, "etf_acoes": 7, "etf_desc": 24, "etf_comp": 30, "qual": 6, "saude_n": 6,
         "div": 6, "div_alt": 6, "div_yield": 6, "payout": 6, "sinal": 14, "saude": 14, "lp_txt": 14}
COLS = [(k, h, _LARG.get(k, w), f, g) for k, h, w, f, g in COLS]
# V190: a nota por ganhos (antes nas folhas «nota alternativa») entra nas folhas Ações / ETFs, logo a seguir ao preço
# noutra moeda (coluna R); a decisão passa para o lado dela
G_COLS = [("g_n5", "Nota por ganhos (0-5: 5 = média das 2 melhores)", 7, "nota", "ganhos"),
          ("g_nota", "Nota por ganhos (rent. ajustada, %/ano)", 9, "pct", "ganhos")]
G_COLS2 = [("g_pond", "Ganhos: rent. anual ponderada (%)", 8, "pct", "ganhos"),
           ("g_simples", "Ganhos: rent. anual simples (%)", 8, "pct", "ganhos"),
           ("g_curva", "Ganhos: rent. da curva, compras periódicas até hoje (%)", 8, "pct", "ganhos"),
           ("g_medida", "Ganhos: medida do preço (P/E, nível 1-5 ou × tendência)", 8, "num1", "ganhos"),
           ("g_vol", "Ganhos: volatilidade anual usada (%)", 8, "pct", "ganhos"),
           ("g_fpreco", "Ganhos: fator preço", 7, "num1", "ganhos"),
           ("g_fvol", "Ganhos: fator volatilidade", 7, "num1", "ganhos"),
           ("g_ref", "Ganhos: ponderada − referência (%)", 8, "pct", "ganhos"),
           ("g_real_t", "Ganhos: real desde a data de cálculo (total, %)", 8, "pct", "ganhos"),
           ("g_real_a", "Ganhos: real desde a data de cálculo (por ano, %)", 8, "pct", "ganhos"),
           ("g_obs", "Ganhos: observações", 30, "txt", "ganhos")]
_dec = next(c for c in COLS if c[0] == "decisao")
_i = [c[0] for c in COLS].index("preco_alt") + 1
COLS = COLS[:_i] + G_COLS + [_dec] + G_COLS2 + [c for c in COLS[_i:] if c[0] != "decisao"]
G_KEYS = [c[0] for c in G_COLS + G_COLS2]
CORRETORAS = ["T212", "TRep", "T212+TRep"]
FONTE9 = {"corr", "qtd", "pcompra", "valor_eur", "lp_eur", "lp_pct", "preco_hora", "sinal", "saude", "lp_txt"}
SIMBOLO = {"EUR": "€", "USD": "$", "GBP": "£", "GBp": "p", "CHF": "CHF", "JPY": "¥", "DKK": "DKK", "SEK": "SEK",
           "NOK": "NOK", "CAD": "C$", "AUD": "A$", "HKD": "HK$", "ARS": "ARS", "BRL": "R$", "TWD": "NT$", "KRW": "₩",
           "CNY": "CN¥", "INR": "₹", "PLN": "zł", "ZAc": "ZAc", "ILA": "ILA", "MXN": "MX$", "SGD": "S$"}
SEM_DECIMAIS = ("JPY", "KRW", "TWD")


def fmt_moeda(m, alt=False):
    """Formato do Excel com o símbolo à esquerda e o número à direita (p. ex. €   123.00); alt: entre parênteses."""
    sim = SIMBOLO.get(m, m or "")
    num = "#,##0" if m in SEM_DECIMAIS else "#,##0.00"
    return f'"({sim}"* {num}")"' if alt else f'"{sim}"* {num}'


def moeda_alt(m):
    return "USD" if m == "EUR" else "EUR"


def preencher_alt(d):
    """Valores noutra moeda (preço × câmbio guardado em fx_alt)."""
    fx = d.get("fx_alt") if isinstance(d.get("fx_alt"), (int, float)) else None
    for k in ALT_KEYS:
        v = d.get(k)
        d[k + "_alt"] = v * fx if (fx and isinstance(v, (int, float))) else None


def cambios_alt(linhas):
    """Câmbio de cada linha para a moeda alternativa (um pedido ao Yahoo para todas as moedas)."""
    moedas = {txt(d.get("moeda")) for d in linhas if txt(d.get("moeda"))}
    try:
        tx = taxas_eur(moedas | {"USD"})
    except Exception:
        tx = {}
    for d in linhas:
        m = txt(d.get("moeda"))
        if m:
            ma = moeda_alt(m)
            fb, fa = (tx.get(m) or (None, None))[1], (tx.get(ma) or (None, None))[0]
            if fb and fa:
                d["moeda_alt"], d["fx_alt"] = ma, fb / fa
        preencher_alt(d)


assert len({h for _, h, *_ in COLS}) == len(COLS), "há dois cabeçalhos iguais em COLS"
COL = {k: j for j, (k, *_) in enumerate(COLS, start=1)}   # chave -> nº da coluna (1 = A)
# cabeçalhos que mudam na folha ETFs (a mesma coluna, outro texto)
CAB_ETF = {"nota": "Nota geral ETF (pesos na folha Ponderações)"}
CAB_RANK = [h for _, h, *_ in COLS]
CAB_ETFS = [CAB_ETF.get(k, h) for k, h, *_ in COLS]
SO_ETF = ("etf_acoes", "etf_desc", "etf_comp")      # colunas só dos ETFs (escondidas na folha Ações)


def cab_ok(cab):
    """O cabeçalho lido do Excel é o desta versão (folha Ações ou ETFs)?"""
    return list(cab) in (CAB_RANK, CAB_ETFS)
L0 = 10   # linha do cabeçalho da tabela (V190: era 6; linhas 6-9: nota por ganhos). Ao ler um livro antigo, o
L0_NOVO = 10   # script vê onde está o cabeçalho (detetar_l0) e usa essa linha até refazer o livro
L0_ANTIGO = 11   # versões anteriores (com a secção ▶ GERAR): para aproveitar o "Ordenar por" ao atualizar
AC_L0 = 5   # linha do cabeçalho da folha Acrescentados

# "Ordenar por" (lista na linha 3 de cada folha): (texto, chave, maior primeiro?)
REC_ORDEM = {"strong_buy": 0, "buy": 1, "hold": 2, "underperform": 3, "sell": 4, "strong_sell": 5}
CRITERIOS_ORD = [
    ("Nota por ganhos (5 → 0)", "g_n5", True),
    ("Nota global (5 → 0)", "nota_geral", True),
    ("Nota geral, sem o preço (5 → 0)", "nota", True),
    ("Pontuação (5 → 1)", "pont", True),
    ("Detidos (x) primeiro", "det", True),
    ("Detidos (x): resultado decrescente", "det_lp_desc", True),
    ("Detidos (x): resultado crescente", "det_lp_asc", True),
    ("Melhor preço para comprar (nível de preço)", "nivel_preco", True),
    ("Melhor preço com nota igual ou maior a (valor ao lado)", "bp_min", True),
    ("Melhor nota com nível de preço igual ou maior a (valor ao lado)", "nt_npmin", True),
    ("Melhor preço com nota global igual ou maior a (valor ao lado)", "bp_gmin", True),
    ("Melhor nota global com nível de preço igual ou maior a (valor ao lado)", "ng_npmin", True),
    ("Analistas (alvo / preço)", "pot", True),
    ("Valor justo (justo / preço)", "justo_pct", True),
    ("Qualidade (0-5)", "qual", True),
    ("Valor transacionado/dia", "valor_tx", True),
    ("Volume (nº de títulos/dia)", "volume", True),
    ("Capitalização", "cap", True),
    ("Subida no último dia", "var1d", True),
    ("Subida no último mês", "var1m", True),
    ("Valorização desde sempre", "val_sempre", True),
    ("Valorização 10 anos", "val10", True),
    ("Valorização 5 anos", "val5", True),
    ("Valorização 3 anos", "val3", True),
    ("Valorização 1 ano", "val1", True),
    ("Recomendação", "rec", False),
    ("Mais abaixo do máx. da semana", "dist0", False),
]
ORD_ANTIGOS = {"Melhor preço com nota ≥ (valor ao lado)": "Melhor preço com nota igual ou maior a (valor ao lado)",
               "Nota por ganhos (maior primeiro)": "Nota por ganhos (5 → 0)"}   # nomes antigos da lista «Ordenar por»
ORD_PADRAO = {1: CRITERIOS_ORD[0][0], 2: CRITERIOS_ORD[0][0]}    # as duas folhas: pela nota geral
# barra de topo (V012): linha 3 = ▶ ATUALIZAR (A:D) · o que atualizar (E:G) · Ordenar por (H, I:L) · ↕ ORDENAR (M:N) ·
# estado (O); linha 4 = ↻ PREÇOS (A:D) · longo prazo mín. (E:F, G) · Situação em (H, I:J, K); linha 5 = bolsas
C_ORD = "I3"
C_ORD_MIN = "L3"     # nota mínima para «Melhor preço com nota ≥»
ORD_MIN = {"v": 4.0}
FICH_ORD_MIN = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_Painel_acoes_ordem_minimo.json")
try:
    with open(FICH_ORD_MIN, encoding="utf-8") as _f:
        ORD_MIN["v"] = float(json.load(_f).get("v", 4.0))
except Exception:
    pass
C_SEL = "E3"          # o que atualizar: Tudo / Assinalados / NG ≥ x
C_AT = "O3"           # "Atualizado em ..."
C_ORD_EST = "O4"
C_PRECOS_EST = "O5"   # estado do ↻ PREÇOS
CEL_ANTIGAS = {"ord": "E3", "lp": "D4", "ref": "H5", "data": "J5", "at": "K3"}   # painéis até à V011
def ler_sel(v):
    """Valor da lista 'o que atualizar' (aceita 'NG ≥ 2,5' / 'NG ≥ 2.5' / 'NG>=2,5'); None se não for nenhum."""
    t = txt(v)
    if not t:
        return None
    if t.lower().startswith("tudo"):
        return SELECOES[0]
    if t.lower().startswith("assinal"):
        return SELECOES[1]
    if t == SELECOES[2]:
        return t
    m = re.search(r"(\d+[.,]?\d*)", t)
    if m:                                  # "NG ≥ 2,5" de versões anteriores, ou só o número
        return f"NG ≥ {float(m.group(1).replace(',', '.')):.2f}".rstrip("0").rstrip(".")
    return None


def ler_sel_valor(sel, valor):
    """'NG ≥ (valor ao lado)' + o número em G3 -> 'NG ≥ 3.2' (None se faltar o número)."""
    if sel != SELECOES[2]:
        return sel
    try:
        v = float(str(valor).replace(",", ".").strip())
    except (TypeError, ValueError):
        print(f"Escolheste 'NG ≥' mas falta o valor em {C_SEL_V}: atualiza-se tudo.")
        return None
    return f"NG ≥ {v:.2f}".rstrip("0").rstrip(".")


SELECOES = ["Tudo", "Assinalados (coluna A)", "NG ≥ (valor ao lado)"]
C_SEL_V = "G3"        # valor da nota geral mínima (com "NG ≥")
MARGEM_200 = 0.03     # regra da média de 200 dias: sai quando o fecho fica 3% abaixo dela; volta a entrar quando fecha acima



def desempate(d):
    """Desempate dentro da mesma pontuação (tuplo para sorted: menor = primeiro).
    Ações: 1.º a qualidade (0-5), 2.º a média do potencial dos analistas e do valor justo (se só houver um, esse).
    ETFs: subida dos últimos 6 meses (momentum)."""
    num = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    if txt(d.get("tipo")) == "ETF":
        m = num("mom6")
        return (-(m if m is not None else -9),)
    q = num("qual")
    v = [x for x in (num("pot"), num("justo_pct")) if x is not None]
    return (-(q if q is not None else -1), -(sum(v) / len(v) if v else -9))


def chave_ordem(criterio):
    """Função de ordenação (para sorted) de um critério; os valores em falta ficam no fim."""
    criterio = ORD_ANTIGOS.get(criterio, criterio)
    _, k, desc = next((c for c in CRITERIOS_ORD if c[0] == criterio), CRITERIOS_ORD[0])

    def f(d):
        v = d.get(k)
        p = d.get("pont")
        p = p if isinstance(p, (int, float)) else 0
        nt = d.get("nota") if isinstance(d.get("nota"), (int, float)) else p
        if k == "bp_min":               # nota ≥ mínimo primeiro, pelo melhor preço (nível); depois os outros
            nv, jp = d.get("nivel_preco"), d.get("justo_pct")
            ok = isinstance(nt, (int, float)) and d.get("nota") is not None and round(nt, 2) >= ORD_MIN["v"] - 1e-9 and \
                isinstance(nv, (int, float))
            if ok:
                return (0, -nv, -(jp if isinstance(jp, (int, float)) else -9), -nt) + desempate(d)
            return (1, 0, 0, -nt, -p) + desempate(d)
        ng = d.get("nota_geral") if isinstance(d.get("nota_geral"), (int, float)) else nt
        if k == "bp_gmin":              # nota global ≥ mínimo primeiro, pelo melhor preço; depois os outros
            nv, jp = d.get("nivel_preco"), d.get("justo_pct")
            if isinstance(ng, (int, float)) and round(ng, 2) >= ORD_MIN["v"] - 1e-9 and isinstance(nv, (int, float)):
                return (0, -nv, -(jp if isinstance(jp, (int, float)) else -9), -ng) + desempate(d)
            return (1, 0, 0, -ng, -p) + desempate(d)
        if k == "ng_npmin":             # nível de preço ≥ mínimo primeiro, pela nota global; depois os outros
            nv = d.get("nivel_preco")
            if isinstance(nv, (int, float)) and round(nv, 2) >= ORD_MIN["v"] - 1e-9:
                return (0, -ng, -nv, -p) + desempate(d)
            return (1, -ng, -p) + desempate(d)
        if k == "nt_npmin":             # nível de preço ≥ mínimo primeiro, pela nota geral; depois os outros
            nv = d.get("nivel_preco")
            if isinstance(nv, (int, float)) and round(nv, 2) >= ORD_MIN["v"] - 1e-9:
                return (0, -nt, -nv, -p) + desempate(d)
            return (1, -nt, -p) + desempate(d)
        if k == "det":                  # detidos primeiro; dentro de cada grupo, pela nota geral e pelo desempate
            return (not e_detido(v), -nt, -p) + desempate(d)
        if k in ("det_lp_desc", "det_lp_asc"):   # detidos primeiro, pelo lucro / prejuízo (€); depois os outros
            lp = d.get("lp_eur")                 # pela nota geral
            tem = e_detido(d.get("det")) or isinstance(lp, (int, float)) and d.get("valor_eur")
            if not tem:
                return (2, 0, -nt, -p) + desempate(d)
            if not isinstance(lp, (int, float)):
                return (1, 0, -nt, -p) + desempate(d)
            return (0, -lp if k == "det_lp_desc" else lp, -nt, -p) + desempate(d)
        if k == "nota":                 # empate: pontuação, depois desempate
            return (not isinstance(v, (int, float)), -(v or 0), -p) + desempate(d)
        if k == "pont":                 # empate: desempate (ações: analistas + valor justo; ETFs: 6 meses)
            return (not isinstance(v, (int, float)), -(v or 0)) + desempate(d)
        if k == "rec":
            r = REC_ORDEM.get(str(v or "").strip().lower())
            return (r is None, r if r is not None else 0) + desempate(d)
        if not isinstance(v, (int, float)):
            return (True, 0)
        return (False, -v if desc else v)
    return f
PROTOCOLO = "painelacoes"   # link do botão ATUALIZAR (registado no Windows pelo próprio script)

VERDE, BRANCO, VERMELHO = "63BE7B", "FFFFFF", "F8696B"


def cor_cabecalho(grupo):
    if grupo == "sel":
        return "C00000"
    if grupo == "ind":
        return "4B2070"
    if grupo == "ganhos":
        return "375623"
    if grupo == "conselho":
        return "C55A11"
    if grupo.startswith("osc"):   # verdes, alternando o tom de período para período
        return "375623" if int(grupo[3:]) % 2 == 0 else "548235"
    if grupo.startswith("px"):    # azuis, alternando
        return "2E75B6" if int(grupo[2:]) % 2 == 0 else "1F5C99"
    if grupo == "compra":
        return "00A050"
    if grupo == "sugerido":
        return "7030A0"
    if grupo == "det":
        return "C00000"
    if grupo == "venda":
        return "C55A11"
    if grupo == "justo":
        return "203864"
    if grupo == "venda":
        return "C55A11"
    if grupo == "stop":
        return "C00000"
    if grupo == "teste":
        return "7F7F7F"
    if grupo == "mov":
        return "806000"
    if grupo == "global":
        return "000000"
    if grupo == "fund":
        return "385723"
    if grupo == "vol":
        return "404040"
    return "1F3864"


# versão do script (sobe a cada alteração): aparece no topo de todas as folhas e na página do telemóvel
VERSAO, DATA_VERSAO = "V192", "25/11/2026"
TXT_VERSAO = f"{VERSAO} ({DATA_VERSAO})"


def marcar_versao(ws):
    """Versão bem visível no canto superior esquerdo (a cor), juntando as primeiras colunas que forem precisas para
    caber (sem as alargar); o título passa para a coluna a seguir."""
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    tit = ws["A1"].value
    j, soma = 1, 0.0
    while True:                                # colunas até somarem ~16 de largura (máx. 6 colunas)
        soma += ws.column_dimensions[get_column_letter(j)].width or 8.43
        if soma >= 16 or j >= 6:
            break
        j += 1
    if tit and tit != TXT_VERSAO:
        ws.cell(1, j + 1, tit).font = Font(size=14, bold=True)
    if j > 1:
        ws.merge_cells(f"A1:{get_column_letter(j)}1")
    c = ws["A1"]
    c.value = TXT_VERSAO
    c.font = Font(size=11, bold=True, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="7030A0")
    c.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    ws.row_dimensions[1].height = 22


def carteira(d):
    """Valor detido, lucro/prejuízo (em euros, ao câmbio de hoje) e % (na moeda do título), a partir do nº de ações
    e do preço médio de compra (por ação, na moeda da cotação)."""
    num = lambda v: v if isinstance(v, (int, float)) and not isinstance(v, bool) else qtd_detida(v)
    q = num(d.get("qtd")) or qtd_detida(d.get("det"))
    pc, p = num(d.get("pcompra")), d.get("preco")
    d["valor_eur"] = d["lp_eur"] = d["lp_pct"] = d["venda_min_compra"] = None
    if isinstance(d.get("pont"), (int, float)) and (d.get("pont") == 1 or d.get("_vigiar")):
        d["decisao"] = decisao_global(d)       # com o preço de compra já conhecido: « · +x% sobre a tua compra»
    if q:
        d["qtd"] = q
        if not e_detido(d.get("det")) or qtd_detida(d.get("det")) is not None:
            d["det"] = "x"                  # o nº de ações passa da coluna A para a sua coluna
    vm = d.get("venda_min")
    if q and isinstance(pc, (int, float)) and pc > 0 and isinstance(vm, (int, float)):
        d["venda_min_compra"] = vm / pc - 1      # se vender na rede de segurança: ganho (+) ou perda (−) sobre a compra
    if not q or not isinstance(p, (int, float)):
        return
    fx = 1.0 if txt(d.get("moeda")) == "EUR" else (d.get("fx_alt") if txt(d.get("moeda_alt")) == "EUR" else None)
    if not isinstance(fx, (int, float)):
        return
    d["valor_eur"] = q * p * fx                # todas as casas decimais (a folha mostra 2)
    if pc:
        d["lp_eur"] = q * (p - pc) * fx
        d["lp_pct"] = p / pc - 1


def fmt_eur(v, sinal=False):
    t = f"{abs(v):,.2f}".replace(",", " ").replace(".", ",")
    return ("+" if v >= 0 else "−") + "€ " + t if sinal else ("−" if v < 0 else "") + "€ " + t


def totais_carteira(linhas, etfs):
    """{'acoes': (compra €, atual €), 'etfs': ..., 'total': ...} só com os títulos com nº de ações e preço de compra."""
    out = {}
    for nome, ls in (("acoes", linhas), ("etfs", etfs)):
        c = a = 0.0
        for d in ls:
            carteira(d)
            if isinstance(d.get("valor_eur"), (int, float)) and isinstance(d.get("lp_eur"), (int, float)):
                a += d["valor_eur"]
                c += d["valor_eur"] - d["lp_eur"]
        out[nome] = (c, a)
    out["total"] = (out["acoes"][0] + out["etfs"][0], out["acoes"][1] + out["etfs"][1])
    out["beta"] = betas_carteira(linhas, etfs)
    return out


FICH_BETAS = os.path.join(PASTA, "_Painel_acoes_betas.json")
FICH_COPIA = os.path.join(PASTA, "_Painel_acoes_copia_seguranca.xlsx")   # o livro antes da última ordenação/preços


def completar_betas(linhas, etfs):
    """Títulos detidos sem beta (folha de uma versão anterior, ou linha não atualizada): vai buscá-lo à última vez
    guardada ou ao Yahoo (só esses, em paralelo). Guarda os betas conhecidos para a próxima vez."""
    try:
        with open(FICH_BETAS, encoding="utf-8") as f:
            cache = json.load(f)
    except Exception:
        cache = {}
    todos = [(d, d in etfs) for d in list(linhas) + list(etfs)]
    for d, _ in todos:
        t = txt(d.get("ticker")).upper()
        if t and isinstance(d.get("beta"), (int, float)):
            cache[t] = d["beta"]
    falta = []
    for d, etf in todos:
        carteira(d)
        if not isinstance(d.get("beta"), (int, float)) and isinstance(d.get("valor_eur"), (int, float)):
            t = txt(d.get("ticker")).upper()
            if t in cache:
                d["beta"] = cache[t]
            elif t:
                falta.append((d, etf, t))
    if falta:
        try:
            from concurrent.futures import ThreadPoolExecutor
            yf = _yf()

            def um(x):
                d, etf, t = x
                try:
                    inf = yf.Ticker(t).info or {}
                    b = _num(inf.get("beta3Year" if etf else "beta")) or _num(inf.get("beta"))
                    return round(b, 2) if b is not None and -3 < b < 6 else None
                except Exception:
                    return None
            with ThreadPoolExecutor(8) as ex:
                for (d, etf, t), b in zip(falta, ex.map(um, falta, timeout=30)):
                    if b is not None:
                        d["beta"] = cache[t] = b
        except Exception:
            pass
    try:
        with open(FICH_BETAS, "w", encoding="utf-8") as f:
            json.dump(cache, f)
    except Exception:
        pass


def betas_carteira(linhas, etfs):
    """Beta médio ponderado pelo valor atual (só os títulos com beta): {'acoes': b, 'etfs': b, 'total': b}."""
    completar_betas(linhas, etfs)
    acc = {"acoes": [0.0, 0.0], "etfs": [0.0, 0.0]}
    for nome, ls in (("acoes", linhas), ("etfs", etfs)):
        for d in ls:
            v, b = d.get("valor_eur"), d.get("beta")
            if isinstance(v, (int, float)) and v > 0 and isinstance(b, (int, float)):
                acc[nome][0] += v * b
                acc[nome][1] += v
    t = [acc["acoes"][0] + acc["etfs"][0], acc["acoes"][1] + acc["etfs"][1]]
    return {k: (round(x[0] / x[1], 2) if x[1] else None) for k, x in (*acc.items(), ("total", t))}


def fmt_beta(b):
    return f"{b:.2f}".replace(".", ",") if isinstance(b, (int, float)) else "–"


def texto_quadro(tot):
    """As 3 linhas do quadro da carteira (ou None se não houver preços de compra)."""
    ct = tot["total"][0]
    if ct <= 0:
        return None
    ls = []
    for nome, k in (("Ações", "acoes"), ("ETFs", "etfs"), ("Total", "total")):
        c, a = tot[k]
        quota = f" ({c / ct:.0%}/carteira)".replace(".", ",") if k != "total" else ""
        var = f" ({a / c - 1:+.2%})".replace(".", ",") if c else ""
        bt = (tot.get("beta") or {}).get(k)
        ls.append(f"{nome}:  valor de compra {fmt_eur(c)}{quota}  /  valor atual {fmt_eur(a)}{var}"
                  + (f"  /  beta {fmt_beta(bt)}" if bt is not None else ""))
    if (tot.get("beta") or {}).get("total") is not None:
        ls.append("Beta = volatilidade comparada com o mercado (S&P 500): 1 = como o mercado; 1,3 = 30% mais")
    return "\n".join(ls)


def folha_carteira(wb, linhas, etfs):
    """Folha Carteira: tabela e 3 gráficos (pie por título, pie ações vs ETFs, barras com lucro a verde / prejuízo a
    vermelho)."""
    from openpyxl.chart import PieChart, BarChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.styles import Font, PatternFill, Alignment
    ws = wb.create_sheet("Carteira")
    ws["A1"] = "Carteira: valor de compra, valor atual e lucro / prejuízo (em euros, ao câmbio de hoje)"
    ws["A1"].font = Font(size=14, bold=True)
    cab = ["Título", "Tipo", "Compra (€)", "Atual (€)", "Lucro / prejuízo (€)", "Lucro / prejuízo (%)",
           "Valor de compra (mantido)", "Lucro", "Prejuízo", "Beta"]
    for j, h in enumerate(cab, 1):
        c = ws.cell(3, j, h)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3B57")
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    dados = []
    tot = totais_carteira(linhas, etfs)                 # (também completa os betas em falta)
    for d in list(linhas) + list(etfs):
        carteira(d)
        if isinstance(d.get("valor_eur"), (int, float)) and isinstance(d.get("lp_eur"), (int, float)):
            a, l = d["valor_eur"], d["lp_eur"]
            dados.append((txt(d.get("nome")) + (f" ({txt(d.get('ticker'))})" if modo_atual() == "todos" else ""),
                          txt(d.get("tipo")) or "Ação", round(a - l, 2), a, l,
                          d["beta"] if isinstance(d.get("beta"), (int, float)) else None))
    dados.sort(key=lambda x: -x[2])
    for i, (n, t, c, a, l, bt) in enumerate(dados, 4):
        for j, v in enumerate([n, t, c, a, l, (a / c - 1) if c else None, min(c, a), max(l, 0), max(-l, 0), bt], 1):
            cel = ws.cell(i, j, v)
            cel.number_format = ("+0.00%;-0.00%;0.00%" if j == 6 else "0.00" if j == 10
                                 else ('"€"* #,##0.00' if j >= 3 else "General"))
        ws.cell(i, 5).font = ws.cell(i, 6).font = Font(bold=True, color="00A050" if l >= 0 else "C00000")
    for col, w in zip("ABCDEFGHIJ", (34, 8, 13, 13, 13, 10, 11, 11, 11, 7)):
        ws.column_dimensions[col].width = w
    if not dados:
        ws["A5"] = "Escreve o nº de ações e o preço médio de compra nas folhas Ações / ETFs para ver a carteira."
        return
    n = len(dados)
    # totais ações vs ETFs
    r0 = 4 + n + 2
    ws.cell(r0, 1, "Totais").font = Font(bold=True)
    ws.cell(r0, 10, "Beta médio").font = Font(bold=True)
    for i, k in zip((r0 + 1, r0 + 2, r0 + 3), ("acoes", "etfs", "total")):
        ws.cell(i, 10, tot["beta"].get(k)).number_format = "0.00"
    ws.cell(r0 + 3, 10).font = Font(bold=True)
    ws.cell(r0 + 5, 1, "Beta = volatilidade comparada com o mercado (S&P 500): 1 = oscila como o mercado; 1,3 = oscila "
                       "30% mais; 0,7 = oscila 30% menos. Beta médio: ponderado pelo valor atual.").font = \
        Font(italic=True, color="595959")
    for i, (nome, k) in enumerate((("Ações", "acoes"), ("ETFs", "etfs")), r0 + 1):
        ws.cell(i, 1, nome)
        ws.cell(i, 3, round(tot[k][0], 2)).number_format = '"€"* #,##0.00'
        ws.cell(i, 4, round(tot[k][1], 2)).number_format = '"€"* #,##0.00'
    ws.cell(r0 + 3, 1, "Total").font = Font(bold=True)
    ws.cell(r0 + 3, 3, round(tot["total"][0], 2)).number_format = '"€"* #,##0.00'
    ws.cell(r0 + 3, 4, round(tot["total"][1], 2)).number_format = '"€"* #,##0.00'

    def rotulos():
        dl = DataLabelList()
        dl.showVal, dl.showPercent, dl.showCatName, dl.showSerName, dl.showLeaderLines = True, True, False, False, True
        dl.numFmt = '"€" #,##0'
        return dl
    p1 = PieChart()
    p1.title = "Carteira por valor de compra (€ e %)"
    p1.add_data(Reference(ws, min_col=3, min_row=3, max_row=3 + n), titles_from_data=True)
    p1.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + n))
    p1.dataLabels = rotulos()
    p1.width, p1.height = 18, 11
    ws.add_chart(p1, "K3")
    p2 = PieChart()
    p2.title = "Ações vs ETFs (valor de compra, € e %)"
    p2.add_data(Reference(ws, min_col=3, min_row=r0 + 1, max_row=r0 + 2), titles_from_data=False)
    p2.set_categories(Reference(ws, min_col=1, min_row=r0 + 1, max_row=r0 + 2))
    p2.dataLabels = rotulos()
    p2.width, p2.height = 12, 9
    ws.add_chart(p2, "K26")
    p2a = PieChart()                                    # ao lado: a mesma divisão, mas pelo valor atual
    p2a.title = "Ações vs ETFs (valor atual, € e %)"
    p2a.add_data(Reference(ws, min_col=4, min_row=r0 + 1, max_row=r0 + 2), titles_from_data=False)
    p2a.set_categories(Reference(ws, min_col=1, min_row=r0 + 1, max_row=r0 + 2))
    p2a.dataLabels = rotulos()
    p2a.width, p2a.height = 12, 9
    ws.add_chart(p2a, "R26")
    b = BarChart()
    b.type, b.grouping, b.overlap = "col", "stacked", 100
    b.title = "Cada título: valor de compra, + lucro (verde) / − prejuízo (vermelho)"
    b.add_data(Reference(ws, min_col=7, max_col=9, min_row=3, max_row=3 + n), titles_from_data=True)
    b.set_categories(Reference(ws, min_col=1, min_row=4, max_row=3 + n))
    for sr, cor in zip(b.series, ("A6A6A6", "00A050", "C00000")):
        sr.graphicalProperties.solidFill = cor
        sr.graphicalProperties.line.solidFill = cor
    b.y_axis.title = "Valor (€)"
    b.x_axis.title = "Título"
    b.y_axis.numFmt = '#,##0" €"'
    b.x_axis.delete = b.y_axis.delete = False           # sem isto o Excel recente esconde os eixos
    b.x_axis.tickLblPos = b.y_axis.tickLblPos = "low"
    from openpyxl.chart.axis import ChartLines
    b.y_axis.majorGridlines = ChartLines()
    b.legend.position = "b"
    b.width, b.height = max(18, min(40, 1.4 * n)), 11
    ws.add_chart(b, "K45")
    # lucros (verdes) e prejuízos (vermelhos): cada fatia = o lucro ou o prejuízo de um título, em valor absoluto
    from openpyxl.chart.series import DataPoint
    lp = sorted([x for x in dados if x[4] > 0], key=lambda x: -x[4]) + sorted([x for x in dados if x[4] < 0], key=lambda x: x[4])
    if lp:
        r1 = r0 + 7
        ws.cell(r1, 1, "Lucros e prejuízos (gráfico circular)").font = Font(bold=True)
        ws.cell(r1 + 1, 1, "Título").font = ws.cell(r1 + 1, 3, "Lucro / prejuízo (€)").font = Font(bold=True)
        ws.cell(r1 + 1, 4, "Tamanho da fatia (€)").font = Font(bold=True)
        for i, (nome, _, c, a, l, *_r) in enumerate(lp, r1 + 2):
            ws.cell(i, 1, f"{nome} ({'+' if l >= 0 else '−'})")
            ws.cell(i, 3, round(l, 2)).number_format = '+"€" #,##0.00;-"€" #,##0.00'
            ws.cell(i, 3).font = Font(bold=True, color="00A050" if l >= 0 else "C00000")
            ws.cell(i, 4, round(abs(l), 2)).number_format = '"€" #,##0.00'
            sg, tot_lp = ("+" if l >= 0 else "−"), sum(abs(x[4]) for x in lp)
            ws.cell(i, 2, f"{nome}: {sg}€ {abs(l):,.0f} ({sg}{abs(l) / tot_lp * 100:.1f}%)"
                    .replace(",", " ").replace(".", ",")).font = Font(color="808080", size=8)
        p3 = PieChart()
        p3.title = "Lucros (verde) e prejuízos (vermelho) por título"
        p3.add_data(Reference(ws, min_col=4, min_row=r1 + 2, max_row=r1 + 1 + len(lp)), titles_from_data=False)
        p3.set_categories(Reference(ws, min_col=2, min_row=r1 + 2, max_row=r1 + 1 + len(lp)))
        VERDES = ["1E8E3E", "34A853", "5BB974", "81C995", "0D652D", "A8DAB5"]
        VERMELHOS = ["D93025", "EA4335", "EE675C", "F28B82", "A50E0E", "F6AEA9"]
        ig = iv = 0
        for k, x in enumerate(lp):
            pt = DataPoint(idx=k)
            if x[4] >= 0:
                cor, ig = VERDES[ig % len(VERDES)], ig + 1
            else:
                cor, iv = VERMELHOS[iv % len(VERMELHOS)], iv + 1
            pt.graphicalProperties.solidFill = cor
            pt.graphicalProperties.line.solidFill = "FFFFFF"
            p3.series[0].dPt.append(pt)
        dl = DataLabelList()                            # rótulos com o sinal: prejuízos com € e % negativos
        dl.showCatName, dl.showVal, dl.showPercent, dl.showSerName, dl.showLeaderLines = True, False, False, False, True
        p3.dataLabels = dl
        p3.legend = None
        p3.width, p3.height = 16, 11
        ws.add_chart(p3, "Y26")


QUADRO = None         # dados da janela da posição (depois de cada ▶ ATUALIZAR)
FICH_POS = os.path.join(PASTA, "_Painel_acoes_posicao.json")
FICH_JANELA = os.path.join(PASTA, "_Painel_acoes_janela.json")     # tamanho e posição da janela


def dados_posicao(linhas, etfs):
    tot = totais_carteira(linhas, etfs)
    tit = []
    for tipo, ls in (("Ação", linhas), ("ETF", etfs)):
        for d in ls:
            if isinstance(d.get("valor_eur"), (int, float)) and isinstance(d.get("lp_eur"), (int, float)):
                q = d.get("qtd") if isinstance(d.get("qtd"), (int, float)) else qtd_detida(d.get("det"))
                tit.append([txt(d.get("nome")), tipo, q, d["valor_eur"] - d["lp_eur"], d["valor_eur"],
                            d.get("beta") if isinstance(d.get("beta"), (int, float)) else None])
    tit.sort(key=lambda x: -x[3])
    return {"quando": f"{datetime.now():%d/%m/%Y %H:%M}", "acoes": tot["acoes"], "etfs": tot["etfs"],
            "total": tot["total"], "beta": tot.get("beta") or {}, "titulos": tit}


def gravar_posicao(dp):
    try:
        with open(FICH_POS, "w", encoding="utf-8") as f:
            json.dump(dp, f, ensure_ascii=False, default=_json_num)
    except Exception:
        pass


def abrir_janela_posicao():
    """Abre a janela da posição num processo à parte (o Excel e a consola ficam livres)."""
    if os.name != "nt":
        return
    import subprocess
    exe = sys.executable
    w = os.path.join(os.path.dirname(exe), "pythonw.exe")
    try:
        subprocess.Popen([w if os.path.exists(w) else exe, os.path.abspath(__file__), "--janela"],
                         creationflags=0x00000008 | 0x00000200, close_fds=True)   # DETACHED_PROCESS | NEW_GROUP
    except Exception as e:
        print("   aviso: não consegui abrir a janela da posição:", e)


def janela_posicao():
    """Janela cinzenta do Windows, redimensionável, com o resumo da carteira e a lista dos títulos. Guarda o tamanho e
    a posição ao fechar (o ▣ VER POSIÇÃO volta a abri-la igual)."""
    import tkinter as tk
    from tkinter import ttk
    try:
        with open(FICH_POS, encoding="utf-8") as f:
            dp = json.load(f)
    except Exception:
        dp = None
    raiz = tk.Tk()
    raiz.title("Posição da carteira")
    try:
        with open(FICH_JANELA, encoding="utf-8") as f:
            raiz.geometry(json.load(f)["geo"])
    except Exception:
        raiz.geometry("760x460+120+120")
    raiz.minsize(480, 260)
    st = ttk.Style(raiz)
    try:
        st.theme_use("vista" if "vista" in st.theme_names() else "clam")
    except Exception:
        pass
    verde, verm = "#1e7e34", "#c00000"
    cab = ttk.Frame(raiz, padding=(12, 10, 12, 4))
    cab.pack(fill="x")
    ttk.Label(cab, text="Posição da carteira", font=("Segoe UI", 13, "bold")).pack(side="left")
    ttk.Label(cab, text=("atualizado em " + dp["quando"]) if dp else "", foreground="#595959").pack(side="right")
    if not dp or not dp["total"][0]:
        ttk.Label(raiz, text="Sem títulos com nº de ações e preço médio de compra (colunas C e D das folhas Ações / ETFs).",
                  padding=12).pack(fill="both", expand=True)
    else:
        res = ttk.Frame(raiz, padding=(12, 4))
        res.pack(fill="x")
        ct = dp["total"][0]
        cols = ("", "Valor de compra", "% da carteira", "Valor atual", "Lucro / prejuízo", "%", "Beta médio")
        bts = dp.get("beta") or {}
        for j, h in enumerate(cols):
            ttk.Label(res, text=h, font=("Segoe UI", 9, "bold"), anchor="e" if j else "w").grid(row=0, column=j, sticky="ew", padx=6)
            res.columnconfigure(j, weight=1)
        for i, (nome, k) in enumerate((("Ações", "acoes"), ("ETFs", "etfs"), ("Total", "total")), 1):
            c, a = dp[k]
            l = a - c
            cor = verde if l >= 0 else verm
            negr = ("Segoe UI", 10, "bold") if k == "total" else ("Segoe UI", 10)
            vals = (nome, fmt_eur(c), f"{c / ct:.1%}".replace(".", ",") if k != "total" else "100%", fmt_eur(a),
                    fmt_eur(l, True), (f"{l / c:+.2%}".replace(".", ",") if c else ""), fmt_beta(bts.get(k)))
            for j, v in enumerate(vals):
                ttk.Label(res, text=v, font=negr, anchor="e" if j else "w",
                          foreground=cor if 4 <= j <= 5 else None).grid(row=i, column=j, sticky="ew", padx=6, pady=1)
        ttk.Label(raiz, padding=(18, 2, 12, 0), foreground="#595959", font=("Segoe UI", 9), justify="left",
                  text="Beta = volatilidade comparada com o mercado (S&P 500): 1 = oscila como o mercado; "
                       "1,3 = oscila 30% mais; 0,7 = oscila 30% menos.\nBeta médio: média dos títulos, "
                       "ponderada pelo valor atual de cada um.").pack(fill="x")
        ttk.Separator(raiz).pack(fill="x", padx=12, pady=6)
        fr = ttk.Frame(raiz, padding=(12, 0))
        fr.pack(fill="both", expand=True)
        cs = ("tit", "tipo", "qtd", "compra", "atual", "lp", "pct", "beta")
        nomes = ("Título", "Tipo", "Nº", "Compra (€)", "Atual (€)", "Lucro / prejuízo (€)", "%", "Beta")
        tv = ttk.Treeview(fr, columns=cs, show="headings", height=8)
        tv.tag_configure("g", foreground=verde)
        tv.tag_configure("p", foreground=verm)
        linhas_t = []                                   # [valores de ordenação por coluna, valores mostrados, tag]
        for n, tp, q, c, a, *resto in dp["titulos"]:
            bt = resto[0] if resto else None
            l = a - c
            qn = q if isinstance(q, (int, float)) else None
            linhas_t.append(((n.lower(), tp, qn, c, a, l, (l / c if c else None), bt),
                             (n, tp, f"{q:,.2f}".replace(",", " ").replace(".", ",") if qn is not None else "",
                              fmt_eur(c), fmt_eur(a), fmt_eur(l, True), f"{l / c:+.2%}".replace(".", ",") if c else "",
                              fmt_beta(bt) if bt is not None else ""),
                             "g" if l >= 0 else "p"))
        estado = {"col": 3, "desc": True}               # começa por valor de compra, do maior para o menor

        def mostrar():
            j, desc = estado["col"], estado["desc"]
            com = [x for x in linhas_t if x[0][j] is not None]
            sem = [x for x in linhas_t if x[0][j] is None]   # vazios sempre no fim
            com.sort(key=lambda x: x[0][j], reverse=desc)
            tv.delete(*tv.get_children())
            for _, vals, tg in com + sem:
                tv.insert("", "end", tags=(tg,), values=vals)
            for i, (cid, h) in enumerate(zip(cs, nomes)):
                tv.heading(cid, text=h + ((" ▼" if desc else " ▲") if i == j else ""))

        def ordenar(j):
            if estado["col"] == j:
                estado["desc"] = not estado["desc"]
            else:                                       # texto: A→Z; números: do maior para o menor
                estado["col"], estado["desc"] = j, j >= 2
            mostrar()
        for i, (cid, h, w, an) in enumerate(zip(cs, nomes, (220, 50, 60, 100, 100, 145, 70, 60),
                                                ("w", "center", "e", "e", "e", "e", "e", "e"))):
            tv.heading(cid, text=h, command=lambda i=i: ordenar(i))
            tv.column(cid, width=w, anchor=an, stretch=True)
        mostrar()
        sb = ttk.Scrollbar(fr, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sb.set)
        tv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

    def fechar():
        try:
            with open(FICH_JANELA, "w", encoding="utf-8") as f:
                json.dump({"geo": raiz.geometry()}, f)
        except Exception:
            pass
        raiz.destroy()
    rod = ttk.Frame(raiz, padding=10)
    rod.pack(fill="x", side="bottom")
    ttk.Button(rod, text="Fechar", command=fechar, width=12).pack(side="right")
    raiz.protocol("WM_DELETE_WINDOW", fechar)
    raiz.bind("<Escape>", lambda e: fechar())
    raiz.attributes("-topmost", True)                   # por cima do Excel ao abrir...
    raiz.after(1500, lambda: raiz.attributes("-topmost", False))   # ...mas depois como uma janela normal
    raiz.mainloop()


def ver_posicao():
    """▣ VER POSIÇÃO: lê as folhas Ações e ETFs como estão (nº de ações e preços de agora) e abre a janela."""
    try:
        ex = Excel(FICHEIRO)
        ultima = get_column_letter(len(COLS))
        r = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)])
        chaves = [k for k, *_ in COLS]
        ls = []
        for idx, tab in zip((1, 2), r):
            ls.append(linhas_por_cabecalho(tab, idx))
        gravar_posicao(dados_posicao(*ls))
    except Exception:
        pass                                            # fica a última posição guardada
    abrir_janela_posicao()


def aplicar_compra(d, qtd, pcompra):
    """Nº de ações e preço de compra guardados (por ticker; por nome se não for 'todas as bolsas')."""
    chaves = [txt(d.get("ticker")).upper()] + ([] if modo_atual() == "todos" else [txt(d.get("nome"))])
    if not isinstance(d.get("qtd"), (int, float)):
        q = next((qtd[k] for k in chaves if k in qtd), None)
        if q is not None:
            d["qtd"] = q
    if not isinstance(d.get("pcompra"), (int, float)):
        pc = next((pcompra[k] for k in chaves if k in pcompra), None)
        if pc is not None:
            d["pcompra"] = pc


def aplicar_corr(d, corr):
    """Corretora guardada (T212 / TRep / T212+TRep), por ticker ou nome."""
    chaves = [txt(d.get("ticker")).upper()] + ([] if modo_atual() == "todos" else [txt(d.get("nome"))])
    if not txt(d.get("corr")):
        d["corr"] = next((corr[k] for k in chaves if k in (corr or {})), None)


def marca_detido(d, det, qtd):
    """O nº de ações (se o souber) ou x, para um título que tens; None se não o tens."""
    chaves = [txt(d.get("ticker")).upper()] + ([] if modo_atual() == "todos" else [txt(d.get("nome"))])
    for k in chaves:                    # em "todas as bolsas" o nº de ações fica só na listagem onde o escreveste
        if k in qtd:
            return qtd[k]
    return "x" if (txt(d.get("nome")) in det or txt(d.get("ticker")).upper() in det) else None


def construir(linhas, atualizado, etfs=(), criterios=None, acrescentados=(), estado=None):
    global L0
    L0 = L0_NOVO                                  # V190: o livro novo tem o cabeçalho na linha 10
    if not any("g_nota" in d for d in list(linhas) + list(etfs or [])):
        g_juntar(linhas, False)                   # nota por ganhos do último cálculo
        g_juntar(etfs or [], True)
    criterios = {**ORD_PADRAO, **(criterios or {})}
    estado = estado or {}
    det = estado.get("det") or set()
    at_a, at_e = atualizado if isinstance(atualizado, tuple) else (atualizado, atualizado)
    cambios_alt(list(linhas) + list(etfs))
    try:
        if calibrar_geral(linhas):
            print("Nota global: " + calib_txt().split(" Só informa")[0])
    except Exception as e:
        print("   aviso: calibração da nota global:", e)
    for d in list(linhas) + list(etfs):
        d["horario"] = horario_bolsa(d.get("ticker")) or d.get("horario")                 # x nos títulos que tens (lidos do painel anterior)
        if not e_detido(d.get("det")):
            d["det"] = marca_detido(d, det, estado.get("qtd") or {})
        aplicar_compra(d, estado.get("qtd") or {}, estado.get("pcompra") or {})
        aplicar_corr(d, estado.get("corr") or {})
        carteira(d)
    wb = Workbook()
    folha(wb.active, FOLHA_A2, "Painel de ações: nota geral (sinal das médias, valor justo e analistas + qualidade e saúde financeira)",
          sorted(linhas, key=chave_ordem(criterios[1])), at_a, True, criterios[1], 1, estado)
    folha(wb.create_sheet(), FOLHA_E2, "Painel de ETFs: nota geral (sinal das médias + nota das ações que os compõem)",
          sorted(etfs, key=chave_ordem(criterios[2])), at_e, False, criterios[2], 2, estado)
    for pf in (FOLHA_PROC, FOLHA_PROC_ETF):
        folha_procurar(wb, pf)
        p = wb[NOME_PROC[pf]]
        for cel, v in estado.get("proc", {}).get(pf, {}).items():
            if v not in (None, ""):
                p[cel] = v
    folha_acrescentados(wb, acrescentados)
    wb.active = 0
    try:
        folha_compra(wb)
    except Exception as e:
        print("   aviso: folha Compra não criada:", e)
    try:
        folha_pesos(wb)
    except Exception as e:
        print("   aviso: folha Ponderações não criada:", e)
    try:
        folha_carteira(wb, linhas, etfs)
    except Exception as e:
        print("   aviso: folha Carteira não criada:", e)
    try:
        folha_simulador(wb)
    except Exception as e:
        print("   aviso: folha Simulador não criada:", e)
    try:
        folha_historico(wb, linhas, etfs)
    except Exception as e:
        print("   aviso: folha Histórico não criada:", e)
    try:
        folha_evolucao(wb)
    except Exception as e:
        print("   aviso: folha Evolução não criada:", e)
    try:
        folha_comp_rent(wb)
    except Exception as e:
        print("   aviso: folha Comp Rent Títulos não criada:", e)
    for etf_, nf_ in ((False, FOLHA_A2), (True, FOLHA_E2)):   # V190: nota por ganhos (parâmetros e botão)
        try:
            g_topo(wb, wb[nf_], etf_)
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"   aviso: parâmetros da nota por ganhos ({nf_}) não criados:", e)
    try:
        folha_comparar(wb)
    except Exception as e:
        print("   aviso: folha Comparar notas não criada:", e)
    if FOLHA_CRT in wb.sheetnames and "Carteira" in wb.sheetnames:   # V169: logo a seguir à Carteira
        wb.move_sheet(FOLHA_CRT, offset=wb.sheetnames.index("Carteira") + 1 - wb.sheetnames.index(FOLHA_CRT))
    for nome in ("Acrescentados", FOLHA_PESOS):       # estas duas no fim (são pouco usadas)
        if nome in wb.sheetnames:
            wb.move_sheet(nome, offset=len(wb.sheetnames) - 1 - wb.sheetnames.index(nome))
    frente = [n_ for n_ in (FOLHA_A2, FOLHA_E2) if n_ in wb.sheetnames]
    wb._sheets = [wb[n_] for n_ in frente] + [w_ for w_ in wb._sheets if w_.title not in frente]
    wb.active = 0
    for ws in wb.worksheets:
        marcar_versao(ws)
    return wb


def _fmt_cel_base(c, k, fmt):
    """Formato de uma célula da tabela das folhas Ações / ETFs."""
    if k == "nome":
        c.font = Font(color="000000")
    if k in NEGRITO:                    # preço atual e níveis de compra e venda a negrito
        c.font = Font(bold=True, color="C00000" if k == "venda_min" else None)
    elif k == "venda_min_pct":          # rede de segurança: a vermelho, é extraordinária
        c.font = Font(color="C00000")
    elif fmt == "alt":                  # o mesmo valor noutra moeda: letra 9
        c.font = Font(size=9, color="C00000" if k == "venda_min_alt" else "595959")
    elif k in ("det", "sel"):           # x = título que tens / a atualizar
        c.font = Font(bold=True, color="C00000")
        c.alignment = Alignment(horizontal="center")
    if k in FONTE9:
        c.font = Font(size=9, bold=bool(c.font and c.font.b))
    if k in ("qtd", "corr"):
        c.alignment = Alignment(horizontal="center")
    if fmt == "preco":                  # o símbolo da moeda vem da formatação condicional (segue a linha)
        c.number_format = "#,##0.00"
    elif fmt == "alt":
        c.number_format = '"("#,##0.00")"'
    elif fmt == "pct":
        c.number_format = "+0.00%;-0.00%;0.00%" if k == "lp_pct" else "+0.0%;-0.0%;0.0%"
    elif fmt == "num1":
        c.number_format = "0.00" if k in ("beta", "r2_1a", "r2_6m", "fator_tend", "rec_media", "g_medida", "g_fpreco",
                                          "g_fvol") else "#,##0.0"
    elif fmt == "qtd":
        c.number_format = "#,##0.00"
    elif fmt == "eur":
        c.number_format = '"€"* #,##0.00'
        if k == "valor_eur":
            c.font = Font(bold=True)
    if k in ("g_nota", "g_n5"):                   # V190: a nota por ganhos a negrito
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")
    if k in ("bolsa", "horario"):                 # bolsa e horário: letra 9
        c.font = Font(size=9)
    elif fmt == "nota":
        c.number_format = "0.0"
        c.alignment = Alignment(horizontal="center")
        if k in ("nota", "nota_geral"):
            c.font = Font(bold=True)
    elif fmt == "pont":
        c.number_format = "General"
        c.alignment = Alignment(horizontal="center")
    elif fmt == "int":
        c.number_format = "#,##0"
    elif fmt == "nivel":
        c.number_format = "0.0"
        c.font = Font(size=9, bold=True)
        c.alignment = Alignment(horizontal="left")


def botao(ws, cel, texto, link, cor, borda_cor, tamanho=12):
    ws[cel] = texto
    ws[cel].hyperlink = link
    ws[cel].font = Font(bold=True, color="FFFFFF", size=tamanho)
    ws[cel].fill = PatternFill("solid", fgColor=cor)
    ws[cel].alignment = Alignment(horizontal="center", vertical="center")
    ws[cel].border = Border(*(Side(style="medium", color=borda_cor),) * 4)


def largura(w, fmt):
    """Largura da coluna: as de números ficam estreitas (cabeçalho em letra 8, em várias linhas)."""
    if fmt in ("pct", "num1", "int", "nota", "pont", "alt", "qtd"):
        return min(w, 8)
    if fmt == "preco":                   # símbolo da moeda à esquerda + número
        return min(w, 10)
    return w


def folha(ws, nome, titulo, linhas, atualizado, ranking, criterio, idx, estado=None):
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import column_index_from_string
    estado = estado or {}
    borda = Border(*(Side(style="thin", color="A6A6A6"),) * 4)
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    lista = lambda itens: '"' + ",".join(itens) + '"'     # opções escritas na própria lista (≤ 255 caracteres)

    def campo(intervalo, valor, fundo=amarelo):
        ws.merge_cells(intervalo)
        ini = intervalo.split(":")[0]
        ws[ini] = valor
        for linha in ws[intervalo]:
            for c in linha:
                c.fill, c.border = fundo, borda
        ws[ini].alignment = Alignment(vertical="center")

    def texto(intervalo, valor):
        ws.merge_cells(intervalo)
        ini = intervalo.split(":")[0]
        ws[ini] = valor
        ws[ini].font = Font(bold=True)
        ws[ini].alignment = Alignment(horizontal="right", vertical="center")

    def lista_val(cel, itens):
        dv = DataValidation(type="list", formula1=lista(itens), allow_blank=False)
        dv.error, dv.errorTitle = "Escolhe um valor da lista.", "Valor inválido"
        ws.add_data_validation(dv)
        dv.add(cel)

    ws.title = nome
    ws["A1"] = titulo
    ws["A1"].font = Font(size=14, bold=True)
    col_u = C_UTIL.rstrip("0123456789")
    ant_u = get_column_letter(column_index_from_string(col_u) - 1)
    ws[f"{ant_u}1"] = "Utilizador:"
    ws[f"{ant_u}1"].font = Font(bold=True)
    ws[f"{ant_u}1"].alignment = Alignment(horizontal="right", vertical="center")
    if idx == 1:                                 # Ações: a escolha (lista); ETFs: mostra a mesma
        ws[C_UTIL] = utilizador()
        dv_u = DataValidation(type="list", formula1='"' + ",".join(UTILIZADORES) + '"', allow_blank=False)
        dv_u.error, dv_u.errorTitle = "Escolhe um utilizador da lista.", "Utilizador"
        ws.add_data_validation(dv_u)
        dv_u.add(C_UTIL)
        ws[C_UTIL].fill = PatternFill("solid", fgColor="FFF2CC")
        ws[C_UTIL].border = borda
    else:
        ws[C_UTIL] = f"='{FOLHA_A2}'!{C_UTIL}"
    ws[C_UTIL].font = Font(bold=True, color="7030A0")
    ws[C_UTIL].alignment = Alignment(vertical="center")
    ws["E2"] = ("Coluna A: x nos títulos a atualizar (com \"Assinalados\" na lista ao lado do ▶ ATUALIZAR, "
                "ou \"NG ≥\" com a nota mínima escrita ao lado; as linhas não atualizadas ficam a rosa até serem atualizadas ou até um \"Tudo\"). Coluna B: x nos títulos que "
                "tens; colunas C e D: nº de ações e preço médio de compra por ação (na moeda da cotação) -> valor detido e "
                "lucro/prejuízo, também na página do telemóvel e na folha Carteira. \"Situação em\": cada folha tem a sua data "
                "(Atual ou uma data passada, escrita em qualquer formato). Só são atualizados os títulos que estão nas linhas "
                "da folha (escondidas ou não): o que apagares deixa de ser procurado. "
                "▶ ATUALIZAR (aceita os avisos do Excel/Windows): atualiza só esta folha (" + ("ações" if ranking else "ETFs")
                + "; a outra fica como estava); o Excel fecha, os dados são atualizados e o painel "
                "volta a abrir; ↻ PREÇOS atualiza só os preços (segundos). \"Ordenar por\" + ↕ ORDENAR reordena logo (sem descarregar nada) e também se aplica em cada "
                "atualização. Por omissão ordena pela NOTA GLOBAL (nota geral + preço). Nota geral: " + (
                    texto_pesos(True) + "; × fator de tendência (folha Ponderações); no máx. 2,5 com saúde em alerta. "
                    if ranking else texto_pesos(False) + " (as ações que faltam na folha Ações são lá acrescentadas); "
                    "× fator de tendência (folha Ponderações). ")
                + ("" if ranking else "Os ETFs não têm preço-alvo nem valor justo (esses critérios deixam-nos no fim). ")
                + LP_AJUDA
                + "▣ VER POSIÇÃO: abre a janela com o valor de compra, o valor atual e o lucro/prejuízo da carteira. "
                + f"Se os botões não responderem, corre uma vez o {os.path.basename(BAT)} (isso ativa-os neste PC).")
    ws["E2"].font = Font(italic=True, size=9, color="595959")
    ws["E2"].alignment = Alignment(vertical="center")
    # linha 2: ▶ ATUALIZAR TUDO (as duas folhas)
    ws.merge_cells("A2:D2")
    botao(ws, "A2", "▶ ATUALIZAR TUDO", f"{PROTOCOLO}://atualizar/tudo", "1F3B57", "0F2236", 11)
    ws.row_dimensions[2].height = 24
    # linha 3: ▶ ATUALIZAR · o que atualizar · Ordenar por · ↕ ORDENAR · data
    ws.merge_cells("A3:D3")
    botao(ws, "A3", "▶ ATUALIZAR " + ("AÇÕES" if ranking else "ETFs"),
          f"{PROTOCOLO}://atualizar/{'acoes' if ranking else 'etfs'}", "00A050", "006E37", 11)
    sel0 = (estado or {}).get("sel" + ("" if ranking else "_e")) or SELECOES[0]
    val0 = None
    if sel0.startswith("NG ≥ ") and sel0 != SELECOES[2]:
        val0, sel0 = float(sel0.split("≥")[1]), SELECOES[2]
    campo("E3:F3", sel0)
    lista_val(C_SEL, SELECOES)
    lc = get_column_letter(COL["corr"])                 # corretora: lista (em branco / T212 / TRep / T212+TRep)
    dvc = DataValidation(type="list", formula1='"' + ",".join(CORRETORAS) + '"', allow_blank=True)
    dvc.error, dvc.errorTitle = "Escolhe T212, TRep ou T212+TRep (ou deixa em branco).", "Corretora"
    ws.add_data_validation(dvc)
    dvc.add(f"{lc}{L0 + 1}:{lc}{L0 + MAX_LINHAS}")
    ws[C_SEL].font = Font(bold=True, size=9)
    ws[C_SEL_V] = val0
    ws[C_SEL_V].fill, ws[C_SEL_V].border = amarelo, borda
    ws[C_SEL_V].font = Font(bold=True, size=9)
    ws[C_SEL_V].number_format = "0.0#"
    ws[C_SEL_V].alignment = Alignment(horizontal="center", vertical="center")
    dvn = DataValidation(type="decimal", operator="between", formula1="0", formula2="5", allow_blank=True)
    dvn.error, dvn.errorTitle = "Nota geral de 0 a 5 (ex.: 3,2).", "Valor inválido"
    ws.add_data_validation(dvn)
    dvn.add(C_SEL_V)
    ws["H3"] = "Ordenar por:"
    ws["H3"].font = Font(bold=True)
    ws["H3"].alignment = Alignment(horizontal="right", vertical="center")
    campo("I3:K3", criterio)
    try:
        with open(FICH_ORD_MIN, encoding="utf-8") as f:
            ORD_MIN["v"] = float(json.load(f).get("v", 4.0))
    except Exception:
        pass
    ws[C_ORD_MIN] = ORD_MIN["v"]
    ws[C_ORD_MIN].number_format = '0.00'
    ws[C_ORD_MIN].fill = PatternFill("solid", fgColor="FFF2CC")
    ws[C_ORD_MIN].font = Font(bold=True)
    ws[C_ORD_MIN].alignment = Alignment(horizontal="center", vertical="center")
    col_l = get_column_letter(len(COLS) + 3)          # opções do "Ordenar por" (lista > 255 caracteres)
    for i, c in enumerate(CRITERIOS_ORD, 1):
        ws[f"{col_l}{i}"] = c[0]
    ws.column_dimensions[col_l].hidden = True
    dv = DataValidation(type="list", formula1=f"${col_l}$1:${col_l}${len(CRITERIOS_ORD)}", allow_blank=False)
    dv.error, dv.errorTitle = "Escolhe um valor da lista.", "Valor inválido"
    ws.add_data_validation(dv)
    dv.add(C_ORD)
    ws.merge_cells("M3:N3")
    botao(ws, "M3", "↕ ORDENAR", f"{PROTOCOLO}://ordenar/{idx}", "2E75B6", "1F4E79")
    ws[C_AT] = atualizado
    ws[C_AT].font = Font(italic=True, color="595959")
    # linha 4: ↻ PREÇOS · longo prazo mínimo · Situação em
    ws.merge_cells("A4:D4")
    botao(ws, "A4", "↻ PREÇOS (rápido)", f"{PROTOCOLO}://precos", "BF8F00", "7F6000", 11)
    texto("E4:F4", "Longo prazo mín./ano:")
    ws["E4"].font = Font(bold=True, size=9)
    ws[C_LPMIN] = (estado or {}).get("lp_min", LP_MIN)
    ws[C_LPMIN].number_format = "0.0%"
    ws[C_LPMIN].fill = PatternFill("solid", fgColor="FFF2CC")
    ws[C_LPMIN].border = Border(*(Side(style="thin", color="A6A6A6"),) * 4)
    ws[C_LPMIN].alignment = Alignment(horizontal="center", vertical="center")
    from openpyxl.comments import Comment
    nota_lp = Comment(LP_AJUDA, "Painel")
    nota_lp.width, nota_lp.height = 440, 360
    ws[C_LPMIN].comment = nota_lp
    ws.merge_cells("M4:N4")
    botao(ws, "M4", "▣ VER POSIÇÃO", f"{PROTOCOLO}://posicao", "595959", "3A3A3A", 10)
    ws["H4"] = "Situação em:"
    ws["H4"].font = Font(bold=True, size=10)
    ws["H4"].alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells("I4:J4")
    est = estado or {}
    suf = "" if ranking else "_e"
    ws[C_REF] = est.get("ref_escolha" + suf) or "Atual"
    ws[C_DATAREF] = ler_data(est.get("ref_data" + suf)) or (est.get("ref_data" + suf) or None)
    for c in (C_REF, "J4", C_DATAREF):
        ws[c].fill = PatternFill("solid", fgColor="FFF2CC")
        ws[c].border = Border(*(Side(style="thin", color="A6A6A6"),) * 4)
    ws[C_DATAREF].number_format = "dd/mm/yyyy"
    from openpyxl.formatting.rule import FormulaRule as _FR      # valores que só contam com uma escolha: cinzento
    _cz, _fz = PatternFill("solid", bgColor="808080"), Font(color="BFBFBF")
    for cel_, f_ in ((C_SEL_V, f'LEFT(${C_SEL[0]}${C_SEL[1:]},2)<>"NG"'),
                     (C_ORD_MIN, f'AND(${C_ORD[0]}${C_ORD[1:]}<>"Melhor preço com nota igual ou maior a (valor ao lado)",${C_ORD[0]}${C_ORD[1:]}<>"Melhor nota com nível de preço igual ou maior a (valor ao lado)",${C_ORD[0]}${C_ORD[1:]}<>"Melhor preço com nota global igual ou maior a (valor ao lado)",${C_ORD[0]}${C_ORD[1:]}<>"Melhor nota global com nível de preço igual ou maior a (valor ao lado)")'),
                     (C_DATAREF, f'${C_REF[0]}${C_REF[1:]}<>"Data (escrever ao lado)"'),
                     (C_LPMIN, f"LOWER(TRIM('{FOLHA_PESOS}'!$E$38))=\"x\"")):     # G4 não conta com o fator
        ws.conditional_formatting.add(cel_, _FR(formula=[f_], fill=_cz, font=_fz))
    dvr = DataValidation(type="list", formula1='"' + ",".join(REFS) + '"', allow_blank=False)
    ws.add_data_validation(dvr)
    dvr.add(C_REF)
    # linha 5: bolsas
    for ini, fim, rot, m, cor, esc in (("A", "D", "◎ LISTA BASE", "base", "7F7F7F", "595959"),
                                       ("E", "F", "€ PRIVILEGIAR EUR", "eur", "2E75B6", "1F4E79"),
                                       ("G", "H", "$ PRIVILEGIAR EUA", "eua", "C55A11", "843C0C"),
                                       ("I", "L", "⊕ TODAS AS BOLSAS", "todos", "548235", "375623")):
        ws.merge_cells(f"{ini}5:{fim}5")
        botao(ws, f"{ini}5", rot + ("  ✓" if m == modo_atual() else ""), f"{PROTOCOLO}://modo/{m}", cor, esc, 9)
    ws[C_PRECOS_EST] = "↻ PREÇOS: só os preços e as contas que dependem deles (segundos); ▶ ATUALIZAR: tudo (3 a 5 min)."
    ws[C_PRECOS_EST].font = Font(italic=True, size=9, color="595959")
    ws[C_ORD_EST].font = Font(italic=True, size=9, color="595959")

    let = get_column_letter
    cabs = CAB_RANK if ranking else CAB_ETFS
    larguras = [largura(w, fmt) for (_, _, w, fmt, _) in COLS]
    # cabeçalhos ao alto nas colunas de números (e nas de x): deixam as colunas bem mais estreitas
    import math
    ALTO = {"pct": 6.5, "nota": 5, "pont": 5, "num1": 6, "int": 6, "qtd": 6.5}
    vertical = set()
    for j, ((k, _, w, fmt, _), h) in enumerate(zip(COLS, cabs)):
        estreita = ALTO.get(fmt) if k not in ("pcompra",) else None
        if k in ("sel", "det", "corr", "desat", "so_caro"):
            estreita = larguras[j]
        if j < 15:                                      # colunas A a O: têm os botões e listas do topo por cima;
            estreita = larguras[j]                      # não ficam mais estreitas (só o texto vai ao alto se ajudar)
            if math.ceil(len(h) / max((estreita - 1) * 1.3, 1)) > 4:
                vertical.add(j)
            continue
        if estreita and math.ceil(len(h) / max((estreita - 1) * 1.3, 1)) > 2:
            vertical.add(j)
            larguras[j] = min(larguras[j], estreita)
    H_ALTO = 47                                         # altura máxima do cabeçalho (pontos), mesmo que algum
    for j in vertical:                                  # título não caiba todo; as colunas com texto ao alto
        n_lin = math.ceil(len(cabs[j]) * 4.4 / (H_ALTO - 10))   # alargam um pouco (no máximo 7,5)
        larguras[j] = max(larguras[j], min(2 * n_lin + 1.5, 7.5))
    for j, ((k, _, w, fmt, grupo), h) in enumerate(zip(COLS, cabs), start=1):
        c = ws.cell(L0, j, h)
        c.font = Font(bold=True, color="FFFFFF", size=8)
        c.fill = PatternFill("solid", fgColor=cor_cabecalho(grupo))
        c.alignment = (Alignment(horizontal="center", vertical="center", wrap_text=True, text_rotation=90)
                       if j - 1 in vertical else Alignment(horizontal="center", vertical="center", wrap_text=True))
        ws.column_dimensions[let(j)].width = larguras[j - 1]
    ws.column_dimensions["A"].width = 3                 # coluna «Atualizar (x)»: estreita (texto ao alto)
    for k_, w_ in (("valor_eur", 10), ("bolsa", 12)):    # mais estreitas
        if k_ in COL:
            ws.column_dimensions[let(COL[k_])].width = w_
    # altura do cabeçalho: chega para o texto mais comprido (letra 8: ≈ 1,3 caracteres por unidade de largura na
    # horizontal; ao alto ≈ 4,4 pontos por carácter, com as linhas que couberem na largura)
    linhas_cab = max([math.ceil(len(h) / max((w - 1) * 1.3, 1)) for j, (h, w) in enumerate(zip(cabs, larguras))
                      if j not in vertical] or [2])
    alto = max([len(h) * 4.4 / max(1, math.floor(w * 5.25 / 10.5)) for j, (h, w) in enumerate(zip(cabs, larguras))
                if j in vertical] or [0])
    ws.row_dimensions[L0].height = min(max(11 * max(linhas_cab, 2) + 4, alto + 8), H_ALTO)
    for i, d in enumerate(linhas, start=1):
        r = L0 + i
        for j, (k, _, _, fmt, grp) in enumerate(COLS, start=1):
            c = ws.cell(r, j, link_nome(vazio_esp(d.get(k)), r) if k == "nome" else vazio_esp(d.get(k)))
            _fmt_cel_base(c, k, fmt)
    n = L0 + max(len(linhas), 1)
    faixa = lambda k: f"{let(COL[k])}{L0+1}:{let(COL[k])}{n}"
    from openpyxl.styles.differential import DifferentialStyle
    from openpyxl.formatting.rule import Rule
    from openpyxl.styles.numbers import NumberFormat
    gm, ga = let(COL["moeda"]), let(COL["moeda_alt"])
    k_preco = [k for k, _, _, f, _ in COLS if f == "preco"]
    k_alt = [k for k, _, _, f, _ in COLS if f == "alt"]
    moedas = sorted({txt(d.get("moeda")) for d in linhas} | {txt(d.get("moeda_alt")) for d in linhas} - {""})
    for i, m in enumerate(moedas):                    # € 123.00 (símbolo à esquerda), conforme a moeda da linha
        for ks, col, alt in ((k_preco, gm, False), (k_alt, ga, True)):
            dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=300 + 2 * i + alt, formatCode=fmt_moeda(m, alt)))
            regra = Rule(type="expression", dxf=dxf, formula=[f'${col}{L0+1}="{m}"'])
            ws.conditional_formatting.add(" ".join(faixa(k) for k in ks), regra)
    from openpyxl.formatting.rule import CellIsRule
    for k in ("lp_eur", "lp_pct"):    # carteira: lucro a verde, prejuízo a vermelho
        ws.conditional_formatting.add(faixa(k), CellIsRule(operator="greaterThan", formula=["0"],
                                                           font=Font(color="00A050", bold=True),
                                                           fill=PatternFill("solid", bgColor="E2EFDA")))
        ws.conditional_formatting.add(faixa(k), CellIsRule(operator="lessThan", formula=["0"],
                                                           font=Font(color="C00000", bold=True),
                                                           fill=PatternFill("solid", bgColor="FCE4D6")))
    for k in ("pot", "justo_pct"):   # vermelho = abaixo do preço atual, verde = acima
        ws.conditional_formatting.add(faixa(k), ColorScaleRule(
            start_type="num", start_value=-0.2, start_color=VERMELHO, mid_type="num", mid_value=0, mid_color=BRANCO,
            end_type="num", end_value=0.3, end_color=VERDE))
    from openpyxl.formatting.rule import CellIsRule   # preço do valor justo: verde se acima do atual, vermelho se abaixo
    ws.conditional_formatting.add(faixa("justo"), CellIsRule(operator="greaterThan", formula=[f"{let(COL['preco'])}{L0+1}"],
                                                             font=Font(color="00A050", bold=True)))
    ws.conditional_formatting.add(faixa("justo"), CellIsRule(operator="lessThan", formula=[f"{let(COL['preco'])}{L0+1}"],
                                                             font=Font(color="C00000", bold=True)))
    from openpyxl.formatting.rule import FormulaRule
    # escalas verde-branco-vermelho, cada coluna com a sua própria escala (mediana = branco)
    verde_em_baixo = dict(start_type="percentile", start_value=5, start_color=VERDE, mid_type="percentile",
                          mid_value=50, mid_color=BRANCO, end_type="max", end_color=VERMELHO)
    verde_em_cima = dict(start_type="min", start_color=VERMELHO, mid_type="percentile", mid_value=50,
                         mid_color=BRANCO, end_type="percentile", end_value=95, end_color=VERDE)
    dc = let(COL["decisao"])
    sc, nvl = let(COL["so_caro"]), let(COL["nivel_preco"])
    # caras: nota geral e nível de preço em azuis (claro → escuro conforme o nível); decisão em azul claro
    for lim, cor, fcor in ((1.25, "1F4E79", "FFFFFF"), (1.5, "2F75B5", "FFFFFF"), (1.75, "9BC2E6", "1F4E79"),
                           (99, "BDD7EE", "1F4E79")):
        for k in ("nota",):
            ws.conditional_formatting.add(faixa(k), FormulaRule(
                formula=[f"AND(${sc}{L0+1}=1,${nvl}{L0+1}<={lim})"], fill=PatternFill("solid", bgColor=cor),
                font=Font(color=fcor, bold=True), stopIfTrue=True))
    # nível de preço: uma cor por rótulo (vermelho → branco → verde), para o número, o nome e a cor baterem certo
    # nível de preço com o texto ao lado do número (1,2 Muito cara … 4,5 Muito barata)
    from openpyxl.styles import PatternFill as _PF
    for i, (cond, rot) in enumerate(rotulos_nivel()):
        dxf = DifferentialStyle(numFmt=NumberFormat(numFmtId=360 + i, formatCode=f'0.0" {rot}"'),
                                fill=_PF(bgColor=COR_NIVEL[rot], fill_type="solid"),
                                font=Font(bold=True, color="FFFFFF" if "muito" in rot else "1C1E21"))
        ws.conditional_formatting.add(faixa("nivel_preco"), Rule(type="expression", dxf=dxf,
                                                                 formula=[cond.format(c=f"${nvl}{L0+1}")]))
    bl = let(COL["beta"])                # beta: verde defensiva, branco ≈ mercado, vermelhos amplifica
    for cond, cor, fc in ((f"${bl}{L0+1}<0.8", "C6E5B3", "1C1E21"), (f"${bl}{L0+1}>1.5", "F8A5A6", "9C0006"),
                          (f"${bl}{L0+1}>1.2", "FCE4D6", "1C1E21")):
        ws.conditional_formatting.add(faixa("beta"), FormulaRule(formula=[f"AND(ISNUMBER(${bl}{L0+1}),{cond})"],
                                      fill=PatternFill("solid", bgColor=cor), font=Font(color=fc), stopIfTrue=True))
    ws.conditional_formatting.add(faixa("decisao"), FormulaRule(       # compra em espera: lilás (o azul é «cara»)
        formula=[f'ISNUMBER(SEARCH(": esperar",{dc}{L0+1}))'], fill=PatternFill("solid", bgColor=COR_ESPERA),
        font=Font(color="5B2C83", bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(faixa("decisao"), FormulaRule(
        formula=[f'ISNUMBER(SEARCH("saúde frágil",{dc}{L0+1}))'], fill=PatternFill("solid", bgColor="F8CBAD"),
        font=Font(color="C00000", bold=True), stopIfTrue=True))
    gl, pl = let(COL["nota"]), let(COL["pont"])
    ws.conditional_formatting.add(faixa("decisao"), FormulaRule(       # decisão: venda pelo sinal a vermelho
        formula=[f"${pl}{L0+1}=1"], fill=PatternFill("solid", bgColor="F8696B"), font=Font(bold=True), stopIfTrue=True))
    ws.conditional_formatting.add(faixa("decisao"), FormulaRule(       # não rent. longo prazo (sinal 2.5): laranja
        formula=[f"${pl}{L0+1}=2.5"], fill=PatternFill("solid", bgColor="F4B183"), stopIfTrue=True))
    ws.conditional_formatting.add(faixa("decisao"), FormulaRule(       # manter (sinal 2): amarelo
        formula=[f"AND(ISNUMBER(${pl}{L0+1}),${pl}{L0+1}<3)"], fill=PatternFill("solid", bgColor="FFEB9C"),
        stopIfTrue=True))
    for lim, cor in ((4.5, "63BE7B"), (3.75, "A9D18E"), (2.75, "E2EFDA"), (2.25, "F4B183"), (1.5, "FFEB9C"), (-99, "F8696B")):
        for k in ("nota",) if "nota_geral" in COL else ("nota", "decisao"):   # nota geral: as cores de sempre
            ws.conditional_formatting.add(faixa(k), FormulaRule(
                formula=[f"AND(ISNUMBER(${gl}{L0+1}),${gl}{L0+1}>={lim})"], fill=PatternFill("solid", bgColor=cor),
                font=Font(bold=lim >= 4.5 or k == "nota"), stopIfTrue=True))
    if "nota_geral" in COL:
        ngl = let(COL["nota_geral"])
        # limites: ações com nível de preço 3,95 / 3,5 / 3; ETFs e ações sem nível de preço 4,5 / 3,75 / 2,75
        for i_, cor in enumerate(("63BE7B", "A9D18E", "E2EFDA", "F4B183", "FFEB9C", "F8696B")):
            la, le = (LIM_ACOES + (2.25, 1.5, -99))[i_], (LIM_ETFS + (2.25, 1.5, -99))[i_]
            lim = f"IF(ISNUMBER(${nvl}{L0+1}),{la},{le})" if la != le else la   # ETFs com nível (das ações) = ações
            ws.conditional_formatting.add(faixa("nota_geral"), FormulaRule(
                formula=[f"AND(ISNUMBER(${ngl}{L0+1}),${ngl}{L0+1}>={lim})"], fill=PatternFill("solid", bgColor=cor),
                font=Font(bold=True), stopIfTrue=True))
        for txt_, cor, b_ in (("Compra forte", "63BE7B", True), ("Compra média", "A9D18E", False),
                              ("Compra", "E2EFDA", False), ("Manter", "FFEB9C", False)):
            ws.conditional_formatting.add(faixa("decisao"), FormulaRule(   # V148: a cor da decisão segue o texto
                formula=[f'LEFT({dc}{L0+1},{len(txt_)})="{txt_}"'], fill=PatternFill("solid", bgColor=cor),
                font=Font(bold=b_), stopIfTrue=True))
    for v, cor in ((5, "63BE7B"), (4, "A9D18E"), (3, "E2EFDA"), (2.5, "F4B183"), (2, "FFEB9C"), (1, "F8696B")):
        for k in ("pont",):                           # pontuação do sinal: 5 verde-escuro … 1 vermelho
            ws.conditional_formatting.add(faixa(k), FormulaRule(
                formula=[f"${pl}{L0+1}={v}"], fill=PatternFill("solid", bgColor=cor),
                font=Font(bold=v in (5, 1)), stopIfTrue=True))
    nt = let(COL["notas"])
    ws.conditional_formatting.add(faixa("notas"), FormulaRule(
        formula=[f'ISNUMBER(SEARCH("VENDA:",{nt}{L0+1}))'], font=Font(color="C00000", bold=True)))
    pc = let(COL["pe"])                               # P/E: verde barato … vermelho caro
    for f, cor in ((f"AND(ISNUMBER({pc}{L0+1}),{pc}{L0+1}<=15)", "C6EFCE"),
                   (f"AND(ISNUMBER({pc}{L0+1}),{pc}{L0+1}>15,{pc}{L0+1}<=25)", "E2EFDA"),
                   (f"AND(ISNUMBER({pc}{L0+1}),{pc}{L0+1}>25,{pc}{L0+1}<=40)", "FFEB9C"),
                   (f"AND(ISNUMBER({pc}{L0+1}),{pc}{L0+1}>40)", "F8CBAD")):
        ws.conditional_formatting.add(faixa("pe"), FormulaRule(formula=[f], fill=PatternFill("solid", bgColor=cor)))
    lc = let(COL["lp_txt"])
    if "venda_min_compra" in COL:                     # rede face à compra: verde se ainda ganhas, vermelho se perdes
        vc = let(COL["venda_min_compra"])
        ws.conditional_formatting.add(faixa("venda_min_compra"), FormulaRule(
            formula=[f"AND(ISNUMBER({vc}{L0+1}),{vc}{L0+1}>=0)"], font=Font(color="1E7B34", bold=True)))
        ws.conditional_formatting.add(faixa("venda_min_compra"), FormulaRule(
            formula=[f"AND(ISNUMBER({vc}{L0+1}),{vc}{L0+1}<0)"], font=Font(color="C00000", bold=True)))
    sv = let(COL["saude"])                            # saúde: alerta a vermelho, OK a verde
    ws.conditional_formatting.add(faixa("saude"), FormulaRule(
        formula=[f'LEFT({sv}{L0+1},6)="ALERTA"'], fill=PatternFill("solid", bgColor="F8CBAD"),
        font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(faixa("saude"), FormulaRule(
        formula=[f'LEFT({sv}{L0+1},2)="OK"'], fill=PatternFill("solid", bgColor="C6EFCE")))
    ea = let(COL["etf_acoes"])                        # nota ponderada das ações do ETF: as cores da nota geral
    for lim, cor in ((4.5, "63BE7B"), (3.75, "A9D18E"), (2.75, "E2EFDA"), (2.25, "F4B183"), (1.5, "FFEB9C"), (-99, "F8696B")):
        ws.conditional_formatting.add(faixa("etf_acoes"), FormulaRule(
            formula=[f"AND(ISNUMBER(${ea}{L0+1}),${ea}{L0+1}>={lim})"], fill=PatternFill("solid", bgColor=cor),
            font=Font(bold=lim >= 4.5), stopIfTrue=True))
    for k, ini in (("qual", 0), ("saude_n", 0)):           # 0 (ou 1) vermelho … 5 verde
        ws.conditional_formatting.add(faixa(k), ColorScaleRule(
            start_type="num", start_value=ini, start_color=VERMELHO, mid_type="num", mid_value=(ini + 5) / 2,
            mid_color=BRANCO, end_type="num", end_value=5, end_color=VERDE))
    sl = let(COL["sinal"])
    for chave, cor in (("compra", "C6EFCE"), ("venda", "F8CBAD")):
        ws.conditional_formatting.add(faixa("sinal"), FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{chave}",{sl}{L0+1}))'], fill=PatternFill("solid", bgColor=cor)))
    cm = let(COL["coment"])
    ws.conditional_formatting.add(faixa("coment"), FormulaRule(
        formula=[f'OR(ISNUMBER(SEARCH("próxima",{cm}{L0+1})),ISNUMBER(SEARCH("rede de segurança",{cm}{L0+1})))'],
        fill=PatternFill("solid", bgColor="FFEB9C")))
    for k, lim in (("var1d", 0.04), ("var1m", 0.15)):   # subidas a verde, descidas a vermelho
        ws.conditional_formatting.add(faixa(k), ColorScaleRule(
            start_type="num", start_value=-lim, start_color=VERMELHO, mid_type="num", mid_value=0, mid_color=BRANCO,
            end_type="num", end_value=lim, end_color=VERDE))
    for i in range(len(PERIODOS)):
        ws.conditional_formatting.add(faixa(f"dist{i}"), ColorScaleRule(**verde_em_baixo))   # mais longe do máx. = verde
        ws.conditional_formatting.add(faixa(f"cima{i}"), ColorScaleRule(**verde_em_cima))    # mais positivo = verde
        ws.conditional_formatting.add(faixa(f"baixo{i}"), ColorScaleRule(**verde_em_baixo))  # mais negativo = verde
    # linhas alternadas a cinzento claro (uma sim, uma não). Conta só as linhas visíveis, por isso continua certo
    # depois de filtrar ou ordenar; nas colunas com escala de cores, as cores da escala ficam por cima.
    from openpyxl.formatting.rule import FormulaRule
    tudo = f"A{L0+1}:{let(len(COLS))}{n}"
    cd, cn, cq, cs = let(COL["det"]), let(COL["nome"]), let(COL["qtd"]), let(COL["desat"])
    par = f"MOD(SUBTOTAL(103,${cn}${L0+1}:${cn}{L0+1}),2)=0"
    tem = (f'OR(LOWER(${cd}{L0+1})="x",AND(ISNUMBER(${cd}{L0+1}),${cd}{L0+1}>0),'
           f'AND(ISNUMBER(${cq}{L0+1}),${cq}{L0+1}>0))')
    x_nome = f"{cd}{L0+1}:{cd}{n} {cn}{L0+1}:{cn}{n}"
    ws.conditional_formatting.add(f"{cn}{L0+1}:{cn}{n}", FormulaRule(   # nomes (com ligação): sempre a preto
        formula=[f"LEN(${cn}{L0+1})>=0"], font=Font(color="000000")))   # (o Excel põe as ligações a azul)
    ws.conditional_formatting.add(x_nome, FormulaRule(formula=[tem], font=Font(bold=True)))   # tens: x e nome a negrito
    ws.conditional_formatting.add(x_nome, FormulaRule(                 # tens: x e nome a amarelo; mais escuro
        formula=[f'AND({tem},{par})'], fill=PatternFill("solid", bgColor="FFD966"), stopIfTrue=True))
    ws.conditional_formatting.add(x_nome, FormulaRule(                 # nas linhas que seriam cinzentas
        formula=[tem], fill=PatternFill("solid", bgColor="FFE699"), stopIfTrue=True))
    ws.conditional_formatting.add(tudo, FormulaRule(                   # não atualizadas na última atualização parcial:
        formula=[f'AND(${cs}{L0+1}=1,{par})'], fill=PatternFill("solid", bgColor="F4B6B6"), stopIfTrue=True))
    ws.conditional_formatting.add(tudo, FormulaRule(                   # vermelho claro / rosa claro em vez dos cinzentos
        formula=[f'${cs}{L0+1}=1'], fill=PatternFill("solid", bgColor="FCE4EC"), stopIfTrue=True))
    if not PESOS.get("usar_analistas", 1):            # V161: analistas fora do sinal: colunas a cinzento, só indicativas
        for k_ in ("alvo", "alvo_alt", "pot", "rec"):
            if k_ in COL:
                cl_ = let(COL[k_])
                ws.conditional_formatting.add(f"{cl_}{L0+1}:{cl_}{n}", FormulaRule(
                    formula=["TRUE"], fill=PatternFill("solid", bgColor="BFBFBF"),
                    font=Font(color="595959", italic=True), stopIfTrue=True))
                ws.cell(L0, COL[k_]).fill = PatternFill("solid", fgColor="808080")
        k0 = min((COL[k_] for k_ in ("alvo", "alvo_alt", "pot", "rec") if k_ in COL), default=None)
        if k0:
            c_ = ws.cell(L0 - 1, k0, "(Só indicativa: não conta para o sinal)")
            c_.font, c_.alignment = Font(bold=True, italic=True, size=8, color="595959"), Alignment(horizontal="left")
    ws.conditional_formatting.add(tudo, FormulaRule(formula=[par], fill=PatternFill("solid", bgColor="D9D9D9")))
    # colunas sem uso nesta folha: na Ações, as dos ETFs; na ETFs, as que nenhum ETF preenche
    for j, (k, *_) in enumerate(COLS, start=1):
        vazia = all(not txt(d.get(k)) for d in linhas)
        if k in OCULTAS or (ranking and k in SO_ETF) or (not ranking and vazia and k not in ("sel", "det", "corr", "nome", "ticker", "preco", "qtd", "pcompra",
                                                                                    "valor_eur", "lp_eur", "lp_pct")):
            ws.column_dimensions[let(j)].hidden = True
    ws.freeze_panes = None                         # V190: sem painéis fixos
    ws.auto_filter.ref = f"A{L0}:{let(len(COLS))}{n}"
    ws[f"B{n+2}"] = ("Fonte: Yahoo Finance. Cada título na sua moeda (Londres em pence). "
                     "Preço atual = o mais recente no momento do ▶ ATUALIZAR (fica fixo até à atualização seguinte), "
                     "incluindo a pré-abertura e o pós-fecho dos EUA; a coluna antes dele diz de que sessão e de que hora é. "
                     "Variação último dia = preço atual face ao fecho anterior (na pré-abertura, face ao último fecho); "
                     "variação último mês = último fecho face ao de há 30 dias. "
                     "Valor transacionado = média dos últimos 3 meses de volume × fecho, convertida em euros (milhões); "
                     "capitalização em euros (mil milhões); volume = nº médio de títulos negociados por dia. "
                     "Máx./mín. = máximos e mínimos diários dos "
                     "últimos 7, 30, 91, 365 e 1095 dias. Média = média dos fechos diários do período. "
                     "Oscilação acima / abaixo da média = máx. / média − 1 e mín. / média − 1. "
                     "Distância ao máx. = preço atual / máximo do período − 1. "
                     "NOTA GLOBAL = nota geral + nível de preço (2 : 1; decide a compra). NOTA GERAL (percentagens na folha Ponderações). Ações: sinal + qualidade + saúde + tendência de longo prazo. ETFs: sinal "
                     "+ nota ponderada das ações do ETF (as maiores posições que o Yahoo dá, até 10, cada uma com a sua nota "
                     "global da folha Ações, pesada pelo peso no ETF; sem composição no Yahoo usa-se um ETF americano do mesmo "
                     "índice, indicado na composição; obrigações e ouro: só a pontuação). Se a qualidade ou a saúde "
                     "faltarem (ETFs, bancos, sem dados), conta a pontuação no lugar delas; com a saúde em ALERTA, a nota geral não "
                     "passa de 2,5, e o mesmo com o longo prazo fraco. A pontuação (e a decisão) continua a ser só o sinal testado: médias, valor justo e analistas. "
                     "QUALIDADE DO NEGÓCIO (0-5) = 5 × média de 3 critérios, cada um de 0 a 1 em linha reta entre mau e bom: ROIC "
                     "5% → 15% (quanto rende cada euro investido no negócio), lucro bruto / ativo total 10% → 30% (poder de preço) e "
                     "FCF / lucro líquido 0,5 → 0,9 (o lucro é dinheiro a sério); precisa de 2 dos 3. O preço (FCF yield, EV/EBITDA) "
                     "fica fora, porque já está na pontuação; só se mostra nos comentários e nas últimas colunas. "
                     "SAÚDE FINANCEIRA (0-5) = 5 × média de: dívida líquida / EBITDA 3× → 1× (utilities e imobiliárias 5× → 2×; com "
                     "mais caixa do que dívida = 1), EBIT / juros 3× → 8× (sem juros = 1) e FCF (positivo = 1, negativo com caixa "
                     "líquida = 0,5, negativo com dívida = 0). ALERTA se a dívida passar o limite mau, se o EBIT não cobrir 3× os "
                     "juros, se o FCF for negativo com dívida ou se o EBITDA for negativo com dívida; uma compra em alerta aparece "
                     "como \"Saúde frágil\". Bancos, seguradoras e financeiras: não se aplica. Dados do Yahoo (últimos 12 meses e "
                     "último relatório anual), indicadores nas últimas colunas; numa data de referência passada são os de hoje; o "
                     "↻ PREÇOS recalcula a nota geral mas não os indicadores. Na mesma nota, desempate: pontuação, qualidade e a "
                     "média de analistas e valor justo. "
                     "Valor justo = lucro por ação dos últimos 12 meses × P/E habitual do título (mediana dos últimos 3 anos); sem histórico "
                     "de lucros, número de Graham. Verde = acima do preço atual, vermelho = abaixo. É um modelo simples, não uma garantia. "
                     "Os 3 níveis (valor e % face ao preço atual; σ = volatilidade mensal do título, σ diário = σ / √21): "
                     "COMPRA MÁXIMA (a ordem limite principal) = preço × exp(σ diário × (−0,75 + 0,25 × tendência dos últimos 20 dias)): "
                     "sempre abaixo do preço (≈ −1,7%), menos em subida franca e mais em queda franca; no teste é executada ~58% das "
                     "vezes em 3 dias e ~71% em 7 (se não for, e o sinal continuar de compra, comprar ao preço do momento). VENDA RAZOÁVEL = preço + 2σ: objetivo de ganho realista para 3-4 meses — referência, "
                     "não uma ordem: no teste, vender tudo num objetivo fixo rendeu menos do que esperar pelo sinal. VENDA MÍNIMA = "
                     "maior fecho dos últimos 6 meses − 4σ: rede de segurança contra uma queda brusca antes de as médias cruzarem "
                     "(no teste, usá-la como stop automático custou um pouco: é para casos extremos). "
                     "Último ano: seguir o sinal = estar comprado só com a média de 100 dias acima da de 200, nos últimos 12 meses "
                     "(começa comprado com 100; vende e volta a comprar à abertura do dia seguinte ao cruzamento; custo de 0,1% por "
                     "operação e 2% ao ano no dinheiro parado); manter = comprar e esperar. "
                     "Resultados passados não garantem os futuros. "
                     "Cores: verde nos extremos (mais longe do máximo, "
                     "maior oscilação para cima ou para baixo), branco na mediana, vermelho no lado oposto. "
                     "O preço-alvo é a média das estimativas dos analistas a 12 meses — não é uma recomendação. "
                     "Indicadores (últimas colunas): PPO = (média exponencial 12 dias / média 26 dias − 1); sinal = média 9 do PPO; "
                     "histograma = PPO − sinal (a subir = tendência a ganhar força); RSI 14 (acima de 70 = esticado, abaixo de 30 = "
                     "sobrevendido); média 200 dias = tendência longa; momentum 12 meses sem o último. "
                     "PONTUAÇÃO (g = média 100 / média 200 − 1; J = valor justo / preço − 1; A = alvo dos analistas / preço − 1): "
                     "5 COMPRA FORTE = (g ≥ 5%, J ≥ 50% e A ≥ 50%) ou (J ≥ 100%, A ≥ 50% e g > 0); muito barata (J ≥ 100% e A ≥ 50%) "
                     "com g < 0: 4 se o preço já estiver acima da média 100 (a queda parou), senão 2; 4 COMPRA MÉDIA = g > 0, J ≥ 25% e "
                     "A ≥ 25%; 3 COMPRA = g > 0, J ≥ 0 e A ≥ 10%; 1 VENDA = média 100 abaixo da 200 e preço abaixo da média 200 (a "
                     "não ser que J e A ≥ 25%), ou preço abaixo da venda mínima; 2 MANTER = o resto. Sem valor justo conta só A; "
                     "ETFs: comparados entre si pela subida dos últimos 6 meses — 5 = entre os 20% que mais subiram e g > 0; "
                     "4 = entre os 40% e g > 0; 3 = g > 0; 1 = g < 0, preço abaixo da média 200 e subida 6 meses negativa; 2 = o "
                     "resto (testado em 5 índices 2013-2023 e 5 ETFs 2016-2026: ficar com os que mais subiram em 6 meses rendeu "
                     "+3 a +4 pontos por ano face a manter todos; amostra pequena). Ações sem analistas: só a tendência (4 = g ≥ "
                     "5% e preço acima da média 200; 3 = g > 0 e preço acima da média 200; 1 = g < 0 e preço abaixo da média 200). "
                     "Comentários: compra / venda próxima = médias a menos "
                     "de 1% de cruzar e a aproximar-se (face a há 5 dias): a compra próxima já conta como subida (\"antecipada\"; "
                     "testado: cruzou em 40 dias 85-95% das vezes e rendeu o mesmo ou mais), a venda próxima limita a 2 (cruzou "
                     "75% das vezes e os meses seguintes renderam metade ou menos); \"para N\" = o que falta para subir de nível. A parte das médias (venda) foi testada; os "
                     "limiares do valor justo e dos analistas não se podem testar no passado (o Yahoo não guarda as estimativas antigas). "
                     "BOLSAS EM EUROS: ações e ETFs fora do euro usam a mesma ação numa bolsa europeia em euros (Xetra, Euronext...), "
                     "confirmada pelo preço; analistas, valor justo, volume e capitalização vêm da bolsa principal, convertidos para "
                     "euros. Botões ◎ LISTA BASE / € PRIVILEGIAR EUR / $ PRIV. EUA / ⊕ TODAS AS BOLSAS: escolhem a bolsa (base = a da "
                     "lista; EUA = a ação ou o ADR americano das empresas europeias; os ETFs europeus não têm equivalente nos EUA; "
                     "todas = uma linha por bolsa: a da lista, a europeia em euros e a americana, para comparar preços e horários) "
                     "e atualizam o painel; nos três primeiros, títulos repetidos só aparecem uma vez. As folhas Procurar seguem o "
                     "mesmo botão: € ou $ propõem a listagem nessa zona, base a com mais negócio, todas mostram cada listagem "
                     "(correspondências em "
                     + os.path.basename(FICH_EUR) + "). "
                     "Regra das médias 100/200 no teste de estratégias (21 títulos, 2010 períodos de 1996 "
                     "a 2026): quedas −12% em vez de −25% a manter, subidas +29% em vez de +40%, equilibrado +4,9% (manter +5,4%), "
                     "pior período −61% (manter −75%), ~3 operações por período; melhor do que 50/200, do que a média 200 sozinha e "
                     "do que qualquer combinação de percentagens VL/VS/CS/CL (mesmo otimizadas título a título). Não ganha a manter "
                     "em subidas longas: é um seguro contra quedas grandes. É informação, não uma recomendação.")
    ws[f"B{n+2}"].font = Font(italic=True, size=9, color="595959")


def guardar_e_abrir(wb):
    filtros = ler_filtros()                         # os teus filtros (todas as folhas), para repor no livro novo
    tmp = os.path.join(os.path.dirname(FICHEIRO), BASE + "_novo.tmp")      # .tmp: o OneDrive não sincroniza
    wb.save(tmp)                                    # primeiro numa cópia (o painel aberto não é tocado)
    if os.name == "nt":
        try:                                        # painel aberto no Excel: guarda e fecha
            Excel(FICHEIRO).ops([{"op": "fechar"}])
        except Exception:
            pass
    for tent in range(30):                          # espera que o Excel / OneDrive larguem o ficheiro (até ~30 s)
        try:
            time.sleep(1 if tent else 2)
            os.replace(tmp, FICHEIRO)
            break
        except PermissionError:
            continue
    else:
        print(f"\nO {os.path.basename(FICHEIRO)} está aberto ou bloqueado. Fecha-o no Excel e volta a correr a atualização.")
        try:
            os.remove(tmp)
        except Exception:
            pass
        input("Carrega em Enter para sair...")
        sys.exit(1)
    if os.name == "nt":
        os.startfile(FICHEIRO)
        repor_filtros(filtros)          # a janela da posição fica preparada mas só abre com ▣ VER POSIÇÃO


def registar_botao():
    """O Excel bloqueia links diretos para ficheiros .bat ("não é possível abrir o ficheiro especificado").
    Por isso o botão aponta para painelacoes://atualizar, e aqui regista-se esse endereço no Windows
    (só para este utilizador, sem precisar de administrador) a correr o Atualizar_painel.bat desta pasta.
    Volta a registar em cada execução, para acompanhar a pasta se ela mudar de sítio."""
    if os.name != "nt":
        return
    try:
        import winreg
        bat = BAT
        cmd_exe = os.environ.get("COMSPEC", r"C:\Windows\System32\cmd.exe")
        comando = f'"{cmd_exe}" /c ""{bat}" "%1""'
        pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        if os.path.exists(pyw):                        # sem janela preta: o script abre a janela de progresso
            comando = f'"{pyw}" "{os.path.abspath(__file__)}" "%1"'
        base = rf"Software\Classes\{PROTOCOLO}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base) as k:
            winreg.SetValueEx(k, "", 0, winreg.REG_SZ, "URL:Painel de acoes")
            winreg.SetValueEx(k, "URL Protocol", 0, winreg.REG_SZ, "")
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, base + r"\shell\open\command") as k:
            winreg.SetValueEx(k, "", 0, winreg.REG_SZ, comando)
    except Exception as e:
        print("Aviso: não foi possível ativar o botão ATUALIZAR do Excel:", e)


def volatilidade(Cl):
    """σ mensal: média do desvio-padrão dos retornos diários dos últimos 63 e 252 dias, × √21."""
    import numpy as np
    r = np.diff(np.log(Cl))
    if len(r) < 40:
        return None
    return float((r[-63:].std() + r[-252:].std()) / 2 * np.sqrt(21))


CURTA = 100            # sinal: média de CURTA dias contra a de 200 (testado: 100/200 foi o melhor compromisso, à frente de 50/200)
STOP_SIGMAS = None     # stop dentro do sinal: desligado (testado: 2,5σ, 3σ e 4σ pioraram o resultado das médias)
REDE_SIGMAS = 4.0      # venda mínima (rede de segurança): 4 σ (mensais) abaixo do maior fecho dos últimos 6 meses
ALVO_SIGMAS = 2.0      # venda razoável: preço + 2 σ (subida típica boa em 3-4 meses)


def sinal_estrategia(Cl, sig, datas=None, i_ini=None, i_fim=None, devolver_serie=False):
    """Sinal do painel, só com fechos (decide no fecho, executa na abertura seguinte): comprado enquanto a média curta
      (CURTA dias) estiver acima da média 200; sai quando passa abaixo da 200; volta a comprar quando passa acima.
      (Com STOP_SIGMAS: sai também se o fecho cair STOP_SIGMAS σ abaixo do maior fecho desde a compra e só volta quando
      o fecho passar acima da média curta — testado e pior, por isso está desligado.)
    Devolve dict(dentro, desde, mmc, mm200, maximo, stop) do último fecho (ou a lista de estados, devolver_serie=True)."""
    import numpy as np
    c = np.asarray(Cl, dtype=float)
    n = len(c)
    if n < 200 or not sig:
        return None
    soma = np.concatenate([[0.0], np.cumsum(c)])
    m50 = np.full(n, np.nan)          # média curta (CURTA dias)
    m200 = np.full(n, np.nan)
    m50[CURTA - 1:] = (soma[CURTA:] - soma[:-CURTA]) / CURTA
    m200[199:] = (soma[200:] - soma[:-200]) / 200
    i0 = 199 if i_ini is None else max(199, i_ini)
    i1 = n - 1 if i_fim is None else i_fim
    dentro = bool(m50[i0] > m200[i0])
    maximo, bloqueado, desde = c[i0], False, i0
    serie = []
    for k in range(i0, i1 + 1):
        alta = m50[k] > m200[k]
        if dentro:
            maximo = max(maximo, c[k])
            if not alta:
                dentro, desde = False, k
            elif STOP_SIGMAS and c[k] < maximo * (1 - STOP_SIGMAS * sig):
                dentro, desde, bloqueado = False, k, True
        else:
            if bloqueado and c[k] > m50[k] and c[k - 1] <= m50[k - 1]:
                bloqueado = False
            if alta and not bloqueado:
                dentro, desde, maximo = True, k, c[k]
        serie.append(dentro)
    if devolver_serie:
        return serie
    data = None
    if datas is not None:
        try:
            data = datas[desde]
        except Exception:
            data = None
    return dict(dentro=dentro, desde=data, mmc=float(m50[i1]), mm200=float(m200[i1]), maximo=float(maximo),
                gap5=float(m50[i1 - 5] / m200[i1 - 5] - 1) if i1 - 5 >= 199 else None,
                stop=float(c[max(0, i1 - 125):i1 + 1].max() * np.exp(-REDE_SIGMAS * sig)),   # rede: sempre
                fecho=float(c[i1]))


# Pontuação 1-5 (g = média 100 / média 200 − 1; J = valor justo / preço − 1; A = preço-alvo dos analistas / preço − 1;
# p200 = preço / média 200 − 1). Do nível mais alto para o mais baixo, fica o primeiro que se cumpre:
#   5 compra forte: (g ≥ 5% e J ≥ 50% e A ≥ 50%) ou (J ≥ 100% e A ≥ 50% e g > 0)
#     muito barata (J ≥ 100% e A ≥ 50%) mas com a média 100 abaixo da 200: 4 se o preço já estiver acima da média 100
#     (a queda parou), 2 se ainda estiver abaixo (a cair; testado: nas descidas, com o preço abaixo da média 100 os
#     6 meses seguintes renderam menos e caíram mais: +10% e −20% contra +15% e −15%)
#   4 compra média: g > 0 e J ≥ 25% e A ≥ 25%
#   3 compra:       g > 0 e J ≥ 0 e A ≥ 10%
#   1 venda (proteção, o que foi testado): g < 0 e preço abaixo da média 200, a não ser que J ≥ 25% e A ≥ 25%;
#                   ou preço abaixo da venda mínima (rede de segurança)
#   2 manter (hold): o resto
# Sem valor justo, decide só pelos analistas; sem analistas (ETFs), só pela tendência: 4 = g ≥ 5% e preço acima da
# média 200; 3 = g > 0 e preço acima da média 200; 1 = g < 0 e preço abaixo da média 200; 2 = o resto.
NIVEIS = {5: "Compra forte", 4: "Compra média", 3: "Compra", 2.5: "Não Rent. Longo Prazo", 2: "Manter (hold)",
          1: "Venda"}
P5_TEND, P5_J, P5_A, P5_J2 = 0.05, 0.50, 0.50, 1.00
P4_J, P4_A = 0.25, 0.25
P3_J, P3_A = 0.0, 0.10
P1_SALVA_J, P1_SALVA_A = 0.25, 0.25
PROXIMO = 0.01        # médias a menos de 1% de cruzar = compra / venda próxima
DIAS_INVERSAO = 30    # cruzamento das médias há menos disto (dias de calendário) = compra / venda devido a inversão
# Longo prazo: inclinação da reta de tendência das cotações mensais desde sempre (o que o Yahoo tiver; com dividendos),
# em subida média por ano. Abaixo de LP_MIN (ou negativa), o que seria compra (3 a 5) passa a 2.5 "Não Rent. Longo Prazo": uma ação que ao longo dos anos
# não sobe (p. ex. a Soitec, que já esteve muito mais alta) não é "compra forte" só porque está num bom momento.
LP_AJUDA = "FATOR DE TENDÊNCIA (folha Ponderações, ligado por omissão): a nota geral é multiplicada por um fator de cerca de 0,60 a 1,09 que sai da média ponderada das inclinações das retas das cotações (desde sempre, 10 anos, 5 anos, último ano, 6 meses; as quedas contam a dobrar). Quem sobe de forma consistente fica com 1,00; quem estagnou ou cai é penalizado aos poucos. A coluna \"Longo prazo\" mostra as retas e o fator. Com o fator ligado, G4 não conta. SEM O FATOR (sem x na folha Ponderações): " + "LONGO PRAZO MÍN./ANO (célula G4, 3% por omissão): para cada título desenha-se a linha reta que melhor acompanha as cotações mensais (com dividendos), uma desde que há dados e outra só dos últimos 5 anos. A inclinação de cada linha diz quanto o título subiu, em média, por ano (p. ex. +8%/ano). Conta a MENOR das duas. Se ficar abaixo do valor de G4 (ou for negativa), o título não tem subido de forma consistente ao longo dos anos: mesmo que esteja num bom momento, a nota fica no máximo em 2,5 (\"Não Rent. Longo Prazo\") em vez de compra. A coluna \"Longo prazo\" mostra as três subidas: desde sempre; 5 anos; 1 ano (a de 1 ano é só informativa). Mais exigente: sobe G4 (p. ex. 5%); para desligar este travão: -100%. "
LP_MIN = 0.03         # valor por defeito; o que conta é o da célula C_LPMIN (escolhido por ti)
C_LPMIN = "G4"


def ler_lp_min(v):
    """'3%', '3', 0.03 -> 0.03 (se vier vazio ou estranho, fica o valor por defeito)."""
    try:
        if isinstance(v, (int, float)):
            x = float(v)
        else:
            x = float(txt(v).replace("%", "").replace(",", ".").replace("−", "-").strip())
        return x / 100 if abs(x) >= 1 else x
    except (TypeError, ValueError):
        return 0.03
LONGO = {}            # ticker -> (subida média por ano, ano do início dos dados)


def longo_prazo(yf, tickers):
    """{ticker: (subida usada, ano inicial, desde sempre, 5 anos, 1 ano)}: retas de tendência dos fechos mensais (até à
    data de referência). A subida usada (teste do longo prazo fraco) é a MENOR das retas desde sempre e dos últimos 5
    anos; a de 1 ano só se mostra."""
    import numpy as np
    out = {}
    tickers = [t for t in dict.fromkeys(tickers) if t]
    if not tickers:
        return out
    print(f"A descarregar a história completa (mensal) de {len(tickers)} títulos, para a tendência de longo prazo...")
    try:
        h = _cortar(_por_ticker(yf.download(tickers, period="max", interval="1mo", group_by="ticker",
                                            auto_adjust=True, threads=True, progress=False), tickers))
    except Exception as e:
        print("   aviso: sem história longa:", e)
        return out
    for t in tickers:
        try:
            c = h[t]["Close"].dropna()
            c = c[c > 0]
            if len(c) < 36:                       # menos de 3 anos: não conta
                continue
            def reta(cc):
                x = np.arange(len(cc)) / 12.0
                return float(np.exp(np.polyfit(x, np.log(cc.to_numpy(dtype=float)), 1)[0]) - 1)
            s_all = reta(c)
            s5 = reta(c.iloc[-61:]) if len(c) >= 61 else None
            s1 = reta(c.iloc[-13:]) if len(c) >= 13 else None
            s10 = reta(c.iloc[-121:]) if len(c) >= 121 else None
            s3 = reta(c.iloc[-37:]) if len(c) >= 37 else None
            usado = min(s_all, s5) if s5 is not None else s_all
            out[t] = (usado, int(c.index[0].year), s_all, s5, s1, s10, s3)
        except Exception:
            continue
    return out
# "Situação em": o painel pode ser calculado como estava numa data passada (preços, médias, sinais e níveis dessa data;
# analistas e valor justo são os de hoje, porque o Yahoo não guarda os antigos). A coluna "Da data de referência até
# hoje" mostra quanto cada título andou desde então (para ver se as recomendações dessa data acertaram).
C_REF, C_DATAREF = "I4", "K4"
REFS = {"Atual": 0, "Data (escrever ao lado)": None, "1 dia": 1, "2 dias": 2, "3 dias": 3, "1 semana": 7,
        "15 dias": 15, "1 mês": 30, "3 meses": 91, "6 meses": 182, "9 meses": 273, "1 ano": 365,
        "1 ano e meio": 548, "2 anos": 730, "2 anos e meio": 913, "3 anos": 1095, "5 anos": 1826}
REF = None            # datetime da situação pedida (None = atual)


def ler_data(v):
    """Data escrita numa célula, venha como vier: data do Excel, número de série do Excel (o que o Excel aberto
    devolve), ou texto dd/mm/aaaa, dd-mm-aaaa, dd.mm.aaaa, dd/mm/aa ou aaaa-mm-dd."""
    from datetime import date
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and 1000 < v < 100000:          # número de série do Excel (dias desde 30/12/1899)
        return datetime(1899, 12, 30) + timedelta(days=int(v))
    t = txt(v)
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(t[:10].strip(), f)
        except ValueError:
            continue
    try:                                                            # p. ex. "2026-08-13 00:00:00"
        return datetime.fromisoformat(t)
    except ValueError:
        return None


def data_ref(escolha, data):
    """Data da situação pedida (None = atual)."""
    escolha = txt(escolha) or "Atual"
    if escolha.startswith("Data ("):            # painéis antigos: "Data (escrever em J5)"
        escolha = "Data (escrever ao lado)"
    if escolha not in REFS or REFS[escolha] == 0:
        return None
    if REFS[escolha] is None:
        d = ler_data(data)
        if d is None:
            raise SystemExit(f"Escreve a data em {C_DATAREF} (dd/mm/aaaa) ou escolhe outra opção em {C_REF}.")
        if d > datetime.now():
            raise SystemExit(f"A data em {C_DATAREF} ({d:%d/%m/%Y}) está no futuro: escreve uma data passada.")
        return d
    return datetime.now() - timedelta(days=REFS[escolha])


def _cortar(hist):
    """Velas só até à data de referência."""
    if REF is None:
        return hist
    out = {}
    for k, v in hist.items():
        if v is None:
            continue
        if getattr(v.index, "tz", None) is not None:
            v = v.copy()
            v.index = v.index.tz_localize(None)
        out[k] = v[v.index <= REF]
    return out
# ETFs: comparados entre si pela subida dos últimos 6 meses (os que sobem mais tendem a continuar; testado em 5 índices
# 2013-2023 e 5 ETFs 2016-2026: ficar com os 2 que mais subiram em 6 meses rendeu +3 a +4 pontos por ano face a manter
# todos, e o filtro 100/200 sozinho rendeu menos). 5 = entre os 20% que mais subiram e média 100 acima da 200;
# 4 = entre os 40% e média 100 acima da 200; 3 = média 100 acima da 200; 1 = média 100 abaixo da 200, preço abaixo da
# média 200 e subida em 6 meses negativa; 2 = o resto.
ETF_TOP5, ETF_TOP4 = 0.20, 0.40


def pontuar_etfs(linhas, preco_de=lambda d: d.get("preco")):
    """Põe em cada ETF o seu lugar pela subida em 6 meses (entre os ETFs com dados) e volta a pontuar."""
    def m6(d):
        p, r = preco_de(d), d.get("ref6m")
        return p / r - 1 if isinstance(p, (int, float)) and isinstance(r, (int, float)) and r else None
    etfs = [d for d in linhas if d.get("tipo") == "ETF"]
    com = sorted([d for d in etfs if m6(d) is not None], key=m6, reverse=True)
    for i, d in enumerate(com):
        d["_pos6"], d["_n6"] = (i / (len(com) - 1) if len(com) > 1 else 0.0), len(com)
    for d in etfs:
        p = preco_de(d)
        pontuar(d, float(p) if isinstance(p, (int, float)) else None)
        d.pop("_pos6", None)
        d.pop("_n6", None)


def pontuar(d, preco):
    """Pontuação, decisão e comentários de uma linha (com o preço dado: o do ▶ ATUALIZAR ou o do ↻ PREÇOS)."""
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    m100, m200 = n("mm100c"), n("mm200c")
    d["pont"] = d["decisao"] = d["coment"] = d["nota"] = None
    if not preco or not m100 or not m200:
        return
    g, p200 = m100 / m200 - 1, preco / m200 - 1
    J, A = n("justo_pct"), n("pot")
    if not PESOS.get("usar_analistas", 1) and txt(d.get("tipo")) != "ETF":
        A = 1e9                                    # V160: sem analistas no sinal: as condições deles contam como cumpridas
    rede = n("venda_min")          # rede de segurança: queda demasiado grande face ao máximo de 6 meses, em qualquer tendência
    so_tend = A is None
    mom6 = preco / n("ref6m") - 1 if n("ref6m") else None
    d["mom6"] = mom6
    pos6, n_etf = d.get("_pos6"), d.get("_n6")              # lugar entre os ETFs pela subida em 6 meses (0 = o 1.º)
    etf_mom = d.get("tipo") == "ETF" and mom6 is not None and pos6 is not None
    jok = lambda x: J is None or J >= x
    barata = J is not None and A is not None and J >= P5_J2 and A >= P5_A
    # compra / venda próxima: médias a menos de 1% de cruzar E a aproximar-se (face a há 5 dias). Testado (25 títulos
    # 2016-26 e 18 ETFs 2001-23): a compra próxima cruzou em 40 dias 85-95% das vezes e os 3 meses seguintes renderam
    # o mesmo ou mais do que já em subida (e caíram menos) -> conta já como subida; a venda próxima cruzou 75% das vezes
    # e os 3-6 meses seguintes renderam metade ou menos -> não se compra (no máximo 2)
    g5 = n("gap5")
    compra_prox = -PROXIMO < g < 0 and g5 is not None and g > g5
    venda_prox = 0 <= g < PROXIMO and g5 is not None and g < g5
    sobe, desce = g > 0 or compra_prox, g < 0 and not compra_prox
    # V144: arranque (só ações): conta como tendência confirmada antes de as médias cruzarem; o preço continua a contar
    # (valor justo / analistas nas condições do sinal 5 e o nível de preço na nota global)
    arr = bool(d.get("arranque")) and d.get("tipo") != "ETF"
    ge, pe = (max(g, P5_TEND), max(p200, 1e-4)) if arr else (g, p200)
    if arr:
        sobe, desce, venda_prox, compra_prox = True, False, False, False
    # V145: subida forte e regular nos últimos 60 dias (só ações, já com tendência): conta como médias ≥ 5% afastadas
    s60 = bool(d.get("sub60")) and d.get("tipo") != "ETF" and g > 0 and p200 > 0 and g < P5_TEND
    if s60:
        ge = max(ge, P5_TEND)
    # V148: recuperação (ver REC_*): hoje, ou nos últimos REC_JANELA dias com o preço ainda acima da média 50
    etf_ = txt(d.get("tipo")) == "ETF"
    rd = d.get("rec_e" if etf_ else "rec_a")
    rd = rd if isinstance(rd, (int, float)) else None
    m50_ = d.get("mm50") if isinstance(d.get("mm50"), (int, float)) else None
    acima50 = bool(m50_) and preco > m50_
    jan_ = int(PESOS.get("rec_jan_etf", REC_JANELA_ETF) if etf_ else PESOS.get("rec_jan", REC_JANELA))
    rec_hoje = recuperacao(d, preco)
    recup = rec_hoje or (rd is not None and rd <= jan_ and acima50)
    pos_rec = not recup and rd is not None and rd <= max(REC_PROTEGE, jan_) and acima50
    if recup or (pos_rec and etf_):              # recuperação (ETFs: também depois): tendência a subir
        sobe, desce, venda_prox = True, False, False
        pe = max(pe, 1e-4)
    elif pos_rec:                                # ações, depois da janela: só não vende por as médias não terem cruzado
        desce = False                            # (nos testes, comprar nesta fase não rendeu mais nas ações)
    queda_vigiar = False
    if rede and preco < rede and p200 < 0:         # queda brusca confirmada pela tendência (abaixo da média 200)
        p = 1
    elif rede and preco < rede:                    # queda brusca, mas ainda acima da média 200: vigiar, não vender
        p, queda_vigiar = 2, True
    elif etf_mom:
        p = (5 if pos6 <= ETF_TOP5 and sobe else 4 if pos6 <= ETF_TOP4 and sobe else 3 if sobe
             else 1 if (p200 < 0 and mom6 < 0) else 2)
    elif so_tend:
        p = 4 if ge >= P5_TEND and pe > 0 else 3 if sobe and pe > 0 else 1 if desce and pe < 0 else 2
    elif (ge >= P5_TEND and jok(P5_J) and A >= P5_A) or (barata and sobe):
        p = 5
    elif barata:
        p = 4 if preco > m100 else 2
    elif sobe and jok(P4_J) and A >= P4_A:
        p = 4
    elif sobe and jok(P3_J) and A >= P3_A:
        p = 3
    elif desce and pe < 0 and not (J is not None and J >= P1_SALVA_J and A >= P1_SALVA_A):
        p = 1
    else:
        p = 2
    abaixo200 = False
    if p >= 3 and pe < 0 and (etf_mom or not so_tend):      # V132: também nos ETFs
        # V130: compra só com o preço acima da média 200 (como quando só há a tendência). Barata face aos analistas
        # / valor justo, mas a cair abaixo da média 200: «Compra …: esperar» em vez de comprar já a meio da queda
        p, abaixo200 = 2, True
    if venda_prox and p > 2:
        p = 2
    d["_recup"] = None
    if recup and not (rede and preco < rede) and (so_tend or etf_mom or (jok(P3_J) and A >= P3_A)):
        p, abaixo200, d["_recup"] = 5, False, True     # com analistas: só se não estiver cara
    lp = lp_travao(d)
    lp_fraco = lp is not None and lp < LP_MIN and not PESOS.get("ft_usar", 1)   # com o fator: sem travão
    limitado = lp_fraco and p >= 3                 # seria compra, mas não rende a longo prazo -> 2.5
    alerta = txt(d.get("saude")).startswith("ALERTA")
    fragil = alerta and p >= 3                     # seria compra, mas a saúde financeira está em alerta
    if limitado:
        p = 2.5
    c = []
    if d.get("_recup"):
        c.append("RECUPERAÇÃO: caiu do máximo de 1 ano e voltou a subir com força (reta 60 dias ≥ +30%/ano, acima da "
                 "média 50" + ("" if etf_ else " e da 200") + ")" + (f", há {int(rd)} dias de bolsa" if rd else "")
                 + ": sinal 5 e o fator ignora a queda do último ano")
    elif recup:
        c.append("recuperação, mas não conta: " + ("abaixo da rede de segurança" if rede and preco < rede
                                                   else "cara face ao valor justo / analistas"))
    elif pos_rec:
        c.append(f"recuperação há {int(rd)} dias de bolsa e o preço acima da média 50: "
                 + ("conta como tendência a subir" if etf_ else "não vende por as médias ainda não terem cruzado"))
    if s60 and not arr:
        c.append(f"subida forte e regular nos últimos {SUB60_D} dias (≥ {SUB60_MIN:+.0%}/ano): conta como tendência "
                 "forte, sem esperar que as médias se afastem 5%")
    if arr and (g < P5_TEND or p200 < 0):
        c.append(f"início de subida depois de {ARR_BASE} dias parada (preço acima do máximo desse período): conta já "
                 "como tendência confirmada, antes de as médias cruzarem")
    if compra_prox:
        c.append(f"compra próxima: média 100 a {g:+.1%} da 200 e a subir (conta já como subida)")
    elif venda_prox:
        c.append(f"venda próxima: média 100 só {g:+.1%} acima da 200 e a descer (não comprar; no máximo 2)")
    elif -PROXIMO < g < 0:
        c.append(f"médias perto ({g:+.1%}) mas a afastar-se")
    elif 0 <= g < PROXIMO:
        c.append(f"médias perto ({g:+.1%}) mas a afastar-se para cima")
    if abaixo200:
        c.append(f"seria compra, mas o preço está abaixo da média 200 ({p200:+.1%}): esperar que recupere")
    if rede and preco < rede:
        c.append("preço abaixo da venda mínima (rede de segurança)" + (
            f", mas ainda acima da média 200 ({p200:+.1%}): queda brusca a vigiar, não se vende já (vende-se se "
            "descer abaixo da média 200)" if queda_vigiar else " e abaixo da média 200: queda confirmada"))
    elif g > 0 and p200 < 0:
        c.append(f"preço abaixo da média 200 ({p200:+.1%}): tendência a enfraquecer")
    elif g < 0 and p200 > 0:
        c.append(f"preço já acima da média 200 ({p200:+.1%}): pode estar a recuperar")
    falta = lambda nome, v, lim: f"{nome} ≥ {lim:+.0%} (agora {v:+.0%})" if v < lim else None
    if not so_tend and p in (2, 3, 4) and not (barata and g <= 0) and not venda_prox:
        if not sobe:
            pedir, alvo_n = [f"média 100 acima da 200 (agora {g:+.1%})"], 3
            if J is not None and J < P3_J:
                pedir.append(falta("valor justo", J, P3_J))
            if A < P3_A:
                pedir.append(falta("analistas", A, P3_A))
        elif p == 4:
            pedir, alvo_n = [falta("tendência", g, P5_TEND), falta("valor justo", J, P5_J) if J is not None else None,
                             falta("analistas", A, P5_A)], 5
        elif p == 3:
            pedir, alvo_n = [falta("valor justo", J, P4_J) if J is not None else None, falta("analistas", A, P4_A)], 4
        else:
            pedir, alvo_n = [falta("valor justo", J, P3_J) if J is not None else None, falta("analistas", A, P3_A)], 3
        pedir = [x for x in pedir if x]
        if pedir:
            c.append(f"para {alvo_n}: " + ", ".join(pedir))
    if barata and g <= 0:
        c.append(f"muito barata mas em descida: {'já' if preco > m100 else 'ainda não'} acima da média 100 "
                 f"({preco / m100 - 1:+.1%}); para 5: média 100 acima da 200 (agora {g:+.1%})")
    if p == 1 and not (rede and preco < rede) and not etf_mom:
        c.append(f"para 2: preço acima da média 200 (agora {p200:+.1%})"
                 + ("" if so_tend else f" ou valor justo e analistas ≥ +{P1_SALVA_A:.0%}"))
    if J is None and not so_tend:
        c.append("sem valor justo: conta só os analistas")
    if so_tend and d.get("tipo") != "ETF":
        c.append("sem preço-alvo dos analistas: conta só a tendência")
    if etf_mom:
        lugar = round(pos6 * (n_etf - 1)) + 1 if n_etf > 1 else 1
        c.append(f"subida 6 meses {mom6:+.1%} ({lugar}.º de {n_etf} ETFs)")
        if p in (3, 4) and g > 0:
            lim = ETF_TOP5 if p == 4 else ETF_TOP4
            c.append(f"para {p + 1}: entre os {lim:.0%} primeiros em subida 6 meses")
        elif g <= 0 and p == 2:
            c.append(f"para 3: média 100 acima da 200 (agora {g:+.1%})")
        elif p == 1:
            c.append(f"para 2: subida 6 meses positiva ou preço acima da média 200 (agora {p200:+.1%})")
    sinal = None
    if compra_prox:
        sinal = "compra próxima"
    elif venda_prox:
        sinal = "venda próxima"
    else:
        m = re.search(r"(subida|descida) desde (\d{2}/\d{2}/\d{4})", str(d.get("tend") or ""))
        if m and ((REF or datetime.now()) - datetime.strptime(m.group(2), "%d/%m/%Y")).days <= DIAS_INVERSAO:
            sinal = "compra devido a inversão" if m.group(1) == "subida" else "venda devido a inversão"
    d["sinal"] = sinal
    notas = re.sub(r"(; )?VENDA( POR PROTEÇÃO)?:[^;]*", "", str(d.get("notas") or "")).strip("; ")
    if p == 1:
        motivo = ("queda brusca confirmada: preço abaixo da venda mínima (rede de segurança) e da média 200"
                  if rede and preco < rede else
                  "tendência de descida (média 100 abaixo da 200 e preço abaixo da média 200), para evitar uma queda "
                  "prolongada e sucessiva")
        notas = (notas + "; " if notas else "") + "VENDA: " + motivo
    d["notas"] = notas or None
    if lp is not None:
        base_lp = re.sub(r":.*$", "", str(d.get("lp_txt") or f"{lp:+.1%}/ano"))
        d["lp_txt"] = (base_lp + (f": não sobe ao longo dos anos (menos de {LP_MIN:.1%}/ano) → 2.5 Não Rent. Longo Prazo"
                                  if limitado else " (fraco)" if lp_fraco else ""))
        if lp_fraco:
            c[:] = [x for x in c if not x.startswith(("para 4", "para 5"))]
        if limitado:
            c.append(f"longo prazo fraco ({lp:+.1%}/ano): 2.5 em vez de compra")
    if alerta:
        c.append("saúde financeira em alerta (" + re.sub(r"^ALERTA: | \(.*$", "", txt(d.get("saude")))
                 + f"): nota geral no máx. {NOTA_ALERTA}".replace(".", ","))
    if txt(d.get("tipo")) != "ETF":
        d["nivel_preco"] = nivel_preco(d)[0]
    cara = motivo_caro(d)
    if cara:
        c.append(cara.replace(".", ",") + " (só aviso: a azul)")
    q, sn = n("qual"), n("saude_n")
    if q is not None:
        c.append(f"qualidade {q:.1f}/5 ({detalhe_qualidade(d)})".replace(".", ",", 1))
    d["pont"], d["coment"] = p, " · ".join(c) or None
    d["_vprox"], d["_cprox"] = venda_prox, compra_prox
    d["_vigiar"] = queda_vigiar
    d["_preco_sinal"] = preco
    d["_venda_tipo"] = ("queda brusca" if rede and preco < rede else "tendência de descida") if p == 1 else None
    d["nota"] = nota_global(d)
    d["decisao"] = decisao_global(d)


def niveis(d, preco):
    """Os 4 níveis (valor e % face ao preço atual): compra razoável, compra máxima, venda razoável, venda mínima."""
    import math
    sig = d.get("sigma") if isinstance(d.get("sigma"), (int, float)) else None
    if not preco or not sig:
        return
    sd = sig / 21 ** 0.5
    t = d.get("tend20") if isinstance(d.get("tend20"), (int, float)) else 0.0
    t = max(-2.0, min(2.0, t))
    teto = preco * math.exp(sd * (TETO_BASE + TETO_TEND * t))
    d["compra"] = min(teto, preco * math.exp(-RAZOAVEL_SIGMAS * sig))
    d["teto"] = teto
    d["venda_alvo"] = preco * (1 + ALVO_SIGMAS * sig)
    stop = d.get("stop") if isinstance(d.get("stop"), (int, float)) else None
    d["venda_min"] = stop if stop else preco * math.exp(-REDE_SIGMAS * sig)   # fora: a rede se comprasses agora
    for k in ("compra", "teto", "venda_alvo", "venda_min"):
        d[k + "_pct"] = d[k] / preco - 1


def ordens(h, preco, compra):
    """Regra da média de 200 dias (estado, níveis de saída e de reentrada), volatilidade σ e teste do último ano."""
    out = {}
    if not preco or h is None or len(h) <= 60:
        return out
    sig = volatilidade(h["Close"].to_numpy(dtype=float))
    if sig:
        out["sigma"] = sig
    h = h.copy()
    for col in ("Open", "High", "Low"):
        h[col] = h[col].fillna(h["Close"])
    Cl = h["Close"].to_numpy(dtype=float)
    Op = h["Open"].to_numpy(dtype=float)
    s = sinal_estrategia(Cl, sig, list(h.index))
    if s:
        out.update(_alta=s["dentro"], _desde=s["desde"], mm100c=s["mmc"], mm200c=s["mm200"], stop=s["stop"],
                   gap5=s["gap5"])
    sdd = (sig or 0) / 21 ** 0.5
    if sdd and len(Cl) > 21:
        import math
        out["tend20"] = max(-2.0, min(2.0, math.log(Cl[-1] / Cl[-21]) / (sdd * 20 ** 0.5)))
    i0 = max(0, len(Cl) - 253)
    out["t_hold"] = Cl[-1] / Cl[i0] - 1
    if len(Cl) >= 200 + 253 and sig:                  # o sinal no último ano (médias e stop com a história anterior)
        serie = sinal_estrategia(Cl, sig, i_ini=i0 - 1, devolver_serie=True)
        out["t_m200"] = simular_sinal(Cl, Op, i0, len(Cl) - 1, serie)
    return out


def simular_sinal(Cl, Op, i0, i1, serie, custo=None, juro=None):
    """Começa comprado com 100 em i0 e segue o sinal (serie[k] = estado no fecho do dia i0 - 1 + k): compra e vende à
    abertura do dia seguinte; o dinheiro parado rende 'juro'. Devolve a rentabilidade."""
    custo = TESTE_CUSTO if custo is None else custo
    juro = TESTE_JURO if juro is None else juro
    T, Sh, jd = 0.0, 100 * (1 - custo) / Cl[i0], juro / 252
    for k in range(i0 + 1, i1 + 1):
        T *= 1 + jd
        quer = serie[k - i0]               # estado no fecho de k-1
        if Sh and not quer:
            T, Sh = T + Sh * Op[k] * (1 - custo), 0.0
        elif not Sh and quer:
            T, Sh = 0.0, T * (1 - custo) / Op[k]
    return (T + Sh * Cl[i1]) / 100 - 1


VJ_PE_HAB = [None]       # PER habitual (mediana de 3 anos) do último valor justo calculado


def valor_justo(tk, h, preco, info):
    """Valor justo = lucro por ação dos últimos 12 meses × P/E habitual do próprio título (mediana dos últimos 3 anos).
    Se não houver histórico de lucros, usa o número de Graham √(22,5 × lucro por ação × valor contabilístico por ação)."""
    import numpy as np
    import pandas as pd
    try:
        quando = pd.Timestamp(REF) if REF is not None else pd.Timestamp.now()
        atras = (pd.Timestamp.now() - quando).days // 91 if REF is not None else 0
        ed = None
        for tent in range(3):                  # o Yahoo às vezes falha este pedido: sem ele caía-se no número de Graham
            try:
                ed = tk.get_earnings_dates(limit=min(100, 24 + atras))
                if ed is not None and len(ed):
                    break
            except Exception:
                ed = None
            time.sleep(2 * (tent + 1))
        col = next(c for c in ed.columns if "Reported" in c)
        eps = ed[col].dropna().astype(float)
        eps.index = pd.to_datetime(eps.index)
        if eps.index.tz is not None:
            eps.index = eps.index.tz_localize(None)
        eps = eps[eps.index <= quando].sort_index()          # só os lucros já divulgados na data
        eps = eps[~eps.index.normalize().duplicated()]
        try:                                   # lucros antigos ainda não ajustados a desdobramentos de ações
            sp = tk.splits
            sp.index = sp.index.tz_localize(None) if sp.index.tz is not None else sp.index
            for data, fator in sp.items():
                if fator > 1:
                    antes, depois = eps[eps.index < data], eps[eps.index >= data]
                    if len(antes) >= 2 and len(depois) >= 2 and \
                            antes.iloc[-4:].abs().mean() / max(depois.iloc[:4].abs().mean(), 1e-9) > fator / 2:
                        eps[eps.index < data] = eps[eps.index < data] / fator
        except Exception:
            pass
        ttm = eps.rolling(4).sum().dropna()
        if len(ttm) >= 2 and ttm.iloc[-1] > 0:
            c = h["Close"]
            e = ttm.reindex(c.index.union(ttm.index)).ffill().reindex(c.index)
            pe = (c / e)[e > 0].dropna()
            pe = pe[(pe > 0) & (pe < 300)]
            if len(pe) > 60:
                lucro, metodo = float(ttm.iloc[-1]), "P/E"
                med5 = float(ttm.iloc[-20:].mean()) if len(ttm) >= 8 else None   # lucro médio dos últimos ~5 anos
                fwd = _num(info.get("forwardEps"))
                # lucro no pico do ciclo (opção da folha Ponderações): só se os analistas também preverem que desce
                # (lucro previsto < 90% do atual); quem cresce de forma estrutural (p. ex. NVIDIA) não é castigado
                if (PESOS.get("pico_lucro") and med5 and med5 > 0 and lucro > PICO_LUCRO * med5
                        and fwd is not None and fwd < 0.9 * lucro):
                    lucro, metodo = med5, "P/E (lucro médio de 5 anos: lucro atual no pico)"
                VJ_PE_HAB[0] = float(pe.median())
                pm = VJ_PE_HAB[0]
                if PESOS.get("cap_pe_sinal") and pm > PESOS["pe_hab_max"]:      # opção: PER habitual limitado
                    pm, metodo = PESOS["pe_hab_max"], metodo + f" (PER habitual limitado a {PESOS['pe_hab_max']:.0f})"
                return float(lucro * pm), metodo
    except Exception:
        pass
    if REF is not None:                         # numa data passada não há o lucro e o valor contabilístico dessa data
        return None, None
    eps1, bv = info.get("trailingEps"), info.get("bookValue")
    if eps1 and bv and eps1 > 0 and bv > 0 and preco:
        g = float(np.sqrt(22.5 * eps1 * bv))
        if g < preco / 20:
            g *= 100                            # Londres: lucros em libras, preço em pence
        elif g > preco * 20:
            g /= 100
        if preco / 20 < g < preco * 20:
            return g, "Graham"
    return None, None


def alvo_em(tk, data):
    """Preço-alvo médio dos analistas numa data passada, a partir do histórico de alterações do Yahoo
    (upgrades_downgrades): para cada corretora, o último preço-alvo dado nos 12 meses antes da data. (alvo, nº)."""
    import pandas as pd
    try:
        ud = tk.upgrades_downgrades
        if ud is None or not len(ud):
            return None, None
        ud = ud.copy()
        if "GradeDate" in ud.columns:
            ud = ud.set_index("GradeDate")
        ud.index = pd.to_datetime(ud.index)
        if ud.index.tz is not None:
            ud.index = ud.index.tz_localize(None)
        col = next((c for c in ud.columns if c.lower() in ("currentpricetarget", "pricetarget")), None)
        firma = next((c for c in ud.columns if c.lower() == "firm"), None)
        if not col or not firma:
            return None, None
        t = pd.Timestamp(data)
        w = ud[(ud.index <= t) & (ud.index > t - pd.Timedelta(days=365))]
        w = w[pd.to_numeric(w[col], errors="coerce") > 0].sort_index()
        if not len(w):
            return None, None
        ult = w.groupby(firma)[col].last().astype(float)
        if len(ult) < 2:
            return None, None
        return float(ult.mean()), int(len(ult))
    except Exception:
        return None, None


def _num(v):
    try:
        v = float(v)
        return None if v != v else v      # NaN -> None
    except (TypeError, ValueError):
        return None


def taxas_eur(moedas):
    """Câmbio para euros de cada moeda do Yahoo: (taxa da moeda principal, taxa do preço). Em Londres (GBp), Joanesburgo
    (ZAc) e Telavive (ILA) o preço vem em cêntimos, mas a capitalização vem na moeda principal."""
    import yfinance as yf
    menor = {"GBp": "GBP", "ZAc": "ZAR", "ILA": "ILS"}
    pedir = sorted({menor.get(m, m) for m in moedas if m and menor.get(m, m) != "EUR"})
    tx = {"EUR": 1.0}
    if pedir:
        try:
            h = yf.download([f"{m}EUR=X" for m in pedir], period="10d", progress=False, auto_adjust=True, threads=True)
            fecho = h["Close"]
            for m in pedir:
                col = f"{m}EUR=X"
                if col in fecho and fecho[col].dropna().size:
                    tx[m] = float(fecho[col].dropna().iloc[-1])
        except Exception as e:
            print("Aviso: sem câmbios para euros:", e)
    out = {}
    for m in moedas:
        t = tx.get(menor.get(m, m)) if m else None
        out[m] = (t, t / 100 if (t and m in menor) else t)
    return out


SESSOES = {"pre": "pré-abertura", "regular": "sessão normal", "post": "pós-fecho"}


def preco_mais_recente(info, ultimo):
    """O preço mais recente que o Yahoo tem: sessão normal, pré-abertura ou pós-fecho (EUA), o que for mais recente.
    Devolve (preço, "sessão dd/mm hh:mm", fecho de referência para a variação do dia)."""
    cand = []
    for chave, preco, hora in (("regular", info.get("regularMarketPrice") or info.get("currentPrice"), info.get("regularMarketTime")),
                               ("pre", info.get("preMarketPrice"), info.get("preMarketTime")),
                               ("post", info.get("postMarketPrice"), info.get("postMarketTime"))):
        if _num(preco):
            cand.append((_num(hora) or 0, chave, float(preco)))
    if not cand:
        return ultimo, ("último fecho" if ultimo else None), None
    hora, sessao, preco = max(cand)                     # o mais recente
    texto = SESSOES[sessao] + (f" {datetime.fromtimestamp(hora):%d/%m %H:%M}" if hora else "")
    # variação do dia: pré-abertura face ao último fecho; sessão normal e pós-fecho face ao fecho anterior
    ref = info.get("regularMarketPrice") if sessao == "pre" else info.get("regularMarketPreviousClose")
    return preco, texto, _num(ref)


def indicadores(fechos):
    """PPO 12/26 com sinal 9 e histograma, variação do histograma em 5 dias, RSI 14 (Wilder), média de 200 dias,
    preço / média 200 dias e momentum 12 meses sem o último mês. Tudo a partir dos fechos diários."""
    c = fechos.dropna().astype(float)
    out = {}
    if len(c) < 40:
        return out
    e12, e26 = c.ewm(span=12, adjust=False).mean(), c.ewm(span=26, adjust=False).mean()
    ppo = e12 / e26 - 1
    sinal = ppo.ewm(span=9, adjust=False).mean()
    hist = ppo - sinal
    out.update(ppo=float(ppo.iloc[-1]), ppo_sinal=float(sinal.iloc[-1]), ppo_hist=float(hist.iloc[-1]),
               ppo_dh=float(hist.iloc[-1] - hist.iloc[-6]) if len(hist) > 6 else None)
    dif = c.diff()
    ganho = dif.clip(lower=0).ewm(alpha=1 / 14, adjust=False).mean()
    perda = (-dif.clip(upper=0)).ewm(alpha=1 / 14, adjust=False).mean()
    if perda.iloc[-1] > 0:
        out["rsi"] = float(100 - 100 / (1 + ganho.iloc[-1] / perda.iloc[-1]))
    elif ganho.iloc[-1] > 0:
        out["rsi"] = 100.0
    if len(c) >= 200:
        mm = float(c.iloc[-200:].mean())
        out.update(mm200=mm, dist200=float(c.iloc[-1] / mm - 1))
    if len(c) >= 253:
        out["mom12"] = float(c.iloc[-22] / c.iloc[-253] - 1)
    return out


# ------------------------------------------------------------------------------------------------ bolsa preferida
# Três modos (botões em cada folha; o escolhido fica guardado em FICH_MODO):
#   base = cada título na bolsa da lista (sem trocas)
#   eur  = de preferência numa bolsa europeia em euros (Xetra, Euronext...)
#   eua  = de preferência numa bolsa americana (ações europeias: a ação ou o ADR nos EUA; os ETFs UCITS não têm
#          equivalente nos EUA e ficam como estão)
# Uma listagem só é aceite se o preço bater com o da bolsa principal (convertido; nos ADR, a um múltiplo de ações) e
# se as variações semanais acompanharem as da principal. Os analistas, o valor justo, o volume e a capitalização vêm
# sempre da bolsa principal e são convertidos. Nestes três modos, títulos repetidos (a mesma ação duas vezes) só
# aparecem uma vez.
#   todos = cada título em todas as bolsas encontradas: a da lista, a europeia em euros e a americana (uma linha cada)
MODOS = {"base": "lista base", "eur": "privilegiar EUR", "eua": "privilegiar EUA", "todos": "todas as bolsas"}
FICH_MODO = os.path.join(PASTA, BASE + "_modo.json")
DIAS_REVER_EUR = 30          # volta a confirmar cada correspondência ao fim deste tempo
FICH_EUR = os.path.join(PASTA, BASE + "_bolsas.json")
SUF_EUR = {".DE": "Xetra", ".F": "Frankfurt", ".AS": "Euronext Amesterdão", ".PA": "Euronext Paris",
           ".BR": "Euronext Bruxelas", ".LS": "Euronext Lisboa", ".MI": "Borsa Italiana", ".MC": "Bolsa de Madrid",
           ".VI": "Viena", ".HE": "Helsínquia", ".IR": "Dublin", ".MU": "Munique", ".SG": "Estugarda",
           ".BE": "Berlim", ".DU": "Düsseldorf", ".HM": "Hamburgo"}
PREF_SUF = [".DE", ".AS", ".PA", ".MI", ".MC", ".BR", ".LS", ".F", ".MU", ".SG", ".BE", ".DU", ".HM", ".VI", ".HE", ".IR"]
MOEDA_SUF = {".L": "GBp", ".SW": "CHF", ".CO": "DKK", ".OL": "NOK", ".ST": "SEK"}
RATIOS_ADR = [0.1, 0.2, 0.25, 0.5, 1, 2, 3, 4, 5, 6, 8, 10]
# palpites (confirmados antes de serem usados; se falharem, procura-se no Yahoo)
CANDIDATOS = {"eur": {
    "AAPL": "APC.DE", "MSFT": "MSF.DE", "NVDA": "NVD.DE", "AMZN": "AMZ.DE", "GOOGL": "ABEA.DE", "META": "FB2A.DE",
    "TSLA": "TL0.DE", "AVGO": "1YD.DE", "BRK-B": "BRYN.DE", "JPM": "CMC.DE", "V": "3V64.DE", "MA": "M4I.DE",
    "LLY": "LLY.DE", "UNH": "UNH.DE", "XOM": "XONA.DE", "CVX": "CHV.DE", "JNJ": "JNJ.DE", "PG": "PRG.DE",
    "WMT": "WMT.DE", "COST": "CTO.DE", "HD": "HDI.DE", "ORCL": "ORC.DE", "NFLX": "NFC.DE", "AMD": "AMD.DE",
    "BAC": "NCB.DE", "KO": "CCC3.DE", "PEP": "PEP.DE", "MRK": "6MK.DE", "ABBV": "4AB.DE", "CRM": "FOO.DE",
    "ADBE": "ADB.DE", "CSCO": "CIS.DE", "INTC": "INL.DE", "QCOM": "QCI.DE", "PLTR": "PTX.DE", "UBER": "UT8.DE",
    "BA": "BCO.DE", "DIS": "WDP.DE", "NKE": "NKE.DE", "MCD": "MDO.DE", "IBM": "IBM.DE", "PFE": "PFE.DE",
    "MU": "MTE.DE", "GS": "GOS.DE", "CAT": "CAT1.DE", "WFC": "NWT.DE", "TSM": "TSFA.DE",
    "SHEL.L": "SHELL.AS", "ULVR.L": "UNA.AS", "AZN.L": "ZEG.DE", "HSBA.L": "HBC1.DE", "BP.L": "BPE5.DE",
    "RIO.L": "RIO1.DE", "GSK.L": "GS7.DE", "NESN.SW": "NESR.DE", "NOVN.SW": "NOT.DE", "ROP.SW": "RHO6.DE",
    "UBSG.SW": "0UB.DE", "ABBN.SW": "ABJ.DE", "NOVO-B.CO": "NOV.DE", "EQNR.OL": "DNQ.DE", "VOLV-B.ST": "VOL3.DE",
    "ERIC-B.ST": "ERCB.DE", "ATCO-A.ST": "ATC.DE"},
  "eua": {
    "SAP.DE": "SAP", "ASML.AS": "ASML", "SHEL.L": "SHEL", "AZN.L": "AZN", "HSBA.L": "HSBC", "ULVR.L": "UL",
    "BP.L": "BP", "RIO.L": "RIO", "GSK.L": "GSK", "DGE.L": "DEO", "BATS.L": "BTI", "REL.L": "RELX", "BARC.L": "BCS",
    "LLOY.L": "LYG", "NOVO-B.CO": "NVO", "NOVN.SW": "NVS", "UBSG.SW": "UBS", "TTE.PA": "TTE", "SAN.PA": "SNY",
    "INGA.AS": "ING", "PHIA.AS": "PHG", "ABI.BR": "BUD", "BBVA.MC": "BBVA", "SAN.MC": "SAN", "TEF.MC": "TEF",
    "ENI.MI": "E", "RACE.MI": "RACE", "STLAM.MI": "STLA", "EQNR.OL": "EQNR", "ERIC-B.ST": "ERIC", "NOKIA.HE": "NOK",
    "DBK.DE": "DB", "NESN.SW": "NSRGY", "ROP.SW": "RHHBY", "ABBN.SW": "ABBNY", "MC.PA": "LVMUY", "OR.PA": "LRLCY",
    "SIE.DE": "SIEGY", "ALV.DE": "ALIZY", "DTE.DE": "DTEGY", "AIR.PA": "EADSY", "SU.PA": "SBGSY", "RMS.PA": "HESAY",
    "MBG.DE": "MBGYY", "BMW.DE": "BMWYY", "VOW3.DE": "VWAGY", "BAS.DE": "BASFY", "BAYN.DE": "BAYRY",
    "IFX.DE": "IFNNY", "ADS.DE": "ADDYY", "IBE.MC": "IBDRY", "ITX.MC": "IDEXY", "ENEL.MI": "ENLAY",
    "UCG.MI": "UNCRY", "ISP.MI": "ISNPY", "PRX.AS": "PROSY", "ADYEN.AS": "ADYEY", "HEIA.AS": "HEINY",
    "AD.AS": "ADRNY", "CS.PA": "AXAHY", "BNP.PA": "BNPQY", "AI.PA": "AIQUY", "EL.PA": "ESLOY", "SAF.PA": "SAFRY",
    "EDP.LS": "EDPFY", "GALP.LS": "GLPEY", "JMT.LS": "JRONY", "ZURN.SW": "ZURVY", "CFR.SW": "CFRUY",
    "ATCO-A.ST": "ATLKY", "VOLV-B.ST": "VLVLY", "RR.L": "RYCEY", "BA.L": "BAESY", "GLEN.L": "GLNCY",
    "LSEG.L": "LNSTY", "RHM.DE": "RNMBY", "DB1.DE": "DBOEY", "WKL.AS": "WTKWY", "MUV2.DE": "MURGY"}}


def sufixo(t):
    t = str(t).upper()
    return t[t.rfind("."):] if "." in t else ""


def em_euros(t):
    return sufixo(t) in SUF_EUR


def nos_eua(t):
    return sufixo(t) == ""


def moeda_prevista(t):
    """Moeda do preço pelo sufixo do ticker (sem sufixo = EUA, dólares)."""
    s = sufixo(t)
    return "EUR" if s in SUF_EUR else MOEDA_SUF.get(s, "USD" if not s else None)


def nome_bolsa(t):
    return "EUA" if nos_eua(t) else SUF_EUR.get(sufixo(t), sufixo(t))


def modo_atual():
    try:
        with open(FICH_MODO, encoding="utf-8") as f:
            m = json.load(f).get("modo")
        return m if m in MODOS else "eur"
    except Exception:
        return "eur"


def guardar_modo(m):
    try:
        with open(FICH_MODO, "w", encoding="utf-8") as f:
            json.dump({"modo": m, "data": f"{datetime.now():%d/%m/%Y %H:%M}"}, f)
    except Exception:
        pass


def fonte_analistas(yf, nome, ticker, preco, moeda):
    """Listagem principal (com analistas) da mesma ação, procurada pelo nome e confirmada pelo preço (±5%, ou um
    múltiplo nos ADR). Guardada em FICH_EUR ("principal"), revista de DIAS_REVER_EUR em DIAS_REVER_EUR dias."""
    try:
        with open(FICH_EUR, encoding="utf-8") as f:
            cache = json.load(f)
    except Exception:
        cache = {}
    cp = cache.setdefault("principal", {})
    v = cp.get(ticker)
    if v and (datetime.now() - datetime.strptime(v.get("data", "01/01/2000"), "%d/%m/%Y")).days <= DIAS_REVER_EUR:
        return v.get("alt")
    escolha = None
    try:
        ref = preco * (taxas_eur({moeda}).get(moeda, (None, None))[1] or 0)
        quotes = yf.Search(nome, max_results=15, news_count=0, lists_count=0).quotes or []
        cand = [q["symbol"] for q in quotes if q.get("quoteType") == "EQUITY" and q.get("symbol")
                and q["symbol"].upper() != ticker.upper()][:6]
        for c in cand:
            inf = yf.Ticker(c).info or {}
            p = _num(inf.get("regularMarketPrice") or inf.get("currentPrice"))
            if not (p and inf.get("targetMeanPrice") and ref):
                continue
            m = inf.get("currency")
            r = p * (taxas_eur({m}).get(m, (None, None))[1] or 0) / ref
            if r and min(abs(r / k - 1) for k in RATIOS_ADR) < 0.05:
                escolha = c
                break
    except Exception:
        escolha = None
    cp[ticker] = {"alt": escolha, "data": datetime.now().strftime("%d/%m/%Y")}
    try:
        with open(FICH_EUR, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=1)
    except Exception:
        pass
    if escolha:
        print(f"   {ticker}: analistas e valor justo vêm de {escolha} (bolsa principal)")
    return escolha


def bolsas_alt(titulos, modo=None):
    """{ticker da lista: ticker a usar} conforme o modo (só os que mudam). Cache em FICH_EUR, por modo."""
    modo = modo or modo_atual()
    if modo in ("base", "todos"):          # todos: o recolher pede as duas (eur e eua) à parte; aqui, a da lista
        return {}
    import yfinance as yf
    alvo_ok = em_euros if modo == "eur" else nos_eua
    try:
        with open(FICH_EUR, encoding="utf-8") as f:
            cache = json.load(f)
    except Exception:
        cache = {}
    cm = cache.setdefault(modo, {})
    hoje = datetime.now()
    fora = [(n, t, etf) for n, c, t, b, etf in titulos if not alvo_ok(t) and moeda_prevista(t)
            and not (modo == "eua" and etf)]
    rever = [(n, t) for n, t, _ in fora if t not in cm or
             (hoje - datetime.strptime(cm[t].get("data", "01/01/2000"), "%d/%m/%Y")).days > DIAS_REVER_EUR]
    if rever:
        print(f"A procurar a mesma ação na bolsa preferida ({MODOS[modo]}) para {len(rever)} títulos "
              f"(só de {DIAS_REVER_EUR} em {DIAS_REVER_EUR} dias)...")
        cand = {}
        for nome, t in rever:
            c = [CANDIDATOS[modo][t.upper()]] if t.upper() in CANDIDATOS[modo] else []
            try:
                for q in yf.Search(nome, max_results=20, news_count=0, lists_count=0).quotes or []:
                    sym = q.get("symbol") or ""
                    if alvo_ok(sym) and q.get("quoteType") in ("EQUITY", "ETF") and sym not in c:
                        c.append(sym)
            except Exception:
                pass
            cand[t] = c[:8]
            time.sleep(0.2)
        todos = sorted({x for v in cand.values() for x in v} | {t for _, t in rever})
        try:
            h = _por_ticker(yf.download(todos, period="3mo", interval="1d", group_by="ticker", auto_adjust=False,
                                        threads=True, progress=False), todos)
        except Exception as e:
            print("Aviso: não consegui confirmar as outras bolsas:", e)
            h = {}
        tx = taxas_eur({moeda_prevista(t) for _, t in rever} | {"USD"} | {moeda_prevista(x) for v in cand.values()
                                                                             for x in v})
        for nome, t in rever:
            escolha, melhor = None, 0.0
            try:
                base = h[t].dropna(subset=["Close"])
                for x in cand.get(t, []):
                    try:
                        hx = h[x].dropna(subset=["Close"])
                        if len(hx) < 10:
                            continue
                        f_x = tx.get(moeda_prevista(x), (None, None))[1]
                        # em Londres há ETFs cotados em dólares: aceita-se o câmbio que bater (pence ou dólares)
                        r = None
                        for m in {moeda_prevista(t), "USD"}:
                            f_t = tx.get(m, (None, None))[1]
                            if f_x and f_t:
                                rr = float(hx["Close"].iloc[-1]) * f_x / (float(base["Close"].iloc[-1]) * f_t)
                                k = min(RATIOS_ADR if modo == "eua" else [1], key=lambda q: abs(rr / q - 1))
                                if r is None or abs(rr / k - 1) < abs(r - 1):
                                    r = rr / k
                        sem = lambda s: s["Close"].resample("W").last().pct_change().dropna()
                        a, b = sem(base).align(sem(hx), join="inner")
                        corr = float(a.corr(b)) if len(a) >= 6 else 0.0
                        neg = float((hx["Close"] * hx["Volume"].fillna(0)).mean()) * (f_x or 1)
                        bonus = (len(PREF_SUF) - PREF_SUF.index(sufixo(x))) if sufixo(x) in PREF_SUF else 0
                        nota = neg * (1 + bonus / 4)
                        if r is not None and 0.95 < r < 1.05 and corr > 0.6 and nota >= melhor:
                            escolha, melhor = x, nota
                    except Exception:
                        continue
            except Exception:
                escolha = None
            cm[t] = {"alt": escolha, "data": hoje.strftime("%d/%m/%Y")}
            print(f"   {t:12} -> {escolha or '(fica a da lista)'}")
        try:
            with open(FICH_EUR, "w", encoding="utf-8") as f:
                json.dump(cache, f, ensure_ascii=False, indent=1)
        except Exception:
            pass
    return {t: cm[t]["alt"] for _, t, _ in fora if cm.get(t, {}).get("alt")}


# Saúde financeira e qualidade do negócio (só ações; dados do Yahoo: últimos 12 meses e último relatório anual).
# Pontos graduais (0 a 1 por critério, em linha reta entre "mau" e "bom"), nota = 5 × média dos critérios com dados:
#   QUALIDADE (o negócio): ROIC 5% → 15% · lucro bruto / ativo 10% → 30% · FCF / lucro 0,5 → 0,9 (precisa de 2 dos 3).
#     O preço (FCF yield, EV/EBITDA) fica fora: já está na pontuação (valor justo e analistas); só se mostra.
#   SAÚDE: dívida líquida / EBITDA 3× → 1× (utilities e imobiliárias 5× → 2×; com mais caixa do que dívida = 1)
#     · EBIT / juros 3× → 8× (sem juros = 1) · FCF positivo = 1, negativo com caixa líquida = 0,5, negativo com dívida = 0.
#   ALERTA (basta um): dívida acima do limite "mau", juros abaixo de 3×, FCF negativo com dívida, EBITDA negativo com
#     dívida -> a nota geral não passa de 2.5.
# NOTA GERAL = 75% pontuação + 12,5% qualidade + 12,5% saúde (o que faltar, p. ex. nos ETFs e nos bancos, conta como a
#   pontuação). Bancos, seguradoras e financeiras: qualidade e saúde não se aplicam.
NOTA_ALERTA = 2.5
PICO_LUCRO = 1.5          # lucro dos últimos 12 meses > 1,5 × a média dos últimos 5 anos -> valor justo com a média
# Ponderações da nota geral (folha Ponderações; o total de cada bloco tem de dar 100%)
PESOS_PROPOSTOS = {"pont": 0.30, "qual": 0.30, "saude": 0.20, "tend": 0.20,
                   "etf_pont": 0.60, "etf_acoes": 0.20, "etf_tend": 0.20, "t10": 0.65, "t5": 0.35,
                   "caro_j": 0.20, "caro_pe": 40.0, "caro_fcf": 0.02, "caro_etf": 0.50, "usar_tend": 1, "usar_caro": 1,
                   "nivel_caro": 2.0, "lp_brando": 0, "pico_lucro": 1,
                   "pe_hab_max": 35.0, "cap_pe_sinal": 1,
                   "ft_usar": 1, "ft_s": 0.10, "ft_10": 0.15, "ft_5": 0.20, "ft_1": 0.35, "ft_6": 0.20,
                   "ft_mult": 2.0, "ft_r2": 1, "ft_lim": 0.60, "ft_vol": 1, "ft_vol0": 0.25, "ft_volpen": 0.03,
                   "ft_sub": 0.05, "ft_sub_etf": 0.02, "ft_vol0_etf": 0.35,
                   "ft_bonus": 1, "ft_bmax": 0.09, "ft_bvol": 0.20, "g_nota": 2.0, "g_preco": 1.0,
                   "rec_jan": 45, "rec_jan_etf": 60, "usar_analistas": 1}
CALIB = {}                 # calibração da nota global com os analistas (último cálculo)
FICH_CALIB = os.path.join(PASTA, BASE + "_calibracao_nota_geral.json")
# Fator de tendência (folha Ponderações): média ponderada das inclinações das retas (desde sempre, 10 e 5 anos com
# fechos mensais; 1 ano e 6 meses com fechos diários), com as quedas multiplicadas (ft_mult) e as subidas recentes
# pesadas pela regularidade (R²). Fator = interpolação da média nesta escala; multiplica a nota geral. Substitui o
# travão do longo prazo (G4 -> 2,5) e o «estabilizou» da nota da tendência.
FT_ESCALA = [(-0.20, 0.70), (0.0, 0.85), (0.05, 0.93), (0.10, 1.00)]   # V108: mais suave
FT_MARCA = "fator V108"     # folha Ponderações (H2): sem isto, os valores do fator lidos são os da versão anterior
FT_PER = [("ft_s", "Desde sempre", "val_sempre", None), ("ft_10", "Últimos 10 anos", "val10", None),
          ("ft_5", "Últimos 5 anos", "val5", None), ("ft_1", "Último ano", "inc1a", "r2_1a"),
          ("ft_6", "Últimos 6 meses", "inc6m", "r2_6m")]
FT_TXT = "Usar o fator (x):"
FT_R2_TXT = "Subidas recentes pela regularidade (x):"
FT_VOL_TXT = "Penalizar a volatilidade (x):"
FT_VOL_MIN = 0.85          # a volatilidade nunca tira mais de 15%
FT_BONUS_TXT = "Bónus de estabilidade (x):"
TRAVOES_CARO = [  # (chave, nome, o que representa, formato)
    ("nivel_caro", "Nível de preço até ao qual é cara",
     "Nível de preço (5 muito barata … 1 muito cara) = 40% preço face ao valor justo, com o PER habitual limitado (30% "
     "abaixo → 5, igual → 3, 50% acima → 1) + 35% PER, o menor entre o atual e o futuro (10 → 5, 22 → 3 (média do "
     "mercado), 30 → 2, 40 → 1) + 25% FCF yield (7% → 5, 3,5% → 3, 2,5% → 2, "
     "1,5% → 1; não nos bancos e seguros). ETFs: média das maiores ações, pesadas. Com este nível ou abaixo conta como "
     "cara: nota geral e nível de preço a azul (a nota não muda).", "0.0"),
]
BLOCOS_PESOS = (("pont", "qual", "saude", "tend"), ("etf_pont", "etf_acoes", "etf_tend"), ("t10", "t5"),
                ("ft_s", "ft_10", "ft_5", "ft_1", "ft_6"))
USAR_TEND_TXT = "Usar este quadro (x):"
USAR_CARO_TXT = "Usar estes avisos (x):"
LP_BRANDO_TXT = "Travão brando (x):"
PICO_TXT = "Lucro no pico (x):"
CAP_TXT = "Também no sinal (x):"
MARCA_PESOS = "pesos V048"           # folha Ponderações já com as percentagens antigas postas uma vez
PESOS_ANTIGOS = {"pont": 0.75, "qual": 0.125, "saude": 0.125, "tend": 0.0, "etf_pont": 0.75, "etf_acoes": 0.25,
                 "etf_tend": 0.0, "usar_tend": 0, "lp_brando": 0, "pico_lucro": 0}   # = classificação de antes da V039
PESOS = dict(PESOS_PROPOSTOS)
PESOS_TXT = [  # (chave, componente, o que representa)
    ("pont", "Sinal (pontuação 1-5)",
     "O momento e o preço: médias 100/200 (tendência), valor justo face ao preço e potencial dos analistas a 12 meses. "
     "É a coluna \"Nota antiga\" (pontuação) e a decisão."),
    ("qual", "Qualidade do negócio (0-5)",
     "Rentabilidade e solidez do negócio: ROIC, lucro bruto sobre ativos e conversão do lucro em free cash flow."),
    ("saude", "Saúde financeira (0-5)",
     "Capacidade de aguentar uma crise: dívida líquida / EBITDA, cobertura dos juros e free cash flow."),
    ("tend", "Tendência de longo prazo (1-5)",
     "Crescimento médio anual do preço (com dividendos) nos últimos 10 e 5 anos, com os pesos do quadro "
     "\"Tendência\" em baixo: 0%/ano ou menos = 1, 15%/ano ou mais = 5. Se os últimos 5 anos ficarem abaixo do "
     "mínimo (G4 da folha Ações), a nota da tendência não passa de 2: um título que estabilizou não pontua muito "
     "(só sem o fator de tendência; com ele, quem estabilizou é penalizado pelo fator)."),
    ("etf_pont", "ETFs: sinal (pontuação 1-5)",
     "Momento do ETF: médias 100/200 e subida dos últimos 6 meses face aos outros ETFs."),
    ("etf_acoes", "ETFs: nota das ações do ETF",
     "Média das notas gerais das maiores posições do ETF (folha Ações), pesadas pelo peso de cada uma no ETF."),
    ("etf_tend", "ETFs: tendência de longo prazo (1-5)",
     "Como nas ações: crescimento médio anual nos últimos 10 e 5 anos (pesos do quadro \"Tendência\")."),
    ("t10", "Últimos 10 anos",
     "Peso do crescimento médio anual dos últimos 10 anos (com menos de 10 anos de história: desde o início)."),
    ("t5", "Últimos 5 anos",
     "Peso do crescimento médio anual dos últimos 5 anos: apanha os títulos que abrandaram ou estabilizaram."),
]
TEND_MAU, TEND_BOM = 0.0, 0.15
FICH_PESOS = os.path.join(PASTA, BASE + "_ponderacoes.json")


def _pesos_ok(p):
    return all(abs(sum(p.get(k, PESOS_PROPOSTOS[k]) for k in b) - 1) < 0.0005 for b in BLOCOS_PESOS
               if (b != ("t10", "t5") or p.get("usar_tend", 1)) and (b[0] != "ft_s" or p.get("ft_usar", 1)))


def pc1(x):
    """Percentagem com 1 casa decimal só quando é precisa: 12,5% / 30%."""
    t = f"{x * 100:.1f}".replace(".", ",")
    return (t[:-2] if t.endswith(",0") else t) + "%"


def usar_pesos(p, origem=""):
    """Põe em uso as ponderações (dicionário com as chaves de PESOS_PROPOSTOS)."""
    q = dict(PESOS_PROPOSTOS)
    q.update({k: float(v) for k, v in (p or {}).items() if k in q and isinstance(v, (int, float))})
    if _pesos_ok(q):
        PESOS.update(q)
        if origem:
            print(f"Ponderações ({origem}): ações {pc1(PESOS['pont'])} sinal + {pc1(PESOS['qual'])} qualidade + "
                  f"{pc1(PESOS['saude'])} saúde + {pc1(PESOS['tend'])} tendência LP; ETFs {pc1(PESOS['etf_pont'])} sinal + "
                  f"{pc1(PESOS['etf_acoes'])} ações do ETF + {pc1(PESOS['etf_tend'])} tendência; tendência "
                  + (f"{pc1(PESOS['t10'])} 10 anos + {pc1(PESOS['t5'])} 5 anos." if PESOS.get("usar_tend", 1)
                     else "pela reta mais fraca (desde sempre / 5 anos)."))
        return True
    return False


def carregar_pesos(ex=None):
    """Lê as ponderações da folha Ponderações (o que lá estiver é o que conta). Total diferente de 100% -> erro."""
    lidos = None
    if ex is not None:
        try:
            r = ex.ops([{"op": "ler", "folha": FOLHA_PESOS, "rng": "E5:E8"}, {"op": "ler", "folha": FOLHA_PESOS, "rng": "E12:E14"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "E18:E19"}, {"op": "ler", "folha": FOLHA_PESOS, "rng": "B14"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "B23:E26"}, {"op": "ler", "folha": FOLHA_PESOS, "rng": "D16:E16"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "D21:E21"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "D28:E28"}, {"op": "ler", "folha": FOLHA_PESOS, "rng": "H1"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "D31:E31"}, {"op": "ler", "folha": FOLHA_PESOS, "rng": "B34:E35"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "B38:E58"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "H2"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "B62:E63"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "B67:E68"},
                        {"op": "ler", "folha": FOLHA_PESOS, "rng": "D71:E71"}])
            antiga = not txt(r[3][0][0]).startswith("ETFs: tendência")   # folha da V039: sem os quadros novos
            if antiga:
                r[1][2][0], r[2][0][0], r[2][1][0] = 0, 0, 0
            v = [x[0] for x in r[0]] + [x[0] for x in r[1]] + [x[0] for x in r[2]]

            def pc(x):
                x = float(str(x).replace("%", "").replace(",", ".").strip()) if x not in (None, "") else 0.0
                return x / 100 if x > 1 else x
            lidos = dict(zip(("pont", "qual", "saude", "tend", "etf_pont", "etf_acoes", "etf_tend", "t10", "t5"), map(pc, v)))
            if antiga:                                            # os quadros novos ainda não existem: os propostos
                lidos.update({k: PESOS_PROPOSTOS[k] for k in ("etf_pont", "etf_acoes", "etf_tend", "t10", "t5")})
            if len(r) > 5 and txt(r[5][0][0]) == USAR_TEND_TXT:      # x: usa o quadro dos 10 / 5 anos
                lidos["usar_tend"] = 1 if txt(r[5][0][1]).lower() == "x" else 0
            else:
                lidos["usar_tend"] = 1
            if len(r) > 6 and txt(r[6][0][0]) == USAR_CARO_TXT:      # x: avisos de cara ligados
                lidos["usar_caro"] = 1 if txt(r[6][0][1]).lower() == "x" else 0
            else:
                lidos["usar_caro"] = 1
            if len(r) > 7 and txt(r[7][0][0]) == LP_BRANDO_TXT:
                lidos["lp_brando"] = 1 if txt(r[7][0][1]).lower() == "x" else 0
            else:
                lidos["lp_brando"] = 0
            if len(r) > 10 and txt(r[10][0][0]) == "PER habitual máximo":
                try:
                    lidos["pe_hab_max"] = float(str(r[10][0][3]).replace(",", "."))
                except (TypeError, ValueError):
                    pass
                lidos["cap_pe_sinal"] = 1 if (txt(r[10][1][2]) == CAP_TXT and txt(r[10][1][3]).lower() == "x") else 0
            lidos["pico_lucro"] = 1 if (len(r) > 9 and txt(r[9][0][0]) == PICO_TXT and txt(r[9][0][1]).lower() == "x") else 0
            ft = r[11] if len(r) > 11 else []
            ft_novo = len(r) > 12 and txt(r[12][0][0]) == FT_MARCA    # folha já com os valores suaves da V108
            for lin in (r[13] if len(r) > 13 else []):                 # pesos da nota global (V110)
                k_ = {"Peso da nota geral": "g_nota", "Peso do nível de preço": "g_preco"}.get(txt(lin[0]))
                if k_:
                    try:
                        lidos[k_] = max(0.0, float(str(lin[3]).replace(",", ".")))
                    except (TypeError, ValueError):
                        pass
            if len(r) > 15 and txt(r[15][0][0]) == USAR_ANALISTAS_TXT:   # V162: x = ignorar os analistas
                lidos["usar_analistas"] = 0 if txt(r[15][0][1]).lower() == "x" else 1
            else:
                lidos["usar_analistas"] = PESOS_PROPOSTOS["usar_analistas"]
            for lin in (r[14] if len(r) > 14 else []):                 # janelas da compra forte (V148)
                k_ = {REC_JAN_TXT: "rec_jan", REC_JAN_ETF_TXT: "rec_jan_etf"}.get(txt(lin[0]))
                if k_:
                    try:
                        lidos[k_] = int(max(0, min(REC_MEMO, round(float(str(lin[3]).replace(",", "."))))))
                    except (TypeError, ValueError):
                        pass
            if ft and txt(ft[0][2]) == FT_TXT:                    # quadro do fator de tendência (V104)
                lidos["ft_usar"] = 1 if txt(ft[0][3]).lower() == "x" else 0
                for (k, nome, _, _), lin in zip(FT_PER, ft[2:7]):
                    if txt(lin[0]) == nome:
                        lidos[k] = pc(lin[3])
                for lin in ft[7:]:
                    nm = txt(lin[0])
                    if not ft_novo and nm in ("Multiplicador das quedas", "Penalização por cada +10% de volatilidade",
                                              "Subida mínima por ano (ações)", "Subida mínima por ano (ETFs)"):
                        continue                         # V108: passam aos valores novos (mais suaves) uma vez
                    try:
                        if nm == "Multiplicador das quedas":
                            lidos["ft_mult"] = float(str(lin[3]).replace(",", "."))
                        elif nm == "Limite de cada inclinação (±)":
                            lidos["ft_lim"] = pc(lin[3])
                        elif nm in ("Volatilidade sem penalização", "Volatilidade sem penalização (ações)"):
                            lidos["ft_vol0"] = pc(lin[3])
                        elif nm == "Volatilidade sem penalização (ETFs)":
                            lidos["ft_vol0_etf"] = pc(lin[3])
                        elif nm == "Subida mínima por ano (ações)":
                            lidos["ft_sub"] = pc(lin[3])
                        elif nm == "Subida mínima por ano (ETFs)":
                            lidos["ft_sub_etf"] = pc(lin[3])
                        elif nm == "Bónus máximo de estabilidade":    # (o rótulo antigo, da V106 a 5%, é ignorado)
                            lidos["ft_bmax"] = pc(lin[3])
                        elif nm == "Volatilidade abaixo da qual há bónus":
                            lidos["ft_bvol"] = pc(lin[3])
                        elif nm == "Penalização por cada +10% de volatilidade":
                            lidos["ft_volpen"] = pc(lin[3])
                    except (TypeError, ValueError):
                        pass
                    if txt(lin[2]) == FT_R2_TXT:
                        lidos["ft_r2"] = 1 if txt(lin[3]).lower() == "x" else 0
                    if txt(lin[2]) == FT_VOL_TXT:
                        lidos["ft_vol"] = 1 if txt(lin[3]).lower() == "x" else 0
                    if txt(lin[2]) == FT_BONUS_TXT:
                        lidos["ft_bonus"] = 1 if txt(lin[3]).lower() == "x" else 0
            if len(r) > 8 and txt(r[8][0][0]) != MARCA_PESOS:         # uma só vez: as percentagens de antigamente
                lidos.update(PESOS_ANTIGOS)
                print("Folha Ponderações: postas as percentagens e os x da classificação antiga (75% sinal, 12,5% "
                      "qualidade, 12,5% saúde; ETFs 75/25; sem o quadro da tendência).")
            trav = r[4] if len(r) > 4 else []
            for (k, nome, _, fm), lin in zip(TRAVOES_CARO, trav):   # travões de caro (se o quadro existir)
                if txt(lin[0]) == nome and lin[3] not in (None, ""):
                    x = float(str(lin[3]).replace("%", "").replace(",", ".").strip())
                    lidos[k] = x / 100 if fm.endswith("%") and x > 1 else x
                    if k == "nivel_caro" and lidos[k] > 3:          # escala antiga (5 = caro): converte
                        lidos[k] = 6 - lidos[k]
                    if k == "nivel_caro" and not 1 <= lidos[k] <= 5:
                        raise SystemExit("Folha Ponderações: o nível de preço a partir do qual é cara tem de estar entre "
                                         "1 e 5.")
        except SystemExit:
            raise
        except Exception:
            lidos = None
    if lidos is not None:
        if not _pesos_ok(lidos):
            ta, te, tt, tf = (sum(lidos.get(k, PESOS_PROPOSTOS[k]) for k in b) for b in BLOCOS_PESOS)
            raise SystemExit(f"Folha Ponderações: os totais têm de dar 100% (ações: {ta:.1%}; ETFs: {te:.1%}; "
                             f"tendência: {tt:.1%}; fator de tendência: {tf:.1%}). Corrige as células a vermelho e "
                             "volta a carregar no botão.")
        usar_pesos(lidos, "folha Ponderações")
        _gravar_json(FICH_PESOS, lidos)
        return
    usar_pesos(_ler_json(FICH_PESOS, {}) or PESOS_PROPOSTOS, "guardadas")


FOLHA_PESOS = "Ponderações"


def _escala(v, pontos):
    """Interpola v numa escala [(valor, nível)] crescente ou decrescente nos valores; níveis 1 a 5."""
    if v is None:
        return None
    xs = [x for x, _ in pontos]
    cres = xs[0] < xs[-1]
    if (v <= xs[0]) if cres else (v >= xs[0]):
        return float(pontos[0][1])
    if (v >= xs[-1]) if cres else (v <= xs[-1]):
        return float(pontos[-1][1])
    for (x0, n0), (x1, n1) in zip(pontos, pontos[1:]):
        if (x0 <= v <= x1) if cres else (x1 <= v <= x0):
            return n0 + (n1 - n0) * (v - x0) / (x1 - x0)
    return None


NIVEL_J = [(0.30, 1), (0.10, 2), (-0.10, 3), (-0.30, 4), (-0.50, 5)]        # preço face ao valor justo
# (escala interna 1 barato … 5 caro; o nível mostrado é 6 − isto). "Justa" (3) = mais ou menos a média do mercado
NIVEL_PE = [(10, 1), (15, 2), (22, 3), (30, 4), (40, 5)]                     # PER (o menor: atual / futuro)
NIVEL_FCF = [(0.07, 1), (0.05, 2), (0.035, 3), (0.025, 4), (0.015, 5)]       # FCF yield


def rotulos_nivel(lim=None):
    """[(condição do Excel com {c} = célula, rótulo)] do nível de preço (5 barata … 1 cara); 'Cara' até ao limite da
    folha Ponderações."""
    lim = PESOS["nivel_caro"] if lim is None else lim
    mc = min(1.5, lim)
    return [(f"AND(ISNUMBER({{c}}),{{c}}<={mc})", "Preço muito caro"), (f"AND({{c}}>{mc},{{c}}<={lim})", "Preço caro"),
            (f"AND({{c}}>{lim},{{c}}<=2.5)", "Preço algo caro"), (f"AND({{c}}>{max(2.5, lim)},{{c}}<3.5)", "Preço justo"),
            (f"AND({{c}}>=3.5,{{c}}<=4.25)", "Preço barato"), (f"AND(ISNUMBER({{c}}),{{c}}>4.25)", "Preço muito barato")]


COR_NIVEL = {"Preço muito caro": "E0484B", "Preço caro": "F8A5A6", "Preço algo caro": "FCE4D6", "Preço justo": "FFFFFF",
             "Preço barato": "C6E5B3", "Preço muito barato": "4E9F5D"}


def rotulo_nivel(v):
    lim = PESOS["nivel_caro"]
    if not isinstance(v, (int, float)):
        return ""
    return ("Preço muito caro" if v <= min(1.5, lim) else "Preço caro" if v <= lim else "Preço algo caro" if v <= 2.5 else
            "Preço justo" if v < 3.5 else "Preço barato" if v <= 4.25 else "Preço muito barato")


def nivel_preco(d):
    """Nível de preço 5 (muito barata) a 1 (muito cara) de uma ação: 40% preço face ao valor justo (PER habitual
    limitado), 35% PER (o menor entre o atual e o futuro), 25% FCF yield (não nos bancos / seguros). (nível, explicação)."""
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    J, fy = n("justo_pct"), n("fcf_yield")
    # valor justo com o PER habitual limitado (evita a "bolha do próprio histórico", p. ex. NVIDIA); se o sinal já usa
    # o limite, o valor justo já vem limitado
    ph, vj, pr = n("pe_hab"), n("justo"), n("preco")
    if ph and vj and pr and ph > PESOS["pe_hab_max"] and not PESOS.get("cap_pe_sinal"):
        J = vj * PESOS["pe_hab_max"] / ph / pr - 1
    pes = [x for x in (n("pe"), n("pe_fwd")) if x is not None]
    pe = min(pes) if pes else None
    if txt(d.get("saude")).startswith("não se aplica"):
        fy = None
    partes = [(0.4, _escala(J, NIVEL_J)), (0.35, _escala(pe, NIVEL_PE)), (0.25, _escala(fy, NIVEL_FCF))]
    partes = [(w, v) for w, v in partes if v is not None]
    if not partes:
        return None, None
    nv = 6 - sum(w * v for w, v in partes) / sum(w for w, _ in partes)       # 5 = muito barata … 1 = muito cara
    det = "; ".join(x for x in (
        (f"preço {-J:+.0%} face ao valor justo".replace("+-", "−").replace("-", "−") if J is not None else None),
        (f"PER {pe:.0f}" if pe is not None else None),
        (f"FCF yield {fy:.1%}" if fy is not None else None)) if x)
    return round(nv, 1), det                      # 1 casa: o número mostrado e o rótulo batem sempre certo


def motivo_caro(d):
    """Porque é que a ação está cara (texto) ou None. Só ações (os ETFs não têm valor justo nem PER)."""
    if not PESOS.get("usar_caro", 1):               # sem x na folha Ponderações: sem avisos
        return None
    nv = d.get("nivel_preco") if isinstance(d.get("nivel_preco"), (int, float)) else None
    if nv is None or nv > PESOS["nivel_caro"]:
        return None
    if txt(d.get("tipo")) == "ETF":
        return f"cara: nível de preço {nv:.1f}/5 (média das maiores ações do ETF; limite {PESOS['nivel_caro']:.1f})"
    return f"cara: nível de preço {nv:.1f}/5 ({nivel_preco(d)[1]}; limite {PESOS['nivel_caro']:.1f})"


def _motivo_caro_antigo(d):
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    J, fy = n("justo_pct"), n("fcf_yield")
    pe = min(x for x in (n("pe"), n("pe_fwd")) if x is not None) if (n("pe") or n("pe_fwd")) else None
    if J is not None and J < -PESOS["caro_j"]:
        return f"cara: preço {-J:.0%} acima do valor justo (limite {PESOS['caro_j']:.0%})"
    # âncora absoluta: PER (o menor entre o atual e o futuro, que favorece quem cresce) alto E FCF yield baixo, juntos
    financeira = txt(d.get("saude")).startswith("não se aplica")
    if pe is not None and pe > PESOS["caro_pe"] and (financeira or fy is None or fy < PESOS["caro_fcf"]):
        return (f"cara: PER {pe:.0f} acima de {PESOS['caro_pe']:.0f}"
                + (f" e FCF yield {fy:.1%} abaixo de {PESOS['caro_fcf']:.1%}" if fy is not None and not financeira else ""))
    return None


def reta_r2(valores, por_ano):
    """(subida por ano, R²) da reta dos mínimos quadrados sobre o logaritmo das cotações (por_ano pontos por ano)."""
    import numpy as np
    y = np.log(np.asarray(valores, dtype=float))
    if len(y) < 20 or not np.all(np.isfinite(y)):
        return None, None
    x = np.arange(len(y)) / por_ano
    b, a = np.polyfit(x, y, 1)
    sy = ((y - y.mean()) ** 2).sum()
    r2 = 1 - ((y - (a + b * x)) ** 2).sum() / sy if sy > 0 else 0.0
    return float(np.exp(b) - 1), float(max(0.0, min(1.0, r2)))


ARR_BASE, ARR_SUB = 120, 30          # arranque: 120 dias de bolsa parada (ou a descer) e depois 30 a subir
ARR_BASE_MAX, ARR_SUB_MIN = 0.05, 0.30   # reta da base ≤ +5%/ano; reta dos 30 dias ≥ +30%/ano; preço acima do máximo
# Testado (42 ações 2016-26; S&P 500 desde 1950): antes de as médias cruzarem, os 6 meses seguintes renderam mais do
# que a média (ações +24% vs +17%; S&P 86% das vezes a ganhar vs 70%); nos ETFs diversificados quase não fez diferença
# (por isso só nas ações). Ao contrário (quebra depois de parada) NÃO prevê mais queda: não há regra de venda.


SUB60_D, SUB60_MIN, SUB60_R2 = 60, 0.40, 0.6   # tendência forte também com a reta dos últimos 60 dias ≥ +40%/ano
# (R² ≥ 0,6) e o preço acima da média 200, sem esperar que as médias se afastem 5% (que chega tarde depois de um
# patamar). Testado (42 ações; S&P 500 desde 1950): as datas promovidas renderam como as de sinal 4 ou mais (S&P
# +11,6% a 12 meses vs +7,0% das que ficam em 3); os 5% entre médias continuam (apanham outros casos).


REC_QUEDA, REC_QUEDA_ETF, REC_I60 = -0.15, -0.20, 0.30   # recuperação (V148)
REC_JANELA, REC_JANELA_ETF, REC_PROTEGE, REC_MEMO = 45, 60, 63, 130
REC_JAN_TXT, REC_JAN_ETF_TXT = "Janela da compra forte (ações)", "Janela da compra forte (ETFs)"
USAR_ANALISTAS_TXT = "Ignorar os analistas no sinal (x):"   # V162 (V160-V161: «Usar …», ignorado)
REC_VOL = 0.4      # a queda exigida é pelo menos 0,4 × a volatilidade anual (Tesla 60% -> queda ≥ 24%): nos títulos
#                    muito voláteis uma queda de 15% é banal (testado: menos sinais e mais certeiros)
# RECUPERAÇÃO: o título caiu do máximo de 1 ano (ações ≥ 15%, ETFs ≥ 20%; e ≥ 0,4 × volatilidade), voltou a subir com força (reta dos últimos
# 60 dias ≥ +30%/ano) e está acima da média 50 (ações: também acima da média 200, recuperação confirmada).
#  -> sinal 5, a decisão sobe um nível e o fator de tendência usa só o longo prazo (as quedas do último ano já não
#     contam), durante o sinal e mais REC_JANELA dias de bolsa (ETFs REC_JANELA_ETF) com o preço acima da média 50;
#  -> depois, até REC_PROTEGE dias com o preço acima da média 50: sem «Venda» por as médias ainda não terem cruzado
#     (ETFs: conta como tendência a subir; nas ações comprar nesta fase não rendeu mais nos testes).
# Testado (V148): 20 ações dos EUA 2008-26, 87% a ganhar nos 6 meses seguintes vs 70% (+13,9% face ao normal do
# título); fora da amostra: 11 outras ações 70% vs 64%, S&P 500 desde 1950 100% vs 70% (26 casos), 27 ETFs 2000-23
# 88% vs 60%. A vantagem cai ao fim de 1-3 meses (por isso a janela). Falha em quedas longas (2000-03).


def recuperacao(d, preco):
    """A regra da recuperação no preço de hoje (com os valores guardados no ▶ ATUALIZAR)."""
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    mx, i60, m50, m200, vol = n("max1a"), n("inc60"), n("mm50"), n("mm200c"), n("vol1a")
    if not (mx and i60 is not None and m50 and preco):
        return False
    etf = txt(d.get("tipo")) == "ETF"
    queda = -max(-(REC_QUEDA_ETF if etf else REC_QUEDA), REC_VOL * (vol or 0))
    if etf:
        return preco / mx - 1 <= queda and i60 >= REC_I60 and preco > m50
    return preco / mx - 1 <= queda and i60 >= REC_I60 and preco > m50 and bool(m200) and preco > m200


def _rec_sinais(P):
    """Sinais diários da recuperação (DataFrame ou Series de cotações): (regra das ações, regra dos ETFs, reta 60 d,
    máximo de 1 ano)."""
    import numpy as np
    import pandas as pd
    df = P.to_frame() if isinstance(P, pd.Series) else P
    df = df.astype(float)
    LD = np.log(df)
    kd = pd.DataFrame({c: np.arange(len(LD), dtype=float) for c in LD.columns}, index=LD.index).where(LD.notna())
    rol = lambda x: x.rolling(61, min_periods=55)
    mk, my, mky, mkk = rol(kd).mean(), rol(LD).mean(), rol(kd * LD).mean(), rol(kd * kd).mean()
    i60 = np.exp((mky - mk * my) / (mkk - mk * mk) * 252) - 1
    m50, m200 = df.rolling(50, min_periods=50).mean(), df.rolling(200, min_periods=200).mean()
    mx = df.rolling(252, min_periods=200).max()
    vol = LD.diff().rolling(252, min_periods=200).std() * np.sqrt(252)
    base = (df > m50) & (i60 >= REC_I60)
    ra = base & (df / mx - 1 <= -np.maximum(-REC_QUEDA, REC_VOL * vol.fillna(0))) & (df > m200)
    re_ = base & (df / mx - 1 <= -np.maximum(-REC_QUEDA_ETF, REC_VOL * vol.fillna(0)))
    return ra, re_, i60, mx


def _dias_desde(b):
    """Dias de bolsa desde o último True (DataFrame booleano) -> DataFrame (NaN se não houve)."""
    import numpy as np
    import pandas as pd
    pos = pd.DataFrame({c: np.arange(len(b), dtype=float) for c in b.columns}, index=b.index)
    ult = pos.where(b.fillna(False).astype(bool)).ffill()
    return pos - ult


def arranque(v):
    """True se as cotações diárias v (as mais recentes no fim) mostram um arranque: base de ARR_BASE dias parada ou a
    descer, depois ARR_SUB dias a subir depressa e o preço acima do máximo da base."""
    import numpy as np
    if v is None or len(v) < ARR_BASE + ARR_SUB + 1:
        return False
    v = np.asarray(v, dtype=float)
    base, rec = v[-(ARR_BASE + ARR_SUB + 1):-(ARR_SUB + 1)], v[-(ARR_SUB + 1):]
    if (base <= 0).any() or (rec <= 0).any():
        return False
    sl = lambda y: float(np.exp(np.polyfit(np.arange(len(y)), np.log(y), 1)[0] * 252) - 1)
    return sl(base) <= ARR_BASE_MAX and sl(rec) >= ARR_SUB_MIN and v[-1] > base.max()


def inclin_recentes(cl):
    """Retas do último ano (252 dias) e dos últimos 6 meses (126 dias) das cotações diárias -> dicionário."""
    out = {}
    cl = cl.dropna()
    cl = cl[cl > 0]
    import numpy as np
    if len(cl) >= 50:
        out["mm50"] = float(cl.iloc[-50:].mean())
    if len(cl) >= ARR_BASE + ARR_SUB + 1:
        out["arranque"] = 1 if arranque(cl.to_numpy(dtype=float)) else 0
    if len(cl) >= 260:                                   # recuperação (V148)
        try:
            ra, re_, i6_, mx_ = _rec_sinais(cl)
            c0 = ra.columns[0]
            out["max1a"] = float(mx_[c0].iloc[-1]) if mx_[c0].notna().iloc[-1] else None
            out["inc60"] = float(i6_[c0].iloc[-1]) if i6_[c0].notna().iloc[-1] else None
            for k_, b_ in (("rec_a", ra), ("rec_e", re_)):
                dd_ = _dias_desde(b_)[c0].iloc[-1]
                out[k_] = int(dd_) if dd_ == dd_ and dd_ <= REC_MEMO else None
        except Exception:
            pass
    if len(cl) >= SUB60_D + 1:
        i60, q60 = reta_r2(cl.iloc[-(SUB60_D + 1):].to_numpy(dtype=float), 252)
        out["sub60"] = 1 if (i60 is not None and q60 is not None and i60 >= SUB60_MIN and q60 >= SUB60_R2) else 0
    for k, r2k, n in (("inc1a", "r2_1a", 252), ("inc6m", "r2_6m", 126)):
        if len(cl) >= int(n * 0.9):
            out[k], out[r2k] = reta_r2(cl.iloc[-n:].to_numpy(dtype=float), 252)
    if len(cl) >= 200:
        r = np.diff(np.log(cl.iloc[-253:].to_numpy(dtype=float)))
        out["vol1a"] = float(r.std() * np.sqrt(252))
    return out


def fator_vol(vol, etf=False):
    """Fator de volatilidade: 1 até ft_vol0 (ETFs: ft_vol0_etf); menos ft_volpen por cada +10%; nunca abaixo de 0,80."""
    if vol is None or not PESOS.get("ft_vol", 1):
        return 1.0
    v0 = PESOS.get("ft_vol0_etf", 0.35) if etf else PESOS.get("ft_vol0", 0.25)
    return max(FT_VOL_MIN, min(1.0, 1 - PESOS.get("ft_volpen", 0.03) * (vol - v0) / 0.10))


def fator_estab(vol, r2, inc1a):
    """Bónus de estabilidade: quem oscila pouco (volatilidade abaixo de ft_bvol) e sobe com regularidade no último
    ano (R² alto e reta a subir) ganha até ft_bmax. Volatilidade ≤ 60% do limite e R² ≥ 0,8 -> bónus inteiro."""
    if not PESOS.get("ft_bonus", 1) or vol is None or r2 is None or inc1a is None or inc1a <= 0:
        return 1.0
    lim = PESOS.get("ft_bvol", 0.20)
    g_vol = max(0.0, min(1.0, (lim - vol) / (lim * 0.4)))
    g_r2 = max(0.0, min(1.0, (r2 - 0.3) / 0.5))
    return 1 + PESOS.get("ft_bmax", 0.09) * g_vol * g_r2


def fator_tend(d):
    """Fator de tendência (folha Ponderações) -> (fator, média ponderada, [(período, inclinação, conta)]) ou None.
    Média ponderada das inclinações das retas (desde sempre, 10, 5 anos, último ano, 6 meses), cada uma limitada a
    ± ft_lim; as quedas contam ft_mult vezes; as subidas do último ano e dos 6 meses contam × R² (regularidade).
    Fator = escala FT_ESCALA (−20%/ano ou pior → 0,70; 0% → 0,85; +5% → 0,93; +10% ou mais → 1,00)."""
    import numpy as np
    if not PESOS.get("ft_usar", 1):
        return None
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    lim, mult = PESOS.get("ft_lim", 0.6), PESOS.get("ft_mult", 2.0)
    etf = txt(d.get("tipo")) == "ETF"
    sub = PESOS.get("ft_sub_etf", 0.02) if etf else PESOS.get("ft_sub", 0.05)   # subida mínima exigida por ano
    partes, soma, sw = [], 0.0, 0.0
    for k, nome, kv, kr in FT_PER:
        v = n(kv)
        if kv == "val_sempre" and v is None:
            v = n("lp")
        if v is None or not PESOS.get(k):
            continue
        c = max(-lim, min(lim, v))
        if c > 0 and kr and PESOS.get("ft_r2", 1) and n(kr) is not None:
            c *= n(kr)
        c -= sub                                 # o que fica abaixo da subida mínima conta como queda
        if c < 0:
            c *= mult
        partes.append((nome, v, c))
        soma += PESOS[k] * c
        sw += PESOS[k]
    if sw < 0.25:                               # quase sem história: sem fator
        return None
    m = soma / sw + sub                          # volta à escala da subida por ano (para mostrar e para a escala)
    f = float(np.interp(m, [x for x, _ in FT_ESCALA], [y for _, y in FT_ESCALA]))   # a escala é a de sempre
    fv = fator_vol(n("vol1a"), etf)
    # quem sobe muito (média ≥ +20%/ano) não é penalizado pela volatilidade (p. ex. Micron); entre +10% e +20%, parte
    fv = 1 - (1 - fv) * max(0.0, min(1.0, (0.20 - m) / 0.10))
    fb = fator_estab(n("vol1a"), n("r2_1a"), n("inc1a"))
    partes.append(("_vol", n("vol1a"), fv))
    partes.append(("_bonus", None, fb))
    return round(f * fv * fb, 2), m, partes


def momento_compra(d):
    """O momento certo para a compra forte (V148): a recuperação. (O arranque depois de um patamar também foi testado
    para isto e piorou: 6 meses depois −4,7% face ao normal, 58% a ganhar; fica só no sinal, como na V144.)"""
    return "recuperação" if d.get("_recup") else None


def fator_tend_atual(d):
    """Fator de tendência; no momento certo (recuperação / arranque) as quedas do último ano e dos 6 meses já não
    contam: o fator usa só as retas de longo prazo (desde sempre, 10 e 5 anos)."""
    if not momento_compra(d):
        return fator_tend(d)
    g = {k: PESOS.get(k) for k in ("ft_1", "ft_6")}
    try:
        PESOS["ft_1"] = PESOS["ft_6"] = 0
        return fator_tend(d)
    finally:
        PESOS.update(g)


def fator_txt(ft):
    """Descrição curta do fator para a coluna Longo prazo e para os comentários."""
    f, m, partes = ft
    pc = lambda v: "0,0%" if abs(v) < 0.0005 else f"{v:+.1%}".replace(".", ",")
    vol = next(((v, fv) for nome, v, fv in partes if nome == "_vol"), (None, 1.0))
    fb = next((x for nome, _, x in partes if nome == "_bonus"), 1.0)
    partes = [x for x in partes if not x[0].startswith("_")]
    quedas = [nome.lower() for nome, v, _ in partes if v < 0]
    fracos = [nome.lower() for nome, v, c in partes if v >= 0 and c < 0]
    ft_ = f / (vol[1] or 1) / (fb or 1)
    if ft_ >= 0.995:
        porque = "sobe de forma consistente"
    else:
        porque = "; ".join(x for x in (("cai: " + ", ".join(quedas)) if quedas else None,
                                       ("sobe pouco: " + ", ".join(fracos)) if fracos else None) if x) \
            or "sobe pouco ou de forma irregular"
    if vol[0] is not None and vol[1] < 0.995:
        porque += f"; volatilidade {vol[0]:.0%} → × {vol[1]:.2f}".replace(".", ",")
    if fb > 1.004:
        porque += f"; estável (oscila pouco e sobe com regularidade) → × {fb:.2f}".replace(".", ",")
    return (f"fator {f:.2f}".replace(".", ",") + f" (média ponderada {pc(m)}/ano; {porque})"
            + (f" → penaliza a nota geral × {f:.2f}".replace(".", ",") if f < 0.90 else
               f" → reduz ligeiramente a nota geral × {f:.2f}".replace(".", ",") if f < 0.995 else
               f" → premeia a nota geral × {f:.2f}".replace(".", ",") + " (no máx. 5)" if f > 1.004 else ""))


def lp_travao(d):
    """Subida anual usada no travão do longo prazo: a reta mais fraca entre desde sempre e 5 anos (como sempre foi) ou,
    com o x do travão brando na folha Ponderações, a MELHOR entre 10 anos e 5 anos (só trava se as duas forem fracas)."""
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    if PESOS.get("lp_brando") and (n("lp10") is not None or n("lp5") is not None):
        return max(x for x in (n("lp10"), n("lp5")) if x is not None)
    return n("lp")


def nota_tendencia(lp, lp10=None, lp5=None):
    """Tendência de longo prazo em nota 1-5: crescimento médio anual dos últimos 10 e 5 anos (pesos da folha
    Ponderações; 0%/ano -> 1, 15%/ano -> 5). Últimos 5 anos abaixo do mínimo (G4) -> no máx. 2 (estabilizou)."""
    if (lp10 is None and lp5 is None) or not PESOS.get("usar_tend", 1):   # sem o quadro: a reta mais fraca
        return None if lp is None else 1 + 4 * _grad(lp, TEND_MAU, TEND_BOM)
    a, b = (lp10 if lp10 is not None else lp5), (lp5 if lp5 is not None else lp10)
    nota = 1 + 4 * _grad(PESOS["t10"] * a + PESOS["t5"] * b, TEND_MAU, TEND_BOM)
    if b < LP_MIN and not PESOS.get("ft_usar", 1):      # com o fator de tendência este travão não se usa
        nota = min(nota, 2.0)
    return nota


def folha_pesos(wb):
    """Folha Ponderações: o quadro das percentagens da nota geral (editável; cada bloco tem de somar 100%)."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.worksheet.datavalidation import DataValidation
    ws = wb.create_sheet(FOLHA_PESOS)
    ws["A1"] = "Ponderações da nota geral"
    fino = Side(style="thin", color="BFBFBF")
    borda = Border(left=fino, right=fino, top=fino, bottom=fino)
    azul, amarelo = PatternFill("solid", fgColor="1F3B57"), PatternFill("solid", fgColor="FFF2CC")
    ws["B2"] = ("Muda as percentagens da coluna E (a amarelo) e carrega em ↻ RECLASSIFICAR E REORDENAR. Cada quadro tem "
                "de somar 100%: se não somar, o total e as células ficam a vermelho e os botões dão erro. Com a saúde em ALERTA ou o longo prazo "
                "abaixo do mínimo (G4 da folha Ações), a nota geral fica no máximo em 2,5, sejam quais forem as "
                "percentagens. As percentagens passam também para a app (atualizações automáticas).")
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B2"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("B2:E2")
    ws.row_dimensions[2].height = 54

    def bloco(r0, titulo, chaves):
        ws.cell(r0 - 1, 2, titulo).font = Font(bold=True, size=12)
        for j, h in enumerate(("Componente", "O que representa", "Proposto", "A usar"), 2):
            c = ws.cell(r0, j, h)
            c.font, c.fill, c.border = Font(bold=True, color="FFFFFF"), azul, borda
            c.alignment = Alignment(horizontal="center", vertical="center")
        txt_k = {k: (n, t) for k, n, t in PESOS_TXT}
        for i, k in enumerate(chaves, r0 + 1):
            n, t = txt_k[k]
            ws.cell(i, 2, n).font = Font(bold=True)
            ws.cell(i, 3, t).alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(i, 4, PESOS_PROPOSTOS[k]).font = Font(color="808080")
            e = ws.cell(i, 5, PESOS[k])
            e.fill, e.font = amarelo, Font(bold=True, size=12)
            for j in range(2, 6):
                ws.cell(i, j).border = borda
            ws.cell(i, 4).number_format = ws.cell(i, 5).number_format = "0.0%"
            ws.cell(i, 4).alignment = ws.cell(i, 5).alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[i].height = 44
        rt = r0 + 1 + len(chaves)
        ws.cell(rt, 2, "Total (tem de ser 100%)").font = Font(bold=True)
        ws.cell(rt, 4, f"=SUM(D{r0 + 1}:D{rt - 1})")
        ws.cell(rt, 5, f"=SUM(E{r0 + 1}:E{rt - 1})")
        for j in (4, 5):
            c = ws.cell(rt, j)
            c.number_format, c.font = "0.0%", Font(bold=True, size=12)
            c.alignment = Alignment(horizontal="center")
            c.border = borda
        vermelho = PatternFill("solid", fgColor="F4B6B6")
        ws.conditional_formatting.add(f"E{r0 + 1}:E{rt}", FormulaRule(formula=[f"ROUND($E${rt},3)<>1"], fill=vermelho,
                                                                       font=Font(bold=True, color="C00000"), stopIfTrue=True))
        ws.conditional_formatting.add(f"E{rt}", FormulaRule(formula=[f"ROUND($E${rt},3)=1"],
                                                            fill=PatternFill("solid", fgColor="C6EFCE")))
        dv = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=False)
        dv.error, dv.errorTitle = "Escreve uma percentagem entre 0% e 100%.", "Ponderação"
        ws.add_data_validation(dv)
        dv.add(f"E{r0 + 1}:E{rt - 1}")
    bloco(4, "Ações", ["pont", "qual", "saude", "tend"])        # linhas 5-8, total 9
    bloco(11, "ETFs", ["etf_pont", "etf_acoes", "etf_tend"])     # linhas 12-14, total 15
    bloco(17, "Tendência de longo prazo (ações e ETFs)", ["t10", "t5"])   # linhas 18-19, total 20
    ws["D16"] = USAR_TEND_TXT
    ws["D16"].alignment = Alignment(horizontal="right", vertical="center")
    ws["D16"].font = Font(bold=True, size=9)
    ws["E16"] = "x" if PESOS.get("usar_tend", 1) else None
    ws["E16"].fill, ws["E16"].border = amarelo, borda
    ws["E16"].font = Font(bold=True, size=12)
    ws["E16"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C16"] = ("Sem x: a tendência usa a reta mais fraca entre \"desde sempre\" e \"últimos 5 anos\" (como antes da "
                 "V040) e este quadro não conta.")
    ws["C16"].font = Font(italic=True, size=8, color="595959")
    ws["C16"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[16].height = 30
    # travões de "caro" (linhas 23-25): não somam 100%; cada um é um limite
    ws.cell(21, 2, "Avisos de preço caro (ações e ETFs): decisão a azul, não mexem na nota").font = Font(bold=True, size=12)
    ws["D21"] = USAR_CARO_TXT
    ws["D21"].alignment = Alignment(horizontal="right", vertical="center")
    ws["D21"].font = Font(bold=True, size=9)
    ws["E21"] = "x" if PESOS.get("usar_caro", 1) else None
    ws["E21"].fill, ws["E21"].border = amarelo, borda
    ws["E21"].font = Font(bold=True, size=12)
    ws["E21"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C21"] = "Sem x: não há avisos de cara (nem nas ações nem nos ETFs)."
    # travão do longo prazo
    ws.cell(27, 2, "Travão de longo prazo (nota geral no máx. 2,5 e \"Não Rent. Longo Prazo\")").font = Font(bold=True, size=12)
    ws["C28"] = ("Sem x (como sempre foi): trava se a reta MAIS FRACA entre \"desde sempre\" e \"últimos 5 anos\" ficar "
                 "abaixo do mínimo (G4 da folha Ações). Com x (mais brando): só trava se os últimos 10 anos E os últimos 5 "
                 "anos ficarem os dois abaixo do mínimo — não castiga quem caiu muito há mais de 10 anos (p. ex. 2008) e "
                 "recuperou.")
    ws["C28"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C28"].font = Font(size=9)
    ws.row_dimensions[28].height = 66
    ws["D28"] = LP_BRANDO_TXT
    ws["D28"].alignment = Alignment(horizontal="right", vertical="center")
    ws["D28"].font = Font(bold=True, size=9)
    ws["E28"] = "x" if PESOS.get("lp_brando") else None
    ws["E28"].fill, ws["E28"].border = amarelo, borda
    ws["E28"].font = Font(bold=True, size=12)
    ws["E28"].alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(30, 2, "Valor justo das empresas cíclicas").font = Font(bold=True, size=12)
    ws["C31"] = ("Com x: se o lucro dos últimos 12 meses estiver mais de 50% acima da média dos últimos 5 anos E os "
                 "analistas previrem que desce (lucro previsto < 90% do atual), o valor justo usa o lucro médio de 5 anos. "
                 "Sem x (como antes da V041): usa sempre o lucro dos últimos 12 meses. Muda o valor justo e por isso a "
                 "pontuação do sinal: só se aplica no ▶ ATUALIZAR (o ↻ RECLASSIFICAR não volta a descarregar os lucros).")
    ws["C31"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C31"].font = Font(size=9)
    ws.row_dimensions[31].height = 66
    ws["D31"] = PICO_TXT
    ws["D31"].alignment = Alignment(horizontal="right", vertical="center")
    ws["D31"].font = Font(bold=True, size=9)
    ws["E31"] = "x" if PESOS.get("pico_lucro") else None
    ws["E31"].fill, ws["E31"].border = amarelo, borda
    ws["E31"].font = Font(bold=True, size=12)
    ws["E31"].alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(33, 2, "PER habitual máximo no valor justo").font = Font(bold=True, size=12)
    ws["C34"] = ("O valor justo é o lucro dos últimos 12 meses × o PER habitual do próprio título (mediana de 3 anos). Nas "
                 "empresas que o mercado pagou muito caro (p. ex. NVIDIA), esse PER é enorme e o valor justo sai "
                 "exagerado. O nível de preço usa sempre o PER habitual limitado a este valor. Com x em \"Também no "
                 "sinal\", o limite aplica-se também ao valor justo da pontuação do sinal (só no ▶ ATUALIZAR; muda a "
                 "classificação de algumas ações de crescimento).")
    ws["C34"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C34"].font = Font(size=9)
    ws.row_dimensions[34].height = 80
    ws["B34"] = "PER habitual máximo"
    ws["B34"].font = Font(bold=True)
    ws["D34"] = PESOS_PROPOSTOS["pe_hab_max"]
    ws["D34"].font = Font(color="808080")
    ws["E34"] = PESOS["pe_hab_max"]
    ws["E34"].fill, ws["E34"].border = amarelo, borda
    ws["E34"].font = Font(bold=True, size=12)
    ws["E34"].alignment = ws["D34"].alignment = Alignment(horizontal="center", vertical="center")
    ws["D35"] = CAP_TXT
    ws["D35"].alignment = Alignment(horizontal="right", vertical="center")
    ws["D35"].font = Font(bold=True, size=9)
    ws["E35"] = "x" if PESOS.get("cap_pe_sinal") else None
    ws["E35"].fill, ws["E35"].border = amarelo, borda
    ws["E35"].font = Font(bold=True, size=12)
    ws["E35"].alignment = Alignment(horizontal="center", vertical="center")
    # --- fator de tendência (V104) ---
    ws.cell(37, 2, "Fator de tendência (multiplica a nota geral; substitui os travões do longo prazo)").font = \
        Font(bold=True, size=12)
    ws["C38"] = ("Para cada título calculam-se as retas (mínimos quadrados sobre o logaritmo das cotações, como nas "
                 "colunas Valorização) de 5 períodos: desde sempre, 10 anos e 5 anos com os fechos mensais; último ano e "
                 "6 meses com os fechos diários. A inclinação de cada reta é a subida média por ano. Faz-se a média "
                 "ponderada destas inclinações com os pesos em baixo, depois de: 0) tirar a cada uma a subida mínima por ano "
                 "(linhas 54-55: 5% nas ações, 2% nos ETFs), para o que sobe pouco contar como queda; 1) limitar cada "
                 "uma a ± o limite (as "
                 "retas curtas, passadas a ano, dão valores enormes); 2) multiplicar as QUEDAS pelo multiplicador (cair "
                 "pesa mais do que subir: não investir no que cai há 1 ou 2 anos); 3) com x, multiplicar as SUBIDAS do "
                 "último ano e dos 6 meses pela regularidade da reta (R², de 0 = aos saltos a 1 = sempre a subir): uma "
                 "subida errática (p. ex. +20% com quedas de 30% pelo meio) conta pouco. A média dá o fator pela escala "
                 "da linha 49, sem degraus. NOTA GERAL = nota × fator. Quem sobe de forma consistente tem 1,00 (não "
                 "muda); quem estagnou ou cai perde até 30%; com a volatilidade (linha 50), quem oscila muito perde até mais "
                 "15%; com o bónus de estabilidade (linha 56), quem oscila pouco e sobe com regularidade ganha até 9%. Com x, o travão de G4 (2,5 \"Não Rent. Longo Prazo\") e o "
                 "«estabilizou» da nota da tendência deixam de se usar. A coluna \"Longo prazo\" mostra o fator e "
                 "porquê. Também se usa no Simulador e no Comparar notas.")
    ws["C38"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C38"].font = Font(size=9)
    ws.row_dimensions[38].height = 200
    ws["D38"] = FT_TXT
    ws["D38"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    ws["D38"].font = Font(bold=True, size=9)
    ws["E38"] = "x" if PESOS.get("ft_usar", 1) else None
    for j, h in enumerate(("Reta", "O que representa", "Proposto", "A usar"), 2):
        c = ws.cell(39, j, h)
        c.font, c.fill, c.border = Font(bold=True, color="FFFFFF"), azul, borda
        c.alignment = Alignment(horizontal="center", vertical="center")
    expl = {"ft_s": "Toda a história: a empresa cria valor ao longo das décadas?",
            "ft_10": "Os últimos 10 anos (com menos história: não entra e os outros pesos sobem).",
            "ft_5": "Os últimos 5 anos: apanha quem abrandou ou estabilizou.",
            "ft_1": "O último ano (cotações diárias): o peso maior, para não comprar o que está a cair agora.",
            "ft_6": "Os últimos 6 meses (cotações diárias): a direção mais recente."}
    for i, (k, nome, _, _) in enumerate(FT_PER, 40):
        ws.cell(i, 2, nome).font = Font(bold=True)
        ws.cell(i, 3, expl[k]).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(i, 4, PESOS_PROPOSTOS[k]).font = Font(color="808080")
        ws.cell(i, 5, PESOS[k]).font = Font(bold=True, size=12)
        ws.cell(i, 5).fill = amarelo
        for j in range(2, 6):
            ws.cell(i, j).border = borda
        ws.cell(i, 4).number_format = ws.cell(i, 5).number_format = "0%"
        ws.cell(i, 4).alignment = ws.cell(i, 5).alignment = Alignment(horizontal="center", vertical="center")
    ws["B45"], ws["E45"] = "Total (tem de dar 100%)", "=SUM(E40:E44)"
    ws["B45"].font, ws["E45"].font, ws["E45"].number_format = Font(bold=True), Font(bold=True), "0%"
    ws["E45"].alignment = Alignment(horizontal="center")
    ws.conditional_formatting.add("E45", FormulaRule(formula=['ABS($E$45-1)>0.0005'],
                                                     fill=PatternFill("solid", bgColor="FF9999")))
    for i, (nome, txt_, kp, fm) in enumerate((
            ("Multiplicador das quedas", "Uma queda conta este número de vezes (2: cair 10%/ano pesa como −20%/ano).",
             "ft_mult", "0.0"),
            ("Limite de cada inclinação (±)", "Cada inclinação conta no máximo isto, a subir ou a descer, antes do "
             "multiplicador (as retas de 6 meses e 1 ano passadas a ano dão saltos enormes).", "ft_lim", "0%")), 46):
        ws.cell(i, 2, nome).font = Font(bold=True)
        ws.cell(i, 3, txt_).alignment = Alignment(wrap_text=True, vertical="center")
        ws.cell(i, 3).font = Font(size=9)
        ws.cell(i, 4, PESOS_PROPOSTOS[kp]).font = Font(color="808080")
        ws.cell(i, 5, PESOS[kp]).font = Font(bold=True, size=12)
        ws.cell(i, 5).fill = amarelo
        ws.cell(i, 4).number_format = ws.cell(i, 5).number_format = fm
        ws.cell(i, 4).alignment = ws.cell(i, 5).alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[i].height = 30
    ws["C48"] = ("Com x: as SUBIDAS do último ano e dos 6 meses contam × R² da reta (0 a 1). Uma subida regular conta "
                 "quase toda; uma subida aos saltos conta pouco. As quedas contam sempre inteiras (× o multiplicador).")
    ws["C48"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["C48"].font = Font(size=9)
    ws.row_dimensions[48].height = 30
    ws["D48"] = FT_R2_TXT
    ws["D48"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    ws["D48"].font = Font(bold=True, size=8)
    ws["E48"] = "x" if PESOS.get("ft_r2", 1) else None
    for c_ in ("E38", "E48"):
        ws[c_].fill, ws[c_].border = amarelo, borda
        ws[c_].font = Font(bold=True, size=12)
        ws[c_].alignment = Alignment(horizontal="center", vertical="center")
    ws["B49"] = "Escala do fator"
    ws["B49"].font = Font(bold=True)
    ws["C49"] = ("Média ponderada −20%/ano ou pior → 0,70 · 0%/ano → 0,85 · +5%/ano → 0,93 · +10%/ano ou mais → 1,00 "
                 "(em linha reta entre estes pontos). Ex.: média −2%/ano → 0,835: nota 4,0 passa a 3,34. Na coluna Longo "
                 "prazo: «penaliza» abaixo de 0,90; entre 0,90 e 1 diz «reduz ligeiramente».")
    ws["C49"].alignment = Alignment(wrap_text=True, vertical="center")
    ws["C49"].font = Font(size=9)
    ws.row_dimensions[49].height = 30
    ws["C50"] = ("Com x: o fator é ainda multiplicado por um fator de volatilidade, para contar com a INCERTEZA de um "
                 "título que oscila muito (sobe e desce aos saltos), mesmo que a tendência seja boa. Volatilidade = "
                 "desvio-padrão anual das variações diárias do último ano (um ETF do mundo anda por 12-15%, uma "
                 "grande empresa estável 20-25%, NVIDIA ≈ 40%, pequenas tecnológicas 60-90%). Até ao valor da linha 51 "
                 "não penaliza; acima, tira o valor da linha 52 por cada +10%; nunca tira mais de 15%. Quem sobe muito (média "
                 "ponderada das retas de +20%/ano ou mais, p. ex. Micron) não é penalizado; entre +10% e +20%, só em parte. Ex. (25% e 3%): "
                 "40% → × 0,955; 50% → × 0,925; 75% ou mais → × 0,85.")
    ws["C50"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C50"].font = Font(size=9)
    ws.row_dimensions[50].height = 75
    ws["D50"] = FT_VOL_TXT
    ws["D50"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    ws["D50"].font = Font(bold=True, size=8)
    ws["E50"] = "x" if PESOS.get("ft_vol", 1) else None
    ws["E50"].fill, ws["E50"].border = amarelo, borda
    ws["E50"].font = Font(bold=True, size=12)
    ws["E50"].alignment = Alignment(horizontal="center", vertical="center")
    for i, (nome, kp) in enumerate((("Volatilidade sem penalização (ações)", "ft_vol0"),
                                    ("Penalização por cada +10% de volatilidade", "ft_volpen"),
                                    ("Volatilidade sem penalização (ETFs)", "ft_vol0_etf"),
                                    ("Subida mínima por ano (ações)", "ft_sub"),
                                    ("Subida mínima por ano (ETFs)", "ft_sub_etf")), 51):
        ws.cell(i, 2, nome).font = Font(bold=True)
        ws.cell(i, 4, PESOS_PROPOSTOS[kp]).font = Font(color="808080")
        ws.cell(i, 5, PESOS[kp]).font = Font(bold=True, size=12)
        ws.cell(i, 5).fill = amarelo
        ws.cell(i, 4).number_format = ws.cell(i, 5).number_format = "0.0%"
        ws.cell(i, 4).alignment = ws.cell(i, 5).alignment = Alignment(horizontal="center", vertical="center")
    ws["C53"] = ("ETFs: um ETF de setor (p. ex. semicondutores) oscila mais do que uma ação grande sem ser mais arriscado "
                 "do que ela, porque está diversificado; por isso só penaliza acima deste valor.")
    ws["C54"] = ("Subida mínima exigida por ano: a cada inclinação tira-se este valor ANTES do multiplicador das quedas. "
                 "Um período que subiu menos do que isto conta como queda (× 2). Ex. (5%): uma reta de +3%/ano conta "
                 "−2% × 2 = −4%; uma de +12% conta +7%. A média volta depois a somar o valor e passa pela escala da "
                 "linha 49: quem sobe sempre mais de 10%/ano fica com 1,00; quem tem períodos abaixo do mínimo perde.")
    ws["C55"] = ("Nos ETFs, mais baixa: um ETF largo (mundo, EUA) rende historicamente 7 a 10%/ano e não deve ser "
                 "castigado por isso.")
    ws["C56"] = ("Com x: quem oscila POUCO e sobe com REGULARIDADE ganha um bónus (fator acima de 1; a nota nunca passa "
                 "de 5). Conta a volatilidade do último ano abaixo da linha 58 e a regularidade da reta do último ano "
                 "(R² de 0,3 → sem bónus a 0,8 → bónus inteiro), com a reta a subir. Bónus inteiro com volatilidade até "
                 "60% do limite (12% com o limite de 20%). Ex.: uma Exxon, J&J ou um ETF do mundo a subir com "
                 "regularidade e a oscilar 13-15% → × 1,05 a 1,09; uma ação a oscilar 30% → sem bónus.")
    ws["C56"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C56"].font = Font(size=9)
    ws.row_dimensions[56].height = 62
    ws["D56"] = FT_BONUS_TXT
    ws["D56"].alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
    ws["D56"].font = Font(bold=True, size=8)
    ws["E56"] = "x" if PESOS.get("ft_bonus", 1) else None
    ws["E56"].fill, ws["E56"].border = amarelo, borda
    ws["E56"].font = Font(bold=True, size=12)
    ws["E56"].alignment = Alignment(horizontal="center", vertical="center")
    for i, (nome, kp) in enumerate((("Bónus máximo de estabilidade", "ft_bmax"), ("Volatilidade abaixo da qual há bónus", "ft_bvol")), 57):
        ws.cell(i, 2, nome).font = Font(bold=True)
        ws.cell(i, 4, PESOS_PROPOSTOS[kp]).font = Font(color="808080")
        ws.cell(i, 5, PESOS[kp]).font = Font(bold=True, size=12)
        ws.cell(i, 5).fill = amarelo
        ws.cell(i, 4).number_format = ws.cell(i, 5).number_format = "0.0%"
        ws.cell(i, 4).alignment = ws.cell(i, 5).alignment = Alignment(horizontal="center", vertical="center")
    for c_, h_ in (("C53", 30), ("C54", 62), ("C55", 30)):
        ws[c_].alignment = Alignment(wrap_text=True, vertical="center")
        ws[c_].font = Font(size=9)
        ws.row_dimensions[int(c_[1:])].height = h_
    ws["H1"] = MARCA_PESOS
    ws["H1"].font = Font(size=7, color="BFBFBF")
    ws["H2"] = FT_MARCA
    ws["H2"].font = Font(size=7, color="BFBFBF")
    # --- nota global (V110) ---
    ws.cell(60, 2, "Nota global = nota geral + nível de preço (coluna \"Nota global\" das folhas Ações e ETFs)").font = \
        Font(bold=True, size=12)
    ws["C61"] = ("A nota geral mede o negócio, a tendência e o momento, mas não diz se o preço é bom para comprar. A NOTA "
                 "GLOBAL junta-lhe o nível de preço (5 muito barata … 1 muito cara): (nota geral × peso E62 + nível de "
                 "preço × peso E63) / (E62 + E63). Com 2 e 1: dois terços nota geral, um terço preço. Sem nível de preço "
                 "(p. ex. ETFs sem ações conhecidas), a nota global é a nota geral. A DECISÃO segue a nota global: ≥ 3,95 compra forte, "
                 "≥ 3,5 compra média, ≥ 3 compra (sem nível de preço: os limites da nota geral, 4,5 / 3,75 / 2,75); "
                 "como sempre, só há compra com o sinal ≥ 3 e venda com o sinal 1. Numa RECUPERAÇÃO depois de uma "
                 "queda (linhas 66-68) a decisão diz «(recuperação)»; o nível continua a ser o da nota global.")
    ws["C61"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C61"].font = Font(size=9)
    ws.row_dimensions[61].height = 90
    for i, (nome, kp, expl_) in enumerate((("Peso da nota geral", "g_nota", "Quanto conta a nota geral (negócio, tendência, momento)."),
                                           ("Peso do nível de preço", "g_preco", "Quanto conta o preço (valor justo, PER, FCF yield).")), 62):
        ws.cell(i, 2, nome).font = Font(bold=True)
        ws.cell(i, 3, expl_).font = Font(size=9)
        ws.cell(i, 4, PESOS_PROPOSTOS[kp]).font = Font(color="808080")
        e_ = ws.cell(i, 5, PESOS[kp])
        e_.fill, e_.border, e_.font = amarelo, borda, Font(bold=True, size=12)
        ws.cell(i, 4).number_format = e_.number_format = "0.0"
        ws.cell(i, 4).alignment = e_.alignment = Alignment(horizontal="center", vertical="center")
    ws["C64"] = calib_txt()
    ws["C64"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C64"].font = Font(size=9, italic=True, color="1F4E79")
    ws.row_dimensions[64].height = 92
    ws["B64"] = "Calibração com os analistas"
    ws["B64"].font = Font(bold=True)
    # --- recuperação: quanto tempo fica a compra forte (V148) ---
    ws.cell(66, 2, "Recuperação: quanto tempo fica assinalada").font = Font(bold=True, size=12)
    ws["C66"] = ("RECUPERAÇÃO (sinal 5, o fator ignora a queda do último ano e a decisão diz «(recuperação)»; o nível "
                 "continua a ser o da nota global): o título caiu do máximo de 1 ano (ações ≥ 15%, ETFs ≥ 20%, "
                 "e pelo menos 0,4 × a volatilidade), voltou a subir com força (reta dos últimos 60 dias ≥ +30%/ano) e está "
                 "acima da média 50 (ações: também da 200). Fica no dia do sinal e mais os dias de bolsa em baixo (≈ 21 por "
                 "mês), enquanto o preço estiver acima da média 50. "
                 "Testes (6 meses depois, face ao normal do título): AÇÕES com 20 dias +7,0% (81% a ganhar) nas 20 dos EUA e "
                 "+4,9% em 11 outras; com 45 dias +6,9% (82%) e +2,1%: a forte dura mais e apanha mais inícios de subida "
                 "(32% vs 27%), mas perde alguma vantagem fora da amostra. ETFs com 60 dias: +12,2% (78%). A vantagem está "
                 "sobretudo no dia do sinal e vai baixando. Os dias contam-se no ▶ ATUALIZAR.")
    ws["C66"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C66"].font = Font(size=9)
    ws.row_dimensions[66].height = 110
    for i, (nome, kp) in enumerate(((REC_JAN_TXT, "rec_jan"), (REC_JAN_ETF_TXT, "rec_jan_etf")), 67):
        ws.cell(i, 2, nome).font = Font(bold=True)
        ws.cell(i, 3, "dias de bolsa depois do sinal (0 = só nos dias do sinal; até " + str(REC_MEMO) + ")").font = Font(size=9)
        ws.cell(i, 4, PESOS_PROPOSTOS[kp]).font = Font(color="808080")
        e_ = ws.cell(i, 5, PESOS[kp])
        e_.fill, e_.border, e_.font = amarelo, borda, Font(bold=True, size=12)
        ws.cell(i, 4).number_format = e_.number_format = "0"
        ws.cell(i, 4).alignment = e_.alignment = Alignment(horizontal="center", vertical="center")
    # --- analistas no sinal (V160) ---
    ws.cell(70, 2, "Analistas no sinal").font = Font(bold=True, size=12)
    ws["D71"] = USAR_ANALISTAS_TXT
    ws["D71"].alignment = Alignment(horizontal="right", vertical="center")
    ws["D71"].font = Font(bold=True, size=9)
    ws["E71"] = None if PESOS.get("usar_analistas", 1) else "x"
    ws["E71"].fill, ws["E71"].border = amarelo, borda
    ws["E71"].font = Font(bold=True, size=12)
    ws["E71"].alignment = Alignment(horizontal="center", vertical="center")
    ws["C71"] = ("Sem x (por omissão): o sinal das ações usa a tendência, o valor justo E os analistas (as duas condições). "
                 "Com x: ignora o preço-alvo dos analistas (as condições deles contam como cumpridas; a coluna fica a "
                 "cinzento, só indicativa), como nas simulações do passado. ATENÇÃO: sem analistas o sinal fica mais "
                 "permissivo (ex.: recuperações que os analistas travavam passam a contar). Antes de usar, testa: põe x e "
                 "carrega em ▶ GERAR EVOLUÇÃO (usa este interruptor), compara com a versão sem x. Teste de 12 datas "
                 "(2023-2025): a ordem pelos analistas não acertou melhor do que a do valor justo, mas o efeito nas "
                 "decisões não foi medido.")
    ws["C71"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C71"].font = Font(size=9)
    ws.row_dimensions[71].height = 80
    ws["C21"].font = Font(italic=True, size=8, color="595959")
    ws["C21"].alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[21].height = 30
    for j, h in enumerate(("Travão", "O que representa", "Proposto", "A usar"), 2):
        c = ws.cell(22, j, h)
        c.font, c.fill, c.border = Font(bold=True, color="FFFFFF"), azul, borda
        c.alignment = Alignment(horizontal="center", vertical="center")
    for i, (k, nome, texto, fm) in enumerate(TRAVOES_CARO, 23):
        ws.cell(i, 2, nome).font = Font(bold=True)
        ws.cell(i, 3, texto).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(i, 4, PESOS_PROPOSTOS[k]).font = Font(color="808080")
        e = ws.cell(i, 5, PESOS[k])
        e.fill, e.font = amarelo, Font(bold=True, size=12)
        for j in range(2, 6):
            ws.cell(i, j).border = borda
        ws.cell(i, 4).number_format = ws.cell(i, 5).number_format = fm
        ws.cell(i, 4).alignment = ws.cell(i, 5).alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[i].height = 44
    dvt = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=False)
    dvt.error, dvt.errorTitle = "Escreve um valor maior do que zero.", "Travão"
    ws.add_data_validation(dvt)
    dvt.add("E23")
    ws.merge_cells("G4:I5")
    botao(ws, "G4", "↻ RECLASSIFICAR E REORDENAR", f"{PROTOCOLO}://reclassificar", "1F3B57", "0F2236", 12)
    ws["G4"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["G6"] = ("Depois de mudares as percentagens, carrega aqui: volta a calcular as notas e as decisões das folhas "
                "Ações e ETFs com as percentagens novas e reordena cada folha pelo seu critério, sem descarregar "
                "nada (é rápido). A app também é atualizada.")
    ws["G6"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["G6"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("G6:I9")
    for col in "GHI":
        ws.column_dimensions[col].width = 14
    for col, w in zip("ABCDE", (2, 30, 70, 11, 11)):
        ws.column_dimensions[col].width = w
    cz, fz = PatternFill("solid", bgColor="808080"), Font(color="BFBFBF")   # sem x: o que depende dele a cinzento
    ws.conditional_formatting.add("D18:E19", FormulaRule(formula=['LOWER(TRIM($E$16))<>"x"'], fill=cz, font=fz))
    ws.conditional_formatting.add(f"D23:E{22 + len(TRAVOES_CARO)}",
                                  FormulaRule(formula=['LOWER(TRIM($E$21))<>"x"'], fill=cz, font=fz))
    ws.conditional_formatting.add("D40:E50", FormulaRule(formula=['LOWER(TRIM($E$38))<>"x"'], fill=cz, font=fz))
    for rng_ in ("D51:E52", "D53:E53"):
        ws.conditional_formatting.add(rng_, FormulaRule(formula=['OR(LOWER(TRIM($E$38))<>"x",LOWER(TRIM($E$50))<>"x")'],
                                                        fill=cz, font=fz))
    ws.conditional_formatting.add("D54:E56", FormulaRule(formula=['LOWER(TRIM($E$38))<>"x"'], fill=cz, font=fz))
    ws.conditional_formatting.add("D57:E58", FormulaRule(formula=['OR(LOWER(TRIM($E$38))<>"x",LOWER(TRIM($E$56))<>"x")'],
                                                         fill=cz, font=fz))
    ws.conditional_formatting.add("D28:E28", FormulaRule(formula=['LOWER(TRIM($E$38))="x"'], fill=cz, font=fz))


def texto_pesos(ranking=True):
    return (f"{pc1(PESOS['pont'])} sinal + {pc1(PESOS['qual'])} qualidade + {pc1(PESOS['saude'])} saúde + "
            f"{pc1(PESOS['tend'])} tendência de longo prazo" if ranking else
            f"{pc1(PESOS['etf_pont'])} sinal + {pc1(PESOS['etf_acoes'])} nota das ações do ETF + {pc1(PESOS['etf_tend'])} "
            f"tendência de longo prazo") + " (folha Ponderações)"


def sobre_compra(d):
    """Nos títulos detidos com preço médio de compra: « · +85% sobre a tua compra: realizar ganho?» (só informação;
    as regras do sinal são as mesmas para todos)."""
    pc, pr = d.get("pcompra"), d.get("preco")
    q = d.get("qtd") if isinstance(d.get("qtd"), (int, float)) else qtd_detida(d.get("det"))
    if not (q or e_detido(d.get("det"))) or not isinstance(pc, (int, float)) or not isinstance(pr, (int, float)) or pc <= 0:
        return ""
    r = pr / pc - 1
    t = f" · {r:+.0%} sobre a tua compra".replace("-", "−")
    return t + (": realizar ganho?" if r > 0 else ": cortar a perda?")


GERAL_LIM = (3.95, 3.5, 3.0)
LIM_ACOES = GERAL_LIM          # ações: nota global (com o nível de preço) -> compra forte / média / compra
LIM_ETFS = (4.5, 3.75, 2.75)   # ETFs e ações sem nível de preço: a nota global é a nota geral (escala não comprimida)
COMPRA_NAS_DESCIDAS = True    # V189: abaixo só da média de 50 dias já não é motivo para esperar
MARGEM_M200 = 0.05            # V191: até 5% abaixo da média de 200 dias também se compra (teste 2002-2025, 5 520 casos:
                              # +9,0% a 1 ano contra +7,5% à espera da média 200 e +10,1% a comprar sempre; 10% piores
                              # −26%, −23% e −28%)
COR_ESPERA = "D9C3E9"          # «Compra …: esperar (…)»: lilás (o azul já é «cara»)     # nota global: compra forte / média / compra (a média com o preço puxa para o meio)


def compra_espera(d):
    """«Compra forte/média/Compra: esperar (…)» quando a empresa e o preço valem a compra (nota global calculada com o
    sinal neutro, 3) mas o momento ainda não (sinal < 3). Diz o que esperar: que estabilize (preço acima da média de
    50 dias) e/ou que chegue à compra máxima (ordem limite)."""
    if d.get("_espera_calc"):
        return None
    d2 = dict(d, pont=3.0, _espera_calc=True)
    try:
        nota_global(d2)
    except Exception:
        return None
    ng, nt = d2.get("nota_geral"), d2.get("nota")
    n_conf = ng if isinstance(ng, (int, float)) and isinstance(d2.get("nivel_preco"), (int, float)) else nt
    if isinstance(ng, (int, float)) and isinstance(d2.get("nivel_preco"), (int, float)):
        nivel = 5 if ng >= GERAL_LIM[0] else 4 if ng >= GERAL_LIM[1] else 3 if ng >= GERAL_LIM[2] else None
    elif isinstance(nt, (int, float)):
        nivel = 5 if nt >= 4.5 else 4 if nt >= 3.75 else 3 if nt >= 2.75 else None
    else:
        nivel = None
    if nivel is None:
        return None
    pr = d.get("preco") if isinstance(d.get("preco"), (int, float)) else d.get("_preco_sinal")
    m50, teto, m200 = d.get("mm50"), d.get("teto"), d.get("mm200c")
    falta = []
    n2 = lambda v: f"{v:,.2f}".replace(",", " ").replace(".", ",")
    lim200 = m200 * (1 - MARGEM_M200) if isinstance(m200, (int, float)) else None
    if isinstance(pr, (int, float)) and lim200 is not None and pr < lim200:
        falta.append(f"recuperar: que o preço volte para cima de {n2(lim200)} = média de 200 dias −{MARGEM_M200:.0%} "
                     f"(agora {n2(pr)}, {pr / m200 - 1:+.0%} face à média)")
    if isinstance(pr, (int, float)) and isinstance(m50, (int, float)) and pr < m50:
        falta.append(f"estabilizar: que o preço volte para cima da média de 50 dias (agora {n2(pr)} < {n2(m50)}, "
                     f"{pr / m50 - 1:+.0%})")
    # V192: a compra máxima é o preço da ordem limite, sempre ~1-2% abaixo do preço (ver niveis): já não conta como
    # «falta» (contava sempre, e por isso a compra nas descidas nunca chegava a aparecer)
    pre = f"nota {n_conf:.2f} ao confirmar; ".replace(".", ",") if isinstance(n_conf, (int, float)) else ""
    abaixo200 = isinstance(pr, (int, float)) and isinstance(m200, (int, float)) and pr < m200
    if COMPRA_NAS_DESCIDAS and (falta or abaixo200) and all(x.startswith("estabilizar") for x in falta):
        # V189: só abaixo da média de 50 dias (acima da de 200 e abaixo da compra máxima) = descida numa empresa boa:
        # compra-se (teste 2002-2025, 191 ações, 3 223 descidas: comprar logo rendeu +11,6% a 1 ano contra +9,0% à
        # espera de voltar acima da média 50, com os casos maus iguais)
        nc_ = f"; nota {n_conf:.2f}".replace(".", ",") if isinstance(n_conf, (int, float)) else ""
        ds_ = [f"{pr / m50 - 1:+.0%} face à média de 50 dias"] if isinstance(m50, (int, float)) and pr < m50 else []
        ds_ += [f"{pr / m200 - 1:+.0%} face à de 200"] if abaixo200 else []
        return f"{NIVEIS[nivel]} (em descida: {', '.join(ds_)}{nc_})"
    if COMPRA_NAS_DESCIDAS:
        falta = [x for x in falta if not x.startswith("estabilizar")] or falta
    if COMPRA_NAS_DESCIDAS and not falta:
        # V192: preço acima das médias (ou perto da de 200) e o sinal ainda por confirmar: sem indício forte de que
        # continue a descer, compra-se
        nc_ = f"; nota {n_conf:.2f}".replace(".", ",") if isinstance(n_conf, (int, float)) else ""
        return f"{NIVEIS[nivel]} (sinal ainda a confirmar{nc_})"
    if not falta:
        falta.append("confirmar a subida (sinal ≥ 3)")
    return f"{NIVEIS[nivel]}: esperar ({pre}{'; '.join(falta)})"


def decisao_global(d):
    """Decisão pela nota global (V114; antes pela nota geral), sem contrariar o sinal: só há compra se a pontuação do sinal for ≥ 3 (e então o
    nível vem da nota geral: ≥ 4,5 forte, ≥ 3,75 média, ≥ 2,75 compra, abaixo disso manter); venda se o sinal for 1."""
    p, nt = d.get("pont"), d.get("nota")
    if not isinstance(p, (int, float)):
        return None
    alerta = txt(d.get("saude")).startswith("ALERTA")
    sobre = sobre_compra(d)
    espera = compra_espera(d) if p < 3 and not alerta else None    # boa e barata, mas ainda não é altura
    tem = e_detido(d.get("det")) or bool(d.get("qtd"))
    if espera and not (tem and (p == 1 or d.get("_vigiar"))):      # nos detidos, a venda/vigiar tem prioridade
        return espera
    if p == 1:
        return NIVEIS[1] + f" ({d.get('_venda_tipo') or 'tendência de descida'})" + sobre
    if d.get("_vigiar") and p == 2:
        return "Manter (queda brusca: vigiar)" + sobre
    if p == 2.5:
        return NIVEIS[2.5] + (" + saúde frágil" if alerta else "")
    cara = ""                                          # "cara" aparece na coluna do nível de preço
    if p < 3:
        return ("Manter (venda próxima)" if d.get("_vprox") else NIVEIS[2]) + cara
    if alerta:
        return "Saúde frágil"
    nt = nt if isinstance(nt, (int, float)) else p
    ng = d.get("nota_geral")
    if isinstance(ng, (int, float)) and isinstance(d.get("nivel_preco"), (int, float)):   # pela nota global (com preço)
        nivel = 5 if ng >= GERAL_LIM[0] else 4 if ng >= GERAL_LIM[1] else 3 if ng >= GERAL_LIM[2] else None
    else:                                              # sem nível de preço: a nota global é a global, limites de sempre
        nivel = 5 if nt >= 4.5 else 4 if nt >= 3.75 else 3 if nt >= 2.75 else None
    mom = momento_compra(d)                            # V151: a recuperação não muda o nível; só fica indicada
    if nivel is None:
        return ("Manter (nota global baixa: negócio, tendência ou preço)" if isinstance(d.get("nivel_preco"), (int, float))
                else "Manter (qualidade / saúde fracas)") + cara
    motivo = f" ({mom})" if mom else ""
    return NIVEIS[nivel] + motivo + (" (antecipada)" if d.get("_cprox") else "") + cara


def nota_global(d):
    """Ações: sinal + qualidade + saúde + tendência LP; ETFs: sinal + nota ponderada das ações (pesos: folha Ponderações).
    O que faltar conta como a pontuação. Travões: saúde em alerta ou longo prazo fraco -> no máx. 2,5."""
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    p = n("pont")
    if p is None:
        return None
    lp = n("lp")
    t = nota_tendencia(lp, n("lp10"), n("lp5"))
    if txt(d.get("tipo")) == "ETF":
        ea = n("etf_acoes")
        nota = (PESOS["etf_pont"] * p + PESOS["etf_acoes"] * (ea if ea is not None else p)
                + PESOS["etf_tend"] * (t if t is not None else p))
    else:
        q, sn = n("qual"), n("saude_n")
        nota = (PESOS["pont"] * p + PESOS["qual"] * (q if q is not None else p)
                + PESOS["saude"] * (sn if sn is not None else p) + PESOS["tend"] * (t if t is not None else p))
    lpt = lp_travao(d)
    ft = fator_tend_atual(d)
    d["fator_tend"] = ft[0] if ft else None
    if ft:                                        # fator de tendência: multiplica a nota (em vez do travão do LP)
        nota = min(5.0, nota * ft[0])
        base_lp = re.sub(r":.*$", "", str(d.get("lp_txt") or ""))
        if base_lp:
            d["lp_txt"] = base_lp + ": " + fator_txt(ft)
    outro = txt(d.get("saude")).startswith("ALERTA") or (not ft and not PESOS.get("ft_usar", 1)
                                                         and lpt is not None and lpt < LP_MIN)
    if txt(d.get("tipo")) != "ETF":
        d["nivel_preco"] = nivel_preco(d)[0]
    # cara: só um aviso (azuis na nota, no nível de preço e na decisão); não trava a nota
    d["so_caro"] = 1 if motivo_caro(d) else None
    d["nota_sc"] = None
    if outro:
        nota = min(nota, NOTA_ALERTA)
    d["nota_geral"] = nota_geral(nota, d.get("nivel_preco"))
    return round(nota, 2)


def nota_geral(nota, nivel, pesos=None):
    """Nota global = (nota geral × g_nota + nível de preço × g_preco) / (g_nota + g_preco); sem nível: a nota geral."""
    if not isinstance(nota, (int, float)):
        return None
    gn, gp = pesos if pesos else (PESOS.get("g_nota", 2.0), PESOS.get("g_preco", 1.0))
    if not isinstance(nivel, (int, float)) or gn + gp <= 0:
        return round(nota, 2)
    return round((gn * nota + gp * nivel) / (gn + gp), 2)


REC_NOTA = {"strong_buy": 1.0, "buy": 2.0, "hold": 3.0, "underperform": 4.0, "sell": 5.0, "strong_sell": 5.0}


def calibrar_geral(acoes):
    """Que pesos (nota geral : nível de preço) fazem a nota global acertar mais com o conselho dos analistas?
    Correlação de ordens (Spearman) entre a nota global e a recomendação média dos analistas (1 compra forte … 5 venda;
    sem ela, a recomendação), nas ações com pelo menos 3 analistas. Só informa (não muda os pesos)."""
    import pandas as pd
    lin = []
    for d in acoes:
        nt, nv = d.get("nota"), d.get("nivel_preco")
        rm = d.get("rec_media") if isinstance(d.get("rec_media"), (int, float)) else REC_NOTA.get(txt(d.get("rec")))
        na = d.get("nanal") if isinstance(d.get("nanal"), (int, float)) else 0
        if isinstance(nt, (int, float)) and isinstance(nv, (int, float)) and rm and na >= 3:
            lin.append((nt, nv, -rm))
    CALIB.clear()
    if len(lin) < 15:
        return CALIB
    df = pd.DataFrame(lin, columns=["nt", "nv", "an"])
    corr = lambda w: float(((1 - w) * df["nt"] + w * df["nv"]).corr(df["an"], method="spearman"))
    import math
    tab = [(round(w, 2), corr(w)) for w in [i / 20 for i in range(0, 21)]]
    tab = [(w, r) for w, r in tab if not math.isnan(r)]
    if not tab:                                   # recomendações todas iguais: não há o que comparar
        return CALIB
    w_b, r_b = max(tab, key=lambda x: x[1])
    gn, gp = PESOS.get("g_nota", 2.0), PESOS.get("g_preco", 1.0)
    w_a = gp / (gn + gp) if gn + gp > 0 else 0
    CALIB.update(n=len(lin), w_melhor=w_b, r_melhor=r_b, r_atual=corr(w_a), r_nota=corr(0.0), r_preco=corr(1.0),
                 w_atual=w_a, data=datetime.now().strftime("%d/%m/%Y %H:%M"), tabela=tab)
    _gravar_json(FICH_CALIB, CALIB)
    return CALIB


def calib_txt():
    c = CALIB or _ler_json(FICH_CALIB, {})
    if not c.get("n"):
        return "Ainda sem calibração: faz um ▶ ATUALIZAR (precisa de pelo menos 15 ações com 3 ou mais analistas)."
    f = lambda r: f"{r:+.2f}".replace(".", ",")
    wb_ = c["w_melhor"]
    razao = "só nota geral" if wb_ == 0 else "só preço" if wb_ == 1 else f"{(1 - wb_) / wb_:.1f} : 1".replace(".", ",")
    return (f"Calibração com os analistas ({c['data']}, {c['n']} ações com 3 ou mais analistas; correlação de ordens com a "
            f"recomendação média, de −1 a +1, maior = acerta mais): com os pesos atuais {f(c['r_atual'])}; só a nota "
            f"global {f(c['r_nota'])}; só o nível de preço {f(c['r_preco'])}; MELHOR: nota geral : preço = {razao} "
            f"({1 - wb_:.0%} / {wb_:.0%}) com {f(c['r_melhor'])}. Só informa: os pesos são os das células E62 e E63. "
            "Atenção: a nota geral já inclui o potencial dos analistas (no sinal), por isso acerta sempre um pouco; "
            "e os analistas também se enganam — não convém afinar demasiado para eles.")
Q_ROIC, Q_GPA, Q_FCFL = (0.05, 0.15), (0.10, 0.30), (0.5, 0.9)
S_DIV, S_DIV_REG, S_JUROS = (3.0, 1.0), (5.0, 2.0), (3.0, 8.0)
SETORES_REG = ("Real Estate", "Utilities")


def _grad(v, mau, bom):
    """0 no 'mau', 1 no 'bom', em linha reta entre os dois (funciona nos dois sentidos)."""
    if v is None:
        return None
    return max(0.0, min(1.0, (v - mau) / (bom - mau)))


_CAMBIO = {}


def fundamentos(tk, info, moeda_cap):
    """Indicadores de saúde e qualidade de uma ação (dicionário com as chaves das colunas)."""
    num = _num
    out = {}
    setor, ind = txt(info.get("sector")), txt(info.get("industry"))
    if setor in ("Financial Services", "Financial") and "Data" not in ind:
        out["saude"] = "não se aplica (setor financeiro: banco, seguro ou financeira)"
        return out
    menor = {"GBp": "GBP", "GBX": "GBP", "ZAc": "ZAR", "ILA": "ILS"}
    fin = menor.get(txt(info.get("financialCurrency")), txt(info.get("financialCurrency"))) or None
    mc = menor.get(moeda_cap, moeda_cap) or None
    cap = num(info.get("marketCap"))
    if cap and fin and mc and fin != mc:                       # capitalização -> moeda das contas
        falta = {m for m in (fin, mc) if m not in _CAMBIO}
        if falta:
            _CAMBIO.update({m: v[0] for m, v in taxas_eur(falta).items()})
        r1, r2 = _CAMBIO.get(mc), _CAMBIO.get(fin)
        cap = cap * r1 / r2 if (r1 and r2) else None

    def ultimo(df, nomes):
        try:
            for n in nomes:
                if df is not None and n in df.index:
                    v = df.loc[n].dropna()
                    if len(v):
                        return float(v.iloc[0])
        except Exception:
            pass
        return None
    def relatorio(nome):
        for tent in range(2):
            try:
                return getattr(tk, nome)
            except Exception:
                time.sleep(2)
        return None
    inc, bs, cf = relatorio("income_stmt"), relatorio("balance_sheet"), relatorio("cashflow")
    ebitda, ni = num(info.get("ebitda")), num(info.get("netIncomeToCommon"))
    # FCF e lucro do último relatório anual (os dois do mesmo ano); o "freeCashflow" do Yahoo é o FCF "levered",
    # que às vezes fica muito abaixo do real (na Nvidia dava FCF/lucro 0,2) -> só se usa se faltar o relatório
    fcf_a = ultimo(cf, ["Free Cash Flow"])
    if fcf_a is None:
        ocf, capex = ultimo(cf, ["Operating Cash Flow"]), ultimo(cf, ["Capital Expenditure"])
        fcf_a = ocf + capex if (ocf is not None and capex is not None) else None
    ni_a = ultimo(inc, ["Net Income", "Net Income Common Stockholders"])
    fcf = fcf_a if fcf_a is not None else num(info.get("freeCashflow"))
    if fcf_a is not None and ni_a is not None:
        ni = ni_a
    divida, caixa = num(info.get("totalDebt")), num(info.get("totalCash"))
    ebit = ultimo(inc, ["EBIT", "Operating Income"])
    juros = ultimo(inc, ["Interest Expense", "Interest Expense Non Operating"])
    bruto = ultimo(inc, ["Gross Profit"]) or num(info.get("grossProfits"))
    imp, antes = ultimo(inc, ["Tax Provision"]), ultimo(inc, ["Pretax Income"])
    ativo = ultimo(bs, ["Total Assets"])
    capital = ultimo(bs, ["Invested Capital"])
    if divida is None:
        divida = ultimo(bs, ["Total Debt"])
    if caixa is None:
        caixa = ultimo(bs, ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"])
    if not capital:
        cp = ultimo(bs, ["Stockholders Equity", "Common Stock Equity"])
        capital = (cp + (divida or 0) - (caixa or 0)) if cp is not None else None
    liquida = (divida or 0) - (caixa or 0) if (divida is not None or caixa is not None) else None
    taxa = min(max(imp / antes, 0.0), 0.35) if (imp is not None and antes and antes > 0) else 0.21
    if ebit is not None and capital and capital > 0:
        out["roic"] = ebit * (1 - taxa) / capital
    if bruto is not None and ativo:
        out["gpa"] = bruto / ativo
    if fcf is not None and ni and ni > 0:
        out["fcf_lucro"] = fcf / ni
    if fcf is not None and cap:
        out["fcf_yield"] = fcf / cap
    if cap and ebitda and ebitda > 0:
        out["ev_ebitda"] = (cap + (liquida or 0)) / ebitda
    if liquida is not None and ebitda and ebitda > 0:
        out["divl_ebitda"] = liquida / ebitda
    if ebit is not None and juros and abs(juros) > 0:
        out["ebit_juros"] = ebit / abs(juros)
    # saúde (gradual) e alerta
    regulada = setor in SETORES_REG
    mau_d, bom_d = S_DIV_REG if regulada else S_DIV
    pts, alerta, ok = [], [], []
    dv, cob = out.get("divl_ebitda"), out.get("ebit_juros")
    fmt = lambda v: f"{v:.1f}".replace(".", ",")
    if liquida is not None and liquida <= 0:
        pts.append(1.0)
        ok.append("mais caixa do que dívida")
    elif dv is not None:
        pts.append(_grad(dv, mau_d, bom_d))
        (alerta if dv > mau_d else ok).append(f"dívida líq. {fmt(dv)}× EBITDA" + (" (setor regulado)" if regulada else ""))
    elif ebitda is not None and ebitda <= 0 and liquida and liquida > 0:
        pts.append(0.0)
        alerta.append("EBITDA negativo com dívida")
    if cob is not None:
        pts.append(_grad(cob, *S_JUROS))
        (alerta if cob < S_JUROS[0] else ok).append(f"juros cobertos {fmt(cob)}×")
    elif ebit is not None and juros is not None and abs(juros) == 0:
        pts.append(1.0)
        ok.append("sem juros")
    if fcf is not None:
        caixa_liq = liquida is not None and liquida <= 0
        pts.append(1.0 if fcf >= 0 else 0.5 if caixa_liq else 0.0)
        if fcf >= 0:
            ok.append("FCF positivo")
        elif caixa_liq:
            ok.append("FCF negativo, mas com caixa líquida")
        else:
            alerta.append("FCF negativo com dívida")
    if pts:
        out["saude_n"] = round(5 * sum(pts) / len(pts), 1)
    if alerta:
        out["saude"] = "ALERTA: " + "; ".join(alerta) + ((" (" + "; ".join(ok) + ")") if ok else "")
    elif ok:
        out["saude"] = "OK: " + "; ".join(ok)
    else:
        out["saude"] = "sem dados no Yahoo"
    # qualidade do negócio (gradual)
    q = [x for x in (_grad(out.get("roic"), *Q_ROIC), _grad(out.get("gpa"), *Q_GPA), _grad(out.get("fcf_lucro"), *Q_FCFL))
         if x is not None]
    if len(q) >= 2:
        out["qual"] = round(5 * sum(q) / len(q), 1)
    return out


def detalhe_qualidade(d):
    """Texto curto: os 3 critérios do negócio (pontos 0-1) e, à parte, o preço (FCF yield, EV/EBITDA)."""
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    partes = []
    for k, rot, lim, f in (("roic", "ROIC", Q_ROIC, "{:.0%}"), ("gpa", "bruto/ativo", Q_GPA, "{:.0%}"),
                           ("fcf_lucro", "FCF/lucro", Q_FCFL, "{:.1f}")):
        v = n(k)
        if v is not None:
            partes.append(f"{rot} {f.format(v)} = {_grad(v, *lim):.1f}".replace(".", ","))
    preco = [f"FCF yield {n('fcf_yield'):.1%}".replace(".", ",") if n("fcf_yield") is not None else None,
             f"EV/EBITDA {n('ev_ebitda'):.0f}×" if n("ev_ebitda") is not None else None]
    preco = [x for x in preco if x]
    return ", ".join(partes) + ((" · preço: " + ", ".join(preco)) if preco else "")


# ---------- composição dos ETFs ----------
# O Yahoo dá a descrição e as maiores posições (até 10) de muitos ETFs; nos UCITS europeus falta muitas vezes. Nesse
# caso usa-se um ETF americano do mesmo índice (só para a composição, indicado no texto). Obrigações e ouro: não.
COMPOSICAO = {}       # ticker do ETF -> {"desc": str, "hold": [[símbolo, nome, peso], ...], "proxy": str|None}
PROXY_ETF = [("health care", "XLV"), ("information technology", "XLK"), ("semicond", "SMH"),
             ("automation & robotics", "BOTZ"), ("nasdaq 100", "QQQ"), ("nasdaq-100", "QQQ"),
             ("small cap value", "VBR"), ("world small cap", "VSS"), ("s&p 500", "VOO"),
             ("high dividend", "VYMI"), ("all-world", "VT"), ("all country world", "ACWI"), ("acwi", "ACWI"),
             ("msci world", "URTH"), ("emerging", "VWO"), ("japan", "EWJ"), ("euro stoxx 50", "FEZ"),
             ("stoxx europe 600", "VGK"), ("developed europe", "VGK"), ("msci europe", "VGK"), ("dax", "EWG"),
             ("cac 40", "EWQ"), ("ftse 100", "EWU"), ("smi", "EWL"), ("ibex", "EWP"), ("ftse mib", "EWI"),
             ("aex", "EWN"), ("omxs30", "EWD")]
SEM_ACOES = ("bond", "obrigaç", "treasury", "gold", "ouro", "govern")
FICH_COMP = os.path.join(PASTA, BASE + "_composicao_etfs.json")
FICH_ACOES_ETF = os.path.join(PASTA, BASE + "_acoes_dos_etfs.json")


def _ler_json(f, vazio):
    try:
        with open(f, encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return vazio


def _gravar_json(f, dados):
    try:
        with open(f, "w", encoding="utf-8") as fh:
            json.dump(dados, fh, ensure_ascii=False, indent=1)
    except Exception as e:
        print("   aviso: não foi possível gravar", os.path.basename(f), e)


def composicao(yf, tk, ticker_orig, ticker, nome, info):
    """Descrição curta e maiores posições de um ETF (do próprio, da listagem em euros ou de um ETF americano do mesmo
    índice)."""
    out = {"desc": None, "hold": [], "proxy": None}
    nl = nome.lower()
    sem_acoes = any(x in nl for x in SEM_ACOES)
    proxy = None if sem_acoes else next((px for chave, px in PROXY_ETF if chave in nl), None)
    cands = [(tk, None)] + ([(yf.Ticker(ticker), None)] if ticker != ticker_orig else []) + \
        ([(yf.Ticker(proxy), proxy)] if proxy else [])
    for t, px in cands:
        desc = th = None
        try:
            fd = t.funds_data
            desc = fd.description
            th = fd.top_holdings
        except Exception:
            pass
        if desc and not out["desc"] and px is None:
            out["desc"] = desc
        if sem_acoes or th is None or not len(th):
            continue
        hold = []
        for sym, row in th.iterrows():
            try:
                w = float(row.get("Holding Percent"))
            except Exception:
                continue
            sym = txt(sym)
            if sym and w > 0 and " " not in sym and len(sym) <= 14:
                hold.append([sym, txt(row.get("Name")) or sym, w])
        if hold:
            out["hold"], out["proxy"] = hold[:10], px
            break
    if not out["desc"]:
        out["desc"] = txt(info.get("longBusinessSummary")) or None
    if out["desc"]:                                        # breve: a 1.ª frase (até ~220 caracteres)
        dd = re.split(r"(?<=[.!?])\s", out["desc"].strip())[0]
        out["desc"] = dd if len(dd) <= 220 else dd[:217].rsplit(" ", 1)[0] + "…"
    return out


def _chaves_acao(d):
    """Símbolos pelos quais uma ação da folha Ações pode aparecer na composição de um ETF."""
    ch = {txt(d.get("codigo")).upper(), txt(d.get("ticker")).upper()}
    m = re.search(r"principal: .* (\S+)$", txt(d.get("bolsa")))
    if m:
        ch.add(m.group(1).upper())
    ch |= {c.split(".")[0] for c in ch if c}
    return {c for c in ch if c}


def _acao_de(sym, indice, nome=None):
    s = sym.upper()
    return indice.get(s) or indice.get(s.split(".")[0]) or (indice.get(_nome_chave(nome)) if nome else None)


REPETIDOS = set()        # tickers que o recolher deixou de fora por serem a mesma empresa de outro título


_SEM_NOME = {"the", "inc", "corp", "corporation", "co", "company", "class", "and", "sa", "nv", "plc", "ag", "se", "ltd",
             "limited", "spa", "ab", "asa", "as", "oyj", "a", "b", "c", "adr", "ord", "shs", "reg", "registered", "nom",
             "de", "pref", "com", "holdings", "holding", "group", "the"}


def _nome_chave(n):
    """Nome sem formas jurídicas nem classes (para reconhecer a mesma empresa com outro ticker, p. ex. GOOG / GOOGL)."""
    w = [x for x in re.findall(r"[a-z0-9]+", txt(n).lower()) if x not in _SEM_NOME]
    return ("nome:" + " ".join(w)) if w else None


def acoes_em_falta(linhas, etfs, comp):
    """Ações das composições dos ETFs que não estão na folha Ações: [(nome, código, ticker, bolsa), ...]."""
    indice = {}
    for d in linhas:
        for c in _chaves_acao(d) | {_nome_chave(d.get("nome"))} - {None}:
            indice.setdefault(c, d)
    falta, vistos = [], {x.upper() for x in comp.get("_sem_dados", []) + comp.get("_apagados", [])}  # sem dados / apagadas
    for e in etfs:
        c = comp.get(txt(e.get("ticker"))) or comp.get(txt(e.get("_orig")))
        for sym, nm, w in (c or {}).get("hold", []):
            if _acao_de(sym, indice, nm) is None and sym.upper() not in vistos:
                vistos.add(sym.upper())
                falta.append((nm, sym.split(".")[0], sym, nome_bolsa(sym)))
    return falta


def ponderar_etfs(linhas, etfs, comp):
    """Nota ponderada das ações de cada ETF (notas gerais da folha Ações, pesadas pelo peso no ETF), descrição e
    composição; recalcula a nota geral dos ETFs."""
    indice = {}
    for d in linhas:
        for c in _chaves_acao(d) | {_nome_chave(d.get("nome"))} - {None}:
            indice.setdefault(c, d)
    for e in etfs:
        c = comp.get(txt(e.get("ticker"))) or comp.get(txt(e.get("_orig"))) or {}
        if c.get("desc"):
            e["etf_desc"] = c["desc"]
        partes, soma_w, soma_wn, total, soma_wc, soma_wcn = [], 0.0, 0.0, 0.0, 0.0, 0.0
        soma_wb, soma_wbn = 0.0, 0.0
        for sym, nm, w in c.get("hold", []):
            total += w
            a = _acao_de(sym, indice, nm)
            nt = a.get("nota") if a is not None and isinstance(a.get("nota"), (int, float)) else None
            if nt is not None:
                soma_w += w
                soma_wn += w * nt
                if isinstance(a.get("nivel_preco"), (int, float)):
                    soma_wc += w * a["nivel_preco"]
                    soma_wcn += w
            if a is not None and isinstance(a.get("beta"), (int, float)):
                soma_wb += w * a["beta"]
                soma_wbn += w
            partes.append(f"{nm} {w:.1%}".replace(".", ",") + (f" ({nt:.1f})".replace(".", ",") if nt is not None else ""))
        if partes:
            e["etf_comp"] = ((f"(aprox. pelo {c['proxy']}, mesmo índice) " if c.get("proxy") else "")
                             + "; ".join(partes) + f" · estas {len(partes)}: {total:.0%} do ETF".replace(".", ","))
        e["etf_acoes"] = round(soma_wn / soma_w, 2) if soma_w > 0 else None
        e["etf_caro"] = None
        e["nivel_preco"] = round(soma_wc / soma_wcn, 1) if soma_wcn >= 0.15 else None
        if soma_wbn >= 0.15:                            # beta do ETF: média das ações conhecidas (senão o do Yahoo)
            e["beta"] = round(soma_wb / soma_wbn, 2)
        e["nota"] = nota_global(e)
        e["decisao"] = decisao_global(e)
        mc = motivo_caro(e)
        if mc:
            base = re.sub(r" · cara: [^·]*", "", txt(e.get("coment")))
            e["coment"] = (base + " · " if base else "") + mc.replace(".", ",") + " (só aviso: a azul)"


def dados_titulo(yf, nome, codigo, ticker_orig, bolsa, etf, hist, eur, tx_orig):
    """Todos os dados de um título (uma linha do painel). hist: {ticker: velas diárias}."""
    ticker = eur.get(ticker_orig) or ticker_orig        # ticker do preço (de preferência em euros)
    h_orig = None
    if ticker != ticker_orig:
        try:
            h_eu = hist[ticker].dropna(subset=["Close"])
            h_orig = hist[ticker_orig].dropna(subset=["Close"])
            if len(h_eu) < min(300, len(h_orig) - 20):    # história curta em euros: fica a bolsa principal
                ticker, h_orig = ticker_orig, None
        except Exception:
            ticker, h_orig = ticker_orig, None
    d = {"nome": nome, "tipo": "ETF" if etf else "Ação", "codigo": codigo, "ticker": ticker, "bolsa": bolsa}
    lp = LONGO.get(ticker_orig) or LONGO.get(ticker)
    if lp:
        f = lambda v: f"{v:+.1%}".replace(".", ",")
        d["lp"] = lp[0]
        s10 = lp[5] if len(lp) > 5 else None
        d["lp10"] = s10 if s10 is not None else lp[2]      # menos de 10 anos de história: conta desde sempre
        d["lp5"] = lp[3] if lp[3] is not None else lp[2]
        s3 = lp[6] if len(lp) > 6 else None
        d["val_sempre"], d["val10"], d["val5"], d["val3"], d["val1"] = lp[2], s10, lp[3], s3, lp[4]
        d["lp_txt"] = (f"{f(lp[2])}/ano desde {lp[1]}" + (f" ; 10 anos {f(s10)}/ano" if s10 is not None else "")
                       + (f" ; 5 anos {f(lp[3])}/ano" if lp[3] is not None else "")
                       + (f" ; 3 anos {f(s3)}/ano" if s3 is not None else "")
                       + (f" ; 1 ano {f(lp[4])}" if lp[4] is not None else ""))
    if ticker != ticker_orig:
        d["bolsa"] = f"{nome_bolsa(ticker)} ({moeda_prevista(ticker)}) · principal: {bolsa} {ticker_orig}"
    ultimo = None
    h = None
    try:
        h = hist[ticker].dropna(subset=["Close"])
        if len(h):
            fim = h.index.max()
            ultimo = float(h["Close"].iloc[-1])
            for i, (_, dias) in enumerate(PERIODOS):
                w = h[h.index > fim - timedelta(days=dias)]
                mx, mn, med = float(w["High"].max()), float(w["Low"].min()), float(w["Close"].mean())
                d.update({f"max{i}": mx, f"min{i}": mn, f"med{i}": med,
                          f"cima{i}": mx / med - 1, f"baixo{i}": mn / med - 1})
            cl = h["Close"]
            d.update(indicadores(cl))
            d.update(inclin_recentes(cl))
            if len(cl) > 126:
                d["ref6m"] = float(cl.iloc[-127])
            if len(cl) >= 2:
                d["var1d"] = float(cl.iloc[-1] / cl.iloc[-2] - 1)
            mes = cl[cl.index <= fim - timedelta(days=30)]
            if len(mes):
                d["var1m"] = float(cl.iloc[-1] / mes.iloc[-1] - 1)
            w3 = (h_orig if h_orig is not None else h).tail(63)   # negócio: o da bolsa principal
            if "Volume" in w3:
                d["volume"] = _num(w3["Volume"].mean())
                d["_valor_local"] = _num((w3["Volume"] * w3["Close"]).mean())
    except Exception:
        pass
    def ler_info(t):
        for tent in range(3):
            try:
                return yf.Ticker(t).info or {}
            except Exception:
                time.sleep(2 * (tent + 1))
        return {}
    info = ler_info(ticker)
    preco, d["preco_hora"], ref_dia = preco_mais_recente(info, ultimo)
    if REF is not None and ultimo:                      # situação numa data passada: o preço é o fecho dessa data
        agora = preco
        preco, ref_dia = ultimo, None
        d["preco_hora"] = f"fecho {h.index[-1]:%d/%m/%Y}"
        if agora:
            d["desde_ref"] = float(agora) / ultimo - 1
    conv = 1.0                                          # bolsa principal -> euros (analistas e valor justo)
    tk = yf.Ticker(ticker)
    if ticker != ticker_orig:
        info_p = info
        info = ler_info(ticker_orig)                    # analistas, lucros e capitalização: bolsa principal
        tk = yf.Ticker(ticker_orig)
        moeda_o = info.get("currency") or moeda_prevista(ticker_orig)
        preco_o = _num(info.get("regularMarketPrice") or info.get("currentPrice"))
        if REF is not None and h_orig is not None and len(h_orig):
            preco_o = float(h_orig["Close"].iloc[-1])   # numa data passada: os fechos dessa data
        if preco and preco_o:                           # conversão pelos dois preços: câmbio e, nos ADR, o múltiplo
            conv = float(preco) / preco_o
        else:
            if moeda_o not in tx_orig:
                tx_orig.update(taxas_eur({moeda_o}))
            conv = tx_orig.get(moeda_o, (None, None))[1] or 0
        d["_moeda_neg"] = info.get("currency") or moeda_prevista(ticker_orig)
        info = {**info, "currency": info_p.get("currency", "EUR")}
    if not etf and preco and not info.get("targetMeanPrice"):
        # sem analistas nesta listagem (p. ex. uma ação francesa acrescentada pela Xetra): procura a bolsa principal
        princ = fonte_analistas(yf, nome, ticker_orig, float(preco), info.get("currency"))
        if princ:
            info_f = ler_info(princ)
            preco_f = _num(info_f.get("regularMarketPrice") or info_f.get("currentPrice"))
            if preco_f and info_f.get("targetMeanPrice"):
                try:
                    ini = ((REF or datetime.now()) - timedelta(days=max(d for _, d in PERIODOS) + 10))
                    h_orig = yf.Ticker(princ).history(start=ini.strftime("%Y-%m-%d"), auto_adjust=False)
                    if h_orig.index.tz is not None:
                        h_orig.index = h_orig.index.tz_localize(None)
                    if REF is not None:
                        h_orig = h_orig[h_orig.index <= REF]
                        preco_f = float(h_orig["Close"].iloc[-1]) if len(h_orig) else preco_f
                except Exception:
                    h_orig = None
                tk, conv = yf.Ticker(princ), float(preco) / preco_f
                d["_moeda_neg"] = info_f.get("currency")
                info = {**info_f, "currency": info.get("currency")}
                d["_fonte"] = princ
    if preco and ref_dia:
        d["var1d"] = float(preco) / ref_dia - 1
    if etf:   # ETFs não têm preço-alvo nem recomendação de analistas
        for chave in ("targetMeanPrice", "numberOfAnalystOpinions", "recommendationKey"):
            info.pop(chave, None)
    alvo = info.get("targetMeanPrice")
    nota = ""
    nanal = info.get("numberOfAnalystOpinions")
    if REF is not None and not etf:                    # numa data passada: alvo reconstruído com o histórico
        a_h, n_h = alvo_em(tk, REF)
        if a_h:
            alvo, nanal = a_h, n_h
            nota = f"alvo de {REF:%d/%m/%Y} reconstruído ({n_h} corretoras)"
        else:
            nota = "alvo de hoje (o Yahoo não tem o histórico dos analistas deste título)"
    if alvo and conv != 1.0:
        alvo = float(alvo) * conv
        nota = ((nota + "; ") if nota else "") + (f"preço na bolsa {nome_bolsa(ticker)}; alvo e valor justo do "
                                                  f"{d.pop('_fonte', None) or ticker_orig} convertidos")
    d.pop("_fonte", None)
    if preco and alvo:
        preco, alvo = float(preco), float(alvo)
        if alvo < preco / 20:
            alvo, nota = alvo * 100, "alvo convertido de GBP para pence"
        elif alvo > preco * 20:
            alvo, nota = alvo / 100, "alvo convertido de pence para GBP"
    if not preco:
        nota = "sem dados no Yahoo"
    if not etf and preco and h is not None and len(h) > 60:
        if conv != 1.0 and h_orig is not None and _num(info.get("regularMarketPrice") or info.get("currentPrice")):
            VJ_PE_HAB[0] = None
            justo, metodo = valor_justo(tk, h_orig, _num(info.get("regularMarketPrice") or info.get("currentPrice")),
                                        info)
            justo = justo * conv if justo else None
        else:
            VJ_PE_HAB[0] = None
            justo, metodo = valor_justo(tk, h, float(preco), info)
        if justo:
            d.update(justo=justo, justo_pct=justo / float(preco) - 1,
                     pe_hab=VJ_PE_HAB[0] if metodo and metodo.startswith("P/E") else None)
            if metodo == "Graham":
                nota = (nota + "; " if nota else "") + "valor justo pelo número de Graham"
            elif metodo and "pico" in metodo:
                nota = (nota + "; " if nota else "") + "valor justo com o lucro médio de 5 anos (o lucro atual está no pico)"
        else:                                           # dizer porque não há valor justo
            eps = _num(info.get("trailingEps"))
            pq = ("lucros dos últimos 12 meses negativos ou nulos (não há P/E)" if eps is not None and eps <= 0 else
                  "o Yahoo não tem os lucros deste título" if eps is None else
                  "lucros sem histórico suficiente e valor contabilístico em falta ou negativo")
            nota = (nota + "; " if nota else "") + "sem valor justo: " + pq
    if not etf and preco:                               # saúde financeira e qualidade (contas da bolsa principal)
        try:
            d.update(fundamentos(tk, info, d.get("_moeda_neg") or info.get("currency")))
        except Exception as e:
            d["saude"] = "sem dados no Yahoo"
            registo(f"fundamentos {ticker_orig}: {e}")
    try:                                                # dividendos (valor na moeda desta cotação = yield × preço)
        px_base = _num(info.get("currentPrice")) or _num(info.get("regularMarketPrice")) or _num(info.get("previousClose"))
        taxa = _num(info.get("dividendRate"))
        dy = taxa / px_base if (taxa and px_base) else _num(info.get("trailingAnnualDividendYield"))
        if etf and not dy:
            dy = _num(info.get("yield"))
        if dy and 0 < dy < 0.5:
            d["div_yield"] = dy
            d["div"] = dy * preco if preco else None
        po = _num(info.get("payoutRatio"))
        if po and po > 0 and not etf:
            d["payout"] = po
    except Exception:
        pass
    pe = _num(info.get("trailingPE"))
    d["pe"] = pe if pe and pe > 0 else None             # sem P/E: lucros negativos ou sem dados
    pf = _num(info.get("forwardPE"))
    d["pe_fwd"] = pf if pf and 0 < pf < 500 else None   # P/E com os lucros previstos para os próximos 12 meses
    d["_cap_local"] = _num(info.get("marketCap"))
    bt = _num(info.get("beta")) if not etf else _num(info.get("beta3Year"))
    d["beta"] = round(bt, 2) if bt is not None and -3 < bt < 6 else None   # ETFs: depois pelas ações (ponderar_etfs)
    d.update(moeda=info.get("currency", ""), preco=preco, alvo=alvo,
             pot=(alvo / preco - 1) if (preco and alvo) else None,
             nanal=nanal, rec=info.get("recommendationKey"), rec_media=_num(info.get("recommendationMean")), notas=nota)
    for i in range(len(PERIODOS)):
        mx = d.get(f"max{i}")
        d[f"dist{i}"] = (float(preco) / mx - 1) if (preco and mx) else None
    try:
        d.update(ordens(h, float(preco) if preco else None, None))
        niveis(d, float(preco) if preco else None)
        if d.get("mm200") and preco:
            d["dist200"] = float(preco) / d["mm200"] - 1      # com o preço mais recente
        alta, desde = d.pop("_alta", None), d.pop("_desde", None)
        if alta is not None:
            d["tend"] = ("↑ subida" if alta else "↓ descida") + (f" desde {desde:%d/%m/%Y}" if desde is not None else "")
        pontuar(d, float(preco) if preco else None)
    except Exception as e:
        print(f"   aviso: {codigo}: ordens não calculadas ({e})")
    d["_info_vazia"] = not info
    if etf:                                             # descrição e composição (para a nota ponderada das ações)
        try:
            c = composicao(yf, tk, ticker_orig, ticker, nome, info)
            COMPOSICAO[ticker_orig] = COMPOSICAO[ticker] = c
            if c.get("desc"):
                d["etf_desc"] = c["desc"]
        except Exception as e:
            registo(f"composição {ticker_orig}: {e}")
    return d


def _por_ticker(df, tickers):
    """Resultado do yf.download -> {ticker: DataFrame} (com um ou vários tickers, colunas por ticker ou por campo)."""
    import pandas as pd
    out = {}
    if df is None or not len(df):
        return out
    if isinstance(df.columns, pd.MultiIndex):
        for t in tickers:
            for nivel in range(df.columns.nlevels):
                if t in df.columns.get_level_values(nivel):
                    try:
                        out[t] = df.xs(t, axis=1, level=nivel)
                    except Exception:
                        pass
                    break
    elif len(tickers) == 1:
        out[tickers[0]] = df
    return out


def _incompleto(d):
    """Faltam dados que deviam vir (preço, história, médias ou, numa ação, a informação do Yahoo)?"""
    return (not d.get("preco") or d.get("max0") is None or d.get("mm200c") is None
            or (d.get("tipo") != "ETF" and d.get("_info_vazia")))


def valores_eur(linhas):
    """Valor transacionado e capitalização em euros (para comparar bolsas com moedas diferentes)."""
    tx = taxas_eur({d.get("moeda") for d in linhas} | {d.get("_moeda_neg") for d in linhas})
    for d in linhas:
        f_maior, f_preco = tx.get(d.pop("_moeda_neg", None) or d.get("moeda"), (None, None))
        v, c = d.pop("_valor_local", None), d.pop("_cap_local", None)
        d["valor_tx"] = v * f_preco / 1e6 if (v and f_preco) else None
        d["cap"] = c * f_maior / 1e9 if (c and f_maior) else None
        d.pop("_info_vazia", None)


def incluir_novos(novos):
    """Depois do ➕ ACRESCENTAR: vai buscar já os dados dos títulos novos e põe-nos nas folhas Ações / ETFs (as outras
    linhas ficam como estão, com os valores do último ▶ ATUALIZAR / ↻ PREÇOS)."""
    global REF, LP_MIN
    _, _, est0 = ler_estado()
    LP_MIN = est0.get("lp_min", LP_MIN)
    REF = data_ref(est0.get("ref_escolha"), est0.get("ref_data"))
    yf = _yf()
    ex = Excel(FICHEIRO)
    ultima = get_column_letter(len(COLS))
    r = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)]
               + [{"op": "ler", "folha": 1, "rng": C_AT}])
    chaves = [k for k, *_ in COLS]
    folhas = {}
    for idx in (1, 2):
        tabela = r[idx - 1]
        if not cab_ok([txt(v) for v in tabela[0]]):
            raise SystemExit("As colunas do painel não são as desta versão: carrega em ▶ ATUALIZAR (os títulos já "
                             "estão na folha Acrescentados e entram aí).")
        folhas[idx] = [linha_lida(chaves, v) for v in tabela[1:] if txt(linha_lida(chaves, v).get("ticker"))]
    titulos = [(n, c, t, b) for n, c, t, b, tipo in novos]
    print(f"A ir buscar os dados de {len(novos)} título(s) novo(s)...")
    eur = bolsas_alt([(n, c, t, b, tipo == "ETF") for n, c, t, b, tipo in novos])
    ja = {txt(d.get("ticker")).upper() for ls in folhas.values() for d in ls}
    rep = [a for a in novos if {a[2].upper(), (eur.get(a[2]) or a[2]).upper()} & ja]
    for a in rep:
        print(f"   {a[0]} ({a[2]}) já está no painel (repetido): não entra outra vez")
    novos = [a for a in novos if a not in rep]
    if not novos:
        return
    titulos = [(n, c, t, b) for n, c, t, b, tipo in novos]
    pedir = [t for _, _, t, _ in titulos] + sorted(set(eur.values()))
    inicio = ((REF or datetime.now()) - timedelta(days=max(d for _, d in PERIODOS) + 10)).strftime("%Y-%m-%d")
    try:
        hist = _cortar(_por_ticker(yf.download(pedir, start=inicio, interval="1d", group_by="ticker",
                                               auto_adjust=False, threads=True, progress=False), pedir))
    except Exception:
        hist = {}
    tx_orig = taxas_eur({moeda_prevista(t) for t in eur} | {"USD"})
    LONGO.update(longo_prazo(yf, [t for _, _, t, _, _ in novos]))
    linhas = []
    for (n, c, t, b, tipo) in novos:
        d = dados_titulo(yf, n, c, t, b, tipo == "ETF", hist, eur, tx_orig)
        if _incompleto(d):                                    # verificação: outra vez, só este
            try:
                p1 = [t] + ([eur[t]] if t in eur else [])
                h1 = _por_ticker(yf.download(p1, start=inicio, interval="1d", group_by="ticker", auto_adjust=False,
                                             threads=False, progress=False), p1)
                d = dados_titulo(yf, n, c, t, b, tipo == "ETF", {**hist, **_cortar(h1)}, eur, tx_orig)
            except Exception:
                pass
        linhas.append(d)
        print(f"   {c:8} preço {d.get('preco')} | alvo {d.get('alvo')}")
    valores_eur(linhas)
    for d in linhas:
        folhas[2 if d.get("tipo") == "ETF" else 1].append(d)
    pontuar_etfs(folhas[2])
    criterios, acrescentados, estado = ler_estado()
    base = re.sub(r" · \+\d+ acrescentado.*$", "", txt(r[2][0][0]) or "Atualizado")
    texto = base + f" · +{len(novos)} acrescentado(s) em {datetime.now():%d/%m/%Y %H:%M}"
    wb = construir(folhas[1], texto, folhas[2], criterios, acrescentados, estado)
    guardar_e_abrir(wb)
    registo(f"acrescentados ao painel: {', '.join(t for _, _, t, _, _ in novos)}")


def recolher(acrescentados=(), so=None, apenas=None, filtro=(None, None)):
    """filtro = (nomes/tickers das ações, dos ETFs) a ir buscar; None = todos."""
    import logging
    import warnings
    import yfinance as yf
    # o yfinance imprime "Failed download / possibly delisted / HTTP Error 404" para códigos sem dados;
    # não fazem mal (a linha fica com "sem dados no Yahoo") — calam-se aqui e há um resumo no fim
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)
    warnings.filterwarnings("ignore")
    # listas do script + títulos da folha Acrescentados (sem repetir tickers)
    vistos = {a[2].upper() for a in ACOES + ETFS}
    extra_a = [(n, c, t, b) for n, c, t, b, tipo in acrescentados if tipo != "ETF" and t.upper() not in vistos]
    extra_e = [(n, c, t, b) for n, c, t, b, tipo in acrescentados if tipo == "ETF" and t.upper() not in vistos]
    lista_a, lista_e = ACOES + extra_a, ETFS + extra_e
    if so == "acoes":
        lista_e, extra_e = [], []
    elif so == "etfs":
        lista_a, extra_a = [], []
    if apenas is not None:                     # só estas ações (as das composições dos ETFs que faltam)
        lista_a, lista_e, extra_a, extra_e = list(apenas), [], list(apenas), []
    else:                                      # só os títulos que estão nas folhas (e, numa atualização parcial,
        dentro = lambda f, n, t: f is None or n in f or t.upper() in f     # só os escolhidos)
        fa, fe = filtro
        lista_a = [x for x in lista_a if dentro(fa, x[0], x[2])]
        lista_e = [x for x in lista_e if dentro(fe, x[0], x[2])]
        extra_a = [x for x in extra_a if dentro(fa, x[0], x[2])]
        extra_e = [x for x in extra_e if dentro(fe, x[0], x[2])]
    if extra_a or extra_e:
        print(f"Títulos acrescentados (folha Acrescentados + ações dos ETFs): {len(extra_a)} ações e {len(extra_e)} ETFs.")
    todos = lista_a + lista_e
    modo = modo_atual()
    lst = [(n, c, t, b, k >= len(lista_a)) for k, (n, c, t, b) in enumerate(todos)]
    if modo == "todos":                        # cada título na bolsa da lista, em euros e nos EUA (uma linha cada)
        alt_e, alt_u = bolsas_alt(lst, "eur"), bolsas_alt(lst, "eua")
        variantes = lambda t: [None, alt_e.get(t), alt_u.get(t)]
    else:
        eur0 = bolsas_alt(lst, modo)
        variantes = lambda t: [eur0.get(t)]
    # repetidos: a mesma ação duas vezes (p. ex. NVDA na lista e NVD.DE acrescentada) -> fica a primeira
    vistos_t, sem_rep, alt_de, n_a = set(), [], [], 0
    rep = []
    for k, (n, c, t, b) in enumerate(todos):
        entrou = False
        for a in variantes(t):
            final = (a or t).upper()
            chaves = {final} if modo == "todos" else {t.upper(), final}
            if chaves & vistos_t:
                continue
            vistos_t |= chaves
            sem_rep.append((n, c, t, b))
            alt_de.append(a)
            n_a += k < len(lista_a)
            entrou = True
        if not entrou:
            rep.append(t)
            REPETIDOS.add(t.upper())
    lista_a, lista_e = sem_rep[:n_a], sem_rep[n_a:]
    todos = sem_rep
    alt_dic = lambda k: {todos[k][2]: alt_de[k]} if alt_de[k] else {}
    if rep:
        print(f"   {len(rep)} título(s) repetido(s) (a mesma empresa com outro ticker) não entra(m): "
              + ", ".join(rep[:12]) + (" …" if len(rep) > 12 else ""))
    print(f"Bolsa: {MODOS[modo]}.")
    tickers = sorted({a[2] for a in todos} | {a for a in alt_de if a})
    LONGO.update(longo_prazo(yf, [a[2] for a in todos]))
    print("A descarregar cotações diárias dos últimos 3 anos...")
    inicio = ((REF or datetime.now()) - timedelta(days=max(d for _, d in PERIODOS) + 10)).strftime("%Y-%m-%d")
    hist = _cortar(_por_ticker(yf.download(tickers, start=inicio, interval="1d", group_by="ticker",
                                           auto_adjust=False, threads=True, progress=False), tickers))
    tx_orig = taxas_eur({moeda_prevista(todos[k][2]) for k in range(len(todos)) if alt_de[k]} | {"USD"})
    linhas = []
    for k, (nome, codigo, ticker_orig, bolsa) in enumerate(todos, start=1):
        etf = k > len(lista_a)
        d = dados_titulo(yf, nome, codigo, ticker_orig, bolsa, etf, hist, alt_dic(k - 1), tx_orig)
        linhas.append(d)
        print(f"{k:3}/{len(todos)} {codigo:8} preço {d.get('preco')} | alvo {d.get('alvo')}")
        time.sleep(0.3)
    # verificação: títulos com dados em falta (p. ex. o Yahoo falhou no pedido em conjunto) -> pede-os um a um, até 2 vezes
    for volta in (1, 2):
        faltam = [k for k, d in enumerate(linhas) if _incompleto(d)]
        if not faltam:
            break
        print(f"\nVerificação ({volta}/2): {len(faltam)} título(s) com dados em falta; a pedir outra vez um a um...")
        for k in faltam:
            nome, codigo, ticker_orig, bolsa = todos[k]
            pedir = [ticker_orig] + ([alt_de[k]] if alt_de[k] else [])
            h1 = {}
            for tent in range(2):
                try:
                    h1 = _por_ticker(yf.download(pedir, start=inicio, interval="1d", group_by="ticker",
                                                 auto_adjust=False, threads=False, progress=False), pedir)
                except Exception:
                    h1 = {}
                if h1:
                    break
                time.sleep(3)
            if not h1:                                         # último recurso: a história do próprio título
                try:
                    h1 = {t: yf.Ticker(t).history(start=inicio, auto_adjust=False) for t in pedir}
                except Exception:
                    h1 = {}
            novo = dados_titulo(yf, nome, codigo, ticker_orig, bolsa, k >= len(lista_a), {**hist, **_cortar(h1)},
                                alt_dic(k), tx_orig)
            if (not _incompleto(novo)) or sum(v is not None for v in novo.values()) > sum(v is not None for v in linhas[k].values()):
                linhas[k] = novo
            print(f"   {codigo:8} {'completo' if not _incompleto(linhas[k]) else 'ainda incompleto'}")
            time.sleep(1)
    pontuar_etfs(linhas)
    for d in linhas:                                           # o que faltar depois disto: dizer porquê, nas notas
        falta = []
        if not d.get("preco"):
            falta.append("preço")
        elif d.get("mm200c") is None:
            falta.append("médias (menos de 200 dias de cotações)" if d.get("max0") is not None else "cotações")
        if d.get("tipo") != "ETF" and d.get("preco") and d.get("alvo") is None:
            falta.append("preço-alvo (o Yahoo não tem analistas para este título)")
        if falta and d.get("notas") != "sem dados no Yahoo":
            d["notas"] = ((d.get("notas") or "") + "; " if d.get("notas") else "") + "falta: " + ", ".join(falta)
        d.pop("_info_vazia", None)
    falhas = [d["ticker"] for d in linhas if d["notas"] == "sem dados no Yahoo"]
    if falhas:
        print(f"\nSem dados no Yahoo ({len(falhas)}): {', '.join(falhas)} — confirma estes códigos nas listas do script.")
    valores_eur(linhas)
    acoes, etfs = linhas[:len(lista_a)], linhas[len(lista_a):]
    return acoes, etfs   # a ordem é aplicada em construir() conforme o "Ordenar por" de cada folha




# ================================================================================================ ligação ao Excel aberto
PS = r"""
$ErrorActionPreference = 'Stop'
$req = Get-Content -Raw -Encoding UTF8 -LiteralPath '@@IN@@' | ConvertFrom-Json
$wb = $null
# 1) procura o livro em TODOS os Excel abertos (tabela de objetos em execução do Windows)
try {
  Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
public static class RotGerarDados {
  [DllImport("ole32.dll")] static extern int GetRunningObjectTable(int r, out IRunningObjectTable p);
  [DllImport("ole32.dll")] static extern int CreateBindCtx(int r, out IBindCtx p);
  public static object Procurar(string nome) {
    IRunningObjectTable rot; IEnumMoniker en; IBindCtx ctx;
    GetRunningObjectTable(0, out rot); rot.EnumRunning(out en); CreateBindCtx(0, out ctx);
    IMoniker[] m = new IMoniker[1];
    while (en.Next(1, m, IntPtr.Zero) == 0) {
      string dn = null;
      try { m[0].GetDisplayName(ctx, null, out dn); } catch { continue; }
      if (dn == null) continue;
      if (dn.EndsWith("\\" + nome, StringComparison.OrdinalIgnoreCase) || dn.EndsWith("/" + nome, StringComparison.OrdinalIgnoreCase)) {
        object o; rot.GetObject(m[0], out o); return o;
      }
    }
    return null;
  }
}
"@ -ErrorAction SilentlyContinue
  $wb = [RotGerarDados]::Procurar($req.livro)
} catch { $wb = $null }
# 2) senão, o Excel "ativo"
$xl = $null
if (-not $wb) {
  try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { exit 3 }
  foreach ($b in $xl.Workbooks) { if ($b.Name -ieq $req.livro) { $wb = $b } }
}
if (-not $wb) {
  try { foreach ($pv in $xl.ProtectedViewWindows) { if ($pv.SourceName -ieq $req.livro) { exit 4 } } } catch {}
  exit 3
}
$res = New-Object System.Collections.ArrayList
foreach ($o in $req.ops) {
  if ($o.op -eq 'autosave') { try { $wb.AutoSaveOn = [bool]$o.valor } catch {}; [void]$res.Add($null); continue }
  if ($o.op -eq 'ativa') { [void]$res.Add([int]$wb.ActiveSheet.Index); continue }
  if ($o.op -eq 'fechar') { try { if (-not $wb.Saved) { $wb.Save() } } catch {}; $wb.Close($false); [void]$res.Add($null); break }
  if ($o.op -eq 'quadro' -or $o.op -eq 'fechar_quadro') {
    foreach ($sh in $wb.Worksheets) {
      $nomes = @(); foreach ($x in $sh.Shapes) { if ($x.Name -like 'Quadro*') { $nomes += $x.Name } }
      foreach ($nm in $nomes) { $sh.Shapes.Item($nm).Delete() }
    }
    if ($o.op -eq 'quadro') {
      foreach ($fi in $o.folhas) {
        $sh = $wb.Worksheets.Item([int]$fi); $rg = $sh.Range($o.ancora)
        $tb = $sh.Shapes.AddTextbox(1, $rg.Left, $rg.Top, [double]$o.largura, [double]$o.altura)
        $tb.Name = 'QuadroCarteira'
        $tb.TextFrame2.TextRange.Text = [string]$o.texto
        $tb.TextFrame2.TextRange.Font.Size = 10; $tb.TextFrame2.TextRange.Font.Bold = 1
        $tb.Fill.ForeColor.RGB = 13431551; $tb.Line.ForeColor.RGB = 8421504; $tb.Placement = 3
        $x = $sh.Shapes.AddShape(9, $rg.Left + [double]$o.largura - 20, $rg.Top + 3, 16, 16)
        $x.Name = 'QuadroFechar'; $x.TextFrame2.TextRange.Text = 'x'; $x.TextFrame2.TextRange.Font.Size = 9
        $x.Fill.ForeColor.RGB = 255; $x.Placement = 3
        [void]$sh.Hyperlinks.Add($x, 'painelacoes://fechar_quadro')
      }
    }
    [void]$res.Add($null); continue
  }
  if ($o.op -eq 'ocultos') {
    if ($o.folha -is [string]) { $sh = $wb.Worksheets.Item($o.folha) } else { $sh = $wb.Worksheets.Item([int]$o.folha) }
    $oc = New-Object System.Collections.ArrayList
    for ($r = [int]$o.r0; $r -lt [int]$o.r0 + [int]$o.max; $r++) {
      $v = $sh.Cells.Item($r, [int]$o.col).Value2
      if (-not $v) { break }
      if ($sh.Rows.Item($r).Hidden) { [void]$oc.Add([string]$v) }
    }
    [void]$res.Add($oc); continue
  }
  if ($o.op -eq 'ler_filtros') {
    $out = New-Object System.Collections.ArrayList
    foreach ($sh in $wb.Worksheets) {
      try {
        if (-not $sh.AutoFilterMode) { continue }
        $af = $sh.AutoFilter; $rg = $af.Range; $fl = New-Object System.Collections.ArrayList
        for ($i = 1; $i -le $af.Filters.Count; $i++) {
          $f = $af.Filters.Item($i)
          if (-not $f.On) { continue }
          $op = [int]$f.Operator
          if ($op -ge 8 -and $op -le 10) { continue }
          $c2 = $null; try { $c2 = $f.Criteria2 } catch {}
          $cab = [string]$sh.Cells.Item($rg.Row, $rg.Column + $i - 1).Value2
          [void]$fl.Add(@{cab = $cab.Trim(); op = $op; c1 = $f.Criteria1; c2 = $c2})
        }
        if ($fl.Count -gt 0) { [void]$out.Add(@{folha = $sh.Name; r = $rg.Row; c = $rg.Column; nr = $rg.Rows.Count; nc = $rg.Columns.Count; filtros = $fl}) }
      } catch {}
    }
    [void]$res.Add($out); continue
  }
  if ($o.op -eq 'copia') {
    try { $wb.SaveCopyAs([string]$o.destino) } catch {}
    [void]$res.Add($null); continue
  }
  if ($o.op -eq 'reaplicar_filtro') {
    try {
      if ($o.folha -is [string]) { $sh = $wb.Worksheets.Item($o.folha) } else { $sh = $wb.Worksheets.Item([int]$o.folha) }
      if ($sh.AutoFilterMode -and $sh.FilterMode) { [void]$sh.AutoFilter.ApplyFilter() }
    } catch {}
    [void]$res.Add($null); continue
  }
  if ($o.op -eq 'mostrar_tudo') {
    try {
      if ($o.folha -is [string]) { $sh = $wb.Worksheets.Item($o.folha) } else { $sh = $wb.Worksheets.Item([int]$o.folha) }
      if ($sh.FilterMode) { $sh.ShowAllData() }
      if ($sh.AutoFilterMode) { $sh.AutoFilter.Range.EntireRow.Hidden = $false }
    } catch {}
    [void]$res.Add($null); continue
  }
  if ($o.op -eq 'aplicar_filtros') {
    $so = $null
    if ($o.folha) { try { if ($o.folha -is [string]) { $so = $wb.Worksheets.Item($o.folha).Name } else { $so = $wb.Worksheets.Item([int]$o.folha).Name } } catch {} }
    $semLinhas = New-Object System.Collections.ArrayList
    foreach ($it in $o.lista) {
      if ($so -and $it.folha -ne $so) { continue }
      try {
        $sh = $wb.Worksheets.Item($it.folha)
        if ($sh.AutoFilterMode) { $rg = $sh.AutoFilter.Range }
        else { $rg = $sh.Range($sh.Cells.Item([int]$it.r, [int]$it.c), $sh.Cells.Item([int]$it.r + [int]$it.nr - 1, [int]$it.c + [int]$it.nc - 1)) }
        $cabs = @(); foreach ($c in $rg.Rows.Item(1).Cells) { $cabs += ([string]$c.Value2).Trim() }
        foreach ($f in $it.filtros) {
          $k = [array]::IndexOf($cabs, [string]$f.cab)
          if ($k -lt 0) { continue }
          $c1 = $f.c1; if ($c1 -is [array]) { $c1 = [object[]]$c1 }
          try {
            if ($f.op -eq 0) { [void]$rg.AutoFilter($k + 1, $c1) }
            elseif (($f.op -eq 1 -or $f.op -eq 2) -and $f.c2 -ne $null) { [void]$rg.AutoFilter($k + 1, $c1, $f.op, $f.c2) }
            else { [void]$rg.AutoFilter($k + 1, $c1, $f.op) }
          } catch {}
        }
        $vis = 0; try { $vis = $sh.AutoFilter.Range.Columns.Item(1).SpecialCells(12).Count } catch {}
        if ($vis -eq 1) { [void]$semLinhas.Add($it.folha) }
      } catch {}
    }
    [void]$res.Add($semLinhas); continue
  }
  if ($o.folha -is [string]) { $ws = $wb.Worksheets.Item($o.folha) } else { $ws = $wb.Worksheets.Item([int]$o.folha) }
  $rg = $ws.Range($o.rng)
  if ($o.op -eq 'ler') {
    $v = $rg.Value2
    $linhas = New-Object System.Collections.ArrayList
    if ($v -is [array]) {
      for ($i = 1; $i -le $v.GetLength(0); $i++) {
        $l = New-Object System.Collections.ArrayList
        for ($j = 1; $j -le $v.GetLength(1); $j++) { $e = $v.GetValue($i, $j); [void]$l.Add($e) }
        [void]$linhas.Add($l)
      }
    } else { $l = New-Object System.Collections.ArrayList; [void]$l.Add($v); [void]$linhas.Add($l) }
    [void]$res.Add($linhas)
  } else {
    $nr = $o.valores.Count; $nc = $o.valores[0].Count
    $arr = New-Object 'object[,]' $nr, $nc
    for ($i = 0; $i -lt $nr; $i++) { for ($j = 0; $j -lt $nc; $j++) { $arr.SetValue($o.valores[$i][$j], $i, $j) } }
    $rg.Value2 = $arr
    [void]$res.Add($null)
  }
}
if ($req.guardar) { try { $wb.Save() } catch {} }
[Console]::OutputEncoding = [Text.Encoding]::UTF8
ConvertTo-Json -InputObject $res -Depth 6 -Compress
"""


VISTA_PROTEGIDA = ("\nO {} está aberto no Excel em VISTA PROTEGIDA (ficheiro vindo da internet ou de um zip), e assim\n"
                   "nenhum programa o consegue ler. No Excel, carrega em 'Ativar Edição' (barra amarela no topo) e volta a carregar no botão.")


def _como_linhas(v):
    if isinstance(v, (tuple, list)):
        return [list(r) if isinstance(r, (tuple, list)) else [r] for r in v]
    return [[v]]


def _ops_pywin32(livro, ops, guardar):
    import win32com.client
    wb = None
    try:                                     # todos os Excel abertos (tabela de objetos em execução)
        import pythoncom
        rot, ctx = pythoncom.GetRunningObjectTable(), pythoncom.CreateBindCtx(0)
        for mk in rot:
            nome = mk.GetDisplayName(ctx, None) or ""
            if re.search(r"[\\/]" + re.escape(livro) + "$", nome, re.I):
                wb = win32com.client.Dispatch(rot.GetObject(mk).QueryInterface(pythoncom.IID_IDispatch))
                break
    except Exception:
        wb = None
    if wb is None:
        xl = win32com.client.GetActiveObject("Excel.Application")
        wb = next((b for b in xl.Workbooks if b.Name.lower() == livro.lower()), None)
        if wb is None:
            try:
                if any(str(pv.SourceName).lower() == livro.lower() for pv in xl.ProtectedViewWindows):
                    raise SystemExit(VISTA_PROTEGIDA.format(livro))
            except SystemExit:
                raise
            except Exception:
                pass
            return None
    res = []
    for o in ops:
        if o["op"] == "autosave":            # Salvamento Automático (OneDrive) desligado / ligado
            try:
                wb.AutoSaveOn = bool(o["valor"])
            except Exception:
                pass
            res.append(None)
            continue
        if o["op"] == "ativa":
            res.append(int(wb.ActiveSheet.Index))
            continue
        if o["op"] == "fechar":
            try:
                if not wb.Saved:                 # guarda o que estiver por guardar, para a cópia local e a da nuvem
                    wb.Save()                    # ficarem iguais antes de o script escrever o ficheiro novo
            except Exception:
                pass
            wb.Close(SaveChanges=False)
            return res + [None]
        if o["op"] in ("quadro", "fechar_quadro"):   # quadro flutuante da carteira (caixa de texto + botão ✕)
            for sh in wb.Worksheets:
                for nm in [x.Name for x in sh.Shapes if x.Name.startswith("Quadro")]:
                    sh.Shapes(nm).Delete()
            if o["op"] == "quadro":
                for fi in o["folhas"]:
                    sh = wb.Worksheets(fi)
                    rg = sh.Range(o["ancora"])
                    tb = sh.Shapes.AddTextbox(1, rg.Left, rg.Top, o["largura"], o["altura"])
                    tb.Name = "QuadroCarteira"
                    tb.TextFrame2.TextRange.Text = o["texto"]
                    tb.TextFrame2.TextRange.Font.Size = 10
                    tb.TextFrame2.TextRange.Font.Bold = True
                    tb.Fill.ForeColor.RGB = 13431551          # amarelo claro
                    tb.Line.ForeColor.RGB = 8421504
                    tb.Placement = 3
                    x = sh.Shapes.AddShape(9, rg.Left + o["largura"] - 20, rg.Top + 3, 16, 16)
                    x.Name = "QuadroFechar"
                    x.TextFrame2.TextRange.Text = "x"
                    x.TextFrame2.TextRange.Font.Size = 9
                    x.Fill.ForeColor.RGB = 255                # vermelho
                    x.Placement = 3
                    sh.Hyperlinks.Add(x, f"{PROTOCOLO}://fechar_quadro")
            res.append(None)
            continue
        if o["op"] == "ocultos":                    # tickers das linhas escondidas (pelo filtro) na folha
            sh, oc = wb.Worksheets(o["folha"]), []
            for r in range(L0 + 1, L0 + 1 + MAX_LINHAS):
                v = sh.Cells(r, o["col"]).Value
                if not v:
                    break
                if sh.Rows(r).Hidden:
                    oc.append(str(v))
            res.append(oc)
            continue
        if o["op"] == "ler_filtros":                # filtros de todas as folhas (por cabeçalho), para os repor
            res.append(_ler_filtros_com(wb))
            continue
        if o["op"] == "aplicar_filtros":
            nome = wb.Worksheets(o["folha"]).Name if o.get("folha") else None
            res.append(_aplicar_filtros_com(wb, [x for x in o["lista"] if nome is None or x["folha"] == nome]))
            continue
        if o["op"] == "copia":                      # cópia de segurança do livro (como está agora no Excel)
            try:
                wb.SaveCopyAs(o["destino"])
            except Exception:
                pass
            res.append(None)
            continue
        if o["op"] == "reaplicar_filtro":           # depois de escrever: o filtro que lá está volta a ser aplicado
            try:
                sh = wb.Worksheets(o["folha"])
                if sh.AutoFilterMode and sh.FilterMode:
                    sh.AutoFilter.ApplyFilter()
            except Exception:
                pass
            res.append(None)
            continue
        if o["op"] == "mostrar_tudo":               # antes de escrever: tira o filtro (as linhas ficam todas visíveis)
            try:
                sh = wb.Worksheets(o["folha"])
                if sh.FilterMode:
                    sh.ShowAllData()
                if sh.AutoFilterMode:
                    sh.AutoFilter.Range.EntireRow.Hidden = False
            except Exception:
                pass
            res.append(None)
            continue
        rg = wb.Worksheets(o["folha"]).Range(o["rng"])
        if o["op"] == "ler":
            res.append(_como_linhas(rg.Value2))
        else:
            rg.Value2 = [list(r) for r in o["valores"]]
            res.append(None)
    if guardar:
        try:
            wb.Save()
        except Exception:
            pass
    return res


def _lista_com(v):
    return [x for x in v] if isinstance(v, (tuple, list)) else v


def _ler_filtros_com(wb):
    """[{folha, rng, filtros: [{cab, op, c1, c2}]}] dos filtros ativos (sem os de cor/ícone)."""
    out = []
    for sh in wb.Worksheets:
        try:
            if not sh.AutoFilterMode:
                continue
            af = sh.AutoFilter
            rg = af.Range
            fl = []
            for i in range(1, af.Filters.Count + 1):
                f = af.Filters(i)
                if not f.On:
                    continue
                op = int(f.Operator or 0)
                if op in (8, 9, 10):
                    continue
                try:
                    c2 = _lista_com(f.Criteria2)
                except Exception:
                    c2 = None
                fl.append({"cab": txt(sh.Cells(rg.Row, rg.Column + i - 1).Value), "op": op,
                           "c1": _lista_com(f.Criteria1), "c2": c2})
            if fl:
                out.append({"folha": sh.Name, "r": rg.Row, "c": rg.Column, "nr": rg.Rows.Count, "nc": rg.Columns.Count,
                            "filtros": fl})
        except Exception:
            pass
    return out


def _aplicar_filtros_com(wb, lista):
    """Repõe os filtros (por cabeçalho). Se uma folha ficar sem nenhuma linha visível, tira o filtro dessa folha
    (nunca ficam as linhas todas escondidas). Devolve os nomes das folhas em que isso aconteceu."""
    sem_linhas = []
    for it in lista or []:
        try:
            sh = wb.Worksheets(it["folha"])
            rg = sh.AutoFilter.Range if sh.AutoFilterMode else sh.Range(
                sh.Cells(it["r"], it["c"]), sh.Cells(it["r"] + it["nr"] - 1, it["c"] + it["nc"] - 1))
            cabs = [txt(c.Value) for c in rg.Rows(1).Cells]
            for f in it["filtros"]:
                if f["cab"] not in cabs:
                    continue
                campo = cabs.index(f["cab"]) + 1
                try:
                    if not f["op"]:
                        rg.AutoFilter(campo, f["c1"])
                    elif f["op"] in (1, 2) and f.get("c2") is not None:
                        rg.AutoFilter(campo, f["c1"], f["op"], f["c2"])
                    else:
                        rg.AutoFilter(campo, f["c1"], f["op"])
                except Exception:
                    pass
            try:
                visiveis = sh.AutoFilter.Range.Columns(1).SpecialCells(12).Count    # 12 = só células visíveis
            except Exception:
                visiveis = 0
            if visiveis == 1:                               # só o cabeçalho visível: fica, mas avisa
                sem_linhas.append(it["folha"])
        except Exception:
            pass
    return sem_linhas


def com_filtros(ex, ops, folhas):
    """Escreve sem estragar filtros: guarda uma cópia do livro, lê os filtros, tira-os das folhas onde se vai
    escrever (NUNCA se escreve numa folha filtrada: o Excel só escreveria nas linhas visíveis), escreve e volta a
    aplicá-los (por cabeçalho). O filtro nunca é retirado, mesmo que esconda todas as linhas."""
    try:
        ex.ops([{"op": "copia", "destino": FICH_COPIA}])
    except Exception:
        pass
    try:
        filtros = ex.ops([{"op": "ler_filtros"}])[0] or []
    except Exception:
        filtros = []
    antes = [{"op": "mostrar_tudo", "folha": f} for f in folhas]
    depois = [{"op": "aplicar_filtros", "folha": f, "lista": filtros} for f in folhas] if filtros else []
    r = ex.ops(antes + list(ops) + depois, guardar=True)
    return r[len(antes):len(antes) + len(ops)] if r else r


def ler_filtros():
    """Filtros do painel aberto no Excel (antes de o ▶ ATUALIZAR o recriar). [] se não estiver aberto."""
    if os.name != "nt" or not os.path.exists(FICHEIRO):
        return []
    livro = os.path.basename(FICHEIRO)
    for f in (_ops_pywin32, _ops_powershell):
        try:
            r = f(livro, [{"op": "ler_filtros"}], False)
            if r is not None:
                return r[0] or []
        except SystemExit:
            return []
        except Exception:
            continue
    return []


def op_ocultos(idx):
    return {"op": "ocultos", "folha": idx, "col": COL["ticker"], "r0": L0 + 1, "max": MAX_LINHAS}


OCULTOS = set()       # tickers escondidos pelos filtros do Excel (não aparecem na página do telemóvel)


def repor_filtros(filtros, quadro=None, espera=45):
    """Depois de abrir o painel novo: volta a aplicar os filtros e mostra o quadro da carteira (espera que o Excel
    abra o livro)."""
    if (not filtros and not quadro) or os.name != "nt":
        return
    livro = os.path.basename(FICHEIRO)
    fim = time.time() + espera
    while time.time() < fim:
        time.sleep(3)
        for f in (_ops_pywin32, _ops_powershell):
            try:
                # sem gravar: o Salvamento Automático do Excel trata disso (gravar daqui ao mesmo tempo que o
                # OneDrive envia o ficheiro dava "conflito de intercalação" / "outra cópia aberta")
                r = f(livro, [{"op": "aplicar_filtros", "lista": filtros or []}, op_ocultos(1), op_ocultos(2)], False)
                if r is not None and quadro:
                    try:
                        f(livro, [{"op": "quadro", "folhas": [1, 2], "texto": quadro,
                                   "ancora": f"{get_column_letter(COL['nome'] + 2)}{L0 + 1}", "largura": 520, "altura": 66}],
                          True)
                        print("Quadro da carteira mostrado (fecha-se no x vermelho).")
                    except Exception as e:
                        print("Aviso: não consegui mostrar o quadro da carteira:", e)
                    quadro = None
                if r is not None:
                    OCULTOS.clear()
                    OCULTOS.update(txt(t).upper() for t in (r[1] or []) + (r[2] or []))
                    if not filtros:
                        return
                    n = sum(len(x["filtros"]) for x in filtros)
                    print(f"Filtros repostos: {n} (em {len(filtros)} folha(s)).")
                    if r[0]:
                        print("Aviso: o filtro de " + ", ".join(r[0]) + " escondia todas as linhas; foi retirado.")
                    return
            except SystemExit:
                return
            except Exception:
                continue
    print("Aviso: não consegui repor os filtros (o Excel demorou a abrir). Volta a aplicá-los à mão.")


SEM_JANELA = {"creationflags": 0x08000000} if os.name == "nt" else {}   # CREATE_NO_WINDOW: sem janela a piscar


def _ops_powershell(livro, ops, guardar):
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
        json.dump({"livro": livro, "ops": ops, "guardar": guardar}, f, ensure_ascii=False)
        fin = f.name
    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8-sig") as f:
        f.write(PS.replace("@@IN@@", fin.replace("'", "''")))
        fps = f.name
    try:
        exe = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), r"System32\WindowsPowerShell\v1.0\powershell.exe")
        if not os.path.exists(exe):
            exe = "powershell"
        r = subprocess.run([exe, "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-ExecutionPolicy", "Bypass",
                            "-File", fps], capture_output=True, timeout=180, **SEM_JANELA)
        if r.returncode == 3:
            return None
        if r.returncode == 4:
            raise SystemExit(VISTA_PROTEGIDA.format(livro))
        if r.returncode != 0:
            raise RuntimeError("PowerShell: " + r.stderr.decode("utf-8", "replace").strip()[:800])
        out = r.stdout.decode("utf-8", "replace").strip()
        return json.loads(out.splitlines()[-1]) if out else []
    finally:
        for p in (fin, fps):
            try:
                os.remove(p)
            except Exception:
                pass


def _ops_ficheiro(caminho, ops, guardar):
    from openpyxl import load_workbook
    wb = load_workbook(caminho)
    def folha(f):                                 # V179: 1 / 2 = Ações(2) / ETFs(2) se existirem (senão, a posição)
        if isinstance(f, int) and NOME_IDX.get(f) in wb.sheetnames:
            return wb[NOME_IDX[f]]
        return wb.worksheets[f - 1] if isinstance(f, int) else wb[f]
    res = []
    for o in ops:
        if o["op"] == "fechar":
            res.append(None)
            continue
        if o["op"] == "ativa":
            res.append(wb.index(wb.active) + 1)
        elif o["op"] == "ler_filtros":
            res.append([])
        elif o["op"] in ("quadro", "fechar_quadro", "autosave"):
            res.append(None)
        elif o["op"] == "ocultos":
            ws, oc = folha(o["folha"]), []
            for r in range(L0 + 1, L0 + 1 + MAX_LINHAS):
                v = ws.cell(r, o["col"]).value
                if not v:
                    break
                if ws.row_dimensions[r].hidden:
                    oc.append(str(v))
            res.append(oc)
        elif o["op"] == "aplicar_filtros":
            res.append([])
        elif o["op"] == "copia":
            try:
                import shutil
                shutil.copy2(caminho, o["destino"])
            except Exception:
                pass
            res.append(None)
        elif o["op"] == "reaplicar_filtro":
            ws = folha(o["folha"])
            if ws.auto_filter.filterColumn:             # sem Excel não dá para avaliar o filtro: mostra as linhas
                for r in range(L0 + 1, ws.max_row + 1):   # (o filtro fica; no Excel: Dados > Reaplicar)
                    ws.row_dimensions[r].hidden = False
            res.append(None)
        elif o["op"] == "mostrar_tudo":
            ws = folha(o["folha"])
            if ws.auto_filter.filterColumn:
                ws.auto_filter.filterColumn = []
                for r in range(L0 + 1, ws.max_row + 1):
                    ws.row_dimensions[r].hidden = False
            res.append(None)
        elif o["op"] == "ler":
            v = folha(o["folha"])[o["rng"]]
            res.append([[c.value for c in r] for r in v] if isinstance(v, tuple) else [[v.value]])
        else:
            ws = folha(o["folha"])
            cel = ws[o["rng"]]
            cel = cel if isinstance(cel, tuple) else ((cel,),)
            for r, linha in zip(cel, o["valores"]):
                for c, val in zip(r, linha):
                    c.value = val
            res.append(None)
    if guardar or any(o["op"] == "escrever" for o in ops):
        wb.save(caminho)
    return res


class Excel:
    """Lê/escreve num livro: no Excel aberto (pywin32 ou, sem ele, PowerShell) ou, se estiver fechado, no ficheiro."""

    def __init__(self, caminho):
        self.caminho, self.livro, self.modo = caminho, os.path.basename(caminho), None

    def ops(self, ops, guardar=False):
        if any(isinstance(o.get("folha"), int) and not isinstance(o.get("folha"), bool) and o.get("folha") in (1, 2)
               for o in ops):
            if not hasattr(self, "_novo2"):           # V179: livro já com «Ações(2)» / «ETFs(2)»? (experimenta uma vez)
                self._novo2 = False
                modo0 = self.modo
                try:
                    self.ops([{"op": "ler", "folha": "§" + NOME_IDX[1], "rng": "A1"}])
                    self._novo2 = True
                except BaseException:
                    pass
                finally:
                    self.modo = modo0
            if self._novo2:
                ops = [{**o, "folha": NOME_IDX[o["folha"]]} if isinstance(o.get("folha"), int) and o["folha"] in NOME_IDX
                       else o for o in ops]
        ops = [{**o, "folha": o["folha"][1:]} if isinstance(o.get("folha"), str) and o["folha"].startswith("§") else o
               for o in ops]                      # (§ = nome sem tradução, só para a experiência acima)
        ops = [{**o, "folha": NOME_PROC[o["folha"]]} if isinstance(o.get("folha"), int) and o.get("folha") in NOME_PROC
               else o for o in ops]               # Proc ações/ETFs pelo nome (a folha Compra vem antes delas)
        if any(o.get("folha") in NOME_PROC.values() for o in ops) and not getattr(self, "_proc_ok", False):
            self._proc_ok = True                  # livro ainda com os nomes antigos (até à V122)? experimenta uma vez
            modo0 = self.modo
            try:
                self.ops([{"op": "ler", "folha": NOME_PROC[FOLHA_PROC], "rng": "A1"}])
            except BaseException:                 # (a falha chega como SystemExit: o ficheiro está aberto no Excel)
                self._proc_ant = True
            finally:
                self.modo = modo0
        if getattr(self, "_proc_ant", False):
            inv = {v: NOME_PROC_ANT[k] for k, v in NOME_PROC.items()}
            ops = [{**o, "folha": inv.get(o.get("folha"), o.get("folha"))} for o in ops]
        if self.modo in (None, "pywin32"):
            try:
                r = _ops_pywin32(self.livro, ops, guardar)
                if r is not None:
                    self.modo = "pywin32"
                    return r
            except SystemExit:
                raise
            except Exception:
                pass
        if os.name == "nt" and self.modo in (None, "ps"):
            r = _ops_powershell(self.livro, ops, guardar)
            if r is not None:
                self.modo = "ps"
                return r
        try:
            self.modo = "ficheiro"
            return _ops_ficheiro(self.caminho, ops, guardar)
        except PermissionError:
            raise SystemExit(f"\nNão consegui ler o {self.livro}: está aberto no Excel e não consegui ligar-me a ele.\n"
                             "Causas habituais:\n"
                             "  - está em Vista Protegida: no Excel carrega em 'Ativar Edição' (barra amarela) e tenta outra vez;\n"
                             "  - o Excel foi aberto como administrador: fecha-o e abre-o normalmente;\n"
                             f"  - o ficheiro aberto não é o desta pasta ({PASTA}).\n"
                             "Se nada disto resolver: guarda-o, fecha-o e volta a carregar no botão (ou duplo clique no .bat).")

    def estado(self, folha, texto, guardar=False, rapido=False):
        """Escreve na célula do estado (com o PowerShell, as mensagens 'rapido' saltam-se: cada chamada demora ~1 s)."""
        if rapido and self.modo != "pywin32":
            return
        try:
            self.ops([{"op": "escrever", "folha": folha, "rng": C_ESTADO, "valores": [[texto]]}], guardar)
        except BaseException:
            pass


def vazio_esp(v):
    """Célula vazia -> um espaço (assim o texto da célula ao lado não se estende por cima dela)."""
    return " " if v is None or (isinstance(v, str) and v == "") else v


def sem_esp(v):
    """Ao ler do Excel: o espaço das células vazias volta a ser vazio (e o nome com ligação volta a ser só o nome)."""
    if isinstance(v, str) and v.startswith("=HYPERLINK("):
        m = re.search(r',"((?:[^"]|"")*)"\)$', v)
        return m.group(1).replace('""', '"') if m else v
    return None if isinstance(v, str) and not v.strip() else v


URL_GRAFICO = "https://finance.yahoo.com/quote/"


def link_nome(nome, r):
    """Nome do título com ligação para o gráfico (Yahoo, no ticker da folha = na moeda da folha). A fórmula usa o
    ticker da própria linha, por isso continua certa depois de ordenar."""
    if not isinstance(nome, str) or not nome.strip() or nome.startswith("="):
        return nome
    tl = get_column_letter(COL["ticker"])
    return f'=HYPERLINK("{URL_GRAFICO}"&${tl}{r}&"/","{nome[:250].replace(chr(34), chr(34) * 2)}")'


def com_links(valores, i0):
    """Linhas a escrever a partir da linha i0: põe a ligação no nome."""
    j = COL["nome"] - 1
    for i, v in enumerate(valores):
        v[j] = link_nome(v[j], i0 + i)
    return valores


def linha_lida(chaves, v):
    return {k: sem_esp(x) for k, x in zip(chaves, v)}


def linhas_por_cabecalho(tab, idx):
    """Linhas da tabela lida (folha idx), pelas colunas desta versão; se a folha ainda for de uma versão anterior
    (ex.: sem a coluna Beta), lê cada coluna pelo nome do cabeçalho. Só para ler (ver posição)."""
    chaves = [k for k, *_ in COLS]
    cab = [txt(v) for v in tab[0]]
    if cab_ok(cab):
        return [d for d in (linha_lida(chaves, v) for v in tab[1:]) if txt(d.get("ticker"))]
    ref = CAB_RANK if idx == 1 else CAB_ETFS
    pos = {k: cab.index(h) for k, h in zip(chaves, ref) if h in cab}
    if "ticker" not in pos:
        raise ValueError("colunas")
    out = []
    for v in tab[1:]:
        d = {k: (sem_esp(v[j]) if j < len(v) else None) for k, j in pos.items()}
        if txt(d.get("ticker")):
            out.append(d)
    return out


def qtd_detida(v):
    """Coluna 'Detido': x (tens, sem dizer quantas) ou o nº de ações (pode ter decimais). Devolve o nº ou None."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v) if v > 0 else None
    t = str(v or "").strip().replace(" ", "")
    if re.fullmatch(r"\d+([.,]\d+)?", t):
        f = float(t.replace(",", "."))
        return f if f > 0 else None
    return None


def e_detido(v):
    return str(v or "").strip().lower() == "x" or qtd_detida(v) is not None


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()



# ================================================================================================ ↕ ORDENAR
MAX_LINHAS = 400


def _ord_min_de(v):
    """Lê o mínimo de L3 (ao lado do «Ordenar por») e guarda-o para a próxima reconstrução da folha."""
    try:
        x = v[0][0] if isinstance(v, (list, tuple)) else v
        x = float(str(x).replace("≥", "").replace(",", ".").strip())
        ORD_MIN["v"] = x
        with open(FICH_ORD_MIN, "w", encoding="utf-8") as f:
            json.dump({"v": x}, f)
    except Exception:
        pass


def ordenar(idx):
    """Reordena a folha idx (1 = Ações, 2 = ETFs) pelo critério escolhido em D3, sem descarregar nada."""
    ex = Excel(FICHEIRO)
    ultima = get_column_letter(len(COLS))
    r = ex.ops([{"op": "ler", "folha": idx, "rng": C_ORD},
                {"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"},
                {"op": "ler", "folha": idx, "rng": C_ORD_MIN}])
    criterio = txt(r[0][0][0]) or ORD_PADRAO[idx]
    _ord_min_de(r[2])
    cab, dados = [txt(v) for v in r[1][0]], r[1][1:]
    chaves = {h: k for k, h, *_ in COLS}
    chaves.update({h: k for k, h in CAB_ETF.items()})
    if [chaves.get(h) for h in cab] != [k for k, *_ in COLS]:
        raise SystemExit("As colunas desta folha não são as desta versão do script: carrega primeiro em ▶ ATUALIZAR.")
    i_tk = COL["ticker"] - 1
    n = 0
    while n < len(dados) and txt(dados[n][i_tk]):
        n += 1
    linhas = [linha_lida([k for k, *_ in COLS], dados[i]) for i in range(n)]
    linhas.sort(key=chave_ordem(criterio))
    valores = com_links([[vazio_esp(d.get(k)) for k, *_ in COLS] for d in linhas], L0 + 1)
    com_filtros(ex, [{"op": "escrever", "folha": idx, "rng": f"A{L0 + 1}:{ultima}{L0 + n}", "valores": valores},
                     {"op": "escrever", "folha": idx, "rng": C_ORD_EST, "valores": [[
                         f"Ordenado por: {criterio} ({datetime.now():%d/%m/%Y %H:%M})"]]}], [idx])
    print(f"{n} linhas ordenadas por: {criterio}")
    registo(f"ordenar folha {idx}: {n} linhas por '{criterio}' (ligação ao Excel: {ex.modo})")
    time.sleep(2)


def ler_estado():
    """Antes de atualizar, lê do painel atual: o "Ordenar por" de cada folha, a folha Procurar e os Acrescentados."""
    criterios, acrescentados = {}, []
    estado = {"proc": {}, "det": set()}
    origem = FICHEIRO
    if not os.path.exists(origem):        # 1.ª vez com o nome novo: aproveita o painel com o nome antigo
        outro = os.path.join(PASTA, (BASE[1:] if BASE.startswith("_") else "_" + BASE) + ".xlsx")
        origem = outro if os.path.exists(outro) else None
    if origem:
        ex = Excel(origem)
        tentar = lambda ops: _tentar(ex, ops)
        r = tentar([{"op": "ler", "folha": 1, "rng": f"A{L0}"}, {"op": "ler", "folha": 1, "rng": f"A{L0_ANTIGO}"}])
        recente = bool(r) and (txt(r[0][0][0]) in ("Empresa / ETF", "Detido (x)", "Atualizar (x)")
                               or txt(r[1][0][0]) == "x")
        novo = bool(r) and txt(r[0][0][0]) == "Atualizar (x)"         # barra de topo da V012
        cel = (lambda k: {"ord": C_ORD, "lp": C_LPMIN, "ref": C_REF, "data": C_DATAREF, "at": C_AT}[k]) if novo \
            else (lambda k: CEL_ANTIGAS[k])
        estado["novo"] = novo
        if r and txt(r[0][0][0]) in ("Detido (x)", "Atualizar (x)"):   # títulos com x (detidos), por nome e por ticker
            ultima = get_column_letter(len(COLS))
            completo = True
            for idx in (1, 2):
                t = tentar([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"}])
                cab = [txt(v) for v in t[0][0]] if t else []
                if "Empresa / ETF" in cab and "Ticker Yahoo" in cab:
                    jn, jt = cab.index("Empresa / ETF"), cab.index("Ticker Yahoo")
                    jd = cab.index("Detido (x)") if "Detido (x)" in cab else 0
                    linhas_t = [v for v in t[0][1:] if txt(v[jt])]
                    estado.setdefault("universo", set()).update(
                        {txt(v[jn]) for v in linhas_t} | {txt(v[jt]).upper() for v in linhas_t})
                    if "Atualizar (x)" in cab:                # as linhas como estão (para a atualização seletiva)
                        por_cab = {h: k for k, h, *_ in COLS}
                        por_cab.update({CAB_ETF.get(k, h): k for k, h, *_ in COLS})
                        chaves = [por_cab.get(h, f"_{j}") for j, h in enumerate(cab)]
                        estado.setdefault("linhas_ant", {})[idx] = [linha_lida(chaves, v) for v in linhas_t]
                    jq = cab.index("Nº de ações detidas") if "Nº de ações detidas" in cab else None
                    jp = cab.index("Preço médio de compra (por ação)") if "Preço médio de compra (por ação)" in cab else None
                    jc = cab.index("Corretora") if "Corretora" in cab else None
                    for v in t[0][1:]:
                        if jc is not None and txt(v[jc]) and txt(v[jt]):
                            estado.setdefault("corr", {}).update({txt(v[jn]): txt(v[jc]), txt(v[jt]).upper(): txt(v[jc])})
                        qq = qtd_detida(v[jq]) if jq is not None else None
                        pp = qtd_detida(v[jp]) if jp is not None else None
                        if qq is not None:
                            estado.setdefault("qtd", {}).update({txt(v[jn]): qq, txt(v[jt]).upper(): qq})
                        if pp is not None:
                            estado.setdefault("pcompra", {}).update({txt(v[jn]): pp, txt(v[jt]).upper(): pp})
                        if qq is not None and not e_detido(v[jd]):
                            estado["det"] |= {txt(v[jn]), txt(v[jt]).upper()}
                        if e_detido(v[jd]):
                            estado["det"] |= {txt(v[jn]), txt(v[jt]).upper()}
                            q = qtd_detida(v[jd])
                            if q is not None:
                                estado.setdefault("qtd", {}).update({txt(v[jn]): q, txt(v[jt]).upper(): q})
                else:
                    completo = False                # uma folha não se leu: não se perde nenhum x guardado
                    estado["universo_incompleto"] = True
            if not completo:
                estado["det"] |= ler_detidos()
                estado["qtd"] = {**ler_qtd(), **estado.get("qtd", {})}
                estado["pcompra"] = {**ler_qtd("pcompra"), **estado.get("pcompra", {})}
                estado["corr"] = {**ler_qtd("corr"), **estado.get("corr", {})}
                print("Aviso: não consegui ler os x de uma das folhas; mantenho também os detidos guardados.")
            guardar_detidos(estado["det"], estado.get("qtd", {}), estado.get("pcompra", {}), estado.get("corr", {}))
            for idx, suf in ((1, ""), (2, "_e")):     # cada folha com a sua data de referência e a sua seleção
                rr = tentar([{"op": "ler", "folha": idx, "rng": cel("ref")}, {"op": "ler", "folha": idx, "rng": cel("data")},
                             {"op": "ler", "folha": idx, "rng": cel("lp")}]
                            + ([{"op": "ler", "folha": idx, "rng": C_SEL}, {"op": "ler", "folha": idx, "rng": C_SEL_V}]
                               if novo else []))
                if rr:
                    esc_r = txt(rr[0][0][0]) or "Atual"
                    esc_r = "Data (escrever ao lado)" if esc_r.startswith("Data (") else esc_r
                    estado["ref_escolha" + suf], estado["ref_data" + suf] = esc_r, rr[1][0][0]
                    if idx == 1 and rr[2][0][0] not in (None, ""):
                        estado["lp_min"] = ler_lp_min(rr[2][0][0])
                    if novo:
                        sel = ler_sel_valor(ler_sel(rr[3][0][0]), rr[4][0][0])
                        print(f"{'Ações' if idx == 1 else 'ETFs'}: lista 'o que atualizar' = {sel or txt(rr[3][0][0]) or '(vazia)'}"
                              f"; {sum(1 for d in estado.get('linhas_ant', {}).get(idx, []) if txt(d.get('sel')))} "
                              f"título(s) com x na coluna A.")
                        if sel:
                            estado["sel" + suf] = sel
        if L0 == L0_NOVO and novo:                   # V190: parâmetros da nota por ganhos (as 2 tabelas de cima)
            for idx, ch in ((1, "gp"), (2, "gp_e")):
                rr = tentar(g_ops_params(idx))
                if rr:
                    estado[ch] = g_vals_params(rr)
        nomes_crit = [c[0] for c in CRITERIOS_ORD]
        for idx in (1, 2):
            r = tentar([{"op": "ler", "folha": idx, "rng": (C_ORD if novo else CEL_ANTIGAS["ord"]) if recente else "D3"},
                        {"op": "ler", "folha": idx, "rng": f"{get_column_letter(COL['nota'])}{L0}"}])
            if r and ORD_ANTIGOS.get(txt(r[0][0][0]), txt(r[0][0][0])) in nomes_crit:
                criterios[idx] = ORD_ANTIGOS.get(txt(r[0][0][0]), txt(r[0][0][0]))
                # 1.ª vez com a nota geral: quem ordenava pela pontuação passa a ordenar pela nota geral
                if criterios[idx] == "Pontuação (5 → 1)" and not txt(r[1][0][0]).startswith(("Nota geral", "Nota global")):
                    criterios[idx] = ORD_PADRAO[idx]
        if recente:
            for pf in (FOLHA_PROC, FOLHA_PROC_ETF):
                cels = (P_TXT, P_CRIT, P_REG, P_N)
                r = tentar([{"op": "ler", "folha": pf, "rng": f"J{L_TAB}"}] +
                           [{"op": "ler", "folha": pf, "rng": c} for c in cels])
                if r and txt(r[0][0][0]) == CAB_P[-1]:
                    estado["proc"][pf] = {c: r[k + 1][0][0] for k, c in enumerate(cels)}
        try:
            acrescentados = ler_acrescentados(ex)
        except SystemExit:
            raise
        except Exception as e:                     # não conseguiu ler a folha: tenta o ficheiro gravado e a cópia de segurança
            registo(f"aviso: folha Acrescentados não lida pelo Excel ({e}); a tentar o ficheiro / a cópia")
            try:
                acrescentados = ler_acrescentados(Excel(origem), so_ficheiro=True)
            except Exception:
                acrescentados = ler_copia_acr()
    elif not acrescentados:
        acrescentados = ler_copia_acr()
    if acrescentados:
        print(f"Folha Acrescentados: {len(acrescentados)} título(s).")
    registo(f"acrescentados lidos: {len(acrescentados)}")
    guardar_copia_acr(acrescentados)
    return criterios, acrescentados, estado


FICH_ACR = os.path.join(PASTA, BASE + "_acrescentados.json")
FICH_DET = os.path.join(PASTA, BASE + "_detidos.json")


def guardar_detidos(det, qtd=None, pcompra=None, corr=None):
    try:
        if os.path.exists(FICH_DET):                 # cópia do anterior (para recuperar se for preciso)
            shutil.copyfile(FICH_DET, FICH_DET.replace(".json", "_anterior.json"))
        with open(FICH_DET, "w", encoding="utf-8") as f:
            json.dump({"det": sorted(x for x in det if x), "qtd": {k: v for k, v in (qtd or {}).items() if k},
                       "pcompra": {k: v for k, v in (pcompra or {}).items() if k},
                       "corr": {k: v for k, v in (corr or {}).items() if k}},
                      f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def ler_detidos():
    try:
        with open(FICH_DET, encoding="utf-8") as f:
            j = json.load(f)
        return set(j.get("det", []) if isinstance(j, dict) else j)
    except Exception:
        return set()


def ler_qtd(chave="qtd"):
    try:
        with open(FICH_DET, encoding="utf-8") as f:
            j = json.load(f)
        return dict(j.get(chave, {})) if isinstance(j, dict) else {}
    except Exception:
        return {}


def guardar_copia_acr(acrescentados):
    """Cópia de segurança da folha Acrescentados (para nunca se perderem se o Excel falhar na leitura)."""
    try:
        with open(FICH_ACR, "w", encoding="utf-8") as f:
            json.dump([list(a) for a in acrescentados], f, ensure_ascii=False, indent=1)
    except Exception:
        pass


def ler_copia_acr():
    try:
        with open(FICH_ACR, encoding="utf-8") as f:
            return [tuple(a) for a in json.load(f)]
    except Exception:
        return []


def _tentar(ex, ops):
    try:
        return ex.ops(ops)
    except SystemExit:
        raise
    except Exception:
        return None


# custos usados no teste do "Último ano: seguir o sinal"
TESTE_CUSTO = 0.001             # custo por operação (comissão + spread), 0,1%
TESTE_JURO = 0.02               # juro anual do dinheiro parado (fora do mercado)


# ================================================================================================ ↻ PREÇOS
def precos_agora(tickers, _repetir=False):
    """Preço mais recente de cada ticker (velas de 1 minuto com pré-abertura e pós-fecho, todos de uma vez) e o fecho do
    dia anterior. Devolve {ticker: (preço, "sessão dd/mm hh:mm", fecho de referência)}."""
    import logging
    import warnings
    import pandas as pd
    import yfinance as yf
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)
    warnings.filterwarnings("ignore")
    if not _repetir:
        print(f"A pedir ao Yahoo os preços de {len(tickers)} títulos...")
    kw = dict(group_by="ticker", threads=True, progress=False, auto_adjust=False)
    minuto = yf.download(tickers, period="2d", interval="1m", prepost=True, **kw)
    diario = yf.download(tickers, period="10d", interval="1d", **kw)

    def serie(df, t):
        try:
            x = df[t]["Close"] if isinstance(df.columns, pd.MultiIndex) else df["Close"]
            return x.dropna()
        except Exception:
            return pd.Series(dtype=float)

    out = {}
    for t in tickers:
        m = serie(minuto, t)
        if not len(m):
            continue
        ts, preco = pd.Timestamp(m.index[-1]), float(m.iloc[-1])
        if ts.tzinfo is None:
            ts = ts.tz_localize("UTC")
        zona = _zona_bolsa(t)                         # o Yahoo junta bolsas de fusos diferentes em UTC
        local = ts.tz_convert(zona) if zona else ts
        sessao = "sessão normal"
        if zona:                                      # EUA/Canadá: sessão normal 9h30-16h00 (hora local da bolsa)
            minutos = local.hour * 60 + local.minute
            sessao = "pré-abertura" if minutos < 570 else "pós-fecho" if minutos >= 960 else "sessão normal"
        hora = datetime.fromtimestamp(ts.timestamp())   # hora deste PC (Lisboa)
        d = serie(diario, t)
        dia = lambda ix: (pd.Timestamp(ix).tz_convert(zona) if (zona and pd.Timestamp(ix).tzinfo) else pd.Timestamp(ix)).date()
        # as velas de 1 minuto do Yahoo às vezes param antes do fim da sessão (p. ex. a TSM ficou nas 14:00 de NY):
        # se a vela diária tiver um dia mais recente, ou se for do mesmo dia e a sessão já tiver acabado sem que as
        # velas de 1 minuto lá cheguem, usa-se o fecho da vela diária
        if len(d):
            d_dia, d_fecho = dia(d.index[-1]), float(d.iloc[-1])
            agora = pd.Timestamp.now(tz=zona) if zona else None
            parou = (zona is not None and d_dia == local.date() and sessao == "sessão normal"
                     and (agora.date() > d_dia or agora.hour * 60 + agora.minute >= 965)
                     and local.hour * 60 + local.minute < 955)
            if d_dia > local.date() or parou:
                preco, sessao = d_fecho, "fecho"
                hora = None
                local = pd.Timestamp(d_dia).tz_localize(zona) if zona else pd.Timestamp(d_dia)
        anteriores = d[[dia(ix) < local.date() for ix in d.index]] if len(d) else d
        ref = float(anteriores.iloc[-1]) if len(anteriores) else None
        out[t] = (preco, f"{sessao} {local:%d/%m}" if hora is None else f"{sessao} {hora:%d/%m %H:%M}", ref)
    faltam = [t for t in tickers if t not in out]
    if faltam and not _repetir:                        # verificação: os que falharam no pedido em conjunto, um a um
        print(f"Verificação: {len(faltam)} sem preço; a pedir outra vez um a um...")
        for t in faltam:
            time.sleep(1)
            try:
                out.update(precos_agora([t], _repetir=True))
            except Exception:
                pass
    return out


# Horário de negociação de cada bolsa (hora local) -> mostra-se em hora de Lisboa (com a hora de verão de cada sítio)
HORARIOS = {"": ("America/New_York", "09:30", "16:00"), ".L": ("Europe/London", "08:00", "16:30"),
            ".IL": ("Europe/London", "08:00", "16:30"), ".IR": ("Europe/Dublin", "08:00", "16:30"),
            ".LS": ("Europe/Lisbon", "08:00", "16:30"), ".DE": ("Europe/Berlin", "09:00", "17:30"),
            ".F": ("Europe/Berlin", "08:00", "22:00"), ".MU": ("Europe/Berlin", "08:00", "22:00"),
            ".SG": ("Europe/Berlin", "08:00", "22:00"), ".BE": ("Europe/Berlin", "08:00", "22:00"),
            ".DU": ("Europe/Berlin", "08:00", "20:00"), ".HM": ("Europe/Berlin", "08:00", "22:00"),
            ".HA": ("Europe/Berlin", "08:00", "22:00"), ".PA": ("Europe/Paris", "09:00", "17:30"),
            ".AS": ("Europe/Amsterdam", "09:00", "17:30"), ".BR": ("Europe/Brussels", "09:00", "17:30"),
            ".MI": ("Europe/Rome", "09:00", "17:30"), ".MC": ("Europe/Madrid", "09:00", "17:30"),
            ".VI": ("Europe/Vienna", "09:00", "17:30"), ".SW": ("Europe/Zurich", "09:00", "17:30"),
            ".ST": ("Europe/Stockholm", "09:00", "17:30"), ".HE": ("Europe/Helsinki", "10:00", "18:30"),
            ".CO": ("Europe/Copenhagen", "09:00", "17:00"), ".OL": ("Europe/Oslo", "09:00", "16:20"),
            ".WA": ("Europe/Warsaw", "09:00", "17:00"), ".T": ("Asia/Tokyo", "09:00", "15:30"),
            ".HK": ("Asia/Hong_Kong", "09:30", "16:00"), ".TW": ("Asia/Taipei", "09:00", "13:30"),
            ".TWO": ("Asia/Taipei", "09:00", "13:30"), ".KS": ("Asia/Seoul", "09:00", "15:30"),
            ".KQ": ("Asia/Seoul", "09:00", "15:30"), ".SS": ("Asia/Shanghai", "09:30", "15:00"),
            ".SZ": ("Asia/Shanghai", "09:30", "15:00"), ".NS": ("Asia/Kolkata", "09:15", "15:30"),
            ".BO": ("Asia/Kolkata", "09:15", "15:30"), ".SI": ("Asia/Singapore", "09:00", "17:00"),
            ".AX": ("Australia/Sydney", "10:00", "16:00"), ".TO": ("America/Toronto", "09:30", "16:00"),
            ".V": ("America/Toronto", "09:30", "16:00"), ".SA": ("America/Sao_Paulo", "10:00", "17:00"),
            ".BA": ("America/Argentina/Buenos_Aires", "11:00", "17:00"), ".MX": ("America/Mexico_City", "08:30", "15:00"),
            ".JO": ("Africa/Johannesburg", "09:00", "17:00"), ".TA": ("Asia/Jerusalem", "09:59", "17:25")}


def horario_bolsa(ticker):
    """'hh:mm-HH:mm' em hora de Lisboa (hoje), pela bolsa do ticker; vazio se não se souber."""
    t = txt(ticker).upper()
    suf = "" if "." not in t else "." + t.rsplit(".", 1)[1]
    h = HORARIOS.get(suf)
    if not h:
        return None
    try:
        import pandas as pd
        zona, a, f = h
        hoje = datetime.now().strftime("%Y-%m-%d")
        conv = lambda hm: pd.Timestamp(f"{hoje} {hm}").tz_localize(zona).tz_convert("Europe/Lisbon").strftime("%H:%M")
        return f"{conv(a)}-{conv(f)}"
    except Exception:
        return None


def _zona_bolsa(ticker):
    """Fuso da bolsa para separar pré-abertura / sessão normal / pós-fecho (só EUA e Canadá têm essas sessões no Yahoo)."""
    if "." not in ticker:
        return "America/New_York"
    if ticker.upper().endswith((".TO", ".V", ".NE", ".CN")):
        return "America/Toronto"
    return None


def recalcular(d, preco, hora, ref):
    """Atualiza numa linha (dict) o preço e tudo o que depende dele, com os valores que já estão na folha.
    V165: também acerta, de forma aproximada, o que o preço de hoje mexe nas médias (50/100/200 dias: o dia de hoje
    passa do preço anterior para o novo) e nos máximos / mínimos (um preço novo acima do máximo passa a ser o máximo)."""
    n = lambda k: d.get(k) if isinstance(d.get(k), (int, float)) else None
    antigo = n("preco")
    if antigo and preco and antigo > 0:
        for k, dias in (("mm50", 50), ("mm100c", 100), ("mm200c", 200), ("mm200", 200)):
            if n(k):
                d[k] = n(k) + (preco - antigo) / dias
        for i in range(len(PERIODOS)):
            mx, mn, med = n(f"max{i}"), n(f"min{i}"), n(f"med{i}")
            if mx and preco > mx:
                d[f"max{i}"] = mx = preco
            if mn and preco < mn:
                d[f"min{i}"] = mn = preco
            if med and mx and mn:
                d[f"cima{i}"], d[f"baixo{i}"] = mx / med - 1, mn / med - 1
        if n("max1a") and preco > n("max1a"):
            d["max1a"] = preco
    d["preco"], d["preco_hora"] = preco, hora
    if n("div"):                                     # o dividendo fica; o yield acompanha o preço
        d["div_yield"] = n("div") / preco
    if ref:
        d["var1d"] = preco / ref - 1
    if n("alvo"):
        d["pot"] = n("alvo") / preco - 1
    if n("justo"):
        d["justo_pct"] = n("justo") / preco - 1
    for i in range(len(PERIODOS)):
        if n(f"max{i}"):
            d[f"dist{i}"] = preco / n(f"max{i}") - 1
    d["stop"] = n("venda_min")                         # a rede de segurança (máximo de 6 meses − 4σ) não muda com o preço
    niveis(d, preco)
    if n("mm200"):
        d["dist200"] = preco / n("mm200") - 1
    pontuar(d, preco)
    preencher_alt(d)
    carteira(d)

def atualizar_precos(reclassificar=False):
    """reclassificar=True (botão da folha Ponderações): só volta a calcular notas e decisões com as percentagens novas
    e reordena, sem descarregar nada.
    ↻ PREÇOS: preços de agora (com pré-abertura e pós-fecho) e as contas que dependem deles, nas folhas Ações e ETFs,
    sem descarregar o resto; depois volta a ordenar cada folha pelo critério escolhido (os x acompanham as linhas)."""
    ex = Excel(FICHEIRO)
    ultima = get_column_letter(len(COLS))
    global LP_MIN
    rr = ex.ops([{"op": "ler", "folha": 1, "rng": C_REF}, {"op": "ler", "folha": 1, "rng": C_LPMIN},
                 {"op": "ler", "folha": 2, "rng": C_REF}])
    if rr[1][0][0] not in (None, ""):
        LP_MIN = ler_lp_min(rr[1][0][0])
    carregar_pesos(ex)
    passada = {idx: (txt(rr[k][0][0]) not in ("", "Atual") and (txt(rr[k][0][0]) in REFS or txt(rr[k][0][0]).startswith("Data ("))) for idx, k in ((1, 0), (2, 2))}
    if all(passada.values()):
        raise SystemExit(f"As duas folhas estão na situação de uma data passada. Para preços de agora escolhe 'Atual' "
                         f"em {C_REF} e carrega em ▶ ATUALIZAR.")
    for idx in (1, 2):
        if passada[idx]:
            print(f"A folha {'Ações' if idx == 1 else 'ETFs'} está numa data passada: os preços dela ficam como estão.")
    try:
        _ord_min_de(ex.ops([{"op": "ler", "folha": 1, "rng": C_ORD_MIN}])[0])
    except Exception:
        pass
    r = ex.ops([op for idx in (1, 2) for op in (
        {"op": "ler", "folha": idx, "rng": C_ORD},
        {"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"})])
    chaves = [k for k, *_ in COLS]
    folhas = {}
    for idx in (1, 2):
        criterio, tabela = txt(r[2 * idx - 2][0][0]) or ORD_PADRAO[idx], r[2 * idx - 1]
        cab = [txt(v) for v in tabela[0]]
        if not cab_ok(cab):
            raise SystemExit("As colunas desta folha não são as desta versão do script: carrega primeiro em ▶ ATUALIZAR.")
        linhas = []
        for v in tabela[1:]:
            d = linha_lida(chaves, v)
            if not txt(d.get("ticker")):
                break
            linhas.append(d)
        folhas[idx] = (criterio, linhas)
    tickers = sorted({txt(d["ticker"]) for _, ls in folhas.values() for d in ls})
    precos = {} if reclassificar else precos_agora(tickers)
    ops, n_ok = [], 0
    if reclassificar:                           # notas das ações primeiro (entram na nota ponderada dos ETFs)
        for d in folhas[1][1]:
            if isinstance(d.get("preco"), (int, float)):
                pontuar(d, d["preco"])
        pontuar_etfs(folhas[2][1])
        ponderar_etfs(folhas[1][1], folhas[2][1], _ler_json(FICH_COMP, {}))
    for idx, (criterio, linhas) in folhas.items():
        for d in linhas:
            p = precos.get(txt(d["ticker"]))
            if p and not passada[idx]:
                recalcular(d, *p)
                n_ok += 1
    if not reclassificar:                         # V165: ETFs depois das ações: sinal (comparação entre ETFs) e
        pontuar_etfs(folhas[2][1])                # nota ponderada das suas ações, com as notas novas das ações
        if not passada[1] and not passada[2]:
            try:
                ponderar_etfs(folhas[1][1], folhas[2][1], _ler_json(FICH_COMP, {}))
            except Exception as e:
                print("   aviso: nota das ações dos ETFs:", e)
    for idx, (criterio, linhas) in folhas.items():
        linhas.sort(key=chave_ordem(criterio))
        if linhas:
            ops.append({"op": "escrever", "folha": idx, "rng": f"A{L0 + 1}:{ultima}{L0 + len(linhas)}",
                        "valores": com_links([[vazio_esp(d.get(k)) for k in chaves] for d in linhas], L0 + 1)})
    sem = len(tickers) - len(precos)
    msg = (f"Reclassificado e reordenado em {datetime.now():%d/%m/%Y %H:%M} com as percentagens da folha Ponderações "
           f"(ações: {texto_pesos(True).split(' (')[0]}). Preços e dados do último ▶ ATUALIZAR / ↻ PREÇOS."
           if reclassificar else None) or (f"Preços atualizados em {datetime.now():%d/%m/%Y %H:%M} ({len(precos)} títulos"
           + (f"; {sem} sem preço novo" if sem else "") + "). O resto é do último ▶ ATUALIZAR.")
    try:
        if calibrar_geral(folhas[1][1]):
            print("Nota global: " + calib_txt().split(" Só informa")[0])
            ops.append({"op": "escrever", "folha": FOLHA_PESOS, "rng": "C64", "valores": [[calib_txt()]]})
    except Exception as e:
        print("   aviso: calibração da nota global:", e)
    for idx in (1, 2):
        ops.append({"op": "escrever", "folha": idx, "rng": C_ORD_EST if reclassificar else C_PRECOS_EST, "valores": [[msg]]})
        ops.append({"op": "escrever", "folha": idx, "rng": "A1", "valores": [[TXT_VERSAO]]})
    com_filtros(ex, ops, [1, 2])                # sem filtro durante a escrita; depois os teus filtros voltam
    try:
        r = ex.ops([op_ocultos(1), op_ocultos(2)])
        OCULTOS.clear()
        OCULTOS.update(txt(t).upper() for t in (r[0] or []) + (r[1] or []))
    except Exception:
        pass
    print(msg)
    sync_dos_precos(folhas[1][1], folhas[2][1])
    gerar_html(folhas[1][1], folhas[2][1], msg)
    registo(("reclassificar" if reclassificar else f"preços: {len(precos)} de {len(tickers)} títulos")
            + f" (ligação ao Excel: {ex.modo})")
    time.sleep(2)


# ================================================================================================ telemóvel
# Em cada ▶ ATUALIZAR e ↻ PREÇOS grava-se _Painel_telemovel.html na mesma pasta (OneDrive): uma página para o telemóvel
# (abre-se pela app OneDrive -> abrir no Chrome; "Adicionar ao ecrã principal" faz um atalho). Só consulta.
FICH_HTML = os.path.join(PASTA, "_Painel_telemovel.html")
HTML_TPL = r'''<!DOCTYPE html>
<html lang="pt">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#1f3b57">
<title>Painel de ações</title>
<style>
:root{--bg:#f4f5f7;--card:#fff;--tx:#1c1e21;--mu:#65676b;--li:#e3e5e8;--ac:#1f3b57}
@media (prefers-color-scheme:dark){:root{--bg:#121417;--card:#1d2024;--tx:#e8eaed;--mu:#9aa0a6;--li:#30343a;--ac:#8ab4f8}}
*{box-sizing:border-box}
body{margin:0;font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;background:var(--bg);color:var(--tx);font-size:15px}
html,body{overscroll-behavior-y:contain}
#ptr{display:none;position:fixed;top:6px;left:50%;transform:translate(-50%,0);z-index:30;background:var(--card);color:var(--tx);
  border:1px solid var(--li);border-radius:18px;padding:7px 14px;font-size:13px;font-weight:600;box-shadow:0 3px 12px rgba(0,0,0,.2)}
#ptr.ok{background:var(--ac);color:#fff}
header{position:sticky;top:0;z-index:5;background:var(--ac);color:#fff;padding:10px 12px 8px}
@media (prefers-color-scheme:dark){header{background:#1f3b57}}
h1{font-size:17px;margin:0 0 2px;font-weight:600}
.at{font-size:11px;opacity:.85}
.tabs{display:flex;gap:6px;margin-top:8px}
.tab{flex:1;border:0;border-radius:8px;padding:8px 4px;background:rgba(255,255,255,.15);color:#fff;font-size:14px}
.tab.on{background:#fff;color:#1f3b57;font-weight:600}
.bar{display:flex;gap:6px;margin-top:8px}.bar select{flex:1.4;min-width:0}#nmb{display:none;align-items:center;gap:8px;margin-top:6px;color:#fff;font-size:14px}#nmb input{flex:0 0 70px;text-align:center;font-weight:700}
input,select{font-size:14px;border:0;border-radius:8px;padding:8px 10px}
input{flex:1;min-width:0}
main{padding:8px 10px 40px}
.card{background:var(--card);border-radius:12px;margin:8px 0;padding:10px 12px;border:1px solid var(--li)}
.top{display:flex;align-items:center;gap:10px}
.nota{min-width:46px;text-align:center;border-radius:10px;padding:6px 4px;font-weight:700;font-size:17px;color:#1c1e21}
.nm{flex:1;min-width:0}
.nm b{display:block;font-size:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.nm span{font-size:12px;color:var(--mu)}
.px{text-align:right;font-size:14px;font-weight:600;white-space:nowrap}
.px span{display:block;font-size:11px;font-weight:400;color:var(--mu)}
.dec{display:inline-block;margin-top:6px;font-size:12px;border-radius:6px;padding:2px 8px;color:#1c1e21}
.nm b .det{color:#d4a017;font-weight:700}
.lv{display:grid;grid-template-columns:auto 1fr 1fr;gap:6px;margin-top:8px;align-items:stretch}
.lv .rt{background:none;padding:6px 2px 6px 0;display:flex;align-items:center;font-size:12px;font-weight:600;color:var(--tx)}
.lv div{background:var(--bg);border-radius:8px;padding:6px 8px;font-size:12px;color:var(--mu)}
.lv b{display:block;font-size:14px;color:var(--tx)}
.lv .red b{color:#d93025}
.more{display:none;margin-top:8px;font-size:13px;line-height:1.45;border-top:1px solid var(--li);padding-top:8px}
.card.open .more{display:block}
.more p{margin:4px 0}
.more i{color:var(--mu);font-style:normal}
.alerta{color:#d93025;font-weight:600}
.vazio{text-align:center;color:var(--mu);margin-top:40px}
.hold{display:inline-block;margin:6px 0 0 8px;font-size:13px;font-weight:600;color:var(--tx)}
.pos{display:grid;grid-template-columns:0.8fr 1fr 1.5fr 1.2fr;gap:5px;margin-top:8px}
.pos div{background:#fff8e1;border-radius:8px;padding:5px 8px;font-size:11px;color:var(--mu)}
@media (prefers-color-scheme:dark){.pos div{background:#3a3220}}
.pos b{display:block;font-size:13px;color:var(--tx)}
.gr{background:var(--card);border:1px solid var(--li);border-radius:12px;margin:10px 0;padding:10px 12px;text-align:center}
.gr h3{font-size:14px;margin:2px 0 6px;text-align:left}
.lg{display:flex;align-items:center;gap:8px;font-size:12px;text-align:left;padding:2px 0}
.lg i{width:11px;height:11px;border-radius:3px;flex:none}.lg span{flex:1}.lg b{font-weight:600}
@media (max-width:420px){.pos{grid-template-columns:1fr 1fr}}
.card.ds{background:#fde7ea}
@media (prefers-color-scheme:dark){.card.ds{background:#3b2228}}
#glob{background:var(--card);border:1px solid var(--li);border-radius:12px;margin:8px 10px 0;padding:8px 12px;font-size:13px}
#glob b{font-size:15px}
.sec{font-size:15px;font-weight:600;margin:14px 2px 4px;padding-bottom:4px;border-bottom:2px solid var(--li)}
.cfg{display:grid;grid-template-columns:1fr 1fr;gap:6px 10px}.cfg label{display:flex;flex-direction:column;font-size:12px;color:var(--mu);gap:3px}
.cfg input{width:100%;flex:none;background:#fff2cc;color:#1c1e21;font-weight:700;font-size:15px;text-align:center;padding:6px}
.chk{display:flex;align-items:center;gap:8px;margin-top:8px;font-size:13px}.chk label{display:flex;align-items:center;gap:8px}
.chk input{flex:none;width:18px;height:18px}
.inf{flex:none;width:20px;height:20px;border-radius:50%;border:1px solid var(--mu);background:none;color:var(--mu);font:italic 700 12px Georgia,serif;padding:0;line-height:18px}
.infx{display:none;color:var(--mu);font-size:11px;margin:3px 0 0 26px}.infx.on{display:block}
.lv3{display:grid;grid-template-columns:1fr 1.2fr 0.9fr;gap:6px;margin-top:8px}
.lv3 div{background:var(--bg);border-radius:8px;padding:6px 8px;font-size:11px;color:var(--mu)}.lv3 b{display:block;font-size:13px;color:var(--tx)}
.lv3 b.cmx{font-size:16px}
.cbt{display:flex;gap:6px;margin-top:8px}.cbt button{flex:1;border:1px solid var(--li);border-radius:8px;padding:6px;background:var(--card);color:var(--tx);font-size:12px}
.cbt button.on{background:#7f7f7f;color:#fff;border-color:#7f7f7f}
.card.fora{opacity:.55}.card.nula{opacity:.6;padding:8px 12px}.card.fora .nm b,.card.fora .px{text-decoration:line-through}
</style>
</head>
<body>
<header>
<div style="display:flex;align-items:center;gap:8px"><h1 style="flex:1">Painel de ações</h1>
<button id="bt" style="display:none;border:0;border-radius:8px;padding:6px 10px;background:rgba(255,255,255,.2);color:#fff;font-size:14px">↻ Atualizar tudo</button></div>
<div id="mn" style="display:none;margin-top:6px;gap:6px;flex-wrap:wrap"><button class="tab" data-a="precos">Preços</button><button class="tab" data-a="completa">Tudo</button><button class="tab" data-a="assinalados">Assinalados</button><input id="ng" class="tab" type="number" min="0" max="5" step="0.1" placeholder="NG ≥" style="width:70px"><button class="tab" id="ngok">OK</button></div>
<div class="at" id="av" style="display:none;margin-top:4px;font-weight:600"></div>
<div class="at" id="at"></div>
<div class="tabs"><button class="tab on" data-t="A">Ações</button><button class="tab" data-t="E">ETFs</button><button class="tab" data-t="X">Detidos</button><button class="tab" data-t="C">Compra</button><button class="tab" data-t="G">Gráficos</button></div>
<div class="bar"><input id="q" type="search" placeholder="Procurar nome ou código"><select id="o">
<option value="g5:d">Nota por ganhos (0-5)</option><option value="ng:d">Nota global (nota geral + preço)</option><option value="nt">Nota geral (sem o preço)</option><option value="bpg">Mais barata primeiro, nota global igual ou maior a…</option><option value="ngp">Melhor nota global, nível de preço igual ou maior a…</option><option value="bp">Melhor preço para comprar</option><option value="bpm">Mais barata primeiro, nota igual ou maior a…</option><option value="ntp">Melhor nota, nível de preço igual ou maior a…</option><option value="pt:d">Pontuação do sinal</option><option value="np:d">Nível de preço (mais barata primeiro)</option><option value="q:d">Qualidade</option><option value="s:d">Saúde financeira</option><option value="o.pot:d">Analistas (alvo / preço)</option><option value="o.justo_pct:d">Valor justo / preço</option><option value="be:a">Beta (mais defensivas primeiro)</option><option value="be:d">Beta (mais voláteis primeiro)</option><option value="o.var1d:d">Subida no último dia</option><option value="o.var1m:d">Subida no último mês</option><option value="o.val_sempre:d">Valorização desde sempre</option><option value="o.val10:d">Valorização 10 anos</option><option value="o.val5:d">Valorização 5 anos</option><option value="o.val3:d">Valorização 3 anos</option><option value="o.val1:d">Valorização 1 ano</option><option value="o.dist0:a">Mais abaixo do máx. da semana</option><option value="vd">Valor detido</option><option value="tmp">Compra máx. mais perto</option><option value="n">Nome</option></select></div><div id="nmb"><span id="nml">Nota geral igual ou maior a</span><input id="nm" type="text" inputmode="decimal" value="4"><span id="nmc" style="opacity:.8"></span></div>
</header>
<div id="glob" style="display:none"></div>
<div id="qd" style="display:none;position:fixed;top:70px;left:8px;right:8px;z-index:20;background:#fffbe6;color:#222;border:1px solid #c9b458;border-radius:12px;padding:10px 34px 10px 12px;box-shadow:0 6px 24px rgba(0,0,0,.25);font-size:13px;line-height:1.5"><button id="qdx" style="position:absolute;top:6px;right:6px;width:24px;height:24px;border:0;border-radius:50%;background:#d93025;color:#fff;font-weight:700">✕</button><div id="qdt"></div></div>
<main id="lst"></main>
<script>
const D=__DATA__;
document.getElementById('at').textContent=D.at;
// ao abrir a app (ou ao voltar a ela): se já houver dados mais recentes publicados, recarrega sozinha
async function verNovos(){
  if(location.protocol==='file:')return;
  try{const v=(await (await fetch('versao.txt?'+Date.now(),{cache:'no-store'})).text()).trim();
    let ult=0;try{ult=+sessionStorage.getItem('rec')||0;}catch(e){}
    if(v&&D.v&&v>D.v&&Date.now()-ult>60000){try{sessionStorage.setItem('rec',Date.now());}catch(e){}
      location.replace(location.pathname+'?v='+v);}}catch(e){}
}
verNovos();
document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible')verNovos();});
setTimeout(()=>{const t=D.a.concat(D.e).filter(d=>d.ce!=null&&d.vd!=null);if(!t.length)return;
  const c=t.reduce((s,d)=>s+d.ce,0),v=t.reduce((s,d)=>s+d.vd,0),l=v-c,cor=l>=0?'#1e8e3e':'#d93025';
  const g=document.getElementById('glob');g.style.display='block';
  const sub=l=>{const t=l.filter(d=>d.ce!=null&&d.vd!=null);return [t.reduce((s,d)=>s+d.ce,0),t.reduce((s,d)=>s+d.vd,0)];};
  const bm=l=>{const t=l.filter(d=>d.vd>0&&d.be!=null),w=t.reduce((s,d)=>s+d.vd,0);return w?t.reduce((s,d)=>s+d.vd*d.be,0)/w:null;};
  const [ca,va]=sub(D.a),[ce,ve]=sub(D.e),lin=(n,cc,vv,q,b)=>`<b>${n}:</b> compra ${eur(cc)}${q?` (${num(cc/c*100,0)}%/carteira)`:''} / atual ${eur(vv)} <b style="color:${vv>=cc?'#1e8e3e':'#d93025'}">(${pct2(cc?vv/cc-1:0)})</b>${b!=null?` · beta <b>${num(b,2)}</b>`:''}`;
  document.getElementById('qdt').innerHTML=[lin('Ações',ca,va,1,bm(D.a)),lin('ETFs',ce,ve,1,bm(D.e)),lin('Total',c,v,0,bm(D.a.concat(D.e)))].join('<br>')
    +'<div style="font-size:11px;color:#666;margin-top:4px">Beta = volatilidade comparada com o mercado (S&amp;P 500): 1 = oscila como o mercado; 1,3 = oscila 30% mais; 0,7 = 30% menos. Médias ponderadas pelo valor atual.</div>';
  document.getElementById('qdx').onclick=e=>{e.stopPropagation();document.getElementById('qd').style.display='none';};
  g.style.cursor='pointer';
  g.onclick=()=>{document.getElementById('qd').style.display='block';};
  const bg=bm(D.a.concat(D.e));
  g.innerHTML=`<div style="display:flex;justify-content:space-between;align-items:baseline"><b>Posição global</b><span style="font-size:11px;color:var(--mu)">▸ ver posição</span></div>`
    +`<div>Investido/atual: <b style="white-space:nowrap">${eur(c)}</b> / <b style="white-space:nowrap">${eur(v)}</b></div>`
    +`<div style="display:flex;justify-content:space-between;align-items:baseline;gap:8px"><span>Resultado: <b style="color:${cor};white-space:nowrap">${l>=0?'+':'−'}${eur(Math.abs(l))} (${pct2(c?v/c-1:0)})</b></span>${bg!=null?`<span style="white-space:nowrap">β <b>${num(bg,2)}</b></span>`:''}</div>`;},0);
if(D.gh&&D.gh.t){
  const bt=document.getElementById('bt'),mn=document.getElementById('mn'),av=document.getElementById('av');
  bt.style.display='block';
  bt.onclick=()=>pedirTudo();
  const ng=document.getElementById('ng');
  document.getElementById('ngok').onclick=()=>{const v=parseFloat(String(ng.value).replace(',','.'));
    if(!(v>=0&&v<=5)){av.style.display='block';av.textContent='Escreve uma nota geral de 0 a 5.';return;}
    pedir({dataset:{a:'ng:'+v}});ng.value='';};
  mn.querySelectorAll('button[data-a]').forEach(b=>b.onclick=()=>pedir(b));
  window.pedirTudo=()=>{
    let u=0;try{u=+sessionStorage.getItem('pedido')||0;}catch(e){}
    if(Date.now()-u<15*60000){av.style.display='block';
      av.textContent='Atualização já pedida há '+Math.max(1,Math.round((Date.now()-u)/60000))+' min: demora cerca de 10-15 min e a página recarrega sozinha.';return;}
    try{sessionStorage.setItem('pedido',Date.now());}catch(e){}
    pedir({dataset:{a:'completa'}});};
  const pedirTudo=window.pedirTudo;
  async function pedir(b){
    mn.style.display='none';av.style.display='block';av.textContent='A pedir a atualização…';
    try{
      const r=await fetch(`https://api.github.com/repos/${D.gh.o}/${D.gh.r}/actions/workflows/painel.yml/dispatches`,{method:'POST',
        headers:{'Authorization':'Bearer '+D.gh.t,'Accept':'application/vnd.github+json'},body:JSON.stringify({ref:'main',inputs:{tipo:b.dataset.a}})});
      if(r.status!==204){try{sessionStorage.removeItem('pedido');}catch(e){} throw new Error('erro '+r.status);}
      av.textContent='Atualização de tudo pedida (cerca de 10-15 min). A página recarrega sozinha quando os dados novos chegarem.';
      const ini=Date.now();
      const t=setInterval(async()=>{
        try{const v=(await (await fetch('versao.txt?'+Date.now(),{cache:'no-store'})).text()).trim();
          if(v&&v!==D.v){clearInterval(t);location.replace(location.pathname+'?v='+v);}}catch(e){}
        if(Date.now()-ini>30*60000){clearInterval(t);av.textContent='Ainda sem dados novos ao fim de 30 min: recarrega mais tarde.';}
      },30000);
    }catch(e){av.textContent='Não foi possível pedir a atualização ('+e.message+').';}
  }
}
// puxar para baixo no topo da lista: pede uma atualização dos preços (sem o token: só vai ver se há dados novos)
(()=>{
  const ptr=document.createElement('div');ptr.id='ptr';document.body.appendChild(ptr);
  let y0=null,dy=0;const LIM=90;
  addEventListener('touchstart',e=>{const qd=document.getElementById('qd');
    y0=(scrollY<=0&&qd.style.display!=='block'&&e.touches.length===1)?e.touches[0].clientY:null;dy=0;},{passive:true});
  addEventListener('touchmove',e=>{if(y0==null)return;dy=e.touches[0].clientY-y0;
    if(dy>12){ptr.style.display='block';ptr.style.transform=`translate(-50%,${Math.min(dy,140)/2}px)`;
      ptr.textContent=dy>LIM?'↻ Solta para atualizar tudo':'↓ Puxa para atualizar tudo';ptr.classList.toggle('ok',dy>LIM);}
    else ptr.style.display='none';},{passive:true});
  addEventListener('touchend',()=>{if(y0==null)return;ptr.style.display='none';y0=null;if(dy>LIM)puxou();});
  function puxou(){
    const av=document.getElementById('av');av.style.display='block';
    if(!window.pedirTudo){av.textContent='A ver se há dados novos…';verNovos().then(()=>{av.textContent='Sem dados novos por agora.';});return;}
    window.pedirTudo();
  }
})();
let aba='A';
const corNivel=v=>{const m=(a,b,t)=>Math.round(a+(b-a)*t),h=(r,g,b)=>'#'+[r,g,b].map(x=>x.toString(16).padStart(2,'0')).join('');
  const t=Math.max(1,Math.min(5,v));
  const c=t<=3?[m(0xF8,255,(t-1)/2),m(0x69,255,(t-1)/2),m(0x6B,255,(t-1)/2)]:[m(255,0x63,(t-3)/2),m(255,0xBE,(t-3)/2),m(255,0x7B,(t-3)/2)];
  return [h(...c),'#1c1e21'];};
const cor=v=>v==null?'#d0d3d8':v>=4.5?'#63BE7B':v>=3.75?'#A9D18E':v>=2.75?'#E2EFDA':v>=2.25?'#F4B183':v>=1.5?'#FFEB9C':'#F8696B';
const corG=v=>v==null?'#d0d3d8':v>=3.95?'#63BE7B':v>=3.5?'#A9D18E':v>=3?'#E2EFDA':v>=2.25?'#F4B183':v>=1.5?'#FFEB9C':'#F8696B';
const num=(v,c)=>v==null?'–':v.toLocaleString('pt-PT',{minimumFractionDigits:c,maximumFractionDigits:c});
const pct=v=>v==null?'':(v>0?'+':'')+num(v*100,1)+'%';
const pct2=v=>v==null?'':(v>0?'+':'')+num(v*100,2)+'%';
const pr=v=>v==null?'–':num(v,v>=1000?0:2);
const eur=v=>'€ '+num(v,2);
const esc=s=>(s==null?'':String(s)).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
function cartao(d){
  const dc=d.dc||'';
  const corDec=d.pt===1?'#F8696B':d.pt===2.5?'#F4B183':corDecF(d);
  const azul=v=>v<=1.25?['#1F4E79','#fff']:v<=1.5?['#2F75B5','#fff']:v<=1.75?['#9BC2E6','#1F4E79']:['#BDD7EE','#1F4E79'];
  const NG=d.ng!=null?d.ng:d.nt, corNG=v=>d.ng!=null&&d.np!=null?corG(v):cor(v);   // nota global (geral + preço)
  const [cn,fn]=d.ca?azul(d.np):[corNG(NG),'inherit'];
  let h=`<div class="card${d.ds?' ds':''}"><div class="top"><div class="nota" style="background:${cn};color:${fn}" title="Nota global (nota geral + preço)${d.np!=null?' · nível de preço '+num(d.np,1)+'/5':''}">${NG==null?'–':num(NG,1)}</div>
  <div class="nm"><b>${d.x?'<span class="det">★ </span>':''}${d.t?`<a href="https://finance.yahoo.com/quote/${encodeURIComponent(d.t)}/" target="_blank" rel="noopener" onclick="event.stopPropagation()" style="color:inherit;text-decoration:underline dotted">${esc(d.n)}</a>`:esc(d.n)}</b><span>${esc(d.c)} · ${esc(d.t)}</span></div>
  <div class="px">${pr(d.p)} ${esc(d.m)}<span>sinal ${d.pt==null?'–':num(d.pt,d.pt%1?1:0)}</span></div></div>
  ${(d.be!=null||(d.pr!=null&&d.pr>0))?`<span style="float:right;display:flex;flex-direction:column;align-items:flex-end;gap:3px;margin-left:4px">${d.pr!=null&&d.pr>0?`<span class="dec" title="PER (preço / lucro dos últimos 12 meses)" style="background:${d.pr<=15?'#C6E5B3':d.pr>40?'#F8A5A6':d.pr>25?'#FCE4D6':'#eceff3'};color:#1c1e21">PE ${num(d.pr,d.pr>=100?0:1)}</span>`:''}${d.be!=null?`<span class="dec" title="Beta: volatilidade face ao S&amp;P 500" style="background:${d.be<0.8?'#C6E5B3':d.be>1.5?'#F8A5A6':d.be>1.2?'#FCE4D6':'#eceff3'};color:#1c1e21">β ${num(d.be,2)}</span>`:''}</span>`:''}<span class="dec" style="background:${corDec}">${esc(dc)}</span>${d.nt!=null&&d.ng!=null?` <span class="dec" style="background:${cor(d.nt)}" title="Nota geral: negócio, tendência e momento, sem o preço">geral ${num(d.nt,1)}</span>`:''}${d.g5!=null?` <span class="dec" style="background:${d.g5>=3.95?'#63BE7B':d.g5>=3.5?'#A9D18E':d.g5>=3?'#E2EFDA':d.g5>=2.25?'#F4B183':d.g5>=1.5?'#FFEB9C':'#F8696B'}" title="Nota por ganhos (0-5): rentabilidade anual ponderada pelo tempo, ajustada ao preço">ganhos ${num(d.g5,1)}${d.gn!=null?' <small style="font-size:80%;font-weight:normal">('+(d.gn>=0?'+':'')+num(d.gn*100,1)+'%/ano)</small>':''}</span>`:''}${d.np!=null?` <span class="dec" style="background:${d.nc||'#eceff3'};color:${/muito/.test(d.nr||'')?'#fff':'#1c1e21'}">${num(d.np,1)} ${esc(d.nr||'')}</span>`:''}${d.qn!=null?`<div class="pos"><div>Nº${d.br?' · '+d.br:''}<b>${num(d.qn,2)}</b></div><div>Preço médio<b>${d.pc!=null?pr(d.pc):'–'}</b></div>
  <div>Valor total<b>${d.vd!=null?eur(d.vd):'–'}${d.lpp!=null?` <span style="color:${d.lpp>=0?'#1e8e3e':'#d93025'}">(${pct2(d.lpp)})</span>`:''}</b></div>
  <div>Lucro / prejuízo<b>${d.le!=null?`<span style="color:${d.le>=0?'#1e8e3e':'#d93025'}">${d.le>=0?'+':'−'}${eur(Math.abs(d.le))}</span>`:'–'}</b></div></div>`:''}
  <div class="lv"><div class="rt">Compra</div><div>Preço compra (Limite)<b>${pr(d.tm)} <small>${pct(d.tmp)}</small></b></div>
  <div style="background:none;border:0"></div>
  <div class="rt">Vendas</div><div>Razoável<b>${pr(d.va)} <small>${pct(d.vap)}</small></b></div>
  <div class="red">Rede de segurança<b>${pr(d.vm)} <small>${pct(d.vmp)}</small></b>${d.vmc!=null?`<small style="display:block;color:${d.vmc>=0?'#1e8e3e':'#d93025'}">${d.vmc>=0?'ganhas':'perdes'} ${pct(Math.abs(d.vmc)).replace('+','')} sobre a tua compra</small>`:''}</div></div><div class="more">`;
  if(d.ea!=null) h+=`<p><i>Nota ponderada das ações:</i> ${num(d.ea,1)}</p>`;
  if(d.be!=null) h+=`<p><i>Beta:</i> ${num(d.be,2)} <small>(${d.be>1.2?'oscila mais do que o mercado':d.be<0.8?'mais defensiva do que o mercado':'acompanha o mercado'})</small></p>`;
  if(d.q!=null||d.s!=null) h+=`<p><i>Qualidade:</i> ${d.q==null?'–':num(d.q,1)} · <i>Saúde:</i> ${d.s==null?'–':num(d.s,1)}</p>`;
  if(d.sd) h+=`<p class="${String(d.sd).startsWith('ALERTA')?'alerta':''}">${esc(d.sd)}</p>`;
  if(d.sg) h+=`<p><i>Médias 100/200:</i> ${esc(d.sg)}</p>`;
  if(d.lp) h+=`<p><i>Longo prazo:</i> ${esc(d.lp)}</p>`;
  if(d.ed) h+=`<p><i>ETF:</i> ${esc(d.ed)}</p>`;
  if(d.ec) h+=`<p><i>Composição:</i> ${esc(d.ec)}</p>`;
  if(d.cm) h+=`<p><i>Comentários:</i> ${esc(d.cm)}</p>`;
  if(d.no) h+=`<p><i>Notas:</i> ${esc(d.no)}</p>`;
  return h+'</div></div>';
}
// ---- separador Compra (como a folha Compra do Excel): quanto investir em cada título
const CMP0={a:3.5,e:3.5,t:1000,p:60,s:1,h:1,vq:1,vd:1,vn:1,fa:[],fe:[]};
// valem sempre as escolhas guardadas mais recentes: as do telemóvel ou as da folha Compra (no último ▶ CALCULAR COMPRA)
const tsX=D.cmp&&D.cmp.ts||0;
const cmpCfg=()=>{let c=null;try{c=JSON.parse(localStorage.getItem('cmp_cfg')||'null');}catch(e){}
  if(c&&(c.ts||0)<tsX)c=null;
  const x=Object.assign({},CMP0,D.cmp||{},c||{});delete x.nq;delete x.nd;return x;};
const cmpX=()=>{let c=null;try{c=JSON.parse(localStorage.getItem('cmp_x')||'null');}catch(e){}
  if(c&&(c.ts||0)<tsX)c=null;
  return c||{nq:(D.cmp&&D.cmp.nq)||[],nd:(D.cmp&&D.cmp.nd)||[]};};
const nPt=v=>{const x=parseFloat(String(v).replace(/\s|€|%/g,'').replace(',','.'));return isNaN(x)?null:x;};
const corT=d=>d.np!=null?corG(d.ngc):cor(d.ngc);   // ações com nível de preço: 3,95/3,5/3; ETFs: 4,5/3,75/2,75
function corDecF(d){const dc=d.dc||'';return dc.includes(': esperar')?'#D9C3E9':/saúde frágil/i.test(dc)?'#F8CBAD':
  dc.startsWith('Venda')?'#F8696B':/Longo Prazo/.test(dc)?'#F4B183':dc.startsWith('Manter')?'#FFEB9C':
  dc.startsWith('Compra forte')?'#63BE7B':dc.startsWith('Compra média')?'#A9D18E':dc.startsWith('Compra')?'#E2EFDA':'#FFEB9C';}
function compra(){
  const c=cmpCfg(),ck=v=>v?' checked':'';
  return `<div class="card" id="cf" style="font-size:13px"><div class="cfg">
  <label>Ações: nota global ≥<input id="ca" inputmode="decimal" value="${num(c.a,2)}"></label>
  <label>ETFs: nota global ≥<input id="ce" inputmode="decimal" value="${num(c.e,2)}"></label>
  <label>Valor total (€)<input id="ct" inputmode="decimal" value="${num(c.t,2)}"></label>
  <label>% em ações<input id="cp" inputmode="decimal" value="${num(c.p,2)}%"></label></div>
  <div class="chk"><label><input type="checkbox" id="cs"${ck(c.s)}> Só os que o painel manda comprar</label><button class="inf" data-i="i1">i</button></div>
  <div class="infx" id="i1">Ligado: só entram decisões «Compra forte / média / Compra» (também «… : esperar»). Desligado: entram também os que estão em Manter (ou Venda) mas têm a nota global acima do mínimo.</div>
  <div class="chk"><label><input type="checkbox" id="ch"${ck(c.h)}> Incluir os que já tenho</label><button class="inf" data-i="i2">i</button></div>
  <div class="infx" id="i2">Ligado: entram também os títulos que já tens (★), para reforçar. Desligado: só os que ainda não tens.</div>
  <div class="chk"><label><input type="checkbox" id="cvq"${ck(c.vq)}> Mostrar os «Não quero»</label></div>
  <div class="chk"><label><input type="checkbox" id="cvd"${ck(c.vd)}> Mostrar os «Não disponível»</label></div>
  <div class="chk"><label><input type="checkbox" id="cvn"${ck(c.vn)}> Mostrar títulos com compra nula</label><button class="inf" data-i="i3">i</button></div>
  <div class="infx" id="i3">Títulos com 0 € (0% do grupo). Cada quadrado desligado esconde sempre: sem este, nenhum título com 0 € aparece; sem os de cima, os «Não quero» / «Não disponível» nunca aparecem. O que não aparece (também pela procura) não recebe valor: o dinheiro do grupo vai só para os que se veem.</div>
  <div class="chk" style="margin-top:6px"><b>📌 Fixos</b><button class="inf" data-i="i4">i</button></div>
  <div class="infx" id="i4">Até 5 ações e 5 ETFs que entram sempre, com a % escrita do valor do seu grupo (ex.: 60% dos ETFs no S&amp;P 500); o resto do grupo divide-se pela nota global pelos outros. Vale a escolha mais recente: a daqui ou a da folha Compra do Excel.</div>
  ${[['a','Ações',D.a],['e','ETFs',D.e]].map(([k,nm,L])=>[0,1,2,3,4].map(i=>{const f=(c['f'+k]||[])[i]||[],t=String(f[0]||'').toUpperCase();
    return `<div style="display:flex;gap:6px;margin:3px 0;align-items:center"><span id="fn${k}${i}" class="nota" style="min-width:40px;font-size:13px;padding:5px 2px"></span><select id="fs${k}${i}" style="flex:1;min-width:0;font-size:13px"><option value="">${nm}: fixo ${i+1} —</option>${L.map(d=>`<option value="${esc(String(d.t||'').toUpperCase())}"${String(d.t||'').toUpperCase()===t?' selected':''}>${esc(d.n)}</option>`).join('')}</select><input id="fp${k}${i}" inputmode="decimal" placeholder="%" style="width:62px;background:#FFF2CC;border:0;border-radius:8px;text-align:center;font-weight:700;color:#1c1e21" value="${f[1]!=null&&f[1]>0?num(f[1]*100,1)+'%':''}"></div>`;}).join('')).join('')}
  <div id="cx" style="color:var(--mu);font-size:11px;margin-top:4px"></div></div><div id="cr"></div>`;
}
function calcCompra(){
  const g=id=>document.getElementById(id),c={a:nPt(g('ca').value)??CMP0.a,e:nPt(g('ce').value)??CMP0.e,
    t:nPt(g('ct').value)??CMP0.t,p:nPt(g('cp').value)??CMP0.p,s:g('cs').checked?1:0,h:g('ch').checked?1:0,
    vq:g('cvq').checked?1:0,vd:g('cvd').checked?1:0,vn:g('cvn').checked?1:0};
  if(c.p>0&&c.p<=1)c.p*=100;
  const porT=new Map([...D.a,...D.e].map(d=>[String(d.t||'').toUpperCase(),d]));
  for(const k of ['a','e'])for(let i=0;i<5;i++){const d=porT.get(g('fs'+k+i).value),b=g('fn'+k+i),v=d?(d.ng??d.nt):null;  // V167: nota atual
    b.textContent=v==null?'':num(v,2);b.style.background=v==null?'transparent':(d.np!=null?corG(v):cor(v));}
  for(const k of ['a','e']){const l=[];for(let i=0;i<5;i++){const t=g('fs'+k+i).value;let v=nPt(g('fp'+k+i).value);
    if(v!=null&&v>1)v/=100;if(t&&v>0)l.push([t,Math.round(v*10000)/10000]);}c['f'+k]=l;}   // V164: fixos
  const c0=cmpCfg();
  if(['a','e','t','p','s','h','vq','vd','vn'].some(k=>c[k]!==c0[k])||JSON.stringify(c.fa)!==JSON.stringify((c0.fa||[]).map(x=>[x[0],x[1]]))||JSON.stringify(c.fe)!==JSON.stringify((c0.fe||[]).map(x=>[x[0],x[1]]))){c.ts=Date.now();try{localStorage.setItem('cmp_cfg',JSON.stringify(c));}catch(e){}}
  const X=cmpX(),nq=new Set(X.nq),nd=new Set(X.nd),q=(document.getElementById('q').value||'').trim().toLowerCase();
  const txtN=(v,L)=>v>=L[0]?'= Compra forte':v>=L[1]?'= Compra média':v>=L[2]?'= Compra':'= abaixo de Compra (entram títulos que o painel não aconselha comprar)';
  const LA=[3.95,3.5,3],LE=[4.5,3.75,2.75];
  g('cx').textContent=`Ações ${txtN(c.a,LA)} · ETFs ${txtN(c.e,LE)} · ações ${eur(c.t*c.p/100)} · ETFs ${eur(c.t*(1-c.p/100))}`;
  let h='';
  const todosT=new Map([...D.a,...D.e].map(d=>[String(d.t||'').toUpperCase(),d]));
  for(const [nome,lista,mn,val,FXL] of [['Ações',D.a,c.a,c.t*c.p/100,c.fa||[]],['ETFs',D.e,c.e,c.t*(1-c.p/100),c.fe||[]]]){
    const esc_=[],fixT=new Set();
    for(const [tk,pc] of FXL){const d=todosT.get(String(tk).toUpperCase());   // V163: fixos (definidos na folha Compra do Excel)
      if(!d||fixT.has(String(tk).toUpperCase())||!(pc>0))continue;fixT.add(String(tk).toUpperCase());
      esc_.push({...d,ngc:d.ng??d.nt??0,fixo:pc});}
    for(const d of lista){
      if(fixT.has(String(d.t||'').toUpperCase()))continue;
      let ng=d.ng??d.nt;const dc=d.dc||'';
      const L=d.np!=null?LA:LE;
      // a nota global da tabela (também nos «esperar»), como no Excel
      if(ng==null||Math.round(ng*100)/100<mn-1e-9)continue;
      if(c.s&&!dc.startsWith('Compra'))continue;
      if(d.x&&!c.h)continue;
      esc_.push({...d,ngc:ng});}
    esc_.sort((a,b)=>(b.fixo!=null)-(a.fixo!=null)||b.ngc-a.ngc);
    const T=d=>String(d.t||'').toUpperCase(),acha=d=>!q||(d.n+' '+d.c+' '+d.t).toLowerCase().includes(q);
    // o que não aparece (procura, «não quero», «não disponível») fica sem valor; cada quadrado desligado esconde sempre
    const fora=d=>nq.has(T(d))||nd.has(T(d))||!acha(d),soma=esc_.filter(d=>!fora(d)&&d.fixo==null).reduce((s,d)=>s+d.ngc,0);
    const sfx=esc_.filter(d=>!fora(d)&&d.fixo!=null).reduce((s,d)=>s+d.fixo,0),resto=Math.max(0,1-sfx);
    const wd=d=>fora(d)?0:d.fixo!=null?d.fixo/Math.max(1,sfx):(!soma?0:resto*d.ngc/soma);
    let tv=0;
    const vis=d=>acha(d)&&(c.vq||!nq.has(T(d)))&&(c.vd||!nd.has(T(d)))&&(c.vn||wd(d)*val>=0.005);   // 0 € (também com o grupo a 0 €)
    const cards=esc_.filter(vis).map(d=>{
      const f=fora(d),w=wd(d),v=w*val,pe=d.p!=null&&d.fx!=null?d.p*d.fx:null,na=pe&&v>0?v/pe:null;
      return `<div class="card cmp${f?' fora':''}"><div class="top"><div class="nota" style="background:${corT(d)}">${num(d.ngc,2)}</div>
      <div class="nm"><b>${d.fixo!=null?'📌 ':''}${d.x?'<span class="det">★ </span>':''}${d.t?`<a href="https://finance.yahoo.com/chart/${encodeURIComponent(d.t)}" target="_blank" rel="noopener" style="color:inherit;text-decoration:underline dotted">${esc(d.n)}</a>`:esc(d.n)}${d.fixo!=null?' <small style="font-weight:400;color:var(--mu)">(fixo)</small>':''}</b><span>${esc(d.c)} · ${d.t?`<a href="https://finance.yahoo.com/chart/${encodeURIComponent(d.t)}" target="_blank" rel="noopener" style="color:inherit">${esc(d.t)} 📈</a>`:''}</span></div>
      <div class="px">${eur(v)}<span>${num(w*100,1)}% do grupo</span></div></div>
      <span class="dec" style="background:${corDecF(d)};${dc_(d)}">${esc(d.dc)}</span>
      <div class="lv3"><div>Preço atual<b>${pr(d.p)} ${esc(d.m)}</b></div><div>Preço compra (Limite)<b class="cmx">${pr(d.tm)} <small>${pct(d.tmp)}</small></b></div>
      <div>Nº de ações<b>${na==null?'–':num(na,2)}</b></div></div>
      <div class="cbt"><button data-k="nq" data-t="${esc(T(d))}" class="${nq.has(T(d))?'on':''}">Não quero</button><button data-k="nd" data-t="${esc(T(d))}" class="${nd.has(T(d))?'on':''}">Não disponível</button></div></div>`;});
    esc_.forEach(d=>{tv+=wd(d)*val;});
    const nesc=esc_.filter(d=>!vis(d)).length;
    h+=`<h2 class="sec">${nome} (${esc_.filter(d=>!fora(d)).length}) · ${eur(tv)}${sfx>0?` <span style="font-size:11px;font-weight:400;color:var(--mu)">· 📌 fixos ${num(Math.min(1,sfx)*100,1)}%</span>`:''}${nesc?` <span style="font-size:11px;font-weight:400;color:var(--mu)">· ${nesc} escondido${nesc>1?'s':''}</span>`:''}</h2>`+(esc_.length?cards.join(''):
      `<div class="card" style="color:var(--mu)">Nenhum título cumpre os critérios: os ${eur(val)} ficam por investir (baixa a nota mínima ou desliga «só com decisão de compra»).</div>`);
  }
  g('cr').innerHTML=h;
}
const dc_=d=>(d.dc||'').includes(': esperar')?'color:#5B2C83;font-weight:700':(d.dc||'').startsWith('Venda')?'font-weight:700':'';
function ligarCompra(){
  const cp=document.getElementById('cp');                       // % sempre com o símbolo (60%, não 0,6)
  cp.addEventListener('focus',()=>{cp.value=cp.value.replace('%','').trim();});
  cp.addEventListener('blur',()=>{let v=nPt(cp.value);if(v!=null&&v>0&&v<=1)v*=100;cp.value=v==null?'':num(v,2)+'%';});
  ['ca','ce','ct','cp'].forEach(id=>{const e=document.getElementById(id);e.addEventListener('input',calcCompra);
    e.addEventListener('keydown',ev=>{if(ev.key==='Enter'){ev.preventDefault();e.blur();}});});
  ['cs','ch','cvq','cvd','cvn'].forEach(id=>document.getElementById(id).addEventListener('change',calcCompra));
  for(const k of ['a','e'])for(let i=0;i<5;i++){document.getElementById('fs'+k+i).addEventListener('change',calcCompra);
    const e=document.getElementById('fp'+k+i);e.addEventListener('input',calcCompra);
    e.addEventListener('focus',()=>{e.value=e.value.replace('%','').trim();});
    e.addEventListener('blur',()=>{let v=nPt(e.value);if(v!=null&&v>0&&v<=1)v*=100;e.value=v==null||!(v>0)?'':num(v,1)+'%';});}
  document.querySelectorAll('#cf .inf').forEach(b=>b.addEventListener('click',ev=>{ev.preventDefault();
    document.getElementById(b.dataset.i).classList.toggle('on');}));
  document.getElementById('cr').addEventListener('click',ev=>{const b=ev.target.closest('button[data-k]');if(!b)return;
    const X=cmpX(),k=b.dataset.k,t=b.dataset.t,s=new Set(X[k]);s.has(t)?s.delete(t):s.add(t);X[k]=[...s];X.ts=Date.now();
    try{localStorage.setItem('cmp_x',JSON.stringify(X));}catch(e){}calcCompra();});
  calcCompra();
}
// ---- separador Gráficos (os mesmos da folha Carteira), em SVG
const CORES=['#4e79a7','#f28e2b','#e15759','#76b7b2','#59a14f','#edc948','#b07aa1','#ff9da7','#9c755f','#bab0ac','#1f77b4','#8c564b'];
const escH=t=>String(t).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
function tarte(itens,titulo,lp){
  const tot=itens.reduce((s,x)=>s+x.v,0);if(!tot)return '';
  const liq=lp?itens.reduce((s,x)=>s+x.s*x.v,0):0;
  const R=90,r=52,cx=110,cy=110;let a0=-Math.PI/2,arcos='';
  itens.forEach((x,i)=>{const a1=a0+2*Math.PI*x.v/tot,g=a1-a0>Math.PI?1:0,cor=x.cor||CORES[i%CORES.length];
    const P=(a,rr)=>`${(cx+rr*Math.cos(a)).toFixed(2)} ${(cy+rr*Math.sin(a)).toFixed(2)}`;
    arcos+=itens.length===1?`<circle cx="${cx}" cy="${cy}" r="${(R+r)/2}" fill="none" stroke="${cor}" stroke-width="${R-r}"/>`:
      `<path d="M${P(a0,R)} A${R} ${R} 0 ${g} 1 ${P(a1,R)} L${P(a1,r)} A${r} ${r} 0 ${g} 0 ${P(a0,r)}Z" fill="${cor}" stroke="var(--card)" stroke-width="1.5"/>`;
    a0=a1;});
  const leg=itens.map((x,i)=>`<div class="lg"><i style="background:${x.cor||CORES[i%CORES.length]}"></i><span>${escH(x.n)}</span><b${lp?` style="color:${x.s>0?'#1e8e3e':'#d93025'}"`:''}>${lp?(x.s>0?'+':'−'):''}${eur(x.v)} (${lp&&x.s<0?'−':''}${num(x.v/tot*100,1)}%)</b></div>`).join('');
  return `<section class="gr"><h3>${titulo}</h3><svg viewBox="0 0 220 220" width="220" height="220" role="img" aria-label="${escH(titulo)}">${arcos}
    <text x="110" y="106" text-anchor="middle" font-size="11" fill="var(--mu)">${lp?'saldo':'total'}</text><text x="110" y="124" text-anchor="middle" font-size="14" font-weight="700" fill="${lp?(liq>=0?'#1e8e3e':'#d93025'):'var(--tx)'}">${lp?(liq>=0?'+':'−'):''}${eur(lp?Math.abs(liq):tot)}</text></svg>${leg}</section>`;
}
function barras(t){
  if(!t.length)return '';
  const mx=Math.max(...t.map(x=>Math.max(x.c,x.a)))*1.08,W=360,L=118,Rm=8,h=26,top=34,H=top+t.length*h+30,esc_=v=>L+(W-L-Rm)*v/mx;
  const passo=(()=>{const b=Math.pow(10,Math.floor(Math.log10(mx/4)));return [1,2,5,10].map(k=>k*b).find(k=>mx/k<=5);})();
  let g='';for(let v=0;v<=mx;v+=passo){const x=esc_(v);g+=`<line x1="${x}" y1="${top-4}" x2="${x}" y2="${H-30}" stroke="var(--li)"/><text x="${x}" y="${top-8}" text-anchor="middle" font-size="9" fill="var(--mu)">${v>=1000?num(v/1000,v%1000?1:0)+' mil':num(v,0)}</text>`;}
  t.forEach((x,i)=>{const y=top+i*h,base=Math.min(x.c,x.a),l=x.a-x.c,cor=l>=0?'#1e8e3e':'#d93025';
    g+=`<text x="${L-6}" y="${y+h/2+4}" text-anchor="end" font-size="10" fill="var(--tx)">${escH(x.n.length>18?x.n.slice(0,17)+'…':x.n)}</text>
      <rect x="${L}" y="${y+5}" width="${(esc_(base)-L).toFixed(1)}" height="${h-10}" fill="#9aa0a6"/>
      <rect x="${esc_(base).toFixed(1)}" y="${y+5}" width="${(esc_(Math.max(x.c,x.a))-esc_(base)).toFixed(1)}" height="${h-10}" fill="${cor}"/>
      <text x="${(esc_(Math.max(x.c,x.a))+3).toFixed(1)}" y="${y+h/2+4}" font-size="9" fill="${cor}" font-weight="700">${x.c?pct2(l/x.c):''}</text>`;});
  g+=`<text x="${L+(W-L)/2}" y="12" text-anchor="middle" font-size="10" font-weight="700" fill="var(--mu)">Valor (€)</text>
    <text x="10" y="${top+t.length*h/2}" font-size="10" font-weight="700" fill="var(--mu)" transform="rotate(-90 10 ${top+t.length*h/2})" text-anchor="middle">Título</text>
    <g font-size="10" fill="var(--tx)"><rect x="${L}" y="${H-18}" width="10" height="10" fill="#9aa0a6"/><text x="${L+14}" y="${H-9}">Compra</text>
    <rect x="${L+70}" y="${H-18}" width="10" height="10" fill="#1e8e3e"/><text x="${L+84}" y="${H-9}">Lucro</text>
    <rect x="${L+130}" y="${H-18}" width="10" height="10" fill="#d93025"/><text x="${L+144}" y="${H-9}">Prejuízo</text></g>`;
  return `<section class="gr"><h3>Cada título: valor de compra, + lucro (verde) / − prejuízo (vermelho)</h3><svg viewBox="0 0 ${W+40} ${H}" width="100%" role="img" aria-label="Lucro e prejuízo por título">${g}</svg></section>`;
}
function graficos(){
  const t=D.a.map(d=>({...d,tp:'Ação'})).concat(D.e.map(d=>({...d,tp:'ETF'}))).filter(d=>d.ce!=null&&d.vd!=null)
    .map(d=>({n:d.n,c:d.ce,a:d.vd,tp:d.tp})).sort((x,y)=>y.c-x.c);
  if(!t.length)return '<p class="vazio">Escreve o nº de ações e o preço médio de compra nas folhas Ações / ETFs para ver os gráficos.</p>';
  const ca=t.filter(x=>x.tp==='Ação').reduce((s,x)=>s+x.c,0), ce=t.filter(x=>x.tp==='ETF').reduce((s,x)=>s+x.c,0);
  const va=t.filter(x=>x.tp==='Ação').reduce((s,x)=>s+x.a,0), ve=t.filter(x=>x.tp==='ETF').reduce((s,x)=>s+x.a,0);
  return tarte(t.map(x=>({n:x.n,v:x.c})),'Carteira por valor de compra (€ e %)')
    +tarte([{n:'Ações',v:ca,cor:'#4e79a7'},{n:'ETFs',v:ce,cor:'#f28e2b'}].filter(x=>x.v),'Ações vs ETFs (valor de compra)')
    +tarte([{n:'Ações',v:va,cor:'#4e79a7'},{n:'ETFs',v:ve,cor:'#f28e2b'}].filter(x=>x.v),'Ações vs ETFs (valor atual)')
    +tarte(lucros(t),'Lucros (verde) e prejuízos (vermelho) por título',true)
    +barras(t);
}
function lucros(t){
  const VD=['#1e8e3e','#34a853','#5bb974','#81c995','#0d652d','#a8dab5'],VM=['#d93025','#ea4335','#ee675c','#f28b82','#a50e0e','#f6aea9'];
  const g=t.filter(x=>x.a>x.c).sort((x,y)=>(y.a-y.c)-(x.a-x.c)), p=t.filter(x=>x.a<x.c).sort((x,y)=>(x.a-x.c)-(y.a-y.c));
  return g.map((x,i)=>({n:x.n,v:x.a-x.c,s:1,cor:VD[i%VD.length]})).concat(p.map((x,i)=>({n:x.n,v:x.c-x.a,s:-1,cor:VM[i%VM.length]})));
}
function minAtivo(){const o=document.getElementById('o').value;return ['bpm','ntp','bpg','ngp'].includes(o);}
function minimo(){const m=parseFloat(String(document.getElementById('nm').value).replace(',','.'));return isNaN(m)?0:m;}
function campoMin(){const o=document.getElementById('o').value,v=['bpm','ntp','bpg','ngp'].includes(o),nm=document.getElementById('nm');
  document.getElementById('nmb').style.display=v&&aba!=='G'?'flex':'none';
  if(v){let x=null;try{x=localStorage.getItem('nm_v_'+o);}catch(e){}const pr_=o==='ntp'||o==='ngp';nm.value=x||(pr_?'3,5':'4');
    document.getElementById('nml').textContent=pr_?'Nível de preço igual ou maior a':o==='bpg'?'Nota global igual ou maior a':'Nota geral igual ou maior a';}}
function mostrar(){
  const q=document.getElementById('q').value.trim().toLowerCase(), o=document.getElementById('o').value;
  const prep=l=>{if(q) l=l.filter(d=>(d.n+' '+d.c+' '+d.t).toLowerCase().includes(q));
    if(o.includes(':')){const [k,dr]=o.split(':'),g=d=>{const v=k.startsWith('o.')?(d.or||{})[k.slice(2)]:d[k];return typeof v==='number'?v:null;};
      return l.slice().sort((a,b)=>{const x=g(a),y=g(b);if(x==null&&y==null)return (b.nt??-9)-(a.nt??-9);if(x==null)return 1;if(y==null)return -1;return dr==='a'?x-y:y-x;});}
    if(o==='ngp'){const mn=minimo(),G=d=>d.ng??d.nt??-9,
        ok=d=>d.np!=null&&Math.round(d.np*100)/100>=mn-1e-9;
      return l.filter(ok).sort((a,b)=>(G(b)-G(a))||((b.np??-9)-(a.np??-9)));}
    if(o==='bpg'){const mn=minimo(),g=d=>d.np??-9,j=d=>(d.or||{}).justo_pct??-9,G=d=>d.ng??d.nt,
        ok=d=>G(d)!=null&&Math.round(G(d)*100)/100>=mn-1e-9&&d.np!=null;
      return l.filter(ok).sort((a,b)=>(g(b)-g(a))||(j(b)-j(a))||((G(b)??-9)-(G(a)??-9)));}
    if(o==='ntp'){const mn=minimo(),
        ok=d=>d.np!=null&&Math.round(d.np*100)/100>=mn-1e-9;   // primeiro os que têm o nível de preço pedido, pela nota
      return l.filter(ok).sort((a,b)=>((b.nt??-9)-(a.nt??-9))||((b.np??-9)-(a.np??-9)));}
    if(o==='bpm'){const mn=minimo(),g=d=>d.np??-9,j=d=>(d.or||{}).justo_pct??-9,
        ok=d=>d.nt!=null&&Math.round(d.nt*100)/100>=mn-1e-9&&d.np!=null;   // primeiro os que cumprem a nota, do mais barato ao mais caro
      return l.filter(ok).sort((a,b)=>(g(b)-g(a))||(j(b)-j(a))||((b.nt??-9)-(a.nt??-9)));}
    if(o==='bp'){const g=d=>d.np??-9,j=d=>(d.or||{}).justo_pct??-9;   // nível de preço (5 = muito barata); empate: valor justo
      return l.slice().sort((a,b)=>(g(b)-g(a))||(j(b)-j(a))||((b.nt??-9)-(a.nt??-9)));}
    return l.slice().sort((a,b)=>o==='n'?String(a.n).localeCompare(b.n):o==='tmp'?(b.tmp??-9)-(a.tmp??-9):
      o==='vd'?(b.vd??-1)-(a.vd??-1):(b.nt??-9)-(a.nt??-9));};
  const tot=l=>{const v=l.filter(d=>d.vd!=null);if(!v.length)return '';
    const lp=l.filter(d=>d.le!=null).reduce((s,d)=>s+d.le,0), ce=l.filter(d=>d.ce!=null).reduce((s,d)=>s+d.ce,0);
    return ' · '+eur(v.reduce((s,d)=>s+d.vd,0))+(l.some(d=>d.le!=null)?` · <span style="color:${lp>=0?'#1e8e3e':'#d93025'}">${lp>=0?'+':'−'}${eur(Math.abs(lp))}${ce?' ('+pct2(lp/ce)+')':''}</span>`:'');};
  const el=document.getElementById('lst');
  let h='';
  document.querySelector('.bar').style.display=aba==='G'?'none':'flex';
  {const o2=document.getElementById('o').value;document.getElementById('nmb').style.display=['bpm','ntp','bpg','ngp'].includes(o2)&&aba!=='G'?'flex':'none';}
  if(aba==='G'){el.innerHTML=graficos();return;}
  if(aba==='C'){document.getElementById('o').style.display='none';document.getElementById('nmb').style.display='none';
    if(!document.getElementById('cf')){el.innerHTML=compra();ligarCompra();}else calcCompra();return;}
  document.getElementById('o').style.display='';
  if(aba==='X'){
    const a=prep(D.a.filter(d=>d.x)), e=prep(D.e.filter(d=>d.x));
    if(minAtivo())document.getElementById('nmc').textContent=(a.length+e.length)+' detido(s)';
    const ta=D.a.concat(D.e).filter(d=>d.x&&d.vd!=null);
    if(ta.length) h+=`<p class="hold" style="margin:8px 2px">Total detido${tot(ta)}</p>`;
    if(a.length) h+=`<h2 class="sec">Ações (${a.length})${tot(a)}</h2>`+a.map(cartao).join('');
    if(e.length) h+=`<h2 class="sec">ETFs (${e.length})${tot(e)}</h2>`+e.map(cartao).join('');
  } else {const ls=prep(aba==='A'?D.a:D.e);h=ls.map(cartao).join('');
    if(minAtivo()){document.getElementById('nmc').textContent=ls.length+(ls.length===1?' título':' títulos');
      if(!ls.length)h='<div class="card" style="text-align:center;color:#666">Nenhum título cumpre o mínimo. Experimenta um valor mais baixo.</div>';}}
  el.innerHTML=h||'<p class="vazio">Nada a mostrar.</p>';
  el.querySelectorAll('.card').forEach(c=>c.addEventListener('click',()=>c.classList.toggle('open')));
}
document.querySelectorAll('.tabs .tab').forEach(b=>b.addEventListener('click',()=>{
  document.querySelectorAll('.tabs .tab').forEach(x=>x.classList.remove('on'));b.classList.add('on');aba=b.dataset.t;mostrar();}));
document.getElementById('q').addEventListener('input',mostrar);
document.getElementById('o').addEventListener('change',()=>{try{localStorage.setItem('ordem_app',document.getElementById('o').value);}catch(e){}campoMin();mostrar();});
try{const o0=localStorage.getItem('ordem_app'),sel=document.getElementById('o');
  if(o0&&[...sel.options].some(x=>x.value===o0)){sel.value=o0;campoMin();}}catch(e){}
function aplicarMin(){const nm=document.getElementById('nm');try{localStorage.setItem('nm_v_'+document.getElementById('o').value,nm.value);}catch(e){}mostrar();}
document.getElementById('nm').addEventListener('input',aplicarMin);
document.getElementById('nm').addEventListener('change',aplicarMin);
document.getElementById('nm').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();e.target.blur();aplicarMin();window.scrollTo(0,0);}});
mostrar();
</script>
</body>
</html>
'''


# ---- publicação no GitHub Pages (endereço https fixo, para o atalho no telemóvel)
# A página vai cifrada (AES-256-GCM, chave da palavra-passe por PBKDF2-SHA256): quem não a souber vê só a caixa.
# Configuração (utilizador, repositório, token e palavra-passe) num ficheiro só deste PC, fora do OneDrive.
GH_USER, GH_REPO = "MiguelAlbergaria13", "Painel-acoes"      # o endereço da página distingue maiúsculas
GH_URL = f"https://{GH_USER.lower()}.github.io/{GH_REPO}/"


DONO = "Miguel Albergaria"                      # só este utilizador publica a app do telemóvel (GitHub)
UTILIZADORES = [DONO, "Frederico Carvalho"]
C_UTIL = "N1"                                    # folha Ações: utilizador deste livro (lista)
FICH_UTIL = os.path.join(PASTA, "_Painel_acoes_utilizador.txt")   # utilizador por defeito de um livro novo
_UTIL = {}


def utilizador():
    """Utilizador deste livro: a célula N1 da folha Ações (como está agora no Excel); sem livro, o do ficheiro
    _Painel_acoes_utilizador.txt da pasta; senão, o dono."""
    if "u" in _UTIL:
        return _UTIL["u"]
    u = None
    if os.path.exists(FICHEIRO):
        try:
            v = Excel(FICHEIRO).ops([{"op": "ler", "folha": 1, "rng": C_UTIL}])[0]
            u = txt(v[0][0] if isinstance(v, (list, tuple)) and v and isinstance(v[0], (list, tuple)) else v)
        except Exception:
            try:
                from openpyxl import load_workbook
                wb_ = load_workbook(FICHEIRO, read_only=True)
                u = txt((wb_[FOLHA_A2] if FOLHA_A2 in wb_.sheetnames else wb_.worksheets[0])[C_UTIL].value)
                wb_.close()
            except Exception:
                u = None
    if not u:
        try:
            with open(FICH_UTIL, encoding="utf-8") as f:
                u = f.read().strip()
        except Exception:
            u = None
    _UTIL["u"] = u or DONO
    return _UTIL["u"]


def e_dono():
    return "--servidor" in sys.argv or utilizador().strip().lower() == DONO.lower()


def ficheiro_link_app():
    """Na pasta do painel: o endereço da app do telemóvel (texto e atalho da Internet)."""
    if "--servidor" in sys.argv or not e_dono():
        return
    try:
        txt_l = (f"App do telemóvel (página atualizada no GitHub, com a palavra-passe):\n{GH_URL}\n\n"
                 "No Android: abre este endereço no Chrome > menu ⋮ > Adicionar ao ecrã principal.\n"
                 "Se o atalho antigo abrir outra página, apaga-o e cria-o outra vez a partir deste endereço.\n")
        for nome, conteudo in (("_App_telemovel_link.txt", txt_l),
                               ("_App_telemovel.url", f"[InternetShortcut]\nURL={GH_URL}\n")):
            f = os.path.join(PASTA, nome)
            if not os.path.exists(f) or open(f, encoding="utf-8").read() != conteudo:
                with open(f, "w", encoding="utf-8") as fh:
                    fh.write(conteudo)
    except Exception:
        pass
FICH_GH = os.path.join(os.environ.get("APPDATA") or PASTA, "PainelAcoes", "github.json")
# cópia na pasta do painel (OneDrive): qualquer PC que abra o painel encontra o token sem o pedir
FICH_GH_PASTA = os.path.join(PASTA, BASE + "_github_config.json")


def _gravar_cfg_gh(cfg):
    for f in (FICH_GH, FICH_GH_PASTA):
        try:
            os.makedirs(os.path.dirname(f), exist_ok=True)
            with open(f, "w", encoding="utf-8") as fh:
                json.dump(cfg, fh)
        except Exception as e:
            print(f"   aviso: não consegui guardar {f} ({e}).")
PBKDF2_ITER = 200000
PUB_TPL = r'''<!DOCTYPE html>
<html lang="pt"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#1f3b57"><meta name="mobile-web-app-capable" content="yes">
<title>Painel de ações</title>
<style>
body{margin:0;font-family:system-ui,-apple-system,Roboto,sans-serif;background:#f4f5f7;color:#1c1e21;display:flex;min-height:100vh;align-items:center;justify-content:center}
@media (prefers-color-scheme:dark){body{background:#121417;color:#e8eaed}}
.cx{width:min(320px,90vw);text-align:center}
h1{font-size:20px;font-weight:600}
input{width:100%;box-sizing:border-box;font-size:17px;padding:12px;border-radius:10px;border:1px solid #bbb;margin:10px 0}
button{width:100%;font-size:17px;padding:12px;border:0;border-radius:10px;background:#1f3b57;color:#fff;margin-top:8px}
button.sec{background:transparent;color:inherit;border:1px solid #999}
label{font-size:14px;display:flex;gap:8px;align-items:center;justify-content:center;margin-top:6px}
#e{color:#d93025;font-size:14px;min-height:20px}
.dig{font-size:54px;margin:4px 0}
</style></head><body>
<form id="f" class="cx" style="display:none"><h1>Painel de ações</h1>
<input id="p" type="password" placeholder="Palavra-passe" autocomplete="current-password">
<button>Abrir</button>
<label id="lb" style="display:none"><input id="b" type="checkbox" checked style="width:auto;margin:0"> pedir impressão digital (ou o PIN / padrão do telemóvel) em vez da palavra-passe</label>
<label id="lm"><input id="m" type="checkbox" checked style="width:auto;margin:0"> lembrar neste telemóvel</label>
<p id="e"></p></form>
<div id="d" class="cx" style="display:none"><h1>Painel de ações</h1><div class="dig">🔒</div>
<button id="bd">Desbloquear com impressão digital</button>
<button id="bp" class="sec">Usar a palavra-passe</button><p id="e2" style="color:#d93025;font-size:14px"></p></div>
<div id="a" class="cx" style="display:none"><h1>Painel de ações</h1><div class="dig">🔒</div>
<p>Passar a pedir a impressão digital (ou o PIN / padrão do telemóvel) para abrir o painel?</p>
<button id="ba">Sim, ativar</button><button id="bn" class="sec">Agora não</button></div>
<script>
const DADOS="__DADOS__", ITER=__ITER__;
const LS={get:k=>{try{return localStorage.getItem(k);}catch(e){return null;}},set:(k,v)=>{try{localStorage.setItem(k,v);}catch(e){}},
  del:k=>{try{localStorage.removeItem(k);}catch(e){}}};
const SS={get:k=>{try{return sessionStorage.getItem(k);}catch(e){return null;}},set:(k,v)=>{try{sessionStorage.setItem(k,v);}catch(e){}}};
const B=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0)), S=u=>btoa(String.fromCharCode(...new Uint8Array(u)));
const rnd=n=>crypto.getRandomValues(new Uint8Array(n)), ver=id=>{['f','d','a'].forEach(x=>document.getElementById(x).style.display=x===id?'block':'none');};
async function abrir(pw){
  const b=B(DADOS), sal=b.slice(0,16), iv=b.slice(16,28), ct=b.slice(28);
  const k0=await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveKey']);
  const k=await crypto.subtle.deriveKey({name:'PBKDF2',salt:sal,iterations:ITER,hash:'SHA-256'},k0,{name:'AES-GCM',length:256},false,['decrypt']);
  const html=new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:iv},k,ct));
  SS.set('painel_s',pw);                 // enquanto a app estiver aberta não volta a pedir (p. ex. ao recarregar dados novos)
  document.open();document.write(html);document.close();
}
// impressão digital: WebAuthn (chave do próprio telemóvel). Com a extensão PRF a palavra-passe fica cifrada com uma chave
// que só a impressão digital liberta; sem PRF fica guardada e a impressão digital é só uma porta.
async function salPrf(){return new Uint8Array(await crypto.subtle.digest('SHA-256',new TextEncoder().encode('painel-acoes-prf-v1')));}
async function bioGet(id){
  const a=await navigator.credentials.get({publicKey:{challenge:rnd(32),allowCredentials:[{type:'public-key',id:B(id)}],
    userVerification:'required',timeout:60000,extensions:{prf:{eval:{first:await salPrf()}}}}});
  const r=a.getClientExtensionResults();
  return r.prf&&r.prf.results&&r.prf.results.first?new Uint8Array(r.prf.results.first):null;
}
const chave=(out,uso)=>crypto.subtle.importKey('raw',out.slice(0,32),{name:'AES-GCM'},false,uso);
async function ativarBio(pw){
  const c=await navigator.credentials.create({publicKey:{challenge:rnd(32),rp:{name:'Painel de ações'},
    user:{id:rnd(16),name:'painel',displayName:'Painel de ações'},
    pubKeyCredParams:[{type:'public-key',alg:-7},{type:'public-key',alg:-257}],
    authenticatorSelection:{authenticatorAttachment:'platform',userVerification:'required',residentKey:'preferred'},
    timeout:60000,extensions:{prf:{}}}});
  const id=S(c.rawId), out=await bioGet(id), rec={id};
  if(out){const iv=rnd(12);rec.iv=S(iv);
    rec.ct=S(await crypto.subtle.encrypt({name:'AES-GCM',iv},await chave(out,['encrypt']),new TextEncoder().encode(pw)));LS.del('painel_pw');}
  else {rec.simples=1;LS.set('painel_pw',pw);}
  LS.set('painel_bio',JSON.stringify(rec));
}
async function abrirBio(){
  const rec=JSON.parse(LS.get('painel_bio')), out=await bioGet(rec.id);
  let pw;
  if(rec.ct){if(!out)throw new Error('sem PRF');
    pw=new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:B(rec.iv)},await chave(out,['decrypt']),B(rec.ct)));}
  else pw=LS.get('painel_pw');
  await abrir(pw);
}
function erroBio(e){document.getElementById('e2').textContent=(e&&e.name==='NotAllowedError')?'Cancelado. Toca no botão para tentar outra vez.':
  (e&&e.name==='OperationError'?'Palavra-passe mudou: usa a palavra-passe.':'Não foi possível ('+((e&&e.message)||e)+').');}
document.getElementById('bd').onclick=()=>abrirBio().catch(erroBio);
document.getElementById('bp').onclick=()=>ver('f');
async function temBio(){try{return !!(window.PublicKeyCredential&&await PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable());}catch(e){return false;}}
(async()=>{
  const bio=await temBio();
  if(bio){document.getElementById('lb').style.display='flex';}
  document.getElementById('lm').style.display=bio?'none':'flex';
  const s=SS.get('painel_s');
  if(s){try{await abrir(s);return;}catch(e){}}
  if(LS.get('painel_bio')&&bio){ver('d');abrirBio().catch(erroBio);return;}
  const g=LS.get('painel_pw');
  if(g&&bio&&!LS.get('painel_bio_nao')){                 // já tinha a palavra-passe guardada: propõe a impressão digital
    ver('a');
    document.getElementById('ba').onclick=()=>ativarBio(g).then(()=>abrir(g)).catch(e=>{alert('Não foi possível ativar: '+((e&&e.message)||e));abrir(g);});
    document.getElementById('bn').onclick=()=>{LS.set('painel_bio_nao','1');abrir(g).catch(()=>ver('f'));};
    return;
  }
  if(g){try{await abrir(g);return;}catch(e){LS.del('painel_pw');}}
  ver('f');
})();
document.getElementById('b').onchange=e=>{document.getElementById('lm').style.display=e.target.checked?'none':'flex';};
document.getElementById('f').addEventListener('submit',async ev=>{
  ev.preventDefault();const pw=document.getElementById('p').value;
  try{await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveKey']);
    const usaBio=document.getElementById('lb').style.display!=='none'&&document.getElementById('b').checked;
    // confirma primeiro que a palavra-passe está certa (abrir só no fim, porque substitui a página)
    const b=B(DADOS);const k0=await crypto.subtle.importKey('raw',new TextEncoder().encode(pw),'PBKDF2',false,['deriveKey']);
    const k=await crypto.subtle.deriveKey({name:'PBKDF2',salt:b.slice(0,16),iterations:ITER,hash:'SHA-256'},k0,{name:'AES-GCM',length:256},false,['decrypt']);
    await crypto.subtle.decrypt({name:'AES-GCM',iv:b.slice(16,28)},k,b.slice(28));
    if(usaBio){try{await ativarBio(pw);}catch(e){alert('Impressão digital não ativada: '+((e&&e.message)||e));}}
    else if(document.getElementById('m').checked){LS.set('painel_pw',pw);}
    await abrir(pw);}
  catch(e){document.getElementById('e').textContent='Palavra-passe errada.';}
});
</script></body></html>
'''


def _ler_cfg_gh():
    """Configuração do GitHub: a deste PC ou, se não houver, a da pasta do painel (OneDrive)."""
    for fich in (FICH_GH, FICH_GH_PASTA):
        try:
            with open(fich, encoding="utf-8") as f:
                cfg = json.load(f)
            if cfg.get("token") or cfg.get("desligado"):
                return cfg
        except Exception:
            pass
    return {}


def _config_github():
    """Lê (ou pergunta, na 1.ª vez) o token e a palavra-passe. None se a publicação estiver desligada."""
    cfg = _ler_cfg_gh()
    if cfg.get("token") and cfg.get("senha"):           # garante as duas cópias (este PC e a pasta do OneDrive)
        for fich in (FICH_GH_PASTA, FICH_GH):
            if not os.path.exists(fich):
                try:
                    os.makedirs(os.path.dirname(fich), exist_ok=True)
                    with open(fich, "w", encoding="utf-8") as f:
                        json.dump(cfg, f)
                    if fich == FICH_GH_PASTA:
                        print(f"   Token guardado também na pasta do painel ({os.path.basename(fich)}): os outros PCs já "
                              "não o pedem.")
                except Exception:
                    pass
    # alternativa à janela preta: ficheiro _Painel_github.txt na pasta do painel, com o token na 1.ª linha e a
    # palavra-passe na 2.ª; o script guarda-os fora do OneDrive e apaga o ficheiro
    ftxt = os.path.join(PASTA, BASE + "_github.txt")
    if os.path.exists(ftxt):
        try:
            with open(ftxt, encoding="utf-8-sig") as f:
                ls = [x.strip() for x in f.read().splitlines() if x.strip()]
            if len(ls) >= 2 and len(ls[1]) >= 8:
                cfg = {"user": GH_USER, "repo": GH_REPO, "token": ls[0], "senha": ls[1]}
                if len(ls) >= 3:
                    cfg["token_botao"] = ls[2]
                _gravar_cfg_gh(cfg)
                os.remove(ftxt)
                print(f"Token e palavra-passe lidos de {os.path.basename(ftxt)} (guardados neste PC e na pasta do painel, "
                      f"em {os.path.basename(FICH_GH_PASTA)}; o ficheiro de texto foi apagado).")
                return cfg
            print(f"   aviso: {os.path.basename(ftxt)} precisa do token na 1.ª linha e da palavra-passe "
                  "(mín. 8 caracteres) na 2.ª.")
        except Exception as e:
            print(f"   aviso: não consegui ler {os.path.basename(ftxt)} ({e}).")
    if cfg.get("desligado"):
        return None
    if cfg.get("token") and cfg.get("senha"):
        return cfg
    # sem configuração neste PC: não pára a atualização a pedir o token; só avisa
    print(f"\n   AVISO: este PC não tem o token do GitHub guardado ({FICH_GH}), por isso a app do telemóvel não foi "
          "atualizada desta vez.")
    print(f"   Para resolver: cria na pasta do painel o ficheiro {BASE}_github.txt com 3 linhas (token de Contents, "
          "palavra-passe da app, token do botão ↻); no ▶ ATUALIZAR seguinte o script guarda-os e apaga o ficheiro.")
    return None


def _cifrar(texto, senha):
    import base64, hashlib
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    except ImportError:
        print("   A instalar o módulo de cifra (cryptography), só na 1.ª vez...")
        subprocess.run([sys.executable, "-m", "pip", "install", "--user", "-q", "cryptography"], timeout=300,
                       **SEM_JANELA)
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    sal, iv = os.urandom(16), os.urandom(12)
    chave = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), sal, PBKDF2_ITER, 32)
    return base64.b64encode(sal + iv + AESGCM(chave).encrypt(iv, texto.encode("utf-8"), None)).decode()


def _decifrar(b64, senha):
    import base64, hashlib
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    b = base64.b64decode(b64)
    chave = hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), b[:16], PBKDF2_ITER, 32)
    return AESGCM(chave).decrypt(b[16:28], b[28:], None).decode("utf-8")


def _json_num(o):
    try:
        return float(o)
    except Exception:
        return str(o)


def pagina_cifrada(html, senha):
    return PUB_TPL.replace("__DADOS__", _cifrar(html, senha)).replace("__ITER__", str(PBKDF2_ITER))


GH_RAMO_PAGINA = "gh-pages"      # a página (e os dados das atualizações automáticas) vivem neste ramo


def _api_github(cfg, caminho, corpo=None, metodo=None):
    import urllib.request
    url = f"https://api.github.com/repos/{cfg.get('user', GH_USER)}/{cfg.get('repo', GH_REPO)}/{caminho}"
    cab = {"Authorization": f"Bearer {cfg['token']}", "Accept": "application/vnd.github+json",
           "User-Agent": "painel-acoes", "X-GitHub-Api-Version": "2022-11-28"}
    req = urllib.request.Request(url, data=json.dumps(corpo).encode() if corpo is not None else None, headers=cab,
                                 method=metodo or ("POST" if corpo is not None else "GET"))
    with urllib.request.urlopen(req, timeout=90) as r:
        t = r.read()
        return json.loads(t) if t else {}


def _garantir_ramo(cfg, ramo):
    """Cria o ramo (a partir do main) se ainda não existir."""
    import urllib.error
    try:
        _api_github(cfg, f"git/ref/heads/{ramo}")
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
        sha = _api_github(cfg, "git/ref/heads/main")["object"]["sha"]
        _api_github(cfg, "git/refs", {"ref": f"refs/heads/{ramo}", "sha": sha})


def _enviar_github(cfg, nome, conteudo, ramo="main"):
    """Cria ou substitui um ficheiro no repositório (API de conteúdos), no ramo indicado."""
    import base64, urllib.request, urllib.error
    if ramo != "main":
        _garantir_ramo(cfg, ramo)
    api = f"https://api.github.com/repos/{cfg.get('user', GH_USER)}/{cfg.get('repo', GH_REPO)}/contents/{nome}"
    cab = {"Authorization": f"Bearer {cfg['token']}", "Accept": "application/vnd.github+json",
           "User-Agent": "painel-acoes", "X-GitHub-Api-Version": "2022-11-28"}
    sha = None
    try:
        with urllib.request.urlopen(urllib.request.Request(api + f"?ref={ramo}", headers=cab), timeout=30) as r:
            sha = json.load(r).get("sha")
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    corpo = {"message": f"{nome} {datetime.now():%d/%m/%Y %H:%M}", "branch": ramo,
             "content": base64.b64encode(conteudo.encode("utf-8")).decode()}
    if sha:
        corpo["sha"] = sha
    req = urllib.request.Request(api, data=json.dumps(corpo).encode(), headers=cab, method="PUT")
    with urllib.request.urlopen(req, timeout=90):
        pass


WORKFLOW_YML = 'name: painel\n\non:\n  schedule:\n    # hora UTC (= Lisboa no inverno; no verão é mais 1 hora em Lisboa)\n    - cron: \'30 7 * * 1-5\'            # completa antes da abertura europeia\n    - cron: \'0 14 * * 1-5\'            # completa antes da abertura americana\n    - cron: \'15 22 * * 1-5\'           # completa depois do fecho americano\n    - cron: \'5,20,35,50 8-21 * * 1-5\' # só preços, de 15 em 15 minutos\n  workflow_dispatch:\n    inputs:\n      tipo:\n        description: \'completa ou precos\'\n        required: false\n        default: \'precos\'\n\npermissions:\n  contents: write\n  pages: write\n\nconcurrency:\n  group: painel\n  cancel-in-progress: false\n\njobs:\n  painel:\n    runs-on: ubuntu-latest\n    timeout-minutes: 50\n    steps:\n      - uses: actions/checkout@v4\n\n      - name: Recuperar dados e caches da última atualização (ramo gh-pages)\n        run: |\n          if git fetch -q origin gh-pages 2>/dev/null; then\n            for f in dados.enc _Painel_acoes_bolsas.json _Painel_acoes_composicao_etfs.json _Painel_acoes_acoes_dos_etfs.json; do\n              git show "FETCH_HEAD:$f" > "$f" 2>/dev/null || rm -f "$f"\n            done\n          fi\n\n      - uses: actions/setup-python@v5\n        with:\n          python-version: \'3.12\'\n\n      - name: Instalar\n        run: pip install -q yfinance openpyxl cryptography pandas numpy\n\n      - name: Atualizar\n        env:\n          PAINEL_SENHA: ${{ secrets.PAINEL_SENHA }}\n        run: |\n          TIPO="${{ github.event.inputs.tipo }}"\n          if [ -z "$TIPO" ]; then\n            case "${{ github.event.schedule }}" in\n              "5,20,35,50 8-21 * * 1-5") TIPO=precos ;;\n              *) TIPO=completa ;;\n            esac\n          fi\n          python _Painel_acoes.py --servidor "$TIPO"\n\n      - name: Publicar no ramo gh-pages (sem histórico, para o repositório não crescer)\n        run: |\n          mkdir -p /tmp/pub\n          for f in index.html versao.txt dados.enc _Painel_acoes_bolsas.json _Painel_acoes_composicao_etfs.json _Painel_acoes_acoes_dos_etfs.json; do\n            [ -f "$f" ] && cp "$f" /tmp/pub/\n          done\n          touch /tmp/pub/.nojekyll\n          cd /tmp/pub\n          git init -q -b gh-pages\n          git config user.name "painel-bot"\n          git config user.email "painel-bot@users.noreply.github.com"\n          git add -A\n          git commit -q -m "painel $(date -u +\'%d/%m/%Y %H:%M\') UTC"\n          git push -q -f "https://x-access-token:${{ secrets.GITHUB_TOKEN }}@github.com/${{ github.repository }}.git" gh-pages:gh-pages\n\n      - name: Pedir a publicação da página\n        run: |\n          curl -s -o /dev/null -w "pages build: %{http_code}\\n" -X POST \\\n            -H "Authorization: Bearer ${{ secrets.GITHUB_TOKEN }}" -H "Accept: application/vnd.github+json" \\\n            "https://api.github.com/repos/${{ github.repository }}/pages/builds"\n'


def _existe_github(cfg, nome, ramo="main"):
    import urllib.request, urllib.error
    api = f"https://api.github.com/repos/{cfg.get('user', GH_USER)}/{cfg.get('repo', GH_REPO)}/contents/{nome}?ref={ramo}"
    cab = {"Authorization": f"Bearer {cfg['token']}", "Accept": "application/vnd.github+json", "User-Agent": "painel-acoes"}
    try:
        with urllib.request.urlopen(urllib.request.Request(api, headers=cab), timeout=30):
            return True
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return False
        raise


def _garantir_workflow(cfg):
    """As atualizações automáticas e o botão ↻ da app precisam de .github/workflows/painel.yml no repositório."""
    import urllib.error
    nome = ".github/workflows/painel.yml"
    try:
        if _existe_github(cfg, nome):
            return
        _enviar_github(cfg, nome, WORKFLOW_YML)
        print("   Atualizações automáticas instaladas no GitHub (.github/workflows/painel.yml).")
    except urllib.error.HTTPError:
        print("   AVISO: o repositório não tem .github/workflows/painel.yml, por isso as atualizações automáticas e o botão ↻ "
              "da app não funcionam. O token não pode criá-lo (precisa da permissão Workflows). Cria-o à mão: no GitHub, "
              "Add file > Create new file, nome .github/workflows/painel.yml, e cola o conteúdo de github\\painel.yml do zip.")
    except Exception:
        pass


def _compra_sync():
    """Escolhas da folha Compra (sem a tabela) para enviar ao GitHub, que refaz a app sem o Excel."""
    c = _ler_json(FICH_COMPRA, {})
    return {k: v for k, v in c.items() if k != "linhas2"} or None


def enviar_compra_github():
    """Depois do ▶ CALCULAR COMPRA: envia logo as escolhas da folha Compra (nas definições cifradas), para a app as
    receber na atualização automática seguinte (até 15 minutos), sem esperar pelo próximo ▶ ATUALIZAR."""
    cfg = _config_github()
    if not cfg or not os.path.exists(FICH_SYNC):
        return
    try:
        with open(FICH_SYNC, encoding="utf-8") as f:
            sync = json.load(f)
        defs = dict(sync, token_botao=cfg.get("token_botao"), enviado=f"{datetime.now():%d/%m/%Y %H:%M}",
                    compra=_compra_sync())
        _enviar_github(cfg, "config.enc", _cifrar(json.dumps(defs, ensure_ascii=False, default=_json_num), cfg["senha"]))
        print("Escolhas da compra enviadas para a app (aparecem na próxima atualização automática, até 15 minutos).")
    except Exception as e:
        print("   aviso: escolhas da compra não enviadas para a app:", e)


def publicar_github(html, versao="", dados=None):
    """Envia a página (cifrada) para o repositório; o GitHub Pages mostra-a em GH_URL ao fim de ~1 minuto. Depois de
    um ▶ ATUALIZAR envia também as definições (cifradas) e o script, para as atualizações automáticas no GitHub."""
    cfg = _config_github()
    if not cfg:
        return
    import urllib.error
    try:
        pag = pagina_cifrada(html, cfg["senha"])
        _enviar_github(cfg, "index.html", pag, GH_RAMO_PAGINA)       # gh-pages (o das atualizações automáticas)
        _enviar_github(cfg, "versao.txt", versao, GH_RAMO_PAGINA)
        _enviar_github(cfg, "index.html", pag)                       # também no main (se o Pages publicar o main)
        _enviar_github(cfg, "versao.txt", versao)
        _enviar_github(cfg, "ocultos.enc", _cifrar(json.dumps(sorted(OCULTOS)), cfg["senha"]))
        if dados:                                # os dados do PC passam a ser a base das atualizações de preços no
            try:                                 # GitHub (senão elas usavam a última completa do servidor)
                dd = {"a": dados[0], "e": dados[1], "quando": f"{datetime.now():%d/%m/%Y %H:%M}"}
                _enviar_github(cfg, "dados.enc", _cifrar(json.dumps(
                    dd, ensure_ascii=False, default=lambda o: _json_num(o) if isinstance(o, (int, float)) or
                    type(o).__module__ == "numpy" else str(o)), cfg["senha"]), GH_RAMO_PAGINA)
            except Exception as e:
                print("   aviso: dados para as atualizações automáticas não enviados:", e)
        with open(os.path.abspath(__file__), encoding="utf-8") as f:   # script sempre (o GitHub corre esta versão)
            _enviar_github(cfg, "_Painel_acoes.py", f.read())
        if SYNC:
            defs = dict(SYNC, token_botao=cfg.get("token_botao"), enviado=f"{datetime.now():%d/%m/%Y %H:%M}",
                        compra=_compra_sync())       # escolhas da folha Compra (o separador Compra da app usa-as)
            _enviar_github(cfg, "config.enc", _cifrar(json.dumps(defs, ensure_ascii=False, default=_json_num), cfg["senha"]))
            try:                                     # cópia local (o ↻ PREÇOS atualiza nela os detidos e volta a enviá-la)
                with open(FICH_SYNC, "w", encoding="utf-8") as f:
                    json.dump(SYNC, f, ensure_ascii=False, default=_json_num)
            except Exception:
                pass
            _garantir_workflow(cfg)
        print(f"Página publicada: {GH_URL}  (fica visível ao fim de 1-2 minutos)"
              + ("; definições enviadas para as atualizações automáticas" if SYNC else ""))
        registo("publicada no GitHub")
    except urllib.error.HTTPError as e:
        msg = {401: "token inválido ou expirado", 403: "o token não tem permissão de escrita neste repositório",
               404: "repositório não encontrado (nome ou acesso do token)"}.get(e.code, f"erro {e.code}")
        print(f"   aviso: não publiquei no GitHub ({msg}). Para mudar o token, cria {BASE}_github.txt na pasta do painel (3 linhas: token, palavra-passe, token do botão) e atualiza outra vez.")
        registo(f"GitHub: {msg}")
    except Exception as e:
        print(f"   aviso: não publiquei no GitHub ({e}).")
        registo(f"GitHub: {e}")


BOTAO = None          # {"o": utilizador, "r": repositório, "t": token só de Actions} -> botão ↻ na página
FICH_SYNC = os.path.join(PASTA, BASE + "_definicoes_app.json")


def sync_dos_precos(linhas, etfs):
    if not e_dono():
        return
    return _sync_dos_precos(linhas, etfs)


def _sync_dos_precos(linhas, etfs):
    """↻ PREÇOS: as definições do último ▶ ATUALIZAR com os detidos (x, nº de ações, preço de compra) como estão agora
    nas folhas, para as atualizações automáticas não voltarem a valores antigos."""
    try:
        with open(FICH_SYNC, encoding="utf-8") as f:
            base = json.load(f)
    except Exception:
        return
    det, qtd, pc, co = set(), {}, {}, {}
    for d in list(linhas) + list(etfs):
        ks = [k for k in (txt(d.get("nome")), txt(d.get("ticker")).upper()) if k]
        q = d.get("qtd") if isinstance(d.get("qtd"), (int, float)) else qtd_detida(d.get("det"))
        if e_detido(d.get("det")) or q:
            det.update(ks)
        if isinstance(q, (int, float)) and q > 0:
            qtd.update({k: q for k in ks})
        if isinstance(d.get("pcompra"), (int, float)):
            pc.update({k: d["pcompra"] for k in ks})
        if txt(d.get("corr")):
            co.update({k: txt(d["corr"]) for k in ks})
    SYNC.clear()
    SYNC.update(base, det=sorted(det), qtd=qtd, pcompra=pc, corr=co, pesos=dict(PESOS))


SYNC = {}             # definições do PC para o servidor (acrescentados, x, modo, longo prazo mínimo, ações dos ETFs)


def valor_detido(d):
    """Valor em euros do que tens (nº de ações × preço, convertido), se souber o nº de ações."""
    carteira(d)
    return d.get("valor_eur")


def pagina_html(linhas, etfs, texto):
    """(html, versão) da página para o telemóvel."""
    def num(v):
        return float(v) if isinstance(v, (int, float)) and v == v else None

    def item(d):
        return {"n": txt(d.get("nome")), "c": txt(d.get("codigo")), "t": txt(d.get("ticker")), "m": txt(d.get("moeda")),
                "p": num(d.get("preco")), "nt": num(d.get("nota")), "ng": num(d.get("nota_geral")), "dc": txt(d.get("decisao")), "pt": num(d.get("pont")),
                "gn": num(d.get("g_nota")), "g5": num(d.get("g_n5")),
                "np": num(d.get("nivel_preco")), "be": num(d.get("beta")), "pr": num(d.get("pe")), "ca": 1 if d.get("so_caro") == 1 else 0,
                "nr": rotulo_nivel(d.get("nivel_preco")) or None,
                "nc": "#" + COR_NIVEL[rotulo_nivel(d.get("nivel_preco"))] if rotulo_nivel(d.get("nivel_preco")) else None,
                "x": e_detido(d.get("det")), "vd": valor_detido(d), "qn": num(d.get("qtd")) or qtd_detida(d.get("det")),
                "br": txt(d.get("corr")) or None, "pc": num(d.get("pcompra")), "le": num(d.get("lp_eur")), "lpp": num(d.get("lp_pct")),
                "ce": ((d["valor_eur"] - d["lp_eur"]) if isinstance(d.get("valor_eur"), (int, float))
                       and isinstance(d.get("lp_eur"), (int, float)) else None),
                "ds": 1 if d.get("desat") == 1 else 0,
                "tm": num(d.get("teto")), "tmp": num(d.get("teto_pct")),
                "fx": (1.0 if txt(d.get("moeda")) == "EUR" else num(d.get("fx_alt")) if txt(d.get("moeda_alt")) == "EUR"
                       else None),
                "cr": num(d.get("compra")), "crp": num(d.get("compra_pct")), "va": num(d.get("venda_alvo")),
                "vap": num(d.get("venda_alvo_pct")), "vm": num(d.get("venda_min")), "vmp": num(d.get("venda_min_pct")), "vmc": num(d.get("venda_min_compra")),
                "q": num(d.get("qual")), "s": num(d.get("saude_n")), "sd": txt(d.get("saude")) or None,
                "sg": txt(d.get("sinal")) or None, "lp": txt(d.get("lp_txt")) or None, "cm": txt(d.get("coment")) or None,
                "no": txt(d.get("notas")) or None, "ed": txt(d.get("etf_desc")) or None,
                "ec": txt(d.get("etf_comp")) or None, "ea": num(d.get("etf_acoes")),
                "or": {k: num(d.get(k)) for k in ("pot", "justo_pct", "var1d", "var1m", "val_sempre", "val10", "val5",
                                                  "val3", "val1", "dist0") if num(d.get(k)) is not None}}
    versao = datetime.now().strftime("%Y%m%d%H%M%S")
    vis = lambda d: txt(d.get("ticker")) and txt(d.get("ticker")).upper() not in OCULTOS
    n_oc = sum(1 for d in list(linhas) + list(etfs) if txt(d.get("ticker")) and not vis(d))
    dados = {"at": f"{texto} · {TXT_VERSAO}" + (f" · {n_oc} escondidos pelo filtro do Excel" if n_oc else ""),
             "a": [item(d) for d in linhas if vis(d)], "e": [item(d) for d in etfs if vis(d)], "v": versao, "gh": BOTAO}
    try:                                       # escolhas da folha Compra (valores por omissão do separador Compra)
        cp_ = _ler_json(FICH_COMPRA, {})
        if cp_:
            dados["cmp"] = {"a": cp_.get("min_a"), "e": cp_.get("min_e"), "t": cp_.get("total"),
                            "p": round(100 * cp_["pct_a"]) if isinstance(cp_.get("pct_a"), (int, float)) else None,
                            "s": 1 if cp_.get("so_compra") else 0, "h": 1 if cp_.get("detidos") else 0,
                            "nq": cp_.get("excluir") or [], "nd": cp_.get("indisp") or [],
                            "fa": cp_.get("fixos_a") or [], "fe": cp_.get("fixos_e") or [],
                            "ts": cp_.get("ts") or int(os.path.getmtime(FICH_COMPRA) * 1000)}
            dados["cmp"] = {k: v for k, v in dados["cmp"].items() if v is not None}
    except Exception:
        pass
    js = json.dumps(dados, ensure_ascii=False).replace("</", "<\\/")
    return HTML_TPL.replace("__DATA__", js), versao


def gerar_html(linhas, etfs, texto):
    """Página para o telemóvel com nomes, notas, decisão e níveis de compra e venda (a situação do momento)."""
    global BOTAO
    if not e_dono():                             # outro utilizador: o livro é só dele; nada vai para o telemóvel
        print(f"Utilizador: {utilizador()} (sem app do telemóvel).")
        return
    try:
        cfg = _ler_cfg_gh()
        if cfg.get("token_botao"):
            BOTAO = {"o": cfg.get("user", GH_USER), "r": cfg.get("repo", GH_REPO), "t": cfg["token_botao"]}
        html, versao = pagina_html(linhas, etfs, texto)
        with open(FICH_HTML, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Página para o telemóvel: {os.path.basename(FICH_HTML)}")
        publicar_github(html, versao, (linhas, etfs))
    except Exception as e:
        print("   aviso: página para o telemóvel não gravada:", e)
        registo(f"telemóvel: {e}")


# ================================================================================================ Procurar / Acrescentados
L_TAB = 12            # linha do cabeçalho da tabela das folhas Procurar
C_ESTADO = "C9"       # estado das folhas Procurar
# Folha 3 = Procurar ações (por nome/ticker, ou as mais significativas de uma região); folha 4 = Procurar ETFs
FOLHA_PROC, FOLHA_PROC_ETF = 3, 4
NOME_PROC = {FOLHA_PROC: "Proc ações", FOLHA_PROC_ETF: "Proc ETFs"}
FOLHA_A2, FOLHA_E2 = "Ações", "ETFs"             # V190: outra vez as folhas Ações / ETFs (com a nota por ganhos)
NOME_IDX = {1: "Ações(2)", 2: "ETFs(2)"}          # livros da V179-V189: as folhas base tinham estes nomes; o código
                                                  # usa 1 e 2 e o Excel.ops traduz enquanto o livro não for refeito
NOME_PROC_ANT = {FOLHA_PROC: "Procurar ações", FOLHA_PROC_ETF: "Procurar ETFs"}     # nomes até à V122
P_TXT, P_CRIT, P_REG, P_N = "C3", "C4", "C5", "C6"
MAXP = 500            # nº máximo de resultados
CAB_P = ["x", "Empresa / ETF", "Tipo", "Código", "Ticker Yahoo", "Bolsa", "Capitalização (mil M USD)",
         "Moeda", "Já está em", "Pontuação / outras bolsas da mesma empresa ou fundo (não se copiam)"]
NC_P = len(CAB_P)
P_TAB = f"A{L_TAB + 1}:{get_column_letter(NC_P)}{L_TAB + MAXP}"
CRITERIOS = ["Maior capitalização", "Mais transacionadas (valor médio/dia 3 meses)",
             "Melhor pontuação (tendência + analistas)", "Maior potencial dos analistas", "Mais subiram em 12 meses",
             "Baratas e a crescer (P/E até 15 e lucros +10%)"] + [f"Maior valorização {h} (% por ano)" for h in
             ("desde sempre", "em 10 anos", "em 5 anos", "em 3 anos", "em 1 ano")]
# horizontes da valorização (iguais às colunas Valorização do painel): critério -> (meses da reta, mínimo de meses)
HORIZ_VAL = {"desde sempre": (None, 36), "em 10 anos": (121, 121), "em 5 anos": (61, 61), "em 3 anos": (37, 37),
             "em 1 ano": (13, 13)}


def _horiz_val(crit):
    m = re.match(r"Maior valorização (.+) \(% por ano\)$", crit or "")
    return (m.group(1),) + HORIZ_VAL[m.group(1)] if m and m.group(1) in HORIZ_VAL else None


def _reta_anual(c):
    """Subida média por ano da reta que melhor acompanha os fechos mensais c (a mesma das colunas Valorização)."""
    import numpy as np
    x = np.arange(len(c)) / 12.0
    return float(np.exp(np.polyfit(x, np.log(c.to_numpy(dtype=float)), 1)[0]) - 1)


def _valorizacoes(yf, tickers, meses, minimo):
    """{ticker: subida média por ano} na janela pedida (meses=None: desde sempre), com fechos mensais e dividendos."""
    out, tickers = {}, [t for t in dict.fromkeys(tickers) if t]
    for i in range(0, len(tickers), 200):
        lote = tickers[i:i + 200]
        print(f"   história mensal: {i + len(lote)} de {len(tickers)}...")
        try:
            h = _por_ticker(yf.download(lote, period="max", interval="1mo", group_by="ticker", auto_adjust=True,
                                        threads=True, progress=False), lote)
        except Exception as e:
            print("   aviso:", e)
            continue
        for t in lote:
            try:
                c = h[t]["Close"].dropna()
                c = c[c > 0]
                if len(c) < minimo:
                    continue
                out[t] = (_reta_anual(c if meses is None else c.iloc[-meses:]), int(c.index[0].year))
            except Exception:
                continue
    return out
CRITERIOS_ETF = ["Maior dimensão (património do fundo)", "Mais transacionados (valor médio/dia)",
                 "Melhor pontuação (subida 6 meses + médias)", "Mais subiram em 12 meses"] + CRITERIOS[6:]
# ETFs UCITS (compráveis na Europa) mais conhecidos por categoria: o Yahoo não tem filtro de ETFs, por isso o ★ parte
# desta lista e ordena-a com os dados do Yahoo (património e negócio); os que o Yahoo não reconhecer ficam de fora.
ETFS_CANDIDATOS = {
    "Mundo": ["IWDA.AS", "VWCE.DE", "VWRL.AS", "XDWD.DE", "SSAC.L", "IUSN.DE", "EUNL.DE"],
    "EUA": ["CSPX.AS", "SXR8.DE", "VUAA.DE", "VUSA.AS", "SXRV.DE", "CNDX.AS", "QDVE.DE", "ZPRV.DE"],
    "Europa": ["EXSA.DE", "IMAE.AS", "VEUR.AS", "MEUD.PA", "EXW1.DE", "EXS1.DE", "CAC.PA", "IMIB.MI", "ISF.L"],
    "Emergentes e Ásia": ["EMIM.AS", "IS3N.DE", "VFEM.AS", "XMME.DE", "IJPA.AS", "VJPN.AS"],
    "Obrigações": ["IEGA.AS", "IEAC.AS", "EUNA.DE", "XGLE.DE", "IDTL.L", "VAGF.DE"],
    "Ouro": ["4GLD.DE", "SGLN.L", "IGLN.L", "EWG2.DE", "SGLD.L"],
    "Setores e temas": ["VVSM.DE", "QDVE.DE", "INRG.L", "2B76.DE", "IUHC.L"],
    "Dividendos": ["VHYL.AS", "TDIV.AS", "ISPA.DE"],
}
CATEGORIAS_ETF = ["Todas"] + list(ETFS_CANDIDATOS)
_EUROPA = ["gb", "de", "fr", "nl", "ch", "es", "it", "se", "dk", "no", "fi", "be", "pt", "ie", "at"]
_ASIA = ["jp", "cn", "hk", "tw", "kr", "sg", "au", "nz", "in"]
_EMERG = ["cn", "tw", "kr", "in", "br", "mx", "za", "sa", "ae", "tr", "pl", "gr", "cl", "id", "th", "my"]
REGIOES = {  # principais mercados primeiro
    "EUA": ["us"], "Europa": _EUROPA, "EUA + Europa": ["us"] + _EUROPA, "Japão": ["jp"], "China": ["cn", "hk"],
    "Ásia-Pacífico": _ASIA, "Emergentes": _EMERG, "Mundo (principais)": ["us"] + _EUROPA + ["jp", "ca", "au", "hk"],
    # países
    "Portugal": ["pt"], "Reino Unido": ["gb"], "Alemanha": ["de"], "França": ["fr"], "Países Baixos": ["nl"],
    "Espanha": ["es"], "Itália": ["it"], "Suíça": ["ch"], "Bélgica": ["be"], "Irlanda": ["ie"], "Áustria": ["at"],
    "Nórdicos": ["se", "dk", "no", "fi"], "Suécia": ["se"], "Dinamarca": ["dk"], "Noruega": ["no"], "Finlândia": ["fi"],
    "Polónia": ["pl"], "Grécia": ["gr"], "Canadá": ["ca"], "Brasil": ["br"], "México": ["mx"], "Argentina": ["ar"],
    "Chile": ["cl"], "Hong Kong": ["hk"], "Taiwan": ["tw"], "Coreia do Sul": ["kr"], "Índia": ["in"],
    "Singapura": ["sg"], "Austrália": ["au"], "Nova Zelândia": ["nz"], "Israel": ["il"], "África do Sul": ["za"],
    "Arábia Saudita": ["sa"], "Emirados": ["ae"], "Turquia": ["tr"], "Indonésia": ["id"], "Tailândia": ["th"],
    "Malásia": ["my"]}
FOLHA_ACR = "Acrescentados"     # pelo nome (V118: a folha Compra mudou a posição das seguintes)
A_TAB = f"A{AC_L0 + 1}:E{AC_L0 + MAX_LINHAS}"   # Empresa / ETF, Código, Ticker Yahoo, Bolsa, Tipo
CAB_A = ["Empresa / ETF", "Código", "Ticker Yahoo", "Bolsa", "Tipo"]


def folha_acrescentados(wb, acrescentados):
    from openpyxl.worksheet.datavalidation import DataValidation
    ws = wb.create_sheet("Acrescentados")
    ws["A1"] = "Títulos acrescentados ao painel (além das listas do script)"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"] = ("Vêm das folhas Procurar ações / Procurar ETFs (➕ ACRESCENTAR MARCADOS) ou escreve-os aqui à mão (o ticker "
                "tem de ser o do Yahoo). Para tirar um título, apaga a linha: deixa de ser procurado no ▶ ATUALIZAR seguinte.")
    ws["A2"].font = Font(italic=True, size=9, color="595959")
    for j, (h, w) in enumerate(zip(CAB_A, (32, 10, 16, 20, 8)), 1):
        c = ws.cell(AC_L0, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(j)].width = w
    for i, linha in enumerate(acrescentados, start=AC_L0 + 1):
        for j, v in enumerate(linha, 1):
            ws.cell(i, j, v)
    dv = DataValidation(type="list", formula1='"Ação,ETF"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"E{AC_L0 + 1}:E{AC_L0 + MAX_LINHAS}")
    ws.freeze_panes = f"A{AC_L0 + 1}"


def ler_acrescentados(ex, so_ficheiro=False):
    """Linhas (nome, código, ticker, bolsa, tipo) da folha Acrescentados (linhas vazias pelo meio são saltadas). Procura a
    folha pelo cabeçalho (a 5.ª nesta versão, a 4.ª nas anteriores)."""
    ler = (lambda ops: _ops_ficheiro(ex.caminho, ops, False)) if so_ficheiro else ex.ops
    for folha_n in (FOLHA_ACR, 6, 5, 4):
        try:
            cab = ler([{"op": "ler", "folha": folha_n, "rng": f"A{AC_L0}:E{AC_L0}"}])[0][0]
        except SystemExit:
            raise
        except Exception:
            continue
        if [txt(v) for v in cab] == CAB_A:
            break
    else:
        raise RuntimeError("folha Acrescentados não encontrada")
    out = []
    for r in ler([{"op": "ler", "folha": folha_n, "rng": A_TAB}])[0]:
        r = (list(r) + [None] * 5)[:5]
        t = txt(r[2])
        if t:
            tipo = "ETF" if txt(r[4]).upper() == "ETF" else "Ação"
            out.append((txt(r[0]) or t, txt(r[1]) or t.split(".")[0], t, txt(r[3]), tipo))
    return out


def folha_procurar(wb, pf):
    from openpyxl.styles import Alignment, Font, PatternFill
    borda = Border(*(Side(style="thin", color="A6A6A6"),) * 4)
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    from openpyxl.worksheet.datavalidation import DataValidation
    etf = pf == FOLHA_PROC_ETF
    qual = "etfs" if etf else "acoes"
    ws = wb.create_sheet(NOME_PROC[pf], pf - 1)
    ws["A1"] = ("Procurar ETFs no Yahoo Finance (qualquer bolsa) e acrescentá-los ao painel" if etf else
                "Procurar ações no Yahoo Finance (qualquer bolsa) e acrescentá-las ao painel")
    ws["A1"].font = Font(size=14, bold=True)
    campos = ((3, "Procurar (nome ou ticker):"), (4, "Ou os mais significativos:"), (5, "Categoria:"),
              (6, "Quantos:")) if etf else (
        (3, "Procurar (nome ou ticker):"), (4, "Ou as mais significativas:"), (5, "Região:"), (6, "Quantas:"))
    for linha, t in campos:
        c = ws.cell(linha, 2, t)
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="right")
        ws.merge_cells(f"C{linha}:H{linha}")
        for col in "CDEFGH":
            ws[f"{col}{linha}"].fill = amarelo
            ws[f"{col}{linha}"].border = borda
    ws[P_TXT] = None
    if True:
        crits, regs = (CRITERIOS_ETF, CATEGORIAS_ETF) if etf else (CRITERIOS, list(REGIOES))
        ws[P_CRIT], ws[P_REG], ws[P_N] = crits[0], ("Todas" if etf else "EUA + Europa"), (30 if etf else 100)
        ws[P_N].alignment = Alignment(horizontal="left")
        def lista(itens, col):                  # lista longa (> 255 caracteres): numa coluna escondida da folha
            if len(",".join(itens)) < 250:
                return '"' + ",".join(itens) + '"'
            for i, v in enumerate(itens, 1):
                ws[f"{col}{i}"] = v
            ws.column_dimensions[col].hidden = True
            return f"${col}$1:${col}${len(itens)}"
        for dv, cel in ((DataValidation(type="list", formula1=lista(crits, "AY")), P_CRIT),
                        (DataValidation(type="list", formula1=lista(regs, "AZ")), P_REG),
                        (DataValidation(type="whole", operator="between", formula1="1", formula2=str(MAXP)), P_N)):
            dv.error, dv.errorTitle, dv.allow_blank = ("Escolhe um valor da lista (Quantas: 1 a %d)." % MAXP,
                                                      "Valor inválido", False)
            ws.add_data_validation(dv)
            dv.add(cel)
    botao(ws, "B8", "▶ PROCURAR", f"{PROTOCOLO}://procurar/{qual}", "2E75B6", "1F4E79")
    ws.merge_cells("D8:F8")
    botao(ws, "D8", "★ MAIS SIGNIFICATIVOS" if etf else "★ MAIS SIGNIFICATIVAS", f"{PROTOCOLO}://top/{qual}",
          "7030A0", "4B2070")
    ws.merge_cells("G8:I8")
    botao(ws, "G8", "➕ ACRESCENTAR MARCADOS", f"{PROTOCOLO}://acrescentar/{qual}", "00A050", "006E37")
    botao(ws, "J8", "↺ ADICIONAR TÍTULOS DA LISTA BASE (exceto se já estão nessa ou noutra moeda)",
          f"{PROTOCOLO}://base/{qual}", "BF8F00", "7F6000", 9)
    ws["J8"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[8].height = 30
    ws["B9"] = "Estado:"
    ws["B9"].alignment = Alignment(horizontal="right")
    ws[C_ESTADO] = ("Escreve um nome e carrega em ▶ PROCURAR, ou escolhe o critério e a " +
                    ("categoria" if etf else "região") + " e carrega em ★.")
    ws[C_ESTADO].font = Font(italic=True, color="595959")
    ws["B10"] = (("▶ PROCURAR: os " + ("ETFs" if etf else "ações") + " que o Yahoo encontra para o texto em C3. ")
                 + ("★ MAIS SIGNIFICATIVOS: os N ETFs europeus (UCITS) maiores (património) ou mais transacionados da "
                    "categoria, a partir de uma lista dos mais conhecidos (o Yahoo não tem filtro de ETFs); a dimensão "
                    "aparece na coluna G. " if etf else
                    "★ MAIS SIGNIFICATIVAS: as N ações maiores (capitalização) ou mais transacionadas (valor "
                    "médio por dia, 3 meses, só acima de 1 000 M de capitalização) da região, ou as que mais se valorizaram (subida média "
                    "por ano da reta das cotações mensais com dividendos, como nas colunas Valorização do painel: desde sempre, "
                    "10, 5, 3 ou 1 ano), escolhidas entre as até 500 maiores da região. ")
                 + "A mesma empresa / fundo em várias bolsas aparece uma só vez, na bolsa com mais negócio (as outras "
                 "ficam na última coluna). Vêm marcados com x os que ainda não estão no painel; os que já lá estão não "
                 "têm x e não se copiam. ➕ ACRESCENTAR MARCADOS: junta os que têm x à folha Acrescentados e às folhas "
                 "Ações / ETFs (com os dados todos). ↺ LISTA BASE: mostra os títulos da lista base do script que não estão "
                 "nas folhas (nem nessa nem noutra bolsa / moeda), marcados com x, para os voltares a acrescentar.")
    ws["B10"].font = Font(italic=True, size=9, color="595959")
    ws["B10"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("B10:J10")
    ws.row_dimensions[10].height = 64
    for j, h in enumerate(CAB_P, start=1):
        c = ws.cell(L_TAB, j, h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="C00000" if j == 1 else "1F3864")
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[L_TAB].height = 30
    from openpyxl.formatting.rule import FormulaRule
    fim = L_TAB + MAXP
    for i in range(L_TAB + 1, fim + 1):
        ws.cell(i, 1).alignment = Alignment(horizontal="center")
        ws.cell(i, 1).font = Font(bold=True, color="C00000")
        ws.cell(i, 7).number_format = "#,##0.0"
    ws.conditional_formatting.add(f"A{L_TAB + 1}:A{fim}", FormulaRule(
        formula=[f'AND($E{L_TAB + 1}<>"",$I{L_TAB + 1}="")'], fill=PatternFill("solid", bgColor="FFF2CC"),
        border=borda))
    ws.conditional_formatting.add(f"B{L_TAB + 1}:J{fim}", FormulaRule(
        formula=[f'AND($E{L_TAB + 1}<>"",$I{L_TAB + 1}<>"")'], font=Font(color="A6A6A6")))
    ws.conditional_formatting.add(f"B{L_TAB + 1}:J{fim}", FormulaRule(
        formula=[f'AND($E{L_TAB + 1}<>"",MOD(ROW(),2)=0)'], fill=PatternFill("solid", bgColor="D9D9D9")))
    for col, w in zip("ABCDEFGHIJ", (4, 36, 8, 10, 16, 20, 14, 8, 14, 30)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"A{L_TAB + 1}"
    ws.auto_filter.ref = f"A{L_TAB}:J{fim}"


def _yf():
    import logging
    import warnings
    logging.getLogger("yfinance").setLevel(logging.CRITICAL)
    warnings.filterwarnings("ignore")
    import yfinance as yf
    return yf


_LEGAL = {"inc", "incorporated", "corp", "corporation", "co", "company", "plc", "sa", "s", "a", "se", "ag", "nv", "n",
          "v", "spa", "p", "ltd", "limited", "group", "holding", "holdings", "the", "class", "ordinary", "shares", "adr",
          "asa", "ab", "oyj", "as", "a/s", "kgaa", "bhd", "sab", "de", "cv", "reg", "registered", "pref"}


def nome_norm(nome):
    """Nome sem acentos, pontuação, conteúdo entre parênteses nem formas jurídicas (para juntar a mesma empresa)."""
    import unicodedata
    s = unicodedata.normalize("NFKD", str(nome or "")).encode("ascii", "ignore").decode().lower()
    s = re.sub(r"\([^)]*\)", " ", s)
    s = re.sub(r"[^a-z0-9& ]+", " ", s)
    return " ".join(w for w in s.split() if w not in _LEGAL)


def _nas_folhas(ex):
    """Tickers, nomes e códigos das linhas das folhas Ações / ETFs -> nome da folha (None se não der para ler)."""
    ultima = get_column_letter(len(COLS))
    try:
        r = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)])
    except Exception:
        return None
    ja, nomes, cods = {}, {}, {}
    for idx, tab in zip((1, 2), r):
        cab = [txt(v) for v in (tab[0] if tab else [])]
        if "Ticker Yahoo" not in cab or "Empresa / ETF" not in cab:
            return None
        jn, jt = cab.index("Empresa / ETF"), cab.index("Ticker Yahoo")
        jc = cab.index("Código") if "Código" in cab else None
        onde = "Ações" if idx == 1 else "ETFs"
        for v in tab[1:]:
            if txt(v[jt]):
                ja.setdefault(txt(v[jt]).upper(), onde)
                nomes.setdefault(nome_norm(v[jn]), onde)
                if _nome_chave(v[jn]):
                    nomes.setdefault(_nome_chave(v[jn]), onde)
                if jc is not None and txt(v[jc]):
                    cods.setdefault(txt(v[jc]).upper(), onde)
    return ja, nomes, cods


def _ja_no_painel(ex):
    """Tickers e nomes já no painel (linhas das folhas Ações / ETFs + folha Acrescentados) -> onde estão."""
    nf = _nas_folhas(ex)
    if nf is None:                         # não se leram as folhas: as listas do script
        ja = {t.upper(): "Ações" for _, _, t, _ in ACOES}
        ja.update({t.upper(): "ETFs" for _, _, t, _ in ETFS})
        nomes = {nome_norm(n): "Ações" for n, _, _, _ in ACOES}
        nomes.update({nome_norm(n): "ETFs" for n, _, _, _ in ETFS})
    else:
        ja, nomes, _ = nf
    for n, c, t, b, tipo in ler_acrescentados(ex):
        ja.setdefault(t.upper(), "Acrescentados")
        nomes.setdefault(nome_norm(n), "Acrescentados")
    try:                                   # a mesma ação noutra bolsa (em euros, nos EUA ou a principal) também conta
        with open(FICH_EUR, encoding="utf-8") as f:
            for cm in json.load(f).values():
                for t, v in cm.items():
                    if v.get("alt") and t.upper() in ja:
                        ja.setdefault(v["alt"].upper(), ja[t.upper()])
    except Exception:
        pass
    nomes.pop("", None)
    return ja, nomes


def _agrupar(yf, quotes, ja, nomes):
    """Junta a mesma empresa / fundo (mesmo nome) listada em várias bolsas e fica com a listagem preferida pelo modo
    (€: uma bolsa em euros; $: uma bolsa americana; lista base: a com mais negócio, em euros, nos últimos 5 dias).
    No modo "todas as bolsas" não junta: cada listagem é uma linha (já está = esse mesmo ticker).
    Devolve [(quote escolhida, outras, onde já está)]."""
    modo = modo_atual()
    if modo == "todos":
        return [(q, [], ja.get(q["symbol"].upper())) for q in quotes]
    pref = (lambda s: 0 if em_euros(s) else 1) if modo == "eur" else (lambda s: 0 if nos_eua(s) else 1) \
        if modo == "eua" else (lambda s: 0)
    grupos = {}
    for q in quotes:
        nome = q.get("longname") or q.get("longName") or q.get("shortname") or q.get("shortName") or q["symbol"]
        grupos.setdefault(nome_norm(nome) or q["symbol"], []).append(q)
    simbolos = [q["symbol"] for g in grupos.values() if len(g) > 1 for q in g]
    neg = {}
    if simbolos:
        try:
            h = _por_ticker(yf.download(simbolos, period="5d", interval="1d", group_by="ticker", auto_adjust=False,
                                        threads=True, progress=False), simbolos)
            tx = taxas_eur({moeda_prevista(s) for s in simbolos} | {"USD"})
            for s in simbolos:
                if s in h and len(h[s].dropna(subset=["Close"])):
                    x = h[s].dropna(subset=["Close"])
                    f = tx.get(moeda_prevista(s) or "USD", (None, None))[1] or 0
                    neg[s] = float((x["Close"] * x["Volume"].fillna(0)).mean()) * f
        except Exception:
            pass
    out = []
    for chave, g in grupos.items():
        g = sorted(g, key=lambda q: (pref(q["symbol"]), -neg.get(q["symbol"], -1)))
        onde = next((ja[q["symbol"].upper()] for q in g if q["symbol"].upper() in ja), None) or nomes.get(chave)
        out.append((g[0], [q["symbol"] for q in g[1:]], onde))
    return out


def _escrever_resultados(ex, pf, linhas, estado):
    linhas = [[None if v in ("", None) else v for v in l] for l in linhas[:MAXP]]
    linhas += [[None] * NC_P] * (MAXP - len(linhas))
    com_filtros(ex, [{"op": "escrever", "folha": pf, "rng": P_TAB, "valores": linhas},
                     {"op": "escrever", "folha": pf, "rng": C_ESTADO, "valores": [[estado]]}], [pf])
    print(estado)


def _procurar_formato_ok(ex, pf):
    if txt(ex.ops([{"op": "ler", "folha": pf, "rng": f"J{L_TAB}"}])[0][0][0]) != CAB_P[-1]:
        raise SystemExit("As folhas Procurar são da versão anterior: carrega primeiro em ▶ ATUALIZAR (o painel passa ao "
                         "formato novo, com Procurar ações e Procurar ETFs separadas).")


def procurar_nome(pf):
    ex = Excel(FICHEIRO)
    _procurar_formato_ok(ex, pf)
    etf = pf == FOLHA_PROC_ETF
    texto = txt(ex.ops([{"op": "ler", "folha": pf, "rng": P_TXT}])[0][0][0])
    if not texto:
        raise SystemExit(f"Escreve em {P_TXT} (folha {NOME_PROC[pf]}) o nome ou o ticker a procurar.")
    print(f"A procurar '{texto}' no Yahoo ({'ETFs' if etf else 'ações'})...")
    yf = _yf()
    try:
        quotes = yf.Search(texto, max_results=40, news_count=0, lists_count=0, raise_errors=True).quotes or []
    except Exception as e:
        raise SystemExit(f"Não consegui procurar no Yahoo ({e}). Confirma a ligação à internet e tenta outra vez.")
    quotes = [q for q in quotes if q.get("symbol") and q.get("quoteType") == ("ETF" if etf else "EQUITY")]
    ja, nomes = _ja_no_painel(ex)
    linhas, novos = [], 0
    for q, outras, onde in _agrupar(yf, quotes, ja, nomes):
        t = q["symbol"]
        novos += not onde
        linhas.append(["x" if not onde else None, q.get("longname") or q.get("shortname") or t,
                       "ETF" if etf else "Ação", t.split(".")[0], t, q.get("exchDisp") or q.get("exchange"), None, None,
                       onde, ", ".join(outras) or None])
    _escrever_resultados(ex, pf, linhas, f"{datetime.now():%d/%m/%Y %H:%M} · '{texto}': {len(linhas)} resultado(s), "
                                         f"{novos} ainda não estão no painel (marcados com x). Tira os x que não quiseres "
                                         "e carrega em ➕ ACRESCENTAR MARCADOS.")
    time.sleep(2)


def _taxas_usd(yf, moedas):
    """Taxa para dólares da moeda principal de cada moeda do Yahoo. Devolve (taxa da capitalização, taxa do preço):
    em Londres (GBp), Joanesburgo (ZAc) e Telavive (ILA) o preço vem em cêntimos mas a capitalização vem na moeda principal."""
    menor = {"GBp": "GBP", "ZAc": "ZAR", "ILA": "ILS"}
    pedir = {menor.get(m, m) for m in moedas if m and menor.get(m, m) != "USD"}
    tx = {"USD": 1.0}
    if pedir:
        try:
            h = yf.download([f"{m}USD=X" for m in pedir], period="10d", progress=False, auto_adjust=True, threads=True)
            fecho = h["Close"] if "Close" in h else h
            for m in pedir:
                col = f"{m}USD=X"
                if col in fecho and fecho[col].dropna().size:
                    tx[m] = float(fecho[col].dropna().iloc[-1])
        except Exception as e:
            print("   aviso: sem câmbios para dólares:", e)
    out = {}
    for m in set(moedas) | {"USD"}:
        if not m:
            continue
        t = tx.get(menor.get(m, m))
        out[m] = (t, t / 100 if (t and m in menor) else t)
    return out


def _rating(x):
    """'1.8 - Buy' -> 1.8 (1 = compra forte … 5 = venda)."""
    try:
        return float(str(x.get("averageAnalystRating") or "").split("-")[0])
    except ValueError:
        return None


def _por_qualidade(yf, cand, n, so_potencial=False):
    """Pontuação como no painel (tendência das médias + potencial dos analistas; sem valor justo, que é demorado):
    1) pré-escolha com os dados do filtro (média 50 acima da 200, opinião dos analistas); 2) para os ~3×N melhores, vai
    buscar o preço-alvo; 3) ordena pela pontuação e, em empate, pelo potencial (ou só pelo potencial)."""
    pre = []
    for chave, cap_usd, x in cand:
        m50, m200, rt = _num(x.get("fiftyDayAverage")), _num(x.get("twoHundredDayAverage")), _rating(x)
        if rt is not None and rt > (3.0 if so_potencial else 2.5):
            continue
        g = m50 / m200 - 1 if (m50 and m200) else None
        if not so_potencial and (g is None or g < -PROXIMO):
            continue
        pre.append(((rt or 3.0) - (g or 0) * 5, cap_usd, x, g))
    pre.sort(key=lambda c: c[0])
    pre = pre[:max(3 * n, 30)]
    print(f"   {len(pre)} pré-escolhidas; a ir buscar o preço-alvo dos analistas...")
    out = []
    for _, cap_usd, x, g in pre:
        try:
            inf = yf.Ticker(x["symbol"]).info or {}
        except Exception:
            continue
        p, alvo = _num(inf.get("regularMarketPrice") or inf.get("currentPrice")), _num(inf.get("targetMeanPrice"))
        if not (p and alvo) or (inf.get("numberOfAnalystOpinions") or 0) < 3:
            continue
        pot = alvo / p - 1
        d = {"tipo": "Ação", "pot": pot, "justo_pct": None, "mm100c": _num(x.get("fiftyDayAverage")),
             "mm200c": _num(x.get("twoHundredDayAverage")), "gap5": None, "venda_min": None}
        pontuar(d, p)
        x["_nota"] = f"pontuação {d.get('pont')} · analistas {pot:+.0%}"
        out.append(((pot if so_potencial else (d.get("pont") or 0) * 10 + pot), cap_usd, x))
        time.sleep(0.15)
    out.sort(key=lambda c: -c[0])
    return out


def mais_significativas():
    ex = Excel(FICHEIRO)
    _procurar_formato_ok(ex, FOLHA_PROC)
    r = ex.ops([{"op": "ler", "folha": FOLHA_PROC, "rng": c} for c in (P_CRIT, P_REG, P_N)])
    crit, reg, n = txt(r[0][0][0]) or CRITERIOS[0], txt(r[1][0][0]) or "EUA + Europa", txt(r[2][0][0]) or "100"
    try:
        n = max(1, min(MAXP, int(float(n))))
    except ValueError:
        n = 100
    if reg not in REGIOES:
        raise SystemExit(f"Região desconhecida: '{reg}'. Escolhe uma da lista em {P_REG}.")
    yf = _yf()
    if not hasattr(yf, "screen") or not hasattr(yf, "EquityQuery"):
        raise SystemExit("O teu yfinance é antigo e não tem o filtro de ações. Atualiza-o (numa linha de comandos):\n"
                         "   python -m pip install -U yfinance")
    if crit not in CRITERIOS:
        crit = CRITERIOS[0]
    volume = crit == CRITERIOS[1]
    qualidade = crit in (CRITERIOS[2], CRITERIOS[3])     # precisam de analistas: universo largo, depois escolhe-se
    hv = _horiz_val(crit)
    filtros, campo = [yf.EquityQuery("gt", ["intradaymarketcap", 1e9 if volume else 0])], "intradaymarketcap"
    if volume:
        campo = "avgdailyvol3m"
    elif qualidade:
        filtros = [yf.EquityQuery("gt", ["intradaymarketcap", 3e8])]    # inclui médias empresas (p. ex. a Soitec)
    elif crit == CRITERIOS[4]:
        filtros, campo = [yf.EquityQuery("gt", ["intradaymarketcap", 1e9])], "fiftytwowkpercentchange"
    elif hv:                                          # valorização: as maiores da região, depois a reta de cada uma
        filtros = [yf.EquityQuery("gt", ["intradaymarketcap", 1e9])]
    elif crit == CRITERIOS[5]:
        filtros = [yf.EquityQuery("gt", ["intradaymarketcap", 5e8]),
                   yf.EquityQuery("btwn", ["peratio.lasttwelvemonths", 0, 15]),
                   yf.EquityQuery("gt", ["epsgrowth.lasttwelvemonths", 10])]
    print(f"A pedir ao Yahoo: {crit} · {reg} · {n}...")
    quotes = []
    for regiao in REGIOES[reg]:
        q = yf.EquityQuery("and", [yf.EquityQuery("eq", ["region", regiao])] + filtros)
        pedidos = min(MAXP, 500 if (qualidade or hv) else n * 2 if volume else n)   # pontuação/analistas/valorização: universo maior
        offset = 0
        while offset < pedidos:
            try:
                res = yf.screen(q, offset=offset, size=min(250, pedidos - offset), sortField=campo, sortAsc=False)
            except Exception as e:
                print(f"   aviso: {regiao}: {e}")
                break
            lote = (res or {}).get("quotes") or []
            for x in lote:
                x["_regiao"] = regiao
            quotes += lote
            if len(lote) < min(250, pedidos - offset):
                break
            offset += len(lote)
        print(f"   {regiao}: {sum(1 for x in quotes if x['_regiao'] == regiao)} títulos")
    if not quotes:
        raise SystemExit("O Yahoo não devolveu resultados. Tenta outra vez daqui a pouco (ou outra região).")
    tx = _taxas_usd(yf, {x.get("currency") for x in quotes})
    vistos, cand = set(), []
    for x in quotes:
        t = x.get("symbol")
        if not t or t in vistos or x.get("quoteType", "EQUITY") != "EQUITY":
            continue
        vistos.add(t)
        f_cap, f_preco = tx.get(x.get("currency") or "USD", (None, None))
        cap = x.get("marketCap")
        cap_usd = cap * f_cap if (cap and f_cap) else None
        if volume:
            vol, preco = x.get("averageDailyVolume3Month"), x.get("regularMarketPrice")
            chave = vol * preco * f_preco if (vol and preco and f_preco) else None
        elif crit == CRITERIOS[4]:
            chave = _num(x.get("fiftyTwoWeekChangePercent"))
            if chave is not None:
                x["_nota"] = f"subida 12 meses {chave / 100:+.0%}"
        else:
            chave = cap_usd
        cand.append((chave, cap_usd, x))
    if hv:
        nome_h, meses, minimo = hv
        print(f"A calcular a valorização {nome_h} de {len(cand)} ações (as maiores da região)...")
        val = _valorizacoes(yf, [x["symbol"] for _, _, x in cand], meses, minimo)
        novo = []
        for _, cap_usd, x in cand:
            v = val.get(x["symbol"])
            if v:
                x["_nota"] = f"valorização {nome_h} {v[0]:+.1%}/ano".replace(".", ",") + (f" (dados desde {v[1]})"
                                                                                       if meses is None else "")
            novo.append((v[0] if v else None, cap_usd, x))
        cand = novo
        print(f"   {len(val)} com história suficiente")
    cand.sort(key=lambda c: (c[0] is None, -(c[0] or 0)))
    if qualidade:
        cand = _por_qualidade(yf, cand, n, so_potencial=crit == CRITERIOS[3])
    ja, nomes = _ja_no_painel(ex)
    linhas, novas, grupos = [], 0, {}
    for chave, cap_usd, x in cand:                       # a mesma empresa em várias bolsas: fica a primeira (a maior)
        k = nome_norm(x.get("longName") or x.get("shortName") or x["symbol"]) or x["symbol"]
        if k in grupos:
            grupos[k][-1].append(x["symbol"])
            continue
        if len(grupos) >= n:
            continue
        grupos[k] = [chave, cap_usd, x, []]
    for k, (chave, cap_usd, x, outras) in grupos.items():
        t = x["symbol"]
        onde = ja.get(t.upper()) or next((ja[o.upper()] for o in outras if o.upper() in ja), None) or nomes.get(k)
        novas += not onde
        if x.get("_nota"):
            outras = [x["_nota"]] + outras
        linhas.append(["x" if not onde else None, x.get("longName") or x.get("shortName") or t, "Ação",
                       t.split(".")[0], t, x.get("fullExchangeName") or x.get("exchange"),
                       round(cap_usd / 1e9, 1) if cap_usd else None, x.get("currency"), onde, ", ".join(outras) or None])
    _escrever_resultados(ex, FOLHA_PROC, linhas, f"{datetime.now():%d/%m/%Y %H:%M} · {crit} · {reg}: {len(linhas)} títulos, "
                                     f"{novas} ainda não estão no painel (marcados com x). "
                                     "Tira os x que não quiseres e carrega em ➕ ACRESCENTAR MARCADOS.")
    time.sleep(2)


def etfs_significativos():
    """★ da folha Procurar ETFs: ordena os ETFs conhecidos da categoria pelo património ou pelo negócio (dados do Yahoo)."""
    ex = Excel(FICHEIRO)
    pf = FOLHA_PROC_ETF
    _procurar_formato_ok(ex, pf)
    r = ex.ops([{"op": "ler", "folha": pf, "rng": c} for c in (P_CRIT, P_REG, P_N)])
    crit, cat, n = txt(r[0][0][0]) or CRITERIOS_ETF[0], txt(r[1][0][0]) or "Todas", txt(r[2][0][0]) or "30"
    try:
        n = max(1, min(MAXP, int(float(n))))
    except ValueError:
        n = 30
    if cat not in CATEGORIAS_ETF:
        raise SystemExit(f"Categoria desconhecida: '{cat}'. Escolhe uma da lista em {P_REG}.")
    cands = list(dict.fromkeys(t for c, ts in ETFS_CANDIDATOS.items() if cat in ("Todas", c) for t in ts))
    yf = _yf()
    print(f"A pedir ao Yahoo os dados de {len(cands)} ETFs ({cat})...")
    dados = []
    for t in cands:
        try:
            inf = yf.Ticker(t).info or {}
        except Exception:
            inf = {}
        p = _num(inf.get("regularMarketPrice") or inf.get("navPrice") or inf.get("previousClose"))
        if not p:
            print(f"   {t}: sem dados no Yahoo")
            continue
        dados.append((t, inf, p))
        time.sleep(0.2)
    tx = _taxas_usd(yf, {inf.get("currency") for _, inf, _ in dados})
    tend = {}
    hv = _horiz_val(crit)
    val = _valorizacoes(yf, [t for t, _, _ in dados], hv[1], hv[2]) if hv else {}
    if crit in CRITERIOS_ETF[2:4]:                     # pontuação / subida: com as cotações do último ano e meio
        lst = [t for t, _, _ in dados]
        try:
            h = _por_ticker(yf.download(lst, period="18mo", interval="1d", group_by="ticker", auto_adjust=True,
                                        threads=True, progress=False), lst)
        except Exception:
            h = {}
        for t in lst:
            try:
                c = h[t]["Close"].dropna()
                m100, m200 = c.tail(100).mean(), c.tail(200).mean()
                tend[t] = dict(g=m100 / m200 - 1 if len(c) >= 200 else None,
                               m6=c.iloc[-1] / c.iloc[-127] - 1 if len(c) > 126 else None,
                               m12=c.iloc[-1] / c.iloc[-253] - 1 if len(c) > 252 else None,
                               p200=c.iloc[-1] / m200 - 1 if len(c) >= 200 else None)
            except Exception:
                pass
        com6 = sorted([t for t in tend if tend[t].get("m6") is not None], key=lambda t: -tend[t]["m6"])
        for i, t in enumerate(com6):
            v = tend[t]
            pos = i / (len(com6) - 1) if len(com6) > 1 else 0
            g, p200, m6 = v.get("g"), v.get("p200"), v.get("m6")
            v["pont"] = (5 if pos <= ETF_TOP5 and (g or 0) > 0 else 4 if pos <= ETF_TOP4 and (g or 0) > 0 else
                         3 if (g or 0) > 0 else 1 if (p200 or 0) < 0 and m6 < 0 else 2)
    linhas_c = []
    for t, inf, p in dados:
        f_maior, f_preco = tx.get(inf.get("currency") or "USD", (None, None))
        aum = _num(inf.get("totalAssets") or inf.get("netAssets"))
        aum_usd = aum * f_maior if (aum and f_maior) else None      # em geral na moeda principal do fundo
        vol = _num(inf.get("averageVolume") or inf.get("averageDailyVolume3Month"))
        neg = vol * p * f_preco if (vol and f_preco) else None
        v = tend.get(t, {})
        if crit == CRITERIOS_ETF[2]:
            chave = (v["pont"] * 10 + v["m6"]) if v.get("pont") else None
            inf["_nota"] = (f"pontuação {v['pont']} · subida 6 meses {v['m6']:+.0%}" if v.get("pont") else None)
        elif crit == CRITERIOS_ETF[3]:
            chave = v.get("m12")
            inf["_nota"] = f"subida 12 meses {v['m12']:+.0%}" if v.get("m12") is not None else None
        elif hv:
            vv = val.get(t)
            chave = vv[0] if vv else None
            inf["_nota"] = ((f"valorização {hv[0]} {vv[0]:+.1%}/ano".replace(".", ",")
                             + (f" (dados desde {vv[1]})" if hv[1] is None else "")) if vv else None)
        else:
            chave = neg if crit == CRITERIOS_ETF[1] else aum_usd
        linhas_c.append((chave, aum_usd, t, inf))
    linhas_c.sort(key=lambda x: (x[0] is None, -(x[0] or 0)))
    ja, nomes = _ja_no_painel(ex)
    grupos = {}
    for chave, aum_usd, t, inf in linhas_c:                # o mesmo fundo em várias bolsas: fica o primeiro (o maior)
        nome = inf.get("longName") or inf.get("shortName") or t
        k = nome_norm(nome) or t
        if k in grupos:
            grupos[k][-1].append(t)
        elif len(grupos) < n:
            grupos[k] = [aum_usd, t, inf, []]
    linhas, novos = [], 0
    for k, (aum_usd, t, inf, outras) in grupos.items():
        onde = ja.get(t.upper()) or next((ja[o.upper()] for o in outras if o.upper() in ja), None) or nomes.get(k)
        novos += not onde
        if inf.get("_nota"):
            outras = [inf["_nota"]] + outras
        linhas.append(["x" if not onde else None, inf.get("longName") or inf.get("shortName") or t, "ETF",
                       t.split(".")[0], t, inf.get("fullExchangeName") or inf.get("exchange"),
                       round(aum_usd / 1e9, 1) if aum_usd else None, inf.get("currency"), onde,
                       ", ".join(outras) or None])
    _escrever_resultados(ex, pf, linhas, f"{datetime.now():%d/%m/%Y %H:%M} · {crit} · {cat}: {len(linhas)} ETFs, "
                                         f"{novos} ainda não estão no painel (marcados com x; a coluna G é o património "
                                         "em mil M USD). Tira os x que não quiseres e carrega em ➕ ACRESCENTAR MARCADOS.")
    time.sleep(2)


def da_lista_base(pf):
    """↺ LISTA BASE: os títulos da lista do script que já não estão nas folhas (nem nessa bolsa / moeda nem noutra)."""
    ex = Excel(FICHEIRO)
    _procurar_formato_ok(ex, pf)
    etf = pf == FOLHA_PROC_ETF
    nf = _nas_folhas(ex)
    if nf is None:
        raise SystemExit("Não consegui ler as folhas Ações / ETFs: carrega primeiro em ▶ ATUALIZAR.")
    ja, nomes, cods = nf
    try:                                   # a mesma ação noutra bolsa (outra moeda) também conta
        with open(FICH_EUR, encoding="utf-8") as f:
            for cm in json.load(f).values():
                for t, v in cm.items():
                    if v.get("alt") and v["alt"].upper() in ja:
                        ja.setdefault(t.upper(), ja[v["alt"].upper()])
    except Exception:
        pass
    acr = {t.upper() for _, _, t, _, _ in ler_acrescentados(ex)}
    linhas = []
    for n, c, t, b in (ETFS if etf else ACOES):
        if (t.upper() in ja or nome_norm(n) in nomes or (_nome_chave(n) and _nome_chave(n) in nomes)
                or (c and c.upper() in cods)):
            continue
        linhas.append(["x", n, "ETF" if etf else "Ação", c, t, b, None, None,
                       "Acrescentados (não está nas folhas)" if t.upper() in acr else None, None])
    _escrever_resultados(ex, pf, linhas, f"{datetime.now():%d/%m/%Y %H:%M} · lista base: {len(linhas)} "
                         f"{'ETF(s)' if etf else 'ação(ões)'} que não estão na folha {'ETFs' if etf else 'Ações'} "
                         "(nem noutra bolsa / moeda), marcados com x. Tira os x que não quiseres e carrega em "
                         "➕ ACRESCENTAR MARCADOS." if linhas else
                         f"{datetime.now():%d/%m/%Y %H:%M} · lista base: estão todos na folha (nesta ou noutra moeda).")
    time.sleep(2)


def acrescentar_marcados(pf):
    ex = Excel(FICHEIRO)
    _procurar_formato_ok(ex, pf)
    res = ex.ops([{"op": "ler", "folha": pf, "rng": P_TAB}])[0]
    atuais = ler_acrescentados(ex)
    ja, nomes = _ja_no_painel(ex)
    novos = 0
    res = [(list(r) + [None] * NC_P)[:NC_P] for r in res]
    for r in res:
        if not txt(r[0]) or not txt(r[4]):
            continue
        t = txt(r[4])
        outras = [o.strip().upper() for o in txt(r[9]).split(",") if o.strip()]
        ja_esta = t.upper() in ja or any(o in ja for o in outras) or nome_norm(r[1]) in nomes
        if not ja_esta:
            atuais.append((txt(r[1]) or t, txt(r[3]) or t.split(".")[0], t, txt(r[5]),
                           "ETF" if txt(r[2]).upper() == "ETF" else "Ação"))
            ja[t.upper()] = "Acrescentados"
            nomes[nome_norm(r[1])] = "Acrescentados"
            novos += 1
        r[0], r[8] = None, ja.get(t.upper()) or next((ja[o] for o in outras if o in ja), None) or nomes.get(nome_norm(r[1]))
    if not novos:
        raise SystemExit(f"Nenhum título novo marcado com x na folha {NOME_PROC[pf]} (os marcados já estavam no painel).")
    linhas = [list(a) for a in atuais][:MAX_LINHAS]
    linhas += [[None] * 5] * (MAX_LINHAS - len(linhas))
    msg = (f"{datetime.now():%d/%m/%Y %H:%M} · {novos} título(s) acrescentado(s) à folha Acrescentados e às folhas "
           "Ações / ETFs.")
    res = [[None if v in ("", None) else v for v in r] for r in res]
    guardar_copia_acr(atuais)
    folha_a = FOLHA_ACR
    com_filtros(ex, [{"op": "escrever", "folha": folha_a, "rng": A_TAB, "valores": linhas},
                     {"op": "escrever", "folha": pf, "rng": P_TAB, "valores": res},
                     {"op": "escrever", "folha": pf, "rng": C_ESTADO, "valores": [[msg]]}], [folha_a, pf])
    print(msg)
    incluir_novos(atuais[-novos:])
    time.sleep(3)


def tirar_marca_internet():
    """Ficheiros vindos da internet (ou de um zip descarregado) abrem no Excel em Vista Protegida, que nenhum programa
    consegue ler. Tira essa marca (Zone.Identifier) aos ficheiros desta ferramenta."""
    if os.name != "nt":
        return
    for f in (FICHEIRO, BAT, os.path.abspath(__file__)):
        try:
            os.remove(f + ":Zone.Identifier")
        except OSError:
            pass


LOG = os.path.join(PASTA, BASE + ".log")


def registo(texto):
    """Acrescenta uma linha ao ficheiro de registo (_Painel_acoes.log), que guarda as últimas 300 linhas."""
    try:
        linhas = open(LOG, encoding="utf-8").read().splitlines()[-299:] if os.path.exists(LOG) else []
        linhas.append(f"{datetime.now():%d/%m/%Y %H:%M:%S}  {texto}")
        open(LOG, "w", encoding="utf-8").write("\n".join(linhas) + "\n")
    except Exception:
        pass


def atualizar_servidor(tipo):
    """No GitHub Actions (sem Excel nem PC): 'completa' (como o ▶ ATUALIZAR) ou 'precos' (como o ↻ PREÇOS, sobre os
    dados da última completa). Escreve index.html (cifrado), versao.txt e dados.enc; o workflow publica-os."""
    global LP_MIN, REF, BOTAO
    senha = os.environ.get("PAINEL_SENHA", "")
    if len(senha) < 8:
        raise SystemExit("Falta o segredo PAINEL_SENHA no repositório (Settings > Secrets and variables > Actions).")
    cfg = {}
    if os.path.exists(os.path.join(PASTA, "config.enc")):
        with open(os.path.join(PASTA, "config.enc"), encoding="utf-8") as f:
            cfg = json.loads(_decifrar(f.read().strip(), senha))
        print(f"Definições do PC de {cfg.get('enviado', '?')}: {len(cfg.get('acrescentados', []))} acrescentados, "
              f"{len(cfg.get('det', []))} detidos, bolsa {cfg.get('modo', 'eur')}.")
    else:
        print("Ainda sem definições do PC (config.enc): uso a lista base. Faz um ▶ ATUALIZAR no PC para as enviar.")
    guardar_modo(cfg.get("modo", "eur"))
    if cfg.get("compra"):                          # escolhas da folha Compra do PC (separador Compra da app)
        _gravar_json(FICH_COMPRA, cfg["compra"])
    if cfg.get("dos_etfs") and not os.path.exists(FICH_ACOES_ETF):
        _gravar_json(FICH_ACOES_ETF, cfg["dos_etfs"])
    LP_MIN, REF = cfg.get("lp_min", LP_MIN), None
    usar_pesos(cfg.get("pesos"), "do PC")
    if cfg.get("token_botao"):
        BOTAO = {"o": GH_USER, "r": GH_REPO, "t": cfg["token_botao"]}
    fo = os.path.join(PASTA, "ocultos.enc")          # linhas escondidas pelos filtros do Excel (enviadas pelo PC)
    if os.path.exists(fo):
        with open(fo, encoding="utf-8") as f:
            OCULTOS.update(json.loads(_decifrar(f.read().strip(), senha)))
    fd = os.path.join(PASTA, "dados.enc")
    if tipo == "precos" and os.path.exists(fd):
        with open(fd, encoding="utf-8") as f:
            dd = json.loads(_decifrar(f.read().strip(), senha))
        linhas, etfs = dd["a"], dd["e"]
        for d in linhas + etfs:
            d["det"] = marca_detido(d, set(cfg.get("det", [])), cfg.get("qtd", {}))
            d["qtd"] = d["pcompra"] = d["corr"] = None
            aplicar_compra(d, cfg.get("qtd", {}), cfg.get("pcompra", {}))
            aplicar_corr(d, cfg.get("corr", {}))
        tickers = sorted({txt(d["ticker"]) for d in linhas + etfs if txt(d.get("ticker"))})
        precos = precos_agora(tickers)
        for d in linhas + etfs:
            p = precos.get(txt(d.get("ticker")))
            if p:
                recalcular(d, *p)
        pontuar_etfs(linhas)
        pontuar_etfs(etfs)
        ponderar_etfs(linhas, etfs, _ler_json(FICH_COMP, {}))
        texto = (f"Preços automáticos de {datetime.now(): %d/%m/%Y %H:%M} UTC ({len(precos)} títulos); notas da "
                 f"atualização completa de {dd.get('quando', '?')}")
        quando = dd.get("quando", "?")
    else:
        if tipo not in ("completa", "precos") and os.path.exists(fd):     # parcial: assinalados ou NG ≥ x
            with open(fd, encoding="utf-8") as f:
                dd = json.loads(_decifrar(f.read().strip(), senha))
            ant = {1: dd["a"], 2: dd["e"]}
            if tipo == "assinalados":
                marc = set(cfg.get("assinalados", []))
                esc = lambda ls: {k for d in ls if txt(d.get("nome")) in marc or txt(d.get("ticker")).upper() in marc
                                  for k in (txt(d.get("nome")), txt(d.get("ticker")).upper())}
            else:
                lim = float(tipo.split(":")[1].replace(",", "."))
                esc = lambda ls: {k for d in ls if isinstance(d.get("nota"), (int, float)) and d["nota"] >= lim
                                  for k in (txt(d.get("nome")), txt(d.get("ticker")).upper())}
            cfg = dict(cfg, _ant=ant, _sel=(esc(ant[1]), esc(ant[2])))
            print(f"Atualização parcial ({tipo}).")
        linhas, etfs, texto = atualizar_tudo(servidor=cfg)
        texto = texto.replace("Atualizado em", "Atualização automática em") + " (hora UTC)"
        quando = f"{datetime.now():%d/%m/%Y %H:%M} UTC"
    cambios_alt(linhas + etfs)                     # câmbio para euros (valor detido)
    html, versao = pagina_html(linhas, etfs, texto)
    with open(os.path.join(PASTA, "index.html"), "w", encoding="utf-8") as f:
        f.write(pagina_cifrada(html, senha))
    with open(os.path.join(PASTA, "versao.txt"), "w", encoding="utf-8") as f:
        f.write(versao)
    with open(fd, "w", encoding="utf-8") as f:
        f.write(_cifrar(json.dumps({"a": linhas, "e": etfs, "quando": quando}, ensure_ascii=False, default=_json_num), senha))
    print(f"Página gerada ({len(linhas)} ações, {len(etfs)} ETFs) · versão {versao}")


def detetar_l0():
    """V190: o cabeçalho das folhas Ações / ETFs passou da linha 6 para a 10. Num livro de uma versão anterior, lê-se
    (até ser refeito pelo ▶ ATUALIZAR) com a linha onde ele está."""
    global L0
    if not os.path.exists(FICHEIRO):
        return
    try:
        r = Excel(FICHEIRO).ops([{"op": "ler", "folha": 1, "rng": "A6"}, {"op": "ler", "folha": 1, "rng": "A10"}])
        a6, a10 = txt(r[0][0][0]), txt(r[1][0][0])
        if a10 == "Atualizar (x)":
            L0 = 10
        elif a6 == "Atualizar (x)":
            L0 = 6
            print("Livro de uma versão anterior (cabeçalho na linha 6): o ▶ ATUALIZAR TUDO passa-o para a nova disposição.")
        registo(f"linha do cabeçalho: {L0}")
    except BaseException as e:
        registo(f"aviso: linha do cabeçalho não detetada ({e}); fica {L0}")


def main():
    if "--janela" in sys.argv:               # janela da posição (processo à parte)
        janela_posicao()
        sys.exit(0)
    ficheiro_link_app()
    if "--servidor" in sys.argv:             # GitHub Actions: python _Painel_acoes.py --servidor completa|precos
        i = sys.argv.index("--servidor")
        tipo = (sys.argv[i + 1] if len(sys.argv) > i + 1 else "completa").strip().lower() or "completa"
        print(f"Painel de ações · script {TXT_VERSAO} · servidor · {tipo}")
        atualizar_servidor(tipo)
        return
    print(f"Painel de ações · script {TXT_VERSAO}")
    registo(f"início · {TXT_VERSAO} · argumentos: {sys.argv[1:] or '(nenhum: duplo clique ou botão sem argumento)'} · bat: {BAT}")
    tirar_marca_internet()
    registar_botao()
    arg = " ".join(sys.argv[1:]).lower()
    if "--modelo" not in arg:
        detetar_l0()
    if "--modelo" in arg:   # só a lista, sem dados (primeiro quadro)
        vazio = lambda lista, tipo: [dict(nome=n, tipo=tipo, codigo=c, ticker=t, bolsa=b, notas="por atualizar")
                                     for n, c, t, b in lista]
        construir(vazio(ACOES, "Ação"), "Ainda sem dados: carrega em ▶ ATUALIZAR", vazio(ETFS, "ETF")).save(FICHEIRO)
        print("modelo criado:", FICHEIRO)
    elif "repor_detidos" in arg:             # V185: ↺ REPOR DETIDOS (de uma cópia)
        repor_detidos()
    elif "posicao" in arg:                   # ▣ VER POSIÇÃO
        ver_posicao()
    elif "fechar_quadro" in arg:             # ✕ do quadro da carteira (versões anteriores)
        Excel(FICHEIRO).ops([{"op": "fechar_quadro"}], guardar=True)
    elif "://historico" in arg:              # folha Histórico: ▶ GERAR HISTÓRICO
        calcular_historico()
    elif "://cfalt" in arg:                  # V190: ▶ CALCULAR NOTA POR GANHOS (folha Ações ou ETFs)
        calcular_alt("etf" in arg.split("://cfalt", 1)[1].lower())
    elif "://comprent" in arg:               # folha Comp Rent Títulos: ▶ CALCULAR
        calcular_comp_rent()
    elif "://evolucao" in arg:               # folha Evolução: ▶ GERAR EVOLUÇÃO
        calcular_evolucao()
    elif "://alvojusto" in arg:              # folha Evolução: ▶ ANALISTAS vs VALOR JUSTO
        comparar_alvo_justo()
    elif "://compra" in arg:                 # folha Compra: ▶ CALCULAR COMPRA
        calcular_compra()
    elif "simular" in arg:                   # folha Simulador: ▶ SIMULAR (calcula à parte)
        simular()
    elif "comparar" in arg:                  # folha Comparar notas: ⚖ COMPARAR
        comparar_notas()
    elif "reclassificar" in arg:             # folha Ponderações: ↻ RECLASSIFICAR E REORDENAR
        atualizar_precos(reclassificar=True)
    elif "ordenar" in arg:                   # ↕ ORDENAR (painelacoes://ordenar/1 ou /2)
        ordenar(2 if re.search(r"ordenar/+2", arg) else 1)
    elif "procurar" in arg:                  # folhas Procurar: ▶ PROCURAR (/acoes ou /etfs)
        procurar_nome(FOLHA_PROC_ETF if "etf" in arg else FOLHA_PROC)
    elif "://top" in arg:                    # folhas Procurar: ★ MAIS SIGNIFICATIVAS / SIGNIFICATIVOS
        if "etf" in arg:
            etfs_significativos()
        else:
            mais_significativas()
    elif "://base" in arg:                   # folhas Procurar: ↺ LISTA BASE (/acoes ou /etfs)
        da_lista_base(FOLHA_PROC_ETF if "etf" in arg else FOLHA_PROC)
    elif "acrescentar" in arg:               # folhas Procurar: ➕ ACRESCENTAR MARCADOS (/acoes ou /etfs)
        acrescentar_marcados(FOLHA_PROC_ETF if "etf" in arg else FOLHA_PROC)
    elif "precos" in arg:                    # ↻ PREÇOS (rápido)
        atualizar_precos()
    elif "://modo" in arg:                   # ◎ / € / $ : bolsa preferida, e atualiza tudo
        m = next((k for k in MODOS if re.search(rf"modo/+{k}", arg)), "eur")
        guardar_modo(m)
        print(f"Bolsa: {MODOS[m]}. A atualizar o painel...")
        atualizar_tudo()
    elif re.search(r"atualizar/+acoes", arg):   # ▶ ATUALIZAR AÇÕES (folha Ações)
        atualizar_tudo("acoes")
    elif re.search(r"atualizar/+etfs", arg):    # ▶ ATUALIZAR ETFs (folha ETFs)
        atualizar_tudo("etfs")
    else:                                    # duplo clique no .bat (ou link antigo): tudo
        atualizar_tudo()


def ler_folha_atual(idx):
    """Linhas e texto de atualização (K3) da folha idx (1 = Ações, 2 = ETFs) do painel atual, para ficarem como estão
    quando só se atualiza a outra folha. None se não der para ler (ou as colunas forem de outra versão)."""
    if not os.path.exists(FICHEIRO):
        return None
    ultima = get_column_letter(len(COLS))
    r = _tentar(Excel(FICHEIRO), [{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"},
                                  {"op": "ler", "folha": idx, "rng": C_AT}])
    if not r or not cab_ok([txt(v) for v in r[0][0]]):
        return None
    chaves = [k for k, *_ in COLS]
    linhas = [linha_lida(chaves, v) for v in r[0][1:] if txt(linha_lida(chaves, v).get("ticker"))]
    return linhas, txt(r[1][0][0]) or "Atualizado antes"


def atualizar_tudo(so=None, servidor=None):
    if True:
        global REF, LP_MIN
        outra = None
        if servidor is not None:                 # no GitHub (sem Excel): definições vindas do PC
            so = None
        if so:                                   # a outra folha fica como está (lida antes de o Excel fechar)
            outra = ler_folha_atual(2 if so == "acoes" else 1)
            if outra is None:
                print("Não foi possível ler a outra folha do painel atual: atualiza-se tudo.")
                so = None
        if servidor is not None:
            criterios = None
            acrescentados = [tuple(a) for a in servidor.get("acrescentados", [])]
            estado = {"det": set(servidor.get("det", [])), "qtd": servidor.get("qtd", {}),
                      "universo": set(servidor.get("universo", [])), "linhas_ant": servidor.get("_ant", {}),
                      "pcompra": servidor.get("pcompra", {}), "corr": servidor.get("corr", {}),
                      "lp_min": servidor.get("lp_min", LP_MIN), "ref_escolha": "Atual"}
        else:
            criterios, acrescentados, estado = ler_estado()
            carregar_pesos(Excel(FICHEIRO) if os.name == "nt" or os.path.exists(FICHEIRO) else None)
            carregar_simulador(Excel(FICHEIRO) if os.name == "nt" or os.path.exists(FICHEIRO) else None)
        LP_MIN = estado.get("lp_min", LP_MIN)
        print(f"Longo prazo mínimo: {LP_MIN:.1%}/ano (abaixo disto, o que seria compra passa a 2.5 Não Rent. Longo Prazo).")
        if not estado.get("det"):
            estado["det"] = ler_detidos()
            estado["qtd"] = ler_qtd()
            estado["pcompra"] = ler_qtd("pcompra")
            estado["corr"] = ler_qtd("corr")
        # cada folha com a sua data de referência
        REF_A = data_ref(estado.get("ref_escolha"), estado.get("ref_data"))
        REF_E = data_ref(estado.get("ref_escolha_e", estado.get("ref_escolha")),
                         estado.get("ref_data_e", estado.get("ref_data")))
        for nm, rf in (("ações", REF_A), ("ETFs", REF_E)):
            if rf is not None:
                print(f"{nm}: situação em {rf:%d/%m/%Y} (preços e médias dessa data; analistas e valor justo de hoje).")
        # só os títulos que estão nas folhas (os que apagou não voltam)
        univ = None if (estado.get("universo_incompleto") or not estado.get("universo")) else estado["universo"]
        # o que atualizar em cada folha: Tudo / Assinalados / NG ≥ x (as outras linhas ficam como estão, a rosa)
        ant = estado.get("linhas_ant", {})

        def escolhidos(idx, sel):
            if not sel or sel == SELECOES[0]:
                return None
            if idx not in ant:
                print(f"{'Ações' if idx == 1 else 'ETFs'}: não consegui ler as linhas da folha: atualiza-se tudo.")
                return None
            if sel.startswith("Assinalados"):
                ls = [d for d in ant[idx] if txt(d.get("sel"))]
                if not ls:
                    print(f"{'Ações' if idx == 1 else 'ETFs'}: 'Assinalados' mas nenhum x na coluna A: não se atualiza nenhum título.")
            else:
                lim = float(sel.split("≥")[1].strip().replace(",", "."))
                ls = [d for d in ant[idx] if isinstance(d.get("nota"), (int, float)) and d["nota"] >= lim]
            print(f"{'Ações' if idx == 1 else 'ETFs'}: atualização parcial ({sel}): {len(ls)} de {len(ant[idx])} títulos.")
            return {txt(d.get("nome")) for d in ls} | {txt(d.get("ticker")).upper() for d in ls}
        esc_a = escolhidos(1, estado.get("sel")) if so in (None, "acoes") and servidor is None else None
        esc_e = escolhidos(2, estado.get("sel_e")) if so in (None, "etfs") and servidor is None else None
        if servidor is not None and servidor.get("_sel"):
            esc_a, esc_e = servidor["_sel"]
        fa = esc_a if esc_a is not None else univ
        fe = esc_e if esc_e is not None else univ
        dos_etfs = [tuple(x) for x in _ler_json(FICH_ACOES_ETF, [])]       # ações acrescentadas pelas composições
        if univ and dos_etfs:                    # ações dos ETFs que apagou da folha: não voltam a ser acrescentadas
            apag = [x for x in dos_etfs if x[0] not in univ and x[2].upper() not in univ]
            if apag:
                dos_etfs = [x for x in dos_etfs if x not in apag]
                _gravar_json(FICH_ACOES_ETF, [list(x) for x in dos_etfs])
                comp0 = _ler_json(FICH_COMP, {})
                comp0["_apagados"] = sorted(set(comp0.get("_apagados", [])) | {x[2] for x in apag})
                _gravar_json(FICH_COMP, comp0)
                print(f"   {len(apag)} ação(ões) dos ETFs apagadas da folha: não voltam a ser acrescentadas.")
        ja = {t.upper() for _, _, t, _, _ in acrescentados}
        acr = list(acrescentados) + [(n, c, t, b, "Ação") for n, c, t, b in dos_etfs if t.upper() not in ja]
        REPETIDOS.clear()
        if so is None and REF_A != REF_E:        # datas diferentes nas duas folhas: cada uma com a sua
            REF = REF_A
            linhas, _ = recolher(acr, "acoes", filtro=(fa, fe))
            REF = REF_E
            _, etfs = recolher(acr, "etfs", filtro=(fa, fe))
            REF = REF_A
        else:
            REF = REF_E if so == "etfs" else REF_A
            linhas, etfs = recolher(acr, so, filtro=(fa, fe))
        base = {a[2].upper() for a in ACOES + ETFS} | {t.upper() for _, _, t, _, _ in acrescentados}
        if dos_etfs:                              # ações dos ETFs que afinal já estavam no painel: saem da lista
            fica = [x for x in dos_etfs if x[2].upper() not in REPETIDOS and x[2].upper() not in base]
            if len(fica) < len(dos_etfs):
                _gravar_json(FICH_ACOES_ETF, [list(x) for x in fica])
                print(f"   {len(dos_etfs) - len(fica)} ação(ões) dos ETFs já estavam no painel com outro ticker: retiradas da lista.")
                dos_etfs = fica
        sit = f" · SITUAÇÃO EM {REF:%d/%m/%Y}" if REF is not None else ""
        agora = f"Atualizado em {datetime.now():%d/%m/%Y %H:%M}{sit} · script {TXT_VERSAO}"
        if so == "acoes":
            etfs, txt_e = outra
            pontuar_etfs(etfs)
        elif so == "etfs":
            linhas, txt_a = outra
        def juntar(novas, idx, esc):
            marcas = {txt(d0.get("ticker")).upper() for d0 in ant.get(idx, []) if txt(d0.get("sel"))}
            for d in novas:                      # as marcas da coluna A ficam
                d["sel"] = "x" if txt(d.get("ticker")).upper() in marcas else None
            if esc is None:                      # atualizada toda: nenhuma linha fica a rosa
                for d in novas:
                    d["desat"] = None
                return novas
            por_tk = {txt(d.get("ticker")).upper(): d for d in novas}
            res = []
            ref_f = REF_A if idx == 1 else REF_E
            for d0 in ant.get(idx, []):
                if ref_f is None:
                    d0["desde_ref"] = None       # "Atual": a coluna fica vazia
                d = por_tk.pop(txt(d0.get("ticker")).upper(), None)
                if d is not None:
                    d["desat"] = None
                    res.append(d)
                else:
                    d0["desat"] = 1              # fica como estava, a vermelho claro / rosa claro
                    res.append(d0)
            return res + list(por_tk.values())
        if so in (None, "acoes"):
            linhas = juntar(linhas, 1, esc_a)
        if so in (None, "etfs"):
            etfs = juntar(etfs, 2, esc_e)
            if esc_e is not None:
                pontuar_etfs(etfs)
        comp = _ler_json(FICH_COMP, {})
        if so != "acoes":                       # composições novas; ações que faltam na folha Ações -> acrescentam-se
            comp.update(COMPOSICAO)
            _gravar_json(FICH_COMP, comp)
            falta = acoes_em_falta(linhas, etfs, comp)
            if falta:
                print(f"\n{len(falta)} ação(ões) das composições dos ETFs não estão na folha Ações: a acrescentar...")
                novas, _ = recolher((), "acoes", apenas=falta)
                ja_tk = {txt(d.get("ticker")).upper() for d in linhas}
                novas = [d for d in novas if d.get("preco") and d.get("pont") is not None
                         and txt(d.get("ticker")).upper() not in ja_tk]
                linhas = linhas + novas
                bons = {txt(d.get("codigo")).upper() for d in novas} | {txt(d.get("ticker")).upper() for d in novas}
                guardar = [list(x) for x in dos_etfs] + [list(x) for x in falta
                                                         if x[2].upper() in bons or x[1].upper() in bons]
                _gravar_json(FICH_ACOES_ETF, guardar)
                comp["_sem_dados"] = sorted(set(comp.get("_sem_dados", [])) |
                                            {x[2] for x in falta if x[2].upper() not in bons and x[1].upper() not in bons})
                _gravar_json(FICH_COMP, comp)
                print(f"   acrescentadas: {len(novas)} de {len(falta)} (as outras não têm dados no Yahoo)")
        ponderar_etfs(linhas, etfs, comp)
        n_alvo = sum(1 for d in linhas if isinstance(d.get("pot"), (int, float)))
        if so is None:
            texto = (f"{agora} · {len(linhas)} ações, {n_alvo} com preço-alvo · {len(etfs)} ETFs")
        elif so == "acoes":
            texto = (f"{agora} · {len(linhas)} ações, {n_alvo} com preço-alvo", txt_e)
        else:
            texto = (txt_a, f"{agora} · {len(etfs)} ETFs")
        if servidor is not None:                # no GitHub: só os dados (a página é feita por atualizar_servidor)
            for d in list(linhas) + list(etfs):
                d["det"] = marca_detido(d, estado.get("det") or set(), estado.get("qtd") or {})
                d["qtd"] = d["pcompra"] = d["corr"] = None
                aplicar_compra(d, estado.get("qtd") or {}, estado.get("pcompra") or {})
                aplicar_corr(d, estado.get("corr") or {})
            return linhas, etfs, texto
        global QUADRO
        QUADRO = dados_posicao(linhas, etfs)
        gravar_posicao(QUADRO)
        assin = {k for ls in estado.get("linhas_ant", {}).values() for d in ls if txt(d.get("sel"))
                 for k in (txt(d.get("nome")), txt(d.get("ticker")).upper())}
        SYNC.update(universo=sorted(estado.get("universo") or []), assinalados=sorted(assin))
        SYNC.update(acrescentados=[list(a) for a in acrescentados], det=sorted(x for x in (estado.get("det") or set()) if x),
                    qtd=estado.get("qtd", {}), pcompra=estado.get("pcompra", {}), corr=estado.get("corr", {}),
                    modo=modo_atual(), lp_min=LP_MIN, pesos=dict(PESOS), dos_etfs=_ler_json(FICH_ACOES_ETF, []))
        g_no_atualizar(linhas, etfs, estado, so)   # V190: nota por ganhos (descarrega 10 anos de cotações)
        wb = construir(linhas, texto, etfs, criterios, acrescentados, estado)
        OCULTOS.clear()
        guardar_e_abrir(wb)                     # abre o livro e repõe os filtros (OCULTOS = as linhas que escondem)
        gerar_html(linhas, etfs, texto if isinstance(texto, str) else texto[0] if so == "acoes" else texto[1])
        qual = {"acoes": f"{len(linhas)} ações ({n_alvo} com preço-alvo); ETFs como estavam",
                "etfs": f"{len(etfs)} ETFs; ações como estavam"}.get(so, f"{len(linhas)} ações ({n_alvo} com preço-alvo) e {len(etfs)} ETFs")
        print(f"\nPainel atualizado: {qual}.")
        mc = [motivo_caro(d) for d in linhas]
        mc = [m for m in mc if m]
        if mc:
            print(f"Caras (nível de preço até {PESOS['nivel_caro']:.1f}, a azul): {len(mc)} de {len(linhas)} ações "
                  "(limite na folha Ponderações).")


# ================================================================================================ Evolução (V152)
FOLHA_EVO = "Evolução"
FICH_EVO = os.path.join(PASTA, BASE + "_evolucao.json")
EVO_NF, EVO_NM = 40, 10            # séries de compra forte (no máximo) e de compra média (as 10 de maior nota)
EVO_PTS = 262                      # pontos por gráfico (5 anos: um a cada ~5 dias de bolsa)
EVO_H = 25                         # linhas por bloco (título, resumo, gráfico, lista)
EVO_COL0 = 30                      # coluna AD: dados dos gráficos (escondidos)
EVO_L0 = 8                         # 1.ª linha dos dados (a de cima tem os nomes das séries)
EVO_NC = 4 + EVO_NM + EVO_NF       # colunas: data, média de todos, compra média (+ a sua média), compra forte (+ a sua média)
EVO_JM = 2 + EVO_NM                # coluna (relativa) da média das «Compra média»
EVO_JF = EVO_NC - 1                # coluna (relativa) da média das «Compra forte»
EVO_V = 2
EVO_COR_F, EVO_COR_M, EVO_COR_T = "1E7B34", "A9D18E", "7F7F7F"
EVO_FIXAS = (("há 5 anos", 60), ("há 3 anos", 36), ("há 1 ano", 12), ("há 6 meses", 6))   # (rótulo, meses)
EVO_NA = 5                         # datas aleatórias
EVO_LIM_PADRAO = ((3.95, 3.5), (3.95, 3.5), (4.5, 3.75))   # V155: (forte, média) ações / ETFs com nível de preço / sem nível
EVO_LIM_ROT = ("Ações (com nível de preço)", "ETFs (com nível de preço)", "Sem nível de preço")
EVO_SALTO, EVO_PARADO = 0.40, 20   # V155: cotações suspeitas: ≥ 2 dias com ±40% ou mais, ou ≥ 20 dias seguidos sem mexer


def _evo_lin(b):
    """1.ª linha do bloco b (0-8 ações, 9-17 ETFs)."""
    n = len(EVO_FIXAS) + EVO_NA
    return 9 + b * EVO_H + (2 if b >= n else 0)


def _evo_col(b):
    return EVO_COL0 + b * EVO_NC


def folha_evolucao(wb):
    """Folha Evolução: em datas passadas (5, 3 e 1 ano, 6 meses e 5 datas aleatórias), que títulos o painel dava como
    «Compra forte» e as 10 primeiras «Compra média» (sem «esperar»), e como evoluíram desde então até hoje."""
    from openpyxl.chart import LineChart, Reference, Series
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    g = _ler_json(FICH_EVO, {})
    if g.get("v") != EVO_V:
        g = {}
    ws = wb.create_sheet(FOLHA_EVO)
    ws["A1"] = "Evolução das compras: o que o painel aconselhava em datas passadas e o que aconteceu depois"
    ws["A1"].font = Font(size=14, bold=True)
    ws["B2"] = ("Carrega em ▶ GERAR EVOLUÇÃO. Para cada data (há 5 anos, 3 anos, 1 ano, 6 meses e 5 datas aleatórias, "
                "novas em cada geração), o painel calcula a decisão de todos os títulos das folhas Ações e ETFs com as "
                "regras e pesos atuais, mas só com o que se sabia nesse dia (como na folha Histórico: cotações e, nas "
                "ações dos EUA e da Europa, os relatórios anuais da SEC e ESEF; sem analistas). Cada gráfico mostra, a partir dessa data (= 100), "
                f"todos os «Compra forte» (verde escuro, até {EVO_NF}) e as {EVO_NM} «Compra média» de maior nota global "
                "(verde claro), com a média de cada grupo numa linha grossa da mesma cor; «esperar» não entra. A linha cinzenta tracejada é a média de todos os títulos do grupo "
                "(ações ou ETFs), para comparar. Cotações em euros, com dividendos, escala logarítmica. Atenção: a lista "
                "é a de hoje (títulos que entraram por terem subido favorecem o resultado) e os fundamentos passados só "
                "existem para as ações dos EUA (SEC) e da Europa (ESEF, desde 2020). Os ETFs recebem, em cada data, a nota e o nível de preço das suas ações "
                "(pela composição de hoje), como na tabela: com nível de preço, compra forte a partir de 3,95. Gerar demora alguns minutos (todas as cotações de 12 anos).")
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B2"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("B2:N2")
    ws.row_dimensions[2].height = 92
    botao(ws, "B4", "▶ GERAR EVOLUÇÃO", "painelacoes://evolucao", "1F4E79", "16365C", 11)
    ws.merge_cells("B4:D5")
    botao(ws, "P4", "▶ ANALISTAS vs VALOR JUSTO (12 datas)", "painelacoes://alvojusto", "7030A0", "4B2070", 10)
    ws.merge_cells("P4:S5")
    ws["P6"] = "Abre um ficheiro à parte: o que previam em 12 datas (3 anos) e o que aconteceu no ano seguinte."
    ws["P6"].font = Font(size=8, italic=True, color="595959")
    ws["B6"] = g.get("gerado") or "Ainda não gerado."
    ws["B6"].font = Font(size=9, color="7030A0")
    lim = g.get("limites") or EVO_LIM_PADRAO           # V155: limites só desta folha (para testar)
    ws["F3"] = "Limites para testar (só nesta folha)"
    ws["F3"].font = Font(bold=True, size=9)
    ws["J3"], ws["K3"] = "Compra forte ≥", "Compra média ≥"
    for c_ in ("J3", "K3"):
        ws[c_].font, ws[c_].alignment = Font(bold=True, size=9), Alignment(horizontal="center", wrap_text=True)
    for i, rot in enumerate(EVO_LIM_ROT):
        ws.cell(4 + i, 6, rot).font = Font(size=9)
        for j in range(2):
            c = ws.cell(4 + i, 10 + j, (lim[i] if i < len(lim) else EVO_LIM_PADRAO[i])[j])
            c.fill, c.number_format = PatternFill("solid", fgColor="FFF2CC"), "0.00"
            c.alignment = Alignment(horizontal="center")
    ws["L4"] = ("Por omissão os do painel (3,95 / 3,5; sem nível 4,5 / 3,75). Só conta quem tem compra nesse dia "
                "(sinal ≥ 3, sem «esperar» nem saúde frágil).")
    ws["L4"].font, ws["L4"].alignment = Font(size=8, italic=True, color="595959"), Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("L4:N6")
    ws.column_dimensions["J"].width = ws.column_dimensions["K"].width = 10
    if g.get("excluidos"):
        ws["B7"] = "Excluídos (cotações suspeitas): " + g["excluidos"]
        ws["B7"].font = Font(size=8, color="C00000")
    ws.column_dimensions["A"].width = 3
    n = len(EVO_FIXAS) + EVO_NA
    for lin, txt_ in ((8, "AÇÕES"), (_evo_lin(n) - 2, "ETFs")):
        c = ws.cell(lin, 2, txt_)
        c.font, c.fill = Font(bold=True, size=12, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
        for j in range(3, 15):
            ws.cell(lin, j).fill = PatternFill("solid", fgColor="1F4E79")
    blocos = g.get("blocos") or []
    for c in range(EVO_COL0, EVO_COL0 + 2 * n * EVO_NC):
        ws.column_dimensions[get_column_letter(c)].hidden = True
    for b in range(2 * n):
        r0, c0 = _evo_lin(b), _evo_col(b)
        bl = blocos[b] if b < len(blocos) else {}
        ws.cell(r0, 2, bl.get("titulo") or ("Ações" if b < n else "ETFs") + " · por gerar").font = Font(bold=True, size=12, color="1F4E79")
        c = ws.cell(r0 + 1, 2, bl.get("resumo"))
        c.font, c.alignment = Font(size=9, bold=True), Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r0 + 1, start_column=2, end_row=r0 + 1, end_column=14)
        ws.row_dimensions[r0 + 1].height = 26
        c = ws.cell(r0 + 22, 2, bl.get("lista"))
        c.font, c.alignment = Font(size=8, color="404040"), Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r0 + 22, start_column=2, end_row=r0 + 23, end_column=14)
        ws.row_dimensions[r0 + 22].height = ws.row_dimensions[r0 + 23].height = 30
        nomes = bl.get("nomes") or (["Data", "Média de todos"] + [f"Compra média {i + 1}" for i in range(EVO_NM)]
                                    + ["Média das compras médias"] + [f"Compra forte {i + 1}" for i in range(EVO_NF)]
                                    + ["Média das compras fortes"])
        for j, nm in enumerate(nomes[:EVO_NC]):
            ws.cell(EVO_L0 - 1, c0 + j, nm)
        for i, lin in enumerate((bl.get("dados") or [])[:EVO_PTS]):
            for j, v in enumerate(lin[:EVO_NC]):
                if j == 0 and isinstance(v, str):
                    try:
                        v = datetime.strptime(v, "%Y-%m-%d")
                    except ValueError:
                        pass
                ws.cell(EVO_L0 + i, c0 + j, v)
        for r in range(EVO_L0, EVO_L0 + EVO_PTS):
            ws.cell(r, c0).number_format = "dd/mm/yyyy"
        ch = LineChart()                                # eixo das datas = só os pontos que existem
        ch.style = 13
        xs = Reference(ws, min_col=c0, min_row=EVO_L0, max_row=EVO_L0 + EVO_PTS - 1)
        for j in range(1, EVO_NC):
            sr = Series(Reference(ws, min_col=c0 + j, min_row=EVO_L0 - 1, max_row=EVO_L0 + EVO_PTS - 1),
                        title_from_data=True)
            if j == 1:
                cor, w, tr = EVO_COR_T, 28000, "dash"
            elif j == EVO_JM:                           # V154: média das verdes claras (grossa)
                cor, w, tr = EVO_COR_M, 57000, None
            elif j == EVO_JF:                           # V154: média das verdes escuras (grossa)
                cor, w, tr = EVO_COR_F, 57000, None
            elif j < EVO_JM:
                cor, w, tr = EVO_COR_M, 15000, None
            else:
                cor, w, tr = EVO_COR_F, 17000, None
            ln = LineProperties(solidFill=cor, w=w)
            if tr:
                ln.prstDash = tr
            sr.graphicalProperties = GraphicalProperties(ln=ln)
            sr.marker = Marker(symbol="none")
            sr.smooth = False
            ch.series.append(sr)
        ch.set_categories(xs)
        ch.legend = None
        ch.x_axis.number_format = "mm-yyyy"
        ch.x_axis.majorGridlines = None
        ch.x_axis.tickLblPos = "low"
        ch.y_axis.number_format = "0"
        ch.y_axis.scaling.logBase = 10
        ch.y_axis.scaling.min = 10                      # V154: o eixo começa em 10
        ch.y_axis.title = "base 100"
        try:
            ch.x_axis.delete = False
            ch.y_axis.delete = False
            from openpyxl.chart.text import RichText
            from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
            for ax in (ch.x_axis, ch.y_axis):
                ax.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=800)),
                                                endParaRPr=CharacterProperties(sz=800))])
        except Exception:
            pass
        ch.width, ch.height = 25, 9.6
        ch.visible_cells_only = False
        ch.display_blanks = "gap"
        ws.add_chart(ch, f"B{r0 + 2}")


def _evo_pct(v):
    return f"{v - 100:+.0f}%".replace("-", "−")


def calcular_evolucao():
    """▶ GERAR EVOLUÇÃO (folha Evolução)."""
    import random
    import pandas as pd
    from datetime import date as date_, timedelta
    ex = Excel(FICHEIRO)
    carregar_pesos(ex)
    ultima = get_column_letter(len(COLS))
    rr = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)])
    linhas, por_t = [], {}
    for i, tab in zip((1, 2), rr):
        for d in linhas_por_cabecalho(tab, i):
            t = txt(d.get("ticker")).upper()
            if t and t not in por_t:
                por_t[t] = d
                linhas.append(d)
    if not linhas:
        print("Evolução: as folhas Ações e ETFs não têm títulos.")
        sys.exit(1)
    tks = list(por_t)
    moedas = {t: txt(por_t[t].get("moeda")) or "EUR" for t in tks}
    tipos = {t: "ETF" if txt(por_t[t].get("tipo")) == "ETF" else "Ação" for t in tks}
    try:                                               # V155: limites desta folha
        rl = ex.ops([{"op": "ler", "folha": FOLHA_EVO, "rng": "J4:K6"}])[0]
        lim = [tuple(_num_ou(rl[i][j], EVO_LIM_PADRAO[i][j]) or EVO_LIM_PADRAO[i][j] for j in range(2)) for i in range(3)]
    except Exception:
        lim = [tuple(x) for x in EVO_LIM_PADRAO]
    # V155: o histórico vem da bolsa principal quando a tabela usa outra (ex.: Equinor em Frankfurt -> Oslo EQNR.OL):
    # as bolsas secundárias têm muitas vezes históricos errados ou parados
    princ = {}
    for t in tks:
        m_ = re.search(r"principal: .* (\S+)$", txt(por_t[t].get("bolsa")))
        if m_ and m_.group(1).upper() != t:
            princ[t] = m_.group(1).upper()
    yf = _yf()
    moedas_p = {}
    for t, pt in princ.items():
        try:
            moedas_p[pt] = yf.Ticker(pt).fast_info.get("currency") or ""
        except Exception:
            moedas_p[pt] = ""
    princ = {t: pt for t, pt in princ.items() if moedas_p.get(pt)}
    hoje = date_.today()
    ini = hoje - timedelta(days=int(12 * 365.25))
    print(f"Evolução: cotações de {len(tks)} títulos (12 anos, para as médias e as retas)"
          + (f"; {len(princ)} pela bolsa principal" if princ else "") + " ...")
    eur, _ = _sim_cotacoes_eur(tks + sorted(set(princ.values())), {**moedas, **moedas_p}, ini, hoje)
    loc = _SIM_LOCAL.get("tab")
    for t, pt in princ.items():                        # usa a principal se tiver cotações de pelo menos 5 anos
        if pt in loc.columns and pt in eur.columns and loc[pt].dropna().size >= 1260:
            loc[t], eur[t] = loc[pt], eur[pt]
    cols = [t for t in tks if loc is not None and t in loc.columns and t in eur.columns]
    corte = loc.index[-1] - pd.Timedelta(days=int(6 * 365.25))
    excl = []
    for t in cols:                                     # V155: cotações suspeitas ficam de fora
        s_ = loc[t][loc.index >= corte].dropna()
        if len(s_) < 30:
            continue
        r_ = s_ / s_.shift(1)
        saltos = int(((r_ >= 1 + EVO_SALTO) | (r_ <= 1 / (1 + EVO_SALTO))).sum())
        igual = (s_.diff() == 0).astype(int)
        parado = int(igual.groupby((igual == 0).cumsum()).sum().max() or 0)
        if saltos >= 2 or parado >= EVO_PARADO:
            excl.append(t)
            print(f"   {t}: cotações suspeitas ({saltos} saltos de ±{EVO_SALTO:.0%}, {parado} dias seguidos parada): fica de fora")
    cols = [t for t in cols if t not in excl]
    if not cols:
        print("Evolução: não consegui as cotações.")
        sys.exit(1)
    loc, eur = loc[cols], eur[cols]
    idx_ = loc.dropna(how="all").index
    fim = idx_[-1]
    dia_bolsa = lambda dt: idx_[max(0, idx_.searchsorted(pd.Timestamp(dt), side="right") - 1)]
    datas = [(r_, dia_bolsa(fim - pd.DateOffset(months=m))) for r_, m in EVO_FIXAS]
    cand = [d for d in idx_ if fim - pd.DateOffset(years=5) <= d <= fim - pd.Timedelta(days=60)]
    escolhidas = []
    random.shuffle(cand)
    for d in cand:                                   # longe (≥ 3 semanas) das outras datas
        if all(abs((d - e).days) >= 21 for e in escolhidas + [x for _, x in datas]):
            escolhidas.append(d)
        if len(escolhidas) == EVO_NA:
            break
    datas += [("data aleatória", d) for d in sorted(escolhidas)]
    dias = sorted({d for _, d in datas})
    sec = {}
    try:
        sec = sec_preparar([por_t[t] for t in cols if tipos[t] != "ETF"], dias[0], hoje)
    except Exception as e:
        print(f"   SEC: {e}")
    print(f"   a calcular a decisão do painel de {len(cols)} títulos em {len(dias)} datas ({len(sec)} com relatórios da SEC / ESEF) ...")
    hold = _sim_etf_hold([por_t[t] for t in cols if tipos[t] == "ETF"], [por_t[t] for t in cols if tipos[t] != "ETF"])
    dec = {}
    notas = _sim_notas_painel(loc, tipos, dias, sec, None, dec, hold)

    def nota(dia, t):
        txt_, ng, nv = dec[(dia, t)]
        if isinstance(ng, (int, float)) and isinstance(nv, (int, float)):
            return float(ng)
        v = notas.at[dia, t] if dia in notas.index else None
        return float(v) if isinstance(v, (int, float)) and pd.notna(v) else 0.0

    nome = lambda t: txt(por_t[t].get("codigo")) or t
    empresa = lambda t: txt(por_t[t].get("nome")) or nome(t)      # V154: as séries com o nome da empresa
    blocos = []
    for grp, rot_g in (("Ação", "Ações"), ("ETF", "ETFs")):
        todos = [t for t in cols if tipos[t] == grp]
        for rot, dia in datas:
            com = [t for t in todos if (dia, t) in dec]

            def classe(t):                             # V155: pelos limites desta folha
                txt_, ng, nv = dec[(dia, t)]
                tx = txt(txt_)
                if not (tx.startswith("Compra") or tx.startswith("Manter (nota global baixa")
                        or tx.startswith("Manter (qualidade / saúde fracas)")) or ": esperar" in tx:
                    return None
                k = (1 if grp == "ETF" else 0) if isinstance(nv, (int, float)) else 2
                v = nota(dia, t)
                return "forte" if v >= lim[k][0] else "media" if v >= lim[k][1] else None
            fortes = sorted([t for t in com if classe(t) == "forte"], key=lambda t: -nota(dia, t))
            medias = sorted([t for t in com if classe(t) == "media"], key=lambda t: -nota(dia, t))
            nf_tot, nm_tot = len(fortes), len(medias)
            fortes, medias = fortes[:EVO_NF], medias[:EVO_NM]
            seg = eur.loc[dia:].ffill()
            base = eur.loc[dia]
            ok = [t for t in todos if pd.notna(base[t]) and base[t] > 0]
            norm = seg[ok] / base[ok] * 100
            media_todos = norm.mean(axis=1)
            npts = len(seg)
            sel_m = [t for t in medias if t in ok]
            sel_f = [t for t in fortes if t in ok]
            # V154: sempre EVO_PTS pontos, do 1.º ao último dia (interpolados nos períodos curtos), para as linhas
            # ocuparem todo o comprimento do gráfico
            import numpy as np
            xs_ = np.linspace(0, npts - 1, EVO_PTS)
            interp = lambda serie: np.interp(xs_, np.arange(npts), serie.to_numpy(dtype=float))
            col_t = interp(media_todos)
            col_m = {t: interp(norm[t]) for t in sel_m}
            col_f = {t: interp(norm[t]) for t in sel_f}
            dados = []
            for i_, x_ in enumerate(xs_):
                k0 = int(np.floor(x_))
                k1 = min(k0 + 1, npts - 1)
                dt_ = seg.index[k0] + (seg.index[k1] - seg.index[k0]) * (x_ - k0)
                lin = [dt_.strftime("%Y-%m-%d"), round(float(col_t[i_]), 2)]
                vm = [float(col_m[t][i_]) for t in sel_m]
                vf = [float(col_f[t][i_]) for t in sel_f]
                lin += [round(v, 2) for v in vm] + [None] * (EVO_NM - len(sel_m))
                lin += [round(sum(vm) / len(vm), 2) if vm else None]
                lin += [round(v, 2) for v in vf] + [None] * (EVO_NF - len(sel_f))
                lin += [round(sum(vf) / len(vf), 2) if vf else None]
                dados.append(lin)
            fim_v = {t: float(norm[t].iloc[-1]) for t in sel_m + sel_f}
            med = lambda ts: sum(fim_v[t] for t in ts) / len(ts) if ts else None
            mdn = lambda ts: sorted(fim_v[t] for t in ts)[len(ts) // 2] if len(ts) % 2 else \
                (sum(sorted(fim_v[t] for t in ts)[len(ts) // 2 - 1:len(ts) // 2 + 1]) / 2 if ts else None)
            partes = []
            mf, mm, mt = med(sel_f), med(sel_m), float(media_todos.iloc[-1])
            partes.append(f"Compra forte: {nf_tot} título(s)" + (f", em média {_evo_pct(mf)} (mediana {_evo_pct(mdn(sel_f))})" if mf is not None else "")
                          + (f" (no gráfico os {EVO_NF} de maior nota)" if nf_tot > EVO_NF else ""))
            partes.append(f"Compra média: {len(sel_m)} de {nm_tot}" + (f", em média {_evo_pct(mm)} (mediana {_evo_pct(mdn(sel_m))})" if mm is not None else ""))
            tod = "Todas as ações" if grp == "Ação" else "Todos os ETFs"
            md_t = float(norm.iloc[-1].median())
            partes.append(f"{tod} ({len(ok)}): {_evo_pct(mt)} (mediana {_evo_pct(md_t)})")
            lst = lambda ts: ", ".join(f"{nome(t)} {_evo_pct(fim_v[t])}" for t in sorted(ts, key=lambda t: -fim_v[t]))
            lista = (("Forte: " + lst(sel_f)) if sel_f else "Forte: nenhum") + "   ·   " + \
                    (("Média: " + lst(sel_m)) if sel_m else "Média: nenhum")
            blocos.append({"titulo": f"{rot_g} · {rot} ({dia:%d/%m/%Y}) → {fim:%d/%m/%Y}",
                           "resumo": " · ".join(partes) + " (até hoje)", "lista": lista,
                           "nomes": ["Data", "Média de " + ("todas as ações" if grp == "Ação" else "todos os ETFs")]
                                    + [empresa(t) for t in sel_m] + [f"—m{i}" for i in range(EVO_NM - len(sel_m))]
                                    + ["Média das compras médias"]
                                    + [empresa(t) for t in sel_f] + [f"—f{i}" for i in range(EVO_NF - len(sel_f))]
                                    + ["Média das compras fortes"],
                           "dados": dados})
            print(f"   {rot_g:5} {dia:%d/%m/%Y}: forte {nf_tot:3d} ({_evo_pct(mf) if mf is not None else '—':>6}) · "
                  f"média {len(sel_m):2d}/{nm_tot:3d} ({_evo_pct(mm) if mm is not None else '—':>6}) · todos {_evo_pct(mt)}")
    gerado = f"Gerado em {datetime.now():%d/%m/%Y %H:%M} · cotações até {fim:%d/%m/%Y}"
    exc_txt = ", ".join(f"{txt(por_t[t].get('nome')) or t} ({t})" for t in excl)
    _gravar_json(FICH_EVO, {"v": EVO_V, "gerado": gerado, "blocos": blocos, "limites": lim, "excluidos": exc_txt})
    ops = [{"op": "escrever", "folha": FOLHA_EVO, "rng": "B6", "valores": [[gerado]]},
           {"op": "escrever", "folha": FOLHA_EVO, "rng": "B7",
            "valores": [["Excluídos (cotações suspeitas): " + exc_txt if exc_txt else None]]}]
    for b, bl in enumerate(blocos):
        r0, c0 = _evo_lin(b), _evo_col(b)
        ops.append({"op": "escrever", "folha": FOLHA_EVO, "rng": f"B{r0}:B{r0 + 1}",
                    "valores": [[bl["titulo"]], [bl["resumo"]]]})
        ops.append({"op": "escrever", "folha": FOLHA_EVO, "rng": f"B{r0 + 22}", "valores": [[bl["lista"]]]})
        grade = [bl["nomes"]] + [[datetime.strptime(l_[0], "%Y-%m-%d")] + l_[1:] for l_ in bl["dados"]]
        grade += [[None] * EVO_NC] * (EVO_PTS + 1 - len(grade))
        ops.append({"op": "escrever", "folha": FOLHA_EVO,
                    "rng": f"{get_column_letter(c0)}{EVO_L0 - 1}:{get_column_letter(c0 + EVO_NC - 1)}{EVO_L0 + EVO_PTS - 1}",
                    "valores": grade})
    ex.ops(ops, guardar=True)
    print("Evolução: pronto (folha Evolução).")



# ================================================================================================ Simulador
FOLHA_COMPRA = "Compra"
FICH_COMPRA = os.path.join(PASTA, BASE + "_compra.json")
COMPRA_PADRAO = {"min_a": 3.5, "min_e": 3.5, "total": 1000.0, "pct_a": 0.6, "so_compra": "x", "detidos": "x"}
CMP_L0 = 21                     # primeira linha da tabela de resultados da folha Compra (V163: 13 -> 19; V166: 21)
CMP_NFIX = 5                    # V163: títulos fixos por grupo (ticker + % do grupo); V166: 3 -> 5
CMP_FIX_L = 13                  # 1.ª linha dos fixos (ações: C nome, D ticker, E %; ETFs: G nome, H ticker, I %)
CMP_FIX_TK = "Ticker (automático)"   # V164: o nome escolhe-se numa lista; o ticker vem sozinho
CMP_MAX = 150
CMP_CAB = ("Não quero (x)", "Não disponível (x)", "Título", "Ticker", "Tipo", "Nota global", "Decisão", "Preço atual",
           "Compra máxima (ordem limite)", "Compra máxima / atual (%)", "Peso no grupo (%)", "Valor a investir (€)",
           "Nº de ações (aprox.)", "Tenho", "Câmbio para €", "Moeda", "Escala", "Visível", "Fixo", "Fixo: % do grupo")
# colunas da tabela (letras)
CQ, CND, CT, CTK, CTP, CNT, CDC, CPR, CMX, CMXP, CPS, CVL, CNA, CTEM, CFX, CMO, CESC, CVIS, CFIX, CFIXP = "ABCDEFGHIJKLMNOPQRST"


def _nota_txt_formula(cel, lims=None):
    """Fórmula que diz a que corresponde a nota global mínima (mesmos limites da decisão)."""
    a, b, c = lims or LIM_ACOES
    return (f'=IF(ISNUMBER({cel}),IF({cel}>={a},"= Compra forte",'
            f'IF({cel}>={b},"= Compra média",IF({cel}>={c},'
            f'"= Compra","= abaixo de Compra (Manter): entram títulos que o painel não aconselha comprar"))),"")')


def folha_compra(wb):
    """Folha Compra (entre ETFs e Procurar): quanto comprar de cada título, a partir de um valor total, da divisão
    ações/ETFs e da nota global mínima de cada grupo. Valor de cada título proporcional à nota global."""
    p = {**COMPRA_PADRAO, **{k: v for k, v in _ler_json(FICH_COMPRA, {}).items() if v not in (None, "")}}
    ws = wb.create_sheet(FOLHA_COMPRA, 2)
    ws["A1"] = "Compra: quanto investir em cada título"
    ws["A1"].font = Font(size=14, bold=True)
    ws["C2"] = ("Escolhe a amarelo e carrega em ▶ CALCULAR COMPRA (só lê as folhas Ações e ETFs, não as muda). Entram "
                "os títulos com nota global igual ou maior ao mínimo de cada grupo (e, com x, só os que têm decisão de "
                "compra: o sinal ≥ 3). O valor de cada grupo divide-se pelos títulos em proporção da nota global (nota "
                "4 recebe o dobro de nota 2). \"Compra máxima\" é o preço até onde vale a pena comprar (ordem limite), "
                "das folhas Ações/ETFs; os preços estão na moeda de cada título. Nº de ações = valor ÷ preço atual em "
                "euros (aproximado). Para tirar um título, põe x na coluna A («Não quero») ou B («Não disponível», p. "
                "ex. a corretora não o tem): fica riscado, com 0 €, e os valores dos outros do mesmo grupo recalculam-se "
                "logo (também ao mudar o valor total ou a % em ações). Os x ficam guardados para os próximos cálculos. "
                "As linhas escondidas com os filtros do cabeçalho também ficam com 0 € (o valor passa para as que se veem). "
                "Decisão a lilás («Compra …: esperar»): título bom e barato cujo momento ainda não é de compra; entra "
                "como os outros (com a nota global da tabela) e a decisão diz o que esperar (preço acima da média 200 "
                "ou 50 dias e/ou até à compra máxima): usa a ordem limite ou põe x em «Não quero» se preferires esperar. "
                "FIXOS (linhas 13-17): até 5 ações e 5 ETFs que entram sempre (mesmo sem cumprir os critérios), com a % "
                "escrita do valor do seu grupo (ex.: 60% dos ETFs no S&P 500); o resto do grupo divide-se pela nota global "
                "pelos outros. Escolhe o nome na lista (pela ordem das folhas Ações / ETFs): o ticker aparece sozinho.")
    ws["C2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["C2"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("C2:N2")
    ws.row_dimensions[2].height = 100
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    borda = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    entradas = (("C4", "Ações: nota global igual ou maior a", "D4", p["min_a"], "0.00"),
                ("C5", "ETFs: nota global igual ou maior a", "D5", p["min_e"], "0.00"),
                ("C6", "Valor total a investir (€)", "D6", p["total"], '#,##0.00 "€"'),
                ("C7", "Distribuição: % em ações", "D7", p["pct_a"], "0%"),
                ("C8", "Só com decisão de compra (sinal ≥ 3) (x)", "D8", p["so_compra"], "@"),
                ("C9", "Incluir os que já tenho (x)", "D9", p["detidos"], "@"))
    for cb, rot, cv, v, fm in entradas:
        ws[cb] = rot
        ws[cb].font = Font(bold=True)
        ws[cb].alignment = Alignment(horizontal="right")
        ws[cv] = v
        ws[cv].fill, ws[cv].border, ws[cv].number_format = amarelo, borda, fm
        ws[cv].font = Font(bold=True, size=12)
        ws[cv].alignment = Alignment(horizontal="center")
    ws["E4"], ws["E5"] = _nota_txt_formula("D4"), _nota_txt_formula("D5", LIM_ETFS)
    ws["E6"] = '=IF(ISNUMBER(D6),"Ações: "&FIXED(D6*D7,2)&" € · ETFs: "&FIXED(D6*(1-D7),2)&" €","")'
    ws["E7"] = '=IF(ISNUMBER(D7),"ETFs: "&TEXT(1-D7,"0%"),"")'
    for c_ in ("E4", "E5", "E6", "E7"):
        ws[c_].font = Font(bold=True, color="7030A0")
    ws["E8"] = "Sem x: entram também os de nota alta com o sinal ainda fraco (a decisão diz porquê)."
    ws["E9"] = "Sem x: só títulos que ainda não tens."
    for c_ in ("E8", "E9"):
        ws[c_].font = Font(italic=True, size=9, color="808080")
    botao(ws, "K4", "▶ CALCULAR COMPRA", "painelacoes://compra", "548235", "375623", 14)
    ws.merge_cells("K4:M5")
    ws["K7"] = ("Escalas da nota global: ações ≥ 3,95 forte · ≥ 3,5 média · ≥ 3 compra; ETFs sem nível de preço "
                "(ouro, obrigações, ou sem ações conhecidas) ≥ 4,5 forte · ≥ 3,75 média · ≥ 2,75 compra")
    ws["K7"].font = Font(size=9, color="595959")
    # --- fixos (V163) ---
    ws["C11"] = "Fixos: entram sempre, com esta % do valor do grupo (o resto divide-se pela nota global)"
    ws["C11"].font = Font(bold=True, size=11, color="1F4E79")
    la, lt = get_column_letter(COL["nome"]), get_column_letter(COL["ticker"])
    nome_de = {}                                  # ticker -> nome, das folhas Ações / ETFs já feitas
    for fl_ in (FOLHA_A2, FOLHA_E2):
        if fl_ in wb.sheetnames:
            w_ = wb[fl_]
            for r_ in range(L0 + 1, L0 + MAX_LINHAS + 1):
                t_, n_ = txt(w_.cell(r_, COL["ticker"]).value).upper(), w_.cell(r_, COL["nome"]).value
                if isinstance(n_, str) and n_.startswith("="):          # =HYPERLINK(…,"Nome")
                    m_ = re.search(r',"((?:[^"]|"")*)"\)$', n_)
                    n_ = m_.group(1).replace('""', '"') if m_ else None
                if t_ and isinstance(n_, str) and n_:
                    nome_de.setdefault(t_, n_)
    from openpyxl.worksheet.datavalidation import DataValidation as DV_
    lng, lnt, lnv = (get_column_letter(COL[k_]) for k_ in ("nota_geral", "nota", "nivel_preco"))
    from openpyxl.formatting.rule import FormulaRule as FRn_
    for (cn, ct, cnt_, cp_, rot, k, fl_) in (("C", "D", "E", "F", "Ações fixas", "fixos_a", FOLHA_A2),
                                             ("G", "H", "I", "J", "ETFs fixos", "fixos_e", FOLHA_E2)):
        dv = DV_(type="list", formula1=f"='{fl_}'!${la}${L0 + 1}:${la}${L0 + MAX_LINHAS}", allow_blank=True)
        dv.error, dv.errorTitle = f"Escolhe um nome da lista (folha {fl_}).", "Fixos"
        ws.add_data_validation(dv)
        dv.add(f"{cn}{CMP_FIX_L}:{cn}{CMP_FIX_L + CMP_NFIX - 1}")
        for c_, h_ in ((cn, rot + " (escolhe na lista)"), (ct, CMP_FIX_TK), (cnt_, "Nota global"), (cp_, "% do grupo")):
            c = ws[f"{c_}{CMP_FIX_L - 1}"]
            c.value = h_
            c.font, c.fill = Font(bold=True, color="FFFFFF", size=9), PatternFill("solid", fgColor="1F4E79")
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        fx_ = p.get(k) or []
        for i in range(CMP_NFIX):
            r = CMP_FIX_L + i
            tk_, pc_, nm_ = (list(fx_[i]) + [None, None, None])[:3] if i < len(fx_) else (None, None, None)
            nm_ = nome_de.get(txt(tk_).upper()) or nm_
            for c_, v_, fm, al in ((cn, nm_, "@", "left"), (cp_, pc_, "0.0%", "center")):
                c = ws[f"{c_}{r}"]
                c.value, c.fill, c.border, c.number_format = v_, amarelo, borda, fm
                c.font, c.alignment = Font(bold=True), Alignment(horizontal=al)
            ws[f"{ct}{r}"] = (f'=IF(TRIM({cn}{r})="","",IFERROR(INDEX(\'{fl_}\'!${lt}:${lt},MATCH({cn}{r},\'{fl_}\'!${la}:${la},0)),'
                              f'"não está na folha {fl_}"))')
            ws[f"{ct}{r}"].font = Font(color="1F4E79", size=9)
            ws[f"{ct}{r}"].alignment = Alignment(horizontal="center")
            mt = f"MATCH({cn}{r},'{fl_}'!${la}:${la},0)"        # V167: nota global atual, com as cores da tabela
            ws[f"{cnt_}{r}"] = (f'=IF(TRIM({cn}{r})="","",IFERROR(IF(ISNUMBER(INDEX(\'{fl_}\'!${lng}:${lng},{mt})),'
                                f'INDEX(\'{fl_}\'!${lng}:${lng},{mt}),INDEX(\'{fl_}\'!${lnt}:${lnt},{mt})),""))')
            ws[f"{cnt_}{r}"].number_format, ws[f"{cnt_}{r}"].font = "0.00", Font(bold=True)
            ws[f"{cnt_}{r}"].alignment = Alignment(horizontal="center")
        rr_ = f"{cnt_}{CMP_FIX_L}:{cnt_}{CMP_FIX_L + CMP_NFIX - 1}"
        c0_ = f"{cnt_}{CMP_FIX_L}"
        esc_ = f"ISNUMBER(INDEX('{fl_}'!${lnv}:${lnv},MATCH(${cn}{CMP_FIX_L},'{fl_}'!${la}:${la},0)))"
        for i_, cor in enumerate(("63BE7B", "A9D18E", "E2EFDA", "F4B183", "FFEB9C", "F8696B")):
            la_, le_ = (LIM_ACOES + (2.25, 1.5, -99))[i_], (LIM_ETFS + (2.25, 1.5, -99))[i_]
            ws.conditional_formatting.add(rr_, FRn_(
                formula=[f"AND(ISNUMBER({c0_}),{c0_}>=IF({esc_},{la_},{le_}))"],
                fill=PatternFill("solid", bgColor=cor), stopIfTrue=True))
        rs = CMP_FIX_L + CMP_NFIX
        sm = "+".join(f"IF(N({cp_}{r})>1,N({cp_}{r})/100,N({cp_}{r}))" for r in range(CMP_FIX_L, rs))   # 50 = 50%
        ws[f"{cn}{rs}"] = (f'=IF(({sm})=0,"",IF(({sm})>1,'
                           f'"Soma acima de 100%: os fixos ficam com o grupo todo (em proporção)",'
                           f'"Soma dos fixos (o resto, "&TEXT(1-({sm}),"0.0%")&", vai pela nota)"))')
        ws[f"{cn}{rs}"].font = Font(italic=True, size=9, color="7030A0")
        ws[f"{cn}{rs}"].alignment = Alignment(horizontal="right")
        ws[f"{cp_}{rs}"] = f"={sm}"
        ws[f"{cp_}{rs}"].number_format, ws[f"{cp_}{rs}"].font = "0.0%", Font(bold=True, color="7030A0")
        ws[f"{cp_}{rs}"].alignment = Alignment(horizontal="center")
    ws[f"C{CMP_L0 - 2}"] = "Resultado"
    ws[f"C{CMP_L0 - 2}"].font = Font(bold=True, size=12)
    for j, h in enumerate(CMP_CAB, 1):
        c = ws.cell(CMP_L0 - 1, j, h)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
        c.alignment, c.border = Alignment(horizontal="center", vertical="center", wrap_text=True), borda
    ws.row_dimensions[CMP_L0 - 1].height = 60
    fmts = {CNT: "0.00", CPR: "#,##0.00", CMX: "#,##0.00", CMXP: "0.0%", CPS: "0.0%", CVL: '#,##0.00 "€"', CNA: "#,##0.00"}
    for r in range(CMP_L0, CMP_L0 + CMP_MAX):
        for col, fm in fmts.items():
            ws[f"{col}{r}"].number_format = fm
        for col in (CQ, CND, CTEM):
            ws[f"{col}{r}"].alignment = Alignment(horizontal="center")
        ws[f"{CMX}{r}"].font = Font(bold=True, size=12)          # compra máxima (ordem limite) em destaque
    for col, w in zip("ABCDEFGHIJKLMN", (5, 8, 34, 11, 7, 8, 30, 12, 12, 10, 9, 13, 10, 7)):
        ws.column_dimensions[col].width = w
    ws.column_dimensions[CFX].hidden = True
    ws.column_dimensions[CMO].hidden = True
    ws.column_dimensions[CESC].hidden = True       # 1 = com nível de preço (limites 3,95/3,5/3), senão 4,5/3,75/2,75
    ws.column_dimensions[CVIS].hidden = True       # 1 = linha à vista; 0 = escondida pelo filtro (fica sem valor)
    ws.column_dimensions[CFIX].hidden = True       # V163: 1 = fixo
    ws.column_dimensions[CFIXP].hidden = True      # V163: % do grupo do fixo (0 se tirado com x / escondido)
    ws.auto_filter.ref = f"A{CMP_L0 - 1}:{CESC}{CMP_L0 + CMP_MAX - 1}"   # filtros: o que se esconde fica a 0 €
    a0, an = CMP_L0, CMP_L0 + CMP_MAX
    from openpyxl.formatting.rule import Rule as Rule_, FormulaRule as FR
    from openpyxl.styles.differential import DifferentialStyle as DS_
    from openpyxl.styles.numbers import NumberFormat as NF_
    for i, m in enumerate(SIMBOLO):                   # preços com o símbolo da moeda da linha (como nas folhas gerais)
        ws.conditional_formatting.add(f"{CPR}{a0}:{CMX}{an}", Rule_(
            type="expression", dxf=DS_(numFmt=NF_(numFmtId=400 + i, formatCode=fmt_moeda(m))),
            formula=[f'${CMO}{a0}="{m}"']))
    ws.conditional_formatting.add(f"{CQ}{a0}:{CND}{an}", FR(          # colunas A e B: amarelas nos títulos
        formula=[f'AND(${CTK}{a0}<>"",LEFT(${CT}{a0},5)<>"Total")'], fill=PatternFill("solid", bgColor="FFF2CC")))
    ws.conditional_formatting.add(f"{CT}{a0}:{CTEM}{an}", FR(          # excluído: cinzento riscado
        formula=[f'OR(LOWER(TRIM(${CQ}{a0}))="x",LOWER(TRIM(${CND}{a0}))="x")'], font=Font(color="A6A6A6", strike=True),
        stopIfTrue=True))
    for col_, cond_ in ((CNT, f'AND(ISNUMBER(${CNT}{a0}),${CNT}{a0}>={{lim}})'),
                        (CDC, f'AND(${CDC}{a0}<>"",ISNUMBER(${CNT}{a0}),${CNT}{a0}>={{lim}})')):
        if col_ == CDC:
            continue                                  # a decisão tem as suas regras mais abaixo
        for i_, cor in enumerate(("63BE7B", "A9D18E", "E2EFDA", "F4B183", "FFEB9C", "F8696B")):
            la, le = (LIM_ACOES + (2.25, 1.5, -99))[i_], (LIM_ETFS + (2.25, 1.5, -99))[i_]
            lim = f'IF(${CESC}{a0}=1,{la},{le})'
            ws.conditional_formatting.add(f"{col_}{a0}:{col_}{an}", FR(
                formula=[cond_.format(lim=lim)], fill=PatternFill("solid", bgColor=cor), stopIfTrue=True))
    ws.conditional_formatting.add(f"{CT}{a0}:{CTEM}{an}", FR(
        formula=[f'OR(LEFT(${CT}{a0},5)="Total",LEFT(${CT}{a0},5)="Lista")'], font=Font(bold=True),
        fill=PatternFill("solid", bgColor="DDEBF7")))
    # decisão: as mesmas cores das folhas Ações / ETFs
    dcl = f"${CDC}{a0}"
    for f_, cor, fnt in ((f'ISNUMBER(SEARCH(": esperar",{dcl}))', COR_ESPERA, Font(color="5B2C83", bold=True)),
                         (f'ISNUMBER(SEARCH("saúde frágil",{dcl}))', "F8CBAD", Font(color="C00000", bold=True)),
                         (f'LEFT({dcl},5)="Venda"', "F8696B", Font(bold=True)),
                         (f'ISNUMBER(SEARCH("Longo Prazo",{dcl}))', "F4B183", None),
                         (f'LEFT({dcl},6)="Manter"', "FFEB9C", None)):
        ws.conditional_formatting.add(f"{CDC}{a0}:{CDC}{an}", FR(formula=[f_], fill=PatternFill("solid", bgColor=cor),
                                                                font=fnt, stopIfTrue=True))
    for txt_, cor, b_ in (("Compra forte", "63BE7B", True), ("Compra média", "A9D18E", False), ("Compra", "E2EFDA", False)):
        ws.conditional_formatting.add(f"{CDC}{a0}:{CDC}{an}", FR(          # V148: pela decisão (forte só no momento certo)
            formula=[f'LEFT({dcl},{len(txt_)})="{txt_}"'], fill=PatternFill("solid", bgColor=cor), font=Font(bold=b_),
            stopIfTrue=True))
    gj = _ler_json(FICH_COMPRA, {})
    res = (gj.get("linhas2") or []) if gj.get("l0") == CMP_L0 else []     # V163: resultados de antes (outra linha) não
    for i, lin in enumerate(res[:CMP_MAX]):
        for j, v in enumerate(lin, 1):
            ws.cell(CMP_L0 + i, j, v)
    ws.freeze_panes = None                         # sem painéis fixos


def _num_ou(v, padrao=None):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    try:
        return float(str(v).replace("€", "").replace("%", "").replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return padrao


def calcular_compra():
    """▶ CALCULAR COMPRA (folha Compra)."""
    ex = Excel(FICHEIRO)
    r = ex.ops([{"op": "ler", "folha": FOLHA_COMPRA, "rng": "D4:D9"}])
    vals = [lin[0] for lin in r[0]]
    p = dict(COMPRA_PADRAO)
    for k, v in zip(("min_a", "min_e", "total", "pct_a"), vals[:4]):
        x = _num_ou(v)
        if x is not None:
            p[k] = x
    if p["pct_a"] > 1:
        p["pct_a"] /= 100
    p["so_compra"] = "x" if txt(vals[4]).lower() == "x" else ""
    p["detidos"] = "x" if txt(vals[5]).lower() == "x" else ""
    g = _ler_json(FICH_COMPRA, {})
    fl = CMP_FIX_L + CMP_NFIX - 1                 # V163: fixos (ticker, % do grupo)
    try:
        cab_fx = txt(ex.ops([{"op": "ler", "folha": FOLHA_COMPRA, "rng": f"D{CMP_FIX_L - 1}"}])[0][0][0])
        cab_t = txt(ex.ops([{"op": "ler", "folha": FOLHA_COMPRA, "rng": f"A{CMP_L0 - 1}"}])[0][0][0])
    except Exception:
        cab_fx = cab_t = ""
    try:
        cab_p = txt(ex.ops([{"op": "ler", "folha": FOLHA_COMPRA, "rng": f"F{CMP_FIX_L - 1}"}])[0][0][0])
    except Exception:
        cab_p = ""
    if cab_t != CMP_CAB[0] or cab_p != "% do grupo":   # V166/V167: folha Compra de outra versão
        raise SystemExit("A folha Compra é de outra versão do script: carrega primeiro em ▶ ATUALIZAR e depois em "
                         "▶ CALCULAR COMPRA (as tuas escolhas ficam guardadas).")
    if cab_fx in ("Ticker", CMP_FIX_TK):          # V163: ticker escrito; V164: nome escolhido na lista (C/G)
        nv = cab_fx == CMP_FIX_TK
        rf = ex.ops([{"op": "ler", "folha": FOLHA_COMPRA, "rng": f"{'C' if nv else 'D'}{CMP_FIX_L}:F{fl}"},
                     {"op": "ler", "folha": FOLHA_COMPRA, "rng": f"{'G' if nv else 'H'}{CMP_FIX_L}:J{fl}"}])
        for k, tab_ in zip(("fixos_a", "fixos_e"), rf):
            lst = []
            for lin in tab_:
                nm_ = txt(lin[0]).strip() if nv else ""
                tk_ = txt(lin[1] if nv else lin[0]).strip().upper()
                if tk_.startswith("NÃO ESTÁ") or tk_.startswith("="):   # sem valor (fórmula por calcular): vai pelo nome
                    tk_ = ""
                pc_ = _num_ou(lin[-1])
                if (tk_ or nm_) and pc_ is not None and pc_ > 0:
                    lst.append([tk_, round(pc_ / 100 if pc_ > 1 else pc_, 4), nm_])
            p[k] = lst
    else:                                         # folha antiga (sem fixos): mantém os guardados
        p["fixos_a"], p["fixos_e"] = g.get("fixos_a") or [], g.get("fixos_e") or []
    excluir, indisp = set(g.get("excluir") or []), set(g.get("indisp") or [])
    try:                                          # o que está marcado agora nas colunas A e B
        t_ = ex.ops([{"op": "ler", "folha": FOLHA_COMPRA, "rng": f"A{CMP_L0}:D{CMP_L0 + CMP_MAX - 1}"}])[0]
        cab = txt(ex.ops([{"op": "ler", "folha": FOLHA_COMPRA, "rng": f"B{CMP_L0 - 1}"}])[0][0][0])
        if cab_fx not in ("Ticker", CMP_FIX_TK):  # V163: folha antiga (tabela noutra linha): fica com os x guardados
            raise ValueError("folha antiga")
        if cab.startswith("Não disponível"):      # folha já com a coluna B (V120)
            vistos = {txt(l_[3]).upper() for l_ in t_ if txt(l_[3])}
            excluir = (excluir - vistos) | {txt(l_[3]).upper() for l_ in t_ if txt(l_[3]) and txt(l_[0]).lower() == "x"}
            indisp = (indisp - vistos) | {txt(l_[3]).upper() for l_ in t_ if txt(l_[3]) and txt(l_[1]).lower() == "x"}
        else:                                     # folha da V118/V119 (sem a coluna B): A = não quero, C = ticker
            vistos = {txt(l_[2]).upper() for l_ in t_ if txt(l_[2])}
            excluir = (excluir - vistos) | {txt(l_[2]).upper() for l_ in t_ if txt(l_[2]) and txt(l_[0]).lower() == "x"}
    except Exception:
        pass
    ultima = get_column_letter(len(COLS))
    r = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)])
    acoes, etfs = (list(linhas_por_cabecalho(tab, i)) for i, tab in zip((1, 2), r))
    ocultos = set()                               # linhas escondidas pelos filtros das folhas Ações / ETFs: ficam de fora
    try:                                          # (como na app: o que não se vê não entra na compra)
        ro = ex.ops([op_ocultos(1), op_ocultos(2)])
        ocultos = {txt(t).upper() for t in (ro[0] or []) + (ro[1] or [])}
    except BaseException:
        pass
    if ocultos:
        acoes = [d for d in acoes if txt(d.get("ticker")).upper() not in ocultos]
        etfs = [d for d in etfs if txt(d.get("ticker")).upper() not in ocultos]
        print(f"   {len(ocultos)} título(s) escondidos pelos filtros das folhas Ações / ETFs ficam de fora da compra.")
    linhas, avisos, espera_l = [], [], []
    L1_, LN_ = CMP_L0, CMP_L0 + CMP_MAX - 1
    R = lambda c, fim=LN_: f"${c}${L1_}:${c}${fim}"
    por_tk, por_nm = {}, {}
    for d in acoes + etfs:
        por_tk.setdefault(txt(d.get("ticker")).upper(), d)
    for d in (acoes, etfs):
        for d_ in d:
            por_nm.setdefault((d is acoes, txt(d_.get("nome")).strip().lower()), d_)
    for k, e_a in (("fixos_a", True), ("fixos_e", False)):   # V164: sem ticker (nome da lista): procura pelo nome
        for f_ in p.get(k) or []:
            if not f_[0] and len(f_) > 2 and f_[2]:
                d_ = por_nm.get((e_a, f_[2].strip().lower())) or por_nm.get((not e_a, f_[2].strip().lower()))
                f_[0] = txt(d_.get("ticker")).upper() if d_ else f_[2]
            if len(f_) > 2 and not f_[2] and por_tk.get(f_[0]):
                f_[2] = txt(por_tk[f_[0]].get("nome"))
    for nome_g, grupo, mn, valor, k_fx, c_fx in (("Ações", acoes, p["min_a"], p["total"] * p["pct_a"], "fixos_a", "F"),
                                                 ("ETFs", etfs, p["min_e"], p["total"] * (1 - p["pct_a"]), "fixos_e", "J")):
        fixos = []                                # V163: (d, linha da % na folha)
        for i_, (tk_, pc_, *_) in enumerate(p.get(k_fx) or []):
            d_ = por_tk.get(tk_)
            if d_ is None:
                avisos.append(f"{nome_g}: o fixo {tk_} não está nas folhas Ações / ETFs (ou está escondido pelos filtros); "
                              "fica de fora.")
                continue
            if all(d_ is not f_[0] for f_ in fixos):
                fixos.append((d_, CMP_FIX_L + i_))
        tk_fix = {txt(f_[0].get("ticker")).upper() for f_ in fixos}
        esc = []
        for d in grupo:
            if txt(d.get("ticker")).upper() in tk_fix:
                continue
            ng = d.get("nota_geral") if isinstance(d.get("nota_geral"), (int, float)) else d.get("nota")
            dc_ = txt(d.get("decisao"))               # V147: a nota global da tabela, também nos «esperar»
            if not isinstance(ng, (int, float)) or round(ng, 2) < mn - 1e-9:
                continue
            if p["so_compra"] and not dc_.startswith("Compra"):
                continue
            tem = e_detido(d.get("det")) or bool(d.get("qtd"))
            if tem and not p["detidos"]:
                continue
            esc.append((ng, d, tem))
        esc.sort(key=lambda x: -x[0])
        tipo = "Ação" if nome_g == "Ações" else "ETF"
        val_g = "$D$6*$D$7" if nome_g == "Ações" else "$D$6*(1-$D$7)"
        if not esc and not fixos:
            avisos.append(f"{nome_g}: nenhum título cumpre os critérios; os € {valor:,.2f} ficam por investir "
                          "(baixa a nota mínima ou tira o x de «só com decisão de compra».)".replace(",", " "))
            continue
        soma_fx = f'SUMIFS({R(CFIXP)},{R(CTP)},"{tipo}")'
        for d, lin_pct in fixos:                  # V163: fixos primeiro (a % da folha, em proporção se passar de 100%)
            r_ = CMP_L0 + len(linhas)
            ng = d.get("nota_geral") if isinstance(d.get("nota_geral"), (int, float)) else d.get("nota")
            pr = d.get("preco")
            fx = 1.0 if txt(d.get("moeda")) == "EUR" else (d.get("fx_alt") if txt(d.get("moeda_alt")) == "EUR" else None)
            tk = txt(d.get("ticker"))
            fora = f'OR(LOWER(TRIM(${CQ}{r_}))="x",LOWER(TRIM(${CND}{r_}))="x")'
            tem = e_detido(d.get("det")) or bool(d.get("qtd"))
            linhas.append(["x" if tk.upper() in excluir else "", "x" if tk.upper() in indisp else "",
                           f"📌 {txt(d.get('nome')) or tk} (fixo)", tk, tipo,
                           round(ng, 2) if isinstance(ng, (int, float)) else None, txt(d.get("decisao")), pr, d.get("teto"),
                           d.get("teto_pct"), f"=IFERROR(${CFIXP}{r_}/MAX(1,{soma_fx}),0)", f"={CPS}{r_}*{val_g}",
                           f'=IF(AND(ISNUMBER({CPR}{r_}),ISNUMBER({CFX}{r_}),{CPR}{r_}>0,{CVL}{r_}>0),'
                           f'{CVL}{r_}/({CPR}{r_}*{CFX}{r_}),"")',
                           "★" if tem else "", fx if isinstance(fx, (int, float)) else None, txt(d.get("moeda")),
                           1 if isinstance(d.get("nivel_preco"), (int, float)) else 0, f"=SUBTOTAL(103,${CT}{r_})",
                           1, f'=IF(OR({fora},${CVIS}{r_}=0),0,IF(N(${c_fx}${lin_pct})>1,N(${c_fx}${lin_pct})/100,N(${c_fx}${lin_pct})))'])
        if not esc:
            avisos.append(f"{nome_g}: além dos fixos, nenhum título cumpre os critérios; o resto do grupo fica por investir.")
        for ng, d, tem in esc:
            r_ = CMP_L0 + len(linhas)
            pr = d.get("preco")
            fx = 1.0 if txt(d.get("moeda")) == "EUR" else (d.get("fx_alt") if txt(d.get("moeda_alt")) == "EUR" else None)
            tk = txt(d.get("ticker"))
            fora = f'OR(LOWER(TRIM(${CQ}{r_}))="x",LOWER(TRIM(${CND}{r_}))="x")'
            peso = (f'=IF(OR({fora},${CVIS}{r_}=0),0,IFERROR(${CNT}{r_}/SUMIFS({R(CNT)},{R(CTP)},"{tipo}",{R(CQ)},"<>x",'
                    f'{R(CND)},"<>x",{R(CVIS)},1,{R(CFIX)},0),0)*MAX(0,1-{soma_fx}))')   # escondido = 0; V163: menos os fixos
            linhas.append(["x" if tk.upper() in excluir else "", "x" if tk.upper() in indisp else "",
                           txt(d.get("nome")) or tk, tk, tipo, round(ng, 2), txt(d.get("decisao")), pr, d.get("teto"),
                           d.get("teto_pct"), peso, f"={CPS}{r_}*{val_g}",
                           f'=IF(AND(ISNUMBER({CPR}{r_}),ISNUMBER({CFX}{r_}),{CPR}{r_}>0,{CVL}{r_}>0),'
                           f'{CVL}{r_}/({CPR}{r_}*{CFX}{r_}),"")',
                           "★" if tem else "", fx if isinstance(fx, (int, float)) else None, txt(d.get("moeda")),
                           1 if isinstance(d.get("nivel_preco"), (int, float)) else 0, f"=SUBTOTAL(103,${CT}{r_})", 0, 0])
        r_ = CMP_L0 + len(linhas)
        lin_t = [""] * len(CMP_CAB)
        lin_t[2] = f"Total {'ações' if nome_g == 'Ações' else 'ETFs'}"
        lin_t[10] = f'=SUMIFS({R(CPS, r_ - 1)},{R(CTP, r_ - 1)},"{tipo}")'
        lin_t[11] = f'=SUMIFS({R(CVL, r_ - 1)},{R(CTP, r_ - 1)},"{tipo}")'
        linhas.append(lin_t)
        linhas.append([""] * len(CMP_CAB))
    if espera_l:                                  # lista de espera (não entra nos valores)
        lin_t = [""] * len(CMP_CAB)
        lin_t[2] = "Lista de espera: boas e baratas, mas ainda não é altura (comprar quando cumprir o que está na decisão)"
        linhas.append(lin_t)
        for ng, d, nome_g in sorted(espera_l, key=lambda x: -x[0]):
            pr = d.get("preco")
            tk = txt(d.get("ticker"))
            linhas.append(["", "x" if tk.upper() in indisp else "", txt(d.get("nome")) or tk, tk,
                           ("Ação" if nome_g == "Ações" else "ETF") + " (espera)", round(ng, 2), txt(d.get("decisao")), pr,
                           d.get("teto"), d.get("teto_pct"), None, None, None,
                           "★" if (e_detido(d.get("det")) or bool(d.get("qtd"))) else "", None, txt(d.get("moeda"))])
        linhas.append([""] * len(CMP_CAB))
    for a in avisos:
        linhas.append(["", "", a] + [""] * (len(CMP_CAB) - 3))
    _gravar_json(FICH_COMPRA, {**p, "excluir": sorted(excluir), "indisp": sorted(indisp), "linhas2": linhas[:CMP_MAX],
                               "l0": CMP_L0, "ts": int(time.time() * 1000)})          # a app usa as escolhas mais recentes
    vazio = [[None] * len(CMP_CAB) for _ in range(CMP_MAX)]
    for i, lin in enumerate(linhas[:CMP_MAX]):
        vazio[i] = ([None if v == "" else v for v in lin] + [None] * len(CMP_CAB))[:len(CMP_CAB)]
    ops_ = [{"op": "escrever", "folha": FOLHA_COMPRA, "rng": f"A{CMP_L0}:{get_column_letter(len(CMP_CAB))}{CMP_L0 + CMP_MAX - 1}",
             "valores": vazio}]
    if cab_fx in ("Ticker", CMP_FIX_TK):          # fixos arrumados (sem buracos, % como fração)
        for k, cn_, cp2_, ct_ in (("fixos_a", "C", "F", "D"), ("fixos_e", "G", "J", "H")):
            lst = [list(x) + [None] * 3 for x in (p.get(k) or [])] + [[None, None, None]] * CMP_NFIX
            rng_ = lambda c: f"{c}{CMP_FIX_L}:{c}{CMP_FIX_L + CMP_NFIX - 1}"
            if cab_fx == CMP_FIX_TK:
                ops_.append({"op": "escrever", "folha": FOLHA_COMPRA, "rng": rng_(cn_), "valores": [[x[2]] for x in lst[:CMP_NFIX]]})
            else:
                ops_.append({"op": "escrever", "folha": FOLHA_COMPRA, "rng": rng_(ct_), "valores": [[x[0]] for x in lst[:CMP_NFIX]]})
            ops_.append({"op": "escrever", "folha": FOLHA_COMPRA, "rng": rng_(cp2_), "valores": [[x[1]] for x in lst[:CMP_NFIX]]})
    ex.ops(ops_, guardar=True)
    n = sum(1 for lin in linhas if lin[3])
    if excluir or indisp:
        print(f"   Fora: {len(excluir)} «não quero» e {len(indisp)} «não disponível» (o valor vai para os outros)")
    print(f"Compra: {n} títulos; ações € {p['total'] * p['pct_a']:,.2f}, ETFs € {p['total'] * (1 - p['pct_a']):,.2f}."
          .replace(",", " "))
    for a in avisos:
        print("   " + a)
    enviar_compra_github()                         # a app do telemóvel passa a usar estas escolhas


FOLHA_HIS = "Histórico"
FICH_HIS = os.path.join(PASTA, BASE + "_historico.json")
HIS_N, HIS_L0, HIS_MAX = 10, 40, 430          # títulos; 1.ª linha dos dados (colunas escondidas); pontos no máximo
HIS_ANOS_MAX = 5
HIS_SER = (("forte", "Compra forte", "17692C"), ("media", "Compra média", "4CAF50"), ("compra", "Compra", "B5DB8F"),
           ("esperar", "Compra, mas esperar", "8E6BBF"), ("manter", "Manter", "A6A6A6"),
           ("venda", "Venda", "E06666"))                                         # séries dos gráficos (6 cores)
HIS_COL0 = 27                                  # coluna AA: datas; depois 6 séries por título
HIS_V = 4                                      # versão do ficheiro do histórico (4: 6 séries)
HIS_CLASSES = (("esperar", "Compra, mas esperar", "8E6BBF"), ("forte", "Compra forte", "1E7B34"), ("media", "Compra média", "63BE7B"), ("compra", "Compra", "A9D18E"),
               ("manter", "Manter", "A6A6A6"), ("venda", "Venda", "E06666"))


def _his_classe(dec):
    dec = txt(dec)
    if ": esperar" in dec:
        return "esperar"
    if dec.startswith("Venda"):
        return "venda"
    if dec.startswith("Compra forte"):
        return "forte"
    if dec.startswith("Compra média"):
        return "media"
    if dec.startswith("Compra"):
        return "compra"
    return "manter"


def folha_historico(wb, linhas=(), etfs=()):
    """Folha Histórico: a decisão do painel em cada dia do último ano (ou do período escolhido), para até 10
    títulos, calculada só com os dados conhecidos em cada dia (cotações; nos EUA também os relatórios da SEC)."""
    from openpyxl.chart import ScatterChart, Reference, Series
    g = _ler_json(FICH_HIS, {})
    tks = g.get("tickers") or [txt(d.get("ticker")) for d in list(linhas) + list(etfs) if e_detido(d.get("det"))][:HIS_N]
    ws = wb.create_sheet(FOLHA_HIS)
    ws["A1"] = "Histórico da classificação: o que o painel teria aconselhado em cada dia"
    ws["A1"].font = Font(size=14, bold=True)
    ws["B2"] = ("Escreve até 10 tickers (como na coluna \"Ticker Yahoo\" das folhas Ações e ETFs) e carrega em ▶ GERAR "
                "HISTÓRICO. Para cada dia do período, o painel calcula a decisão com as regras e pesos atuais, mas só com "
                "o que se sabia nesse dia: as cotações (médias, rede de segurança, retas, fator de tendência) e, nas ações "
                "dos EUA, os relatórios anuais já publicados na SEC (qualidade, saúde, PER, FCF, valor justo e nível de "
                "preço); nas europeias, os relatórios anuais ESEF (desde 2020, sem trimestres). Fora dos EUA e da Europa, sem nível de preço, a nota "
                "global é a nota geral e os limites são os da nota geral (4,5 / 3,75 / 2,75), por isso a compra forte é "
                "rara. Os analistas não têm histórico: o sinal é o da tendência. O lucro por ação é o dos últimos 4 trimestres "
                "já divulgados (10-Q e 10-K), como na tabela. O ÚLTIMO PONTO é a decisão da tabela (com analistas), para bater com ela. "
                "Período em anos em C4 (=4/52 para 4 semanas); um ponto a cada 3 a 5 dias de bolsa (C5), até 5 anos. Cores: verde escuro compra forte, verde compra média, "
                "verde claro compra, lilás «compra, mas esperar» (boa e barata, mas a tendência ainda não virou), cinzento "
                "manter, vermelho venda.")
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B2"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("B2:N2")
    ws.row_dimensions[2].height = 80
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    borda = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
    ws["B4"], ws["C4"] = f"Período em anos (até {HIS_ANOS_MAX})", g.get("anos") or 1      # V156: letra legível
    ws["B5"], ws["C5"] = "Dias de bolsa por ponto (3 a 21)", g.get("passo") or 5
    for c_ in ("B4", "B5"):
        ws[c_].font = Font(bold=True, size=9)
        ws[c_].alignment = Alignment(wrap_text=False, vertical="center")
    ws.row_dimensions[4].height = ws.row_dimensions[5].height = 14     # linhas baixas
    for c_ in ("C4", "C5"):
        ws[c_].fill, ws[c_].border, ws[c_].font = amarelo, borda, Font(bold=True, size=9)
        ws[c_].alignment = Alignment(horizontal="center", vertical="center")
    ws["C4"].number_format = "0.00"
    ws["D4"] = '=IF(ISNUMBER(C4),TEXT(C4*52,"0")&" semanas","")'
    ws["D4"].font = Font(size=9, color="7030A0")
    botao(ws, "K4", "▶ GERAR HISTÓRICO", "painelacoes://historico", "1F4E79", "16365C", 10)
    ws.merge_cells("K4:M5")
    cab = ("Ticker", "Título", "Pontos em compra: esperar", "Compra forte", "Compra média", "Compra", "Manter", "Venda",
           "Decisão no último ponto")
    for j, h in enumerate(cab, 2):
        c = ws.cell(6, j, h)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
        c.alignment, c.border = Alignment(horizontal="center", vertical="center", wrap_text=True), borda
    ws.row_dimensions[6].height = 30
    for i in range(HIS_N):
        c = ws.cell(7 + i, 2, tks[i] if i < len(tks) else None)
        c.fill, c.border, c.font = amarelo, borda, Font(bold=True)
        for j, v in enumerate(((g.get("resumo") or {}).get(txt(c.value).upper(), []) if g.get("v") in (3, HIS_V) else [])[:8], 3):
            ws.cell(7 + i, j, v)
    for col, w in zip("BCDEFGHIJ", (30, 30, 10, 10, 10, 10, 10, 10, 34)):
        ws.column_dimensions[col].width = w

    # dados (escondidos) e gráficos
    NS = len(HIS_SER)
    ws.cell(HIS_L0 - 1, HIS_COL0, "Data")
    for k in range(HIS_N):
        for m, (_, nome, _) in enumerate(HIS_SER):
            ws.cell(HIS_L0 - 1, HIS_COL0 + 1 + NS * k + m, nome)
    for c in range(HIS_COL0, HIS_COL0 + 1 + NS * HIS_N):
        ws.column_dimensions[get_column_letter(c)].hidden = True
    dados = g.get("dados") if g.get("v") == HIS_V else None
    if dados:
        for i, lin in enumerate(dados[:HIS_MAX]):
            for j, v in enumerate(lin):
                if j == 0 and isinstance(v, str):
                    try:
                        v = datetime.strptime(v, "%Y-%m-%d")
                    except ValueError:
                        pass
                ws.cell(HIS_L0 + i, HIS_COL0 + j, v)
    for r in range(HIS_L0, HIS_L0 + HIS_MAX):
        ws.cell(r, HIS_COL0).number_format = "dd/mm/yyyy"
    for k in range(HIS_N):
        lin_g = 19 + 18 * (k // 2)
        col_g = "B" if k % 2 == 0 else "H"
        ws[f"{col_g}{lin_g}"] = f"=IF(B{7 + k}=\"\",\"\",B{7 + k}&\" · \"&C{7 + k})"
        ws[f"{col_g}{lin_g}"].font = Font(bold=True, size=12, color="1F4E79")
        ch = ScatterChart()                             # XY: o eixo das datas ajusta-se aos pontos que existem
        ch.style = 13
        xs = Reference(ws, min_col=HIS_COL0, min_row=HIS_L0, max_row=HIS_L0 + HIS_MAX - 1)
        for m in range(NS):
            col = HIS_COL0 + 1 + NS * k + m
            sr = Series(Reference(ws, min_col=col, min_row=HIS_L0 - 1, max_row=HIS_L0 + HIS_MAX - 1), xs,
                        title_from_data=True)
            ch.series.append(sr)
        ch.x_axis.number_format = "dd-mm-yyyy"
        ch.x_axis.tickLblPos = "low"                    # datas em baixo, dentro da área do gráfico
        try:                                            # letra pequena (7) e inclinada
            from openpyxl.chart.text import RichText
            from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties
            ch.x_axis.txPr = RichText(bodyPr=RichTextProperties(rot=-2700000, vert="horz"),
                                      p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=700)),
                                                   endParaRPr=CharacterProperties(sz=700))])
            ch.y_axis.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=800)),
                                                   endParaRPr=CharacterProperties(sz=800))])
        except Exception:
            pass
        ch.y_axis.number_format = "#,##0.00"
        try:
            ch.x_axis.delete = False
            ch.y_axis.delete = False
        except Exception:
            pass
        from openpyxl.chart.marker import Marker
        from openpyxl.chart.shapes import GraphicalProperties
        from openpyxl.drawing.line import LineProperties
        for s_, (_, _, cor) in zip(ch.series, HIS_SER):
            s_.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=cor, w=22000))
            s_.marker = Marker(symbol="circle", size=4)
            s_.marker.graphicalProperties = GraphicalProperties(solidFill=cor, ln=LineProperties(solidFill=cor))
            s_.smooth = False
        ch.legend.position = "b"
        ch.width, ch.height = 15.5, 8
        ch.visible_cells_only = False
        ch.display_blanks = "gap"
        ws.add_chart(ch, f"{col_g}{lin_g + 1}")


def calcular_historico():
    """▶ GERAR HISTÓRICO (folha Histórico)."""
    import pandas as pd
    from datetime import date as date_
    ex = Excel(FICHEIRO)
    r = ex.ops([{"op": "ler", "folha": FOLHA_HIS, "rng": "C4"}, {"op": "ler", "folha": FOLHA_HIS, "rng": f"B7:B{6 + HIS_N}"},
                {"op": "ler", "folha": FOLHA_HIS, "rng": "C5"}])
    passo = int(round(_num_ou(r[2][0][0], 5) or 5)) if len(r) > 2 else 5
    passo = max(3, min(21, passo))                 # nunca menos de 3 dias de bolsa
    anos = _num_ou(r[0][0][0], 1) or 1
    if anos > 20:                                  # folha da V121 (dias de bolsa)
        anos = anos / 252
    anos = max(7 / 365.25, min(HIS_ANOS_MAX, anos))     # de 1 semana a 5 anos (p. ex. =4/52 para 4 semanas)
    tks = list(dict.fromkeys(txt(l_[0]).upper() for l_ in r[1] if txt(l_[0])))   # sem repetidos
    if not tks:
        print("Histórico: escreve pelo menos um ticker na coluna B (linhas 7 a 16).")
        sys.exit(1)
    carregar_pesos(ex)
    ultima = get_column_letter(len(COLS))
    rr = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)])
    todas = {txt(d.get("ticker")).upper(): d for i, tab in zip((1, 2), rr) for d in linhas_por_cabecalho(tab, i)}
    por_cod = {}                                   # a mesma empresa noutra bolsa (ex.: MU nos EUA e MTE.DE na tabela)
    for d in todas.values():
        for c_ in (txt(d.get("codigo")).upper(), txt(d.get("ticker")).upper().split(".")[0]):
            if c_:
                por_cod.setdefault(c_, d)
    linhas = []
    for t in tks:
        if t in todas:
            linhas.append(todas[t])
        elif t.split(".")[0] in por_cod:           # da tabela: nome e decisão; cotações e moeda as deste ticker
            d0 = por_cod[t.split(".")[0]]
            todas[t] = d0
            linhas.append({**d0, "ticker": t, "moeda": None})
            print(f"   {t}: na tabela como {txt(d0.get('ticker'))} (decisão de hoje dessa linha)")
        else:
            linhas.append({"ticker": t, "tipo": "Ação"})
            print(f"   {t}: não está nas folhas Ações / ETFs: o último ponto é o calculado (sem analistas)")
    moedas = {t: txt(d.get("moeda")) or "EUR" for t, d in zip(tks, linhas)}
    tipos = {t: txt(d.get("tipo")) or "Ação" for t, d in zip(tks, linhas)}
    hoje = date_.today()
    ini = hoje.replace(year=hoje.year - 12)
    # o sinal dos ETFs compara a subida de 6 meses com a dos OUTROS ETFs da tabela: também são precisos
    extra = ([t for t, d in todas.items() if txt(d.get("tipo")) == "ETF" and t not in tks]
             if any(tipos[t] == "ETF" for t in tks) else [])
    for t in extra:
        moedas[t], tipos[t] = txt(todas[t].get("moeda")) or "EUR", "ETF"
    print(f"Histórico: cotações de {len(tks)} títulos (12 anos, para as retas)"
          + (f" e de mais {len(extra)} ETFs da tabela (para comparar a subida de 6 meses)" if extra else "") + " ...")
    _sim_cotacoes_eur(tks + extra, moedas, ini, hoje)
    loc = _SIM_LOCAL.get("tab")
    cols = [t for t in tks if loc is not None and t in loc.columns]
    falta = [t for t in tks if t not in cols]
    if not cols:
        print("Histórico: não consegui as cotações destes tickers.")
        sys.exit(1)
    loc_todos = loc[cols + [t for t in extra if t in loc.columns]]
    loc = loc[cols]
    idx_ = loc.dropna(how="all").index
    idx_ = idx_[idx_ >= idx_[-1] - pd.Timedelta(days=int(anos * 365.25))]
    dias = list(idx_[::-1][::passo][::-1])         # um ponto a cada «passo» dias de bolsa, a acabar no último dia
    if len(dias) > HIS_MAX:                         # 5 anos a cada 3 dias ≈ 420 pontos: cabe; mais do que isso, alarga
        passo = -(-len(idx_) // HIS_MAX)
        dias = list(idx_[::-1][::passo][::-1])
    sec = {}
    try:
        sec = sec_preparar([d for d in linhas if txt(d.get("tipo")) != "ETF" and txt(d.get("ticker")).upper() in cols],
                           dias[0], hoje)
    except Exception as e:
        print(f"   SEC: {e}")
    print(f"   a calcular a decisão do painel em {len(dias)} pontos (um a cada {passo} dias de bolsa) ({len(sec)} com relatórios da SEC / ESEF) ...")
    dec = {}
    _sim_notas_painel(loc_todos, tipos, dias, sec, None, dec)
    for t, d in zip(tks, linhas):                  # último ponto = a decisão da tabela (tem analistas e dados de hoje)
        dt_ = txt(d.get("decisao"))
        if t not in cols or not dias:
            continue
        calc = txt((dec.get((dias[-1], t)) or [""])[0])
        if dt_:
            dec[(dias[-1], t)] = (dt_, d.get("nota_geral"), d.get("nivel_preco"))
            print(f"   {t}: último ponto = decisão da tabela «{dt_[:60]}» (calculada sem analistas: «{calc[:60]}»)")
        else:
            print(f"   {t}: sem decisão na tabela; último ponto calculado «{calc[:60]}»")
    dados, resumo = [], {}
    NS = len(HIS_SER)
    for dia in dias:
        dados.append([dia.to_pydatetime()] + [None] * (NS * HIS_N))
    ordem = [c for c, _, _ in HIS_CLASSES]
    serie_de = lambda c: c                         # uma série (cor) por classe
    ords = [c for c, _, _ in HIS_SER]
    for k, t in enumerate(tks):
        if t not in cols:
            resumo[t] = [txt(todas.get(t, {}).get("nome")), "sem cotações"]
            continue
        cont = {c: 0 for c in ordem}
        ant = None
        for i, dia in enumerate(dias):
            pr = loc.at[dia, t]
            dcs = dec.get((dia, t))
            if dcs is None or pd.isna(pr):
                ant = None
                continue
            c = _his_classe(dcs[0])
            cont[c] += 1
            sc = serie_de(c)
            dados[i][1 + NS * k + ords.index(sc)] = float(pr)
            if ant and ant != sc:                    # liga a linha: a semana da mudança também na série anterior
                dados[i][1 + NS * k + ords.index(ant)] = float(pr)
            ant = sc
        ult = next((dec[(d_, t)][0] for d_ in reversed(dias) if (d_, t) in dec), "")
        resumo[t] = [txt(todas.get(t, {}).get("nome")) or t] + [cont[c] for c in ordem] + [ult]
    _gravar_json(FICH_HIS, {"v": HIS_V, "tickers": tks, "anos": anos, "passo": passo, "resumo": resumo,
                            "dados": [[d[0].strftime("%Y-%m-%d")] + d[1:] for d in dados]})
    vazio = [[None] * (1 + NS * HIS_N) for _ in range(HIS_MAX)]
    for i, lin in enumerate(dados):
        vazio[i] = lin
    tab_res = [(resumo.get(t, []) + [None] * 8)[:8] for t in tks] + [[None] * 8] * (HIS_N - len(tks))
    ops = [{"op": "escrever", "folha": FOLHA_HIS, "rng": f"C7:J{6 + HIS_N}", "valores": tab_res},
           {"op": "escrever", "folha": FOLHA_HIS,
            "rng": f"{get_column_letter(HIS_COL0)}{HIS_L0}:{get_column_letter(HIS_COL0 + NS * HIS_N)}{HIS_L0 + HIS_MAX - 1}",
            "valores": vazio}]
    ex.ops(ops, guardar=True)
    for t in tks:
        r_ = resumo.get(t, [])
        if len(r_) >= 8:
            print(f"   {t:10} esperar {r_[1]:3d} · forte {r_[2]:3d} · média {r_[3]:3d} · compra {r_[4]:3d} · "
                  f"manter {r_[5]:3d} · venda {r_[6]:3d}")
    if falta:
        print("   sem cotações:", ", ".join(falta))


FOLHA_SIM = "Simulador"
FICH_SIM = os.path.join(PASTA, "_Painel_acoes_simulador.json")   # escolhas e último resultado do simulador
SIM_PERIODOS = ["Nenhum", "Semanal", "Mensal", "Trimestral", "Semestral", "Anual"]
SIM_COMPRAR = ["1 · Os meus títulos (com x)",
               "2 · Os N melhores pela nota geral de hoje",
               "3 · Nota geral de hoje ≥ valor",
               "4 · A carteira da tabela (tickers e pesos)",
               "5 · Os N melhores pela nota do painel em cada data",
               "6 · Nota do painel em cada data ≥ valor",
               "7 · N ações + N ETFs, nota geral de hoje",
               "8 · N ações + N ETFs, nota do painel em cada data",
               "9 · Os N com melhor preço (nível de preço) e nota geral ≥ mínimo (hoje)",
               "10 · Os N com melhor preço em cada data e nota ≥ mínimo (só ações dos EUA, dados da SEC)",
               "11 · Todos com nível de preço ≥ F6 e nota geral ≥ mínimo (hoje)",
               "12 · Todos com nível de preço ≥ F6 e nota ≥ mínimo em cada data (só ações dos EUA, dados da SEC)"]


def _sim_cod(t):
    """Número (1 a 8) da escolha «Comprar» (aceita também os textos das versões anteriores)."""
    t = txt(t).strip()
    m_ = re.match(r"(\d+)", t)
    if m_:
        return int(m_.group(1))
    lt = t.lower()
    if lt.startswith("n primeiras ações"):
        return 8 if "cada data" in lt else 7
    if lt.startswith("os n com melhor preço"):
        return 9
    for k, ini in ((1, "os meus"), (2, "os n primeiros"), (3, "nota geral"), (4, "carteira"),
                   (5, "n primeiros pela nota de preços"), (6, "nota de preços")):
        if lt.startswith(ini):
            return k
    return 2
SIM_RODAR = ["Só compras", "Rodar a carteira"]
SIM_ENTRE = ["Ações e ETFs", "Só ações", "Só ETFs"]
SIM_PADRAO = {"ini_sim": "Sim", "ini_val": 10000, "per": "Mensal", "per_val": 200, "comprar": SIM_COMPRAR[1],
              "infl": "Portugal",
              "n": 10, "entre": SIM_ENTRE[0], "d_ini": "01/01/2016", "d_fim": "", "bench": "IWDA.AS",
              "carteira": [], "rodar": SIM_RODAR[0], "n_a": 5, "v_a": 150, "n_e": 3, "v_e": 100,
              "nota_min": 4}
SIM_LIN_CART, SIM_MAX_CART = 16, 30          # carteira escolhida: B16:C45
SIM_LIN_TIT, SIM_MAX_TIT = 16, 60            # resultado por título: N16:R75
SIM_R0 = 15                                  # resumo do resultado: E15:F25
SIM_LIN_SER, SIM_MAX_SER = 14, 150           # série para o gráfico (colunas U:X; sempre 150 pontos)
SIM = {}


def _sim_num(v):
    """Número (aceita texto com vírgula); senão None."""
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v) if v == v else None
    try:
        return float(str(v).replace("€", "").replace("%", "").replace(" ", "").replace(",", "."))
    except Exception:
        return None


def _sim_guardados():
    try:
        with open(FICH_SIM, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _sim_gravar(**kw):
    g = _sim_guardados()
    g.update(kw)
    try:
        with open(FICH_SIM, "w", encoding="utf-8") as f:
            json.dump(g, f, ensure_ascii=False, default=str)
    except Exception:
        pass


def _sim_data(v):
    """Data escrita de qualquer forma (data do Excel, dd/mm/aaaa, aaaa-mm-dd) -> datetime (ou None)."""
    from datetime import datetime as dt_, timedelta
    if v in (None, ""):
        return None
    if isinstance(v, dt_):
        return dt_(v.year, v.month, v.day)
    if hasattr(v, "year") and hasattr(v, "month"):
        return dt_(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and 20000 < v < 80000:          # número de série do Excel
        return dt_(1899, 12, 30) + timedelta(days=int(v))
    s = str(v).strip()
    for f in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y", "%Y/%m/%d", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt_.strptime(s[:19] if "%H" in f else s.split(" ")[0], f)
        except Exception:
            pass
    return None


def ler_simulador(ex):
    """Escolhas da folha Simulador (como estão agora no Excel)."""
    r = ex.ops([{"op": "ler", "folha": FOLHA_SIM, "rng": "C4:F13"},
                {"op": "ler", "folha": FOLHA_SIM, "rng": f"B{SIM_LIN_CART}:C{SIM_LIN_CART + SIM_MAX_CART - 1}"}])
    v, cart = r[0], r[1]
    g = lambda i, j: sem_esp(v[i][j]) if i < len(v) and j < len(v[i]) else None
    p = {"ini_sim": txt(g(0, 0)) or "Não", "ini_val": _sim_num(g(0, 3)) or 0, "per": txt(g(1, 0)) or "Nenhum",
         "per_val": _sim_num(g(1, 3)) or 0, "comprar": txt(g(2, 0)) or SIM_COMPRAR[0], "n": _sim_num(g(2, 3)),
         "entre": txt(g(3, 0)) or SIM_ENTRE[0], "d_ini": g(4, 0), "d_fim": g(4, 3), "bench": txt(g(5, 0)),
         "rodar": txt(g(6, 0)) or SIM_RODAR[0], "n_a": _sim_num(g(7, 0)), "v_a": _sim_num(g(7, 3)),
         "n_e": _sim_num(g(8, 0)), "v_e": _sim_num(g(8, 3)), "nota_min": _sim_num(g(9, 0)),
         "carteira": [[txt(sem_esp(a)).upper(), _sim_num(sem_esp(b))] for a, b, *_ in (cart or [])
                      if re.fullmatch(r"[A-Za-z0-9.\-^=]{1,15}", txt(sem_esp(a)) or "")]}
    for k in ("d_ini", "d_fim"):
        d = _sim_data(p[k])
        p[k] = d.strftime("%d/%m/%Y") if d else ""
    p["comprar"] = SIM_COMPRAR[_sim_cod(p["comprar"]) - 1]
    return p


def carregar_simulador(ex=None):
    """Antes de reconstruir o livro: guarda as escolhas do Simulador (para a folha nova as ter)."""
    if ex is None:
        return
    try:
        p = ler_simulador(ex)
        SIM.update(p)
        _sim_gravar(param=p)
    except Exception:
        pass


def folha_simulador(wb):
    """Folha Simulador: investimento inicial e reforços periódicos nos títulos escolhidos, entre duas datas
    (cotações com dividendos, em euros). Calculado à parte, no botão ▶ SIMULAR."""
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.chart import LineChart, Reference
    g = _sim_guardados()
    p = {**SIM_PADRAO, **{k: v for k, v in {**(g.get("param") or {}), **SIM}.items() if v not in (None, "")}}
    p["comprar"] = SIM_COMPRAR[_sim_cod(p["comprar"]) - 1]
    ws = wb.create_sheet(FOLHA_SIM)
    ws["A1"] = "Simulador: investimento inicial + reforços periódicos"
    ws["A1"].font = Font(size=14, bold=True)
    ws["B2"] = ("Escolhe a amarelo e carrega em ▶ SIMULAR (calcula à parte: não mexe nas outras folhas nem na app). "
                "Compra nas datas de cada reforço, ao fecho, com as cotações ajustadas a dividendos (reinvestidos) e em "
                "euros ao câmbio de cada dia; sem comissões nem impostos. O dinheiro de cada compra divide-se em partes "
                "iguais pelos títulos (na carteira escolhida, pelos pesos); um título ainda sem cotação nessa data fica "
                "de fora dessa compra. Atenção: escolher pela nota de HOJE e simular o passado favorece o resultado "
                "(os títulos com boa nota hoje são muitas vezes os que subiram).")
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B2"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("B2:L2")
    ws.row_dimensions[2].height = 52
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    fino = Side(style="thin", color="BFBFBF")
    borda = Border(left=fino, right=fino, top=fino, bottom=fino)

    def lista(cel, itens):
        dv = DataValidation(type="list", formula1='"' + ",".join(itens) + '"', allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(cel)

    def entrada(cel, v, fmt=None):
        c = ws[cel]
        c.value = v
        c.fill, c.border, c.font = amarelo, borda, Font(bold=True)
        if fmt:
            c.number_format = fmt
    linhas_e = [("Investimento inicial", "C4", p["ini_sim"], ["Sim", "Não"], "Valor (€)", "F4", p["ini_val"]),
                ("Reforço periódico", "C5", p["per"], SIM_PERIODOS, "Valor de cada reforço (€)", "F5", p["per_val"]),
                ("Comprar", "C6", p["comprar"], SIM_COMPRAR, "N (2, 5, 9, 10) · nota mín. (3, 6) · nível mín. (11, 12)", "F6", p["n"]),
                ("Entre", "C7", p["entre"], SIM_ENTRE, None, None, None),
                ("Data inicial", "C8", _sim_data(p["d_ini"]) or p["d_ini"], None, "Data final (vazio = hoje)", "F8",
                 _sim_data(p["d_fim"]) or p["d_fim"]),
                ("Comparar com (ticker Yahoo)", "C9", p["bench"], None, None, None, None),
                ("Ao reclassificar (em cada data)", "C10", p.get("rodar") or SIM_RODAR[0], SIM_RODAR, None, None, None),
                ("Ações: N primeiras", "C11", p.get("n_a"), None, "€ em ações por reforço", "F11", p.get("v_a")),
                ("ETFs: N primeiros", "C12", p.get("n_e"), None, "€ em ETFs por reforço", "F12", p.get("v_e")),
                ("Nota mínima (5, 6, 8-12)", "C13", p.get("nota_min"), None, None, None, None)]
    for rot, cel, v, itens, rot2, cel2, v2 in linhas_e:
        r = int(cel[1:])
        ws[f"B{r}"] = rot
        ws[f"B{r}"].font = Font(bold=True)
        entrada(cel, v, "dd/mm/yyyy" if r == 8 else None)
        ws.merge_cells(f"C{r}:D{r}")
        if itens:
            lista(cel, itens)
        if rot2:
            ws[f"E{r}"] = rot2
            ws[f"E{r}"].alignment = Alignment(horizontal="right")
            entrada(cel2, v2, "dd/mm/yyyy" if r == 8 else ('"€" #,##0' if r in (4, 5, 11, 12) else "General"))
    ws["G10"] = ("(só nas escolhas «em cada data»: «Só compras» compra os melhores de cada data e guarda o que já tem; "
                 "«Rodar a carteira» vende o que saiu e reparte tudo pelos escolhidos, sem comissões nem impostos)")
    ws["G10"].font = Font(italic=True, size=9, color="808080")
    ws["G9"] = "(ex.: IWDA.AS = ações mundiais em €; vazio = sem comparação)"
    from openpyxl.formatting.rule import FormulaRule
    cod_f = 'IFERROR(VALUE(LEFT($C$6,FIND(" ",$C$6&" ")-1)),0)'
    cinza = PatternFill("solid", bgColor="808080")
    _nn, _nu = '"N: quantos títulos →"', '"(F6 não se usa nesta opção)"'
    ws["E6"] = (f'=CHOOSE({cod_f}+1,{_nu},{_nu},{_nn},"Nota geral mínima (hoje) →",{_nu},{_nn},'
                f'"Nota mínima em cada data →",{_nu},{_nu},{_nn},{_nn},"Nível de preço mínimo (5 = muito barata) →",'
                f'"Nível de preço mínimo em cada data (5 = muito barata) →")')
    ws["E6"].font = Font(bold=True, color="7030A0")
    usa = {"F6": (2, 3, 5, 6, 9, 10, 11, 12), "C7:D7": (1, 2, 3, 5, 6, 9, 11), "C10:D10": (5, 6, 8, 10, 12),
           "C11:D11": (7, 8), "F11": (7, 8), "C12:D12": (7, 8), "F12": (7, 8), "C13:D13": (5, 6, 8, 9, 10, 11, 12),
           "F5": (1, 2, 3, 4, 5, 6, 9, 10, 11, 12),
           f"B{SIM_LIN_CART}:C{SIM_LIN_CART + SIM_MAX_CART - 1}": (4,)}
    for rng_, cods in usa.items():                       # não conta para a escolha feita em «Comprar» -> cinzento
        f_ = "NOT(OR(" + ",".join(f"{cod_f}={c}" for c in cods) + "))"
        ws.conditional_formatting.add(rng_, FormulaRule(formula=[f_], fill=cinza, font=Font(color="BFBFBF")))
    ws["E13"] = ("(opções 5, 6, 8, 10, 12: não se compra nenhum título com nota abaixo disto na data da compra — só com "
                 "dados antigos a nota das ações vai no máximo a ≈4,2; opções 9 e 11: só entram os de nota geral de hoje "
                 "≥ isto; vazio = sem mínimo)")
    ws["E13"].font = Font(italic=True, size=9, color="808080")
    ws["G11"] = ("(só em «N primeiras ações + N primeiros ETFs»: cada reforço põe estes valores em cada parte, na "
                 "periodicidade escolhida; o investimento inicial reparte-se na mesma proporção)")
    ws["G11"].font = Font(italic=True, size=9, color="808080")
    ws["G9"].font = Font(italic=True, size=9, color="808080")
    botao(ws, "H4", "▶ SIMULAR", "painelacoes://simular", "7030A0", "4B0082", 14)
    ws.merge_cells("H4:J5")
    for col, w in zip("ABCDEFGHIJKLMNOPQRST", (2, 30, 14, 14, 50, 16, 3, 12, 12, 12, 3, 3, 3, 34, 11, 13, 13, 13, 13, 11)):
        ws.column_dimensions[col].width = w
    # carteira escolhida
    ws[f"B{SIM_LIN_CART - 2}"] = "Carteira escolhida (só com «Carteira escolhida»)"
    ws[f"B{SIM_LIN_CART - 2}"].font = Font(bold=True, size=12)
    for j, h in ((2, "Ticker Yahoo (ex.: MSFT, ASML.AS, IWDA.AS)"), (3, "Peso (%)")):
        c = ws.cell(SIM_LIN_CART - 1, j, h)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3B57")
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    cart = [c for c in (p.get("carteira") or []) if re.fullmatch(r"[A-Za-z0-9.\-^=]{1,15}", txt(c[0]) or "")]
    for i in range(SIM_MAX_CART):
        r = SIM_LIN_CART + i
        t, w = (cart[i] + [None, None])[:2] if i < len(cart) else (None, None)
        entrada(f"B{r}", t)
        entrada(f"C{r}", w, "0.0")
    ws[f"B{SIM_LIN_CART + SIM_MAX_CART}"] = "Pesos vazios = partes iguais."
    ws[f"B{SIM_LIN_CART + SIM_MAX_CART}"].font = Font(italic=True, size=9, color="808080")
    # resultados (preenchidos pelo ▶ SIMULAR)
    ws[f"E{SIM_R0 - 1}"] = "Resultado"
    ws[f"E{SIM_R0 - 1}"].font = Font(bold=True, size=12)
    for j, h in enumerate(("Título", "Ticker", "Comprado (€)", "Vendido (€)", "Valor final (€)", "Resultado (€)",
                           "Resultado (%)"), 14):
        c = ws.cell(SIM_LIN_TIT - 1, j, h)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F3B57")
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.cell(SIM_LIN_TIT - 2, 14, "Por título").font = Font(bold=True, size=12)
    for i in range(SIM_MAX_TIT):
        r = SIM_LIN_TIT + i
        for j in (16, 17, 18):
            ws.cell(r, j).number_format = '"€" #,##0.00'
        ws.cell(r, 19).number_format = '+"€" #,##0.00;-"€" #,##0.00'
        ws.cell(r, 20).number_format = "+0.0%;-0.0%;0.0%"
    for i in range(SIM_MAX_SER):
        r = SIM_LIN_SER + i
        ws.cell(r, 21).number_format = "dd/mm/yyyy"
        for j in (22, 23, 24, 25):
            ws.cell(r, j).number_format = '"€" #,##0'
    for j, h in enumerate(("Data", "Investido", "Carteira", "Comparação", "Carteira descontada da inflação"), 21):
        ws.cell(SIM_LIN_SER - 1, j, h).font = Font(bold=True)
    for col in "UVWXY":
        ws.column_dimensions[col].hidden = True
    ch = LineChart()
    ch.title = "Valor da carteira vs dinheiro investido (€)"
    ch.add_data(Reference(ws, min_col=22, max_col=25, min_row=SIM_LIN_SER - 1, max_row=SIM_LIN_SER + SIM_MAX_SER - 1),
                titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=21, min_row=SIM_LIN_SER, max_row=SIM_LIN_SER + SIM_MAX_SER - 1))
    ch.y_axis.number_format = '#,##0'
    ch.x_axis.number_format = "mm/yyyy"
    ch.x_axis.majorTimeUnit = "years"
    ch.y_axis.title, ch.x_axis.title = "€", None
    try:
        ch.x_axis.delete = False
        ch.y_axis.delete = False
    except Exception:
        pass
    for s_, cor in zip(ch.series, ("A6A6A6", "1F77B4", "FF7F0E", "C00000")):
        s_.graphicalProperties.line.solidFill = cor
        s_.graphicalProperties.line.width = 22000
        s_.smooth = False
    ch.series[3].graphicalProperties.line.dashStyle = "dash"
    ch.width, ch.height = 24, 11
    ch.visible_cells_only = False                   # os dados estão em colunas escondidas
    ws.add_chart(ch, f"B{SIM_LIN_CART + SIM_MAX_CART + 2}")
    res = g.get("res")
    if res and res.get("v") == 3:                                         # o último resultado continua à vista depois de um ▶ ATUALIZAR
        for cel, v in res.get("resumo", []):
            ws[cel] = v
        for i, lin in enumerate(res.get("titulos", [])[:SIM_MAX_TIT]):
            for j, v in enumerate(lin, 14):
                ws.cell(SIM_LIN_TIT + i, j, v)
        for i, lin in enumerate(res.get("serie", [])[:SIM_MAX_SER]):
            for j, v in enumerate(lin, 21):
                ws.cell(SIM_LIN_SER + i, j, _sim_data(v) if j == 21 else v)
    _sim_formatar_resumo(ws)


SIM_RESUMO = ["Total investido", "Valor final", "Ganho (€)", "Ganho (% do investido)",
              "Rendimento anual (TIR, conta as datas dos reforços)",
              "Valor final descontado da inflação (em € da data inicial)", "Ganho real (acima da inflação)",
              "Inflação acumulada (data inicial → final)", "Maior queda pelo caminho (sem contar reforços)",
              "Comparação: valor final com o mesmo dinheiro", "Comparação: rendimento anual (TIR)", "Títulos comprados",
              "Período", "Notas"]


def _sim_formatar_resumo(ws):
    fmts = ['"€" #,##0.00', '"€" #,##0.00', '+"€" #,##0.00;-"€" #,##0.00', "+0.0%;-0.0%", "+0.00%;-0.00%",
            '"€" #,##0.00', '+"€" #,##0.00;-"€" #,##0.00', "+0.0%;-0.0%", "0.0%", '"€" #,##0.00', "+0.00%;-0.00%",
            "General", "General", "General"]
    for i, (rot, f) in enumerate(zip(SIM_RESUMO, fmts), SIM_R0):
        ws[f"E{i}"] = rot
        ws[f"E{i}"].alignment = Alignment(horizontal="right")
        ws[f"F{i}"].number_format = f
        ws[f"F{i}"].font = Font(bold=True, size=12 if i < SIM_R0 + 5 else 10)
    fim_ = SIM_R0 + len(SIM_RESUMO) - 1
    ws[f"F{fim_}"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"F{fim_}:L{fim_ + 5}")


def _xirr(fluxos):
    """TIR anual de [(data, valor)] (negativo = dinheiro posto, positivo = valor final)."""
    if not fluxos or all(v >= 0 for _, v in fluxos) or all(v <= 0 for _, v in fluxos):
        return None
    t0 = fluxos[0][0]
    f = lambda r: sum(v / (1 + r) ** ((d - t0).days / 365.25) for d, v in fluxos)
    lo, hi = -0.99, 10.0
    if f(lo) * f(hi) > 0:
        return None
    for _ in range(200):
        m = (lo + hi) / 2
        if f(lo) * f(m) <= 0:
            hi = m
        else:
            lo = m
    return (lo + hi) / 2


def _sim_limpar(s):
    """Corrige erros conhecidos das cotações do Yahoo: saltos de ~100× (pence/libras trocados) e picos isolados
    (um dia 3× acima/abaixo dos vizinhos)."""
    s = s[s > 0].astype(float).copy()
    if len(s) < 5:
        return s
    v = list(s.values)
    out, fator = [v[0]], 1.0
    for k in range(1, len(v)):
        r = v[k] * fator / out[-1]
        if 50 < r < 200:                            # passou a vir 100× maior: volta à escala de antes
            fator /= 100.0
        elif 1 / 200 < r < 1 / 50:                  # passou a vir 100× menor
            fator *= 100.0
        out.append(v[k] * fator)
    s[:] = out
    med = s.rolling(7, center=True, min_periods=3).median()
    mau = (s / med > 3) | (s / med < 1 / 3)
    return s.mask(mau).ffill()


FICH_INFL = os.path.join(PASTA, "_Painel_acoes_inflacao.json")
# Portugal, IHPC, variação média anual (%) — só se o Eurostat não responder e não houver cópia guardada
INFL_PT_ANUAL = {2000: 2.8, 2001: 4.4, 2002: 3.7, 2003: 3.3, 2004: 2.5, 2005: 2.1, 2006: 3.0, 2007: 2.4, 2008: 2.7,
                 2009: -0.9, 2010: 1.4, 2011: 3.6, 2012: 2.8, 2013: 0.4, 2014: -0.2, 2015: 0.5, 2016: 0.6,
                 2017: 1.6, 2018: 1.2, 2019: 0.3, 2020: -0.1, 2021: 0.9, 2022: 8.1, 2023: 5.3, 2024: 2.7,
                 2025: 2.2}


def _sim_inflacao():
    """Índice de preços mensal de Portugal (IHPC, Eurostat prc_hicp_midx, 2015 = 100), como Series de fins de mês.
    Guarda uma cópia; sem net usa a cópia ou a tabela anual. -> (Series, fonte)"""
    import pandas as pd
    url = ("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/prc_hicp_midx"
           "?format=JSON&lang=EN&geo=PT&coicop=CP00&unit=I15")
    dados, fonte = None, None
    try:
        import urllib.request
        with urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"}),
                                    timeout=25) as r:
            js = json.loads(r.read().decode("utf-8"))
        tempo = js["dimension"]["time"]["category"]["index"]
        val = js["value"]
        dados = {t: float(val[str(i)]) for t, i in tempo.items() if str(i) in val}
        if len(dados) > 24:
            fonte = "Eurostat (IHPC Portugal, mensal)"
            try:
                with open(FICH_INFL, "w", encoding="utf-8") as f:
                    json.dump({"dados": dados, "quando": datetime.now().strftime("%d/%m/%Y")}, f)
            except Exception:
                pass
        else:
            dados = None
    except Exception:
        dados = None
    if dados is None:
        try:
            with open(FICH_INFL, encoding="utf-8") as f:
                g = json.load(f)
            dados, fonte = g["dados"], f"Eurostat (cópia de {g.get('quando', '?')})"
        except Exception:
            dados = None
    if dados:
        s = pd.Series({pd.Timestamp(int(k[:4]), int(k[5:7]), 1) + pd.offsets.MonthEnd(0): v
                       for k, v in dados.items()}).sort_index()
    else:                                           # tabela anual -> mensal (crescimento igual em cada mês)
        fonte = "tabela anual aproximada (sem ligação ao Eurostat)"
        nivel, pts = 100.0, {}
        for a in sorted(INFL_PT_ANUAL):
            m = (1 + INFL_PT_ANUAL[a] / 100) ** (1 / 12)
            for mes in range(1, 13):
                nivel *= m
                pts[pd.Timestamp(a, mes, 1) + pd.offsets.MonthEnd(0)] = nivel
        s = pd.Series(pts)
    return s, fonte


def _sim_infl_em(serie_i, datas):
    """Nível do índice em cada data (o do mês; depois do último mês conhecido, continua ao ritmo dos últimos 12)."""
    import pandas as pd
    ult = serie_i.index[-1]
    taxa_m = (serie_i.iloc[-1] / serie_i.iloc[-13]) ** (1 / 12) if len(serie_i) > 13 else 1.0
    out = []
    for d in datas:
        fm = pd.Timestamp(d) + pd.offsets.MonthEnd(0)
        if fm <= ult:
            pos = serie_i.index.searchsorted(fm)
            out.append(float(serie_i.iloc[min(pos, len(serie_i) - 1)]) if pos < len(serie_i) or True else None)
        else:
            meses = (fm.year - ult.year) * 12 + fm.month - ult.month
            out.append(float(serie_i.iloc[-1]) * taxa_m ** meses)
    return out


# ================================================================ fundamentos históricos (SEC, só empresas dos EUA)
PASTA_SEC = os.path.join(PASTA, "_Painel_acoes_sec")             # cópias dos dados da SEC (renovadas a cada 30 dias)
FICH_SEC_CONTACTO = os.path.join(PASTA, "_Painel_acoes_sec_contacto.txt")  # opcional: e-mail para o pedido à SEC
SEC_FLUXO = {
    "eps": ["EarningsPerShareDiluted", "EarningsPerShareBasic"],
    "ni": ["NetIncomeLoss", "ProfitLoss", "NetIncomeLossAvailableToCommonStockholdersBasic"],
    "ocf": ["NetCashProvidedByUsedInOperatingActivities",
            "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"],
    "capex": ["PaymentsToAcquirePropertyPlantAndEquipment", "PaymentsToAcquireProductiveAssets"],
    "bruto": ["GrossProfit"],
    "receita": ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax", "SalesRevenueNet"],
    "custo": ["CostOfRevenue", "CostOfGoodsAndServicesSold", "CostOfGoodsSold"],
    "ebit": ["OperatingIncomeLoss"],
    "juros": ["InterestExpense", "InterestExpenseNonoperating", "InterestExpenseDebt"],
    "imp": ["IncomeTaxExpenseBenefit"],
    "antes": ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest",
              "IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments"],
    "da": ["DepreciationDepletionAndAmortization", "DepreciationAmortizationAndAccretionNet",
           "DepreciationAndAmortization", "Depreciation"],
    "acoes": ["WeightedAverageNumberOfDilutedSharesOutstanding"],
}
SEC_INSTANTE = {
    "ativo": ["Assets"],
    "cp": ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
    "caixa": ["CashAndCashEquivalentsAtCarryingValue",
              "CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents"],
    "div_lp": ["LongTermDebtNoncurrent", "LongTermDebt"],
    "div_cp": ["LongTermDebtCurrent", "DebtCurrent", "ShortTermBorrowings"],
}


def _sec_ua():
    try:
        with open(FICH_SEC_CONTACTO, encoding="utf-8") as f:
            c = f.read().strip()
        if c:
            return f"PainelAcoes/1.0 ({c})"
    except Exception:
        pass
    return "PainelAcoes/1.0 (uso pessoal; painel.acoes@example.com)"


def _sec_get(url, destino, dias=30):
    """JSON da SEC, com cópia local (válida «dias» dias). Sem net: usa a cópia, mesmo velha."""
    import urllib.request
    import gzip
    os.makedirs(PASTA_SEC, exist_ok=True)
    if os.path.exists(destino) and time.time() - os.path.getmtime(destino) < dias * 86400:
        with open(destino, encoding="utf-8") as f:
            return json.load(f)
    try:
        rq = urllib.request.Request(url, headers={"User-Agent": _sec_ua(), "Accept-Encoding": "gzip"})
        with urllib.request.urlopen(rq, timeout=40) as r:
            dados = r.read()
            if r.headers.get("Content-Encoding") == "gzip":
                dados = gzip.decompress(dados)
        js = json.loads(dados.decode("utf-8"))
        with open(destino, "w", encoding="utf-8") as f:
            json.dump(js, f)
        time.sleep(0.15)                                 # a SEC pede no máximo 10 pedidos por segundo
        return js
    except Exception:
        if os.path.exists(destino):
            with open(destino, encoding="utf-8") as f:
                return json.load(f)
        raise


def _sec_ciks():
    js = _sec_get("https://www.sec.gov/files/company_tickers.json", os.path.join(PASTA_SEC, "tickers.json"))
    return {txt(v.get("ticker")).upper().replace(".", "-"): int(v["cik_str"]) for v in js.values()}


def _sec_anuais(js):
    """Relatórios anuais (10-K) de uma empresa: [{fim, disponivel, eps, ni, ...}], cada um com os valores tal como
    foram divulgados pela primeira vez (sem olhar para correções futuras)."""
    from datetime import datetime as dt_
    g = (js.get("facts") or {}).get("us-gaap") or {}
    if not g:
        return []
    porfim = {}

    def ler(chave, conceitos, fluxo):
        for pri, c in enumerate(conceitos):                # o 1.º conceito manda; os seguintes só preenchem anos
            u = (g.get(c) or {}).get("units") or {}
            vals = u.get("USD") or u.get("USD/shares") or u.get("shares") or []
            for x in vals:
                if x.get("form") not in ("10-K", "10-K/A") or x.get("fp") != "FY" or "end" not in x:
                    continue
                if fluxo:
                    if "start" not in x:
                        continue
                    dur = (dt_.strptime(x["end"], "%Y-%m-%d") - dt_.strptime(x["start"], "%Y-%m-%d")).days
                    if not 330 <= dur <= 400:
                        continue
                elif "start" in x:
                    continue
                reg = porfim.setdefault(x["end"], {})
                ant = reg.get(chave)
                if ant is None or pri < ant[2] or (pri == ant[2] and x["filed"] < ant[1]):
                    reg[chave] = (float(x["val"]), x["filed"], pri)
    for k, cs in SEC_FLUXO.items():
        ler(k, cs, True)
    for k, cs in SEC_INSTANTE.items():
        ler(k, cs, False)
    out = []
    for fim, reg in porfim.items():
        if "eps" not in reg and "ni" not in reg:
            continue
        r = {k: v[0] for k, v in reg.items()}
        base = reg.get("eps") or reg.get("ni")
        r["fim"], r["disponivel"] = fim, base[1]
        if "bruto" not in r and "receita" in r and "custo" in r:
            r["bruto"] = r["receita"] - r["custo"]
        r["divida"] = (r.get("div_lp") or 0.0) + (r.get("div_cp") or 0.0) if ("div_lp" in r or "div_cp" in r) else None
        out.append(r)
    out.sort(key=lambda r: r["disponivel"])
    return out


def _sec_ttm_eps(js):
    """Lucro por ação dos últimos 12 meses (4 trimestres) em cada data de divulgação, só com o que já tinha saído:
    [(disponivel, eps_12m, fim)], por data. Trimestres dos 10-Q; o 4.º = ano (10-K) − 3 trimestres."""
    from datetime import datetime as dt_
    g = (js.get("facts") or {}).get("us-gaap") or {}
    per = {}                                           # (início, fim) -> (valor, divulgado), 1.ª divulgação
    for pri, c in enumerate(SEC_FLUXO["eps"]):
        for x in ((g.get(c) or {}).get("units") or {}).get("USD/shares") or []:
            if x.get("form") not in ("10-Q", "10-K", "10-Q/A", "10-K/A") or "start" not in x or "end" not in x:
                continue
            k = (x["start"], x["end"])
            ant = per.get(k)
            if ant is None or pri < ant[2] or (pri == ant[2] and x["filed"] < ant[1]):
                per[k] = (float(x["val"]), x["filed"], pri)
    dd = lambda a, b: (dt_.strptime(b, "%Y-%m-%d") - dt_.strptime(a, "%Y-%m-%d")).days
    tri = {}                                           # fim -> (início, eps, divulgado)
    for (a, b), (v, f, _) in per.items():
        if 80 <= dd(a, b) <= 100:
            tri[b] = (a, v, f)
    for (a, b), (v, f, _) in per.items():              # 4.º trimestre a partir do ano
        if 330 <= dd(a, b) <= 400 and b not in tri:
            qs = [q for e, q in tri.items() if a <= q[0] and e < b]
            if len(qs) == 3:
                ini4 = max(e for e, q in tri.items() if a <= q[0] and e < b)
                tri[b] = (ini4, v - sum(q[1] for q in qs), f)
    fins = sorted(tri)
    out = []
    for i in range(3, len(fins)):
        q4 = fins[i - 3:i + 1]
        if dd(q4[0], q4[-1]) > 300:                     # 4 trimestres seguidos (≈ 9 meses entre o 1.º e o último fim)
            continue
        disp = max(tri[e][2] for e in q4)
        out.append((disp, sum(tri[e][1] for e in q4), q4[-1]))
    out.sort()
    return out


def _qual_saude(v, financeira=False):
    """Qualidade (0-5) e saúde (0-5, com ALERTA) com as mesmas regras de fundamentos(), a partir de números."""
    out = {}
    ni, fcf, ebit, juros = v.get("ni"), v.get("fcf"), v.get("ebit"), v.get("juros")
    bruto, ativo, cp = v.get("bruto"), v.get("ativo"), v.get("cp")
    divida, caixa, imp, antes, da = v.get("divida"), v.get("caixa"), v.get("imp"), v.get("antes"), v.get("da")
    ebitda = (ebit + da) if (ebit is not None and da is not None) else None
    capital = (cp + (divida or 0) - (caixa or 0)) if cp is not None else None
    liquida = (divida or 0) - (caixa or 0) if (divida is not None or caixa is not None) else None
    taxa = min(max(imp / antes, 0.0), 0.35) if (imp is not None and antes and antes > 0) else 0.21
    if ebit is not None and capital and capital > 0:
        out["roic"] = ebit * (1 - taxa) / capital
    if bruto is not None and ativo:
        out["gpa"] = bruto / ativo
    if fcf is not None and ni and ni > 0:
        out["fcf_lucro"] = fcf / ni
    q = [x for x in (_grad(out.get("roic"), *Q_ROIC), _grad(out.get("gpa"), *Q_GPA),
                     _grad(out.get("fcf_lucro"), *Q_FCFL)) if x is not None]
    if len(q) >= 2:
        out["qual"] = round(5 * sum(q) / len(q), 1)
    if financeira:
        out["saude"] = "não se aplica (setor financeiro: banco, seguro ou financeira)"
        return out
    mau_d, bom_d = S_DIV
    pts, alerta = [], []
    dv = liquida / ebitda if (liquida is not None and ebitda and ebitda > 0) else None
    cob = ebit / abs(juros) if (ebit is not None and juros and abs(juros) > 0) else None
    if liquida is not None and liquida <= 0:
        pts.append(1.0)
    elif dv is not None:
        pts.append(_grad(dv, mau_d, bom_d))
        if dv > mau_d:
            alerta.append(f"dívida líq. {dv:.1f}× EBITDA")
    elif ebitda is not None and ebitda <= 0 and liquida and liquida > 0:
        pts.append(0.0)
        alerta.append("EBITDA negativo com dívida")
    if cob is not None:
        pts.append(_grad(cob, *S_JUROS))
        if cob < S_JUROS[0]:
            alerta.append(f"juros cobertos {cob:.1f}×")
    elif ebit is not None and juros is not None and abs(juros) == 0:
        pts.append(1.0)
    if fcf is not None:
        caixa_liq = liquida is not None and liquida <= 0
        pts.append(1.0 if fcf >= 0 else 0.5 if caixa_liq else 0.0)
        if fcf < 0 and not caixa_liq:
            alerta.append("FCF negativo com dívida")
    if pts:
        out["saude_n"] = round(5 * sum(pts) / len(pts), 1)
    out["saude"] = ("ALERTA: " + "; ".join(alerta)) if alerta else "OK"
    return out


class SecHist:
    """Fundamentos de uma empresa dos EUA em qualquer data, só com os relatórios anuais já divulgados nessa data
    (lucro por ação, PER, PER habitual de 3 anos, valor justo, FCF yield, qualidade e saúde)."""

    def __init__(self, anuais, splits, preco_usd, financeira=False, ttm=None):
        import pandas as pd
        self.ttm = ttm or []                          # lucro por ação dos últimos 4 trimestres (10-Q)
        self.ttm_disp = [pd.Timestamp(x[0]) for x in self.ttm]
        self.an = anuais
        self.disp = [pd.Timestamp(r["disponivel"]) for r in anuais]
        self.splits = splits if splits is not None else pd.Series(dtype=float)
        self.p = preco_usd.dropna()
        self.fin = financeira
        self._pe_mes = None

    def _fator(self, desde):
        s = self.splits[self.splits.index > desde] if len(self.splits) else []
        f = 1.0
        for v in (s.values if len(s) else []):
            if v and v > 0:
                f *= float(v)
        return f

    def registo(self, dia):
        import bisect
        import pandas as pd
        i = bisect.bisect_right(self.disp, pd.Timestamp(dia)) - 1
        if i < 0:
            return None
        r = dict(self.an[i])
        f = self._fator(pd.Timestamp(r["disponivel"]))   # ações desdobradas depois: em ações de hoje
        if r.get("eps") is None and r.get("ni") is not None and r.get("acoes"):
            r["eps"] = r["ni"] / r["acoes"]
        if r.get("eps") is not None:
            r["eps"] = r["eps"] / f
        j = bisect.bisect_right(self.ttm_disp, pd.Timestamp(dia)) - 1
        if j >= 0 and self.ttm_disp[j] >= pd.Timestamp(r["disponivel"]):   # trimestres mais recentes do que o ano
            r["eps"] = self.ttm[j][1] / self._fator(self.ttm_disp[j])
        if r.get("acoes") is not None:
            r["acoes"] = r["acoes"] * f
        if r.get("ocf") is not None and r.get("capex") is not None:
            r["fcf"] = r["ocf"] - abs(r["capex"])
        return r

    def _pe_mensal(self):
        import pandas as pd
        if self._pe_mes is None:
            pm = self.p.resample("ME").last().dropna()
            vals = {}
            for d_, pr in pm.items():
                r = self.registo(d_)
                if r and r.get("eps") and r["eps"] > 0:
                    vals[d_] = pr / r["eps"]
            self._pe_mes = pd.Series(vals, dtype=float)
        return self._pe_mes

    def em(self, dia):
        """Campos do painel nesta data (só com o que se sabia): justo_pct, justo, pe, pe_hab, fcf_yield, qual,
        saude_n, saude."""
        import pandas as pd
        r = self.registo(dia)
        if not r:
            return {}
        pr = self.p[self.p.index <= pd.Timestamp(dia)]
        if not len(pr):
            return {}
        preco = float(pr.iloc[-1])
        out = _qual_saude(r, self.fin)
        eps = r.get("eps")
        if eps and eps > 0:
            out["pe"] = preco / eps
            pem = self._pe_mensal()
            pem = pem[(pem.index <= pd.Timestamp(dia)) & (pem.index > pd.Timestamp(dia) - pd.Timedelta(days=3 * 366))]
            if len(pem) >= 12:
                ph = float(pem.median())
                out["pe_hab"] = ph
                pm = min(ph, PESOS["pe_hab_max"]) if PESOS.get("cap_pe_sinal") else ph
                out["justo"] = eps * pm
                out["justo_pct"] = out["justo"] / preco - 1
        if r.get("fcf") is not None and r.get("acoes"):
            out["fcf_yield"] = r["fcf"] / (preco * r["acoes"])
        out["_preco_usd"] = preco
        return out


def sec_preparar(linhas_acoes, ini, fim):
    """{ticker Yahoo: SecHist}: EUA pela SEC (10-K e 10-Q); V157: Europa pelos relatórios anuais ESEF."""
    out = _sec_preparar_eua(linhas_acoes, ini, fim)
    resto = [d for d in linhas_acoes if txt(d.get("ticker")).upper() not in out and "." in txt(d.get("ticker"))]
    if resto:
        try:
            out.update(esef_preparar(resto, ini, fim))
        except Exception as e:
            print(f"   ESEF: {e}")
    return out


def _sec_preparar_eua(linhas_acoes, ini, fim):
    """{ticker Yahoo: SecHist} para as ações da lista cuja empresa está na SEC (relatórios 10-K, EUA)."""
    import pandas as pd
    from datetime import timedelta
    try:
        ciks = _sec_ciks()
    except Exception as e:
        print(f"   SEC indisponível ({e}): sem fundamentos históricos.")
        return {}
    pares = []
    for d in linhas_acoes:
        t = txt(d.get("ticker")).upper()
        cod = txt(d.get("codigo")).upper().replace(".", "-").replace(" ", "")
        for c in (cod, t if "." not in t else ""):
            if c and c in ciks:
                pares.append((t, c, ciks[c], txt(d.get("saude")).startswith("não se aplica")))
                break
    if not pares:
        return {}
    print(f"   SEC: relatórios anuais de {len(pares)} empresas dos EUA ...")
    yf = _yf()
    us = sorted({c for _, c, _, _ in pares})
    try:
        h = _por_ticker(yf.download(us, start=(ini - timedelta(days=4 * 366)).strftime("%Y-%m-%d"),
                                    end=(fim + timedelta(days=1)).strftime("%Y-%m-%d"), interval="1d",
                                    group_by="ticker", auto_adjust=False, progress=False, threads=True), us)
    except Exception:
        h = {}
    out = {}
    for t, c, cik, fin in pares:
        try:
            js = _sec_get(f"https://data.sec.gov/api/xbrl/companyfacts/CIK{cik:010d}.json",
                          os.path.join(PASTA_SEC, f"CIK{cik:010d}.json"))
            an = _sec_anuais(js)
            df = h.get(c) if hasattr(h, "get") else None
            if not an or df is None or "Close" not in df:
                continue
            p = df["Close"].dropna()
            if getattr(p.index, "tz", None) is not None:
                p.index = p.index.tz_localize(None)
            p.index = p.index.normalize()
            try:
                sp = yf.Ticker(c).splits
                if getattr(sp.index, "tz", None) is not None:
                    sp.index = sp.index.tz_localize(None)
            except Exception:
                sp = None
            try:
                ttm = _sec_ttm_eps(js)
            except Exception:
                ttm = []
            out[t] = SecHist(an, sp, p, fin, ttm)
        except Exception:
            continue
    print(f"   SEC: fundamentos históricos de {len(out)} empresas.")
    return out


# ======================================================= fundamentos históricos (ESEF: relatórios anuais europeus, V157)
# Desde 2020/2021 as empresas cotadas na UE/EEE/Reino Unido publicam o relatório anual em formato digital (ESEF).
# O repositório público filings.xbrl.org guarda-os (grátis). Em cada data conta só o último relatório anual já
# publicado (data de publicação; sem ela, fim do ano + 4 meses, o prazo legal). Só relatórios anuais: o lucro por ação
# é o do último ano (não há trimestres como na SEC). Analistas continuam a não existir para o passado.
PASTA_ESEF = os.path.join(PASTA, "_Painel_acoes_esef")
FICH_ESEF_LEI = os.path.join(PASTA, "_Painel_acoes_esef_lei.json")    # ticker -> LEI (automático; podes corrigir)
ESEF_URL = "https://filings.xbrl.org"
ESEF_DIAS_PUB = 120                # sem data de publicação: fim do ano + 4 meses
ESEF_FLUXO = {
    "eps": ["ifrs-full:DilutedEarningsLossPerShare", "ifrs-full:BasicEarningsLossPerShare",
            "ifrs-full:DilutedEarningsLossPerShareFromContinuingOperations",
            "ifrs-full:BasicEarningsLossPerShareFromContinuingOperations"],
    "ni": ["ifrs-full:ProfitLossAttributableToOwnersOfParent", "ifrs-full:ProfitLoss"],
    "ocf": ["ifrs-full:CashFlowsFromUsedInOperatingActivities",
            "ifrs-full:CashFlowsFromUsedInOperatingActivitiesContinuingOperations"],
    "capex": ["ifrs-full:PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities",
              "ifrs-full:PurchaseOfPropertyPlantAndEquipment",
              "ifrs-full:PurchaseOfPropertyPlantAndEquipmentIntangibleAssetsOtherThanGoodwillInvestmentPropertyAndOtherNoncurrentAssets"],
    "bruto": ["ifrs-full:GrossProfit"],
    "receita": ["ifrs-full:Revenue", "ifrs-full:RevenueFromContractsWithCustomers"],
    "custo": ["ifrs-full:CostOfSales"],
    "ebit": ["ifrs-full:ProfitLossFromOperatingActivities"],
    "juros": ["ifrs-full:InterestExpense", "ifrs-full:FinanceCosts", "ifrs-full:InterestExpenseOnBorrowings"],
    "imp": ["ifrs-full:IncomeTaxExpenseContinuingOperations"],
    "antes": ["ifrs-full:ProfitLossBeforeTax"],
    "da": ["ifrs-full:DepreciationAndAmortisationExpense",
           "ifrs-full:AdjustmentsForDepreciationAndAmortisationExpense",
           "ifrs-full:DepreciationAmortisationAndImpairmentLossReversalOfImpairmentLossRecognisedInProfitOrLoss",
           "ifrs-full:AdjustmentsForDepreciationAmortisationAndImpairmentLossReversalOfImpairmentLossRecognisedInProfitOrLoss",
           "ifrs-full:DepreciationExpense"],
    "acoes": ["ifrs-full:AdjustedWeightedAverageShares", "ifrs-full:WeightedAverageShares"],
}
ESEF_INSTANTE = {
    "ativo": ["ifrs-full:Assets"],
    "cp": ["ifrs-full:EquityAttributableToOwnersOfParent", "ifrs-full:Equity"],
    "caixa": ["ifrs-full:CashAndCashEquivalents"],
    "div_lp": ["ifrs-full:NoncurrentPortionOfNoncurrentBorrowings", "ifrs-full:LongtermBorrowings",
               "ifrs-full:NoncurrentBorrowings"],
    "div_cp": ["ifrs-full:CurrentBorrowingsAndCurrentPortionOfNoncurrentBorrowings",
               "ifrs-full:CurrentPortionOfLongtermBorrowings", "ifrs-full:ShorttermBorrowings",
               "ifrs-full:CurrentBorrowings"],
    "div_tot": ["ifrs-full:Borrowings"],
}
_ESEF_DIM_OK = {"concept", "entity", "period", "unit", "language", "noteId"}


def _esef_get(url, destino=None, dias=7, binario=False):
    """GET com cópia local (válida «dias» dias; sem net usa a cópia velha). Devolve o JSON (ou bytes)."""
    import urllib.request
    import gzip
    if destino and os.path.exists(destino) and time.time() - os.path.getmtime(destino) < dias * 86400:
        with open(destino, "rb") as f:
            dados = f.read()
        return dados if binario else json.loads(dados.decode("utf-8"))
    try:
        rq = urllib.request.Request(url, headers={"User-Agent": _sec_ua(), "Accept-Encoding": "gzip",
                                                  "Accept": "application/vnd.api+json, application/json"})
        with urllib.request.urlopen(rq, timeout=90) as r:
            dados = r.read()
            if r.headers.get("Content-Encoding") == "gzip":
                dados = gzip.decompress(dados)
        if destino:
            os.makedirs(os.path.dirname(destino), exist_ok=True)
            with open(destino, "wb") as f:
                f.write(dados)
        time.sleep(0.2)
        return dados if binario else json.loads(dados.decode("utf-8"))
    except Exception:
        if destino and os.path.exists(destino):
            with open(destino, "rb") as f:
                dados = f.read()
            return dados if binario else json.loads(dados.decode("utf-8"))
        raise


def _esef_filings(lei):
    """{"nome", "pais", "rel": [...]} dos relatórios de uma empresa (pela API, filtrada pelo LEI; cópia de 7 dias)."""
    import urllib.parse
    filtro = json.dumps([{"name": "entity.identifier", "op": "eq", "val": lei}])
    url = (f"{ESEF_URL}/api/filings?include=entity&page%5Bsize%5D=200&filter="
           + urllib.parse.quote(filtro))
    destino = os.path.join(PASTA_ESEF, "filings", f"{lei}.json")
    try:
        js = _esef_get(url, destino, dias=7)
    except Exception:                                  # outra forma do filtro (a simples)
        js = _esef_get(f"{ESEF_URL}/api/filings?include=entity&page%5Bsize%5D=200&filter%5Bentity.identifier%5D={lei}",
                       destino, dias=7)
    inc = {str(e.get("id")): (e.get("attributes") or {}) for e in (js.get("included") or [])
           if e.get("type") == "entity"}
    ent = {"nome": "", "pais": "", "rel": []}
    for f in js.get("data") or []:
        a = f.get("attributes") or {}
        ju = a.get("json_url")
        rid = str((((f.get("relationships") or {}).get("entity") or {}).get("data") or {}).get("id"))
        ea = inc.get(rid, {})
        if not ju or not a.get("period_end"):
            continue
        if txt(ea.get("identifier")).upper() != lei and lei not in ju:   # o filtro falhou: não é desta empresa
            continue
        ent["nome"] = ent["nome"] or txt(ea.get("name"))
        ent["pais"] = ent["pais"] or txt(a.get("country"))
        ent["rel"].append({"fim": a["period_end"][:10], "pub": txt(a.get("published"))[:10] or None,
                           "json": ju, "erros": a.get("error_count") or 0, "added": txt(a.get("date_added"))[:10]})
    return ent


def _esef_lei_de_nome(nome):
    """LEI pelo nome (só se o nome sem formas jurídicas for exatamente igual ao de uma empresa do repositório)."""
    import urllib.parse
    k = _nome_chave(nome)
    if not k:
        return None
    pal = k.split(":", 1)[1].split()
    if not pal:
        return None
    filtro = json.dumps([{"name": "name", "op": "ilike", "val": f"%{pal[0]}%"}])
    js = _esef_get(f"{ESEF_URL}/api/entities?page%5Bsize%5D=200&filter=" + urllib.parse.quote(filtro),
                   os.path.join(PASTA_ESEF, "nomes", re.sub(r"[^a-z0-9]", "_", pal[0]) + ".json"), dias=30)
    cand = {txt((e.get("attributes") or {}).get("identifier")).upper() for e in (js.get("data") or [])
            if _nome_chave((e.get("attributes") or {}).get("name")) == k}
    cand = {c for c in cand if re.fullmatch(r"[A-Z0-9]{20}", c)}
    return cand.pop() if len(cand) == 1 else None


def _esef_data(s, fim_excl=True):
    """«2023-01-01T00:00:00» (fim exclusivo do XBRL) -> date de 31/12/2022; «2022-12-31» -> essa data."""
    from datetime import datetime as dt_, timedelta
    s = txt(s)
    if "T" in s:
        d = dt_.strptime(s[:19], "%Y-%m-%dT%H:%M:%S")
        if fim_excl and d.hour == 0 and d.minute == 0 and d.second == 0:
            d -= timedelta(days=1)
        return d.date()
    return dt_.strptime(s[:10], "%Y-%m-%d").date()


def _esef_extrair(js, fim_rel):
    """Do xBRL-JSON de um relatório: {data de fim do ano: {eps, ni, ...}} (o ano do relatório e o anterior, que vem
    como comparativo) e a moeda."""
    from datetime import datetime as dt_
    fim_rel = dt_.strptime(fim_rel[:10], "%Y-%m-%d").date()
    alvo = {}
    for k, cs in ESEF_FLUXO.items():
        for pri, c in enumerate(cs):
            alvo.setdefault(c, []).append((k, pri, True))
    for k, cs in ESEF_INSTANTE.items():
        for pri, c in enumerate(cs):
            alvo.setdefault(c, []).append((k, pri, False))
    anos, moedas = {}, {}
    for f in (js.get("facts") or {}).values():
        dm = f.get("dimensions") or {}
        c = dm.get("concept")
        if c not in alvo or set(dm) - _ESEF_DIM_OK:
            continue
        try:
            v = float(f.get("value"))
        except (TypeError, ValueError):
            continue
        per = txt(dm.get("period"))
        try:
            if "/" in per:
                a, b = per.split("/")
                ini, fim = _esef_data(a, False), _esef_data(b)
                dur = (fim - ini).days
            else:
                ini, fim, dur = None, _esef_data(per), None
        except Exception:
            continue
        if abs((fim - fim_rel).days) > 10 and abs((fim_rel - fim).days - 365) > 12:
            continue                                   # só o ano do relatório e o anterior
        un = txt(dm.get("unit"))
        m_ = re.match(r"iso4217:([A-Z]{3})", un)
        if m_:
            moedas[m_.group(1)] = moedas.get(m_.group(1), 0) + 1
        for k, pri, fluxo in alvo[c]:
            if fluxo != (dur is not None):
                continue
            if fluxo and not 330 <= dur <= 400:
                continue
            chave = fim if abs((fim - fim_rel).days) <= 10 else ("ant", fim)
            reg = anos.setdefault(chave, {})
            if k not in reg or pri < reg[k][1]:
                reg[k] = (v, pri)
    moeda = max(moedas, key=moedas.get) if moedas else None
    out = {}
    for chave, reg in anos.items():
        fim = chave[1] if isinstance(chave, tuple) else chave
        r = {k: v[0] for k, v in reg.items()}
        if "eps" not in r and "ni" not in r:
            continue
        if "bruto" not in r and "receita" in r and "custo" in r:
            r["bruto"] = r["receita"] - abs(r["custo"])
        if "div_lp" in r or "div_cp" in r:
            r["divida"] = (r.get("div_lp") or 0.0) + (r.get("div_cp") or 0.0)
        elif "div_tot" in r:
            r["divida"] = r["div_tot"]
        else:
            r["divida"] = None
        if r.get("juros") is not None:
            r["juros"] = abs(r["juros"])
        out[(fim.isoformat(), isinstance(chave, tuple))] = r
    return out, moeda


def _esef_anuais(lei, ent):
    """Relatórios anuais de uma empresa no formato de _sec_anuais (+ moeda), com cópia local do que se extraiu."""
    from datetime import datetime as dt_, timedelta
    destino = os.path.join(PASTA_ESEF, f"{lei}.json")
    guard = _ler_json(destino, {})
    por_fim = {}
    for r in ent["rel"]:                               # um relatório por ano: o com menos erros; depois o 1.º
        a = por_fim.get(r["fim"])
        if a is None or (r["erros"], r["added"] or "9") < (a["erros"], a["added"] or "9"):
            por_fim[r["fim"]] = r
    anos, moeda, mudou = {}, guard.get("moeda"), False
    for fim, r in sorted(por_fim.items()):
        k = f"{fim}|{r['json']}"
        if k not in guard:
            try:
                js = _esef_get(ESEF_URL + r["json"])
                ex_, m_ = _esef_extrair(js, fim)
                guard[k] = {"anos": [[f, ant, v] for (f, ant), v in ex_.items()], "moeda": m_}
            except Exception:                          # sem net / ficheiro com erro: tenta da próxima vez
                continue
            mudou = True
        g = guard[k]
        moeda = g.get("moeda") or moeda
        pub = r.get("pub")
        fim_d = dt_.strptime(fim, "%Y-%m-%d")
        if pub:
            try:
                p_ = dt_.strptime(pub, "%Y-%m-%d")
                pub = p_ if fim_d < p_ <= fim_d + timedelta(days=400) else None
            except ValueError:
                pub = None
        disp = pub or (fim_d + timedelta(days=ESEF_DIAS_PUB))
        for f, ant, v in g.get("anos", []):
            if ant:                                    # comparativo: só serve para anos sem relatório próprio
                if f in por_fim or f in anos:
                    continue
                d_ = dt_.strptime(f, "%Y-%m-%d") + timedelta(days=ESEF_DIAS_PUB)
            else:
                d_ = disp
            anos[f] = {**v, "fim": f, "disponivel": d_.strftime("%Y-%m-%d")}
    if mudou:
        guard["moeda"] = moeda
        os.makedirs(PASTA_ESEF, exist_ok=True)
        _gravar_json(destino, guard)
    out = sorted(anos.values(), key=lambda r: r["disponivel"])
    return out, moeda


def _esef_isin(t):
    try:
        i = txt(_yf().Ticker(t).isin).upper()
        return i if re.fullmatch(r"[A-Z]{2}[A-Z0-9]{9}[0-9]", i) else None
    except Exception:
        return None


def _esef_lei_de_isin(isin):
    """LEI do emitente pelo ISIN (GLEIF). Sem net: levanta a exceção (para não gravar «não encontrado»)."""
    js = _esef_get(f"https://api.gleif.org/api/v1/lei-records?filter%5Bisin%5D={isin}",
                   os.path.join(PASTA_ESEF, "gleif", f"{isin}.json"), dias=365)
    d = js.get("data") or []
    return txt(d[0].get("id")).upper() if d else None


def esef_preparar(linhas_acoes, ini, fim):
    """{ticker Yahoo: SecHist} para as ações europeias com relatórios ESEF (as dos EUA já vêm da SEC)."""
    import pandas as pd
    from datetime import timedelta
    if not linhas_acoes:
        return {}
    mapa = _ler_json(FICH_ESEF_LEI, {})                # ticker -> LEI ou "" (procurado e não encontrado)
    mudou, pares, ents, falhas = False, [], {}, 0
    for d in linhas_acoes:
        t = txt(d.get("ticker")).upper()
        if not t or falhas >= 3:                       # sem net / repositório em baixo: não insiste
            continue
        lei = mapa.get(t)
        if lei is None:                                # 1.ª vez: ISIN -> LEI (GLEIF); senão pelo nome exato
            isin = _esef_isin(t)
            if isin and isin[:2] in ("US", "JP", "KR", "TW", "CN", "HK", "CA", "AU", "IN", "BR"):
                lei = ""                               # fora da Europa: não há ESEF
            else:
                try:
                    lei = (_esef_lei_de_isin(isin) if isin else None) or _esef_lei_de_nome(d.get("nome")) or ""
                except Exception:                      # sem net: tenta outra vez da próxima
                    falhas += 1
                    continue
            mapa[t] = lei
            mudou = True
        if lei:
            try:
                e = _esef_filings(lei)
            except Exception:
                falhas += 1
                continue
            if e["rel"]:
                ents[lei] = e
                pares.append((d, t, lei))
    if mudou:
        os.makedirs(PASTA_ESEF, exist_ok=True)
        _gravar_json(FICH_ESEF_LEI, mapa)
    if falhas >= 3:
        print("   ESEF: o repositório filings.xbrl.org não respondeu (sem net?): sem fundamentos europeus desta vez.")
    if not pares:
        return {}
    print(f"   ESEF: relatórios anuais de {len(pares)} empresas europeias (a 1.ª vez demora) ...")
    yf = _yf()
    hist = {}                                          # ticker das cotações: a bolsa principal, se houver
    for d, t, _ in pares:
        m_ = re.search(r"principal: .* (\S+)$", txt(d.get("bolsa")))
        hist[t] = m_.group(1).upper() if m_ else t
    pedir = sorted(set(hist.values()))
    try:
        h = _por_ticker(yf.download(pedir, start=(ini - timedelta(days=4 * 366)).strftime("%Y-%m-%d"),
                                    end=(fim + timedelta(days=1)).strftime("%Y-%m-%d"), interval="1d",
                                    group_by="ticker", auto_adjust=False, progress=False, threads=True), pedir)
    except Exception:
        h = {}
    fx_cache = {}

    def fx(de, para):
        if de == para:
            return 1.0
        k = f"{de}{para}=X"
        if k not in fx_cache:
            try:
                s = yf.download(k, start=(ini - timedelta(days=4 * 366 + 15)).strftime("%Y-%m-%d"),
                                end=(fim + timedelta(days=1)).strftime("%Y-%m-%d"), interval="1d",
                                auto_adjust=True, progress=False)["Close"]
                if hasattr(s, "columns"):
                    s = s.iloc[:, 0]
                s = s.dropna()
                if not len(s):
                    fx_cache[k] = None
                elif not isinstance(s.index, pd.DatetimeIndex):   # sem datas: câmbio fixo (o último)
                    fx_cache[k] = float(s.iloc[-1])
                else:
                    if getattr(s.index, "tz", None) is not None:
                        s.index = s.index.tz_localize(None)
                    s.index = s.index.normalize()
                    fx_cache[k] = s[~s.index.duplicated(keep="last")]
            except Exception:
                fx_cache[k] = None
        return fx_cache[k]
    out, sem = {}, []
    for d, t, lei in pares:
        try:
            an, moeda_r = _esef_anuais(lei, ents[lei])
            th = hist[t]
            df = h.get(th) if hasattr(h, "get") else None
            if (not an or df is None or "Close" not in df) and th != t:
                th = t
                df = _por_ticker(yf.download([t], start=(ini - timedelta(days=4 * 366)).strftime("%Y-%m-%d"),
                                             end=(fim + timedelta(days=1)).strftime("%Y-%m-%d"), interval="1d",
                                             group_by="ticker", auto_adjust=False, progress=False), [t]).get(t)
            if not an or df is None or "Close" not in df or not moeda_r:
                sem.append(t)
                continue
            p = df["Close"].dropna()
            if getattr(p.index, "tz", None) is not None:
                p.index = p.index.tz_localize(None)
            p.index = p.index.normalize()
            p = p[~p.index.duplicated(keep="last")]
            try:
                mt = (yf.Ticker(th).fast_info.get("currency") if th != t else txt(d.get("moeda"))) or ""
            except Exception:
                mt = txt(d.get("moeda"))
            fat = 0.01 if mt in ("GBp", "GBX", "ZAc", "ILA") else 1.0
            mt = {"GBp": "GBP", "GBX": "GBP", "ZAc": "ZAR", "ILA": "ILS"}.get(mt, mt.upper() or "EUR")
            f_ = fx(mt, moeda_r)
            if f_ is None:
                sem.append(t)
                continue
            p = p * fat * (f_ if isinstance(f_, float) else f_.reindex(p.index.union(f_.index)).ffill().bfill().reindex(p.index))
            try:
                sp = yf.Ticker(th).splits
                if getattr(sp.index, "tz", None) is not None:
                    sp.index = sp.index.tz_localize(None)
            except Exception:
                sp = None
            out[t] = SecHist(an, sp, p.dropna(), txt(d.get("saude")).startswith("não se aplica"), None)
        except Exception:
            sem.append(t)
    print(f"   ESEF: fundamentos históricos de {len(out)} empresas europeias"
          + (f" (sem dados utilizáveis: {', '.join(sem[:12])}{'…' if len(sem) > 12 else ''})" if sem else "") + ".")
    return out


# ======================================================= analistas vs valor justo, comparados com o que aconteceu (V158)
FICH_ALVO_JUSTO = os.path.join(PASTA, BASE + "_analistas_vs_justo.xlsx")
AJ_MIN_FIRMAS = 3                  # consenso só com pelo menos 3 casas de análise
AJ_DATAS, AJ_PASSO_MESES = 12, 3   # V159: 12 datas de partida, de 3 em 3 meses (a mais recente há 1 ano)
AJ_MAX = 3.0                       # potencial acima de +300% ou abaixo de −90%: dado errado (fica de fora)


def _aj_historico(tk):
    """Histórico de alterações dos analistas (data, casa, alvo) e desdobramentos de um ticker do Yahoo."""
    import pandas as pd
    ud = getattr(tk, "upgrades_downgrades", None)
    if ud is None or not len(ud) or "currentPriceTarget" not in ud.columns:
        ud = None
    else:
        ud = ud.reset_index()
        col_d = "GradeDate" if "GradeDate" in ud.columns else ud.columns[0]
        ud = ud.rename(columns={col_d: "_dia"})
        ud["_dia"] = pd.to_datetime(ud["_dia"], errors="coerce")
        if getattr(ud["_dia"].dt, "tz", None) is not None:
            ud["_dia"] = ud["_dia"].dt.tz_localize(None)
        ud["_alvo"] = pd.to_numeric(ud["currentPriceTarget"], errors="coerce")
        ud = ud[(ud["_alvo"] > 0) & ud["_dia"].notna()].sort_values("_dia")
    try:
        sp = tk.splits
        if getattr(sp.index, "tz", None) is not None:
            sp.index = sp.index.tz_localize(None)
    except Exception:
        sp = None
    return ud, sp


def _aj_consenso(ud, sp, dia):
    """Preço-alvo médio em «dia»: o último alvo de cada casa de análise nos 12 meses antes, em ações de hoje
    (desdobramentos depois de «dia» corrigidos, como as cotações do Yahoo). -> (alvo, nº de casas)."""
    import pandas as pd
    if ud is None or not len(ud):
        return None, 0
    dia = pd.Timestamp(dia)
    x = ud[(ud["_dia"] <= dia) & (ud["_dia"] > dia - pd.Timedelta(days=365))]
    if not len(x):
        return None, 0
    ult = x.groupby("Firm").tail(1)
    if len(ult) < AJ_MIN_FIRMAS:
        return None, len(ult)
    fator = 1.0
    if sp is not None and len(sp):
        s_ = sp[(sp.index > dia) & (sp > 0)]
        fator = float(s_.prod()) if len(s_) else 1.0
    return float(ult["_alvo"].mean()) / (fator or 1.0), len(ult)


def comparar_alvo_justo():
    """▶ ANALISTAS vs VALOR JUSTO (folha Evolução): em 12 datas (de 3 em 3 meses, a mais recente há 1 ano), quanto
    prometiam o preço-alvo dos analistas e o valor justo (lucros × PER habitual, relatórios SEC / ESEF), e quanto
    subiu de facto cada ação no ano seguinte. V159: preço e analistas da bolsa principal (mesma moeda). Ficheiro à parte."""
    import pandas as pd
    from datetime import date as date_, timedelta
    from openpyxl import Workbook
    from openpyxl.chart import ScatterChart, BarChart, Reference, Series
    from openpyxl.chart.marker import Marker
    ex = Excel(FICHEIRO)
    carregar_pesos(ex)
    ultima = get_column_letter(len(COLS))
    tab = ex.ops([{"op": "ler", "folha": 1, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"}])[0]
    acoes = [d for d in linhas_por_cabecalho(tab, 1) if txt(d.get("tipo")) != "ETF"]
    por_t = {}
    for d in acoes:
        por_t.setdefault(txt(d.get("ticker")).upper(), d)
    hist = {}                                          # ticker da tabela -> ticker da bolsa principal
    for t, d in por_t.items():
        m_ = re.search(r"principal: .* (\S+)$", txt(d.get("bolsa")))
        hist[t] = m_.group(1).upper() if m_ else t
    hoje = date_.today()
    anos_dados = 1 + (AJ_DATAS - 1) * AJ_PASSO_MESES / 12 + 1
    ini = hoje - timedelta(days=int(anos_dados * 365.25) + 30)
    yf = _yf()
    pedir = sorted(set(hist.values()))
    print(f"Analistas vs valor justo: cotações de {len(pedir)} ações (bolsa principal) ...")
    try:
        h = _por_ticker(yf.download(pedir, start=ini.strftime("%Y-%m-%d"), end=(hoje + timedelta(days=1)).strftime("%Y-%m-%d"),
                                    interval="1d", group_by="ticker", auto_adjust=False, progress=False, threads=True), pedir)
    except Exception as e:
        print(f"   sem cotações: {e}")
        sys.exit(1)
    P = {}
    for t, th in hist.items():
        df = h.get(th) if hasattr(h, "get") else None
        if df is None or "Close" not in df:
            continue
        s = df["Close"].dropna()
        if len(s) < 300:
            continue
        if getattr(s.index, "tz", None) is not None:
            s.index = s.index.tz_localize(None)
        s.index = s.index.normalize()
        P[t] = _sim_limpar(s[~s.index.duplicated(keep="last")])
    if not P:
        print("   sem cotações.")
        sys.exit(1)
    fim = max(s.index[-1] for s in P.values())
    datas = [pd.Timestamp(fim) - pd.DateOffset(years=1) - pd.DateOffset(months=AJ_PASSO_MESES * k)
             for k in range(AJ_DATAS)][::-1]
    print(f"   {len(datas)} datas de partida, de {datas[0]:%d/%m/%Y} a {datas[-1]:%d/%m/%Y}, cada uma com 1 ano à frente")
    sec = {}
    try:
        sec = sec_preparar([por_t[t] for t in P], datas[0].to_pydatetime().date() - timedelta(days=30), hoje)
    except Exception as e:
        print(f"   SEC/ESEF: {e}")
    print(f"   valor justo: {len(sec)} ações com relatórios · a ler o histórico dos analistas ...")
    analistas = {}
    for i, t in enumerate(P):
        try:
            analistas[t] = _aj_historico(yf.Ticker(hist[t]))
        except Exception:
            analistas[t] = (None, None)
        if (i + 1) % 25 == 0:
            print(f"   {i + 1}/{len(P)} ...")
    casos = []
    for dia in datas:
        dia1 = dia + pd.DateOffset(years=1)
        for t, s in P.items():
            s0, s1 = s[s.index <= dia], s[s.index <= dia1]
            if not len(s0) or s0.index[-1] < dia - pd.Timedelta(days=10) or s1.index[-1] < dia1 - pd.Timedelta(days=10):
                continue
            p0, p1 = float(s0.iloc[-1]), float(s1.iloc[-1])
            ud, sp = analistas.get(t, (None, None))
            alvo, nf = _aj_consenso(ud, sp, dia)
            pa = alvo / p0 - 1 if alvo else None
            if pa is not None and not (-0.9 < pa < AJ_MAX):
                pa = None
            jp = None
            if t in sec:
                try:
                    jp = sec[t].em(dia).get("justo_pct")
                except Exception:
                    jp = None
                if jp is not None and not (-0.99 < jp < 10):
                    jp = None
            if pa is None and jp is None:
                continue
            casos.append({"dia": dia, "t": t, "th": hist[t], "nome": txt(por_t[t].get("nome")) or t, "p0": p0,
                          "p1": p1, "alvo": alvo, "nf": nf, "pa": pa, "pj": jp, "real": p1 / p0 - 1})
    if not casos:
        print("   sem casos.")
        sys.exit(1)
    C = pd.DataFrame(casos)

    def stats(df, chave):
        """Por data: correlação de ordem e terço de cima − terço de baixo; depois a média das datas."""
        x = df[df[chave].notna()]
        porr = []
        for dia, g in x.groupby("dia"):
            if len(g) < 6:
                continue
            q = g.sort_values(chave)
            k = max(2, len(q) // 3)
            porr.append({"dia": dia, "n": len(g), "rho": float(g[chave].rank().corr(g["real"].rank())),
                         "top": float(q["real"].iloc[-k:].median()), "baixo": float(q["real"].iloc[:k].median())})
        if not porr:
            return None, []
        R = pd.DataFrame(porr)
        return {"n": int(len(x)), "datas": len(R), "prev": float(x[chave].median()), "real": float(x["real"].median()),
                "rho": float(R["rho"].mean()), "rho_pos": float((R["rho"] > 0).mean()),
                "top": float(R["top"].mean()), "baixo": float(R["baixo"].mean()),
                "erro": float((x[chave] - x["real"]).abs().median()),
                "sentido": float(((x[chave] > 0) == (x["real"] > 0)).mean())}, porr
    amb = C[C["pa"].notna() & C["pj"].notna()]
    r_a, pd_a = stats(C, "pa")
    r_j, pd_j = stats(C, "pj")
    r_aa, pd_aa = stats(amb, "pa")
    r_jj, pd_jj = stats(amb, "pj")
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    ws["A1"] = (f"Analistas vs valor justo: {len(datas)} datas de partida ({datas[0]:%d/%m/%Y} a {datas[-1]:%d/%m/%Y}), "
                f"cada uma com o ano seguinte · {len(C)} casos")
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("Preço-alvo = média do último alvo de cada casa de análise nos 12 meses anteriores (histórico do Yahoo, "
                f"≥ {AJ_MIN_FIRMAS} casas). Valor justo = lucro por ação × PER habitual de 3 anos, só com os relatórios já "
                "publicados (SEC nos EUA, ESEF na Europa). Preços e alvos da bolsa principal, na mesma moeda; variação real "
                "= só o preço, 1 ano depois. Correlação e terços: calculados em cada data e depois a média das datas.")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A2"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("A2:K2")
    ws.row_dimensions[2].height = 62
    cab = ["", "Casos", "Datas", "Previsto (mediana)", "Real (mediana)", "Correlação de ordem (média das datas)",
           "Datas com correlação positiva", "Real: terço com maior previsão", "Real: terço com menor previsão",
           "Erro (mediana)", "Acertou no sentido"]
    for j, c_ in enumerate(cab, 1):
        c = ws.cell(4, j, c_)
        c.font, c.alignment = Font(bold=True, color="FFFFFF"), Alignment(wrap_text=True, horizontal="center")
        c.fill = PatternFill("solid", fgColor="1F4E79")
    ws.row_dimensions[4].height = 48
    rot = [("Analistas (todas com alvo)", r_a), ("Valor justo (todas com relatórios)", r_j),
           ("Analistas (os mesmos casos)", r_aa), ("Valor justo (os mesmos casos)", r_jj)]
    for i, (nm, r) in enumerate(rot):
        ws.cell(5 + i, 1, nm).font = Font(bold=True)
        if not r:
            ws.cell(5 + i, 2, "poucos casos")
            continue
        vals = [r["n"], r["datas"], r["prev"], r["real"], r["rho"], r["rho_pos"], r["top"], r["baixo"], r["erro"], r["sentido"]]
        for j, v in enumerate(vals, 2):
            c = ws.cell(5 + i, j, v)
            c.number_format = "0" if j in (2, 3) else "0.00" if j == 6 else "0%"
            c.alignment = Alignment(horizontal="center")
    ws["A10"] = ("Como ler: compara sobretudo as duas últimas linhas (os mesmos casos). Correlação de ordem > 0 = as ações "
                 "com maior previsão subiram mais no ano seguinte; o melhor sinal tem maior correlação, mais datas positivas "
                 "e maior diferença entre o terço de cima e o de baixo. «Previsto» acima do «Real» = otimismo.")
    ws["A10"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A10"].font = Font(size=9)
    ws.merge_cells("A10:K10")
    ws.row_dimensions[10].height = 46
    ws.column_dimensions["A"].width = 32
    for col in "BCDEFGHIJK":
        ws.column_dimensions[col].width = 13
    # por data (os mesmos casos)
    wdd = wb.create_sheet("Por data")
    for j, c_ in enumerate(["Data de partida", "Casos", "Correlação analistas", "Correlação valor justo",
                            "Analistas: cima − baixo", "Valor justo: cima − baixo"], 1):
        c = wdd.cell(1, j, c_)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(wrap_text=True, horizontal="center")
    ma = {x["dia"]: x for x in (pd_aa if pd_aa else pd_a)}   # os mesmos casos; sem eles, todos
    mj = {x["dia"]: x for x in (pd_jj if pd_jj else pd_j)}
    dias_ = sorted(set(ma) | set(mj))
    for i, dia in enumerate(dias_, 2):
        a_, j_ = ma.get(dia), mj.get(dia)
        vals = [dia.to_pydatetime(), (a_ or j_)["n"], a_ and a_["rho"], j_ and j_["rho"],
                a_ and a_["top"] - a_["baixo"], j_ and j_["top"] - j_["baixo"]]
        for j, v in enumerate(vals, 1):
            c = wdd.cell(i, j, v)
            c.number_format = "dd/mm/yyyy" if j == 1 else "0" if j == 2 else "0.00" if j in (3, 4) else "0%"
    for col, w in zip("ABCDEF", (13, 8, 13, 13, 13, 13)):
        wdd.column_dimensions[col].width = w
    if dias_:
        ch = BarChart()
        ch.title, ch.style = "Correlação de ordem em cada data", 10
        ch.add_data(Reference(wdd, min_col=3, max_col=4, min_row=1, max_row=len(dias_) + 1), titles_from_data=True)
        ch.set_categories(Reference(wdd, min_col=1, min_row=2, max_row=len(dias_) + 1))
        ch.x_axis.number_format = "mm/yy"
        ch.x_axis.tickLblPos = "low"
        ch.x_axis.delete = ch.y_axis.delete = False
        ch.series[0].graphicalProperties.solidFill = "1F4E79"
        ch.series[1].graphicalProperties.solidFill = "1E7B34"
        ch.width, ch.height = 22, 9
        ws.add_chart(ch, "A12")
    # casos
    wd = wb.create_sheet("Casos")
    cab2 = ["Data de partida", "Ticker (tabela)", "Ticker usado", "Empresa", "Preço de partida", "Preço 1 ano depois",
            "Variação real", "Preço-alvo de então", "Casas de análise", "Previsão analistas", "Previsão valor justo",
            "Erro analistas", "Erro valor justo"]
    for j, c_ in enumerate(cab2, 1):
        c = wd.cell(1, j, c_)
        c.font, c.fill = Font(bold=True, color="FFFFFF"), PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(wrap_text=True, horizontal="center")
    C = C.sort_values(["dia", "real"], ascending=[True, False])
    for i, l in enumerate(C.to_dict("records"), 2):
        vals = [l["dia"].to_pydatetime(), l["t"], l["th"], l["nome"], l["p0"], l["p1"], l["real"],
                l["alvo"] if pd.notna(l["alvo"]) else None, l["nf"] or None,
                l["pa"] if pd.notna(l["pa"]) else None, l["pj"] if pd.notna(l["pj"]) else None,
                (l["pa"] - l["real"]) if pd.notna(l["pa"]) else None, (l["pj"] - l["real"]) if pd.notna(l["pj"]) else None]
        for j, v in enumerate(vals, 1):
            c = wd.cell(i, j, v)
            if j == 1:
                c.number_format = "dd/mm/yyyy"
            elif j in (5, 6, 8):
                c.number_format = "0.00"
            elif j >= 7 and j != 9:
                c.number_format = "0%"
    wd.freeze_panes = "E2"
    wd.auto_filter.ref = f"A1:M{len(C) + 1}"
    for col, w in zip("ABCDEFGHIJKLM", (11, 10, 10, 24, 10, 10, 10, 10, 8, 10, 10, 10, 10)):
        wd.column_dimensions[col].width = w
    n = len(C) + 1
    for k_, (col, titulo, cor) in enumerate(((10, "Analistas: previsto (x) vs real (y)", "1F4E79"),
                                             (11, "Valor justo: previsto (x) vs real (y)", "1E7B34"))):
        ch = ScatterChart()
        ch.title, ch.style = titulo, 13
        ch.x_axis.title, ch.y_axis.title = "previsto", "real"
        ch.x_axis.number_format = ch.y_axis.number_format = "0%"
        ch.x_axis.tickLblPos = ch.y_axis.tickLblPos = "low"
        ch.x_axis.delete = ch.y_axis.delete = False
        sr = Series(Reference(wd, min_col=7, min_row=2, max_row=n), Reference(wd, min_col=col, min_row=2, max_row=n),
                    title=titulo)
        sr.marker = Marker(symbol="circle", size=4)
        sr.marker.graphicalProperties.solidFill = cor
        sr.marker.graphicalProperties.line.solidFill = cor
        sr.graphicalProperties.line.noFill = True
        ch.series.append(sr)
        ch.legend = None
        ch.width, ch.height = 14, 9
        ws.add_chart(ch, f"{'A' if k_ == 0 else 'F'}31")
    try:
        wb.save(FICH_ALVO_JUSTO)
    except PermissionError:
        print(f"   fecha o ficheiro {os.path.basename(FICH_ALVO_JUSTO)} e volta a carregar no botão.")
        sys.exit(1)
    for nm, r in rot:
        if r:
            print(f"   {nm:34} {r['n']:4d} casos · correlação {r['rho']:+.2f} (positiva em {r['rho_pos']:.0%} das datas)"
                  f" · terço de cima {r['top']:+.0%} / de baixo {r['baixo']:+.0%}")
    print(f"Pronto: {FICH_ALVO_JUSTO}")
    try:
        os.startfile(FICH_ALVO_JUSTO)
    except Exception:
        pass


# ======================================================= Comp Rent Títulos (V168): rentabilidade anual de cada título
FOLHA_CRT = "Comp Rent Títulos"
FICH_CRT = os.path.join(PASTA, BASE + "_comp_rent.json")
CRT_N = 20                         # títulos
CRT_L1 = 12                        # 1.ª linha dos títulos (cabeçalho na 11)
CRT_PTS = 400                      # pontos por curva (no máximo)
CRT_CN = 39                        # coluna AM: lista de nomes (para as listas de escolha)
CRT_C0 = 40                        # coluna AN: dados dos gráficos (escondidos)
CRT_D0 = 4                         # 1.ª linha dos dados (a 3 tem os nomes das séries)
CRT_PADRAO = {"anos": 5.0, "passo": 5, "tmin": 12, "pontos": 4, "meia": 3.0}   # V170: meia-vida 3 anos
CRT_COLS = ("Rent. anual simples (do início)", "Rent. anual ponderada (meia-vida)", "Ponderada − referência",
            "Volatilidade anual", "Ponderada ÷ volatilidade", "% do tempo acima da referência", "Ordem")
CRT_LBL = 37                       # 1.ª linha dos gráficos de cada título
CRT_H = 18                         # linhas por par de gráficos


def _crt_col(j, k):
    """Coluna dos dados: k = 'c' curva do título j, 'fx'/'fy' curva ajustada, 'mx'/'my' média ponderada, 'r' referência."""
    if k == "d":
        return CRT_C0
    if k == "c":
        return CRT_C0 + 1 + j
    if k in ("fx", "fy"):
        return CRT_C0 + 1 + CRT_N + 2 * j + (k == "fy")
    if k in ("mx", "my"):
        return CRT_C0 + 1 + 3 * CRT_N + 2 * j + (k == "my")
    return CRT_C0 + 1 + 5 * CRT_N                      # referência


def _nomes_das_folhas(wb):
    """[(nome, ticker)] pela ordem das folhas Ações e depois ETFs (nomes das fórmulas HYPERLINK)."""
    out, vistos = [], set()
    for fl_ in (FOLHA_A2, FOLHA_E2):
        if fl_ not in wb.sheetnames:
            continue
        w_ = wb[fl_]
        for r_ in range(L0 + 1, L0 + MAX_LINHAS + 1):
            t_, n_ = txt(w_.cell(r_, COL["ticker"]).value).upper(), w_.cell(r_, COL["nome"]).value
            if isinstance(n_, str) and n_.startswith("="):
                m_ = re.search(r',"((?:[^"]|"")*)"\)$', n_)
                n_ = m_.group(1).replace('""', '"') if m_ else None
            if t_ and isinstance(n_, str) and n_ and n_ not in vistos:
                vistos.add(n_)
                out.append((n_, t_))
    return out


def folha_comp_rent(wb):
    """Folha Comp Rent Títulos: para até 20 títulos, a rentabilidade anual média que se teria tido comprando em cada data
    (de há N anos até hoje menos um tempo mínimo) e mantendo até hoje; curva ajustada (3 ou 4 pontos), média com
    meia-vida (mais peso ao recente), comparação com uma referência e com o risco; gráfico geral e um por título."""
    from openpyxl.chart import ScatterChart, BarChart, Reference, Series
    from openpyxl.chart.marker import Marker
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.worksheet.datavalidation import DataValidation as DV_
    g = _ler_json(FICH_CRT, {})
    ws = wb.create_sheet(FOLHA_CRT)
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    fina = Side(style="thin", color="BFBFBF")
    borda = Border(left=fina, right=fina, top=fina, bottom=fina)
    ws["A1"] = "Comparar a rentabilidade dos títulos: quanto se teria ganho por ano, comprando em cada data e mantendo até hoje"
    ws["A1"].font = Font(size=14, bold=True)
    ws["B2"] = ("Escolhe até 20 títulos (listas com os nomes das folhas Ações e ETFs) e carrega em ▶ CALCULAR. Para cada "
                "data, de há «extensão» anos até hoje menos o «tempo mínimo», de «periodicidade» em «periodicidade» dias de "
                "bolsa, a curva mostra a rentabilidade anual média de comprar nessa data e manter até hoje (em euros, com "
                "dividendos). Perto de hoje a anualização multiplica o ruído (1 mês com +5% = +80%/ano), por isso o tempo "
                "mínimo deve ser ≥ 6 meses (12 por omissão). A curva ajustada liga 3 ou 4 pontos (a média da curva em cada "
                "terço / quarto do período). A média ponderada usa as rentabilidades mensais, cada mês com metade do peso "
                "do mês «meia-vida» anos mais recente (3 anos por omissão: análise de 5 anos 1,5 a 2; de 10-15 anos 3; de 20 ou mais 4 a 5): mais peso ao passado recente. «Ponderada − "
                "referência»: pontos por ano acima / abaixo do título de referência (mesmo cálculo). «Ponderada ÷ "
                "volatilidade»: rentabilidade por unidade de risco. «% do tempo acima»: em quantas datas da curva comprar o "
                "título rendeu mais do que comprar a referência. Passado não garante futuro.")
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B2"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("B2:K2")
    ws.row_dimensions[2].height = 118
    nomes = _nomes_das_folhas(wb)
    for i, (n_, t_) in enumerate(nomes[:600], 1):
        ws.cell(i, CRT_CN, n_)
    ws.column_dimensions[get_column_letter(CRT_CN)].hidden = True
    lista = f"=${get_column_letter(CRT_CN)}$1:${get_column_letter(CRT_CN)}${max(1, len(nomes))}"
    p = {**CRT_PADRAO, **(g.get("p") or {})}
    ref_nome = p.get("ref") or next((n_ for n_, t_ in nomes if t_ == "SXR8.DE"), None) \
        or next((n_ for n_, t_ in nomes if "S&P 500" in n_), None)
    entradas = (("Extensão da análise (anos)", p["anos"], "0.00"), ("Periodicidade (dias de bolsa entre pontos)", p["passo"], "0"),
                ("Tempo mínimo até hoje (meses; ≥ 6)", p["tmin"], "0"), ("Pontos da curva ajustada (3 ou 4)", p["pontos"], "0"),
                ("Meia-vida da média ponderada (anos; 3 por omissão)", p["meia"] or 3.0, "0.00"),
                ("Referência (para comparar)", ref_nome, "@"))
    for i, (rot, v_, fm) in enumerate(entradas):
        r = 4 + i
        ws.cell(r, 2, rot).font = Font(bold=True, size=9)
        c = ws.cell(r, 4, v_)
        c.fill, c.border, c.number_format = amarelo, borda, fm
        c.font, c.alignment = Font(bold=True), Alignment(horizontal="center")
    ws.cell(4, 2).alignment = Alignment(horizontal="left")
    dv = DV_(type="list", formula1=lista, allow_blank=True)
    dv.error, dv.errorTitle = "Escolhe um nome da lista.", "Títulos"
    ws.add_data_validation(dv)
    dv.add("D9")
    dv.add(f"C{CRT_L1}:C{CRT_L1 + CRT_N - 1}")
    ws.merge_cells("D9:F9")
    botao(ws, "F4", "▶ CALCULAR", "painelacoes://comprent", "1F4E79", "16365C", 12)
    ws.merge_cells("F4:G5")
    ws["F7"] = g.get("gerado") or "Ainda não calculado."
    ws["F7"].font = Font(size=9, color="7030A0")
    cab = ["N.º", "Título (escolhe na lista)", "Ticker"] + list(CRT_COLS)
    for j, h_ in enumerate(cab, 2):
        c = ws.cell(CRT_L1 - 1, j, h_)
        c.font, c.fill = Font(bold=True, color="FFFFFF", size=9), PatternFill("solid", fgColor="1F4E79")
        c.alignment, c.border = Alignment(horizontal="center", vertical="center", wrap_text=True), borda
    ws.row_dimensions[CRT_L1 - 1].height = 45
    la, lt = get_column_letter(COL["nome"]), get_column_letter(COL["ticker"])
    tit = g.get("titulos") or []
    res = g.get("res") or []
    for i in range(CRT_N):
        r = CRT_L1 + i
        ws.cell(r, 2, i + 1).alignment = Alignment(horizontal="center")
        c = ws.cell(r, 3, tit[i] if i < len(tit) else None)
        c.fill, c.border, c.font = amarelo, borda, Font(bold=True)
        ws.cell(r, 4, f'=IF(TRIM(C{r})="","",IFERROR(INDEX(\'{FOLHA_A2}\'!${lt}:${lt},MATCH(C{r},\'{FOLHA_A2}\'!${la}:${la},0)),'
                      f'IFERROR(INDEX(\'{FOLHA_E2}\'!${lt}:${lt},MATCH(C{r},\'{FOLHA_E2}\'!${la}:${la},0)),"não está nas tabelas")))')
        ws.cell(r, 4).font, ws.cell(r, 4).alignment = Font(size=9, color="1F4E79"), Alignment(horizontal="center")
        lin = res[i] if i < len(res) else [None] * len(CRT_COLS)
        for j, fm in enumerate(("0.0%", "0.0%", "+0.0%;-0.0%;0.0%", "0.0%", "0.00", "0%", "0")):
            c = ws.cell(r, 5 + j, lin[j] if j < len(lin) else None)
            c.number_format, c.border, c.alignment = fm, borda, Alignment(horizontal="center")
    from openpyxl.formatting.rule import ColorScaleRule
    for col in ("F", "G", "I"):
        ws.conditional_formatting.add(f"{col}{CRT_L1}:{col}{CRT_L1 + CRT_N - 1}", ColorScaleRule(
            start_type="min", start_color="F8696B", mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="63BE7B"))
    for col, w in zip("ABCDEFGHIJK", (2, 5, 34, 11, 12, 12, 12, 11, 11, 11, 8)):
        ws.column_dimensions[col].width = w
    # dados (escondidos)
    for c_ in range(CRT_C0, _crt_col(0, "r") + 1):
        ws.column_dimensions[get_column_letter(c_)].hidden = True
    dados = g.get("dados") or {}
    ws.cell(CRT_D0 - 1, CRT_C0, "Data")
    for j in range(CRT_N):
        ws.cell(CRT_D0 - 1, _crt_col(j, "c"), (dados.get("nomes") or [None] * CRT_N)[j] or " ")
        ws.cell(CRT_D0 - 1, _crt_col(j, "fy"), "Curva ajustada")
        ws.cell(CRT_D0 - 1, _crt_col(j, "my"), "Média ponderada")
    ws.cell(CRT_D0 - 1, _crt_col(0, "r"), dados.get("ref_nome") or "Referência")
    for i, v_ in enumerate((dados.get("ref") or [])[:CRT_PTS]):
        ws.cell(CRT_D0 + i, _crt_col(0, "r"), v_)
    for i, lin in enumerate((dados.get("linhas") or [])[:CRT_PTS]):
        for k_, v_ in enumerate(lin):
            if k_ == 0 and isinstance(v_, str):
                try:
                    v_ = datetime.strptime(v_, "%Y-%m-%d")
                except ValueError:
                    pass
            ws.cell(CRT_D0 + i, CRT_C0 + k_, v_)
    for r in range(CRT_D0, CRT_D0 + CRT_PTS):
        ws.cell(r, CRT_C0).number_format = "dd/mm/yyyy"
    for j, blk in enumerate(dados.get("aux") or []):
        for i, (x_, y_) in enumerate(blk.get("f") or []):
            ws.cell(CRT_D0 + i, _crt_col(j, "fx"), datetime.strptime(x_, "%Y-%m-%d"))
            ws.cell(CRT_D0 + i, _crt_col(j, "fy"), y_)
        for i, (x_, y_) in enumerate(blk.get("m") or []):
            ws.cell(CRT_D0 + i, _crt_col(j, "mx"), datetime.strptime(x_, "%Y-%m-%d"))
            ws.cell(CRT_D0 + i, _crt_col(j, "my"), y_)

    def linha(sr, cor, w, tr=None, liso=False, marca=None):
        ln = LineProperties(solidFill=cor, w=w)
        if tr:
            ln.prstDash = tr
        sr.graphicalProperties = GraphicalProperties(ln=ln)
        sr.marker = Marker(symbol=marca or "none", size=7 if marca else None)
        if marca:
            sr.marker.graphicalProperties = GraphicalProperties(solidFill=cor)
            sr.marker.graphicalProperties.line.solidFill = cor
        sr.smooth = liso

    def eixos(ch):
        ch.visible_cells_only = False                  # os dados estão em colunas escondidas
        ch.x_axis.number_format, ch.y_axis.number_format = "mm/yy", "0%"
        ch.x_axis.tickLblPos = "low"
        ch.x_axis.delete = ch.y_axis.delete = False
        ch.x_axis.majorGridlines = None
    xs = Reference(ws, min_col=CRT_C0, min_row=CRT_D0, max_row=CRT_D0 + CRT_PTS - 1)
    cores = ("1F77B4", "FF7F0E", "2CA02C", "D62728", "9467BD", "8C564B", "E377C2", "7F7F7F", "BCBD22", "17BECF",
             "393B79", "AD494A", "637939", "8C6D31", "843C39", "7B4173", "3182BD", "E6550D", "31A354", "756BB1")
    ch = ScatterChart()                                 # gráfico geral
    ch.title, ch.style = "Rentabilidade anual de comprar em cada data e manter até hoje (todos)", 13
    for j in range(CRT_N):
        sr = Series(Reference(ws, min_col=_crt_col(j, "c"), min_row=CRT_D0 - 1, max_row=CRT_D0 + CRT_PTS - 1), xs,
                    title_from_data=True)
        linha(sr, cores[j], 15000)
        ch.series.append(sr)
    sr = Series(Reference(ws, min_col=_crt_col(0, "r"), min_row=CRT_D0 - 1, max_row=CRT_D0 + CRT_PTS - 1), xs,
                title_from_data=True)
    linha(sr, "000000", 28000, "dash")
    ch.series.append(sr)
    eixos(ch)
    ch.legend.position = "r"
    ch.width, ch.height = 30, 14
    ws.add_chart(ch, "M3")
    bc = BarChart()                                    # resumo: ponderada e simples
    bc.type, bc.style = "bar", 10
    bc.title = "Rentabilidade anual: ponderada (recente pesa mais) e simples"
    bc.add_data(Reference(ws, min_col=6, max_col=6, min_row=CRT_L1 - 1, max_row=CRT_L1 + CRT_N - 1), titles_from_data=True)
    bc.add_data(Reference(ws, min_col=5, max_col=5, min_row=CRT_L1 - 1, max_row=CRT_L1 + CRT_N - 1), titles_from_data=True)
    bc.set_categories(Reference(ws, min_col=3, min_row=CRT_L1, max_row=CRT_L1 + CRT_N - 1))
    bc.series[0].graphicalProperties.solidFill = "1E7B34"
    bc.series[1].graphicalProperties.solidFill = "A9D18E"
    bc.y_axis.number_format = "0%"
    bc.x_axis.scaling.orientation = "maxMin"
    bc.x_axis.delete = bc.y_axis.delete = False
    bc.y_axis.tickLblPos = bc.x_axis.tickLblPos = "low"
    bc.width, bc.height = 30, 13
    ws.add_chart(bc, "M33")
    for j in range(CRT_N):                             # um gráfico por título
        r0 = CRT_LBL + (j // 2) * CRT_H
        c0 = "B" if j % 2 == 0 else "G"
        lb = (g.get("rotulos") or [None] * CRT_N)[j] if j < len(g.get("rotulos") or []) else None
        ws[f"{c0}{r0}"] = lb or f"{j + 1}. (por escolher)"
        ws[f"{c0}{r0}"].font = Font(bold=True, size=10, color="1F4E79")
        ch = ScatterChart()
        ch.style = 13
        sr = Series(Reference(ws, min_col=_crt_col(j, "c"), min_row=CRT_D0 - 1, max_row=CRT_D0 + CRT_PTS - 1), xs,
                    title="Comprando nessa data")
        linha(sr, "1F4E79", 15000)
        ch.series.append(sr)
        sr = Series(Reference(ws, min_col=_crt_col(0, "r"), min_row=CRT_D0 - 1, max_row=CRT_D0 + CRT_PTS - 1), xs,
                    title_from_data=True)
        linha(sr, "A6A6A6", 15000, "dash")
        ch.series.append(sr)
        sr = Series(Reference(ws, min_col=_crt_col(j, "fy"), min_row=CRT_D0, max_row=CRT_D0 + 3),
                    Reference(ws, min_col=_crt_col(j, "fx"), min_row=CRT_D0, max_row=CRT_D0 + 3), title="Curva ajustada")
        linha(sr, "ED7D31", 28000, None, True, "circle")
        ch.series.append(sr)
        sr = Series(Reference(ws, min_col=_crt_col(j, "my"), min_row=CRT_D0, max_row=CRT_D0 + 1),
                    Reference(ws, min_col=_crt_col(j, "mx"), min_row=CRT_D0, max_row=CRT_D0 + 1), title="Média ponderada")
        linha(sr, "1E7B34", 28000, "sysDash")
        ch.series.append(sr)
        eixos(ch)
        ch.legend.position = "b"
        ch.width, ch.height = 15.5, 7.8
        ws.add_chart(ch, f"{c0}{r0 + 1}")


def calcular_comp_rent():
    """▶ CALCULAR (folha Comp Rent Títulos)."""
    import numpy as np
    import pandas as pd
    from datetime import date as date_, timedelta
    ex = Excel(FICHEIRO)
    ultima = get_column_letter(len(COLS))
    rr = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)]
                + [{"op": "ler", "folha": FOLHA_CRT, "rng": "D4:D9"},
                   {"op": "ler", "folha": FOLHA_CRT, "rng": f"C{CRT_L1}:C{CRT_L1 + CRT_N - 1}"}])
    por_nome, por_t = {}, {}
    for i, tab in zip((1, 2), rr[:2]):
        for d in linhas_por_cabecalho(tab, i):
            t = txt(d.get("ticker")).upper()
            if t:
                por_t.setdefault(t, d)
                por_nome.setdefault(txt(d.get("nome")).strip().lower(), d)
    v = [lin[0] for lin in rr[2]]
    p = {"anos": _num_ou(v[0], CRT_PADRAO["anos"]) or CRT_PADRAO["anos"],
         "passo": int(max(1, _num_ou(v[1], CRT_PADRAO["passo"]) or CRT_PADRAO["passo"])),
         "tmin": max(1.0, _num_ou(v[2], CRT_PADRAO["tmin"]) or CRT_PADRAO["tmin"]),
         "pontos": 3 if (_num_ou(v[3], 4) or 4) <= 3 else 4,
         "meia": _num_ou(v[4]) if _num_ou(v[4]) and _num_ou(v[4]) > 0 else None,
         "ref": txt(v[5]).strip() or None}
    titulos = [txt(lin[0]).strip() or None for lin in rr[3]]
    esc = []                                           # (posição, nome, d)
    for i, n_ in enumerate(titulos):
        if n_:
            d = por_nome.get(n_.lower())
            if d is None:
                print(f"   «{n_}» não está nas folhas Ações / ETFs: fica de fora.")
            else:
                esc.append((i, n_, d))
    d_ref = por_nome.get((p["ref"] or "").lower())
    if not esc:
        print("Comp Rent Títulos: escolhe pelo menos um título na lista.")
        sys.exit(1)
    hist = {}                                          # ticker da tabela -> ticker das cotações (bolsa principal)
    moedas = {}
    yf = _yf()
    for d in [x[2] for x in esc] + ([d_ref] if d_ref else []):
        t = txt(d.get("ticker")).upper()
        m_ = re.search(r"principal: .* (\S+)$", txt(d.get("bolsa")))
        th = t
        if m_ and m_.group(1).upper() != t:
            try:
                mo = yf.Ticker(m_.group(1).upper()).fast_info.get("currency")
                if mo:
                    th, moedas[m_.group(1).upper()] = m_.group(1).upper(), mo
            except Exception:
                pass
        hist[t] = th
        moedas.setdefault(t, txt(d.get("moeda")) or "EUR")
    hoje = date_.today()
    ini = hoje - timedelta(days=int(p["anos"] * 365.25) + 5)
    print(f"Comp Rent Títulos: {len(esc)} títulos, {p['anos']:.2f} anos, um ponto a cada {p['passo']} dias de bolsa, "
          f"até hoje menos {p['tmin']:.0f} meses ...")
    eur, _ = _sim_cotacoes_eur(sorted(set(hist.values()) | set(hist)), moedas, ini, hoje)
    P = {}
    for t, th in hist.items():
        s = eur[th] if th in eur.columns and eur[th].dropna().size > 20 else (eur[t] if t in eur.columns else None)
        if s is not None and s.dropna().size > 20:
            P[t] = s.dropna()
    if not P:
        print("   sem cotações.")
        sys.exit(1)
    fim = max(s.index[-1] for s in P.values())
    inicio = pd.Timestamp(fim) - pd.Timedelta(days=int(p["anos"] * 365.25))
    corte = pd.Timestamp(fim) - pd.DateOffset(months=int(round(p["tmin"])))
    dias = pd.DatetimeIndex(sorted(set().union(*[set(s.index) for s in P.values()])))
    dias = dias[(dias >= inicio) & (dias <= corte)]
    passo = p["passo"]
    while len(dias) / passo > CRT_PTS:
        passo += 1
    grelha = list(dias[::passo])
    if not grelha:
        print("   período curto demais (extensão menor do que o tempo mínimo).")
        sys.exit(1)
    meia = p["meia"] or 3.0

    def curva(s):
        pf = float(s.iloc[-1])
        out = []
        for d in grelha:
            s0 = s[s.index <= d]
            if not len(s0) or s.index[0] > d or s0.index[-1] < d - pd.Timedelta(days=7):
                out.append(None)
                continue
            anos_ = (s.index[-1] - d).days / 365.25
            out.append(float((pf / float(s0.iloc[-1])) ** (1 / anos_) - 1) if anos_ > 0 else None)
        return out

    def medidas(s):
        s = s[s.index >= inicio]
        if len(s) < 30:
            return None, None, None
        anos_ = (s.index[-1] - s.index[0]).days / 365.25
        simples = float((s.iloc[-1] / s.iloc[0]) ** (1 / anos_) - 1) if anos_ > 0.2 else None
        m = s.resample("ME").last()
        m = pd.concat([pd.Series([s.iloc[0]], index=[s.index[0]]), m]).sort_index()
        m = m[~m.index.duplicated(keep="last")]
        r = np.log(m / m.shift(1)).dropna()
        if not len(r):
            return simples, None, None
        idade = np.array([(s.index[-1] - d).days / 365.25 for d in r.index])
        w = 0.5 ** (idade / meia)
        pond = float(np.exp(12 * float((w * r.values).sum() / w.sum())) - 1)
        vol = float(np.log(s / s.shift(1)).dropna().std() * np.sqrt(252))
        return simples, pond, vol
    t_ref = txt(d_ref.get("ticker")).upper() if d_ref else None
    c_ref = curva(P[t_ref]) if t_ref in P else [None] * len(grelha)
    m_ref = medidas(P[t_ref]) if t_ref in P else (None, None, None)
    res = [[None] * len(CRT_COLS) for _ in range(CRT_N)]
    curvas, aux, rot = [[None] * len(grelha) for _ in range(CRT_N)], [{} for _ in range(CRT_N)], [None] * CRT_N
    nomes_s = [None] * CRT_N
    for j, n_, d in esc:
        t = txt(d.get("ticker")).upper()
        nomes_s[j] = n_
        if t not in P:
            rot[j] = f"{j + 1}. {n_}: sem cotações"
            continue
        c = curva(P[t])
        curvas[j] = c
        simples, pond, vol = medidas(P[t])
        val = [(g_, y_) for g_, y_ in zip(grelha, c) if y_ is not None]
        if val:
            partes = np.array_split(np.arange(len(val)), p["pontos"])
            aux[j]["f"] = [((val[pt[0]][0] + (val[pt[-1]][0] - val[pt[0]][0]) / 2).strftime("%Y-%m-%d"),
                            round(float(np.mean([val[k][1] for k in pt])), 5)) for pt in partes if len(pt)]
            aux[j]["m"] = [(val[0][0].strftime("%Y-%m-%d"), round(pond, 5) if pond is not None else None),
                           (val[-1][0].strftime("%Y-%m-%d"), round(pond, 5) if pond is not None else None)]
        dif = pond - m_ref[1] if pond is not None and m_ref[1] is not None else None
        ambos = [(a, b) for a, b in zip(c, c_ref) if a is not None and b is not None]
        acima = sum(a > b for a, b in ambos) / len(ambos) if ambos else None
        res[j] = [simples, pond, dif, vol, (pond / vol) if pond is not None and vol else None, acima, None]
        desde = f" (só desde {val[0][0]:%m/%Y})" if val and val[0][0] > grelha[0] + pd.Timedelta(days=40) else ""
        rot[j] = (f"{j + 1}. {n_}{desde}: simples {simples:+.1%} · ponderada {pond:+.1%}"
                  if simples is not None and pond is not None else f"{j + 1}. {n_}{desde}").replace("+-", "−")
    ordem = sorted([j for j in range(CRT_N) if res[j][2] is not None or res[j][1] is not None],
                   key=lambda j: -(res[j][2] if res[j][2] is not None else res[j][1]))
    for k, j in enumerate(ordem, 1):
        res[j][6] = k
    linhas = []
    for i, d in enumerate(grelha):
        linhas.append([d.strftime("%Y-%m-%d")] + [None if curvas[j][i] is None else round(curvas[j][i], 5)
                                                   for j in range(CRT_N)])
    gerado = (f"Calculado em {datetime.now():%d/%m/%Y %H:%M} · {grelha[0]:%d/%m/%Y} a {grelha[-1]:%d/%m/%Y} "
              f"({len(grelha)} pontos, de {passo} em {passo} dias de bolsa) · meia-vida {meia:.2f} anos"
              + (f" · referência: {p['ref']} (ponderada {m_ref[1]:+.1%})" if m_ref[1] is not None else ""))
    dados = {"nomes": nomes_s, "ref_nome": p["ref"] or "Referência", "linhas": linhas, "aux": aux,
             "ref": [None if x is None else round(x, 5) for x in c_ref]}
    _gravar_json(FICH_CRT, {"p": p, "titulos": titulos, "res": res, "rotulos": rot, "gerado": gerado, "dados": dados})
    # escrever na folha
    from datetime import datetime as dt_
    C = lambda c: get_column_letter(c)
    ops = [{"op": "escrever", "folha": FOLHA_CRT, "rng": "F7", "valores": [[gerado]]},
           {"op": "escrever", "folha": FOLHA_CRT, "rng": f"E{CRT_L1}:K{CRT_L1 + CRT_N - 1}", "valores": res}]
    grade = [["Data"] + [nomes_s[j] or " " for j in range(CRT_N)]]
    grade += [[dt_.strptime(l_[0], "%Y-%m-%d")] + l_[1:] for l_ in linhas]
    grade += [[None] * (CRT_N + 1)] * (CRT_PTS + 1 - len(grade))
    ops.append({"op": "escrever", "folha": FOLHA_CRT, "rng": f"{C(CRT_C0)}{CRT_D0 - 1}:{C(CRT_C0 + CRT_N)}{CRT_D0 + CRT_PTS - 1}",
                "valores": grade})
    gr = [[p["ref"] or "Referência"]] + [[x if x is None else round(x, 5)] for x in c_ref]
    gr += [[None]] * (CRT_PTS + 1 - len(gr))
    ops.append({"op": "escrever", "folha": FOLHA_CRT, "rng": f"{C(_crt_col(0, 'r'))}{CRT_D0 - 1}:{C(_crt_col(0, 'r'))}{CRT_D0 + CRT_PTS - 1}",
                "valores": gr})
    blocof = [[None] * (4 * CRT_N) for _ in range(4)]
    for j in range(CRT_N):
        for i, (x_, y_) in enumerate((aux[j].get("f") or [])[:4]):
            blocof[i][2 * j], blocof[i][2 * j + 1] = dt_.strptime(x_, "%Y-%m-%d"), y_
        for i, (x_, y_) in enumerate((aux[j].get("m") or [])[:2]):
            blocof[i][2 * CRT_N + 2 * j], blocof[i][2 * CRT_N + 2 * j + 1] = dt_.strptime(x_, "%Y-%m-%d"), y_
    ops.append({"op": "escrever", "folha": FOLHA_CRT, "rng": f"{C(_crt_col(0, 'fx'))}{CRT_D0}:{C(_crt_col(CRT_N - 1, 'my'))}{CRT_D0 + 3}",
                "valores": blocof})
    for j in range(CRT_N):
        r0 = CRT_LBL + (j // 2) * CRT_H
        ops.append({"op": "escrever", "folha": FOLHA_CRT, "rng": f"{'B' if j % 2 == 0 else 'G'}{r0}",
                    "valores": [[rot[j] or f"{j + 1}. (por escolher)"]]})
    ex.ops(ops, guardar=True)
    for k, j in enumerate(ordem, 1):
        r = res[j]
        print(f"   {k:2d}. {nomes_s[j][:34]:34} simples {r[0]:+6.1%} · ponderada {r[1]:+6.1%}"
              + (f" · vs referência {r[2]:+6.1%}" if r[2] is not None else ""))
    print("Comp Rent Títulos: pronto.")


# ======================================================= Ações-CFAlt (V171): classificação alternativa das ações
FICH_ALT_G = {False: os.path.join(PASTA, BASE + "_cfalt_acoes.json"), True: os.path.join(PASTA, BASE + "_cfalt_etfs.json")}
ALT_V = 9                          # muda -> os parâmetros voltam aos aconselhados uma vez
ALT_METODOS = ("P/E (sem ele: nível; sem ele: tendência)", "Nível do painel (sem ele: tendência)", "Nível do painel",
               "Distância à tendência", "P/E")
ALT_BASES = ("Ponderada mensal", "Curva da Comp Rent (compras periódicas até hoje)")   # V178
ALT_PADRAO = {"data": None, "anos": 10.0, "meia": 3.0, "met": ALT_METODOS[0], "m": 4.0, "n": 3.0, "pe_ref": 20.0,
              "vol_ref": 0.25, "pe_max": 100.0, "pe_min": 8.0, "vol_pond": "x", "hmin": 3.0, "k": None, "ref": None,
              "verde": 0.10, "base": ALT_BASES[0], "passo": 30.0, "tmin": 6.0}       # V190: m = 4, n = 3
# V190: nas folhas Ações / ETFs, 2 tabelas de 9 parâmetros nas linhas 1-9 (rótulo, valor em 2 colunas)
ALT_PAR = (("met", "Medida do preço", "@"), ("base", "Base da nota", "@"), ("ref", "Referência (ponderada − ref.)", "@"),
           ("data", "Data de cálculo (vazio = hoje)", "dd/mm/yyyy"), ("anos", "Extensão da análise (anos)", "0.00"),
           ("meia", "Meia-vida (anos)", "0.00"), ("passo", "Curva: periodicidade (dias)", "0"),
           ("tmin", "Curva: tempo mínimo (meses)", "0"), ("k", "Peso qualidade + saúde (k)", "0.0"),
           ("m", "Raiz do fator preço (m)", "0.0"), ("n", "Raiz do fator volatilidade (n)", "0.0"),
           ("pe_ref", "P/E de referência (fator 1)", "0.0"), ("pe_max", "P/E máximo (sem lucros: este)", "0"),
           ("pe_min", "P/E mínimo (abaixo: este)", "0.0"), ("vol_ref", "Volat. de referência (abaixo: 1)", "0%"),
           ("vol_pond", "Volatilidade ponderada (x)", "@"), ("hmin", "Histórico mínimo (anos)", "0.0"))
ALT_TAB = ((16, 19, 21), (22, 26, 28))   # (rótulo de, valor de, valor até): P-R → S:U e V-Y → Z:AB
G_C_GER = "E6"                     # «Calculado em …» (E6:O7)


def _alt_cel(k):
    """(linha, coluna) do valor do parâmetro k."""
    i = [x[0] for x in ALT_PAR].index(k)
    return 1 + i % 9, ALT_TAB[i // 9][1]


def _alt_aconselhado(k):
    v = ALT_PADRAO.get(k)
    if k in ("data", "ref"):
        return ""
    if k == "met":
        return " (acons.: P/E)"
    if k == "base":
        return " (acons.: mensal)"
    if v in (None, ""):
        return " (vazio = não entra)" if k in ("n", "k") else " (acons.: vazio)"
    if k in ("vol_ref", "verde"):
        return f" (acons.: {v:.0%})"
    if isinstance(v, float):
        return f" (acons.: {v:g})"
    return f" (acons.: {v})"


def g_params(etf):
    """Parâmetros guardados da nota por ganhos (ou os aconselhados)."""
    g = _ler_json(FICH_ALT_G[etf], {})
    return {**ALT_PADRAO, **(g.get("p") or {})} if g.get("v") == ALT_V else dict(ALT_PADRAO)


def g_topo(wb, ws, etf):
    """V190/V192: na folha Ações / ETFs, linhas 6-8 (à esquerda) o estado da nota por ganhos e o ↺ REPOR DETIDOS;
    linhas 1-9 (à direita) as 2 tabelas de parâmetros. Também as cores das colunas da nota por ganhos."""
    from openpyxl.worksheet.datavalidation import DataValidation as DV_
    from openpyxl.formatting.rule import FormulaRule
    g = _ler_json(FICH_ALT_G[etf], {})
    p = g_params(etf)
    amarelo = PatternFill("solid", fgColor="FFF2CC")
    cinza = PatternFill("solid", fgColor="F2F2F2")
    fina, grossa = Side(style="thin", color="A6A6A6"), Side(style="medium", color="595959")
    let = get_column_letter
    ws.merge_cells("A6:D6")
    botao(ws, "A6", "↺ REPOR DETIDOS (cópia)", f"{PROTOCOLO}://repor_detidos", "7030A0", "4B2070", 9)
    ws.merge_cells("E6:O7")
    ws[G_C_GER] = g.get("gerado") or "Nota por ganhos: calcula-se em cada ▶ ATUALIZAR (ações, ETFs ou tudo)."
    ws[G_C_GER].font = Font(size=9, color="7030A0")
    ws[G_C_GER].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A8:O9")
    ws["A8"] = ("NOTA POR GANHOS = rentabilidade anual ponderada pelo tempo (euros, com dividendos) ÷ fator do preço "
                "[÷ fator da volatilidade] [× qualidade + saúde]; a nota 0-5 é proporcional (5 = média das 2 melhores). "
                "Calcula-se em cada ▶ ATUALIZAR. Explicação completa no fim da tabela.")
    ws["A8"].font = Font(italic=True, size=8, color="595959")
    ws["A8"].alignment = Alignment(wrap_text=True, vertical="top")
    for i, (k, rot, fm) in enumerate(ALT_PAR):
        r, cv = _alt_cel(k)
        c0, _, cf_ = ALT_TAB[i // 9]
        ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=cv - 1)
        c = ws.cell(r, c0, rot + _alt_aconselhado(k))
        c.font, c.fill = Font(bold=True, size=9), cinza
        c.alignment = Alignment(horizontal="right", vertical="center", shrink_to_fit=True)
        ws.merge_cells(start_row=r, start_column=cv, end_row=r, end_column=cf_)
        c = ws.cell(r, cv, p.get(k))
        c.fill, c.number_format = amarelo, fm
        c.font = Font(bold=True, size=9)
        c.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    for t_, (c0, cv, cf_) in enumerate(ALT_TAB):        # molduras: linhas finas, contorno grosso
        n_lin = len([1 for i in range(len(ALT_PAR)) if i // 9 == t_])
        for r in range(1, n_lin + 1):
            for c_ in range(c0, cf_ + 1):
                ws.cell(r, c_).border = Border(left=grossa if c_ == c0 else fina, right=grossa if c_ == cf_ else fina,
                                               top=grossa if r == 1 else fina, bottom=grossa if r == n_lin else fina)
    for cel_ in (C_AT, C_ORD_EST, C_PRECOS_EST):           # textos de estado na coluna O: em 2 linhas (à direita
        ws[cel_].alignment = Alignment(wrap_text=True, vertical="center")   # estão os parâmetros)
        ws[cel_].font = Font(italic=True, size=7, color="595959")
    for r_ in range(1, 10):                             # alturas fixas (o texto em 2 linhas não alarga as linhas)
        ws.row_dimensions[r_].height = max(ws.row_dimensions[r_].height or 0, 20) if r_ <= 5 else 15
    nomes_ = _nomes_das_folhas(wb)                      # lista de nomes para a referência (coluna escondida)
    cl_n = len(COLS) + 12
    for i_, (n_, t_) in enumerate(nomes_[:600], 1):
        ws.cell(i_, cl_n, n_)
    ws.column_dimensions[let(cl_n)].hidden = True
    rr_, cr_ = _alt_cel("ref")
    if not ws.cell(rr_, cr_).value:
        ws.cell(rr_, cr_).value = next((n_ for n_, t_ in nomes_ if t_ == "SXR8.DE"), None) or \
            next((n_ for n_, t_ in nomes_ if "S&P 500" in n_), None)
    for k_, form_ in (("ref", f"=${let(cl_n)}$1:${let(cl_n)}${max(1, len(nomes_))}"),
                      ("base", '"' + ",".join(ALT_BASES) + '"'), ("met", '"' + ",".join(ALT_METODOS) + '"')):
        dv_ = DV_(type="list", formula1=form_, allow_blank=k_ == "ref")
        ws.add_data_validation(dv_)
        r_, c_ = _alt_cel(k_)
        dv_.add(f"{let(c_)}{r_}")
    ult_ = L0
    while txt(ws.cell(ult_ + 1, COL["ticker"]).value):
        ult_ += 1
    ws.cell(ult_ + 4, 1, (
        "NOTA POR GANHOS (colunas S-T e V-AF; a nota 0-5 é proporcional: 5 = média das 2 melhores, nunca abaixo de 0 nem "
        "acima de 5; cores como as da nota global) = rentabilidade anual ponderada pelo tempo ÷ fator do preço [÷ fator da "
        "volatilidade] [× peso da qualidade e saúde]. Rentabilidade (euros, com dividendos), à escolha: PONDERADA MENSAL "
        "(rentabilidades mensais da extensão escolhida, cada mês com metade do peso do mês «meia-vida» anos mais recente; a "
        "mesma da Comp Rent Títulos) ou CURVA DA COMP RENT (o que se ganhava por ano comprando de X em X dias de bolsa e "
        "guardando até hoje, sem as compras dos últimos meses pedidos: soma dos ganhos ÷ soma dos anos guardados, mais peso às "
        "compras recentes). PREÇO: P/E (fator = (P/E ÷ P/E de referência)^(1/m); sem lucros ou acima do máximo conta o máximo; "
        "abaixo do mínimo conta o mínimo; ETFs sem P/E: o nível de preço, (3 ÷ nível)^(1/m), ou a distância à tendência). "
        "VOLATILIDADE (com n): (vol ÷ referência)^(1/n), nunca menos de 1 (só penaliza). QUALIDADE + SAÚDE (com k): × (média ÷ "
        "3)^(1/k); nos ETFs, a nota global. Rentabilidade negativa: os fatores multiplicam. Sem nota (com o motivo nas "
        "observações): histórico curto, cotações suspeitas ou que não batem com a valorização desta folha. DATA DE CÁLCULO "
        "passada: só com as cotações até essa data e as colunas «real desde a data» mostram o que aconteceu. Calcula-se em "
        "cada ▶ ATUALIZAR. TESTE (180 ações e 36 ETFs, 2008-2025): sinal fraco mas real a 3 anos; volatilidade "
        "e tendência não acrescentaram. É baseada no passado: não garante o futuro."))
    ws.cell(ult_ + 4, 1).font = Font(italic=True, size=9, color="595959")
    c5 = let(COL["g_n5"])
    for lim, cor in ((GERAL_LIM[0], "63BE7B"), (GERAL_LIM[1], "A9D18E"), (GERAL_LIM[2], "E2EFDA"), (2.25, "F4B183"),
                     (1.5, "FFEB9C"), (-99, "F8696B")):
        ws.conditional_formatting.add(
            f"{c5}{L0 + 1}:{let(COL['g_nota'])}{L0 + MAX_LINHAS}",
            FormulaRule(formula=[f"AND(ISNUMBER(${c5}{L0 + 1}),${c5}{L0 + 1}>={lim})"],
                        fill=PatternFill("solid", bgColor=cor), font=Font(bold=True), stopIfTrue=True))


def g_juntar(dados, etf):
    """Os valores da nota por ganhos guardados (último cálculo) entram nos títulos, pelo ticker."""
    g = _ler_json(FICH_ALT_G[etf], {})
    cf = g.get("cf") if g.get("v") == ALT_V and isinstance(g.get("cf"), dict) else {}
    for d in dados:
        v = cf.get(txt(d.get("ticker")).upper())
        for k in G_KEYS:
            d[k] = (v or {}).get(k)


def _alt_medidas(s, inicio, meia):
    """(simples, ponderada com meia-vida, volatilidade, data de início, volatilidade ponderada pelo tempo) de uma série
    de cotações (euros, com dividendos)."""
    import numpy as np
    import pandas as pd
    s = s[s.index >= inicio].dropna()
    if len(s) < 60:
        return None, None, None, None, None
    anos_ = (s.index[-1] - s.index[0]).days / 365.25
    simples = float((s.iloc[-1] / s.iloc[0]) ** (1 / anos_) - 1) if anos_ > 0.5 else None
    m = s.resample("ME").last()
    m = pd.concat([pd.Series([s.iloc[0]], index=[s.index[0]]), m]).sort_index()
    m = m[~m.index.duplicated(keep="last")]
    r = np.log(m / m.shift(1)).dropna()
    if not len(r):
        return simples, None, None, s.index[0], None
    idade = np.array([(s.index[-1] - d).days / 365.25 for d in r.index])
    w = 0.5 ** (idade / meia)
    pond = float(np.exp(12 * float((w * r.values).sum() / w.sum())) - 1)
    rd = np.log(s / s.shift(1)).dropna()
    vol = float(rd.std() * np.sqrt(252))
    wd = 0.5 ** (np.array([(s.index[-1] - d).days / 365.25 for d in rd.index]) / meia)   # V171: mais peso ao recente
    md = float((wd * rd.values).sum() / wd.sum())
    vol_p = float(np.sqrt((wd * (rd.values - md) ** 2).sum() / wd.sum() * 252))
    return simples, pond, vol, s.index[0], vol_p


def _alt_curva(s, inicio, passo, tmin, meia):
    """V178: rentabilidade anual de comprar de «passo» em «passo» dias de bolsa e guardar até ao fim da série (como a
    curva da Comp Rent Títulos), sem as compras dos últimos «tmin» meses; média com mais peso às compras recentes
    (meia-vida). None se não houver pontos."""
    import math
    import pandas as pd
    s = s[s.index >= inicio].dropna()
    if len(s) < 30:
        return None
    fim_, pf = s.index[-1], float(s.iloc[-1])
    corte = fim_ - pd.DateOffset(months=int(round(tmin)))
    soma = pesos = 0.0
    for d_, v_ in list(s.items())[::max(1, int(passo))]:
        if d_ > corte:
            break
        anos_ = (fim_ - d_).days / 365.25
        if anos_ <= 0 or v_ <= 0:
            continue
        w = 0.5 ** (anos_ / meia)               # V182: ganho total ÷ tempo total (com pesos), e não a média das
        soma += w * math.log(pf / float(v_))    # rentabilidades por ano de cada compra: uma compra de há 7 meses que
        pesos += w * anos_                      # subiu 60% já não vale «+120%/ano» com peso máximo
    return math.exp(soma / pesos) - 1 if pesos else None


def g_ler_params(vals):
    """Parâmetros lidos das 2 tabelas (lista pela ordem de ALT_PAR)."""
    p = {}
    for (k, _, _), v0 in zip(ALT_PAR, vals):
        if isinstance(v0, str) and v0.startswith("="):
            v0 = None
        if k == "data":
            p[k] = v0
        elif k == "ref":
            p[k] = txt(v0).strip() or None
        elif k in ("k", "n"):
            v_ = _num_ou(v0)
            p[k] = v_ if v_ is not None and v_ > 0 else None
        elif k == "base":
            p[k] = ALT_BASES[1] if txt(v0).strip().lower().startswith("curva") else ALT_BASES[0]
        elif k == "met":
            v_ = txt(v0).strip()
            p[k] = next((m_ for m_ in ALT_METODOS if m_.lower() == v_.lower()), ALT_METODOS[0])
        elif k == "vol_pond":
            p[k] = "x" if txt(v0).strip().lower() == "x" else ""
        else:
            v_ = _num_ou(v0)
            p[k] = v_ if v_ is not None and v_ > 0 else ALT_PADRAO[k]
    if p["vol_ref"] > 1.5:
        p["vol_ref"] /= 100
    return p


def g_ops_params(idx):
    """Leituras das 2 tabelas de parâmetros (9 linhas cada) da folha idx."""
    return [{"op": "ler", "folha": idx, "rng": f"{get_column_letter(cv)}1:{get_column_letter(cv)}9"} for _, cv, _ in ALT_TAB]


def g_vals_params(blocos):
    return [(lin[0] if isinstance(lin, (list, tuple)) else lin) for bl in blocos for lin in bl]


def g_calcular(acoes, todos_nm, p, etf_):
    """Nota por ganhos de cada título (acoes: dicionários com as colunas do painel). Devolve ({ticker: {g_…}}, texto)."""
    import pandas as pd
    from datetime import date as date_, timedelta
    nome_f = FOLHA_E2 if etf_ else FOLHA_A2
    p = dict(p, grupo="ETFs" if etf_ else "Ações")
    yf = _yf()
    hist, moedas = {}, {}
    d_ref = todos_nm.get((p.get("ref") or "").strip().lower())
    for d in acoes + ([d_ref] if d_ref is not None else []):   # cotações da bolsa principal (mais fiáveis)
        t = txt(d.get("ticker")).upper()
        moedas[t] = txt(d.get("moeda")) or "EUR"
        m_ = re.search(r"principal: .* (\S+)$", txt(d.get("bolsa")))
        hist[t] = t
        if m_ and m_.group(1).upper() != t:
            try:
                mo = yf.Ticker(m_.group(1).upper()).fast_info.get("currency")
                if mo:
                    hist[t], moedas[m_.group(1).upper()] = m_.group(1).upper(), mo
            except Exception:
                pass
    hoje = date_.today()
    dref = None                                        # V173: data de cálculo (análise no passado)
    v_ = p.get("data")
    try:
        if isinstance(v_, (int, float)) and v_ > 20000:
            dref = pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v_))
        elif hasattr(v_, "year"):
            dref = pd.Timestamp(v_.year, v_.month, v_.day)
        elif txt(v_).strip():
            dref = pd.to_datetime(txt(v_).strip(), dayfirst=True)
    except Exception:
        dref = None
    if dref is not None and dref.date() >= hoje - timedelta(days=3):
        dref = None
    base = dref.date() if dref is not None else hoje
    ini = base - timedelta(days=int(p["anos"] * 365.25) + 10)
    print(f"{nome_f}: cotações de {len(acoes)} {p['grupo']} ({p['anos']:.2f} anos"
          + (f", análise em {dref:%d/%m/%Y}" if dref is not None else "") + ") ...")
    eur, _ = _sim_cotacoes_eur(sorted(set(hist.values()) | set(hist)), moedas, ini, hoje)
    idx_ = eur.dropna(how="all").index
    fim = idx_[idx_ <= dref][-1] if dref is not None and (idx_ <= dref).any() else idx_[-1]
    inicio = pd.Timestamp(fim) - pd.Timedelta(days=int(p["anos"] * 365.25))
    pond_ref = None
    if d_ref is not None:
        t_r = txt(d_ref.get("ticker")).upper()
        for c_ in (hist.get(t_r, t_r), t_r):
            if c_ in eur.columns and eur[c_].dropna().size > 60:
                pond_ref = _alt_medidas(eur[c_][eur[c_].index <= fim], inicio, p["meia"])[1]
                break
    curvas_ = {}
    if pond_ref is not None and p.get("base") == ALT_BASES[1]:   # V178: referência na mesma base
        t_r = txt(d_ref.get("ticker")).upper()
        for c_ in (hist.get(t_r, t_r), t_r):
            if c_ in eur.columns and eur[c_].dropna().size > 60:
                pond_ref = _alt_curva(eur[c_][eur[c_].index <= fim], inicio, p["passo"], p["tmin"], p["meia"])
                break
    nt_painel = {}
    for d in acoes:
        v_ = d.get("nota_geral") if isinstance(d.get("nota_geral"), (int, float)) else d.get("nota")
        nt_painel[txt(d.get("ticker")).upper()] = v_ if isinstance(v_, (int, float)) else None
    pos_painel = {t: k for k, t in enumerate(sorted([t for t in nt_painel if nt_painel[t] is not None],
                                                    key=lambda t: -nt_painel[t]), 1)}
    linhas = []
    for d in acoes:
        t = txt(d.get("ticker")).upper()
        s = None
        for c_ in (hist[t], t):
            if c_ in eur.columns and eur[c_].dropna().size > 60:
                s = eur[c_]
                break
        obs = []
        if s is None:
            linhas.append([None, txt(d.get("nome")) or t, t, None, None, None, None, None, None, None,
                           txt(d.get("decisao")), nt_painel.get(t), pos_painel.get(t), None, None, None, None,
                           "sem cotações"])
            continue
        s_tudo = s.dropna()
        s = s[s.index <= fim]                          # V173: só o que se sabia na data de cálculo
        real_t = real_a = None
        if dref is not None and len(s.dropna()) and len(s_tudo):
            p0, p1 = float(s.dropna().iloc[-1]), float(s_tudo.iloc[-1])
            real_t = p1 / p0 - 1
            anos_r = (s_tudo.index[-1] - s.dropna().index[-1]).days / 365.25
            real_a = (1 + real_t) ** (1 / anos_r) - 1 if anos_r >= 0.25 else None
        s_ = s[s.index >= inicio].dropna()                 # V172: cotações suspeitas e histórico curto ficam sem nota
        r_ = s_ / s_.shift(1)
        salto_ = ((r_ >= 1 + EVO_SALTO) | (r_ <= 1 / (1 + EVO_SALTO))).to_numpy()
        v_s = s_.to_numpy(dtype=float)                 # V177: só conta o salto que volta atrás em poucos dias (erro de
        saltos = 0                                     # cotação); uma queda verdadeira (ex.: Soitec −30%) não conta
        for i_ in [i for i in range(1, len(v_s)) if salto_[i]]:
            if any(abs(v_s[j] / v_s[i_ - 1] - 1) <= 0.15 for j in range(i_ + 1, min(len(v_s), i_ + 6))):
                saltos += 1
        if int(salto_.sum()) >= 8:
            saltos = max(saltos, 2)
        igual = (s_.diff() == 0).astype(int)
        parado = int(igual.groupby((igual == 0).cumsum()).sum().max() or 0) if len(igual) else 0
        anos_h = (s_.index[-1] - s_.index[0]).days / 365.25 if len(s_) > 1 else 0
        motivo = (f"cotações suspeitas ({saltos} saltos de ±{EVO_SALTO:.0%} num dia)" if saltos >= 2 else
                  f"cotações suspeitas ({parado} dias seguidos paradas)" if parado >= EVO_PARADO else
                  f"histórico curto ({anos_h:.1f} anos)" if anos_h < min(p["hmin"], p["anos"] - 0.05) else None)
        if motivo:
            linhas.append([None, txt(d.get("nome")) or t, t, None, None, None, None, None, None, None,
                           txt(d.get("decisao")), nt_painel.get(t), pos_painel.get(t), None, real_t, real_a, None,
                           motivo + ": sem nota"])
            continue
        simples, pond, vol, d0, vol_p = _alt_medidas(s, inicio, p["meia"])
        k_v = "val10" if anos_h >= 9.5 else "val5" if anos_h >= 4.5 else "val3"   # V184: confere com a folha base
        vb = d.get(k_v)
        if dref is None and isinstance(vb, (int, float)) and isinstance(simples, (int, float)):
            vb = vb / 100 if abs(vb) > 3 else vb
            if abs(simples - vb) > max(0.15, 0.5 * abs(vb)):
                linhas.append([None, txt(d.get("nome")) or t, t, None, None, simples, None, None, None, None,
                               txt(d.get("decisao")), nt_painel.get(t), pos_painel.get(t), None, real_t, real_a, None,
                               f"cotações de {hist.get(t, t)} não batem com a folha base ({simples:+.0%}/ano aqui, "
                               f"{vb:+.0%}/ano lá): sem nota"])
                continue
        curva_ = _alt_curva(s, inicio, p["passo"], p["tmin"], p["meia"])   # V178
        curvas_[t] = curva_
        pond_m = pond
        if p.get("base") == ALT_BASES[1]:
            pond = curva_
        if p.get("vol_pond") == "x" and vol_p:
            vol = vol_p
        if d0 is not None and d0 > inicio + pd.Timedelta(days=60):
            obs.append(f"só desde {d0:%m/%Y}")
        nv = d.get("nivel_preco") if isinstance(d.get("nivel_preco"), (int, float)) and d.get("nivel_preco") > 0 else None
        tend = None                                    # preço de hoje ÷ reta de crescimento (log) do período
        try:
            import numpy as np
            x_ = np.array([(dd - s_.index[0]).days for dd in s_.index], dtype=float)
            a_, b_ = np.polyfit(x_, np.log(s_.values.astype(float)), 1)
            tend = float(np.exp(np.log(float(s_.iloc[-1])) - (a_ * x_[-1] + b_)))
        except Exception:
            tend = None
        met = p.get("met") or ALT_METODOS[0]
        pe0 = d.get("pe") if isinstance(d.get("pe"), (int, float)) and d.get("pe") > 0 else None
        usar = ("nivel" if met == "Nível do painel" else "tend" if met == "Distância à tendência" else
                "pe" if met == "P/E" else
                ("pe" if pe0 is not None else "nivel" if nv is not None else "tend") if met == ALT_METODOS[0] else
                ("nivel" if nv is not None else "tend"))
        if dref is not None and usar != "tend":        # nível / P/E são os de hoje: numa data passada vai pela tendência
            usar = "tend"
            obs.append("data passada: preço pela tendência")
        pe = d.get("pe") if isinstance(d.get("pe"), (int, float)) else None
        if usar == "nivel" and nv is None:
            pe_u, f_pe = None, 1.0
            obs.append("sem nível de preço: fator 1")
        elif usar == "nivel":
            pe_u, f_pe = nv, (3.0 / nv) ** (1 / p["m"])
            obs.append(f"nível {nv:.1f}")
        elif usar == "tend" and tend is None:
            pe_u, f_pe = None, 1.0
            obs.append("sem tendência: fator 1")
        elif usar == "tend":
            pe_u, f_pe = tend, tend ** (1 / p["m"])
            obs.append(f"preço {tend:.2f}× a tendência")
        elif (pe is None or pe <= 0) and etf_:        # ETF sem P/E: o fator do preço não conta
            pe_u = p["pe_ref"]
            obs.append("sem P/E: fator 1")
        elif pe is None or pe <= 0:
            pe_u = p["pe_max"]
            obs.append("sem lucros: P/E máximo")
        else:
            pe_u = min(pe, p["pe_max"])
            if pe > p["pe_max"]:
                obs.append(f"P/E {pe:.0f}: conta {p['pe_max']:.0f}")
            pe_min_ = min(p.get("pe_min") or 0, p["pe_ref"])     # V180: P/E muito baixo (lucro extraordinário, dados
            if pe_min_ and pe_u < pe_min_:                         # errados) não pode multiplicar a nota sem limite
                obs.append(f"P/E {pe:.1f}: conta {pe_min_:g}")
                pe_u = pe_min_
        if usar == "pe":
            f_pe = (pe_u / p["pe_ref"]) ** (1 / p["m"])
            if pe0 is not None and pe0 <= p["pe_max"]:
                obs.append(f"P/E {pe0:.1f}")
        f_vol = 1.0 if not p.get("n") else max(1.0, (vol / p["vol_ref"]) ** (1 / p["n"])) if vol else None
        # V183: o fator volatilidade só penaliza (nunca < 1): menos oscilação do que a referência não dá prémio
        nota = None
        if pond is not None and f_vol:
            f = f_pe * f_vol
            q_, s_n = d.get("qual"), d.get("saude_n")   # V177: qualidade e saúde (não repetem preço nem tendência)
            if isinstance(q_, (int, float)) and isinstance(s_n, (int, float)) and q_ > 0 and s_n > 0:
                ng_, qual_txt = (q_ + s_n) / 2, "qualidade e saúde"
            else:
                ng_, qual_txt = nt_painel.get(t), "nota do painel"
            if p.get("k") and isinstance(ng_, (int, float)) and ng_ > 0 and dref is None:
                f = f / (ng_ / 3.0) ** (1 / p["k"])
                obs.append(f"{qual_txt} {ng_:.2f}")
            nota = pond / f if pond >= 0 else pond * f
        linhas.append([None, txt(d.get("nome")) or t, t, nota, pond_m, simples, pe_u, vol, f_pe, f_vol,
                       txt(d.get("decisao")), nt_painel.get(t), pos_painel.get(t), None, real_t, real_a,
                       (pond - pond_ref) if pond is not None and pond_ref is not None else None,
                       " · ".join(obs) or None])
    linhas.sort(key=lambda l: (l[3] is None, -(l[3] or 0)))
    for k, l in enumerate(linhas, 1):
        if l[3] is not None:
            l[0] = k
    rnd = lambda v: round(v, 5) if isinstance(v, float) else v
    gerado = (f"Nota por ganhos calculada em {datetime.now():%d/%m/%Y %H:%M} · {inicio:%d/%m/%Y} a {fim:%d/%m/%Y} · "
              f"{len(acoes)} {p['grupo']} · base: {p.get('base')}"
              + (f" (de {p['passo']:g} em {p['passo']:g} dias, sem os últimos {p['tmin']:g} meses)"
                 if p.get('base') == ALT_BASES[1] else "")
              + f" · preço: {p.get('met')} · m = {p['m']:g}, n = {(format(p['n'], 'g') if p.get('n') else 'sem volatilidade')}"
              + (f", k = {p['k']:g}" if p.get("k") else "") + f" · meia-vida {p['meia']:g} anos")
    par_ = [(l[0], l[12]) for l in linhas if l[0] is not None and l[12] is not None]
    if len(par_) >= 6:
        dpp = pd.DataFrame(par_, columns=["a", "p"])
        gerado += f" · parecença com a ordem da nota global: {float(dpp['a'].rank().corr(dpp['p'].rank())):+.2f} (1 = igual)"
    if pond_ref is not None:
        gerado += f" · referência {p.get('ref')}: {pond_ref:+.1%}"
    if dref is not None:
        com = [l for l in linhas if l[3] is not None and l[14] is not None]
        if len(com) >= 6:
            dfc = pd.DataFrame([(l[3], l[14]) for l in com], columns=["n", "r"])
            rho = float(dfc["n"].rank().corr(dfc["r"].rank()))
            k3 = max(2, len(com) // 3)
            cima, baixo = float(dfc["r"].iloc[:k3].median()), float(dfc["r"].iloc[-k3:].median())
            gerado += (f" · ANÁLISE EM {fim:%d/%m/%Y}: desde então, o terço de cima variou {cima:+.0%} e o de baixo "
                       f"{baixo:+.0%} (medianas); correlação nota × real {rho:+.2f}")
    res = {}
    melhores = sorted((l[3] for l in linhas if isinstance(l[3], float) and l[3] > 0), reverse=True)[:2]
    ref5 = sum(melhores) / len(melhores) if melhores else None          # V192: 5 = média das 2 melhores
    n5_ = lambda v: round(max(0.0, min(5.0, 5 * v / ref5)), 2) if isinstance(v, float) and ref5 else None
    if ref5:
        gerado += f" · nota 5 = {ref5:+.1%}/ano (média das 2 melhores)"
    for l in linhas:
        res[l[2]] = {"g_n5": n5_(l[3]), "g_nota": rnd(l[3]), "g_pond": rnd(l[4]), "g_simples": rnd(l[5]),
                     "g_curva": rnd(curvas_.get(l[2])), "g_medida": rnd(l[6]), "g_vol": rnd(l[7]),
                     "g_fpreco": rnd(l[8]), "g_fvol": rnd(l[9]), "g_ref": rnd(l[16]), "g_real_t": rnd(l[14]),
                     "g_real_a": rnd(l[15]), "g_obs": l[17]}
    p_g = {k: v for k, v in p.items() if k != "grupo"}
    p_g["data"] = dref.strftime("%Y-%m-%d") if dref is not None else None
    _gravar_json(FICH_ALT_G[etf_], {"v": ALT_V, "p": p_g, "cf": res, "gerado": gerado})
    for l in linhas[:10]:
        if l[3] is not None:
            print(f"   {l[0]:3d}. {str(l[1])[:30]:30} nota por ganhos {l[3]:+6.1%} · fator preço {l[8]:4.2f}")
    return res, gerado


def _g_todos_nm(tabs):
    todos_nm = {}
    for i_, tb_ in tabs:
        for d_ in linhas_por_cabecalho(tb_, i_):
            todos_nm.setdefault(txt(d_.get("nome")).strip().lower(), d_)
    return todos_nm


def calcular_alt(etf_=False):
    """▶ CALCULAR NOTA POR GANHOS (folha Ações ou ETFs): calcula e escreve as colunas da nota por ganhos em cada linha
    (pelo ticker, sem mexer nas outras colunas nem na ordem); se a folha estiver ordenada pela nota por ganhos,
    reordena-a a seguir."""
    ex = Excel(FICHEIRO)
    idx = 2 if etf_ else 1
    ultima = get_column_letter(len(COLS))
    rr = ex.ops([{"op": "ler", "folha": 1, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"},
                 {"op": "ler", "folha": 2, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"},
                 {"op": "ler", "folha": idx, "rng": C_ORD}] + g_ops_params(idx))
    if not cab_ok([txt(v) for v in rr[idx - 1][0]]):
        raise SystemExit("As colunas desta folha não são as desta versão: carrega primeiro em ▶ ATUALIZAR.")
    p = g_ler_params(g_vals_params(rr[3:]))
    acoes, vistos = [], set()
    for d in linhas_por_cabecalho(rr[idx - 1], idx):
        t = txt(d.get("ticker")).upper()
        if t and t not in vistos:
            vistos.add(t)
            acoes.append(d)
    if not acoes:
        raise SystemExit("A folha não tem títulos.")
    res, gerado = g_calcular(acoes, _g_todos_nm(((1, rr[0]), (2, rr[1]))), p, etf_)
    # escreve só as colunas g_* (a nota; e o bloco a seguir à decisão), linha a linha pela ordem da folha
    jt = COL["ticker"] - 1
    tks = [txt(v[jt]).upper() for v in rr[idx - 1][1:]]
    n = 0
    while n < len(tks) and tks[n]:
        n += 1
    c1, c1b, c2a, c2b = COL["g_n5"], COL["g_nota"], COL[G_COLS2[0][0]], COL[G_COLS2[-1][0]]
    v1 = [[(res.get(t) or {}).get("g_n5"), (res.get(t) or {}).get("g_nota")] for t in tks[:n]]
    v2 = [[(res.get(t) or {}).get(k) for k, *_ in G_COLS2] for t in tks[:n]]
    L = get_column_letter
    com_filtros(ex, [{"op": "escrever", "folha": idx, "rng": f"{L(c1)}{L0 + 1}:{L(c1b)}{L0 + n}", "valores": v1},
                     {"op": "escrever", "folha": idx, "rng": f"{L(c2a)}{L0 + 1}:{L(c2b)}{L0 + n}", "valores": v2},
                     {"op": "escrever", "folha": idx, "rng": G_C_GER, "valores": [[gerado]]}], [idx])
    print(f"{FOLHA_E2 if etf_ else FOLHA_A2}: nota por ganhos escrita em {n} linhas.")
    if txt(rr[2][0][0]).startswith("Nota por ganhos"):
        print("A folha está ordenada pela nota por ganhos: a reordenar...")
        ordenar(idx)


def g_no_atualizar(linhas, etfs, estado, so=None):
    """▶ ATUALIZAR: calcula a nota por ganhos dos grupos atualizados (com os parâmetros lidos da folha) e junta-a às
    linhas; nos outros fica a do último cálculo."""
    todos_nm = {}
    for d_ in list(linhas) + list(etfs):
        todos_nm.setdefault(txt(d_.get("nome")).strip().lower(), d_)
    for etf_, dados in ((False, linhas), (True, etfs)):
        g_juntar(dados, etf_)
        if so is not None and so != ("etfs" if etf_ else "acoes"):
            continue
        vals = (estado or {}).get("gp_e" if etf_ else "gp")
        p = g_ler_params(vals) if vals else g_params(etf_)
        try:
            print(f"\n{FOLHA_E2 if etf_ else FOLHA_A2}: a calcular a nota por ganhos...")
            res, gerado = g_calcular([d for d in dados if txt(d.get("ticker"))], todos_nm, p, etf_)
            for d in dados:
                v = res.get(txt(d.get("ticker")).upper()) or {}
                for k in G_KEYS:
                    d[k] = v.get(k)
        except BaseException as e:
            print(f"   aviso: nota por ganhos não calculada ({e}); fica a do último cálculo.")


FICH_SINC_ANTES = os.path.join(PASTA, BASE + "_antes_sincronizar.json")


def _ui_escolher(titulo, itens, texto):
    """Escolha de um item: janela (sem consola) ou pergunta na consola. Índice (0...) ou None."""
    if sys.stdout is not None and sys.stdin is not None:
        print(texto)
        for i_, t_ in enumerate(itens, 1):
            print(f"  {i_:2d}. {t_}")
        e_ = input("Número (Enter = sair): ").strip()
        return int(e_) - 1 if e_.isdigit() and 1 <= int(e_) <= len(itens) else None
    import tkinter as tk
    from tkinter import ttk
    r_ = {"i": None}
    w = tk.Tk()
    w.title(titulo)
    w.geometry("720x460")
    ttk.Label(w, text=texto, wraplength=690, padding=10, font=("Segoe UI", 10)).pack(fill="x")
    lb = tk.Listbox(w, font=("Consolas", 10), activestyle="dotbox")
    for t_ in itens:
        lb.insert("end", t_)
    lb.pack(fill="both", expand=True, padx=10)
    fb = ttk.Frame(w, padding=10)
    fb.pack(fill="x")

    def ok_():
        sel_ = lb.curselection()
        if sel_:
            r_["i"] = sel_[0]
            w.destroy()
    ttk.Button(fb, text="Cancelar", command=w.destroy).pack(side="right")
    ttk.Button(fb, text="Usar esta cópia", command=ok_).pack(side="right", padx=8)
    lb.bind("<Double-Button-1>", lambda e: ok_())
    w.mainloop()
    return r_["i"]


def _ui_confirmar(titulo, texto, pergunta):
    if sys.stdout is not None and sys.stdin is not None:
        print(texto)
        return input(pergunta + " (s/n): ").strip().lower() in ("s", "sim", "y")
    import tkinter as tk
    from tkinter import ttk
    r_ = {"ok": False}
    w = tk.Tk()
    w.title(titulo)
    w.geometry("980x560")
    t_ = tk.Text(w, font=("Consolas", 9), wrap="none")
    t_.insert("1.0", texto)
    t_.configure(state="disabled")
    t_.pack(fill="both", expand=True, padx=10, pady=10)
    fb = ttk.Frame(w, padding=10)
    fb.pack(fill="x")
    ttk.Label(fb, text=pergunta, font=("Segoe UI", 10, "bold")).pack(side="left")

    def sim_():
        r_["ok"] = True
        w.destroy()
    ttk.Button(fb, text="Não", command=w.destroy).pack(side="right")
    ttk.Button(fb, text="Sim, repor", command=sim_).pack(side="right", padx=8)
    w.mainloop()
    return r_["ok"]


def _ui_aviso(titulo, texto):
    if sys.stdout is not None:
        print(texto)
        return
    import tkinter as tk
    from tkinter import messagebox
    w = tk.Tk()
    w.withdraw()
    messagebox.showinfo(titulo, texto, parent=w)
    w.destroy()


def repor_detidos():
    """V185: repõe o x, o nº de ações, o preço de compra e a corretora nas folhas Ações / ETFs
    a partir de uma cópia: as definições enviadas para a app (config.enc no GitHub, com histórico) ou as cópias locais
    (_detidos.json / _detidos_anterior.json). Mostra as diferenças e só escreve se confirmares."""
    fontes = []                                       # (descrição, função que devolve o dicionário)
    cfg = _config_github()
    if cfg:
        try:
            cms = _api_github(cfg, "commits?path=config.enc&sha=main&per_page=40")
            for c_ in cms:
                quando = datetime.strptime(c_["commit"]["committer"]["date"], "%Y-%m-%dT%H:%M:%SZ")
                quando += (datetime.now() - datetime.utcnow())
                sha_ = c_["sha"]

                def ler_(sha_=sha_):
                    import base64
                    j_ = _api_github(cfg, f"contents/config.enc?ref={sha_}")
                    return json.loads(_decifrar(base64.b64decode(j_["content"]).decode("utf-8"), cfg["senha"]))
                fontes.append((f"app (GitHub) · {quando:%d/%m/%Y %H:%M}", ler_, quando))
        except Exception as e:
            print("Não consegui ler o histórico do GitHub:", e)
    for idx_ in (1, 2):                                # V188: cópias feitas antes de cada sincronização (V184-V187)
        f_ = FICH_SINC_ANTES.replace(".json", f"_{idx_}.json")
        if os.path.exists(f_):
            def ler_s(f_=f_):
                j_ = _ler_json(f_, {})
                det_, q_, p_, c_ = [], {}, {}, {}
                for t_, (sel_, dt_, qt_, pc_, co_) in (j_.get("linhas") or {}).items():
                    q0 = qt_ if isinstance(qt_, (int, float)) else qtd_detida(dt_)
                    if e_detido(dt_) or q0:
                        det_.append(t_)
                    if isinstance(q0, (int, float)) and q0 > 0:
                        q_[t_] = q0
                    if isinstance(pc_, (int, float)):
                        p_[t_] = pc_
                    if txt(co_).strip():
                        c_[t_] = txt(co_).strip()
                return {"det": det_, "qtd": q_, "pcompra": p_, "corr": c_, "_so": idx_}
            j0 = _ler_json(f_, {})
            q_ = datetime.fromtimestamp(os.path.getmtime(f_))
            fontes.append((f"cópia antes da sincronização ({FOLHA_A2 if idx_ == 1 else FOLHA_E2}) · {j0.get('quando') or f'{q_:%d/%m/%Y %H:%M}'}",
                           ler_s, q_))
    for f_ in (FICH_DET.replace(".json", "_anterior.json"), FICH_DET):
        if os.path.exists(f_):
            q_ = datetime.fromtimestamp(os.path.getmtime(f_))
            fontes.append((f"cópia local {os.path.basename(f_)} · {q_:%d/%m/%Y %H:%M}", lambda f_=f_: _ler_json(f_, {}), q_))
    if not fontes:
        _ui_aviso("Repor detidos", "Não encontrei nenhuma cópia dos detidos.")
        return
    i_esc = _ui_escolher("Painel de ações · Repor detidos", [d_ for d_, _, _ in fontes],
                         "Cópias dos títulos detidos (x, nº de ações, preço de compra, corretora), mais recentes primeiro. "
                         "Escolhe uma de ANTES do problema. A seguir vês as alterações e só se confirmares é que se escreve.")
    if i_esc is None:
        return
    try:
        j = fontes[i_esc][1]() or {}
    except Exception as e:
        _ui_aviso("Repor detidos", f"Não consegui ler essa cópia: {e}")
        return
    det = set(j.get("det") or [])
    qtd, pc, co = j.get("qtd") or {}, j.get("pcompra") or {}, j.get("corr") or {}
    ex = Excel(FICHEIRO)
    ultima = get_column_letter(len(COLS))
    jt, jn, jh = COL["ticker"] - 1, COL["nome"] - 1, COL["corr"] - 1
    lh = get_column_letter(COL["corr"])
    plano, mud = {}, []
    for idx in (1, 2):
        plano[idx] = []
        if j.get("_so") and j["_so"] != idx:
            continue
        rr = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0 + 1}:{ultima}{L0 + MAX_LINHAS}"}])[0]
        ops_ = []
        for i_, v in enumerate(rr):
            t_, n_ = txt(v[jt]).upper(), txt(v[jn])
            m_ = re.search(r',"([^"]*)"\)\s*$', n_) if n_.startswith("=") else None
            n_ = m_.group(1) if m_ else n_
            if not t_:
                continue
            ks = [k for k in (t_, n_) if k]
            g_ = lambda dd: next((dd[k] for k in ks if k in dd), None)
            tem = any(k in det for k in ks) or bool(g_(qtd))
            novo = ["x" if tem else " ", g_(qtd) if tem and g_(qtd) is not None else " ",
                    g_(pc) if g_(pc) is not None else " ", g_(co) or " "]
            atual = [v[1], v[2], v[3], v[jh]]
            norm = lambda x: "" if x is None or txt(x).strip() == "" else (round(float(x), 6) if isinstance(x, (int, float)) else txt(x).strip().lower())
            if [norm(x) for x in novo] != [norm(x) for x in atual]:
                mud.append(f"   {n_[:28]:28} {t_:10} detido/nº/preço/corretora: "
                           f"{[txt(x) for x in atual]} -> {[txt(x) for x in novo]}")
                ops_.append({"op": "escrever", "folha": idx, "rng": f"B{L0 + 1 + i_}:D{L0 + 1 + i_}", "valores": [novo[:3]]})
                ops_.append({"op": "escrever", "folha": idx, "rng": f"{lh}{L0 + 1 + i_}", "valores": [[novo[3]]]})
        plano[idx] = ops_
    if not mud:
        _ui_aviso("Repor detidos", "Essa cópia é igual ao que está nas folhas: nada a mudar.")
        return
    if not _ui_confirmar("Painel de ações · Repor detidos", f"Cópia: {fontes[i_esc][0]}\nAlterações ({len(mud)} títulos):"
                         f"  [atual] -> [reposto]\n\n" + "\n".join(m_.strip() for m_ in mud),
                         "Repor estes valores?"):
        _ui_aviso("Repor detidos", "Nada mudado.")
        return
    for idx in (1, 2):
        if plano[idx]:
            com_filtros(ex, plano[idx], [idx])
    _ui_aviso("Repor detidos", "Títulos detidos repostos. Confere nas folhas; o próximo ▶ ATUALIZAR envia-os também "
                               "para a app.")


def esperar_excel(espera=90):
    """Depois de o painel voltar a abrir: espera que o Excel o tenha aberto (para escrever nele e não no ficheiro)."""
    if os.name != "nt":
        return
    livro, fim = os.path.basename(FICHEIRO), time.time() + espera
    while time.time() < fim:
        time.sleep(3)
        for f in (_ops_pywin32, _ops_powershell):
            try:
                if f(livro, [{"op": "ler", "folha": 1, "rng": "A1"}], False) is not None:
                    return
            except Exception:
                pass


def sem_formula(v):
    """Valor a escrever pelo Excel: as fórmulas (nome com ligação) ficam; o resto como está."""
    return v


_SIM_LOCAL = {}


def _sim_etf_hold(etfs, acoes):
    """V153: {ETF: [(ticker da ação na folha Ações, peso), ...]} pela composição guardada dos ETFs (a de hoje), para o
    Simulador e a Evolução darem aos ETFs, em cada data, a nota ponderada e o nível de preço das suas ações."""
    comp = _ler_json(FICH_COMP, {})
    indice = {}
    for d in acoes:
        for c in _chaves_acao(d) | {_nome_chave(d.get("nome"))} - {None}:
            indice.setdefault(c, d)
    out = {}
    for e in etfs:
        c = comp.get(txt(e.get("ticker"))) or comp.get(txt(e.get("_orig"))) or {}
        h = {}
        for sym, nm, w in c.get("hold", []):
            a = _acao_de(sym, indice, nm)
            ta = txt(a.get("ticker")).upper() if a is not None else ""
            if ta and isinstance(w, (int, float)) and w > 0:
                h[ta] = h.get(ta, 0.0) + float(w)
        if h:
            out[txt(e.get("ticker")).upper()] = list(h.items())
    return out


def _sim_notas_painel(P, tipos, dias, sec=None, niveis=None, decisoes=None, etf_hold=None):
    """Nota geral do PAINEL (as mesmas regras: pontuar + nota_global) em cada dia de «dias», só com as cotações até
    esse dia (sem analistas, valor justo, qualidade e saúde, que não há para datas passadas: contam como o sinal,
    como no painel). P = cotações diárias na moeda da bolsa. -> DataFrame (dias × tickers)."""
    import numpy as np
    import pandas as pd
    P = P.astype(float)
    m100 = P.rolling(CURTA, min_periods=CURTA).mean()
    m50 = P.rolling(50, min_periods=50).mean()
    m200 = P.rolling(200, min_periods=200).mean()
    gap = m100 / m200 - 1
    r = np.log(P).diff()
    sig = (r.rolling(63, min_periods=40).std(ddof=0) + r.rolling(252, min_periods=40).std(ddof=0)) / 2 * np.sqrt(21)
    stop = P.rolling(126, min_periods=1).max() * np.exp(-REDE_SIGMAS * sig)
    ref6 = P.shift(126)
    nval = P.notna().cumsum()
    M = np.log(P.resample("ME").last())

    def inclin(w, minimo, expand=False):
        k = pd.DataFrame({c: np.arange(len(M), dtype=float) for c in M.columns}, index=M.index).where(M.notna())
        rol = (lambda x: x.expanding(min_periods=minimo)) if expand else (lambda x: x.rolling(w, min_periods=minimo))
        mk, my, mky, mkk = rol(k).mean(), rol(M).mean(), rol(k * M).mean(), rol(k * k).mean()
        return np.exp((mky - mk * my) / (mkk - mk * mk) * 12) - 1
    s_all, s5, s10 = inclin(None, 36, True), inclin(61, 61), inclin(121, 121)
    LD = np.log(P)                                   # retas diárias do último ano e dos 6 meses (fator de tendência)
    kd = pd.DataFrame({c: np.arange(len(LD), dtype=float) for c in LD.columns}, index=LD.index).where(LD.notna())

    def inclin_d(w):
        rol = lambda x: x.rolling(w, min_periods=int(w * 0.9))
        mk, my, mky, mkk, myy = rol(kd).mean(), rol(LD).mean(), rol(kd * LD).mean(), rol(kd * kd).mean(), rol(LD * LD).mean()
        cov, vk, vy = mky - mk * my, mkk - mk * mk, myy - my * my
        return np.exp(cov / vk * 252) - 1, (cov * cov / (vk * vy)).clip(0, 1)
    i1, q1 = inclin_d(252)
    i6, q6 = inclin_d(126)
    i60s, q60s = inclin_d(SUB60_D + 1)
    ra_, re_, _, mx1 = _rec_sinais(P)                  # recuperação (V148)
    rda, rde = _dias_desde(ra_), _dias_desde(re_)
    sub60 = (i60s >= SUB60_MIN) & (q60s >= SUB60_R2)
    iB = inclin_d(ARR_BASE)[0].shift(ARR_SUB + 1)            # arranque: base parada e depois subida rápida
    iR = inclin_d(ARR_SUB + 1)[0]
    mxB = P.rolling(ARR_BASE, min_periods=int(ARR_BASE * 0.9)).max().shift(ARR_SUB + 1)
    arrq = (iB <= ARR_BASE_MAX) & (iR >= ARR_SUB_MIN) & (P > mxB)
    v1 = r.rolling(252, min_periods=200).std() * np.sqrt(252)
    dias = [d for d in dias if d in P.index]
    out = pd.DataFrame(index=pd.DatetimeIndex(sorted(set(dias))), columns=P.columns, dtype=float)
    etfs = [t for t in P.columns if tipos.get(t) == "ETF"]
    ordem_cols = [t for t in P.columns if tipos.get(t) != "ETF"] + etfs     # V153: as ações antes dos ETFs (que usam as notas delas)
    for dia in out.index:
        dia_nt, dia_nv = {}, {}
        fm = M.index[M.index.searchsorted(dia, side="right") - 1] if M.index.searchsorted(dia, side="right") else None
        mes_fechado = fm if fm is not None and fm < dia else (M.index[M.index.get_loc(fm) - 1]
                                                               if fm is not None and M.index.get_loc(fm) > 0 else None)
        m6 = {t: P.at[dia, t] / ref6.at[dia, t] - 1 for t in etfs
              if pd.notna(ref6.at[dia, t]) and ref6.at[dia, t] and pd.notna(P.at[dia, t])}
        ordem = sorted(m6, key=lambda t: -m6[t])
        pos6 = {t: (i / (len(ordem) - 1) if len(ordem) > 1 else 0.0) for i, t in enumerate(ordem)}
        for t in ordem_cols:
            pr = P.at[dia, t]
            if pd.isna(pr) or nval.at[dia, t] < 200 or pd.isna(m200.at[dia, t]):
                continue
            d = {"tipo": tipos.get(t) or "Ação", "mm100c": float(m100.at[dia, t]), "mm200c": float(m200.at[dia, t]),
                 "mm50": float(m50.at[dia, t]) if pd.notna(m50.at[dia, t]) else None, "preco": float(P.at[dia, t]),
                 "arranque": 1 if bool(arrq.at[dia, t]) else 0, "sub60": 1 if bool(sub60.at[dia, t]) else 0,
                 "max1a": float(mx1.at[dia, t]) if pd.notna(mx1.at[dia, t]) else None,
                 "inc60": float(i60s.at[dia, t]) if pd.notna(i60s.at[dia, t]) else None}
            for k_, df_ in (("rec_a", rda), ("rec_e", rde)):
                v_ = df_.at[dia, t]
                if pd.notna(v_) and v_ <= REC_MEMO:
                    d[k_] = int(v_)
            pos5 = P.index.get_loc(dia) - 5
            if pos5 >= 0 and pd.notna(gap.iloc[pos5][t]):
                d["gap5"] = float(gap.iloc[pos5][t])
            if pd.notna(stop.at[dia, t]):
                d["venda_min"] = d["stop"] = float(stop.at[dia, t])
            if pd.notna(ref6.at[dia, t]):
                d["ref6m"] = float(ref6.at[dia, t])
            if mes_fechado is not None and pd.notna(s_all.at[mes_fechado, t]):
                sa = float(s_all.at[mes_fechado, t])
                v5 = float(s5.at[mes_fechado, t]) if pd.notna(s5.at[mes_fechado, t]) else None
                v10 = float(s10.at[mes_fechado, t]) if pd.notna(s10.at[mes_fechado, t]) else None
                d["lp"] = min(sa, v5) if v5 is not None else sa
                d["lp10"] = v10 if v10 is not None else sa
                d["lp5"] = v5 if v5 is not None else sa
                d["val_sempre"], d["val10"], d["val5"] = sa, v10, v5
            for k_, df_ in (("inc1a", i1), ("r2_1a", q1), ("inc6m", i6), ("r2_6m", q6), ("vol1a", v1)):
                if pd.notna(df_.at[dia, t]):
                    d[k_] = float(df_.at[dia, t])
            if t in pos6:
                d["_pos6"], d["_n6"] = pos6[t], len(ordem)
            if sec and t in sec:                        # EUA: fundamentos dos relatórios já divulgados nesse dia
                try:
                    f_ = sec[t].em(dia)
                    pu = f_.pop("_preco_usd", None)
                    d.update(f_)
                    if pu:
                        d["preco"] = pu
                except Exception:
                    pass
            if etf_hold and t in etf_hold:              # V153: ETF = as suas ações nesse dia (como na tabela de hoje)
                sw = swn = swc = swcn = 0.0
                for a_, w_ in etf_hold[t]:
                    if a_ in dia_nt:
                        sw, swn = sw + w_, swn + w_ * dia_nt[a_]
                    if a_ in dia_nv:
                        swc, swcn = swc + w_ * dia_nv[a_], swcn + w_
                if sw > 0:
                    d["etf_acoes"] = round(swn / sw, 2)
                if swcn >= 0.15:                        # como na tabela: ≥ 15% do ETF em ações com nível de preço
                    d["nivel_preco"] = round(swc / swcn, 1)
            pontuar(d, float(pr))
            try:
                nt = nota_global(d)
            except Exception:
                nt = None
            d["nota"] = nt                              # V148: a decisão usa a nota (sem ela usava o sinal)
            if nt is not None and tipos.get(t) != "ETF":
                dia_nt[t] = nt
                if isinstance(d.get("nivel_preco"), (int, float)):
                    dia_nv[t] = d["nivel_preco"]
            if nt is not None:
                out.at[dia, t] = nt
                if decisoes is not None:                # folha Histórico: a decisão do painel nesse dia
                    decisoes[(dia, t)] = (decisao_global(d), d.get("nota_geral"), d.get("nivel_preco"))
            if niveis is not None and t in niveis.columns and isinstance(d.get("nivel_preco"), (int, float)):
                niveis.at[dia, t] = d["nivel_preco"]
    return out


def _sim_cotacoes_eur(tickers, moedas, ini, fim):
    """Cotações diárias ajustadas (dividendos reinvestidos) em euros: DataFrame (dias × tickers)."""
    import pandas as pd
    from datetime import timedelta
    yf = _yf()
    pedir = sorted(set(tickers))
    h = _por_ticker(yf.download(pedir, start=(ini - timedelta(days=10)).strftime("%Y-%m-%d"),
                                end=(fim + timedelta(days=1)).strftime("%Y-%m-%d"), interval="1d", group_by="ticker",
                                auto_adjust=True, progress=False, threads=True), pedir)
    precos = {}
    for t in pedir:
        df = h.get(t)
        if df is None or not len(df) or "Close" not in df:
            continue
        s = df["Close"].dropna()
        if getattr(s.index, "tz", None) is not None:
            s.index = s.index.tz_localize(None)
        s.index = s.index.normalize()
        precos[t] = _sim_limpar(s[~s.index.duplicated(keep="last")])
    menor = {"GBp": ("GBP", 0.01), "GBX": ("GBP", 0.01), "ZAc": ("ZAR", 0.01), "ILA": ("ILS", 0.01)}
    precisa = {}
    for t in precos:
        m = moedas.get(t) or "EUR"
        base, fat = menor.get(m, (m.upper() if m else "EUR", 1.0))
        precisa[t] = (base, fat)
    pares = sorted({b for b, _ in precisa.values() if b != "EUR"})
    fx = {}
    if pares:
        hf = _por_ticker(yf.download([f"{b}EUR=X" for b in pares], start=(ini - timedelta(days=15)).strftime("%Y-%m-%d"),
                                     end=(fim + timedelta(days=1)).strftime("%Y-%m-%d"), interval="1d",
                                     group_by="ticker", auto_adjust=True, progress=False, threads=True),
                         [f"{b}EUR=X" for b in pares])
        for b in pares:
            df = hf.get(f"{b}EUR=X")
            if df is not None and len(df) and "Close" in df:
                s = df["Close"].dropna()
                if not isinstance(s.index, pd.DatetimeIndex):   # sem datas: câmbio fixo (o último)
                    fx[b] = float(s.iloc[-1])
                    continue
                if getattr(s.index, "tz", None) is not None:
                    s.index = s.index.tz_localize(None)
                s.index = s.index.normalize()
                fx[b] = s[~s.index.duplicated(keep="last")]
    sem_cambio = []
    tab = pd.DataFrame(precos).sort_index()
    for t in list(tab.columns):
        b, fat = precisa[t]
        if b == "EUR":
            tab[t] = tab[t] * fat
        elif b in fx and not hasattr(fx[b], "index"):
            tab[t] = tab[t] * fat * fx[b]
        elif b in fx:
            tab[t] = tab[t] * fat * fx[b].reindex(tab.index.union(fx[b].index)).ffill().bfill().reindex(tab.index)
        else:
            sem_cambio.append(t)
            tab = tab.drop(columns=[t])
    tab = tab.ffill()
    _SIM_LOCAL["tab"] = pd.DataFrame(precos).sort_index().ffill()       # na moeda de cada bolsa (para a nota do painel)
    return tab.loc[(tab.index >= pd.Timestamp(ini) - pd.Timedelta(days=10)) & (tab.index <= pd.Timestamp(fim))], sem_cambio


def _sim_datas_reforco(ini, fim, per):
    from datetime import timedelta
    import pandas as pd
    passo = {"Semanal": None, "Mensal": 1, "Trimestral": 3, "Semestral": 6, "Anual": 12}.get(per, -1)
    out, d, k = [], ini, 0
    if passo == -1:
        return []
    while d <= fim:
        if d > ini:
            out.append(d)
        k += 1
        d = ini + timedelta(weeks=k) if passo is None else (pd.Timestamp(ini) + pd.DateOffset(months=passo * k)).to_pydatetime()
    return out


def _sim_correr(tab, pesos, compras, fim):
    """Compra nas datas indicadas (dia de bolsa seguinte), dividindo pelos pesos dos títulos com cotação.
    -> (unidades, investido por título, série diária de valor, série diária de investido, fluxos)."""
    import pandas as pd
    idx = tab.index
    unid = {t: 0.0 for t in pesos}
    inv_t = {t: 0.0 for t in pesos}
    caixa, fluxos = 0.0, []
    planos = {}
    for d, v in compras:
        pos = idx.searchsorted(pd.Timestamp(d))
        if pos >= len(idx):
            continue
        planos.setdefault(idx[pos], 0.0)
        planos[idx[pos]] += v
    val, inv, investido = [], [], 0.0
    for dia in idx:
        if dia in planos:
            v = planos[dia] + caixa
            fluxos.append((dia.to_pydatetime(), -planos[dia]))
            investido += planos[dia]
            com = {t: w for t, w in pesos.items() if t in tab.columns and pd.notna(tab.at[dia, t]) and tab.at[dia, t] > 0}
            sw = sum(com.values())
            if sw > 0:
                for t, w in com.items():
                    parte = v * w / sw
                    unid[t] += parte / tab.at[dia, t]
                    inv_t[t] += parte
                caixa = 0.0
            else:
                caixa = v                                # ainda nenhum título com cotação: fica à espera
        tot = caixa + sum(unid[t] * tab.at[dia, t] for t in unid if t in tab.columns and pd.notna(tab.at[dia, t]))
        val.append(tot)
        inv.append(investido)
    serie_v = pd.Series(val, index=idx)
    serie_i = pd.Series(inv, index=idx)
    if len(idx):
        fluxos.append((idx[-1].to_pydatetime(), float(serie_v.iloc[-1])))
    return unid, inv_t, serie_v, serie_i, fluxos


def _sim_maior_queda(serie_v, serie_i):
    """Maior queda do índice da carteira (rentabilidade diária sem contar os reforços)."""
    import numpy as np
    v, i = serie_v.values, serie_i.values
    idx_ = 1.0
    pico, pior = 1.0, 0.0
    for k in range(1, len(v)):
        entrou = i[k] - i[k - 1]
        if v[k - 1] > 0:
            idx_ *= (v[k] - entrou) / v[k - 1]
        pico = max(pico, idx_)
        pior = min(pior, idx_ / pico - 1)
    return pior


def _sim_notas_precos(tab):
    """Nota de preços (1 a 5) de cada título em cada dia, só com as cotações até esse dia:
    40% tendência de longo prazo (reta dos preços mensais: 65% 10 anos + 35% 5 anos; mínimo 3 anos),
    35% sinal das médias (preço vs médias de 100 e 200 dias), 25% momento (subida em 6 meses)."""
    import numpy as np
    import pandas as pd
    P = tab
    ma100 = P.rolling(100, min_periods=100).mean()
    ma200 = P.rolling(200, min_periods=200).mean()
    mom = P / P.shift(126) - 1
    M = np.log(P.resample("ME").last())
    k = pd.Series(np.arange(len(M), dtype=float), index=M.index)

    def inclin(w, minimo):
        kk = pd.DataFrame({c: k for c in M.columns})
        kk = kk.where(M.notna())
        mk = kk.rolling(w, min_periods=minimo).mean()
        my = M.rolling(w, min_periods=minimo).mean()
        mky = (kk * M).rolling(w, min_periods=minimo).mean()
        mkk = (kk * kk).rolling(w, min_periods=minimo).mean()
        b = (mky - mk * my) / (mkk - mk * mk)
        return np.exp(b * 12) - 1
    s10, s5 = inclin(120, 120), inclin(60, 36)
    tend = (0.65 * s10 + 0.35 * s5).where(s10.notna(), s5)
    tend_d = tend.reindex(P.index.union(tend.index)).ffill().reindex(P.index)
    lin = lambda x, a, b: (1 + 4 * (x - a) / (b - a)).clip(1, 5)
    n_tend = lin(tend_d, -0.05, 0.15)
    n_mom = lin(mom, -0.15, 0.25)
    sinal = pd.DataFrame(3.0, index=P.index, columns=P.columns)
    acima200, acima100, cruz = P > ma200, P > ma100, ma100 > ma200
    sinal = sinal.mask(acima200, 4.0).mask(acima200 & acima100 & cruz, 5.0)
    sinal = sinal.mask(~acima200, 2.0).mask(~acima200 & ~acima100 & ~cruz, 1.0)
    sinal = sinal.where(ma200.notna())
    nota = 0.40 * n_tend + 0.35 * sinal + 0.25 * n_mom
    return nota.where(n_tend.notna() & sinal.notna() & n_mom.notna())


def _sim_correr_dinamico(tab, notas, compras, modo, limite, rodar, minimo=None, ordem=None, verificar=None):
    """Em cada compra escolhe pela nota de preços desse dia (N primeiros ou nota ≥); só compras ou rodar a carteira
    (vende o que saiu e reparte tudo em partes iguais pelos escolhidos)."""
    import pandas as pd
    idx = tab.index
    planos = {}
    for d, v in compras:
        pos = idx.searchsorted(pd.Timestamp(d))
        if pos < len(idx):
            planos[idx[pos]] = planos.get(idx[pos], 0.0) + v
    unid, inv_t, vezes, vendido = {}, {}, {}, {}
    caixa, fluxos, investido, val, inv = 0.0, [], 0.0, [], []
    verificar = set(verificar or [])
    sem_escolha = 0
    for dia in idx:
        px = tab.loc[dia]
        # dia de compra; ou, com dinheiro à espera, um dia de verificação (1.º dia de bolsa de cada mês)
        if dia in planos or (caixa > 0.01 and dia in verificar):
            entra = planos.get(dia, 0.0)
            if entra:
                fluxos.append((dia.to_pydatetime(), -entra))
                investido += entra
            nt = notas.loc[dia].dropna()
            nt = nt[[t for t in nt.index if pd.notna(px.get(t)) and px.get(t) > 0]]
            if minimo is not None:
                nt = nt[nt >= minimo]
            if ordem is not None:                       # ordena por outra coisa (nível de preço) entre os que passam
                ov = ordem.loc[dia].reindex(nt.index).dropna()
                nt = ov
            esc = list(nt.sort_values(ascending=False).index[:int(limite)]) if modo == "n" else \
                list(nt[nt >= limite].index)
            for t in esc:
                vezes[t] = vezes.get(t, 0) + 1
            if not esc and entra:
                sem_escolha += 1
            if rodar and esc and entra:
                for t in list(unid):                    # vende tudo (fica em caixa) e reparte pelos escolhidos
                    if unid[t] and pd.notna(px.get(t)):
                        caixa += unid[t] * px[t]
                        vendido[t] = vendido.get(t, 0.0) + unid[t] * px[t]
                    unid[t] = 0.0
            dinheiro = entra + caixa
            if esc:
                for t in esc:
                    parte = dinheiro / len(esc)
                    unid[t] = unid.get(t, 0.0) + parte / px[t]
                    inv_t[t] = inv_t.get(t, 0.0) + parte
                caixa = 0.0
            else:
                caixa = dinheiro
        tot = caixa + sum(u * px[t] for t, u in unid.items() if u and pd.notna(px.get(t)))
        val.append(tot)
        inv.append(investido)
    sv, si = pd.Series(val, index=idx), pd.Series(inv, index=idx)
    if len(idx):
        fluxos.append((idx[-1].to_pydatetime(), float(sv.iloc[-1])))
    _SIM_LOCAL["espera"] = (caixa, sem_escolha)
    return unid, inv_t, sv, si, fluxos, vezes, vendido


def simular():
    """▶ SIMULAR (folha Simulador): calcula à parte e escreve o resultado só nessa folha."""
    from datetime import datetime as dt_
    import pandas as pd
    ex = Excel(FICHEIRO)
    try:
        carregar_pesos(ex)                                # as percentagens da folha Ponderações (como estão agora)
    except SystemExit:
        raise
    except Exception:
        pass
    p = ler_simulador(ex)
    _sim_gravar(param=p)
    ini = _sim_data(p["d_ini"])
    fim = _sim_data(p["d_fim"]) or dt_.now()
    if not ini or ini >= fim:
        print("Simulador: escreve uma data inicial anterior à data final.")
        sys.exit(1)
    ultima = get_column_letter(len(COLS))
    r = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)])
    acoes, etfs = (linhas_por_cabecalho(tab, i) for i, tab in zip((1, 2), r))
    acoes, etfs = list(acoes), list(etfs)
    univ = {"Só ações": acoes, "Só ETFs": etfs}.get(p["entre"], acoes + etfs)
    moedas = {txt(d.get("ticker")).upper(): txt(d.get("moeda")) for d in acoes + etfs if txt(d.get("ticker"))}
    nomes = {txt(d.get("ticker")).upper(): txt(d.get("nome")) for d in acoes + etfs if txt(d.get("ticker"))}
    escolha = p["comprar"]
    cod = _sim_cod(escolha)
    dinamico = cod in (5, 6, 8, 10, 12)
    duplo = cod in (7, 8)
    grupos = None
    nmin = p.get("nota_min") if (dinamico or cod in (9, 11)) else None   # nota mínima: 5, 6, 8, 9, 10, 11, 12
    nota_de = {txt(d.get("ticker")).upper(): _sim_num(d.get("nota")) for d in acoes + etfs if txt(d.get("ticker"))}
    passa = lambda d: nmin is None or (_sim_num(d.get("nota")) is not None and _sim_num(d["nota"]) >= nmin)
    if not dinamico:                                      # nota mínima: só entram os que a cumprem (nota de hoje)
        acoes_c, etfs_c = [d for d in acoes if passa(d)], [d for d in etfs if passa(d)]
    else:
        acoes_c, etfs_c = acoes, etfs
    univ_c = {"Só ações": acoes_c, "Só ETFs": etfs_c}.get(p["entre"], acoes_c + etfs_c)
    if duplo:                                             # duas partes: ações e ETFs, cada uma com o seu valor
        for k in ("n_a", "v_a", "n_e", "v_e"):              # vazios (folha de uma versão anterior): os padrões
            if p.get(k) is None:
                p[k] = SIM_PADRAO[k]
        n_a, n_e = int(p.get("n_a") or 0), int(p.get("n_e") or 0)
        if dinamico:
            ga = [txt(d["ticker"]).upper() for d in acoes if txt(d.get("ticker"))] if n_a else []
            ge = [txt(d["ticker"]).upper() for d in etfs if txt(d.get("ticker"))] if n_e else []
        else:
            top = lambda ls, n: [txt(d["ticker"]).upper() for d in sorted(
                [d for d in ls if isinstance(_sim_num(d.get("nota")), float)], key=lambda d: -_sim_num(d["nota"]))[:n]]
            ga, ge = top(acoes_c, n_a), top(etfs_c, n_e)
        grupos = [(ga, n_a, float(p.get("v_a") or 0)), (ge, n_e, float(p.get("v_e") or 0))]
        grupos = [g_ for g_ in grupos if g_[0] and g_[1] > 0]
        pesos = {t: 1.0 for g_ in grupos for t in g_[0]}
    elif cod in (10, 12):
        pesos = {txt(d["ticker"]).upper(): 1.0 for d in acoes if txt(d.get("ticker"))}
    elif cod == 11:                                       # todos com nível de preço ≥ F6 (e nota ≥ mínimo, acima)
        nv_min = p["n"] if p["n"] is not None else 3.5
        sel = [d for d in univ_c if isinstance(_sim_num(d.get("nivel_preco")), float)
               and _sim_num(d["nivel_preco"]) >= nv_min]
        pesos = {txt(d["ticker"]).upper(): 1.0 for d in sel}
    elif dinamico:                                          # todos os títulos do universo; escolhe-se em cada data
        pesos = {txt(d["ticker"]).upper(): 1.0 for d in univ if txt(d.get("ticker"))}
    elif cod == 9:                                        # melhor preço (nível de preço) entre os de nota ≥ mínimo
        n = int(p["n"] or 10)
        cand = [d for d in univ_c if isinstance(_sim_num(d.get("nivel_preco")), float)]
        sel = sorted(cand, key=lambda d: (-_sim_num(d["nivel_preco"]), -(_sim_num(d.get("justo_pct")) or -9),
                                          -(_sim_num(d.get("nota")) or 0)))[:n]
        pesos = {txt(d["ticker"]).upper(): 1.0 for d in sel}
    elif cod == 1:
        sel = [d for d in univ_c if e_detido(d.get("det")) or (_sim_num(d.get("qtd")) or 0) > 0]
        pesos = {txt(d["ticker"]).upper(): 1.0 for d in sel}
    elif cod == 2:
        n = int(p["n"] or 10)
        sel = sorted([d for d in univ_c if isinstance(_sim_num(d.get("nota")), (int, float))], key=lambda d: -_sim_num(d["nota"]))[:n]
        pesos = {txt(d["ticker"]).upper(): 1.0 for d in sel}
    elif cod == 3:
        m = p["n"] if p["n"] is not None else 4
        sel = [d for d in univ_c if isinstance(_sim_num(d.get("nota")), (int, float)) and _sim_num(d["nota"]) >= m]
        pesos = {txt(d["ticker"]).upper(): 1.0 for d in sel}
    else:
        pesos = {}
        for t, w in p["carteira"]:
            pesos[t] = pesos.get(t, 0) + (w if isinstance(w, (int, float)) and w > 0 else 0)
        if pesos and not any(pesos.values()):
            pesos = {t: 1.0 for t in pesos}
        pesos = {t: w for t, w in pesos.items() if w > 0 and
                 (nmin is None or nota_de.get(t) is None or nota_de[t] >= nmin)}
    if not pesos:
        print(f"Simulador: nenhum título para «{escolha}». Verifica a escolha (e o x/quantidades, a nota, a nota mínima ou a carteira).")
        sys.exit(1)
    bench = txt(p.get("bench")).upper()
    falta_moeda = [t for t in list(pesos) + ([bench] if bench else []) if t not in moedas]
    if falta_moeda:                                       # títulos fora das folhas: moeda pelo Yahoo
        yf = _yf()
        for t in falta_moeda:
            try:
                moedas[t] = yf.Ticker(t).fast_info.get("currency") or ""
            except Exception:
                moedas[t] = ""
            try:
                nomes.setdefault(t, txt(yf.Ticker(t).info.get("shortName")) or t)
            except Exception:
                nomes.setdefault(t, t)
    print(f"Simulador: {len(pesos)} título(s) · {escolha} · de {ini:%d/%m/%Y} a {fim:%d/%m/%Y} ...")
    ini_dados = ini.replace(year=ini.year - 11) if dinamico else ini   # nota de preços: precisa de 10 anos antes
    hold, extra_h = {}, []
    if dinamico:                                          # V153: as ações dos ETFs (nota e nível de preço em cada data)
        try:
            hold = _sim_etf_hold([d for d in etfs if txt(d.get("ticker")).upper() in pesos], acoes)
        except Exception:
            hold = {}
        extra_h = sorted({a for h in hold.values() for a, _ in h} - set(pesos) - {bench})
    tab_h, sem_cambio = _sim_cotacoes_eur(list(pesos) + ([bench] if bench else []) + extra_h, moedas, ini_dados, fim)
    tab_h = tab_h.drop(columns=[t for t in extra_h if t in tab_h.columns])
    notas_p = niveis_p = None
    dias_v = []
    _SIM_LOCAL["espera"] = (0.0, 0)
    sec = {}
    tipos = {txt(d.get("ticker")).upper(): ("ETF" if d in etfs else "Ação") for d in acoes + etfs if txt(d.get("ticker"))}
    tab = tab_h.loc[tab_h.index >= pd.Timestamp(ini)]
    if not len(tab):
        print("Simulador: sem cotações nesse período.")
        sys.exit(1)
    compras = []
    ini_ef = max(ini, tab.index[0].to_pydatetime())       # sem cotações antes disto: a simulação começa aqui
    if p["ini_sim"].lower().startswith("s") and (p["ini_val"] or 0) > 0:
        compras.append((ini_ef, float(p["ini_val"])))
    if p["per"] != "Nenhum" and (p["per_val"] or 0) > 0 and not duplo:
        compras += [(d, float(p["per_val"])) for d in _sim_datas_reforco(ini, fim, p["per"]) if d >= ini_ef]
    compras_g = []
    if duplo:                                             # cada parte com o seu valor por reforço
        per = p["per"] if p["per"] != "Nenhum" else "Mensal"
        datas = [d for d in _sim_datas_reforco(ini, fim, per) if d >= ini_ef]
        soma_v = sum(v for _, _, v in grupos)
        for tk, n, v in grupos:
            cg = [(ini_ef, float(p["ini_val"]) * (v / soma_v if soma_v else 1 / len(grupos)))] if compras else []
            cg += [(d, v) for d in datas] if v > 0 else []
            compras_g.append(cg)
        compras = sorted([c for cg in compras_g for c in cg])
    if not compras:
        print("Simulador: sem dinheiro investido (investimento inicial e/ou reforço).")
        sys.exit(1)
    sem_dados = [t for t in pesos if t not in tab.columns]
    pesos_ok = {t: w for t, w in pesos.items() if t in tab.columns}
    vezes, vendido = {}, {}
    if dinamico:                                          # a nota do painel só nos dias de compra
        dias_c = sorted({tab.index[tab.index.searchsorted(pd.Timestamp(d))] for d, _ in compras
                         if tab.index.searchsorted(pd.Timestamp(d)) < len(tab.index)})
        dias_v = sorted(set(pd.Series(tab.index, index=tab.index).groupby(tab.index.to_period("M")).first()))
        dias_c = sorted(set(dias_c) | set(dias_v))
        loc = _SIM_LOCAL.get("tab")
        cols = [t for t in pesos if t in tab.columns and loc is not None and t in loc.columns]
        cols_h = [t for t in extra_h if loc is not None and t in loc.columns]
        sec = {}
        try:                                              # EUA: qualidade, saúde e preço pelos relatórios da SEC
            sec = sec_preparar([d for d in acoes if txt(d.get("ticker")).upper() in cols + cols_h], ini, fim)
        except Exception as e:
            print(f"   SEC: {e}")
        print(f"   a calcular a nota do painel de {len(cols)} títulos em {len(dias_c)} datas (só com os dados de cada data) ...")
        niveis_p = pd.DataFrame(index=pd.DatetimeIndex(dias_c), columns=cols, dtype=float)
        notas_p = _sim_notas_painel(loc[cols + cols_h], tipos, dias_c, sec, niveis_p, None, hold)[cols].reindex(tab.index)
        niveis_p = niveis_p.reindex(tab.index)
    if duplo:                                             # corre cada parte e soma
        unid, inv_t, fluxos = {}, {}, []
        sv = si = None
        for (tk, n, v), cg in zip(grupos, compras_g):
            tk = [t for t in tk if t in tab.columns]
            if not tk or not cg:
                continue
            if dinamico:
                u_, i_, sv_, si_, f_, vz, vd_ = _sim_correr_dinamico(tab[tk], notas_p.reindex(columns=tk), cg, "n", n,
                                                                     p.get("rodar") == SIM_RODAR[1], nmin,
                                                                     verificar=dias_v)
                vezes.update(vz)
                vendido.update(vd_)
            else:
                u_, i_, sv_, si_, f_ = _sim_correr(tab, {t: 1.0 for t in tk}, cg, fim)
            unid.update(u_)
            inv_t.update(i_)
            fluxos += f_[:-1]
            sv = sv_ if sv is None else sv + sv_
            si = si_ if si is None else si + si_
        if sv is None:
            print("Simulador: sem títulos com cotações para as duas partes.")
            sys.exit(1)
        fluxos = sorted(fluxos) + [(sv.index[-1].to_pydatetime(), float(sv.iloc[-1]))]
        pesos_ok = {t: 1.0 for t in inv_t}
        unid = {t: unid.get(t, 0.0) for t in pesos_ok}
    elif dinamico:
        modo = "n" if cod in (5, 10) else "nota"          # 6 e 12: todos os que passam o limite
        limite = (p["n"] or 10) if modo == "n" else (p["n"] if p["n"] is not None else 3.5)
        unid, inv_t, sv, si, fluxos, vezes, vendido = _sim_correr_dinamico(
            tab[list(pesos_ok)], notas_p.reindex(columns=list(pesos_ok)), compras, modo, limite, p.get("rodar") == SIM_RODAR[1], nmin,
            niveis_p.reindex(columns=list(pesos_ok)) if cod in (10, 12) else None, dias_v)
        pesos_ok = {t: 1.0 for t in inv_t}
        unid = {t: unid.get(t, 0.0) for t in pesos_ok}
    else:
        unid, inv_t, sv, si, fluxos = _sim_correr(tab, pesos_ok, compras, fim)
    investido, final = float(si.iloc[-1]), float(sv.iloc[-1])
    tir = _xirr(fluxos)
    queda = _sim_maior_queda(sv, si)
    b_final = b_tir = None
    sb = None
    if bench and bench in tab.columns:
        _, _, sb, _, fb = _sim_correr(tab, {bench: 1.0}, compras, fim)
        b_final, b_tir = float(sb.iloc[-1]), _xirr(fb)
    notas = []
    if sem_dados:
        notas.append("sem cotações no período: " + ", ".join(sem_dados))
    if sem_cambio:
        notas.append("sem câmbio para euros: " + ", ".join(sem_cambio))
    if bench and bench not in tab.columns:
        notas.append(f"comparação {bench}: sem cotações")
    tarde = [t for t in pesos_ok if tab[t].first_valid_index() is not None and
             tab[t].first_valid_index() > tab.index[0] + pd.Timedelta(days=10)]
    if tarde:
        notas.append("só com cotações depois do início (entram quando começam): " + ", ".join(tarde[:12]))
    esp_cx, esp_n = _SIM_LOCAL.get("espera", (0.0, 0))
    if dinamico and esp_n:
        notas.append(f"em {esp_n} data(s) de compra nenhum título cumpria os critérios: o dinheiro ficou à espera e "
                     "entrou no 1.º mês em que algum passou")
    if dinamico and esp_cx > 0.01:
        notas.append(f"ATENÇÃO: € {esp_cx:,.2f} ficaram por investir até ao fim (nenhum título voltou a cumprir os "
                     "critérios) — experimenta baixar a nota mínima (C13) ou o nível mínimo (F6)".replace(",", " "))
    if nmin is not None:
        notas.append(f"só se compram títulos com nota ≥ {nmin:g} na data da compra" if dinamico else
                     f"só entram títulos com nota geral de hoje ≥ {nmin:g}")
    if ini_ef > ini + pd.Timedelta(days=7):
        notas.append(f"sem cotações antes de {ini_ef:%d/%m/%Y}: a simulação começa aí")
    if dinamico:
        notas.append(f"reclassificado em cada data pela nota do painel (as mesmas regras da nota geral) só com os "
                     f"dados de cada data; ações dos EUA com qualidade, saúde e preço pelos relatórios anuais da SEC já "
                     f"divulgados nessa data ({len(sec)} empresas); nas outras, sem qualidade e saúde (contam como o "
                     f"sinal); analistas nunca (não há histórico); {p.get('rodar')}; {len(vezes)} títulos diferentes "
                     "escolhidos ao longo do tempo; a lista de títulos é a de hoje (as que desapareceram não entram)")
    elif cod in (2, 3, 7, 9, 11):
        notas.append("escolha feita pela nota de hoje: o passado fica favorecido")
    infl_s, infl_fonte = None, None
    try:
        infl_s, infl_fonte = _sim_inflacao()
    except Exception as e:
        notas.append(f"inflação indisponível ({e})")
    infl_inv = None                                 # investido corrigido pela inflação, em cada dia da série

    def corrigido(datas_):
        niv_d = _sim_infl_em(infl_s, datas_)
        niv_c = _sim_infl_em(infl_s, [d for d, _ in compras])
        out = []
        for dd, nd in zip(datas_, niv_d):
            out.append(sum(v * nd / nc for (dc, v), nc in zip(compras, niv_c) if pd.Timestamp(dc) <= pd.Timestamp(dd)))
        return out
    if infl_s is not None:
        infl_fim = corrigido([sv.index[-1]])[0]
        n0, n1 = _sim_infl_em(infl_s, [sv.index[0], sv.index[-1]])
        infl_acum = n1 / n0 - 1
        notas.append(f"inflação: {infl_fonte}")
    else:
        infl_fim = infl_acum = None
    final_real = final / (1 + infl_acum) if infl_acum is not None else None
    vals = [investido, final, final - investido, (final / investido - 1) if investido else None, tir,
            final_real, (final - infl_fim) if infl_fim is not None else None, infl_acum, queda or 0,
            b_final if b_final is not None else (f"(sem {bench})" if bench else "–"), b_tir,
            f"{len(pesos_ok)} de {len(pesos)}" if not dinamico else
            f"{sum(1 for t in unid if unid[t] > 0)} na carteira no fim; {len(vezes)} escolhidos pelo caminho",
            f"{tab.index[0]:%d/%m/%Y} a {tab.index[-1]:%d/%m/%Y} · {len(compras)} compra(s)",
            "; ".join(notas) or "–"]
    resumo = [(f"F{SIM_R0 + k}", v) for k, v in enumerate(vals)]
    titulos = []
    ultimo = tab.ffill().iloc[-1]
    for t in pesos_ok:
        comp = inv_t.get(t, 0.0)
        if comp <= 0:
            continue
        vf = unid.get(t, 0.0) * float(ultimo[t]) if t in ultimo and pd.notna(ultimo[t]) else 0.0
        vd = vendido.get(t, 0.0)
        res_ = vf + vd - comp
        titulos.append([nomes.get(t, t), t, round(comp, 2), round(vd, 2), round(vf, 2), round(res_, 2), res_ / comp])
    titulos.sort(key=lambda x: -x[4])
    n_ = len(sv.index)                              # 150 pontos espaçados (o gráfico usa sempre as 150 linhas)
    pts = [sv.index[round(k * (n_ - 1) / (SIM_MAX_SER - 1))] for k in range(SIM_MAX_SER)] if n_ >= SIM_MAX_SER \
        else list(sv.index)
    if infl_s is not None:                          # carteira em euros da data inicial (descontada da inflação)
        niv_p = _sim_infl_em(infl_s, pts)
        n_ini = _sim_infl_em(infl_s, [sv.index[0]])[0]
        ci = [float(sv[d_]) * n_ini / nv for d_, nv in zip(pts, niv_p)]
    else:
        ci = [None] * len(pts)
    serie = [[d_.to_pydatetime(), round(float(si[d_]), 2), round(float(sv[d_]), 2),
              round(float(sb[d_]), 2) if sb is not None else None, round(c_, 2) if c_ is not None else None]
             for d_, c_ in zip(pts, ci)]
    vazio = lambda nl, nc: [[None] * nc for _ in range(nl)]
    tit_w = (titulos + vazio(SIM_MAX_TIT, 7))[:SIM_MAX_TIT]
    ser_w = (serie + vazio(SIM_MAX_SER, 5))[:SIM_MAX_SER]
    ops = [{"op": "escrever", "folha": FOLHA_SIM, "rng": cel, "valores": [[v]]} for cel, v in resumo]
    ops += [{"op": "escrever", "folha": FOLHA_SIM, "rng": f"N{SIM_LIN_TIT}:T{SIM_LIN_TIT + SIM_MAX_TIT - 1}",
             "valores": tit_w},
            {"op": "escrever", "folha": FOLHA_SIM, "rng": f"U{SIM_LIN_SER}:Y{SIM_LIN_SER + SIM_MAX_SER - 1}",
             "valores": ser_w}]
    ex.ops(ops, guardar=True)
    _sim_gravar(res={"v": 3, "resumo": resumo, "titulos": titulos,
                     "serie": [[d_.strftime("%Y-%m-%d"), a, b, c, e] for d_, a, b, c, e in serie]})
    print(f"Investido € {investido:,.2f} -> valor final € {final:,.2f} ({(final / investido - 1):+.1%}); "
          f"TIR {tir:+.2%}/ano" if tir is not None else f"Investido € {investido:,.2f} -> € {final:,.2f}")
    if b_final is not None:
        print(f"Comparação {bench}: € {b_final:,.2f}" + (f" (TIR {b_tir:+.2%}/ano)" if b_tir is not None else ""))


FOLHA_CMP = "Comparar notas"


CMP_TIT = ["Top 10 pela NOTA DO PAINEL SÓ COM DADOS ANTIGOS — as regras da nota geral sem o que não há para datas "
           "passadas (analistas, valor justo, qualidade, saúde). Ao lado: a nota geral de hoje e o lugar nela",
           "Top 10 pela NOTA DE PREÇOS — a das opções 5, 6 e 8 do Simulador (tendência 40%, médias 35%, momento 25%). "
           "Ao lado: a nota geral de hoje e o lugar nela"]
CMP_COLS = (2, 8)                              # B, H
CMP_LINS = (6, 23)                             # ações, ETFs


def folha_comparar(wb):
    """Folha «Comparar notas»: a nota geral (todos os critérios) vs as notas que se conseguem calcular para datas
    antigas (só cotações). Preenchida pelo botão ⚖ COMPARAR."""
    ws = wb.create_sheet(FOLHA_CMP)
    ws["A1"] = "As notas de datas antigas são parecidas com a nota geral de hoje? (para ver se o Simulador é verosímil)"
    ws["A1"].font = Font(size=14, bold=True)
    ws["B2"] = ("Hoje há todos os dados; numa data antiga só há cotações. Os dois quadros classificam os títulos de hoje "
                "só com o que se saberia no passado e mostram, ao lado, a nota geral (com tudo) e o lugar que cada um "
                "tem nela. Quanto mais títulos em comum com o top 10 da nota geral e mais perto de 1 a correlação, "
                "mais verosímil é simular o passado com essa nota. Em baixo: as percentagens usadas em cada nota.")
    ws["B2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B2"].font = Font(italic=True, size=9, color="595959")
    ws.merge_cells("B2:L2")
    ws.row_dimensions[2].height = 40
    botao(ws, "N3", "⚖ COMPARAR", "painelacoes://comparar", "7030A0", "4B0082", 13)
    ws.merge_cells("N3:P4")
    for col, w in zip("ABCDEFGHIJKLMNOP", (2, 5, 30, 9, 9, 9, 2, 5, 30, 9, 9, 9, 2, 10, 10, 10)):
        ws.column_dimensions[col].width = w
    pc = lambda k: f"{PESOS.get(k, 0) * 100:.0f}%"
    linhas_p = [
        "Percentagens usadas em cada nota (ações):",
        f"• Nota geral (com tudo): sinal {pc('pont')} + qualidade {pc('qual')} + saúde {pc('saude')} + tendência "
        f"{pc('tend')} (tendência = {pc('t10')} 10 anos + {pc('t5')} 5 anos). O sinal (1 a 5) usa as médias 100/200 E os "
        "analistas e o valor justo; sem analistas fica no máximo 4.",
        f"• Nota do painel só com dados antigos: sem qualidade nem saúde, que contam como o sinal (é a regra do painel "
        f"para o que falta) -> na prática sinal {PESOS.get('pont', 0) * 100 + PESOS.get('qual', 0) * 100 + PESOS.get('saude', 0) * 100:.0f}% "
        f"+ tendência {pc('tend')}, e o sinal só pelas médias (máximo 4).",
        "• Nota de preços: tendência de longo prazo 40% + sinal das médias 35% + momento 6 meses 25%.",
        f"ETFs: nota geral = sinal {pc('etf_pont')} + nota das ações que têm {pc('etf_acoes')} + tendência "
        f"{pc('etf_tend')}; com dados antigos a parte das ações conta como o sinal."]
    for i, t in enumerate(linhas_p):
        c = ws.cell(40 + i, 2, t)
        c.font = Font(bold=(i == 0), size=10 if i == 0 else 9)
    g = _sim_guardados().get("cmp3")
    if g:
        for cel, v in g:
            ws[cel] = v
    _cmp_formatar(ws)


def _cmp_formatar(ws):
    az = PatternFill("solid", fgColor="1F3B57")
    nomes = ("Nota dados antigos", "Nota preços")
    for r0 in CMP_LINS:
        ws.cell(r0, 2).font = Font(bold=True, size=12)
        for k, c0 in enumerate(CMP_COLS):
            t = ws.cell(r0 + 1, c0, CMP_TIT[k])
            t.font = Font(bold=True, size=9, color="1F3B57")
            t.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r0 + 1, start_column=c0, end_row=r0 + 1, end_column=c0 + 4)
            outra = nomes[k]
            hs = ("Pos.", "Título", outra, "Nota geral", "Pos. na global")
            for j, h in enumerate(hs, c0):
                c = ws.cell(r0 + 2, j, h)
                c.font, c.fill = Font(bold=True, color="FFFFFF", size=9), az
                c.alignment = Alignment(horizontal="center", wrap_text=True, vertical="center")
            for i in range(10):
                for j in (c0 + 2, c0 + 3):
                    ws.cell(r0 + 3 + i, j).number_format = "0.0"
            ws.cell(r0 + 14, c0).font = Font(italic=True, size=9)
        ws.row_dimensions[r0 + 1].height = 50
        ws.row_dimensions[r0 + 2].height = 30


def _nota_dados_antigos(c, tipo, pos6=None, n6=None, extra=None):
    """Nota geral do painel calculada só com as cotações até ao último dia de c (sem analistas, valor justo,
    qualidade e saúde: o que falta conta como o sinal, como no painel)."""
    import numpy as np
    c = c.dropna()
    c = c[c > 0]
    if len(c) < 200:
        return None
    sig = volatilidade(c.to_numpy(dtype=float))
    s_ = sinal_estrategia(c.to_numpy(dtype=float), sig)
    if not s_:
        return None
    d = {"tipo": tipo, "mm100c": s_["mmc"], "mm200c": s_["mm200"], "gap5": s_["gap5"], "venda_min": s_["stop"],
         "stop": s_["stop"]}
    if len(c) > 126:
        d["ref6m"] = float(c.iloc[-127])
    m = c.resample("ME").last().dropna()
    if len(m) >= 36:
        reta = lambda cc: float(np.exp(np.polyfit(np.arange(len(cc)) / 12.0, np.log(cc.to_numpy(dtype=float)), 1)[0]) - 1)
        s_all = reta(m)
        s5 = reta(m.iloc[-61:]) if len(m) >= 61 else None
        s10 = reta(m.iloc[-121:]) if len(m) >= 121 else None
        d["lp"] = min(s_all, s5) if s5 is not None else s_all
        d["lp10"] = s10 if s10 is not None else s_all
        d["lp5"] = s5 if s5 is not None else s_all
        d["val_sempre"], d["val10"], d["val5"] = s_all, s10, s5
    d.update(inclin_recentes(c))
    if pos6 is not None:
        d["_pos6"], d["_n6"] = pos6, n6
    if extra:
        ex_ = dict(extra)
        pu = ex_.pop("_preco_usd", None)
        d.update(ex_)
        if pu:
            d["preco"] = pu
    pontuar(d, float(c.iloc[-1]))
    try:
        return nota_global(d)
    except Exception:
        return None


def comparar_notas():
    """⚖ COMPARAR: nota geral da folha vs nota de preços e vs nota do painel só com dados antigos (hoje)."""
    from datetime import datetime as dt_
    import pandas as pd
    ex = Excel(FICHEIRO)
    ultima = get_column_letter(len(COLS))
    r = ex.ops([{"op": "ler", "folha": idx, "rng": f"A{L0}:{ultima}{L0 + MAX_LINHAS}"} for idx in (1, 2)])
    grupos = [("AÇÕES", list(linhas_por_cabecalho(r[0], 1))), ("ETFs", list(linhas_por_cabecalho(r[1], 2)))]
    todos = sorted({txt(d.get("ticker")).upper() for _, ls in grupos for d in ls if txt(d.get("ticker"))})
    hoje = dt_.now()
    try:
        carregar_pesos(ex)                                # as percentagens da folha Ponderações (como estão agora)
    except SystemExit:
        raise
    except Exception:
        pass
    print(f"Comparar notas: {len(todos)} títulos; a descarregar 11 anos de cotações ...")
    tab, _ = _sim_cotacoes_eur(todos, {t: "EUR" for t in todos}, hoje.replace(year=hoje.year - 11), hoje)
    notas_p = _sim_notas_precos(tab).ffill().iloc[-1]
    sec = {}
    try:
        sec = sec_preparar(grupos[0][1], hoje.replace(year=hoje.year - 4), hoje)
    except Exception as e:
        print(f"   SEC: {e}")
    m6 = {}
    for t in tab.columns:                           # ETFs: lugar pela subida em 6 meses (como no painel)
        cc = tab[t].dropna()
        if len(cc) > 126:
            m6[t] = float(cc.iloc[-1] / cc.iloc[-127] - 1)
    cel = []
    for (nome, ls), r0 in zip(grupos, CMP_LINS):
        etf = nome == "ETFs"
        if etf:
            ord6 = sorted([txt(d.get("ticker")).upper() for d in ls if txt(d.get("ticker")).upper() in m6],
                          key=lambda t: -m6[t])
            pos = {t: (i / (len(ord6) - 1) if len(ord6) > 1 else 0.0) for i, t in enumerate(ord6)}
        lin = []
        for d in ls:
            t = txt(d.get("ticker")).upper()
            ng = _sim_num(d.get("nota"))
            if not t or ng is None:
                continue
            npz = notas_p.get(t)
            npz = float(npz) if npz is not None and pd.notna(npz) else None
            nh = None
            if t in tab.columns:
                ext = None
                if not etf and t in sec:
                    try:
                        ext = sec[t].em(tab.index[-1])
                    except Exception:
                        ext = None
                nh = _nota_dados_antigos(tab[t], "ETF" if etf else "Ação", pos.get(t) if etf else None,
                                         len(ord6) if etf else None, ext)
            lin.append((txt(d.get("nome")) or t, ng, npz, nh))
        if not lin:
            continue
        por_g = sorted(lin, key=lambda x: -x[1])
        pos_g = {x[0]: i + 1 for i, x in enumerate(por_g)}
        resumo = []
        for k, (c0, idx_n, rot) in enumerate(zip(CMP_COLS, (3, 2), ("painel com dados antigos", "nota de preços"))):
            lista = sorted([x for x in lin if x[idx_n] is not None], key=lambda x: -x[idx_n])
            for i in range(10):
                x = lista[i] if i < len(lista) else None
                vals = [i + 1, x[0], x[idx_n], x[1], pos_g.get(x[0])] if x else [None] * 5
                for j, v in enumerate(vals):
                    cel.append((f"{get_column_letter(c0 + j)}{r0 + 3 + i}", v))
            comuns = len({x[0] for x in por_g[:10]} & {x[0] for x in lista[:10]})
            rho = (pd.Series([x[1] for x in lista]).rank().corr(pd.Series([x[idx_n] for x in lista]).rank())
                   if len(lista) > 5 else None)
            cel.append((f"{get_column_letter(c0)}{r0 + 14}",
                        f"{comuns} dos 10 primeiros em comum com o top 10 da nota geral; correlação das ordens "
                        f"{rho:+.2f} ({len(lista)} títulos; 1 = mesma ordem, 0 = sem relação)" if rho is not None
                        else f"{comuns} dos 10 primeiros em comum com o top 10 da nota geral"))
            resumo.append(f"{rot} {comuns}/10" + (f" (correlação {rho:+.2f})" if rho is not None else ""))
            print(f"   {nome} · {resumo[-1]}")
        cel.append((f"B{r0}", f"{nome} — em comum com a nota geral: " + " · ".join(resumo)))
    cel.append(("N6", f"Calculado em {hoje:%d/%m/%Y %H:%M}"))
    limpar = {f"{get_column_letter(c)}{r}": None for r0 in CMP_LINS for r in [r0] + list(range(r0 + 3, r0 + 15))
              for c in range(2, 13)}
    limpar.update(dict(cel))
    ex.ops([{"op": "escrever", "folha": FOLHA_CMP, "rng": k, "valores": [[v]]} for k, v in limpar.items()],
           guardar=True)
    _sim_gravar(cmp3=cel)


FICH_TEMPOS = os.path.join(PASTA, "_Painel_acoes_tempos.json")   # durações das últimas execuções (estimativa)


CANCELAR = {"pedido": False, "critico": 0}       # botão Cancelar da janela de progresso


def _zona_critica(f):
    """Enquanto se escreve no Excel / se grava o ficheiro, um Cancelar espera pelo fim dessa escrita
    (para o livro nunca ficar a meio); logo a seguir, o script para."""
    def g(*a, **k):
        CANCELAR["critico"] += 1
        try:
            return f(*a, **k)
        finally:
            CANCELAR["critico"] -= 1
            if CANCELAR["pedido"] and CANCELAR["critico"] == 0:
                registo("cancelado pelo utilizador")
                os._exit(2)
    return g


Excel.ops = _zona_critica(Excel.ops)
guardar_e_abrir = _zona_critica(guardar_e_abrir)


def acao_do_botao(arg):
    """(chave, descrição) da ação pedida, para a janela de progresso e para a estimativa do tempo."""
    a = arg.lower()
    for chave, padrao, txt_ in (
            ("reclassificar", "reclassificar", "Reclassificar e reordenar"), ("simular", "simular", "Simulador"), ("compra", "://compra", "Calcular compra"), ("historico", "://historico", "Histórico da classificação"), ("evolucao", "://evolucao", "Evolução das compras"), ("comprent", "://comprent", "Comparar rentabilidade"), ("cfalt", "://cfalt", "Classificação alternativa"), ("repor", "repor_detidos", "Repor detidos"), ("alvojusto", "://alvojusto", "Analistas vs valor justo"), ("comparar", "comparar", "Comparar notas"),
            ("ordenar", "ordenar", "Ordenar"),
            ("procurar", "procurar", "Procurar"), ("top", "://top", "Mais significativas"),
            ("base", "://base", "Lista base"), ("acrescentar", "acrescentar", "Acrescentar marcados"),
            ("precos", "precos", "Atualizar preços"), ("modo", "://modo", "Mudar de bolsa e atualizar tudo"),
            ("acoes", "atualizar/acoes", "Atualizar ações"), ("acoes", "atualizar//acoes", "Atualizar ações"),
            ("etfs", "atualizar/etfs", "Atualizar ETFs"), ("etfs", "atualizar//etfs", "Atualizar ETFs")):
        if padrao in a:
            return chave, txt_
    return "tudo", "Atualizar tudo"


def com_janela_progresso(funcao):
    """Corre funcao() numa janela do Windows (em vez da consola): mostra as mensagens, o tempo decorrido e o tempo
    estimado (pelas últimas execuções da mesma ação) e fecha-se sozinha no fim. Se houver erro, fica aberta."""
    import tkinter as tk
    from tkinter import ttk
    import threading
    import queue
    import time
    import builtins
    chave, titulo = acao_do_botao(" ".join(sys.argv[1:]))
    try:
        with open(FICH_TEMPOS, encoding="utf-8") as f:
            tempos = json.load(f)
    except Exception:
        tempos = {}
    hist = [t for t in tempos.get(chave, []) if isinstance(t, (int, float)) and t > 0]
    estimado = sorted(hist)[len(hist) // 2] if hist else None
    fila = queue.Queue()
    estado = {"fim": None, "erro": False}

    dec2 = re.compile(r"(?<![\w.])(-?\d+\.\d{3,})(?![\w.])")    # só na janela: números com 2 casas

    class Escritor:
        def write(self, txt_):
            if txt_:
                fila.put(dec2.sub(lambda m: f"{float(m.group(1)):.2f}", txt_.replace("\r", "")))

        def flush(self):
            pass
    sys.stdout = sys.stderr = Escritor()
    builtins.input = lambda prompt="": (print(prompt) if prompt else None) or ""   # nunca fica à espera

    def trabalho():
        try:
            import pythoncom                            # o Excel (COM) fora da thread principal
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            funcao()
            registo("fim: ok")
        except SystemExit as e:
            if e.code not in (None, 0):
                registo(f"parou: {e.code}")
                estado["erro"] = True
                if not isinstance(e.code, int):
                    print(e.code)
            else:
                registo("fim: ok")
        except Exception:
            registo("ERRO: " + traceback.format_exc().strip().replace("\n", " | "))
            traceback.print_exc()
            estado["erro"] = True
        estado["fim"] = time.time()

    raiz = tk.Tk()
    raiz.title(f"Painel de ações · {titulo}")
    raiz.geometry("760x430")
    raiz.minsize(480, 260)
    try:
        raiz.call("tk", "scaling", raiz.winfo_fpixels("1i") / 72.0)
    except Exception:
        pass
    st = ttk.Style()
    try:
        st.theme_use("vista")
    except Exception:
        pass
    topo = ttk.Frame(raiz, padding=(12, 10, 12, 4))
    topo.pack(fill="x")
    ttk.Label(topo, text=f"{titulo}  ·  {TXT_VERSAO}", font=("Segoe UI", 11, "bold")).pack(side="left")
    barra = ttk.Progressbar(raiz, mode="determinate" if estimado else "indeterminate", maximum=1000)
    barra.pack(fill="x", padx=12, pady=(4, 2))
    if not estimado:
        barra.start(15)
    rot = ttk.Label(raiz, padding=(12, 0, 12, 4), foreground="#404040")
    rot.pack(fill="x")
    baixo = ttk.Frame(raiz, padding=(12, 4, 12, 10))   # botões: arrumados antes do texto, para ficarem sempre
    baixo.pack(side="bottom", fill="x")                # à vista (o texto encolhe se a janela for pequena)
    caixa = ttk.Frame(raiz, padding=(12, 4, 12, 4))
    caixa.pack(fill="both", expand=True)
    txt_w = tk.Text(caixa, font=("Consolas", 9), wrap="word", relief="solid", borderwidth=1, background="#FBFBFB",
                    height=10)
    sb = ttk.Scrollbar(caixa, command=txt_w.yview)
    txt_w.configure(yscrollcommand=sb.set, state="disabled")
    sb.pack(side="right", fill="y")
    txt_w.pack(side="left", fill="both", expand=True)
    txt_w.tag_configure("erro", foreground="#C00000")
    bt = ttk.Button(baixo, text="Fechar", command=raiz.destroy, state="disabled")
    bt.pack(side="right")

    def cancelar():
        from tkinter import messagebox
        if estado["fim"] is not None or CANCELAR["pedido"]:
            return
        if not messagebox.askyesno("Cancelar", f"Cancelar «{titulo}»?\n\nO que já foi escrito no Excel fica; "
                                   "o resto não é feito.", parent=raiz):
            return
        CANCELAR["pedido"] = True
        if CANCELAR["critico"] == 0:
            registo("cancelado pelo utilizador")
            os._exit(2)
        bc.configure(state="disabled", text="A cancelar...")
        print("\nA cancelar: espero só que acabe a escrita no Excel que está a decorrer...")
    bc = ttk.Button(baixo, text="Cancelar", command=cancelar)
    bc.pack(side="right", padx=(0, 8))
    manter = {"sim": False}
    bm = ttk.Button(baixo, text="Manter aberta", command=lambda: manter.update(sim=True) or bm.pack_forget())
    t0 = time.time()

    def mmss(x):
        x = int(round(x))
        return f"{x // 60}:{x % 60:02d}"

    def fechar_pedido():
        if estado["fim"] is not None:
            raiz.destroy()                              # durante o trabalho não fecha (o Excel ficaria a meio)
    raiz.protocol("WM_DELETE_WINDOW", fechar_pedido)

    def ciclo():
        novo = []
        while True:
            try:
                novo.append(fila.get_nowait())
            except queue.Empty:
                break
        if novo:
            txt_w.configure(state="normal")
            for x in novo:
                txt_w.insert("end", x, "erro" if ("Traceback" in x or "Error" in x or "ERRO" in x) else ())
            txt_w.see("end")
            txt_w.configure(state="disabled")
        agora = time.time()
        if estado["fim"] is None:
            dec = agora - t0
            if estimado:
                barra["value"] = min(dec / estimado, 0.99) * 1000
                falta = estimado - dec
                rot["text"] = (f"Decorrido {mmss(dec)}  /  estimado ~{mmss(estimado)}  ({min(dec / estimado, 0.99):.0%})"
                               + (f"  ·  faltam ~{mmss(falta)}" if falta > 0 else "  ·  quase a acabar..."))
            else:
                rot["text"] = f"Decorrido {mmss(dec)}  (primeira vez desta ação: ainda sem estimativa)"
            raiz.after(250, ciclo)
            return
        dur = estado["fim"] - t0
        barra.stop()
        barra.configure(mode="determinate")
        barra["value"] = 1000
        bt.configure(state="normal")
        bc.pack_forget()
        if estado["erro"]:
            rot.configure(text=f"Parou ao fim de {mmss(dur)}: vê as mensagens acima.", foreground="#C00000")
            raiz.bell()
            return
        if not estado.get("gravado"):
            estado["gravado"] = True
            try:
                tempos[chave] = (hist + [round(dur, 1)])[-5:]
                with open(FICH_TEMPOS, "w", encoding="utf-8") as f:
                    json.dump(tempos, f)
            except Exception:
                pass
            estado["fecha"] = agora                     # sem erro: fecha logo no fim
        if manter["sim"]:
            rot.configure(text=f"Concluído em {mmss(dur)}.")
            return
        resta = estado["fecha"] - agora
        if resta <= 0:
            raiz.destroy()
            return
        rot.configure(text=f"Concluído em {mmss(dur)}. A janela fecha-se em {int(resta) + 1} s.", foreground="#1E8E3E")
        raiz.after(250, ciclo)
    threading.Thread(target=trabalho, daemon=True).start()
    raiz.after(100, ciclo)
    raiz.mainloop()


def usar_janela_progresso():
    """Janela de progresso: no Windows, nos botões que trabalham (não na janela da posição nem no servidor), quando
    não há consola (pythonw) ou quando se pede com --janela-progresso."""
    if os.name != "nt" and "--janela-progresso" not in sys.argv:
        return False
    a = " ".join(sys.argv[1:]).lower()
    if any(x in a for x in ("--servidor", "--janela", "--modelo", "posicao", "fechar_quadro", "repor_detidos")) and \
            "--janela-progresso" not in sys.argv:
        return False
    return sys.stdout is None or "--janela-progresso" in sys.argv


if __name__ == "__main__" and usar_janela_progresso():
    if "--janela-progresso" in sys.argv:
        sys.argv.remove("--janela-progresso")
    com_janela_progresso(main)
elif __name__ == "__main__":
    try:
        main()
        registo("fim: ok")
    except SystemExit as e:
        if e.code not in (None, 0):
            registo(f"parou: {e.code}")
            print(e.code)
            if "--servidor" in sys.argv:
                sys.exit(1)
            input("\nCarrega em Enter para fechar...")
        else:
            registo("fim: ok")
    except Exception:
        registo("ERRO: " + traceback.format_exc().strip().replace("\n", " | "))
        traceback.print_exc()
        if "--servidor" in sys.argv:
            sys.exit(1)
        input("\nOcorreu um erro (acima). Carrega em Enter para fechar...")
