# -*- coding: utf-8 -*-
"""
Check SAMP (ANEEL open data) x Mercado Receita (MR), one reference month.

Purpose: validate the MR sent to the tariff process against what the
distributor declared to ANEEL in the SAMP (Sistema de Acompanhamento de
Informacoes de Mercado para Regulacao Economica). Both come from the same
billing, so quantities should match exactly; revenue is compared with the MR
FinTotAj (tariff of application net of subsidies).

Grain: TIPO_MERCADO x OPCAO x SUBGRUPO x MODALIDADE x CLASSE x SUBCLASSE x
DETALHE x POSTO x AGENTE.

Normalizations (MR -> SAMP vocabulary):
  - OPCAO "LIVRE APE" -> LIVRE (SAMP keeps APE in DETALHE)
  - OPCAO "ERC" -> LIVRE with DETALHE "ERC" (SAMP layout)
  - CLASSE "Serviço Público" -> "Serviço público"
  - AGENTE names mapped to the SAMP accessing-agent names (AGENTE_MR)
  - TipoMercado "Sistema de Compensação GD II - SCEE" (MR only) is kept as is;
    its SAMP counterpart is "Receita energia compensada (R$)"

Usage:
    python check_samp_mr.py 2026-03
Reads  SAMP/data/samp-<yyyy>.parquet and the MR of the month (SP/BW/3-Checks/py/le_mr.py)
Writes SAMP/1-Entrega/checks/resultados/check_samp_mr_<ref>.xlsx
"""
import os
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(os.environ.get("PXQ_ROOT", r"C:\python_projects\claude\p_x_q")) / "SP" / "BW" / "3-Checks" / "py"))
from le_mr import carrega_mr  # noqa: E402

BASE = Path(os.environ.get("PXQ_ROOT", r"C:\python_projects\claude\p_x_q"))
REF = sys.argv[1] if len(sys.argv) > 1 else "2026-03"
CNPJ_SP = "61695227000193"

ARQ_SAMP = BASE / "SAMP" / "data" / f"samp-{REF[:4]}.parquet"
SAIDA = BASE / "SAMP" / "1-Entrega" / "checks" / "resultados" / f"check_samp_mr_{REF}.xlsx"

GRAO = ["TIPO_MERCADO", "OPCAO", "SUBGRUPO", "MODALIDADE", "CLASSE",
        "SUBCLASSE", "DETALHE", "POSTO", "AGENTE"]
# measures compared on both sides
MEDIDAS = ["DEMANDA_KW", "ENERGIA_TUSD_MWH", "ENERGIA_TE_MWH", "UC",
           "RECEITA_BRL", "PIS_COFINS_BRL", "ICMS_BRL"]
TOL = {"DEMANDA_KW": 1.0, "ENERGIA_TUSD_MWH": 1.0, "ENERGIA_TE_MWH": 1.0, "UC": 1.0,
       "RECEITA_BRL": 1000.0, "PIS_COFINS_BRL": 1000.0, "ICMS_BRL": 1000.0}

# MR agent name -> SAMP accessing-agent name
AGENTE_MR = {
    "Neoenergia Elektro": "ELEKTRO REDES S.A.",
    "CPFL Piratininga": "COMPANHIA PIRATININGA DE FORCA E LUZ",
    "Ceris": "COOPERATIVA DE ELETRIFICAÇÃO DA REGIÃO DE ITAPECERICA DA SERRA",
    "PCH RASGAO": "Rasgão",
    "UTE FERNANDO GASPARIAN (N.PIRAT)": "Nova Piratininga (Antiga Fernando Gasparian)",
    "UTE SAO JOAO BIOGAS": "São João Biogás",
    "PCH PIRAPORA": "Pirapora",
    "TECIPAR": "Tecipar",
}

# SAMP metric (DscDetalheMercado) -> internal measure; scale converts units
METRICA_SAMP = {
    "Demanda Faturada (kW)": ("DEMANDA_KW", 1.0),
    "Energia TUSD (kWh)": ("ENERGIA_TUSD_MWH", 1e-3),
    "Energia TE (kWh)": ("ENERGIA_TE_MWH", 1e-3),
    "Número de consumidores": ("UC", 1.0),
    "Número de Consumidores": ("UC", 1.0),
    "Receita Energia (R$)": ("RECEITA_BRL", 1.0),
    "Receita Demanda (R$)": ("RECEITA_BRL", 1.0),
    "Receita energia compensada (R$)": ("RECEITA_BRL", 1.0),
    "PIS/PASEP (R$)": ("PIS_COFINS_BRL", 1.0),
    "COFINS (R$)": ("PIS_COFINS_BRL", 1.0),
    "PIS/COFINS (R$)": ("PIS_COFINS_BRL", 1.0),
    "ICMS (R$)": ("ICMS_BRL", 1.0),
    # memo only (no MR counterpart)
    "Receita Ultrapassagem Demanda (R$)": ("MEMO_ULTRAPASSAGEM_BRL", 1.0),
    "Receita Bandeiras (R$)": ("MEMO_BANDEIRAS_BRL", 1.0),
    "Ultrapassagem Demanda (kW)": ("MEMO_ULTRAPASSAGEM_KW", 1.0),
    "Energia compensada (kWh)": ("MEMO_ENERGIA_COMPENSADA_MWH", 1e-3),
    "Demanda Contratada kW": ("MEMO_DEMANDA_CONTRATADA_KW", 1.0),
    "PROINFA (kWh)": ("MEMO_PROINFA_MWH", 1e-3),
    "Energia Consumida (kWh)": ("MEMO_ENERGIA_CONSUMIDA_MWH", 1e-3),
    "Energia Injetada (kWh)": ("MEMO_ENERGIA_INJETADA_MWH", 1e-3),
    # discount percentages are rates, not additive - handled separately
    "Desconto Demanda %": (None, 1.0),
    "Desconto Energia TUSD %": (None, 1.0),
    "Desconto Energia TE %": (None, 1.0),
    "Desconto de energia compensada %": (None, 1.0),
}


def norm(df):
    """Common text normalization of the grain columns."""
    for c in GRAO:
        df[c] = df[c].fillna("Não se aplica").astype(str).str.strip()
    df["CLASSE"] = df["CLASSE"].replace({"Serviço Público": "Serviço público"})
    return df


# ------------------------------------------------------------------ SAMP side
samp = pd.read_parquet(ARQ_SAMP)
# the CNPJ is read as an integer: compare as a 14-digit string
samp = samp[samp["NumCNPJAgenteDistribuidora"].astype(str).str.zfill(14) == CNPJ_SP]
samp = samp[pd.to_datetime(samp["DatCompetencia"]).dt.strftime("%Y-%m") == REF].copy()
assert len(samp), f"SAMP sem linhas para {REF}"
novas = set(samp["DscDetalheMercado"]) - set(METRICA_SAMP)
assert not novas, f"metricas do SAMP sem mapeamento: {novas}"

samp = samp.rename(columns={
    "NomTipoMercado": "TIPO_MERCADO", "DscOpcaoEnergia": "OPCAO",
    "DscSubGrupoTarifario": "SUBGRUPO", "DscModalidadeTarifaria": "MODALIDADE",
    "DscClasseConsumoMercado": "CLASSE", "DscSubClasseConsumidor": "SUBCLASSE",
    "DscDetalheConsumidor": "DETALHE", "DscPostoTarifario": "POSTO",
    "NomAgenteAcessante": "AGENTE"})
samp = norm(samp)
samp["MEDIDA"] = samp["DscDetalheMercado"].map(lambda m: METRICA_SAMP[m][0])
samp["VALOR"] = samp["VlrMercado"] * samp["DscDetalheMercado"].map(lambda m: METRICA_SAMP[m][1])

# discount rates, kept apart (weighted later by the quantity they apply to)
desc = samp[samp["DscDetalheMercado"].str.contains("%")].pivot_table(
    index=GRAO, columns="DscDetalheMercado", values="VlrMercado", aggfunc="mean")

s = samp[samp["MEDIDA"].notna()].pivot_table(
    index=GRAO, columns="MEDIDA", values="VALOR", aggfunc="sum", fill_value=0.0)

# ------------------------------------------------------------------ MR side
# MR of the month: single-month file or tariff-period file (le_mr.py)
mr, ARQ_MR = carrega_mr(BASE / "SP" / "BW", REF)
mr = mr[mr["Subgrupo"].notna()].copy()
mr["TIPO_MERCADO"] = mr["TipoMercado"]
mr["OPCAO"] = mr["OPÇÃO"]
mr["DETALHE"] = mr["Detalhe"]
mr.loc[mr["OPCAO"] == "ERC", "DETALHE"] = "ERC"
mr["OPCAO"] = mr["OPCAO"].replace({"LIVRE APE": "LIVRE", "ERC": "LIVRE"})
mr["SUBGRUPO"] = mr["Subgrupo"]
mr["MODALIDADE"] = mr["Modalidade"]
mr["CLASSE"] = mr["Classe"]
mr["SUBCLASSE"] = mr["Subclasse"]
mr["POSTO"] = mr["Posto"]
mr["AGENTE"] = mr["Agente"].replace(AGENTE_MR)
mr = norm(mr)
for c in ["D", "Daj", "TUSD_E", "TUSD_Eaj", "TE_E", "TE_Eaj", "UC", "FinTotAj",
          "PisPasep", "Cofins", "ICMS", "SubsidioTotal", "Total com Tar de aplicação"]:
    mr[c] = mr[c].fillna(0.0)
mr["DEMANDA_KW"] = mr["D"]
mr["ENERGIA_TUSD_MWH"] = mr["TUSD_E"]
mr["ENERGIA_TE_MWH"] = mr["TE_E"]
mr["RECEITA_BRL"] = mr["FinTotAj"]
mr["PIS_COFINS_BRL"] = mr["PisPasep"] + mr["Cofins"]
mr["ICMS_BRL"] = mr["ICMS"]
mr["MR_DAJ_KW"] = mr["Daj"]
mr["MR_TUSD_EAJ_MWH"] = mr["TUSD_Eaj"]
mr["MR_SUBSIDIO_BRL"] = mr["SubsidioTotal"]
mr["MR_RECEITA_TA_BRL"] = mr["Total com Tar de aplicação"]
m = mr.groupby(GRAO)[MEDIDAS + ["MR_DAJ_KW", "MR_TUSD_EAJ_MWH", "MR_SUBSIDIO_BRL",
                                "MR_RECEITA_TA_BRL"]].sum()

# ------------------------------------------------------------------ compare
for c in MEDIDAS:
    if c not in s.columns:
        s[c] = 0.0
grao = s.join(m, how="outer", lsuffix="_SAMP", rsuffix="_MR")
grao = grao.rename(columns={c: f"{c}_SAMP" for c in s.columns if c not in MEDIDAS})
grao = grao.fillna(0.0)
grao["ORIGEM"] = np.select(
    [grao.index.isin(s.index) & grao.index.isin(m.index), grao.index.isin(s.index)],
    ["AMBOS", "SO_SAMP"], "SO_MR")
for c in MEDIDAS:
    grao[f"{c}_DELTA"] = grao[f"{c}_SAMP"] - grao[f"{c}_MR"]
grao["DIVERGE"] = np.logical_or.reduce(
    [grao[f"{c}_DELTA"].abs() > TOL[c] for c in MEDIDAS])
grao = grao.reset_index()

# summary per OPCAO and per TIPO_MERCADO
def resumo(chave):
    linhas = []
    for c in MEDIDAS:
        g = grao.groupby(chave)[[f"{c}_SAMP", f"{c}_MR"]].sum()
        g.columns = ["SAMP", "MR"]
        g.loc["TOTAL"] = g.sum()
        g["DELTA"] = g["SAMP"] - g["MR"]
        g["DELTA_%"] = np.where(g["MR"] != 0, g["DELTA"] / g["MR"].replace(0, np.nan), np.nan)
        g.insert(0, "MEDIDA", c)
        linhas.append(g.reset_index())
    return pd.concat(linhas, ignore_index=True)

res_opcao = resumo("OPCAO")
res_tipo = resumo("TIPO_MERCADO")

# discount check (free market / generation): SAMP declared % x MR 1 - Qaj/Q
dsc = grao.set_index(GRAO)[["DEMANDA_KW_MR", "MR_DAJ_KW", "ENERGIA_TUSD_MWH_MR",
                            "MR_TUSD_EAJ_MWH", "MR_SUBSIDIO_BRL"]].join(desc, how="outer")
dsc = dsc[dsc.filter(like="%").notna().any(axis=1) | (dsc["MR_SUBSIDIO_BRL"].fillna(0) != 0)]
dsc["MR_DESC_DEMANDA_%"] = np.where(dsc["DEMANDA_KW_MR"] > 0,
                                    100 * (1 - dsc["MR_DAJ_KW"] / dsc["DEMANDA_KW_MR"].replace(0, np.nan)), np.nan)
dsc["MR_DESC_ENERGIA_TUSD_%"] = np.where(dsc["ENERGIA_TUSD_MWH_MR"] > 0,
                                         100 * (1 - dsc["MR_TUSD_EAJ_MWH"] / dsc["ENERGIA_TUSD_MWH_MR"].replace(0, np.nan)), np.nan)
dsc = dsc.reset_index()

with pd.ExcelWriter(SAIDA, engine="openpyxl") as xw:
    res_opcao.to_excel(xw, sheet_name="0_Resumo_Opcao", index=False)
    res_tipo.to_excel(xw, sheet_name="1_Resumo_TipoMercado", index=False)
    grao[grao["DIVERGE"]].to_excel(xw, sheet_name="2_Divergencias", index=False)
    grao.to_excel(xw, sheet_name="3_Grao_Completo", index=False)
    dsc.to_excel(xw, sheet_name="4_Descontos", index=False)

pd.set_option("display.width", 250)
pd.set_option("display.max_rows", 200)
pd.set_option("display.max_columns", 30)
print(res_opcao.to_string())
print(res_tipo.to_string())
print("linhas no grao:", len(grao), "| divergentes:", int(grao["DIVERGE"].sum()))
print(grao["ORIGEM"].value_counts().to_string())
print("Saida:", SAIDA)
