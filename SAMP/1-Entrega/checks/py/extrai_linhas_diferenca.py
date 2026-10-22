# -*- coding: utf-8 -*-
"""
Extract the raw SAMP rows behind the SAMP x Mercado Receita (MR) revenue gap.

Groups are compared at the full grain (tipo de mercado, opcao, subgrupo, modalidade,
classe, subclasse, detalhe, posto, agente). Revenue on both sides excludes the
compensated energy (SAMP "Receita energia compensada" x MR line "Sistema de
Compensacao GD II - SCEE"): that is a presentation difference only (projeto_samp.md 3).
A group is kept when |SAMP - MR| > LIMITE_BRL. Every raw SAMP row of a kept group is
written (all metrics, original layout), with the group totals appended.

Usage:
    python extrai_linhas_diferenca.py 2026-03
Writes SAMP/1-Entrega/checks/resultados/samp_linhas_diferenca_<ref>.csv (";" separator, "," decimal)
"""
import sys

import numpy as np
import pandas as pd

import lib_samp as eng

REF = sys.argv[1] if len(sys.argv) > 1 else "2026-03"
LIMITE_BRL = 1000.0
SAIDA = eng.ROOT / "SAMP" / "1-Entrega" / "checks" / "resultados" / f"samp_linhas_diferenca_{REF}.csv"
G = eng.GRUPO_SAMP

# ------------------------------------------------------------------ SAMP (raw rows)
samp = eng.le_samp(REF)
samp["CLASSE_SAMP"] = samp["CLASSE_SAMP"].replace({"Serviço Público": "Serviço público"})
receita = samp["DscDetalheMercado"].isin(["Receita Energia (R$)", "Receita Demanda (R$)"])
energia = samp["DscDetalheMercado"] == "Energia TUSD (kWh)"
demanda = samp["DscDetalheMercado"] == "Demanda Faturada (kW)"
g_s = pd.DataFrame({
    "RECEITA_SAMP": samp[receita].groupby(G)["VlrMercado"].sum(),
    "ENERGIA_TUSD_MWH_SAMP": samp[energia].groupby(G)["VlrMercado"].sum() / 1000.0,
    "DEMANDA_KW_SAMP": samp[demanda].groupby(G)["VlrMercado"].sum(),
}).fillna(0.0)

# ------------------------------------------------------------------ MR (same grain)
mr, _ = eng._carrega_mr(eng.ROOT / "SP" / "BW", REF)  # single-month or tariff-period MR file
mr = mr[mr["Subgrupo"].notna() & (mr["TipoMercado"] != "Sistema de Compensação GD II - SCEE")].copy()
mr["DETALHE_SAMP"] = np.where(mr["OPÇÃO"] == "ERC", "ERC", mr["Detalhe"])
mr["OPCAO"] = mr["OPÇÃO"].replace({"LIVRE APE": "LIVRE", "ERC": "LIVRE"})
agente = {v: k for k, v in eng.ACESSANTE_PCAT.items()}
agente["TECIPAR"] = "Tecipar"
mr = mr.rename(columns={"TipoMercado": "TIPO_MERCADO", "Subgrupo": "SUBGRUPO",
                        "Modalidade": "MODALIDADE", "Classe": "CLASSE_SAMP",
                        "Subclasse": "SUBCLASSE_SAMP", "Posto": "POSTO"})
mr["AGENTE"] = mr["Agente"].replace(agente)
for c in G:
    mr[c] = mr[c].fillna(eng.NSA).astype(str).str.strip()
mr["CLASSE_SAMP"] = mr["CLASSE_SAMP"].replace({"Serviço Público": "Serviço público"})
g_m = mr.groupby(G).agg(RECEITA_MR=("FinTotAj", "sum"), ENERGIA_TUSD_MWH_MR=("TUSD_E", "sum"),
                        DEMANDA_KW_MR=("D", "sum"))

# ------------------------------------------------------------------ gap per group
g = g_s.join(g_m, how="outer").fillna(0.0)
g["DELTA_RECEITA"] = g["RECEITA_SAMP"] - g["RECEITA_MR"]
g["RS_MWH_SAMP"] = np.where(g["ENERGIA_TUSD_MWH_SAMP"] != 0,
                            g["RECEITA_SAMP"] / g["ENERGIA_TUSD_MWH_SAMP"].replace(0, np.nan), np.nan)
g["RS_MWH_MR"] = np.where(g["ENERGIA_TUSD_MWH_MR"] != 0,
                          g["RECEITA_MR"] / g["ENERGIA_TUSD_MWH_MR"].replace(0, np.nan), np.nan)
sel = g[g["DELTA_RECEITA"].abs() > LIMITE_BRL].copy()
sel["ORDEM_GRUPO"] = sel["DELTA_RECEITA"].abs().rank(ascending=False, method="first").astype(int)
so_mr = sel[~sel.index.isin(g_s.index)]

# ------------------------------------------------------------------ raw rows of the kept groups
raw = samp.merge(sel.reset_index(), on=G, how="inner")
raw = raw.sort_values(["ORDEM_GRUPO", "DscDetalheMercado"])
cols_orig = ["DatCompetencia", "SigAgenteDistribuidora", "TIPO_MERCADO", "OPCAO", "SUBGRUPO",
             "MODALIDADE", "CLASSE_SAMP", "SUBCLASSE_SAMP", "DETALHE_SAMP", "POSTO", "AGENTE",
             "IdeNucleoCeg", "DscDetalheMercado", "VlrMercado"]
cols_grupo = ["ORDEM_GRUPO", "RECEITA_SAMP", "RECEITA_MR", "DELTA_RECEITA", "ENERGIA_TUSD_MWH_SAMP",
              "ENERGIA_TUSD_MWH_MR", "RS_MWH_SAMP", "RS_MWH_MR", "DEMANDA_KW_SAMP", "DEMANDA_KW_MR"]
out = raw[cols_orig + cols_grupo].rename(columns={
    "TIPO_MERCADO": "NomTipoMercado", "OPCAO": "DscOpcaoEnergia", "SUBGRUPO": "DscSubGrupoTarifario",
    "MODALIDADE": "DscModalidadeTarifaria", "CLASSE_SAMP": "DscClasseConsumoMercado",
    "SUBCLASSE_SAMP": "DscSubClasseConsumidor", "DETALHE_SAMP": "DscDetalheConsumidor",
    "POSTO": "DscPostoTarifario", "AGENTE": "NomAgenteAcessante"})
out.to_csv(SAIDA, sep=";", decimal=",", index=False, encoding="utf-8-sig")

pd.set_option("display.width", 250)
pd.set_option("display.max_rows", 100)
print(f"grupos com |SAMP - MR| > R$ {LIMITE_BRL:,.0f}: {len(sel)} | linhas brutas do SAMP: {len(out)}")
print(f"delta dos grupos selecionados: R$ {sel['DELTA_RECEITA'].sum():,.2f} "
      f"| delta total (todos os grupos): R$ {g['DELTA_RECEITA'].sum():,.2f}")
if len(so_mr):
    print(f"ATENCAO: {len(so_mr)} grupo(s) so no MR (sem linha no SAMP), delta R$ {so_mr['DELTA_RECEITA'].sum():,.2f}")
print(sel.reset_index().groupby(["OPCAO", "TIPO_MERCADO"])["DELTA_RECEITA"].agg(["count", "sum"]).round(0))
print(sel.reset_index().sort_values("ORDEM_GRUPO").head(15)[
    ["ORDEM_GRUPO", "TIPO_MERCADO", "OPCAO", "SUBGRUPO", "MODALIDADE", "SUBCLASSE_SAMP", "DETALHE_SAMP",
     "POSTO", "RECEITA_SAMP", "RECEITA_MR", "DELTA_RECEITA", "RS_MWH_SAMP", "RS_MWH_MR"]].round(2).to_string())
print("saida:", SAIDA)
