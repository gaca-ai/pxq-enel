# -*- coding: utf-8 -*-
"""
BW PxQ and SAMP PxQ, both against the Mercado Receita (MR), by market option and
family, absolute (R$ / MWh / kW) and relative. Basis of projeto_samp.md 4.4.

ERC is billed outside the BW, so it is left out of the BW x MR side (as in
check_mr_bloco2.py) and, since 29/09/2026, also from the SAMP side (rule unknown).

Usage:
    python compara_bw_samp_mr.py 2026-03
Writes SAMP/1-Entrega/checks/resultados/compara_bw_samp_mr_<ref>.xlsx (one sheet, long table)
"""
import sys

import pandas as pd

import lib_samp as eng

REF = sys.argv[1] if len(sys.argv) > 1 else "2026-03"
SAIDA = eng.ROOT / "SAMP" / "1-Entrega" / "checks" / "resultados" / f"compara_bw_samp_mr_{REF}.xlsx"

samp = eng.carrega_pxq_samp(REF)
bw = eng.carrega_pxq_bw(REF)
mr_sem = eng.familias_mr(REF, incluir_erc=False)


def por_opcao(df):
    g = df.groupby("OPCAO")[eng.MEDIDAS].sum()
    g.loc["TOTAL"] = g.sum()
    return g


linhas = []
for nome, px, mr in (("BW", bw, mr_sem), ("SAMP", samp, mr_sem)):
    a, b = por_opcao(px), por_opcao(mr)
    for op in a.index.union(b.index):
        for m in eng.MEDIDAS:
            va = a.loc[op, m] if op in a.index else 0.0
            vb = b.loc[op, m] if op in b.index else 0.0
            linhas.append({"FONTE_PXQ": nome, "OPCAO": op, "MEDIDA": m, "PXQ": va, "MR": vb,
                           "DELTA": va - vb, "DELTA_%": (va - vb) / vb if vb else None})
t = pd.DataFrame(linhas)
t.to_excel(SAIDA, index=False)
w = t.pivot_table(index=["OPCAO", "MEDIDA"], columns="FONTE_PXQ", values=["DELTA", "DELTA_%"])
pd.set_option("display.width", 250)
pd.set_option("display.max_rows", 200)
print(w.round(4).to_string())
print("saida:", SAIDA)
