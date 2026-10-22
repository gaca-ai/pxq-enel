# -*- coding: utf-8 -*-
"""
Reverse check: one PxQ built from the best source of each segment, compared with
the full BW PxQ of the same month.

Consolidated PxQ:
  - CATIVO: BW PxQ rows (flag "Cliente Cativo"), bloco 2 as delivered;
  - LIVRE, GERACAO, DISTRIBUICAO, SUPRIMENTO: SAMP PxQ (PxQ_SAMP.ipynb output).
    SAMP LIVRE already carries the BW hybrid customer (energy and demand) and the
    EMAE pumping stations, so nothing else comes from the BW. ERC is out of both
    sides since 29/09/2026 (not in the BW; left out of the SAMP PxQ).

Usage:
    python consolida_bw_samp.py 2026-03
Writes SAMP/1-Entrega/checks/resultados/consolidado_bw_samp_x_bw_<ref>.xlsx
"""
import sys

import numpy as np
import pandas as pd

import lib_samp as eng

REF = sys.argv[1] if len(sys.argv) > 1 else "2026-03"
SAIDA = eng.ROOT / "SAMP" / "1-Entrega" / "checks" / "resultados" / f"consolidado_bw_samp_x_bw_{REF}.xlsx"

bw = eng.carrega_pxq_bw(REF)
samp = eng.carrega_pxq_samp(REF)

cativo_bw = bw[bw["FLAG"] == "Cliente Cativo"].assign(ORIGEM="BW cativo")
resto_samp = samp[samp["OPCAO"] != "CATIVO"].assign(ORIGEM="SAMP")

COLS = ["ORIGEM", "OPCAO", "SUBGRUPO", "RECEITA_BRL"] + eng.MEDIDAS
cons = pd.concat([cativo_bw[COLS], resto_samp[COLS]], ignore_index=True)

med_c, med_b, chaves = eng.alinha(eng.medidas(cons), eng.medidas(bw))

# what each origin brings, against what the BW had for the same perimeter
orig = cons.groupby("ORIGEM")[["RECEITA_BRL", "MONETARIO_TOTAL"]].sum()
bw_livre_resto = bw[bw["FLAG"] != "Cliente Cativo"]
memoria = [
    ("Receita BW cativo (entra igual nos dois lados)", orig.loc["BW cativo", "RECEITA_BRL"],
     bw.loc[bw["FLAG"] == "Cliente Cativo", "RECEITA_BRL"].sum()),
    ("Receita livres + híbrido + SD: SAMP x BW",
     orig.drop(index="BW cativo")["RECEITA_BRL"].sum(), bw_livre_resto["RECEITA_BRL"].sum()),
    ("Resíduo não ajustável do SAMP (receita sem tarifa, suprimento TE)",
     resto_samp["RESIDUO_NAO_AJUSTAVEL"].sum(), None),
]
detalhe = cons.copy()
eng.escreve_comparacao(
    SAIDA, f"PxQ consolidado (BW cativo + SAMP demais opções) x PxQ BW - {REF}",
    "Consolidado [2]", "PxQ BW [2]", med_c, med_b, chaves, memoria=memoria,
    detalhe=detalhe, nome_detalhe="Base_Consolidado")

pd.set_option("display.width", 250)
print(f"consolidado x BW | {REF}")
for _, m, tt, _u in eng.ABAS:
    a, b = med_c[m].sum(), med_b[m].sum()
    print(f"  {tt:<26} consolidado {a:>18,.2f} | BW {b:>18,.2f} | delta {a - b:>16,.2f} "
          f"({100 * (a - b) / b if b else float('nan'):+.2f}%)")
for rot, a, b in memoria:
    print(f"  {rot:<85} {a:>16,.2f}" + (f" | {b:>16,.2f}" if b is not None else ""))
d = (med_c - med_b).groupby(level=0).sum()[["MONETARIO_TOTAL"] + eng.FAMILIAS + ["ENERGIA_TUSD_MWH", "DEMANDA_KW"]]
print("\ndelta por opcao (consolidado - BW):")
print((d).round(0).to_string())
print("saida:", SAIDA)
