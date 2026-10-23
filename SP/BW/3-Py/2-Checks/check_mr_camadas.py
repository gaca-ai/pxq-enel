# -*- coding: utf-8 -*-
"""
Check PxQ bloco 2 x Mercado Receita (MR) by tariff layer, month by month.

Question answered: how far is the PxQ from the MR of the tariff adjustment in
each family of the economic base (BE), and in the CVA and financial (FIN)
layers taken as totals? No split by option, class or subgroup.

Same rules as check_mr_bloco2.py: ERC/ERCAP out; the MR subsidy is spread over
its components by the MR's own quantity abatement (Daj/D, Eaj/E), so both sides
carry the discounts inside the components. MR columns are split by layer from
their names: CVA_* = CVA, FinCT_*/finCT_*/FINCT_* = FIN, the rest = BE.

Usage:
    python check_mr_camadas.py 2026-01 2026-05
Reads  1-Entrega/pxq/pxq_<ref>.xlsx (sheet Familias) and the MR of each month (le_mr.py)
Writes 2-Checks/consolidado_check_mr_camadas_<ini>_a_<fim>.xlsx
"""
import os
import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
from le_mr import FAMILIA_MR, PARCELA_TE, carrega_mr  # noqa: E402

BASE = Path(os.environ.get("PXQ_BASE", r"C:\python_projects\claude\p_x_q\SP\BW"))
INI, FIM = (sys.argv[1], sys.argv[2]) if len(sys.argv) > 2 else ("2026-01", "2026-05")
REFS = [p.strftime("%Y-%m") for p in pd.period_range(INI, FIM, freq="M")]
SAIDA = BASE / "2-Checks" / f"consolidado_check_mr_camadas_{INI}_a_{FIM}.xlsx"

FAMILIAS = ["FIO A", "FIO B", "ENCARGOS", "PERDAS", "ENERGIA", "RI", "OUTROS"]
LINHAS = [f"BE · {f}" for f in FAMILIAS] + ["CVA (total)", "FIN (total)", "TOTAL"]


def camada(col):
    """Layer of an MR component column, from its name."""
    if col.startswith("CVA_"):
        return "CVA"
    if col.lower().startswith("finct_"):
        return "FIN"
    return "BE"


def lado_mr(ref):
    """MR of the month by layer: BE per family, CVA and FIN totals (R$)."""
    mr, _ = carrega_mr(BASE, ref)
    mr = mr[mr["Subgrupo"].notna()].copy()
    mr = mr[mr["OPÇÃO"].replace({"LIVRE APE": "LIVRE"}) != "ERC"].copy()
    for c in ["D", "Daj", "TUSD_E", "TUSD_Eaj", "TE_E", "TE_Eaj"]:
        mr[c] = mr[c].fillna(0.0)
    # subsidy spread by the MR's own abatement factor (same as check_mr_bloco2.py)
    f_tusd = np.where(mr["D"] != 0, mr["Daj"] / mr["D"].replace(0, np.nan),
                      np.where(mr["TUSD_E"] != 0, mr["TUSD_Eaj"] / mr["TUSD_E"].replace(0, np.nan), 1.0))
    f_te = np.where(mr["TE_E"] != 0, mr["TE_Eaj"] / mr["TE_E"].replace(0, np.nan), 1.0)
    f_tusd, f_te = np.nan_to_num(f_tusd, nan=1.0), np.nan_to_num(f_te, nan=1.0)
    saida = dict.fromkeys(LINHAS, 0.0)
    for fam, cols in FAMILIA_MR.items():
        for c in cols:
            v = (mr[c].fillna(0.0) * (f_te if c in PARCELA_TE else f_tusd)).sum()
            k = f"BE · {fam}" if camada(c) == "BE" else f"{camada(c)} (total)"
            saida[k] += v
    saida["TOTAL"] = sum(v for k, v in saida.items() if k != "TOTAL")
    return saida


def lado_pxq(ref):
    """PxQ bloco 2 of the month by layer (tariff rows only)."""
    fam = pd.read_excel(BASE / "1-Entrega" / "pxq" / f"pxq_{ref}.xlsx", sheet_name="Familias")
    fam = fam[fam["Tarifável?"] == "SIM"]
    saida = {f"BE · {f}": fam[f"[2] BE · {f}"].fillna(0.0).sum() for f in FAMILIAS}
    saida["CVA (total)"] = fam[[f"[2] CVA · {f}" for f in FAMILIAS]].fillna(0.0).sum().sum()
    saida["FIN (total)"] = fam[[f"[2] FIN · {f}" for f in FAMILIAS]].fillna(0.0).sum().sum()
    saida["TOTAL"] = sum(saida.values())
    return saida


pxq = pd.DataFrame({ref: lado_pxq(ref) for ref in REFS}).reindex(LINHAS)
mr = pd.DataFrame({ref: lado_mr(ref) for ref in REFS}).reindex(LINHAS)
pxq["Acumulado"], mr["Acumulado"] = pxq.sum(axis=1), mr.sum(axis=1)
delta = pxq - mr
# relative to the MR magnitude: FIN is negative, so the sign follows PxQ - MR
pct = (delta / mr.abs().replace(0, np.nan)).where(mr.abs() > 0)

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    pct.to_excel(w, sheet_name="1_Delta_pct")
    delta.to_excel(w, sheet_name="2_Delta_R$")
    pxq.to_excel(w, sheet_name="3_PxQ_bloco2_R$")
    mr.to_excel(w, sheet_name="4_MR_R$")
    for nome in w.book.sheetnames:
        ws = w.book[nome]
        ws.column_dimensions["A"].width = 18
        for row in ws.iter_rows(min_row=2, min_col=2):
            for c in row:
                c.number_format = "0.00%" if nome.startswith("1_") else "#,##0"

pd.set_option("display.width", 200)
print((100 * pct).round(2).to_string())
print(f"\nsaida: {SAIDA}")
