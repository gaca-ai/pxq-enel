# -*- coding: utf-8 -*-
"""
Test 2 (projeto_samp.md section 4): PxQ fed by SAMP quantities and PCAT tariffs,
checked against the Mercado Receita (MR) by market option, subgroup and family.
PxQ from PxQ_SAMP.ipynb (read, not recalculated); MR families as in SP/BW/.../check_mr_bloco2.py, ERC left out on both sides (29/09/2026).

Usage:
    python pxq_samp_check_mr.py 2026-03
Reads  SAMP/1-Entrega/pxq/pxq_samp_<ref>.xlsx (PxQ_SAMP.ipynb)
Writes SAMP/1-Entrega/checks/resultados/check_pxq_samp_mr_<ref>.xlsx
"""
import sys

import pandas as pd

import lib_samp as eng

REF = sys.argv[1] if len(sys.argv) > 1 else "2026-03"
SAIDA = eng.ROOT / "SAMP" / "1-Entrega" / "checks" / "resultados" / f"check_pxq_samp_mr_{REF}.xlsx"

base = eng.carrega_pxq_samp(REF)
mr = eng.familias_mr(REF, incluir_erc=False)
med_px, med_mr, chaves = eng.alinha(eng.medidas(base), eng.medidas(mr))

memoria = [
    ("Receita SAMP (alvo do bloco 2)", base["RECEITA_BRL"].sum(), None),
    ("Receita SAMP sem tarifa no bloco 1 (resíduo não ajustável)", base["RESIDUO_NAO_AJUSTAVEL"].sum(), None),
    ("Bloco 1 PxQ SAMP (TA cheia) x MR Total com Tar de aplicação",
     base["BLOCO1_TA"].sum(), mr["Total com Tar de aplicação"].sum()),
    ("1ª passada PxQ SAMP (desconto sobre Fio A+B) x MR Subs_CarFI (sinal invertido)",
     base["AJUSTE_1A_BRL"].sum(), -mr["Subs_CarFI"].sum()),
    ("2ª passada PxQ SAMP (fechamento de conta)", base["AJUSTE_2A_BRL"].sum(), None),
]
conciliacao = [
    ("MR - resíduo de arredondamento (NAO_ALOCADO)", 0.0, float(mr["NAO_ALOCADO"].sum())),
    ("MR - subsídio não explicado por abatimento de quantidade", 0.0,
     float(mr["SUBSIDIO_NAO_DISTRIBUIDO"].sum())),
]
cols_det = (eng.GRUPO_SAMP + ["CLASSE", "SUBCLASSE", "DETALHE", "ACESSANTE", "CHAVE_E", "REGRA_E",
                              "CHAVE_D", "REGRA_D", "DEMANDA_KW", "E_TUSD_MWH", "E_TE_MWH",
                              "DESC_DEM_PCT", "DESC_E_PCT", "REGIME", "RECEITA_BRL", "BLOCO1_TA",
                              "DESC_FIO", "AJUSTE_1A_BRL", "AJUSTE_2A_BRL", "BLOCO2_TA",
                              "RESIDUO_NAO_AJUSTAVEL"]
            + [f"B1_{f}" for f in eng.FAMILIAS] + [f"B2_{f}" for f in eng.FAMILIAS])
eng.escreve_comparacao(
    SAIDA, f"PxQ SAMP (quantidade SAMP x tarifa PCAT) bloco 2 x Mercado Receita - {REF}",
    "PxQ SAMP [2]", "MR", med_px, med_mr, chaves, conciliacao=conciliacao, memoria=memoria,
    detalhe=base[cols_det], nome_detalhe="Base_PxQ_SAMP")

sem = base[((base["E_TUSD_MWH"] + base["E_TE_MWH"] != 0) & base["CHAVE_E"].isna())
           | ((base["DEMANDA_KW"] != 0) & base["CHAVE_D"].isna())]
print(f"grupos SAMP: {len(base)} | sem tarifa: {len(sem)}")
print("fallbacks de chave:", pd.concat([base["REGRA_E"], base["REGRA_D"]]).value_counts().to_dict())
for rot, vp, vm in memoria:
    print(f"  {rot:<80} PxQ {vp:>18,.2f}" + (f" | MR {vm:>18,.2f} | delta {vp - vm:>15,.2f}" if vm is not None else ""))
for _, m, tt, _u in eng.ABAS:
    a, b = med_px[m].sum(), med_mr[m].sum()
    print(f"  {tt:<26} PxQ SAMP [2] {a:>18,.2f} | MR {b:>18,.2f} | delta {a - b:>16,.2f} "
          f"({100 * (a - b) / b if b else float('nan'):+.2f}%)")
print("saida:", SAIDA)
