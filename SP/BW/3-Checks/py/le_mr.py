# -*- coding: utf-8 -*-
"""Load the Mercado Receita (MR) rows of one reference month.

Source order (first file that has the month wins):
  1. src/Mercado_Receita_Enel_SP_<yyyy.mm>.xlsx  (single-month file)
  2. src/MR_SP_periodo_tarifario*.xlsx           (full tariff period, filtered by AnoMes)

Months whose rows are all TipoMercado = PROJETADO are refused: they are a copy
of a previous month, not realized market (e.g. jun/26 in the 2025-26 period).
"""
from pathlib import Path

import pandas as pd


def carrega_mr(base: Path, ref: str) -> tuple[pd.DataFrame, Path]:
    """Return (MR rows of month `ref` = 'yyyy-mm', source file)."""
    raw = Path(base) / "src"
    mes = pd.Timestamp(f"{ref}-01")

    # 1. single-month file keeps its own total rows; callers drop rows without Subgrupo
    arq_mes = raw / f"Mercado_Receita_Enel_SP_{ref.replace('-', '.')}.xlsx"
    if arq_mes.exists():
        return pd.read_excel(arq_mes), arq_mes

    # 2. tariff-period files
    for arq in sorted(raw.glob("MR_SP_periodo_tarifario*.xlsx")):
        df = pd.read_excel(arq)
        df = df[pd.to_datetime(df["AnoMes"], errors="coerce") == mes].copy()
        if df.empty:
            continue
        if (df["TipoMercado"].astype(str).str.upper() == "PROJETADO").all():
            raise ValueError(f"MR de {ref} em {arq.name} e PROJETADO (copia de outro mes) - nao usar no check")
        return df, arq

    raise FileNotFoundError(f"nenhum arquivo do Mercado Receita em {raw} contem o mes {ref}")


# MR component columns by family: economic + CVA + FinCT (shared by
# check_mr_bloco2.py and check_mr_camadas.py). Duplicated aggregates stay out on purpose.
FAMILIA_MR = {
    "FIO A": ["TUST_RB", "CVA_TUST_RB", "finCT_TUST_RB",
              "TUST_FR", "CVA_TUST_FR", "finCT_TUST_FR",
              "Conexao_T", "FinCT_Conexao_T", "Conexao_D", "FinCT_Conexao_D",
              "CUSD", "FinCT_CUSD",
              "TUSDG_T", "CVA_TUSDG_T", "FINCT_TUSDG_T",
              "ITAIPU", "CVA_TranItaipu", "FinCT_Itaipu",
              "TUST_ITAIPU", "CVA_TUST_ITAIPU", "FINCT_TUST_ITAIPU",
              "TUST_CI", "CVA_TUST_CI", "FINCT_TUST_CI"],
    "FIO B": ["DISTRIBUICAO", "FinCT_ParcelaB"],
    "ENCARGOS": ["covidTusdEco", "CVA_CovidTusd", "FinCT_CDECOVIDTUSD",
                 "TFSEE", "finCT_TFSEE", "PeD", "FinCT_Ped", "ONS", "FinCT_ONS",
                 "CCC", "CVA_CCC", "FinCT_CCC", "CDE", "CVA_CDE", "FinCT_CDE",
                 "PROINFA", "CVA_PROINFA", "finCT_PROINFA",
                 "CVA_liminar", "FINCT_Liminar",
                 "TUSDG_ONS", "CVA_TUSDG_ONS", "FINCT_TUSDG_ONS",
                 "covidTeEco", "CVA_CovidTe_ou_CVA_CDE_E", "FinCT_CDECOVIDTE_ou_FinCT_CDE_E",
                 "CDE_Elet_Eco", "CVA_CDE_Elet", "FinCT_CDE_Elet",
                 "CDE_GD_Eco", "CVA_CDE_GD", "FinCT_CDE_GD",
                 "ESSERR", "CVA_ESSERR", "FinCT_ESSERR",
                 "CFURH", "CVA_CFURH", "FinCT_CFURH",
                 "RGR", "FINCT_RGR", "CDE_E", "CVA_CDE_E", "FinCT_CDE_E"],
    "PERDAS": ["perdastec", "CVA_perdastec", "FinCT_perdastec",
               "perdasrb_d", "CVA_perdasrb_d", "FINCT_perdasrb_d",
               "perdasnt", "CVA_perdasnt", "FinCT_perdasnt",
               "perdasrb_c", "CVA_perdasrb_c", "FINCT_perdasrb_c"],
    "ENERGIA": ["EnergiaRevenda", "CVA_EnergiaRevenda", "FinCT_EnergiaRevenda",
                "ANGRA_Eco", "CVA_ANGRA", "FinCT_ANGRA"],
    "RI": ["RI", "FINCT_RI"],
    "OUTROS": ["OUTROS", "FinCT_Subsidio",
               "Subs_BENEFICIO_L14299_TUSD", "Subs_BENEFICIO_L14299_TE"],
}
# MR columns that belong to the TE parcel (energy tariff); everything else is
# TUSD. PeD and FinCT_Subsidio are single MR columns covering TUSD and TE - they
# are treated as TUSD (their TE share is small).
PARCELA_TE = {
    "EnergiaRevenda", "CVA_EnergiaRevenda", "FinCT_EnergiaRevenda",
    "ANGRA_Eco", "CVA_ANGRA", "FinCT_ANGRA",
    "ESSERR", "CVA_ESSERR", "FinCT_ESSERR",
    "CFURH", "CVA_CFURH", "FinCT_CFURH",
    "covidTeEco", "CVA_CovidTe_ou_CVA_CDE_E", "FinCT_CDECOVIDTE_ou_FinCT_CDE_E",
    "CDE_Elet_Eco", "CVA_CDE_Elet", "FinCT_CDE_Elet",
    "CDE_GD_Eco", "CVA_CDE_GD", "FinCT_CDE_GD",
    "CDE_E", "CVA_CDE_E", "FinCT_CDE_E",
    "perdasrb_c", "CVA_perdasrb_c", "FINCT_perdasrb_c",
    "ITAIPU", "CVA_TranItaipu", "FinCT_Itaipu",
    "TUST_ITAIPU", "CVA_TUST_ITAIPU", "FINCT_TUST_ITAIPU",
    "TUST_CI", "CVA_TUST_CI", "FINCT_TUST_CI",
    "Subs_BENEFICIO_L14299_TE",
}
