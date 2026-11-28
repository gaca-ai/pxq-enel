# -*- coding: utf-8 -*-
"""Read one tariff vintage from the ANEEL open-data set "Componentes Tarifarias da TE e da TUSD"
(dadosabertos.aneel.gov.br, one parquet per year) and return it in the layout of tarifa_long.csv.

Used by extrai_pcat.py for the historical vintages (fonte = "ANEEL"). The open data is published
monthly and may lag, so recent vintages keep coming from the PCAT; for history it avoids the
layout changes between PCAT workbooks.

Rules (validated against the PCAT vintages RJ 2025/2026 and CE 2025/2026, 28/11/2026):
  - the parquet carries exact duplicate rows (about half of the file): they are dropped first,
    otherwise every tariff doubles; a key repeated with DIFFERENT values stops the run;
  - layers: "Tarifa de Aplicacao" -> TA, "Base Economica" -> BE, "CVA" -> CVA; the open data has no
    financial layer, so FIN = TA - BE - CVA, item by item (identity checked on all PCAT vintages);
  - units "R$/MWh" / "R$/kW" -> "MWh" / "kW"; all other key labels are identical to the PCAT;
  - component names follow COMPONENTES below; TE_Liminar + TUSD_Liminar -> Liminar1;
    TUSD_RGR has no column in the table and is dropped only while it is zero.
"""
from pathlib import Path

import pandas as pd

ATRIBUTOS = ["SUBGRUPO", "MODALIDADE", "CLASSE", "SUBCLASSE", "DETALHE",
             "POSTO", "UNIDADE", "ACESSANTE"]

# open-data column -> tariff-table column
CAMPOS = {
    "DscSubGrupoTarifario": "SUBGRUPO", "DscModalidadeTarifaria": "MODALIDADE",
    "DscClasseConsumidor": "CLASSE", "DscSubClasseConsumidor": "SUBCLASSE",
    "DscDetalheConsumidor": "DETALHE", "DscPostoTarifario": "POSTO", "DscUnidade": "UNIDADE",
    "SigNomeAgenteAcessante": "ACESSANTE", "DscBaseTarifaria": "CAMADA",
    "DscComponenteTarifario": "COMPONENTE", "VlrComponenteTarifario": "VALOR",
}
CAMADAS = {"Tarifa de Aplicação": "TA", "Base Econômica": "BE", "CVA": "CVA"}
UNIDADES = {"R$/MWh": "MWh", "R$/kW": "kW"}
# distributor code of the tariff table -> SigNomeAgente of the open data
AGENTES = {"RJ": "ENEL RJ", "CE": "ENEL CE", "SP": "ENEL SP"}

# open-data component -> tariff-table column (names not listed here are identical in both)
COMPONENTES = {
    "TE": "Total TE", "TUSD": "Total TUSD",
    "TE_SUBSIDIO": "TE Subsidio", "TUSD_SUBSIDIO": "TUSD Subsidio",
    "TUSD_OUTROS": "TUSD Outros", "TE_BENEFICIO_L14299": "TE BENEFICIO_L14299",
    "TE_Liminar": "Liminar1", "TUSD_Liminar": "Liminar1",
}
IDENTICOS = {
    "TE_ANGRA", "TE_CDE_COVID", "TE_CDE_ELET", "TE_CDE_GD", "TE_CFURH", "TE_ENERGIA", "TE_ESSERR",
    "TE_PeD", "TE_Per_RB", "TE_TRANSPORTE_ITAIPU", "TE_TUST_CI", "TE_TUST_ITAIPU",
    "TUSD BENEFICIO_L14299", "TUSDG_ONS", "TUSDG_T", "TUSD_CCC", "TUSD_CCD", "TUSD_CCT", "TUSD_CDE",
    "TUSD_CDE_COVID", "TUSD_CUSD", "TUSD_FR", "TUSD_FioB", "TUSD_ONS", "TUSD_PNT", "TUSD_PROINFA",
    "TUSD_PT", "TUSD_PeD", "TUSD_Per_RB_D", "TUSD_RB", "TUSD_RI", "TUSD_TFSEE",
}
DESCARTADOS_SE_ZERO = {"TUSD_RGR"}
TOTAIS = ["Total TUSD", "Total TE"]
TOL_TOTAL = 0.02   # R$/MWh or R$/kW: published total vs sum of its components (rounding)


def le_vigencia(pasta, dist, inicio, fim):
    """TA, BE, CVA and FIN of one vintage, wide (one column per item), keyed by CAMADA + ATRIBUTOS.

    pasta  : folder with componentes-tarifarias-<year>.parquet
    inicio : DatInicioVigencia of the vintage (the file is chosen by its year)
    fim    : expected DatFimVigencia (checked)
    """
    arquivo = Path(pasta) / f"componentes-tarifarias-{inicio.year}.parquet"
    if not arquivo.exists():
        raise FileNotFoundError(f"{arquivo} nao encontrado")
    bruto = pd.read_parquet(arquivo)
    df = bruto[(bruto["SigNomeAgente"] == AGENTES[dist])
               & (bruto["DatInicioVigencia"].astype(str) == inicio.isoformat())]
    if df.empty:
        raise ValueError(f"{dist}: vigencia iniciada em {inicio} ausente em {arquivo.name}")
    fins = set(df["DatFimVigencia"].astype(str))
    if fins != {fim.isoformat()}:
        raise ValueError(f"{dist} {inicio}: fim da vigencia na ANEEL {sorted(fins)}, esperado {fim}")

    # exact duplicates are a feature of the file; conflicting duplicates are an error
    n_bruto = len(df)
    df = df.drop_duplicates()
    n_dup = n_bruto - len(df)
    df = df.rename(columns=CAMPOS)[list(CAMPOS.values())]
    df["CAMADA"] = df["CAMADA"].map(CAMADAS)
    df["UNIDADE"] = df["UNIDADE"].map(UNIDADES)
    if df["CAMADA"].isna().any() or df["UNIDADE"].isna().any():
        raise ValueError(f"{dist} {inicio}: base tarifaria ou unidade fora do de-para")
    conflito = df.duplicated(subset=["CAMADA"] + ATRIBUTOS + ["COMPONENTE"]).sum()
    if conflito:
        raise ValueError(f"{dist} {inicio}: {conflito} componente(s) repetido(s) com valores diferentes")

    desconhecidos = set(df["COMPONENTE"]) - set(COMPONENTES) - IDENTICOS - DESCARTADOS_SE_ZERO
    if desconhecidos:
        raise ValueError(f"{dist} {inicio}: componentes fora do de-para: {sorted(desconhecidos)}")
    descartar = df["COMPONENTE"].isin(DESCARTADOS_SE_ZERO)
    if (df.loc[descartar, "VALOR"].abs() > 1e-9).any():
        raise ValueError(f"{dist} {inicio}: {sorted(DESCARTADOS_SE_ZERO)} com valor diferente de zero - "
                         f"criar a coluna na tabela de tarifas")
    df = df[~descartar].copy()
    df["COMPONENTE"] = df["COMPONENTE"].map(lambda c: COMPONENTES.get(c, c))

    larga = df.pivot_table(index=["CAMADA"] + ATRIBUTOS, columns="COMPONENTE", values="VALOR",
                           aggfunc="sum", fill_value=0.0).reset_index()
    larga.columns.name = None
    itens = sorted(c for c in larga.columns if c not in ["CAMADA"] + ATRIBUTOS + TOTAIS)

    # published totals must match the sum of their items
    for total, prefixo in (("Total TE", "TE"), ("Total TUSD", "TUSD")):
        # Liminar1 joins TE and TUSD court-order items, so it stays out of both checks
        soma = larga[[i for i in itens if i.startswith(prefixo)]].sum(axis=1)
        d = (larga[total] - soma).abs().max()
        if d > TOL_TOTAL:
            raise ValueError(f"{dist} {inicio}: {total} difere da soma dos itens em {d:.4f}")

    # same keys in the three published layers
    chaves = {c: set(map(tuple, g[ATRIBUTOS].values)) for c, g in larga.groupby("CAMADA")}
    if set(chaves) != {"TA", "BE", "CVA"} or not (chaves["TA"] == chaves["BE"] == chaves["CVA"]):
        raise ValueError(f"{dist} {inicio}: chaves diferentes entre TA, BE e CVA")

    # FIN = TA - BE - CVA, merged on the key (never positional)
    valores = TOTAIS + itens
    por = {c: larga[larga["CAMADA"] == c].set_index(ATRIBUTOS)[valores] for c in ("TA", "BE", "CVA")}
    fin = (por["TA"] - por["BE"] - por["CVA"]).reset_index()
    fin.insert(0, "CAMADA", "FIN")
    tabela = pd.concat([larga, fin], ignore_index=True)

    print(f"{dist} {inicio} (ANEEL, {arquivo.name}): {len(chaves['TA'])} chaves, {len(itens)} itens | "
          f"{n_dup} de {n_bruto} linhas eram duplicatas exatas (descartadas)")
    return tabela[["CAMADA"] + ATRIBUTOS + valores], itens
