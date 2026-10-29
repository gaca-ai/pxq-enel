"""
Extrai do relatorio mensal da RNF (aba "Relatorio SAP_ZUBIL_146") somente as colunas usadas no
calculo da abertura da RNF e do layout receita_nao_faturada (notebook 1-Entrega/Abertura_RNF.ipynb),
mais as quantidades de reativo (informativas), e grava um Excel com:
- aba ZUBIL_<ref>: linhas de detalhe (grupos ELGA/ELGB), sem subtotais, com a linha de origem;
- aba Colunas: letra, rotulo no ZUBIL e uso de cada coluna no calculo.

Uso:
    python 3-Py/2-Checks/extrai_zubil_colunas.py 2026-08
Saida:
    2-Checks/zubil/ZUBIL_146_<ref>_colunas_calculo.xlsx
"""
import os
import sys
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter as letra
from pyxlsb import open_workbook

BASE = Path(os.environ.get("RNF_BASE", r"C:\python_projects\claude\p_x_q\SP\RNF"))

# ZUBIL columns (0-based) that feed the RNF calculation and the receita_nao_faturada layout -> use
COLS = {2: "Classe principal SAP", 6: "Competência", 7: "Lote", 8: "Grupo (ELGA/ELGB)", 9: "Subclasse SAP",
        10: "Categoria tarifária", 18: "Valor TE (cativo)", 20: "Valor TUSD TE (cativo; demanda no livre)",
        22: "Valor TUSD (livre)", 26: "Valor reativo", 34: "Energia TE (cativo)", 38: "Energia TUSD (livre)",
        48: "Demanda (cativo e livre)", 58: "Valor bandeira",
        42: "Consumo reativo (informativo: o reativo entra só pelo valor, coluna UD/ER do layout)",
        54: "Demanda reativa (informativo: o reativo entra só pelo valor, coluna UD/ER do layout)"}
NUMERICAS = (18, 20, 22, 26, 34, 38, 42, 48, 54, 58)
COLS = dict(sorted(COLS.items()))                 # report column order


def extrai(ref):
    """Write the reduced ZUBIL of month `ref` (AAAA-MM) and return the output path."""
    aaaa, mm = ref[:4], ref[5:7]
    arq = BASE / "src" / f"RNF CCS {mm}{aaaa} Fator K Nova Metodologia.xlsb"
    assert arq.exists(), f"arquivo nao encontrado: {arq}"
    with open_workbook(str(arq)) as wb:
        abas = {n.strip(): n for n in wb.sheets}
        with wb.get_sheet(abas["Relatório SAP_ZUBIL_146"]) as sh:
            linhas = [[c.v for c in r] for r in sh.rows()]
    cab = linhas[2]                                   # header row of the sheet (row 3)
    assert str(cab[10]).strip() == "Categoria Tarifa", "layout do ZUBIL mudou"

    # detail rows only: subtotal rows have no ELGA/ELGB group
    dados = []
    for n, r in enumerate(linhas[3:], start=4):
        r = r + [None] * 70
        if str(r[8]).strip() in ("ELGA", "ELGB"):
            dados.append([n] + [r[j] for j in COLS])
    nomes = {j: f"{letra(j + 1)} - {str(cab[j]).strip()}" for j in COLS}
    df = pd.DataFrame(dados, columns=["Linha no relatório"] + [nomes[j] for j in COLS])
    for j in NUMERICAS:
        df[nomes[j]] = pd.to_numeric(df[nomes[j]], errors="coerce").fillna(0.0)
    legenda = pd.DataFrame([(letra(j + 1), str(cab[j]).strip(), uso) for j, uso in COLS.items()],
                           columns=["Coluna", "Rótulo no ZUBIL", "Uso no cálculo"])

    saida = BASE / "2-Checks" / "zubil" / f"ZUBIL_146_{ref}_colunas_calculo.xlsx"
    saida.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(saida, engine="openpyxl") as w:
        df.to_excel(w, sheet_name=f"ZUBIL_{ref}", index=False)
        legenda.to_excel(w, sheet_name="Colunas", index=False)
        ws = w.book[f"ZUBIL_{ref}"]
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        for i in range(1, ws.max_column + 1):
            ws.column_dimensions[letra(i)].width = 18
    print(f"{ref}: {len(df)} linhas -> {saida}")
    return saida


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit(__doc__)
    for ref in sys.argv[1:]:
        extrai(ref)
