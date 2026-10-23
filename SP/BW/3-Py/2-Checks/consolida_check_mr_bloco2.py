# -*- coding: utf-8 -*-
"""Consolidate the monthly check_mr_bloco2_<ref>.xlsx files into one workbook.

Goal: see whether the PxQ bloco 2 x Mercado Receita (MR) differences stay at the
same level month after month or swing and offset each other.

Usage:
    python consolida_check_mr_bloco2.py 2026-01 2026-05
        reads  2-Checks/resultados/check_mr_bloco2_<ref>.xlsx (every month in the range)
        writes 2-Checks/consolidado_check_mr_bloco2_<ini>_a_<fim>.xlsx
    python consolida_check_mr_bloco2.py 2026-01 2026-05 --fonte samp
        same, for the PxQ fed by SAMP (SAMP/1-Entrega/checks/py/pxq_samp_check_mr.py):
        reads  SAMP/1-Entrega/checks/resultados/check_pxq_samp_mr_<ref>.xlsx
        writes SAMP/1-Entrega/checks/consolidado_check_pxq_samp_mr_<ini>_a_<fim>.xlsx
Both inputs share the layout of check_mr_bloco2.py (tabs 1 to 11, header on row 3,
columns OPCAO | SUBGRUPO | PxQ | MR).

Monthly deltas are recomputed from the PxQ [2] and MR values (the monthly files
carry the deltas as formulas). Statistics per row are Excel formulas:
  Acumulado      = sum of the monthly deltas (a swing that offsets the next month
                   shows up as a small accumulated value with a large Desvio)
  Media / Desvio = mean and sample standard deviation of the monthly deltas
  CV             = Desvio / |Media|; low = stable level, high = the delta swings
  Meses mesmo sinal = months whose delta has the sign of Acumulado
RESIDUO = Monetario total - sum of the families (MR rounding/not allocated, MR
subsidy not explained by abatement, PxQ non-adjustable residual).
"""
import argparse
import os
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(os.environ.get("PXQ_BASE", r"C:\python_projects\claude\p_x_q\SP\BW"))

arg = argparse.ArgumentParser(description="Consolidate monthly PxQ x MR checks")
arg.add_argument("ini", nargs="?", default="2026-01")
arg.add_argument("fim", nargs="?", default="2026-05")
arg.add_argument("--fonte", choices=["bw", "samp"], default="bw",
                 help="bw = PxQ from the BW (default); samp = PxQ from the SAMP")
ARGS = arg.parse_args()
INI, FIM = ARGS.ini, ARGS.fim

# input folder, monthly file prefix, output file and label of the PxQ side per source
FONTES = {
    "bw": (BASE / "2-Checks" / "resultados", "check_mr_bloco2",
           BASE / "2-Checks", "consolidado_check_mr_bloco2", "PxQ bloco 2 (BW)"),
    "samp": (BASE.parents[1] / "SAMP" / "1-Entrega" / "checks" / "resultados", "check_pxq_samp_mr",
             BASE.parents[1] / "SAMP" / "1-Entrega" / "checks", "consolidado_check_pxq_samp_mr",
             "PxQ SAMP bloco 2"),
}
DIR_RES, PREFIXO, DIR_SAIDA, PREFIXO_SAIDA, ROTULO_PXQ = FONTES[ARGS.fonte]
SAIDA = DIR_SAIDA / f"{PREFIXO_SAIDA}_{INI}_a_{FIM}.xlsx"

MESES = [p.strftime("%Y-%m") for p in pd.period_range(INI, FIM, freq="M")]
NOME_MES = {"01": "jan", "02": "fev", "03": "mar", "04": "abr", "05": "mai", "06": "jun",
            "07": "jul", "08": "ago", "09": "set", "10": "out", "11": "nov", "12": "dez"}
ROTULO = {m: f"{NOME_MES[m[5:]]}/{m[2:4]}" for m in MESES}

# monthly tab -> (measure label, unit group)
ABAS = {
    "1_Energia_TUSD_MWh": ("Energia TUSD (MWh)", "QTD"),
    "2_Energia_TE_MWh": ("Energia TE (MWh)", "QTD"),
    "3_Demanda_kW": ("Demanda (kW)", "QTD"),
    "4_Monetario_Total_BRL": ("Monetário total", "BRL"),
    "5_FIO_A": ("FIO A", "BRL"),
    "6_FIO_B": ("FIO B", "BRL"),
    "7_ENCARGOS": ("ENCARGOS", "BRL"),
    "8_PERDAS": ("PERDAS", "BRL"),
    "9_ENERGIA": ("ENERGIA", "BRL"),
    "10_RI": ("RI", "BRL"),
    "11_OUTROS": ("OUTROS", "BRL"),
}
FAMILIAS = ["FIO A", "FIO B", "ENCARGOS", "PERDAS", "ENERGIA", "RI", "OUTROS"]
QTDS = ["Energia TUSD (MWh)", "Energia TE (MWh)", "Demanda (kW)"]
ORDEM_OPCAO = ["CATIVO", "DISTRIBUIÇÃO", "GERAÇÃO", "LIVRE", "SUPRIMENTO"]

# ------------------------------------------------------------------ read months
partes = []
for ref in MESES:
    arq = DIR_RES / f"{PREFIXO}_{ref}.xlsx"
    if not arq.exists():
        raise FileNotFoundError(f"falta {arq.name} - rodar o check do mes {ref} antes")
    for aba, (medida, unidade) in ABAS.items():
        d = pd.read_excel(arq, sheet_name=aba, header=2, usecols=[0, 1, 2, 3])
        d.columns = ["OPCAO", "SUBGRUPO", "PXQ", "MR"]
        # detail rows only: subtotals and total are rebuilt here
        d = d[d["SUBGRUPO"].notna() & d["OPCAO"].notna()].copy()
        d["REF"], d["MEDIDA"], d["UNIDADE"] = ref, medida, unidade
        partes.append(d)
base = pd.concat(partes, ignore_index=True)
base[["PXQ", "MR"]] = base[["PXQ", "MR"]].astype(float).fillna(0.0)
base["DELTA"] = base["PXQ"] - base["MR"]

# Families do not add up exactly to the monetary total: the MR side keeps the
# rounding / not-allocated residual and the subsidy not explained by abatement,
# and the PxQ side keeps the non-adjustable residual (see 0_Resumo of each month).
# That gap becomes an explicit RESIDUO row so every block closes to the total.
chave = ["REF", "OPCAO", "SUBGRUPO"]
fam = base[base["MEDIDA"].isin(FAMILIAS)].groupby(chave)[["PXQ", "MR"]].sum()
tot = base[base["MEDIDA"] == "Monetário total"].set_index(chave)[["PXQ", "MR"]]
res = (tot - fam.reindex(tot.index).fillna(0.0)).reset_index()
res["MEDIDA"], res["UNIDADE"] = "RESÍDUO", "BRL"
res["DELTA"] = res["PXQ"] - res["MR"]
base = pd.concat([base, res[base.columns]], ignore_index=True)
FAMILIAS = FAMILIAS + ["RESÍDUO"]


def pivota(df, linhas):
    """Months in columns for PXQ, MR and DELTA, keyed by `linhas`."""
    g = df.groupby(linhas + ["REF"], sort=False)[["PXQ", "MR", "DELTA"]].sum().reset_index()
    return g.pivot_table(index=linhas, columns="REF", values=["PXQ", "MR", "DELTA"],
                         aggfunc="sum", fill_value=0.0, sort=False)


def ordena_opcao(df, col="OPCAO"):
    ordem = {o: i for i, o in enumerate(ORDEM_OPCAO + ["TOTAL"])}
    return df.assign(_o=df[col].map(ordem).fillna(99)).sort_values("_o", kind="stable").drop(columns="_o")


# ------------------------------------------------------------------ layout helpers
FONTE_T = Font(bold=True, size=12)
NEGRITO = Font(bold=True)
CAB = PatternFill("solid", fgColor="1F4E78")
CAB_F = Font(bold=True, color="FFFFFF")
SUB = PatternFill("solid", fgColor="DDEBF7")
EST = PatternFill("solid", fgColor="FFF2CC")
FINO = Side(style="thin", color="BFBFBF")
BORDA = Border(top=FINO, bottom=FINO, left=FINO, right=FINO)
FMT_BRL = '#,##0;[Red]-#,##0'
FMT_PCT = '0.00%;[Red]-0.00%'
FMT_CV = '0.00'

ESTATS = ["Acumulado", "Média", "Desvio", "Mín", "Máx", "CV", "Meses mesmo sinal"]


def escreve_bloco(ws, lin, titulo, chaves, rotulos_chave, valores, fmt, com_estat=True,
                  negrito=None, pct_de=None):
    """Write one block: key columns, one column per month, then statistics.

    valores: DataFrame indexed like `chaves`, columns = MESES.
    pct_de : optional (numerador_df, denominador_df) to write delta % instead.
    Returns the next free row.
    """
    ws.cell(row=lin, column=1, value=titulo).font = FONTE_T
    lin += 1
    cab = rotulos_chave + [ROTULO[m] for m in MESES] + (ESTATS if com_estat else [])
    for j, t in enumerate(cab, start=1):
        c = ws.cell(row=lin, column=j, value=t)
        c.fill, c.font, c.border = CAB, CAB_F, BORDA
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    lin += 1
    nk, nm = len(rotulos_chave), len(MESES)
    for i, chave in enumerate(chaves):
        chave = chave if isinstance(chave, tuple) else (chave,)
        for j, v in enumerate(chave, start=1):
            ws.cell(row=lin, column=j, value=v).border = BORDA
        for k, m in enumerate(MESES):
            v = float(valores.loc[chave if len(chave) > 1 else chave[0], m])
            c = ws.cell(row=lin, column=nk + 1 + k, value=v)
            c.number_format, c.border = fmt, BORDA
        if com_estat:
            a = f"{get_column_letter(nk + 1)}{lin}"
            b = f"{get_column_letter(nk + nm)}{lin}"
            rng = f"{a}:{b}"
            col = nk + nm + 1
            ac = f"{get_column_letter(col)}{lin}"
            me = f"{get_column_letter(col + 1)}{lin}"
            de = f"{get_column_letter(col + 2)}{lin}"
            formulas = [f"=SUM({rng})", f"=AVERAGE({rng})", f"=STDEV({rng})",
                        f"=MIN({rng})", f"=MAX({rng})",
                        f'=IF(ABS({me})<1E-9,"",{de}/ABS({me}))',
                        f'=IF({ac}>=0,COUNTIF({rng},">=0"),COUNTIF({rng},"<0"))']
            fmts = [fmt, fmt, fmt, fmt, fmt, FMT_CV, "0"]
            for k, (f, fm) in enumerate(zip(formulas, fmts)):
                c = ws.cell(row=lin, column=col + k, value=f)
                c.number_format, c.border, c.fill = fm, BORDA, EST
        if negrito is not None and negrito(chave):
            for j in range(1, len(cab) + 1):
                ws.cell(row=lin, column=j).font = NEGRITO
                if not ws.cell(row=lin, column=j).fill.fgColor.rgb == EST.fgColor.rgb:
                    ws.cell(row=lin, column=j).fill = SUB
        lin += 1
    return lin + 2


def larguras(ws, nk):
    for j in range(1, nk + 1):
        ws.column_dimensions[get_column_letter(j)].width = 18
    for j in range(nk + 1, nk + len(MESES) + len(ESTATS) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 15


def pct(delta, mr):
    return (delta / mr.where(mr.abs() > 1e-9)).fillna(0.0)


wb = Workbook()

# ------------------------------------------------------------------ 0_Resumo
ws = wb.active
ws.title = "0_Resumo"
ws.cell(row=1, column=1, value=f"{ROTULO_PXQ} x Mercado Receita - consolidado {ROTULO[INI]} a {ROTULO[FIM]}").font = Font(bold=True, size=14)
ws.cell(row=2, column=1, value=f"Delta = {ROTULO_PXQ} - MR. Acumulado pequeno com Desvio alto = diferenças que se compensam entre meses; "
                               "CV baixo e todos os meses com o mesmo sinal = diferença estável.")
brl = base[base["UNIDADE"] == "BRL"]
p_tot = pivota(brl, ["MEDIDA"])
ordem_med = ["Monetário total"] + FAMILIAS
p_tot = p_tot.reindex(ordem_med)
lin = 4
lin = escreve_bloco(ws, lin, "Delta R$ por família (todas as opções)", ordem_med, ["Família"],
                    p_tot["DELTA"], FMT_BRL, negrito=lambda k: k[0] == "Monetário total")
lin = escreve_bloco(ws, lin, "Delta % por família (delta / MR)", ordem_med, ["Família"],
                    pct(p_tot["DELTA"], p_tot["MR"]), FMT_PCT, negrito=lambda k: k[0] == "Monetário total")
q = base[base["UNIDADE"] == "QTD"]
p_q = pivota(q, ["MEDIDA"]).reindex(QTDS)
lin = escreve_bloco(ws, lin, "Delta de quantidade (todas as opções)", QTDS, ["Medida"], p_q["DELTA"], FMT_BRL)
lin = escreve_bloco(ws, lin, "Delta % de quantidade", QTDS, ["Medida"], pct(p_q["DELTA"], p_q["MR"]), FMT_PCT)
larguras(ws, 1)
ws.column_dimensions["A"].width = 22

# ------------------------------------------------------------------ 1 and 2: OPCAO x family
def por_opcao_familia(valor):
    """Rows OPCAO x (Monetário total + families) plus TOTAL; `valor` = 'DELTA' or '%'."""
    p = pivota(brl, ["OPCAO", "MEDIDA"])
    t = pivota(brl, ["MEDIDA"])
    t.index = pd.MultiIndex.from_tuples([("TOTAL", m) for m in t.index], names=["OPCAO", "MEDIDA"])
    p = pd.concat([p, t])
    chaves = [(o, m) for o in ORDEM_OPCAO + ["TOTAL"] for m in ordem_med if (o, m) in p.index]
    p = p.reindex(chaves)
    v = p["DELTA"] if valor == "DELTA" else pct(p["DELTA"], p["MR"])
    return chaves, v


for nome, valor, fmt, titulo in [("1_Opcao_Familia_R$", "DELTA", FMT_BRL, "Delta R$ por opção e família"),
                                 ("2_Opcao_Familia_%", "%", FMT_PCT, "Delta % por opção e família (delta / MR)")]:
    ws = wb.create_sheet(nome)
    chaves, v = por_opcao_familia(valor)
    escreve_bloco(ws, 1, titulo, chaves, ["Opção", "Família"], v, fmt,
                  negrito=lambda k: k[1] == "Monetário total" or k[0] == "TOTAL")
    larguras(ws, 2)

# ------------------------------------------------------------------ 3: quantities by OPCAO
ws = wb.create_sheet("3_Quantidades")
pq = pivota(q, ["OPCAO", "MEDIDA"])
chaves = [(o, m) for o in ORDEM_OPCAO for m in QTDS if (o, m) in pq.index]
pq = pq.reindex(chaves)
lin = escreve_bloco(ws, 1, "Delta de quantidade por opção", chaves, ["Opção", "Medida"], pq["DELTA"], FMT_BRL)
escreve_bloco(ws, lin, "Delta % de quantidade por opção (delta / MR)", chaves, ["Opção", "Medida"],
              pct(pq["DELTA"], pq["MR"]), FMT_PCT)
larguras(ws, 2)

# ------------------------------------------------------------------ 4: subgroup detail
ws = wb.create_sheet("4_Detalhe_Subgrupo_R$")
pd_ = pivota(brl, ["OPCAO", "SUBGRUPO", "MEDIDA"])
pd_ = pd_[(pd_["PXQ"].abs().sum(axis=1) + pd_["MR"].abs().sum(axis=1)) > 0.5]
chaves = [k for o in ORDEM_OPCAO for k in pd_.index if k[0] == o]
chaves = sorted(chaves, key=lambda k: (ORDEM_OPCAO.index(k[0]), k[1], ordem_med.index(k[2])))
escreve_bloco(ws, 1, "Delta R$ por opção, subgrupo e família", chaves, ["Opção", "Subgrupo", "Família"],
              pd_.reindex(chaves)["DELTA"], FMT_BRL, negrito=lambda k: k[2] == "Monetário total")
larguras(ws, 3)
ws.freeze_panes = "D3"

# ------------------------------------------------------------------ 5: long base for pivots
ws = wb.create_sheet("5_Base_longa")
cols = ["REF", "MEDIDA", "UNIDADE", "OPCAO", "SUBGRUPO", "PXQ", "MR", "DELTA"]
nomes = ["Mês", "Medida", "Unidade", "Opção", "Subgrupo", ROTULO_PXQ, "MR", "Delta"]
for j, t in enumerate(nomes, start=1):
    c = ws.cell(row=1, column=j, value=t)
    c.fill, c.font = CAB, CAB_F
for i, r in enumerate(ordena_opcao(base[cols]).itertuples(index=False), start=2):
    for j, v in enumerate(r, start=1):
        c = ws.cell(row=i, column=j, value=v)
        if j >= 6:
            c.number_format = FMT_BRL
for j in range(1, 9):
    ws.column_dimensions[get_column_letter(j)].width = 18
ws.auto_filter.ref = f"A1:H{len(base) + 1}"
ws.freeze_panes = "A2"

for w in wb.worksheets[:4]:
    w.freeze_panes = w.freeze_panes or "B3"
wb.save(SAIDA)

# ------------------------------------------------------------------ console summary
print(f"consolidado {ROTULO_PXQ} x MR | {', '.join(ROTULO[m] for m in MESES)}")
t = p_tot["DELTA"]
for med in ordem_med:
    linha = " | ".join(f"{t.loc[med, m] / 1e6:8.2f}" for m in MESES)
    print(f"  {med:<16} R$ mi: {linha} | acumulado {t.loc[med].sum() / 1e6:8.2f}")
print(f"saida: {SAIDA}")
