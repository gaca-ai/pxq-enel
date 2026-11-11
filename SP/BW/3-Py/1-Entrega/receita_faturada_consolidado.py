# -*- coding: utf-8 -*-
"""
Consolidado anual do layout Receita Faturada por familia, com os meses lado a lado e as linhas por
classe (estrutura da aba Parcela B do arquivo de explicacao ACT).

Uso:
    python receita_faturada_consolidado.py 2026
Le     1-Entrega/receita_faturada/receita_faturada_<ano>-<mm>.xlsx (meses que existirem)
Grava  1-Entrega/receita_faturada/consolidado_familias_<ano>.xlsx

Abas:
- Total: FAT + REFAT, monetario total, energia, tarifa media e a participacao de cada familia no
  monetario total, por linha (%).
- Uma aba por familia (FAT + REFAT): monetario total, energia, tarifa media e monetario detalhado
  (cada coluna ou grupo de colunas do layout que compoe a familia).
- Livre por classe (FAT + REFAT): o Mercado Livre aberto por classe, monetario total, energia, tarifa
  media e o monetario de cada familia.
- CNR: so o CNR (TOI), monetario total, energia, tarifa media e o monetario de cada familia.
- Conferencia: soma das familias x gravado por tipo de faturamento na aba Conferencia de cada layout.

Regras:
- Monetario: soma das linhas D e E de cada coluna do layout (toda a receita), em MR$.
- Energia: coluna CONSUMO das linhas E, em MWh; a demanda e a demanda de ultrapassagem ficam de fora.
  Excecao: Geracao e aberta em Geracao - Energia (linhas E) e Geracao - Demanda (*) (linhas D). Na
  energia, a linha Geracao - Demanda traz a demanda faturada (MW), so para a tarifa media dessa linha
  (R$/MW); ela nao soma na energia da Geracao nem no Total. A tarifa media da Geracao e do Total usa
  todo o monetario (D + E) e so a energia em MWh.
- Tarifa media: monetario / energia de cada linha, em R$/MWh (nao calculada para Bandeira e Tributos).
- Linhas: Total; Mercado Cativo e as classes do cativo (SCEE - GD1/GD2 somados na classe); Mercado
  Livre (so o total; aberto por classe na aba Livre por classe); Mercado de Suprimento; Mercado de Uso
  Distribuicao; Geracao (Energia e Demanda).
- Colunas: jan a dez do ano e YTD; meses sem layout ficam em branco. Totais, YTD, tarifa media e
  participacoes sao formulas.
"""
import os
import sys
import datetime as dt
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

BASE = Path(os.environ.get("PXQ_BASE", r"C:\python_projects\claude\p_x_q\SP\BW"))
DIR_RF = BASE / "1-Entrega" / "receita_faturada"
ANO = int(sys.argv[1]) if len(sys.argv) > 1 else max(int(p.stem[-7:-3]) for p in DIR_RF.glob("receita_faturada_????-??.xlsx"))
SAIDA = DIR_RF / f"consolidado_familias_{ANO}.xlsx"
TIPOS_BASE, TIPOS_CNR = ["FAT", "REFAT"], ["CNR"]

# family -> list of (detail label, layout columns summed in it); tariff average flag
FAMILIAS = [
    ("Parcela B", [("Parcela B", ["Parcela B"])], True),
    ("RI", [("RI", ["RI"])], True),
    ("Compra de Energia", [("Compra de Energia", ["Compra de Energia"])], True),
    ("Perdas", [("Perdas", ["Perdas"])], True),
    ("Transmissão", [("Rede Básica", ["Rede Básica"]), ("Transporte Itaipu", ["Transporte Itaipu"]),
                     ("CUSD", ["CUSD"]), ("Conexão T", ["Conexão T"])], True),
    ("Encargos", [("CDE", ["CDE USO", "CDE CONTAS", "CDE ELET"]), ("ESS/ERR/ERCAP", ["ESS/ERR/ERCAP"]),
                  ("Outros encargos (PROINFA, P&D, TFSEE, ONS)", ["PROINFA", "P&D", "TFSEE", "ONS"])], True),
    ("Financeiros", [("Base Financeira", ["Base Financeira"])], True),
    ("CVA", [("Base CVA", ["Base CVA"])], True),
    ("Bandeira", [("Bandeira", ["Bandeira"])], False),
    ("UD-ER e Outros", [("UD/ER", ["UD/ER"]), ("Outros", ["Outros"]), ("Covid Migrantes", ["Covid Migrantes"])], True),
    ("Tributos", [("PIS/COFINS", ["PIS/COFINS"]), ("ICMS", ["ICMS"])], False),
]
TODAS = [c for _, det, _ in FAMILIAS for _, cols in det for c in cols]
FAM_COLS = [(n, [c for _, cs in det for c in cs]) for n, det, _ in FAMILIAS]

CLASSES = [("Residencial", "Residencial"), ("Residencial Baixa Renda", "Residencial (Baixa renda)"),
           ("Industrial", "Industrial"), ("Comercial", "Comercial"), ("Rural", "Rural"),
           ("Poderes Públicos", "Poderes Públicos"), ("Iluminação Pública", "Iluminação Pública"),
           ("Serviços Públicos", "Serviços Públicos"), ("Consumo Próprio", "Próprio")]
GER, GER_E, GER_D = "Geração", "Geração – Energia", "Geração – Demanda (*)"
MERCADOS = [("LIVRE", "Mercado Livre"), ("SUPRIMENTO", "Mercado de Suprimento"),
            ("USO D", "Mercado de Uso Distribuição")]
OPCOES_CATIVO = {"CATIVO", "SCEE - GD1", "SCEE - GD2"}
MESES = [f"{ANO}-{m:02d}" for m in range(1, 13)]
OBS = ("(*) Geração – Demanda: na energia, a quantidade é a demanda faturada (MW), usada só na tarifa média "
       "dessa linha (R$/MW); não soma na energia da Geração nem no Total. A tarifa média da Geração e do Total "
       "usa todo o monetário da geração (D + E) e só a energia em MWh.")

# row layouts: (label, kind) with kind in total | grupo | item | ger | ger_e | ger_d
ROT_GERAL = ([("Total", "total"), ("Mercado Cativo", "grupo")] + [(r, "item") for _, r in CLASSES]
             + [(r, "mercado") for _, r in MERCADOS] + [(GER, "ger"), (GER_E, "ger_e"), (GER_D, "ger_d")])
ROT_LIVRE = [("Mercado Livre", "total")] + [(r, "item") for _, r in CLASSES if r != "Residencial (Baixa renda)"]


def linha(op, cl, un, livre):
    """Report row of a layout option x class x unit (livre=True: free market by class only)."""
    if livre:
        return dict(CLASSES).get(cl) if op == "LIVRE" else None
    if op in OPCOES_CATIVO:
        return dict(CLASSES).get(cl)
    if op == "GERAÇÃO":
        return GER_E if un == "E" else GER_D
    return dict(MERCADOS).get(op)


def le_mes(ref, tipos, livre=False):
    """Layout of month ref for the billing types: MR$ per layout column and quantity per report row."""
    arq = DIR_RF / f"receita_faturada_{ref}.xlsx"
    if not arq.exists():
        return None
    d = pd.read_excel(arq, sheet_name="Receita Faturada", header=2)
    d = d[d["OpcaoCompra"].notna() & d["TipoFaturamento"].isin(tipos)].copy()
    d["LINHA"] = [linha(o, c, u, livre) for o, c, u in zip(d["OpcaoCompra"], d["Classe"], d["Unidade"])]
    if livre:
        d = d[d["OpcaoCompra"] == "LIVRE"]
    sem = d[d["LINHA"].isna()]
    if len(sem):
        raise ValueError(f"{ref}: opcao x classe sem linha no relatorio: {sorted(set(zip(sem.OpcaoCompra, sem.Classe)))}")
    val = d.groupby("LINHA")[TODAS].sum() / 1000.0                  # R$ mil -> MR$
    # quantity: MWh of the E rows; for Geracao - Demanda, the billed demand (MW) of the D rows
    qtd = d[(d["Unidade"] == "E") | (d["LINHA"] == GER_D)].groupby("LINHA")["CONSUMO"].sum()
    val["QTD"] = qtd
    return val.fillna(0.0)


DADOS = {"base": {r: le_mes(r, TIPOS_BASE) for r in MESES},
         "cnr": {r: le_mes(r, TIPOS_CNR) for r in MESES},
         "livre": {r: le_mes(r, TIPOS_BASE, livre=True) for r in MESES}}
com_dado = [r for r in MESES if DADOS["base"][r] is not None]
if not com_dado:
    sys.exit(f"nenhum receita_faturada de {ANO} em {DIR_RF}")

# ------------------------------------------------ workbook
AZUL = PatternFill("solid", fgColor="1F4E79")
CINZA = PatternFill("solid", fgColor="D9E1F2")
FMT = {"MR$": '#,##0.00;-#,##0.00;"-"', "MWh": '#,##0;-#,##0;"-"', "R$/MWh": '#,##0.00;-#,##0.00;"-"',
       "%": '0.0%;-0.0%;"-"'}
C0, CY = 3, 3 + 12                                                # first month column (C) and YTD column (O)


def cabecalho(ws, r0, texto):
    ws.cell(r0, 2, texto).font = Font(bold=True, color="FFFFFF")
    for c in range(2, CY + 1):
        ws.cell(r0, c).fill = AZUL
    for k in range(12):
        cel = ws.cell(r0, C0 + k, dt.datetime(ANO, k + 1, 1))
        cel.number_format = "mmm/yyyy"
        cel.font = Font(bold=True, color="FFFFFF")
        cel.alignment = Alignment(horizontal="center")
    ws.cell(r0, CY, "YTD").font = Font(bold=True, color="FFFFFF")
    ws.cell(r0, CY).alignment = Alignment(horizontal="center")


def bloco(ws, r0, titulo, unidade, rotulos, dados, valor=None, razao=None):
    """One block (header + rows) at row r0. valor(frame, label) gives data rows; razao=(pos_num, pos_den,
    factor) builds ratio formulas (tariff average or share). Returns {label: row} and the next free row."""
    cabecalho(ws, r0, f"{titulo} ({unidade})")
    pos = {rot: r0 + 1 + i for i, (rot, _) in enumerate(rotulos)}
    tipo = dict(rotulos)
    for rot, r in pos.items():
        ws.cell(r, 2, rot)
        k = tipo[rot]
        if k in ("total", "grupo", "mercado", "ger"):
            ws.cell(r, 2).font = Font(bold=True)
        if k in ("item", "ger_e", "ger_d"):
            ws.cell(r, 2).alignment = Alignment(indent=2)
        if k == "total":
            for c in range(2, CY + 1):
                ws.cell(r, c).fill = CINZA
    itens = [rot for rot, k in rotulos if k == "item"]
    filhos_total = [rot for rot, k in rotulos if k in ("grupo", "mercado", "ger")]
    energia = unidade == "MWh"
    for k, ref in enumerate(MESES + [None]):
        c = CY if ref is None else C0 + k
        X = L(c)
        for rot, r in pos.items():
            kind = tipo[rot]
            if razao is not None:
                pn, pd_, fator = razao
                f = f'=IF(OR({X}{pd_[rot]}="",{X}{pd_[rot]}=0),"",{X}{pn[rot]}*{fator}/{X}{pd_[rot]})'
            elif ref is None:
                f = f"=SUM({L(C0)}{r}:{L(CY - 1)}{r})"
            elif dados[ref] is None:
                continue
            elif kind == "total" and filhos_total:
                f = "=" + "+".join(f"{X}{pos[x]}" for x in filhos_total)
            elif kind in ("total", "grupo"):
                f = f"=SUM({X}{pos[itens[0]]}:{X}{pos[itens[-1]]})"
            elif kind == "ger":                                      # energy: only the E rows (MWh)
                f = f"={X}{pos[GER_E]}" if energia else f"={X}{pos[GER_E]}+{X}{pos[GER_D]}"
            else:
                f = valor(dados[ref], rot)
            cel = ws.cell(r, c, f)
            cel.number_format = FMT["R$/MWh" if (razao is not None and unidade.startswith("R$")) else unidade]
            if kind == "total" or ref is None:
                cel.font = Font(bold=True)
    return pos, r0 + len(rotulos) + 3


def v_cols(cols):
    return lambda fr, rot: float(fr[cols].sum(axis=1).get(rot, 0.0))


def v_qtd(fr, rot):
    return float(fr["QTD"].get(rot, 0.0))


def prepara(ws, titulo, rotulos):
    ws.cell(1, 2, titulo).font = Font(bold=True, size=12)
    ws.cell(2, 2, "MR$ = milhões de reais; energia em MWh (só linhas de energia); tarifa média em R$/MWh").font = \
        Font(italic=True, size=9, color="808080")
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 36
    for c in range(C0, CY + 1):
        ws.column_dimensions[L(c)].width = 13


def topo(ws, rotulos, dados, cols, com_tarifa, nome):
    """Monetario total, Energia and Tarifa media blocks; returns positions and next row."""
    pos_m, r = bloco(ws, 4, f"Monetário total – {nome}", "MR$", rotulos, dados, valor=v_cols(cols))
    pos_e, r = bloco(ws, r, "Energia", "MWh", rotulos, dados, valor=v_qtd)
    if com_tarifa:
        _, r = bloco(ws, r, f"Tarifa média – {nome}", "R$/MWh", rotulos, dados, razao=(pos_m, pos_e, 1000000))
    return pos_m, r


def obs(ws, r, rotulos):
    if any(k == "ger_d" for _, k in rotulos):
        ws.cell(r, 2, OBS).font = Font(italic=True, size=9, color="808080")
        return r + 2
    return r


wb = openpyxl.Workbook()
wb.remove(wb.active)

# Total (FAT + REFAT): top blocks + share of each family in the total monetary value, by row
ws_t = wb.create_sheet("Total")
prepara(ws_t, "Total – receita faturada (FAT + REFAT, sem CNR)", ROT_GERAL)
pos_tot, r = topo(ws_t, ROT_GERAL, DADOS["base"], TODAS, True, "Total")
abas_fam = {}
for nome, det, tar in FAMILIAS:
    ws = wb.create_sheet(nome[:31])
    abas_fam[nome] = ws
    prepara(ws, f"{nome} – receita faturada (FAT + REFAT, sem CNR)", ROT_GERAL)
    cols = [c for _, cs in det for c in cs]
    pos_f, rf = topo(ws, ROT_GERAL, DADOS["base"], cols, tar, nome)
    if len(det) > 1 or det[0][0] != nome:
        for rot, cs in det:
            extra = "" if cs == [rot] else f" ({' + '.join(cs)})"
            _, rf = bloco(ws, rf, f"Monetário detalhado – {rot}{extra}", "MR$", ROT_GERAL, DADOS["base"], valor=v_cols(cs))
    obs(ws, rf, ROT_GERAL)
    # share of the family in the total monetary value (formulas pointing to the family sheet)
    cabecalho(ws_t, r, f"Participação de {nome} no monetário total (%)")
    for rot, rr in pos_tot.items():
        lin = r + 1 + list(pos_tot).index(rot)
        ws_t.cell(lin, 2, rot)
        if dict(ROT_GERAL)[rot] in ("item", "ger_e", "ger_d"):
            ws_t.cell(lin, 2).alignment = Alignment(indent=2)
        else:
            ws_t.cell(lin, 2).font = Font(bold=True)
        for c in range(C0, CY + 1):
            X = L(c)
            ref_f = f"'{ws.title}'!{X}{pos_f[rot]}"
            cel = ws_t.cell(lin, c, f'=IF(OR({X}{rr}="",{X}{rr}=0),"",{ref_f}/{X}{rr})')
            cel.number_format = FMT["%"]
    r += len(ROT_GERAL) + 3
obs(ws_t, r, ROT_GERAL)

# Livre por classe (FAT + REFAT)
ws_l = wb.create_sheet("Livre por classe")
prepara(ws_l, "Mercado Livre por classe – receita faturada (FAT + REFAT, sem CNR)", ROT_LIVRE)
_, r = topo(ws_l, ROT_LIVRE, DADOS["livre"], TODAS, True, "Mercado Livre")
for nome, cols in FAM_COLS:
    _, r = bloco(ws_l, r, f"Monetário – {nome}", "MR$", ROT_LIVRE, DADOS["livre"], valor=v_cols(cols))

# CNR: same layout as the Total of FAT + REFAT, with the monetary value of each family below
ws_c = wb.create_sheet("CNR")
prepara(ws_c, "CNR (TOI) – receita faturada", ROT_GERAL)
_, r = topo(ws_c, ROT_GERAL, DADOS["cnr"], TODAS, True, "CNR")
for nome, cols in FAM_COLS:
    _, r = bloco(ws_c, r, f"Monetário – {nome}", "MR$", ROT_GERAL, DADOS["cnr"], valor=v_cols(cols))
obs(ws_c, r, ROT_GERAL)

# reconciliation: families x Conferencia of each month, by billing type
ws = wb.create_sheet("Conferencia")
ws.append(["Mês", "FAT + REFAT: soma das famílias (MR$)", "FAT + REFAT: gravado no layout (MR$)", "Diferença (MR$)",
           "CNR: soma das famílias (MR$)", "CNR: gravado no layout (MR$)", "Diferença (MR$)"])
for ref in com_dado:
    conf = pd.read_excel(DIR_RF / f"receita_faturada_{ref}.xlsx", sheet_name="Conferencia", header=None)
    lin = [ref]
    for chave, tipos in (("base", TIPOS_BASE), ("cnr", TIPOS_CNR)):
        grav = sum(float(conf.loc[conf[0].eq(t), 1].iloc[0]) for t in tipos) / 1e6
        soma = float(DADOS[chave][ref][TODAS].sum().sum())
        lin += [soma, grav, soma - grav]
        if abs(soma - grav) > 0.001:
            print(f"ATENCAO {ref} {chave}: familias {soma:,.3f} x Conferencia {grav:,.3f} MR$")
    ws.append(lin)
for row in ws.iter_rows(min_row=2):
    for cel in row[1:]:
        cel.number_format = "#,##0.000"
for c in ws[1]:
    c.font = Font(bold=True)
    c.alignment = Alignment(wrap_text=True)
ws.column_dimensions["A"].width = 10
for x in "BCDEFG":
    ws.column_dimensions[x].width = 22
ws.append([])
ws.append(["Meses sem arquivo receita_faturada ficam em branco nas demais abas."])

wb.save(SAIDA)
print(f"meses: {', '.join(com_dado)} | saida: {SAIDA}")
