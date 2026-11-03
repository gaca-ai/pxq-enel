# -*- coding: utf-8 -*-
"""
Check of the subsidies granted, month by month, between the billing system report (subsidy items of the
BW billing base) and the homologated x realized tariff check (Bl1 x Bl2), by item, customer type and class.

Items:
  Tarifa Social       report: ZREDU4            | Bl1 x Bl2: sheet desconto_tarifa_social_dmr
  Fonte incentivada   report: ZTUS98 + ZTUS99   | Bl1 x Bl2: sheet desconto_fonte_incentivada
Positive = discount granted (ZREDU4 comes positive in the report; ZTUS98/99 negative, sign inverted).

Usage:
    python check_subsidios_vs_basefaturamento.py 2026-01 2026-08
Reads  src/MM-AAAA_ENEL SP_RECEITA_com BAIXA RENDA E FONTE INCENTIVADA.csv (billing team)
       2-Checks/receita_faturada_Bl1_vs_Bl2/receita_faturada_Bl1_vs_Bl2_<ref>.xlsx
Writes 2-Checks/check_subsidios_vs_basefaturamento_<ini>_a_<fim>.xlsx (MR$)

Rules:
- Report: sum of "Montante variável"; subtotal rows ("Resultado", "Result.global") are dropped.
- Customer type: Cliente Cativo -> CATIVO; Cliente Livre and Cliente Hibrido -> LIVRE (as in the layout).
- Class in the layout standard: Residencial with Baixa Renda, BPC, Indigena, Quilombola or Desconto Social
  -> Residencial Baixa Renda; Poder Publico -> Poderes Publicos; Servico Publico -> Servicos Publicos.
- No split by billing type: both sides sum normal billing, rebilling and TOI (CNR) of the accounting month.
- Bl1 x Bl2: tariff columns of the layout (Parcela B to Base CVA), D and E rows summed. The Tarifa Social of
  the Bl1 x Bl2 is the DMR of the informe de gestao, which has no class detail (all in Residencial Baixa Renda).
- Month without Bl1 x Bl2 file: that side stays blank, with a warning.
"""
import os
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl.styles import Alignment, Font

BASE = Path(os.environ.get("PXQ_BASE", r"C:\python_projects\claude\p_x_q\SP\BW"))
INI, FIM = (sys.argv[1], sys.argv[2]) if len(sys.argv) > 2 else ("2026-01", "2026-08")
REFS = [p.strftime("%Y-%m") for p in pd.period_range(INI, FIM, freq="M")]
DIR_BL = BASE / "2-Checks" / "receita_faturada_Bl1_vs_Bl2"
SAIDA = BASE / "2-Checks" / f"check_subsidios_vs_basefaturamento_{INI}_a_{FIM}.xlsx"
MR = 1e6                                                       # R$ -> MR$

ITEM_REL = {"ZREDU4": ("Tarifa Social", 1.0), "ZTUS98": ("Fonte incentivada", -1.0),
            "ZTUS99": ("Fonte incentivada", -1.0)}
ABA_BL = {"Tarifa Social": "desconto_tarifa_social_dmr", "Fonte incentivada": "desconto_fonte_incentivada"}
TIPO = {"Cliente Cativo": "CATIVO", "Cliente Livre": "LIVRE", "Cliente Híbrido": "LIVRE"}
CLASSE = {"Residencial": "Residencial", "Industrial": "Industrial", "Comercial": "Comercial", "Rural": "Rural",
          "Poder Público": "Poderes Públicos", "Iluminação Pública": "Iluminação Pública",
          "Serviço Público": "Serviços Públicos", "Consumo Próprio": "Consumo Próprio"}
BAIXA_RENDA = ("Baixa Renda", "BPC", "Indígena", "Quilombola", "Desconto Social")
CHAVE = ["Item", "Tipo de cliente", "Classe"]


def classe(texto):
    """Layout class from the report's 'Classe/Sub-classe' (e.g. 'Residencial-Baixa Renda BPC')."""
    principal, _, sub = str(texto).partition("-")
    if principal == "Residencial" and any(b in sub for b in BAIXA_RENDA):
        return "Residencial Baixa Renda"
    return CLASSE.get(principal, principal)


def num_br(s):
    """Brazilian number text ('1.234,56') -> float."""
    return pd.to_numeric(s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")


def le_relatorio(ref):
    """Subsidies of the billing report for month ref, MR$ by item x type x class (None if no file)."""
    arq = BASE / "src" / f"{ref[5:]}-{ref[:4]}_ENEL SP_RECEITA_com BAIXA RENDA E FONTE INCENTIVADA.csv"
    if not arq.exists():
        print(f"ATENCAO {ref}: relatorio de subsidios ausente ({arq.name})")
        return None
    cab = pd.read_csv(arq, sep=";", header=None, nrows=4, dtype=str, encoding="utf-8")
    # fixed positions of the BW export, checked against the header rows
    assert cab.iloc[3, 0] == "Tipo item documento" and cab.iloc[3, 2] == "Classe/Sub-classe", arq.name
    assert cab.iloc[3, 6] == "Flag Cli.Liv./ Cat." and cab.iloc[1, cab.shape[1] - 1] == "Montante variável", arq.name
    d = pd.read_csv(arq, sep=";", header=None, skiprows=4, dtype=str, encoding="utf-8")
    d = d[(d[6] != "Resultado") & (d[0] != "Result.global")]   # subtotal rows of the export
    desconhecido = set(d[0]) - set(ITEM_REL)
    if desconhecido:
        print(f"ATENCAO {ref}: itens nao previstos no relatorio, ignorados: {sorted(desconhecido)}")
        d = d[d[0].isin(ITEM_REL)]
    out = pd.DataFrame({
        "Item": d[0].map(lambda x: ITEM_REL[x][0]),
        "Tipo de cliente": d[6].map(TIPO).fillna(d[6]),
        "Classe": d[2].map(classe),
        "v": num_br(d[d.columns[-1]]).fillna(0.0) * d[0].map(lambda x: ITEM_REL[x][1]) / MR,
    })
    return out.groupby(CHAVE)["v"].sum()


def le_bl(ref):
    """Tarifa Social (DMR) and fonte incentivada of the Bl1 x Bl2 file, MR$ by item x type x class."""
    arq = DIR_BL / f"receita_faturada_Bl1_vs_Bl2_{ref}.xlsx"
    if not arq.exists():
        print(f"ATENCAO {ref}: Bl1 x Bl2 ausente ({arq.name}); rodar gera_receita_faturada_bl1_vs_bl2.py")
        return None
    partes = []
    for item, aba in ABA_BL.items():
        d = pd.read_excel(arq, sheet_name=aba, header=2)
        d = d[d["OpcaoCompra"].notna()]
        cols = list(d.columns[list(d.columns).index("Parcela B"):list(d.columns).index("Base CVA") + 1])
        v = d[cols].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1) / 1000.0   # R$ mil -> MR$
        partes.append(pd.DataFrame({"Item": item, "Tipo de cliente": d["OpcaoCompra"], "Classe": d["Classe"], "v": v}))
    return pd.concat(partes).groupby(CHAVE)["v"].sum()


rel = {ref: le_relatorio(ref) for ref in REFS}
bl = {ref: le_bl(ref) for ref in REFS}
indice = sorted(set().union(*[s.index for s in list(rel.values()) + list(bl.values()) if s is not None]))
# keep only keys with value on some side
cols, delta = {}, {}
for ref in REFS:
    r = rel[ref].reindex(indice) if rel[ref] is not None else pd.Series(np.nan, index=pd.MultiIndex.from_tuples(indice))
    b = bl[ref].reindex(indice) if bl[ref] is not None else pd.Series(np.nan, index=pd.MultiIndex.from_tuples(indice))
    if rel[ref] is not None:
        r = r.fillna(0.0)
    if bl[ref] is not None:
        b = b.fillna(0.0)
    cols[f"{ref} Relatório"], cols[f"{ref} Bl1×Bl2"], cols[f"{ref} Δ"] = r.values, b.values, (r - b).values
    delta[ref] = (r - b).values
comp = pd.DataFrame(cols, index=pd.MultiIndex.from_tuples(indice, names=CHAVE))
comp = comp[comp.abs().sum(axis=1) > 0.0005]
ok = [ref for ref in REFS if rel[ref] is not None and bl[ref] is not None]   # months with both sides
comp["Acumulado Relatório"] = comp[[f"{r} Relatório" for r in ok]].sum(axis=1)
comp["Acumulado Bl1×Bl2"] = comp[[f"{r} Bl1×Bl2" for r in ok]].sum(axis=1)
comp["Acumulado Δ"] = comp["Acumulado Relatório"] - comp["Acumulado Bl1×Bl2"]
comp["Acumulado Δ %"] = (comp["Acumulado Δ"] / comp["Acumulado Bl1×Bl2"].abs()).where(comp["Acumulado Bl1×Bl2"].abs() > 1e-6)


def com_totais(df):
    """Subtotal per item and grand total rows."""
    blocos = []
    for item, g in df.groupby(level=0, sort=False):
        blocos.append(g)
        s = g.sum(min_count=1).to_frame().T
        s.index = pd.MultiIndex.from_tuples([(item, "Total", "")], names=CHAVE)
        blocos.append(s)
    t = df.sum(min_count=1).to_frame().T
    t.index = pd.MultiIndex.from_tuples([("TOTAL", "", "")], names=CHAVE)
    out = pd.concat(blocos + [t])
    if "Acumulado Δ %" in out:
        out["Acumulado Δ %"] = (out["Acumulado Δ"] / out["Acumulado Bl1×Bl2"].abs()).where(
            out["Acumulado Bl1×Bl2"].abs() > 1e-6)
    return out


comp_t = com_totais(comp)
dlt = comp[[f"{r} Δ" for r in REFS] + ["Acumulado Δ"]]
dlt.columns = REFS + ["Acumulado"]
dlt_t = com_totais(dlt)

LEIA_ME = [
    "Check dos subsídios concedidos: relatório do sistema de faturamento x Bl1 x Bl2 (tarifa homologada x realizada)",
    "",
    "Valores em MR$ (milhões de reais). Positivo = desconto concedido. Δ = relatório - Bl1 x Bl2.",
    "Itens: Tarifa Social = ZREDU4 (relatório) x Tarifa Social - DMR (Bl1 x Bl2);",
    "fonte incentivada = ZTUS98 + ZTUS99 (relatório, sinal invertido) x fonte incentivada (Bl1 x Bl2).",
    "Tipo de cliente: Cliente Cativo = CATIVO; Cliente Livre e Cliente Híbrido = LIVRE.",
    "Classe no padrão do layout Receita Faturada (Baixa Renda, BPC, Indígena, Quilombola e Desconto Social em",
    "Residencial Baixa Renda).",
    "",
    "SEM ABERTURA POR TIPO DE FATURAMENTO: os dois lados somam faturamento normal (FAT), refaturamento (REFAT)",
    "e TOI (CNR) do mês contábil. Não há separação entre FAT, REFAT e TOI nesta planilha.",
    "",
    "A Tarifa Social do Bl1 x Bl2 é a DMR do informe de gestão, que não abre por classe: fica toda em",
    "Residencial Baixa Renda. A abertura por classe só traz informação nova na fonte incentivada.",
    "Mês sem arquivo do Bl1 x Bl2 (ou do relatório): o lado ausente fica em branco e o mês sai do acumulado.",
    "",
    f"Meses no acumulado: {', '.join(ok) if ok else 'nenhum'}",
    "Fontes: src/MM-AAAA_ENEL SP_RECEITA_com BAIXA RENDA E FONTE INCENTIVADA.csv e "
    "2-Checks/receita_faturada_Bl1_vs_Bl2/receita_faturada_Bl1_vs_Bl2_<ref>.xlsx",
    "Gerado por 3-Py/2-Checks/check_subsidios_vs_basefaturamento.py",
]

with pd.ExcelWriter(SAIDA, engine="openpyxl") as w:
    pd.DataFrame({"Leia-me": LEIA_ME}).to_excel(w, sheet_name="Leia-me", index=False)
    comp_t.to_excel(w, sheet_name="Comparacao")
    dlt_t.to_excel(w, sheet_name="Delta_MR$")
    w.book["Leia-me"].column_dimensions["A"].width = 115
    for nome in ("Comparacao", "Delta_MR$"):
        ws = w.book[nome]
        for L, larg in zip("ABC", (18, 16, 26)):
            ws.column_dimensions[L].width = larg
        ultima = ws.max_column
        for c in ws[1]:
            c.font = Font(bold=True)
            c.alignment = Alignment(wrap_text=True, horizontal="center")
        for row in ws.iter_rows(min_row=2):
            negrito = row[1].value == "Total" or row[0].value == "TOTAL"
            for cel in row:
                if cel.column > 3:
                    pct = nome == "Comparacao" and cel.column == ultima
                    cel.number_format = "0.0%" if pct else "#,##0.00"
                if negrito:
                    cel.font = Font(bold=True)
        ws.freeze_panes = None

pd.set_option("display.width", 220)
print(dlt_t.round(2).to_string())
print(f"\nsaida: {SAIDA}")
