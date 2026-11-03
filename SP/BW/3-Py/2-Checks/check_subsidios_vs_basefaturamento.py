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
Writes 2-Checks/check_subsidios_vs_basefaturamento_<ini>_a_<fim>.xlsx (MR$), in the layout requested by the
       area (model "Comparativo"): sheets Total, Tarifa Social and Fonte incentivada, each with the blocks
       Base Faturamento, PxQ and Delta, months Jan-Dec + YTD (B2 = last closed month), live formulas

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

import openpyxl
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


GRUPOS = [("CATIVO", "Cativo"), ("LIVRE", "Livre")]
# class order and labels requested by the area (Residencial Baixa Renda right after Residencial)
ORDEM_CLASSE = ["Residencial", "Residencial Baixa Renda", "Industrial", "Comercial", "Rural", "Poder Público",
                "Iluminação Pública", "Serviço Público", "Consumo Próprio", "Suprimento", "Uso Distribuição",
                "Uso Geração"]
ROTULO_CLASSE = {"Poderes Públicos": "Poder Público", "Serviços Públicos": "Serviço Público"}
ABAS = [("Total", ["Tarifa Social", "Fonte incentivada"]), ("Tarifa Social", ["Tarifa Social"]),
        ("Fonte incentivada", ["Fonte incentivada"])]
MIN_MR = 0.0005                    # a class row is kept when some month has at least R$ 500 on either side
FMT = '#,##0.00;-#,##0.00;"-"'

rel = {ref: le_relatorio(ref) for ref in REFS}
bl = {ref: le_bl(ref) for ref in REFS}
ok = [ref for ref in REFS if rel[ref] is not None and bl[ref] is not None]   # months with both sides
ultimo = ok[-1] if ok else REFS[0]
ANO = int(REFS[0][:4])
MESES = [f"{ANO}-{m:02d}" for m in range(1, 13)]                  # calendar-year columns, as in the model


def valores(lado, itens):
    """{month: Series indexed by (type, class)} of one side for the given items (None = no file)."""
    out = {}
    for ref in MESES:
        s = lado.get(ref)
        if s is None:
            out[ref] = None
            continue
        s = s[s.index.get_level_values(0).isin(itens)]
        s = s.rename(index=ROTULO_CLASSE, level=2)
        out[ref] = s.groupby(level=[1, 2]).sum()
    return out


def linhas_aba(base, pxq):
    """Class rows with value in some month on either side, grouped by customer type, in the layout order."""
    vistos = {}
    for lado in (base, pxq):
        for s in lado.values():
            if s is None:
                continue
            for k, v in s.items():
                vistos[k] = max(vistos.get(k, 0.0), abs(v))
    grupos = []
    for tipo, rotulo in GRUPOS:
        cls = [c for (t, c), m in vistos.items() if t == tipo and m >= MIN_MR]
        cls.sort(key=lambda c: (ORDEM_CLASSE.index(c) if c in ORDEM_CLASSE else len(ORDEM_CLASSE), c))
        if cls:
            grupos.append((tipo, rotulo, cls))
    outros = sorted({t for (t, _), m in vistos.items() if m >= MIN_MR} - {t for t, _ in GRUPOS})
    if outros:
        print(f"ATENCAO: tipos de cliente fora de Cativo/Livre ignorados: {outros}")
    return grupos


def escreve_aba(ws, base, pxq):
    """Model layout: B2 = last closed month; blocks Base Faturamento, PxQ and Delta (Base - PxQ)."""
    from openpyxl.utils import get_column_letter as L
    grupos = linhas_aba(base, pxq)
    col = {ref: 3 + k for k, ref in enumerate(MESES)}                # C..N
    c_ytd = 15                                                       # O
    ws["B2"] = pd.Timestamp(f"{ultimo}-01").to_pydatetime()
    ws["B2"].number_format = "dd/mm/yyyy"
    ws["B2"].font = Font(bold=True)
    for ref, c in col.items():
        ws.cell(3, c, f'=IF({L(c)}4>$B$2,"YTG","YTD")').alignment = Alignment(horizontal="center")

    def bloco(r0, titulo, lado=None, ref_bloco=None):
        """Header at row r0; returns {row key: row number} and the last row."""
        ws.cell(r0, 2, titulo).font = Font(bold=True)
        for ref, c in col.items():
            cel = ws.cell(r0, c, pd.Timestamp(f"{MESES[0]}-01").to_pydatetime() if c == 3 else f"=EDATE({L(c - 1)}{r0},1)")
            cel.number_format = "mmm/yy"
            cel.font = Font(bold=True)
            cel.alignment = Alignment(horizontal="center")
        ws.cell(r0, c_ytd, "YTD").font = Font(bold=True)
        pos, r = {}, r0 + 1
        pos["TOTAL"] = r
        r += 1
        for tipo, rotulo, cls in grupos:
            pos[tipo] = r
            r += 1
            for c_ in cls:
                pos[(tipo, c_)] = r
                r += 1
        ultima = r - 1
        # labels
        ws.cell(pos["TOTAL"], 2, "Total").font = Font(bold=True)
        for tipo, rotulo, cls in grupos:
            ws.cell(pos[tipo], 2, rotulo).font = Font(bold=True)
            for c_ in cls:
                ws.cell(pos[(tipo, c_)], 2, c_)
        # values / formulas
        for ref, c in list(col.items()) + [(None, c_ytd)]:
            X = L(c)
            tem = ref is None or (lado is not None and lado[ref] is not None) or (
                ref_bloco is not None and ref in ok)
            if not tem:
                continue
            for tipo, rotulo, cls in grupos:
                for c_ in cls:
                    rr = pos[(tipo, c_)]
                    if ref_bloco is not None:                         # Delta block
                        f = f"={X}{ref_bloco[0][(tipo, c_)]}-{X}{ref_bloco[1][(tipo, c_)]}"
                    elif ref is None:
                        f = f'=SUMIF({L(3)}$3:{L(14)}$3,"YTD",{L(3)}{rr}:{L(14)}{rr})'
                    else:
                        f = float(lado[ref].get((tipo, c_), 0.0))
                    ws.cell(rr, c, f)
                ini, fim = pos[(tipo, cls[0])], pos[(tipo, cls[-1])]
                ws.cell(pos[tipo], c, f"=SUM({X}{ini}:{X}{fim})" if ref_bloco is None
                        else f"={X}{ref_bloco[0][tipo]}-{X}{ref_bloco[1][tipo]}").font = Font(bold=True)
            ws.cell(pos["TOTAL"], c, ("=SUM(" + ",".join(f"{X}{pos[t]}" for t, _, _ in grupos) + ")") if ref_bloco is None
                    else f"={X}{ref_bloco[0]['TOTAL']}-{X}{ref_bloco[1]['TOTAL']}").font = Font(bold=True)
        for row in ws.iter_rows(min_row=r0 + 1, max_row=ultima, min_col=3, max_col=c_ytd):
            for cel in row:
                cel.number_format = FMT
        return pos, ultima

    pos_b, fim_b = bloco(4, "Base Faturamento", lado=base)
    ws.cell(fim_b + 3, 2, "* Comparativo dos subsídios concedidos (MR$)").font = Font(italic=True)
    r_flag = fim_b + 5
    for ref, c in col.items():
        ws.cell(r_flag, c, f"={L(c)}$3").alignment = Alignment(horizontal="center")
    pos_p, fim_p = bloco(r_flag + 1, "PxQ", lado=pxq)
    bloco(fim_p + 6, "Delta", ref_bloco=(pos_b, pos_p))
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 26
    for c in range(3, c_ytd + 1):
        ws.column_dimensions[L(c)].width = 11


LEIA_ME = [
    "Check dos subsídios concedidos: relatório do sistema de faturamento (Base Faturamento) x Bl1 x Bl2 (PxQ)",
    "",
    "Valores em MR$ (milhões de reais). Positivo = desconto concedido. Delta = Base Faturamento - PxQ.",
    "Abas: Total (Tarifa Social + fonte incentivada), Tarifa Social e Fonte incentivada.",
    "Tarifa Social = ZREDU4 (relatório) x Tarifa Social - DMR (Bl1 x Bl2);",
    "fonte incentivada = ZTUS98 + ZTUS99 (relatório, sinal invertido) x fonte incentivada (Bl1 x Bl2).",
    "Tipo de cliente: Cliente Cativo = Cativo; Cliente Livre e Cliente Híbrido = Livre.",
    "Classe no padrão do layout Receita Faturada (Baixa Renda, BPC, Indígena, Quilombola e Desconto Social em",
    "Residencial Baixa Renda). Só aparecem as classes com valor em algum mês.",
    "",
    "B2 = último mês fechado (último mês com os dois lados). Meses até B2 = YTD; seguintes = YTG.",
    "A coluna YTD soma só os meses YTD. Mês sem arquivo de um dos lados fica em branco; o Delta só é",
    "calculado nos meses com os dois lados. B2 pode ser alterado à mão.",
    "",
    "SEM ABERTURA POR TIPO DE FATURAMENTO: os dois lados somam faturamento normal (FAT), refaturamento (REFAT)",
    "e TOI (CNR) do mês contábil. Não há separação entre FAT, REFAT e TOI nesta planilha.",
    "O subsídio só é calculado para o TOI da Baixa Renda (Tarifa Social); a fonte incentivada não se aplica",
    "ao TOI, nos dois lados.",
    "",
    "A Tarifa Social do Bl1 x Bl2 é a DMR do informe de gestão, que não abre por classe: fica toda em",
    "Residencial Baixa Renda. A abertura por classe só traz informação nova na fonte incentivada.",
    "",
    f"Meses com os dois lados: {', '.join(ok) if ok else 'nenhum'}",
    "Fontes: src/MM-AAAA_ENEL SP_RECEITA_com BAIXA RENDA E FONTE INCENTIVADA.csv e "
    "2-Checks/receita_faturada_Bl1_vs_Bl2/receita_faturada_Bl1_vs_Bl2_<ref>.xlsx",
    "Gerado por 3-Py/2-Checks/check_subsidios_vs_basefaturamento.py",
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Leia-me"
for linha in LEIA_ME:
    ws.append([linha])
ws["A1"].font = Font(bold=True)
ws.column_dimensions["A"].width = 115
resumo = []
for nome, itens in ABAS:
    base, pxq = valores(rel, itens), valores(bl, itens)
    escreve_aba(wb.create_sheet(nome), base, pxq)
    b = sum(base[r].sum() for r in ok)
    x = sum(pxq[r].sum() for r in ok)
    resumo.append((nome, b, x, b - x))
wb.save(SAIDA)

print(f"YTD ate {ultimo} (MR$):")
for nome, b, x, d in resumo:
    print(f"  {nome:18s} Base Faturamento {b:8.2f} | PxQ {x:8.2f} | Delta {d:6.2f}")
print(f"saida: {SAIDA}")
