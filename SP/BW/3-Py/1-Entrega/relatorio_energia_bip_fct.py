# -*- coding: utf-8 -*-
"""
Relatorio de energia realizada para BIP/FCT: a tabela tabPxQ do modelo do orcamento (aba PXQ),
preenchida com o realizado do P x Q, sem vinculos externos.

Uso:
    python relatorio_energia_bip_fct.py 2026            # todos os meses com pxq e receita_faturada
    python relatorio_energia_bip_fct.py 2026 2026-10    # so os meses indicados
Le, por mes:
    1-Entrega/pxq/pxq_<ref>.xlsx                    grupos do pxq (bloco 2), pelo conversor do layout
    1-Entrega/pxq_regulacao/PxQ_regulacao_<ref>.xlsx camadas CVA e FIN de cada item
    1-Entrega/receita_faturada/receita_faturada_<ref>.xlsx    conferencia
    tbl/tarifas/tarifa_long.csv                     tarifas das faixas do Desconto Social
Grava
    1-Entrega/relatorio_energia_bip_fct/relatorio_energia_bip_fct_<ano>_agregado.xlsx
        Grandezas FATURADA_MWH e DEMANDA_KW (FAT + REFAT + TOI)
    1-Entrega/relatorio_energia_bip_fct/relatorio_energia_bip_fct_<ano>_por_tipo.xlsx
        Grandezas FAT_MWH, REFAT_MWH, TOI_MWH, DEMANDA_FAT_KW, DEMANDA_REFAT_KW, DEMANDA_TOI_KW
    Abas: PXQ (tabela tabPxQ, colunas do modelo), Conferencia e Leia-me.

Premissas (procedimento, "Relatorio de energia realizada para BIP/FCT: premissas"):
- Linhas tarifaveis do bloco 2 (entrega) do pxq, com a classificacao do layout Receita Faturada
  (gera_receita_faturada.converte, parametro detalhe). Nao tarifaveis (bandeira, tributos, UD/ER,
  Outros, Covid Migrantes) e o residuo sem item tarifario ficam fora.
- Volume: energia em MWh (linhas E) e demanda em kW (linhas D). REFAT e TOI no mes de faturamento.
- Receita em R$ mil, por familia: Fio A, Fio B, Encargos, Perdas, Energia e RI (camada BE, pelo de-para
  do gera_de_para_regulacao_MR.py), CVA (camada CVA); Financeiros = Total - familias (camada FIN e itens
  sem familia); Total = tarifa de aplicacao do bloco 2. Tarifa media = receita / volume x 1000.
- Sem abertura de SCEE (somado na classe), APE e fonte incentivada (somados no livre) e irrigacao.
- Faixas sociais pela estimativa do P x Q (% energia na faixa 01): Tarifa Social - faixa 01 com receita
  zero (isenta); Desconto Social repartido item a item pelas tarifas das faixas 01 e 02.
- <500 = Ceris (agente do SD).
"""
import os
import sys
import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

BASE = Path(os.environ.get("PXQ_BASE", r"C:\python_projects\claude\p_x_q\SP\BW"))
sys.path.insert(0, str(BASE / "3-Py" / "1-Entrega"))
sys.path.insert(0, str(BASE / "3-Py" / "2-Checks"))
_argv, sys.argv = sys.argv, sys.argv[:1]            # gera_de_para_regulacao_MR reads sys.argv on import
import gera_receita_faturada as grf                   # noqa: E402
from gera_de_para_regulacao_MR import ITENS as DE_PARA_MR   # noqa: E402
sys.argv = _argv

ANO = int(sys.argv[1]) if len(sys.argv) > 1 else dt.date.today().year
DIR_PXQ = BASE / "1-Entrega" / "pxq"
DIR_REG = BASE / "1-Entrega" / "pxq_regulacao"
DIR_RF = BASE / "1-Entrega" / "receita_faturada"
ARQ_TARIFA = BASE / "tbl" / "tarifas" / "tarifa_long.csv"
DIR_SAIDA = BASE / "1-Entrega" / "relatorio_energia_bip_fct"
NSA = "Não se aplica"
COL_F01 = "% energia na faixa 01 (estimada pela receita)"

FAMILIAS = ["Fio B", "RI", "Energia", "Perdas", "Fio A", "Encargos", "CVA"]      # template order
FAM_PXQ = {"FIO B": "Fio B", "RI": "RI", "ENERGIA": "Energia", "PERDAS": "Perdas", "FIO A": "Fio A",
           "ENCARGOS": "Encargos"}                                               # OUTROS -> Financeiros
NATUREZA = {"CATIVO": "CATIVO", "LIVRE": "LIVRE", "USO D": "DISTRIBUIÇÃO", "GERAÇÃO": "GERAÇÃO",
            "SUPRIMENTO": "SUPRIMENTO"}
CLASSE = {"Residencial": "RESIDENCIAL", "Residencial Baixa Renda": "RESIDENCIAL", "Industrial": "INDUSTRIAL",
          "Comercial": "COMERCIAL", "Rural": "RURAL", "Poderes Públicos": "PODER PÚBLICO",
          "Iluminação Pública": "ILUMINAÇÃO PÚBLICA", "Serviços Públicos": "SERVIÇO PÚBLICO",
          "Consumo Próprio": "CONSUMO PRÓPRIO", "SUPRIMENTO": "SUPRIMENTO", "USO D": "USO DISTRIBUIÇÃO",
          "GERAÇÃO": "USO GERAÇÃO"}
POSTO = {"P": "Ponta", "FP": "Fora ponta", "INT": "Intermediário"}
SUBGRUPOS = ["A3a", "A2", "A4", "AS", "B1", "B2", "B3", "B4"]
MODALIDADES = ["Convencional", "Distribuição", "Geração", "Branca", "Azul", "Verde"]
SUB_TS = ("Residencial Tarifa Social – faixa 01", "Residencial Tarifa Social – faixa 02")
SUB_DS = ("Residencial Desconto Social – faixa 01", "Residencial Desconto Social – faixa 02")
GRANDEZA = {"agregado": {("E", t): "FATURADA_MWH" for t in grf.TIPOS} | {("D", t): "DEMANDA_KW" for t in grf.TIPOS},
            "por_tipo": {("E", "FAT"): "FAT_MWH", ("E", "REFAT"): "REFAT_MWH", ("E", "CNR"): "TOI_MWH",
                         ("D", "FAT"): "DEMANDA_FAT_KW", ("D", "REFAT"): "DEMANDA_REFAT_KW",
                         ("D", "CNR"): "DEMANDA_TOI_KW"}}
CHAVES = ["Grandeza", "Natureza Mercado", "SUBGRUPO", "Modalidade", "Classe", "Subclasse", "Detalhe", "Posto",
          "Unidade", "Agente", "Mes_Ano"]
REC = ["Receita - Total"] + [f"Receita - {f}" for f in FAMILIAS] + ["Receita - Financeiros"]
TAR = ["Tarifa - Total"] + [f"Tarifa - {f}" for f in FAMILIAS]


def tabela_tarifas():
    """tarifa_long with the pxq key (same rule as the notebook, Passo 4) and the attributes."""
    t = pd.read_csv(ARQ_TARIFA, sep=";", decimal=",", encoding="utf-8-sig")
    suf = {"Ponta": "P", "Fora ponta": "FP", "Intermediário": "INT", NSA: ""}
    pos = [("D" if u == "kW" else "E") + suf.get(p, "") for u, p in zip(t["UNIDADE"], t["POSTO"])]
    pos = [NSA if (p == "E" and s == SUB_DS[1]) else p for p, s in zip(pos, t["SUBCLASSE"])]
    t["CHAVE"] = (t["SUBGRUPO"] + t["MODALIDADE"] + t["CLASSE"] + t["SUBCLASSE"] + t["DETALHE"] + t["ACESSANTE"]
                  + pd.Series(pos, index=t.index) + t["UNIDADE"])
    t["VIGENCIA_INICIO"] = pd.to_datetime(t["VIGENCIA_INICIO"])
    t["VIGENCIA_FIM"] = pd.to_datetime(t["VIGENCIA_FIM"])
    return t


TARIFAS = tabela_tarifas()
ATRIB = (TARIFAS.drop_duplicates("CHAVE").set_index("CHAVE")
         [["SUBGRUPO", "MODALIDADE", "SUBCLASSE", "DETALHE", "ACESSANTE", "POSTO", "UNIDADE"]])


def atributos(chave):
    """Attributes of a pxq key: from the tariff table or, for the SD keys, read off the key itself."""
    if chave in ATRIB.index:
        return ATRIB.loc[chave].to_dict()
    s = str(chave)
    sg = next((x for x in SUBGRUPOS if s.startswith(x)), None)
    md = next((x for x in MODALIDADES if s[len(sg or ""):].startswith(x)), None)
    un = "kW" if s.endswith("kW") else "MWh"
    meio = s[len(sg or "") + len(md or ""):-len(un)]
    posto = NSA
    for ab, nome in (("INT", "Intermediário"), ("FP", "Fora ponta"), ("P", "Ponta")):
        if meio.endswith(("D" if un == "kW" else "E") + ab):
            posto, meio = nome, meio[:-(len(ab) + 1)]
            break
    else:
        meio = meio[:-1]                                    # drop the E/D prefix of the period
    acess = meio.replace(NSA, "") or NSA                    # SD keys: class, subclass, detail = NSA
    if sg is None or md is None:
        raise ValueError(f"chave sem atributos na tabela de tarifas: {chave}")
    return {"SUBGRUPO": sg, "MODALIDADE": md, "SUBCLASSE": NSA, "DETALHE": NSA, "ACESSANTE": acess,
            "POSTO": posto, "UNIDADE": un}


def tarifa_itens(ref, subclasse, itens):
    """TA item tariffs of B1 Convencional Residencial <subclasse> (MWh) in force at the month end."""
    fim = pd.Timestamp(f"{ref}-01") + pd.offsets.MonthEnd(0)
    t = TARIFAS[(TARIFAS["CAMADA"] == "TA") & (TARIFAS["SUBGRUPO"] == "B1") & (TARIFAS["MODALIDADE"] == "Convencional")
                & (TARIFAS["SUBCLASSE"] == subclasse) & (TARIFAS["DETALHE"] == NSA) & (TARIFAS["UNIDADE"] == "MWh")
                & (TARIFAS["VIGENCIA_INICIO"] <= fim) & (TARIFAS["VIGENCIA_FIM"] >= fim)]
    if t.empty:
        raise ValueError(f"{ref}: sem tarifa {subclasse} na tabela de tarifas")
    return t.iloc[0][itens].astype(float)


def linhas_mes(ref, modelo):
    """Tariff rows of the month: key attributes, volume and R$ by item and layer (bloco 2)."""
    det = {}
    grf.converte(ref, *modelo, grava=False, detalhe=det)
    t, so_ta, tar_ta = det["tar"], det["so_ta"].loc[det["tar"].index], det["tar_ta"]
    itens = [c[4:] for c in det["cols_comp"]]
    comp = pd.read_excel(DIR_PXQ / f"pxq_{ref}.xlsx", "Componentes", usecols=["Chave", "TOTAL DEMANDA", COL_F01])
    reg = pd.read_excel(DIR_REG / f"PxQ_regulacao_{ref}.xlsx")
    if not (reg.loc[t.index, "Chave"].astype(str).values == t["Chave"].astype(str).values).all():
        raise ValueError(f"{ref}: PxQ_regulacao fora da ordem das linhas do pxq (regerar o notebook)")
    be = t[det["cols_comp"]].fillna(0.0).copy()
    be.loc[so_ta] = tar_ta.loc[so_ta, det["cols_comp"]].fillna(0.0).values    # SD keys with only TA
    be.columns = itens
    cva = reg.loc[t.index, [f"CVA {i}" for i in itens]].fillna(0.0).set_axis(itens, axis=1)
    fin = reg.loc[t.index, [f"FIN {i}" for i in itens]].fillna(0.0).set_axis(itens, axis=1)
    cva.loc[so_ta], fin.loc[so_ta] = 0.0, 0.0
    fam = {i: FAM_PXQ.get((DE_PARA_MR.get(i) or ("OUTROS",))[0]) for i in itens}

    at = pd.DataFrame([atributos(c) for c in t["Chave"]], index=t.index)
    base = pd.DataFrame({
        "TIPO": t["TIPO"], "UN": t["UNIDADE"], "OPCAO": t["OPCAO"], "CLASSE_LAY": t["CLASSE"],
        "Natureza Mercado": t["OPCAO"].map(NATUREZA), "Classe": t["CLASSE"].map(CLASSE),
        "SUBGRUPO": at["SUBGRUPO"].str.upper(), "Modalidade": at["MODALIDADE"].str.upper(),
        "Subclasse": at["SUBCLASSE"], "Detalhe": np.where(at["ACESSANTE"] == "Ceris", "<500", NSA),
        "Posto": at["POSTO"], "Unidade": at["UNIDADE"].str.upper(),
        "Agente": np.where(t["Source"] == "SD", at["ACESSANTE"], NSA),
        "Volume": np.where(t["UNIDADE"] == "D", comp.loc[t.index, "TOTAL DEMANDA"], t["CONSUMO"]),
        "F01": comp.loc[t.index, COL_F01].fillna(0.0).clip(0.0, 1.0)}, index=t.index)
    if base["Classe"].isna().any() or base["Natureza Mercado"].isna().any():
        raise ValueError(f"{ref}: classe ou mercado sem de-para: {base.loc[base['Classe'].isna(), 'CLASSE_LAY'].unique()}")

    def receitas(b, c, f, peso=None):
        """R$ mil by family for item frames b (BE), c (CVA), f (FIN), optionally weighted per item."""
        if peso is not None:
            b, c, f = b * peso, c * peso, f * peso
        out = pd.DataFrame(index=b.index)
        out["Receita - Total"] = (b.sum(axis=1) + c.sum(axis=1) + f.sum(axis=1)) / 1000.0
        for nome in FAMILIAS[:-1]:
            out[f"Receita - {nome}"] = b[[i for i in itens if fam[i] == nome]].sum(axis=1) / 1000.0
        out["Receita - CVA"] = c.sum(axis=1) / 1000.0
        out["Receita - Financeiros"] = out["Receita - Total"] - out[[f"Receita - {x}" for x in FAMILIAS]].sum(axis=1)
        return out

    partes = []
    ts = base["Subclasse"].eq("Baixa Renda")
    ds = base["Subclasse"].eq(SUB_DS[1])
    resto = ~ts & ~ds
    partes.append(base[resto].join(receitas(be[resto], cva[resto], fin[resto])))
    if ts.any():                                        # Tarifa Social: faixa 01 is exempt (no revenue)
        b = base[ts]
        f1 = b.assign(Subclasse=SUB_TS[0], Volume=b["Volume"] * b["F01"])
        f2 = b.assign(Subclasse=SUB_TS[1], Volume=b["Volume"] * (1.0 - b["F01"]))
        partes += [f1.join(receitas(be[ts] * 0.0, cva[ts] * 0.0, fin[ts] * 0.0)), f2.join(receitas(be[ts], cva[ts], fin[ts]))]
    if ds.any():                                        # Desconto Social: item shares by the faixa tariffs
        b = base[ds]
        t1, t2 = tarifa_itens(ref, SUB_DS[0], itens), tarifa_itens(ref, SUB_DS[1], itens)
        w1 = pd.DataFrame(np.outer(b["F01"], t1), index=b.index, columns=itens)
        w2 = pd.DataFrame(np.outer(1.0 - b["F01"], t2), index=b.index, columns=itens)
        s1 = (w1 / (w1 + w2).replace(0.0, np.nan)).fillna(0.0)
        f1 = b.assign(Subclasse=SUB_DS[0], Volume=b["Volume"] * b["F01"])
        f2 = b.assign(Subclasse=SUB_DS[1], Volume=b["Volume"] * (1.0 - b["F01"]))
        partes += [f1.join(receitas(be[ds], cva[ds], fin[ds], s1)), f2.join(receitas(be[ds], cva[ds], fin[ds], 1.0 - s1))]
    out = pd.concat(partes)
    out["Subclasse"] = out["Subclasse"].replace({"Baixa Renda": SUB_TS[1]})
    out["Mes_Ano"] = pd.Timestamp(f"{ref}-01")
    return out


def layout_conf(ref):
    """Receita Faturada layout of the month: CONSUMO (MWh / kW) and tariff total (R$ mil) by option x class x unit."""
    d = pd.read_excel(DIR_RF / f"receita_faturada_{ref}.xlsx", sheet_name="Receita Faturada", header=2)
    d = d[d["OpcaoCompra"].notna() & d["TipoFaturamento"].notna()]
    ini, fim = list(d.columns).index("Parcela B"), list(d.columns).index("Base CVA")
    d["TOTAL_TAR"] = d.iloc[:, ini:fim + 1].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1)
    d["VOL"] = np.where(d["Unidade"] == "D", d["CONSUMO"] * 1000.0, d["CONSUMO"])     # MW -> kW
    return d.groupby(["OpcaoCompra", "Classe", "Unidade"])[["VOL", "TOTAL_TAR"]].sum()


def agrega(linhas, opcao):
    df = linhas.copy()
    df["Grandeza"] = [GRANDEZA[opcao][(u, tp)] for u, tp in zip(df["UN"], df["TIPO"])]
    g = df.groupby(CHAVES, dropna=False)[["Volume"] + REC].sum().reset_index()
    for f, r in zip(TAR, REC[:-1]):
        g[f] = (g[r] / g["Volume"].replace(0.0, np.nan) * 1000.0)
    g["Chave 1"] = g[["SUBGRUPO", "Modalidade", "Classe", "Subclasse", "Detalhe", "Posto", "Unidade", "Agente"]].agg(" - ".join, axis=1)
    g["Chave 2"] = g[["SUBGRUPO", "Modalidade", "Detalhe", "Posto", "Unidade", "Agente"]].agg(" - ".join, axis=1)
    g["Corresp"] = None
    return g[["Chave 1", "Chave 2", "Corresp"] + CHAVES + ["Volume"] + TAR + REC]


def grava(tab, conf, opcao):
    DIR_SAIDA.mkdir(parents=True, exist_ok=True)
    arq = DIR_SAIDA / f"relatorio_energia_bip_fct_{ANO}_{opcao}.xlsx"
    leia = pd.DataFrame({"Relatório de energia realizada para BIP/FCT": [
        f"Opção: {'FAT + REFAT + TOI agregados (FATURADA_MWH, DEMANDA_KW)' if opcao == 'agregado' else 'aberto por tipo de faturamento (FAT, REFAT, TOI)'}.",
        "Fonte: bloco 2 (entrega) do P × Q, só a receita tarifável; mesma classificação do layout Receita Faturada.",
        "Volume: energia em MWh e demanda em kW. REFAT e TOI no mês de faturamento.",
        "Receita em R$ mil por família (Fio B, RI, Energia, Perdas, Fio A, Encargos, CVA); Financeiros = Total − famílias.",
        "Tarifa média = receita ÷ volume × 1.000 (R$/MWh ou R$/kW).",
        "SCEE somado na classe; APE e fonte incentivada somados no livre; irrigação/aquicultura sem abertura.",
        "Faixas sociais pela estimativa do P × Q: Tarifa Social – faixa 01 com receita zero; Desconto Social repartido pelas tarifas das faixas.",
        "Fora do relatório: bandeira, tributos, UD/ER, Outros, Covid Migrantes e resíduo sem item tarifário.",
        "Conferência: volume e receita contra o layout Receita Faturada (CONSUMO e total Parcela B a Base CVA).",
        "Premissas: procedimento, 'Relatório de energia realizada para BIP/FCT: premissas'."]})
    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        tab.to_excel(w, sheet_name="PXQ", index=False, startrow=2, startcol=1)
        conf.to_excel(w, sheet_name="Conferencia", index=False)
        leia.to_excel(w, sheet_name="Leia-me", index=False)
    wb = load_workbook(arq)
    ws = wb["PXQ"]
    n = len(tab) + 3
    ultima = get_column_letter(1 + tab.shape[1])
    tb = Table(displayName="tabPxQ", ref=f"B3:{ultima}{n}")
    tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tb)
    col = {c: get_column_letter(2 + i) for i, c in enumerate(tab.columns)}
    ws[f"{col['Volume']}2"] = "=SUBTOTAL(109,tabPxQ[Volume])"
    for r in REC:
        ws[f"{col[r]}2"] = f"=SUBTOTAL(109,tabPxQ[{r}])"
    for r in range(4, n + 1):
        ws[f"{col['Mes_Ano']}{r}"].number_format = "mmm/yyyy"
        for c in ["Volume"] + TAR + REC:
            ws[f"{col[c]}{r}"].number_format = "#,##0.00;-#,##0.00;\"-\""
    for c, l in col.items():
        ws.column_dimensions[l].width = 14 if c not in ("Chave 1", "Chave 2") else 40
    ws.freeze_panes = "F4"
    wc = wb["Conferencia"]
    for c in wc[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True)
    for row in wc.iter_rows(min_row=2):
        for c in row[4:]:
            c.number_format = "#,##0.000;-#,##0.000;\"-\""
    wb["Leia-me"].column_dimensions["A"].width = 130
    wb.save(arq)
    return arq


if __name__ == "__main__":
    modelo = grf.le_modelo()
    meses = sys.argv[2:] or [f"{ANO}-{m:02d}" for m in range(1, 13)
                             if (DIR_PXQ / f"pxq_{ANO}-{m:02d}.xlsx").exists()
                             and (DIR_RF / f"receita_faturada_{ANO}-{m:02d}.xlsx").exists()]
    todas, conf = [], []
    for ref in meses:
        lin = linhas_mes(ref, modelo)
        todas.append(lin)
        rel = lin.groupby(["OPCAO", "CLASSE_LAY", "UN"])[["Volume", "Receita - Total"]].sum()
        lay = layout_conf(ref)
        k = rel.index.union(lay.index)
        r, l = rel.reindex(k, fill_value=0.0), lay.reindex(k, fill_value=0.0)
        for i in k:
            conf.append([ref, *i, r.loc[i, "Volume"], l.loc[i, "VOL"], r.loc[i, "Volume"] - l.loc[i, "VOL"],
                         r.loc[i, "Receita - Total"], l.loc[i, "TOTAL_TAR"], r.loc[i, "Receita - Total"] - l.loc[i, "TOTAL_TAR"]])
        dv = max(abs(x[6]) for x in conf if x[0] == ref)
        dr = max(abs(x[9]) for x in conf if x[0] == ref)
        print(f"{ref}: {len(lin)} linhas | volume x layout: maior diferenca {dv:,.6f} | receita x layout: "
              f"maior diferenca R$ mil {dr:,.6f}")
    conf = pd.DataFrame(conf, columns=["Mês", "Opção (layout)", "Classe (layout)", "Unidade",
                                       "Volume relatório (MWh / kW)", "Volume layout (MWh / kW)", "Diferença volume",
                                       "Receita relatório (R$ mil)", "Receita layout, Parcela B a Base CVA (R$ mil)",
                                       "Diferença receita"])
    linhas = pd.concat(todas)
    for opcao in ("agregado", "por_tipo"):
        arq = grava(agrega(linhas, opcao), conf, opcao)
        print(f"gravado: {arq}")
