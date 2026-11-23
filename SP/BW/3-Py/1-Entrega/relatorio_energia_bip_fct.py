# -*- coding: utf-8 -*-
"""
Relatorio de energia realizada para BIP/FCT: a tabela tabPxQ do modelo do orcamento (aba PXQ),
preenchida com o realizado do P x Q, sem vinculos externos.

Uso:
    python relatorio_energia_bip_fct.py 2026            # todos os meses com pxq e receita_faturada
    python relatorio_energia_bip_fct.py 2026 2026-10    # so os meses indicados
    python relatorio_energia_bip_fct.py 2026 --tipo-rnf=PROVISAO   # RNF pela provisao (padrao: SALDO)
Le, por mes:
    1-Entrega/pxq/pxq_<ref>.xlsx                    grupos do pxq (bloco 2), pelo conversor do layout
    1-Entrega/pxq_regulacao/PxQ_regulacao_<ref>.xlsx camadas CVA e FIN de cada item
    1-Entrega/receita_faturada/receita_faturada_<ref>.xlsx    conferencia
    tbl/tarifas/tarifa_long.csv                     tarifas das faixas do Desconto Social
    ../RNF/1-Entrega/rnf/rnf_<ref>.xlsx             aba Componentes_unidade (bloco 2 do RNF por linha, unidade e
                                                    camada); no SALDO, tambem a do mes anterior
    ../RNF/1-Entrega/receita_nao_faturada/receita_nao_faturada_<ref>.xlsx   conferencia do RNF
Grava
    1-Entrega/relatorio_energia_bip_fct/relatorio_energia_bip_fct_<ano>_agregado.xlsx
        Grandezas FATURADA_MWH e DEMANDA_KW (FAT + REFAT + TOI)
    1-Entrega/relatorio_energia_bip_fct/relatorio_energia_bip_fct_<ano>_por_tipo.xlsx
        Grandezas FAT_MWH, REFAT_MWH, TOI_MWH, DEMANDA_FAT_KW, DEMANDA_REFAT_KW, DEMANDA_TOI_KW
    Nos dois: RNF_MWH e DEMANDA_RNF_KW, pelo TIPO_RNF (SALDO = provisao do mes - provisao do mes anterior;
    PROVISAO = provisao do mes; com PROVISAO o arquivo leva o sufixo _rnf_provisao).
    Abas: PXQ (tabela tabPxQ, colunas do modelo), Subsidio (tabela tabSubsidio), Conferencia e Leia-me.
Subsidio: desconto reembolsado pela CDE estimado no Bl1 x Bl2 (2-Checks/gera_receita_faturada_bl1_vs_bl2.py),
    Tarifa Social - DMR e fonte incentivada, linha a linha do pxq, nas chaves do relatorio (so faturamento);
    confere com 2-Checks/receita_faturada_Bl1_vs_Bl2/receita_faturada_Bl1_vs_Bl2_<ref>.xlsx (aba conciliacao).

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
- RNF: mesma regra, a partir da abertura do notebook Abertura_RNF (bloco 2 por item e camada). Energia
  em MWh e demanda em kW; faixas sociais pela estimativa do RNF (% energia na faixa 02).
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
import gera_receita_faturada_bl1_vs_bl2 as bvb        # noqa: E402
sys.argv = _argv

ANO = int(sys.argv[1]) if len(sys.argv) > 1 else dt.date.today().year
TIPO_RNF = "SALDO"            # "SALDO" (provision of the month - provision of the previous month) or "PROVISAO"
for _a in sys.argv[2:]:
    if _a.lower().startswith("--tipo-rnf="):
        TIPO_RNF = _a.split("=", 1)[1].strip().upper()
if TIPO_RNF not in ("SALDO", "PROVISAO"):
    raise ValueError(f"TIPO_RNF invalido: {TIPO_RNF} (use SALDO ou PROVISAO)")
DIR_PXQ = BASE / "1-Entrega" / "pxq"
DIR_REG = BASE / "1-Entrega" / "pxq_regulacao"
DIR_RF = BASE / "1-Entrega" / "receita_faturada"
ARQ_TARIFA = BASE / "tbl" / "tarifas" / "tarifa_long.csv"
DIR_SAIDA = BASE / "1-Entrega" / "relatorio_energia_bip_fct"
DIR_RNF = BASE.parent / "RNF" / "1-Entrega" / "rnf"
DIR_RNF_LAY = BASE.parent / "RNF" / "1-Entrega" / "receita_nao_faturada"
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
GRANDEZA_RNF = {("E", "RNF"): "RNF_MWH", ("D", "RNF"): "DEMANDA_RNF_KW"}
GRANDEZA = {"agregado": {("E", t): "FATURADA_MWH" for t in grf.TIPOS} | {("D", t): "DEMANDA_KW" for t in grf.TIPOS}
                        | GRANDEZA_RNF,
            "por_tipo": {("E", "FAT"): "FAT_MWH", ("E", "REFAT"): "REFAT_MWH", ("E", "CNR"): "TOI_MWH",
                         ("D", "FAT"): "DEMANDA_FAT_KW", ("D", "REFAT"): "DEMANDA_REFAT_KW",
                         ("D", "CNR"): "DEMANDA_TOI_KW"} | GRANDEZA_RNF}
CHAVES = ["Grandeza", "Natureza Mercado", "SUBGRUPO", "Modalidade", "Classe", "Subclasse", "Detalhe", "Posto",
          "Unidade", "Agente", "Mes_Ano"]
REC = ["Receita - Total"] + [f"Receita - {f}" for f in FAMILIAS] + ["Receita - Financeiros"]
TAR = ["Tarifa - Total"] + [f"Tarifa - {f}" for f in FAMILIAS]
SUB_ORIGEM = {"[1]": "Subsídio - Tarifa Social", "[A]": "Subsídio - Fonte Incentivada"}   # by first block of the origin
SUB = (["Subsídio - Total"] + list(SUB_ORIGEM.values()) + [f"Subsídio - {f}" for f in FAMILIAS]
       + ["Subsídio - Financeiros"])


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


def familia(itens):
    """Family of each tariff item (de-para of gera_de_para_regulacao_MR); OUTROS / unmapped -> None (Financeiros)."""
    return {i: FAM_PXQ.get((DE_PARA_MR.get(i) or ("OUTROS",))[0]) for i in itens}


def abre(base, be, cva, fin, ref):
    """R$ mil by family from the item frames be (BE), cva (CVA), fin (FIN), in R$, aligned with base.
    Opens the social faixas: Tarifa Social (Subclasse "Baixa Renda") - faixa 01 exempt (no revenue);
    Desconto Social (faixa 02 key) - item shares by the faixa tariffs in force at the end of ref."""
    itens = list(be.columns)
    fam = familia(itens)

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
    return out


def linhas_mes(ref, modelo):
    """Tariff rows of the month: key attributes, volume and R$ by item and layer (bloco 2)."""
    det = {}
    grf.converte(ref, *modelo, grava=False, detalhe=det)
    t, so_ta, tar_ta = det["tar"], det["so_ta"].loc[det["tar"].index], det["tar_ta"]
    itens = [c[4:] for c in det["cols_comp"]]
    reg = pd.read_excel(DIR_REG / f"PxQ_regulacao_{ref}.xlsx")
    if not (reg.loc[t.index, "Chave"].astype(str).values == t["Chave"].astype(str).values).all():
        raise ValueError(f"{ref}: PxQ_regulacao fora da ordem das linhas do pxq (regerar o notebook)")
    be = t[det["cols_comp"]].fillna(0.0).copy()
    be.loc[so_ta] = tar_ta.loc[so_ta, det["cols_comp"]].fillna(0.0).values    # SD keys with only TA
    be.columns = itens
    cva = reg.loc[t.index, [f"CVA {i}" for i in itens]].fillna(0.0).set_axis(itens, axis=1)
    fin = reg.loc[t.index, [f"FIN {i}" for i in itens]].fillna(0.0).set_axis(itens, axis=1)
    cva.loc[so_ta], fin.loc[so_ta] = 0.0, 0.0
    base = base_linhas(ref, t)
    out = abre(base, be, cva, fin, ref)
    out["Mes_Ano"] = pd.Timestamp(f"{ref}-01")
    return out


def base_linhas(ref, t):
    """Report attributes, volume (MWh / kW) and % faixa 01 of the tariff rows t of the layout converter."""
    comp = pd.read_excel(DIR_PXQ / f"pxq_{ref}.xlsx", "Componentes", usecols=["Chave", "TOTAL DEMANDA", COL_F01])
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
    return base


def subsidio_mes(ref, modelo):
    """Subsidy reimbursed by the CDE, row by row of the pxq, from the Bl1 x Bl2 blocks (same prepara):
    Tarifa Social - DMR = [1] - [D], fonte incentivada = [A] - [P]; R$ mil by family, positive = discount.
    Faixas: Tarifa Social - faixa 01 = its whole homologated value x DMR factor (revenue zero), faixa 02 =
    the rest; Desconto Social by the item shares of the faixa tariffs (as the revenue)."""
    f_dmr, _, _ = bvb.fator_dmr(ref)
    det = {}
    grf.converte(ref, *modelo, grava=False, detalhe=det, prepara=lambda c, b: bvb.prepara(c, b, f_dmr))
    t, ta = det["tar"], det["tar_ta"]
    itens = [c[4:] for c in det["cols_comp"]]
    fam = familia(itens)
    base = base_linhas(ref, t)
    f01 = base["F01"]
    ts = base["Subclasse"].eq("Baixa Renda")
    ds = base["Subclasse"].eq(SUB_DS[1])

    def col(df, bloco):
        return df[[f"{bloco} {i}" for i in itens]].apply(pd.to_numeric, errors="coerce").fillna(0.0).set_axis(itens, axis=1)

    def cva_tot(bloco):
        return pd.to_numeric(t[f"{bloco} Total CVA"], errors="coerce").fillna(0.0)

    def familias(o_ta, o_be, o_cva):
        """R$ mil: total (TA), BE families, CVA layer and Financeiros = total - families."""
        out = pd.DataFrame(index=o_ta.index)
        out["Subsídio - Total"] = o_ta.sum(axis=1) / 1000.0
        for nome in FAMILIAS[:-1]:
            out[f"Subsídio - {nome}"] = o_be[[i for i in itens if fam[i] == nome]].sum(axis=1) / 1000.0
        out["Subsídio - CVA"] = o_cva / 1000.0
        out["Subsídio - Financeiros"] = out["Subsídio - Total"] - out[[f"Subsídio - {x}" for x in FAMILIAS]].sum(axis=1)
        return out

    if ds.any():
        t1, t2 = tarifa_itens(ref, SUB_DS[0], itens), tarifa_itens(ref, SUB_DS[1], itens)
        w1 = pd.DataFrame(np.outer(f01[ds], t1), index=base.index[ds], columns=itens)
        w2 = pd.DataFrame(np.outer(1.0 - f01[ds], t2), index=base.index[ds], columns=itens)
        s1 = (w1 / (w1 + w2).replace(0.0, np.nan)).fillna(0.0)
    resto, faixa01 = [], []                        # per origin: family frames of (whole / faixa 02) and faixa 01
    for _, _, de, para in bvb.ABAS_ORIGEM:
        o_ta, o_be, o_cva = col(ta, de) - col(ta, para), col(t, de) - col(t, para), cva_tot(de) - cva_tot(para)
        p_ta, p_be = o_ta * 0.0, o_be * 0.0
        p_cva = o_cva * 0.0
        if ts.any():
            if de == "[1]":                        # Tarifa Social - DMR: faixa 01 is exempt (whole homologated value)
                p_ta.loc[ts] = (f_dmr * col(ta, "[1]").mul(f01, axis=0)).loc[ts]
                p_be.loc[ts] = (f_dmr * col(t, "[1]").mul(f01, axis=0)).loc[ts]
                p_cva.loc[ts] = (f_dmr * cva_tot("[1]") * f01).loc[ts]
            else:                                  # other origins on Tarifa Social rows: by volume
                p_ta.loc[ts], p_be.loc[ts] = o_ta.mul(f01, axis=0).loc[ts], o_be.mul(f01, axis=0).loc[ts]
                p_cva.loc[ts] = (o_cva * f01).loc[ts]
        if ds.any():
            p_ta.loc[ds], p_be.loc[ds] = o_ta.loc[ds] * s1, o_be.loc[ds] * s1
            sh = (p_ta.loc[ds].sum(axis=1) / o_ta.loc[ds].sum(axis=1).replace(0.0, np.nan)).fillna(0.0)
            p_cva.loc[ds] = o_cva.loc[ds] * sh
        r, f = familias(o_ta - p_ta, o_be - p_be, o_cva - p_cva), familias(p_ta, p_be, p_cva)
        r[SUB_ORIGEM[de]], f[SUB_ORIGEM[de]] = r["Subsídio - Total"], f["Subsídio - Total"]
        resto.append(r)
        faixa01.append(f)
    soma = lambda partes: sum(x.reindex(columns=SUB, fill_value=0.0) for x in partes)    # noqa: E731
    r, f = soma(resto), soma(faixa01)
    social = ts | ds
    linhas = [base[~social].join(r[~social])]
    for m, nomes in ((ts, SUB_TS), (ds, SUB_DS)):
        if m.any():
            b = base[m]
            linhas += [b.assign(Subclasse=nomes[0], Volume=b["Volume"] * b["F01"]).join(f[m]),
                       b.assign(Subclasse=nomes[1], Volume=b["Volume"] * (1.0 - b["F01"])).join(r[m])]
    out = pd.concat(linhas)
    neg = out[out["Subclasse"].isin([SUB_TS[1]]) & (out["Subsídio - Tarifa Social"] < -1e-6)]
    if len(neg):
        print(f"ATENCAO {ref}: Tarifa Social - faixa 02 com subsidio negativo em {len(neg)} linhas "
              f"(R$ mil {neg['Subsídio - Tarifa Social'].sum():,.3f}): % faixa 01 acima do desconto da linha")
    out = out[out[SUB].abs().max(axis=1) > 1e-9]
    out["Mes_Ano"] = pd.Timestamp(f"{ref}-01")
    return out


def confere_subsidio(ref, sub):
    """Subsidy totals of the month by origin (R$ mil) against the Bl1 x Bl2 reconciliation (column Layout, R$)."""
    arq = bvb.DIR_SAIDA / f"receita_faturada_Bl1_vs_Bl2_{ref}.xlsx"
    if not arq.exists():
        raise FileNotFoundError(f"{ref}: sem {arq.name} (rodar 2-Checks/gera_receita_faturada_bl1_vs_bl2.py)")
    c = pd.read_excel(arq, "conciliacao", header=None, usecols=[0, 1]).dropna().set_index(0)[1]
    out = []
    for (_, rot, de, _), coluna in zip(bvb.ABAS_ORIGEM, SUB_ORIGEM.values()):
        rel, bl = float(sub[coluna].sum()), float(c.loc[rot]) / 1000.0
        out.append([ref, f"Subsídio {rot} x Bl1 × Bl2 (conciliacao)", "(todas)", "(todas)", "(todas)",
                    None, None, None, rel, bl, rel - bl])
    print(f"{ref} Subsidio x Bl1 x Bl2: " + " | ".join(f"{x[1][9:].split(' x ')[0]} R$ mil {x[8]:,.3f} "
                                                        f"(diferenca {x[10]:,.6f})" for x in out))
    return out


def componentes_rnf(ref):
    """Componentes_unidade sheet of the RNF of the month (bloco 2 by RNF line x unit x layer, MR$)."""
    arq = DIR_RNF / f"rnf_{ref}.xlsx"
    if not arq.exists():
        raise FileNotFoundError(f"{ref}: sem {arq} (rodar o Abertura_RNF do mes)")
    if "Componentes_unidade" not in pd.ExcelFile(arq).sheet_names:
        raise ValueError(f"{ref}: {arq.name} sem a aba Componentes_unidade (regerar o Abertura_RNF do mes)")
    return pd.read_excel(arq, "Componentes_unidade")


def provisao_rnf(ref, sinal=1.0):
    """RNF provision of ref in the report rows (sinal = -1 for the reversal of the previous month)."""
    d = componentes_rnf(ref)
    itens = [str(c)[4:] for c in d.columns if str(c).startswith("[2] ")]
    d = d.set_index(["Linha", "Unidade"])
    cam = {c: d[d["Camada"] == c] for c in ("BE", "CVA", "FIN")}
    a = cam["BE"]
    if not (cam["CVA"].index.equals(a.index) and cam["FIN"].index.equals(a.index)) or not a.index.is_unique:
        raise ValueError(f"{ref}: Componentes_unidade com linhas fora de ordem entre as camadas")
    rs = {c: x[[f"[2] {i}" for i in itens]].fillna(0.0).set_axis(itens, axis=1) * 1e6 * sinal   # MR$ -> R$
          for c, x in cam.items()}
    un = a.index.get_level_values("Unidade")
    benef = a["Benefício social"]
    sub = np.select([benef.eq("BAIXA_RENDA"), benef.eq("DESCONTO_SOCIAL")], ["Baixa Renda", SUB_DS[1]],
                    a["Subclasse (tarifa)"].astype(str))
    base = pd.DataFrame({
        "TIPO": "RNF", "UN": un, "OPCAO": a["Mercado"], "CLASSE_LAY": a["Classe"],
        "Natureza Mercado": a["Mercado"].map(NATUREZA), "Classe": a["Classe"],
        "SUBGRUPO": a["Subgrupo"].str.upper(), "Modalidade": a["Modalidade"].str.upper(),
        "Subclasse": sub, "Detalhe": NSA,
        "Posto": np.where(un == "D", a["Posto da demanda"], a["Posto da energia"]),
        "Unidade": np.where(un == "D", "KW", "MWH"), "Agente": NSA,
        "Volume": a["Volume (MWh ou MW)"].fillna(0.0) * np.where(un == "D", 1000.0, 1.0) * sinal,   # MW -> kW
        "F01": (1.0 - a["% energia na faixa 02 (estimada)"]).where(benef.notna(), 0.0).fillna(0.0).clip(0.0, 1.0)},
        index=a.index)
    sem = ~base["Classe"].isin(set(CLASSE.values())) | base["Natureza Mercado"].isna()
    if sem.any():
        raise ValueError(f"{ref}: RNF com classe ou mercado sem de-para: {base.loc[sem, ['OPCAO', 'Classe']].drop_duplicates().values}")
    tem = (base["Volume"] != 0.0) | (sum(x.abs().sum(axis=1) for x in rs.values()) != 0.0)    # drop empty units
    return abre(base[tem], rs["BE"][tem], rs["CVA"][tem], rs["FIN"][tem], ref)


def linhas_rnf(ref):
    """RNF rows of the month by TIPO_RNF: PROVISAO = provision of ref; SALDO = provision of ref - provision of
    the previous month (the reversal), each opened with its own attributes and tariffs."""
    out = provisao_rnf(ref)
    if TIPO_RNF == "SALDO":
        ant = (pd.Timestamp(f"{ref}-01") - pd.offsets.MonthBegin(1)).strftime("%Y-%m")
        out = pd.concat([out, provisao_rnf(ant, -1.0)])
    out["Mes_Ano"] = pd.Timestamp(f"{ref}-01")
    return out


def layout_conf(arq, aba, classe_modelo=False):
    """Layout (Receita Faturada or receita_nao_faturada): CONSUMO (MWh / kW) and tariff total (R$ mil) by option x
    class x unit; classe_modelo=True groups by the template class (RNF lines carry the template class)."""
    d = pd.read_excel(arq, sheet_name=aba, header=2)
    d = d[d["OpcaoCompra"].notna() & d["TipoFaturamento"].notna()]
    ini, fim = list(d.columns).index("Parcela B"), list(d.columns).index("Base CVA")
    d["TOTAL_TAR"] = d.iloc[:, ini:fim + 1].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum(axis=1)
    d["VOL"] = np.where(d["Unidade"] == "D", d["CONSUMO"] * 1000.0, d["CONSUMO"])     # MW -> kW
    if classe_modelo:
        sem = d["Classe"].map(CLASSE).isna()
        if sem.any():
            raise ValueError(f"{arq.name} ({aba}): classe sem de-para: {d.loc[sem, 'Classe'].unique()}")
        d["Classe"] = d["Classe"].map(CLASSE)
    return d.groupby(["OpcaoCompra", "Classe", "Unidade"])[["VOL", "TOTAL_TAR"]].sum()


def confere(ref, nome, rel, lay):
    """Conference rows (report x layout) by option x class x unit; prints the largest differences."""
    k = rel.index.union(lay.index)
    r, l = rel.reindex(k, fill_value=0.0), lay.reindex(k, fill_value=0.0)
    out = [[ref, nome, *i, r.loc[i, "Volume"], l.loc[i, "VOL"], r.loc[i, "Volume"] - l.loc[i, "VOL"],
            r.loc[i, "Receita - Total"], l.loc[i, "TOTAL_TAR"], r.loc[i, "Receita - Total"] - l.loc[i, "TOTAL_TAR"]]
           for i in k]
    print(f"{ref} {nome}: volume x layout: maior diferenca {max(abs(x[7]) for x in out):,.6f} | receita x layout: "
          f"maior diferenca R$ mil {max(abs(x[10]) for x in out):,.6f} | receita R$ mil {sum(x[8] for x in out):,.3f}")
    return out


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


def agrega_subsidio(sub, opcao):
    df = sub.copy()
    df["Grandeza"] = [GRANDEZA[opcao][(u, tp)] for u, tp in zip(df["UN"], df["TIPO"])]
    g = df.groupby(CHAVES, dropna=False)[["Volume"] + SUB].sum().reset_index()
    g["Subsídio médio (R$/MWh ou R$/kW)"] = g["Subsídio - Total"] / g["Volume"].replace(0.0, np.nan) * 1000.0
    return g[CHAVES + ["Volume", "Subsídio médio (R$/MWh ou R$/kW)"] + SUB]


def grava(tab, conf, opcao, sub):
    DIR_SAIDA.mkdir(parents=True, exist_ok=True)
    arq = DIR_SAIDA / f"relatorio_energia_bip_fct_{ANO}_{opcao}{'' if TIPO_RNF == 'SALDO' else '_rnf_provisao'}.xlsx"
    leia = pd.DataFrame({"Relatório de energia realizada para BIP/FCT": [
        f"Opção: {'FAT + REFAT + TOI agregados (FATURADA_MWH, DEMANDA_KW)' if opcao == 'agregado' else 'aberto por tipo de faturamento (FAT, REFAT, TOI)'}.",
        "Fonte: bloco 2 (entrega) do P × Q, só a receita tarifável; mesma classificação do layout Receita Faturada.",
        "Volume: energia em MWh e demanda em kW. REFAT e TOI no mês de faturamento.",
        (f"RNF (RNF_MWH, DEMANDA_RNF_KW) pelo {'saldo: provisão do mês − provisão do mês anterior' if TIPO_RNF == 'SALDO' else 'valor da provisão do mês'}"
         "; mesma abertura do faturamento, a partir do bloco 2 do RNF (notebook Abertura_RNF)."),
        "Receita em R$ mil por família (Fio B, RI, Energia, Perdas, Fio A, Encargos, CVA); Financeiros = Total − famílias.",
        "Tarifa média = receita ÷ volume × 1.000 (R$/MWh ou R$/kW).",
        "SCEE somado na classe; APE e fonte incentivada somados no livre; irrigação/aquicultura sem abertura.",
        "Faixas sociais pela estimativa do P × Q: Tarifa Social – faixa 01 com receita zero; Desconto Social repartido pelas tarifas das faixas.",
        "Fora do relatório: bandeira, tributos, UD/ER, Outros, Covid Migrantes e resíduo sem item tarifário.",
        "Conferência: volume e receita contra o layout Receita Faturada e o receita_nao_faturada (CONSUMO e total Parcela B a Base CVA).",
        "Aba Subsidio: desconto reembolsado pela CDE estimado no Bl1 × Bl2 (Tarifa Social – DMR e fonte incentivada), só faturamento, R$ mil, positivo = desconto.",
        "Subsidio: Tarifa Social – faixa 01 = todo o valor homologado da faixa (isenta) × fator DMR; faixa 02 = o restante. Fora: isenção de encargos do APE, SCEE – Baixa Renda, Desconto Social e 2ª passada.",
        "Subsidio: volume = todo o volume da chave (não só dos clientes com desconto); subsídio médio = subsídio ÷ volume × 1.000.",
        "Subsidio: total por mês e origem confere com a aba conciliacao do Bl1 × Bl2 do mês (Conferencia).",
        "Premissas: procedimento, 'Relatório de energia realizada para BIP/FCT: premissas'."]})
    with pd.ExcelWriter(arq, engine="openpyxl") as w:
        tab.to_excel(w, sheet_name="PXQ", index=False, startrow=2, startcol=1)
        sub.to_excel(w, sheet_name="Subsidio", index=False, startrow=2, startcol=1)
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
    wsub = wb["Subsidio"]
    ns = len(sub) + 3
    tb = Table(displayName="tabSubsidio", ref=f"B3:{get_column_letter(1 + sub.shape[1])}{ns}")
    tb.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    wsub.add_table(tb)
    cs = {c: get_column_letter(2 + i) for i, c in enumerate(sub.columns)}
    wsub["B1"] = "Subsídio reembolsado pela CDE (Bl1 × Bl2), R$ mil; positivo = desconto"
    for r in ["Volume"] + SUB:
        wsub[f"{cs[r]}2"] = f"=SUBTOTAL(109,tabSubsidio[{r}])"
    for r in range(4, ns + 1):
        wsub[f"{cs['Mes_Ano']}{r}"].number_format = "mmm/yyyy"
        for c in ["Volume", "Subsídio médio (R$/MWh ou R$/kW)"] + SUB:
            wsub[f"{cs[c]}{r}"].number_format = "#,##0.00;-#,##0.00;\"-\""
    for l in cs.values():
        wsub.column_dimensions[l].width = 14
    wsub.freeze_panes = "D4"
    wc = wb["Conferencia"]
    for c in wc[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True)
    for row in wc.iter_rows(min_row=2):
        for c in row[5:]:
            c.number_format = "#,##0.000;-#,##0.000;\"-\""
    wb["Leia-me"].column_dimensions["A"].width = 130
    wb.save(arq)
    return arq


if __name__ == "__main__":
    modelo = grf.le_modelo()
    meses = [a for a in sys.argv[2:] if not a.startswith("--")] or [f"{ANO}-{m:02d}" for m in range(1, 13)
                             if (DIR_PXQ / f"pxq_{ANO}-{m:02d}.xlsx").exists()
                             and (DIR_RF / f"receita_faturada_{ANO}-{m:02d}.xlsx").exists()]
    todas, conf, subs = [], [], []
    for ref in meses:
        lin = linhas_mes(ref, modelo)
        rel = lin.groupby(["OPCAO", "CLASSE_LAY", "UN"])[["Volume", "Receita - Total"]].sum()
        conf += confere(ref, "Faturamento x Receita Faturada", rel,
                        layout_conf(DIR_RF / f"receita_faturada_{ref}.xlsx", "Receita Faturada"))
        rnf = linhas_rnf(ref)
        rel = rnf.groupby(["OPCAO", "Classe", "UN"])[["Volume", "Receita - Total"]].sum()
        aba = "Saldo" if TIPO_RNF == "SALDO" else "Provisao"
        conf += confere(ref, f"RNF ({TIPO_RNF}) x receita_nao_faturada ({aba})", rel,
                        layout_conf(DIR_RNF_LAY / f"receita_nao_faturada_{ref}.xlsx", aba, classe_modelo=True))
        todas += [lin, rnf]
        sm = subsidio_mes(ref, modelo)
        conf += confere_subsidio(ref, sm)
        subs.append(sm)
    conf = pd.DataFrame(conf, columns=["Mês", "Base", "Opção (layout)", "Classe", "Unidade",
                                       "Volume relatório (MWh / kW)", "Volume layout (MWh / kW)", "Diferença volume",
                                       "Receita relatório (R$ mil)", "Receita layout, Parcela B a Base CVA (R$ mil)",
                                       "Diferença receita"])
    linhas, subs = pd.concat(todas), pd.concat(subs)
    for opcao in ("agregado", "por_tipo"):
        arq = grava(agrega(linhas, opcao), conf, opcao, agrega_subsidio(subs, opcao))
        print(f"gravado: {arq}")
