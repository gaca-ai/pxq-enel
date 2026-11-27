# -*- coding: utf-8 -*-
"""Extract the homologated tariff (TA, BE, CVA, FIN) from the PCAT workbooks of one or more
distributors and build ONE long table, one row per (distribuidora x vigencia x camada x chave),
with VIGENCIA_INICIO / VIGENCIA_FIM as real dates.

Common version of SP/BW/3-Py/1-Entrega/extrai_pcat.py (SP MVP, left untouched). The reading
logic is the same; what changes:
  - the PCAT files and vintages are configured per distributor (DISTRIBUIDORAS below);
  - a vintage is identified by VIGENCIA (e.g. "2026_recalculo"), not only by the year, because
    a distributor can have more than one vintage in the same year (CE 2026: annual adjustment
    from 22/04 and the UBP recalculation from 26/08);
  - the output carries the columns DISTRIBUIDORA and VIGENCIA.

Adding a new vintage means adding one entry to DISTRIBUIDORAS and re-running this script.

Columns are matched by NAME (header row, row 3 of each PCAT sheet), never by fixed position, so a
component added or removed by ANEEL between vintages (e.g. TE_ANGRA, introduced in 2026) does not
break the read -- it is reported below and filled with 0.0 in the vintages where it is absent.

Usage:
  python extrai_pcat.py                       -> RJ and CE, output common/tbl/tarifas/tarifa_long.csv
  python extrai_pcat.py --distribuidoras RJ   -> RJ only
  python extrai_pcat.py --distribuidoras SP --saida <file>   -> SP, only to check against the MVP
"""
import os
import argparse
import datetime as dt
from pathlib import Path

import pandas as pd
import openpyxl

ROOT = Path(os.environ.get("PXQ_ROOT", r"C:\python_projects\claude\p_x_q"))
SAIDA_PADRAO = ROOT / "common" / "tbl" / "tarifas" / "tarifa_long.csv"
PCAT_CE_RJ = ROOT / "CE_RJ" / "FAT" / "src" / "PCAT"

# One entry per PCAT vintage, per distributor. Vintages must be contiguous (no gap, no overlap).
# Sources of the dates: REH / Notas Tecnicas of each process (see CE_RJ/FAT/README, section 3).
DISTRIBUIDORAS = {
    # SP: kept only to check this script against the MVP output (SP/BW/tbl/tarifas)
    "SP": [
        {"vigencia": "2024", "ano": 2024,
         "arquivo": ROOT / "SP" / "BW" / "src" / "pcats" / "390 PCAT Enel SP 2024 V02.xlsx",
         "inicio": dt.date(2024, 7, 4), "fim": dt.date(2025, 7, 3)},
        {"vigencia": "2025", "ano": 2025,
         "arquivo": ROOT / "SP" / "BW" / "src" / "pcats" / "390 PCAT Enel SP 2025 V02_25jun TA.xlsx",
         "inicio": dt.date(2025, 7, 4), "fim": dt.date(2026, 7, 3)},
        {"vigencia": "2026", "ano": 2026,
         "arquivo": ROOT / "SP" / "BW" / "src" / "pcats" / "390 PCAT Enel SP 2026 V02_23junho.xlsx",
         "inicio": dt.date(2026, 7, 4), "fim": dt.date(2027, 7, 3)},
    ],
    # RJ: annual adjustment on 15/03 (REH 3.435/2025; Nota Tecnica 33/2026). PCAT 2024 not received.
    "RJ": [
        {"vigencia": "2025", "ano": 2025,
         "arquivo": PCAT_CE_RJ / "RJ" / "383 PCAT Enel RJ 2025 V02.xlsx",
         "inicio": dt.date(2025, 3, 15), "fim": dt.date(2026, 3, 14)},
        {"vigencia": "2026", "ano": 2026,
         "arquivo": PCAT_CE_RJ / "RJ" / "383 PCAT Enel RJ 2026 V02.xlsx",
         "inicio": dt.date(2026, 3, 15), "fim": dt.date(2027, 3, 14)},
    ],
    # CE: annual adjustment on 22/04 (REH 3.445/2025; REH 3.574/2026). 2026 was recalculated with
    # the UBP funds from 26/08/2026 (Despacho ANEEL 3.324/2026), acting as a new vintage.
    # PCAT 2024 not received.
    "CE": [
        {"vigencia": "2025", "ano": 2025,
         "arquivo": PCAT_CE_RJ / "CE" / "39 PCAT Enel CE 2025 V02.xlsx",
         "inicio": dt.date(2025, 4, 22), "fim": dt.date(2026, 4, 21)},
        {"vigencia": "2026_reajuste", "ano": 2026,
         "arquivo": PCAT_CE_RJ / "CE" / "39 PCAT Enel CE 2026 V02_reajuste.xlsx",
         "inicio": dt.date(2026, 4, 22), "fim": dt.date(2026, 8, 25)},
        {"vigencia": "2026_recalculo", "ano": 2026,
         "arquivo": PCAT_CE_RJ / "CE" / "39 PCAT Enel CE 2026 V02_recalculo.xlsx",
         "inicio": dt.date(2026, 8, 26), "fim": dt.date(2027, 4, 21)},
    ],
}

ABAS = {"TA": "TA - Aplicação", "BE": "TA - BE", "CVA": "TA - CVA"}
ATRIBUTOS = ["SUBGRUPO", "MODALIDADE", "CLASSE", "SUBCLASSE", "DETALHE",
             "POSTO", "UNIDADE", "ACESSANTE"]
TOTAIS = ["Total TUSD", "Total TE"]

# Desconto Social (Lei 15.235/2025, REN ANEEL 1147/2025). PCATs published before
# the benefit existed have no row for it, but it is billed since 2026. The tariff is fully
# defined by the regulation, so the missing rows are derived from the B1 Residencial row of
# the same vintage:
#   faixa 01 (up to 120 kWh): Residencial without the exempt items
#            (PRORET 7.1 par. 18.V - TUSD alineas d/f; par. 23.V - TE alinea f);
#   faixa 02 (above 120 kWh): Residencial as is (REN 1147, art. 179 par. 4 II).
# Derived rows carry ORIGEM = "DERIVADA_RESIDENCIAL"; PCAT rows carry "PCAT".
# In vintages that already publish the benefit the native rows are kept and the
# derivation only runs as a check against them.
DS_FAIXA_01 = "Residencial Desconto Social – faixa 01"
DS_FAIXA_02 = "Residencial Desconto Social – faixa 02"
ISENTOS_DESCONTO_SOCIAL = ["TUSD_CDE", "TUSD_CDE_COVID", "TE_CDE_GD"]


# TE of the supply to distributors (A4 Distribuição keys). Vintages may publish these keys with
# TE = 0 in the "TA - ..." sheets; the TE of supply is only in the TE sheets, row "Energia
# suprimento permissionária / concessionária" (<500GWh), with the layers apart: "TE BE" (BE),
# "TE CVA" (CVA) and "TE BF" (financeiros, FIN). Those rows fill the TE items of the keys whose
# TE is zero; ORIGEM = "TE_SUPRIMENTO". Keys that already have TE are not touched.
ABAS_TE_SUPRIMENTO = {"BE": "TE BE", "CVA": "TE CVA", "FIN": "TE BF"}
# Supply row used for each (distribuidora, acessante): (row name, market band of column E).
# Default (SP MVP rule): concessionária, band <500GWh.
LINHA_SUPRIMENTO_PADRAO = ("Energia suprimento concessionária", "<500GWh")
LINHA_SUPRIMENTO = {
    ("SP", "Ceris"): ("Energia suprimento permissionária", "<500GWh"),
    # RJ cooperatives (permissionárias)
    ("RJ", "CERCI"): ("Energia suprimento permissionária", "<500GWh"),
    ("RJ", "Ceral Araruama"): ("Energia suprimento permissionária", "<500GWh"),
    ("RJ", "Ceres"): ("Energia suprimento permissionária", "<500GWh"),
    # RJ concessionárias: the 2026 PCAT only publishes the concessionária row with band
    # "Não se aplica" (no <500GWh row). Provisional choice, to be confirmed (CE_RJ/FAT/sessions).
    ("RJ", "Energisa Minas Rio"): ("Energia suprimento concessionária", "Não se aplica"),
    ("RJ", "Neoenergia Elektro"): ("Energia suprimento concessionária", "Não se aplica"),
}
# item label in the TE sheets (row 4) -> item column of the tariff table
ITEM_TE = {
    "P&D": "TE_PeD", "ESS/ERR": "TE_ESSERR", "CFURH": "TE_CFURH", "CDE Covid TE": "TE_CDE_COVID",
    "CDE ELET": "TE_CDE_ELET", "CDE GD": "TE_CDE_GD", "ENERGIA REVENDA": "TE_ENERGIA",
    "ITAIPU": "TE_TRANSPORTE_ITAIPU", "TUST ITAIPU": "TE_TUST_ITAIPU", "TUST CI": "TE_TUST_CI",
    "SUBSIDIO": "TE Subsidio", "PERDAS RB/C": "TE_Per_RB", "ANGRA": "TE_ANGRA",
}
# items that only exist in some vintages (TE_ANGRA: from 2026); absent -> 0.0
ITEM_TE_OPCIONAL = {"TE_ANGRA"}

COLUNAS_ID = ["DISTRIBUIDORA", "VIGENCIA", "ANO_VIGENCIA", "VIGENCIA_INICIO", "VIGENCIA_FIM", "CAMADA"]


def s(v):
    return "" if v is None else str(v).strip()


def num(v):
    if v is None or v == "":
        return 0.0
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def le_aba(caminho, aba):
    """Read one PCAT sheet. Header (row 3) drives which column is which --
    never a fixed index -- so the sheet can gain or lose item columns between
    vintages without breaking the read."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb[aba]
    linhas = list(ws.iter_rows(values_only=True))
    wb.close()

    cabecalho = linhas[2]
    idx = {s(v): i for i, v in enumerate(cabecalho) if s(v)}
    faltando = [a for a in ATRIBUTOS + TOTAIS if a not in idx]
    if faltando:
        raise ValueError(f"{caminho.name} / {aba}: colunas ausentes no cabecalho: {faltando}")

    reservadas = set(ATRIBUTOS) | set(TOTAIS)
    itens = [s(v) for v in cabecalho if s(v) and s(v) not in reservadas]

    registros = []
    for r in linhas[3:]:
        if not s(r[idx["SUBGRUPO"]]):
            continue
        reg = {a: s(r[idx[a]]) for a in ATRIBUTOS}
        for t in TOTAIS:
            reg[t] = num(r[idx[t]])
        for it in itens:
            reg[it] = num(r[idx[it]])
        registros.append(reg)
    return pd.DataFrame(registros), itens


def deriva_desconto_social(tabela, dist):
    """Add the Desconto Social rows to the vintages whose PCAT does not have them."""
    tabela = tabela.copy()
    tabela["ORIGEM"] = "PCAT"
    nativas = tabela[tabela["SUBCLASSE"].isin([DS_FAIXA_01, DS_FAIXA_02])]
    modalidades = sorted(nativas["MODALIDADE"].unique())
    if not modalidades:
        raise ValueError(f"{dist}: nenhuma vigencia traz o Desconto Social - sem referencia de modalidades")

    base = tabela[(tabela["SUBGRUPO"] == "B1") & (tabela["SUBCLASSE"] == "Residencial")
                  & tabela["MODALIDADE"].isin(modalidades)]
    vigencias_nativas = set(nativas["VIGENCIA"])
    chave = ["VIGENCIA", "CAMADA"] + ATRIBUTOS
    novas = []
    for vig, g in base.groupby("VIGENCIA"):
        f2 = g.copy()
        f2["SUBCLASSE"] = DS_FAIXA_02
        f1 = g.copy()
        f1["SUBCLASSE"] = DS_FAIXA_01
        for it in ISENTOS_DESCONTO_SOCIAL:
            total = "Total TE" if it.startswith("TE") else "Total TUSD"
            f1[total] = f1[total] - f1[it]
            f1[it] = 0.0
        derivadas = pd.concat([f1, f2], ignore_index=True)
        derivadas["ORIGEM"] = "DERIVADA_RESIDENCIAL"

        if vig in vigencias_nativas:
            # check only: how far the rule is from the published rows
            cmp_ = derivadas.merge(nativas, on=chave, suffixes=("_DER", "_PCAT"))
            ta = cmp_[cmp_["CAMADA"] == "TA"]
            for t in TOTAIS:
                d = (ta[f"{t}_DER"] - ta[f"{t}_PCAT"]).abs().max()
                print(f"{dist} {vig}: Desconto Social publicado no PCAT - regra derivada x PCAT, "
                      f"{t} (TA), diferenca maxima {d:.2f} R$/MWh")
            continue

        novas.append(derivadas)
        print(f"{dist} {vig}: Desconto Social ausente no PCAT - {len(derivadas)} linhas derivadas da "
              f"Residencial ({', '.join(modalidades)})")

    if novas:
        tabela = pd.concat([tabela] + novas, ignore_index=True)
    return tabela


def le_te_suprimento(caminho):
    """{(camada, linha, faixa): {item: R$/MWh}} of the supply rows of the TE sheets. Columns located
    by the item label (row 4), never by position; the row total (last column) is kept as a check.
    A row repeated in the sheet must carry the same values."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    out = {}
    try:
        for camada, aba in ABAS_TE_SUPRIMENTO.items():
            if aba not in wb.sheetnames:
                raise ValueError(f"{caminho.name}: aba {aba} ausente")
            linhas = list(wb[aba].iter_rows(values_only=True))
            rotulos = [s(v) for v in linhas[3]]
            col = {ITEM_TE[r]: j for j, r in enumerate(rotulos) if r in ITEM_TE}
            faltam = sorted(set(ITEM_TE.values()) - set(col) - ITEM_TE_OPCIONAL)
            if faltam:
                raise ValueError(f"{caminho.name} / {aba}: itens ausentes no cabecalho: {faltam}")
            total = max(j for j, v in enumerate(linhas[2]) if s(v))   # row total: last labelled column of row 3
            for r in linhas[4:]:
                nome = s(r[1])
                if nome.startswith("Energia suprimento"):
                    faixa = s(r[4])
                    valores = {it: num(r[j]) for it, j in col.items()}
                    soma = sum(valores.values())
                    if abs(soma - num(r[total])) > 0.01:
                        raise ValueError(f"{caminho.name} / {aba} / {nome} / {faixa}: itens somam "
                                         f"{soma:.4f}, total da linha {num(r[total]):.4f}")
                    chave = (camada, nome, faixa)
                    if chave in out and any(abs(out[chave][it] - x) > 1e-9 for it, x in valores.items()):
                        raise ValueError(f"{caminho.name} / {aba} / {nome} / {faixa}: linha repetida "
                                         f"com valores diferentes")
                    out[chave] = valores
    finally:
        wb.close()
    return out


def completa_te_suprimento(tabela, vigencias, dist):
    """Fill the TE items of the A4 Distribuição MWh keys whose TE is zero, from the supply rows of
    the TE sheets of the same vintage (layers BE, CVA, FIN; TA = BE + CVA + FIN)."""
    tabela = tabela.copy()
    for v in vigencias:
        alvo = ((tabela["VIGENCIA"] == v["vigencia"]) & (tabela["MODALIDADE"] == "Distribuição")
                & (tabela["UNIDADE"] == "MWh"))
        sem_te = tabela.loc[alvo & (tabela["CAMADA"] == "TA") & (tabela["Total TE"].abs() < 1e-9), "ACESSANTE"]
        if sem_te.empty:
            continue
        sup = le_te_suprimento(v["arquivo"])
        if not sup:
            # no supply row in the TE sheets: the distributor only provides network use (TUSD)
            # to these acessantes, so their TE is zero by definition (e.g. CE)
            print(f"{dist} {v['vigencia']}: PCAT sem linha de suprimento nas abas de TE - TE mantida "
                  f"zero para {sorted(sem_te.unique())} (somente uso da rede)")
            continue
        for acessante in sorted(sem_te.unique()):
            linha, faixa = LINHA_SUPRIMENTO.get((dist, acessante), LINHA_SUPRIMENTO_PADRAO)
            ta = dict.fromkeys(ITEM_TE.values(), 0.0)
            for camada in ("BE", "CVA", "FIN"):
                if (camada, linha, faixa) not in sup:
                    raise ValueError(f"{dist} {v['vigencia']}: linha '{linha}' / '{faixa}' ausente na aba "
                                     f"{ABAS_TE_SUPRIMENTO[camada]}")
                vals = sup[(camada, linha, faixa)]
                m = alvo & (tabela["CAMADA"] == camada) & (tabela["ACESSANTE"] == acessante)
                for it, x in vals.items():
                    if it not in tabela.columns:
                        tabela[it] = 0.0
                    tabela.loc[m, it] = x
                    ta[it] += x
                tabela.loc[m, "Total TE"] = sum(vals.values())
                tabela.loc[m, "ORIGEM"] = "TE_SUPRIMENTO"
            m = alvo & (tabela["CAMADA"] == "TA") & (tabela["ACESSANTE"] == acessante)
            for it, x in ta.items():
                tabela.loc[m, it] = x
            tabela.loc[m, "Total TE"] = sum(ta.values())
            tabela.loc[m, "ORIGEM"] = "TE_SUPRIMENTO"
            print(f"{dist} {v['vigencia']}: TE de suprimento completada para {acessante} ({linha}, {faixa}): "
                  f"TA {sum(ta.values()):.2f} R$/MWh")
    return tabela


def extrai_distribuidora(dist, vigencias):
    """Long tariff table of one distributor (all its vintages and layers)."""
    blocos = []
    itens_por_vig = {}
    pior_fin_geral = 0.0

    for v in vigencias:
        camadas, itens_por_aba = {}, {}
        for camada, aba in ABAS.items():
            df, itens = le_aba(v["arquivo"], aba)
            camadas[camada] = df
            itens_por_aba[camada] = set(itens)

        n = {c: len(camadas[c]) for c in camadas}
        if len(set(n.values())) != 1:
            raise ValueError(f"{dist} {v['vigencia']}: numero de linhas nao bate entre TA/BE/CVA: {n}")

        # duplicate-key check per layer
        for camada, df in camadas.items():
            dups = df.duplicated(subset=ATRIBUTOS).sum()
            if dups:
                print(f"ATENCAO {dist} {v['vigencia']} / {camada}: {dups} chave(s) duplicada(s) na aba")

        itens_desta_vigencia = sorted(set().union(*itens_por_aba.values()))
        itens_por_vig[v["vigencia"]] = set(itens_desta_vigencia)
        for camada, df in camadas.items():
            for it in itens_desta_vigencia:
                if it not in df.columns:
                    df[it] = 0.0

        # FIN = TA - BE - CVA, merged on the 8-attribute key (never positional)
        merge_cols = ATRIBUTOS
        m = camadas["TA"].merge(camadas["BE"], on=merge_cols, suffixes=("", "_BE")) \
                          .merge(camadas["CVA"], on=merge_cols, suffixes=("", "_CVA"))
        fin = camadas["TA"][ATRIBUTOS].copy()
        for it in TOTAIS + itens_desta_vigencia:
            fin[it] = m[it] - m[f"{it}_BE"] - m[f"{it}_CVA"]
        camadas["FIN"] = fin

        # validation: TA = BE + CVA + FIN, item a item, for this vintage
        chk = camadas["TA"].merge(camadas["BE"], on=merge_cols, suffixes=("", "_BE")) \
                            .merge(camadas["CVA"], on=merge_cols, suffixes=("", "_CVA")) \
                            .merge(camadas["FIN"], on=merge_cols, suffixes=("", "_FIN"))
        pior = 0.0
        for it in TOTAIS + itens_desta_vigencia:
            d = (chk[it] - (chk[f"{it}_BE"] + chk[f"{it}_CVA"] + chk[f"{it}_FIN"])).abs().max()
            pior = max(pior, d)
        pior_fin_geral = max(pior_fin_geral, pior)
        print(f"{dist} {v['vigencia']}: {len(camadas['TA'])} chaves, {len(itens_desta_vigencia)} itens | "
              f"TA=BE+CVA+FIN, diferenca maxima {pior:.10f}")

        for camada, df in camadas.items():
            df = df.copy()
            df.insert(0, "CAMADA", camada)
            df.insert(0, "VIGENCIA_FIM", v["fim"])
            df.insert(0, "VIGENCIA_INICIO", v["inicio"])
            df.insert(0, "ANO_VIGENCIA", v["ano"])
            df.insert(0, "VIGENCIA", v["vigencia"])
            df.insert(0, "DISTRIBUIDORA", dist)
            blocos.append(df)

    # report item drift between consecutive vintages
    ordem = [v["vigencia"] for v in sorted(vigencias, key=lambda v: v["inicio"])]
    for prev, cur in zip(ordem, ordem[1:]):
        novos = itens_por_vig[cur] - itens_por_vig[prev]
        sumidos = itens_por_vig[prev] - itens_por_vig[cur]
        if novos:
            print(f"{dist}: itens novos em {cur} (nao existiam em {prev}): {sorted(novos)}")
        if sumidos:
            print(f"{dist}: itens que sumiram em {cur} (existiam em {prev}): {sorted(sumidos)}")
    if pior_fin_geral >= 0.01:
        raise ValueError(f"{dist}: validacao TA=BE+CVA+FIN falhou: diferenca maxima {pior_fin_geral}")

    todos_itens = sorted(set().union(*itens_por_vig.values()))
    tabela = pd.concat(blocos, ignore_index=True)
    for it in todos_itens:
        if it not in tabela.columns:
            tabela[it] = 0.0
    tabela = tabela[COLUNAS_ID + ATRIBUTOS + TOTAIS + todos_itens].fillna(0.0)
    tabela = deriva_desconto_social(tabela, dist)
    tabela = completa_te_suprimento(tabela, vigencias, dist)

    # sanity: vigencias contiguous, no gap or overlap
    ordenadas = sorted(vigencias, key=lambda v: v["inicio"])
    for a, b in zip(ordenadas, ordenadas[1:]):
        gap = (b["inicio"] - a["fim"]).days
        if gap != 1:
            print(f"ATENCAO {dist}: vigencia {a['vigencia']} termina em {a['fim']} e {b['vigencia']} "
                  f"comeca em {b['inicio']} -- {'sobreposicao' if gap < 1 else 'buraco'} de {1 - gap} dia(s)")
    return tabela


def main():
    ap = argparse.ArgumentParser(description="PCAT -> tarifa_long.csv (common, per distributor)")
    ap.add_argument("--distribuidoras", nargs="+", default=["RJ", "CE"], choices=sorted(DISTRIBUIDORAS))
    ap.add_argument("--saida", type=Path, default=SAIDA_PADRAO)
    args = ap.parse_args()

    tabelas = [extrai_distribuidora(d, DISTRIBUIDORAS[d]) for d in args.distribuidoras]

    # union of items across distributors; an item absent in one distributor is 0.0 there
    itens = sorted(set().union(*[set(t.columns) for t in tabelas]) - set(COLUNAS_ID + ATRIBUTOS + TOTAIS + ["ORIGEM"]))
    colunas = COLUNAS_ID + ATRIBUTOS + TOTAIS + itens + ["ORIGEM"]
    tabela = pd.concat(tabelas, ignore_index=True)
    tabela[itens] = tabela[itens].fillna(0.0)
    tabela = tabela[colunas]

    args.saida.parent.mkdir(parents=True, exist_ok=True)
    tabela.to_csv(args.saida, sep=";", decimal=",", index=False, encoding="utf-8-sig")
    print(f"\ngravado: {args.saida}  ({len(tabela)} linhas, {len(tabela.columns)} colunas)")
    for d in args.distribuidoras:
        print(f"{d}: vigencias {[v['vigencia'] for v in DISTRIBUIDORAS[d]]}")
    print(f"itens (uniao de todas as vigencias e distribuidoras): {len(itens)}")


if __name__ == "__main__":
    main()
