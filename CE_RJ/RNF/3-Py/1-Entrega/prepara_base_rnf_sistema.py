"""
Consolidate the RNF system files (CE / RJ) into a single parquet file.

The commercial system delivers one text file per voltage level (AT and BT),
one line per customer, ';' separated, '.' as decimal, values padded with spaces.
The BT file is ~1 GB per month, so it is read in chunks and written
incrementally with a pyarrow ParquetWriter (memory stays bounded).

Usage (from CE_RJ/RNF):
    python 3-Py/1-Entrega/prepara_base_rnf_sistema.py --ref 2026-09 --nivel BT --uf RJ
    python 3-Py/1-Entrega/prepara_base_rnf_sistema.py --ref 2026-09 --nivel AT --uf RJ --entrada "src/CLIENTES_ALTA_*.txt"

Output: src/parquet/<UF>_RNF_SISTEMA_<NIVEL>_<AAAAMM>.parquet
"""

import argparse
import glob
import re
import sys
import unicodedata
from pathlib import Path

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

# ---------------------------------------------------------------------------
# Parameters
# ---------------------------------------------------------------------------
BASE = Path(__file__).resolve().parents[2]          # CE_RJ/RNF
DIR_SRC = BASE / "src"
DIR_SAIDA = DIR_SRC / "parquet"
CHUNK_ROWS = 500_000                                 # rows per chunk (tune to available memory)

# Default input pattern per voltage level (overridable with --entrada)
PADRAO_ENTRADA = {
    "AT": "CLIENTES_ALTA_*.txt",
    "BT": "CLIENTES_BAIXA_*.txt",
}

# Expected layout (40 fields, same for AT and BT), in file order
LAYOUT = [
    "NUMERO CLIENTE", "CONTA CONTRATO", "CLIENTE NOVO", "LOTE ANTERIOR", "LEITURA ANTERIOR",
    "ULTIMO LOTE", "ULTIMA LEITURA", "DIAS FATURADOS", "DIAS FATURADOS DENTRO DO MÊS DE REFERÊNCIA",
    "DIAS FATURADOS DO MÊS ANTERIOR DE REFERÊNCIA", "DIAS ESTIMADOS", "CONSUMO LIDO",
    "COMPRAS PARA O PERIODO FATURADO", "COMPRAS PARA O PERIODO ESTIMADO", "FATOR",
    "RNF SEM FATOR", "RNF COM FATOR", "CLASSE", "SUBCLASSE", "CODIGO DO MUNICIPIO", "CODIGO TARIFA",
    "BANDEIRA TARIFÁRIA SEM TRIBUTO", "BANDEIRA TARIFÁRIA COM TRIBUTO", "PREÇO PRO RATA TARIFA",
    "PREÇO MÉDIO", "VALOR MONETÁRIO", "DEMANDA FATURADA",
    "RENDA NÃO FATURADA DA DEMANDA FATURADA SEM ENCARGOS",
    "PIS/COFINS REFERENTE A RENDA NÃO FATURADA DA DEMANDA FATURADA", "CONSUMO FATURADO",
    "RNF SEM FATOR (CONSFAT)", "RNF COM FATOR (CONSFAT)", "BANDEIRA TARIFÁRIA SEM TRIBUTO (CONSFAT)",
    "BANDEIRA TARIFÁRIA COM TRIBUTO (CONSFAT)", "PREÇO PRO RATA TARIFA (CONSFAT)",
    "PREÇO MÉDIO (CONSFAT)", "VALOR MONETÁRIO (CONSFAT)", "DEMANDA FATURADA (CONSFAT)",
    "RENDA NÃO FATURADA DA DEMANDA FATURADA SEM ENCARGOS (CONSFAT)",
    "PIS/COFINS REFERENTE A RENDA NÃO FATURADA DA DEMANDA FATURADA (CONSFAT)",
]

# Field types (by original name); everything not listed here is numeric (float64)
CAMPOS_TEXTO = {
    "NUMERO CLIENTE", "CONTA CONTRATO", "CLIENTE NOVO", "LOTE ANTERIOR", "ULTIMO LOTE",
    "CLASSE", "SUBCLASSE", "CODIGO DO MUNICIPIO", "CODIGO TARIFA",
}
CAMPOS_DATA = {"LEITURA ANTERIOR", "ULTIMA LEITURA"}                  # AAAAMMDD
CAMPOS_INTEIRO = {
    "DIAS FATURADOS", "DIAS FATURADOS DENTRO DO MÊS DE REFERÊNCIA",
    "DIAS FATURADOS DO MÊS ANTERIOR DE REFERÊNCIA", "DIAS ESTIMADOS",
}


def snake(nome: str) -> str:
    """Original field name -> ASCII snake_case (parquet column name)."""
    s = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode()
    s = re.sub(r"[^0-9a-zA-Z]+", "_", s).strip("_").lower()
    return s


COLUNAS = {c: snake(c) for c in LAYOUT}

# Explicit arrow schema, so every chunk is written with identical types
SCHEMA = pa.schema(
    [
        (COLUNAS[c],
         pa.string() if c in CAMPOS_TEXTO
         else pa.date32() if c in CAMPOS_DATA
         else pa.int32() if c in CAMPOS_INTEIRO
         else pa.float64())
        for c in LAYOUT
    ]
    + [("arquivo_origem", pa.string())]
)


def ler_cabecalho(arq: Path, encoding: str) -> list:
    """Read the header line and return the stripped field names."""
    with open(arq, encoding=encoding) as f:
        return [c.strip() for c in f.readline().rstrip("\r\n").split(";")]


def detectar_encoding(arq: Path) -> str:
    """UTF-8 first; fall back to latin-1 if the header does not decode."""
    try:
        with open(arq, encoding="utf-8") as f:
            f.readline()
        return "utf-8"
    except UnicodeDecodeError:
        return "latin-1"


def converter(chunk: pd.DataFrame, origem: str) -> pd.DataFrame:
    """Strip padding and cast each field to its type."""
    chunk = chunk.apply(lambda s: s.str.strip())
    out = pd.DataFrame(index=chunk.index)
    for c in LAYOUT:
        col = chunk[c].replace("", None)
        if c in CAMPOS_TEXTO:
            out[COLUNAS[c]] = col
        elif c in CAMPOS_DATA:
            out[COLUNAS[c]] = pd.to_datetime(col, format="%Y%m%d", errors="coerce").dt.date
        elif c in CAMPOS_INTEIRO:
            out[COLUNAS[c]] = pd.to_numeric(col, errors="coerce").astype("Int32")
        else:
            out[COLUNAS[c]] = pd.to_numeric(col, errors="coerce")
    out["arquivo_origem"] = origem
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description="RNF system files (CE/RJ) -> parquet")
    ap.add_argument("--ref", required=True, help="reference month AAAA-MM")
    ap.add_argument("--nivel", required=True, choices=["AT", "BT"], help="voltage level")
    ap.add_argument("--uf", required=True, choices=["CE", "RJ"], help="distribution company")
    ap.add_argument("--entrada", help="glob pattern of input files (default: src/" + "<padrao do nivel>)")
    ap.add_argument("--chunk", type=int, default=CHUNK_ROWS, help="rows per chunk")
    args = ap.parse_args()

    padrao = args.entrada or str(DIR_SRC / PADRAO_ENTRADA[args.nivel])
    arquivos = sorted(Path(p) for p in glob.glob(padrao))
    if not arquivos:
        sys.exit(f"Nenhum arquivo encontrado: {padrao}")

    DIR_SAIDA.mkdir(parents=True, exist_ok=True)
    aaaamm = args.ref.replace("-", "")
    saida = DIR_SAIDA / f"{args.uf}_RNF_SISTEMA_{args.nivel}_{aaaamm}.parquet"

    # Running totals for the conference print
    tot = {"linhas": 0, "rnf_com_fator": 0.0, "valor_monetario": 0.0,
           "demanda": 0.0, "pis_cofins_dem": 0.0, "bandeira_com_trib": 0.0}
    clientes = set()

    writer = pq.ParquetWriter(saida, SCHEMA, compression="zstd")
    try:
        for arq in arquivos:
            enc = detectar_encoding(arq)
            cab = ler_cabecalho(arq, enc)
            if cab != LAYOUT:
                faltam = [c for c in LAYOUT if c not in cab]
                sobram = [c for c in cab if c not in LAYOUT]
                writer.close()
                saida.unlink(missing_ok=True)
                sys.exit(f"Layout diferente do esperado em {arq.name}. Faltam: {faltam}. Sobram: {sobram}")

            print(f"Lendo {arq.name} ({arq.stat().st_size / 1e6:,.0f} MB, {enc})")
            leitor = pd.read_csv(arq, sep=";", dtype=str, encoding=enc, names=LAYOUT, header=0,
                                 keep_default_na=False, chunksize=args.chunk)
            for i, chunk in enumerate(leitor, start=1):
                df = converter(chunk, arq.name)
                writer.write_table(pa.Table.from_pandas(df, schema=SCHEMA, preserve_index=False))

                # Conference totals
                tot["linhas"] += len(df)
                tot["rnf_com_fator"] += df["rnf_com_fator"].sum()
                tot["valor_monetario"] += df["valor_monetario"].sum()
                tot["demanda"] += df["renda_nao_faturada_da_demanda_faturada_sem_encargos"].sum()
                tot["pis_cofins_dem"] += df["pis_cofins_referente_a_renda_nao_faturada_da_demanda_faturada"].sum()
                tot["bandeira_com_trib"] += df["bandeira_tarifaria_com_tributo"].sum()
                clientes.update(df["numero_cliente"].dropna().unique())
                print(f"  bloco {i}: {tot['linhas']:,} linhas acumuladas")
    finally:
        writer.close()

    # Conference: energy in MWh, values in MR$
    print("\nGravado:", saida)
    print(f"Linhas: {tot['linhas']:,} | clientes distintos: {len(clientes):,}"
          + ("" if len(clientes) == tot["linhas"] else "  ATENCAO: cliente repetido"))
    print(f"Energia RNF com fator: {tot['rnf_com_fator'] / 1e3:,.0f} MWh")
    print(f"Valor da energia: MR$ {tot['valor_monetario'] / 1e6:,.2f}")
    print(f"Demanda (sem encargos): MR$ {tot['demanda'] / 1e6:,.2f}")
    print(f"PIS/COFINS da demanda: MR$ {tot['pis_cofins_dem'] / 1e6:,.2f}")
    print(f"Bandeira (com tributo): MR$ {tot['bandeira_com_trib'] / 1e6:,.2f}")


if __name__ == "__main__":
    main()
