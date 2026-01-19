import pandas as pd
from pathlib import Path
import unicodedata

DATA = Path("01_data/02_interim/ATOS_clean.csv")

MAP_STATUS = Path("04_docs/escopo/mapeamento_status_financeiro.csv")
MAP_SERV = Path("04_docs/escopo/mapeamento_servicos_autofill.csv")

OUT_DIR = Path("01_data/03_processed")
OUT_DIR.mkdir(parents=True, exist_ok=True)

OUT = OUT_DIR / "atos_enriquecido.csv"


def norm_key(x: str) -> str:
    """Normaliza texto p/ chave: lowercase, sem acento, sem espaços duplos."""
    if x is None:
        return ""
    s = str(x).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))  # remove acentos
    s = " ".join(s.split())  # remove espaços extras
    return s


def main():
    df = pd.read_csv(DATA)

    # datas
    df["AUTORIZAÇÃO"] = pd.to_datetime(df["AUTORIZAÇÃO"], errors="coerce")
    df["TÉRMINO"] = pd.to_datetime(df["TÉRMINO"], errors="coerce")

    # SLA (dias)
    df["sla_dias"] = (df["TÉRMINO"] - df["AUTORIZAÇÃO"]).dt.days

    # receita
    df["receita"] = pd.to_numeric(df["RECEITA"], errors="coerce").fillna(0)

    # --- STATUS -> billing_status ---
    ms = pd.read_csv(MAP_STATUS)
    df["status_key"] = df["STATUS"].apply(norm_key)
    ms["status_key"] = ms["STATUS"].apply(norm_key)

    df = df.merge(ms[["status_key", "billing_status"]], on="status_key", how="left")

    # --- SERVIÇO -> tipo_servico/categoria ---
    mserv = pd.read_csv(MAP_SERV)

    # garante que a coluna existe (pode vir com espaços/encoding)
    # pega a primeira coluna que contenha "SERVI" no nome
    serv_col = None
    for c in mserv.columns:
        if "SERV" in str(c).upper():
            serv_col = c
            break
    if serv_col is None:
        raise ValueError(f"Não achei a coluna de SERVIÇO no mapeamento. Colunas: {mserv.columns.tolist()}")

    df["servico_key"] = df["SERVIÇO"].apply(norm_key)
    mserv["servico_key"] = mserv[serv_col].apply(norm_key)

    df = df.merge(
        mserv[["servico_key", "tipo_servico", "categoria"]],
        on="servico_key",
        how="left"
    )

    # preenchimentos
    df["billing_status"] = df["billing_status"].fillna("NAO_MAPEADO")
    df["tipo_servico"] = df["tipo_servico"].fillna("NAO_MAPEADO")
    df["categoria"] = df["categoria"].fillna("NAO_MAPEADO")

    # limpa chaves auxiliares
    df = df.drop(columns=["status_key", "servico_key"], errors="ignore")

    df.to_csv(OUT, index=False)

    print("OK! Base enriquecida salva em:", OUT.resolve())
    print("\nTipo de serviço (top 10):")
    print(df["tipo_servico"].value_counts(dropna=False).head(10))
    print("\nBilling status (contagem):")
    print(df["billing_status"].value_counts(dropna=False))


if __name__ == "__main__":
    main()
