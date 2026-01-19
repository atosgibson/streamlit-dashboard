import pandas as pd
from pathlib import Path

RAW = Path("01_data/01_raw/ATOS.xlsx")

OUT_DOC = Path("04_docs/escopo")
OUT_DOC.mkdir(parents=True, exist_ok=True)

OUT_INTERIM = Path("01_data/02_interim")
OUT_INTERIM.mkdir(parents=True, exist_ok=True)


def main():
    # header=1 -> usa a 2ª linha do Excel como cabeçalho (onde está CLIENTE, PONTO, ...)
    df = pd.read_excel(RAW, header=1)

    # remove colunas Unnamed (lixo)
    df = df.loc[:, ~df.columns.astype(str).str.contains(r"^Unnamed")]

    # padroniza nomes das colunas
    df.columns = [str(c).strip() for c in df.columns]

    # força identificadores e textos como string (evita erro no parquet)
    text_cols = ["CLIENTE", "PONTO", "UF", "GESTOR", "CHAMADO", "OS", "SERVIÇO", "STATUS"]
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype("string")

    # tenta converter datas
    date_cols = ["AUTORIZAÇÃO", "TÉRMINO"]
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # qualquer coluna ainda "object" vira string (robusto pro pyarrow)
    obj_cols = df.select_dtypes(include=["object"]).columns
    for col in obj_cols:
        df[col] = df[col].astype("string")

    # salva versão limpa
    clean_csv = OUT_INTERIM / "ATOS_clean.csv"
    clean_parquet = OUT_INTERIM / "ATOS_clean.parquet"
    df.to_csv(clean_csv, index=False)
    df.to_parquet(clean_parquet, index=False)

    # gera dicionário de dados
    rows = []
    for col in df.columns:
        s = df[col]
        examples = [x for x in s.dropna().astype(str).head(3).tolist()]
        rows.append({
            "coluna": col,
            "tipo_pandas": str(s.dtype),
            "n_linhas": len(s),
            "n_vazios": int(s.isna().sum()),
            "exemplo_1": examples[0] if len(examples) > 0 else "",
            "exemplo_2": examples[1] if len(examples) > 1 else "",
            "exemplo_3": examples[2] if len(examples) > 2 else "",
        })

    dic = pd.DataFrame(rows).sort_values("coluna")
    dic_path = OUT_DOC / "dicionario_dados_ATOS.csv"
    dic.to_csv(dic_path, index=False)

    print("OK! ATOS_clean salvo em:", clean_csv.resolve())
    print("OK! ATOS_clean (parquet) salvo em:", clean_parquet.resolve())
    print("OK! Dicionário salvo em:", dic_path.resolve())
    print("\nColunas (limpas):", list(df.columns))
    print("\nAmostra (5 linhas):")
    print(df.head(5))


if __name__ == "__main__":
    main()
