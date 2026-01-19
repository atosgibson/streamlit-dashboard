import pandas as pd
from pathlib import Path

DATA = Path("01_data/02_interim/ATOS_clean.csv")
OUT = Path("04_docs/escopo/mapeamento_servicos.csv")

df = pd.read_csv(DATA)
servicos = (
    df["SERVIÇO"]
    .dropna()
    .astype(str)
    .str.strip()
    .drop_duplicates()
    .sort_values()
)

map_df = pd.DataFrame({
    "SERVIÇO": servicos,
    "tipo_servico": "",   # você vai preencher
    "categoria": ""       # opcional
})

map_df.to_csv(OUT, index=False)
print("OK! Arquivo criado em:", OUT.resolve())
print("Total de serviços:", len(map_df))
print("\nAmostra:")
print(map_df.head(10))
