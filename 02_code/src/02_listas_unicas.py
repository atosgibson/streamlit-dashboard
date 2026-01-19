import pandas as pd
from pathlib import Path

DATA = Path("01_data/02_interim/ATOS_clean.csv")
OUT = Path("04_docs/escopo")
OUT.mkdir(parents=True, exist_ok=True)

df = pd.read_csv(DATA)

servicos = sorted(df["SERVIÇO"].dropna().astype(str).str.strip().unique())
status = sorted(df["STATUS"].dropna().astype(str).str.strip().unique())

# salva em txt pra você ver fácil
(OUT / "lista_servicos.txt").write_text("\n".join(servicos), encoding="utf-8")
(OUT / "lista_status.txt").write_text("\n".join(status), encoding="utf-8")

print("OK! Salvei:")
print(" -", (OUT / "lista_servicos.txt").resolve())
print(" -", (OUT / "lista_status.txt").resolve())

print("\nSERVIÇOS (primeiros 30):")
print(servicos[:30])
print("\nSTATUS (todos):")
print(status)
