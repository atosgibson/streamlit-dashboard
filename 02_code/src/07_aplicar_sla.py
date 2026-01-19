import pandas as pd
from pathlib import Path
import unicodedata

DATA = Path("01_data/03_processed/atos_enriquecido.csv")
SLA_CAD = Path("04_docs/escopo/cadastro_sla.csv")
OUT = Path("01_data/04_exports")
OUT.mkdir(parents=True, exist_ok=True)

def norm(x: str) -> str:
    if x is None:
        return ""
    s = str(x).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = " ".join(s.split())
    return s

# carrega base
df = pd.read_csv(DATA)
df["CLIENTE"] = df["CLIENTE"].astype(str).str.strip()
df["tipo_servico"] = df["tipo_servico"].astype(str).str.strip()
df["sla_dias"] = pd.to_numeric(df["sla_dias"], errors="coerce")

# carrega cadastro SLA
sla = pd.read_csv(SLA_CAD)
sla["cliente"] = sla["cliente"].astype(str).str.strip()
sla["tipo_servico"] = sla["tipo_servico"].astype(str).str.strip()
sla["sla_dias"] = pd.to_numeric(sla["sla_dias"], errors="coerce")

# cria chaves normalizadas
df["k_cli"] = df["CLIENTE"].apply(norm)
df["k_tipo"] = df["tipo_servico"].apply(norm)

sla["k_cli"] = sla["cliente"].apply(norm)
sla["k_tipo"] = sla["tipo_servico"].apply(norm)

# separa regras específicas e padrão (*)
sla_especifico = sla[sla["cliente"] != "*"].copy()
sla_padrao = sla[sla["cliente"] == "*"].copy()

# 1) tenta casar regra específica (cliente + tipo)
df = df.merge(
    sla_especifico[["k_cli", "k_tipo", "sla_dias"]].rename(columns={"sla_dias": "sla_meta_dias"}),
    on=["k_cli", "k_tipo"],
    how="left"
)

# 2) onde não casou, aplica regra padrão por tipo (* + tipo)
df = df.merge(
    sla_padrao[["k_tipo", "sla_dias"]].rename(columns={"sla_dias": "sla_meta_padrao"}),
    on="k_tipo",
    how="left"
)

df["sla_meta_dias"] = df["sla_meta_dias"].fillna(df["sla_meta_padrao"])
df = df.drop(columns=["sla_meta_padrao"])

# resultado SLA
def sla_result(row):
    if pd.isna(row["sla_dias"]) or pd.isna(row["sla_meta_dias"]):
        return "SEM_DADO"
    return "DENTRO" if row["sla_dias"] <= row["sla_meta_dias"] else "FORA"

df["sla_resultado"] = df.apply(sla_result, axis=1)

# export 1: detalhe (se quiser depois usar no dashboard)
df_out = df.drop(columns=["k_cli", "k_tipo"], errors="ignore")
df_out.to_csv(OUT / "atos_com_sla.csv", index=False)

# export 2: resumo por tipo_servico
sla_por_tipo = (
    df_out.groupby("tipo_servico")
    .agg(
        qtd=("OS", "count"),
        sla_media=("sla_dias", "mean"),
        sla_mediana=("sla_dias", "median"),
        meta_media=("sla_meta_dias", "mean"),
        dentro=("sla_resultado", lambda s: (s == "DENTRO").sum()),
        fora=("sla_resultado", lambda s: (s == "FORA").sum()),
        sem_dado=("sla_resultado", lambda s: (s == "SEM_DADO").sum()),
    )
    .reset_index()
    .sort_values("qtd", ascending=False)
)
sla_por_tipo.to_csv(OUT / "kpi_sla_por_tipo_regras.csv", index=False)

# export 3: resumo por cliente
sla_por_cliente = (
    df_out.groupby("CLIENTE")
    .agg(
        qtd=("OS", "count"),
        dentro=("sla_resultado", lambda s: (s == "DENTRO").sum()),
        fora=("sla_resultado", lambda s: (s == "FORA").sum()),
        sem_dado=("sla_resultado", lambda s: (s == "SEM_DADO").sum()),
    )
    .reset_index()
    .sort_values("qtd", ascending=False)
)
sla_por_cliente.to_csv(OUT / "kpi_sla_por_cliente.csv", index=False)

print("OK! SLA aplicado.")
print("Gerados:")
print(" - atos_com_sla.csv")
print(" - kpi_sla_por_tipo_regras.csv")
print(" - kpi_sla_por_cliente.csv")
