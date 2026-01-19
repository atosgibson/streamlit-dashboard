import pandas as pd
from pathlib import Path

DATA = Path("01_data/03_processed/atos_enriquecido.csv")
OUT = Path("01_data/04_exports")
OUT.mkdir(parents=True, exist_ok=True)

df = pd.read_csv(DATA)

# datas
df["AUTORIZAÇÃO"] = pd.to_datetime(df["AUTORIZAÇÃO"], errors="coerce")
df["TÉRMINO"] = pd.to_datetime(df["TÉRMINO"], errors="coerce")

# mês de autorização
df["mes_autorizacao"] = df["AUTORIZAÇÃO"].dt.to_period("M").astype(str)

# 1) Chamadas por cliente (CHAMADO único)
chamadas_cliente = (
    df.dropna(subset=["CLIENTE", "CHAMADO"])
      .groupby("CLIENTE")["CHAMADO"]
      .nunique()
      .reset_index(name="qtd_chamados")
      .sort_values("qtd_chamados", ascending=False)
)
chamadas_cliente.to_csv(OUT / "kpi_chamadas_por_cliente.csv", index=False)

# 2) Demandas por mês
demandas_mes = (
    df.dropna(subset=["mes_autorizacao"])
      .groupby("mes_autorizacao")["OS"]
      .count()
      .reset_index(name="qtd_linhas")
      .sort_values("mes_autorizacao")
)
demandas_mes.to_csv(OUT / "kpi_demandas_por_mes.csv", index=False)

# 3) Financeiro (receita por billing_status)
financeiro_status = (
    df.groupby("billing_status")["receita"]
      .sum()
      .reset_index(name="receita_total")
      .sort_values("receita_total", ascending=False)
)
financeiro_status.to_csv(OUT / "kpi_financeiro_por_status.csv", index=False)

# Pendências (as 2 principais)
pendente_faturamento = df.loc[df["billing_status"] == "PENDENTE_FATURAMENTO", "receita"].sum()
faturado_pendente = df.loc[df["billing_status"] == "FATURADO_PENDENTE", "receita"].sum()
total_receita = df["receita"].sum()

resumo_fin = pd.DataFrame([{
    "receita_total": total_receita,
    "pendente_faturamento": pendente_faturamento,
    "faturado_pendente_receber": faturado_pendente,
}])
resumo_fin.to_csv(OUT / "kpi_financeiro_resumo.csv", index=False)

# 4) SLA (provisório): meta padrão 30 dias (depois vamos criar cadastro de SLAs)
SLA_PADRAO_DIAS = 30
df["sla_dias"] = pd.to_numeric(df["sla_dias"], errors="coerce")

df["sla_resultado"] = df["sla_dias"].apply(lambda x: "SEM_DADO" if pd.isna(x) else ("DENTRO" if x <= SLA_PADRAO_DIAS else "FORA"))

sla_por_tipo = (
    df.groupby("tipo_servico")
      .agg(
          qtd=("OS", "count"),
          sla_media=("sla_dias", "mean"),
          sla_mediana=("sla_dias", "median"),
          dentro=("sla_resultado", lambda s: (s == "DENTRO").sum()),
          fora=("sla_resultado", lambda s: (s == "FORA").sum()),
          sem_dado=("sla_resultado", lambda s: (s == "SEM_DADO").sum()),
      )
      .reset_index()
      .sort_values("qtd", ascending=False)
)
sla_por_tipo.to_csv(OUT / "kpi_sla_por_tipo.csv", index=False)

print("OK! KPIs exportados em:", OUT.resolve())
print("Arquivos gerados:")
for p in sorted(OUT.glob("kpi_*.csv")):
    print(" -", p.name)
