import pandas as pd
from pathlib import Path

MAP_IN = Path("04_docs/escopo/mapeamento_servicos.csv")
MAP_OUT = Path("04_docs/escopo/mapeamento_servicos_autofill.csv")

df = pd.read_csv(MAP_IN)

def classify(service: str) -> tuple[str, str]:
    s = str(service).strip().lower()

    # Emergencial
    if "emergenc" in s:
        return "Emergencial", "Atendimento"

    # Laudos
    if "laudo" in s:
        return "Laudos", "Técnico"

    # Pintura
    if "pintura" in s:
        return "Pintura", "Manutenção"

    # Hidráulica (água/fria, esgoto, desentupimento, infiltração)
    if any(k in s for k in ["desentup", "tubula", " af", "af ", "infiltra"]):
        return "Hidráulica", "Manutenção"

    # Porta & Acesso (portas, fechaduras, dobradiças, mola aérea, gradil)
    if any(k in s for k in ["porta", "fechadura", "dobradi", "mola", "gradil"]):
        return "Portas & Acessos", "Manutenção"

    # Fachada / Revestimento
    if any(k in s for k in ["fachada", "acm"]):
        return "Fachada & Revestimento", "Reforma/Manutenção"

    # Coberta / Telhado
    if "coberta" in s or "telhado" in s or "tehado" in s:
        return "Coberta/Telhado", "Manutenção/Reforma"

    # Acessibilidade / Sinalização
    if "acessibil" in s:
        return "Acessibilidade", "Adequação"
    if "sinaliza" in s:
        return "Sinalização", "Adequação"

    # Civil (reforma/layout/parede/escada/elevador/cantoneira)
    if any(k in s for k in ["reforma", "layout", "parede", "escada", "elevador", "cantoneira"]):
        return "Civil", "Obra/Adequação"

    return "", ""


# Trata NaN como vazio
df["tipo_servico"] = df.get("tipo_servico", "").fillna("").astype(str)
df["categoria"] = df.get("categoria", "").fillna("").astype(str)

# Preenche apenas o que está vazio
new_tipo = []
new_cat = []

for _, row in df.iterrows():
    cur_tipo = str(row["tipo_servico"]).strip()
    cur_cat = str(row["categoria"]).strip()

    if cur_tipo != "":
        new_tipo.append(cur_tipo)
        new_cat.append(cur_cat)
        continue

    t, c = classify(row["SERVIÇO"])
    new_tipo.append(t)
    new_cat.append(c)

df["tipo_servico"] = new_tipo
df["categoria"] = new_cat

df.to_csv(MAP_OUT, index=False)

pendentes = df[df["tipo_servico"].fillna("").astype(str).str.strip() == ""]
print("OK! Salvei em:", MAP_OUT.resolve())
print("Serviços não classificados:", len(pendentes))
if len(pendentes) > 0:
    print(pendentes[["SERVIÇO"]].head(20).to_string(index=False))
