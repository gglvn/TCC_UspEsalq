"""
unir_dados.py
=============
Une os CSVs das imobiliárias em um único arquivo consolidado.

Arquivos esperados na mesma pasta:
    imoveis_Biguaçu_SC.csv
    imoveis_florianopolis.csv
    imoveis_Palhoça_SC.csv
    imoveis_São_José_SC.csv
    imoveis_gralha_grande_florianopolis_corrigido.csv

Saída:
    imoveis_consolidado.csv
    imoveis_consolidado.xlsx

Como rodar:
    python unir_dados.py
"""

import os
import pandas as pd

PASTA = os.path.dirname(os.path.abspath(__file__))

ARQUIVOS = [
    "imoveis_Biguaçu_SC.csv",
    "imoveis_florianopolis.csv",
    "imoveis_Palhoça_SC.csv",
    "imoveis_São_José_SC.csv",
    "imoveis_gralha_grande_florianopolis_corrigido.csv",
]

COLUNAS_ESPERADAS = [
    "imobiliaria", "tipo", "rua", "bairro", "cidade",
    "tamanho_m2", "quartos", "banheiros", "vagas",
    "valor", "latitude", "longitude",
    "amenidades", "url",
]

# ─── Carrega e inspeciona ──────────────────────────────────────────────────────

print("=" * 60)
print("CARREGANDO ARQUIVOS")
print("=" * 60)

dfs = []
for arquivo in ARQUIVOS:
    caminho = os.path.join(PASTA, arquivo)
    if not os.path.exists(caminho):
        print(f"  ⚠ Não encontrado: {arquivo}")
        continue
    df = pd.read_csv(caminho, encoding="utf-8-sig", low_memory=False)
    print(f"  ✓ {arquivo}: {len(df)} linhas | colunas: {list(df.columns)}")
    dfs.append(df)

if not dfs:
    print("Nenhum arquivo encontrado. Verifique a pasta.")
    exit()

# ─── Une ──────────────────────────────────────────────────────────────────────

df = pd.concat(dfs, ignore_index=True)
print(f"\nTotal bruto após união: {len(df)} linhas")

# ─── Garante colunas padrão ───────────────────────────────────────────────────

for col in COLUNAS_ESPERADAS:
    if col not in df.columns:
        df[col] = None

df = df[COLUNAS_ESPERADAS]

# ─── Limpeza básica ───────────────────────────────────────────────────────────

# Remove duplicatas por URL
antes = len(df)
df = df.drop_duplicates(subset=["url"], keep="first")
print(f"Duplicatas removidas por URL: {antes - len(df)}")

# Remove linhas sem valor e sem m²  (provavelmente incompletas)
antes = len(df)
df = df[df["valor"].notna() | df["tamanho_m2"].notna()]
print(f"Linhas sem valor E sem m² removidas: {antes - len(df)}")

# Normaliza cidade (capitaliza, remove espaços extras)
df["cidade"] = df["cidade"].astype(str).str.strip().str.title()
df["bairro"] = df["bairro"].astype(str).str.strip().str.title()
df["tipo"]   = df["tipo"].astype(str).str.strip().str.title()

# Converte colunas numéricas
for col in ["tamanho_m2", "quartos", "banheiros", "vagas", "valor",
            "latitude", "longitude"]:
    df[col] = pd.to_numeric(df[col], errors="coerce")

# ─── Relatório ────────────────────────────────────────────────────────────────

print(f"\n{'='*60}")
print(f"DATASET CONSOLIDADO: {len(df)} imóveis")
print(f"{'='*60}")

print(f"\nPor imobiliária:")
print(df["imobiliaria"].value_counts().to_string())

print(f"\nPor cidade:")
print(df["cidade"].value_counts().to_string())

print(f"\nPor tipo:")
print(df["tipo"].value_counts().to_string())

print(f"\nPreenchimento dos campos:")
for col in COLUNAS_ESPERADAS:
    preenchido = df[col].notna().sum()
    pct = 100 * preenchido / len(df)
    print(f"  {col:<20} {preenchido:>6} ({pct:.1f}%)")

print(f"\nEstatísticas de valor (R$):")
print(df["valor"].describe().apply(lambda x: f"R$ {x:,.2f}").to_string())

# ─── Salva ────────────────────────────────────────────────────────────────────

saida_csv  = os.path.join(PASTA, "imoveis_consolidado.csv")
saida_xlsx = os.path.join(PASTA, "imoveis_consolidado.xlsx")

df.to_csv(saida_csv, index=False, encoding="utf-8-sig")
df.to_excel(saida_xlsx, index=False)

print(f"\n✅ Salvo em:")
print(f"   {saida_csv}")
print(f"   {saida_xlsx}")
