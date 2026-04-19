"""
sanitizar_dados.py
==================
Sanitização completa do dataset consolidado de imóveis.

Etapas:
  1. Simplificação dos tipos de imóveis
  2. Vagas NULL → 0
  3. Normalização e correção de bairros (mapa explícito + fuzzy matching)
  4. Criação da variável 'regiao'
  5. Colunas binárias de amenidades
  6. Filtragem final

Entrada:  imoveis_consolidado.csv
Saída:    imoveis_sanitizado.csv / imoveis_sanitizado.xlsx

Como rodar:
    pip install rapidfuzz
    python sanitizar_dados.py
"""

import os
import unicodedata
from collections import Counter
import pandas as pd
import numpy as np
from rapidfuzz import process, fuzz

PASTA     = os.path.dirname(os.path.abspath(__file__))
ENTRADA   = os.path.join(PASTA, "imoveis_consolidado.csv")
SAIDA_CSV = os.path.join(PASTA, "imoveis_sanitizado.csv")
SAIDA_XLSX= os.path.join(PASTA, "imoveis_sanitizado.xlsx")

# ─── Helpers ──────────────────────────────────────────────────────────────────

def remover_acentos(texto):
    if pd.isna(texto):
        return ""
    return (unicodedata.normalize("NFD", str(texto))
            .encode("ascii", "ignore").decode("utf-8").lower().strip())

def normalizar_cidade(cidade):
    return remover_acentos(str(cidade)).replace("-", " ").strip()

# ─── 0. Carrega ───────────────────────────────────────────────────────────────

df = pd.read_csv(ENTRADA, encoding="utf-8-sig", low_memory=False)
print(f"Carregado: {len(df):,} imóveis\n")

# ─── 1. Simplificação de tipos ────────────────────────────────────────────────
print("=" * 60)
print("ETAPA 1: Simplificação de tipos")
print("=" * 60)

MAPA_TIPOS = {
    "apartamento":              "Apartamento",
    "apartamento garden":       "Apartamento",
    "apartamento duplex":       "Apartamento",
    "duplex":                   "Apartamento",
    "estúdio":                  "Apartamento",
    "studio":                   "Apartamento",
    "flat":                     "Apartamento",
    "jk/kitnet":                "Apartamento",
    "kitnet":                   "Apartamento",
    "kitnet/conjugado":         "Apartamento",
    "cobertura":                "Apartamento",
    "casa":                     "Casa",
    "casa duplex":              "Casa",
    "casa em condomínio":       "Casa",
    "casa geminada":            "Casa",
    "casa geminada/sobrados":   "Casa",
    "casa sobrado":             "Casa",
    "sobrado":                  "Casa",
    "condomínio fechado":       "Casa",
}

antes = df["tipo"].value_counts().to_dict()
df["tipo"] = df["tipo"].str.strip().str.lower().map(MAPA_TIPOS).fillna(df["tipo"])

print("Antes → Depois:")
for tipo_orig, qtd in sorted(antes.items(), key=lambda x: -x[1]):
    novo = MAPA_TIPOS.get(str(tipo_orig).lower(), tipo_orig)
    print(f"  {tipo_orig:<30} → {novo} ({qtd})")
print(f"\nTipos após simplificação:")
print(df["tipo"].value_counts().to_string())

# ─── 2. Vagas NULL → 0 ────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("ETAPA 2: Vagas NULL → 0")
print("=" * 60)

nulos_vagas = df["vagas"].isna().sum()
df["vagas"] = df["vagas"].fillna(0).astype(int)
print(f"  {nulos_vagas} valores nulos preenchidos com 0")

# ─── 3. Normalização de bairros ───────────────────────────────────────────────
print("\n" + "=" * 60)
print("ETAPA 3: Normalização de bairros")
print("=" * 60)

# Todas as chaves devem estar em minúsculo sem acento (saída de remover_acentos)
MAPA_BAIRROS_EXPLICITO = {
    # ── Florianópolis ─────────────────────────────────────────────────────────
    "abraao":                                   ("Abraão",                      None),
    "acores":                                   ("Açores",                      None),
    "agronomica":                               ("Agronômica",                  None),
    "alto ribeirao":                            ("Alto Ribeirão",               None),
    "barra da lagoa":                           ("Barra da Lagoa",              None),
    "cacupe":                                   ("Cacupé",                      None),
    "morro das pedras":                         ("Morro das Pedras",            None),
    "ponta das canas":                          ("Ponta das Canas",             None),
    "santo antonio de lisboa":                  ("Santo Antônio de Lisboa",     None),
    "vargem do bom jesus":                      ("Vargem do Bom Jesus",         None),
    "vargem grande":                            ("Vargem Grande",               None),
    "vargem pequena":                           ("Vargem Pequena",              None),
    "ribeirao da ilha":                         ("Ribeirão da Ilha",            None),
    "saco dos limoes":                          ("Saco dos Limões",             None),
    "pantano do sul":                           ("Pântano do Sul",              None),
    "balneario dos acores":                     ("Açores",                      None),
    "baleario dos acores":                      ("Açores",                      None),
    "armacao":                                  ("Armação",                     None),
    "armacao do pantano do sul":                ("Armação",                     None),
    "armacao do pantano do sul ens brito":      ("Armação",                     None),
    "armacao do pantano do sul":                ("Armação",                     None),
    "armacao do pantano do sul":                ("Armação",                     None),
    "balneario do estreito":                    ("Estreito",                    None),
    "beiramar":                                 ("Centro",                      None),
    "beira mar":                                ("Centro",                      None),
    "beira mar norte":                          ("Centro",                      None),
    "canajure":                                 ("Canasvieiras",                None),
    "canto da lagoa":                           ("Canto da Lagoa",              None),
    "corrego grande":                           ("Córrego Grande",              None),
    "corrego grande":                           ("Córrego Grande",              None),
    "costeira do pirajubae":                    ("Costeira do Pirajubaé",       None),
    "ingleses centro":                          ("Ingleses",                    None),
    "ingleses norte":                           ("Ingleses",                    None),
    "ingleses sul":                             ("Ingleses",                    None),
    "ingleses florianopolis ag 4 cod te1470":   ("Ingleses",                    None),
    "ingleses do rio vermelho":                 ("Rio Vermelho",                None),
    "jurere":                                   ("Jurerê",                      None),
    "jurere internacional":                     ("Jurerê Internacional",        None),
    "lagoa da conceicao":                       ("Lagoa da Conceição",          None),
    "lagoa":                                    ("Lagoa da Conceição",          None),
    "novo campeche":                            ("Campeche",                    None),
    "campeche leste":                           ("Campeche",                    None),
    "campeche sul":                             ("Campeche",                    None),
    "praia forte":                              ("Forte",                       None),
    "praia do santinho":                        ("Santinho",                    None),
    "parque sao jorge":                         ("Saco Grande",                 None),
    "ponta das  canas":                         ("Ponta das Canas",             None),
    "ressacada":                                ("Ressacada",                   None),
    "rio tavares central":                      ("Rio Tavares",                 None),
    "rio tavares do norte":                     ("Rio Tavares",                 None),
    "sao joao do rio vermelho":                 ("Rio Vermelho",                None),
    "sao joao do rio vermelho":                 ("Rio Vermelho",                None),
    "santo antonio":                            ("Santo Antônio de Lisboa",     None),
    "tapera da base":                           ("Tapera da Base",              None),
    "agronomica florianopolis ag 55 cod ap10327": ("Agronômica",                None),
    "gaivotas":                                 ("Ingleses",                    None),
    "estrada sao roque":                        ("Estreito",                    "Florianópolis"), 
    # ── São José ──────────────────────────────────────────────────────────────
    "areias sao":                               ("Areias",                      "São José"),
    "barreiros sao":                            ("Barreiros",                   "São José"),
    "bosque das mansoes sao":                   ("Bosque das Mansões",          "São José"),
    "bosque sao":                               ("Bosque das Mansões",          "São José"),
    "campina sao":                              ("Campinas",                    "São José"),
    "bela vista sao":                           ("Bela Vista",                  "São José"),
    "campinas sao":                             ("Campinas",                    "São José"),
    "centro historico sao":                     ("Centro",                      "São José"),
    "centro sao":                               ("Centro",                      "São José"),
    "colonia santana sao":                      ("Colônia Santana",             "São José"),
    "fazenda santo antonio sao":                ("Fazenda Santo Antônio",       "São José"),
    "flor de napolis sao":                      ("Flor de Nápolis",             "São José"),
    "floresta sao":                             ("Floresta",                    "São José"),
    "forquilhas sao":                           ("Forquilhas",                  "São José"),
    "forquilhinha sao":                         ("Forquilhinhas",               "São José"),
    "forquilhinha":                             ("Forquilhinhas",               "São José"),
    "ipiranga sao":                             ("Ipiranga",                    "São José"),
    "jardim botanico sao":                      ("Jardim Botânico",             "São José"),
    "jardim cidade de florianopolis sao":       ("Jardim Cidade",               "São José"),
    "jardim cidade de florianopolis":           ("Jardim Cidade",               "São José"),
    "jardim marajoara":                         ("Jardim Marajoara",            "São José"),
    "kobrasol sao":                             ("Kobrasol",                    "São José"),
    "lisboa iii sao":                           (None,                          None),
    "mathias velho sao":                        (None,                          None),
    "nossa senhora do rosario sao":             ("Nossa Senhora do Rosário",    "São José"),
    "parque primavera sao":                     ("Parque Primavera",            "São José"),
    "picadas do sul sao":                       ("Picadas do Sul",              "São José"),
    "ponta de baixo sao":                       ("Ponta de Baixo",              "São José"),
    "potecas sao":                              ("Potecas",                     "São José"),
    "praia comprida sao":                       ("Praia Comprida",              "São José"),
    "real parque sao":                          ("Real Parque",                 "São José"),
    "rocado sao":                               ("Roçado",                      "São José"),
    "sao jose sao":                             (None,                          None),
    "sao luiz sao":                             ("São Luís",                    "São José"),
    "serraria sao":                             ("Serraria",                    "São José"),
    "sertao do imaruim sao":                    ("Sertão do Maruim",            "São José"),
    "sertao do maruim sao":                     ("Sertão do Maruim",            "São José"),
    "sertao do imaruim":                        ("Sertão do Maruim",            "São José"),
    "sertaozinho sao":                          ("Sertãozinho",                 "São José"),
    # ── Palhoça ───────────────────────────────────────────────────────────────
    "aririu":                                   ("Aririú",                      "Palhoça"),
    "barra do aririu":                          ("Barra do Aririú",             "Palhoça"),  # chave em minúsculo sem acento
    "enseada do brito ens brito":               ("Enseada de Brito",            "Palhoça"),
    "passagem de maciambu ens brito":           ("Massiambu",                   "Palhoça"),
    "pinheira ens brito":                       ("Pinheira",                    "Palhoça"),
    "pinheira (ens brito)":                     ("Pinheira",                    "Palhoça"),
    "praia do meio ens brito":                  ("Praia do Meio",               "Palhoça"),
    "praia do sonho ens brito":                 ("Praia do Sonho",              "Palhoça"),
    "praia do sonho (ens brito)":               ("Praia do Sonho",              "Palhoça"),
    "passagem do massiambu":                    ("Massiambu",                   "Palhoça"),
    "nova palhoca":                             ("Centro",                      "Palhoça"),
    "parque residencial pagani":                ("Pagani",                      "Palhoça"),
    "sao sebastiao":                            ("São Sebastião",               "Palhoça"),
    "sao sebastiao":                            ("São Sebastião",               "Palhoça"),
    "aririu da formiga":                        ("Aririú",                      "Palhoça"),
    "alto aririu":                              ("Alto Aririú",                 "Palhoça"),
    "sertao do imaruim":                        ("Sertão do Imaruim",           "Palhoça"),
    # ── Biguaçu ───────────────────────────────────────────────────────────────
    "areias de cima guaporanga":                ("Areias de Cima",              "Biguaçu"),
    "cachoeiras guaporanga":                    ("Cachoeiras",                  "Biguaçu"),
    "sao miguel guaporanga":                    ("São Miguel",                  "Biguaçu"),
    "tijuquinhas guaporanga":                   ("Tijuquinhas",                 "Biguaçu"),
    "tijuquinhas (guaporanga)":                 ("Tijuquinhas",                 "Biguaçu"),
    "area rural de biguacu":                    ("Área Rural de Biguaçu",       "Biguaçu"),
    "sorocaba de dentro":                       ("Área Rural de Biguaçu",       "Biguaçu"), 
    "tres riachos":                             ("Área Rural de Biguaçu",       "Biguaçu"),
    "jardim":                                   ("Jardim Janaína",              "Biguaçu"),
    "santa catarina":                           ("Área Rural de Biguaçu",       "Biguaçu"),
    "estrada são roque":                        ("Área Rural de Biguaçu",       "Biguaçu"), 
    # ── Múltiplas cidades ─────────────────────────────────────────────────────
    "nossa senhora do rosario":                 ("Nossa Senhora do Rosário",    None),
    "rocado":                                   ("Roçado",                      None),
    "guarda do cubatao":                        ("Guarda do Cubatão",           None),
    "fazenda santo antonio":                    ("Fazenda Santo Antônio",       None),
    "sertao do maruim":                         ("Sertão do Maruim",            None),
}

BAIRROS_FLORIANOPOLIS = [
    "Agronômica", "Centro", "Itacorubi", "João Paulo", "Monte Verde",
    "Pantanal", "Prainha", "Santa Mônica", "Saco dos Limões", "Saco Grande",
    "Córrego Grande", "Trindade", "José Mendes", "Carvoeira", "Serrinha",
    "Abraão", "Balneário", "Bom Abrigo", "Capoeiras", "Canto",
    "Coqueiros", "Estreito", "Itaguaçu", "Jardim Atlântico", "Monte Cristo",
    "Coloninha", "Ilhota", "Sapé",
    "Cachoeira do Bom Jesus", "Canasvieiras", "Daniela", "Ingleses",
    "Jurerê", "Jurerê Internacional", "Ponta das Canas", "Praia Brava",
    "Forte", "Ratones", "Sambaqui", "Santinho",
    "Santo Antônio de Lisboa", "Vargem do Bom Jesus", "Vargem Grande",
    "Vargem Pequena", "Cacupé",
    "Barra da Lagoa", "Costa da Lagoa", "Lagoa da Conceição",
    "Moçambique", "Rio Vermelho", "São João do Rio Vermelho", "Praia Mole",
    "Alto Ribeirão", "Armação", "Campeche", "Carianos", "Costeira do Pirajubaé",
    "Morro das Pedras", "Pântano do Sul", "Ribeirão da Ilha",
    "Rio Tavares", "Tapera", "Açores",
]

BAIRROS_SAO_JOSE = [
    "Areias", "Barreiros", "Bela Vista", "Campinas", "Centenário",
    "Cidade Jardim", "Colônia Santana", "Cruzeiro", "Dona Donata",
    "Espinheiros", "Flor de Nápolis", "Forquilhas", "Forquilhinhas",
    "Ipiranga", "Jardim Santiago", "Kobrasol", "Nossa Senhora do Rosário",
    "Pedregal", "Picadas do Sul", "Potecas", "Praia Comprida", "Real Parque",
    "Roçado", "Santa Terezinha", "Serraria", "Sertão do Maruim",
    "Ponta de Baixo", "Bosque das Mansões", "Floresta",
]

BAIRROS_PALHOCA = [
    "Aririú", "Alto Aririú", "Bela Vista", "Brejaru", "Caminho Novo",
    "Dom Bosco", "Enseada de Brito", "Guarda do Cubatão", "Jardim Eldorado",
    "Madri", "Massiambu", "Pagani", "Parque Residencial Pagani",
    "Passa Vinte", "Pedra Branca", "Pinheira", "Ponte do Imaruim",
    "Porto Grande", "Praia de Fora", "Praia do Sonho", "Rio Grande",
    "Santa Regina", "São Sebastião", "Barra do Aririú",
]

BAIRROS_BIGUACU = [
    "Areias de Cima", "Beira Rio", "Boa Vista", "Bom Viver", "Cachoeiras", "Centro",
    "Deltaville", "Encruzilhada", "Fundos", "Jardim Janaína", "Mar das Pedras",
    "Morro da Bina", "Prado de Baixo", "Praia João Rosa", "Rio Caveiras",
    "São Miguel", "Saudade", "Tijuquinhas", "Universitário", "Vendaval",
    "Área Rural de Biguaçu",
]

BAIRROS_POR_CIDADE = {
    "florianopolis": BAIRROS_FLORIANOPOLIS,
    "sao jose":      BAIRROS_SAO_JOSE,
    "palhoca":       BAIRROS_PALHOCA,
    "biguacu":       BAIRROS_BIGUACU,
}

def corrigir_bairro_cidade(bairro, cidade):
    """Primeiro tenta mapa explícito, depois fuzzy matching."""
    if pd.isna(bairro) or str(bairro).strip() in ("", "nan"):
        return bairro, cidade
    chave = remover_acentos(str(bairro)).lower().strip()
    # 1. Mapa explícito
    if chave in MAPA_BAIRROS_EXPLICITO:
        novo_bairro, nova_cidade = MAPA_BAIRROS_EXPLICITO[chave]
        return (novo_bairro if novo_bairro else None), (nova_cidade if nova_cidade else cidade)
    # 2. Fuzzy matching na lista da cidade
    cidade_norm = normalizar_cidade(cidade)
    lista = BAIRROS_POR_CIDADE.get(cidade_norm, [])
    if lista:
        resultado = process.extractOne(bairro, lista, scorer=fuzz.token_sort_ratio)
        if resultado and resultado[1] >= 85:
            return resultado[0], cidade
    return bairro, cidade

print("  Aplicando correção de bairros (mapa explícito + fuzzy)...")
df["bairro_original"] = df["bairro"]
df["cidade_original"] = df["cidade"]

resultados = df.apply(
    lambda row: corrigir_bairro_cidade(row["bairro"], row["cidade"]), axis=1
)
df["bairro"] = [r[0] for r in resultados]
df["cidade"] = [r[1] for r in resultados]

corrigidos_b = df[df["bairro"] != df["bairro_original"]][["bairro_original", "bairro", "cidade"]]
corrigidos_c = df[df["cidade"] != df["cidade_original"]][["bairro_original", "cidade_original", "cidade"]]
print(f"  {len(corrigidos_b)} bairros corrigidos")
print(f"  {len(corrigidos_c)} cidades corrigidas")
print("\n  Top correções de bairro:")
print(corrigidos_b.value_counts().head(30).to_string())
if len(corrigidos_c):
    print("\n  Correções de cidade:")
    print(corrigidos_c.value_counts().head(20).to_string())

df.drop(columns=["bairro_original", "cidade_original"], inplace=True)

# Corrige grafia de cidades (caso venham sem acento do CSV)
df["cidade"] = df["cidade"].replace({
    "Biguacu":  "Biguaçu",
    "Palhoca":  "Palhoça",
    "Sao Jose": "São José",
})

antes_remocao = len(df)
df = df[df["bairro"].notna()]
print(f"\n  {antes_remocao - len(df)} imóveis com bairro inválido removidos")

# ─── 4. Região ────────────────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("ETAPA 4: Criação da variável 'regiao'")
print("=" * 60)

REGIAO_FLORIPA = {
    # Central (Distrito Sede — Lei 5504)
    "Agronômica": "Central", "Carvoeira": "Central", "Centro": "Central",
    "Córrego Grande": "Central", "Costeira do Pirajubaé": "Central",
    "Itacorubi": "Central", "João Paulo": "Central", "José Mendes": "Central",
    "Monte Verde": "Central", "Pantanal": "Central", "Prainha": "Central",
    "Saco dos Limões": "Central", "Saco Grande": "Central",
    "Santa Mônica": "Central", "Trindade": "Central",
    # Continental (distritos Estreito + Coqueiros)
    "Abraão": "Continental", "Balneário": "Continental", "Bom Abrigo": "Continental",
    "Canto": "Continental", "Capoeiras": "Continental", "Coloninha": "Continental",
    "Coqueiros": "Continental", "Estreito": "Continental", "Itaguaçu": "Continental",
    "Jardim Atlântico": "Continental", "Monte Cristo": "Continental",
    # Norte (distritos Cachoeira do Bom Jesus + Canasvieiras + Ratones + Santo Antônio de Lisboa)
    "Cachoeira do Bom Jesus": "Norte", "Cacupé": "Norte", "Canajurê": "Norte",
    "Canasvieiras": "Norte", "Daniela": "Norte", "Forte": "Norte",
    "Ingleses": "Norte", "Ingleses Centro": "Norte", "Ingleses Norte": "Norte",
    "Ingleses Sul": "Norte", "Jurerê": "Norte", "Jurerê Internacional": "Norte",
    "Ponta das Canas": "Norte", "Praia Brava": "Norte", "Ratones": "Norte",
    "Sambaqui": "Norte", "Santinho": "Norte", "Santo Antônio de Lisboa": "Norte",
    "Vargem do Bom Jesus": "Norte", "Vargem Grande": "Norte", "Vargem Pequena": "Norte",
    # Leste (distritos Barra da Lagoa + Lagoa da Conceição + São João do Rio Vermelho)
    "Barra da Lagoa": "Leste", "Canto da Lagoa": "Leste",
    "Lagoa da Conceição": "Leste", "Praia Mole": "Leste",
    "Rio Vermelho": "Leste", "São João do Rio Vermelho": "Leste",
    # Sul (distritos Campeche + Pântano do Sul + Ribeirão da Ilha + Tapera da Base)
    "Açores": "Sul", "Alto Ribeirão": "Sul", "Armação": "Sul",
    "Campeche": "Sul", "Campeche Central": "Sul", "Campeche Leste": "Sul",
    "Campeche Sul": "Sul", "Carianos": "Sul", "Morro das Pedras": "Sul",
    "Pântano do Sul": "Sul", "Ressacada": "Sul", "Ribeirão da Ilha": "Sul",
    "Rio Tavares": "Sul", "Rio Tavares Central": "Sul",
    "Tapera": "Sul", "Tapera da Base": "Sul",
}

def atribuir_regiao(row):
    cidade = normalizar_cidade(str(row["cidade"]))
    if "florianopolis" in cidade or "florianópolis" in cidade:
        return REGIAO_FLORIPA.get(row["bairro"], "Outro")
    elif "sao jose" in cidade or "são josé" in cidade:
        return "São José"
    elif "palhoca" in cidade or "palhoça" in cidade:
        return "Palhoça"
    elif "biguacu" in cidade or "biguaçu" in cidade:
        return "Biguaçu"
    return str(row["cidade"]).strip()

df["regiao"] = df.apply(atribuir_regiao, axis=1)

print("  Distribuição por região:")
print(df["regiao"].value_counts().to_string())

outros_floripa = df[
    (df["regiao"] == "Outro") &
    (df["cidade"].str.lower().str.contains("florian", na=False))
]["bairro"].value_counts()
if len(outros_floripa) > 0:
    print(f"\n  ⚠ Bairros de Florianópolis sem região ({len(outros_floripa)} únicos):")
    print(outros_floripa.head(30).to_string())

df.to_csv(SAIDA_CSV, index=False, encoding="utf-8-sig")
print("\n  💾 Checkpoint salvo.")

# ─── 5. Colunas binárias de amenidades ───────────────────────────────────────
print("\n" + "=" * 60)
print("ETAPA 5: Colunas binárias de amenidades")
print("=" * 60)

AMENIDADES = [
    ("piscina",         ["piscina"]),
    ("sacada",          ["sacada", "varanda"]),
    ("churrasqueira",   ["churrasqueira"]),
    ("imovel_mobiliado",["mobiliado", "mobiliada"]),
    ("elevador",        ["elevador"]),
    ("academia",        ["academia"]),
    ("salao_de_festas", ["salão de festas", "salao de festas", "salão festas", "salao festas"]),
]

amenidades_lower = df["amenidades"].str.lower().fillna("")

for col, termos in AMENIDADES:
    mask = amenidades_lower.str.contains("|".join(termos), regex=True, na=False)
    df[col] = mask.astype(int)
    print(f"  {col:<25} {df[col].sum():>5} imóveis ({100*df[col].mean():.1f}%)")

df.to_csv(SAIDA_CSV, index=False, encoding="utf-8-sig")
print("\n  💾 Checkpoint salvo.")

# ─── 6. Filtros de limpeza ────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("ETAPA 6: Filtros de limpeza")
print("=" * 60)
antes = len(df)

# Valor > 50 milhões → erro de coleta (faltou dividir por 100)
mask_valor_alto = df["valor"] > 50_000_000
n_corrigidos = mask_valor_alto.sum()
df.loc[mask_valor_alto, "valor"] = df.loc[mask_valor_alto, "valor"] / 100
print(f"  {n_corrigidos} imóveis com valor > R$50M corrigidos (÷100)")

# Valor < 90.000 ou nulo
antes_filtro = len(df)
df = df[df["valor"].notna() & (df["valor"] >= 90_000)]
print(f"  {antes_filtro - len(df)} imóveis removidos: valor < R$90k ou nulo")

# Tamanho < 10 m² ou nulo
antes_filtro = len(df)
df = df[df["tamanho_m2"].notna() & (df["tamanho_m2"] >= 10)]
print(f"  {antes_filtro - len(df)} imóveis removidos: tamanho_m2 < 10 ou nulo")

# Tamanho > 1000 m²
antes_filtro = len(df)
df = df[df["tamanho_m2"] <= 1_000]
print(f"  {antes_filtro - len(df)} imóveis removidos: tamanho_m2 > 1.000 m²")

# Quartos >= 10 ou nula
antes_filtro = len(df)
df = df[df["quartos"].notna() & (df["quartos"] < 10)]
print(f"  {antes_filtro - len(df)} imóveis removidos: quartos >= 10")

# Banheiros nulo
antes_filtro = len(df)
df = df[df["banheiros"].notna()]
print(f"  {antes_filtro - len(df)} imóveis removidos: banheiros nulo")

# Vagas > 12
antes_filtro = len(df)
df = df[df["vagas"].isna() | (df["vagas"] <= 12)]
print(f"  {antes_filtro - len(df)} imóveis removidos: vagas > 12")

print(f"\n  Total removido: {antes - len(df)} imóveis ({antes:,} → {len(df):,})")

# ─── 7. Relatório final ───────────────────────────────────────────────────────
print("\n" + "=" * 60)
print("DATASET SANITIZADO")
print("=" * 60)
print(f"  Total de imóveis: {len(df):,}")

print(f"\n  Por tipo:")
print(df["tipo"].value_counts().to_string())

print(f"\n  Por cidade:")
print(df["cidade"].value_counts().to_string())

print(f"\n  Por região:")
print(df["regiao"].value_counts().to_string())

print(f"\n  Preenchimento geral:")
for col in ["bairro", "cidade", "tamanho_m2", "quartos", "banheiros",
            "vagas", "valor", "latitude", "longitude", "regiao"]:
    pct = 100 * df[col].notna().mean()
    print(f"    {col:<20} {pct:.1f}%")

# ─── 8. Salva ─────────────────────────────────────────────────────────────────
df.to_csv(SAIDA_CSV, index=False, encoding="utf-8-sig")
df.to_excel(SAIDA_XLSX, index=False)
print(f"\n✅ Salvo em:")
print(f"   {SAIDA_CSV}")
print(f"   {SAIDA_XLSX}")
