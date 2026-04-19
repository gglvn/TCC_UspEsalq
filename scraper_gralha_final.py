"""
Scraper - Gralha Imóveis (Grande Florianópolis)
================================================
Etapa 1: Chama diretamente a API JSON do site para coletar todos os imóveis

Etapa 2: Salva tudo em Excel (.xlsx) e CSV (.csv).

Como rodar:
    pip install requests pandas
    python scraper_gralha_final.py
"""

import re
import os
import json
import time
import requests
import pandas as pd

# ─── Configurações ─────────────────────────────────────────────────────────────

BASE_URL = "https://www.gralhaimoveis.com.br"
API_URL  = "https://www.gralhaimoveis.com.br/api/anuncios/search"

CIDADES = [
    {"tipoId":0,"tipo":"Cidades","titulo":"Florianópolis - SC","slug":"cidade+sc+florianopolis",
     "bairroId":0,"bairro":None,"cidadeId":2,"cidade":"Florianópolis","estadoSigla":"SC",
     "empreendimento":None,"condominio":None,"agrupamentoId":None,"agrupamento":None},
    {"tipoId":0,"tipo":"Cidades","titulo":"São José - SC","slug":"cidade+sc+sao-jose",
     "bairroId":0,"bairro":None,"cidadeId":4,"cidade":"São José","estadoSigla":"SC",
     "empreendimento":None,"condominio":None,"agrupamentoId":None,"agrupamento":None},
    {"tipoId":0,"tipo":"Cidades","titulo":"Palhoça - SC","slug":"cidade+sc+palhoca",
     "bairroId":0,"bairro":None,"cidadeId":6,"cidade":"Palhoça","estadoSigla":"SC",
     "empreendimento":None,"condominio":None,"agrupamentoId":None,"agrupamento":None},
    {"tipoId":0,"tipo":"Cidades","titulo":"Biguacu - SC","slug":"cidade+sc+biguacu",
     "bairroId":0,"bairro":None,"cidadeId":7,"cidade":"Biguacu","estadoSigla":"SC",
     "empreendimento":None,"condominio":None,"agrupamentoId":None,"agrupamento":None},
]

# MODO TESTE: Pra não ter a resposta do algoritmo apenas depois de muito tempo rodando, foi criada essa seção para teste com poucas páginas e anúncios.
# True: coleta 3 páginas da API e detalha 20 imóveis individualmente.
# False: coleta tudo.
MODO_TESTE    = True
PAGINAS_TESTE = 3
# 

TIPOS_ACEITOS = {"casa", "apartamento", "cobertura", "flat", "estúdio", "studio", "duplex", "kitnet"} # excluindo terrenos, chácaras, galpões, salas comerciais, etc.
PAUSA_API     = 0.3   # entre chamadas à API

PASTA_SCRIPT  = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_SAIDA = os.path.join(PASTA_SCRIPT, "imoveis_gralha_grande_florianopolis")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "pt-BR,pt;q=0.9",
    "Referer": "https://www.gralhaimoveis.com.br/",
}

# Funções auxiliares

def extrair_valor(v):
    if not v:
        return None
    if isinstance(v, (int, float)):
        valor = float(v)
        # Se vier em centavos (ex: 1247018688 = R$ 12.470.186,88)
        if valor > 50_000_000:
            return valor / 100
        return valor
    # fallback para strings formatadas
    s = re.sub(r'[R$\s]', '', str(v))
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s) if s else None
    except ValueError:
        return None

def extrair_m2(v):
    if not v:
        return None
    val = re.sub(r'\.(\d{3})(?!\d)', r'\1', str(v))
    val = val.replace(",", ".")
    try:
        return float(val)
    except ValueError:
        return None


def chamar_api(pagina, tentativas=3, pausa_erro=5):
    params = {
        "suggest":    json.dumps(CIDADES, ensure_ascii=False, separators=(",", ":")),
        "finalidade": "venda",
        "page":       pagina,
    }
    for tentativa in range(1, tentativas + 1):
        try:
            r = requests.get(API_URL, params=params, headers=HEADERS, timeout=20)
            r.raise_for_status()
            return r.json()
        except Exception as e:
            print(f"    ⚠ Erro na API página {pagina} (tentativa {tentativa}/{tentativas}): {e}")
            if tentativa < tentativas:
                time.sleep(pausa_erro)
    return None


# Etapa 1: Coletar dados via API

def coletar_dados_api():
    paginas_max = PAGINAS_TESTE if MODO_TESTE else 9999
    modo_label  = f"[MODO TESTE: {paginas_max} páginas]" if MODO_TESTE else "[todas as páginas]"

    print(f"\n{'='*60}")
    print(f"ETAPA 1: Coletando via API {modo_label}")
    print(f"{'='*60}")

    # Faz uma chamada de teste para inspecionar a estrutura do JSON
    print("  Verificando estrutura da API")
    data_teste = chamar_api(1)
    if not data_teste:
        print(" API não respondeu")
        return []

    print(f"  Campos retornados: {list(data_teste.keys()) if isinstance(data_teste, dict) else 'lista direta'}")

    # Detecta onde estão os itens
    if isinstance(data_teste, list):
        chave_itens = None
    elif isinstance(data_teste, dict):
        for chave in ["data", "anuncios", "results", "items", "imoveis", "listings"]:
            if chave in data_teste and isinstance(data_teste[chave], list):
                chave_itens = chave
                break
        else:
            chave_itens = None
            # Mostra o JSON bruto para diagnóstico
            print(f"  Estrutura desconhecida. Primeiros 500 chars:")
            print(f"  {str(data_teste)[:500]}")

    def extrair_itens(data):
        if isinstance(data, list):
            return data
        if chave_itens:
            return data.get(chave_itens, [])
        return []

    # Inspeciona campos do primeiro item
    itens_teste = extrair_itens(data_teste)
    if itens_teste:
        print(f"  Campos do primeiro item: {list(itens_teste[0].keys())}")

    todos  = []
    vistos = set()

    erros_consecutivos = 0
    for pagina in range(1, paginas_max + 1):
        if pagina > 1:  # página 1 já foi chamada
            data = chamar_api(pagina)
        else:
            data = data_teste

        if not data:
            erros_consecutivos += 1
            if erros_consecutivos >= 5:
                print(f"  5 erros consecutivos — encerrando coleta")
                break
            print(f" Pulando página {pagina} e continuando")
            continue
        erros_consecutivos = 0

        itens = extrair_itens(data)
        if not itens:
            print(f"  Página {pagina}: 0 itens. fim da listagem")
            break

        novos = 0
        for item in itens:
            # Filtra tipos de imóveis não desejados
            tipo_item = (item.get("tipo") or "").strip()
            if tipo_item.lower() not in TIPOS_ACEITOS:
                continue

            slug       = item.get("url") or ""
            imovel_id  = item.get("id") or ""
            imovel_url = f"{BASE_URL}/imovel/{slug}/{imovel_id}" if slug and imovel_id else None
            if not imovel_url or imovel_url in vistos:
                continue
            vistos.add(imovel_url)

            dados = {
                "url":         imovel_url,
                "imobiliaria": "Gralha Imóveis",
                "tipo":        tipo_item,
                "rua":         (f"{item['logradouro']}, {item['numero']}".strip(", ")
                               if item.get("logradouro") else None),
                "bairro":      item.get("bairro") or None,
                "cidade":      item.get("cidade") or None,
                "tamanho_m2":  extrair_m2(item.get("areaConstruida")),
                "quartos":     int(item["quartos"]) if item.get("quartos") else None,
                "banheiros":   int(item["banheiros"]) if item.get("banheiros") else None,
                "vagas":       int(item["vagas"]) if item.get("vagas") else None,
                "valor":       extrair_valor(item.get("valorVenda") or item.get("valorPromocional")),
                "amenidades":  ", ".join(item["caracteristicas"]) if item.get("caracteristicas") else None,
                "data_construcao": None,
                "latitude":        float(item["latitude"]) if item.get("latitude") else None,
                "longitude":       float(item["longitude"]) if item.get("longitude") else None,
            }

            todos.append(dados)
            novos += 1

        print(f"  Página {pagina}: {novos} novos (total: {len(todos)})")
        time.sleep(PAUSA_API)

    print(f"\nTotal coletado via API: {len(todos)} imóveis\n")
    return todos


# Etapa 2: Salvando dados

def salvar(lista_dados):
    df = pd.DataFrame(lista_dados)
    colunas = [
        "imobiliaria", "tipo", "rua", "bairro", "cidade",
        "tamanho_m2", "quartos", "banheiros", "vagas",
        "valor", "latitude", "longitude",
        "amenidades", "url",
    ]
    colunas_existentes = [c for c in colunas if c in df.columns]
    df = df[colunas_existentes]
    df.to_excel(f"{ARQUIVO_SAIDA}.xlsx", index=False)
    df.to_csv(f"{ARQUIVO_SAIDA}.csv",   index=False, encoding="utf-8-sig")
    print(f"  Salvo em: {os.path.abspath(ARQUIVO_SAIDA + '.xlsx')}")
    return df


# Função Principal

def main():
    print("\n🏠 SCRAPER - GRALHA IMÓVEIS (GRANDE FLORIANÓPOLIS)")
    print("=" * 60)

    lista_dados = coletar_dados_api()
    if not lista_dados:
        print("Nenhum imóvel encontrado")
        return

    print("Salvando resultado")
    df = salvar(lista_dados)

    print(f"\n{'='*60}")
    print(f"CONCLUÍDO - {len(df)} imóveis coletados")
    print(f"{'='*60}\n")

    colunas_preview = [c for c in ["tipo","bairro","cidade","tamanho_m2","quartos","banheiros","vagas","valor"] if c in df.columns]
    print(df[colunas_preview].head(10).to_string(index=False))


if __name__ == "__main__":
    main()
