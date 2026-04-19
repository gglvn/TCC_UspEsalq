"""
Scraper - Crédito Real Imóveis (Florianópolis)
================================================
Etapa 1: Coleta os dados básicos (tipo, rua, bairro, cidade, m², quartos,
         vagas, valor) diretamente dos cards de listagem.

Etapa 2: Entra em cada página individual com Playwright (navegador real)
         para pegar: banheiros, amenidades, data de construção, lat/lng.

Etapa 3: Salva tudo em Excel (.xlsx) e CSV (.csv).

Como rodar:
    pip install requests beautifulsoup4 pandas openpyxl playwright
    python -m playwright install chromium
    python scraper.py
"""

import re
import os
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

# ─── Configurações ─────────────────────────────────────────────────────────────

# ─── Cidade ────────────────────────────────────────────────────────────────────
# Mude apenas esta variável para trocar de cidade.
# Opções prontas:
#   "Florianópolis_SC"
#   "Biguaçu_SC"
#   "Palhoça_SC"
#   "São José_SC"
CIDADE = "Biguaçu_SC"
# ───────────────────────────────────────────────────────────────────────────────

BASE_URL = "https://www.creditoreal.com.br"

LISTAGEM_URL = (
    "https://www.creditoreal.com.br/vendas/residenciais"
    "?cityState=" + CIDADE.replace(" ", "+") + "&valueType=true&page={page}"
)

# Total de páginas da cidade escolhida (verifique no site antes de rodar)
# Florianópolis: 157 | Biguaçu: 14 | Palhoça: 36 | São José: 54
TOTAL_PAGINAS = 36  # valor alto como segurança — o scraper para sozinho quando não há mais imóveis

# ─── MODO TESTE ────────────────────────────────────────────────────────────────
# Com MODO_TESTE = True: coleta 2 páginas e detalha apenas 5 imóveis (~1 minuto).
# Quando tudo estiver certo, mude para False para rodar completo.
MODO_TESTE    = False
PAGINAS_TESTE = 2
IMOVEIS_TESTE = 5
# ───────────────────────────────────────────────────────────────────────────────

PAUSA_LISTAGEM = 0.5   # segundos entre páginas de listagem
PAUSA_IMOVEL   = 0.5   # segundos entre páginas individuais

# Salva os arquivos na mesma pasta onde está o scraper.py
PASTA_SCRIPT  = os.path.dirname(os.path.abspath(__file__))
CIDADE_SLUG   = CIDADE.replace(" ", "_").replace("/", "_")
ARQUIVO_SAIDA = os.path.join(PASTA_SCRIPT, f"imoveis_{CIDADE_SLUG}")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9",
}

# ─── Funções auxiliares ────────────────────────────────────────────────────────

def get_page(url, tentativas=3):
    for i in range(tentativas):
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            print(f"    ⚠ Tentativa {i+1}/{tentativas} falhou: {e}")
            time.sleep(3)
    print(f"    ✗ Falha definitiva: {url}")
    return None


def slug_para_texto(slug):
    """Converte 'praia-mole' → 'Praia Mole'"""
    return slug.replace("-", " ").title()


def extrair_bairro_da_url(url):
    """
    Extrai o bairro a partir do slug da URL.
    Ex: .../apartamento-em-centro-florianopolis-ag-280-cod-74081121 → 'Centro'
    Ex: .../casa-em-praia-mole-florianopolis-ag-280-cod-74063568    → 'Praia Mole'
    """
    slug = url.rstrip("/").split("/")[-1]
    slug = re.sub(r"-ag-\d+-cod-\d+$", "", slug)   # remove sufixo -ag-XXX-cod-XXX
    slug = re.sub(r"-[a-zA-Z]+$", "", slug)          # remove nome da cidade
    partes = slug.split("-em-")
    if len(partes) >= 2:
        return slug_para_texto(partes[-1])
    return None


def extrair_valor(texto):
    """'1.250.000' → 1250000.0"""
    if not texto:
        return None
    nums = re.sub(r"[^\d]", "", texto)
    return float(nums) if nums else None


# ─── Etapa 1: Coletar dados dos cards de listagem ─────────────────────────────

def coletar_dados_listagem():
    """
    Percorre as páginas de listagem e extrai de cada card:
    imobiliaria, tipo, rua, bairro, cidade, m², quartos, vagas, valor, url.
    """
    paginas    = PAGINAS_TESTE if MODO_TESTE else TOTAL_PAGINAS
    modo_label = f"[MODO TESTE: {paginas} páginas]" if MODO_TESTE else f"[{paginas} páginas]"

    print(f"\n{'='*60}")
    print(f"ETAPA 1: Coletando dados dos cards {modo_label}")
    print(f"{'='*60}")

    todos  = []
    vistos = set()

    for pagina in range(1, paginas + 1):
        url = LISTAGEM_URL.format(page=pagina)
        print(f"  Página {pagina}/{paginas}: {url}")

        resp = get_page(url)
        if not resp:
            continue

        soup  = BeautifulSoup(resp.text, "html.parser")
        cards = soup.find_all("a", href=re.compile(r"^/vendas/imovel/"))

        novos = 0
        for card in cards:
            href       = card.get("href", "")
            imovel_url = BASE_URL + href

            if imovel_url in vistos:
                continue
            vistos.add(imovel_url)

            h2 = card.find("h2")
            if not h2:
                continue

            texto_h2 = h2.get_text(" ", strip=True)

            dados = {
                "url":             imovel_url,
                "imobiliaria":     "Crédito Real",
                "tipo":            None,
                "rua":             None,
                "bairro":          None,
                "cidade":          None,
                "tamanho_m2":      None,
                "quartos":         None,
                "banheiros":       None,
                "vagas":           None,
                "valor":           None,
                "amenidades":      None,
                "data_construcao": None,
                "latitude":        None,
                "longitude":       None,
            }

# ── Tipo, Rua, Bairro, Cidade — lidos dos spans do h2 ─────────
            tipo_span   = h2.find("span", class_="imovel-type")
            rua_span    = h2.find("span", class_="wzvrd")
            local_span  = h2.find("span", class_="fwxJQc")   # "Bairro, Cidade"

            dados["tipo"] = tipo_span.get_text(strip=True)  if tipo_span  else None
            dados["rua"]  = rua_span.get_text(strip=True)   if rua_span   else None

            if local_span:
                local_txt = local_span.get_text(strip=True)  # ex: "Boa Vista, Biguaçu"
                if "," in local_txt:
                    bairro_txt, cidade_txt = local_txt.rsplit(",", 1)
                    dados["bairro"] = bairro_txt.strip()
                    dados["cidade"] = cidade_txt.strip()
                else:
                    dados["bairro"] = local_txt
                    dados["cidade"] = CIDADE.split("_")[0]
            else:
                dados["cidade"] = CIDADE.split("_")[0]

# ── Ignorar terrenos ───────────────────────────────────────────
            if dados["tipo"] and "terreno" in dados["tipo"].lower():
                continue

# ── m², quartos, vagas ─────────────────────────────────────────
            texto_card = card.get_text(" ", strip=True)

            m2 = re.search(r"(\d[\d.,]*)\s*m[²2]", texto_card)
            if m2:
                try:
                    dados["tamanho_m2"] = float(
                        m2.group(1).replace(".", "").replace(",", ".")
                    )
                except ValueError:
                    pass

            qts = re.search(r"(\d+)\s*quart[oa]s?", texto_card, re.I)
            if qts:
                dados["quartos"] = int(qts.group(1))

            vagas = re.search(r"(\d+)\s*vaga[s]?", texto_card, re.I)
            if vagas:
                dados["vagas"] = int(vagas.group(1))

            # ── Valor ──────────────────────────────────────────────────────
            match_de  = re.search(r"De:\s*R\$\s*([\d.,]+)", texto_card)
            match_por = re.search(r"Por:\s*R\$\s*([\d.,]+)", texto_card)
            match_rs  = re.search(r"R\$\s*([\d.,]+)", texto_card)

            if match_de and match_por:
                dados["valor"] = extrair_valor(match_por.group(1))
            elif match_rs:
                dados["valor"] = extrair_valor(match_rs.group(1))

            todos.append(dados)
            novos += 1

        print(f"    → {novos} novos (total: {len(todos)})")

        # Para automaticamente quando a página não tiver novos imóveis
        if novos == 0:
            print(f"  ℹ Nenhum imóvel novo na página {pagina} — encerrando coleta.")
            break

        time.sleep(PAUSA_LISTAGEM)

    print(f"\n✓ Total coletado da listagem: {len(todos)} imóveis\n")
    return todos


# ─── Etapa 2: Complementar com dados das páginas individuais ──────────────────

AMENIDADES_LISTA = [
    # Infraestrutura do condomínio
    "piscina", "elevador", "salão de festas", "academia", "churrasqueira",
    "playground", "portaria 24h", "portaria", "segurança 24h", "câmeras",
    "gerador", "sauna", "spa", "quadra", "espaço gourmet", "terraço",
    "coworking", "bicicletário", "jardim", "área de lazer",
    "salão de jogos", "espaço kids", "piscina infantil",
    # Tags que o site usa nas páginas individuais
    "aceita pet", "sem mobilia", "mobiliado", "semi mobiliado",
    "desocupado", "vista mar", "vista para o mar", "frente mar",
    "varanda gourmet", "varanda",
]


def complementar_com_playwright(lista_dados):
    """
    Abre as páginas individuais com Playwright e preenche:
    banheiros, amenidades, data_construcao, latitude, longitude.
    """
    imoveis    = lista_dados[:IMOVEIS_TESTE] if MODO_TESTE else lista_dados
    total      = len(imoveis)
    modo_label = f"[MODO TESTE: {total} imóveis]" if MODO_TESTE else f"[{total} imóveis]"

    print(f"ETAPA 2: Abrindo páginas individuais {modo_label}")
    print(f"  (Um navegador Chromium invisível será aberto — é normal)")
    print("=" * 60)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        pw_page = browser.new_page(extra_http_headers=HEADERS)

        for i, dados in enumerate(imoveis, start=1):
            url = dados["url"]
            print(f"  [{i}/{total}] {url}")

            try:
                pw_page.goto(url, wait_until="domcontentloaded", timeout=30000)
                time.sleep(1)
                html        = pw_page.content()
                soup        = BeautifulSoup(html, "html.parser")
                texto       = soup.get_text(" ", strip=True)
                texto_lower = texto.lower()

                # ── Banheiros ─────────────────────────────────────────────
                match = re.search(r"(\d+)\s*banheir[oa]s?", texto, re.I)
                if match:
                    dados["banheiros"] = int(match.group(1))

                # ── Amenidades ────────────────────────────────────────────
                encontradas = [a for a in AMENIDADES_LISTA if a in texto_lower]
                dados["amenidades"] = ", ".join(encontradas) if encontradas else None

                # ── Data de Construção ────────────────────────────────────
                # O site mostra "X Ano de construção" (X pode ser um ano ou uma idade)
                match = re.search(r"(\d+)\s*[Aa]no[s]?\s*de\s*constru[çc][ãa]o", texto)
                if match:
                    valor = int(match.group(1))
                    if valor > 1900:       # é um ano (ex: 2019)
                        dados["data_construcao"] = valor
                    elif valor < 100:      # é uma idade em anos (ex: 5)
                        dados["data_construcao"] = 2025 - valor

                # ── Latitude e Longitude ──────────────────────────────────
                coords = re.findall(r"-\d{1,2}\.\d{4,}", html)
                coords_unicas = list(dict.fromkeys(coords))
                if len(coords_unicas) >= 2:
                    dados["latitude"]  = float(coords_unicas[0])
                    dados["longitude"] = float(coords_unicas[1])

            except Exception as e:
                print(f"    ⚠ Erro: {e}")

            # Salva progresso a cada 50 imóveis
            if i % 50 == 0:
                print(f"\n  💾 Salvando progresso ({i} imóveis)...\n")
                salvar(lista_dados)

            time.sleep(PAUSA_IMOVEL)

        browser.close()

    print(f"\n✓ Etapa 2 concluída.\n")
    return lista_dados


# ─── Etapa 3: Salvar ──────────────────────────────────────────────────────────

def salvar(lista_dados):
    df = pd.DataFrame(lista_dados)

    colunas = [
        "imobiliaria", "tipo", "rua", "bairro", "cidade",
        "tamanho_m2", "quartos", "banheiros", "vagas",
        "valor", "data_construcao",
        "latitude", "longitude",
        "amenidades", "url",
    ]
    colunas_existentes = [c for c in colunas if c in df.columns]
    df = df[colunas_existentes]

    df.to_excel(f"{ARQUIVO_SAIDA}.xlsx", index=False)
    df.to_csv(f"{ARQUIVO_SAIDA}.csv",   index=False, encoding="utf-8-sig")

    print(f"  Salvo em: {os.path.abspath(ARQUIVO_SAIDA + '.xlsx')}")
    return df


# ─── Principal ────────────────────────────────────────────────────────────────

def main():
    print("\n🏠 SCRAPER - CRÉDITO REAL IMÓVEIS (FLORIANÓPOLIS)")
    print("=" * 60)

    lista_dados = coletar_dados_listagem()
    if not lista_dados:
        print("Nenhum imóvel encontrado. Verifique a URL.")
        return

    print("💾 Salvando checkpoint etapa 1...")
    salvar(lista_dados)

    lista_dados = complementar_com_playwright(lista_dados)

    print("ETAPA 3: Salvando resultado final...")
    df = salvar(lista_dados)

    print(f"\n{'='*60}")
    print(f"✅ CONCLUÍDO! {len(df)} imóveis coletados.")
    print(f"{'='*60}\n")

    colunas_preview = [c for c in ["tipo", "bairro", "tamanho_m2", "quartos", "banheiros", "vagas", "valor"] if c in df.columns]
    print(df[colunas_preview].to_string(index=False))


if __name__ == "__main__":
    main()
