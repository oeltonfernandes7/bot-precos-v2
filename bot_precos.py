"""
bot_precos.py
Bot principal de pesquisa de preços em farmácias.
Lê produtos.xlsx → busca preços nos sites → salva resultado.csv
"""

import csv
import json
import logging
import os
import random
import re
import time
import urllib.parse
from datetime import datetime

import pandas as pd
import requests

import config

# ─────────────────────────────────────────────
# CONFIGURAÇÃO DE LOGS PROFISSIONAIS
# ─────────────────────────────────────────────

def configurar_logging():
    """Configura logging profissional com rotação automática de arquivos."""
    # Criar diretório de logs se não existir
    os.makedirs("logs", exist_ok=True)
    
    # Nome do arquivo de log com timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    arquivo_log = f"logs/log_{timestamp}.txt"
    
    # Configurar logging
    logging.basicConfig(
        filename=arquivo_log,
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        encoding='utf-8'
    )
    
    # Também logar no console
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
    console.setFormatter(formatter)
    logging.getLogger('').addHandler(console)
    
    # Rotação: manter apenas os últimos 30 arquivos de log
    arquivos_log = sorted([f for f in os.listdir("logs") if f.startswith("log_") and f.endswith(".txt")])
    if len(arquivos_log) > 30:
        for arquivo_antigo in arquivos_log[:-30]:
            os.remove(os.path.join("logs", arquivo_antigo))
            logging.info(f"Log antigo removido: {arquivo_antigo}")
    
    logging.info("Sistema de logs configurado. Arquivo: " + arquivo_log)
    return arquivo_log

# ─────────────────────────────────────────────
# GESTÃO DE FONTES DINÂMICAS
# ─────────────────────────────────────────────

def carregar_fontes():
    """Carrega fontes ativas do fontes.json."""
    try:
        with open("fontes.json", "r", encoding="utf-8") as f:
            dados = json.load(f)
        fontes_ativas = [fonte for fonte in dados.get("fontes", []) if fonte.get("status") == "ativa"]
        logging.info(f"Fontes carregadas: {len(fontes_ativas)} ativas de {len(dados.get('fontes', []))} total")
        return fontes_ativas
    except FileNotFoundError:
        logging.error("Arquivo fontes.json não encontrado")
        return []
    except Exception as e:
        logging.error(f"Erro ao carregar fontes.json: {e}")
        return []

def salvar_fontes(fontes, pendentes=None):
    """Salva fontes atualizadas no fontes.json."""
    dados = {"fontes": fontes}
    if pendentes:
        dados["pendentes"] = pendentes
    try:
        with open("fontes.json", "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=2, ensure_ascii=False)
        logging.info("fontes.json atualizado")
    except Exception as e:
        logging.error(f"Erro ao salvar fontes.json: {e}")

# ─────────────────────────────────────────────
# USER-AGENTS ROTATIVOS — evita bloqueio por fingerprint
# ─────────────────────────────────────────────

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36 Edg/123.0.0.0",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 14_4_1) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4.1 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
]


def headers_rotativos(referer: str = "") -> dict:
    ua = random.choice(USER_AGENTS)
    h = {
        "User-Agent": ua,
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    }
    if referer:
        h["Referer"] = referer
    return h


# Playwright importado apenas se necessário
_playwright_disponivel = True
try:
    from playwright.sync_api import sync_playwright
except ImportError:
    _playwright_disponivel = False


# ─────────────────────────────────────────────
# UTILITÁRIOS
# ─────────────────────────────────────────────

def pausar():
    t = random.uniform(config.TEMPO_ESPERA_MIN, config.TEMPO_ESPERA_MAX)
    print(f"  ⏳ Aguardando {t:.1f}s...")
    time.sleep(t)


def limpar_preco(texto: str) -> str:
    if not texto:
        return ""
    # Mantém apenas dígitos, vírgula e ponto
    preco = re.sub(r"[^\d,.]", "", texto.strip())
    return preco


def carregar_estrategia() -> dict:
    if os.path.exists("estrategia.json"):
        with open("estrategia.json", "r", encoding="utf-8") as f:
            estrategia = json.load(f)
        logging.info(f"Estratégia carregada de 'estrategia.json' ({len(estrategia)} sites)")
        return estrategia
    logging.warning("'estrategia.json' não encontrado. Usando método padrão REQUESTS_HTML para todos os sites.")
    return {}


def carregar_produtos() -> list[dict]:
    try:
        df = pd.read_excel(config.ARQUIVO_ENTRADA, dtype=str)
        df.columns = [c.strip().lower() for c in df.columns]

        # Aceita colunas "ean", "codigo", "codigo_barras" ou "barcode"
        col_ean = next(
            (c for c in df.columns if c in ["ean", "codigo", "codigo_barras", "barcode"]),
            None,
        )
        col_nome = next(
            (c for c in df.columns if c in ["nome", "produto", "descricao", "description"]),
            None,
        )

        if col_ean is None:
            raise ValueError("Coluna EAN não encontrada. Esperado: 'ean', 'codigo', 'codigo_barras' ou 'barcode'.")

        produtos = []
        for _, row in df.iterrows():
            ean = str(row[col_ean]).strip().split(".")[0]  # remove decimais do Excel
            nome = str(row[col_nome]).strip() if col_nome else ""
            if ean:
                produtos.append({"ean": ean, "nome": nome})

        print(f"[ENTRADA] {len(produtos)} produtos carregados de '{config.ARQUIVO_ENTRADA}'")
        logging.info(f"{len(produtos)} produtos carregados de '{config.ARQUIVO_ENTRADA}'")
        return produtos

    except FileNotFoundError:
        print(f"[ERRO] Arquivo '{config.ARQUIVO_ENTRADA}' não encontrado.")
        logging.error(f"Arquivo '{config.ARQUIVO_ENTRADA}' não encontrado.")
        return []


# ─────────────────────────────────────────────
# BUSCADORES — Camada 1: Mercado Livre
# ─────────────────────────────────────────────

def buscar_mercadolivre_api(ean: str) -> list[dict]:
    url = f"https://api.mercadolivre.com/sites/MLB/search?q={ean}"
    resultados = []
    try:
        resp = requests.get(url, headers=headers_rotativos(), timeout=15)
        dados = resp.json()
        for item in dados.get("results", [])[:3]:
            preco = item.get("price", "")
            titulo = item.get("title", "")
            link = item.get("permalink", "")
            resultados.append({
                "preco": f"R$ {preco:.2f}".replace(".", ",") if preco else "",
                "titulo": titulo,
                "link": link,
            })
        if resultados:
            logging.info(f"EAN encontrado na API Mercado Livre: R$ {resultados[0]['preco']}")
    except Exception as e:
        logging.error(f"Erro na API Mercado Livre: {e}")
    return resultados


def buscar_mercadolivre_playwright(termo: str) -> list[dict]:
    """Abre lista.mercadolivre.com.br com Playwright e captura o primeiro preço."""
    if not _playwright_disponivel:
        return []

    termo_url = urllib.parse.quote(termo)
    url = f"https://lista.mercadolivre.com.br/{termo_url}"
    ua = random.choice(USER_AGENTS)
    resultados = []

    # Seletores em ordem de prioridade (do mais específico ao genérico)
    SELETORES_ML = [
        ".andes-money-amount__fraction",
        ".price-tag-fraction",
        "[class*='price'] span",
        "[class*='andes-money'] span",
    ]

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=not config.VISIVEL)
            context = browser.new_context(
                user_agent=ua,
                locale="pt-BR",
                viewport={"width": 1366, "height": 768},
            )
            page = context.new_page()
            page.goto(url, wait_until="networkidle", timeout=40000)
            page.wait_for_timeout(4000)

            preco_txt = ""
            titulo_txt = ""

            for seletor in SELETORES_ML:
                el = page.query_selector(seletor)
                if el:
                    txt = limpar_preco(el.inner_text())
                    if txt and len(txt) >= 2:
                        preco_txt = txt
                        break

            titulo_el = page.query_selector(
                ".ui-search-item__title, [class*='item__title'], h2"
            )
            if titulo_el:
                titulo_txt = titulo_el.inner_text().strip()

            browser.close()

            if preco_txt:
                resultados.append({
                    "preco": f"R$ {preco_txt}",
                    "titulo": titulo_txt,
                    "link": url,
                })
                logging.info(f"EAN encontrado no Playwright Mercado Livre: R$ {preco_txt}")
            else:
                logging.warning("Nenhum preço encontrado na página do Mercado Livre")

    except Exception as e:
        logging.error(f"Erro no Playwright Mercado Livre: {e}")

    return resultados


# ─────────────────────────────────────────────
# BUSCADORES — Camada 2: JSON VTEX
# ─────────────────────────────────────────────

def buscar_vtex_json(site_url: str, ean: str) -> list[dict]:
    url = (
        f"{site_url}/api/catalog_system/pub/products/search"
        f"?fq=alternateIds_Ean:{ean}"
    )
    resultados = []
    try:
        resp = requests.get(url, headers=config.HEADERS, timeout=15)
        produtos = resp.json()
        for p in produtos[:3]:
            nome_prod = p.get("productName", "")
            items = p.get("items", [])
            for item in items[:1]:
                sellers = item.get("sellers", [])
                for seller in sellers[:1]:
                    preco = seller.get("commertialOffer", {}).get("Price", "")
                    link = p.get("link", "")
                    resultados.append({
                        "preco": f"R$ {preco:.2f}".replace(".", ",") if preco else "",
                        "titulo": nome_prod,
                        "link": link,
                    })
    except Exception as e:
        logging.error(f"Erro no JSON VTEX: {e}")
    return resultados


# ─────────────────────────────────────────────
# BUSCADORES — Camada 2b: JSON específico Drogasil
# ─────────────────────────────────────────────

def buscar_drogasil_json(ean: str) -> list[dict]:
    """Tenta múltiplos endpoints antes do Playwright — JSON → safedata → buscapagina → HTML."""
    base = "https://www.drogasil.com.br"
    sessao = requests.Session()
    sessao.headers.update(headers_rotativos(referer=base))

    # ── Fase 1: endpoints JSON ──────────────────────────────────
    endpoints_json = [
        f"{base}/api/catalog_system/pub/products/search?fq=alternateIds_Ean:{ean}",
        f"{base}/api/catalog_system/pub/products/search?fq=ProductRefId:{ean}",
        f"{base}/api/io/search-api/pub/products_search?query={ean}&count=5",
        f"{base}/search?q={ean}&_format=json",
    ]

    for url in endpoints_json:
        try:
            resp = sessao.get(url, timeout=15)
            ct = resp.headers.get("content-type", "")
            if "json" not in ct and resp.text.strip().startswith("<"):
                continue
            dados = resp.json()
            print(f"    [DROGASIL JSON] OK: {url.split('com.br')[1].split('?')[0]}")
            logging.info(f"Drogasil JSON OK: {url.split('com.br')[1].split('?')[0]}")
            if isinstance(dados, list) and dados:
                return _extrair_vtex(dados, base)
            if isinstance(dados, dict) and dados.get("products"):
                return _extrair_vtex(dados["products"], base)
            try:
                prods = dados["data"]["productSearch"]["products"]
                if prods:
                    return _extrair_vtex(prods, base)
            except (KeyError, TypeError):
                pass
        except Exception:
            pass

    # ── Fase 2: SafeData API (VTEX CRM/search) ─────────────────
    try:
        url_safe = f"{base}/api/io/safedata/VI/search?q={ean}&_fields=Preco,NomeProduto"
        resp = sessao.get(url_safe, timeout=15)
        if resp.status_code == 200 and "json" in resp.headers.get("content-type", ""):
            dados = resp.json()
            print(f"    [DROGASIL SafeData] Status {resp.status_code}, itens: {len(dados)}")
            logging.info(f"Drogasil SafeData: Status {resp.status_code}, itens: {len(dados)}")
            if isinstance(dados, list) and dados:
                item = dados[0]
                preco = item.get("Preco", 0)
                nome_prod = item.get("NomeProduto", "")
                if preco:
                    return [{"preco": f"R$ {float(preco):.2f}".replace(".", ","),
                             "titulo": nome_prod, "link": f"{base}/search?q={ean}"}]
    except Exception as e:
        logging.warning(f"Drogasil SafeData erro: {e}")

    # ── Fase 3: buscapagina (retorna HTML com dados JSON embeddados) ──
    try:
        url_busca = f"{base}/buscapagina?fq=alternateIds_Ean:{ean}&_from=0&_to=3"
        resp = sessao.get(url_busca, timeout=15)
        if resp.status_code == 200:
            html = resp.text
            # buscapagina pode retornar JSON dentro de script ou HTML com preços
            precos = re.findall(r'"Price"\s*:\s*([\d.]+)', html)
            nomes = re.findall(r'"productName"\s*:\s*"([^"]+)"', html)
            if precos:
                preco_val = float(precos[0])
                nome_val = nomes[0] if nomes else ""
                print(f"    [DROGASIL buscapagina] Preco encontrado: {preco_val}")
                logging.info(f"Drogasil buscapagina: Preço encontrado R$ {preco_val:.2f}")
                return [{"preco": f"R$ {preco_val:.2f}".replace(".", ","),
                         "titulo": nome_val, "link": f"{base}/search?q={ean}"}]
    except Exception as e:
        logging.warning(f"Drogasil buscapagina erro: {e}")

    # ── Fase 4: requests com headers completos de navegador ────
    try:
        url_html = f"{base}/search?q={ean}"
        hdrs = {
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "pt-BR,pt;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Referer": base,
        }
        resp = requests.get(url_html, headers=hdrs, timeout=15)
        print(f"    [DROGASIL HTML] Status {resp.status_code}, tamanho {len(resp.text)} bytes")
        logging.info(f"Drogasil HTML: Status {resp.status_code}, tamanho {len(resp.text)} bytes")
        if len(resp.text) > 1000:
            precos = re.findall(r'"Price"\s*:\s*([\d.]+)', resp.text)
            nomes = re.findall(r'"productName"\s*:\s*"([^"]+)"', resp.text)
            if not precos:
                precos = re.findall(r'R\$\s*[\d.,]+', resp.text)
                if precos:
                    return [{"preco": precos[0].strip(), "titulo": "", "link": url_html}]
            if precos:
                preco_val = float(precos[0])
                nome_val = nomes[0] if nomes else ""
                return [{"preco": f"R$ {preco_val:.2f}".replace(".", ","),
                         "titulo": nome_val, "link": url_html}]
    except Exception as e:
        logging.warning(f"Drogasil HTML erro: {e}")

    print(f"    [DROGASIL] Todas as fases falharam, usando Playwright")
    logging.warning("Drogasil: Todas as fases falharam, usando Playwright")
    return []


# ─────────────────────────────────────────────
# BUSCADORES — Camada 2c: Pacheco (DPSP / VTEX)
# ─────────────────────────────────────────────

def buscar_pacheco(ean: str, nome_produto: str = "") -> list[dict]:
    """Pacheco (grupo DPSP, VTEX): JSON → buscapagina → requests HTML → Playwright."""
    base = "https://www.pacheco.com.br"
    sessao = requests.Session()
    sessao.headers.update(headers_rotativos(referer=base))

    # ── Camada 1: endpoints VTEX JSON ──────────────────────────
    endpoints_vtex = [
        f"{base}/api/catalog_system/pub/products/search?fq=alternateIds_Ean:{ean}",
        f"{base}/api/catalog_system/pub/products/search?fq=ProductRefId:{ean}",
        f"{base}/api/io/search-api/pub/products_search?query={ean}&count=5",
    ]

    for url in endpoints_vtex:
        try:
            resp = sessao.get(url, timeout=15)
            ct = resp.headers.get("content-type", "")
            if "json" not in ct and resp.text.strip().startswith("<"):
                continue
            dados = resp.json()
            print(f"    [PACHECO JSON] OK: {url.split('com.br')[1].split('?')[0]}")
            logging.info(f"Pacheco JSON OK: {url.split('com.br')[1].split('?')[0]}")
            if isinstance(dados, list) and dados:
                return _extrair_vtex(dados, base)
            if isinstance(dados, dict) and dados.get("products"):
                return _extrair_vtex(dados["products"], base)
            try:
                prods = dados["data"]["productSearch"]["products"]
                if prods:
                    return _extrair_vtex(prods, base)
            except (KeyError, TypeError):
                pass
        except Exception:
            pass

    # ── Camada 1b: buscapagina (HTML com JSON embeddado) ───────
    try:
        url_bp = f"{base}/buscapagina?fq=alternateIds_Ean:{ean}&_from=0&_to=3"
        resp = sessao.get(url_bp, timeout=15)
        html_bp = resp.text
        precos_bp = re.findall(r'"Price"\s*:\s*([\d.]+)', html_bp)
        nomes_bp = re.findall(r'"productName"\s*:\s*"([^"]+)"', html_bp)
        if precos_bp:
            preco_val = float(precos_bp[0])
            print(f"    [PACHECO buscapagina] Preco: {preco_val}")
            logging.info(f"Pacheco buscapagina: Preço R$ {preco_val:.2f}")
            return [{"preco": f"R$ {preco_val:.2f}".replace(".", ","),
                     "titulo": nomes_bp[0] if nomes_bp else "",
                     "link": f"{base}/search?q={ean}"}]
    except Exception as e:
        logging.warning(f"Pacheco buscapagina erro: {e}")

    # ── Camada 2: requests com headers completos de navegador ──
    try:
        url_html = f"{base}/search?q={ean}"
        hdrs = {
            "User-Agent": random.choice(USER_AGENTS),
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "pt-BR,pt;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Connection": "keep-alive",
            "Referer": base,
        }
        resp = requests.get(url_html, headers=hdrs, timeout=15)
        print(f"    [PACHECO HTML] Status {resp.status_code}, tamanho {len(resp.text)} bytes")
        logging.info(f"Pacheco HTML: Status {resp.status_code}, tamanho {len(resp.text)} bytes")
        if len(resp.text) > 1000:
            precos = re.findall(r'"Price"\s*:\s*([\d.]+)', resp.text)
            nomes = re.findall(r'"productName"\s*:\s*"([^"]+)"', resp.text)
            if not precos:
                precos_raw = re.findall(r'R\$\s*[\d.,]+', resp.text)
                if precos_raw:
                    return [{"preco": precos_raw[0].strip(), "titulo": "", "link": url_html}]
            if precos:
                preco_val = float(precos[0])
                return [{"preco": f"R$ {preco_val:.2f}".replace(".", ","),
                         "titulo": nomes[0] if nomes else "", "link": url_html}]
    except Exception as e:
        logging.warning(f"Pacheco HTML erro: {e}")

    # ── Camada 3: Playwright ────────────────────────────────────
    print(f"    [PACHECO] Usando Playwright")
    logging.warning("Pacheco: Usando Playwright")
    return buscar_playwright(base, ean, "Pacheco", nome_produto)


def _extrair_vtex(produtos: list, base_url: str = "") -> list[dict]:
    resultados = []
    for p in produtos[:3]:
        nome_prod = p.get("productName", p.get("productTitle", ""))
        link = p.get("link", p.get("linkText", ""))
        if link and not link.startswith("http"):
            link = f"{base_url}/{link.lstrip('/')}"
            if not link.endswith("/p"):
                link += "/p"
        items = p.get("items", p.get("skus", []))
        for item in items[:1]:
            sellers = item.get("sellers", [])
            for seller in sellers[:1]:
                preco = seller.get("commertialOffer", {}).get("Price", 0)
                if preco:
                    resultados.append({
                        "preco": f"R$ {preco:.2f}".replace(".", ","),
                        "titulo": nome_prod,
                        "link": link,
                    })
    return resultados


# ─────────────────────────────────────────────
# BUSCADORES — Camada 3: Playwright (JS pesado)
# ─────────────────────────────────────────────

SELETORES_PRECO = (
    "[class*='price']:not([class*='original']):not([class*='from']), "
    "[class*='preco'], [class*='valor'], "
    "[data-testid*='price'], [itemprop='price'], "
    ".shelf-item__buy-btn, .price"
)

SELETORES_TITULO = (
    "[class*='product-name'], [class*='productName'], "
    "[class*='shelf-item__title'], [class*='title'], h2, h3"
)


def buscar_playwright(site_url: str, ean: str, nome_site: str, nome_produto: str = "") -> list[dict]:
    if not _playwright_disponivel:
        print(f"    [AVISO] Playwright nao instalado: python -m pip install playwright && python -m playwright install chromium")
        logging.warning("Playwright não instalado")
        return []

    resultados = []
    ua = random.choice(USER_AGENTS)

    def _tentar_busca(page, termo: str) -> list[dict]:
        url_busca = f"{site_url}/search?q={termo}"
        page.goto(url_busca, wait_until="networkidle", timeout=40000)
        page.wait_for_timeout(5000)  # aguarda JS carregar conteúdo dinâmico

        elementos_preco = page.query_selector_all(SELETORES_PRECO)
        titulos = page.query_selector_all(SELETORES_TITULO)

        achados = []
        for i, el in enumerate(elementos_preco[:3]):
            preco_txt = limpar_preco(el.inner_text())
            titulo_txt = titulos[i].inner_text().strip() if i < len(titulos) else ""
            if preco_txt and len(preco_txt) >= 3:
                achados.append({
                    "preco": f"R$ {preco_txt}",
                    "titulo": titulo_txt,
                    "link": url_busca,
                })
        return achados

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=not config.VISIVEL)
            context = browser.new_context(
                user_agent=ua,
                locale="pt-BR",
                viewport={"width": 1366, "height": 768},
                extra_http_headers={
                    "Accept-Language": "pt-BR,pt;q=0.9",
                    "Referer": site_url,
                },
            )
            page = context.new_page()

            # Tentativa 1: busca pelo EAN
            resultados = _tentar_busca(page, ean)

            # Tentativa 2: busca pelo nome do produto (fallback)
            if not resultados and nome_produto:
                print(f"    [PLAYWRIGHT] EAN sem resultado, tentando por nome: {nome_produto[:40]}")
                logging.info(f"Playwright {nome_site}: Tentando por nome: {nome_produto[:40]}")
                resultados = _tentar_busca(page, nome_produto)

            browser.close()

    except Exception as e:
        logging.error(f"Erro no Playwright {nome_site}: {e}")

    return resultados


# ─────────────────────────────────────────────
# BUSCADORES — Camada 4: requests + HTML simples
# ─────────────────────────────────────────────

def buscar_requests_html(site_url: str, ean: str) -> list[dict]:
    url_busca = f"{site_url}/search?q={ean}"
    resultados = []
    try:
        resp = requests.get(url_busca, headers=config.HEADERS, timeout=15)
        html = resp.text

        # Regex genérico para preços em BRL
        precos = re.findall(r"R\$\s*[\d.,]+", html)
        precos_unicos = list(dict.fromkeys(precos))[:3]

        for preco in precos_unicos:
            resultados.append({
                "preco": preco.strip(),
                "titulo": "",
                "link": url_busca,
            })

    except Exception as e:
        logging.error(f"Erro no requests HTML: {e}")
    return resultados


# ─────────────────────────────────────────────
# DISPATCHER — escolhe método por site
# ─────────────────────────────────────────────

def buscar_em_site(site: dict, ean: str, estrategia: dict, nome_produto: str = "") -> dict:
    nome = site["nome"]
    url = site["url"]
    dominio = url.replace("https://www.", "").replace("https://", "")

    info = estrategia.get(dominio, {})
    metodo = info.get("metodo_recomendado", "REQUESTS_HTML")

    print(f"  [{nome}] Método: {metodo}")
    logging.info(f"Iniciando pesquisa do produto: {nome_produto or ean} em {nome} usando {metodo}")
    resultados = []

    if metodo == "API_MERCADOLIVRE" or "mercadolivre" in dominio:
        # Playwright primeiro (mais confiável para ML); API como fallback
        termo_busca = ean if ean else nome_produto
        resultados = buscar_mercadolivre_playwright(termo_busca)
        if not resultados:
            print(f"    [ML] Playwright sem resultado, tentando API")
            logging.warning("Mercado Livre: Playwright sem resultado, tentando API")
            resultados = buscar_mercadolivre_api(ean)
        if not resultados and nome_produto and nome_produto != ean:
            print(f"    [ML] API sem resultado, tentando Playwright por nome")
            logging.warning("Mercado Livre: API sem resultado, tentando Playwright por nome")
            resultados = buscar_mercadolivre_playwright(nome_produto)

    elif "drogasil" in dominio:
        resultados = buscar_drogasil_json(ean)
        if not resultados:
            resultados = buscar_playwright(url, ean, nome, nome_produto)

    elif "pacheco" in dominio:
        resultados = buscar_pacheco(ean, nome_produto)

    elif metodo == "JSON_VTEX":
        resultados = buscar_vtex_json(url, ean)

    elif metodo == "PLAYWRIGHT":
        resultados = buscar_playwright(url, ean, nome, nome_produto)

    else:  # REQUESTS_HTML ou fallback
        resultados = buscar_requests_html(url, ean)

    if resultados:
        melhor = resultados[0]
        print(f"    → {melhor['preco']} | {melhor['titulo'][:60]}")
        logging.info(f"EAN encontrado em {nome}: {melhor['preco']}")
        return {
            "site": nome,
            "preco": melhor["preco"],
            "titulo": melhor["titulo"],
            "link": melhor["link"],
        }

    print(f"    → Nao encontrado")
    logging.warning(f"EAN não encontrado em {nome}")
    return {"site": nome, "preco": "Nao encontrado", "titulo": "", "link": ""}


# ─────────────────────────────────────────────
# SAÍDA — salvar linha no CSV
# ─────────────────────────────────────────────

def inicializar_csv(fontes_ativas):
    cabecalho = ["EAN", "Nome", "Data"] + [fonte["coluna_csv"] for fonte in fontes_ativas] + [f"Link_{fonte['coluna_csv']}" for fonte in fontes_ativas]
    if not os.path.exists(config.ARQUIVO_SAIDA):
        with open(config.ARQUIVO_SAIDA, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(cabecalho)
        logging.info(f"CSV inicializado: {config.ARQUIVO_SAIDA}")


def salvar_linha(ean: str, nome: str, resultados_por_site: dict, fontes_ativas):
    data = datetime.now().strftime("%d/%m/%Y %H:%M")
    precos = [resultados_por_site.get(fonte["nome"], {}).get("preco", "") for fonte in fontes_ativas]
    links = [resultados_por_site.get(fonte["nome"], {}).get("link", "") for fonte in fontes_ativas]

    linha = [ean, nome, data] + precos + links

    with open(config.ARQUIVO_SAIDA, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f, delimiter=";")
        writer.writerow(linha)
    logging.info(f"Linha salva no CSV: EAN {ean}")


# ─────────────────────────────────────────────
# FUNÇÕES PARA INTERFACE STREAMLIT
# ─────────────────────────────────────────────

def buscar_produto(ean: str, nome: str, fontes_ativas: list, estrategia: dict) -> dict:
    """Busca preço de um produto em todas as fontes ativas."""
    resultados = {}
    for fonte in fontes_ativas:
        resultado = buscar_em_site(fonte, ean, estrategia, nome_produto=nome)
        resultados[fonte["nome"]] = resultado
        pausar()
    return resultados

def main():
    # Configurar logging
    arquivo_log = configurar_logging()
    
    print("=" * 55)
    print("  BOT DE PESQUISA DE PREÇOS EM FARMÁCIAS v2.0")
    print("=" * 55)
    logging.info("Bot de preços v2.0 iniciado")

    # Carregar fontes dinâmicas
    fontes_ativas = carregar_fontes()
    if not fontes_ativas:
        logging.error("Nenhuma fonte ativa encontrada. Verifique fontes.json")
        return

    estrategia = carregar_estrategia()

    produtos = carregar_produtos()
    if not produtos:
        return

    inicializar_csv(fontes_ativas)

    inicio = config.INICIAR_DO_PRODUTO - 1
    total = len(produtos)

    print(f"\n[INÍCIO] Processando produtos {inicio + 1} a {total}\n")
    logging.info(f"Iniciando processamento de {total - inicio} produtos")

    # Estatísticas para resumo final
    total_pesquisados = 0
    encontrados_por_site = {fonte["nome"]: 0 for fonte in fontes_ativas}
    nao_encontrados_por_site = {fonte["nome"]: 0 for fonte in fontes_ativas}
    tempo_inicio = time.time()

    for idx, produto in enumerate(produtos[inicio:], start=inicio + 1):
        ean = produto["ean"]
        nome = produto["nome"]

        print(f"\n[{idx}/{total}] EAN: {ean} | {nome}")
        print("-" * 45)
        logging.info(f"Processando produto {idx}/{total}: EAN {ean} - {nome}")

        resultados_por_site = {}

        # Ordem aleatória entre produtos diferentes — reduz padrão detectável
        fontes_ordem = random.sample(fontes_ativas, len(fontes_ativas))
        for fonte in fontes_ordem:
            resultado = buscar_em_site(fonte, ean, estrategia, nome_produto=nome)
            resultados_por_site[fonte["nome"]] = resultado
            if resultado["preco"] != "Nao encontrado":
                encontrados_por_site[fonte["nome"]] += 1
            else:
                nao_encontrados_por_site[fonte["nome"]] += 1
            pausar()

        salvar_linha(ean, nome, resultados_por_site, fontes_ativas)
        print(f"  ✓ Salvo em '{config.ARQUIVO_SAIDA}'")
        total_pesquisados += 1

        # Pausa extra entre lotes
        if idx % config.LOTE == 0 and idx < total:
            pausa_lote = random.uniform(15, 25)
            print(f"\n[LOTE] Pausa de {pausa_lote:.0f}s entre lotes...\n")
            logging.info(f"Pausa entre lotes: {pausa_lote:.0f}s")
            time.sleep(pausa_lote)

    # Resumo final
    tempo_total = time.time() - tempo_inicio
    logging.info("=" * 50)
    logging.info("RESUMO DA EXECUÇÃO")
    logging.info("=" * 50)
    logging.info(f"Total pesquisados: {total_pesquisados}")
    for fonte in fontes_ativas:
        nome_fonte = fonte["nome"]
        logging.info(f"{nome_fonte}: {encontrados_por_site[nome_fonte]} encontrados, {nao_encontrados_por_site[nome_fonte]} não encontrados")
    logging.info(f"Tempo total: {tempo_total:.2f}s")
    logging.info(f"Próximo INICIAR_DO_PRODUTO sugerido: {total + 1}")
    logging.info("=" * 50)

    print("\n" + "=" * 55)
    print(f"[CONCLUÍDO] {total - inicio} produtos processados.")
    print(f"[SAÍDA] Resultados em '{config.ARQUIVO_SAIDA}'")
    print(f"[LOGS] Arquivo de log: {arquivo_log}")
    print("=" * 55)


if __name__ == "__main__":
    main()
