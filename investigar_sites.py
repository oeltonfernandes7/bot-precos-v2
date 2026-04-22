"""
investigar_sites.py
Roda UMA ÚNICA VEZ antes do bot principal.
Detecta automaticamente qual tecnologia cada site usa e gera estrategia.json.
"""

import json
import requests
from config import HEADERS

# ─────────────────────────────────────────────
# SITES A INVESTIGAR
# ─────────────────────────────────────────────

SITES = [
    "https://www.drogasil.com.br",
    "https://www.farmaciaindiana.com.br",
    "https://www.mercadolivre.com.br",
    "https://www.pacheco.com.br",
]

# EAN genérico para testes
EAN_TESTE = "7891058013202"

# Timeout padrão das requisições (segundos)
TIMEOUT = 15


# ─────────────────────────────────────────────
# VERIFICAÇÃO 1 — Detectar plataforma (VTEX, etc.)
# ─────────────────────────────────────────────

def detectar_plataforma(site: str) -> str:
    print(f"\n[INVESTIGANDO] {site}")
    try:
        resp = requests.get(site, headers=HEADERS, timeout=TIMEOUT)
        headers_lower = {k.lower(): v for k, v in resp.headers.items()}
        html = resp.text.lower()

        # Checagem por headers HTTP
        for header in headers_lower:
            if "vtex" in header:
                print(f"  [PLATAFORMA DETECTADA] {site} → VTEX (via header: {header})")
                return "VTEX"

        # Checagem por conteúdo HTML
        if "__vtex" in html or "vtex.com" in html or "vtexcommercestable" in html:
            print(f"  [PLATAFORMA DETECTADA] {site} → VTEX (via HTML)")
            return "VTEX"

        print(f"  [PLATAFORMA] {site} → Desconhecida / Genérica")
        return "GENÉRICA"

    except Exception as e:
        print(f"  [ERRO] Falha ao acessar {site}: {e}")
        return "ERRO"


# ─────────────────────────────────────────────
# VERIFICAÇÃO 2 — Testar API oficial do Mercado Livre
# ─────────────────────────────────────────────

def testar_api_mercadolivre() -> bool:
    url = f"https://api.mercadolivre.com/sites/MLB/search?q={EAN_TESTE}"
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        dados = resp.json()
        if "results" in dados:
            print(f"  [API OK] Mercado Livre → Camada 1 disponível (API oficial)")
            return True
        print(f"  [API FALHOU] Resposta sem 'results': {list(dados.keys())}")
        return False
    except Exception as e:
        print(f"  [ERRO] API Mercado Livre: {e}")
        return False


# ─────────────────────────────────────────────
# VERIFICAÇÃO 3 — Testar endpoint JSON interno (VTEX)
# ─────────────────────────────────────────────

def testar_json_vtex(site: str) -> bool:
    url = (
        f"{site}/api/catalog_system/pub/products/search"
        f"?fq=alternateIds_Ean:{EAN_TESTE}"
    )
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        dados = resp.json()
        if isinstance(dados, list) and len(dados) > 0:
            print(f"  [JSON OK] {site} → Camada 2 disponível (endpoint VTEX)")
            return True
        print(f"  [JSON VAZIO] {site} → Endpoint VTEX retornou lista vazia")
        return False
    except Exception as e:
        print(f"  [JSON FALHOU] {site}: {e}")
        return False


# ─────────────────────────────────────────────
# VERIFICAÇÃO 4 — Detectar necessidade de Playwright
# ─────────────────────────────────────────────

def precisa_playwright(site: str) -> bool:
    url_busca = f"{site}/search?q={EAN_TESTE}"
    try:
        resp = requests.get(url_busca, headers=HEADERS, timeout=TIMEOUT)
        html = resp.text.lower()

        sinais_js = [
            "noscript",
            "enable javascript",
            "javascript is required",
            "__next_data__",
            "window.__remixcontext",
            "react-root",
        ]
        for sinal in sinais_js:
            if sinal in html:
                print(f"  [JAVASCRIPT PESADO] {site} → Playwright necessário (sinal: {sinal})")
                return True

        # Heurística: página muito curta provavelmente é shell JS
        if len(resp.text) < 3000:
            print(f"  [JAVASCRIPT PESADO] {site} → Playwright necessário (HTML muito curto: {len(resp.text)} bytes)")
            return True

        print(f"  [HTML OK] {site} → requests suficiente, Playwright não necessário")
        return False

    except Exception as e:
        print(f"  [ERRO] Verificação Playwright em {site}: {e}")
        return True  # em caso de dúvida, usa Playwright


# ─────────────────────────────────────────────
# EXECUÇÃO PRINCIPAL
# ─────────────────────────────────────────────

def investigar():
    estrategia = {}

    for site in SITES:
        nome = site.replace("https://www.", "").replace("https://", "")
        info = {
            "url": site,
            "plataforma": None,
            "api_oficial": False,
            "json_vtex": False,
            "playwright": False,
            "metodo_recomendado": None,
        }

        plataforma = detectar_plataforma(site)
        info["plataforma"] = plataforma

        # Mercado Livre → testa API própria
        if "mercadolivre" in site:
            info["api_oficial"] = testar_api_mercadolivre()
            info["metodo_recomendado"] = "API_MERCADOLIVRE" if info["api_oficial"] else "PLAYWRIGHT"

        elif plataforma == "VTEX":
            info["json_vtex"] = testar_json_vtex(site)
            if info["json_vtex"]:
                info["metodo_recomendado"] = "JSON_VTEX"
            else:
                info["playwright"] = precisa_playwright(site)
                info["metodo_recomendado"] = "PLAYWRIGHT" if info["playwright"] else "REQUESTS_HTML"

        else:
            info["playwright"] = precisa_playwright(site)
            info["metodo_recomendado"] = "PLAYWRIGHT" if info["playwright"] else "REQUESTS_HTML"

        estrategia[nome] = info
        print(f"  → Método recomendado: {info['metodo_recomendado']}\n")

    # Salvar resultado
    with open("estrategia.json", "w", encoding="utf-8") as f:
        json.dump(estrategia, f, ensure_ascii=False, indent=2)

    print("=" * 50)
    print("[CONCLUÍDO] Arquivo 'estrategia.json' gerado com sucesso.")
    print("=" * 50)
    return estrategia


if __name__ == "__main__":
    investigar()
