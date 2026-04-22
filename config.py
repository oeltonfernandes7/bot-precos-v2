# ─────────────────────────────────────────────
# CONFIGURAÇÕES DO BOT DE PESQUISA DE PREÇOS
# ─────────────────────────────────────────────

# Quantidade de produtos processados por lote antes de pausar
LOTE = 6

# Arquivo Excel de entrada com a lista de produtos (coluna EAN obrigatória)
ARQUIVO_ENTRADA = "produtos.xlsx"

# Arquivo CSV onde os resultados serão salvos
ARQUIVO_SAIDA = "resultado.csv"

# Tempo mínimo de espera (segundos) entre buscas — evita bloqueio por bot
TEMPO_ESPERA_MIN = 3

# Tempo máximo de espera (segundos) entre buscas — variação aleatória
TEMPO_ESPERA_MAX = 7

# Número do produto a partir do qual iniciar (1 = do início, útil para retomar)
INICIAR_DO_PRODUTO = 1

# False = navegador oculto (mais rápido) | True = navegador visível (para depuração)
VISIVEL = False

# ─────────────────────────────────────────────
# CABEÇALHOS HTTP — simula navegador real
# ─────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# Fontes dinâmicas - não editar manualmente, usar fontes.json
SITES = []  # Mantido para compatibilidade, mas usa fontes.json
