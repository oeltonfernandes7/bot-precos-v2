# launcher.py — Ponto de entrada do executavel BotPrecos
# Inicia o servidor Streamlit e abre o navegador automaticamente.
# Ao fechar o navegador, encerra o processo.

import sys
import os
import threading
import webbrowser
import socket
import time

PORT = 8501
URL = f"http://localhost:{PORT}"
GRACE_PERIOD = 30  # segundos sem conexao antes de encerrar


def caminho_base():
    """Retorna o diretorio raiz, funcionando tanto normal quanto compilado."""
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def aguardar_porta(porta=PORT, timeout=60):
    """Fica tentando conectar na porta ate o servidor responder."""
    inicio = time.time()
    while time.time() - inicio < timeout:
        try:
            with socket.create_connection(("localhost", porta), timeout=1):
                return True
        except OSError:
            time.sleep(0.5)
    return False


def ha_browser_conectado(porta=PORT):
    """
    Verifica se algum navegador tem conexao TCP estabelecida com o Streamlit.
    Retorna True se nao conseguir verificar (evita encerramento indevido).
    """
    try:
        import psutil

        for conn in psutil.net_connections(kind="tcp"):
            if (
                conn.laddr
                and conn.laddr.port == porta
                and conn.status == "ESTABLISHED"
            ):
                return True
        return False
    except Exception:
        # psutil indisponivel ou sem permissao — nao encerra automaticamente
        return True


def abrir_browser_e_monitorar():
    """
    Thread auxiliar:
      1. Aguarda o servidor ficar pronto.
      2. Abre o navegador.
      3. Monitora conexoes e encerra o processo quando o browser for fechado.
    """
    if not aguardar_porta():
        # Servidor nao respondeu no tempo limite — encerra tudo
        os._exit(1)

    # Abre o navegador apos o servidor estar pronto
    webbrowser.open(URL)

    # Aguarda a primeira conexao do navegador (ate 60 s)
    for _ in range(20):
        if ha_browser_conectado():
            break
        time.sleep(3)

    # Monitora em loop: se nao houver conexao por GRACE_PERIOD segundos, encerra
    sem_conexao_desde = None

    while True:
        time.sleep(3)
        if ha_browser_conectado():
            sem_conexao_desde = None
        else:
            if sem_conexao_desde is None:
                sem_conexao_desde = time.time()
            elif time.time() - sem_conexao_desde > GRACE_PERIOD:
                # Navegador fechado — encerra o servidor e o processo
                os._exit(0)


if __name__ == "__main__":
    base = caminho_base()
    app_path = os.path.join(base, "app.py")

    # Inicia thread que abre o browser e monitora conexoes
    t = threading.Thread(target=abrir_browser_e_monitorar, daemon=True)
    t.start()

    # Inicia o servidor Streamlit no processo principal
    sys.argv = [
        "streamlit",
        "run",
        app_path,
        f"--server.port={PORT}",
        "--server.headless=true",
        "--browser.gatherUsageStats=false",
        "--server.enableCORS=false",
    ]

    from streamlit.web import cli as stcli

    stcli.main()
