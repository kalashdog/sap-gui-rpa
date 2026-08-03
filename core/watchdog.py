import psutil
import subprocess
import os
import logging
import threading

def _iniciar(nome, caminho, args=None):
    if not caminho or not os.path.exists(caminho):
        logging.error(f"[Watchdog] Executável '{nome}' não encontrado: {caminho}")
        return
    try:
        subprocess.Popen([caminho] + (args or []))
        logging.info(f"[Watchdog] '{nome}' iniciado com sucesso.")
    except Exception as e:
        logging.error(f"[Watchdog] Falha ao abrir '{nome}': {e}")

def _resolver_caminho(candidatos):
    return next((c for c in candidatos if c and os.path.exists(c)), None)

def watchdog_infraestrutura():
    logging.info("[Watchdog] Verificando infraestrutura...")

    try:
        ativos = {p.name().lower() for p in psutil.process_iter()}
    except Exception as e:
        logging.error(f"[Watchdog] Erro ao listar processos: {e}")
        return

    local   = os.environ.get("LOCALAPPDATA", "")

    # Onedrive
    if "onedrive.exe" not in ativos:
        logging.warning("[Watchdog] OneDrive não encontrado. Reabrindo...")
        caminho = _resolver_caminho([
            os.path.join(local, "Microsoft", "OneDrive", "OneDrive.exe"),
            r"C:\Program Files\Microsoft OneDrive\OneDrive.exe",
        ])
        _iniciar("OneDrive", caminho, ["/background"])

    # SAP Logon fk
    ferramenta_remota = {"anydesk.exe", "sap logon.exe"}
    if not (ativos & ferramenta_remota):
        logging.warning("[Watchdog] Ferramenta não encontrada. Tentando reabrir...")
        od = os.environ.get("OneDriveCommercial") or os.environ.get("OneDrive")
        caminho = _resolver_caminho([
            os.path.join(od, "SESÉ DASHBOARD", "Anchieta Dados", "000 - Dashboard Dados", ".shared", "SAP Logon.exe") if od else None,
            os.path.join(od, "SESÉ DASHBOARD", "002 - Filiais database", "000 - Global", ".Assets", "VW", "SAP Logon.exe") if od else None,
        ])
        _iniciar("AnyDesk", caminho)


DEFAULT_JOB_TIMEOUT = 600  # 10 minuto

class JobWatchdog:
    """
    Watchdog por job: encerra processos SAP se uma operação ultrapassar o timeout acima.
    Se o timeout expirar, mata saplogon.exe (o proximo request vai abrir o sap dnv).
    """
    def __init__(self, timeout: int, job_name: str = ""):
        self.timeout = timeout
        self.job_name = job_name
        self._timer = None
        self._triggered = False

    def _on_timeout(self):
        self._triggered = True
        logging.error(f"[WATCHDOG] Timeout de {self.timeout}s atingido para '{self.job_name}'. Encerrando SAP...")
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", "saplogon.exe"],
                capture_output=True, timeout=10
            )
            logging.warning("[WATCHDOG] Processos SAP encerrados com sucesso.")
        except Exception as e:
            logging.error(f"[WATCHDOG] Falha ao encerrar SAP: {e}")

    def __enter__(self):
        self._timer = threading.Timer(self.timeout, self._on_timeout)
        self._timer.daemon = True
        self._timer.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self._timer:
            self._timer.cancel()
        return False

    @property
    def was_triggered(self):
        return self._triggered
