"""
Orquestrador inteligente - loop continuo com estado por job.
request -> espera -> extract -> repeat.
"""
import json
import logging
import os
import sys
import threading
import requests
import pythoncom
from datetime import datetime, date
import subprocess

from core.utils import get_onedrive_path
from core.connection import SAPConnection
from config.settings import settings
import transactions.request as req_module
from transactions.extract import extract_sp02_job
from core.watchdog import watchdog_infraestrutura, JobWatchdog, DEFAULT_JOB_TIMEOUT

app_data = os.environ.get('LOCALAPPDATA')
if not app_data:
    app_data = os.path.expanduser('~')

BASE_DIR = os.path.join(app_data, "HubSeseRPA")
STATE_DIR = os.path.join(BASE_DIR, "state")
LOGS_DIR = os.path.join(BASE_DIR, "logs")

os.makedirs(STATE_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)

# Configuração de Logs
from logging.handlers import RotatingFileHandler
log_file = os.path.join(LOGS_DIR, "rpa_execution.log")
handlers = [
    RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=3, encoding='utf-8'),
    logging.StreamHandler()
]
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', handlers=handlers)

MASTER_PLANT = "01-Anchieta"
CYCLE_WAIT = 300
FIREBASE_URL = os.getenv("FIREBASE_URL", "")
TIMEOUT = 600

FIREBASE_SECRET = None
try:
    from core.bootstrap import FIREBASE_RTDB_SECRET, FIREBASE_URL as B_URL  # type: ignore
    FIREBASE_SECRET = FIREBASE_RTDB_SECRET
    if B_URL:
        FIREBASE_URL = B_URL
except ImportError:
    pass

if not FIREBASE_SECRET:
    FIREBASE_SECRET = os.getenv("FIREBASE_RTDB_SECRET")

if not FIREBASE_SECRET or not FIREBASE_URL:
    logging.warning("Módulo de segurança Bootstrap ausente ou incompleto. O RPA não conseguirá enviar telemetria.")
def _get_job_handler(job_key: str, job_data: dict):
    """
    Descobre dinamicamente a função de request no módulo transactions.request.
    """

    func_name = f"request_{job_key.lower()}"
    if hasattr(req_module, func_name):
        return getattr(req_module, func_name)
    
    transaction = job_data.get("transaction", "").lower()
    transaction_clean = transaction.replace("/", "").replace("\\", "")
    func_name_tx = f"request_{transaction_clean}"
    if hasattr(req_module, func_name_tx):
        return getattr(req_module, func_name_tx)
        
    return None


class JobState:
    def __init__(self, plant_id):
        self._plant_id = plant_id
        os.makedirs(STATE_DIR, exist_ok=True)
        self._plant_file = os.path.join(STATE_DIR, f"{plant_id}.json")
        self._global_file = os.path.join(STATE_DIR, "global.json")
        self._plant_data = self._load(self._plant_file)
        self._global_data = self._load(self._global_file)

    def _load(self, path):
        try:
            with open(path, "r") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return {}

    def _save_plant(self):
        with open(self._plant_file, "w") as f:
            json.dump(self._plant_data, f, indent=2)

    def _save_global(self):
        with open(self._global_file, "w") as f:
            json.dump(self._global_data, f, indent=2)

    def _is_global(self, job_data):
        return job_data.get("scope") == "global"

    def _get_store(self, job_data):
        return self._global_data if self._is_global(job_data) else self._plant_data

    def get(self, job_key, job_data):
        return self._get_store(job_data).get(job_key, {})

    def mark_requested(self, job_key, job_data):
        self._get_store(job_data)[job_key] = {
            "requested": datetime.now().isoformat(),
            "extracted": None,
            "date": date.today().isoformat()
        }
        if self._is_global(job_data):
            self._save_global()
        else:
            self._save_plant()

    def mark_extracted(self, job_key, job_data):
        store = self._get_store(job_data)
        if job_key in store:
            store[job_key]["extracted"] = datetime.now().isoformat()
            if self._is_global(job_data):
                self._save_global()
            else:
                self._save_plant()

    def should_request(self, job_key, job_data):
        s = self.get(job_key, job_data)
        today = date.today().isoformat()

        if not s or s.get("date") != today:
            return True

        if job_data.get("once_per_day"):
            return False

        if s.get("extracted"):
            return True

        req = s.get("requested")
        if req:
            elapsed = (datetime.now() - datetime.fromisoformat(req)).total_seconds()
            return elapsed >= TIMEOUT

        return True

    def needs_extraction(self, job_key, job_data):
        if not job_data.get("background_job", True):
            return False
        s = self.get(job_key, job_data)
        return (s.get("date") == date.today().isoformat()
                and s.get("requested")
                and not s.get("extracted"))

_last_reported_pct = {}

def report_status(plant_id, job_key, status_text, pct, gui_callback=None):
    global _last_reported_pct
    
    if gui_callback:
        gui_callback(status_text, pct)
        
    if not FIREBASE_URL:
        return
        
    try:
        job_id = f"{plant_id}_{job_key}"
        last_pct = _last_reported_pct.get(job_id, -100)
        
        is_error = "ERRO" in status_text.upper()
        is_finished = pct >= 100
        jumped_enough = abs(pct - last_pct) >= 20 
        
        if is_error or is_finished or jumped_enough:
            payload = {
                "Title": plant_id,
                "JobAtual": job_key,
                "Status": status_text,
                "Concluidos": int(pct),
                "TotalJobs": 100,
                "UltimaAtualizacao": datetime.now().isoformat()
            }
            
            firebase_url = f"{FIREBASE_URL}/rpa_monitor/{plant_id}.json?auth={FIREBASE_SECRET}"
            
            requests.put(firebase_url, json=payload, timeout=3)
            
            _last_reported_pct[job_id] = pct
            
    except Exception as e:
        logging.warning(f"Failed to report status to Cloud: {e}")


def run_plant(plant_id: str, gui_callback=None, stop_event=None):
    pythoncom.CoInitialize()
    logging.info(f"Starting orchestrator for plant: {plant_id}")

    try:
        if plant_id not in settings.config.get("plants", {}):
            logging.error(f"Plant '{plant_id}' not found in config.")
            return

        conn = SAPConnection(plant_id)
        try:
            conn.connect()
            conn.ensure_logged_in()
        except Exception as e:
            err = f"Erro de Conexão SAP: {e}"
            logging.error(err)
            if gui_callback: gui_callback(err, 0)
            return

        # WAKE UP BAT AUTOMATIONS (após SAP carregado e logado)
        try:
            plant_config = settings.config["plants"].get(plant_id, {})
            folder_name = plant_config.get("folder_name", "")
            caminho_onedrive = get_onedrive_path()
            
            if folder_name and caminho_onedrive:
                # O OneDrive corporativo guarda os dados dentro da sub-pasta SESÉ DASHBOARD
                sese_dashboard = os.path.join(caminho_onedrive, "SESÉ DASHBOARD")
                base_dir = sese_dashboard if os.path.exists(sese_dashboard) else caminho_onedrive
                
                bat_dir = os.path.normpath(os.path.join(base_dir, folder_name, "002 - Automacao", "Manual"))
                if os.path.exists(bat_dir):
                    bats_found = False
                    for file in os.listdir(bat_dir):
                        if file.lower().endswith('.bat'):
                            bat_path = os.path.join(bat_dir, file)
                            subprocess.Popen(bat_path, cwd=bat_dir, shell=True, creationflags=0x08000000)
                            msg = f"Processo .BAT ativado: {file}"
                            logging.info(msg)
                            if gui_callback: gui_callback(msg, 5)
                            bats_found = True
                    if not bats_found:
                        logging.info("Nenhum arquivo .bat encontrado na pasta de automação.")
                else:
                    logging.info(f"Pasta de automação manual não existe: {bat_dir}")
        except Exception as e:
            err_msg = f"Aviso: Não foi possível rodar automações BAT secundárias: {e}"
            logging.warning(err_msg)
            if gui_callback: gui_callback(err_msg, 5)
            


        jobs = settings.config.get("jobs", {})

        while stop_event is None or not stop_event.is_set():
            watchdog_infraestrutura()
            
            if not conn.check_connection():
                logging.warning("Conexão SAP inativa. Tentando reconectar...")
                try:
                    conn.connect()
                    conn.ensure_logged_in()
                except Exception as e:
                    logging.error(f"Falha ao reconectar: {e}")
                    if stop_event: stop_event.wait(60)
                    else: threading.Event().wait(60)
                    continue

            state = JobState(plant_id)
            t0 = datetime.now()
            logging.info(f"=== Cycle {t0.strftime('%H:%M:%S')} ===")
            
            cycle_has_error = False

            jobs_to_request = [
                k for k, v in jobs.items() 
                if v.get("active") 
                and plant_id in v.get("plant_params", {}) 
                and (v.get("scope") != "global" or plant_id == MASTER_PLANT) 
                and state.should_request(k, v)
            ]
            req_total = len(jobs_to_request)
            req_count = 0

            for job_key, job_data in jobs.items():
                if stop_event and stop_event.is_set(): break
                if not job_data.get("active") or plant_id not in job_data.get("plant_params", {}):
                    continue

                if job_data.get("scope") == "global" and plant_id != MASTER_PLANT:
                    continue

                if not state.should_request(job_key, job_data):
                    logging.info(f"[SKIP] {job_key}")
                    continue

                t_code = job_data.get("transaction")
                
                pct = (req_count / req_total * 30) if req_total > 0 else 30
                report_status(plant_id, job_key, f"Solicitando relatório ({t_code})...", pct, gui_callback)
                
                logging.info(f"[REQUEST] {job_key} ({t_code})")
                try:
                    job_timeout = job_data.get("timeout", DEFAULT_JOB_TIMEOUT)
                    conn.start_transaction(t_code)
                    func = _get_job_handler(job_key, job_data)
                    if func:
                        with JobWatchdog(timeout=job_timeout, job_name=job_key):
                            func(conn.session, plant_id, job_key)
                        state.mark_requested(job_key, job_data)
                        logging.info(f"[OK] {job_key}")
                        
                        pct = (req_count / req_total * 30) if req_total > 0 else 30
                        report_status(plant_id, job_key, "Solicitação concluída.", pct, gui_callback)
                    else:
                        logging.warning(f"[WARN] No handler for '{job_key}'")
                except Exception as e:
                    logging.error(f"[ERROR] {job_key}: {e}")
                    report_status(plant_id, job_key, f"ERRO na solicitação: {str(e)[:50]}", pct, gui_callback)
                    cycle_has_error = True

                req_count += 1

            pending = [(k, v) for k, v in jobs.items()
                        if v.get("active")
                        and plant_id in v.get("plant_params", {})
                        and (v.get("scope") != "global" or plant_id == MASTER_PLANT)
                        and state.needs_extraction(k, v)]

            if pending and (not stop_event or not stop_event.is_set()):
                logging.info("[WAIT] Aguardando 120s para jobs processarem...")
                
                for i in range(4):
                    if stop_event and stop_event.is_set(): break
                    current_pct = 30 + (i * 7.5) 
                    if not cycle_has_error:
                        report_status(plant_id, "SISTEMA", f"Processando SAP... ({(i+1)*30}s/120s)", current_pct, gui_callback)
                    if stop_event: stop_event.wait(30)
                    else: threading.Event().wait(30)

                if (not stop_event or not stop_event.is_set()) and not conn.check_connection():
                    logging.warning("Conexão SAP inativa após espera de 120s. Tentando reconectar para extração...")
                    try:
                        conn.connect()
                        conn.ensure_logged_in()
                    except Exception as e:
                        logging.error(f"Falha ao reconectar durante extração: {e}")

                if (not stop_event or not stop_event.is_set()) and conn.check_connection():
                    logging.info(f"[SP02] {len(pending)} pending")
                    try:
                        conn.start_transaction("SP02")
                        ext_total = len(pending)
                        ext_count = 0
                        for job_key, job_data in pending:
                            if stop_event and stop_event.is_set(): break
                            pct = 60 + (ext_count / ext_total * 40) if ext_total > 0 else 60
                            report_status(plant_id, job_key, "Extraindo spool na SP02...", pct, gui_callback)
                            logging.info(f"[EXTRACT] {job_key}")
                            ext_timeout = job_data.get("timeout", DEFAULT_JOB_TIMEOUT)
                            with JobWatchdog(timeout=ext_timeout, job_name=f"SP02_{job_key}"):
                                extracted = extract_sp02_job(conn.session, plant_id, job_key, job_data)
                            if extracted:
                                state.mark_extracted(job_key, job_data)
                                logging.info(f"[OK] {job_key} extracted")
                                
                                ext_count += 1
                                pct = 60 + (ext_count / ext_total * 40) if ext_total > 0 else 60
                                report_status(plant_id, job_key, "Extração salva com sucesso!", pct, gui_callback)
                            else:
                                ext_count += 1
                    except Exception as e:
                        logging.error(f"[ERROR] SP02: {e}")
                        report_status(plant_id, "SP02", f"ERRO na extração: {str(e)[:50]}", 60, gui_callback)
                        cycle_has_error = True

            if stop_event and stop_event.is_set(): break

            wait = int(CYCLE_WAIT - (datetime.now() - t0).total_seconds())
            if wait > 0:
                if not cycle_has_error:
                    report_status(plant_id, "SISTEMA", f"Repouso. Próximo ciclo em {wait}s.", 100, gui_callback)
                else:
                    report_status(plant_id, "SISTEMA", f"Repouso após erros. Próximo ciclo em {wait}s.", 100, gui_callback)
                    
                logging.info(f"=== Next cycle in {wait}s ===")
                if stop_event: stop_event.wait(wait)
                else: threading.Event().wait(wait)

    finally:
        logging.info("Orchestrator shutting down.")
        if gui_callback: gui_callback("Trabalho finalizado ou interrompido.", 0)
        pythoncom.CoUninitialize()
