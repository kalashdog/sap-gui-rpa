import os
import json
import requests
from pathlib import Path
from dotenv import load_dotenv
import logging

load_dotenv()

# Tenta carregar a URL do Firebase (do bootstrap seguro ou do env local)
FIREBASE_URL = os.getenv("FIREBASE_URL", "")
try:
    from core.bootstrap import FIREBASE_URL as B_URL # type: ignore
    if B_URL:
        FIREBASE_URL = B_URL
except ImportError:
    pass

class Settings:
    def __init__(self):
        app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        self.cache_dir = Path(app_data) / "HubSeseRPA" / "config"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.cache_path = self.cache_dir / "rpa_jobs_cache.json"
        
        self.config = self._load_config()
        
        self.dynamic_user = None
        self.dynamic_pwd = None
        self.export_base_path = None

    def _load_config(self):
        if not FIREBASE_URL:
            logging.warning("FIREBASE_URL não configurado. Tentando carregar do cache local...")
            return self._load_from_cache()

        url = f"{FIREBASE_URL}/rpa_jobs.json"
        try:
            response = requests.get(url, timeout=10)
            response.raise_for_status()
            data = response.json()
            if not data:
                raise ValueError("JSON vazio retornado do Firebase.")
            
            # Salva o novo cache
            with open(self.cache_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
                
            return data
            
        except Exception as e:
            logging.warning(f"Falha ao carregar configuração do Firebase ({e}). Tentando fallback offline...")
            return self._load_from_cache()

    def _load_from_cache(self):
        # Fallback de dev (se existir o arquivo antigo na pasta local)
        local_dev_path = Path(__file__).parent / "rpa_jobs.json"
        target_path = self.cache_path if self.cache_path.exists() else local_dev_path
            
        if not target_path.exists():
            raise RuntimeError("Nenhum cache de configuração encontrado e Firebase indisponível. Impossível iniciar.")
            
        with open(target_path, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError as e:
                raise ValueError(f"Erro ao analisar o cache json: {e}")

    def get_credentials(self, plant_id: str) -> tuple:
        if self.dynamic_user and self.dynamic_pwd:
            return (self.dynamic_user, self.dynamic_pwd)

        try:
            code = self.config["plants"][plant_id]["code"]
        except KeyError:
            raise ValueError(f"Planta '{plant_id}' não encontrada na configuração.")
            
        user = os.getenv(f"{code}_USER")
        password = os.getenv(f"{code}_PASS")
        
        if not user or not password:
            raise ValueError(f"Credenciais ausentes para a planta '{plant_id}' (esperado {code}_USER e {code}_PASS).")
            
        return (user, password)

settings = Settings()
