import os
import json
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

class Settings:
    def __init__(self):
        self.config_path = Path(__file__).parent / "sapscripts_config.json"
        self.config = self._load_config()
        
        self.dynamic_user = None
        self.dynamic_pwd = None
        self.export_base_path = None

    def _load_config(self):
        if not self.config_path.exists():
            raise RuntimeError("Arquivo de configuração ausente. Não é possível iniciar o RPA.")
            
        with open(self.config_path, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError as e:
                raise ValueError(f"Erro ao analisar sapscripts_config.json: {e}")

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
