import os
import winreg
import sys
import logging
from config.settings import settings

def get_onedrive_path():
    possible_paths = set()
    
    try:
        base_key = r"Software\Microsoft\OneDrive\Accounts"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_key) as key:
            num_subkeys = winreg.QueryInfoKey(key)[0] 
            
            for i in range(num_subkeys):
                subkey_name = winreg.EnumKey(key, i) 
                with winreg.OpenKey(key, subkey_name) as subkey:
                    try:
                        user_folder, _ = winreg.QueryValueEx(subkey, "UserFolder")
                        if user_folder and os.path.exists(user_folder):
                            possible_paths.add(user_folder)
                    except FileNotFoundError:
                        pass
    except Exception as e:
        logging.warning(f"Erro ao ler contas do OneDrive: {e}")

    for env in ["OneDriveCommercial", "OneDrive", "OneDriveConsumer"]:
        val = os.environ.get(env)
        if val and os.path.exists(val):
            possible_paths.add(val)

    for path in possible_paths:
        low_path = path.lower()
        if "sese" in low_path or "sesé" in low_path:
            return path
            
    for path in possible_paths:
        if os.path.exists(os.path.join(path, "SESÉ DASHBOARD")):
            return path
    return list(possible_paths)[0] if possible_paths else None


def get_target_export_dir(plant_id: str) -> str:
    base_dir = getattr(settings, 'export_base_path', None)
    
    if not base_dir:
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
            
    plant_path = os.path.join(base_dir, plant_id)
    os.makedirs(plant_path, exist_ok=True)
    
    return plant_path