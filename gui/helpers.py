import os
import json
import keyring

def load_prefs() -> dict:
    from gui.constants import PREFS_FILE
    try:
        if os.path.exists(PREFS_FILE):
            with open(PREFS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def save_prefs(prefs: dict) -> None:
    from gui.constants import PREFS_FILE
    try:
        with open(PREFS_FILE, "w", encoding="utf-8") as f:
            json.dump(prefs, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

def theme_to_mode(label: str) -> str:
    return {"Sistema": "system", "Claro": "light", "Escuro": "dark"}.get(label, "dark")

def short_path(path: str, n: int = 48) -> str:
    if not path:
        return "-"
    return path if len(path) <= n else f"…{path[-n:]}"

def safe_del_pwd(service: str, username: str) -> None:
    try:
        keyring.delete_password(service, username)
    except Exception:
        pass
