import os
import sys
import json
import subprocess
import shutil
import tkinter as tk
import winreg
from tkinter import messagebox

APP_DATA = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
BIN_DIR = os.path.join(APP_DATA, "HubSeseRPA", "bin")
VERSION_FILE = os.path.join(APP_DATA, "HubSeseRPA", "current_version.txt")

os.makedirs(BIN_DIR, exist_ok=True)

def get_onedrive_updates_folder():
    """Search the hidden OneDrive update folder in OneDrive local or SharePoint via Registry"""
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
    except Exception:
        pass

    for env in ["OneDriveCommercial", "OneDrive", "OneDriveConsumer"]:
        val = os.environ.get(env)
        if val and os.path.exists(val):
            possible_paths.add(val)
            
    for path in possible_paths:
        sese_path = os.path.join(path, "SESÉ DASHBOARD")
        if os.path.exists(sese_path):
            return os.path.join(sese_path, "002 - Filiais database", "000 - Global", ".rpa_update")
            
    return None

def show_msg(title, message, is_error=False):
    root = tk.Tk()
    root.withdraw()
    if is_error:
        messagebox.showerror(title, message)
    else:
        messagebox.showinfo(title, message)
    root.destroy()

def get_current_version():
    if os.path.exists(VERSION_FILE):
        with open(VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return "0.0.0"

def set_current_version(version):
    with open(VERSION_FILE, 'w', encoding='utf-8') as f:
        f.write(version)

def main():
    try:
        updates_dir = get_onedrive_updates_folder()
        if not updates_dir or not os.path.exists(updates_dir):
            raise Exception("A pasta do SharePoint (SESÉ DASHBOARD) não foi encontrada ou não está sincronizada no seu computador.")
        info_file = os.path.join(updates_dir, "update_info.json")
        if not os.path.exists(info_file):
            raise Exception("O arquivo de configuração de versão (update_info.json) não está na pasta.")

        with open(info_file, 'r', encoding='utf-8') as f:
            cloud_data = json.load(f)

        cloud_version = cloud_data.get("version", "0.0.0").strip()
        cloud_filename = cloud_data.get("filename", "").strip()
        cloud_status = cloud_data.get("status", "ATIVO").strip().upper()
        cloud_message = cloud_data.get("message", "").strip()

        if cloud_status == "INATIVO":
            show_msg("Sistema em Manutenção", f"O RPA está temporariamente suspenso.\n\nMotivo: {cloud_message}")
            sys.exit(0)
        local_version = get_current_version()
        local_exe_path = os.path.join(BIN_DIR, cloud_filename)

        if cloud_version != local_version or not os.path.exists(local_exe_path):
            source_exe = os.path.join(updates_dir, cloud_filename)

            if not os.path.exists(source_exe):
                show_msg("Aguardando Sincronização", f"A versão {cloud_version} foi lançada, mas o seu OneDrive ainda está a descarregar o ficheiro.\n\nAguarde o ícone da nuvem do Windows terminar e tente novamente.")
                sys.exit(0)

            show_msg("Atualização Detectada", f"A instalar a versão {cloud_version} do Hub Sesé RPA...\n\nNovidades: {cloud_message}")

            shutil.copy2(source_exe, local_exe_path)
            set_current_version(cloud_version)


        args = [local_exe_path] + sys.argv[1:]
        subprocess.Popen(args)

    except Exception as e:
        local_version = get_current_version()
        if local_version != "0.0.0":
            exes = [f for f in os.listdir(BIN_DIR) if f.endswith('.exe')]
            if exes:
                subprocess.Popen([os.path.join(BIN_DIR, exes[0])] + sys.argv[1:])
                sys.exit(0)
                
        show_msg("Erro de Inicialização", f"Não foi possível iniciar o RPA.\n\nDetalhe: {e}", True)

if __name__ == "__main__":
    main()
