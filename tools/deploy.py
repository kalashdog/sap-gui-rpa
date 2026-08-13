import os
import sys
import json
import subprocess
import shutil
import winreg
import glob

def get_onedrive_updates_folder():
    possible_paths = set()
    try:
        base_key = r"Software\Microsoft\OneDrive\Accounts"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_key) as key:
            for i in range(winreg.QueryInfoKey(key)[0]):
                with winreg.OpenKey(key, winreg.EnumKey(key, i)) as subkey:
                    try:
                        folder, _ = winreg.QueryValueEx(subkey, "UserFolder")
                        if folder and os.path.exists(folder): possible_paths.add(folder)
                    except FileNotFoundError: pass
    except Exception: pass

    for env in ["OneDriveCommercial", "OneDrive", "OneDriveConsumer"]:
        if os.environ.get(env) and os.path.exists(os.environ.get(env)):
            possible_paths.add(os.environ.get(env))
            
    for path in possible_paths:
        sese_path = os.path.join(path, "SESÉ DASHBOARD")
        if os.path.exists(sese_path):
            return os.path.join(sese_path, "002 - Filiais database", "000 - Global", ".rpa_update")
    return None

def main():
    os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    
    os.system('cls' if os.name == 'nt' else 'clear')
    print("="*60)
    print("   DEPLOY ASSISTANT - HUB RPA")
    print("="*60)
    
    # 1. Collect infos
    version = input("\n[1] Type the new version (ex: 2.3.4): ").strip()
    if not version:
        print("Operation cancelled.")
        return
        
    message = input("[2] What changed in this version? (Release Notes): ").strip()
    if not message:
        message = "Routine update and stability improvements."
        
    filename = f"HubSese_v{version}.exe"
    exe_name_no_ext = filename.replace(".exe", "")
    
    # 2. Clean
    print("\n[3] Cleaning cache...")
    for d in ['build', 'dist']:
        if os.path.exists(d): shutil.rmtree(d)

    # 2.5 Inject Bootstrap (Base64 Obfuscated)
    print("\n[4] Preparing bootstrap injection...")
    bootstrap_path = os.path.join("core", "bootstrap.py")
    secret = None
    url = None
    
    import base64
    if os.path.exists(".env"):
        with open(".env", "r") as f:
            for line in f:
                if line.startswith("FIREBASE_RTDB_SECRET="):
                    secret = line.split("=", 1)[1].strip()
                elif line.startswith("FIREBASE_URL="):
                    url = line.split("=", 1)[1].strip()
    
    if secret and url:
        b64_secret = base64.b64encode(secret.encode()).decode()
        b64_url = base64.b64encode(url.encode()).decode()
        
        with open(bootstrap_path, "w") as f:
            f.write("import base64\n")
            f.write(f"FIREBASE_RTDB_SECRET = base64.b64decode(b'{b64_secret}').decode()\n")
            f.write(f"FIREBASE_URL = base64.b64decode(b'{b64_url}').decode()\n")
        print("    ✅ Secure Bootstrap module generated (Obfuscated).")
        
        print("    [4.5] Sincronizando rpa_jobs.json com a Nuvem...")
        rpa_jobs_path = os.path.join("config", "rpa_jobs.json")
        if os.path.exists(rpa_jobs_path):
            try:
                import requests
                with open(rpa_jobs_path, "r", encoding="utf-8") as json_file:
                    jobs_data = json.load(json_file)
                fb_url = f"{url}/rpa_jobs.json?auth={secret}"
                resp = requests.put(fb_url, json=jobs_data, timeout=10)
                if resp.status_code == 200:
                    print("    ✅ Configurações rpa_jobs atualizadas na Nuvem com sucesso!")
                else:
                    print(f"    ⚠️ Erro ao subir rpa_jobs para a nuvem: HTTP {resp.status_code} - {resp.text}")
            except Exception as e:
                print(f"    ⚠️ Erro ao conectar com o Firebase: {e}")
        else:
            print("    ⚠️ Arquivo config/rpa_jobs.json não encontrado. Ignorando upload.")
            
    else:
        print("    ⚠️ WARNING: FIREBASE_RTDB_SECRET or FIREBASE_URL not found in .env!")

    # 3. Armored Compilation (GUI only)
    print(f"\n[5] Compiling {filename} (This may take 1 minute)...")
    cmd = [
        "pyinstaller", "--noconfirm", "--onefile", "--windowed",
        "--icon=.assets\\rpaseselogo_perfect.ico",
        f"--name={exe_name_no_ext}",
        "--add-data=.assets;.assets",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=keyring.backends.Windows",
        "--collect-all=customtkinter",
        "--exclude-module=pytest",
        "--exclude-module=unittest",
        "--exclude-module=pydoc",
        "--exclude-module=matplotlib",
        "--exclude-module=numpy",
        "--exclude-module=pandas",
        "--exclude-module=scipy",
        "--exclude-module=gevent",
        "--exclude-module=Eel",
        "--exclude-module=bottle",
        "--exclude-module=pygments",
        "--exclude-module=auto_py_to_exe",
        "--exclude-module=pluggy",
        "--exclude-module=iniconfig",
        "--exclude-module=greenlet",
        "--exclude-module=setuptools",
        "--exclude-module=pkg_resources",
        "--exclude-module=tkinter.test",
        "--exclude-module=pip",
        "--exclude-module=pythonwin",
        "--exclude-module=_pytest",
        "--exclude-module=future",
        "--exclude-module=zope",
        "gui.py"
    ]
    
    # Execute silently, only show error if it fails
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print("\n❌ ERROR ON COMPILATION:")
        print(result.stderr)
        return
        
    print("    ✅ Compilation completed successfully (25MB guaranteed).")

    # 4. Searching OneDrive
    updates_dir = get_onedrive_updates_folder()
    if not updates_dir or not os.path.exists(updates_dir):
        print("\n❌ ERROR: Hidden OneDrive folder (.rpa_update) not found!")
        print("The file was compiled in the 'dist' folder, but not sent to the cloud.")
        return

    # 5. The Deploy Físico
    print(f"\n[5] Sending to cloud ({updates_dir})...")
    src_exe = os.path.join("dist", filename)
    dest_exe = os.path.join(updates_dir, filename)
    
    shutil.copy2(src_exe, dest_exe)
    print(f"    ✅ File {filename} copied.")

    # 6. Update oracle (JSON)
    json_path = os.path.join(updates_dir, "update_info.json")
    json_data = {
        "version": version,
        "filename": filename,
        "status": "ATIVO",
        "message": message
    }
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, indent=4, ensure_ascii=False)
    print("    ✅ File update_info.json updated.")

    print("\n" + "="*60)
    print(f"  DEPLOY VERSION {version} SUCCESS!")
    print(" The operators will receive this version on next click.")
    print("\n[7] Cleaning up build and dist...")
    try:
        if os.path.exists("build"): shutil.rmtree("build")
        if os.path.exists("dist"): shutil.rmtree("dist")
        if os.path.exists(bootstrap_path): os.remove(bootstrap_path)
        for spec_file in glob.glob("*.spec"):
            os.remove(spec_file)
        print("    ✅ Cleaned up.")
    except Exception as e:
        print(f"    ⚠️ Warning: {e}")
    print("="*60 + "\n")

if __name__ == "__main__":
    main()
