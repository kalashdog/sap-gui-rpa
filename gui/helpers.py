import os
import json
import keyring
import ctypes

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

def setup_window(app):
    from gui.constants import APP_TITLE, WIN_W, WIN_H, BG_APP, get_asset_path
    app.title(APP_TITLE)
    app.resizable(False, False)
    app.configure(fg_color=BG_APP)
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('sese.rpa.logistica.v2')
        icon_path = get_asset_path(".assets/rpaseselogo_perfect.ico")
        if os.path.exists(icon_path): app.iconbitmap(icon_path)
    except Exception: pass
    
    x = (app.winfo_screenwidth() - WIN_W) // 2
    y = (app.winfo_screenheight() - WIN_H) // 2
    app.geometry(f"{WIN_W}x{WIN_H}+{x}+{y}")


def build_shell(app):
    import webbrowser
    import customtkinter as ctk
    from PIL import Image
    from gui.constants import (
        APP_VERSION, SHAREPOINT_LINK,
        FONT, SIDEBAR_W,
        ACCENT, PANEL, BORDER, TEXT, TEXT_MUTED,
        SIDEBAR_BG, SIDEBAR_MUTED, SIDEBAR_CARD_BG, SIDEBAR_CARD_BORDER,
        SIDEBAR_BTN_HOVER, get_asset_path
    )

    # Base grid
    app.grid_columnconfigure(0, weight=0, minsize=SIDEBAR_W)
    app.grid_columnconfigure(1, weight=1)
    app.grid_rowconfigure(0, weight=1)

    # Sidebar
    sb = ctk.CTkFrame(app, width=SIDEBAR_W, corner_radius=0, fg_color=SIDEBAR_BG)
    sb.grid(row=0, column=0, sticky="nsew")
    sb.grid_propagate(False)
    app.sidebar = sb

    app.sb_top = ctk.CTkFrame(sb, fg_color="transparent")
    app.sb_top.pack(side="top", fill="x")

    brand = ctk.CTkFrame(app.sb_top, fg_color="transparent")
    brand.pack(fill="x", padx=24, pady=(24, 16))

    try:
        img = Image.open(get_asset_path(".assets/sese_white.png"))
        ctk.CTkLabel(brand, image=ctk.CTkImage(light_image=img, dark_image=img, size=(160, 48)), text="").pack(anchor="w")
    except Exception:
        ctk.CTkLabel(brand, text="RPA SESÉ", width=48, height=48, corner_radius=14, fg_color=ACCENT, text_color="white", font=(FONT, 18, "bold")).pack(anchor="w")

    ctk.CTkLabel(brand, text="Hub de Dashboards", font=(FONT, 22, "bold"), text_color="white").pack(anchor="w", pady=(12, 0))
    ctk.CTkLabel(brand, text="RPA SESÉ • SAP", font=(FONT, 12), text_color=SIDEBAR_MUTED).pack(anchor="w", pady=(2, 0))

    app.env_card = ctk.CTkFrame(app.sb_top, corner_radius=12, fg_color=SIDEBAR_CARD_BG, border_width=1, border_color=SIDEBAR_CARD_BORDER)
    app.env_card.pack(fill="x", padx=24, pady=(0, 12))

    env_hdr = ctk.CTkFrame(app.env_card, fg_color="transparent")
    env_hdr.pack(fill="x", padx=14, pady=(12, 8))

    ctk.CTkLabel(env_hdr, text="AMBIENTE", font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED).pack(side="left")
    ctk.CTkButton(env_hdr, text="🔄", width=28, height=28, corner_radius=6, fg_color="transparent", hover_color=SIDEBAR_BTN_HOVER, text_color="#CBD5E1", font=(FONT, 14), command=lambda: app._check_env(navigate=not app._running())).pack(side="right", padx=(4, 0))
    ctk.CTkButton(env_hdr, text="🌐", width=28, height=28, corner_radius=6, fg_color="transparent", hover_color=SIDEBAR_BTN_HOVER, text_color="#CBD5E1", font=(FONT, 14), command=lambda: webbrowser.open_new_tab(SHAREPOINT_LINK)).pack(side="right")

    app.env_badge = ctk.CTkLabel(env_hdr, text="Verificando…", height=22, corner_radius=999, fg_color=("#1E293B", "#1E293B"), text_color="#E2E8F0", font=(FONT, 10, "bold"))
    app.env_badge.pack(side="right", padx=(0, 8))

    app.env_path = ctk.CTkLabel(app.env_card, text="Verificando OneDrive…", wraplength=SIDEBAR_W - 76, justify="left", font=(FONT, 11), text_color="#CBD5E1")
    app.env_path.pack(anchor="w", padx=14, pady=(0, 14))

    app.sb_bottom = ctk.CTkFrame(sb, fg_color="transparent")
    app.sb_bottom.pack(side="bottom", fill="x")

    ctk.CTkLabel(app.sb_bottom, text=f"v{APP_VERSION}  •  VINICIUS LIMA", font=(FONT, 10), text_color="#475569").pack(side="bottom", anchor="w", padx=24, pady=(8, 18))

    seg = ctk.CTkSegmentedButton(app.sb_bottom, values=["Sistema", "Claro", "Escuro"], command=app._set_theme, corner_radius=12)
    seg.pack(side="bottom", fill="x", padx=24, pady=(4, 12))
    seg.set(app._prefs.get("theme", "Escuro"))

    ctk.CTkLabel(app.sb_bottom, text="TEMA", font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED).pack(side="bottom", anchor="w", padx=24, pady=(8, 4))

    app.jobs_container = ctk.CTkScrollableFrame(sb, fg_color="transparent", corner_radius=0)

    # Main panel
    app.main_panel = ctk.CTkFrame(app, fg_color="transparent")
    app.main_panel.grid(row=0, column=1, sticky="nsew", padx=(12, 18), pady=18)
    app.main_panel.grid_rowconfigure(0, weight=1)
    app.main_panel.grid_columnconfigure(0, weight=1)

    app.main_card = ctk.CTkFrame(app.main_panel, corner_radius=20, fg_color=PANEL, border_width=1, border_color=BORDER)
    app.main_card.grid(row=0, column=0, sticky="nsew")
    app.main_card.grid_rowconfigure(2, weight=1)
    app.main_card.grid_columnconfigure(0, weight=1)

    hdr = ctk.CTkFrame(app.main_card, fg_color="transparent")
    hdr.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 10))

    app.page_tag = ctk.CTkLabel(hdr, text="", font=(FONT, 10, "bold"), text_color=ACCENT)
    app.page_tag.pack(anchor="w")

    app.page_title = ctk.CTkLabel(hdr, text="", font=(FONT, 22, "bold"), text_color=TEXT)
    app.page_title.pack(anchor="w", pady=(2, 0))

    app.page_sub = ctk.CTkLabel(hdr, text="", font=(FONT, 11), text_color=TEXT_MUTED, wraplength=320, justify="left")
    app.page_sub.pack(anchor="w", pady=(4, 0))

    ctk.CTkFrame(app.main_card, height=1, fg_color=BORDER).grid(row=1, column=0, sticky="ew", padx=24)

    app.content = ctk.CTkFrame(app.main_card, fg_color="transparent")
    app.content.grid(row=2, column=0, sticky="nsew", padx=24, pady=(12, 20))
