"""
RpaGUI - Main application window.

Orchestrates the sidebar, navigation, and page lifecycle.
Pages are rendered by the modules in gui.pages.
"""
import os
import sys
import re
import threading
import webbrowser
import ctypes
import time
import subprocess
from datetime import datetime, timedelta
from tkinter import messagebox
import tkinter as tk

import customtkinter as ctk
import keyring
from PIL import Image

from core.orchestrator import run_plant
from config.settings import settings
from core.utils import get_onedrive_path

from gui.constants import (
    APP_TITLE, APP_VERSION, EXPECTED_FOLDER, SHAREPOINT_LINK,
    FONT, FONT_MONO, SIDEBAR_W, MAIN_W, WIN_W, WIN_H,
    ES_CONTINUOUS, ES_SYSTEM_REQUIRED, ES_DISPLAY_REQUIRED,
    ACCENT, ACCENT_HOVER, ACCENT_SOFT, ACCENT_TEXT,
    BG_APP, PANEL, PANEL_ALT, BORDER, TEXT, TEXT_MUTED,
    SIDEBAR_BG, SIDEBAR_MUTED, SIDEBAR_CARD_BG, SIDEBAR_CARD_BORDER,
    SIDEBAR_BTN_HOVER, SECONDARY, SECONDARY_HOVER,
    SUCCESS_FG, SUCCESS_TEXT, WARNING_FG, WARNING_TEXT,
    DANGER, DANGER_HOVER, ERROR_FG, ERROR_TEXT,
    get_asset_path,
)
from gui.helpers import load_prefs, save_prefs, theme_to_mode, short_path, safe_del_pwd

from gui.pages import setup as setup_page
from gui.pages import login as login_page
from gui.pages import progress as progress_page


class RpaGUI(ctk.CTk):

    def __init__(self):
        self._prefs = load_prefs()
        ctk.set_appearance_mode(theme_to_mode(self._prefs.get("theme", "Escuro")))

        super().__init__()

        self.title(APP_TITLE)
        self.geometry(f"{WIN_W}x{WIN_H}")

        try:
            myappid = 'sese.rpa.logistica.v2'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        try:
            icon_path = get_asset_path(".assets/rpaseselogo_perfect.ico")
            if os.path.exists(icon_path):
                self.iconbitmap(icon_path)
        except Exception as e:
            print(f"Erro ao carregar ícone: {e}")

        self.resizable(False, False)
        self.configure(fg_color=BG_APP)
        self._center(WIN_W, WIN_H)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # State
        self.stop_event = threading.Event()
        self.worker_thread = None
        self._execution_gen = 0
        self.onedrive_base = None
        self.caminho_onedrive = None
        self.env_ready = False
        self.current_pct = 0
        self.last_status = None
        self.execution_finished = False
        self.current_user = ""
        self.current_plant = ""
        self.password_visible = False
        self.timer_running = False
        self.start_time = None
        self.remember_var = tk.BooleanVar(value=True)
        self.autostart_var = tk.BooleanVar(value=self._prefs.get("autostart", False))

        self._reset_refs()
        self._build_shell()
        self._check_env(navigate=True)

        self._prevent_sleep()

        if "--autostart" in sys.argv:
            print("Autostart detetado. Aguardando 30s para estabilização da rede...")
            self.after(30000, self.trigger_auto_start)

    # Dynamic widget references
    def _reset_refs(self):
        for attr in (
            "combo_planta", "input_user", "input_pwd", "form_msg",
            "btn_start", "btn_toggle_pwd", "lbl_status", "lbl_percent",
            "lbl_timer", "progress_bar", "status_badge", "log_box",
            "btn_stop", "btn_back", "lbl_hint",
        ):
            setattr(self, attr, None)

    def _ok(self, w) -> bool:
        try:
            return w is not None and w.winfo_exists()
        except Exception:
            return False

    def _running(self) -> bool:
        return self.worker_thread is not None and self.worker_thread.is_alive()

    def _center(self, w: int, h: int) -> None:
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _on_close(self):
        if self._running():
            if not messagebox.askyesno(
                "Encerrar", "O robô está em execução.\nDeseja parar e fechar?"
            ):
                return
            self.stop_event.set()
        self.timer_running = False
        self._allow_sleep()
        self.destroy()

    def _prevent_sleep(self):
        """Injects the 'Active Display' command into the Windows Kernel."""
        try:
            if sys.platform == "win32":
                ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED)
        except Exception as e:
            print(f"Erro ao bloquear suspensão: {e}")

    def _allow_sleep(self):
        """Returns power control to Windows."""
        try:
            if sys.platform == "win32":
                ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        except Exception:
            pass

    def toggle_autostart(self):
        import win32com.client
        startup_folder = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs\Startup")
        shortcut_path = os.path.join(startup_folder, "HubSeseRPA.lnk")

        if self.caminho_onedrive:
            target_path = os.path.join(self.caminho_onedrive, "002 - Filiais database", "007 - RPA SAP", "RPA_Sese.exe")
        else:
            target_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)

        if self.autostart_var.get():
            try:
                # fallback pra versao velha ja baixada se der pau no onedrive
                if not os.path.exists(target_path):
                    target_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
                    
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = target_path
                shortcut.Arguments = "--autostart"
                shortcut.WorkingDirectory = os.path.dirname(target_path)
                shortcut.IconLocation = target_path
                shortcut.save()
            except Exception as e:
                print(f"Erro ao criar atalho: {e}")
        else:
            if os.path.exists(shortcut_path):
                os.remove(shortcut_path)

        self._prefs["autostart"] = self.autostart_var.get()
        save_prefs(self._prefs)

    def trigger_auto_start(self):
        if keyring.get_password("RPA_SESE_USER", "default"):
            self._start()

    
    #  SHELL
    def _build_shell(self):
        self.grid_columnconfigure(0, weight=0, minsize=SIDEBAR_W)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        #  SIDEBAR 
        self.sidebar = ctk.CTkFrame(
            self, width=SIDEBAR_W, corner_radius=0, fg_color=SIDEBAR_BG
        )
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_propagate(False)
        self._build_sidebar()

        #  MAIN PANEL   
        self.main_panel = ctk.CTkFrame(self, fg_color="transparent")
        self.main_panel.grid(row=0, column=1, sticky="nsew", padx=(12, 18), pady=18)
        self.main_panel.grid_rowconfigure(0, weight=1)
        self.main_panel.grid_columnconfigure(0, weight=1)

        self.main_card = ctk.CTkFrame(
            self.main_panel, corner_radius=20,
            fg_color=PANEL, border_width=1, border_color=BORDER
        )
        self.main_card.grid(row=0, column=0, sticky="nsew")
        self.main_card.grid_rowconfigure(2, weight=1)
        self.main_card.grid_columnconfigure(0, weight=1)

        # Header
        hdr = ctk.CTkFrame(self.main_card, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew", padx=24, pady=(20, 10))

        self.page_tag = ctk.CTkLabel(
            hdr, text="", font=(FONT, 10, "bold"), text_color=ACCENT
        )
        self.page_tag.pack(anchor="w")

        self.page_title = ctk.CTkLabel(
            hdr, text="", font=(FONT, 22, "bold"), text_color=TEXT
        )
        self.page_title.pack(anchor="w", pady=(2, 0))

        self.page_sub = ctk.CTkLabel(
            hdr, text="", font=(FONT, 11), text_color=TEXT_MUTED,
            wraplength=MAIN_W - 60, justify="left"
        )
        self.page_sub.pack(anchor="w", pady=(4, 0))

        ctk.CTkFrame(
            self.main_card, height=1, fg_color=BORDER
        ).grid(row=1, column=0, sticky="ew", padx=24)

        self.content = ctk.CTkFrame(self.main_card, fg_color="transparent")
        self.content.grid(row=2, column=0, sticky="nsew", padx=24, pady=(12, 20))

    #  Sidebar  
    def _build_sidebar(self):
        sb = self.sidebar
        wrap = SIDEBAR_W - 48

        #  Brand  
        brand = ctk.CTkFrame(sb, fg_color="transparent")
        brand.pack(fill="x", padx=24, pady=(24, 16))

        try:
            logo_path = get_asset_path(".assets/sese_white.png")
            logo_img = ctk.CTkImage(
                light_image=Image.open(logo_path),
                dark_image=Image.open(logo_path),
                size=(160, 48)
            )
            logo_label = ctk.CTkLabel(brand, image=logo_img, text="")
            logo_label.pack(anchor="w")
        except Exception:
            badge = ctk.CTkLabel(
                brand, text="RPA SESÉ", width=48, height=48,
                corner_radius=14, fg_color=ACCENT,
                text_color="white", font=(FONT, 18, "bold")
            )
            badge.pack(anchor="w")

        ctk.CTkLabel(
            brand, text="Hub de Dashboards",
            font=(FONT, 22, "bold"), text_color="white"
        ).pack(anchor="w", pady=(12, 0))

        ctk.CTkLabel(
            brand, text="RPA SESÉ • SAP",
            font=(FONT, 12), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", pady=(2, 0))

        ctk.CTkLabel(
            brand,
            text=(
                "Painel centralizado para autenticação, "
                "execução e acompanhamento do robô de "
                "automação SAP com exportação em nuvem."
            ),
            wraplength=wrap, justify="left",
            font=(FONT, 12), text_color="#CBD5E1"
        ).pack(anchor="w", pady=(14, 0))

        #  Environment card 
        env = ctk.CTkFrame(
            sb, corner_radius=16,
            fg_color=SIDEBAR_CARD_BG,
            border_width=1, border_color=SIDEBAR_CARD_BORDER
        )
        env.pack(fill="x", padx=24, pady=(8, 12))

        ctk.CTkLabel(
            env, text="AMBIENTE",
            font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", padx=14, pady=(14, 8))

        self.env_badge = ctk.CTkLabel(
            env, text="Verificando…",
            width=140, height=26, corner_radius=999,
            fg_color=("#1E293B", "#1E293B"),
            text_color="#E2E8F0", font=(FONT, 11, "bold")
        )
        self.env_badge.pack(anchor="w", padx=14)

        self.env_path = ctk.CTkLabel(
            env, text="Verificando OneDrive…",
            wraplength=wrap - 28, justify="left",
            font=(FONT, 11), text_color="#CBD5E1"
        )
        self.env_path.pack(anchor="w", padx=14, pady=(10, 14))

        #  Quick actions 
        ctk.CTkLabel(
            sb, text="AÇÕES RÁPIDAS",
            font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", padx=24, pady=(4, 8))

        ctk.CTkButton(
            sb, text="🌐  Abrir SharePoint",
            height=36, corner_radius=12,
            fg_color=ACCENT, hover_color=ACCENT_HOVER,
            font=(FONT, 12, "bold"),
            command=lambda: webbrowser.open_new_tab(SHAREPOINT_LINK)
        ).pack(fill="x", padx=24, pady=4)

        ctk.CTkButton(
            sb, text="🔄  Revalidar ambiente",
            height=36, corner_radius=12,
            fg_color="transparent", hover_color=SIDEBAR_BTN_HOVER,
            border_width=1, border_color="#334155",
            font=(FONT, 12, "bold"),
            command=lambda: self._check_env(navigate=not self._running())
        ).pack(fill="x", padx=24, pady=4)

        #  Theme 
        ctk.CTkLabel(
            sb, text="TEMA",
            font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", padx=24, pady=(16, 8))

        seg = ctk.CTkSegmentedButton(
            sb, values=["Sistema", "Claro", "Escuro"],
            command=self._set_theme, corner_radius=12
        )
        seg.pack(fill="x", padx=24)
        seg.set(self._prefs.get("theme", "Escuro"))

        #  Footer 
        ctk.CTkLabel(
            sb, text=f"v{APP_VERSION}  •  VINICIUS LIMA",
            font=(FONT, 10), text_color="#475569"
        ).pack(side="bottom", anchor="w", padx=24, pady=(8, 18))

    def _set_theme(self, sel):
        ctk.set_appearance_mode(theme_to_mode(sel))
        self._prefs["theme"] = sel
        save_prefs(self._prefs)

    
    #  ENVIRONMENT
    def _check_env(self, navigate: bool = True):
        self.onedrive_base = get_onedrive_path()
        self.caminho_onedrive = None
        self.env_ready = False

        if self.onedrive_base:
            p = os.path.join(self.onedrive_base, EXPECTED_FOLDER)
            if os.path.exists(p):
                self.caminho_onedrive = p
                self.env_ready = True

        self._update_env()
        if navigate:
            (self._show_login if self.env_ready else self._show_setup)()

    def _update_env(self):
        if self.env_ready:
            self.env_badge.configure(
                text="✓  Pronto", fg_color=ACCENT_SOFT, text_color=ACCENT_TEXT
            )
            self.env_path.configure(
                text=f"Pasta validada:\n{short_path(self.caminho_onedrive, 48)}"
            )
        elif self.onedrive_base:
            self.env_badge.configure(
                text="⚠  Pendente", fg_color=WARNING_FG, text_color=WARNING_TEXT
            )
            self.env_path.configure(
                text=f"OneDrive encontrado, mas '{EXPECTED_FOLDER}' não existe."
            )
        else:
            self.env_badge.configure(
                text="✕  Sem OneDrive", fg_color=ERROR_FG, text_color=ERROR_TEXT
            )
            self.env_path.configure(text="Nenhum OneDrive encontrado.")

    
    #  VISUAL HELPERS
    def _header(self, tag: str, title: str, sub: str):
        self.page_tag.configure(text=tag.upper())
        self.page_title.configure(text=title)
        self.page_sub.configure(text=sub)

    def _clear(self):
        for w in self.content.winfo_children():
            w.destroy()
        self._reset_refs()

    def _card(self, parent, alt: bool = False, **kw):
        return ctk.CTkFrame(
            parent, corner_radius=14,
            fg_color=PANEL_ALT if alt else PANEL,
            border_width=1, border_color=BORDER, **kw
        )

    def _step(self, parent, n, title, desc):
        c = self._card(parent, alt=True)
        ctk.CTkLabel(
            c, text=n, width=26, height=26, corner_radius=999,
            fg_color=ACCENT, text_color="white", font=(FONT, 11, "bold")
        ).pack(anchor="w", padx=12, pady=(12, 8))
        ctk.CTkLabel(
            c, text=title, font=(FONT, 13, "bold"), text_color=TEXT
        ).pack(anchor="w", padx=12)
        ctk.CTkLabel(
            c, text=desc, wraplength=150, justify="left",
            font=(FONT, 10), text_color=TEXT_MUTED
        ).pack(anchor="w", padx=12, pady=(4, 12))
        return c

    def _mini_card(self, parent, title, value):
        c = self._card(parent, alt=True)
        ctk.CTkLabel(
            c, text=title.upper(),
            font=(FONT, 9, "bold"), text_color=TEXT_MUTED
        ).pack(anchor="w", padx=10, pady=(10, 4))
        ctk.CTkLabel(
            c, text=value or "-",
            font=(FONT, 12, "bold"), text_color=TEXT,
            wraplength=100, justify="left"
        ).pack(anchor="w", padx=10, pady=(0, 10))
        return c

    
    #  PAGE NAVIGATION
    def _show_setup(self):
        setup_page.show(self)

    def _show_login(self):
        login_page.show(self)

    def _show_progress(self):
        progress_page.show(self)

    
    #  LOGIN HELPERS
    def _toggle_pwd(self):
        self.password_visible = not self.password_visible
        self.input_pwd.configure(show="" if self.password_visible else "•")
        self.btn_toggle_pwd.configure(
            text="Ocultar" if self.password_visible else "Mostrar"
        )

    def _form_msg(self, msg: str, kind: str = "info"):
        if not self._ok(self.form_msg):
            return
        colors = {
            "info": TEXT_MUTED, "success": ACCENT,
            "warning": WARNING_TEXT, "error": ERROR_TEXT
        }
        self.form_msg.configure(text=msg, text_color=colors.get(kind, TEXT_MUTED))

    
    #  PROGRESS HELPERS
    def _tick(self):
        """Updates the elapsed-time timer every second."""
        if self.timer_running and self._ok(self.lbl_timer):
            s = int(time.time() - self.start_time)
            self.lbl_timer.configure(text=f"⏱ {timedelta(seconds=s)}")
            self.after(1000, self._tick)
        elif self.timer_running and not self._ok(self.lbl_timer):
            self.timer_running = False

    def _log(self, txt: str):
        if not self._ok(self.log_box):
            return
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}]  {txt}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    
    #  STATUS CALLBACK
    def _make_status_callback(self, gen: int):
        """Returns a closure bound to a specific execution generation.
        If the generation has been superseded, the callback becomes a no-op,
        preventing zombie threads from polluting the GUI."""
        def cb(text, pct):
            if self._execution_gen != gen:
                return
            try:
                self.after(0, self._apply, text, pct)
            except Exception:
                pass
        return cb

    def _apply(self, text: str, pct: float):
        try:
            pct = max(0, min(100, int(float(pct))))
        except Exception:
            pct = self.current_pct
        self.current_pct = pct

        if self._ok(self.lbl_status):
            self.lbl_status.configure(text=text)
        if self._ok(self.progress_bar):
            self.progress_bar.set(pct / 100.0)
        if self._ok(self.lbl_percent):
            self.lbl_percent.configure(text=f"{pct}%")

        if text and text != self.last_status:
            self._log(text)
            self.last_status = text

        if self.execution_finished:
            return

        low = (text or "").lower()
        if "erro" in low or "falha" in low or "exception" in low:
            self._badge("error", "Erro transient")
        elif pct >= 100:
            self._badge("success", "Ciclo Concluído")
        else:
            self._badge("info", "Em execução")

    def _badge(self, kind: str, text: str):
        if not self._ok(self.status_badge):
            return
        s = {
            "info":    (ACCENT_SOFT, ACCENT_TEXT),
            "success": (SUCCESS_FG,  SUCCESS_TEXT),
            "warning": (WARNING_FG,  WARNING_TEXT),
            "error":   (ERROR_FG,    ERROR_TEXT),
        }.get(kind, (ACCENT_SOFT, ACCENT_TEXT))
        self.status_badge.configure(text=text, fg_color=s[0], text_color=s[1])

    def _finish(self, kind: str):
        if self.execution_finished:
            return
        self.execution_finished = True
        self.timer_running = False

        badge_map = {
            "success": ("success", "Concluído"),
            "warning": ("warning", "Interrompido"),
            "error":   ("error",   "Erro"),
        }
        bk, bt = badge_map.get(kind, ("info", "Finalizado"))
        self._badge(bk, bt)

        if self._ok(self.btn_stop):
            self.btn_stop.configure(state="disabled")
        if self._ok(self.btn_back):
            labels = {
                "success": "▶  Nova execução",
                "warning": "↩  Voltar",
                "error":   "🔧  Tentar novamente"
            }
            self.btn_back.configure(
                state="normal", text=labels.get(kind, "Nova execução")
            )
        if self._ok(self.lbl_hint):
            hints = {
                "success": "Concluído com sucesso!",
                "warning": "Interrompido pelo usuário.",
                "error":   "Erro detectado - revise o log."
            }
            self.lbl_hint.configure(text=hints.get(kind, ""))

    
    #  ACTIONS
    def _start(self):
        if not self.env_ready:
            self._show_setup()
            return

        user = (self.input_user.get().strip() if self.input_user else "")
        pwd = (self.input_pwd.get() if self.input_pwd else "")
        planta = (self.combo_planta.get().strip() if self.combo_planta else "")

        if not user or not pwd:
            self._form_msg("Preencha usuário e senha.", "warning")
            return
        if len(user) != 7:
            self._form_msg("Usuário SAP: exatamente 7 caracteres.", "warning")
            return
        if len(pwd) < 12:
            self._form_msg("Senha: mínimo 12 caracteres.", "warning")
            return
        if not planta or planta == "Nenhuma":
            self._form_msg("Selecione uma planta válida.", "warning")
            return

        self._prefs["last_plant"] = planta
        save_prefs(self._prefs)

        stored = keyring.get_password("RPA_SESE_USER", "default")
        if self.remember_var.get():
            keyring.set_password("RPA_SESE_USER", "default", user)
            keyring.set_password("RPA_SESE_PWD", user, pwd)
        else:
            if stored:
                safe_del_pwd("RPA_SESE_PWD", stored)
            safe_del_pwd("RPA_SESE_USER", "default")
            safe_del_pwd("RPA_SESE_PWD", user)

        settings.dynamic_user = user
        settings.dynamic_pwd = pwd
        settings.export_base_path = self.caminho_onedrive

        self.current_user = user
        self.current_plant = planta
        self.current_pct = 0
        self.last_status = None

        self.stop_event.set()

        if self.worker_thread is not None and self.worker_thread.is_alive():
            self.worker_thread.join(timeout=5)
            if self.worker_thread.is_alive():
                pass

        self.stop_event = threading.Event()
        self._execution_gen += 1
        current_gen = self._execution_gen
        self._show_progress()
        self._log(f"Execução iniciada - planta '{self.current_plant}'.")

        # WAKE UP ETL
        try:
            etl_launcher_path = os.path.join(self.caminho_onedrive, "002 - Filiais database", "006 - ETL", "ETL_Sese.exe")
            local_appdata = os.environ.get("LOCALAPPDATA", "")
            etl_local_path = os.path.join(local_appdata, "HubSeseRPA", "ETL", "bin", "ETL_Sese.exe")

            etl_exe_path = etl_launcher_path if os.path.exists(etl_launcher_path) else etl_local_path

            if os.path.exists(etl_exe_path):
                os.system("taskkill /F /IM HubSese_ETL*.exe /T 2>nul")
                clean_plant = re.sub(r'[\d\-]', '', self.current_plant).split()[0].strip().lower()
                subprocess.Popen([etl_exe_path, clean_plant], creationflags=0x08000000)
                self._log(f"Processo ETL ativado em background com planta: {clean_plant}.")
            else:
                self._log("Aviso: Executável do ETL não foi encontrado.")
                self._log(f"  Tentou: {etl_launcher_path}")
                self._log(f"  Tentou: {etl_local_path}")
        except Exception as e:
            self._log(f"Aviso: Não foi possível acordar o ETL: {e}")


        self.start_time = time.time()
        self.timer_running = True
        self._tick()

        gui_cb = self._make_status_callback(current_gen)
        self.worker_thread = threading.Thread(
            target=self._worker, args=(planta, gui_cb), daemon=True
        )
        self.worker_thread.start()

    def _worker(self, planta: str, gui_cb):
        final_kind = "success"
        try:
            run_plant(planta, gui_cb, self.stop_event)
            if self.stop_event.is_set():
                gui_cb("Processo interrompido pelo usuário.", 0)
                final_kind = "warning"
            else:
                gui_cb("Processo finalizado completamente.", 100)
                final_kind = "success"
        except Exception as e:
            gui_cb(f"Erro fatal: {e}", self.current_pct)
            final_kind = "error"
        finally:
            self.after(0, lambda k=final_kind: self._finish(k))

    def _stop(self):
        if not self._running():
            return
        self.stop_event.set()
        if self._ok(self.lbl_status):
            self.lbl_status.configure(text="Processo interrompido pelo usuário.")
        self._log("Interrupção solicitada.")
        self._finish("warning")


if __name__ == "__main__":
    app = RpaGUI()
    app.mainloop()
