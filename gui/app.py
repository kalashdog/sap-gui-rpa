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
    APP_VERSION, EXPECTED_FOLDER, SHAREPOINT_LINK,
    FONT, SIDEBAR_W,
    ES_CONTINUOUS, ES_SYSTEM_REQUIRED, ES_DISPLAY_REQUIRED,
    ACCENT, ACCENT_SOFT, ACCENT_TEXT,
    PANEL, PANEL_ALT, BORDER, TEXT, TEXT_MUTED,
    SIDEBAR_BG, SIDEBAR_MUTED, SIDEBAR_CARD_BG, SIDEBAR_CARD_BORDER,
    SIDEBAR_BTN_HOVER,
    SUCCESS_FG, SUCCESS_TEXT, WARNING_FG, WARNING_TEXT,
    ERROR_FG, ERROR_TEXT,
    get_asset_path,
)
from gui.helpers import load_prefs, save_prefs, theme_to_mode, short_path, safe_del_pwd, setup_window, build_shell

from gui.pages import setup as setup_page
from gui.pages import login as login_page
from gui.pages import progress as progress_page


class RpaGUI(ctk.CTk):

    def __init__(self):
        self._prefs = load_prefs()
        ctk.set_appearance_mode(theme_to_mode(self._prefs.get("theme", "Escuro")))

        super().__init__()
        setup_window(self)
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
        build_shell(self)
        self._check_env(navigate=True)

        self._set_sleep(True)

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



    def _on_close(self):
        if self._running():
            if not messagebox.askyesno(
                "Encerrar", "O robô está em execução.\nDeseja parar e fechar?"
            ):
                return
            self.stop_event.set()
        self.timer_running = False
        self._set_sleep(False)
        self.destroy()

    def _set_sleep(self, block: bool):
        try:
            if sys.platform == "win32":
                ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | (ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED if block else 0))
        except Exception: pass

    def toggle_autostart(self):
        import win32com.client
        shortcut = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs\Startup\HubSeseRPA.lnk")
        if self.autostart_var.get():
            try:
                exe = os.path.join(self.caminho_onedrive or "", "002 - Filiais database", "007 - RPA SAP", "RPA_Sese.exe")
                exe = exe if os.path.exists(exe) else (sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__))
                s = win32com.client.Dispatch("WScript.Shell").CreateShortCut(shortcut)
                s.Targetpath, s.Arguments, s.WorkingDirectory, s.IconLocation = exe, "--autostart", os.path.dirname(exe), exe
                s.save()
            except Exception: pass
        elif os.path.exists(shortcut):
            os.remove(shortcut)
        self._prefs["autostart"] = self.autostart_var.get()
        save_prefs(self._prefs)

    def trigger_auto_start(self):
        if keyring.get_password("RPA_SESE_USER", "default"):
            self._start()

    def _set_theme(self, sel):
        ctk.set_appearance_mode(theme_to_mode(sel))
        self._prefs["theme"] = sel
        save_prefs(self._prefs)

    
    #  ENVIRONMENT
    def _check_env(self, navigate: bool = True):
        self.onedrive_base = get_onedrive_path()
        self.caminho_onedrive = os.path.join(self.onedrive_base, EXPECTED_FOLDER) if self.onedrive_base else None
        self.env_ready = bool(self.caminho_onedrive and os.path.exists(self.caminho_onedrive))
        self._update_env()
        if navigate: (self._show_login if self.env_ready else self._show_setup)()

    def _update_env(self):
        if self.env_ready:
            self.env_badge.configure(text="  ✓ Pronto  ", fg_color=ACCENT_SOFT, text_color=ACCENT_TEXT)
            self.env_path.pack_forget()
            self.jobs_container.pack(side="top", fill="both", expand=True, padx=12, pady=(0, 8))
        else:
            self.jobs_container.pack_forget()
            self.env_path.pack(anchor="w", padx=14, pady=(0, 14))
            if self.onedrive_base:
                self.env_badge.configure(text="  ⚠ Pendente  ", fg_color=WARNING_FG, text_color=WARNING_TEXT)
                self.env_path.configure(text=f"OneDrive encontrado, mas '{EXPECTED_FOLDER}' não existe.")
            else:
                self.env_badge.configure(text="  ✕ Sem OneDrive  ", fg_color=ERROR_FG, text_color=ERROR_TEXT)
                self.env_path.configure(text="Nenhum OneDrive encontrado.")

    def on_plant_select(self, plant_id: str):
        self._prefs["last_plant"] = plant_id
        save_prefs(self._prefs)
        for w in self.jobs_container.winfo_children(): w.destroy()
            
        from config.settings import settings
        jobs = settings.config.get("jobs", {})
        
        plant_jobs_count = sum(1 for j in jobs.values() if plant_id in j.get("plant_params", {}))
        
        ctk.CTkLabel(
            self.jobs_container, text=f"JOBS DISPONÍVEIS ({plant_id[3:]}): {plant_jobs_count}",
            font=(FONT, 10, "bold"), text_color=SIDEBAR_MUTED
        ).pack(anchor="w", pady=(0, 8), padx=12)
        
        sorted_jobs = sorted(jobs.items(), key=lambda item: item[1].get("dashboard", ""))
        for job_key, job_data in sorted_jobs:
            if plant_id in job_data.get("plant_params", {}):
                is_active = job_data.get("active", False)
                bg_color, border_color = (SIDEBAR_CARD_BG, SIDEBAR_CARD_BORDER) if is_active else (("#FEF2F2", "#2A1010"), ("#FECACA", "#451A1A"))
                
                card = ctk.CTkFrame(self.jobs_container, corner_radius=8, fg_color=bg_color, border_width=1, border_color=border_color)
                card.pack(fill="x", padx=12, pady=4)

                hdr = ctk.CTkFrame(card, fg_color="transparent")
                hdr.pack(fill="x", padx=10, pady=(8, 2))
                ctk.CTkLabel(hdr, text=f"▶ {job_key}", font=(FONT, 11, "bold"), text_color="#E2E8F0").pack(side="left")
                
                s_txt, s_col = ("● ON", SUCCESS_TEXT) if is_active else ("○ OFF", ERROR_TEXT)
                ctk.CTkLabel(hdr, text=s_txt, font=(FONT, 10, "bold"), text_color=s_col).pack(side="right")

                info_row = ctk.CTkFrame(card, fg_color="transparent")
                info_row.pack(fill="x", padx=10, pady=(0, 8))
                
                ctk.CTkLabel(info_row, text=f"Transação: {job_data.get('transaction', 'N/A')}", font=(FONT, 10), text_color=SIDEBAR_MUTED).pack(side="left")
                ctk.CTkLabel(info_row, text=job_data.get('dashboard', 'N/A')[6:], font=(FONT, 10), text_color=SIDEBAR_MUTED).pack(side="right")
                
        if not plant_jobs_count:
            ctk.CTkLabel(
                self.jobs_container, text="Nenhum job configurado para esta planta.",
                font=(FONT, 11), text_color=SIDEBAR_MUTED, wraplength=SIDEBAR_W - 50, justify="left"
            ).pack(anchor="w", padx=12, pady=10)

    
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
        if not self.timer_running: return
        if self._ok(self.lbl_timer):
            self.lbl_timer.configure(text=f"⏱ {timedelta(seconds=int(time.time() - self.start_time))}")
            self.after(1000, self._tick)
        else: self.timer_running = False

    def _log(self, txt: str):
        if self._ok(self.log_box):
            self.log_box.configure(state="normal")
            self.log_box.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}]  {txt}\n")
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
        try: pct = max(0, min(100, int(float(pct))))
        except Exception: pct = self.current_pct
        self.current_pct = pct

        if self._ok(self.lbl_status): self.lbl_status.configure(text=text)
        if self._ok(self.progress_bar): self.progress_bar.set(pct / 100.0)
        if self._ok(self.lbl_percent): self.lbl_percent.configure(text=f"{pct}%")

        if text and text != self.last_status:
            self._log(text)
            self.last_status = text

        if self.execution_finished: return

        low = (text or "").lower()
        if any(x in low for x in ("erro", "falha", "exception")): self._badge("error", "Erro transient")
        elif pct >= 100: self._badge("success", "Ciclo Concluído")
        else: self._badge("info", "Em execução")

    def _badge(self, kind: str, text: str):
        if self._ok(self.status_badge):
            fg, txt = {"success": (SUCCESS_FG, SUCCESS_TEXT), "warning": (WARNING_FG, WARNING_TEXT), "error": (ERROR_FG, ERROR_TEXT)}.get(kind, (ACCENT_SOFT, ACCENT_TEXT))
            self.status_badge.configure(text=text, fg_color=fg, text_color=txt)

    def _finish(self, kind: str):
        if self.execution_finished: return
        self.execution_finished = True
        self.timer_running = False

        cfg = {
            "success": ("Concluído", "▶  Nova execução", "Concluído com sucesso!"),
            "warning": ("Interrompido", "↩  Voltar", "Interrompido pelo usuário."),
            "error":   ("Erro", "🔧  Tentar novamente", "Erro detectado - revise o log.")
        }.get(kind, ("Finalizado", "Nova execução", ""))
        
        self._badge(kind if kind in ["success", "warning", "error"] else "info", cfg[0])

        if self._ok(self.btn_stop): self.btn_stop.configure(state="disabled")
        if self._ok(self.btn_back): self.btn_back.configure(state="normal", text=cfg[1])
        if self._ok(self.lbl_hint): self.lbl_hint.configure(text=cfg[2])

    
    #  ACTIONS
    def _start(self):
        if not self.env_ready:
            self._show_setup()
            return

        user = (self.input_user.get().strip() if self.input_user else "")
        pwd = (self.input_pwd.get() if self.input_pwd else "")
        planta = (self.combo_planta.get().strip() if self.combo_planta else "")

        err = next((m for c, m in [
            (not user or not pwd, "Preencha usuário e senha."),
            (len(user) != 7, "Usuário SAP: exatamente 7 caracteres."),
            (len(pwd) < 12, "Senha: mínimo 12 caracteres."),
            (not planta or planta == "Nenhuma", "Selecione uma planta válida.")
        ] if c), None)
        if err: return self._form_msg(err, "warning")

        self._prefs["last_plant"] = planta
        save_prefs(self._prefs)

        stored = keyring.get_password("RPA_SESE_USER", "default")
        if self.remember_var.get():
            keyring.set_password("RPA_SESE_USER", "default", user)
            keyring.set_password("RPA_SESE_PWD", user, pwd)
        else:
            if stored: safe_del_pwd("RPA_SESE_PWD", stored)
            safe_del_pwd("RPA_SESE_USER", "default")
            safe_del_pwd("RPA_SESE_PWD", user)

        settings.dynamic_user, settings.dynamic_pwd, settings.export_base_path = user, pwd, self.caminho_onedrive
        self.current_user, self.current_plant, self.current_pct, self.last_status = user, planta, 0, None

        self.stop_event.set()
        if self._running(): self.worker_thread.join(timeout=5)

        self.stop_event = threading.Event()
        self._execution_gen += 1
        self._show_progress()
        self._log(f"Execução iniciada - planta '{self.current_plant}'.")

        try:
            etl_exe = os.path.join(self.caminho_onedrive, "002 - Filiais database", "006 - ETL", "ETL_Sese.exe")
            if os.path.exists(etl_exe):
                os.system("taskkill /F /IM HubSese_ETL*.exe /T 2>nul")
                clean_plant = re.sub(r'[\d\-]', '', self.current_plant).split()[0].strip().lower()
                subprocess.Popen([etl_exe, clean_plant], creationflags=0x08000000)
                self._log(f"Processo ETL ativado: {clean_plant}")
        except Exception: pass

        self.start_time, self.timer_running = time.time(), True
        self._tick()

        self.worker_thread = threading.Thread(target=self._worker, args=(planta, self._make_status_callback(self._execution_gen)), daemon=True)
        self.worker_thread.start()

    def _worker(self, planta: str, gui_cb):
        try:
            run_plant(planta, gui_cb, self.stop_event)
            is_stop = self.stop_event.is_set()
            gui_cb("Processo interrompido pelo usuário." if is_stop else "Processo finalizado completamente.", 0 if is_stop else 100)
            final_kind = "warning" if is_stop else "success"
        except Exception as e:
            gui_cb(f"Erro fatal: {e}", self.current_pct)
            final_kind = "error"
        finally:
            self.after(0, lambda k=final_kind: self._finish(k))

    def _stop(self):
        if not self._running(): return
        self.stop_event.set()
        if self._ok(self.lbl_status): self.lbl_status.configure(text="Processo interrompido pelo usuário.")
        self._log("Interrupção solicitada.")
        self._finish("warning")


if __name__ == "__main__":
    app = RpaGUI()
    app.mainloop()
