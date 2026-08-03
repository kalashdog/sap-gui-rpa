"""
Login Page - SAP credentials input and plant selection.
"""
import tkinter as tk
import customtkinter as ctk
import keyring

from config.settings import settings
from gui.constants import (
    FONT, MAIN_W, ACCENT, ACCENT_HOVER,
    SECONDARY, SECONDARY_HOVER, BORDER,
    TEXT, TEXT_MUTED, WARNING_TEXT,
)
from gui.helpers import short_path

def show(app) -> None:
    """Renders the SAP login page inside the app's content frame."""
    app._clear()
    app.password_visible = False

    app._header(
        "Autenticação SAP",
        "Iniciar automação",
        "Informe credenciais e selecione a planta."
    )

    # Export path info card
    info = app._card(app.content, alt=True)
    info.pack(fill="x", pady=(0, 12))

    row_info = ctk.CTkFrame(info, fg_color="transparent")
    row_info.pack(fill="x", padx=14, pady=10)

    ctk.CTkLabel(
        row_info, text="✓  Exportação:",
        font=(FONT, 11, "bold"), text_color=ACCENT
    ).pack(side="left")

    ctk.CTkLabel(
        row_info, text=short_path(app.caminho_onedrive, 30),
        font=(FONT, 11), text_color=TEXT
    ).pack(side="left", padx=(6, 0))

    # Form
    form = ctk.CTkFrame(app.content, fg_color="transparent")
    form.pack(fill="x", pady=(4, 0))

    # Plant selector
    plants = list(settings.config.get("plants", {}).keys()) or ["Nenhuma"]

    ctk.CTkLabel(
        form, text="Planta",
        font=(FONT, 11, "bold"), text_color=TEXT_MUTED
    ).pack(anchor="w", pady=(0, 4))

    app.combo_planta = ctk.CTkComboBox(
        form, values=plants, height=36,
        corner_radius=10, border_width=1, border_color=BORDER,
        font=(FONT, 12), dropdown_font=(FONT, 12)
    )
    app.combo_planta.pack(anchor="w", fill="x", pady=(0, 10))

    last = app._prefs.get("last_plant")
    app.combo_planta.set(last if last in plants else plants[0])

    # Username
    ctk.CTkLabel(
        form, text="Usuário SAP",
        font=(FONT, 11, "bold"), text_color=TEXT_MUTED
    ).pack(anchor="w", pady=(0, 4))

    app.input_user = ctk.CTkEntry(
        form, placeholder_text="Ex.: FV2WL5N",
        height=36, corner_radius=10,
        border_width=1, border_color=BORDER, font=(FONT, 12)
    )
    app.input_user.pack(anchor="w", fill="x", pady=(0, 10))

    # Password
    ctk.CTkLabel(
        form, text="Senha SAP",
        font=(FONT, 11, "bold"), text_color=TEXT_MUTED
    ).pack(anchor="w", pady=(0, 4))

    pwd_row = ctk.CTkFrame(form, fg_color="transparent")
    pwd_row.pack(anchor="w", fill="x", pady=(0, 8))

    app.input_pwd = ctk.CTkEntry(
        pwd_row, placeholder_text="Mínimo 12 caracteres",
        show="•", height=36, corner_radius=10,
        border_width=1, border_color=BORDER, font=(FONT, 12)
    )
    app.input_pwd.pack(side="left", fill="x", expand=True, padx=(0, 6))

    app.btn_toggle_pwd = ctk.CTkButton(
        pwd_row, text="Mostrar", width=72, height=36,
        corner_radius=10, fg_color=SECONDARY,
        hover_color=SECONDARY_HOVER, text_color=TEXT,
        font=(FONT, 11), command=app._toggle_pwd
    )
    app.btn_toggle_pwd.pack(side="right")

    # Pre-fill from keyring
    saved_u = keyring.get_password("RPA_SESE_USER", "default")
    if saved_u:
        app.input_user.insert(0, saved_u)
        saved_p = keyring.get_password("RPA_SESE_PWD", saved_u)
        if saved_p:
            app.input_pwd.insert(0, saved_p)
            app.remember_var.set(True)
        else:
            app.remember_var.set(False)
    else:
        app.remember_var.set(True)

    # Remember checkbox
    ctk.CTkCheckBox(
        form, text="Lembrar credenciais",
        variable=app.remember_var, font=(FONT, 11)
    ).pack(anchor="w", pady=(0, 6))

    # Autostart checkbox
    ctk.CTkCheckBox(
        form, text="Iniciar automaticamente com o Windows",
        variable=app.autostart_var, font=(FONT, 11),
        command=app.toggle_autostart
    ).pack(anchor="w", pady=(0, 6))

    # Feedback label
    app.form_msg = ctk.CTkLabel(
        form, text="", font=(FONT, 11), text_color=TEXT_MUTED
    )
    app.form_msg.pack(anchor="w", pady=(0, 8))

    # Start button
    app.btn_start = ctk.CTkButton(
        form, text="▶   Iniciar RPA",
        height=42, corner_radius=12,
        fg_color=ACCENT, hover_color=ACCENT_HOVER,
        font=(FONT, 14, "bold"), command=app._start
    )
    app.btn_start.pack(fill="x", pady=(0, 6))

    ctk.CTkLabel(
        form, text="Ou pressione Enter no campo de senha.",
        font=(FONT, 10), text_color=TEXT_MUTED
    ).pack(anchor="w")

    app.input_pwd.bind("<Return>", lambda _: app._start())

    if "Nenhuma" in plants[0]:
        app.btn_start.configure(state="disabled")
        app._form_msg("Nenhuma planta configurada.", "warning")

    # Focus management
    if not saved_u:
        app.input_user.focus()
    elif not app.input_pwd.get():
        app.input_pwd.focus()
