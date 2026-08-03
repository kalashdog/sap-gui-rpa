"""
Setup Page - Guides the user through OneDrive/SharePoint environment setup.
"""
import webbrowser
import customtkinter as ctk

from gui.constants import (
    FONT, MAIN_W, EXPECTED_FOLDER, SHAREPOINT_LINK,
    ACCENT, ACCENT_HOVER, SECONDARY, SECONDARY_HOVER,
    WARNING_BORDER, WARNING_TEXT, TEXT, TEXT_MUTED,
)


def show(app) -> None:
    """Renders the environment setup page inside the app's content frame."""
    app._clear()
    app._header(
        "Configuração",
        "Prepare o ambiente",
        "Vincule o SharePoint ao OneDrive antes de usar o robô."
    )

    # Alert card
    alert = ctk.CTkFrame(
        app.content, corner_radius=14,
        fg_color=("#FFFBEB", "#1C1917"),
        border_width=1, border_color=WARNING_BORDER
    )
    alert.pack(fill="x", pady=(0, 12))

    ctk.CTkLabel(
        alert, text="⚠  Ambiente não configurado",
        font=(FONT, 13, "bold"), text_color=WARNING_TEXT
    ).pack(anchor="w", padx=14, pady=(12, 4))

    ctk.CTkLabel(
        alert,
        text=f"A pasta '{EXPECTED_FOLDER}' precisa existir no OneDrive.",
        wraplength=MAIN_W - 80, justify="left",
        font=(FONT, 11), text_color=WARNING_TEXT
    ).pack(anchor="w", padx=14, pady=(0, 12))

    # Steps 2×2 grid
    grid = ctk.CTkFrame(app.content, fg_color="transparent")
    grid.pack(fill="x", pady=(0, 12))
    grid.grid_columnconfigure(0, weight=1)
    grid.grid_columnconfigure(1, weight=1)

    app._step(grid, "1", "Abrir SharePoint",
              "Abra a biblioteca no navegador."
              ).grid(row=0, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
    app._step(grid, "2", "Adicionar ao OneDrive",
              "Clique 'Adicionar atalho'."
              ).grid(row=0, column=1, sticky="nsew", padx=(4, 0), pady=(0, 4))
    app._step(grid, "3", "Validar nome",
              f"Deve ser '{EXPECTED_FOLDER}'."
              ).grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(4, 0))
    app._step(grid, "4", "Verificar",
              "Volte e clique 'Já vinculei'."
              ).grid(row=1, column=1, sticky="nsew", padx=(4, 0), pady=(4, 0))

    # Action buttons
    ctk.CTkButton(
        app.content, text="🌐   Abrir SharePoint",
        height=40, corner_radius=12,
        fg_color=ACCENT, hover_color=ACCENT_HOVER,
        font=(FONT, 13, "bold"),
        command=lambda: webbrowser.open_new_tab(SHAREPOINT_LINK)
    ).pack(fill="x", pady=(8, 8))

    ctk.CTkButton(
        app.content, text="✓   Já vinculei, validar",
        height=40, corner_radius=12,
        fg_color=SECONDARY, hover_color=SECONDARY_HOVER,
        text_color=TEXT, font=(FONT, 13, "bold"),
        command=lambda: app._check_env(navigate=True)
    ).pack(fill="x")
