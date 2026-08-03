"""
Progress Page - Real-time execution monitoring with log, timer, and progress bar.
"""
import customtkinter as ctk

from gui.constants import (
    FONT, FONT_MONO, MAIN_W,
    ACCENT, ACCENT_SOFT, ACCENT_TEXT,
    SECONDARY, SECONDARY_HOVER, BORDER,
    TEXT, TEXT_MUTED, DANGER, DANGER_HOVER,
)


def show(app) -> None:
    """Renders the execution progress page inside the app's content frame."""
    app._clear()
    app.execution_finished = False

    app._header(
        "Execução",
        "Em andamento",
        "Acompanhe o progresso em tempo real."
    )

    # Mini-cards summary row
    row = ctk.CTkFrame(app.content, fg_color="transparent")
    row.pack(fill="x", pady=(0, 10))
    for c in range(3):
        row.grid_columnconfigure(c, weight=1)

    app._mini_card(row, "Planta", app.current_plant).grid(
        row=0, column=0, sticky="nsew", padx=(0, 3))
    app._mini_card(row, "Usuário", app.current_user).grid(
        row=0, column=1, sticky="nsew", padx=3)
    app._mini_card(row, "Export", "OneDrive").grid(
        row=0, column=2, sticky="nsew", padx=(3, 0))

    # Status card
    sc = app._card(app.content, alt=True)
    sc.pack(fill="x", pady=(0, 10))

    top = ctk.CTkFrame(sc, fg_color="transparent")
    top.pack(fill="x", padx=14, pady=(12, 6))

    app.status_badge = ctk.CTkLabel(
        top, text="Em execução", width=110, height=24,
        corner_radius=999, fg_color=ACCENT_SOFT,
        text_color=ACCENT_TEXT, font=(FONT, 10, "bold")
    )
    app.status_badge.pack(side="left")

    app.lbl_timer = ctk.CTkLabel(
        top, text="⏱ 00:00:00",
        font=(FONT, 11), text_color=TEXT_MUTED
    )
    app.lbl_timer.pack(side="left", padx=(10, 0))

    app.lbl_percent = ctk.CTkLabel(
        top, text="0%",
        font=(FONT, 18, "bold"), text_color=TEXT
    )
    app.lbl_percent.pack(side="right")

    app.lbl_status = ctk.CTkLabel(
        sc, text="Inicializando SAP…",
        font=(FONT, 12), text_color=TEXT,
        wraplength=MAIN_W - 80, justify="left"
    )
    app.lbl_status.pack(anchor="w", padx=14, pady=(0, 8))

    app.progress_bar = ctk.CTkProgressBar(
        sc, height=6, corner_radius=3, progress_color=ACCENT
    )
    app.progress_bar.set(0)
    app.progress_bar.pack(fill="x", padx=14, pady=(0, 12))

    #  FOOTER (Buttons) - packed BEFORE log to guarantee space 
    footer = ctk.CTkFrame(app.content, fg_color="transparent")
    footer.pack(side="bottom", fill="x", pady=(10, 0))

    app.lbl_hint = ctk.CTkLabel(
        footer, text="Processando passos da planta…",
        font=(FONT, 10), text_color=TEXT_MUTED
    )
    app.lbl_hint.pack(side="bottom", anchor="w", pady=(6, 0))

    app.btn_back = ctk.CTkButton(
        footer, text="Nova execução",
        height=36, corner_radius=12,
        fg_color=SECONDARY, hover_color=SECONDARY_HOVER,
        text_color=TEXT, state="disabled",
        font=(FONT, 12, "bold"), command=app._show_login
    )
    app.btn_back.pack(side="bottom", fill="x")

    app.btn_stop = ctk.CTkButton(
        footer, text="⏹   Parar",
        height=38, corner_radius=12,
        fg_color=DANGER, hover_color=DANGER_HOVER,
        font=(FONT, 13, "bold"), command=app._stop
    )
    app.btn_stop.pack(side="bottom", fill="x", pady=(0, 6))

    # LOG (Expands into remaining middle space) 
    lc = app._card(app.content, alt=True)
    lc.pack(side="top", fill="both", expand=True, pady=(0, 0))

    ctk.CTkLabel(
        lc, text="LOG", font=(FONT, 10, "bold"), text_color=TEXT_MUTED
    ).pack(anchor="w", padx=14, pady=(12, 6))

    app.log_box = ctk.CTkTextbox(
        lc, corner_radius=8, font=(FONT_MONO, 10),
        fg_color=("gray96", "#0d1117"),
        border_width=1, border_color=BORDER
    )
    app.log_box.pack(fill="both", expand=True, padx=14, pady=(0, 14))
    app.log_box.configure(state="disabled")
