import os
import sys
import customtkinter as ctk
from dotenv import load_dotenv

load_dotenv()
ctk.set_default_color_theme("blue")

APP_TITLE = "Hub Sesé • RPA Logística"
EXPECTED_FOLDER = "SESÉ DASHBOARD"
PREFS_FILE = os.path.join(os.path.expanduser("~"), ".hub_sese_rpa_ui.json")
SHAREPOINT_LINK = os.getenv("SHAREPOINT_URL", "")

def get_app_version() -> str:
    try:
        app_data = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        version_file = os.path.join(app_data, "HubSeseRPA", "current_version.txt")
        if os.path.exists(version_file):
            with open(version_file, "r", encoding="utf-8") as f:
                return f.read().strip()
    except Exception:
        pass
    return "Dev Build"

def get_asset_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller."""
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base_path, relative_path)

APP_VERSION = get_app_version()

# Windows API Constants for Insomnia Mode
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

#  FONTS
FONT = "Segoe UI"
FONT_MONO = "Consolas"

#  DIMENSIONS
SIDEBAR_W = 350
MAIN_W = 380
WIN_W = SIDEBAR_W + MAIN_W + 36  # 36 = padding
WIN_H = 640

#  COLOR PALETTE
ACCENT = "#0066F9"
ACCENT_HOVER = "#0052C7"
ACCENT_SOFT = ("#E8F1FF", "#0A2A5E")
ACCENT_TEXT = ("#0B4FCF", "#B9D4FF")

BG_APP = ("#EEF4FF", "#020817")
PANEL = ("#FFFFFF", "#0F172A")
PANEL_ALT = ("#F8FBFF", "#111827")

SIDEBAR_BG = ("#0B1220", "#030712")
SIDEBAR_MUTED = "#94A3B8"
SIDEBAR_CARD_BG = ("#111827", "#0B1220")
SIDEBAR_CARD_BORDER = "#1F2937"
SIDEBAR_BTN_HOVER = ("#1E293B", "#0F172A")

BORDER = ("#D9E6FF", "#1F2937")
TEXT = ("#0F172A", "#F8FAFC")
TEXT_MUTED = ("#64748B", "#94A3B8")

SECONDARY = ("#EAF1FF", "#1E293B")
SECONDARY_HOVER = ("#D7E6FF", "#334155")

SUCCESS_FG = ("#DCFCE7", "#14532D")
SUCCESS_TEXT = ("#166534", "#BBF7D0")
WARNING_FG = ("#FEF3C7", "#78350F")
WARNING_TEXT = ("#92400E", "#FDE68A")
WARNING_BORDER = ("#FDE68A", "#78350F")
DANGER = "#EF4444"
DANGER_HOVER = "#DC2626"
ERROR_FG = ("#FEE2E2", "#7F1D1D")
ERROR_TEXT = ("#B91C1C", "#FCA5A5")