import os
import time
import subprocess
import win32com.client
import pythoncom

from config.settings import settings

class SAPConnection:
    """
    Gerencia a conexão SAP GUI, login e navegação de transações.
    """
    def __init__(self, plant_id: str):
        self.plant_id = plant_id
        self.session = None
        self._already_logged_in = False

    def connect(self):
        """
          1. Tenta anexar à sessão JÁ LOGADA (Children(0).Children(0))
          2. Só abre nova conexão + login se não existir sessão ativa
        """
        SapGuiAuto = None

        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        except Exception:
            # Tenta encontrar o saplogon em locais comuns
            sap_paths = [
                r"C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe",
                r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplogon.exe"
            ]
            
            sap_exe = next((p for p in sap_paths if os.path.exists(p)), None)
            
            if not sap_exe:
                raise RuntimeError("SAP Logon não encontrado nos caminhos padrões.")
                
            subprocess.Popen(sap_exe)
            time.sleep(5)
            try:
                SapGuiAuto = win32com.client.GetObject("SAPGUI")
            except Exception as e:
                raise RuntimeError(f"SAP GUI não está rodando e a inicialização automática falhou: {e}")

        application = SapGuiAuto.GetScriptingEngine

        try:
            connection = application.Children(0)
            self.session = connection.Children(0)
            _ = self.session.findById("wnd[0]")
            self._already_logged_in = True
            return
        except Exception:
            self._already_logged_in = False

        conexao_realizada = False
        possiveis_nomes = [
            "VW Brazil - MAIIS Produção [P04]",
            "VW Brasil - MAIIS Produção [P04]",
            "VW Brazil - MAIIS Produção    [P04] Link",
            "VW Brasil - MAIIS Produção    [P04] Link",
            "VW Brazil - MAIIS Produção [P04] Link",
            "VW Brasil - MAIIS Produção [P04] Link"
        ]  # ********** depois pensar numa forma mais inteligente pra essa parte

        for nome in possiveis_nomes:
            try:
                connection = application.OpenConnection(nome, True)
                self.session = connection.Children(0)
                self._already_logged_in = False
                conexao_realizada = True
                break
            except Exception:
                continue

        if not conexao_realizada:
            raise RuntimeError("Falha ao abrir nova conexão SAP: Nenhum dos nomes (com ou sem Link/Brasil/Brazil) funcionou.")

    def ensure_logged_in(self):
        if not self.session:
            raise RuntimeError("Sessão SAP não estabelecida. Chame connect() primeiro.")

        if self._already_logged_in:
            return

        try:
            user_field = self.session.findById("wnd[0]/usr/txtRSYST-BNAME")
            password_field = self.session.findById("wnd[0]/usr/pwdRSYST-BCODE")
            
            user, password = settings.get_credentials(self.plant_id)
            
            user_field.Text = user
            password_field.Text = password
            password_field.SetFocus()
            
            self.session.findById("wnd[0]").sendVKey(0)
            
            time.sleep(1)
            
        except pythoncom.com_error:
            pass
        except Exception as e:
            raise RuntimeError(f"Ocorreu um erro inesperado durante o login no SAP: {e}")

    def start_transaction(self, t_code: str):
        if not self.session:
            raise RuntimeError("Sessão SAP não estabelecida. Chame connect() primeiro.")

        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/o"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").maximize()
            self.session.findById("wnd[0]/tbar[0]/okcd").Text = f"/n{t_code}"
            self.session.findById("wnd[0]").sendVKey(0)
            
        except pythoncom.com_error as e:
            raise RuntimeError(f"Falha ao iniciar transação {t_code}. Erro COM: {e}")
        except Exception as e:
            raise RuntimeError(f"Erro inesperado ao iniciar transação {t_code}: {e}")

    def check_connection(self):
        if not self.session:
            return False
        
        try:
            _ = self.session.findById("wnd[0]")
            return True
        except Exception:
            self.session = None
            return False
