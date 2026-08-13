"""
Scripts de transação SAP.
Contém as lógicas de navegação e extração de dados no SAP GUI via COM.
"""
import os
import time
from datetime import datetime, timedelta
from typing import Tuple, Dict, Any

import pythoncom
import win32com.client

from config.settings import settings
from core.utils import get_target_export_dir


def _resolve_export_path(plant_id: str, job_key: str, default_name: str) -> Tuple[str, str, Dict[str, Any]]:
    plant_config = settings.config["plants"].get(plant_id, {})
    folder_name = plant_config.get("folder_name", plant_id)
    inner_base_path = plant_config.get("base_path", "")
    
    job_data = settings.config["jobs"][job_key]
    params = job_data["plant_params"][plant_id]
    local_extract = params.get("local_extract", "")
    name_file = params.get("name_file", default_name).format(date=datetime.now())
    
    use_staging = job_data.get("use_staging", False)
    
    if use_staging:
        plant_name = plant_config.get("name", plant_id).lower()
        local_appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
        full_path = os.path.normpath(os.path.join(
            local_appdata, "HubSeseRPA", "ETL", "staging", plant_name, local_extract
        ))
    else:
        base_plant_path = get_target_export_dir(folder_name)
        full_path = os.path.normpath(os.path.join(base_plant_path, inner_base_path, local_extract))
        
    os.makedirs(full_path, exist_ok=True)
    return full_path, name_file, params

def close_excel(filename):
    time.sleep(3)
    try:
        excel = win32com.client.GetObject(Class="Excel.Application")
        for wb in excel.Workbooks:
            if filename.lower() in wb.Name.lower():
                wb.Close(SaveChanges=False)
                break
    except Exception:
        pass

def send_to_background(session: Any, spool_name: str, printer: str = "locl") -> None:
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[2]").Select()

        try:
            session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").Text
        except pythoncom.com_error:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()

        try:
            session.findById("wnd[1]/usr/ctxtPRI_PARAMS-PDEST").Text = printer
        except pythoncom.com_error:
            pass

        session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").Text = spool_name
        session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PRTXT").Text = spool_name
        session.findById("wnd[1]/usr/subSUBSCREEN:SAPLSPRI:0600/txtPRI_PARAMS-PLIST").SetFocus()
        session.findById("wnd[1]/tbar[0]/btn[13]").press()

        try:
            session.findById("wnd[2]").sendVKey(0)
        except pythoncom.com_error:
            pass

        session.findById("wnd[1]/usr/btnSOFORT_PUSH").press()
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        session.findById("wnd[0]").sendVKey(3)
    except pythoncom.com_error as e:
        raise RuntimeError(f"Falha ao enviar job '{spool_name}' para background: {e}")

def export_xxl(session: Any, path: str, filename: str, shell_id: str = None) -> None:
    """
    Exporta para XXL no formato .tmp para evitar que o Excel abra automaticamente e depois volta pra xlsx.
    """
    try:
        nome_base, extensao = os.path.splitext(filename)
        temp_filename = f"{nome_base}.tmp"

        if shell_id:
            shell = session.findById(shell_id)
        else:
            try:
                shell = session.findById("wnd[0]/shellcont[1]/shell")
            except pythoncom.com_error:
                shell = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")

        shell.pressToolbarContextButton("&MB_EXPORT")
        shell.selectContextMenuItem("&XXL")
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        session.findById("wnd[1]/usr/ctxtDY_PATH").Text = path
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = temp_filename
        session.findById("wnd[1]/tbar[0]/btn[0]").press()

        try:
            session.findById("wnd[1]/tbar[0]/btn[11]").press()
        except pythoncom.com_error:
            pass

        caminho_temp = os.path.abspath(os.path.join(path, temp_filename))
        caminho_final = os.path.abspath(os.path.join(path, filename))

        timeout = 45 
        start = time.time()

        while not os.path.exists(caminho_temp):
            if time.time() - start > timeout:
                raise TimeoutError(f"SAP demorou muito para exportar {temp_filename}")
            time.sleep(0.5)

        time.sleep(2.0)

        if os.path.exists(caminho_final):
            os.remove(caminho_final)

        os.rename(caminho_temp, caminho_final)

    except pythoncom.com_error as e:
        raise RuntimeError(f"Export XXL falhou para {filename}: {e}")
    except Exception as e:
        raise RuntimeError(f"Erro no processamento XXL de {filename}: {e}")

def _get_params(job_key: str, plant_id: str) -> Tuple[Dict[str, Any], str]:
    cfg = settings.config["jobs"][job_key]
    plant_params = cfg["plant_params"][plant_id]
    spool = plant_params.get("spool_name") or cfg["spool_name"]
    return plant_params, spool

# LT23 (CORINGA GENÉRICO - via variante externa)
def request_lt23(session, plant_id: str, job_key: str):
    """LT23 via variante externa - a variante ja contém lgnum, datas e radio."""
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    if "lgnum" in params:
        session.findById("wnd[0]/usr/ctxtT1_LGNUM").Text = params["lgnum"]

    send_to_background(session, spool, params.get("printer", "locl"))

# LT22 (CORINGA GENÉRICO - via variante externa)
def request_lt22(session, plant_id: str, job_key: str):
    """LT22 via variante externa - a variante ja contém lgnum, datas e radio."""
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    
    if "lgnum" in params:
        session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]

    send_to_background(session, spool, params.get("printer", "locl"))

# LT23
def request_atend_linha_pendentes(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/radT1_OFFTA").Select()
    session.findById("wnd[0]/usr/ctxtT1_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params["variant"]
    session.findById("wnd[0]/usr/ctxtLISTV").SetFocus()
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool, params.get("printer", "locl"))

# LT23
def request_lt23_fifo1(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtT1_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/txtT1_TANUM-LOW").Text = ""
    session.findById("wnd[0]/usr/txtT1_TANUM-HIGH").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtAENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtMLANGU-LOW").Text = ""
    session.findById("wnd[1]").sendVKey(0)

    session.findById("wnd[0]/usr/radT1_ALLTA").Select()
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)

# LT23
def request_lt23_cofre1(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtT1_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/txtT1_TANUM-LOW").Text = ""
    session.findById("wnd[0]/usr/txtT1_TANUM-HIGH").Text = ""
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtAENAME-LOW").Text = ""
    session.findById("wnd[1]/usr/txtMLANGU-LOW").Text = ""
    session.findById("wnd[1]").sendVKey(0)

    session.findById("wnd[0]/usr/radT1_ALLTA").Select()
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  MB51_EMPURRADA 
def request_mb51_empurrada(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=4)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = params["werks"]
    session.findById("wnd[0]/usr/ctxtLGORT-LOW").Text = params["lgort"]
    session.findById("wnd[0]/usr/ctxtBWART-LOW").Text = "311"
    session.findById("wnd[0]/usr/ctxtBUDAT-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtBUDAT-HIGH").SetFocus()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/radRFLAT_L").Select()
    session.findById("wnd[0]/usr/ctxtALV_DEF").Text = params["variant"]
    session.findById("wnd[0]/usr/ctxtALV_DEF").SetFocus()
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool, params.get("printer", "locl"))


# MB51 (CORINGA GENÉRICO - com variante externa)
def request_mb51(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]").sendVKey(0)

    if "werks" in params:
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = params["werks"]

    send_to_background(session, spool)


#  LX03 (FIFO2, ZONA, BLOQUEADOS) 
def request_lx03(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = params["lgnum"]
    if "lgtyp" in params:
        session.findById("wnd[0]/usr/ctxtS1_LGTYP-LOW").Text = params["lgtyp"]
    session.findById("wnd[0]/usr/chkPMITB").Selected = True
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LX02 (IMP, BESI3) 
def request_lx02(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtS1_LGNUM").Text = params["lgnum"]
    
    if "werks" in params:
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = params["werks"]
        
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = params["variant"]
    session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus()

    send_to_background(session, spool)


#  LT22_IMP2 
def request_lt22_imp2(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = params["lgtyp"]
    session.findById("wnd[0]/usr/radT3_OFFTA").Select()
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = ""
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LT22_IMP3 
def request_lt22_imp3(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = datetime.now().replace(day=1).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = params["lgtyp"]
    session.findById("wnd[0]/usr/radT3_ALLTA").Select()
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params["variant"]
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LT22_ALERTAOP (alerta no teams)
def request_lt22_alertaop(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    data_ini = (datetime.now() - timedelta(days=1)).strftime("%d.%m.%Y")
    data_fim = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = params["lgtyp"]
    
    if params.get("radio") == "ALLTA":
        session.findById("wnd[0]/usr/radT3_ALLTA").Select()
    
    session.findById("wnd[0]/usr/ctxtBDATU-LOW").Text = data_ini
    session.findById("wnd[0]/usr/ctxtBDATU-HIGH").Text = data_fim
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params.get("variant", "")
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)


#  LT22_ZONA 
def request_lt22_zona(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtT3_LGNUM").Text = params["lgnum"]
    session.findById("wnd[0]/usr/ctxtT3_LGTYP-LOW").Text = params["lgtyp"]
    
    if params.get("radio") == "ALLTA":
        session.findById("wnd[0]/usr/radT3_ALLTA").Select()
    else:
        session.findById("wnd[0]/usr/radT3_OFFTA").SetFocus()
        
    session.findById("wnd[0]/usr/ctxtLISTV").Text = params.get("variant", "")
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool)

request_lt22_zona_geral = request_lt22_zona


#  VL06I_FORNEC (e VL06I_FORNEC2 via lgnum) 
def request_vl06i(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)
    today = datetime.now().strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/btnBUTTON7").press()
    
    lgnum = params.get("lgnum")
    vstel = params.get("vstel", "")

    if lgnum:
        # CTB parte 2: filtra por número de depósito (lgnum) em vez de ponto de expedição
        session.findById("wnd[0]/usr/ctxtIF_VSTEL-LOW").Text = ""
        session.findById("wnd[0]/usr/ctxtIT_LGNUM-LOW").Text = lgnum
    elif isinstance(vstel, list):
        session.findById("wnd[0]/usr/btn%_IF_VSTEL_%_APP_%-VALU_PUSH").press()
        for idx, val in enumerate(vstel):
            session.findById(f"wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,{idx}]").Text = val
        session.findById("wnd[1]").sendVKey(8)
    else:
        session.findById("wnd[0]/usr/ctxtIF_VSTEL-LOW").Text = vstel
        
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-LOW").Text = params["date_low"]
    session.findById("wnd[0]/usr/ctxtIT_LFDAT-HIGH").Text = today
    session.findById("wnd[0]/usr/ctxtIT_WBSTK-LOW").SetFocus()
    session.findById("wnd[0]").sendVKey(2)

    try:
        session.findById("wnd[1]/usr/cntlMY_TOOLBAR_CONTAINER/shellcont/shell").pressButton("EXCL")
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").currentCellColumn = "TEXT"
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "0"
        session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").doubleClickCurrentCell()
    except pythoncom.com_error:
        pass

    session.findById("wnd[0]/usr/ctxtIT_WBSTK-LOW").Text = "c"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkIF_ITEM").Selected = True

    send_to_background(session, spool)


#  MAISEWM021R_EMBALAGEM 
def request_maisewm021r_embalagem(session: Any, plant_id: str, job_key: str) -> None:
    params, spool = _get_params(job_key, plant_id)
    today = datetime.now()
    first_day = today.replace(day=1).strftime("%d.%m.%Y")

    session.findById("wnd[0]/usr/btnREL").press()
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = params["werks"]
    session.findById("wnd[0]/usr/ctxtS_DATUM-LOW").Text = first_day
    session.findById("wnd[0]/usr/ctxtS_DATUM-HIGH").Text = today.strftime("%d.%m.%Y")
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/radP_MOV").Select()

    send_to_background(session, spool)


#  MB52_AUTO 
def request_mb52(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/txtV-LOW").Text = params["variant"]
    session.findById("wnd[1]/tbar[0]/btn[8]").press()

    send_to_background(session, spool)


#  AL11_BESI3 (Foreground) 
def request_al11_besi3(session, plant_id: str, job_key: str):

    full_path, name_file, _ = _resolve_export_path(plant_id, job_key, "besi3.txt")

    folders = ["\\\\10.135.7.23\\files\\PRD\\interfaces", "pp", "inbound", "BESI3", "5100", "Backup"]
    grid_id = "wnd[0]/usr/cntlGRID1/shellcont/shell"

    for folder in folders:
        i = 0
        col = "DIRNAME" if "\\\\" in folder else "NAME"
        while True:
            try:
                val = session.findById(grid_id).GetCellValue(i, col)
                if val == folder:
                    session.findById(grid_id).setCurrentCell(i, col)
                    session.findById(grid_id).doubleClickCurrentCell()
                    break
                i += 1
            except pythoncom.com_error:
                break

    session.findById(grid_id).setCurrentCell(-1, "MOD_DATE")
    session.findById(grid_id).selectColumn("MOD_DATE")
    session.findById("wnd[0]/tbar[1]/btn[40]").press()

    session.findById(grid_id).currentCellRow = -1
    session.findById(grid_id).selectColumn("USEABLE")
    session.findById("wnd[0]/tbar[1]/btn[29]").press()
    session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").Text = "x"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById(grid_id).selectedRows = "0"
    session.findById(grid_id).doubleClickCurrentCell()

    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = full_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = name_file
    session.findById("wnd[1]/tbar[0]/btn[11]").press()
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)


#  PKMC_GERAL (Foreground) 
def request_pkmc(session, plant_id: str, job_key: str):

    full_path, name_file, params = _resolve_export_path(plant_id, job_key, "PKMC.XLSX")

    session.findById("wnd[0]/usr/ssubCCY_AND_SELECTION:SAPLMPK_CCY_UI:0111/subSELECTION:SAPLMPK_CCY_UI:0113/ctxtRMPKR-WERKS").Text = params["werks"]
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ssubCCY_AND_SELECTION:SAPLMPK_CCY_UI:0111/subSELECTION:SAPLMPK_CCY_UI:0113/btnESEL").press()
    session.findById("wnd[1]/usr/ctxtRANG_MAT-LOW").Text = ""
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    time.sleep(2)

    grid_pkmc = "wnd[0]/usr/ssubCCY_AND_SELECTION:SAPLMPK_CCY_UI:0111/subCCY:SAPLMPK_CCY_UI:0130/subBIGGRIDCONTAINER:SAPLMPK_CCY_UI:0135/cntlAVAILABLE_CONTROLCYCLES/shellcont/shell"
    session.findById(grid_pkmc).pressToolbarContextButton("&MB_VARIANT")
    session.findById(grid_pkmc).selectContextMenuItem("&LOAD")

    variant_grid = "wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell"
    i = 0
    while True:
        try:
            val = session.findById(variant_grid).GetCellValue(i, "VARIANT")
            if val == params["variant"]:
                session.findById(variant_grid).selectedRows = str(i)
                session.findById(variant_grid).clickCurrentCell()
                break
            i += 1
        except pythoncom.com_error:
            break

    export_xxl(session, full_path, name_file, shell_id=grid_pkmc)
    session.findById("wnd[0]").sendVKey(3)


#  MD04_GLOBAL (Foreground) 
def request_md04_global(session, plant_id: str, job_key: str):
    full_path, name_file, params = _resolve_export_path(plant_id, job_key, "MD04_full.XLSX")

    session.findById("wnd[0]/usr/tabsTAB300/tabpF02").Select()

    tab = "wnd[0]/usr/tabsTAB300/tabpF02/ssubINCLUDE300:SAPMM61R:0212"

    session.findById(f"{tab}/ctxtRM61R-WERKS2").Text = params["werks"]
    session.findById("wnd[0]").sendVKey(0)
    session.findById(f"{tab}/radRM61R-CLSKZ").Select()
    session.findById(f"{tab}/ctxtRM61R-CLASS").Text = params["class"]
    session.findById(f"{tab}/ctxtRM61R-KLART").Text = params["klart"]
    session.findById(f"{tab}/ctxtRM61R-KLART").SetFocus()
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").sendVKey(0)

    tbl = "wnd[0]/usr/subVALUATION_DYNPRO:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S"
    for i in range(5):
        try:
            session.findById(f"{tbl}/ctxtRCTMS-MWERT[1,{i}]").Text = "*"
        except pythoncom.com_error:
            pass

    session.findById("wnd[0]/mbar/menu[4]/menu[0]").Select()
    session.findById("wnd[1]/usr/tabsPARAM/tabpSUCH").Select()
    session.findById("wnd[1]/usr/tabsPARAM/tabpSUCH/ssubSUB:SAPLCLPR:0110/txtRMCLPAR-DAR_MAX_HITS").Text = "0"
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()

    export_xxl(session, full_path, name_file)

    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

#  SSC0_USERS_SYNC (Foreground / Exce) 
def request_users_sync(session, plant_id: str, job_key: str):
    """
    o ETL exporta uma lista de usuarios SAP novos (que não estão na planilha com nome e sobrenome) em um arquivo txt.
    e então esse script detecta se há um user novo, se houver ele procura o nome no SAP e insere na planilha excel, 
    depois exclui o user do arquivo para não processar novamente
    """
    import win32com.client
    
    local_appdata = os.environ.get('LOCALAPPDATA', os.path.expanduser('~'))
    staging_file = os.path.join(local_appdata, "HubSeseRPA", "ETL", "staging", "global", "sapusers.txt")
    
    if not os.path.exists(staging_file):
        return 
        
    with open(staging_file, "r", encoding="utf-8") as f:
        usuarios = f.read().splitlines()
        
    if not usuarios:
        return 
        
    resultados = []
    for user_id in usuarios:
        if not user_id.strip():
            continue
            
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSSC0"
        session.findById("wnd[0]").sendVKey(0)
        
        try:
            session.findById("wnd[1]/usr/radSOS04-S_USR_SEL").Select()
            session.findById("wnd[1]/usr/txtSOS04-S_ADR_NAME").Text = user_id
            session.findById("wnd[1]/tbar[0]/btn[0]").Press()
            
            titulo = session.findById("wnd[0]/titl").Text
            if ":" in titulo:
                nome = titulo.split(":", 1)[1].strip()
                resultados.append((user_id, nome))
            else:
                resultados.append((user_id, "DESLIGADO"))
        except Exception:
            resultados.append((user_id, "DESLIGADO"))
            
    if not resultados:
        open(staging_file, 'w').close()
        return
        
    full_path, name_file, params = _resolve_export_path(plant_id, job_key, "scan user.xlsm")
    excel_path = os.path.join(full_path, name_file)
    
    plant_map = {
        "01-Anchieta": "ANCHIETA",
        "02-Taubate": "TAUBATE",
        "03-Curitiba": "CURITIBA",
        "04-SaoCarlos": "SÃO CARLOS"
    }
    planta_formatada = plant_map.get(plant_id, "DESCONHECIDO")
    
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb = excel.Workbooks.Open(excel_path)
        ws = wb.Worksheets(1)
        
        last_row = ws.Cells(ws.Rows.Count, "B").End(-4162).Row
        
        for user_id, nome in resultados:
            last_row += 1
            ws.Cells(last_row, 2).Value = user_id
            ws.Cells(last_row, 3).Value = nome
            ws.Cells(last_row, 4).Value = "NOVO"
            ws.Cells(last_row, 7).Value = "DEFINIR LOCAL"
            ws.Cells(last_row, 8).Value = planta_formatada
            ws.Cells(last_row, 9).Value = "A DEFINIR"
            
        wb.Save()
    except Exception as e:
        print(f"Erro ao salvar excel: {e}")
    finally:
        try:
            wb.Close(SaveChanges=False)
        except: pass
        excel.Quit()

    open(staging_file, 'w').close()


#  PK05_GERAL (Foreground) 
def request_pk05(session, plant_id: str, job_key: str):
    full_path, name_file, params = _resolve_export_path(plant_id, job_key, "PK05.txt")

    # Fechar possível aviso de permissão de edição (usuários sem acesso de edição)
    try:
        session.findById("wnd[1]/usr/sub:SAPLSVIX:0100/ctxtD0100_FIELD_TAB-LOWER_LIMIT[0,37]").Text = params["werks"]
    except Exception:
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press
        except Exception:
            session.findById("wnd[1]").sendVKey(0)
        time.sleep(0.5)
        session.findById("wnd[1]/usr/sub:SAPLSVIX:0100/ctxtD0100_FIELD_TAB-LOWER_LIMIT[0,37]").Text = params["werks"]
    session.findById("wnd[1]").sendVKey(0)

    session.findById("wnd[0]/mbar/menu[0]/menu[7]").Select()
    session.findById("wnd[0]/tbar[1]/btn[45]").press()
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()

    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = full_path
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = name_file

    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/tbar[0]/btn[11]").press()

    session.findById("wnd[0]").sendVKey(3)
    session.findById("wnd[0]").sendVKey(3)

# /VWK/MAIREIM015
def request_maireim015_superbesi(session, plant_id: str, job_key: str):
    params, spool = _get_params(job_key, plant_id)

    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = params.get("werks", "5100")
    session.findById("wnd[0]/usr/txtS_COBERT-LOW").Text = "-99999999" 
    session.findById("wnd[0]/usr/txtS_COBERT-HIGH").Text = "999999999"
    session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").Text = "*"
    session.findById("wnd[0]/usr/ctxtS_DISPO-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtS_LIFNR-LOW").Text = ""
    session.findById("wnd[0]/usr/ctxtP_VARIAN").Text = params.get("variant", "/NEWBESI")
    session.findById("wnd[0]").sendVKey(0)

    send_to_background(session, spool, params.get("printer", "locl"))
