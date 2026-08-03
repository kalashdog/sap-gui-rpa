"""
Funções genéricas para extrair e baixar spools do SAP.
"""
import os
import logging
from datetime import datetime
import pythoncom

from config.settings import settings
from core.utils import get_target_export_dir

def extract_sp02_job(session, plant_id: str, job_key: str, job_data: dict):
    try:
        plant_config = settings.config["plants"].get(plant_id, {})
        folder_name = plant_config.get("folder_name", plant_id)
        inner_base_path = plant_config.get("base_path", "")
        
        plant_params = job_data.get("plant_params", {}).get(plant_id, {})
        local_extract = plant_params.get("local_extract", "")
        name_file = plant_params.get("name_file", f"{job_key}.txt").format(date=datetime.now())
        spool_name = plant_params.get("spool_name") or job_data.get("spool_name", job_key)
        
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

        # Snapshot de madrugada: entre 03:00-04:00, exporta com nome especial (1x/dia)
        madrugada_name = plant_params.get("name_file_madrugada")
        if madrugada_name:
            now = datetime.now()
            if 3 <= now.hour < 4:
                madrugada_path = os.path.join(full_path, madrugada_name)
                already_exported = False
                if os.path.exists(madrugada_path):
                    mod_date = datetime.fromtimestamp(os.path.getmtime(madrugada_path)).date()
                    already_exported = (mod_date == now.date())
                if not already_exported:
                    name_file = madrugada_name
        
        logging.info(f"Extraction for '{job_key}' (spool: '{spool_name}'): folder '{full_path}', file '{name_file}'")
        
        for i in range(3, 31):
            try:
                if session.findById(f"wnd[0]/usr/lbl[51,{i}]").Text == spool_name:
                    session.findById(f"wnd[0]/usr/chk[1,{i}]").Selected = True
                    session.findById(f"wnd[0]/usr/lbl[14,{i}]").SetFocus()
                    session.findById("wnd[0]").sendVKey(2)
                    session.findById("wnd[0]/tbar[1]/btn[48]").press()
                    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    session.findById("wnd[1]/usr/ctxtDY_PATH").Text = full_path
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = name_file
                    session.findById("wnd[1]/tbar[0]/btn[11]").press()
                    session.findById("wnd[0]").sendVKey(3)
                    session.findById("wnd[0]/tbar[1]/btn[14]").press()
                    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
                    return True
                    
            except pythoncom.com_error:
                continue
                
        logging.info(f"Spool '{spool_name}' ainda nao apareceu no SP02. Tentando no proximo ciclo.")
        return False
        
    except KeyError as e:
        logging.error(f"Config ausente para extracao de '{job_key}': {e}")
        return False
    except pythoncom.com_error as e:
        logging.error(f"Erro COM SAP na extracao de '{job_key}': {e}")
        return False

