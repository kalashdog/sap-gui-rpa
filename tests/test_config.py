import pytest 
from config.settings import settings

def test_config_loads_plants():
    """Testa se o arquivo sapscripts_config.json é carregado corretamente e contém as plantas."""
    assert "plants" in settings.config
    assert "01-Anchieta" in settings.config["plants"]
    
    plant_anchieta = settings.config["plants"]["01-Anchieta"]
    assert plant_anchieta["code"] == "ANC"

def test_config_loads_jobs():
    """Testa se os jobs foram carregados corretamente."""
    assert "jobs" in settings.config
    assert "MB52_AUTO" in settings.config["jobs"]
    
    mb52 = settings.config["jobs"]["MB52_AUTO"]
    assert mb52["transaction"] == "MB52"
    assert mb52["dashboard"] == "000 - Compartilhado"
