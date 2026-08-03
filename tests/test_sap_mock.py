import pytest 
from unittest.mock import MagicMock, patch
from transactions.request import request_mb52

def test_request_mb52_mocked_sap():
    mock_session = MagicMock()
    mock_txt_variant = MagicMock()
    
    def mock_find_by_id(element_id):
        if "txtV-LOW" in element_id:
            return mock_txt_variant
        return MagicMock()
        
    mock_session.findById.side_effect = mock_find_by_id
    
    with patch("transactions.request.send_to_background") as mock_send_bg:
        request_mb52(mock_session, plant_id="01-Anchieta", job_key="MB52_AUTO")
        
        assert mock_txt_variant.Text == "/VINIAUTO52"
        mock_send_bg.assert_called_once()
        args, kwargs = mock_send_bg.call_args
        assert args[1] == "MB52AUTO"
