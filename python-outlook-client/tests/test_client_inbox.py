import pytest
from outlook_client.client import OutlookClient
from outlook_client.errors import GraphAPIError

def test_get_inbox_success(mocker):
    client = OutlookClient("test-id")
    client.access_token = "valid-token"
    
    mock_resp = mocker.Mock()
    mock_resp.status_code = 200
    mock_resp.json.return_value = {"value": [{"subject": "Hello"}]}
    mocker.patch('requests.get', return_value=mock_resp)
    
    messages = client.get_inbox(top=1)
    assert len(messages) == 1
    assert messages[0]["subject"] == "Hello"

def test_get_inbox_api_error(mocker):
    client = OutlookClient("test-id")
    client.access_token = "valid-token"
    
    mock_resp = mocker.Mock()
    mock_resp.status_code = 401
    mock_resp.text = "Unauthorized"
    mocker.patch('requests.get', return_value=mock_resp)
    
    with pytest.raises(GraphAPIError):
        client.get_inbox()
