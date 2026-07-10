import os
import msal
import json
import requests
from outlook_client.client import OutlookClient

def test_outlook_client_init(mocker):
    mocker.patch('os.path.exists', return_value=False)
    
    mock_response = mocker.Mock()
    mock_response.status_code = 200
    mock_response.text = json.dumps({
        "token_endpoint": "https://login.microsoftonline.com/test-tenant/oauth2/v2.0/token",
        "authorization_endpoint": "https://login.microsoftonline.com/test-tenant/oauth2/v2.0/authorize",
        "issuer": "https://login.microsoftonline.com/test-tenant/v2.0"
    })
    mocker.patch('requests.Session.get', return_value=mock_response)
    
    client = OutlookClient(client_id="test-id", tenant_id="test-tenant")
    
    assert client.client_id == "test-id"
    assert client.authority == "https://login.microsoftonline.com/test-tenant"
    assert isinstance(client.app, msal.PublicClientApplication)
