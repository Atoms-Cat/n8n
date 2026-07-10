import pytest
from outlook_client.client import OutlookClient
from outlook_client.errors import AuthError

def test_login_with_browser_success(mocker):
    client = OutlookClient("test-id")
    mocker.patch.object(client.app, 'acquire_token_interactive', return_value={"access_token": "token123"})
    mocker.patch.object(client, '_save_cache')
    
    client.login_with_browser()
    assert client.access_token == "token123"
    client._save_cache.assert_called_once()

def test_login_with_sso_failure(mocker):
    client = OutlookClient("test-id")
    mocker.patch.object(client.app, 'get_accounts', return_value=[])
    
    with pytest.raises(AuthError):
        client.login_with_sso()
