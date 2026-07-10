import pytest
from outlook_client.client import OutlookClient

def test_send_email_success(mocker, tmp_path):
    client = OutlookClient("test-id")
    client.access_token = "valid-token"
    
    mock_resp = mocker.Mock()
    mock_resp.status_code = 202
    mock_post = mocker.patch('requests.post', return_value=mock_resp)
    
    # Create dummy attachment
    dummy_file = tmp_path / "test.txt"
    dummy_file.write_text("hello")
    
    client.send_email(["test@example.com"], "Subject", "Body", attachments=[str(dummy_file)])
    
    mock_post.assert_called_once()
    kwargs = mock_post.call_args.kwargs
    payload = kwargs['json']
    assert payload["message"]["subject"] == "Subject"
    assert len(payload["message"]["attachments"]) == 1
    assert payload["message"]["attachments"][0]["name"] == "test.txt"
