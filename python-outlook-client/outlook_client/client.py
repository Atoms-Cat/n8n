import os
import msal
import requests
from .errors import AuthError, GraphAPIError, NetworkError

class OutlookClient:
    def __init__(self, client_id: str, tenant_id: str = "common", client_secret: str = None, http_client=None):
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.client_secret = client_secret
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.token_cache = msal.SerializableTokenCache()
        self.cache_path = os.path.expanduser("~/.outlook_client_token.json")
        
        if os.path.exists(self.cache_path):
            with open(self.cache_path, "r") as f:
                self.token_cache.deserialize(f.read())
                
        if self.client_secret:
            self.app = msal.ConfidentialClientApplication(
                self.client_id, authority=self.authority,
                client_credential=self.client_secret, token_cache=self.token_cache,
                http_client=http_client
            )
        else:
            self.app = msal.PublicClientApplication(
                self.client_id, authority=self.authority, token_cache=self.token_cache,
                http_client=http_client
            )
        self.access_token = None

    def _save_cache(self):
        if self.token_cache.has_state_changed:
            with open(self.cache_path, "w") as f:
                f.write(self.token_cache.serialize())

    def login_with_browser(self):
        scopes = ["https://graph.microsoft.com/.default"]
        result = self.app.acquire_token_interactive(scopes=scopes)
        if "access_token" in result:
            self.access_token = result["access_token"]
            self._save_cache()
        else:
            raise AuthError(result.get("error_description", "Unknown auth error"))

    def login_with_username_password(self, username, password):
        scopes = ["https://graph.microsoft.com/.default"]
        result = self.app.acquire_token_by_username_password(username, password, scopes=scopes)
        if "access_token" in result:
            self.access_token = result["access_token"]
            self._save_cache()
        else:
            raise AuthError(result.get("error_description", "Unknown auth error"))

    def login_with_sso(self, username_hint=None):
        scopes = ["https://graph.microsoft.com/.default"]
        accounts = self.app.get_accounts(username=username_hint)
        if accounts:
            result = self.app.acquire_token_silent(scopes, account=accounts[0])
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                self._save_cache()
                return
        raise AuthError("No cached token found or silent auth failed")

    def _get_headers(self):
        if not self.access_token:
            raise AuthError("Not authenticated. Call a login method first.")
        return {"Authorization": f"Bearer {self.access_token}", "Content-Type": "application/json"}

    def get_inbox(self, top: int = 10, unread_only: bool = False) -> list:
        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top={top}"
        if unread_only:
            url += "&$filter=isRead eq false"
        
        try:
            response = requests.get(url, headers=self._get_headers())
        except requests.RequestException as e:
            raise NetworkError(str(e))
            
        if response.status_code == 200:
            return response.json().get("value", [])
        else:
            raise GraphAPIError(response.status_code, response.text)
