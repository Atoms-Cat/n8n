import os
import msal

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
