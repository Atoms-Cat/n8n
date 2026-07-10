class OutlookClientError(Exception):
    pass

class AuthError(OutlookClientError):
    pass

class NetworkError(OutlookClientError):
    pass

class GraphAPIError(OutlookClientError):
    def __init__(self, status_code, message):
        super().__init__(f"{status_code}: {message}")
        self.status_code = status_code
