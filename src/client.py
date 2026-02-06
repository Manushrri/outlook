"""
Microsoft Outlook Client - OAuth2 Authentication with Microsoft Graph API
"""

import os
import json
import webbrowser
from pathlib import Path
from typing import Optional
import msal
import requests
from dotenv import load_dotenv

load_dotenv()


class OutlookClient:
    """Client for Microsoft Outlook using Microsoft Graph API with OAuth2."""
    
    GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0"
    AUTHORITY = "https://login.microsoftonline.com/common"
    
    # Microsoft Graph API scopes for Outlook
    # Note: offline_access is automatically included by MSAL
    SCOPES = [
        "User.Read",
        "Mail.Read",
        "Mail.ReadWrite",
        "Mail.Send",
        "Calendars.Read",
        "Calendars.ReadWrite",
        "Contacts.Read",
        "Contacts.ReadWrite",
        "MailboxSettings.Read",
        "MailboxSettings.ReadWrite"
    ]
    
    def __init__(self):
        self.client_id = os.getenv("OUTLOOK_CLIENT_ID")
        self.client_secret = os.getenv("OUTLOOK_CLIENT_SECRET")
        # Use the nativeclient redirect URI
        self.redirect_uri = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        self.token_cache_path = Path(__file__).parent.parent / ".token_cache.json"
        
        if not self.client_id:
            raise ValueError("OUTLOOK_CLIENT_ID environment variable is required")
        
        # Initialize MSAL application
        self.app = self._create_msal_app()
        self.access_token: Optional[str] = None
        
        # Try to load cached token
        self._load_cached_token()
    
    def _create_msal_app(self):
        """Create MSAL application for authentication."""
        cache = msal.SerializableTokenCache()
        
        if self.token_cache_path.exists():
            cache.deserialize(self.token_cache_path.read_text())
        
        # Use PublicClientApplication for CLI/MCP apps
        app = msal.PublicClientApplication(
            client_id=self.client_id,
            authority=self.AUTHORITY,
            token_cache=cache
        )
        
        return app
    
    def _save_token_cache(self):
        """Save token cache to file."""
        if self.app.token_cache.has_state_changed:
            self.token_cache_path.write_text(self.app.token_cache.serialize())
    
    def _load_cached_token(self):
        """Try to get token from cache."""
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(self.SCOPES, account=accounts[0])
            if result and "access_token" in result:
                self.access_token = result["access_token"]
                return True
        return False
    
    def authenticate_interactive(self) -> bool:
        """
        Authenticate using device code flow.
        User enters a code at microsoft.com/devicelogin
        """
        print("Starting Microsoft authentication...")
        
        # Use device code flow
        flow = self.app.initiate_device_flow(scopes=self.SCOPES)
        
        if "user_code" not in flow:
            error = flow.get('error_description', flow.get('error', 'Unknown error'))
            print(f"[ERROR] Failed to start authentication: {error}")
            return False
        
        print()
        print("=" * 60)
        print("AUTHENTICATION REQUIRED")
        print("=" * 60)
        print()
        print("1. Open this URL in your browser:")
        print(f"   {flow['verification_uri']}")
        print()
        print(f"2. Enter code: {flow['user_code']}")
        print()
        print("3. Sign in with your Microsoft account")
        print()
        print("=" * 60)
        print("Waiting for sign-in (this may take a minute)...")
        
        # Wait for user to complete authentication
        result = self.app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            self._save_token_cache()
            print()
            print("[OK] Authentication successful!")
            return True
        else:
            error = result.get('error_description', result.get('error', 'Unknown error'))
            print(f"[FAILED] {error}")
            return False
    
    def is_authenticated(self) -> bool:
        """Check if client has valid authentication."""
        return self.access_token is not None
    
    def get_headers(self) -> dict:
        """Get headers for API requests."""
        if not self.access_token:
            raise Exception("Not authenticated. Call authenticate_interactive() first.")
        
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
    
    def request(self, method: str, endpoint: str, **kwargs) -> dict:
        """Make authenticated request to Microsoft Graph API."""
        url = f"{self.GRAPH_API_ENDPOINT}{endpoint}"
        
        # Get base headers and merge with any custom headers passed in kwargs
        headers = self.get_headers()
        if "headers" in kwargs:
            custom_headers = kwargs.pop("headers")
            if custom_headers:
                headers.update(custom_headers)
        
        response = requests.request(method, url, headers=headers, **kwargs)
        
        if response.status_code == 401:
            # Token expired, try to refresh
            if self._load_cached_token():
                headers = self.get_headers()
                response = requests.request(method, url, headers=headers, **kwargs)
            else:
                raise Exception("Authentication expired. Please re-authenticate.")
        
        if not response.ok:
            error_msg = f"{response.status_code} Client Error: {response.reason}"
            try:
                error_data = response.json()
                if "error" in error_data:
                    error_info = error_data["error"]
                    if isinstance(error_info, dict):
                        error_msg += f" for url: {url}"
                        if "message" in error_info:
                            error_msg += f"\nError: {error_info['message']}"
                        if "code" in error_info:
                            error_msg += f"\nCode: {error_info['code']}"
                        if "details" in error_info:
                            error_msg += f"\nDetails: {error_info['details']}"
                    else:
                        error_msg += f" for url: {url}\nError: {error_info}"
            except:
                error_msg += f" for url: {url}"
            raise requests.exceptions.HTTPError(error_msg, response=response)
        
        return response.json() if response.content else {}
    
    def get(self, endpoint: str, **kwargs) -> dict:
        """GET request to Microsoft Graph API."""
        return self.request("GET", endpoint, **kwargs)
    
    def post(self, endpoint: str, **kwargs) -> dict:
        """POST request to Microsoft Graph API."""
        return self.request("POST", endpoint, **kwargs)
    
    def patch(self, endpoint: str, **kwargs) -> dict:
        """PATCH request to Microsoft Graph API."""
        return self.request("PATCH", endpoint, **kwargs)
    
    def delete(self, endpoint: str, **kwargs) -> dict:
        """DELETE request to Microsoft Graph API."""
        return self.request("DELETE", endpoint, **kwargs)
    
    # Basic test methods
    def get_me(self) -> dict:
        """Get current user profile."""
        return self.get("/me")
    
    def get_mailbox_settings(self) -> dict:
        """Get user's mailbox settings."""
        return self.get("/me/mailboxSettings")


# Singleton instance
_client: Optional[OutlookClient] = None


def get_client() -> OutlookClient:
    """Get or create Outlook client instance."""
    global _client
    if _client is None:
        _client = OutlookClient()
    return _client


