"""
Configuration settings for Microsoft Outlook MCP Server
"""

import os
from dataclasses import dataclass
from typing import Optional
from dotenv import load_dotenv

load_dotenv()


@dataclass
class Settings:
    """Application settings loaded from environment variables."""
    
    # Microsoft Azure AD / OAuth settings
    client_id: str = os.getenv("OUTLOOK_CLIENT_ID", "")
    client_secret: Optional[str] = os.getenv("OUTLOOK_CLIENT_SECRET")
    redirect_uri: str = os.getenv("OUTLOOK_REDIRECT_URI", "https://login.microsoftonline.com/common/oauth2/nativeclient")
    token_path: Optional[str] = os.getenv("OUTLOOK_TOKEN_PATH")
    
    # Microsoft Graph API settings
    graph_api_endpoint: str = os.getenv("GRAPH_API_ENDPOINT", "https://graph.microsoft.com/v1.0")
    authority: str = os.getenv("AUTHORITY", "https://login.microsoftonline.com/common")
    
    # Backend API settings (optional, for token provider)
    backend_api_url: Optional[str] = os.getenv("BACKEND_API_URL")
    backend_api_key: Optional[str] = os.getenv("BACKEND_API_KEY")
    mcp_identifier: str = os.getenv("MCP_IDENTIFIER", "outlook-mcp")
    agent_id: Optional[str] = os.getenv("AGENT_ID")
    
    # Scopes for Microsoft Graph API
    scopes: list = None
    
    def __post_init__(self):
        self.scopes = [
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


# Singleton settings instance
settings = Settings()

