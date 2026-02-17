"""
Microsoft Outlook Attachment Tools
"""

import base64
import os
from pathlib import Path
from typing import Optional

from src.workspace_utils import resolve_workspace_file, to_filename


def create_attachment_upload_session(
    client,
    message_id: str,
    attachmentItem: Optional[dict] = None,
    file_path: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Create an upload session for large (>3 MB) message attachments.
    Use when you need to upload attachments in chunks instead of a single request.

    Get message_id from list_messages, search_messages (pick the 'id' of a draft
    message), or create_draft (the returned 'id'). The message must be a draft.

    You can provide either:
    1. file_path (recommended): Path to file relative to WORKSPACE_PATH - will auto-detect name and size
    2. attachmentItem dict: {"attachmentType": "file", "name": "report.pdf", "size": 5242880}

    Args:
        client: The OutlookClient instance
        message_id: The ID of the draft message to create the upload session for.
                    Get from list_messages (folder='drafts'), search_messages, or create_draft.
        attachmentItem: Optional attachment metadata dict with attachmentType, name, and size.
                        Ignored if file_path is provided.
        file_path: Optional path to file (relative to WORKSPACE_PATH). If provided, will auto-detect name and size.
        user_id: Optional user ID (defaults to 'me')

    Returns:
        dict with 'successful', 'data' (containing uploadUrl and expiration), and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }

        # Handle file_path: auto-detect name and size (restricted to WORKSPACE_PATH)
        if file_path:
            try:
                resolved = resolve_workspace_file(file_path, must_exist=True)
                file_path_obj = Path(resolved)
                
                # Get file size
                file_size = file_path_obj.stat().st_size
                
                # Auto-detect name from file
                file_name = file_path_obj.name
                
                # Build attachmentItem from file_path
                attachmentItem = {
                    "attachmentType": "file",
                    "name": file_name,
                    "size": file_size
                }
                
            except (PermissionError, ValueError, FileNotFoundError) as sec_err:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"File error: {str(sec_err)}",
                }
            except Exception as file_error:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Error reading file: {str(file_error)}"
                }

        # Validate attachmentItem
        if not attachmentItem or not isinstance(attachmentItem, dict):
            return {
                "successful": False,
                "data": {},
                "error": "Either provide 'file_path' (relative to WORKSPACE_PATH) or 'attachmentItem' dict with 'attachmentType', 'name', and 'size'. Example: {\"attachmentType\": \"file\", \"name\": \"report.pdf\", \"size\": 5242880}"
            }

        required_fields = ["attachmentType", "name", "size"]
        missing = [f for f in required_fields if f not in attachmentItem]
        if missing:
            return {
                "successful": False,
                "data": {},
                "error": f"attachmentItem is missing required fields: {', '.join(missing)}. Example: {{\"attachmentType\": \"file\", \"name\": \"report.pdf\", \"size\": 5242880}}"
            }

        user = user_id if user_id else "me"
        endpoint = f"/{user}/messages/{message_id}/attachments/createUploadSession"

        payload = {
            "AttachmentItem": attachmentItem
        }

        result = client.post(endpoint, json=payload)

        return {
            "successful": True,
            "data": result
        }

    except Exception as e:
        error_msg = str(e)
        if "400" in error_msg or "Bad Request" in error_msg:
            error_msg += "\n\nCommon issues:\n"
            error_msg += "1. The message_id must be for a DRAFT message\n"
            error_msg += "2. Get draft IDs from: list_messages with folder='drafts' or create_draft\n"
            error_msg += "3. attachmentItem must include: attachmentType, name, size\n"
            error_msg += "4. attachmentType should be 'file'"
        return {
            "successful": False,
            "data": {},
            "error": error_msg
        }


def download_outlook_attachment(
    client,
    message_id: str,
    attachment_id: str,
    file_name: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Downloads a specific file attachment from an email message in a Microsoft Outlook mailbox.
    The attachment must contain 'contentBytes' (binary data) and not be a link or embedded item.
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message containing the attachment
        attachment_id: The ID of the attachment to download
        file_name: The name to save the file as
        user_id: Optional user ID (defaults to 'me')
    
    Returns:
        dict with 'successful', 'data', and optional 'error' fields
    """
    try:
        if not client.is_authenticated():
            return {
                "successful": False,
                "data": {},
                "error": "Not authenticated. Please authenticate first."
            }
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/messages/{message_id}/attachments/{attachment_id}"
        
        # Get the attachment metadata and content
        result = client.get(endpoint)
        
        # Check if it has contentBytes
        if "contentBytes" not in result:
            return {
                "successful": False,
                "data": {},
                "error": "Attachment does not contain downloadable content (contentBytes). It may be a link or embedded item."
            }
        
        # Resolve destination path inside workspace (no absolute paths allowed)
        try:
            dest_path = resolve_workspace_file(file_name, must_exist=False)
        except (PermissionError, ValueError, FileNotFoundError) as e:
            return {
                "successful": False,
                "data": {},
                "error": str(e),
            }

        # Decode and save the file securely inside workspace
        content_bytes = base64.b64decode(result["contentBytes"])

        with open(dest_path, "wb") as f:
            f.write(content_bytes)
        
        return {
            "successful": True,
            "data": {
                # Return workspace-relative filename so callers never see full paths
                "file_name": to_filename(dest_path),
                "size": len(content_bytes),
                "content_type": result.get("contentType", "unknown"),
                "name": result.get("name", file_name)
            }
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }

