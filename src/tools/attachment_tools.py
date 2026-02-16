"""
Microsoft Outlook Attachment Tools
"""

import base64
from typing import Optional


def create_attachment_upload_session(
    client,
    message_id: str,
    attachmentItem: dict,
    user_id: Optional[str] = None
) -> dict:
    """
    Create an upload session for large (>3 MB) message attachments.
    Use when you need to upload attachments in chunks instead of a single request.

    Get message_id from list_messages, search_messages (pick the 'id' of a draft
    message), or create_draft (the returned 'id'). The message must be a draft.

    attachmentItem must include:
      - attachmentType: "file"
      - name: file name (e.g. "large-report.pdf")
      - size: total file size in bytes

    Example attachmentItem:
      {"attachmentType": "file", "name": "report.pdf", "size": 5242880}

    Args:
        client: The OutlookClient instance
        message_id: The ID of the draft message to create the upload session for.
                    Get from list_messages (folder='drafts'), search_messages, or create_draft.
        attachmentItem: Attachment metadata dict with attachmentType, name, and size.
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

        if not attachmentItem or not isinstance(attachmentItem, dict):
            return {
                "successful": False,
                "data": {},
                "error": "attachmentItem must be a dict with 'attachmentType', 'name', and 'size'. Example: {\"attachmentType\": \"file\", \"name\": \"report.pdf\", \"size\": 5242880}"
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
        
        # Decode and save the file
        content_bytes = base64.b64decode(result["contentBytes"])
        
        with open(file_name, "wb") as f:
            f.write(content_bytes)
        
        return {
            "successful": True,
            "data": {
                "file_name": file_name,
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

