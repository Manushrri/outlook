"""
Microsoft Outlook Mail Tools
"""

import base64
import mimetypes
from pathlib import Path
from typing import Optional, List


def add_mail_attachment(
    client,
    message_id: str,
    name: str,
    odata_type: str,
    contentBytes: Optional[str] = None,
    file_path: Optional[str] = None,
    contentId: Optional[str] = None,
    contentLocation: Optional[str] = None,
    contentType: Optional[str] = None,
    isInline: Optional[bool] = None,
    item: Optional[dict] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Add an attachment to an email message.
    Use when you have a message id and need to attach a small (<3 MB) file or reference.
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to attach to
        name: The name of the attachment
        odata_type: The OData type of the attachment (e.g., "#microsoft.graph.fileAttachment")
        contentBytes: Base64-encoded content of the file (optional if file_path is provided)
        file_path: Path to the file to attach (will be automatically encoded to base64)
        contentId: Optional content ID for inline attachments
        contentLocation: Optional content location URL
        contentType: Optional MIME type of the attachment (auto-detected from file_path if not provided)
        isInline: Whether the attachment is inline
        item: Optional item data for item attachments
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
        
        # Handle file_path: read file and encode to base64
        if file_path:
            try:
                file_path_obj = Path(file_path)
                if not file_path_obj.exists():
                    return {
                        "successful": False,
                        "data": {},
                        "error": f"File not found: {file_path}"
                    }
                
                # Check file size (3 MB limit)
                file_size = file_path_obj.stat().st_size
                if file_size > 3 * 1024 * 1024:  # 3 MB
                    return {
                        "successful": False,
                        "data": {},
                        "error": f"File size ({file_size / 1024 / 1024:.2f} MB) exceeds 3 MB limit. Use a smaller file or upload via other method."
                    }
                
                # Read file and encode to base64
                with open(file_path_obj, 'rb') as f:
                    file_content = f.read()
                    contentBytes = base64.b64encode(file_content).decode('utf-8')
                
                # Auto-detect content type if not provided
                if not contentType:
                    detected_type, _ = mimetypes.guess_type(str(file_path_obj))
                    if detected_type:
                        contentType = detected_type
                
                # Use file name if name not provided
                if not name or name.strip() == "":
                    name = file_path_obj.name
                    
            except Exception as file_error:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Error reading file: {str(file_error)}"
                }
        
        # Validate that contentBytes is provided (either directly or via file_path)
        if not contentBytes or not contentBytes.strip():
            return {
                "successful": False,
                "data": {},
                "error": "Either contentBytes (base64-encoded) or file_path must be provided."
            }
        
        # First, check if the message is a draft (attachments can only be added to drafts)
        user = user_id if user_id else "me"
        check_endpoint = f"/{user}/messages/{message_id}?$select=isDraft"
        
        try:
            message_info = client.get(check_endpoint)
            if not message_info.get("isDraft", False):
                return {
                    "successful": False,
                    "data": {},
                    "error": "Cannot add attachment to received/sent message. Only draft messages can have attachments added. Please use a draft message ID. Get draft message IDs using list_messages with folder='drafts' or create a draft first using create_draft_email."
                }
        except Exception as check_error:
            # If we can't check (e.g., message doesn't exist), proceed and let the API return the error
            pass
        
        # Build the attachment payload
        attachment_data = {
            "@odata.type": odata_type,
            "name": name,
            "contentBytes": contentBytes
        }
        
        # Add optional fields if provided
        if contentId is not None:
            attachment_data["contentId"] = contentId
        if contentLocation is not None:
            attachment_data["contentLocation"] = contentLocation
        if contentType is not None:
            attachment_data["contentType"] = contentType
        if isInline is not None:
            attachment_data["isInline"] = isInline
        if item is not None and isinstance(item, dict) and len(item) > 0:
            attachment_data["item"] = item
        
        # Determine the endpoint
        endpoint = f"/{user}/messages/{message_id}/attachments"
        
        # Make the API call
        result = client.post(endpoint, json=attachment_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        error_msg = str(e)
        # Provide helpful guidance for common errors
        if "400" in error_msg:
            error_msg += "\n\nCommon issues:\n"
            error_msg += "1. The message_id must be for a DRAFT message (not sent/received)\n"
            error_msg += "2. Get draft message IDs using: list_messages with folder='drafts'\n"
            error_msg += "3. Or create a draft first using: create_draft_email\n"
            error_msg += "4. Either provide contentBytes (base64-encoded) OR file_path (path to file)\n"
            error_msg += "5. File size must be less than 3 MB\n"
            error_msg += "6. Do NOT include empty objects {} for optional fields like 'item'"
        return {
            "successful": False,
            "data": {},
            "error": error_msg
        }


def delete_message(
    client,
    message_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Permanently delete an Outlook email message by its message_id.
    Use when removing unwanted messages, cleaning up drafts, or performing
    mailbox maintenance.

    Get message_id from list_messages or search_messages (pick the 'id' field
    of the message you want to delete).

    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to delete.
                    Get from list_messages or search_messages.
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

        user = user_id if user_id else "me"

        # First, verify the message exists and get its details for confirmation
        check_endpoint = f"/{user}/messages/{message_id}?$select=id,subject,from,receivedDateTime,isDraft"
        try:
            message_info = client.get(check_endpoint)
        except Exception as check_error:
            error_msg = str(check_error)
            if "404" in error_msg or "Not Found" in error_msg:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Message not found. The message_id '{message_id}' does not exist or has already been deleted. Get a valid message_id from list_messages or search_messages."
                }
            return {
                "successful": False,
                "data": {},
                "error": f"Could not verify message: {error_msg}"
            }

        # Now delete the verified message
        endpoint = f"/{user}/messages/{message_id}"

        # DELETE returns 204 No Content on success
        client.delete(endpoint)

        # Return info about what was deleted
        deleted_subject = message_info.get("subject", "Unknown")
        deleted_from = message_info.get("from", {}).get("emailAddress", {}).get("address", "Unknown")
        was_draft = message_info.get("isDraft", False)

        return {
            "successful": True,
            "data": {
                "message": "Message deleted successfully",
                "deleted_subject": deleted_subject,
                "deleted_from": deleted_from,
                "was_draft": was_draft
            }
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def create_draft(
    client,
    subject: str,
    body: str,
    to_recipients: List[str],
    cc_recipients: Optional[List[str]] = None,
    bcc_recipients: Optional[List[str]] = None,
    is_html: Optional[bool] = None,
    conversation_id: Optional[str] = None,
    attachment: Optional[dict] = None
) -> dict:
    """
    Creates an Outlook email draft with subject, body, recipients, and an optional attachment.
    Supports creating drafts as part of existing conversation threads.
    
    Args:
        client: The OutlookClient instance
        subject: The subject of the email
        body: The body content of the email
        to_recipients: List of recipient email addresses
        cc_recipients: Optional list of CC email addresses
        bcc_recipients: Optional list of BCC email addresses
        is_html: Whether body is HTML
        conversation_id: Optional conversation ID for threading
        attachment: Optional attachment dict with name, contentType, contentBytes
    
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
        
        # Build the draft payload
        draft_data = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body
            },
            "toRecipients": [
                {"emailAddress": {"address": email}} for email in to_recipients
            ]
        }
        
        if cc_recipients:
            draft_data["ccRecipients"] = [
                {"emailAddress": {"address": email}} for email in cc_recipients
            ]
        
        if bcc_recipients:
            draft_data["bccRecipients"] = [
                {"emailAddress": {"address": email}} for email in bcc_recipients
            ]
        
        if conversation_id:
            draft_data["conversationId"] = conversation_id
        
        # Make the API call to create draft
        endpoint = "/me/messages"
        result = client.post(endpoint, json=draft_data)
        
        # Add attachment if provided
        if attachment and result.get("id"):
            message_id = result["id"]
            attachment_data = {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": attachment.get("name"),
                "contentType": attachment.get("contentType", "application/octet-stream"),
                "contentBytes": attachment.get("contentBytes")
            }
            attachment_endpoint = f"/me/messages/{message_id}/attachments"
            client.post(attachment_endpoint, json=attachment_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def create_draft_reply(
    client,
    message_id: str,
    comment: Optional[str] = None,
    cc_emails: Optional[List[str]] = None,
    bcc_emails: Optional[List[str]] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Creates a draft reply in the specified user's Outlook mailbox to an existing message.
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to reply to
        comment: Optional comment/reply text
        cc_emails: Optional list of CC email addresses
        bcc_emails: Optional list of BCC email addresses
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
        
        # Build the reply payload
        reply_data = {}
        
        if comment:
            reply_data["comment"] = comment
        
        # Add recipients if provided
        message_updates = {}
        if cc_emails:
            message_updates["ccRecipients"] = [
                {"emailAddress": {"address": email}} for email in cc_emails
            ]
        if bcc_emails:
            message_updates["bccRecipients"] = [
                {"emailAddress": {"address": email}} for email in bcc_emails
            ]
        
        if message_updates:
            reply_data["message"] = message_updates
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/messages/{message_id}/createReply"
        
        # Make the API call
        result = client.post(endpoint, json=reply_data if reply_data else None)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def forward_message(
    client,
    message_id: str,
    to_recipients: List[str],
    comment: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Forward an existing email message to new recipients.
    Use when you need to send an existing email to someone else.

    Get message_id from list_messages or search_messages (pick the 'id' field
    of the message you want to forward).

    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to forward.
                    Get from list_messages or search_messages.
        to_recipients: List of recipient email addresses to forward to.
        comment: Optional message to include with the forwarded email.
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

        if not to_recipients:
            return {
                "successful": False,
                "data": {},
                "error": "to_recipients list cannot be empty. Provide at least one email address."
            }

        user = user_id if user_id else "me"

        # Build the forward payload
        forward_data = {
            "toRecipients": [
                {"emailAddress": {"address": email}} for email in to_recipients
            ]
        }

        if comment is not None:
            forward_data["comment"] = comment

        endpoint = f"/{user}/messages/{message_id}/forward"

        # forward action returns no content on success (202)
        client.post(endpoint, json=forward_data)

        return {
            "successful": True,
            "data": {"message": "Message forwarded successfully"}
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def get_message(
    client,
    message_id: str,
    select: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieves a specific email message by its ID from the user's Outlook mailbox.
    Use the 'select' parameter to include specific fields like 'internetMessageHeaders'.
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to retrieve
        select: Optional comma-separated list of properties to select
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
        
        # Build query parameters
        params = {}
        if select:
            params["$select"] = select
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/messages/{message_id}"
        
        # Make the API call
        result = client.get(endpoint, params=params if params else None)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def move_message(
    client,
    message_id: str,
    destination_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Move a message to another folder within the specified user's mailbox.
    This creates a new copy of the message in the destination folder and removes the original message.
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to move
        destination_id: The ID of the destination folder (or well-known name like 'inbox', 'drafts', 'deleteditems')
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
        
        # Build the move payload
        move_data = {
            "destinationId": destination_id
        }
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/messages/{message_id}/move"
        
        # Make the API call
        result = client.post(endpoint, json=move_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def reply_email(
    client,
    message_id: str,
    comment: str,
    cc_emails: Optional[List[str]] = None,
    bcc_emails: Optional[List[str]] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Sends a plain text reply to an Outlook email message, identified by message_id,
    allowing optional CC and BCC recipients.
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to reply to
        comment: The reply text/comment to send
        cc_emails: Optional list of CC email addresses
        bcc_emails: Optional list of BCC email addresses
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
        
        # Build the reply payload
        reply_data = {
            "comment": comment
        }
        
        # Add recipients if provided
        message_updates = {}
        if cc_emails:
            message_updates["ccRecipients"] = [
                {"emailAddress": {"address": email}} for email in cc_emails
            ]
        if bcc_emails:
            message_updates["bccRecipients"] = [
                {"emailAddress": {"address": email}} for email in bcc_emails
            ]
        
        if message_updates:
            reply_data["message"] = message_updates
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/messages/{message_id}/reply"
        
        # Make the API call (reply action returns no content on success)
        client.post(endpoint, json=reply_data)
        
        return {
            "successful": True,
            "data": {"message": "Reply sent successfully"}
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def search_messages(
    client,
    query: str,
    fromEmail: Optional[str] = None,
    subject: Optional[str] = None,
    hasAttachments: Optional[bool] = None,
    from_index: Optional[int] = None,
    size: Optional[int] = None,
    enable_top_results: Optional[bool] = None
) -> dict:
    """
    Searches messages in a Microsoft 365 or enterprise Outlook account mailbox,
    supporting filters for sender, subject, attachments, pagination, and sorting by relevance or date.
    
    Args:
        client: The OutlookClient instance
        query: The search query string
        fromEmail: Optional sender email address to filter by
        subject: Optional subject to search for
        hasAttachments: Optional filter for messages with attachments
        from_index: Optional starting index for pagination
        size: Optional number of results to return
        enable_top_results: Optional flag to enable top results sorting
    
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
        
        # Normalize empty strings to None for optional parameters
        if fromEmail == "":
            fromEmail = None
        if subject == "":
            subject = None
        if query == "":
            return {
                "successful": False,
                "data": {},
                "error": "query parameter cannot be empty"
            }
        
        # Build query parameters using $filter (more reliable than $search for personal accounts)
        params = {}
        filter_parts = []
        
        # Escape single quotes in query strings to prevent OData injection
        def escape_odata_string(s: str) -> str:
            return s.replace("'", "''")
        
        # Build filter expressions
        # Note: contains() on bodyPreview may not be supported in all contexts
        # So we'll search in subject only, or use subject parameter for subject-specific search
        if query:
            # Search query in subject (bodyPreview contains() may not be supported)
            escaped_query = escape_odata_string(query)
            filter_parts.append(f"contains(subject, '{escaped_query}')")
        
        if subject:
            # Additional subject filter
            escaped_subject = escape_odata_string(subject)
            filter_parts.append(f"contains(subject, '{escaped_subject}')")
        
        if fromEmail:
            escaped_email = escape_odata_string(fromEmail)
            filter_parts.append(f"from/emailAddress/address eq '{escaped_email}'")
        
        if hasAttachments is not None:
            filter_parts.append(f"hasAttachments eq {str(hasAttachments).lower()}")
        
        # Combine all filters
        if filter_parts:
            params["$filter"] = " and ".join(filter_parts)
        
        if size is not None:
            params["$top"] = size
        if from_index is not None:
            params["$skip"] = from_index
        
        # Determine the endpoint
        endpoint = "/me/messages"
        
        # Make the API call
        result = client.get(endpoint, params=params if params else None)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def update_email(
    client,
    message_id: str,
    subject: Optional[str] = None,
    body: Optional[dict] = None,
    to_recipients: Optional[List[str]] = None,
    cc_recipients: Optional[List[str]] = None,
    bcc_recipients: Optional[List[str]] = None,
    importance: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Updates specified properties of an existing email message.
    message_id must identify a valid message within the specified user_id's mailbox.
    
    NOTE: Only draft messages can be updated. Received messages cannot be modified.
    Use outlook_list_messages to find draft messages (isDraft: true).
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to update (must be a draft message)
        subject: Optional subject of the email
        body: Optional body object with contentType and content, e.g., {"contentType": "text", "content": "Hello"}
        to_recipients: Optional list of TO recipient email addresses
        cc_recipients: Optional list of CC recipient email addresses
        bcc_recipients: Optional list of BCC recipient email addresses
        importance: Optional importance level (low, normal, high)
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
        
        # First, check if the message is a draft
        user = user_id if user_id else "me"
        check_endpoint = f"/{user}/messages/{message_id}?$select=isDraft"
        
        try:
            message_info = client.get(check_endpoint)
            if not message_info.get("isDraft", False):
                return {
                    "successful": False,
                    "data": {},
                    "error": "Cannot update received message. Only draft messages can be updated. Please use a draft message ID or create a draft first using outlook_create_draft."
                }
        except Exception as check_error:
            # If we can't check, proceed anyway and let the API return the error
            pass
        
        # Build the message update payload
        message_data = {}
        
        if subject is not None:
            message_data["subject"] = subject
        
        if body is not None:
            # Validate body format
            if isinstance(body, dict):
                if "contentType" not in body or "content" not in body:
                    return {
                        "successful": False,
                        "data": {},
                        "error": "Body must be a dict with 'contentType' and 'content' fields, e.g., {'contentType': 'text', 'content': 'Hello'}"
                    }
                message_data["body"] = body
            else:
                return {
                    "successful": False,
                    "data": {},
                    "error": "Body must be a dict with 'contentType' and 'content' fields, e.g., {'contentType': 'text', 'content': 'Hello'}"
                }
        
        if to_recipients is not None:
            message_data["toRecipients"] = [
                {"emailAddress": {"address": email}} for email in to_recipients
            ]
        
        if cc_recipients is not None:
            message_data["ccRecipients"] = [
                {"emailAddress": {"address": email}} for email in cc_recipients
            ]
        
        if bcc_recipients is not None:
            message_data["bccRecipients"] = [
                {"emailAddress": {"address": email}} for email in bcc_recipients
            ]
        
        if importance is not None:
            message_data["importance"] = importance
        
        # Check if we have at least one field to update
        if not message_data:
            return {
                "successful": False,
                "data": {},
                "error": "At least one field (subject, body, to_recipients, cc_recipients, bcc_recipients, or importance) must be provided to update."
            }
        
        # Determine the endpoint
        endpoint = f"/{user}/messages/{message_id}"
        
        # Make the API call
        result = client.patch(endpoint, json=message_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        error_msg = str(e)
        # Provide more helpful error messages
        if "400" in error_msg or "Bad Request" in error_msg:
            if "draft" in error_msg.lower() or "cannot" in error_msg.lower():
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Cannot update this message. Only draft messages can be updated. Error: {error_msg}"
                }
        return {
            "successful": False,
            "data": {},
            "error": error_msg
        }


def batch_move_messages(
    client,
    message_ids: List[str],
    destination_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Batch-move up to 20 Outlook messages to a destination folder in a single
    Microsoft Graph $batch call. Use when moving multiple messages to avoid
    per-message move API calls.

    Get message_ids from list_messages or search_messages (pick the 'id' field
    of each message you want to move). Get destination_id from list_mail_folders
    (use the 'id' of the target folder, or a well-known name like 'inbox',
    'drafts', 'deleteditems', 'sentitems').

    Args:
        client: The OutlookClient instance
        message_ids: List of message IDs to move (max 20)
        destination_id: The destination folder ID or well-known name.
                        Get this from list_mail_folders.
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

        if not message_ids:
            return {
                "successful": False,
                "data": {},
                "error": "message_ids list cannot be empty."
            }

        if len(message_ids) > 20:
            return {
                "successful": False,
                "data": {},
                "error": "Cannot batch-move more than 20 messages at once. Please split into smaller batches."
            }

        user = user_id if user_id else "me"

        # Build individual requests for the $batch payload
        requests_list = []
        for idx, msg_id in enumerate(message_ids):
            requests_list.append({
                "id": str(idx + 1),
                "method": "POST",
                "url": f"/{user}/messages/{msg_id}/move",
                "headers": {"Content-Type": "application/json"},
                "body": {"destinationId": destination_id}
            })

        batch_payload = {"requests": requests_list}

        # POST to the $batch endpoint
        result = client.post("/$batch", json=batch_payload)

        # Summarise per-request outcomes
        responses = result.get("responses", [])
        failures = [r for r in responses if r.get("status", 0) >= 400]

        if failures:
            return {
                "successful": False,
                "data": result,
                "error": f"{len(failures)} of {len(message_ids)} move operations failed. Check 'data.responses' for details."
            }

        return {
            "successful": True,
            "data": result
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def batch_update_messages(
    client,
    updates: List[dict],
    user_id: Optional[str] = None
) -> dict:
    """
    Batch-update up to 20 Outlook messages per call using Microsoft Graph JSON
    batching. Use when marking multiple messages read/unread or updating other
    properties to avoid per-message PATCH calls.

    Each item in 'updates' must contain a 'message_id' key and one or more
    updatable properties such as:
      - isRead (bool): mark read/unread
      - categories (list[str]): assign categories
      - importance (str): 'low', 'normal', or 'high'
      - flag (dict): e.g. {"flagStatus": "flagged"}
      - inferenceClassification (str): 'focused' or 'other'

    Get message_id values from list_messages or search_messages (pick the 'id'
    field of each message you want to update).

    Example 'updates':
      [
        {"message_id": "AAMk...", "isRead": true},
        {"message_id": "AAMk...", "isRead": false, "importance": "high"}
      ]

    Args:
        client: The OutlookClient instance
        updates: List of dicts, each with 'message_id' and properties to update (max 20)
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

        if not updates:
            return {
                "successful": False,
                "data": {},
                "error": "updates list cannot be empty."
            }

        if len(updates) > 20:
            return {
                "successful": False,
                "data": {},
                "error": "Cannot batch-update more than 20 messages at once. Please split into smaller batches."
            }

        user = user_id if user_id else "me"

        # Build individual requests for the $batch payload
        requests_list = []
        for idx, update in enumerate(updates):
            msg_id = update.get("message_id")
            if not msg_id:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Update at index {idx} is missing required 'message_id' field."
                }

            # Build the PATCH body (everything except message_id)
            patch_body = {k: v for k, v in update.items() if k != "message_id"}

            if not patch_body:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Update at index {idx} has no properties to update besides 'message_id'."
                }

            requests_list.append({
                "id": str(idx + 1),
                "method": "PATCH",
                "url": f"/{user}/messages/{msg_id}",
                "headers": {"Content-Type": "application/json"},
                "body": patch_body
            })

        batch_payload = {"requests": requests_list}

        # POST to the $batch endpoint
        result = client.post("/$batch", json=batch_payload)

        # Summarise per-request outcomes
        responses = result.get("responses", [])
        failures = [r for r in responses if r.get("status", 0) >= 400]

        if failures:
            return {
                "successful": False,
                "data": result,
                "error": f"{len(failures)} of {len(updates)} update operations failed. Check 'data.responses' for details."
            }

        return {
            "successful": True,
            "data": result
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def permanent_delete_message(
    client,
    message_id: str,
    mail_folder_id: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Permanently delete an Outlook message by moving it to the Purges folder
    in the dumpster. Unlike standard delete_message, this action makes the
    message UNRECOVERABLE.

    IMPORTANT: This is NOT the same as delete_message — permanentDelete is
    irreversible. Not available in US Government L4, L5 (DOD), or
    China (21Vianet) deployments.

    Get message_id from list_messages or search_messages (pick the 'id' field).
    Optionally provide mail_folder_id from list_mail_folders if the message
    is in a specific folder.

    Args:
        client: The OutlookClient instance
        message_id: The ID of the message to permanently delete.
                    Get from list_messages or search_messages.
        mail_folder_id: Optional folder ID containing the message.
                        Get from list_mail_folders.
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

        user = user_id if user_id else "me"

        # Verify the message exists first
        check_endpoint = f"/{user}/messages/{message_id}?$select=id,subject,isDraft"
        try:
            message_info = client.get(check_endpoint)
        except Exception as check_error:
            error_msg = str(check_error)
            if "404" in error_msg or "Not Found" in error_msg:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Message not found. The message_id '{message_id}' does not exist or has already been deleted. Get a valid message_id from list_messages or search_messages."
                }
            return {
                "successful": False,
                "data": {},
                "error": f"Could not verify message: {error_msg}"
            }

        # Build the endpoint for permanentDelete
        if mail_folder_id:
            endpoint = f"/{user}/mailFolders/{mail_folder_id}/messages/{message_id}/permanentDelete"
        else:
            endpoint = f"/{user}/messages/{message_id}/permanentDelete"

        # POST to permanentDelete (returns 204 No Content on success)
        client.post(endpoint)

        deleted_subject = message_info.get("subject", "Unknown")

        return {
            "successful": True,
            "data": {
                "message": "Message permanently deleted (unrecoverable)",
                "deleted_subject": deleted_subject
            }
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def query_emails(
    client,
    folder: Optional[str] = None,
    filter: Optional[str] = None,
    orderby: Optional[str] = None,
    select: Optional[List[str]] = None,
    skip: Optional[int] = None,
    top: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Query Outlook emails within a SINGLE folder using OData filters.
    Build precise server-side filters for dates, read status, importance,
    subjects, attachments, and conversations. Best for structured queries
    on message metadata within a specific folder. Returns up to 100 messages
    per request with pagination support.

    - Searches SINGLE folder only (inbox, sentitems, etc.) - NOT across all folders
    - For cross-folder/mailbox-wide search: Use search_messages
    - Server-side filters: dates, importance, isRead, hasAttachments, subjects, conversationId
    - CRITICAL: Always check response['@odata.nextLink'] for pagination
    - Limitations: Recipient/body filtering requires search_messages

    Get folder name or ID from list_mail_folders. Common well-known names:
    'inbox', 'drafts', 'sentitems', 'deleteditems', 'junkemail', 'archive'.

    Filter examples:
      - "isRead eq false"
      - "importance eq 'high'"
      - "hasAttachments eq true"
      - "receivedDateTime ge 2026-02-01T00:00:00Z"
      - "contains(subject, 'meeting')"

    Args:
        client: The OutlookClient instance
        folder: Folder name or ID. Get from list_mail_folders.
                Defaults to inbox.
        filter: OData filter expression
        orderby: Property to order by (e.g. 'receivedDateTime desc')
        select: List of properties to select
        skip: Number of items to skip (pagination)
        top: Max number of messages to return (max 100)
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

        # Build query parameters
        params = {}
        if filter:
            params["$filter"] = filter
        if orderby:
            params["$orderby"] = orderby
        if select:
            params["$select"] = ",".join(select)
        if skip is not None:
            params["$skip"] = skip
        if top is not None:
            if top > 100:
                top = 100  # Microsoft Graph max
            params["$top"] = top

        user = user_id if user_id else "me"
        folder_name = folder if folder else "inbox"
        endpoint = f"/{user}/mailFolders/{folder_name}/messages"

        result = client.get(endpoint, params=params if params else None)

        return {
            "successful": True,
            "data": result
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def send_draft(
    client,
    message_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Send an existing draft message. Use after creating a draft (via
    create_draft) when you want to deliver it to recipients immediately.

    Get message_id from create_draft (the returned 'id'), or from
    list_messages with folder='drafts' (pick the 'id' field of the
    draft you want to send).

    Args:
        client: The OutlookClient instance
        message_id: The ID of the draft message to send.
                    Get from create_draft or list_messages (folder='drafts').
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

        user = user_id if user_id else "me"

        # Verify the message is a draft first
        check_endpoint = f"/{user}/messages/{message_id}?$select=id,subject,isDraft"
        try:
            message_info = client.get(check_endpoint)
            if not message_info.get("isDraft", False):
                return {
                    "successful": False,
                    "data": {},
                    "error": "This message is not a draft. Only draft messages can be sent. Get a draft message_id from create_draft or list_messages (folder='drafts')."
                }
        except Exception as check_error:
            error_msg = str(check_error)
            if "404" in error_msg or "Not Found" in error_msg:
                return {
                    "successful": False,
                    "data": {},
                    "error": f"Message not found. The message_id '{message_id}' does not exist. Get a valid draft message_id from create_draft or list_messages (folder='drafts')."
                }
            # Continue and let the API handle it
            pass

        endpoint = f"/{user}/messages/{message_id}/send"

        # POST to send (returns 202 Accepted with no content on success)
        client.post(endpoint)

        subject = message_info.get("subject", "Unknown") if 'message_info' in dir() else "Unknown"

        return {
            "successful": True,
            "data": {
                "message": "Draft sent successfully",
                "sent_subject": subject
            }
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def send_email(
    client,
    subject: str,
    body: str,
    to_email: str,
    to_name: Optional[str] = None,
    cc_emails: Optional[List[str]] = None,
    bcc_emails: Optional[List[str]] = None,
    is_html: Optional[bool] = None,
    attachment: Optional[dict] = None,
    save_to_sent_items: Optional[bool] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Sends an email with subject, body, recipients, and an optional attachment via Microsoft Graph API.
    Attachments require a non-empty file with valid name and mimetype.
    
    Args:
        client: The OutlookClient instance
        subject: The subject of the email
        body: The body content of the email
        to_email: The primary recipient email address
        to_name: Optional name of the primary recipient
        cc_emails: Optional list of CC email addresses
        bcc_emails: Optional list of BCC email addresses
        is_html: Whether body is HTML (default: False)
        attachment: Optional attachment dict with name, contentType, contentBytes
        save_to_sent_items: Whether to save the email to Sent Items (default: True)
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
        
        # Build the message
        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": to_email,
                        "name": to_name if to_name else to_email
                    }
                }
            ]
        }
        
        # Add CC recipients
        if cc_emails:
            message["ccRecipients"] = [
                {"emailAddress": {"address": email}} for email in cc_emails
            ]
        
        # Add BCC recipients
        if bcc_emails:
            message["bccRecipients"] = [
                {"emailAddress": {"address": email}} for email in bcc_emails
            ]
        
        # Add attachment if provided
        if attachment:
            message["attachments"] = [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment.get("name"),
                    "contentType": attachment.get("contentType", "application/octet-stream"),
                    "contentBytes": attachment.get("contentBytes")
                }
            ]
        
        # Build the send mail payload
        send_data = {
            "message": message
        }
        
        if save_to_sent_items is not None:
            send_data["saveToSentItems"] = save_to_sent_items
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/sendMail"
        
        # Make the API call (sendMail returns no content on success)
        client.post(endpoint, json=send_data)
        
        return {
            "successful": True,
            "data": {"message": "Email sent successfully"}
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }