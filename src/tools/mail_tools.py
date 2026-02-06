"""
Microsoft Outlook Mail Tools
"""

import base64
from typing import Optional, List


def add_mail_attachment(
    client,
    message_id: str,
    name: str,
    odata_type: str,
    contentBytes: str,
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
        contentBytes: Base64-encoded content of the file
        contentId: Optional content ID for inline attachments
        contentLocation: Optional content location URL
        contentType: Optional MIME type of the attachment
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
        
        # Validate that contentBytes is not empty
        if not contentBytes or not contentBytes.strip():
            return {
                "successful": False,
                "data": {},
                "error": "contentBytes cannot be empty. Provide base64-encoded file content."
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
            error_msg += "4. contentBytes must be valid base64-encoded content\n"
            error_msg += "5. File size must be less than 3 MB\n"
            error_msg += "6. Do NOT include empty objects {} for optional fields like 'item'"
        return {
            "successful": False,
            "data": {},
            "error": error_msg
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