"""
Microsoft Teams Chat Tools
"""

from typing import Optional


def list_chats(
    client,
    top: Optional[int] = None,
    filter: Optional[str] = None,
    orderby: Optional[str] = None,
    expand: Optional[str] = None
) -> dict:
    """
    List Teams chats for the signed-in user. Use when you need chat IDs
    and topics to select a chat for further actions like list_chat_messages.

    Args:
        client: The OutlookClient instance
        top: Optional max number of chats to return.
        filter: Optional OData filter expression.
        orderby: Optional property to order by.
        expand: Optional property to expand (e.g. 'members', 'lastMessagePreview').

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

        params = {}
        if top is not None:
            params["$top"] = top
        if filter:
            params["$filter"] = filter
        if orderby:
            params["$orderby"] = orderby
        if expand:
            params["$expand"] = expand

        endpoint = "/me/chats"

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


def pin_message(
    client,
    chat_id: str,
    message_url: str
) -> dict:
    """
    Pin a message in a Teams chat. Use when you want to mark an important
    message for quick access.

    Get chat_id from list_chats (pick the 'id' field of the chat).
    Get message_url by constructing it from chat_id and message_id:
      https://graph.microsoft.com/v1.0/chats/{chat_id}/messages/{message_id}
    Get message_id from list_chat_messages (pick the 'id' field of the
    message you want to pin).

    Args:
        client: The OutlookClient instance
        chat_id: The ID of the chat. Get from list_chats.
        message_url: The full URL of the message to pin.
                     Format: https://graph.microsoft.com/v1.0/chats/{chat_id}/messages/{message_id}

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

        endpoint = f"/chats/{chat_id}/pinnedMessages"

        pin_data = {
            "message@odata.bind": message_url
        }

        result = client.post(endpoint, json=pin_data)

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


def list_chat_messages(
    client,
    chat_id: str,
    top: Optional[int] = None,
    filter: Optional[str] = None,
    orderby: Optional[str] = None
) -> dict:
    """
    List messages in a Teams chat. Use when you need message IDs to select
    a specific message for further actions.

    Get chat_id from list_chats (pick the 'id' field of the chat you want
    to read messages from).

    Args:
        client: The OutlookClient instance
        chat_id: The ID of the chat. Get from list_chats.
        top: Optional max number of messages to return.
        filter: Optional OData filter expression.
        orderby: Optional property to order by.

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

        params = {}
        if top is not None:
            params["$top"] = top
        if filter:
            params["$filter"] = filter
        if orderby:
            params["$orderby"] = orderby

        endpoint = f"/me/chats/{chat_id}/messages"

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

