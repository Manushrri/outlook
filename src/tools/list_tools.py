"""
Microsoft Outlook List Tools
"""

from typing import Optional, List


def list_calendars(
    client,
    select: Optional[List[str]] = None,
    filter: Optional[str] = None,
    orderby: Optional[List[str]] = None,
    top: Optional[int] = None,
    skip: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    List calendars in the signed-in user's mailbox.
    Use when you need to retrieve calendars with optional OData queries.
    
    Args:
        client: The OutlookClient instance
        select: Optional list of properties to select
        filter: Optional OData filter expression
        orderby: Optional list of properties to order by
        top: Optional number of items to return
        skip: Optional number of items to skip
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
            params["$select"] = ",".join(select)
        if filter:
            params["$filter"] = filter
        if orderby:
            params["$orderby"] = ",".join(orderby)
        if top is not None:
            params["$top"] = top
        if skip is not None:
            params["$skip"] = skip
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/calendars"
        
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


def list_event_attachments(
    client,
    event_id: str,
    select: Optional[List[str]] = None,
    filter: Optional[str] = None,
    orderby: Optional[List[str]] = None,
    top: Optional[int] = None,
    skip: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    List attachments for a specific Outlook calendar event.
    Use when you have an event ID and need to view its attachments.
    
    Args:
        client: The OutlookClient instance
        event_id: The ID of the event
        select: Optional list of properties to select
        filter: Optional OData filter expression
        orderby: Optional list of properties to order by
        top: Optional number of items to return
        skip: Optional number of items to skip
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
            params["$select"] = ",".join(select)
        if filter:
            params["$filter"] = filter
        if orderby:
            params["$orderby"] = ",".join(orderby)
        if top is not None:
            params["$top"] = top
        if skip is not None:
            params["$skip"] = skip
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/events/{event_id}/attachments"
        
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


def list_outlook_attachments(
    client,
    message_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Lists metadata (like name, size, and type, but not contentBytes) 
    for all attachments of a specified Outlook email message.
    
    Args:
        client: The OutlookClient instance
        message_id: The ID of the message
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
        endpoint = f"/{user}/messages/{message_id}/attachments"
        
        # Select metadata fields only (exclude contentBytes)
        params = {
            "$select": "id,name,contentType,size,isInline,lastModifiedDateTime"
        }
        
        # Make the API call
        result = client.get(endpoint, params=params)
        
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


def list_reminders(
    client,
    startDateTime: str,
    endDateTime: str,
    userId: Optional[str] = None
) -> dict:
    """
    Retrieve reminders for events occurring within a specified time range.
    Use when you need to see upcoming reminders between two datetimes.
    
    Args:
        client: The OutlookClient instance
        startDateTime: Start of the time range (ISO 8601 format)
        endDateTime: End of the time range (ISO 8601 format)
        userId: Optional user ID (defaults to 'me')
    
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
        user = userId if userId else "me"
        endpoint = f"/{user}/reminderView(startDateTime='{startDateTime}',endDateTime='{endDateTime}')"
        
        # Make the API call
        result = client.get(endpoint)
        
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


def list_child_mail_folders(
    client,
    parent_folder_id: str,
    filter: Optional[str] = None,
    include_hidden_folders: Optional[bool] = None,
    select: Optional[List[str]] = None,
    top: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    List subfolders (childFolders) under a specified Outlook mail folder.
    Use when navigating nested folder hierarchies or checking if a folder
    has subfolders.

    Get parent_folder_id from list_mail_folders (pick the 'id' field of
    the parent folder you want to explore). You can also use well-known
    folder names like 'inbox', 'drafts', 'sentitems', 'deleteditems'.

    Args:
        client: The OutlookClient instance
        parent_folder_id: The ID of the parent mail folder.
                          Get from list_mail_folders (use the 'id' field)
                          or a well-known name like 'inbox'.
        filter: Optional OData filter expression
        include_hidden_folders: Whether to include hidden folders
        select: Optional list of properties to select
        top: Optional max number of folders to return
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
        if include_hidden_folders:
            params["includeHiddenFolders"] = "true"
        if select:
            params["$select"] = ",".join(select)
        if top is not None:
            params["$top"] = top

        user = user_id if user_id else "me"
        endpoint = f"/{user}/mailFolders/{parent_folder_id}/childFolders"

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


def list_mail_folder_messages(
    client,
    mail_folder_id: str,
    filter: Optional[str] = None,
    orderby: Optional[str] = None,
    select: Optional[List[str]] = None,
    skip: Optional[int] = None,
    top: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    List messages from a specific mail folder including subfolders.
    Use when you need to retrieve messages from a particular folder or
    subfolder by its ID or well-known name.

    Get mail_folder_id from list_mail_folders (pick the 'id' field of the
    folder you want to read) or list_child_mail_folders for subfolders.
    You can also use well-known names like 'inbox', 'drafts', 'sentitems',
    'deleteditems'.

    Args:
        client: The OutlookClient instance
        mail_folder_id: The ID of the mail folder.
                        Get from list_mail_folders or list_child_mail_folders.
        filter: Optional OData filter expression
        orderby: Optional property to order by (e.g. 'receivedDateTime desc')
        select: Optional list of properties to select
        skip: Optional number of items to skip
        top: Optional max number of messages to return
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
            params["$top"] = top

        user = user_id if user_id else "me"
        endpoint = f"/{user}/mailFolders/{mail_folder_id}/messages"

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


def list_contacts(
    client,
    contact_folder_id: Optional[str] = None,
    filter: Optional[str] = None,
    orderby: Optional[List[str]] = None,
    select: Optional[List[str]] = None,
    top: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieves a user's Microsoft Outlook contacts, from the default or a specified contact folder.
    
    Args:
        client: The OutlookClient instance
        contact_folder_id: Optional contact folder ID to retrieve contacts from a specific folder
        filter: Optional OData filter expression
        orderby: Optional list of properties to order by
        select: Optional list of properties to select
        top: Optional number of items to return
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
            params["$orderby"] = ",".join(orderby)
        if select:
            params["$select"] = ",".join(select)
        if top is not None:
            params["$top"] = top
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        if contact_folder_id:
            endpoint = f"/{user}/contactFolders/{contact_folder_id}/contacts"
        else:
            endpoint = f"/{user}/contacts"
        
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


def list_events(
    client,
    expand_recurring_events: Optional[bool] = None,
    filter: Optional[str] = None,
    orderby: Optional[List[str]] = None,
    select: Optional[List[str]] = None,
    skip: Optional[int] = None,
    timezone: Optional[str] = None,
    top: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieves events from a user's Outlook calendar via Microsoft Graph API,
    supporting pagination, filtering, property selection, sorting, and timezone specification.
    
    Args:
        client: The OutlookClient instance
        expand_recurring_events: If True, uses calendarView to expand recurring events
        filter: Optional OData filter expression
        orderby: Optional list of properties to order by
        select: Optional list of properties to select
        skip: Optional number of items to skip
        timezone: Optional timezone for the response (e.g., 'Pacific Standard Time')
        top: Optional number of items to return
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
            params["$orderby"] = ",".join(orderby)
        if select:
            params["$select"] = ",".join(select)
        if skip is not None:
            params["$skip"] = skip
        if top is not None:
            params["$top"] = top
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/events"
        
        # Build custom headers for timezone
        headers = {}
        if timezone:
            headers["Prefer"] = f'outlook.timezone="{timezone}"'
        
        # Make the API call
        result = client.get(endpoint, params=params if params else None, headers=headers if headers else None)
        
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


def list_mail_folders(
    client,
    include_hidden_folders: Optional[bool] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Lists a user's top-level mail folders. Use when you need folders like Inbox, Drafts, Sent Items.
    Set include_hidden_folders=True to include hidden folders.
    
    Args:
        client: The OutlookClient instance
        include_hidden_folders: Whether to include hidden folders
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
        if include_hidden_folders:
            params["includeHiddenFolders"] = "true"
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/mailFolders"
        
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


def list_messages(
    client,
    folder: Optional[str] = None,
    categories: Optional[List[str]] = None,
    conversationId: Optional[str] = None,
    from_address: Optional[str] = None,
    has_attachments: Optional[bool] = None,
    importance: Optional[str] = None,
    is_read: Optional[bool] = None,
    orderby: Optional[List[str]] = None,
    received_date_time_ge: Optional[str] = None,
    received_date_time_gt: Optional[str] = None,
    received_date_time_le: Optional[str] = None,
    received_date_time_lt: Optional[str] = None,
    select: Optional[List[str]] = None,
    sent_date_time_gt: Optional[str] = None,
    sent_date_time_lt: Optional[str] = None,
    skip: Optional[int] = None,
    subject: Optional[str] = None,
    subject_contains: Optional[str] = None,
    subject_endswith: Optional[str] = None,
    subject_startswith: Optional[str] = None,
    top: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieves a list of email messages from a specified mail folder in an Outlook mailbox,
    with options for filtering (including by conversationId to get all messages in a thread),
    pagination, and sorting. Ensure 'user_id' and 'folder' are valid, and all date/time 
    strings are in ISO 8601 format.
    
    Args:
        client: The OutlookClient instance
        folder: Mail folder ID or well-known name (inbox, drafts, sentitems, deleteditems)
        categories: Optional list of categories to filter by
        conversationId: Optional conversation ID to get all messages in a thread
        from_address: Optional sender email address to filter by
        has_attachments: Optional filter for messages with attachments
        importance: Optional importance filter (low, normal, high)
        is_read: Optional filter for read/unread messages
        orderby: Optional list of properties to order by
        received_date_time_ge: Optional filter for receivedDateTime >= value (ISO 8601)
        received_date_time_gt: Optional filter for receivedDateTime > value (ISO 8601)
        received_date_time_le: Optional filter for receivedDateTime <= value (ISO 8601)
        received_date_time_lt: Optional filter for receivedDateTime < value (ISO 8601)
        select: Optional list of properties to select
        sent_date_time_gt: Optional filter for sentDateTime > value (ISO 8601)
        sent_date_time_lt: Optional filter for sentDateTime < value (ISO 8601)
        skip: Optional number of items to skip
        subject: Optional exact subject to filter by
        subject_contains: Optional substring to search in subject
        subject_endswith: Optional suffix to search in subject
        subject_startswith: Optional prefix to search in subject
        top: Optional number of items to return
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
        
        # Build filter expressions
        filter_parts = []
        
        if conversationId:
            filter_parts.append(f"conversationId eq '{conversationId}'")
        if from_address:
            filter_parts.append(f"from/emailAddress/address eq '{from_address}'")
        if has_attachments is not None:
            filter_parts.append(f"hasAttachments eq {str(has_attachments).lower()}")
        if importance:
            filter_parts.append(f"importance eq '{importance}'")
        if is_read is not None:
            filter_parts.append(f"isRead eq {str(is_read).lower()}")
        if received_date_time_ge:
            filter_parts.append(f"receivedDateTime ge {received_date_time_ge}")
        if received_date_time_gt:
            filter_parts.append(f"receivedDateTime gt {received_date_time_gt}")
        if received_date_time_le:
            filter_parts.append(f"receivedDateTime le {received_date_time_le}")
        if received_date_time_lt:
            filter_parts.append(f"receivedDateTime lt {received_date_time_lt}")
        if sent_date_time_gt:
            filter_parts.append(f"sentDateTime gt {sent_date_time_gt}")
        if sent_date_time_lt:
            filter_parts.append(f"sentDateTime lt {sent_date_time_lt}")
        if subject:
            filter_parts.append(f"subject eq '{subject}'")
        if subject_contains:
            filter_parts.append(f"contains(subject, '{subject_contains}')")
        if subject_startswith:
            filter_parts.append(f"startswith(subject, '{subject_startswith}')")
        if subject_endswith:
            filter_parts.append(f"endswith(subject, '{subject_endswith}')")
        if categories:
            # Filter by categories using 'any' lambda
            cat_filters = [f"categories/any(c:c eq '{cat}')" for cat in categories]
            filter_parts.extend(cat_filters)
        
        # Build query parameters
        params = {}
        if filter_parts:
            params["$filter"] = " and ".join(filter_parts)
        if orderby:
            params["$orderby"] = ",".join(orderby)
        if select:
            params["$select"] = ",".join(select)
        if skip is not None:
            params["$skip"] = skip
        if top is not None:
            params["$top"] = top
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        if folder:
            endpoint = f"/{user}/mailFolders/{folder}/messages"
        else:
            endpoint = f"/{user}/messages"
        
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
