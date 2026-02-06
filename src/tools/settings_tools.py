"""
Microsoft Outlook Settings Tools
"""

from typing import Optional, List, Literal


def get_mailbox_settings(
    client,
    select: Optional[List[str]] = None,
    expand: Optional[List[str]] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieve mailbox settings.
    Use when you need to view settings such as automatic replies, time zone, 
    and working hours for the signed-in or specified user.
    
    Args:
        client: The OutlookClient instance
        select: Optional list of properties to select
        expand: Optional list of properties to expand
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
        if expand:
            params["$expand"] = ",".join(expand)
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/mailboxSettings"
        
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


def get_mail_delta(
    client,
    folder_id: Optional[str] = None,
    select: Optional[List[str]] = None,
    expand: Optional[List[str]] = None,
    filter: Optional[str] = None,
    orderby: Optional[List[str]] = None,
    search: Optional[str] = None,
    top: Optional[int] = None,
    skip: Optional[int] = None,
    count: Optional[bool] = None,
    delta_token: Optional[str] = None,
    skip_token: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieve incremental changes (delta) of messages in a mailbox.
    Use when syncing mailbox updates since last checkpoint.
    
    Args:
        client: The OutlookClient instance
        folder_id: Optional folder ID to get delta for specific folder
        select: Optional list of properties to select
        expand: Optional list of properties to expand
        filter: Optional OData filter expression
        orderby: Optional list of properties to order by
        search: Optional search query
        top: Optional number of items to return
        skip: Optional number of items to skip
        count: Whether to include count of items
        delta_token: Token from previous delta call for incremental sync
        skip_token: Token for pagination
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
        if expand:
            params["$expand"] = ",".join(expand)
        if filter:
            params["$filter"] = filter
        if orderby:
            params["$orderby"] = ",".join(orderby)
        if search:
            params["$search"] = search
        if top is not None:
            params["$top"] = top
        if skip is not None:
            params["$skip"] = skip
        if count is not None:
            params["$count"] = str(count).lower()
        if delta_token:
            params["$deltatoken"] = delta_token
        if skip_token:
            params["$skiptoken"] = skip_token
        
        # Determine the endpoint
        # Note: For personal accounts, delta queries require a specific folder
        user = user_id if user_id else "me"
        folder = folder_id if folder_id else "inbox"
        endpoint = f"/{user}/mailFolders/{folder}/messages/delta"
        
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


def get_mail_tips(
    client,
    EmailAddresses: List[str],
    MailTipsOptions: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieve mail tips such as automatic replies and mailbox full status.
    Use when you need to check recipient status before sending mail.
    
    Args:
        client: The OutlookClient instance
        EmailAddresses: List of email addresses to get mail tips for
        MailTipsOptions: Comma-separated mail tip options (e.g., "automaticReplies,mailboxFullStatus")
                        Valid options: automaticReplies, customMailTip, deliveryRestriction,
                        externalMemberCount, mailboxFullStatus, maxMessageSize, moderationStatus,
                        recipientScope, recipientSuggestions, totalMemberCount
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
        
        # Build the request payload
        payload = {
            "EmailAddresses": EmailAddresses,
            "MailTipsOptions": MailTipsOptions
        }
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/getMailTips"
        
        # Make the API call
        result = client.post(endpoint, json=payload)
        
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


def get_supported_languages(
    client,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieve supported languages in the user's mailbox.
    Use when you need to display or select from available mailbox languages.
    
    Args:
        client: The OutlookClient instance
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
        endpoint = f"/{user}/outlook/supportedLanguages"
        
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


def get_supported_time_zones(
    client,
    timeZoneStandard: Optional[Literal["Windows", "Iana"]] = None
) -> dict:
    """
    Retrieve supported time zones in the user's mailbox.
    Use when you need a list of time zones to display or choose from for event scheduling.
    
    Args:
        client: The OutlookClient instance
        timeZoneStandard: Optional time zone standard ("Windows" or "Iana")
    
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
        if timeZoneStandard:
            params["TimeZoneStandard"] = timeZoneStandard
        
        # Endpoint for supported time zones
        endpoint = "/me/outlook/supportedTimeZones"
        
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


def update_mailbox_settings(
    client,
    automaticRepliesSetting: Optional[dict] = None,
    language: Optional[dict] = None,
    timeZone: Optional[str] = None,
    workingHours: Optional[dict] = None
) -> dict:
    """
    Tool to update mailbox settings for the signed-in user.
    Use when you need to configure automatic replies, default time zone, language, or working hours.
    Example: schedule automatic replies for vacation.
    
    Args:
        client: The OutlookClient instance
        automaticRepliesSetting: Optional automatic replies setting object
            Example: {
                "status": "scheduled",  # "disabled", "alwaysEnabled", "scheduled"
                "externalAudience": "all",  # "none", "contactsOnly", "all"
                "scheduledStartDateTime": {"dateTime": "2026-02-10T00:00:00", "timeZone": "UTC"},
                "scheduledEndDateTime": {"dateTime": "2026-02-15T23:59:59", "timeZone": "UTC"},
                "internalReplyMessage": "I'm out of office...",
                "externalReplyMessage": "I'm out of office..."
            }
        language: Optional language object with locale and displayName
            Example: {"locale": "en-US", "displayName": "English (United States)"}
        timeZone: Optional time zone string (e.g., "Pacific Standard Time")
        workingHours: Optional working hours object
            Example: {
                "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
                "startTime": "09:00:00",
                "endTime": "17:00:00",
                "timeZone": {"name": "Pacific Standard Time"}
            }
    
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
        
        # Build the mailbox settings update payload
        settings_data = {}
        
        # Validate and add automaticRepliesSetting
        if automaticRepliesSetting is not None:
            if not isinstance(automaticRepliesSetting, dict) or len(automaticRepliesSetting) == 0:
                return {
                    "successful": False,
                    "data": {},
                    "error": "automaticRepliesSetting must be a non-empty dictionary. Example: {\"status\": \"scheduled\", \"externalAudience\": \"all\", \"scheduledStartDateTime\": {\"dateTime\": \"2026-02-10T00:00:00\", \"timeZone\": \"UTC\"}, \"scheduledEndDateTime\": {\"dateTime\": \"2026-02-15T23:59:59\", \"timeZone\": \"UTC\"}, \"internalReplyMessage\": \"I'm out of office...\", \"externalReplyMessage\": \"I'm out of office...\"}"
                }
            # Validate required status field
            if "status" not in automaticRepliesSetting:
                return {
                    "successful": False,
                    "data": {},
                    "error": "automaticRepliesSetting must include 'status' field. Valid values: 'disabled', 'alwaysEnabled', 'scheduled'"
                }
            settings_data["automaticRepliesSetting"] = automaticRepliesSetting
        
        # Validate and add language - must be non-empty dict with locale
        if language is not None:
            if not isinstance(language, dict) or len(language) == 0:
                return {
                    "successful": False,
                    "data": {},
                    "error": "language must be a non-empty dictionary with 'locale' and 'displayName' fields. Example: {\"locale\": \"en-US\", \"displayName\": \"English (United States)\"}. If you don't want to change language, omit this field entirely."
                }
            if "locale" not in language:
                return {
                    "successful": False,
                    "data": {},
                    "error": "language must include 'locale' field. Example: {\"locale\": \"en-US\", \"displayName\": \"English (United States)\"}"
                }
            settings_data["language"] = language
        
        # Validate and add timeZone - must be non-empty string
        if timeZone is not None:
            if not isinstance(timeZone, str) or not timeZone.strip():
                return {
                    "successful": False,
                    "data": {},
                    "error": "timeZone must be a non-empty string. Example: 'Pacific Standard Time' or 'UTC'. If you don't want to change time zone, omit this field entirely."
                }
            settings_data["timeZone"] = timeZone
        
        # Validate and add workingHours - must be non-empty dict
        if workingHours is not None:
            if not isinstance(workingHours, dict) or len(workingHours) == 0:
                return {
                    "successful": False,
                    "data": {},
                    "error": "workingHours must be a non-empty dictionary. Example: {\"daysOfWeek\": [\"monday\", \"tuesday\", \"wednesday\", \"thursday\", \"friday\"], \"startTime\": \"09:00:00\", \"endTime\": \"17:00:00\", \"timeZone\": {\"name\": \"Pacific Standard Time\"}}. If you don't want to change working hours, omit this field entirely."
                }
            # Validate required fields
            if "daysOfWeek" not in workingHours or "startTime" not in workingHours or "endTime" not in workingHours:
                return {
                    "successful": False,
                    "data": {},
                    "error": "workingHours must include 'daysOfWeek', 'startTime', and 'endTime' fields. Example: {\"daysOfWeek\": [\"monday\", \"tuesday\"], \"startTime\": \"09:00:00\", \"endTime\": \"17:00:00\", \"timeZone\": {\"name\": \"Pacific Standard Time\"}}"
                }
            settings_data["workingHours"] = workingHours
        
        # Ensure at least one field is provided
        if len(settings_data) == 0:
            return {
                "successful": False,
                "data": {},
                "error": "At least one setting must be provided. Provide automaticRepliesSetting, language, timeZone, or workingHours."
            }
        
        # Determine the endpoint
        endpoint = "/me/mailboxSettings"
        
        # Make the API call
        result = client.patch(endpoint, json=settings_data)
        
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