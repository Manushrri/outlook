"""
Microsoft Outlook Contact Tools
"""

from typing import Optional, List


def create_contact(
    client,
    givenName: Optional[str] = None,
    surname: Optional[str] = None,
    displayName: Optional[str] = None,
    emailAddresses: Optional[List[dict]] = None,
    businessPhones: Optional[List[str]] = None,
    mobilePhone: Optional[str] = None,
    homePhone: Optional[str] = None,
    companyName: Optional[str] = None,
    department: Optional[str] = None,
    jobTitle: Optional[str] = None,
    officeLocation: Optional[str] = None,
    birthday: Optional[str] = None,
    categories: Optional[List[str]] = None,
    notes: Optional[str] = None,
    folderId: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Creates a new contact in a Microsoft Outlook user's contacts folder.
    
    Args:
        client: The OutlookClient instance
        givenName: First name
        surname: Last name
        displayName: Display name
        emailAddresses: List of email address dicts [{"address": "email@example.com", "name": "Name"}]
        businessPhones: List of business phone numbers
        mobilePhone: Mobile phone number
        homePhone: Home phone number
        companyName: Company name
        department: Department
        jobTitle: Job title
        officeLocation: Office location
        birthday: Birthday (ISO 8601 date)
        categories: List of categories
        notes: Notes about the contact
        folderId: Optional contact folder ID to store the contact in a specific folder. Get folder_id from get_contact_folders tool. If not provided, contact is created in the default contacts folder.
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
        
        # Build the contact payload
        contact_data = {}
        
        if givenName is not None:
            contact_data["givenName"] = givenName
        if surname is not None:
            contact_data["surname"] = surname
        if displayName is not None:
            contact_data["displayName"] = displayName
        if emailAddresses is not None:
            contact_data["emailAddresses"] = emailAddresses
        if businessPhones is not None:
            contact_data["businessPhones"] = businessPhones
        if mobilePhone is not None:
            contact_data["mobilePhone"] = mobilePhone
        if homePhone is not None:
            contact_data["homePhones"] = [homePhone]
        if companyName is not None:
            contact_data["companyName"] = companyName
        if department is not None:
            contact_data["department"] = department
        if jobTitle is not None:
            contact_data["jobTitle"] = jobTitle
        if officeLocation is not None:
            contact_data["officeLocation"] = officeLocation
        if birthday is not None:
            contact_data["birthday"] = birthday
        if categories is not None:
            contact_data["categories"] = categories
        if notes is not None:
            contact_data["personalNotes"] = notes
        
        # Determine the endpoint based on whether folderId is provided
        user = user_id if user_id else "me"
        if folderId:
            endpoint = f"/{user}/contactFolders/{folderId}/contacts"
        else:
            endpoint = f"/{user}/contacts"
        
        # Make the API call
        result = client.post(endpoint, json=contact_data)
        
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


def get_contact(
    client,
    contact_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieves a specific Outlook contact by its contact ID.
    
    Args:
        client: The OutlookClient instance
        contact_id: The ID of the contact to retrieve
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
        endpoint = f"/{user}/contacts/{contact_id}"
        
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


def get_contact_folders(
    client,
    select: Optional[List[str]] = None,
    filter: Optional[str] = None,
    orderby: Optional[List[str]] = None,
    top: Optional[int] = None,
    skip: Optional[int] = None,
    expand: Optional[List[str]] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieves a list of contact folders in the signed-in user's mailbox.
    Use after authentication when you need to browse or select among contact folders.
    
    Args:
        client: The OutlookClient instance
        select: Optional list of properties to select
        filter: Optional OData filter expression
        orderby: Optional list of properties to order by
        top: Optional number of items to return
        skip: Optional number of items to skip
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
        if filter:
            params["$filter"] = filter
        if orderby:
            params["$orderby"] = ",".join(orderby)
        if top is not None:
            params["$top"] = top
        if skip is not None:
            params["$skip"] = skip
        if expand:
            params["$expand"] = ",".join(expand)
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/contactFolders"
        
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


def delete_contact(
    client,
    contact_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Permanently deletes an existing contact from the user's Outlook contacts.
    
    Args:
        client: The OutlookClient instance
        contact_id: The ID of the contact to delete
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
        endpoint = f"/{user}/contacts/{contact_id}"
        
        # Make the API call
        client.delete(endpoint)
        
        return {
            "successful": True,
            "data": {"message": "Contact deleted successfully"}
        }
        
    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def update_contact(
    client,
    contact_id: str,
    givenName: Optional[str] = None,
    surname: Optional[str] = None,
    displayName: Optional[str] = None,
    emailAddresses: Optional[List[dict]] = None,
    businessPhones: Optional[List[str]] = None,
    mobilePhone: Optional[str] = None,
    homePhones: Optional[List[str]] = None,
    companyName: Optional[str] = None,
    department: Optional[str] = None,
    jobTitle: Optional[str] = None,
    officeLocation: Optional[str] = None,
    birthday: Optional[str] = None,
    categories: Optional[List[str]] = None,
    notes: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Updates an existing Outlook contact, identified by contact_id for the specified user_id,
    requiring at least one other field to be modified.
    
    Args:
        client: The OutlookClient instance
        contact_id: The ID of the contact to update
        givenName: Optional first name
        surname: Optional last name
        displayName: Optional display name
        emailAddresses: Optional list of email address dicts [{"address": "email@example.com", "name": "Name"}]
        businessPhones: Optional list of business phone numbers
        mobilePhone: Optional mobile phone number
        homePhones: Optional list of home phone numbers
        companyName: Optional company name
        department: Optional department
        jobTitle: Optional job title
        officeLocation: Optional office location
        birthday: Optional birthday (ISO 8601 date)
        categories: Optional list of categories
        notes: Optional notes about the contact
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
        
        # Build the contact update payload
        contact_data = {}
        
        if givenName is not None:
            contact_data["givenName"] = givenName
        if surname is not None:
            contact_data["surname"] = surname
        if displayName is not None:
            contact_data["displayName"] = displayName
        if emailAddresses is not None:
            contact_data["emailAddresses"] = emailAddresses
        if businessPhones is not None:
            contact_data["businessPhones"] = businessPhones
        if mobilePhone is not None:
            contact_data["mobilePhone"] = mobilePhone
        if homePhones is not None:
            contact_data["homePhones"] = homePhones
        if companyName is not None:
            contact_data["companyName"] = companyName
        if department is not None:
            contact_data["department"] = department
        if jobTitle is not None:
            contact_data["jobTitle"] = jobTitle
        if officeLocation is not None:
            contact_data["officeLocation"] = officeLocation
        if birthday is not None:
            contact_data["birthday"] = birthday
        if categories is not None:
            contact_data["categories"] = categories
        if notes is not None:
            contact_data["personalNotes"] = notes
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/contacts/{contact_id}"
        
        # Make the API call
        result = client.patch(endpoint, json=contact_data)
        
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


def create_contact_folder(
    client,
    displayName: str,
    parentFolderId: Optional[str] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Create a new contact folder in the user's mailbox.
    Use when needing to organize contacts into custom folders.
    
    Args:
        client: The OutlookClient instance
        displayName: The display name of the contact folder
        parentFolderId: Optional parent folder ID to create subfolder
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
        
        # Build the folder payload
        folder_data = {
            "displayName": displayName
        }
        
        # Add optional fields if provided
        if parentFolderId is not None:
            folder_data["parentFolderId"] = parentFolderId
        
        # Determine the endpoint
        user = user_id if user_id else "me"
        endpoint = f"/{user}/contactFolders"
        
        # Make the API call
        result = client.post(endpoint, json=folder_data)
        
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

