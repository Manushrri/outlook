"""
Microsoft Outlook Category Tools
"""

from typing import Optional, Literal, List


def get_master_categories(
    client,
    select: Optional[List[str]] = None,
    filter: Optional[str] = None,
    orderby: Optional[List[str]] = None,
    top: Optional[int] = None,
    skip: Optional[int] = None,
    user_id: Optional[str] = None
) -> dict:
    """
    Retrieve the user's master category list.
    Use when you need to get all categories defined for the user.
    
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
        endpoint = f"/{user}/outlook/masterCategories"
        
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


def delete_master_category(
    client,
    category_id: str,
    user_id: Optional[str] = None
) -> dict:
    """
    Delete a category from the user's master category list.
    Use when removing unused or obsolete categories from the mailbox.

    Get category_id from get_master_categories (pick the 'id' field of the
    category you want to delete).

    Args:
        client: The OutlookClient instance
        category_id: The ID of the category to delete.
                     Get from get_master_categories (use the 'id' field).
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
        endpoint = f"/{user}/outlook/masterCategories/{category_id}"

        # DELETE returns 204 No Content on success
        client.delete(endpoint)

        return {
            "successful": True,
            "data": {"message": f"Category '{category_id}' deleted successfully"}
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def create_master_category(
    client,
    displayName: str,
    color: Optional[Literal[
        "preset0", "preset1", "preset2", "preset3", "preset4", "preset5",
        "preset6", "preset7", "preset8", "preset9", "preset10", "preset11",
        "preset12", "preset13", "preset14", "preset15", "preset16", "preset17",
        "preset18", "preset19", "preset20", "preset21", "preset22", "preset23",
        "preset24"
    ]] = None
) -> dict:
    """
    Create a new category in the user's master category list.
    Use after selecting a unique display name.
    
    Args:
        client: The OutlookClient instance
        displayName: The display name of the category
        color: Optional color preset (preset0 through preset24)
    
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
        
        # Build the category payload
        category_data = {
            "displayName": displayName
        }
        
        # Add optional fields if provided
        if color is not None:
            category_data["color"] = color
        
        # Endpoint for master categories
        endpoint = "/me/outlook/masterCategories"
        
        # Make the API call
        result = client.post(endpoint, json=category_data)
        
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

