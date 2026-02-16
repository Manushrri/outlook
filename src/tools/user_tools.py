"""
Microsoft Outlook User Tools
"""

from typing import Optional, List


def list_users(
    client,
    filter: Optional[str] = None,
    select: Optional[List[str]] = None,
    skip: Optional[int] = None,
    top: Optional[int] = None
) -> dict:
    """
    List users in Microsoft Entra ID. Use when you need to retrieve a
    paginated list of users, optionally filtering or selecting specific
    properties.

    Common select values: 'displayName', 'mail', 'userPrincipalName', 'id',
    'jobTitle', 'department', 'officeLocation'.

    Common filter examples:
      - "startswith(displayName, 'John')"
      - "mail eq 'john@example.com'"
      - "department eq 'Engineering'"

    Note: This requires User.Read.All or User.ReadBasic.All permission
    for listing other users. For the signed-in user only, use get_profile.

    Args:
        client: The OutlookClient instance
        filter: Optional OData filter expression
        select: Optional list of properties to select
        skip: Optional number of items to skip
        top: Optional max number of users to return

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
        if select:
            params["$select"] = ",".join(select)
        if skip is not None:
            params["$skip"] = skip
        if top is not None:
            params["$top"] = top

        endpoint = "/users"

        result = client.get(endpoint, params=params if params else None)

        return {
            "successful": True,
            "data": result
        }

    except Exception as e:
        error_msg = str(e)
        # Handle non-JSON responses (common with personal accounts)
        if "Expecting" in error_msg and "delimiter" in error_msg:
            error_msg = (
                "This endpoint requires a work/school (organizational) account "
                "with User.Read.All or User.ReadBasic.All permission. "
                "It is not available for personal Microsoft accounts (outlook.com, hotmail.com, live.com). "
                "For the signed-in user's own profile, use get_profile instead."
            )
        return {
            "successful": False,
            "data": {},
            "error": error_msg
        }

