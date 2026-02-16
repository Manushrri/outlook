"""
Microsoft Outlook Email Rule Tools
"""

from typing import Optional


def delete_email_rule(
    client,
    ruleId: str
) -> dict:
    """
    Delete an email rule from the user's inbox.

    Get ruleId from list_email_rules (pick the 'id' field of the rule you
    want to delete). If list_email_rules is not yet available, you can find
    rules via the Microsoft Graph endpoint /me/mailFolders/inbox/messageRules.

    Args:
        client: The OutlookClient instance
        ruleId: The ID of the email rule to delete.
                Get from list_email_rules (use the 'id' field).

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

        endpoint = f"/me/mailFolders/inbox/messageRules/{ruleId}"

        # DELETE returns 204 No Content on success
        client.delete(endpoint)

        return {
            "successful": True,
            "data": {"message": f"Email rule '{ruleId}' deleted successfully"}
        }

    except Exception as e:
        return {
            "successful": False,
            "data": {},
            "error": str(e)
        }


def list_email_rules(
    client,
    top: Optional[int] = None
) -> dict:
    """
    List all email rules from the user's inbox.
    Use when you need to see existing rules before creating, updating,
    or deleting them.

    The returned rules include their 'id' field which you can pass to
    delete_email_rule to remove a specific rule.

    Args:
        client: The OutlookClient instance
        top: Optional max number of rules to return

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

        endpoint = "/me/mailFolders/inbox/messageRules"

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


def update_email_rule(
    client,
    ruleId: str,
    displayName: Optional[str] = None,
    conditions: Optional[dict] = None,
    actions: Optional[dict] = None,
    isEnabled: Optional[bool] = None,
    sequence: Optional[int] = None
) -> dict:
    """
    Update an existing email rule. Provide only the fields you want to change.

    Get ruleId from list_email_rules (pick the 'id' field of the rule you
    want to update).

    Args:
        client: The OutlookClient instance
        ruleId: The ID of the email rule to update.
                Get from list_email_rules (use the 'id' field).
        displayName: Optional new display name for the rule
        conditions: Optional new conditions (replaces existing)
        actions: Optional new actions (replaces existing)
        isEnabled: Optional enable/disable the rule
        sequence: Optional new order of the rule in the list

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

        # Build the update payload — only include provided fields
        rule_data = {}
        if displayName is not None:
            rule_data["displayName"] = displayName
        if conditions is not None:
            rule_data["conditions"] = conditions
        if actions is not None:
            rule_data["actions"] = actions
        if isEnabled is not None:
            rule_data["isEnabled"] = isEnabled
        if sequence is not None:
            if sequence < 1:
                return {
                    "successful": False,
                    "data": {},
                    "error": "sequence must be a positive integer (1 or greater)."
                }
            rule_data["sequence"] = sequence

        if not rule_data:
            return {
                "successful": False,
                "data": {},
                "error": "At least one field (displayName, conditions, actions, isEnabled, sequence) must be provided to update."
            }

        endpoint = f"/me/mailFolders/inbox/messageRules/{ruleId}"

        result = client.patch(endpoint, json=rule_data)

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


def create_email_rule(
    client,
    displayName: str,
    conditions: dict,
    actions: dict,
    isEnabled: Optional[bool] = None,
    sequence: Optional[int] = None
) -> dict:
    """
    Create email rule filter with conditions and actions.
    
    Args:
        client: The OutlookClient instance
        displayName: The display name of the rule
        conditions: The conditions that trigger the rule
        actions: The actions to perform when conditions are met
        isEnabled: Whether the rule is enabled (default True)
        sequence: The order of the rule in the rule list
    
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
        
        # Validate conditions structure
        if not conditions or not isinstance(conditions, dict):
            return {
                "successful": False,
                "data": {},
                "error": "conditions must be a non-empty dictionary. Example: {\"fromAddresses\": [{\"emailAddress\": {\"address\": \"email@example.com\"}}]}"
            }
        
        # Validate actions structure
        if not actions or not isinstance(actions, dict):
            return {
                "successful": False,
                "data": {},
                "error": "actions must be a non-empty dictionary. Example: {\"delete\": true} or {\"moveToFolder\": \"folderId\"}"
            }
        
        # Validate that at least one condition field is present
        valid_condition_fields = [
            "fromAddresses", "sentToAddresses", "subjectContains", "bodyContains",
            "hasAttachments", "importance", "bodyOrSubjectContains", "categories",
            "flag", "fromAddressContains", "isApprovalRequest", "isAutomaticForward",
            "isAutomaticReply", "isEncrypted", "isMeetingRequest", "isMeetingResponse",
            "isNonDeliveryReport", "isPermissionControlled", "isReadReceipt", "isSigned",
            "isVoicemail", "messageActionFlag", "notSentToMe", "sentCcMe", "sentOnlyToMe",
            "sentToAddresses", "sentToMe", "sentToOrCcMe", "sensitivity", "withinSizeRange"
        ]
        has_valid_condition = any(field in conditions for field in valid_condition_fields)
        if not has_valid_condition:
            return {
                "successful": False,
                "data": {},
                "error": f"conditions must contain at least one valid field. Valid fields include: {', '.join(valid_condition_fields[:10])}... Use 'fromAddresses' for sender filtering."
            }
        
        # Validate that at least one action field is present
        valid_action_fields = [
            "assignCategories", "copyToFolder", "delete", "forwardAsAttachmentTo",
            "forwardTo", "markAsRead", "markImportance", "moveToFolder",
            "permanentDelete", "redirectTo", "stopProcessingRules"
        ]
        has_valid_action = any(field in actions for field in valid_action_fields)
        if not has_valid_action:
            return {
                "successful": False,
                "data": {},
                "error": f"actions must contain at least one valid field. Valid fields include: {', '.join(valid_action_fields)}. Use 'delete' for deletion or 'moveToFolder' with a folder ID."
            }
        
        # Build the rule payload - Microsoft Graph API expects messageRulePredicates and messageRuleActions
        rule_data = {
            "displayName": displayName,
            "conditions": conditions,  # This should be messageRulePredicates structure
            "actions": actions  # This should be messageRuleActions structure
        }
        
        # Add optional fields if provided
        if isEnabled is not None:
            rule_data["isEnabled"] = isEnabled
        # Sequence must be a positive integer (>= 1), cannot be 0
        if sequence is not None:
            if sequence < 1:
                return {
                    "successful": False,
                    "data": {},
                    "error": "sequence must be a positive integer (1 or greater). Cannot be 0 or negative."
                }
            rule_data["sequence"] = sequence
        
        # Endpoint for inbox rules
        endpoint = "/me/mailFolders/inbox/messageRules"
        
        # Make the API call
        result = client.post(endpoint, json=rule_data)
        
        return {
            "successful": True,
            "data": result
        }
        
    except Exception as e:
        error_msg = str(e)
        # Provide helpful guidance for common errors
        if "400" in error_msg or "Bad Request" in error_msg:
            error_msg += "\n\nCommon issues:\n"
            error_msg += "1. Ensure 'fromAddresses' uses format: [{\"emailAddress\": {\"address\": \"email@example.com\"}}]\n"
            error_msg += "2. For 'moveToFolder' or 'copyToFolder', get folder ID using: list_mail_folders\n"
            error_msg += "3. 'delete' action should be: {\"delete\": true}\n"
            error_msg += "4. Check that conditions and actions contain at least one valid field\n"
            error_msg += "5. Ensure the email address in fromAddresses is valid"
        return {
            "successful": False,
            "data": {},
            "error": error_msg
        }

