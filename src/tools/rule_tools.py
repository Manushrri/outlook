"""
Microsoft Outlook Email Rule Tools
"""

from typing import Optional


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

