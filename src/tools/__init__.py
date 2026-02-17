# Outlook MCP Tools
from .mail_tools import (
    add_mail_attachment,
    batch_move_messages,
    batch_update_messages,
    create_draft,
    create_draft_reply,
    delete_message,
    forward_message,
    get_message,
    move_message,
    permanent_delete_message,
    query_emails,
    reply_email,
    search_messages,
    send_draft,
    send_email,
    update_email
)
from .calendar_tools import (
    add_event_attachment,
    create_calendar,
    create_event,
    decline_event,
    delete_event,
    find_meeting_times,
    get_calendar_view,
    get_event,
    get_schedule,
    update_calendar_event
)
from .contact_tools import (
    create_contact,
    create_contact_folder,
    delete_contact,
    get_contact,
    get_contact_folders,
    update_contact
)
from .rule_tools import create_email_rule, delete_email_rule, list_email_rules, update_email_rule
from .folder_tools import create_mail_folder, delete_mail_folder
from .category_tools import create_master_category, delete_master_category, get_master_categories
from .attachment_tools import create_attachment_upload_session, download_outlook_attachment
from .settings_tools import (
    get_mailbox_settings,
    get_mail_delta,
    get_mail_tips,
    get_supported_languages,
    get_supported_time_zones,
    update_mailbox_settings
)
from .list_tools import (
    list_calendars,
    list_child_mail_folders,
    list_contacts,
    list_event_attachments,
    list_events,
    list_mail_folder_messages,
    list_mail_folders,
    list_messages,
    list_outlook_attachments,
    list_reminders,
    list_files
)
from .user_tools import list_users
from .teams_tools import (
    list_chats,
    list_chat_messages,
    pin_message
)
from .profile_tools import get_profile

__all__ = [
    "add_mail_attachment",
    "batch_move_messages",
    "batch_update_messages",
    "create_attachment_upload_session",
    "create_draft",
    "create_draft_reply",
    "decline_event",
    "delete_email_rule",
    "delete_master_category",
    "delete_message",
    "find_meeting_times",
    "forward_message",
    "get_calendar_view",
    "get_message",
    "move_message",
    "reply_email",
    "search_messages",
    "send_email",
    "update_email",
    "add_event_attachment",
    "create_calendar",
    "create_event",
    "delete_event",
    "get_event",
    "get_schedule",
    "update_calendar_event",
    "create_contact",
    "create_contact_folder",
    "delete_contact",
    "get_contact",
    "get_contact_folders",
    "update_contact",
    "create_email_rule",
    "create_mail_folder",
    "delete_mail_folder",
    "create_master_category",
    "get_master_categories",
    "download_outlook_attachment",
    "get_mailbox_settings",
    "get_mail_delta",
    "get_mail_tips",
    "get_supported_languages",
    "get_supported_time_zones",
    "update_mailbox_settings",
    "list_calendars",
    "list_chat_messages",
    "list_chats",
    "list_child_mail_folders",
    "list_contacts",
    "list_email_rules",
    "list_event_attachments",
    "list_events",
    "list_mail_folder_messages",
    "list_mail_folders",
    "list_messages",
    "list_outlook_attachments",
    "list_reminders",
    "list_users",
    "list_files",
    "permanent_delete_message",
    "pin_message",
    "query_emails",
    "send_draft",
    "update_email_rule",
    "get_profile"
]
