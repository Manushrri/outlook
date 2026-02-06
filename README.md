# Microsoft Outlook MCP Server

A modular MCP (Model Context Protocol) server for Microsoft Outlook integration. Provides 43 tools for managing emails, calendars, contacts, and mailbox settings via Microsoft Graph API.

## Features

- 📧 **Email Management**: Send, reply, search, move, and manage drafts
- 📅 **Calendar**: Create, update, delete events; manage calendars and reminders
- 👥 **Contacts**: Create, update, delete contacts and contact folders
- 📎 **Attachments**: Download and manage email and event attachments
- 📊 **Mailbox Settings**: Configure automatic replies, time zones, working hours
- 🔐 **OAuth2**: Automatic token storage and refresh (device code flow)
- 🧩 **Modular**: Clean separation of tools by category

## Project Structure

```
microsoft-outlook/
├── src/                    # Source code
│   ├── __init__.py         # Package marker
│   ├── main.py             # FastMCP server (main entry point)
│   ├── client.py           # Outlook OAuth/MSAL client
│   ├── config.py           # Configuration settings
│   └── tools/              # Modular tool implementations
│       ├── __init__.py     # Tools package marker
│       ├── mail_tools.py   # 10 mail tools
│       ├── list_tools.py   # 8 list tools
│       ├── calendar_tools.py # 9 calendar tools
│       ├── contact_tools.py  # 7 contact tools
│       ├── attachment_tools.py # 1 attachment tool
│       ├── folder_tools.py    # 2 folder tools
│       ├── category_tools.py  # 2 category tools
│       ├── profile_tools.py   # 1 profile tool
│       ├── rule_tools.py      # 1 rule tool
│       └── settings_tools.py # 6 settings tools
├── run_server.py           # Convenience wrapper for src.main
├── test_auth.py            # Authentication bootstrap (REQUIRED - run once)
├── tools_manifest.json     # Tool definitions for dynamic loading
├── .token_cache.json       # Saved tokens (auto-generated)
└── .env                    # Your credentials (create this)
```

## Quick Start

### 1. Install Dependencies

```bash
# Create virtual environment
python -m venv .venv

# Activate virtual environment
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### 2. Create Azure AD Application

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **New registration**
4. Configure:
   - **Name**: `Outlook MCP Server` (or any name)
   - **Supported account types**: Choose based on your needs
   - **Redirect URI**: Optional for device code flow, but you can add:
     - Platform: **Mobile and desktop applications**
     - URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
5. Click **Register**

### 3. Configure API Permissions

1. In your app registration, go to **API permissions**
2. Click **Add a permission** → **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:
   - `User.Read` - Sign in and read user profile
   - `Mail.Read` - Read user mail
   - `Mail.ReadWrite` - Read and write user mail
   - `Mail.Send` - Send mail as user
   - `Calendars.Read` - Read user calendars
   - `Calendars.ReadWrite` - Read and write user calendars
   - `Contacts.Read` - Read user contacts
   - `Contacts.ReadWrite` - Read and write user contacts
   - `MailboxSettings.Read` - Read mailbox settings
   - `MailboxSettings.ReadWrite` - Read and write mailbox settings
   - `offline_access` - Maintain access to data (for refresh tokens)
4. Click **Grant admin consent** (if you're an admin)

### 4. Get Client ID and Client Secret

1. **Get Client ID:**
   - In your app registration, go to **Overview**
   - Copy the **Application (client) ID** → this is your `OUTLOOK_CLIENT_ID`

2. **Create and Get Client Secret:**
   - Go to **Certificates & secrets**
   - Click **New client secret**
   - Add a description and choose expiration
   - Click **Add**
   - **IMPORTANT**: Copy the **Value** (not the Secret ID) - you will **NOT** see it again!
   - This **Value** is your `OUTLOOK_CLIENT_SECRET`

### 5. Configure Environment

Create `.env` file:

```env
# REQUIRED - Application (client) ID from Azure Portal Overview
OUTLOOK_CLIENT_ID=your-application-client-id-here

# REQUIRED - Client secret Value (not Secret ID) from Certificates & secrets
OUTLOOK_CLIENT_SECRET=your-client-secret-value-here

# Optional - Default is already set
OUTLOOK_REDIRECT_URI=https://login.microsoftonline.com/common/oauth2/nativeclient
```

**⚠️ Important:** When copying the client secret, make sure you copy the **Value** column, not the **Secret ID**. The Value is only shown once when you create the secret!

### 6. Run Authentication (REQUIRED - Run Once)

```bash
python test_auth.py
```

This will:
- Start device code flow
- Show you a code and URL (`https://microsoft.com/devicelogin`)
- You sign in and enter the code
- Accept permissions once
- Save tokens to `.token_cache.json`

### 7. Run the Server

```bash
# Using uv with wrapper (recommended)
uv run run_server.py

# Or using python with wrapper
python run_server.py

# Or directly with src/main.py
uv run src/main.py
# or
python -m src.main
```

**Note:** `run_server.py` is a convenience wrapper that calls `src.main.main()`. Both methods work identically.

---

## Complete Token & Authentication Guide

### What You Need to Know

**You DON'T need to manually manage tokens!** Everything is auto-handled after running `test_auth.py` once.

### The Complete OAuth Flow

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                    STEP 1: Run test_auth.py                                │
│                    $ python test_auth.py                                   │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  STEP 2: Device Code Flow Initiated                                        │
│  - MSAL initiates device code flow                                         │
│  - Terminal shows:                                                          │
│    "To sign in, use a web browser to open the page                          │
│     https://microsoft.com/devicelogin                                       │
│     and enter the code XXXX-XXXX to authenticate."                          │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  STEP 3: User Authentication                                               │
│  - User opens https://microsoft.com/devicelogin                             │
│  - User enters the code                                                    │
│  - User signs in with Microsoft account                                    │
│  - Microsoft shows consent screen for all requested scopes                  │
│  - User accepts permissions                                                 │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  STEP 4: Exchange Device Code for Access Token                             │
│                                                                             │
│  POST https://login.microsoftonline.com/common/oauth2/v2.0/token           │
│    grant_type=urn:ietf:params:oauth:grant-type:device_code                 │
│    &client_id={client_id}                                                  │
│    &device_code={device_code}                                              │
│                                                                             │
│  Returns:                                                                    │
│  {                                                                          │
│    "access_token": "eyJ0eXAi...",        <- Short-lived (1 hour)            │
│    "refresh_token": "0.AXcA...",        <- Long-lived (90 days)            │
│    "expires_in": 3600,                   <- 1 hour                          │
│    "token_type": "Bearer"                                                 │
│  }                                                                          │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  STEP 5: Token Saved to .token_cache.json                                   │
│                                                                             │
│  MSAL serializes token cache:                                               │
│  {                                                                          │
│    "AccessToken": {                                                         │
│      "secret": "eyJ0eXAi...",            <- access_token                    │
│      "expires_on": "1769688710"         <- Expiry timestamp                │
│    },                                                                       │
│    "RefreshToken": {                                                        │
│      "secret": "0.AXcA...",             <- refresh_token                   │
│    },                                                                       │
│    "Account": {                                                             │
│      "username": "user@example.com"                                         │
│    }                                                                        │
│  }                                                                          │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Understanding the Tokens

| Token | What It Is | Valid For | How It's Obtained |
|-------|-----------|-----------|-------------------|
| `access_token` | Short-lived access token | 1 hour | Device code flow → token endpoint |
| `refresh_token` | Long-lived refresh token | 90 days | Included in initial token response |
| `expires_on` | Token expiry timestamp | - | From API response |

### How Token Refresh Works

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                        TOKEN LIFECYCLE                                      │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  Hour 0    Hour 0.5   Hour 0.9      Hour 1                                  │
│    │         │         │             │                                      │
│    │◄────────│─────────│─────────────┤                                      │
│    │         │         │   ▲         │                                      │
│    │  Token Valid      │   │         │                                      │
│    │                   │   │         │                                      │
│    │              ┌────┴───┴────┐    │                                      │
│    │              │ AUTO-REFRESH │    │                                      │
│    │              │ TRIGGERED    │    │                                      │
│    │              │ (on 401      │    │                                      │
│    │              │  error)      │    │                                      │
│    │              └──────────────┘    │                                      │
│                                                                             │
│  Key: When API returns 401, client automatically:                           │
│       1. Uses refresh_token to get new access_token                         │
│       2. Retries the original request                                       │
│       3. Updates .token_cache.json                                          │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

**Refresh Flow:**

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  STEP 1: API Call Returns 401                                               │
│  GET https://graph.microsoft.com/v1.0/me/messages                          │
│    Authorization: Bearer {expired_token}                                    │
│    → 401 Unauthorized                                                       │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  STEP 2: Silent Token Refresh                                               │
│                                                                             │
│  POST https://login.microsoftonline.com/common/oauth2/v2.0/token           │
│    grant_type=refresh_token                                                 │
│    &client_id={client_id}                                                  │
│    &refresh_token={refresh_token}                                          │
│    &scope={requested_scopes}                                               │
│                                                                             │
│  Returns: New access_token + refresh_token                                  │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  STEP 3: Retry Original Request                                            │
│  GET https://graph.microsoft.com/v1.0/me/messages                          │
│    Authorization: Bearer {new_token}                                        │
│    → 200 OK                                                                 │
└─────────────────────────────────────────────────────────────────────────────┘
```

### How Each Tool Uses Tokens

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         TOOL EXECUTION FLOW                                 │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  1. LOAD TOKENS from .token_cache.json                                      │
│     - MSAL deserializes cache                                               │
│     - Extracts access_token and refresh_token                               │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  2. MAKE API CALL with access_token                                         │
│     GET https://graph.microsoft.com/v1.0/{endpoint}                        │
│       Authorization: Bearer {access_token}                                  │
└─────────────────────────────────────────────────────────────────────────────┘
                                    │
                                    ▼
┌─────────────────────────────────────────────────────────────────────────────┐
│  3. CHECK RESPONSE                                                          │
│     ┌──────────────┬────────────────────────────────┐                      │
│     │ Status Code │ Action                          │                      │
│     ├──────────────┼────────────────────────────────┤                      │
│     │ 200 OK      │ Return result                   │                      │
│     │ 401         │ Refresh token → Retry request   │                      │
│     │ 4xx/5xx     │ Return error                    │                      │
│     └──────────────┴────────────────────────────────┘                      │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Available Tools (43 total)

### Mail (10 tools)
| Tool | Description |
|------|-------------|
| `send_email` | Send a new email with subject, body, recipients, optional attachment |
| `create_draft` | Create a draft email (subject, body, recipients, optional attachment) |
| `create_draft_reply` | Create a draft reply to an existing message |
| `reply_email` | Send a plain text reply to a received message |
| `get_message` | Get one email message by `message_id` (optionally with headers) |
| `list_messages` | List messages from a folder with rich filters (inbox, drafts, etc.) |
| `search_messages` | Search messages by text, sender, subject, attachments, pagination |
| `update_email` | Update a **draft** message (subject, body, recipients, importance) |
| `move_message` | Move a message to another folder by `message_id` |
| `add_mail_attachment` | Attach a small file (<3 MB) to an existing message |

### Calendar (9 tools)
| Tool | Description |
|------|-------------|
| `create_calendar` | Create a new calendar |
| `list_calendars` | List calendars for the signed-in user |
| `create_event` | Create a new event (subject, body, start/end, time zone, attendees) |
| `list_events` | List events with filters, ordering, selection, timezone preference |
| `get_event` | Get full details for a specific event by `event_id` |
| `update_calendar_event` | Update event (subject, body, time, location, attendees, categories) |
| `delete_event` | Delete an event (optionally notify attendees) |
| `list_reminders` | Get reminders for events in a time range |
| `get_schedule` | Get free/busy availability for email addresses over a time window |

### Contacts (7 tools)
| Tool | Description |
|------|-------------|
| `create_contact` | Create a new contact in the default contacts folder |
| `list_contacts` | List contacts from the default or a specific contact folder |
| `get_contact` | Get details for a contact by `contact_id` |
| `update_contact` | Update an existing contact (name, emails, phones, company info) |
| `delete_contact` | Delete a contact by `contact_id` |
| `create_contact_folder` | Create a new contact folder |
| `get_contact_folders` | List contact folders with optional filters and expansions |

### Attachments (4 tools)
| Tool | Description |
|------|-------------|
| `download_attachment` | Download a specific attachment by `message_id` + `attachment_id` |
| `list_attachments` | List attachment metadata (name, size, type) for a message |
| `add_event_attachment` | Attach a file or item to a calendar event by `event_id` |
| `list_event_attachments` | List attachments for a specific calendar event |

### Folders & Rules (3 tools)
| Tool | Description |
|------|-------------|
| `list_mail_folders` | List top-level mail folders (Inbox, Drafts, Sent Items, etc.) |
| `create_mail_folder` | Create a new mail folder |
| `delete_mail_folder` | Delete an existing mail folder by `folder_id` |
| `create_email_rule` | Create a mail rule with conditions and actions |

### Settings & Profile (9 tools)
| Tool | Description |
|------|-------------|
| `get_profile` | Get the user's Outlook profile (basic user information) |
| `get_mailbox_settings` | Read mailbox settings (automatic replies, time zone, working hours) |
| `update_mailbox_settings` | Update automatic replies, language, time zone, or working hours |
| `get_mail_delta` | Get delta changes for messages (for sync scenarios) |
| `get_mail_tips` | Get mail tips (automatic replies, mailbox full, etc.) for recipients |
| `get_supported_languages` | List mailbox-supported languages |
| `get_supported_time_zones` | List supported time zones (Windows or IANA) |
| `create_master_category` | Create a new category in the user's master category list |
| `get_master_categories` | List all master categories |

---

## How to Get IDs for API Calls

### Getting `message_id` (for email operations)

```
Step 1: Call list_messages or search_messages
        └── Returns array of messages

Step 2: Each message has an "id" field
        └── This is your message_id

Example Response:
{
  "value": [
    {
      "id": "AQMkADAwATM0MDAAMi05OQBhNi1jNGUwLTAwAi0wMAoARgAAA7Z-YNwBy6BIv42xibO5ymcHAA3Ko5voYV5DnI2jTT2tVUkAAAIBDAAAAA3Ko5voYV5DnI2jTT2tVUkAAABOAGDKAAAA",
      "subject": "Hello",
      "from": {
        "emailAddress": {
          "address": "sender@example.com"
        }
      }
    }
  ]
}
```

### Getting `event_id` (for calendar operations)

```
Step 1: Call list_events
        └── Returns array of events

Step 2: Each event has an "id" field
        └── This is your event_id

Example Response:
{
  "value": [
    {
      "id": "AQMkADAwATM0MDAAMi05OQBhNi1jNGUwLTAwAi0wMAoARgAAA7Z-YNwBy6BIv42xibO5ymcHAA3Ko5voYV5DnI2jTT2tVUkAAAIBDQAAAA3Ko5voYV5DnI2jTT2tVUkAAABOADuAAAAA",
      "subject": "Team Meeting",
      "start": {
        "dateTime": "2026-02-05T10:00:00",
        "timeZone": "UTC"
      }
    }
  ]
}
```

### Getting `contact_id` (for contact operations)

```
Step 1: Call list_contacts
        └── Returns array of contacts

Step 2: Each contact has an "id" field
        └── This is your contact_id

Example Response:
{
  "value": [
    {
      "id": "AQMkADAwATM0MDAAMi05OQBhNi1jNGUwLTAwAi0wMAoARgAAA7Z-YNwBy6BIv42xibO5ymcHAA3Ko5voYV5DnI2jTT2tVUkAAAIBDgAAAA3Ko5voYV5DnI2jTT2tVUkAAABOABRgAAAA",
      "displayName": "John Doe",
      "emailAddresses": [
        {
          "address": "john@example.com"
        }
      ]
    }
  ]
}
```

### Getting `folder_id` (for folder operations)

```
Step 1: Call list_mail_folders
        └── Returns array of folders

Step 2: Each folder has an "id" field
        └── This is your folder_id

Example Response:
{
  "value": [
    {
      "id": "AQMkADAwATM0MDAAMi05OQBhNi1jNGUwLTAwAi0wMAoARgAAA7Z-YNwBy6BIv42xibO5ymcHAA3Ko5voYV5DnI2jTT2tVUkAAAIBDAAAAA3Ko5voYV5DnI2jTT2tVUkAAABOAGDKAAAA",
      "displayName": "Inbox",
      "childFolderCount": 0
    }
  ]
}

Alternatively, use well-known names:
- "inbox" - Inbox folder
- "drafts" - Drafts folder
- "sentitems" - Sent Items folder
- "deleteditems" - Deleted Items folder
```

### Getting `attachment_id` (for attachment operations)

```
Step 1: Call list_attachments with message_id
        └── Returns array of attachments

Step 2: Each attachment has an "id" field
        └── This is your attachment_id

Example Response:
{
  "value": [
    {
      "id": "AQMkADAwATM0MDAAMi05OQBhNi1jNGUwLTAwAi0wMAoARgAAA7Z-YNwBy6BIv42xibO5ymcHAA3Ko5voYV5DnI2jTT2tVUkAAAIBDAAAAA3Ko5voYV5DnI2jTT2tVUkAAABOAGDKAAAA",
      "name": "document.pdf",
      "contentType": "application/pdf",
      "size": 1024
    }
  ]
}
```

---

## Usage Examples

### Send an Email

```
// Step 1: Send email
send_email
  subject: "Hello from MCP"
  body: "This is a test email sent via Outlook MCP Server"
  to_email: "recipient@example.com"
  is_html: false

// Response: { "successful": true, "data": { "message": "Email sent successfully" } }
```

### Create and Update a Draft

```
// Step 1: Create draft
create_draft
  subject: "Draft Email"
  body: "This is a draft"
  to_recipients: ["recipient@example.com"]

// Response: { "successful": true, "data": { "id": "AQMkADAw..." } }

// Step 2: Update draft (use the id from Step 1)
update_email
  message_id: "AQMkADAw..."
  subject: "Updated Draft Email"
  body: { "contentType": "text", "content": "Updated content" }
```

### Create a Calendar Event

```
// Step 1: Create event
create_event
  subject: "Team Meeting"
  body: "Discussing project progress"
  start_datetime: "2026-02-05T10:00:00"
  end_datetime: "2026-02-05T11:00:00"
  time_zone: "UTC"
  attendees_info: [
    {
      "emailAddress": {
        "address": "attendee@example.com",
        "name": "John Doe"
      },
      "type": "required"
    }
  ]

// Response: { "successful": true, "data": { "id": "AQMkADAw..." } }
```

### Search and Reply to Messages

```
// Step 1: Search for messages
search_messages
  query: "important"
  fromEmail: "sender@example.com"
  size: 10

// Step 2: Get message details (use id from search results)
get_message
  message_id: "AQMkADAw..."

// Step 3: Reply to message
reply_email
  message_id: "AQMkADAw..."
  comment: "Thanks for your message!"
```

### Download an Attachment

```
// Step 1: List attachments for a message
list_attachments
  message_id: "AQMkADAw..."

// Step 2: Download specific attachment
download_attachment
  message_id: "AQMkADAw..."
  attachment_id: "AQMkADAw..."
  file_name: "document.pdf"
```

### Update Mailbox Settings

```
// Update automatic replies for vacation
update_mailbox_settings
  automaticRepliesSetting: {
    "status": "scheduled",
    "externalAudience": "all",
    "scheduledStartDateTime": {
      "dateTime": "2026-02-10T00:00:00",
      "timeZone": "UTC"
    },
    "scheduledEndDateTime": {
      "dateTime": "2026-02-15T23:59:59",
      "timeZone": "UTC"
    },
    "internalReplyMessage": "I'm on vacation",
    "externalReplyMessage": "I'm on vacation"
  }
```

---

## Configuration

### Required .env Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `OUTLOOK_CLIENT_ID` | Yes | Azure App Registration **Application (client) ID** from Overview page |
| `OUTLOOK_CLIENT_SECRET` | Yes | Client secret **Value** (not Secret ID) from Certificates & secrets - copy the Value column, shown only once! |
| `OUTLOOK_REDIRECT_URI` | No | Default: `https://login.microsoftonline.com/common/oauth2/nativeclient` |

**⚠️ Important:** When creating a client secret in Azure Portal:
- Go to **Certificates & secrets** → **New client secret**
- After creating, you'll see two columns: **Secret ID** and **Value**
- Copy the **Value** column (not the Secret ID) - this is your `OUTLOOK_CLIENT_SECRET`
- The Value is only displayed once and cannot be retrieved later!

### Auto-Detected (No Setup Needed)

| Variable | Description |
|----------|-------------|
| `user_id` | Defaults to `"me"` (current authenticated user) |
| All scopes | Hard-coded in `src/config.py`, requested automatically |

### Token Storage (.token_cache.json)

```json
{
  "AccessToken": {
    "secret": "eyJ0eXAi...",
    "expires_on": "1769688710"
  },
  "RefreshToken": {
    "secret": "0.AXcA..."
  },
  "Account": {
    "username": "user@example.com"
  }
}
```

---

## Troubleshooting

| Error | Cause | Solution |
|-------|-------|----------|
| `AADSTS65001: The user or administrator has not consented` | Missing API permissions | Add required permissions in Azure Portal and grant consent |
| `AADSTS700016: Application not found` | Invalid client ID | Double-check `OUTLOOK_CLIENT_ID` in `.env` |
| `Authentication expired. Please re-authenticate.` | Refresh token expired | Delete `.token_cache.json` and run `test_auth.py` again |
| `404 Client Error: Not Found` | Invalid message/event ID | Verify the ID exists using `list_messages` or `list_events` |
| `400 Client Error: Bad Request` | Invalid request format | Check parameter format (e.g., body must be `{"contentType": "text", "content": "..."}`) |
| `Only draft messages can be updated` | Trying to update received message | Use `list_messages` with `folder="drafts"` to get draft message IDs |
| `Body must be a dict with 'contentType' and 'content' fields` | Invalid body format | Use `{"contentType": "text", "content": "your text"}` or `{"contentType": "html", "content": "<p>your html</p>"}` |

---

## MCP Configuration

Add to your MCP client config (Cursor/Claude):

```json
{
  "mcpServers": {
    "outlook-mcp": {
      "command": "uv",
      "args": ["run", "run_server.py"],
      "cwd": "C:\\Users\\manus\\OneDrive\\Desktop\\microsoft oulook"
    }
  }
}
```

---

## Requirements

- Python 3.10+
- Microsoft account (personal or work/school)
- Azure AD App Registration with Microsoft Graph permissions
- Internet connection for Microsoft Graph API

---

## License

MIT License
