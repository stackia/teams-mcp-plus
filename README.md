# Teams MCP Plus

[![npm version](https://img.shields.io/npm/v/teams-mcp-plus.svg)](https://www.npmjs.com/package/teams-mcp-plus)
[![npm downloads](https://img.shields.io/npm/dm/teams-mcp-plus.svg)](https://www.npmjs.com/package/teams-mcp-plus)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![GitHub stars](https://img.shields.io/github/stars/stackia/teams-mcp.svg)](https://github.com/stackia/teams-mcp/stargazers)

A Microsoft Teams MCP server with multi-tenant support. Connect multiple organizations in one server, with isolated credentials and explicit tenant selection for chats, channels, users, search, and file operations.

Based on [Floris Cornel’s Teams MCP](https://github.com/floriscornel/teams-mcp), extended and maintained in this repository.

## 📦 Installation

Authenticate each tenant first (replace `<tenant-id>` with its Microsoft Entra directory GUID):

```bash
npx -y teams-mcp-plus@latest authenticate --tenant <tenant-id> --name Work
npx -y teams-mcp-plus@latest tenants
```

Then add the following configuration in Cursor/Claude/VS Code:

```json
{
  "mcpServers": {
    "teams-mcp-plus": {
      "command": "npx",
      "args": ["-y", "teams-mcp-plus@latest"]
    }
  }
}
```

## 🚀 Features

### 🔐 Authentication
- Multiple tenants in one server, with isolated credentials and per-call `tenantId` selection
- OAuth 2.0 device code authentication flow with Microsoft Graph
- Secure token management, cache persistence, and refresh token renewal
- Authentication status checking and logout support
- Read-only mode with reduced scopes
- Direct `AUTH_TOKEN` support for pre-issued Microsoft Graph access tokens

### 👥 User Management
- Get current user information
- Search users by name or email
- Retrieve detailed user profiles
- Access organizational directory data

### 🏢 Microsoft Teams Integration
- **Teams Management**
  - List user's joined teams
  - Access team details and metadata

- **Channel Operations**
  - List channels within teams
  - Retrieve channel messages and replies
  - Send messages to team channels
  - Reply to existing channel threads
  - Edit and soft delete channel messages and replies
  - Support for message importance levels (`normal`, `high`, `urgent`)
  - Support for inline image attachments via URL or base64 data

- **Team Members**
  - List team members and their roles
  - Access member information
  - Search users for `@mentions`

### 💬 Chat & Messaging
- **1:1 and Group Chats**
  - List user's chats
  - Create new 1:1 or group conversations
  - Retrieve chat message history with filtering, ordering, and pagination
  - Fetch all available messages via `@odata.nextLink` pagination
  - Send messages to existing chats
  - Edit previously sent chat messages
  - Soft delete chat messages

### ✏️ Message Management
- **Edit & Delete**
  - Update (edit) sent messages in chats and channels
  - Soft delete messages in chats and channels (marks as deleted without permanent removal)
  - Only message senders can update/delete their own messages
  - Support for Markdown formatting, mentions, and importance levels on edits

### 📎 Media & Attachments
- **Hosted Content**
  - Download hosted content (images, files) from chat and channel messages
  - Access inline images and attachments shared in conversations
  - Optionally save hosted content directly to disk

- **File Upload**
  - Upload and send any file type (PDF, DOCX, XLSX, ZIP, images, etc.) to channels and chats
  - Large file support (>4 MB) via resumable upload sessions
  - Channel uploads go to SharePoint and chat uploads go to OneDrive
  - Optional message text, custom filename, formatting, and importance levels

### 🔍 Advanced Search & Discovery
- **Message Search**
  - Search across all Teams channels and chats using Microsoft Search API
  - Support for KQL (Keyword Query Language) syntax
  - Filter by sender, mentions, attachments, read state, and date ranges
  - Get recent messages with advanced filtering options
  - Find messages mentioning the current user

## Rich Message Formatting Support

The following tools support rich message formatting in Teams channels and chats:
- `send_channel_message`
- `send_chat_message`
- `reply_to_channel_message`
- `update_channel_message`
- `update_chat_message`
- `send_file_to_channel`
- `send_file_to_chat`

### Format Options

You can specify the `format` parameter to control the message formatting:
- `text` (default): Plain text
- `markdown`: Markdown formatting (bold, italic, lists, links, code, etc.) converted to sanitized HTML

When `format` is set to `markdown`, the message content is converted to HTML using a secure markdown parser and sanitized to remove potentially dangerous content before being sent to Teams.

If `format` is not specified, the message will be sent as plain text.

### Example Usage

```json
{
  "teamId": "...",
  "channelId": "...",
  "message": "**Bold text** and _italic text_\n\n- List item 1\n- List item 2\n\n[Link](https://example.com)",
  "format": "markdown",
  "importance": "high"
}
```

```json
{
  "chatId": "...",
  "message": "Simple plain text message",
  "format": "text"
}
```

### Security Features

- **HTML Sanitization**: All markdown content is converted to HTML and sanitized to remove potentially dangerous elements (scripts, event handlers, etc.)
- **Allowed Tags**: Only safe HTML tags are permitted (p, strong, em, a, ul, ol, li, h1-h6, code, pre, etc.)
- **Safe Attributes**: Only safe attributes are allowed
- **XSS Prevention**: Content is automatically sanitized to prevent cross-site scripting attacks

### Supported Markdown Features

- **Text formatting**: Bold (`**text**`), italic (`_text_`), strikethrough (`~~text~~`)
- **Links**: `[text](url)`
- **Lists**: Bulleted (`- item`) and numbered (`1. item`)
- **Code**: Inline `` `code` `` and fenced code blocks
- **Headings**: `# H1` through `###### H6`
- **Blockquotes**: `> quoted text`
- **Tables**: GitHub-flavored markdown tables

## LLM-Friendly Content Format

Messages retrieved from the Microsoft Graph API are returned as raw HTML containing Teams-specific tags. To make this content more consumable by AI assistants, the following tools support automatic HTML-to-Markdown conversion:

- `get_chat_messages`
- `get_channel_messages`
- `get_channel_message_replies`
- `search_messages`
- `get_my_mentions`

### Content Format Options

Use the `contentFormat` parameter to control how message content is returned:
- `markdown` (default): Converts Teams HTML to clean Markdown, optimized for LLM consumption
- `raw`: Returns the original HTML from the Microsoft Graph API

### What Gets Converted

| HTML Element                           | Markdown Output                                           |
| -------------------------------------- | --------------------------------------------------------- |
| `<at id="0">Name</at>` (Teams mention) | `@Name` (multi-word names merged using mentions metadata) |
| `<strong>text</strong>`                | `**text**`                                                |
| `<em>text</em>`                        | `*text*`                                                  |
| `<code>text</code>`                    | `` `text` ``                                              |
| `<a href="url">text</a>`               | `[text](url)`                                             |
| `<ul><li>item</li></ul>`               | `- item`                                                  |
| `<table>...</table>`                   | GFM Markdown table                                        |
| `<attachment id="...">`                | `{attachment:id}`                                         |
| `<systemEventMessage/>`                | *(removed)*                                               |
| `<hr>`                                 | `---`                                                     |
| `&nbsp;`, `&amp;`, etc.                | Decoded to plain characters                               |

### Attachment Metadata

Messages that contain file attachments or inline images include an `attachments` array in the response with metadata for each attachment (id, name, contentType, contentUrl, thumbnailUrl). The inline `{attachment:id}` markers in the markdown content correlate with entries in this array, allowing consumers to identify and download attachments via `download_message_hosted_content` or `download_chat_hosted_content`.

### Example Usage

```json
{
  "chatId": "19:meeting_...",
  "limit": 10,
  "contentFormat": "markdown"
}
```

To get the original HTML:

```json
{
  "chatId": "19:meeting_...",
  "limit": 10,
  "contentFormat": "raw"
}
```

## 📦 Installation

```bash
# Use the project's Node.js LTS version
nvm install
nvm use

# Install dependencies
npm install

# Build the project
npm run build

# Set up authentication
npm run auth -- --tenant <tenant-id>
```

## 🔧 Configuration

### Prerequisites
- Node.js 22.22.2+, 24.15.0+, or 26+ (Node.js 20 is no longer supported)
- Microsoft 365 account with appropriate permissions
- Microsoft Graph delegated permissions for the scopes below

### Required Microsoft Graph Permissions

**Full mode (default):**
- `User.Read` - Read user profile
- `User.ReadBasic.All` - Read basic user info
- `Team.ReadBasic.All` - Read team information
- `Channel.ReadBasic.All` - Read channel information
- `ChannelMessage.Read.All` - Read channel messages
- `ChannelMessage.Send` - Send channel messages and replies
- `ChannelMessage.ReadWrite` - Edit and delete channel messages
- `Chat.Read` - Read chat messages (included via read-only scopes)
- `Chat.ReadWrite` - Create and manage chats, send/edit/delete chat messages (supersedes `Chat.Read`)
- `TeamMember.Read.All` - Read team members
- `Files.ReadWrite.All` - Required for file uploads to channels and chats

**Read-only mode** (`TEAMS_MCP_READ_ONLY=true`) — only these scopes are requested:
- `User.Read`
- `User.ReadBasic.All`
- `Team.ReadBasic.All`
- `Channel.ReadBasic.All`
- `ChannelMessage.Read.All`
- `TeamMember.Read.All`
- `Chat.Read`

### Multiple Tenants

One MCP server can access several organizations concurrently. Authenticate **each target tenant separately**, including organizations where you are a guest. Use the Microsoft Entra directory **tenant ID (GUID)**; `common`, `organizations`, and tenant domains are not accepted. Each tenant stores one selected account; authenticating again replaces that tenant's active login. Different tenants may use the same home account or entirely different users. `--name` is a display label, not a routing alias.

For a source checkout, build first, then connect two tenants:

```bash
npm ci
npm run build
node dist/index.js authenticate --tenant aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa --name Work
node dist/index.js authenticate --tenant bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb --name Customer --read-only
node dist/index.js tenants
node dist/index.js check
```

Replace the example IDs with real tenant IDs. The account must belong to or be invited into each tenant, and that tenant must permit the application and required Graph permissions. A login in the home tenant does not automatically authorize guest tenants.

Point your MCP client at the built checkout:

```json
{
  "mcpServers": {
    "teams": {
      "command": "node",
      "args": ["/absolute/path/to/teams-mcp/dist/index.js"],
      "env": {
        "TEAMS_MCP_CONFIG_DIR": "/absolute/path/to/shared/teams-credentials"
      }
    }
  }
}
```

Omit `TEAMS_MCP_CONFIG_DIR` to use `~/.teams-mcp-plus`; if you set it, use the same value when authenticating. For the published package, use `npx -y teams-mcp-plus@latest` instead of `node dist/index.js`.

Call `list_tenants` first. Every other tool accepts an optional `tenantId`:

```json
{ "name": "list_teams", "arguments": { "tenantId": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa" } }
```

```json
{ "name": "search_messages", "arguments": { "tenantId": "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb", "query": "release" } }
```

Selection order is explicit tool `tenantId`, server `--tenant` / `TEAMS_MCP_TENANT_ID`, then the sole connected tenant. With multiple connections and no default, omission returns a selection error. An unknown or logged-out explicit/default tenant always fails; it never falls back to another tenant. Keep resource IDs with the tenant that produced them; results and searches are scoped to one tenant per call, not aggregated across organizations.

```bash
# Optional server default; individual tool calls can still target another tenant
node dist/index.js --tenant aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa
# Live check of one tenant; without --tenant/default, check covers all tenants
node dist/index.js check --tenant bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb
# Logout requires an explicit target, even if an environment default is configured
node dist/index.js logout --tenant bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb
node dist/index.js logout --all
```

`tenants` / `list_tenants` list local connections; they do not guarantee that consent or refresh tokens are still valid. `check` and `auth_status` acquire a token and call Graph `/me`. `check` exits nonzero if any checked connection fails. New logins and logout are detected by running servers without a restart; requests already submitted to Graph cannot be recalled. Logout clears local files, not Microsoft account sessions or consent. An injected `AUTH_TOKEN` remains active until removed from the environment and the server restarted.

**Isolation design:** tenant-specific authorities are used for both device-code login and silent refresh. Account selection matches the persisted home account ID, local account ID, and tenant ID instead of taking the first cached account. A tool receives an immutable tenant-bound service, also passed to file uploads and user lookups. Each login has a separate cache revision; stale processes cannot overwrite the active login. Previous cache revisions are removed when that tenant is logged out. This follows Microsoft's guidance on [MSAL account selection](https://learn.microsoft.com/en-us/entra/msal/javascript/node/accounts) and [authority configuration](https://learn.microsoft.com/en-us/entra/identity-platform/msal-client-application-configuration).

### Authentication Modes

**Full access:**

```bash
npx teams-mcp-plus@latest authenticate --tenant <tenant-id>
```

**Read-only access:**

```bash
npx teams-mcp-plus@latest authenticate --tenant <tenant-id> --read-only
```

**Direct token injection with an existing Microsoft Graph JWT:**

```json
{
  "mcpServers": {
    "teams-mcp-plus": {
      "command": "npx",
      "args": ["-y", "teams-mcp-plus@latest"],
      "env": {
        "AUTH_TOKEN": "<jwt-for-https://graph.microsoft.com>"
      }
    }
  }
}
```

### Token Storage

- Each tenant has its own directory: `~/.teams-mcp-plus/<tenant-id>/`.
- `profile.json` records the selected account, label, granted scopes, and login revision.
- `<revision>.cache.json` contains the MSAL tokens for that login. Files use mode `0600`, and newly created directories use `0700` on POSIX systems. Writes are atomic. These files contain credentials in plaintext; protect the directory.
- `TEAMS_MCP_CONFIG_DIR` overrides the storage root. All CLI and MCP processes that share connections must use the same directory.
- Old single-account files are neither read nor migrated. Re-authenticate each tenant; obsolete legacy files can be removed manually.

## 🛠️ Usage

### Starting the Server
```bash
# Development mode with hot reload
npm run dev

# Production mode
npm run build && node dist/index.js

# Start in read-only mode (disables all write tools)
TEAMS_MCP_READ_ONLY=true node dist/index.js
```

### CLI Commands

```bash
npx teams-mcp-plus@latest authenticate --tenant <tenant-id>              # Authenticate with full scopes
npx teams-mcp-plus@latest authenticate --tenant <tenant-id> --read-only  # Authenticate with read-only scopes
npx teams-mcp-plus@latest check                     # Check authentication status
npx teams-mcp-plus@latest logout --tenant <tenant-id>                    # Clear authentication
npx teams-mcp-plus@latest auth --tenant <tenant-id>   # Alias for authenticate
npx teams-mcp-plus@latest                           # Start MCP server (default)
```

### Environment Variables

- `TEAMS_MCP_READ_ONLY=true` - Start the MCP server in read-only mode
- `TEAMS_MCP_TENANT_ID=<tenant-id>` - Default tenant; `--tenant` takes precedence. Explicit tool `tenantId` overrides both.
- `TEAMS_MCP_CONFIG_DIR=<path>` - Shared credential directory (default: `~/.teams-mcp-plus`).
- `AUTH_TOKEN=<jwt>` - Pre-issued Graph token, used only for the tenant in its `tid` claim. Other tenants continue to use their own MSAL credentials. The token must have a valid Graph audience and unexpired `exp`; it is never written to disk or refreshed.

### Read-Only Mode

The server supports a read-only mode that disables all write operations (sending messages, creating chats, uploading files, editing/deleting messages) and requests only read-permission scopes from Microsoft Graph.

**Enable read-only mode** using either:
- Environment variable: `TEAMS_MCP_READ_ONLY=true`
- CLI flag: `--read-only`

**Authenticate with reduced scopes:**
```bash
npx teams-mcp-plus@latest authenticate --tenant <tenant-id> --read-only
```

**MCP server configuration (read-only):**
```json
{
  "mcpServers": {
    "teams-mcp-plus": {
      "command": "npx",
      "args": ["-y", "teams-mcp-plus@latest"],
      "env": {
        "TEAMS_MCP_READ_ONLY": "true"
      }
    }
  }
}
```

**Switching modes:** Scope grants are stored per tenant. A full-mode server can expose write tools while a tenant still has read-only grants; that tenant's write requests will fail with a Graph permission error. Re-authenticate that tenant without `--read-only` to request write permissions:
```bash
npx teams-mcp-plus@latest authenticate --tenant <tenant-id>
```

**Read-only tools (18):**
`list_tenants`, `auth_status`, `get_current_user`, `search_users`, `get_user`, `list_teams`, `list_channels`, `get_channel_messages`, `get_channel_message_replies`, `list_team_members`, `search_users_for_mentions`, `download_message_hosted_content`, `list_chats`, `list_chat_members`, `get_chat_messages`, `download_chat_hosted_content`, `search_messages`, `get_my_mentions`

**Write tools disabled in read-only mode (15):**
`set_chat_read_state`, `send_channel_message`, `reply_to_channel_message`, `update_channel_message`, `delete_channel_message`, `send_file_to_channel`, `send_chat_message`, `create_chat`, `update_chat_message`, `delete_chat_message`, `send_file_to_chat`, `set_channel_message_reaction`, `unset_channel_message_reaction`, `set_chat_message_reaction`, `unset_chat_message_reaction`

### Available MCP Tools

#### Authentication
- `list_tenants` - List tenant IDs, labels, account names, scopes, and the configured default (no token refresh)
- `auth_status` - Check current authentication status

#### User Operations
- `get_current_user` - Get authenticated user information
- `search_users` - Search for users by name or email
- `get_user` - Get detailed user information by ID or email

#### Teams Operations
- `list_teams` - List user's joined teams
- `list_channels` - List channels in a specific team
- `get_channel_messages` - List thread roots, or pass `messageId` to read one root; add `replyId` to read a reply within that thread
- `get_channel_message_replies` - List replies within the thread identified by its root `messageId`
- `send_channel_message` - Start a new channel thread by posting its root message
- `reply_to_channel_message` - Post a reply in the existing thread identified by its root `messageId`
- `update_channel_message` - Edit a previously sent channel message or reply
- `delete_channel_message` - Soft delete a channel message or reply
- `list_team_members` - List members of a specific team
- `search_users_for_mentions` - Search for team members to @mention in messages
- `send_file_to_channel` - Post a file in a new thread, or pass the root `messageId` to reply in an existing thread

A channel thread consists of a root message and its replies. For channel tools,
`messageId` identifies the root message; `replyId`, where supported, identifies an
individual reply within that thread. `reply_to_channel_message` appends to the thread
and does not select an individual reply to quote. These correspond to Graph's
[new message](https://learn.microsoft.com/en-us/graph/api/chatmessage-post?view=graph-rest-1.0)
and [thread reply](https://learn.microsoft.com/en-us/graph/api/chatmessage-post-replies?view=graph-rest-1.0) endpoints.

#### Chat Operations
- `list_chats` - List all user's chats (1:1, group, and meeting), with read status and latest-message previews; use `unreadOnly: true` to filter unread chats
- `get_chat_messages` - List chat messages with pagination and filters, or pass `messageId` to read one
- `list_chat_members` - List all chat members with membership IDs, user IDs, names, emails, tenant IDs, roles, and visible history start times
- `set_chat_read_state` - Mark a chat read (`isRead: true`) or unread (`isRead: false`) for the current user
- `send_chat_message` - Send a message to a chat
- `create_chat` - Create a new 1:1 or group chat
- `update_chat_message` - Edit a previously sent chat message
- `delete_chat_message` - Soft delete a chat message
- `send_file_to_chat` - Upload a local file and send it as a message to a chat

To find unread chats, call `list_chats` with:

```json
{
  "tenantId": "<tenant-id from list_tenants>",
  "unreadOnly": true
}
```

The tool lists chats newest message first (`$orderby=lastMessagePreview/createdDateTime desc`),
follows all chat pages, and compares `lastMessagePreview.createdDateTime` with
`viewpoint.lastMessageReadDateTime`, as described in the
[Microsoft Graph documentation](https://learn.microsoft.com/en-us/graph/api/chat-list?view=graph-rest-1.0#example-4-list-chats-along-with-the-preview-of-the-last-message-sent-in-the-chat).
It does not use the Search API's `IsRead` filter. Results retain the existing chat-list array
format and include `isUnread`, `isHidden`, `lastMessageReadDateTime`, and `lastMessagePreview`
(message ID, Markdown content, sender name, and creation time). Hidden chats are included.
Missing or invalid timestamps produce `isUnread: null`; these chats are excluded when
`unreadOnly` is true. If none match, the response notes any chats with unknown read status.
Omit `unreadOnly` (or set it to false) to list all chats, including those with unknown status.
This indicates messages after your read position, not other participants' Seen receipts.
It covers chats, not channels, and does not change read state. To fetch the messages,
use `get_chat_messages` with the returned chat ID and `since: lastMessageReadDateTime`.

Single-message reads keep the `{ totalReturned, hasMore, messages }` response envelope and
support `contentFormat` (`markdown` or `raw`). Chat list filters, sorting, and pagination
are ignored when `messageId` is supplied; channel `limit` is also ignored for single reads.
A channel `replyId` requires the parent `messageId`. Reads do not mark messages read.

```json
{ "name": "get_chat_messages", "arguments": { "chatId": "<chat>", "messageId": "<message>" } }
{ "name": "get_channel_messages", "arguments": { "teamId": "<team>", "channelId": "<channel>", "messageId": "<parent>", "replyId": "<reply>" } }
{ "name": "list_chat_members", "arguments": { "chatId": "<chat>" } }
{ "name": "set_chat_read_state", "arguments": { "chatId": "<chat>", "isRead": true } }
```

`set_chat_read_state` uses the existing delegated `Chat.ReadWrite` permission and is disabled
in read-only mode. When marking unread, omit `lastMessageReadDateTime` to mark the latest
message unread, or supply an ISO timestamp to mark messages after that time unread.
This timestamp is rejected with `isRead: true`. See Graph's
[mark read](https://learn.microsoft.com/en-us/graph/api/chat-markchatreadforuser?view=graph-rest-1.0)
and [mark unread](https://learn.microsoft.com/en-us/graph/api/chat-markchatunreadforuser?view=graph-rest-1.0) APIs.
Member `id` identifies the membership record; use `userId` for mentions and user lookup.
All calls support the existing optional `tenantId` selector.

#### Media Operations
- `download_message_hosted_content` - Download hosted content (images, files) from channel messages
- `download_chat_hosted_content` - Download hosted content (images, files) from chat messages

#### Search Operations
- `search_messages` - Search across all Teams messages using KQL syntax
- `get_my_mentions` - Find recent messages mentioning the current user

## 📋 Examples

### Authentication

First, authenticate with Microsoft Graph:

```bash
# Full access (default)
npx teams-mcp-plus@latest authenticate --tenant <tenant-id>

# Read-only (reduced permission scopes)
npx teams-mcp-plus@latest authenticate --tenant <tenant-id> --read-only
```

Check your authentication status:

```bash
npx teams-mcp-plus@latest check
```

Logout if needed:

```bash
npx teams-mcp-plus@latest logout --tenant <tenant-id>
```

### Chat Pagination Example

```json
{
  "chatId": "19:meeting_...",
  "limit": 100,
  "fetchAll": true,
  "orderBy": "createdDateTime",
  "descending": true,
  "contentFormat": "markdown"
}
```

### Channel Message with Mentions and Image

```json
{
  "teamId": "team-id",
  "channelId": "channel-id",
  "message": "Please review **today's update**",
  "format": "markdown",
  "importance": "high",
  "mentions": [
    {
      "mention": "alex.chen",
      "userId": "00000000-0000-0000-0000-000000000000"
    }
  ],
  "imageUrl": "https://example.com/status.png"
}
```

### File Upload Example

```json
{
  "chatId": "19:meeting_...",
  "filePath": "/absolute/path/to/report.pdf",
  "message": "Please review the attached report",
  "format": "markdown"
}
```

### Integrating with Cursor/Claude

This MCP server is designed to work with AI assistants like Claude/Cursor/VS Code through the Model Context Protocol.

```json
{
  "mcpServers": {
    "teams-mcp-plus": {
      "command": "npx",
      "args": ["-y", "teams-mcp-plus@latest"]
    }
  }
}
```

## 🔒 Security

- All authentication is handled through Microsoft's OAuth 2.0 flow or a caller-provided Microsoft Graph token
- **Refresh token support**: Access tokens are automatically renewed using cached refresh tokens, so you don't need to re-authenticate every hour
- Credentials are isolated per tenant under `~/.teams-mcp-plus/<tenant-id>/`; logout removes all login generations for that tenant.
- Markdown content is sanitized before sending HTML to Teams
- `AUTH_TOKEN` routing checks its Graph audience, tenant ID, and expiry. Microsoft Graph validates the token signature and permissions.
- No sensitive data is logged or exposed
- Follows Microsoft Graph API security best practices

## 📝 License

MIT License - see LICENSE file for details

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Run build, linting, and tests
5. Submit a pull request

## 📞 Support

For issues and questions:
- Check the existing GitHub issues
- Review Microsoft Graph API documentation
- Ensure proper authentication and permissions are configured
