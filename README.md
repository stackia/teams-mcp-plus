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

### 🔐 Authentication & Multiple Tenants

- Connect multiple organizations with isolated credentials and explicit tenant selection per call
- Authenticate through Microsoft Graph OAuth device login, with persistent token caching and automatic refresh
- List connected tenants, check live authentication, and log out through the CLI
- Run in read-only mode with reduced permissions, or supply an existing Graph access token

### 👥 Users, Teams & Members

- Look up the current user or another user's profile
- Search users by name or email, including IDs and display text for `@mentions`
- List joined teams, their channels, and team members with roles
- Read chat member details, including user IDs, emails, tenant IDs, roles, and visible history start times

### 💬 Chats & Read Status

- List 1:1, group, and meeting chats with participants and latest-message previews, newest message first
- Find unread chats based on the current user's read position
- Mark chats read or unread for the current user
- Create 1:1 and group chats, with optional group topics
- Read chat history with time filters, ordering, pagination, and automatic retrieval of all available pages
- Read individual chat messages by ID, including multiple IDs in one call

### 🧵 Channel Threads & Messaging

- List channel thread roots and thread replies, or read specific messages and replies by ID, individually or in batches
- Start a new channel thread or reply within an existing thread
- Send chat messages or quote and reply to an existing chat message
- Edit or soft delete your own chat messages, channel messages, and channel replies
- Add or remove emoji reactions on chat messages, channel messages, and channel replies
- Compose messages with plain text or sanitized Markdown, `@mentions`, and importance levels
- Read message bodies as Markdown by default, or retain the original HTML

### 📎 Files & Inline Content

- Upload local files to chats through OneDrive or to channel threads through SharePoint
- Post channel files in a new thread or as a reply in an existing thread
- Upload large files through resumable upload sessions
- Include inline images in channel messages from a URL or base64 data
- Download hosted content from chat messages, channel messages, and channel replies as base64 or save it to disk

### 🔍 Message Search

- Search accessible chat and channel messages within a tenant using Microsoft Search and KQL
- Combine keywords, phrases, Boolean operators, sender, attachment, mention, and date filters
- Find messages mentioning the current user, with an optional recent-hours window
- Retrieve paginated results with optional relevance ranking and Markdown or raw message bodies

## Rich Message Formatting Support

The following tools support rich message formatting in Teams channels and chats:
- `send_channel_message`
- `send_chat_message`
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
- `search_messages`

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
- `ChatMessage.Send` - Send quoted chat replies
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

The server supports a read-only mode that disables all write tools, including sending or changing messages, reactions, chat read state, chat creation, and file uploads. Authentication in this mode requests only read-permission scopes from Microsoft Graph.

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

### Available MCP Tools

Full mode exposes **26 tools: 14 read tools and 12 write tools**. Read-only mode exposes only
those marked **Read** below. Tool schemas provide parameter details.

#### Authentication & Tenants

| Tool | Access | Function |
| --- | --- | --- |
| `list_tenants` | Read | List connected tenants, accounts, granted scopes, and the configured default. |
| `auth_status` | Read | Check live authentication and the signed-in user for a tenant. |

#### Users & Membership

| Tool | Access | Function |
| --- | --- | --- |
| `get_user` | Read | Retrieve the current user's profile or look up another user. |
| `search_users` | Read | Search users by name or email, including IDs and mention text. |
| `list_team_members` | Read | List a team's members and their roles. |
| `list_chat_members` | Read | List chat members with user and tenant IDs, emails, roles, and visible history start times. |

#### Chats

| Tool | Access | Function |
| --- | --- | --- |
| `list_chats` | Read | List chats with participants, latest-message previews, and read status; optionally return only unread chats. |
| `get_chat_messages` | Read | Retrieve chat history with filtering, ordering, and pagination, or read specific messages individually or in batches. |
| `create_chat` | Write | Create a 1:1 or group chat. |
| `set_chat_read_state` | Write | Mark a chat read or unread for the current user. |
| `send_chat_message` | Write | Send a chat message or quote and reply to an existing message. |
| `update_chat_message` | Write | Edit a chat message you sent. |
| `delete_chat_message` | Write | Soft delete a chat message you sent. |
| `set_chat_message_reaction` | Write | Add or remove a reaction on a chat message. |

#### Teams & Channels

| Tool | Access | Function |
| --- | --- | --- |
| `list_teams` | Read | List teams the current user has joined. |
| `list_channels` | Read | List a team's channels and their basic details. |
| `get_channel_messages` | Read | List thread roots or replies, or read specific channel messages or replies individually or in batches. |
| `send_channel_message` | Write | Start a new channel thread or reply within an existing thread. |
| `update_channel_message` | Write | Edit a channel message or reply you sent. |
| `delete_channel_message` | Write | Soft delete a channel message or reply you sent. |
| `set_channel_message_reaction` | Write | Add or remove a reaction on a channel message or reply. |

#### Files & Hosted Content

| Tool | Access | Function |
| --- | --- | --- |
| `send_file_to_chat` | Write | Upload a local file to OneDrive and share it in a chat. |
| `send_file_to_channel` | Write | Upload a local file to SharePoint and share it in a new or existing channel thread. |
| `download_chat_hosted_content` | Read | Download inline hosted content from a chat message, optionally saving it to disk. |
| `download_message_hosted_content` | Read | Download inline hosted content from a channel message or reply, optionally saving it to disk. |

#### Search

| Tool | Access | Function |
| --- | --- | --- |
| `search_messages` | Read | Search chat and channel messages with KQL, current-user mentions, and optional time filtering. |

Unread chat detection compares the latest message time with the current user's read position;
it does not use Search's `IsRead` filter or other participants' Seen receipts. Chats with unknown
read status are excluded from unread-only results. Reading messages does not mark them read.

Quoted chat replies require `ChatMessage.Send`. Existing connections without this permission
must authenticate again in full mode; injected access tokens must also include this scope.

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

### Read Multiple Messages

Pass an array of IDs to read several messages in one call:

```json
{ "name": "get_chat_messages", "arguments": { "chatId": "<chat>", "messageId": ["<message-1>", "<message-2>"] } }
```

For channels, use an array of root message IDs, or a single root ID with an array of reply IDs:

```json
{ "name": "get_channel_messages", "arguments": { "teamId": "<team>", "channelId": "<channel>", "messageId": ["<root-1>", "<root-2>"] } }
{ "name": "get_channel_messages", "arguments": { "teamId": "<team>", "channelId": "<channel>", "messageId": "<root>", "replyId": ["<reply-1>", "<reply-2>"] } }
```

Batches accept up to 50 IDs and query sequentially in input order. Successful reads appear in
`messages`; failures appear in `errors` with their requested IDs. The entire call is marked as
an error only when every read fails. List filters, sorting, and limits do not apply to these
ID-based reads. Existing single-ID calls and list operations remain available.

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
