# microsoft-teams-mcp MCP server

An MCP Server with a tool for Microsoft Teams chat notifications.

> [!WARNING]  
> This is provided for reference and wasn't tested with MCP clients other than VS Code.

## Components

### Tools

The server implements one tool:
- send-notification: Sends a notification message to Microsoft Teams
  - Takes "message" and "project" as required string arguments
  - Supports Markdown formatting for messages
  - Uses Azure AD authentication to securely communicate with Teams

## Configuration

This requires a Microsoft Teams bot to use for the notifications. You can use [my example Notification Bot](https://github.com/therealjohn/TeamsNotificationBotMCP) created with [Teams Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/teams-toolkit-fundamentals).

The server requires the following environment variables to be set:

- `BOT_ENDPOINT`: The URL endpoint of your Microsoft Teams bot
- `MICROSOFT_APP_ID`: Application (client) ID from Azure AD app registration
- `MICROSOFT_APP_PASSWORD`: Client secret from Azure AD app registration
- `MICROSOFT_APP_TENANT_ID`: Your Azure AD tenant ID
- `EMAIL`: The email address for the user receiving notifications

You can set these in a `.env` file in the project root directory.

## Quickstart

### Install

#### VS Code

This was tested using MCP support in VS Code, which at the time of creating this was available only in VS Code Insiders.

Add this to the VS Code Insiders Settings (JSON)

```
"mcp": {
  "inputs": [],
  "servers": {
      "MicrosoftTeams": {
          "command": "uv",
          "args": [
              "--directory",
              "<path/to/the/project>/microsoft-teams-mcp",
              "run",
              "microsoft-teams-mcp"
          ],
          "env": {
              "BOT_ENDPOINT": "<endpoint or dev tunnel URL of Teams bot>/api/notification",
              "MICROSOFT_APP_ID": "<microsoft-entra-client-id>",
              "MICROSOFT_APP_PASSWORD": "<microsoft-entra-client-secret>",
              "MICROSOFT_APP_TENANT_ID": "<microsoft-entra-tenant-id>",
              "EMAIL": "<your-email-in-teams>",
          }
      }
  }
    }
```

## Development

### Building

To prepare the package for distribution:

1. Sync dependencies and update lockfile:
```bash
uv sync
```

2. Build package distributions:
```bash
uv build
```
