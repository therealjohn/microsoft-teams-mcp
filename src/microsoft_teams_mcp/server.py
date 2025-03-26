import asyncio
import os
import msal

from mcp.server.models import InitializationOptions
import mcp.types as types
from mcp.server import NotificationOptions, Server
import mcp.server.stdio
from dotenv import load_dotenv
import aiohttp
from typing import Dict, List, Optional, Tuple

load_dotenv()

SERVER_NAME = "microsoft-teams-mcp"
SERVER_VERSION = "0.1.0"
TOOL_NAME = "send-notification"
REQUIRED_ENV_VARS = ["BOT_ENDPOINT", "MICROSOFT_APP_ID", "MICROSOFT_APP_PASSWORD", "MICROSOFT_APP_TENANT_ID", "EMAIL"]

server = Server(SERVER_NAME)

async def get_auth_token(app_id: str, app_password: str, tenant_id: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Get authentication token using MSAL client credentials flow.
    
    Args:
        app_id: The application ID
        app_password: The application password/secret
        tenant_id: The tenant ID
        
    Returns:
        Tuple containing (access_token, error_message)
    """
    try:
        app = msal.ConfidentialClientApplication(
            client_id=app_id,
            client_credential=app_password,
            authority=f"https://login.microsoftonline.com/{tenant_id}"
        )
        
        scopes = [f"{app_id}/.default"]
        
        result = app.acquire_token_for_client(scopes)
        
        if "access_token" not in result:
            error_msg = result.get("error_description", "Failed to acquire token")
            return None, error_msg
            
        return result["access_token"], None
    except Exception as e:
        return None, str(e)

def validate_environment_variables() -> Tuple[Dict[str, str], List[str]]:
    """
    Validate required environment variables.
    
    Returns:
        Tuple containing (env_vars_dict, missing_vars_list)
    """
    env_vars = {}
        
    for var_name in REQUIRED_ENV_VARS:
        env_vars[var_name] = os.getenv(var_name)
        
    missing_vars = [var_name for var_name in REQUIRED_ENV_VARS if not env_vars[var_name]]
            
    return env_vars, missing_vars

async def send_notification(
    bot_endpoint: str, 
    access_token: str, 
    email: Optional[str], 
    message: str, 
    project: str
) -> Tuple[bool, Optional[str]]:
    """
    Send notification to the Teams bot endpoint.
    
    Args:
        bot_endpoint: The bot endpoint URL
        access_token: Authentication token
        email: User email (optional)
        message: Notification message
        project: Project name
    
    Returns:
        Tuple containing (success, error_message)
    """
    try:        
        payload = {
            "email": email,
            "message": message,
            "project": project,
        }
        
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        async with aiohttp.ClientSession() as session:
            async with session.post(bot_endpoint, json=payload, headers=headers) as response:
                if response.status >= 400:
                    response_text = await response.text()
                    return False, f"HTTP {response.status} - {response_text}"
                
                return True, None
                
    except Exception as e:
        return False, str(e)

@server.list_tools()
async def handle_list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name=TOOL_NAME,
            description="Send a notification message to the user. Supports markdown formatting for messages. Use backticks for code blocks and inline code. Use square brackets for placeholders.",
            inputSchema={
                "type": "object",
                "properties": {
                    "message": {"type": "string"},
                    "project": {"type": "string"},
                },
                "required": ["message", "project"],
            },
        )
    ]

@server.call_tool()
async def handle_call_tool(
    name: str, arguments: dict | None
) -> list[types.TextContent | types.ImageContent | types.EmbeddedResource]:
    if name != TOOL_NAME:
        raise ValueError(f"Unknown tool: {name}")

    if not arguments:
        raise ValueError("Missing arguments")

    message = arguments.get("message")
    project = arguments.get("project")

    if not message or not project:
        raise ValueError("Missing message or project")

    env_vars, missing_vars = validate_environment_variables()
    
    if missing_vars:
        return [
            types.TextContent(
                type="text",
                text=f"Missing required environment variables: {', '.join(missing_vars)}"
            )
        ]
    
    try:
        access_token, error = await get_auth_token(
            env_vars["MICROSOFT_APP_ID"], 
            env_vars["MICROSOFT_APP_PASSWORD"], 
            env_vars["MICROSOFT_APP_TENANT_ID"]
        )
        
        if error:
            return [
                types.TextContent(
                    type="text",
                    text=f"Authentication failed: {error}"
                )
            ]
        
        success, error_msg = await send_notification(
            env_vars["BOT_ENDPOINT"],
            access_token,
            env_vars["EMAIL"],
            message,
            project
        )
        
        if not success:
            return [
                types.TextContent(
                    type="text",
                    text=f"Failed to send notification: {error_msg}"
                )
            ]
        
        return [
            types.TextContent(
                type="text",
                text=f"Sent notification message for project '{project}' with content: {message}",
            )
        ]
                
    except Exception as e:
        return [
            types.TextContent(
                type="text",
                text=f"Error sending notification: {str(e)}"
            )
        ]

async def run_server():
    async with mcp.server.stdio.stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            InitializationOptions(
                server_name=SERVER_NAME,
                server_version=SERVER_VERSION,
                capabilities=server.get_capabilities(
                    notification_options=NotificationOptions(),
                    experimental_capabilities={},
                ),
            ),
        )

if __name__ == "__main__":
    asyncio.run(run_server())