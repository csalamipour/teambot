#!/usr/bin/env python3
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

import os
import sys
import traceback
from datetime import datetime
from http import HTTPStatus
from aiohttp import web
from aiohttp.web import Request, Response
from botbuilder.core import (
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    ConversationState,
    MemoryStorage,
    TurnContext,
)
from botbuilder.core.integration import aiohttp_error_middleware
from botbuilder.schema import Activity

from teams_bot import TeamsBot

# Read environment variables
PORT = int(os.environ.get("PORT", 3978))
APP_ID = os.environ.get("MicrosoftAppId", "")
APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")

# Create adapter
SETTINGS = BotFrameworkAdapterSettings(app_id=APP_ID, app_password=APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)

# Global error handling
async def on_error(context: TurnContext, error: Exception):
    # Print error details
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Send error message to user
    await context.send_activity("Sorry, an error occurred processing your request.")
    
    # Create a trace activity that contains the error object
    if context.activity.channel_id == "msteams":
        await context.send_activity("To continue, please try sending your message again or type '/start' to restart.")

ADAPTER.on_turn_error = on_error

# Create bot instance
BOT = TeamsBot()

# Create HTTP server
APP = web.Application(middlewares=[aiohttp_error_middleware])

# Listen for incoming requests on /api/messages
async def messages(req: Request) -> Response:
    # Main bot message handler
    if "application/json" in req.headers["Content-Type"]:
        body = await req.json()
    else:
        return Response(status=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)

    activity = Activity().deserialize(body)
    auth_header = req.headers["Authorization"] if "Authorization" in req.headers else ""

    # Call bot's on_turn method
    response = await ADAPTER.process_activity(activity, auth_header, BOT.on_turn)
    if response:
        return Response(body=response, status=HTTPStatus.OK)
    return Response(status=HTTPStatus.OK)

# Setup HTTP server routing
APP.router.add_post("/api/messages", messages)

if __name__ == "__main__":
    try:
        # Display startup info
        print("=" * 50)
        print(" Product Management Bot Server")
        print("=" * 50)
        print(f" Bot endpoint: http://localhost:{PORT}/api/messages")
        print(" Press Ctrl+C to exit")
        print("=" * 50)

        # Start the server
        web.run_app(APP, host="0.0.0.0", port=PORT)
    except Exception as error:
        print(f"Error starting the server: {error}")
        raise
