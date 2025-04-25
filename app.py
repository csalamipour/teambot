import os
import sys
import traceback
import logging
import tempfile
import json
from datetime import datetime
from http import HTTPStatus
from typing import Dict, Any, Optional, List

from fastapi import FastAPI, Request, Response, HTTPException, UploadFile, File, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from botbuilder.core import (
    BotFrameworkAdapterSettings,
    TurnContext,
    BotFrameworkAdapter
)
from botbuilder.schema import Activity, ActivityTypes, Attachment, ConversationReference

import requests
import aiohttp
import asyncio

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# Your FastAPI backend URL - already deployed
API_BASE_URL = "https://copilotv2.azurewebsites.net"

# Dictionary to store conversation state for each user
# Key: conversation_id, Value: dict with assistant_id, session_id, etc.
conversation_states = {}

# App credentials from environment variables
APP_ID = os.environ.get("MicrosoftAppId", "")
APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")

# Create adapter with proper settings
SETTINGS = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)

# Create FastAPI app
app = FastAPI(title="Teams Product Management Bot")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Catch-all for errors
async def on_error(context: TurnContext, error: Exception):
    # Print the error to the console
    logger.error(f"\n [on_turn_error] unhandled error: {error}")
    traceback.print_exc()
    
    # Send a message to the user
    await context.send_activity("The bot encountered an error. Please try again.")
    
    # Send a trace activity if we're talking to the Bot Framework Emulator
    if context.activity.channel_id == "emulator":
        # Create a trace activity that contains the error object
        trace_activity = Activity(
            label="TurnError",
            name="on_turn_error Trace",
            timestamp=datetime.utcnow(),
            type=ActivityTypes.trace,
            value=f"{error}",
            value_type="https://www.botframework.com/schemas/error",
        )
        # Send a trace activity, which will be displayed in Bot Framework Emulator
        await context.send_activity(trace_activity)

# Assign the error handler
ADAPTER.on_turn_error = on_error

# Create typing indicator activity
def create_typing_activity() -> Activity:
    return Activity(
        type=ActivityTypes.typing,
        channel_id="msteams"
    )

# Bot logic handler
async def bot_logic(turn_context: TurnContext):
    # Get the conversation reference for later use
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_id = conversation_reference.conversation.id
    
    # Initialize state for this conversation if not exists
    if conversation_id not in conversation_states:
        conversation_states[conversation_id] = {
            "assistant_id": None,
            "session_id": None,
            "vector_store_id": None,
            "uploaded_files": []
        }
    
    state = conversation_states[conversation_id]
    
    # Handle different activity types
    if turn_context.activity.type == ActivityTypes.message:
        # Handle file attachments
        if turn_context.activity.attachments and len(turn_context.activity.attachments) > 0:
            await handle_file_upload(turn_context, state)
        # Handle text messages
        elif turn_context.activity.text:
            await handle_text_message(turn_context, state)
    
    # Handle conversation update (bot added to conversation)
    elif turn_context.activity.type == ActivityTypes.conversation_update:
        if turn_context.activity.members_added:
            for member in turn_context.activity.members_added:
                if member.id != turn_context.activity.recipient.id:
                    # Bot was added - send welcome message
                    await send_welcome_message(turn_context)

# Function to handle file uploads
async def handle_file_upload(turn_context: TurnContext, state):
    for attachment in turn_context.activity.attachments:
        try:
            # Send typing indicator
            await turn_context.send_activity(create_typing_activity())
            
            # Download the file content
            file_content = await download_attachment(turn_context, attachment)
            if not file_content:
                await turn_context.send_activity(f"Sorry, I couldn't download the file '{attachment.name}'.")
                continue
                
            # Message user that file is being processed
            await turn_context.send_activity(f"Processing file: '{attachment.name}'...")
            
            # If no assistant yet, initialize chat first
            if not state["assistant_id"]:
                await initialize_chat(turn_context, state)
            
            # Create a temporary file to handle the upload properly
            with tempfile.NamedTemporaryFile(delete=False, suffix='_' + attachment.name) as temp:
                temp.write(file_content)
                temp_path = temp.name
            
            try:
                # Upload file to the backend
                with open(temp_path, 'rb') as file:
                    files = {"file": (attachment.name, file)}
                    data = {"assistant": state["assistant_id"]}
                    
                    # Add session if available
                    if state["session_id"]:
                        data["session"] = state["session_id"]
                        
                    response = requests.post(
                        f"{API_BASE_URL}/upload-file", 
                        files=files,
                        data=data
                    )
                
                if response.status_code == 200:
                    result = response.json()
                    state["uploaded_files"].append(attachment.name)
                    await turn_context.send_activity(f"File '{attachment.name}' uploaded successfully!")
                    
                    # If it's an image, show the analysis
                    if "processing_method" in result and result["processing_method"] == "thread_message":
                        await turn_context.send_activity("Here's my analysis of the image:")
                        await send_message(turn_context, state)
                else:
                    await turn_context.send_activity(f"Failed to upload file. Status code: {response.status_code}")
                    if response.text:
                        try:
                            error_json = response.json()
                            await turn_context.send_activity(f"Error details: {json.dumps(error_json)}")
                        except:
                            await turn_context.send_activity(f"Error details: {response.text[:500]}")
            finally:
                # Clean up the temporary file
                try:
                    os.unlink(temp_path)
                except:
                    pass
                
        except Exception as e:
            await turn_context.send_activity(f"Error processing file: {str(e)}")
            logger.error(f"Error processing file: {str(e)}")
            traceback.print_exc()

# Download attachment content from Teams
async def download_attachment(turn_context: TurnContext, attachment: Attachment):
    try:
        if attachment.content_url:
            connector = turn_context.adapter.create_connector_client(
                turn_context.activity.service_url
            )
            
            response = await connector.attachments.get_attachment_content(
                attachment.content_url,
            )
            
            if response:
                return response
            
        return None
    except Exception as e:
        logger.error(f"Error downloading attachment: {str(e)}")
        traceback.print_exc()
        return None

# Function to handle text messages
async def handle_text_message(turn_context: TurnContext, state):
    user_message = turn_context.activity.text.strip()
    
    # If no assistant yet, initialize chat with the message as context
    if not state["assistant_id"]:
        await initialize_chat(turn_context, state, context=user_message)
        return
    
    # Send typing indicator
    await turn_context.send_activity(create_typing_activity())
    
    # Send message to the backend and get response
    params = {
        "session": state["session_id"],
        "assistant": state["assistant_id"],
        "prompt": user_message,
    }
    
    try:
        # Get non-streaming response
        response = requests.get(f"{API_BASE_URL}/chat", params=params)
        
        if response.status_code == 200:
            result = response.json()
            assistant_response = result.get("response", "I'm sorry, I couldn't process your request.")
            
            # Split long responses into chunks if needed (Teams has message size limits)
            if len(assistant_response) > 7000:
                chunks = [assistant_response[i:i+7000] for i in range(0, len(assistant_response), 7000)]
                for i, chunk in enumerate(chunks):
                    if i == 0:
                        await turn_context.send_activity(chunk)
                    else:
                        await turn_context.send_activity(f"(continued) {chunk}")
            else:
                await turn_context.send_activity(assistant_response)
        else:
            await turn_context.send_activity(f"Failed to get a response. Status code: {response.status_code}")
            try:
                error_json = response.json()
                await turn_context.send_activity(f"Error details: {json.dumps(error_json)}")
            except:
                await turn_context.send_activity(f"Error details: {response.text[:500]}")
            
    except Exception as e:
        await turn_context.send_activity(f"Error processing your message: {str(e)}")
        logger.error(f"Error in handle_text_message: {str(e)}")
        traceback.print_exc()

# Initialize chat with the backend
async def initialize_chat(turn_context: TurnContext, state, context=None):
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Prepare data for initialization
        data = {}
        if context:
            data["context"] = context
            
        # Make initiate-chat request
        response = requests.post(f"{API_BASE_URL}/initiate-chat", data=data)
        
        if response.status_code == 200:
            result = response.json()
            state["assistant_id"] = result["assistant"]
            state["session_id"] = result["session"]
            state["vector_store_id"] = result["vector_store"]
            
            # Tell the user chat was initialized
            await turn_context.send_activity("Hi! I'm the Product Management Bot. I'm ready to help you with your product management tasks.")
            
            if context:
                await turn_context.send_activity(f"I've initialized with your context: '{context}'")
                # Also send the first response
                await send_message(turn_context, state)
        else:
            await turn_context.send_activity(f"Failed to initialize chat. Status code: {response.status_code}")
            try:
                error_json = response.json()
                await turn_context.send_activity(f"Error details: {json.dumps(error_json)}")
            except:
                await turn_context.send_activity(f"Error details: {response.text[:500]}")
    
    except Exception as e:
        await turn_context.send_activity(f"Error initializing chat: {str(e)}")
        logger.error(f"Error in initialize_chat: {str(e)}")
        traceback.print_exc()

# Send a message without user input (used after file upload or initialization)
async def send_message(turn_context: TurnContext, state):
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Get the latest message from the thread
        params = {
            "session": state["session_id"],
            "assistant": state["assistant_id"],
        }
        
        response = requests.get(f"{API_BASE_URL}/chat", params=params)
        
        if response.status_code == 200:
            result = response.json()
            assistant_response = result.get("response", "")
            
            if assistant_response:
                # Split long responses if needed
                if len(assistant_response) > 7000:
                    chunks = [assistant_response[i:i+7000] for i in range(0, len(assistant_response), 7000)]
                    for i, chunk in enumerate(chunks):
                        if i == 0:
                            await turn_context.send_activity(chunk)
                        else:
                            await turn_context.send_activity(f"(continued) {chunk}")
                else:
                    await turn_context.send_activity(assistant_response)
            
    except Exception as e:
        await turn_context.send_activity(f"Error getting response: {str(e)}")
        logger.error(f"Error in send_message: {str(e)}")
        traceback.print_exc()

# Send welcome message when bot is added
async def send_welcome_message(turn_context: TurnContext):
    welcome_text = (
        "# Welcome to the Product Management Bot! 👋\n\n"
        "I'm here to help you with your product management tasks. I can:\n\n"
        "- Create and edit product requirements documents\n"
        "- Analyze data from CSV and Excel files\n"
        "- Answer questions about uploaded documents\n"
        "- Analyze images and provide insights\n\n"
        "To get started, you can:\n"
        "- Send me a message with your request\n"
        "- Upload a file for analysis\n"
        "- Ask me to create a PRD\n\n"
        "How can I assist you today?"
    )
    
    await turn_context.send_activity(welcome_text)

# FastAPI endpoint to handle Bot Framework messages
@app.post("/api/messages")
async def messages(req: Request) -> Response:
    # Check content type
    if "application/json" not in req.headers.get("Content-Type", ""):
        return Response(content="Unsupported Media Type", status_code=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)
    
    # Get the request body
    body = await req.json()
    
    # Parse the activity
    activity = Activity().deserialize(body)
    
    # Get authentication header
    auth_header = req.headers.get("Authorization", "")
    
    # Process the activity
    try:
        response = await ADAPTER.process_activity(activity, auth_header, bot_logic)
        if response:
            return Response(content=json.dumps(response.body), status_code=response.status)
        return Response(status_code=HTTPStatus.OK)
    except Exception as e:
        # Log any errors
        logger.error(f"Error processing message: {str(e)}")
        traceback.print_exc()
        return Response(content=str(e), status_code=HTTPStatus.INTERNAL_SERVER_ERROR)

# Simple health check endpoint
@app.get("/health")
async def health_check():
    return {"status": "ok", "service": "Teams Product Management Bot"}

# Root path redirect to health
@app.get("/")
async def root():
    return {"status": "ok", "message": "Product Management Bot is running. Use the /api/messages endpoint."}

# Run the app with uvicorn if executed directly
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
