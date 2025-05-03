import logging
import threading
import tempfile
import os
import json
import asyncio
import base64
import mimetypes
import traceback
import time
import re
import copy
import sys
from io import StringIO
from typing import Optional, List, Dict, Any, Tuple, Union, Callable, Literal
from http import HTTPStatus
from datetime import datetime

# FastAPI imports
from fastapi import FastAPI, Request, Response, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

# Azure OpenAI imports
from openai import AzureOpenAI

# Bot Framework imports
from botbuilder.core import (
    BotFrameworkAdapterSettings,
    TurnContext,
    BotFrameworkAdapter,
    CardFactory,
    MemoryStorage
)
from botbuilder.schema import (
    Activity, 
    ActivityTypes, 
    Attachment, 
    ConversationReference,
    ChannelAccount,
    ConversationAccount,
    Entity
)

from botbuilder.schema.teams import (
    FileDownloadInfo,
    FileConsentCard,
    FileConsentCardResponse,
    FileInfoCard,
)
from botbuilder.schema.teams.additional_properties import ContentType

# Teams AI imports
from teams.streaming import StreamingResponse
from teams.streaming.streaming_channel_data import StreamingChannelData
from teams.streaming.streaming_entity import StreamingEntity
from teams.ai.citations.citations import Appearance, SensitivityUsageInfo
from teams.ai.citations import AIEntity, ClientCitation
from teams.ai.prompts.message import Citation

import uuid
from typing import Dict, List, Deque
from collections import deque
import threading

# Dictionary to store pending messages for each conversation
pending_messages = {}
# Lock for thread-safe operations on the pending_messages dict
pending_messages_lock = threading.Lock()
# Dictionary for tracking active runs
active_runs = {}
# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(module)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("pmbot.log")
    ]
)
logger = logging.getLogger("pmbot")

# Azure OpenAI client configuration
AZURE_ENDPOINT = "https://kb-stellar.openai.azure.com/"  # Replace with your endpoint if different
AZURE_API_KEY = "bc0ba854d3644d7998a5034af62d03ce"  # Replace with your key if different
AZURE_API_VERSION = "2024-05-01-preview"

# App credentials from environment variables for Bot Framework
APP_ID = os.environ.get("MicrosoftAppId", "")
APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")

# Dictionary to store conversation state for each user in Teams
# Key: conversation_id, Value: dict with assistant_id, session_id, etc.
conversation_states = {}
# Add this after the conversation_states declaration
conversation_states_lock = threading.Lock()
# Simple status updates for long-running operations
operation_statuses = {}

# Create adapter with proper settings for Bot Framework
SETTINGS = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)

# Directory for file handling
FILE_DIRECTORY = "files/"
os.makedirs(FILE_DIRECTORY, exist_ok=True)

# Create FastAPI app
app = FastAPI(title="Product Management and Teams Bot")

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def create_client():
    """Creates an AzureOpenAI client instance."""
    return AzureOpenAI(
        azure_endpoint=AZURE_ENDPOINT,
        api_key=AZURE_API_KEY,
        api_version=AZURE_API_VERSION,
    )

# Create typing indicator activity for Teams
def create_typing_activity() -> Activity:
    return Activity(
        type=ActivityTypes.typing,
        channel_id="msteams"
    )

async def handle_thread_recovery(turn_context: TurnContext, state, error_message):
    """Handles recovery from thread or assistant errors with improved user isolation"""
    # Get user identity for safety checks and logging
    user_id = turn_context.activity.from_property.id if hasattr(turn_context.activity, 'from_property') else "unknown"
    conversation_id = TurnContext.get_conversation_reference(turn_context.activity).conversation.id
    
    # Increment recovery attempts (with thread safety)
    with conversation_states_lock:
        state["recovery_attempts"] = state.get("recovery_attempts", 0) + 1
        state["last_error"] = error_message
        recovery_attempts = state["recovery_attempts"]
    
    # Log recovery attempt with user context
    logging.info(f"Attempting recovery for user {user_id} (attempt #{recovery_attempts}): {error_message}")
    
    # If too many recovery attempts, suggest a fresh start
    if recovery_attempts >= 3:
        # Reset the recovery counter
        with conversation_states_lock:
            state["recovery_attempts"] = 0
        
        # Send error message with new chat card
        await turn_context.send_activity(f"I'm having trouble maintaining our conversation. Let's start a new chat session.")
        await send_new_chat_card(turn_context)
        return
    
    # ALWAYS create new resources on recovery - NEVER try to reuse existing ones
    try:
        client = create_client()
        
        # Send a message to indicate recovery
        recovery_message = "I encountered an issue with our conversation. Creating a fresh session while keeping our context."
        await turn_context.send_activity(recovery_message)
        
        # Create completely new resources
        try:
            # Create a new vector store
            vector_store = client.beta.vector_stores.create(
                name=f"recovery_user_{user_id}_convo_{conversation_id}_{int(time.time())}"
            )
            
            # Create a new assistant with a unique name
            unique_name = f"recovery_assistant_user_{user_id}_{int(time.time())}"
            assistant_obj = client.beta.assistants.create(
                name=unique_name,
                model="gpt-4o-mini",
                instructions="You are a helpful assistant recovering from a system error. Please continue the conversation naturally.",
                tools=[{"type": "file_search"}],
                tool_resources={"file_search": {"vector_store_ids": [vector_store.id]}},
            )
            
            # Create a new thread
            thread = client.beta.threads.create()
            
            # Update state with new resources (thread safe)
            with conversation_states_lock:
                old_thread = state.get("session_id")
                state["assistant_id"] = assistant_obj.id
                state["session_id"] = thread.id
                state["vector_store_id"] = vector_store.id
                state["active_run"] = False
            
            # Clear any active runs
            if old_thread in active_runs:
                del active_runs[old_thread]
            
            logging.info(f"Recovery successful for user {user_id}: Created new assistant {assistant_obj.id} and thread {thread.id}")
            
        except Exception as creation_error:
            logging.error(f"Failed to create new resources during recovery for user {user_id}: {creation_error}")
            # If we fail to create new resources, reset state and try fresh initialization
            with conversation_states_lock:
                state["assistant_id"] = None
                state["session_id"] = None
                state["vector_store_id"] = None
                state["active_run"] = False
            
            await turn_context.send_activity("I'm still having trouble. Starting completely fresh.")
            await initialize_chat(turn_context, state)
            
    except Exception as recovery_error:
        # If recovery fails, suggest a new chat
        logging.error(f"Recovery attempt failed for user {user_id}: {recovery_error}")
        await turn_context.send_activity("I'm still having trouble with our conversation. Let's start a new chat session.")
        await send_new_chat_card(turn_context)

def create_new_chat_card():
    """Creates an adaptive card for starting a new chat session"""
    card = {
        "type": "AdaptiveCard",
        "version": "1.0",
        "body": [
            {
                "type": "TextBlock",
                "text": "Start a New Conversation",
                "size": "large",
                "weight": "bolder"
            },
            {
                "type": "TextBlock",
                "text": "Click the button below to start a fresh conversation with me.",
                "wrap": True
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Start New Chat",
                "data": {
                    "action": "new_chat"
                }
            }
        ]
    }
    return CardFactory.adaptive_card(card)

async def send_new_chat_card(turn_context: TurnContext):
    """Sends a card with a button to start a new chat session"""
    reply = _create_reply(turn_context.activity)
    reply.attachments = [create_new_chat_card()]
    await turn_context.send_activity(reply)

async def handle_card_actions(turn_context: TurnContext, action_data):
    """Handles actions from adaptive cards"""
    try:
        if action_data.get("action") == "new_chat":
            # Get conversation ID
            conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
            conversation_id = conversation_reference.conversation.id
            
            # Reset conversation state
            if conversation_id in conversation_states:
                # Clear any pending messages
                with pending_messages_lock:
                    if conversation_id in pending_messages and pending_messages[conversation_id]:
                        # Now process pending messages in a smarter way
                        pending_messages[conversation_id].clear()
                
                # Send typing indicator
                await turn_context.send_activity(create_typing_activity())
                
                # Initialize new chat
                await initialize_chat(turn_context, None)  # Pass None to force new state creation
            else:
                await initialize_chat(turn_context, None)
    except Exception as e:
        logging.error(f"Error handling card action: {e}")
        await turn_context.send_activity(f"I couldn't start a new chat. Please try again later.")

# ----- Teams Bot Logic Functions -----

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

def _create_reply(activity, text=None, text_format=None):
    """Helper method to create a reply message."""
    return Activity(
        type=ActivityTypes.message,
        timestamp=datetime.utcnow(),
        from_property=ChannelAccount(id=activity.recipient.id, name=activity.recipient.name),
        recipient=ChannelAccount(id=activity.from_property.id, name=activity.from_property.name),
        reply_to_id=activity.id,
        service_url=activity.service_url,
        channel_id=activity.channel_id,
        conversation=ConversationAccount(
            is_group=activity.conversation.is_group,
            id=activity.conversation.id,
            name=activity.conversation.name,
        ),
        text=text or "",
        text_format=text_format or None,
        locale=activity.locale,
    )

# Bot logic handler
async def bot_logic(turn_context: TurnContext):
    # Get the conversation reference for later use
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_id = conversation_reference.conversation.id
    
    # Extract user identity for security validation
    user_id = turn_context.activity.from_property.id if hasattr(turn_context.activity, 'from_property') else "unknown"
    channel_id = turn_context.activity.channel_id
    tenant_id = getattr(turn_context.activity.conversation, 'tenant_id', 'unknown')
    
    # Create a user security fingerprint for verification
    user_security_fingerprint = f"{user_id}_{tenant_id}_{channel_id}"
    
    # Log incoming activity with user context
    logging.info(f"Processing activity type {turn_context.activity.type} from user {user_id} in conversation {conversation_id}")
    
    # Thread-safe state initialization
    with conversation_states_lock:
        if conversation_id not in conversation_states:
            # Create new state for this conversation
            conversation_states[conversation_id] = {
                "assistant_id": None,
                "session_id": None,
                "vector_store_id": None,
                "uploaded_files": [],
                "recovery_attempts": 0,
                "last_error": None,
                "active_run": False,
                "user_id": user_id,
                "tenant_id": tenant_id,
                "security_fingerprint": user_security_fingerprint,
                "creation_time": time.time()
            }
        else:
            # Verify user identity to prevent cross-contamination
            stored_user_id = conversation_states[conversation_id].get("user_id")
            stored_fingerprint = conversation_states[conversation_id].get("security_fingerprint")
            
            # If user mismatch detected, create fresh state
            if stored_user_id and stored_user_id != user_id:
                logging.warning(f"SECURITY ALERT: User mismatch in conversation {conversation_id}! Expected {stored_user_id}, got {user_id}")
                
                # Create fresh state to avoid cross-contamination
                conversation_states[conversation_id] = {
                    "assistant_id": None,
                    "session_id": None,
                    "vector_store_id": None,
                    "uploaded_files": [],
                    "recovery_attempts": 0,
                    "last_error": None,
                    "active_run": False,
                    "user_id": user_id,
                    "tenant_id": tenant_id,
                    "security_fingerprint": user_security_fingerprint,
                    "creation_time": time.time()
                }
                
                # Clear any pending messages for security
                with pending_messages_lock:
                    if conversation_id in pending_messages:
                        pending_messages[conversation_id].clear()
                
                logging.info(f"Created fresh state for user {user_id} in conversation {conversation_id} after user mismatch")
            elif stored_fingerprint and stored_fingerprint != user_security_fingerprint:
                logging.warning(f"SECURITY ALERT: Security fingerprint mismatch in conversation {conversation_id}!")
                
                # Update fingerprint but keep existing state if user_id matches
                # This handles cases where other attributes might change but it's still the same user
                conversation_states[conversation_id]["security_fingerprint"] = user_security_fingerprint
                logging.info(f"Updated security fingerprint for user {user_id}")
    
    # Get state after all security checks
    state = conversation_states[conversation_id]
    
    # Handle different activity types
    if turn_context.activity.type == ActivityTypes.message:
        # Initialize pending messages queue if not exists (thread-safe)
        with pending_messages_lock:
            if conversation_id not in pending_messages:
                pending_messages[conversation_id] = deque()
        
        # Check if we have text content first
        has_text = turn_context.activity.text and turn_context.activity.text.strip()
        
        # Check for file attachments
        has_file_attachments = False
        has_file_content_message = False
        file_caption = None
        
        if turn_context.activity.attachments and len(turn_context.activity.attachments) > 0:
            for attachment in turn_context.activity.attachments:
                if hasattr(attachment, 'content_type') and attachment.content_type == ContentType.FILE_DOWNLOAD_INFO:
                    has_file_attachments = True
                    # Check if there's also a message with the file (caption)
                    if has_text:
                        has_file_content_message = True
                        file_caption = turn_context.activity.text.strip()
                    break
        
        # Check for session timeout (24 hours)
        session_timeout = 86400  # 24 hours in seconds
        current_time = time.time()
        with conversation_states_lock:
            creation_time = state.get("creation_time", current_time)
            session_age = current_time - creation_time
            
            # Force session refresh if too old
            if session_age > session_timeout and state.get("session_id"):
                logging.info(f"Session timeout for user {user_id}: age={session_age}s - Creating fresh session")
                # Keep user ID but reset all resources
                state["assistant_id"] = None
                state["session_id"] = None
                state["vector_store_id"] = None
                state["uploaded_files"] = []
                state["recovery_attempts"] = 0
                state["creation_time"] = current_time
                
                # Clear any pending messages
                with pending_messages_lock:
                    if conversation_id in pending_messages:
                        pending_messages[conversation_id].clear()
                
                await turn_context.send_activity("Your previous session has expired. Creating a new session for you.")
        
        # Track if thread is currently processing (thread-safe)
        is_thread_busy = False
        with conversation_states_lock:
            is_thread_busy = state.get("active_run", False)
        
        # If thread is busy and we have a text message, queue it
        if is_thread_busy and has_text and not has_file_attachments:
            with pending_messages_lock:
                pending_messages[conversation_id].append(turn_context.activity.text.strip())
            await turn_context.send_activity("I'm still working on your previous request. I'll address this message next.")
            return
        
        # Prioritize text processing if we have text content (even if there are non-file attachments)
        if has_text and not has_file_attachments:
            try:
                await handle_text_message(turn_context, state)
            except Exception as e:
                logging.error(f"Error in handle_text_message for user {user_id}: {e}")
                traceback.print_exc()
                # Attempt recovery
                await handle_thread_recovery(turn_context, state, str(e))
        
        # Process file attachments with or without caption
        elif has_file_attachments:
            try:
                await handle_file_upload(turn_context, state, file_caption)
            except Exception as e:
                logging.error(f"Error in handle_file_upload for user {user_id}: {e}")
                traceback.print_exc()
                # Attempt recovery
                await handle_thread_recovery(turn_context, state, str(e))
        
        # Fallback for messages with neither text nor file attachments
        else:
            # This handles cases where Teams might send empty messages or special activities
            logger.info(f"Received message without text or file attachments from user {user_id}")
            
            # Retrieve current assistant_id (thread-safe)
            current_assistant_id = None
            with conversation_states_lock:
                current_assistant_id = state.get("assistant_id")
                
            if not current_assistant_id:
                try:
                    await initialize_chat(turn_context, state)
                except Exception as e:
                    logging.error(f"Error in initialize_chat for user {user_id}: {e}")
                    # Attempt recovery
                    await handle_thread_recovery(turn_context, state, str(e))
            else:
                await turn_context.send_activity("I didn't receive any text or files. How can I help you?")
    
    # Handle Teams file consent card responses
    elif turn_context.activity.type == ActivityTypes.invoke:
        if turn_context.activity.name == "fileConsent/invoke":
            await handle_file_consent_response(turn_context, turn_context.activity.value)
        elif turn_context.activity.name == "adaptiveCard/action":
            # Handle adaptive card actions (for new chat button)
            await handle_card_actions(turn_context, turn_context.activity.value)
    
    # Handle conversation update (bot added to conversation)
    elif turn_context.activity.type == ActivityTypes.conversation_update:
        if turn_context.activity.members_added:
            for member in turn_context.activity.members_added:
                if member.id != turn_context.activity.recipient.id:
                    # Bot was added - send welcome message with new chat card
                    await send_welcome_message(turn_context)

async def handle_file_consent_response(turn_context: TurnContext, file_consent_response: FileConsentCardResponse):
    """Handle file consent card response."""
    if file_consent_response.action == "accept":
        await handle_file_consent_accept(turn_context, file_consent_response)
    else:
        await handle_file_consent_decline(turn_context, file_consent_response)

async def handle_file_consent_accept(turn_context: TurnContext, file_consent_response: FileConsentCardResponse):
    """Handles file upload when the user accepts the file consent."""
    file_path = os.path.join(FILE_DIRECTORY, file_consent_response.context["filename"])
    file_size = os.path.getsize(file_path)

    headers = {
        "Content-Length": f"\"{file_size}\"",
        "Content-Range": f"bytes 0-{file_size-1}/{file_size}"
    }
    try:
        import requests
        response = requests.put(
            file_consent_response.upload_info.upload_url, open(file_path, "rb"), headers=headers
        )

        if response.status_code in [200, 201]:
            await file_upload_complete(turn_context, file_consent_response)
        else:
            await file_upload_failed(turn_context, "Unable to upload file.")
    except Exception as e:
        logger.error(f"Error uploading file to Teams: {str(e)}")
        await file_upload_failed(turn_context, f"Upload failed: {str(e)}")

async def handle_file_consent_decline(turn_context: TurnContext, file_consent_response: FileConsentCardResponse):
    """Handles file upload when the user declines the file consent."""
    filename = file_consent_response.context["filename"]
    reply = _create_reply(turn_context.activity, f"Declined. We won't upload file <b>{filename}</b>.", "xml")
    await turn_context.send_activity(reply)

async def file_upload_complete(turn_context: TurnContext, file_consent_response: FileConsentCardResponse):
    """Handles successful file upload."""
    upload_info = file_consent_response.upload_info
    download_card = FileInfoCard(
        unique_id=upload_info.unique_id,
        file_type=upload_info.file_type
    )

    attachment = Attachment(
        content=download_card.serialize(),
        content_type=ContentType.FILE_INFO_CARD,
        name=upload_info.name,
        content_url=upload_info.content_url
    )

    reply = _create_reply(turn_context.activity, f"<b>File uploaded.</b> Your file <b>{upload_info.name}</b> is ready to download", "xml")
    reply.attachments = [attachment]
    await turn_context.send_activity(reply)

async def file_upload_failed(turn_context: TurnContext, error: str):
    """Handles file upload failure."""
    reply = _create_reply(turn_context.activity, f"<b>File upload failed.</b> Error: <pre>{error}</pre>", "xml")
    await turn_context.send_activity(reply)

async def download_file(turn_context: TurnContext, attachment: Attachment):
    """Handles file download from Teams."""
    try:
        file_download = FileDownloadInfo.deserialize(attachment.content)
        file_path = os.path.join(FILE_DIRECTORY, attachment.name)

        # Ensure the file directory exists
        os.makedirs(FILE_DIRECTORY, exist_ok=True)

        import requests
        response = requests.get(file_download.download_url, allow_redirects=True)
        if response.status_code == 200:
            with open(file_path, "wb") as f:
                f.write(response.content)
                
            # Check file type and reject if necessary
            file_ext = os.path.splitext(attachment.name)[1].lower()
            if file_ext in ['.csv', '.xlsx', '.xls', '.xlsm']:
                await turn_context.send_activity("Sorry, CSV and Excel files are not supported. Please upload PDF, DOC, DOCX, or TXT files only.")
                # Delete the file
                os.remove(file_path)
                return None
                
            await turn_context.send_activity(f"Completed downloading <b>{attachment.name}</b>")
            return file_path
        else:
            await file_upload_failed(turn_context, "Download failed.")
            return None
    except Exception as e:
        logger.error(f"Error downloading file from Teams: {str(e)}")
        await file_upload_failed(turn_context, f"Download failed: {str(e)}")
        return None

# Function to handle file uploads
async def handle_file_upload(turn_context: TurnContext, state, message_text=None):
    """Handle file uploads from Teams with optional message text"""
    
    for attachment in turn_context.activity.attachments:
        try:
            # Send typing indicator
            await turn_context.send_activity(create_typing_activity())
            
            # Check if it's a file download info
            if hasattr(attachment, 'content_type') and attachment.content_type == ContentType.FILE_DOWNLOAD_INFO:
                # Download the file using the Teams-specific logic
                file_path = await download_file(turn_context, attachment)
                
                if not file_path:
                    # File was either not downloaded or rejected
                    continue
                    
                # Check file extension to ensure we only accept supported types
                file_ext = os.path.splitext(attachment.name)[1].lower()
                if file_ext in ['.csv', '.xlsx', '.xls', '.xlsm']:
                    await turn_context.send_activity("Sorry, CSV and Excel files are not supported. Please upload PDF, DOC, DOCX, or TXT files only.")
                    continue
                
                # Process the file with message text if provided
                await process_uploaded_file(turn_context, state, file_path, attachment.name, message_text)
            else:
                # Only prompt for file uploads if this is actually a file-related attachment
                # but not in the expected format (prevents the message when dealing with non-file attachments)
                file_related_types = [
                    ContentType.FILE_CONSENT_CARD,
                    ContentType.FILE_INFO_CARD,
                    "application/vnd.microsoft.teams.file."
                ]
                
                is_file_related = False
                if hasattr(attachment, 'content_type'):
                    for file_type in file_related_types:
                        if file_type in attachment.content_type:
                            is_file_related = True
                            break
                
                if is_file_related:
                    await turn_context.send_activity("Please upload a file using the file upload feature in Teams.")
                # If it's not file-related, we don't need to send any message
                
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            traceback.print_exc()
            await turn_context.send_activity(f"Error processing file: {str(e)}")

async def upload_file_to_openai_thread(client: AzureOpenAI, file_content: bytes, filename: str, thread_id: str, message_content: str = None):
    """
    Uploads a file directly to OpenAI and attaches it to a thread.
    
    Args:
        client: Azure OpenAI client
        file_content: Raw file content bytes
        filename: Name of the file
        thread_id: Thread ID to attach the file to
        message_content: Optional message content to include with the file
        
    Returns:
        Dictionary with upload result information
    """
    try:
        # Create a temporary file for upload
        with tempfile.NamedTemporaryFile(delete=False, suffix='_' + filename) as temp:
            temp.write(file_content)
            temp_path = temp.name
        
        logging.info(f"Uploading file {filename} directly to OpenAI for thread {thread_id}")
        
        try:
            # Upload the file to OpenAI
            with open(temp_path, "rb") as file_data:
                file_obj = client.files.create(
                    file=file_data,
                    purpose="assistants"
                )
            
            file_id = file_obj.id
            logging.info(f"File uploaded to OpenAI with ID: {file_id}")
            
            # Create a message with the file attachment
            message_text = message_content or f"I've uploaded a file named '{filename}'. Please analyze this file."
            
            # Create a message with the file attachment
            message = client.beta.threads.messages.create(
                thread_id=thread_id,
                role="user",
                content=message_text,
                attachments=[{
                    "file_id": file_id,
                    "tools": [{"type": "file_search"}]
                  }]
            )
            
            logging.info(f"File {filename} (ID: {file_id}) attached to thread {thread_id}")
            
            return {
                "message": f"File {filename} uploaded and attached to thread",
                "filename": filename,
                "file_id": file_id,
                "thread_id": thread_id,
                "processing_method": "thread_attachment"
            }
            
        finally:
            # Clean up temporary file
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
    except Exception as e:
        logging.error(f"Error uploading file to OpenAI: {str(e)}")
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Failed to upload file to OpenAI: {str(e)}")

async def process_uploaded_file(turn_context: TurnContext, state, file_path: str, filename: str, message_text: str = None):
    """Process an uploaded file after it's been downloaded, with optional message text"""
    # Message user that file is being processed
    if message_text:
        await turn_context.send_activity(f"Processing file: '{filename}' with your message: '{message_text}'...")
    else:
        await turn_context.send_activity(f"Processing file: '{filename}'...")
    
    # If no assistant yet, initialize chat first
    if not state["assistant_id"]:
        await initialize_chat(turn_context, state)
    
    try:
        # Read the file content
        with open(file_path, 'rb') as file:
            file_content = file.read()
            
            # Check file type
            file_ext = os.path.splitext(filename)[1].lower()
            is_image = file_ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
            is_document = file_ext in ['.pdf', '.doc', '.docx', '.txt', '.md', '.html', '.json']
            is_csv_excel = file_ext in ['.csv', '.xlsx', '.xls', '.xlsm']
            
            if is_csv_excel:
                await turn_context.send_activity("Sorry, CSV and Excel files are not supported. Please upload PDF, DOC, DOCX, or TXT files only.")
                return
            
            # Process based on file type
            client = create_client()
            
            if is_image:
                # Analyze image - same as before but add message text if provided
                analysis_text = await image_analysis_internal(file_content, filename)
                
                # Add analysis to the thread
                if state["session_id"]:
                    # Create content with analysis and optional message
                    content_text = f"Analysis result for uploaded image '{filename}':\n{analysis_text}"
                    
                    # If there's a message, include it first
                    if message_text:
                        content_text = f"User message: {message_text}\n\n{content_text}"
                        
                    client.beta.threads.messages.create(
                        thread_id=state["session_id"],
                        role="user",
                        content=content_text
                    )
                    
                    # Add image file awareness
                    await add_file_awareness_internal(
                        state["session_id"], 
                        {
                            "name": filename,
                            "type": "image",
                            "processing_method": "thread_message",
                            "with_message": message_text is not None
                        }
                    )
                    
                    await turn_context.send_activity(f"Image '{filename}' processed successfully!")
                    await turn_context.send_activity("Here's my analysis of the image:")
                    await turn_context.send_activity(analysis_text)
                else:
                    await turn_context.send_activity("Cannot process image: No active conversation session.")
                    
            elif is_document:
                # Use the new direct file upload approach for documents
                if state["assistant_id"] and state["session_id"]:
                    # Send a typing indicator
                    await turn_context.send_activity(create_typing_activity())
                    
                    # Upload the file directly to the thread
                    if message_text:
                        message_content = f"{message_text}\n\nI've also uploaded a document named '{filename}'. Please use this document to answer my questions."
                    else:
                        message_content = f"I've uploaded a document named '{filename}'. Please use this document to answer my questions."
                    
                    # Use the new OpenAI direct file upload approach
                    try:
                        result = await upload_file_to_openai_thread(
                            client,
                            file_content,
                            filename,
                            state["session_id"],
                            message_content
                        )
                        
                        # Add to the list of uploaded files
                        state["uploaded_files"].append(filename)
                        
                        # Add file awareness to the thread
                        await add_file_awareness_internal(
                            state["session_id"],
                            {
                                "name": filename,
                                "type": file_ext[1:] if file_ext else "document",
                                "processing_method": "thread_attachment",
                                "with_message": message_text is not None
                            }
                        )
                        
                        await turn_context.send_activity(f"Document '{filename}' uploaded successfully! You can now ask questions about it.")
                        
                    except Exception as upload_error:
                        logger.error(f"Error uploading file to OpenAI: {str(upload_error)}")
                        await turn_context.send_activity(f"Error uploading document: {str(upload_error)}")
                        
                        # Fall back to vector store approach if direct upload fails
                        logger.info(f"Falling back to vector store approach for document '{filename}'")
                        
                        # Create a temporary file for vector store upload
                        with tempfile.NamedTemporaryFile(delete=False, suffix='_' + filename) as temp:
                            temp.write(file_content)
                            temp_path = temp.name
                        
                        try:
                            # Get current vector store ID
                            vector_store_id = state["vector_store_id"]
                            if not vector_store_id:
                                # Create a new vector store if needed
                                vector_store = client.beta.vector_stores.create(name=f"Assistant_{state['assistant_id']}_Store")
                                vector_store_id = vector_store.id
                                state["vector_store_id"] = vector_store_id
                            
                            # Upload to vector store
                            with open(temp_path, "rb") as file_stream:
                                file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
                                    vector_store_id=vector_store_id,
                                    files=[file_stream]
                                )
                            
                            # Add user message separately if provided
                            if message_text:
                                client.beta.threads.messages.create(
                                    thread_id=state["session_id"],
                                    role="user",
                                    content=message_text
                                )
                            
                            # Update assistant with file_search tool
                            assistant_obj = client.beta.assistants.retrieve(assistant_id=state["assistant_id"])
                            has_file_search = False
                            
                            for tool in assistant_obj.tools:
                                if hasattr(tool, 'type') and tool.type == "file_search":
                                    has_file_search = True
                                    break
                            
                            if not has_file_search:
                                current_tools = list(assistant_obj.tools)
                                current_tools.append({"type": "file_search"})
                                
                                client.beta.assistants.update(
                                    assistant_id=state["assistant_id"],
                                    tools=current_tools,
                                    tool_resources={"file_search": {"vector_store_ids": [vector_store_id]}}
                                )
                            
                            # Add file awareness
                            await add_file_awareness_internal(
                                state["session_id"],
                                {
                                    "name": filename,
                                    "type": file_ext[1:] if file_ext else "document",
                                    "processing_method": "vector_store",
                                    "with_message": message_text is not None
                                }
                            )
                            
                            # Add to the list of uploaded files
                            state["uploaded_files"].append(filename)
                            
                            await turn_context.send_activity(f"Document '{filename}' uploaded to vector store successfully as a fallback method. You can now ask questions about it.")
                            
                        except Exception as fallback_error:
                            await turn_context.send_activity(f"Failed to upload document via fallback method: {str(fallback_error)}")
                        finally:
                            # Clean up temp file
                            if os.path.exists(temp_path):
                                os.remove(temp_path)
                else:
                    await turn_context.send_activity("Cannot process document: No active assistant or session.")
            else:
                await turn_context.send_activity(f"Unsupported file type: {file_ext}. Please upload PDF, DOC, DOCX, TXT files, or images.")
    except Exception as e:
        logger.error(f"Error processing file '{filename}': {e}")
        traceback.print_exc()
        await turn_context.send_activity(f"Error processing file: {str(e)}")
    finally:
        # Clean up file
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except OSError as e:
                logger.error(f"Error removing file {file_path}: {e}")

# Function to send file to Teams
async def send_file_to_teams(turn_context: TurnContext, filename: str):
    """Sends a file to the user in Teams using file consent card."""
    file_path = os.path.join(FILE_DIRECTORY, filename)
    if not os.path.exists(file_path):
        await file_upload_failed(turn_context, "File not found.")
        return

    file_size = os.path.getsize(file_path)
    consent_context = {"filename": filename}

    file_card = FileConsentCard(
        description="This is the file I want to send you",
        size_in_bytes=file_size,
        accept_context=consent_context,
        decline_context=consent_context
    )

    attachment = Attachment(
        content=file_card.serialize(),
        content_type=ContentType.FILE_CONSENT_CARD,
        name=filename
    )

    reply = _create_reply(turn_context.activity)
    reply.attachments = [attachment]
    await turn_context.send_activity(reply)

# Thread summarization helper function
async def summarize_thread_if_needed(client: AzureOpenAI, thread_id: str, state: dict, threshold: int = 30):
    """
    Checks if a thread needs summarization and performs the summarization if necessary.
    
    Args:
        client: Azure OpenAI client
        thread_id: The thread ID to check
        state: The conversation state dictionary
        threshold: Message count threshold before summarization (default: 30)
    
    Returns:
        bool: True if summarization was performed, False otherwise
    """
    try:
        # Check if we've already summarized recently
        last_summarization = state.get("last_summarization_time", 0)
        current_time = time.time()
        
        # Don't summarize more often than every 10 minutes
        if current_time - last_summarization < 600:  # 600 seconds = 10 minutes
            return False
            
        # Retrieve thread messages
        messages = client.beta.threads.messages.list(
            thread_id=thread_id,
            order="asc",
            limit=100  # Get up to 100 messages to check count
        )
        
        # Count messages
        message_count = len(messages.data)
        logging.info(f"Thread {thread_id} has {message_count} messages")
        
        # If below threshold, no need to summarize
        if message_count < threshold:
            return False
            
        # Determine how many messages to summarize (leave 5-10 recent messages untouched)
        messages_to_keep = 7  # Keep the 7 most recent messages
        messages_to_summarize = message_count - messages_to_keep
        
        if messages_to_summarize <= 5:  # Not worth summarizing if too few
            return False
            
        # Get messages to summarize (all except the most recent)
        messages_list = list(messages.data)
        messages_to_summarize_list = messages_list[:-messages_to_keep]
        
        # Convert messages to a format suitable for summarization
        conversation_text = ""
        for msg in messages_to_summarize_list:
            role = "User" if msg.role == "user" else "Assistant"
            
            # Extract the text content from the message
            content_text = ""
            for content_part in msg.content:
                if content_part.type == 'text':
                    content_text += content_part.text.value
            
            conversation_text += f"{role}: {content_text}\n\n"
        
        # If we have a very long conversation, we need to be selective
        if len(conversation_text) > 12000:  # Truncate if too long
            conversation_text = conversation_text[:4000] + "\n...[middle of conversation omitted]...\n" + conversation_text[-8000:]
        
        # Create a new thread for summarization to avoid conflicts
        summary_thread = client.beta.threads.create()
        
        # Add the conversation to summarize
        client.beta.threads.messages.create(
            thread_id=summary_thread.id,
            role="user",
            content=f"Please create a concise but comprehensive summary of the following conversation. Focus on key points, decisions, and important context that would be needed for continuing the conversation effectively:\n\n{conversation_text}"
        )
        
        # Run the summarization with a different assistant
        summary_run = client.beta.threads.runs.create(
            thread_id=summary_thread.id,
            assistant_id=state["assistant_id"],  # Use the same assistant
            instructions="Create a concise but comprehensive summary of the conversation provided. Focus on extracting key points, decisions, and important context that would be needed for continuing the conversation effectively. Format the summary in clear sections with bullet points where appropriate."
        )
        
        # Wait for completion
        max_wait = 60  # Maximum wait time in seconds
        wait_interval = 2  # Check interval in seconds
        elapsed = 0
        
        while elapsed < max_wait:
            run_status = client.beta.threads.runs.retrieve(
                thread_id=summary_thread.id,
                run_id=summary_run.id
            )
            
            if run_status.status == "completed":
                # Get the summary
                summary_messages = client.beta.threads.messages.list(
                    thread_id=summary_thread.id,
                    order="desc",
                    limit=1
                )
                
                if summary_messages.data:
                    # Extract the summary text
                    summary_text = ""
                    for content_part in summary_messages.data[0].content:
                        if content_part.type == 'text':
                            summary_text += content_part.text.value
                    
                    # Create a new thread with the summary as context
                    new_thread = client.beta.threads.create()
                    
                    # Add the summary as a system message in the new thread
                    client.beta.threads.messages.create(
                        thread_id=new_thread.id,
                        role="user",
                        content=f"CONVERSATION SUMMARY: {summary_text}\n\nPlease acknowledge this conversation summary and continue the conversation based on this context.",
                        metadata={"type": "conversation_summary"}
                    )
                    
                    # Get a response acknowledging the summary
                    acknowledgement_run = client.beta.threads.runs.create(
                        thread_id=new_thread.id,
                        assistant_id=state["assistant_id"]
                    )
                    
                    # Wait for acknowledgement
                    await asyncio.sleep(5)
                    
                    # Add the most recent messages to the new thread to maintain continuity
                    for recent_msg in messages_list[-messages_to_keep:]:
                        # Extract content
                        content_text = ""
                        for content_part in recent_msg.content:
                            if content_part.type == 'text':
                                content_text += content_part.text.value
                        
                        # Add to new thread
                        client.beta.threads.messages.create(
                            thread_id=new_thread.id,
                            role=recent_msg.role,
                            content=content_text
                        )
                    
                    # Update the state with the new thread ID
                    old_thread_id = state["session_id"]
                    state["session_id"] = new_thread.id
                    state["last_summarization_time"] = current_time
                    state["active_run"] = False
                    
                    # Update active_runs dictionary
                    if old_thread_id in active_runs:
                        del active_runs[old_thread_id]
                    
                    logging.info(f"Summarized thread {old_thread_id} and created new thread {new_thread.id}")
                    return True
            
            elif run_status.status in ["failed", "cancelled", "expired"]:
                logging.error(f"Summary generation failed with status: {run_status.status}")
                return False
            
            await asyncio.sleep(wait_interval)
            elapsed += wait_interval
        
        logging.warning(f"Summary generation timed out after {max_wait} seconds")
        return False
        
    except Exception as e:
        logging.error(f"Error summarizing thread: {str(e)}")
        traceback.print_exc()
        return False

# Modified handle_text_message with thread summarization
async def handle_text_message(turn_context: TurnContext, state):
    user_message = turn_context.activity.text.strip()
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_id = conversation_reference.conversation.id
    
    # Extract user identity for security validation
    user_id = turn_context.activity.from_property.id if hasattr(turn_context.activity, 'from_property') else "unknown"
    
    # Thread-safe access to state values
    with conversation_states_lock:
        stored_user_id = state.get("user_id")
        stored_assistant_id = state.get("assistant_id")
        stored_session_id = state.get("session_id")
    
    # Verify user identity matches state (double-check)
    if stored_user_id and stored_user_id != user_id:
        logging.warning(f"SECURITY ALERT: User mismatch detected in handle_text_message! Expected {stored_user_id}, got {user_id}")
        # This is a severe security issue - reinitialize chat for this user
        await turn_context.send_activity("For security reasons, I need to create a new conversation session.")
        await initialize_chat(turn_context, None, context=user_message)
        return
    
    # Record this user's message processing (audit trail)
    logging.info(f"Processing message from user {user_id} in conversation {conversation_id}: {user_message[:50]}...")
    
    # If no assistant yet, initialize chat with the message as context
    if not stored_assistant_id:
        await initialize_chat(turn_context, state, context=user_message)
        return
    
    # Send typing indicator
    await turn_context.send_activity(create_typing_activity())
    
    # Check if thread needs summarization (with thread safety)
    summarized = False
    if stored_session_id:
        client = create_client()
        summarized = await summarize_thread_if_needed(client, stored_session_id, state, threshold=30)
        
        if summarized:
            # Update stored_session_id after summarization (thread may have changed)
            with conversation_states_lock:
                stored_session_id = state.get("session_id")
                
            await turn_context.send_activity("I've summarized our previous conversation to maintain context while keeping the conversation focused.")
    
    # Mark thread as busy (thread-safe)
    with conversation_states_lock:
        state["active_run"] = True
        current_session_id = state.get("session_id")
    
    if current_session_id:
        active_runs[current_session_id] = True
    
    try:
        # Double-verify resources before proceeding
        client = create_client()
        validation = await validate_resources(client, current_session_id, stored_assistant_id)
        
        # If any resource is invalid, force recovery
        if not validation["thread_valid"] or not validation["assistant_valid"]:
            logging.warning(f"Resource validation failed for user {user_id}: thread_valid={validation['thread_valid']}, assistant_valid={validation['assistant_valid']}")
            raise Exception("Invalid conversation resources detected - forcing recovery")
            
        # Use streaming if supported by the channel
        supports_streaming = turn_context.activity.channel_id == "msteams"
        
        if supports_streaming:
            # Use enhanced streaming with Teams AI library
            await stream_response_with_teams_ai(turn_context, state, user_message)
        else:
            # Call the internal function directly without HTTP calls
            result = await process_conversation_internal(
                client=client,
                session=current_session_id,
                prompt=user_message,
                assistant=stored_assistant_id,
                stream_output=False
            )
            
            # Extract text from response
            if isinstance(result, dict) and "response" in result:
                assistant_response = result["response"]
            else:
                assistant_response = "I'm sorry, I couldn't process your request."
                
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
        
        # Mark thread as no longer busy (thread-safe)
        with conversation_states_lock:
            state["active_run"] = False
            current_session_id = state.get("session_id")
        
        if current_session_id in active_runs:
            del active_runs[current_session_id]
        
        # Process any pending messages
        with pending_messages_lock:
            if conversation_id in pending_messages and pending_messages[conversation_id]:
                # Process queued messages
                next_messages = list(pending_messages[conversation_id])
                pending_messages[conversation_id].clear()
                
                if len(next_messages) == 1:
                    # Handle single follow-up message
                    next_message = next_messages[0]
                    await turn_context.send_activity("Now addressing your follow-up message...")
                    
                    # Instead of creating a new context, use the existing one with modified text
                    original_text = turn_context.activity.text
                    turn_context.activity.text = next_message
                    
                    try:
                        # Process with the same context - just different text
                        await handle_text_message(turn_context, state)
                    finally:
                        # Restore original text when done
                        turn_context.activity.text = original_text
                else:
                    # Process multiple messages as a batch
                    await turn_context.send_activity(f"Now addressing your {len(next_messages)} follow-up messages together...")
                    
                    # Use existing context with combined messages
                    original_text = turn_context.activity.text
                    combined_message = "\n\n".join([f"Question {i+1}: {msg}" for i, msg in enumerate(next_messages)])
                    turn_context.activity.text = f"Please answer all of these questions:\n{combined_message}"
                    
                    try:
                        # Process with the same context
                        await handle_text_message(turn_context, state)
                    finally:
                        # Restore original text
                        turn_context.activity.text = original_text
            
    except Exception as e:
        # Mark thread as no longer busy even on error (thread-safe)
        with conversation_states_lock:
            state["active_run"] = False
            current_session_id = state.get("session_id")
            
        if current_session_id in active_runs:
            del active_runs[current_session_id]
            
        # Don't show raw error details to users
        logging.error(f"Error in handle_text_message for user {user_id}: {str(e)}")
        traceback.print_exc()
        await turn_context.send_activity("I'm sorry, I encountered a problem while processing your message. Please try again.")

# Implement streaming with Teams AI integration
async def stream_response_with_teams_ai(turn_context: TurnContext, state, user_message):
    """
    Stream responses using the Teams AI library's StreamingResponse class.
    
    Args:
        turn_context: The TurnContext object
        state: The conversation state
        user_message: The user's message
    """
    try:
        client = create_client()
        thread_id = state["session_id"]
        assistant_id = state["assistant_id"]
        
        # Mark run as active in state
        state["active_run"] = True
        active_runs[thread_id] = True
        
        # Create a StreamingResponse instance from Teams AI
        streamer = StreamingResponse(turn_context)
        
        # Send initial informative update
        streamer.queue_informative_update("Processing your request...")
        
        try:
            # First, add the user message to the thread
            if user_message:
                try:
                    # Check for any existing active runs
                    try:
                        runs = client.beta.threads.runs.list(thread_id=thread_id, limit=1)
                        if runs.data:
                            latest_run = runs.data[0]
                            if latest_run.status in ["in_progress", "queued", "requires_action"]:
                                active_run_id = latest_run.id
                                logging.info(f"Found existing active run {active_run_id} with status {latest_run.status}")
                                
                                # Cancel the active run
                                client.beta.threads.runs.cancel(thread_id=thread_id, run_id=active_run_id)
                                logging.info(f"Requested cancellation of pre-existing run {active_run_id}")
                                
                                # Wait briefly for cancellation to take effect
                                await asyncio.sleep(2)
                    except Exception as check_e:
                        logging.warning(f"Error checking for existing runs: {check_e}")
                    
                    # Add the user message to the thread
                    client.beta.threads.messages.create(
                        thread_id=thread_id,
                        role="user",
                        content=user_message
                    )
                    logging.info(f"Added user message to thread {thread_id}")
                    
                except Exception as msg_e:
                    if "while a run" in str(msg_e):
                        logging.warning(f"Could not add message due to active run: {msg_e}")
                        # Create a new thread as fallback
                        new_thread = client.beta.threads.create()
                        thread_id = new_thread.id
                        state["session_id"] = thread_id
                        logging.info(f"Created new thread {thread_id} due to message add failure")
                        
                        # Add message to the new thread
                        client.beta.threads.messages.create(
                            thread_id=thread_id,
                            role="user",
                            content=user_message
                        )
                        logging.info(f"Added user message to new thread {thread_id}")
                    else:
                        logging.error(f"Error adding message to thread: {msg_e}")
                        await streamer.end_stream()
                        await turn_context.send_activity("I'm having trouble processing your request. Please try again.")
                        return
            
            # Create a run for the assistant
            run = client.beta.threads.runs.create(
                thread_id=thread_id,
                assistant_id=assistant_id
            )
            run_id = run.id
            
            # Poll for run completion with streaming updates
            max_wait_time = 120  # seconds
            wait_interval = 1.5  # seconds
            elapsed_time = 0
            last_message_length = 0
            completed = False
            
            # Send an initial thinking update
            streamer.queue_informative_update("Thinking...")
            
            # Main polling loop
            while elapsed_time < max_wait_time and not completed:
                # Get run status
                run_status = client.beta.threads.runs.retrieve(
                    thread_id=thread_id,
                    run_id=run_id
                )
                
                if run_status.status == "completed":
                    # Run completed, get the final message
                    messages = client.beta.threads.messages.list(
                        thread_id=thread_id,
                        order="desc",
                        limit=1
                    )
                    
                    if messages.data:
                        latest_message = messages.data[0]
                        message_text = ""
                        
                        # Extract text content
                        for content_part in latest_message.content:
                            if content_part.type == 'text':
                                message_text += content_part.text.value
                        
                        # Queue the full text as the final chunk
                        streamer.queue_text_chunk(message_text)
                    
                    completed = True
                    break
                
                elif run_status.status in ["failed", "cancelled", "expired"]:
                    streamer.queue_text_chunk(f"I encountered an error with status: {run_status.status}. Please try again.")
                    completed = True
                    break
                
                elif run_status.status == "requires_action":
                    # Handle function calling if needed
                    streamer.queue_informative_update("Performing additional actions...")
                    # TODO: Implement function calling logic if needed
                    
                elif run_status.status == "in_progress":
                    # Attempt to show partial updates every few seconds
                    try:
                        # Get the latest message being generated
                        messages = client.beta.threads.messages.list(
                            thread_id=thread_id,
                            order="desc",
                            limit=1
                        )
                        
                        if messages.data and messages.data[0].role == "assistant":
                            latest_message = messages.data[0]
                            current_text = ""
                            
                            # Extract text content
                            for content_part in latest_message.content:
                                if content_part.type == 'text':
                                    current_text += content_part.text.value
                            
                            # Only queue an update if we have new content
                            if len(current_text) > last_message_length:
                                # Queue only the new part of the message
                                new_content = current_text[last_message_length:]
                                streamer.queue_text_chunk(new_content)
                                last_message_length = len(current_text)
                        else:
                            # No assistant message yet, send a typing indicator
                            streamer.queue_informative_update("Still working...")
                    except Exception as stream_e:
                        logging.warning(f"Error getting partial updates: {stream_e}")
                        # Continue with polling; this is non-fatal
                
                # Wait before next poll, respecting Teams rate limits
                await asyncio.sleep(wait_interval)
                elapsed_time += wait_interval
            
            # If we didn't complete within the time limit
            if not completed:
                streamer.queue_text_chunk("\n\nI'm still working on your request but it's taking longer than expected. Here's what I have so far, and I'll continue processing in the background.")
            
            # Enable feedback loop for the final message
            streamer.set_feedback_loop(True)
            streamer.set_generated_by_ai_label(True)
            
            # End the stream
            await streamer.end_stream()
            
        except Exception as e:
            logging.error(f"Error in streaming: {e}")
            traceback.print_exc()
            
            try:
                # Try to end the stream gracefully
                streamer.queue_text_chunk("I'm sorry, I encountered an error while processing your request.")
                await streamer.end_stream()
            except Exception as stream_end_error:
                # If that fails, just send a direct message
                logging.error(f"Error ending stream: {stream_end_error}")
                await turn_context.send_activity("I encountered an error while processing your request. Please try again.")
        
        finally:
            # Ensure run is marked as complete
            with conversation_states_lock:
                state["active_run"] = False
            if thread_id in active_runs:
                del active_runs[thread_id]
    
    except Exception as outer_e:
        logging.error(f"Outer error in stream_response_with_teams_ai: {str(outer_e)}")
        traceback.print_exc()
        
        # Send a user-friendly error message
        await turn_context.send_activity("I encountered a problem while processing your request. Please try again or start a new chat.")
        
        # Mark as complete
        with conversation_states_lock:
            state["active_run"] = False
        if state.get("session_id") in active_runs:
            del active_runs[state.get("session_id", "")]

async def validate_resources(client: AzureOpenAI, thread_id: Optional[str], assistant_id: Optional[str]) -> Dict[str, bool]:
    """
    Validates that the given thread_id and assistant_id exist and are accessible.
    
    Args:
        client (AzureOpenAI): The Azure OpenAI client instance
        thread_id (Optional[str]): The thread ID to validate, or None
        assistant_id (Optional[str]): The assistant ID to validate, or None
        
    Returns:
        Dict[str, bool]: Dictionary with "thread_valid" and "assistant_valid" flags
    """
    result = {
        "thread_valid": False,
        "assistant_valid": False
    }
    
    # Validate thread if provided
    if thread_id:
        try:
            # Attempt to retrieve thread
            thread = client.beta.threads.retrieve(thread_id=thread_id)
            result["thread_valid"] = True
            logging.info(f"Thread validation: {thread_id} is valid")
        except Exception as e:
            result["thread_valid"] = False
            logging.warning(f"Thread validation: {thread_id} is invalid - {str(e)}")
    
    # Validate assistant if provided
    if assistant_id:
        try:
            # Attempt to retrieve assistant
            assistant = client.beta.assistants.retrieve(assistant_id=assistant_id)
            result["assistant_valid"] = True
            logging.info(f"Assistant validation: {assistant_id} is valid")
        except Exception as e:
            result["assistant_valid"] = False
            logging.warning(f"Assistant validation: {assistant_id} is invalid - {str(e)}")
    
    return result

async def image_analysis_internal(image_data: bytes, filename: str, prompt: Optional[str] = None) -> str:
    """Analyzes an image using Azure OpenAI vision capabilities and returns the analysis text."""
    try:
        client = create_client()
        ext = os.path.splitext(filename)[1].lower()
        b64_img = base64.b64encode(image_data).decode("utf-8")
        # Try guessing mime type, default to jpeg if extension isn't standard or determinable
        mime, _ = mimetypes.guess_type(filename)
        if not mime or not mime.startswith('image'):
            mime = f"image/{ext[1:]}" if ext and ext[1:] in ['jpg', 'jpeg', 'png', 'gif', 'webp'] else "image/jpeg"

        data_url = f"data:{mime};base64,{b64_img}"

        default_prompt = (
            "Analyze this image and provide a thorough summary including all elements. "
            "If there's any text visible, include all the textual content. Describe:"
        )
        combined_prompt = f"{default_prompt} {prompt}" if prompt else default_prompt

        # Use the existing client
        response = client.chat.completions.create(
            model="gpt-4o-mini",  # Ensure this model supports vision
            messages=[{
                "role": "user",
                "content": [
                    {"type": "text", "text": combined_prompt},
                    {"type": "image_url", "image_url": {"url": data_url, "detail": "high"}}
                ]
            }],
            max_tokens=5000  # Increased max_tokens for potentially more detailed analysis
        )

        analysis_text = response.choices[0].message.content
        return analysis_text if analysis_text else "No analysis content received."

    except Exception as e:
        logging.error(f"Image analysis error for {filename}: {e}")
        return f"Error analyzing image '{filename}': {str(e)}"

# Helper function to update user persona context
async def update_context_internal(client: AzureOpenAI, thread_id: str, context: str):
    """Updates the user persona context in a thread by adding/replacing a special message."""
    if not context:
        return

    try:
        # Get existing messages to check for previous context
        messages = client.beta.threads.messages.list(
            thread_id=thread_id,
            order="desc",
            limit=20  # Check recent messages is usually sufficient
        )

        # Look for previous context messages to avoid duplication
        previous_context_message_id = None
        for msg in messages.data:
            if hasattr(msg, 'metadata') and msg.metadata and msg.metadata.get('type') == 'user_persona_context':
                previous_context_message_id = msg.id
                break

        # If found, delete previous context message to replace it
        if previous_context_message_id:
            try:
                client.beta.threads.messages.delete(
                    thread_id=thread_id,
                    message_id=previous_context_message_id
                )
                logging.info(f"Deleted previous context message {previous_context_message_id} in thread {thread_id}")
            except Exception as e:
                logging.error(f"Error deleting previous context message {previous_context_message_id}: {e}")
            # Continue even if delete fails to add the new context

        # Add new context message
        client.beta.threads.messages.create(
            thread_id=thread_id,
            role="user",
            content=f"USER PERSONA CONTEXT: {context}",
            metadata={"type": "user_persona_context"}
        )

        logging.info(f"Updated user persona context in thread {thread_id}")
    except Exception as e:
        logging.error(f"Error updating context in thread {thread_id}: {e}")
        # Continue the flow even if context update fails

# Function to add file awareness to the assistant
async def add_file_awareness_internal(thread_id: str, file_info: Dict[str, Any]):
    """Adds file awareness to the assistant by sending a message about the file."""
    if not file_info:
        return

    try:
        client = create_client()
        
        # Create a message that informs the assistant about the file
        file_type = file_info.get("type", "unknown")
        file_name = file_info.get("name", "unnamed_file")
        processing_method = file_info.get("processing_method", "")

        awareness_message = f"FILE INFORMATION: A file named '{file_name}' of type '{file_type}' has been uploaded and processed. "

        if processing_method == "thread_message":
            awareness_message += "This image has been analyzed and the descriptive content has been added to this thread."
        elif processing_method == "vector_store":
            awareness_message += "This file has been added to the vector store and its content is available for search."
        else:
            awareness_message += "This file has been processed."

        # Send the message to the thread
        client.beta.threads.messages.create(
            thread_id=thread_id,
            role="user",  # Sending as user so assistant 'sees' it as input/instruction
            content=awareness_message,
            metadata={"type": "file_awareness", "processed_file": file_name}
        )

        logging.info(f"Added file awareness for '{file_name}' ({processing_method}) to thread {thread_id}")
    except Exception as e:
        logging.error(f"Error adding file awareness for '{file_name}' to thread {thread_id}: {e}")
        # Continue the flow even if adding awareness fails

# Initialize chat with the backend
async def initialize_chat(turn_context: TurnContext, state=None, context=None):
    """Initialize a new chat session with the backend - with improved user isolation"""
    # Get the conversation reference including user identity information
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_id = conversation_reference.conversation.id
    user_id = turn_context.activity.from_property.id if hasattr(turn_context.activity, 'from_property') else None
    
    # Create a unique identifier that includes both conversation and user 
    unique_user_key = f"{conversation_id}_{user_id}" if user_id else conversation_id
    
    # Thread-safe state initialization
    if state is None:
        with conversation_states_lock:
            conversation_states[conversation_id] = {
                "assistant_id": None,
                "session_id": None,
                "vector_store_id": None,
                "uploaded_files": [],
                "recovery_attempts": 0,
                "last_error": None,
                "active_run": False,
                "user_id": user_id,  # Store the user ID for additional verification
                "creation_time": time.time()
            }
            state = conversation_states[conversation_id]
            
            # Clear any pending messages
            with pending_messages_lock:
                if conversation_id in pending_messages:
                    pending_messages[conversation_id].clear()
    
    try:
        # Always verify user before proceeding
        if user_id and state.get("user_id") and state.get("user_id") != user_id:
            logging.warning(f"User mismatch detected! Expected {state.get('user_id')}, got {user_id}")
            # Create a fresh state since this appears to be a different user with same conversation ID
            with conversation_states_lock:
                conversation_states[conversation_id] = {
                    "assistant_id": None,
                    "session_id": None, 
                    "vector_store_id": None,
                    "uploaded_files": [],
                    "recovery_attempts": 0,
                    "last_error": None,
                    "active_run": False,
                    "user_id": user_id,
                    "creation_time": time.time()
                }
                state = conversation_states[conversation_id]
                
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Log initialization attempt with user details for traceability
        logger.info(f"Initializing chat for user {user_id} in conversation {conversation_id} with context: {context}")
        
        # Define system prompt here instead of relying on external variable
        system_prompt = '''
You are a Product Management AI Co-Pilot that helps create documentation and analyze various file types. Your capabilities vary based on the type of files uploaded.

### Understanding File Types and Processing Methods:

1. **Documents (PDF, DOC, TXT, etc.)** - When users upload these files, you should:
   - Use your file_search capability to extract relevant information
   - Quote information directly from the documents when answering questions
   - Always reference the specific filename when sharing information from a document

2. **Images** - When users upload images, you should:
   - Refer to the analysis that was automatically added to the conversation
   - Use details from the image analysis to answer questions
   - Acknowledge when information might not be visible in the image

3. **Unsupported File Types**:
   - CSV and Excel files are not supported by this system
   - If users ask about analyzing spreadsheets, kindly inform them that this feature is not available

### PRD Generation Excellence:

When creating a PRD (Product Requirements Document), develop a comprehensive and professional document with these mandatory sections:

1. **Product Overview:**
   - Product Manager: [Name and contact details]
   - Product Name: [Clear, concise name]
   - Date: [Current date and version]
   - Vision Statement: [Compelling, aspirational vision in 1-2 sentences]

2. **Problem and Customer Analysis:**
   - Customer Problem: [Clearly articulated problem statement]
   - Market Opportunity: [Quantified TAM/SAM/SOM when possible]
   - Personas: [Detailed primary and secondary user personas]
   - User Stories: [Key scenarios from persona perspective]

3. **Strategic Elements:**
   - Executive Summary: [Brief overview of product and value proposition]
   - Business Objectives: [Measurable goals with KPIs]
   - Success Metrics: [Specific metrics to track success]

4. **Detailed Requirements:**
   - Key Features: [Prioritized feature list with clear descriptions]
   - Functional Requirements: [Detailed specifications for each feature]
   - Non-Functional Requirements: [Performance, security, scalability, etc.]
   - Technical Specifications: [Relevant architecture and technical details]

5. **Implementation Planning:**
   - Milestones: [Phased delivery timeline with key dates]
   - Dependencies: [Internal and external dependencies]
   - Risks and Mitigations: [Potential challenges and contingency plans]

6. **Appendices:**
   - Supporting Documents: [Research findings, competitive analysis, etc.]
   - Open Questions: [Items requiring further investigation]

If any information is unavailable, clearly mark sections as "[To be determined]" and request specific clarification from the user. When creating a PRD, maintain a professional, clear, and structured format with appropriate headers and bullet points.

### Professional Assistance Guidelines:

- Demonstrate expertise and professionalism in all responses
- Proactively seek clarification when details are missing or ambiguous
- Ask specific questions about file names, requirements, or expectations when needed
- Provide context for why you need certain information to deliver better results
- Structure responses clearly with appropriate formatting for readability
- Always reference files by their exact filenames
- Use tools appropriately based on file type
- If asked about CSV/Excel data analysis, politely explain this is not supported
- Acknowledge limitations and be transparent when information is unavailable
- Balance detail with conciseness based on the user's needs
- When in doubt about requirements, ask targeted questions rather than making assumptions

Remember to be thorough yet efficient with your responses, anticipating follow-up needs while addressing the immediate question.
'''
        
        # ALWAYS create a new assistant and thread for this user - never reuse
        client = create_client()
        
        # Create a vector store
        try:
            vector_store = client.beta.vector_stores.create(
                name=f"user_{user_id}_convo_{conversation_id}_{int(time.time())}"
            )
            logging.info(f"Created vector store: {vector_store.id} for user {user_id}")
        except Exception as e:
            logging.error(f"Failed to create vector store for user {user_id}: {e}")
            raise HTTPException(status_code=500, detail="Failed to create vector store")

        # Include file_search tool
        assistant_tools = [{"type": "file_search"}]
        assistant_tool_resources = {
            "file_search": {"vector_store_ids": [vector_store.id]}
        }

        # Create the assistant with a unique name including user identifiers
        try:
            unique_name = f"pm_copilot_user_{user_id}_convo_{conversation_id}_{int(time.time())}"
            assistant = client.beta.assistants.create(
                name=unique_name,
                model="gpt-4o-mini",
                instructions=system_prompt,  # Now correctly defined
                tools=assistant_tools,
                tool_resources=assistant_tool_resources,
            )
            logging.info(f'Created assistant {assistant.id} for user {user_id}')
        except Exception as e:
            logging.error(f"Failed to create assistant for user {user_id}: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to create assistant: {e}")

        # Create a thread
        try:
            thread = client.beta.threads.create()
            logging.info(f"Created thread {thread.id} for user {user_id}")
        except Exception as e:
            logging.error(f"Failed to create thread for user {user_id}: {e}")
            raise HTTPException(status_code=500, detail=f"Failed to create thread: {e}")

        # Update state with new resources
        with conversation_states_lock:
            state["assistant_id"] = assistant.id
            state["session_id"] = thread.id
            state["vector_store_id"] = vector_store.id
            state["active_run"] = False
            state["recovery_attempts"] = 0
            state["user_identifier"] = unique_user_key  # Store the unique key for verification
            
        # If context is provided, add it as user persona context
        if context:
            await update_context_internal(client, thread.id, context)
            
        # Tell the user chat was initialized
        await turn_context.send_activity("Hi! I'm the Product Management Bot. I'm ready to help you with your product management tasks.")
        
        if context:
            await turn_context.send_activity(f"I've initialized with your context: '{context}'")
            # Also send the first response
            await send_message(turn_context, state)
            
    except Exception as e:
        await turn_context.send_activity(f"Error initializing chat: {str(e)}")
        logger.error(f"Error in initialize_chat for user {user_id}: {str(e)}")
        traceback.print_exc()

# Send a message without user input (used after file upload or initialization)
async def send_message(turn_context: TurnContext, state):
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Use streaming if supported by the channel
        supports_streaming = turn_context.activity.channel_id == "msteams"
        
        if supports_streaming:
            # Use streaming for response
            await stream_response_with_teams_ai(turn_context, state, None)
        else:
            # Call internal function directly to get latest message
            client = create_client()
            result = await process_conversation_internal(
                client=client,
                session=state["session_id"],
                assistant=state["assistant_id"],
                prompt=None,
                stream_output=False
            )
            
            if isinstance(result, dict) and "response" in result:
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
        "- Answer questions about uploaded documents (PDF, DOC, TXT)\n"
        "- Analyze images and provide insights\n\n"
        "To get started, you can:\n"
        "- Send me a message with your request\n"
        "- Upload a document for analysis\n"
        "- Ask me to create a PRD\n\n"
        "Note: CSV and Excel files are not supported.\n\n"
        "How can I assist you today?"
    )
    
    await turn_context.send_activity(welcome_text)
    
    # Also send the new chat card
    await send_new_chat_card(turn_context)

# ----- Common API Functions -----

async def process_conversation_internal(
    client: AzureOpenAI,
    session: Optional[str] = None,
    prompt: Optional[str] = None,
    assistant: Optional[str] = None,
    stream_output: bool = True
):
    """
    Core function to process conversation with the assistant.
    This function handles both streaming and non-streaming modes with robust Stream object handling.
    """
    try:
        # Create defaults if not provided
        if not assistant:
            logging.warning(f"No assistant ID provided, creating a default one.")
            try:
                assistant_obj = client.beta.assistants.create(
                    name="default_conversation_assistant",
                    model="gpt-4o-mini",
                    instructions="You are a helpful conversation assistant.",
                )
                assistant = assistant_obj.id
            except Exception as e:
                logging.error(f"Failed to create default assistant: {e}")
                raise HTTPException(status_code=500, detail="Failed to create default assistant")

        if not session:
            logging.warning(f"No session (thread) ID provided, creating a new one.")
            try:
                thread = client.beta.threads.create()
                session = thread.id
            except Exception as e:
                logging.error(f"Failed to create default thread: {e}")
                raise HTTPException(status_code=500, detail=f"Failed to create default thread: {e}")
        
        # Validate resources if provided 
        validation = await validate_resources(client, session, assistant)
        
        # Create new thread if invalid
        if not validation["thread_valid"]:
            logging.warning(f"Invalid thread ID: {session}, creating a new one")
            try:
                thread = client.beta.threads.create()
                session = thread.id
                logging.info(f"Created recovery thread: {session}")
            except Exception as e:
                logging.error(f"Failed to create recovery thread: {e}")
                raise HTTPException(status_code=500, detail="Failed to create a valid conversation thread")
        
        # Create new assistant if invalid
        if not validation["assistant_valid"]:
            logging.warning(f"Invalid assistant ID: {assistant}, creating a new one")
            try:
                assistant_obj = client.beta.assistants.create(
                    name=f"recovery_assistant_{int(time.time())}",
                    model="gpt-4o-mini",
                    instructions="You are a helpful assistant recovering from a system error.",
                )
                assistant = assistant_obj.id
                logging.info(f"Created recovery assistant: {assistant}")
            except Exception as e:
                logging.error(f"Failed to create recovery assistant: {e}")
                raise HTTPException(status_code=500, detail="Failed to create a valid assistant")
        
        # Check if there's an active run before adding a message
        active_run = False
        run_id = None
        try:
            # List runs to check for active ones
            runs = client.beta.threads.runs.list(thread_id=session, limit=1)
            if runs.data:
                latest_run = runs.data[0]
                if latest_run.status in ["in_progress", "queued", "requires_action"]:
                    active_run = True
                    run_id = latest_run.id
                    logging.warning(f"Active run {run_id} detected with status {latest_run.status}")
        except Exception as e:
            logging.warning(f"Error checking for active runs: {e}")
            # Continue anyway - we'll handle failure when adding messages

        # Add user message to the thread if prompt is given
        if prompt:
            max_retries = 5
            base_retry_delay = 3
            success = False
            
            # Handle active run if found
            if active_run and run_id:
                try:
                    # Cancel the run
                    client.beta.threads.runs.cancel(thread_id=session, run_id=run_id)
                    logging.info(f"Requested cancellation of active run {run_id}")
                    
                    # Wait for run to be fully canceled
                    cancel_wait_time = 5
                    max_cancel_wait = 30
                    wait_time = 0
                    
                    while wait_time < max_cancel_wait:
                        await asyncio.sleep(cancel_wait_time)
                        wait_time += cancel_wait_time
                        
                        # Check if run is actually canceled or completed
                        try:
                            run_status = client.beta.threads.runs.retrieve(thread_id=session, run_id=run_id)
                            if run_status.status in ["cancelled", "completed", "failed", "expired"]:
                                logging.info(f"Run {run_id} is now in state {run_status.status} after waiting {wait_time}s")
                                break
                            else:
                                logging.warning(f"Run {run_id} still in state {run_status.status} after waiting {wait_time}s")
                                # Gradually increase wait time
                                cancel_wait_time = min(cancel_wait_time * 1.5, 10)
                        except Exception as status_e:
                            logging.warning(f"Error checking run status after cancellation: {status_e}")
                            break
                    
                    # If we've waited the maximum time and run is still active, create a new thread
                    if wait_time >= max_cancel_wait:
                        logging.warning(f"Unable to cancel run {run_id} after waiting {wait_time}s, creating new thread")
                        thread = client.beta.threads.create()
                        session = thread.id
                        logging.info(f"Created new thread {session} due to stuck run")
                        active_run = False
                except Exception as cancel_e:
                    logging.error(f"Error canceling run {run_id}: {cancel_e}")
                    # Create a new thread as fallback
                    try:
                        thread = client.beta.threads.create()
                        session = thread.id
                        logging.info(f"Created new thread {session} after failed run cancellation")
                        active_run = False
                    except Exception as thread_e:
                        logging.error(f"Failed to create new thread after cancellation error: {thread_e}")
                        raise HTTPException(status_code=500, detail="Failed to handle active run and create new thread")
            
            # Now try to add the message with retries
            retry_delay = base_retry_delay
            for attempt in range(max_retries):
                try:
                    client.beta.threads.messages.create(
                        thread_id=session,
                        role="user",
                        content=prompt
                    )
                    logging.info(f"Added user message to thread {session} (attempt {attempt+1})")
                    success = True
                    break
                except Exception as e:
                    if "while a run" in str(e) and attempt < max_retries - 1:
                        logging.warning(f"Failed to add message (attempt {attempt+1}), run is still active. Retrying in {retry_delay}s: {e}")
                        await asyncio.sleep(retry_delay)
                        retry_delay *= 2  # Exponential backoff
                        
                        # If we're still having issues after multiple attempts, create a new thread
                        if attempt >= 2:  # After 3rd attempt
                            try:
                                logging.warning("Creating new thread due to persistent run issues")
                                thread = client.beta.threads.create()
                                old_session = session
                                session = thread.id
                                logging.info(f"Switched from thread {old_session} to new thread {session}")
                                # Add the message to the new thread
                                client.beta.threads.messages.create(
                                    thread_id=session,
                                    role="user",
                                    content=prompt
                                )
                                success = True
                                break
                            except Exception as new_thread_e:
                                logging.error(f"Error creating new thread during retries: {new_thread_e}")
                    else:
                        logging.error(f"Failed to add message to thread {session}: {e}")
                        if attempt == max_retries - 1:
                            raise HTTPException(status_code=500, detail="Failed to add message to conversation thread")
            
            if not success:
                raise HTTPException(status_code=500, detail="Failed to add message to conversation thread after retries")
        
        # For streaming mode (/conversation endpoint)
        if stream_output:
            # For API endpoints, we'll use a simpler approach than the Teams integration
            async def async_generator():
                try:
                    # Create run with stream=True
                    run = client.beta.threads.runs.create(
                        thread_id=session,
                        assistant_id=assistant,
                        stream=True
                    )
                    
                    # Handle the stream based on available methods
                    if hasattr(run, "iter_chunks"):
                        # Using iter_chunks synchronous iterator
                        logging.info("Using iter_chunks() for API streaming")
                        for chunk in run.iter_chunks():
                            text_piece = ""
                            
                            if hasattr(chunk, "data") and hasattr(chunk.data, "delta"):
                                delta = chunk.data.delta
                                if hasattr(delta, "content") and delta.content:
                                    for content in delta.content:
                                        if content.type == "text" and hasattr(content.text, "value"):
                                            text_piece = content.text.value
                                            
                            if text_piece:
                                yield text_piece
                                # Small delay to make it work with asyncio
                                await asyncio.sleep(0.01)
                                
                    elif hasattr(run, "events"):
                        # Using events iterator
                        logging.info("Using events iterator for API streaming")
                        for event in run.events:
                            if event.event == "thread.message.delta":
                                if hasattr(event.data, "delta") and hasattr(event.data.delta, "content"):
                                    for content in event.data.delta.content:
                                        if content.type == "text" and hasattr(content.text, "value"):
                                            yield content.text.value
                                            await asyncio.sleep(0.01)
                    else:
                        # Fallback to polling
                        logging.info("Using fallback polling for API streaming")
                        yield "Processing your request...\n"
                        
                        run_id = run.id
                        max_wait_time = 90  # seconds
                        wait_interval = 2   # seconds
                        elapsed_time = 0
                        
                        while elapsed_time < max_wait_time:
                            run_status = client.beta.threads.runs.retrieve(
                                thread_id=session, 
                                run_id=run_id
                            )
                            
                            if run_status.status == "completed":
                                yield "\n"  # Clear the progress line
                                
                                # Get the complete message
                                messages = client.beta.threads.messages.list(
                                    thread_id=session,
                                    order="desc",
                                    limit=1
                                )
                                
                                if messages.data:
                                    latest_message = messages.data[0]
                                    for content_part in latest_message.content:
                                        if content_part.type == 'text':
                                            yield content_part.text.value
                                break
                            
                            elif run_status.status in ["failed", "cancelled", "expired"]:
                                yield f"\nError: Run ended with status {run_status.status}. Please try again."
                                break
                            
                            yield "."  # Show progress
                            await asyncio.sleep(wait_interval)
                            elapsed_time += wait_interval
                        
                        if elapsed_time >= max_wait_time:
                            yield "\nResponse timed out. Please try again."
                
                except Exception as e:
                    logging.error(f"Error in streaming generation: {e}")
                    yield f"\n[ERROR] An error occurred while generating the response: {str(e)}. Please try again.\n"
            
            # Return streaming generator
            return async_generator()
        
        # Handle non-streaming mode (/chat endpoint)
        else:
            # For non-streaming mode, we'll use a completely different approach
            full_response = ""
            try:
                # Create a run without streaming
                run = client.beta.threads.runs.create(
                    thread_id=session,
                    assistant_id=assistant
                )
                run_id = run.id
                logging.info(f"Created run {run_id} for thread {session} (non-streaming mode)")
                
                # Poll for run completion
                max_poll_attempts = 60  # 5 minute timeout with 5 second intervals
                poll_interval = 5  # seconds
                
                for attempt in range(max_poll_attempts):
                    try:
                        run_status = client.beta.threads.runs.retrieve(
                            thread_id=session,
                            run_id=run_id
                        )
                        
                        logging.info(f"Run status poll {attempt+1}/{max_poll_attempts}: {run_status.status}")
                        
                        # Handle completed run
                        if run_status.status == "completed":
                            # Get the latest message
                            messages = client.beta.threads.messages.list(
                                thread_id=session,
                                order="desc",
                                limit=1
                            )
                            
                            if messages and messages.data:
                                latest_message = messages.data[0]
                                for content_part in latest_message.content:
                                    if content_part.type == 'text':
                                        full_response += content_part.text.value
                                
                                logging.info(f"Successfully retrieved final response")
                            break  # Exit the polling loop
                        
                        # Handle failed/cancelled/expired run
                        elif run_status.status in ["failed", "cancelled", "expired"]:
                            logging.error(f"Run ended with status: {run_status.status}")
                            return {"response": f"Sorry, I encountered an error and couldn't complete your request. Run status: {run_status.status}. Please try again."}
                        
                        # Continue polling if still in progress
                        if attempt < max_poll_attempts - 1:
                            await asyncio.sleep(poll_interval)
                            
                    except Exception as poll_e:
                        logging.error(f"Error polling run status (attempt {attempt+1}): {poll_e}")
                        await asyncio.sleep(poll_interval)
                        
                # If we still don't have a response, try one more time to get the latest message
                if not full_response:
                    try:
                        messages = client.beta.threads.messages.list(
                            thread_id=session,
                            order="desc",
                            limit=1
                        )
                        
                        if messages and messages.data:
                            latest_message = messages.data[0]
                            for content_part in latest_message.content:
                                if content_part.type == 'text':
                                    full_response += content_part.text.value
                    except Exception as final_e:
                        logging.error(f"Error retrieving final message: {final_e}")
                
                # Final fallback if we still don't have a response
                if not full_response:
                    full_response = "I processed your request, but couldn't generate a proper response. Please try again or rephrase your question."

                return {"response": full_response}
                
            except Exception as e:
                logging.error(f"Error in non-streaming response generation: {e}")
                return {
                    "response": "An error occurred while processing your request. Please try again."
                }
        
    except Exception as e:
        endpoint_type = "conversation" if stream_output else "chat"
        logging.error(f"Error in /{endpoint_type} endpoint setup: {e}")
        if stream_output:
            # For streaming, we need to return a generator that yields the error
            async def error_generator():
                yield f"Error: {str(e)}"
            return error_generator()
        else:
            # For non-streaming, return a JSON response with the error
            return {"response": f"Error: {str(e)}"}

# ----- FastAPI Endpoints -----

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
    return {"status": "ok", "service": "Product Management and Teams Bot"}

# Root path redirect to health
@app.get("/")
async def root():
    return {"status": "ok", "message": "Product Management and Teams Bot is running."}

# Run the app with uvicorn if executed directly
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    print(f"Starting FastAPI server on http://0.0.0.0:{port}")
    uvicorn.run(app, host="0.0.0.0", port=port)
