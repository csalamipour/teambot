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
from typing import Optional, List, Dict, Any, Tuple, Union
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
)

from botbuilder.schema.teams import (
    FileDownloadInfo,
    FileConsentCard,
    FileConsentCardResponse,
    FileInfoCard,
)
from botbuilder.schema.teams.additional_properties import ContentType
import uuid
from typing import Dict, List, Deque
from collections import deque
import threading
from typing import Dict, List, Optional, Any
from botbuilder.core import CardFactory, Storage, TurnContext
from botbuilder.schema import Activity, ActivityTypes, ChannelAccount, ConversationAccount, Attachment
from teams.state import MemoryBase  # Import MemoryBase from Teams library
import logging
import time
import threading
import asyncio
# Dictionary to store pending messages for each conversation
pending_messages = {}
# Lock for thread-safe operations on the pending_messages dict
pending_messages_lock = threading.Lock()
# Dictionary for tracking active runs
active_runs = {}
# Active streamers by conversation ID
active_streamers = {}
# Lock for thread-safe operations on active_streamers
active_streamers_lock = threading.Lock()

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
def create_message_card(message_text):
    """Creates an adaptive card for displaying message text"""
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "version": "1.6",
        "type": "AdaptiveCard",
        "body": [{"type": "TextBlock", "wrap": True, "text": message_text}]
    }
    return CardFactory.adaptive_card(card)

class TeamsStreamingResponse:
    """Handles streaming responses to Teams using proper Teams streaming protocols with Teams AI compatibility"""
    
    def __init__(self, turn_context):
        self.turn_context = turn_context
        self.conversation_id = TurnContext.get_conversation_reference(turn_context.activity).conversation.id
        self.message_parts = []
        self.stream_id = None
        self.sequence_number = 1
        self.last_update_time = 0
        self.min_update_interval = 0.7  # Minimum time between updates in seconds
        self.active = True
        self.complete = False
        self.attachments = None  # Store attachments for final message
        self.message = ""  # Property required by end_stream_handler
        
    async def initialize(self):
        """Initialize the stream with first informative message"""
        # Register this streamer in active streamers dict
        with active_streamers_lock:
            active_streamers[self.conversation_id] = self
        
        # Send initial informative message
        message = "Generating response..."
        await self._send_streaming_update(message, "start")
        
    async def _send_streaming_update(self, text, stream_type="continue"):
        """Send a streaming update to Teams with proper sequencing"""
        try:
            # FIXED: Verify correct stream types based on activity type
            if self.complete:
                # Final messages must be message type with end stream type
                activity_type = ActivityTypes.message
                stream_type = "end"
            else:
                # In-progress messages must be typing with start/continue
                activity_type = ActivityTypes.typing
                if stream_type not in ["start", "continue"]:
                    stream_type = "start" if self.stream_id is None else "continue"
            
            # Create the activity for a streaming update
            activity = Activity(
                type=activity_type,
                text=text,
                channel_id="msteams",
                entities=[{
                    "type": "streaminfo",
                    "streamType": stream_type
                }]
            )
            
            # Add sequence number for non-final messages
            if not self.complete:
                activity.entities[0]["streamSequence"] = self.sequence_number
                self.sequence_number += 1
            
            # Add stream ID for all but the first message
            if self.stream_id is not None:
                activity.entities[0]["streamId"] = self.stream_id
            
            # Add attachments to final message if available
            if self.complete and self.attachments:
                activity.attachments = self.attachments
            
            # Send the activity
            response = await self.turn_context.send_activity(activity)
            
            # Store the stream ID from the first response
            if self.stream_id is None and response is not None:
                self.stream_id = response.id
                logging.info(f"Initialized stream with ID: {self.stream_id}")
                
            # Update last update time
            self.last_update_time = time.time()
            
            return True
        except Exception as e:
            logging.error(f"Error sending streaming update: {e}")
            self.active = False
            return False
            
    async def send_typing_indicator(self):
        """Sends a dedicated typing indicator to Teams (non-streaming)"""
        if not self.active:
            return
            
        try:
            activity = Activity(
                type=ActivityTypes.typing,
                channel_id="msteams"
            )
            await self.turn_context.send_activity(activity)
        except Exception as e:
            logging.error(f"Error sending typing indicator: {e}")
    
    async def queue_update(self, text_chunk):
        """Queues and potentially sends a text update"""
        if not self.active or self.complete:
            return
            
        # Add to the accumulated text
        self.message_parts.append(text_chunk)
        self.message = "".join(self.message_parts)  # Update self.message property for Teams AI compatibility
        
        # Check if we should send an update
        current_time = time.time()
        
        # Send update if sufficient time has passed
        if (current_time - self.last_update_time) >= self.min_update_interval:
            await self._send_streaming_update(self.message)
            
            # Periodically add a typing indicator to keep the user informed
            if self.sequence_number % 5 == 0:
                await self.send_typing_indicator()
    
    def get_full_message(self):
        """Gets the complete message from all chunks"""
        return self.message
    
    def set_attachments(self, attachments):
        """Sets attachments for final message (required by end_stream_handler)"""
        self.attachments = attachments
        logging.info(f"Set {len(attachments)} attachments for streaming response")
    
    async def send_informative_update(self, message):
        """Sends an informative update to the user"""
        if not self.active or self.complete:
            return False
            
        # For first message, use "start", otherwise use "continue"
        stream_type = "start" if self.stream_id is None else "continue"
        return await self._send_streaming_update(message, stream_type)
        
    async def send_final_message(self, attachments=None):
        """Sends the final complete message with proper stream ending"""
        if not self.active:
            return False
            
        if self.complete:
            logging.warning("Attempted to send final message for already completed stream")
            return False
            
        try:
            # Mark stream as complete before sending (avoid race conditions)
            self.complete = True
            
            # Set attachments if provided
            if attachments:
                self.attachments = attachments
            
            # If no attachments are set and end_stream_handler hasn't been called yet,
            # create a default card
            if not self.attachments:
                # Create an adaptive card with the message
                message_card = create_message_card(self.message)
                self.attachments = [message_card]
            
            # Send the final message with proper streaming format and attachments
            await self._send_streaming_update(self.message, "end")
            
            # Clean up resources
            with active_streamers_lock:
                if self.conversation_id in active_streamers:
                    del active_streamers[self.conversation_id]
            
            return True
            
        except Exception as e:
            logging.error(f"Error sending final streaming message: {e}")
            
            # Try to send as a regular message if streaming fails
            try:
                # Use the attachments if they exist, otherwise create a default card
                if not self.attachments:
                    message_card = create_message_card(self.message)
                    all_attachments = [message_card]
                else:
                    all_attachments = self.attachments
                
                # Create a regular message activity with the card
                activity = Activity(
                    type=ActivityTypes.message,
                    channel_id="msteams",
                    attachments=all_attachments
                )
                
                await self.turn_context.send_activity(activity)
                return True
            except Exception as fallback_e:
                logging.error(f"Failed to send final message even as regular message: {fallback_e}")
                return False
            finally:
                # Clean up resources
                with active_streamers_lock:
                    if self.conversation_id in active_streamers:
                        del active_streamers[self.conversation_id]
                        
    async def abort_streaming(self):
        """Aborts the streaming session and cleans up resources"""
        self.active = False
        self.complete = True
        
        # Clean up resources
        with active_streamers_lock:
            if self.conversation_id in active_streamers:
                del active_streamers[self.conversation_id]
        
        logging.info(f"Aborted streaming session for conversation {self.conversation_id}")
        return True
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
async def end_stream_handler(
    context: TurnContext,
    state: MemoryBase,
    response: Any,  # Using Any instead of PromptResponse[str] for flexibility
    streamer: TeamsStreamingResponse,
):
    """
    Handles the end of streaming by creating an Adaptive Card with the response.
    Called by the Teams AI framework when streaming is complete.
    
    Args:
        context: The turn context
        state: The conversation state
        response: The response from the model
        streamer: The streaming response object
    """
    if not streamer:
        return
    
    try:
        # Create an adaptive card with the full message
        card = CardFactory.adaptive_card(
            {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "version": "1.6",
                "type": "AdaptiveCard",
                "body": [
                    {
                        "type": "TextBlock", 
                        "wrap": True, 
                        "text": streamer.message
                    }
                ]
            }
        )
        
        # Set the attachment on the streamer
        streamer.set_attachments([card])
        
        logging.info("End stream handler completed successfully with Adaptive Card")
    except Exception as e:
        logging.error(f"Error in end_stream_handler: {e}")
def update_operation_status(operation_id: str, status: str, progress: float, message: str):
    """Update the status of a long-running operation."""
    operation_statuses[operation_id] = {
        "status": status,
        "progress": progress,
        "message": message,
        "updated_at": time.time()
    }
    logging.info(f"Operation {operation_id}: {status} - {progress:.0f}% - {message}")

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
                # Process any queued messages
                with pending_messages_lock:
                    if conversation_id in pending_messages and pending_messages[conversation_id]:
                        # Get the next message
                        next_message = pending_messages[conversation_id].popleft()
                        await turn_context.send_activity("Now addressing your follow-up message...")
                        
                        # Save original text
                        original_text = turn_context.activity.text
                        
                        try:
                            # Set the new message text
                            turn_context.activity.text = next_message
                            
                            # Process the message using the same turn context
                            await handle_text_message(turn_context, state)
                        finally:
                            # Restore original text
                            turn_context.activity.text = original_text
                        
                        # Send typing indicator
                        await turn_context.send_activity(create_typing_activity())
                        
                        # Initialize new chat
                        await initialize_chat(turn_context, None)  # Pass None to force new state creation
                    else:
                        await initialize_chat(turn_context, None)
    except Exception as e:
        logging.error(f"Error handling card action: {e}")
        await turn_context.send_activity(f"I couldn't start a new chat. Please try again later.")

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

# Parallel message processing control
# This is a semaphore that limits parallel message processing to avoid overwhelming the API
message_processing_semaphore = asyncio.Semaphore(3)  # Allow up to 3 messages to be processed in parallel

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
            # Always send a typing indicator first
            await turn_context.send_activity(create_typing_activity())
            
            with pending_messages_lock:
                pending_messages[conversation_id].append(turn_context.activity.text.strip())
                queue_length = len(pending_messages[conversation_id])
            
            # Always let the user know we've queued their message for processing
            await turn_context.send_activity(f"I'm still working on your previous request. This message has been queued ({queue_length} in queue).")
            
            # Check if there's an active streamer for this conversation
            with active_streamers_lock:
                active_streamer = active_streamers.get(conversation_id)
                
            # If there's an active streamer, update the informative message to show progress
            if active_streamer and active_streamer.active and not active_streamer.complete:
                await active_streamer.send_informative_update(f"Still working... ({queue_length} new message(s) queued)")
            
            return
        
        # Acquire the semaphore for parallel processing
        async with message_processing_semaphore:
            # Prioritize text processing if we have text content (even if there are non-file attachments)
            if has_text and not has_file_attachments:
                try:
                    # Send typing indicator immediately
                    await turn_context.send_activity(create_typing_activity())
                    await handle_text_message(turn_context, state)
                except Exception as e:
                    logging.error(f"Error in handle_text_message for user {user_id}: {e}")
                    traceback.print_exc()
                    # Attempt recovery
                    await handle_thread_recovery(turn_context, state, str(e))
            
            # Process file attachments with or without caption
            elif has_file_attachments:
                try:
                    # Send typing indicator immediately
                    await turn_context.send_activity(create_typing_activity())
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

async def process_uploaded_file(turn_context: TurnContext, state, file_path: str, filename: str, message_text: str = None):
    """Process an uploaded file after it's been downloaded, with optional message text"""
    # Initialize streaming response
    streamer = TeamsStreamingResponse(turn_context)
    await streamer.initialize()
    
    # If no assistant yet, initialize chat first
    if not state["assistant_id"]:
        await initialize_chat(turn_context, state)
    
    try:
        # Read the file content
        with open(file_path, 'rb') as file:
            file_content = file.read()
            
            # Update streaming with progress
            await streamer.send_informative_update(f"Processing file: '{filename}'...")
            
            # Check file type
            file_ext = os.path.splitext(filename)[1].lower()
            is_image = file_ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
            is_document = file_ext in ['.pdf', '.doc', '.docx', '.txt', '.md', '.html', '.json']
            is_csv_excel = file_ext in ['.csv', '.xlsx', '.xls', '.xlsm']
            
            if is_csv_excel:
                await streamer.send_final_message("Sorry, CSV and Excel files are not supported. Please upload PDF, DOC, DOCX, or TXT files only.")
                return
            
            # Process based on file type
            client = create_client()
            
            if is_image:
                # Update streaming with progress
                await streamer.send_informative_update(f"Analyzing image '{filename}'...")
                
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
                    
                    # Update streaming with final response
                    final_message = f"Image '{filename}' processed successfully!\n\nHere's my analysis of the image:\n\n{analysis_text}"
                    await streamer.queue_update(final_message)
                    await streamer.send_final_message()
                else:
                    await streamer.send_final_message("Cannot process image: No active conversation session.")
                    
            elif is_document:
                # Use the new direct file upload approach for documents
                if state["assistant_id"] and state["session_id"]:
                    # Update streaming with progress
                    await streamer.send_informative_update(f"Uploading document '{filename}'...")
                    
                    # Upload the file directly to the thread
                    if message_text:
                        message_content = f"{message_text}\n\nI've also uploaded a document named '{filename}'. Please use this document to answer my questions."
                    else:
                        message_content = f"I've uploaded a document named '{filename}'. Please use this document to answer my questions."
                    
                    # Use the new OpenAI direct file upload approach
                    try:
                        await streamer.send_informative_update(f"Attaching document to conversation...")
                        
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
                        
                        # Send final streaming message
                        await streamer.queue_update(f"Document '{filename}' uploaded successfully! You can now ask questions about it.")
                        await streamer.send_final_message()
                        
                    except Exception as upload_error:
                        logger.error(f"Error uploading file to OpenAI: {str(upload_error)}")
                        await streamer.send_informative_update(f"Error uploading document. Trying alternate method...")
                        
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
                                await streamer.send_informative_update(f"Creating storage for the document...")
                                vector_store = client.beta.vector_stores.create(name=f"Assistant_{state['assistant_id']}_Store")
                                vector_store_id = vector_store.id
                                state["vector_store_id"] = vector_store_id
                            
                            # Upload to vector store
                            await streamer.send_informative_update(f"Uploading document to storage...")
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
                                await streamer.send_informative_update(f"Configuring assistant to access your document...")
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
                            
                            # Send final success message
                            await streamer.queue_update(f"Document '{filename}' uploaded to vector store successfully as a fallback method. You can now ask questions about it.")
                            await streamer.send_final_message()
                            
                        except Exception as fallback_error:
                            await streamer.queue_update(f"Failed to upload document via fallback method: {str(fallback_error)}")
                            await streamer.send_final_message()
                        finally:
                            # Clean up temp file
                            if os.path.exists(temp_path):
                                os.remove(temp_path)
                else:
                    await streamer.send_final_message("Cannot process document: No active assistant or session.")
            else:
                await streamer.send_final_message(f"Unsupported file type: {file_ext}. Please upload PDF, DOC, DOCX, TXT files, or images.")
    except Exception as e:
        logger.error(f"Error processing file '{filename}': {e}")
        traceback.print_exc()
        
        # Try to send final message through streamer first
        try:
            if streamer and streamer.active and not streamer.complete:
                await streamer.queue_update(f"Error processing file: {str(e)}")
                await streamer.send_final_message()
            else:
                await turn_context.send_activity(f"Error processing file: {str(e)}")
        except:
            # Fallback to regular message
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

# Improved handle_text_message with proper streaming support
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
    
    # Initialize streaming response handler
    streamer = TeamsStreamingResponse(turn_context)
    await streamer.initialize()
    
    # If no assistant yet, initialize chat with the message as context
    if not stored_assistant_id:
        await streamer.send_informative_update("Setting up a new conversation...")
        await initialize_chat(turn_context, state, context=user_message)
        return
    
    # Send typing indicator
    await turn_context.send_activity(create_typing_activity())
    
    # Check if thread needs summarization (with thread safety)
    summarized = False
    if stored_session_id:
        client = create_client()
        
        # Send informative update about checking context
        await streamer.send_informative_update("Checking conversation context...")
        
        summarized = await summarize_thread_if_needed(client, stored_session_id, state, threshold=30)
        
        if summarized:
            # Update stored_session_id after summarization (thread may have changed)
            with conversation_states_lock:
                stored_session_id = state.get("session_id")
                
            await streamer.send_informative_update("Summarized our previous conversation to maintain context while keeping the conversation focused.")
    
    # Mark thread as busy (thread-safe)
    with conversation_states_lock:
        state["active_run"] = True
        current_session_id = state.get("session_id")
    
    if current_session_id:
        active_runs[current_session_id] = True
    
    try:
        # Double-verify resources before proceeding
        client = create_client()
        
        # Send progress update
        await streamer.send_informative_update("Validating conversation resources...")
        
        validation = await validate_resources(client, current_session_id, stored_assistant_id)
        
        # If any resource is invalid, force recovery
        if not validation["thread_valid"] or not validation["assistant_valid"]:
            logging.warning(f"Resource validation failed for user {user_id}: thread_valid={validation['thread_valid']}, assistant_valid={validation['assistant_valid']}")
            
            # Send error message through streamer and end streaming
            await streamer.queue_update("I encountered an issue with our conversation resources. Creating a fresh session...")
            await streamer.send_final_message()
            
            # Force recovery
            raise Exception("Invalid conversation resources detected - forcing recovery")
        
        # Send progress update
        await streamer.send_informative_update("Processing your message...")
        
        # Add message to thread
        client.beta.threads.messages.create(
            thread_id=current_session_id,
            role="user",
            content=user_message
        )
        
        # Create run with proper handling
        run = client.beta.threads.runs.create(
            thread_id=current_session_id,
            assistant_id=stored_assistant_id
        )
        
        run_id = run.id
        logging.info(f"Created run {run_id} for thread {current_session_id}")
        
        # Poll for run completion with streaming updates
        max_wait_time = 120  # seconds
        wait_interval = 1   # seconds
        elapsed_time = 0
        last_progress_update = 0
        progress_update_interval = 5  # seconds
        
        while elapsed_time < max_wait_time:
            # Check for cancellation (user sent a new message)
            with pending_messages_lock:
                has_pending_messages = conversation_id in pending_messages and len(pending_messages[conversation_id]) > 0
            
            # If user has sent a new message, consider cancelling this run
            if has_pending_messages and elapsed_time > 10:  # Only cancel if we've been running for at least 10 seconds
                logging.info(f"User sent new messages while processing run {run_id}, considering early completion")
                
                # Try to get what we have so far
                try:
                    # Get the latest message that might be available
                    messages = client.beta.threads.messages.list(
                        thread_id=current_session_id,
                        order="desc",
                        limit=1
                    )
                    
                    if messages.data and messages.data[0].role == "assistant":
                        # We have a partial response from the assistant, send it
                        latest_message = messages.data[0]
                        partial_response = ""
                        for content_part in latest_message.content:
                            if content_part.type == 'text':
                                partial_response += content_part.text.value
                        
                        if partial_response.strip():
                            # We have a partial response to send
                            await streamer.queue_update(partial_response)
                            await streamer.queue_update("\n\n[Note: This is a partial response as you've sent a new message.]")
                            await streamer.send_final_message()
                            
                            # Update state
                            with conversation_states_lock:
                                state["active_run"] = False
                            if current_session_id in active_runs:
                                del active_runs[current_session_id]
                            
                            # Try to cancel the run
                            try:
                                client.beta.threads.runs.cancel(
                                    thread_id=current_session_id,
                                    run_id=run_id
                                )
                                logging.info(f"Cancelled run {run_id} due to new messages")
                            except:
                                pass
                            
                            return
                except Exception as e:
                    logging.error(f"Error checking for partial response: {e}")
            
            # Send periodic progress updates to keep the user informed
            current_time = time.time()
            if current_time - last_progress_update >= progress_update_interval:
                # Update progress message based on elapsed time
                if elapsed_time < 5:
                    await streamer.send_informative_update("Processing your message...")
                elif elapsed_time < 15:
                    await streamer.send_informative_update("Thinking about your question...")
                elif elapsed_time < 30:
                    await streamer.send_informative_update("Still working on a comprehensive response...")
                else:
                    await streamer.send_informative_update(f"This is taking longer than expected, but I'm still working on it... ({elapsed_time}s)")
                
                last_progress_update = current_time
            
            # Check run status
            run_status = client.beta.threads.runs.retrieve(
                thread_id=current_session_id, 
                run_id=run_id
            )
            
            if run_status.status == "completed":
                # Get the complete message
                messages = client.beta.threads.messages.list(
                    thread_id=current_session_id,
                    order="desc",
                    limit=1
                )
                
                if messages.data:
                    latest_message = messages.data[0]
                    response_text = ""
                    for content_part in latest_message.content:
                        if content_part.type == 'text':
                            response_text += content_part.text.value
                    
                    # Stream full response through streamer
                    await streamer.queue_update(response_text)
                    await streamer.send_final_message()
                else:
                    await streamer.queue_update("I processed your request, but couldn't generate a proper response.")
                    await streamer.send_final_message()
                
                break
            
            elif run_status.status in ["failed", "cancelled", "expired"]:
                logging.error(f"Run ended with status: {run_status.status}")
                await streamer.queue_update(f"Sorry, I encountered an error and couldn't complete your request. Run status: {run_status.status}. Please try again.")
                await streamer.send_final_message()
                break
            
            await asyncio.sleep(wait_interval)
            elapsed_time += wait_interval
        
        # Handle timeout
        if elapsed_time >= max_wait_time:
            logging.warning(f"Run {run_id} timed out after {max_wait_time} seconds")
            await streamer.queue_update("I'm sorry, it's taking longer than expected to generate a response. Please try again or rephrase your question.")
            await streamer.send_final_message()
            
            # Try to cancel the run
            try:
                client.beta.threads.runs.cancel(
                    thread_id=current_session_id,
                    run_id=run_id
                )
                logging.info(f"Cancelled timed out run {run_id}")
            except:
                pass
        
        # Mark thread as no longer busy (thread-safe)
        with conversation_states_lock:
            state["active_run"] = False
            current_session_id = state.get("session_id")
        
        if current_session_id in active_runs:
            del active_runs[current_session_id]
        
        # Process any queued messages
        with pending_messages_lock:
            if conversation_id in pending_messages and pending_messages[conversation_id]:
                # Get the next message
                next_message = pending_messages[conversation_id].popleft()
                await turn_context.send_activity("Now addressing your follow-up message...")
                
                # Create a new turn context for the next message
                new_activity = copy.deepcopy(turn_context.activity)
                new_activity.text = next_message
                new_turn_context = TurnContext(ADAPTER, new_activity)
                
                # Process the next message
                await handle_text_message(new_turn_context, state)
            
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
        
        # Send error through streamer if active
        if streamer and streamer.active and not streamer.complete:
            await streamer.queue_update("I'm sorry, I encountered a problem while processing your message. Please try again.")
            await streamer.send_final_message()
        else:
            await turn_context.send_activity("I'm sorry, I encountered a problem while processing your message. Please try again.")

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
    
    # Create a streamer for a better user experience
    streamer = TeamsStreamingResponse(turn_context)
    await streamer.initialize()
    
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
        await streamer.send_informative_update("Creating a new conversation...")
        
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
            await streamer.send_informative_update("Creating storage resources...")
            vector_store = client.beta.vector_stores.create(
                name=f"user_{user_id}_convo_{conversation_id}_{int(time.time())}"
            )
            logging.info(f"Created vector store: {vector_store.id} for user {user_id}")
        except Exception as e:
            logging.error(f"Failed to create vector store for user {user_id}: {e}")
            await streamer.queue_update("I encountered an error while setting up your conversation storage.")
            await streamer.send_final_message()
            raise HTTPException(status_code=500, detail="Failed to create vector store")

        # Include file_search tool
        assistant_tools = [{"type": "file_search"}]
        assistant_tool_resources = {
            "file_search": {"vector_store_ids": [vector_store.id]}
        }

        # Create the assistant with a unique name including user identifiers
        try:
            await streamer.send_informative_update("Creating your personal assistant...")
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
            await streamer.queue_update("I encountered an error while creating your personal assistant.")
            await streamer.send_final_message()
            raise HTTPException(status_code=500, detail=f"Failed to create assistant: {e}")

        # Create a thread
        try:
            await streamer.send_informative_update("Setting up your conversation...")
            thread = client.beta.threads.create()
            logging.info(f"Created thread {thread.id} for user {user_id}")
        except Exception as e:
            logging.error(f"Failed to create thread for user {user_id}: {e}")
            await streamer.queue_update("I encountered an error while setting up your conversation thread.")
            await streamer.send_final_message()
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
            await streamer.send_informative_update("Adding your context to the conversation...")
            await update_context_internal(client, thread.id, context)
            
        # Compose greeting messages
        welcome_message = "Hi! I'm the Product Management Bot. I'm ready to help you with your product management tasks."
        
        if context:
            welcome_message += f"\n\nI've initialized with your context: '{context}'"
        
        # Send final welcome message through streamer
        await streamer.queue_update(welcome_message)
        await streamer.send_final_message()
            
        # Also process the first message if context was provided
        if context:
            # Process the context as a message
            await send_message(turn_context, state)
            
    except Exception as e:
        # Try to send error through streamer if active
        if streamer and streamer.active and not streamer.complete:
            await streamer.queue_update(f"Error initializing chat: {str(e)}")
            await streamer.send_final_message()
        else:
            await turn_context.send_activity(f"Error initializing chat: {str(e)}")
            
        logger.error(f"Error in initialize_chat for user {user_id}: {str(e)}")
        traceback.print_exc()

# Send a message without user input (used after file upload or initialization)
async def send_message(turn_context: TurnContext, state):
    try:
        # Create streaming response handler
        streamer = TeamsStreamingResponse(turn_context)
        await streamer.initialize()
        
        # Call internal function directly to get latest message
        client = create_client()
        
        # Send informative update
        await streamer.send_informative_update("Processing your request...")
        
        # Create a run to process the current thread state
        run = client.beta.threads.runs.create(
            thread_id=state["session_id"],
            assistant_id=state["assistant_id"]
        )
        
        # Poll for completion
        max_wait_time = 90  # seconds
        wait_interval = 1   # seconds
        elapsed_time = 0
        last_progress_update = 0
        progress_update_interval = 5  # seconds
        
        while elapsed_time < max_wait_time:
            # Send periodic progress updates
            current_time = time.time()
            if current_time - last_progress_update >= progress_update_interval:
                if elapsed_time < 10:
                    await streamer.send_informative_update("Analyzing your request...")
                elif elapsed_time < 30:
                    await streamer.send_informative_update("Working on a comprehensive response...")
                else:
                    await streamer.send_informative_update(f"Still processing... ({elapsed_time}s)")
                
                last_progress_update = current_time
            
            # Check run status
            run_status = client.beta.threads.runs.retrieve(
                thread_id=state["session_id"],
                run_id=run.id
            )
            
            if run_status.status == "completed":
                # Get the complete message
                messages = client.beta.threads.messages.list(
                    thread_id=state["session_id"],
                    order="desc",
                    limit=1
                )
                
                if messages.data:
                    latest_message = messages.data[0]
                    
                    # Extract the message text
                    response_text = ""
                    for content_part in latest_message.content:
                        if content_part.type == 'text':
                            response_text += content_part.text.value
                    
                    # Send through streamer
                    await streamer.queue_update(response_text)
                    await streamer.send_final_message()
                else:
                    await streamer.queue_update("I processed your request, but couldn't generate a response.")
                    await streamer.send_final_message()
                
                break
            
            elif run_status.status in ["failed", "cancelled", "expired"]:
                logging.error(f"Run ended with status: {run_status.status}")
                await streamer.queue_update(f"Sorry, I encountered an error processing your request. Run status: {run_status.status}. Please try again.")
                await streamer.send_final_message()
                break
            
            # Wait before next check
            await asyncio.sleep(wait_interval)
            elapsed_time += wait_interval
        
        # Handle timeout
        if elapsed_time >= max_wait_time:
            logging.warning(f"Run timed out after {max_wait_time} seconds")
            await streamer.queue_update("I'm sorry, it's taking longer than expected to generate a response. Please try again or rephrase your question.")
            await streamer.send_final_message()
            
    except Exception as e:
        logging.error(f"Error in send_message: {str(e)}")
        traceback.print_exc()
        
        # Try to send through streamer if active
        if streamer and streamer.active and not streamer.complete:
            await streamer.queue_update(f"Error getting response: {str(e)}")
            await streamer.send_final_message()
        else:
            await turn_context.send_activity(f"Error getting response: {str(e)}")

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

# Internal implementation of process_conversation
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
                raise HTTPException(status_code=500, detail="Failed to create default thread")
        
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
            max_retries = 5  # Increased from 3 to 5
            base_retry_delay = 3  # Increased from 2 to 3 seconds
            success = False
            
            # Handle active run if found
            if active_run and run_id:
                try:
                    # Cancel the run
                    client.beta.threads.runs.cancel(thread_id=session, run_id=run_id)
                    logging.info(f"Requested cancellation of active run {run_id}")
                    
                    # Wait for run to be fully canceled - this is the key improvement
                    cancel_wait_time = 5  # Wait 5 seconds initially after cancellation request
                    max_cancel_wait = 30  # Maximum time to wait for cancellation
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

# ----- API Endpoints -----

# Status endpoint
@app.get("/operation-status/{operation_id}")
async def check_operation_status(operation_id: str):
    """Check the status of a long-running operation."""
    if operation_id not in operation_statuses:
        return JSONResponse(
            status_code=404,
            content={"error": f"No operation found with ID {operation_id}"}
        )
    
    return JSONResponse(content=operation_statuses[operation_id])

# Internal implementation of initiate_chat that can be called directly
async def initiate_chat_internal(client: AzureOpenAI, context: Optional[str] = None, file: Optional[UploadFile] = None):
    """Internal implementation of initiate_chat that can be called directly by the Teams bot."""
    logging.info("Initiating new chat session...")

    # Create a vector store up front
    try:
        vector_store = client.beta.vector_stores.create(name=f"chat_init_store_{int(time.time())}")
        logging.info(f"Vector store created: {vector_store.id}")
    except Exception as e:
        logging.error(f"Failed to create vector store: {e}")
        raise HTTPException(status_code=500, detail="Failed to create vector store")

    # Include file_search tool
    assistant_tools = [
        {"type": "file_search"}
    ]
    
    assistant_tool_resources = {
        "file_search": {"vector_store_ids": [vector_store.id]}
    }

    # Use the improved system prompt
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
    
    # Create the assistant
    try:
        assistant = client.beta.assistants.create(
            name=f"pm_copilot_{int(time.time())}",
            model="gpt-4o-mini",  # Ensure this model is deployed
            instructions=system_prompt,
            tools=assistant_tools,
            tool_resources=assistant_tool_resources,
        )
        logging.info(f'Assistant created: {assistant.id}')
    except Exception as e:
        logging.error(f"An error occurred while creating the assistant: {e}")
        # Attempt to clean up vector store if assistant creation fails
        try:
            client.beta.vector_stores.delete(vector_store_id=vector_store.id)
            logging.info(f"Cleaned up vector store {vector_store.id} after assistant creation failure.")
        except Exception as cleanup_e:
            logging.error(f"Failed to cleanup vector store {vector_store.id} after error: {cleanup_e}")
        raise HTTPException(status_code=500, detail=f"An error occurred while creating assistant: {e}")

    # Create a thread
    try:
        thread = client.beta.threads.create()
        logging.info(f"Thread created: {thread.id}")
    except Exception as e:
        logging.error(f"An error occurred while creating the thread: {e}")
        # Attempt cleanup
        try:
            client.beta.assistants.delete(assistant_id=assistant.id)
            logging.info(f"Cleaned up assistant {assistant.id} after thread creation failure.")
        except Exception as cleanup_e:
            logging.error(f"Failed to cleanup assistant {assistant.id} after error: {cleanup_e}")
        try:
            client.beta.vector_stores.delete(vector_store_id=vector_store.id)
            logging.info(f"Cleaned up vector store {vector_store.id} after thread creation failure.")
        except Exception as cleanup_e:
            logging.error(f"Failed to cleanup vector store {vector_store.id} after error: {cleanup_e}")
        raise HTTPException(status_code=500, detail=f"An error occurred while creating the thread: {e}")

    # If context is provided, add it as user persona context
    if context:
        await update_context_internal(client, thread.id, context)
    # Errors handled within update_context

    # If a file is provided, upload and process it
    if file:
        filename = file.filename
        file_content = await file.read()
        file_path = os.path.join('/tmp/', filename)  # Use /tmp or a configurable temp dir

        try:
            with open(file_path, 'wb') as f:
                f.write(file_content)

            # Determine file type
            file_ext = os.path.splitext(filename)[1].lower()
            is_csv_excel = file_ext in ['.csv', '.xlsx', '.xls', '.xlsm']
            # Check MIME type as well for broader image support
            mime_type, _ = mimetypes.guess_type(filename)
            is_image = file_ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'] or (mime_type and mime_type.startswith('image/'))
            is_document = file_ext in ['.pdf', '.doc', '.docx', '.txt', '.md', '.html', '.json']  # Common types for vector store

            # Reject CSV/Excel files
            if is_csv_excel:
                file_info = {
                    "name": filename,
                    "type": "unsupported"
                }
                
                # Add unsupported file warning message to thread
                client.beta.threads.messages.create(
                    thread_id=thread.id,
                    role="user",
                    content=f"Warning: The file '{filename}' is a CSV/Excel file which is not supported. Please upload PDF, DOC, DOCX, or TXT files only."
                )
                logging.info(f"Rejected unsupported file type: {filename}")
            elif is_image:
                # Analyze image and add analysis text to the thread
                analysis_text = await image_analysis_internal(file_content, filename, None)
                client.beta.threads.messages.create(
                    thread_id=thread.id,
                    role="user",  # Add analysis as user message for context
                    content=f"Analysis result for uploaded image '{filename}':\n{analysis_text}"
                )
                file_info = {
                    "name": filename,
                    "type": "image",
                    "processing_method": "thread_message"
                }
                await add_file_awareness_internal(thread.id, file_info)
                logging.info(f"Added image analysis for '{filename}' to thread {thread.id}")
            elif is_document:
                # Upload to vector store
                with open(file_path, "rb") as file_stream:
                    file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
                        vector_store_id=vector_store.id,
                        files=[file_stream]
                    )
                file_info = {
                    "name": filename,
                    "type": file_ext[1:] if file_ext else "document",
                    "processing_method": "vector_store"
                }
                await add_file_awareness_internal(thread.id, file_info)
                logging.info(f"File '{filename}' uploaded to vector store {vector_store.id}: status={file_batch.status}, count={file_batch.file_counts.total}")
            else:
                logging.warning(f"File type for '{filename}' not explicitly handled for upload, skipping specific processing.")
                file_info = {
                    "name": filename,
                    "type": "unknown"
                }

        except Exception as e:
            logging.error(f"Error processing uploaded file '{filename}': {e}")
            # Don't raise HTTPException here, allow response with IDs but log error
        finally:
            # Clean up temporary file
            if os.path.exists(file_path):
                try:
                    os.remove(file_path)
                except OSError as e:
                    logging.error(f"Error removing temporary file {file_path}: {e}")

    res = {
        "message": "Chat initiated successfully.",
        "assistant": assistant.id,
        "session": thread.id,  # Use 'session' for thread_id consistency with other endpoints
        "vector_store": vector_store.id
    }

    return res

# FastAPI endpoint for initiate_chat
@app.post("/initiate-chat")
async def initiate_chat(
    context: Optional[str] = Form(None),
    file: Optional[UploadFile] = File(None)
):
    """
    Initiates a new assistant, session (thread), and vector store.
    Optionally uploads a file and sets user context.
    """
    try:
        client = create_client()
        result = await initiate_chat_internal(client, context, file)
        return JSONResponse(result, status_code=200)
    except HTTPException as http_e:
        raise http_e
    except Exception as e:
        logging.error(f"Error in /initiate-chat endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to initiate chat: {str(e)}")

# Internal implementation of co_pilot that can be called directly
async def co_pilot_internal(client: AzureOpenAI, assistant_id: str, vector_store_id: str, context: Optional[str] = None):
    """
    Internal implementation of co_pilot that can be called directly by the Teams bot.
    
    Args:
        client: Azure OpenAI client instance
        assistant_id: Assistant ID
        vector_store_id: Vector store ID
        context: Optional context for the thread
    
    Returns:
        Dictionary with assistant, session, and vector_store IDs
    """
    try:
        # Retrieve the assistant to verify it exists
        try:
            assistant_obj = client.beta.assistants.retrieve(assistant_id=assistant_id)
            logging.info(f"Using existing assistant: {assistant_id}")
        except Exception as e:
            logging.error(f"Error retrieving assistant {assistant_id}: {e}")
            raise HTTPException(status_code=404, detail=f"Assistant not found: {assistant_id}")

        # Verify the vector store exists
        try:
            # Just try to retrieve it to verify it exists
            client.beta.vector_stores.retrieve(vector_store_id=vector_store_id)
            logging.info(f"Using existing vector store: {vector_store_id}")
        except Exception as e:
            logging.error(f"Error retrieving vector store {vector_store_id}: {e}")
            raise HTTPException(status_code=404, detail=f"Vector store not found: {vector_store_id}")

        # Ensure assistant has the right tools and vector store is linked
        current_tools = assistant_obj.tools if assistant_obj.tools else []
        
        # Check for file_search tool, add if missing
        if not any(tool.type == "file_search" for tool in current_tools if hasattr(tool, 'type')):
            current_tools.append({"type": "file_search"})
            logging.info(f"Adding file_search tool to assistant {assistant_id}")

        # Prepare tool resources
        tool_resources = {
            "file_search": {"vector_store_ids": [vector_store_id]},
        }

        # Update the assistant with tools and vector store
        client.beta.assistants.update(
            assistant_id=assistant_id,
            tools=current_tools,
            tool_resources=tool_resources
        )
        logging.info(f"Updated assistant {assistant_id} with tools and vector store {vector_store_id}")

        # Create a new thread
        thread = client.beta.threads.create()
        thread_id = thread.id
        logging.info(f"Created new thread: {thread_id} for assistant {assistant_id}")

        # If context is provided, add it to the thread
        if context:
            await update_context_internal(client, thread_id, context)
            logging.info(f"Added context to thread {thread_id}")

        # Return the same structure as initiate-chat
        return {
            "message": "Chat initiated successfully.",
            "assistant": assistant_id,
            "session": thread_id,
            "vector_store": vector_store_id
        }

    except HTTPException:
        # Re-raise HTTP exceptions to preserve their status codes
        raise
    except Exception as e:
        logging.error(f"Error in co_pilot_internal: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to process co-pilot request: {str(e)}")

# FastAPI endpoint for co-pilot
@app.post("/co-pilot")
async def co_pilot(
    assistant: str = Form(...),
    vector_store: str = Form(...),
    context: Optional[str] = Form(None)
):
    """
    Sets context for a chatbot, creates a new thread using existing assistant and vector store.
    Required parameters: assistant_id, vector_store_id
    Optional parameters: context
    Returns: Same structure as initiate-chat
    """
    try:
        client = create_client()
        result = await co_pilot_internal(client, assistant, vector_store, context)
        return JSONResponse(result, status_code=200)
    except HTTPException as http_e:
        raise http_e
    except Exception as e:
        logging.error(f"Error in /co-pilot endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to process co-pilot request: {str(e)}")

# FastAPI endpoint for upload_file
@app.post("/upload-file")
async def upload_file(
    file: UploadFile = File(...),
    assistant: str = Form(...),
    session: Optional[str] = Form(None),
    context: Optional[str] = Form(None),
    prompt: Optional[str] = Form(None)
):
    """
    Uploads a file and associates it with the given assistant.
    Handles different file types appropriately.
    """
    try:
        client = create_client()
        result = await upload_file_internal(client, file, assistant, session, context, prompt)
        return JSONResponse(result, status_code=200)
    except HTTPException as http_e:
        raise http_e
    except Exception as e:
        logging.error(f"Error in /upload-file endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to upload file: {str(e)}")

# FastAPI endpoint for conversation (streaming)
@app.get("/conversation")
async def conversation(
    session: Optional[str] = None,
    prompt: Optional[str] = None,
    assistant: Optional[str] = None
):
    """
    Handles conversation queries with streaming response.
    """
    try:
        client = create_client()
        response_stream = await process_conversation_internal(client, session, prompt, assistant, stream_output=True)
        return StreamingResponse(response_stream, media_type="text/event-stream")
    except HTTPException as http_e:
        raise http_e
    except Exception as e:
        logging.error(f"Error in /conversation endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to process conversation: {str(e)}")

# FastAPI endpoint for chat (non-streaming)
@app.get("/chat")
async def chat(
    session: Optional[str] = None,
    prompt: Optional[str] = None,
    assistant: Optional[str] = None
):
    """
    Handles conversation queries and returns the full response as JSON.
    Uses the same logic as the streaming endpoint but returns the complete response.
    """
    try:
        client = create_client()
        result = await process_conversation_internal(client, session, prompt, assistant, stream_output=False)
        return JSONResponse(result)
    except HTTPException as http_e:
        raise http_e
    except Exception as e:
        logging.error(f"Error in /chat endpoint: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to process chat: {str(e)}")

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
