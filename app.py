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
    FileDownloadInfo
    FileConsentCard,
    FileConsentCardResponse,
    FileInfoCard,
)
from botbuilder.schema.teams.additional_properties import ContentType

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
    
    # Handle Teams file consent card responses
    elif turn_context.activity.type == ActivityTypes.invoke:
        if turn_context.activity.name == "fileConsent/invoke":
            await handle_file_consent_response(turn_context, turn_context.activity.value)
    
    # Handle conversation update (bot added to conversation)
    elif turn_context.activity.type == ActivityTypes.conversation_update:
        if turn_context.activity.members_added:
            for member in turn_context.activity.members_added:
                if member.id != turn_context.activity.recipient.id:
                    # Bot was added - send welcome message
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
async def handle_file_upload(turn_context: TurnContext, state):
    """Handle file uploads from Teams"""
    
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
                
                # Process the file as normal with direct function calls
                await process_uploaded_file(turn_context, state, file_path, attachment.name)
            else:
                # Not a valid file attachment
                await turn_context.send_activity("Please upload a file using the file upload feature in Teams.")
                
        except Exception as e:
            logger.error(f"Error processing file: {str(e)}")
            traceback.print_exc()
            await turn_context.send_activity(f"Error processing file: {str(e)}")

async def process_uploaded_file(turn_context: TurnContext, state, file_path: str, filename: str):
    """Process an uploaded file after it's been downloaded"""
    # Message user that file is being processed
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
            if is_image:
                # Analyze image
                analysis_text = await image_analysis_internal(file_content, filename)
                
                # Add analysis to the thread
                if state["session_id"]:
                    client = create_client()
                    client.beta.threads.messages.create(
                        thread_id=state["session_id"],
                        role="user",
                        content=f"Analysis result for uploaded image '{filename}':\n{analysis_text}"
                    )
                    
                    # Add image file awareness - direct function call
                    await add_file_awareness_internal(
                        state["session_id"], 
                        {
                            "name": filename,
                            "type": "image",
                            "processing_method": "thread_message"
                        }
                    )
                    
                    await turn_context.send_activity(f"Image '{filename}' processed successfully!")
                    await turn_context.send_activity("Here's my analysis of the image:")
                    await turn_context.send_activity(analysis_text)
                else:
                    await turn_context.send_activity("Cannot process image: No active conversation session.")
                    
            elif is_document:
                # If assistant and session exist, upload to vector store
                if state["assistant_id"] and state["session_id"]:
                    client = create_client()
                    
                    # Get current vector store IDs
                    assistant_obj = client.beta.assistants.retrieve(assistant_id=state["assistant_id"])
                    vector_store_ids = []
                    
                    if hasattr(assistant_obj, 'tool_resources') and assistant_obj.tool_resources:
                        file_search_resources = getattr(assistant_obj.tool_resources, 'file_search', None)
                        if file_search_resources and hasattr(file_search_resources, 'vector_store_ids'):
                            vector_store_ids = list(file_search_resources.vector_store_ids)
                    
                    # Ensure a vector store exists
                    if not vector_store_ids:
                        vector_store = client.beta.vector_stores.create(name=f"Assistant_{state['assistant_id']}_Store")
                        vector_store_ids = [vector_store.id]
                        state["vector_store_id"] = vector_store.id
                    else:
                        state["vector_store_id"] = vector_store_ids[0]
                    
                    # Create a temporary file for uploading to vector store
                    with tempfile.NamedTemporaryFile(delete=False, suffix='_' + filename) as temp:
                        temp.write(file_content)
                        temp_path = temp.name
                    
                    try:
                        # Upload to vector store
                        with open(temp_path, "rb") as file_stream:
                            file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
                                vector_store_id=vector_store_ids[0],
                                files=[file_stream]
                            )
                        
                        # Add file awareness
                        await add_file_awareness_internal(
                            state["session_id"], 
                            {
                                "name": filename,
                                "type": file_ext[1:] if file_ext else "document",
                                "processing_method": "vector_store"
                            }
                        )
                        
                        await turn_context.send_activity(f"File '{filename}' uploaded successfully! You can now ask questions about it.")
                        state["uploaded_files"].append(filename)
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

# Function to handle text messages
async def handle_text_message(turn_context: TurnContext, state):
    user_message = turn_context.activity.text.strip()
    
    # If no assistant yet, initialize chat with the message as context
    if not state["assistant_id"]:
        await initialize_chat(turn_context, state, context=user_message)
        return
    
    # Send typing indicator
    await turn_context.send_activity(create_typing_activity())
    
    try:
        # Use streaming if supported by the channel
        supports_streaming = turn_context.activity.channel_id == "msteams"
        
        if supports_streaming:
            # Use streaming response for Teams
            await stream_response_to_teams(turn_context, state, user_message)
        else:
            # Call the internal function directly without HTTP calls
            client = create_client()
            result = await process_conversation_internal(
                client=client,
                session=state["session_id"],
                prompt=user_message,
                assistant=state["assistant_id"],
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
            
    except Exception as e:
        await turn_context.send_activity(f"Error processing your message: {str(e)}")
        logger.error(f"Error in handle_text_message: {str(e)}")
        traceback.print_exc()

# Initialize chat with the backend
async def initialize_chat(turn_context: TurnContext, state, context=None):
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Call internal function directly
        client = create_client()
        result = await initiate_chat_internal(client, context=context)
        
        if result and isinstance(result, dict):
            state["assistant_id"] = result.get("assistant")
            state["session_id"] = result.get("session")
            state["vector_store_id"] = result.get("vector_store")
            
            # Tell the user chat was initialized
            await turn_context.send_activity("Hi! I'm the Product Management Bot. I'm ready to help you with your product management tasks.")
            
            if context:
                await turn_context.send_activity(f"I've initialized with your context: '{context}'")
                # Also send the first response
                await send_message(turn_context, state)
        else:
            await turn_context.send_activity(f"Failed to initialize chat. Please try again.")
            if isinstance(result, str):
                await turn_context.send_activity(f"Error details: {result}")
    
    except Exception as e:
        await turn_context.send_activity(f"Error initializing chat: {str(e)}")
        logger.error(f"Error in initialize_chat: {str(e)}")
        traceback.print_exc()

# Send a message without user input (used after file upload or initialization)
async def send_message(turn_context: TurnContext, state):
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
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

# Stream response to Teams
async def stream_response_to_teams(turn_context: TurnContext, state, user_message):
    try:
        client = create_client()
        
        # First, send an informative update
        await turn_context.send_activity(create_typing_activity())
        informative_activity = Activity(
            type="message",
            text="Searching through documents...",
            channel_id="msteams",
            entities=[{
                "type": "streaminfo",
                "streamType": "informative",
                "streamSequence": 1
            }]
        )
        await turn_context.send_activity(informative_activity)
        
        # Start streaming process by calling internal function
        stream_generator = await process_conversation_internal(
            client=client,
            session=state["session_id"],
            prompt=user_message,
            assistant=state["assistant_id"],
            stream_output=True
        )
        
        # Process the streaming response
        if hasattr(stream_generator, "__aiter__"):
            stream_id = f"a-{int(time.time())}"
            sequence = 2
            current_text = ""
            
            # Keep track of time to avoid exceeding streaming time limits
            start_time = time.time()
            max_stream_time = 110  # Maximum streaming time in seconds
            
            async for chunk in stream_generator:
                if time.time() - start_time > max_stream_time:
                    logger.warning("Streaming exceeded maximum time limit - sending final message")
                    break
                    
                # Add chunk to current text
                if chunk:
                    current_text += chunk
                    
                    # Send update every 500ms or so (rate limit-aware)
                    streaming_activity = Activity(
                        type="typing",
                        text=current_text,
                        channel_id="msteams",
                        entities=[{
                            "type": "streaminfo",
                            "streamId": stream_id,
                            "streamType": "streaming",
                            "streamSequence": sequence
                        }]
                    )
                    await turn_context.send_activity(streaming_activity)
                    sequence += 1
                    
                    # Slow down to avoid rate limiting
                    await asyncio.sleep(0.75)
            
            # Send final message
            final_activity = Activity(
                type="message",
                text=current_text,
                channel_id="msteams",
                entities=[{
                    "type": "streaminfo",
                    "streamId": stream_id,
                    "streamType": "final"
                }]
            )
            await turn_context.send_activity(final_activity)
        else:
            # Fallback to non-streaming
            if isinstance(stream_generator, dict) and "response" in stream_generator:
                await turn_context.send_activity(stream_generator["response"])
            else:
                await turn_context.send_activity("I've processed your request, but encountered an issue with streaming. Please try again.")
                
    except Exception as e:
        logger.error(f"Error in stream_response_to_teams: {str(e)}")
        traceback.print_exc()
        await turn_context.send_activity(f"Error processing streaming response: {str(e)}")
        
        # Fallback to non-streaming
        try:
            client = create_client()
            result = await process_conversation_internal(
                client=client,
                session=state["session_id"],
                prompt=user_message,
                assistant=state["assistant_id"],
                stream_output=False
            )
            
            if isinstance(result, dict) and "response" in result:
                await turn_context.send_activity(result["response"])
        except:
            await turn_context.send_activity("I'm sorry, I couldn't process your request due to a technical issue.")

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

# ----- Common API Functions -----

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
            max_tokens=1000  # Increased max_tokens for potentially more detailed analysis
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
                await add_file_awareness_internal(client, thread.id, file_info)
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
                await add_file_awareness_internal(client, thread.id, file_info)
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

# Internal implementation of upload_file that can be called directly
async def upload_file_internal(client: AzureOpenAI, file: UploadFile, assistant: str, session: Optional[str] = None, context: Optional[str] = None, prompt: Optional[str] = None):
    """
    Internal implementation of upload_file that can be called directly by the Teams bot.
    
    Args:
        client: Azure OpenAI client instance
        file: The uploaded file
        assistant: Assistant ID
        session: Optional session ID
        context: Optional context
        prompt: Optional prompt for image analysis
        
    Returns:
        Dictionary with upload result information
    """
    filename = file.filename
    file_path = f"/tmp/{filename}"
    uploaded_file_details = {}  # To return info about the uploaded file

    try:
        # Save the uploaded file locally and get the data
        file_content = await file.read()
        with open(file_path, "wb") as temp_file:
            temp_file.write(file_content)

        # Determine file type
        file_ext = os.path.splitext(filename)[1].lower()
        is_csv_excel = file_ext in ['.csv', '.xlsx', '.xls', '.xlsm']
        mime_type, _ = mimetypes.guess_type(filename)
        is_image = file_ext in ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'] or (mime_type and mime_type.startswith('image/'))
        is_document = file_ext in ['.pdf', '.doc', '.docx', '.txt', '.md', '.html', '.json']

        # Retrieve the assistant
        assistant_obj = client.beta.assistants.retrieve(assistant_id=assistant)
        
        # Get current vector store IDs first
        vector_store_ids = []
        if hasattr(assistant_obj, 'tool_resources') and assistant_obj.tool_resources:
            file_search_resources = getattr(assistant_obj.tool_resources, 'file_search', None)
            if file_search_resources and hasattr(file_search_resources, 'vector_store_ids'):
                vector_store_ids = list(file_search_resources.vector_store_ids)
        
        # Handle CSV/Excel files - reject them
        if is_csv_excel:
            uploaded_file_details = {
                "message": "CSV and Excel files are not supported. Please upload PDF, DOC, DOCX, or TXT files.",
                "filename": filename,
                "type": "unsupported",
                "processing_method": "rejected"
            }
            
            # If session provided, add warning message
            if session:
                client.beta.threads.messages.create(
                    thread_id=session,
                    role="user",
                    content=f"Warning: The file '{filename}' is a CSV/Excel file which is not supported. Please upload PDF, DOC, DOCX, or TXT files only."
                )
                
            logging.info(f"Rejected unsupported file type: {filename}")
                
        # Handle document files
        elif is_document:
            # Ensure a vector store is linked or create one
            if not vector_store_ids:
                logging.info(f"No vector store linked to assistant {assistant}. Creating and linking a new one.")
                vector_store = client.beta.vector_stores.create(name=f"Assistant_{assistant}_Store")
                vector_store_ids = [vector_store.id]

            vector_store_id_to_use = vector_store_ids[0]  # Use the first linked store

            # Upload to vector store
            with open(file_path, "rb") as file_stream:
                file_batch = client.beta.vector_stores.file_batches.upload_and_poll(
                    vector_store_id=vector_store_id_to_use,
                    files=[file_stream]
                )
            uploaded_file_details = {
                "message": "File successfully uploaded to vector store.",
                "filename": filename,
                "vector_store_id": vector_store_id_to_use,
                "processing_method": "vector_store",
                "batch_status": file_batch.status
            }
            
            # If session provided, add file awareness message
            if session:
                await add_file_awareness_internal(
                    thread_id=session, 
                    file_info={
                        "name": filename,
                        "type": file_ext[1:] if file_ext else "document",
                        "processing_method": "vector_store"
                    }
                )
                
            logging.info(f"Uploaded '{filename}' to vector store {vector_store_id_to_use} for assistant {assistant}")
            
            # Update assistant with file_search if needed
            try:
                has_file_search = False
                for tool in assistant_obj.tools:
                    if hasattr(tool, 'type') and tool.type == "file_search":
                        has_file_search = True
                        break
                
                if not has_file_search:
                    # Get full list of tools while preserving any existing tools
                    current_tools = list(assistant_obj.tools)
                    current_tools.append({"type": "file_search"})
                    
                    # Update the assistant
                    client.beta.assistants.update(
                        assistant_id=assistant,
                        tools=current_tools,
                        tool_resources={"file_search": {"vector_store_ids": vector_store_ids}}
                    )
                    logging.info(f"Added file_search tool to assistant {assistant}")
                else:
                    # Just update the vector store IDs if needed
                    client.beta.assistants.update(
                        assistant_id=assistant,
                        tool_resources={"file_search": {"vector_store_ids": vector_store_ids}}
                    )
                    logging.info(f"Updated vector_store_ids for assistant {assistant}")
            except Exception as e:
                logging.error(f"Error updating assistant with file_search: {e}")
                # Continue without failing the whole request
                
        # Handle image files
        elif is_image and session:
            analysis_text = await image_analysis_internal(file_content, filename, prompt)
            client.beta.threads.messages.create(
                thread_id=session,
                role="user",
                content=f"Analysis result for uploaded image '{filename}':\n{analysis_text}"
            )
            uploaded_file_details = {
                "message": "Image successfully analyzed and analysis added to thread.",
                "filename": filename,
                "thread_id": session,
                "processing_method": "thread_message"
            }
            
            # Add file awareness message
            if session:
                await add_file_awareness_internal(
                    thread_id=session, 
                    file_info={
                        "name": filename,
                        "type": "image",
                        "processing_method": "thread_message"
                    }
                )
                
            logging.info(f"Analyzed image '{filename}' and added to thread {session}")
        elif is_image:
            uploaded_file_details = {
                "message": "Image uploaded but not analyzed as no session/thread ID was provided.",
                "filename": filename,
                "processing_method": "skipped_analysis"
            }
            logging.warning(f"Image '{filename}' uploaded for assistant {assistant} but no thread ID provided.")
        else:
            uploaded_file_details = {
                "message": "Unsupported file type. Please upload PDF, DOC, DOCX, TXT files, or images.",
                "filename": filename,
                "type": "unsupported",
                "processing_method": "rejected"
            }
            logging.warning(f"Rejected unsupported file type: {filename}")

        # --- Update Context (if provided and thread exists) ---
        if context and session:
            await update_context_internal(client, session, context)

        return uploaded_file_details
    
    except Exception as e:
        logging.error(f"Error uploading file '{filename}' for assistant {assistant}: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to upload or process file: {str(e)}")
    finally:
        # Clean up temporary file
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except OSError as e:
                logging.error(f"Error removing temporary file {file_path}: {e}")

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
    This function handles both streaming and non-streaming modes.
    
    Args:
        client: Azure OpenAI client instance
        session: Thread ID
        prompt: User message
        assistant: Assistant ID
        stream_output: If True, returns a streaming response, otherwise collects and returns full response
        
    Returns:
        Either a streaming response generator or a dictionary with the full response
    """
    try:
        # Validate resources if provided 
        if session or assistant:
            validation = await validate_resources(client, session, assistant)
            
            # Create new thread if invalid
            if session and not validation["thread_valid"]:
                logging.warning(f"Invalid thread ID: {session}, creating a new one")
                try:
                    thread = client.beta.threads.create()
                    session = thread.id
                    logging.info(f"Created recovery thread: {session}")
                except Exception as e:
                    logging.error(f"Failed to create recovery thread: {e}")
                    raise HTTPException(status_code=500, detail="Failed to create a valid conversation thread")
            
            # Create new assistant if invalid
            if assistant and not validation["assistant_valid"]:
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
        
        # Create defaults if not provided
        if not assistant:
            logging.warning(f"No assistant ID provided for /{('conversation' if stream_output else 'chat')}, creating a default one.")
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
            logging.warning(f"No session (thread) ID provided for /{('conversation' if stream_output else 'chat')}, creating a new one.")
            try:
                thread = client.beta.threads.create()
                session = thread.id
            except Exception as e:
                logging.error(f"Failed to create default thread: {e}")
                raise HTTPException(status_code=500, detail="Failed to create default thread")

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
            max_retries = 3
            retry_delay = 2  # seconds
            success = False
            
            for attempt in range(max_retries):
                try:
                    if active_run and run_id:
                        # If there's an active run, check if it's still active or can be cancelled
                        try:
                            run_status = client.beta.threads.runs.retrieve(thread_id=session, run_id=run_id)
                            if run_status.status in ["in_progress", "queued"]:
                                # Option 1: Cancel the run
                                client.beta.threads.runs.cancel(thread_id=session, run_id=run_id)
                                logging.info(f"Cancelled active run {run_id} to allow new message")
                                time.sleep(1)  # Brief delay after cancellation
                            elif run_status.status == "requires_action":
                                # For requires_action, we can submit empty tool outputs to move forward
                                client.beta.threads.runs.submit_tool_outputs(
                                    thread_id=session,
                                    run_id=run_id,
                                    tool_outputs=[{"tool_call_id": "dummy", "output": "Cancelled by new request"}]
                                )
                                logging.info(f"Submitted empty tool outputs to finish run {run_id}")
                                time.sleep(1)  # Brief delay after submission
                            # If run is already completed or failed, we can proceed
                        except Exception as run_e:
                            logging.warning(f"Error handling active run: {run_e}")
                            # Continue anyway - we'll try to add message

                    # Try to add the message
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
                        logging.warning(f"Failed to add message (attempt {attempt+1}), run is active. Retrying in {retry_delay}s: {e}")
                        time.sleep(retry_delay)
                        retry_delay *= 2  # Exponential backoff
                    else:
                        logging.error(f"Failed to add message to thread {session}: {e}")
                        if attempt == max_retries - 1:
                            raise HTTPException(status_code=500, detail="Failed to add message to conversation thread")
            
            if not success:
                raise HTTPException(status_code=500, detail="Failed to add message to conversation thread after retries")
        
        # Handle non-streaming mode (/chat endpoint)
        if not stream_output:
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
                            time.sleep(poll_interval)
                            
                    except Exception as poll_e:
                        logging.error(f"Error polling run status (attempt {attempt+1}): {poll_e}")
                        time.sleep(poll_interval)
                        
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
        
        # For streaming mode (/conversation endpoint)
        # Create run and stream the response
        async def async_generator():
            try:
                streaming_run = client.beta.threads.runs.create(
                    thread_id=session,
                    assistant_id=assistant,
                    stream=True
                )
                
                async for event in streaming_run:
                    if event.event == "thread.message.delta" and hasattr(event.data, "delta") and hasattr(event.data.delta, "content") and event.data.delta.content:
                        for content in event.data.delta.content:
                            if content.type == "text" and hasattr(content.text, "value"):
                                yield content.text.value
                                
            except Exception as e:
                logging.error(f"Error in streaming generation: {e}")
                yield "\n[ERROR] An error occurred while generating the response. Please try again.\n"
        
        # Return streaming response
        return async_generator()
        
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

# ----- Teams Bot API Endpoints -----

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
