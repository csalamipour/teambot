import os
import sys
import asyncio
import aiohttp
import base64
import tempfile
from typing import List, Dict, Any, Optional
from botbuilder.core import ActivityHandler, TurnContext, CardFactory, MessageFactory
from botbuilder.schema import (
    Activity, ActivityTypes, Attachment, AttachmentData, 
    ConversationReference, HeroCard, CardImage, CardAction, 
    ActionTypes, ChannelAccount
)
from botbuilder.schema.teams import (
    FileConsentCard, FileConsentCardResponse, FileDownloadInfo,
    TeamsChannelAccount
)
from botframework.connector import Channels
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings
from botbuilder.schema.teams import MessagingExtensionAction, MessagingExtensionResult

# API endpoint URL - update with your FastAPI server address
API_BASE_URL = "copilotv2.azurewebsites.net"  # Change this to your deployed backend URL

# Store conversation state for each user
conversation_state = {}

class TeamsBot(ActivityHandler):
    def __init__(self):
        """Initialize the TeamsBot."""
        self.client_session = None
        super().__init__()
    
    async def on_turn(self, turn_context: TurnContext):
        """Process incoming activities and initialize HTTP session."""
        # Initialize aiohttp session if not already created
        if self.client_session is None:
            self.client_session = aiohttp.ClientSession()
        
        # Process the activity
        await super().on_turn(turn_context)
    
    async def on_conversation_update_activity(self, turn_context: TurnContext):
        """Handle conversation update activities."""
        # When the bot is added to a conversation
        if self._is_bot_added_activity(turn_context.activity):
            await self._send_welcome_message(turn_context)
        
        # Call the parent handler
        await super().on_conversation_update_activity(turn_context)
    
    async def on_message_activity(self, turn_context: TurnContext):
        """Handle message activities (text messages from users)."""
        # Get the conversation reference for storing state
        conversation_ref = TurnContext.get_conversation_reference(turn_context.activity)
        conversation_id = self._get_conversation_id(conversation_ref)
        
        # Initialize user state if not exists
        if conversation_id not in conversation_state:
            conversation_state[conversation_id] = {
                "assistant_id": None,
                "session_id": None,
                "vector_store_id": None,
                "thread_name": None,
                "uploaded_files": []
            }
        
        # Get the text from the activity
        text = turn_context.activity.text.strip()
        
        # Check for commands
        if text.lower() == "/start":
            await self._handle_start_command(turn_context, conversation_id)
        elif text.lower() == "/newthread":
            await self._handle_new_thread_command(turn_context, conversation_id)
        elif text.lower() == "/help":
            await self._send_help_message(turn_context)
        elif text.lower() == "/status":
            await self._handle_status_command(turn_context, conversation_id)
        else:
            # Process as a regular message for conversation
            await self._handle_conversation(turn_context, conversation_id, text)
    
    async def on_teams_file_consent_accept(self, turn_context: TurnContext, file_consent_card_response: FileConsentCardResponse):
        """Handle when a user accepts a file upload request."""
        try:
            conversation_ref = TurnContext.get_conversation_reference(turn_context.activity)
            conversation_id = self._get_conversation_id(conversation_ref)
            
            # Download the file that the user has accepted to upload
            file_data = await self._download_file(file_consent_card_response.upload_info.upload_url)
            
            # Get the file name
            file_name = file_consent_card_response.upload_info.name
            
            # Upload the file to our backend API
            await self._upload_file_to_api(turn_context, conversation_id, file_name, file_data)
            
            # Send acknowledgment
            await turn_context.send_activity(f"File '{file_name}' uploaded successfully!")
            
        except Exception as e:
            await turn_context.send_activity(f"Error processing file: {str(e)}")
    
    async def on_teams_file_consent_decline(self, turn_context: TurnContext, file_consent_card_response: FileConsentCardResponse):
        """Handle when a user declines a file upload request."""
        await turn_context.send_activity("File upload was canceled.")
    
    async def _handle_start_command(self, turn_context: TurnContext, conversation_id: str):
        """Initialize a new chat session."""
        try:
            # Show typing indicator
            await self._send_typing_indicator(turn_context)
            
            # Call initiate-chat endpoint
            async with self.client_session.post(f"{API_BASE_URL}/initiate-chat") as response:
                if response.status == 200:
                    result = await response.json()
                    
                    # Store the session data
                    conversation_state[conversation_id]["assistant_id"] = result.get("assistant")
                    conversation_state[conversation_id]["session_id"] = result.get("session")
                    conversation_state[conversation_id]["vector_store_id"] = result.get("vector_store")
                    conversation_state[conversation_id]["thread_name"] = "Default Thread"
                    
                    await turn_context.send_activity(
                        "✅ New assistant created! You can start chatting now.\n\n"
                        "📤 Upload files by attaching them to a message.\n"
                        "🧵 Create a new thread with /newthread\n"
                        "❓ Get help with /help"
                    )
                else:
                    error_text = await response.text()
                    await turn_context.send_activity(f"Error creating assistant: {error_text}")
        except Exception as e:
            await turn_context.send_activity(f"Error starting chat: {str(e)}")
    
    async def _handle_new_thread_command(self, turn_context: TurnContext, conversation_id: str):
        """Create a new thread with the existing assistant."""
        try:
            # Get current state
            state = conversation_state.get(conversation_id, {})
            assistant_id = state.get("assistant_id")
            vector_store_id = state.get("vector_store_id")
            
            if not assistant_id or not vector_store_id:
                await turn_context.send_activity("Please start a chat first with /start")
                return
            
            # Show typing indicator
            await self._send_typing_indicator(turn_context)
            
            # Call co-pilot endpoint
            data = {
                "assistant": assistant_id,
                "vector_store": vector_store_id,
                "context": "",  # Optional context could be added here
                "thread_name": f"Teams Thread {len(conversation_state)}"
            }
            
            async with self.client_session.post(
                f"{API_BASE_URL}/co-pilot", 
                data=data
            ) as response:
                if response.status == 200:
                    result = await response.json()
                    
                    # Update session ID
                    conversation_state[conversation_id]["session_id"] = result.get("session")
                    
                    await turn_context.send_activity("✅ New thread created! You can continue chatting.")
                else:
                    error_text = await response.text()
                    await turn_context.send_activity(f"Error creating new thread: {error_text}")
        except Exception as e:
            await turn_context.send_activity(f"Error creating new thread: {str(e)}")
    
    async def _handle_status_command(self, turn_context: TurnContext, conversation_id: str):
        """Show current conversation status."""
        state = conversation_state.get(conversation_id, {})
        
        if not state.get("assistant_id"):
            await turn_context.send_activity("No active chat session. Start with /start")
            return
        
        status_message = (
            "📊 **Current Status**\n\n"
            f"👨‍💼 Assistant ID: `{state.get('assistant_id', 'None')}`\n"
            f"💬 Session ID: `{state.get('session_id', 'None')}`\n"
            f"🧵 Thread Name: {state.get('thread_name', 'Default')}\n"
            f"📁 Uploaded Files: {len(state.get('uploaded_files', []))}"
        )
        
        if state.get("uploaded_files"):
            status_message += "\n\n**Files:**\n" + "\n".join(
                [f"- {file}" for file in state.get("uploaded_files", [])]
            )
        
        await turn_context.send_activity(status_message)
    
    async def _handle_conversation(self, turn_context: TurnContext, conversation_id: str, text: str):
        """Handle a regular conversation message."""
        # Get current state
        state = conversation_state.get(conversation_id, {})
        
        # Check if we have an active session
        if not state.get("session_id") or not state.get("assistant_id"):
            await turn_context.send_activity(
                "You need to start a chat first. Use /start to create a new assistant."
            )
            return
        
        # Check for file attachments
        if turn_context.activity.attachments and len(turn_context.activity.attachments) > 0:
            for attachment in turn_context.activity.attachments:
                # Handle file attachment
                if attachment.content_type != "text/html":
                    await self._handle_file_attachment(turn_context, conversation_id, attachment)
            return
        
        try:
            # Show typing indicator
            await self._send_typing_indicator(turn_context)
            
            # Call conversation endpoint
            session_id = state.get("session_id")
            assistant_id = state.get("assistant_id")
            
            # We'll use the non-streaming endpoint for Teams
            params = {
                "session": session_id,
                "assistant": assistant_id,
                "prompt": text,
            }
            
            async with self.client_session.get(
                f"{API_BASE_URL}/chat", 
                params=params
            ) as response:
                if response.status == 200:
                    result = await response.json()
                    assistant_response = result.get("response", "No response received")
                    
                    # Send the response
                    await turn_context.send_activity(assistant_response)
                else:
                    error_text = await response.text()
                    await turn_context.send_activity(f"Error getting response: {error_text}")
        except Exception as e:
            await turn_context.send_activity(f"Error processing message: {str(e)}")
    
    async def _handle_file_attachment(self, turn_context: TurnContext, conversation_id: str, attachment: Attachment):
        """Handle file attachments from the user."""
        try:
            # Get current state
            state = conversation_state.get(conversation_id, {})
            
            # Check if we have an active session
            if not state.get("assistant_id"):
                await turn_context.send_activity(
                    "You need to start a chat first before uploading files. Use /start to create a new assistant."
                )
                return
            
            # Teams doesn't provide direct file content, so we need to use FileConsentCard
            filename = attachment.name or "uploaded_file"
            
            # Create consent card for the user to approve file upload
            consent_context = { "conversationId": conversation_id, "filename": filename }
            
            # Create file consent card
            file_card = FileConsentCard(
                description=f"Please approve to upload and process '{filename}'",
                accept_context=consent_context,
                decline_context=consent_context
            )
            
            consent_attachment = CardFactory.file_consent_card(file_card)
            
            await turn_context.send_activity(MessageFactory.attachment(consent_attachment))
            
        except Exception as e:
            await turn_context.send_activity(f"Error processing file attachment: {str(e)}")
    
    async def _upload_file_to_api(self, turn_context: TurnContext, conversation_id: str, filename: str, file_data: bytes):
        """Upload a file to the backend API."""
        try:
            # Get current state
            state = conversation_state.get(conversation_id, {})
            assistant_id = state.get("assistant_id")
            session_id = state.get("session_id")
            
            if not assistant_id:
                await turn_context.send_activity("No active assistant. Please start a chat with /start first.")
                return
            
            # Show typing indicator
            await self._send_typing_indicator(turn_context)
            
            # Save file to temp location
            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(filename)[1]) as tmp:
                tmp.write(file_data)
                tmp_path = tmp.name
            
            try:
                # Create form data
                data = aiohttp.FormData()
                data.add_field('file', 
                            open(tmp_path, 'rb'),
                            filename=filename,
                            content_type='application/octet-stream')
                data.add_field('assistant', assistant_id)
                
                # Add session ID if available
                if session_id:
                    data.add_field('session', session_id)
                
                # Upload file
                async with self.client_session.post(
                    f"{API_BASE_URL}/upload-file", 
                    data=data
                ) as response:
                    if response.status == 200:
                        result = await response.json()
                        
                        # Add to uploaded files list
                        state["uploaded_files"].append(filename)
                        
                        # Notify the user
                        await turn_context.send_activity(f"File '{filename}' uploaded and processed successfully!")
                    else:
                        error_text = await response.text()
                        await turn_context.send_activity(f"Error uploading file: {error_text}")
            finally:
                # Clean up temp file
                os.unlink(tmp_path)
                
        except Exception as e:
            await turn_context.send_activity(f"Error uploading file: {str(e)}")
    
    async def _download_file(self, download_url: str) -> bytes:
        """Download a file from a URL."""
        async with self.client_session.get(download_url) as response:
            if response.status == 200:
                return await response.read()
            else:
                error_text = await response.text()
                raise Exception(f"Error downloading file: {error_text}")
    
    async def _send_welcome_message(self, turn_context: TurnContext):
        """Send a welcome message when the bot is added to a conversation."""
        welcome_text = (
            "# 👋 Welcome to the Product Management Bot!\n\n"
            "I can help you create documentation and analyze various file types.\n\n"
            "## Getting Started\n"
            "- Use **/start** to create a new assistant\n"
            "- Simply send messages to chat with me\n"
            "- Attach files to upload them for analysis\n"
            "- Use **/newthread** to create a new thread\n"
            "- Use **/help** to see all available commands\n\n"
            "Let's get started! Type **/start** to begin."
        )
        await turn_context.send_activity(welcome_text)
    
    async def _send_help_message(self, turn_context: TurnContext):
        """Send a help message with available commands."""
        help_text = (
            "# 🛠️ Product Management Bot - Help\n\n"
            "## Available Commands\n"
            "- **/start** - Create a new assistant and start a chat\n"
            "- **/newthread** - Create a new thread with the current assistant\n"
            "- **/status** - Show current session information\n"
            "- **/help** - Show this help message\n\n"
            "## Features\n"
            "- Chat with the AI assistant by typing messages\n"
            "- Upload files by attaching them to a message\n"
            "- Analyze CSV/Excel files automatically\n"
            "- Generate PRDs and other product documentation\n"
        )
        await turn_context.send_activity(help_text)
    
    async def _send_typing_indicator(self, turn_context: TurnContext):
        """Send a typing indicator to show the bot is processing."""
        typing_activity = Activity(
            type=ActivityTypes.typing,
            recipient=turn_context.activity.from_property,
            from_property=turn_context.activity.recipient
        )
        await turn_context.send_activity(typing_activity)
    
    def _is_bot_added_activity(self, activity: Activity) -> bool:
        """Check if the activity is about the bot being added to conversation."""
        members_added = activity.members_added or []
        return any(member.id == activity.recipient.id for member in members_added)
    
    def _get_conversation_id(self, conversation_ref: ConversationReference) -> str:
        """Get a unique ID for the conversation."""
        return f"{conversation_ref.channel_id}:{conversation_ref.conversation.id}"
