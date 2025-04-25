import os
import requests
import json
import logging
import tempfile
from botbuilder.core import BotFrameworkAdapter, BotFrameworkAdapterSettings, TurnContext, CardFactory
from botbuilder.schema import Activity, Attachment, HeroCard, CardAction, ActionTypes
from botbuilder.schema import ResourceResponse, ActivityTypes, MessageFactory
from aiohttp import web
from aiohttp.web import Request, Response
import aiohttp
import asyncio
import traceback

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# Log startup for debugging
logger.info("Bot application starting...")

# API Base URL - Update this to your FastAPI deployment
API_BASE_URL = "https://copilotv2.azurewebsites.net"
logger.info(f"Using API base URL: {API_BASE_URL}")

# Bot credentials - Replace with your values from Azure Bot registration
APP_ID = os.environ.get("MicrosoftAppId", "")
APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")
logger.info(f"Bot credentials - App ID exists: {bool(APP_ID)}, Password exists: {bool(APP_PASSWORD)}")

# Configure the Bot Framework adapter
SETTINGS = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)

# Error handler
async def on_error(context: TurnContext, error: Exception):
    logger.error(f"Error: {str(error)}")
    logger.error(traceback.format_exc())
    await context.send_activity("Sorry, an error occurred. Please try again.")
    
ADAPTER.on_turn_error = on_error

# Store session state (in-memory for demo, use Azure Storage or Redis for production)
SESSION_STORE = {}

# Initialize or get session data
def get_session_data(user_id):
    if user_id not in SESSION_STORE:
        logger.info(f"Creating new session data for user {user_id}")
        SESSION_STORE[user_id] = {
            "assistant_id": None,
            "session_id": None,
            "vector_store_id": None,
            "uploaded_files": []
        }
    return SESSION_STORE[user_id]

# Main bot functionality
class ProductManagementBot:
    async def on_turn(self, turn_context: TurnContext):
        logger.info(f"Processing activity type: {turn_context.activity.type}")
        
        if turn_context.activity.type == ActivityTypes.message:
            await self.on_message_activity(turn_context)
        elif turn_context.activity.type == ActivityTypes.conversation_update:
            # Handle conversation update (e.g., bot added to conversation)
            for member in turn_context.activity.members_added or []:
                if member.id != turn_context.activity.recipient.id:
                    # New user added - send welcome message
                    logger.info("Sending welcome message to new user")
                    await self.send_welcome_message(turn_context)

    async def send_welcome_message(self, turn_context: TurnContext):
        welcome_text = (
            "# Welcome to the Product Management Bot!\n\n"
            "I can help you create documentation and analyze various file types. "
            "You can interact with me in the following ways:\n\n"
            "- **Upload a file** for analysis (CSV, Excel, PDF, etc.)\n"
            "- **Chat with me** to create documentation\n"
            "- **Start a new session** with `/new`\n"
            "- **Create a new thread** with `/newthread`\n\n"
            "Type `/help` for more details on what I can do!"
        )
        await turn_context.send_activity(welcome_text)
        
        # Create a card with buttons for common actions
        card = HeroCard(
            title="What would you like to do?",
            buttons=[
                CardAction(
                    type=ActionTypes.im_back,
                    title="Start New Session",
                    value="/new"
                ),
                CardAction(
                    type=ActionTypes.im_back,
                    title="Help",
                    value="/help"
                )
            ]
        )
        
        message = MessageFactory.attachment(CardFactory.hero_card(card))
        await turn_context.send_activity(message)

    async def on_message_activity(self, turn_context: TurnContext):
        user_id = turn_context.activity.from_property.id
        channel_id = turn_context.activity.channel_id
        conversation_id = turn_context.activity.conversation.id
        
        # Create a unique ID for this user-conversation combo
        user_session_id = f"{user_id}_{conversation_id}_{channel_id}"
        logger.info(f"Processing message for session: {user_session_id}")
        
        session_data = get_session_data(user_session_id)
        
        # Get the message text
        message_text = turn_context.activity.text.strip() if turn_context.activity.text else ""
        
        # Check for file attachments
        if turn_context.activity.attachments and len(turn_context.activity.attachments) > 0:
            logger.info(f"Processing {len(turn_context.activity.attachments)} file attachments")
            await self._process_file_attachment(turn_context, session_data)
            return
        
        # Check for commands
        if message_text.startswith("/"):
            logger.info(f"Processing command: {message_text}")
            await self._process_command(turn_context, message_text, session_data)
            return
        
        # Process regular message as conversation
        logger.info("Processing regular conversation message")
        await self._process_conversation(turn_context, message_text, session_data)
    
    async def _process_file_attachment(self, turn_context: TurnContext, session_data):
        # Process each attachment
        for attachment in turn_context.activity.attachments:
            # Check file type
            content_type = attachment.content_type
            file_name = attachment.name or "uploaded_file"
            
            logger.info(f"Processing file attachment: {file_name} ({content_type})")
            await turn_context.send_activity(f"Processing your file: {file_name}...")
            
            try:
                # Download the file content
                file_content = await self._get_file_content(turn_context, attachment)
                if not file_content:
                    logger.error(f"Could not download file: {file_name}")
                    await turn_context.send_activity("Could not download the file. Please try again.")
                    continue
                
                logger.info(f"Downloaded file content: {len(file_content)} bytes")
                
                # Create a temporary file
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_name)[1]) as temp_file:
                    temp_file.write(file_content)
                    temp_path = temp_file.name
                    logger.info(f"Saved to temporary file: {temp_path}")
                
                # Determine appropriate API call based on session state
                if not session_data["assistant_id"]:
                    # Initialize a new chat with the file
                    logger.info("Creating new assistant with file")
                    await turn_context.send_activity("Creating a new assistant with your file...")
                    
                    with open(temp_path, 'rb') as file:
                        files = {'file': (file_name, file)}
                        logger.info(f"Calling /initiate-chat endpoint with file: {file_name}")
                        
                        try:
                            response = requests.post(f"{API_BASE_URL}/initiate-chat", files=files)
                            logger.info(f"Received response: status code {response.status_code}")
                            
                            if response.status_code == 200:
                                data = response.json()
                                logger.info(f"Response data: {json.dumps(data)}")
                                
                                session_data["assistant_id"] = data["assistant"]
                                session_data["session_id"] = data["session"]
                                session_data["vector_store_id"] = data["vector_store"]
                                session_data["uploaded_files"].append(file_name)
                                
                                logger.info(f"Updated session data with new IDs")
                                await turn_context.send_activity(f"✅ Created a new assistant and uploaded {file_name}. Ask me anything about this file!")
                            else:
                                logger.error(f"Failed API call: {response.status_code}, Response: {response.text[:200]}")
                                await turn_context.send_activity(f"❌ Failed to initialize chat with your file. Error: {response.status_code}")
                        except Exception as e:
                            logger.error(f"Exception calling API: {str(e)}")
                            logger.error(traceback.format_exc())
                            await turn_context.send_activity(f"❌ Error connecting to backend service: {str(e)}")
                else:
                    # Upload to existing session
                    logger.info(f"Uploading file to existing session: assistant={session_data['assistant_id']}, session={session_data['session_id']}")
                    with open(temp_path, 'rb') as file:
                        files = {'file': (file_name, file)}
                        data = {"assistant": session_data["assistant_id"]}
                        
                        if session_data["session_id"]:
                            data["session"] = session_data["session_id"]
                        
                        logger.info(f"Calling /upload-file endpoint with file: {file_name}")
                        try:
                            response = requests.post(f"{API_BASE_URL}/upload-file", files=files, data=data)
                            logger.info(f"Received response: status code {response.status_code}")
                            
                            if response.status_code == 200:
                                session_data["uploaded_files"].append(file_name)
                                await turn_context.send_activity(f"✅ File '{file_name}' uploaded successfully! Ask me anything about this file.")
                            else:
                                logger.error(f"Failed API call: {response.status_code}, Response: {response.text[:200]}")
                                await turn_context.send_activity(f"❌ Failed to upload file. Error: {response.status_code}")
                        except Exception as e:
                            logger.error(f"Exception calling API: {str(e)}")
                            logger.error(traceback.format_exc())
                            await turn_context.send_activity(f"❌ Error connecting to backend service: {str(e)}")
                
                # Clean up the temporary file
                try:
                    os.unlink(temp_path)
                    logger.info(f"Cleaned up temporary file: {temp_path}")
                except Exception as e:
                    logger.warning(f"Could not clean up temporary file: {str(e)}")
                    
            except Exception as e:
                logger.error(f"Error processing file: {str(e)}")
                logger.error(traceback.format_exc())
                await turn_context.send_activity(f"❌ Error processing your file: {str(e)}")
    
    async def _get_file_content(self, turn_context: TurnContext, attachment: Attachment):
        try:
            logger.info(f"Getting file content: content_url={bool(attachment.content_url)}, content present={hasattr(attachment, 'content') and bool(attachment.content)}")
            
            # If content is already available in the attachment
            if hasattr(attachment, 'content') and attachment.content:
                logger.info("Using attachment content directly")
                if isinstance(attachment.content, str):
                    return attachment.content.encode('utf-8')
                return attachment.content
            
            # If there's a content URL, download the file
            if attachment.content_url:
                logger.info(f"Downloading from content URL: {attachment.content_url}")
                
                try:
                    connector = turn_context.adapter.create_connector_client(turn_context.activity.service_url)
                    
                    if "sharepoint" in attachment.content_url.lower() or "teams" in attachment.content_url.lower():
                        # For Teams/Sharepoint files, use connector client for authentication
                        logger.info("Using connector client for Teams/Sharepoint file")
                        response = await connector.conversations.get_attachment_file(
                            attachment.content_url
                        )
                        logger.info(f"Downloaded file with connector client, size: {len(response) if response else 'unknown'}")
                        return response
                    else:
                        # For other URLs, try a direct download
                        logger.info("Using direct HTTP download")
                        async with aiohttp.ClientSession() as session:
                            async with session.get(attachment.content_url) as response:
                                if response.status == 200:
                                    content = await response.read()
                                    logger.info(f"Downloaded file, size: {len(content)}")
                                    return content
                                else:
                                    logger.error(f"Failed to download file, status: {response.status}")
                except Exception as e:
                    logger.error(f"Exception downloading file: {str(e)}")
                    logger.error(traceback.format_exc())
            
            # If we couldn't get the content
            logger.error("Could not get file content")
            return None
        except Exception as e:
            logger.error(f"Error in _get_file_content: {str(e)}")
            logger.error(traceback.format_exc())
            return None
    
    async def _process_command(self, turn_context: TurnContext, command: str, session_data):
        cmd = command.lower()
        
        if cmd == "/start" or cmd == "/new":
            # Start a new chat
            logger.info("Processing /new command")
            await turn_context.send_activity("Creating a new assistant...")
            
            try:
                response = requests.post(f"{API_BASE_URL}/initiate-chat")
                logger.info(f"Received response: status code {response.status_code}")
                
                if response.status_code == 200:
                    data = response.json()
                    logger.info(f"Response data: {json.dumps(data)}")
                    
                    session_data["assistant_id"] = data["assistant"]
                    session_data["session_id"] = data["session"]
                    session_data["vector_store_id"] = data["vector_store"]
                    session_data["uploaded_files"] = []
                    
                    await turn_context.send_activity("✅ Created a new assistant. How can I help you today?")
                else:
                    logger.error(f"Failed API call: {response.status_code}, Response: {response.text[:200]}")
                    await turn_context.send_activity(f"❌ Failed to create assistant. Error: {response.status_code}")
            except Exception as e:
                logger.error(f"Exception calling API: {str(e)}")
                logger.error(traceback.format_exc())
                await turn_context.send_activity(f"❌ Error connecting to backend service: {str(e)}")
        
        elif cmd == "/newthread":
            # Create a new thread with existing assistant
            logger.info("Processing /newthread command")
            
            if not session_data["assistant_id"]:
                logger.warning("No assistant ID available for /newthread")
                await turn_context.send_activity("❌ Please start a new chat first using /new")
                return
            
            await turn_context.send_activity("Creating a new thread...")
            
            try:
                data = {
                    "assistant": session_data["assistant_id"],
                    "vector_store": session_data["vector_store_id"]
                }
                
                logger.info(f"Calling /co-pilot endpoint with data: {json.dumps(data)}")
                response = requests.post(f"{API_BASE_URL}/co-pilot", data=data)
                logger.info(f"Received response: status code {response.status_code}")
                
                if response.status_code == 200:
                    data = response.json()
                    logger.info(f"Response data: {json.dumps(data)}")
                    session_data["session_id"] = data["session"]
                    
                    await turn_context.send_activity("✅ Created a new thread. How can I help you today?")
                else:
                    logger.error(f"Failed API call: {response.status_code}, Response: {response.text[:200]}")
                    await turn_context.send_activity(f"❌ Failed to create new thread. Error: {response.status_code}")
            except Exception as e:
                logger.error(f"Exception calling API: {str(e)}")
                logger.error(traceback.format_exc())
                await turn_context.send_activity(f"❌ Error connecting to backend service: {str(e)}")
        
        elif cmd == "/clear":
            # Clear the current thread history
            logger.info("Processing /clear command")
            
            if not session_data["session_id"]:
                logger.warning("No session ID available for /clear")
                await turn_context.send_activity("❌ No active thread to clear.")
                return
                
            await turn_context.send_activity("Clearing chat history for current thread...")
            
            try:
                # We don't have a direct clear-chat endpoint, so we'll just create a new thread
                data = {
                    "assistant": session_data["assistant_id"],
                    "vector_store": session_data["vector_store_id"]
                }
                
                logger.info(f"Calling /co-pilot endpoint to create new thread: {json.dumps(data)}")
                response = requests.post(f"{API_BASE_URL}/co-pilot", data=data)
                logger.info(f"Received response: status code {response.status_code}")
                
                if response.status_code == 200:
                    data = response.json()
                    logger.info(f"Response data: {json.dumps(data)}")
                    session_data["session_id"] = data["session"]
                    
                    await turn_context.send_activity("✅ Chat history cleared. You're now in a new thread.")
                else:
                    logger.error(f"Failed API call: {response.status_code}, Response: {response.text[:200]}")
                    await turn_context.send_activity(f"❌ Failed to clear chat history. Error: {response.status_code}")
            except Exception as e:
                logger.error(f"Exception calling API: {str(e)}")
                logger.error(traceback.format_exc())
                await turn_context.send_activity(f"❌ Error connecting to backend service: {str(e)}")
        
        elif cmd == "/files":
            # Show uploaded files
            logger.info("Processing /files command")
            
            if not session_data["uploaded_files"]:
                await turn_context.send_activity("No files have been uploaded in this session.")
            else:
                files_list = "\n".join([f"- {file}" for file in session_data["uploaded_files"]])
                await turn_context.send_activity(f"Uploaded files:\n{files_list}")
        
        elif cmd == "/help":
            logger.info("Processing /help command")
            
            help_text = """
# Available Commands

- **/new** or **/start** - Start a new chat with a new assistant
- **/newthread** - Create a new thread with the current assistant
- **/clear** - Clear chat history (creates a new thread)
- **/files** - List uploaded files
- **/help** - Show this help message

## Features
- **Upload files** directly to be analyzed (CSV, Excel, PDF, images, etc.)
- **Create documentation** by chatting with the assistant
- **Analyze data** by uploading data files

## Tips
- If you upload a file, you can ask questions about it
- Start a new thread when changing topics
- Clear chat history to start fresh while keeping the context
            """
            await turn_context.send_activity(help_text)
        
        elif cmd == "/debug":
            # Debug command - show session info (admin only)
            logger.info("Processing /debug command")
            debug_info = f"""
# Debug Information
- Assistant ID: {session_data['assistant_id']}
- Session ID: {session_data['session_id']}
- Vector Store ID: {session_data['vector_store_id']}
- Uploaded files: {', '.join(session_data['uploaded_files']) if session_data['uploaded_files'] else 'None'}
- API Base URL: {API_BASE_URL}
            """
            await turn_context.send_activity(debug_info)
        
        else:
            logger.warning(f"Unknown command: {command}")
            await turn_context.send_activity(f"Unknown command: {command}\nUse /help to see available commands.")
    
    async def _process_conversation(self, turn_context: TurnContext, message: str, session_data):
        # Check if we have an active session
        if not session_data["assistant_id"] or not session_data["session_id"]:
            # Initialize a new chat
            logger.info("No active session, creating new chat")
            await turn_context.send_activity("Starting a new chat session...")
            
            try:
                response = requests.post(f"{API_BASE_URL}/initiate-chat")
                logger.info(f"Received response: status code {response.status_code}")
                
                if response.status_code == 200:
                    data = response.json()
                    logger.info(f"Response data: {json.dumps(data)}")
                    
                    session_data["assistant_id"] = data["assistant"]
                    session_data["session_id"] = data["session"]
                    session_data["vector_store_id"] = data["vector_store"]
                    
                    await turn_context.send_activity("✅ Created a new assistant. Now processing your message...")
                else:
                    logger.error(f"Failed API call: {response.status_code}, Response: {response.text[:200]}")
                    await turn_context.send_activity(f"❌ Failed to create assistant. Error: {response.status_code}")
                    return
            except Exception as e:
                logger.error(f"Exception calling API: {str(e)}")
                logger.error(traceback.format_exc())
                await turn_context.send_activity(f"❌ Error connecting to backend service: {str(e)}")
                return
        
        # Show typing indicator
        logger.info("Sending typing indicator")
        await turn_context.send_activities([Activity(type=ActivityTypes.typing)])
        
        # Use the non-streaming chat endpoint
        params = {
            "session": session_data["session_id"],
            "assistant": session_data["assistant_id"],
            "prompt": message,
        }
        
        logger.info(f"Calling /chat endpoint with params: {json.dumps(params)}")
        
        try:
            # Make the request to the API
            response = requests.get(f"{API_BASE_URL}/chat", params=params)
            logger.info(f"Received response: status code {response.status_code}")
            
            if response.status_code == 200:
                data = response.json()
                response_text = data.get("response", "I didn't receive a proper response. Please try again.")
                logger.info(f"Response length: {len(response_text)} characters")
                
                # Check if the response is too long for a single message
                if len(response_text) > 4000:
                    # Split into chunks
                    logger.info(f"Response too long, splitting into chunks")
                    chunks = [response_text[i:i+4000] for i in range(0, len(response_text), 4000)]
                    for i, chunk in enumerate(chunks):
                        logger.info(f"Sending chunk {i+1}/{len(chunks)}")
                        await turn_context.send_activity(chunk)
                else:
                    # Send the full response
                    logger.info("Sending full response")
                    await turn_context.send_activity(response_text)
            else:
                logger.error(f"Failed API call: {response.status_code}, Response: {response.text[:200]}")
                error_text = f"❌ Failed to get a response. Status code: {response.status_code}"
                try:
                    error_data = response.json()
                    error_text += f"\nError: {json.dumps(error_data)}"
                except:
                    error_text += f"\nError: {response.text[:100]}"
                
                await turn_context.send_activity(error_text)
        
        except Exception as e:
            logger.error(f"Error in conversation: {str(e)}")
            logger.error(traceback.format_exc())
            await turn_context.send_activity(f"❌ Error processing your message: {str(e)}")

# HTTP server setup
APP = web.Application()

# Process bot messages
async def messages(req: Request) -> Response:
    # Convert from aiohttp request to botbuilder activity
    if "application/json" in req.headers["Content-Type"]:
        body = await req.json()
    else:
        return Response(status=415)

    # Create an activity object from the received payload
    activity = Activity().deserialize(body)
    auth_header = req.headers["Authorization"] if "Authorization" in req.headers else ""

    # Call the bot's on_turn function and get a response
    response = Response(status=201)
    try:
        bot_response = await ADAPTER.process_activity(activity, auth_header, bot.on_turn)
        if bot_response:
            return Response(body=json.dumps(bot_response.body), status=bot_response.status)
    except Exception as e:
        logger.error(f"Error processing activity: {str(e)}")
        logger.error(traceback.format_exc())
        return Response(status=500)
    
    return response

# Create the bot instance
bot = ProductManagementBot()

# Set up HTTP route for Bot Framework
APP.router.add_post("/api/messages", messages)

# Health check endpoint
async def health_check(req: Request) -> Response:
    return Response(text="Bot is running!", status=200)

APP.router.add_get("/", health_check)

# For Azure Bot Service - add direct access to health check
APP.router.add_get("/api/messages", health_check)

# For detailed debug info
async def debug_info(req: Request) -> Response:
    debug_text = {
        "status": "running",
        "api_base_url": API_BASE_URL,
        "bot_credentials": {
            "app_id_exists": bool(APP_ID),
            "password_exists": bool(APP_PASSWORD)
        },
        "active_sessions": len(SESSION_STORE),
        "python_version": os.environ.get("PYTHONVERSION", "unknown"),
        "server_time": str(asyncio.get_event_loop().time())
    }
    return Response(text=json.dumps(debug_text, indent=2), status=200, content_type="application/json")

APP.router.add_get("/debug", debug_info)

# Run the server
if __name__ == "__main__":
    try:
        port = int(os.environ.get("PORT", 8080))
        host = os.environ.get("HOST", "0.0.0.0")
        
        logger.info(f"Starting web server on {host}:{port}")
        web.run_app(APP, host=host, port=port)
    except Exception as e:
        logger.error(f"Error running app: {e}")
        logger.error(traceback.format_exc())
