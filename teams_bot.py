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

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# API Base URL - Update this to your FastAPI deployment
API_BASE_URL = "https://copilotv2.azurewebsites.net"

# Bot credentials - Replace with your values from Azure Bot registration
APP_ID = os.environ.get("MicrosoftAppId", "")
APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")

# Configure the Bot Framework adapter
SETTINGS = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)

# Error handler
async def on_error(context: TurnContext, error: Exception):
    logger.error(f"Error: {str(error)}")
    await context.send_activity("Sorry, an error occurred. Please try again.")
    
ADAPTER.on_turn_error = on_error

# Store session state (in-memory for demo, use Azure Storage or Redis for production)
SESSION_STORE = {}

# Initialize or get session data
def get_session_data(user_id):
    if user_id not in SESSION_STORE:
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
        if turn_context.activity.type == ActivityTypes.message:
            await self.on_message_activity(turn_context)
        elif turn_context.activity.type == ActivityTypes.conversation_update:
            # Handle conversation update (e.g., bot added to conversation)
            for member in turn_context.activity.members_added or []:
                if member.id != turn_context.activity.recipient.id:
                    # New user added - send welcome message
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
        session_data = get_session_data(user_session_id)
        
        # Get the message text
        message_text = turn_context.activity.text.strip() if turn_context.activity.text else ""
        
        # Check for file attachments
        if turn_context.activity.attachments and len(turn_context.activity.attachments) > 0:
            await self._process_file_attachment(turn_context, session_data)
            return
        
        # Check for commands
        if message_text.startswith("/"):
            await self._process_command(turn_context, message_text, session_data)
            return
        
        # Process regular message as conversation
        await self._process_conversation(turn_context, message_text, session_data)
    
    async def _process_file_attachment(self, turn_context: TurnContext, session_data):
        # Process each attachment
        for attachment in turn_context.activity.attachments:
            # Check file type
            content_type = attachment.content_type
            file_name = attachment.name or "uploaded_file"
            
            await turn_context.send_activity(f"Processing your file: {file_name}...")
            
            try:
                # Download the file content
                file_content = await self._get_file_content(turn_context, attachment)
                if not file_content:
                    await turn_context.send_activity("Could not download the file. Please try again.")
                    continue
                
                # Create a temporary file
                with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_name)[1]) as temp_file:
                    temp_file.write(file_content)
                    temp_path = temp_file.name
                
                # Determine appropriate API call based on session state
                if not session_data["assistant_id"]:
                    # Initialize a new chat with the file
                    await turn_context.send_activity("Creating a new assistant with your file...")
                    
                    with open(temp_path, 'rb') as file:
                        files = {'file': (file_name, file)}
                        response = requests.post(f"{API_BASE_URL}/initiate-chat", files=files)
                    
                    if response.status_code == 200:
                        data = response.json()
                        session_data["assistant_id"] = data["assistant"]
                        session_data["session_id"] = data["session"]
                        session_data["vector_store_id"] = data["vector_store"]
                        session_data["uploaded_files"].append(file_name)
                        
                        await turn_context.send_activity(f"✅ Created a new assistant and uploaded {file_name}. Ask me anything about this file!")
                    else:
                        await turn_context.send_activity(f"❌ Failed to initialize chat with your file. Error: {response.status_code}")
                else:
                    # Upload to existing session
                    with open(temp_path, 'rb') as file:
                        files = {'file': (file_name, file)}
                        data = {"assistant": session_data["assistant_id"]}
                        
                        if session_data["session_id"]:
                            data["session"] = session_data["session_id"]
                            
                        response = requests.post(f"{API_BASE_URL}/upload-file", files=files, data=data)
                    
                    if response.status_code == 200:
                        session_data["uploaded_files"].append(file_name)
                        await turn_context.send_activity(f"✅ File '{file_name}' uploaded successfully! Ask me anything about this file.")
                    else:
                        await turn_context.send_activity(f"❌ Failed to upload file. Error: {response.status_code}")
                
                # Clean up the temporary file
                try:
                    os.unlink(temp_path)
                except:
                    pass
                    
            except Exception as e:
                logger.error(f"Error processing file: {str(e)}")
                await turn_context.send_activity(f"❌ Error processing your file: {str(e)}")
    
    async def _get_file_content(self, turn_context: TurnContext, attachment: Attachment):
        try:
            # If content is already available in the attachment
            if hasattr(attachment, 'content') and attachment.content:
                if isinstance(attachment.content, str):
                    return attachment.content.encode('utf-8')
                return attachment.content
            
            # If there's a content URL, download the file
            if attachment.content_url:
                connector = turn_context.adapter.create_connector_client(turn_context.activity.service_url)
                
                if "sharepoint" in attachment.content_url.lower() or "teams" in attachment.content_url.lower():
                    # For Teams/Sharepoint files, use connector client for authentication
                    response = await connector.conversations.get_attachment_file(
                        attachment.content_url
                    )
                    return response
                else:
                    # For other URLs, try a direct download
                    async with aiohttp.ClientSession() as session:
                        async with session.get(attachment.content_url) as response:
                            if response.status == 200:
                                return await response.read()
            
            # If we couldn't get the content
            return None
        except Exception as e:
            logger.error(f"Error downloading file: {str(e)}")
            return None
    
    async def _process_command(self, turn_context: TurnContext, command: str, session_data):
        cmd = command.lower()
        
        if cmd == "/start" or cmd == "/new":
            # Start a new chat
            await turn_context.send_activity("Creating a new assistant...")
            
            response = requests.post(f"{API_BASE_URL}/initiate-chat")
            
            if response.status_code == 200:
                data = response.json()
                session_data["assistant_id"] = data["assistant"]
                session_data["session_id"] = data["session"]
                session_data["vector_store_id"] = data["vector_store"]
                session_data["uploaded_files"] = []
                
                await turn_context.send_activity("✅ Created a new assistant. How can I help you today?")
            else:
                await turn_context.send_activity(f"❌ Failed to create assistant. Error: {response.status_code}")
        
        elif cmd == "/newthread":
            # Create a new thread with existing assistant
            if not session_data["assistant_id"]:
                await turn_context.send_activity("❌ Please start a new chat first using /new")
                return
            
            await turn_context.send_activity("Creating a new thread...")
            
            data = {
                "assistant": session_data["assistant_id"],
                "vector_store": session_data["vector_store_id"]
            }
            
            response = requests.post(f"{API_BASE_URL}/co-pilot", data=data)
            
            if response.status_code == 200:
                data = response.json()
                session_data["session_id"] = data["session"]
                
                await turn_context.send_activity("✅ Created a new thread. How can I help you today?")
            else:
                await turn_context.send_activity(f"❌ Failed to create new thread. Error: {response.status_code}")
        
        elif cmd == "/clear":
            # Clear the current thread history
            if not session_data["session_id"]:
                await turn_context.send_activity("❌ No active thread to clear.")
                return
                
            await turn_context.send_activity("Clearing chat history for current thread...")
            
            # We don't have a direct clear-chat endpoint, so we'll just create a new thread
            data = {
                "assistant": session_data["assistant_id"],
                "vector_store": session_data["vector_store_id"]
            }
            
            response = requests.post(f"{API_BASE_URL}/co-pilot", data=data)
            
            if response.status_code == 200:
                data = response.json()
                session_data["session_id"] = data["session"]
                
                await turn_context.send_activity("✅ Chat history cleared. You're now in a new thread.")
            else:
                await turn_context.send_activity(f"❌ Failed to clear chat history. Error: {response.status_code}")
        
        elif cmd == "/files":
            # Show uploaded files
            if not session_data["uploaded_files"]:
                await turn_context.send_activity("No files have been uploaded in this session.")
            else:
                files_list = "\n".join([f"- {file}" for file in session_data["uploaded_files"]])
                await turn_context.send_activity(f"Uploaded files:\n{files_list}")
        
        elif cmd == "/help":
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
        
        else:
            await turn_context.send_activity(f"Unknown command: {command}\nUse /help to see available commands.")
    
    async def _process_conversation(self, turn_context: TurnContext, message: str, session_data):
        # Check if we have an active session
        if not session_data["assistant_id"] or not session_data["session_id"]:
            # Initialize a new chat
            await turn_context.send_activity("Starting a new chat session...")
            
            response = requests.post(f"{API_BASE_URL}/initiate-chat")
            
            if response.status_code == 200:
                data = response.json()
                session_data["assistant_id"] = data["assistant"]
                session_data["session_id"] = data["session"]
                session_data["vector_store_id"] = data["vector_store"]
                
                await turn_context.send_activity("✅ Created a new assistant. Now processing your message...")
            else:
                await turn_context.send_activity(f"❌ Failed to create assistant. Error: {response.status_code}")
                return
        
        # Show typing indicator
        await turn_context.send_activities([Activity(type=ActivityTypes.typing)])
        
        # Use the non-streaming chat endpoint
        params = {
            "session": session_data["session_id"],
            "assistant": session_data["assistant_id"],
            "prompt": message,
        }
        
        try:
            # Make the request to the API
            response = requests.get(f"{API_BASE_URL}/chat", params=params)
            
            if response.status_code == 200:
                data = response.json()
                response_text = data.get("response", "I didn't receive a proper response. Please try again.")
                
                # Check if the response is too long for a single message
                if len(response_text) > 4000:
                    # Split into chunks
                    chunks = [response_text[i:i+4000] for i in range(0, len(response_text), 4000)]
                    for chunk in chunks:
                        await turn_context.send_activity(chunk)
                else:
                    # Send the full response
                    await turn_context.send_activity(response_text)
            else:
                error_text = f"❌ Failed to get a response. Status code: {response.status_code}"
                try:
                    error_data = response.json()
                    error_text += f"\nError: {json.dumps(error_data)}"
                except:
                    error_text += f"\nError: {response.text[:100]}"
                
                await turn_context.send_activity(error_text)
        
        except Exception as e:
            logger.error(f"Error in conversation: {str(e)}")
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
    bot_response = await ADAPTER.process_activity(activity, auth_header, bot.on_turn)
    if bot_response:
        return Response(body=json.dumps(bot_response.body), status=bot_response.status)
    return Response(status=201)

# Create the bot instance
bot = ProductManagementBot()

# Set up HTTP route for Bot Framework
APP.router.add_post("/api/messages", messages)

# Health check endpoint
async def health_check(req: Request) -> Response:
    return Response(text="Bot is running!", status=200)

APP.router.add_get("/", health_check)

# Run the server
if __name__ == "__main__":
    try:
        port = int(os.environ.get("PORT", 3978))
        web.run_app(APP, host="0.0.0.0", port=port)
    except Exception as e:
        logger.error(f"Error running app: {e}")
