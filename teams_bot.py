import os
import sys
import traceback
import requests
import json
import tempfile
from botbuilder.core import BotFrameworkAdapter, TurnContext, MessageFactory
from botbuilder.schema import Activity, ActivityTypes, Attachment
from botbuilder.integration.aiohttp import CloudAdapter, ConfigurationBotFrameworkAuthentication
from aiohttp import web
from aiohttp.web import Request, Response

# Your FastAPI backend URL - already deployed
API_BASE_URL = "https://copilotv2.azurewebsites.net"

# Dictionary to store conversation state for each user
# Key: conversation_id, Value: dict with assistant_id, session_id, etc.
conversation_states = {}

# App credentials from environment variables
APP_ID = os.environ.get("MicrosoftAppId", "")
APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")

# Check if credentials are provided
if not APP_ID or not APP_PASSWORD:
    print("ERROR: MicrosoftAppId and MicrosoftAppPassword must be set in environment variables")
    print("These should be configured in the Azure App Service Configuration/Application Settings")

# For aiohttp setup with Bot Framework
CONFIG_AUTH_PROVIDER = ConfigurationBotFrameworkAuthentication(
    app_id=APP_ID,
    app_password=APP_PASSWORD
)
ADAPTER = CloudAdapter(CONFIG_AUTH_PROVIDER)

# Main bot message handler
async def bot_main(req: Request) -> Response:
    if "application/json" in req.headers["Content-Type"]:
        body = await req.json()
    else:
        return Response(status=415)

    # Create a TurnContext from the incoming activity
    activity = Activity().deserialize(body)
    
    try:
        response = Response()
        await ADAPTER.process(req, response, activity, bot_logic)
        return response
    except Exception as e:
        print(f"Error processing activity: {e}")
        traceback.print_exc()
        return Response(status=500)

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
            await turn_context.send_activity(MessageFactory.typing())
            
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
            print(f"Error processing file: {str(e)}")
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
        print(f"Error downloading attachment: {str(e)}")
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
    await turn_context.send_activity(MessageFactory.typing())
    
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
        print(f"Error in handle_text_message: {str(e)}")
        traceback.print_exc()

# Initialize chat with the backend
async def initialize_chat(turn_context: TurnContext, state, context=None):
    try:
        # Send typing indicator
        await turn_context.send_activity(MessageFactory.typing())
        
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
        print(f"Error in initialize_chat: {str(e)}")
        traceback.print_exc()

# Send a message without user input (used after file upload or initialization)
async def send_message(turn_context: TurnContext, state):
    try:
        # Send typing indicator
        await turn_context.send_activity(MessageFactory.typing())
        
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
        print(f"Error in send_message: {str(e)}")
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

# Set up the web app
app = web.Application()
app.router.add_post("/api/messages", bot_main)

if __name__ == "__main__":
    try:
        # Get port from environment for Azure deployment
        PORT = int(os.environ.get("PORT", 8000))
        
        # Start the web server
        web.run_app(app, host="0.0.0.0", port=PORT)
    except Exception as e:
        print(f"Error starting bot: {e}")
        traceback.print_exc()
        sys.exit()
