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
from typing import Optional, List, Dict, Any, Tuple, Union, Callable, Literal, Deque
from http import HTTPStatus
from datetime import datetime

# FastAPI imports
from fastapi import FastAPI, Request, Response, UploadFile, File, Form, HTTPException, BackgroundTasks
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware

# Azure OpenAI imports
from openai import AzureOpenAI, APIError, APIConnectionError, APITimeoutError

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

# Import Teams AI StreamingResponse class if available
try:
    from teams.streaming import StreamingResponse
    from teams.streaming.streaming_channel_data import StreamingChannelData
    from teams.streaming.streaming_entity import StreamingEntity
    from teams.ai.citations.citations import Appearance, SensitivityUsageInfo
    from teams.ai.citations import AIEntity, ClientCitation
    from teams.ai.prompts.message import Citation
    TEAMS_AI_AVAILABLE = True
except ImportError:
    TEAMS_AI_AVAILABLE = False
    logging.warning("Teams AI library not available. Using custom streaming implementation.")

import uuid
from collections import deque
# Add these imports at the top
from msal import ConfidentialClientApplication
import requests
from urllib.parse import urlparse, parse_qs
# Dictionary to store pending messages for each conversation
pending_messages = {}
# Lock for thread-safe operations on the pending_messages dict
pending_messages_lock = threading.Lock()
# Dictionary for tracking active runs
active_runs = {}
# Active runs lock
active_runs_lock = threading.Lock()
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
# AZURE_ENDPOINT = "https://kb-stellar.openai.azure.com/"  # Replace with your endpoint if different
# AZURE_API_KEY = "bc0ba854d3644d7998a5034af62d03ce"  # Replace with your key if different
# AZURE_API_VERSION = "2024-05-01-preview"
AZURE_ENDPOINT = os.environ.get("OPENAI_ENDPOINT", "")
AZURE_API_KEY = os.environ.get("OPENAI_KEY", "")
AZURE_API_VERSION = os.environ.get("OPENAI_API_VERSION", "")
# App credentials from environment variables for Bot Framework
APP_ID = os.environ.get("MicrosoftAppId", "")
APP_PASSWORD = os.environ.get("MicrosoftAppPassword", "")
TENANT_ID = os.environ.get("TENANT_ID", "")
CLIENT_ID = os.environ.get("MicrosoftAppId", APP_ID)  # Reuse your existing app ID
CLIENT_SECRET = os.environ.get("MicrosoftAppPassword", APP_PASSWORD)  # Reuse your existing password
GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
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

# Define system prompt here instead of relying on external variable
SYSTEM_PROMPT = '''
You are an Email and Chat Assistant that helps manage communications and analyze various file types. Your capabilities vary based on the type of files uploaded.

### Understanding File Types and Processing Methods:
1. **Documents (PDF, DOC, TXT, etc.)** - When users upload these files, you should:
   - Extract relevant information to assist with communications
   - Quote information directly from the documents when answering questions
   - Always reference the specific filename when sharing information from a document

2. **Images** - When users upload images, you should:
   - Refer to the analysis that was automatically added to the conversation
   - Use details from the image analysis to answer questions
   - Acknowledge when information might not be visible in the image

3. **Unsupported File Types**:
   - CSV and Excel files are not supported by this system
   - If users ask about analyzing spreadsheets, kindly inform them that this feature is not available

### Email Assistant Capabilities:
1. **Email Drafting:**
   - Create professional, well-structured emails based on user requirements
   - Adapt tone and formality based on the recipient and purpose
   - Include appropriate greetings, body content, and sign-offs

2. **Email Analysis:**
   - Summarize long email threads or conversations
   - Extract key action items and deadlines from emails
   - Identify important information from uploaded email documents

3. **Response Templates:**
   - Generate templates for common email scenarios (introduction, follow-up, request, etc.)
   - Customize templates based on specific requirements
   - Provide multiple options when appropriate

### Chat Assistant Capabilities:
1. **Message Crafting:**
   - Create clear, concise messages for various chat platforms
   - Adapt style for different communication channels (professional chat, casual messaging, etc.)
   - Suggest appropriate emojis and formatting when relevant

2. **Conversation Management:**
   - Help formulate responses to complex messages
   - Suggest follow-up questions or topics to maintain engagement
   - Provide conversation starters for different situations

3. **Meeting Coordination:**
   - Draft messages to schedule, reschedule, or cancel meetings
   - Create clear agenda items for upcoming discussions
   - Help compose follow-up messages after meetings

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
EMAIL_PROMPT = '''
# First Choice Debt Relief - Email Communication Guidelines

## Core Principles
- BE BRIEF: Keep emails concise and to-the-point: no one reads long emails
- BE DIRECT: Get to the point quickly with clear messages
- BE PROFESSIONAL: Maintain a polished, expert tone
- BE PURPOSEFUL: Every sentence should serve a clear function

## Company Profile
First Choice Debt Relief offers debt resolution services with 17+ years of experience, helping clients reduce their debt burden through negotiated settlements, legal coordination, and financial education.

## Professional Brand Voice
- Confident but not pushy
- Empathetic without being overly emotional
- Clear and straightforward language
- Solution-oriented rather than problem-focused

## Email Structure Requirements
1. Short, descriptive subject lines (5-7 words)
2. Direct greeting with client's name
3. Purpose stated in first sentence/paragraph
4. 2-4 short paragraphs maximum (3-4 sentences each)
5. Bullet points for multiple items (when needed)
6. Clear call-to-action
7. Professional sign-off
8. Include company name in signature

## Email Communication Channels

### 1. CUSTOMER SERVICE CHANNEL
Used for client updates, account maintenance, and responding to existing client inquiries.
**Tone**: Helpful, informative, supportive
**Key characteristics**: Clear explanations, specific next steps, reassurance

### 2. SALES CHANNEL
Used for communicating with prospects and leads about program benefits and enrollment.
**Tone**: Consultative, solution-oriented, value-focused
**Key characteristics**: Clear value propositions, personalized options, gentle urgency

### 3. INTRODUCTION CHANNEL
Used for initial contact with potential clients.
**Tone**: Professional, informative, non-pushy
**Key characteristics**: Brief introduction, clear value proposition, easy next steps

## Email Templates By Category

### CUSTOMER SERVICE TEMPLATES:

#### 1. WELCOME EMAIL
**Purpose**: Sent to new clients upon enrollment to welcome them to the program
**When to use**: Immediately after client enrollment is confirmed
**Key elements**: Program confirmation, what to expect next, contact information

Example:
Subject: Welcome to First Choice Debt Relief!
---
Hi {FIRSTNAME},

Welcome to First Choice Debt Relief! We're excited to have you on board. You've officially been approved and enrolled in our Debt Resolution Program — your journey to financial freedom starts now.

Please take a few moments to review your Program Guide, which includes important details about what to expect, how settlements work, and how to make the most of your program.

If you have any questions, we're just an email or call away.

Sincerely,
The FCDR Team

#### 2. LEGAL UPDATE EMAIL
**Purpose**: Update clients when their account has legal action pending
**When to use**: When a client's account involves legal proceedings
**Key elements**: Status update, explanation of legal provider's role, information about potential settlements

Example:
Subject: Update Regarding Your Legal Account
---
Hi {FIRSTNAME},

I'm reaching out with a quick update on your legal case. Your assigned legal provider is actively working on your behalf, and we're staying in close communication with their office to support the process.

Important: Your legal provider may contact you directly, especially if a potential settlement becomes available. If that happens, please connect with us before making any decisions. We'll help you review the offer based on your available funds and program progress so you can make the most informed decision.

If you're able to contribute additional funds — through a one-time deposit or an increase in your monthly draft — this may help resolve the account faster and give your legal provider more flexibility during negotiations. Just let us know if that's something you'd like to explore.

We're here to support you every step of the way. Feel free to reply to this email or call us at 800-985-9319 with any questions.

Best regards,
First Choice Debt Relief - Client Services

#### 3. LOST SETTLEMENT EMAIL
**Purpose**: Alert clients about a missed settlement payment and its implications
**When to use**: When a client's scheduled payment fails and jeopardizes a settlement
**Key elements**: Clear explanation of the situation, potential consequences, urgent call to action

Example:
Subject: Missed Settlement Payment – Immediate Attention Needed
---
Hi {FIRSTNAME},

We're reaching out regarding a missed payment tied to one of your settlements. This payment was scheduled to be drafted from your {GATEWAY} account, but due to insufficient funds, it could not be processed.

Unfortunately, when a settlement payment is missed, the agreement is typically voided. This means:
- The savings originally negotiated could be lost
- Past payments may be applied to the full balance owed
- The account may revert to the original amount, plus possible interest or fees

At this time, we've paused any future payments to the creditor. However, in some cases, acting quickly may allow us to reinstate the settlement or renegotiate similar terms.

We understand this can be stressful, and we're here to help. Please call us at (714) 589-2245 as soon as possible so we can review your options and help preserve your progress.

We look forward to helping you get back on track.

Sincerely,
First Choice Debt Relief - Client Services

#### 4. LEGAL CONFIRMATION EMAIL
**Purpose**: Confirm receipt of legal documents and explain next steps
**When to use**: When a client has submitted lawsuit documents
**Key elements**: Confirmation of receipt, explanation of legal review process, request for additional funds if needed

Example:
Subject: Lawsuit Document Received – Legal Review in Progress
---
Hi {FIRSTNAME},

We've received the lawsuit related to your enrolled account and have forwarded it to your Legal Plan provider for review. Our office will work closely with your assigned legal representative to help bring this matter to resolution.

With over 17 years of experience resolving cases like this, you can trust that you're in capable hands. You have a highly experienced and dedicated team working on your behalf.

If you're able to deposit additional funds — either as a one-time amount or by increasing your monthly draft — please let us know. This may help expedite the resolution of your account.

Important: Your assigned law office may contact you directly regarding possible settlement offers. If that happens, please speak with our team before making any decisions. We'll help you review your funds and make sure the offer aligns with your program.

If you have any questions, feel free to reply to this email or give us a call at 800-985-9319.

Thank you,
First Choice Debt Relief – Client Support Team
800-985-9319

#### 5. PAYMENT RETURNED EMAIL
**Purpose**: Notify clients about a returned payment and request contact
**When to use**: When a client's payment is returned or rejected
**Key elements**: Payment status notification, request for contact, mention of potential settlement impact

Example:
Subject: Returned Payment – Please Contact Us
---
Hi {FIRSTNAME},

We wanted to let you know that your most recent program payment was returned. When you have a moment, please reach out—even if you're not yet able to reschedule the payment.

Talking with us gives us a chance to go over your options and review any potential program impacts. If you're currently in the middle of a settlement term, it's especially important to stay on track, as a delayed payment could affect your savings agreement.

Our goal is to help you stay on course and succeed in resolving your debt. Please don't hesitate to contact us—we'll work with you to accommodate your needs.

Best regards,
Client Services Team
First Choice Debt Relief
Phone: 800-985-9319
Email: service@firstchoicedebtrelief.com

### SALES TEMPLATES:

#### 1. SALES QUOTE EMAIL
**Purpose**: Provide a formal quote after initial consultation
**When to use**: As a follow-up after speaking with a prospect who showed interest
**Key elements**: Brief recap of conversation, program quote, monthly savings, call to action

Example:
Subject: Your Pre-Approved Debt Relief Quote
---
Hi {FIRSTNAME},

It's been a few days since we last spoke, so I wanted to give you a snapshot of your quote should you still be interested. If you have some questions, let me know and we could hop on a call, and I can also go over the loan option that is offered within the program.

Below you will find your approved quote for the program. As you will see, you could save significantly on a monthly basis. Through this program, your credit effects may have a shorter timeframe than out of the plan because you are working on eliminating your debt quickly, versus years of minimum payments.

Feel free to contact me back by email or phone if you have any further questions or concerns. You can contact me on my direct line at [YOUR_PHONE].

Thank you,
[YOUR_NAME]
First Choice Debt Relief

#### 2. SALES ANALYSIS EMAIL
**Purpose**: Provide a financial analysis to demonstrate program value
**When to use**: When a prospect needs more details about how the program will impact their finances
**Key elements**: Financial analysis, current vs. program comparison, expiring quote mention

Example:
Subject: Your Personal Financial Analysis
---
Hi {FIRSTNAME},

I wanted to provide you with a brief analysis of your current financial situation. Please review so you can see where you stand.

If you have any questions, you can call me at [YOUR_PHONE].

I have included your quote which expires soon.

As you can see, your debts are like an anchor holding you back, not just affecting your credit score and utilization, but your financial well-being. Our solution provides you with monthly relief on your payment, relief from high interest, relief from your credit utilization, and helps you become debt-free YEARS faster compared to just minimum payments.

Thank you,
[YOUR_NAME]
First Choice Debt Relief

#### 3. SALES OVERVIEW EMAIL
**Purpose**: Provide high-level program overview to interested prospects
**When to use**: For prospects who need a concise explanation of program benefits
**Key elements**: Pre-approval notification, program highlights, financial freedom positioning

Example:
Subject: Pre-Approved for Our Debt Resolution Plan
---
Hi {FIRSTNAME},

This is [YOUR_NAME] from First Choice. I have great news, you are pre-approved for our debt resolution plan!

The monthly payment is for an estimated program at an affordable rate. That payment includes everything; the cost of the program and payments to the creditors. There are no pre-payment penalties, you can always pay more, and we'll just get the job done faster.

Our solution provides real financial freedom with a clear end date, unlike minimum payments that can keep you in debt for 15+ years.

I'd be happy to discuss this with you and answer any questions you might have. Feel free to call me at [YOUR_PHONE].

Thank you,
[YOUR_NAME]
First Choice Debt Relief

#### 4. SALES QUICK QUOTE EMAIL
**Purpose**: Provide a quick quote to capture interest of busy prospects
**When to use**: For initial outreach to prospects who might need immediate information
**Key elements**: Low payment option, savings emphasis, time-limited offer, call to action

Example:
Subject: Your Debt Consolidation Quote - Lower Monthly Payment
---
Hi {FIRSTNAME},

This is [YOUR_NAME] from First Choice. We got you a low payment option to consolidate your debt, saving you a significant amount every month compared to what you are paying now.

This quote is valid for a limited time. If you are still serious about consolidating and getting that lower payment, please give me a call at [YOUR_PHONE].

Our goal is to help you get your life back financially!

Thank you,
[YOUR_NAME]
First Choice Debt Relief

#### 5. SALES GENERIC EMAIL
**Purpose**: For custom sales communications that don't fit other templates
**When to use**: When a unique sales situation requires personalized messaging
**Key elements**: Clear value proposition, personal connection, strong call to action

Example:
Subject: [Customized Based on Situation]
---
Hi {FIRSTNAME},

[Customized content addressing specific prospect needs and situation]

Our debt resolution program at First Choice Debt Relief has helped thousands of clients reduce their debt burden and achieve financial freedom faster than they thought possible.

I'd be happy to discuss how we might be able to help with your specific situation. Please feel free to reach out at [YOUR_PHONE] at your convenience.

Thank you,
[YOUR_NAME]
First Choice Debt Relief

### INTRODUCTION TEMPLATES:

#### 1. INTRODUCTION EMAIL
**Purpose**: Make initial contact with potential clients
**When to use**: For first outreach to new leads or referrals
**Key elements**: Brief company introduction, value proposition, invitation to learn more

Example:
Subject: Introduction from First Choice Debt Relief
---
Hi {FIRSTNAME},

My name is [YOUR_NAME] from First Choice Debt Relief. I'm reaching out because we specialize in helping people overcome overwhelming debt and regain financial control.

Based on our initial analysis, we may be able to offer you a program that could significantly reduce your monthly payments and help you become debt-free in a shorter timeframe than making minimum payments.

Would you be interested in learning more about your options? I'd be happy to provide you with a free consultation to discuss your specific situation and how we might be able to help.

Feel free to reach out to me directly at [YOUR_PHONE] or simply reply to this email to schedule a time to chat.

Best regards,
[YOUR_NAME]
First Choice Debt Relief

#### 2. FOLLOW-UP EMAIL
**Purpose**: Follow up after initial contact with no response
**When to use**: When there has been no response to previous outreach
**Key elements**: Reference to previous contact, gentle reminder, invitation to respond

Example:
Subject: Follow-up from First Choice Debt Relief
---
Hi {FIRSTNAME},

I hope this email finds you well. I'm following up on our previous conversation about your debt relief options.

I understand that taking steps to address financial challenges can require careful consideration, and I want to assure you that we're here to help whenever you're ready to move forward.

If you have any questions about our debt resolution program or would like to revisit the details we discussed, please don't hesitate to reach out. I'm available at [YOUR_PHONE] or you can simply reply to this email.

Looking forward to hearing from you.

Best regards,
[YOUR_NAME]
First Choice Debt Relief

#### 3. GENERIC EMAIL
**Purpose**: For communications that don't fit standard templates
**When to use**: When unique situations require custom messaging
**Key elements**: Clear purpose statement, structured information, professional tone

Example:
Subject: [Customized Based on Purpose]
---
Hi {FIRSTNAME},

[Customized content addressing specific situation]

At First Choice Debt Relief, we're committed to helping our clients achieve financial freedom through personalized debt resolution strategies.

[Additional custom content relevant to the specific purpose]

Please don't hesitate to contact us at 800-985-9319 if you have any questions or need further assistance.

Best regards,
[YOUR_NAME]
First Choice Debt Relief

## Email Structure Guidelines

1. **Subject Lines**
   - Clear and specific
   - 40-60 characters in length
   - Include action words for urgent communications
   - Avoid all caps, excessive punctuation, and spam trigger words

2. **Greetings**
   - Use the client's first name when available
   - Default to "Hi [FirstName]," for most communications
   - "Dear [FirstName]," for more formal communications

3. **Body Content**
   - Begin with clear purpose statement
   - Use short paragraphs (3-5 sentences maximum)
   - Include bullet points for multiple items or steps
   - Highlight important information in its own paragraph
   - Use bold formatting sparingly for critical information

4. **Closings**
   - Include clear next steps or call to action
   - Provide relevant contact information
   - Use appropriate sign-off based on context:
     * "Sincerely," for formal communications
     * "Best regards," for standard communications
     * "Thank you," for sales communications
   - Include name, title (if applicable), and company name

5. **Contact Information**
   - Customer Service: 800-985-9319
   - Urgent Account Issues: 714-589-2245
   - Email: service@firstchoicedebtrelief.com

## Special Considerations

1. **When Discussing Settlements**
   - Be clear about potential savings without making guarantees
   - Emphasize the importance of staying on track with payments
   - Explain potential consequences of missed payments

2. **When Discussing Legal Matters**
   - Emphasize partnership with legal providers
   - Advise consulting with First Choice before accepting offers
   - Maintain professional, confident tone without creating unnecessary alarm

3. **When Communicating Urgency**
   - Clearly state timeframes
   - Explain potential consequences of inaction
   - Provide direct contact information
   - Use "Urgent" in subject line only for truly time-sensitive matters

4. **When Replying to Client Queries**
   - Address all questions raised
   - Provide additional relevant information
   - End with invitation for further questions
   - Aim for same-day or next-business-day response

## MOST IMPORTANT INSTRUCTIONS
- OUTPUT THE EMAIL TEXT ONLY - NO COMMENTARY OR EXPLANATIONS BEFORE OR AFTER
- DO NOT include phrases like "Here's the email" or "Let me know if you need changes"
- DO NOT write about what you're doing or explain your process
- DO NOT ask follow-up questions after the email
- SIMPLY WRITE THE EMAIL CONTENT EXACTLY AS IT SHOULD APPEAR, STARTING WITH THE GREETING
- DO NOT ASK QUESTIONS after the email or explain what you've done
- ONLY OUTPUT THE EXACT EMAIL CONTENT as if it were being sent directly
- THE FIRST LINE OF YOUR RESPONSE MUST BE THE EMAIL GREETING (e.g., "Hi [Name],")
- END WITH THE SIGNATURE - NOTHING AFTER THAT


'''

def get_sharepoint_access_token():
    """Gets an access token for Microsoft Graph API to access SharePoint/OneDrive files."""
    try:
        app = ConfidentialClientApplication(
            client_id=CLIENT_ID,
            client_credential=CLIENT_SECRET,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}"
        )
        
        result = app.acquire_token_silent(scopes=GRAPH_SCOPE, account=None)
        
        if not result:
            logging.info("No suitable token in cache, getting new token")
            result = app.acquire_token_for_client(scopes=GRAPH_SCOPE)
        
        if "access_token" in result:
            logging.info("Access token obtained successfully")
            return result["access_token"]
        else:
            logging.error(f"Error getting token: {result.get('error')}")
            logging.error(f"Error description: {result.get('error_description')}")
            return None
    except Exception as e:
        logging.error(f"Exception in get_sharepoint_access_token: {str(e)}")
        return None

async def download_sharepoint_file(file_url, filename):
    """
    Downloads a file from SharePoint using Microsoft Graph API.
    
    Args:
        file_url: The URL to the file
        filename: The filename to save as
        
    Returns:
        str: Local file path if successful, None otherwise
    """
    try:
        # Parse the URL to extract important components
        parsed_url = urlparse(file_url)
        
        # Get access token
        access_token = get_sharepoint_access_token()
        if not access_token:
            logging.error("Failed to get access token for SharePoint file")
            return None
        
        # Extract drive ID and item ID from URL if present
        # This is a simplified approach, actual extraction might depend on your specific URL format
        path_parts = parsed_url.path.split('/')
        
        # If this is a sharing URL, we need to resolve it first
        if "sharepointonline.com" in parsed_url.netloc and "/s/" in parsed_url.path:
            # This is a sharing link, we need to resolve it
            sharing_url = file_url
            headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            }
            
            # Use Graph API to resolve the sharing link
            resolve_url = "https://graph.microsoft.com/v1.0/shares/u!{encoded_url}/driveItem"
            # URL needs to be encoded properly for the API
            encoded_url = sharing_url.replace(':', '%3A').replace('/', '%2F')
            
            resolve_response = requests.get(
                resolve_url.format(encoded_url=encoded_url),
                headers=headers
            )
            
            if resolve_response.status_code != 200:
                logging.error(f"Failed to resolve sharing URL: {resolve_response.status_code} - {resolve_response.text}")
                return None
                
            item_info = resolve_response.json()
            # Now we can download using the item's download URL
            download_url = item_info.get('@microsoft.graph.downloadUrl')
            
            if not download_url:
                logging.error("Could not get download URL from sharing link resolution")
                return None
                
            # Download the file
            download_response = requests.get(download_url)
            if download_response.status_code != 200:
                logging.error(f"Failed to download file: {download_response.status_code}")
                return None
                
            # Save the file locally
            file_path = os.path.join(FILE_DIRECTORY, filename)
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            
            with open(file_path, 'wb') as f:
                f.write(download_response.content)
                
            logging.info(f"Successfully downloaded SharePoint file to {file_path}")
            return file_path
            
        else:
            # Handle other SharePoint/OneDrive URL formats
            logging.error(f"Unrecognized SharePoint URL format: {file_url}")
            return None
            
    except Exception as e:
        logging.error(f"Error downloading SharePoint file: {str(e)}")
        traceback.print_exc()
        return None
def create_typing_stop_activity():
    """Creates an activity to explicitly stop the typing indicator"""
    return Activity(
        type=ActivityTypes.message,
        text="",  # Empty text message to replace typing indicator
        value={"action": "stop_typing"}  # Metadata for debugging
    )
# Custom TeamsStreamingResponse for better control when official library not available
class TeamsStreamingResponse:
    """Handles streaming responses to Teams in a more controlled way"""
    
    def __init__(self, turn_context):
        self.turn_context = turn_context
        self.message_parts = []
        self.is_first_update = True
        self.stream_id = f"stream_{int(time.time())}"
        self.last_update_time = 0
        self.min_update_interval = 1.5  # Minimum time between updates in seconds (Teams requirement)
        
    async def send_typing_indicator(self):
        """Sends a typing indicator to Teams"""
        await self.turn_context.send_activity(create_typing_activity())
    
    async def queue_update(self, text_chunk):
        """Queues and potentially sends a text update"""
        # Add to the accumulated text
        self.message_parts.append(text_chunk)
        
        # Check if we should send an update
        current_time = time.time()
        if (current_time - self.last_update_time) >= self.min_update_interval:
            await self.send_typing_indicator()
            self.last_update_time = current_time
    
    def get_full_message(self):
        """Gets the complete message from all chunks"""
        return "".join(self.message_parts)
    
    async def send_final_message(self):
        """Sends the final complete message, split if necessary"""
        complete_message = self.get_full_message()
        
        # Split long messages if needed (Teams has message size limits)
        if len(complete_message) > 7000:
            chunks = [complete_message[i:i+7000] for i in range(0, len(complete_message), 7000)]
            for i, chunk in enumerate(chunks):
                if i == 0:
                    await self.turn_context.send_activity(chunk)
                else:
                    await self.turn_context.send_activity(f"(continued) {chunk}")
        else:
            await self.turn_context.send_activity(complete_message)
        await self.turn_context.send_activity(create_typing_stop_activity())
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
        recovery_message = "Creating a fresh session while keeping our context."
        await turn_context.send_activity(recovery_message)
        
        # Create completely new resources
        try:
            # Create a new vector store
            vector_store = client.vector_stores.create(
                name=f"recovery_user_{user_id}_convo_{conversation_id}_{int(time.time())}"
            )
            
            # Create a new assistant with a unique name
            unique_name = f"recovery_assistant_user_{user_id}_{int(time.time())}"
            assistant_obj = client.beta.assistants.create(
                name=unique_name,
                model="gpt-4.1-mini",
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
            
            # Clear any active runs (thread safe)
            with active_runs_lock:
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

        # FALLBACK: Use direct completion API if everything else fails
        try:
            await send_fallback_response(turn_context, "I'm having trouble with our conversation system. Let me try to help directly. What can I assist you with?")
        except Exception as fallback_error:
            logging.error(f"Even fallback failed for user {user_id}: {fallback_error}")
def create_channel_selection_card():
    """Creates an adaptive card for selecting email channels (Sales or Customer Service)"""
    card = {
        "type": "AdaptiveCard",
        "version": "1.0",
        "body": [
            {
                "type": "TextBlock",
                "text": "First Choice Debt Relief Email Templates",
                "size": "large",
                "weight": "bolder"
            },
            {
                "type": "TextBlock",
                "text": "Please select an email channel:",
                "wrap": True
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Introduction Email",
                "data": {
                    "action": "select_channel",
                    "channel": "intro"
                }
            },
            {
                "type": "Action.Submit",
                "title": "Sales Email",
                "data": {
                    "action": "select_channel",
                    "channel": "sales"
                }
            },
            {
                "type": "Action.Submit",
                "title": "Customer Service Email",
                "data": {
                    "action": "select_channel",
                    "channel": "customer_service"
                }
            }
        ]
    }
    
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    return attachment
async def send_fallback_response(turn_context: TurnContext, user_message: str):
    """Last resort fallback using direct completion API"""
    try:
        client = create_client()
        
        # Send a typing indicator first
        await turn_context.send_activity(create_typing_activity())
        
        # Get user's message if not provided
        if not user_message:
            if hasattr(turn_context.activity, 'text'):
                user_message = turn_context.activity.text.strip()
            else:
                user_message = "Hello, I need your help."
        
        # Create a simple completion request with minimal context
        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            messages=[
                {"role": "system", "content": "You are a helpful product management assistant. Keep your response concise and helpful."},
                {"role": "user", "content": user_message}
            ],
            max_tokens=1000
        )
        
        # Send the response back
        if response.choices and response.choices[0].message.content:
            await turn_context.send_activity(response.choices[0].message.content)
        else:
            await turn_context.send_activity("I'm sorry, I'm having trouble generating a response right now. Please try again later.")
    
    except Exception as e:
        logging.error(f"Fallback response generation failed: {e}")
        await turn_context.send_activity("I'm experiencing technical difficulties right now. Please try again in a moment.")
def create_welcome_card():
    """Creates an enhanced welcome card with modern features"""
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "style": "emphasis",
                "bleed": True,
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "Image",
                                        "url": "https://adaptivecards.io/content/email.png",
                                        "size": "Small",
                                        "altText": "Email assistant icon"
                                    }
                                ],
                                "verticalContentAlignment": "Center"
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Email & Chat Assistant",
                                        "wrap": True,
                                        "size": "Large",
                                        "weight": "Bolder",
                                        "color": "Accent"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": "Your AI-powered communication partner",
                                        "wrap": True,
                                        "isSubtle": True
                                    }
                                ],
                                "verticalContentAlignment": "Center"
                            }
                        ]
                    }
                ]
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "👋 Welcome! I'm here to help with your communication needs.",
                        "wrap": True,
                        "size": "Medium",
                        "weight": "Bolder",
                        "spacing": "Medium"
                    },
                    {
                        "type": "TextBlock",
                        "text": "I can help you with:",
                        "wrap": True,
                        "spacing": "Medium"
                    },
                    {
                        "type": "FactSet",
                        "facts": [
                            {
                                "title": "📧",
                                "value": "Drafting professional emails"
                            },
                            {
                                "title": "📄",
                                "value": "Analyzing documents (PDF, DOC, TXT)"
                            },
                            {
                                "title": "🖼️",
                                "value": "Describing and analyzing images"
                            },
                            {
                                "title": "💬",
                                "value": "Answering questions and providing assistance"
                            }
                        ]
                    },
                    {
                        "type": "TextBlock",
                        "text": "⚠️ Note: CSV and Excel files are not supported",
                        "wrap": True,
                        "color": "Attention",
                        "isSubtle": True,
                        "spacing": "Small"
                    }
                ]
            },
            {
                "type": "Container",
                "style": "good",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Get Started",
                        "wrap": True,
                        "size": "Medium",
                        "weight": "Bolder"
                    },
                    {
                        "type": "TextBlock",
                        "text": "Select an option below or simply type a message to begin.",
                        "wrap": True
                    }
                ],
                "spacing": "Medium"
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "✉️ Create Email Template",
                "style": "positive",
                "data": {
                    "action": "create_email"
                }
            },
            {
                "type": "Action.Submit",
                "title": "📁 Upload a Document",
                "style": "default",
                "data": {
                    "action": "show_upload_info"
                }
            },
            {
                "type": "Action.ShowCard",
                "title": "❓ Help & Tips",
                "card": {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Quick Tips:",
                            "wrap": True,
                            "weight": "Bolder"
                        },
                        {
                            "type": "TextBlock",
                            "text": "• Type '/email' to create an email template anytime\n• Upload files using the paperclip button in Teams\n• Ask specific questions about uploaded documents\n• For best results, be clear and detailed in your requests",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Sample queries:",
                            "wrap": True,
                            "weight": "Bolder",
                            "spacing": "Medium"
                        },
                        {
                            "type": "TextBlock",
                            "text": "\"Draft a follow-up email to the marketing team\"\n\"Summarize the key points in this document\"\n\"Help me write a meeting invitation for Friday\"",
                            "wrap": True
                        }
                    ]
                }
            }
        ]
    }
    
    return CardFactory.adaptive_card(card)

async def send_welcome_message(turn_context: TurnContext):
    """Sends enhanced welcome message with modern adaptive card"""
    welcome_card = create_welcome_card()
    
    reply = _create_reply(turn_context.activity)
    reply.attachments = [welcome_card]
    await turn_context.send_activity(reply)

def create_edit_email_card(original_email):
    """
    Creates an adaptive card for email editing.
    
    Args:
        original_email: The original email text to edit
    
    Returns:
        Attachment: The card attachment
    """
    card = {
        "type": "AdaptiveCard",
        "version": "1.0",
        "body": [
            {
                "type": "TextBlock",
                "text": "Edit Email",
                "size": "large",
                "weight": "bolder"
            },
            {
                "type": "TextBlock",
                "text": "Current Email:",
                "wrap": True
            },
            {
                "type": "TextBlock",
                "text": original_email,
                "wrap": True,
                "spacing": "Small"
            },
            {
                "type": "TextBlock",
                "text": "What changes would you like to make?",
                "wrap": True
            },
            {
                "type": "Input.Text",
                "id": "edit_instructions",
                "placeholder": "E.g., 'Make it more concise', 'Add more details about payment options', 'Change the tone to be more urgent'",
                "isMultiline": True
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Apply Changes",
                "data": {
                    "action": "apply_email_edits"
                }
            },
            {
                "type": "Action.Submit",
                "title": "Cancel",
                "data": {
                    "action": "cancel_edit"
                }
            }
        ]
    }
    
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    return attachment
async def send_edit_email_card(turn_context: TurnContext, state):
    """
    Sends an email editing card to the user.
    
    Args:
        turn_context: The turn context
        state: The conversation state containing the last generated email
    """
    with conversation_states_lock:
        original_email = state.get("last_generated_email", "")
    
    if not original_email:
        await turn_context.send_activity("I couldn't find a recently generated email to edit. Please create a new email first.")
        return
    
    reply = _create_reply(turn_context.activity)
    reply.attachments = [create_edit_email_card(original_email)]
    await turn_context.send_activity(reply)
async def apply_email_edits(turn_context: TurnContext, state, edit_instructions):
    """
    Applies edits to the previously generated email.
    
    Args:
        turn_context: The turn context
        state: The conversation state
        edit_instructions: Instructions for editing the email
    """
    # Send typing indicator
    await turn_context.send_activity(create_typing_activity())
    
    # Get the original email and template data
    with conversation_states_lock:
        original_email = state.get("last_generated_email", "")
        template_id = state.get("last_email_template", "generic")
        email_data = state.get("last_email_data", {})
    
    if not original_email:
        await turn_context.send_activity("I couldn't find the original email to edit. Please create a new email.")
        return
    
    # Create prompt for editing
    prompt = f"Edit the following email based on these instructions: {edit_instructions}\n\n"
    prompt += "ORIGINAL EMAIL:\n"
    prompt += f"{original_email}\n\n"
    prompt += "Please provide the complete revised email with all changes incorporated."
    
    # Initialize chat if needed
    if not state.get("assistant_id"):
        await initialize_chat(turn_context, state)
    
    # Use the existing process_conversation_internal function to get AI response
    client = create_client()
    result = await process_conversation_internal(
        client=client,
        session=state["session_id"],
        prompt=prompt,
        assistant=state["assistant_id"],
        stream_output=False
    )
    
    # Extract and format the edited email
    if isinstance(result, dict) and "response" in result:
        edited_email = result["response"]
        
        # Update the saved email
        with conversation_states_lock:
            state["last_generated_email"] = edited_email
        
        # Create an email result card
        email_card = {
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "Edited Email",
                    "size": "large",
                    "weight": "bolder"
                },
                {
                    "type": "TextBlock",
                    "text": edited_email,
                    "wrap": True
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Edit Again",
                    "data": {
                        "action": "edit_email"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Create Another Email",
                    "data": {
                        "action": "create_email"
                    }
                }
            ]
        }
        
        # Create attachment
        attachment = Attachment(
            content_type="application/vnd.microsoft.card.adaptive",
            content=email_card
        )
        
        reply = _create_reply(turn_context.activity)
        reply.attachments = [attachment]
        await turn_context.send_activity(reply)
    else:
        await turn_context.send_activity("I'm sorry, I couldn't edit the email. Please try again with different instructions.")
# Add this to your handle_card_actions function
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
                    if conversation_id in pending_messages:
                        pending_messages[conversation_id].clear()
                
                # Send typing indicator
                await turn_context.send_activity(create_typing_activity())
                
                # Initialize new chat
                await initialize_chat(turn_context, None)  # Pass None to force new state creation
            else:
                await initialize_chat(turn_context, None)
        elif action_data.get("action") == "generate_email":
            # Get conversation state
            conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
            conversation_id = conversation_reference.conversation.id
            state = conversation_states[conversation_id]
            
            # Get template type
            template_id = action_data.get("template", "generic")
            
            # Extract common fields
            recipient = action_data.get("recipient", "")
            instructions = action_data.get("instructions", "")
            chain = action_data.get("chain", "")
            has_attachments = action_data.get("hasAttachments", "false") == "true"
            
            # Extract template-specific fields
            firstname = action_data.get("firstname", "")
            gateway = action_data.get("gateway", "")
            subject = action_data.get("subject", "")
            
            # Generate email using AI
            await generate_email(
                turn_context, 
                state, 
                template_id, 
                recipient, 
                firstname, 
                gateway, 
                subject, 
                instructions, 
                chain, 
                has_attachments
            )
        elif action_data.get("action") == "create_email":
            # Send channel selection card
            await send_email_card(turn_context, "channel_selection")
        elif action_data.get("action") == "select_channel":
            # Get the selected channel
            channel = action_data.get("channel", "intro")
            
            # Send the template selection card for this channel
            await send_email_card(turn_context, "selection", channel)
        elif action_data.get("action") == "select_template":
            # Get the selected template
            template = action_data.get("template", "generic")
            
            # Send the appropriate template card
            await send_email_card(turn_context, template)
        elif action_data.get("action") == "edit_email":
            # Get conversation state
            conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
            conversation_id = conversation_reference.conversation.id
            state = conversation_states[conversation_id]
            
            # Send edit email card
            await send_edit_email_card(turn_context, state)
        elif action_data.get("action") == "apply_email_edits":
            # Get conversation state
            conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
            conversation_id = conversation_reference.conversation.id
            state = conversation_states[conversation_id]
            
            # Get edit instructions
            edit_instructions = action_data.get("edit_instructions", "")
            
            # Apply edits
            await apply_email_edits(turn_context, state, edit_instructions)
        # Add this to the handle_card_actions function
        elif action_data.get("action") == "regenerate_email":
            # Get conversation state
            conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
            conversation_id = conversation_reference.conversation.id
            state = conversation_states[conversation_id]
            
            # Extract parameters
            template_id = action_data.get("template", "generic")
            recipient = action_data.get("recipient", "")
            firstname = action_data.get("firstname", "")
            gateway = action_data.get("gateway", "")
            subject = action_data.get("subject", "")
            instructions = action_data.get("instructions", "")
            chain = action_data.get("chain", "")
            has_attachments = action_data.get("hasAttachments", "false") == "true"
            
            # Get previous instructions and add a request for more creativity/variation
            with conversation_states_lock:
                previous_instructions = state.get("last_email_data", {}).get("instructions", "")
            
            # Combine previous instructions with request for variation
            if previous_instructions:
                new_instructions = f"{previous_instructions}\n\nPlease provide a different variation of this email with alternative wording while keeping the same message intent."
            else:
                new_instructions = "Please provide a different variation of this email with alternative wording while keeping the same message intent."
            
            # Generate a new version
            await generate_email(
                turn_context, 
                state, 
                template_id, 
                recipient, 
                firstname, 
                gateway, 
                subject, 
                new_instructions, 
                chain, 
                has_attachments
            )
        elif action_data.get("action") == "cancel_edit":
            # Cancel edit and go back to last generated email
            conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
            conversation_id = conversation_reference.conversation.id
            state = conversation_states[conversation_id]
            
            with conversation_states_lock:
                original_email = state.get("last_generated_email", "")
            
            if original_email:
                # Create an email result card
                email_card = {
                    "type": "AdaptiveCard",
                    "version": "1.0",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Generated Email",
                            "size": "large",
                            "weight": "bolder"
                        },
                        {
                            "type": "TextBlock",
                            "text": original_email,
                            "wrap": True
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.Submit",
                            "title": "Edit This Email",
                            "data": {
                                "action": "edit_email"
                            }
                        },
                        {
                            "type": "Action.Submit",
                            "title": "Create Another Email",
                            "data": {
                                "action": "create_email"
                            }
                        }
                    ]
                }
                
                attachment = Attachment(
                    content_type="application/vnd.microsoft.card.adaptive",
                    content=email_card
                )
                
                reply = _create_reply(turn_context.activity)
                reply.attachments = [attachment]
                await turn_context.send_activity(reply)
            else:
                await send_email_card(turn_context)
    except Exception as e:
        logging.error(f"Error handling card action: {e}")
        await turn_context.send_activity(f"I couldn't process your request. Please try again later.")
def get_template_title(template_id):
    """
    Returns the human-readable title for a template ID.
    
    Args:
        template_id (str): Template identifier
    
    Returns:
        str: Human-readable template title
    """
    template_titles = {
        # Customer service templates
        "welcome": "Welcome Email",
        "legal_update": "Legal Document Update",
        "lost_settlement": "Lost Settlement",
        "legal_confirmation": "Legal Document Confirmation",
        "payment_returned": "Payment Returned",
        
        # Sales templates
        "sales_quote": "Initial Quote Email",
        "sales_analysis": "Financial Analysis Email",
        "sales_overview": "Program Overview Email",
        "sales_generic": "Generic Sales Email",
        "sales_quick_quote": "Quick Quote Email",  # New template
        
        # Intro templates
        "introduction": "Introduction Email",
        "followup": "Follow-up Email",
        "generic": "Generic Email"
    }
    
    return template_titles.get(template_id, "Email Template")
def get_template_content(template_id, **kwargs):
    """
    Returns the base content for a specific template with placeholders.
    
    Args:
        template_id (str): Template identifier
        **kwargs: Key-value pairs for template placeholders
    
    Returns:
        tuple: (subject, content) tuple with template content
    """
    # Default placeholder values
    firstname = kwargs.get('firstname', '{FIRSTNAME}')
    gateway = kwargs.get('gateway', '{GATEWAY}')
    
    # Customer service templates
    templates = {
        "welcome": (
            "Welcome to First Choice Debt Relief!",
            f"Hi {firstname},\n\n"
            "Welcome to First Choice Debt Relief! We're excited to have you on board. "
            "You've officially been approved and enrolled in our Debt Resolution Program — "
            "your journey to financial freedom starts now.\n\n"
            "Please take a few moments to review your Program Guide, which includes important "
            "details about what to expect, how settlements work, and how to make the most of your program.\n\n"
            "If you have any questions, we're just an email or call away.\n\n"
            "Sincerely,\n"
            "The FCDR Team"
        ),
        "legal_update": (
            "Update Regarding Your Legal Account",
            f"Hi {firstname},\n\n"
            "I'm reaching out with a quick update on your legal case. Your assigned legal provider "
            "is actively working on your behalf, and we're staying in close communication with their "
            "office to support the process.\n\n"
            "Important: Your legal provider may contact you directly, especially if a potential settlement "
            "becomes available. If that happens, please connect with us before making any decisions. "
            "We'll help you review the offer based on your available funds and program progress so "
            "you can make the most informed decision.\n\n"
            "If you're able to contribute additional funds — through a one-time deposit or an increase "
            "in your monthly draft — this may help resolve the account faster and give your legal provider "
            "more flexibility during negotiations. Just let us know if that's something you'd like to explore.\n\n"
            "We're here to support you every step of the way. Feel free to reply to this email or "
            "call us at 800-985-9319 with any questions.\n\n"
            "Best regards,\n"
            "First Choice Debt Relief - Client Services"
        ),
        "lost_settlement": (
            "Missed Settlement Payment – Immediate Attention Needed",
            f"Hi {firstname},\n\n"
            f"We're reaching out regarding a missed payment tied to one of your settlements. "
            f"This payment was scheduled to be drafted from your {gateway} account, but due to "
            f"insufficient funds, it could not be processed.\n\n"
            f"Unfortunately, when a settlement payment is missed, the agreement is typically voided. This means:\n"
            f"- The savings originally negotiated could be lost\n"
            f"- Past payments may be applied to the full balance owed\n"
            f"- The account may revert to the original amount, plus possible interest or fees\n\n"
            f"At this time, we've paused any future payments to the creditor. However, in some cases, "
            f"acting quickly may allow us to reinstate the settlement or renegotiate similar terms.\n\n"
            f"We understand this can be stressful, and we're here to help. Please call us at (714) 589-2245 "
            f"as soon as possible so we can review your options and help preserve your progress.\n\n"
            f"We look forward to helping you get back on track.\n\n"
            f"Sincerely,\n"
            f"First Choice Debt Relief - Client Services"
        ),
        "legal_confirmation": (
            "Lawsuit Document Received – Legal Review in Progress",
            f"Hi {firstname},\n\n"
            f"We've received the lawsuit related to your enrolled account and have forwarded it to your "
            f"Legal Plan provider for review. Our office will work closely with your assigned legal "
            f"representative to help bring this matter to resolution.\n\n"
            f"With over 17 years of experience resolving cases like this, you can trust that you're in capable hands. "
            f"You have a highly experienced and dedicated team working on your behalf.\n\n"
            f"If you're able to deposit additional funds — either as a one-time amount or by increasing your "
            f"monthly draft — please let us know. This may help expedite the resolution of your account.\n\n"
            f"Important: Your assigned law office may contact you directly regarding possible settlement offers. "
            f"If that happens, please speak with our team before making any decisions. We'll help you review your "
            f"funds and make sure the offer aligns with your program.\n\n"
            f"If you have any questions, feel free to reply to this email or give us a call at 800-985-9319.\n\n"
            f"Thank you,\n"
            f"First Choice Debt Relief – Client Support Team\n"
            f"800-985-9319"
        ),
        "payment_returned": (
            "Returned Payment – Please Contact Us",
            f"Hi {firstname},\n\n"
            f"We wanted to let you know that your most recent program payment was returned. "
            f"When you have a moment, please reach out—even if you're not yet able to reschedule the payment.\n\n"
            f"Talking with us gives us a chance to go over your options and review any potential program impacts. "
            f"If you're currently in the middle of a settlement term, it's especially important to stay on track, "
            f"as a delayed payment could affect your savings agreement.\n\n"
            f"Our goal is to help you stay on course and succeed in resolving your debt. "
            f"Please don't hesitate to contact us—we'll work with you to accommodate your needs.\n\n"
            f"Best regards,\n"
            f"Client Services Team\n"
            f"First Choice Debt Relief\n"
            f"Phone: 800-985-9319\n"
            f"Email: service@firstchoicedebtrelief.com"
        ),
        # Sales templates
        "sales_quote": (
            "Your Pre-Approved Debt Relief Quote",
            f"Hi {firstname},\n\n"
            f"It's been a few days since we last spoke, so I wanted to give you a snapshot of your quote should you still be interested. "
            f"If you have some questions, let me know and we could hop on a call, and I can also go over the loan option that is offered within the program.\n\n"
            f"Below you will find your approved quote for the program. As you will see, you could save significantly on a monthly basis. "
            f"Through this program, your credit effects may have a shorter timeframe than out of the plan because you are working on eliminating your debt quickly, "
            f"versus years of minimum payments.\n\n"
            f"Feel free to contact me back by email or phone if you have any further questions or concerns. "
            f"You can contact me on my direct line at [YOUR_PHONE].\n\n"
            f"Thank you,\n"
            f"[YOUR_NAME]\n"
            f"First Choice Debt Relief\n"
        ),
        "sales_analysis": (
            "Your Personal Financial Analysis",
            f"Hi {firstname},\n\n"
            f"I wanted to provide you with a brief analysis of your current financial situation. Please review so you can see where you stand.\n\n"
            f"If you have any questions, you can call me at [YOUR_PHONE].\n\n"
            f"I have included your quote which expires soon.\n\n"
            f"As you can see, your debts are like an anchor holding you back, not just affecting your credit score and utilization, but your financial well-being. "
            f"Our solution provides you with monthly relief on your payment, relief from high interest, relief from your credit utilization, "
            f"and helps you become debt-free YEARS faster compared to just minimum payments.\n\n"
            f"Thank you,\n"
            f"[YOUR_NAME]\n"
            f"First Choice Debt Relief\n"
        ),
        "sales_overview": (
            "Pre-Approved for Our Debt Resolution Plan",
            f"Hi {firstname},\n\n"
            f"This is [YOUR_NAME] from First Choice. I have great news, you are pre-approved for our debt resolution plan!\n\n"
            f"The monthly payment is for an estimated program at an affordable rate. That payment includes everything; "
            f"the cost of the program and payments to the creditors. There are no pre-payment penalties, you can always pay more, "
            f"and we'll just get the job done faster.\n\n"
            f"Our solution provides real financial freedom with a clear end date, unlike minimum payments that can keep you in debt for 15+ years.\n\n"
            f"I'd be happy to discuss this with you and answer any questions you might have. Feel free to call me at [YOUR_PHONE].\n\n"
            f"Thank you,\n"
            f"[YOUR_NAME]\n"
            f"First Choice Debt Relief\n"
        ),
        "sales_quick_quote": (
            "Your Debt Consolidation Quote - Lower Monthly Payment",
            f"Hi {firstname},\n\n"
            f"This is [YOUR_NAME] from First Choice. We got you a low payment option to consolidate your debt, "
            f"saving you a significant amount every month compared to what you are paying now.\n\n"
            f"This quote is valid for a limited time. If you are still serious about consolidating and getting that lower payment, "
            f"please give me a call at [YOUR_PHONE].\n\n"
            f"Our goal is to help you get your life back financially!\n\n"
            f"Thank you,\n"
            f"[YOUR_NAME]\n"
            f"First Choice Debt Relief\n"
        ),
        # Intro templates
        "introduction": (
            "Introduction from First Choice Debt Relief",
            f"Hi {firstname},\n\n"
            f"My name is [YOUR_NAME] from First Choice Debt Relief. I'm reaching out because we specialize in helping people overcome overwhelming debt "
            f"and regain financial control.\n\n"
            f"Based on our initial analysis, we may be able to offer you a program that could significantly reduce your monthly payments and "
            f"help you become debt-free in a shorter timeframe than making minimum payments.\n\n"
            f"Would you be interested in learning more about your options? I'd be happy to provide you with a free consultation "
            f"to discuss your specific situation and how we might be able to help.\n\n"
            f"Feel free to reach out to me directly at [YOUR_PHONE] or simply reply to this email to schedule a time to chat.\n\n"
            f"Best regards,\n"
            f"[YOUR_NAME]\n"
            f"First Choice Debt Relief\n"
        ),
        "followup": (
            "Follow-up from First Choice Debt Relief",
            f"Hi {firstname},\n\n"
            f"I hope this email finds you well. I'm following up on our previous conversation about your debt relief options.\n\n"
            f"I understand that taking steps to address financial challenges can require careful consideration, "
            f"and I want to assure you that we're here to help whenever you're ready to move forward.\n\n"
            f"If you have any questions about our debt resolution program or would like to revisit the details we discussed, "
            f"please don't hesitate to reach out. I'm available at [YOUR_PHONE] or you can simply reply to this email.\n\n"
            f"Looking forward to hearing from you.\n\n"
            f"Best regards,\n"
            f"[YOUR_NAME]\n"
            f"First Choice Debt Relief\n"
        )
    }
    
    # Default to empty template if not found
    return templates.get(template_id, ("", ""))
def get_template_channel(template_id):
    """
    Returns the channel for a given template ID.
    
    Args:
        template_id (str): Template identifier
    
    Returns:
        str: Channel name
    """
    # Sales templates
    if template_id.startswith("sales_"):
        return "sales"
    # Customer service templates
    elif template_id in ["welcome", "legal_update", "lost_settlement", "legal_confirmation", "payment_returned"]:
        return "customer_service"
    # Introduction templates
    elif template_id in ["introduction", "followup", "generic"]:
        return "intro"
    # Default
    else:
        return "intro"
def create_email_card(template_mode="selection", channel=None):
    """
    Creates an adaptive card for email composition with template selection.
    
    Args:
        template_mode (str): Mode of the card - "selection", "generic", or specific template name
        channel (str): Email channel - "sales" or "customer_service"
    """
    if template_mode == "selection":
        # Template selection card based on channel
        card = {
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                {
                    "type": "TextBlock",
                    "text": f"First Choice Debt Relief {channel.replace('_', ' ').title() if channel else ''} Email Templates",
                    "size": "large",
                    "weight": "bolder"
                },
                {
                    "type": "TextBlock",
                    "text": "Please select an email template:",
                    "wrap": True
                }
            ],
            "actions": []
        }
        
        # Add actions based on channel
        if channel == "sales":
            card["actions"] = [
                {
                    "type": "Action.Submit",
                    "title": "Quick Quote Email",
                    "data": {
                        "action": "select_template",
                        "template": "sales_quick_quote"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Initial Quote Email",
                    "data": {
                        "action": "select_template",
                        "template": "sales_quote"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Financial Analysis Email",
                    "data": {
                        "action": "select_template",
                        "template": "sales_analysis"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Program Overview Email",
                    "data": {
                        "action": "select_template",
                        "template": "sales_overview"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Generic Sales Email",
                    "data": {
                        "action": "select_template",
                        "template": "sales_generic"
                    }
                }
            ]
        elif channel == "customer_service":
            card["actions"] = [
                {
                    "type": "Action.Submit",
                    "title": "Welcome Email",
                    "data": {
                        "action": "select_template",
                        "template": "welcome"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Legal Document Update",
                    "data": {
                        "action": "select_template",
                        "template": "legal_update"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Lost Settlement",
                    "data": {
                        "action": "select_template",
                        "template": "lost_settlement"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Legal Document Confirmation",
                    "data": {
                        "action": "select_template",
                        "template": "legal_confirmation"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Payment Returned",
                    "data": {
                        "action": "select_template",
                        "template": "payment_returned"
                    }
                }
            ]
        elif channel == "intro":
            card["actions"] = [
                {
                    "type": "Action.Submit",
                    "title": "Introduction Email",
                    "data": {
                        "action": "select_template",
                        "template": "introduction"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Follow-up Email",
                    "data": {
                        "action": "select_template",
                        "template": "followup"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Generic Email",
                    "data": {
                        "action": "select_template",
                        "template": "generic"
                    }
                }
            ]
        
        # Add back button
        card["actions"].append({
            "type": "Action.Submit",
            "title": "Back to Channels",
            "data": {
                "action": "create_email"
            }
        })
            
    elif template_mode == "generic" or template_mode == "sales_generic":
        # Generic email card
        card = {
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "First Choice Debt Relief - Email Creator",
                    "size": "large",
                    "weight": "bolder"
                },
                {
                    "type": "TextBlock",
                    "text": "Recipient (Optional)",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "recipient",
                    "placeholder": "Enter recipient(s)"
                },
                {
                    "type": "TextBlock",
                    "text": "Subject",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "subject",
                    "placeholder": "Enter email subject"
                },
                {
                    "type": "TextBlock",
                    "text": "Instructions",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "instructions",
                    "placeholder": "Describe what you want in this email, including any specific points to include or avoid",
                    "isMultiline": True
                },
                {
                    "type": "TextBlock",
                    "text": "Previous Email (for replies)",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "chain",
                    "placeholder": "Paste previous email if this is a reply",
                    "isMultiline": True
                },
                {
                    "type": "Input.Toggle",
                    "id": "hasAttachments",
                    "title": "Mention attachments in email?",
                    "value": "false"
                },
                {
                    "type": "TextBlock",
                    "text": "Note: This only mentions attachments in the text. To actually attach files, you'll need to add them when sending the email in your email client.",
                    "wrap": True,
                    "isSubtle": True,
                    "size": "small"
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Generate Email",
                    "data": {
                        "action": "generate_email",
                        "template": template_mode
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Back to Templates",
                    "data": {
                        "action": "select_channel",
                        "channel": "sales" if template_mode == "sales_generic" else "intro"
                    }
                }
            ]
        }
    else:
        # Template-specific card
        card = {
            "type": "AdaptiveCard",
            "version": "1.0",
            "body": [
                {
                    "type": "TextBlock",
                    "text": f"First Choice Debt Relief - {get_template_title(template_mode)}",
                    "size": "large",
                    "weight": "bolder"
                },
                {
                    "type": "TextBlock",
                    "text": "Recipient (Optional)",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "recipient",
                    "placeholder": "Enter recipient(s)"
                },
                {
                    "type": "TextBlock",
                    "text": "Client First Name",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "firstname",
                    "placeholder": "Enter client's first name"
                }
            ]
        }
        
        # Add template-specific fields
        if template_mode == "lost_settlement":
            card["body"].extend([
                {
                    "type": "TextBlock",
                    "text": "Payment Gateway",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "gateway",
                    "placeholder": "Enter payment gateway (e.g., bank account)"
                }
            ])
        
        # Add instructions field for all templates
        card["body"].extend([
            {
                "type": "TextBlock",
                "text": "Instructions (Optional)",
                "wrap": True
            },
            {
                "type": "Input.Text",
                "id": "instructions",
                "placeholder": "Any specific details or modifications to the template - your instructions will take priority over the template",
                "isMultiline": True
            },
            {
                "type": "Input.Toggle",
                "id": "hasAttachments",
                "title": "Mention attachments in email?",
                "value": "false"
            },
            {
                "type": "TextBlock",
                "text": "Note: This only mentions attachments in the text. To actually attach files, you'll need to add them when sending the email in your email client.",
                "wrap": True,
                "isSubtle": True,
                "size": "small"
            }
        ])
        
        # Add actions
        card["actions"] = [
            {
                "type": "Action.Submit",
                "title": "Generate Email",
                "data": {
                    "action": "generate_email",
                    "template": template_mode
                }
            },
            {
                "type": "Action.Submit",
                "title": "Back to Templates",
                "data": {
                    "action": "select_channel",
                    "channel": get_template_channel(template_mode)
                }
            }
        ]
    
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    return attachment
async def send_email_card(turn_context: TurnContext, template_mode="channel_selection", channel=None):
    """
    Sends an email composer card to the user.
    
    Args:
        turn_context: The turn context
        template_mode: The template mode to display
        channel: Email channel if in selection mode
    """
    reply = _create_reply(turn_context.activity)
    
    if template_mode == "channel_selection":
        reply.attachments = [create_channel_selection_card()]
    elif template_mode == "selection":
        reply.attachments = [create_email_card(template_mode, channel)]
    else:
        reply.attachments = [create_email_card(template_mode)]
    
    await turn_context.send_activity(reply)
async def handle_info_request(turn_context: TurnContext, info_type: str):
    """Handles requests for information about uploads or help"""
    if info_type == "upload":
        upload_info_card = {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "Container",
                    "style": "attention",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "How to Upload Files",
                            "size": "Large",
                            "weight": "Bolder",
                            "horizontalAlignment": "Center"
                        }
                    ],
                    "bleed": True
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "To upload and analyze files:",
                            "wrap": True,
                            "weight": "Bolder"
                        },
                        {
                            "type": "TextBlock",
                            "text": "1. Click the paperclip icon in the Teams chat input area\n2. Select your file from your device\n3. Send the file to me\n4. Once uploaded, you can ask questions about the file",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Supported File Types:",
                            "wrap": True,
                            "weight": "Bolder",
                            "spacing": "Medium"
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {
                                    "title": "Documents",
                                    "value": "PDF, DOC, DOCX, TXT"
                                },
                                {
                                    "title": "Images",
                                    "value": "JPG, JPEG, PNG, GIF, BMP"
                                },
                                {
                                    "title": "Not Supported",
                                    "value": "CSV, XLSX, XLS, XLSM"
                                }
                            ]
                        }
                    ]
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Back to Menu",
                    "data": {
                        "action": "new_chat"
                    }
                }
            ]
        }
        
        attachment = Attachment(
            content_type="application/vnd.microsoft.card.adaptive",
            content=upload_info_card
        )
        
        reply = _create_reply(turn_context.activity)
        reply.attachments = [attachment]
        await turn_context.send_activity(reply)
        
    elif info_type == "help":
        help_card = {
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "type": "AdaptiveCard",
            "version": "1.5",
            "body": [
                {
                    "type": "Container",
                    "style": "accent",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Available Commands & Tips",
                            "size": "Large",
                            "weight": "Bolder",
                            "horizontalAlignment": "Center"
                        }
                    ],
                    "bleed": True
                },
                {
                    "type": "Container",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Text Commands:",
                            "wrap": True,
                            "weight": "Bolder"
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {
                                    "title": "/email",
                                    "value": "Create a new email template"
                                },
                                {
                                    "title": "create email",
                                    "value": "Create a new email template"
                                },
                                {
                                    "title": "write email",
                                    "value": "Create a new email template"
                                },
                                {
                                    "title": "email template",
                                    "value": "Create a new email template"
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "Working with Files:",
                            "wrap": True,
                            "weight": "Bolder",
                            "spacing": "Medium"
                        },
                        {
                            "type": "TextBlock",
                            "text": "• Upload files using the paperclip icon in Teams\n• Ask questions about uploaded documents\n• Request analysis or summaries of documents\n• Reference file content in email drafts",
                            "wrap": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Sample Requests:",
                            "wrap": True,
                            "weight": "Bolder",
                            "spacing": "Medium"
                        },
                        {
                            "type": "TextBlock",
                            "text": "\"Write a professional email to the marketing team about the new product launch\"\n\n\"Summarize the key points from the document I just uploaded\"\n\n\"Draft a meeting invitation for a project kickoff on Friday\"",
                            "wrap": True
                        }
                    ]
                }
            ],
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Email Templates",
                    "data": {
                        "action": "show_template_categories"
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Back to Menu",
                    "data": {
                        "action": "new_chat"
                    }
                }
            ]
        }
        
        attachment = Attachment(
            content_type="application/vnd.microsoft.card.adaptive",
            content=help_card
        )
        
        reply = _create_reply(turn_context.activity)
        reply.attachments = [attachment]
        await turn_context.send_activity(reply)


# Example of handling email generation result
def create_email_result_card(email_text):
    """Creates a card displaying the generated email with copy options"""
    card = {
        "type": "AdaptiveCard",
        "version": "1.3",
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "body": [
            {
                "type": "TextBlock",
                "text": "Generated Email Template",
                "size": "large",
                "weight": "bolder",
                "horizontalAlignment": "center",
                "color": "accent"
            },
            {
                "type": "Container",
                "style": "emphasis",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": email_text,
                        "wrap": True,
                        "spacing": "medium"
                    }
                ]
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Create Another Email",
                "style": "positive",
                "data": {
                    "action": "create_email"
                }
            },
            {
                "type": "Action.Submit",
                "title": "Return to Home",
                "data": {
                    "action": "new_chat"
                }
            }
        ]
    }
    
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    return attachment
def create_template_selection_card():
    """Creates an adaptive card for selecting email template categories"""
    card = {
        "type": "AdaptiveCard",
        "version": "1.0",
        "body": [
            {
                "type": "TextBlock",
                "text": "Email Template Categories",
                "size": "large",
                "weight": "bolder"
            },
            {
                "type": "TextBlock",
                "text": "Select a template category to start your email",
                "wrap": True
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "📩 Introduction",
                                                "data": {
                                                    "action": "template_category",
                                                    "category": "introduction"
                                                },
                                                "style": "positive"
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "🔄 Follow-up",
                                                "data": {
                                                    "action": "template_category",
                                                    "category": "followup"
                                                }
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "📝 Request",
                                                "data": {
                                                    "action": "template_category",
                                                    "category": "request"
                                                }
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "🙏 Thank You",
                                                "data": {
                                                    "action": "template_category",
                                                    "category": "thankyou"
                                                }
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "📊 Status Update",
                                                "data": {
                                                    "action": "template_category",
                                                    "category": "status"
                                                }
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "📅 Meeting",
                                                "data": {
                                                    "action": "template_category",
                                                    "category": "meeting"
                                                }
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "⚠️ Urgent",
                                                "data": {
                                                    "action": "template_category",
                                                    "category": "urgent"
                                                }
                                            }
                                        ]
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "ActionSet",
                                        "actions": [
                                            {
                                                "type": "Action.Submit",
                                                "title": "✨ Custom",
                                                "data": {
                                                    "action": "template_category",
                                                    "category": "custom"
                                                }
                                            }
                                        ]
                                    }
                                ]
                            }
                        ]
                    }
                ]
            }
        ]
    }
    
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    return attachment
async def send_template_selection_card(turn_context: TurnContext):
    """Sends a template selection card to the user"""
    reply = _create_reply(turn_context.activity)
    reply.attachments = [create_template_selection_card()]
    await turn_context.send_activity(reply)
async def handle_template_selection(turn_context: TurnContext, category: str, state):
    """Handles a user's selection of an email template category"""
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Get template prompt for the selected category
        template_prompt = get_template_prompt(category)
        
        # Get the category-specific email card
        await send_category_email_card(turn_context, category)
        
    except Exception as e:
        logging.error(f"Error handling template selection: {e}")
        await turn_context.send_activity("I couldn't process your template selection. Please try again.")
def get_template_prompt(category: str) -> str:
    """Returns the specialized prompt for the selected template category"""
    # Base prompt structure with enhanced instructions
    base_prompt = {
        "introduction": """
You are composing an introduction email where the recipient doesn't know you. Your role is to create a professional, concise introduction email that accomplishes the following:
1. Opens with a friendly but professional greeting
2. Clearly identifies who you are and your organization
3. Explains your purpose for writing with specific value proposition
4. Includes a clear, low-pressure next step or call to action
5. Ends with a professional sign-off

Your tone should be warm, professional, and confident without being pushy. Avoid lengthy paragraphs - keep sentences short and focused. Use bullet points for any list of benefits or key points.

FORMAT THE EMAIL WITH:
- Greeting on its own line
- 3-4 short paragraphs with clear spacing between them
- Call to action as its own paragraph
- Professional signature
""",
        "followup": """
You are crafting a follow-up email to maintain momentum after a previous interaction. Your role is to create a message that:
1. References the specific previous interaction with date and context
2. Provides a concise summary of what was discussed/agreed upon
3. Clearly states the purpose of the follow-up (next steps, additional information, etc.)
4. Includes any relevant updates since the last interaction
5. Ends with a specific action item or question

Maintain a helpful, proactive tone that shows attention to detail and respect for the recipient's time. Avoid appearing passive-aggressive about response times - assume positive intent. Keep the email under 10 sentences total.

FORMAT THE EMAIL WITH:
- Brief, specific subject line referencing previous interaction
- Friendly opening acknowledging previous contact
- 2-3 concise paragraphs
- Clear next step or question highlighted in some way
- Professional but warm closing
""",
        "request": """
You are writing an email to make a specific request. Your role is to create a persuasive but respectful email that:
1. Opens with context that establishes relevance to the recipient
2. Clearly defines the specific request with all necessary details
3. Explains the rationale and benefits of fulfilling the request
4. Acknowledges any imposition and expresses appreciation
5. Provides a clear timeframe and process for response

Maintain a confident but courteous tone that respects the recipient's authority while clearly communicating the importance of your request. Avoid vague language - be specific about what you're asking for and why it matters.

FORMAT THE EMAIL WITH:
- Subject line that clearly indicates a request
- Brief context establishing relationship or relevance
- Detailed but concise explanation of the request
- Clear statement of timeline and preferred response method
- Appreciative closing
""",
        "thankyou": """
You are writing a thank-you email expressing genuine appreciation. Your role is to create a sincere, specific email that:
1. Clearly states what you're thankful for with specific details
2. Explains the positive impact or difference their action made
3. Includes a personal touch that shows authentic appreciation
4. If appropriate, mentions how you plan to pay it forward or reciprocate
5. Ends with warm, genuine closing

Maintain a warm, sincere tone throughout. Avoid generic platitudes - be specific about what was done and why it mattered. Keep the email concise but not rushed - quality over quantity.

FORMAT THE EMAIL WITH:
- Subject line clearly indicating gratitude
- Immediate, direct expression of thanks in first line
- 1-2 specific paragraphs detailing the impact
- Warm, personal closing
""",
        "status": """
You are creating a project status update email. Your role is to craft a clear, informative update that:
1. Starts with an executive summary of overall status (on track, at risk, etc.)
2. Provides specific updates on key workstreams with metrics where relevant
3. Clearly identifies any blockers, risks, or issues requiring attention
4. Outlines specific next steps and timeline
5. Includes any requests for input or decisions needed

Maintain a factual, solutions-oriented tone. Avoid placing blame or making excuses for delays. Use visual hierarchy (bullet points, bold text) to improve scannability. Keep the update concise but comprehensive.

FORMAT THE EMAIL WITH:
- Subject line with project name and update period
- Executive summary (1-2 sentences)
- Progress section with bullet points for each workstream
- Risks/issues section if applicable
- Next steps section with dates
- Clear signature with your role
""",
        "meeting": """
You are scheduling or following up on a meeting. Your role is to create a clear, actionable email that:
1. States the purpose of the meeting concisely
2. Provides essential logistical details (date, time, location/link)
3. Includes a brief agenda with time allocations
4. Specifies any preparation required from participants
5. Clarifies next steps or follow-ups expected after the meeting

Maintain an efficient, respectful tone that values everyone's time. Avoid unnecessary details - focus on what participants need to know. Make the email scannable for busy professionals.

FORMAT THE EMAIL WITH:
- Subject line with meeting purpose and date
- Brief context paragraph
- Clearly formatted logistics (When, Where, Who)
- Numbered or bulleted agenda
- Any preparation requirements clearly highlighted
- Professional closing
""",
        "urgent": """
You are writing an email requiring urgent attention. Your role is to create a clear, impactful message that:
1. Immediately identifies the urgent situation in the first sentence
2. Explains the specific impact or consequences if not addressed
3. Provides clear, specific actions needed with deadlines
4. Includes all necessary information to take action without follow-up questions
5. Offers availability for immediate discussion if needed

Maintain a serious, focused tone without creating panic. Avoid false urgency or hyperbole - be truthful about the timeline and impact. Use formatting to highlight key information and required actions.

FORMAT THE EMAIL WITH:
- Subject line starting with "URGENT:" followed by specific issue
- First sentence clearly stating the situation and timeline
- Brief explanation paragraph (3-5 sentences maximum)
- Bulleted list of required actions with deadlines
- Contact information for immediate response
- Professional but urgent closing
""",
        "custom": """
You are creating a custom professional email. Your role is to produce a well-structured message that:
1. Has a clear purpose identifiable in the first paragraph
2. Maintains appropriate tone for the business relationship
3. Includes all necessary information without unnecessary details
4. Has a logical flow from introduction to conclusion
5. Ends with a clear next step or call to action

Adapt your tone to match the context of the relationship and purpose. Use appropriate formality based on the recipient and situation. Make the email scannable with appropriate paragraph breaks and formatting.

FORMAT THE EMAIL WITH:
- Descriptive subject line 
- Clear introduction establishing purpose
- Logically organized body paragraphs
- Specific closing with next steps
- Professional signature
"""
    }
    
    # Return the prompt for the requested category
    return base_prompt.get(category, base_prompt["custom"])

async def send_category_email_card(turn_context: TurnContext, category: str):
    """Sends a category-specific email composition card with advanced adaptive card features"""
    # Get category-specific default values and placeholders
    defaults = get_category_defaults(category)
    
    # Base card structure with advanced features
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "TextBlock",
                "size": "medium",
                "weight": "bolder",
                "text": f"{defaults['title']} Email Template",
                "horizontalAlignment": "center",
                "wrap": True,
                "style": "heading"
            },
            {
                "type": "Container",
                "style": "emphasis",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Basic Information",
                        "weight": "bolder",
                        "size": "medium"
                    },
                    {
                        "type": "Input.Text",
                        "label": "Recipient",
                        "id": "recipient",
                        "placeholder": defaults["recipient_placeholder"],
                        "style": "text",
                        "isRequired": True,
                        "errorMessage": "Recipient is required"
                    },
                    {
                        "type": "Input.Text",
                        "label": "Subject",
                        "id": "subject",
                        "placeholder": defaults["subject_placeholder"],
                        "value": defaults["subject_default"],
                        "style": "text",
                        "isRequired": True,
                        "errorMessage": "Subject is required"
                    }
                ]
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Generate Email",
                "data": {
                    "action": "generate_category_email",
                    "category": category
                }
            },
            {
                "type": "Action.ShowCard",
                "title": "Advanced Options",
                "card": {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Style and Formatting",
                            "weight": "bolder"
                        },
                        {
                            "type": "Input.ChoiceSet",
                            "id": "formality",
                            "label": "Formality Level",
                            "style": "compact",
                            "choices": [
                                {
                                    "title": "Formal",
                                    "value": "formal"
                                },
                                {
                                    "title": "Semi-formal",
                                    "value": "semi-formal"
                                },
                                {
                                    "title": "Casual",
                                    "value": "casual"
                                }
                            ],
                            "value": "semi-formal"
                        },
                        {
                            "type": "Input.Toggle",
                            "id": "use_bullets",
                            "title": "Use bullet points for lists",
                            "valueOn": "true",
                            "valueOff": "false",
                            "value": "true"
                        },
                        {
                            "type": "Input.Number",
                            "id": "max_length",
                            "label": "Target Length (sentences)",
                            "placeholder": "10-15",
                            "min": 5,
                            "max": 30,
                            "value": 12
                        }
                    ]
                }
            },
            {
                "type": "Action.Submit",
                "title": "Back to Categories",
                "data": {
                    "action": "show_template_categories"
                }
            }
        ]
    }
    
    # Add purpose/details section based on category
    purpose_container = {
        "type": "Container",
        "items": [
            {
                "type": "TextBlock",
                "text": "Email Content",
                "weight": "bolder",
                "size": "medium"
            },
            {
                "type": "Input.Text",
                "label": "Purpose/Details",
                "id": "topic",
                "placeholder": defaults["purpose_placeholder"],
                "isMultiline": True,
                "style": "text",
                "isRequired": True,
                "errorMessage": "Please provide content details"
            }
        ]
    }
    
    # Add category-specific sections
    if category == "followup":
        followup_container = {
            "type": "Container",
            "items": [
                {
                    "type": "TextBlock",
                    "text": "Follow-up Details",
                    "weight": "bolder",
                    "size": "medium"
                },
                {
                    "type": "Input.Date",
                    "label": "Previous Interaction Date",
                    "id": "interaction_date"
                },
                {
                    "type": "Input.Text",
                    "label": "Previous Email/Conversation",
                    "id": "chain",
                    "placeholder": "Paste previous email or summarize conversation",
                    "isMultiline": True
                },
                {
                    "type": "Input.ChoiceSet",
                    "label": "Follow-up Type",
                    "id": "followup_type",
                    "style": "expanded",
                    "choices": [
                        {
                            "title": "Request update on prior discussion",
                            "value": "request_update"
                        },
                        {
                            "title": "Provide additional information",
                            "value": "provide_info"
                        },
                        {
                            "title": "Schedule next steps",
                            "value": "schedule_next"
                        },
                        {
                            "title": "Other (specify in details)",
                            "value": "other"
                        }
                    ],
                    "value": "request_update"
                }
            ]
        }
        card["body"].append(followup_container)
    
    elif category == "request":
        request_container = {
            "type": "Container",
            "items": [
                {
                    "type": "TextBlock",
                    "text": "Request Details",
                    "weight": "bolder",
                    "size": "medium"
                },
                {
                    "type": "Input.Text",
                    "label": "Requested Action",
                    "id": "requested_action",
                    "placeholder": "Specific action you're requesting",
                    "isRequired": True,
                    "errorMessage": "Please specify the requested action"
                },
                {
                    "type": "Input.Date",
                    "label": "Deadline",
                    "id": "deadline"
                },
                {
                    "type": "Input.ChoiceSet",
                    "label": "Priority Level",
                    "id": "priority",
                    "style": "compact",
                    "choices": [
                        {
                            "title": "High",
                            "value": "high"
                        },
                        {
                            "title": "Medium",
                            "value": "medium"
                        },
                        {
                            "title": "Low",
                            "value": "low"
                        }
                    ],
                    "value": "medium"
                }
            ]
        }
        card["body"].append(request_container)
    
    elif category == "meeting":
        meeting_container = {
            "type": "Container",
            "items": [
                {
                    "type": "TextBlock",
                    "text": "Meeting Details",
                    "weight": "bolder",
                    "size": "medium"
                },
                {
                    "type": "Input.Date",
                    "label": "Meeting Date",
                    "id": "meeting_date",
                    "isRequired": True,
                    "errorMessage": "Please select a meeting date"
                },
                {
                    "type": "Input.Time",
                    "label": "Meeting Time",
                    "id": "meeting_time",
                    "isRequired": True,
                    "errorMessage": "Please select a meeting time"
                },
                {
                    "type": "Input.Text",
                    "label": "Location/Link",
                    "id": "meeting_location",
                    "placeholder": "Physical location or virtual meeting link"
                },
                {
                    "type": "Input.Text",
                    "label": "Agenda Items",
                    "id": "agenda",
                    "placeholder": "List main points to discuss",
                    "isMultiline": True
                },
                {
                    "type": "Input.ChoiceSet",
                    "label": "Meeting Type",
                    "id": "meeting_type",
                    "style": "expanded",
                    "choices": [
                        {
                            "title": "Initial discussion",
                            "value": "initial"
                        },
                        {
                            "title": "Project update",
                            "value": "update"
                        },
                        {
                            "title": "Decision making",
                            "value": "decision"
                        },
                        {
                            "title": "Brainstorming session",
                            "value": "brainstorm"
                        },
                        {
                            "title": "Other (specify in details)",
                            "value": "other"
                        }
                    ],
                    "value": "initial"
                }
            ]
        }
        card["body"].append(meeting_container)
    
    elif category == "status":
        status_container = {
            "type": "Container",
            "items": [
                {
                    "type": "TextBlock",
                    "text": "Status Update Details",
                    "weight": "bolder",
                    "size": "medium"
                },
                {
                    "type": "Input.Text",
                    "label": "Project/Initiative Name",
                    "id": "project_name",
                    "placeholder": "Name of the project or initiative"
                },
                {
                    "type": "Input.ChoiceSet",
                    "label": "Overall Status",
                    "id": "overall_status",
                    "style": "expanded",
                    "choices": [
                        {
                            "title": "On track (Green)",
                            "value": "on_track"
                        },
                        {
                            "title": "At risk (Yellow)",
                            "value": "at_risk"
                        },
                        {
                            "title": "Off track (Red)",
                            "value": "off_track"
                        },
                        {
                            "title": "Completed (Blue)",
                            "value": "completed"
                        }
                    ],
                    "value": "on_track"
                },
                {
                    "type": "Input.Text",
                    "label": "Key Accomplishments",
                    "id": "accomplishments",
                    "placeholder": "List major achievements since last update",
                    "isMultiline": True
                },
                {
                    "type": "Input.Text",
                    "label": "Challenges/Blockers",
                    "id": "blockers",
                    "placeholder": "List any challenges or blockers",
                    "isMultiline": True
                }
            ]
        }
        card["body"].append(status_container)

    # Add purpose section after category-specific sections
    card["body"].append(purpose_container)
    
    # Common optional fields for all categories
    optional_container = {
        "type": "Container",
        "items": [
            {
                "type": "TextBlock",
                "text": "Additional Options",
                "weight": "bolder",
                "size": "medium"
            },
            {
                "type": "Input.Text",
                "label": "Key Points to Include",
                "id": "dos",
                "placeholder": "Important points, tone preferences, etc.",
                "isMultiline": True
            },
            {
                "type": "Input.Text",
                "label": "Points to Avoid",
                "id": "donts",
                "placeholder": "Topics to avoid, sensitive issues, etc.",
                "isMultiline": True
            },
            {
                "type": "Input.ChoiceSet",
                "label": "File Attachments",
                "id": "attachment_type",
                "style": "compact",
                "choices": [
                    {
                        "title": "No attachments",
                        "value": "none"
                    },
                    {
                        "title": "Reference uploaded files",
                        "value": "reference"
                    },
                    {
                        "title": "Will send attachments later",
                        "value": "later"
                    }
                ],
                "value": "none"
            }
        ]
    }
    
    card["body"].append(optional_container)
    
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    reply = _create_reply(turn_context.activity)
    reply.attachments = [attachment]
    await turn_context.send_activity(reply)

def get_category_defaults(category: str) -> dict:
    """Returns default values and placeholders for the selected category"""
    defaults = {
        "introduction": {
            "title": "Introduction",
            "recipient_placeholder": "Name of person you're introducing yourself to",
            "subject_placeholder": "Introduction - [Your Name] from [Your Company]",
            "subject_default": "Introduction - [Your Name] from [Your Company]",
            "purpose_placeholder": "Why you're reaching out and what value you can provide"
        },
        "followup": {
            "title": "Follow-Up",
            "recipient_placeholder": "Name of person you're following up with",
            "subject_placeholder": "Follow-up on our [meeting/conversation] about [topic]",
            "subject_default": "Follow-up on our discussion",
            "purpose_placeholder": "Key points from previous interaction and purpose of follow-up"
        },
        "request": {
            "title": "Request",
            "recipient_placeholder": "Name of person you're making the request to",
            "subject_placeholder": "Request: [Brief description of what you're requesting]",
            "subject_default": "Request: ",
            "purpose_placeholder": "Context and details of your request, including why it's important"
        },
        "thankyou": {
            "title": "Thank You",
            "recipient_placeholder": "Name of person you're thanking",
            "subject_placeholder": "Thank you for [what you're thanking them for]",
            "subject_default": "Thank you for your help",
            "purpose_placeholder": "Specific details about what you're thankful for and the impact it had"
        },
        "status": {
            "title": "Status Update",
            "recipient_placeholder": "Name(s) of person/team receiving the update",
            "subject_placeholder": "[Project Name]: Status Update - [Date/Period]",
            "subject_default": "Project Status Update",
            "purpose_placeholder": "Overall status, key accomplishments, challenges, and next steps"
        },
        "meeting": {
            "title": "Meeting",
            "recipient_placeholder": "Name(s) of meeting attendees",
            "subject_placeholder": "[Meeting Type]: [Topic] - [Date]",
            "subject_default": "Meeting Invitation",
            "purpose_placeholder": "Purpose of the meeting and expected outcomes"
        },
        "urgent": {
            "title": "Urgent",
            "recipient_placeholder": "Name of person who needs to take action",
            "subject_placeholder": "URGENT: [Specific issue requiring immediate attention]",
            "subject_default": "URGENT: Action Required",
            "purpose_placeholder": "Description of the urgent situation, impact, and required actions"
        },
        "custom": {
            "title": "Custom",
            "recipient_placeholder": "Enter recipient name(s)",
            "subject_placeholder": "Enter a clear, descriptive subject line",
            "subject_default": "",
            "purpose_placeholder": "Describe the purpose of your email and any specific details"
        }
    }
    
    return defaults.get(category, defaults["custom"])

async def generate_category_email(turn_context: TurnContext, state, category: str, form_data: dict):
    """Generates an email using AI based on enhanced category template and provided parameters using streaming mode"""
    # Extract common form data
    recipient = form_data.get("recipient", "")
    subject = form_data.get("subject", "")
    topic = form_data.get("topic", "")
    dos = form_data.get("dos", "")
    donts = form_data.get("donts", "")
    
    # Extract advanced options
    formality = form_data.get("formality", "semi-formal")
    use_bullets = form_data.get("use_bullets", "true") == "true"
    max_length = form_data.get("max_length", "12")
    attachment_type = form_data.get("attachment_type", "none")
    
    # Get the specialized template prompt for this category
    template_prompt = get_template_prompt(category)
    
    # Create prompt for the AI with category-specific instructions
    prompt = f"You are generating a {category} email following these specialized instructions:\n\n{template_prompt}\n\n"
    prompt += f"To: {recipient or 'Appropriate recipient'}\n"
    prompt += f"Subject: {subject or 'Appropriate subject based on context'}\n"
    prompt += f"Purpose/Details: {topic or 'Unspecified'}\n"
    
    # Add formality level instruction
    prompt += f"\nFORMATTING INSTRUCTIONS:\n"
    prompt += f"- Use a {formality} tone throughout the email\n"
    
    if use_bullets:
        prompt += "- Use bullet points for any lists or multiple items\n"
    else:
        prompt += "- Use paragraph format instead of bullet points\n"
        
    prompt += f"- Target length: Approximately {max_length} sentences total\n"
    
    # Add category-specific form data
    if category == "followup":
        interaction_date = form_data.get("interaction_date", "")
        previous_communication = form_data.get("chain", "")
        followup_type = form_data.get("followup_type", "request_update")
        
        prompt += f"Previous Interaction Date: {interaction_date}\n"
        prompt += f"Previous Communication: {previous_communication}\n"
        prompt += f"Follow-up Type: {followup_type}\n"
    
    elif category == "request":
        requested_action = form_data.get("requested_action", "")
        deadline = form_data.get("deadline", "")
        priority = form_data.get("priority", "medium")
        
        prompt += f"Requested Action: {requested_action}\n"
        prompt += f"Deadline: {deadline}\n"
        prompt += f"Priority Level: {priority}\n"
    
    elif category == "meeting":
        meeting_date = form_data.get("meeting_date", "")
        meeting_time = form_data.get("meeting_time", "")
        meeting_location = form_data.get("meeting_location", "")
        agenda = form_data.get("agenda", "")
        meeting_type = form_data.get("meeting_type", "initial")
        
        prompt += f"Meeting Date: {meeting_date}\n"
        prompt += f"Meeting Time: {meeting_time}\n"
        prompt += f"Location/Link: {meeting_location}\n"
        prompt += f"Agenda Items: {agenda}\n"
        prompt += f"Meeting Type: {meeting_type}\n"
    
    elif category == "status":
        project_name = form_data.get("project_name", "")
        overall_status = form_data.get("overall_status", "on_track")
        accomplishments = form_data.get("accomplishments", "")
        blockers = form_data.get("blockers", "")
        
        prompt += f"Project/Initiative Name: {project_name}\n"
        prompt += f"Overall Status: {overall_status}\n"
        prompt += f"Key Accomplishments: {accomplishments}\n"
        prompt += f"Challenges/Blockers: {blockers}\n"
    
    # Add common optional fields
    if dos:
        prompt += f"Important points to include: {dos}\n"
    if donts:
        prompt += f"Points to avoid: {donts}\n"
    
    # Handle attachments instruction
    if attachment_type == "reference":
        prompt += "\nIMPORTANT: The user has indicated there are file attachments for this email. "
        prompt += f"If any files have been uploaded to this conversation, use your file_search tool to retrieve relevant information related to '{subject} {topic}' "
        prompt += "and incorporate key insights into the email content."
        prompt += "\nInclude a line at the end mentioning that documents are attached for reference."
    elif attachment_type == "later":
        prompt += "\nInclude a line at the end mentioning that relevant documents will be sent in a follow-up email."
    
    # Initialize chat if needed
    if not state.get("assistant_id"):
        await initialize_chat(turn_context, state)
    
    # Send typing indicator immediately
    await turn_context.send_activity(create_typing_activity())
    
    # Create a background task for email generation
    email_text = ""
    
    # Create a client
    client = create_client()
    
    # Start the streaming process with typing indicators
    try:
        # Add the prompt to the thread
        thread_id = state.get("session_id")
        assistant_id = state.get("assistant_id")
        
        # Add message to thread to generate email
        client.beta.threads.messages.create(
            thread_id=thread_id,
            role="user",
            content=prompt
        )
        
        # Mark thread as busy (thread-safe)
        with conversation_states_lock:
            state["active_run"] = True
        
        with active_runs_lock:
            active_runs[thread_id] = True
        
        # Use streaming mode for better UX
        if TEAMS_AI_AVAILABLE:
            # Create a custom collector for the email text
            class EmailCollector:
                def __init__(self):
                    self.complete_text = ""
                
                def collect_text(self, text):
                    self.complete_text += text
            
            collector = EmailCollector()
            
            # Create a wrapper function for the streaming process
            async def streaming_wrapper(tc, state, msg=None):
                # Use enhanced streaming with Teams AI library
                streamer = StreamingResponse(tc)
                
                # Track the run ID for proper cleanup
                run_id = None
                
                try:
                    # Create run with streaming
                    run = client.beta.threads.runs.create(
                        thread_id=thread_id,
                        assistant_id=assistant_id,
                        stream=True
                    )
                    
                    run_id = run.id
                    
                    # Process the streaming response
                    previous_text = ""
                    for chunk in run.iter_chunks():
                        if hasattr(chunk, "data") and hasattr(chunk.data, "delta"):
                            delta = chunk.data.delta
                            if hasattr(delta, "content") and delta.content:
                                for content in delta.content:
                                    if content.type == "text" and hasattr(content.text, "value"):
                                        text_piece = content.text.value
                                        # Collect the text for later use
                                        collector.collect_text(text_piece)
                                        # No need to send the email text as it's generated
                
                    # Don't send the response now; we'll use it to build the card
                    
                except Exception as e:
                    logging.error(f"Error in streaming email generation: {e}")
                    await tc.send_activity("I encountered an error while generating your email template. Please try again.")
                finally:
                    # Clean up active runs
                    with conversation_states_lock:
                        state["active_run"] = False
                    
                    with active_runs_lock:
                        if thread_id in active_runs:
                            del active_runs[thread_id]
            
            # Send progress message and typing indicators
            await turn_context.send_activity("Generating your email template...")
            
            # Start a typing indicator task
            typing_task = asyncio.create_task(send_periodic_typing(turn_context, 4))
            
            try:
                # Run the streaming process to collect the email text
                await streaming_wrapper(turn_context, state)
                
                # Get the generated email text from the collector
                email_text = collector.complete_text
            finally:
                # Cancel the typing indicator task
                typing_task.cancel()
                try:
                    await typing_task
                except asyncio.CancelledError:
                    pass
                await turn_context.send_activity(create_typing_stop_activity())
        else:
            # Use custom streaming implementation if Teams AI not available
            # Setup a custom collector similar to above
            class CustomEmailCollector:
                def __init__(self):
                    self.complete_text = ""
                
                def add_text(self, text):
                    self.complete_text += text
            
            collector = CustomEmailCollector()
            
            # Send a progress message
            await turn_context.send_activity("Generating your email template...")
            
            # Start typing indicator task
            typing_task = asyncio.create_task(send_periodic_typing(turn_context, 4))
            
            try:
                # Create a run
                run = client.beta.threads.runs.create(
                    thread_id=thread_id,
                    assistant_id=assistant_id
                )
                
                run_id = run.id
                
                # Poll for completion with typing indicators
                max_wait_time = 120  # Maximum wait time in seconds
                wait_interval = 2    # Check interval in seconds
                elapsed_time = 0
                
                while elapsed_time < max_wait_time:
                    # Check run status
                    run_status = client.beta.threads.runs.retrieve(
                        thread_id=thread_id,
                        run_id=run_id
                    )
                    
                    # Check for completion
                    if run_status.status == "completed":
                        # Get the complete message
                        messages = client.beta.threads.messages.list(
                            thread_id=thread_id,
                            order="desc",
                            limit=1
                        )
                        
                        if messages.data:
                            latest_message = messages.data[0]
                            message_text = ""
                            
                            for content_part in latest_message.content:
                                if content_part.type == 'text':
                                    message_text += content_part.text.value
                            
                            # Collect the complete email text
                            collector.add_text(message_text)
                            break
                            
                    # Check for failure states
                    elif run_status.status in ["failed", "cancelled", "expired"]:
                        logging.error(f"Run {run_id} ended with status: {run_status.status}")
                        await turn_context.send_activity(f"I encountered an issue while generating the email template. Please try again.")
                        break
                    
                    # Wait before next check
                    await asyncio.sleep(wait_interval)
                    elapsed_time += wait_interval
                
                # Get the collected email text
                email_text = collector.complete_text
                
            finally:
                # Clean up
                with conversation_states_lock:
                    state["active_run"] = False
                
                with active_runs_lock:
                    if thread_id in active_runs:
                        del active_runs[thread_id]
                
                # Cancel typing indicator task
                typing_task.cancel()
                try:
                    await typing_task
                except asyncio.CancelledError:
                    pass
                await turn_context.send_activity(create_typing_stop_activity())
        
        # If we have email text, create and send the card
        if email_text:
            # Create an enhanced email result card
            email_card = {
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "type": "AdaptiveCard",
                "version": "1.5",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": f"Generated {category.title()} Email",
                        "size": "large",
                        "weight": "bolder",
                        "horizontalAlignment": "center",
                        "wrap": True,
                        "style": "heading"
                    },
                    {
                        "type": "Container",
                        "style": "accent",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": subject,
                                "weight": "bolder",
                                "wrap": True
                            },
                            {
                                "type": "TextBlock",
                                "text": f"To: {recipient}",
                                "wrap": True
                            }
                        ],
                        "bleed": True
                    },
                    {
                        "type": "Container",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": email_text,
                                "wrap": True
                            }
                        ],
                        "style": "default"
                    }
                ],
                "actions": [
                    {
                        "type": "Action.Submit",
                        "title": "Create Another Email",
                        "data": {
                            "action": "show_template_categories"
                        }
                    },
                    {
                        "type": "Action.ShowCard",
                        "title": "Edit Email",
                        "card": {
                            "type": "AdaptiveCard",
                            "body": [
                                {
                                    "type": "Input.Text",
                                    "label": "Edit Content",
                                    "id": "edit_content",
                                    "isMultiline": True,
                                    "value": email_text
                                }
                            ],
                            "actions": [
                                {
                                    "type": "Action.Submit",
                                    "title": "Update Email",
                                    "data": {
                                        "action": "update_email_content"
                                    }
                                }
                            ]
                        }
                    }
                ]
            }
            
            # Create attachment
            attachment = Attachment(
                content_type="application/vnd.microsoft.card.adaptive",
                content=email_card
            )
            
            reply = _create_reply(turn_context.activity)
            reply.attachments = [attachment]
            await turn_context.send_activity(reply)
        else:
            await turn_context.send_activity("I'm sorry, I couldn't generate the email template. Please try again.")
            
    except Exception as e:
        logging.error(f"Error generating email: {e}")
        traceback.print_exc()
        
        # Clean up active runs on error
        with conversation_states_lock:
            state["active_run"] = False
        
        with active_runs_lock:
            if thread_id in active_runs:
                del active_runs[thread_id]
        
        await turn_context.send_activity("I'm sorry, I encountered an error while generating your email template. Please try again.")

async def generate_email(turn_context: TurnContext, state, template_id, recipient=None, firstname=None, gateway=None, subject=None, instructions=None, chain=None, has_attachments=False):
    """
    Generates an email using AI based on template or provided parameters with enhanced flexibility and creativity.
    
    Args:
        turn_context: The turn context
        state: The conversation state
        template_id: The template ID to use
        recipient: The recipient's email (optional)
        firstname: The client's first name (optional)
        gateway: The payment gateway (for lost settlement template) (optional)
        subject: The email subject (for generic template) (optional)
        instructions: Additional instructions for customization (optional)
        chain: Previous email chain (optional)
        has_attachments: Whether to mention attachments
    """
    # Send typing indicator
    await turn_context.send_activity(create_typing_activity())
    
    # Initialize chat if needed
    if not state.get("assistant_id"):
        await initialize_chat(turn_context, state)
    
    try:
        # Get template channel, title and type information
        template_channel = get_template_channel(template_id)
        template_title = get_template_title(template_id)
        
        # Get base template content if using a template
        template_subject = ""
        template_content = ""
        
        if template_id != "generic":
            template_subject, template_content = get_template_content(
                template_id, 
                firstname=firstname or "{FIRSTNAME}",
                gateway=gateway or "{GATEWAY}"
            )
        
        # Create a detailed request with the EMAIL_PROMPT as context and balanced instructions for creativity
        creative_instruction = """
GUIDANCE ON TEMPLATE USE VS. CREATIVITY:
1. Use the specified template as a STARTING POINT only.
2. PRIORITIZE any specific user instructions over rigid template adherence.
3. Feel free to IMPROVE and PERSONALIZE the template language to sound more natural.
4. BE CREATIVE while maintaining First Choice Debt Relief's professional brand voice.
5. Focus on CLARITY, EMPATHY, and EFFECTIVENESS rather than strict format adherence.
6. The email should be READY TO SEND without further editing.
7. Include all standard parts (greeting, body, closing, signature) in a professional format.
8. Focus on accomplishing the EMAIL'S PURPOSE rather than following exact template wording.
"""
        
        # Create the specific email request
        email_request = f"""
I need to draft a {template_title} for First Choice Debt Relief.

EMAIL SPECIFICATIONS:
- TYPE: {template_title}
- CHANNEL: {template_channel.replace('_', ' ').title() if template_channel else 'Generic'}
"""

        # Add recipient information
        if recipient:
            email_request += f"- RECIPIENT: {recipient}\n"
        
        # Add subject line information
        if template_id == "generic" and subject:
            email_request += f"- SUBJECT: {subject}\n"
        elif template_subject:
            email_request += f"- SUBJECT: {template_subject}\n"
        
        # Add specific parameters
        if firstname:
            email_request += f"- CLIENT NAME: {firstname}\n"
        
        if gateway and template_id == "lost_settlement":
            email_request += f"- PAYMENT GATEWAY: {gateway}\n"
        
        # Add attachments note if needed
        if has_attachments:
            email_request += "- ATTACHMENTS: Yes (please mention files are attached)\n"
        
        # Add previous email thread if available
        if chain:
            email_request += f"\nPREVIOUS EMAIL THREAD:\n```\n{chain}\n```\n"
        
        # Add template content as reference
        if template_content:
            email_request += f"\nTEMPLATE REFERENCE (use as starting point only):\n```\n{template_content}\n```\n"
        
        # Add special instructions with absolute highest priority
        if instructions:
            email_request += f"\nUSER INSTRUCTIONS (HIGHEST PRIORITY - MUST FOLLOW):\n{instructions}\n"
        
        # Create a client and build the full prompt with the EMAIL_PROMPT
        client = create_client()
        
        # Combine everything into a structured prompt
        full_prompt = f"{EMAIL_PROMPT}\n\n{creative_instruction}\n\n{email_request}"
        
        # Add specific template-based guidance
        if template_id == "welcome":
            full_prompt += "\nThis is a WELCOME EMAIL: Focus on enthusiasm, clear next steps, and making the client feel supported as they begin their debt resolution journey."
        elif template_id == "legal_update":
            full_prompt += "\nThis is a LEGAL UPDATE EMAIL: Balance providing important information without causing unnecessary alarm. Emphasize partnership and support through the legal process."
        elif template_id == "lost_settlement":
            full_prompt += "\nThis is a LOST SETTLEMENT EMAIL: Clearly communicate urgency while maintaining empathy. Focus on immediate action steps and how we can help recover the situation."
        elif template_id == "sales_quote" or template_id == "sales_quick_quote":
            full_prompt += "\nThis is a SALES QUOTE EMAIL: Focus on the value proposition, potential savings, and creating gentle urgency to act. Be consultative rather than pushy."
        elif template_id.startswith("sales_"):
            full_prompt += "\nThis is a SALES EMAIL: Balance providing valuable information with a clear call to action. Focus on how our service addresses the client's financial pain points."
        elif template_id == "introduction":
            full_prompt += "\nThis is an INTRODUCTION EMAIL: Make a strong first impression with a clear, concise introduction and value proposition. Keep it professional and non-pushy."
        elif template_id == "followup":
            full_prompt += "\nThis is a FOLLOW-UP EMAIL: Gently remind the recipient of previous contact without being pushy. Show understanding of their decision-making process."
        
        # Make the final request
        result = await process_conversation_internal(
            client=client,
            session=state["session_id"],
            prompt=full_prompt,
            assistant=state["assistant_id"],
            stream_output=False
        )
        
        # Log the template used for analytics
        logging.info(f"Generated {template_title} email for user {state.get('user_id', 'unknown')}")
        
        # Extract and format the email
        if isinstance(result, dict) and "response" in result:
            email_text = result["response"]
            
            # Clean up the email text - remove any markdown formatting artifacts
            email_text = email_text.replace("```", "").strip()
            
            # If the email starts with "Subject:", extract just the body
            if email_text.lower().startswith("subject:"):
                # Try to find the first line break after "Subject:"
                subject_end_idx = email_text.find("\n")
                if subject_end_idx > 0:
                    email_text = email_text[subject_end_idx+1:].strip()
            
            # Save the generated email in the state for potential editing
            with conversation_states_lock:
                state["last_generated_email"] = email_text
                state["last_email_template"] = template_id
                state["last_email_data"] = {
                    "recipient": recipient,
                    "firstname": firstname,
                    "gateway": gateway,
                    "subject": subject,
                    "instructions": instructions,
                    "chain": chain,
                    "has_attachments": has_attachments
                }
            
            # Create an enhanced email result card with rich formatting
            email_card = {
                "type": "AdaptiveCard",
                "version": "1.3",
                "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                "body": [
                    {
                        "type": "TextBlock",
                        "text": f"Generated {template_title}",
                        "size": "large",
                        "weight": "bolder",
                        "horizontalAlignment": "center",
                        "color": "accent"
                    },
                    {
                        "type": "Container",
                        "style": "emphasis",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": email_text,
                                "wrap": True,
                                "spacing": "medium"
                            }
                        ]
                    }
                ],
                "actions": [
                    {
                        "type": "Action.Submit",
                        "title": "Edit This Email",
                        "style": "default",
                        "data": {
                            "action": "edit_email"
                        }
                    },
                    {
                        "type": "Action.Submit",
                        "title": "Create Another Email",
                        "style": "positive",
                        "data": {
                            "action": "create_email"
                        }
                    },
                    {
                        "type": "Action.Submit",
                        "title": "Regenerate",
                        "style": "default",
                        "data": {
                            "action": "regenerate_email",
                            "template": template_id,
                            "recipient": recipient,
                            "firstname": firstname,
                            "gateway": gateway,
                            "subject": subject,
                            "instructions": instructions,
                            "chain": chain,
                            "hasAttachments": "true" if has_attachments else "false"
                        }
                    }
                ]
            }
            
            # Add copy to clipboard action for easy copying
            email_card["actions"].append({
                "type": "Action.ShowCard",
                "title": "Copy Text",
                "card": {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Email text ready to copy:",
                            "wrap": True
                        },
                        {
                            "type": "Input.Text",
                            "id": "email_text_to_copy",
                            "value": email_text,
                            "isMultiline": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Select the text above and use Ctrl+C to copy",
                            "wrap": True,
                            "size": "small",
                            "isSubtle": True
                        }
                    ]
                }
            })
            
            # Create attachment
            attachment = Attachment(
                content_type="application/vnd.microsoft.card.adaptive",
                content=email_card
            )
            
            reply = _create_reply(turn_context.activity)
            reply.attachments = [attachment]
            await turn_context.send_activity(reply)
        else:
            await turn_context.send_activity("I'm sorry, I couldn't generate the email template. Please try again.")
    
    except Exception as e:
        # Handle and log any errors
        logging.error(f"Error generating email: {str(e)}")
        traceback.print_exc()
        await turn_context.send_activity(f"I encountered an error while generating your email template. Please try again or select a different template.")
        
        # Try fallback direct completion
        try:
            await send_fallback_response(turn_context, f"Please generate a professional {get_template_title(template_id)} email for First Choice Debt Relief.")
        except Exception as fallback_error:
            logging.error(f"Fallback email generation also failed: {fallback_error}")
async def send_periodic_typing(turn_context: TurnContext, interval_seconds: int):
    """Sends typing indicators periodically until the task is cancelled"""
    try:
        while True:
            await turn_context.send_activity(create_typing_activity())
            await asyncio.sleep(interval_seconds)
    except asyncio.CancelledError:
        # Task was cancelled, exit cleanly
        pass
async def send_new_chat_card(turn_context: TurnContext):
    """Sends an enhanced card with buttons to start a new chat session"""
    reply = _create_reply(turn_context.activity)
    reply.attachments = [create_new_chat_card()]
    await turn_context.send_activity(reply)


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

    # Try to recover with fallback response
    try:
        await send_fallback_response(context, None)
    except:
        pass  # If even this fails, just continue

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
# Modified bot_logic function to properly handle email card submissions
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
                "creation_time": time.time(),
                "last_activity_time": time.time()
            }
        else:
            # Update last activity time
            conversation_states[conversation_id]["last_activity_time"] = time.time()
            
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
                    "creation_time": time.time(),
                    "last_activity_time": time.time()
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
        # First, check if this is a card submission before checking for text content
        # This is the key fix for email card submissions
        value_data = getattr(turn_context.activity, 'value', None)
        if value_data:
            logging.info(f"Card submission detected: {value_data}")
            try:
                # Handle card submission directly
                await handle_card_actions(turn_context, value_data)
                return  # Exit early since we've handled the card action
            except Exception as card_e:
                logging.error(f"Error handling card submission: {card_e}")
                await turn_context.send_activity("I had trouble processing your form submission. Please try again.")
                return
        
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
                state["last_activity_time"] = current_time
                
                # Clear any pending messages
                with pending_messages_lock:
                    if conversation_id in pending_messages:
                        pending_messages[conversation_id].clear()
                
                await turn_context.send_activity("Your previous session has expired. Creating a new session for you.")
        
        # Track if thread is currently processing (thread-safe)
        is_thread_busy = False
        with conversation_states_lock:
            is_thread_busy = state.get("active_run", False)
            
            # Double-check with active_runs for consistency
            with active_runs_lock:
                thread_id = state.get("session_id")
                if thread_id:
                    if thread_id in active_runs:
                        is_thread_busy = True
                        state["active_run"] = True
                    elif state.get("active_run", False):
                        # State says active but active_runs doesn't have it - fix the inconsistency
                        state["active_run"] = False
                        is_thread_busy = False
        
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
                await turn_context.send_activity("To upload files, use the paperclip icon and select from your device storage only - do not use OneDrive or shared locations. Text, PDF, Image and Doc files are presently supported.")
    
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
# async def handle_file_upload(turn_context: TurnContext, state, message_text=None):
#     """Handle file uploads from Teams with clear messaging about supported types"""
    
#     for attachment in turn_context.activity.attachments:
#         try:
#             # Send typing indicator
#             await turn_context.send_activity(create_typing_activity())
            
#             # Check if it's a direct file upload (locally uploaded file)
#             if hasattr(attachment, 'content_type') and attachment.content_type == ContentType.FILE_DOWNLOAD_INFO:
#                 # Download the file using the Teams-specific logic
#                 file_path = await download_file(turn_context, attachment)
                
#                 if not file_path:
#                     # File was either not downloaded or rejected
#                     continue
                    
#                 # Check file extension to ensure we only accept supported types
#                 file_ext = os.path.splitext(attachment.name)[1].lower()
#                 if file_ext in ['.csv', '.xlsx', '.xls', '.xlsm']:
#                     await turn_context.send_activity("Sorry, CSV and Excel files are not supported. Please upload PDF, DOC, DOCX, or TXT files only.")
#                     continue
                
#                 # Process the file with message text if provided
#                 await process_uploaded_file(turn_context, state, file_path, attachment.name, message_text)
#             else:
#                 # Check if this is likely an OneDrive or SharePoint file
#                 is_internal_file = False
#                 if hasattr(attachment, 'content_type'):
#                     internal_file_indicators = [
#                         "sharepoint", 
#                         "onedrive", 
#                         "vnd.microsoft.teams.file", 
#                         "application/vnd.microsoft.teams.file"
#                     ]
                    
#                     for indicator in internal_file_indicators:
#                         if indicator.lower() in attachment.content_type.lower():
#                             is_internal_file = True
#                             break
                
#                 if is_internal_file:
#                     # Provide clear message that only local uploads are supported
#                     await turn_context.send_activity("I'm sorry, but I can only process files uploaded directly from your device. Files shared from OneDrive, SharePoint, or other internal sources are not currently supported. Please download the file to your device first, then upload it directly.")
#                 else:
#                     # For other attachment types, provide general guidance
#                     await turn_context.send_activity("To upload a file, please use the file upload feature in Teams to send files directly from your device. Click the paperclip icon in the chat input area to upload a file.")
                
#         except Exception as e:
#             logger.error(f"Error processing file: {str(e)}")
#             traceback.print_exc()
#             await turn_context.send_activity(f"Error processing file: {str(e)}")
async def handle_file_upload(turn_context: TurnContext, state, message_text=None):
    """Handle file uploads from Teams with support for both direct uploads and SharePoint guest links"""
    
    for attachment in turn_context.activity.attachments:
        try:
            # Send typing indicator
            await turn_context.send_activity(create_typing_activity())
            
            # Check if it's a direct file upload (locally uploaded file)
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
                # Check if this is likely an OneDrive or SharePoint file
                is_internal_file = False
                sharepoint_url = None

                if hasattr(attachment, 'content_type'):
                    internal_file_indicators = [
                        "sharepoint", 
                        "onedrive", 
                        "vnd.microsoft.teams.file", 
                        "application/vnd.microsoft.teams.file"
                    ]
                    
                    for indicator in internal_file_indicators:
                        if indicator.lower() in attachment.content_type.lower():
                            is_internal_file = True
                            break

                if is_internal_file:
                    # Try to handle SharePoint/OneDrive files
                    try:
                        # Extract the SharePoint URL from the attachment
                        if hasattr(attachment, 'content_url'):
                            sharepoint_url = attachment.content_url
                        elif hasattr(attachment, 'content') and isinstance(attachment.content, dict):
                            # Check various possible content structures
                            if 'downloadUrl' in attachment.content:
                                sharepoint_url = attachment.content['downloadUrl']
                            elif 'webUrl' in attachment.content:
                                sharepoint_url = attachment.content['webUrl']
                            elif 'contentUrl' in attachment.content:
                                sharepoint_url = attachment.content['contentUrl']
                        # Sometimes the URL might be in the contentUrl property directly
                        elif hasattr(attachment, 'contentUrl'):
                            sharepoint_url = attachment.contentUrl
                            
                        # Log the attachment structure for debugging
                        logging.info(f"SharePoint attachment structure: {vars(attachment)}")
                        
                        if sharepoint_url:
                            await turn_context.send_activity(f"Processing shared file from SharePoint/OneDrive: {attachment.name}")
                            
                            # Download the file using SharePoint authentication
                            file_path = await download_sharepoint_file(sharepoint_url, attachment.name)
                            
                            if file_path:
                                # Check file extension to ensure we only accept supported types
                                file_ext = os.path.splitext(attachment.name)[1].lower()
                                if file_ext in ['.csv', '.xlsx', '.xls', '.xlsm']:
                                    await turn_context.send_activity("Sorry, CSV and Excel files are not supported. Please upload PDF, DOC, DOCX, or TXT files only.")
                                    # Clean up the downloaded file
                                    if os.path.exists(file_path):
                                        try:
                                            os.remove(file_path)
                                        except Exception as del_e:
                                            logging.error(f"Error removing rejected file: {del_e}")
                                    continue
                                
                                # Process the file as normal
                                await process_uploaded_file(turn_context, state, file_path, attachment.name, message_text)
                            else:
                                await turn_context.send_activity("I had trouble accessing that shared file. Please check permissions or try uploading directly from your device.")
                        else:
                            # If URL extraction failed, log the attachment details for debugging
                            logging.error(f"Could not extract SharePoint URL from attachment: {attachment}")
                            logging.error(f"Attachment type: {type(attachment)}")
                            if hasattr(attachment, 'content'):
                                logging.error(f"Content type: {type(attachment.content)}")
                                if isinstance(attachment.content, dict):
                                    logging.error(f"Content keys: {attachment.content.keys()}")
                            
                            await turn_context.send_activity("I couldn't extract the file URL from the shared file. Please try uploading directly from your device.")
                    except Exception as sp_e:
                        logging.error(f"Error processing SharePoint file: {str(sp_e)}")
                        traceback.print_exc()
                        await turn_context.send_activity(f"Error processing shared file: {str(sp_e)}. Please try uploading directly from your device.")
                else:
                    # For other attachment types, provide general guidance
                    await turn_context.send_activity("To upload a file, please use the file upload feature in Teams to send files directly from your device or share from SharePoint/OneDrive. Click the paperclip icon in the chat input area to upload a file.")
                
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
                                vector_store = client.vector_stores.create(name=f"Assistant_{state['assistant_id']}_Store")
                                vector_store_id = vector_store.id
                                state["vector_store_id"] = vector_store_id
                            
                            # Upload to vector store
                            with open(temp_path, "rb") as file_stream:
                                file_batch = client.vector_stores.file_batches.upload_and_poll(
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
                    
                    # Update active_runs dictionary (thread-safe)
                    with active_runs_lock:
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
    if user_message.lower() in ["/email", "create email", "write email", "email template", "email"]:
        await send_email_card(turn_context)
        return
    if user_message.lower() in ["/new", "/reset", "new chat", "start over", "reset chat"]:
        await handle_new_chat_command(turn_context, state, conversation_id)
        return
        
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
    current_session_id = None
    with conversation_states_lock:
        state["active_run"] = True
        current_session_id = state.get("session_id")
    
    if current_session_id:
        with active_runs_lock:
            active_runs[current_session_id] = True
    
    try:
        # Double-verify resources before proceeding
        client = create_client()
        validation = await validate_resources(client, current_session_id, stored_assistant_id)
        
        # If any resource is invalid, force recovery
        if not validation["thread_valid"] or not validation["assistant_valid"]:
            logging.warning(f"Resource validation failed for user {user_id}: thread_valid={validation['thread_valid']}, assistant_valid={validation['assistant_valid']}")
            raise Exception("Invalid conversation resources detected - forcing recovery")
            
        # Use the optimal streaming approach based on available libraries and preferences
        if TEAMS_AI_AVAILABLE:
            # Use enhanced streaming with Teams AI library
            await stream_with_teams_ai(turn_context, state, user_message)
        else:
            # Use custom TeamsStreamingResponse if Teams AI library is not available
            await stream_with_custom_implementation(turn_context, state, user_message)
        
        # Mark thread as no longer busy (thread-safe)
        with conversation_states_lock:
            state["active_run"] = False
            current_session_id = state.get("session_id")
        
        with active_runs_lock:
            if current_session_id in active_runs:
                del active_runs[current_session_id]
        
        # Process any pending messages
        await process_pending_messages(turn_context, state, conversation_id)
            
    except Exception as e:
        # Mark thread as no longer busy even on error (thread-safe)
        with conversation_states_lock:
            state["active_run"] = False
            current_session_id = state.get("session_id")
            
        with active_runs_lock:
            if current_session_id in active_runs:
                del active_runs[current_session_id]
            
        # Don't show raw error details to users
        logging.error(f"Error in handle_text_message for user {user_id}: {str(e)}")
        traceback.print_exc()
        await turn_context.send_activity("I'm sorry, I encountered a problem while processing your message. Please try again.")
        
        # Try a fallback direct completion if there was a severe error
        try:
            await send_fallback_response(turn_context, user_message)
        except Exception as fallback_error:
            logging.error(f"Fallback response also failed: {fallback_error}")

# Modified process_pending_messages function to fix the run conflict
async def process_pending_messages(turn_context: TurnContext, state, conversation_id):
    """Process any pending messages in the queue safely"""
    with pending_messages_lock:
        if conversation_id in pending_messages and pending_messages[conversation_id]:
            # Process one message at a time to avoid race conditions
            if len(pending_messages[conversation_id]) > 0:
                next_message = pending_messages[conversation_id].popleft()
                await turn_context.send_activity("Now addressing your follow-up message...")
                
                # IMPORTANT: Don't modify the original turn_context
                # Instead, directly process the message through OpenAI API
                try:
                    # Get the thread and assistant IDs
                    thread_id = state.get("session_id")
                    assistant_id = state.get("assistant_id")
                    
                    if not thread_id or not assistant_id:
                        await turn_context.send_activity("I'm having trouble with your follow-up question. Let's start a new conversation.")
                        return
                    
                    # Create a new client
                    client = create_client()
                    
                    # Send typing indicator
                    await turn_context.send_activity(create_typing_activity())
                    
                    # Check for any existing active runs and cancel them first
                    try:
                        runs = client.beta.threads.runs.list(thread_id=thread_id, limit=1)
                        if runs.data:
                            latest_run = runs.data[0]
                            if latest_run.status in ["in_progress", "queued", "requires_action"]:
                                logging.info(f"Cancelling active run {latest_run.id} before processing follow-up")
                                client.beta.threads.runs.cancel(thread_id=thread_id, run_id=latest_run.id)
                                await asyncio.sleep(2)  # Wait for cancellation to take effect
                    except Exception as cancel_e:
                        logging.warning(f"Error checking or cancelling runs: {cancel_e}")
                    
                    # CRITICAL: Wait to ensure no active runs
                    active_run_found = True
                    max_wait = 5  # Maximum 5 seconds to wait
                    start_time = time.time()
                    
                    while active_run_found and (time.time() - start_time) < max_wait:
                        try:
                            # Check if any active runs still exist
                            runs = client.beta.threads.runs.list(thread_id=thread_id, limit=1)
                            active_run_found = False
                            
                            if runs.data:
                                for run in runs.data:
                                    if run.status in ["in_progress", "queued", "requires_action"]:
                                        active_run_found = True
                                        logging.info(f"Still waiting for run {run.id} to cancel...")
                                        await asyncio.sleep(1)
                                        break
                            
                            if not active_run_found:
                                break
                        except Exception:
                            break  # If we can't check, just proceed
                    
                    # Add the follow-up message to the thread
                    client.beta.threads.messages.create(
                        thread_id=thread_id,
                        role="user",
                        content=next_message
                    )
                    
                    # Process the response with streaming
                    if TEAMS_AI_AVAILABLE:
                        await stream_with_teams_ai(turn_context, state, None)
                    else:
                        await stream_with_custom_implementation(turn_context, state, None)
                        
                except Exception as e:
                    logging.error(f"Error processing follow-up: {e}")
                    traceback.print_exc()
                    await turn_context.send_activity(f"I had trouble processing your follow-up. Please try asking again.")
# Add this to the end of handle_text_message function
async def ensure_no_active_runs(client, thread_id):
    """Ensure there are no active runs on the thread"""
    try:
        runs = client.beta.threads.runs.list(thread_id=thread_id, limit=1)
        active_run_found = False
        
        if runs.data:
            for run in runs.data:
                if run.status in ["in_progress", "queued", "requires_action"]:
                    # Try to cancel it
                    logging.info(f"Cancelling active run {run.id}")
                    try:
                        client.beta.threads.runs.cancel(thread_id=thread_id, run_id=run.id)
                    except Exception as cancel_e:
                        logging.warning(f"Error cancelling run {run.id}: {cancel_e}")
                    
                    active_run_found = True
        
        # If we found an active run, wait for it to be cancelled
        if active_run_found:
            await asyncio.sleep(2)  # Give the cancellation time to take effect
    except Exception as e:
        logging.warning(f"Error checking for active runs: {e}")
# Right before the function returns
async def cleanup_after_message():
    """Ensure all state is properly reset after message processing"""
    # Force reset of active run markers
    with conversation_states_lock:
        state["active_run"] = False
    
    with active_runs_lock:
        if current_session_id in active_runs:
            del active_runs[current_session_id]
            
    # Send a special "ready" signal to the client
    # This doesn't need to be visible, but helps ensure the bot is ready
    try:
        activity = Activity(
            type="event",
            name="bot_ready",
            value={"timestamp": str(time.time())}
        )
        await turn_context.send_activity(activity)
    except:
        pass  # Ignore errors with the ready signal
async def stream_with_teams_ai(turn_context: TurnContext, state, user_message):
    """
    Stream responses using the Teams AI library's StreamingResponse class with improved run handling
    
    Args:
        turn_context: The TurnContext object
        state: The conversation state
        user_message: The user's message (can be None for follow-up processing)
    """
    try:
        client = create_client()
        thread_id = state["session_id"]
        assistant_id = state["assistant_id"]
        
        # Create a StreamingResponse instance from Teams AI
        streamer = StreamingResponse(turn_context)
        
        # Send initial informative update
        streamer.queue_informative_update("Processing your request...")
        
        # Track the run ID for proper cleanup
        run_id = None
        
        try:
            # First, add the user message to the thread if provided
            if user_message:
                # Always check for and cancel any active runs first
                try:
                    runs = client.beta.threads.runs.list(thread_id=thread_id, limit=1)
                    if runs.data:
                        for run in runs.data:
                            if run.status in ["in_progress", "queued", "requires_action"]:
                                logging.info(f"Cancelling active run {run.id} before processing new message")
                                try:
                                    client.beta.threads.runs.cancel(thread_id=thread_id, run_id=run.id)
                                    # Wait for cancellation to take effect
                                    cancel_wait_start = time.time()
                                    max_cancel_wait = 5  # Maximum seconds to wait
                                    
                                    # Poll until cancellation completes or times out
                                    while time.time() - cancel_wait_start < max_cancel_wait:
                                        await asyncio.sleep(1)
                                        try:
                                            status = client.beta.threads.runs.retrieve(
                                                thread_id=thread_id, 
                                                run_id=run.id
                                            )
                                            if status.status in ["cancelled", "completed", "failed", "expired"]:
                                                logging.info(f"Run {run.id} is now in state {status.status}")
                                                break
                                        except Exception:
                                            # If we can't retrieve the run, assume it's gone
                                            break
                                except Exception as cancel_e:
                                    logging.warning(f"Error cancelling run: {cancel_e}")
                                    # If cancellation fails, create a new thread
                                    new_thread = client.beta.threads.create()
                                    old_thread_id = thread_id
                                    thread_id = new_thread.id
                                    with conversation_states_lock:
                                        state["session_id"] = thread_id
                                    logging.info(f"Created new thread {thread_id} after cancel failure (replacing {old_thread_id})")
                except Exception as check_e:
                    logging.warning(f"Error checking for active runs: {check_e}")
                
                # Now add the message with retries
                max_retries = 3
                added = False
                
                for retry in range(max_retries):
                    try:
                        client.beta.threads.messages.create(
                            thread_id=thread_id,
                            role="user",
                            content=user_message
                        )
                        logging.info(f"Added user message to thread {thread_id} (attempt {retry+1})")
                        added = True
                        break
                    except Exception as add_e:
                        if "already has an active run" in str(add_e) and retry < max_retries - 1:
                            logging.warning(f"Thread busy on attempt {retry+1}, waiting before retry")
                            await asyncio.sleep(2 * (retry + 1))  # Exponential backoff
                        elif retry == max_retries - 1:
                            # Final attempt - create new thread
                            try:
                                new_thread = client.beta.threads.create()
                                old_thread_id = thread_id
                                thread_id = new_thread.id
                                with conversation_states_lock:
                                    state["session_id"] = thread_id
                                logging.info(f"Created new thread {thread_id} after message add failures (replacing {old_thread_id})")
                                
                                # Try adding to the new thread
                                client.beta.threads.messages.create(
                                    thread_id=thread_id,
                                    role="user",
                                    content=user_message
                                )
                                added = True
                                logging.info(f"Added message to new thread {thread_id}")
                            except Exception as new_thread_e:
                                logging.error(f"Error creating new thread for message: {new_thread_e}")
                                raise
                        else:
                            logging.error(f"Error adding message on attempt {retry+1}: {add_e}")
                
                if not added:
                    raise Exception("Failed to add message after multiple attempts")

            # Create a run with proper error handling
            try:
                # Send typing indicator
                await turn_context.send_activity(create_typing_activity())
                
                # Create run - with retry logic
                max_run_retries = 3
                run_created = False
                
                for run_retry in range(max_run_retries):
                    try:
                        # Create the run
                        run = client.beta.threads.runs.create(
                            thread_id=thread_id,
                            assistant_id=assistant_id
                        )
                        run_id = run.id
                        logging.info(f"Created run {run_id} for thread {thread_id} (attempt {run_retry+1})")
                        run_created = True
                        break
                    except Exception as run_e:
                        if "already has an active run" in str(run_e) and run_retry < max_run_retries - 1:
                            logging.warning(f"Thread has active run on attempt {run_retry+1}, waiting before retry")
                            await asyncio.sleep(2 * (run_retry + 1))
                        elif run_retry == max_run_retries - 1:
                            # Final attempt - create new thread
                            try:
                                new_thread = client.beta.threads.create()
                                old_thread_id = thread_id
                                thread_id = new_thread.id
                                with conversation_states_lock:
                                    state["session_id"] = thread_id
                                logging.info(f"Created new thread {thread_id} after run creation failures (replacing {old_thread_id})")
                                
                                # If we had a user message, add it to the new thread
                                if user_message:
                                    client.beta.threads.messages.create(
                                        thread_id=thread_id,
                                        role="user",
                                        content=user_message
                                    )
                                
                                # Now try creating a run on the new thread
                                run = client.beta.threads.runs.create(
                                    thread_id=thread_id,
                                    assistant_id=assistant_id
                                )
                                run_id = run.id
                                run_created = True
                                logging.info(f"Created run {run_id} on new thread {thread_id}")
                            except Exception as new_thread_run_e:
                                logging.error(f"Error creating run on new thread: {new_thread_run_e}")
                                raise
                        else:
                            logging.error(f"Error creating run on attempt {run_retry+1}: {run_e}")
                
                if not run_created:
                    raise Exception("Failed to create run after multiple attempts")
                        
            except Exception as run_create_e:
                logging.error(f"Error creating run: {run_create_e}")
                raise

            # Now handle the streaming with buffer management
            buffer = []
            last_chunk_time = time.time()
            completed = False
            
            try:
                # Poll for the run result instead of streaming to avoid race conditions
                # This is more reliable than streaming for Teams
                max_wait_time = 120  # Maximum wait time in seconds
                wait_interval = 2    # Check interval in seconds
                elapsed_time = 0
                
                # Send initial typing indicator
                await turn_context.send_activity(create_typing_activity())
                
                while elapsed_time < max_wait_time:
                    # Check run status
                    try:
                        run_status = client.beta.threads.runs.retrieve(
                            thread_id=thread_id,
                            run_id=run_id
                        )
                        
                        # Send typing indicator periodically
                        if elapsed_time % 8 == 0:
                            await turn_context.send_activity(create_typing_activity())
                        
                        # Check for completion
                        if run_status.status == "completed":
                            logging.info(f"Run {run_id} completed successfully")
                            completed = True
                            
                            # Get the complete message
                            messages = client.beta.threads.messages.list(
                                thread_id=thread_id,
                                order="desc",
                                limit=1
                            )
                            
                            if messages.data:
                                latest_message = messages.data[0]
                                message_text = ""
                                
                                for content_part in latest_message.content:
                                    if content_part.type == 'text':
                                        message_text += content_part.text.value
                                
                                # Queue the complete message
                                if message_text:
                                    streamer.queue_text_chunk(message_text)
                            
                            break
                            
                        # Check for failure states
                        elif run_status.status in ["failed", "cancelled", "expired"]:
                            logging.error(f"Run {run_id} ended with status: {run_status.status}")
                            streamer.queue_text_chunk(f"\n\nI encountered an issue while processing your request (status: {run_status.status}). Please try again.")
                            break
                            
                        # Check for partial results every 5 seconds during in_progress state
                        elif run_status.status == "in_progress" and elapsed_time % 5 == 0:
                            try:
                                messages = client.beta.threads.messages.list(
                                    thread_id=thread_id,
                                    order="desc",
                                    limit=1
                                )
                                
                                if messages.data and messages.data[0].role == "assistant":
                                    latest_message = messages.data[0]
                                    current_text = ""
                                    
                                    for content_part in latest_message.content:
                                        if content_part.type == 'text':
                                            current_text += content_part.text.value
                                    
                                    # Only update if we have new content since last check
                                    if current_text and current_text != "".join(buffer):
                                        buffer = [current_text]  # Replace buffer with current full text
                                        
                                        # Only send updates at reasonable intervals for Teams
                                        current_time = time.time()
                                        if current_time - last_chunk_time >= 1.5:  # Teams requires 1.5s between msgs
                                            streamer.queue_text_chunk(current_text)
                                            last_chunk_time = current_time
                            except Exception as check_e:
                                logging.warning(f"Error checking for partial messages: {check_e}")
                                # Continue - don't break the loop for this error
                    
                    except Exception as status_e:
                        logging.warning(f"Error checking run status: {status_e}")
                        # Continue polling despite the error
                    
                    # Wait before next check
                    await asyncio.sleep(wait_interval)
                    elapsed_time += wait_interval
                
                # If we timed out without completing
                if not completed and elapsed_time >= max_wait_time:
                    logging.warning(f"Timed out waiting for run {run_id} to complete")
                    
                    # Try to get whatever we have so far
                    try:
                        messages = client.beta.threads.messages.list(
                            thread_id=thread_id,
                            order="desc",
                            limit=1
                        )
                        
                        if messages.data and messages.data[0].role == "assistant":
                            current_text = ""
                            for content_part in messages.data[0].content:
                                if content_part.type == 'text':
                                    current_text += content_part.text.value
                            
                            if current_text:
                                streamer.queue_text_chunk("I'm taking longer than expected. Here's what I have so far:\n\n")
                                streamer.queue_text_chunk(current_text)
                            else:
                                streamer.queue_text_chunk("I'm taking longer than expected to generate a response. Please try again with a simpler request.")
                        else:
                            streamer.queue_text_chunk("I'm taking longer than expected to generate a response. Please try again.")
                    except Exception as timeout_e:
                        logging.error(f"Error retrieving partial message after timeout: {timeout_e}")
                        streamer.queue_text_chunk("I'm taking longer than expected to generate a response. Please try again.")
                
            except Exception as poll_e:
                logging.error(f"Error polling run: {poll_e}")
                streamer.queue_text_chunk("I encountered an error while generating a response. Please try again.")
            
            # Enable feedback loop for the final message
            streamer.set_feedback_loop(True)
            streamer.set_generated_by_ai_label(True)
            
            # End the stream
            await streamer.end_stream()
            
        except Exception as inner_e:
            logging.error(f"Error in streaming process: {inner_e}")
            traceback.print_exc()
            
            try:
                # Try to end the stream gracefully
                streamer.queue_text_chunk("I'm sorry, I encountered an error while processing your request.")
                await streamer.end_stream()
            except Exception as end_error:
                logging.error(f"Failed to end stream properly: {end_error}")
                # At this point, just send a direct message
                await turn_context.send_activity("I encountered an error while processing your request. Please try again.")
        
        finally:
            # Always clean up active runs to prevent lingering state issues
            try:
                # Mark thread as no longer busy in the state
                with conversation_states_lock:
                    state["active_run"] = False
                
                # Remove from active runs tracking
                with active_runs_lock:
                    if thread_id in active_runs:
                        del active_runs[thread_id]
                try:
                    await turn_context.send_activity(create_typing_stop_activity())
                except Exception as typing_stop_error:
                    logging.error(f"Error stopping typing indicator: {typing_stop_error}")
                # Try to cancel the run if it's still active
                if run_id:
                    try:
                        client.beta.threads.runs.cancel(thread_id=thread_id, run_id=run_id)
                        logging.info(f"Cancelled run {run_id} during cleanup")
                    except Exception:
                        # Ignore cancellation errors during cleanup
                        pass
            except Exception as cleanup_e:
                logging.warning(f"Error during run cleanup: {cleanup_e}")
                
    except Exception as outer_e:
        logging.error(f"Outer error in stream_with_teams_ai: {str(outer_e)}")
        traceback.print_exc()
        
        # Send a user-friendly error message
        await turn_context.send_activity("I encountered a problem while processing your request. Please try again or start a new chat.")
        
        # Try a fallback response
        await send_fallback_response(turn_context, user_message or "How can I help you?")

async def stream_with_custom_implementation(turn_context: TurnContext, state, user_message):
    """
    Use a custom streaming implementation when Teams AI library is not available
    
    Args:
        turn_context: The TurnContext object
        state: The conversation state
        user_message: The user's message
    """
    try:
        client = create_client()
        thread_id = state["session_id"]
        assistant_id = state["assistant_id"]
        
        # Create our custom streaming response handler
        streamer = TeamsStreamingResponse(turn_context)
        
        # Send initial typing indicator
        await streamer.send_typing_indicator()
        
        try:
            # First, add the user message to the thread
            if user_message:
                try:
                    # Check for any existing active runs first
                    runs = client.beta.threads.runs.list(thread_id=thread_id, limit=1)
                    active_run_found = False
                    
                    if runs.data:
                        for run in runs.data:
                            if run.status in ["in_progress", "queued", "requires_action"]:
                                active_run_found = True
                                # Try to cancel it
                                try:
                                    client.beta.threads.runs.cancel(thread_id=thread_id, run_id=run.id)
                                    logging.info(f"Requested cancellation of run {run.id}")
                                    await asyncio.sleep(2)  # Wait for cancellation to take effect
                                except Exception as cancel_e:
                                    logging.warning(f"Failed to cancel run {run.id}: {cancel_e}")
                    
                    # If we had an active run and couldn't cancel it, create a new thread
                    if active_run_found:
                        try:
                            # Create a new thread instead of trying to use the busy one
                            new_thread = client.beta.threads.create()
                            thread_id = new_thread.id
                            with conversation_states_lock:
                                state["session_id"] = thread_id
                            logging.info(f"Created new thread {thread_id} to avoid active run conflicts")
                        except Exception as thread_create_e:
                            logging.error(f"Failed to create new thread: {thread_create_e}")
                    
                    # Add the message (with retries)
                    message_added = False
                    max_retries = 3
                    
                    for retry in range(max_retries):
                        try:
                            client.beta.threads.messages.create(
                                thread_id=thread_id,
                                role="user",
                                content=user_message
                            )
                            message_added = True
                            break
                        except Exception as add_e:
                            if retry < max_retries - 1:
                                await asyncio.sleep(2)
                                logging.warning(f"Retrying message add after error: {add_e}")
                            else:
                                raise
                    
                    if not message_added:
                        raise Exception("Failed to add message after multiple attempts")
                
                except Exception as msg_e:
                    logging.error(f"Failed to add message to thread: {msg_e}")
                    # Create a new thread and try again
                    try:
                        new_thread = client.beta.threads.create()
                        thread_id = new_thread.id
                        with conversation_states_lock:
                            state["session_id"] = thread_id
                        
                        client.beta.threads.messages.create(
                            thread_id=thread_id,
                            role="user",
                            content=user_message
                        )
                        logging.info(f"Created new thread {thread_id} and added message after failure")
                    except Exception as recovery_e:
                        logging.error(f"Recovery attempt failed: {recovery_e}")
                        raise Exception("Could not create thread or add message")
            
            # Create a run to generate a response
            run = None
            try:
                # Create a run
                run = client.beta.threads.runs.create(
                    thread_id=thread_id,
                    assistant_id=assistant_id
                )
                run_id = run.id
                logging.info(f"Created run {run_id}")
            except Exception as run_e:
                logging.error(f"Error creating run: {run_e}")
                raise
            
            # Poll the run until completion
            accumulated_text = ""
            max_wait_time = 120  # Maximum wait time in seconds
            wait_interval = 2  # seconds
            elapsed_time = 0
            last_message_check_time = 0
            message_check_interval = 5  # Check for partial messages every 5 seconds
            
            while elapsed_time < max_wait_time:
                # Send a typing indicator
                if elapsed_time % 8 == 0:  # Send typing indicator every ~8 seconds
                    await streamer.send_typing_indicator()
                
                # Check run status
                try:
                    run_status = client.beta.threads.runs.retrieve(
                        thread_id=thread_id,
                        run_id=run_id
                    )
                    
                    # Check for partial messages if enough time has passed
                    current_time = time.time()
                    if (current_time - last_message_check_time) >= message_check_interval:
                        last_message_check_time = current_time
                        
                        # Get the latest messages
                        messages = client.beta.threads.messages.list(
                            thread_id=thread_id,
                            order="desc",
                            limit=1
                        )
                        
                        if messages.data and messages.data[0].role == "assistant":
                            message_text = ""
                            for content_part in messages.data[0].content:
                                if content_part.type == 'text':
                                    message_text += content_part.text.value
                            
                            # If we have new text, add it to our buffer
                            if message_text and message_text != accumulated_text:
                                # Get just the new part
                                new_text = message_text[len(accumulated_text):]
                                if new_text:
                                    # Queue this update
                                    await streamer.queue_update(new_text)
                                    accumulated_text = message_text
                    
                    # Handle completed run
                    if run_status.status == "completed":
                        # Get the final message
                        messages = client.beta.threads.messages.list(
                            thread_id=thread_id,
                            order="desc",
                            limit=1
                        )
                        
                        if messages.data:
                            message_text = ""
                            for content_part in messages.data[0].content:
                                if content_part.type == 'text':
                                    message_text += content_part.text.value
                            
                            # If there's new text we haven't sent yet
                            if message_text and message_text != accumulated_text:
                                new_text = message_text[len(accumulated_text):]
                                if new_text:
                                    # Add this to our stream buffer
                                    await streamer.queue_update(new_text)
                                    accumulated_text = message_text
                        
                        # Send the final complete message
                        await streamer.send_final_message()
                        return
                    
                    # Handle failed run
                    elif run_status.status in ["failed", "cancelled", "expired"]:
                        logging.error(f"Run ended with status: {run_status.status}")
                        await turn_context.send_activity(f"I'm sorry, I encountered an issue while processing your request (status: {run_status.status}). Please try again.")
                        return
                
                except Exception as poll_e:
                    logging.error(f"Error polling run status: {poll_e}")
                    # Continue polling - don't break the loop for transient errors
                
                # Wait before checking again
                await asyncio.sleep(wait_interval)
                elapsed_time += wait_interval
            
            # If we get here, we timed out
            logging.warning(f"Timed out waiting for run {run_id} to complete")
            await turn_context.send_activity("I'm sorry, it's taking longer than expected to process your request. Here's what I have so far:")
            
            # Send whatever we've accumulated
            if accumulated_text:
                await turn_context.send_activity(accumulated_text)
            else:
                await turn_context.send_activity("I couldn't generate a response. Please try again or ask in a different way.")
        
        except Exception as e:
            logging.error(f"Error in custom streaming: {e}")
            traceback.print_exc()
            await turn_context.send_activity("I encountered an error while processing your request. Please try again.")
            await turn_context.send_activity(create_typing_stop_activity())
            # Try a fallback direct completion
            await send_fallback_response(turn_context, user_message)
    
    except Exception as outer_e:
        logging.error(f"Outer error in stream_with_custom_implementation: {str(outer_e)}")
        traceback.print_exc()
        await turn_context.send_activity("I'm experiencing technical difficulties. Please try again later.")

async def poll_for_message(client, thread_id, streamer):
    """
    Poll for messages and send any updates via the streamer.
    Used as a fallback when streaming fails.
    """
    try:
        # Get the latest message
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
            
            # Queue the complete message if we have it
            if message_text:
                streamer.queue_text_chunk(message_text)
                return
        
        streamer.queue_text_chunk("I processed your request but couldn't generate a proper response. Please try again.")
        
    except Exception as e:
        logging.error(f"Error in poll_for_message: {e}")
        streamer.queue_text_chunk("I encountered an error while retrieving the response. Please try again.")

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
            model="gpt-4.1-mini",  # Ensure this model supports vision
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
        elif processing_method == "thread_attachment":
            awareness_message += "This file has been attached to the thread and can be accessed via file search."
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
                "creation_time": time.time(),
                "last_activity_time": time.time()
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
                    "creation_time": time.time(),
                    "last_activity_time": time.time()
                }
                state = conversation_states[conversation_id]
                
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Log initialization attempt with user details for traceability
        logger.info(f"Initializing chat for user {user_id} in conversation {conversation_id} with context: {context}")
        
        # ALWAYS create a new assistant and thread for this user - never reuse
        client = create_client()
        
        # Create a vector store
        try:
            vector_store = client.vector_stores.create(
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
                model="gpt-4.1-mini",
                instructions=SYSTEM_PROMPT,
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
        await turn_context.send_activity("Hi! I'm the AI Assistant here to help you with your tasks.")
        
        if context:
            #await turn_context.send_activity(f"I've initialized with your context: '{context}'")
            # Also send the first response
            await send_message(turn_context, state)
            
    except Exception as e:
        await turn_context.send_activity(f"Error initializing chat: {str(e)}")
        logger.error(f"Error in initialize_chat for user {user_id}: {str(e)}")
        traceback.print_exc()
        
        # Try a fallback response if everything else fails
        try:
            await send_fallback_response(turn_context, context or "How can I help you with product management today?")
        except Exception as fallback_e:
            logging.error(f"Even fallback failed during initialization: {fallback_e}")

# Send a message without user input (used after file upload or initialization)
async def send_message(turn_context: TurnContext, state):
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Use streaming if supported by the channel
        supports_streaming = turn_context.activity.channel_id == "msteams"
        
        if TEAMS_AI_AVAILABLE and supports_streaming:
            # Use streaming for response
            await stream_with_teams_ai(turn_context, state, None)
        elif supports_streaming:
            # Use custom streaming implementation
            await stream_with_custom_implementation(turn_context, state, None)
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
        
        # Use fallback if everything else fails
        try:
            await send_fallback_response(turn_context, "Hello, how can I help you with product management today?")
        except:
            pass  # Last resort is to simply give up

# Send welcome message when bot is added

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
    This function handles both streaming and non-streaming modes.
    """
    try:
        # Create defaults if not provided
        if not assistant:
            logging.warning(f"No assistant ID provided, creating a default one.")
            try:
                assistant_obj = client.beta.assistants.create(
                    name="default_conversation_assistant",
                    model="gpt-4.1-mini",
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
                    model="gpt-4.1-mini",
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
                """
                Enhanced async generator for streaming assistant responses
                with improved handling of events, tool calls, and error recovery.
                """
                buffer = []
                run_id = None
                completed = False
                tool_outputs_submitted = False
                wait_for_final_response = False
                latest_message_id = None
                last_yield_time = time.time()
                
                try:
                    # Get most recent message ID before run
                    try:
                        pre_run_messages = client.beta.threads.messages.list(
                            thread_id=session,
                            order="desc",
                            limit=1
                        )
                        if pre_run_messages and pre_run_messages.data:
                            latest_message_id = pre_run_messages.data[0].id
                            logging.info(f"Latest message before run: {latest_message_id}")
                    except Exception as e:
                        logging.warning(f"Could not get latest message before run: {e}")
                    
                    # Try to use the streaming API directly if available
                    try:
                        # Check if stream method is available
                        if hasattr(client.beta.threads.runs, 'stream'):
                            logging.info("Using beta.threads.runs.stream() for streaming")
                            
                            with client.beta.threads.runs.stream(
                                thread_id=session,
                                assistant_id=assistant,
                            ) as stream:
                                for event in stream:
                                    # Store run ID
                                    if hasattr(event, 'data') and hasattr(event.data, 'id') and event.event == "thread.run.created":
                                        run_id = event.data.id
                                        logging.info(f"Created streaming run {run_id}")
                                        
                                    # Handle message creation events
                                    if event.event == "thread.message.created":
                                        logging.info(f"New message created: {event.data.id}")
                                        if tool_outputs_submitted and event.data.id != latest_message_id:
                                            wait_for_final_response = True
                                            latest_message_id = event.data.id
                                            
                                    # Handle text deltas (the actual content streaming)
                                    if event.event == "thread.message.delta":
                                        delta = event.data.delta
                                        if delta.content:
                                            for content_part in delta.content:
                                                if content_part.type == 'text' and content_part.text:
                                                    text_value = content_part.text.value
                                                    if text_value:
                                                        # Add to buffer
                                                        buffer.append(text_value)
                                                        
                                                        # Yield chunks either when buffer gets large enough
                                                        # or when enough time has passed since last yield
                                                        current_time = time.time()
                                                        if len(buffer) >= 3 or (current_time - last_yield_time >= 0.5 and buffer):
                                                            joined_text = ''.join(buffer)
                                                            yield joined_text
                                                            buffer = []
                                                            last_yield_time = current_time
                                    
                                    # Handle run completion
                                    if event.event == "thread.run.completed":
                                        logging.info(f"Run completed: {event.data.id}")
                                        completed = True
                                        
                                        # Yield any remaining text
                                        if buffer:
                                            joined_text = ''.join(buffer)
                                            yield joined_text
                                            buffer = []
                                    
                                    # Handle tool calls
                                    elif event.event == "thread.run.requires_action":
                                        if event.data.required_action.type == "submit_tool_outputs":
                                            tool_calls = event.data.required_action.submit_tool_outputs.tool_calls
                                            
                                            # For now, just log and send a message about processing
                                            tool_call_message = "\n[Processing additional actions...]\n"
                                            yield tool_call_message
                                            
                                            logging.info(f"Run requires action: {len(tool_calls)} tool calls")
                                            
                                            # Create empty outputs array - in future this would handle actual tool calls
                                            tool_outputs = []
                                            for tool_call in tool_calls:
                                                # Log tool call for debugging
                                                logging.info(f"Tool call: {tool_call.function.name} - {tool_call.function.arguments[:100]}...")
                                                
                                                # Add a placeholder result
                                                tool_outputs.append({
                                                    "tool_call_id": tool_call.id,
                                                    "output": "This function is not yet implemented in this version."
                                                })
                                            
                                            # Submit the (placeholder) outputs
                                            try:
                                                client.beta.threads.runs.submit_tool_outputs(
                                                    thread_id=session,
                                                    run_id=event.data.id,
                                                    tool_outputs=tool_outputs
                                                )
                                                tool_outputs_submitted = True
                                                logging.info(f"Submitted tool outputs for run {event.data.id}")
                                                yield "\n[Continuing with response...]\n"
                                            except Exception as submit_e:
                                                logging.error(f"Error submitting tool outputs: {submit_e}")
                                                yield f"\n[Error handling tools: {str(submit_e)}]\n"
                                
                                # Yield any remaining buffer at the end
                                if buffer:
                                    joined_text = ''.join(buffer)
                                    yield joined_text
                                    buffer = []
                                
                                # Exit if completed
                                if completed:
                                    return
                        else:
                            raise NotImplementedError("Stream method not available")
                            
                    except (NotImplementedError, AttributeError) as stream_not_available:
                        # Fallback to iter_chunks if stream is not available
                        logging.info(f"Direct streaming not available: {stream_not_available}. Falling back to iter_chunks")
                        
                        # Create run with stream=True for iter_chunks approach
                        run = client.beta.threads.runs.create(
                            thread_id=session,
                            assistant_id=assistant,
                            stream=True
                        )
                        
                        run_id = run.id
                        
                        # Use iter_chunks if available
                        if hasattr(run, "iter_chunks"):
                            logging.info(f"Using iter_chunks() for streaming run {run_id}")
                            
                            for chunk in run.iter_chunks():
                                text_piece = ""
                                
                                if hasattr(chunk, "data") and hasattr(chunk.data, "delta"):
                                    delta = chunk.data.delta
                                    if hasattr(delta, "content") and delta.content:
                                        for content in delta.content:
                                            if content.type == "text" and hasattr(content.text, "value"):
                                                text_piece = content.text.value
                                                
                                if text_piece:
                                    # Add to buffer
                                    buffer.append(text_piece)
                                    
                                    # Yield chunks periodically
                                    current_time = time.time()
                                    if len(buffer) >= 3 or (current_time - last_yield_time >= 0.5 and buffer):
                                        joined_text = ''.join(buffer)
                                        yield joined_text
                                        buffer = []
                                        last_yield_time = current_time
                                
                                # Small delay to work with asyncio
                                await asyncio.sleep(0.01)
                            
                            # Yield any remaining text
                            if buffer:
                                joined_text = ''.join(buffer)
                                yield joined_text
                                buffer = []
                                    
                        # Fallback to events iterator
                        elif hasattr(run, "events"):
                            logging.info(f"Using events iterator for streaming run {run_id}")
                            
                            for event in run.events:
                                if event.event == "thread.message.delta":
                                    if hasattr(event.data, "delta") and hasattr(event.data.delta, "content"):
                                        for content in event.data.delta.content:
                                            if content.type == "text" and hasattr(content.text, "value"):
                                                # Add to buffer
                                                buffer.append(content.text.value)
                                                
                                                # Yield chunks periodically
                                                current_time = time.time()
                                                if len(buffer) >= 3 or (current_time - last_yield_time >= 0.5 and buffer):
                                                    joined_text = ''.join(buffer)
                                                    yield joined_text
                                                    buffer = []
                                                    last_yield_time = current_time
                                                
                                                # Small delay
                                                await asyncio.sleep(0.01)
                            
                            # Yield any remaining text
                            if buffer:
                                joined_text = ''.join(buffer)
                                yield joined_text
                                buffer = []
                        
                        # Final fallback to polling
                        else:
                            logging.info(f"Falling back to polling for streaming run {run_id}")
                            yield "Processing your request...\n"
                            
                            max_wait_time = 90  # seconds
                            wait_interval = 2   # seconds
                            elapsed_time = 0
                            last_status = None
                            
                            while elapsed_time < max_wait_time:
                                try:
                                    run_status = client.beta.threads.runs.retrieve(
                                        thread_id=session, 
                                        run_id=run_id
                                    )
                                    
                                    # Only log when status changes
                                    if last_status != run_status.status:
                                        logging.info(f"Run {run_id} status: {run_status.status}")
                                        last_status = run_status.status
                                    
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
                                            message_text = ""
                                            
                                            for content_part in latest_message.content:
                                                if content_part.type == 'text':
                                                    message_text += content_part.text.value
                                            
                                            # Split long responses into chunks for better streaming
                                            if len(message_text) > 500:
                                                # Use sentence-aware chunking if possible
                                                sentences = message_text.split('. ')
                                                current_chunk = ""
                                                
                                                for sentence in sentences:
                                                    current_chunk += sentence + '. '
                                                    
                                                    if len(current_chunk) >= 200:
                                                        yield current_chunk
                                                        current_chunk = ""
                                                        await asyncio.sleep(0.05)  # Small delay between chunks
                                                
                                                # Yield any remaining text
                                                if current_chunk:
                                                    yield current_chunk
                                            else:
                                                # For shorter responses, just yield the whole thing
                                                yield message_text
                                        break
                                    
                                    elif run_status.status in ["failed", "cancelled", "expired"]:
                                        yield f"\nError: Run ended with status {run_status.status}. Please try again."
                                        break
                                    
                                    elif run_status.status == "requires_action":
                                        yield "\n[Run requires additional actions which cannot be handled in polling mode.]\n"
                                        # Try to cancel the run since we can't handle actions in polling mode
                                        try:
                                            client.beta.threads.runs.cancel(
                                                thread_id=session,
                                                run_id=run_id
                                            )
                                            logging.info(f"Cancelled run {run_id} that required actions in polling mode")
                                        except Exception as cancel_e:
                                            logging.error(f"Failed to cancel run requiring actions: {cancel_e}")
                                        break
                                    
                                    yield "."  # Show progress
                                    await asyncio.sleep(wait_interval)
                                    elapsed_time += wait_interval
                                    
                                except Exception as poll_e:
                                    logging.error(f"Error polling run status: {poll_e}")
                                    yield "E"  # Show error in progress
                                    await asyncio.sleep(wait_interval)
                                    elapsed_time += wait_interval
                            
                            if elapsed_time >= max_wait_time:
                                yield "\nResponse timed out. Please try again with a simpler request."
                
                except Exception as e:
                    error_details = traceback.format_exc()
                    logging.error(f"Error in streaming generation: {e}\n{error_details}")
                    yield f"\n[ERROR] An error occurred while generating the response: {str(e)}. Please try again.\n"
            # Return streaming generator
            return async_generator()
            # async def async_generator():
            #     try:
            #         # Create run with stream=True
            #         run = client.beta.threads.runs.create(
            #             thread_id=session,
            #             assistant_id=assistant,
            #             stream=True
            #         )
                    
            #         # Handle the stream based on available methods
            #         if hasattr(run, "iter_chunks"):
            #             # Using iter_chunks synchronous iterator
            #             logging.info("Using iter_chunks() for API streaming")
            #             for chunk in run.iter_chunks():
            #                 text_piece = ""
                            
            #                 if hasattr(chunk, "data") and hasattr(chunk.data, "delta"):
            #                     delta = chunk.data.delta
            #                     if hasattr(delta, "content") and delta.content:
            #                         for content in delta.content:
            #                             if content.type == "text" and hasattr(content.text, "value"):
            #                                 text_piece = content.text.value
                                            
            #                 if text_piece:
            #                     yield text_piece
            #                     # Small delay to make it work with asyncio
            #                     await asyncio.sleep(0.01)
                                
            #         elif hasattr(run, "events"):
            #             # Using events iterator
            #             logging.info("Using events iterator for API streaming")
            #             for event in run.events:
            #                 if event.event == "thread.message.delta":
            #                     if hasattr(event.data, "delta") and hasattr(event.data.delta, "content"):
            #                         for content in event.data.delta.content:
            #                             if content.type == "text" and hasattr(content.text, "value"):
            #                                 yield content.text.value
            #                                 await asyncio.sleep(0.01)
            #         else:
            #             # Fallback to polling
            #             logging.info("Using fallback polling for API streaming")
            #             yield "Processing your request...\n"
                        
            #             run_id = run.id
            #             max_wait_time = 90  # seconds
            #             wait_interval = 2   # seconds
            #             elapsed_time = 0
                        
            #             while elapsed_time < max_wait_time:
            #                 run_status = client.beta.threads.runs.retrieve(
            #                     thread_id=session, 
            #                     run_id=run_id
            #                 )
                            
            #                 if run_status.status == "completed":
            #                     yield "\n"  # Clear the progress line
                                
            #                     # Get the complete message
            #                     messages = client.beta.threads.messages.list(
            #                         thread_id=session,
            #                         order="desc",
            #                         limit=1
            #                     )
                                
            #                     if messages.data:
            #                         latest_message = messages.data[0]
            #                         for content_part in latest_message.content:
            #                             if content_part.type == 'text':
            #                                 yield content_part.text.value
            #                     break
                            
            #                 elif run_status.status in ["failed", "cancelled", "expired"]:
            #                     yield f"\nError: Run ended with status {run_status.status}. Please try again."
            #                     break
                            
            #                 yield "."  # Show progress
            #                 await asyncio.sleep(wait_interval)
            #                 elapsed_time += wait_interval
                        
            #             if elapsed_time >= max_wait_time:
            #                 yield "\nResponse timed out. Please try again."
                
            #     except Exception as e:
            #         logging.error(f"Error in streaming generation: {e}")
            #         yield f"\n[ERROR] An error occurred while generating the response: {str(e)}. Please try again.\n"
            
            # # Return streaming generator
            # return async_generator()
        
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
