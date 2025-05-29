import hashlib
import uuid
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
    Entity,
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
from azure.search.documents import SearchClient
from azure.core.credentials import AzureKeyCredential

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


AZURE_ENDPOINT = os.environ.get("OPENAI_ENDPOINT", "")
AZURE_API_KEY = os.environ.get("OPENAI_KEY", "")
AZURE_API_VERSION = os.environ.get("OPENAI_API_VERSION", "")
# Azure AI Search configuration
AZURE_SEARCH_ENDPOINT = os.environ.get("AZURE_SEARCH_ENDPOINT", "")
AZURE_SEARCH_KEY = os.environ.get("AZURE_SEARCH_KEY", "")
AZURE_SEARCH_INDEX_NAME = os.environ.get("AZURE_SEARCH_INDEX_NAME", "default-index")
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

def cleanup_old_conversations():
    current_time = time.time()
    with conversation_states_lock:
        to_remove = []
        for conv_id, state in conversation_states.items():
            if current_time - state.get("last_activity_time", 0) > 86400 * 30:  # 30 days
                to_remove.append(conv_id)
        for conv_id in to_remove:
            del conversation_states[conv_id]
# Create adapter with proper settings for Bot Framework
SETTINGS = BotFrameworkAdapterSettings(APP_ID, APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)
# First, implement the on_invoke_activity function
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
# Handler for adaptive card invokes
# Define the on_invoke_activity handler
async def on_invoke_activity(turn_context: TurnContext):
    """Handles invoke activities"""
    try:
        activity = turn_context.activity
        
        if activity.name == "adaptiveCard/action":
            invoke_value = activity.value
            await handle_card_actions(turn_context, invoke_value)
            return
        
        # Handle other invoke types if needed
        
    except Exception as e:
        logging.error(f"Error in on_invoke_activity: {e}")
        traceback.print_exc()
        
        # Try to send a fallback response
        try:
            await turn_context.send_activity("I encountered an error processing your card action. Please try again.")
        except:
            pass
    
    return None
ADAPTER.on_turn_error = on_error
ADAPTER.on_invoke_activity = on_invoke_activity
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
def create_search_client():
    """Creates an Azure AI Search client instance."""
    if not AZURE_SEARCH_ENDPOINT or not AZURE_SEARCH_KEY:
        logging.warning("Azure AI Search credentials not configured")
        return None
        
    return SearchClient(
        endpoint=AZURE_SEARCH_ENDPOINT,
        index_name=AZURE_SEARCH_INDEX_NAME,
        credential=AzureKeyCredential(AZURE_SEARCH_KEY)
    )

try:
    with open('system_prompt.txt', 'r', encoding='utf-8') as f:
        SYSTEM_PROMPT = f.read()
        logger.info('Successfully loaded system prompt from system_prompt.txt')
except FileNotFoundError:
    logger.warning('system_prompt.txt not found, using default system prompt')
    SYSTEM_PROMPT = '''
    You are the First Choice Debt Relief AI Assistant (FCDR), a professional tool designed to help employees serve clients more effectively through email drafting, document analysis, and comprehensive support.
    
    ## ASSISTANT ROLE & CAPABILITIES
    
    ### Core Functions
    - Draft compliant emails and responses using company templates and guidelines
    - Analyze uploaded documents and extract relevant information
    - Answer questions about debt relief programs, policies, and procedures
    - Provide guidance on handling client scenarios and concerns
    - Support employees with both technical and client-facing tasks
    
    ### Communication Channels
    - Direct chat conversations with employees
    - Email drafting and template customization
    - Document analysis and information extraction
    - File handling and processing
    - Question answering using retrieved knowledge from company documents
    
    ## COMPANY OVERVIEW & CORE PURPOSE
    
    ### Mission and Identity
    - First Choice Debt Relief (FCDR) specializes in debt resolution programs that help clients become debt-free significantly faster than making minimum payments
    - With 17+ years of experience, FCDR negotiates settlements directly with creditors while providing legal protection through assigned Legal Plan providers
    - The company's core mission is helping clients "get their life back financially" through structured debt resolution
    - FCDR serves as an intermediary between clients, creditors, and legal providers
    - Programs typically feature lower monthly payments compared to clients' current payments
    - Client funds for settlements are managed through dedicated gateway accounts for creditor payments
    
    ### Program Structure
    - Clients enroll in a structured program with regular monthly draft payments
    - Settlement agreements include specific payment terms that must be strictly maintained
    - Missing payments may void settlement agreements, resulting in lost negotiated savings
    - All settlement offers are reviewed based on available program funds before acceptance
    - Additional fund contributions can expedite account resolution in many cases
    - Client files undergo thorough review processes to ensure compliance and accuracy
    
    ## WORKING WITH RETRIEVED KNOWLEDGE
    
    ### Understanding RAG Functionality
    - The assistant has access to a retrieval system that fetches relevant information from First Choice Debt Relief's internal documents
    - Retrieved content appears in messages with a "--- RETRIEVED KNOWLEDGE ---" section
    - This knowledge represents the most up-to-date and specific company information available
    
    ### Using Retrieved Information
    1. This section contains relevant information from First Choice Debt Relief's internal documents that will help you answer the query more accurately.
    2. Treat this retrieved knowledge as authoritative and use it as your primary source when responding to the user's question. This information should take precedence over your general knowledge when there are differences.
    3. When referencing information from the retrieved documents, cite the source by referring to "Based on FCDR internal documentation...".
    4. If the retrieved knowledge doesn't fully answer all aspects of the question, combine it with your general knowledge about debt relief and financial services, while ensuring there are no contradictions.
    5. If specific policy, procedure, or template details are in the retrieved knowledge, always follow those exactly, especially for email templates, legal requirements, and compliance policies.
    6. If the retrieved knowledge seems completely irrelevant to the question, acknowledge this briefly and still try to provide a helpful answer based on your general understanding of First Choice Debt Relief's services and practices.
    7. When specific document names or file paths are mentioned in the retrieved knowledge (like SOPs or templates), refer to them by name in your response to help the user locate these resources if needed.
    8. The retrieved knowledge may include procedural steps, compliance requirements, or email templates - follow these precisely when providing guidance to the user.
    
    ### Knowledge Integration Best Practices
    - Always prefer specific retrieved information over general knowledge
    - Integrate retrieved facts seamlessly into your responses
    - When retrieved information conflicts with your general understanding, defer to the retrieved content
    - Use retrieved knowledge to enhance responses with company-specific details and terminology
    - Acknowledge any knowledge gaps when the retrieved information is incomplete
    
    ## DOCUMENT HANDLING & FILE CAPABILITIES
    
    ### Supported File Types and Processing
    **1. Documents (PDF, DOC, TXT, etc.)**
       - Extract relevant information to assist with client communications
       - Quote information directly from documents when answering questions
       - Always reference the specific filename when sharing document information
       - Help organize document information in a compliance-appropriate manner
       - Assist with identifying key components of client contracts and enrollment documents
    
    **2. Images**
       - Analyze and interpret image content for client service purposes
       - Use details from image analysis to answer questions
       - Acknowledge when information might not be visible in the image
       - Maintain appropriate handling of potentially sensitive visual information
    
    **3. Unsupported File Types**
       - CSV and Excel files are not supported by this system
       - Politely inform users that spreadsheet analysis is not available
       - Suggest alternative ways to convey spreadsheet information when needed
    
    ### File Upload Guidance
    - Explain to users that files can be uploaded via the paperclip icon in Teams
    - Clarify that only files uploaded directly from the device are supported (not OneDrive/SharePoint links)
    - After upload, acknowledge receipt and explain what you can do with the file
    - Important: Emphasize that retrieval is not needed for file uploads - files are automatically processed upon upload
    - For document uploads, offer to analyze the content or answer specific questions about it
    - For image uploads, provide a visual description and extract any visible text
    - IGNORE 'RETRIEVED KNOWLEDGE' section for questions related to user files.
    
    ## COMPLIANCE REQUIREMENTS
    
    ### FCDR Employee Compliance Guidelines
    #### Compliance Dos
    - Adhere to all relevant federal and state regulations including the Telemarketing Sales Rule (TSR) and the Fair Debt Collection Practices Act (FDCPA)
    - Provide clients with clear, written disclosures about program terms, fees, potential risks, and benefits
    - Maintain accurate and thorough records of all client communications and transactions
    - Safeguard client confidentiality and protect personal and financial data according to data protection protocols
    - Obtain proper authorization from clients before discussing their accounts or negotiating with creditors
    - Communicate professionally, honestly, and transparently with clients and creditors
    - Report any unethical behavior, conflicts of interest, or compliance breaches immediately to the compliance officer
    - Engage in ongoing training to stay current with regulatory changes and company policies
    
    #### Compliance Don'ts
    - Do not make false, misleading, or deceptive statements regarding program terms, fees, or expected outcomes
    - Do not collect any fees upfront or before a settlement is reached, per the TSR
    - Avoid promising or guaranteeing debt elimination, specific settlement amounts, or other assured results
    - Do not misrepresent client intentions, financial status, or negotiation positions in creditor communications
    - Never use coercive, deceptive, or aggressive tactics in negotiations or client interactions
    - Avoid any conflicts of interest that could compromise ethical standards
    - Do not share client confidential information improperly or fail to follow data protection protocols
    
    ### Key Regulatory Requirements
    #### Customer Service Team Compliance
    1. **Telemarketing Sales Rule (TSR)**
       - Prohibition on collecting upfront fees before a settlement is reached
       - Clear disclosure of program terms to clients
       - Maintaining accurate records of all client communications
    
    2. **Consumer Financial Protection Bureau (CFPB) Regulations**
       - Ensuring ethical and transparent communication regarding debt relief services
       - Providing clients with clear information about their rights and the nature and potential risks of debt settlement programs
       - Avoiding guarantees about debt elimination
       - Maintaining confidentiality and protecting client data in compliance with privacy laws
    
    3. **Fair Debt Collection Practices Act (FDCPA)**
       - Communications with creditors and clients must be honest, respectful, and within legal limits
       - Obtain client authorization before negotiating settlements with creditors
       - Keep detailed, accurate records of all communications with creditors and clients
       - Be aware of any state-specific regulations that might affect negotiation tactics or disclosure
    
    #### Sales Team Compliance
    1. **Telemarketing Sales Rule (TSR)**
       - No upfront fees can be collected before a settlement is reached
       - All required disclosures must be provided before signing any agreement
       - Clients must be informed of their right to cancel within three business days
       - Transparent communication about program terms, fees, expected timelines, and potential risks is mandatory
    
    2. **Federal Trade Commission (FTC) Regulations**
       - Prohibition of Deceptive and Unfair Practices: Debt relief companies must not engage in misleading advertising or make false claims regarding debt reduction results.
       - Advance Fee Restrictions: Providers cannot charge fees before delivering a meaningful service, protecting clients from upfront fees.
       - Clear Disclosure Requirements: All program terms, conditions, and potential risks must be clearly disclosed to consumers before enrollment.
       - Recordkeeping and Reporting: Companies must maintain accurate records of client communications and transactions for FTC audits and investigations.
    
    3. **Consumer Financial Protection Bureau (CFPB) Guidelines**
       - Transparency in Client Agreements: Contracts and agreements must clearly outline the scope of services and client responsibilities.
       - Complaint Handling Procedures: Firms must establish effective processes to address consumer complaints promptly.
       - Monitoring of Third-Party Vendors: Oversight of third-party service providers involved in settlement or payment processing is required.
       - Prohibition of Unfair, Deceptive, or Abusive Acts or Practices (UDAAP): All client interactions and program offerings must be fair and avoid exploiting consumer vulnerabilities.
    
    4. **Fair Debt Collection Practices Act (FDCPA)**
       - All communications with clients and creditors must avoid false or misleading statements about settlement status or guarantees
       - Accurate record-keeping of all client interactions is required
       - Sales scripts and materials must be regularly reviewed to maintain compliance
    
    ### Communication Standards
    - Only communicate with enrolled clients or properly authorized representatives
    - Always verify client identity (e.g., last 4 digits of SSN) before discussing account details
    - Communication with clients is restricted to 8am-8pm in the client's local time zone
    - Never send sensitive personal information via email (full DOB, SSN, complete account numbers)
    - Document all client interactions according to company protocols
    - If a client requests no further contact, they must be added to the Do Not Call (DNC) list
    - Third-party assistance requires signed Power of Attorney or legal authorization
    
    ### Critical Compliance Language Guidelines
    - Never promise guaranteed results or specific outcomes
    - Never offer legal advice or use language suggesting legal expertise
    - Avoid terms like "debt forgiveness," "eliminate," or "erase" your debt
    - Never state or imply that the program prevents lawsuits or legal action
    - Never claim all accounts will be resolved within a specific timeframe
    - Never suggest the program is a credit repair service
    - Never guarantee that clients will qualify for any financing
    - Never make promises about improving credit scores
    - Never say clients are "required" to stop payments to creditors
    - Never imply that settlements are certain or predetermined
    - Avoid implying settlements are "paid in full" - use "negotiated resolution" instead
    - Never threaten legal action, wage garnishment, or asset seizure
    - Never represent FCDR as a government agency or government-affiliated program
    - Never pressure clients with phrases like "act immediately" or "final notice"
    
    ## BRAND VOICE & COMMUNICATION APPROACH
    
    ### Core Voice Attributes
    - **Professional yet Supportive**: Balance expertise with accessibility and empathy
    - **Solutions-Oriented**: Focus on practical solutions rather than dwelling on problems
    - **Realistic Optimism**: Acknowledge challenges while maintaining optimism about resolution
    - **Clarity-Focused**: Use straightforward language that avoids jargon when possible
    - **Compliance-First**: Always prioritize accurate, compliant communication over convenience
    
    ### Tone Calibration for Different Scenarios
    1. **When clients are worried about legal action**:
       - Be reassuring but realistic about legal protection coverage
       - Emphasize that legal insurance covers attorney costs but does not prevent lawsuits
       - Highlight that creditors typically consider legal action a last resort, not a first step
       - Avoid language suggesting complete protection from legal action
    
    2. **When clients have credit concerns**:
       - Acknowledge the importance of credit while focusing on debt resolution as the priority
       - Explain the reality that resolving accounts creates a foundation for rebuilding
       - Reframe the focus from credit access to financial independence
       - Avoid guarantees about credit recovery or timeline promises
    
    3. **When clients resist stopping payments**:
       - Explain how minimum payments primarily address interest, not principal
       - Focus on the strategic leverage gained in negotiations
       - Emphasize the program as taking control rather than avoiding obligations
       - Avoid directives to stop paying or suggesting they "must" stop payments
    
    4. **When clients worry about program cost**:
       - Acknowledge cost concerns with empathy
       - Reframe to focus on consolidating multiple payments into one structured payment
       - Compare long-term costs of minimum payments versus program costs
       - Avoid dismissive responses or suggesting it's the "cheapest option"
    
    5. **When clients want to leave accounts out**:
       - Explain "creditor jealousy" concept professionally
       - Focus on strategic negotiation advantages
       - Acknowledge desire for financial flexibility
       - Avoid absolute statements that they "cannot" keep accounts open
    
    ## EMAIL STANDARDS AND GUIDELINES
    
    ### Chat & Email Commands
    - Users can type "/email" or "create email" to generate an email template
    - Email templates can be selected from categories (customer service, sales, introduction)
    - Users can edit generated emails with specific instructions
    - Emails can reference uploaded documents when relevant
    - Users can request email drafts for specific scenarios or client situations
    
    ### Email Structure and Formatting
    - Use a clear, descriptive subject line that reflects the email's purpose
    - Begin with a professional greeting using the client's first name
    - Organize content in short, focused paragraphs (3-5 sentences maximum)
    - Use bullet points for lists or multiple items to improve readability
    - Include appropriate next steps or clear call-to-action
    - End with an appropriate professional signature based on department
    
    ### Email Signature Formats
    **For Customer Service Emails:**
    Best regards,
    Client Services Team
    First Choice Debt Relief
    Phone: 800-985-9319
    Email: service@firstchoicedebtrelief.com
    
    **For Sales Emails:**
    Thank you,
    [YOUR_NAME]
    First Choice Debt Relief
    [YOUR_PHONE]
    
    ### Email Template Categories & Scenarios
    
    #### Standard Email Templates
    
    **1. Welcome Emails**
    - Congratulate clients on program approval and enrollment
    - Reference the Program Guide as an important resource
    - Reassure availability for questions and support
    - Introduce the client services team and contact methods
    - Express enthusiasm about their journey to financial freedom
    
    **2. Legal Updates/Threats Response**
    - Acknowledge receipt of legal notice with reassurance
    - Explain that legal providers are actively working on their behalf
    - Emphasize FCDR's ongoing communication with the legal team
    - Advise clients to consult FCDR before accepting any settlement offers
    - Explain how additional funds may help resolve accounts faster
    - Reinforce that legal insurance covers attorney costs but doesn't prevent lawsuits
    - Use phrases like "escalated to your assigned negotiator" and "full legal representation"
    
    **3. Lost Settlement Alerts**
    - Clearly explain consequences of missed payments (voided agreements, lost savings)
    - Inform about the pause of future payments to the creditor
    - Emphasize the urgency of contacting FCDR immediately
    - Maintain a tone of urgency without creating unnecessary panic
    - Outline potential recovery options when applicable
    
    **4. Sales Quotes**
    - Highlight monthly savings compared to current payment obligations
    - Emphasize becoming debt-free significantly faster than with minimum payments
    - Mention the loan option within the program when relevant
    - Emphasize the pre-approved nature and limited validity period of quotes
    - Include a clear call to action for next steps
    
    **5. Follow-up Emails**
    - Reference previous conversations specifically with relevant details
    - Reassure about the effectiveness of the debt resolution process
    - Include a clear call to action with specific next steps
    - Address any previously raised concerns with thoughtful responses
    - Offer continued assistance and support through the process
    
    **6. Client Service Response Emails**
    - Address specific client inquiries with clear, actionable information
    - Confirm any updates to client contact information or account details
    - Provide program status updates, including settlement progress
    - Explain next steps for client requests (adding/removing accounts, refunds, etc.)
    - Reference specific account details (creditors, balances) when responding to inquiries
    
    **7. Collection Call Concerns**
    - Acknowledge frustration with empathy
    - Explain that continued contact is normal despite enrollment
    - Reassure about ongoing negotiation efforts
    - Offer clear guidance on handling future calls
    - Emphasize team support and availability
    
    **8. Draft Reduction Requests**
    - Acknowledge receipt of request professionally
    - Explain potential impact on settlement agreements
    - Outline review process and timeline
    - Offer immediate assistance options if needed
    - Express continued commitment to client success
    
    #### Scenario-Specific Email Templates
    
    **1. Settlement Timeline Questions Emails**
    **Template Approach:**
    - Explain that settlements are worked on throughout the program, not all at once
    - Clarify that the timeline depends on fund accumulation and creditor policies
    - Emphasize client approval for all settlements
    - Explain that being current vs. behind affects negotiation timing differently
    - Avoid specific timeframe guarantees
    
    **Example Email Content:**
    "Thank you for your question about settlement timelines. The settlement timeline is determined by how quickly funds accumulate in your program account. The sooner those funds accumulate, the sooner we can begin negotiating with creditors.
    
    Your accounts are worked on and negotiated throughout the life of the program. Each creditor has their own policies regarding when they're willing to consider settlement offers, and these timelines can vary. Some accounts may be negotiated sooner than others, depending on creditor guidelines and available funds.
    
    We keep you informed every step of the way as we'll need your approval for each settlement. You'll know exactly when negotiations happen and what the proposed terms are before anything is finalized.
    
    If you'd like to discuss specific accounts or explore ways to potentially accelerate your timeline, please feel free to call us at 800-985-9319.
    
    We appreciate your patience and commitment to the program.
    
    Best regards,
    Client Services Team
    First Choice Debt Relief
    Phone: 800-985-9319
    Email: service@firstchoicedebtrelief.com"
    
    **2. Credit Concerns Response Emails**
    **Template Approach:**
    - Acknowledge credit importance
    - Reframe focus from credit as a borrowing tool to financial independence
    - Explain that resolving balances creates a foundation for rebuilding
    - Clarify that if payments are already behind, impact is already occurring
    - Avoid guarantees about credit improvement timelines
    
    **Example Email Content:**
    "Thank you for sharing your concerns about your credit. I completely understand that this is an important aspect of your financial picture, and it's natural to be concerned about it.
    What we've seen is that by resolving these accounts, clients can actually set themselves up to rebuild on a stronger foundation. While the program is focused on debt resolution rather than credit improvement, the goal is to help you become debt-free significantly faster than making minimum payments, which gives you more financial flexibility in the long run.
    The current focus is on getting you out of debt so you can keep more of your money each month instead of paying toward interest and minimums. Once your debts are resolved, you'll be in a better position to rebuild your credit profile if that's important to you.
    If you have specific questions or concerns about your individual situation, please don't hesitate to call us at 800-985-9319, and we can discuss this in more detail.
    We're here to support you throughout this journey to financial freedom.
    
    Best regards,
    Client Services Team
    First Choice Debt Relief
    Phone: 800-985-9319
    Email: service@firstchoicedebtrelief.com"
    
    **3. Legal Protection Emails**
    **Template Approach:**
    - Explain that legal insurance covers attorney costs if legal action occurs
    - Clarify that insurance cannot prevent lawsuits from happening
    - Emphasize FCDR's coordination with legal providers
    - Describe creditors' typical escalation process before legal action
    - Avoid language suggesting complete protection from legal action
    
    **Example Email Content:**
    "I'm reaching out with a quick update on your legal case. Your assigned legal provider is actively working on your behalf, and we're staying in close communication with their office to support the process.
    Important: Your legal provider may contact you directly, especially if a potential settlement becomes available. If that happens, please connect with us before making any decisions. We'll help you review the offer based on your available funds and program progress so you can make the most informed decision.
    If you're able to contribute additional funds — through a one-time deposit or an increase in your monthly draft — this may help resolve the account faster and give your legal provider more flexibility during negotiations. Just let us know if that's something you'd like to explore.
    We're here to support you every step of the way. Feel free to reply to this email or call us at 800-985-9319 with any questions.
    
    Best regards,
    First Choice Debt Relief - Client Services"
    
    **4. Program Cost Concerns Emails**
    **Template Approach:**
    - Acknowledge concern with empathy
    - Explain how minimum payments primarily go to interest, not principal
    - Reframe as redirecting existing payments more effectively
    - Compare long-term interest costs to program costs when appropriate
    - Avoid dismissive responses about affordability
    
    **Example Email Content:**
    "Thank you for expressing your concerns about the program cost. I completely understand that when you're already juggling multiple payments, this can feel like an additional burden.
    I'd like to offer a different perspective: With your current debt payments, a significant portion goes straight to interest and minimum payments, which means you're spending more in the long run just to maintain your current position. Through our program, we're consolidating those payments and focusing on reducing what you owe, not just covering interest.
    If you continued making minimum payments, you'd likely pay significantly more in interest alone than you would in this program. Our goal is to help you become debt-free faster and save money long-term.
    That said, if you'd like to discuss your specific financial situation and explore potential adjustments to make the program more manageable, please call us at 800-985-9319. We're committed to finding a solution that works for your unique circumstances.
    We're here to support you on your journey to financial freedom.
    
    Best regards,
    Client Services Team
    First Choice Debt Relief
    Phone: 800-985-9319
    Email: service@firstchoicedebtrelief.com"
    
    **5. Account Exclusion Emails**
    **Template Approach:**
    - Acknowledge desire to keep accounts as backup
    - Focus on freeing up cash flow by resolving balances
    - Explain strategic negotiation benefits
    - Address maxed-out cards realistically
    - Avoid demanding account closure or suggesting they "must" close accounts
    
    **Example Email Content:**
    "Thank you for your inquiry about leaving certain accounts out of your program. I understand the desire to maintain some financial flexibility by keeping certain accounts open.
    When negotiating with creditors, we need to be strategic. If one account is being resolved while another is left out, it can create what we call 'creditor jealousy.' Essentially, some creditors might question why one account is receiving assistance while theirs isn't, which can impact how willing they are to work with us.
    However, I notice that we've already structured your program to exclude [specific accounts] to maintain some flexibility for you. The primary goal is to help you free up cash flow, reduce your balances, and regain financial control.
    If you'd like to discuss specific accounts or have concerns about your current program structure, please call us at 800-985-9319 so we can review your particular situation in detail.
    We appreciate your commitment to the program and are here to support your financial recovery.
    
    Best regards,
    Client Services Team
    First Choice Debt Relief
    Phone: 800-985-9319
    Email: service@firstchoicedebtrelief.com"
    
    **6. Loan Qualification Issues Emails**
    **Template Approach:**
    - Acknowledge frustration empathetically
    - Explain that pre-qualification is based on initial data
    - Clarify how changing circumstances affect loan approval
    - Offer information about future options after program progress
    - Avoid guarantees about future loan qualification
    
    **Example Email Content:**
    "I understand your frustration regarding the loan qualification. The pre-qualification is based on initial data, but the final approval considers your current financial situation. If things have changed — like missed payments or higher balances — that can impact the outcome.
    The good news is, our program is still designed to get you where you need to be financially, and after 8-12 consistent payments, you can reapply for the loan with potentially better terms. During this time, we'll continue working to resolve your accounts according to the program.
    Please know that we're still committed to helping you achieve your financial goals, even if the path looks slightly different than initially expected. This temporary setback doesn't change the overall effectiveness of the debt resolution strategy.
    If you'd like to discuss your specific situation in more detail or explore other options, please call us at 800-985-9319. We're here to support you throughout this process.
    
    Best regards,
    Client Services Team
    First Choice Debt Relief
    Phone: 800-985-9319
    Email: service@firstchoicedebtrelief.com"
    
    **7. Decision Uncertainty Emails**
    **Template Approach:**
    - Break down available options clearly
    - Compare debt resolution to alternatives (minimum payments, loans)
    - Address specific concerns about chosen option
    - Provide realistic benefits without overpromising
    - Avoid pressuring language or creating artificial urgency
    
    **Example Email Content:**
    "Thank you for sharing your thoughts with me. I completely understand feeling uncertain about which direction to take with your finances.
    Let's break down your options realistically:
    1. Continuing with minimum payments: Based on what you've shared, this would keep you in debt for many years longer and cost significantly more in interest over time.
    2. Debt consolidation loan: While this could simplify payments, the interest rates available to you currently may not provide much savings, and it doesn't address the underlying debt amount.
    3. Debt resolution program: This option is designed to reduce both your monthly payment and the total amount paid over time, helping you become debt-free faster than minimum payments.
    Every financial situation is unique, and there's no perfect solution for everyone. Our goal is to help you find the approach that balances immediate relief with long-term financial health.
    If you'd like to discuss your specific concerns in more detail or get answers to any additional questions, please feel free to call us at 800-985-9319. We're here to help you make the decision that's right for you.
    
    Best regards,
    Client Services Team
    First Choice Debt Relief
    Phone: 800-985-9319
    Email: service@firstchoicedebtrelief.com"
    
    ### Email Best Practices - DO:
    - Emphasize the program's ability to help clients become debt-free YEARS faster than minimum payments
    - Highlight monthly payment relief compared to current debt obligations
    - Acknowledge when settlement payments are tied to specific creditor agreements
    - Emphasize FCDR's role as an intermediary between clients and creditors/legal providers
    - Explain that additional fund contributions can expedite account resolution when appropriate
    - Highlight the importance of maintaining settlement terms to preserve negotiated savings
    - Use beneficial phrases like "financial freedom," "relief from high interest," and "relief from credit utilization"
    - Follow specific template structures when provided with exact formatting
    - Verify information before including it in communications
    - Document all client interactions according to company protocols
    - Maintain a solution-oriented approach to all client concerns
    - Thank clients for reaching out and contacting FCDR
    - Validate client concerns while providing reassurance
    - Use estimates rather than definitive statements
    - Highlight client control over settlement decisions
    - Encourage open communication and questions
    
    ### Email Best Practices - DON'T:
    - Quote specific debt reduction percentages or exact savings amounts unless explicitly provided
    - Make guarantees about specific settlement outcomes or timeframes
    - Suggest clients contact creditors directly - always direct them to FCDR
    - Use high-pressure sales tactics or create artificial urgency
    - Add commentary before or after requested email templates
    - Mention competitors' debt programs or directly compare services
    - Include complex financial or legal terminology without clear explanation
    - Use aggressive language about creditors or the client's financial situation
    - Communicate with unauthorized third parties about client accounts
    - Send sensitive information through unsecured channels
    - Initiate contact outside permitted hours (8am-8pm client local time)
    - Express personal opinions about financial matters or client situations
    - Dismiss or minimize client concerns (e.g., "Don't worry, it's no big deal")
    - Imply that settlements are certain or predetermined
    - Guarantee that legal insurance will prevent lawsuits
    
    ## INDUSTRY TERMINOLOGY & JARGON
    
    ### Key Terms & Definitions
    - **Settlement**: Agreement between creditor and client to resolve debt for less than full balance
    - **Gateway/Custodial Account**: Dedicated account where client funds are held for settlements
    - **Creditor**: Original lender or current debt owner (bank, credit card company, collection agency)
    - **Legal Plan**: Third-party legal service that provides representation for clients facing lawsuits
    - **Draft/Program Payment**: Regular client payment into their dedicated account
    - **Settlement Percentage**: The portion of original debt that will be paid in settlement (e.g., 40%)
    - **Program Length**: The estimated duration of a client's debt resolution program
    - **Service Fee**: Fees charged by FCDR for debt negotiation and program management
    - **Letter of Authorization (LOA)**: Document authorizing FCDR to communicate with creditors
    - **Debt Resolution Agreement (DRA)**: Primary contract between client and FCDR
    - **Summons/Complaint**: Legal documents initiating a lawsuit from creditor against client
    
    ### Usage Guidelines
    - Use industry terminology appropriately when communicating with employees
    - Provide brief explanations when using specialized terms in client-facing communications
    - Maintain consistent terminology across related communications
    - Recognize department-specific terminology differences
    - Adapt language complexity based on the employee's role and expertise
    
    ## RESPONSE APPROACH & STANDARDS
    
    ### Response Prioritization
    - Address safety, compliance, and time-sensitive issues first
    - Break down complex requests into clearly defined components
    - Create structured responses with headers, bullet points, or numbered lists for clarity
    - For multi-part questions, maintain the same order as in the original request
    - Flag which items require immediate action versus future consideration
    
    ### Organization Principles
    - Prioritize actionable information at the beginning of responses
    - Suggest batching similar tasks when multiple requests are presented
    - Identify dependencies between tasks and suggest logical sequencing
    - Recommend appropriate delegation when tasks span multiple departments
    - Balance thoroughness with conciseness based on urgency and importance
    
    ### Professional Response Approach
    1. Demonstrate financial expertise while maintaining accessible language
    2. Approach all inquiries with a solution-focused mindset aligned with the company mission
    3. When discussing financial matters, balance honesty about challenges with optimism about resolution
    4. Maintain appropriate professional boundaries while showing genuine concern for clients
    5. Provide context for how recommendations support the client's financial recovery journey
    6. Structure all information clearly and logically to prioritize comprehension
    7. Reference client files or documents by their exact names for clarity and record-keeping
    8. Explain financial concepts at an appropriate level based on context
    9. Seek clarification on financial details when necessary for accurate assistance
    10. For questions outside debt relief, provide helpful, professional responses while maintaining quality standards
    
    ### Casual Conversation & Chitchat
    - Display a friendly, personable demeanor while maintaining professional boundaries
    - Show measured enthusiasm and positivity that reflects FCDR's supportive culture
    - Exhibit a light sense of humor appropriate for workplace interactions
    - Demonstrate emotional intelligence by recognizing and responding to social cues
    - Balance warmth with professionalism, avoiding overly casual or informal language
    - Engage naturally in brief small talk while gently steering toward productivity
    - Respond to personal questions with appropriate, general answers that don't overshare
    - Show interest in user experiences without prying or asking personal questions
    - Acknowledge special occasions (holidays, company milestones) with brief, appropriate messages
    - Participate in light team-building conversations while maintaining a service-oriented focus
    
    ## ERROR HANDLING & LIMITATIONS
    
    ### Knowledge Boundaries
    - Acknowledge when a request requires information not available in your training
    - Clearly communicate limits without apologizing excessively or being defensive
    - Offer alternative approaches when you cannot fulfill the exact request
    - Suggest resources or colleagues who might have the specialized information needed
    - Never guess about compliance-related matters or specific client accounts
    
    ### Request Clarification
    - Ask specific questions to narrow down ambiguous requests
    - Seek clarification on account details, client information, or process steps when needed
    - When documents or emails contain unclear elements, request specific clarification
    - Verify understanding of complex requests by summarizing before proceeding
    - Be direct about what additional information would help you provide better assistance
    
    ### Sensitive Information
    - Immediately flag if users are sharing information that shouldn't be communicated in this channel
    - Redirect requests for sensitive client information to appropriate secure systems
    - Remind users about proper channels for sharing protected information when relevant
    - Never store or repeat sensitive information like SSNs, full account numbers, or complete DOBs
    - Guide users to redact sensitive information when sharing documents for review
    
    
    PS: Remember to embody First Choice Debt Relief's commitment to helping clients achieve financial freedom through every interaction, supporting employees in providing exceptional service at each client touchpoint.
    PS: Remember to use "RETRIEVED KNOWLEDGE" to enrich your response (if relevant and applicable)
    '''
def create_unified_email_card(state=None, active_view="main", email_type=None):
    """
    Creates a simplified adaptive card for email creation with just Client Service and Sales Service options.
    
    Args:
        state: The conversation state (optional)
        active_view: Which view to display (main, form)
        email_type: Type of email (client_service, sales_service)
    
    Returns:
        Attachment: The complete adaptive card
    """
    # Base card structure
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "style": "emphasis",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "First Choice Debt Relief Email Creator",
                        "size": "large",
                        "weight": "bolder",
                        "horizontalAlignment": "center"
                    }
                ],
                "bleed": True
            }
        ]
    }
    
    # For the main view, show the category selector
    if active_view == "main":
        # Main container for selection
        main_container = {
            "type": "Container",
            "items": [
                {
                    "type": "TextBlock",
                    "text": "Select an email category to get started",
                    "wrap": True,
                    "weight": "bolder",
                    "size": "medium",
                    "horizontalAlignment": "center",
                    "spacing": "medium"
                },
                {
                    "type": "ColumnSet",
                    "columns": [
                        {
                            "type": "Column",
                            "width": "stretch",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": "📧 Client Service",
                                    "weight": "bolder",
                                    "horizontalAlignment": "center"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": "Service emails to existing clients",
                                    "wrap": True,
                                    "size": "small",
                                    "horizontalAlignment": "center",
                                    "spacing": "none"
                                },
                                {
                                    "type": "ActionSet",
                                    "actions": [
                                        {
                                            "type": "Action.Submit",
                                            "title": "Select",
                                            "style": "positive",
                                            "data": {
                                                "action": "view_change",
                                                "view": "form",
                                                "email_type": "client_service"
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
                                    "type": "TextBlock",
                                    "text": "💼 Sales Service",
                                    "weight": "bolder",
                                    "horizontalAlignment": "center"
                                },
                                {
                                    "type": "TextBlock",
                                    "text": "Sales emails and program offerings",
                                    "wrap": True,
                                    "size": "small",
                                    "horizontalAlignment": "center",
                                    "spacing": "none"
                                },
                                {
                                    "type": "ActionSet",
                                    "actions": [
                                        {
                                            "type": "Action.Submit",
                                            "title": "Select",
                                            "style": "positive",
                                            "data": {
                                                "action": "view_change",
                                                "view": "form",
                                                "email_type": "sales_service"
                                            }
                                        }
                                    ]
                                }
                            ]
                        }
                    ]
                },
                {
                    "type": "Container",
                    "style": "attention",
                    "items": [
                        {
                            "type": "TextBlock",
                            "text": "Compliance Reminder",
                            "weight": "bolder",
                            "horizontalAlignment": "center",
                            "size": "small"
                        },
                        {
                            "type": "TextBlock",
                            "text": "Always ensure communications follow FCDR compliance guidelines.",
                            "wrap": True,
                            "size": "small",
                            "horizontalAlignment": "center"
                        }
                    ],
                    "spacing": "medium"
                }
            ]
        }
        card["body"].append(main_container)
        
    # Form view for email creation
    elif active_view == "form":
        title = "Client Service Email" if email_type == "client_service" else "Sales Service Email"
        card["body"][0]["items"][0]["text"] = title
        
        # Create form container
        form_container = {
            "type": "Container",
            "items": [
                {
                    "type": "TextBlock",
                    "text": "Email Information",
                    "wrap": True,
                    "weight": "bolder",
                    "size": "medium",
                    "spacing": "medium"
                },
                {
                    "type": "TextBlock",
                    "text": "Recipient",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "recipient",
                    "placeholder": "Enter recipient name or email"
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
                    "placeholder": "Describe what you want in this email - be specific about the purpose, tone, and key points to include",
                    "isMultiline": True,
                    "style": "text"
                },
                {
                    "type": "TextBlock",
                    "text": "Additional Context (Optional)",
                    "wrap": True
                },
                {
                    "type": "Input.Text",
                    "id": "context",
                    "placeholder": "Any additional context, previous conversations, or specific details",
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
            ]
        }
        
        # Add buttons
        form_container["items"].append({
            "type": "ActionSet",
            "actions": [
                {
                    "type": "Action.Submit",
                    "title": "Generate Email",
                    "style": "positive",
                    "data": {
                        "action": "generate_email",
                        "email_type": email_type
                    }
                },
                {
                    "type": "Action.Submit",
                    "title": "Back",
                    "data": {
                        "action": "view_change",
                        "view": "main"
                    }
                }
            ],
            "spacing": "medium"
        })
        
        card["body"].append(form_container)
    
    # Create attachment
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    return attachment
async def retrieve_documents(query, top=5, mode="openai"):
    """
    Retrieves documents from either OpenAI API (default) or Azure AI Search.
    Falls back to Azure Search if OpenAI retrieval fails.
    
    Args:
        query (str): The search query
        top (int): Maximum number of results to return
        mode (str): Search mode - "openai" (default) or "azure_search"
    
    Returns:
        list: List of document dictionaries with title and content keys
              Example: [{"title": "Document Name", "content": "Document content text..."}]
              Returns empty list if query is unrelated or no documents found
    """
    try:
        # If mode is explicitly set to azure_search, use that
        if mode == "azure_search":
            return await _retrieve_with_azure_search(query, top)
        # Default to OpenAI with fallback to Azure Search
        return await _retrieve_with_openai(query, top)
    except Exception as e:
        logging.error(f"Error retrieving documents: {e}")
        import traceback
        traceback.print_exc()
        return []

async def _retrieve_with_azure_search(query, top=5):
    """Azure Search implementation for document retrieval"""
    try:
        search_client = create_search_client()
        if not search_client:
            return []
            
        # Use only basic search parameters that work with any SDK version
        results = search_client.search(
            search_text=query,
            top=top
        )
        
        documents = []
        
        # Process search results
        for item in results:
            # Try to get content from various possible field names
            content = None
            for field_name in ["chunk", "content", "text"]:
                if field_name in item:
                    content = item[field_name]
                    if content:
                        break
            
            if not content:
                # If we can't find a content field, look for any string field
                for key, value in item.items():
                    if isinstance(value, str) and len(value) > 50:
                        content = value
                        break
            
            if not content:
                continue
            
            # Try to get a title
            title = None
            for title_field in ["title", "name", "filename"]:
                if title_field in item:
                    title = item[title_field]
                    if title:
                        break
            
            if not title:
                # Use a key as title if available
                for key_field in ["id", "key", "chunk_id"]:
                    if key_field in item:
                        title = f"Document {item[key_field]}"
                        break
                        
            if not title:
                title = "Unknown Document"
            
            documents.append({
                "title": title,
                "content": content
            })
        
        return documents
        
    except Exception as e:
        logging.error(f"Azure Search retrieval error: {e}")
        return []

async def _retrieve_with_openai(query, top=5, max_retries=1):
    """OpenAI implementation for document retrieval with retry logic for JSON format issues"""
    original_error = None
    
    for retry_count in range(max_retries + 1):  # +1 for initial attempt
        try:
            client = create_client()
            if not client:
                return []
            
            # Create a prompt for the OpenAI model to retrieve relevant information
            system_prompt = """You are a high-precision retrieval system for First Choice Debt Relief (FCDR) documentation.

            ## YOUR ROLE
            Your task is to accurately retrieve relevant information from First Choice Debt Relief's knowledge base in response to queries.
            You will extract knowledge from your training on debt relief concepts, policies, procedures, and FCDR-specific information.
            You must maintain the exact format requested and never invent or hallucinate document names or content.
            
            ## DOMAIN KNOWLEDGE
            First Choice Debt Relief specializes in:
            - Debt resolution programs
            - Settlement negotiation with creditors
            - Client enrollment and onboarding
            - Compliance and legal protection
            - Payment processing and management
            - Financial counseling and education
            - Creditor relationships and communication
            - Program management and client services
            
            ## DOCUMENT FORMAT REQUIREMENTS
            1. The "title" field MUST contain an EXACT document name with proper extension, such as:
               - "Client Enrollment Guide.pdf"
               - "Settlement Procedures Manual.docx"
               - "Compliance Handbook v3.2.pdf"
               - "Legal Protection Plan Overview.ppt"
               - "Client Services Training Manual.pdf"
               - "Creditor Communication Protocols.doc"
               - "Payment Gateway Management Guide.pdf"
               - "First Choice Debt Relief Program Overview.docx"
            
            2. The "content" field MUST contain specific, relevant information that directly addresses the query.
               - Content should be substantive and detailed
               - Information should be coherent and complete
               - Content should be clearly related to the document title
               - Do not include placeholder or generic content
            
            ## RELEVANCE DETERMINATION
            - ONLY return documents that contain information DIRECTLY relevant to the query
            - Return an empty array if the query is unrelated to debt relief or FCDR
            - Return an empty array if you are uncertain about the information
            - Do not stretch relevance - be conservative in your selections
            - Prioritize content that specifically answers the query vs. general information
            
            ## RESPONSE FORMAT
            You MUST structure your response as valid JSON with a documents array.
            Each document needs both a "title" (document name) and "content" (relevant text) field.
            """
            
            # Add feedback about previous error if this is a retry
            format_reminder = ""
            if retry_count > 0 and original_error:
                format_reminder = f"""
                IMPORTANT FORMAT CORRECTION NEEDED: Previous attempt failed with error: {original_error}
                You MUST return a valid JSON object with the exact format specified below.
                DO NOT add any text before or after the JSON object.
                The response MUST contain a 'documents' array (even if empty).
                Each document MUST have both 'title' and 'content' fields.
                """
            
            user_prompt = f"""Based on the following query, provide relevant information from First Choice Debt Relief knowledge base.
            
            Query: {query}
            INSTRUCTIONS:
            1. Search for information specifically related to this query
            2. Only return information if it's directly relevant to First Choice Debt Relief operations or debt relief concepts
            3. Return an empty array if no relevant information exists or if you're uncertain
            4. Include real document names with extensions (e.g., .pdf, .docx, .ppt) in the title field
            5. Provide substantial, relevant content in the content field
            6. Return no more than {top} of the most relevant documents
            {format_reminder}
            
            REQUIRED FORMAT:
            {{
                "response_type": "json",
                "documents": [
                    {{
                        "title": "Exact Name of the Source Document (e.g., 'ClientOnboardingManual.pdf', 'LegalDisclosurePolicy.ppt')",
                        "content": "Relevant content from source document"
                    }},
                    {{
                        "title": "Exact Name of the Source Document (e.g., 'SettlementGuide.docx', 'ComplianceManual.pdf')",
                        "content": "Relevant content from source document"
                    }}
                ]
            }}
            
            If the query is unrelated to First Choice Debt Relief or debt relief programs, or if you don't have relevant information, return:
            {{
                "response_type": "json",
                "documents": []
            }}
            
            Return no more than {top} document entries. Each document must have both a title (document name) and content fields."""
            
            # Call OpenAI API
            response = client.chat.completions.create(
                model="gpt-4.1-mini",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                response_format={"type": "json_object"},
                temperature=0
            )
            
            # Parse the response
            response_text = response.choices[0].message.content
            response_json = json.loads(response_text)
            
            # Extract documents from the response
            if "documents" in response_json and isinstance(response_json["documents"], list):
                # Ensure each document has both title and content
                documents = []
                for doc in response_json["documents"]:
                    if isinstance(doc, dict) and "title" in doc and "content" in doc:
                        documents.append({
                            "title": doc["title"],
                            "content": doc["content"]
                        })
                return documents
            else:
                # Format problem with the JSON structure - missing documents field
                error_msg = "JSON response missing 'documents' field or not in expected format"
                logging.warning(f"Invalid response format from OpenAI: {error_msg}")
                
                # Store error for retry and continue to next attempt if retries remaining
                if retry_count < max_retries:
                    original_error = error_msg
                    logging.info(f"Retrying OpenAI retrieval (attempt {retry_count + 1}/{max_retries + 1})")
                    continue
                return []
                
        except json.JSONDecodeError as json_error:
            # JSON parsing error - the model didn't return valid JSON
            error_msg = f"Failed to parse JSON: {str(json_error)}"
            logging.error(error_msg)
            
            # Store error for retry and continue to next attempt if retries remaining
            if retry_count < max_retries:
                original_error = error_msg
                logging.info(f"Retrying OpenAI retrieval (attempt {retry_count + 1}/{max_retries + 1})")
                continue
            return []
            
        except Exception as e:
            logging.error(f"OpenAI retrieval error: {e}")
            raise  # Re-raise other types of errors to trigger fallback
    
    # If we get here, all retries have failed
    logging.error("All retrieval attempts failed")
    return []
def create_new_chat_card():
    """Creates an adaptive card for starting a new chat"""
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "TextBlock",
                "text": "Start a New Conversation",
                "size": "medium",
                "weight": "bolder",
                "horizontalAlignment": "center"
            },
            {
                "type": "TextBlock",
                "text": "Your previous conversation has ended. Would you like to start a new one?",
                "wrap": True,
                "horizontalAlignment": "center"
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Start New Chat",
                "data": {
                    "action": "new_chat"
                },
                "style": "positive"
            }
        ]
    }
    
    return Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
async def handle_new_chat_command(turn_context: TurnContext, state, conversation_id):
    """Handles commands to start a new chat or reset the current chat"""
    # Send typing indicator
    await turn_context.send_activity(create_typing_activity())
    
    # Clear any pending messages for this conversation
    with pending_messages_lock:
        if conversation_id in pending_messages:
            pending_messages[conversation_id].clear()
    
    # Send a message informing the user
    await turn_context.send_activity("Starting a new conversation...")
    
    # Initialize a new chat
    await initialize_chat(turn_context, None)  # Pass None to force new state creation

# Add this to your bot initialization code

async def on_adaptive_card_invoke(turn_context: TurnContext, invoke_value: dict):
    """Handler for Adaptive Card invocations"""
    try:
        # Extract the action data from the invoke value
        action_data = invoke_value.get("action", {})
        
        # Process the action through our card handler
        await handle_card_actions(turn_context, action_data)
        
        # Return a successful response
        return {
            "statusCode": 200,
            "type": "application/vnd.microsoft.activity.message",
            "value": "Action processed successfully"
        }
    except Exception as e:
        logging.error(f"Error handling adaptive card invoke: {e}")
        traceback.print_exc()
        
        # Return an error response
        return {
            "statusCode": 500,
            "type": "application/vnd.microsoft.error",
            "value": str(e)
        }
# async def handle_new_chat_command(turn_context: TurnContext, state, conversation_id):
#     """Handles commands to start a new chat or reset the current chat"""
#     # Send typing indicator
#     await turn_context.send_activity(create_typing_activity())
    
#     # Clear any pending messages for this conversation
#     with pending_messages_lock:
#         if conversation_id in pending_messages:
#             pending_messages[conversation_id].clear()
    
#     # Send a message informing the user
#     await turn_context.send_activity("Starting a new conversation...")
    
#     # Initialize a new chat
#     await initialize_chat(turn_context, None)  # Pass None to force new state creation
def create_typing_stop_activity():
    """Creates an activity to explicitly stop the typing indicator"""
    return Activity(
        type=ActivityTypes.typing,  # Use typing type, not message
        channel_id="msteams",
        value={"isTyping": False}  # Signal to stop typing
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
    def replace_buffer_with(self, text: str) -> None:
        """Make *text* the only thing that will be sent in send_final_message()."""
        self.message_parts.clear()
        self.message_parts.append(text)
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

async def handle_thread_recovery(turn_context: TurnContext, state, error_message, recovery_context=None):
    """
    Handles recovery from thread or assistant errors with enhanced self-healing capabilities.
    
    Args:
        turn_context: The turn context
        state: The conversation state
        error_message: The error that triggered recovery
        recovery_context: Additional context for recovery (optional)
    """
    # Get user identity for safety checks and logging
    user_id = turn_context.activity.from_property.id if hasattr(turn_context.activity, 'from_property') else "unknown"
    conversation_id = TurnContext.get_conversation_reference(turn_context.activity).conversation.id
    
    # Increment recovery attempts (with thread safety)
    with conversation_states_lock:
        state["recovery_attempts"] = state.get("recovery_attempts", 0) + 1
        state["last_error"] = error_message
        state["error_history"] = state.get("error_history", [])
        state["error_history"].append({
            "error": str(error_message)[:200],
            "timestamp": datetime.now().isoformat(),
            "context": recovery_context
        })
        # Keep only last 5 errors
        if len(state["error_history"]) > 5:
            state["error_history"] = state["error_history"][-5:]
        
        recovery_attempts = state["recovery_attempts"]
    
    # Log recovery attempt with user context
    logging.info(f"Recovery attempt #{recovery_attempts} for user {user_id}: {error_message}")
    
    # Analyze error pattern to determine best recovery strategy
    error_str = str(error_message).lower()
    recovery_strategy = determine_recovery_strategy(error_str, state.get("error_history", []))
    
    try:
        # Apply recovery strategy based on error type
        if recovery_strategy == "rate_limit":
            # Rate limit error - wait and retry
            wait_time = min(recovery_attempts * 5, 30)  # Progressive backoff, max 30s
            await turn_context.send_activity(f"I'm experiencing high demand. Please wait {wait_time} seconds...")
            await asyncio.sleep(wait_time)
            
            # Reset recovery attempts for rate limits
            with conversation_states_lock:
                state["recovery_attempts"] = 0
            
            # Try to continue with the original request
            return True
            
        elif recovery_strategy == "thread_not_found":
            # Thread is gone - create new one but try to preserve context
            await turn_context.send_activity("I need to refresh our conversation. One moment...")
            
            # Preserve important context
            preserved_context = await preserve_conversation_context(state)
            
            # Create new resources
            success = await create_fresh_resources(turn_context, state, user_id, conversation_id, preserved_context)
            
            if success:
                await turn_context.send_activity("I've refreshed our conversation and I'm ready to continue. What can I help you with?")
                return True
            else:
                raise Exception("Failed to create fresh resources")
                
        elif recovery_strategy == "assistant_error":
            # Assistant configuration issue - recreate assistant
            await turn_context.send_activity("Updating my configuration. This will just take a moment...")
            
            client = create_client()
            
            # Try to preserve the thread if possible
            thread_id = state.get("session_id")
            if thread_id and await verify_thread_exists(client, thread_id):
                # Thread is okay, just recreate assistant
                new_assistant = await create_new_assistant(client, user_id, conversation_id)
                
                with conversation_states_lock:
                    state["assistant_id"] = new_assistant.id
                    state["recovery_attempts"] = 0
                
                await turn_context.send_activity("Configuration updated. I'm ready to help!")
                return True
            else:
                # Both thread and assistant need recreation
                success = await create_fresh_resources(turn_context, state, user_id, conversation_id)
                if success:
                    await turn_context.send_activity("I've completely refreshed our session. How can I help you?")
                    return True
                else:
                    raise Exception("Failed to recover assistant and thread")
                    
        elif recovery_strategy == "network_error":
            # Network/connection issue - wait and retry with exponential backoff
            wait_time = min(2 ** recovery_attempts, 16)  # Exponential backoff, max 16s
            await turn_context.send_activity(f"Connection issue detected. Retrying in {wait_time} seconds...")
            await asyncio.sleep(wait_time)
            
            # Test connection
            if await test_openai_connection():
                with conversation_states_lock:
                    state["recovery_attempts"] = 0
                return True
            else:
                raise Exception("Persistent network connection issue")
                
        elif recovery_strategy == "resource_cleanup":
            # Too many attempts or persistent errors - full cleanup
            await turn_context.send_activity("I'm performing a complete refresh to resolve persistent issues...")
            
            # Clean up all resources
            await cleanup_all_resources(state)
            
            # Create everything fresh
            success = await create_fresh_resources(turn_context, state, user_id, conversation_id)
            
            if success:
                # Reset all error tracking
                with conversation_states_lock:
                    state["recovery_attempts"] = 0
                    state["error_history"] = []
                    state["last_error"] = None
                
                await turn_context.send_activity("Complete refresh successful! I'm ready to help you.")
                return True
            else:
                raise Exception("Failed to perform complete resource cleanup and recreation")
                
        else:
            # Unknown error - try generic recovery
            if recovery_attempts < 3:
                # Try creating fresh resources
                await turn_context.send_activity("I'm experiencing an issue. Let me try a different approach...")
                
                success = await create_fresh_resources(turn_context, state, user_id, conversation_id)
                if success:
                    return True
                else:
                    raise Exception("Generic recovery failed")
            else:
                # Too many attempts - suggest manual reset
                raise Exception("Multiple recovery attempts failed")
                
    except Exception as recovery_error:
        logging.error(f"Recovery attempt #{recovery_attempts} failed: {recovery_error}")
        
        # If we've tried too many times, suggest a fresh start
        if recovery_attempts >= 3:
            # Reset the recovery counter
            with conversation_states_lock:
                state["recovery_attempts"] = 0
                state["fallback_mode"] = True
            
            # Send error message with new chat card
            await turn_context.send_activity(
                "I'm having persistent trouble with our conversation. "
                "Let's start fresh with a new chat session for the best experience."
            )
            await send_new_chat_card(turn_context)
            
            # Enable fallback mode for this conversation
            return False
        else:
            # Try fallback response mode
            await turn_context.send_activity(
                "I'm still having some trouble, but I'll do my best to help you. "
                "You can also try starting a new chat if issues persist."
            )
            
            # Set fallback mode
            with conversation_states_lock:
                state["fallback_mode"] = True
            
            return False


def determine_recovery_strategy(error_str: str, error_history: list) -> str:
    """Determine the best recovery strategy based on error patterns"""
    
    # Check for rate limit errors
    if any(indicator in error_str for indicator in ["rate limit", "429", "too many requests"]):
        return "rate_limit"
    
    # Check for thread not found errors
    if any(indicator in error_str for indicator in ["thread", "not found", "404", "invalid thread"]):
        return "thread_not_found"
    
    # Check for assistant errors
    if any(indicator in error_str for indicator in ["assistant", "not found", "invalid assistant"]):
        return "assistant_error"
    
    # Check for network errors
    if any(indicator in error_str for indicator in ["network", "connection", "timeout", "unreachable"]):
        return "network_error"
    
    # Check error history for patterns
    if len(error_history) >= 3:
        # If we have multiple similar errors, do a full cleanup
        recent_errors = [e["error"] for e in error_history[-3:]]
        if len(set(recent_errors)) == 1:  # Same error repeating
            return "resource_cleanup"
    
    return "unknown"


async def preserve_conversation_context(state: dict) -> dict:
    """Preserve important conversation context during recovery"""
    context = {}
    
    with conversation_states_lock:
        # Preserve email-related state
        if state.get("last_generated_email"):
            context["last_email"] = {
                "type": state.get("last_email_type"),
                "email": state.get("last_generated_email")[:500],  # First 500 chars
                "data": state.get("last_email_data")
            }
        
        # Preserve uploaded files list
        if state.get("uploaded_files"):
            context["uploaded_files"] = state.get("uploaded_files")
        
        # Preserve any custom context
        if state.get("user_context"):
            context["user_context"] = state.get("user_context")
    
    return context


async def create_fresh_resources(turn_context: TurnContext, state: dict, user_id: str, conversation_id: str, preserved_context: dict = None) -> bool:
    """Create completely fresh resources for the conversation"""
    try:
        client = create_client()
        
        # Create a new vector store
        vector_store = client.vector_stores.create(
            name=f"recovery_user_{user_id}_convo_{conversation_id}_{int(time.time())}"
        )
        
        # Create a new assistant
        assistant = await create_new_assistant(client, user_id, conversation_id, vector_store.id)
        
        # Create a new thread
        thread = client.beta.threads.create()
        
        # If we have preserved context, add it to the thread
        if preserved_context:
            context_message = "Previous conversation context:\n"
            if "last_email" in preserved_context:
                context_message += f"- Last generated {preserved_context['last_email']['type']} email\n"
            if "uploaded_files" in preserved_context:
                context_message += f"- Uploaded files: {', '.join(preserved_context['uploaded_files'])}\n"
            
            client.beta.threads.messages.create(
                thread_id=thread.id,
                role="user",
                content=context_message,
                metadata={"type": "recovered_context"}
            )
        
        # Update state with new resources
        with conversation_states_lock:
            old_thread = state.get("session_id")
            state["assistant_id"] = assistant.id
            state["session_id"] = thread.id
            state["vector_store_id"] = vector_store.id
            state["active_run"] = False
            
            # Restore preserved context
            if preserved_context and "last_email" in preserved_context:
                state["last_generated_email"] = preserved_context["last_email"]["email"]
                state["last_email_type"] = preserved_context["last_email"]["type"]
                state["last_email_data"] = preserved_context["last_email"]["data"]
        
        # Clear any active runs
        with active_runs_lock:
            if old_thread in active_runs:
                del active_runs[old_thread]
        
        logging.info(f"Successfully created fresh resources for user {user_id}")
        return True
        
    except Exception as e:
        logging.error(f"Failed to create fresh resources: {e}")
        return False


async def create_new_assistant(client: AzureOpenAI, user_id: str, conversation_id: str, vector_store_id: str = None) -> any:
    """Create a new assistant with proper configuration"""
    tools = [{"type": "file_search"}]
    tool_resources = {}
    
    if vector_store_id:
        tool_resources["file_search"] = {"vector_store_ids": [vector_store_id]}
    
    unique_name = f"recovery_assistant_user_{user_id}_{int(time.time())}"
    
    assistant = client.beta.assistants.create(
        name=unique_name,
        model="gpt-4.1-mini",
        instructions=SYSTEM_PROMPT,
        tools=tools,
        tool_resources=tool_resources if tool_resources else None,
    )
    
    return assistant


async def verify_thread_exists(client: AzureOpenAI, thread_id: str) -> bool:
    """Verify if a thread still exists"""
    try:
        thread = client.beta.threads.retrieve(thread_id=thread_id)
        return True
    except:
        return False


async def test_openai_connection() -> bool:
    """Test if we can connect to OpenAI"""
    try:
        client = create_client()
        # Try a simple API call
        client.models.list()
        return True
    except:
        return False


async def cleanup_all_resources(state: dict) -> None:
    """Clean up all OpenAI resources for a conversation"""
    try:
        client = create_client()
        
        # Try to delete assistant
        if state.get("assistant_id"):
            try:
                client.beta.assistants.delete(state["assistant_id"])
                logging.info(f"Deleted assistant {state['assistant_id']}")
            except:
                pass
        
        # Note: We can't delete threads or vector stores via API
        # Just log for tracking
        if state.get("session_id"):
            logging.info(f"Abandoning thread {state['session_id']}")
        
        if state.get("vector_store_id"):
            logging.info(f"Abandoning vector store {state['vector_store_id']}")
            
    except Exception as e:
        logging.error(f"Error during resource cleanup: {e}")

async def send_fallback_response(turn_context: TurnContext, user_message: str = None, context: dict = None):
    """
    Enhanced fallback response system with multiple tiers of fallback.
    
    Args:
        turn_context: The turn context
        user_message: The user's message (optional)
        context: Additional context for the fallback response
    """
    # Get conversation state
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_id = conversation_reference.conversation.id
    
    # Check if we have state
    state = conversation_states.get(conversation_id, {})
    fallback_level = state.get("fallback_level", 0)
    
    try:
        # Send typing indicator first
        await turn_context.send_activity(create_typing_activity())
        
        # Get user's message if not provided
        if not user_message:
            if hasattr(turn_context.activity, 'text'):
                user_message = turn_context.activity.text.strip()
            else:
                user_message = "Hello, I need your help."
        
        # Tier 1: Try direct completion with system prompt
        if fallback_level == 0:
            try:
                logging.info(f"Fallback Tier 1: Direct completion for user message: {user_message[:100]}...")
                
                client = create_client()
                
                # Build a context-aware fallback prompt
                fallback_prompt = f"""You are First Choice Debt Relief's AI Assistant in fallback mode. 
The main system is experiencing issues, but you should still try to help the user.

User's message: {user_message}

Provide a helpful response following these guidelines:
1. Be warm and professional
2. If this is about email generation, provide general guidance
3. If this is about debt relief, share general program information
4. Always maintain compliance - no guarantees or specific promises
5. Suggest they can try again or start a new chat if needed

Remember: You're helping an FCDR employee, not a client."""
                
                # Try with Azure Search for context
                response = client.chat.completions.create(
                    model="gpt-4.1-mini",
                    messages=[
                        {"role": "system", "content": fallback_prompt},
                        {"role": "user", "content": user_message}
                    ],
                    max_tokens=1000,
                    temperature=0.7,
                    extra_body={
                        "data_sources": [{
                            "type": "azure_search",
                            "parameters": {
                                "endpoint": AZURE_SEARCH_ENDPOINT,
                                "index_name": AZURE_SEARCH_INDEX_NAME,
                                "semantic_configuration": "default",
                                "query_type": "simple",
                                "fields_mapping": {},
                                "in_scope": True,
                                "role_information": SYSTEM_PROMPT[:1000],  # First 1000 chars
                                "strictness": 3,
                                "top_n_documents": 3,
                                "authentication": {
                                    "type": "api_key",
                                    "key": AZURE_SEARCH_KEY
                                }
                            }
                        }]
                    } if AZURE_SEARCH_ENDPOINT and AZURE_SEARCH_KEY else {}
                )
                
                if response.choices and response.choices[0].message.content:
                    fallback_text = response.choices[0].message.content
                    
                    # Add a note about fallback mode
                    fallback_text += "\n\n*Note: I'm operating in limited mode due to technical issues. For full functionality, you may want to start a new chat.*"
                    
                    await turn_context.send_activity(fallback_text)
                    
                    # Update fallback success
                    with conversation_states_lock:
                        state["fallback_success"] = True
                    
                    return
                else:
                    raise Exception("No response from fallback completion")
                    
            except Exception as tier1_error:
                logging.error(f"Fallback Tier 1 failed: {tier1_error}")
                fallback_level = 1
        
        # Tier 2: Simple completion without search
        if fallback_level == 1:
            try:
                logging.info("Fallback Tier 2: Simple completion without search")
                
                client = create_client()
                
                # Simpler prompt without search
                simple_response = client.chat.completions.create(
                    model="gpt-4.1-mini",
                    messages=[
                        {"role": "system", "content": "You are a helpful AI assistant for First Choice Debt Relief employees. Keep responses brief and helpful."},
                        {"role": "user", "content": user_message}
                    ],
                    max_tokens=500,
                    temperature=0.5
                )
                
                if simple_response.choices and simple_response.choices[0].message.content:
                    response_text = simple_response.choices[0].message.content
                    response_text += "\n\n*I'm in basic mode. Some features may be limited.*"
                    
                    await turn_context.send_activity(response_text)
                    return
                else:
                    raise Exception("No response from simple completion")
                    
            except Exception as tier2_error:
                logging.error(f"Fallback Tier 2 failed: {tier2_error}")
                fallback_level = 2
        
        # Tier 3: Context-based template responses
        if fallback_level >= 2:
            logging.info("Fallback Tier 3: Template-based response")
            
            # Analyze the user message for intent
            user_message_lower = user_message.lower()
            
            # Email-related fallback
            if any(keyword in user_message_lower for keyword in ["email", "template", "draft", "write"]):
                template_response = """I understand you need help with email creation. While I'm experiencing technical difficulties, here's what you can do:

1. **For Client Service Emails**: Focus on being supportive and solution-oriented. Always include the team signature with 800-985-9319.

2. **For Sales Emails**: Emphasize benefits without guarantees. Include your name and direct phone number.

3. **Key Compliance Reminders**:
   - Never promise specific outcomes
   - Avoid "debt forgiveness" language
   - Don't guarantee timeline or results
   
Would you like to try starting a new chat for full email generation capabilities?"""
                
                await turn_context.send_activity(template_response)
                
            # General help fallback
            elif any(keyword in user_message_lower for keyword in ["help", "assist", "support"]):
                help_response = """I'm here to help, though I'm currently in limited mode. I can assist with:

- General questions about FCDR policies
- Basic email guidance
- Compliance reminders
- Document-related queries

For full functionality, including email generation and document analysis, please try starting a new chat session.

What specific area do you need help with?"""
                
                await turn_context.send_activity(help_response)
                
            # Unknown intent fallback
            else:
                generic_response = """I'm experiencing some technical limitations right now, but I'm still here to help as best I can.

You can:
- Ask general questions about FCDR
- Get basic email writing tips
- Review compliance guidelines
- Start a new chat for full functionality

What would you like to know about?"""
                
                await turn_context.send_activity(generic_response)
            
            # Suggest new chat
            await send_new_chat_card(turn_context)
            
    except Exception as final_error:
        logging.error(f"All fallback tiers failed: {final_error}")
        
        # Ultimate fallback - just acknowledge and suggest new chat
        try:
            await turn_context.send_activity(
                "I apologize, but I'm unable to process your request right now due to technical issues. "
                "Please start a new chat session for the best experience, or contact IT support if the problem persists."
            )
            await send_new_chat_card(turn_context)
        except:
            # Even sending a message failed - log and give up
            logging.critical(f"Critical failure: Unable to send any response to user in conversation {conversation_id}")
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

def create_edit_email_card(original_email, email_id=None):
    """
    Creates an enhanced adaptive card for email editing with compliance guidance.
    
    Args:
        original_email: The original email text to edit
        email_id: Optional email ID for tracking
    
    Returns:
        Attachment: The card attachment
    """
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "style": "emphasis",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Edit Email",
                        "size": "large",
                        "weight": "bolder",
                        "horizontalAlignment": "center"
                    }
                ],
                "bleed": True
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Current Email:",
                        "wrap": True,
                        "weight": "bolder"
                    },
                    {
                        "type": "Container",
                        "style": "default",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": original_email,
                                "wrap": True
                            }
                        ],
                        "separator": True
                    }
                ],
                "spacing": "medium"
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "What changes would you like to make?",
                        "wrap": True,
                        "weight": "bolder"
                    },
                    {
                        "type": "Input.Text",
                        "id": "edit_instructions",
                        "placeholder": "E.g., 'Make it more concise', 'Add more details about payment options', 'Change the tone to be more urgent'",
                        "isMultiline": True,
                        "style": "text",
                        "height": "stretch"
                    }
                ],
                "spacing": "medium"
            },
            {
                "type": "Container",
                "style": "warning",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Compliance Reminder",
                        "weight": "bolder",
                        "size": "small"
                    },
                    {
                        "type": "TextBlock",
                        "text": "• Avoid making guarantees\n• Dont commit to timelines\n• Maintain professional tone",
                        "wrap": True,
                        "size": "small"
                    }
                ],
                "spacing": "medium"
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Apply Changes",
                "style": "positive",
                "data": {
                    "action": "apply_email_edits",
                    "email_id": email_id  # Include email_id
                }
            },
            {
                "type": "Action.Submit",
                "title": "Cancel",
                "data": {
                    "action": "cancel_edit",
                    "email_id": email_id  # Include email_id
                }
            }
        ]
    }
    
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    return attachment
async def send_edit_email_card(turn_context: TurnContext, state, email_id=None):
    """
    Sends an email editing card to the user.
    
    Args:
        turn_context: The turn context
        state: The conversation state containing the last generated email
        email_id: Optional specific email ID to edit
    """
    with conversation_states_lock:
        # If email_id is provided, try to get that specific email
        if email_id and "email_history" in state and email_id in state["email_history"]:
            email_data = state["email_history"][email_id]
            original_email = email_data.get("email_text", "")
        else:
            # Fallback to last generated email
            email_id = state.get("last_email_id")
            if email_id and "email_history" in state and email_id in state["email_history"]:
                email_data = state["email_history"][email_id]
                original_email = email_data.get("email_text", "")
            else:
                # Final fallback to old method
                original_email = state.get("last_generated_email", "")
    
    if not original_email:
        await turn_context.send_activity("I couldn't find a recently generated email to edit. Please create a new email first.")
        return
    
    edit_card = create_edit_email_card(original_email, email_id)
    await send_card_response(turn_context, edit_card)
async def apply_email_edits(turn_context: TurnContext, state, edit_instructions):
    """
    Applies edits to the previously generated email with intelligent retrieval support.
    
    Args:
        turn_context: The turn context
        state: The conversation state
        edit_instructions: Instructions for editing the email
    """
    # Send typing indicator
    await turn_context.send_activity(create_typing_activity())
    
    # Get the original email and metadata
    with conversation_states_lock:
        original_email = state.get("last_generated_email", "")
        email_type = state.get("last_email_type", "client_service")
        email_data = state.get("last_email_data", {})
    
    if not original_email:
        await turn_context.send_activity("I couldn't find the original email to edit. Please create a new email.")
        return
    
    # Build retrieval query based on email type and edit instructions
    retrieval_query = ""
    if email_type == "client_service":
        retrieval_query = "customer service email templates compliance guidelines "
    else:
        retrieval_query = "sales email templates compliance guidelines "
    
    retrieval_query += f"{edit_instructions} {email_data.get('subject', '')} {email_data.get('instructions', '')}"
    
    relevant_docs = await retrieve_documents(retrieval_query, top=3)
    
    # Format retrieved context
    retrieved_context = ""
    if relevant_docs:
        retrieved_context = "\n\n--- RETRIEVED KNOWLEDGE FOR EDITS ---\n\n"
        for doc in relevant_docs:
            if isinstance(doc, dict):
                content = doc.get("content", "")
                if content:
                    retrieved_context += f"{content[:1500]}...\n\n" if len(content) > 1500 else f"{content}\n\n"
    
    # Create prompt for editing
    email_category_text = "This is a Customer Service email." if email_type == "client_service" else "This is a Sales email."
    
    prompt = f"""Edit the following email based on these instructions: {edit_instructions}

{email_category_text}

ORIGINAL EMAIL:
{original_email}

{retrieved_context}

CRITICAL REQUIREMENTS:
1. Maintain the warm, human-like tone - don't make it sound robotic
2. Keep the same department signature format
3. Apply all compliance guidelines
4. Make the requested changes while preserving the email's purpose
5. Use appropriate template language if switching to a different type of email

COMPLIANCE REMINDERS:
- NEVER promise guaranteed results or specific outcomes
- NEVER offer legal advice
- NEVER use terms like 'debt forgiveness,' 'eliminate,' or 'erase' your debt
- NEVER state or imply that the program prevents lawsuits
- Use phrases like 'negotiated resolution' instead of 'paid in full'

Please provide the complete revised email with all changes incorporated."""
    
    # Initialize chat if needed
    if not state.get("assistant_id"):
        await initialize_chat(turn_context, state)
    
    try:
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
            
            email_card = create_email_result_card(edited_email)
            await send_card_response(turn_context, email_card)
        else:
            await turn_context.send_activity("I'm sorry, I couldn't edit the email. Please try again.")
    except Exception as e:
        logging.error(f"Error editing email: {str(e)}")
        traceback.print_exc()
        await turn_context.send_activity("I encountered an error while editing your email. Please try again.")
async def handle_card_actions(turn_context: TurnContext, action_data):
    """Handles actions from adaptive cards with simplified email UI"""
    try:
        conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
        conversation_id = conversation_reference.conversation.id
        
        # Check if we have state for this conversation
        if conversation_id not in conversation_states:
            # Initialize state if needed
            await initialize_chat(turn_context, None)
        
        state = conversation_states[conversation_id]
        
        # Check if busy for any action that will generate content
        action_type = action_data.get("action")
        
        # List of actions that require processing
        processing_actions = ["generate_email", "apply_email_edits", "create_email"]
        
        if action_type in processing_actions:
            if check_thread_busy(state):
                # Get current operation for better messaging
                active_operation = state.get("active_operation", "another request")
                
                # Create an informative card about the busy state
                busy_card = {
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
                                    "text": "⏳ Processing in Progress",
                                    "size": "medium",
                                    "weight": "bolder",
                                    "horizontalAlignment": "center"
                                }
                            ],
                            "bleed": True
                        },
                        {
                            "type": "Container",
                            "items": [
                                {
                                    "type": "TextBlock",
                                    "text": f"I'm currently busy with {active_operation}.",
                                    "wrap": True
                                },
                                {
                                    "type": "TextBlock",
                                    "text": "Please wait a moment and try again. Most operations complete within 10-30 seconds.",
                                    "wrap": True,
                                    "isSubtle": True
                                }
                            ]
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.Submit",
                            "title": "Try Again",
                            "data": action_data,
                            "style": "positive"
                        }
                    ]
                }
                
                attachment = Attachment(
                    content_type="application/vnd.microsoft.card.adaptive",
                    content=busy_card
                )
                
                await send_card_response(turn_context, attachment)
                return
        
        # Handle view changes for the unified card
        if action_type == "view_change":
            view = action_data.get("view", "main")
            email_type = action_data.get("email_type", None)
            
            # Create the unified card with appropriate view
            unified_card = create_unified_email_card(state, view, email_type)
            
            # Send the updated card
            await send_card_response(turn_context, unified_card)
            return
        
        # Handle email generation from unified card
        elif action_type == "generate_email":
            # Get email type
            email_type = action_data.get("email_type", "client_service")
            
            # Extract form fields
            recipient = action_data.get("recipient", "")
            subject = action_data.get("subject", "")
            instructions = action_data.get("instructions", "")
            context = action_data.get("context", "")
            has_attachments = action_data.get("hasAttachments", "false") == "true"
            
            # Validate required fields
            if not instructions:
                # Show error card
                error_card = {
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
                                    "text": "⚠️ Missing Required Information",
                                    "size": "medium",
                                    "weight": "bolder",
                                    "color": "attention"
                                }
                            ],
                            "bleed": True
                        },
                        {
                            "type": "TextBlock",
                            "text": "Please provide instructions for the email. This helps me understand what kind of email you need.",
                            "wrap": True
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.Submit",
                            "title": "Go Back",
                            "data": {
                                "action": "view_change",
                                "view": "form",
                                "email_type": email_type
                            }
                        }
                    ]
                }
                
                attachment = Attachment(
                    content_type="application/vnd.microsoft.card.adaptive",
                    content=error_card
                )
                
                await send_card_response(turn_context, attachment)
                return
            
            # Generate email using AI with retrieval
            await generate_email(
                turn_context, 
                state, 
                email_type,
                recipient, 
                subject, 
                instructions, 
                context,
                has_attachments
            )
            return
        
        # Handle other actions as before
        elif action_type == "new_chat":
            await handle_new_chat_command(turn_context, state, conversation_id)
            return
            
        elif action_type == "edit_email":
            # Get the email ID if provided
            email_id = action_data.get("email_id")
            await send_edit_email_card(turn_context, state, email_id)
            return
            
        elif action_type == "apply_email_edits":
            edit_instructions = action_data.get("edit_instructions", "")
            if not edit_instructions:
                await turn_context.send_activity("Please provide instructions for how you'd like to edit the email.")
                return
            await apply_email_edits(turn_context, state, edit_instructions)
            return
            
        elif action_type == "cancel_edit":
            # Get email ID from action data
            email_id = action_data.get("email_id")
            
            with conversation_states_lock:
                if email_id and "email_history" in state and email_id in state["email_history"]:
                    original_email = state["email_history"][email_id]["email_text"]
                else:
                    original_email = state.get("last_generated_email", "")
            
            if original_email:
                result_card = create_email_result_card(original_email)
                await send_card_response(turn_context, result_card)
            else:
                # Go back to main view if no email
                unified_card = create_unified_email_card(state, "main")
                await send_card_response(turn_context, unified_card)
            return
        
        elif action_type == "show_upload_info":
            await handle_info_request(turn_context, "upload")
            return
        
        elif action_type == "show_template_categories":
            # Show main email card
            unified_card = create_unified_email_card(state, "main")
            await send_card_response(turn_context, unified_card)
            return
        
        # Default to showing main view
        else:
            logging.warning(f"Unknown card action: {action_type}")
            unified_card = create_unified_email_card(state, "main")
            await send_card_response(turn_context, unified_card)
            return
            
    except Exception as e:
        logging.error(f"Error handling card action: {e}")
        traceback.print_exc()
        
        # Try to send a friendly error message
        try:
            await turn_context.send_activity("I encountered an error processing your request. Please try again.")
        except:
            pass  # If even error message fails, just log it
async def send_email_card(turn_context: TurnContext, template_mode="main", channel=None):
    """
    Sends a simplified unified email card.
    
    Args:
        turn_context: The turn context
        template_mode: The view to display (main or form)
        channel: Not used in simplified version
    """
    # Get state
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_id = conversation_reference.conversation.id
    
    if conversation_id in conversation_states:
        state = conversation_states[conversation_id]
    else:
        state = None
    
    # Always show main view when called from commands
    card = create_unified_email_card(state, "main")
    
    # Send the card
    await send_card_response(turn_context, card)
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


async def send_card_response(turn_context: TurnContext, attachment):
    """
    Properly creates and sends a response with an attachment.
    
    Args:
        turn_context: The turn context
        attachment: The attachment to send
    """
    reply = Activity(
        type=ActivityTypes.message,
        attachments=[attachment],
        input_hint=None  # Explicitly set input_hint
    )
    
    # Copy necessary properties from the incoming activity
    if hasattr(turn_context.activity, 'conversation'):
        reply.conversation = turn_context.activity.conversation
    
    if hasattr(turn_context.activity, 'from_property'):
        reply.recipient = ChannelAccount(
            id=turn_context.activity.from_property.id,
            name=turn_context.activity.from_property.name,
        )
    
    if hasattr(turn_context.activity, 'recipient'):
        reply.from_property = ChannelAccount(
            id=turn_context.activity.recipient.id,
            name=turn_context.activity.recipient.name,
        )
    
    if hasattr(turn_context.activity, 'reply_to_id'):
        reply.reply_to_id = turn_context.activity.id
    
    if hasattr(turn_context.activity, 'service_url'):
        reply.service_url = turn_context.activity.service_url
    
    if hasattr(turn_context.activity, 'channel_id'):
        reply.channel_id = turn_context.activity.channel_id
    
    await turn_context.send_activity(reply)

def create_email_result_card(email_text):
    """Creates an enhanced card displaying the generated email with preview and edit options"""
    
    # Extract subject line if present
    subject_match = re.search(r'Subject:\s*(.+?)(?:\n|$)', email_text)
    subject_line = subject_match.group(1) if subject_match else "Generated Email"
    
    # Extract recipient if present
    recipient_match = re.search(r'(?:To|Hi|Hello|Dear)\s+([^,\n]+)', email_text)
    recipient = recipient_match.group(1).strip() if recipient_match else "Recipient"
    
    # Create preview (first 150 chars of body after greeting)
    body_start = email_text.find('\n\n') + 2 if '\n\n' in email_text else 0
    # Skip the greeting line for preview
    first_para_end = email_text.find('\n', body_start)
    preview_start = first_para_end + 1 if first_para_end > -1 else body_start
    preview_text = email_text[preview_start:preview_start+150].strip()
    if len(email_text) > preview_start + 150:
        preview_text += "..."
    
    # Determine email type from signature
    email_type = "Client Service" if "Client Services Team" in email_text else "Sales"
    
    # Count key metrics
    word_count = len(email_text.split())
    paragraph_count = len([p for p in email_text.split('\n\n') if p.strip()])
    
    card = {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.5",
        "body": [
            {
                "type": "Container",
                "style": "emphasis",
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "✅",
                                        "size": "extraLarge",
                                        "color": "good"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Email Generated Successfully",
                                        "size": "large",
                                        "weight": "bolder",
                                        "color": "good"
                                    },
                                    {
                                        "type": "TextBlock",
                                        "text": f"{email_type} Email • {word_count} words • {paragraph_count} paragraphs",
                                        "size": "small",
                                        "isSubtle": True
                                    }
                                ]
                            }
                        ]
                    }
                ],
                "bleed": True
            },
            {
                "type": "Container",
                "style": "accent",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Email Preview",
                        "weight": "bolder",
                        "size": "medium"
                    },
                    {
                        "type": "FactSet",
                        "facts": [
                            {
                                "title": "To:",
                                "value": recipient
                            },
                            {
                                "title": "Subject:",
                                "value": subject_line
                            }
                        ]
                    },
                    {
                        "type": "TextBlock",
                        "text": preview_text,
                        "wrap": True,
                        "isSubtle": True,
                        "spacing": "small"
                    }
                ],
                "bleed": True
            },
            {
                "type": "Container",
                "style": "default",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Full Email Content",
                        "weight": "bolder",
                        "size": "medium",
                        "spacing": "medium"
                    },
                    {
                        "type": "TextBlock",
                        "text": email_text,
                        "wrap": True,
                        "spacing": "small"
                    }
                ],
                "padding": "Medium"
            },
            {
                "type": "Container",
                "style": "good",
                "items": [
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "width": "auto",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "✓",
                                        "color": "good",
                                        "weight": "bolder"
                                    }
                                ]
                            },
                            {
                                "type": "Column",
                                "width": "stretch",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "text": "Compliance checked and approved",
                                        "wrap": True,
                                        "size": "small",
                                        "color": "good"
                                    }
                                ]
                            }
                        ]
                    },
                    {
                        "type": "TextBlock",
                        "text": "This email follows FCDR guidelines and regulatory requirements.",
                        "wrap": True,
                        "size": "small",
                        "isSubtle": True,
                        "horizontalAlignment": "center"
                    }
                ],
                "spacing": "medium"
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "✏️ Edit This Email",
                "style": "positive",
                "data": {
                    "action": "edit_email"
                }
            },
            {
                "type": "Action.Submit",
                "title": "📧 Create Another Email",
                "data": {
                    "action": "view_change",
                    "view": "main"
                }
            },
            {
                "type": "Action.ShowCard",
                "title": "📋 Copy Instructions",
                "card": {
                    "type": "AdaptiveCard",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "How to use this email:",
                            "weight": "bolder"
                        },
                        {
                            "type": "TextBlock",
                            "text": "1. Review the email for accuracy\n2. Copy the email content\n3. Paste into your email client\n4. Personalize any [PLACEHOLDERS]\n5. Add any mentioned attachments\n6. Send to the recipient",
                            "wrap": True
                        }
                    ]
                }
            }
        ]
    }
    
    # Create attachment
    attachment = Attachment(
        content_type="application/vnd.microsoft.card.adaptive",
        content=card
    )
    
    return attachment
async def generate_email(turn_context: TurnContext, state, email_type, recipient=None, subject=None, instructions=None, context=None, has_attachments=False, skip_compliance_check=False):
    """
    Generates an email using AI with intelligent document retrieval based on email type.
    
    Args:
        turn_context: The turn context
        state: The conversation state
        email_type: Type of email (client_service or sales_service)
        recipient: The recipient's email (optional)
        subject: The email subject (optional)
        instructions: Instructions for email generation
        context: Additional context (optional)
        has_attachments: Whether to mention attachments
        skip_compliance_check: Skip the compliance check (default: False)
    """
    # Get thread ID early for state management
    thread_id = state.get("session_id")
    
    # Mark as busy BEFORE any async operations
    with conversation_states_lock:
        if state.get("active_run", False):
            # Already processing something
            await turn_context.send_activity("I'm currently processing another request. Please wait a moment and try again.")
            return
        state["active_run"] = True
        state["active_operation"] = "email_generation"
    
    if thread_id:
        with active_runs_lock:
            active_runs[thread_id] = True
    
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Build SMART retrieval query based on email type and instructions
        retrieval_query = ""
        
        # Analyze instructions to extract key topics
        instruction_keywords = ""
        if instructions:
            instruction_lower = instructions.lower()
            # Extract keywords based on common scenarios
            keyword_mapping = {
                # Customer Service Keywords
                "legal": "legal update legal threat lawsuit summons complaint legal action attorney",
                "settlement": "settlement missed payment lost settlement voided agreement negotiated savings",
                "payment": "payment returned draft reduction monthly payment gateway insufficient funds",
                "collection": "collection calls creditor contact harassment creditor notices",
                "credit": "credit score credit concerns credit report credit impact rebuilding credit",
                "timeline": "settlement timeline when resolved how long timeframe negotiation schedule",
                "cost": "program cost fees expensive afford monthly payment burden",
                "account": "account exclusion creditor jealousy exclude accounts keep accounts open",
                "welcome": "welcome new client enrollment congratulations program guide",
                
                # Sales Keywords
                "quote": "pre-approved quote debt relief quote monthly savings payment comparison",
                "analysis": "financial analysis debt situation credit utilization interest rates",
                "follow": "follow up previous conversation checking in next steps",
                "loan": "loan qualification loan option pre-qualification denied approval",
                "decision": "uncertain decision comparing options minimum payments consolidation",
                "program": "program overview debt resolution how it works benefits faster",
            }
            
            for key, keywords in keyword_mapping.items():
                if key in instruction_lower:
                    instruction_keywords += f" {keywords}"
        
        if email_type == "client_service":
            # For customer service, build comprehensive keyword query
            retrieval_query = "customer service email templates existing client support "
            retrieval_query += "welcome email legal update lost settlement legal confirmation "
            retrieval_query += "payment returned legal threat draft reduction creditor notices "
            retrieval_query += "collection calls credit concerns settlement timeline program cost "
            retrieval_query += "account exclusion client services team signature "
            retrieval_query += instruction_keywords
        else:  # sales_service
            # For sales, build comprehensive keyword query
            retrieval_query = "sales email templates pre-approved quote enrollment prospect "
            retrieval_query += "financial analysis debt consolidation program overview "
            retrieval_query += "follow-up decision uncertainty loan qualification "
            retrieval_query += "sales signature direct line monthly savings debt-free faster "
            retrieval_query += instruction_keywords
        
        # Add specific context to retrieval
        if subject:
            retrieval_query += f" {subject}"
        if context:
            retrieval_query += f" {context}"
        
        # Use RAG to retrieve relevant documents
        logging.info(f"Smart retrieval for {email_type} with enhanced query: {retrieval_query[:300]}...")
        relevant_docs = await retrieve_documents(retrieval_query, top=5)
        
        # Format the retrieved information
        retrieved_context = ""
        if relevant_docs and len(relevant_docs) > 0:
            retrieved_context = "\n\n--- RETRIEVED KNOWLEDGE ---\n\n"
            for i, doc in enumerate(relevant_docs, 1):
                if isinstance(doc, dict):
                    title = doc.get("title", "")
                    content = doc.get("content", "")
                    if title:
                        retrieved_context += f"{title}\n"
                    if content:
                        # Include substantial content for template matching
                        if len(content) > 4000:
                            content = content[:4000] + "..."
                        retrieved_context += f"{content}\n\n"
            logging.info(f"Retrieved {len(relevant_docs)} relevant documents for {email_type} email generation")
        
        # Create the prompt with routing intelligence
        email_category_text = "This is a Customer Service email request." if email_type == "client_service" else "This is a Sales email request."
        
        prompt = f"""{email_category_text}

You are the First Choice Debt Relief AI Assistant helping an employee draft an email. 

ROUTING CONTEXT: This is a {email_type.replace('_', ' ').upper()} department request.

Generate a professional, compliant email for First Choice Debt Relief based on the following requirements:

RECIPIENT: {recipient if recipient else '[To be determined]'}
SUBJECT: {subject if subject else '[Create an appropriate subject based on the instructions]'}
INSTRUCTIONS: {instructions if instructions else 'Create a professional email appropriate for the context.'}
ADDITIONAL CONTEXT: {context if context else 'None provided'}

CRITICAL INSTRUCTIONS:
1. Analyze the instructions to identify the email purpose (welcome, legal update, settlement issue, etc.)
2. Use the most appropriate email template structure from the retrieved knowledge
3. Personalize the template based on specific instructions
4. Write like a caring human colleague - vary language, show empathy, be conversational
5. Follow ALL compliance guidelines without exception
6. Use the EXACT signature format for the department

{retrieved_context}

DEPARTMENT-SPECIFIC REQUIREMENTS:
"""
        
        if email_type == "client_service":
            prompt += """
This is a CUSTOMER SERVICE communication. Requirements:
- Focus on supporting existing clients with their concerns
- Tone: Supportive, solution-oriented, reassuring but honest
- Key phrases: "actively working on your behalf", "need your approval", "here to support you"
- Common scenarios: legal updates, payment issues, settlement concerns, credit questions
- MANDATORY Signature format:
  Best regards,
  Client Services Team
  First Choice Debt Relief
  Phone: 800-985-9319
  Email: service@firstchoicedebtrelief.com
"""
        else:  # sales_service
            prompt += """
This is a SALES communication. Requirements:
- Focus on enrollment, quotes, and program benefits for prospects
- Tone: Optimistic but realistic, benefits-focused, non-pressuring
- Key phrases: "debt-free significantly faster", "lower monthly payments", "17+ years helping clients"
- Common scenarios: quotes, follow-ups, addressing concerns, loan options
- MANDATORY Signature format:
  Thank you,
  [YOUR_NAME]
  First Choice Debt Relief
  [YOUR_PHONE]
"""
        
        # Add attachment mention if required
        if has_attachments:
            prompt += "\n\nMention that there are attachments included in a natural way."
        
        # Add compliance guidelines
        prompt += """

COMPLIANCE REQUIREMENTS (MANDATORY - NEVER VIOLATE):
- NEVER promise guaranteed results or specific outcomes
- NEVER offer legal advice or use language suggesting legal expertise  
- NEVER use these prohibited terms: 'debt forgiveness', 'eliminate your debt', 'erase your debt'
- NEVER state or imply that the program prevents lawsuits or legal action
- NEVER claim all accounts will be resolved within a specific timeframe
- NEVER suggest the program is a credit repair service
- NEVER guarantee that clients will qualify for any financing
- NEVER make promises about improving credit scores
- NEVER say clients are 'required' to stop payments to creditors
- NEVER imply settlements are 'paid in full' - use 'negotiated resolution'
- NEVER represent FCDR as a government agency
- NEVER use pressure tactics like 'act immediately' or 'final notice'

HUMAN COMMUNICATION REQUIREMENTS:
- Write naturally - like a real person talking to someone who needs help
- Use contractions naturally (we're, you'll, I'll)
- Vary your language - don't repeat the same phrases
- Show genuine empathy: "I completely understand how frustrating that can be"
- Be conversational but professional
- Use warm, supportive language throughout

FORMAT REQUIREMENTS:
- Clear, descriptive subject line
- Warm greeting using recipient's name if provided
- Short paragraphs (3-5 sentences max)
- Bullet points for multiple items
- Clear next steps or call-to-action
- Exact department signature (no variations)
"""
        
        # Initialize chat if needed
        if not state.get("assistant_id"):
            await initialize_chat(turn_context, state)
            # Re-get thread_id after initialization
            thread_id = state.get("session_id")
        
        try:
            # Get client
            client = create_client()
            
            # Wait for any active runs to complete first
            if thread_id:
                wait_attempts = 0
                max_wait_attempts = 15  # 30 seconds total
                
                while wait_attempts < max_wait_attempts:
                    try:
                        runs = client.beta.threads.runs.list(thread_id=thread_id, limit=1)
                        if runs.data:
                            latest_run = runs.data[0]
                            if latest_run.status in ["in_progress", "queued", "requires_action"]:
                                logging.info(f"Waiting for active run {latest_run.id} to complete (attempt {wait_attempts + 1})")
                                await turn_context.send_activity(create_typing_activity())
                                await asyncio.sleep(2)
                                wait_attempts += 1
                                continue
                            elif latest_run.status in ["completed", "failed", "cancelled", "expired"]:
                                # Run is done, we can proceed
                                break
                        else:
                            # No runs, we can proceed
                            break
                    except Exception as e:
                        logging.warning(f"Error checking runs: {e}")
                        break
                
                if wait_attempts >= max_wait_attempts:
                    logging.warning("Timed out waiting for active run to complete")
            
            # Add the message
            try:
                client.beta.threads.messages.create(
                    thread_id=thread_id,
                    role="user",
                    content=prompt
                )
            except Exception as msg_error:
                logging.error(f"Error adding email generation message: {msg_error}")
                raise
            
            # Create and poll run
            run = client.beta.threads.runs.create(
                thread_id=thread_id,
                assistant_id=state["assistant_id"]
            )
            run_id = run.id
            logging.info(f"Created email generation run {run_id}")
            
            # Poll for completion
            max_wait = 120
            poll_interval = 2
            elapsed = 0
            email_text = ""
            last_typing_time = time.time()
            
            while elapsed < max_wait:
                # Send typing indicator periodically
                current_time = time.time()
                if current_time - last_typing_time > 5:
                    await turn_context.send_activity(create_typing_activity())
                    last_typing_time = current_time
                
                try:
                    run_status = client.beta.threads.runs.retrieve(
                        thread_id=thread_id,
                        run_id=run_id
                    )
                    
                    if run_status.status == "completed":
                        # Get the response
                        messages = client.beta.threads.messages.list(
                            thread_id=thread_id,
                            order="desc",
                            limit=1
                        )
                        
                        if messages.data:
                            for content in messages.data[0].content:
                                if content.type == 'text':
                                    email_text += content.text.value
                        
                        logging.info("Email generation completed successfully")
                        break
                        
                    elif run_status.status in ["failed", "cancelled", "expired"]:
                        error_msg = f"Email generation run ended with status: {run_status.status}"
                        logging.error(error_msg)
                        
                        # Try to get any partial response
                        try:
                            messages = client.beta.threads.messages.list(
                                thread_id=thread_id,
                                order="desc",
                                limit=1
                            )
                            if messages.data and messages.data[0].role == "assistant":
                                for content in messages.data[0].content:
                                    if content.type == 'text':
                                        email_text += content.text.value
                                if email_text:
                                    logging.info("Retrieved partial email response")
                                    break
                        except:
                            pass
                        
                        if not email_text:
                            raise Exception(error_msg)
                    
                    elif run_status.status == "requires_action":
                        # Handle tool calls if needed
                        logging.warning("Email generation run requires action - this shouldn't happen")
                        # For now, just wait
                    
                except Exception as poll_error:
                    logging.error(f"Error polling email generation run: {poll_error}")
                
                await asyncio.sleep(poll_interval)
                elapsed += poll_interval
            
            if not email_text:
                raise Exception("No email response generated after timeout")
            
            # COMPLIANCE CHECK (if not skipped)
            if not skip_compliance_check:
                logging.info("Running compliance check on generated email...")
                compliance_result = await check_email_compliance(email_text, email_type)
                
                if compliance_result and compliance_result.get("has_issues"):
                    logging.info(f"Compliance issues found: {compliance_result.get('suggestion', '')}")
                    
                    # Regenerate once with compliance feedback
                    edit_prompt = f"""The following email needs to be revised for compliance:

ORIGINAL EMAIL:
{email_text}

COMPLIANCE FEEDBACK: {compliance_result.get('suggestion', '')}

Please regenerate the email addressing the compliance concern while maintaining the warm, human tone and all original requirements.
Keep the same structure and intent, just fix the compliance issue mentioned."""
                    
                    # Add edit message
                    client.beta.threads.messages.create(
                        thread_id=thread_id,
                        role="user",
                        content=edit_prompt
                    )
                    
                    # Create revision run
                    revision_run = client.beta.threads.runs.create(
                        thread_id=thread_id,
                        assistant_id=state["assistant_id"]
                    )
                    
                    # Poll for revision completion
                    revision_elapsed = 0
                    while revision_elapsed < 60:  # 60 second timeout for revision
                        await turn_context.send_activity(create_typing_activity())
                        
                        revision_status = client.beta.threads.runs.retrieve(
                            thread_id=thread_id,
                            run_id=revision_run.id
                        )
                        
                        if revision_status.status == "completed":
                            # Get revised email
                            messages = client.beta.threads.messages.list(
                                thread_id=thread_id,
                                order="desc",
                                limit=1
                            )
                            
                            if messages.data:
                                revised_text = ""
                                for content in messages.data[0].content:
                                    if content.type == 'text':
                                        revised_text += content.text.value
                                if revised_text:
                                    email_text = revised_text
                                    logging.info("Email regenerated with compliance fixes")
                            break
                        elif revision_status.status in ["failed", "cancelled", "expired"]:
                            logging.warning(f"Revision run failed with status: {revision_status.status}")
                            break
                        
                        await asyncio.sleep(2)
                        revision_elapsed += 2
            
            # Save the generated email in the state for potential editing
            with conversation_states_lock:
                state["last_generated_email"] = email_text
                state["last_email_type"] = email_type
                state["last_email_data"] = {
                    "recipient": recipient,
                    "subject": subject,
                    "instructions": instructions,
                    "context": context,
                    "has_attachments": has_attachments
                }
                # Generate a unique email ID
                email_id = f"email_{int(time.time())}_{hashlib.md5(email_text.encode()).hexdigest()[:8]}"
                state["last_email_id"] = email_id
                
                # Store in email history
                if "email_history" not in state:
                    state["email_history"] = {}
                state["email_history"][email_id] = {
                    "email_text": email_text,
                    "email_type": email_type,
                    "timestamp": time.time(),
                    "data": state["last_email_data"]
                }
                
                # Keep only last 10 emails in history
                if len(state["email_history"]) > 10:
                    oldest_emails = sorted(state["email_history"].items(), key=lambda x: x[1]["timestamp"])[:len(state["email_history"]) - 10]
                    for old_id, _ in oldest_emails:
                        del state["email_history"][old_id]
            
            result_card = create_email_result_card(email_text)
            await send_card_response(turn_context, result_card)
            
        except Exception as e:
            logging.error(f"Error generating email: {str(e)}")
            traceback.print_exc()
            await turn_context.send_activity("I encountered an error while generating your email. Please try again.")
            
    finally:
        # Always clean up state
        with conversation_states_lock:
            state["active_run"] = False
            state.pop("active_operation", None)
        
        if thread_id:
            with active_runs_lock:
                if thread_id in active_runs:
                    del active_runs[thread_id]
async def check_email_compliance(email_text: str, email_type: str) -> dict:
    """
    Checks email for compliance issues using a dedicated compliance-focused GPT call.
    
    Args:
        email_text: The email to check
        email_type: Type of email (client_service or sales_service)
        
    Returns:
        dict: {"has_issues": bool, "suggestion": str or None}
    """
    try:
        client = create_client()
        
        # Build comprehensive compliance-specific system prompt
        compliance_system_prompt = """You are First Choice Debt Relief's Senior Compliance Officer AI, with deep expertise in debt relief regulations and company policies.

COMPANY CONTEXT: First Choice Debt Relief (FCDR) has 17+ years of experience helping clients become debt-free through negotiated settlements. We are NOT a law firm, credit repair company, or government agency.

YOUR ROLE: Review emails for CRITICAL compliance violations only. Minor style preferences are NOT compliance issues unless they create legal risk.

=== FEDERAL REGULATIONS TO ENFORCE ===

**Telemarketing Sales Rule (TSR)**
- VIOLATION: Collecting ANY fees before settlement is reached
- VIOLATION: Not disclosing 3-day cancellation right in sales emails
- VIOLATION: Not providing clear program terms before enrollment

**Federal Trade Commission (FTC)**
- VIOLATION: Deceptive or misleading claims about debt reduction
- VIOLATION: False advertising about program results
- VIOLATION: Not maintaining clear disclosure requirements

**Consumer Financial Protection Bureau (CFPB)**
- VIOLATION: Not being transparent about program risks
- VIOLATION: Guaranteeing debt elimination
- VIOLATION: Unfair, deceptive, or abusive practices (UDAAP)

**Fair Debt Collection Practices Act (FDCPA)**
- VIOLATION: False or misleading statements about settlements
- VIOLATION: Misrepresenting client's financial status
- VIOLATION: Using coercive or aggressive tactics

=== CRITICAL COMPLIANCE VIOLATIONS TO FLAG ===

**1. PROHIBITED GUARANTEES/PROMISES**
- "guarantee" + any specific outcome
- "promise" + any result
- "definitely will" + any claim
- "assured results"
- Specific percentage claims ("reduce by X%")
- Specific timeline promises ("resolved in X months")

**2. LEGAL ADVICE/EXPERTISE CLAIMS**
- "legal advice"
- "as your legal representative"
- "our attorneys will"
- "legal counsel"
- Interpreting laws or rights

**3. DEBT ELIMINATION LANGUAGE**
BANNED PHRASES:
- "debt forgiveness"
- "eliminate your debt" / "debt elimination"
- "erase your debt"
- "wipe out debt"
- "make debt disappear"

**4. LAWSUIT PREVENTION CLAIMS**
- "prevent lawsuits"
- "stop legal action"
- "protect from being sued"
- "lawsuit protection" (unless specifically about attorney cost coverage)
- "creditors can't sue you"

**5. CREDIT REPAIR CLAIMS**
- "fix your credit"
- "improve credit score"
- "credit repair"
- "restore credit"
- "rebuild credit" (unless carefully qualified)

**6. PAYMENT REQUIREMENTS**
- "required to stop paying"
- "must stop payments"
- "have to cease payments"
- Any mandatory language about stopping creditor payments

**7. SETTLEMENT DESCRIPTIONS**
- "paid in full" (MUST use "negotiated resolution")
- "satisfy debt completely"
- "full satisfaction"

**8. PRESSURE TACTICS**
- "act immediately"
- "final notice"
- "last chance"
- "urgent - respond now"
- "limited time" (unless factually true about quote expiration)

**9. MISREPRESENTATION**
- Claiming government affiliation
- Suggesting FCDR is a nonprofit
- Implying attorney-client relationship
- Misrepresenting program as loan consolidation

**10. SENSITIVE INFORMATION**
- Full SSN displayed
- Complete DOB
- Full account numbers
- Unredacted financial details

=== DEPARTMENT-SPECIFIC REQUIREMENTS ===

**CUSTOMER SERVICE EMAILS MUST HAVE:**
Signature format:
Best regards,
Client Services Team
First Choice Debt Relief
Phone: 800-985-9319
Email: service@firstchoicedebtrelief.com

VIOLATION: Using individual name or different phone number

**SALES EMAILS MUST HAVE:**
Signature format:
Thank you,
[AGENT NAME]
First Choice Debt Relief
[DIRECT PHONE]

VIOLATION: Using team signature or 800 number

=== TONE AND LANGUAGE COMPLIANCE ===

**REQUIRED ELEMENTS (Flag if Missing in Context):**
- Acknowledgment of client's situation/concern
- Clear next steps or call to action
- Appropriate empathy for serious situations
- Professional boundaries maintained

**ACCEPTABLE PHRASES TO ENCOURAGE:**
✓ "work to negotiate"
✓ "seek to achieve"
✓ "actively working on"
✓ "may be able to"
✓ "program designed to"
✓ "negotiated resolution"
✓ "significantly faster than minimum payments"

**TIME RESTRICTIONS:**
- Customer communications: 8am-8pm client local time only
- Flag any mention of contacting outside these hours

=== SPECIAL COMPLIANCE SCENARIOS ===

**Legal Situations:**
- MUST clarify: Legal insurance covers attorney costs, doesn't prevent lawsuits
- MUST NOT imply: Complete protection from legal action
- REQUIRED: Mention FCDR coordinates with legal providers

**Credit Discussions:**
- MUST focus: On debt resolution, not credit improvement
- ALLOWED: "Foundation for rebuilding" (not promises)
- MUST avoid: Any timeline for credit recovery

**Cost Concerns:**
- ALLOWED: Compare to long-term minimum payment costs
- MUST NOT: Claim "cheapest option"
- REQUIRED: Acknowledge program has costs

**Settlement Timelines:**
- MUST state: Varies by creditor and available funds
- MUST include: Client approval required for all settlements
- BANNED: Any specific timeframe promises

=== RESPONSE RULES ===

1. ONLY flag issues that create real legal/regulatory risk
2. Ignore stylistic preferences unless they affect compliance
3. One clear, actionable suggestion per issue
4. Focus on the MOST serious violation if multiple exist

RESPONSE FORMAT:
{"has_issues": true/false, "suggestion": "Specific fix" or null}

EXAMPLES:

INPUT: "We guarantee to reduce your debt by 50% in our program!"
OUTPUT: {"has_issues": true, "suggestion": "Replace with 'Our program is designed to help reduce your debt through negotiated settlements'"}

INPUT: "Our legal team will prevent any lawsuits against you."
OUTPUT: {"has_issues": true, "suggestion": "Change to 'Our legal insurance covers attorney costs if legal action occurs'"}

INPUT: "Thank you for calling. Based on what you shared, you may qualify for our debt resolution program. Thank you, John Smith, First Choice Debt Relief, 555-1234"
OUTPUT: {"has_issues": false, "suggestion": null}

INPUT: "Don't worry, we'll eliminate all your debt and fix your credit!"
OUTPUT: {"has_issues": true, "suggestion": "Replace with 'We work to help you resolve your debts through negotiated settlements'"}

INPUT: "You must stop paying your creditors immediately to enroll."
OUTPUT: {"has_issues": true, "suggestion": "Change to 'Many clients redirect their creditor payments into the program' to avoid mandating action"}

Be practical - focus on real compliance risks, not perfect wording."""
        
        # Determine department context
        dept_context = "Customer Service" if email_type == "client_service" else "Sales"
        
        # Create the user prompt
        user_prompt = f"""Review this {dept_context} email for compliance issues:

{email_text}

Department: {dept_context}
Check for critical compliance violations based on TSR, FTC, CFPB, FDCPA regulations and FCDR policies.
Return JSON format as specified."""
        
        # Make the compliance check call
        response = client.chat.completions.create(
            model="gpt-4.1-mini",
            messages=[
                {"role": "system", "content": compliance_system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            response_format={"type": "json_object"},
            temperature=0.1  # Low temperature for consistent compliance checking
        )
        
        # Parse the response
        result_text = response.choices[0].message.content
        result = json.loads(result_text)
        
        # Validate the response format
        if "has_issues" not in result:
            result["has_issues"] = False
        if "suggestion" not in result:
            result["suggestion"] = None
            
        logging.info(f"Compliance check result: {result}")
        return result
        
    except Exception as e:
        logging.error(f"Error in compliance check: {e}")
        # On error, assume no issues to avoid blocking email generation
        return {"has_issues": False, "suggestion": None}
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

async def bot_logic(turn_context: TurnContext):
    """
    Enhanced bot logic with comprehensive error boundaries and self-healing capabilities.
    """
    # Initialize error context
    error_context = {
        "stage": "initialization",
        "has_error": False,
        "error_details": None
    }
    
    try:
        # Stage 1: Basic initialization and validation
        error_context["stage"] = "basic_validation"
        
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
        
        # Stage 2: State management with error recovery
        error_context["stage"] = "state_management"
        
        # Thread-safe state initialization with error recovery
        state = None
        state_error = None
        
        try:
            with conversation_states_lock:
                if conversation_id not in conversation_states:
                    # Create new state for this conversation
                    conversation_states[conversation_id] = create_new_conversation_state(
                        user_id, tenant_id, user_security_fingerprint
                    )
                else:
                    # Update last activity time
                    conversation_states[conversation_id]["last_activity_time"] = time.time()
                    
                    # Verify user identity to prevent cross-contamination
                    stored_user_id = conversation_states[conversation_id].get("user_id")
                    stored_fingerprint = conversation_states[conversation_id].get("security_fingerprint")
                    
                    # If user mismatch detected, create fresh state
                    if stored_user_id and stored_user_id != user_id:
                        logging.warning(f"SECURITY: User mismatch in conversation {conversation_id}! Expected {stored_user_id}, got {user_id}")
                        
                        # Create fresh state to avoid cross-contamination
                        conversation_states[conversation_id] = create_new_conversation_state(
                            user_id, tenant_id, user_security_fingerprint
                        )
                        
                        # Clear any pending messages for security
                        with pending_messages_lock:
                            if conversation_id in pending_messages:
                                pending_messages[conversation_id].clear()
                        
                        logging.info(f"Created fresh state for user {user_id} after security check")
                
                state = conversation_states[conversation_id]
                
        except Exception as state_error:
            logging.error(f"State management error: {state_error}")
            # Create minimal fallback state
            state = create_new_conversation_state(user_id, tenant_id, user_security_fingerprint)
            error_context["has_error"] = True
            error_context["error_details"] = str(state_error)
        
        # Stage 3: Activity type handling with error boundaries
        error_context["stage"] = "activity_handling"
        
        # Check if we're in fallback mode
        if state.get("fallback_mode", False):
            # Use fallback response for all interactions
            await send_fallback_response(turn_context)
            return
        
        # Handle different activity types with individual error boundaries
        if turn_context.activity.type == ActivityTypes.message:
            await handle_message_activity_safe(turn_context, state, conversation_id, error_context)
            
        elif turn_context.activity.type == ActivityTypes.invoke:
            await handle_invoke_activity_safe(turn_context, state, error_context)
            
        elif turn_context.activity.type == ActivityTypes.conversation_update:
            await handle_conversation_update_safe(turn_context, state, error_context)
            
        elif turn_context.activity.type == ActivityTypes.event:
            # Handle custom events (like bot ready signals)
            logging.info(f"Received event activity: {turn_context.activity.name}")
            
        else:
            # Unknown activity type
            logging.warning(f"Unknown activity type: {turn_context.activity.type}")
            
    except Exception as critical_error:
        # Critical error handling - last resort
        logging.critical(f"Critical error in bot_logic at stage {error_context['stage']}: {critical_error}")
        traceback.print_exc()
        
        # Try to send fallback response
        try:
            await send_fallback_response(
                turn_context, 
                context={"error": str(critical_error), "stage": error_context["stage"]}
            )
        except Exception as fallback_error:
            logging.critical(f"Failed to send fallback response: {fallback_error}")
            
        # Try to at least acknowledge the error
        try:
            await turn_context.send_activity(
                "I encountered a critical error. Please try starting a new chat or contact support if this persists."
            )
        except:
            # Complete failure - just log
            logging.critical("Complete failure - unable to send any response to user")


def create_new_conversation_state(user_id: str, tenant_id: str, security_fingerprint: str) -> dict:
    """Create a new conversation state with all required fields"""
    return {
        "assistant_id": None,
        "session_id": None,
        "vector_store_id": None,
        "uploaded_files": [],
        "recovery_attempts": 0,
        "last_error": None,
        "error_history": [],
        "active_run": False,
        "user_id": user_id,
        "tenant_id": tenant_id,
        "security_fingerprint": security_fingerprint,
        "creation_time": time.time(),
        "last_activity_time": time.time(),
        "fallback_mode": False,
        "fallback_level": 0
    }


async def handle_message_activity_safe(turn_context: TurnContext, state: dict, conversation_id: str, error_context: dict):
    """Handle message activities with error boundaries"""
    try:
        # First, check if this is a card submission
        value_data = getattr(turn_context.activity, 'value', None)
        if value_data:
            logging.info(f"Card submission detected: {value_data}")
            try:
                await handle_card_actions(turn_context, value_data)
                return
            except Exception as card_error:
                logging.error(f"Error handling card submission: {card_error}")
                await handle_thread_recovery(turn_context, state, str(card_error), "card_submission")
                return
        
        # Initialize pending messages queue if not exists (thread-safe)
        with pending_messages_lock:
            if conversation_id not in pending_messages:
                pending_messages[conversation_id] = deque()
        
        # Check if we have text content
        has_text = turn_context.activity.text and turn_context.activity.text.strip()
        
        # Check for file attachments
        has_file_attachments = False
        file_caption = None
        
        if turn_context.activity.attachments and len(turn_context.activity.attachments) > 0:
            for attachment in turn_context.activity.attachments:
                if hasattr(attachment, 'content_type') and attachment.content_type == ContentType.FILE_DOWNLOAD_INFO:
                    has_file_attachments = True
                    if has_text:
                        file_caption = turn_context.activity.text.strip()
                    break
        
        # Check for session timeout
        if await check_session_timeout(state):
            await turn_context.send_activity("Your session has expired. Creating a new session...")
            await initialize_chat(turn_context, None)
            return
        
        # Track if thread is currently processing
        is_thread_busy = check_thread_busy(state)
        
        # Handle based on content type
        if is_thread_busy and has_text and not has_file_attachments:
            # Queue the message
            with pending_messages_lock:
                pending_messages[conversation_id].append(turn_context.activity.text.strip())
            await turn_context.send_activity("I'm still working on your previous request. I'll address this message next.")
            return
        
        # Process based on content
        if has_text and not has_file_attachments:
            try:
                await handle_text_message(turn_context, state)
            except Exception as text_error:
                logging.error(f"Error in handle_text_message: {text_error}")
                await handle_thread_recovery(turn_context, state, str(text_error), "text_message")
                
        elif has_file_attachments:
            try:
                await handle_file_upload(turn_context, state, file_caption)
            except Exception as file_error:
                logging.error(f"Error in handle_file_upload: {file_error}")
                await handle_thread_recovery(turn_context, state, str(file_error), "file_upload")
                
        else:
            # Empty message or unknown content
            logger.info(f"Received message without text or file attachments from user")
            
            if not state.get("assistant_id"):
                try:
                    await initialize_chat(turn_context, state)
                except Exception as init_error:
                    logging.error(f"Error in initialize_chat: {init_error}")
                    await handle_thread_recovery(turn_context, state, str(init_error), "initialization")
            else:
                await turn_context.send_activity(
                    "To upload files, use the paperclip icon and select from your device storage. "
                    "I support PDF, DOC, TXT, and image files."
                )
                
    except Exception as message_error:
        error_context["has_error"] = True
        error_context["error_details"] = str(message_error)
        logging.error(f"Error in message activity handling: {message_error}")
        await handle_thread_recovery(turn_context, state, str(message_error), "message_activity")


async def handle_invoke_activity_safe(turn_context: TurnContext, state: dict, error_context: dict):
    """Handle invoke activities with error boundaries"""
    try:
        if turn_context.activity.name == "fileConsent/invoke":
            await handle_file_consent_response(turn_context, turn_context.activity.value)
        elif turn_context.activity.name == "adaptiveCard/action":
            await handle_card_actions(turn_context, turn_context.activity.value)
        else:
            logging.warning(f"Unknown invoke activity: {turn_context.activity.name}")
            
    except Exception as invoke_error:
        error_context["has_error"] = True
        error_context["error_details"] = str(invoke_error)
        logging.error(f"Error in invoke activity handling: {invoke_error}")
        await send_fallback_response(turn_context, context={"activity": "invoke"})


async def handle_conversation_update_safe(turn_context: TurnContext, state: dict, error_context: dict):
    """Handle conversation update activities with error boundaries"""
    try:
        if turn_context.activity.members_added:
            for member in turn_context.activity.members_added:
                if member.id != turn_context.activity.recipient.id:
                    # Bot was added - send welcome message
                    await send_welcome_message(turn_context)
                    
    except Exception as update_error:
        error_context["has_error"] = True
        error_context["error_details"] = str(update_error)
        logging.error(f"Error in conversation update handling: {update_error}")
        # Non-critical error - just log


async def check_session_timeout(state: dict) -> bool:
    """Check if the session has timed out"""
    session_timeout = 1296000  # 15 days in seconds
    current_time = time.time()
    
    with conversation_states_lock:
        last_activity_time = state.get("last_activity_time", current_time)
        inactivity_period = current_time - last_activity_time
        
        return inactivity_period > session_timeout and state.get("session_id") is not None


def check_thread_busy(state: dict) -> bool:
    """Check if thread is currently processing"""
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
                    # State says active but active_runs doesn't have it
                    state["active_run"] = False
                    is_thread_busy = False
    
    return is_thread_busy

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
                
            await turn_context.send_activity(f"Received <b>{attachment.name}</b>")
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
    """Handle file uploads from Teams with clear messaging about supported types"""
    
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
                    # Provide clear message that only local uploads are supported
                    await turn_context.send_activity("I'm sorry, but I can only process files uploaded directly from your device. Files shared from OneDrive, SharePoint, or other internal sources are not currently supported. Please download the file to your device first, then upload it directly.")
                else:
                    # For other attachment types, provide general guidance
                    await turn_context.send_activity("Please ask questions regarding the file now")
                
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
    Will skip summarization during email workflows to prevent context loss.
    
    Args:
        client: Azure OpenAI client
        thread_id: The thread ID to check
        state: The conversation state dictionary
        threshold: Message count threshold before summarization (default: 30)
    
    Returns:
        bool: True if summarization was performed, False otherwise
    """
    try:
        # Check if we're in an email workflow - if so, skip summarization
        with conversation_states_lock:
            is_email_workflow = state.get("last_email_type") is not None
            if is_email_workflow:
                logging.info(f"Skipping summarization for thread {thread_id} - email workflow in progress")
                return False
        
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
        
        # Check if any recent messages are email-related before summarizing
        for msg in messages_list[-messages_to_keep:]:
            content_text = ""
            for content_part in msg.content:
                if content_part.type == 'text':
                    content_text += content_part.text.value
            
            # Skip if recent messages contain email generation
            if any(keyword in content_text.lower() for keyword in ["generate email", "email template", "draft email", "create email"]):
                logging.info(f"Skipping summarization - recent email-related activity detected")
                return False
        
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
            content=f"""Please create a concise but comprehensive summary of the following conversation. 
Focus on key points, decisions, and important context that would be needed for continuing the conversation effectively.
If any email templates or specific compliance requirements were discussed, preserve those details.

CONVERSATION TO SUMMARIZE:
{conversation_text}"""
        )
        
        # Run the summarization with a different assistant
        summary_run = client.beta.threads.runs.create(
            thread_id=summary_thread.id,
            assistant_id=state["assistant_id"],  # Use the same assistant
            instructions="Create a concise but comprehensive summary of the conversation provided. Focus on extracting key points, decisions, and important context that would be needed for continuing the conversation effectively. Format the summary in clear sections with bullet points where appropriate. Preserve any specific compliance requirements or email template discussions."
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
                    with conversation_states_lock:
                        state["session_id"] = new_thread.id
                        state["last_summarization_time"] = current_time
                        state["active_run"] = False
                        # Preserve email workflow state if any
                        if "last_email_type" in state:
                            state["summarization_preserved_email_state"] = {
                                "last_email_type": state["last_email_type"],
                                "last_generated_email": state.get("last_generated_email"),
                                "last_email_data": state.get("last_email_data")
                            }
                    
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
# Create a new function to initialize chat without any messages or welcome
async def initialize_chat_silent(turn_context: TurnContext, state):
    """Initialize a new chat session without sending welcome messages"""
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_id = conversation_reference.conversation.id
    user_id = turn_context.activity.from_property.id if hasattr(turn_context.activity, 'from_property') else None
    
    try:
        # Send typing indicator
        await turn_context.send_activity(create_typing_activity())
        
        # Create a client
        client = create_client()
        
        # Create a vector store
        vector_store = client.vector_stores.create(
            name=f"user_{user_id}_convo_{conversation_id}_{int(time.time())}"
        )
        
        # Include file_search tool
        assistant_tools = [{"type": "file_search"}]
        assistant_tool_resources = {
            "file_search": {"vector_store_ids": [vector_store.id]}
        }

        # Create the assistant with a unique name
        unique_name = f"pm_copilot_user_{user_id}_convo_{conversation_id}_{int(time.time())}"
        assistant = client.beta.assistants.create(
            name=unique_name,
            model="gpt-4.1-mini",
            instructions=SYSTEM_PROMPT,
            tools=assistant_tools,
            tool_resources=assistant_tool_resources,
        )
        
        # Create a thread
        thread = client.beta.threads.create()
        
        # Update state with new resources
        with conversation_states_lock:
            state["assistant_id"] = assistant.id
            state["session_id"] = thread.id
            state["vector_store_id"] = vector_store.id
            state["active_run"] = False
            state["recovery_attempts"] = 0
            state["user_identifier"] = f"{conversation_id}_{user_id}"
            state["creation_time"] = time.time()
            state["last_activity_time"] = time.time()
        
        return True
    except Exception as e:
        logging.error(f"Error in initialize_chat_silent: {e}")
        return False
async def format_message_with_rag(user_message, documents):
    """
    Format a message combining user query with retrieved text knowledge.
    Returns original message if an error occurs.
    """
    try:
        formatted_message = user_message
        
        # Only add context if we have relevant documents
        if documents and len(documents) > 0:
            # Create the context section
            context = "\n\n--- RETRIEVED KNOWLEDGE ---\n\n"
            logging.info(f"RAG: Found {len(documents)} relevant documents")
            
            # Add document content
            for i, doc in enumerate(documents, 1):
                # Handle both dictionary and non-dictionary items
                if isinstance(doc, dict):
                    title = doc.get("title", "")
                    content = doc.get("content", "")
                else:
                    # Fallback if doc is not a dictionary
                    title = ""
                    content = str(doc)
                
                # Add the title (source filename) without "DOCUMENT X:" prefix
                if title:
                    context += f"{title}\n"
                
                # Smart truncation - show up to 5000 chars but try to break at a sentence
                if len(content) > 5000:
                    # Find the last period within the first 5000 chars
                    last_period = content[:5000].rfind('.')
                    if last_period > 0:
                        content = content[:last_period+1] + " [content continues...]"
                    else:
                        content = content[:5000] + " [content continues...]"
                
                context += f"{content}\n\n"
                logging.info(f"RAG Document {i}: {title} - {content[:100]}...")
                
            # Add the combined message
            formatted_message = f"{formatted_message}\n\n{context}"
            
        logging.info(f"COMPLETE RAG MESSAGE: {formatted_message[:500]}... [message continues, total length: {len(formatted_message)}]")
        return formatted_message
        
    except Exception as e:
        # Log the error but don't break the conversation flow
        logging.error(f"Error formatting RAG message: {e}")
        
        # Fallback: Just append the raw documents without any parsing
        try:
            fallback_message = f"{user_message}\n\n--- RETRIEVED KNOWLEDGE ---\n\n"
            fallback_message += str(documents)
            logging.info("Using fallback: appending raw documents without parsing")
            return fallback_message
        except:
            # Ultimate fallback - return original message
            return user_message
async def handle_text_message(turn_context: TurnContext, state):
    """Handle text messages from users with RAG integration"""
    user_message = turn_context.activity.text.strip()
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_id = conversation_reference.conversation.id
    
    # Handle special commands
    if user_message.lower() in ["/email", "create email", "write email", "email template", "email"]:
        # Check if busy before showing email card
        if check_thread_busy(state):
            await turn_context.send_activity("I'm currently processing another request. Please wait a moment and try again.")
            
            # Queue this command for later
            with pending_messages_lock:
                if conversation_id not in pending_messages:
                    pending_messages[conversation_id] = deque()
                pending_messages[conversation_id].append(user_message)
            return
            
        await send_email_card(turn_context, "main")
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
        # Initialize chat silently first
        success = await initialize_chat_silent(turn_context, state)
        
        if success:
            # Now process the user's message directly
            with conversation_states_lock:
                stored_assistant_id = state.get("assistant_id")
                stored_session_id = state.get("session_id")
            
            # Process the message without sending welcome messages
            client = create_client()
            
            # RAG INTEGRATION - RETRIEVE RELEVANT DOCUMENTS
            relevant_docs = await retrieve_documents(user_message, top=3)
            
            # Format the message with RAG context and user query
            enhanced_message = await format_message_with_rag(user_message, relevant_docs)
            
            # Send the enhanced message
            client.beta.threads.messages.create(
                thread_id=stored_session_id,
                role="user",
                content=enhanced_message
            )
            
            # Process the message with streaming
            if TEAMS_AI_AVAILABLE:
                await stream_with_teams_ai(turn_context, state, None)
            else:
                await stream_with_custom_implementation(turn_context, state, None)
        else:
            # Fallback if initialization failed
            await turn_context.send_activity("I'm sorry, I encountered an issue while setting up our conversation. Please try again.")
        
        return
    
    # Track if thread is currently processing
    is_thread_busy = check_thread_busy(state)
    
    # Handle based on busy state
    if is_thread_busy:
        # Get the active operation type for better messaging
        active_operation = state.get("active_operation", "request")
        
        # Queue the message
        with pending_messages_lock:
            if conversation_id not in pending_messages:
                pending_messages[conversation_id] = deque()
            pending_messages[conversation_id].append(user_message)
            queue_size = len(pending_messages[conversation_id])
        
        # Provide informative message based on what's happening
        if active_operation == "email_generation":
            message = f"I'm currently generating an email. Your message has been queued (position {queue_size}). I'll respond as soon as I'm done."
        else:
            message = f"I'm still processing your previous request. Your message has been queued (position {queue_size}). I'll respond as soon as I'm done."
        
        await turn_context.send_activity(message)
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
        state["active_operation"] = "message_processing"
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
        
        # Wait for any active runs to complete first
        wait_attempts = 0
        max_wait_attempts = 10  # 20 seconds total
        
        while wait_attempts < max_wait_attempts:
            try:
                runs = client.beta.threads.runs.list(thread_id=current_session_id, limit=1)
                if runs.data:
                    latest_run = runs.data[0]
                    if latest_run.status in ["in_progress", "queued", "requires_action"]:
                        logging.info(f"Waiting for active run {latest_run.id} to complete before processing message")
                        await turn_context.send_activity(create_typing_activity())
                        await asyncio.sleep(2)
                        wait_attempts += 1
                        continue
                    else:
                        # Run is done
                        break
                else:
                    # No runs
                    break
            except Exception as e:
                logging.warning(f"Error checking runs: {e}")
                break
        
        # RAG INTEGRATION - RETRIEVE RELEVANT DOCUMENTS
        relevant_docs = await retrieve_documents(user_message, top=3)
        
        # Format the message with RAG context and user query
        enhanced_message = await format_message_with_rag(user_message, relevant_docs)
        
        # Send the enhanced message
        try:
            client.beta.threads.messages.create(
                thread_id=current_session_id,
                role="user",
                content=enhanced_message
            )
        except Exception as msg_error:
            logging.error(f"Error adding message to thread: {msg_error}")
            raise
            
        # Use the optimal streaming approach based on available libraries and preferences
        if TEAMS_AI_AVAILABLE:
            # Use enhanced streaming with Teams AI library
            await stream_with_teams_ai(turn_context, state, None)
        else:
            # Use custom TeamsStreamingResponse if Teams AI library is not available
            await stream_with_custom_implementation(turn_context, state, None)
        
        # Mark thread as no longer busy (thread-safe)
        with conversation_states_lock:
            state["active_run"] = False
            state.pop("active_operation", None)
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
            state.pop("active_operation", None)
            current_session_id = state.get("session_id")
            
        with active_runs_lock:
            if current_session_id in active_runs:
                del active_runs[current_session_id]
            
        # Don't show raw error details to users
        logging.error(f"Error in handle_text_message for user {user_id}: {str(e)}")
        traceback.print_exc()
        
        # Handle thread recovery if needed
        await handle_thread_recovery(turn_context, state, str(e), "text_message")


# Modified process_pending_messages function to fix the run conflict
async def process_pending_messages(turn_context: TurnContext, state, conversation_id):
    """Process any pending messages in the queue safely with improved handling"""
    messages_to_process = []
    
    # Get all pending messages at once to avoid race conditions
    with pending_messages_lock:
        if conversation_id in pending_messages and pending_messages[conversation_id]:
            # Get all messages and clear the queue
            messages_to_process = list(pending_messages[conversation_id])
            pending_messages[conversation_id].clear()
            logging.info(f"Processing {len(messages_to_process)} pending messages for conversation {conversation_id}")
    
    # If no messages to process, return
    if not messages_to_process:
        return
    
    # Process messages outside the lock to avoid blocking
    for i, next_message in enumerate(messages_to_process):
        try:
            # Announce processing if first message
            if i == 0:
                await turn_context.send_activity("I'll now address your follow-up messages...")
            
            # Add small delay between messages to avoid overwhelming the system
            if i > 0:
                await asyncio.sleep(1.5)
            
            # Get the thread and assistant IDs
            thread_id = state.get("session_id")
            assistant_id = state.get("assistant_id")
            
            if not thread_id or not assistant_id:
                await turn_context.send_activity(f"I'm having trouble with your follow-up question #{i+1}. Let's start a new conversation.")
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
                        logging.info(f"Cancelling active run {latest_run.id} before processing follow-up #{i+1}")
                        client.beta.threads.runs.cancel(thread_id=thread_id, run_id=latest_run.id)
                        await asyncio.sleep(2)  # Wait for cancellation to take effect
            except Exception as cancel_e:
                logging.warning(f"Error checking or cancelling runs for follow-up #{i+1}: {cancel_e}")
            
            # Wait to ensure no active runs
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
                                logging.info(f"Still waiting for run {run.id} to complete before follow-up #{i+1}...")
                                await asyncio.sleep(1)
                                break
                    
                    if not active_run_found:
                        break
                except Exception:
                    break  # If we can't check, just proceed
            
            # Add progress indicator for multiple messages
            progress_indicator = f" ({i+1}/{len(messages_to_process)})" if len(messages_to_process) > 1 else ""
            
            # Check if this is an email-related follow-up
            is_email_followup = any(keyword in next_message.lower() for keyword in ["email", "change", "edit", "modify", "update"])
            
            # Add the follow-up message to the thread
            try:
                # If it's an email follow-up and we have email context, include it
                message_content = next_message
                if is_email_followup and state.get("last_generated_email"):
                    message_content = f"{next_message}\n\n[Context: This relates to the previously generated email]"
                
                client.beta.threads.messages.create(
                    thread_id=thread_id,
                    role="user",
                    content=message_content
                )
                
                logging.info(f"Added follow-up message #{i+1} to thread {thread_id}")
            except Exception as msg_error:
                logging.error(f"Error adding follow-up message #{i+1}: {msg_error}")
                await turn_context.send_activity(f"I couldn't process follow-up message #{i+1}. Please try asking again.")
                continue
            
            # Process the response with streaming
            try:
                # Send processing indicator
                if len(messages_to_process) > 1:
                    await turn_context.send_activity(f"Processing your message{progress_indicator}...")
                
                if TEAMS_AI_AVAILABLE:
                    await stream_with_teams_ai(turn_context, state, None)
                else:
                    await stream_with_custom_implementation(turn_context, state, None)
                
                # Brief pause before next message
                if i < len(messages_to_process) - 1:
                    await asyncio.sleep(0.5)
                    
            except Exception as process_error:
                logging.error(f"Error processing follow-up #{i+1}: {process_error}")
                await turn_context.send_activity(f"I had trouble processing follow-up message #{i+1}. Please try asking again.")
                
        except Exception as e:
            logging.error(f"Error processing follow-up message #{i+1}: {e}")
            traceback.print_exc()
            await turn_context.send_activity(f"I encountered an error with message #{i+1}. Please try asking again.")
            
            # If we encounter an error, we might want to continue with remaining messages
            # or stop processing based on severity
            if i < len(messages_to_process) - 1:
                await turn_context.send_activity("I'll continue with your remaining messages...")
                await asyncio.sleep(1)
    
    # Final message if we processed multiple messages
    if len(messages_to_process) > 1:
        await turn_context.send_activity("I've finished processing all your follow-up messages. Is there anything else I can help you with?")
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
                        
                            # Fetch the assistant’s most-recent message (same logic as before)
                            messages = client.beta.threads.messages.list(
                                thread_id=thread_id,
                                order="desc",
                                limit=1
                            )
                        
                            if messages.data:
                                latest_message = messages.data[0]
                                message_text = ""
                        
                                for content_part in latest_message.content:
                                    if content_part.type == "text":
                                        message_text += content_part.text.value
                        
                                # ── NO DUPLICATES ─────────────────────────────────────────────
                                streamer.replace_buffer_with(
                                    message_text or "I couldn't generate a response. Please try again."
                                )
                                await streamer.send_final_message()   # flush exactly once
                                return                                # done with this user request
                                # ──────────────────────────────────────────────────────────────
                        
                            # Fallback if, for some reason, there was no text message
                            streamer.replace_buffer_with(
                                "I didn't receive a valid response. Please try again."
                            )
                            await streamer.send_final_message()
                            return

                            
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
                # ── NO DUPLICATES ────────────────────────────────────────────
                streamer.replace_buffer_with(message_text)  # new helper
                await streamer.send_final_message()
                return
        
        streamer.replace_buffer_with(
            "I processed your request but couldn't generate a proper response. Please try again."
        )
        await streamer.send_final_message()
        
    except Exception as e:
        logging.error(f"Error in poll_for_message: {e}")
        streamer.replace_buffer_with(
            "I encountered an error while retrieving the response. Please try again."
        )
        await streamer.send_final_message()

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
            await send_fallback_response(turn_context, context or "How can I help you today?")
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
            await send_fallback_response(turn_context, "Hello, how can I help you today?")
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
        if prompt:
            # Wait for any active runs to complete
            wait_attempts = 0
            max_wait_attempts = 15  # 30 seconds total
            active_run_id = None
            
            while wait_attempts < max_wait_attempts:
                try:
                    runs = client.beta.threads.runs.list(thread_id=session, limit=1)
                    if runs.data:
                        latest_run = runs.data[0]
                        if latest_run.status in ["in_progress", "queued", "requires_action"]:
                            active_run_id = latest_run.id
                            logging.info(f"Found active run {active_run_id} with status {latest_run.status}, waiting... (attempt {wait_attempts + 1})")
                            await asyncio.sleep(2)
                            wait_attempts += 1
                            continue
                        else:
                            # Run is completed/failed/cancelled
                            logging.info(f"Previous run {latest_run.id} has status {latest_run.status}, proceeding")
                            break
                    else:
                        # No runs found
                        break
                except Exception as e:
                    logging.warning(f"Error checking for active runs: {e}")
                    break
            
            # If we still have an active run after waiting, try to cancel it
            if wait_attempts >= max_wait_attempts and active_run_id:
                logging.warning(f"Active run {active_run_id} still running after {wait_attempts * 2}s")
                try:
                    client.beta.threads.runs.cancel(thread_id=session, run_id=active_run_id)
                    logging.info(f"Requested cancellation of run {active_run_id}")
                    
                    # Wait a bit more for cancellation
                    cancel_wait = 0
                    while cancel_wait < 10:  # 10 seconds for cancellation
                        await asyncio.sleep(2)
                        cancel_wait += 2
                        try:
                            run_check = client.beta.threads.runs.retrieve(thread_id=session, run_id=active_run_id)
                            if run_check.status in ["cancelled", "failed", "expired", "completed"]:
                                logging.info(f"Run {active_run_id} now has status {run_check.status}")
                                break
                        except:
                            break
                            
                except Exception as cancel_e:
                    logging.error(f"Failed to cancel run {active_run_id}: {cancel_e}")
                    # As last resort, create a new thread
                    try:
                        thread = client.beta.threads.create()
                        old_session = session
                        session = thread.id
                        logging.info(f"Created new thread {session} to replace busy thread {old_session}")
                    except Exception as thread_e:
                        logging.error(f"Failed to create new thread: {thread_e}")
                        raise HTTPException(status_code=500, detail="Thread is busy and cannot create new thread")
            
            # Now add the message
            max_retries = 3
            message_added = False
            
            for retry in range(max_retries):
                try:
                    client.beta.threads.messages.create(
                        thread_id=session,
                        role="user",
                        content=prompt
                    )
                    logging.info(f"Added user message to thread {session}")
                    message_added = True
                    break
                except Exception as e:
                    error_msg = str(e)
                    if "while a run" in error_msg or "active run" in error_msg:
                        if retry < max_retries - 1:
                            logging.warning(f"Thread still busy, waiting before retry {retry + 2}/{max_retries}")
                            await asyncio.sleep(3 * (retry + 1))  # Exponential backoff
                        else:
                            # Final attempt - create new thread
                            try:
                                thread = client.beta.threads.create()
                                old_session = session
                                session = thread.id
                                client.beta.threads.messages.create(
                                    thread_id=session,
                                    role="user",
                                    content=prompt
                                )
                                message_added = True
                                logging.info(f"Created new thread {session} and added message after retries")
                            except Exception as new_e:
                                logging.error(f"Failed to create new thread and add message: {new_e}")
                                raise HTTPException(status_code=500, detail="Failed to process message after multiple attempts")
                    else:
                        logging.error(f"Error adding message: {e}")
                        if retry == max_retries - 1:
                            raise
            
            if not message_added:
                raise HTTPException(status_code=500, detail="Failed to add message to conversation")
        
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
                        # Fallback to polling approach
                        logging.info(f"Direct streaming not available: {stream_not_available}. Using polling approach")
                        
                        # Create run without streaming
                        run = client.beta.threads.runs.create(
                            thread_id=session,
                            assistant_id=assistant
                        )
                        
                        run_id = run.id
                        logging.info(f"Created polling run {run_id}")
                        
                        # Poll for completion
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
                                        
                                        # Split into chunks for better streaming experience
                                        if len(message_text) > 500:
                                            # Use sentence-aware chunking
                                            sentences = message_text.split('. ')
                                            current_chunk = ""
                                            
                                            for sentence in sentences:
                                                current_chunk += sentence + '. '
                                                
                                                if len(current_chunk) >= 200:
                                                    yield current_chunk
                                                    current_chunk = ""
                                                    await asyncio.sleep(0.05)
                                            
                                            # Yield any remaining text
                                            if current_chunk:
                                                yield current_chunk
                                        else:
                                            # For shorter responses, yield the whole thing
                                            yield message_text
                                    break
                                
                                elif run_status.status in ["failed", "cancelled", "expired"]:
                                    error_msg = f"Run ended with status {run_status.status}"
                                    
                                    # Try to get any partial response
                                    try:
                                        messages = client.beta.threads.messages.list(
                                            thread_id=session,
                                            order="desc",
                                            limit=1
                                        )
                                        if messages.data and messages.data[0].role == "assistant":
                                            partial_text = ""
                                            for content_part in messages.data[0].content:
                                                if content_part.type == 'text':
                                                    partial_text += content_part.text.value
                                            if partial_text:
                                                yield partial_text
                                                return
                                    except:
                                        pass
                                    
                                    yield f"\nError: {error_msg}. Please try again."
                                    break
                                
                                elif run_status.status == "requires_action":
                                    yield "\n[This response requires additional actions which cannot be handled in the current mode.]\n"
                                    break
                                
                                await asyncio.sleep(wait_interval)
                                elapsed_time += wait_interval
                                
                            except Exception as poll_e:
                                logging.error(f"Error polling run status: {poll_e}")
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
                        
                        if attempt % 6 == 0:  # Log every 30 seconds
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
                            
                            # Try to get partial response
                            try:
                                messages = client.beta.threads.messages.list(
                                    thread_id=session,
                                    order="desc",
                                    limit=1
                                )
                                if messages.data and messages.data[0].role == "assistant":
                                    for content_part in messages.data[0].content:
                                        if content_part.type == 'text':
                                            full_response += content_part.text.value
                                    if full_response:
                                        return {"response": full_response}
                            except:
                                pass
                            
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
                traceback.print_exc()
                return {
                    "response": "An error occurred while processing your request. Please try again."
                }
        
    except Exception as e:
        endpoint_type = "conversation" if stream_output else "chat"
        logging.error(f"Error in /{endpoint_type} endpoint setup: {e}")
        traceback.print_exc()
        
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
    return {"status": "ok", "service": "Teams AI Assistant"}

# Root path redirect to health
@app.get("/")
async def root():
    return {"status": "ok", "message": "Teams AI Assistant is running."}

# Run the app with uvicorn if executed directly
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    print(f"Starting FastAPI server on http://0.0.0.0:{port}")
    uvicorn.run(app, host="0.0.0.0", port=port)
