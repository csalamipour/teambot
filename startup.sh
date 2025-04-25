#!/bin/bash

# Make sure script is executable:
# chmod +x startup.sh

# This is the startup script for Azure App Service
# It will install dependencies and start the bot

echo "Starting Teams Bot deployment process..."

# Install any dependencies
echo "Installing dependencies..."
pip install -r requirements.txt

# Start the Python application
echo "Starting bot application..."
python teams_bot.py
