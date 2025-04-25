#!/bin/bash
# This script helps with deploying your Teams bot to Azure App Service

# Make sure you're logged in to Azure CLI first:
# az login

# Set variables (change these to your values)
RESOURCE_GROUP="copilot-updated"
APP_SERVICE_NAME="product-management-bot"
LOCATION="eastus"  # or your preferred region
PYTHON_VERSION="3.9"

# Create resource group if it doesn't exist
echo "Creating or ensuring resource group exists..."
az group create --name $RESOURCE_GROUP --location $LOCATION

# Create App Service Plan
echo "Creating App Service Plan..."
az appservice plan create --name "${APP_SERVICE_NAME}-plan" \
                         --resource-group $RESOURCE_GROUP \
                         --sku B1 \
                         --is-linux

# Create Web App
echo "Creating Web App..."
az webapp create --name $APP_SERVICE_NAME \
                --resource-group $RESOURCE_GROUP \
                --plan "${APP_SERVICE_NAME}-plan" \
                --runtime "PYTHON:${PYTHON_VERSION}"

# Configure startup command
echo "Setting startup command..."
az webapp config set --name $APP_SERVICE_NAME \
                    --resource-group $RESOURCE_GROUP \
                    --startup-file "startup.sh"

# Configure app settings (environment variables)
echo "Setting environment variables..."
az webapp config appsettings set --name $APP_SERVICE_NAME \
                                --resource-group $RESOURCE_GROUP \
                                --settings MicrosoftAppId="your-app-id" \
                                          MicrosoftAppPassword="your-app-password" \
                                          PORT="8080" \
                                          PYTHONPATH="."

echo "Deployment configuration complete!"
echo "Now set up GitHub Actions for continuous deployment"
echo "or use: az webapp deployment source config --name $APP_SERVICE_NAME --resource-group $RESOURCE_GROUP --repo-url YOUR_GITHUB_REPO --branch main --manual-integration"
