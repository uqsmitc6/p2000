#!/bin/bash
# =============================================================
# Deploy p2001 (UQ Slide Converter) to Azure App Service
# =============================================================
#
# Prerequisites:
#   - Azure CLI installed (brew install azure-cli)
#   - Logged in (az login)
#   - Docker installed and running
#
# This script:
#   1. Creates an Azure Container Registry (ACR)
#   2. Builds and pushes the Docker image
#   3. Creates a Web App on your existing App Service plan
#   4. Configures environment variables
#
# Usage:
#   chmod +x deploy-azure.sh
#   ./deploy-azure.sh
#
# After first deploy, to UPDATE the app with code changes:
#   ./deploy-azure.sh --update
# =============================================================

set -e

# --- Configuration ---
RESOURCE_GROUP="bschool-learning-tools"
APP_SERVICE_PLAN="ASP-bschoollearningtools-83d7"
LOCATION="australiaeast"
APP_NAME="p2001"
ACR_NAME="bschoollearningtools"   # Must be globally unique, alphanumeric only

# Docker image tag
IMAGE_TAG="${ACR_NAME}.azurecr.io/${APP_NAME}:latest"

# --- Colours for output ---
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No colour

echo -e "${GREEN}=== p2001 Azure Deployment ===${NC}"
echo "Resource group: ${RESOURCE_GROUP}"
echo "App Service plan: ${APP_SERVICE_PLAN}"
echo "App name: ${APP_NAME}"
echo ""

# --- Check if this is an update-only run ---
if [[ "$1" == "--update" ]]; then
    echo -e "${YELLOW}Update mode: rebuilding and pushing image only${NC}"

    # Login to ACR
    az acr login --name "${ACR_NAME}"

    # Build and push
    echo -e "${GREEN}Building Docker image...${NC}"
    docker build -t "${IMAGE_TAG}" .

    echo -e "${GREEN}Pushing to ACR...${NC}"
    docker push "${IMAGE_TAG}"

    # Restart the app to pull the new image
    echo -e "${GREEN}Restarting app...${NC}"
    az webapp restart --name "${APP_NAME}" --resource-group "${RESOURCE_GROUP}"

    echo -e "${GREEN}Done! App will be live in ~60 seconds at:${NC}"
    echo "https://${APP_NAME}.azurewebsites.net"
    exit 0
fi

# --- Step 1: Create ACR (if it doesn't exist) ---
echo -e "${GREEN}Step 1: Setting up Container Registry...${NC}"
if az acr show --name "${ACR_NAME}" --resource-group "${RESOURCE_GROUP}" &>/dev/null; then
    echo "  ACR '${ACR_NAME}' already exists — skipping"
else
    echo "  Creating ACR '${ACR_NAME}'..."
    az acr create \
        --name "${ACR_NAME}" \
        --resource-group "${RESOURCE_GROUP}" \
        --sku Basic \
        --location "${LOCATION}" \
        --admin-enabled true
fi

# Login to ACR
az acr login --name "${ACR_NAME}"

# --- Step 2: Build and push Docker image ---
echo -e "${GREEN}Step 2: Building and pushing Docker image...${NC}"
docker build -t "${IMAGE_TAG}" .
docker push "${IMAGE_TAG}"

# --- Step 3: Get ACR credentials ---
echo -e "${GREEN}Step 3: Retrieving ACR credentials...${NC}"
ACR_USERNAME=$(az acr credential show --name "${ACR_NAME}" --query "username" -o tsv)
ACR_PASSWORD=$(az acr credential show --name "${ACR_NAME}" --query "passwords[0].value" -o tsv)
ACR_SERVER="${ACR_NAME}.azurecr.io"

# --- Step 4: Create Web App ---
echo -e "${GREEN}Step 4: Creating Web App '${APP_NAME}'...${NC}"
if az webapp show --name "${APP_NAME}" --resource-group "${RESOURCE_GROUP}" &>/dev/null; then
    echo "  Web App '${APP_NAME}' already exists — updating container"
    az webapp config container set \
        --name "${APP_NAME}" \
        --resource-group "${RESOURCE_GROUP}" \
        --container-image-name "${IMAGE_TAG}" \
        --container-registry-url "https://${ACR_SERVER}" \
        --container-registry-user "${ACR_USERNAME}" \
        --container-registry-password "${ACR_PASSWORD}"
else
    az webapp create \
        --name "${APP_NAME}" \
        --resource-group "${RESOURCE_GROUP}" \
        --plan "${APP_SERVICE_PLAN}" \
        --container-image-name "${IMAGE_TAG}" \
        --container-registry-url "https://${ACR_SERVER}" \
        --container-registry-user "${ACR_USERNAME}" \
        --container-registry-password "${ACR_PASSWORD}"
fi

# --- Step 5: Configure app settings ---
echo -e "${GREEN}Step 5: Configuring app settings...${NC}"
az webapp config appsettings set \
    --name "${APP_NAME}" \
    --resource-group "${RESOURCE_GROUP}" \
    --settings \
        WEBSITES_PORT=8501 \
        DOCKER_ENABLE_CI=false

# --- Step 6: Set secrets (prompt user) ---
echo ""
echo -e "${YELLOW}=== Secrets Configuration ===${NC}"
echo "You need to set these secrets in the Azure Portal:"
echo "  1. Go to: https://portal.azure.com"
echo "  2. Navigate to: App Services → ${APP_NAME} → Settings → Configuration"
echo "  3. Add these Application Settings:"
echo ""
echo "     ANTHROPIC_API_KEY = (your Anthropic API key)"
echo "     GOOGLE_SHEETS_WEBHOOK_URL = (your Apps Script webhook URL)"
echo ""
echo "  Or set them via CLI:"
echo "     az webapp config appsettings set --name ${APP_NAME} --resource-group ${RESOURCE_GROUP} --settings ANTHROPIC_API_KEY='sk-ant-...' GOOGLE_SHEETS_WEBHOOK_URL='https://script.google.com/...'"
echo ""

# --- Step 7: Enable logging ---
echo -e "${GREEN}Step 6: Enabling container logging...${NC}"
az webapp log config \
    --name "${APP_NAME}" \
    --resource-group "${RESOURCE_GROUP}" \
    --docker-container-logging filesystem

# --- Done ---
echo ""
echo -e "${GREEN}=== Deployment Complete ===${NC}"
echo "App URL: https://${APP_NAME}.azurewebsites.net"
echo ""
echo "First startup may take 2-3 minutes (Docker image is ~500MB)."
echo ""
echo "Useful commands:"
echo "  View logs:    az webapp log tail --name ${APP_NAME} --resource-group ${RESOURCE_GROUP}"
echo "  Restart:      az webapp restart --name ${APP_NAME} --resource-group ${RESOURCE_GROUP}"
echo "  Update code:  ./deploy-azure.sh --update"
echo "  SSH into app: az webapp ssh --name ${APP_NAME} --resource-group ${RESOURCE_GROUP}"
