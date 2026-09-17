#!/usr/bin/env bash
set -euo pipefail

APP_NAME="${APP_NAME:-weekly-sharepoint-compliance-agent}"
TENANT_ID="${TENANT_ID:-}"
REDIRECT_URI="${REDIRECT_URI:-https://localhost}"

if [[ -z "$TENANT_ID" ]]; then
  echo "Set TENANT_ID before running this script."
  exit 1
fi

az login

APP_ID=$(az ad app create \
  --display-name "$APP_NAME" \
  --sign-in-audience AzureADMyOrg \
  --web-redirect-uris "$REDIRECT_URI" \
  --query appId -o tsv)

echo "APP_ID=$APP_ID"

echo "Add Microsoft Graph permissions in Entra for the app:"
echo "- Files.Read.All"
echo "- Sites.Read.All"
echo "- Mail.Send"
echo "- ChannelMessage.Send"
echo "- TeamworkAppInstallation.ReadWriteForTeam.All"
echo "- SecurityEvents.Read.All"
echo "- User.Read"

echo "Then grant admin consent in the Entra portal."

echo "App registration complete. Keep the APP_ID values for your .env file."
