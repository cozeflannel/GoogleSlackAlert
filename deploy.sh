#!/bin/bash
set -e

# ─────────────────────────────────────────────────────────────────────────────
# 1. Environment Setup
# ─────────────────────────────────────────────────────────────────────────────
echo "🚀 Starting Full System Deployment..."

# Load environment variables from .env
if [ -f .env ]; then
  export $(grep -v '^#' .env | xargs)
else
  echo "❌ Error: .env file not found. Please create one based on .env.example"
  exit 1
fi

# Validate critical variables
REQUIRED_VARS=("SUPABASE_PROJECT_ID" "SUPABASE_URL" "SUPABASE_SERVICE_ROLE_KEY" "GAS_SCRIPT_ID" "GAS_DEPLOYMENT_ID")
for var in "${REQUIRED_VARS[@]}"; do
  if [ -z "${!var}" ]; then
    echo "❌ Error: $var is not set in .env"
    exit 1
  fi
done

# Check for required CLIs
if ! command -v supabase &> /dev/null; then
  echo "❌ Error: Supabase CLI not found. Please install it first."
  exit 1
fi

if ! command -v clasp &> /dev/null; then
  echo "❌ Error: clasp CLI not found. Please install it (npm install -g @google/clasp)."
  exit 1
fi

# ─────────────────────────────────────────────────────────────────────────────
# 2. Supabase Deployment
# ─────────────────────────────────────────────────────────────────────────────
echo "📦 Step 1/4: Syncing Secrets to Supabase..."
# We push developer credentials from .env to the Supabase cloud environment
supabase secrets set SLACK_CLIENT_ID="$SLACK_CLIENT_ID" 
                   SLACK_CLIENT_SECRET="$SLACK_CLIENT_SECRET" 
                   SLACK_SIGNING_SECRET="$SLACK_SIGNING_SECRET" 
                   SUPABASE_URL="$SUPABASE_URL" 
                   SUP_SECRET_KEY="$SUPABASE_SERVICE_ROLE_KEY" 
                   --project-ref "$SUPABASE_PROJECT_ID"

echo "⚡ Step 2/4: Deploying Edge Functions..."
# This uploads the Deno code in /supabase/functions/bot to the cloud
supabase functions deploy bot --project-ref "$SUPABASE_PROJECT_ID"

# ─────────────────────────────────────────────────────────────────────────────
# 3. Apps Script Deployment
# ─────────────────────────────────────────────────────────────────────────────
echo "📝 Step 3/4: Pushing Code to Google Apps Script..."
# 'push' uploads the local JS files to the GAS editor
clasp push

echo "🚀 Step 4/4: Deploying GAS Production Version..."
# 'deploy' takes that code and creates a live, versioned deployment
clasp deploy -i "$GAS_DEPLOYMENT_ID" -d "Production"

echo "─────────────────────────────────────────────────────────────────────────────"
echo "✅ DEPLOYMENT COMPLETE!"
echo "Your changes are now live on both Supabase and Google Apps Script."
echo "─────────────────────────────────────────────────────────────────────────────"
