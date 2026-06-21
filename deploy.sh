#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
# deploy.sh  —  One-command deploy for GoogleSlackAlert
#
# Run:  ./deploy.sh
#
# What it does (in order):
#   1.  Checks dependencies (clasp, jq, curl)
#   2.  Runs `clasp push` to upload all source files
#   3.  Finds your existing Production deployment and redeploys it
#       (creates one if none exists yet)
#   4.  Saves the Web App URL to DEPLOYED_WEBAPP_URL in Script Properties
#       by writing+pushing+running a tiny temp function, then cleans it up
#   5.  Updates the Slack app's Interactivity, Event Subscriptions, and
#       OAuth redirect URLs via the Slack API
#   6.  Prints a clean summary
#
# One-time setup:
#   cp .env.example .env   # then fill in the values
#   chmod +x deploy.sh
#   clasp login
#
# ─────────────────────────────────────────────────────────────────────────────

set -euo pipefail

# ── Load .env ────────────────────────────────────────────────────────────────
if [[ -f ".env" ]]; then
  set -o allexport
  # shellcheck disable=SC1091
  source .env
  set +o allexport
fi

# ── Config (override via .env or environment) ─────────────────────────────────
SLACK_APP_ID="${SLACK_APP_ID:-A0BAGJL0933}"
SLACK_USER_TOKEN="${SLACK_USER_TOKEN:-}"
CLASP_DEPLOYMENT_NAME="${CLASP_DEPLOYMENT_NAME:-Production}"
UPDATE_SLACK_URLS="${UPDATE_SLACK_URLS:-true}"
SUPABASE_PROJECT_ID="${SUPABASE_PROJECT_ID:-}"

# ── Colours ───────────────────────────────────────────────────────────────────
RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'
CYAN='\033[0;36m'; BOLD='\033[1m'; NC='\033[0m'
info()    { echo -e "${CYAN}[INFO]${NC}  $*"; }
success() { echo -e "${GREEN}[OK]${NC}    $*"; }
warn()    { echo -e "${YELLOW}[WARN]${NC}  $*"; }
error()   { echo -e "${RED}[ERROR]${NC} $*" >&2; }
header()  { echo -e "\n${BOLD}${CYAN}━━━  $*  ━━━${NC}\n"; }

# ─────────────────────────────────────────────────────────────────────────────
header "Step 0 — Pre-flight checks"
# ─────────────────────────────────────────────────────────────────────────────

for cmd in clasp jq curl; do
  if command -v "$cmd" &>/dev/null; then
    success "$cmd  →  $(command -v "$cmd")"
  else
    error "$cmd not found."
    case "$cmd" in
      clasp) echo "  npm install -g @google/clasp && clasp login" ;;
      jq)    echo "  brew install jq   (or: apt-get install jq)" ;;
      curl)  echo "  Should be pre-installed on macOS/Linux." ;;
    esac
    exit 1
  fi
done

if [[ ! -f ".clasp.json" ]]; then
  error ".clasp.json not found — run 'clasp create --type standalone' or 'clasp clone <scriptId>' first."
  exit 1
fi

SCRIPT_ID=$(jq -r '.scriptId // empty' .clasp.json)
if [[ -z "$SCRIPT_ID" ]]; then
  error "Could not read scriptId from .clasp.json"
  exit 1
fi
success "scriptId  →  $SCRIPT_ID"

# ─────────────────────────────────────────────────────────────────────────────
header "Step 1 — Push source files"
# ─────────────────────────────────────────────────────────────────────────────

info "Running: clasp push --force"
clasp push --force
success "Source files pushed."

# ─────────────────────────────────────────────────────────────────────────────
header "Step 2 — Create / update deployment"
# ─────────────────────────────────────────────────────────────────────────────

# List current deployments and find the one labelled with our name.
# `clasp deployments` output looks like:
#   - AKfycby... @1.   "Production"
#   - AKfycby... @HEAD "HEAD"
info "Fetching existing deployments..."
DEPLOYMENTS=$(clasp deployments 2>&1)
echo "$DEPLOYMENTS"

# Extract the deployment ID for our named deployment (not HEAD)
EXISTING_ID=$(echo "$DEPLOYMENTS" \
  | grep -v 'HEAD' \
  | grep -o '\- [A-Za-z0-9_-]\{20,\}' \
  | awk '{print $2}' \
  | head -1 || true)

if [[ -n "$EXISTING_ID" ]]; then
  info "Found existing deployment: $EXISTING_ID"
  info "Redeploying with new version..."
  DEPLOY_OUT=$(clasp deploy \
    --deploymentId "$EXISTING_ID" \
    --description  "$CLASP_DEPLOYMENT_NAME" 2>&1)
  echo "$DEPLOY_OUT"
  DEPLOYMENT_ID="$EXISTING_ID"
else
  info "No existing non-HEAD deployment found. Creating new deployment..."
  DEPLOY_OUT=$(clasp deploy --description "$CLASP_DEPLOYMENT_NAME" 2>&1)
  echo "$DEPLOY_OUT"
  DEPLOYMENT_ID=$(echo "$DEPLOY_OUT" \
    | grep -o '\- [A-Za-z0-9_-]\{20,\}' \
    | awk '{print $2}' \
    | head -1 || true)
fi

if [[ -z "$DEPLOYMENT_ID" ]]; then
  error "Could not determine deployment ID. Check output above."
  exit 1
fi

success "Deployment ID  →  $DEPLOYMENT_ID"

# ─────────────────────────────────────────────────────────────────────────────
header "Step 3 — Resolve Web App URL"
# ─────────────────────────────────────────────────────────────────────────────

WEBAPP_URL="https://script.google.com/macros/s/${DEPLOYMENT_ID}/exec"
OAUTH_REDIRECT="${WEBAPP_URL}?action=oauth_callback"

success "Web App URL    →  $WEBAPP_URL"
success "OAuth Redirect →  $OAUTH_REDIRECT"

# ─────────────────────────────────────────────────────────────────────────────
header "Step 4 — Save DEPLOYED_WEBAPP_URL to Script Properties"
# ─────────────────────────────────────────────────────────────────────────────
# We write a tiny temporary function, push it, run it via 'clasp run',
# then delete the file and re-push to clean up.
# If 'clasp run' isn't available (no GCP project / Apps Script API not enabled),
# we print the manual step instead.

HELPER_FILE="src/__deploy_helper__.js"

cat > "$HELPER_FILE" <<JSEOF
// AUTO-GENERATED by deploy.sh — do not commit this file
function __setDeployedUrl__() {
  PropertiesService.getScriptProperties()
    .setProperty('DEPLOYED_WEBAPP_URL', '${WEBAPP_URL}');
  Logger.log('DEPLOYED_WEBAPP_URL set to: ${WEBAPP_URL}');
}
JSEOF

info "Pushing helper function..."
clasp push --force 2>/dev/null

info "Running __setDeployedUrl__ in Apps Script..."
if clasp run __setDeployedUrl__ 2>&1; then
  success "DEPLOYED_WEBAPP_URL saved to Script Properties."
else
  warn "clasp run failed. This requires:"
  warn "  • Apps Script API enabled in your GCP project"
  warn "  • A GCP project linked to this Apps Script"
  warn "  • clasp login with the correct account"
  echo ""
  echo -e "${YELLOW}Manual fallback:${NC}"
  echo "  Open the Apps Script editor → run setupDeveloperCredentials()"
  echo "  and set DEPLOYED_WEBAPP_URL = '${WEBAPP_URL}'"
fi

info "Cleaning up helper file..."
rm -f "$HELPER_FILE"
clasp push --force 2>/dev/null
success "Helper cleaned up."

# ─────────────────────────────────────────────────────────────────────────────
header "Step 4.5 — Update & Deploy Supabase Edge Function"
# ─────────────────────────────────────────────────────────────────────────────

if [[ -z "$SUPABASE_PROJECT_ID" ]]; then
  warn "SUPABASE_PROJECT_ID not set. Skipping Supabase deployment."
  SUPABASE_FUNCTION_URL="$WEBAPP_URL" # Fallback to GAS directly
else
  info "Setting GAS_WEBAPP_URL secret in Supabase..."
  if supabase secrets set GAS_WEBAPP_URL="$WEBAPP_URL" --project-ref "$SUPABASE_PROJECT_ID"; then
    success "Secret GAS_WEBAPP_URL updated in Supabase."
  else
    warn "Failed to set Supabase secret. Make sure you are logged in to Supabase CLI."
  fi

  info "Deploying Supabase edge function 'bot'..."
  if supabase functions deploy bot --project-ref "$SUPABASE_PROJECT_ID"; then
    success "Supabase edge function deployed."
  else
    warn "Failed to deploy Supabase edge function."
  fi

  SUPABASE_FUNCTION_URL="https://${SUPABASE_PROJECT_ID}.supabase.co/functions/v1/bot"
  success "Supabase Function URL →  $SUPABASE_FUNCTION_URL"
fi

# ─────────────────────────────────────────────────────────────────────────────
header "Step 5 — Update Slack App URLs"
# ─────────────────────────────────────────────────────────────────────────────

print_manual_slack_steps() {
  echo ""
  echo -e "${YELLOW}Paste these URLs into your Slack app manually:${NC}"
  echo ""
  echo "  1. Interactivity & Shortcuts"
  echo "     https://api.slack.com/apps/${SLACK_APP_ID}/interactive-messages"
  echo "     Request URL: ${SUPABASE_FUNCTION_URL}"
  echo ""
  echo "  2. Event Subscriptions"
  echo "     https://api.slack.com/apps/${SLACK_APP_ID}/event-subscriptions"
  echo "     Request URL: ${SUPABASE_FUNCTION_URL}"
  echo ""
  echo "  3. OAuth & Permissions → Redirect URLs"
  echo "     https://api.slack.com/apps/${SLACK_APP_ID}/oauth"
  echo "     Add:  ${OAUTH_REDIRECT}"
  echo ""
}

if [[ "$UPDATE_SLACK_URLS" != "true" ]]; then
  warn "UPDATE_SLACK_URLS != true — skipping Slack update."
  print_manual_slack_steps
elif [[ -z "$SLACK_USER_TOKEN" ]]; then
  warn "SLACK_USER_TOKEN not set in .env — cannot update Slack automatically."
  print_manual_slack_steps
else
  info "Updating Slack app manifest..."

  # The apps.manifest.update endpoint requires an app-configuration token
  # (xoxp- with apps:write scope, obtained via config token flow).
  # We attempt it; if it fails we fall back to manual steps.
  MANIFEST_PAYLOAD=$(cat <<JSON
{
  "app_id": "${SLACK_APP_ID}",
  "manifest": {
    "settings": {
      "interactivity": {
        "is_enabled": true,
        "request_url": "${SUPABASE_FUNCTION_URL}"
      },
      "event_subscriptions": {
        "request_url": "${SUPABASE_FUNCTION_URL}",
        "bot_events": ["app_home_opened", "app_mention"]
      }
    },
    "oauth_config": {
      "redirect_urls": ["${OAUTH_REDIRECT}"]
    }
  }
}
JSON
)

  RESPONSE=$(curl -s -X POST \
    "https://slack.com/api/apps.manifest.update" \
    -H "Authorization: Bearer ${SLACK_USER_TOKEN}" \
    -H "Content-Type: application/json" \
    -d "$MANIFEST_PAYLOAD")

  SLACK_OK=$(echo "$RESPONSE" | jq -r '.ok' 2>/dev/null || echo "false")

  if [[ "$SLACK_OK" == "true" ]]; then
    success "Slack app URLs updated via API."
  else
    SLACK_ERR=$(echo "$RESPONSE" | jq -r '.error // "unknown"' 2>/dev/null || echo "parse error")
    warn "Slack manifest API returned error: $SLACK_ERR"
    warn "(This usually means your token needs the 'apps:write' scope,"
    warn " which requires a config-level token from the Slack app config flow.)"
    print_manual_slack_steps
  fi
fi

# ─────────────────────────────────────────────────────────────────────────────
header "Deploy complete"
# ─────────────────────────────────────────────────────────────────────────────

echo -e "${BOLD}Script ID:${NC}      $SCRIPT_ID"
echo -e "${BOLD}Deployment ID:${NC}  $DEPLOYMENT_ID"
echo -e "${BOLD}Web App URL:${NC}    $WEBAPP_URL"
echo -e "${BOLD}Supabase URL:${NC}   ${SUPABASE_FUNCTION_URL:-}"
echo -e "${BOLD}OAuth Redirect:${NC} $OAUTH_REDIRECT"
echo ""
echo -e "${GREEN}Done. Next deploy:  ${BOLD}./deploy.sh${NC}"