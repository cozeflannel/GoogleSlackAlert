# GoogleSlackAlert

A dynamic Google Workspace Add-on built with Google Apps Script that monitors Google Sheets and sends alerts via Slack and Gmail when specific conditions are met.

## Features

*   **Dynamic Configuration:** A user-friendly HTML sidebar built into Google Sheets allows users to map their sheet columns (Name, Email, Status, Extra Info).
*   **Slack Integration:** Securely connect your Slack workspace via OAuth. Receive rich Block Kit messages in a designated channel.
*   **Actionable Alerts:** Slack messages and Emails include a "Take Action" button. Users can resolve alerts directly from Slack or their email, which automatically updates the originating Google Sheet and logs the resolution.
*   **Downstream Notifications (Cascade):** Optionally configure the bot to notify another sheet/user when an alert is resolved, keeping dependent workflows in sync.
*   **Multi-tenant Architecture:** Stores configurations per-spreadsheet using `DocumentProperties`, allowing a single deployment to serve multiple independent spreadsheets.

## Architecture

This project is built as a Standalone Apps Script project, intended to be deployed as an **Editor Add-on** for Google Workspace.

*   `src/Menu.js`: Injects the custom menu into the Google Sheet.
*   `src/Sidebar.html` & `src/ConfigServer.js`: Handles the UI and backend logic for user configuration and Slack OAuth flow.
*   `src/Main.js`: Contains the core logic to evaluate sheet data against user-defined conditions and trigger alerts.
*   `src/Slack.js`: Manages Slack API interactions (posting messages, opening modals, updating the App Home).
*   `src/Gmail.js`: Manages sending HTML emails with action links and resolving alerts.
*   `src/Webhooks.js`: The `doPost` and `doGet` entry points for receiving callbacks from Slack (OAuth, interactivity) and Gmail action links.

## Setup Instructions

### 1. Prerequisites

*   [Node.js and npm](https://nodejs.org/) installed.
*   [clasp](https://github.com/google/clasp) (Command Line Apps Script Projects) installed globally: `npm install -g @google/clasp`
*   A Google Cloud Project (GCP) with the **Google Workspace Marketplace SDK** and **Google Apps Script API** enabled.

### 2. Local Setup

1.  Clone this repository.
2.  Login to clasp: `clasp login`
3.  Create a new Apps Script project: `clasp create --type standalone --title "SheetAlerts"`
    *   This will create a `.clasp.json` file linking your local directory to the new script.
4.  Push the code to your Apps Script project: `clasp push`

### 3. Slack App Setup

1.  Go to [api.slack.com/apps](https://api.slack.com/apps) and create a new Slack App.
2.  In **OAuth & Permissions**, add the following Redirect URL:
    *   `https://script.google.com/macros/s/[YOUR_DEPLOYMENT_ID]/exec?action=oauth_callback`
    *   *(Note: You'll need to deploy your Apps Script project as a Web App to get this URL).*
3.  Add the following Bot Token Scopes: `chat:write`, `channels:read`, `channels:join`, `app_mentions:read`.
4.  Enable **Interactivity & Shortcuts** and set the Request URL to your Apps Script Web App URL.
5.  Enable **Event Subscriptions** and subscribe to `app_home_opened` (Set the Request URL to your Web App URL).
6.  Copy your **Client ID** and **Client Secret** from the "Basic Information" page.

### 4. Apps Script Configuration

1.  Open your Apps Script project in the browser (`clasp open`).
2.  Open `src/ConfigServer.js` and locate the `setupDeveloperCredentials()` function.
3.  Replace `'YOUR_CLIENT_ID_HERE'` and `'YOUR_CLIENT_SECRET_HERE'` with your actual Slack credentials.
4.  Run the `setupDeveloperCredentials()` function once from the editor to save these to the script's secure properties.

### 5. Deployment & Testing

To test the Add-on on a specific spreadsheet:
1.  In the Apps Script editor, go to **Deploy > Test deployments**.
2.  Select **Editor Add-on**.
3.  Select a Test Document (your Google Sheet) and click **Install**.
4.  Open the Google Sheet, refresh, and look for **Extensions > SheetAlerts > Configure Alerts**.

To enable the webhooks for Slack OAuth and interactivity:
1.  Go to **Deploy > New deployment**.
2.  Select **Web app**.
3.  Execute as: **User accessing the web app**.
4.  Who has access: **Anyone**.
5.  Deploy and copy the Web App URL to use in your Slack App configuration.
