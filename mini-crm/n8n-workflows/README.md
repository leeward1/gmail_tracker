# n8n Daily Reminder Workflow for Gmail CRM

This workflow sends daily email reminders with priority-based handling: email responses take priority over meeting follow-ups.

## Priority System

| Priority | Type | When Created | Email Subject |
|----------|------|--------------|---------------|
| **1 (High)** | Email Response | Received email not yet replied to | 🔴 RESPOND: contact@email.com |
| **2 (Normal)** | Post-Meeting | Meeting occurred, need follow-up | 📋 Follow-up: contact@email.com |

**Key behavior:** If a contact has both an unanswered email AND a past meeting, only the email reminder is shown (higher priority). Once you respond (mark complete), the meeting reminder can be created.

## How It Works

### Gmail Sync (Email Reminders)
1. Syncs Gmail threads and detects received emails
2. If last message was **received** (not sent by you):
   - Creates reminder with **priority 1** (email-response)
   - Reminder due 1 day after email received
3. Email reminders override meeting reminders (higher priority)

### Calendar Sync (Meeting Reminders)
1. Syncs past and future calendar events
2. For past meetings with external participants:
   - Creates reminder with **priority 2** (post-meeting)
   - Only if no email reminder exists for that contact
3. Updates `CRM_NextMeetingDate` for upcoming meetings

### n8n Workflow (Daily at 9 AM)
1. Reads all contacts from sheet
2. Filters for due reminders not yet complete
3. Sorts by priority (email first, then meeting)
4. Sends styled email based on priority type
5. Logs reminder to `CRM_ReminderHistory`

## New Columns

| Column | Description |
|--------|-------------|
| `CRM_ReminderPriority` | 1 = email (high), 2 = meeting (normal) |
| `CRM_PendingEmailThreadId` | Thread ID of email awaiting response |

## Scenario Examples

### Scenario 1: Meeting then Email
1. You have a meeting with John → `post-meeting` reminder created (P2)
2. John sends you an email → `email-response` reminder replaces it (P1)
3. You get reminded to respond to the email first
4. Mark complete → next Calendar Sync can create new meeting reminder

### Scenario 2: Email Only
1. New contact Sarah emails you → `email-response` reminder created (P1)
2. Daily reminder: "🔴 RESPOND: sarah@company.com"
3. You reply and mark complete → reminder stops

### Scenario 3: Meeting Only
1. Meeting with Mike, he doesn't email → `post-meeting` reminder (P2)
2. Daily reminder: "📋 Follow-up: mike@company.com"
3. You send follow-up and mark complete → reminder stops

## Setup Instructions

### Step 1: Run Setup Columns
In Google Sheets: **CRM Menu** → **Setup Columns**

### Step 2: Import Workflow
1. Open n8n
2. **Workflows** → **Import from File**
3. Select `daily-reminder-workflow.json`

### Step 3: Update Placeholders
Replace in the workflow:
- `YOUR_SHEET_ID` - Google Sheet ID from URL
- `YOUR_EMAIL@gmail.com` - Your email
- `YOUR_SHEET_URL` - Full sheet URL

### Step 4: Connect Credentials
- Google Sheets OAuth2 API
- Gmail OAuth2

### Step 5: Activate
Toggle workflow **Active**.

## Reminder History Format

Each reminder is logged with priority:
```
2024-01-15 | P1 | email-response: Re: Project Update
2024-01-16 | P1 | email-response: Re: Project Update
2024-01-17 | P2 | post-meeting: Sales Call
```

## Flow Diagram

```
                    ┌─────────────────┐
                    │  Gmail Sync     │
                    │  (received      │
                    │   emails)       │
                    └────────┬────────┘
                             │
                    ┌────────▼────────┐
                    │ Email received? │
                    └────────┬────────┘
                             │ Yes
                    ┌────────▼────────┐
                    │ Set reminder    │
                    │ Priority = 1    │
                    │ (overrides P2)  │
                    └────────┬────────┘
                             │
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
                             │
                    ┌────────▼────────┐
                    │ Calendar Sync   │
                    │ (past meetings) │
                    └────────┬────────┘
                             │
                    ┌────────▼────────┐
                    │ Has P1 reminder?│
                    └────────┬────────┘
                             │ No
                    ┌────────▼────────┐
                    │ Set reminder    │
                    │ Priority = 2    │
                    └────────┬────────┘
                             │
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
                             │
                    ┌────────▼────────┐
                    │ n8n workflow    │
                    │ (daily 9 AM)    │
                    └────────┬────────┘
                             │
                    ┌────────▼────────┐
                    │ Sort by priority│
                    │ Send emails     │
                    │ Log to history  │
                    └─────────────────┘
```
