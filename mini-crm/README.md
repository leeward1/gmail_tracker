# Gmail CRM - Hybrid Architecture

A lightweight CRM system that tracks email interactions and sends smart follow-up reminders. Built with a hybrid architecture using **n8n** for email ingestion/sending and **Google Apps Script** for data enrichment.

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                            GMAIL CRM PIPELINE                                │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐                   │
│  │   STAGE 1    │    │   STAGE 2    │    │   STAGE 3    │                   │
│  │  n8n (hourly)│───▶│ Apps Script  │───▶│  n8n (hourly)│                   │
│  │              │    │ (every 15min)│    │              │                   │
│  └──────────────┘    └──────────────┘    └──────────────┘                   │
│         │                   │                   │                            │
│         ▼                   ▼                   ▼                            │
│  ┌──────────────┐    ┌──────────────┐    ┌──────────────┐                   │
│  │    Gmail     │    │   Google     │    │    Gmail     │                   │
│  │   Ingestion  │    │   Sheets     │    │   Send       │                   │
│  │              │    │  (Database)  │    │   Reminders  │                   │
│  └──────────────┘    └──────────────┘    └──────────────┘                   │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Stage 1: Gmail Ingestion (n8n)
**File:** `n8n-workflows/01-gmail-ingestion.json`
**Schedule:** Every hour

1. Fetches recent emails from Gmail (last 2 days)
2. Filters out spam, promotions, notifications, and internal emails
3. Transforms email data and generates Gmail search links
4. Deduplicates against existing records
5. Appends new interactions to `CRM_Interactions` sheet

### Stage 2: Enrichment (Apps Script)
**File:** `CRM_Stable.gs` (copy to Code.gs in Apps Script)
**Schedule:** Every 15 minutes (automatic trigger)

1. Reads unprocessed interactions from `CRM_Interactions`
2. Creates/updates contact records in `CRM_Contacts`
3. Creates reminder entries in `CRM_Reminders` for received emails
4. Marks interactions as processed

### Stage 3: Reminder Sender (n8n)
**File:** `n8n-workflows/02-reminder-sender.json`
**Schedule:** Every hour

1. Reads due reminders from `CRM_Reminders` (status=queued, nextAttemptAt <= now)
2. Acquires lock (sets status=sending, lockExpiresAt)
3. Sends reminder email with Gmail search link
4. Marks as sent or handles failures with exponential backoff

---

## Google Sheets Structure

The system uses 5 tabs in your Google Sheet:

### CRM_Contacts
One row per unique contact.

| Column | Description |
|--------|-------------|
| contactId | Unique ID (C + timestamp) |
| email | Contact's email address |
| name | Contact name (manual entry) |
| company | Company name (manual entry) |
| lastSeenAt | Last interaction date |
| lastDirection | "Sent" or "Received" |
| lastSubject | Subject of last email |
| lastMessageDate | Date of last message |
| notes | Free-form notes |

### CRM_Interactions
Append-only log of all email interactions.

| Column | Description |
|--------|-------------|
| interactionId | Unique ID (I + timestamp) |
| messageId | Gmail message ID |
| contactEmail | Contact's email |
| date | Email date |
| subject | Email subject |
| direction | "Sent" or "Received" |
| fromAddress | Full from address |
| snippet | First 200 chars of email body |
| gmailSearchLink | Direct link to email in Gmail |
| fallbackSearch | Gmail search query if link fails |
| processedAt | When Apps Script processed this row |

### CRM_Reminders
Queue for follow-up reminders with state machine.

| Column | Description |
|--------|-------------|
| reminderId | Unique ID (R + timestamp) |
| contactEmail | Contact to remind about |
| contactName | Contact name |
| type | "email-response" or "meeting-followup" |
| status | queued/sending/sent/failed/abandoned |
| nextAttemptAt | When to attempt sending |
| lockExpiresAt | Lock timeout for stale recovery |
| attemptCount | Number of send attempts |
| lastError | Error message if failed |
| gmailSearchLink | Link to original email |
| fallbackSearch | Backup search query |
| subject | Original email subject |
| snippet | Email preview |
| createdAt | When reminder was created |
| sentAt | When reminder was sent |
| idempotencyKey | Prevents double-sends |
| sentKeys | History of sent idempotency keys |

### CRM_Runs
Observability log for debugging.

| Column | Description |
|--------|-------------|
| runId | Unique ID (RUN + timestamp) |
| stage | Which stage ran |
| startedAt | Run start time |
| completedAt | Run end time |
| status | success/error |
| itemsProcessed | Total items processed |
| itemsSucceeded | Successful items |
| itemsFailed | Failed items |

### CRM_Settings
Configuration key-value pairs.

| Key | Description |
|-----|-------------|
| EXCLUDE_DOMAINS | Domains to ignore (sendgrid.net, etc.) |
| EXCLUDE_PREFIXES | Email prefixes to ignore (noreply, etc.) |
| MY_DOMAIN | Your company domain (knostic.ai) |
| REMINDER_DAYS | Days before creating reminder |
| ARCHIVE_DAYS | Days before archiving old interactions |

---

## Installation

### 1. Google Apps Script Setup

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Delete existing code and paste contents of `CRM_Stable.gs`
4. Save (Cmd+S)
5. Click **Project Settings** (gear icon) → Check "Show appsscript.json"
6. Update `appsscript.json` manifest:
   ```json
   {
     "timeZone": "America/New_York",
     "dependencies": {},
     "exceptionLogging": "STACKDRIVER",
     "runtimeVersion": "V8",
     "oauthScopes": [
       "https://www.googleapis.com/auth/spreadsheets",
       "https://www.googleapis.com/auth/script.scriptapp",
       "https://www.googleapis.com/auth/gmail.readonly"
     ]
   }
   ```
7. Refresh your Google Sheet
8. From menu: **Stable CRM → Setup Tabs**
9. From menu: **Stable CRM → Setup Triggers**
10. Grant permissions when prompted

### 2. n8n Setup

1. Import `n8n-workflows/01-gmail-ingestion.json`
   - Select your Gmail OAuth credentials
   - Select your Google Sheets credentials
   - Activate the workflow

2. Import `n8n-workflows/02-reminder-sender.json`
   - Select your Gmail OAuth credentials
   - Select your Google Sheets credentials
   - Activate the workflow

---

## Configuration

### Your Email & Sheet ID
Already configured in the workflow files:
- Email: `gadi@knostic.ai`
- Sheet ID: `19TkAK9YN0OlpDe2t6QDytAzc1Bctqm9jh2XyqAvWx8g`

### Exclude Domains
Edit `CRM_Settings` sheet to add domains you want to ignore:
```
sendgrid.net,mailchimp.com,mailgun.org,amazonses.com,postmarkapp.com
```

### Exclude Prefixes
Email address prefixes to ignore:
```
noreply,no-reply,notifications,mailer-daemon,bounce
```

---

## How It Works

### Email Flow
```
Incoming Email
      │
      ▼
┌─────────────────┐
│ n8n Stage 1     │ Hourly: Fetch from Gmail
│ - Filter spam   │
│ - Skip internal │
│ - Deduplicate   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ CRM_Interactions│ Append-only log
└────────┬────────┘
         │
         ▼
┌─────────────────┐
│ Apps Script     │ Every 15 min
│ Stage 2         │
│ - Update contact│
│ - Create remind │
└────────┬────────┘
         │
    ┌────┴────┐
    ▼         ▼
┌───────┐ ┌─────────┐
│Contact│ │Reminder │
│Updated│ │Queued   │
└───────┘ └────┬────┘
               │
               ▼
┌─────────────────┐
│ n8n Stage 3     │ Hourly
│ - Send reminder │
│ - Handle errors │
└─────────────────┘
```

### Reminder State Machine
```
                    ┌─────────┐
                    │ queued  │ Initial state
                    └────┬────┘
                         │ nextAttemptAt <= now
                         ▼
                    ┌─────────┐
           ┌────────│ sending │────────┐
           │        └─────────┘        │
           │ success              failure
           ▼                           ▼
      ┌─────────┐              ┌─────────┐
      │  sent   │              │ failed  │──┐
      └─────────┘              └─────────┘  │
                                     │      │ attemptCount < 4
                                     │      │ backoff & retry
                                     │◀─────┘
                                     │
                                     │ attemptCount >= 4
                                     ▼
                              ┌───────────┐
                              │ abandoned │
                              └───────────┘
```

### Backoff Schedule
| Attempt | Wait Time |
|---------|-----------|
| 1st retry | 5 minutes |
| 2nd retry | 30 minutes |
| 3rd retry | 4 hours |
| 4th+ | Abandoned |

---

## Troubleshooting

### Check Run History
- In Google Sheet: **Stable CRM → View Run History**
- Shows all Apps Script runs with success/error status

### Check n8n Executions
- In n8n: Click on workflow → Executions tab
- Shows all runs with input/output data

### Common Issues

**"Required sheets not found"**
- Run **Stable CRM → Setup Tabs** first

**Permission errors**
- Re-run **Setup Triggers** and grant permissions
- Check `appsscript.json` has correct oauth scopes

**No emails being ingested**
- Check n8n workflow is activated
- Verify Gmail credentials are connected
- Check Gmail search filters aren't too restrictive

**Reminders not sending**
- Check `CRM_Reminders` has rows with status=queued
- Verify n8n reminder workflow is activated
- Check for errors in n8n execution history

---

## File Structure

```
mini-crm/
├── README.md                     # This file
├── CRM_Stable.gs                 # Apps Script code (copy to Code.gs)
├── Code.gs                       # Original CRM code (legacy)
└── n8n-workflows/
    ├── 01-gmail-ingestion.json   # Stage 1: Email ingestion
    └── 02-reminder-sender.json   # Stage 3: Send reminders
```

---

## Future Improvements

- [ ] Add meeting follow-up reminders (calendar integration)
- [ ] Enrich contacts with company data (Clearbit, etc.)
- [ ] Add email templates for common responses
- [ ] Dashboard for CRM metrics
- [ ] Slack/Teams notifications for urgent reminders
- [ ] AI-powered email categorization

---

## License

MIT
