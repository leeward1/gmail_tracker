# Gmail Email Tracker for Google Sheets - Refactored

A comprehensive Gmail tracking solution that monitors sent and received emails, tracks responses, manages campaigns, and provides analytics.

## Preserving Your Existing Data

**Your manual data in columns A, B, C, etc. will NOT be overwritten.** The tracker adds its columns AFTER your existing data.

### Setting the Start Column

Three ways to control where tracker columns begin:

| Method | How | Best For |
|--------|-----|----------|
| **Auto-detect** | Run `setupSheet()` | First-time setup - detects last column, starts after it |
| **Menu option** | Gmail Tracker > Settings > Set Tracker Start Column | Interactive - prompts for column number with current state shown |
| **Hardcode** | Edit `MANUAL_START_COLUMN = 4` at top of script | If you always want column D (change 4 to your column number) |

### Example: Your data is in columns A-C

1. **Option A (Auto):** Just run `setupSheet()` - it will detect column C has data and start tracker at column D
2. **Option B (Menu):** Go to Gmail Tracker > Settings > Set Tracker Start Column, enter `4`
3. **Option C (Code):** Change line 44 to `const MANUAL_START_COLUMN = 4;`

The tracker will NEVER write to columns A, B, or C - your existing data is safe.

## How to Configure

### Required Script Properties

The script uses `PropertiesService` for runtime configuration. All keys are auto-managed, but you can pre-configure:

| Property Key | Description | Example Value |
|-------------|-------------|---------------|
| `EMAIL_TRACKER_COLUMN_OFFSET` | Column where tracker data starts | `4` (Column D) |
| `EMAIL_TRACKER_START_DATE` | Only track emails after this date | `2024-01-01` |
| `EMAIL_TRACKER_LAST_SYNC_AT` | Timestamp of last successful sync | Auto-managed |
| `EMAIL_TRACKER_SUBJECT_FILTERS` | JSON array of subject filters | `["Sales","Meeting"]` |

### Required Sheet Names

| Sheet Name | Purpose |
|-----------|---------|
| `Sheet1` | Main tracker sheet (configurable via `CONFIG.SHEET_NAME`) |
| `Subject Analytics` | Analytics dashboard (auto-created) |
| `Campaigns` | Campaign management (auto-created) |
| `Contact Notes` | Contact notes storage (auto-created) |
| `Health` | Sync run diagnostics (auto-created) |

### Required Headers (36 columns)

The tracker uses these headers, auto-created during setup:

1. Message ID, Thread ID, Direction, From, To, CC, BCC
2. Date, Subject, Body Preview, Full Body, Status, Thread Count
3. Attachments, Attachment Names
4. Reply 1-3 (Date, From, Subject, Body for each)
5. Last Meeting, Next Meeting, Meeting Title, Meeting History, Meeting Participants, Meeting Host
6. Contact Notes, Related Contacts, Thread Contacts

### Gmail Query Format

The script builds queries automatically based on configuration:

```
(in:sent OR in:inbox) after:2024-01-01 -label:Tracked -category:promotions -category:social
```

**Query Components:**
- `in:sent` - Included if `CONFIG.TRACK_SENT = true`
- `in:inbox` - Included if `CONFIG.TRACK_RECEIVED = true`
- `after:YYYY/MM/DD` - Based on last sync or start date
- `-label:Tracked` - Excludes already-processed emails
- `-category:*` - Excludes configured categories

## How to Run

### Main Entrypoints

| Function | Description | Recommended Trigger |
|----------|-------------|---------------------|
| `runTracker()` | Full sync + reply check | Manual or hourly |
| `trackEmails()` | Sync new emails only | Time-driven, every 5 min |
| `checkForReplies()` | Check for replies only | Time-driven, every 10 min |
| `refreshAll()` | trackEmails + checkForReplies | Manual |
| `updateAnalytics()` | Refresh analytics sheet | Manual or daily |
| `checkAndSendReminders()` | Send reminder emails | Time-driven, hourly |

### Setting Up Triggers

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Click the clock icon (Triggers) in the left sidebar
4. Add triggers:
   - `trackEmails` - Time-driven, every 5 minutes
   - `checkForReplies` - Time-driven, every 10 minutes
   - `checkAndSendReminders` - Time-driven, every 1 hour

### First-Time Setup

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Paste the script code
4. Save and run `setupSheet()` once
5. Authorize when prompted
6. Configure triggers as described above

## What Changed (Refactoring Improvements)

### Architecture

- **TSS-style separation**: Code organized into CONFIG, UTILS, REPOSITORIES, SERVICES, CONTROLLERS
- **Single source of truth**: All headers defined in `CONFIG.HEADERS` array
- **Dependency injection pattern**: Repositories abstract data access from business logic

### Idempotent Column Management

- `SheetRepository.ensureHeaders()` - Only adds missing headers, never duplicates
- `SheetRepository.getHeaderIndexMap()` - Dynamic header -> column index mapping
- Running setup multiple times produces identical sheet structure

### Security Framework

- **Input sanitization**: `Utils.sanitizeText()` removes control characters
- **Formula injection prevention**: Leading `=`, `+`, `-`, `@`, `\t`, `\r` characters are escaped with `'`
- **Least privilege**: Only required Gmail/Calendar scopes used
- **No secrets in code**: All config via PropertiesService

### Comprehensive Logging

- **`RunSummary` object**: Tracks all metrics during sync
- **`printSummary()` function**: Outputs detailed run report to logs:
  - Run timestamp and duration
  - Spreadsheet ID/URL and sheet names
  - Rows processed/updated/created
  - Errors encountered
  - Gmail actions taken

### Performance Optimizations

- Batch reads/writes with `getValues()`/`setValues()`
- Sheet references cached, not re-fetched in loops
- Lock service prevents concurrent execution
- Incremental sync with overlap window

### Error Handling

- Try-catch around all external API calls
- Errors logged to Health sheet
- Optional email notifications on errors
- Individual message errors don't abort entire sync

### Code Quality

- ES6-compatible for Apps Script
- Consistent naming conventions
- Clear section organization
- JSDoc-style comments

## Scopes Required

The script requires these OAuth scopes (automatically prompted):

```
https://www.googleapis.com/auth/gmail.readonly
https://www.googleapis.com/auth/gmail.labels
https://www.googleapis.com/auth/gmail.send
https://www.googleapis.com/auth/spreadsheets
https://www.googleapis.com/auth/calendar.readonly
https://www.googleapis.com/auth/script.scriptapp
```

## Configuration Options

Edit the `CONFIG` object at the top of the script:

```javascript
const CONFIG = {
  // Sheet names
  SHEET_NAME: 'Sheet1',

  // Email matching column (0 to disable)
  EMAIL_MATCH_COLUMN: 3,  // Column C

  // Internal domains to exclude
  INTERNAL_DOMAINS: ['yourcompany.com'],
  EXCLUDE_INTERNAL_ONLY: true,

  // What to track
  TRACK_SENT: true,
  TRACK_RECEIVED: true,

  // Gmail categories to exclude
  EXCLUDE_PROMOTIONS: true,
  EXCLUDE_SOCIAL: true,
  EXCLUDE_UPDATES: true,
  EXCLUDE_FORUMS: true,

  // Performance limits
  MAX_THREADS_PER_RUN: 50,
  MAX_ROWS_TO_CHECK: 500,

  // Automation
  REMINDER_DAYS_DEFAULT: 7,
  REMINDER_EMAIL_ENABLED: true,
  EMAIL_ON_ERROR: false
};
```

## Troubleshooting

### "Sheet not found" error
Run `setupSheet()` first to create the tracker sheet and headers.

### Duplicate columns appearing
This refactored version prevents duplicates. If you have existing duplicates, delete them manually and run `fixHeaderLabels()`.

### Emails not being tracked
1. Check the Health sheet for error logs
2. Verify your Gmail query in the Health sheet's "Query Used" column
3. Ensure the `Tracked` label exists in Gmail
4. Run `resetSyncCursor()` to force a full re-scan

### Performance issues
1. Reduce `MAX_THREADS_PER_RUN` if hitting timeout
2. Reduce `MAX_ROWS_TO_CHECK` for reply checking
3. Set `STORE_FULL_BODY = false` to reduce data storage
