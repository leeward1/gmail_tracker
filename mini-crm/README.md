# Mini Salesforce CRM for Google Sheets

A lightweight, modular CRM system built on Google Sheets + Apps Script with Gmail and Gong integrations.

## Architecture

```
mini-crm/
├── 00_Config.gs      # Configuration, constants, sheet/header definitions
├── 01_Utils.gs       # Utility functions (IDs, sanitization, parsing)
├── 02_Repositories.gs # Data access layer (sheet operations, upserts)
├── 03_Services.gs    # Business logic (CRM, Gmail, Gong services)
├── 04_Controllers.gs # Entry points (menu handlers, triggers)
├── 05_Tests.gs       # TDD test suite
└── README.md         # This file
```

### Layered Design

| Layer | Purpose | Files |
|-------|---------|-------|
| **Config** | Constants, headers, defaults | `00_Config.gs` |
| **Utils** | Pure functions (no side effects) | `01_Utils.gs` |
| **Repositories** | Data access, sheet operations | `02_Repositories.gs` |
| **Services** | Business logic, integrations | `03_Services.gs` |
| **Controllers** | User-facing entry points | `04_Controllers.gs` |

## Quick Start

1. **Copy all `.gs` files** to your Apps Script project
2. **Run `setupCrm()`** to create sheets and load defaults
3. **Configure settings** in the `Config` sheet
4. **Set up triggers** for automatic syncing

## Sheets Created

| Sheet | Purpose |
|-------|---------|
| `Leads` | Sales leads before qualification |
| `Contacts` | Qualified contacts (ONE ROW PER CONTACT) |
| `Accounts` | Companies/organizations |
| `Opportunities` | Sales pipeline deals |
| `Activities` | Tasks, calls, meetings, notes |
| `Email_Log` | Gmail integration log |
| `Gong_Calls` | Gong call integration log |
| `Config` | Runtime configuration |
| `CRM_Health` | Sync run diagnostics |

## Key Features

### One Row Per Contact (HARD RULE)

The system enforces strict contact deduplication:

1. **Primary match**: Email (normalized, lowercase)
2. **Fallback match**: Phone number OR (First + Last + Company)
3. **Upsert pattern**: Always search before insert; update if exists

```javascript
// This will NEVER create duplicate contacts
ContactRepo.upsertContact({ Email: 'john@example.com', First_Name: 'John' });
ContactRepo.upsertContact({ Email: 'JOHN@EXAMPLE.COM', Last_Name: 'Doe' }); // Updates same contact
```

### Idempotent Header Management

Running setup multiple times will NOT create duplicate columns:

```javascript
setupCrm(); // Creates headers
setupCrm(); // Safe to rerun - no duplicates
setupCrm(); // Still safe
```

### Editable Columns (Manual Data Preservation)

Certain columns are marked as **editable** and will NOT be overwritten during sync operations. This allows you to add notes, assign owners, or customize records without losing your changes when Gmail or Gong syncs run.

**Default Editable Columns:**

| Sheet | Editable Columns |
|-------|-----------------|
| Contacts | Notes, Owner |
| Leads | Notes, Owner |
| Accounts | Notes, Industry, Owner, Account_Status |
| Opportunities | Notes, Owner, Next_Step, Stage, Amount, Probability, Close_Date |
| Activities | Description, Status |
| Gong_Calls | Notes_Summary, Next_Steps |

**How it works:**
```javascript
// First sync creates a contact
ContactRepo.upsertContact({ Email: 'john@example.com', First_Name: 'John' });

// User manually adds notes in the sheet
// Notes column now contains: "Met at conference, interested in product"

// Later sync updates the contact - NOTES ARE PRESERVED
ContactRepo.upsertContact({ Email: 'john@example.com', Company: 'Acme Corp' });
// Notes still contains: "Met at conference, interested in product"
```

**Adding a new editable column:**
1. Add the column name to the appropriate header array in `CRM_HEADERS` (in `00_Config.gs`)
2. Add the column name to the `EDITABLE_COLUMNS` array for that sheet
3. Run `setupCrm()` to add the new column

```javascript
// Example: Making "Custom_Field" editable for Contacts
// In 00_Config.gs:

// 1. Add to headers
CRM_HEADERS.CONTACTS = [..., 'Custom_Field'];

// 2. Add to editable list
EDITABLE_COLUMNS.CONTACTS = ['Notes', 'Owner', 'Custom_Field'];
```

**Override editable preservation:**
```javascript
// Force overwrite editable columns (use with caution)
ContactRepo.upsertContact(data, { preserveEditable: false });
```

### Gmail Integration

- Syncs emails based on configurable query
- Deduplicates by Gmail MessageId
- Auto-creates contacts from email participants
- Labels processed threads

### Gong Integration

- Fetches calls via Gong API
- Deduplicates by Gong Call ID
- Links calls to contacts by participant email
- Links to accounts by email domain

## Configuration

### Config Sheet Settings

| Key | Default | Description |
|-----|---------|-------------|
| `GMAIL_ENABLED` | TRUE | Enable Gmail sync |
| `GMAIL_QUERY` | `newer_than:7d ...` | Gmail search query |
| `GMAIL_MAX_THREADS` | 50 | Max threads per sync |
| `GONG_ENABLED` | FALSE | Enable Gong sync |
| `GONG_SYNC_DAYS_BACK` | 30 | Days of call history |
| `AUTO_CREATE_CONTACTS` | TRUE | Auto-create contacts from emails/calls |
| `AUTO_CREATE_ACCOUNTS` | TRUE | Auto-create accounts from domains |

### Gong API Setup

1. Get API access token from Gong
2. Open **Apps Script > Project Settings > Script Properties**
3. Add property: `CRM_GONG_ACCESS_TOKEN` = your token
4. Optionally: `CRM_GONG_BASE_URL` = custom base URL

## Entry Points

### Setup

| Function | Description |
|----------|-------------|
| `setupCrm()` | Initialize all sheets and config (idempotent) |

### Sync

| Function | Description | Recommended Trigger |
|----------|-------------|---------------------|
| `ingestGmail()` | Sync Gmail emails | Every 10-15 minutes |
| `syncGongCalls()` | Sync Gong calls | Every hour |
| `syncAll()` | Run all syncs | Manual or daily |

### CRM Operations

| Function | Description |
|----------|-------------|
| `createLead(data)` | Create a new lead |
| `convertLead(leadId)` | Convert lead to contact |
| `createOpportunity(data)` | Create opportunity |
| `addActivity(data)` | Log an activity |

### Utilities

| Function | Description |
|----------|-------------|
| `printSummary()` | Print last run summary to logs |
| `runAllTests()` | Run test suite |

## Testing

Run the test suite:

```javascript
runAllTests(); // Full test suite
runQuickTest(); // Quick smoke test
```

Tests cover:
- Header idempotency (no duplicate columns)
- Contact deduplication (one row per contact)
- Gmail MessageId deduplication
- Gong Call ID deduplication
- Contact matching (email-first, fallback logic)
- Security (formula injection prevention)

## Security

### Implemented Protections

1. **Formula Injection Prevention**: All input sanitized before sheet write
2. **Secrets in PropertiesService**: No credentials in code
3. **Error Sanitization**: Sensitive data redacted from logs
4. **Rate Limiting**: Retry with backoff for API calls
5. **Lock Service**: Prevents concurrent sync execution

### OAuth Scopes Required

- `gmail.readonly` - Read emails
- `gmail.labels` - Manage labels
- `spreadsheets` - Read/write sheets
- `script.external_request` - Gong API calls

## Extending the CRM

### Adding a New Sheet

1. Add to `CRM_SHEETS` in `00_Config.gs`
2. Add headers to `CRM_HEADERS` in `00_Config.gs`
3. Create repository in `02_Repositories.gs` if needed
4. Add service methods in `03_Services.gs`

### Adding a New Integration

1. Add config keys to `DEFAULT_CONFIG`
2. Create service object in `03_Services.gs`
3. Add controller function in `04_Controllers.gs`
4. Add to menu in `onOpen()`

### Adding New Tests

Add test functions in `05_Tests.gs`:

```javascript
function testMyFeature() {
  Logger.log('Running my feature tests...');
  TestRunner.assertEqual(actual, expected, 'description');
  // ...
}

// Add to runAllTests()
```

## Troubleshooting

### "Sheet not found" error
Run `setupCrm()` first.

### Gmail not syncing
1. Check `GMAIL_ENABLED` in Config sheet
2. Check query in `GMAIL_QUERY`
3. View `CRM_Health` sheet for errors

### Gong API errors
1. Verify API token in Script Properties
2. Check `GONG_ENABLED` is TRUE
3. View logs: View > Logs

### Duplicate contacts appearing
This should not happen. If it does:
1. Run `runAllTests()` to verify deduplication
2. Check if contacts were created before system was set up
3. Review `Contact_Key` column for matching issues
