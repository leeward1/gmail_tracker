/**
 * STABLE CRM SYSTEM v2
 *
 * Hybrid architecture:
 * - n8n: Gmail ingestion, cleaning, reminder sending
 * - Apps Script: Enrichment only (time-triggered, batched)
 *
 * Sheet Tabs:
 * - CRM_Contacts: One row per contact
 * - CRM_Interactions: Append-only email log
 * - CRM_Reminders: Queue with state machine
 * - CRM_Runs: Observability log
 * - CRM_Settings: Configuration
 */

// ============================================================================
// CONFIGURATION
// ============================================================================

var CRM_CONFIG = {
  TABS: {
    CONTACTS: 'CRM_Contacts',
    INTERACTIONS: 'CRM_Interactions',
    REMINDERS: 'CRM_Reminders',
    RUNS: 'CRM_Runs',
    SETTINGS: 'CRM_Settings',
    ARCHIVE: 'CRM_Interactions_Archive'
  },

  BATCH_SIZE: 100,  // Process 100 rows per run to avoid timeout
  ARCHIVE_DAYS: 90, // Archive interactions older than 90 days

  // Backoff schedule (in minutes)
  BACKOFF: {
    1: 5,      // 1st retry: 5 min
    2: 30,     // 2nd retry: 30 min
    3: 240,    // 3rd retry: 4 hours
    MAX_ATTEMPTS: 4  // After 4 attempts, mark as abandoned
  },

  // Reminder statuses
  STATUS: {
    QUEUED: 'queued',
    SENDING: 'sending',
    SENT: 'sent',
    FAILED: 'failed',
    ABANDONED: 'abandoned'
  }
};

// ============================================================================
// SETUP - Create all required tabs
// ============================================================================

/**
 * Setup the stable CRM system - creates all required tabs
 */
function setupStableCRM() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var response = ui.alert('Setup Stable CRM',
    'This will create the following tabs:\n\n' +
    '• CRM_Contacts - One row per contact\n' +
    '• CRM_Interactions - Email log (append-only)\n' +
    '• CRM_Reminders - Reminder queue\n' +
    '• CRM_Runs - Run history\n' +
    '• CRM_Settings - Configuration\n\n' +
    'Existing tabs with these names will NOT be overwritten.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO);

  if (response !== ui.Button.YES) return;

  var created = [];
  var existed = [];

  // CRM_Contacts
  if (!ss.getSheetByName(CRM_CONFIG.TABS.CONTACTS)) {
    var contactsSheet = ss.insertSheet(CRM_CONFIG.TABS.CONTACTS);
    contactsSheet.getRange(1, 1, 1, 9).setValues([[
      'contactId', 'email', 'name', 'company', 'lastSeenAt',
      'lastDirection', 'lastSubject', 'lastMessageDate', 'notes'
    ]]);
    contactsSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    contactsSheet.setFrozenRows(1);
    created.push('CRM_Contacts');
  } else {
    existed.push('CRM_Contacts');
  }

  // CRM_Interactions
  if (!ss.getSheetByName(CRM_CONFIG.TABS.INTERACTIONS)) {
    var interactionsSheet = ss.insertSheet(CRM_CONFIG.TABS.INTERACTIONS);
    interactionsSheet.getRange(1, 1, 1, 11).setValues([[
      'interactionId', 'messageId', 'contactEmail', 'date', 'subject',
      'direction', 'fromAddress', 'snippet', 'gmailSearchLink',
      'fallbackSearch', 'processedAt'
    ]]);
    interactionsSheet.getRange(1, 1, 1, 11).setFontWeight('bold');
    interactionsSheet.setFrozenRows(1);
    created.push('CRM_Interactions');
  } else {
    existed.push('CRM_Interactions');
  }

  // CRM_Reminders
  if (!ss.getSheetByName(CRM_CONFIG.TABS.REMINDERS)) {
    var remindersSheet = ss.insertSheet(CRM_CONFIG.TABS.REMINDERS);
    remindersSheet.getRange(1, 1, 1, 17).setValues([[
      'reminderId', 'contactEmail', 'contactName', 'type', 'status',
      'nextAttemptAt', 'lockExpiresAt', 'attemptCount', 'lastError',
      'gmailSearchLink', 'fallbackSearch', 'subject', 'snippet',
      'createdAt', 'sentAt', 'idempotencyKey', 'sentKeys'
    ]]);
    remindersSheet.getRange(1, 1, 1, 17).setFontWeight('bold');
    remindersSheet.setFrozenRows(1);
    created.push('CRM_Reminders');
  } else {
    existed.push('CRM_Reminders');
  }

  // CRM_Runs
  if (!ss.getSheetByName(CRM_CONFIG.TABS.RUNS)) {
    var runsSheet = ss.insertSheet(CRM_CONFIG.TABS.RUNS);
    runsSheet.getRange(1, 1, 1, 8).setValues([[
      'runId', 'stage', 'startedAt', 'completedAt', 'status',
      'itemsProcessed', 'itemsSucceeded', 'itemsFailed'
    ]]);
    runsSheet.getRange(1, 1, 1, 8).setFontWeight('bold');
    runsSheet.setFrozenRows(1);
    created.push('CRM_Runs');
  } else {
    existed.push('CRM_Runs');
  }

  // CRM_Settings
  if (!ss.getSheetByName(CRM_CONFIG.TABS.SETTINGS)) {
    var settingsSheet = ss.insertSheet(CRM_CONFIG.TABS.SETTINGS);
    settingsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    settingsSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    settingsSheet.setFrozenRows(1);

    // Add default settings
    settingsSheet.getRange(2, 1, 5, 2).setValues([
      ['EXCLUDE_DOMAINS', 'sendgrid.net,mailchimp.com,mailgun.org,amazonses.com,postmarkapp.com'],
      ['EXCLUDE_PREFIXES', 'noreply,no-reply,notifications,mailer-daemon,bounce'],
      ['MY_DOMAIN', 'knostic.ai'],
      ['REMINDER_DAYS', '1'],  // Days after received email to create reminder
      ['ARCHIVE_DAYS', '90']
    ]);
    created.push('CRM_Settings');
  } else {
    existed.push('CRM_Settings');
  }

  ui.alert('Setup Complete',
    'Created: ' + (created.length > 0 ? created.join(', ') : 'None') + '\n' +
    'Already existed: ' + (existed.length > 0 ? existed.join(', ') : 'None'),
    ui.ButtonSet.OK);
}

// ============================================================================
// STAGE 2: ENRICHMENT (Apps Script - Time Triggered)
// ============================================================================

/**
 * Process unprocessed interactions and create reminders
 * This is triggered by a time-based trigger (every 15 min)
 * OPTIMIZED: Batches all updates to minimize API calls
 */
function processInteractions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var interactionsSheet = ss.getSheetByName(CRM_CONFIG.TABS.INTERACTIONS);
  var remindersSheet = ss.getSheetByName(CRM_CONFIG.TABS.REMINDERS);
  var contactsSheet = ss.getSheetByName(CRM_CONFIG.TABS.CONTACTS);

  if (!interactionsSheet || !remindersSheet) {
    Logger.log('Required sheets not found. Run setupStableCRM first.');
    return;
  }

  var runId = logRunStart('Stage2-Enrichment');
  var stats = { processed: 0, remindersCreated: 0, contactsUpdated: 0, errors: 0 };

  try {
    var settings = getStableSettings();
    var myDomain = settings['MY_DOMAIN'] || 'knostic.ai';

    // Get interactions where processedAt is empty
    var lastRow = interactionsSheet.getLastRow();
    if (lastRow < 2) {
      logRunEnd(runId, 'success', stats);
      return stats;
    }

    var headers = getSheetHeaders(interactionsSheet);
    var data = interactionsSheet.getRange(2, 1, lastRow - 1, interactionsSheet.getLastColumn()).getValues();

    var processedAtCol = headers['processedAt'];
    var directionCol = headers['direction'];
    var contactEmailCol = headers['contactEmail'];
    var subjectCol = headers['subject'];
    var gmailSearchLinkCol = headers['gmailSearchLink'];
    var fallbackSearchCol = headers['fallbackSearch'];
    var snippetCol = headers['snippet'];
    var dateCol = headers['date'];

    var now = new Date();

    // Collect rows to process and batch updates
    var rowsToMark = [];  // Row numbers to mark as processed
    var contactUpdates = {};  // Email -> update data
    var remindersToCreate = [];

    // Pre-load existing reminders for deduplication
    var existingReminders = loadExistingReminders(remindersSheet);

    for (var i = 0; i < data.length && rowsToMark.length < CRM_CONFIG.BATCH_SIZE; i++) {
      var row = data[i];
      var processedAt = row[processedAtCol - 1];

      // Skip if already processed
      if (processedAt) continue;

      stats.processed++;
      rowsToMark.push(i + 2);

      var direction = row[directionCol - 1];
      var contactEmail = row[contactEmailCol - 1];
      var subject = row[subjectCol - 1];
      var gmailSearchLink = row[gmailSearchLinkCol - 1];
      var fallbackSearch = row[fallbackSearchCol - 1];
      var snippet = row[snippetCol - 1];
      var emailDate = row[dateCol - 1];

      // Collect contact update (will dedupe by email)
      contactUpdates[contactEmail.toLowerCase()] = {
        email: contactEmail,
        lastSeenAt: emailDate,
        lastDirection: direction,
        lastSubject: subject,
        lastMessageDate: emailDate
      };

      // If direction is "Received", queue reminder creation
      if (direction === 'Received') {
        var reminderKey = contactEmail.toLowerCase() + '|' + subject;
        if (!existingReminders[reminderKey]) {
          remindersToCreate.push({
            contactEmail: contactEmail,
            contactName: '',
            type: 'email-response',
            gmailSearchLink: gmailSearchLink,
            fallbackSearch: fallbackSearch,
            subject: subject,
            snippet: snippet
          });
          existingReminders[reminderKey] = true;  // Prevent duplicates in same batch
        }
      }
    }

    // Batch update contacts
    if (Object.keys(contactUpdates).length > 0) {
      stats.contactsUpdated = batchUpdateContacts(contactsSheet, contactUpdates);
    }

    // Batch create reminders
    if (remindersToCreate.length > 0) {
      batchCreateReminders(remindersSheet, remindersToCreate);
      stats.remindersCreated = remindersToCreate.length;
    }

    // Batch mark as processed
    if (rowsToMark.length > 0) {
      for (var j = 0; j < rowsToMark.length; j++) {
        data[rowsToMark[j] - 2][processedAtCol - 1] = now;
      }
      interactionsSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }

    logRunEnd(runId, 'success', stats);

  } catch (e) {
    logRunEnd(runId, 'error', stats, e.toString());
  }

  return stats;
}

/**
 * Load existing reminders into a lookup map for fast deduplication
 */
function loadExistingReminders(remindersSheet) {
  var map = {};
  var lastRow = remindersSheet.getLastRow();
  if (lastRow < 2) return map;

  var headers = getSheetHeaders(remindersSheet);
  var data = remindersSheet.getRange(2, 1, lastRow - 1, remindersSheet.getLastColumn()).getValues();

  var emailCol = headers['contactEmail'] - 1;
  var subjectCol = headers['subject'] - 1;
  var statusCol = headers['status'] - 1;

  for (var i = 0; i < data.length; i++) {
    var status = data[i][statusCol];
    if (status !== CRM_CONFIG.STATUS.SENT && status !== CRM_CONFIG.STATUS.ABANDONED) {
      var key = (data[i][emailCol] || '').toLowerCase() + '|' + data[i][subjectCol];
      map[key] = true;
    }
  }
  return map;
}

/**
 * Batch create multiple reminders at once
 */
function batchCreateReminders(remindersSheet, reminders) {
  if (reminders.length === 0) return;

  var now = new Date();
  var rows = [];

  for (var i = 0; i < reminders.length; i++) {
    var data = reminders[i];
    var reminderId = 'R' + now.getTime() + i;
    var idempotencyKey = reminderId + '-0';

    rows.push([
      reminderId,
      data.contactEmail,
      data.contactName || '',
      data.type,
      CRM_CONFIG.STATUS.QUEUED,
      now,
      '',
      0,
      '',
      data.gmailSearchLink || '',
      data.fallbackSearch || '',
      data.subject || '',
      data.snippet || '',
      now,
      '',
      idempotencyKey,
      ''
    ]);
  }

  var lastRow = remindersSheet.getLastRow();
  remindersSheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Batch update contacts - single write operation
 */
function batchUpdateContacts(contactsSheet, updates) {
  var lastRow = contactsSheet.getLastRow();
  var headers = getSheetHeaders(contactsSheet);
  var emailCol = headers['email'];
  var updated = 0;

  // Load existing contacts
  var existingData = [];
  var emailToRow = {};
  if (lastRow >= 2) {
    existingData = contactsSheet.getRange(2, 1, lastRow - 1, contactsSheet.getLastColumn()).getValues();
    for (var i = 0; i < existingData.length; i++) {
      emailToRow[existingData[i][emailCol - 1].toLowerCase()] = i;
    }
  }

  var newContacts = [];

  for (var email in updates) {
    var data = updates[email];
    var rowIdx = emailToRow[email];

    if (rowIdx !== undefined) {
      // Update existing row in memory
      if (headers['lastSeenAt']) existingData[rowIdx][headers['lastSeenAt'] - 1] = data.lastSeenAt;
      if (headers['lastDirection']) existingData[rowIdx][headers['lastDirection'] - 1] = data.lastDirection;
      if (headers['lastSubject']) existingData[rowIdx][headers['lastSubject'] - 1] = data.lastSubject;
      if (headers['lastMessageDate']) existingData[rowIdx][headers['lastMessageDate'] - 1] = data.lastMessageDate;
      updated++;
    } else {
      // Queue new contact
      var contactId = 'C' + new Date().getTime() + newContacts.length;
      newContacts.push([
        contactId,
        data.email,
        '',
        '',
        data.lastSeenAt || new Date(),
        data.lastDirection || '',
        data.lastSubject || '',
        data.lastMessageDate || '',
        ''
      ]);
      updated++;
    }
  }

  // Write updates back
  if (existingData.length > 0) {
    contactsSheet.getRange(2, 1, existingData.length, existingData[0].length).setValues(existingData);
  }

  // Append new contacts
  if (newContacts.length > 0) {
    contactsSheet.getRange(lastRow + 1, 1, newContacts.length, newContacts[0].length).setValues(newContacts);
  }

  return updated;
}

// ============================================================================
// ARCHIVE - Move old interactions to archive
// ============================================================================

/**
 * Archive interactions older than 90 days
 * Run monthly via time trigger
 */
function archiveOldInteractions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var interactionsSheet = ss.getSheetByName(CRM_CONFIG.TABS.INTERACTIONS);

  if (!interactionsSheet) return;

  // Create archive sheet if needed
  var archiveSheet = ss.getSheetByName(CRM_CONFIG.TABS.ARCHIVE);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(CRM_CONFIG.TABS.ARCHIVE);
    // Copy headers
    var headers = interactionsSheet.getRange(1, 1, 1, interactionsSheet.getLastColumn()).getValues();
    archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    archiveSheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold');
    archiveSheet.setFrozenRows(1);
  }

  var settings = getStableSettings();
  var archiveDays = parseInt(settings['ARCHIVE_DAYS']) || CRM_CONFIG.ARCHIVE_DAYS;
  var cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - archiveDays);

  var lastRow = interactionsSheet.getLastRow();
  if (lastRow < 2) return;

  var headers = getSheetHeaders(interactionsSheet);
  var dateCol = headers['date'];
  var data = interactionsSheet.getRange(2, 1, lastRow - 1, interactionsSheet.getLastColumn()).getValues();

  var rowsToArchive = [];
  var rowsToDelete = [];

  for (var i = data.length - 1; i >= 0; i--) {
    var rowDate = new Date(data[i][dateCol - 1]);
    if (rowDate < cutoffDate) {
      rowsToArchive.push(data[i]);
      rowsToDelete.push(i + 2);  // Row number in sheet
    }
  }

  if (rowsToArchive.length > 0) {
    // Reverse rowsToArchive so it's in chronological order when appended
    rowsToArchive.reverse();

    // Append to archive
    archiveSheet.getRange(archiveSheet.getLastRow() + 1, 1, rowsToArchive.length, rowsToArchive[0].length)
      .setValues(rowsToArchive);

    // Delete from main sheet (from bottom to top - rowsToDelete is already in descending order)
    for (var j = 0; j < rowsToDelete.length; j++) {
      interactionsSheet.deleteRow(rowsToDelete[j]);
    }

    Logger.log('Archived ' + rowsToArchive.length + ' interactions');
  }
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Get headers as a map: {headerName: columnNumber}
 */
function getSheetHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) map[String(headers[i])] = i + 1;
  }
  return map;
}

/**
 * Get settings from CRM_Settings sheet
 */
function getStableSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingsSheet = ss.getSheetByName(CRM_CONFIG.TABS.SETTINGS);
  if (!settingsSheet) return {};

  var lastRow = settingsSheet.getLastRow();
  if (lastRow < 2) return {};

  var data = settingsSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var settings = {};
  for (var i = 0; i < data.length; i++) {
    if (data[i][0]) {
      settings[data[i][0]] = data[i][1];
    }
  }
  return settings;
}

/**
 * Log run start
 */
function logRunStart(stage) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var runsSheet = ss.getSheetByName(CRM_CONFIG.TABS.RUNS);
  if (!runsSheet) return null;

  var runId = 'RUN' + new Date().getTime();
  runsSheet.appendRow([
    runId,
    stage,
    new Date(),
    '',  // completedAt
    'running',
    0, 0, 0
  ]);
  return runId;
}

/**
 * Log run end
 */
function logRunEnd(runId, status, stats, errorMsg) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var runsSheet = ss.getSheetByName(CRM_CONFIG.TABS.RUNS);
  if (!runsSheet || !runId) return;

  var lastRow = runsSheet.getLastRow();
  var headers = getSheetHeaders(runsSheet);
  var data = runsSheet.getRange(2, 1, lastRow - 1, 1).getValues();

  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === runId) {
      var row = i + 2;
      runsSheet.getRange(row, headers['completedAt']).setValue(new Date());
      runsSheet.getRange(row, headers['status']).setValue(status + (errorMsg ? ': ' + errorMsg.substring(0, 100) : ''));
      runsSheet.getRange(row, headers['itemsProcessed']).setValue(stats.processed || 0);
      runsSheet.getRange(row, headers['itemsSucceeded']).setValue(stats.remindersCreated || stats.sent || 0);
      runsSheet.getRange(row, headers['itemsFailed']).setValue(stats.errors || 0);
      return;
    }
  }
}

// ============================================================================
// TRIGGERS
// ============================================================================

/**
 * Setup automatic triggers for the CRM pipeline
 * - processInteractions: every 15 minutes
 * - archiveOldInteractions: monthly
 */
function setupTriggers() {
  var ui = SpreadsheetApp.getUi();

  // Remove existing triggers for these functions first
  var existingTriggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < existingTriggers.length; i++) {
    var funcName = existingTriggers[i].getHandlerFunction();
    if (funcName === 'processInteractions' || funcName === 'archiveOldInteractions') {
      ScriptApp.deleteTrigger(existingTriggers[i]);
      removed++;
    }
  }

  // Create new triggers
  ScriptApp.newTrigger('processInteractions')
    .timeBased()
    .everyMinutes(15)
    .create();

  ScriptApp.newTrigger('archiveOldInteractions')
    .timeBased()
    .onMonthDay(1)
    .atHour(3)
    .create();

  ui.alert('Triggers Setup Complete',
    'Removed ' + removed + ' existing trigger(s).\n\n' +
    'Created:\n' +
    '• processInteractions: every 15 minutes\n' +
    '• archiveOldInteractions: 1st of each month at 3am',
    ui.ButtonSet.OK);
}

/**
 * Remove all CRM triggers
 */
function removeTriggers() {
  var existingTriggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  for (var i = 0; i < existingTriggers.length; i++) {
    var funcName = existingTriggers[i].getHandlerFunction();
    if (funcName === 'processInteractions' || funcName === 'archiveOldInteractions') {
      ScriptApp.deleteTrigger(existingTriggers[i]);
      removed++;
    }
  }
  SpreadsheetApp.getUi().alert('Removed ' + removed + ' trigger(s).');
}

// ============================================================================
// MENU
// ============================================================================

/**
 * Add Stable CRM menu items
 */
function addStableCRMMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 Stable CRM')
    .addItem('📋 Setup Tabs', 'setupStableCRM')
    .addItem('⏰ Setup Triggers', 'setupTriggers')
    .addItem('🚫 Remove Triggers', 'removeTriggers')
    .addSeparator()
    .addItem('🔄 Process Interactions (Manual)', 'processInteractions')
    .addItem('📦 Archive Old Interactions', 'archiveOldInteractions')
    .addSeparator()
    .addItem('📊 View Run History', 'viewRunHistory')
    .addToUi();
}

/**
 * View run history
 */
function viewRunHistory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var runsSheet = ss.getSheetByName(CRM_CONFIG.TABS.RUNS);

  if (!runsSheet) {
    SpreadsheetApp.getUi().alert('CRM_Runs sheet not found. Run Setup first.');
    return;
  }

  ss.setActiveSheet(runsSheet);
}

// Add menu on open
function onOpen() {
  addStableCRMMenu();
}
