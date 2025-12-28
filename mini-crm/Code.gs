/**
 * SIMPLE GMAIL CRM
 *
 * Modular architecture for easy extension:
 * - INTEGRATIONS object: Add new integrations (Gong, Slack, etc.) here
 * - Each integration has: columns, sync function, enabled flag
 *
 * Sheets:
 * - Sheet1 = Contacts (one row per contact)
 * - Email_History = One row per contact with ALL email history
 * - CRM_Settings = Configuration
 */

// ============================================================================
// CONFIG
// ============================================================================

var CONFIG = {
  CONTACTS_SHEET: 'Sheet1',
  HISTORY_SHEET: 'Email_History',
  SETTINGS_SHEET: 'CRM_Settings',
  LOG_SHEET: 'CRM_Log',
  EMAIL_COLUMN: 3,  // Column C
  DATA_START_ROW: 2,
  GMAIL_QUERY: 'newer_than:2y',
  MAX_THREADS: 200,  // Reduced for performance (runs faster, can run multiple times)
  CALENDAR_DAYS_BACK: 730,
  MAX_LOG_ROWS: 100,

  // Follow-up configuration
  CRM_LABELS: {
    NEEDS_RESPONSE: 'CRM/Needs-Response',
    FOLLOW_UP: 'CRM/Follow-Up',
    WAITING_REPLY: 'CRM/Waiting-Reply'
  },
  MY_EMAILS: [],  // Will be populated from Session + settings
  MAX_THREADS_PER_LABEL: 2000,
  PAGE_SIZE: 100
};

// ============================================================================
// MODULAR INTEGRATIONS
// Add new integrations here. Each needs: columns, enabled, sync function
// ============================================================================

var INTEGRATIONS = {
  // Core CRM columns (always enabled)
  core: {
    enabled: true,
    columns: ['CRM_ID', 'CRM_HistoryLink']
  },

  // Gmail integration
  gmail: {
    enabled: true,
    columns: ['CRM_LastDate', 'CRM_LastSubject', 'CRM_LastPreview', 'CRM_Direction']
  },

  // Calendar integration
  calendar: {
    enabled: true,
    columns: ['CRM_MeetingDate', 'CRM_MeetingTitle', 'CRM_MeetingHost', 'CRM_MeetingParticipants']
  },

  // Gong integration (disabled - ready to enable later)
  gong: {
    enabled: false,
    columns: ['CRM_GongLink', 'CRM_LastCallDate', 'CRM_LastCallTitle'],
    // To enable: set enabled: true and implement syncGong() function
  },

  // Campaign tracking (columns added dynamically when contact joins a campaign)
  campaigns: {
    enabled: true,
    columns: []  // CRM_CampaignName, CRM_CampaignStatus added only when needed
  },

  // Follow-up tracking (enhanced with timestamps and aging)
  followups: {
    enabled: true,
    columns: [
      'CRM_FollowUpStatus',      // Values: Needs Response, Follow-Up, Waiting Reply
      'CRM_FollowUpStatusSetAt', // Datetime when status was set
      'CRM_LastInboundAt',       // Last email received from contact
      'CRM_LastOutboundAt',      // Last email sent to contact
      'CRM_DaysInStatus',        // Days since status was set
      'CRM_FollowUpSource',      // Label + subject snippet
      'CRM_LastThreadId'         // Gmail thread ID for the winning thread
    ]
  }

  // Add more integrations here:
  // slack: { enabled: false, columns: ['CRM_SlackChannel'] },
  // hubspot: { enabled: false, columns: ['CRM_HubspotLink'] },
};

// ============================================================================
// MENU
// ============================================================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('📧 Gmail CRM')
    .addItem('🔄 Sync All', 'syncAll')
    .addItem('📧 Sync Gmail Only', 'syncGmail')
    .addItem('📅 Sync Calendar Only', 'syncCalendar')
    .addItem('⚙️ Setup', 'setup')
    .addSeparator()
    .addSubMenu(ui.createMenu('📢 Campaigns')
      .addItem('New Campaign', 'createCampaign')
      .addItem('Sync Campaign Status', 'syncCampaigns')
      .addItem('View Campaigns', 'viewCampaigns'))
    .addSubMenu(ui.createMenu('🔔 Follow-ups')
      .addItem('Setup Follow-up Labels', 'setupFollowUpLabels')
      .addItem('Scan for Needs Response', 'scanNeedsResponse')
      .addItem('Sync Follow-up Status', 'syncFollowUps')
      .addSeparator()
      .addItem('Apply Formatting & Validation', 'setupFollowUpFormatting'))
    .addSubMenu(ui.createMenu('🚫 Exclude Settings')
      .addItem('Exclude Emails', 'addExcludeEmail')
      .addItem('Exclude Domains', 'addExcludeDomain')
      .addItem('Exclude Subject Keywords', 'addExcludeSubject')
      .addItem('Exclude Internal Only', 'toggleExcludeInternal')
      .addItem('Exclude Promotional', 'toggleExcludePromotional')
      .addSeparator()
      .addItem('Clear All Excludes', 'clearAllExcludes')
      .addItem('View All Settings', 'viewSettings'))
    .addItem('📅 Set Gmail Date Range', 'setGmailDateRange')
    .addSeparator()
    .addItem('📊 Stats', 'showStats')
    .addItem('📋 View Logs', 'viewLogs')
    .addItem('🔧 Debug Headers', 'debugHeaders')
    .addItem('🆔 Assign Missing IDs', 'assignMissingIds');

  menu.addToUi();
}

// ============================================================================
// SETUP
// ============================================================================

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Collect all enabled columns from integrations
  var allColumns = [];
  for (var key in INTEGRATIONS) {
    if (INTEGRATIONS[key].enabled) {
      allColumns = allColumns.concat(INTEGRATIONS[key].columns);
    }
  }

  // Setup Sheet1 columns
  var contacts = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  if (contacts) {
    addHeadersIfMissing(contacts, allColumns);
  }

  // Setup Email_History sheet
  var history = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!history) {
    history = ss.insertSheet(CONFIG.HISTORY_SHEET);
  }

  var historyHeaders = ['Contact_Email', 'ContactLink', 'TotalEmails', 'EmailHistory'];
  var existingHeaders = history.getLastColumn() > 0 ? history.getRange(1, 1, 1, history.getLastColumn()).getValues()[0] : [];

  if (existingHeaders.length === 0 || existingHeaders[0] !== 'Contact_Email') {
    history.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]);
    history.getRange(1, 1, 1, historyHeaders.length).setFontWeight('bold').setBackground('#e8f0fe');
    history.setFrozenRows(1);
    history.setColumnWidth(4, 600);
  }

  // Setup Settings sheet
  setupSettingsSheet(ss);

  // Check if internal domain is set
  var settings = getSettings();
  if (!settings['INTERNAL_DOMAIN']) {
    var domainResponse = ui.prompt('Set Internal Domain',
      'Enter your company email domain (e.g., "company.com"):\n\n' +
      'This is used to filter out internal-only email threads.',
      ui.ButtonSet.OK_CANCEL);

    if (domainResponse.getSelectedButton() === ui.Button.OK && domainResponse.getResponseText().trim()) {
      setSetting('INTERNAL_DOMAIN', domainResponse.getResponseText().trim().toLowerCase());
    }
  }

  // Show enabled integrations
  var enabledList = [];
  for (var k in INTEGRATIONS) {
    if (INTEGRATIONS[k].enabled) enabledList.push(k);
  }

  ui.alert('Setup Complete',
    'Columns added for: ' + enabledList.join(', ') + '\n\n' +
    'Sheets:\n- Sheet1 (contacts)\n- Email_History\n- CRM_Settings',
    ui.ButtonSet.OK);
}

function setupSettingsSheet(ss) {
  var settings = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  if (!settings) {
    settings = ss.insertSheet(CONFIG.SETTINGS_SHEET);
    settings.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Description']]);
    settings.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
    settings.setFrozenRows(1);

    settings.appendRow(['EXCLUDE_INTERNAL', 'TRUE', 'Skip emails with only internal addresses']);
    settings.appendRow(['EXCLUDE_EMAILS', '', 'Comma-separated email addresses to exclude']);
    settings.appendRow(['EXCLUDE_SUBJECTS', '', 'Comma-separated keywords to exclude']);
    settings.appendRow(['EXCLUDE_DOMAINS', '', 'Comma-separated domains to exclude']);
    settings.appendRow(['INTERNAL_DOMAIN', '', 'Your company domain']);
    settings.appendRow(['EXCLUDE_PROMOTIONAL', 'TRUE', 'Skip promotional/marketing emails']);
    settings.appendRow(['GMAIL_DAYS_BACK', '730', 'How many days back to sync Gmail (e.g., 365, 730, 1095)']);

    settings.setColumnWidth(2, 300);
    settings.setColumnWidth(3, 400);
  }
}

function addHeadersIfMissing(sheet, headers) {
  var lastCol = sheet.getLastColumn();
  var existing = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var nextCol = lastCol + 1;

  for (var i = 0; i < headers.length; i++) {
    if (existing.indexOf(headers[i]) === -1) {
      sheet.getRange(1, nextCol).setValue(headers[i]).setFontWeight('bold').setBackground('#e8f0fe');
      nextCol++;
    }
  }
}

// ============================================================================
// SETTINGS
// ============================================================================

function getSettings() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var settings = {};
  for (var i = 0; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function setSetting(key, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) {
    if (sheet.getRange(i, 1).getValue() === key) {
      sheet.getRange(i, 2).setValue(value);
      return;
    }
  }
}

function addExcludeEmail() {
  var ui = SpreadsheetApp.getUi();
  var current = getSettings()['EXCLUDE_EMAILS'] || '';

  var response = ui.prompt('Exclude Email Patterns',
    'Enter email addresses or patterns (comma-separated).\n\n' +
    'Examples:\n' +
    '- Full email: spam@company.com\n' +
    '- Prefix: no-reply@, noreply@, notifications@\n' +
    '- Suffix: @newsletter.com\n\n' +
    'Current: ' + (current || '(none)'),
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK && response.getResponseText().trim()) {
    var combined = current ? current + ',' + response.getResponseText().trim().toLowerCase() : response.getResponseText().trim().toLowerCase();
    setSetting('EXCLUDE_EMAILS', combined);
    ui.alert('Updated', 'Exclude email patterns: ' + combined, ui.ButtonSet.OK);
  }
}

function addExcludeSubject() {
  var ui = SpreadsheetApp.getUi();
  var current = getSettings()['EXCLUDE_SUBJECTS'] || '';

  var response = ui.prompt('Exclude Subject Keywords',
    'Enter keywords (comma-separated).\n\nCurrent: ' + (current || '(none)'),
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK && response.getResponseText().trim()) {
    var combined = current ? current + ',' + response.getResponseText().trim() : response.getResponseText().trim();
    setSetting('EXCLUDE_SUBJECTS', combined);
    ui.alert('Updated', 'Exclude subjects: ' + combined, ui.ButtonSet.OK);
  }
}

function addExcludeDomain() {
  var ui = SpreadsheetApp.getUi();
  var current = getSettings()['EXCLUDE_DOMAINS'] || '';

  var response = ui.prompt('Exclude Domains',
    'Enter domains (comma-separated).\n\nCurrent: ' + (current || '(none)'),
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK && response.getResponseText().trim()) {
    var combined = current ? current + ',' + response.getResponseText().trim() : response.getResponseText().trim();
    setSetting('EXCLUDE_DOMAINS', combined);
    ui.alert('Updated', 'Exclude domains: ' + combined, ui.ButtonSet.OK);
  }
}

function toggleExcludeInternal() {
  var ui = SpreadsheetApp.getUi();
  var settings = getSettings();
  var internalDomain = settings['INTERNAL_DOMAIN'] || '';

  if (!internalDomain) {
    var response = ui.prompt('Set Internal Domain', 'Enter your company domain:', ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() === ui.Button.OK && response.getResponseText().trim()) {
      setSetting('INTERNAL_DOMAIN', response.getResponseText().trim().toLowerCase());
      setSetting('EXCLUDE_INTERNAL', 'TRUE');
      ui.alert('Updated', 'Internal domain set. Internal emails will be excluded.', ui.ButtonSet.OK);
    }
  } else {
    var newValue = settings['EXCLUDE_INTERNAL'] === 'TRUE' ? 'FALSE' : 'TRUE';
    setSetting('EXCLUDE_INTERNAL', newValue);
    ui.alert('Updated', 'Exclude internal: ' + newValue, ui.ButtonSet.OK);
  }
}

function toggleExcludePromotional() {
  var ui = SpreadsheetApp.getUi();
  var settings = getSettings();
  var current = settings['EXCLUDE_PROMOTIONAL'];

  // If setting doesn't exist, add it
  if (current === undefined) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
    if (sheet) {
      sheet.appendRow(['EXCLUDE_PROMOTIONAL', 'TRUE', 'Skip promotional/marketing emails']);
    }
    ui.alert('Updated', 'Promotional emails will be excluded.', ui.ButtonSet.OK);
    return;
  }

  var newValue = current === 'TRUE' ? 'FALSE' : 'TRUE';
  setSetting('EXCLUDE_PROMOTIONAL', newValue);
  ui.alert('Updated', 'Exclude promotional: ' + newValue, ui.ButtonSet.OK);
}

function clearAllExcludes() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Clear All Excludes',
    'This will clear:\n' +
    '- Exclude emails\n' +
    '- Exclude subject keywords\n' +
    '- Exclude domains\n\n' +
    'Internal domain and toggle settings will be kept.\n\nContinue?',
    ui.ButtonSet.YES_NO);

  if (response === ui.Button.YES) {
    setSetting('EXCLUDE_EMAILS', '');
    setSetting('EXCLUDE_SUBJECTS', '');
    setSetting('EXCLUDE_DOMAINS', '');
    ui.alert('Cleared', 'Exclude emails, subjects, and domains have been cleared.', ui.ButtonSet.OK);
  }
}

function viewSettings() {
  var settings = getSettings();
  var gmailDays = settings['GMAIL_DAYS_BACK'] || '730';
  var years = (parseInt(gmailDays, 10) / 365).toFixed(1);
  SpreadsheetApp.getUi().alert('Settings',
    'Exclude Internal: ' + (settings['EXCLUDE_INTERNAL'] || 'FALSE') + '\n' +
    'Internal Domain: ' + (settings['INTERNAL_DOMAIN'] || '(not set)') + '\n' +
    'Exclude Emails: ' + (settings['EXCLUDE_EMAILS'] || '(none)') + '\n' +
    'Exclude Subjects: ' + (settings['EXCLUDE_SUBJECTS'] || '(none)') + '\n' +
    'Exclude Domains: ' + (settings['EXCLUDE_DOMAINS'] || '(none)') + '\n' +
    'Exclude Promotional: ' + (settings['EXCLUDE_PROMOTIONAL'] !== 'FALSE' ? 'TRUE' : 'FALSE') + '\n' +
    'Gmail Date Range: ' + gmailDays + ' days (~' + years + ' years)',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function setGmailDateRange() {
  var ui = SpreadsheetApp.getUi();
  var settings = getSettings();
  var currentDays = settings['GMAIL_DAYS_BACK'] || '730';

  var response = ui.prompt('Set Gmail Date Range',
    'Enter how many days back to sync Gmail emails.\n\n' +
    'Examples:\n' +
    '- 365 = 1 year\n' +
    '- 730 = 2 years\n' +
    '- 1095 = 3 years\n' +
    '- 1825 = 5 years\n\n' +
    'Current: ' + currentDays + ' days\n\n' +
    'Note: Larger ranges will take longer to sync.',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() === ui.Button.OK) {
    var days = parseInt(response.getResponseText().trim(), 10);
    if (isNaN(days) || days < 1) {
      ui.alert('Error', 'Please enter a valid number of days (e.g., 365, 730, 1095).', ui.ButtonSet.OK);
      return;
    }

    // Add or update the setting
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var settingsSheet = ss.getSheetByName(CONFIG.SETTINGS_SHEET);

    if (settingsSheet) {
      var found = false;
      var lastRow = settingsSheet.getLastRow();
      for (var i = 2; i <= lastRow; i++) {
        if (settingsSheet.getRange(i, 1).getValue() === 'GMAIL_DAYS_BACK') {
          settingsSheet.getRange(i, 2).setValue(String(days));
          found = true;
          break;
        }
      }
      if (!found) {
        settingsSheet.appendRow(['GMAIL_DAYS_BACK', String(days), 'How many days back to sync Gmail (e.g., 365, 730, 1095)']);
      }
    }

    var years = (days / 365).toFixed(1);
    ui.alert('Date Range Updated',
      'Gmail sync will now go back ' + days + ' days (~' + years + ' years).\n\n' +
      'Run "Sync Gmail" to pull in older emails.',
      ui.ButtonSet.OK);
  }
}

// ============================================================================
// SYNC ALL
// ============================================================================

function syncAll() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Syncing...', 'Gmail CRM', -1);

  var results = [];

  if (INTEGRATIONS.gmail.enabled) {
    ss.toast('Syncing Gmail...', 'Gmail CRM', -1);
    var gmail = syncGmailInternal();
    results.push('GMAIL: ' + gmail.processed + ' threads, ' + gmail.newContacts + ' new, ' + gmail.updated + ' updated');
  }

  if (INTEGRATIONS.calendar.enabled) {
    ss.toast('Syncing Calendar...', 'Gmail CRM', -1);
    var cal = syncCalendarInternal();
    results.push('CALENDAR: ' + cal.events + ' events, ' + cal.updated + ' contacts updated');
  }

  if (INTEGRATIONS.campaigns.enabled) {
    ss.toast('Syncing Campaigns...', 'Gmail CRM', -1);
    var camp = syncCampaignsInternal();
    results.push('CAMPAIGNS: ' + camp.campaigns + ' campaigns, ' + camp.updated + ' contacts updated');
  }

  if (INTEGRATIONS.followups.enabled) {
    ss.toast('Syncing Follow-ups...', 'Gmail CRM', -1);
    var followup = syncFollowUpsInternal();
    results.push('FOLLOW-UPS: ' + followup.followups + ' follow-ups, ' + followup.needsResponse + ' need response');
  }

  // Add more integrations here when enabled
  // if (INTEGRATIONS.gong.enabled) { ... }

  ss.toast('Sync complete!', 'Gmail CRM', 3);
  ui.alert('Sync Complete', results.join('\n\n'), ui.ButtonSet.OK);
}

// ============================================================================
// GMAIL SYNC
// ============================================================================

function syncGmail() {
  var result = syncGmailInternal();
  SpreadsheetApp.getUi().alert('Gmail Sync Complete',
    'Threads scanned: ' + result.processed + '\n' +
    'New contacts: ' + result.newContacts + '\n' +
    'Existing updated: ' + result.updated + '\n' +
    'Excluded: ' + result.excluded + '\n' +
    'Duplicates removed: ' + result.duplicates + '\n' +
    'Blocked removed: ' + result.removed + '\n' +
    'Contacts in sheet: ' + result.totalContacts + '\n' +
    'Header columns: ' + result.headerDebug,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function debugHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contactsSheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var headers = getHeaderMap(contactsSheet);
  var emailIndex = buildEmailIndex(contactsSheet);
  var myEmail = Session.getActiveUser().getEmail().toLowerCase();

  var results = [];
  results.push('My email: ' + myEmail);
  results.push('CRM_LastDate: Col ' + headers['CRM_LastDate']);
  results.push('---');

  // Get a few Gmail threads and show what emails we find
  var threads = GmailApp.search(CONFIG.GMAIL_QUERY, 0, 5);
  results.push('Gmail threads found: ' + threads.length);

  var matchCount = 0;
  var writeCount = 0;

  for (var i = 0; i < threads.length && i < 3; i++) {
    var messages = threads[i].getMessages();
    var contactEmail = findExternalEmail(messages, myEmail);
    var subject = threads[i].getFirstMessageSubject() || '(no subject)';

    results.push('---');
    results.push('Thread ' + (i+1) + ': ' + subject.substring(0, 30));
    results.push('Contact found: ' + (contactEmail || 'NONE'));

    if (contactEmail) {
      var row = emailIndex[contactEmail.toLowerCase()];
      results.push('Row in sheet: ' + (row || 'NOT IN SHEET'));

      if (row && headers['CRM_LastDate']) {
        matchCount++;
        // Actually write to test
        contactsSheet.getRange(row, headers['CRM_LastDate']).setValue(new Date());
        contactsSheet.getRange(row, headers['CRM_LastSubject']).setValue(subject);
        contactsSheet.getRange(row, headers['CRM_Direction']).setValue('Test');
        writeCount++;
        results.push('WROTE to row ' + row);
      }
    }
  }

  results.push('---');
  results.push('Matches: ' + matchCount + ', Writes: ' + writeCount);

  SpreadsheetApp.getUi().alert('Debug', results.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncGmailInternal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contactsSheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var historySheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);

  if (!historySheet) return { processed: 0, newContacts: 0, updated: 0, excluded: 0, removed: 0, duplicates: 0 };

  var headers = getHeaderMap(contactsSheet);
  if (!headers['CRM_ID']) return { processed: 0, newContacts: 0, updated: 0, excluded: 0, removed: 0, duplicates: 0 };

  // Load exclusion settings
  var settings = getSettings();
  var excludeInternal = settings['EXCLUDE_INTERNAL'] === 'TRUE';
  var internalDomain = (settings['INTERNAL_DOMAIN'] || '').toLowerCase().trim();
  var excludeEmails = parseList(settings['EXCLUDE_EMAILS']);
  var excludeSubjects = parseList(settings['EXCLUDE_SUBJECTS']);
  var excludeDomains = parseList(settings['EXCLUDE_DOMAINS']);
  var excludePromotional = settings['EXCLUDE_PROMOTIONAL'] !== 'FALSE'; // Default TRUE

  // Get Gmail date range from settings (default to CONFIG value)
  var gmailDaysBack = parseInt(settings['GMAIL_DAYS_BACK'], 10) || 730;
  var gmailQuery = 'newer_than:' + gmailDaysBack + 'd';

  // Remove duplicates first
  var duplicatesRemoved = removeDuplicateContacts(contactsSheet, historySheet);

  // Remove existing contacts that now match exclude patterns
  var removed = removeExcludedContacts(contactsSheet, historySheet, excludeEmails, excludeDomains);

  // Build indexes (after cleanup)
  var emailIndex = buildEmailIndex(contactsSheet);
  var historyIndex = buildHistoryIndex(historySheet);
  var historyHeaders = getHeaderMap(historySheet);

  var myEmail = Session.getActiveUser().getEmail().toLowerCase();
  var threads = GmailApp.search(gmailQuery, 0, CONFIG.MAX_THREADS);

  // Debug: check which header columns exist
  var headerDebug = 'ID:' + (headers['CRM_LastDate'] ? headers['CRM_LastDate'] : 'X');

  var stats = { processed: 0, newContacts: 0, updated: 0, excluded: 0, removed: removed, duplicates: duplicatesRemoved, totalContacts: Object.keys(emailIndex).length, headerDebug: headerDebug };
  var sheetUrl = ss.getUrl();

  // Track new contacts to add
  var newContactRows = [];

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var messages = thread.getMessages();
    if (!messages || messages.length === 0) continue;

    // Get ALL external emails from this thread (not just one)
    var allExternalEmails = findAllExternalEmails(messages, myEmail);
    if (allExternalEmails.length === 0) continue;

    var subject = (thread.getFirstMessageSubject() || '').toLowerCase();
    var lastMsg = messages[messages.length - 1];
    var fromEmail = extractEmail(lastMsg.getFrom());
    var direction = (fromEmail === myEmail) ? 'Sent' : 'Received';

    // Process each external email found in the thread
    var threadProcessed = false;
    for (var e = 0; e < allExternalEmails.length; e++) {
      var contactEmail = allExternalEmails[e];
      var domain = contactEmail.split('@')[1] || '';
      var normalized = contactEmail.toLowerCase().trim();

      // Apply exclusions
      if (shouldExclude(normalized, subject, domain, excludeEmails, excludeSubjects, excludeDomains, excludeInternal, internalDomain, excludePromotional, thread)) {
        stats.excluded++;
        continue;
      }

      // Check if this contact exists in our sheet
      var contactRow = emailIndex[normalized];

      if (contactRow) {
        // Contact exists - update their Gmail columns
        stats.updated++;

        // Ensure CRM_ID exists for existing contacts
        if (headers['CRM_ID']) {
          var existingId = contactsSheet.getRange(contactRow, headers['CRM_ID']).getValue();
          if (!existingId) {
            contactsSheet.getRange(contactRow, headers['CRM_ID']).setValue('C' + Date.now() + i + e);
          }
        }

        // Update Gmail columns directly using setValue
        if (headers['CRM_LastDate']) {
          contactsSheet.getRange(contactRow, headers['CRM_LastDate']).setValue(lastMsg.getDate());
        }
        if (headers['CRM_LastSubject']) {
          contactsSheet.getRange(contactRow, headers['CRM_LastSubject']).setValue(thread.getFirstMessageSubject() || '');
        }
        if (headers['CRM_LastPreview']) {
          contactsSheet.getRange(contactRow, headers['CRM_LastPreview']).setValue((lastMsg.getPlainBody() || '').substring(0, 200).replace(/\s+/g, ' '));
        }
        if (headers['CRM_Direction']) {
          contactsSheet.getRange(contactRow, headers['CRM_Direction']).setValue(direction);
        }

        // Update Email History
        updateEmailHistory(historySheet, historyHeaders, normalized, contactEmail, lastMsg, thread, direction, historyIndex);

        // Update history link
        var historyRow = historyIndex[normalized];
        var contactLinkCol = historyHeaders['ContactLink'];
        if (historyRow && contactLinkCol && headers['CRM_HistoryLink']) {
          var historyLink = sheetUrl + '#gid=' + historySheet.getSheetId() + '&range=A' + historyRow;
          contactsSheet.getRange(contactRow, headers['CRM_HistoryLink']).setValue(historyLink);
          historySheet.getRange(historyRow, contactLinkCol).setValue(
            sheetUrl + '#gid=' + contactsSheet.getSheetId() + '&range=A' + contactRow
          );
        }

        threadProcessed = true;
      } else if (e === 0) {
        // Only add new contact for the PRIMARY external email (first one found)
        // This prevents creating hundreds of contacts from CC lists
        contactRow = contactsSheet.getLastRow() + 1 + newContactRows.length;
        newContactRows.push({ row: contactRow, email: contactEmail, id: 'C' + Date.now() + i });
        emailIndex[normalized] = contactRow;
        stats.newContacts++;

        // Update Gmail columns for new contact
        if (headers['CRM_LastDate']) {
          contactsSheet.getRange(contactRow, headers['CRM_LastDate']).setValue(lastMsg.getDate());
        }
        if (headers['CRM_LastSubject']) {
          contactsSheet.getRange(contactRow, headers['CRM_LastSubject']).setValue(thread.getFirstMessageSubject() || '');
        }
        if (headers['CRM_LastPreview']) {
          contactsSheet.getRange(contactRow, headers['CRM_LastPreview']).setValue((lastMsg.getPlainBody() || '').substring(0, 200).replace(/\s+/g, ' '));
        }
        if (headers['CRM_Direction']) {
          contactsSheet.getRange(contactRow, headers['CRM_Direction']).setValue(direction);
        }

        // Update Email History
        updateEmailHistory(historySheet, historyHeaders, normalized, contactEmail, lastMsg, thread, direction, historyIndex);

        // Update history link
        var historyRowNew = historyIndex[normalized];
        var contactLinkColNew = historyHeaders['ContactLink'];
        if (historyRowNew && contactLinkColNew && headers['CRM_HistoryLink']) {
          var historyLinkNew = sheetUrl + '#gid=' + historySheet.getSheetId() + '&range=A' + historyRowNew;
          contactsSheet.getRange(contactRow, headers['CRM_HistoryLink']).setValue(historyLinkNew);
          historySheet.getRange(historyRowNew, contactLinkColNew).setValue(
            sheetUrl + '#gid=' + contactsSheet.getSheetId() + '&range=A' + contactRow
          );
        }

        threadProcessed = true;
      }
    }

    if (threadProcessed) {
      stats.processed++;
    }
  }

  // Write new contacts
  for (var n = 0; n < newContactRows.length; n++) {
    var newRow = newContactRows[n];
    contactsSheet.getRange(newRow.row, CONFIG.EMAIL_COLUMN).setValue(newRow.email);
    if (headers['CRM_ID']) {
      contactsSheet.getRange(newRow.row, headers['CRM_ID']).setValue(newRow.id);
    }
  }

  return stats;
}

function updateEmailHistory(sheet, headers, normalized, email, msg, thread, direction, index) {
  // Validate required headers exist
  var emailHistoryCol = headers['EmailHistory'];
  var totalEmailsCol = headers['TotalEmails'];
  var contactEmailCol = headers['Contact_Email'];

  if (!emailHistoryCol || !totalEmailsCol || !contactEmailCol) return;

  var emailEntry = formatEmailEntry(msg, thread, direction);

  // Create a unique key for this email (date + subject) to detect duplicates
  var emailKey = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') +
                 '|' + (thread.getFirstMessageSubject() || '');

  if (index[normalized]) {
    var row = index[normalized];
    if (!row) return;
    var existing = sheet.getRange(row, emailHistoryCol).getValue() || '';

    // Check if this email is already in history (by checking the date/subject key)
    if (existing.indexOf(emailKey.split('|')[0]) > -1 &&
        existing.indexOf(emailKey.split('|')[1]) > -1) {
      // Already logged, skip
      return;
    }

    var newHistory = emailEntry + '\n---\n' + existing;
    var entries = newHistory.split('\n---\n');
    if (entries.length > 10) entries = entries.slice(0, 10);
    sheet.getRange(row, emailHistoryCol).setValue(entries.join('\n---\n'));
    var count = parseInt(sheet.getRange(row, totalEmailsCol).getValue(), 10) || 0;
    sheet.getRange(row, totalEmailsCol).setValue(count + 1);
  } else {
    var newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, contactEmailCol).setValue(email);
    sheet.getRange(newRow, totalEmailsCol).setValue(1);
    sheet.getRange(newRow, emailHistoryCol).setValue(emailEntry);
    index[normalized] = newRow;
  }
}

// ============================================================================
// CALENDAR SYNC
// ============================================================================

function syncCalendar() {
  var result = syncCalendarInternal();
  SpreadsheetApp.getUi().alert('Calendar Sync Complete',
    'Events: ' + result.events + '\n' +
    'Contacts updated: ' + result.updated,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncCalendarInternal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contactsSheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var headers = getHeaderMap(contactsSheet);

  if (!headers['CRM_MeetingDate']) return { events: 0, updated: 0 };

  var emailIndex = buildEmailIndex(contactsSheet);
  var myEmail = Session.getActiveUser().getEmail().toLowerCase();

  var now = new Date();
  var startDate = new Date(now.getTime() - CONFIG.CALENDAR_DAYS_BACK * 24 * 60 * 60 * 1000);

  var events = CalendarApp.getDefaultCalendar().getEvents(startDate, now);
  var stats = { events: 0, updated: 0 };
  var meetingMap = {};

  for (var i = 0; i < events.length; i++) {
    var event = events[i];
    stats.events++;

    var title = event.getTitle() || '(no title)';
    var eventDate = event.getStartTime();
    var guests = event.getGuestList(true);
    var organizer = event.getCreators()[0] || myEmail;

    var participantEmails = [];
    for (var j = 0; j < guests.length; j++) {
      var guestEmail = guests[j].getEmail().toLowerCase();
      if (guestEmail && guestEmail !== myEmail) {
        participantEmails.push(guestEmail);
      }
    }

    for (var k = 0; k < participantEmails.length; k++) {
      var pEmail = participantEmails[k];
      if (!meetingMap[pEmail] || eventDate > meetingMap[pEmail].date) {
        meetingMap[pEmail] = {
          date: eventDate,
          title: title,
          host: organizer,
          participants: participantEmails.join(', ')
        };
      }
    }
  }

  for (var email in meetingMap) {
    var normalized = email.toLowerCase().trim();
    if (emailIndex[normalized]) {
      var row = emailIndex[normalized];
      var m = meetingMap[email];
      setCell(contactsSheet, row, headers['CRM_MeetingDate'], m.date);
      setCell(contactsSheet, row, headers['CRM_MeetingTitle'], m.title);
      setCell(contactsSheet, row, headers['CRM_MeetingHost'], m.host);
      setCell(contactsSheet, row, headers['CRM_MeetingParticipants'], m.participants);

      // Ensure CRM_ID exists for contacts updated via calendar
      if (headers['CRM_ID']) {
        var existingId = contactsSheet.getRange(row, headers['CRM_ID']).getValue();
        if (!existingId) {
          contactsSheet.getRange(row, headers['CRM_ID']).setValue('C' + Date.now() + row);
        }
      }

      stats.updated++;
    }
  }

  return stats;
}

// ============================================================================
// CAMPAIGN TRACKING
// ============================================================================

function createCampaign() {
  var ui = SpreadsheetApp.getUi();

  var response = ui.prompt('Create New Campaign',
    'Enter the exact email subject line for this campaign:\n\n' +
    '(This will be used to track responses)',
    ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var subject = response.getResponseText().trim();
  if (!subject) {
    ui.alert('Error', 'Please enter a subject line.', ui.ButtonSet.OK);
    return;
  }

  var nameResponse = ui.prompt('Campaign Name',
    'Enter a short name for this campaign:\n\n' +
    '(e.g., "Q1 Outreach", "Product Launch")',
    ui.ButtonSet.OK_CANCEL);

  if (nameResponse.getSelectedButton() !== ui.Button.OK) return;

  var campaignName = nameResponse.getResponseText().trim() || subject.substring(0, 30);

  // Generate unique sheet name (handle collisions)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var baseSheetName = 'Campaign_' + campaignName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20);
  var detailSheetName = baseSheetName;
  var counter = 1;

  while (ss.getSheetByName(detailSheetName)) {
    detailSheetName = baseSheetName + '_' + counter;
    counter++;
    if (counter > 100) {
      ui.alert('Error', 'Too many campaigns with similar names. Please use a more unique name.', ui.ButtonSet.OK);
      return;
    }
  }

  // Create or get Campaigns sheet
  var campaignSheet = ss.getSheetByName('Campaigns');
  if (!campaignSheet) {
    campaignSheet = ss.insertSheet('Campaigns');
    campaignSheet.getRange(1, 1, 1, 7).setValues([['Campaign_Name', 'Subject', 'Created', 'Total_Sent', 'Responded', 'No_Response', 'Sheet_Name']]);
    campaignSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8f0fe');
    campaignSheet.setFrozenRows(1);
    campaignSheet.setColumnWidth(2, 300);
  }

  // Add campaign (including sheet name for lookup)
  var newRow = campaignSheet.getLastRow() + 1;
  campaignSheet.getRange(newRow, 1, 1, 7).setValues([[campaignName, subject, new Date(), 0, 0, 0, detailSheetName]]);

  // Create campaign detail sheet (using unique name generated earlier)
  var detailSheet = ss.insertSheet(detailSheetName);
  {
    detailSheet.getRange(1, 1, 1, 5).setValues([['Email', 'Contact_Link', 'Status', 'Last_Updated', 'Response_Date']]);
    detailSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e8f0fe');
    detailSheet.setFrozenRows(1);
    detailSheet.setColumnWidth(1, 250);
    detailSheet.setColumnWidth(3, 120);
  }

  // Populate with contacts who received this campaign (search sent emails)
  var myEmail = Session.getActiveUser().getEmail().toLowerCase();
  var query = 'from:me subject:"' + subject + '"';
  var threads = GmailApp.search(query, 0, 100);

  var recipients = {};
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var msg = messages[j];
      if (extractEmail(msg.getFrom()) === myEmail) {
        var toEmails = msg.getTo().split(',');
        for (var k = 0; k < toEmails.length; k++) {
          var email = extractEmail(toEmails[k]);
          if (email && email !== myEmail) {
            recipients[email] = true;
          }
        }
        var ccEmails = (msg.getCc() || '').split(',');
        for (var l = 0; l < ccEmails.length; l++) {
          var ccEmail = extractEmail(ccEmails[l]);
          if (ccEmail && ccEmail !== myEmail) {
            recipients[ccEmail] = true;
          }
        }
      }
    }
  }

  // Add recipients to detail sheet
  var contactsSheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var emailIndex = buildEmailIndex(contactsSheet);
  var sheetUrl = ss.getUrl();
  var count = 0;

  for (var email in recipients) {
    var contactRow = emailIndex[email.toLowerCase()];
    var contactLink = contactRow ? sheetUrl + '#gid=' + contactsSheet.getSheetId() + '&range=A' + contactRow : '';
    detailSheet.appendRow([email, contactLink, 'No Response', new Date(), '']);
    count++;
  }

  // Update campaign totals
  campaignSheet.getRange(newRow, 4).setValue(count);
  campaignSheet.getRange(newRow, 6).setValue(count);

  // Activate the detail sheet
  ss.setActiveSheet(detailSheet);

  ui.alert('Campaign Created',
    'Campaign: ' + campaignName + '\n' +
    'Subject: ' + subject + '\n' +
    'Recipients found: ' + count + '\n\n' +
    'Run "Sync Campaign Status" to check for responses.',
    ui.ButtonSet.OK);
}

function syncCampaigns() {
  var result = syncCampaignsInternal();
  SpreadsheetApp.getUi().alert('Campaign Sync Complete',
    'Campaigns checked: ' + result.campaigns + '\n' +
    'Contacts updated: ' + result.updated + '\n' +
    'Responses found: ' + result.responses + '\n' +
    'Bounces found: ' + result.bounces,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncCampaignsInternal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var campaignSheet = ss.getSheetByName('Campaigns');

  if (!campaignSheet || campaignSheet.getLastRow() < 2) {
    return { campaigns: 0, updated: 0, responses: 0, bounces: 0 };
  }

  // Get campaigns with sheet names (columns 1-7, or fallback to 1-2 for old format)
  var lastCol = campaignSheet.getLastColumn();
  var campaigns = campaignSheet.getRange(2, 1, campaignSheet.getLastRow() - 1, Math.min(lastCol, 7)).getValues();
  var myEmail = Session.getActiveUser().getEmail().toLowerCase();

  var stats = { campaigns: 0, updated: 0, responses: 0, bounces: 0 };
  var contactsSheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var headers = getHeaderMap(contactsSheet);
  var emailIndex = buildEmailIndex(contactsSheet);

  // Add campaign columns dynamically if they don't exist yet
  var campaignColumnsAdded = false;
  if (!headers['CRM_CampaignName'] || !headers['CRM_CampaignStatus']) {
    addHeadersIfMissing(contactsSheet, ['CRM_CampaignName', 'CRM_CampaignStatus']);
    headers = getHeaderMap(contactsSheet); // Refresh headers
    campaignColumnsAdded = true;
  }

  for (var i = 0; i < campaigns.length; i++) {
    var campaignName = campaigns[i][0];
    var subject = campaigns[i][1];
    if (!subject) continue;

    stats.campaigns++;

    // Use stored sheet name if available (column 7), otherwise generate from name
    var detailSheetName = campaigns[i][6] || ('Campaign_' + String(campaignName).replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20));
    var detailSheet = ss.getSheetByName(detailSheetName);
    if (!detailSheet || detailSheet.getLastRow() < 2) continue;

    // Get all recipients from detail sheet
    var detailData = detailSheet.getRange(2, 1, detailSheet.getLastRow() - 1, 5).getValues();
    var recipientRows = {};
    for (var j = 0; j < detailData.length; j++) {
      var recipientEmail = detailData[j][0];
      if (recipientEmail && String(recipientEmail).indexOf('@') > -1) {
        recipientRows[String(recipientEmail).toLowerCase()] = j + 2;
      }
    }

    // Search for responses (emails TO me with similar subject)
    var responseQuery = 'to:me subject:"' + subject.replace(/^(Re: |Fwd: )/gi, '') + '"';
    var responseThreads = GmailApp.search(responseQuery, 0, 100);

    var responded = 0;
    var noResponse = 0;

    for (var t = 0; t < responseThreads.length; t++) {
      var messages = responseThreads[t].getMessages();
      for (var m = 0; m < messages.length; m++) {
        var fromEmail = extractEmail(messages[m].getFrom());
        var detailRow = recipientRows[fromEmail];
        if (fromEmail && fromEmail !== myEmail && detailRow && detailRow > 0) {
          var currentStatus = safeGetValue(detailSheet, detailRow, 3);
          if (currentStatus !== 'Responded') {
            safeSetValue(detailSheet, detailRow, 3, 'Responded');
            safeSetValue(detailSheet, detailRow, 4, new Date());
            safeSetValue(detailSheet, detailRow, 5, messages[m].getDate());
            stats.responses++;
            stats.updated++;

            // Update contact sheet
            var contactRow = emailIndex[fromEmail];
            if (contactRow && headers['CRM_CampaignName'] && headers['CRM_CampaignStatus']) {
              setCell(contactsSheet, contactRow, headers['CRM_CampaignName'], campaignName);
              setCell(contactsSheet, contactRow, headers['CRM_CampaignStatus'], 'Responded');
            }
          }
          responded++;
          delete recipientRows[fromEmail];
        }
      }
    }

    // Check for bounces
    var bounceQuery = 'from:mailer-daemon subject:"' + subject + '"';
    var bounceThreads = GmailApp.search(bounceQuery, 0, 50);

    for (var b = 0; b < bounceThreads.length; b++) {
      var bounceMessages = bounceThreads[b].getMessages();
      for (var bm = 0; bm < bounceMessages.length; bm++) {
        var body = bounceMessages[bm].getPlainBody() || '';
        // Try to extract bounced email from body
        for (var recipEmail in recipientRows) {
          if (body.toLowerCase().indexOf(recipEmail) > -1) {
            var bRow = recipientRows[recipEmail];
            if (bRow && bRow > 0) {
              safeSetValue(detailSheet, bRow, 3, 'Bounced');
              safeSetValue(detailSheet, bRow, 4, new Date());
              stats.bounces++;
              stats.updated++;

              var cRow = emailIndex[recipEmail];
              if (cRow && headers['CRM_CampaignStatus']) {
                setCell(contactsSheet, cRow, headers['CRM_CampaignStatus'], 'Bounced');
              }
            }
            delete recipientRows[recipEmail];
          }
        }
      }
    }

    // Count remaining as no response
    noResponse = Object.keys(recipientRows).length;

    // Update campaign summary
    var summaryRow = i + 2;
    campaignSheet.getRange(summaryRow, 5).setValue(responded);
    campaignSheet.getRange(summaryRow, 6).setValue(noResponse);
  }

  return stats;
}

function viewCampaigns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var campaignSheet = ss.getSheetByName('Campaigns');

  if (!campaignSheet) {
    SpreadsheetApp.getUi().alert('No Campaigns',
      'No campaigns found. Use "New Campaign" to create one.',
      SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  ss.setActiveSheet(campaignSheet);
}

// ============================================================================
// FOLLOW-UP & GMAIL LABELS (Enhanced with auto-transitions and aging)
// ============================================================================

// Legacy compatibility - use CONFIG.CRM_LABELS instead
var FOLLOWUP_LABELS = {
  FOLLOWUP: 'CRM/Follow-Up',
  NEEDS_RESPONSE: 'CRM/Needs-Response',
  WAITING: 'CRM/Waiting-Reply'
};

/**
 * Get list of "my" email addresses for determining sent vs received
 */
function getMyEmails() {
  var emails = {};
  // Primary email
  emails[Session.getActiveUser().getEmail().toLowerCase()] = true;

  // Check settings for aliases
  var settings = getSettings();
  var aliases = settings['MY_EMAIL_ALIASES'] || '';
  if (aliases) {
    var aliasList = aliases.split(',');
    for (var i = 0; i < aliasList.length; i++) {
      var alias = aliasList[i].trim().toLowerCase();
      if (alias) emails[alias] = true;
    }
  }

  return emails;
}

/**
 * Check if an email is "from me"
 */
function isFromMe(email, myEmails) {
  if (!email) return false;
  return myEmails[email.toLowerCase()] === true;
}

/**
 * Get thread activity (inbound/outbound timestamps)
 */
function getThreadActivity(thread, myEmails) {
  var messages = thread.getMessages();
  var result = {
    lastInboundAt: null,
    lastOutboundAt: null,
    lastMessageAt: null,
    lastMessageFromMe: false,
    subject: thread.getFirstMessageSubject() || '(no subject)',
    threadId: thread.getId(),
    contactEmails: {}
  };

  for (var i = 0; i < messages.length; i++) {
    var msg = messages[i];
    var msgDate = msg.getDate();
    var fromEmail = extractEmail(msg.getFrom());
    var fromMe = isFromMe(fromEmail, myEmails);

    // Track last message
    if (!result.lastMessageAt || msgDate > result.lastMessageAt) {
      result.lastMessageAt = msgDate;
      result.lastMessageFromMe = fromMe;
    }

    if (fromMe) {
      // Outbound
      if (!result.lastOutboundAt || msgDate > result.lastOutboundAt) {
        result.lastOutboundAt = msgDate;
      }
      // Collect external recipients
      var toList = (msg.getTo() || '').split(',');
      var ccList = (msg.getCc() || '').split(',');
      var allRecipients = toList.concat(ccList);
      for (var r = 0; r < allRecipients.length; r++) {
        var recip = extractEmail(allRecipients[r]);
        if (recip && !isFromMe(recip, myEmails)) {
          result.contactEmails[recip] = true;
        }
      }
    } else {
      // Inbound
      if (!result.lastInboundAt || msgDate > result.lastInboundAt) {
        result.lastInboundAt = msgDate;
      }
      // Track sender as contact
      if (fromEmail) {
        result.contactEmails[fromEmail] = true;
      }
    }
  }

  return result;
}

/**
 * Get all threads with a label using pagination (no 100-thread limit)
 */
function getAllLabeledThreads(labelName, maxThreads) {
  var label = GmailApp.getUserLabelByName(labelName);
  if (!label) return [];

  var allThreads = [];
  var start = 0;
  var pageSize = CONFIG.PAGE_SIZE || 100;
  var cap = maxThreads || CONFIG.MAX_THREADS_PER_LABEL || 2000;

  while (allThreads.length < cap) {
    var threads = label.getThreads(start, pageSize);
    if (threads.length === 0) break;

    allThreads = allThreads.concat(threads);
    start += pageSize;

    if (threads.length < pageSize) break; // No more pages
  }

  return allThreads.slice(0, cap);
}

/**
 * Apply Mode B auto-transitions based on reply activity
 */
function applyAutoTransitions(thread, activity, myEmails) {
  var transitions = { from: null, to: null };
  var labels = thread.getLabels();
  var labelNames = {};

  for (var i = 0; i < labels.length; i++) {
    labelNames[labels[i].getName()] = labels[i];
  }

  var hasNeedsResponse = labelNames[CONFIG.CRM_LABELS.NEEDS_RESPONSE];
  var hasFollowUp = labelNames[CONFIG.CRM_LABELS.FOLLOW_UP];
  var hasWaitingReply = labelNames[CONFIG.CRM_LABELS.WAITING_REPLY];

  // Rule 1: Needs-Response + I replied → remove NR, add Waiting-Reply
  if (hasNeedsResponse && activity.lastOutboundAt && activity.lastInboundAt) {
    if (activity.lastOutboundAt > activity.lastInboundAt) {
      transitions.from = 'Needs Response';
      transitions.to = 'Waiting Reply';
      thread.removeLabel(hasNeedsResponse);
      var waitingLabel = GmailApp.getUserLabelByName(CONFIG.CRM_LABELS.WAITING_REPLY);
      if (!waitingLabel) waitingLabel = GmailApp.createLabel(CONFIG.CRM_LABELS.WAITING_REPLY);
      if (!hasWaitingReply) thread.addLabel(waitingLabel);
    }
  }

  // Rule 2: Follow-Up + I replied → convert to Waiting-Reply
  if (hasFollowUp && activity.lastOutboundAt && activity.lastInboundAt) {
    if (activity.lastOutboundAt > activity.lastInboundAt) {
      transitions.from = 'Follow-Up';
      transitions.to = 'Waiting Reply';
      thread.removeLabel(hasFollowUp);
      var waitingLabel2 = GmailApp.getUserLabelByName(CONFIG.CRM_LABELS.WAITING_REPLY);
      if (!waitingLabel2) waitingLabel2 = GmailApp.createLabel(CONFIG.CRM_LABELS.WAITING_REPLY);
      if (!hasWaitingReply) thread.addLabel(waitingLabel2);
    }
  }

  // Rule 3: Waiting-Reply + they replied → remove WR, add Needs-Response
  if (hasWaitingReply && activity.lastInboundAt && activity.lastOutboundAt) {
    if (activity.lastInboundAt > activity.lastOutboundAt) {
      transitions.from = 'Waiting Reply';
      transitions.to = 'Needs Response';
      thread.removeLabel(hasWaitingReply);
      var needsLabel = GmailApp.getUserLabelByName(CONFIG.CRM_LABELS.NEEDS_RESPONSE);
      if (!needsLabel) needsLabel = GmailApp.createLabel(CONFIG.CRM_LABELS.NEEDS_RESPONSE);
      if (!hasNeedsResponse) thread.addLabel(needsLabel);
    }
  }

  return transitions;
}

/**
 * Get current label status for a thread
 */
function getThreadLabelStatus(thread) {
  var labels = thread.getLabels();
  for (var i = 0; i < labels.length; i++) {
    var name = labels[i].getName();
    if (name === CONFIG.CRM_LABELS.NEEDS_RESPONSE) return 'Needs Response';
    if (name === CONFIG.CRM_LABELS.FOLLOW_UP) return 'Follow-Up';
    if (name === CONFIG.CRM_LABELS.WAITING_REPLY) return 'Waiting Reply';
  }
  return null;
}

/**
 * Determine winner thread for a contact with multiple labeled threads
 * Priority: Needs Response > Follow-Up > Waiting Reply
 * Tie-breaker: most recent relevant timestamp
 */
function selectWinnerThread(candidates) {
  if (candidates.length === 0) return null;
  if (candidates.length === 1) return candidates[0];

  var priorityOrder = { 'Needs Response': 1, 'Follow-Up': 2, 'Waiting Reply': 3 };

  candidates.sort(function(a, b) {
    // First by priority
    var priorityDiff = (priorityOrder[a.status] || 99) - (priorityOrder[b.status] || 99);
    if (priorityDiff !== 0) return priorityDiff;

    // Tie-breaker by relevant timestamp
    var aTime, bTime;
    if (a.status === 'Needs Response') {
      aTime = a.activity.lastInboundAt ? a.activity.lastInboundAt.getTime() : 0;
      bTime = b.activity.lastInboundAt ? b.activity.lastInboundAt.getTime() : 0;
    } else if (a.status === 'Waiting Reply') {
      aTime = a.activity.lastOutboundAt ? a.activity.lastOutboundAt.getTime() : 0;
      bTime = b.activity.lastOutboundAt ? b.activity.lastOutboundAt.getTime() : 0;
    } else {
      aTime = a.activity.lastMessageAt ? a.activity.lastMessageAt.getTime() : 0;
      bTime = b.activity.lastMessageAt ? b.activity.lastMessageAt.getTime() : 0;
    }

    return bTime - aTime; // Descending (newer first)
  });

  return candidates[0];
}

function setupFollowUpLabels() {
  var ui = SpreadsheetApp.getUi();
  var created = [];

  var labelsToCreate = [
    CONFIG.CRM_LABELS.NEEDS_RESPONSE,
    CONFIG.CRM_LABELS.FOLLOW_UP,
    CONFIG.CRM_LABELS.WAITING_REPLY
  ];

  for (var i = 0; i < labelsToCreate.length; i++) {
    var labelName = labelsToCreate[i];
    var label = GmailApp.getUserLabelByName(labelName);
    if (!label) {
      GmailApp.createLabel(labelName);
      created.push(labelName);
    }
  }

  if (created.length > 0) {
    ui.alert('Labels Created',
      'Created Gmail labels:\n' + created.join('\n') + '\n\n' +
      'You can now apply these labels to emails in Gmail.',
      ui.ButtonSet.OK);
  } else {
    ui.alert('Labels Ready',
      'All follow-up labels already exist:\n' +
      labelsToCreate.join('\n'),
      ui.ButtonSet.OK);
  }
}

function scanNeedsResponse() {
  var ui = SpreadsheetApp.getUi();
  var myEmails = getMyEmails();

  // Get or create the Needs Response label
  var label = GmailApp.getUserLabelByName(CONFIG.CRM_LABELS.NEEDS_RESPONSE);
  if (!label) {
    label = GmailApp.createLabel(CONFIG.CRM_LABELS.NEEDS_RESPONSE);
  }

  // Search for emails in inbox where last message is NOT from me
  var threads = GmailApp.search('is:inbox to:me newer_than:30d -label:"' + CONFIG.CRM_LABELS.NEEDS_RESPONSE + '"', 0, 100);

  var tagged = 0;
  var questionPatterns = [
    /\?/,
    /can you/i,
    /could you/i,
    /would you/i,
    /please.*(?:send|provide|share|confirm|let me know)/i,
    /waiting.*(?:response|reply|hear)/i,
    /get back to/i,
    /thoughts\??/i,
    /what do you think/i,
    /your input/i
  ];

  for (var i = 0; i < threads.length; i++) {
    var thread = threads[i];
    var activity = getThreadActivity(thread, myEmails);

    // Only tag if the last message is NOT from me
    if (activity.lastMessageFromMe) continue;

    // Check if message contains questions or requests
    var messages = thread.getMessages();
    var lastMsg = messages[messages.length - 1];
    var body = lastMsg.getPlainBody() || '';
    var subject = lastMsg.getSubject() || '';
    var text = subject + ' ' + body;

    var needsResponse = false;
    for (var p = 0; p < questionPatterns.length; p++) {
      if (questionPatterns[p].test(text)) {
        needsResponse = true;
        break;
      }
    }

    if (needsResponse) {
      thread.addLabel(label);
      tagged++;
    }
  }

  ui.alert('Scan Complete',
    'Emails tagged as needing response: ' + tagged + '\n\n' +
    'Check your Gmail for the "' + CONFIG.CRM_LABELS.NEEDS_RESPONSE + '" label.',
    ui.ButtonSet.OK);
}

function syncFollowUps() {
  var result = syncFollowUpsInternal();
  SpreadsheetApp.getUi().alert('Follow-up Sync Complete',
    'Threads scanned: ' + result.threadsScanned + '\n' +
    'Auto-transitions: ' + result.transitions + '\n' +
    'Contacts needing response: ' + result.needsResponse + '\n' +
    'Contacts with follow-ups: ' + result.followups + '\n' +
    'Contacts waiting reply: ' + result.waitingReply + '\n' +
    'Statuses cleared: ' + result.cleared + '\n' +
    'Contacts updated: ' + result.updated,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncFollowUpsInternal() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contactsSheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var headers = getHeaderMap(contactsSheet);

  // Ensure new columns exist
  var requiredCols = [
    'CRM_FollowUpStatus', 'CRM_FollowUpStatusSetAt', 'CRM_LastInboundAt',
    'CRM_LastOutboundAt', 'CRM_DaysInStatus', 'CRM_FollowUpSource', 'CRM_LastThreadId'
  ];
  addHeadersIfMissing(contactsSheet, requiredCols);
  headers = getHeaderMap(contactsSheet);

  if (!headers['CRM_FollowUpStatus']) {
    return { threadsScanned: 0, transitions: 0, followups: 0, needsResponse: 0, waitingReply: 0, updated: 0, cleared: 0 };
  }

  var emailIndex = buildEmailIndex(contactsSheet);
  var myEmails = getMyEmails();
  var stats = {
    threadsScanned: 0,
    transitions: 0,
    followups: 0,
    needsResponse: 0,
    waitingReply: 0,
    updated: 0,
    cleared: 0
  };

  // Collect all thread data by contact email
  var contactThreads = {}; // email -> [{ status, activity, thread }]

  // Process each label type with pagination
  var labelTypes = [
    { name: CONFIG.CRM_LABELS.NEEDS_RESPONSE, status: 'Needs Response' },
    { name: CONFIG.CRM_LABELS.FOLLOW_UP, status: 'Follow-Up' },
    { name: CONFIG.CRM_LABELS.WAITING_REPLY, status: 'Waiting Reply' }
  ];

  for (var lt = 0; lt < labelTypes.length; lt++) {
    var labelInfo = labelTypes[lt];
    var threads = getAllLabeledThreads(labelInfo.name, CONFIG.MAX_THREADS_PER_LABEL);

    for (var t = 0; t < threads.length; t++) {
      var thread = threads[t];
      stats.threadsScanned++;

      var activity = getThreadActivity(thread, myEmails);

      // Apply auto-transitions (Mode B)
      var transition = applyAutoTransitions(thread, activity, myEmails);
      if (transition.to) {
        stats.transitions++;
        logInfo('syncFollowUps', 'Transition: ' + transition.from + ' -> ' + transition.to + ' | ' + activity.subject.substring(0, 40));
      }

      // Get current status after transitions
      var currentStatus = getThreadLabelStatus(thread);
      if (!currentStatus) continue; // Label was removed by transition

      // Map thread to all contact emails
      for (var email in activity.contactEmails) {
        var normalized = email.toLowerCase().trim();
        if (!contactThreads[normalized]) {
          contactThreads[normalized] = [];
        }
        contactThreads[normalized].push({
          status: currentStatus,
          activity: activity,
          thread: thread
        });
      }
    }
  }

  // Read existing sheet data for batch update
  var lastRow = contactsSheet.getLastRow();
  var existingData = {};
  if (lastRow >= CONFIG.DATA_START_ROW) {
    var emailData = contactsSheet.getRange(CONFIG.DATA_START_ROW, CONFIG.EMAIL_COLUMN, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
    var statusCol = headers['CRM_FollowUpStatus'];
    var statusSetAtCol = headers['CRM_FollowUpStatusSetAt'];

    if (statusCol) {
      var statusData = contactsSheet.getRange(CONFIG.DATA_START_ROW, statusCol, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
      var statusSetAtData = statusSetAtCol ?
        contactsSheet.getRange(CONFIG.DATA_START_ROW, statusSetAtCol, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues() : null;

      for (var r = 0; r < emailData.length; r++) {
        var email = (emailData[r][0] || '').toString().toLowerCase().trim();
        if (email) {
          existingData[email] = {
            row: CONFIG.DATA_START_ROW + r,
            status: statusData[r][0] || '',
            statusSetAt: statusSetAtData ? statusSetAtData[r][0] : null
          };
        }
      }
    }
  }

  // Prepare batch updates
  var updates = []; // { row, col, value }
  var now = new Date();

  // Process each contact with labeled threads
  for (var contactEmail in contactThreads) {
    var candidates = contactThreads[contactEmail];
    var winner = selectWinnerThread(candidates);
    if (!winner) continue;

    var row = emailIndex[contactEmail];
    if (!row) continue; // Contact not in sheet

    var existing = existingData[contactEmail] || {};
    var statusChanged = existing.status !== winner.status;

    // Update status
    updates.push({ row: row, col: headers['CRM_FollowUpStatus'], value: winner.status });

    // Update StatusSetAt if status changed or wasn't set
    if (statusChanged || !existing.statusSetAt) {
      updates.push({ row: row, col: headers['CRM_FollowUpStatusSetAt'], value: now });
    }

    // Update timestamps
    if (headers['CRM_LastInboundAt'] && winner.activity.lastInboundAt) {
      updates.push({ row: row, col: headers['CRM_LastInboundAt'], value: winner.activity.lastInboundAt });
    }
    if (headers['CRM_LastOutboundAt'] && winner.activity.lastOutboundAt) {
      updates.push({ row: row, col: headers['CRM_LastOutboundAt'], value: winner.activity.lastOutboundAt });
    }

    // Update DaysInStatus
    if (headers['CRM_DaysInStatus']) {
      var setAt = statusChanged ? now : (existing.statusSetAt || now);
      var days = Math.floor((now.getTime() - new Date(setAt).getTime()) / 86400000);
      updates.push({ row: row, col: headers['CRM_DaysInStatus'], value: days });
    }

    // Update source info
    if (headers['CRM_FollowUpSource']) {
      var relevantDate = winner.status === 'Needs Response' ? winner.activity.lastInboundAt :
                         winner.status === 'Waiting Reply' ? winner.activity.lastOutboundAt :
                         winner.activity.lastMessageAt;
      var dateStr = relevantDate ? Utilities.formatDate(relevantDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
      var source = winner.status + ' | "' + winner.activity.subject.substring(0, 30) + '..." | ' + dateStr;
      updates.push({ row: row, col: headers['CRM_FollowUpSource'], value: source });
    }

    // Update thread ID
    if (headers['CRM_LastThreadId']) {
      updates.push({ row: row, col: headers['CRM_LastThreadId'], value: winner.activity.threadId });
    }

    // Track stats
    if (winner.status === 'Needs Response') stats.needsResponse++;
    else if (winner.status === 'Follow-Up') stats.followups++;
    else if (winner.status === 'Waiting Reply') stats.waitingReply++;

    if (statusChanged) stats.updated++;

    // Remove from existing to track what needs clearing
    delete existingData[contactEmail];
  }

  // Clear statuses for contacts no longer in any labeled threads
  for (var oldEmail in existingData) {
    var oldData = existingData[oldEmail];
    if (oldData.status) {
      updates.push({ row: oldData.row, col: headers['CRM_FollowUpStatus'], value: '' });
      if (headers['CRM_FollowUpStatusSetAt']) {
        updates.push({ row: oldData.row, col: headers['CRM_FollowUpStatusSetAt'], value: '' });
      }
      if (headers['CRM_DaysInStatus']) {
        updates.push({ row: oldData.row, col: headers['CRM_DaysInStatus'], value: '' });
      }
      if (headers['CRM_FollowUpSource']) {
        updates.push({ row: oldData.row, col: headers['CRM_FollowUpSource'], value: '' });
      }
      if (headers['CRM_LastThreadId']) {
        updates.push({ row: oldData.row, col: headers['CRM_LastThreadId'], value: '' });
      }
      stats.cleared++;
      stats.updated++;
    }
  }

  // Apply all updates
  for (var u = 0; u < updates.length; u++) {
    var upd = updates[u];
    if (upd.row && upd.col) {
      contactsSheet.getRange(upd.row, upd.col).setValue(upd.value);
    }
  }

  logInfo('syncFollowUps', 'Completed: ' + stats.threadsScanned + ' threads, ' + stats.transitions + ' transitions, ' + stats.updated + ' contacts updated');

  return stats;
}

/**
 * Setup conditional formatting and data validation for follow-up columns
 */
function setupFollowUpFormatting() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contactsSheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var headers = getHeaderMap(contactsSheet);
  var ui = SpreadsheetApp.getUi();

  if (!headers['CRM_FollowUpStatus']) {
    ui.alert('Error', 'Follow-up columns not found. Please run Setup first.', ui.ButtonSet.OK);
    return;
  }

  var lastRow = Math.max(contactsSheet.getLastRow(), 100);
  var statusCol = headers['CRM_FollowUpStatus'];
  var statusRange = contactsSheet.getRange(CONFIG.DATA_START_ROW, statusCol, lastRow, 1);

  // Add data validation dropdown
  var validStatuses = ['', 'Needs Response', 'Follow-Up', 'Waiting Reply'];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(validStatuses, true)
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(rule);

  // Clear existing conditional formatting rules for this column
  var rules = contactsSheet.getConditionalFormatRules();
  var newRules = [];
  for (var i = 0; i < rules.length; i++) {
    var ranges = rules[i].getRanges();
    var keepRule = true;
    for (var r = 0; r < ranges.length; r++) {
      if (ranges[r].getColumn() === statusCol) {
        keepRule = false;
        break;
      }
    }
    if (keepRule) newRules.push(rules[i]);
  }

  // Add new conditional formatting rules
  // Needs Response - Red/urgent
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Needs Response')
    .setBackground('#f4cccc')
    .setFontColor('#990000')
    .setBold(true)
    .setRanges([statusRange])
    .build());

  // Follow-Up - Orange/medium
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Follow-Up')
    .setBackground('#fce5cd')
    .setFontColor('#b45f06')
    .setBold(true)
    .setRanges([statusRange])
    .build());

  // Waiting Reply - Yellow/low
  newRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Waiting Reply')
    .setBackground('#fff2cc')
    .setFontColor('#7f6000')
    .setRanges([statusRange])
    .build());

  // Days in status - add gradient if column exists
  if (headers['CRM_DaysInStatus']) {
    var daysCol = headers['CRM_DaysInStatus'];
    var daysRange = contactsSheet.getRange(CONFIG.DATA_START_ROW, daysCol, lastRow, 1);

    // > 7 days warning
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(7)
      .setBackground('#ffcdd2')
      .setRanges([daysRange])
      .build());

    // 3-7 days caution
    newRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(3, 7)
      .setBackground('#ffe0b2')
      .setRanges([daysRange])
      .build());
  }

  contactsSheet.setConditionalFormatRules(newRules);

  // Freeze header row if not already
  if (contactsSheet.getFrozenRows() < 1) {
    contactsSheet.setFrozenRows(1);
  }

  // Enable filter if not already
  var filter = contactsSheet.getFilter();
  if (!filter) {
    var dataRange = contactsSheet.getDataRange();
    dataRange.createFilter();
  }

  ui.alert('Formatting Applied',
    'Added:\n' +
    '• Status dropdown validation\n' +
    '• Conditional formatting (red/orange/yellow)\n' +
    '• Days aging warning colors\n' +
    '• Header row frozen\n' +
    '• Filter enabled',
    ui.ButtonSet.OK);
}

// ============================================================================
// FUTURE: GONG SYNC (Template - uncomment and implement when ready)
// ============================================================================

/*
function syncGong() {
  var result = syncGongInternal();
  SpreadsheetApp.getUi().alert('Gong Sync Complete',
    'Calls: ' + result.calls + '\nContacts updated: ' + result.updated,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function syncGongInternal() {
  // TODO: Implement Gong API integration
  // 1. Get API token from settings
  // 2. Fetch calls from Gong API
  // 3. Match participants to contacts
  // 4. Update CRM_GongLink, CRM_LastCallDate, CRM_LastCallTitle columns
  return { calls: 0, updated: 0 };
}
*/

// ============================================================================
// HELPERS
// ============================================================================

function removeDuplicateContacts(contactsSheet, historySheet) {
  var lastRow = contactsSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  var emails = contactsSheet.getRange(CONFIG.DATA_START_ROW, CONFIG.EMAIL_COLUMN, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  var seen = {};
  var duplicateRows = [];

  // Find duplicate rows (keep first occurrence, mark later ones for deletion)
  for (var i = 0; i < emails.length; i++) {
    var email = (emails[i][0] || '').toString().toLowerCase().trim();
    if (!email) continue;

    if (seen[email]) {
      // This is a duplicate - mark for deletion
      duplicateRows.push(i + CONFIG.DATA_START_ROW);
    } else {
      seen[email] = true;
    }
  }

  // Delete from bottom up to preserve row numbers
  duplicateRows.sort(function(a, b) { return b - a; });
  for (var d = 0; d < duplicateRows.length; d++) {
    contactsSheet.deleteRow(duplicateRows[d]);
  }

  // Also remove duplicates from history sheet
  if (historySheet && duplicateRows.length > 0) {
    var historyLastRow = historySheet.getLastRow();
    if (historyLastRow >= 2) {
      var historyEmails = historySheet.getRange(2, 1, historyLastRow - 1, 1).getValues();
      var historySeen = {};
      var historyDuplicates = [];

      for (var h = 0; h < historyEmails.length; h++) {
        var hEmail = (historyEmails[h][0] || '').toString().toLowerCase().trim();
        if (!hEmail) continue;

        if (historySeen[hEmail]) {
          historyDuplicates.push(h + 2);
        } else {
          historySeen[hEmail] = true;
        }
      }

      historyDuplicates.sort(function(a, b) { return b - a; });
      for (var hd = 0; hd < historyDuplicates.length; hd++) {
        historySheet.deleteRow(historyDuplicates[hd]);
      }
    }
  }

  return duplicateRows.length;
}

function removeExcludedContacts(contactsSheet, historySheet, excludeEmails, excludeDomains) {
  if (excludeEmails.length === 0 && excludeDomains.length === 0) return 0;

  var lastRow = contactsSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return 0;

  var emails = contactsSheet.getRange(CONFIG.DATA_START_ROW, CONFIG.EMAIL_COLUMN, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  var rowsToDelete = [];
  var emailsToDelete = {};

  // Find rows to delete (check from bottom up for safe deletion)
  for (var i = emails.length - 1; i >= 0; i--) {
    var email = (emails[i][0] || '').toString().toLowerCase().trim();
    if (!email) continue;

    var domain = email.split('@')[1] || '';
    var shouldRemove = false;

    // Check email patterns
    for (var j = 0; j < excludeEmails.length; j++) {
      if (email.indexOf(excludeEmails[j]) > -1) {
        shouldRemove = true;
        break;
      }
    }

    // Check domains
    if (!shouldRemove && excludeDomains.indexOf(domain) > -1) {
      shouldRemove = true;
    }

    if (shouldRemove) {
      rowsToDelete.push(i + CONFIG.DATA_START_ROW);
      emailsToDelete[email] = true;
    }
  }

  // Delete rows from contacts sheet (already sorted bottom up)
  for (var k = 0; k < rowsToDelete.length; k++) {
    contactsSheet.deleteRow(rowsToDelete[k]);
  }

  // Batch delete from history sheet
  if (historySheet && Object.keys(emailsToDelete).length > 0) {
    var historyLastRow = historySheet.getLastRow();
    if (historyLastRow >= 2) {
      var historyEmails = historySheet.getRange(2, 1, historyLastRow - 1, 1).getValues();
      var historyRowsToDelete = [];
      for (var h = historyEmails.length - 1; h >= 0; h--) {
        var hEmail = (historyEmails[h][0] || '').toString().toLowerCase().trim();
        if (emailsToDelete[hEmail]) {
          historyRowsToDelete.push(h + 2);
        }
      }
      // Delete history rows (already sorted bottom up)
      for (var d = 0; d < historyRowsToDelete.length; d++) {
        historySheet.deleteRow(historyRowsToDelete[d]);
      }
    }
  }

  return rowsToDelete.length;
}

function parseList(str) {
  return (str || '').toLowerCase().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
}

function shouldExclude(email, subject, domain, excludeEmails, excludeSubjects, excludeDomains, excludeInternal, internalDomain, excludePromotional, thread) {
  // Check email patterns (supports full email, prefix like "no-reply@", or suffix like "@newsletter.com")
  for (var i = 0; i < excludeEmails.length; i++) {
    var pattern = excludeEmails[i];
    if (email === pattern) return true;  // Exact match
    if (email.indexOf(pattern) > -1) return true;  // Partial match (e.g., "no-reply@" matches "no-reply@company.com")
  }

  // Check subject keywords
  if (excludeSubjects.some(function(kw) { return subject.indexOf(kw) > -1; })) return true;

  // Check domains
  if (excludeDomains.indexOf(domain) > -1) return true;

  // Check if thread is internal-only (no external participants)
  if (excludeInternal && internalDomain && thread) {
    var hasExternal = threadHasExternalParticipant(thread, internalDomain);
    if (!hasExternal) return true; // Exclude if ALL participants are internal
  }

  // Check Gmail's built-in categories (Promotions, Social, Updates, Forums)
  if (excludePromotional && thread) {
    var labels = thread.getLabels();
    for (var labelIdx = 0; labelIdx < labels.length; labelIdx++) {
      var labelName = labels[labelIdx].getName().toUpperCase();
      if (labelName === 'CATEGORY_PROMOTIONS' || labelName === 'CATEGORY_SOCIAL' ||
          labelName === 'CATEGORY_UPDATES' || labelName === 'CATEGORY_FORUMS') {
        return true;
      }
    }
  }

  return false;
}

function threadHasExternalParticipant(thread, internalDomain) {
  try {
    var messages = thread.getMessages();
    for (var i = 0; i < messages.length; i++) {
      var msg = messages[i];

      // Check From
      var fromEmail = extractEmail(msg.getFrom());
      if (fromEmail && !fromEmail.endsWith('@' + internalDomain)) {
        return true;
      }

      // Check To
      var toEmails = (msg.getTo() || '').split(',');
      for (var j = 0; j < toEmails.length; j++) {
        var toEmail = extractEmail(toEmails[j]);
        if (toEmail && !toEmail.endsWith('@' + internalDomain)) {
          return true;
        }
      }

      // Check CC
      var ccEmails = (msg.getCc() || '').split(',');
      for (var k = 0; k < ccEmails.length; k++) {
        var ccEmail = extractEmail(ccEmails[k]);
        if (ccEmail && !ccEmail.endsWith('@' + internalDomain)) {
          return true;
        }
      }
    }
  } catch (e) {
    // If error, don't exclude
    return true;
  }

  return false; // All participants are internal
}

function formatEmailEntry(msg, thread, direction) {
  var date = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var subject = thread.getFirstMessageSubject() || '(no subject)';
  var preview = (msg.getPlainBody() || '').substring(0, 300).replace(/\s+/g, ' ').trim();
  return '[' + date + '] ' + direction + '\nSubject: ' + subject + '\n' + preview;
}

function buildEmailIndex(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return {};

  var emails = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.EMAIL_COLUMN, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  var index = {};

  for (var i = 0; i < emails.length; i++) {
    var email = emails[i][0];
    if (email && String(email).indexOf('@') > -1) {
      var normalized = String(email).toLowerCase().trim();
      if (!index[normalized]) index[normalized] = CONFIG.DATA_START_ROW + i;
    }
  }
  return index;
}

function buildHistoryIndex(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var emails = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var index = {};

  for (var i = 0; i < emails.length; i++) {
    var email = emails[i][0];
    if (email) {
      var normalized = String(email).toLowerCase().trim();
      if (!index[normalized]) index[normalized] = i + 2;
    }
  }
  return index;
}

function getHeaderMap(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return {};

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) map[String(headers[i])] = i + 1;
  }
  return map;
}

function findExternalEmail(messages, myEmail) {
  for (var i = messages.length - 1; i >= 0; i--) {
    var from = extractEmail(messages[i].getFrom());
    if (from && from !== myEmail) return from;
    var to = extractEmail(messages[i].getTo());
    if (to && to !== myEmail) return to;
  }
  return null;
}

/**
 * Find ALL external emails in a thread (from, to, cc)
 * Returns array of unique external emails, with the primary contact first
 */
function findAllExternalEmails(messages, myEmail) {
  var emails = {};
  var primaryEmail = null;

  // Process messages from newest to oldest
  for (var i = messages.length - 1; i >= 0; i--) {
    var msg = messages[i];

    // From
    var from = extractEmail(msg.getFrom());
    if (from && from !== myEmail) {
      if (!primaryEmail) primaryEmail = from;
      emails[from] = true;
    }

    // To (may have multiple)
    var toList = (msg.getTo() || '').split(',');
    for (var t = 0; t < toList.length; t++) {
      var toEmail = extractEmail(toList[t]);
      if (toEmail && toEmail !== myEmail) {
        if (!primaryEmail) primaryEmail = toEmail;
        emails[toEmail] = true;
      }
    }

    // CC (may have multiple)
    var ccList = (msg.getCc() || '').split(',');
    for (var c = 0; c < ccList.length; c++) {
      var ccEmail = extractEmail(ccList[c]);
      if (ccEmail && ccEmail !== myEmail) {
        emails[ccEmail] = true;
      }
    }
  }

  // Build result array with primary email first
  var result = [];
  if (primaryEmail) {
    result.push(primaryEmail);
    delete emails[primaryEmail];
  }

  // Add remaining emails
  for (var email in emails) {
    result.push(email);
  }

  return result;
}

function extractEmail(str) {
  if (!str) return '';
  var match = String(str).match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/i);
  return match ? match[0].toLowerCase().trim() : '';
}

function setCell(sheet, row, col, value) {
  if (row && col && row > 0 && col > 0 && value !== undefined) {
    sheet.getRange(row, col).setValue(value);
  }
}

// Safe getRange wrappers to prevent null errors
function safeGetValue(sheet, row, col) {
  if (!sheet || !row || !col || row < 1 || col < 1) return '';
  try {
    return sheet.getRange(row, col).getValue() || '';
  } catch (e) {
    logError('safeGetValue', e.message);
    return '';
  }
}

function safeSetValue(sheet, row, col, value) {
  if (!sheet || !row || !col || row < 1 || col < 1) return;
  try {
    sheet.getRange(row, col).setValue(value);
  } catch (e) {
    logError('safeSetValue', e.message);
  }
}

// ============================================================================
// ERROR LOGGING
// ============================================================================

function logError(functionName, message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);

    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.LOG_SHEET);
      logSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Function', 'Message', 'Type']]);
      logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f0fe');
      logSheet.setFrozenRows(1);
      logSheet.setColumnWidth(1, 150);
      logSheet.setColumnWidth(2, 150);
      logSheet.setColumnWidth(3, 400);
    }

    // Add log entry
    logSheet.insertRowAfter(1);
    logSheet.getRange(2, 1, 1, 4).setValues([[new Date(), functionName, message, 'ERROR']]);

    // Trim old logs
    var lastRow = logSheet.getLastRow();
    if (lastRow > CONFIG.MAX_LOG_ROWS + 1) {
      logSheet.deleteRows(CONFIG.MAX_LOG_ROWS + 2, lastRow - CONFIG.MAX_LOG_ROWS - 1);
    }
  } catch (e) {
    // Can't log the logging error, fail silently
  }
}

function logInfo(functionName, message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);

    if (!logSheet) return; // Don't create sheet just for info logs

    logSheet.insertRowAfter(1);
    logSheet.getRange(2, 1, 1, 4).setValues([[new Date(), functionName, message, 'INFO']]);

    // Trim old logs
    var lastRow = logSheet.getLastRow();
    if (lastRow > CONFIG.MAX_LOG_ROWS + 1) {
      logSheet.deleteRows(CONFIG.MAX_LOG_ROWS + 2, lastRow - CONFIG.MAX_LOG_ROWS - 1);
    }
  } catch (e) {
    // Fail silently
  }
}

function viewLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET);

  if (!logSheet) {
    SpreadsheetApp.getUi().alert('No Logs', 'No log sheet found. Logs are created when errors occur.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  ss.setActiveSheet(logSheet);
}

function showStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contacts = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var history = ss.getSheetByName(CONFIG.HISTORY_SHEET);

  var contactCount = contacts ? Object.keys(buildEmailIndex(contacts)).length : 0;
  var historyCount = history ? Math.max(0, history.getLastRow() - 1) : 0;

  var enabled = [];
  for (var k in INTEGRATIONS) {
    if (INTEGRATIONS[k].enabled) enabled.push(k);
  }

  var settings = getSettings();
  var internalDomain = settings['INTERNAL_DOMAIN'] || '(not set)';

  SpreadsheetApp.getUi().alert('Stats',
    'Contacts: ' + contactCount + '\n' +
    'History rows: ' + historyCount + '\n' +
    'Internal domain: ' + internalDomain + '\n\n' +
    'Enabled integrations: ' + enabled.join(', '),
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Assigns CRM_ID to all contacts that don't have one.
 * This ensures every contact has a unique ID for tracking across integrations.
 */
function assignMissingIds() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var contactsSheet = ss.getSheetByName(CONFIG.CONTACTS_SHEET);
  var headers = getHeaderMap(contactsSheet);

  if (!headers['CRM_ID']) {
    SpreadsheetApp.getUi().alert('Error', 'CRM_ID column not found. Please run Setup first.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var lastRow = contactsSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('No Contacts', 'No contacts found to assign IDs.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  var idCol = headers['CRM_ID'];
  var emailCol = CONFIG.EMAIL_COLUMN;

  // Read all IDs and emails at once for performance
  var idData = contactsSheet.getRange(CONFIG.DATA_START_ROW, idCol, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  var emailData = contactsSheet.getRange(CONFIG.DATA_START_ROW, emailCol, lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();

  var assigned = 0;
  var baseTime = Date.now();

  for (var i = 0; i < idData.length; i++) {
    var existingId = idData[i][0];
    var email = emailData[i][0];

    // Only assign ID if the row has an email but no ID
    if (email && !existingId) {
      var newId = 'C' + baseTime + (CONFIG.DATA_START_ROW + i);
      contactsSheet.getRange(CONFIG.DATA_START_ROW + i, idCol).setValue(newId);
      assigned++;
    }
  }

  SpreadsheetApp.getUi().alert('IDs Assigned',
    'Assigned CRM_ID to ' + assigned + ' contacts.\n' +
    'Total contacts: ' + emailData.length,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

