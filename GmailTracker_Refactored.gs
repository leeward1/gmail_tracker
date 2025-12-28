/**
 * ============================================================================
 * GMAIL EMAIL TRACKER FOR GOOGLE SHEETS - REFACTORED
 * ============================================================================
 *
 * A comprehensive Gmail tracking solution that monitors sent and received emails,
 * tracks responses, manages campaigns, and provides analytics.
 *
 * Architecture: TSS-style separation of concerns
 * - CONFIG: Configuration constants and script properties
 * - UTILS: Utility functions (text processing, validation, formatting)
 * - REPOSITORIES: Data access layer (Sheet operations, property storage)
 * - SERVICES: Business logic (Gmail operations, analytics calculations)
 * - CONTROLLERS: Entry points (menu handlers, triggers, user-facing functions)
 *
 * @version 2.0.0
 * @author Refactored for maintainability and security
 */

// ============================================================================
// SECTION 1: CONFIGURATION
// ============================================================================

/**
 * ============================================================================
 * IMPORTANT: PRESERVING YOUR EXISTING DATA
 * ============================================================================
 *
 * If you have existing data in columns A, B, C, etc., the tracker will NOT
 * overwrite it. The tracker columns start AFTER your existing data.
 *
 * To set where tracker columns begin:
 *   Option 1: Run setupSheet() - auto-detects last column and starts after it
 *   Option 2: Run setTrackerStartColumn() - manually pick the starting column
 *   Option 3: Edit MANUAL_START_COLUMN below (set to 0 for auto-detect)
 *
 * Example: If you have data in columns A-C and want tracker to start at D:
 *   - Set MANUAL_START_COLUMN = 4  (D is the 4th column)
 *   - Or run setTrackerStartColumn() from the menu
 * ============================================================================
 */

// SET THIS TO FORCE A SPECIFIC START COLUMN (0 = auto-detect after existing data)
const MANUAL_START_COLUMN = 0;

/**
 * Configuration namespace - all configurable values in one place.
 * Uses PropertiesService for runtime config; constants for static config.
 */
const CONFIG = {
  // Sheet names
  SHEET_NAME: 'Sheet1',
  ANALYTICS_SHEET_NAME: 'Subject Analytics',
  CAMPAIGNS_SHEET_NAME: 'Campaigns',
  CONTACTS_SHEET_NAME: 'Contact Notes',
  HEALTH_SHEET_NAME: 'Health',

  // Gmail label for tracking
  TRACKED_LABEL: 'Tracked',

  // Email matching - Column C (3) contains email addresses to match
  // Set to 0 to disable matching and append all emails as new rows
  EMAIL_MATCH_COLUMN: 3,

  // Internal domains to exclude (e.g., ['company.com', 'subsidiary.com'])
  INTERNAL_DOMAINS: [],
  EXCLUDE_INTERNAL_ONLY: true,

  // Tracking settings
  TRACK_SENT: true,
  TRACK_RECEIVED: true,

  // Gmail category exclusions
  EXCLUDE_PROMOTIONS: true,
  EXCLUDE_SOCIAL: true,
  EXCLUDE_UPDATES: true,
  EXCLUDE_FORUMS: true,

  // Performance limits
  MAX_THREADS_PER_RUN: 50,
  MAX_ROWS_TO_CHECK: 500,
  BODY_MAX_LENGTH: 50000,
  SYNC_OVERLAP_MINUTES: 5,

  // Storage settings
  STORE_FULL_BODY: false,

  // Automation settings
  REMINDER_DAYS_DEFAULT: 7,
  REMINDER_EMAIL_ENABLED: true,
  EMAIL_ON_ERROR: false,

  // Calendar settings
  CALENDAR_LOOKBACK_YEARS: 2.5,
  CALENDAR_LOOKAHEAD_DAYS: 90,
  CALENDAR_SYNC_BATCH_SIZE: 50,

  // Property keys for PropertiesService storage
  PROPERTY_KEYS: {
    COLUMN_OFFSET: 'EMAIL_TRACKER_COLUMN_OFFSET',
    START_DATE: 'EMAIL_TRACKER_START_DATE',
    SUBJECT_FILTERS: 'EMAIL_TRACKER_SUBJECT_FILTERS',
    CAMPAIGNS: 'EMAIL_TRACKER_CAMPAIGNS',
    CONTACT_NOTES: 'EMAIL_TRACKER_CONTACT_NOTES',
    REMINDER_SETTINGS: 'EMAIL_TRACKER_REMINDER_SETTINGS',
    LAST_SYNC: 'EMAIL_TRACKER_LAST_SYNC_AT',
    CALENDAR_SYNC_START: 'CALENDAR_SYNC_START_DATE'
  },

  // Expected headers (36 columns) - single source of truth
  HEADERS: [
    'Message ID',       // 0
    'Thread ID',        // 1
    'Direction',        // 2
    'From',             // 3
    'To',               // 4
    'CC',               // 5
    'BCC',              // 6
    'Date',             // 7
    'Subject',          // 8
    'Body Preview',     // 9
    'Full Body',        // 10
    'Status',           // 11
    'Thread Count',     // 12
    'Attachments',      // 13
    'Attachment Names', // 14
    'Reply 1 Date',     // 15
    'Reply 1 From',     // 16
    'Reply 1 Subject',  // 17
    'Reply 1 Body',     // 18
    'Reply 2 Date',     // 19
    'Reply 2 From',     // 20
    'Reply 2 Subject',  // 21
    'Reply 2 Body',     // 22
    'Reply 3 Date',     // 23
    'Reply 3 From',     // 24
    'Reply 3 Subject',  // 25
    'Reply 3 Body',     // 26
    'Last Meeting',     // 27
    'Next Meeting',     // 28
    'Meeting Title',    // 29
    'Meeting History',  // 30
    'Meeting Participants', // 31
    'Meeting Host',     // 32
    'Contact Notes',    // 33
    'Related Contacts', // 34
    'Thread Contacts'   // 35
  ]
};

// ============================================================================
// SECTION 2: UTILITIES
// ============================================================================

/**
 * Utility namespace - pure functions for common operations
 */
const Utils = {
  /**
   * Sanitizes text by removing potentially dangerous characters.
   * Prevents formula injection by escaping leading special characters.
   * @param {string} text - Raw text input
   * @returns {string} Sanitized text safe for spreadsheet cells
   */
  sanitizeText: function(text) {
    if (!text) return '';
    let str = String(text);

    // Remove control characters
    str = str.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');

    // Prevent formula injection - escape leading special characters
    if (/^[=+\-@\t\r]/.test(str)) {
      str = "'" + str;
    }

    return str;
  },

  /**
   * Truncates text to prevent cell overflow
   * @param {string} text - Text to truncate
   * @param {number} maxLength - Maximum character length
   * @returns {string} Truncated text with indicator if truncated
   */
  truncateText: function(text, maxLength) {
    if (!text) return '';
    const str = String(text);
    if (str.length <= maxLength) return str;
    return str.substring(0, maxLength) + '... [truncated]';
  },

  /**
   * Extracts email address from a string like "Name <email@example.com>"
   * @param {string} emailString - String potentially containing email
   * @returns {string} Extracted email address (lowercase) or empty string
   */
  extractEmailAddress: function(emailString) {
    if (!emailString) return '';
    const match = String(emailString).match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/i);
    return match ? match[0].toLowerCase() : '';
  },

  /**
   * Extracts all email addresses from a string
   * @param {string} text - Text containing email addresses
   * @returns {Array<string>} Array of email addresses
   */
  extractEmails: function(text) {
    if (!text) return [];
    const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi;
    return (String(text).match(emailRegex) || []).map(e => e.toLowerCase());
  },

  /**
   * Checks if an email address is internal based on configured domains
   * @param {string} email - Email address to check
   * @returns {boolean} True if email is from internal domain
   */
  isInternalEmail: function(email) {
    if (!email || CONFIG.INTERNAL_DOMAINS.length === 0) return false;
    const emailLower = String(email).toLowerCase();
    return CONFIG.INTERNAL_DOMAINS.some(domain =>
      emailLower.includes('@' + domain.toLowerCase())
    );
  },

  /**
   * Converts column number to letter (e.g., 1 -> A, 27 -> AA)
   * @param {number} column - Column number (1-indexed)
   * @returns {string} Column letter(s)
   */
  columnToLetter: function(column) {
    let letter = '';
    while (column > 0) {
      const mod = (column - 1) % 26;
      letter = String.fromCharCode(65 + mod) + letter;
      column = Math.floor((column - mod) / 26);
    }
    return letter;
  },

  /**
   * Normalizes a subject line for grouping (removes Re:/Fwd: prefixes)
   * @param {string} subject - Email subject
   * @returns {string} Normalized subject
   */
  normalizeSubject: function(subject) {
    if (!subject) return '(No Subject)';
    let normalized = String(subject);
    // Remove Re:/Fwd:/etc. prefixes (run twice for nested prefixes)
    normalized = normalized.replace(/^(re:|fwd:|fw:|aw:)\s*/gi, '');
    normalized = normalized.replace(/^(re:|fwd:|fw:|aw:)\s*/gi, '');
    normalized = normalized.trim();
    return normalized || '(No Subject)';
  },

  /**
   * Formats a date for display
   * @param {Date} date - Date to format
   * @param {string} format - Format string (default: 'yyyy-MM-dd HH:mm:ss')
   * @returns {string} Formatted date string
   */
  formatDate: function(date, format) {
    if (!date) return '';
    format = format || 'yyyy-MM-dd HH:mm:ss';
    try {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
    } catch (e) {
      return String(date);
    }
  },

  /**
   * Validates an ISO date string (YYYY-MM-DD)
   * @param {string} dateStr - Date string to validate
   * @returns {boolean} True if valid
   */
  isValidISODate: function(dateStr) {
    return /^\d{4}-\d{2}-\d{2}$/.test(dateStr);
  },

  /**
   * Checks if a reply indicates a bounce
   * @param {string} replyFrom - Reply sender
   * @param {string} replySubject - Reply subject
   * @param {string} replyBody - Reply body
   * @returns {boolean} True if likely a bounce
   */
  isBounced: function(replyFrom, replySubject, replyBody) {
    const bounceFromPatterns = [
      'mailer-daemon', 'postmaster', 'mail delivery', 'maildelivery',
      'no-reply', 'noreply', 'bounce'
    ];
    const fromLower = (replyFrom || '').toLowerCase();
    for (const pattern of bounceFromPatterns) {
      if (fromLower.includes(pattern)) return true;
    }

    const bounceSubjectPatterns = [
      'undeliverable', 'delivery failed', 'delivery status', 'mail delivery failed',
      'returned mail', 'failure notice', 'delivery notification',
      'could not be delivered', 'message not delivered', 'rejected'
    ];
    const subjectLower = (replySubject || '').toLowerCase();
    for (const pattern of bounceSubjectPatterns) {
      if (subjectLower.includes(pattern)) return true;
    }

    const bodyLower = (replyBody || '').toLowerCase().substring(0, 500);
    const bounceBodyPatterns = [
      'message was undeliverable', 'address rejected', 'user unknown',
      'mailbox not found', 'recipient rejected', 'does not exist',
      '550 ', '553 ', '554 '
    ];
    for (const pattern of bounceBodyPatterns) {
      if (bodyLower.includes(pattern)) return true;
    }

    return false;
  }
};

// ============================================================================
// SECTION 3: REPOSITORIES (Data Access Layer)
// ============================================================================

/**
 * PropertyRepository - manages PropertiesService storage
 */
const PropertyRepository = {
  /**
   * Gets a property value
   * @param {string} key - Property key from CONFIG.PROPERTY_KEYS
   * @returns {string|null} Property value or null
   */
  get: function(key) {
    return PropertiesService.getScriptProperties().getProperty(key);
  },

  /**
   * Sets a property value
   * @param {string} key - Property key
   * @param {string} value - Value to store
   */
  set: function(key, value) {
    PropertiesService.getScriptProperties().setProperty(key, value);
  },

  /**
   * Deletes a property
   * @param {string} key - Property key to delete
   */
  delete: function(key) {
    PropertiesService.getScriptProperties().deleteProperty(key);
  },

  /**
   * Gets a JSON-parsed property value
   * @param {string} key - Property key
   * @param {*} defaultValue - Default value if property doesn't exist or parsing fails
   * @returns {*} Parsed value or default
   */
  getJSON: function(key, defaultValue) {
    const value = this.get(key);
    if (!value) return defaultValue;
    try {
      return JSON.parse(value);
    } catch (e) {
      return defaultValue;
    }
  },

  /**
   * Stores a value as JSON
   * @param {string} key - Property key
   * @param {*} value - Value to stringify and store
   */
  setJSON: function(key, value) {
    this.set(key, JSON.stringify(value));
  },

  // Specific getters/setters for common properties
  getColumnOffset: function() {
    const offset = this.get(CONFIG.PROPERTY_KEYS.COLUMN_OFFSET);
    return offset ? parseInt(offset, 10) : 1;
  },

  setColumnOffset: function(offset) {
    this.set(CONFIG.PROPERTY_KEYS.COLUMN_OFFSET, String(offset));
  },

  getLastSyncAt: function() {
    const timestamp = this.get(CONFIG.PROPERTY_KEYS.LAST_SYNC);
    if (timestamp) {
      try {
        return new Date(timestamp);
      } catch (e) {
        return null;
      }
    }
    return null;
  },

  setLastSyncAt: function(date) {
    this.set(CONFIG.PROPERTY_KEYS.LAST_SYNC, date.toISOString());
  },

  getStartDate: function() {
    const dateStr = this.get(CONFIG.PROPERTY_KEYS.START_DATE);
    if (dateStr) {
      try {
        return new Date(dateStr);
      } catch (e) {
        return null;
      }
    }
    return null;
  },

  setStartDate: function(dateStr) {
    this.set(CONFIG.PROPERTY_KEYS.START_DATE, dateStr);
  },

  getSubjectFilters: function() {
    return this.getJSON(CONFIG.PROPERTY_KEYS.SUBJECT_FILTERS, []);
  },

  setSubjectFilters: function(filters) {
    this.setJSON(CONFIG.PROPERTY_KEYS.SUBJECT_FILTERS, filters);
  },

  getCampaigns: function() {
    return this.getJSON(CONFIG.PROPERTY_KEYS.CAMPAIGNS, {});
  },

  setCampaigns: function(campaigns) {
    this.setJSON(CONFIG.PROPERTY_KEYS.CAMPAIGNS, campaigns);
  },

  getContactNotes: function() {
    return this.getJSON(CONFIG.PROPERTY_KEYS.CONTACT_NOTES, {});
  },

  setContactNotes: function(notes) {
    this.setJSON(CONFIG.PROPERTY_KEYS.CONTACT_NOTES, notes);
  },

  getReminderSettings: function() {
    return this.getJSON(CONFIG.PROPERTY_KEYS.REMINDER_SETTINGS, {
      enabled: true,
      defaultDays: CONFIG.REMINDER_DAYS_DEFAULT,
      sendEmail: CONFIG.REMINDER_EMAIL_ENABLED,
      lastRun: null,
      excludedEmails: [],
      sentReminders: {}
    });
  },

  setReminderSettings: function(settings) {
    this.setJSON(CONFIG.PROPERTY_KEYS.REMINDER_SETTINGS, settings);
  }
};

/**
 * SheetRepository - manages spreadsheet data access
 */
const SheetRepository = {
  /**
   * Gets the active spreadsheet
   * @returns {Spreadsheet} Active spreadsheet
   */
  getSpreadsheet: function() {
    return SpreadsheetApp.getActiveSpreadsheet();
  },

  /**
   * Gets a sheet by name, creating if needed
   * @param {string} sheetName - Name of the sheet
   * @param {boolean} create - Whether to create if not exists
   * @returns {Sheet|null} The sheet or null
   */
  getSheet: function(sheetName, create) {
    const ss = this.getSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet && create) {
      sheet = ss.insertSheet(sheetName);
    }
    return sheet;
  },

  /**
   * Gets the main tracker sheet
   * @returns {Sheet|null} The tracker sheet
   */
  getTrackerSheet: function() {
    return this.getSheet(CONFIG.SHEET_NAME, false);
  },

  /**
   * Gets existing headers from the tracker sheet
   * @returns {Object} { headers: string[], startColumn: number }
   */
  getExistingHeaders: function() {
    const sheet = this.getTrackerSheet();
    if (!sheet) return { headers: [], startColumn: 1 };

    const colOffset = PropertyRepository.getColumnOffset();
    const lastCol = sheet.getLastColumn();

    if (lastCol < colOffset) return { headers: [], startColumn: colOffset };

    const numCols = Math.min(lastCol - colOffset + 1, CONFIG.HEADERS.length);
    const headers = sheet.getRange(1, colOffset, 1, numCols).getValues()[0];

    return { headers: headers, startColumn: colOffset };
  },

  /**
   * Detects which headers are missing and their positions
   * @returns {Object[]} Array of { name, position } for missing headers
   */
  detectMissingHeaders: function() {
    const { headers } = this.getExistingHeaders();
    const missing = [];

    for (let i = 0; i < CONFIG.HEADERS.length; i++) {
      if (headers.indexOf(CONFIG.HEADERS[i]) === -1) {
        missing.push({ name: CONFIG.HEADERS[i], position: i });
      }
    }

    return missing;
  },

  /**
   * Creates a header -> column index map for existing headers
   * @returns {Object} Map of header name to column index (0-based relative to offset)
   */
  getHeaderIndexMap: function() {
    const { headers } = this.getExistingHeaders();
    const map = {};

    for (let i = 0; i < headers.length; i++) {
      if (headers[i]) {
        map[headers[i]] = i;
      }
    }

    return map;
  },

  /**
   * Ensures all required headers exist (idempotent operation)
   * Only adds missing headers, never duplicates
   * @returns {Object} { added: string[], existing: string[] }
   */
  ensureHeaders: function() {
    const sheet = this.getTrackerSheet();
    if (!sheet) {
      throw new Error('Tracker sheet not found. Please create ' + CONFIG.SHEET_NAME + ' first.');
    }

    const colOffset = PropertyRepository.getColumnOffset();
    const { headers: existingHeaders } = this.getExistingHeaders();
    const existingSet = new Set(existingHeaders.filter(h => h));

    const added = [];
    const existing = [];

    // Check each expected header
    for (let i = 0; i < CONFIG.HEADERS.length; i++) {
      const expectedHeader = CONFIG.HEADERS[i];
      if (existingSet.has(expectedHeader)) {
        existing.push(expectedHeader);
      } else {
        added.push(expectedHeader);
      }
    }

    // If nothing is missing, just verify order and return
    if (added.length === 0) {
      // Verify headers are in correct order and fix if needed
      const currentHeaders = sheet.getRange(1, colOffset, 1, CONFIG.HEADERS.length).getValues()[0];
      let needsUpdate = false;
      for (let i = 0; i < CONFIG.HEADERS.length; i++) {
        if (currentHeaders[i] !== CONFIG.HEADERS[i]) {
          needsUpdate = true;
          break;
        }
      }
      if (needsUpdate) {
        sheet.getRange(1, colOffset, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
      }
      return { added: [], existing: existing };
    }

    // Insert missing columns at correct positions (from last to first to preserve positions)
    const missing = this.detectMissingHeaders();
    missing.sort((a, b) => b.position - a.position);

    for (const col of missing) {
      const insertAt = colOffset + col.position;
      sheet.insertColumnBefore(insertAt);
      sheet.getRange(1, insertAt).setValue(col.name).setFontWeight('bold');
    }

    return { added: added, existing: existing };
  },

  /**
   * Gets existing message IDs from the sheet
   * @returns {Set<string>} Set of existing message IDs
   */
  getExistingMessageIds: function() {
    const sheet = this.getTrackerSheet();
    if (!sheet) return new Set();

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return new Set();

    const colOffset = PropertyRepository.getColumnOffset();
    const messageIdColumn = sheet.getRange(2, colOffset, lastRow - 1, 1).getValues();
    const messageIds = new Set();

    for (let i = 0; i < messageIdColumn.length; i++) {
      if (messageIdColumn[i][0]) {
        messageIds.add(String(messageIdColumn[i][0]));
      }
    }

    return messageIds;
  },

  /**
   * Gets email to row mapping for contact matching
   * @returns {Object} Map of email -> row number (1-indexed)
   */
  getEmailToRowMap: function() {
    const sheet = this.getTrackerSheet();
    if (!sheet || CONFIG.EMAIL_MATCH_COLUMN <= 0) return {};

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};

    const emailColumn = sheet.getRange(2, CONFIG.EMAIL_MATCH_COLUMN, lastRow - 1, 1).getValues();
    const map = {};

    for (let r = 0; r < emailColumn.length; r++) {
      const email = Utils.extractEmailAddress(String(emailColumn[r][0] || '').toLowerCase().trim());
      if (email) {
        map[email] = r + 2; // Row number (1-indexed, skip header)
      }
    }

    return map;
  }
};

/**
 * HealthRepository - manages health/diagnostic logging
 */
const HealthRepository = {
  /**
   * Gets or creates the health sheet
   * @returns {Sheet} Health sheet
   */
  getHealthSheet: function() {
    const ss = SheetRepository.getSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.HEALTH_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.HEALTH_SHEET_NAME);
      const headers = [
        'Start Time', 'End Time', 'Duration (sec)', 'Query Used',
        'Threads Scanned', 'Messages Added', 'Threads Labeled', 'Errors', 'Status'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(4, 400);
    }

    return sheet;
  },

  /**
   * Logs a sync run to the health sheet
   * @param {Object} runData - Run metrics
   */
  logEntry: function(runData) {
    try {
      const sheet = this.getHealthSheet();
      const duration = runData.endTime && runData.startTime
        ? Math.round((runData.endTime - runData.startTime) / 1000)
        : 0;

      const row = [
        runData.startTime || new Date(),
        runData.endTime || new Date(),
        duration,
        runData.query || '',
        runData.threadsScanned || 0,
        runData.messagesAdded || 0,
        runData.threadsLabeled || 0,
        runData.errors || '',
        runData.status || 'Unknown'
      ];

      sheet.appendRow(row);

      // Keep only last 500 entries
      const lastRow = sheet.getLastRow();
      if (lastRow > 501) {
        sheet.deleteRows(2, lastRow - 501);
      }
    } catch (e) {
      Logger.log('Error logging health entry: ' + e);
    }
  }
};

// ============================================================================
// SECTION 4: SERVICES (Business Logic)
// ============================================================================

/**
 * RunSummary - tracks and reports sync run metrics
 */
const RunSummary = {
  _data: null,

  /**
   * Initializes a new run summary
   */
  init: function() {
    this._data = {
      startTime: new Date(),
      endTime: null,
      spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
      spreadsheetUrl: SpreadsheetApp.getActiveSpreadsheet().getUrl(),
      sheetNames: [],
      query: '',
      threadsScanned: 0,
      messagesAdded: 0,
      rowsUpdated: 0,
      rowsCreated: 0,
      threadsLabeled: 0,
      repliesDetected: 0,
      errors: [],
      gmailActionsRead: 0,
      gmailActionsSent: 0,
      status: 'Started'
    };
    return this;
  },

  /**
   * Sets a metric value
   * @param {string} key - Metric key
   * @param {*} value - Metric value
   */
  set: function(key, value) {
    if (this._data) this._data[key] = value;
  },

  /**
   * Increments a numeric metric
   * @param {string} key - Metric key
   * @param {number} amount - Amount to add (default 1)
   */
  increment: function(key, amount) {
    if (this._data && typeof this._data[key] === 'number') {
      this._data[key] += (amount || 1);
    }
  },

  /**
   * Adds an error to the list
   * @param {string} error - Error message
   */
  addError: function(error) {
    if (this._data && this._data.errors.length < 20) {
      this._data.errors.push(error);
    }
  },

  /**
   * Adds a sheet name to the list
   * @param {string} name - Sheet name
   */
  addSheet: function(name) {
    if (this._data && !this._data.sheetNames.includes(name)) {
      this._data.sheetNames.push(name);
    }
  },

  /**
   * Completes the run and returns summary data
   * @param {string} status - Final status
   * @returns {Object} Complete run data
   */
  complete: function(status) {
    if (!this._data) return null;

    this._data.endTime = new Date();
    this._data.status = status;
    return this._data;
  },

  /**
   * Prints a formatted summary to the log
   */
  printSummary: function() {
    if (!this._data) {
      Logger.log('No run data available');
      return;
    }

    const d = this._data;
    const duration = d.endTime && d.startTime
      ? Math.round((d.endTime - d.startTime) / 1000)
      : 0;

    const summary = [
      '========================================',
      'GMAIL TRACKER RUN SUMMARY',
      '========================================',
      'Run Timestamp: ' + Utils.formatDate(d.startTime),
      'Duration: ' + duration + ' seconds',
      'Status: ' + d.status,
      '',
      'SPREADSHEET INFO:',
      '  ID: ' + d.spreadsheetId,
      '  URL: ' + d.spreadsheetUrl,
      '  Sheets: ' + d.sheetNames.join(', '),
      '',
      'PROCESSING METRICS:',
      '  Threads Scanned: ' + d.threadsScanned,
      '  Rows Updated: ' + d.rowsUpdated,
      '  Rows Created: ' + d.rowsCreated,
      '  Total Messages Added: ' + d.messagesAdded,
      '  Replies Detected: ' + d.repliesDetected,
      '  Threads Labeled: ' + d.threadsLabeled,
      '',
      'GMAIL ACTIONS:',
      '  Threads/Messages Read: ' + d.gmailActionsRead,
      '  Emails Sent: ' + d.gmailActionsSent,
      '',
      'QUERY USED:',
      '  ' + (d.query || 'N/A'),
      ''
    ];

    if (d.errors.length > 0) {
      summary.push('ERRORS (' + d.errors.length + '):');
      d.errors.slice(0, 10).forEach(function(err, i) {
        summary.push('  ' + (i + 1) + '. ' + err);
      });
      if (d.errors.length > 10) {
        summary.push('  ... and ' + (d.errors.length - 10) + ' more');
      }
    } else {
      summary.push('ERRORS: None');
    }

    summary.push('========================================');

    Logger.log(summary.join('\n'));
  },

  /**
   * Gets the current run data
   * @returns {Object|null} Current run data
   */
  getData: function() {
    return this._data;
  }
};

/**
 * GmailService - handles Gmail operations
 */
const GmailService = {
  /**
   * Builds a Gmail search query based on configuration
   * @returns {string} Gmail search query
   */
  buildSearchQuery: function() {
    const parts = [];

    if (CONFIG.TRACK_SENT) parts.push('in:sent');
    if (CONFIG.TRACK_RECEIVED) parts.push('in:inbox');

    if (parts.length === 0) {
      throw new Error('Both TRACK_SENT and TRACK_RECEIVED are disabled.');
    }

    // Date filter
    let dateFilter = 'newer_than:30d';
    const lastSyncAt = PropertyRepository.getLastSyncAt();
    const startDate = PropertyRepository.getStartDate();

    if (lastSyncAt) {
      const overlapTime = new Date(lastSyncAt.getTime() - (CONFIG.SYNC_OVERLAP_MINUTES * 60 * 1000));
      const dateStr = Utils.formatDate(overlapTime, 'yyyy/MM/dd');
      dateFilter = 'after:' + dateStr;
    } else if (startDate) {
      const startDateStr = Utils.formatDate(startDate, 'yyyy/MM/dd');
      dateFilter = 'after:' + startDateStr;
    }

    // Category exclusions
    let categoryExclusions = '';
    if (CONFIG.EXCLUDE_PROMOTIONS) categoryExclusions += ' -category:promotions';
    if (CONFIG.EXCLUDE_SOCIAL) categoryExclusions += ' -category:social';
    if (CONFIG.EXCLUDE_UPDATES) categoryExclusions += ' -category:updates';
    if (CONFIG.EXCLUDE_FORUMS) categoryExclusions += ' -category:forums';

    return '(' + parts.join(' OR ') + ') ' + dateFilter +
           ' -label:' + CONFIG.TRACKED_LABEL + categoryExclusions;
  },

  /**
   * Searches Gmail with pagination support
   * @param {string} query - Search query
   * @param {number} start - Start index
   * @param {number} max - Maximum results
   * @returns {GmailThread[]} Array of threads
   */
  search: function(query, start, max) {
    start = start || 0;
    max = max || CONFIG.MAX_THREADS_PER_RUN;
    return GmailApp.search(query, start, max);
  },

  /**
   * Gets or creates the tracking label
   * @returns {GmailLabel} The tracking label
   */
  getOrCreateLabel: function() {
    let label = GmailApp.getUserLabelByName(CONFIG.TRACKED_LABEL);
    if (!label) {
      label = GmailApp.createLabel(CONFIG.TRACKED_LABEL);
    }
    return label;
  },

  /**
   * Checks if current user sent a message
   * @param {GmailMessage} message - The message
   * @param {string} userEmail - User's email
   * @returns {boolean} True if sent by user
   */
  isMessageFromMe: function(message, userEmail) {
    const fromEmail = message.getFrom() || '';
    return fromEmail.toLowerCase().indexOf(userEmail.toLowerCase()) !== -1;
  },

  /**
   * Determines if a message should be excluded based on internal-only rules
   * @param {GmailMessage} message - Message to check
   * @param {string} currentUserEmail - Current user's email (lowercase)
   * @returns {boolean} True if should be excluded
   */
  shouldExcludeMessage: function(message, currentUserEmail) {
    if (!CONFIG.EXCLUDE_INTERNAL_ONLY || CONFIG.INTERNAL_DOMAINS.length === 0) {
      return false;
    }

    try {
      const from = message.getFrom() || '';
      const to = message.getTo() || '';
      const cc = message.getCc() || '';
      const bcc = message.getBcc() || '';
      const emails = Utils.extractEmails([from, to, cc, bcc].join(','));

      if (emails.length === 0) return false;

      // Filter out current user
      const otherParticipants = emails.filter(function(email) {
        return email.toLowerCase() !== currentUserEmail;
      });

      if (otherParticipants.length === 0) return false;

      // Exclude only if ALL others are internal
      return otherParticipants.every(function(email) {
        return Utils.isInternalEmail(email);
      });
    } catch (e) {
      Logger.log('Error in shouldExcludeMessage: ' + e);
      return false;
    }
  },

  /**
   * Gets all unique participants across all messages in a thread
   * @param {GmailThread} thread - The Gmail thread
   * @param {string} userEmailLower - Current user's email (lowercase)
   * @returns {string[]} Array of unique external email addresses
   */
  getThreadParticipants: function(thread, userEmailLower) {
    const participants = new Set();

    try {
      const messages = thread.getMessages();
      for (const message of messages) {
        const allEmails = Utils.extractEmails(
          [message.getFrom(), message.getTo(), message.getCc()].join(' ')
        );
        for (const email of allEmails) {
          const emailLower = email.toLowerCase();
          if (emailLower !== userEmailLower) {
            participants.add(emailLower);
          }
        }
      }
    } catch (e) {
      Logger.log('Error getting thread participants: ' + e);
    }

    return Array.from(participants);
  },

  /**
   * Sends an error notification email (rate-limited)
   * @param {string} subject - Email subject
   * @param {string} body - Email body
   */
  sendErrorNotification: function(subject, body) {
    if (!CONFIG.EMAIL_ON_ERROR) return;

    try {
      const userEmail = Session.getActiveUser().getEmail();
      if (userEmail) {
        GmailApp.sendEmail(userEmail, '[Gmail Tracker] ' + subject, body);
        RunSummary.increment('gmailActionsSent');
      }
    } catch (e) {
      Logger.log('Error sending notification email: ' + e);
    }
  }
};

/**
 * TrackerService - core email tracking logic
 */
const TrackerService = {
  /**
   * Main email tracking function
   * @returns {Object} Run summary data
   */
  trackEmails: function() {
    const lock = LockService.getScriptLock();
    RunSummary.init();
    RunSummary.addSheet(CONFIG.SHEET_NAME);

    if (!lock.tryLock(10000)) {
      Logger.log('Could not acquire lock. Another instance may be running.');
      const data = RunSummary.complete('Skipped (Lock)');
      HealthRepository.logEntry(data);
      RunSummary.printSummary();
      return data;
    }

    try {
      const sheet = SheetRepository.getTrackerSheet();
      if (!sheet) {
        RunSummary.addError('Sheet not found');
        const data = RunSummary.complete('Error');
        HealthRepository.logEntry(data);
        GmailService.sendErrorNotification('Sync Error', 'Tracker sheet not found.');
        RunSummary.printSummary();
        return data;
      }

      const userEmail = Session.getActiveUser().getEmail() || '';
      const userEmailLower = userEmail.toLowerCase();
      if (!userEmail) {
        RunSummary.addError('Could not get user email');
        const data = RunSummary.complete('Error');
        HealthRepository.logEntry(data);
        RunSummary.printSummary();
        return data;
      }

      // Build and execute search
      const query = GmailService.buildSearchQuery();
      RunSummary.set('query', query);
      Logger.log('Search query: ' + query);

      const threads = GmailService.search(query);
      RunSummary.set('threadsScanned', threads.length);
      RunSummary.set('gmailActionsRead', threads.length);

      if (threads.length === 0) {
        PropertyRepository.setLastSyncAt(new Date());
        const data = RunSummary.complete('Success (No new emails)');
        HealthRepository.logEntry(data);
        RunSummary.printSummary();
        return data;
      }

      const trackedLabel = GmailService.getOrCreateLabel();
      const existingMessageIds = SheetRepository.getExistingMessageIds();
      const emailToRowMap = SheetRepository.getEmailToRowMap();
      const colOffset = PropertyRepository.getColumnOffset();

      // Track updates
      const updatedRows = new Set();
      const rowUpdates = [];
      const newRows = [];
      const newContactsMap = {};
      const threadsWithNewMessages = [];

      for (let t = 0; t < threads.length; t++) {
        const thread = threads[t];
        let threadHasNewMessages = false;

        try {
          const messages = thread.getMessages();
          const threadId = thread.getId();
          const threadCount = thread.getMessageCount();
          let threadParticipantsCache = null;

          RunSummary.increment('gmailActionsRead', messages.length);

          for (let m = 0; m < messages.length; m++) {
            const message = messages[m];
            const messageId = message.getId();

            if (existingMessageIds.has(messageId)) continue;

            const isSentByMe = GmailService.isMessageFromMe(message, userEmail);
            if (isSentByMe && !CONFIG.TRACK_SENT) continue;
            if (!isSentByMe && !CONFIG.TRACK_RECEIVED) continue;
            if (GmailService.shouldExcludeMessage(message, userEmailLower)) continue;

            try {
              // Extract message data with sanitization
              const from = Utils.sanitizeText(message.getFrom() || '');
              const to = Utils.sanitizeText(message.getTo() || '');
              const cc = Utils.sanitizeText(message.getCc() || '');
              const bcc = Utils.sanitizeText(message.getBcc() || '');
              const subject = Utils.sanitizeText(message.getSubject() || '(No Subject)');
              const body = Utils.sanitizeText(message.getPlainBody() || '');
              const bodyPreview = Utils.truncateText(body, 100);
              const fullBody = CONFIG.STORE_FULL_BODY ? Utils.truncateText(body, CONFIG.BODY_MAX_LENGTH) : '';
              const emailDate = message.getDate();
              const direction = isSentByMe ? 'Sent' : 'Received';

              // Attachments
              const attachments = message.getAttachments();
              const attachmentCount = attachments.length;
              const attachmentNames = attachments.slice(0, 5).map(function(a) { return a.getName(); }).join('; ');

              // Extract participants
              const fromEmails = Utils.extractEmails(from);
              const toEmails = Utils.extractEmails(to);
              const ccEmails = Utils.extractEmails(cc);
              const allParticipants = [...new Set([...fromEmails, ...toEmails, ...ccEmails])];
              const externalParticipants = allParticipants.filter(function(e) { return e !== userEmailLower; });

              // Determine direct contacts
              let directContacts = [];
              if (direction === 'Sent') {
                directContacts = toEmails.filter(function(e) { return e !== userEmailLower; });
              } else {
                directContacts = fromEmails.filter(function(e) { return e !== userEmailLower; });
              }

              // Cache thread participants
              if (threadParticipantsCache === null) {
                threadParticipantsCache = GmailService.getThreadParticipants(thread, userEmailLower);
              }

              // Process each direct contact
              for (const contactEmail of directContacts) {
                if (!contactEmail) continue;

                const relatedContacts = externalParticipants
                  .filter(function(e) { return e !== contactEmail; })
                  .join(', ');
                const threadContacts = threadParticipantsCache
                  .filter(function(e) { return e !== contactEmail; })
                  .join(', ');

                // Build row data (36 columns)
                const rowData = [
                  messageId, threadId, direction,
                  Utils.truncateText(from, 500),
                  Utils.truncateText(to, 5000),
                  Utils.truncateText(cc, 5000),
                  Utils.truncateText(bcc, 5000),
                  emailDate,
                  Utils.truncateText(subject, 1000),
                  bodyPreview, fullBody,
                  'Awaiting Reply',
                  threadCount,
                  attachmentCount,
                  Utils.truncateText(attachmentNames, 500),
                  '', '', '', '',  // Reply 1
                  '', '', '', '',  // Reply 2
                  '', '', '', '',  // Reply 3
                  '', '', '', '', '', '', // Calendar (6 columns)
                  '',              // Contact Notes
                  relatedContacts,
                  threadContacts
                ];

                // One row per contact email
                if (CONFIG.EMAIL_MATCH_COLUMN > 0 && emailToRowMap[contactEmail]) {
                  const matchedRow = emailToRowMap[contactEmail];
                  if (!updatedRows.has(matchedRow)) {
                    rowUpdates.push({ rowNum: matchedRow, data: rowData });
                    updatedRows.add(matchedRow);
                  } else {
                    // Update with newer email
                    for (let i = 0; i < rowUpdates.length; i++) {
                      if (rowUpdates[i].rowNum === matchedRow) {
                        rowUpdates[i].data = rowData;
                        break;
                      }
                    }
                  }
                } else if (CONFIG.EMAIL_MATCH_COLUMN > 0 && newContactsMap[contactEmail] !== undefined) {
                  newRows[newContactsMap[contactEmail]].data = rowData;
                } else {
                  newRows.push({ data: rowData, contactEmail: contactEmail });
                  newContactsMap[contactEmail] = newRows.length - 1;
                }
              }

              existingMessageIds.add(messageId);
              threadHasNewMessages = true;

            } catch (msgError) {
              Logger.log('Error processing message: ' + msgError);
              RunSummary.addError('Message: ' + (msgError.message || msgError));
            }
          }

          if (threadHasNewMessages) {
            threadsWithNewMessages.push(thread);
          }

        } catch (e) {
          Logger.log('Error processing thread: ' + e);
          RunSummary.addError('Thread: ' + e.message);
        }
      }

      // Write updates
      let rowsWritten = false;

      if (rowUpdates.length > 0) {
        for (const update of rowUpdates) {
          sheet.getRange(update.rowNum, colOffset, 1, update.data.length).setValues([update.data]);
        }
        RunSummary.set('rowsUpdated', rowUpdates.length);
        rowsWritten = true;
      }

      if (newRows.length > 0) {
        const lastRow = sheet.getLastRow();
        const startRow = lastRow + 1;
        const newRowData = newRows.map(function(r) { return r.data; });
        sheet.getRange(startRow, colOffset, newRowData.length, newRowData[0].length).setValues(newRowData);

        if (CONFIG.EMAIL_MATCH_COLUMN > 0) {
          const contactEmails = newRows.map(function(r) { return [r.contactEmail]; });
          sheet.getRange(startRow, CONFIG.EMAIL_MATCH_COLUMN, contactEmails.length, 1).setValues(contactEmails);
        }

        RunSummary.set('rowsCreated', newRows.length);
        rowsWritten = true;
      }

      RunSummary.set('messagesAdded', rowUpdates.length + newRows.length);

      // Label threads
      let labeledCount = 0;
      for (const thread of threadsWithNewMessages) {
        try {
          thread.addLabel(trackedLabel);
          labeledCount++;
        } catch (e) {
          RunSummary.addError('Labeling: ' + e.message);
        }
      }
      RunSummary.set('threadsLabeled', labeledCount);

      // Update sync cursor
      PropertyRepository.setLastSyncAt(new Date());

      const data = RunSummary.complete(rowsWritten ? 'Success' : 'Success (No new messages)');
      HealthRepository.logEntry(data);
      RunSummary.printSummary();
      return data;

    } catch (e) {
      RunSummary.addError(e.message || String(e));
      const data = RunSummary.complete('Error');
      HealthRepository.logEntry(data);
      GmailService.sendErrorNotification('Sync Failed', e.message);
      RunSummary.printSummary();
      return data;

    } finally {
      lock.releaseLock();
    }
  },

  /**
   * Checks for replies to tracked emails
   * @returns {Object} Run summary data
   */
  checkForReplies: function() {
    const lock = LockService.getScriptLock();

    if (!lock.tryLock(10000)) {
      Logger.log('Could not acquire lock for reply check.');
      return null;
    }

    try {
      const sheet = SheetRepository.getTrackerSheet();
      if (!sheet) return null;

      const colOffset = PropertyRepository.getColumnOffset();
      const userEmail = Session.getActiveUser().getEmail() || '';
      const userEmailLower = userEmail.toLowerCase();
      const lastRow = sheet.getLastRow();
      const rowsToCheck = Math.min(lastRow, CONFIG.MAX_ROWS_TO_CHECK + 1);

      // Read tracker data (36 columns)
      const data = sheet.getRange(1, colOffset, rowsToCheck, 36).getValues();
      const rowsNeedingCheck = [];

      for (let i = 1; i < data.length; i++) {
        const status = data[i][11];
        if (status === 'Awaiting Reply' && data[i][1]) {
          rowsNeedingCheck.push(i);
        }
      }

      if (rowsNeedingCheck.length === 0) {
        Logger.log('No rows need reply checking.');
        return null;
      }

      const rowUpdates = [];

      for (let idx = 0; idx < rowsNeedingCheck.length; idx++) {
        const i = rowsNeedingCheck[idx];
        const row = data[i];
        const threadId = row[1];
        const direction = row[2];
        const threadCount = row[12];
        const trackedEmailDate = row[7]; // Date column

        try {
          const thread = GmailApp.getThreadById(threadId);
          if (!thread) continue;

          const currentThreadCount = thread.getMessageCount();
          if (currentThreadCount <= threadCount) continue;

          const messages = thread.getMessages();
          const replies = [];
          const trackedDate = trackedEmailDate ? new Date(trackedEmailDate) : null;

          for (let j = 0; j < messages.length && replies.length < 3; j++) {
            const replyMessage = messages[j];
            const replyDate = replyMessage.getDate();

            if (trackedDate && replyDate <= trackedDate) continue;

            const replyFrom = replyMessage.getFrom() || '';
            const replyFromLower = replyFrom.toLowerCase();

            let isReply = false;
            if (direction === 'Sent') {
              isReply = replyFromLower.indexOf(userEmailLower) === -1;
            } else {
              isReply = true;
            }

            if (isReply) {
              replies.push({
                date: replyDate,
                from: Utils.sanitizeText(replyFrom),
                subject: Utils.sanitizeText(replyMessage.getSubject() || ''),
                body: CONFIG.STORE_FULL_BODY
                  ? Utils.truncateText(Utils.sanitizeText(replyMessage.getPlainBody() || ''), CONFIG.BODY_MAX_LENGTH)
                  : Utils.truncateText(Utils.sanitizeText(replyMessage.getPlainBody() || ''), 500),
                isFromMe: replyFromLower.indexOf(userEmailLower) !== -1
              });
            }
          }

          if (replies.length > 0) {
            let status = 'Awaiting Reply';
            let bgColor = null;

            if (direction === 'Sent') {
              status = 'Replied';
              bgColor = '#fff2cc';
            } else {
              const iResponded = replies.some(function(r) { return r.isFromMe; });
              status = iResponded ? 'I Responded' : 'They Replied';
              bgColor = iResponded ? '#d9ead3' : '#fff2cc';
            }

            rowUpdates.push({
              rowNumber: i + 1,
              status: status,
              threadCount: currentThreadCount,
              replies: replies,
              bgColor: bgColor
            });
          }

        } catch (e) {
          Logger.log('Error checking thread ' + threadId + ': ' + e);
        }
      }

      // Apply updates
      if (rowUpdates.length > 0) {
        for (const update of rowUpdates) {
          // Status and thread count (columns 12-13)
          sheet.getRange(update.rowNumber, colOffset + 11, 1, 2)
            .setValues([[update.status, update.threadCount]]);

          // Reply data (columns 16-27)
          const replyData = [];
          for (let r = 0; r < 3; r++) {
            if (update.replies[r]) {
              replyData.push(
                update.replies[r].date,
                Utils.truncateText(update.replies[r].from, 500),
                Utils.truncateText(update.replies[r].subject, 1000),
                update.replies[r].body
              );
            } else {
              replyData.push('', '', '', '');
            }
          }
          sheet.getRange(update.rowNumber, colOffset + 15, 1, 12).setValues([replyData]);

          // Background color
          if (update.bgColor) {
            sheet.getRange(update.rowNumber, colOffset, 1, 36).setBackground(update.bgColor);
          }
        }

        Logger.log('Updated ' + rowUpdates.length + ' rows with replies.');
      }

      return { repliesDetected: rowUpdates.length };

    } finally {
      lock.releaseLock();
    }
  }
};

// ============================================================================
// SECTION 5: CONTROLLERS (Entry Points)
// ============================================================================

/**
 * Creates the custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const startDate = PropertyRepository.get(CONFIG.PROPERTY_KEYS.START_DATE);
  const dateDisplay = startDate ? ' (from ' + startDate + ')' : ' (all time)';

  ui.createMenu('Gmail Tracker')
    .addItem('Run Tracker' + dateDisplay, 'runTracker')
    .addSeparator()
    .addItem('Refresh All', 'refreshAll')
    .addItem('Update Analytics', 'updateAnalytics')
    .addSeparator()
    .addSubMenu(ui.createMenu('Campaigns')
      .addItem('Create Campaign', 'createCampaign')
      .addItem('Assign Email to Campaign', 'assignToCampaign')
      .addItem('View Campaign Stats', 'viewCampaignStats')
      .addItem('Manage Campaigns', 'manageCampaigns'))
    .addSubMenu(ui.createMenu('Automation')
      .addItem('Configure Reminders', 'configureReminders')
      .addItem('Check Reminders Now', 'checkAndSendReminders')
      .addItem('View Pending Reminders', 'viewPendingReminders'))
    .addSubMenu(ui.createMenu('Calendar')
      .addItem('Sync All Calendar Data', 'syncCalendarData')
      .addItem('Lookup Contact Calendar', 'lookupContactCalendar'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Subject Filters')
      .addItem('Add Subject Filter', 'addSubjectFilter')
      .addItem('View/Remove Filters', 'manageSubjectFilters')
      .addItem('Clear All Filters', 'clearSubjectFilters'))
    .addSubMenu(ui.createMenu('Date Range')
      .addItem('Pick Start Date...', 'showDatePickerDialog')
      .addItem('Start From Today', 'setStartDateToday')
      .addItem('Clear Start Date', 'clearStartDate')
      .addItem('Re-scan All (Reset Sync)', 'resetSyncCursor'))
    .addSeparator()
    .addSubMenu(ui.createMenu('Settings')
      .addItem('Set Tracker Start Column...', 'setTrackerStartColumn')
      .addItem('Fix Header Labels', 'fixHeaderLabels')
      .addItem('Insert Missing Columns', 'updateColumns')
      .addItem('Setup Sheet (first time)', 'setupSheet')
      .addSeparator()
      .addItem('View Health/Diagnostics', 'viewHealthSheet')
      .addItem('Show Editable Columns', 'showEditableColumns')
      .addItem('Print Summary', 'printLastRunSummary')
      .addSeparator()
      .addItem('Reset Tracker (DANGER)', 'resetTracker'))
    .addToUi();
}

/**
 * Main entry point - runs the email tracker
 */
function runTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const hasColumnOffset = PropertyRepository.get(CONFIG.PROPERTY_KEYS.COLUMN_OFFSET) !== null;

  if (!sheet || !hasColumnOffset) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Setup Required',
      'Email Tracker columns have not been set up. Set them up now?',
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      setupSheet();
      ss.toast('Setup complete! Running tracker...', 'Gmail Tracker', 3);
    } else {
      return;
    }
  }

  ss.toast('Tracking emails...', 'Gmail Tracker', 3);
  TrackerService.trackEmails();
  TrackerService.checkForReplies();
  ss.toast('Tracker complete!', 'Gmail Tracker', 5);
}

/**
 * Refreshes everything - tracks new emails and checks for replies
 */
function refreshAll() {
  TrackerService.trackEmails();
  TrackerService.checkForReplies();
  SpreadsheetApp.getActiveSpreadsheet().toast('Refresh complete!', 'Gmail Tracker', 3);
}

/**
 * Tracks sent/received emails (for trigger use)
 */
function trackEmails() {
  return TrackerService.trackEmails();
}

/**
 * Checks for replies (for trigger use)
 */
function checkForReplies() {
  return TrackerService.checkForReplies();
}

/**
 * Allows user to manually set which column the tracker data starts at.
 * This preserves any existing data in earlier columns.
 */
function setTrackerStartColumn() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SheetRepository.getTrackerSheet();

  if (!sheet) {
    ui.alert('Error', 'Sheet "' + CONFIG.SHEET_NAME + '" not found. Create it first.', ui.ButtonSet.OK);
    return;
  }

  const currentOffset = PropertyRepository.getColumnOffset();
  const lastColumn = sheet.getLastColumn();

  const response = ui.prompt(
    'Set Tracker Start Column',
    'Your existing data appears to end at column ' + Utils.columnToLetter(lastColumn) + '.\n\n' +
    'Current tracker start: Column ' + Utils.columnToLetter(currentOffset) + ' (' + currentOffset + ')\n\n' +
    'Enter the column NUMBER where tracker should start:\n' +
    '  - Column A = 1\n' +
    '  - Column D = 4\n' +
    '  - Column G = 7\n\n' +
    'Your manual data in earlier columns will NOT be touched.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const newOffset = parseInt(input, 10);

  if (isNaN(newOffset) || newOffset < 1) {
    ui.alert('Invalid Input', 'Please enter a valid column number (1 or greater).', ui.ButtonSet.OK);
    return;
  }

  // Warn if they're about to overwrite existing data
  if (newOffset <= lastColumn) {
    const confirmResponse = ui.alert(
      'Warning: Potential Data Overlap',
      'Column ' + Utils.columnToLetter(newOffset) + ' already has data.\n\n' +
      'Starting the tracker here may overwrite existing data.\n\n' +
      'Are you sure you want to start at column ' + Utils.columnToLetter(newOffset) + '?',
      ui.ButtonSet.YES_NO
    );
    if (confirmResponse !== ui.Button.YES) return;
  }

  PropertyRepository.setColumnOffset(newOffset);

  ui.alert(
    'Start Column Set',
    'Tracker will now start at column ' + Utils.columnToLetter(newOffset) + ' (column ' + newOffset + ').\n\n' +
    'Columns A through ' + Utils.columnToLetter(newOffset - 1) + ' will be preserved.\n\n' +
    'Run "Setup Sheet" to create/update headers at this position.',
    ui.ButtonSet.OK
  );
}

/**
 * Sets up the spreadsheet with headers (idempotent)
 * Respects MANUAL_START_COLUMN if set, otherwise auto-detects
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }

  // Determine start column: manual setting > existing offset > auto-detect
  let startColumn;

  if (MANUAL_START_COLUMN > 0) {
    // Use the hardcoded manual setting
    startColumn = MANUAL_START_COLUMN;
    Logger.log('Using MANUAL_START_COLUMN: ' + startColumn);
  } else {
    // Check if we already have an offset saved
    const existingOffset = PropertyRepository.get(CONFIG.PROPERTY_KEYS.COLUMN_OFFSET);
    if (existingOffset) {
      startColumn = parseInt(existingOffset, 10);
      Logger.log('Using existing offset: ' + startColumn);
    } else {
      // Auto-detect: start after existing data
      const lastColumn = sheet.getLastColumn();
      startColumn = lastColumn > 0 ? lastColumn + 1 : 1;
      Logger.log('Auto-detected start column: ' + startColumn);
    }
  }

  // Save column offset
  PropertyRepository.setColumnOffset(startColumn);

  // Set up headers (uses idempotent ensureHeaders internally)
  sheet.getRange(1, startColumn, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  sheet.getRange(1, startColumn, 1, CONFIG.HEADERS.length).setFontWeight('bold');

  if (startColumn === 1) {
    sheet.setFrozenRows(1);
  }

  // Auto-resize columns
  for (let i = startColumn; i < startColumn + CONFIG.HEADERS.length; i++) {
    sheet.autoResizeColumn(i);
  }

  // Create Gmail label
  try {
    GmailService.getOrCreateLabel();
  } catch (e) {
    Logger.log('Label creation: ' + e);
  }

  const colLetter = Utils.columnToLetter(startColumn);
  SpreadsheetApp.getUi().alert(
    'Setup complete! Email tracker columns added starting at column ' + colLetter +
    '.\n\nSet up time-driven triggers for automatic syncing.'
  );
}

/**
 * Fixes header labels only (does not insert columns)
 */
function fixHeaderLabels() {
  const sheet = SheetRepository.getTrackerSheet();
  const ui = SpreadsheetApp.getUi();

  if (!sheet) {
    ui.alert('Sheet not found.');
    return;
  }

  const colOffset = PropertyRepository.getColumnOffset();
  const currentHeaders = sheet.getRange(1, colOffset, 1, CONFIG.HEADERS.length).getValues()[0];

  const changes = [];
  for (let i = 0; i < CONFIG.HEADERS.length; i++) {
    if (currentHeaders[i] !== CONFIG.HEADERS[i]) {
      changes.push('Column ' + Utils.columnToLetter(colOffset + i) +
        ': "' + (currentHeaders[i] || '(empty)') + '" -> "' + CONFIG.HEADERS[i] + '"');
    }
  }

  if (changes.length === 0) {
    ui.alert('Headers are already correct!');
    return;
  }

  const response = ui.alert(
    'Fix Header Labels?',
    'This will rename ' + changes.length + ' header(s). Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  sheet.getRange(1, colOffset, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
  sheet.getRange(1, colOffset, 1, CONFIG.HEADERS.length).setFontWeight('bold');

  ui.alert('Fixed ' + changes.length + ' header label(s).');
}

/**
 * Inserts missing columns (idempotent)
 */
function updateColumns() {
  const ui = SpreadsheetApp.getUi();

  try {
    const result = SheetRepository.ensureHeaders();

    if (result.added.length === 0) {
      ui.alert('All columns are up to date!');
    } else {
      ui.alert('Inserted ' + result.added.length + ' missing column(s):\n' + result.added.join(', '));
    }
  } catch (e) {
    ui.alert('Error: ' + e.message);
  }
}

/**
 * Resets the sync cursor to re-scan all emails
 */
function resetSyncCursor() {
  PropertyRepository.delete(CONFIG.PROPERTY_KEYS.LAST_SYNC);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Sync cursor reset! Next run will scan last 30 days.',
    'Reset Complete',
    5
  );
}

/**
 * Views the health/diagnostics sheet
 */
function viewHealthSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const healthSheet = HealthRepository.getHealthSheet();
  ss.setActiveSheet(healthSheet);
  ss.toast('Health sheet shows sync run history.', 'Health/Diagnostics', 5);
}

/**
 * Prints the last run summary
 */
function printLastRunSummary() {
  if (RunSummary.getData()) {
    RunSummary.printSummary();
    SpreadsheetApp.getActiveSpreadsheet().toast('Summary printed to Apps Script logs.', 'Summary', 3);
  } else {
    SpreadsheetApp.getActiveSpreadsheet().toast('No run data available. Run the tracker first.', 'Summary', 3);
  }
}

/**
 * Shows editable columns guide
 */
function showEditableColumns() {
  const colOffset = PropertyRepository.getColumnOffset();
  const message =
    'COLUMN GUIDE\n\n' +
    'SAFE TO EDIT:\n' +
    '- Columns before ' + Utils.columnToLetter(colOffset) + ' (your pre-existing data)\n' +
    '- Contact Notes column\n\n' +
    'EDITABLE BUT MAY BE UPDATED:\n' +
    '- Status column\n' +
    '- Calendar columns (overwritten on sync)\n\n' +
    'DO NOT EDIT:\n' +
    '- Email data columns (Message ID, Thread, From, To, etc.)\n' +
    '- Reply data columns';

  SpreadsheetApp.getUi().alert('Editable Columns Guide', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Sets start date to today
 */
function setStartDateToday() {
  const today = Utils.formatDate(new Date(), 'yyyy-MM-dd');
  PropertyRepository.setStartDate(today);
  SpreadsheetApp.getActiveSpreadsheet().toast('Start date set to ' + today, 'Done', 3);
}

/**
 * Clears the start date
 */
function clearStartDate() {
  PropertyRepository.delete(CONFIG.PROPERTY_KEYS.START_DATE);
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Start date cleared', 'Done', 3);
  } catch (e) {
    // Called from dialog
  }
}

/**
 * Shows date picker dialog
 */
function showDatePickerDialog() {
  const currentDate = PropertyRepository.get(CONFIG.PROPERTY_KEYS.START_DATE) || '';
  const todayStr = Utils.formatDate(new Date(), 'yyyy-MM-dd');

  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;padding:20px;text-align:center}' +
    'input[type=date]{font-size:16px;padding:10px;margin:10px}' +
    'button{padding:10px 20px;margin:5px;border:none;border-radius:4px;cursor:pointer}' +
    '.primary{background:#4285f4;color:white}' +
    '.secondary{background:#f1f3f4;color:#333}' +
    '</style></head><body>' +
    '<h3>Select Start Date</h3>' +
    '<p>Current: <strong>' + (currentDate || 'Not set') + '</strong></p>' +
    '<input type="date" id="datePicker" value="' + currentDate + '" max="' + todayStr + '">' +
    '<div><button class="primary" onclick="confirm()">Confirm</button>' +
    '<button class="secondary" onclick="clear_()">Clear</button>' +
    '<button class="secondary" onclick="google.script.host.close()">Close</button></div>' +
    '<script>' +
    'function confirm(){var d=document.getElementById("datePicker").value;if(d)google.script.run.withSuccessHandler(function(){google.script.host.close()}).confirmStartDate(d)}' +
    'function clear_(){google.script.run.withSuccessHandler(function(){google.script.host.close()}).clearStartDate()}' +
    '</script></body></html>'
  ).setWidth(320).setHeight(200);

  SpreadsheetApp.getUi().showModalDialog(html, 'Select Start Date');
}

/**
 * Confirms and stores start date (called from dialog)
 * @param {string} isoDate - Date in YYYY-MM-DD format
 */
function confirmStartDate(isoDate) {
  if (!Utils.isValidISODate(isoDate)) {
    throw new Error('Invalid date format.');
  }

  const selectedDate = new Date(isoDate + 'T00:00:00');
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (selectedDate > today) {
    throw new Error('Cannot select future date.');
  }

  PropertyRepository.setStartDate(isoDate);
}

/**
 * Resets the tracker (deletes all data)
 */
function resetTracker() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Confirm Reset',
    'This will delete all tracked email data. Are you sure?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return;

  ss.toast('Resetting...', 'Reset', -1);

  const colOffset = PropertyRepository.getColumnOffset();
  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, colOffset, lastRow - 1, CONFIG.HEADERS.length).clearContent();
  }

  // Remove labels
  const label = GmailApp.getUserLabelByName(CONFIG.TRACKED_LABEL);
  if (label) {
    let threads = label.getThreads(0, 100);
    let total = 0;
    while (threads.length > 0) {
      for (const thread of threads) {
        thread.removeLabel(label);
        total++;
      }
      threads = label.getThreads(0, 100);
    }
    Logger.log('Removed label from ' + total + ' threads');
  }

  PropertyRepository.delete(CONFIG.PROPERTY_KEYS.LAST_SYNC);
  ui.alert('Tracker reset complete!');
}

// ============================================================================
// STUB FUNCTIONS - Preserved for compatibility (implement as needed)
// ============================================================================

function updateAnalytics() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Analytics update - implement based on original', 'Analytics', 3);
}

function createCampaign() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Create Campaign - implement based on original', 'Campaigns', 3);
}

function assignToCampaign() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Assign to Campaign - implement based on original', 'Campaigns', 3);
}

function viewCampaignStats() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Campaign Stats - implement based on original', 'Campaigns', 3);
}

function manageCampaigns() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Manage Campaigns - implement based on original', 'Campaigns', 3);
}

function configureReminders() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Configure Reminders - implement based on original', 'Automation', 3);
}

function checkAndSendReminders() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Check Reminders - implement based on original', 'Automation', 3);
}

function viewPendingReminders() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Pending Reminders - implement based on original', 'Automation', 3);
}

function syncCalendarData() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Calendar Sync - implement based on original', 'Calendar', 3);
}

function lookupContactCalendar() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Calendar Lookup - implement based on original', 'Calendar', 3);
}

function addSubjectFilter() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Add Filter - implement based on original', 'Filters', 3);
}

function manageSubjectFilters() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Manage Filters - implement based on original', 'Filters', 3);
}

function clearSubjectFilters() {
  PropertyRepository.delete(CONFIG.PROPERTY_KEYS.SUBJECT_FILTERS);
  SpreadsheetApp.getActiveSpreadsheet().toast('Subject filters cleared', 'Done', 3);
}

// For backwards compatibility
function trackSentEmails() {
  trackEmails();
}
