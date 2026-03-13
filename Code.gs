/**
 * ================================================================================
 * PROJECT MANAGEMENT TASK NOTIFICATION SYSTEM
 * ================================================================================
 * 
 * This Google Apps Script adds intelligent task change notifications to your
 * project management spreadsheet.
 * 
 * Features:
 * - Automatically detects changes in tasks across all sheets
 * - Sends consolidated emails only to assignees whose tasks changed
 * - Tracks full history for comparison and audit trail
 * - Shows exactly what changed (old → new values)
 * 
 * Column Structure Expected:
 * A: Task | B: Priority | C: Assignee (email) | D: Status | E: Init Date
 * F: Finish Date | G: Hits | H: Product | I: Notes
 * 
 * ================================================================================
 */




// ================================================================================
// SECTION 1: CONFIGURATION
// ================================================================================

const CONFIG = {
  // External history and backup storage
  HISTORY: {
    FOLDER_NAME: 'history_and_backup',
    SPREADSHEET_NAME: 'history_backup',
    RETENTION_DAYS: 45,
    TASKS_CURRENT_SHEET_NAME: 'Tasks_Current',
    MEETINGS_CURRENT_SHEET_NAME: 'Meetings_Current',
    TASKS_BACKUP_SHEET_PREFIX: 'Tasks_Backup',
    MEETINGS_BACKUP_SHEET_PREFIX: 'Meetings_Backup'
  },

  // Column indices (0-based, matching spreadsheet structure)
  COLUMNS: {
    TASK: 0,        // Column A
    PRIORITY: 1,    // Column B
    EMAIL: 2,       // Column C (Assignee email)
    STATUS: 3,      // Column D
    INIT_DATE: 4,   // Column E
    FINISH_DATE: 5, // Column F
    HITS: 6,        // Column G (not monitored for changes)
    PRODUCT: 7,     // Column H
    NOTES: 8        // Column I
  },

  // Email configuration
  EMAIL_SUBJECT: 'Actualización de tareas - se requiere su atención',

  // Spreadsheet UI
  SPREADSHEET_MENU_NAME: '📧 Notificaciones',
  SPREADSHEET_MENU_ITEM: 'Notificar tareas a los asignados',

  // Spreadsheet structure
  HEADER_ROW: 1,
  FIRST_DATA_ROW: 2,

  // Fields to monitor for changes (excludes HITS as it changes too frequently)
  MONITORED_FIELDS: ['TASK', 'PRIORITY', 'EMAIL', 'STATUS', 'INIT_DATE', 'FINISH_DATE', 'PRODUCT', 'NOTES'],

  // Field display names for emails
  FIELD_NAMES: {
    TASK: 'Tarea',
    PRIORITY: 'Prioridad',
    EMAIL: 'Asignado',
    STATUS: 'Estado',
    INIT_DATE: 'Inicio',
    FINISH_DATE: 'Finalización',
    HITS: 'Hits',
    PRODUCT: 'Producto',
    NOTES: 'Notas'
  },

  // Note: Task ID must remain at index 2
  HISTORY_SHEET_HEADERS: [
    'Timestamp', 'Hoja', 'ID Tarea', 'Tarea', 'Prioridad', 'Asignado',
    'Estado', 'Inicio', 'Finalización', 'Hits', 'Producto', 'Notas',
    'Acción', 'Notificado'
  ],

  ALERTS: {
    NO_TASKS: 'No se encontraron tareas en ninguna hoja. Por favor, agregue tareas primero.',
    SUMMARY_TITLE: '✓ Resumen de notificaciones',
    SUMMARY_MESSAGE: (totalTasks, sheetsProcessed, newCount, changedCount, deletedCount, unchangedCount, emailsSent) =>
      `• ${totalTasks} tarea(s) escaneada(s) en ${sheetsProcessed} hoja(s)\n` +
      `• ${newCount} tarea(s) nueva(s) detectada(s)\n` +
      `• ${changedCount} tarea(s) cambiada(s)\n` +
      `• ${deletedCount} tarea(s) eliminada(s)\n` +
      `• ${unchangedCount} tarea(s) sin cambios (no notificadas)\n\n` +
      `📧 Correos enviados a ${emailsSent} asignado(s)\n\n` +
      `Snapshot guardado en ${CONFIG.HISTORY.SPREADSHEET_NAME}/${CONFIG.HISTORY.TASKS_CURRENT_SHEET_NAME}.`,
    EMAIL_NOTIFICATION_FAILED: (email) => `Error al enviar correo a ${email}.`
  },

  // ================================================================================
  // MEETINGS CONFIGURATION
  // ================================================================================

  MEETINGS: {
    SHEET_NAME: 'Reuniones',
    EMAIL_SUBJECT: 'Reunión de equipo - convocatoria',
    MENU_ITEM: 'Notificar reuniones',
    COMPLETED_STATUS: 'Completada',
    CALENDAR_DEFAULT_DURATION_MINUTES: 30,

    COLUMNS: {
      TITLE: 0,         // Columna A
      ATTENDEES: 1,     // Columna B
      STATUS: 2,        // Columna C
      DATE: 3,          // Columna D (Fecha)
      TIME: 4,          // Columna E (Hora)
      AGENDA: 5,        // Columna F
      DOCUMENTATION: 6  // Columna G
    },

    HEADER_ROW: 1,
    FIRST_DATA_ROW: 2,

    MONITORED_FIELDS: ['TITLE', 'ATTENDEES', 'STATUS', 'DATE', 'TIME', 'AGENDA', 'DOCUMENTATION'],

    FIELD_NAMES: {
      TITLE: 'Título',
      ATTENDEES: 'Asistentes',
      STATUS: 'Estado',
      DATE: 'Fecha',
      TIME: 'Hora',
      AGENDA: 'Agenda',
      DOCUMENTATION: 'Documentación'
    },

    HISTORY_HEADERS: [
      'Timestamp', 'ID Reunión', 'Título', 'Asistentes',
      'Estado', 'Fecha_hora', 'Agenda', 'Documentación',
      'Acción', 'Notificado'
    ],

    ALERTS: {
      NO_MEETINGS: 'No se encontraron reuniones en la hoja. Por favor, agregue reuniones primero.',
      SUMMARY_TITLE: '✓ Resumen de notificaciones de reuniones',
      SUMMARY_MESSAGE: (totalMeetings, newCount, changedCount, deletedCount, postponedCount, cancelledCount, unchangedCount, emailsSent) =>
        `• ${totalMeetings} reunión(es) escaneada(s)\n` +
        `• ${newCount} reunión(es) nueva(s) detectada(s)\n` +
        `• ${changedCount} reunión(es) cambiada(s)\n` +
        `• ${deletedCount} reunión(es) eliminada(s)\n` +
        `• ${postponedCount} reunión(es) pospuesta(s) detectada(s)\n` +
        `• ${cancelledCount} reunión(es) cancelada(s) detectada(s)\n` +
        `• ${unchangedCount} reunión(es) sin cambios (no notificadas)\n\n` +
        `📧 Correos enviados a ${emailsSent} asistente(s)\n\n` +
        `Snapshot guardado en ${CONFIG.HISTORY.SPREADSHEET_NAME}/${CONFIG.HISTORY.MEETINGS_CURRENT_SHEET_NAME}.`
    }
  }
};


// ================================================================================
// SECTION 2: MENU SETUP
// ================================================================================

/**
 * Creates custom menu when spreadsheet opens
 * This function runs automatically on spreadsheet open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(CONFIG.SPREADSHEET_MENU_NAME)
    .addItem(CONFIG.SPREADSHEET_MENU_ITEM, 'notifyTaskAssignees')
    .addItem(CONFIG.MEETINGS.MENU_ITEM, 'notifyMeetingAttendees')
    .addToUi();
}


// ================================================================================
// SECTION 3: MAIN ORCHESTRATOR
// ================================================================================

/**
 * Main function that orchestrates the entire notification process
 * Called when user clicks "Notify task assignees" menu button
 */
function notifyTaskAssignees() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Step 1: Ensure History sheet exists
    createHistorySheet();

    // Step 1.1: Archive records older than retention window
    try {
      rotateTasksHistoryIfNeeded();
    } catch (archiveError) {
      Logger.log(`Task history rotation warning: ${archiveError.toString()}`);
    }

    // Step 2: Get all current tasks from all sheets
    const currentTasks = getAllTasks();

    if (currentTasks.length === 0) {
      SpreadsheetApp.getUi().alert(CONFIG.ALERTS.NO_TASKS);
      return;
    }

    // Step 3: Load previous snapshot from History sheet
    const previousSnapshot = getPreviousSnapshot();

    // Step 4: Compare current vs previous and detect changes
    const changes = compareSnapshots(currentTasks, previousSnapshot);

    // Step 5: Send notifications only for changed tasks
    const notificationStats = sendChangeNotifications(changes, spreadsheet.getUrl());

    // Step 6: Save current snapshot to History sheet
    saveSnapshot(currentTasks, changes);

    // Step 7: Show summary to user
    const totalTasks = currentTasks.length;
    const sheetsProcessed = new Set(currentTasks.map(t => t.sheetName)).size;
    const newCount = changes.new.length;
    const changedCount = changes.changed.length;
    const deletedCount = changes.deleted.length;
    const unchangedCount = changes.unchanged.length;

    const summary = CONFIG.ALERTS.SUMMARY_MESSAGE(
      totalTasks,
      sheetsProcessed,
      newCount,
      changedCount,
      deletedCount,
      unchangedCount,
      notificationStats.emailsSent
    );

    SpreadsheetApp.getUi().alert(CONFIG.ALERTS.SUMMARY_TITLE, summary, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in notifyTaskAssignees: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}

/**
 * Main function for meeting notifications
 * Called when user clicks "Notify meetings" menu button
 */
function notifyMeetingAttendees() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Validate required Reuniones structure before processing.
    const structureCheck = validateMeetingsSheetStructure(spreadsheet);
    if (!structureCheck.isValid) {
      SpreadsheetApp.getUi().alert(structureCheck.message);
      return;
    }

    // Step 1: Ensure Meetings History sheet exists
    createMeetingsHistorySheet();

    // Step 1.1: Archive records older than retention window
    try {
      rotateMeetingsHistoryIfNeeded();
    } catch (archiveError) {
      Logger.log(`Meetings history rotation warning: ${archiveError.toString()}`);
    }

    // Step 2: Get all current meetings from Reuniones sheet
    const currentMeetings = getAllMeetings();

    if (currentMeetings.length === 0) {
      SpreadsheetApp.getUi().alert(CONFIG.MEETINGS.ALERTS.NO_MEETINGS);
      return;
    }

    // Step 3: Load previous snapshot from Meetings History sheet
    const previousSnapshot = getPreviousMeetingsSnapshot();

    // Step 4: Compare current vs previous and detect changes
    const changes = compareMeetingsSnapshots(currentMeetings, previousSnapshot);

    // Step 5: Send notifications only for changed meetings
    const notificationStats = sendMeetingNotifications(changes, spreadsheet.getUrl());

    // Step 6: Save current snapshot to Meetings History sheet
    saveMeetingsSnapshot(currentMeetings, changes);

    // Step 7: Show summary to user
    const totalMeetings = currentMeetings.length;
    const newCount = changes.new.length;
    const changedCount = changes.changed.length;
    const deletedCount = changes.deleted.length;
    const postponedCount = changes.postponed.length;
    const cancelledCount = changes.cancelled.length;
    const unchangedCount = changes.unchanged.length;

    const summary = CONFIG.MEETINGS.ALERTS.SUMMARY_MESSAGE(
      totalMeetings,
      newCount,
      changedCount,
      deletedCount,
      postponedCount,
      cancelledCount,
      unchangedCount,
      notificationStats.emailsSent
    );

    SpreadsheetApp.getUi().alert(CONFIG.MEETINGS.ALERTS.SUMMARY_TITLE, summary, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (error) {
    Logger.log('Error in notifyMeetingAttendees: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error: ' + error.toString());
  }
}



// ================================================================================
// SECTION 4: TASK DATA COLLECTION
// ================================================================================

/**
 * Scans all sheets (except History) and extracts all tasks
 * @returns {Array} Array of task objects with all fields
 */
function getAllTasks() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const allTasks = [];

  sheets.forEach(sheet => {
    // Skip Meetings sheet
    if (
      sheet.getName() === CONFIG.MEETINGS.SHEET_NAME
    ) {
      return;
    }

    const lastRow = sheet.getLastRow();

    // Skip empty sheets or sheets with only headers
    if (lastRow < CONFIG.FIRST_DATA_ROW) {
      return;
    }

    // Get all data rows (skip header)
    const dataRange = sheet.getRange(CONFIG.FIRST_DATA_ROW, 1, lastRow - CONFIG.HEADER_ROW, 9);
    const data = dataRange.getValues();

    data.forEach((row, index) => {
      const actualRowIndex = CONFIG.FIRST_DATA_ROW + index;
      const taskOrTasks = getTaskFromRow(sheet, row, actualRowIndex);

      // getTaskFromRow now returns null or an array of task objects
      if (taskOrTasks && Array.isArray(taskOrTasks)) {
        // Add all tasks from the array (one per assignee)
        allTasks.push(...taskOrTasks);
      }
    });
  });

  return allTasks;
}

/**
 * Scans the Reuniones sheet and extracts all meetings
 * @returns {Array} Array of meeting objects with all fields
 */
function getAllMeetings() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.MEETINGS.SHEET_NAME);
  const allMeetings = [];

  // Return empty array if Reuniones sheet doesn't exist
  if (!sheet) {
    Logger.log('Reuniones sheet not found');
    return allMeetings;
  }

  const lastRow = sheet.getLastRow();

  // Skip if sheet is empty or has only headers
  if (lastRow < CONFIG.MEETINGS.FIRST_DATA_ROW) {
    return allMeetings;
  }

  // Get all data rows (skip header) using the strict split format: Fecha (D) + Hora (E)
  const dataRange = sheet.getRange(CONFIG.MEETINGS.FIRST_DATA_ROW, 1, lastRow - CONFIG.MEETINGS.HEADER_ROW, 7);
  const data = dataRange.getValues();

  data.forEach((row, index) => {
    const actualRowIndex = CONFIG.MEETINGS.FIRST_DATA_ROW + index;
    const meetingOrMeetings = getMeetingFromRow(sheet, row, actualRowIndex);

    // getMeetingFromRow returns null or an array of meeting objects
    if (meetingOrMeetings && Array.isArray(meetingOrMeetings)) {
      // Add all meetings from the array (one per attendee)
      allMeetings.push(...meetingOrMeetings);
    }
  });

  return allMeetings;
}

/**
 * Validates that Reuniones uses the required split Fecha/Hora structure.
 * @param {Spreadsheet} spreadsheet - Active spreadsheet
 * @returns {{isValid: boolean, message: string}} Validation result
 */
function validateMeetingsSheetStructure(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(CONFIG.MEETINGS.SHEET_NAME);

  if (!sheet) {
    return {
      isValid: false,
      message: `No se encontró la hoja "${CONFIG.MEETINGS.SHEET_NAME}".`
    };
  }

  if (sheet.getLastColumn() < 7) {
    return {
      isValid: false,
      message:
        'La hoja Reuniones requiere 7 columnas en este orden: ' +
        'Título, Asistentes, Estado, Fecha, Hora, Agenda, Documentación.'
    };
  }

  const headers = sheet.getRange(CONFIG.MEETINGS.HEADER_ROW, 1, 1, 7).getValues()[0]
    .map(value => normalizeStringForComparison(value).toLowerCase());

  const expectedHeaders = ['titulo', 'asistentes', 'estado', 'fecha', 'hora', 'agenda', 'documentacion'];

  const headersMatch = expectedHeaders.every((expected, index) => {
    const actual = removeDiacritics(headers[index]);
    return actual === expected;
  });

  if (!headersMatch) {
    return {
      isValid: false,
      message:
        'Los encabezados de Reuniones no cumplen el formato requerido. ' +
        'Use exactamente: Título, Asistentes, Estado, Fecha, Hora, Agenda, Documentación.'
    };
  }

  return { isValid: true, message: '' };
}

/**
 * Extracts meeting data from a spreadsheet row
 * @param {Sheet} sheet - The sheet containing the meeting
 * @param {Array} row - Row data array
 * @param {number} rowIndex - Actual row number in sheet
 * @returns {Array|null} Array of meeting objects (one per attendee) or null if no valid emails
 */
function getMeetingFromRow(sheet, row, rowIndex) {
  const attendeesString = row[CONFIG.MEETINGS.COLUMNS.ATTENDEES];

  // Skip rows without attendees
  if (!attendeesString || attendeesString.toString().trim() === '') {
    return null;
  }

  // Parse multiple emails from the string
  const emailList = parseMultipleEmails(attendeesString);

  // Skip if no valid emails found
  if (emailList.length === 0) {
    return null;
  }

  const title = sanitizeValue(row[CONFIG.MEETINGS.COLUMNS.TITLE]) || 'Reunión';
  const dateTimeParts = extractMeetingDateTimeParts(row);

  // Create one meeting object per attendee
  return emailList.map(email => ({
    meetingId: generateMeetingId(sheet.getName(), rowIndex, title),
    sheetName: sheet.getName(),
    rowNumber: rowIndex,
    title: title,
    attendees: attendeesString.toString().trim(), // Keep full list for display
    email: email, // Individual email for routing
    status: sanitizeValue(row[CONFIG.MEETINGS.COLUMNS.STATUS]),
    datetime: dateTimeParts.datetime,
    date: dateTimeParts.date,
    time: dateTimeParts.time,
    agenda: sanitizeValue(dateTimeParts.agenda),
    documentation: sanitizeValue(dateTimeParts.documentation)
  }));
}

/**
 * Generates unique meeting ID from sheet name, row number, and title
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowIndex - Row number
 * @param {string} title - Title of the meeting
 * @returns {string} Unique meeting ID
 */
function generateMeetingId(sheetName, rowIndex, title) {
  return `${sheetName}_Row${rowIndex}_${title.substring(0, 30).replace(/\s+/g, '_')}`;
}


/**
 * Extracts task data from a spreadsheet row
 * @param {Sheet} sheet - The sheet containing the task
 * @param {Array} row - Row data array
 * @param {number} rowIndex - Actual row number in sheet
 * @returns {Array|null} Array of task objects (one per assignee) or null if no valid emails
 */
function getTaskFromRow(sheet, row, rowIndex) {
  const emailString = row[CONFIG.COLUMNS.EMAIL];

  // Skip rows without email
  if (!emailString || emailString.toString().trim() === '') {
    return null;
  }

  // Parse multiple emails from the string
  const emailList = parseMultipleEmails(emailString);

  // Skip if no valid emails found
  if (emailList.length === 0) {
    return null;
  }

  // Create one task object per assignee
  return emailList.map(email => ({
    taskId: generateTaskId(sheet.getName(), rowIndex),
    sheetName: sheet.getName(),
    rowNumber: rowIndex,
    task: sanitizeValue(row[CONFIG.COLUMNS.TASK]),
    priority: sanitizeValue(row[CONFIG.COLUMNS.PRIORITY]),
    assignees: emailString.toString().trim(), // Keep full list for display
    email: email, // Individual email for routing
    status: sanitizeValue(row[CONFIG.COLUMNS.STATUS]),
    initDate: row[CONFIG.COLUMNS.INIT_DATE],
    finishDate: row[CONFIG.COLUMNS.FINISH_DATE],
    hits: sanitizeValue(row[CONFIG.COLUMNS.HITS]),
    product: sanitizeValue(row[CONFIG.COLUMNS.PRODUCT]),
    notes: sanitizeValue(row[CONFIG.COLUMNS.NOTES])
  }));
}


// ================================================================================
// SECTION 5: HISTORY MANAGEMENT
// ================================================================================

/**
 * Gets or creates the Drive folder that stores history and backups.
 * Folder is created inside the same parent folder as the active board spreadsheet.
 * @returns {Folder} Destination folder
 */
function getOrCreateHistoryFolder() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const boardFile = DriveApp.getFileById(activeSpreadsheet.getId());
  const parents = boardFile.getParents();

  if (!parents.hasNext()) {
    throw new Error('El tablero principal debe estar dentro de una carpeta de Drive para crear history_and_backup en esa misma ubicación.');
  }

  const boardParentFolder = parents.next();
  const folders = boardParentFolder.getFoldersByName(CONFIG.HISTORY.FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  const folder = boardParentFolder.createFolder(CONFIG.HISTORY.FOLDER_NAME);
  Logger.log(`Created history folder in board parent folder: ${CONFIG.HISTORY.FOLDER_NAME}`);
  return folder;
}

/**
 * Gets or creates the external spreadsheet that stores current history and backups.
 * @returns {Spreadsheet} History spreadsheet
 */
function getOrCreateHistorySpreadsheet() {
  const folder = getOrCreateHistoryFolder();
  const files = folder.getFilesByName(CONFIG.HISTORY.SPREADSHEET_NAME);

  while (files.hasNext()) {
    const file = files.next();
    try {
      return SpreadsheetApp.openById(file.getId());
    } catch (error) {
      Logger.log(`Skipping non-spreadsheet file while searching history file: ${file.getName()}`);
    }
  }

  const spreadsheet = SpreadsheetApp.create(CONFIG.HISTORY.SPREADSHEET_NAME);
  const file = DriveApp.getFileById(spreadsheet.getId());
  folder.addFile(file);

  // Keep root uncluttered if script has permission to remove from My Drive root.
  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (error) {
    Logger.log('Could not remove history spreadsheet from root folder. Continuing.');
  }

  Logger.log(`Created history spreadsheet: ${CONFIG.HISTORY.SPREADSHEET_NAME}`);
  return spreadsheet;
}

/**
 * Ensures a sheet exists with expected headers in the external history spreadsheet.
 * @param {Spreadsheet} historySpreadsheet - External history spreadsheet
 * @param {string} sheetName - Target sheet name
 * @param {Array} headers - Required header row
 * @returns {Sheet} Existing or newly created sheet
 */
function getOrCreateHistoryDataSheet(historySpreadsheet, sheetName, headers) {
  let sheet = historySpreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = historySpreadsheet.insertSheet(sheetName);
  }

  const mustInitializeHeaders = sheet.getLastRow() < 1;

  if (mustInitializeHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const headersMatch = headers.every((header, index) =>
      normalizeStringForComparison(existingHeaders[index]) === normalizeStringForComparison(header)
    );

    if (!headersMatch) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * Builds a deterministic signature for one history row.
 * @param {Array} row - History row values
 * @returns {string} Unique row signature
 */
function buildHistoryRowSignature(row) {
  const timezone = Session.getScriptTimeZone();
  return row.map(value => {
    if (value instanceof Date) {
      return Utilities.formatDate(value, timezone, "yyyy-MM-dd'T'HH:mm:ss");
    }
    return normalizeStringForComparison(value);
  }).join('\u241F');
}

/**
 * Appends only rows that do not already exist in a target history sheet.
 * @param {Sheet} targetSheet - Backup target sheet
 * @param {Array<Array>} rows - Rows to append
 * @returns {number} Number of rows appended
 */
function appendUniqueRowsToHistorySheet(targetSheet, rows) {
  if (!targetSheet || !rows || rows.length === 0) {
    return 0;
  }

  const existingSignatures = new Set();
  const headersLength = rows[0].length;

  if (targetSheet.getLastRow() > 1) {
    const existing = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, headersLength).getValues();
    existing.forEach(row => existingSignatures.add(buildHistoryRowSignature(row)));
  }

  const uniqueRows = rows.filter(row => {
    const signature = buildHistoryRowSignature(row);
    if (existingSignatures.has(signature)) {
      return false;
    }
    existingSignatures.add(signature);
    return true;
  });

  if (uniqueRows.length === 0) {
    return 0;
  }

  targetSheet.getRange(targetSheet.getLastRow() + 1, 1, uniqueRows.length, uniqueRows[0].length).setValues(uniqueRows);
  return uniqueRows.length;
}

/**
 * Computes deterministic 45-day window bounds for backup page names.
 * @param {Date} timestamp - Row timestamp
 * @returns {{start: Date, end: Date, startKey: string, endKey: string}} Window details
 */
function getBackupWindowBounds(timestamp) {
  const timezone = Session.getScriptTimeZone();
  const windowMs = CONFIG.HISTORY.RETENTION_DAYS * 24 * 60 * 60 * 1000;
  const windowIndex = Math.floor(timestamp.getTime() / windowMs);
  const start = new Date(windowIndex * windowMs);
  const end = new Date((windowIndex + 1) * windowMs - 1);

  return {
    start: start,
    end: end,
    startKey: Utilities.formatDate(start, timezone, 'yyyy-MM-dd'),
    endKey: Utilities.formatDate(end, timezone, 'yyyy-MM-dd')
  };
}

/**
 * Rotates old rows from a current history page into 45-day backup pages.
 * @param {string} currentSheetName - Current active history sheet
 * @param {Array} headers - Expected headers
 * @param {string} backupPrefix - Prefix for backup page names
 */
function rotateHistorySheetIfNeeded(currentSheetName, headers, backupPrefix) {
  const historySpreadsheet = getOrCreateHistorySpreadsheet();
  const currentSheet = getOrCreateHistoryDataSheet(historySpreadsheet, currentSheetName, headers);

  if (currentSheet.getLastRow() <= 1) {
    return;
  }

  const now = new Date();
  const retentionMs = CONFIG.HISTORY.RETENTION_DAYS * 24 * 60 * 60 * 1000;
  const rows = currentSheet.getRange(2, 1, currentSheet.getLastRow() - 1, headers.length).getValues();
  const keepRows = [];
  const archiveGroups = new Map();

  rows.forEach(row => {
    const timestamp = parseLocalizedDateOrDateTime(row[0]);

    // Keep unknown timestamp formats in current history to avoid accidental data loss.
    if (!timestamp) {
      keepRows.push(row);
      return;
    }

    const ageMs = now.getTime() - timestamp.getTime();
    if (ageMs < retentionMs) {
      keepRows.push(row);
      return;
    }

    const window = getBackupWindowBounds(timestamp);
    const groupKey = `${window.startKey}_${window.endKey}`;

    if (!archiveGroups.has(groupKey)) {
      archiveGroups.set(groupKey, { window: window, rows: [] });
    }

    archiveGroups.get(groupKey).rows.push(row);
  });

  if (archiveGroups.size === 0) {
    return;
  }

  archiveGroups.forEach(group => {
    const backupSheetName = `${backupPrefix}_${group.window.startKey}_to_${group.window.endKey}`;
    const backupSheet = getOrCreateHistoryDataSheet(historySpreadsheet, backupSheetName, headers);
    const appendedCount = appendUniqueRowsToHistorySheet(backupSheet, group.rows);
    Logger.log(`Archived ${appendedCount} row(s) into ${backupSheetName}`);
  });

  const previousLastRow = currentSheet.getLastRow();
  currentSheet.getRange(2, 1, previousLastRow - 1, headers.length).clearContent();

  if (keepRows.length > 0) {
    currentSheet.getRange(2, 1, keepRows.length, headers.length).setValues(keepRows);
  }

  const desiredRows = Math.max(keepRows.length + 1, 2);
  if (currentSheet.getMaxRows() > desiredRows) {
    currentSheet.deleteRows(desiredRows + 1, currentSheet.getMaxRows() - desiredRows);
  }
}

/**
 * Rotates task history when rows are older than retention threshold.
 */
function rotateTasksHistoryIfNeeded() {
  rotateHistorySheetIfNeeded(
    CONFIG.HISTORY.TASKS_CURRENT_SHEET_NAME,
    CONFIG.HISTORY_SHEET_HEADERS,
    CONFIG.HISTORY.TASKS_BACKUP_SHEET_PREFIX
  );
}

/**
 * Rotates meetings history when rows are older than retention threshold.
 */
function rotateMeetingsHistoryIfNeeded() {
  rotateHistorySheetIfNeeded(
    CONFIG.HISTORY.MEETINGS_CURRENT_SHEET_NAME,
    CONFIG.MEETINGS.HISTORY_HEADERS,
    CONFIG.HISTORY.MEETINGS_BACKUP_SHEET_PREFIX
  );
}

/**
 * Creates the History sheet if it doesn't exist
 */
function createHistorySheet() {
  const historySpreadsheet = getOrCreateHistorySpreadsheet();
  getOrCreateHistoryDataSheet(
    historySpreadsheet,
    CONFIG.HISTORY.TASKS_CURRENT_SHEET_NAME,
    CONFIG.HISTORY_SHEET_HEADERS
  );
}

/**
 * Loads the most recent snapshot of each task from History sheet
 * @returns {Map} Map of taskId+email to most recent task state
 */
function getPreviousSnapshot() {
  const historySpreadsheet = getOrCreateHistorySpreadsheet();
  const historySheet = getOrCreateHistoryDataSheet(
    historySpreadsheet,
    CONFIG.HISTORY.TASKS_CURRENT_SHEET_NAME,
    CONFIG.HISTORY_SHEET_HEADERS
  );

  const snapshot = new Map();

  if (!historySheet || historySheet.getLastRow() <= 1) {
    // No history exists yet
    return snapshot;
  }

  // Get all history data
  const lastRow = historySheet.getLastRow();
  const data = historySheet.getRange(2, 1, lastRow - 1, CONFIG.HISTORY_SHEET_HEADERS.length).getValues();

  // Build map of most recent state for each task+email combination
  // Process rows from newest to oldest (reverse)
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const taskId = row[2]; // Task ID column
    const email = row[5]; // Email column

    // Create unique key combining taskId and email
    const key = `${taskId}|${email}`;

    // Only store the first (most recent) occurrence of each task+email combination
    if (!snapshot.has(key)) {
      snapshot.set(key, {
        taskId: taskId,
        sheetName: row[1],
        task: row[3],
        priority: row[4],
        assignees: row[5], // Will be treated as assignees field
        email: email,
        status: row[6],
        initDate: row[7],
        finishDate: row[8],
        hits: row[9],
        product: row[10],
        notes: row[11],
        action: row[12]
      });
    }
  }

  return snapshot;
}

/**
 * Saves current snapshot to History sheet
 * @param {Array} tasks - Current task list
 * @param {Object} changes - Change detection results
 */
function saveSnapshot(tasks, changes) {
  const historySpreadsheet = getOrCreateHistorySpreadsheet();
  const historySheet = getOrCreateHistoryDataSheet(
    historySpreadsheet,
    CONFIG.HISTORY.TASKS_CURRENT_SHEET_NAME,
    CONFIG.HISTORY_SHEET_HEADERS
  );

  if (!historySheet) {
    return;
  }

  const timestamp = formatDateTime(new Date());
  const rows = [];

  // Create a set of notified task keys (taskId + email)
  const notifiedKeys = new Set();
  const newKeys = new Set();
  const changedKeys = new Set();
  changes.new.forEach(t => {
    const key = `${t.taskId}|${t.email}`;
    notifiedKeys.add(key);
    newKeys.add(key);
  });
  changes.changed.forEach(t => {
    const key = `${t.task.taskId}|${t.task.email}`;
    notifiedKeys.add(key);
    changedKeys.add(key);
  });

  // Add current state of all tasks
  tasks.forEach(task => {
    const currentKey = `${task.taskId}|${task.email}`;
    let action = 'UNCHANGED';
    if (newKeys.has(currentKey)) {
      action = 'NEW';
    } else if (changedKeys.has(currentKey)) {
      action = 'CHANGED';
    }

    const notified = notifiedKeys.has(currentKey) ? 'YES' : 'NO';

    rows.push([
      timestamp,
      task.sheetName,
      task.taskId,
      task.task,
      task.priority,
      task.email,
      task.status,
      formatDate(task.initDate),
      formatDate(task.finishDate),
      task.hits,
      task.product,
      task.notes,
      action,
      notified
    ]);
  });

  // Add deleted tasks
  changes.deleted.forEach(task => {
    rows.push([
      timestamp,
      task.sheetName,
      task.taskId,
      task.task,
      task.priority,
      task.email,
      task.status,
      formatDate(task.initDate),
      formatDate(task.finishDate),
      task.hits,
      task.product,
      task.notes,
      'DELETED',
      'YES'
    ]);
  });

  if (rows.length > 0) {
    const lastRow = historySheet.getLastRow();
    historySheet.getRange(lastRow + 1, 1, rows.length, 14).setValues(rows);
  }
}

/**
 * Creates the Meetings History sheet if it doesn't exist
 */
function createMeetingsHistorySheet() {
  const historySpreadsheet = getOrCreateHistorySpreadsheet();
  getOrCreateHistoryDataSheet(
    historySpreadsheet,
    CONFIG.HISTORY.MEETINGS_CURRENT_SHEET_NAME,
    CONFIG.MEETINGS.HISTORY_HEADERS
  );
}

/**
 * Loads the most recent snapshot of each meeting from Meetings History sheet
 * @returns {Map} Map of meetingId+email to most recent meeting state
 */
function getPreviousMeetingsSnapshot() {
  const historySpreadsheet = getOrCreateHistorySpreadsheet();
  const historySheet = getOrCreateHistoryDataSheet(
    historySpreadsheet,
    CONFIG.HISTORY.MEETINGS_CURRENT_SHEET_NAME,
    CONFIG.MEETINGS.HISTORY_HEADERS
  );

  const snapshot = new Map();

  if (!historySheet || historySheet.getLastRow() <= 1) {
    // No history exists yet
    return snapshot;
  }

  // Get all history data
  const lastRow = historySheet.getLastRow();
  const data = historySheet.getRange(2, 1, lastRow - 1, CONFIG.MEETINGS.HISTORY_HEADERS.length).getValues();

  // Build map of most recent state for each meeting+email combination
  // Process rows from newest to oldest (reverse)
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const meetingId = row[1]; // Meeting ID column
    const email = row[3]; // Individual attendee email column

    // Create unique key combining meetingId and email
    const key = `${meetingId}|${email}`;

    // Only store the first (most recent) occurrence of each meeting+email combination
    if (!snapshot.has(key)) {
      // Parse the stored datetime so date/time normalization works correctly
      // against current meeting data that has properly separated date/time values.
      const parsedDatetime = parseLocalizedDateOrDateTime(row[5]);

      snapshot.set(key, {
        meetingId: meetingId,
        title: row[2],
        attendees: row[3],
        email: email,
        status: row[4],
        datetime: parsedDatetime || row[5],
        date: parsedDatetime || row[5],
        time: parsedDatetime || row[5],
        agenda: row[6],
        documentation: row[7],
        action: row[8]
      });
    }
  }

  return snapshot;
}

/**
 * Saves current meetings snapshot to Meetings History sheet
 * @param {Array} meetings - Current meeting list
 * @param {Object} changes - Change detection results
 */
function saveMeetingsSnapshot(meetings, changes) {
  const historySpreadsheet = getOrCreateHistorySpreadsheet();
  const historySheet = getOrCreateHistoryDataSheet(
    historySpreadsheet,
    CONFIG.HISTORY.MEETINGS_CURRENT_SHEET_NAME,
    CONFIG.MEETINGS.HISTORY_HEADERS
  );

  if (!historySheet) {
    return;
  }

  const timestamp = formatDateTime(new Date());
  const rows = [];
  const postponedPreviousKeys = new Set(
    changes.postponed.map(changeObj => `${changeObj.previousMeeting.meetingId}|${changeObj.previousMeeting.email}`)
  );

  // Create a set of notified meeting keys (meetingId + email)
  const notifiedKeys = new Set();
  const newKeys = new Set();
  const changedKeys = new Set();
  const postponedKeys = new Set();
  const cancelledKeys = new Set();
  changes.new.forEach(m => {
    const key = `${m.meetingId}|${m.email}`;
    notifiedKeys.add(key);
    newKeys.add(key);
  });
  changes.changed.forEach(m => {
    const key = `${m.meeting.meetingId}|${m.meeting.email}`;
    notifiedKeys.add(key);
    changedKeys.add(key);
  });
  changes.postponed.forEach(m => {
    const key = `${m.meeting.meetingId}|${m.meeting.email}`;
    notifiedKeys.add(key);
    postponedKeys.add(key);
  });
  changes.cancelled.forEach(m => {
    const key = `${m.meeting.meetingId}|${m.meeting.email}`;
    notifiedKeys.add(key);
    cancelledKeys.add(key);
  });

  // Add current state of all meetings
  meetings.forEach(meeting => {
    const currentKey = `${meeting.meetingId}|${meeting.email}`;
    let action = 'UNCHANGED';
    if (postponedKeys.has(currentKey)) {
      action = 'POSTPONED';
    } else if (cancelledKeys.has(currentKey)) {
      action = 'CANCELLED';
    } else if (newKeys.has(currentKey)) {
      action = 'NEW';
    } else if (changedKeys.has(currentKey)) {
      action = 'CHANGED';
    }

    const notified = notifiedKeys.has(currentKey) ? 'YES' : 'NO';

    rows.push([
      timestamp,
      meeting.meetingId,
      meeting.title,
      meeting.email, // Individual email for this attendee
      meeting.status,
      formatMeetingDateTime(meeting.datetime),
      meeting.agenda,
      meeting.documentation,
      action,
      notified
    ]);
  });

  // Add deleted meetings
  changes.deleted.forEach(meeting => {
    const deletedKey = `${meeting.meetingId}|${meeting.email}`;

    // Skip synthetic deleted rows generated for postponed meetings.
    // Persisting these would overwrite the latest baseline and retrigger notifications.
    if (postponedPreviousKeys.has(deletedKey)) {
      return;
    }

    rows.push([
      timestamp,
      meeting.meetingId,
      meeting.title,
      meeting.email,
      meeting.status,
      formatMeetingDateTime(meeting.datetime),
      meeting.agenda,
      meeting.documentation,
      'DELETED',
      'YES'
    ]);
  });

  if (rows.length > 0) {
    const lastRow = historySheet.getLastRow();
    historySheet.getRange(lastRow + 1, 1, rows.length, 10).setValues(rows);
  }
}



// ================================================================================
// SECTION 6: CHANGE DETECTION
// ================================================================================

/**
 * Compares current tasks with previous snapshot to detect changes
 * @param {Array} currentTasks - Current task list
 * @param {Map} previousSnapshot - Previous task states (keyed by taskId+email)
 * @returns {Object} Categorized changes: {new, changed, deleted, unchanged}
 */
function compareSnapshots(currentTasks, previousSnapshot) {
  const changes = {
    new: [],
    changed: [],
    deleted: [],
    unchanged: []
  };

  const currentKeys = new Set();
  const matchedPreviousKeys = new Set();

  // Candidate lookup by email for fallback matching when row numbers changed.
  const previousByEmail = new Map();
  previousSnapshot.forEach((previousTask, key) => {
    if (previousTask.action === 'DELETED') {
      return;
    }

    if (!previousByEmail.has(previousTask.email)) {
      previousByEmail.set(previousTask.email, []);
    }

    previousByEmail.get(previousTask.email).push({ key, task: previousTask, used: false });
  });

  // Check each current task
  currentTasks.forEach(currentTask => {
    const directKey = `${currentTask.taskId}|${currentTask.email}`;
    let resolvedKey = directKey;
    let previousTask = previousSnapshot.get(directKey);

    // Fallback reconciliation: same assignee + same content means row-move, not new task.
    if (!previousTask) {
      const candidates = previousByEmail.get(currentTask.email) || [];
      const currentFingerprint = buildTaskFingerprint(currentTask);
      const candidate = candidates.find(entry => !entry.used && buildTaskFingerprint(entry.task) === currentFingerprint);

      if (candidate) {
        previousTask = candidate.task;
        candidate.used = true;
        resolvedKey = candidate.key;
        currentTask.taskId = previousTask.taskId;
      }
    }

    currentKeys.add(resolvedKey);

    if (!previousTask) {
      changes.new.push(currentTask);
      return;
    }

    matchedPreviousKeys.add(resolvedKey);
    const fieldChanges = detectFieldChanges(currentTask, previousTask);

    if (fieldChanges.length > 0) {
      changes.changed.push({
        task: currentTask,
        changes: fieldChanges,
        previousTask: previousTask
      });
    } else {
      changes.unchanged.push(currentTask);
    }
  });

  // Check for deleted tasks (or removed assignees)
  previousSnapshot.forEach((previousTask, key) => {
    if (!currentKeys.has(key) && !matchedPreviousKeys.has(key) && previousTask.action !== 'DELETED') {
      changes.deleted.push(previousTask);
    }
  });

  return changes;
}

/**
 * Detects which specific fields changed between two task versions
 * @param {Object} currentTask - Current task state
 * @param {Object} previousTask - Previous task state
 * @returns {Array} Array of change objects: [{field, oldValue, newValue}]
 */
function detectFieldChanges(currentTask, previousTask) {
  const fieldChanges = [];

  CONFIG.MONITORED_FIELDS.forEach(fieldKey => {
    let currentValue, previousValue;

    // Map field keys to task object properties
    switch (fieldKey) {
      case 'TASK':
        currentValue = currentTask.task;
        previousValue = previousTask.task;
        break;
      case 'PRIORITY':
        currentValue = currentTask.priority;
        previousValue = previousTask.priority;
        break;
      case 'EMAIL':
        currentValue = currentTask.email;
        previousValue = previousTask.email;
        break;
      case 'STATUS':
        currentValue = currentTask.status;
        previousValue = previousTask.status;
        break;
      case 'INIT_DATE':
        currentValue = formatDate(currentTask.initDate);
        previousValue = formatDate(previousTask.initDate);
        break;
      case 'FINISH_DATE':
        currentValue = formatDate(currentTask.finishDate);
        previousValue = formatDate(previousTask.finishDate);
        break;
      case 'PRODUCT':
        currentValue = currentTask.product;
        previousValue = previousTask.product;
        break;
      case 'NOTES':
        currentValue = currentTask.notes;
        previousValue = previousTask.notes;
        break;
    }

    // Normalize for comparison
    let currentNorm;
    let previousNorm;

    if (fieldKey === 'INIT_DATE' || fieldKey === 'FINISH_DATE') {
      currentNorm = normalizeDateForComparison(currentTask[fieldKey === 'INIT_DATE' ? 'initDate' : 'finishDate']);
      previousNorm = normalizeDateForComparison(previousTask[fieldKey === 'INIT_DATE' ? 'initDate' : 'finishDate']);
    } else {
      currentNorm = normalizeStringForComparison(currentValue);
      previousNorm = normalizeStringForComparison(previousValue);
    }

    if (currentNorm !== previousNorm) {
      fieldChanges.push({
        field: fieldKey,
        fieldName: CONFIG.FIELD_NAMES[fieldKey],
        oldValue: previousValue || '(empty)',
        newValue: currentValue || '(empty)'
      });
    }
  });

  return fieldChanges;
}

/**
 * Generates unique task ID from sheet name and row number
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowIndex - Row number
 * @returns {string} Unique task ID
 */
function generateTaskId(sheetName, rowIndex) {
  return `${sheetName}_Row${rowIndex}`;
}

/**
 * Compares current meetings with previous snapshot to detect changes
 * @param {Array} currentMeetings - Current meeting list
 * @param {Map} previousSnapshot - Previous meeting states (keyed by meetingId+email)
 * @returns {Object} Categorized changes: {new, changed, deleted, unchanged}
 */
function compareMeetingsSnapshots(currentMeetings, previousSnapshot) {
  const changes = {
    new: [],
    changed: [],
    deleted: [],
    unchanged: [],
    postponed: [],
    cancelled: []
  };

  const currentKeys = new Set();
  const matchedPreviousKeys = new Set();

  // Candidate lookup by attendee for fallback matching when row numbers changed.
  const previousByEmail = new Map();
  previousSnapshot.forEach((previousMeeting, key) => {
    if (previousMeeting.action === 'DELETED') {
      return;
    }

    if (!previousByEmail.has(previousMeeting.email)) {
      previousByEmail.set(previousMeeting.email, []);
    }

    previousByEmail.get(previousMeeting.email).push({ key, meeting: previousMeeting, used: false });
  });

  // Check each current meeting
  currentMeetings.forEach(currentMeeting => {
    const directKey = `${currentMeeting.meetingId}|${currentMeeting.email}`;
    let resolvedKey = directKey;
    let previousMeeting = previousSnapshot.get(directKey);

    // Fallback reconciliation: same attendee + same content means row-move, not new meeting.
    if (!previousMeeting) {
      const candidates = previousByEmail.get(currentMeeting.email) || [];
      const currentFingerprint = buildMeetingFingerprint(currentMeeting);
      const candidate = candidates.find(entry => !entry.used && buildMeetingFingerprint(entry.meeting) === currentFingerprint);

      if (candidate) {
        previousMeeting = candidate.meeting;
        candidate.used = true;
        resolvedKey = candidate.key;
        currentMeeting.meetingId = previousMeeting.meetingId;
      }
    }

    currentKeys.add(resolvedKey);

    if (!previousMeeting) {
      if (isCancelledMeetingStatus(currentMeeting.status)) {
        changes.cancelled.push({
          meeting: currentMeeting,
          changes: [],
          previousMeeting: currentMeeting
        });
      } else {
        changes.new.push(currentMeeting);
      }
      return;
    }

    matchedPreviousKeys.add(resolvedKey);

    // Already-notified cancelled meetings stay tracked but should not notify again.
    if (previousMeeting.action === 'CANCELLED' && isCancelledMeetingStatus(currentMeeting.status)) {
      changes.unchanged.push(currentMeeting);
      return;
    }

    const fieldChanges = detectMeetingFieldChanges(currentMeeting, previousMeeting);

    if (fieldChanges.length > 0) {
      if (isPostponedMeetingReschedule(currentMeeting, fieldChanges)) {
        changes.postponed.push({
          meeting: currentMeeting,
          changes: fieldChanges,
          previousMeeting: previousMeeting
        });

        // Notify postponed meetings as delete + create actions.
        changes.new.push(currentMeeting);
        changes.deleted.push(previousMeeting);
      } else if (isCancelledMeetingUpdate(currentMeeting, previousMeeting)) {
        changes.cancelled.push({
          meeting: currentMeeting,
          changes: fieldChanges,
          previousMeeting: previousMeeting
        });
      } else {
        changes.changed.push({
          meeting: currentMeeting,
          changes: fieldChanges,
          previousMeeting: previousMeeting
        });
      }
    } else {
      changes.unchanged.push(currentMeeting);
    }
  });

  // Check for deleted meetings (or removed attendees)
  previousSnapshot.forEach((previousMeeting, key) => {
    if (!currentKeys.has(key) && !matchedPreviousKeys.has(key) && previousMeeting.action !== 'DELETED') {
      changes.deleted.push(previousMeeting);
    }
  });

  return changes;
}

/**
 * Detects which specific fields changed between two meeting versions
 * @param {Object} currentMeeting - Current meeting state
 * @param {Object} previousMeeting - Previous meeting state
 * @returns {Array} Array of change objects: [{field, oldValue, newValue}]
 */
function detectMeetingFieldChanges(currentMeeting, previousMeeting) {
  const fieldChanges = [];

  CONFIG.MEETINGS.MONITORED_FIELDS.forEach(fieldKey => {
    let currentValue, previousValue;

    // Map field keys to meeting object properties
    switch (fieldKey) {
      case 'TITLE':
        currentValue = currentMeeting.title;
        previousValue = previousMeeting.title;
        break;
      case 'ATTENDEES': {
        currentValue = currentMeeting.attendees;
        previousValue = previousMeeting.attendees;

        // Ignore list churn for this attendee when they remain invited in both versions.
        const currentEmails = parseMultipleEmails(currentMeeting.attendees);
        const previousEmails = parseMultipleEmails(previousMeeting.attendees);
        const attendeeEmail = currentMeeting.email || previousMeeting.email;

        if (
          attendeeEmail &&
          currentEmails.includes(attendeeEmail) &&
          previousEmails.includes(attendeeEmail)
        ) {
          return;
        }
        break;
      }
      case 'STATUS':
        currentValue = currentMeeting.status;
        previousValue = previousMeeting.status;
        break;
      case 'DATE':
        currentValue = formatDate(currentMeeting.date || currentMeeting.datetime);
        previousValue = formatDate(previousMeeting.date || previousMeeting.datetime);
        break;
      case 'TIME':
        currentValue = formatTime(currentMeeting.time || currentMeeting.datetime);
        previousValue = formatTime(previousMeeting.time || previousMeeting.datetime);
        break;
      case 'AGENDA':
        currentValue = currentMeeting.agenda;
        previousValue = previousMeeting.agenda;
        break;
      case 'DOCUMENTATION':
        currentValue = currentMeeting.documentation;
        previousValue = previousMeeting.documentation;
        break;
    }

    // Normalize for comparison
    const currentNorm = fieldKey === 'DATE'
      ? normalizeDateForComparison(currentMeeting.date || currentMeeting.datetime)
      : fieldKey === 'TIME'
        ? normalizeTimeForComparison(currentMeeting.time || currentMeeting.datetime)
        : normalizeStringForComparison(currentValue);

    const previousNorm = fieldKey === 'DATE'
      ? normalizeDateForComparison(previousMeeting.date || previousMeeting.datetime)
      : fieldKey === 'TIME'
        ? normalizeTimeForComparison(previousMeeting.time || previousMeeting.datetime)
        : normalizeStringForComparison(previousValue);

    if (currentNorm !== previousNorm) {
      fieldChanges.push({
        field: fieldKey,
        fieldName: CONFIG.MEETINGS.FIELD_NAMES[fieldKey],
        oldValue: previousValue || '(empty)',
        newValue: currentValue || '(empty)'
      });
    }
  });

  return fieldChanges;
}

/**
 * Detects postponed meetings that should be handled as delete/create notifications.
 * Rules:
 * - Current status must be "Pospuesta"
 * - At least date or time changed
 * - No other fields changed except optional status update
 * @param {Object} currentMeeting - Current meeting state
 * @param {Array} fieldChanges - Detected field-level changes
 * @returns {boolean} True when meeting should be treated as postponed
 */
function isPostponedMeetingReschedule(currentMeeting, fieldChanges) {
  if (!currentMeeting || !fieldChanges || fieldChanges.length === 0) {
    return false;
  }

  const normalizedStatus = removeDiacritics(
    normalizeStringForComparison(currentMeeting.status).toLowerCase()
  );

  if (normalizedStatus !== 'pospuesta') {
    return false;
  }

  const changedFields = fieldChanges.map(change => change.field);
  const hasDateOrTimeChange = changedFields.includes('DATE') || changedFields.includes('TIME');

  if (!hasDateOrTimeChange) {
    return false;
  }

  const allowedFields = new Set(['DATE', 'TIME', 'STATUS']);
  return changedFields.every(field => allowedFields.has(field));
}

/**
 * Checks whether a meeting status represents a cancelled meeting.
 * @param {string} status - Meeting status text
 * @returns {boolean} True when status is cancelled-like
 */
function isCancelledMeetingStatus(status) {
  const normalizedStatus = removeDiacritics(
    normalizeStringForComparison(status).toLowerCase()
  );

  return normalizedStatus === 'cancelada'
    || normalizedStatus === 'cancelado'
    || normalizedStatus === 'cancelled'
    || normalizedStatus === 'canceled';
}

/**
 * Detects cancelled meetings that should be notified in a dedicated category.
 * @param {Object} currentMeeting - Current meeting state
 * @param {Object} previousMeeting - Previous meeting state
 * @returns {boolean} True when meeting should be treated as cancelled
 */
function isCancelledMeetingUpdate(currentMeeting, previousMeeting) {
  if (!currentMeeting || !previousMeeting) {
    return false;
  }

  if (!isCancelledMeetingStatus(currentMeeting.status)) {
    return false;
  }

  // Notify as cancelled when transitioning from any non-cancelled state,
  // or when no prior cancelled action exists in history yet.
  return !isCancelledMeetingStatus(previousMeeting.status)
    || previousMeeting.action !== 'CANCELLED';
}



// ================================================================================
// SECTION 7: EMAIL NOTIFICATION SYSTEM
// ================================================================================

/**
 * Sends consolidated email notifications to assignees with changed tasks
 * @param {Object} changes - Detected changes
 * @param {string} spreadsheetUrl - URL of the spreadsheet
 * @returns {Object} Notification statistics
 */
function sendChangeNotifications(changes, spreadsheetUrl) {
  const tasksByAssignee = groupTasksByAssignee(changes);
  let emailsSent = 0;

  tasksByAssignee.forEach((tasks, email) => {
    try {
      if (!isValidEmail(email)) {
        Logger.log(`Skipping invalid email: ${email}`);
        return;
      }

      const emailBody = buildEmailBody(
        email,
        tasks.new,
        tasks.changed,
        tasks.deleted,
        spreadsheetUrl
      );

      const emailHtmlBody = buildTaskEmailHtml(
        email,
        tasks.new,
        tasks.changed,
        tasks.deleted,
        spreadsheetUrl
      );

      MailApp.sendEmail({
        to: email,
        subject: CONFIG.EMAIL_SUBJECT,
        body: emailBody,
        htmlBody: emailHtmlBody
      });

      emailsSent++;
      Logger.log(`Email sent to: ${email}`);

    } catch (error) {
      Logger.log(`Failed to send email to ${email}: ${error.toString()}`);
      showUiAlertSafe(CONFIG.ALERTS.EMAIL_NOTIFICATION_FAILED(email));
    }
  });

  return { emailsSent: emailsSent };
}

/**
 * Groups tasks by assignee email address
 * @param {Object} changes - Detected changes
 * @returns {Map} Map of email to categorized tasks
 */
function groupTasksByAssignee(changes) {
  const tasksByAssignee = new Map();

  // Helper function to add task to assignee's list
  const addToAssignee = (email, category, task) => {
    if (!tasksByAssignee.has(email)) {
      tasksByAssignee.set(email, { new: [], changed: [], deleted: [] });
    }
    tasksByAssignee.get(email)[category].push(task);
  };

  // Add new tasks
  changes.new.forEach(task => {
    addToAssignee(task.email, 'new', task);
  });

  // Add changed tasks
  changes.changed.forEach(changeObj => {
    addToAssignee(changeObj.task.email, 'changed', changeObj);
  });

  // Add deleted tasks (use previous assignee)
  changes.deleted.forEach(task => {
    addToAssignee(task.email, 'deleted', task);
  });

  return tasksByAssignee;
}

/**
 * Builds plain text email body with task change details
 * @param {string} assigneeEmail - Recipient email
 * @param {Array} newTasks - New tasks
 * @param {Array} changedTasks - Changed tasks with change details
 * @param {Array} deletedTasks - Deleted tasks
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @returns {string} Formatted email body
 */
function buildEmailBody(assigneeEmail, newTasks, changedTasks, deletedTasks, spreadsheetUrl) {
  const totalChanges = newTasks.length + changedTasks.length + deletedTasks.length;
  let body = `Hola,\n\n`;
  body += `Tienes ${totalChanges} tarea(s) que fueron actualizadas recientemente en el tablero de gestión de proyectos:\n\n`;
  body += `═══════════════════════════════════════════════════════\n\n`;

  // Section 1: New tasks
  if (newTasks.length > 0) {
    newTasks.forEach(task => {
      body += `✨ NUEVA TAREA ASIGNADA A TI en la hoja ${task.sheetName}\n\n`;
      body += `Tarea: ${task.task}\n`;
      body += `  Prioridad: ${task.priority}\n`;
      body += `  Estado: ${task.status}\n`;
      body += `  Producto: ${task.product}\n`;
      body += `  Fecha de inicio: ${formatDate(task.initDate)}\n`;
      body += `  Fecha de finalización: ${formatDate(task.finishDate)}\n`;
      if (task.notes) {
        body += `  Notas: ${task.notes}\n`;
      }
      // Show all assignees if multiple
      if (task.assignees) {
        body += `  Asignado a: ${task.assignees}\n`;
        const assigneeCount = parseMultipleEmails(task.assignees).length;
        if (assigneeCount > 1) {
          body += `  → Esta tarea fue asignada a ti y ${assigneeCount - 1} colaborador(es) más\n`;
        } else {
          body += `  → Esta tarea fue asignada recientemente a ti\n`;
        }
      } else {
        body += `  → Esta tarea fue asignada recientemente a ti\n`;
      }
      body += `\n`;
      body += `───────────────────────────────────────────────────────\n\n`;
    });
  }

  // Section 2: Changed tasks
  if (changedTasks.length > 0) {
    changedTasks.forEach(changeObj => {
      const task = changeObj.task;
      const changes = changeObj.changes;

      body += `📝 TAREA ACTUALIZADA en la hoja ${task.sheetName}\n\n`;
      body += `Tarea: ${task.task}\n`;

      // Show all fields, highlighting changed ones
      changes.forEach(change => {
        if (change.field === 'TASK') {
          // Task name changed - already shown above
          body += `  (Nombre de la tarea cambiado de: "${change.oldValue}")\n`;
        } else if (change.field === 'PRIORITY') {
          body += `  Prioridad: ${change.oldValue} → ${change.newValue} ⚠️\n`;
        } else if (change.field === 'STATUS') {
          body += `  Estado: ${change.oldValue} → ${change.newValue} ✓\n`;
        } else if (change.field === 'INIT_DATE') {
          body += `  Fecha de inicio: ${change.oldValue} → ${change.newValue}\n`;
        } else if (change.field === 'FINISH_DATE') {
          body += `  Fecha de finalización: ${change.oldValue} → ${change.newValue}\n`;
        } else if (change.field === 'PRODUCT') {
          body += `  Producto: ${change.oldValue} → ${change.newValue}\n`;
        } else if (change.field === 'EMAIL') {
          body += `  Asignado a: ${change.oldValue} → ${change.newValue}\n`;
        } else if (change.field === 'NOTES') {
          body += `  Notas: ${change.oldValue} → ${change.newValue}\n`;
        }
      });

      // Show unchanged fields
      if (!changes.find(c => c.field === 'PRIORITY')) {
        body += `  Prioridad: ${task.priority}\n`;
      }
      if (!changes.find(c => c.field === 'STATUS')) {
        body += `  Estado: ${task.status}\n`;
      }
      if (!changes.find(c => c.field === 'PRODUCT')) {
        body += `  Producto: ${task.product}\n`;
      }
      if (!changes.find(c => c.field === 'INIT_DATE')) {
        body += `  Fecha de inicio: ${formatDate(task.initDate)}\n`;
      }
      if (!changes.find(c => c.field === 'FINISH_DATE')) {
        body += `  Fecha de finalización: ${formatDate(task.finishDate)}\n`;
      }
      if (!changes.find(c => c.field === 'NOTES') && task.notes) {
        body += `  Notas: ${task.notes}\n`;
      }
      // Show all assignees
      if (!changes.find(c => c.field === 'EMAIL') && task.assignees) {
        body += `  Asignado a: ${task.assignees}\n`;
      }

      body += `  \n`;
      body += `  → ${changes.length} cambios desde la última notificación\n\n`;
      body += `───────────────────────────────────────────────────────\n\n`;
    });
  }

  // Section 3: Deleted tasks
  if (deletedTasks.length > 0) {
    deletedTasks.forEach(task => {
      body += `🗑️ TAREA ELIMINADA en la hoja ${task.sheetName}\n\n`;
      body += `Tarea: ${task.task}\n`;
      body += `  → Esta tarea fue eliminada de tus asignaciones\n\n`;
      body += `───────────────────────────────────────────────────────\n\n`;
    });
  }

  body += `═══════════════════════════════════════════════════════\n\n`;
  body += `Ver el tablero completo: ${spreadsheetUrl}\n\n`;
  body += `Notificación enviada: ${formatDateTime(new Date())}\n\n`;
  body += `---\n`;
  body += `ℹ️ Estás recibiendo esto porque las tareas asignadas a ti han cambiado.\n`;

  return body;
}

/**
 * Builds HTML email body with task change details.
 * @param {string} assigneeEmail - Recipient email
 * @param {Array} newTasks - New tasks
 * @param {Array} changedTasks - Changed tasks with change details
 * @param {Array} deletedTasks - Deleted tasks
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @returns {string} Formatted HTML body
 */
function buildTaskEmailHtml(assigneeEmail, newTasks, changedTasks, deletedTasks, spreadsheetUrl) {
  const totalChanges = newTasks.length + changedTasks.length + deletedTasks.length;
  let html = '';

  html += '<div style="margin:0;padding:0;background:#f5f7fb;">';
  html += '<div style="max-width:700px;margin:0 auto;padding:24px 14px;font-family:\'Segoe UI\', Arial, sans-serif;color:#1f2937;">';
  html += '<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:14px;padding:20px;">';
  html += `<h2 style="margin:0 0 10px 0;font-size:22px;line-height:1.2;color:#111827;">Tareas actualizadas</h2>`;
  html += `<p style="margin:0 0 16px 0;font-size:14px;color:#4b5563;">Tienes <strong>${totalChanges}</strong> tarea(s) actualizadas recientemente.</p>`;

  if (newTasks.length > 0) {
    newTasks.forEach(task => {
      const assigneeCount = parseMultipleEmails(task.assignees || '').length;
      const footerText = assigneeCount > 1
        ? `Esta tarea fue asignada a ti y ${assigneeCount - 1} colaborador(es) más.`
        : 'Esta tarea fue asignada recientemente a ti.';
      html += buildTaskCardHtml('✨ NUEVA TAREA ASIGNADA', '#e8f7ee', '#2f855a', task, null, footerText);
    });
  }

  if (changedTasks.length > 0) {
    changedTasks.forEach(changeObj => {
      const task = changeObj.task;
      const changesHtml = buildTaskChangesHtml(changeObj.changes);
      html += buildTaskCardHtml('📝 TAREA ACTUALIZADA', '#fff8e8', '#9a6700', task, changesHtml, `${changeObj.changes.length} cambio(s) desde la última notificación.`);
    });
  }

  if (deletedTasks.length > 0) {
    deletedTasks.forEach(task => {
      html += buildTaskCardHtml('🗑️ TAREA ELIMINADA', '#fdecec', '#b42318', task, null, 'Esta tarea fue eliminada de tus asignaciones.', {
        showPriority: false,
        showStatus: false,
        showProduct: false,
        showInitDate: false,
        showFinishDate: false,
        showAssignees: false,
        showNotes: false
      });
    });
  }

  html += '<div style="margin-top:18px;padding-top:14px;border-top:1px solid #e5e7eb;">';
  html += `<p style="margin:0 0 8px 0;font-size:13px;"><a href="${escapeHtml(spreadsheetUrl)}" target="_blank" style="color:#1d4ed8;text-decoration:none;">Ver el tablero completo</a></p>`;
  html += `<p style="margin:0;font-size:12px;color:#6b7280;">Notificación enviada: ${escapeHtml(formatDateTime(new Date()))}</p>`;
  html += '<p style="margin:10px 0 0 0;font-size:12px;color:#6b7280;">Estás recibiendo esto porque las tareas asignadas a ti han cambiado.</p>';
  html += '</div>';

  html += '</div>';
  html += '</div>';
  html += '</div>';

  return html;
}

/**
 * Builds an HTML card for a task item.
 * @param {string} title - Card title label
 * @param {string} badgeBackground - Badge background color
 * @param {string} badgeColor - Badge text color
 * @param {Object} task - Task object
 * @param {string|null} changesHtml - Changes block HTML
 * @param {string} footerText - Card footer text
 * @param {Object=} displayOptions - Optional field visibility flags
 * @returns {string} HTML card
 */
function buildTaskCardHtml(title, badgeBackground, badgeColor, task, changesHtml, footerText, displayOptions) {
  let card = '';
  const resolvedDisplayOptions = Object.assign({
    showPriority: true,
    showStatus: true,
    showProduct: true,
    showInitDate: true,
    showFinishDate: true,
    showAssignees: true,
    showNotes: true
  }, displayOptions || {});

  card += '<div style="margin:0 0 14px 0;padding:16px;border:1px solid #e5e7eb;border-radius:12px;background:#ffffff;">';
  card += `<div style="display:inline-block;font-size:12px;font-weight:700;padding:4px 10px;border-radius:999px;background:${badgeBackground};color:${badgeColor};margin-bottom:10px;">${escapeHtml(title)}</div>`;
  card += `<h3 style="margin:0 0 10px 0;font-size:18px;color:#111827;">${escapeHtml(task.task || '(sin nombre)')}</h3>`;

  if (task.sheetName) {
    card += `<p style="margin:0 0 4px 0;font-size:14px;"><strong>Hoja:</strong> ${escapeHtml(task.sheetName)}</p>`;
  }

  if (resolvedDisplayOptions.showPriority && task.priority) {
    card += `<p style="margin:0 0 4px 0;font-size:14px;"><strong>Prioridad:</strong> ${escapeHtml(task.priority)}</p>`;
  }

  if (resolvedDisplayOptions.showStatus && task.status) {
    card += `<p style="margin:0 0 4px 0;font-size:14px;"><strong>Estado:</strong> ${escapeHtml(task.status)}</p>`;
  }

  if (resolvedDisplayOptions.showProduct && task.product) {
    card += `<p style="margin:0 0 4px 0;font-size:14px;"><strong>Producto:</strong> ${escapeHtml(task.product)}</p>`;
  }

  const initDate = formatDate(task.initDate);
  if (resolvedDisplayOptions.showInitDate && initDate) {
    card += `<p style="margin:0 0 4px 0;font-size:14px;"><strong>Fecha de inicio:</strong> ${escapeHtml(initDate)}</p>`;
  }

  const finishDate = formatDate(task.finishDate);
  if (resolvedDisplayOptions.showFinishDate && finishDate) {
    card += `<p style="margin:0 0 10px 0;font-size:14px;"><strong>Fecha de finalización:</strong> ${escapeHtml(finishDate)}</p>`;
  }

  if (resolvedDisplayOptions.showAssignees && task.assignees) {
    card += `<p style="margin:0 0 8px 0;font-size:13px;"><strong>Asignado a:</strong> ${escapeHtml(task.assignees)}</p>`;
  }

  if (changesHtml) {
    card += changesHtml;
  }

  if (resolvedDisplayOptions.showNotes && task.notes) {
    card += '<p style="margin:10px 0 4px 0;font-size:13px;font-weight:700;color:#374151;">Notas</p>';
    card += `<p style="margin:0;font-size:13px;color:#4b5563;line-height:1.45;">${textToHtml(task.notes)}</p>`;
  }

  card += `<p style="margin:10px 0 0 0;font-size:12px;color:#6b7280;">${escapeHtml(footerText)}</p>`;
  card += '</div>';

  return card;
}

/**
 * Builds HTML list of task changes.
 * @param {Array} changes - Changed fields
 * @returns {string} HTML block
 */
function buildTaskChangesHtml(changes) {
  if (!changes || changes.length === 0) {
    return '';
  }

  let html = '';
  html += '<div style="margin-top:8px;padding:10px;background:#fffbeb;border:1px solid #f3e8bb;border-radius:8px;">';
  html += '<p style="margin:0 0 6px 0;font-size:13px;font-weight:700;color:#7c5a00;">Cambios detectados</p>';
  html += '<ul style="margin:0;padding-left:18px;font-size:13px;color:#4b5563;">';

  changes.forEach(change => {
    html += `<li style="margin:0 0 4px 0;">${escapeHtml(formatTaskChangeText(change))}</li>`;
  });

  html += '</ul>';
  html += '</div>';
  return html;
}

/**
 * Formats one changed task field for display.
 * @param {Object} change - Change object
 * @returns {string} Human-readable change text
 */
function formatTaskChangeText(change) {
  if (change.field === 'TASK') {
    return `Nombre de tarea: "${change.oldValue}" -> "${change.newValue}"`;
  }

  if (change.field === 'PRIORITY') {
    return `Prioridad: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'STATUS') {
    return `Estado: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'INIT_DATE') {
    return `Fecha de inicio: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'FINISH_DATE') {
    return `Fecha de finalización: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'PRODUCT') {
    return `Producto: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'EMAIL') {
    return `Asignado a: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'NOTES') {
    return 'Notas actualizadas';
  }

  return `${change.field}: ${change.oldValue} -> ${change.newValue}`;
}

/**
 * Sends consolidated email notifications to attendees with changed meetings
 * @param {Object} changes - Detected changes
 * @param {string} spreadsheetUrl - URL of the spreadsheet
 * @returns {Object} Notification statistics
 */
function sendMeetingNotifications(changes, spreadsheetUrl) {
  const meetingsByAttendee = groupMeetingsByAttendee(changes);
  let emailsSent = 0;
  const ownerEmail = resolveNotificationSenderEmail();

  meetingsByAttendee.forEach((meetings, email) => {
    try {
      if (!isValidEmail(email)) {
        Logger.log(`Skipping invalid meeting email: ${email}`);
        return;
      }

      const emailBody = buildMeetingEmailBody(
        email,
        meetings.new,
        meetings.changed,
        meetings.deleted,
        meetings.postponed,
        meetings.cancelled,
        spreadsheetUrl,
        ownerEmail
      );

      const emailHtmlBody = buildMeetingEmailHtml(
        email,
        meetings.new,
        meetings.changed,
        meetings.deleted,
        meetings.postponed,
        meetings.cancelled,
        spreadsheetUrl,
        ownerEmail
      );

      const subject = buildMeetingEmailSubject(meetings);

      MailApp.sendEmail({
        to: email,
        subject: subject,
        body: emailBody,
        htmlBody: emailHtmlBody
      });

      emailsSent++;
      Logger.log(`Meeting notification email sent to: ${email}`);

    } catch (error) {
      Logger.log(`Failed to send meeting email to ${email}: ${error.toString()}`);
      showUiAlertSafe(CONFIG.ALERTS.EMAIL_NOTIFICATION_FAILED(email));
    }
  });

  return { emailsSent: emailsSent };
}

/**
 * Builds a contextual subject line for meeting notifications.
 * Marks emails as postponed when they contain only postponed meetings.
 * @param {Object} meetingsByType - Recipient meetings grouped by type
 * @returns {string} Subject text
 */
function buildMeetingEmailSubject(meetingsByType) {
  const newCount = (meetingsByType.new || []).length;
  const changedCount = (meetingsByType.changed || []).length;
  const deletedCount = (meetingsByType.deleted || []).length;
  const postponedCount = (meetingsByType.postponed || []).length;
  const cancelledCount = (meetingsByType.cancelled || []).length;

  if (cancelledCount > 0 && newCount === 0 && changedCount === 0 && deletedCount === 0 && postponedCount === 0) {
    return `${CONFIG.MEETINGS.EMAIL_SUBJECT} - Cancelada`;
  }

  if (postponedCount > 0 && newCount === 0 && changedCount === 0 && deletedCount === 0 && cancelledCount === 0) {
    return `${CONFIG.MEETINGS.EMAIL_SUBJECT} - Pospuesta`;
  }

  return CONFIG.MEETINGS.EMAIL_SUBJECT;
}

/**
 * Groups meetings by attendee email address
 * Filters out meetings with status "Completada" (finished meetings don't need notifications)
 * @param {Object} changes - Detected changes
 * @returns {Map} Map of email to categorized meetings
 */
function groupMeetingsByAttendee(changes) {
  const meetingsByAttendee = new Map();

  const postponedCurrentKeys = new Set(
    changes.postponed.map(changeObj => `${changeObj.meeting.meetingId}|${changeObj.meeting.email}`)
  );
  const postponedPreviousKeys = new Set(
    changes.postponed.map(changeObj => `${changeObj.previousMeeting.meetingId}|${changeObj.previousMeeting.email}`)
  );

  // Helper function to add meeting to attendee's list
  const addToAttendee = (email, category, meeting) => {
    if (!meetingsByAttendee.has(email)) {
      meetingsByAttendee.set(email, { new: [], changed: [], deleted: [], postponed: [], cancelled: [] });
    }
    meetingsByAttendee.get(email)[category].push(meeting);
  };

  // Add new meetings (skip completed ones)
  changes.new.forEach(meeting => {
    const meetingKey = `${meeting.meetingId}|${meeting.email}`;
    if (postponedCurrentKeys.has(meetingKey)) {
      return;
    }

    if (meeting.status !== CONFIG.MEETINGS.COMPLETED_STATUS) {
      addToAttendee(meeting.email, 'new', meeting);
    }
  });

  // Add changed meetings (skip completed ones)
  changes.changed.forEach(changeObj => {
    if (changeObj.meeting.status !== CONFIG.MEETINGS.COMPLETED_STATUS) {
      addToAttendee(changeObj.meeting.email, 'changed', changeObj);
    }
  });

  // Add deleted meetings (use previous attendee)
  changes.deleted.forEach(meeting => {
    const meetingKey = `${meeting.meetingId}|${meeting.email}`;
    if (postponedPreviousKeys.has(meetingKey)) {
      return;
    }

    addToAttendee(meeting.email, 'deleted', meeting);
  });

  // Add postponed meetings as their own notification type (skip completed status)
  changes.postponed.forEach(changeObj => {
    if (changeObj.meeting.status !== CONFIG.MEETINGS.COMPLETED_STATUS) {
      addToAttendee(changeObj.meeting.email, 'postponed', changeObj);
    }
  });

  // Add cancelled meetings as their own notification type.
  changes.cancelled.forEach(changeObj => {
    addToAttendee(changeObj.meeting.email, 'cancelled', changeObj);
  });

  return meetingsByAttendee;
}

/**
 * Builds plain text email body with meeting change details
 * @param {string} attendeeEmail - Recipient email
 * @param {Array} newMeetings - New meetings
 * @param {Array} changedMeetings - Changed meetings with change details
 * @param {Array} deletedMeetings - Deleted meetings
 * @param {Array} postponedMeetings - Postponed meetings with change details
 * @param {Array} cancelledMeetings - Cancelled meetings with change details
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @param {string} ownerEmail - Notification sender email
 * @returns {string} Formatted email body
 */
function buildMeetingEmailBody(attendeeEmail, newMeetings, changedMeetings, deletedMeetings, postponedMeetings, cancelledMeetings, spreadsheetUrl, ownerEmail) {
  const totalChanges = newMeetings.length + changedMeetings.length + deletedMeetings.length + postponedMeetings.length + cancelledMeetings.length;
  let body = `Hola,\n\n`;
  body += `Tienes ${totalChanges} reunión(es) programadas o actualizadas:\n\n`;
  body += `═══════════════════════════════════════════════════════\n\n`;

  // Section 1: New meetings
  if (newMeetings.length > 0) {
    newMeetings.forEach(meeting => {
      body += `✨ NUEVA REUNIÓN PROGRAMADA\n\n`;
      body += `📅 ${meeting.title}\n`;
      body += `  Fecha y hora: ${formatMeetingDateTime(meeting.datetime)}\n`;
      body += `  Estado: ${meeting.status}\n`;
      body += `  Asistentes: ${meeting.attendees}\n`;
      body += `\n`;
      if (meeting.agenda) {
        body += `  Agenda:\n`;
        body += `  ${meeting.agenda}\n`;
        body += `\n`;
      }
      if (meeting.documentation) {
        body += `  📎 Documentación:\n`;
        body += `  ${meeting.documentation}\n`;
        body += `  (Asegúrate de tener acceso a los documentos compartidos)\n`;
        body += `\n`;
      }

      body += buildMeetingCalendarActionsText(meeting, spreadsheetUrl, {
        includeAdd: true,
        includeUpdate: true,
        includeDelete: true,
        includeMessageAll: true
      }, attendeeEmail, ownerEmail);

      body += `  → Esta reunión fue agendada recientemente\n\n`;
      body += `───────────────────────────────────────────────────────\n\n`;
    });
  }

  // Section 2: Changed meetings
  if (changedMeetings.length > 0) {
    changedMeetings.forEach(changeObj => {
      const meeting = changeObj.meeting;
      const changes = changeObj.changes;

      body += `📝 REUNIÓN ACTUALIZADA\n\n`;
      body += `📅 ${meeting.title}\n`;

      // Show all fields, highlighting changed ones
      changes.forEach(change => {
        if (change.field === 'TITLE') {
          body += `  (Título cambiado de: "${change.oldValue}")\n`;
        } else if (change.field === 'DATE') {
          body += `  Fecha: ${change.oldValue} → ${change.newValue} ⚠️\n`;
        } else if (change.field === 'TIME') {
          body += `  Hora: ${change.oldValue} → ${change.newValue} ⚠️\n`;
        } else if (change.field === 'STATUS') {
          body += `  Estado: ${change.oldValue} → ${change.newValue} ✓\n`;
        } else if (change.field === 'ATTENDEES') {
          body += `  Asistentes actualizados: ${change.newValue}\n`;// ${change.oldValue} → ${change.newValue}\n`;
        } else if (change.field === 'AGENDA') {
          body += `  Agenda actualizada\n`;
        } else if (change.field === 'DOCUMENTATION') {
          body += `  Documentación actualizada\n`;
        }
      });

      // Show unchanged fields
      if (!changes.find(c => c.field === 'DATE') && !changes.find(c => c.field === 'TIME')) {
        body += `  Fecha y hora: ${formatMeetingDateTime(meeting.datetime)}\n`;
      }
      if (!changes.find(c => c.field === 'STATUS')) {
        body += `  Estado: ${meeting.status}\n`;
      }
      if (!changes.find(c => c.field === 'ATTENDEES')) {
        body += `  Asistentes: ${meeting.attendees}\n`;
      }

      body += `\n`;

      if (!changes.find(c => c.field === 'AGENDA')) {
        if (meeting.agenda) {
          body += `  Agenda:\n`;
          body += `  ${meeting.agenda}\n`;
          body += `\n`;
        }
      } else {
        body += `  Nueva agenda:\n`;
        body += `  ${meeting.agenda || '(sin agenda)'}\n`;
        body += `\n`;
      }

      if (meeting.documentation) {
        if (changes.find(c => c.field === 'DOCUMENTATION')) {
          body += `  📎 Documentación actualizada:\n`;
        } else {
          body += `  📎 Documentación:\n`;
        }
        body += `  ${meeting.documentation}\n`;
        body += `\n`;
      }

      body += buildMeetingCalendarActionsText(meeting, spreadsheetUrl, {
        includeAdd: false,
        includeUpdate: true,
        includeDelete: true,
        includeMessageAll: true
      }, attendeeEmail, ownerEmail);

      body += `  → ${changes.length} cambio(s) desde la última notificación\n\n`;
      body += `───────────────────────────────────────────────────────\n\n`;
    });
  }

  // Section 3: Deleted meetings
  if (deletedMeetings.length > 0) {
    deletedMeetings.forEach(meeting => {
      body += `🗑️ REUNIÓN CANCELADA/ELIMINADA\n\n`;
      body += `📅 ${meeting.title}\n`;
      body += `  Fecha y hora: ${formatMeetingDateTime(meeting.datetime)}\n`;
      // body += `  Estado: ${meeting.status}\n`;
      if (meeting.attendees) {
        body += `  Asistentes: ${meeting.attendees}\n`;
      }
      body += `\n`;

      body += buildMeetingCalendarActionsText(meeting, spreadsheetUrl, {
        includeAdd: false,
        includeUpdate: false,
        includeDelete: true,
        includeMessageAll: true
      }, attendeeEmail, ownerEmail);

      body += `  → Esta reunión fue eliminada del calendario\n\n`;
      body += `───────────────────────────────────────────────────────\n\n`;
    });
  }

  // Section 4: Postponed meetings
  if (postponedMeetings.length > 0) {
    postponedMeetings.forEach(changeObj => {
      const meeting = changeObj.meeting;
      const previousMeeting = changeObj.previousMeeting;

      body += `⏰ REUNIÓN POSPUESTA\n\n`;
      body += `📅 ${meeting.title}\n`;
      body += `  Fecha y hora anterior: ${formatMeetingDateTime(previousMeeting.datetime)}\n`;
      body += `  Nueva fecha y hora: ${formatMeetingDateTime(meeting.datetime)}\n`;
      body += `  Estado: ${meeting.status}\n`;
      if (meeting.attendees) {
        body += `  Asistentes: ${meeting.attendees}\n`;
      }
      body += `\n`;

      if (meeting.agenda) {
        body += `  Agenda:\n`;
        body += `  ${meeting.agenda}\n`;
        body += `\n`;
      }

      if (meeting.documentation) {
        body += `  📎 Documentación:\n`;
        body += `  ${meeting.documentation}\n`;
        body += `\n`;
      }

      body += buildMeetingCalendarActionsText(meeting, spreadsheetUrl, {
        includeAdd: false,
        includeUpdate: true,
        includeDelete: false,
        includeMessageAll: true
      }, attendeeEmail, ownerEmail);

      body += `  → Esta reunión fue pospuesta\n\n`;
      body += `───────────────────────────────────────────────────────\n\n`;
    });
  }

  // Section 5: Cancelled meetings
  if (cancelledMeetings.length > 0) {
    cancelledMeetings.forEach(changeObj => {
      const meeting = changeObj.meeting;
      const previousMeeting = changeObj.previousMeeting;

      body += `❌ REUNIÓN CANCELADA\n\n`;
      body += `📅 ${meeting.title}\n`;
      body += `  Fecha y hora original: ${formatMeetingDateTime(previousMeeting.datetime)}\n`;
      body += `  Estado: ${meeting.status}\n`;
      if (meeting.attendees) {
        body += `  Asistentes: ${meeting.attendees}\n`;
      }
      body += `\n`;

      body += buildMeetingCalendarActionsText(meeting, spreadsheetUrl, {
        includeAdd: false,
        includeUpdate: false,
        includeDelete: true,
        includeMessageAll: true
      }, attendeeEmail, ownerEmail);

      body += `  → Esta reunión fue cancelada\n\n`;
      body += `───────────────────────────────────────────────────────\n\n`;
    });
  }

  body += `═══════════════════════════════════════════════════════\n\n`;
  body += `Ver el calendario completo: ${spreadsheetUrl}\n\n`;
  body += `Notificación enviada: ${formatDateTime(new Date())}\n\n`;
  body += `---\n`;
  body += `ℹ️ Estás recibiendo esto porque fuiste invitado a reunión(es) que cambiaron.\n`;

  return body;
}

/**
 * Builds HTML email body with meeting change details and action buttons.
 * @param {string} attendeeEmail - Recipient email
 * @param {Array} newMeetings - New meetings
 * @param {Array} changedMeetings - Changed meetings with change details
 * @param {Array} deletedMeetings - Deleted meetings
 * @param {Array} postponedMeetings - Postponed meetings with change details
 * @param {Array} cancelledMeetings - Cancelled meetings with change details
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @param {string} ownerEmail - Notification sender email
 * @returns {string} Formatted HTML body
 */
function buildMeetingEmailHtml(attendeeEmail, newMeetings, changedMeetings, deletedMeetings, postponedMeetings, cancelledMeetings, spreadsheetUrl, ownerEmail) {
  const totalChanges = newMeetings.length + changedMeetings.length + deletedMeetings.length + postponedMeetings.length + cancelledMeetings.length;
  let html = '';

  html += '<div style="margin:0;padding:0;background:#f5f7fb;">';
  html += '<div style="max-width:700px;margin:0 auto;padding:24px 14px;font-family:\'Segoe UI\', Arial, sans-serif;color:#1f2937;">';
  html += '<div style="background:#ffffff;border:1px solid #e5e7eb;border-radius:14px;padding:20px;">';
  html += `<h2 style="margin:0 0 10px 0;font-size:22px;line-height:1.2;color:#111827;">Reuniones actualizadas</h2>`;
  html += `<p style="margin:0 0 16px 0;font-size:14px;color:#4b5563;">Tienes <strong>${totalChanges}</strong> reunión(es) programadas o actualizadas.</p>`;

  if (newMeetings.length > 0) {
    newMeetings.forEach(meeting => {
      html += buildMeetingCardHtml('✨ NUEVA REUNIÓN PROGRAMADA', '#e8f7ee', '#2f855a', meeting, null, spreadsheetUrl, {
        includeAdd: true,
        includeUpdate: true,
        includeDelete: true,
        includeMessageAll: true
      }, 'Esta reunión fue agendada recientemente.', attendeeEmail, ownerEmail);
    });
  }

  if (changedMeetings.length > 0) {
    changedMeetings.forEach(changeObj => {
      const meeting = changeObj.meeting;
      const changesHtml = buildMeetingChangesHtml(changeObj.changes);
      html += buildMeetingCardHtml('📝 REUNIÓN ACTUALIZADA', '#fff8e8', '#9a6700', meeting, changesHtml, spreadsheetUrl, {
        includeAdd: false,
        includeUpdate: true,
        includeDelete: true,
        includeMessageAll: true
      }, `${changeObj.changes.length} cambio(s) desde la última notificación.`, attendeeEmail, ownerEmail);
    });
  }

  if (deletedMeetings.length > 0) {
    deletedMeetings.forEach(meeting => {
      html += buildMeetingCardHtml('🗑️ REUNIÓN CANCELADA/ELIMINADA', '#fdecec', '#b42318', meeting, null, spreadsheetUrl, {
        includeAdd: false,
        includeUpdate: false,
        includeDelete: true,
        includeMessageAll: true
      }, 'Esta reunión fue eliminada del calendario.', attendeeEmail, ownerEmail);
    });
  }

  if (postponedMeetings.length > 0) {
    postponedMeetings.forEach(changeObj => {
      const meeting = changeObj.meeting;
      const previousMeeting = changeObj.previousMeeting;
      const postponedChangesHtml =
        '<div style="margin-top:8px;padding:10px;background:#ecfeff;border:1px solid #bae6fd;border-radius:8px;">' +
        '<p style="margin:0 0 6px 0;font-size:13px;font-weight:700;color:#075985;">Reprogramación detectada</p>' +
        `<p style="margin:0 0 4px 0;font-size:13px;color:#0f172a;"><strong>Fecha y hora anterior:</strong> ${escapeHtml(formatMeetingDateTime(previousMeeting.datetime))}</p>` +
        `<p style="margin:0;font-size:13px;color:#0f172a;"><strong>Nueva fecha y hora:</strong> ${escapeHtml(formatMeetingDateTime(meeting.datetime))}</p>` +
        '</div>';

      html += buildMeetingCardHtml('⏰ REUNIÓN POSPUESTA', '#e0f2fe', '#075985', meeting, postponedChangesHtml, spreadsheetUrl, {
        includeAdd: false,
        includeUpdate: true,
        includeDelete: false,
        includeMessageAll: true
      }, 'Esta reunión fue pospuesta.', attendeeEmail, ownerEmail);
    });
  }

  if (cancelledMeetings.length > 0) {
    cancelledMeetings.forEach(changeObj => {
      const meeting = changeObj.meeting;
      const previousMeeting = changeObj.previousMeeting;
      const cancelledChangesHtml =
        '<div style="margin-top:8px;padding:10px;background:#fff1f2;border:1px solid #fecdd3;border-radius:8px;">' +
        '<p style="margin:0 0 6px 0;font-size:13px;font-weight:700;color:#9f1239;">Cancelación detectada</p>' +
        `<p style="margin:0;font-size:13px;color:#0f172a;"><strong>Fecha y hora original:</strong> ${escapeHtml(formatMeetingDateTime(previousMeeting.datetime))}</p>` +
        '</div>';

      html += buildMeetingCardHtml('❌ REUNIÓN CANCELADA', '#ffe4e6', '#9f1239', meeting, cancelledChangesHtml, spreadsheetUrl, {
        includeAdd: false,
        includeUpdate: false,
        includeDelete: true,
        includeMessageAll: true
      }, 'Esta reunión fue cancelada.', attendeeEmail, ownerEmail);
    });
  }

  html += '<div style="margin-top:18px;padding-top:14px;border-top:1px solid #e5e7eb;">';
  html += `<p style="margin:0 0 8px 0;font-size:13px;"><a href="${escapeHtml(spreadsheetUrl)}" target="_blank" style="color:#1d4ed8;text-decoration:none;">Ver el calendario completo</a></p>`;
  html += `<p style="margin:0;font-size:12px;color:#6b7280;">Notificación enviada: ${escapeHtml(formatDateTime(new Date()))}</p>`;
  html += '<p style="margin:10px 0 0 0;font-size:12px;color:#6b7280;">Estás recibiendo esto porque fuiste invitado a reunión(es) que cambiaron.</p>';
  html += '</div>';

  html += '</div>';
  html += '</div>';
  html += '</div>';

  return html;
}

/**
 * Builds an HTML card for a meeting item.
 * @param {string} title - Card title label
 * @param {string} badgeBackground - Badge background color
 * @param {string} badgeColor - Badge text color
 * @param {Object} meeting - Meeting object
 * @param {string|null} changesHtml - Changes block HTML
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @param {Object} actionOptions - Which calendar actions to include
 * @param {string} footerText - Card footer text
 * @param {string} attendeeEmail - Recipient email
 * @param {string} ownerEmail - Notification sender email
 * @returns {string} HTML card
 */
function buildMeetingCardHtml(title, badgeBackground, badgeColor, meeting, changesHtml, spreadsheetUrl, actionOptions, footerText, attendeeEmail, ownerEmail) {
  let card = '';

  card += '<div style="margin:0 0 14px 0;padding:16px;border:1px solid #e5e7eb;border-radius:12px;background:#ffffff;">';
  card += `<div style="display:inline-block;font-size:12px;font-weight:700;padding:4px 10px;border-radius:999px;background:${badgeBackground};color:${badgeColor};margin-bottom:10px;">${escapeHtml(title)}</div>`;
  card += `<h3 style="margin:0 0 10px 0;font-size:18px;color:#111827;">${escapeHtml(meeting.title || '(sin título)')}</h3>`;
  card += `<p style="margin:0 0 4px 0;font-size:14px;"><strong>Fecha y hora:</strong> ${escapeHtml(formatMeetingDateTime(meeting.datetime))}</p>`;

  if (meeting.status) {
    card += `<p style="margin:0 0 4px 0;font-size:14px;"><strong>Estado:</strong> ${escapeHtml(meeting.status)}</p>`;
  }

  if (meeting.attendees) {
    card += `<p style="margin:0 0 10px 0;font-size:14px;"><strong>Asistentes:</strong> ${escapeHtml(meeting.attendees)}</p>`;
  }

  if (changesHtml) {
    card += changesHtml;
  }

  if (meeting.agenda) {
    card += `<p style="margin:10px 0 4px 0;font-size:13px;font-weight:700;color:#374151;">Agenda</p>`;
    card += `<p style="margin:0;font-size:13px;color:#4b5563;line-height:1.45;">${textToHtml(meeting.agenda)}</p>`;
  }

  if (meeting.documentation) {
    card += `<p style="margin:10px 0 4px 0;font-size:13px;font-weight:700;color:#374151;">Documentación</p>`;
    card += `<p style="margin:0;font-size:13px;color:#4b5563;line-height:1.45;">${textToHtml(meeting.documentation)}</p>`;
  }

  card += buildMeetingCalendarActionsHtml(meeting, spreadsheetUrl, actionOptions, attendeeEmail, ownerEmail);

  card += `<p style="margin:10px 0 0 0;font-size:12px;color:#6b7280;">${escapeHtml(footerText)}</p>`;
  card += '</div>';

  return card;
}

/**
 * Builds HTML list of meeting changes.
 * @param {Array} changes - Changed fields
 * @returns {string} HTML block
 */
function buildMeetingChangesHtml(changes) {
  if (!changes || changes.length === 0) {
    return '';
  }

  let html = '';
  html += '<div style="margin-top:8px;padding:10px;background:#fffbeb;border:1px solid #f3e8bb;border-radius:8px;">';
  html += '<p style="margin:0 0 6px 0;font-size:13px;font-weight:700;color:#7c5a00;">Cambios detectados</p>';
  html += '<ul style="margin:0;padding-left:18px;font-size:13px;color:#4b5563;">';

  changes.forEach(change => {
    html += `<li style="margin:0 0 4px 0;">${escapeHtml(formatMeetingChangeText(change))}</li>`;
  });

  html += '</ul>';
  html += '</div>';
  return html;
}

/**
 * Formats one changed meeting field for display.
 * @param {Object} change - Change object
 * @returns {string} Human-readable change text
 */
function formatMeetingChangeText(change) {
  if (change.field === 'TITLE') {
    return `Título cambiado de: "${change.oldValue}" a "${change.newValue}"`;
  }

  if (change.field === 'DATE') {
    return `Fecha: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'TIME') {
    return `Hora: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'STATUS') {
    return `Estado: ${change.oldValue} -> ${change.newValue}`;
  }

  if (change.field === 'ATTENDEES') {
    return `Asistentes actualizados: ${change.newValue}`;
  }

  if (change.field === 'AGENDA') {
    return 'Agenda actualizada';
  }

  if (change.field === 'DOCUMENTATION') {
    return 'Documentación actualizada';
  }

  return `${change.field}: ${change.oldValue} -> ${change.newValue}`;
}

/**
 * Renders HTML calendar action buttons for meeting emails.
 * @param {Object} meeting - Meeting object
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @param {Object} options - Which actions to include
 * @param {string} attendeeEmail - Recipient email
 * @param {string} ownerEmail - Notification sender email
 * @returns {string} Formatted HTML block
 */
function buildMeetingCalendarActionsHtml(meeting, spreadsheetUrl, options, attendeeEmail, ownerEmail) {
  const actions = buildMeetingCalendarActions(meeting, spreadsheetUrl, attendeeEmail, ownerEmail);
  let html = '';

  html += '<div style="margin-top:12px;padding:10px;background:#f8fafc;border:1px solid #e5e7eb;border-radius:10px;">';
  html += '<p style="margin:0 0 8px 0;font-size:13px;font-weight:700;color:#334155;">Acciones en Google Calendar</p>';

  if (options.includeAdd) {
    if (actions.add) {
      html += `<a href="${escapeHtml(actions.add)}" target="_blank" style="display:inline-block;margin:0 8px 8px 0;padding:8px 12px;background:#1d4ed8;color:#ffffff;text-decoration:none;border-radius:8px;font-size:12px;font-weight:700;">Agregar a Google Calendar</a>`;
    } else {
      html += '<p style="margin:0 0 8px 0;font-size:12px;color:#b45309;">Agregar a Google Calendar no disponible (fecha/hora inválida).</p>';
    }
  }

  if (options.includeUpdate) {
    html += `<a href="${escapeHtml(actions.update)}" target="_blank" style="display:inline-block;margin:0 8px 8px 0;padding:8px 12px;background:#0f766e;color:#ffffff;text-decoration:none;border-radius:8px;font-size:12px;font-weight:700;">Actualizar reunión en Google Calendar</a>`;
    html += '<p style="margin:0 0 8px 0;font-size:12px;color:#64748b;">Abre Google Calendar para localizar la reunión y editarla manualmente.</p>';
  }

  if (options.includeDelete) {
    html += `<a href="${escapeHtml(actions.delete)}" target="_blank" style="display:inline-block;margin:0 8px 8px 0;padding:8px 12px;background:#b42318;color:#ffffff;text-decoration:none;border-radius:8px;font-size:12px;font-weight:700;">Eliminar reunión en Google Calendar</a>`;
    html += '<p style="margin:0;font-size:12px;color:#64748b;">Abre Google Calendar para localizar la reunión y eliminarla manualmente.</p>';
  }

  if (options.includeMessageAll) {
    if (actions.messageAll) {
      html += `<a href="${escapeHtml(actions.messageAll)}" target="_blank" style="display:inline-block;margin:8px 8px 8px 0;padding:8px 12px;background:#7c3aed;color:#ffffff;text-decoration:none;border-radius:8px;font-size:12px;font-weight:700;">Enviar mensaje a todos los asistentes</a>`;
      html += '<p style="margin:0;font-size:12px;color:#64748b;">Abre tu cliente de correo y prepara un mensaje para asistentes + propietario de la reunión.</p>';
    } else {
      html += '<p style="margin:8px 0 0 0;font-size:12px;color:#b45309;">No se pudo generar el enlace para enviar mensaje a todos (sin correos válidos).</p>';
    }
  }

  html += '</div>';
  return html;
}

/**
 * Escapes basic HTML special characters.
 * @param {any} value - Value to escape
 * @returns {string} Escaped HTML string
 */
function escapeHtml(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Converts plain text to HTML-safe text preserving line breaks.
 * @param {any} value - Plain text value
 * @returns {string} HTML-safe text with <br> line breaks
 */
function textToHtml(value) {
  return escapeHtml(value).replace(/\n/g, '<br>');
}



// ================================================================================
// SECTION 8: UTILITY FUNCTIONS
// ================================================================================

/**
 * Validates email address format
 * @param {string} email - Email address to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
  if (!email) return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email.toString().trim());
}

/**
 * Parses multiple email addresses from a string
 * Supports comma and semicolon separators
 * @param {string} emailString - String containing one or more emails
 * @returns {Array} Array of unique, valid email addresses
 */
function parseMultipleEmails(emailString) {
  if (!emailString) return [];

  // Normalize: replace semicolons with commas
  const normalized = emailString.toString().replace(/;/g, ',');

  // Split by commas, trim, lowercase, and filter valid emails
  const emails = normalized.split(',')
    .map(email => email.trim().toLowerCase())
    .filter(email => email && isValidEmail(email));

  // Remove duplicates using Set
  return [...new Set(emails)];
}

/**
 * Formats date value to dd-mm-yyyy string
 * @param {Date|string} dateValue - Date to format
 * @returns {string} Formatted date string
 */
function formatDate(dateValue) {
  if (!dateValue) return '';

  try {
    let date;
    if (dateValue instanceof Date) {
      date = dateValue;
    } else if (typeof dateValue === 'string') {
      date = new Date(dateValue);
    } else {
      return String(dateValue);
    }

    if (isNaN(date.getTime())) {
      return String(dateValue);
    }

    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();

    return `${day}-${month}-${year}`;
  } catch (error) {
    return String(dateValue);
  }
}

/**
 * Formats date and time for timestamps
 * @param {Date} date - Date object
 * @returns {string} Formatted date-time string
 */
function formatDateTime(date) {
  const day = String(date.getDate()).padStart(2, '0');
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');

  return `${day}-${month}-${year} ${hours}:${minutes}`;
}

/**
 * Formats time value to HH:mm string.
 * @param {Date|string} timeValue - Time value to format
 * @returns {string} Formatted time string
 */
function formatTime(timeValue) {
  if (!timeValue) return '';

  try {
    const parsed = parseLocalizedDateOrDateTime(timeValue);
    if (parsed) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'HH:mm');
    }

    if (timeValue instanceof Date) {
      return Utilities.formatDate(timeValue, Session.getScriptTimeZone(), 'HH:mm');
    }

    const raw = String(timeValue).trim();
    const hhmmMatch = raw.match(/^(\d{1,2}):(\d{2})$/);
    if (hhmmMatch) {
      const hh = hhmmMatch[1].padStart(2, '0');
      const mm = hhmmMatch[2];
      return `${hh}:${mm}`;
    }

    return raw;
  } catch (error) {
    return String(timeValue);
  }
}

/**
 * Formats meeting date-time for display in emails
 * @param {Date|string} datetimeValue - Datetime value from Google Sheets
 * @returns {string} Formatted datetime string "dd-mm-yyyy HH:mm"
 */
function formatMeetingDateTime(datetimeValue) {
  if (!datetimeValue) return '(no especificada)';

  try {
    let date;
    if (datetimeValue instanceof Date) {
      date = datetimeValue;
    } else if (typeof datetimeValue === 'string') {
      date = new Date(datetimeValue);
    } else {
      return String(datetimeValue);
    }

    if (isNaN(date.getTime())) {
      return String(datetimeValue);
    }

    // Reuse formatDateTime() for consistent formatting
    return formatDateTime(date);

  } catch (error) {
    return String(datetimeValue);
  }
}

/**
 * Formats datetime for Google Calendar links.
 * @param {Date} date - Date object
 * @returns {string} Formatted datetime string yyyyMMdd'T'HHmmss
 */
function formatGoogleCalendarDateTime(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyyMMdd'T'HHmmss");
}

/**
 * Adds minutes to a date and returns a new Date object.
 * @param {Date} date - Source date
 * @param {number} minutes - Minutes to add
 * @returns {Date} New date with minutes added
 */
function addMinutesToDate(date, minutes) {
  const result = new Date(date.getTime());
  result.setMinutes(result.getMinutes() + minutes);
  return result;
}

/**
 * Builds meeting details text for Google Calendar links.
 * @param {Object} meeting - Meeting object
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @returns {string} Details text
 */
function buildMeetingCalendarDetails(meeting, spreadsheetUrl) {
  const parts = [];

  if (meeting.agenda) {
    parts.push(`Agenda:\n${meeting.agenda}`);
  }

  if (meeting.documentation) {
    parts.push(`Documentación:\n${meeting.documentation}`);
  }

  if (spreadsheetUrl) {
    parts.push(`Calendario completo:\n${spreadsheetUrl}`);
  }

  return parts.join('\n\n');
}

/**
 * Builds a Google Calendar prefilled create-event URL.
 * @param {Object} meeting - Meeting object
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @returns {string|null} Calendar URL or null when datetime is invalid
 */
function buildGoogleCalendarAddUrl(meeting, spreadsheetUrl) {
  const startDate = parseLocalizedDateOrDateTime(meeting.datetime);
  if (!startDate) {
    return null;
  }

  const endDate = addMinutesToDate(startDate, CONFIG.MEETINGS.CALENDAR_DEFAULT_DURATION_MINUTES);
  const timezone = Session.getScriptTimeZone();
  const details = buildMeetingCalendarDetails(meeting, spreadsheetUrl);

  const params = [
    'action=TEMPLATE',
    `text=${encodeURIComponent(meeting.title || 'Reunión')}`,
    `dates=${encodeURIComponent(`${formatGoogleCalendarDateTime(startDate)}/${formatGoogleCalendarDateTime(endDate)}`)}`,
    `details=${encodeURIComponent(details)}`,
    `ctz=${encodeURIComponent(timezone)}`
  ];

  return `https://calendar.google.com/calendar/render?${params.join('&')}`;
}

/**
 * Builds a Google Calendar search URL to help users locate a meeting manually.
 * @param {Object} meeting - Meeting object
 * @returns {string} Calendar search URL
 */
function buildGoogleCalendarSearchUrl(meeting) {
  const queryParts = [meeting.title || 'reunión'];

  if (meeting.datetime) {
    queryParts.push(formatMeetingDateTime(meeting.datetime));
  }

  return `https://calendar.google.com/calendar/u/0/r/search?q=${encodeURIComponent(queryParts.join(' '))}`;
}

/**
 * Builds calendar action URLs for meeting emails.
 * @param {Object} meeting - Meeting object
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @returns {Object} Calendar action URLs
 */
function buildMeetingCalendarActions(meeting, spreadsheetUrl, attendeeEmail, ownerEmail) {
  const searchUrl = buildGoogleCalendarSearchUrl(meeting);

  return {
    add: buildGoogleCalendarAddUrl(meeting, spreadsheetUrl),
    update: searchUrl,
    delete: searchUrl,
    messageAll: buildMeetingMessageAllMailtoUrl(meeting, attendeeEmail, ownerEmail)
  };
}

/**
 * Renders plain-text calendar action links for meeting emails.
 * @param {Object} meeting - Meeting object
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @param {Object} options - Which actions to include
 * @param {string} attendeeEmail - Recipient email
 * @param {string} ownerEmail - Notification sender email
 * @returns {string} Formatted text block
 */
function buildMeetingCalendarActionsText(meeting, spreadsheetUrl, options, attendeeEmail, ownerEmail) {
  const actions = buildMeetingCalendarActions(meeting, spreadsheetUrl, attendeeEmail, ownerEmail);
  const lines = [];

  lines.push('  Acciones en Google Calendar:');

  if (options.includeAdd) {
    if (actions.add) {
      lines.push(`  - [Agregar a Google Calendar] ${actions.add}`);
    } else {
      lines.push('  - [Agregar a Google Calendar] No disponible (fecha/hora inválida)');
    }
  }

  if (options.includeUpdate) {
    lines.push(`  - [Actualizar reunión en Google Calendar] ${actions.update}`);
    lines.push('    (Abre Google Calendar para localizar la reunión y editarla manualmente)');
  }

  if (options.includeDelete) {
    lines.push(`  - [Eliminar reunión en Google Calendar] ${actions.delete}`);
    lines.push('    (Abre Google Calendar para localizar la reunión y eliminarla manualmente)');
  }

  if (options.includeMessageAll) {
    if (actions.messageAll) {
      lines.push(`  - [Enviar mensaje a todos los asistentes] ${actions.messageAll}`);
      lines.push('    (Incluye asistentes y propietario de la reunión en tu correo)');
    } else {
      lines.push('  - [Enviar mensaje a todos los asistentes] No disponible (sin correos válidos)');
    }
  }

  lines.push('');
  return lines.join('\n') + '\n';
}

/**
 * Resolves notification sender email to be used as meeting owner.
 * @returns {string} Sender email or empty string when unavailable
 */
function resolveNotificationSenderEmail() {
  try {
    const activeEmail = sanitizeValue(Session.getActiveUser().getEmail()).toLowerCase();
    if (isValidEmail(activeEmail)) {
      return activeEmail;
    }
  } catch (error) {
    Logger.log(`Active user email unavailable: ${error.toString()}`);
  }

  try {
    const effectiveEmail = sanitizeValue(Session.getEffectiveUser().getEmail()).toLowerCase();
    if (isValidEmail(effectiveEmail)) {
      return effectiveEmail;
    }
  } catch (error) {
    Logger.log(`Effective user email unavailable: ${error.toString()}`);
  }

  Logger.log('Notification sender email unavailable; message-all action will use attendees only.');
  return '';
}

/**
 * Builds a normalized and deduplicated recipient list for message-all action.
 * Includes meeting attendees and owner (notification sender).
 * @param {Object} meeting - Meeting object
 * @param {string} attendeeEmail - Notification recipient email
 * @param {string} ownerEmail - Notification sender email
 * @returns {Array<string>} Recipient emails
 */
function buildMeetingMessageRecipients(meeting, attendeeEmail, ownerEmail) {
  const attendees = parseMultipleEmails(meeting.attendees || '');
  const recipients = [];

  attendees.forEach(email => {
    if (isValidEmail(email)) {
      recipients.push(email.toLowerCase());
    }
  });

  if (isValidEmail(ownerEmail)) {
    recipients.push(ownerEmail.toLowerCase());
  }

  if (isValidEmail(attendeeEmail)) {
    recipients.push(attendeeEmail.toLowerCase());
  }

  return [...new Set(recipients)];
}

/**
 * Builds a mailto URL for messaging all attendees and owner.
 * @param {Object} meeting - Meeting object
 * @param {string} attendeeEmail - Notification recipient email
 * @param {string} ownerEmail - Notification sender email
 * @returns {string|null} Mailto URL or null if no recipients available
 */
function buildMeetingMessageAllMailtoUrl(meeting, attendeeEmail, ownerEmail) {
  const recipients = buildMeetingMessageRecipients(meeting, attendeeEmail, ownerEmail);
  if (recipients.length === 0) {
    return null;
  }

  const subject = `Mensaje sobre reunión: ${meeting.title || 'Reunión'}`;
  const bodyLines = [
    'Hola equipo,',
    '',
    `Comparto este mensaje sobre la reunión "${meeting.title || 'Reunión'}".`,
    `Fecha y hora: ${formatMeetingDateTime(meeting.datetime)}`,
    '',
    'Mensaje:'
  ];

  const toParam = recipients.join(',');
  const subjectParam = encodeURIComponent(subject);
  const bodyParam = encodeURIComponent(bodyLines.join('\n'));

  return `mailto:${toParam}?subject=${subjectParam}&body=${bodyParam}`;
}

/**
 * Parses localized date strings (dd-mm-yyyy or dd-mm-yyyy HH:mm) safely.
 * @param {Date|string} value - Date-like value
 * @returns {Date|null} Parsed date or null if not parseable
 */
function parseLocalizedDateOrDateTime(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  if (value instanceof Date) {
    return isNaN(value.getTime()) ? null : value;
  }

  const raw = String(value).trim();

  // dd-mm-yyyy HH:mm
  const dateTimeMatch = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{4})\s+(\d{1,2}):(\d{2})$/);
  if (dateTimeMatch) {
    const day = Number(dateTimeMatch[1]);
    const month = Number(dateTimeMatch[2]);
    const year = Number(dateTimeMatch[3]);
    const hour = Number(dateTimeMatch[4]);
    const minute = Number(dateTimeMatch[5]);
    const parsed = new Date(year, month - 1, day, hour, minute, 0, 0);

    if (
      parsed.getFullYear() === year &&
      parsed.getMonth() === month - 1 &&
      parsed.getDate() === day &&
      parsed.getHours() === hour &&
      parsed.getMinutes() === minute
    ) {
      return parsed;
    }
    return null;
  }

  // dd-mm-yyyy
  const dateMatch = raw.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (dateMatch) {
    const day = Number(dateMatch[1]);
    const month = Number(dateMatch[2]);
    const year = Number(dateMatch[3]);
    const parsed = new Date(year, month - 1, day, 0, 0, 0, 0);

    if (
      parsed.getFullYear() === year &&
      parsed.getMonth() === month - 1 &&
      parsed.getDate() === day
    ) {
      return parsed;
    }
    return null;
  }

  // Final fallback for ISO or other parser-friendly values
  const fallback = new Date(raw);
  return isNaN(fallback.getTime()) ? null : fallback;
}

/**
 * Normalizes arbitrary values to comparable strings.
 * @param {any} value - Any comparable value
 * @returns {string} Normalized value for equality checks
 */
function normalizeStringForComparison(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  return String(value).trim();
}

/**
 * Removes diacritics (accents) for robust header comparisons.
 * @param {string} value - Input text
 * @returns {string} Text without diacritics
 */
function removeDiacritics(value) {
  return normalizeStringForComparison(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/**
 * Normalizes date values to yyyy-mm-dd for stable comparisons.
 * @param {Date|string} value - Date value from sheet/history
 * @returns {string} Normalized date string or empty string
 */
function normalizeDateForComparison(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  const date = parseLocalizedDateOrDateTime(value);

  if (!date) {
    return normalizeStringForComparison(value);
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Normalizes datetime values to yyyy-mm-dd HH:mm for stable comparisons.
 * @param {Date|string} value - Datetime value from sheet/history
 * @returns {string} Normalized datetime string or empty string
 */
function normalizeDateTimeForComparison(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  const date = parseLocalizedDateOrDateTime(value);

  if (!date) {
    return normalizeStringForComparison(value);
  }

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

/**
 * Normalizes time values to HH:mm for stable comparisons.
 * @param {Date|string} value - Time value from sheet/history
 * @returns {string} Normalized time string or empty string
 */
function normalizeTimeForComparison(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }

  const parsed = parseLocalizedDateOrDateTime(value);
  if (parsed) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'HH:mm');
  }

  const raw = String(value).trim();
  const hhmmMatch = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (hhmmMatch) {
    return `${hhmmMatch[1].padStart(2, '0')}:${hhmmMatch[2]}`;
  }

  return raw;
}

/**
 * Extracts meeting datetime/date/time from the split row format.
 * @param {Array} row - Meeting row
 * @returns {Object} Extracted meeting date-time fields and content columns
 */
function extractMeetingDateTimeParts(row) {
  const date = row[CONFIG.MEETINGS.COLUMNS.DATE] || '';
  const time = row[CONFIG.MEETINGS.COLUMNS.TIME] || '';

  return {
    datetime: buildDateTimeFromDateAndTime(date, time),
    date: date,
    time: time,
    agenda: row[CONFIG.MEETINGS.COLUMNS.AGENDA],
    documentation: row[CONFIG.MEETINGS.COLUMNS.DOCUMENTATION]
  };
}

/**
 * Combines separate date/time inputs into a Date when possible.
 * @param {Date|string} dateValue - Date value
 * @param {Date|string} timeValue - Time value
 * @returns {Date|string} Date object when parseable; otherwise combined/raw string
 */
function buildDateTimeFromDateAndTime(dateValue, timeValue) {
  const datePart = normalizeDateForComparison(dateValue);
  const timePart = normalizeTimeForComparison(timeValue);

  if (!datePart && !timePart) {
    return '';
  }

  if (datePart && timePart) {
    const match = datePart.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    const timeMatch = timePart.match(/^(\d{2}):(\d{2})$/);

    if (match && timeMatch) {
      return new Date(
        Number(match[1]),
        Number(match[2]) - 1,
        Number(match[3]),
        Number(timeMatch[1]),
        Number(timeMatch[2]),
        0,
        0
      );
    }

    return `${datePart} ${timePart}`;
  }

  return datePart || timePart;
}

/**
 * Creates an order-independent fingerprint for task reconciliation.
 * @param {Object} task - Task object
 * @returns {string} Fingerprint
 */
function buildTaskFingerprint(task) {
  return [
    normalizeStringForComparison(task.sheetName),
    normalizeStringForComparison(task.email),
    normalizeStringForComparison(task.task),
    normalizeStringForComparison(task.priority),
    normalizeStringForComparison(task.status),
    normalizeDateForComparison(task.initDate),
    normalizeDateForComparison(task.finishDate),
    normalizeStringForComparison(task.product),
    normalizeStringForComparison(task.notes)
  ].join('|');
}

/**
 * Creates an order-independent fingerprint for meeting reconciliation.
 * @param {Object} meeting - Meeting object
 * @returns {string} Fingerprint
 */
function buildMeetingFingerprint(meeting) {
  return [
    normalizeStringForComparison(meeting.email),
    normalizeStringForComparison(meeting.title),
    normalizeDateTimeForComparison(meeting.datetime),
    normalizeStringForComparison(meeting.status),
    normalizeStringForComparison(meeting.agenda),
    normalizeStringForComparison(meeting.documentation)
  ].join('|');
}


/**
 * Sanitizes cell values for safe handling
 * @param {any} value - Cell value
 * @returns {string} Sanitized string value
 */
function sanitizeValue(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  return String(value).trim();
}

/**
 * Shows an alert when UI context is available; otherwise logs a fallback message.
 * @param {string} message - Alert message
 */
function showUiAlertSafe(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (error) {
    Logger.log(`UI alert unavailable. Message: ${message}. Reason: ${error.toString()}`);
  }
}
