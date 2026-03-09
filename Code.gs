/**
 * ================================================================================
 * PROJECT MANAGEMENT TASK NOTIFICATION SYSTEM
 * ================================================================================
 * 
 * This Google Apps Script adds intelligent task change notifications to your
 * project management spreadsheet.MP
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
  // History tracking sheet name
  HISTORY_SHEET_NAME: '_task_history',

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
  SPREADSHEET_MENU_ITEM: 'Notificar a los asignados',

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
      `Snapshot guardado en la hoja ${CONFIG.HISTORY_SHEET_NAME}.`
  },

  // ================================================================================
  // MEETINGS CONFIGURATION
  // ================================================================================

  MEETINGS: {
    SHEET_NAME: 'Reuniones',
    HISTORY_SHEET_NAME: '_meetings_history',
    EMAIL_SUBJECT: 'Reunión de equipo - convocatoria',
    MENU_ITEM: 'Notificar reuniones',

    COLUMNS: {
      TITLE: 0,         // Columna A
      ATTENDEES: 1,     // Columna B
      STATUS: 2,        // Columna C
      DATETIME: 3,      // Columna D
      AGENDA: 4,        // Columna E
      DOCUMENTATION: 5  // Columna F
    },

    HEADER_ROW: 1,
    FIRST_DATA_ROW: 2,

    MONITORED_FIELDS: ['TITLE', 'ATTENDEES', 'STATUS', 'DATETIME', 'AGENDA', 'DOCUMENTATION'],

    FIELD_NAMES: {
      TITLE: 'Título',
      ATTENDEES: 'Asistentes',
      STATUS: 'Estado',
      DATETIME: 'Fecha y hora',
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
      SUMMARY_MESSAGE: (totalMeetings, newCount, changedCount, deletedCount, unchangedCount, emailsSent) =>
        `• ${totalMeetings} reunión(es) escaneada(s)\n` +
        `• ${newCount} reunión(es) nueva(s) detectada(s)\n` +
        `• ${changedCount} reunión(es) cambiada(s)\n` +
        `• ${deletedCount} reunión(es) eliminada(s)\n` +
        `• ${unchangedCount} reunión(es) sin cambios (no notificadas)\n\n` +
        `📧 Correos enviados a ${emailsSent} asistente(s)\n\n` +
        `Snapshot guardado en la hoja ${CONFIG.MEETINGS.HISTORY_SHEET_NAME}.`
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
    const startTime = new Date();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Step 1: Ensure History sheet exists
    createHistorySheet();

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
    const startTime = new Date();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Step 1: Ensure Meetings History sheet exists
    createMeetingsHistorySheet();

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
    const unchangedCount = changes.unchanged.length;

    const summary = CONFIG.MEETINGS.ALERTS.SUMMARY_MESSAGE(
      totalMeetings,
      newCount,
      changedCount,
      deletedCount,
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
    // Skip the History sheet
    if (sheet.getName() === CONFIG.HISTORY_SHEET_NAME) {
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

  // Get all data rows (skip header)
  const dataRange = sheet.getRange(CONFIG.MEETINGS.FIRST_DATA_ROW, 1, lastRow - CONFIG.MEETINGS.HEADER_ROW, 6);
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

  // Create one meeting object per attendee
  return emailList.map(email => ({
    meetingId: generateMeetingId(sheet.getName(), rowIndex),
    sheetName: sheet.getName(),
    rowNumber: rowIndex,
    title: sanitizeValue(row[CONFIG.MEETINGS.COLUMNS.TITLE]),
    attendees: attendeesString.toString().trim(), // Keep full list for display
    email: email, // Individual email for routing
    status: sanitizeValue(row[CONFIG.MEETINGS.COLUMNS.STATUS]),
    datetime: row[CONFIG.MEETINGS.COLUMNS.DATETIME], // Date object from Google Sheets
    agenda: sanitizeValue(row[CONFIG.MEETINGS.COLUMNS.AGENDA]),
    documentation: sanitizeValue(row[CONFIG.MEETINGS.COLUMNS.DOCUMENTATION])
  }));
}

/**
 * Generates unique meeting ID from sheet name and row number
 * @param {string} sheetName - Name of the sheet
 * @param {number} rowIndex - Row number
 * @returns {string} Unique meeting ID
 */
function generateMeetingId(sheetName, rowIndex) {
  return `${sheetName}_Row${rowIndex}`;
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
 * Creates the History sheet if it doesn't exist
 */
function createHistorySheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET_NAME);

  if (!historySheet) {
    historySheet = spreadsheet.insertSheet(CONFIG.HISTORY_SHEET_NAME);

    historySheet.getRange(1, 1, 1, CONFIG.HISTORY_SHEET_HEADERS.length).setValues([CONFIG.HISTORY_SHEET_HEADERS]);
    historySheet.getRange(1, 1, 1, CONFIG.HISTORY_SHEET_HEADERS.length).setFontWeight('bold');
    historySheet.setFrozenRows(1);

    Logger.log('Created History sheet');
  }
}

/**
 * Loads the most recent snapshot of each task from History sheet
 * @returns {Map} Map of taskId+email to most recent task state
 */
function getPreviousSnapshot() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET_NAME);

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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = spreadsheet.getSheetByName(CONFIG.HISTORY_SHEET_NAME);

  if (!historySheet) {
    return;
  }

  const timestamp = formatDateTime(new Date());
  const rows = [];

  // Create a set of notified task IDs
  const notifiedIds = new Set();
  changes.new.forEach(t => notifiedIds.add(t.taskId));
  changes.changed.forEach(t => notifiedIds.add(t.task.taskId));

  // Add current state of all tasks
  tasks.forEach(task => {
    let action = 'UNCHANGED';
    if (changes.new.find(t => t.taskId === task.taskId)) {
      action = 'NEW';
    } else if (changes.changed.find(t => t.task.taskId === task.taskId)) {
      action = 'CHANGED';
    }

    const notified = notifiedIds.has(task.taskId) ? 'YES' : 'NO';

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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = spreadsheet.getSheetByName(CONFIG.MEETINGS.HISTORY_SHEET_NAME);

  if (!historySheet) {
    historySheet = spreadsheet.insertSheet(CONFIG.MEETINGS.HISTORY_SHEET_NAME);

    historySheet.getRange(1, 1, 1, CONFIG.MEETINGS.HISTORY_HEADERS.length).setValues([CONFIG.MEETINGS.HISTORY_HEADERS]);
    historySheet.getRange(1, 1, 1, CONFIG.MEETINGS.HISTORY_HEADERS.length).setFontWeight('bold');
    historySheet.setFrozenRows(1);

    Logger.log('Created Meetings History sheet');
  }
}

/**
 * Loads the most recent snapshot of each meeting from Meetings History sheet
 * @returns {Map} Map of meetingId+email to most recent meeting state
 */
function getPreviousMeetingsSnapshot() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = spreadsheet.getSheetByName(CONFIG.MEETINGS.HISTORY_SHEET_NAME);

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
    const email = row[3]; // Attendees column (we'll extract individual email)

    // Create unique key combining meetingId and email
    const key = `${meetingId}|${email}`;

    // Only store the first (most recent) occurrence of each meeting+email combination
    if (!snapshot.has(key)) {
      snapshot.set(key, {
        meetingId: meetingId,
        title: row[2],
        attendees: row[3],
        email: email,
        status: row[4],
        datetime: row[5],
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
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const historySheet = spreadsheet.getSheetByName(CONFIG.MEETINGS.HISTORY_SHEET_NAME);

  if (!historySheet) {
    return;
  }

  const timestamp = formatDateTime(new Date());
  const rows = [];

  // Create a set of notified meeting IDs
  const notifiedIds = new Set();
  changes.new.forEach(m => notifiedIds.add(m.meetingId));
  changes.changed.forEach(m => notifiedIds.add(m.meeting.meetingId));

  // Add current state of all meetings
  meetings.forEach(meeting => {
    let action = 'UNCHANGED';
    if (changes.new.find(m => m.meetingId === meeting.meetingId)) {
      action = 'NEW';
    } else if (changes.changed.find(m => m.meeting.meetingId === meeting.meetingId)) {
      action = 'CHANGED';
    }

    const notified = notifiedIds.has(meeting.meetingId) ? 'YES' : 'NO';

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

  // Check each current task
  currentTasks.forEach(currentTask => {
    const key = `${currentTask.taskId}|${currentTask.email}`;
    currentKeys.add(key);

    if (!previousSnapshot.has(key)) {
      // New task (or new assignee to existing task)
      changes.new.push(currentTask);
    } else {
      // Existing task+email - check for changes
      const previousTask = previousSnapshot.get(key);
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
    }
  });

  // Check for deleted tasks (or removed assignees)
  previousSnapshot.forEach((previousTask, key) => {
    if (!currentKeys.has(key) && previousTask.action !== 'DELETED') {
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
    const currentNorm = String(currentValue || '').trim();
    const previousNorm = String(previousValue || '').trim();

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
    unchanged: []
  };

  const currentKeys = new Set();

  // Check each current meeting
  currentMeetings.forEach(currentMeeting => {
    const key = `${currentMeeting.meetingId}|${currentMeeting.email}`;
    currentKeys.add(key);

    if (!previousSnapshot.has(key)) {
      // New meeting (or new attendee to existing meeting)
      changes.new.push(currentMeeting);
    } else {
      // Existing meeting+email - check for changes
      const previousMeeting = previousSnapshot.get(key);
      const fieldChanges = detectMeetingFieldChanges(currentMeeting, previousMeeting);

      if (fieldChanges.length > 0) {
        changes.changed.push({
          meeting: currentMeeting,
          changes: fieldChanges,
          previousMeeting: previousMeeting
        });
      } else {
        changes.unchanged.push(currentMeeting);
      }
    }
  });

  // Check for deleted meetings (or removed attendees)
  previousSnapshot.forEach((previousMeeting, key) => {
    if (!currentKeys.has(key) && previousMeeting.action !== 'DELETED') {
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
      case 'ATTENDEES':
        currentValue = currentMeeting.attendees;
        previousValue = previousMeeting.attendees;
        break;
      case 'STATUS':
        currentValue = currentMeeting.status;
        previousValue = previousMeeting.status;
        break;
      case 'DATETIME':
        currentValue = formatMeetingDateTime(currentMeeting.datetime);
        previousValue = formatMeetingDateTime(previousMeeting.datetime);
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
    const currentNorm = String(currentValue || '').trim();
    const previousNorm = String(previousValue || '').trim();

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
      const emailBody = buildEmailBody(
        email,
        tasks.new,
        tasks.changed,
        tasks.deleted,
        spreadsheetUrl
      );

      MailApp.sendEmail({
        to: email,
        subject: CONFIG.EMAIL_SUBJECT,
        body: emailBody
      });

      emailsSent++;
      Logger.log(`Email sent to: ${email}`);

    } catch (error) {
      Logger.log(`Failed to send email to ${email}: ${error.toString()}`);
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
      body += `  Prioridad: ${task.priority}\n`;
      body += `  Estado: ${task.status}\n`;
      body += `  Producto: ${task.product}\n`;
      if (task.assignees) {
        body += `  Asignado a: ${task.assignees}\n`;
      }
      if (task.notes) {
        body += `  Notas: ${task.notes}\n`;
      }
      body += `  \n`;
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
 * Sends consolidated email notifications to attendees with changed meetings
 * @param {Object} changes - Detected changes
 * @param {string} spreadsheetUrl - URL of the spreadsheet
 * @returns {Object} Notification statistics
 */
function sendMeetingNotifications(changes, spreadsheetUrl) {
  const meetingsByAttendee = groupMeetingsByAttendee(changes);
  let emailsSent = 0;

  meetingsByAttendee.forEach((meetings, email) => {
    try {
      const emailBody = buildMeetingEmailBody(
        email,
        meetings.new,
        meetings.changed,
        meetings.deleted,
        spreadsheetUrl
      );

      MailApp.sendEmail({
        to: email,
        subject: CONFIG.MEETINGS.EMAIL_SUBJECT,
        body: emailBody
      });

      emailsSent++;
      Logger.log(`Meeting notification email sent to: ${email}`);

    } catch (error) {
      Logger.log(`Failed to send meeting email to ${email}: ${error.toString()}`);
    }
  });

  return { emailsSent: emailsSent };
}

/**
 * Groups meetings by attendee email address
 * Filters out meetings with status "Completada" (finished meetings don't need notifications)
 * @param {Object} changes - Detected changes
 * @returns {Map} Map of email to categorized meetings
 */
function groupMeetingsByAttendee(changes) {
  const meetingsByAttendee = new Map();

  // Helper function to add meeting to attendee's list
  const addToAttendee = (email, category, meeting) => {
    if (!meetingsByAttendee.has(email)) {
      meetingsByAttendee.set(email, { new: [], changed: [], deleted: [] });
    }
    meetingsByAttendee.get(email)[category].push(meeting);
  };

  // Add new meetings (skip completed ones)
  changes.new.forEach(meeting => {
    if (meeting.status !== 'Completada') {
      addToAttendee(meeting.email, 'new', meeting);
    }
  });

  // Add changed meetings (skip completed ones)
  changes.changed.forEach(changeObj => {
    if (changeObj.meeting.status !== 'Completada') {
      addToAttendee(changeObj.meeting.email, 'changed', changeObj);
    }
  });

  // Add deleted meetings (use previous attendee)
  changes.deleted.forEach(meeting => {
    addToAttendee(meeting.email, 'deleted', meeting);
  });

  return meetingsByAttendee;
}

/**
 * Builds plain text email body with meeting change details
 * @param {string} attendeeEmail - Recipient email
 * @param {Array} newMeetings - New meetings
 * @param {Array} changedMeetings - Changed meetings with change details
 * @param {Array} deletedMeetings - Deleted meetings
 * @param {string} spreadsheetUrl - Spreadsheet URL
 * @returns {string} Formatted email body
 */
function buildMeetingEmailBody(attendeeEmail, newMeetings, changedMeetings, deletedMeetings, spreadsheetUrl) {
  const totalChanges = newMeetings.length + changedMeetings.length + deletedMeetings.length;
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
        } else if (change.field === 'DATETIME') {
          body += `  Fecha y hora: ${change.oldValue} → ${change.newValue} ⚠️\n`;
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
      if (!changes.find(c => c.field === 'DATETIME')) {
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
      body += `  → Esta reunión fue eliminada del calendario\n\n`;
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
