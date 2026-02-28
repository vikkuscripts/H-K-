// Full Code.gs — Daily-reset + fetch/update + migration helpers
const SHEET_ID = "18k583-5dNp6JdslfJaPVmoVGijsxrinyI5aUKnPQTYI"; // replace if needed
const SHEET_ROOMS = "Rooms";
const SHEET_AREA  = "Area";
const SHEET_STAFF = "Staff_Sheet";
const LAST_RESET_PROP = 'lastResetDate_v1'; // script property key

function doGet(e) {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Housekeeping - Hotel Room Management')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Top-level data fetch called by the frontend.
 * Ensures daily reset runs automatically once per day (fallback)
 * before returning dataset.
 */
function getInitialData() {
  // ensure reset runs once per day (fallback if trigger not set)
  try {
    ensureDailyReset();
  } catch (err) {
    Logger.log('ensureDailyReset error: ' + (err && err.toString()));
    // continue; we still return data even if reset check fails
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);

  const staffData = fetchStaff(ss);
  const areaData = fetchAreas(ss);
  const roomData = fetchRooms(ss);

  const counts = {
    totalRooms: roomData.length,
    dirty: roomData.filter(r => r.status === 'Dirty').length,
    inProgress: roomData.filter(r => r.status === 'In Progress').length,
    clean: roomData.filter(r => r.status === 'Clean').length,

    totalAreas: areaData.length,
    areaDirty: areaData.filter(a => a.status === 'Dirty').length,
    areaInProgress: areaData.filter(a => a.status === 'In Progress').length,
    areaClean: areaData.filter(a => a.status === 'Clean').length,
  };

  return {
    staff: staffData,
    areas: areaData,
    rooms: roomData,
    counts: counts
  };
}

function dailyResetToDirty() {
  Logger.log('dailyResetToDirty started at ' + new Date());
  const ss = SpreadsheetApp.openById(SHEET_ID);
  resetRooms(ss);
  resetAreas(ss);

  // persist last reset date so fallback won't re-run same day
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  PropertiesService.getScriptProperties().setProperty(LAST_RESET_PROP, today);
  Logger.log('dailyResetToDirty completed at ' + new Date());
}

/**
 * Ensure reset runs once per calendar day.
 * This is a safe fallback run inside getInitialData() so the app works
 * even if the trigger wasn't created or had issues.
 */
function ensureDailyReset() {
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const props = PropertiesService.getScriptProperties();
  const last = props.getProperty(LAST_RESET_PROP);

  if (last === today) {
    // already ran today
    return;
  }

  // run reset now and store date
  dailyResetToDirty();
  // dailyResetToDirty() writes LAST_RESET_PROP itself
}

function resetRooms(ss) {
  const sheet = ss.getSheetByName(SHEET_ROOMS);
  if (!sheet) {
    Logger.log('resetRooms: Rooms sheet not found');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Read all column A to check floor rows quickly
  const colA = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    const rowIndex = i + 2; // actual sheet row
    const a = String(colA[i][0] || '').toLowerCase();
    if (a.indexOf('floor') !== -1) continue; // skip header-like rows
    sheet.getRange(rowIndex, 2).setValue('Dirty'); // B
    sheet.getRange(rowIndex, 3).clearContent();   // C Assigned
    sheet.getRange(rowIndex, 4).clearContent();   // D TimeIn
    sheet.getRange(rowIndex, 5).clearContent();   // E TimeOut
  }
}


function resetAreas(ss) {
  const sheet = ss.getSheetByName(SHEET_AREA);
  if (!sheet) {
    Logger.log('resetAreas: Area sheet not found');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    const rowIndex = i + 2;
    const nm = String(names[i][0] || '').trim();
    if (!nm) continue;
    sheet.getRange(rowIndex, 2).setValue('Dirty'); // C
    sheet.getRange(rowIndex, 3).clearContent();   // D assignedTo
    sheet.getRange(rowIndex, 4).clearContent();   // F timeIn
    sheet.getRange(rowIndex, 5).clearContent();   // G timeOut
  }
}

/* -------------
   Trigger helpers
   ------------- */

/**
 * Create an installable daily trigger for dailyResetToDirty().
 * Run this manually once from Apps Script editor: select function createDailyTrigger and click Run.
 * It will remove any duplicate triggers created previously for the same function.
 */
function createDailyTrigger() {
  // Remove existing triggers for dailyResetToDirty to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'dailyResetToDirty') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create a time-driven trigger that runs daily at 01:00 (1 AM)
  // You can change the hour as you prefer in the UI triggers too.
  ScriptApp.newTrigger('dailyResetToDirty')
    .timeBased()
    .everyDays(1)
    .atHour(1)   // server timezone; adjust if you want another hour
    .create();

  Logger.log('createDailyTrigger: daily trigger created (1 AM)');
}

/**
 * Manual admin function to run reset now.
 * Use from Apps Script editor -> select manualResetNow -> Run
 */
function manualResetNow() {
  dailyResetToDirty();
  Logger.log('manualResetNow executed at ' + new Date());
}

/* ---------------------------
   Fetch helpers (rooms/areas/staff)
   --------------------------- */

function fetchStaff(ss) {
  const sheet = ss.getSheetByName(SHEET_STAFF);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const staff = [];
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    const role = String(values[i][1] || '').trim();
    if (name) staff.push({ name: name, role: role });
  }
  return staff;
}

/*
  fetchAreas: read the Area sheet properly.
  Column map (must match updateArea below):
    A (0) => name
    B (1) => location
    C (2) => status
    D (3) => assignedTo
    E (4) => optional
    F (5) => timeIn
    G (6) => timeOut
  id returned is sheet row number (so updateArea(row) writes to same row).
*/
function fetchAreas(ss) {
  const sheet = ss.getSheetByName(SHEET_AREA);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const areas = [];
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][0] || '').trim();
    if (!name) continue;

    const location = String(values[i][1] || '').trim() || 'Unassigned';
    const status = String(values[i][2] || 'Dirty').trim() || 'Dirty';
    const assignedTo = String(values[i][3] || '').trim();
    const timeIn = formatTimeFromSheet(values[i][5]);   // F
    const timeOut = formatTimeFromSheet(values[i][6]);  // G

    areas.push({
      id: i + 1,
      name: name,
      location: location,
      status: status,
      assignedTo: assignedTo,
      timeIn: timeIn,
      timeOut: timeOut
    });
  }
  return areas;
}

function fetchRooms(ss) {
  const sheet = ss.getSheetByName(SHEET_ROOMS);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  const rooms = [];
  let currentFloor = "General";

  // Col A: "Floor X" OR "Room no"
  // Col B: Status | Col C: Assigned To | Col D: Time In | Col E: Time Out
  for (let i = 1; i < values.length; i++) {
    const colA = String(values[i][0] || '').trim();
    if (!colA) continue;

    if (colA.toLowerCase().indexOf('floor') !== -1) {
      currentFloor = colA;
      continue;
    }

    rooms.push({
      id: i + 1,
      roomNumber: colA,
      floor: currentFloor,
      status: String(values[i][1] || 'Dirty').trim(),
      assignedTo: String(values[i][2] || '').trim(),
      timeIn: formatTimeFromSheet(values[i][3]),
      timeOut: formatTimeFromSheet(values[i][4])
    });
  }
  return rooms;
}

function formatTimeFromSheet(val) {
  if (!val) return '';
  if (val instanceof Date) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "M/d/yyyy, h:mm:ss a");
    } catch(e) {
      return String(val);
    }
  }
  return String(val).trim();
}

/* ---------------------------
   Update handlers called by frontend
   --------------------------- */

function updateRoom(data) {
  Logger.log('updateRoom payload: %s', JSON.stringify(data));
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_ROOMS);
  if (!sheet) throw new Error("Rooms sheet not found.");

  const row = Number(data.id);
  if (!row || row <= 1) throw new Error("Invalid row id for room: " + data.id);

  sheet.getRange(row, 2).setValue(data.status || 'Dirty');

  if (data.assignedTo !== undefined) sheet.getRange(row, 3).setValue(data.assignedTo);

  if (data.setTimeIn) {
    const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy, h:mm:ss a");
    sheet.getRange(row, 4).setValue(timeStr);
    sheet.getRange(row, 5).clearContent();
  }

  if (data.setTimeOut) {
    const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy, h:mm:ss a");
    sheet.getRange(row, 5).setValue(timeStr);
  }

  if (data.reset) {
    sheet.getRange(row, 3).clearContent();
    sheet.getRange(row, 4).clearContent();
    sheet.getRange(row, 5).clearContent();
  }

  return getInitialData();
}

function updateArea(data) {
  Logger.log('updateArea payload: %s', JSON.stringify(data));
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_AREA);
  if (!sheet) throw new Error("Area sheet not found.");

  const row = Number(data.id);
  if (!row || row <= 1) throw new Error("Invalid row id for area: " + data.id);

  // Write status to column C (3)
  sheet.getRange(row, 3).setValue(data.status || 'Dirty');

  if (data.assignedTo !== undefined) sheet.getRange(row, 4).setValue(data.assignedTo);

  if (data.setTimeIn) {
    const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy, h:mm:ss a");
    sheet.getRange(row, 6).setValue(timeStr); // F
    sheet.getRange(row, 7).clearContent();    // G cleared on restart
  }

  if (data.setTimeOut) {
    const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy, h:mm:ss a");
    sheet.getRange(row, 7).setValue(timeStr); // G
  }

  if (data.reset) {
    sheet.getRange(row, 4).clearContent(); // D assignee
    sheet.getRange(row, 6).clearContent(); // F timeIn
    sheet.getRange(row, 7).clearContent(); // G timeOut
  }

  return getInitialData();
}

/* ---------------------------
   Migration / Repair helpers
   --------------------------- */

/**
 * Quick repair: if some time strings ended up in column E (5) or other
 * wrong columns, this helper will move them to proper F (6) / G (7)
 * if the destination column is empty. This is safe and logs changes.
 *
 * Run manually once from Apps Script editor if you see times in wrong column.
 */
function migrateAreaTimesIfNeeded() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_AREA);
  if (!sheet) throw new Error("Area sheet not found.");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'No rows to check';

  const range = sheet.getRange(2, 1, lastRow - 1, 8);
  const values = range.getValues();
  let moved = 0;

  for (let i = 0; i < values.length; i++) {
    const rowIndex = i + 2;
    const colE = String(values[i][4] || '').trim(); // E
    const colF = String(values[i][5] || '').trim(); // F
    const colG = String(values[i][6] || '').trim(); // G

    // Simple heuristic: if E looks like a datetime and F is empty -> move E->F
    if (colE && looksLikeDateTime(colE) && !colF) {
      sheet.getRange(rowIndex, 6).setValue(colE);
      sheet.getRange(rowIndex, 5).clearContent();
      moved++;
      Logger.log(`migrateAreaTimesIfNeeded: moved E->F at row ${rowIndex}`);
    }

    // If F contains two datetimes separated by " - " or newline, try to split
    if (colF && !colG) {
      // check if contains " - " or newline as a simple separator
      if (colF.indexOf(' - ') !== -1) {
        const parts = colF.split(' - ').map(s => s.trim());
        if (looksLikeDateTime(parts[0]) && looksLikeDateTime(parts[1])) {
          sheet.getRange(rowIndex, 6).setValue(parts[0]);
          sheet.getRange(rowIndex, 7).setValue(parts[1]);
          moved++;
          Logger.log(`migrateAreaTimesIfNeeded: split F->F&G at row ${rowIndex}`);
        }
      }
    }
  }

  return 'migrateAreaTimesIfNeeded completed, moved: ' + moved;
}

function looksLikeDateTime(s) {
  if (!s) return false;
  // simple check for pattern like "M/d/yyyy" or "yyyy" etc.
  return /[0-9]{1,2}\/[0-9]{1,2}\/[0-9]{4}/.test(s) || /[0-9]{4}-[0-9]{2}-[0-9]{2}/.test(s);
}
