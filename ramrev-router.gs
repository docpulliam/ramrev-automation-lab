/**
 * RAM REV Command Center Router
 * Trigger: Command_Center!Status -> "Completed"
 * Routes row to Stripe Refunds or Month tab (Week sections)
 * Marks Command_Center row in Notes (AUTO-MOVED ...) to prevent duplicates
 */

const CONFIG = {
  // Sheet/tab names (must match exactly)
  COMMAND_CENTER: "Command_Center",
  ACTIVE_MONTH_SHEET_NAME: "Feb 26 Refund Req",
  STRIPE_SHEET_NAME: "Stripe Refunds",

  // Only run when Status becomes exactly "Completed" (case-insensitive)
  STATUS_DONE_VALUES: ["completed"],

  // If Refund Type contains these (case-insensitive), route to Stripe
  STRIPE_REFUND_TYPE_KEYWORDS: ["stripe", "return to stripe"],

  // Column headers in Command_Center (row 1)
  CC_HEADERS: {
    firstName: "First Name",
    lastName: "Last Name",
    memberId: "Member ID",
    email: "Email",
    phone: "Phone",
    rawAddress: "Raw Address",
    refundType: "Refund Type",
    status: "Status",
    notes: "Notes"
  },

  // Marker written into Notes to prevent double-processing
  NOTES_MOVED_MARKER: "AUTO-MOVED"
};

/**
 * Installable trigger: From spreadsheet -> On edit
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const ss = e.source;
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.COMMAND_CENTER) return;

    const headers = getHeaderMap_(sh);
    const statusCol = headers[CONFIG.CC_HEADERS.status];
    if (!statusCol) return;

    // only react to edits in Status column
    if (e.range.getColumn() !== statusCol) return;

    const row = e.range.getRow();
    if (row < 2) return;

    const newVal = String(e.value || "").trim().toLowerCase();
    if (!CONFIG.STATUS_DONE_VALUES.includes(newVal)) return;

    // Read row into object
    const rowObj = getRowObject_(sh, row);

    // Duplicate prevention: if already moved, do nothing
    const existingNotes = String(rowObj[CONFIG.CC_HEADERS.notes] || "");
    if (existingNotes.toUpperCase().includes(CONFIG.NOTES_MOVED_MARKER)) return;

    // Validate required fields
    const memberId = String(rowObj[CONFIG.CC_HEADERS.memberId] || "").trim();
    if (!memberId) throw new Error("Member ID is required before completing.");

    const address = String(rowObj[CONFIG.CC_HEADERS.rawAddress] || "").trim();
    if (!address) throw new Error("Raw Address is blank — paste address before completing.");

    const refundTypeRaw = String(rowObj[CONFIG.CC_HEADERS.refundType] || "").trim();
    const refundType = refundTypeRaw.toLowerCase();

    const isStripeByType = CONFIG.STRIPE_REFUND_TYPE_KEYWORDS.some(k => refundType.includes(k));
    const isCARFallback = memberId.toUpperCase().startsWith("CAR");

    if (isStripeByType || isCARFallback) {
      routeToStripe_(ss, rowObj);
    } else {
      routeToMonthWeek_(ss, rowObj);
    }

    // Mark Notes instead of changing Status (avoids data validation)
    const notesCol = headers[CONFIG.CC_HEADERS.notes];
    if (notesCol) {
      sh.getRange(row, notesCol).setValue(
        `${CONFIG.NOTES_MOVED_MARKER} ${new Date().toLocaleString()}`
      );
    }

  } catch (err) {
    console.error(err);
    SpreadsheetApp.getUi().alert("Automation stopped:\n\n" + err.message);
  }
}

/** ===== ROUTE: STRIPE ===== */
function routeToStripe_(ss, rowObj) {
  const sh = ss.getSheetByName(CONFIG.STRIPE_SHEET_NAME);
  if (!sh) throw new Error("Stripe sheet not found: " + CONFIG.STRIPE_SHEET_NAME);

  const headers = getHeaderMap_(sh);
  const today = new Date();

  const writeMap = {
    "First Name": rowObj[CONFIG.CC_HEADERS.firstName] || "",
    "Last Name": rowObj[CONFIG.CC_HEADERS.lastName] || "",
    "Member ID": rowObj[CONFIG.CC_HEADERS.memberId] || "",
    "Email": rowObj[CONFIG.CC_HEADERS.email] || "",
    "Phone": rowObj[CONFIG.CC_HEADERS.phone] || "",
    "Full Address": rowObj[CONFIG.CC_HEADERS.rawAddress] || "",
    "Date Requested": today,
    "Refund Type": "Stripe",
    "Completed?": "Complete"
  };

  // Optional columns if they exist
  if (headers["Date Processed"]) writeMap["Date Processed"] = today;
  if (headers["Date Complete"]) writeMap["Date Complete"] = today;

  appendRowByHeaderMap_(sh, writeMap);
}

/** ===== ROUTE: MONTH TAB -> WEEK SECTION ===== */
function routeToMonthWeek_(ss, rowObj) {
  const sh = ss.getSheetByName(CONFIG.ACTIVE_MONTH_SHEET_NAME);
  if (!sh) throw new Error("Month sheet not found: " + CONFIG.ACTIVE_MONTH_SHEET_NAME);

  const headers = getHeaderMap_(sh);
  const today = new Date();

  // Week placement based on TODAY (since you complete during call)
  const weekNum = computeWeekOfMonth_FirstFullWeek_(today);
  const weekLabel = `Week ${weekNum}`;

  const targetRow = findFirstEmptyRowUnderWeek_(sh, weekLabel);
  if (!targetRow) throw new Error("Could not find an open row under " + weekLabel);

  // Keep Refund Type from Command Center (e.g., Check) unless blank
  const refundType = String(rowObj[CONFIG.CC_HEADERS.refundType] || "").trim() || "Check";

  const writeMap = {
    "First Name": rowObj[CONFIG.CC_HEADERS.firstName] || "",
    "Last Name": rowObj[CONFIG.CC_HEADERS.lastName] || "",
    "Member ID": rowObj[CONFIG.CC_HEADERS.memberId] || "",
    "Email": rowObj[CONFIG.CC_HEADERS.email] || "",
    "Phone": rowObj[CONFIG.CC_HEADERS.phone] || "",
    "Full Address": rowObj[CONFIG.CC_HEADERS.rawAddress] || "",
    "Date Requested": today,
    "Refund Type": refundType
  };

  // Optional columns if they exist
  if (headers["Completed?"]) writeMap["Completed?"] = "Requested";
  if (headers["Date Processed"]) writeMap["Date Processed"] = today;

  writeRowByHeaderMapAt_(sh, targetRow, writeMap);
}

/** ===== HELPERS ===== */

function getHeaderMap_(sheet) {
  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headerRow.forEach((h, i) => {
    const key = String(h || "").trim();
    if (key) map[key] = i + 1; // 1-based
  });
  return map;
}

function getRowObject_(sheet, row) {
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(row, 1, 1, lastCol).getValues()[0];

  const obj = {};
  for (let i = 0; i < headerRow.length; i++) {
    const key = String(headerRow[i] || "").trim();
    if (key) obj[key] = values[i];
  }
  return obj;
}

function appendRowByHeaderMap_(sheet, writeMap) {
  const headers = getHeaderMap_(sheet);

  // Find first truly-empty row by checking column A downwards (safer for "table" formatting)
  const row = findFirstEmptyRowInColumn_(sheet, 1);

  Object.keys(writeMap).forEach(h => {
    const col = headers[h];
    if (!col) return; // skip missing optional headers
    sheet.getRange(row, col).setValue(writeMap[h]);
  });
}

function writeRowByHeaderMapAt_(sheet, row, writeMap) {
  const headers = getHeaderMap_(sheet);
  Object.keys(writeMap).forEach(h => {
    const col = headers[h];
    if (!col) return;
    sheet.getRange(row, col).setValue(writeMap[h]);
  });
}

function findFirstEmptyRowInColumn_(sheet, col) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  const vals = sheet.getRange(1, col, lastRow, 1).getValues().flat();
  for (let r = 2; r <= vals.length; r++) {
    if (String(vals[r - 1] || "").trim() === "") return r;
  }
  return lastRow + 1;
}

/**
 * Week 1 starts on first Monday that begins a full Mon–Sun week inside the month.
 */
function computeWeekOfMonth_FirstFullWeek_(dt) {
  const y = dt.getFullYear();
  const m = dt.getMonth();
  const firstDay = new Date(y, m, 1);
  const lastDay = new Date(y, m + 1, 0);

  const day = firstDay.getDay(); // Sun=0..Sat=6
  const offsetToMonday = (day === 0) ? 1 : (day <= 1 ? 1 - day : 8 - day);
  let firstMonday = new Date(y, m, 1 + offsetToMonday - 1);

  // Ensure it's a full week
  if (addDays_(firstMonday, 6) > lastDay) firstMonday = firstDay;

  if (stripTime_(dt) < stripTime_(firstMonday)) return 1;

  const diffDays = Math.floor((stripTime_(dt) - stripTime_(firstMonday)) / 86400000);
  return Math.floor(diffDays / 7) + 1;
}

function stripTime_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
}

function addDays_(d, n) {
  const x = new Date(d);
  x.setDate(x.getDate() + n);
  return x;
}

/**
 * Finds "Week X" in column A, then first blank row under it before next Week header.
 */
function findFirstEmptyRowUnderWeek_(sheet, weekLabel) {
  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues().flat().map(v => String(v || "").trim());

  const idx = colA.findIndex(v => v.toLowerCase() === weekLabel.toLowerCase());
  if (idx === -1) return null;

  for (let i = idx + 1; i < colA.length; i++) {
    const v = colA[i].toLowerCase();
    if (v.startsWith("week ")) break;

    // Treat blank row as open slot
    if (v === "") return i + 1;

    // Also treat template rows that literally say "First Name" as open
    if (v === "first name") return i + 1;
  }

  return lastRow + 1;
}
