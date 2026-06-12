/**
 * AddRoomColumn — inserts a scoring column for a new room at the RIGHT end of the
 * room block (just left of the Total Score column). Inserting there is INSIDE the
 * "Audit Location Scores" merged band, so the band extends over the new column —
 * inserting at the far-left edge of a merge would leave the new column outside it.
 *
 * The neighbouring room column is copied wholesale so the new room inherits the
 * merges, the Score % formula in the band, and all formatting; the per-item score
 * cells are then blanked. The header gets the room name in a unique rotating color.
 *
 * Menu: 6S Audit Tools ▸ Add room column   (wired up in onOpen.gs)
 */

var AUDIT_SHEET_NAME = '6S Audit Sheet'; // the tab these tools act on
var QUESTION_COL     = 3;                // fallback if the question header isn't found

function addRoomColumn() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUDIT_SHEET_NAME) || ss.getActiveSheet();

  var resp = ui.prompt('Add room', 'Name of the room / area to add:', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  var roomName = resp.getResponseText().trim();
  if (!roomName) { ui.alert('No room name entered — nothing added.'); return; }

  var lastRow = sheet.getLastRow();
  var headerRow = findHeaderRow_(sheet, lastRow);
  var qCol = findQuestionCol_(sheet, headerRow) || QUESTION_COL;
  var totalCol = findColByHeader_(sheet, headerRow, 'total score');
  if (!totalCol || totalCol <= qCol + 1) {
    ui.alert('Could not find the room block — need a "Total Score" column to the right of the rooms.');
    return;
  }

  // Insert just LEFT of the Total Score column (inside the merged band so it extends).
  sheet.insertColumnsBefore(totalCol, 1);
  var newCol = totalCol;     // the inserted column
  var srcCol = newCol - 1;   // previous right-most room — copied as the template

  // Copy the neighbour room column wholesale (format + Score % formula + merges),
  // then blank the per-item score cells so the new room starts empty.
  sheet.getRange(1, srcCol, lastRow, 1).copyTo(sheet.getRange(1, newCol, lastRow, 1));
  sheet.setColumnWidth(newCol, sheet.getColumnWidth(srcCol));

  var colA = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, 1).getValues();
  for (var i = 0; i < colA.length; i++) {
    if (isItemNumber_(colA[i][0])) {
      sheet.getRange(headerRow + 1 + i, newCol).clearContent();
    }
  }

  // Name + colour the header.
  var color = nextRoomColor_();
  sheet.getRange(headerRow, newCol)
    .setValue(roomName)
    .setFontWeight('bold')
    .setFontColor(color.font)
    .setBackground(color.bg)
    .setHorizontalAlignment('center')
    .setWrap(true);

  rebuildRoomTotals_(sheet);
  ui.alert('Added room column "' + roomName + '".');
}

/* ---------- helpers ---------- */

/** Header row = the one with "No." in col A or "Check Item" in col B. */
function findHeaderRow_(sheet, lastRow) {
  var max = Math.min(lastRow, 30);
  var vals = sheet.getRange(1, 1, max, 3).getValues();
  for (var i = 0; i < vals.length; i++) {
    var a = String(vals[i][0]).trim().toLowerCase();
    var b = String(vals[i][1]).trim().toLowerCase();
    if (a === 'no.' || b === 'check item') return i + 1;
  }
  return 7; // default for this template
}

/** True if a column-A value is an item number — numeric (1) or "1.0"-style text. */
function isItemNumber_(v) {
  if (typeof v === 'number') return v > 0;
  var s = String(v).trim();
  return /^\d+(\.\d+)?$/.test(s) && parseFloat(s) > 0;
}

/** Find the question column by its header text ("description" / "audit question"). */
function findQuestionCol_(sheet, headerRow) {
  var lastCol = sheet.getLastColumn();
  var hdr = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  for (var c = 0; c < hdr.length; c++) {
    var t = String(hdr[c]).toLowerCase();
    if (t.indexOf('description') >= 0 || t.indexOf('audit question') >= 0) return c + 1;
  }
  return 0;
}

/**
 * Next room header color, rotating through a fixed palette so consecutive rooms
 * get distinct colors. Position is remembered per document.
 */
function nextRoomColor_() {
  var palette = [
    { bg: '#1F4E79', font: '#FFFFFF' }, // blue
    { bg: '#2E7D32', font: '#FFFFFF' }, // green
    { bg: '#8E44AD', font: '#FFFFFF' }, // purple
    { bg: '#C0392B', font: '#FFFFFF' }, // red
    { bg: '#D68910', font: '#FFFFFF' }, // amber
    { bg: '#16A085', font: '#FFFFFF' }, // teal
    { bg: '#AD1457', font: '#FFFFFF' }, // magenta
    { bg: '#2C3E50', font: '#FFFFFF' }  // slate
  ];
  var props = PropertiesService.getDocumentProperties();
  var i = parseInt(props.getProperty('roomColorIndex') || '0', 10);
  props.setProperty('roomColorIndex', String(i + 1));
  return palette[i % palette.length];
}
