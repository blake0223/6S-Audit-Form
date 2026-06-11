/**
 * AddRoomColumn — inserts an individual scoring column for a new room,
 * immediately to the right of the questions column. The question column is found
 * by HEADER TEXT (not a fixed letter) so it keeps working as columns shift right.
 *
 * It sets the new column header to the room name in a unique rotating color.
 * No data validation and no totals are added — each room column is simply its own
 * 0–3 scoring column (the 0–3 rule is enforced elsewhere).
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

  var headerRow   = findHeaderRow_(sheet, sheet.getLastRow());
  var questionCol = findQuestionCol_(sheet, headerRow) || QUESTION_COL;

  // New blank scoring column right of the questions column.
  sheet.insertColumnsAfter(questionCol, 1);
  var newCol = questionCol + 1;

  // Header = room name, in a unique rotating color.
  var color = nextRoomColor_();
  sheet.getRange(headerRow, newCol)
    .setValue(roomName)
    .setFontWeight('bold')
    .setFontColor(color.font)
    .setBackground(color.bg)
    .setHorizontalAlignment('center')
    .setWrap(true);

  sheet.setColumnWidth(newCol, 90);
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
