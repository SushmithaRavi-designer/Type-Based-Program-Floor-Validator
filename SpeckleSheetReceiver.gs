/**
 * SPECKLE AUTOMATE → GOOGLE SHEETS RECEIVER
 * ==========================================
 * Deploy this as a Web App (one-time setup) and paste the URL into
 * Speckle Automate Function Settings → "Google Apps Script Web App URL"
 *
 * SETUP STEPS:
 *  1. Go to https://script.google.com → New project
 *  2. Paste this entire file, replacing the default code
 *  3. Click Deploy → New deployment
 *  4. Type: Web app
 *  5. Execute as: Me
 *  6. Who has access: Anyone   ← required so Automate can POST to it
 *  7. Click Deploy → copy the Web App URL
 *  8. Paste that URL into Speckle Function Settings → "Google Apps Script Web App URL"
 *
 * Sheet 1 "Analysis"        — Level/Program/Area/Status per occupancy
 * Sheet 2 "Occupancy Timing" — Area at each time band per occupancy (from formula)
 */

// ── Column definitions ───────────────────────────────────────────────────────

var ANALYSIS_COLUMNS = [
  "Occupancy", "Level", "Program", "Area", "Status",
  "Area_OffPeak", "Area_Morning", "Area_Afternoon", "Area_Evening",
];

var TIMING_COLUMNS = [
  "Occupancy", "Time Band", "Time Range (s)", "Example Clock", "Area (mm)", "Area (m²)",
];

// ── POST handler ─────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    var payload    = JSON.parse(e.postData.contents);
    var sheetTitle = payload.sheetTitle  || "Program_Floor_Analysis";
    var rows       = payload.rows        || [];
    var timingRows = payload.timingRows  || [];

    var spreadsheet = _getOrCreateSpreadsheet(sheetTitle);

    // ── Sheet 1: Analysis ─────────────────────────────────────────────────
    _writeSheet(spreadsheet, "Analysis", ANALYSIS_COLUMNS, rows, "#1a73e8");

    // ── Sheet 2: Occupancy Timing ─────────────────────────────────────────
    if (timingRows.length > 1) {   // >1 because first row is the header dict
      // Strip the header row (it's a dict row with column names as values)
      var dataRows = timingRows.slice(1);
      _writeSheet(spreadsheet, "Occupancy Timing", TIMING_COLUMNS, dataRows, "#0f9d58");
    }

    // Move Analysis sheet to front
    var analysisSheet = spreadsheet.getSheetByName("Analysis");
    if (analysisSheet) spreadsheet.setActiveSheet(analysisSheet);
    spreadsheet.moveActiveSheet(1);

    return _jsonResponse({ status: "ok", sheetUrl: spreadsheet.getUrl() });

  } catch (err) {
    return _jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── GET: health check ────────────────────────────────────────────────────────
function doGet(e) {
  return _jsonResponse({ status: "ok", message: "Speckle Sheet Receiver is running." });
}

// ── Generic sheet writer ──────────────────────────────────────────────────────
function _writeSheet(spreadsheet, sheetName, columns, rows, headerColor) {
  var sheet = _getOrCreateSheet(spreadsheet, sheetName);
  sheet.clearContents();
  sheet.clearFormats();

  if (!rows || rows.length === 0) {
    sheet.getRange(1, 1).setValue("No data.");
    return;
  }

  // Header row
  sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  var hRange = sheet.getRange(1, 1, 1, columns.length);
  hRange.setBackground(headerColor);
  hRange.setFontColor("#ffffff");
  hRange.setFontWeight("bold");
  sheet.setFrozenRows(1);

  // Data rows — skip empty separator dicts
  var dataValues = rows
    .filter(function(row) { return row && Object.keys(row).length > 0; })
    .map(function(row) {
      return columns.map(function(col) {
        var v = row[col];
        return (v === undefined || v === null) ? "" : v;
      });
    });

  if (dataValues.length > 0) {
    sheet.getRange(2, 1, dataValues.length, columns.length).setValues(dataValues);
  }

  // Alternating row colors + highlight summary rows
  for (var i = 0; i < dataValues.length; i++) {
    var rowNum  = i + 2;
    var firstCell = dataValues[i][0];   // Occupancy column value
    var range   = sheet.getRange(rowNum, 1, 1, columns.length);

    if (firstCell === "SUMMARY" || firstCell === "Total Area (m²)" ||
        firstCell === "OK Entries" || firstCell === "MONO-FUNCTIONAL Entries") {
      range.setBackground("#fff3cd");
      range.setFontWeight("bold");
    } else {
      range.setBackground(i % 2 === 0 ? "#f8f9fa" : "#ffffff");
    }
  }

  // Auto-resize columns
  for (var c = 1; c <= columns.length; c++) {
    sheet.autoResizeColumn(c);
  }

  // Timestamp note
  sheet.getRange(1, columns.length + 2)
       .setValue("Updated: " + new Date().toLocaleString());
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function _getOrCreateSpreadsheet(title) {
  var files = DriveApp.getFilesByName(title);
  while (files.hasNext()) {
    var file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return SpreadsheetApp.openById(file.getId());
    }
  }
  return SpreadsheetApp.create(title);
}

function _getOrCreateSheet(spreadsheet, sheetName) {
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  return sheet;
}

function _jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
