// Shared config for registration Apps Script project.
// Keep Form / Spreadsheet IDs and sheet tab names here.

var REG_CONFIG = {
  FORM_ID: "1wmpmUTbWPKz-HoW43HQXeuTAXxbBHFcNNDi6Mjh1YV4",
  SPREADSHEET_ID: "1xQGr-yOPB_f5-mztIb9veFj3Ltu7K0HNTcSisV8KnS4",

  // 2026 spreadsheet uses the default Forms tab name first
  RESPONSES_SHEET_NAMES: [
    "Form Responses 1",
    "FormResponses",
    "Form Responses",
    "Responses"
  ],
  DIVISION_SHEET_NAMES: [
    "DivisionList",
    "Division List",
    "Divisions"
  ]
};

function getRegistrationSpreadsheet() {
  return SpreadsheetApp.openById(REG_CONFIG.SPREADSHEET_ID);
}

/**
 * Find a sheet by trying several names. Throws with available tab names if none match.
 */
function getSheetByKnownNames_(ss, candidates, label) {
  for (var i = 0; i < candidates.length; i++) {
    var sheet = ss.getSheetByName(candidates[i]);
    if (sheet) {
      return sheet;
    }
  }

  var available = ss.getSheets().map(function (s) {
    return s.getName();
  }).join(", ");

  throw new Error(
    label + " sheet not found. Tried: [" + candidates.join(", ") + "]. " +
    "Available tabs: [" + available + "]. " +
    "Update REG_CONFIG in config.gs, or run createDivisionListTemplate() if this is the division sheet."
  );
}

function getFormResponsesSheet(ss) {
  ss = ss || getRegistrationSpreadsheet();
  return getSheetByKnownNames_(ss, REG_CONFIG.RESPONSES_SHEET_NAMES, "Form responses");
}

function getDivisionListSheet(ss) {
  ss = ss || getRegistrationSpreadsheet();
  return getSheetByKnownNames_(ss, REG_CONFIG.DIVISION_SHEET_NAMES, "Division list");
}

/**
 * One-time helper: creates an empty DivisionList tab with headers.
 * Then paste/copy your division rows from the 2025 spreadsheet (columns A–E).
 *
 * A: Division code | B: Age range | C: Gender | D: Ranking | E: Event (Kata|Kumite)
 */
function createDivisionListTemplate() {
  var ss = getRegistrationSpreadsheet();
  var existing = null;
  for (var i = 0; i < REG_CONFIG.DIVISION_SHEET_NAMES.length; i++) {
    existing = ss.getSheetByName(REG_CONFIG.DIVISION_SHEET_NAMES[i]);
    if (existing) break;
  }

  if (existing) {
    Logger.log('Division sheet already exists: "' + existing.getName() + '"');
    return existing;
  }

  var sheet = ss.insertSheet("DivisionList");
  sheet.getRange(1, 1, 1, 5).setValues([[
    "DivisionCode",
    "AgeRange",
    "Gender",
    "Ranking",
    "Event"
  ]]);
  sheet.setFrozenRows(1);
  Logger.log(
    "Created DivisionList with headers. " +
    "Copy your division data rows from the 2025 spreadsheet into columns A–E, then run assignDivisionCode()."
  );
  return sheet;
}
