// Shared config for registration Apps Script project.
// Keep Form / Spreadsheet IDs and sheet tab names here.

var REG_CONFIG = {
  FORM_ID: "1wmpmUTbWPKz-HoW43HQXeuTAXxbBHFcNNDi6Mjh1YV4",
  SPREADSHEET_ID: "1xQGr-yOPB_f5-mztIb9veFj3Ltu7K0HNTcSisV8KnS4",

  // Preferred tab names, then common Google Forms defaults
  RESPONSES_SHEET_NAMES: [
    "FormResponses",
    "Form Responses 1",
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
    "Update REG_CONFIG in config.gs to match your tab name."
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
