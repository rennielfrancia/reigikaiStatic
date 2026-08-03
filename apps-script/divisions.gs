// Division code assignment (Google Apps Script)
// Maps FormResponses rows to DivisionList codes by age, gender, ranking, and event.
// Writes:
//   - Combined division codes in Q (original)
//   - Kata-only codes in Y
//   - Kumite-only codes in Z
// Payment plugin fields R–X are left untouched.
// Not loaded by the static website — kept here for version control / edits.
//
// Sheets: FormResponses, DivisionList
// Spreadsheet ID: 1xQGr-yOPB_f5-mztIb9veFj3Ltu7K0HNTcSisV8KnS4
//
// DivisionList columns A–E: Code | Age range | Gender | Ranking | Event (Kata|Kumite)

var DIVISION_COLS = {
  GENDER: 3,             // C
  AGE: 4,                // D
  RANKING: 5,            // E
  EVENT: 12,             // L
  DIVISION_COMBINED: 17, // Q — original combined codes (before payment block R–X)
  // R–X reserved for payment plugin — do not write here
  DIVISION_KATA: 25,     // Y — Kata only
  DIVISION_KUMITE: 26    // Z — Kumite only
};

function assignDivisionCode() {
  var ss = SpreadsheetApp.openById("1xQGr-yOPB_f5-mztIb9veFj3Ltu7K0HNTcSisV8KnS4");
  var sheet = ss.getSheetByName("FormResponses");
  var divisionSheet = ss.getSheetByName("DivisionList");
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No response rows to process.");
    return;
  }

  var divisionData = divisionSheet.getRange(1, 1, divisionSheet.getLastRow(), 5).getValues();
  var c = DIVISION_COLS;

  for (var i = 2; i <= lastRow; i++) {
    var existingCombined = sheet.getRange(i, c.DIVISION_COMBINED).getValue();
    var existingKata = sheet.getRange(i, c.DIVISION_KATA).getValue();
    var existingKumite = sheet.getRange(i, c.DIVISION_KUMITE).getValue();

    var ageInput = sheet.getRange(i, c.AGE).getValue();
    var gender = sheet.getRange(i, c.GENDER).getValue();
    var rankingInput = sheet.getRange(i, c.RANKING).getValue();
    var eventInput = sheet.getRange(i, c.EVENT).getValue();

    var eventTypes = parseEventTypes(eventInput);
    if (eventTypes.length === 0) {
      Logger.log("Row " + i + ": no recognizable event type, skipping.");
      continue;
    }

    // Skip when combined exists and every requested event column is already filled
    var needsKata = eventTypes.indexOf("Kata") !== -1;
    var needsKumite = eventTypes.indexOf("Kumite") !== -1;
    var kataDone = !needsKata || !!existingKata;
    var kumiteDone = !needsKumite || !!existingKumite;
    if (existingCombined && kataDone && kumiteDone) {
      continue;
    }

    var ageRange = resolveAgeRange(ageInput);
    var rankingLevel = resolveRankingLevel(rankingInput);

    Logger.log("Row " + i + " | rank=" + rankingLevel + " age=" + ageRange + " events=" + eventTypes.join(","));

    var kataCodes = [];
    var kumiteCodes = [];

    for (var e = 0; e < eventTypes.length; e++) {
      var eventType = eventTypes[e];
      for (var j = 0; j < divisionData.length; j++) {
        if (
          divisionData[j][1].toString() == ageRange &&
          divisionData[j][2] == gender &&
          divisionData[j][3] == rankingLevel &&
          divisionData[j][4] == eventType
        ) {
          if (eventType === "Kata") {
            kataCodes.push(divisionData[j][0]);
          } else if (eventType === "Kumite") {
            kumiteCodes.push(divisionData[j][0]);
          }
        }
      }
    }

    var combinedCodes = kataCodes.concat(kumiteCodes);

    if (!existingCombined) {
      sheet.getRange(i, c.DIVISION_COMBINED).setValue(combinedCodes.join(", "));
    }
    if (needsKata && !existingKata) {
      sheet.getRange(i, c.DIVISION_KATA).setValue(kataCodes.join(", "));
    }
    if (needsKumite && !existingKumite) {
      sheet.getRange(i, c.DIVISION_KUMITE).setValue(kumiteCodes.join(", "));
    }

    // Clear the unused event column only when we are assigning fresh for that row's events
    // (leave the other sport blank so filters stay clean)
    if (needsKata && !needsKumite && !existingKumite) {
      sheet.getRange(i, c.DIVISION_KUMITE).setValue("");
    }
    if (needsKumite && !needsKata && !existingKata) {
      sheet.getRange(i, c.DIVISION_KATA).setValue("");
    }
  }

  Logger.log("Division codes assigned (combined + Kata + Kumite columns).");
}

/** Detect Kata / Kumite from the form event label (works with old $ and 2026 pricing text). */
function parseEventTypes(eventInput) {
  var text = (eventInput || "").toString();
  var eventTypes = [];

  if (/kata\s*\+\s*kumite/i.test(text) || /both/i.test(text)) {
    eventTypes.push("Kata");
    eventTypes.push("Kumite");
    return eventTypes;
  }

  // Prefer explicit "only" labels when present
  var kataOnly = /kata\s*only/i.test(text);
  var kumiteOnly = /kumite\s*only/i.test(text);

  if (kataOnly) eventTypes.push("Kata");
  if (kumiteOnly) eventTypes.push("Kumite");

  if (eventTypes.length > 0) return eventTypes;

  // Fallback: bare mentions
  if (/kata/i.test(text)) eventTypes.push("Kata");
  if (/kumite/i.test(text)) eventTypes.push("Kumite");

  return eventTypes;
}

function resolveAgeRange(ageInput) {
  var raw = (ageInput || "").toString().toLowerCase().trim();

  if (raw === "5 and below") return "5 and below";
  if (raw === "18 - 34" || raw === "18-34") return "18-34";
  if (raw === "35+") return "35+";

  var age = parseInt(raw, 10);
  if (isNaN(age)) return "";

  if (age >= 6 && age <= 7) return "6-7";
  if (age >= 8 && age <= 9) return "8-9";
  if (age >= 10 && age <= 11) return "10-11";
  if (age >= 12 && age <= 13) return "12-13";
  if (age >= 14 && age <= 15) return "14-15";
  if (age >= 16 && age <= 17) return "16-17";
  return "";
}

function resolveRankingLevel(rankingInput) {
  var text = (rankingInput || "").toString().toLowerCase();
  if (text.indexOf("beginner") !== -1) return "Beginner";
  if (text.indexOf("novice") !== -1) return "Novice";
  if (text.indexOf("intermediate") !== -1) return "Intermediate";
  if (text.indexOf("advanced") !== -1) return "Advanced";
  return "";
}
