// Form dropdown maintenance (Google Apps Script)
// Keeps Dojo Name / Instructor Name list items in sync with form responses,
// and replaces "Other" selections with free-text entries.
// Requires config.gs in the same Apps Script project.
// Not loaded by the static website — kept here for version control / edits.

function updateDojoNameDropdown() {
  var formId = REG_CONFIG.FORM_ID;
  var sheet = getFormResponsesSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No response rows yet — skipping dojo dropdown update.");
    return;
  }

  var dropdownColumn = 6;  // Column F: Dojo Name dropdown selections
  var newEntryColumn = 7;  // Column G: Free-text dojo name entries
  var numRows = lastRow - 1;

  var dropdownValues = sheet.getRange(2, dropdownColumn, numRows, 1).getValues().flat();
  var newEntries = sheet.getRange(2, newEntryColumn, numRows, 1).getValues().flat().filter(String).map(toPascalCase);

  var form = FormApp.openById(formId);
  var formItems = form.getItems(FormApp.ItemType.LIST);
  var listItem = formItems.find(item => item.getTitle().includes("Dojo Name"))?.asListItem();

  if (!listItem) {
    Logger.log("List item with title containing 'Dojo Name' not found.");
    return;
  }

  var existingChoices = listItem.getChoices().map(choice => choice.getValue());
  var combinedValues = [...new Set(existingChoices.concat(dropdownValues, newEntries))];
  var sortedValues = combinedValues.filter(v => v !== "Other").sort();
  sortedValues.push("Other");

  if (sortedValues.join() !== existingChoices.join()) {
    listItem.setChoiceValues(sortedValues);
    Logger.log("Dropdown dojo name updated with new entries.");
  } else {
    Logger.log("No new dojo names to add.");
  }
}

function updateInstructorNameDropdown() {
  var formId = REG_CONFIG.FORM_ID;
  var sheet = getFormResponsesSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("No response rows yet — skipping instructor dropdown update.");
    return;
  }

  var dropdownColumn = 8;  // Column H: Instructor Name dropdown selections
  var newEntryColumn = 9;  // Column I: Free-text instructor name entries
  var numRows = lastRow - 1;

  var dropdownValues = sheet.getRange(2, dropdownColumn, numRows, 1).getValues().flat();
  var newEntries = sheet.getRange(2, newEntryColumn, numRows, 1).getValues().flat().filter(String).map(toPascalCase);

  var form = FormApp.openById(formId);
  var formItems = form.getItems(FormApp.ItemType.LIST);
  var listItem = formItems.find(item => item.getTitle().includes("Instructor Name"))?.asListItem();

  if (!listItem) {
    Logger.log("List item with title containing 'Instructor Name' not found.");
    return;
  }

  var existingChoices = listItem.getChoices().map(choice => choice.getValue());
  var combinedValues = [...new Set(existingChoices.concat(dropdownValues, newEntries))];
  var sortedValues = combinedValues.filter(v => v !== "Other").sort();
  sortedValues.push("Other");

  if (sortedValues.join() !== existingChoices.join()) {
    listItem.setChoiceValues(sortedValues);
    Logger.log("Dropdown instructor name updated with new entries.");
  } else {
    Logger.log("No new instructor names to add.");
  }
}

function toPascalCase(str) {
  return String(str)
    .toLowerCase()
    .split(/[\s_-]+/)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1))
    .join(" ");
}

function replaceDojoWithUserInput() {
  var sheet = getFormResponsesSheet();
  var dropdownColumn = 6;
  var otherColumn = 7;
  var lastRow = sheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var dropdownValue = sheet.getRange(i, dropdownColumn).getValue();
    var otherValue = sheet.getRange(i, otherColumn).getValue();

    if (dropdownValue === "Other" && otherValue) {
      sheet.getRange(i, dropdownColumn).setValue(otherValue);
    }
  }

  Logger.log("Dropdown dojo column updated with user-entered values!");
}

function replaceInstructorWithUserInput() {
  var sheet = getFormResponsesSheet();
  var dropdownColumn = 8;
  var otherColumn = 9;
  var lastRow = sheet.getLastRow();

  for (var i = 2; i <= lastRow; i++) {
    var dropdownValue = sheet.getRange(i, dropdownColumn).getValue();
    var otherValue = sheet.getRange(i, otherColumn).getValue();

    if (dropdownValue === "Other" && otherValue) {
      sheet.getRange(i, dropdownColumn).setValue(otherValue);
    }
  }

  Logger.log("Dropdown instructor column updated with user-entered values!");
}
