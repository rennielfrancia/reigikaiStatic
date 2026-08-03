// Form dropdown maintenance (Google Apps Script)
// Keeps Dojo Name / Instructor Name list items in sync with FormResponses,
// and replaces "Other" selections with free-text entries.
// Not loaded by the static website — kept here for version control / edits.
//
// Form ID: 1ParlaBQqUhifqiWHyFnk-tyojxxG8_wUtr-Wfq-15kY
// Spreadsheet ID: 1_ox1_IopSiFfPvlLjTrJVdjmtYND1NjXdtNBpuWUim4
// Sheet: FormResponses

function updateDojoNameDropdown() {
    // Google Form and Sheet setup
    var formId = "1ParlaBQqUhifqiWHyFnk-tyojxxG8_wUtr-Wfq-15kY";  // Replace with your Google Form ID
    var sheet = SpreadsheetApp.openById("1_ox1_IopSiFfPvlLjTrJVdjmtYND1NjXdtNBpuWUim4").getSheetByName("FormResponses"); // Sheet containing responses
    var dropdownColumn = 6;  // Column F: Dojo Name dropdown selections
    var newEntryColumn = 7;  // Column G: Free-text dojo name entries

    // Fetch dojo names from sheet
    var dropdownValues = sheet.getRange(2, dropdownColumn, sheet.getLastRow() - 1, 1).getValues().flat();
    var newEntries = sheet.getRange(2, newEntryColumn, sheet.getLastRow() - 1, 1).getValues().flat().filter(String).map(toPascalCase);

    // Open the form and find the relevant dropdown item
    var form = FormApp.openById(formId);
    var formItems = form.getItems(FormApp.ItemType.LIST);
    var listItem = formItems.find(item => item.getTitle().includes("Dojo Name"))?.asListItem();

    if (!listItem) {
        Logger.log("List item with title containing 'Dojo Name' not found.");
        return;
    }

    // Get existing choices from the form
    var existingChoices = listItem.getChoices().map(choice => choice.getValue());

    // Combine existing choices with new entries and deduplicate
    var combinedValues = [...new Set(existingChoices.concat(dropdownValues, newEntries))];
    var sortedValues = combinedValues.filter(v => v !== "Other").sort();  // Remove "Other" before sorting
    sortedValues.push("Other");  // Add "Other" back to the end

    // Update form only if values have changed
    if (sortedValues.join() !== existingChoices.join()) {
        listItem.setChoiceValues(sortedValues);
        Logger.log("Dropdown dojo name updated with new entries.");
    } else {
        Logger.log("No new dojo names to add.");
    }
}

// Converts a string to Pascal Case (e.g., "john smith" → "John Smith")
function toPascalCase(str) {
    return str
        .toLowerCase()
        .split(/[\s_-]+/)
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(" ");
}



function updateInstructorNameDropdown() {
    // Google Form and Sheet setup
    var formId = "1ParlaBQqUhifqiWHyFnk-tyojxxG8_wUtr-Wfq-15kY";  // Replace with your Google Form ID
    var sheet = SpreadsheetApp.openById("1_ox1_IopSiFfPvlLjTrJVdjmtYND1NjXdtNBpuWUim4").getSheetByName("FormResponses"); // Sheet containing responses
    var dropdownColumn = 8;  // Column H: Instructor Name dropdown selections
    var newEntryColumn = 9;  // Column I: Free-text instructor name entries

    // Fetch instructor names from sheet
    var dropdownValues = sheet.getRange(2, dropdownColumn, sheet.getLastRow() - 1, 1).getValues().flat();
    var newEntries = sheet.getRange(2, newEntryColumn, sheet.getLastRow() - 1, 1).getValues().flat().filter(String).map(toPascalCase);

    // Open the form and find the relevant dropdown item
    var form = FormApp.openById(formId);
    var formItems = form.getItems(FormApp.ItemType.LIST);
    var listItem = formItems.find(item => item.getTitle().includes("Instructor Name"))?.asListItem();

    if (!listItem) {
        Logger.log("List item with title containing 'Instructor Name' not found.");
        return;
    }

    // Get existing choices from the form
    var existingChoices = listItem.getChoices().map(choice => choice.getValue());

    // Combine existing choices with new entries and deduplicate
    var combinedValues = [...new Set(existingChoices.concat(dropdownValues, newEntries))];
    var sortedValues = combinedValues.filter(v => v !== "Other").sort();  // Remove "Other" before sorting
    sortedValues.push("Other");  // Add "Other" back to the end

    // Update form only if values have changed
    if (sortedValues.join() !== existingChoices.join()) {
        listItem.setChoiceValues(sortedValues);
        Logger.log("Dropdown instructor name updated with new entries.");
    } else {
        Logger.log("No new instructor names to add.");
    }
}

// Converts a string to Pascal Case (e.g., "jane doe" → "Jane Doe")
function toPascalCase(str) {
    return str
        .toLowerCase()
        .split(/[\s_-]+/)
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(" ");
}




function replaceDojoWithUserInput() {    
    var sheet = SpreadsheetApp.openById("1_ox1_IopSiFfPvlLjTrJVdjmtYND1NjXdtNBpuWUim4").getSheetByName("FormResponses"); // Adjust sheet name
    var dropdownColumn = 6;  // Adjust to match dropdown column (e.g., B)
    var otherColumn = 7;  // Adjust to match short-text column (e.g., C)
    var lastRow = sheet.getLastRow();

    // Loop through responses and replace "Other" with user input
    for (var i = 2; i <= lastRow; i++) {
        var dropdownValue = sheet.getRange(i, dropdownColumn).getValue();
        var otherValue = sheet.getRange(i, otherColumn).getValue();

        if (dropdownValue === "Other" && otherValue) {
            sheet.getRange(i, dropdownColumn).setValue(otherValue); // Replace "Other" with user input
        }
    }

    Logger.log("Dropdown dojo column updated with user-entered values!");
}

function replaceInstructorWithUserInput() {    
    var sheet = SpreadsheetApp.openById("1_ox1_IopSiFfPvlLjTrJVdjmtYND1NjXdtNBpuWUim4").getSheetByName("FormResponses"); // Adjust sheet name
    var dropdownColumn = 8;  // Adjust to match dropdown column (e.g., B)
    var otherColumn = 9;  // Adjust to match short-text column (e.g., C)
    var lastRow = sheet.getLastRow();

    // Loop through responses and replace "Other" with user input
    for (var i = 2; i <= lastRow; i++) {
        var dropdownValue = sheet.getRange(i, dropdownColumn).getValue();
        var otherValue = sheet.getRange(i, otherColumn).getValue();

        if (dropdownValue === "Other" && otherValue) {
            sheet.getRange(i, dropdownColumn).setValue(otherValue); // Replace "Other" with user input
        }
    }

    Logger.log("Dropdown instructor column updated with user-entered values!");
}
