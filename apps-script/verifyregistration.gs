// Registration Verification API (Google Apps Script)
// Deployed as a web app; called by Registration-Verification.html via POST.
// Requires config.gs in the same Apps Script project.
// Not loaded by the static website — kept here for version control / edits.

// This handles GET requests (when someone visits the URL directly in browser)
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({
      status: "Registration Verification API",
      message: "This endpoint accepts POST requests for registration verification",
      usage: "Send POST request with paymentCode parameter"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

// This handles POST requests (when your website sends verification requests)
function doPost(e) {
  return verifyRegistration(e);
}

// Your actual verification logic
function verifyRegistration(e) {
  var sheet = getFormResponsesSheet();
  const paymentCode = e.parameter.paymentCode;
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][17] === paymentCode) {
      const dojoName = data[i][6] || data[i][5];
      
      // Get the actual hyperlink URL from the cell
      const orderIdCell = sheet.getRange(i + 1, 18); // Row i+1, Column R (18)
      const orderIdUrl = orderIdCell.getRichTextValue().getLinkUrl();
      const orderIdText = data[i][17]; // Display text
      
      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          data: {
            participantName: data[i][1],
            gender: data[i][2],
            age: data[i][3],
            ranking: data[i][4],
            dojoName: dojoName,
            divisionCode: data[i][16],
            payableOrderId: orderIdText,
            payableOrderIdUrl: orderIdUrl || null, // The actual URL
            payableStatus: data[i][19]
          }
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({
      success: false,
      message: "Payment code not found"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
