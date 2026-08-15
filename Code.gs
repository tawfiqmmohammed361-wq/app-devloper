// Google Apps Script for handling form submissions
// Deploy as: New Deployment > Web app > Execute as Me > Anyone

const SHEET_NAME = "Responses";
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");

function doGet(e) {
  const name = e.parameter.name || "";
  const email = e.parameter.email || "";
  const phone = e.parameter.phone || "";

  return appendToSheet(name, email, phone);
}

function doPost(e) {
  let name = "";
  let email = "";
  let phone = "";

  try {
    if (e.postData) {
      const contentType = e.postData.type;
      
      if (contentType === "application/x-www-form-urlencoded") {
        const params = e.parameter;
        name = params.name || "";
        email = params.email || "";
        phone = params.phone || "";
      } else if (contentType === "application/json") {
        const jsonData = JSON.parse(e.postData.contents);
        name = jsonData.name || "";
        email = jsonData.email || "";
        phone = jsonData.phone || "";
      }
    } else if (e.parameter) {
      name = e.parameter.name || "";
      email = e.parameter.email || "";
      phone = e.parameter.phone || "";
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: "Error parsing request: " + err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return appendToSheet(name, email, phone);
}

function appendToSheet(name, email, phone) {
  try {
    // Get the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if sheet exists, if not create it
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      
      // Add headers with bold formatting
      const headerRange = sheet.getRange(1, 1, 1, 4);
      headerRange.setValues([["Timestamp", "Name", "Email", "Phone"]]);
      headerRange.setFontWeight("bold");
      
      // Auto-fit columns
      sheet.autoResizeColumns(1, 4);
    }

    // Get the next row
    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;

    // Create timestamp
    const timestamp = new Date().toLocaleString("en-IN", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: true
    });

    // Append data
    sheet.getRange(nextRow, 1, 1, 4).setValues([[timestamp, name, email, phone]]);

    // Return success response
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: "Response recorded successfully",
      timestamp: timestamp
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
