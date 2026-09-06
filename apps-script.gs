const SHEET_NAME = "IT Students";

function doGet() {
  return ContentService
    .createTextOutput("SIWES IT Placement Information API is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    // Create sheet and headers if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        "Timestamp",
        "Submission ID",
        "Full Name",
        "Matriculation Number",
        "Course of Study",
        "Level of Study",
        "SIWES Year",
        "Email Address",
        "Phone Number",
        "Nationality",
        "Employer Name",
        "Employer Address",
        "Attachment From",
        "Attachment To",
        "Bank Name",
        "Account Number",
        "Sort Code"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }

    // Parse incoming JSON data
    const data = JSON.parse(e.postData.contents);

    // Generate unique submission ID
    const submissionID = "SUB-" + Utilities.getUuid();
    const timestamp = new Date();

    // Format row data matching input order
    const row = [
      timestamp,
      submissionID,
      data.name || "",
      data.matric || "",
      data.course || "",
      data.level || "",
      data.year || "",
      data.email || "",
      data.phone || "",
      data.nationality || "",
      data.employerName || "",
      data.employerAddress || "",
      data.fromDate || "",
      data.toDate || "",
      data.bank || "",
      "'" + (data.accountNumber || ""), // Prepended apostrophe preserves leading zeros
      "'" + (data.sortCode || "")
    ];

    sheet.appendRow(row);

    // Return response in the format your HTML expects
    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        submissionID: submissionID,
        status: "Pending",
        message: "Placement record submitted successfully",
        timestamp: timestamp.toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader("Access-Control-Allow-Origin", "*")
      .addHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
      .addHeader("Access-Control-Allow-Headers", "Content-Type");

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: "Submission failed: " + error.toString(),
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON)
      .addHeader("Access-Control-Allow-Origin", "*")
      .addHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
      .addHeader("Access-Control-Allow-Headers", "Content-Type");

  } finally {
    lock.releaseLock();
  }
}

// Handle OPTIONS requests for CORS preflight
function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .addHeader("Access-Control-Allow-Origin", "*")
    .addHeader("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
    .addHeader("Access-Control-Allow-Headers", "Content-Type")
    .setMimeType(ContentService.MimeType.TEXT);
}
