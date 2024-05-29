# Email-Automation
This repo contains Java code/logic that can be used for customized email automation (generally used for marketing purpose).

## Approach -
Uses links of uploaded Google Drive folder and files that are used to access the code as well as the required media.

1. The following code takes in "first Name" and "Email" as the dynamic fields. Main content is stored in "folderId" and "fileName".
2. A "On edit" trigger is put in-place upon the function "sendEmailToNewEntry" to automate the process of sending emails in real-time.

## Code - 
"""
// Function to send automated email to new entries
function sendEmailToNewEntry() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var emailColumn = 5; // Assuming email is in the fifth column
  var lastSentRow = getLastSentRow();
  var emailSent = false;

  // Iterate through new entries starting from the last sent row + 1
  for (var i = lastSentRow + 1; i <= lastRow; i++) {
    var emailAddress = sheet.getRange(i, emailColumn).getValue(); // Get email from the current row

    // Log the email address retrieved
    Logger.log("Email from spreadsheet: " + emailAddress);

    // Send email if not sent already and email is valid
    if (emailAddress && !isEmailSent(emailAddress) && isValidEmail(emailAddress)) {
      var firstName = sheet.getRange(i, 1).getValue(); // Assuming first name is in the first column (column A)
      var subject = "Welcome to GNYAN.AI :)";
      var message = "Dear " + firstName + ",\n\n\n" + getEmailMessage();
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: message
      });

      // Record sent email
      recordSentEmail(emailAddress);
      Logger.log("Email sent to " + firstName + " at " + emailAddress);
      emailSent = true;
    } else {
      Logger.log("Skipped sending email: Invalid email address or already sent.");
    }
  }

  // Log if no email was sent
  if (!emailSent) {
    Logger.log("No new emails to send.");
  }
}

// Function to validate an email address
function isValidEmail(email) {
  // Regular expression for email validation
  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailPattern.test(email);
}

// Function to get the last row number that has been sent
function getLastSentRow() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sentSheet = ss.getSheetByName("SentEmails");
  if (!sentSheet) return 0; // If the SentEmails sheet doesn't exist, assume no emails have been sent yet

  var lastRow = sentSheet.getLastRow();
  return lastRow;
}

// Function to check if email has already been sent
function isEmailSent(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sentSheet = ss.getSheetByName("SentEmails");
  if (!sentSheet) return false; // If the SentEmails sheet doesn't exist, assume the email hasn't been sent

  var emailColumn = sentSheet.getRange("A:A").getValues(); // Assuming email is in column A
  for (var i = 0; i < emailColumn.length; i++) {
    if (emailColumn[i][0] == email) {
      return true;
    }
  }
  return false;
}

// Function to record sent email
function recordSentEmail(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sentSheet = ss.getSheetByName("SentEmails");
  if (!sentSheet) { // If the SentEmails sheet doesn't exist, create it
    sentSheet = ss.insertSheet("SentEmails");
    sentSheet.appendRow(["Email Address", "Timestamp"]); // Header row
  }

  var timestamp = new Date();
  sentSheet.appendRow([email, timestamp]); // Record email address and timestamp
}

// Function to get email message from HTML file in Google Drive
function getEmailMessage() {
  var folderId = '1IPL_qXhhVMwEgUSBJ9KPYJjbTQXtDnsl';
  var fileName = 'new-email.html';

  var files = DriveApp.getFolderById(folderId).getFilesByName(fileName);
  if (files.hasNext()) {
    var file = files.next();
    var content = file.getBlob().getDataAsString();
    return content;
  } else {
    throw new Error("HTML file not found");
  }
}

// Create an onEdit trigger to send email to new entry
function createOnEditTrigger() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('sendEmailToNewEntry')
    .forSpreadsheet(sheet)
    .onEdit()
    .create();
}
"""
