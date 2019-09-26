// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'YES';

var emailAddress = 'george.tucker@diversdirect.com'; // Change this to update recipient of email!

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; // First row of data to process
  var numRows = 7; // Number of rows to process
  // Fetch the range of cells A2:B3 (changed numRows from 3 to 7 hoping this will reach to column H
  var dataRange = sheet.getRange(startRow, 1, numRows, 6);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];

    var message = "from " + row[1] + "about " + row[3]; 
    var urgency = row[4]; // column E
    var emailSent = row[7]; // column G
    if (emailSent != 'YES') { // Prevents sending duplicates
      var subject = 'New product update request!' + urgency;
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(startRow + i, 7).setValue(EMAIL_SENT); 
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}



//function sendEmails() { // Get the sheet where the data is, in sheet 'system' 
//  var sheet = SpreadsheetApp.getActiveSheet();
//  var startRow = 2; // First row of data to process since there is a header row 
//  var numRows = sheet.getRange(1,7).getValue(); // The way scripts thinks about cell references is Row Column. So E1 is 1, 5. Row one, column 5. 
//  
//  // So remember the Row, Column order? The longer way or writing this in this function is: Row, Column, Number of Rows, Number of Columns
//  var dataRange = sheet.getRange(startRow, 1, numRows, 2) // sheet.getRange(Row, Column, Number of Rows, Number of Columns) becomes sheet.getRange(startRow, 1, numRows, 2)
//  // Why are there names in here? Well we defined startRow and numRows above! 1 is the first column and 2 means the first two columns where email and message are situated.
//  
//  var data = dataRange.getValues(); // This processes the emails you want to send 
//  for (i in data) 
//  { var row = data[i];  
//    var message = row[2]; // Column C
//    var urgency = row[4]; // column E
//    var emailSent = row[6]; // column G
//    if (emailSent != EMAIL_SENT) { // Prevents sending duplicates
//      var subject = 'New product update request!' + urgency;
//      MailApp.sendEmail(emailAddress, subject, message);
//      sheet.getRange(startRow + i, 7).setValue(EMAIL_SENT); // changed value 3 to 7 hoping this will update the right cell
//      // Make sure the cell is updated right away in case the script is interrupted
//      SpreadsheetApp.flush();
//    }
//  }
//}
