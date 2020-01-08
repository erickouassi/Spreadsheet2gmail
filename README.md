# Spreadsheet2gmail

```
--Code.gs--
function sendEmails() {
   var sheet = SpreadsheetApp.getActiveSheet();
   var range = sheet.getRange(1, 2);  // Fetch the range of cells B1:B1
   var subject = range.getValues();   // Fetch value for subject line from above range
   var range = sheet.getRange(1, 9);  // Fetch the range of cells I1:I1
   var numRows = range.getValues();   // Fetch value for number of emails from above range
   var startRow = 4;                  // First row of data to process
   var dataRange = sheet.getRange(startRow, 1, numRows,9 ) // Fetch the range of cells A4:I_
   var data = dataRange.getValues();  // Fetch values for each row in the Range.
   for (i in data) {
      var row = data[i];
      var emailAddress = row[0];      // First column
      var message = row[8];           // Ninth column
      MailApp.sendEmail(emailAddress, subject, message);
   }
}

```
