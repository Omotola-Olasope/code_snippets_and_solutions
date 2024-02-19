function sendWeeklyDataEmail() {
    
    const sourceSheetName = "Form Responses 1"; // Name of the sheet with data
    const targetSheetName = "Sheet1"; // Name of the sheet for extracted data
    //const recipientEmail = "tola.olasope@gmail.com"; // Email address to send to for testing purposes
    const subject = "Weekly Report on Foreigners Movement, MMIA"; // Email subject
    // Array of email recipients
    const recipients = ["tola.olasope@gmail.com", "sys.mmia@gmail.com", "xtraconceptsmedia@gmail.com", "tunrayookusaga@gmail.com", "yettynuga730@gmail.com", "steveposby@gmail.com"];


    // Get the source and target sheets
    const sourceSheet = SpreadsheetApp.openById("1H2XDXdVGMy5sWopiC4AL5b-jDy3agqBX5wlTnPCjsaQ").getSheetByName(sourceSheetName);
    const targetSheet = SpreadsheetApp.openById("1jV7bKwrNK4oUWze-TsXmayM5rSbW3LXkNQGraMmlmwA").getSheetByName(targetSheetName);

    // Set the time zone to "Africa/Lagos"
    const timeZone = "Africa/Lagos";

    // Get today's date and calculate last week's dates in Nigerian time
    const today = new Date();
    const lastWeekStart = Utilities.formatDate(new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000), timeZone, "yyyy-MM-dd HH:mm:ss");
    const lastWeekEnd = Utilities.formatDate(new Date(today.getTime() - 60 * 1000), timeZone, "yyyy-MM-dd HH:mm:ss");


    // Filter data based on date range
    const dataRange = sourceSheet.getDataRange();
    const filteredData = dataRange.getValues().filter(row =>
        Utilities.formatDate(new Date(row[0]), timeZone, "yyyy-MM-dd HH:mm:ss") >= lastWeekStart &&
        Utilities.formatDate(new Date(row[0]), timeZone, "yyyy-MM-dd HH:mm:ss") <= lastWeekEnd
    );

    // Clear and write data to the target sheet
    targetSheet.clearContents();
    targetSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData.map(row => row.map(value => value.toString()))); // change all data to string to avoing timezone issues in date


    // Create attachment from target sheet
    const csvString = targetSheet.getRange(1, 1, targetSheet.getLastRow(), targetSheet.getLastColumn()).getValues().map(row => row.join(',')).join('\n');
    const attachment = Utilities.newBlob(csvString, "text/csv", targetSheetName + ".csv");

    // Send email with attachment for testing purposes
    //GmailApp.sendEmail(recipientEmail, subject, "", {attachments: [attachment]});

    // Send email with attachment to multiple recipients
    GmailApp.sendEmail(recipients.join(','), subject, "", {attachments: [attachment]});

    // Clear target sheet for next week
    targetSheet.clearContents();
}

// Set up weekly trigger in Nigerian time
function createTrigger() {
    ScriptApp.newTrigger("sendWeeklyDataEmail")
        .timeBased()
        .everyWeeks(1)  // Set to trigger every 1 week
        .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)  // Set the trigger day to Wednesday
        .inTimezone("Africa/Lagos") // Set the time zone to Nigeria
        .atHour(7)
        .create();
}