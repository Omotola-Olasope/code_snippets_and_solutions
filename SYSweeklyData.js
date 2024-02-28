function sendWeeklyData() {
    const sourceSheetName = "Form Responses 1"; // Name of the sheet with data
    const targetSheetName = "Sheet1"; // Name of the sheet for extracted data
    
    // Get the source and target sheets
    const sourceSheet = SpreadsheetApp.openById("1H2XDXdVGMy5sWopiC4AL5b-jDy3agqBX5wlTnPCjsaQ").getSheetByName(sourceSheetName);
    const targetSheet = SpreadsheetApp.openById("1x3VsPaiAM3kzJ0v2bZQUPeP5CSCDnpO180EqP7gQWsY").getSheetByName(targetSheetName);

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

    // Clear and write optimized data to the target sheet
    targetSheet.clearContents();
    targetSheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData.map(row => row.map(value => value.toString()))); // change all data to string to avoing timezone issues in date
}


function editAndEmailData() {
    const sourceSheetName2 = "Sheet1"; // Name of the sheet with data to be edited
    const targetSheetName2 = "Sheet2"; // Name of the sheet for data to be mailed
    
    //const recipientEmail = "tola.olasope@gmail.com"; // Email address to send to for testing purposes
    const subject = "Weekly Report on Foreigners Movement, MMIA"; // Email subject
    
    // Array of email recipients
    const recipients = ["tola.olasope@gmail.com", "sys.mmia@gmail.com", "xtraconceptsmedia@gmail.com", "tunrayookusaga@gmail.com", "yettynuga730@gmail.com", "steveposby@gmail.com"];

    // Get sheet
    const sourceSheet2 = SpreadsheetApp.openById("1x3VsPaiAM3kzJ0v2bZQUPeP5CSCDnpO180EqP7gQWsY").getSheetByName(sourceSheetName2);
    const targetSheet2 = SpreadsheetApp.openById("1x3VsPaiAM3kzJ0v2bZQUPeP5CSCDnpO180EqP7gQWsY").getSheetByName(targetSheetName2);
    
    // Get all Data to be edited in the Sheet 
    const dataToEdit = sourceSheet2.getDataRange().getValues().map(row => row.map(value => value.toString()));

    // Function to titlecase a string
    function titleCase(str) {
        return str.toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
    }

    // Optimize data before sending via email
    const optimizedData = dataToEdit.map(row => {
        // Get the date string from the 7th column
        const dateString = row[6];

        // Convert the date string to a Date object
        const dateObject = new Date(dateString);

        // Format the date using Utilities.formatDate()
        const formattedDate = Utilities.formatDate(dateObject, "GMT+0100", "EEE MMM dd yyyy");

        // Update the 7th column with the formatted date
        row[6] = formattedDate;

        // Merge values in the 10th and 11th column
        row[9] = row[9] + " " + row[10];

        // Select only specific columns (2, 3, 4, 5, 6, 7, 8, 9, 10, and 13)
        // Titlecase the values in the first 3 columns 
        return [titleCase(row[1]), titleCase(row[2]), titleCase(row[3]), row[4], row[5], row[6], titleCase(row[7]), titleCase(row[8]), row[9], row[12]]; 
    });

    // Clear and write optimized data to the target sheet
    targetSheet2.clearContents();
    targetSheet2.getRange(1, 1, optimizedData.length, optimizedData[0].length).setValues(optimizedData.map(row => row.map(value => value.toString()))); // change all data to string to avoing timezone issues in date

    // Create attachment from the sheet
    const csvString = targetSheet2.getRange(1, 1, targetSheet2.getLastRow(), targetSheet2.getLastColumn())
      .getValues()
      .map(row => row.map(value => (value instanceof Date) ? Utilities.formatDate(value, "GMT+0100", "EEE MMM dd yyyy") : value.toString()))
      .map(row => row.join(','))
      .join('\n');

    const attachment = Utilities.newBlob(csvString, "text/csv", targetSheetName2 + ".csv");

    // Send email with attachment for testing purposes
    //GmailApp.sendEmail(recipientEmail, subject, "", {attachments: [attachment]});

    // Send email with attachment to multiple recipients
    GmailApp.sendEmail(recipients.join(','), subject, "", {attachments: [attachment]});

    // Clear target sheet for next week
    //targetSheet.clearContents();
}

// Set up weekly trigger in Nigerian time for sendWeeklyData
function createWeeklyTriggerForSendData() {
    return ScriptApp.newTrigger("sendWeeklyData")
        .timeBased()
        .everyWeeks(1)  // Set to trigger every 1 week
        .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)  // Set the trigger day to Wednesday
        .inTimezone("Africa/Lagos") // Set the time zone to Nigeria
        .atHour(7)
        .create();
}

// Set up trigger for editAndEmailData to run after sendWeeklyData
function createTriggerForEditAndEmailData() {
    const scriptProperties = PropertiesService.getScriptProperties();
    const triggerId = scriptProperties.getProperty('editAndEmailDataTriggerId');

    // Delete existing trigger if it exists
    if (triggerId) {
        ScriptApp.getProjectTriggers().forEach(trigger => {
            if (trigger.getUniqueId() === triggerId) {
                ScriptApp.deleteTrigger(trigger);
            }
        });
    }

    // Set up new trigger
    const trigger = ScriptApp.newTrigger('editAndEmailData')
        .timeBased()
        .everyWeeks(1)  // Set to trigger every 1 week
        .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)  // Set the trigger day to Wednesday
        .inTimezone("Africa/Lagos") // Set the time zone to Nigeria
        .atHour(8)  // Adjust the hour as needed
        .create();

    // Store the trigger ID in script properties
    scriptProperties.setProperty('editAndEmailDataTriggerId', trigger.getUniqueId());
}



// Function to run both sendWeeklyData and editAndEmailData
//function runWeeklyTasks() {
    // Run sendWeeklyData first
//    sendWeeklyData();

    // Run editAndEmailData immediately after sendWeeklyData
//    editAndEmailData();
//}

// Set up weekly trigger for sendWeeklyData
//createWeeklyTriggerForSendData();

// Set up trigger for editAndEmailData to run after sendWeeklyData
//createTriggerForEditAndEmailData();

// Uncomment the line below if you want to run both functions immediately
// runWeeklyTasks();




