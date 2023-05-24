function onChange(e) {
  Logger.log(e.changeType);
  if (e.changeType == "INSERT_ROW") {
    Logger.log("It Works 1");
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    let meetingCalendar = CalendarApp.getDefaultCalendar();
    var sheet = spreadsheet.getSheetByName("prospect_requests");
    var targetSheet = spreadsheet.getSheetByName("sheduled_meet");

    var data = sheet.getDataRange().getValues();
    data.splice(0, 1);
    var today = new Date();
    Logger.log(data.length);

    if (data.length != 0) {
      Logger.log("It Works 2");

      for (var i = data.length - 1; i >= 0; i--) {
        var prospect = data[i][0];
        var email = data[i][2];
        var meetingStart = new Date(data[i][3]);
        var meetingEnd = new Date(new Date(data[i][3]).getTime() + 30 * 60000);

        Logger.log(prospect);

        if (meetingStart > today) {
          var event = meetingCalendar.createEvent(
            `Fuel Up - Website Design Agency | ${prospect}`,
            meetingStart,
            meetingEnd,
            {
              guests: email,
              sendInvites: true,
              description: "Meeting with " + prospect,
              location: "Meeting Room",
            }
          );

          // Add yourself as a guest
          event.addGuest("fuelup.yourpresence@gmail.com");

          event.addEmailReminder(30);

          // Move the row to the target sheet
          targetSheet.appendRow(data[i]);

          // Delete the row from the original sheet
          sheet.deleteRow(i + 2);

          // Send email to yourself
          MailApp.sendEmail(
            "fuelup.yourpresence@gmail.com",
            "New Calendar Event",
            "A new event has been created for " + prospect + "."
          );

          // Send email to the prospect
          MailApp.sendEmail(
            email,
            "New Calendar Event",
            "You have been invited to a meeting for " + prospect + "."
          );
        }
      }
    }
  }
}
