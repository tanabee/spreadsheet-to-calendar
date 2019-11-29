var calendarId = PropertiesService
  .getScriptProperties()
  .getProperty('CALENDAR_ID');

function registerEvents() {
  var values = SpreadsheetApp
    .getActiveSheet()
    .getDataRange()
    .getValues();

  values.shift();

  var eventIds = values.map(function (e) {
    var event = CalendarApp
      .getCalendarById(calendarId)
      .createEvent(
        e[0],
        new Date(e[4] + 'T' + e[5]),
        new Date(e[4] + 'T' + e[6]),
        {
          description: e[1] + '<br>Speaker: ' + e[2],
          location: e[3],
        }
      );
    return [event.getId()];
  });

  SpreadsheetApp
    .getActiveSheet()
    .getRange('H2:H' + (eventIds.length+1))
    .setValues(eventIds);
}

function deleteEvents() {
  var values = SpreadsheetApp
    .getActiveSheet()
    .getDataRange()
    .getValues();

  values.shift();

  values.forEach(function (event) {
    CalendarApp
      .getCalendarById(calendarId)
      .getEventById(event[7])
      .deleteEvent();
  });

  SpreadsheetApp
    .getActiveSheet()
    .getRange('H2:H' + (values.length+1))
    .clear();
}
