// 設定
function getConfig() {
  return {
    spreadSheetId: '',
    spreadSheetTabName: ''
  }
}

// カレンダーにイベントを登録
function registerAll() {
  var config      = getConfig();
  var spreadSheet = SpreadsheetApp.openById(config.spreadSheetId);
  var sheet       = spreadSheet.getSheetByName(config.spreadSheetTabName);
  var eventIds = sheet
    .getDataRange()
    .getValues()
    .filter(function (e, i) {
      return i !== 0 && e[6] === '';
    }).map(function (e) {
      var calendarEvent = CalendarApp.getDefaultCalendar().createEvent(
        e[0], e[1], e[2],
        {
          description: e[3],
          location: e[4],
          guests: e[5]
        });
      return [calendarEvent.getId()];
    });

  if (eventIds.length === 0) return;

  sheet.getRange('G2:G' + (eventIds.length+1) ).setValues(eventIds);
}

// 登録されたイベントをキャンセルして、シートから削除
function cancelAll() {
  var range       = 'G2:G1000';
  var config      = getConfig();
  var spreadSheet = SpreadsheetApp.openById(config.spreadSheetId);
  var sheet       = spreadSheet.getSheetByName(config.spreadSheetTabName);
  sheet
    .getRange(range)
    .getValues()
    .filter(function (eventId) {
      return eventId[0] !== '';
    }).forEach(function (id) {
      CalendarApp.getDefaultCalendar().getEventById(id).deleteEvent();
    });
  sheet.getRange(range).clear();
}
