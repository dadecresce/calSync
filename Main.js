// Main.gs

function avvio() {
  const today = new Date();
  const { startTime, endTime } = getTimeRange(today);
  
  const calendar = CalendarApp.getCalendarById(getConfig().calendarId);
  const events = calendar.getEvents(startTime, endTime);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const techMapping = buildTechMapping(ss);
  
  const allEvents = updateEvents(ss, events, techMapping);
  updateServices(ss, allEvents, techMapping);
  verifyTechResponses(ss, allEvents);
  verifyServiceResponses(ss);
  verifyIdLinkInResponses(ss);
}

function createDailyTrigger() {
  ScriptApp.newTrigger('avvio')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
}