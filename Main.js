// Main.gs

function avvio() {
  const today = new Date();
  const { startTime, endTime } = getTimeRange(today);
  
  const config = getConfig();
  let allCalendarEvents = [];
  
  // Raccogliamo gli eventi da tutti i calendari configurati
  config.calendarIds.forEach(calendarId => {
    try {
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (calendar) {
        const events = calendar.getEvents(startTime, endTime);
        allCalendarEvents = allCalendarEvents.concat(events);
        Logger.log(`Trovati ${events.length} eventi nel calendario ${calendarId}`);
      } else {
        Logger.log(`Impossibile accedere al calendario: ${calendarId}`);
      }
    } catch (error) {
      Logger.log(`Errore nell'accesso al calendario ${calendarId}: ${error.toString()}`);
    }
  });
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const techMapping = buildTechMapping(ss);
  
  const allEvents = updateEvents(ss, allCalendarEvents, techMapping);
  updateServices(ss, allEvents, techMapping);
  verifyTechResponses(ss, allEvents);
  verifyServiceResponses(ss);
  verifyIdLinkInResponses(ss);
  
  // Aggiorna i calendari dei tecnici
  updateTechCalendars(ss, allEvents, techMapping);
}

function createDailyTrigger() {
  ScriptApp.newTrigger('avvio')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
}