// UpdateEvents.js

function updateEvents(ss, calendarEvents, techMapping) {
  const config = getConfig();
  const eventsSheet = ss.getSheetByName(config.sheetNames.eventi);
  
  // Preparation: create arrays to store existing and new events
  let existingEvents = [];
  if (eventsSheet.getLastRow() >= 2) {
    existingEvents = eventsSheet.getRange(2, 1, eventsSheet.getLastRow() - 1, 25).getValues();
  }
  const existingIds = existingEvents.map(row => row[0]);
  
  const newEvents = [];
  const allEvents = [];
  
  // Process each calendar event
  calendarEvents.forEach(calEvent => {
    try {
      const eventId = calEvent.getId();
      const title = calEvent.getTitle();
      const description = calEvent.getDescription();
      const location = calEvent.getLocation() || "";
      const startTime = calEvent.getStartTime();
      const endTime = calEvent.getEndTime();
      const formattedDate = Utilities.formatDate(startTime, "GMT+1", "dd-MM-yyyy");
      
      // Convert calendar event title and location - we'll assume title contains event name and location
      // This logic might need adjustment based on your actual calendar event format
      let eventName = title;
      let eventLocation = location;
      
      // Determine the calendar name
      let calendarName = "";
      try {
        calendarName = calEvent.getOriginalCalendarId() ? 
                      CalendarApp.getCalendarById(calEvent.getOriginalCalendarId()).getName() : 
                      "Calendario principale";
      } catch (e) {
        calendarName = "Calendario principale";
      }
      
      // Default structure for an event row (25 columns)
      const eventRow = [
        eventId,                                // ID Evento
        formattedDate,                          // Data in formato "dd-MM-yyyy"
        new Date(startTime),                    // Data come oggetto Date
        eventName,                              // Nome evento
        eventLocation,                          // Luogo
        "", "", "", "", "",                     // Tecnici 1-5 (colonne F-J)
        "", "", "", "", "",                     // ID Tecnici 1-5 (colonne K-O)
        "",                                     // Mezzo/Veicolo (colonna P)
        startTime.toISOString(),                // Data e ora inizio (ISO)
        endTime.toISOString(),                  // Data e ora fine (ISO)
        description || "",                      // Descrizione evento
        "", "", "", "", "",                     // Info extra per tecnici 1-5 (colonne T-X)
        calendarName                            // Nome calendario (colonna Y)
      ];
      
      // Check if this event already exists in our spreadsheet
      const indexEvent = existingIds.indexOf(eventId);
      
      if (indexEvent > -1) {
        // Event exists - update it and preserve tecnici and other data
        const existingEvent = existingEvents[indexEvent];
        
        // Copy existing technicians and IDs
        for (let i = 0; i < 5; i++) {
          eventRow[5 + i] = existingEvent[5 + i]; // Tech names
          eventRow[10 + i] = existingEvent[10 + i]; // Tech IDs
        }
        
        // Copy vehicle information
        eventRow[15] = existingEvent[15]; // Mezzo
        
        // Copy extra info for technicians
        for (let i = 0; i < 5; i++) {
          eventRow[19 + i] = existingEvent[19 + i];
        }
        
        // Update the event in the spreadsheet
        const rowIndex = indexEvent + 2;
        eventsSheet.getRange(rowIndex, 1, 1, 25).setValues([eventRow]);
        eventsSheet.getRange(rowIndex, 1).setBackground('yellow');
      } else {
        // New event - add it to the array of new events
        newEvents.push(eventRow);
      }
      
      // Add this event to the allEvents array (to be returned)
      allEvents.push(eventRow);
    } catch (error) {
      Logger.log(`Errore nell'elaborazione dell'evento: ${error.toString()}`);
    }
  });
  
  // If there are new events, append them to the spreadsheet
  if (newEvents.length > 0) {
    eventsSheet.getRange(eventsSheet.getLastRow() + 1, 1, newEvents.length, 25).setValues(newEvents);
  }
  
  return allEvents;
}