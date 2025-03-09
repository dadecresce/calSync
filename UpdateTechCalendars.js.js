// UpdateTechCalendars.js

function updateTechCalendars(ss, allEvents, techMapping) {
  const config = getConfig();
  
  // Filtra solo gli eventi con tecnici assegnati
  allEvents.forEach(event => {
    const eventId = event[0];
    const eventName = event[3]; // Nome evento
    const luogo = event[4];     // Luogo/indirizzo
    const startTimeISO = event[16];
    const endTimeISO = event[17];
    const description = event[18];
    const extraInfos = event.slice(19, 24);
    
    // Itera sui tecnici assegnati all'evento
    for (let i = 0; i < 5; i++) {
      const techName = event[5 + i];
      if (techName && techName.trim() !== "" && techMapping.hasOwnProperty(techName.toLowerCase())) {
        const tech = techMapping[techName.toLowerCase()];
        
        // Verifica se il tecnico ha un calendario configurato
        if (tech.calendarId && tech.calendarId.trim() !== "") {
          try {
            const techCalendar = CalendarApp.getCalendarById(tech.calendarId);
            if (techCalendar) {
              // Crea titolo evento per il calendario del tecnico
              const techEventTitle = `${eventName} - ${luogo}`;
              
              // Recupera il link al form precompilato
              const idLink = `${eventId}${techName}`;
              let formLink = "";
              
              // Cerca il link al form nella scheda servizi
              const servicesSheet = ss.getSheetByName(config.sheetNames.servizi);
              if (servicesSheet) {
                const servicesData = servicesSheet.getDataRange().getValues();
                for (let j = 1; j < servicesData.length; j++) {
                  if (servicesData[j][4] && servicesData[j][4].toString().trim().toLowerCase() === idLink.toLowerCase()) {
                    formLink = servicesData[j][5]; // Colonna F contiene il link al form
                    break;
                  }
                }
              }
              
              // Crea descrizione per il calendario del tecnico
              let extraInfo = extraInfos[i] || "";
              extraInfo = (typeof extraInfo === "string") ? extraInfo.trim() : "";
              
              let fullDescription = description || "";
              if (extraInfo) {
                fullDescription = `${fullDescription}\n\nNote: ${extraInfo}`;
              }
              
              // Aggiungi il link al form precompilato
              if (formLink) {
                fullDescription = `${fullDescription}\n\nCompila il form di servizio: ${formLink}`;
              }
              
              // Crea o aggiorna l'evento nel calendario del tecnico
              const startTime = new Date(startTimeISO);
              const endTime = new Date(endTimeISO);
              
              // Cerca se esiste già un evento con lo stesso ID nel calendario del tecnico
              // Utilizziamo una proprietà estesa per memorizzare l'ID dell'evento originale
              const existingEvents = techCalendar.getEvents(
                new Date(startTime.getTime() - 3600000), // 1 ora prima
                new Date(endTime.getTime() + 3600000)    // 1 ora dopo
              );
              
              let techEvent = null;
              let eventFound = false;
              
              // Prima verifica se c'è un evento con l'ID originale nelle proprietà estese
              for (let j = 0; j < existingEvents.length; j++) {
                const currentEvent = existingEvents[j];
                try {
                  const properties = currentEvent.getAllExtendedProperties();
                  if (properties && properties.shared && properties.shared.originalEventId === eventId) {
                    techEvent = currentEvent;
                    eventFound = true;
                    break;
                  }
                } catch (propError) {
                  // Ignora errori nelle proprietà estese
                }
              }
              
              // Se non è stato trovato un evento con l'ID originale, cerca eventi con titolo simile
              if (!eventFound) {
                for (let j = 0; j < existingEvents.length; j++) {
                  if (existingEvents[j].getTitle().includes(eventName) || 
                      existingEvents[j].getTitle().includes(luogo)) {
                    techEvent = existingEvents[j];
                    eventFound = true;
                    break;
                  }
                }
              }
              
              if (eventFound && techEvent) {
                // Aggiorna l'evento esistente
                techEvent.setTitle(techEventTitle);
                techEvent.setDescription(fullDescription);
                if (luogo && luogo.trim() !== "") {
                  techEvent.setLocation(luogo);
                }
                techEvent.setTime(startTime, endTime);
                
                // Assicuriamoci che l'ID originale sia salvato nelle proprietà estese
                try {
                  techEvent.setExtendedProperty('originalEventId', eventId);
                } catch (propError) {
                  Logger.log(`Impossibile impostare la proprietà estesa: ${propError.toString()}`);
                }
                
                Logger.log(`Evento aggiornato nel calendario di ${tech.name}: ${techEventTitle}`);
              } else {
                // Crea un nuovo evento
                techEvent = techCalendar.createEvent(
                  techEventTitle,
                  startTime,
                  endTime,
                  {
                    description: fullDescription,
                    location: luogo
                  }
                );
                
                // Salva l'ID originale nelle proprietà estese
                try {
                  techEvent.setExtendedProperty('originalEventId', eventId);
                } catch (propError) {
                  Logger.log(`Impossibile impostare la proprietà estesa: ${propError.toString()}`);
                }
                
                Logger.log(`Nuovo evento creato nel calendario di ${tech.name}: ${techEventTitle}`);
              }
            }
          } catch (error) {
            Logger.log(`Errore nell'aggiornamento del calendario per ${tech.name}: ${error.toString()}`);
          }
        }
      }
    }
  });
}