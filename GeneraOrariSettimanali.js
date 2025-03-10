// GeneraOrariSettimanali.gs
function generaOrariSettimanali() {
  // Usa la configurazione esistente
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const today = new Date();
  const { startTime, endTime } = getTimeRange(today);
  
  // Ottieni la mappatura dei tecnici
  const techMapping = buildTechMapping(ss);
  
  // Ottieni tutti gli eventi dai calendari configurati
  let allCalendarEvents = [];
  config.calendarIds.forEach(calendarId => {
    try {
      const calendar = CalendarApp.getCalendarById(calendarId);
      if (calendar) {
        const events = calendar.getEvents(startTime, endTime);
        allCalendarEvents = allCalendarEvents.concat(events);
        Logger.log(`Trovati ${events.length} eventi nel calendario ${calendarId}`);
      }
    } catch (error) {
      Logger.log(`Errore nell'accesso al calendario ${calendarId}: ${error.toString()}`);
    }
  });
  
  // Processa gli eventi per generare il messaggio settimanale
  const eventsSheet = ss.getSheetByName(config.sheetNames.eventi);
  const servicesSheet = ss.getSheetByName(config.sheetNames.servizi);
  
  // Ottieni eventi esistenti dalla spreadsheet
  const allEvents = updateEvents(ss, allCalendarEvents, techMapping);
  
  // Per ogni tecnico, genera un messaggio WhatsApp con tutti i suoi eventi della settimana
  for (const techNameLower in techMapping) {
    const tech = techMapping[techNameLower];
    if (!tech.phone) continue; // Salta i tecnici senza numero di telefono
    
    // Filtra gli eventi per questo tecnico
    const techEvents = [];
    allEvents.forEach(event => {
      for (let i = 0; i < 5; i++) {
        const eventTechName = event[5 + i];
        if (eventTechName && eventTechName.toLowerCase() === techNameLower) {
          techEvents.push(event);
          break;
        }
      }
    });
    
    if (techEvents.length === 0) continue; // Salta se non ci sono eventi per questo tecnico
    
    // Costruisci il messaggio con tutti gli eventi della settimana
    let message = `Ciao ${tech.name}, ecco i tuoi eventi per questa settimana:\n\n`;
    
    techEvents.forEach((event, index) => {
      const eventId = event[0];
      const eventDate = Utilities.formatDate(new Date(event[2]), "GMT+1", "dd/MM/yyyy");
      const eventName = event[3];
      const luogo = event[4];
      const startTimeISO = event[16];
      const endTimeISO = event[17];
      const startTime = Utilities.formatDate(new Date(startTimeISO), "GMT+1", "HH:mm");
      const endTime = Utilities.formatDate(new Date(endTimeISO), "GMT+1", "HH:mm");
      const vehicle = event[15] || "Non specificato";
      
      // Trova il link al form per questo evento/tecnico
      const idLink = `${eventId}${tech.name}`;
      let formLink = "";
      
      // Cerca nella scheda servizi
      const servicesData = servicesSheet.getDataRange().getValues();
      for (let j = 1; j < servicesData.length; j++) {
        if (servicesData[j][4] && servicesData[j][4].toString().trim() === idLink) {
          formLink = servicesData[j][5]; // Colonna F contiene il link al form
          break;
        }
      }
      
      // Aggiungi informazioni sull'evento
      message += `EVENTO ${index + 1}:\n`;
      message += `Data: ${eventDate}\n`;
      message += `Titolo: ${eventName}\n`;
      message += `Luogo: ${luogo}\n`;
      message += `Orario: ${startTime} - ${endTime}\n`;
      message += `Mezzo: ${vehicle}\n`;
      
      // Aggiungi il link al form se disponibile
      if (formLink) {
        message += `Link per conferma: ${formLink}\n`;
      }
      
      message += "\n";
    });
    
    message += "Per favore, conferma tutti i tuoi eventi compilando i relativi moduli. Grazie!";
    
    // Genera il link WhatsApp
    const whatsappUrl = `https://wa.me/${tech.phone}?text=${encodeURIComponent(message)}`;
    
    // Crea uno short link per WhatsApp
    const shortWhatsappUrl = createShortUrl(whatsappUrl);
    
    // Aggiungi una riga con le info del tecnico e il link WhatsApp
    const weekRange = `${Utilities.formatDate(startTime, "GMT+1", "dd/MM/yyyy")} - ${Utilities.formatDate(endTime, "GMT+1", "dd/MM/yyyy")}`;
    Logger.log(`Generato link WhatsApp per ${tech.name}: ${shortWhatsappUrl}`);
    
    // Puoi anche salvare questi link in una nuova scheda se necessario
    // Ad esempio:
    /*
    let schedulesSheet = ss.getSheetByName("orari_settimanali");
    if (!schedulesSheet) {
      schedulesSheet = ss.insertSheet("orari_settimanali");
      schedulesSheet.appendRow(["Tecnico", "Settimana", "Numero Eventi", "Link WhatsApp"]);
    }
    schedulesSheet.appendRow([tech.name, weekRange, techEvents.length, shortWhatsappUrl]);
    */
  }
  
  Logger.log("Link WhatsApp settimanali generati con successo");
  return "Operazione completata con successo";
}

// Se non hai già la funzione createShortUrl, assicurati di includerla
// Questa è identica a quella nel tuo file UpdateServices.js
function createShortUrl(longUrl) {
  if (typeof longUrl !== 'string' || !longUrl) return '';
  
  try {
    const encodedUrl = encodeURIComponent(longUrl);
    const response = UrlFetchApp.fetch(`https://tinyurl.com/api-create.php?url=${encodedUrl}`);
    
    if (response.getResponseCode() === 200) {
      return response.getContentText();
    } else {
      Logger.log('Errore nella creazione dello short URL: ' + response.getContentText());
      return longUrl;
    }
  } catch (error) {
    Logger.log('Errore nel servizio di shortening: ' + error.toString());
    return longUrl;
  }
}