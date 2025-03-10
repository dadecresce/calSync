/**
 * Genera un messaggio di riepilogo settimanale per ciascun tecnico e lo salva nel foglio tecnici
 */
function GeneraMessaggioSettimana() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig();
  const eventsSheet = ss.getSheetByName(config.sheetNames.eventi);
  const techSheet = ss.getSheetByName(config.sheetNames.tecnici);
  
  // Verifica che esista la colonna "messaggio Settimana" nel foglio tecnici
  let headers = techSheet.getRange(1, 1, 1, techSheet.getLastColumn()).getValues()[0];
  let messaggioColIndex = headers.indexOf("messaggio Settimana") + 1;
  
  // Se la colonna non esiste, la aggiungiamo
  if (messaggioColIndex === 0) {
    messaggioColIndex = techSheet.getLastColumn() + 1;
    techSheet.getRange(1, messaggioColIndex).setValue("messaggio Settimana");
  }
  
  // Recupera i dati dei tecnici
  const techData = techSheet.getDataRange().getValues();
  const techMap = {};
  
  // Crea una mappa dei tecnici (nome -> riga)
  for (let i = 1; i < techData.length; i++) {
    const techName = techData[i][1].toString().trim(); // Nome del tecnico in colonna B
    if (techName) {
      techMap[techName.toLowerCase()] = i + 1; // Salva l'indice della riga (+1 perché gli indici partono da 0)
    }
  }
  
  // Ottieni la data odierna e il range della settimana
  const today = new Date();
  const { startTime, endTime } = getTimeRange(today);
  
  // Recupera tutti gli eventi della settimana
  let eventsData = [];
  if (eventsSheet.getLastRow() >= 2) {
    eventsData = eventsSheet.getRange(2, 1, eventsSheet.getLastRow() - 1, 25).getValues();
  }
  
  // Inizializza un oggetto per tracciare gli impegni di ciascun tecnico
  const techEvents = {};
  
  // Esamina tutti gli eventi
  eventsData.forEach(event => {
    // Controlla se l'evento è nella settimana corrente
    const eventDate = new Date(event[2]); // Colonna C: data evento
    
    // Se l'evento è nel range della settimana
    if (eventDate >= startTime && eventDate <= endTime) {
      const eventName = event[3]; // Colonna D: nome evento
      const location = event[4]; // Colonna E: luogo
      const startTimeISO = event[16]; // Colonna Q: orario inizio (ISO)
      const endTimeISO = event[17]; // Colonna R: orario fine (ISO)
      
      // Formatta orari di inizio e fine
      const startDateTime = new Date(startTimeISO);
      const endDateTime = new Date(endTimeISO);
      const startTimeFormatted = Utilities.formatDate(startDateTime, "GMT+1", "HH:mm");
      const endTimeFormatted = Utilities.formatDate(endDateTime, "GMT+1", "HH:mm");
      
      // Esamina tutti i tecnici assegnati all'evento
      for (let i = 0; i < 5; i++) {
        const techName = event[5 + i]; // Colonne F-J: nomi tecnici
        if (techName && techName.trim() !== "") {
          const techKey = techName.toLowerCase().trim();
          
          // Inizializza l'array degli eventi per questo tecnico se non esiste
          if (!techEvents[techKey]) {
            techEvents[techKey] = [];
          }
          
          // Aggiungi l'evento all'array degli eventi del tecnico
          techEvents[techKey].push({
            date: startDateTime,
            startTime: startTimeFormatted,
            endTime: endTimeFormatted,
            eventName: eventName,
            location: location
          });
        }
      }
    }
  });
  
  // Array con i nomi dei giorni della settimana abbreviati
  const giorniSettimana = ["Dom", "Lun", "Mar", "Mer", "Gio", "Ven", "Sab"];
  
  // Per ogni tecnico, genera il messaggio settimanale
  Object.keys(techEvents).forEach(techKey => {
    const events = techEvents[techKey];
    
    // Ordina gli eventi per data
    events.sort((a, b) => a.date - b.date);
    
    // Raggruppa gli eventi per giorno della settimana
    const eventiPerGiorno = {};
    events.forEach(event => {
      const giornoSettimana = event.date.getDay(); // 0-6 (Dom-Sab)
      const giorno = giorniSettimana[giornoSettimana];
      
      if (!eventiPerGiorno[giorno]) {
        eventiPerGiorno[giorno] = [];
      }
      
      eventiPerGiorno[giorno].push(event);
    });
    
    // Costruisci il messaggio
    let messaggio = `Impegni della settimana dal ${Utilities.formatDate(startTime, "GMT+1", "dd/MM")} al ${Utilities.formatDate(endTime, "GMT+1", "dd/MM")}:\n\n`;
    
    // Aggiungi gli eventi per ogni giorno
    giorniSettimana.forEach(giorno => {
      const eventiGiorno = eventiPerGiorno[giorno];
      if (eventiGiorno && eventiGiorno.length > 0) {
        messaggio += `${giorno}: `;
        
        // Aggiungi tutti gli eventi di questo giorno
        const eventiFormattati = eventiGiorno.map(event => 
          `${event.startTime} - ${event.endTime} ${event.eventName} @ ${event.location}`
        );
        
        messaggio += eventiFormattati.join(", ") + "\n";
      }
    });
    
    // Aggiungi nota finale
    messaggio += "\nPer ulteriori dettagli, controlla i link nei singoli messaggi di servizio.";
    
    // Salva il messaggio nella riga del tecnico
    const techRow = techMap[techKey];
    if (techRow) {
      techSheet.getRange(techRow, messaggioColIndex).setValue(messaggio);
      Logger.log(`Messaggio settimanale salvato per ${techKey}`);
    }
  });
  
  // Notifica completamento
  SpreadsheetApp.getUi().alert("Messaggi settimanali generati e salvati con successo!");
}

/**
 * Funzione per generare e inviare via WhatsApp il messaggio settimanale a un tecnico
 */
function inviaMessaggioSettimanaleTecnico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const techSheet = ss.getSheetByName(getConfig().sheetNames.tecnici);
  
  // Ottieni la riga selezionata
  const selectedRow = techSheet.getActiveCell().getRow();
  if (selectedRow < 2) {
    SpreadsheetApp.getUi().alert("Seleziona una riga che contiene i dati di un tecnico.");
    return;
  }
  
  // Trova l'indice della colonna "messaggio Settimana"
  const headers = techSheet.getRange(1, 1, 1, techSheet.getLastColumn()).getValues()[0];
  const messaggioColIndex = headers.indexOf("messaggio Settimana") + 1;
  
  if (messaggioColIndex === 0) {
    SpreadsheetApp.getUi().alert("Colonna 'messaggio Settimana' non trovata. Genera prima i messaggi.");
    return;
  }
  
  // Leggi i dati del tecnico
  const techName = techSheet.getRange(selectedRow, 2).getValue(); // Nome in colonna B
  const techPhone = techSheet.getRange(selectedRow, 3).getValue(); // Telefono in colonna C
  const messaggioSettimana = techSheet.getRange(selectedRow, messaggioColIndex).getValue();
  
  if (!messaggioSettimana) {
    SpreadsheetApp.getUi().alert("Nessun messaggio settimanale trovato per questo tecnico.");
    return;
  }
  
  if (!techPhone) {
    SpreadsheetApp.getUi().alert("Numero di telefono mancante per questo tecnico.");
    return;
  }
  
  // Crea il messaggio WhatsApp
  const message = encodeURIComponent(`Ciao ${techName},\n\n${messaggioSettimana}`);
  const whatsappUrl = `https://wa.me/${techPhone.toString().replace(/[^0-9]/g, "")}?text=${message}`;
  
  // Crea short URL se possibile
  let finalUrl;
  try {
    finalUrl = createShortUrl(whatsappUrl);
  } catch (error) {
    Logger.log('Errore nella creazione dello short URL: ' + error.toString());
    finalUrl = whatsappUrl;
  }
  
  // Apri WhatsApp Web
  const htmlOutput = HtmlService
    .createHtmlOutput(`<script>window.open("${finalUrl}", "_blank"); google.script.host.close();</script>`)
    .setWidth(1)
    .setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Apertura WhatsApp Web...");
}

/**
 * Aggiunge voci di menu per la gestione dei messaggi settimanali
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Comunicazioni')
    .addItem('Genera Messaggi Settimanali', 'GeneraMessaggioSettimana')
    .addItem('Invia Messaggio Settimanale al Tecnico Selezionato', 'inviaMessaggioSettimanaleTecnico')
    .addSeparator()
    .addItem('Invia Messaggio WhatsApp', 'sendWhatsAppMessage')
    .addToUi();
}