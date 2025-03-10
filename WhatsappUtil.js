/**
 * Generates a WhatsApp Web link with a precomposed message
 * 
 * @param {string} phoneNumber - Phone number in international format without leading + (e.g. "393471234567")
 * @param {string} eventName - Name of the event
 * @param {string} location - Location of the event
 * @param {string} formDate - Date in format "yyyy-MM-dd"
 * @param {string} startTime - Start time in format "HH:mm"
 * @param {string} endTime - End time in format "HH:mm"
 * @param {string} vehicle - Vehicle information
 * @param {string} description - Event description
 * @param {string} techName - Technician name
 * @param {string} formLink - Link to the form (optional)
 * @return {string} The WhatsApp Web link with precomposed message
 */
function generateWhatsAppLink(phoneNumber, eventName, location, formDate, startTime, endTime, vehicle, description, techName, formLink) {
  if (!phoneNumber) {
    return "";
  }
  
  // Clean the phone number (remove spaces, dashes, etc.)
  phoneNumber = phoneNumber.toString().replace(/[^0-9]/g, "");
  
  // Compose the message
  const message = encodeURIComponent(
    `Ciao ${techName}, ecco il tuo evento:\n` +
    `Luogo: ${eventName}\n` +
    `Titolo: ${location}\n` +
    `Data: ${formDate}\n` +
    `Inizio: ${startTime}\n` +
    `Fine: ${endTime}\n` +
    `Mezzo: ${vehicle}\n` +
    `Descrizione: ${description}\n` +
    (formLink ? `Link: ${formLink}` : "")
  );
  
  // Create the WhatsApp URL
  const whatsappUrl = `https://wa.me/${phoneNumber}?text=${message}`;
  
  // Create short URL if possible
  try {
    return createShortUrl(whatsappUrl);
  } catch (error) {
    Logger.log('Errore nella creazione dello short URL: ' + error.toString());
    return whatsappUrl;
  }
}

/**
 * Example usage: Generate and open a WhatsApp Web link
 * This function can be called from a custom menu or button
 */
function sendWhatsAppMessage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const selection = sheet.getActiveCell();
  const row = selection.getRow();
  
  // Get data from the active row (adjust column indexes as needed)
  const techName = sheet.getRange(row, 3).getValue(); // Column C: Tech name
  const phoneNumber = sheet.getRange(row, 8).getValue(); // Example: Column H could contain phone numbers
  const eventName = sheet.getRange(row, 4).getValue(); // Column D: Event name
  const location = sheet.getRange(row, 5).getValue(); // Column E: Location
  const formDate = Utilities.formatDate(new Date(sheet.getRange(row, 2).getValue()), "GMT+1", "yyyy-MM-dd"); // Column B: Date
  const startTime = Utilities.formatDate(new Date(sheet.getRange(row, 17).getValue()), "GMT+1", "HH:mm"); // Column Q: Start time
  const endTime = Utilities.formatDate(new Date(sheet.getRange(row, 18).getValue()), "GMT+1", "HH:mm"); // Column R: End time
  const vehicle = sheet.getRange(row, 16).getValue(); // Column P: Vehicle
  const description = sheet.getRange(row, 19).getValue(); // Column S: Description
  const formLink = sheet.getRange(row, 6).getValue(); // Column F: Form link
  
  const whatsappUrl = generateWhatsAppLink(
    phoneNumber,
    eventName,
    location,
    formDate,
    startTime,
    endTime,
    vehicle,
    description,
    techName,
    formLink
  );
  
  if (whatsappUrl) {
    // Open the WhatsApp Web link
    const htmlOutput = HtmlService
      .createHtmlOutput(`<script>window.open("${whatsappUrl}", "_blank"); google.script.host.close();</script>`)
      .setWidth(1)
      .setHeight(1);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Apertura WhatsApp Web...");
  } else {
    SpreadsheetApp.getUi().alert("Impossibile creare il link WhatsApp. Verifica che il numero di telefono sia valido.");
  }
}

/**
 * Add a custom menu item to send WhatsApp messages
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Comunicazioni')
    .addItem('Invia Messaggio WhatsApp', 'sendWhatsAppMessage')
    .addToUi();
}