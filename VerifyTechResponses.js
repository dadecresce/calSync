// VerifyTechResponses.gs

function verifyTechResponses(ss, allEvents) {
  const config = getConfig();
  const eventsSheet = ss.getSheetByName(config.sheetNames.eventi);
  let responsesSheet = ss.getSheetByName(config.sheetNames.risposte);
  if (!responsesSheet) {
    responsesSheet = ss.insertSheet(config.sheetNames.risposte);
  }
  
  let responses = [];
  if (responsesSheet.getLastRow() >= 2) {
    responses = responsesSheet.getRange(2, 1, responsesSheet.getLastRow() - 1, 10).getValues();
  }
  const responseIdLinks = responses.map(row => row[9]);
  
  const today = new Date();
  const oneWeekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
  
  allEvents.forEach((event, index) => {
    const eventId = event[0];
    const date = event[2];
    let eventDate;
    if (typeof date === 'string') {
      const parts = date.split("-");
      eventDate = new Date(parts[2], parts[1] - 1, parts[0]);
    } else if (date instanceof Date) {
      eventDate = date;
    } else {
      eventDate = new Date(date);
    }
    
    const rowIndex = index + 2;
    const technicians = [event[5], event[6], event[7], event[8], event[9]];
    technicians.forEach((tech, techIndex) => {
      if (tech && tech.trim() !== "") {
        const idLink = `${eventId}${tech}`;
        const techCell = eventsSheet.getRange(rowIndex, 6 + techIndex);
        if (responseIdLinks.map(x => x.toString().trim().toLowerCase()).includes(idLink.toString().trim().toLowerCase())) {
          techCell.setBackground('#00FF00');
        } else if (eventDate > oneWeekAgo) {
          techCell.setBackground('#FFFF00');
        } else {
          techCell.setBackground('#FF0000');
        }
      }
    });
  });
}