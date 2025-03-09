// VerifyServiceResponses.gs

function verifyServiceResponses(ss) {
  const config = getConfig();
  const servicesSheet = ss.getSheetByName(config.sheetNames.servizi);
  let responsesSheet = ss.getSheetByName(config.sheetNames.risposte);
  if (!responsesSheet) {
    responsesSheet = ss.insertSheet(config.sheetNames.risposte);
  }
  
  let servicesData = [];
  if (servicesSheet.getLastRow() >= 2) {
    servicesData = servicesSheet.getRange(2, 1, servicesSheet.getLastRow() - 1, 6).getValues();
  }
  
  let responses = [];
  if (responsesSheet.getLastRow() >= 2) {
    responses = responsesSheet.getRange(2, 1, responsesSheet.getLastRow() - 1, 10).getValues();
  }
  const responseIdLinks = responses.map(row => row[9]);
  let normalizedResponseIdLinks = responseIdLinks.map(x => x.toString().trim().toLowerCase());
  
  for (let i = 0; i < servicesData.length; i++) {
    let idLink = servicesData[i][4];
    if (idLink) {
      idLink = idLink.toString().trim().toLowerCase();
      if (normalizedResponseIdLinks.includes(idLink)) {
        servicesSheet.getRange(i + 2, 5).setBackground('green');
      }
    }
  }
}