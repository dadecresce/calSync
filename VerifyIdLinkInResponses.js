// VerifyIdLinkInResponses.gs

function verifyIdLinkInResponses(ss) {
  const config = getConfig();
  const servicesSheet = ss.getSheetByName(config.sheetNames.servizi);
  let responsesSheet = ss.getSheetByName(config.sheetNames.risposte);
  if (!responsesSheet) {
    responsesSheet = ss.insertSheet(config.sheetNames.risposte);
  }
  
  // Recupera gli IDLink dalla colonna E di "servizi"
  let servicesData = [];
  if (servicesSheet.getLastRow() >= 2) {
    servicesData = servicesSheet.getRange(2, 5, servicesSheet.getLastRow() - 1, 1).getValues();
  }
  const serviceIdLinks = servicesData.map(row => row[0].toString().trim().toLowerCase());
  
  // Recupera gli ID dalla colonna A di "risposte"
  let responses = [];
  if (responsesSheet.getLastRow() >= 2) {
    responses = responsesSheet.getRange(2, 1, responsesSheet.getLastRow() - 1, 1).getValues();
  }
  const responseIds = responses.map(row => row[0].toString().trim().toLowerCase());
  
  // Colora la colonna E in "servizi"
  servicesData.forEach((row, index) => {
    const idLink = row[0].toString().trim().toLowerCase();
    const rowIndex = index + 2;
    const idLinkCell = servicesSheet.getRange(rowIndex, 5);
    
    if (idLink && responseIds.includes(idLink)) {
      idLinkCell.setBackground('#00FF00');
    } else {
      idLinkCell.setBackground(null);
    }
  });
}