// ConfigUtils.gs

function getConfig() {
  return {
    calendarIds: [
      'service3civette@gmail.com',
      'udlnhgidrei32nla3ufs8cr6v0@group.calendar.google.com'
    ],
    formUrl: 'https://docs.google.com/forms/d/14ErbViiMIM4DsYpg2Z0nI5JotqlIexxMyGwcCkYHF0g/viewform',
    sheetNames: {
      eventi: 'eventi',
      servizi: 'servizi',
      risposte: 'risposte',
      tecnici: 'tecnici'
    }
  };
}

function getTimeRange(today) {
  const startTime = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 1, 0, 0);
  const daysUntilNextSunday = (today.getDay() === 0) ? 7 : 7 - today.getDay();
  const endTime = new Date(today.getFullYear(), today.getMonth(), today.getDate() + daysUntilNextSunday, 23, 59, 59, 999);
  return { startTime, endTime };
}

function buildTechMapping(ss) {
  const config = getConfig();
  const techSheet = ss.getSheetByName(config.sheetNames.tecnici);
  let techMapping = {};
  if (techSheet) {
    const techData = techSheet.getDataRange().getValues();
    for (let i = 1; i < techData.length; i++) {
      let techID = techData[i][0].toString().trim();      // Colonna A: IDTecnico
      let techName = techData[i][1].toString().trim();    // Colonna B: Nome
      let techPhone = techData[i][2] ? techData[i][2].toString().trim() : ""; // Colonna C: Numero di Telefono
      let techCalendarId = techData[i][3] ? techData[i][3].toString().trim() : ""; // Colonna D: ID Calendario
      if (techName) {
        techMapping[techName.toLowerCase()] = { 
          id: techID, 
          name: techName, 
          phone: techPhone,
          calendarId: techCalendarId
        };
      }
    }
  }
  return techMapping;
}