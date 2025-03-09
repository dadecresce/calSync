// UpdateEvents.gs

function updateEvents(ss, events, techMapping) {
  const config = getConfig();
  const eventsSheet = ss.getSheetByName(config.sheetNames.eventi);
  
  let existingEvents = [];
  if (eventsSheet.getLastRow() >= 2) {
    existingEvents = eventsSheet.getRange(2, 1, eventsSheet.getLastRow() - 1, 24).getValues();
  }
  const existingEventIds = existingEvents.map(row => row[0]);
  
  const eventDataNew = [];
  events.forEach(event => {
    const eventId = event.getId();
    const eventDate = Utilities.formatDate(event.getStartTime(), "GMT+1", "dd-MM-yyyy");
    let title = event.getTitle();
    
    const regex = /^(.*?)\s*(?:-|@)\s*(.*?)\s*\((.*?)\)\s*(.*)?$/;
    const match = title.match(regex);
    
    let idCivette = "", luogo = "", name = "";
    let techNames = [];
    let techExtras = [];
    let noteFromParentheses = "";
    let vehicle = "";
    
    if (match) {
      const group1 = match[1].trim();
      const group1Regex = /^(\d+)\s+(.*)$/;
      const group1Match = group1.match(group1Regex);
      if (group1Match) {
        idCivette = group1Match[1];
        name = group1Match[2];
      } else {
        name = group1;
      }
      luogo = match[2].trim();
      let rawGroup = match[3].trim();
      let tokens = rawGroup.split(/[+\/]/).map(token => token.trim()).filter(token => token !== "");
      let tokenObjects = tokens.map(token => {
        let m = token.match(/^([A-Za-zÀ-ÖØ-öø-ÿ]+)\s*(.*)$/);
        if (m) {
          return { name: m[1], extra: (m[2] && typeof m[2] === 'string') ? m[2].trim() : "" };
        }
        return { name: token, extra: "" };
      });
      let areAllTech = tokenObjects.length > 0 && tokenObjects.every(obj => techMapping.hasOwnProperty(obj.name.toLowerCase()));
      if (areAllTech) {
        tokenObjects.forEach(obj => {
          techNames.push(obj.name);
          techExtras.push(obj.extra);
        });
      } else {
        noteFromParentheses = rawGroup;
      }
      vehicle = match[4] || "";
    } else {
      name = title;
    }
    
    if (!vehicle || vehicle.trim() === "") {
      vehicle = "Mezzo Proprio";
    }
    
    let techColumns = Array(5).fill("");
    let extraColumns = Array(5).fill("");
    if (techNames.length > 0) {
      for (let i = 0; i < techNames.length && i < 5; i++) {
        techColumns[i] = techNames[i];
        extraColumns[i] = techExtras[i] || "";
      }
    } else if (noteFromParentheses) {
      techColumns[0] = noteFromParentheses;
      extraColumns[0] = noteFromParentheses;
    }
    
    const techIDs = techColumns.map(tech => {
      if (tech && tech.trim() !== "" && techMapping.hasOwnProperty(tech.toLowerCase())) {
        return techMapping[tech.toLowerCase()].id;
      }
      return "";
    });
    
    let description = event.getDescription() || "";
    if (noteFromParentheses && techNames.length === 0) {
      description = `[Note: ${noteFromParentheses}] ` + description;
    }
    
    const eventInfo = [
      eventId, idCivette, eventDate, name, luogo
    ].concat(techColumns).concat(techIDs).concat([vehicle, event.getStartTime().toISOString(), event.getEndTime().toISOString(), description]).concat(extraColumns);
    
    if (!existingEventIds.includes(eventId)) {
      eventDataNew.push(eventInfo);
    } else {
      const rowIndex = existingEventIds.indexOf(eventId) + 2;
      eventsSheet.getRange(rowIndex, 1, 1, 24).setValues([eventInfo]);
      eventsSheet.getRange(rowIndex, 1).setBackground('yellow');
    }
  });
  
  if (eventDataNew.length > 0) {
    eventsSheet.getRange(eventsSheet.getLastRow() + 1, 1, eventDataNew.length, 24).setValues(eventDataNew);
  }
  
  return eventsSheet.getRange(2, 1, eventsSheet.getLastRow() - 1, 24).getValues();
}