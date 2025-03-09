// UpdateServices.gs

function updateServices(ss, allEvents, techMapping) {
    const config = getConfig();
    const servicesSheet = ss.getSheetByName(config.sheetNames.servizi);
    
    let existingServices = [];
    if (servicesSheet.getLastRow() >= 2) {
      existingServices = servicesSheet.getRange(2, 1, servicesSheet.getLastRow() - 1, 7).getValues();
    }
    const existingIdLinks = existingServices.map(row => row[4]);
    
    const serviceDataNew = [];
    const formUrl = config.formUrl;
    
    allEvents.forEach(event => {
      const eventId = event[0];
      const eventName = event[3]; // Titolo evento (D)
      const luogo = event[4];     // Luogo (E)
      const vehicle = event[15];
      const startTimeISO = event[16];
      const endTimeISO = event[17];
      const description = event[18];
      const extraInfos = event.slice(19, 24);
      
      const formDate = Utilities.formatDate(new Date(startTimeISO), "GMT+1", "yyyy-MM-dd");
      const startTimeFormatted = Utilities.formatDate(new Date(startTimeISO), "GMT+1", "HH:mm");
      const endTimeFormatted = Utilities.formatDate(new Date(endTimeISO), "GMT+1", "HH:mm");
      const sentinelValue = (luogo === "Teatro Cervia") ? "Sì" : "No";
      
      for (let i = 0; i < 5; i++) {
        const rawTech = event[5 + i];
        const techID = event[10 + i];
        if (rawTech && rawTech.trim() !== "" && techMapping.hasOwnProperty(rawTech.toLowerCase())) {
          let standardizedTech = techMapping[rawTech.toLowerCase()].name;
          let techPhone = techMapping[rawTech.toLowerCase()].phone || "";
          let extraInfo = extraInfos[i] || "";
          extraInfo = (typeof extraInfo === "string") ? extraInfo.trim() : "";
          let noteText = extraInfo ? extraInfo : (standardizedTech.toLowerCase().includes("mont") ? "Solo Montaggio" : "");
          const fullDescription = noteText ? (noteText + " " + description) : description;
          const idLink = `${eventId}${rawTech}`;
          
          // Link per il Google Form (Titolo prima di Luogo, già corretto)
          const prefilledUrl =
            `${formUrl}?entry.392873233=${encodeURIComponent(idLink)}` +
            `&entry.70294579=${encodeURIComponent(eventName)}` + // Titolo
            `&entry.1074121414=${encodeURIComponent(luogo)}` +   // Luogo
            `&entry.418786620=${encodeURIComponent(formDate)}` +
            `&entry.1259553055=${encodeURIComponent(rawTech)}` +
            `&entry.1518244967=${encodeURIComponent(vehicle)}` +
            `&entry.2116091289=${encodeURIComponent(vehicle)}` +
            `&entry.1222319648=${encodeURIComponent(startTimeFormatted)}` +
            `&entry.776840624=${encodeURIComponent(endTimeFormatted)}` +
            `&entry.204817070=${encodeURIComponent(fullDescription)}` +
            `&entry.1169178161=${encodeURIComponent(sentinelValue)}`;
          
          // Link WhatsApp (Titolo prima di Luogo, allineato al form)
          let whatsappUrl = "";
          if (techPhone) {
            const message = encodeURIComponent(
              `Ciao ${standardizedTech}, ecco il tuo evento:\n` +
              `Luogo: ${eventName}\n` + // Titolo evento
              `Titolo: ${luogo}\n` +      // Luogo
              `Data: ${formDate}\n` +
              `Inizio: ${startTimeFormatted}\n` +
              `Fine: ${endTimeFormatted}\n` +
              `Mezzo: ${vehicle}\n` +
              `Descrizione: ${fullDescription}\n` +
              `Link: ${prefilledUrl}`
            );
            whatsappUrl = `https://wa.me/${techPhone}?text=${message}`;
          }
          
          const serviceRowData = [
            eventId,
            techID,
            standardizedTech,
            noteText,
            idLink,
            prefilledUrl,
            whatsappUrl
          ];
          
          if (existingIdLinks.indexOf(idLink) > -1) {
            const indexService = existingIdLinks.indexOf(idLink);
            const rowIndex = indexService + 2;
            servicesSheet.getRange(rowIndex, 1, 1, 7).setValues([serviceRowData]);
            servicesSheet.getRange(rowIndex, 1).setBackground('yellow');
          } else {
            serviceDataNew.push(serviceRowData);
          }
        }
      }
    });
    
    if (serviceDataNew.length > 0) {
      servicesSheet.getRange(servicesSheet.getLastRow() + 1, 1, serviceDataNew.length, 7).setValues(serviceDataNew);
    }
  }