function downloadProjectAsZip() {
  var projectId = ScriptApp.getScriptId();
  var url = "https://script.googleapis.com/v1/projects/" + projectId + "/content";

  var options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (!json.files) {
    Logger.log("Errore: Nessun file trovato nel progetto.");
    return;
  }

  // Trova la cartella del file Google Sheets
  var sheetFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  var parentFolder = sheetFile.getParents().next(); // Prende la cartella in cui si trova il file

  // Crea l'archivio ZIP
  var zipContent = [];
  
  json.files.forEach(function(file) {
    zipContent.push(Utilities.newBlob(file.source, "text/plain", file.name + "." + file.type.toLowerCase()));
  });

  var zipFile = Utilities.zip(zipContent, "GoogleAppsScript.zip");
  
  // Salva lo ZIP nella cartella del Google Sheets
  var zipDriveFile = parentFolder.createFile(zipFile);
  
  Logger.log("ZIP salvato in: " + zipDriveFile.getUrl());
}
function debugDownloadProject() {
  var projectId = ScriptApp.getScriptId();
  var url = "https://script.googleapis.com/v1/projects/" + projectId + "/content";

  var options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText()); // Controlla cosa restituisce l'API
}
