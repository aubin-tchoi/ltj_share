function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Archivage")
  .addItem("Archiver le fichier", "archiving")
  .addToUi();
}

function archiving() {
  
  const ssOld = SpreadsheetApp.getActiveSpreadsheet();
  const ssNew = SpreadsheetApp.create(ssOld.getName());
  ssOld.getSheets().forEach(function(s) {s.copyTo(ssNew);});
  
  var file = DriveApp.getFileById(ssNew.getId());
  DriveApp.getFolderById("1LOk5WHg2hwLNVwuLG31625JY-nfUbHrW").addFile(file);
  
  ssNew.deleteSheet(ssNew.getSheetByName("Feuille 1"));
  ssNew.getSheets().forEach(function(s) {s.setName(s.getName().match(/Copie de (.+)/)[1]);});
  
  confirm = Browser.msgBox("Archivage", 
                           "Le fichier a été archivé dans le dossier 'Archives', son contenu sera maintenant effacé.",
                           Browser.Buttons.OK_CANCEL);
  if (confirm == "cancel") {
    return;
  }
  
  try {
    ssOld.getSheets().forEach(function(s) {s.getRange(2, 1, (s.getLastRow() - 1), s.getLastColumn())
    .clearContent()
    .setBackground('white');
                                          });
  } catch(e) {Logger.log(e)}
  
}
