function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Archivage")
  .addItem("Archiver le fichier", "archiving")
  .addToUi();
}

function archiving() {
  
  const ssOld = SpreadsheetApp.getActiveSpreadsheet(),
      ssNew = SpreadsheetApp.create(ssOld.getName()),
      erase_from = 2,
      folder_id = "0AKsJ624j8TbXUk9PVA";
  
  ssOld.getSheets().forEach(function(s) {s.copyTo(ssNew);});
  
  var file = DriveApp.getFileById(ssNew.getId());
  DriveApp.getFolderById(folder_id).addFile(file);
  
  ssNew.deleteSheet(ssNew.getSheetByName("Feuille 1"));
  ssNew.getSheets().forEach(function(s) {s.setName(s.getName().match(/Copie de (.+)/)[1]);});
  
  confirm = Browser.msgBox("Archivage", 
                           "Le fichier a été archivé dans le dossier 'Archives', son contenu sera maintenant effacé.",
                           Browser.Buttons.OK_CANCEL);
  if (confirm == "cancel") {
    return;
  }
  
  try {
    ssOld.getSheets().forEach(function(s) {s.getRange(erase_from, 1, (s.getLastRow() + 1 - erase_from), s.getLastColumn())
    .clearContent()
    .setBackground('white');
                                          });
  } catch(e) {Logger.log(e)}
  
}
