function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Import de données")
  .addItem("Importer les données d'un autre Google Sheets", "merge")
  .addToUi();
}

function merge() {
  
  function properCase(phrase) {
    return (phrase.toLowerCase().replace(/([^a-zÀÁÂÄÇÈÉÊËÌÍÎÏÒÓÔÖŒÙÚÛÜàáâãäçèéêëìíîïôöùúûüýÿ])([a-zÀÁÂÄÇÈÉÊËÌÍÎÏÒÓÔÖŒÙÚÛÜàáâãäçèéêëìíîïôöùúûüýÿ])(?=[a-zÀÁÂÄÇÈÉÊËÌÍÎÏÒÓÔÖŒÙÚÛÜàáâãäçèéêëìíîïôöùúûüýÿ]{2})|^([a-zÀÁÂÄÇÈÉÊËÌÍÎÏÒÓÔÖŒÙÚÛÜàáâãäçèéêëìíîïôöùúûüýÿ])/g, function(_, g1, g2, g3) {
      return (typeof g1 === 'undefined') ? g3.toUpperCase() : g1 + g2.toUpperCase(); }));
  }
  
  function merge2(sheetscr, sheetdst, columns, hdsidx) {
    
    const heads = sheetscr.getRange(hdsidx, 1, (sheetscr.getLastRow() + 1 - hdsidx), sheetscr.getLastColumn()).getDisplayValues().shift();
    var data = sheetscr.getRange((hdsidx + 1), 1, (sheetscr.getLastRow() - hdsidx), sheetscr.getLastColumn()).getDisplayValues()
    .map(function(row) {row[heads.indexOf("Tél")] = "'" + row[heads.indexOf("Tél")];
                        row[heads.indexOf("Téléphone")] = "'" + row[heads.indexOf("Téléphone")];
                        row[heads.indexOf("Numéro de portable")] = "'" + row[heads.indexOf("Numéro de portable")];
                        row[heads.indexOf("Numéro de tél")] = "'" + row[heads.indexOf("Numéro de tél")];
                        row[heads.indexOf("Nom")] = properCase(row[heads.indexOf("Nom")]);
                        row[heads.indexOf("Prénom")] = properCase(row[heads.indexOf("Prénom")]);
                        return (row);})
    .map(function(r) {var row = []; columns.forEach(function(el) {row.push(r[heads.indexOf(el)]);}); return (row);})
    
    if (data.length > 0) {
      sheetdst.getRange(1, 1, 1, columns.length)
      .setValues([columns])
      .setBackgroundRGB(255, 0, 127);
      sheetdst.getRange(sheetdst.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
    }
  }
  
  sourceURL = Browser.inputBox("Import d'un fichier", 
                            "Entrez l'URL du fichier source :",
                            Browser.Buttons.OK_CANCEL);
  if (sourceURL === "cancel" || sourceURL == "") {
    return;
  }
  
  sourcename = Browser.inputBox("Import d'un fichier", 
                                "Entrez le nom de l'onglet à importer :",
                                Browser.Buttons.OK_CANCEL);
  if (sourcename === "cancel" || sourcename == "") {
    return;
  }
  
  colonnesstr = Browser.inputBox("Import d'un fichier", 
                              "Entrez le nom des colonnes à importer en les séparant d'une virgule (pour toutes les importer, entrez 'Toutes' sans guillemets):",
                              Browser.Buttons.OK_CANCEL);
  if (colonnesstr === "cancel" || colonnesstr == "") {
    return;
  }
  
  position_header = Browser.inputBox("Import d'un fichier", 
                                     "Entrez le numéro de la ligne d'en-tête :",
                                     Browser.Buttons.OK_CANCEL);
  if (position_header === "cancel" || position_header == "") {
    return;
  }
  
  const source = SpreadsheetApp.openByUrl(sourceURL).getSheetByName(sourcename),
      destination = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (colonnesstr.toLowerCase() == "toutes") {
    var colonnes = source.getRange(position_header, 1, 1, source.getLastColumn()).getDisplayValues().shift()
    } else {
      colonnes = colonnesstr.split(", ");
    }
  
  merge2(source, destination, colonnes, position_header);
  
}
