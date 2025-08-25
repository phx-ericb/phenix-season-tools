var Library = {
  importerDonneesSaison: importerDonneesSaison,
  initSeasonFile: initSeasonFile,
  nettoyerConversionsSaison: nettoyerConversionsSaison,

  // Diff v0.7
  diffInscriptions: diffInscriptions_,
  diffArticles: diffArticles_,

  // Emails
  sendPendingOutbox: sendPendingOutbox,

  // Règles
  evaluateSeasonRules: evaluateSeasonRules,

  // Exports — références uniquement
  writeRetroMembresSheet: writeRetroMembresSheet,
  exportRetroMembresXlsxToDrive: exportRetroMembresXlsxToDrive,
  writeRetroGroupesSheet: writeRetroGroupesSheet,
  exportRetroGroupesXlsxToDrive: exportRetroGroupesXlsxToDrive,
  writeRetroGroupeArticlesSheet: writeRetroGroupeArticlesSheet,
  exportRetroGroupeArticlesXlsxToDrive: exportRetroGroupeArticlesXlsxToDrive
};

// wrapper pratique pour lancer l’export
function runExportRetroMembres() {
  return exportRetroMembresXlsxToDrive(SpreadsheetApp.getActiveSpreadsheet().getId());
}
