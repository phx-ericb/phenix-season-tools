// --- Shim logger global: évite "ReferenceError: log_ is not defined"
if (typeof log_ !== 'function') {
  var log_ = function(code, details) {
    try {
      var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      Logger.log('[' + ts + '][' + code + '] ' + (details || ''));
      // Si tu veux écrire aussi dans la feuille IMPORT_LOG du classeur de saison:
      try {
        var ss = SpreadsheetApp.openById(getSeasonId_());
        var sh = ss.getSheetByName('IMPORT_LOG');
        if (sh) sh.appendRow([ts, String(code || ''), String(details || '')]);
      } catch (e) { /* silencieux si la feuille n’existe pas */ }
    } catch (e) {
      // rien d'autre à faire
    }
  };
}

function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || '';
  var file = (view === 'ui') ? 'ui' : 'index';
  return HtmlService.createTemplateFromFile(file).evaluate()
    .setTitle('Back-Office Phénix')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// utilitaire include() inchangé
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function includeEval(filename) {
  // Évalue le fichier comme un template Apps Script, puis renvoie son HTML résolu
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}
