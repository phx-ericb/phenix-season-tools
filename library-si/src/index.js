// index.js
// ⬇️ si ces symboles viennent de source.js, ils existent déjà globalement :
/* global _rulesBuildErrorsFast_, _rulesWriteFull_, _toPassportSet_, getSeasonId_, getSeasonSpreadsheet_ */

function _rulesBuildErrorsIncrFast_(passports, ss) {
  ss = ss || getSeasonSpreadsheet_(getSeasonId_());
  var set = (typeof _toPassportSet_ === 'function') ? _toPassportSet_(passports) : new Set((passports||[]).map(String));
  return _rulesBuildErrorsFast_(ss, set);
}

// (optionnel) runner FULL coté lib – pratique pour tests directs
function runEvaluateRulesFast_(ss) {
  ss = ss || getSeasonSpreadsheet_(getSeasonId_());
  var res = _rulesBuildErrorsFast_(ss, /*touchedSet*/ null);
  _rulesWriteFull_(ss, res.errors, res.header);
  return { written: res.errors.length, ledger: res.ledgerCount, joueurs: res.joueursCount };
}

var Library = {
  // Import
  importerDonneesSaison: importerDonneesSaison,
  initSeasonFile: initSeasonFile,
  nettoyerConversionsSaison: nettoyerConversionsSaison,

  // Diff v0.7
  diffInscriptions: diffInscriptions_,
  diffArticles: diffArticles_,

  // Emails / Queue
  sendPendingOutbox: sendPendingOutbox,

  // FAST mail pipeline helpers (noms canoniques)
  enqueueWelcomeFromJoueursFast_: (typeof enqueueWelcomeFromJoueursFast_ === 'function') ? enqueueWelcomeFromJoueursFast_ : undefined,
  enqueueValidationMailsFromErreursFast_: (typeof enqueueValidationMailsFromErreursFast_ === 'function') ? enqueueValidationMailsFromErreursFast_ : undefined,

  // --- Back-compat API (alias vers FAST) ---
  enqueueInscriptionNewBySectors: (typeof enqueueWelcomeFromJoueursFast_ === 'function')
    ? function(seasonId){ return enqueueWelcomeFromJoueursFast_(seasonId); }
    : undefined,
  enqueueValidationEmailsByErrorCode: (typeof enqueueValidationMailsFromErreursFast_ === 'function')
    ? function(seasonId, code){ return enqueueValidationMailsFromErreursFast_(seasonId, code); }
    : undefined,

  // Règles
  evaluateSeasonRules: evaluateSeasonRules,             // legacy (fallback)
  runEvaluateRulesFast_: runEvaluateRulesFast_,         // FAST: full build+write
  _rulesBuildErrorsIncrFast_: _rulesBuildErrorsIncrFast_, // FAST: incr build (touched)

  // Exports — références uniquement
  writeRetroMembresSheet: writeRetroMembresSheet,
  exportRetroMembresXlsxToDrive: exportRetroMembresXlsxToDrive,
  writeRetroGroupesSheet: writeRetroGroupesSheet,
  exportRetroGroupesXlsxToDrive: exportRetroGroupesXlsxToDrive,
  writeRetroGroupeArticlesSheet: writeRetroGroupeArticlesSheet,
  exportRetroGroupeArticlesXlsxToDrive: exportRetroGroupeArticlesXlsxToDrive,

  // Augmentations
  runPostImportAugmentations: runPostImportAugmentations_,

    // --- Late registration utilities (exported from external_utils.js) ---
  late_onEditHandler: (typeof late_onEditHandler === 'function') ? late_onEditHandler : undefined,
  late_sendAssignmentEmail: (typeof late_sendAssignmentEmail === 'function') ? late_sendAssignmentEmail : undefined,
   // --- ÉCRITURE dans la feuille tardive (appelée par __sendSummariesV2__) ---
  // (Expose-la surtout si tu veux aussi la tester/appeler manuellement)
  writeLateRegistrations_: (typeof writeLateRegistrations_ === 'function') ? writeLateRegistrations_ : undefined,

  // --- Optionnel : sync générique si tu veux l’appeler depuis ton flow d’import ---
  late_syncPlayersToTargetSheet: (typeof late_syncPlayersToTargetSheet === 'function') ? late_syncPlayersToTargetSheet : undefined,


};

// wrapper pratique pour lancer l’export
function runExportRetroMembres() {
  return exportRetroMembresXlsxToDrive(SpreadsheetApp.getActiveSpreadsheet().getId());
}

