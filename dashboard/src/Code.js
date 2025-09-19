/** =========================================================================
 *  Dashboard runners – Phénix (v1.1)
 *  - Pas d’appel à getActiveSpreadsheet()
 *  - Passe toujours l’ID du classeur saison à la librairie
 *  - Endpoints UI: getRecentActivity(), getDashboardMetrics()
 *  - Logger unifié appendImportLog_ (1 ou 3 arguments)
 *  - Aligne les runners UI sur les exporteurs «*ToDrive»
 *  - Variante incrémentale via onlyPassports (ou LAST_TOUCHED_PASSPORTS)
 * ========================================================================= */

function seedSeasonYearOnce() {
  var id = getSeasonId_();
  setParamValue('SEASON_YEAR', 2025);
  setParamValue('RETRO_MEMBER_MAX_U', 18); // optionnel
}

// Alias vers la lib (ajuste "SI" si ton alias diffère)
var LIB = (SI && SI.Library) ? SI.Library : null;

/** ============================ Runners principaux ============================ */
/** Code.gs — wrappers exposés à l’UI */

/** Import du fichier le plus récent (Validation_Membres) → upsert dans le CENTRAL.MEMBRES_GLOBAL */
function runImportValidationMembres() {
  var centralId = readParamValue('GLOBAL_MEMBRES_SHEET_ID');
  if (!centralId) {
    throw new Error("PARAMS.GLOBAL_MEMBRES_SHEET_ID est vide : indique l'ID du classeur central MEMBRES_GLOBAL.");
  }
  // IMPORTANT: importValidationMembresToGlobal_ ouvre le classeur CENTRAL (centralId)
  // et écrit dans sa feuille MEMBRES_GLOBAL (pas dans le classeur saison).
  return importValidationMembresToGlobal_(centralId);
}

/** Exports entraîneurs (membres + groupes) */
function runExportEntraineursMembres() {
  return exportRetroEntraineursMembresXlsxToDrive(getSeasonId_(), {});
}
function runExportEntraineursGroupes() {
  return exportRetroEntraineursGroupesXlsxToDrive(getSeasonId_(), {});
}

/** ============================== Outils debug =============================== */
function debugSeasonId() {
  var id = getSeasonId_();
  try {
    var f = DriveApp.getFileById(id);
    Logger.log('Drive OK: name=%s, mime=%s, url=%s, trashed=%s', f.getName(), f.getMimeType(), f.getUrl(), f.isTrashed());
  } catch (e) {
    Logger.log('DriveApp.getFileById FAILED: %s', e);
  }
  try {
    var ss = SpreadsheetApp.openById(id);
    Logger.log('Spreadsheet OK: title=%s, sheets=%s', ss.getName(), ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
  } catch (e) {
    Logger.log('SpreadsheetApp.openById FAILED: %s', e);
  }
}

/** ============================ Helpers internes ============================ */
function _nowMs_() { return (new Date()).getTime(); }

/** Passeports "touchés" par l'import courant */
function _collectTouchedFromImpRes_(impRes) {
  // a) API dédiée de la lib si dispo
  try {
    if (typeof getLastTouchedPassports_ === 'function') {
      var list = getLastTouchedPassports_();
      if (Array.isArray(list) && list.length) return list.slice();
    }
  } catch (e) { /* no-op */ }

  // b) Sinon, analyser le résultat d'import
  var set = new Set();
  try {
    if (impRes && impRes.incrInscriptions && Array.isArray(impRes.incrInscriptions.touchedPassports)) {
      impRes.incrInscriptions.touchedPassports.forEach(function(p){
        var n = normalizePassportPlain8_(p);
        if (n) set.add(n);
      });
    }
    if (impRes && impRes.incrArticles && Array.isArray(impRes.incrArticles.touchedPassports)) {
      impRes.incrArticles.touchedPassports.forEach(function(p){
        var n = normalizePassportPlain8_(p);
        if (n) set.add(n);
      });
    }
  } catch (e) { /* no-op */ }

  return Array.from(set);
}

/** Passeports invalides (photo/casier) présents localement dans MEMBRES_GLOBAL */
function _collectInvalidPassportsFromLocal_() {
  var ss = getSeasonSpreadsheet_(getSeasonId_());
  var sh = ss.getSheetByName(readParamValue('SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL');
  var set = new Set();
  if (!sh || sh.getLastRow() < 2) return set;

  var V = sh.getDataRange().getValues();
  var H = V[0];
  var cPass = H.indexOf('Passeport');
  var cPif  = H.indexOf('PhotoInvalideFlag');
  var cCfl  = H.indexOf('CasierExpireFlag');
  var cPh   = H.indexOf('PhotoExpireLe');

  // Fallback si flags absents : déduire via PhotoExpireLe / CasierExpiré
  var cCas = H.indexOf('CasierExpiré');
  var seasonYear = Number(readParamValue('SEASON_YEAR') || new Date().getFullYear());
  var cutoffNextJan1 = (seasonYear + 1) + '-01-01';

  for (var r = 1; r < V.length; r++) {
    var p = normalizePassportPlain8_(V[r][cPass]);
    if (!p) continue;

    var photoInvalid = (cPif >= 0)
      ? Number(V[r][cPif] || 0)
      : ((String(V[r][cPh] || '') && String(V[r][cPh]) < cutoffNextJan1) ? 1 : 0);

    var casInvalid = (cCfl >= 0)
      ? Number(V[r][cCfl] || 0)
      : ((cCas >= 0) ? Number(V[r][cCas] || 0) : 0);

    if (photoInvalid === 1 || casInvalid === 1) set.add(p);
  }
  return set;
}

/** Liste des passeports "touchés" par le dernier import (fallback sûr) */
function _getLastTouchedPassportsSafe_() {
  try {
    if (typeof getLastTouchedPassports_ === 'function') {
      var list = getLastTouchedPassports_();
      if (Array.isArray(list)) return list.slice();
    }
  } catch (e) { /* no-op */ }
  return []; // fallback silencieux
}

