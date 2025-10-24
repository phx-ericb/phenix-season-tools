/** =========================================================
 * server_entrypoint.js  —  Point d'entrée minimal & commun
 * - doGet() pour servir index/ui
 * - include()/includeEval() pour les templates HTML
 * - API_listApis() pour introspection
 * - API_ping() fumée rapide
 * - Helpers génériques partagés (norm/resolve/open)
 * 
 * ⚠️ Aucune API métier ici (JOUEURS / ERREURS / ANNULATIONS).
 *    Celles-ci doivent vivre dans leurs fichiers dédiés :
 *    - server_joueurs.js
 *    - server_errors.js
 *    - server_annulations.js
 * ========================================================= */

/* ---------- Logger shim (optionnel) ---------- */
if (typeof log_ !== 'function') {
  var log_ = function (code, details) {
    try {
      var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      Logger.log('[' + ts + '][' + code + '] ' + (details || ''));
      // Optionnel : journaliser dans IMPORT_LOG si présent
      try {
        if (typeof getSeasonSpreadsheet_ === 'function' && typeof getSeasonId_ === 'function') {
          var ss = getSeasonSpreadsheet_(getSeasonId_());
          var sh = ss && ss.getSheetByName('IMPORT_LOG');
          if (sh) sh.appendRow([ts, String(code || ''), String(details || '')]);
        }
      } catch (e) { /* silencieux */ }
    } catch (e) {}
  };
}

/* ---------- Web entry ---------- */
function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || '';
  var file = (view === 'ui') ? 'ui' : 'index';
  return HtmlService.createTemplateFromFile(file).evaluate()
    .setTitle('Gestion Phénix')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Templates helpers ---------- */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
function includeEval(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

/* ---------- Introspection & fumée ---------- */
function API_listApis() {
  // Liste toutes les fonctions globales qui commencent par API_
  return Object.getOwnPropertyNames(globalThis)
    .filter(function (n) { return /^API_/.test(n) && typeof globalThis[n] === 'function'; })
    .sort();
}
function API_ping() { return 'pong'; }

/* ---------- Helpers génériques partagés ---------- */

// Normalisation d’étiquettes (accents -> ASCII, minuscule, sans ponctuation)
// Utilisable par d’autres modules (JOUEURS/ERREURS/ANNULATIONS)
var ER_norm = this.ER_norm || function (s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '')
    .trim();
};

// Résolution de l’ID de saison
// Préférence : getSeasonId_() si présent, sinon ScriptProperties.SEASON_SHEET_ID
var ER_resolveSeasonId_ = this.ER_resolveSeasonId_ || function (overrideId) {
  if (overrideId) return overrideId;
  if (typeof getSeasonId_ === 'function') {
    var sid = getSeasonId_();
    if (sid) return sid;
  }
  var sidProp = PropertiesService.getScriptProperties().getProperty('SEASON_SHEET_ID');
  if (sidProp) return sidProp;
  throw new Error("Aucun ID de saison : définis SEASON_SHEET_ID ou implémente getSeasonId_().");
};

// Ouverture du classeur de saison
// Préférence : getSeasonSpreadsheet_() si présent, sinon openById
var ER_openSeasonSpreadsheet_ = this.ER_openSeasonSpreadsheet_ || function (seasonId) {
  if (typeof getSeasonSpreadsheet_ === 'function') {
    var ss = getSeasonSpreadsheet_(seasonId);
    if (ss) return ss;
  }
  return SpreadsheetApp.openById(seasonId);
};
