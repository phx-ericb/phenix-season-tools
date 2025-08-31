/** =========================================================================
 *  Dashboard runners pour la librairie Phénix (v0.7)
 *  - AUCUN appel à getActiveSpreadsheet() ici (dashboard non lié à un Sheets)
 *  - Passe toujours l’ID du classeur saison à la librairie
 *  - Possibilité de stocker l’ID dans Script Properties (PHENIX_SEASON_SHEET_ID)
 * ========================================================================= */

function seedSeasonYearOnce() {
  var id = getSeasonId_(); 
  setParamValue('SEASON_YEAR', 2025);
  setParamValue('RETRO_MEMBER_MAX_U', 18); // optionnel
}


/// Alias vers la lib (ajuste "SI" si ton alias est différent)
var LIB = SI && SI.Library ? SI.Library : null;

/** ======================== Config de la cible saison ======================== */
/** 1) Option A : définis la constante ci-dessous et basta */
var SEASON_SHEET_ID = '1IVVHi17Jyo8jvWtrSuenbPW8IyEZqlY1bXx-WbnXPkk'; // ← colle l'ID du classeur saison ici ou laisse vide

/** 2) Option B : stocke l’ID en Script Property une fois pour toutes
 *    exécute setSeasonSheetIdOnce() une seule fois puis laisse SEASON_SHEET_ID vide
 */
function setSeasonSheetIdOnce() {
  var id = '1IVVHi17Jyo8jvWtrSuenbPW8IyEZqlY1bXx-WbnXPkk';
  var props = PropertiesService.getScriptProperties();
  props.setProperty('PHENIX_SEASON_SHEET_ID', id);
  props.setProperty('ACTIVE_SEASON_ID', id);

  // Ajoute au registre si absent (pour l’UI)
  var ss = SpreadsheetApp.openById(id);
  var list = JSON.parse(props.getProperty('SEASONS_JSON') || '[]');
  if (!list.some(function(s){ return s.id === id; })) {
    list.push({ id:id, title:ss.getName(), url:ss.getUrl() });
    props.setProperty('SEASONS_JSON', JSON.stringify(list));
  }
}


/** Récupère l’ID du classeur saison (constante > ScriptProperty) */
function getSeasonId_() {
  var props = PropertiesService.getScriptProperties();
  var id =
    props.getProperty('ACTIVE_SEASON_ID') ||
    (SEASON_SHEET_ID && String(SEASON_SHEET_ID).trim()) ||
    props.getProperty('PHENIX_SEASON_SHEET_ID') ||
    props.getProperty('SEASON_SPREADSHEET_ID');

  if (!id) {
    throw new Error(
      "Aucun ID de classeur saison. Définis SEASON_SHEET_ID " +
      "ou exécute setSeasonSheetIdOnce() / clique 'Définir active' dans l'UI."
    );
  }
  return String(id).trim();
}


/** ============================ Runners principaux ============================ */
/** Export XLSX — Rétro : Membres */
function runExportRetroMembres() {
  if (!LIB) throw new Error('Librairie indisponible: vérifie l’alias (ex: SI.Library).');
  return LIB.exportRetroMembresXlsxToDrive(getSeasonId_());
}

/** Export XLSX — Rétro : Groupes (si exposé par la lib) */
function runExportRetroGroupes() {
  if (!LIB || typeof LIB.exportRetroGroupesXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroGroupesXlsxToDrive indisponible dans la lib.');
  }
  return LIB.exportRetroGroupesXlsxToDrive(getSeasonId_());
}

/** Export XLSX — Rétro : Groupe Articles (si exposé par la lib) */
function runExportRetroGroupeArticles() {
  if (!LIB || typeof LIB.exportRetroGroupeArticlesXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroGroupeArticlesXlsxToDrive indisponible dans la lib.');
  }
  return LIB.exportRetroGroupeArticlesXlsxToDrive(getSeasonId_());
}

/** Importer/mettre à jour les données (si tu utilises l’import automatisé) */
function runImporterDonneesSaison() {
  if (!LIB || typeof LIB.importerDonneesSaison !== 'function') {
    throw new Error('Fonction importerDonneesSaison indisponible dans la lib.');
  }
  return LIB.importerDonneesSaison(getSeasonId_());
}

/** Appliquer les règles (remplit ERREURS) */
function runEvaluateRules() {
  if (!LIB || typeof LIB.evaluateSeasonRules !== 'function') {
    throw new Error('Fonction evaluateSeasonRules indisponible dans la lib.');
  }
  return LIB.evaluateSeasonRules(getSeasonId_());
}


/** ============================== Outils debug =============================== */
/** Vérifie que la lib expose bien les fonctions clés */
function debugRetroFns() {
  if (!LIB) { Logger.log('LIB indisponible'); return; }
  Logger.log('typeof writeRetroMembresSheet            = %s', typeof LIB.writeRetroMembresSheet);
  Logger.log('typeof exportRetroMembresXlsxToDrive     = %s', typeof LIB.exportRetroMembresXlsxToDrive);
  Logger.log('typeof exportRetroGroupesXlsxToDrive     = %s', typeof LIB.exportRetroGroupesXlsxToDrive);
  Logger.log('typeof exportRetroGroupeArticlesXlsxToDrive = %s', typeof LIB.exportRetroGroupeArticlesXlsxToDrive);
  Logger.log('typeof importerDonneesSaison             = %s', typeof LIB.importerDonneesSaison);
  Logger.log('typeof evaluateSeasonRules               = %s', typeof LIB.evaluateSeasonRules);
  Logger.log('typeof sendPendingOutbox                 = %s', typeof LIB.sendPendingOutbox);
}

/** Sanity-check sur l’ID (fichier/feuille) sans dépendre des helpers de la lib */
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
    Logger.log('Spreadsheet OK: title=%s, sheets=%s', ss.getName(), ss.getSheets().map(function(s){return s.getName();}).join(', '));
  } catch (e) {
    Logger.log('SpreadsheetApp.openById FAILED: %s', e);
  }
}

/** Dernières lignes d’IMPORT_LOG (lecture directe) */
function debug_tailImportLog(n) {
  n = n || 30;
  var ss = SpreadsheetApp.openById(getSeasonId_());
  var sh = ss.getSheetByName('IMPORT_LOG');
  if (!sh || sh.getLastRow() < 2) { Logger.log('IMPORT_LOG vide.'); return; }
  var last = sh.getLastRow();
  var start = Math.max(2, last - n + 1);
  var vals = sh.getRange(start, 1, last - start + 1, Math.min(sh.getLastColumn(), 3)).getValues();
  vals.forEach(function(r){ Logger.log('%s | %s | %s', r[0], r[1], r[2]); });
}

/** Écriture simple d’un param dans PARAMS (sans dépendre des helpers internes de la lib) */
function setParamValue(key, value) {
  var ss = SpreadsheetApp.openById(getSeasonId_());
  var sh = ss.getSheetByName('PARAMS') || ss.insertSheet('PARAMS');
  if (sh.getLastRow() < 1) sh.getRange(1,1,1,4).setValues([['Clé','Valeur','Type','Description']]);

  var last = sh.getLastRow();
  if (last < 2) { sh.appendRow([key, value, '', '']); return; }

  var keys = sh.getRange(2,1,last-1,1).getValues().map(function(r){return String(r[0]||'');});
  var row = -1;
  for (var i=0;i<keys.length;i++){ if (keys[i] === key) { row = 2+i; break; } }
  if (row === -1) sh.appendRow([key, value, '', '']);
  else sh.getRange(row,2).setValue(value);
}

/** Mini smoke test complet : import → règles → export Membres */
function smoke_test() {
  if (!LIB) throw new Error('Librairie indisponible');
  var id = getSeasonId_();
  Logger.log('---- RUN importerDonneesSaison ----');
  if (typeof LIB.importerDonneesSaison === 'function') Logger.log(LIB.importerDonneesSaison(id));
  Logger.log('---- RUN evaluateSeasonRules ----');
  if (typeof LIB.evaluateSeasonRules === 'function') Logger.log(JSON.stringify(LIB.evaluateSeasonRules(id)));
  Logger.log('---- RUN exportRetroMembresXlsxToDrive ----');
  Logger.log(JSON.stringify(LIB.exportRetroMembresXlsxToDrive(id)));
  Logger.log('---- TAIL IMPORT_LOG ----');
  debug_tailImportLog(40);
}

function resetSeasonData(seasonId){
  return _wrap('resetSeasonData', function(){
    var ss = SpreadsheetApp.openById(seasonId);

    var targets = [
      'MODIFS_INSCRIPTIONS','STAGING_INSCRIPTIONS','STAGING_ARTICLES',
      'ARTICLES','INSCRIPTIONS',
      'ANNULATIONS_INSCRIPTIONS','ANNULATIONS_ARTICLES',
      'IMPORT_LOG','MAIL_OUTBOX','ERREURS'
    ];

    var cleared = [];
    targets.forEach(function(name){
      var sh = ss.getSheetByName(name);
      if (!sh) return;
      var last = sh.getLastRow();
      if (last > 1){
        sh.getRange(2,1,last-1, Math.max(1, sh.getLastColumn())).clearContent();
      }
      cleared.push(name);
    });

    return _ok({ cleared: cleared }, 'Réinitialisation complétée');
  });
}

/**
 * Import → Exports rétroaction (groupes & groupe articles) → Log dans IMPORT_LOG.
 * Utilise tes runners existants (runImportDry / runExport).
 */
function runImportAndExports(){
  var seasonId = getSeasonId_(); // déjà défini dans ton projet
  var out = { ok:false, steps: [] };

  // 1) Import (dry)
  var imp = runImportDry(seasonId);           // fourni par server_exports.js
  out.steps.push({ step: 'import', res: imp });

  // 2) Exports rétroaction (XLSX)
  var ex1 = runExport(seasonId, 'groupes');   // fourni par server_exports.js
  out.steps.push({ step: 'export_groupes', res: ex1 });

  var ex2 = runExport(seasonId, 'groupArticles'); // fourni par server_exports.js
  out.steps.push({ step: 'export_groupArticles', res: ex2 });

  // 3) Log IMPORT_LOG
  appendImportLog_({
    type: 'IMPORT',
    details: 'Import exécuté + Exports rétroaction générés (groupes & groupe articles)'
  });

  out.ok = true;
  return out;
}

/** Écrit une ligne dans IMPORT_LOG (Date | Type | Détails). Crée la feuille au besoin. */
function appendImportLog_(entry){
  var ss = SpreadsheetApp.openById(getSeasonId_());
  var name = 'IMPORT_LOG';
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,3).setValues([['Date','Type','Détails']]);
  }
  sh.appendRow([ new Date(), entry.type || 'INFO', entry.details || '' ]);
}
