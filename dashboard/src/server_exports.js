/* ============================== server_exports.js — v1.2 ==============================
 * - Exports complets garantis: on neutralise LAST_TOUCHED_PASSPORTS quand nécessaire
 * - Respecte le flag INCREMENTAL_ON (PARAMS) pour couper l'incrémental
 * - Conserve la compatibilité avec l’UI existante
 * ============================================================================== */

function _getTouchedPassportsArray_(){
  try{
    var s = (PropertiesService.getDocumentProperties().getProperty('LAST_TOUCHED_PASSPORTS') || '').trim();
    if (!s) return [];
    if (s[0] === '[') { // JSON
      var arr = JSON.parse(s);
      return (Array.isArray(arr)?arr:[]).map(function(x){return String(x||'').trim();}).filter(Boolean);
    }
    return s.split(',').map(function(x){return String(x||'').trim();}).filter(Boolean);
  }catch(e){ return []; }
}

/** Désactive temporairement l’incrémental basé sur LAST_TOUCHED_PASSPORTS pour la durée du callback. */
function _withIncrementalDisabled_(fn){
  var dp = PropertiesService.getDocumentProperties();
  var key = 'LAST_TOUCHED_PASSPORTS';
  var had = dp.getProperty(key);
  try {
    if (had !== null && had !== undefined) dp.deleteProperty(key);
    return fn();
  } finally {
    if (had !== null && had !== undefined) dp.setProperty(key, had);
  }
}
/** Import-only (scan → staging → finals → archive) via LIB.
 * Pas de startImportRun_ ici : le flow entoure déjà l'appel.
 */



function runImporterDonneesSaison() {
  var seasonId = getSeasonId_();
  var ssSeason = getSeasonSpreadsheet_(seasonId);

  // --- Central : ID à récupérer dans les params de la saison
  var centralId = readParam_(ssSeason, 'GLOBAL_MEMBRES_SHEET_ID') || '';
  if (!centralId) {
    throw new Error('Paramètre GLOBAL_MEMBRES_SHEET_ID manquant dans le fichier de saison.');
  }

  // 0) Mise à jour des MEMBRES_GLOBAL central à partir du fichier XLSX Spordle (Validation_Membres)
  try {
    importValidationMembresToGlobal_(centralId);
    appendImportLog_(ssSeason, 'VM_IMPORT_OK', { central: centralId });
  } catch (e0) {
    appendImportLog_(ssSeason, 'VM_IMPORT_FAIL', String(e0));
    // on n’arrête pas complètement, mais note qu’ici si le central n’est pas à jour → la suite verra les vieilles données
  }

  // 1) Synchro du sous-ensemble saison à partir du central
  try {
    var syncRes = syncMembresGlobalSubsetFromCentral_(seasonId, centralId);
    appendImportLog_(ssSeason, 'SYNC_MEMBRES_GLOBAL_OK', syncRes);
  } catch (e1) {
    appendImportLog_(ssSeason, 'SYNC_MEMBRES_GLOBAL_FAIL', String(e1));
  }

  // 2) Import classique via la lib (prioritaire)
  if (typeof LIB !== 'undefined' && LIB && typeof LIB.importerDonneesSaison === 'function') {
    var res = LIB.importerDonneesSaison(seasonId);
    try { appendImportLog_(ssSeason, 'RUN_IMPORT_LIB', { ok: !!res && res.ok === true }); } catch (_) {}
    return res;
  }

  // 3) Repli global si la fonction est chargée sans namespace
  if (typeof importerDonneesSaison === 'function') {
    var res2 = importerDonneesSaison(seasonId);
    try { appendImportLog_(ssSeason, 'RUN_IMPORT_GLOBAL', { ok: !!res2 && res2.ok === true }); } catch (_) {}
    return res2;
  }

  throw new Error('importerDonneesSaison introuvable (ni LIB.importerDonneesSaison, ni globale).');
}


/** Export XLSX — Rétro : Membres (COMPLET GARANTI) */
function runExportRetroMembres() {
  if (!LIB || typeof LIB.exportRetroMembresXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroMembresXlsxToDrive indisponible dans la lib.');
  }
  // Certains implémentations lisent LAST_TOUCHED_PASSPORTS si présent.
  // On le neutralise temporairement pour forcer un export complet.
  var seasonId = getSeasonId_();
  return _withIncrementalDisabled_(function(){
    return LIB.exportRetroMembresXlsxToDrive(seasonId);
  });
}

/** Export XLSX — Rétro : Membres (FORCÉ incrémental, si INCREMENTAL_ON=TRUE et set non vide) */
function runExportRetroMembresIncr() {
  if (!LIB || typeof LIB.exportRetroMembresXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroMembresXlsxToDrive indisponible dans la lib.');
  }
  if (!_retroIncrementalOn_()) {
    // Incrémental désactivé par param → on bascule en complet
    return runExportRetroMembres();
  }
  var list = (typeof getLastTouchedPassports_==='function') ? getLastTouchedPassports_() : _getTouchedPassportsArray_();
  if (!Array.isArray(list) || list.length === 0) {
    // Rien à incrémenter → complet
    return runExportRetroMembres();
  }
  return LIB.exportRetroMembresXlsxToDrive(getSeasonId_(), { onlyPassports: list });
}

/** Export XLSX — Rétro : Groupes (ALL = Groupes + GroupeArticles) — COMPLET GARANTI */
function runExportRetroGroupes() {
  if (!LIB || typeof LIB.exportRetroGroupesAllXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroGroupesAllXlsxToDrive indisponible dans la lib.');
  }
  var seasonId = getSeasonId_();
  return _withIncrementalDisabled_(function(){
    return LIB.exportRetroGroupesAllXlsxToDrive(seasonId);
  });
}

/** Export XLSX — Rétro : Groupes (ALL) FORCÉ incrémental, si INCREMENTAL_ON=TRUE et set non vide */
function runExportRetroGroupesIncr() {
  if (!LIB || typeof LIB.exportRetroGroupesAllXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroGroupesAllXlsxToDrive indisponible dans la lib.');
  }
  if (!_retroIncrementalOn_()) {
    return runExportRetroGroupes();
  }
  var list = (typeof getLastTouchedPassports_==='function') ? getLastTouchedPassports_() : _getTouchedPassportsArray_();
  if (!Array.isArray(list) || list.length === 0) {
    return runExportRetroGroupes();
  }
  return LIB.exportRetroGroupesAllXlsxToDrive(getSeasonId_(), { onlyPassports: list });
}

// Runner unique : Import -> Exports rétro (Membres + Groupes ALL) -> Log + Sync coachs
function runImportAndExports(){
  var out = { ok:false, steps: [] };

  var seasonId = getSeasonId_();
  var ctx = startImportRun_({ seasonId: seasonId, source: 'dashboard' });

  try {
    // 1) Import
    if (!LIB || typeof LIB.importerDonneesSaison !== 'function') {
      throw new Error('Fonction importerDonneesSaison indisponible dans la lib.');
    }
    var t0 = _nowMs_();
    var impRes = LIB.importerDonneesSaison(seasonId);
    var t1 = _nowMs_();

    appendImportLog_({
      type: 'IMPORT_DONE',
      details: { runId: ctx.runId, elapsedMs: (t1 - t0), summary: summarizeImportResult_(impRes) }
    });
    out.steps.push({ step: 'import', res: impRes, elapsedMs: (t1 - t0) });

    // 2) Exports rétro — on appelle les versions COMPLÈTES (elles neutralisent l’incrémental si présent)
    try {
      var t2 = _nowMs_();
      var mRes = runExportRetroMembres();
      var t3 = _nowMs_();
      appendImportLog_({
        type:'EXPORT_RETRO_MEMBRES_OK',
        details: { runId: ctx.runId, elapsedMs: (t3 - t2), name: mRes && mRes.name, rows: mRes && mRes.rows, filtered: mRes && mRes.filtered }
      });
      out.steps.push({ step:'export_membres', res:mRes, elapsedMs:(t3 - t2) });
    } catch(eM){
      appendImportLog_({ type:'EXPORT_RETRO_MEMBRES_FAIL', details: { runId: ctx.runId, error: String(eM) } });
    }

    try {
      var t4 = _nowMs_();
      var gRes = runExportRetroGroupes();
      var t5 = _nowMs_();
      appendImportLog_({
        type:'EXPORT_RETRO_GROUPES_ALL_OK',
        details: { runId: ctx.runId, elapsedMs: (t5 - t4), name: gRes && gRes.name, rows: gRes && gRes.rows, filtered: gRes && gRes.filtered }
      });
      out.steps.push({ step:'export_groupes_all', res:gRes, elapsedMs:(t5 - t4) });
    } catch(eG){
      appendImportLog_({ type:'EXPORT_RETRO_GROUPES_ALL_FAIL', details: { runId: ctx.runId, error: String(eG) } });
    }

    // 3) Sync coachs
    var t6 = _nowMs_();
    try {
      var x = SR_syncInscriptionsEntraineurs_(SpreadsheetApp.openById(seasonId));
      var t7 = _nowMs_();
      appendImportLog_({ type:'COACHS_SYNC_DONE', details: { runId: ctx.runId, rows: (x && x.total || 0), elapsedMs:(t7 - t6) } });
    } catch(e) {
      appendImportLog_({ type:'COACHS_SYNC_FAIL', details: { runId: ctx.runId, error: String(e) } });
    }

    out.ok = true;
    return out;

  } finally {
    endImportRun_(ctx);
  }
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
  Logger.log(JSON.stringify(_withIncrementalDisabled_(function(){ return LIB.exportRetroMembresXlsxToDrive(id); })));
  Logger.log('---- TAIL IMPORT_LOG ----');
  debug_tailImportLog(40);
}

/** Appliquer les règles (remplit ERREURS) */
function runEvaluateRules() {
  if (!LIB || typeof LIB.evaluateSeasonRules !== 'function') {
    throw new Error('Fonction evaluateSeasonRules indisponible dans la lib.');
  }
  return LIB.evaluateSeasonRules(getSeasonId_());
}

function runImportDry(seasonId) {
  return _wrap('runImportDry', function(){
    var params=getParams(seasonId); if(!params.ok) throw new Error(params.error);
    var was=!!params.data.DRY_RUN; setParams(seasonId, Object.assign({}, params.data, { DRY_RUN:true }));
    var sum; try { sum = SI.Library.importerDonneesSaison(seasonId); } finally { setParams(seasonId, Object.assign({}, params.data, { DRY_RUN:was })); }
    return _ok(sum || { ok:true }, 'Import (dry) done');
  });
}

function testRules(seasonId) {
  return _wrap('testRules', function(){
    var res = SI.Library.evaluateSeasonRules(seasonId);
    return _ok(res || { ok:true, errors:0, warns:0, total:0 }, 'Rules evaluated');
  });
}

/**
 * Lance un export complet (ou incrémental si options.onlyPassports est fourni et autorisé)
 * type ∈ { 'members', 'groupes', 'groupArticles' }
 */
function runExport(seasonId, type, options) {
  return _wrap('runExport', function(){
    options = options || {};
    var incrRequested = Array.isArray(options.onlyPassports) && options.onlyPassports.length > 0;
    var allowIncr = _retroIncrementalOn_();

    if (type==='members'){
      if(!SI.Library.exportRetroMembresXlsxToDrive) throw new Error('exportRetroMembresXlsxToDrive not available');
      if (incrRequested && allowIncr) {
        return _ok(SI.Library.exportRetroMembresXlsxToDrive(seasonId, { onlyPassports: options.onlyPassports }), 'Export Membres (incr) lancé');
      }
      return _ok(_withIncrementalDisabled_(function(){ return SI.Library.exportRetroMembresXlsxToDrive(seasonId); }), 'Export Membres (full) lancé');
    }

    if (type==='groupes'){
      if(!SI.Library.exportRetroGroupesAllXlsxToDrive) throw new Error('exportRetroGroupesAllXlsxToDrive not available');
      if (incrRequested && allowIncr) {
        return _ok(SI.Library.exportRetroGroupesAllXlsxToDrive(seasonId, { onlyPassports: options.onlyPassports }), 'Export Groupes (incr) lancé');
      }
      return _ok(_withIncrementalDisabled_(function(){ return SI.Library.exportRetroGroupesAllXlsxToDrive(seasonId); }), 'Export Groupes (full) lancé');
    }

    if (type==='groupArticles'){
      if(!SI.Library.exportRetroGroupeArticlesXlsxToDrive) throw new Error('exportRetroGroupeArticlesXlsxToDrive not available');
      if (incrRequested && allowIncr) {
        return _ok(SI.Library.exportRetroGroupeArticlesXlsxToDrive(seasonId, { onlyPassports: options.onlyPassports }), 'Export Groupe Articles (incr) lancé');
      }
      return _ok(_withIncrementalDisabled_(function(){ return SI.Library.exportRetroGroupeArticlesXlsxToDrive(seasonId); }), 'Export Groupe Articles (full) lancé');
    }

    throw new Error('Unknown export type: ' + type);
  });
}

/** Version pratique: force l’incrémental en lisant LAST_TOUCHED_PASSPORTS (si autorisé) */
function runExportIncr(seasonId, type){
  if (!_retroIncrementalOn_()) {
    // Incrémental off → export complet
    return runExport(seasonId, type, {});
  }
  var list = _getTouchedPassportsArray_();
  if (!list.length) {
    // Rien à incrémenter → complet
    return runExport(seasonId, type, {});
  }
  return runExport(seasonId, type, { onlyPassports: list });
}

function _retroIncrementalOn_() {
  var v = String(readParamValue('INCREMENTAL_ON') || '1').toLowerCase();
  return v === '1' || v === 'true' || v === 'yes' || v === 'oui';
}

function previewExport(seasonId, type, maxRows) {
  return _wrap('previewExport', function(){
    maxRows = maxRows || 20;
    var out;
    if (type==='members')            out = SI.Library.buildRetroMembresRows(seasonId);
    else if (type==='groupes')       out = SI.Library.buildRetroGroupesRows(seasonId);
    else if (type==='groupArticles') out = SI.Library.buildRetroGroupeArticlesRows(seasonId);
    else throw new Error('Unknown export type: ' + type);

    var header = out.header || [];
    var rows   = out.rows || [];
    return _ok({ header: header, total: rows.length, preview: rows.slice(0, maxRows), nbCols: out.nbCols || header.length }, 'Preview ready');
  });
}

// ---- Courriels d’erreur (réutilise le moteur de secteurs) ----
function previewErrorForPassport(seasonId, secteurId, passport, errorItem){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var row = _findRowByPassport_(ss, passport);
    if(!row) throw new Error('Passeport introuvable: ' + passport);

    var sectors = getMailSectors(seasonId);
    if(!sectors.ok) throw new Error(sectors.error || 'Erreur lecture secteurs');
    var items = sectors.data.items || [];
    var it = null;
    for (var i=0;i<items.length;i++){ if(String(items[i].SecteurId||'')===String(secteurId||'')){ it=items[i]; break; } }
    if(!it) throw new Error('Secteur introuvable: ' + secteurId);

    var payload = buildDataFromRow_(row);
    errorItem = errorItem || {};
    payload.error_code = String(errorItem.code || '');
    payload.error_label = String(errorItem.label || '');
    payload.error_details = String(errorItem.details || '');

    var sb = _resolveSubjectBody_(ss, it, payload);
    var to = _resolveToForRow_(ss, row, it);

    return { ok:true, data:{ subject: sb.subject, bodyHtml: sb.bodyHtml, to: to } };
  } catch(e){ return { ok:false, error:String(e) }; }
}

function sendErrorTest(seasonId, item, passport, toTest, errorItem){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var row = _findRowByPassport_(ss, passport);
    if(!row) throw new Error('Passeport introuvable: ' + passport);

    var it = Object.assign({}, item || {});
    var payload = buildDataFromRow_(row);
    errorItem = errorItem || {};
    payload.error_code = String(errorItem.code || '');
    payload.error_label = String(errorItem.label || '');
    payload.error_details = String(errorItem.details || '');

    var sb = _resolveSubjectBody_(ss, it, payload);
    var to = String(toTest||'').trim(); if (!to) to = Session.getActiveUser().getEmail();
    var fromName = readParam_(ss, 'MAIL_FROM') || undefined;

    MailApp.sendEmail({ to: to, subject: sb.subject, htmlBody: sb.bodyHtml, name: fromName });

    return { ok:true };
  } catch(e){ return { ok:false, error:String(e) }; }
}

function debugRetroFns() {
  if (!LIB) { Logger.log('LIB indisponible'); return; }
  Logger.log('typeof exportRetroMembresXlsxToDrive         = %s', typeof LIB.exportRetroMembresXlsxToDrive);
  Logger.log('typeof exportRetroGroupesAllXlsxToDrive      = %s', typeof LIB.exportRetroGroupesAllXlsxToDrive);
  Logger.log('typeof importerDonneesSaison                 = %s', typeof LIB.importerDonneesSaison);
  Logger.log('typeof evaluateSeasonRules                   = %s', typeof LIB.evaluateSeasonRules);
  Logger.log('typeof sendPendingOutbox                     = %s', typeof LIB.sendPendingOutbox);
}


/************* EXPORTS — INCR *************/
function runRetroExportsIncr(passports, options){
  options = options || {};
  var ss = getSeasonSpreadsheet_(getSeasonId_());
  var pass = (passports||[]).map(normalizePassportPlain8_).filter(Boolean);
  if (!pass.length) return { ok:true, note:'no-passports', files:[] };

  var allowIncr = _retroIncrementalOn_(); // déjà dans ton fichier
  if (!allowIncr) {
    // si incr OFF, bascule au full (mais filtrage côté lib sera neutralisé par nos wrappers)
    return { ok:true, note:'incr-off', files: [
      runExportRetroMembres(),
      runExportRetroGroupes()
    ]};
  }

  // 1) build rows complets
  var memb = SI.Library.buildRetroMembresRows(getSeasonId_());            // {header, rows, nbCols}
  var g    = SI.Library.buildRetroGroupesRows(getSeasonId_());
  var ga   = (typeof SI.Library.buildRetroGroupeArticlesRows==='function')
             ? SI.Library.buildRetroGroupeArticlesRows(getSeasonId_())
             : { header:g.header, rows:[], nbCols:g.nbCols };

  // 2) filtrer par passeports (y compris statut annulé, qui sera naturellement absent s’il n’est plus actif)
  var set = new Set(pass);
  var membRows = memb.rows.filter(function(r){ return set.has(normalizePassportPlain8_(r[0])); });
  var gRows    = g.rows.filter(function(r){ return set.has(normalizePassportPlain8_(r[0])); });
  var gaRows   = ga.rows.filter(function(r){ return set.has(normalizePassportPlain8_(r[0])); });

  // 3) COMBINED ou séparés
  var combined = !!options.combined;  // ex. from PARAMS ('RETRO_COMBINED_XLSX'==TRUE)
  if (combined) {
    var file = _exportCombinedXlsx_({
      Membres:{ header:memb.header, rows:membRows, nbCols:memb.nbCols, name:'Retro_Membres' },
      Groupes:{ header:g.header, rows:gRows, nbCols:g.nbCols, name:'Retro_Groupes' },
      GroupArt:{ header:ga.header, rows:gaRows, nbCols:ga.nbCols, name:'Retro_GroupeArticles' }
    });
    return { ok:true, files:[ file ] };
  } else {
    var f1 = _exportSingleSheetXlsx_('Retro_Membres', memb.header, membRows, memb.nbCols, 'Export temporaire - Retro Membres');
    var merged = _mergeRows_(g.header, gRows, ga.header, gaRows); // garde ton header de base
    var f2 = _exportSingleSheetXlsx_('Retro_Groupes_All', merged.header, merged.rows, merged.nbCols, 'Export temporaire - Import Retro Groupes All');
    return { ok:true, files:[ f1, f2 ] };
  }
}

// utilitaires
function _mergeRows_(h1, r1, h2, r2){
  // ici on concatène simplement; si les headers sont identiques c’est trivial
  if (JSON.stringify(h1) !== JSON.stringify(h2)) {
    // si besoin, fais une vraie union des colonnes (pour l’instant on assume identiques)
  }
  return { header: h1, rows: [].concat(r1||[], r2||[]), nbCols: (h1||[]).length };
}

function _exportCombinedXlsx_(tabs){
  var temp = SpreadsheetApp.create('Export temporaire - Retro Combined');
  var first = temp.getSheets()[0]; first.setName('Retro_Membres');
  // Membres
  _writeSheet_(first, tabs.Membres.header, tabs.Membres.rows, tabs.Membres.nbCols);
  // Groupes
  var shG = temp.insertSheet('Retro_Groupes');
  _writeSheet_(shG, tabs.Groupes.header, tabs.Groupes.rows, tabs.Groupes.nbCols);
  // GroupArticles
  var shGA = temp.insertSheet('Retro_GroupeArticles');
  _writeSheet_(shGA, tabs.GroupArt.header, tabs.GroupArt.rows, tabs.GroupArt.nbCols);

  SpreadsheetApp.flush();
  var blob = _exportSpreadsheetAsXlsx_(temp);
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HHmm');
  blob.setName('Export_Retro_Combined_' + ts + '_INCR.xlsx');

  var folderId = readParam_(SpreadsheetApp.openById(getSeasonId_()), PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID) || '';
  var dest = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = dest.createFile(blob);

  DriveApp.getFileById(temp.getId()).setTrashed(true);
  appendImportLog_(SpreadsheetApp.openById(getSeasonId_()), 'RETRO_COMBINED_XLSX_OK', file.getName() + ' -> ' + dest.getName());
  return { fileId:file.getId(), name:file.getName() };
}

function _exportSingleSheetXlsx_(sheetName, header, rows, nbCols, tempTitle){
  var temp = SpreadsheetApp.create(tempTitle || 'Export temporaire');
  var tmp = temp.getSheets()[0];
  tmp.setName('Export');
  _writeSheet_(tmp, header, rows, nbCols);
  SpreadsheetApp.flush();

  var blob = _exportSpreadsheetAsXlsx_(temp);
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HHmm');
  blob.setName(sheetName + '_' + ts + '_INCR.xlsx');

  var ss = getSeasonSpreadsheet_(getSeasonId_());
  var folderId = readParam_(ss, PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID) || '';
  var dest = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var f = dest.createFile(blob);
  DriveApp.getFileById(temp.getId()).setTrashed(true);

  appendImportLog_(ss, 'RETRO_XLSX_OK', f.getName() + ' -> ' + dest.getName() + ' (rows=' + rows.length + ')');
  return { fileId:f.getId(), name:f.getName(), rows: rows.length };
}

function _writeSheet_(sh, header, rows, nbCols){
  var all = [header].concat(rows || []);
  if (typeof normalizePassportToText8_ === 'function') {
    for (var i=1;i<all.length;i++){ if(all[i] && all[i].length) all[i][0] = normalizePassportToText8_(all[i][0]); }
  }
  if (all.length) {
    sh.getRange(1,1,all.length, nbCols).setValues(all);
    if (all.length > 1) sh.getRange(2,1,all.length-1,1).setNumberFormat('@');
  }
}
function _exportSpreadsheetAsXlsx_(ssTmp){
  var url = 'https://docs.google.com/spreadsheets/d/' + ssTmp.getId() + '/export?format=xlsx';
  return UrlFetchApp.fetch(url, { headers:{ Authorization:'Bearer '+ScriptApp.getOAuthToken() } }).getBlob();
}
// Perf cache commun (optionnel)
var __expPerfCache = null;
function _expPerfCacheBegin_(ss){
  __expPerfCache = {
    params: _readParamsAsMap_(ss),
    membres: _readMembresGlobalAsMap_(ss)
  };
}
function _expPerfCacheEnd_(){ __expPerfCache = null; }
