/** @OnlyCurrentDoc
 * Lecture du journal d'import depuis la feuille IMPORT_LOG,
 * avec un filtrage "essentiel" par défaut.
 */

// ==== CONFIG ====
var IMPORT_LOG_SHEET = 'IMPORT_LOG';
var IMPORT_LOG_DATE_FORMAT = 'yyyy-MM-dd HH:mm:ss';

// Types d’événements "essentiels" (ajuste à ta guise)
var ESSENTIAL_EVENTS = {
  RUN_IMPORT_START: 1,
  RUN_IMPORT_END: 1,
  CONVERT_ID: 1,
  CONVERT_OK: 1,
  DIFF_OK: 1,
  EXPORT_OK: 1,
  EXPORTS_DONE: 1,
  OUTBOX_ENQUEUED: 1,
  SEND_OUTBOX_DONE: 1,
  RULES_OK: 1,
  RULES_FAIL_WRAP: 1,
};

function getRecentActivity(mode, limit){
  return _wrap('getRecentActivity', function(){
    var ss = getSeasonSpreadsheet_(getSeasonId_());
    var sh = getSheetOrCreate_(ss, 'IMPORT_LOG', ['Horodatage','Action','Détails']);
    var last = sh.getLastRow(); if (last<2) return [];
    var L = Math.min(limit||50, last-1);
    var vals = sh.getRange(last-L+1,1,L,3).getDisplayValues();
    var essential = (String(mode||'').toLowerCase()!=='all');
    var keep = essential
      ? /^(IMPORT|RULES_|QUEUE_|MAIL_|COACHS_)/ // ajuste au besoin
      : /.*/;
    return vals
      .reverse()
      .map(function(r){ return { date:r[0]||'', type:r[1]||'', details:r[2]||'' }; })
      .filter(function(x){ return keep.test(x.type||''); });
  });
}


// Version “all” si tu veux l’exposer aussi
function getRecentActivityAll(limit) {
  return getImportLogSummary_({ mode: 'all', limit: limit || 50 });
}

// ===== Implémentation =====
function getImportLogSummary_(opts) {
  var sh = SpreadsheetApp.getActive().getSheetByName(IMPORT_LOG_SHEET);
  if (!sh) return []; // pas de log, pas de crise :)

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // Header
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  var idxDate = pickIdx_(header, ['Date', 'Datetime', 'Timestamp', 'Horodatage']);
  var idxType = pickIdx_(header, ['Type', 'Code', 'Event', 'Événement']);
  var idxMsg  = pickIdx_(header, ['Details', 'Détails', 'Message', 'Msg', 'Info', 'Détail']);

  // fallback simple si on ne trouve pas
  if (idxDate < 0) idxDate = 0;
  if (idxType < 0) idxType = Math.min(1, lastCol - 1);
  if (idxMsg  < 0) idxMsg  = Math.min(2, lastCol - 1);

  // Lis les X dernières lignes brutes (on lit plus large pour pouvoir filtrer)
  var rawBudget = Math.min(lastRow - 1, (opts.limit || 50) * 5);
  var startRow = Math.max(2, lastRow - rawBudget + 1);
  var raw = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  var tz = Session.getScriptTimeZone();
  var essentials = (String(opts.mode || 'essential').toLowerCase() === 'essential');

  // Map -> objets {date,type,details}
  var mapped = raw.map(function(row) {
    var d = row[idxDate];
    var ds = (d instanceof Date)
      ? Utilities.formatDate(d, tz, IMPORT_LOG_DATE_FORMAT)
      : String(d || '');
    var type = String(row[idxType] || '');
    var details = String(row[idxMsg] || '');

    return { date: ds, type: type, details: details };
  });

  // Filtrage “essentiel”
  if (essentials) {
    mapped = mapped.filter(function(x) {
      var t = String(x.type || '').toUpperCase();
      return !!ESSENTIAL_EVENTS[t];
    });
  }

  // On veut les plus récents en bas (chronologique) et limiter
  // On part de la fin (plus récent), on garde 'limit', puis on remet à l'endroit
  var out = [];
  for (var i = mapped.length - 1; i >= 0 && out.length < (opts.limit || 50); i--) {
    out.push(mapped[i]);
  }
  out.reverse();
  return out;
}

function pickIdx_(header, candidates) {
  var map = header.reduce(function(acc, h, i) { acc[String(h).trim().toLowerCase()] = i; return acc; }, {});
  for (var i = 0; i < candidates.length; i++) {
    var k = String(candidates[i]).trim().toLowerCase();
    if (k in map) return map[k];
  }
  // essaie par "includes" (ex.: "Détails (FR)")
  for (var j = 0; j < header.length; j++) {
    var h = String(header[j]).trim().toLowerCase();
    for (var c = 0; c < candidates.length; c++) {
      if (h.indexOf(String(candidates[c]).toLowerCase()) >= 0) return j;
    }
  }
  return -1;
}
function startImportRun_(meta){
  var runId = Utilities.getUuid();
  _setRunCtxProps_({
    PHENIX_IMPORT_RUN_ID: runId,
    PHENIX_IMPORT_RUN_ACTIVE: '1',
    PHENIX_IMPORT_RUN_STARTED_AT: String(new Date().toISOString())
  });
  appendImportLog_({ type:'RUN_IMPORT_START', details: { runId: runId, details:'Déclenchement via '+(meta && meta.source || 'unknown'), seasonId:(meta && meta.seasonId)||null } });
  return { runId: runId };
}
function endImportRun_(ctx){
  var t = String(new Date().toISOString());
  appendImportLog_({ type:'RUN_IMPORT_END', details: { runId: ctx && ctx.runId || null, at: t, message:'Terminé' } });
  _setRunCtxProps_({ PHENIX_IMPORT_RUN_ACTIVE: '0' });
}
function _getRunCtxProps_(){
  var props = PropertiesService.getScriptProperties();
  return {
    runId: props.getProperty('PHENIX_IMPORT_RUN_ID') || '',
    active: props.getProperty('PHENIX_IMPORT_RUN_ACTIVE') || '0',
    startedAt: props.getProperty('PHENIX_IMPORT_RUN_STARTED_AT') || ''
  };
}

function _setRunCtxProps_(kv){
  var props = PropertiesService.getScriptProperties();
  Object.keys(kv||{}).forEach(function(k){ props.setProperty(k, String(kv[k])); });
}
/** Écrit une ligne dans IMPORT_LOG (Date | Type | Détails). Accepte 1 ou 3 arguments.
 *  Améliorations:
 *   - Ajoute runId et lateFlag (si le run est clos)
 *   - Sanitize les détails volumineux (touchedPassports -> count + sample) sauf si VERBOSE_TOUCHED_PASSPORTS=1
 *   - Supporte objet details (sera JSON.stringify compact)
 */
function appendImportLog_(a, b, c){
  var ss = getSeasonSpreadsheet_(getSeasonId_());
  var name = 'IMPORT_LOG';
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,3).setValues([['Date','Type','Détails']]);
  }

  // 1) Résolution type / details
  var type, details;
  if (arguments.length === 1 && a && typeof a === 'object') {
    type = a.type || 'INFO';
    details = a.details || '';
  } else if (arguments.length >= 3) {
    type = b || 'INFO';
    details = c || '';
  } else {
    type = 'INFO';
    details = String(a == null ? '' : a);
  }

  // 2) Contexte de run & late flag
  var ctx = _getRunCtxProps_(); // {runId, active, startedAt}
  var late = (ctx && ctx.active === '0'); // si déjà end

  // 3) Sanitize & enrich details
  var enriched = _enrichAndSanitizeDetails_(details, {
    runId: ctx && ctx.runId || null,
    late: late
  });

  // 4) Append
  sh.appendRow([ new Date(), late ? ('LATE_' + type) : type, enriched ]);
}
/** ---------- Details Sanitizer ---------- */
function _enrichAndSanitizeDetails_(details, extra){
  // Convertir en objet si possible
  var obj, rawIsObject = false;
  if (details && typeof details === 'object') {
    obj = details;
    rawIsObject = true;
  } else if (typeof details === 'string') {
    try { obj = JSON.parse(details); } catch(e){ obj = { message: details }; }
  } else {
    obj = { message: String(details) };
  }

  // Inject runId/late si fournis
  if (extra && extra.runId && !obj.runId) obj.runId = extra.runId;
  if (extra && typeof extra.late === 'boolean') obj.late = extra.late;

  // Réduction des payloads trop gros (touchedPassports)
  try {
    var verbose = String(readParamValue('VERBOSE_TOUCHED_PASSPORTS') || '').trim() === '1';
    if (!verbose && obj && obj.touchedPassports && Array.isArray(obj.touchedPassports)) {
      var arr = obj.touchedPassports;
      obj.touchedPassportsCount = arr.length;
      obj.touchedPassportsSample = arr.slice(0, 5);
      delete obj.touchedPassports;
    }
  } catch(e){ /* no-op */ }

  // Failsafe: si le JSON devient trop gros, tronquer le champ message
  var str = JSON.stringify(obj);
  if (str && str.length > 60000) { // évite dépassement GAS
    obj._truncated = true;
    if (obj.message && typeof obj.message === 'string' && obj.message.length > 2000) {
      obj.message = obj.message.slice(0, 2000) + '…[truncated]';
    }
    str = JSON.stringify(obj);
  }

  return str;
}
function debug_tailImportLog(n) {
  n = n || 30;
  var ss = getSeasonSpreadsheet_(getSeasonId_());
  var sh = ss.getSheetByName('IMPORT_LOG');
  if (!sh || sh.getLastRow() < 2) { Logger.log('IMPORT_LOG vide.'); return; }
  var last = sh.getLastRow();
  var start = Math.max(2, last - n + 1);
  var vals = sh.getRange(start, 1, last - start + 1, Math.min(sh.getLastColumn(), 3)).getValues();
  vals.forEach(function(r){ Logger.log('%s | %s | %s', r[0], r[1], r[2]); });
}
/** ---------- Mini résumé import pour lisibilité ---------- */
function summarizeImportResult_(impRes){
  // On essaye d’extraire des compteurs usuels si présents
  var o = {};
  try {
    if (!impRes) return { note:'no-result' };
    // Inscriptions
    if (impRes.incrInscriptions) {
      var i = impRes.incrInscriptions;
      o.inscriptions = {
        added: Number(i.added||0),
        updated: Number(i.updated||0),
        invalidated: Number(i.annuls||0), // renommage logique
        outbox: Number(i.outbox||0),
        mods: Number(i.mods||0),
        uniquePassports: Array.isArray(i.touchedPassports) ? new Set(i.touchedPassports).size : 0,
        recordsTouched: Array.isArray(i.touchedPassports) ? i.touchedPassports.length : 0
      };
    }
    // Articles
    if (impRes.incrArticles) {
      var a = impRes.incrArticles;
      o.articles = {
        added: Number(a.added||0),
        updated: Number(a.updated||0),
        invalidated: Number(a.annuls||0),
        uniquePassports: Array.isArray(a.touchedPassports) ? new Set(a.touchedPassports).size : 0,
        recordsTouched: Array.isArray(a.touchedPassports) ? a.touchedPassports.length : 0
      };
    }
    // Règles
    if (impRes.rules) {
      o.rules = {
        written: Number(impRes.rules.written||0),
        inscPlay: Number(impRes.rules.inscScanned||impRes.rules.inscPlay||0),
        artPlay: Number(impRes.rules.artScanned||impRes.rules.artPlay||0),
        source: impRes.rules.source || null
      };
    }
  } catch(e){
    o.error = String(e);
  }
  return o;
}