/** ============================= server_import_logs.js ===============================
 * Journal d'import: lecture/écriture robustes dans la feuille IMPORT_LOG.
 * - appendImportLog_(ss, code, payload) → unique signature (PAS de surcharge)
 * - startImportRun_(opts) / endImportRun_(ctx) → ne font jamais planter le flow
 * - API_IMPORT_tail(limit) → lecture "snapshot" non-bloquante pour l'UI
 */

// ==== CONFIG ====
var IMPORT_LOG_SHEET = 'IMPORT_LOG';
var IMPORT_LOG_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

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

// ===================== LECTURE — résumé récent =====================

function getRecentActivity(mode, limit) {
  if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
    return { ok: false, busy: true, at: Date.now() };
  }

  return _wrap('getRecentActivity', function () {
    var sid = (typeof ER_resolveSeasonId_ === 'function') ? ER_resolveSeasonId_() : getSeasonId_();
    var ss = getSeasonSpreadsheet_(sid);
    var sh = getSheetOrCreate_(ss, IMPORT_LOG_SHEET, ['Date', 'Type', 'Détails']);
    var last = sh.getLastRow();
    if (last < 2) return [];

    var L = Math.min(limit || 50, last - 1);
    var vals = sh.getRange(last - L + 1, 1, L, Math.min(sh.getLastColumn(), 3)).getDisplayValues();
    var essentials = (String(mode || '').toLowerCase() !== 'all');

    var keep = essentials
      ? /^(RUN_IMPORT_|CONVERT_|DIFF_|EXPORT|RULES_|OUTBOX_|SEND_|MAIL_|QUEUE_|COACHS_)/
      : /.*/;

    return vals
      .reverse()
      .map(function (r) { return { date: r[0] || '', type: r[1] || '', details: r[2] || '' }; })
      .filter(function (x) { return keep.test(x.type || ''); });
  });
}

function getRecentActivityAll(limit) {
  return getImportLogSummary_({ mode: 'all', limit: limit || 50 });
}

// ===== Implémentation de lecture plus souple (multi headers) =====
function getImportLogSummary_(opts) {
  if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
    return { ok: false, busy: true, at: Date.now() };
  }

  var sid = (typeof ER_resolveSeasonId_ === 'function') ? ER_resolveSeasonId_() : getSeasonId_();
  var ss = getSeasonSpreadsheet_(sid);
  var sh = ss.getSheetByName(IMPORT_LOG_SHEET);
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  // Header
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  var idxDate = pickIdx_(header, ['Date', 'Datetime', 'Timestamp', 'Horodatage']);
  var idxType = pickIdx_(header, ['Type', 'Code', 'Event', 'Événement', 'Action']);
  var idxMsg  = pickIdx_(header, ['Details', 'Détails', 'Message', 'Msg', 'Info', 'Détail']);

  if (idxDate < 0) idxDate = 0;
  if (idxType < 0) idxType = Math.min(1, lastCol - 1);
  if (idxMsg  < 0) idxMsg  = Math.min(2, lastCol - 1);

  // Lire large puis filtrer
  var rawBudget = Math.min(lastRow - 1, (opts.limit || 50) * 5);
  var startRow = Math.max(2, lastRow - rawBudget + 1);
  var raw = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

  var tz = Session.getScriptTimeZone();
  var essentials = (String(opts.mode || 'essential').toLowerCase() === 'essential');

  var mapped = raw.map(function (row) {
    var d = row[idxDate];
    var ds = (d instanceof Date)
      ? Utilities.formatDate(d, tz, IMPORT_LOG_DATE_FORMAT)
      : String(d || '');
    var type = String(row[idxType] || '');
    var details = String(row[idxMsg] || '');
    return { date: ds, type: type, details: details };
  });

  if (essentials) {
    mapped = mapped.filter(function (x) { return !!ESSENTIAL_EVENTS[String(x.type || '').toUpperCase()]; });
  }

  // Récents en bas, limité
  var out = [];
  for (var i = mapped.length - 1; i >= 0 && out.length < (opts.limit || 50); i--) out.push(mapped[i]);
  out.reverse();
  return out;
}

function pickIdx_(header, candidates) {
  var map = header.reduce(function (acc, h, i) { acc[String(h).trim().toLowerCase()] = i; return acc; }, {});
  for (var i = 0; i < candidates.length; i++) {
    var k = String(candidates[i]).trim().toLowerCase();
    if (k in map) return map[k];
  }
  // includes (ex.: "Détails (FR)")
  for (var j = 0; j < header.length; j++) {
    var h = String(header[j]).trim().toLowerCase();
    for (var c = 0; c < candidates.length; c++) {
      if (h.indexOf(String(candidates[c]).toLowerCase()) >= 0) return j;
    }
  }
  return -1;
}

// ===================== TAIL — snapshot non-bloquant =====================

function API_IMPORT_tail(limit) {
  // si un import roule, évite d'ouvrir le classeur et signale busy
  if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
    return { ok: false, busy: true, at: Date.now() };
  }
  limit = Math.max(50, Math.min(Number(limit || 200), 500)); // borne
  var sid;
  try { sid = (typeof ER_resolveSeasonId_ === 'function') ? ER_resolveSeasonId_() : getSeasonId_(); }
  catch (e) { return { ok: false, reason: 'no-season' }; }

  var c = CacheService.getScriptCache();
  var K = 'IMPORT_LOG_TAIL_' + sid;
  var hit = c.get(K);
  if (hit) { try { return JSON.parse(hit); } catch (_) { c.remove(K); } }

  // Lock très court pour éviter la bataille avec un import en cours
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(400)) {
    return { ok: false, busy: true, at: Date.now() };
  }

  try {
    var ss = getSeasonSpreadsheet_(sid);
    var sh = ss.getSheetByName(IMPORT_LOG_SHEET);
    if (!sh) return { ok: true, rows: [], at: Date.now() };

    var last = sh.getLastRow();
    if (!last) return { ok: true, rows: [], at: Date.now() };

    var n = Math.max(1, last - limit + 1);
    var rng = sh.getRange(n, 1, last - n + 1, sh.getLastColumn()).getValues();
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

    var rows = rng.map(function (r) {
      var o = {};
      for (var i = 0; i < headers.length; i++) o[headers[i]] = r[i];
      return o;
    });

    var payload = { ok: true, rows: rows, lastRow: last, at: Date.now() };
    try { c.put(K, JSON.stringify(payload), 10); } catch (_) { }
    return payload;

  } catch (e) {
    return { ok: false, error: String(e), at: Date.now() };

  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

// ===================== ÉCRITURE — durcie =====================

// Append bas niveau avec lock court (retourne true si écrit, sinon false)
function _log_tryAppend_(ss, rows2d) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(1500)) return false; // ← un peu plus tolérant hors-import
  try {
    var sh = ss.getSheetByName(IMPORT_LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(IMPORT_LOG_SHEET);
      sh.getRange(1, 1, 1, 3).setValues([['Date', 'Type', 'Détails']]);
    }
    var last = sh.getLastRow();
    var rng = sh.getRange(last + 1, 1, rows2d.length, rows2d[0].length);
    rng.setValues(rows2d);
    return true;
  } finally {
    try { lock.releaseLock(); } catch (_) { }
  }
}

// Backlog en ScriptProperties si on ne peut pas écrire tout de suite
function _log_bufferPush_(rows2d) {
  var k = 'IMPORT_LOG_BACKLOG';
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty(k) || '[]';
  var arr;
  try { arr = JSON.parse(cur) || []; } catch (_) { arr = []; }
  if (arr.length > 2000) arr = arr.slice(-1500);
  arr.push(rows2d);
  props.setProperty(k, JSON.stringify(arr));
}

function API_IMPORT_flushBacklog() {
  // Si un import roule encore, ne tente pas de flusher (évite la contention)
  if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
    return { ok: false, busy: true, flushed: 0 };
  }

  var sid = (typeof ER_resolveSeasonId_ === 'function') ? ER_resolveSeasonId_() : getSeasonId_();
  var ss = getSeasonSpreadsheet_(sid);

  var k = 'IMPORT_LOG_BACKLOG';
  var props = PropertiesService.getScriptProperties();
  var cur = props.getProperty(k) || '[]';
  var arr;
  try { arr = JSON.parse(cur) || []; } catch (_) { arr = []; }
  if (!arr.length) return { ok: true, flushed: 0 };

  var flushed = 0;
  for (var i = 0; i < arr.length; i++) {
    var ok = _log_tryAppend_(ss, arr[i]);
    if (!ok) break;
    flushed++;
  }
  if (flushed) {
    props.setProperty(k, JSON.stringify(arr.slice(flushed)));
  }
  return { ok: true, flushed: flushed };
}

/**
 * Append public robuste:
 *  - signature unique: (ss, code, payload)
 *  - non-bloquant pendant l'import (backlog direct si PHENIX_IMPORT_LOCK)
 *  - 5 essais avec backoff hors-import, sinon buffer
 */
function appendImportLog_(ss, code, payload) {
  // Pendant un import: n'écrit pas dans la feuille, empile en backlog (non-bloquant)
  try {
    if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
      var nowLock = new Date();
      var rowLock = [
        Utilities.formatDate(nowLock, Session.getScriptTimeZone(), IMPORT_LOG_DATE_FORMAT),
        String(code || ''),
        typeof payload === 'string' ? payload : JSON.stringify(payload || null)
      ];
      _log_bufferPush_([rowLock]);
      return true;
    }
  } catch (_) {}

  var now = new Date();
  var row = [
    Utilities.formatDate(now, Session.getScriptTimeZone(), IMPORT_LOG_DATE_FORMAT),
    String(code || ''),
    typeof payload === 'string' ? payload : JSON.stringify(payload || null)
  ];
  var rows2d = [row];

  // Essais hors-import
  var base = 250;
  for (var a = 1; a <= 5; a++) {
    try {
      if (ss && _log_tryAppend_(ss, rows2d)) return true;
    } catch (e) {
      // no-op, on retente
    }
    Utilities.sleep(base * Math.pow(2, a - 1)); // 250,500,1000,2000,4000
  }

  // Buffer en dernier recours
  _log_bufferPush_(rows2d);
  return false;
}

// ===================== Cycle de run (start/end) =====================

function startImportRun_(opts) {
  var sid = (opts && opts.seasonId) ||
            (typeof ER_resolveSeasonId_ === 'function' ? ER_resolveSeasonId_() : getSeasonId_());
  var ss = getSeasonSpreadsheet_(sid);
  var ctx = {
    runId: Utilities.getUuid(),
    seasonId: sid,
    ss: ss,
    source: (opts && opts.source) || 'manual',
    startedAt: new Date()
  };

  // Persistance légère (facultatif)
  _setRunCtxProps_({
    PHENIX_IMPORT_RUN_ID: ctx.runId,
    PHENIX_IMPORT_RUN_ACTIVE: '1',
    PHENIX_IMPORT_RUN_STARTED_AT: String(ctx.startedAt.toISOString())
  });

  appendImportLog_(ss, 'RUN_IMPORT_START', { runId: ctx.runId, source: ctx.source });
  return ctx;
}

function endImportRun_(ctx) {
  try {
    if (!ctx) return true;
    appendImportLog_(ctx.ss, 'RUN_IMPORT_END', {
      runId: ctx.runId,
      durMs: Date.now() - ctx.startedAt.getTime()
    });
    // flush opportuniste du backlog (ne bloque pas si un import est signalé encore actif)
    try { API_IMPORT_flushBacklog(); } catch (_) { }
  } catch (_) {
    // ne fait pas planter le flow
  } finally {
    _setRunCtxProps_({ PHENIX_IMPORT_RUN_ACTIVE: '0' });
  }
  return true;
}

// ===================== Utils internes =====================

function _getRunCtxProps_() {
  var props = PropertiesService.getScriptProperties();
  return {
    runId: props.getProperty('PHENIX_IMPORT_RUN_ID') || '',
    active: props.getProperty('PHENIX_IMPORT_RUN_ACTIVE') || '0',
    startedAt: props.getProperty('PHENIX_IMPORT_RUN_STARTED_AT') || ''
  };
}

function _setRunCtxProps_(kv) {
  var props = PropertiesService.getScriptProperties();
  Object.keys(kv || {}).forEach(function (k) { props.setProperty(k, String(kv[k])); });
}

/** Mini résumé import (inchangé, utilitaire) */
function summarizeImportResult_(impRes) {
  var o = {};
  try {
    if (!impRes) return { note: 'no-result' };
    if (impRes.incrInscriptions) {
      var i = impRes.incrInscriptions;
      o.inscriptions = {
        added: Number(i.added || 0),
        updated: Number(i.updated || 0),
        invalidated: Number(i.annuls || 0),
        outbox: Number(i.outbox || 0),
        mods: Number(i.mods || 0),
        uniquePassports: Array.isArray(i.touchedPassports) ? new Set(i.touchedPassports).size : 0,
        recordsTouched: Array.isArray(i.touchedPassports) ? i.touchedPassports.length : 0
      };
    }
    if (impRes.incrArticles) {
      var a = impRes.incrArticles;
      o.articles = {
        added: Number(a.added || 0),
        updated: Number(a.updated || 0),
        invalidated: Number(a.annuls || 0),
        uniquePassports: Array.isArray(a.touchedPassports) ? new Set(a.touchedPassports).size : 0,
        recordsTouched: Array.isArray(a.touchedPassports) ? a.touchedPassports.length : 0
      };
    }
    if (impRes.rules) {
      o.rules = {
        written: Number(impRes.rules.written || 0),
        inscPlay: Number(impRes.rules.inscScanned || impRes.rules.inscPlay || 0),
        artPlay: Number(impRes.rules.artScanned || impRes.rules.artPlay || 0),
        source: impRes.rules.source || null
      };
    }
  } catch (e) {
    o.error = String(e);
  }
  return o;
}

/** Debug helper */
function debug_tailImportLog(n) {
  n = n || 30;
  var sid = (typeof ER_resolveSeasonId_ === 'function') ? ER_resolveSeasonId_() : getSeasonId_();
  var ss = getSeasonSpreadsheet_(sid);
  var sh = ss.getSheetByName(IMPORT_LOG_SHEET);
  if (!sh || sh.getLastRow() < 2) { Logger.log('IMPORT_LOG vide.'); return; }
  var last = sh.getLastRow();
  var start = Math.max(2, last - n + 1);
  var vals = sh.getRange(start, 1, last - start + 1, Math.min(sh.getLastColumn(), 3)).getValues();
  vals.forEach(function (r) { Logger.log('%s | %s | %s', r[0], r[1], r[2]); });
}

// ===================== Helper de wrap (optionnel) =====================
// Si tu as déjà un _wrap global, tu peux retirer ceci.
function _wrap(name, fn) {
  try { return fn(); }
  catch (e) { return { ok: false, error: '[' + name + '] ' + String(e) }; }
}
