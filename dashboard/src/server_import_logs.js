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
