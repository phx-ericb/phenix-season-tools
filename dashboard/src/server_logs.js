function getLogs(seasonId, sheet, limit) {
  return _wrap('getLogs', function(){
    sheet = sheet || 'IMPORT_LOG'; limit = limit || 100;
    var ss = getSeasonSpreadsheet_(seasonId);
    var sh = ss.getSheetByName(sheet); if (!sh) throw new Error('Log sheet not found: ' + sheet);
    var last = sh.getLastRow(); if (last === 0) return _ok([]);
    var start = Math.max(2, last - limit + 1);
    var rng = sh.getRange(start, 1, last - start + 1, sh.getLastColumn()).getValues();
    var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
    var idxTime = headers.indexOf('When') > -1 ? headers.indexOf('When') : 0;
    var idxType = headers.indexOf('Type') > -1 ? headers.indexOf('Type') : 1;
    var idxMsg  = headers.indexOf('Message') > -1 ? headers.indexOf('Message') : 2;
    var idxPayload = headers.indexOf('Payload') > -1 ? headers.indexOf('Payload') : -1;
    var out = rng.map(function(r){ return { when: r[idxTime] instanceof Date ? r[idxTime].toISOString() : String(r[idxTime]||''), type: String(r[idxType]||''), message: String(r[idxMsg]||''), payload: (idxPayload>-1 ? r[idxPayload] : null) }; });
    return _ok(out);
  });
}
