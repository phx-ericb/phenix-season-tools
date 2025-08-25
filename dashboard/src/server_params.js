function getParams(seasonId) {
  return _wrap('getParams', function(){
    var sh = getSeasonSpreadsheet_(seasonId).getSheetByName('PARAMS'); if (!sh) throw new Error('PARAMS sheet missing');
    var values = sh.getDataRange().getValues(); var out = {};
    for (var i=1; i<values.length; i++) { var k = values[i][0]; if (!k) continue; out[k] = _coerceByType_(k, values[i][1]); }
    return _ok(out);
  });
}
function getParamsDetailed(seasonId) {
  return _wrap('getParamsDetailed', function(){
    var sh = getSeasonSpreadsheet_(seasonId).getSheetByName('PARAMS'); if (!sh) throw new Error('PARAMS sheet missing');
    var values = sh.getDataRange().getValues(); if (!values.length) return _ok({ items: [] });
    var head = values[0].map(String);
    function idxOf(names){ for (var i=0;i<head.length;i++){ var h=String(head[i]).toLowerCase(); if (names.indexOf(h)>-1) return i; } return -1; }
    var ik=idxOf(['clé','cle','key']), iv=idxOf(['valeur','value']), it=idxOf(['type']), id=idxOf(['description']);
    var items=[], byKey={};
    for (var r=1;r<values.length;r++) { var row=values[r], k=row[ik]; if(!k)continue;
      var e={ key:String(k), value:row[iv], type:it>-1?String(row[it]||''):'', description:id>-1?String(row[id]||''):'' };
      e.value = _coerceByType_(e.key, e.value); items.push(e); byKey[e.key]=e.value;
    }
    return _ok({ items: items, byKey: byKey });
  });
}
function setParams(seasonId, params) {
  return _wrap('setParams', function(){
    var sh = getSeasonSpreadsheet_(seasonId).getSheetByName('PARAMS') || getSeasonSpreadsheet_(seasonId).insertSheet('PARAMS');
    if (sh.getLastRow() === 0) sh.appendRow(['Clé','Valeur','Type','Description']);
    var rng=sh.getDataRange().getValues(), mapRow={}; for (var i=1;i<rng.length;i++) if (rng[i][0]) mapRow[rng[i][0]] = i+1;
    if (!Array.isArray(params)) {
      Object.keys(params).forEach(function(k){ var v=_coerceByType_(k, params[k]); if (mapRow[k]) sh.getRange(mapRow[k],2).setValue(v); else sh.appendRow([k,v,PARAM_SCHEMA[k]||'','']); });
      return _ok(null,'Params saved');
    }
    params.forEach(function(e){ var k=e.key, v=_coerceByType_(k,e.value), t=e.type||(PARAM_SCHEMA[k]||''), d=e.description||''; if(mapRow[k]) sh.getRange(mapRow[k],1,1,4).setValues([[k,v,t,d]]); else sh.appendRow([k,v,t,d]); });
    return _ok(null,'Params saved');
  });
}
