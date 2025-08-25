function getRules(seasonId) {
  return _wrap('getRules', function(){
    var ss = getSeasonSpreadsheet_(seasonId);
    var sh = ss.getSheetByName('RETRO_RULES_JSON');
    var text = sh && sh.getLastRow()>0 ? String(sh.getRange(1,1).getValue()||'') : '';
    if (!text) {
      var p=ss.getSheetByName('PARAMS'); if (p) { var values=p.getDataRange().getValues();
        for (var i=1;i<values.length;i++) if (values[i][0]==='RETRO_RULES_JSON') { text=String(values[i][1]||''); break; }
      }
    }
    var parsed=null, ok=true, err=''; if (text && String(text).trim()) { try { parsed=JSON.parse(text);} catch(e){ ok=false; err=String(e);} }
    return _ok({ jsonText:text, parsed:parsed, parsedOk:ok, error:err });
  });
}
function setRules(seasonId, jsonText) {
  return _wrap('setRules', function(){
    try { if (jsonText && jsonText.trim()) JSON.parse(jsonText); } catch(e) { throw new Error('Invalid JSON: ' + e); }
    var sh = getSeasonSpreadsheet_(seasonId).getSheetByName('RETRO_RULES_JSON') || getSeasonSpreadsheet_(seasonId).insertSheet('RETRO_RULES_JSON');
    sh.clear(); sh.getRange(1,1).setValue(jsonText);
    return _ok(null,'Rules saved');
  });
}
