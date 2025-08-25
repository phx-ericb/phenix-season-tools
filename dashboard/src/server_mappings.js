function listMappingSheets(seasonId) {
  return _wrap('listMappingSheets', function(){
    var ss = getSeasonSpreadsheet_(seasonId);
    var names = ss.getSheets().map(function(s){ return s.getName(); })
      .filter(function(n){ return /^MAPPINGS/.test(n) || /^GROUPES/.test(n); });
    return _ok(names);
  });
}
function getMappings(seasonId, sheetName) {
  return _wrap('getMappings', function(){
    var sh = getSeasonSpreadsheet_(seasonId).getSheetByName(sheetName); if(!sh) throw new Error('Sheet not found: '+sheetName);
    var values = sh.getDataRange().getValues(); if (!values.length) return _ok({ sheet:sheetName, headers:[], rows:[] });
    var headerIdx=-1; for (var r=0;r<values.length;r++){ var nonEmpty=values[r].filter(function(x){return String(x||'').trim()!=='';}).length; if(nonEmpty>=2){ headerIdx=r; break; } }
    if (headerIdx===-1) return _ok({ sheet:sheetName, headers:[], rows:[] });
    var headers=values[headerIdx].map(String), rows=values.slice(headerIdx+1);
    return _ok({ sheet:sheetName, headers:headers, rows:rows });
  });
}
function setMappings(seasonId, sheetName, rows) {
  return _wrap('setMappings', function(){
    var sh = getSeasonSpreadsheet_(seasonId).getSheetByName(sheetName) || getSeasonSpreadsheet_(seasonId).insertSheet(sheetName);
    sh.clear(); if (!rows || !rows.length) return _ok(null,'Empty mapping');
    sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);
    return _ok(null,'Mappings saved');
  });
}
function importMappingsCsv(seasonId, sheetName, csv) { return _wrap('importMappingsCsv', function(){ var arr = Utilities.parseCsv(csv); return setMappings(seasonId, sheetName, arr); }); }
function exportMappingsCsv(seasonId, sheetName) {
  return _wrap('exportMappingsCsv', function(){
    var data = getMappings(seasonId, sheetName); if (!data.ok) return data;
    var rows = [data.data.headers].concat(data.data.rows||[]);
    function esc(v){ v=(v===null||v===undefined)?'':String(v); if (/[",\n]/.test(v)) v='"'+v.replace(/"/g,'""')+'"'; return v; }
    var csv = rows.map(function(r){ return r.map(esc).join(','); }).join('\n');
    return _ok({ filename: sheetName + '.csv', mimeType: 'text/csv', content: csv });
  });
}
