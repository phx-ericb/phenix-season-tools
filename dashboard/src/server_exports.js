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
function runExport(seasonId, type, options) {
  return _wrap('runExport', function(){
    options = options || {};
    if (type==='members')       { if(!SI.Library.exportRetroMembresXlsx) throw new Error('exportRetroMembresXlsx not available in lib v0.7'); return _ok(SI.Library.exportRetroMembresXlsx(seasonId, options), 'Export Membres lancé'); }
    if (type==='groupes')       { if(!SI.Library.exportRetroGroupesXlsx) throw new Error('exportRetroGroupesXlsx not available in lib v0.7'); return _ok(SI.Library.exportRetroGroupesXlsx(seasonId, options), 'Export Groupes lancé'); }
    if (type==='groupArticles') { if(!SI.Library.exportRetroGroupeArticlesXlsx) throw new Error('exportRetroGroupeArticlesXlsx not available in lib v0.7'); return _ok(SI.Library.exportRetroGroupeArticlesXlsx(seasonId, options), 'Export Groupe Articles lancé'); }
    throw new Error('Unknown export type: ' + type);
  });
}
function previewExport(seasonId, type, maxRows) {
  return _wrap('previewExport', function(){
    maxRows = maxRows || 20;
    var out;
    if (type==='members')       out = SI.Library.buildRetroMembresRows(seasonId);
    else if (type==='groupes')  out = SI.Library.buildRetroGroupesRows(seasonId);
    else if (type==='groupArticles') out = SI.Library.buildRetroGroupeArticlesRows(seasonId);
    else throw new Error('Unknown export type: ' + type);

    var header = out.header || [];
    var rows   = out.rows || [];
    return _ok({
      header: header,
      total: rows.length,
      preview: rows.slice(0, maxRows),
      nbCols: out.nbCols || header.length
    }, 'Preview ready');
  });
}
