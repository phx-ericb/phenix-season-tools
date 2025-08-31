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

// ---- Courriels D'ERREUR (réutilise le moteur de secteurs) ----
// errorItem = { code: 'MISSING_CDP', label: 'CDP manquant', details: '...' }

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

    // payload utilisateur + enrichissements erreur
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

    MailApp.sendEmail({
      to: to,
      subject: sb.subject,
      htmlBody: sb.bodyHtml,
      name: fromName
    });

    return { ok:true };
  } catch(e){ return { ok:false, error:String(e) }; }
}
