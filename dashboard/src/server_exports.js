/* ============================== server_exports.js — v1.1 ==============================
 * - Met à jour les endpoints d’export pour utiliser les nouveaux noms «*ToDrive»
 * - Ajoute le support du **filtrage incrémental** via options.onlyPassports
 * - Fournit des helpers pour lire/forcer les «touchedPassports» depuis DocumentProperties
 * - Conserve la compatibilité avec l’UI existante (preview, testRules, etc.)
 *
 * NOTE: suppose la présence de SI.Library (alias LIB) côté librairie.
 */


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
 * Lance un export complet (ou incrémental si options.onlyPassports est fourni)
 * type ∈ { 'members', 'groupes', 'groupArticles' }
 */
function runExport(seasonId, type, options) {
  return _wrap('runExport', function(){
    options = options || {};
    if (type==='members'){
      if(!SI.Library.exportRetroMembresXlsxToDrive) throw new Error('exportRetroMembresXlsxToDrive not available');
      return _ok(SI.Library.exportRetroMembresXlsxToDrive(seasonId, options), 'Export Membres lancé');
    }
    if (type==='groupes'){
      if(!SI.Library.exportRetroGroupesAllXlsxToDrive) throw new Error('exportRetroGroupesAllXlsxToDrive not available');
      return _ok(SI.Library.exportRetroGroupesAllXlsxToDrive(seasonId, options), 'Export Groupes (ALL) lancé');
    }
    if (type==='groupArticles'){
      if(!SI.Library.exportRetroGroupeArticlesXlsxToDrive) throw new Error('exportRetroGroupeArticlesXlsxToDrive not available');
      return _ok(SI.Library.exportRetroGroupeArticlesXlsxToDrive(seasonId, options), 'Export Groupe Articles lancé');
    }
    throw new Error('Unknown export type: ' + type);
  });
}

/** Version pratique: force l’incrémental en lisant LAST_TOUCHED_PASSPORTS */
function runExportIncr(seasonId, type){
  var list = _getTouchedPassportsArray_();
  return runExport(seasonId, type, { onlyPassports: list });
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

// ---- Courriels D'ERREUR (réutilise le moteur de secteurs) ----
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