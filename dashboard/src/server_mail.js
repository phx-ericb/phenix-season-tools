// server_mail.js
/**
 * server_mail.js — endpoints Dashboard (UI “Courriels”)
 * - Gère les secteurs d’envoi (CRUD) dans la feuille MAIL_SECTEURS
 * - Aperçu & envoi test d’un secteur sur un passeport
 *
 * Dépendances (si déjà fournies ailleurs, on ne les redéfinit pas) :
 *   getSeasonSpreadsheet_, readParam_, readSheetAsObjects_, renderTemplate_,
 *   deriveSectorFromRow_, deriveUFromRow_, collectEmailsFromRow_,
 *   buildDataFromRow_
 * Sinon, on active des fallbacks compacts & sûrs.
 *
 * UI attendu (ui_js_mail.html) appelle :
 *   getMailSectors(seasonId)
 *   upsertMailSector(seasonId, item)
 *   deleteMailSector(seasonId, secteurId)
 *   duplicateMailSector(seasonId, secteurId)
 *   previewSectorForPassport(seasonId, secteurId, passport, itemOverride)
 *   sendSectorTest(seasonId, item, passport, toTest)
 */

/* -------------------- Fallbacks légers (si absents) -------------------- */
if (typeof getSeasonSpreadsheet_ !== 'function') {
  function getSeasonSpreadsheet_(id) { if (!id) throw new Error('seasonId requis'); return SpreadsheetApp.openById(id); }
}
if (typeof readParam_ !== 'function') {
  function readParam_(ss, key){
    var sh = ss.getSheetByName('PARAMS');
    if (sh && sh.getLastRow() >= 1) {
      var rng = sh.getRange(1,1,sh.getLastRow(),2).getValues();
      for (var i=0;i<rng.length;i++){ if (String(rng[i][0]||'').trim() === key) return String(rng[i][1]||'').trim(); }
    }
    var props = PropertiesService.getDocumentProperties();
    return String(props.getProperty(key)||'').trim();
  }
}
if (typeof readSheetAsObjects_ !== 'function') {
  function readSheetAsObjects_(ssId, sheetName){
    var ss = (typeof ssId==='string') ? SpreadsheetApp.openById(ssId) : ssId;
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
      return { sheet: sh || ss.insertSheet(sheetName), headers: [], rows: [] };
    }
    var values = sh.getRange(1,1,sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    var headers = values[0].map(String);
    var out = [];
    for (var r=1;r<values.length;r++){
      var row = {}; for (var c=0;c<headers.length;c++) row[headers[c]] = values[r][c];
      out.push(row);
    }
    return { sheet: sh, headers: headers, rows: out };
  }
}
if (typeof renderTemplate_ !== 'function') {
  function renderTemplate_(tpl, data){
    tpl = String(tpl==null?'':tpl);
    return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function(_m, k){ return (data!=null && data.hasOwnProperty(k)) ? String(data[k]||'') : ''; });
  }
}
if (typeof collectEmailsFromRow_ !== 'function') {
  function _normEmail_(s){ s=String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s.trim(); }
  function collectEmailsFromRow_(row, fieldsCsv){
    var fields = (fieldsCsv?fieldsCsv.split(','):['Courriel','Parent 1 - Courriel','Parent 2 - Courriel']).map(function(x){return String(x||'').trim();}).filter(Boolean);
    var bag = {};
    fields.forEach(function(f){
      var v = row[f]; if (!v) return;
      String(v).split(/[;,]/).forEach(function(e){ e=_normEmail_(e); if (e) bag[e]=true; });
    });
    return Object.keys(bag).join(',');
  }
}
if (typeof deriveUFromRow_ !== 'function') {
  function _parseSeasonYear_(s){ var m=String(s||'').match(/(20\d{2})/); return m?parseInt(m[1],10):(new Date()).getFullYear(); }
  function _birthYearFromRow_(row){
    var y=row['Année de naissance']||row['Annee de naissance']||row['Annee']||'';
    if (y && /^\d{4}$/.test(String(y))) return parseInt(y,10);
    var dob=row['Date de naissance']||'';
    if (dob){
      var s=String(dob), m=s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if(m) return parseInt(m[1],10);
      var m2=s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/); if(m2) return parseInt(m2[3],10);
    }
    return null;
  }
  function _computeUForYear_(by, sy){ if(!by||!sy) return ''; var u=sy-by; return (u>=4 && u<=21)?('U'+u):''; }
  function deriveUFromRow_(row){
    var cat=row['Catégorie']||row['Categorie']||'';
    if (cat && /^U\d{1,2}/i.test(cat)) return cat.toUpperCase().replace(/\s+/g,'');
    var sy = _parseSeasonYear_(row['Saison']||''); var by=_birthYearFromRow_(row);
    return _computeUForYear_(by, sy);
  }
}
if (typeof deriveSectorFromRow_ !== 'function') {
  function deriveSectorFromRow_(row){
    var U = deriveUFromRow_(row)||''; var n = parseInt(String(U).replace(/^U/i,''),10);
    if (!n || isNaN(n)) return 'AUTRES';
    if (n>=4 && n<=8)  return 'U4-U8';
    if (n>=9 && n<=12) return 'U9-U12';
    if (n>=13 && n<=18) return 'U13-U18';
    return 'AUTRES';
  }
}
if (typeof buildDataFromRow_ !== 'function') {
  function _genreInitFromRow_(row){
    var g=String(row['Identité de genre']||row['Identité de Genre']||row['Genre']||'').toUpperCase();
    if (!g) return ''; if (g[0]==='M') return 'M'; if (g[0]==='F') return 'F'; return 'X';
  }
function buildDataFromRow_(row){
  var U = deriveUFromRow_(row)||''; 
  var n = parseInt(String(U).replace(/^U/i,''),10);
  var U2 = !isNaN(n) ? ('U'+(n<10?('0'+n):n)) : '';

  var prenomRaw = row['Prénom']||row['Prenom']||'';
  var nomRaw    = row['Nom']||'';

  var prenomPC  = _toProperCase_(prenomRaw);
  var nomPC     = _toProperCase_(nomRaw);

  return {
    passeport: row['Passeport #']||'',
    nom: nomPC,
    prenom: prenomPC,
    nomcomplet: (prenomPC + ' ' + nomPC).trim(),
    saison: row['Saison']||'',
    frais: row['Nom du frais']||row['Frais']||row['Produit']||'',
    categorie: row['Catégorie']||row['Categorie']||'',
    secteur: deriveSectorFromRow_(row)||'',
    U: U||'', U2: U2||'', U_num: n||'',
    genre: row['Identité de genre']||row['Identité de Genre']||row['Genre']||'',
    genreInitiale: _genreInitFromRow_(row)
  };
}
}

function _toProperCase_(s) {
  s = String(s||'').toLowerCase();
  return s.replace(/\b\w/g, function(c){ return c.toUpperCase(); });
}


/* Helpers locaux MAIL_SECTEURS (header-safe) */
/* -------------------- Constantes feuille secteurs -------------------- */
var MAIL_SECTORS_SHEET  = 'MAIL_SECTEURS';
var MAIL_SECTORS_HEADER = ['SecteurId','Label','Umin','Umax','Genre','To','Cc','ReplyTo','SubjectTpl','BodyTpl','AttachIdsCSV','Active','ErrorCode'];

function fetchStagingRowByKeyHash_(ss, keyHash){
  keyHash = String(keyHash||'').trim();
  if (!keyHash) return null;
  var names = ['STAGING_INSCRIPTIONS', 'STAGING_ARTICLES'];
  for (var si=0; si<names.length; si++){
    var d = readSheetAsObjects_(ss.getId(), names[si]);
    var rows = d.rows || [];
    for (var i=0; i<rows.length; i++){
      if (String(rows[i]['KeyHash']||'').trim() === keyHash) return rows[i];
    }
  }
  return null;
}

/** "id1, id2 ; id3" -> ['id1','id2','id3'] */
function _parseAttachIdsCsv_(csv) {
  if (!csv) return [];
  return String(csv)
    .split(/[,\s;]+/)
    .map(function(s){ return String(s||'').trim(); })
    .filter(Boolean);
}

/** Charge les blobs Drive à partir d’IDs. Ignore l’ID si introuvable / accès refusé. */
function _getAttachBlobs_(ids) {
  var blobs = [];
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    try {
      var f = DriveApp.getFileById(id);
      blobs.push(f.getBlob());
    } catch (e) {
      try { Logger.log('ATTACH_WARN: ' + id + ' (' + e + ')'); } catch(_){}
    }
  }
  return blobs;
}

function _probeAttachIds_(csv) {
  var ids = _parseAttachIdsCsv_(csv);
  var res = [];
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    try {
      var f = DriveApp.getFileById(id);
      res.push({
        id: id,
        ok: true,
        name: f.getName(),
        mimeType: f.getMimeType(),
        size: (function(b){ try { return b.getBytes().length; } catch(e){ return null; } })(f.getBlob())
      });
    } catch (e) {
      res.push({ id: id, ok: false, error: String(e) });
    }
  }
  return res;
}

function _blobsFromProbe_(probe) {
  var blobs = [];
  for (var i = 0; i < probe.length; i++) {
    var p = probe[i];
    if (!p.ok) continue;
    try { blobs.push(DriveApp.getFileById(p.id).getBlob()); } catch(e){}
  }
  return blobs;
}

/** (optionnel) strip du HTML pour textBody/plain */
function _stripHtml_(html) {
  html = String(html || '');
  html = html.replace(/<style[\s\S]*?<\/style>/gi, '')
             .replace(/<script[\s\S]*?<\/script>/gi, '')
             .replace(/<[^>]+>/g, ' ')
             .replace(/&nbsp;/g, ' ')
             .replace(/\s{2,}/g, ' ')
             .trim();
  return html;
}

function _ms_sheet_(ss){
  var sh = ss.getSheetByName(MAIL_SECTORS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MAIL_SECTORS_SHEET);
    sh.getRange(1,1,1,MAIL_SECTORS_HEADER.length).setValues([MAIL_SECTORS_HEADER]);
  } else {
    // upgrade header si incomplet
    var need = sh.getLastColumn() < MAIL_SECTORS_HEADER.length;
    if (need) sh.insertColumnsAfter(sh.getLastColumn()||1, MAIL_SECTORS_HEADER.length - (sh.getLastColumn()||1));
    sh.getRange(1,1,1,MAIL_SECTORS_HEADER.length).setValues([MAIL_SECTORS_HEADER]);
  }
  return sh;
}
function _ms_headers_(){ return MAIL_SECTORS_HEADER.slice(); }

function _ms_readAll_(ss){
  var sh = _ms_sheet_(ss);
  var last = sh.getLastRow(); if (last < 2) return [];
  var lastCol = sh.getLastColumn();
  var vals = sh.getRange(1,1,last, lastCol).getValues();
  var headers = vals[0].map(String), idx = {}; headers.forEach(function(h,i){ idx[h]=i; });
  var items = [];
  for (var r=1; r<vals.length; r++){
    var V = vals[r];
    function cell(key){ var i=idx[key]; return (i==null?'':V[i]); }
    var it = {
      SecteurId: String(cell('SecteurId')||'').trim(),
      Label:     String(cell('Label')||'').trim(),
      Umin:      Number(cell('Umin')||''),
      Umax:      Number(cell('Umax')||''),
      Genre:     String(cell('Genre')||'*').trim()||'*',
      To:        String(cell('To')||'').trim(),
      Cc:        String(cell('Cc')||'').trim(),
      ReplyTo:   String(cell('ReplyTo')||'').trim(),
      SubjectTpl:String(cell('SubjectTpl')||'').trim(),
      BodyTpl:   String(cell('BodyTpl')||'').trim(),
      AttachIdsCSV: String(cell('AttachIdsCSV')||'').trim(),
      Active:    String(cell('Active')).toString().toLowerCase() !== 'false',
      ErrorCode: String(cell('ErrorCode')||'').trim(),
      _row:      r+1
    };
    items.push(it);
  }
  items.sort(function(a,b){ return (a.Umin||0)-(b.Umin||0) || (a.Umax||0)-(b.Umax||0) || String(a.Label||'').localeCompare(String(b.Label||'')); });
  return items;
}

function _ms_writeRow_(sh, rowIndex, item){
  var v = [
    item.SecteurId||'',
    item.Label||'',
    Number(item.Umin||'')||'',
    Number(item.Umax||'')||'',
    (item.Genre||'*'),
    item.To||'',
    item.Cc||'',
    item.ReplyTo||'',
    item.SubjectTpl||'',
    item.BodyTpl||'',
    item.AttachIdsCSV||'',
    (String(item.Active).toLowerCase()==='false' ? 'FALSE' : 'TRUE'),
    (item.ErrorCode||'')
  ];
  sh.getRange(rowIndex, 1, 1, MAIL_SECTORS_HEADER.length).setValues([v]);
}

/* -------------------- Résolution To/subject/body pour un row -------------------- */
function _resolveToForRow_(ss, row, sectorItem){
  // Priorité : sectorItem.To si défini, sinon TO_FIELDS_INSCRIPTIONS sur la ligne, sinon fallback param global
  if (sectorItem && sectorItem.To) return sectorItem.To;
  var csv = readParam_(ss, 'TO_FIELDS_INSCRIPTIONS') || 'Courriel,Parent 1 - Courriel,Parent 2 - Courriel';
  var to = collectEmailsFromRow_(row, csv);
  if (to) return to;
  return readParam_(ss, 'MAIL_TO_NEW_INSCRIPTIONS') || '';
}
function _resolveCcForRow_(ss, sectorItem){
  if (sectorItem && sectorItem.Cc) return sectorItem.Cc;
  return readParam_(ss, 'MAIL_CC_NEW_INSCRIPTIONS') || '';
}
function _resolveSubjectBody_(ss, sectorItem, payload){
  var subj = (sectorItem && sectorItem.SubjectTpl) ? renderTemplate_(sectorItem.SubjectTpl, payload) : '';
  var body = (sectorItem && sectorItem.BodyTpl) ? renderTemplate_(sectorItem.BodyTpl, payload) : '';
  if (!subj) subj = renderTemplate_(readParam_(ss, 'MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT') || 'Bienvenue {{prenom}} – {{frais}}', payload);
  if (!body) body = renderTemplate_(readParam_(ss, 'MAIL_TEMPLATE_INSCRIPTION_NEW_BODY') || 'Bonjour {{prenom}},<br>Nous confirmons {{frais}} ({{saison}}).', payload);
  return { subject: subj, bodyHtml: body };
}

/* -------------------- Recherche ligne INSCRIPTIONS par passeport -------------------- */
function _findFinalRowByPassport_(ss, passport){
  passport = String(passport||'').trim();
  if (!passport) return null;
  var data = readSheetAsObjects_(ss.getId(), 'INSCRIPTIONS');
  var rows = data.rows||[];
  for (var i=0;i<rows.length;i++){
    var r = rows[i];
    var p = String(r['Passeport #']||'').trim();
    if (p === passport) return r;
    if (p && p.replace(/^0+/, '') === passport.replace(/^0+/, '')) return r;
  }
  return null;
}

/* ==================== API UI ==================== */

// Liste
function getMailSectors(seasonId){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var items = _ms_readAll_(ss);
    return { ok:true, data:{ items: items } };
  } catch(e){ return { ok:false, error: String(e) }; }
}

// Create/Update (par SecteurId ou _row si fourni)
function upsertMailSector(seasonId, item){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var sh = _ms_sheet_(ss);
    var items = _ms_readAll_(ss);
    var row = 0;

    if (item && item._row) {
      row = parseInt(item._row, 10) || 0;
    } else if (item && item.SecteurId) {
      var id = String(item.SecteurId).trim();
      for (var i=0;i<items.length;i++){ if (String(items[i].SecteurId||'')===id) { row = items[i]._row; break; } }
    }

    if (!row) {
      row = sh.getLastRow() + 1;
      if (row === 1) {
        sh.getRange(1,1,1,MAIL_SECTORS_HEADER.length).setValues([MAIL_SECTORS_HEADER]);
        row = 2;
      }
    }

    _ms_writeRow_(sh, row, item||{});
    return { ok:true };
  } catch(e){ return { ok:false, error:String(e) }; }
}

// Delete
function deleteMailSector(seasonId, secteurId){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var sh = _ms_sheet_(ss);
    var items = _ms_readAll_(ss);
    var id = String(secteurId||'').trim();
    var row = 0;
    for (var i=0;i<items.length;i++){ if (String(items[i].SecteurId||'')===id) { row = items[i]._row; break; } }
    if (!row) return { ok:false, error:'Secteur introuvable' };
    sh.deleteRow(row);
    return { ok:true };
  } catch(e){ return { ok:false, error:String(e) }; }
}

// Duplicate (retourne la copie pour ouverture immédiate dans le modal)
function duplicateMailSector(seasonId, secteurId){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var sh = _ms_sheet_(ss);
    var items = _ms_readAll_(ss);
    var id = String(secteurId||'').trim();
    var src = null;
    for (var i=0;i<items.length;i++){ if (String(items[i].SecteurId||'')===id) { src = items[i]; break; } }
    if (!src) return { ok:false, error:'Secteur introuvable' };

    var base = src.SecteurId || 'SEC';
    var copyId = base + '_COPY';
    var exists = items.some(function(x){ return String(x.SecteurId||'')===copyId; });
    if (exists) copyId = base + '_' + Date.now();

    var copy = Object.assign({}, src, {
      SecteurId: copyId,
      Label: (src.Label||src.SecteurId||'') + ' (copie)',
      _row: null
    });

    var row = sh.getLastRow() + 1; if (row===1) { sh.getRange(1,1,1,MAIL_SECTORS_HEADER.length).setValues([MAIL_SECTORS_HEADER]); row=2; }
    _ms_writeRow_(sh, row, copy);

    return { ok:true, data:{ item: copy } };
  } catch(e){ return { ok:false, error:String(e) }; }
}

// Aperçu sujet/body + To résolu
function previewSectorForPassport(seasonId, secteurId, passport, itemOverride){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var row = _findFinalRowByPassport_(ss, passport);
    if (!row) return { ok:false, error:'Passeport introuvable dans INSCRIPTIONS' };

    var items = _ms_readAll_(ss);
    var current = null;
    var id = String(secteurId||'').trim();
    for (var i=0;i<items.length;i++){ if (String(items[i].SecteurId||'')===id) { current = items[i]; break; } }
    var it = Object.assign({}, current||{}, itemOverride||{});

    var payload = buildDataFromRow_(row);
    var sb = _resolveSubjectBody_(ss, it, payload);
    var to = _resolveToForRow_(ss, row, it);

    return { ok:true, data:{ subject: sb.subject, bodyHtml: sb.bodyHtml, to: to } };
  } catch(e){ return { ok:false, error:String(e) }; }
}

// Envoi test (HTML) vers toTest (sinon email du script) — n’écrit rien dans OUTBOX
function sendSectorTest(seasonId, item, passport, toTest){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var row = _findFinalRowByPassport_(ss, passport);
    if (!row) return { ok:false, error:'Passeport introuvable dans INSCRIPTIONS' };

    var payload = buildDataFromRow_(row);
    var sb = _resolveSubjectBody_(ss, item||{}, payload);

    var to = String(toTest||'').trim() || Session.getActiveUser().getEmail();
    var cc = item && item.Cc ? String(item.Cc).trim() : '';
    var replyTo = item && item.ReplyTo ? String(item.ReplyTo).trim() : '';
    var fromName = readParam_(ss, 'MAIL_FROM') || undefined;

    var probe = _probeAttachIds_(item && item.AttachIdsCSV);
    var blobs = _blobsFromProbe_(probe);

    MailApp.sendEmail({
      to: to,
      subject: sb.subject,
      htmlBody: sb.bodyHtml,
      name: fromName,
      cc: cc || undefined,
      replyTo: replyTo || undefined,
      attachments: (blobs.length ? blobs : undefined)
    });

    return { ok:true, data:{ 
      attached: blobs.length,
      probe: probe
    }};
  } catch(e){ 
    return { ok:false, error:String(e) }; 
  }
}

// Lance le worker d’envoi d’outbox (respecte MAIL_BATCH_MAX et DRY_RUN)
function runSendPendingOutbox(seasonSheetId) {
  try {
    var sid = seasonSheetId || (typeof getSeasonId_ === 'function' ? getSeasonId_() : null);

    var implFn = null;
    if (typeof sendPendingOutbox === 'function') {
      implFn = sendPendingOutbox;                   
    } else if (typeof LIB !== 'undefined' && LIB && typeof LIB.sendPendingOutbox === 'function') {
      implFn = LIB.sendPendingOutbox;               
    } else if (typeof Library !== 'undefined' && Library && typeof Library.sendPendingOutbox === 'function') {
      implFn = Library.sendPendingOutbox;
    }

    if (!implFn) {
      throw new Error('sendPendingOutbox introuvable : ni locale ni dans la librairie (vérifie le nom d’alias de la lib).');
    }

    var res = implFn(sid);
    return { ok: true, data: res, impl: (implFn === sendPendingOutbox ? 'local' : 'lib') };
  } catch (e) {
    return { ok: false, error: '' + e };
  }
}

// Aperçu ERREUR (utilise les templates du secteur courant + ajoute {{error_*}})
function previewErrorForPassport(seasonId, secteurId, passport, errorItem){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var row = _findFinalRowByPassport_(ss, passport);
    if (!row) return { ok:false, error:'Passeport introuvable dans INSCRIPTIONS' };

    var items = _ms_readAll_(ss), current=null;
    for (var i=0;i<items.length;i++){ if (String(items[i].SecteurId||'')===String(secteurId||'')) { current=items[i]; break; } }
    var it = current || {};

    var payload = buildDataFromRow_(row);
    payload.error_code    = String(errorItem && errorItem.code || '').trim();
    payload.error_label   = String(errorItem && errorItem.label || '').trim();
    payload.error_details = String(errorItem && errorItem.details || '').trim();

    var subj = renderTemplate_(it.SubjectTpl||'', payload);
    var body = renderTemplate_(it.BodyTpl||'', payload);
    if (!subj) subj = 'Validation requise – '+(payload.nomcomplet||'')+' ('+(payload.U||'')+')';
    if (!body) body = 'Bonjour '+(payload.prenom||'')+',<br>Merci de valider : <b>'+(payload.error_label||payload.error_code||'')+'</b><br><small>'+(payload.error_details||'')+'</small>';

    var to = _resolveToForRow_(ss, row, it);
    return { ok:true, data:{ subject: subj, bodyHtml: body, to: to } };
  } catch(e){ return { ok:false, error:String(e) }; }
}

// Envoi test ERREUR (n’écrit rien dans OUTBOX)
function sendErrorTest(seasonId, secteurItem, passport, toTest, errorItem){
  try{
    var ss = getSeasonSpreadsheet_(seasonId);
    var row = _findFinalRowByPassport_(ss, passport);
    if (!row) return { ok:false, error:'Passeport introuvable dans INSCRIPTIONS' };

    var payload = buildDataFromRow_(row);
    payload.error_code    = String(errorItem && errorItem.code || '').trim();
    payload.error_label   = String(errorItem && errorItem.label || '').trim();
    payload.error_details = String(errorItem && errorItem.details || '').trim();

    var subj = renderTemplate_(secteurItem && secteurItem.SubjectTpl || '', payload);
    var body = renderTemplate_(secteurItem && secteurItem.BodyTpl || '', payload);
    if (!subj) subj = 'Validation requise – '+(payload.nomcomplet||'')+' ('+(payload.U||'')+')';
    if (!body) body = 'Bonjour '+(payload.prenom||'')+',<br>Merci de valider : <b>'+(payload.error_label||payload.error_code||'')+'</b>';

    var to = String(toTest||'').trim() || Session.getActiveUser().getEmail();
    var fromName = readParam_(ss, 'MAIL_FROM') || undefined;

    MailApp.sendEmail({ to: to, subject: subj, htmlBody: body, name: fromName });
    return { ok:true };
  } catch(e){ return { ok:false, error:String(e) }; }
}

// Queue des nouveaux courriels en mode FAST (dashboard → lib)
function runQueueNewFast_(ss){
  var sid = (ss && typeof ss.getId === 'function') ? ss.getId()
           : (typeof getSeasonId_ === 'function' ? getSeasonId_() : null);

  var lib = (typeof LIB !== 'undefined' && LIB) ? LIB
          : ((typeof Library !== 'undefined' && Library) ? Library : null);

  if (!lib) throw new Error('Librairie Phénix non chargée (alias LIB/Library manquant).');

  // Assure que MAIL_OUTBOX a les colonnes lisibles (L/M/N) dès maintenant
  try {
    var ssLocal = (typeof SpreadsheetApp !== 'undefined') ? SpreadsheetApp.openById(sid) : null;
    if (ssLocal && typeof upgradeMailOutboxForDisplay_ === 'function') upgradeMailOutboxForDisplay_(ssLocal);
  } catch (_e) {}

  var __t0 = Date.now();

  // Enqueue bienvenue (JOUEURS) + validations (ERREURS)
  var r1 = (typeof lib.enqueueWelcomeFromJoueursFast_ === 'function')
             ? lib.enqueueWelcomeFromJoueursFast_(sid) : { queued: 0 };
  var __tWelcome = Date.now();
  appendImportLog_(ss, 'QUEUE_TIMING', JSON.stringify({ step:'fast_welcome_done', ms:(__tWelcome-__t0), queued:r1.queued||0 }));

  var r2 = (typeof lib.enqueueValidationMailsFromErreursFast_ === 'function')
             ? lib.enqueueValidationMailsFromErreursFast_(sid) : { queued: 0 };
  var __tErrors = Date.now();
  appendImportLog_(ss, 'QUEUE_TIMING', JSON.stringify({ step:'fast_errors_done', ms:(__tErrors-__tWelcome), queued:r2.queued||0 }));
  appendImportLog_(ss, 'QUEUE_TIMING', JSON.stringify({ step:'fast_total', ms:(__tErrors-__t0), queued_total:(r1.queued||0)+(r2.queued||0) }));
  Logger.log(JSON.stringify({ welcome:r1, errors:r2 }));

  return { queued: (r1.queued || 0) + (r2.queued || 0) };
}

// Worker (optionnel) – brancher sur la fonction existante runSendPendingOutbox
function runMailWorker(ss){
  var sid = (ss && typeof ss.getId === 'function') ? ss.getId()
           : (typeof getSeasonId_ === 'function' ? getSeasonId_() : null);
  var res = runSendPendingOutbox(sid);
  return (res && res.data) ? res.data : { processed: 0 };
}

/**
 * Queue all new emails (LEGACY mode) — redirigé vers FAST.
 */
function runQueueNew_(ss) {
  // rétrocompat : on redirige vers la voie FAST
  return runQueueNewFast_(ss);
}
/**
 * 🔗 Pipeline sélectionné par le flow (server_flow.runImportRulesExportsFull)
 * Appelé avec (ss, 'AFTER') — enchaîne enqueue FAST + worker; loggue dans IMPORT_LOG.
 */
function runMailPipelineSelected_(ssOrId, stage){
  // Resolve season + spreadsheet
  var sid = (ssOrId && typeof ssOrId.getId === 'function') ? ssOrId.getId()
           : (typeof ssOrId === 'string' ? ssOrId
           : (typeof getSeasonId_ === 'function' ? getSeasonId_() : null));
  var ss  = (ssOrId && typeof ssOrId.getId === 'function') ? ssOrId : SpreadsheetApp.openById(sid);
  var stg = String(stage||'AFTER').toUpperCase();

  try { if (typeof appendImportLog_==='function') appendImportLog_(ss,'MAIL_PIPELINE_START', JSON.stringify({ stage: stg })); } catch(_){}

  // 1) Enqueue FAST – INSCRIPTION_NEW (JOUEURS)
  var qWelcome = (typeof enqueueWelcomeFromJoueursFast_==='function')
    ? enqueueWelcomeFromJoueursFast_(sid)
    : (typeof LIB!=='undefined' && LIB.enqueueWelcomeFromJoueursFast_ ? LIB.enqueueWelcomeFromJoueursFast_(sid) : { queued: 0 });
  try { if (typeof appendImportLog_==='function') appendImportLog_(ss,'MAIL_QUEUE_WELCOME_OK', JSON.stringify(qWelcome)); } catch(_){}

  // 2) Enqueue FAST – Erreurs (ERREURS → MAIL_SECTEURS.ErrorCode)
  var qErrors = (typeof enqueueValidationMailsFromErreursFast_==='function')
    ? enqueueValidationMailsFromErreursFast_(sid)
    : (typeof LIB!=='undefined' && LIB.enqueueValidationMailsFromErreursFast_ ? LIB.enqueueValidationMailsFromErreursFast_(sid) : { queued: 0 });
  try { if (typeof appendImportLog_==='function') appendImportLog_(ss,'MAIL_QUEUE_ERRORS_OK', JSON.stringify(qErrors)); } catch(_){}

  // 3) Worker – envoie pending (honore DRY_RUN / DRY_REDIRECT_EMAIL)
  var workerRes = (typeof sendPendingOutbox==='function')
    ? sendPendingOutbox(sid)
    : (typeof LIB!=='undefined' && LIB.sendPendingOutbox ? LIB.sendPendingOutbox(sid) : { processed: 0, sent: 0, errors: 0 });
  try {
    if (typeof appendImportLog_==='function') {
      appendImportLog_(ss,'MAIL_WORKER_OK', JSON.stringify(workerRes));
      appendImportLog_(ss,'MAIL_PIPELINE_END', JSON.stringify({ stage: stg, queued_total: (qWelcome.queued||0)+(qErrors.queued||0) }));
    }
  } catch(_){}

  return { ok:true, queued: (qWelcome.queued||0)+(qErrors.queued||0), welcome: qWelcome, errors: qErrors, worker: workerRes };
}
