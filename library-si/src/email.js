/**
 * email.gs — v0.7
 * - Dépile MAIL_OUTBOX (Status=pending) et envoie les courriels INSCRIPTION_NEW
 * - Concatène TO depuis Courriel, Parent 1 - Courriel, Parent 2 - Courriel (configurable)
 * - Dans la même passe, envoie un RÉSUMÉ PAR SECTEUR (U4-U8, U9-U12, U13-U18) avec CSV en PJ
 *
 * Dépendances attendues (si non présentes, fallbacks ci-dessous):
 *  - getSeasonSpreadsheet_, ensureMailOutbox_, getMailOutboxHeaders_, getHeadersIndex_
 *  - readParam_, readSheetAsObjects_, getSheetOrCreate_, appendImportLog_
 *  - deriveSectorFromRow_, collectEmailsFromRow_
 *  - SHEETS, PARAM_KEYS
 */

/* ======================== Fallbacks (seulement si absents) ======================== */
if (typeof SHEETS === 'undefined') {
  var SHEETS = {
    INSCRIPTIONS: 'INSCRIPTIONS',
    MAIL_OUTBOX: 'MAIL_OUTBOX',
    MAIL_LOG: 'MAIL_LOG',
    PARAMS: 'PARAMS'
  };
}
if (typeof PARAM_KEYS === 'undefined') {
  var PARAM_KEYS = {
    KEY_COLS: 'KEY_COLS',
    DRY_RUN: 'DRY_RUN',
    MAIL_FROM: 'MAIL_FROM',
    MAIL_BATCH_MAX: 'MAIL_BATCH_MAX',
    TO_FIELDS_INSCRIPTIONS: 'TO_FIELDS_INSCRIPTIONS',
    MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT: 'MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT',
    MAIL_TEMPLATE_INSCRIPTION_NEW_BODY: 'MAIL_TEMPLATE_INSCRIPTION_NEW_BODY',
    MAIL_TO_SUMMARY_U4U8: 'MAIL_TO_SUMMARY_U4U8',
    MAIL_CC_SUMMARY_U4U8: 'MAIL_CC_SUMMARY_U4U8',
    MAIL_TO_SUMMARY_U9U12: 'MAIL_TO_SUMMARY_U9U12',
    MAIL_CC_SUMMARY_U9U12: 'MAIL_CC_SUMMARY_U9U12',
    MAIL_TO_SUMMARY_U13U18: 'MAIL_TO_SUMMARY_U13U18',
    MAIL_CC_SUMMARY_U13U18: 'MAIL_CC_SUMMARY_U13U18',
    MAIL_TEMPLATE_SUMMARY_SUBJECT: 'MAIL_TEMPLATE_SUMMARY_SUBJECT',
    MAIL_TEMPLATE_SUMMARY_BODY: 'MAIL_TEMPLATE_SUMMARY_BODY'
  };
}
if (typeof getSeasonSpreadsheet_ !== 'function') {
  function getSeasonSpreadsheet_(seasonSheetId){
    if (!seasonSheetId) throw new Error('seasonSheetId manquant');
    return SpreadsheetApp.openById(seasonSheetId);
  }
}
if (typeof getSheetOrCreate_ !== 'function') {
  function getSheetOrCreate_(ss, name, header){
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      if (header && header.length) sh.getRange(1,1,1,header.length).setValues([header]);
    } else if (header && header.length && sh.getLastRow() === 0) {
      sh.getRange(1,1,1,header.length).setValues([header]);
    }
    return sh;
  }
}
if (typeof getMailOutboxHeaders_ !== 'function') {
function getMailOutboxHeaders_() {
  return [
    'Type','To','Cc','Sujet','Corps','Attachments','KeyHash','Status','CreatedAt','SentAt','Error',
    // --- nouvelles colonnes d’info (lecture/tri) :
    'Passeport8','Nom','Prénom','NomComplet','Saison','Frais','EmailsCandidates'
  ];
}
}
if (typeof ensureMailOutbox_ !== 'function') {
/** Crée/upgrade l’OUTBOX: si l’entête existante est un préfixe de la nouvelle, on complète proprement. */
function ensureMailOutbox_(ss) {
  var sheetName = (typeof SHEETS !== 'undefined' && SHEETS.MAIL_OUTBOX) ? SHEETS.MAIL_OUTBOX : 'MAIL_OUTBOX';
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  var want = getMailOutboxHeaders_();

  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,want.length).setValues([want]);
    return sh;
  }
    var firstRow = sh.getRange(1,1,1,Math.max(want.length, sh.getLastColumn())).getDisplayValues()[0];
  // si l’entête actuelle est un préfixe de la nouvelle → on complete à droite
  var isPrefix = true;
  for (var i=0;i<Math.min(firstRow.length, want.length); i++) {
    if ((firstRow[i]||'') !== want[i]) { isPrefix = false; break; }
  }
  if (isPrefix && firstRow.length < want.length) {
    sh.getRange(1, firstRow.length+1, 1, want.length-firstRow.length).setValues([want.slice(firstRow.length)]);
    return sh;
  }

  // Si entête incompatible → on insère une nouvelle ligne d’entête propre (préserve l’historique).
  sh.insertRowsBefore(1,1);
  sh.getRange(1,1,1,want.length).setValues([want]);
  return sh;
}
}
if (typeof getHeadersIndex_ !== 'function') {
  function getHeadersIndex_(sh, width){
    var headers = sh.getRange(1,1,1,width||sh.getLastColumn()).getValues()[0].map(String);
    var idx = {}; headers.forEach(function(h,i){ idx[h]=i+1; }); return idx; // 1-based
  }
}
if (typeof readParam_ !== 'function') {
  function readParam_(ss, key){
    var sh = ss.getSheetByName(SHEETS.PARAMS);
    if (sh) {
      var last = sh.getLastRow();
      if (last >= 1) {
        var data = sh.getRange(1,1,last,2).getValues();
        for (var i=0;i<data.length;i++){
          if ((data[i][0]+'').trim() === key) return (data[i][1]+'').trim();
        }
      }
    }
    var props = PropertiesService.getDocumentProperties();
    return (props.getProperty(key) || '').trim();
  }
}
if (typeof readSheetAsObjects_ !== 'function') {
  function readSheetAsObjects_(ssId, sheetName){
    var ss = SpreadsheetApp.openById(ssId);
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
      return { sheet: sh || ss.insertSheet(sheetName), headers: [], rows: [] };
    }
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    var values = sh.getRange(1,1,lastRow,lastCol).getDisplayValues(); // display pour préserver zéros
    var headers = values[0].map(String);
    var rows = [];
    for (var r=1; r<values.length; r++){
      var o = {};
      for (var c=0; c<headers.length; c++) o[headers[c]] = values[r][c];
      rows.push(o);
    }
    return { sheet: sh, headers: headers, rows: rows };
  }
}
if (typeof appendImportLog_ !== 'function') {
  function appendImportLog_(ss, action, details){
    var sh = getSheetOrCreate_(ss, 'IMPORT_LOG', ['Horodatage','Action','Détails']);
    sh.appendRow([new Date(), action, details||'']);
  }
}
/* --- utils pour U et secteur (fallback) --- */
if (typeof deriveSectorFromRow_ !== 'function') {
  function parseSeasonYear_(s){ var m=(String(s||'').match(/(20\d{2})/)); return m?parseInt(m[1],10):(new Date()).getFullYear(); }
  function birthYearFromRow_(row){
    var y = row['Année de naissance']||row['Annee de naissance']||row['Annee']||'';
    if (y && /^\d{4}$/.test(String(y))) return parseInt(y,10);
    var dob = row['Date de naissance']||'';
    if (dob){
      var s=String(dob), m=s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if (m) return parseInt(m[1],10);
      var m2=s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/); if (m2) return parseInt(m2[3],10);
    }
    return null;
  }
  function computeUForYear_(by, sy){ if(!by||!sy) return null; var u=sy-by; return (u>=4&&u<=21)?('U'+u):null; }
  function deriveUFromRow_(row){
    var cat=row['Catégorie']||row['Categorie']||''; if (cat && /^U\d{1,2}/i.test(cat)) return cat.toUpperCase().replace(/\s+/g,'');
    var U = computeUForYear_(birthYearFromRow_(row), parseSeasonYear_(row['Saison']||''));
    return U||'';
  }
  function deriveSectorFromRow_(row){
    var U = deriveUFromRow_(row), n=parseInt(String(U).replace(/^U/i,''),10);
    if (!n||isNaN(n)) return 'AUTRES';
    if (n>=4 && n<=8) return 'U4-U8';
    if (n>=9 && n<=12) return 'U9-U12';
    if (n>=13 && n<=18) return 'U13-U18';
    return 'AUTRES';
  }
}
if (typeof collectEmailsFromRow_ !== 'function') {
  function norm_(s){ s=String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s.trim(); }
  function collectEmailsFromRow_(row, fieldsCsv){
    var fields=(fieldsCsv&&fieldsCsv.length)?fieldsCsv.split(',').map(function(x){return x.trim();}).filter(Boolean):['Courriel','Parent 1 - Courriel','Parent 2 - Courriel'];
    var set={};
    fields.forEach(function(f){
      var v=row[f]; if(!v) return;
      String(v).split(/[;,]/).forEach(function(e){ e=norm_(e); if(!e) return; set[e]=true; });
    });
    return Object.keys(set).join(',');
  }
}
/* ====================== Fin des fallbacks ====================== */


/* ============================ Core v0.7 ============================ */

function renderTemplate_(tpl, data) {
  tpl = String(tpl == null ? '' : tpl);
  return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function(_, key){
    var v = data.hasOwnProperty(key) ? data[key] : '';
    return (v == null ? '' : String(v));
  });
}

function buildDataFromRow_(row) {
  var data = {
    passeport: row['Passeport #'] || '',
    nom: row['Nom'] || '',
    prenom: row['Prénom'] || row['Prenom'] || '',
    nomcomplet: (((row['Prénom']||row['Prenom']||'') + ' ' + (row['Nom']||'')).trim()),
    saison: row['Saison'] || '',
    frais: row['Nom du frais'] || row['Frais'] || row['Produit'] || '',
    categorie: row['Catégorie'] || row['Categorie'] || '',
    secteur: deriveSectorFromRow_(row) || ''
  };
  return data;
}

function resolveRecipient_(ss, type, row) {
  var to = '', cc = '';
  if (type === 'INSCRIPTION_NEW') {
    var csv = readParam_(ss, PARAM_KEYS.TO_FIELDS_INSCRIPTIONS) || 'Courriel,Parent 1 - Courriel,Parent 2 - Courriel';
    to = collectEmailsFromRow_(row, csv);
  }
  // Fallback éventuels par PARAMS (peu utilisé ici)
  if (!to) {
    var t = readParam_(ss, 'MAIL_TO_NEW_INSCRIPTIONS') || '';
    var c = readParam_(ss, 'MAIL_CC_NEW_INSCRIPTIONS') || '';
    to = t; cc = c;
  }
  return { to: to, cc: cc };
}

function fetchFinalRowByKeyHash_(ss, keyHash) {
  // key = base64websafe decode → "Passeport||Nom du frais||Saison"
  var keyStr = Utilities.newBlob(Utilities.base64DecodeWebSafe(keyHash)).getDataAsString();
  var finals = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var keyColsCsv = readParam_(ss, PARAM_KEYS.KEY_COLS) || 'Passeport #,Nom du frais,Saison';
  var keyCols = keyColsCsv.split(',').map(function(x){ return x.trim(); });

  for (var i=0;i<finals.rows.length;i++){
    var r = finals.rows[i];
    var k = keyCols.map(function(kc){ return r[kc] == null ? '' : String(r[kc]); }).join('||');
    if (k === keyStr) return r;
  }
  return null;
}

/**
 * Dépile et envoie les emails + résumés par secteur
 */
function sendPendingOutbox(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var shOut = ensureMailOutbox_(ss);
  var headers = getMailOutboxHeaders_();
  var idx = getHeadersIndex_(shOut, headers.length);

  var last = shOut.getLastRow();
  if (last < 2) return { processed: 0, summaries: false };

  var batchMax = parseInt(readParam_(ss, PARAM_KEYS.MAIL_BATCH_MAX) || '50', 10);
  var dry = (readParam_(ss, PARAM_KEYS.DRY_RUN) || 'FALSE').toUpperCase() === 'TRUE';
  var fromName = readParam_(ss, PARAM_KEYS.MAIL_FROM) || undefined;

  var data = shOut.getRange(2,1,last-1,headers.length).getValues();
  var processed = 0;

  // Pour le résumé par secteur
  var processedNew = []; // {row, data}

  for (var i=0; i<data.length && processed < batchMax; i++) {
    var row = {};
    headers.forEach(function(h, j){ row[h] = data[i][j]; });

    if (String(row['Status']).toLowerCase() !== 'pending') continue;
    if (row['SentAt']) continue;

    var type = row['Type'];
    var keyHash = row['KeyHash'];

    // On va chercher la ligne INSCRIPTION correspondante
    var fRow = fetchFinalRowByKeyHash_(ss, keyHash) || {};
    var payload = buildDataFromRow_(fRow);
    var rcpt = resolveRecipient_(ss, type, fRow);

    var subject = '', body = '';

    if (type === 'INSCRIPTION_NEW') {
      subject = renderTemplate_(readParam_(ss, PARAM_KEYS.MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT) || 'Bienvenue {{prenom}} – {{frais}}', payload);
      body    = renderTemplate_(readParam_(ss, PARAM_KEYS.MAIL_TEMPLATE_INSCRIPTION_NEW_BODY)    || 'Bonjour {{prenom}},<br>Nous confirmons votre inscription à {{frais}} ({{saison}}).', payload);
    } else {
      // On ignore les autres types en v0.7 (annulations gérées côté Rétroaction)
      shOut.getRange(i+2, idx['Status']).setValue('skipped');
      continue;
    }

    if (!dry) {
      MailApp.sendEmail({
        to: rcpt.to || '',
        cc: rcpt.cc || '',
        subject: subject,
        htmlBody: body,
        name: fromName
      });
    }

    // Marquage "sent"
    shOut.getRange(i+2, idx['SentAt']).setValue(new Date());
    shOut.getRange(i+2, idx['Status']).setValue('sent');

    // Log
    var shLog = getSheetOrCreate_(ss, SHEETS.MAIL_LOG, ['Type','To','Sujet','KeyHash','SentAt','Result']);
    shLog.appendRow([type, rcpt.to, subject, keyHash, new Date(), dry ? 'DRY_RUN' : 'SENT']);

    processed++;
    if (type === 'INSCRIPTION_NEW') processedNew.push({ row: fRow, data: payload });
  }

  // Résumés par secteur (si on a envoyé au moins un INSCRIPTION_NEW)
  if (processedNew.length) {
    var groups = { 'U4-U8':[], 'U9-U12':[], 'U13-U18':[] };
    processedNew.forEach(function(x){
      var sec = x.data.secteur || deriveSectorFromRow_(x.row);
      if (groups[sec]) groups[sec].push(x);
    });

    var tplSub = readParam_(ss, PARAM_KEYS.MAIL_TEMPLATE_SUMMARY_SUBJECT) || 'Nouveaux inscrits – {{secteur}} – {{date}}';
    var tplBody = readParam_(ss, PARAM_KEYS.MAIL_TEMPLATE_SUMMARY_BODY) || 'Bonjour,<br>Veuillez trouver la liste des nouveaux inscrits {{secteur}} en pièce jointe.<br><br>Bonne journée.';

    Object.keys(groups).forEach(function(sec){
      var arr = groups[sec];
      if (!arr.length) return;

      var toKey = sec === 'U4-U8' ? PARAM_KEYS.MAIL_TO_SUMMARY_U4U8 : (sec === 'U9-U12' ? PARAM_KEYS.MAIL_TO_SUMMARY_U9U12 : PARAM_KEYS.MAIL_TO_SUMMARY_U13U18);
      var ccKey = sec === 'U4-U8' ? PARAM_KEYS.MAIL_CC_SUMMARY_U4U8 : (sec === 'U9-U12' ? PARAM_KEYS.MAIL_CC_SUMMARY_U9U12 : PARAM_KEYS.MAIL_CC_SUMMARY_U13U18);
      var to = readParam_(ss, toKey) || '';
      var cc = readParam_(ss, ccKey) || '';

      // Compose CSV
      var headersCSV = ['Passeport','NomComplet','Saison','Frais','Categorie','Secteur'];
      var lines = [headersCSV.join(',')];
      arr.forEach(function(x){
        var r = x.row;
        var nomc = ((r['Prénom']||r['Prenom']||'') + ' ' + (r['Nom']||'')).trim();
        var vals = [
          (r['Passeport #']||''),
          nomc,
          (r['Saison']||''),
          (r['Nom du frais']||r['Frais']||r['Produit']||''),
          (r['Catégorie']||r['Categorie']||''),
          (deriveSectorFromRow_(r)||'')
        ].map(function(v){
          v = String(v).replace(/"/g,'""');
          if (/[",\n;]/.test(v)) v = '"' + v + '"';
          return v;
        });
        lines.push(vals.join(','));
      });
      var csv = lines.join('\n');
      var blob = Utilities.newBlob(csv, 'text/csv', 'Nouveaux_' + sec.replace(/[^A-Za-z0-9\-]/g,'') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.csv');

      var subject = renderTemplate_(tplSub, { secteur: sec, date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') });
      var body = renderTemplate_(tplBody, { secteur: sec });

      if (!dry && to) {
        MailApp.sendEmail({
          to: to,
          cc: cc,
          subject: subject,
          htmlBody: body,
          attachments: [blob],
          name: fromName
        });
      }
    });
  }

  appendImportLog_(ss, 'MAIL_WORKER', JSON.stringify({ processed: processed, summaries: processedNew.length>0 }));
  return { processed: processed, summaries: processedNew.length>0 };
}
