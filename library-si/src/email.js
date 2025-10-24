/**
 * email.gs — v0.8 (secteurs configurables + attachments Drive par secteur)
 *
 * Ajouts v0.8 :
 *  - Feuille "MAIL_SECTEURS" (config UI) : SecteurId, Label, Umin, Umax, Genre, To, Cc, ReplyTo,
 *    SubjectTpl, BodyTpl (HTML), AttachIdsCSV, Active
 *  - Sélection du secteur par U (et Genre optionnel), puis templates + attachments sectoriels
 *    pour les envois "INSCRIPTION_NEW" de MAIL_OUTBOX
 *  - Si pas de secteur applicable → fallback sur paramètres globaux existants
 *
 * On conserve :
 *  - Le worker sendPendingOutbox() qui dépile MAIL_OUTBOX
 *  - Les résumés par secteur U4-U8 / U9-U12 / U13-U18 (CSV en PJ) tels que v0.7
 *
 * Dépendances (gérées par la lib utils.js si présentes) :
 *  - getSeasonSpreadsheet_, ensureMailOutbox_, getMailOutboxHeaders_, getHeadersIndex_,
 *    readParam_, readSheetAsObjects_, getSheetOrCreate_, appendImportLog_,
 *    deriveSectorFromRow_, collectEmailsFromRow_, PARAM_KEYS, SHEETS
 */

/* ======================== Fallbacks (identiques v0.7, abrégés) ======================== */
if (typeof SHEETS === 'undefined') {
  var SHEETS = { INSCRIPTIONS: 'INSCRIPTIONS', MAIL_OUTBOX: 'MAIL_OUTBOX', PARAMS: 'PARAMS' };
}
if (typeof PARAM_KEYS === 'undefined') {
  var PARAM_KEYS = {
    KEY_COLS: 'KEY_COLS', DRY_RUN: 'DRY_RUN', MAIL_FROM: 'MAIL_FROM', MAIL_BATCH_MAX: 'MAIL_BATCH_MAX',
    TO_FIELDS_INSCRIPTIONS: 'TO_FIELDS_INSCRIPTIONS',
    MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT: 'MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT',
    MAIL_TEMPLATE_INSCRIPTION_NEW_BODY: 'MAIL_TEMPLATE_INSCRIPTION_NEW_BODY',
    MAIL_TO_SUMMARY_U4U8: 'MAIL_TO_SUMMARY_U4U8', MAIL_CC_SUMMARY_U4U8: 'MAIL_CC_SUMMARY_U4U8',
    MAIL_TO_SUMMARY_U9U12: 'MAIL_TO_SUMMARY_U9U12', MAIL_CC_SUMMARY_U9U12: 'MAIL_CC_SUMMARY_U9U12',
    MAIL_TO_SUMMARY_U13U18: 'MAIL_TO_SUMMARY_U13U18', MAIL_CC_SUMMARY_U13U18: 'MAIL_CC_SUMMARY_U13U18',
    MAIL_TEMPLATE_SUMMARY_SUBJECT: 'MAIL_TEMPLATE_SUMMARY_SUBJECT',
    MAIL_TEMPLATE_SUMMARY_BODY: 'MAIL_TEMPLATE_SUMMARY_BODY'
  };
}
if (typeof getSeasonSpreadsheet_ !== 'function') { function getSeasonSpreadsheet_(id) { if (!id) throw new Error('seasonSheetId manquant'); return SpreadsheetApp.openById(id); } }
if (typeof getSheetOrCreate_ !== 'function') {
  function getSheetOrCreate_(ss, name, header) {
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); if (header && header.length) sh.getRange(1, 1, 1, header.length).setValues([header]); }
    else if (header && header.length && sh.getLastRow() === 0) { sh.getRange(1, 1, 1, header.length).setValues([header]); }
    return sh;
  }
}
if (typeof getHeadersIndex_ !== 'function') {
  function getHeadersIndex_(sh, width) { var headers = sh.getRange(1, 1, 1, width || sh.getLastColumn()).getValues()[0].map(String); var idx = {}; headers.forEach(function (h, i) { idx[h] = i + 1; }); return idx; }
}
if (typeof readParam_ !== 'function') {
  function readParam_(ss, key) {
    var sh = ss.getSheetByName(SHEETS.PARAMS);
    if (sh) { var last = sh.getLastRow(); if (last >= 1) { var data = sh.getRange(1, 1, last, 2).getValues(); for (var i = 0; i < data.length; i++) { if ((data[i][0] + '').trim() === key) return (data[i][1] + '').trim(); } } }
    var props = PropertiesService.getDocumentProperties(); return (props.getProperty(key) || '').trim();
  }
}
if (typeof readSheetAsObjects_ !== 'function') {
  function readSheetAsObjects_(ssId, sheetName) {
    var ss = SpreadsheetApp.openById(ssId); var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) return { sheet: sh || ss.insertSheet(sheetName), headers: [], rows: [] };
    var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    var headers = values[0].map(String), rows = [];
    for (var r = 1; r < values.length; r++) { var o = {}; for (var c = 0; c < headers.length; c++) o[headers[c]] = values[r][c]; rows.push(o); }
    return { sheet: sh, headers: headers, rows: rows };
  }
}
if (typeof appendImportLog_ !== 'function') {
  function appendImportLog_(ss, action, details) { var sh = getSheetOrCreate_(ss, 'IMPORT_LOG', ['Horodatage', 'Action', 'Détails']); sh.appendRow([new Date(), action, details || '']); }
}
/* U/secteur fallbacks */
if (typeof deriveSectorFromRow_ !== 'function') {
  function parseSeasonYear_(s) { var m = (String(s || '').match(/(20\d{2})/)); return m ? parseInt(m[1], 10) : (new Date()).getFullYear(); }
  function birthYearFromRow_(row) {
    var y = row['Année de naissance'] || row['Annee de naissance'] || row['Annee'] || ''; if (y && /^\d{4}$/.test(String(y))) return parseInt(y, 10);
    var dob = row['Date de naissance'] || ''; if (dob) { var s = String(dob), m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if (m) return parseInt(m[1], 10); var m2 = s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/); if (m2) return parseInt(m2[3], 10); }
    return null;
  }
  function computeUForYear_(by, sy) { if (!by || !sy) return null; var u = sy - by; return (u >= 4 && u <= 21) ? ('U' + u) : null; }
  function deriveUFromRow_(row) {
    var cat = row['Catégorie'] || row['Categorie'] || ''; if (cat && /^U\d{1,2}/i.test(cat)) return cat.toUpperCase().replace(/\s+/g, '');
    var U = computeUForYear_(birthYearFromRow_(row), parseSeasonYear_(row['Saison'] || '')); return U || '';
  }
  function deriveSectorFromRow_(row) {
    var U = deriveUFromRow_(row), n = parseInt(String(U).replace(/^U/i, ''), 10);
    if (!n || isNaN(n)) return 'AUTRES';
    if (n >= 4 && n <= 8) return 'U4-U8'; if (n >= 9 && n <= 12) return 'U9-U12'; if (n >= 13 && n <= 18) return 'U13-U18';
    return 'AUTRES';
  }
}
if (typeof collectEmailsFromRow_ !== 'function') {
  function norm_(s) { s = String(s == null ? '' : s); try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return s.trim(); }
  function collectEmailsFromRow_(row, fieldsCsv) {
    var fields = (fieldsCsv && fieldsCsv.length) ? fieldsCsv.split(',').map(function (x) { return x.trim(); }).filter(Boolean) : ['Courriel', 'Parent 1 - Courriel', 'Parent 2 - Courriel'];
    var set = {}; fields.forEach(function (f) { var v = row[f]; if (!v) return; String(v).split(/[;,]/).forEach(function (e) { e = norm_(e); if (!e) return; set[e] = true; }); });
    return Object.keys(set).join(',');
  }
}

// Fallback minimal si la lib ne fournit pas _stripHtml_
if (typeof _stripHtml_ !== 'function') {
  function _stripHtml_(html) {
    html = String(html == null ? '' : html);
    // virer les tags + compacter les espaces
    return html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  }
}


// ===== Coach detection shim (lib-safe) =====
// Utilise la version lib (rules.js) si dispo, sinon fallback autonome.
// Zéro SpreadsheetApp.getActive() — on passe "ss".

function _coachCsv_(ss) {
  var csv = '';
  try { if (typeof readParam_ === 'function') csv = readParam_(ss, 'RETRO_COACH_FEES_CSV') || ''; } catch (_) { }
  if (!csv) csv = 'Entraîneurs, Entraineurs, Entraîneur, Entraineur, Coach, Coaches';
  return csv;
}
function _normNoAccentsLower_(s) {
  s = String(s == null ? '' : s).trim();
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (_) { }
  return s.toLowerCase();
}
function _isCoachFeeByNameSafe_(ss, name) {
  // 1) Si la lib rules.js fournit isCoachFeeByName_, on l'utilise
  try { if (typeof isCoachFeeByName_ === 'function') return !!isCoachFeeByName_(ss, name); } catch (_) { }
  // 2) Fallback par mots-clés
  var v = _normNoAccentsLower_(name);
  if (!v) return false;
  var toks = _coachCsv_(ss).split(',').map(_normNoAccentsLower_).filter(Boolean);
  if (toks.some(function (t) { return v === t || v.indexOf(t) >= 0; })) return true;
  // filet
  return /(entraineur|entra[îi]neur|coach)/i.test(String(name || ''));
}
function _isCoachMemberSafe_(ss, row) {
  var name = row ? (row['Nom du frais'] || row['Frais'] || row['Produit'] || '') : '';
  // 1) Si la lib rules.js fournit isCoachMember_, on l'utilise
  try { if (typeof isCoachMember_ === 'function') return !!isCoachMember_(ss, row); } catch (_) { }
  // 2) Fallback
  return _isCoachFeeByNameSafe_(ss, name);
}



/* ======================== Helpers communs ======================== */

function _rg_csvEsc_(v) { v = v == null ? '' : String(v).replace(/"/g, '""'); return /[",\n;]/.test(v) ? ('"' + v + '"') : v; }

/** "id1, id2 ; id3" -> ['id1','id2','id3'] */
function _parseAttachIdsCsv_(csv) {
  if (!csv) return [];
  return String(csv)
    .split(/[,\s;]+/)
    .map(function (s) { return String(s || '').trim(); })
    .filter(Boolean);
}

/** IDs Drive -> Array<Blob> (ignore IDs invalides / accès refusé) */
function _getAttachBlobsByIds_(ids) {
  var blobs = [];
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    try {
      var f = DriveApp.getFileById(id);
      blobs.push(f.getBlob());
    } catch (e) {
      try { Logger.log('ATTACH_WARN ' + id + ': ' + e); } catch (_) { }
    }
  }
  return blobs;
}

/** CSV d’IDs Drive -> Array<Blob> (dedup simple via nom+size) */
function _attachmentsFromCsv_(csv) {
  var ids = _parseAttachIdsCsv_(csv);
  var blobs = _getAttachBlobsByIds_(ids);
  if (!blobs || !blobs.length) return [];
  // dédoublonnage basique : nom+taille (évite doublons si sector+row incluent le même ID)
  var seen = {};
  var unique = [];
  for (var i = 0; i < blobs.length; i++) {
    var b = blobs[i];
    var key = (b.getName() || '') + '|' + (b.getBytes() ? b.getBytes().length : b.getDataAsString().length);
    if (!seen[key]) { seen[key] = 1; unique.push(b); }
  }
  return unique;
}



function _normText(s) { s = String(s == null ? '' : s); try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return s.trim(); }
function renderTemplate_(tpl, data) { tpl = String(tpl == null ? '' : tpl); return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function (_, k) { var v = (data.hasOwnProperty(k) ? data[k] : ''); return (v == null ? '' : String(v)); }); }
// Helpers (ajoute-les si non présents dans le fichier)
function _genreInitFromRow_(row) {
  var lbl = row['Identité de genre'] || row['Identité de Genre'] || row['Genre'] || '';
  var g = String(lbl || '').toUpperCase().trim();
  if (!g) return 'X';
  if (g[0] === 'M') return 'M';
  if (g[0] === 'F') return 'F';
  return 'X';
}
function _U_U2_FromRow_(row) {
  var U = deriveUFromRow_(row) || '';                 // utilise computeUForYear_/birthYearFromRow_/parseSeasonYear_
  var n = parseInt(String(U).replace(/^U/i, ''), 10);
  var U2 = (!isNaN(n) ? ('U' + (n < 10 ? ('0' + n) : n)) : '');
  return { U: U, U2: U2, n: (isNaN(n) ? '' : n) };
}
function buildDataFromRow_(row) {
  var d = _U_U2_FromRow_(row);

  var prenomRaw = row['Prénom'] || row['Prenom'] || '';
  var nomRaw = row['Nom'] || '';

  var prenomPC = _toProperCase_(prenomRaw);
  var nomPC = _toProperCase_(nomRaw);

  return {
    passeport: row['Passeport #'] || '',
    nom: nomPC,
    prenom: prenomPC,
    nomcomplet: (prenomPC + ' ' + nomPC).trim(),
    saison: row['Saison'] || '',
    frais: row['Nom du frais'] || row['Frais'] || row['Produit'] || '',
    categorie: row['Catégorie'] || row['Categorie'] || '',
    secteur: deriveSectorFromRow_(row) || '',
    U: d.U || '',
    U2: d.U2 || '',
    U_num: d.n,
    genre: row['Identité de genre'] || row['Identité de Genre'] || row['Genre'] || '',
    genreInitiale: _genreInitFromRow_(row)
  };
}

function _toProperCase_(s) {
  s = String(s || '').toLowerCase();
  return s.replace(/\b\w/g, function (c) { return c.toUpperCase(); });
}


/** ======================== MAIL_SECTEURS (+ErrorCode) ======================== */
var MAIL_SECTORS_SHEET = 'MAIL_SECTEURS';
// On ajoute ErrorCode en DERNIER pour compat rétro
var MAIL_SECTORS_HEADER = [
  'SecteurId', 'Label', 'Umin', 'Umax', 'Genre',
  'To', 'Cc', 'ReplyTo', 'SubjectTpl', 'BodyTpl',
  'AttachIdsCSV', 'Active', 'ErrorCode'
];



function _getMailSectorsSheet_(ss) {
  // 1) ta feuille réelle
  var sh = ss.getSheetByName(MAIL_SECTORS_SHEET);
  // 2) compat (ancien nom)
  if (!sh) sh = ss.getSheetByName(MAIL_SECTORS_SHEET);
  // 3) sinon: on crée MINIMAL (sans valeurs par défaut)
  if (!sh) {
    sh = ss.insertSheet(MAIL_SECTORS_SHEET);
    sh.getRange(1, 1, 1, MAIL_SECTORS_HEADER.length).setValues([MAIL_SECTORS_HEADER]);
  }
  // upgrade doux: s'il manque "ErrorCode", on l’ajoute à la fin (sans réordonner)
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  if (hdr.indexOf('ErrorCode') < 0) {
    sh.insertColumnAfter(lastCol);
    sh.getRange(1, lastCol + 1).setValue('ErrorCode');
  }
  return sh;
}

/** Lecture robuste par NOMS de colonnes (compat v0.8 sans ErrorCode) */
function _loadMailSectors_(ss) {
  var sh = _getMailSectorsSheet_(ss);
  var lastRow = sh.getLastRow(); if (lastRow < 2) return [];
  var lastCol = sh.getLastColumn();
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(String);
  var idx = {}; headers.forEach(function (h, i) { idx[h] = i; });
  function val(r, k) { var i = idx[k]; return (i == null ? '' : values[r][i]); }

  var out = [];
  for (var r = 1; r < values.length; r++) {
    out.push({
      id: String(val(r, 'SecteurId') || '').trim(),
      label: String(val(r, 'Label') || '').trim(),
      Umin: Number(val(r, 'Umin') || ''),
      Umax: Number(val(r, 'Umax') || ''),
      genre: (String(val(r, 'Genre') || '*').trim().toUpperCase() || '*'),
      to: String(val(r, 'To') || '').trim(),
      cc: String(val(r, 'Cc') || '').trim(),
      replyTo: String(val(r, 'ReplyTo') || '').trim(),
      subj: String(val(r, 'SubjectTpl') || '').trim(),
      body: String(val(r, 'BodyTpl') || '').trim(),
      attachCsv: String(val(r, 'AttachIdsCSV') || '').trim(),
      active: String(val(r, 'Active')).toString().toLowerCase() !== 'false',
      errorCode: String(val(r, 'ErrorCode') || '').trim()
    });
  }
  return out.filter(function (s) { return s.active; }).sort(function (a, b) { return (a.Umin || 0) - (b.Umin || 0); });
}

/** match secteur avec filtre optionnel sur ErrorCode */
function _matchSectorForType_(sectors, payload, type) {
  var U_num = payload.U_num, g = (payload.genreInitiale || '').toUpperCase() || '*';
  var isNew = (type === 'INSCRIPTION_NEW');
  for (var i = 0; i < sectors.length; i++) {
    var s = sectors[i]; if (!s.active) continue;
    var okU = (U_num >= (s.Umin || 0)) && (U_num <= (s.Umax || 0));
    var okG = (s.genre === '*' || !s.genre) ? true : (s.genre === g || (s.genre === 'X' && (g === 'X' || g === '')));
    var okErr = isNew ? (!s.errorCode) : (String(s.errorCode || '') === String(type || ''));
    if (okU && okG && okErr) return s;
  }
  return null;
}


/* ============================ Worker ============================ */

/**
 * Dépile MAIL_OUTBOX et envoie les courriels (consomme le snapshot écrit à l’enqueue).
 * En fin de run, envoie 3 résumés CSV (INSCRIPTION_NEW uniquement) : U4-U8, U9-U12, U13-U18.
 */
function sendPendingOutbox(seasonId) {
  var ss = getSeasonSpreadsheet_(seasonId);

  // --- Params
  var dry = String(readParam_(ss, 'DRY_RUN') || 'FALSE').toUpperCase() === 'TRUE';
  var redirect = String(readParam_(ss, 'DRY_REDIRECT_EMAIL') || '').trim(); // utilisé SEULEMENT en DRY
  var fromName = String(readParam_(ss, 'MAIL_FROM') || 'Robot Courriels').trim();
  var useGmailApp = String(readParam_(ss, 'MAIL_USE_GMAILAPP') || 'TRUE').toUpperCase() === 'TRUE';
  var batchMax = parseInt(String(readParam_(ss, 'MAIL_BATCH_MAX') || '200'), 10) || 200;

  var sh = upgradeMailOutboxForDisplay_(ss);
  var idx = getHeadersIndex_(sh, sh.getLastColumn());
  var last = sh.getLastRow();
  if (last < 2) return { processed: 0, sent: 0, errors: 0, summaries: false };

  var w = sh.getLastColumn();
  var values = sh.getRange(2, 1, last - 1, w).getValues();
  function col(name) { return (idx[name] || 0) - 1; }

  var cType = col('Type'), cStatus = col('Status'), cTo = col('To'), cCc = col('Cc'), cReplyTo = col('ReplyTo');
  var cSubj = col('Sujet'), cBody = col('Corps'), cAtt = col('Attachments'), cSent = col('SentAt'), cErr = col('Error');
  var cSec = col('SecteurId');
  var cPasseport = idx['Passeport'] ? (idx['Passeport'] - 1) : -1;
  var cNomComplet = idx['NomComplet'] ? (idx['NomComplet'] - 1) : -1;
  var cFrais = idx['Frais'] ? (idx['Frais'] - 1) : -1;

  // --- MAIL_SECTEURS (pour déterminer la "band")
  var sectors = (readSheetAsObjects_(ss.getId(), 'MAIL_SECTEURS').rows || []).map(function (s) {
    return {
      SecteurId: String(s['SecteurId'] || '').trim(),
      Label: String(s['Label'] || '').trim(),
      Umin: parseInt(String(s['Umin'] || '').replace(/[^\d]/g, ''), 10) || 0,
      Umax: parseInt(String(s['Umax'] || '').replace(/[^\d]/g, ''), 10) || 0,
      Genre: String(s['Genre'] || '*').trim().toUpperCase() || '*'
    };
  });
  var sectorById = {};
  for (var i = 0; i < sectors.length; i++) sectorById[sectors[i].SecteurId] = sectors[i];

  function bandOfSecteur_(sid) {
    var s = sectorById[String(sid || '').trim()];
    if (!s) return '';
    var n1 = s.Umin, n2 = s.Umax;
    if (n1 >= 4 && n2 <= 8) return 'U4-U8';
    if (n1 >= 9 && n2 <= 12) return 'U9-U12';
    if (n1 >= 13 && n2 <= 18) return 'U13-U18';
    if (n2 <= 8) return 'U4-U8';
    if (n2 <= 12) return 'U9-U12';
    if (n2 <= 18) return 'U13-U18';
    return '';
  }

  function _normalizeEmailsCsv_(csv) {
    return String(csv || '')
      .split(/[;,]/).map(function (s) { return s.trim(); })
      .filter(Boolean)
      .filter(function (e) { return !/noreply|no-reply|invalid|example/.test(e.toLowerCase()); })
      .join(', ');
  }

  // --- Stat counters + collecte des envois INSCRIPTION_NEW de CE run
  var now = new Date();
  var sent = 0, errs = 0, proc = 0;
  var sentNewIns = []; // { p8, nomc, frais, secteurId }

  for (var r = 0; r < values.length && proc < batchMax; r++) {
    var row = values[r];
    if (String(row[cStatus] || '').trim().toLowerCase() !== 'pending') continue;
    proc++;

    var type = String(row[cType] || '').trim();
    var to = _normalizeEmailsCsv_(row[cTo] || '');      // <-- To provient de l'OUTBOX (déjà résolu en amont)
    var cc = _normalizeEmailsCsv_(row[cCc] || '');
    var replyTo = cReplyTo >= 0 ? String(row[cReplyTo] || '').trim() : '';
    var subj = String(row[cSubj] || '');
    var body = String(row[cBody] || '');
    var attachCsv = String(row[cAtt] || '');

    // --- Pas de destinataire → erreur
    if (!to) {
      row[cStatus] = 'error';
      row[cErr] = 'NO_RECIPIENT';
      values[r] = row;
      errs++;
      continue;
    }

    // --- Redirect uniquement en DRY
    var finalTo = to, finalCc = cc, finalSubj = subj, finalBodyHtml = body, finalReplyTo = replyTo;
    if (dry && redirect) {
      finalCc = '';
      finalSubj = '[DRY→' + redirect + '] ' + subj;
      var dbg = '<div style="font:12px monospace;color:#555">[DRY_REDIRECT] original To=' + _rg_csvEsc_(to) + '; Cc=' + _rg_csvEsc_(cc) + '</div>';
      finalBodyHtml = dbg + '\n' + body;
      finalTo = redirect;
    }

    // Attachments
    var blobs = _attachmentsFromCsv_(attachCsv);

    try {
      if (!dry) {
        if (useGmailApp) {
          var pt = (typeof _stripHtml_ === 'function') ? _stripHtml_(finalBodyHtml) : String(finalBodyHtml || '').replace(/<[^>]+>/g, ' ');
          GmailApp.sendEmail(finalTo, finalSubj, pt, {
            htmlBody: finalBodyHtml,
            cc: finalCc || undefined,
            name: fromName,
            replyTo: finalReplyTo || undefined,
            attachments: (blobs && blobs.length) ? blobs : undefined
          });
        } else {
          MailApp.sendEmail({
            to: finalTo, subject: finalSubj, htmlBody: finalBodyHtml,
            cc: finalCc || undefined, name: fromName,
            replyTo: finalReplyTo || undefined,
            attachments: (blobs && blobs.length) ? blobs : undefined
          });
        }
        // PROD → sent
        row[cStatus] = 'sent';
        row[cSent] = now;
        row[cErr] = '';
        sent++;
      } else {
        // DRY → pas d’envoi réel; on marque 'dry'
        row[cStatus] = 'dry';
        row[cSent] = now;     // timestamp utile pour audit
        row[cErr] = '';
      }

      // Collecte les INSCRIPTION_NEW réellement traités dans CE run
      if (String(type || '').toUpperCase() === 'INSCRIPTION_NEW') {
        var p8 = cPasseport >= 0 ? String(row[cPasseport] || '').trim() : '';
        var secteurId = cSec >= 0 ? String(row[cSec] || '').trim() : '';
        var nomc = cNomComplet >= 0 ? String(row[cNomComplet] || '').trim() : '';
        var frais = cFrais >= 0 ? String(row[cFrais] || '').trim() : '';
        sentNewIns.push({ p8: p8, nomc: nomc, frais: frais, secteurId: secteurId });
      }

    } catch (e) {
      row[cStatus] = 'error';
      row[cErr] = String(e);
      errs++;
    }

    values[r] = row;
  }

  // Flush
  if (proc > 0) {
    sh.getRange(2, 1, values.length, w).setValues(values);
  }

  // --- Résumés CSV (V2.0) basés sur les mails effectivement traités (INSCRIPTION_NEW)
// --- Résumés CSV (V2.0) basés sur les mails effectivement traités (INSCRIPTION_NEW)
var summaries = false;
(function __sendSummariesV2__() {
  var dry2      = String(readParam_(ss, 'DRY_RUN') || '').toUpperCase() === 'TRUE';
  var redirect2 = String(readParam_(ss, 'DRY_REDIRECT_EMAIL') || '').trim();
  var allowSend = !dry2 || (!!redirect2);

  var candidates = Array.isArray(sentNewIns) ? sentNewIns : [];
  summaries = candidates.length > 0;

  if (!candidates.length) {
    try { appendImportLog_(ss, 'MAIL_SUMMARIES_NONE', JSON.stringify({ reason:'no_sent_new', dry: dry2 })); } catch(_) {}
    return;
  }
  if (!allowSend) {
    try { appendImportLog_(ss, 'MAIL_SUMMARIES_SKIPPED', JSON.stringify({ reason:'dry_no_redirect', count: candidates.length })); } catch(_) {}
    return;
  }

  try {
    // Helpers
    function tz()     { return Session.getScriptTimeZone() || 'America/Toronto'; }
    function fmtDate(d){ return Utilities.formatDate(d, tz(), 'yyyy-MM-dd HH:mm'); }
    function to8(x)  { return String(x||'').replace(/\D/g,'').slice(-8).padStart(8,'0'); }

    // JOUEURS pour enrichir (noms + articles)
    var J      = readSheetAsObjects_(ss.getId(), SHEETS.JOUEURS).rows || [];
    var JByP8  = {};
    J.forEach(function(row){
      var p8j = to8(row['Passeport #'] || row['Passeport'] || '');
      if (p8j) JByP8[p8j] = row;
    });

    // Fonction de mappage SecteurId -> bande
    function inferBandFromSecteurId_(secteurId) {
      var band = bandOfSecteur_(secteurId);
      if (band) return band;
      var s = String(secteurId||'').toUpperCase();
      if (/(U4|U5|U6|U7|U8)/.test(s))  return 'U4-U8';
      if (/(U9|U10|U11|U12)/.test(s)) return 'U9-U12';
      if (/(U13|U14|U15|U16|U17|U18)/.test(s)) return 'U13-U18';
      return '';
    }

    // Regrouper les nouveaux inscrits par bande
    var byBand = { 'U4-U8': [], 'U9-U12': [], 'U13-U18': [] };
    candidates.forEach(function(it) {
      var p8 = to8(it.p8 || '');
      var jr = p8 ? (JByP8[p8] || {}) : {};
      var secteurId = it.secteurId || jr.SecteurId || '';
      var band = inferBandFromSecteurId_(secteurId);
      if (!band || !byBand[band]) return;

      var prenom = String(jr['Prénom'] || jr['Prenom'] || '');
      var nom    = String(jr['Nom'] || '');
      if ((!prenom || !nom) && it.nomc) {
        var parts = String(it.nomc).trim().split(/\s+/);
        if (parts.length > 1) {
          prenom = prenom || parts[0];
          nom    = nom    || parts.slice(1).join(' ');
        } else {
          nom = nom || it.nomc;
        }
      }
      var mainFee = String(it.frais || '');
      var autres  = extractOtherProducts_(jr, mainFee);

      byBand[band].push({
        p8:    p8,
        nom:   nom   || '',
        prenom:prenom|| '',
        frais: mainFee || '',
        autres: autres || ''
      });
    });

    // Modèles de sujet et de corps
    var subjTpl = readParam_(ss, 'MAIL_TEMPLATE_SUMMARY_SUBJECT') || 'Nouveaux inscrits — {{band}} — {{date}}';
    var bodyTpl = readParam_(ss, 'MAIL_TEMPLATE_SUMMARY_BODY')    || 'Bonjour,<br>Veuillez trouver en pièce jointe la liste des nouveaux inscrits {{band}} ({{count}} membres).';

    // Paramètres tardifs
    var lateDateU9U12   = readParam_(ss, 'LATE_DATE_U9U12')   || '';
    var lateDateU13U18  = readParam_(ss, 'LATE_DATE_U13U18')  || '';
    var lateSheetU9U12  = readParam_(ss, 'LATE_SHEET_U9U12')  || '';
    var lateSheetU13U18 = readParam_(ss, 'LATE_SHEET_U13U18') || '';

    function parseDate_(s){
      if (!s) return null;
      var p = s.split(/[-/]/);
      return new Date(p[0], p[1]-1, p[2]);
    }
    var today = new Date();
    var dU9   = parseDate_(lateDateU9U12);
    var dU13  = parseDate_(lateDateU13U18);

    // Écriture dans les feuilles tardives
    try {
      if ((!dU9 || today >= dU9) && lateSheetU9U12 && byBand['U9-U12'].length) {
        appendImportLog_(ss,'LATE_REG_ROWS',JSON.stringify({band:'U9-U12', count: byBand['U9-U12'].length}));
        writeLateRegistrations_(lateSheetU9U12, 'U9-U12', byBand['U9-U12']);
      }
      if ((!dU13 || today >= dU13) && lateSheetU13U18 && byBand['U13-U18'].length) {
        appendImportLog_(ss,'LATE_REG_ROWS',JSON.stringify({band:'U13-U18', count: byBand['U13-U18'].length}));
        writeLateRegistrations_(lateSheetU13U18, 'U13-U18', byBand['U13-U18']);
      }
    } catch(e) {
      appendImportLog_(ss, 'LATE_REG_ERROR', String(e));
    }

    // Préparer des liens vers les feuilles tardives
    var linkU9  = lateSheetU9U12  ? 'https://docs.google.com/spreadsheets/d/' + lateSheetU9U12  + '/edit' : '';
    var linkU13 = lateSheetU13U18 ? 'https://docs.google.com/spreadsheets/d/' + lateSheetU13U18 + '/edit' : '';

    // Fonction d’envoi des résumés
    function sendSummaryFor_(bandKey, toKey, ccKey) {
      var list = byBand[bandKey] || [];
      if (!list.length) return;

      var toS = String(readParam_(ss, toKey) || '').trim();
      var ccS = String(readParam_(ss, ccKey) || '').trim();
      var toFinal = (!dry2 ? toS : (redirect2 || ''));
      var ccFinal = (!dry2 ? ccS : '');
      if (!toFinal) return;

      var csvRows = [['Passeport #','Nom','Prénom','Nom du frais','Autres articles']];
      list.forEach(function(L){
        csvRows.push([L.p8||'', L.nom||'', L.prenom||'', L.frais||'', L.autres||'']);
      });
      var csv = csvRows.map(function(arr){ return arr.map(_rg_csvEsc_).join(','); }).join('\n');
      var filename = 'nouveaux_'+bandKey+'_'+fmtDate(new Date()).replace(/[^\d\-]/g,'')+'.csv';
      var blob = Utilities.newBlob(csv, 'text/csv', filename);

      var basePayload = { band: bandKey, date: fmtDate(new Date()), count: list.length };
      var subj = renderTemplate_(subjTpl, basePayload);
      var body = renderTemplate_(bodyTpl, basePayload);

      var link = (bandKey === 'U9-U12') ? linkU9 : (bandKey === 'U13-U18' ? linkU13 : '');
      if (link) {
        body += '<br><br><b>Consultez la feuille des inscriptions tardives&nbsp;:</b> '
              + '<a href="'+link+'" target="_blank">'+bandKey+'</a>';
      }
      if (dry2 && redirect2) subj = '[DRY] ' + subj;

      if (useGmailApp) {
        var pt = (typeof _stripHtml_ === 'function') ? _stripHtml_(body)
                 : String(body||'').replace(/<[^>]+>/g,' ').replace(/\s{2,}/g,' ').trim();
        GmailApp.sendEmail(toFinal, subj, pt, { htmlBody: body, cc: (ccFinal||undefined), name: fromName, attachments: [blob] });
      } else {
        MailApp.sendEmail({ to: toFinal, subject: subj, htmlBody: body, name: fromName, cc: (ccFinal||undefined), attachments: [blob] });
      }
    }

    // Envoi des résumés pour chaque secteur
    sendSummaryFor_('U4-U8',  'MAIL_TO_SUMMARY_U4U8',  'MAIL_CC_SUMMARY_U4U8');
    sendSummaryFor_('U9-U12', 'MAIL_TO_SUMMARY_U9U12', 'MAIL_CC_SUMMARY_U9U12');
    sendSummaryFor_('U13-U18','MAIL_TO_SUMMARY_U13U18','MAIL_CC_SUMMARY_U13U18');

    // Log final des envois
    try {
      var c1 = (byBand['U4-U8']   || []).length,
          c2 = (byBand['U9-U12']  || []).length,
          c3 = (byBand['U13-U18'] || []).length;
      appendImportLog_(ss, 'MAIL_SUMMARIES_SENT', JSON.stringify({ U4U8:c1, U9U12:c2, U13U18:c3, dry: dry2, redirected: !!redirect2 }));
    } catch(_) {}
  } catch (err) {
    try { appendImportLog_(ss, 'MAIL_SUMMARIES_ERROR', String(err)); } catch(__){}
  }
})();

  return { processed: proc, sent: sent, errors: errs, summaries: summaries };
}

/**
 * Extrait une liste lisible des autres articles/frais achetés par le membre,
 * en excluant le "frais principal" (mainFee). Accepte des schémas variés:
 * - colonnes nommées "Articles", "Autres articles", "Nom du frais", "Frais 1..n",
 *   "Article 1..n", "Produit 1..n", etc.
 * Retourne une chaîne (ex: "Chandail | Short | Bas").
 *
 * @param {Object} jr      Ligne JOUEURS (objet clé→valeur)
 * @param {string} mainFee Nom du frais principal (à exclure si présent)
 * @return {string}
 */
function extractOtherProducts_(jr, mainFee) {
  if (!jr || typeof jr !== 'object') return '';

  var main = String(mainFee || '').trim().toLowerCase();
  var out = [];
  var seen = Object.create(null);

  function pushUnique(s) {
    var v = String(s || '').trim();
    if (!v) return;
    var key = v.toLowerCase();
    if (main && key === main) return;             // exclure le frais principal
    if (seen[key]) return;
    seen[key] = true;
    out.push(v);
  }

  function splitMulti(s) {
    return String(s || '')
      .split(/[;,|]/g)
      .map(function(t){ return t.trim(); })
      .filter(Boolean);
  }

  // 1) Champs agrégés courants
  ['Autres articles', 'Articles', 'Autres', 'Autres produits'].forEach(function(k){
    if (jr.hasOwnProperty(k)) {
      splitMulti(jr[k]).forEach(pushUnique);
    }
  });

  // 2) Champs “Nom du frais” (peut contenir 1 élément; on l’exclura si == main)
  ['Nom du frais', 'Frais', 'Produit', 'Article'].forEach(function(k){
    if (jr.hasOwnProperty(k)) {
      splitMulti(jr[k]).forEach(pushUnique);
    }
  });

  // 3) Séries indexées (Article 1..n, Frais 1..n, Produit 1..n)
  Object.keys(jr).forEach(function(k){
    if (/(?:^|\s)(article|frais|produit)\s*\d+$/i.test(k)) {
      pushUnique(jr[k]);
    }
  });

  // 4) Fallback “large filet” : toute colonne contenant article|frais|produit
  // (utile si le schéma varie; garde ça en dernier pour éviter le bruit)
  Object.keys(jr).forEach(function(k){
    if (/article|frais|produit/i.test(k) &&
        !/courriel|email|e\-?mail|naissance|date|équipe|passport|passeport/i.test(k)) {
      splitMulti(jr[k]).forEach(pushUnique);
    }
  });

  return out.join(' | ');
}




/** === OUTBOX helpers (corrigés) ===
 * Nouvelles colonnes officielles: SecteurId, ReplyTo
 * - Création: utilise l’ordre ci-dessous.
 * - Migration: ajoute en fin de feuille toute colonne manquante.
 */
function getMailOutboxHeaders_() {
  // Colonnes “techniques” (avant) + ajout SecteurId et ReplyTo
  return [
    'Type',       // p.ex. INSCRIPTION_NEW ou code d’erreur
    'To',
    'Cc',
    'ReplyTo',    // nouveau
    'Sujet',
    'Corps',
    'Attachments',
    'KeyHash',
    'Status',     // pending|processing|sent|error
    'CreatedAt',
    'SentAt',
    'Error',
    'SecteurId',  // ID exact du secteur matché dans MAIL_SECTEURS
  ];
}
function ensureMailOutbox_(ss) {
  var headers = getMailOutboxHeaders_();
  var sh = ss.getSheetByName(SHEETS.MAIL_OUTBOX);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.MAIL_OUTBOX);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  // Si la feuille existe mais est vide → pose l’entête complète
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  // Migration douce: ajoute les colonnes manquantes en fin
  var first = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  var have = {}; first.forEach(function (h) { have[h] = true; });
  var toAdd = headers.filter(function (h) { return !have[h]; });

  if (toAdd.length) {
    sh.insertColumnsAfter(sh.getLastColumn() || 1, toAdd.length);
    var newHeader = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    // Écrit les noms des nouvelles colonnes en fin
    for (var i = 0; i < toAdd.length; i++) {
      sh.getRange(1, newHeader.length - toAdd.length + i + 1).setValue(toAdd[i]);
    }
  }
  return sh;
}

/** Ajoute (si besoin) des colonnes lisibles à droite: Passeport, NomComplet, Frais (inchangé) */
function upgradeMailOutboxForDisplay_(ss) {
  var sh = ensureMailOutbox_(ss); // crée/migre si besoin
  var firstRow = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
  var have = {}; firstRow.forEach(function (h) { have[String(h || '')] = true; });
  var add = [];
  ['Passeport', 'NomComplet', 'Frais'].forEach(function (h) {
    if (!have[h]) add.push(h);
  });
  if (!add.length) return sh;

  // append les nouvelles colonnes et inscrit l’entête
  sh.insertColumnsAfter(sh.getLastColumn() || 1, add.length);
  var hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  for (var i = 0; i < add.length; i++) {
    sh.getRange(1, hdr.length - add.length + i + 1).setValue(add[i]);
  }
  return sh;
}


/**
 * Ajoute N lignes dans MAIL_OUTBOX (idempotence gérée par l'appelant).
 * Optionnel: widthOpt pour forcer la largeur d'écriture (ex.: headers.length du call-site).
 */
function enqueueOutboxRows_(ssId, rows, widthOpt) {
  if (!rows || !rows.length) return 0;

  var ss = ensureSpreadsheet_(ssId);
  var sh = upgradeMailOutboxForDisplay_(ss);

  // Largeur d'écriture: priorité au widthOpt (si fourni), sinon largeur réelle de la feuille
  var W = Number(widthOpt || sh.getLastColumn() || 0);
  if (!W) return 0; // rien à écrire si pas d'entêtes

  // Normaliser chaque ligne à W colonnes (NE PAS tronquer ce que l'appelant nous a donné)
  var toWrite = rows.map(function (r) {
    var out = new Array(W);
    for (var i = 0; i < W; i++) out[i] = (r && i < r.length) ? r[i] : '';
    return out;
  });

  var start = sh.getLastRow() + 1;
  // Pas besoin d'insérer des lignes manuellement; setValues étend la feuille si nécessaire
  sh.getRange(start, 1, toWrite.length, W).setValues(toWrite);

  return toWrite.length;
}



// utils.js (ou email.js, même fichier que tes enqueues)
function _normFold_(s) { s = String(s == null ? '' : s); try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return s.toUpperCase().trim(); }

function isExcludedMember_(ss, row) {
  var map = readSheetAsObjects_(ss.getId(), SHEETS.MAPPINGS);
  var maps = (map && map.rows) ? map.rows : [];
  var hay = _normFold_((row['Catégorie'] || row['Categorie'] || '') + ' ' + (row['Nom du frais'] || row['Frais'] || row['Produit'] || ''));
  for (var i = 0; i < maps.length; i++) {
    var m = maps[i];
    if (String(m['Type'] || '').toLowerCase() !== 'member') continue;
    var excl = String(m['Exclude'] || '').toLowerCase() === 'true';
    if (!excl) continue;
    var ali = _normFold_(m['AliasContains'] || m['Alias'] || '');
    if (!ali) continue;
    if (hay.indexOf(ali) !== -1) return true;
  }
  return false;
}



// --- Utilitaires JOUEURS ---
function _firstNonEmpty_() {
  for (var i = 0; i < arguments.length; i++) {
    var v = String(arguments[i] || '').trim(); if (v) return v;
  }
  return '';
}
function _uNumberFromJ_(row, ss) {
  // 1) Priorité: Age numérique (déjà dans JOUEURS)
  var a = String(row.Age || row['Âge'] || '').trim();
  var age = a ? parseInt(a, 10) : 0;
  if (age) return age;

  // 2) Fallback: Date de naissance + année de saison
  var dob = row.DateNaissance || row['Date de naissance'] || '';
  if (dob && typeof readParam_ === 'function' && typeof parseSeasonYear_ === 'function' && typeof _extractBirthYearLoose_ === 'function') {
    var seasonY = parseSeasonYear_(readParam_(ss, 'SEASON_LABEL') || '') || (new Date()).getFullYear();
    var by = _extractBirthYearLoose_(dob);
    var a2 = by ? (seasonY - by) : 0;
    if (a2) return a2;
  }

  // 3) Dernier recours: U / AgeBracket
  //    → si c’est une plage “U4-U8”, on prend la BORNE HAUTE (8), pas la basse.
  var u = String(row.U || row.U2 || row.AgeBracket || '').toUpperCase();
  var m = u.match(/U\s*0?(\d+)(?:\s*-\s*U?\s*0?(\d+))?/);
  if (m) {
    var lo = parseInt(m[1], 10);
    var hi = m[2] ? parseInt(m[2], 10) : lo;
    return hi;
  }

  return 0;
}


function _normalizeP8_(p) {
  return (p == null || p === '') ? '' : String(p).replace(/\D/g, '').padStart(8, '0');
}

// Famille logique: “WELCOME/INSCRIPTION” vs “ERROR/VALIDATION”
function _typeFamily_(t) {
  var T = String(t || '').toUpperCase();
  if (T.indexOf('ERROR') > -1 || T.indexOf('VALIDATION') > -1) return 'ERR';
  if (T.indexOf('WELCOME') > -1 || T.indexOf('INSCRIPTION') > -1) return 'WELCOME';
  return T;
}

// Clé naturelle pour dédup:
// - WELCOME:      W|<p8>|<saison>
// - ERROR/VALID:  E|<p8>|<saison>|<code>
function _naturalOutboxKey_(row, saisonFallback) {
  var fam = _typeFamily_(row['Type'] || row.Type);
  var p8 = _normalizeP8_(row['Passeport #'] || row['Passeport'] || '');
  var saison = row['Saison'] || saisonFallback || '';
  if (!p8 || !saison) return '';
  if (fam === 'WELCOME') {
    return 'W|' + p8 + '|' + saison;
  }
  if (fam === 'ERR') {
    var code = String(row['ErrorCode'] || row['Code'] || row['ErreurCode'] || row['SecteurId'] || row['SectorId'] || '').toUpperCase();
    return code ? ('E|' + p8 + '|' + saison + '|' + code) : '';
  }
  return '';
}

// Lit MAIL_OUTBOX et prépare les sets de dédup (clé naturelle + keyhash legacy)
function _loadOutboxDedupSets_(ss, saison) {
  var out = readSheetAsObjects_(ss.getId(), 'MAIL_OUTBOX');
  var sentNat = new Set();
  var legacyHash = new Set();
  (out.rows || []).forEach(function (r) {
    var st = String(r['Status'] || r['Statut'] || '').toLowerCase();
    if (st === 'sent' || st === 'envoye' || st === 'queued' || st === 'enqueued') {
      var nk = _naturalOutboxKey_(r, r['Saison'] || saison);
      if (nk) sentNat.add(nk);
    }
    if (r['KeyHash']) legacyHash.add(String(r['KeyHash']));
  });
  return { sentNat: sentNat, legacyHash: legacyHash };
}


function _genreInit_(g) { return (String(g || '').toUpperCase().charAt(0) || ''); }

/**
 * Enfile des courriels "INSCRIPTION_NEW" à partir de JOUEURS, en gelant le Secteur et ses templates.
 * - Source unique: JOUEURS (aucun fallback INSCRIPTIONS).
 * - N’enfile QUE si un secteur actif (MAIL_SECTEURS) match (Umin/Umax + Genre {F|M|X|*}).
 * - Écrit un snapshot secteur dans MAIL_OUTBOX: SecteurId, To, Cc, ReplyTo, SubjectTpl, BodyTpl, AttachIdsCSV.
 * - Zéro recalcul au worker: tout est prêt dans la ligne OUTBOX.
 *
 * @param {string} seasonSheetId
 * @param {Array<string>|Set<string>=} passportsOpt  // (optionnel) filtre sur un sous-ensemble de passeports (p8)
 * @return {{enqueued:number, skipped_no_sector:number, skipped_invalid:number, dup_skipped:number}}
 */
function enqueueWelcomeFromJoueursFast_(seasonSheetId, passportsOpt) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var saisonLbl = readParam_(ss, 'SEASON_LABEL') || '';

  // --- Types d’email + param ON/OFF élite
  var TYPE_REG = 'INSCRIPTION_NEW';
  var TYPE_ELITE = 'INSCRIPTION_ELITE_NEW';
  var eliteEnqueueEnabled = (function () {
    var v = String(readParam_(ss, 'MAIL_ELITE_WELCOME_ENABLED') || 'FALSE').trim().toUpperCase();
    return (v === 'TRUE' || v === '1' || v === 'YES' || v === 'OUI');
  })();

  // --- 1) OUTBOX + headers (forcer l’ordre final souhaité)
  var shOut = upgradeMailOutboxForDisplay_(ss);
  var desiredHeaders = [
    'Type', 'To', 'Cc', 'Sujet', 'Corps', 'Attachments', 'KeyHash', 'Status', 'CreatedAt', 'SentAt', 'Error',
    'Passeport', 'NomComplet', 'Frais', 'SecteurId', 'ReplyTo'
  ];

  // récupère headers actuels (si vide, crée la ligne d’entêtes)
  var lastCol = Math.max(shOut.getLastColumn(), desiredHeaders.length);
  var headers = (lastCol > 0) ? shOut.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(String) : [];
  if (!headers || !headers.length || headers.every(function (h) { return !String(h || '').trim(); })) {
    shOut.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    headers = desiredHeaders.slice();
    lastCol = desiredHeaders.length;
  }

  // assure présence de toutes les colonnes, dans l’ordre exact demandé
  (function ensureHeadersOrder_() {
    // map actuel
    var map = {}; for (var i = 0; i < headers.length; i++) map[String(headers[i] || '').trim()] = i;
    // étend si nécessaire
    if (headers.length < desiredHeaders.length) {
      shOut.insertColumnsAfter(headers.length, desiredHeaders.length - headers.length);
      headers.length = desiredHeaders.length;
    }
    // écrit l’ordre exact
    shOut.getRange(1, 1, 1, desiredHeaders.length).setValues([desiredHeaders]);
    headers = desiredHeaders.slice();
  })();

  // index rapide
  var idx = {}; for (var i = 0; i < headers.length; i++) idx[headers[i]] = i;

  // raccourci d’accès
  function col(name) { return (name in idx) ? idx[name] : -1; }

  var cType = col('Type');
  var cTo = col('To');
  var cCc = col('Cc');
  var cSujet = col('Sujet');
  var cCorps = col('Corps');
  var cAttach = col('Attachments');
  var cKeyHash = col('KeyHash');
  var cStatus = col('Status');
  var cCreatedAt = col('CreatedAt');
  var cSentAt = col('SentAt');
  var cError = col('Error');
  var cPasseport = col('Passeport');
  var cNomComplet = col('NomComplet');
  var cFrais = col('Frais');
  var cSecteurId = col('SecteurId');
  var cReplyTo = col('ReplyTo');

  // --- 2) JOUEURS (map) + filtre optionnel
  var joueursRows = (readSheetAsObjects_(ss.getId(), SHEETS.JOUEURS).rows || []);
  var filterSet = (function toSet_(x) {
    if (!x) return null;
    if (x instanceof Set) return x;
    if (Array.isArray(x)) return new Set(x.map(function (v) { return String(v || '').trim(); }));
    var out = new Set(); try { Object.keys(x).forEach(function (k) { if (x[k]) out.add(String(k).trim()); }); } catch (e) { }
    return out;
  })(passportsOpt);

  function _normP8_(p) { return (p == null || p === '') ? '' : String(p).replace(/\D/g, '').padStart(8, '0'); }

  var joueursByP = {};
  for (var r = 0; r < joueursRows.length; r++) {
    var row = joueursRows[r];
    var p = _normP8_(row['Passeport #'] || row['Passeport'] || row['PS'] || '');
    if (!p) continue;
    if (filterSet && !filterSet.has(p)) continue;
    joueursByP[p] = row;
  }
  var allP = Object.keys(joueursByP);

  // --- 3) MAIL_SECTEURS actifs (pour To/Cc/Subject/Body)
  var sectorsObj = readSheetAsObjects_(ss.getId(), 'MAIL_SECTEURS');
  var sectors = (sectorsObj.rows || []).filter(function (s) {
    var active = String(s['Active'] || '').trim().toUpperCase();
    return (active === 'TRUE' || active === '1' || active === 'YES' || active === 'OUI');
  }).map(function (s) {
    return {
      SecteurId: String(s['SecteurId'] || s['SecteurID'] || '').trim(),
      Label: String(s['Label'] || '').trim(),
      Umin: parseInt(String(s['Umin'] || '').replace(/[^\d]/g, ''), 10) || 0,
      Umax: parseInt(String(s['Umax'] || '').replace(/[^\d]/g, ''), 10) || 0,
      Genre: String(s['Genre'] || '*').trim().toUpperCase() || '*',
      To: String(s['To'] || '').trim(),
      Cc: String(s['Cc'] || '').trim(),
      ReplyTo: String(s['ReplyTo'] || '').trim(),
      SubjectTpl: String(s['SubjectTpl'] || '').trim(),
      BodyTpl: String(s['BodyTpl'] || '').trim(),
      AttachIdsCSV: String(s['AttachIdsCSV'] || '').trim(),
      ErrorCode: String(s['ErrorCode'] || '').trim()
    };
  });

  var secById = {}; sectors.forEach(function (s) { if (s.SecteurId) secById[s.SecteurId] = s; });

  // --- 4) Dédup OUTBOX (Type|KeyHash) + dédup naturel W|TYPE|p8|saison
  var existing = {};
  var sentNatWelcome = new Set(); // W|TYPE|p8|saison
  var last = shOut.getLastRow();
  if (last >= 2) {
    var w2 = shOut.getLastColumn();
    var data = shOut.getRange(2, 1, last - 1, w2).getDisplayValues();
    var H = shOut.getRange(1, 1, 1, w2).getDisplayValues()[0].map(String);
    var colType = H.indexOf('Type');
    var colKH = H.indexOf('KeyHash');
    var colPassTxt = H.indexOf('Passeport');
    var colStatus = H.indexOf('Status');
    var colCreated = H.indexOf('CreatedAt');

    if (colType >= 0 && colKH >= 0) {
      for (var i = 0; i < data.length; i++) {
        var t = String(data[i][colType] || '').trim();
        var k = String(data[i][colKH] || '').trim();
        if (t && k) existing[t + '|' + k] = 1;
      }
    }

    var y = (function (lbl) { var m = String(lbl || '').match(/(\d{4})/); return m ? (+m[1]) : (new Date()).getFullYear(); })(saisonLbl);
    for (var j = 0; j < data.length; j++) {
      var tR = (colType >= 0) ? String(data[j][colType] || '').toUpperCase() : '';
      var st = (colStatus >= 0) ? String(data[j][colStatus] || '').toLowerCase() : '';
      var pT = (colPassTxt >= 0) ? String(data[j][colPassTxt] || '') : '';
      var cA = (colCreated >= 0) ? new Date(data[j][colCreated]) : null;
      if (!/WELCOME|INSCRIPTION/.test(tR)) continue;
      if (!pT) continue;
      if (!(st === 'sent' || st === 'queued' || st === 'pending')) continue;
      if (!(cA && !isNaN(+cA) && cA.getFullYear() === y)) continue;
      var p8x = _normP8_(pT);
      if (p8x && tR) sentNatWelcome.add('W|' + tR + '|' + p8x + '|' + saisonLbl);
    }
  }

  // --- helpers
  function _containsEliteKeyword_(s) {
    var x; try { x = String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase(); }
    catch (e) { x = String(s || '').toLowerCase(); }
    return /(?:^|[^a-z0-9])d1\+(?=$|[^a-z0-9])|(?:^|[^a-z0-9])cfp(?=$|[^a-z0-9])|(?:^|[^a-z0-9])ldp(?=$|[^a-z0-9])|ligue\s*2|ligue\s*3/.test(x);
  }
  function _makeKeyHash_(p8, saison, type) {
    try {
      var s = JSON.stringify({ p: p8, s: saison || '', t: (type || '') });
      var dig = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
      return Utilities.base64Encode(dig);
    } catch (e) { return (p8 + '|' + (saison || '') + '|' + (type || '')); }
  }
  function _genreInit_(g) {
    var s = String(g || '').trim().toUpperCase();
    if (s.startsWith('M')) return 'M';
    if (s.startsWith('F')) return 'F';
    return 'X';
  }
  function _resolveToFromJoueur_(jr, sec) {
    var sTo = String((sec && (sec.To || sec.to)) || '').trim();
    if (sTo) return sTo;
    var csv = String((jr && jr.Courriels) || '').trim();
    if (!csv) {
      var c1 = (jr && jr.CourrielPrimaire) || '';
      var p1 = (jr && jr.CourrielParent1) || '';
      var p2 = (jr && jr.CourrielParent2) || '';
      csv = [c1, p1, p2].filter(Boolean).join('; ');
    }
    csv = csv.replace(/[|]/g, ';').replace(/\s*;\s*/g, '; ').replace(/^[;,\s]+|[;,\s]+$/g, '');
    return csv;
  }
  function _programBandFromU_(U_num) {
    if (U_num >= 4 && U_num <= 8) return 'U4-U8';
    if (U_num >= 9 && U_num <= 12) return 'U9-U12';
    if (U_num >= 13 && U_num <= 18) return 'U13-U18';
    return '';
  }
  function _bandFromText_(s) {
    var t = String(s || '').toUpperCase();
    var m = t.match(/U\s*0?(\d{1,2})(?:\s*-\s*U?\s*0?(\d{1,2}))?/);
    if (!m) return '';
    var lo = parseInt(m[1], 10);
    var hi = m[2] ? parseInt(m[2], 10) : lo;
    if (hi < lo) { var tmp = lo; lo = hi; hi = tmp; }
    if (lo >= 4 && hi <= 8) return 'U4-U8';
    if (lo >= 9 && hi <= 12) return 'U9-U12';
    if (lo >= 13 && hi <= 18) return 'U13-U18';
    if (lo <= 8 && hi <= 12) return 'U9-U12';
    if (lo <= 12 && hi >= 13) return 'U13-U18';
    return '';
  }
  function _matchSectorForType_(U_num, G, TYPE) {
    var cand = sectors.filter(function (s) {
      var okU = (s.Umin <= U_num && U_num <= s.Umax);
      var okG = (s.Genre === G || s.Genre === 'X' || s.Genre === '*');
      if (!okU || !okG) return false;
      var err = String(s.ErrorCode || '').trim();
      if (TYPE === TYPE_ELITE) return (err === TYPE_ELITE);
      return (err === '');
    });
    if (!cand.length) return null;
    function scoreOf_(s) {
      var gScore = (s.Genre === G) ? 3 : (s.Genre === 'X' ? 2 : 1);
      var width = Math.max(1, s.Umax - s.Umin + 1);
      return (gScore * 1000) - width;
    }
    cand.sort(function (a, b) { return scoreOf_(b) - scoreOf_(a); });
    return cand[0];
  }

  // --- 4bis) LEDGER : INSCRIPTIONS actives (élites vs normales)
  var led = readSheetAsObjects_(ss.getId(), SHEETS.ACHATS_LEDGER);
  var rowsL = (led.rows || []).filter(function (L) {
    if (String(L['Saison'] || '') !== saisonLbl) return false;
    if ((Number(L['Status']) || 0) !== 1) return false;
    if ((Number(L['isIgnored']) || 0) === 1) return false;
    return true;
  });
  var hasInscByP = {};
  var hasEliteInscByP = {};
  (rowsL || []).forEach(function (L) {
    var t = String(L['Type'] || '').toUpperCase();
    if (t !== 'INSCRIPTION') return;
    var name = String(L['NomFrais'] || L['Nom du frais'] || L['Frais'] || L['Produit'] || '');
    var p8 = _normP8_(L['Passeport #'] || L['Passeport'] || L['PS'] || '');
    if (!p8) return;
    hasInscByP[p8] = true;
    if (_containsEliteKeyword_(name)) hasEliteInscByP[p8] = true;
  });

  // --- 5) Build OUTBOX (1 pass)
  var now = new Date();
  var rowsOut = [];
  var stats = { enqueued: 0, skipped_no_sector: 0, skipped_invalid: 0, dup_skipped: 0, skipped_elite_off: 0, skipped_no_insc: 0, skipped_adapte: 0 };

  for (var i = 0; i < allP.length; i++) {
    var p8 = allP[i];
    var row = joueursByP[p8];
    if (!row) { stats.skipped_invalid++; continue; }

    if (!hasInscByP[p8]) { stats.skipped_no_insc++; continue; }

    // ⛔️ soccer adapté → skip (pas de welcome par secteur d’âge)
    var isAdapteFlag = String(row.isAdapte || row['isAdapte'] || row['Adapté'] || row['Adapte'] || row['Adaptée'] || '').trim().toUpperCase();
    if (isAdapteFlag === '1' || isAdapteFlag === 'TRUE') { stats.skipped_adapte++; continue; }

    var U_num = Number(_uNumberFromJ_(row, ss) || 0);
    if (!U_num) { stats.skipped_invalid++; continue; }
    var G = _genreInit_(row['Genre'] || row['Identité de genre'] || row['Identité de Genre'] || '');

    var isElite = !!hasEliteInscByP[p8];
    var TYPE = isElite ? TYPE_ELITE : TYPE_REG;
    if (isElite && !eliteEnqueueEnabled) { stats.skipped_elite_off++; continue; }

    // Secteur (modèles / adresses)
    var sec = _matchSectorForType_(U_num, G, TYPE);
    if (!sec) {
      // fallback explicite: SecteurId déjà présent sur JOUEURS
      var sidJ = String(row['SecteurId'] || row['SecteurID'] || '').trim();
      var s2 = sidJ ? secById[sidJ] : null;
      if (s2) {
        var err = String(s2.ErrorCode || '').trim();
        var okT = (TYPE === TYPE_ELITE) ? (err === TYPE_ELITE) : (err === '');
        if (okT) sec = s2;
      }
    }
    if (!sec) { stats.skipped_no_sector++; continue; }

    // === SecteurId OUTBOX — lire JOUEURS, sinon fallback sur le secteur réel ===
    // 1) SecteurId OUTBOX dérivé (avec fallback sec.SecteurId) – OK
    var abRaw = String(row.AgeBracket || '').trim();
    var pbRaw = String(row.ProgramBand || row['Program Band'] || row['Program_Band'] || '').trim();
    var secteurIdOutbox =
      _bandFromText_(abRaw)
      || _bandFromText_(pbRaw)
      || _programBandFromU_(Number(row.U || U_num || 0))
      || (sec && sec.SecteurId)
      || '';

    // 2) Destinataires
    var toCsv = _resolveToFromJoueur_(row, sec);
    if (!toCsv) { stats.skipped_invalid++; continue; }

    // 3) Données pour le templating (AVANT de calculer subject/body)
    var data = buildDataFromRow_(row) || {};
    data.U_num = U_num;
    if (!data.U) data.U = 'U' + U_num;
    if (!data.U2) data.U2 = 'U' + (U_num < 10 ? ('0' + U_num) : U_num);
    if (!data.genreInitiale) data.genreInitiale = G;

    // 4) Templating (subject/body) — ICI, AVANT d’écrire 'line[...]'
    var subject = renderTemplate_(sec.SubjectTpl || '', data);
    var body = renderTemplate_(sec.BodyTpl || '', data);

    // 5) Key/dédup
    var keyHash = _makeKeyHash_(p8, data.saison || saisonLbl, TYPE);
    var natKey = 'W|' + TYPE + '|' + p8 + '|' + saisonLbl;
    if (sentNatWelcome.has(natKey)) { stats.dup_skipped++; continue; }
    if (existing[TYPE + '|' + keyHash]) { stats.dup_skipped++; continue; }

    // 6) Écriture EN 1 PASSE
    var line = new Array(headers.length); for (var z = 0; z < line.length; z++) line[z] = '';
    if (cType >= 0) line[cType] = TYPE;
    if (cTo >= 0) line[cTo] = toCsv;
    if (cCc >= 0) line[cCc] = sec.Cc || '';
    if (cSujet >= 0) line[cSujet] = subject;     // ← subject est maintenant défini
    if (cCorps >= 0) line[cCorps] = body;
    if (cAttach >= 0) line[cAttach] = sec.AttachIdsCSV || '';
    if (cKeyHash >= 0) line[cKeyHash] = keyHash;
    if (cStatus >= 0) line[cStatus] = 'pending';
    if (cCreatedAt >= 0) line[cCreatedAt] = now;
    if (cSentAt >= 0) line[cSentAt] = '';
    if (cError >= 0) line[cError] = '';

    if (cPasseport >= 0) line[cPasseport] = (typeof normalizePassportToText8_ === 'function') ? normalizePassportToText8_(p8) : String(p8 || '');
    if (cNomComplet >= 0) line[cNomComplet] = data.nomcomplet || (((row['Prénom'] || row['Prenom'] || '') + ' ' + (row['Nom'] || '')).trim());
    if (cFrais >= 0) line[cFrais] = sec.Label || '';
    if (cSecteurId >= 0) line[cSecteurId] = secteurIdOutbox;
    if (cReplyTo >= 0) line[cReplyTo] = sec.ReplyTo || '';

    rowsOut.push(line);
    existing[TYPE + '|' + keyHash] = 1;
    sentNatWelcome.add(natKey);

  }

  // --- 6) Écriture (1 pass)
  if (rowsOut.length) {
    enqueueOutboxRows_(ss.getId(), rowsOut, headers.length);
  }


  return {
    enqueued: rowsOut.length,
    skipped_no_sector: stats.skipped_no_sector,
    skipped_invalid: stats.skipped_invalid,
    dup_skipped: stats.dup_skipped,
    skipped_elite_off: stats.skipped_elite_off,
    skipped_no_insc: stats.skipped_no_insc,
    skipped_adapte: stats.skipped_adapte
  };
}




// === FAST: enfile tous les e-mails de VALIDATION en 1 passe (à partir de ERREURS) ===
function enqueueValidationMailsFromErreursFast_(seasonId, codeFilterOpt) {
  var ss = SpreadsheetApp.openById(seasonId);
  var saisonLbl = readParam_(ss, 'SEASON_LABEL') || '';
  var shO = upgradeMailOutboxForDisplay_(ss);
  var hdr = getMailOutboxHeaders_();
  var idx = getHeadersIndex_(shO, hdr.length);

  // ---- Certaines erreurs doivent quand même partir même sans entrée INSCRIPTIONS
  var ALLOW_ERR_WITHOUT_FINAL = {
    'U13U18_CAMP_SEUL': true
    // ajoute ici d'autres codes qui n'ont pas forcément de "final" (si besoin)
  };

  // ---- Dédup OUTBOX existant
  var existing = {};               // legacy Type||KeyHash
  var sentNatErr = new Set();      // naturel E|p8|saison|code
  var last = shO.getLastRow();
  if (last >= 2) {
    var V = shO.getRange(2, 1, last - 1, hdr.length).getDisplayValues();
    var H = shO.getRange(1, 1, 1, hdr.length).getDisplayValues()[0].map(String);
    var iT = H.indexOf('Type');
    var iKH = H.indexOf('KeyHash');
    var iP = H.indexOf('Passeport');
    var iSt = H.indexOf('Status');
    var iCr = H.indexOf('CreatedAt');

    // legacy
    if (iT >= 0 && iKH >= 0) {
      for (var i = 0; i < V.length; i++) {
        var t = String(V[i][iT] || '').trim();
        var kh = String(V[i][iKH] || '').trim();
        if (t && kh) existing[t + '||' + kh] = true;
      }
    }

    // naturel (même saison, via CreatedAt.year)
    var y = (function (lbl) { var m = String(lbl || '').match(/(\d{4})/); return m ? (+m[1]) : (new Date()).getFullYear(); })(saisonLbl);
    function to8(x) { return String(x || '').replace(/\D/g, '').slice(-8).padStart(8, '0'); }
    for (var j = 0; j < V.length; j++) {
      var tR = (iT >= 0) ? String(V[j][iT] || '').trim().toUpperCase() : '';
      var st = (iSt >= 0) ? String(V[j][iSt] || '').trim().toLowerCase() : '';
      var pT = (iP >= 0) ? String(V[j][iP] || '') : '';
      var cA = (iCr >= 0) ? new Date(V[j][iCr]) : null;

      if (!(st === 'sent' || st === 'queued' || st === 'pending')) continue;
      if (!(cA && !isNaN(+cA) && cA.getFullYear() === y)) continue;

      var p8 = to8(pT);
      if (!p8) continue;

      // Ici, Type en outbox == code d’erreur (U7_8_SANS_2E_SEANCE, etc.)
      if (tR && p8) sentNatErr.add('E|' + p8 + '|' + saisonLbl + '|' + tR);
    }
  }

  // ---- Secteurs (ceux avec ErrorCode)
  var sectors = _loadMailSectors_(ss).filter(function (s) { return !!String(s.errorCode || '').trim(); });
  if (!sectors.length) return { scanned: 0, matched: 0, deduped: 0, queued: 0 };

  var secByCode = {};
  sectors.forEach(function (s) { secByCode[String(s.errorCode).trim()] = s; });

  // ---- Helpers
  function to8(x) { return String(x || '').replace(/\D/g, '').slice(-8).padStart(8, '0'); }
  function _firstNonEmpty_() { for (var i = 0; i < arguments.length; i++) { var v = String(arguments[i] || '').trim(); if (v) return v; } return ''; }

  // ---- Index INSCRIPTIONS (p8||Saison → {row,kh})
  var finals = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS).rows || [];
  var keyColsCsv = readParam_(ss, PARAM_KEYS.KEY_COLS) || 'Passeport #,Saison';
  var keyCols = keyColsCsv.split(',').map(function (x) { return x.trim(); });

  var finalByPass8Season = {};
  for (var iF = 0; iF < finals.length; iF++) {
    var R = finals[iF] || {};
    var pass = String(R['Passeport #'] || R['Passeport'] || '').trim(); if (!pass) continue;
    var p8 = to8(pass);
    var saz = String(R['Saison'] || '').trim();

    var keyStr = keyCols.map(function (kc) { return R[kc] == null ? '' : String(R[kc]); }).join('||');
    var kh = Utilities.base64EncodeWebSafe(Utilities.newBlob(keyStr).getBytes());

    finalByPass8Season[p8 + '||' + saz] = { row: R, kh: kh };
  }

  // ---- Index JOUEURS (p8 → row)
  var J = readSheetAsObjects_(ss.getId(), SHEETS.JOUEURS).rows || [];
  var JByP8 = {};
  for (var jx = 0; jx < J.length; jx++) {
    var JJ = J[jx] || {};
    var p8j = to8(JJ['Passeport #'] || JJ['Passeport'] || '');
    if (p8j) JByP8[p8j] = JJ;
  }

  // ---- Lecture ERREURS
  var E = readSheetAsObjects_(ss.getId(), SHEETS.ERREURS).rows || [];
  if (!E.length) return { scanned: 0, matched: 0, deduped: 0, queued: 0 };

  var outRows = [];
  var denorm = [];
  var scanned = 0, matched = 0, deduped = 0, queued = 0;
  var only = codeFilterOpt ? String(codeFilterOpt).trim() : '';

  // --- Collecteur pour log propre (un seul append)
  var fallbackHits = [];
  var byCode = {};

  for (var r = 0; r < E.length; r++) {
    var e = E[r]; scanned++;

    var code = String(e['Code'] || e['ErrorCode'] || e['Type'] || '').trim();
    if (!code) continue;
    if (only && code !== only) continue;

    var sec = secByCode[code];
    if (!sec) continue; // IMPORTANT: il faut un secteur configuré pour ce code

    var passRaw = e['Passeport #'] || e['Passeport'] || '';
    var saison = String(e['Saison'] || '').trim();
    if (!passRaw || !saison) continue;

    var p8 = to8(passRaw);
    if (!p8) continue;

    var fk = p8 + '||' + saison;
    var k = finalByPass8Season[fk] || { row: null, kh: null };

    // ---- Fallback si pas d'INSCRIPTIONS pour ce code (ex.: U13U18_CAMP_SEUL)
    if ((!k.kh || !k.row) && ALLOW_ERR_WITHOUT_FINAL[code] === true) {
      var keyStrFallback = [p8, saison, code, 'NO_INSCR'].join('||');
      var khFallback = Utilities.base64EncodeWebSafe(Utilities.newBlob(keyStrFallback).getBytes());
      k = { row: {}, kh: khFallback };
      fallbackHits.push({ code: code, passport8: p8, saison: saison });
      byCode[code] = (byCode[code] || 0) + 1;
    }

    // Si on n'a toujours pas de kh (et pas autorisé en fallback), on skip
    if ((!k.kh || !k.row) && !ALLOW_ERR_WITHOUT_FINAL[code]) continue;

    if (typeof _isCoachMemberSafe_ === 'function' && _isCoachMemberSafe_(ss, k.row)) continue;
    if (typeof isExcludedMember_ === 'function' && isExcludedMember_(ss, k.row)) continue;

    // dédup naturel (E|p8|saisonCourante|code) — borne à la saison courante
    var natKey = 'E|' + p8 + '|' + saisonLbl + '|' + code;
    if (sentNatErr.has(natKey)) { deduped++; continue; }

    // dédup legacy (Type||KeyHash), ici Type == code
    var existingKey = code + '||' + k.kh;
    if (existing[existingKey]) { deduped++; continue; }

    var jr = JByP8[p8] || {};
    var email = _firstNonEmpty_(jr.Courriels, jr.CourrielPrimaire, jr.CourrielParent1, jr.CourrielParent2);

    var payload = {};
    try {
      payload = buildDataFromRow_(k.row) || {};
      if (jr) {
        if (!payload.prenom) payload.prenom = jr['Prénom'] || jr['Prenom'] || '';
        if (!payload.nomcomplet) payload.nomcomplet = (jr.NomComplet || (((jr['Prénom'] || jr['Prenom'] || '') + ' ' + (jr.Nom || '')).trim()));
      }
      payload.error_code = code;
      payload.error_label = String(e['Message'] || '').trim() || code;
      payload.error_details = String(e['Contexte'] || '').trim();
    } catch (_) { }

    var subj = sec.subj ? renderTemplate_(sec.subj, payload) : '';
    var body = sec.body ? renderTemplate_(sec.body, payload) : '';

    var arr = new Array(hdr.length).fill('');
    function put(col, val) { var i = idx[col]; if (i) arr[i - 1] = val; }

    put('Type', code);
    put('SecteurId', sec.id || sec.SecteurId || '');
    put('To', (sec.to || sec.To || email || ''));
    if (sec.cc) put('Cc', sec.cc);
    if (sec.replyTo) put('ReplyTo', sec.replyTo);
    if (subj) put('Sujet', subj);
    if (body) put('Corps', body);
    if (sec.attachCsv) put('Attachments', sec.attachCsv);
    put('KeyHash', k.kh);
    if (idx['NaturalKey']) put('NaturalKey', natKey);
    put('Status', 'pending');
    put('CreatedAt', new Date());
    put('Error', JSON.stringify({ label: payload.error_label || code, details: payload.error_details || '' }));

    outRows.push(arr);
    denorm.push({ row: k.row, jr: jr, err: e });
    existing[existingKey] = true;
    sentNatErr.add(natKey);
    matched++; queued++;
  }

  if (!outRows.length) {
    // log propre des fallbacks même s'il n'y a rien à écrire
    if (fallbackHits.length) {
      try {
        appendImportLog_(ss, 'MAIL_QUEUE_ERRORS_NO_FINAL_FALLBACK',
          JSON.stringify({ count: fallbackHits.length, byCode: byCode }));
      } catch (_) { }
    }
    return { scanned: scanned, matched: matched, deduped: deduped, queued: 0 };
  }

  // ---- Écriture batch
  var startRow = shO.getLastRow() + 1;
  enqueueOutboxRows_(ss.getId(), outRows);

  // ---- Backfill lisible (non-bloquant)
  try {
    var n = denorm.length;
    var pass = new Array(n), nomc = new Array(n), frais = new Array(n);
    for (var z = 0; z < n; z++) {
      var src = denorm[z] || {};
      var r0 = (src.row || {});
      var jr0 = (src.jr || {});
      var e0 = (src.err || {});
      var ptxt = (typeof normalizePassportToText8_ === 'function')
        ? normalizePassportToText8_(r0['Passeport #'] || r0['Passeport'] || e0['Passeport #'] || e0['Passeport'] || '')
        : String(r0['Passeport #'] || r0['Passeport'] || e0['Passeport #'] || e0['Passeport'] || '');
      pass[z] = [ptxt];
      nomc[z] = [(jr0.NomComplet || (((jr0['Prénom'] || jr0['Prenom'] || '') + ' ' + (jr0.Nom || '')).trim()))];
      frais[z] = [e0['Nom du frais'] || e0['Frais'] || e0['Produit'] || r0['Nom du frais'] || r0['Frais'] || r0['Produit'] || ''];
    }
    if (idx['Passeport']) shO.getRange(startRow, idx['Passeport'], n, 1).setValues(pass);
    if (idx['NomComplet']) shO.getRange(startRow, idx['NomComplet'], n, 1).setValues(nomc);
    if (idx['Frais']) shO.getRange(startRow, idx['Frais'], n, 1).setValues(frais);
  } catch (_) { /* non bloquant */ }

  // ---- Log unique des fallbacks, après écriture
  if (fallbackHits.length) {
    try {
      appendImportLog_(ss, 'MAIL_QUEUE_ERRORS_NO_FINAL_FALLBACK',
        JSON.stringify({ count: fallbackHits.length, byCode: byCode }));
    } catch (_) { }
  }

  return { scanned: scanned, matched: matched, deduped: deduped, queued: queued };
}

/**
 * Écrit / fusionne les inscriptions tardives dans la feuille cible.
 * Appelée par __sendSummariesV2__ :
 *   writeLateRegistrations_(spreadsheetId, bandKey, list)
 * where list = [{p8, nom, prenom, frais, autres}, ...]
 *
 * Sélection de l’onglet (dans le classeur targetSpreadsheetId), par priorité :
 *   1) readParam_(ss, 'LATE_SHEET_TAB_'+BAND)  // ex. 'LATE_SHEET_TAB_U9_U12'
 *   2) readParam_(ss, 'LATE_SHEET_TAB')
 *   3) 'Inscriptions' (par défaut)
 *
 * Upsert basé sur "Passeport #", normalisé à 8 chiffres.
 *
 * @param {string} targetSpreadsheetId
 * @param {'U4-U8'|'U9-U12'|'U13-U18'} bandKey
 * @param {Array<Object>} rows
 * @return {number} nb de lignes ajoutées
 */
function writeLateRegistrations_(targetSpreadsheetId, bandKey, rows) {
  if (!targetSpreadsheetId || !Array.isArray(rows) || !rows.length) return 0;

  var ss = SpreadsheetApp.openById(targetSpreadsheetId);
  var sheetName = (function() {
    try {
      // Permet d’avoir un onglet différent par bande (ex.: LATE_SHEET_TAB_U9_U12)
      var bandToken = String(bandKey || '')
        .replace(/-/g, '_')
        .replace(/\s/g, '_'); // 'U9-U12' => 'U9_U12'
      var perBand = (typeof readParam_ === 'function')
        ? (readParam_(ss, 'LATE_SHEET_TAB_' + bandToken) || '')
        : '';
      if (perBand) return String(perBand);

      var generic = (typeof readParam_ === 'function')
        ? (readParam_(ss, 'LATE_SHEET_TAB') || '')
        : '';
      return generic ? String(generic) : 'Inscriptions';
    } catch (_) {
      return 'Inscriptions';
    }
  })();

  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  // En-têtes cibles “compatibles assignation”
  // (Tu peux ajuster l’ordre, mais garde au minimum Passeport #, Prénom, Nom, Équipe, Courriel envoyé)
  var wantedHeaders = [
    'Passeport #',
    'Prénom',
    'Nom',
    'Courriel',
    'Date de naissance',
    'Équipe',
    'Courriel envoyé',
    'Nom du frais',
    'Autres articles'
  ];

  // Pose les en-têtes si feuille vierge
  var headers;
  if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) {
    headers = wantedHeaders.slice();
    sh.clear();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    // Si des colonnes manquent, on les append à droite (idempotent)
    var missing = wantedHeaders.filter(function(h){ return headers.indexOf(h) < 0; });
    if (missing.length) {
      var start = headers.length + 1;
      // Étend la largeur
      sh.insertColumnsAfter(headers.length, missing.length);
      // Écrit l’entête manquante
      sh.getRange(1, start, 1, missing.length).setValues([missing]);
      headers = headers.concat(missing);
    }
  }

  var keyHeader = 'Passeport #';
  var keyIdx = headers.indexOf(keyHeader);
  if (keyIdx < 0) throw new Error('Colonne clé "' + keyHeader + '" introuvable dans la feuille cible');

  // Index des passeports déjà présents
  var existing = {};
  var lastRow = sh.getLastRow();
  if (lastRow >= 2) {
    var keyCol = keyIdx + 1;
    var colVals = sh.getRange(2, keyCol, lastRow - 1, 1).getValues();
    for (var i = 0; i < colVals.length; i++) {
      var v = __toP8(String(colVals[i][0] || '').replace(/^'/,'').trim());
      if (v) existing[v] = true;
    }
  }

  // Mappe les objets {p8, nom, prenom, frais, autres} vers row alignée sur headers
  var rowsToAppend = [];
  for (var r = 0; r < rows.length; r++) {
    var it = rows[r] || {};
    var p8 = __toP8(it.p8 || '');
    if (!p8) continue; // pas de clé, on ignore
    if (existing[p8]) continue; // déjà là

    var out = new Array(headers.length);
    for (var c = 0; c < headers.length; c++) {
      var H = headers[c];

      if (H === 'Passeport #') {
        out[c] = p8;
      } else if (H === 'Prénom') {
        out[c] = it.prenom || '';
      } else if (H === 'Nom') {
        out[c] = it.nom || '';
      } else if (H === 'Nom du frais') {
        out[c] = it.frais || '';
      } else if (H === 'Autres articles') {
        out[c] = it.autres || '';
      } else if (H === 'Équipe') {
        out[c] = ''; // laissé vide pour l’assignation humaine + onEdit handler
      } else if (H === 'Courriel envoyé') {
        out[c] = ''; // horodaté plus tard par l’onEdit courriel
      } else {
        // Colonnes “confort” qu’on ne connaît pas (Courriel, Date de naissance, etc.)
        out[c] = '';
      }
    }

    rowsToAppend.push(out);
    existing[p8] = true;
  }

  if (rowsToAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  }

  // Optionnel: journaliser localement
  try {
    appendImportLog_(ss, 'LATE_REG_WRITE_OK', JSON.stringify({
      band: bandKey,
      added: rowsToAppend.length,
      total_input: rows.length,
      sheet: sheetName
    }));
  } catch (_) {}

  return rowsToAppend.length;

  /** Normalise un passeport en 8 chiffres (garde num., tronque à droite, pad à gauche). */
  function __toP8(x) {
    var s = String(x || '').replace(/\D/g, '');
    if (!s) return '';
    s = s.slice(-8);
    while (s.length < 8) s = '0' + s;
    return s;
  }
}

