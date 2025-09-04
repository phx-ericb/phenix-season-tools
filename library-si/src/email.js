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
  var SHEETS = { INSCRIPTIONS:'INSCRIPTIONS', MAIL_OUTBOX:'MAIL_OUTBOX', MAIL_LOG:'MAIL_LOG', PARAMS:'PARAMS' };
}
if (typeof PARAM_KEYS === 'undefined') {
  var PARAM_KEYS = {
    KEY_COLS:'KEY_COLS', DRY_RUN:'DRY_RUN', MAIL_FROM:'MAIL_FROM', MAIL_BATCH_MAX:'MAIL_BATCH_MAX',
    TO_FIELDS_INSCRIPTIONS:'TO_FIELDS_INSCRIPTIONS',
    MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT:'MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT',
    MAIL_TEMPLATE_INSCRIPTION_NEW_BODY:'MAIL_TEMPLATE_INSCRIPTION_NEW_BODY',
    MAIL_TO_SUMMARY_U4U8:'MAIL_TO_SUMMARY_U4U8', MAIL_CC_SUMMARY_U4U8:'MAIL_CC_SUMMARY_U4U8',
    MAIL_TO_SUMMARY_U9U12:'MAIL_TO_SUMMARY_U9U12', MAIL_CC_SUMMARY_U9U12:'MAIL_CC_SUMMARY_U9U12',
    MAIL_TO_SUMMARY_U13U18:'MAIL_TO_SUMMARY_U13U18', MAIL_CC_SUMMARY_U13U18:'MAIL_CC_SUMMARY_U13U18',
    MAIL_TEMPLATE_SUMMARY_SUBJECT:'MAIL_TEMPLATE_SUMMARY_SUBJECT',
    MAIL_TEMPLATE_SUMMARY_BODY:'MAIL_TEMPLATE_SUMMARY_BODY'
  };
}
if (typeof getSeasonSpreadsheet_ !== 'function') { function getSeasonSpreadsheet_(id){ if(!id)throw new Error('seasonSheetId manquant'); return SpreadsheetApp.openById(id); } }
if (typeof getSheetOrCreate_ !== 'function') {
  function getSheetOrCreate_(ss, name, header){
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); if (header && header.length) sh.getRange(1,1,1,header.length).setValues([header]); }
    else if (header && header.length && sh.getLastRow() === 0){ sh.getRange(1,1,1,header.length).setValues([header]); }
    return sh;
  }
}
if (typeof getHeadersIndex_ !== 'function') {
  function getHeadersIndex_(sh, width){ var headers=sh.getRange(1,1,1,width||sh.getLastColumn()).getValues()[0].map(String); var idx={}; headers.forEach(function(h,i){ idx[h]=i+1; }); return idx; }
}
if (typeof readParam_ !== 'function') {
  function readParam_(ss, key){
    var sh = ss.getSheetByName(SHEETS.PARAMS);
    if (sh) { var last=sh.getLastRow(); if (last>=1){ var data=sh.getRange(1,1,last,2).getValues(); for (var i=0;i<data.length;i++){ if ((data[i][0]+'').trim()===key) return (data[i][1]+'').trim(); } } }
    var props = PropertiesService.getDocumentProperties(); return (props.getProperty(key)||'').trim();
  }
}
if (typeof readSheetAsObjects_ !== 'function') {
  function readSheetAsObjects_(ssId, sheetName){
    var ss = SpreadsheetApp.openById(ssId); var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow()<1 || sh.getLastColumn()<1) return { sheet: sh||ss.insertSheet(sheetName), headers: [], rows: [] };
    var values = sh.getRange(1,1,sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    var headers = values[0].map(String), rows=[];
    for (var r=1;r<values.length;r++){ var o={}; for (var c=0;c<headers.length;c++) o[headers[c]]=values[r][c]; rows.push(o); }
    return { sheet: sh, headers: headers, rows: rows };
  }
}
if (typeof appendImportLog_ !== 'function') {
  function appendImportLog_(ss, action, details){ var sh=getSheetOrCreate_(ss,'IMPORT_LOG',['Horodatage','Action','Détails']); sh.appendRow([new Date(),action,details||'']); }
}
/* U/secteur fallbacks */
if (typeof deriveSectorFromRow_ !== 'function') {
  function parseSeasonYear_(s){ var m=(String(s||'').match(/(20\d{2})/)); return m?parseInt(m[1],10):(new Date()).getFullYear(); }
  function birthYearFromRow_(row){
    var y=row['Année de naissance']||row['Annee de naissance']||row['Annee']||''; if (y && /^\d{4}$/.test(String(y))) return parseInt(y,10);
    var dob=row['Date de naissance']||''; if (dob){ var s=String(dob), m=s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if(m) return parseInt(m[1],10); var m2=s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/); if(m2) return parseInt(m2[3],10); }
    return null;
  }
  function computeUForYear_(by, sy){ if(!by||!sy) return null; var u=sy-by; return (u>=4&&u<=21)?('U'+u):null; }
  function deriveUFromRow_(row){
    var cat=row['Catégorie']||row['Categorie']||''; if (cat && /^U\d{1,2}/i.test(cat)) return cat.toUpperCase().replace(/\s+/g,'');
    var U = computeUForYear_(birthYearFromRow_(row), parseSeasonYear_(row['Saison']||'')); return U||'';
  }
  function deriveSectorFromRow_(row){
    var U=deriveUFromRow_(row), n=parseInt(String(U).replace(/^U/i,''),10);
    if(!n||isNaN(n)) return 'AUTRES';
    if (n>=4 && n<=8) return 'U4-U8'; if (n>=9 && n<=12) return 'U9-U12'; if (n>=13 && n<=18) return 'U13-U18';
    return 'AUTRES';
  }
}
if (typeof collectEmailsFromRow_ !== 'function') {
  function norm_(s){ s=String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s.trim(); }
  function collectEmailsFromRow_(row, fieldsCsv){
    var fields=(fieldsCsv&&fieldsCsv.length)?fieldsCsv.split(',').map(function(x){return x.trim();}).filter(Boolean):['Courriel','Parent 1 - Courriel','Parent 2 - Courriel'];
    var set={}; fields.forEach(function(f){ var v=row[f]; if(!v) return; String(v).split(/[;,]/).forEach(function(e){ e=norm_(e); if(!e) return; set[e]=true; }); });
    return Object.keys(set).join(',');
  }
}


// ===== Coach detection shim (lib-safe) =====
// Utilise la version lib (rules.js) si dispo, sinon fallback autonome.
// Zéro SpreadsheetApp.getActive() — on passe "ss".

function _coachCsv_(ss){
  var csv = '';
  try { if (typeof readParam_ === 'function') csv = readParam_(ss, 'RETRO_COACH_FEES_CSV') || ''; } catch(_){}
  if (!csv) csv = 'Entraîneurs, Entraineurs, Entraîneur, Entraineur, Coach, Coaches';
  return csv;
}
function _normNoAccentsLower_(s){
  s = String(s == null ? '' : s).trim();
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g,''); } catch(_){}
  return s.toLowerCase();
}
function _isCoachFeeByNameSafe_(ss, name){
  // 1) Si la lib rules.js fournit isCoachFeeByName_, on l'utilise
  try { if (typeof isCoachFeeByName_ === 'function') return !!isCoachFeeByName_(ss, name); } catch(_){}
  // 2) Fallback par mots-clés
  var v = _normNoAccentsLower_(name);
  if (!v) return false;
  var toks = _coachCsv_(ss).split(',').map(_normNoAccentsLower_).filter(Boolean);
  if (toks.some(function(t){ return v === t || v.indexOf(t) >= 0; })) return true;
  // filet
  return /(entraineur|entra[îi]neur|coach)/i.test(String(name||''));
}
function _isCoachMemberSafe_(ss, row){
  var name = row ? (row['Nom du frais'] || row['Frais'] || row['Produit'] || '') : '';
  // 1) Si la lib rules.js fournit isCoachMember_, on l'utilise
  try { if (typeof isCoachMember_ === 'function') return !!isCoachMember_(ss, row); } catch(_){}
  // 2) Fallback
  return _isCoachFeeByNameSafe_(ss, name);
}



/* ======================== Helpers communs ======================== */

function _rg_csvEsc_(v){ v=v==null?'':String(v).replace(/"/g,'""'); return /[",\n;]/.test(v)?('"'+v+'"'):v; }

/** "id1, id2 ; id3" -> ['id1','id2','id3'] */
function _parseAttachIdsCsv_(csv) {
  if (!csv) return [];
  return String(csv)
    .split(/[,\s;]+/)
    .map(function(s){ return String(s||'').trim(); })
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
    } catch(e) {
      try { Logger.log('ATTACH_WARN ' + id + ': ' + e); } catch(_) {}
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
  for (var i=0;i<blobs.length;i++){
    var b = blobs[i];
    var key = (b.getName()||'') + '|' + (b.getBytes() ? b.getBytes().length : b.getDataAsString().length);
    if (!seen[key]) { seen[key]=1; unique.push(b); }
  }
  return unique;
}



function _normText(s){ s=String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s.trim(); }
function renderTemplate_(tpl, data){ tpl=String(tpl==null?'':tpl); return tpl.replace(/{{\s*([\w.]+)\s*}}/g,function(_,k){ var v=(data.hasOwnProperty(k)?data[k]:''); return (v==null?'':String(v)); }); }
// Helpers (ajoute-les si non présents dans le fichier)
function _genreInitFromRow_(row){
  var lbl = row['Identité de genre'] || row['Identité de Genre'] || row['Genre'] || '';
  var g = String(lbl||'').toUpperCase().trim();
  if (!g) return 'X';
  if (g[0] === 'M') return 'M';
  if (g[0] === 'F') return 'F';
  return 'X';
}
function _U_U2_FromRow_(row){
  var U = deriveUFromRow_(row) || '';                 // utilise computeUForYear_/birthYearFromRow_/parseSeasonYear_
  var n = parseInt(String(U).replace(/^U/i,''), 10);
  var U2 = (!isNaN(n) ? ('U' + (n < 10 ? ('0' + n) : n)) : '');
  return { U: U, U2: U2, n: (isNaN(n) ? '' : n) };
}
function buildDataFromRow_(row) {
  var d = _U_U2_FromRow_(row);

  var prenomRaw = row['Prénom']||row['Prenom']||'';
  var nomRaw    = row['Nom']||'';

  var prenomPC  = _toProperCase_(prenomRaw);
  var nomPC     = _toProperCase_(nomRaw);

  return {
    passeport:  row['Passeport #'] || '',
    nom:        nomPC,
    prenom:     prenomPC,
    nomcomplet: (prenomPC + ' ' + nomPC).trim(),
    saison:     row['Saison'] || '',
    frais:      row['Nom du frais'] || row['Frais'] || row['Produit'] || '',
    categorie:  row['Catégorie'] || row['Categorie'] || '',
    secteur:    deriveSectorFromRow_(row) || '',
    U:          d.U || '',
    U2:         d.U2 || '',
    U_num:      d.n,
    genre:      row['Identité de genre'] || row['Identité de Genre'] || row['Genre'] || '',
    genreInitiale: _genreInitFromRow_(row)
  };
}

function _toProperCase_(s) {
  s = String(s||'').toLowerCase();
  return s.replace(/\b\w/g, function(c){ return c.toUpperCase(); });
}


/** Renvoie true si la ligne INSCRIPTIONS matche un mapping member avec Exclude=TRUE */
function matchesMemberExcludeMapping_(ss, row){
  var mapObj = readSheetAsObjects_(ss.getId(), 'MAPPINGS');
  var maps = (mapObj && mapObj.rows) ? mapObj.rows : [];
  if (!maps.length) return false;

  function norm(s){ s = String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s.toUpperCase().trim(); }
  var txt = norm( (row['Catégorie']||row['Categorie']||'') + ' ' + (row['Nom du frais']||row['Frais']||row['Produit']||'') );

  // U / Genre de la ligne
  var U = deriveUFromRow_(row)||''; var m = String(U).match(/U\s*-?\s*(\d{1,2})/i);
  var uNum = m ? parseInt(m[1],10) : null;
  var g0 = (row['Identité de genre']||row['Identité de Genre']||row['Genre']||''); var G = (String(g0).trim().toUpperCase()[0]||'X');

  for (var i=0;i<maps.length;i++){
    var mp = maps[i];
    if (String(mp['Type']||'').toLowerCase() !== 'member') continue;

    // Exclude truthy ?
    var ex = String(mp['Exclude']||'').trim().toUpperCase();
    var isExcluded = (ex==='TRUE' || ex==='1' || ex==='YES' || ex==='OUI');
    if (!isExcluded) continue;

    // AliasContains
    var alias = norm(mp['AliasContains']||mp['Alias']||'');
    if (alias && txt.indexOf(alias) === -1) continue;

    // Umin/Umax
    var Umin = parseInt(mp['Umin']||'',10), Umax = parseInt(mp['Umax']||'',10);
    if (uNum!=null && !isNaN(Umin) && uNum < Umin) continue;
    if (uNum!=null && !isNaN(Umax) && uNum > Umax) continue;

    // Genre
    var mg = String(mp['Genre']||'*').trim().toUpperCase();
    if (mg && mg !== '*'){
      if (mg==='X'){ /* ok pour inconnu */ }
      else if (G !== mg) continue;
    }

    // -> toutes les conditions sont OK
    return true;
  }
  return false;
}

/* ======================== MAIL_SECTEURS ======================== */
/** ======================== MAIL_SECTEURS (+ErrorCode) ======================== */
var MAIL_SECTORS_SHEET = 'MAIL_SECTEURS';
// On ajoute ErrorCode en DERNIER pour compat rétro
var MAIL_SECTORS_HEADER = [
  'SecteurId','Label','Umin','Umax','Genre',
  'To','Cc','ReplyTo','SubjectTpl','BodyTpl',
  'AttachIdsCSV','Active','ErrorCode'
];

/** Crée/upgrade le header si besoin (ajoute ErrorCode si absent) */
function _ensureMailSectorsSheet_(ss){
  var sh = ss.getSheetByName(MAIL_SECTORS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(MAIL_SECTORS_SHEET);
    sh.getRange(1,1,1,MAIL_SECTORS_HEADER.length).setValues([MAIL_SECTORS_HEADER]);
    // quelques valeurs par défaut (confirmation, pas d'ErrorCode)
    var defaults = [
      ['U9','U9',9,9,'*','','','','Confirmation U9 – {{nomcomplet}}','Bonjour {{prenom}},<br>Bienvenue!','',true,''],
      ['U10','U10',10,10,'*','','','','Confirmation U10 – {{nomcomplet}}','Bonjour {{prenom}},<br>Bienvenue!','',true,''],
      ['U11','U11',11,11,'*','','','','Confirmation U11 – {{nomcomplet}}','Bonjour {{prenom}},<br>Bienvenue!','',true,''],
      ['U12','U12',12,12,'*','','','','Confirmation U12 – {{nomcomplet}}','Bonjour {{prenom}},<br>Bienvenue!','',true,'']
    ];
    sh.getRange(2,1,defaults.length, MAIL_SECTORS_HEADER.length).setValues(defaults);
    return sh;
  }
  // upgrade header si ErrorCode manquant
  var lastCol = sh.getLastColumn()||0;
  var hdr = sh.getRange(1,1,1, Math.max(lastCol, MAIL_SECTORS_HEADER.length)).getValues()[0].map(String);
  // Si le header exact n'est pas celui attendu → on réécrit la ligne d'entête complète
  var needUpgrade = MAIL_SECTORS_HEADER.some(function(h, i){ return (hdr[i]||'') !== h; });
  if (needUpgrade) {
    // Étend la feuille si nécessaire
    if (sh.getLastColumn() < MAIL_SECTORS_HEADER.length) {
      sh.insertColumnsAfter(sh.getLastColumn() || 1, MAIL_SECTORS_HEADER.length - (sh.getLastColumn()||1));
    }
    sh.getRange(1,1,1,MAIL_SECTORS_HEADER.length).setValues([MAIL_SECTORS_HEADER]);
  }
  return sh;
}

/** Lecture robuste par NOMS de colonnes (compat v0.8 sans ErrorCode) */
function _loadMailSectors_(ss){
  var sh = _ensureMailSectorsSheet_(ss);
  var lastRow = sh.getLastRow(); if (lastRow < 2) return [];
  var lastCol = sh.getLastColumn();
  var values = sh.getRange(1,1,lastRow, lastCol).getValues();
  var headers = values[0].map(String);
  var idx = {}; headers.forEach(function(h,i){ idx[h]=i; });

  function val(row, key){ var i=idx[key]; return (i==null? '' : values[row][i]); }

  var out = [];
  for (var r=1; r<values.length; r++){
    out.push({
      id: String(val(r,'SecteurId')||'').trim(),
      label: String(val(r,'Label')||'').trim(),
      Umin: Number(val(r,'Umin')||''),
      Umax: Number(val(r,'Umax')||''),
      genre: (String(val(r,'Genre')||'*').trim().toUpperCase()||'*'),
      to: String(val(r,'To')||'').trim(),
      cc: String(val(r,'Cc')||'').trim(),
      replyTo: String(val(r,'ReplyTo')||'').trim(),
      subj: String(val(r,'SubjectTpl')||'').trim(),
      body: String(val(r,'BodyTpl')||'').trim(),
      attachCsv: String(val(r,'AttachIdsCSV')||'').trim(),
      active: String(val(r,'Active')).toString().toLowerCase() !== 'false',
      errorCode: String(val(r,'ErrorCode')||'').trim() // '' = confirmations
    });
  }
  return out.filter(function(s){ return s.active; }).sort(function(a,b){ return (a.Umin||0)-(b.Umin||0); });
}

/** match secteur avec filtre optionnel sur ErrorCode */
function _matchSectorForType_(sectors, payload, type){
  var U_num = payload.U_num, g = (payload.genreInitiale||'').toUpperCase()||'*';
  var isNew = (type === 'INSCRIPTION_NEW');
  for (var i=0;i<sectors.length;i++){
    var s = sectors[i]; if (!s.active) continue;
    var okU = (U_num >= (s.Umin||0)) && (U_num <= (s.Umax||0));
    var okG = (s.genre === '*' || !s.genre) ? true : (s.genre === g || (s.genre==='X' && (g==='X'||g==='')));
    var okErr = isNew ? (!s.errorCode) : (String(s.errorCode||'') === String(type||''));
    if (okU && okG && okErr) return s;
  }
  return null;
}

function _attachmentsFromCsv_(idsCsv){
  var out = [];
  String(idsCsv||'').split(/[,\s;]+/).filter(Boolean).forEach(function(id){
    try{ out.push(DriveApp.getFileById(id).getBlob()); }catch(e){}
  });
  return out;
}

/* ============================ Worker ============================ */

/**
 * Envoie les emails en attente (MAIL_OUTBOX)
 * - Type = INSCRIPTION_NEW : applique secteur (si trouvé) pour Sujet/Body/Attachments/To/Cc/ReplyTo
 *   sinon fallback global (PARAMS).
 * - Résumés par secteur (legacy v0.7) conservés.
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
  var useGmailApp = (readParam_(ss, PARAM_KEYS.MAIL_USE_GMAILAPP) || 'FALSE').toUpperCase() === 'TRUE';

  var sectors = _loadMailSectors_(ss);

  var data = shOut.getRange(2, 1, last - 1, headers.length).getValues();
  var processed = 0;
  var processedNew = [];           // pour les résumés CSV
  var mailLogBuffer = [];          // <-- on accumule les lignes pour MAIL_LOG
  var sentRows = [];               // (optionnel) indices des lignes OUTBOX traitées

  function _normUpper_(s) {
    s = String(s == null ? '' : s);
    try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) {}
    return s.trim().toUpperCase();
  }
  function _composeText_(row) {
    return _normUpper_((row['Catégorie'] || row['Categorie'] || '') + ' ' + (row['Nom du frais'] || row['Frais'] || row['Produit'] || ''));
  }
  function _excludedReasonForMail_(ss, row, type) {
    if (type !== 'INSCRIPTION_NEW') return null;
    try {
      if (typeof matchesMemberExcludeMapping_ === 'function' && matchesMemberExcludeMapping_(ss, row)) return 'excluded_by_mappings';
    } catch(e){}
    try {
      if (typeof isAdapteMember_ === 'function' && isAdapteMember_(row)) return 'excluded_adapte_rule';
    } catch(e){}

   // ⛔ Coach : pas de confirmations
   try {
     if (_isCoachMemberSafe_(ss, row)) return 'excluded_coach_fee';
   } catch(e){}

    var txt = _composeText_(row);
    if (txt.indexOf('ADAPTE') !== -1) return 'excluded_adapte_text';
    if (txt.indexOf('ADULTE') !== -1) return 'excluded_adulte_text';
    return null;
  }

  for (var i = 0; i < data.length && processed < batchMax; i++) {
    var rowArr = data[i];
    var row = {}; headers.forEach(function (h, j) { row[h] = rowArr[j]; });

    if (String(row['Status']).toLowerCase() !== 'pending') continue;
    if (row['SentAt']) continue;

    var type = row['Type'];
    var keyHash = row['KeyHash'];

    var fRow = fetchFinalRowByKeyHash_(ss, keyHash) || {};
    var payload = buildDataFromRow_(fRow);

    // Backfill affichage (Passeport formaté + Nom/Frais) si vides
    try {
      var headersNow = shOut.getRange(1, 1, 1, shOut.getLastColumn()).getValues()[0].map(String);
      var idxCols = {}; headersNow.forEach(function (h, ii) { idxCols[h] = ii + 1; });
      function isEmptyCol(col) { var c = idxCols[col]; return c && !String(shOut.getRange(i+2, c).getValue() || '').trim(); }
      if (fRow && (isEmptyCol('Passeport') || isEmptyCol('NomComplet') || isEmptyCol('Frais'))) {
        if (idxCols['Passeport']) {
          var passText8 = (typeof normalizePassportToText8_ === 'function') ? normalizePassportToText8_(fRow['Passeport #']) : String(fRow['Passeport #']||'');
          shOut.getRange(i+2, idxCols['Passeport']).setValue(passText8);
        }
        if (idxCols['NomComplet']) {
          var nomc = (((fRow['Prénom']||fRow['Prenom']||'')+' '+(fRow['Nom']||'')).trim());
          shOut.getRange(i+2, idxCols['NomComplet']).setValue(nomc);
        }
        if (idxCols['Frais']) {
          var frais = fRow['Nom du frais']||fRow['Frais']||fRow['Produit']||'';
          shOut.getRange(i+2, idxCols['Frais']).setValue(frais);
        }
      }
    } catch(e){}

    // Exclusions confirmations
    var excludedReason = _excludedReasonForMail_(ss, fRow, type);
    if (excludedReason) {
      shOut.getRange(i + 2, idx['Status']).setValue('skipped_excluded');
      shOut.getRange(i + 2, idx['Error']).setValue(excludedReason);
      shOut.getRange(i + 2, idx['SentAt']).setValue(new Date());
      processed++; sentRows.push(i+2);
      var resultTag = dry ? 'DRY_RUN_SKIP' : 'SKIP_EXCLUDED';
      mailLogBuffer.push([type, (row['To']||''), '(skipped) '+(row['Sujet']||''), keyHash, new Date(), resultTag]);
      continue;
    }

// -- Résolution destinataires (secteur > finale > STAGING fallback)
var rcptDefault = { to:'', cc:'' };

// overrides secteur (si un secteur s'applique plus tard)
if (s && s.to) rcptDefault.to = s.to;
if (s && s.cc) rcptDefault.cc = s.cc;

// sinon: essaie sur la ligne finale
if (!rcptDefault.to) {
  var csv = readParam_(ss, PARAM_KEYS.TO_FIELDS_INSCRIPTIONS) || 'Courriel,Parent 1 - Courriel,Parent 2 - Courriel';
  var toFinal = collectEmailsFromRow_(fRow || {}, csv);
  if (toFinal) rcptDefault.to = toFinal;
}

// si encore vide ET type ≠ INSCRIPTION_NEW: secours via STAGING
if (!rcptDefault.to && type !== 'INSCRIPTION_NEW') {
  var sRow = fetchStagingRowByKeyHash_(ss, keyHash);
  if (sRow) {
    var csv2 = readParam_(ss, PARAM_KEYS.TO_FIELDS_INSCRIPTIONS) || 'Courriel,Parent 1 - Courriel, Parent 2 - Courriel';
    var toStage = collectEmailsFromRow_(sRow, csv2);
    if (toStage) rcptDefault.to = toStage;

    // bonus: en profite pour enrichir le payload (prenom/nom/U...) si vide
    try {
      var p2 = buildDataFromRow_(sRow);
      ['prenom','nom','nomcomplet','U','U2','U_num','categorie','secteur'].forEach(function(k){
        if (!payload[k] && p2[k]) payload[k] = p2[k];
      });
    } catch(_) {}
  }
}

// CC par défaut selon tes PARAMS « confirmation » (ou ajoute des PARAMS spécifiques si tu veux)
if (!rcptDefault.cc) rcptDefault.cc = readParam_(ss, PARAM_KEYS.MAIL_CC_NEW_INSCRIPTIONS) || '';


var subject = (row['Sujet'] || '').trim();
    var bodyHtml = (row['Corps'] || '').trim();
    var attachments = _attachmentsFromCsv_(row['Attachments']);

    var isNew = (type === 'INSCRIPTION_NEW');
    if (!isNew) {
      try {
        var err = JSON.parse(String(row['Error'] || '{}'));
        payload.error_code = type;
        payload.error_label = err.label || '';
        payload.error_details = err.details || '';
      } catch (e) {
        payload.error_code = type;
        payload.error_label = '';
        payload.error_details = '';
      }
    }

    var s = _matchSectorForType_(sectors, payload, type);
    if (type === 'INSCRIPTION_NEW' && !s) {
      shOut.getRange(i + 2, idx['Status']).setValue('skipped_no_sector');
      shOut.getRange(i + 2, idx['SentAt']).setValue(new Date());
      processed++; sentRows.push(i+2);
      mailLogBuffer.push([type, (row['To']||''), '(no_sector) '+(row['Sujet']||''), keyHash, new Date(), dry ? 'DRY_RUN_SKIP' : 'SKIP_NO_SECTOR']);
      continue;
    }

    if (!subject && s && s.subj) subject = renderTemplate_(s.subj, payload);
    if (!bodyHtml && s && s.body) bodyHtml = renderTemplate_(s.body, payload);

    if (isNew) {
      if (!subject) subject = renderTemplate_(readParam_(ss, PARAM_KEYS.MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT) || 'Bienvenue {{prenom}} – {{frais}}', payload);
      if (!bodyHtml) bodyHtml = renderTemplate_(readParam_(ss, PARAM_KEYS.MAIL_TEMPLATE_INSCRIPTION_NEW_BODY) || 'Bonjour {{prenom}},<br>Nous confirmons votre inscription à {{frais}} ({{saison}}).', payload);
    } else {
      if (!subject) subject = 'Validation requise – '+(payload.nomcomplet||'')+' ('+(payload.U||'')+')';
      if (!bodyHtml) bodyHtml = 'Bonjour '+(payload.prenom||'')+',<br><br>Nous avons remarqué un point à valider pour '+(payload.saison||'')+
        ' ('+(payload.U||'')+') : <b>'+(payload.error_label||type)+'</b><br><small>'+(payload.error_details||'')+'</small>';
    }

    var to = (row['To'] || '').trim() || (s && s.to) || rcptDefault.to || '';
    var cc = (row['Cc'] || '').trim() || (s && s.cc) || rcptDefault.cc || '';
    var replyTo = (s && s.replyTo) || undefined;

    if (s && s.attachCsv) attachments = attachments.concat(_attachmentsFromCsv_(s.attachCsv));
// === Resolve DRY behavior ===
var dryRun = (readParam_(ss, PARAM_KEYS.DRY_RUN) || 'FALSE').toUpperCase() === 'TRUE';
var redirect = readParam_(ss, 'DRY_REDIRECT_EMAIL'); // new param in PARAMS

// Build a safe recipient set for DRY
var resolvedTo = to;
var resolvedCc = cc;
var resolvedReplyTo = replyTo;
var resolvedSubject = subject;
var resolvedHtml = bodyHtml;

if (dryRun && redirect) {
  // Redirige tout vers l’adresse de test; garde l’original en mémo dans le corps
  resolvedSubject = '[DRY → ' + redirect + '] ' + subject;

  // Ajoute un encart debug au bas du HTML (les headers custom ne sont pas supportés)
  var debugBlock =
    '<hr style="margin:16px 0;border:0;border-top:1px solid #ddd">' +
    '<div style="font:12px/1.4 system-ui,Arial,sans-serif;color:#555">' +
      '<div><b>DRY RUN</b> : ce message a été redirigé.</div>' +
      '<div><b>Destinataires (original)</b> : ' + _rg_csvEsc_(to) + (cc ? (' | CC: ' + _rg_csvEsc_(cc)) : '') + '</div>' +
      (replyTo ? ('<div><b>Reply-To</b> : ' + _rg_csvEsc_(replyTo) + '</div>') : '') +
    '</div>';

  resolvedHtml = (bodyHtml || '') + debugBlock;

  // Forcer l’envoi vers l’adresse de test
  resolvedTo = redirect;
  resolvedCc = '';          // neutre en DRY
  resolvedReplyTo = '';     // neutre en DRY
}

if (dryRun && !redirect) { appendImportLog_(ss, 'MAIL_DRY_NO_REDIRECT', 'DRY_RUN sans DRY_REDIRECT_EMAIL: envoi bloqué'); return; }


// === Envoi (inchangé, mais avec les variables "resolved*") ===
if (useGmailApp) {
  GmailApp.sendEmail(resolvedTo, resolvedSubject, '', {
    htmlBody: resolvedHtml,
    cc: resolvedCc || undefined,
    replyTo: resolvedReplyTo || undefined,
    attachments: (attachments.length ? attachments : undefined),
    name: fromName
  });
} else {
  MailApp.sendEmail({
    to: resolvedTo,
    cc: resolvedCc || undefined,
    replyTo: resolvedReplyTo || undefined,
    subject: resolvedSubject,
    htmlBody: resolvedHtml,
    attachments: (attachments.length ? attachments : null),
    name: fromName
  });
}


    // marque envoyé (par ligne; on peut optimiser si les lignes sont contiguës)
    shOut.getRange(i + 2, idx['SentAt']).setValue(new Date());
    shOut.getRange(i + 2, idx['Status']).setValue(dry ? 'dry_run' : 'sent');
    processed++; sentRows.push(i+2);

    // log dans le buffer
    mailLogBuffer.push([type, to, subject, keyHash, new Date(), dry ? 'DRY_RUN' : (useGmailApp ? 'SENT_GMAILAPP' : 'SENT_MAILAPP')]);

    if (isNew) processedNew.push({ row: fRow, data: payload });
  }

  // Écrit MAIL_LOG en batch
  if (mailLogBuffer.length) {
    var shLog = getSheetOrCreate_(ss, SHEETS.MAIL_LOG, ['Type','To','Sujet','KeyHash','SentAt','Result']);
    var start = shLog.getLastRow() + 1;
    shLog.insertRowsAfter(shLog.getLastRow(), mailLogBuffer.length);
    shLog.getRange(start, 1, mailLogBuffer.length, 6).setValues(mailLogBuffer);
  }

  // Résumés CSV (confirmations)
  if (processedNew.length) {
    var groups = { 'U4-U8': [], 'U9-U12': [], 'U13-U18': [] };
    processedNew.forEach(function (x) {
      var sec = x.data.secteur || deriveSectorFromRow_(x.row);
      if (groups[sec]) groups[sec].push(x);
    });
    var tplSub = readParam_(ss, PARAM_KEYS.MAIL_TEMPLATE_SUMMARY_SUBJECT) || 'Nouveaux inscrits – {{secteur}} – {{date}}';
    var tplBody = readParam_(ss, PARAM_KEYS.MAIL_TEMPLATE_SUMMARY_BODY) || 'Bonjour,<br>Veuillez trouver la liste des nouveaux inscrits {{secteur}} en pièce jointe.<br><br>Bonne journée.';
    Object.keys(groups).forEach(function (sec) {
      var arr = groups[sec]; if (!arr.length) return;
      var toKey = sec === 'U4-U8' ? PARAM_KEYS.MAIL_TO_SUMMARY_U4U8 : (sec === 'U9-U12' ? PARAM_KEYS.MAIL_TO_SUMMARY_U9U12 : PARAM_KEYS.MAIL_TO_SUMMARY_U13U18);
      var ccKey = sec === 'U4-U8' ? PARAM_KEYS.MAIL_CC_SUMMARY_U4U8 : (sec === 'U9-U12' ? PARAM_KEYS.MAIL_CC_SUMMARY_U9U12 : PARAM_KEYS.MAIL_CC_SUMMARY_U13U18);
      var to = readParam_(ss, toKey) || '', cc = readParam_(ss, ccKey) || '';

      var headersCSV = ['Passeport', 'NomComplet', 'Saison', 'Frais', 'Categorie', 'Secteur'];
      var lines = [headersCSV.join(',')];
      arr.forEach(function (x) {
        var r = x.row, nomc = ((r['Prénom'] || r['Prenom'] || '') + ' ' + (r['Nom'] || '')).trim();
        var passText8 = (typeof normalizePassportToText8_ === 'function') ? normalizePassportToText8_(r['Passeport #']) : String(r['Passeport #']||'');
        var vals = [passText8, nomc, (r['Saison'] || ''), (r['Nom du frais'] || r['Frais'] || r['Produit'] || ''), (r['Catégorie'] || r['Categorie'] || ''), (deriveSectorFromRow_(r) || '')]
          .map(function (v) { v = String(v).replace(/"/g, '""'); if (/[",\n;]/.test(v)) v = '"' + v + '"'; return v; });
        lines.push(vals.join(','));
      });
      var csv = lines.join('\n');
      var blob = Utilities.newBlob(csv, 'text/csv', 'Nouveaux_' + sec.replace(/[^A-Za-z0-9\-]/g, '') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + '.csv');

      var subject = renderTemplate_(tplSub, { secteur: sec, date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') });
      var body = renderTemplate_(tplBody, { secteur: sec });

      if (!dry && to) {
        if (useGmailApp) {
          GmailApp.sendEmail(to, subject, '', { htmlBody: body, cc: cc || undefined, attachments: [blob], name: fromName });
        } else {
          MailApp.sendEmail({ to: to, cc: cc, subject: subject, htmlBody: body, attachments: [blob], name: fromName });
        }
      }
    });
  }

  appendImportLog_(ss, 'MAIL_WORKER', JSON.stringify({ processed: processed, summaries: processedNew.length > 0 }));
  return { processed: processed, summaries: processedNew.length > 0 };
}



/**
 * Enfile les courriels "INSCRIPTION_NEW" UNIQUEMENT pour les membres
 * qui matchent au moins un secteur de confirmation (ErrorCode vide).
 * -> pas de doublons (clé: Type||KeyHash)
 */
function enqueueInscriptionNewBySectors(seasonSheetId){
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var outSh   = upgradeMailOutboxForDisplay_(ss);
  var headers = getMailOutboxHeaders_();
  var idx     = getHeadersIndex_(outSh, headers.length);

  // (1) Secteurs de confirmation (ErrorCode vide)
  var sectors = _loadMailSectors_(ss).filter(function(s){ return !String(s.errorCode||'').trim(); });
  if (!sectors.length) { appendImportLog_(ss, 'QUEUE_NEW', '0 sectors (no confirmation sectors)'); return { queued: 0 }; }

  // (2) Index OUTBOX existant (Type||KeyHash)
  var existing = {};
  var outLast = outSh.getLastRow();
  if (outLast >= 2) {
    var outVals = outSh.getRange(2,1,outLast-1, headers.length).getValues();
    var iT = idx['Type']-1, iKH = idx['KeyHash']-1;
    for (var i=0;i<outVals.length;i++){
      var t = String(outVals[i][iT]||'').trim();
      var kh= String(outVals[i][iKH]||'').trim();
      if (t && kh) existing[t+'||'+kh] = true;
    }
  }

  // (3) Balayage des INSCRIPTIONS et construction du buffer
  var finals = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS).rows || [];
  var now = new Date();
  var rows = [];         // buffer pour enqueueOutboxRows_
  var denormSrc = [];    // lignes sources pour backfill

  finals.forEach(function(r){
    // KeyHash stable
    // ⛔ On n'envoie pas de confirmations aux "frais coach"
    if (_isCoachMemberSafe_(ss, r)) return;
    var keyColsCsv = readParam_(ss, PARAM_KEYS.KEY_COLS) || 'Passeport #,Saison';
    var keyStr = keyColsCsv.split(',').map(function(kc){ kc=kc.trim(); return r[kc] == null ? '' : String(r[kc]); }).join('||');
    var kh = Utilities.base64EncodeWebSafe(Utilities.newBlob(keyStr).getBytes());
    if (existing['INSCRIPTION_NEW||'+kh]) return;

    // Exclusion (member Exclude=TRUE)
    if (isExcludedMember_(ss, r)) return;

    // Match secteur
    var payload = buildDataFromRow_(r) || {};
    if (payload.U_num == null) {
      var m = String(deriveUFromRow_(r)||'').match(/U\s*-?\s*(\d{1,2})/i);
      payload.U_num = m ? parseInt(m[1],10) : 0;
      payload.genreInitiale = payload.genreInitiale || '';
    }
    var s = _matchSectorForType_(sectors, payload, 'INSCRIPTION_NEW');
    if (!s) return;

    // Ligne OUTBOX « squelette »
    var arr = new Array(headers.length).fill('');
    function set(col, val){ var i = idx[col]; if (i) arr[i-1] = val; }
    set('Type', 'INSCRIPTION_NEW');
    set('Status', 'pending');
    set('KeyHash', kh);
    set('CreatedAt', now);
    rows.push(arr);
    denormSrc.push(r);

    existing['INSCRIPTION_NEW||'+kh] = true;
  });

  // (4) Écriture batch
  if (!rows.length) { appendImportLog_(ss, 'QUEUE_NEW', JSON.stringify({ queued: 0, sectors: sectors.length })); return { queued: 0 }; }

  var startRow = outSh.getLastRow() + 1;
  enqueueOutboxRows_(ss.getId(), rows);

  // (5) Backfill lisibles en bloc (Passeport formaté)
  try {
    var n = denormSrc.length;
    var pass = new Array(n), nomc = new Array(n), frais = new Array(n);
    for (var k=0; k<n; k++){
      var r = denormSrc[k];
      var passText8 = (typeof normalizePassportToText8_ === 'function')
        ? normalizePassportToText8_(r['Passeport #'])
        : String(r['Passeport #']||'');
      pass[k]  = [ passText8 ];
      nomc[k]  = [ (((r['Prénom']||r['Prenom']||'') + ' ' + (r['Nom']||'')).trim()) ];
      frais[k] = [ r['Nom du frais'] || r['Frais'] || r['Produit'] || '' ];
    }
    if (idx['Passeport'])   outSh.getRange(startRow, idx['Passeport'],   n, 1).setValues(pass);
    if (idx['NomComplet'])  outSh.getRange(startRow, idx['NomComplet'],  n, 1).setValues(nomc);
    if (idx['Frais'])       outSh.getRange(startRow, idx['Frais'],       n, 1).setValues(frais);
  } catch(e) {
    // non bloquant
  }

  appendImportLog_(ss, 'QUEUE_NEW', JSON.stringify({ queued: rows.length, sectors: sectors.length }));
  return { queued: rows.length };
}


/* ======================== Enqueue error emails ======================== */
function enqueueValidationEmailsByErrorCode(seasonSheetId, errorCode){
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  if (!errorCode) throw new Error('errorCode requis');

  var errRows = (readSheetAsObjects_(ss.getId(), SHEETS.ERREURS).rows || []);
  var outSh   = upgradeMailOutboxForDisplay_(ss);
  var headers = getMailOutboxHeaders_();
  var idx     = getHeadersIndex_(outSh, headers.length);

  // Index OUTBOX existant (Type||KeyHash)
  var existing = {};
  var outLast = outSh.getLastRow();
  if (outLast >= 2) {
    var outVals = outSh.getRange(2,1,outLast-1, headers.length).getValues();
    var iT = idx['Type']-1, iKH = idx['KeyHash']-1;
    for (var i=0;i<outVals.length;i++){
      var t = String(outVals[i][iT]||'').trim();
      var kh= String(outVals[i][iKH]||'').trim();
      if (t && kh) existing[t+'||'+kh] = true;
    }
  }

  // Helper: retrouve {kh,row} à partir (passeport,saison)
  function keyHashAndRowFor(passport, saison){
    var finals = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS).rows || [];
    var found = null;
    for (var i=0;i<finals.length;i++){
      var p = String(finals[i]['Passeport #']||'').trim();
      if (p === passport || (p && p.replace(/^0+/, '') === String(passport||'').replace(/^0+/, ''))) {
        if (!saison || String(finals[i]['Saison']||'').trim() === String(saison||'').trim()) { found = finals[i]; break; }
      }
    }
    if (!found) return { kh:null, row:null };
    var keyColsCsv = readParam_(ss, PARAM_KEYS.KEY_COLS) || 'Passeport #,Saison';
    var keyCols = keyColsCsv.split(',').map(function(x){ return x.trim(); });
    var keyStr = keyCols.map(function(kc){ return found[kc] == null ? '' : String(found[kc]); }).join('||');
    var kh = Utilities.base64EncodeWebSafe(Utilities.newBlob(keyStr).getBytes());
    return { kh: kh, row: found };
  }

  // Buffers batch
  var outRows = [];
  var denorm = []; // garde la row source (finals) alignée sur outRows

  // Construction batch
  for (var r=0; r<errRows.length; r++){
    var e = errRows[r];
    if (String(e['Type']||'') !== errorCode) continue;
    var passport = String(e['Passeport']||e['Passeport #']||'').trim();
    if (!passport) continue;
    var saison   = String(e['Saison']||'').trim();

    var k = keyHashAndRowFor(passport, saison);
    if (!k.kh) continue;

// ⛔ Jamais d'e-mail "erreur" aux coachs
if (_isCoachMemberSafe_(ss, k.row)) continue;

    // skip si déjà en OUTBOX
    if (existing[errorCode+'||'+k.kh]) continue;

    // exclusions globales (ex.: Adapté/Adulte) — on n’ennuie pas ces membres avec des mails d’erreurs
    if (typeof isExcludedMember_ === 'function' && isExcludedMember_(ss, k.row)) continue;

    // payload d’erreur compact
    var errPayload = {
      label: String(e['Message']||'').trim() || errorCode,
      details: String(e['Contexte']||'').trim()
    };

    // construit une ligne outbox squelettique (respecte l’ordre headers)
    var arr = new Array(headers.length).fill('');
    function set(col, val){ var i = idx[col]; if (i) arr[i-1] = val; }
    set('Type', errorCode);
    set('Status', 'pending');
    set('KeyHash', k.kh);
    set('CreatedAt', new Date());
    set('Error', JSON.stringify(errPayload));

    outRows.push(arr);
    denorm.push(k.row);
    existing[errorCode+'||'+k.kh] = true;
  }

  if (!outRows.length) {
    appendImportLog_(ss, 'QUEUE_ERRMAIL', JSON.stringify({ code: errorCode, queued: 0 }));
    return { queued: 0 };
  }

  // Écriture batch dans MAIL_OUTBOX
  var startRow = outSh.getLastRow() + 1;
  enqueueOutboxRows_(ss.getId(), outRows);

  // Backfill lisibles en bloc
  try {
    var n = denorm.length;
    var pass = new Array(n), nomc = new Array(n), frais = new Array(n);
    for (var k=0; k<n; k++){
      var r0 = denorm[k];
      pass[k]  = [ r0['Passeport #'] || '' ];
      nomc[k]  = [ (((r0['Prénom']||r0['Prenom']||'') + ' ' + (r0['Nom']||'')).trim()) ];
      frais[k] = [ r0['Nom du frais'] || r0['Frais'] || r0['Produit'] || '' ];
    }
    if (idx['Passeport'])   outSh.getRange(startRow, idx['Passeport'],   n, 1).setValues(pass);
    if (idx['NomComplet'])  outSh.getRange(startRow, idx['NomComplet'],  n, 1).setValues(nomc);
    if (idx['Frais'])       outSh.getRange(startRow, idx['Frais'],       n, 1).setValues(frais);
  } catch (eBF) { /* non bloquant */ }

  appendImportLog_(ss, 'QUEUE_ERRMAIL', JSON.stringify({ code: errorCode, queued: outRows.length }));
  return { queued: outRows.length };
}




/* === Récupération ligne finale par KeyHash (inchangé v0.7) === */
function fetchFinalRowByKeyHash_(ss, keyHash) {
  var keyStr = Utilities.newBlob(Utilities.base64DecodeWebSafe(keyHash)).getDataAsString(); // "Passeport||Saison" par défaut
  var finals = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var keyColsCsv = readParam_(ss, PARAM_KEYS.KEY_COLS) || 'Passeport #,Saison';
  var keyCols = keyColsCsv.split(',').map(function(x){ return x.trim(); });
  for (var i=0;i<finals.rows.length;i++){
    var r = finals.rows[i];
    var k = keyCols.map(function(kc){ return r[kc] == null ? '' : String(r[kc]); }).join('||');
    if (k === keyStr) return r;
  }
  return null;
}

/* === Résolution des destinataires par défaut (inchangé v0.7) === */
function resolveRecipient_(ss, type, row) {
  var to = '', cc = '';
  if (type === 'INSCRIPTION_NEW') {
    var csv = readParam_(ss, PARAM_KEYS.TO_FIELDS_INSCRIPTIONS) || 'Courriel,Parent 1 - Courriel,Parent 2 - Courriel';
    to = collectEmailsFromRow_(row, csv);
  }
  if (!to) { to = readParam_(ss, 'MAIL_TO_NEW_INSCRIPTIONS') || ''; cc = readParam_(ss, 'MAIL_CC_NEW_INSCRIPTIONS') || ''; }
  return { to: to, cc: cc };
}

/* === OUTBOX helpers (inchangés) === */
function getMailOutboxHeaders_(){ return ['Type','To','Cc','Sujet','Corps','Attachments','KeyHash','Status','CreatedAt','SentAt','Error']; }
function ensureMailOutbox_(ss){
  var headers=getMailOutboxHeaders_(); var sh=ss.getSheetByName(SHEETS.MAIL_OUTBOX);
  if (!sh){ sh=ss.insertSheet(SHEETS.MAIL_OUTBOX); sh.getRange(1,1,1,headers.length).setValues([headers]); return sh; }
  var last=sh.getLastRow(); if (last===0){ sh.getRange(1,1,1,headers.length).setValues([headers]); return sh; }
  var first=sh.getRange(1,1,1,headers.length).getValues()[0]; var ok=headers.every(function(h,i){ return String(first[i]||'')===h; });
  if(!ok){ sh.insertRowsBefore(1,1); sh.getRange(1,1,1,headers.length).setValues([headers]); }
  return sh;
}

/** Ajoute (si besoin) 3 colonnes lisibles en fin de MAIL_OUTBOX: Passeport, NomComplet, Frais */
function upgradeMailOutboxForDisplay_(ss){
  var sh = ensureMailOutbox_(ss); // crée si besoin
  var firstRow = sh.getLastColumn() ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] : [];
  var have = {}; firstRow.forEach(function(h){ have[String(h||'')] = true; });
  var add = [];
  ['Passeport','NomComplet','Frais'].forEach(function(h){
    if (!have[h]) add.push(h);
  });
  if (!add.length) return sh;

  // append les nouvelles colonnes et inscrit l'entête
  sh.insertColumnsAfter(sh.getLastColumn() || 1, add.length);
  var hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  for (var i=0;i<add.length;i++){
    sh.getRange(1, hdr.length - add.length + 1 + i).setValue(add[i]);
  }
  return sh;
}

/** Remplit les colonnes lisibles pour la ligne outbox nouvellement insérée */
function fillOutboxDenorm_(sh, finalRowObj){
  if (!finalRowObj) return;
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1,1,1,lastCol).getValues()[0].map(String);
  var idx = {}; headers.forEach(function(h,i){ idx[h]=i+1; });
  function set(colName, val){
    var c = idx[colName]; if (!c) return;
    sh.getRange(lastRow, c).setValue(val);
  }
  var passeport = finalRowObj['Passeport #'] || '';
  var nomc = (((finalRowObj['Prénom']||finalRowObj['Prenom']||'') + ' ' + (finalRowObj['Nom']||'')).trim());
  var frais = finalRowObj['Nom du frais'] || finalRowObj['Frais'] || finalRowObj['Produit'] || '';
  set('Passeport', passeport);
  set('NomComplet', nomc);
  set('Frais', frais);
}
// utils.js (ou email.js, même fichier que tes enqueues)
function _normFold_(s){ s=String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s.toUpperCase().trim(); }

function isExcludedMember_(ss, row){
  var map = readSheetAsObjects_(ss.getId(), SHEETS.MAPPINGS);
  var maps = (map && map.rows) ? map.rows : [];
  var hay = _normFold_((row['Catégorie']||row['Categorie']||'') + ' ' + (row['Nom du frais']||row['Frais']||row['Produit']||''));
  for (var i=0;i<maps.length;i++){
    var m = maps[i];
    if (String(m['Type']||'').toLowerCase() !== 'member') continue;
    var excl = String(m['Exclude']||'').toLowerCase()==='true';
    if (!excl) continue;
    var ali = _normFold_(m['AliasContains']||m['Alias']||'');
    if (!ali) continue;
    if (hay.indexOf(ali) !== -1) return true;
  }
  return false;
}
