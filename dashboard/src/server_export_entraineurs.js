/***** server_exports_entraineurs.js (drop-in) *****/

/* ===== Fallbacks légers (si non fournis par la lib) ===== */
if (typeof SHEETS === 'undefined') {
  var SHEETS = { INSCRIPTIONS:'INSCRIPTIONS', MAIL_OUTBOX:'MAIL_OUTBOX', MAIL_LOG:'MAIL_LOG', PARAMS:'PARAMS' };
}
if (typeof getSeasonSpreadsheet_ !== 'function') {
  function getSeasonSpreadsheet_(id){ if(!id) throw new Error('seasonSheetId manquant'); return SpreadsheetApp.openById(id); }
}
if (typeof readParam_ !== 'function') {
  function readParam_(ss, key){
    var sh = ss.getSheetByName(SHEETS.PARAMS);
    if (sh) {
      var last = sh.getLastRow();
      if (last>=1){
        var data = sh.getRange(1,1,last,2).getValues();
        for (var i=0;i<data.length;i++) if ((data[i][0]+'').trim()===key) return (data[i][1]+'').trim();
      }
    }
    var props = PropertiesService.getDocumentProperties(); return (props.getProperty(key)||'').trim();
  }
}
if (typeof appendImportLog_ !== 'function') {
  function appendImportLog_(ss, action, details){
    var sh = ss.getSheetByName('IMPORT_LOG') || ss.insertSheet('IMPORT_LOG');
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,3).setValues([['Horodatage','Action','Détails']]);
    sh.appendRow([new Date(), action, details || '']);
  }
}
if (typeof normalizePassportPlain8_ !== 'function') {
  function normalizePassportPlain8_(p){
    var s = String(p==null?'':p).replace(/\D/g,''); return ('00000000'+s).slice(-8);
  }
}

/* ===== Helpers tableurs ===== */
function _getSheet_(ss, name, createIfMissing){
  var sh = ss.getSheetByName(name);
  if (!sh && createIfMissing) sh = ss.insertSheet(name);
  return sh;
}
function _readSheetAsObjectsByName_(ss, sheetName){
  var sh = _getSheet_(ss, sheetName, false);
  if (!sh || sh.getLastRow()<1 || sh.getLastColumn()<1) return {sheet: sh, headers:[], rows:[]};
  var values = sh.getRange(1,1,sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var H = (values[0]||[]).map(String), out=[];
  for (var r=1;r<values.length;r++){
    var o={}; for (var c=0;c<H.length;c++) o[H[c]] = values[r][c]; out.push(o);
  }
  return { sheet: sh, headers: H, rows: out };
}

/* ===== Runners appelés par l’UI ===== */
function runExportEntraineursMembres(){
  var sid = (typeof getSeasonId_==='function') ? getSeasonId_() : null;
  if (!sid) throw new Error('SeasonId manquant');
  return exportRetroEntraineursMembresXlsxToDrive(sid, {});
}

function runExportEntraineursGroupes(){
  var sid = (typeof getSeasonId_==='function') ? getSeasonId_() : null;
  if (!sid) throw new Error('SeasonId manquant');
  return exportRetroEntraineursGroupesXlsxToDrive(sid, {});
}


/* ===== Exports (nouvelle logique) ===== */

/** Exporte la liste des entraîneurs (membres) valides pour la saison */
function exportRetroEntraineursMembresXlsxToDrive(seasonId, opts) {
  var ss = getSeasonSpreadsheet_(seasonId);
  var membres = _ex_fetchMembresEntraineurs_(ss, seasonId); // basé sur INSCRIPTIONS_ENTRAINEURS + MEMBRES_GLOBAL

  var headers = [
    'Passeport','Prenom','Nom','DateNaissance','Genre','StatutMembre',
    'PhotoExpireLe','CasierExpiré'
  ];
  var rows = membres.map(function(m){
    return [
      m.Passeport, m.Prenom, m.Nom, m.DateNaissance, m.Genre, m.StatutMembre,
      m.PhotoExpireLe, m.CasierExpiré
    ];
  });

  return _ex_writeXlsxToDrive_(ss, "retro-entraineurs-membres-"+seasonId+".xlsx", headers, rows, opts);
}

/** Exporte les groupes (rôles) des entraîneurs valides */
function exportRetroEntraineursGroupesXlsxToDrive(seasonId, opts) {
  var ss = getSeasonSpreadsheet_(seasonId);
  var roles = _ex_fetchRolesFilteredOnValidCoaches_(ss, seasonId); // filtré sur coachs valides

  var headers = ['Passeport','Type','Groupe','Commentaire'];
  var rows = roles.map(function(r){
    var type = 'entraineur';
    var groupe = (r.Role||'').trim();
    // (optionnel) enrichir si besoin :
    var extra = [];
    if (r.Categorie) extra.push(r.Categorie);
    if (r.Equipe)    extra.push(r.Equipe);
    if (extra.length) groupe += ' | ' + extra.join(' | ');
    return [r.Passeport, type, groupe, r.Commentaire || ''];
  });

  return _ex_writeXlsxToDrive_(ss, "retro-entraineurs-groupes-"+seasonId+".xlsx", headers, rows, opts);
}

/* ===== Data helpers ===== */

/** Retourne Set de passeports coachs valides (INSCRIPTIONS_ENTRAINEURS, Saison = seasonId, CoachValid=1) */
function _ex_validCoachPassportsSet_(ss, seasonId){
  var x = _readSheetAsObjectsByName_(ss, 'INSCRIPTIONS_ENTRAINEURS');
  var set = new Set();
  if (!x.rows.length) return set;

  x.rows.forEach(function(r){
    if ((String(r['Saison']||'').trim() !== String(seasonId||'').trim())) return;
    var ok = (String(r['CoachValid']||'').toLowerCase()==='true') || Number(r['CoachValid']||0)===1;
    if (!ok) return;
    var p8 = normalizePassportPlain8_(r['Passeport']||r['Passeport #']||'');
    if (p8) set.add(p8);
  });
  return set;
}

/** Jointure MEMBRES_GLOBAL x INSCRIPTIONS_ENTRAINEURS (valides) */
function _ex_fetchMembresEntraineurs_(ss, seasonId) {
  var validSet = _ex_validCoachPassportsSet_(ss, seasonId);

  // fallback si pas de vue => on retombe sur ROLES (compat)
  if (!validSet.size) return _ex_fetchMembresFromRolesFallback_(ss, seasonId);

  var sh = _getSheet_(ss, (readParam_(ss,'SHEET_MEMBRES_GLOBAL')||'MEMBRES_GLOBAL'), false);
  if (!sh || sh.getLastRow()<2) return [];

  var V = sh.getDataRange().getValues();
  var H = V[0].map(String);
  var c = {}; H.forEach(function(h,i){ c[h]=i; });

  function val(r,k){ var i=c[k]; return (i==null)?'':V[r][i]; }

  var out = [];
  for (var r=1;r<V.length;r++){
    var pass = normalizePassportPlain8_(val(r,'Passeport'));
    if (!pass || !validSet.has(pass)) continue;
    out.push({
      Passeport: pass,
      Prenom: String(val(r,'Prenom')||''),
      Nom: String(val(r,'Nom')||''),
      DateNaissance: String(val(r,'DateNaissance')||''),
      Genre: String(val(r,'Genre')||''),
      StatutMembre: String(val(r,'StatutMembre')||''),
      PhotoExpireLe: String(val(r,'PhotoExpireLe')||''),
      CasierExpiré:  String(val(r,'CasierExpiré')||val(r,'CasierExpire')||'')
    });
  }
  return out;
}

/** Fallback historique: si la vue coach n’existe pas, partir des ROLES */
function _ex_fetchMembresFromRolesFallback_(ss, seasonId){
  var roles = _ex_fetchRolesRaw_(ss, seasonId);
  if (!roles.length) return [];
  var passports = Array.from(new Set(roles.map(function(r){ return normalizePassportPlain8_(r.Passeport); })));

  var sh = _getSheet_(ss, (readParam_(ss,'SHEET_MEMBRES_GLOBAL')||'MEMBRES_GLOBAL'), false);
  if (!sh || sh.getLastRow()<2) return [];

  var V = sh.getDataRange().getValues();
  var H = V[0].map(String);
  var c = {}; H.forEach(function(h,i){ c[h]=i; });

  var idx = new Map();
  for (var r=1;r<V.length;r++){
    var pass = normalizePassportPlain8_(V[r][c['Passeport']]); if (!pass) continue;
    idx.set(pass, V[r]);
  }

  var out = [];
  passports.forEach(function(p){
    var row = idx.get(p); if (!row) return;
    out.push({
      Passeport: p,
      Prenom: row[c['Prenom']]||'',
      Nom: row[c['Nom']]||'',
      DateNaissance: row[c['DateNaissance']]||'',
      Genre: row[c['Genre']]||'',
      StatutMembre: row[c['StatutMembre']]||'',
      PhotoExpireLe: row[c['PhotoExpireLe']]||'',
      CasierExpiré:  row[c['CasierExpiré']]||row[c['CasierExpire']]||''
    });
  });
  return out;
}

/** ROLES “raw” (toutes lignes valides pour la saison) */
function _ex_fetchRolesRaw_(ss, seasonId){
  var sh = _getSheet_(ss, 'ENTRAINEURS_ROLES', false);
  if (!sh || sh.getLastRow()<2) return [];
  var V = sh.getDataRange().getValues();
  var H = V[0].map(String), c={}; H.forEach(function(h,i){ c[h]=i; });

  var rows = [];
  for (var r=1;r<V.length;r++){
    var saison = String(V[r][c['Saison']]||'').trim();
    if (!saison || saison !== String(seasonId)) continue;
    var pass = normalizePassportPlain8_(V[r][c['Passeport']]||'');
    if (!pass) continue;
    rows.push({
      Passeport: pass,
      Saison: saison,
      Role: String(V[r][c['Role']]||'').trim(),
      Categorie: String(V[r][c['Categorie']]||'').trim(),
      Equipe: String(V[r][c['Equipe']]||'').trim(),
      Commentaire: String(V[r][c['Commentaire']]||'').trim()
    });
  }
  return rows;
}

/** ROLES filtrés sur coachs valides (INSCRIPTIONS_ENTRAINEURS) */
function _ex_fetchRolesFilteredOnValidCoaches_(ss, seasonId){
  var rows = _ex_fetchRolesRaw_(ss, seasonId);
  var validSet = _ex_validCoachPassportsSet_(ss, seasonId);
  if (!validSet.size) return rows; // fallback si pas de vue
  return rows.filter(function(r){ return validSet.has(r.Passeport); });
}

/* ===== XLSX writer ===== */
function _ex_writeXlsxToDrive_(ss, filename, headers, rows, opts) {
  var tmp = SpreadsheetApp.create('TMP_'+filename);
  var sh = tmp.getSheets()[0];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows && rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);

  var blob = DriveApp.getFileById(tmp.getId())
    .getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    .setName(filename);

  var folderId = (opts && opts.folderId) || readParam_(ss, 'DRIVE_FOLDER_EXPORTS');
  var folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = folder.createFile(blob);

  // ménage
  DriveApp.getFileById(tmp.getId()).setTrashed(true);

  appendImportLog_(ss, 'EXPORT_DONE', filename);
  return { fileId: file.getId(), name: filename };
}

/* ===== KPIs (coach = via INSCRIPTIONS_ENTRAINEURS) ===== */
function getKpisBoth() {
  var seasonId = (typeof getSeasonId_==='function' ? getSeasonId_() : null);
  var ss = getSeasonSpreadsheet_(seasonId);

  // ===== Entraîneurs via INSCRIPTIONS_ENTRAINEURS =====
  var coach = _readSheetAsObjectsByName_(ss, 'INSCRIPTIONS_ENTRAINEURS');
  var kE = { photosEchues:0, photosARenouv:0, casiersExp:0, total:0 };
  if (coach.rows.length){
    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    coach.rows.forEach(function(r){
      if (String(r['Saison']||'').trim() !== String(seasonId)) return;
      kE.total++;
      var photoBad = (String(r['PhotoInvalideFlag']||'').toLowerCase()==='true') || Number(r['PhotoInvalideFlag']||0)===1;
      var dueLe    = String(r['PhotoInvalideDuesLe']||'');
      var casBad   = (String(r['CasierExpireFlag']||'').toLowerCase()==='true') || Number(r['CasierExpireFlag']||0)===1;

      if (photoBad) kE.photosEchues++;
      else if (dueLe && dueLe <= today) kE.photosARenouv++;

      if (casBad) kE.casiersExp++;
    });
  }

  // ===== Joueurs (conservé, set depuis feuille finale INSCRIPTIONS) =====
  var kJ = { photosEchues:0, photosARenouv:0, total:0 };

  var shJ = ss.getSheetByName(SHEETS.INSCRIPTIONS) || ss.getSheetByName('inscriptions');
  var shMG= ss.getSheetByName(readParam_(ss,'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL');

  if (shJ && shMG && shJ.getLastRow()>1 && shMG.getLastRow()>1){
    var J = shJ.getDataRange().getValues(); var hj = J[0], jp = hj.indexOf('Passeport')>=0 ? hj.indexOf('Passeport') : hj.indexOf('Passeport #');
    var joueursSet = new Set();
    for (var r=1;r<J.length;r++){
      var p = normalizePassportPlain8_(J[r][jp]);
      if (p) joueursSet.add(p);
    }
    var V = shMG.getDataRange().getValues(); var H = V[0];
    var cPass=H.indexOf('Passeport'), cPh=H.indexOf('PhotoExpireLe'), cFlag=H.indexOf('PhotoInvalide'), cDue=H.indexOf('PhotoInvalideDuesLe');

    var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    for (var i=1;i<V.length;i++){
      var p = normalizePassportPlain8_(V[i][cPass]); if (!p || !joueursSet.has(p)) continue;
      kJ.total++;
      var flag = String(V[i][cFlag]||'').toLowerCase()==='true' || Number(V[i][cFlag]||0)===1;
      var due  = String(V[i][cDue]||'');
      if (flag) kJ.photosEchues++;
      else if (due && due <= today) kJ.photosARenouv++;
    }
  }

  return { joueurs: kJ, entraineurs: kE };
}

/* ===== Liste joueurs à corriger (conservée) ===== */
function getJoueursPhotosProblemes() {
  var seasonId = (typeof getSeasonId_==='function' ? getSeasonId_() : null);
  var ss = getSeasonSpreadsheet_(seasonId);

  var shMG = ss.getSheetByName(readParam_(ss,'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL');
  var shJ  = ss.getSheetByName(SHEETS.INSCRIPTIONS) || ss.getSheetByName('inscriptions');
  if (!shMG || !shJ || shMG.getLastRow()<2 || shJ.getLastRow()<2) return [];

  // set des joueurs inscrits
  var J = shJ.getDataRange().getValues();
  var hj = J[0], jp = hj.indexOf('Passeport')>=0 ? hj.indexOf('Passeport') : hj.indexOf('Passeport #');
  var joueursSet = new Set();
  for (var r=1;r<J.length;r++){
    var p = normalizePassportPlain8_(J[r][jp]); if (p) joueursSet.add(p);
  }

  var V = shMG.getDataRange().getValues(), H = V[0];
  var cPass=H.indexOf('Passeport'), cNom=H.indexOf('Nom'), cPre=H.indexOf('Prenom'),
      cPh=H.indexOf('PhotoExpireLe'), cFlag=H.indexOf('PhotoInvalide'), cDue=H.indexOf('PhotoInvalideDuesLe');
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var out = [];
  for (var i=1;i<V.length;i++){
    var p = normalizePassportPlain8_(V[i][cPass]);
    if (!p || !joueursSet.has(p)) continue;

    var photo = String(V[i][cPh] || '');
    var flag  = String(V[i][cFlag]||'').toLowerCase()==='true' || Number(V[i][cFlag]||0)===1;
    var due   = String(V[i][cDue]||'');

    if (!flag && !(due && due <= today)) continue;

    var statut = flag ? 'ÉCHUE' : 'À RENOUVELER';
    out.push({
      Passeport: p,
      Nom: String(V[i][cNom]||''),
      Prenom: String(V[i][cPre]||''),
      PhotoExpireLe: photo || '',
      StatutPhoto: statut
    });
  }

  out.sort(function(a,b){
    var A=(a.Nom||'')+(a.Prenom||''); var B=(b.Nom||'')+(b.Prenom||'');
    return A.localeCompare(B,'fr');
  });
  return out;
}
