/***** server_membres.js *****/
// --- Shims utilitaires (si absents du projet courant)
if (typeof getSheetOrCreate_ !== 'function') {
  function getSheetOrCreate_(ss, name, headersOpt) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      if (headersOpt && headersOpt.length) {
        sh.getRange(1, 1, 1, headersOpt.length).setValues([headersOpt]);
      }
    }
    return sh;
  }
}
// --- Shims passeport (utilise tes fonctions si elles existent déjà)
if (typeof normalizePassportPlain8_ !== 'function') {
  function normalizePassportPlain8_(v) {
    if (typeof normalizePassportToText8_ === 'function') {
      return normalizePassportToText8_(v); // ta version texte 8 chars si dispo
    }
    var s = String(v || '').replace(/\D/g, '');
    return s ? s.padStart(8, '0') : '';
  }
}

if (typeof getSheet_ !== 'function') {
  // Alias pratique (certains modules appellent getSheet_(..., true))
  function getSheet_(ss, name, createIfMissing) {
    return createIfMissing ? getSheetOrCreate_(ss, name) : ss.getSheetByName(name);
  }
}

/**
 * Importe le dernier export Validation_Membres du dossier paramétré
 * et UPSERT dans le classeur CENTRAL (targetSpreadsheetId) la feuille MEMBRES_GLOBAL.
 *
 * - Lit les paramètres dans le classeur SAISON (getSeasonId_()):
 *   - DRIVE_FOLDER_VALIDATION_MEMBRES
 *   - SEASON_YEAR
 *   - PHOTO_INVALID_FROM_MMDD
 * - Écrit dans le classeur CENTRAL (targetSpreadsheetId):
 *   MEMBRES_GLOBAL avec l'entête normalisée (voir colOrder).
 *
 * Attentes helpers existants:
 *  - normalizePassportPlain8_, _vm_findLatestActiveFile_, _vm_readFileAsSheetValues_,
 *    _vm_indexCols_, _vm_toISODate_, _vm_to01_, _vm_hashRow_, _vm_groupRuns_,
 *    getSheetOrCreate_, log_, _vm_nowISO_, _vm_rowValues_, _vm_notifyImportMembres_
 */
function importValidationMembresToGlobal_(targetSpreadsheetId) {
  // --- SAISON: params & dossier source ---
  var seasonId = getSeasonId_(); // on s'appuie sur ta saison active
  var ssSeason = SpreadsheetApp.openById(seasonId);

  var folderId = readParam_ ? readParam_(ssSeason, 'DRIVE_FOLDER_VALIDATION_MEMBRES')
                            : (readParamValue('DRIVE_FOLDER_VALIDATION_MEMBRES') || '');
  if (!folderId) throw new Error('Paramètre DRIVE_FOLDER_VALIDATION_MEMBRES manquant (classeur saison).');

  var seasonYear = Number( (readParam_ ? readParam_(ssSeason, 'SEASON_YEAR') : readParamValue('SEASON_YEAR')) || new Date().getFullYear() );
  var photoInvalidFrom = String( (readParam_ ? readParam_(ssSeason, 'PHOTO_INVALID_FROM_MMDD') : readParamValue('PHOTO_INVALID_FROM_MMDD')) || '04-01' ).trim();

  // --- CENTRAL: feuille cible ---
  if (!targetSpreadsheetId) throw new Error('ID du classeur CENTRAL manquant (targetSpreadsheetId).');
  var ssCentral = SpreadsheetApp.openById(targetSpreadsheetId);

  // --- Fichier source ---
  var found = _vm_findLatestActiveFile_(folderId);
  var file = found && found.file;
  if (!file) {
    log_('VM_IMPORT_NOFILE', 'Aucun fichier actif dans le dossier Validation_Membres.');
    return { updated: 0, unchanged: 0, created: 0 };
  }
  log_('VM_IMPORT_START', 'Lecture: ' + file.getName() + ' (' + file.getId() + ')');

  var values = _vm_readFileAsSheetValues_(file);
  if (!values || values.length < 2) throw new Error('Fichier Validation_Membres vide ou illisible.');

  var headers = values[0].map(function(h){ return String(h).trim(); });
  var idx = _vm_indexCols_(headers, [
    'Passeport #', 'Prénom', 'Nom', 'Date de naissance',
    'Identité de genre', 'Statut du membre',
    'Date d\'expiration de la photo de profil',
    'Vérification du casier judiciaire est expiré'
  ]);

  var seasonInvalidDate = seasonYear + '-' + photoInvalidFrom;   // ex 2025-04-01
  var cutoffNextJan1    = (seasonYear + 1) + '-01-01';           // ex 2026-01-01

  // --- Build: passeport -> objet cible (aligné EXACTEMENT sur l'entête finale) ---
  var targetByPassport = new Map();
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row || row.length === 0) continue;

    var passportRaw = row[idx['Passeport #']];
    var passport = normalizePassportPlain8_(passportRaw);
    if (!passport) continue;

    var prenom   = String(row[idx['Prénom']] || '').trim();
    var nom      = String(row[idx['Nom']] || '').trim();
    var dobRaw   = row[idx['Date de naissance']];
    var genre    = String(row[idx['Identité de genre']] || '').trim();
    var statut   = String(row[idx['Statut du membre']] || '').trim();
    var photoExp = _vm_toISODate_(row[idx['Date d\'expiration de la photo de profil']]);
    var casier01 = _vm_to01_(row[idx['Vérification du casier judiciaire est expiré']]); // 1 = expiré

    var photoInvalide = (!photoExp || photoExp < cutoffNextJan1) ? 1 : 0;
    var photoDuesLe   = photoInvalide ? seasonInvalidDate : '';

    // ⚠️ IMPORTANT: utiliser les mêmes clés que l'entête finale (CasierExpiré avec accent)
    var obj = {
      'Passeport': passport,
      'Prenom': prenom,
      'Nom': nom,
      'DateNaissance': _vm_toISODate_(dobRaw),
      'Genre': genre,
      'StatutMembre': statut,
      'PhotoExpireLe': photoExp,
      'PhotoInvalideDuesLe': photoDuesLe,
      'PhotoInvalide': photoInvalide,
      'CasierExpiré': casier01,          // <-- clé EXACTE = entête
      'SeasonYear': seasonYear
    };

    obj.RowHash = _vm_hashRow_(obj);
    targetByPassport.set(passport, obj);
  }

  // --- Upsert CENTRAL.MEMBRES_GLOBAL (diff par RowHash) ---
  var sheetName = (readParam_ ? readParam_(ssSeason, 'SHEET_MEMBRES_GLOBAL') : readParamValue('SHEET_MEMBRES_GLOBAL')) || 'MEMBRES_GLOBAL';
  var sh = getSheetOrCreate_(ssCentral, sheetName);

  var colOrder = [
    'Passeport','Prenom','Nom','DateNaissance','Genre','StatutMembre',
    'PhotoExpireLe','PhotoInvalideDuesLe','PhotoInvalide',
    'CasierExpiré','SeasonYear','RowHash','LastUpdate'
  ];
  _vm_ensureHeader_(sh, colOrder);

  var data   = sh.getDataRange().getValues();
  var header = data[0];
  var colIdx = {};
  header.forEach(function(h, i){ colIdx[String(h)] = i; });

  // index existant: passeport -> {rowIndex, RowHash}
  var existingIdx = new Map();
  for (var i = 1; i < data.length; i++) {
    var pass = String(data[i][colIdx['Passeport']] || '').trim();
    if (pass) {
      existingIdx.set(pass, {
        rowIndex: i,
        hash: String(data[i][colIdx['RowHash']] || '').trim()
      });
    }
  }

  var nowISO = _vm_nowISO_();
  var toWrite = [];   // [rowIndex (1-based), rowValues]
  var toAppend = [];  // [rowValues]
  var updated = 0, created = 0, unchanged = 0;

  targetByPassport.forEach(function(obj, pass){
    var exists = existingIdx.get(pass);
    if (exists) {
      if (exists.hash === obj.RowHash) { unchanged++; return; }
      var rowValsU = _vm_rowValues_(colOrder, obj, nowISO);
      toWrite.push([exists.rowIndex + 1, rowValsU]); // +1: entête
      updated++;
    } else {
      var rowValsA = _vm_rowValues_(colOrder, obj, nowISO);
      toAppend.push(rowValsA);
      created++;
    }
  });

  // Écritures batchées
  if (toWrite.length) {
    toWrite.sort(function(a,b){ return a[0]-b[0]; });
    var runs = _vm_groupRuns_(toWrite);
    runs.forEach(function(run){
      var startRow = run[0][0];
      var block = run.map(function(x){ return x[1]; });
      sh.getRange(startRow, 1, block.length, colOrder.length).setValues(block);
    });
  }
  if (toAppend.length) {
    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, toAppend.length, colOrder.length).setValues(toAppend);
  }

  log_('VM_IMPORT_SUMMARY', 'created=' + created + ', updated=' + updated + ', unchanged=' + unchanged + ', total=' + targetByPassport.size);
  _vm_notifyImportMembres_(ssSeason, file, {created:created, updated:updated, unchanged:unchanged});

  return {created: created, updated: updated, unchanged: unchanged};
}
/** Initialiser la saison si membres_global est vide (l'onglet de la saison) */

function runSyncMembresGlobalSubsetFromCentral(){
  var seasonId = getSeasonId_();
  var centralId = readParamValue('GLOBAL_MEMBRES_SHEET_ID');
  if (!centralId) throw new Error('GLOBAL_MEMBRES_SHEET_ID manquant dans PARAMS.');
  var res = syncMembresGlobalSubsetFromCentral_(seasonId, centralId);
  appendImportLog_({ type:'MEMBRES_GLOBAL_SUBSET_OK', details: res });
  return res;
}

function syncMembresGlobalSubsetFromCentral_(seasonId, centralId){
  var ssSeason = SpreadsheetApp.openById(seasonId);
  var ssCentral = SpreadsheetApp.openById(centralId);

  // Collecter les passeports de la saison (inscriptions + coachs)
  var need = buildSeasonPassportSet_(ssSeason);

  var shC = ssCentral.getSheetByName('MEMBRES_GLOBAL');
  if (!shC || shC.getLastRow() < 2) throw new Error('MEMBRES_GLOBAL central vide.');
  var VC = shC.getDataRange().getValues();
  var HC = VC[0];

  function col_(name){ return HC.indexOf(name); }
  var cPass = col_('Passeport'),
      cPre  = col_('Prenom'),
      cNom  = col_('Nom'),
      cDOB  = col_('DateNaissance'),
      cGen  = col_('Genre'),
      cStat = col_('StatutMembre'),
      cPh   = col_('PhotoExpireLe'),
      cPid  = col_('PhotoInvalideDuesLe'),
      cPif  = col_('PhotoInvalide'),
      cCas  = col_('CasierExpiré');

  var mapC = {};
  for (var r=1; r<VC.length; r++){
    var p = normalizePassportPlain8_(VC[r][cPass]);
    if (p) mapC[p] = VC[r];
  }

  var sheetNameLocal = readParamValue('SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL';
  var shL = ssSeason.getSheetByName(sheetNameLocal) || ssSeason.insertSheet(sheetNameLocal);
  shL.clear();
  var header = ['Passeport','Prenom','Nom','DateNaissance','Genre','StatutMembre',
                'PhotoExpireLe','PhotoInvalideDuesLe','PhotoInvalide','CasierExpiré'];
  shL.getRange(1,1,1,header.length).setValues([header]);

  var out = [];
  need.forEach(function(p){
    var row = mapC[p];
    if (!row) return; // passeport non trouvé dans central
    out.push([
      p,
      row[cPre], row[cNom], row[cDOB], row[cGen], row[cStat],
      row[cPh], row[cPid], row[cPif], row[cCas]
    ]);
  });

  if (out.length) {
    shL.getRange(2,1,out.length, header.length).setValues(out);
  }

  return {written: out.length, sheet: sheetNameLocal};
}
function buildSeasonPassportSet_(ss){
  var set = new Set();

  function collect_(sheetName){
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    var vals = sh.getDataRange().getValues();
    var header = vals[0];
    var ciP = header.indexOf('Passeport #');
    if (ciP < 0) return;
    for (var r=1; r<vals.length; r++){
      var p = normalizePassportPlain8_(vals[r][ciP]);
      if (p) set.add(p);
    }
  }

  collect_('INSCRIPTIONS');
  collect_('INSCRIPTIONS_ENTRAINEURS');
  return set;
}


/** Upsert depuis le classeur central pour une liste de passeports
 *  - On lit la ligne centrale si elle existe et on remplace/insère la ligne locale correspondante
 *  - Conserve l’entête locale utilisée par ton subset (Passeport, Prenom, Nom, DOB, Genre, PhotoExpireLe, PhotoInvalideFlag, PhotoInvalideDuesLe, CasierExpiré, CasierExpireFlag)
 */
function upsertMembresFromCentralByPassports_(passports) {
  if (!passports || passports.length === 0) return { upserted: 0, missing: 0 };

  var seasonId = getSeasonId_();
  var ssSeason = SpreadsheetApp.openById(seasonId);
  var globalId = readParamValue('GLOBAL_MEMBRES_SHEET_ID');
  if (!globalId) throw new Error('GLOBAL_MEMBRES_SHEET_ID manquant dans PARAMS.');
  var ssGlobal = SpreadsheetApp.openById(globalId);

  var sheetNameLocal  = readParamValue('SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL';
  var sheetNameGlobal = 'MEMBRES_GLOBAL';

  var shG = ssGlobal.getSheetByName(sheetNameGlobal);
  if (!shG || shG.getLastRow() < 2) throw new Error('MEMBRES_GLOBAL introuvable dans le classeur central.');

  var VG = shG.getDataRange().getValues();
  var HG = VG[0];
  function colG_(name){ return HG.indexOf(name); }

  var gPass = colG_('Passeport'),
      gPre  = colG_('Prenom'),
      gNom  = colG_('Nom'),
      gDOB  = colG_('DateNaissance'),
      gGen  = colG_('Genre'),
      gPh   = colG_('PhotoExpireLe'),
      gCas  = colG_('CasierExpiré');

  if (gPass < 0) throw new Error("Colonne 'Passeport' manquante dans le global.");

  // Index global par passeport
  var mapG = {};
  for (var r=1; r<VG.length; r++){
    var p = normalizePassportPlain8_(VG[r][gPass]);
    if (p) mapG[p] = VG[r];
  }

  // Prépare la feuille locale (créée si absente)
  var shL = ssSeason.getSheetByName(sheetNameLocal);
  if (!shL) {
    shL = ssSeason.insertSheet(sheetNameLocal);
    shL.getRange(1,1,1,10).setValues([[
      'Passeport','Prenom','Nom','DateNaissance','Genre',
      'PhotoExpireLe','PhotoInvalideFlag','PhotoInvalideDuesLe',
      'CasierExpiré','CasierExpireFlag'
    ]]);
  }
  var VL = shL.getDataRange().getValues();
  var HL = VL[0];
  function colL_(name){ return HL.indexOf(name); }

  var lPass = colL_('Passeport'),
      lPre  = colL_('Prenom'),
      lNom  = colL_('Nom'),
      lDOB  = colL_('DateNaissance'),
      lGen  = colL_('Genre'),
      lPh   = colL_('PhotoExpireLe'),
      lPif  = colL_('PhotoInvalideFlag'),
      lPid  = colL_('PhotoInvalideDuesLe'),
      lCas  = colL_('CasierExpiré'),
      lCfl  = colL_('CasierExpireFlag');

  // Index local par passeport → N° de ligne (1-based)
  var rowIndexLocal = {};
  for (var r2=1; r2<VL.length; r2++){
    var p2 = normalizePassportPlain8_(VL[r2][lPass]);
    if (p2) rowIndexLocal[p2] = r2 + 1;
  }

  var seasonYear = Number(readParamValue('SEASON_YEAR') || new Date().getFullYear());
  var mmdd = (readParamValue('PHOTO_INVALID_FROM_MMDD') || '04-01').trim();
  var dueDate = seasonYear + '-' + mmdd;
  var cutoffNextJan1 = (seasonYear + 1) + '-01-01';

  var upserted = 0, missing = 0;
  passports.forEach(function(pp){
    var p = normalizePassportPlain8_(pp);
    if (!p) return;
    var rowG = mapG[p];
    if (!rowG) { missing++; return; }

    var prenom = String(rowG[gPre]||''),
        nom    = String(rowG[gNom]||''),
        dob    = String(rowG[gDOB]||''),
        gen    = String(rowG[gGen]||''),
        photo  = String(rowG[gPh] || ''),
        casExp = Number(rowG[gCas]||0);

    var pFlag  = (photo && photo < cutoffNextJan1) ? 1 : 0;
    var pDue   = pFlag ? dueDate : '';
    var cFlag  = (casExp === 1) ? 1 : 0;

    var rowValues = [ p, prenom, nom, dob, gen, photo, pFlag, pDue, casExp, cFlag ];

    var targetRow = rowIndexLocal[p];
    if (targetRow) {
      shL.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      shL.appendRow(rowValues);
      rowIndexLocal[p] = shL.getLastRow();
    }
    upserted++;
  });

  return { upserted: upserted, missing: missing };
}

/** Job planifié : rafraîchir MEMBRES_GLOBAL saison depuis le central (full subset) */
function nightlyRefreshMembresGlobalSubset(){
  var seasonId = getSeasonId_();
  var centralId = readParamValue('GLOBAL_MEMBRES_SHEET_ID');
  if (!centralId) throw new Error('GLOBAL_MEMBRES_SHEET_ID manquant dans PARAMS.');

  try {
    var res = syncMembresGlobalSubsetFromCentral_(seasonId, centralId);
    appendImportLog_({ type:'NIGHTLY_MEMBRES_GLOBAL_SUBSET_OK', details: res });
    return res;
  } catch(e) {
    appendImportLog_({ type:'NIGHTLY_MEMBRES_GLOBAL_SUBSET_FAIL', details: { error: String(e) } });
    throw e;
  }
}



/* === Helpers spécifiques Validation_Membres === */

/***** === Helpers d’ouverture/convert .xlsx → Google Sheet (autonomes) === *****/

// Conversion .xlsx -> Google Sheets (compatible Drive Advanced v2 et v3).
// Nécessite le service avancé "Drive API" activé (Éditeur > Services avancés de Google > Drive API).
function _vm_convertXlsxBlobToSpreadsheet_(blob, name) {
  var filename = (name || 'TMP_Import') + '';

  // v3: Drive.Files.create(resource, mediaData)
  if (typeof Drive !== 'undefined' &&
      Drive.Files &&
      typeof Drive.Files.create === 'function') {
    var resourceV3 = { name: filename, mimeType: MimeType.GOOGLE_SHEETS };
    var fileV3 = Drive.Files.create(resourceV3, blob);
    return SpreadsheetApp.openById(fileV3.id);
  }

  // v2: Drive.Files.insert(resource, mediaData)
  if (typeof Drive !== 'undefined' &&
      Drive.Files &&
      typeof Drive.Files.insert === 'function') {
    var resourceV2 = { title: filename, mimeType: MimeType.GOOGLE_SHEETS };
    var fileV2 = Drive.Files.insert(resourceV2, blob);
    return SpreadsheetApp.openById(fileV2.id);
  }

  // Service pas dispo -> message explicite
  throw new Error('Le service avancé Drive n’est pas activé (ou API non disponible). ' +
                  'Active Drive API dans "Services avancés de Google" et réessaie.');
}

// Ouvre un Spreadsheet à partir d’un Blob .xlsx OU d’un fichier Drive.
// - Si Google Sheet: ouvre direct
// - Si .xlsx: convertit via _vm_convertXlsxBlobToSpreadsheet_
function _vm_ensureSpreadsheetFromFile_(srcFileOrBlob, optName) {
  try {
    // 1) Fichier Drive ? Ouvre direct si déjà un Google Sheet.
    if (srcFileOrBlob && typeof srcFileOrBlob.getId === 'function') {
      var f = srcFileOrBlob;
      var ct = (f.getMimeType && f.getMimeType()) || (f.getBlob && f.getBlob().getContentType()) || '';
      if (String(ct).indexOf('application/vnd.google-apps.spreadsheet') >= 0) {
        return SpreadsheetApp.openById(f.getId());
      }
      // sinon convertit le blob
      var b = f.getBlob();
      return _vm_convertXlsxBlobToSpreadsheet_(b, f.getName());
    }
  } catch (e) {
    // on tentera la voie blob
  }

  // 2) Blob ?
  if (srcFileOrBlob && typeof srcFileOrBlob.getContentType === 'function') {
    var blob = srcFileOrBlob;
    var name = optName || 'TMP_Import.xlsx';
    var ct2  = (blob.getContentType && blob.getContentType()) || '';

    if (String(ct2).indexOf('spreadsheetml') >= 0 || /\.xlsx$/i.test(name)) {
      return _vm_convertXlsxBlobToSpreadsheet_(blob, name);
    }

    // fallback: créer un fichier, puis ouvrir (rare)
    var tmp = DriveApp.createFile(blob);
    try {
      return SpreadsheetApp.openById(tmp.getId());
    } finally {
      try { tmp.setTrashed(true); } catch (_) {}
    }
  }

  throw new Error('Format non supporté pour _vm_ensureSpreadsheetFromFile_.');
}


function _vm_findLatestActiveFile_(folderId) {
  const folder = DriveApp.getFolderById(folderId);
  // Exclure sous-dossier "Archives"
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    files.push(f);
  }
  if (!files.length) return { file: null, source: null };
  // Choisir le plus récent par date de dernière mise à jour
  files.sort((a,b) => b.getLastUpdated() - a.getLastUpdated());
  const file = files[0];
  return { file, source: 'folder' };
}

function _vm_readFileAsSheetValues_(file) {
  // Ouvre le fichier Drive (Google Sheet ou .xlsx) et retourne les valeurs de la 1re feuille.
  // Requiert les helpers _vm_ensureSpreadsheetFromFile_ et _vm_convertXlsxBlobToSpreadsheet_ si .xlsx.
  if (!file) throw new Error('Aucun fichier fourni à _vm_readFileAsSheetValues_.');

  var blob, name, contentType;
  try {
    blob = file.getBlob();
    name = file.getName ? file.getName() : 'Fichier';
    contentType = (blob && blob.getContentType && blob.getContentType().toLowerCase()) || '';
  } catch (e) {
    // Si jamais getBlob() échoue, on tente l’ouverture directe
    contentType = '';
    name = (file && file.getName && file.getName()) || 'Fichier';
  }

  // Cas .xlsx (Office Open XML)
  if (contentType.indexOf('spreadsheetml') >= 0 || /\.xlsx$/i.test(String(name))) {
    var ssTmp = _vm_ensureSpreadsheetFromFile_(blob || file, name); // conversion si besoin
    var sh1 = ssTmp.getSheets()[0];
    var rng1 = sh1.getDataRange();
    return rng1 ? rng1.getValues() : [];
  }

  // Cas Google Sheet natif
  try {
    var ss = SpreadsheetApp.openById(file.getId());
    var sh = ss.getSheets()[0];
    var rng = sh.getDataRange();
    return rng ? rng.getValues() : [];
  } catch (e) {
    // Dernier recours : si on n’a pas pu ouvrir par Id, on tente via conversion
    var ssFallback = _vm_ensureSpreadsheetFromFile_(blob || file, name);
    var shFallback = ssFallback.getSheets()[0];
    var rngFallback = shFallback.getDataRange();
    return rngFallback ? rngFallback.getValues() : [];
  }
}

function _vm_indexCols_(headers, names) {
  const map = {};
  names.forEach(n => {
    const idx = headers.findIndex(h => String(h).trim().toLowerCase() === String(n).trim().toLowerCase());
    if (idx < 0) throw new Error('Colonne introuvable: ' + n);
    map[n] = idx;
  });
  return map;
}

function _vm_toISODate_(val) {
  if (!val) return '';
  try {
    if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const s = String(val).trim().replace(/\//g,'-');
    // formats possibles: yyyy-mm-dd, dd-mm-yyyy, etc. — on laisse simple
    const d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e){}
  return '';
}

function _vm_to01_(v) {
  const s = String(v).trim().toLowerCase();
  if (s === '1' || s === 'true' || s === 'oui' || s === 'yes') return 1;
  return 0;
}

function _vm_hashRow_(obj) {
  const payload = [
    obj.Passeport, obj.Prenom, obj.Nom, obj.DateNaissance, obj.Genre, obj.StatutMembre,
    obj.PhotoExpireLe, obj.PhotoInvalideDuesLe, obj.PhotoInvalide,
    obj.CasierExpire, obj.SeasonYear
  ].map(x => String(x ?? '')).join('|');
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload))
         .substring(0, 22); // raccourci
}

function _vm_nowISO_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function _vm_rowValues_(order, obj, nowISO) {
  const map = {
    'Passeport': obj.Passeport,
    'Prenom': obj.Prenom,
    'Nom': obj.Nom,
    'DateNaissance': obj.DateNaissance,
    'Genre': obj.Genre,
    'StatutMembre': obj.StatutMembre,
    'PhotoExpireLe': obj.PhotoExpireLe,
    'PhotoInvalideDuesLe': obj.PhotoInvalideDuesLe,
    'PhotoInvalide': obj.PhotoInvalide,
    'CasierExpiré': obj.CasierExpire,
    'SeasonYear': obj.SeasonYear,
    'RowHash': obj.RowHash,
    'LastUpdate': nowISO
  };
  return order.map(k => map[k] ?? '');
}

function _vm_ensureHeader_(sh, order) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,order.length).setValues([order]);
    return;
  }
  const hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const same = hdr.length === order.length && hdr.every((h,i)=>String(h)===order[i]);
  if (!same) {
    sh.clear();
    sh.getRange(1,1,1,order.length).setValues([order]);
  }
}

function _vm_groupRuns_(pairs /* [rowIndex, rowValues] trié par rowIndex */) {
  const runs = [];
  let cur = [];
  for (let i=0;i<pairs.length;i++){
    if (!cur.length) { cur.push(pairs[i]); continue; }
    const prevRow = cur[cur.length-1][0];
    const thisRow = pairs[i][0];
    if (thisRow === prevRow + 1) cur.push(pairs[i]);
    else { runs.push(cur); cur = [pairs[i]]; }
  }
  if (cur.length) runs.push(cur);
  return runs;
}

function _vm_notifyImportMembres_(ss, file, stats) {
  try {
    const to = readParam_(ss, 'MAIL_ON_IMPORT_MEMBRES');
    if (!to) return;
    const subject = 'Validation_Membres — Import complété';
    const body = `Fichier: ${file.getName()}\ncreated=${stats.created}\nupdated=${stats.updated}\nunchanged=${stats.unchanged}`;
    MailApp.sendEmail(to, subject, body);
  } catch(e){
    log_('VM_NOTIFY_FAIL', String(e));
  }
}

function apiListJoueursPage(opt){
  opt = opt || {};
  var offset = Number(opt.offset||0), limit = Math.min(Number(opt.limit||50), 200);
  var search = String(opt.search||'').trim().toLowerCase();
  var band   = String(opt.band||'').toUpperCase();    // U4-U8 | U9-U12 | U13-U18 | Adulte | ''
  var adapte = String(opt.adapte||'').toLowerCase();  // '1'|'0'|''
  var photo  = String(opt.photo||'').toLowerCase();   // 'missing'|'expired'|'soon'|''

  var ss  = getSeasonSpreadsheet_(getSeasonId_());
  var tab = readSheetAsObjects_(ss.getId(), 'JOUEURS'); // {header, rows}
  var H = tab.header||[], rows = tab.rows||[];

  var iPass = H.indexOf('Passeport #');
  var iNom  = H.indexOf('Nom'), iPre = H.indexOf('Prénom');
  var iCour = H.indexOf('Courriels'), iBand = H.indexOf('AgeBracket')>=0?H.indexOf('AgeBracket'):H.indexOf('ProgramBand');
  var iAda  = H.indexOf('isAdapte'), iPhoStr = H.indexOf('PhotoStr'), iPhoDate = H.indexOf('PhotoExpireLe');

  function match(r){
    // search nom/prenom/passeport/courriel
    if (search){
      var blob = (String(r[iPass]||'')+' '+String(r[iNom]||'')+' '+String(r[iPre]||'')+' '+String(r[iCour]||'')).toLowerCase();
      if (blob.indexOf(search) === -1) return false;
    }
    if (band && String(r[iBand]||'').toUpperCase() !== band) return false;
    if (adapte){
      var a = String(r[iAda]||'').toLowerCase();
      var yes = (a==='1'||a==='true'||a==='oui');
      if (adapte==='1' && !yes) return false;
      if (adapte==='0' && yes)  return false;
    }
    if (photo){
      var s = String(r[iPhoStr]||'').toLowerCase();
      var d = String(r[iPhoDate]||'');
      if (photo==='missing' && !(s.indexOf('aucune')!==-1 || (!d && !s))) return false;
      if (photo==='expired' && s.indexOf('expir')===-1) return false;
      if (photo==='soon'    && s.indexOf('bient')===-1) return false;
    }
    return true;
  }

  var all = rows.filter(match);
  var page = all.slice(offset, offset+limit).map(function(r){
    return {
      passeport: r[iPass]||'',
      nom: r[iNom]||'',
      prenom: r[iPre]||'',
      courriels: r[iCour]||'',
      band: r[iBand]||'',
      adapte: r[iAda]||'',
      photo: r[iPhoStr]||'',
      photoDate: r[iPhoDate]||''
    };
  });

  return { total: all.length, offset: offset, limit: limit, rows: page };
}

function apiGetJoueurDetail(passeport){
  var ss = getSeasonSpreadsheet_(getSeasonId_());
  var J = readSheetAsObjects_(ss.getId(), 'JOUEURS');
  var L = readSheetAsObjects_(ss.getId(), 'ACHATS_LEDGER');
  function norm(p){ return String(p||'').replace(/\D/g,'').padStart(8,'0'); }
  var p8 = norm(passeport);

  var H=J.header||[], R=J.rows||[];
  var idx = {
    pass: H.indexOf('Passeport #'), nom:H.indexOf('Nom'), pre:H.indexOf('Prénom'),
    dob:H.indexOf('DateNaissance'), genre:H.indexOf('Genre'), cour:H.indexOf('Courriels'),
    band: H.indexOf('AgeBracket')>=0?H.indexOf('AgeBracket'):H.indexOf('ProgramBand'),
    adapte:H.indexOf('isAdapte'), photo:H.indexOf('PhotoStr'), photoDate:H.indexOf('PhotoExpireLe')
  };
  var rec = null;
  for (var i=0;i<R.length;i++){
    var rp = norm(R[i][idx.pass]||''); if (rp===p8){ rec = R[i]; break; }
  }
  if (!rec) return { notFound:true };

  var detail = {
    passeport: rec[idx.pass]||'',
    nom: rec[idx.nom]||'',
    prenom: rec[idx.pre]||'',
    dob: rec[idx.dob]||'',
    genre: rec[idx.genre]||'',
    courriels: rec[idx.cour]||'',
    band: rec[idx.band]||'',
    adapte: rec[idx.adapte]||'',
    photo: rec[idx.photo]||'',
    photoDate: rec[idx.photoDate]||''
  };

  // activités actives, saison courante, non ignorées
  var saison = readParam_(ss, 'SEASON_LABEL')||'';
  var acts = (L.rows||[]).filter(function(r){
    if (String(r['Saison']||'')!==saison) return false;
    if ((+r['Status']||0)!==1) return false;
    if ((+r['isIgnored']||0)===1) return false;
    return norm(r['Passeport #']||r['Passeport']||'')===p8;
  }).map(function(r){
    return {
      type: r['Type'],
      nom: r['NomFrais'] || r['Nom du frais'] || r['Frais'] || r['Produit'] || '',
      tags: r['Tags'] || '',
      band: r['ProgramBand'] || ''
    };
  });

  return { joueur: detail, activites: acts };
}
