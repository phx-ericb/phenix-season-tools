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

/** Import principal: lit le dernier fichier du dossier Validation_Membres et upsert dans SHEET_MEMBRES_GLOBAL */
function importValidationMembresToGlobal_(seasonId) {
    const ss = SpreadsheetApp.openById(seasonId);
  const folderId = readParam_(ss, 'DRIVE_FOLDER_VALIDATION_MEMBRES');
  if (!folderId) throw new Error('Paramètre DRIVE_FOLDER_VALIDATION_MEMBRES manquant.');
  const seasonYear = Number(readParam_(ss, 'SEASON_YEAR') || new Date().getFullYear());
  const photoInvalidFrom = (readParam_(ss, 'PHOTO_INVALID_FROM_MMDD') || '04-01').trim();

  const { file, source } = _vm_findLatestActiveFile_(folderId);
  if (!file) {
    log_('VM_IMPORT_NOFILE', 'Aucun fichier actif dans le dossier Validation_Membres.');
    return { updated: 0, unchanged: 0, created: 0 };
  }

  log_('VM_IMPORT_START', `Lecture: ${file.getName()} (${file.getId()})`);

  const values = _vm_readFileAsSheetValues_(file);
  if (!values || values.length < 2) throw new Error('Fichier Validation_Membres vide ou illisible.');

  const headers = values[0].map(h => String(h).trim());
  const idx = _vm_indexCols_(headers, [
    'Passeport #', 'Prénom', 'Nom', 'Date de naissance',
    'Identité de genre', 'Statut du membre',
    'Date d\'expiration de la photo de profil',
    'Vérification du casier judiciaire est expiré'
  ]);

  const seasonInvalidDate = `${seasonYear}-${photoInvalidFrom}`; // ex 2025-04-01

  // Build target rows keyed by passport
  const targetByPassport = new Map();
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row || row.length === 0) continue;

  const passportRaw = row[idx['Passeport #']];
const passport = normalizePassportPlain8_(passportRaw); // utilise ta fonction existante
if (!passport) continue;


    const prenom   = String(row[idx['Prénom']] || '').trim();
    const nom      = String(row[idx['Nom']] || '').trim();
    const dobRaw   = row[idx['Date de naissance']];
    const genre    = String(row[idx['Identité de genre']] || '').trim();
    const statut   = String(row[idx['Statut du membre']] || '').trim();
    const photoExp = _vm_toISODate_(row[idx['Date d\'expiration de la photo de profil']]);
    const casier   = _vm_to01_(row[idx['Vérification du casier judiciaire est expiré']]); // 1 si expiré

// invalide si la photo expire avant le 1er janvier de l'année suivante
const cutoffNextJan1 = (seasonYear + 1) + '-01-01'; // ex.: 2026-01-01
const photoInvalide = (!photoExp || photoExp < cutoffNextJan1) ? 1 : 0;
const photoDuesLe   = photoInvalide ? seasonInvalidDate : '';

    // Base row object
    const obj = {
      Passeport: passport,
      Prenom: prenom,
      Nom: nom,
      DateNaissance: _vm_toISODate_(dobRaw),
      Genre: genre,
      StatutMembre: statut,
      PhotoExpireLe: photoExp,
      PhotoInvalideDuesLe: photoDuesLe,
      PhotoInvalide: photoInvalide,
      CasierExpire: casier,
      SeasonYear: seasonYear
    };

    const rowHash = _vm_hashRow_(obj);
    obj.RowHash = rowHash;

    targetByPassport.set(passport, obj);
  }

  // Upsert dans MEMBRES_GLOBAL avec diff sur RowHash
  const sheetName = readParam_(ss, 'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL';
const sh = getSheetOrCreate_(ss, sheetName);
  const colOrder = [
    'Passeport','Prenom','Nom','DateNaissance','Genre','StatutMembre',
    'PhotoExpireLe','PhotoInvalideDuesLe','PhotoInvalide',
    'CasierExpiré','SeasonYear','RowHash','LastUpdate'
  ];
  _vm_ensureHeader_(sh, colOrder);

  const data = sh.getDataRange().getValues();
  const header = data[0];
  const colIdx = {};
  header.forEach((h, i) => colIdx[String(h)] = i);

  // index existant: passeport -> {rowIndex, RowHash}
  const existingIdx = new Map();
  for (let r = 1; r < data.length; r++) {
    const pass = String(data[r][colIdx['Passeport']] || '').trim();
    if (pass) {
      existingIdx.set(pass, {
        rowIndex: r,
        hash: String(data[r][colIdx['RowHash']] || '').trim()
      });
    }
  }

  const nowISO = _vm_nowISO_();

  const toWrite = [];     // [rowIndex, rowValues]
  const toAppend = [];    // [rowValues]
  let updated = 0, created = 0, unchanged = 0;

  targetByPassport.forEach((obj, pass) => {
    if (existingIdx.has(pass)) {
      const { rowIndex, hash } = existingIdx.get(pass);
      if (hash === obj.RowHash) {
        unchanged++;
        return; // rien à faire
      }
      // update ligne existante
      const rowVals = _vm_rowValues_(colOrder, obj, nowISO);
      toWrite.push([rowIndex+1, rowVals]); // +1 car getRange est 1-based avec header
      updated++;
    } else {
      // append nouvelle ligne
      const rowVals = _vm_rowValues_(colOrder, obj, nowISO);
      toAppend.push(rowVals);
      created++;
    }
  });

  // Écritures batchées
  if (toWrite.length) {
    // Grouper par blocs contigus? Ici simple: on écrit une par une en Range.setValues (coûteux si 1000+)
    // Optimisation: trier par rowIndex et regrouper par runs contigus
    toWrite.sort((a,b) => a[0]-b[0]);
    const runs = _vm_groupRuns_(toWrite);
    runs.forEach(run => {
      const startRow = run[0][0];
      const block = run.map(x => x[1]);
      sh.getRange(startRow, 1, block.length, block[0].length).setValues(block);
    });
  }

if (toAppend.length) {
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, toAppend.length, colOrder.length).setValues(toAppend);
}


  log_('VM_IMPORT_SUMMARY', `created=${created}, updated=${updated}, unchanged=${unchanged}, total=${targetByPassport.size}`);
  _vm_notifyImportMembres_(ss, file, {created, updated, unchanged});
  return {created, updated, unchanged};
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
