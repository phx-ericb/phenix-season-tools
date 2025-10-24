/***** utils.js — révisé (robuste ID/URL/DriveFile/Spreadsheet) *****/

/** Définis une saison active (script properties) */
function setSeasonIdGlobalOnce() {
  var id = '1IVVHi17Jyo8jvWtrSuenbPW8IyEZqlY1bXx-WbnXPkk'; // ton ID
  var p = PropertiesService.getScriptProperties();
  p.setProperty('ACTIVE_SEASON_ID', id);
  p.setProperty('PHENIX_SEASON_SHEET_ID', id); // alias lu par la lib
  p.setProperty('SEASON_SPREADSHEET_ID', id);  // alias supplémentaire
}

/* -------------------- Helpers robustes Spreadsheet -------------------- */

if (typeof extractSpreadsheetId_ !== 'function') {
  /** Extrait un ID depuis un ID pur OU une URL Google Sheets */
  function extractSpreadsheetId_(s) {
    var m = String(s || '').match(/[-\w]{25,}/);
    if (m && m[0]) return m[0];
    throw new Error('extractSpreadsheetId_: ID/URL invalide: ' + s);
  }
}

if (typeof _debugType_ !== 'function') {
  /** Petit helper pour logger le type d’input passé aux utils */
  function _debugType_(x) {
    if (x == null) return 'null/undefined';
    if (typeof x === 'string') return 'string';
    if (typeof x.getSheetByName === 'function') return 'Spreadsheet';
    if (typeof x.getId === 'function') return 'DriveFile(id=' + x.getId() + ')';
    return Object.prototype.toString.call(x);
  }
}

/** Garantit un Spreadsheet (accepte Spreadsheet | id|url string | DriveFile | rien) */
function ensureSpreadsheet_(ssOrId) {
  // 1) Déjà un Spreadsheet ?
  if (ssOrId && typeof ssOrId.getSheetByName === 'function' && typeof ssOrId.getId === 'function') {
    return ssOrId;
  }
  // 2) DriveFile ?
  if (ssOrId && typeof ssOrId.getId === 'function' && typeof ssOrId.getBlob === 'function') {
    return SpreadsheetApp.openById(ssOrId.getId());
  }
  // 3) String: ID pur ou URL
  if (typeof ssOrId === 'string' && ssOrId.trim()) {
    return SpreadsheetApp.openById(extractSpreadsheetId_(ssOrId));
  }

  // 4) Rien de fourni → essaie les Script Properties (IDs de saison)
  try {
    var props = PropertiesService.getScriptProperties();
    var pid =
      (props && (props.getProperty('ACTIVE_SEASON_ID')
              ||  props.getProperty('PHENIX_SEASON_SHEET_ID')
              ||  props.getProperty('SEASON_SPREADSHEET_ID'))) || '';
    if (pid) return SpreadsheetApp.openById(pid);
  } catch (_){}

  // 5) Dernier filet: classeur actif — et si absent, on lève une erreur claire
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (_){}

  throw new Error('ensureSpreadsheet_: impossible d’obtenir un Spreadsheet valide (input=' + _debugType_(ssOrId) + ')');
}

/* ---------- DROP-IN: assertions & shims Spreadsheet ---------- */

// log neutre si absent
if (typeof log_ !== 'function') {
  function log_(action, details) {
    try {
      var ss = (typeof getSeasonSpreadsheet_ === 'function') ? getSeasonSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
      if (ss && typeof appendImportLog_ === 'function') {
        appendImportLog_(ss, String(action || 'LOG'), (typeof details === 'string' ? details : JSON.stringify(details || {})));
      }
    } catch (_){}
  }
}

/** Vérifie/convertit en Spreadsheet et *garantit* un objet valide, sinon jette une erreur claire. */
if (typeof assertSpreadsheet_ !== 'function') {
  function assertSpreadsheet_(ssOrId, contextMsg) {
    var ss = ensureSpreadsheet_(ssOrId);
    if (!ss || typeof ss.getSheetByName !== 'function') {
      throw new Error('assertSpreadsheet_: Spreadsheet invalide' + (contextMsg ? ' (' + contextMsg + ')' : ''));
    }
    return ss;
  }
}

/** Variante pratique : renvoie directement une feuille, jette si absente. */
if (typeof assertSheet_ !== 'function') {
  function assertSheet_(ssOrId, sheetName, contextMsg) {
    var ss = assertSpreadsheet_(ssOrId, contextMsg);
    var sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error('assertSheet_: feuille "' + sheetName + '" introuvable' + (contextMsg ? ' (' + contextMsg + ')' : ''));
    return sh;
  }
}

/** Sécurisée: essaye d’ouvrir une feuille; si absente et header fourni → la crée. */
if (typeof getSheetOrCreate_ !== 'function') {
  function getSheetOrCreate_(ssOrId, name, header) {
    var ss = assertSpreadsheet_(ssOrId, 'getSheetOrCreate_');
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


/** Ouvre le classeur de la saison (robuste sur Spreadsheet/DriveFile/URL/ID) */
function getSeasonSpreadsheet_(seasonSheetLike) {
  // Si on nous donne quelque chose (Spreadsheet/DriveFile/URL/ID), on laisse ensureSpreadsheet_ gérer.
  if (seasonSheetLike) return ensureSpreadsheet_(seasonSheetLike);

  // Sinon, on tente via Script Properties
  var props = PropertiesService.getScriptProperties();
  var id =
    (props && (props.getProperty('ACTIVE_SEASON_ID')
            ||  props.getProperty('PHENIX_SEASON_SHEET_ID')
            ||  props.getProperty('SEASON_SPREADSHEET_ID'))) || '';
  if (id) return ensureSpreadsheet_(id);

  // Puis l’actif si présent
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (_){}

  throw new Error("getSeasonSpreadsheet_: seasonSheetId manquant. Passe un Spreadsheet/ID/URL, ou définis ACTIVE_SEASON_ID/PHENIX_SEASON_SHEET_ID.");
}

/* -------------------- Sheets utils -------------------- */

/** Récupère une feuille, ou la crée vide avec un header si fourni */
function getSheetOrCreate_(ss, name, header) {
  ss = ensureSpreadsheet_(ss);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header && header.length) {
      sh.getRange(1,1,1,header.length).setValues([header]);
    }
  } else if (header && header.length && sh.getLastRow() === 0) {
    sh.getRange(1,1,1,header.length).setValues([header]);
  }
  return sh;
}

/** Assure les feuilles cœur (et leurs entêtes) */
function ensureCoreSheets_(ss) {
  ss = ensureSpreadsheet_(ss);
  getSheetOrCreate_(ss, SHEETS.PARAMS);
  getSheetOrCreate_(ss, SHEETS.IMPORT_LOG, ['Horodatage','Action','Détails']);
  ensureMailOutbox_(ss);
    // ERREURS enrichi v0.7
  getSheetOrCreate_(ss, SHEETS.ERREURS, ['Passeport','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']);
  // Annulations (avec colonnes "actifs restants")
  getSheetOrCreate_(ss, SHEETS.ANNULATIONS_INSCRIPTIONS, ['Horodatage','Passeport','Nom','Prénom','NomComplet','Saison','Frais','DateAnnulation','A_ENCORE_ACTIF','ACTIFS_RESTANTS']);
  getSheetOrCreate_(ss, SHEETS.ANNULATIONS_ARTICLES,     ['Horodatage','Passeport','Nom','Prénom','NomComplet','Saison','Frais','DateAnnulation','A_ENCORE_ACTIF','ACTIFS_RESTANTS']);
  // Modifs
}

/** Lit un paramètre depuis PARAMS (fallback DocumentProperties) */
function readParam_(ss, key) {
  ss = ensureSpreadsheet_(ss);
  var sh = ss.getSheetByName(SHEETS.PARAMS);
  if (sh) {
    var last = sh.getLastRow();
    if (last >= 1) {
      var data = sh.getRange(1,1,last,2).getValues();
      for (var i=0;i<data.length;i++) {
        if ((data[i][0]+'').trim() === key) {
          return (data[i][1]+'').trim();
        }
      }
    }
  }
  var props = PropertiesService.getDocumentProperties();
  return (props.getProperty(key) || '').trim();
}

/** Petit logger d’import */
function appendImportLog_(ss, action, details) {
  ss = ensureSpreadsheet_(ss);
  var sh = getSheetOrCreate_(ss, SHEETS.IMPORT_LOG, ['Horodatage','Action','Détails']);
  sh.appendRow([new Date(), action, details || '']);
}

/** Copie le tableau values dans la feuille cible (full refresh), en conservant les dimensions. */
function overwriteSheet_(sheet, values) {
  sheet.clearContents();
  if (!values || !values.length) return;
  sheet.getRange(1,1,values.length, values[0].length).setValues(values);
}

/** Essaie de forcer la 1ère colonne en texte (utile pour passeports, leading zeros) */
function prefixFirstColWithApostrophe_(values) {
  if (!values || !values.length) return values;
  for (var r=1; r<values.length; r++) { // commence à 1 pour ne pas toucher l'entête
    if (values[r][0] !== '' && values[r][0] != null) {
      values[r][0] = "'" + values[r][0];
    }
  }
  return values;
}

/** Crée (si absent) et renvoie le sous-dossier d’archive daté sous le dossier d’imports */
function ensureArchiveSubfolder_(importsFolderId) {
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), ARCHIVE.DATE_FMT);
  var importsFolder = DriveApp.getFolderById(importsFolderId);
  var it = importsFolder.getFoldersByName(ARCHIVE.ROOT_SUBFOLDER);
  var archives = it.hasNext() ? it.next() : importsFolder.createFolder(ARCHIVE.ROOT_SUBFOLDER);
  var it2 = archives.getFoldersByName(dateStr);
  var dated = it2.hasNext() ? it2.next() : archives.createFolder(dateStr);
  return dated;
}

/** Helper simple : assure un sous-dossier (par nom) sous un parent DriveApp.Folder */
function ensureSubfolderUnder_(parentFolder, name) {
  var it = parentFolder.getFoldersByName(name);
  return it.hasNext() ? it.next() : parentFolder.createFolder(name);
}

/** Déplace un fichier vers un dossier (Drive v3), compatible Shared Drives. — FIX mediaData=null */
function moveFileIdToFolderIdV3_(fileId, targetFolderId) {
  if (!(Drive && Drive.Files && typeof Drive.Files.update === 'function')) {
    throw new Error('Drive v3 indisponible (Services avancés).');
  }
  var meta = Drive.Files.get(fileId, { supportsAllDrives: true });
  var parents = (meta.parents || []).map(function(p){ return (typeof p === 'string') ? p : p.id; });
  var removeArg = parents.join(',');
  Drive.Files.update(
    {},                  // resource
    fileId,
    null,                // mediaData: null pour une update sans upload
    { addParents: targetFolderId, removeParents: removeArg, supportsAllDrives: true }
  );
}

/** Lit une feuille en objets [{col:val,...}] + métadonnées.
 *  Utilise getDisplayValues() pour préserver les zéros en tête (ex.: Passeport #).
 */
function readSheetAsObjects_(ssId, sheetName) {
  var sh = SpreadsheetApp.openById(extractSpreadsheetId_(ssId)).getSheetByName(sheetName);
  if (!sh) return { headers: [], rows: [], sheet: null, lastRow: 0, lastCol: 0 };

  var rng = sh.getDataRange();
  var disp = rng.getDisplayValues(); // <- clé pour garder "00247789"
  if (!disp || !disp.length) return { headers: [], rows: [], sheet: sh, lastRow: 0, lastCol: 0 };

  var headers = disp[0].map(function(h){ return String(h || ''); });
  var rows = [];
  for (var r = 1; r < disp.length; r++) {
    var o = {};
    for (var c = 0; c < headers.length; c++) o[headers[c]] = disp[r][c];
    rows.push(o);
  }
  return { headers: headers, rows: rows, sheet: sh, lastRow: rng.getNumRows(), lastCol: rng.getNumColumns() };
}

/** KEY_COLS depuis PARAMS (fallback raisonnable) */
function getKeyColsFromParams_(ss) {
  ss = ensureSpreadsheet_(ss);
  var raw = (readParam_(ss, PARAM_KEYS.KEY_COLS) || 'Passeport #,Saison');
  return raw.split(',').map(function(s){ return s.trim(); }).filter(String);
}

/** Construit la clé concat (ordre des KEY_COLS) */
function buildKeyFromRow_(row, keyCols) {
  return keyCols.map(function(k){ return String(row[k] == null ? '' : row[k]).trim(); }).join('||');
}

/** True si l'objet-ligne est totalement vide (toutes colonnes vides/NULL). */
function isRowCompletelyEmpty_(rowObj) {
  var keys = Object.keys(rowObj || {});
  for (var i=0;i<keys.length;i++) {
    var v = rowObj[keys[i]];
    if (v !== '' && v != null) return false;
  }
  return true;
}

/** True si toutes les colonnes de clé sont non vides (après trim). */
function isKeyComplete_(rowObj, keyCols) {
  for (var i=0;i<keyCols.length;i++) {
    var v = rowObj[keyCols[i]];
    if (v == null) return false;
    if (String(v).trim() === '') return false;
  }
  return true;
}

/** MD5 helpers (exclut colonnes de contrôle si présentes) */
function bytesToHex_(bytes) {
  var out = [];
  for (var i=0;i<bytes.length;i++) {
    var s = (bytes[i] & 0xFF).toString(16);
    out.push(s.length === 1 ? '0'+s : s);
  }
  return out.join('');
}
function computeRowHash_(row) {
  var copy = {};
  Object.keys(row).forEach(function(k){
    if (k === CONTROL_COLS.ROW_HASH
     || k === CONTROL_COLS.CANCELLED
     || k === CONTROL_COLS.EXCLUDE_FROM_EXPORT
     || k === CONTROL_COLS.LAST_MODIFIED_AT) return; // v0.7: exclut LAST_MODIFIED_AT
    copy[k] = row[k];
  });
  var str = JSON.stringify(copy);
  var md5 = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str, Utilities.Charset.UTF_8);
  return bytesToHex_(md5);
}

/** Écrit dans la feuille staging (efface avant), crée si absente.
 *  - Injecte 'Saison' si absente (depuis PARAMS!SEASON_LABEL)
 *  - Normalise 'Passeport #' en TEXTE avec zéro-padding (8) + apostrophe
 */
function writeStaging_(seasonSs, stagingSheetName, values) {
  seasonSs = ensureSpreadsheet_(seasonSs);
  var sh = getSheetOrCreate_(seasonSs, stagingSheetName);
  sh.clearContents();
  if (!values || !values.length) return;

  var headers = values[0].map(function(h){ return String(h || ''); });
  var passIdx = headers.indexOf('Passeport #');

  var hasSaison = headers.indexOf('Saison') !== -1;
  var seasonLabel = hasSaison ? '' : (readParam_(seasonSs, 'SEASON_LABEL') || '');

  // Construire l'entête finale
  var outHeaders = headers.slice();
  if (!hasSaison) outHeaders.push('Saison');

  // Construire le tableau normalisé
  var out = new Array(values.length);
  out[0] = outHeaders;

  for (var r = 1; r < values.length; r++) {
    var row = values[r].slice();

    // Normaliser Passeport # → "'00001234" (texte, longueur 8 si numérique)
    if (passIdx >= 0) {
      row[passIdx] = normalizePassportToText8_(row[passIdx]);
    }

    // Injecter Saison si absente
    if (!hasSaison) row.push(seasonLabel);

    out[r] = row;
  }

  // Écriture en bloc
  sh.getRange(1, 1, out.length, out[0].length).setValues(out);

  // Verrouiller les formats colonnes (texte)
  if (passIdx >= 0) {
    var colP = passIdx + 1;
    if (sh.getLastRow() > 1) sh.getRange(2, colP, sh.getLastRow() - 1, 1).setNumberFormat('@');
  }
  if (!hasSaison) {
    var colS = outHeaders.indexOf('Saison') + 1;
    if (sh.getLastRow() > 1) sh.getRange(2, colS, sh.getLastRow() - 1, 1).setNumberFormat('@');
  }
}

/** Normalise Passeport # en TEXTE avec padding 8 (retourne "'00123456") */
function normalizePassportToText8_(val) {
  if (val == null) return '';
  var s = String(val).trim();
  if (s === '') return '';
  if (s[0] === "'") s = s.slice(1);
  if (/^\d+$/.test(s)) {
    if (s.length < 8) s = ('00000000' + s).slice(-8);
  }
  return "'" + s;
}
/** Version "export" : chaîne "00123456" (sans apostrophe) */
function normalizePassportPlain8_(val) {
  var s = normalizePassportToText8_(val);
  return (s && s[0] === "'") ? s.slice(1) : s;
}

/* ---- Params & statut helpers ---- */
function _paramOr_(ss, key, dflt) {
  ss = ensureSpreadsheet_(ss);
  var v = readParam_(ss, key);
  return (v == null || v === '') ? String(dflt) : String(v);
}
function _norm_(s) {
  s = String(s == null ? '' : s);
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch(e) {}
  return s.toLowerCase().trim();
}
function _isCancelledStatus_(val, cancelListCsv) {
  var norm = _norm_(val);
  var list = String(cancelListCsv || '').split(',').map(function(x){return _norm_(x);}).filter(Boolean);
  return list.indexOf(norm) >= 0;
}

/* ---- MAIL_OUTBOX helpers ---- */
function getMailOutboxHeaders_() {
  return ['Type','To','Cc','Sujet','Corps','Attachments','KeyHash','Status','CreatedAt','SentAt','Error'];
}
function ensureMailOutbox_(ss) {
  ss = ensureSpreadsheet_(ss);
  var headers = getMailOutboxHeaders_();
  var sh = ss.getSheetByName(SHEETS.MAIL_OUTBOX);

  if (!sh) {
    sh = ss.insertSheet(SHEETS.MAIL_OUTBOX);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }
  var last = sh.getLastRow();

  if (last === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }

  var first = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var ok = headers.every(function (h, i) { return String(first[i] || '') === h; });
  if (!ok) {
    sh.insertRowsBefore(1, 1);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}



function setCancelFlags_(sheet, rowNumber, headers, cancelled, exclude) {
  var idxC = headers.indexOf(CONTROL_COLS.CANCELLED);
  var idxE = headers.indexOf(CONTROL_COLS.EXCLUDE_FROM_EXPORT);
  if (idxC >= 0) sheet.getRange(rowNumber, idxC + 1).setValue(!!cancelled);
  if (idxE >= 0) sheet.getRange(rowNumber, idxE + 1).setValue(!!exclude);
}

function getHeadersIndex_(sh, width) {
  var headers = sh.getRange(1,1,1, width || sh.getLastColumn()).getValues()[0].map(String);
  var idx = {};
  headers.forEach(function(h, i) { idx[h] = i + 1; }); // 1-based
  return idx;
}

/* ====== Helpers v0.7 (nouveaux) ====== */
function norm_(s) {
  s = String(s == null ? '' : s);
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch(e){}
  return s.trim();
}
function normLower_(s){ return norm_(s).toLowerCase(); }

function parseSeasonYear_(saisonLabel) {
  var m = String(saisonLabel||'').match(/(20\d{2})/);
  return m ? parseInt(m[1],10) : (new Date()).getFullYear();
}

function birthYearFromRow_(row) {
  var y = row['Année de naissance'] || row['Annee de naissance'] || row['Annee'] || '';
  if (y && /^\d{4}$/.test(String(y))) return parseInt(y,10);
  var dob = row['Date de naissance'] || row['Naissance'] || '';
  if (dob) {
    var s = String(dob);
    var m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) return parseInt(m[1],10);
    var m2 = s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
    if (m2) return parseInt(m2[3],10);
  }
  return null;
}
function computeUForYear_(birthYear, seasonYear) {
  if (!birthYear || !seasonYear) return null;
  var u = seasonYear - birthYear;
  return (u >= 4 && u <= 21) ? ('U' + u) : null;
}
function deriveUFromRow_(row) {
  var cat = row['Catégorie'] || row['Categorie'] || '';
  if (cat && /^U\d{1,2}/i.test(cat)) return cat.toUpperCase().replace(/\s+/g,'');
  var seasonYear = parseSeasonYear_(row['Saison'] || '');
  var by = birthYearFromRow_(row);
  var U = computeUForYear_(by, seasonYear);
  return U || '';
}
function deriveSectorFromRow_(row) {
  var U = deriveUFromRow_(row); // ex U10
  var n = parseInt(String(U).replace(/^U/i,''),10);
  if (!n || isNaN(n)) return 'U?';
  if (n >= 4 && n <= 8)  return 'U4-U8';
  if (n >= 9 && n <= 12) return 'U9-U12';
  if (n >= 13 && n <= 18) return 'U13-U18';
  return 'AUTRES';
}

function collectEmailsFromRow_(row, fieldsCsv) {
  var fields = (fieldsCsv && fieldsCsv.length)
    ? fieldsCsv.split(',').map(function(x){return x.trim();}).filter(Boolean)
    : ['Courriel','Parent 1 - Courriel','Parent 2 - Courriel'];
  var set = {};
  fields.forEach(function(f){
    var v = row[f];
    if (!v) return;
    String(v).split(/[;,]/).forEach(function(e){
      e = norm_(e);
      if (!e) return;
      set[e] = true;
    });
  });
  return Object.keys(set).join(',');
}

/** Liste des occurrences actives (non annulées) pour un passeport dans une feuille donnée */
function listActiveOccurrencesForPassport_(ss, sheetName, passport) {
  ss = ensureSpreadsheet_(ss);
  var info = readSheetAsObjects_(ss.getId(), sheetName);
  var act = info.rows.filter(function(r){
    return norm_(r['Passeport #']) === norm_(passport)
        && String(r[CONTROL_COLS.CANCELLED]||'').toLowerCase() !== 'true'
        && String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT]||'').toLowerCase() !== 'true';
  });
  return act.map(function(r){
    return (r['Nom du frais'] || r['Frais'] || r['Produit'] || '').toString();
  });
}

/** Conversion simple JSON → string "compact" */
function jsonCompact_(obj) {
  try { return JSON.stringify(obj); } catch(e){ return '{}'; }
}

/** Harmonise une valeur "Genre" provenant de MAPPINGS. */
function _normMapGenre(g) {
  var v = String(g == null ? '' : g).trim().toUpperCase();
  if (!v || v === '*') return '*';

  var vClean = v.replace(/\s+/g, '');
  var AS_X   = ['MIXTE','MIX','NB','NONBINAIRE','NON-BINAIRE','X'];
  var AS_ALL = ['MF','M/F','F/M','FEM/MASC','MASC/FEM'];

  if (AS_X.indexOf(vClean)   >= 0) return 'X';
  if (AS_ALL.indexOf(vClean) >= 0) return '*';
  if (vClean === 'M' || vClean === 'F') return vClean;

  return '*';
}

/* ===== Lecture MAPPINGS unifiés (incl. ExclusiveGroup) ===== */
function _loadUnifiedGroupMappings_(ss) {
  ss = ensureSpreadsheet_(ss);

  var sh = ss.getSheetByName(SHEETS.MAPPINGS) || ss.getSheetByName('MAPPINGS');
  if (!sh) return [];

  var data = sh.getDataRange().getValues();
  var headerIdx = -1;
  for (var r = 0; r < data.length; r++) {
    var row = data[r] || [];
    var nonEmpty = 0;
    for (var c = 0; c < row.length; c++) {
      if (String(row[c] || '').trim() !== '') { nonEmpty++; if (nonEmpty >= 2) break; }
    }
    if (nonEmpty >= 2) { headerIdx = r; break; }
  }
  if (headerIdx === -1) return [];

  var H = (data[headerIdx] || []).map(function(h){ return String(h || '').trim(); });
  function idx(k) { var i = H.indexOf(k); return i < 0 ? null : i; }

  var iType = idx('Type'), iAli = idx('AliasContains'), iUmin = idx('Umin'), iUmax = idx('Umax'),
      iGen  = idx('Genre'), iG   = idx('GroupeFmt'),     iC    = idx('CategorieFmt'),
      iEx   = idx('Exclude'), iPr = idx('Priority'), iX = idx('ExclusiveGroup'), iCode = idx('Code');

  if (iType == null || iAli == null) return [];

  function _t(v){ return String(v == null ? '' : v).replace(/\u00A0/g, ' ').trim(); }
  function _toIntOrNull(v){ var s=_t(v); if(!s) return null; var n=parseInt(s,10); return isNaN(n)?null:n; }

  var out = [];
  for (var r2 = headerIdx + 1; r2 < data.length; r2++) {
    var row2 = data[r2] || [];

    var isEmpty = true;
    for (var c2 = 0; c2 < row2.length; c2++) {
      if (_t(row2[c2]) !== '') { isEmpty = false; break; }
    }
    if (isEmpty) continue;

    var type = (iType == null) ? '' : String(_t(row2[iType])).toLowerCase();
    var ali  = (iAli  == null) ? '' : _t(row2[iAli]);
    var umin = (iUmin == null) ? null : _toIntOrNull(row2[iUmin]);
    var umax = (iUmax == null) ? null : _toIntOrNull(row2[iUmax]);
    var gen  = (iGen  == null) ? '*' : _normMapGenre(row2[iGen]);
    var grp  = (iG    == null) ? '' : _t(row2[iG]);
    var cat  = (iC    == null) ? '' : _t(row2[iC]);
    var ex   = (iEx   == null) ? '' : String(_t(row2[iEx])).toLowerCase();
    var pri  = (iPr   == null) ? null : _toIntOrNull(row2[iPr]);
    var exg  = (iX    == null) ? '' : _t(row2[iX]);
    var code = (iCode == null) ? '' : _t(row2[iCode]);

    out.push({
      Type: type,
      AliasContains: ali,
      Umin: umin,
      Umax: umax,
      Genre: gen,
      GroupeFmt: grp,
      CategorieFmt: cat,
      Exclude: ex === 'true',
      Priority: (pri == null ? 100 : pri),
      ExclusiveGroup: exg,
      Code: code
    });
  }

  out.sort(function(a, b) {
    if (a.Priority !== b.Priority) return b.Priority - a.Priority;
    return (b.AliasContains || '').length - (a.AliasContains || '').length;
  });

  return out;
}

/* ========= MEMBRES_GLOBAL: index & cache ========= */

var __MG_CACHE = (typeof __MG_CACHE !== 'undefined') ? __MG_CACHE : { at:0, key:'', data:null };

if (typeof normalizePassportPlain8_ !== 'function') {
  function normalizePassportPlain8_(v){
    var s = String(v == null ? '' : v).replace(/\D/g,'').trim();
    if (!s) return '';
    s = s.slice(-8);
    while (s.length < 8) s = '0' + s;
    return s;
  }
}

function _mg_cacheKey_(ss){
  ss = ensureSpreadsheet_(ss);
  try {
    var name = (typeof readParam_==='function' ? (readParam_(ss,'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL') : 'MEMBRES_GLOBAL');
    var sh = ss.getSheetByName(name);
    if (!sh) return ss.getId() + '|MG:none';
    var lr = sh.getLastRow() || 0, lc = sh.getLastColumn() || 0;
    var hdr = (lr >= 1 && lc >= 1) ? sh.getRange(1,1,1, Math.min(lc,30)).getDisplayValues()[0].join('|') : '';
    var sig = [ss.getId(), name, lr, lc, hdr].join('|');
    var md5 = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, sig);
    var b64 = Utilities.base64EncodeWebSafe(md5).slice(0,22);
    return 'MG:' + b64;
  } catch(e) {
    return ss.getId() + '|MG:fallback';
  }
}

function getMembresIndex_(ss){
  ss = ensureSpreadsheet_(ss);
  var key = _mg_cacheKey_(ss);
  var now = Date.now();

  if (__MG_CACHE.data && __MG_CACHE.key === key && (now - __MG_CACHE.at) < 5*60*1000) {
    return __MG_CACHE.data;
  }

  try {
    var sc = CacheService.getScriptCache();
    var cached = sc.get(key);
    if (cached) {
      try {
        var parsed = JSON.parse(cached);
        __MG_CACHE = { at: now, key: key, data: parsed };
        return parsed;
      } catch(_){}
    }
  } catch(_){}

  var name = (typeof readParam_==='function' ? (readParam_(ss,'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL') : 'MEMBRES_GLOBAL');
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 1) {
    var empty = { byPass:{}, headers:[], rows:[], key:key };
    __MG_CACHE = { at: now, key: key, data: empty };
    return empty;
  }

  var V = sh.getDataRange().getValues();
  var H = V[0].map(String);
  var col = {};
  H.forEach(function(h,i){ col[h]=i; });
  function colIdx(n){ var i = col[n]; return (typeof i==='number') ? i : -1; }

  var cPass = colIdx('Passeport'),
      cPhEx = colIdx('PhotoExpireLe'),
      cPhInv= colIdx('PhotoInvalide'),
      cCas  = colIdx('CasierExpiré') >= 0 ? colIdx('CasierExpiré') : colIdx('CasierExpire'),
      cLast = colIdx('LastUpdate'),
      cDue  = colIdx('PhotoInvalideDuesLe'),
      cStat = colIdx('StatutMembre');

  var byPass = Object.create(null);

  for (var r=1; r<V.length; r++){
    var row = V[r];
    var p8 = normalizePassportPlain8_(row[cPass]);
    if (!p8) continue;

    var photoInv = String(row[cPhInv] == null ? '' : row[cPhInv]).toLowerCase();
    var casVal   = row[cCas];

    byPass[p8] = {
      Passeport: p8,
      PhotoExpireLe: (cPhEx >= 0 ? String(row[cPhEx]||'') : ''),
      PhotoInvalide: (photoInv === 'true' || photoInv === '1') ? 1 : 0,
      CasierExpire: (String(casVal||'').toLowerCase()==='1' || String(casVal||'').toLowerCase()==='true') ? 1 : 0,
      PhotoInvalideDuesLe: (cDue >= 0 ? String(row[cDue]||'') : ''),
      LastUpdate: (cLast >= 0 ? String(row[cLast]||'') : ''),
      StatutMembre: (cStat >= 0 ? String(row[cStat]||'') : '')
    };
  }

  var result = { byPass: byPass, headers: H, key: key };

  __MG_CACHE = { at: now, key: key, data: result };
  try { CacheService.getScriptCache().put(key, JSON.stringify(result), 300); } catch(_){}

  return result;
}

/** Accès feuille sécurisé (throw si manquante) */
function safeGetSheet_(ssOrId, name) {
  var ss = ensureSpreadsheet_(ssOrId);
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Feuille introuvable "' + name + '" dans ss=' + ss.getId());
  return sh;
}

/* ====== Allow-Orphan — Sources: MAPPINGS, PARAMS, RULES_JSON ====== */

/** Optionnel : si ta feuille MAPPINGS a une colonne "AllowOrphan" (TRUE/FALSE),
 *  tu peux l’exposer via catalog.match(...) → item.AllowOrphan.
 *  Ci-dessous, on tolère son absence sans erreur. */

/** Parsing simple d’un JSON paramétré côté PARAMS, clé RULES_JSON (optionnel) */
function _readRulesJson_(ss){
  try {
    var raw = (typeof readParam_==='function') ? (readParam_(ss, 'RULES_JSON') || '') : '';
    if (!raw) return null;
    return JSON.parse(String(raw));
  } catch(e){ return null; }
}

function _csvToSetU_(csv){
  var S = Object.create(null);
  String(csv||'').split(',').map(function(t){ return String(t||'').trim(); })
    .filter(Boolean).forEach(function(k){ S[k.toUpperCase()] = true; });
  return S;
}

function _safeRegex_(s){
  if (!s) return null;
  try { return new RegExp(String(s), 'i'); } catch(_){ return null; }
}

/** Construit un "matcher" d’allow-orphan à partir de PARAMS et RULES_JSON */
function _buildAllowOrphanConfig_(ss){
  // PARAMS (tous optionnels)
  var codesCSV  = (typeof readParam_==='function') ? readParam_(ss, 'ORPHAN_ALLOW_CODES')  : '';
  var groupsCSV = (typeof readParam_==='function') ? readParam_(ss, 'ORPHAN_ALLOW_GROUPS') : '';
  var regexStr  = (typeof readParam_==='function') ? readParam_(ss, 'ORPHAN_ALLOW_REGEX')  : '';

  var allowCodes  = _csvToSetU_(codesCSV);
  var allowGroups = _csvToSetU_(groupsCSV);
  var allowRE     = _safeRegex_(regexStr);

  // RULES_JSON supporte des règles structurées
  // Exemple:
  // { "articles": [
  //   { "action":"allow_orphan", "code":"FRAIS_TEMP" },
  //   { "action":"allow_orphan", "group":"PHOTOS" },
  //   { "action":"allow_orphan", "contains":"Don" }
  // ] }
  var rj = _readRulesJson_(ss);
  var rjCodes = Object.create(null), rjGroups = Object.create(null), rjContains = [];
  if (rj && Array.isArray(rj.articles)) {
    rj.articles.forEach(function(rule){
      if (!rule || String(rule.action||'').toLowerCase() !== 'allow_orphan') return;
      if (rule.code)  rjCodes[String(rule.code).toUpperCase()] = true;
      if (rule.group) rjGroups[String(rule.group).toUpperCase()] = true;
      if (rule.contains) rjContains.push(String(rule.contains));
    });
  }

  return {
    allowCodes:  allowCodes,
    allowGroups: allowGroups,
    allowRE:     allowRE,
    rjCodes:     rjCodes,
    rjGroups:    rjGroups,
    rjContains:  rjContains
  };
}

/** Test principal: faut-il **tolérer** un article orphelin ? */
function isAllowedOrphan_(ss, a, item, raw, allowCfg){
  // 0) Filet de sécurité: si l’article est annulé/exclu, il n’atteint normalement pas ce code
  //    (isActive_ le filtre déjà). On ne refait rien ici.

  // 1) Par mapping (MAPPINGS): item.AllowOrphan === true
  if (item && item.AllowOrphan === true) return true;

  // 2) Par PARAMS / RULES_JSON (codes / groupes)
  var codeU = (item && item.Code) ? String(item.Code).toUpperCase() : '';
  if (codeU && (allowCfg.allowCodes[codeU] || allowCfg.rjCodes[codeU])) return true;

  var grpU = (item && item.ExclusiveGroup) ? String(item.ExclusiveGroup).toUpperCase() : '';
  if (grpU && (allowCfg.allowGroups[grpU] || allowCfg.rjGroups[grpU])) return true;

  // 3) Par regex (PARAMS) ou "contains" (RULES_JSON) sur le libellé brut
  var lib = String(raw||'');
  if (allowCfg.allowRE && allowCfg.allowRE.test(lib)) return true;
  if (allowCfg.rjContains && allowCfg.rjContains.length){
    var L = lib.toLowerCase();
    for (var i=0;i<allowCfg.rjContains.length;i++){
      if (L.indexOf(String(allowCfg.rjContains[i]||'').toLowerCase()) !== -1) return true;
    }
  }

  return false;
}
/** Écrit un tableau d’objets en une passe (clear + setValues) */
function writeObjectsToSheet_(ssOrId, sheetName, rows, header) {
  var ss = ensureSpreadsheet_(ssOrId);
  var sh = getSheetOrCreate_(ss, sheetName);

  rows = rows || [];
  // Déterminer l'entête
  var hdr = (header && header.length)
    ? header.slice()
    : (rows.length ? Object.keys(rows[0]) : []);

  // Construire la matrice [ [hdr...], [row...], ... ]
  var out = [hdr];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i], arr = new Array(hdr.length);
    for (var c = 0; c < hdr.length; c++) {
      var k = hdr[c];
      arr[c] = (r && r[k] != null) ? r[k] : '';
    }
    out.push(arr);
  }

  overwriteSheet_(sh, out); // clear + setValues en 1 seul call
}

/** Append d’objets (respecte l’entête déjà en place, crée si absent) */
function appendObjectsToSheet_(ssOrId, sheetName, rows) {
  var ss = ensureSpreadsheet_(ssOrId);
  var sh = getSheetOrCreate_(ss, sheetName);
  rows = rows || [];
  if (!rows.length) return 0;

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  var hdr;

  if (lastRow >= 1 && lastCol >= 1) {
    hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  } else {
    // Initialiser l’entête d’après le 1er objet
    hdr = Object.keys(rows[0] || {});
    if (!hdr.length) return 0;
    sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
    lastRow = 1;
    lastCol = hdr.length;
  }

  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i], arr = new Array(hdr.length);
    for (var c = 0; c < hdr.length; c++) {
      var k = hdr[c];
      arr[c] = (r && r[k] != null) ? r[k] : '';
    }
    out.push(arr);
  }

  sh.getRange(lastRow + 1, 1, out.length, hdr.length).setValues(out);
  return out.length;
}


function slugify_(s){ return String(s||'').toLowerCase()
  .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
  .replace(/[^a-z0-9]+/g,'-').replace(/^-+|-+$/g,''); }

function deriveTags_(name){
  const s=String(name||'').toLowerCase();
  const tags=new Set();
  if (/camp/.test(s)) tags.add('camp');
  if (/\bcdp\b|centre de d[eé]veloppement/.test(s)) tags.add('cdp');
  if (/entra[îi]neur|coach/.test(s)) tags.add('coach');
  if (/futsal/.test(s)) tags.add('futsal');
  if (/gardien/.test(s)) tags.add('gardien');
  if (/adulte|s[eé]nior/.test(s)) tags.add('adulte');
  // U-bands
  const m = s.match(/\bu-?\s?(\d{1,2})/i);
  if (m){
    const u = +m[1];
    if (u<=8) tags.add('u4u8');
    else if (u<=12) tags.add('u9u12');
    else if (u<=18) tags.add('u13u18');
  }
  if (/saison/i.test(s)) tags.add('inscription_normale');
  if (/camp de s[eé]lection|s[eé]lection/.test(s)) tags.add('camp_selection');
  return Array.from(tags);
}

function deriveCatCodeFromTags_(tags){
  if (tags.includes('cdp')) return 'CDP';
  if (tags.includes('camp')) return 'CAMP';
  if (tags.includes('futsal')) return 'FUTSAL';
  if (tags.includes('inscription_normale')) return 'SEASON';
  if (tags.includes('coach')) return 'COACH';
  return 'OTHER';
}

function deriveAudience_(tags){
  if (tags.includes('coach')) return 'Entraîneur';
  if (tags.includes('adulte')) return 'Adulte';
  return 'Joueur';
}

function deriveIsCoachFee_(tags){ return tags.includes('coach') ? 1 : 0; }

function deriveProgramBand_(tags){
  if (tags.includes('u4u8')) return 'U4-U8';
  if (tags.includes('u9u12')) return 'U9-U12';
  if (tags.includes('u13u18')) return 'U13-U18';
  if (tags.includes('adulte')) return 'Adulte';
  return '';
}

function derivePaymentStatus_(due, paid, rest){
  const d=Number(due)||0, p=Number(paid)||0;
  const r = (rest===''||rest==null) ? (d-p) : Number(rest)||0;
  if (r<=0 && (d>0 || p>0)) return 'Paid';
  if (p>0 && r>0) return 'Partial';
  return 'Unpaid';
}

function makePS_(p, saison){ return (p ? String(p).padStart(8,'0') : '') + '|' + String(saison||''); }

function pickPrimaryEmail_(emailsStr){
  const arr = String(emailsStr||'').split(/[;,]/).map(s=>s.trim()).filter(Boolean);
  const bad = /noreply|no-reply|invalid|test|example/i;
  const firstGood = arr.find(e=>!bad.test(e));
  return firstGood || (arr[0]||'');
}

function pickAgeBracketFromLedgerRows_(rows){ // rows = lignes LEDGER pour ce passeport
  // priorité : U4-U8 > U9-U12 > U13-U18 > Adulte (choisis l'info la plus précise)
  const bands = new Set(rows.map(r=>r.ProgramBand).filter(Boolean));
  if (bands.has('U4-U8'))  return 'U4-U8';
  if (bands.has('U9-U12')) return 'U9-U12';
  if (bands.has('U13-U18'))return 'U13-U18';
  if (bands.has('Adulte')) return 'Adulte';
  return '';
}

