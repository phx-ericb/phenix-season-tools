function setSeasonIdGlobalOnce() {
  var id = '1IVVHi17Jyo8jvWtrSuenbPW8IyEZqlY1bXx-WbnXPkk'; // ton ID
  var p = PropertiesService.getScriptProperties();
  p.setProperty('ACTIVE_SEASON_ID', id);
  p.setProperty('PHENIX_SEASON_SHEET_ID', id); // alias lu par la lib
  p.setProperty('SEASON_SPREADSHEET_ID', id);  // alias supplémentaire
}


/** Ouvre le classeur de la saison ou lève une erreur claire */
function getSeasonSpreadsheet_(seasonSheetId) {
  if (seasonSheetId) return SpreadsheetApp.openById(seasonSheetId);

  // ↙️ mêmes clés que celles que tu utilises dans Code.js
  var props = PropertiesService.getScriptProperties();
  var id =
    props.getProperty('ACTIVE_SEASON_ID') ||
    props.getProperty('PHENIX_SEASON_SHEET_ID') ||
    props.getProperty('SEASON_SPREADSHEET_ID');

  if (id) return SpreadsheetApp.openById(id);

  // Dernier filet: si exécuté depuis un classeur (rare côté lib), prends l’actif
  try {
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}

  throw new Error("seasonSheetId manquant. Passe l’ID ou définis ACTIVE_SEASON_ID/PHENIX_SEASON_SHEET_ID.");
}


/** Récupère une feuille, ou la crée vide avec un header si fourni */
function getSheetOrCreate_(ss, name, header) {
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
  getSheetOrCreate_(ss, SHEETS.PARAMS);
  getSheetOrCreate_(ss, SHEETS.IMPORT_LOG, ['Horodatage','Action','Détails']);
  ensureMailOutbox_(ss);
  getSheetOrCreate_(ss, SHEETS.MAIL_LOG, ['Type','To','Sujet','KeyHash','SentAt','Result']);
  // ERREURS enrichi v0.7
  getSheetOrCreate_(ss, SHEETS.ERREURS, ['Passeport','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']);
  getSheetOrCreate_(ss, SHEETS.EXPORT_LOG, ['Horodatage','ExportType','Fichier','Checksum','Lien']);
  // Annulations (avec colonnes "actifs restants")
  getSheetOrCreate_(ss, SHEETS.ANNULATIONS_INSCRIPTIONS, ['Horodatage','Passeport','Nom','Prénom','NomComplet','Saison','Frais','DateAnnulation','A_ENCORE_ACTIF','ACTIFS_RESTANTS']);
  getSheetOrCreate_(ss, SHEETS.ANNULATIONS_ARTICLES,     ['Horodatage','Passeport','Nom','Prénom','NomComplet','Saison','Frais','DateAnnulation','A_ENCORE_ACTIF','ACTIFS_RESTANTS']);
  // Modifs
  getSheetOrCreate_(ss, SHEETS.MODIFS_INSCRIPTIONS, ['Horodatage','Passeport','Nom','Prénom','NomComplet','Saison','ChangedFieldsJSON']);
}

/** Lit un paramètre depuis PARAMS (fallback DocumentProperties) */
function readParam_(ss, key) {
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
  var sh = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
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

/** Normalise Passeport #:
 * - null/'' → ''
 * - si numérique → pad à 8 (garde >8 tel quel)
 * - retour toujours en TEXTE via apostrophe en tête
 */
function normalizePassportToText8_(val) {
  if (val == null) return '';
  var s = String(val).trim();
  if (s === '') return '';
  if (s[0] === "'") s = s.slice(1); // strip apostrophe éventuelle
  if (/^\d+$/.test(s)) {
    if (s.length < 8) s = ('00000000' + s).slice(-8);
  }
  return "'" + s; // force texte
}
/** Version "export" : réutilise normalizePassportToText8_ mais retire l’apostrophe.
 *  -> On conserve le padding à 8, on écrit une chaîne "00123456" (sans '),
 *  -> et on met la colonne A en format texte côté export.
 */
function normalizePassportPlain8_(val) {
  var s = normalizePassportToText8_(val); // ex: "'00123456"
  return (s && s[0] === "'") ? s.slice(1) : s; // -> "00123456"
}


// ---- Params & statut helpers ----
function _paramOr_(ss, key, dflt) {
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

// ---- MAIL_OUTBOX helpers (globalisés) ----
function getMailOutboxHeaders_() {
  return ['Type','To','Cc','Sujet','Corps','Attachments','KeyHash','Status','CreatedAt','SentAt','Error'];
}
/** MAIL_OUTBOX: garantit l’entête attendue (et répare si besoin). */
function ensureMailOutbox_(ss) {
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

/**
 * Ajoute N lignes dans MAIL_OUTBOX (idempotence via KeyHash+Type gérée par l’appelant).
 * rows = array of [Type,To,Cc,Sujet,Corps,Attachments,KeyHash,'pending',now,'','']
 */
/** Ecrit les lignes en respectant la largeur d’entête (pad/troncature) */
function enqueueOutboxRows_(ssId, rows) {
  if (!rows || !rows.length) return 0;
  var ss = SpreadsheetApp.openById(ssId);
  var sh = ensureMailOutbox_(ss);
  var headers = getMailOutboxHeaders_();
  var W = headers.length;

  // normalise la largeur de chaque ligne
  var toWrite = rows.map(function(r){
    var arr = (r && r.slice) ? r.slice(0, W) : [];
    while (arr.length < W) arr.push('');
    return arr;
  });

  var start = sh.getLastRow() + 1;
  if (toWrite.length) {
    sh.insertRowsAfter(sh.getLastRow(), toWrite.length);
    sh.getRange(start, 1, toWrite.length, W).setValues(toWrite);
  }
  return toWrite.length;
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

/** ====== Helpers v0.7 (nouveaux) ====== */
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

/* ===== Lecture MAPPINGS unifiés (incl. ExclusiveGroup) ===== */
function _loadUnifiedGroupMappings_(ss) {
  // 1) feuille
  var sh = ss.getSheetByName(SHEETS.MAPPINGS) || ss.getSheetByName('MAPPINGS');
  if (!sh) return [];

  // 2) données + détection d’entête « première ligne valable » (≥2 cellules non vides)
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

  // 3) index des colonnes
  var H = (data[headerIdx] || []).map(function (h) { return String(h || '').trim(); });
  function idx(k) { var i = H.indexOf(k); return i < 0 ? null : i; }

  var iType = idx('Type'), iAli = idx('AliasContains'), iUmin = idx('Umin'), iUmax = idx('Umax'),
      iGen  = idx('Genre'), iG   = idx('GroupeFmt'),     iC    = idx('CategorieFmt'),
      iEx   = idx('Exclude'), iPr = idx('Priority'), iX = idx('ExclusiveGroup'), iCode = idx('Code');

  if (iType == null || iAli == null) return [];

  // 4) lecture des lignes
// 4) lecture des lignes (REMPLACE TOUT CE BLOC)
var out = [];
function _t(v) { return String(v == null ? '' : v).trim(); }            // trim simple
function _tu(v) { return _t(v).toUpperCase(); }                         // trim + upper
function _tl(v) { return _t(v).toLowerCase(); }                         // trim + lower
function _toIntOrNull(v) {
  var s = _t(v); if (s === '') return null;
  var n = parseInt(s, 10); return isNaN(n) ? null : n;
}

for (var r = headerIdx + 1; r < data.length; r++) {
  var row = data[r] || [];
  // ligne vide ?
  var isEmpty = true;
  for (var c = 0; c < row.length; c++) {
    if (_t(row[c]) !== '') { isEmpty = false; break; }
  }
  if (isEmpty) continue;

  var type = (iType == null) ? '' : _tl(row[iType]);                    // <-- TRIM + lower
  var ali  = (iAli  == null) ? '' : _t(row[iAli]);                      // <-- TRIM
  var umin = (iUmin == null) ? null : _toIntOrNull(row[iUmin]);
  var umax = (iUmax == null) ? null : _toIntOrNull(row[iUmax]);
  var gen  = (iGen  == null) ? '*' : _tu(row[iGen] || '*');             // <-- TRIM + upper
  var grp  = (iG    == null) ? '' : _t(row[iG]);                        // <-- TRIM
  var cat  = (iC    == null) ? '' : _t(row[iC]);                        // <-- TRIM
  var ex   = (iEx   == null) ? '' : _tl(row[iEx]);                      // <-- TRIM + lower
  var pri  = (iPr   == null) ? null : _toIntOrNull(row[iPr]);
  var exg  = (iX    == null) ? '' : _t(row[iX]);                        // <-- TRIM
  var code = (iCode == null) ? '' : _t(row[iCode]);                     // <-- TRIM

  out.push({
    Type: type,                                // "article" / "member"
    AliasContains: ali,                        // propre (trim)
    Umin: umin,                                // nombre ou null
    Umax: umax,
    Genre: gen || '*',                         // "M" | "F" | "X" | "*"
    GroupeFmt: grp,
    CategorieFmt: cat,
    Exclude: ex === 'true',
    Priority: (pri == null ? 100 : pri),
    ExclusiveGroup: exg,
    Code: code
  });
}

// 5) tri (inchangé)
out.sort(function (a, b) {
  if (a.Priority !== b.Priority) return b.Priority - a.Priority;
  return (b.AliasContains || '').length - (a.AliasContains || '').length;
});
return out;

}
