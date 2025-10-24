/***** server_membres.js *****
 * Chaîne complète: GLOBAL → SAISON (subset) → JOUEURS (maj photo)
 * - API_VM_importToGlobal(force)
 * - API_VM_refreshSeasonSubset()
 * - API_VM_fullRefresh(force)
 *
 * Hypothèses:
 *  - Fonctions utilitaires potentiellement déjà présentes: readParam_, readParamValue,
 *    getSeasonId_, _vm_findLatestActiveFile_, _vm_readFileAsSheetValues_,
 *    _vm_indexCols_, _vm_toISODate_, _vm_to01_, _vm_hashRow_, _vm_groupRuns_,
 *    _vm_nowISO_, _vm_rowValues_, log_, normalizePassportToText8_, getSheetOrCreate_.
 *  - Ce fichier fournit des “shims” qui s’activent seulement si tes versions n’existent pas.
 */

/* ────────────────────────────────────────────────────────────────────────── */
/* Shims utilitaires (activés seulement si absents du projet courant)        */
/* ────────────────────────────────────────────────────────────────────────── */

if (typeof getSheetOrCreate_ !== 'function') {
  function getSheetOrCreate_(ss, name, headersOpt) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (headersOpt && headersOpt.length && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headersOpt.length).setValues([headersOpt]);
    }
    return sh;
  }
}

// --- Lecture générique d’un fichier (Sheet, CSV/TSV, XLSX) en 2D array [rows][cols]
if (typeof _vm_readFileAsSheetValues_ !== 'function') {

  function _vm_readFileAsSheetValues_(file, wantedHeadersOpt) {
    // wantedHeadersOpt: tableau de libellés probables pour aider à choisir la bonne feuille (facultatif)
    // Retour: Array<Array<any>> incluant l’en-tête en [0]
    if (!file) throw new Error('_vm_readFileAsSheetValues_: file manquant');

    var mime = (file.getMimeType && file.getMimeType()) || '';
    var name = (file.getName && file.getName()) || '';
    var blob = (file.getBlob && file.getBlob()) ? file.getBlob() : null;

    // 1) Google Sheet natif
    if (mime === 'application/vnd.google-apps.spreadsheet') {
      var ss = SpreadsheetApp.openById(file.getId());
      return _vm__readBestSheetValues_(ss, wantedHeadersOpt);
    }

    // 2) CSV / texte délimité
    if (/^text\/|csv|tab-separated|tsv|plain|application\/vnd\.ms-excel$/.test(mime) || /\.(csv|tsv)$/i.test(name)) {
      // lire comme texte puis parser
      var text = blob ? blob.getDataAsString() : '';
      if (!text) return [];
      var delim = _vm__detectDelimiter_(text); // ; , \t
      var rows = Utilities.parseCsv(text, delim);
      return rows || [];
    }

// 3) Excel (xlsx/xls) → convertir vers Google Sheet puis lire
if (
  /\.xlsx$/i.test(name) || /\.xls$/i.test(name) ||
  /application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet/.test(mime) ||
  /application\/vnd\.ms-excel/.test(mime)
) {
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error("Lecture XLSX requiert le service avancé Drive activé + API Drive activée (v2 ou v3).");
  }

  // Ressource cible = Google Sheet
  var resource = {
    title: name.replace(/\.(xlsx|xls)$/i, '') + ' (temp import)',
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };

  // v3: create(...) ; v2: insert(...)
  var useCreate = (typeof Drive.Files.create === 'function');
  var useInsert = (typeof Drive.Files.insert === 'function');
  if (!useCreate && !useInsert) {
    throw new Error("Drive.Files.create/insert introuvable. Vérifie que Drive avancé (v3 ou v2) est bien ON.");
  }

  // upload + conversion
  var tempFile = useCreate
    ? Drive.Files.create(resource, blob)          // v3
    : Drive.Files.insert(resource, blob);         // v2

  try {
    var ssConv = SpreadsheetApp.openById(tempFile.id);
    var vals   = _vm__readBestSheetValues_(ssConv, wantedHeadersOpt);

    // Nettoyage (v3: delete, v2: remove)
    try {
      if (typeof Drive.Files.delete === 'function') Drive.Files.delete(tempFile.id); // v3
      else if (typeof Drive.Files.remove === 'function') Drive.Files.remove(tempFile.id); // v2
    } catch (_) {}
    return vals;
  } catch (e) {
    try {
      if (typeof Drive.Files.delete === 'function') Drive.Files.delete(tempFile.id);
      else if (typeof Drive.Files.remove === 'function') Drive.Files.remove(tempFile.id);
    } catch (_) {}
    throw e;
  }
}

    // 4) Dernier recours: essayer CSV
    if (blob) {
      var txt = blob.getDataAsString();
      if (txt) {
        var d = _vm__detectDelimiter_(txt);
        var r = Utilities.parseCsv(txt, d);
        if (r && r.length) return r;
      }
    }

    throw new Error("Type de fichier non supporté pour VM: " + name + " (" + mime + ")");
  }

  // --- Choisit la meilleure feuille d’un classeur puis renvoie toutes les valeurs
  function _vm__readBestSheetValues_(ss, wantedHeadersOpt) {
    var sheets = ss.getSheets();
    if (!sheets || !sheets.length) return [];

    // Heuristique: choisir la feuille qui a (1) des en-têtes “prometteurs” si dispos, sinon (2) le plus de cellules non vides
    var best = null, bestScore = -1, bestVals = null;

    for (var i = 0; i < sheets.length; i++) {
      var sh = sheets[i];
      var lr = sh.getLastRow(), lc = sh.getLastColumn();
      if (lr < 1 || lc < 1) continue;

      var vals = sh.getRange(1, 1, lr, lc).getValues();
      if (!vals || !vals.length) continue;

      var header = (vals[0] || []).map(function(h){ return String(h||'').trim().toLowerCase(); });

      var score = 0;
      if (wantedHeadersOpt && wantedHeadersOpt.length) {
        // bonus par correspondance d'en-têtes exactes/partielles
        for (var k=0;k<wantedHeadersOpt.length;k++) {
          var needle = String(wantedHeadersOpt[k]).trim().toLowerCase();
          if (header.indexOf(needle) !== -1) score += 5;
          else {
            // match partiel
            for (var h=0;h<header.length;h++) {
              if (header[h].indexOf(needle) !== -1) { score += 2; break; }
            }
          }
        }
      }
      // densité: nb cellules non vides (cap à 1000 pour ne pas gonfler démesurément)
      var nonEmpty = 0;
      for (var r=0;r<Math.min(vals.length, 1000); r++) {
        for (var c=0;c<Math.min(vals[r].length, 50); c++) {
          if (vals[r][c] !== '' && vals[r][c] !== null) nonEmpty++;
        }
      }
      score += Math.min(nonEmpty, 1000) / 50; // +20 max

      if (score > bestScore) { bestScore = score; best = sh; bestVals = vals; }
    }

    return bestVals || [];
  }

  // --- Détecte ; , ou \t comme délimiteur probable
  function _vm__detectDelimiter_(text) {
    // échantillon de quelques lignes
    var lines = text.split(/\r?\n/).slice(0, 5);
    var counts = { ',':0, ';':0, '\t':0 };
    lines.forEach(function(line){
      counts[',']  += (line.match(/,/g)  || []).length;
      counts[';']  += (line.match(/;/g)  || []).length;
      counts['\t'] += (line.match(/\t/g) || []).length;
    });
    // priorité: la plus fréquente; tie-breaker: ; puis , puis \t (souvent VM FR = ;)
    var best = ','; var max = counts[','];
    if (counts[';'] > max) { best = ';'; max = counts[';']; }
    if (counts['\t'] > max) { best = '\t'; max = counts['\t']; }
    return best;
  }
}


if (typeof getSheet_ !== 'function') {
  function getSheet_(ss, name, createIfMissing) {
    return createIfMissing ? getSheetOrCreate_(ss, name) : ss.getSheetByName(name);
  }
}

if (typeof normalizePassportPlain8_ !== 'function') {
  function normalizePassportPlain8_(v) {
    if (typeof normalizePassportToText8_ === 'function') {
      return normalizePassportToText8_(v);
    }
    var s = String(v || '').replace(/\D/g, '');
    return s ? s.padStart(8, '0') : '';
  }
}

// --- Fallback si la fonction n'est pas visible dans l'environnement actuel
if (typeof _vm_findLatestActiveFile_ !== 'function') {
  function _vm_findLatestActiveFile_(folderId) {
    if (!folderId) return { file: null, source: 'fallback:no-folder' };
    var folder = DriveApp.getFolderById(folderId);
    var it = folder.getFiles();
    var latest = null, latestTs = -1;

    while (it.hasNext()) {
      var f = it.next();
      // on prend tout type de fichier; la lecture/parse sera gérée plus loin (_vm_readFileAsSheetValues_)
      var ts = (f.getLastUpdated && +f.getLastUpdated()) || 0;
      if (ts > latestTs) { latest = f; latestTs = ts; }
    }
    return { file: latest, source: 'fallback' };
  }
}


/* ────────────────────────────────────────────────────────────────────────── */
/* Helpers locaux                                                             */
/* ────────────────────────────────────────────────────────────────────────── */

function _readParamFromSeasonOrGlobal_(ssSeason, key, fallback) {
  try {
    if (typeof readParam_ === 'function') {
      var v1 = readParam_(ssSeason, key);
      if (v1) return v1;
    }
  } catch (_) {}
  try {
    if (typeof readParamValue === 'function') {
      var v2 = readParamValue(key);
      if (v2) return v2;
    }
  } catch (_) {}
  return fallback || '';
}

function _indexByHeader_(headers, wantedList) {
  var idx = -1;
  var map = {};
  headers.forEach(function(h, i) { map[String(h).trim().toLowerCase()] = i; });
  for (var k = 0; k < wantedList.length; k++) {
    var w = String(wantedList[k]).trim().toLowerCase();
    if (w in map) return map[w];
  }
  for (var j = 0; j < wantedList.length; j++) {
    var needle = String(wantedList[j]).trim().toLowerCase();
    for (var key in map) if (key.indexOf(needle) !== -1) return map[key];
  }
  return idx;
}

function _readSheet_(sh) {
  var lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return { headers: [], rows: [] };
  var values = sh.getRange(1, 1, lr, lc).getValues();
  var headers = values.shift() || [];
  return { headers: headers, rows: values };
}

function _writeWholeSheet_(sh, headers, rows) {
  sh.clearContents();
  if (headers && headers.length) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows && rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

/* ────────────────────────────────────────────────────────────────────────── */
/* IMPORT GLOBAL: lit dernier fichier VM et upsert CENTRAL.MEMBRES_GLOBAL     */
/* ────────────────────────────────────────────────────────────────────────── */

function importValidationMembresToGlobal_(targetSpreadsheetId) {
  var seasonId = getSeasonId_();
  var ssSeason = SpreadsheetApp.openById(seasonId);

  var folderId = _readParamFromSeasonOrGlobal_(ssSeason, 'DRIVE_FOLDER_VALIDATION_MEMBRES', '');
  if (!folderId) throw new Error('DRIVE_FOLDER_VALIDATION_MEMBRES manquant (saison).');

  var seasonYear = Number(_readParamFromSeasonOrGlobal_(ssSeason, 'SEASON_YEAR', new Date().getFullYear()));
  var photoInvalidFrom = String(_readParamFromSeasonOrGlobal_(ssSeason, 'PHOTO_INVALID_FROM_MMDD', '04-01')).trim();

  if (!targetSpreadsheetId) throw new Error('ID du classeur CENTRAL manquant.');
  var ssCentral = SpreadsheetApp.openById(targetSpreadsheetId);

  var found = _vm_findLatestActiveFile_(folderId);
  var file = found && found.file;
  if (!file) {
    if (typeof log_ === 'function') log_('VM_IMPORT_NOFILE', 'Aucun fichier actif dans Validation_Membres.');
    return { created: 0, updated: 0, unchanged: 0 };
  }
  if (typeof log_ === 'function') log_('VM_IMPORT_START', file.getName() + ' (' + file.getId() + ')');

  var values = _vm_readFileAsSheetValues_(file);
  if (!values || values.length < 2) throw new Error('Fichier Validation_Membres vide/illisible.');

  var headers = values[0].map(function(h){ return String(h).trim(); });

  // Index principaux (via ton util; sinon fallback strict)
  var idx = (typeof _vm_indexCols_ === 'function')
    ? _vm_indexCols_(headers, [
        'Passeport #','Prénom','Nom','Date de naissance',
        'Identité de genre','Statut du membre',
        "Date d'expiration de la photo de profil",
        'Vérification du casier judiciaire est expiré',
        'Courriel','Email','Parent 1 - Courriel','Parent 2 - Courriel'
      ])
    : {};

  function findIdx_(label) {
    var want = String(label).trim().toLowerCase();
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i]).trim().toLowerCase() === want) return i;
    }
    return -1;
  }

  function idxOf(label) {
    return (idx && label in idx && idx[label] >= 0) ? idx[label] : findIdx_(label);
  }

  var iPass   = idxOf('Passeport #');
  var iPrenom = idxOf('Prénom');
  var iNom    = idxOf('Nom');
  var iDOB    = idxOf('Date de naissance');
  var iGenre  = idxOf('Identité de genre');
  var iStatut = idxOf('Statut du membre');
  var iPhoto  = idxOf("Date d'expiration de la photo de profil");
  var iCasier = idxOf('Vérification du casier judiciaire est expiré');

  var iCourriel     = idxOf('Courriel');
  var iEmail        = (iCourriel < 0) ? idxOf('Email') : -1;
  var iParent1Mail  = idxOf('Parent 1 - Courriel');
  var iParent2Mail  = idxOf('Parent 2 - Courriel');

  var seasonInvalidDate = seasonYear + '-' + photoInvalidFrom; // p.ex. 2026-04-01
  var cutoffNextJan1    = (seasonYear + 1) + '-01-01';

  function pickPrimaryEmail_(parts) {
    var bad = /noreply|no-reply|invalid|example|test/i;
    var list = [];
    parts.forEach(function(p){
      String(p || '').split(/[;,]/).forEach(function(s){
        var t = s.trim();
        if (t) list.push(t);
      });
    });
    for (var k = 0; k < list.length; k++) if (!bad.test(list[k])) return list[k];
    return list[0] || '';
  }

  var targetByPassport = new Map();

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row || row.length === 0) continue;

    var passport = normalizePassportPlain8_(row[iPass]);
    if (!passport) continue;

    var prenom   = String(row[iPrenom] || '').trim();
    var nom      = String(row[iNom] || '').trim();
    var dob      = (typeof _vm_toISODate_ === 'function') ? _vm_toISODate_(row[iDOB]) : row[iDOB];
    var genre    = String(row[iGenre] || '').trim();
    var statut   = String(row[iStatut] || '').trim();
    var photoExp = (typeof _vm_toISODate_ === 'function') ? _vm_toISODate_(row[iPhoto]) : row[iPhoto];
    var casier01 = (typeof _vm_to01_ === 'function') ? _vm_to01_(row[iCasier]) : (row[iCasier] ? 1 : 0);

    var photoInvalide = (!photoExp || String(photoExp) < cutoffNextJan1) ? 1 : 0;
    var photoDuesLe   = photoInvalide ? seasonInvalidDate : '';

    var courriel = pickPrimaryEmail_([
      (iCourriel >= 0 ? row[iCourriel] : ''),
      (iEmail    >= 0 ? row[iEmail]    : ''),
      (iParent1Mail >= 0 ? row[iParent1Mail] : ''),
      (iParent2Mail >= 0 ? row[iParent2Mail] : '')
    ]);

    var obj = {
      'Passeport': passport,
      'Prenom': prenom,
      'Nom': nom,
      'DateNaissance': dob,
      'Genre': genre,
      'StatutMembre': statut,
      'PhotoExpireLe': photoExp,
      'PhotoInvalideDuesLe': photoDuesLe,
      'PhotoInvalide': photoInvalide,
      'CasierExpiré': casier01,
      'SeasonYear': seasonYear,
      'Courriel': courriel
    };

    obj.RowHash = (typeof _vm_hashRow_ === 'function') ? _vm_hashRow_(obj) : JSON.stringify(obj);
    targetByPassport.set(passport, obj);
  }

  var sheetName = _readParamFromSeasonOrGlobal_(ssSeason, 'SHEET_MEMBRES_GLOBAL', 'MEMBRES_GLOBAL');
  var sh = getSheetOrCreate_(ssCentral, sheetName);

  var colOrder = [
    'Passeport','Prenom','Nom','DateNaissance','Genre','StatutMembre',
    'PhotoExpireLe','PhotoInvalideDuesLe','PhotoInvalide',
    'CasierExpiré','SeasonYear','RowHash','LastUpdate','Courriel'
  ];

  if (typeof _vm_ensureHeader_ === 'function') {
    _vm_ensureHeader_(sh, colOrder);
  } else {
    var hdr = sh.getLastRow() ? sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] : [];
    if (!hdr || hdr.join('|') !== colOrder.join('|')) {
      sh.clear();
      sh.getRange(1,1,1,colOrder.length).setValues([colOrder]);
    }
  }

  var data = sh.getDataRange().getValues();
  var header = data[0];
  var colIdx = {};
  header.forEach(function(h, i){ colIdx[String(h)] = i; });

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

  var nowISO = (typeof _vm_nowISO_ === 'function') ? _vm_nowISO_() : new Date().toISOString();
  function buildRow_(obj) {
    if (typeof _vm_rowValues_ === 'function') return _vm_rowValues_(colOrder, obj, nowISO);
    var out = [];
    colOrder.forEach(function(k){
      if (k === 'LastUpdate') out.push(nowISO);
      else out.push(obj.hasOwnProperty(k) ? obj[k] : '');
    });
    return out;
  }

  var toWrite = [];   // [rowIndex(1-based), rowValues]
  var toAppend = [];  // [rowValues]
  var updated = 0, created = 0, unchanged = 0;

  targetByPassport.forEach(function(obj, pass){
    var ex = existingIdx.get(pass);
    if (ex) {
      if (ex.hash === obj.RowHash) { unchanged++; return; }
      toWrite.push([ex.rowIndex + 1, buildRow_(obj)]);
      updated++;
    } else {
      toAppend.push(buildRow_(obj));
      created++;
    }
  });

  if (toWrite.length) {
    toWrite.sort(function(a,b){ return a[0] - b[0]; });
    var runs = (typeof _vm_groupRuns_ === 'function') ? _vm_groupRuns_(toWrite) : [toWrite];
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

  if (typeof log_ === 'function') log_('VM_IMPORT_SUMMARY', 'created=' + created + ', updated=' + updated + ', unchanged=' + unchanged);
  if (typeof _vm_notifyImportMembres_ === 'function') {
    try { _vm_notifyImportMembres_({created:created, updated:updated, unchanged:unchanged}); } catch(_) {}
  }

  return { created: created, updated: updated, unchanged: unchanged };
}

/* ────────────────────────────────────────────────────────────────────────── */
/* SUBSET SAISON: GLOBAL → MEMBRES_GLOBAL_SAISON (filtré par passeports)     */
/* ────────────────────────────────────────────────────────────────────────── */

// GLOBAL -> (filtré) -> SAISON.MEMBRES_GLOBAL
function syncMembresGlobalSubsetFromCentral_(seasonId, centralId) {
  var ssSeason = SpreadsheetApp.openById(seasonId);
  var ssGlobal = SpreadsheetApp.openById(centralId);

  var globalSheetName  = (typeof readParam_==='function' ? readParam_(ssSeason,'SHEET_MEMBRES_GLOBAL') : '') || 'MEMBRES_GLOBAL';
  var seasonSheetName  = globalSheetName; // 👈 même nom dans la SAISON
  var joueursSheetName = (typeof readParam_==='function' ? readParam_(ssSeason,'SHEET_JOUEURS') : '') || 'JOUEURS';

  var shGlobal  = ssGlobal.getSheetByName(globalSheetName);
  if (!shGlobal) throw new Error('GLOBAL: feuille "' + globalSheetName + '" introuvable.');
  var shJoueurs = ssSeason.getSheetByName(joueursSheetName);
  if (!shJoueurs) throw new Error('SAISON: feuille "' + joueursSheetName + '" introuvable.');

  // Lire global + joueurs
  var lrG = shGlobal.getLastRow(), lcG = shGlobal.getLastColumn();
  if (lrG < 1 || lcG < 1) throw new Error('GLOBAL vide.');
  var valsG = shGlobal.getRange(1,1,lrG,lcG).getValues();
  var header = valsG.shift();

  var lrJ = shJoueurs.getLastRow(), lcJ = shJoueurs.getLastColumn();
  if (lrJ < 1 || lcJ < 1) throw new Error('JOUEURS vide.');
  var valsJ = shJoueurs.getRange(1,1,lrJ,lcJ).getValues();
  var HJ = valsJ[0]; valsJ.shift();

  // index colonnes "Passeport"
  function idxOf(H, list) {
    var low = {}; H.forEach(function(h,i){ low[String(h).trim().toLowerCase()] = i; });
    for (var k=0;k<list.length;k++) {
      var n = String(list[k]).toLowerCase(); if (n in low) return low[n];
    }
    for (var k2=0;k2<list.length;k2++) {
      var needle = String(list[k2]).toLowerCase();
      for (var key in low) if (key.indexOf(needle) !== -1) return low[key];
    }
    return -1;
  }
  var iPassG = idxOf(header, ['passeport #','passeport','passport','no passeport']);
  var iPassJ = idxOf(HJ,     ['passeport #','passeport','passport','no passeport']);
  if (iPassG < 0) throw new Error('GLOBAL: colonne Passeport # introuvable.');
  if (iPassJ < 0) throw new Error('JOUEURS: colonne Passeport # introuvable.');

  var seasonPass = new Set(valsJ.map(function(r){ return String(r[iPassJ]||'').trim(); }).filter(Boolean));

  // Filtrer global
  var filtered = [];
  for (var i=0;i<valsG.length;i++) {
    var p = String(valsG[i][iPassG]||'').trim();
    if (p && seasonPass.has(p)) filtered.push(valsG[i]);
  }

  // Écrire dans SAISON.MEMBRES_GLOBAL (même nom), en remplaçant le contenu
  var shSeasonMG = ssSeason.getSheetByName(seasonSheetName) || ssSeason.insertSheet(seasonSheetName);
  shSeasonMG.clearContents();
  shSeasonMG.getRange(1,1,1,header.length).setValues([header]);
  if (filtered.length) {
    shSeasonMG.getRange(2,1,filtered.length,header.length).setValues(filtered);
  }

  return { sheet: seasonSheetName, kept: filtered.length, totalGlobal: valsG.length };
}


/* ────────────────────────────────────────────────────────────────────────── */
/* APPLY → JOUEURS: met à jour PhotoExpireLe (+ flag Photo expirée si présent)*/
/* ────────────────────────────────────────────────────────────────────────── */

function applySeasonVMToJoueurs_() {
  var seasonId = getSeasonId_();
  var ssSeason = SpreadsheetApp.openById(seasonId);

  var subsetSheetName  = _readParamFromSeasonOrGlobal_(ssSeason, 'SHEET_MEMBRES_GLOBAL_SAISON', 'MEMBRES_GLOBAL_SAISON');
  var joueursSheetName = _readParamFromSeasonOrGlobal_(ssSeason, 'SHEET_JOUEURS', 'JOUEURS');

  var shSubset  = ssSeason.getSheetByName(subsetSheetName);
  var shJoueurs = ssSeason.getSheetByName(joueursSheetName);
  if (!shSubset)  throw new Error('Subset "' + subsetSheetName + '" introuvable.');
  if (!shJoueurs) throw new Error('JOUEURS "' + joueursSheetName + '" introuvable.');

  var S = _readSheet_(shSubset);
  var J = _readSheet_(shJoueurs);
  if (!S.headers.length || !S.rows.length) return { updated: 0, reason: 'subset vide' };
  if (!J.headers.length) return { updated: 0, reason: 'joueurs vide' };

  var idxS_pass = _indexByHeader_(S.headers, ['Passeport #','Passeport','Passport','No passeport']);
  var idxS_exp  = _indexByHeader_(S.headers, ['PhotoExpireLe',"Date d'expiration de la photo","Photo expire le","Photo Expire Le"]);
  if (idxS_pass < 0) throw new Error('Subset: colonne Passeport # introuvable.');
  if (idxS_exp  < 0) throw new Error('Subset: colonne PhotoExpireLe introuvable.');

  var idxJ_pass      = _indexByHeader_(J.headers, ['Passeport #','Passeport','Passport','No passeport']);
  var idxJ_photoDate = _indexByHeader_(J.headers, ['PhotoExpireLe',"Date d'expiration de la photo","Photo expire le"]);
  var idxJ_photoFlag = _indexByHeader_(J.headers, ['Photo expirée','Photo est expirée','Statut photo','PhotoStatus']);

  if (idxJ_pass < 0) throw new Error('JOUEURS: colonne Passeport # introuvable.');
  if (idxJ_photoDate < 0 && idxJ_photoFlag < 0) return { updated: 0, reason: 'aucune colonne cible photo dans JOUEURS' };

  var mapExp = new Map();
  for (var i = 0; i < S.rows.length; i++) {
    var r = S.rows[i];
    var p = String(r[idxS_pass] || '').trim();
    var d = r[idxS_exp];
    mapExp.set(p, d);
  }

  var today = new Date();
  var td = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  var updates = 0;
  var out = J.rows.map(function(row){
    var p = String(row[idxJ_pass] || '').trim();
    if (!p) return row;
    if (!mapExp.has(p)) return row;

    var exp = mapExp.get(p);
    var changed = false;

    if (idxJ_photoDate >= 0) {
      var cur = row[idxJ_photoDate];
      var same = (cur && exp && String(cur) === String(exp)) || (!cur && !exp);
      if (!same) { row[idxJ_photoDate] = exp || ''; changed = true; }
    }

    if (idxJ_photoFlag >= 0) {
      var isExpired = false;
      if (exp instanceof Date) {
        var ed = new Date(exp.getFullYear(), exp.getMonth(), exp.getDate());
        isExpired = ed < td;
      } else if (exp) {
        var parsed = new Date(exp);
        if (!isNaN(parsed.getTime())) {
          var ed2 = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
          isExpired = ed2 < td;
        }
      }
      var curFlag = row[idxJ_photoFlag];
      var nextFlag = isExpired ? 1 : 0;
      if (String(curFlag) !== String(nextFlag)) { row[idxJ_photoFlag] = nextFlag; changed = true; }
    }

    if (changed) updates++;
    return row;
  });

  if (updates > 0) {
    shJoueurs.getRange(2, 1, out.length, J.headers.length).setValues(out);
  }

  return { updated: updates, rows: out.length };
}

/* ────────────────────────────────────────────────────────────────────────── */
/* APIs chaînées                                                             */
/* ────────────────────────────────────────────────────────────────────────── */

function API_VM_importToGlobal(force) {
  if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
    return { ok: false, busy: true, reason: 'import-running' };
  }
  var seasonId = getSeasonId_();
  var ssSeason = SpreadsheetApp.openById(seasonId);
  var centralId = _readParamFromSeasonOrGlobal_(ssSeason, 'GLOBAL_MEMBRES_SHEET_ID', '');
  if (!centralId) throw new Error('GLOBAL_MEMBRES_SHEET_ID manquant dans PARAMS.');

  // Détection de nouveauté côté fichier VM
  var folderId = _readParamFromSeasonOrGlobal_(ssSeason, 'DRIVE_FOLDER_VALIDATION_MEMBRES', '');
  if (!folderId) throw new Error('DRIVE_FOLDER_VALIDATION_MEMBRES manquant dans PARAMS (saison).');

  var found = _vm_findLatestActiveFile_(folderId);
  var file  = found && found.file;
  if (!file) return { ok: true, skipped: true, reason: 'Aucun fichier VM' };

  var sp = PropertiesService.getScriptProperties();
  var lastId  = sp.getProperty('VM_LAST_FILE_ID') || '';
  var lastMts = Number(sp.getProperty('VM_LAST_FILE_MTIME') || '0');
  var thisId  = file.getId();
  var thisMts = (file.getLastUpdated && +file.getLastUpdated()) || 0;

  if (!force && lastId === thisId && thisMts <= lastMts) {
    return { ok: true, skipped: true, reason: 'Déjà à jour (même fichier VM)' };
  }

  var res = importValidationMembresToGlobal_(centralId);

  sp.setProperty('VM_LAST_FILE_ID', thisId);
  if (thisMts) sp.setProperty('VM_LAST_FILE_MTIME', String(thisMts));

  return { ok: true, skipped: false, import: res };
}

function API_VM_refreshSeasonSubset() {
  if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
    return { ok: false, busy: true, reason: 'import-running' };
  }
  var seasonId  = getSeasonId_();
  var ssSeason  = SpreadsheetApp.openById(seasonId);
  var centralId = (typeof readParam_==='function' ? readParam_(ssSeason,'GLOBAL_MEMBRES_SHEET_ID') : '') || '';
  if (!centralId) throw new Error('GLOBAL_MEMBRES_SHEET_ID manquant dans PARAMS.');

  // 1) filtre → écrit SAISON.MEMBRES_GLOBAL (même nom)
  var subsetRes = syncMembresGlobalSubsetFromCentral_(seasonId, centralId);

  // 2) propage vers JOUEURS (PhotoExpireLe)
  var applied   = applySeasonMembresToJoueurs_();

  // 3) recalc PhotoStr
  var photoRefreshed = null;
  try {
    if (typeof refreshPhotoStrInJoueurs_ === 'function') {
      photoRefreshed = refreshPhotoStrInJoueurs_(ssSeason);
      // Si la fonction n'a rien écrit (ou a échoué silencieusement), fallback
      if (!photoRefreshed || photoRefreshed.updated === 0) {
        photoRefreshed = _fallbackRefreshPhotoStr_(ssSeason);
      }
    } else if (typeof refreshJoueursPhotoStr_ === 'function') {
      photoRefreshed = refreshJoueursPhotoStr_(ssSeason);
      if (!photoRefreshed || photoRefreshed.updated === 0) {
        photoRefreshed = _fallbackRefreshPhotoStr_(ssSeason);
      }
    } else {
      photoRefreshed = _fallbackRefreshPhotoStr_(ssSeason);
    }
  } catch (e) {
    // quoi qu’il arrive, on assure un recalcul
    photoRefreshed = _fallbackRefreshPhotoStr_(ssSeason);
  }

  return { ok: true, subset: subsetRes, appliedToJoueurs: applied, photoStatus: photoRefreshed };
}

function API_VM_fullRefresh(force) {
  if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
    return { ok: false, busy: true, reason: 'import-running' };
  }
  var r1 = API_VM_importToGlobal(!!force);   // peut retourner skipped:true
  var r2 = API_VM_refreshSeasonSubset();     // subset + application JOUEURS
  return { ok: true, steps: { importGlobal: r1, refreshSubset: r2 } };
}

// MEMBRES_GLOBAL (saison) -> JOUEURS : met à jour PhotoExpireLe + (optionnel) flag/texte
function applySeasonMembresToJoueurs_() {
  var seasonId = getSeasonId_();
  var ss = SpreadsheetApp.openById(seasonId);

  var mgName = (typeof readParam_==='function' ? readParam_(ss, 'SHEET_MEMBRES_GLOBAL') : '') || 'MEMBRES_GLOBAL';
  var jName  = (typeof readParam_==='function' ? readParam_(ss, 'SHEET_JOUEURS') : '') || 'JOUEURS';

  var shMG = ss.getSheetByName(mgName);
  var shJ  = ss.getSheetByName(jName);
  if (!shMG) throw new Error('SAISON: "'+mgName+'" introuvable.');
  if (!shJ)  throw new Error('SAISON: "'+jName+'" introuvable.');

  var lrM = shMG.getLastRow(), lcM = shMG.getLastColumn();
  if (lrM < 2) return { updated: 0, reason: 'MEMBRES_GLOBAL vide' };
  var MG = shMG.getRange(1,1,lrM,lcM).getValues(); var HM = MG.shift();

  var lrJ = shJ.getLastRow(), lcJ = shJ.getLastColumn();
  if (lrJ < 2) return { updated: 0, reason: 'JOUEURS vide' };
  var J = shJ.getRange(1,1,lrJ,lcJ).getValues(); var HJ = J.shift();

  function idxOf(H, list) {
    var low = {}; H.forEach(function(h,i){ low[String(h).trim().toLowerCase()] = i; });
    for (var k=0;k<list.length;k++) { var n = String(list[k]).toLowerCase(); if (n in low) return low[n]; }
    for (var k2=0;k2<list.length;k2++) { var needle = String(list[k2]).toLowerCase(); for (var key in low) if (key.indexOf(needle)!==-1) return low[key]; }
    return -1;
  }

  var iPassM = idxOf(HM, ['passeport #','passeport','passport','no passeport']);
  var iExpM  = idxOf(HM, ['photoexpirele',"date d'expiration de la photo","photo expire le"]);
  if (iPassM < 0 || iExpM < 0) return { updated: 0, reason: 'colonnes manquantes dans MEMBRES_GLOBAL' };

  var iPassJ = idxOf(HJ, ['passeport #','passeport','passport','no passeport']);
  var iExpJ  = idxOf(HJ, ['photoexpirele',"date d'expiration de la photo","photo expire le"]);
  var iStrJ  = idxOf(HJ, ['photostr','photo str','statut photo']);
  if (iPassJ < 0) return { updated: 0, reason: 'colonne Passeport # introuvable dans JOUEURS' };

  // index MEMBRES_GLOBAL par passeport
  var m = new Map();
  for (var r=0;r<MG.length;r++) {
    var p = String(MG[r][iPassM]||'').trim(); if (!p) continue;
    m.set(p, MG[r][iExpM]);
  }

  // MAJ JOUEURS
  var updates = 0;
  for (var r2=0;r2<J.length;r2++) {
    var p2 = String(J[r2][iPassJ]||'').trim(); if (!p2 || !m.has(p2)) continue;
    var exp = m.get(p2);

    var changed = false;
    if (iExpJ >= 0) {
      var cur = J[r2][iExpJ];
      var same = (cur && exp && String(cur)===String(exp)) || (!cur && !exp);
      if (!same) { J[r2][iExpJ] = exp || ''; changed = true; }
    }

    // si on a aussi une colonne PhotoStr, on la recalculera après via source.js
    if (changed) updates++;
  }

  if (updates) {
    shJ.getRange(2,1,J.length,HJ.length).setValues(J);
  }
  return { updated: updates, rows: J.length };
}
// -- Helpers d'entêtes souples
function _idxOfHeaderSoft_(H, wantedList) {
  var map = {};
  H.forEach(function(h,i){ map[String(h||'').trim().toLowerCase()] = i; });
  // exact
  for (var k=0;k<wantedList.length;k++) {
    var w = String(wantedList[k]).trim().toLowerCase();
    if (w in map) return map[w];
  }
  // contains
  for (var j=0;j<wantedList.length;j++) {
    var needle = String(wantedList[j]).trim().toLowerCase();
    for (var key in map) if (key.indexOf(needle) !== -1) return map[key];
  }
  return -1;
}

// -- Fallback: recalcule et écrit PhotoStr si la fonction de source.js n'a pas pu le faire.
function _fallbackRefreshPhotoStr_(ss) {
  ss = ss || SpreadsheetApp.openById(getSeasonId_());
  var sh = ss.getSheetByName('JOUEURS');
  if (!sh || sh.getLastRow() < 2) return { ok:false, reason:'JOUEURS vide' };

  var n = sh.getLastRow() - 1, lc = sh.getLastColumn();
  var H  = sh.getRange(1,1,1,lc).getValues()[0];
  var V  = sh.getRange(2,1,n,lc).getValues();

  var iExp  = _idxOfHeaderSoft_(H, ['PhotoExpireLe',"Date d'expiration de la photo","Photo expire le"]);
  var iAge  = _idxOfHeaderSoft_(H, ['Age','Âge']);
  var iAda  = _idxOfHeaderSoft_(H, ['isAdapte','Adapté','Programme adapté']);
  var iHasI = _idxOfHeaderSoft_(H, ['hasInscription','Inscription','A une inscription']);
  var iStr  = _idxOfHeaderSoft_(H, ['PhotoStr','Statut photo','Photo Str']);

  if (iStr < 0) return { ok:false, reason:'colonne PhotoStr/Statut photo introuvable' };

  // cutoff: on réutilise ta logique si dispo, sinon aujourd’hui
  var cutoffAbs = (typeof readParam_==='function' && typeof PARAM_KEYS!=='undefined')
    ? (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '')
    : '';
  var cutoff = cutoffAbs ? new Date(cutoffAbs)
              : (typeof _getPhotoCutoffDate_==='function' ? _getPhotoCutoffDate_(ss) : new Date());

  function truthy(v){
    v = String(v||'').toUpperCase();
    return v==='1'||v==='TRUE'||v==='OUI'||v==='YES';
  }
  function needPhoto(ageVal, adapteVal, hasInsVal) {
    var age = parseInt(String(ageVal||''),10);
    if (!isNaN(age) && age < 8) return false;
    if (truthy(adapteVal)) return false;
    if (iHasI >= 0) return truthy(hasInsVal);
    return true; // si la colonne n’existe pas, on assume requis
  }
  function statusFor(expVal) {
    if (!expVal && expVal !== 0) return 'Aucune photo';
    var d = (expVal instanceof Date) ? expVal : new Date(expVal);
    if (isNaN(+d)) return 'Aucune photo';
    return (d < cutoff) ? 'Expirée' : 'Valide';
  }

  var out = new Array(n);
  for (var r=0;r<n;r++){
    var age  = (iAge>=0)  ? V[r][iAge]  : '';
    var ada  = (iAda>=0)  ? V[r][iAda]  : '';
    var hasI = (iHasI>=0) ? V[r][iHasI] : '1';
    var exp  = (iExp>=0)  ? V[r][iExp]  : '';
    out[r] = [ needPhoto(age,ada,hasI) ? statusFor(exp) : 'Non requis' ];
  }

  sh.getRange(2, iStr+1, n, 1).setValues(out);
  return { ok:true, updated:n };
}
