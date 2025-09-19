/**
 * v0.8 — Import & Archives (incrémental v1)
 * - Logger compatible (gère appendImportLog_ 3-args ou 1-arg)
 * - Après diffs: collecte des passeports touchés depuis STAGING
 * - Maintenance ERREURS ciblée (dédup + purge erreurs résolues) pour les passeports touchés
 * - Exports rétro: Membres + Groupes (ALL)
 */
// ============ import.js ============
// v0.8 — Import & Archives (incrémental v1) — VERSION NETTOYÉE
/**
 * v0.9 — Import ONLY
 * - Scan + convert → STAGING
 * - STAGING → FINALS (FULL ou INCR selon INCREMENTAL_ON)
 * - Archivage (respecte DRY_RUN / MOVE_CONVERTED_TO_ARCHIVE)
 * - Résumé + persist des passeports "touchés" (pour l’INCR)
 * - Aucune règle / aucun queue mail / aucun worker / aucun export ici
 */
function importerDonneesSaison(seasonSheetId) {
  var ss  = getSeasonSpreadsheet_(seasonSheetId);
  ensureCoreSheets_(ss);
  var sid = seasonSheetId || getSeasonId_();

  // --- PARAMS ---
  var folderId   = readParam_(ss, PARAM_KEYS.DRIVE_FOLDER_IMPORTS);
  var patInsc    = readParam_(ss, PARAM_KEYS.FILE_PATTERN_INSCRIPTIONS) || 'inscriptions';
  var patArt     = readParam_(ss, PARAM_KEYS.FILE_PATTERN_ARTICLES)     || 'articles';
  var dryRun     = (readParam_(ss, PARAM_KEYS.DRY_RUN) || 'FALSE').toUpperCase() === 'TRUE';
  var moveConv   = (readParam_(ss, PARAM_KEYS.MOVE_CONVERTED_TO_ARCHIVE) || 'FALSE').toUpperCase() === 'TRUE';
  var incrOn     = (readParam_(ss, PARAM_KEYS.INCREMENTAL_ON) || 'TRUE').toUpperCase() === 'TRUE';

  if (!folderId) throw new Error('PARAM manquant : ' + PARAM_KEYS.DRIVE_FOLDER_IMPORTS);

  // --- Scan des fichiers sources ---
  var folder = DriveApp.getFolderById(folderId);
  var filesInfo = scanImportFiles_(folder, [patInsc, patArt]);

  // --- 1) Conversion -> STAGING ---
  var converted = [];
  filesInfo.forEach(function (info) {
    var target = inferTargetFromName_(info.name, patInsc, patArt);
    if (!target) return;

    try {
      var tempSheetId = ensureSpreadsheetFromFile_(info.id, folderId);
      log_(ss, 'CONVERT_ID', info.name + ' -> tempId=' + tempSheetId);

      var values = readFirstSheetValues_(tempSheetId); // tableau 2D
      var stagingName = (target === 'INSCRIPTIONS') ? SHEETS.STAGING_INSCRIPTIONS : SHEETS.STAGING_ARTICLES;

      // Normalisations (Passeport en texte + Saison injectée si absente)
      writeStaging_(ss, stagingName, values);

      log_(ss, 'CONVERT_OK', info.name + ' -> ' + stagingName + ' (' + (values ? values.length : 0) + ' lignes)');
      converted.push({ fileId: info.id, name: info.name, target: target, tempId: tempSheetId });
    } catch (e) {
      var msg = (e && e.message) ? e.message : String(e);
      if (e && e.name === 'UNSUPPORTED_MIME')         log_(ss, 'SKIP_UNSUPPORTED_MIME', info.name + ' (mime=' + info.mime + ')');
      else if (e && e.name === 'CONVERTED_NOT_SHEET') log_(ss, 'SKIP_BAD_CONVERSION', info.name + ' -> ' + msg);
      else                                            log_(ss, 'CONVERT_FAIL', info.name + ' -> ' + msg);
    }
  });

  // --- 2) STAGING -> FINALS ---
  var incrResInsc = null, incrResArt = null;
  if (!incrOn) {
    // FULL overwrite: copie brute du STAGING
    if (converted.some(function (c) { return c.target === 'INSCRIPTIONS'; })) {
      var shStInsc = ss.getSheetByName(SHEETS.STAGING_INSCRIPTIONS);
      if (shStInsc && shStInsc.getLastRow() > 0) {
        var vInsc = shStInsc.getRange(1,1, shStInsc.getLastRow(), shStInsc.getLastColumn()).getValues();
        vInsc = prefixFirstColWithApostrophe_(vInsc); // compat legacy
        overwriteSheet_(getSheetOrCreate_(ss, SHEETS.INSCRIPTIONS), vInsc);
        log_(ss, 'WRITE_OK', 'INSCRIPTIONS <- STAGING (' + vInsc.length + ' lignes)');
      }
    }
    if (converted.some(function (c) { return c.target === 'ARTICLES'; })) {
      var shStArt = ss.getSheetByName(SHEETS.STAGING_ARTICLES);
      if (shStArt && shStArt.getLastRow() > 0) {
        var vArt = shStArt.getRange(1,1, shStArt.getLastRow(), shStArt.getLastColumn()).getValues();
        overwriteSheet_(getSheetOrCreate_(ss, SHEETS.ARTICLES), vArt);
        log_(ss, 'WRITE_OK', 'ARTICLES <- STAGING (' + vArt.length + ' lignes)');
      }
    }
  } else {
    // INCR: applique diffs (renvoie touchedPassports pour chaque table)
    incrResInsc = applyIncrementalForInscriptions_(sid);
    incrResArt  = applyIncrementalForArticles_(sid);
  }

  // --- 3) Archivage des sources (+ Converted si demandé) ---
  if (!dryRun && converted.length) {
    var archiveFolder = ensureArchiveSubfolder_(folderId);
    converted.forEach(function (c) {
      moveFileIdToFolderIdV3_(c.fileId, archiveFolder.getId());
    });
    log_(ss, 'ARCHIVE_OK', 'Fichiers sources déplacés vers ' + archiveFolder.getName());

    if (moveConv) {
      var convertedFolder = ensureSubfolderUnder_(archiveFolder, 'Converted');
      converted.forEach(function (c) {
        try { moveFileIdToFolderIdV3_(c.tempId, convertedFolder.getId()); }
        catch (err) { log_(ss, 'ARCHIVE_CONVERTED_FAIL', c.name + ' -> ' + err); }
      });
      log_(ss, 'ARCHIVE_CONVERTED_OK', 'Feuilles [CONVERTI] déplacées vers ' + convertedFolder.getName());
    }
  } else if (dryRun) {
    log_(ss, 'ARCHIVE_SKIP', 'DRY_RUN=TRUE (pas d’archivage)');
  }

  // --- 4) Récap SCAN_OK ---
  var rowsInsc = 0, rowsArt = 0;
  try {
    var shInsc2 = ss.getSheetByName(SHEETS.INSCRIPTIONS);
    if (shInsc2) rowsInsc = Math.max(0, shInsc2.getLastRow() - 1);
    var shArt2 = ss.getSheetByName(SHEETS.ARTICLES);
    if (shArt2) rowsArt = Math.max(0, shArt2.getLastRow() - 1);
  } catch (_e2) {}
  var summaryText = Utilities.formatString(
    'Fichiers détectés: %s, INSCRIPTIONS: %s lignes, ARTICLES: %s lignes%s%s',
    filesInfo.length, rowsInsc, rowsArt, (dryRun ? ' (DRY_RUN ON)' : ''), (incrOn ? ' [INCR]' : ' [FULL]')
  );
  log_(ss, 'SCAN_OK', summaryText);

  // --- 4b) Persistance des "touched" (INCR seulement)
  var touchedFromDiffs = unionPassportArrays_(
    (incrResInsc && incrResInsc.touchedPassports) || [],
    (incrResArt  && incrResArt.touchedPassports)  || []
  );
  try {
    var props = PropertiesService.getDocumentProperties();
    props.setProperty('LAST_TOUCHED_PASSPORTS', (touchedFromDiffs || []).join(','));
    props.setProperty('LAST_TOUCHED_PASSPORTS_JSON', JSON.stringify(touchedFromDiffs || []));
  } catch (__){}

  // IMPORTANT : on s'arrête ici. Pas de règles, pas de trim, pas de queue, pas de worker, pas d'exports.
  return {
    ok: true,
    filesDetected: filesInfo.length,
    rowsInscriptions: rowsInsc,
    rowsArticles: rowsArt,
    mode: incrOn ? 'INCR' : 'FULL',
    converted: converted.map(function(c){ return { name: c.name, target: c.target }; }),
    touchedPassports: touchedFromDiffs
  };
}




/* ========= Logger compatible (1-arg ou 3-args) ========= */
function log_(ssOrNull, type, details) {
  try {
    if (typeof appendImportLog_ === 'function') {
      // Essai API «lib» (ss, type, details)
      try { appendImportLog_(ssOrNull, type, details); return; } catch (_e) { }
      // Repli API «dashboard» ({type, details})
      try { appendImportLog_({ type: type || 'INFO', details: details || '' }); return; } catch (_e2) { }
    }
  } catch (_) { /* no-op */ }
  // Dernier repli : écriture directe
  try {
    var ss = ssOrNull && ssOrNull.getId ? ssOrNull : SpreadsheetApp.openById(getSeasonId_());
    var name = 'IMPORT_LOG';
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1, 1, 1, 3).setValues([['Date', 'Type', 'Détails']]);
    sh.appendRow([new Date(), type || 'INFO', details || '']);
  } catch (__) { /* ignore */ }
}

/** Liste les fichiers dont le nom contient l’un des patterns (niveau courant uniquement) */
function scanImportFiles_(folder, namePatterns) {
  var it = folder.getFiles();
  var res = [];
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    var lower = name.toLowerCase();
    var mime = f.getMimeType();

    if (name.indexOf('[CONVERTI]') === 0) continue;

    var matchName = namePatterns.some(function (p) {
      return p && lower.indexOf((p + '').toLowerCase()) !== -1;
    });
    if (!matchName) continue;

    res.push({ id: f.getId(), name: name, url: f.getUrl(), mime: mime });
  }
  return res;
}

/** Optionnel : init rapide d’un classeur saison (crée les feuilles socle si absentes) */
function initSeasonFile(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  ensureCoreSheets_(ss);
  log_(ss, 'INIT_SAISON', 'Feuilles socle vérifiées/créées.');
}

function isSpreadsheetMime_(mime) {
  return mime === 'application/vnd.google-apps.spreadsheet';
}
function isConvertibleSpreadsheetMime_(mime) {
  return mime === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || // .xlsx
    mime === 'application/vnd.ms-excel' ||                                        // .xls
    mime === 'text/csv';                                                          // .csv
}
function extToContentType_(name) {
  var lower = (name || '').toLowerCase();
  if (lower.endsWith('.xlsx')) return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  if (lower.endsWith('.xls')) return 'application/vnd.ms-excel';
  if (lower.endsWith('.csv')) return 'text/csv';
  return null;
}

/**
 * Retourne l'ID d'un Spreadsheet lisible:
 * - si fileId est déjà un Google Sheet -> retourne fileId
 * - si XLS/XLSX/CSV -> crée un Google Sheet converti dans le même dossier et retourne son ID
 * - sinon -> lève "UNSUPPORTED_MIME:<mime>"
 */
function ensureSpreadsheetFromFile_(fileId, parentFolderId) {
  var src = DriveApp.getFileById(fileId);
  var mime = src.getMimeType();

  if (isSpreadsheetMime_(mime)) return src.getId();

  if (!isConvertibleSpreadsheetMime_(mime)) {
    var e = new Error('UNSUPPORTED_MIME:' + mime);
    e.name = 'UNSUPPORTED_MIME';
    throw e;
  }

  if (!(Drive && Drive.Files && typeof Drive.Files.create === 'function')) {
    throw new Error('Drive v3 non disponible. Active le service avancé Drive (Library + consommateur) et redéploie.');
  }

  var blob = src.getBlob();
  var ct = extToContentType_(src.getName());
  if (ct) blob.setContentType(ct);

  var resource = {
    name: '[CONVERTI] ' + src.getName().replace(/\.xlsx?$/i, '').replace(/\.csv$/i, ''),
    mimeType: 'application/vnd.google-apps.spreadsheet',
    parents: [parentFolderId]
  };

  var created = Drive.Files.create(resource, blob, { supportsAllDrives: true });
  Utilities.sleep(800);

  var createdFile = DriveApp.getFileById(created.id);
  var createdMime = createdFile.getMimeType();
  if (!isSpreadsheetMime_(createdMime)) {
    var err = new Error('CONVERTED_NOT_SHEET:' + createdMime);
    err.name = 'CONVERTED_NOT_SHEET';
    throw err;
  }
  return created.id;
}

/**
 * Ouvre un Spreadsheet et lit la première feuille, avec retry si le fichier vient d'être créé.
 */
function readFirstSheetValues_(spreadsheetId) {
  try {
    var f = DriveApp.getFileById(spreadsheetId);
    var mime = f.getMimeType();
    if (mime !== 'application/vnd.google-apps.spreadsheet') {
      Utilities.sleep(1200);
      mime = DriveApp.getFileById(spreadsheetId).getMimeType();
      if (mime !== 'application/vnd.google-apps.spreadsheet') {
        throw new Error('Le fichier converti n’est pas un Google Sheet (mime=' + mime + ').');
      }
    }
  } catch (e) {
    Utilities.sleep(1500);
  }

  var ss = SpreadsheetApp.openById(spreadsheetId);
  var sh = ss.getSheets()[0];
  return sh.getDataRange().getDisplayValues();
}

/** Tente d’inférer la cible (INSCRIPTIONS/ARTICLES) d’après le nom de fichier */
function inferTargetFromName_(name, patInsc, patArt) {
  var lower = (name || '').toLowerCase();
  if (lower.indexOf((patInsc + '').toLowerCase()) >= 0) return 'INSCRIPTIONS';
  if (lower.indexOf((patArt + '').toLowerCase()) >= 0) return 'ARTICLES';
  return null;
}

/* ======================= Incr v1 — apply diffs ======================= */
function applyIncrementalForInscriptions_(seasonSheetId) {
  if (typeof diffInscriptions_ !== 'function') {
    throw new Error('diffInscriptions_ manquant. Assure-toi d’avoir inclus diff_inscriptions.gs v0.7.');
  }
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var res = diffInscriptions_(seasonSheetId);
  log_(ss, 'INCR_INSCRIPTIONS_OK', JSON.stringify(res));
  return res;
}
function applyIncrementalForArticles_(seasonSheetId) {
  if (typeof diffArticles_ !== 'function') {
    throw new Error('diffArticles_ manquant. Assure-toi d’avoir inclus diff_articles.gs v0.7.');
  }
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var res = diffArticles_(seasonSheetId);
  log_(ss, 'INCR_ARTICLES_OK', JSON.stringify(res));
  return res;
}


/**
 * Déplace les fichiers [CONVERTI] du dossier d’import vers Archives/YYYY-MM-DD/Converted/
 * Déplace aussi tout autre fichier non-Sheet (erreur de conversion) vers Archives/YYYY-MM-DD/BadConversions/
 */
function nettoyerConversionsSaison(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var folderId = readParam_(ss, PARAM_KEYS.DRIVE_FOLDER_IMPORTS);
  if (!folderId) throw new Error('PARAM manquant : ' + PARAM_KEYS.DRIVE_FOLDER_IMPORTS);

  var importsFolder = DriveApp.getFolderById(folderId);
  var archiveFolder = ensureArchiveSubfolder_(folderId);
  var convFolder = ensureSubfolderUnder_(archiveFolder, 'Converted');
  var badFolder = ensureSubfolderUnder_(archiveFolder, 'BadConversions');

  var moved = { converted: [], bad: [] };
  var it = importsFolder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    var mime = f.getMimeType();

    if (name.indexOf('[CONVERTI]') === 0) {
      moveFileIdToFolderIdV3_(f.getId(), convFolder.getId());
      moved.converted.push(name);
    } else if (mime !== 'application/vnd.google-apps.spreadsheet') {
      moveFileIdToFolderIdV3_(f.getId(), badFolder.getId());
      moved.bad.push(name);
    }
  }

  log_(ss, 'CLEAN_CONVERTED', JSON.stringify(moved));
  return moved;
}
/* ======================= Incr v1 — helpers ERREURS ======================= */


function trimErreursIncremental_(ssOrId, passports) {
  var ss = ensureSpreadsheet_(ssOrId); // <= CLÉ
  if (!passports || passports.length === 0) {
    log_ && log_('TRIM_ERR_SKIP', { reason: 'no-passports' });
    return { trimmed: 0, resolved: 0, total: 0 };
  }
  var errSh = safeGetSheet_(ss, 'ERREURS');
  var range = errSh.getDataRange();
  var values = range.getValues(); // IMPORTANT: objets Date natifs (pas displayValues)
  if (!values || values.length <= 1) {
    return { trimmed: 0, resolved: 0, total: 0 };
  }

  var header = values[0];
  var idx = {};
  for (var i = 0; i < header.length; i++) idx[header[i]] = i;

  function keyOf(row) {
    return [
      row[idx['Passeport #']] || '',
      row[idx['Type']] || '',
      row[idx['Saison']] || '',
      row[idx['Frais']] || '',
      row[idx['Message']] || ''
    ].join('¦');
  }

  // 1) Déduplication: on garde la plus récente par clé
  var best = new Map();
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var k = keyOf(row);
    var created = row[idx['CreatedAt']] instanceof Date ? row[idx['CreatedAt']] : new Date();
    if (!best.has(k) || created > best.get(k).created) {
      best.set(k, { row: row, r: r, created: created });
    }
  }

  // 2) Reconstruire la table “dédupliquée”
  var dedupRows = [header];
  best.forEach(function (obj) { dedupRows.push(obj.row); });
  if (dedupRows.length !== values.length) {
    errSh.clearContents();
    errSh.getRange(1,1,dedupRows.length,dedupRows[0].length).setValues(dedupRows);
    values = dedupRows;
  }

  // 3) Purge des erreurs “résolues” pour les passeports touchés
  var setTouched = new Set(passports.map(String));
  // Charge INSCRIPTIONS finales 1 seule fois
  var inscSh = safeGetSheet_(ss, 'INSCRIPTIONS');
  var inscVals = inscSh.getDataRange().getValues();
  var inscHeader = inscVals[0] || [];
  var iIdx = {};
  for (var j = 0; j < inscHeader.length; j++) iIdx[inscHeader[j]] = j;
  var inscPassSet = new Set();
  for (var rr = 1; rr < inscVals.length; rr++) {
    var p = (inscVals[rr][iIdx['Passeport #']] || '').toString().trim();
    if (p) inscPassSet.add(p);
  }

  var kept = [header];
  var resolved = 0;
  for (var r2 = 1; r2 < values.length; r2++) {
    var row2 = values[r2];
    var p2 = (row2[idx['Passeport #']] || '').toString().trim();
    if (setTouched.has(p2)) {
      // logique “résolue”: si le passeport existe encore dans INSCRIPTIONS et que l’erreur est de type orphelin/doublon, on peut la purger
      var type = (row2[idx['Type']] || '').toString();
      if (type === 'MembreIntrouvable' || type === 'DoubleInscription') {
        if (inscPassSet.has(p2)) { resolved++; continue; }
      }
      // autres types: on garde (elles ont été recalculées juste avant de toute façon)
    }
    kept.push(row2);
  }

  if (kept.length !== values.length) {
    errSh.clearContents();
    errSh.getRange(1,1,kept.length,kept[0].length).setValues(kept);
  }

  return { trimmed: values.length - kept.length, resolved: resolved, total: kept.length - 1 };
}


/* ======================= Utilitaires incrémentaux (union + cache) ======================= */
function unionPassportArrays_() {
  var set = {};
  for (var i = 0; i < arguments.length; i++) {
    var arr = arguments[i] || [];
    for (var j = 0; j < arr.length; j++) {
      var p = (arr[j] || '').toString().trim();
      if (!p) continue;
      set[p] = true;
    }
  }
  return Object.keys(set);
}
function getLastTouchedPassports_() {
  try {
    var raw = (PropertiesService.getDocumentProperties().getProperty('LAST_TOUCHED_PASSPORTS') || '').trim();
    if (!raw) return [];
    if (raw.charAt(0) === '[') {
      var arr = JSON.parse(raw);
      return (Array.isArray(arr) ? arr : []).map(function (x) { return String(x || '').trim(); }).filter(Boolean);
    }
    // CSV
    return raw.split(',').map(function (x) { return String(x || '').trim(); }).filter(Boolean);
  } catch (e) {
    return [];
  }
}
