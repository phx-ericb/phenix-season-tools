/**
 * v0.5 — Import & Archives
 * Appel : Library.importerDonneesSaison('<ID_CLASSEUR_SAISON>')cd dash
 */

function importerDonneesSaison(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  ensureCoreSheets_(ss);

  // --- PARAMS ---
  var folderId      = readParam_(ss, PARAM_KEYS.DRIVE_FOLDER_IMPORTS);
  var patInsc       = readParam_(ss, PARAM_KEYS.FILE_PATTERN_INSCRIPTIONS) || 'inscriptions';
  var patArt        = readParam_(ss, PARAM_KEYS.FILE_PATTERN_ARTICLES)     || 'articles';
  var dryRun        = (readParam_(ss, PARAM_KEYS.DRY_RUN) || 'FALSE').toUpperCase() === 'TRUE';
  var moveConverted = (readParam_(ss, PARAM_KEYS.MOVE_CONVERTED_TO_ARCHIVE) || 'FALSE').toUpperCase() === 'TRUE';
  var incrOn        = (readParam_(ss, PARAM_KEYS.INCREMENTAL_ON) || 'TRUE').toUpperCase() === 'TRUE';

  if (!folderId) {
    throw new Error('PARAM manquant : ' + PARAM_KEYS.DRIVE_FOLDER_IMPORTS);
  }

  // --- Scan des fichiers sources ---
  var folder    = DriveApp.getFolderById(folderId);
  var filesInfo = scanImportFiles_(folder, [patInsc, patArt]);

  // --- 1) Conversion -> STAGING ---
  var converted = [];
  filesInfo.forEach(function (info) {
    var target = inferTargetFromName_(info.name, patInsc, patArt);
    if (!target) return;

    try {
      var tempSheetId = ensureSpreadsheetFromFile_(info.id, folderId);
      appendImportLog_(ss, 'CONVERT_ID', info.name + ' -> tempId=' + tempSheetId);

      var values = readFirstSheetValues_(tempSheetId); // tableau 2D
      var stagingName = (target === 'INSCRIPTIONS') ? SHEETS.STAGING_INSCRIPTIONS : SHEETS.STAGING_ARTICLES;

      // Écrit le staging avec nos normalisations (Passeport # texte + Saison injectée si absente)
      writeStaging_(ss, stagingName, values);

      appendImportLog_(ss, 'CONVERT_OK',
        info.name + ' -> ' + stagingName + ' (' + (values ? values.length : 0) + ' lignes)'
      );

      converted.push({ fileId: info.id, name: info.name, target: target, tempId: tempSheetId });
    } catch (e) {
      var msg = (e && e.message) ? e.message : ('' + e);
      if (e && e.name === 'UNSUPPORTED_MIME') {
        appendImportLog_(ss, 'SKIP_UNSUPPORTED_MIME', info.name + ' (mime=' + info.mime + ')');
      } else if (e && e.name === 'CONVERTED_NOT_SHEET') {
        appendImportLog_(ss, 'SKIP_BAD_CONVERSION', info.name + ' -> ' + msg);
      } else {
        appendImportLog_(ss, 'CONVERT_FAIL', info.name + ' -> ' + msg);
      }
    }
  });

  // --- 2) STAGING -> FINALS ---
  if (!incrOn) {
    // === MODE FULL REFRESH (legacy) ===
    if (converted.some(function (c) { return c.target === 'INSCRIPTIONS'; })) {
      var shStInsc = ss.getSheetByName(SHEETS.STAGING_INSCRIPTIONS);
      if (shStInsc && shStInsc.getLastRow() > 0) {
        var vInsc = shStInsc.getRange(1, 1, shStInsc.getLastRow(), shStInsc.getLastColumn()).getValues();
        vInsc = prefixFirstColWithApostrophe_(vInsc); // legacy : garde le comportement historique
        var shInsc = getSheetOrCreate_(ss, SHEETS.INSCRIPTIONS);
        overwriteSheet_(shInsc, vInsc);
        appendImportLog_(ss, 'WRITE_OK', 'INSCRIPTIONS <- STAGING (' + vInsc.length + ' lignes)');
      }
    }
    if (converted.some(function (c) { return c.target === 'ARTICLES'; })) {
      var shStArt = ss.getSheetByName(SHEETS.STAGING_ARTICLES);
      if (shStArt && shStArt.getLastRow() > 0) {
        var vArt = shStArt.getRange(1, 1, shStArt.getLastRow(), shStArt.getLastColumn()).getValues();
        var shArt = getSheetOrCreate_(ss, SHEETS.ARTICLES);
        overwriteSheet_(shArt, vArt);
        appendImportLog_(ss, 'WRITE_OK', 'ARTICLES <- STAGING (' + vArt.length + ' lignes)');
      }
    }
  } else {
    // === MODE INCRÉMENTAL (v0.7) ===
    applyIncrementalForInscriptions_(seasonSheetId);
    applyIncrementalForArticles_(seasonSheetId);
  }

  // --- 3) Archivage des sources (+ Converted si demandé) ---
  if (!dryRun && converted.length) {
    var archiveFolder = ensureArchiveSubfolder_(folderId);
    converted.forEach(function (c) {
      moveFileIdToFolderIdV3_(c.fileId, archiveFolder.getId());
    });
    appendImportLog_(ss, 'ARCHIVE_OK', 'Fichiers sources déplacés vers ' + archiveFolder.getName());

    if (moveConverted) {
      var convertedFolder = ensureSubfolderUnder_(archiveFolder, 'Converted');
      converted.forEach(function (c) {
        try {
          moveFileIdToFolderIdV3_(c.tempId, convertedFolder.getId());
        } catch (err) {
          appendImportLog_(ss, 'ARCHIVE_CONVERTED_FAIL', c.name + ' -> ' + err);
        }
      });
      appendImportLog_(ss, 'ARCHIVE_CONVERTED_OK', 'Feuilles [CONVERTI] déplacées vers ' + convertedFolder.getName());
    }
  } else if (dryRun) {
    appendImportLog_(ss, 'ARCHIVE_SKIP', 'DRY_RUN=TRUE (pas d’archivage)');
  }

  // --- 4) Récap SCAN_OK (avec vrais totaux en finals, même en INCR) ---
  var rowsInsc = 0, rowsArt = 0;
  try {
    var shInsc2 = ss.getSheetByName(SHEETS.INSCRIPTIONS);
    if (shInsc2) rowsInsc = Math.max(0, shInsc2.getLastRow() - 1);
    var shArt2 = ss.getSheetByName(SHEETS.ARTICLES);
    if (shArt2) rowsArt = Math.max(0, shArt2.getLastRow() - 1);
  } catch (e2) {
    // no-op
  }

  var summary = Utilities.formatString(
    'Fichiers détectés: %s, INSCRIPTIONS: %s lignes, ARTICLES: %s lignes%s%s',
    filesInfo.length,
    rowsInsc,
    rowsArt,
    dryRun ? ' (DRY_RUN ON)' : '',
    incrOn ? ' [INCR]' : ' [FULL]'
  );
  appendImportLog_(ss, 'SCAN_OK', summary);
  // À la toute fin, après SCAN_OK
try { evaluateSeasonRules(seasonSheetId || getSeasonId_()); } catch(e) { appendImportLog_(ss, 'RULES_FAIL', ''+e); }

  return summary;
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

    // 1) ignorer ce qu'on a créé nous-mêmes
    if (name.indexOf('[CONVERTI]') === 0) continue;

    // 2) matcher par nom
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
  appendImportLog_(ss, 'INIT_SAISON', 'Feuilles socle vérifiées/créées.');
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
  if (lower.endsWith('.xls'))  return 'application/vnd.ms-excel';
  if (lower.endsWith('.csv'))  return 'text/csv';
  return null;
}

/**
 * Retourne l'ID d'un Spreadsheet lisible:
 * - si fileId est déjà un Google Sheet -> retourne fileId
 * - si XLS/XLSX/CSV -> crée un Google Sheet converti dans le même dossier et retourne son ID
 * - sinon -> lève "UNSUPPORTED_MIME:<mime>"
 */
function ensureSpreadsheetFromFile_(fileId, parentFolderId) {
  var src  = DriveApp.getFileById(fileId);
  var mime = src.getMimeType();

  // 1) Déjà un Google Sheet ?
  if (isSpreadsheetMime_(mime)) {
    return src.getId();
  }

  // 2) Convertissable ?
  if (!isConvertibleSpreadsheetMime_(mime)) {
    var e = new Error('UNSUPPORTED_MIME:' + mime);
    e.name = 'UNSUPPORTED_MIME';
    throw e;
  }

  // 3) v3 create avec content-type explicite
  if (!(Drive && Drive.Files && typeof Drive.Files.create === 'function')) {
    throw new Error('Drive v3 non disponible. Active le service avancé Drive (Library + consommateur) et redéploie.');
  }

  var blob = src.getBlob();
  var ct = extToContentType_(src.getName());
  if (ct) blob.setContentType(ct); // aide l'importeur Drive

  var resource = {
    name: '[CONVERTI] ' + src.getName().replace(/\.xlsx?$/i, '').replace(/\.csv$/i, ''),
    mimeType: 'application/vnd.google-apps.spreadsheet',
    parents: [parentFolderId]
  };

  var created = Drive.Files.create(resource, blob, { supportsAllDrives: true });
  Utilities.sleep(800); // marge propagation

  // 4) Validation post-création : doit être un Spreadsheet
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
  // 1) Vérifier le mimeType côté Drive (doit être un google-apps.spreadsheet)
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
    throw new Error('Impossible de valider le fichier converti (' + spreadsheetId + '): ' + e);
  }

  // 2) Ouverture avec retries exponentiels (propagation Drive -> Spreadsheet)
  var lastErr = null;
  for (var attempt = 1; attempt <= 5; attempt++) {
    try {
      var ss = SpreadsheetApp.openById(spreadsheetId);
      var shs = ss.getSheets();
      if (!shs || !shs.length) return [];
      var sh = shs[0];
      var lastRow = sh.getLastRow();
      var lastCol = sh.getLastColumn();
      if (lastRow === 0 || lastCol === 0) return [];
      return sh.getRange(1, 1, lastRow, lastCol).getValues();
    } catch (err) {
      lastErr = err;
      var waitMs = [500, 1000, 2000, 3000, 5000][attempt - 1];
      Utilities.sleep(waitMs);
    }
  }
  throw new Error('SpreadsheetApp.openById a échoué après retries (' + spreadsheetId + '): ' + lastErr);
}

/** Détermine la cible (INSCRIPTIONS/ARTICLES) en fonction du nom de fichier et des patterns */
function inferTargetFromName_(fileName, patInsc, patArt) {
  var n = (fileName || '').toLowerCase();
  if (patInsc && n.indexOf(patInsc.toLowerCase()) !== -1) return 'INSCRIPTIONS';
  if (patArt && n.indexOf(patArt.toLowerCase()) !== -1) return 'ARTICLES';
  return null;
}

/**
 * Déplace les fichiers [CONVERTI] du dossier d’import vers Archives/YYYY-MM-DD/Converted/
 * Déplace aussi tout autre fichier non-Sheet (erreur de conversion) vers Archives/YYYY-MM-DD/BadConversions/
 */
function nettoyerConversionsSaison(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var folderId = readParam_(ss, PARAM_KEYS.DRIVE_FOLDER_IMPORTS);
  if (!folderId) {
    throw new Error('PARAM manquant : ' + PARAM_KEYS.DRIVE_FOLDER_IMPORTS);
  }

  var importsFolder = DriveApp.getFolderById(folderId);
  var archiveFolder = ensureArchiveSubfolder_(folderId);
  var convFolder = ensureSubfolderUnder_(archiveFolder, 'Converted');
  var badFolder  = ensureSubfolderUnder_(archiveFolder, 'BadConversions');

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

  appendImportLog_(ss, 'CLEAN_CONVERTED', JSON.stringify(moved));
  return moved;
}

/* ======================= v0.7: wrappers INCR ======================= */
/** Appelle le diff INSCRIPTIONS et log le résultat, pour compat avec import.js */
function applyIncrementalForInscriptions_(seasonSheetId) {
  if (typeof diffInscriptions_ !== 'function') {
    throw new Error('diffInscriptions_ manquant. Assure-toi d’avoir inclus diff_inscriptions.gs v0.7.');
  }
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var res = diffInscriptions_(seasonSheetId);
  appendImportLog_(ss, 'INCR_INSCRIPTIONS_OK', JSON.stringify(res));
  return res;
}
/** Appelle le diff ARTICLES et log le résultat, pour compat avec import.js */
function applyIncrementalForArticles_(seasonSheetId) {
  if (typeof diffArticles_ !== 'function') {
    throw new Error('diffArticles_ manquant. Assure-toi d’avoir inclus diff_articles.gs v0.7.');
  }
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var res = diffArticles_(seasonSheetId);
  appendImportLog_(ss, 'INCR_ARTICLES_OK', JSON.stringify(res));
  return res;
}
