/**
 * v0.8 — Import & Archives (incrémental v1)
 * - Logger compatible (gère appendImportLog_ 3-args ou 1-arg)
 * - Après diffs: collecte des passeports touchés depuis STAGING
 * - Maintenance ERREURS ciblée (dédup + purge erreurs résolues) pour les passeports touchés
 * - Exports rétro: Membres + Groupes (ALL)
 */
// ============ import.js ============
// v0.8 — Import & Archives (incrémental v1) — VERSION NETTOYÉE
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

      // Normalisations (Passeport texte + Saison injectée si absente)
      writeStaging_(ss, stagingName, values);

      log_(ss, 'CONVERT_OK', info.name + ' -> ' + stagingName + ' (' + (values ? values.length : 0) + ' lignes)');
      converted.push({ fileId: info.id, name: info.name, target: target, tempId: tempSheetId });
    } catch (e) {
      var msg = (e && e.message) ? e.message : String(e);
      if (e && e.name === 'UNSUPPORTED_MIME')        log_(ss, 'SKIP_UNSUPPORTED_MIME', info.name + ' (mime=' + info.mime + ')');
      else if (e && e.name === 'CONVERTED_NOT_SHEET') log_(ss, 'SKIP_BAD_CONVERSION', info.name + ' -> ' + msg);
      else                                            log_(ss, 'CONVERT_FAIL', info.name + ' -> ' + msg);
    }
  });

  // --- 2) STAGING -> FINALS ---
  var incrResInsc = null, incrResArt = null;
  if (!incrOn) {
    if (converted.some(function (c) { return c.target === 'INSCRIPTIONS'; })) {
      var shStInsc = ss.getSheetByName(SHEETS.STAGING_INSCRIPTIONS);
      if (shStInsc && shStInsc.getLastRow() > 0) {
        var vInsc = shStInsc.getRange(1,1, shStInsc.getLastRow(), shStInsc.getLastColumn()).getValues();
        vInsc = prefixFirstColWithApostrophe_(vInsc); // legacy compat
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
    // Incrémental
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
  var summary = Utilities.formatString(
    'Fichiers détectés: %s, INSCRIPTIONS: %s lignes, ARTICLES: %s lignes%s%s',
    filesInfo.length, rowsInsc, rowsArt, (dryRun ? ' (DRY_RUN ON)' : ''), (incrOn ? ' [INCR]' : ' [FULL]')
  );
  log_(ss, 'SCAN_OK', summary);

  // --- 4b) Calcul "touchés" + maintenance ERREURS ciblée (v2 — diff-only)
  try {
    // Uniquement les passeports touchés par les diffs incrémentaux
    var touchedFromDiffs = unionPassportArrays_(
      (incrResInsc && incrResInsc.touchedPassports) || [],
      (incrResArt && incrResArt.touchedPassports) || []
    );
    var touchedAll = touchedFromDiffs; // on NE scanne PLUS le STAGING ici

    // Persist pour debug / autres jobs (CSV + JSON)
    var props = PropertiesService.getDocumentProperties();
    try {
      props.setProperty('LAST_TOUCHED_PASSPORTS', (touchedAll || []).join(','));
      props.setProperty('LAST_TOUCHED_PASSPORTS_JSON', JSON.stringify(touchedAll || []));
    } catch (__){}

    // Règles ciblées (incrémental)
    try {
      evaluateSeasonRules(seasonSheetId || getSeasonId_(), touchedAll);
    } catch (e) {
      log_(ss, 'RULES_FAIL_WRAP', '' + e);
    }

    // Maintenance erreurs ciblée (dédup + purge des résolues) sur les mêmes passeports
    try {
      var trimRes = trimErreursIncremental_(seasonSheetId || getSeasonId_(), touchedAll);
      log_(ss, 'ERR_MAINTENANCE', JSON.stringify(trimRes));
    } catch (eTrim) {
      log_(ss, 'ERR_MAINTENANCE_FAIL', '' + eTrim);
    }
  } catch (eOuter) {
    log_(ss, 'ERR_MAINTENANCE_WRAP_FAIL', '' + eOuter);
  }


  // --- 5) Enqueue des courriels ---
  try { enqueueInscriptionNewBySectors(sid); }
  catch (eQ) { log_(ss, 'QUEUE_NEW_FAIL', String(eQ)); }

  ['U9_12_SANS_CDP', 'U7_8_SANS_2E_SEANCE'].forEach(function (code) {
    try {
      var qRes = enqueueValidationEmailsByErrorCode(sid, code);
      log_(ss, 'QUEUE_ERRMAIL_OK', JSON.stringify({ code: code, queued: qRes.queued }));
    } catch (eQE) {
      log_(ss, 'QUEUE_ERRMAIL_FAIL', JSON.stringify({ code: code, error: String(eQE) }));
    }
  });

  // --- 6) Worker d’envoi ---
  try {
    log_(ss, 'MAIL_WORKER_BEGIN', 'start');
    var workerRes = sendPendingOutbox(sid);
    log_(ss, 'MAIL_WORKER_DONE', JSON.stringify(workerRes));
  } catch (eW) {
    log_(ss, 'MAIL_WORKER_FAIL', String(eW));
  }

  // --- 7) Exports rétro ---
  // ⚠️ SUPPRIMÉ ICI pour éviter les doublons — le runner UI s’en charge (voir Code.js).

  return summary;
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


/** Trim ciblé de la feuille ERREURS : supprime doublons et purges résolues pour passeports touchés */
function trimErreursIncremental_(seasonSheetId, touchedPassports) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var shE = ss.getSheetByName(SHEETS.ERREURS);
  if (!shE) return { deduped: 0, removedResolved: 0, scanned: 0 };

  var rng = shE.getDataRange();
  var val = rng.getDisplayValues();
  if (!val || val.length <= 1) return { deduped: 0, removedResolved: 0, scanned: 0 };

  var headers = val[0];
  var data = val.slice(1);

  function hIdx(name) { return headers.indexOf(name); }
  var iPass = hIdx('Passeport');
  var iNom = hIdx('Nom');
  var iPre = hIdx('Prénom');
  var iScope = hIdx('Scope');
  var iType = hIdx('Type');
  var iSais = hIdx('Saison');
  var iFrais = hIdx('Frais');
  var iMsg = hIdx('Message');
  var iCtx = hIdx('Contexte');
  var iCreated = hIdx('CreatedAt');

  // Dédup (key simple: Passeport||Type||Saison||Frais||Message)
  function keyOf_(row) { return [row[iPass], row[iType], row[iSais], row[iFrais], row[iMsg]].join('||'); }

  var keepNewest = {}; // key -> {t:timestamp, r:rowIndex}
  var toDelete = [];
  data.forEach(function (row, idx) {
    var k = keyOf_(row);
    var created = row[iCreated] instanceof Date ? row[iCreated].getTime() : new Date(row[iCreated] || new Date()).getTime();
    if (!(k in keepNewest) || keepNewest[k].t < created) {
      keepNewest[k] = { t: created, r: idx };
    }
  });
  data.forEach(function (row, idx) {
    var k = keyOf_(row);
    if (keepNewest[k] && keepNewest[k].r !== idx) toDelete.push(idx);
  });

  // Purge résolues pour les passeports touchés (quelques types fréquents couverts)
  var touchedSet = {};
  (touchedPassports || []).forEach(function (p) { touchedSet[String(p || '').trim()] = true; });

  var removedResolved = 0, scanned = 0;
  data.forEach(function (row, idx) {
    var pass = String(row[iPass] || '').trim();
    if (!pass || !touchedSet[pass]) return;
    // Ne re-supprime pas une ligne déjà marquée pour dédup
    if (toDelete.indexOf(idx) >= 0) return;

    var scope = String(row[iScope] || '').trim();
    var type = String(row[iType] || '').trim();
    var sais = String(row[iSais] || '').trim();
    var psKey = pass + '||' + sais;
    var okStill = true; // par défaut on garde

    scanned++;

    // Exemple de règles de purge «résolu» (à étendre au besoin)
    if (type === 'MembreIntrouvable' || type === 'DoubleInscription') {
      // Considère comme résolu si la clé passeport/saison réapparaît dans FINALS:INSCRIPTIONS
      try {
        var finalsI = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
        var cols = getKeyColsFromParams_(ss);
        var idxByKey = {};
        finalsI.rows.forEach(function (r) { idxByKey[buildKeyFromRow_(r, cols)] = true; });
        // s’il y a au moins UNE ligne de cette saison pour ce passeport -> résolu
        var resolved = finalsI.rows.some(function (r) { return String(r['Passeport #'] || '').trim() === pass && String(r['Saison'] || '').trim() === sais; });
        if (resolved) { okStill = false; }
      } catch (_) { }
    }

    if (!okStill) {
      toDelete.push(idx);
      removedResolved++;
      log_(ss, 'ERR_RESOLVED', JSON.stringify({
        passeport: pass, type: type, saison: sais
      }));
    }
  });

  // Suppressions physiques (du bas vers le haut)
  toDelete.sort(function (a, b) { return b - a; });
  toDelete.forEach(function (idx) {
    shE.deleteRow(idx + 2); // +2 = header + 1-index
  });

  return { deduped: (toDelete.length - removedResolved), removedResolved: removedResolved, scanned: scanned };
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
