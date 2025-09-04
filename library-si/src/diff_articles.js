/**
 * v0.8 — Diff incrémental ARTICLES (optimisé)
 * - Clé = buildArticleKey_ (Passeport||Saison||Frais-normalisé)
 * - Nouveaux: append + ROW_HASH
 * - Modifiés: update + ROW_HASH + LAST_MODIFIED_AT + log (si statut change)
 * - Annulations (batch):
 *    a) disparition -> CANCELLED/EXCLUDE + LAST_MODIFIED_AT + log ANNULATIONS_ARTICLES
 *    b) statut "annulé" en staging -> idem
 * - ✨ Nouveau: retourne touchedPassports (ensemble des passeports touchés)
 */

// --- helper clé unique ARTICLES (Passeport||Saison||Frais-normalisé)
function buildArticleKey_(row) {
  var p8 = (typeof normalizePassportPlain8_ === 'function')
    ? normalizePassportPlain8_(row['Passeport #'] || row['Passeport'] || '')
    : String(row['Passeport #'] || row['Passeport'] || '').trim();

  var s  = String(row['Saison'] || '').trim();
  var lib = (row['Nom du frais'] || row['Frais'] || row['Produit'] || '');
  lib = String(lib).toLowerCase().replace(/\s+/g,' ').trim();

  return [p8, s, lib].join('||');
}

if (typeof CONTROL_COLS === 'undefined') {
  var CONTROL_COLS = { ROW_HASH:'ROW_HASH', CANCELLED:'CANCELLED', EXCLUDE_FROM_EXPORT:'EXCLUDE_FROM_EXPORT', LAST_MODIFIED_AT:'LAST_MODIFIED_AT' };
}

if (typeof _isCancelledStatus_ !== 'function') {
  function _isCancelledStatus_(val, cancelListCsv) {
    var norm = _norm_(val);
    var list = String(cancelListCsv || '').split(',').map(function(x){return _norm_(x);}).filter(Boolean);
    return list.indexOf(norm) >= 0;
  }
}

function ensureControlCols_(finalsInfo) {
  var sh = finalsInfo.sheet;
  var headers = finalsInfo.headers.slice();
  function ensureCol_(name) {
    var idx = headers.indexOf(name);
    if (idx >= 0) return;
    headers.push(name);
    sh.getRange(1, headers.length).setValue(name);
  }
  ensureCol_(CONTROL_COLS.ROW_HASH);
  ensureCol_(CONTROL_COLS.CANCELLED);
  ensureCol_(CONTROL_COLS.EXCLUDE_FROM_EXPORT);
  ensureCol_(CONTROL_COLS.LAST_MODIFIED_AT);
  finalsInfo.headers = headers;
}
function getKeyColsFromParams_(ss) {
  var csv = readParam_(ss, PARAM_KEYS.KEY_COLS) || 'Passeport #,Saison';
  return csv.split(',').map(function(x){ return x.trim(); }).filter(Boolean);
}
function buildKeyFromRow_(row, keyCols) {
  var parts = keyCols.map(function(k){ return (row[k] == null ? '' : String(row[k])); });
  return parts.join('||');
}
function ensureLogSheetWithHeaders_(ss, sheetName, headers) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    var first = sh.getRange(1,1,1,headers.length).getDisplayValues()[0];
    var ok = headers.every(function(h,i){ return String(first[i]||'') === headers[i]; });
    if (!ok) sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}
function rowIsActive_(row, statusCol, cancelListCsv) {
  var can = String(row[CONTROL_COLS.CANCELLED]||'').toLowerCase()==='true';
  var exc = String(row[CONTROL_COLS.EXCLUDE_FROM_EXPORT]||'').toLowerCase()==='true';
  var st  = row[statusCol];
  return !can && !exc && !_isCancelledStatus_(st, cancelListCsv);
}
function occurrenceLabel_(row) {
  var saison = row['Saison'] || '';
  var frais  = row['Nom du frais'] || row['Frais'] || row['Produit'] || '';
  return (saison + ' — ' + frais).trim();
}
function jsonCompact_(obj){ try { return JSON.stringify(obj); } catch(e){ return '{}'; } }

function diffArticles_(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var staging = readSheetAsObjects_(ss.getId(), SHEETS.STAGING_ARTICLES);
  var finals  = readSheetAsObjects_(ss.getId(), SHEETS.ARTICLES);
  var shFinal = finals.sheet;

  ensureControlCols_(finals);
  var keyCols = getKeyColsFromParams_(ss);
  var statusCol = readParam_(ss, PARAM_KEYS.STATUS_COL_ARTICLES) || 'Statut de l\'inscription';
  var cancelListCsv = readParam_(ss, PARAM_KEYS.STATUS_CANCEL_VALUES) || 'Annulé,Annule,Cancelled';

  // Index finals + occurrences actives par passeport
  var idxFinalByKey = {};
  var activeByPassport = {};
  finals.rows.forEach(function(r, i){
    r.__rownum__ = i + 2;
    r.__key__ = buildArticleKey_(r);
    idxFinalByKey[r.__key__] = r;

    var pass = String(r['Passeport #']||'').trim();
    if (!pass) return;
    if (rowIsActive_(r, statusCol, cancelListCsv)) {
      if (!activeByPassport[pass]) activeByPassport[pass] = [];
      activeByPassport[pass].push(occurrenceLabel_(r));
    }
  });

  var toAppend = [], toUpdate = [];
  var touchedSet = {}; // ✨ nouveau
  function touch_(row){
    var p = String((row && (row['Passeport #'] || row['Passeport'])) || '').trim();
    if (!p) return;
    try { if (typeof normalizePassportPlain8_ === 'function') p = normalizePassportPlain8_(p); } catch(_){ }
    touchedSet[p] = true;
  }

  var HEADERS_ANN = ['Horodatage','Passeport','Nom','Prénom','NomComplet','Saison','Frais','DateAnnulation','A_ENCORE_ACTIF','ACTIFS_RESTANTS'];
  var annRows = [];

  function rowHash_(r){ return computeRowHash_(r); }

  // --- Parcours staging
  staging.rows.forEach(function(sRow){
    var key = buildArticleKey_(sRow);
    var fRow = idxFinalByKey[key];
    var sCancelled = _isCancelledStatus_(sRow[statusCol], cancelListCsv);

    if (!fRow) {
      // NEW
      var newRow = {};
      finals.headers.forEach(function(h){ newRow[h] = sRow[h] == null ? '' : sRow[h]; });
      newRow[CONTROL_COLS.CANCELLED] = !!sCancelled;
      newRow[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = !!sCancelled;
      newRow[CONTROL_COLS.ROW_HASH] = rowHash_(newRow);
      toAppend.push(newRow);
      touch_(newRow);

      if (sCancelled) {
        var passN = String(newRow['Passeport #']||'').trim();
        var actifsN = (activeByPassport[passN]||[]).slice(0);
        annRows.push([
          new Date(),
          normalizePassportToText8_(passN),
          newRow['Nom'] || '',
          newRow['Prénom'] || newRow['Prenom'] || '',
          ((newRow['Prénom']||newRow['Prenom']||'') + ' ' + (newRow['Nom']||'')).trim(),
          newRow['Saison'] || '',
          newRow['Nom du frais'] || newRow['Frais'] || newRow['Produit'] || '',
          readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ? (newRow[ readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ] || '') : '',
          actifsN.length > 0,
          actifsN.join(' ; ')
        ]);
      }
      return;
    }

    // EXISTANT → merge
    var oldHash = fRow[CONTROL_COLS.ROW_HASH] || '';
    var merged = {};
    finals.headers.forEach(function(h){ merged[h] = (sRow[h] == null ? fRow[h] : sRow[h]); });

    var newCancelled = _isCancelledStatus_(merged[statusCol], cancelListCsv);
    merged[CONTROL_COLS.CANCELLED] = !!newCancelled;
    merged[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = !!newCancelled;

    var newHash = computeRowHash_(merged);
    if (newHash !== oldHash) {
      merged[CONTROL_COLS.ROW_HASH] = newHash;
      merged[CONTROL_COLS.LAST_MODIFIED_AT] = new Date();
      toUpdate.push({ rownum: fRow.__rownum__, data: merged });
      touch_(merged);
    } else if (newCancelled !== (String(fRow[CONTROL_COLS.CANCELLED]||'').toLowerCase()==='true')) {
      // Changement de statut => annulation
      toUpdate.push({ rownum: fRow.__rownum__, data: merged });
      touch_(merged);

      var passU = String(merged['Passeport #']||'').trim();
      var list = (activeByPassport[passU]||[]).slice(0);
      var occ = occurrenceLabel_(merged);
      var idx = list.indexOf(occ);
      if (idx >= 0) list.splice(idx,1);

      annRows.push([
        new Date(),
        normalizePassportToText8_(passU),
        merged['Nom'] || '',
        merged['Prénom'] || merged['Prenom'] || '',
        ((merged['Prénom']||merged['Prenom']||'') + ' ' + (merged['Nom']||'')).trim(),
        merged['Saison'] || '',
        merged['Nom du frais'] || merged['Frais'] || merged['Produit'] || '',
        readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ? (merged[ readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ] || '') : '',
        list.length > 0,
        list.join(' ; ')
      ]);
    }
  });

  // --- Disparitions = annulations (⚠️ ajoute LAST_MODIFIED_AT + touch_)
  var indexStagingKeys = {};
  staging.rows.forEach(function(s){ indexStagingKeys[buildArticleKey_(s)] = true; });
  finals.rows.forEach(function(fRow){
    if (!indexStagingKeys[fRow.__key__]) {
      fRow[CONTROL_COLS.CANCELLED] = true;
      fRow[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = true;
      fRow[CONTROL_COLS.LAST_MODIFIED_AT] = new Date(); // ✨
      toUpdate.push({ rownum: fRow.__rownum__, data: fRow });
      touch_(fRow);

      var passD = String(fRow['Passeport #']||'').trim();
      var listD = (activeByPassport[passD]||[]).slice(0);
      var occD = occurrenceLabel_(fRow);
      var ix = listD.indexOf(occD);
      if (ix >= 0) listD.splice(ix,1);

      annRows.push([
        new Date(),
        normalizePassportToText8_(passD),
        fRow['Nom'] || '',
        fRow['Prénom'] || fRow['Prenom'] || '',
        ((fRow['Prénom']||fRow['Prenom']||'') + ' ' + (fRow['Nom']||'')).trim(),
        fRow['Saison'] || '',
        fRow['Nom du frais'] || fRow['Frais'] || fRow['Produit'] || '',
        readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ? (fRow[ readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ] || '') : '',
        listD.length > 0,
        listD.join(' ; ')
      ]);
    }
  });

  // --- Écritures finals
  if (toAppend.length) {
    var startRow = shFinal.getLastRow() + 1;
    shFinal.insertRowsAfter(shFinal.getLastRow(), toAppend.length);
    var headers = finals.headers;
    var values = toAppend.map(function (r) { return headers.map(function (h) { return r[h]; }); });
    shFinal.getRange(startRow, 1, toAppend.length, headers.length).setValues(values);
  }
  if (toUpdate.length) {
    var headersU = finals.headers;
    toUpdate.forEach(function (up) {
      var rowArr = headersU.map(function (h) { return up.data[h]; });
      shFinal.getRange(up.rownum, 1, 1, headersU.length).setValues([rowArr]);
    });
  }

  // --- Logs batch (annulations)
  if (annRows.length) {
    var shAnn = ensureLogSheetWithHeaders_(ss, SHEETS.ANNULATIONS_ARTICLES, HEADERS_ANN);
    var startA = shAnn.getLastRow() + 1;
    shAnn.insertRowsAfter(shAnn.getLastRow(), annRows.length);
    shAnn.getRange(startA, 1, annRows.length, HEADERS_ANN.length).setValues(annRows);
  }

  var touchedPassports = Object.keys(touchedSet);
  return {
    added: toAppend.length,
    updated: toUpdate.length,
    annuls: annRows.length,
    touchedPassports: touchedPassports
  };
}
