/**
 * v0.7 — Diff incrémental INSCRIPTIONS (optimisé)
 * - Nouveaux: append + ROW_HASH + OUTBOX (INSCRIPTION_NEW)
 * - Modifiés: update ciblé + ROW_HASH + LAST_MODIFIED_AT + log MODIFS_INSCRIPTIONS (batch)
 * - Annulations (batch):
 *    a) disparition -> CANCELLED/EXCLUDE + log ANNULATIONS_INSCRIPTIONS
 *    b) statut "annulé" en staging -> idem
 */

/*** Fallbacks sûrs ***/



if (typeof CONTROL_COLS === 'undefined') {
  var CONTROL_COLS = { ROW_HASH: 'ROW_HASH', CANCELLED: 'CANCELLED', EXCLUDE_FROM_EXPORT: 'EXCLUDE_FROM_EXPORT', LAST_MODIFIED_AT: 'LAST_MODIFIED_AT' };
}

if (typeof _isCancelledStatus_ !== 'function') {
  function _isCancelledStatus_(val, cancelListCsv) {
    var norm = _norm_(val);
    var list = String(cancelListCsv || '').split(',').map(function (x) { return _norm_(x); }).filter(Boolean);
    return list.indexOf(norm) >= 0;
  }
}




/*** Helpers locaux ***/
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
  return csv.split(',').map(function (x) { return x.trim(); }).filter(Boolean);
}
function buildKeyFromRow_(row, keyCols) {
  var parts = keyCols.map(function (k) { return (row[k] == null ? '' : String(row[k])); });
  return parts.join('||');
}
function ensureLogSheetWithHeaders_(ss, sheetName, headers) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sh;
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    var first = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    var ok = headers.every(function (h, i) { return String(first[i] || '') === headers[i]; });
    if (!ok) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}
function rowIsActive_(row, statusCol, cancelListCsv) {
  var can = String(row[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true';
  var exc = String(row[CONTROL_COLS.EXCLUDE_FROM_EXPORT] || '').toLowerCase() === 'true';
  var st = row[statusCol];
  return !can && !exc && !_isCancelledStatus_(st, cancelListCsv);
}
function occurrenceLabel_(row) {
  var saison = row['Saison'] || '';
  var frais = row['Nom du frais'] || row['Frais'] || row['Produit'] || '';
  return (saison + ' — ' + frais).trim();
}

/** Diff principal */
function diffInscriptions_(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var staging = readSheetAsObjects_(ss.getId(), SHEETS.STAGING_INSCRIPTIONS);
  var finals = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var shFinal = finals.sheet;

  ensureControlCols_(finals);

  var keyCols = getKeyColsFromParams_(ss);
  var statusCol = readParam_(ss, PARAM_KEYS.STATUS_COL_INSCRIPTIONS) || 'Statut de l\'inscription';
  var cancelListCsv = readParam_(ss, PARAM_KEYS.STATUS_CANCEL_VALUES) || 'Annulé,Annule,Cancelled';

  // index finals par clé et par passeport (actifs)
  var idxFinalByKey = {};
  var activeByPassport = {}; // { '00001234': [ '2025-Ete — U10 F ...', ... ] }
  finals.rows.forEach(function (r, i) {
    r.__rownum__ = i + 2;
    r.__key__ = buildKeyFromRow_(r, keyCols);
    idxFinalByKey[r.__key__] = r;

    var pass = String(r['Passeport #'] || '').trim();
    if (!pass) return;
    if (rowIsActive_(r, statusCol, cancelListCsv)) {
      if (!activeByPassport[pass]) activeByPassport[pass] = [];
      activeByPassport[pass].push(occurrenceLabel_(r));
    }
  });

  var toAppend = [];
  var toUpdate = [];
  var outboxRows = [];

  // Logs batch
  var HEADERS_ANN = ['Horodatage', 'Passeport', 'Nom', 'Prénom', 'NomComplet', 'Saison', 'Frais', 'DateAnnulation', 'A_ENCORE_ACTIF', 'ACTIFS_RESTANTS'];
  var annRows = [];
  var HEADERS_MOD = ['Horodatage', 'Passeport', 'Nom', 'Prénom', 'NomComplet', 'Saison', 'ChangedFieldsJSON'];
  var modRows = [];

  function rowHash_(r) { return computeRowHash_(r); }

  // --- 1) Parcours du staging
  staging.rows.forEach(function (sRow) {
    var key = buildKeyFromRow_(sRow, keyCols);
    var fRow = idxFinalByKey[key];
    var sCancelled = _isCancelledStatus_(sRow[statusCol], cancelListCsv);

    if (!fRow) {
      // NEW
      var newRow = {};
      finals.headers.forEach(function (h) { newRow[h] = sRow[h] == null ? '' : sRow[h]; });
      newRow[CONTROL_COLS.CANCELLED] = !!sCancelled;
      newRow[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = !!sCancelled;
      newRow[CONTROL_COLS.ROW_HASH] = rowHash_(newRow);
      toAppend.push(newRow);

      if (!sCancelled) {
        // Outbox (INSCRIPTION_NEW)
        var keyHash = Utilities.base64EncodeWebSafe(key);
        // --- enrichit l’OUTBOX (lisible dans la feuille) ---
        var passRaw = newRow['Passeport #'] || '';
        var passText8 = (typeof normalizePassportToText8_ === 'function')
          ? normalizePassportToText8_(passRaw)
          : String(passRaw); // fallback

        var prenom = newRow['Prénom'] || newRow['Prenom'] || '';
        var nom = newRow['Nom'] || '';
        var nomComplet = ((prenom || '') + ' ' + (nom || '')).trim();
        var saison = newRow['Saison'] || '';
        var frais = newRow['Nom du frais'] || newRow['Frais'] || newRow['Produit'] || '';

        // candidats emails (pour lecture; le worker résoudra To/Cc final)
        var colsCsv = readParam_(ss, 'TO_FIELDS_INSCRIPTIONS') || readParam_(ss, 'TO_FIELD_INSCRIPTIONS') ||
          'Courriel,Parent 1 - Courriel,Parent 2 - Courriel';
        var emailsCandidates = '';
        if (typeof collectEmailsFromRow_ === 'function') {
          emailsCandidates = collectEmailsFromRow_(newRow, colsCsv) || '';
        } else {
          // petit fallback : concatène les 3 colonnes usuelles si présentes
          var cand = [];
          ['Courriel', 'Parent 1 - Courriel', 'Parent 2 - Courriel'].forEach(function (c) {
            var v = newRow[c]; if (v) cand.push(String(v).trim());
          });
          emailsCandidates = cand.filter(Boolean).join(',');
        }

        // construit la ligne alignée sur l’entête enrichie
        var hdr = getMailOutboxHeaders_();
        var out = new Array(hdr.length).fill('');

        // colonnes “système”
        out[0] = 'INSCRIPTION_NEW';  // Type
        out[1] = '';                  // To (résolu par le worker)
        out[2] = '';                  // Cc
        out[3] = '';                  // Sujet
        out[4] = '';                  // Corps
        out[5] = '';                  // Attachments
        out[6] = keyHash;             // KeyHash
        out[7] = 'pending';           // Status
        out[8] = new Date();          // CreatedAt
        out[9] = '';                  // SentAt
        out[10] = '';                  // Error

        // nouvelles colonnes d’info (debug/tri)
        out[11] = passText8;           // Passeport8
        out[12] = nom;                 // Nom
        out[13] = prenom;              // Prénom
        out[14] = nomComplet;          // NomComplet
        out[15] = saison;              // Saison
        out[16] = frais;               // Frais
        out[17] = emailsCandidates;    // EmailsCandidates

        outboxRows.push(out);

      } else {
        // arrive déjà annulé → log (batch)
        var pass = String(newRow['Passeport #'] || '').trim();
        var actifs = (activeByPassport[pass] || []).slice(0); // autres lignes actives (s'il y en a)
        annRows.push([
          new Date(),
          normalizePassportToText8_(pass),
          newRow['Nom'] || '',
          newRow['Prénom'] || newRow['Prenom'] || '',
          ((newRow['Prénom'] || newRow['Prenom'] || '') + ' ' + (newRow['Nom'] || '')).trim(),
          newRow['Saison'] || '',
          newRow['Nom du frais'] || newRow['Frais'] || newRow['Produit'] || '',
          readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ? (newRow[readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL)] || '') : '',
          actifs.length > 0,
          actifs.join(' ; ')
        ]);
      }
      return;
    }

    // EXISTANT → merge
    var oldHash = fRow[CONTROL_COLS.ROW_HASH] || '';
    var merged = {};
    finals.headers.forEach(function (h) { merged[h] = (sRow[h] == null ? fRow[h] : sRow[h]); });

    var newCancelled = _isCancelledStatus_(merged[statusCol], cancelListCsv);
    merged[CONTROL_COLS.CANCELLED] = !!newCancelled;
    merged[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = !!newCancelled;

    var newHash = computeRowHash_(merged);
    if (newHash !== oldHash) {
      merged[CONTROL_COLS.ROW_HASH] = newHash;
      merged[CONTROL_COLS.LAST_MODIFIED_AT] = new Date();

      // Log modifications (batch)
      var changed = {};
      finals.headers.forEach(function (h) {
        if (String(fRow[h] || '') !== String(merged[h] || '')) changed[h] = { old: fRow[h] || '', nu: merged[h] || '' };
      });
      modRows.push([
        new Date(),
        normalizePassportToText8_(merged['Passeport #'] || ''),
        merged['Nom'] || '',
        merged['Prénom'] || merged['Prenom'] || '',
        ((merged['Prénom'] || merged['Prenom'] || '') + ' ' + (merged['Nom'] || '')).trim(),
        merged['Saison'] || '',
        jsonCompact_(changed)
      ]);

      toUpdate.push({ rownum: fRow.__rownum__, data: merged });

      // si ce merge entraîne une annulation, on la loguera plus bas (via test du statut)
      if (newCancelled && !rowIsActive_(fRow, statusCol, cancelListCsv)) {
        // already cancelled before, nothing new
      }
    } else if (newCancelled !== (String(fRow[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true')) {
      // Changement de statut uniquement → log annulation (batch)
      toUpdate.push({ rownum: fRow.__rownum__, data: merged });

      var passU = String(merged['Passeport #'] || '').trim();
      var list = (activeByPassport[passU] || []).slice(0);
      // retirer l'occurrence en cours (elle est en train de passer à annulée)
      var occ = occurrenceLabel_(merged);
      var idx = list.indexOf(occ);
      if (idx >= 0) list.splice(idx, 1);

      annRows.push([
        new Date(),
        normalizePassportToText8_(passU),
        merged['Nom'] || '',
        merged['Prénom'] || merged['Prenom'] || '',
        ((merged['Prénom'] || merged['Prenom'] || '') + ' ' + (merged['Nom'] || '')).trim(),
        merged['Saison'] || '',
        merged['Nom du frais'] || merged['Frais'] || merged['Produit'] || '',
        readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ? (merged[readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL)] || '') : '',
        list.length > 0,
        list.join(' ; ')
      ]);
    }
  });

  // --- 2) Disparitions = annulations
  var indexStagingKeys = {};
  staging.rows.forEach(function (s) { indexStagingKeys[buildKeyFromRow_(s, keyCols)] = true; });

  finals.rows.forEach(function (fRow) {
    if (!indexStagingKeys[fRow.__key__]) {
      fRow[CONTROL_COLS.CANCELLED] = true;
      fRow[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = true;
      toUpdate.push({ rownum: fRow.__rownum__, data: fRow });

      var passD = String(fRow['Passeport #'] || '').trim();
      var listD = (activeByPassport[passD] || []).slice(0);
      var occD = occurrenceLabel_(fRow);
      var ix = listD.indexOf(occD);
      if (ix >= 0) listD.splice(ix, 1);

      annRows.push([
        new Date(),
        normalizePassportToText8_(passD),
        fRow['Nom'] || '',
        fRow['Prénom'] || fRow['Prenom'] || '',
        ((fRow['Prénom'] || fRow['Prenom'] || '') + ' ' + (fRow['Nom'] || '')).trim(),
        fRow['Saison'] || '',
        fRow['Nom du frais'] || fRow['Frais'] || fRow['Produit'] || '',
        readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ? (fRow[readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL)] || '') : '',
        listD.length > 0,
        listD.join(' ; ')
      ]);
    }
  });

  // --- 3) Écritures finals
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

  // --- 4) Logs batch (annulations / modifications)
  if (annRows.length) {
    var shAnn = ensureLogSheetWithHeaders_(ss, SHEETS.ANNULATIONS_INSCRIPTIONS, HEADERS_ANN);
    var startA = shAnn.getLastRow() + 1;
    shAnn.insertRowsAfter(shAnn.getLastRow(), annRows.length);
    shAnn.getRange(startA, 1, annRows.length, HEADERS_ANN.length).setValues(annRows);
  }
  if (modRows.length) {
    var shMod = ensureLogSheetWithHeaders_(ss, SHEETS.MODIFS_INSCRIPTIONS, HEADERS_MOD);
    var startM = shMod.getLastRow() + 1;
    shMod.insertRowsAfter(shMod.getLastRow(), modRows.length);
    shMod.getRange(startM, 1, modRows.length, HEADERS_MOD.length).setValues(modRows);
  }

  // --- 5) Outbox des nouveaux
  if (outboxRows.length) enqueueOutboxRows_(ss.getId(), outboxRows);

  return { added: toAppend.length, updated: toUpdate.length, outbox: outboxRows.length, annuls: annRows.length, mods: modRows.length };
}
