/**
 * v0.8.1 — Diff incrémental INSCRIPTIONS (optimisé + filtre entraîneurs)
 * - Nouveaux: append + ROW_HASH + OUTBOX (INSCRIPTION_NEW)  ❗(hors entraîneurs)
 * - Modifiés: update ciblé + ROW_HASH + LAST_MODIFIED_AT + log MODIFS_INSCRIPTIONS (batch)
 * - Annulations (batch):
 *    a) disparition -> CANCELLED/EXCLUDE + LAST_MODIFIED_AT + log ANNULATIONS_INSCRIPTIONS
 *    b) statut "annulé" en staging -> idem
 * - ✨ Nouveau (v0.8): retourne touchedPassports (ensemble des passeports touchés)
 * - ✨ v0.8.1: entraîneurs exclus du FINAL; lignes existantes d'entraîneurs forcées en annulées/exclues
 */

/*** Fallbacks sûrs ***/
if (typeof CONTROL_COLS === 'undefined') {
  var CONTROL_COLS = {
    ROW_HASH: 'ROW_HASH',
    CANCELLED: 'CANCELLED',
    EXCLUDE_FROM_EXPORT: 'EXCLUDE_FROM_EXPORT',
    LAST_MODIFIED_AT: 'LAST_MODIFIED_AT'
  };
}
if (typeof _norm_ !== 'function') {
  function _norm_(s){
    s = String(s == null ? '' : s).trim().toLowerCase();
    try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g,''); } catch(_){}
    return s;
  }
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

/*** Détection entraîneur (lib -> fallback) ***/
function _isCoachFromFeeName_(ss, feeRaw){
  var fee = String(feeRaw || '');
  // lib si dispo
  try { if (typeof isCoachFeeByName_ === 'function') return !!isCoachFeeByName_(ss, fee); } catch(_){}
  // param CSV + fallback lexical
  var v = _norm_(fee);
  if (!v) return false;
  var csv = (typeof readParam_==='function') ? (readParam_(ss,'RETRO_COACH_FEES_CSV')||'') : '';
  var toks = csv.split(',').map(_norm_).filter(Boolean);
  if (toks.length){
    if (toks.indexOf(v) >= 0) return true;
    for (var i=0;i<toks.length;i++) if (v.indexOf(toks[i]) >= 0) return true;
  }
  return /(entraineur|entra[îi]neur|coach)/i.test(fee);
}
function _isCoachMemberSafe_(ss, row){
  try { if (typeof isCoachMember_ === 'function') return !!isCoachMember_(ss, row); } catch(_){}
  var fee = (row && (row['Nom du frais']||row['Frais']||row['Produit'])) || '';
  return _isCoachFromFeeName_(ss, fee);
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
  var touchedSet = {}; // ✨ nouveau
  function touch_(row){
    var p = String((row && (row['Passeport #'] || row['Passeport'])) || '').trim();
    if (!p) return;
    try { if (typeof normalizePassportPlain8_ === 'function') p = normalizePassportPlain8_(p); } catch(_){ }
    touchedSet[p] = true;
  }

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

    // *** Détermine si COACH ***
    var isCoach = _isCoachMemberSafe_(ss, sRow);

    if (!fRow) {
      // NEW
      if (isCoach) {
        // → on ignore totalement les coachs côté FINAL (pas d'append, pas d'outbox, pas de log)
        return;
      }

      var newRow = {};
      finals.headers.forEach(function (h) { newRow[h] = sRow[h] == null ? '' : sRow[h]; });
      newRow[CONTROL_COLS.CANCELLED] = !!sCancelled;
      newRow[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = !!sCancelled;
      newRow[CONTROL_COLS.ROW_HASH] = rowHash_(newRow);
      toAppend.push(newRow);
      touch_(newRow);

      if (!sCancelled) {
        // Outbox (INSCRIPTION_NEW) — on compte seulement ici; l'enqueue réel se fait ailleurs si souhaité
        var keyHash = Utilities.base64EncodeWebSafe(key);

        var passRaw = newRow['Passeport #'] || '';
        var passText8 = (typeof normalizePassportToText8_ === 'function')
          ? normalizePassportToText8_(passRaw)
          : String(passRaw);

        var prenom = newRow['Prénom'] || newRow['Prenom'] || '';
        var nom = newRow['Nom'] || '';
        var nomComplet = ((prenom || '') + ' ' + (nom || '')).trim();
        var saison = newRow['Saison'] || '';
        var frais = newRow['Nom du frais'] || newRow['Frais'] || newRow['Produit'] || '';

        var colsCsv = readParam_(ss, 'TO_FIELDS_INSCRIPTIONS') || readParam_(ss, 'TO_FIELD_INSCRIPTIONS') ||
          'Courriel,Parent 1 - Courriel,Parent 2 - Courriel';
        var emailsCandidates = '';
        if (typeof collectEmailsFromRow_ === 'function') {
          emailsCandidates = collectEmailsFromRow_(newRow, colsCsv) || '';
        } else {
          var cand = [];
          ['Courriel', 'Parent 1 - Courriel', 'Parent 2 - Courriel'].forEach(function (c) {
            var v = newRow[c]; if (v) cand.push(String(v).trim());
          });
          emailsCandidates = cand.filter(Boolean).join(',');
        }

        var hdr = getMailOutboxHeaders_();
        var out = new Array(hdr.length).fill('');
        out[0]  = 'INSCRIPTION_NEW';
        out[6]  = keyHash;
        out[7]  = 'pending';
        out[8]  = new Date();
        // colonnes facultatives de lisibilité (si présentes côté projet):
        // on met ce qu'on peut sans casser (les index exacts peuvent différer selon upgrade local)
        try {
          var idxHdr = {};
          hdr.forEach(function(h,i){ idxHdr[h]=i; });
          function setIf(colName, val){ var i = idxHdr[colName]; if (typeof i==='number' && i>=0) out[i]=val; }
          setIf('Passeport', passText8);
          setIf('NomComplet', nomComplet);
          setIf('Frais', frais);
        } catch(_){}

        outboxRows.push(out);
      } else {
        // arrive déjà annulé → log (batch)
        var pass = String(newRow['Passeport #'] || '').trim();
        var actifs = (activeByPassport[pass] || []).slice(0); // autres lignes actives
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

    // EXISTANT
    if (isCoach) {
      // Politique v0.8.1 : un coach ne doit PAS rester dans FINAL → force CANCEL/EXCLUDE si pas déjà fait
      var alreadyCancelled = String(fRow[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true';
      var alreadyExcluded  = String(fRow[CONTROL_COLS.EXCLUDE_FROM_EXPORT] || '').toLowerCase() === 'true';
      if (!alreadyCancelled || !alreadyExcluded) {
        fRow[CONTROL_COLS.CANCELLED] = true;
        fRow[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = true;
        fRow[CONTROL_COLS.LAST_MODIFIED_AT] = new Date();
        toUpdate.push({ rownum: fRow.__rownum__, data: fRow });
        touch_(fRow);

        var passC = String(fRow['Passeport #'] || '').trim();
        var listC = (activeByPassport[passC] || []).slice(0);
        var occC = occurrenceLabel_(fRow);
        var ixC = listC.indexOf(occC);
        if (ixC >= 0) listC.splice(ixC, 1);

        annRows.push([
          new Date(),
          normalizePassportToText8_(passC),
          fRow['Nom'] || '',
          fRow['Prénom'] || fRow['Prenom'] || '',
          ((fRow['Prénom'] || fRow['Prenom'] || '') + ' ' + (fRow['Nom'] || '')).trim(),
          fRow['Saison'] || '',
          fRow['Nom du frais'] || fRow['Frais'] || fRow['Produit'] || '',
          readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL) ? (fRow[readParam_(ss, PARAM_KEYS.STATUS_CANCEL_DATE_COL)] || '') : '',
          listC.length > 0,
          listC.join(' ; ')
        ]);
      }
      return; // rien d'autre à faire pour un coach
    }

    // EXISTANT non-coach → merge normal
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
      touch_(merged);

    } else if (newCancelled !== (String(fRow[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true')) {
      // Changement de statut uniquement → log annulation (batch)
      toUpdate.push({ rownum: fRow.__rownum__, data: merged });
      touch_(merged);

      var passU = String(merged['Passeport #'] || '').trim();
      var list = (activeByPassport[passU] || []).slice(0);
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

  // --- 2) Disparitions = annulations (⚠️ ajoute LAST_MODIFIED_AT + touch_)
  var indexStagingKeys = {};
  staging.rows.forEach(function (s) { indexStagingKeys[buildKeyFromRow_(s, keyCols)] = true; });

  finals.rows.forEach(function (fRow) {
    if (!indexStagingKeys[fRow.__key__]) {
      fRow[CONTROL_COLS.CANCELLED] = true;
      fRow[CONTROL_COLS.EXCLUDE_FROM_EXPORT] = true;
      fRow[CONTROL_COLS.LAST_MODIFIED_AT] = new Date(); // ✨ pour qu'un export incrémente bien
      toUpdate.push({ rownum: fRow.__rownum__, data: fRow });
      touch_(fRow);

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

  // --- 5) Outbox des nouveaux (toujours désactivé ici; le worker s'en charge ailleurs)
  // if (outboxRows.length) enqueueOutboxRows_(ss.getId(), outboxRows);

  var touchedPassports = Object.keys(touchedSet);
  return {
    added: toAppend.length,
    updated: toUpdate.length,
    outbox: outboxRows.length,
    annuls: annRows.length,
    mods: modRows.length,
    touchedPassports: touchedPassports
  };
}
