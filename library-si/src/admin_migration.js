/**
 * One-shot: normalise INSCRIPTIONS et refabrique ROW_HASH, puis dé-duplique par clé.
 * - Passeport # -> texte + zero-pad(8)
 * - Saison -> injectée si vide avec PARAMS!SEASON_LABEL
 * - ROW_HASH recalculé
 * - Duplicates par clé -> garde 1ère ligne, soft-cancel les autres (+ log dans ANNULATIONS_INSCRIPTIONS)
 */
function oneTimeNormalizeFinalsAndRehash_(seasonSpreadsheetId) {
  var ss = SpreadsheetApp.openById(seasonSpreadsheetId);
  var sh = ss.getSheetByName(SHEETS.INSCRIPTIONS);
  if (!sh || sh.getLastRow() < 2) throw new Error('INSCRIPTIONS vide ou manquant.');

  // Lire PARAMS
  var seasonLabel = (readParam_(ss, 'SEASON_LABEL') || '').trim();
  var keyCols = (readParam_(ss, PARAM_KEYS.KEY_COLS) || 'Passeport #,Saison')
                 .split(',').map(function(s){return s.trim();}).filter(String);

  // Données en display (préserve zéros)
  var rng = sh.getDataRange();
  var disp = rng.getDisplayValues();
  var headers = disp[0];
  var idx = {}; headers.forEach(function(h,i){ idx[h]=i; });

  // Assurer colonnes de contrôle
  function ensureCol_(name) {
    if (idx[name] != null) return;
    headers.push(name);
    sh.getRange(1, headers.length).setValue(name);
    idx[name] = headers.length - 1;
  }
  ensureCol_(CONTROL_COLS.ROW_HASH);
  ensureCol_(CONTROL_COLS.CANCELLED);
  ensureCol_(CONTROL_COLS.EXCLUDE_FROM_EXPORT);
  if (idx['Saison'] == null) {
    headers.push('Saison');
    sh.getRange(1, headers.length).setValue('Saison');
    idx['Saison'] = headers.length - 1;
  }

  // Travail en mémoire
  var data = disp.slice(1); // sans header
  var changed = 0;

  // Normalise chaque ligne
  for (var r=0; r<data.length; r++) {
    var row = data[r];

    // Passeport # -> texte + zero-pad(8)
    if (idx['Passeport #'] != null) {
      row[idx['Passeport #']] = normalizePassportToText8_(row[idx['Passeport #']]).slice(1); // sans l'apostrophe pour l'affichage
    }

    // Saison injectée si absente/vide
    if (idx['Saison'] != null) {
      var v = String(row[idx['Saison']] || '').trim();
      if (v === '' && seasonLabel) { row[idx['Saison']] = seasonLabel; changed++; }
    }

    // Recompute ROW_HASH à partir d'un objet propre
    var obj = {};
    headers.forEach(function(h,i) {
      if (h === CONTROL_COLS.ROW_HASH) return;
      obj[h] = row[i];
    });
    var hash = computeRowHash_(obj);
    row[idx[CONTROL_COLS.ROW_HASH]] = hash;

    // Reset flags si ligne active
    if (row[idx[CONTROL_COLS.CANCELLED]] === '') row[idx[CONTROL_COLS.CANCELLED]] = false;
    if (row[idx[CONTROL_COLS.EXCLUDE_FROM_EXPORT]] === '') row[idx[CONTROL_COLS.EXCLUDE_FROM_EXPORT]] = false;
  }

  // Dé-dup par clé
  var keyToFirst = {};
  var dupCount = 0;
  var annu = getSheetOrCreate_(ss, SHEETS.ANNULATIONS_INSCRIPTIONS, ['Horodatage','Key','RowId']);

  for (var r=0; r<data.length; r++) {
    var row = data[r];
    var key = keyCols.map(function(k){ return String(row[idx[k]]||'').trim(); }).join('||');
    if (!keyToFirst[key]) {
      keyToFirst[key] = r; // garde la première occurrence
    } else {
      // Soft-cancel les duplicates
      row[idx[CONTROL_COLS.CANCELLED]] = true;
      row[idx[CONTROL_COLS.EXCLUDE_FROM_EXPORT]] = true;
      dupCount++;
      annu.appendRow([new Date(), key, r+2]); // +2: header + base-1
    }
  }

  // Écrire en bloc (sans réécrire l'entête)
  sh.getRange(2, 1, data.length, headers.length).setValues(data);

  // Formats
  ensureFinalsColumnFormats_({ sheet: sh, headers: headers });

  appendImportLog_(ss, 'FINALS_MIGRATE_OK', JSON.stringify({
    normalized: data.length,
    changed: changed,
    duplicatesSoftCancelled: dupCount
  }));
}

/**
 * v0.9 — Migration MAPPINGS (sections -> entête unique)
 * - Concatène anciennes sections: ARTICLES, GROUPES, GROUPES_ARTICLES
 * - Construit une seule table à entête unifiée (Type=article|member)
 * - Sauvegarde un backup: MAPPINGS_backup_YYYYMMDD-HHmm
 */
function migrateMappingsToUnified_(seasonSpreadsheetId) {
  var ss = SpreadsheetApp.openById(seasonSpreadsheetId);
  var sh = ss.getSheetByName(SHEETS.MAPPINGS) || ss.insertSheet(SHEETS.MAPPINGS);
  var last = sh.getLastRow();
  var values = last ? sh.getDataRange().getValues() : [];

  // Déjà unifié ?
  if (last && String(values[0][0]).trim() === 'Type') {
    appendImportLog_(ss, 'MAPPINGS_ALREADY_UNIFIED', '');
    return { ok:true, migrated:false };
  }

  // Helpers
  function s(x){ return String(x == null ? '' : x).trim(); }
  function i(x){ var n = parseInt(x,10); return isNaN(n) ? '' : n; }
  function b(x){ var v = String(x||'').toLowerCase(); return (v==='true' || v==='oui' || v==='1'); }

  var out = [];
  function parseSection(label, type){
    for (var r=0; r<values.length; r++){
      if (s(values[r][0]).toUpperCase() === label) {
        var header = (r+1 < values.length) ? values[r+1].map(s) : [];
        var idx = {}; header.forEach(function(h,j){ idx[h]=j; });
        for (var k=r+2; k<values.length; k++){
          var row = values[k] || [];
          var alias = s(row[idx['AliasContains']] || row[idx['Alias']] || '');
          var code  = s(row[idx['Code']]);
          if (!alias && !code) break; // fin de section
          out.push([
            type,                       // Type
            alias,                      // AliasContains
            i(row[idx['Umin']]),        // Umin
            i(row[idx['Umax']]),        // Umax
            s(row[idx['Genre']]),       // Genre
            s(row[idx['GroupeFmt']]),   // GroupeFmt
            s(row[idx['CategorieFmt']]),// CategorieFmt
            b(row[idx['Exclude']]),     // Exclude
            i(row[idx['Priority']]),    // Priority
            code,                       // Code
            s(row[idx['ExclusiveGroup']]) // ExclusiveGroup
          ]);
        }
      }
    }
  }

  // Concatène toutes les sections rencontrées (si présentes)
  parseSection('ARTICLES', 'article');
  parseSection('GROUPES_ARTICLES', 'article');
  parseSection('GROUPES', 'member');

  // Backup
  if (last) {
    var backup = ss.insertSheet('MAPPINGS_backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(),'yyyyMMdd-HHmm'));
    backup.getRange(1,1,last,sh.getLastColumn()).setValues(sh.getDataRange().getValues());
  }

  // Réécriture unifiée
  var header = ['Type','AliasContains','Umin','Umax','Genre','GroupeFmt','CategorieFmt','Exclude','Priority','Code','ExclusiveGroup'];
  sh.clear();
  sh.getRange(1,1,1,header.length).setValues([header]);
  if (out.length) sh.getRange(2,1,out.length,header.length).setValues(out);
  sh.autoResizeColumns(1, header.length);

  appendImportLog_(ss, 'MAPPINGS_MIGRATE_OK', 'rows='+out.length);
  return { ok:true, migrated:true, rows:out.length };
}

