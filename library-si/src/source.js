/*************************************************
 *  Guard: PARAMS & SHEETS keys
 *************************************************/
(function bootstrapAugmentations_() {
  try {
    if (typeof SHEETS === 'undefined') { this.SHEETS = {}; }
    if (!SHEETS.ACHATS_LEDGER) SHEETS.ACHATS_LEDGER = 'ACHATS_LEDGER';
    if (!SHEETS.JOUEURS) SHEETS.JOUEURS = 'JOUEURS';

    if (typeof PARAM_KEYS === 'undefined') { this.PARAM_KEYS = {}; }
    if (!PARAM_KEYS.LEDGER_ENABLED) PARAM_KEYS.LEDGER_ENABLED = 'LEDGER_ENABLED';
    if (!PARAM_KEYS.JOUEURS_ENABLED) PARAM_KEYS.JOUEURS_ENABLED = 'JOUEURS_ENABLED';
    if (!PARAM_KEYS.RETRO_MEMBRES_READ_SRC) PARAM_KEYS.RETRO_MEMBRES_READ_SRC = 'RETRO_MEMBRES_READ_SOURCE';
    if (!PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL = 'RETRO_PHOTO_INCLUDE_COL';
  } catch (e) { }
})();

/*************************************************
 *  Hook d’orchestration post-import (FULL/INCR)
 *************************************************/
/**
 * A appeler à la fin d’un run (FULL/INCR) après les diffs et les règles.
 * @param {string} seasonSheetId
 * @param {Array<string>|Set<string>} touchedPassports  // [] en FULL
 * @param {{isFull?:boolean, isDryRun?:boolean}} opts
 */

function runPostImportAugmentations_(seasonSheetId, touchedPassports, opts) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var isFull = !!(opts && opts.isFull);
  var isDryRun = !!(opts && opts.isDryRun);

  var ledgerOn = String(readParam_(ss, 'LEDGER_ENABLED') || 'FALSE').toUpperCase() === 'TRUE';
  var joueursOn = String(readParam_(ss, 'JOUEURS_ENABLED') || 'FALSE').toUpperCase() === 'TRUE';

  if (isDryRun) {
    try { appendImportLog_(ss, 'AUG_SKIP', 'DRY_RUN=TRUE'); } catch (e) { }
    return;
  }

  var touchedSet = (function _toSet_(arr) {
    if (!arr) return new Set();
    if (arr instanceof Set) return arr;
    if (Array.isArray(arr)) return new Set(arr.map(function (x) { return String(x).trim(); }));
    var out = new Set(); try { Object.keys(arr).forEach(function (k) { if (arr[k]) out.add(String(k).trim()); }); } catch (e) { }
    return out;
  })(touchedPassports);

  try {
    if (ledgerOn) {
      if (isFull) {
        buildAchatsLedger_(ss);
        appendImportLog_(ss, 'LEDGER_BUILD_OK', JSON.stringify({ mode: 'FULL' }));
      } else {
        updateAchatsLedgerForPassports_(ss, touchedSet);
        appendImportLog_(ss, 'LEDGER_INCR_OK', JSON.stringify({ mode: 'INCR', count: touchedSet.size || 0 }));
      }
    } else {
      appendImportLog_(ss, 'LEDGER_DISABLED', '{}');
    }
  } catch (e) {
    appendImportLog_(ss, 'LEDGER_FAIL', String(e));
    throw e;
  }

  try {
    if (joueursOn) {
      if (isFull) {
        var res = buildJoueursIndex_(ss);                      // <- récupère {header, rows}
        writeObjectsToSheet_(ss, SHEETS.JOUEURS, res.rows, res.header);  // <- ÉCRITURE FULL
        refreshPhotoStrInJoueurs_(ss);  // ← ici, après l’écriture FULL
        appendImportLog_(ss, 'JOUEURS_BUILD_OK', JSON.stringify({ mode: 'FULL', count: (res.rows || []).length }));
      } else {
        updateJoueursForPassports_(ss, touchedSet);
        refreshJoueursPhotoStr_(ss);
        appendImportLog_(ss, 'JOUEURS_INCR_OK', JSON.stringify({ mode: 'INCR', count: touchedSet.size || 0 }));
      }

    } else {
      appendImportLog_(ss, 'JOUEURS_DISABLED', '{}');
    }
  } catch (e) {
    appendImportLog_(ss, 'JOUEURS_FAIL', String(e));
    throw e;
  }
}


function _augLog_(code, msg) {
  try {
    if (typeof appendImportLog_ === 'function') appendImportLog_({ type: code, details: msg });
    else if (typeof addLogLine_ === 'function') addLogLine_(code, msg);
    else console && console.log && console.log(code, msg);
  } catch (e) { }
}

function buildRetroMembresRowsSelected_(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var src = String(readParam_(ss, PARAM_KEYS.RETRO_MEMBRES_READ_SRC) || 'LEGACY').toUpperCase();
  if (src === 'JOUEURS' && typeof buildRetroMembresRowsFromJoueurs_ === 'function') {
    if (typeof appendImportLog_ === 'function')
      appendImportLog_({ type: 'RETRO_MEMBRES_SOURCE', details: { source: 'JOUEURS' } });
    return buildRetroMembresRowsFromJoueurs_(seasonSheetId);
  }
  if (typeof appendImportLog_ === 'function')
    appendImportLog_({ type: 'RETRO_MEMBRES_SOURCE', details: { source: 'LEGACY' } });
  return buildRetroMembresRows(seasonSheetId);
}

function _pickSheetName_(ss, finalName, stagingName) {
  try { return ss.getSheetByName(finalName) ? finalName : stagingName; }
  catch (e) { return stagingName; }
}


/** =========================
 *  ACHATS_LEDGER — FULL
 *  ========================= */
function buildAchatsLedger_(ss) {
  var saison = readParam_(ss, 'SEASON_LABEL') || '';

  // --- ignore list robuste (CSV vide => défaut; Set/Array/Object/CSV OK)
  var DEFAULT_IGNORE = 'senior,u-se,adulte,ligue';
  var ignoreCsvRaw = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV);
  var ignoreCsv = (ignoreCsvRaw && String(ignoreCsvRaw).trim()) ? ignoreCsvRaw : DEFAULT_IGNORE;

  function toIgnoreArr_(x) {
    if (!x) return [];
    if (Array.isArray(x)) return x;
    if (typeof x === 'string') return x.split(',');
    if (x && typeof x.forEach === 'function' && x.add) { var arr = []; x.forEach(function (v) { arr.push(v); }); return arr; }
    if (typeof x === 'object') return Object.keys(x);
    return [];
  }
  var ignoreList = toIgnoreArr_(_compileCsvToSet_(ignoreCsv)).map(function (s) {
    return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  });

  function matchIgnore_(name) {
    var s = String(name || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    for (var i = 0; i < ignoreList.length; i++) {
      var k = ignoreList[i]; if (!k) continue;
      if (s.indexOf(k) !== -1) return true;
    }
    return false;
  }

  // --- I/O
  var shInscName = _pickSheetName_(ss, SHEETS.INSCRIPTIONS_FINAL, SHEETS.INSCRIPTIONS);
  var shArtName  = _pickSheetName_(ss, SHEETS.ARTICLES_FINAL, SHEETS.ARTICLES);
  var insc = readSheetAsObjects_(ss.getId(), shInscName);
  var art  = readSheetAsObjects_(ss.getId(), shArtName);

  appendImportLog_(ss, 'LEDGER_INPUT_SHEETS', JSON.stringify({ insc: shInscName, art: shArtName }));
  appendImportLog_(ss, 'LEDGER_INPUT_ROWS',   JSON.stringify({ insc: (insc.rows || []).length, art: (art.rows || []).length }));

  // --- helpers
  var normP = (typeof normalizePassport8_ === 'function')
    ? normalizePassport8_
    : function (p) { return (p == null || p === '') ? '' : String(p).replace(/\D/g, '').padStart(8, '0'); };

  function slugify_(s) {
    return String(s || '').toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  }
  function _nl_(s){ try { return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase(); } catch(e){ return String(s||'').toLowerCase(); } }
  function _isEliteName_(name){
    var x = _nl_(name);
    // D1+ (attention au '+'), CFP, LDP, Ligue 2/3
    return /(?:^|[^a-z0-9])d1\+(?=$|[^a-z0-9])|(?:^|[^a-z0-9])cfp(?=$|[^a-z0-9])|(?:^|[^a-z0-9])ldp(?=$|[^a-z0-9])|ligue\s*2|ligue\s*3/.test(x);
  }

  // Tagger **local** (ne pose PAS inscription_* ici ; on l’ajoute après selon Type)
  function deriveTagsLocal_(name) {
    var s = _nl_(name);
    var tags = [];

    // adapté
    var isAdapt = /\badapte\b/.test(s) || /\badapte?\b/.test(s);
    if (isAdapt) tags.push('adapte');

    // thématiques
    if (/\bcamp de selection\b|\bselection\b/.test(s)) tags.push('camp_selection');
    if (/\bcamp\b/.test(s)) tags.push('camp');
    if (/\bcdp\b|centre de developpement/.test(s)) tags.push('cdp');
    if (/entra[iî]neur|coach/.test(s)) tags.push('coach');
    if (/futsal/.test(s)) tags.push('futsal');
    if (/gardien/.test(s)) tags.push('gardien');
    if (/\badulte\b|seniors?/.test(s)) tags.push('adulte');

    // bandes U
    var m = s.match(/\bu-?\s?(\d{1,2})\b/);
    if (m) {
      var u = +m[1];
      if (u <= 8) tags.push('u4u8');
      else if (u <= 12) tags.push('u9u12');
      else if (u <= 18) tags.push('u13u18');
    }
    return tags;
  }

  function catFromTags_(tags) {
    if (tags.includes('cdp')) return 'CDP';
    if (tags.includes('camp')) return 'CAMP';
    if (tags.includes('futsal')) return 'FUTSAL';
    if (tags.includes('inscription_normale')) return 'SEASON';
    if (tags.includes('coach')) return 'COACH';
    return 'OTHER';
  }
  function audienceFromTags_(tags) {
    if (tags.includes('coach')) return 'Entraîneur';
    if (tags.includes('adulte')) return 'Adulte';
    return 'Joueur';
  }
  function programBandFromTags_(tags) {
    if (tags.includes('u4u8')) return 'U4-U8';
    if (tags.includes('u9u12')) return 'U9-U12';
    if (tags.includes('u13u18')) return 'U13-U18';
    if (tags.includes('adulte')) return 'Adulte';
    return '';
  }
  function parseMoney_(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    var s = String(v).replace(/\s/g, '').replace(/[^\d.,-]/g, '');
    var parts = s.split(',');
    if (parts.length > 1 && parts[parts.length - 1].length === 2 && s.indexOf('.') === -1) {
      s = s.replace(/\./g, '').replace(',', '.');
    } else {
      s = s.replace(/,/g, '');
    }
    var n = Number(s); return isNaN(n) ? 0 : n;
  }
  function paymentStatus_(due, paid, rest) {
    var d = parseMoney_(due), p = parseMoney_(paid);
    var r = (rest === '' || rest == null) ? (d - p) : parseMoney_(rest);
    if (r <= 0 && (d > 0 || p > 0)) return 'Paid';
    if (p > 0 && r > 0) return 'Partial';
    return 'Unpaid';
  }
  function makePS_(p, s) { return (p ? String(p).padStart(8, '0') : '') + '|' + String(s || ''); }
  function makeRowHash_(obj) {
    try {
      var s = JSON.stringify({ p: obj['Passeport #'], t: obj['Type'], n: obj['NomFrais'], sa: obj['Saison'] });
      var dig = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
      return Utilities.base64Encode(dig);
    } catch (e) { return ''; }
  }

  // --- build (mêmes règles que l’INCR)
  var rows = [];

  function consider_(type, r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p) return;

    var name = r['Nom du frais'] || r['Frais'] || r['Produit'] || r['NomFrais'] || '';
    var status = _isActiveRow_(r) ? 1 : 0;

    var amtDue  = r['Montant dû'] || r['Montant du'] || r['Montant_dû'] || r['MontantDu'] || r['Due'] || '';
    var amtPaid = r['Montant payé'] || r['Montant paye'] || r['MontantPayé'] || r['MontantPaye'] || r['Paid'] || '';
    var amtRest = r['Montant dû restant'] || r['Montant restant'] || r['Restant'] || r['Due restant'] || r['DueRestant'] || r['Balance'] || '';

    var tags = deriveTagsLocal_(name);
    var isEliteName = _isEliteName_(name);
    var isAdapt = tags.indexOf('adapte') >= 0;

    // --- Normalisation des tags selon le Type ---
    if (type === 'ARTICLE') {
      // ARTICLE → jamais inscription_*, mais marquage article_elite si applicable
      tags = tags.filter(function(t){ return t !== 'inscription_normale' && t !== 'inscription_elite'; });
      if (isEliteName && tags.indexOf('article_elite') < 0) tags.push('article_elite');
    } else if (type === 'INSCRIPTION') {
      // INSCRIPTION → jamais article_elite
      tags = tags.filter(function(t){ return t !== 'article_elite'; });
      if (isEliteName) {
        if (tags.indexOf('inscription_elite') < 0) tags.push('inscription_elite');
        tags = tags.filter(function(t){ return t !== 'inscription_normale'; });
      } else {
        var hasSaison = /saison/i.test(_nl_(name));
        tags = tags.filter(function(t){ return t !== 'inscription_elite'; });
        if (hasSaison && !isAdapt) {
          if (tags.indexOf('inscription_normale') < 0) tags.push('inscription_normale');
        } else {
          tags = tags.filter(function(t){ return t !== 'inscription_normale'; });
        }
      }
    }

    var isIgn = (tags.includes('adulte') || matchIgnore_(name)) ? 1 : 0;

    var row = {
      'Passeport #': p,
      'Type': type,
      'NomFrais': name,
      'Status': status,
      'isIgnored': isIgn,
      'RowHash': r['ROW_HASH'] || r['RowHash'] || '',
      'Saison': saison,
      'PS': makePS_(p, saison),
      'MapKey': slugify_(name),
      'Tags': tags.join(','),
      'CatCode': catFromTags_(tags),
      'Audience': audienceFromTags_(tags),
      'isCoachFee': tags.includes('coach') ? 1 : 0,
      'ProgramBand': programBandFromTags_(tags),
      'AmountDue': parseMoney_(amtDue),
      'AmountPaid': parseMoney_(amtPaid),
      'AmountRestant': parseMoney_(amtRest),
      'PaymentStatus': paymentStatus_(amtDue, amtPaid, amtRest),
      'Qty': r['Quantité'] || r['Qty'] || 1,
      'Date de la facture': r['Date de la facture'] || r['DateFacture'] || r['Date'] || '',
      'CreatedAt': new Date(),
      'UpdatedAt': new Date()
    };
    if (!row['RowHash']) row['RowHash'] = makeRowHash_(row);

    rows.push(row);
  }

  (insc.rows || []).forEach(function (r) { consider_('INSCRIPTION', r); });
  (art.rows  || []).forEach(function (r) { consider_('ARTICLE', r); });

writeObjectsToSheet_(ss, 'ACHATS_LEDGER', rows, [
  'Passeport #', 'Type', 'NomFrais', 'Status', 'isIgnored', 'RowHash', 'Saison',
  'PS', 'MapKey', 'Tags', 'CatCode', 'Audience', 'isCoachFee', 'ProgramBand',
  'AmountDue', 'AmountPaid', 'AmountRestant', 'PaymentStatus', 'Qty',
  'Date de la facture',
  'CreatedAt', 'UpdatedAt'
]);
}


/** =========================================
 *  ACHATS_LEDGER — INCR (par passeports)
 *  ========================================= */
function updateAchatsLedgerForPassports_(ss, touchedPassports) {
  var saison = readParam_(ss, 'SEASON_LABEL') || '';

  // Normalise passeports
  var normP = (typeof normalizePassport8_ === 'function')
    ? normalizePassport8_
    : function (p) { return (p == null || p === '') ? '' : String(p).replace(/\D/g, '').padStart(8, '0'); };

  // accepte Array, Set, objet {p:1}, etc.
  var rawSet = _toPassportSet_(touchedPassports);
  var touchedSet = new Set(Array.from(rawSet).map(normP).filter(Boolean));
  if (!touchedSet.size) return;

  // Ignore list
  var DEFAULT_IGNORE = 'senior,u-se,adulte,ligue';
  var ignoreCsvRaw = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV);
  var ignoreCsv = (ignoreCsvRaw && String(ignoreCsvRaw).trim()) ? ignoreCsvRaw : DEFAULT_IGNORE;

  function toIgnoreArr_(x) {
    if (!x) return [];
    if (Array.isArray(x)) return x;
    if (typeof x === 'string') return x.split(',');
    if (x && typeof x.forEach === 'function' && x.add) { var a = []; x.forEach(function(v){ a.push(v); }); return a; }
    if (typeof x === 'object') return Object.keys(x);
    return [];
  }
  var ignoreList = toIgnoreArr_(_compileCsvToSet_(ignoreCsv)).map(function(s){
    return String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  });
  function matchIgnore_(name) {
    var s = String(name || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    for (var i = 0; i < ignoreList.length; i++) { var k = ignoreList[i]; if (k && s.indexOf(k) !== -1) return true; }
    return false;
  }

  // Helpers (mêmes que FULL)
  function slugify_(s) {
    return String(s || '').toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  }
  function _nl_(s){ try { return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase(); } catch(e){ return String(s||'').toLowerCase(); } }
  function _isEliteName_(name){
    var x = _nl_(name);
    return /(?:^|[^a-z0-9])d1\+(?=$|[^a-z0-9])|(?:^|[^a-z0-9])cfp(?=$|[^a-z0-9])|(?:^|[^a-z0-9])ldp(?=$|[^a-z0-9])|ligue\s*2|ligue\s*3/.test(x);
  }
  function deriveTagsLocal_(name) {
    var s = _nl_(name);
    var tags = [];
    var isAdapt = /\badapte\b/.test(s) || /\badapte?\b/.test(s);
    if (isAdapt) tags.push('adapte');
    if (/\bcamp de selection\b|\bselection\b/.test(s)) tags.push('camp_selection');
    if (/\bcamp\b/.test(s)) tags.push('camp');
    if (/\bcdp\b|centre de developpement/.test(s)) tags.push('cdp');
    if (/entra[iî]neur|coach/.test(s)) tags.push('coach');
    if (/futsal/.test(s)) tags.push('futsal');
    if (/gardien/.test(s)) tags.push('gardien');
    if (/\badulte\b|seniors?/.test(s)) tags.push('adulte');
    var m = s.match(/\bu-?\s?(\d{1,2})\b/);
    if (m) {
      var u = +m[1];
      if (u <= 8) tags.push('u4u8');
      else if (u <= 12) tags.push('u9u12');
      else if (u <= 18) tags.push('u13u18');
    }
    return tags;
  }
  function catFromTags_(tags) {
    if (tags.includes('cdp')) return 'CDP';
    if (tags.includes('camp')) return 'CAMP';
    if (tags.includes('futsal')) return 'FUTSAL';
    if (tags.includes('inscription_normale')) return 'SEASON';
    if (tags.includes('coach')) return 'COACH';
    return 'OTHER';
  }
  function audienceFromTags_(tags) { return tags.includes('coach') ? 'Entraîneur' : (tags.includes('adulte') ? 'Adulte' : 'Joueur'); }
  function programBandFromTags_(tags) {
    if (tags.includes('u4u8')) return 'U4-U8';
    if (tags.includes('u9u12')) return 'U9-U12';
    if (tags.includes('u13u18')) return 'U13-U18';
    if (tags.includes('adulte')) return 'Adulte';
    return '';
  }
  function parseMoney_(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    var s = String(v).replace(/\s/g, '').replace(/[^\d.,-]/g, '');
    var parts = s.split(',');
    if (parts.length > 1 && parts[parts.length - 1].length === 2 && s.indexOf('.') === -1) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
    var n = Number(s); return isNaN(n) ? 0 : n;
  }
  function paymentStatus_(due, paid, rest) {
    var d = parseMoney_(due), p = parseMoney_(paid);
    var r = (rest === '' || rest == null) ? (d - p) : parseMoney_(rest);
    if (r <= 0 && (d > 0 || p > 0)) return 'Paid';
    if (p > 0 && r > 0) return 'Partial';
    return 'Unpaid';
  }
  function makePS_(p, s) { return (p ? String(p).padStart(8, '0') : '') + '|' + String(s || ''); }
  function makeRowHash_(obj) {
    try {
      var s = JSON.stringify({ p: obj['Passeport #'], t: obj['Type'], n: obj['NomFrais'], sa: obj['Saison'] });
      var dig = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
      return Utilities.base64Encode(dig);
    } catch (e) { return ''; }
  }

  // Lire les sources finales
  var shInscName = _pickSheetName_(ss, SHEETS.INSCRIPTIONS_FINAL, SHEETS.INSCRIPTIONS);
  var shArtName  = _pickSheetName_(ss, SHEETS.ARTICLES_FINAL, SHEETS.ARTICLES);
  var insc = readSheetAsObjects_(ss.getId(), shInscName);
  var art  = readSheetAsObjects_(ss.getId(), shArtName);

  // 1) Purge ciblée dans LEDGER (passeports touchés / saison courante)
  var ledger   = readSheetAsObjects_(ss.getId(), 'ACHATS_LEDGER');
  var existing = ledger.rows || [];
  var kept = existing.filter(function (r) {
    var p = normP(r['Passeport #'] || '');
    return !(touchedSet.has(p) && r['Saison'] === saison);
  });
  writeObjectsToSheet_(ss, 'ACHATS_LEDGER', kept, [
    'Passeport #', 'Type', 'NomFrais', 'Status', 'isIgnored', 'RowHash', 'Saison',
    'PS', 'MapKey', 'Tags', 'CatCode', 'Audience', 'isCoachFee', 'ProgramBand',
    'AmountDue', 'AmountPaid', 'AmountRestant', 'PaymentStatus', 'Qty',
    'CreatedAt', 'UpdatedAt'
  ]);

  // 2) Rebuild pour ces passeports (depuis FINAL)
  var toAppend = [];
  function consider_(type, r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p || !touchedSet.has(p)) return;

    var name   = r['Nom du frais'] || r['Frais'] || r['Produit'] || r['NomFrais'] || '';
    var status = _isActiveRow_(r) ? 1 : 0;

    var amtDue  = r['Montant dû'] || r['Montant du'] || r['Montant_dû'] || r['MontantDu'] || r['Due'] || '';
    var amtPaid = r['Montant payé'] || r['Montant paye'] || r['MontantPayé'] || r['MontantPaye'] || r['Paid'] || '';
    var amtRest = r['Montant dû restant'] || r['Montant restant'] || r['Restant'] || r['Due restant'] || r['DueRestant'] || r['Balance'] || '';

    var tags = deriveTagsLocal_(name);
    var isEliteName = _isEliteName_(name);
    var isAdapt = tags.indexOf('adapte') >= 0;

    // --- Normalisation des tags selon le Type (identique au FULL) ---
    if (type === 'ARTICLE') {
      tags = tags.filter(function(t){ return t !== 'inscription_normale' && t !== 'inscription_elite'; });
      if (isEliteName && tags.indexOf('article_elite') < 0) tags.push('article_elite');
    } else if (type === 'INSCRIPTION') {
      tags = tags.filter(function(t){ return t !== 'article_elite'; });
      if (isEliteName) {
        if (tags.indexOf('inscription_elite') < 0) tags.push('inscription_elite');
        tags = tags.filter(function(t){ return t !== 'inscription_normale'; });
      } else {
        var hasSaison = /saison/i.test(_nl_(name));
        tags = tags.filter(function(t){ return t !== 'inscription_elite'; });
        if (hasSaison && !isAdapt) {
          if (tags.indexOf('inscription_normale') < 0) tags.push('inscription_normale');
        } else {
          tags = tags.filter(function(t){ return t !== 'inscription_normale'; });
        }
      }
    }

    var isIgn = (tags.includes('adulte') || matchIgnore_(name)) ? 1 : 0;

    var row = {
      'Passeport #': p,
      'Type': type,
      'NomFrais': name,
      'Status': status,
      'isIgnored': isIgn,
      'RowHash': r['ROW_HASH'] || r['RowHash'] || '',
      'Saison': saison,
      'PS': makePS_(p, saison),
      'MapKey': slugify_(name),
      'Tags': tags.join(','),
      'CatCode': catFromTags_(tags),
      'Audience': audienceFromTags_(tags),
      'isCoachFee': tags.includes('coach') ? 1 : 0,
      'ProgramBand': programBandFromTags_(tags),
      'AmountDue': parseMoney_(amtDue),
      'AmountPaid': parseMoney_(amtPaid),
      'AmountRestant': parseMoney_(amtRest),
      'PaymentStatus': paymentStatus_(amtDue, amtPaid, amtRest),
      'Qty': r['Quantité'] || r['Qty'] || 1,
      'CreatedAt': new Date(),
      'UpdatedAt': new Date()
    };
    if (!row['RowHash']) row['RowHash'] = makeRowHash_(row);
    toAppend.push(row);
  }

  (insc.rows || []).forEach(function (r) { consider_('INSCRIPTION', r); });
  (art.rows  || []).forEach(function (r) { consider_('ARTICLE', r); });

  if (toAppend.length) appendObjectsToSheet_(ss, 'ACHATS_LEDGER', toAppend);
}


// --- Helpers FAST pour sérialiser les activités ---
// Retourne une version "light" des inscriptions
function _simplifyInscriptions(rows) {
  rows = rows || [];
  return rows
    .filter(function (r) {
      return String(r.Type || '').toUpperCase() === 'INSCRIPTION'; // ⬅️ Type strict
    })
    .map(function (r) {
      return {
        Produit: r['NomFrais'] || r['Nom du frais'] || r['Frais'] || r['Produit'] || '',
        Date: r['Date'] || r['Date d’achat'] || r['AchatDate'] || r['Achat'] || '',
        Montant: r['Montant'] || r['Prix'] || r['Total'] || r['AmountPaid'] || r['AmountDue'] || '',
        Tags: r['Tags'] || ''
      };
    });
}

function _simplifyArticles(rows) {
  rows = rows || [];
  return rows
    .filter(function (r) {
      return String(r.Type || '').toUpperCase() === 'ARTICLE'; // ⬅️ Type strict
    })
    .map(function (r) {
      return {
        Produit: r['NomFrais'] || r['Nom du frais'] || r['Frais'] || r['Produit'] || '',
        Date: r['Date'] || r['Date d’achat'] || r['AchatDate'] || r['Achat'] || '',
        Montant: r['Montant'] || r['Prix'] || r['Total'] || r['AmountPaid'] || r['AmountDue'] || '',
        Tags: r['Tags'] || ''
      };
    });
}


function _sx_absWarnDate_(ss) {
  var raw = readParam_(ss, 'RETRO_PHOTO_WARN_ABS_DATE') || '';
  if (!raw) return null;
  try { var d = new Date(raw); return isNaN(+d) ? null : d; } catch (e) { return null; }
}
function _sx_needsPhoto_(age, isAdapte, hasInscription) {
  var a = parseInt(String(age || ''), 10);
  var need = (a >= 8) && !isAdapte && !!hasInscription;
  return need;
}
function _sx_photoStr_(ss, exp, age, isAdapte, hasInscription) {
  if (!_sx_needsPhoto_(age, isAdapte, hasInscription)) return '';
  if (!exp) return 'Aucune photo';
  var abs = _sx_absWarnDate_(ss);
  try {
    var d = (exp instanceof Date) ? exp : new Date(exp);
    if (isNaN(+d)) return 'Aucune photo';
    if (abs && d < abs) return 'Expirée';
    return 'Valide';
  } catch (e) { return 'Aucune photo'; }
}



function refreshPhotoStrInJoueurs_(ssOrId) {
  // 1) classeur saison
  var ss = (ssOrId && typeof ssOrId.getId === 'function')
    ? ssOrId
    : ensureSpreadsheet_(ssOrId || getSeasonSpreadsheet_(getSeasonId_()));

  // 2) lecture JOUEURS (header + valeurs)
  var sh = ss.getSheetByName('JOUEURS');
  if (!sh || sh.getLastRow() < 2) return { updated: 0 };
  var H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  function c(name) { var i = H.indexOf(name); return i < 0 ? -1 : i; }
  var ciAge = c('Age');
  var ciAda = c('isAdapte');
  var ciHasIn = c('hasInscription');
  var ciExp = c('PhotoExpireLe');
  var ciStr = c('PhotoStr');
  var ciBand = c('AgeBracket'); // (non utilisé pour l’instant, on le garde si jamais)

  if (ciStr < 0) throw new Error("Colonne 'PhotoStr' introuvable dans JOUEURS.");

  var n = sh.getLastRow() - 1;
  var lc = sh.getLastColumn();
  // Lire les valeurs "brutes" (Dates/Numbers), pas les display values
  var vals = sh.getRange(2, 1, n, lc).getValues();

  // 3) cutoff: date absolue si présente, sinon logique standard
  var cutoffAbs = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';
  var cutoff = cutoffAbs ? new Date(cutoffAbs) : _getPhotoCutoffDate_(ss);

  function truthy(v) {
    v = String(v || '').toUpperCase();
    return v === '1' || v === 'TRUE' || v === 'OUI' || v === 'YES';
  }
  function needPhoto(ageVal, isAdapteVal, hasInsVal /*, bracket */) {
    var age = parseInt(String(ageVal || ''), 10);
    if (!isNaN(age) && age < 8) return false;
    if (truthy(isAdapteVal)) return false;
    // hasInscription: FALSE/Non => on ignore (champ vide)
    if (!truthy(hasInsVal)) return false;
    return true;
  }
  function statusFor(expVal, cutoffDate) {
    if (!expVal && expVal !== 0) return 'Aucune photo';
    var d = (expVal instanceof Date) ? expVal : new Date(expVal);
    if (isNaN(+d)) return 'Aucune photo';
    return (d < cutoffDate) ? 'Expirée' : 'Valide';
  }

  // 4) calcule la nouvelle colonne
  var outCol = new Array(n);
  for (var r = 0; r < n; r++) {
    var age = ciAge >= 0 ? vals[r][ciAge] : '';
    var adap = ciAda >= 0 ? vals[r][ciAda] : '';
    var hasIn = ciHasIn >= 0 ? vals[r][ciHasIn] : '';
    var exp = ciExp >= 0 ? vals[r][ciExp] : '';

    var str = needPhoto(age, adap, hasIn /*, vals[r][ciBand]*/) ? statusFor(exp, cutoff) : 'Non requis';
    outCol[r] = [str];
  }

  // 5) écrit UNIQUEMENT la colonne PhotoStr (en texte)
  sh.getRange(2, ciStr + 1, n, 1).setValues(outCol);

  return { updated: n };
}




/** =========================
 *  JOUEURS — FULL
 *  ========================= */
function buildJoueursIndex_(ss) {
  // --- setup & params
  ss = ensureSpreadsheet_(ss || getSeasonSpreadsheet_(getSeasonId_()));
  var saison = readParam_(ss, 'SEASON_LABEL') || '';
  var adapteCsv = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || 'adapté,adapte,adapte';
  var rxAdapte = _compileKeywordsToRegex_(adapteCsv);

  var normP = (typeof normalizePassport8_ === 'function')
    ? normalizePassport8_
    : function (p) { return (p == null || p === '') ? '' : String(p).replace(/\D/g, '').padStart(8, '0'); };

  // --- sources finales
  var shInscName = _pickSheetName_(ss, SHEETS.INSCRIPTIONS_FINAL, SHEETS.INSCRIPTIONS);
  var shArtName = _pickSheetName_(ss, SHEETS.ARTICLES_FINAL, SHEETS.ARTICLES);
  var inscF = readSheetAsObjects_(ss.getId(), shInscName);
  var artF = readSheetAsObjects_(ss.getId(), shArtName);

  // --- helpers Age (locaux à la fonction)
  function _seasonYear_() {
    var y = parseInt(String(readParam_(ss, 'SEASON_YEAR') || ''), 10);
    if (!isNaN(y) && y > 1900) return y;
    var lbl = String(readParam_(ss, 'SEASON_LABEL') || '');
    var m = lbl.match(/(20\d{2})/);
    return m ? parseInt(m[1], 10) : (new Date()).getFullYear();
  }
  function _birthYearFrom_(v) {
    if (!v && v !== 0) return null;
    if (v instanceof Date && !isNaN(+v)) return v.getFullYear();
    var s = String(v || '').trim(); if (!s) return null;
    var m1 = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if (m1) return +m1[1];
    var m2 = s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/); if (m2) return +m2[3];
    return null;
  }
  var __SY = _seasonYear_();




  // --- identité par passeport (priorité INSCRIPTIONS, fallback ARTICLES)
  var idByPass = {};
  (inscF.rows || []).forEach(function (r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p) return;
    var id = idByPass[p];
    if (!id) {
      idByPass[p] = {
        Nom: r['Nom de famille'] || r['Nom'] || '',
        Prenom: r['Prénom'] || r['Prenom'] || '',
        DateNaissance: r['Date de naissance'] || r['DateNaissance'] || '',
        Age: r['Âge'] || r['Age'] || '',   // <— ajout (accent toléré)
        Genre: r['Identité de genre'] || r['Sexe'] || r['Genre'] || '',
        Adresse: r['Adresse'] || '',
        Ville: r['Ville'] || '',
        Province: r['Province'] || '',
        CodePostal: r['Code postal'] || r['CodePostal'] || '',
        Téléphone1: r['Téléphone1'] || r['Tel1'] || r['Téléphone'] || '',
        Téléphone2: r['Téléphone2'] || r['Tel2'] || '',
        PhotoExpireLe: r['PhotoExpireLe'] || '',
        isAdapte: String(r['Programme adapté'] || r['Programme'] || '').match(rxAdapte) ? 1 : 0,
        typeMembre: r['Type de membre'] || r['typeMembre'] || ''
      };
    } else {
      if (!id.Genre && (r['Identité de genre'] || r['Sexe'] || r['Genre'])) id.Genre = r['Identité de genre'] || r['Sexe'] || r['Genre'] || '';
      if (!id.PhotoExpireLe && r['PhotoExpireLe']) id.PhotoExpireLe = r['PhotoExpireLe'];
      if (!id.isAdapte) id.isAdapte = String(r['Programme adapté'] || r['Programme'] || '').match(rxAdapte) ? 1 : 0;
      if (!id.typeMembre && (r['Type de membre'] || r['typeMembre'])) id.typeMembre = r['Type de membre'] || r['typeMembre'] || '';
    }
  });
  (artF.rows || []).forEach(function (r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p) return;
    var id = idByPass[p];
    if (!id) {
      idByPass[p] = {
        Nom: r['Nom de famille'] || r['Nom'] || '',
        Prenom: r['Prénom'] || r['Prenom'] || '',
        DateNaissance: r['Date de naissance'] || r['DateNaissance'] || r['Naissance'] || '',
        Age: r['Âge'] || r['Age'] || '',   // <— ajout (accent toléré)
        Genre: r['Sexe'] || r['Genre'] || r['Identité de genre'] || '',
        Adresse: r['Adresse'] || '',
        Ville: r['Ville'] || '',
        Province: r['Province'] || '',
        CodePostal: r['Code postal'] || r['CodePostal'] || '',
        Téléphone1: r['Téléphone1'] || r['Tel1'] || r['Téléphone'] || '',
        Téléphone2: r['Téléphone2'] || r['Tel2'] || '',
        PhotoExpireLe: r['PhotoExpireLe'] || '',
        isAdapte: 0,
        typeMembre: r['Type de membre'] || r['typeMembre'] || ''
      };
    } else {
      if (!id.Nom) id.Nom = r['Nom de famille'] || r['Nom'] || '';
      if (!id.Prenom) id.Prenom = r['Prénom'] || r['Prenom'] || '';
      if (!id.DateNaissance) id.DateNaissance = r['Date de naissance'] || r['DateNaissance'] || r['Naissance'] || '';
      if (!id.Genre) id.Genre = r['Sexe'] || r['Genre'] || r['Identité de genre'] || '';
      if (!id.PhotoExpireLe) id.PhotoExpireLe = r['PhotoExpireLe'] || '';
      if (!id.Adresse) id.Adresse = r['Adresse'] || '';
      if (!id.Ville) id.Ville = r['Ville'] || '';
      if (!id.Province) id.Province = r['Province'] || '';
      if (!id.CodePostal) id.CodePostal = r['Code postal'] || r['CodePostal'] || '';
      if (!id.Téléphone1) id.Téléphone1 = r['Téléphone1'] || r['Tel1'] || r['Téléphone'] || '';
      if (!id.Téléphone2) id.Téléphone2 = r['Téléphone2'] || r['Tel2'] || '';
      if (!id.typeMembre && (r['Type de membre'] || r['typeMembre'])) id.typeMembre = r['Type de membre'] || r['typeMembre'] || '';
    }
  });

  // --- MEMBRES_GLOBAL : PhotoExpireLe en priorité si JOUEURS vide
  var mgSheet = ss.getSheetByName(readParam_(ss, 'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL');
  if (mgSheet) {
    var mg = mgSheet.getDataRange().getValues(), H = mg[0] || [];
    var idxP = H.indexOf('Passeport #'); if (idxP < 0) idxP = H.indexOf('Passeport');
    var idxPhoto = H.indexOf('PhotoExpireLe');
    if (idxP >= 0 && idxPhoto >= 0) {
      for (var i = 1; i < mg.length; i++) {
        var pid = normP(mg[i][idxP]); if (!pid) continue;
        var ph = mg[i][idxPhoto];
        if (ph && idByPass[pid] && !idByPass[pid].PhotoExpireLe) idByPass[pid].PhotoExpireLe = ph;
      }
    }
  }

  // --- LEDGER : activités par passeport (actifs, saison courante, non ignorés)
  var ledger = readSheetAsObjects_(ss.getId(), SHEETS.ACHATS_LEDGER);
  var actByP = new Map();
  (ledger.rows || []).forEach(function (r) {
    if (r['Saison'] !== saison) return;
    if ((Number(r['Status']) || 0) !== 1) return;
    if ((Number(r['isIgnored']) || 0) === 1) return;
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p) return;
    if (!actByP.has(p)) actByP.set(p, []);
    var tags = (r['Tags'] ? String(r['Tags']).split(',').map(function (x) { return x.trim(); }).filter(Boolean) : []);
    actByP.get(p).push({
      Type: r['Type'],
      Tags: tags,
      Audience: r['Audience'] || '',
      ProgramBand: r['ProgramBand'] || '',
      NomFrais: r['NomFrais'] || r['Nom du frais'] || r['Frais'] || r['Produit'] || ''
    });
  });

  // --- COURRIELS : index en un seul passage (évite O(N²))
  var emailsByPass = new Map();
  function _pushEmail(p, eCsv) {
    if (!eCsv) return;
    var set = emailsByPass.get(p); if (!set) { set = new Set(); emailsByPass.set(p, set); }
    String(eCsv).split(/[;,]/).forEach(function (x) { x = x.trim(); if (x) set.add(x); });
  }
  (inscF.rows || []).forEach(function (r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p) return;
    var e = (typeof collectEmailsFromRow_ === 'function')
      ? collectEmailsFromRow_(r, 'Courriel,Parent 1 - Courriel,Parent 2 - Courriel')
      : (r['Courriel'] || r['Parent 1 - Courriel'] || r['Parent 2 - Courriel'] || '');
    _pushEmail(p, e);
  });
  (artF.rows || []).forEach(function (r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p) return;
    _pushEmail(p, r['Courriel']);
  });

  // --- cutoff photo calculé une seule fois (pas de readParam_ par joueur)
  var cutoff = _getPhotoCutoffDate_(ss);
  function photoStatusWithCutoff(d, ageBracket, isAdapteFlag) {
    if (isAdapteFlag === true || /^1$|^true$/i.test(String(isAdapteFlag)) || _isPhotoNotRequiredBracket_(ageBracket)) {
      return 'Non requis';
    }
    if (!d) return 'Invalide (aucune photo)';
    try {
      var dt = (d instanceof Date) ? d : new Date(d);
      if (isNaN(+dt)) return 'Invalide (aucune photo)';
      return (dt < cutoff) ? 'Invalide (expirée)' : 'Valide';
    } catch (e) {
      return 'Invalide (aucune photo)';
    }
  }

  // --- utils
  var pickAgeBracketFromLedgerRows_ = (typeof this.pickAgeBracketFromLedgerRows_ === 'function')
    ? this.pickAgeBracketFromLedgerRows_
    : function (rows) {
      var set = {};
      for (var i = 0; i < (rows || []).length; i++) {
        var b = rows[i].ProgramBand || '';
        if (b) set[b] = 1;
      }
      if (set['U4-U8']) return 'U4-U8';
      if (set['U9-U12']) return 'U9-U12';
      if (set['U13-U18']) return 'U13-U18';
      if (set['Adulte']) return 'Adulte';
      return '';
    };

  function pickPrimaryEmail_(emailsStr) {
    var arr = String(emailsStr || '').split(/[;,]/).map(function (s) { return s.trim(); }).filter(Boolean);
    var bad = /noreply|no-reply|invalid|test|example/i;
    for (var i = 0; i < arr.length; i++) if (!bad.test(arr[i])) return arr[i];
    return arr[0] || '';
  }

  // --- build sortie
  var out = [];
  var seen = {};
  Object.keys(idByPass).forEach(function (p) { seen[p] = 1; });
  Array.from(actByP.keys()).forEach(function (p) { seen[p] = 1; });

  Object.keys(seen).forEach(function (p) {
    var id = idByPass[p] || {};
    var act = actByP.get(p) || [];

    // Aucun acte admissible → on skip (comme avant)
    if (!act.length) return;

    var nom = id.Nom || '';
    var prenom = id.Prenom || '';
    var fullName = (prenom + ' ' + nom).trim();

    var dob = id.DateNaissance || '';
    var age = (function () {
      var a = parseInt(String(id.Age || ''), 10);
      if (!isNaN(a) && a > 0 && a < 99) return a;
      var by = _birthYearFrom_(dob);
      if (!by) return '';
      var val = __SY - by;
      return (val > 0 && val < 99) ? val : '';
    })(); var genre = (id.Genre || '').toString().trim();
    if (genre) { var g0 = genre.toUpperCase().charAt(0); genre = (g0 === 'M' || g0 === 'F') ? g0 : ''; }

    var adresse = id.Adresse || '';
    var ville = id.Ville || '';
    var province = id.Province || '';
    var codePostal = id.CodePostal || '';
    var tel1 = id.Téléphone1 || '';
    var tel2 = id.Téléphone2 || '';
    var parent1 = id.Parent1 || id['Nom du parent 1'] || '';
    var parent2 = id.Parent2 || id['Nom du parent 2'] || '';
    var mail1 = id.CourrielParent1 || '';
    var mail2 = id.CourrielParent2 || '';

    var typeMem = String(id.typeMembre || '').trim();
    if (!typeMem && act.length) {
      var hasAdulteTag = act.some(function (r) { return String(r.Tags || '').toLowerCase().indexOf('adulte') >= 0; });
      if (hasAdulteTag) typeMem = 'ADULTE';
    }

    var ageBracket = pickAgeBracketFromLedgerRows_(act);
    var programBand = ageBracket;
    act.forEach(function (r) { if (r.ProgramBand) programBand = r.ProgramBand; });

    var hasInscription = act.some(function (r) {
      var t = String(r.Type || '').toLowerCase();
      var tg = String(r.Tags || '').toLowerCase();
      return t === 'inscription' || tg.indexOf('inscription') >= 0 || tg.indexOf('inscription_normale') >= 0;
    });
    var hasCamp = act.some(function (r) { return String(r.Tags || '').toLowerCase().indexOf('camp') >= 0; });
    var isCoach = act.some(function (r) { return String(r.Tags || '').toLowerCase().indexOf('coach') >= 0; });

    var inU9U12 = /U9|U10|U11|U12/.test(String(programBand || '').toUpperCase());
    var cdpCount = '';
    if (inU9U12) {
      cdpCount = act.filter(function (r) {
        return String(r.ProgramBand || '') === 'U9-U12' && String(r.Tags || '').toLowerCase().indexOf('cdp') >= 0;
      }).length || 0;
    }

    var isAdapte = id.isAdapte ? 1 : 0;
    if (!isAdapte && act.length) {
      isAdapte = act.some(function (r) { return String(r.Tags || '').toLowerCase().indexOf('adapte') >= 0; }) ? 1 : 0;
      if (!isAdapte && rxAdapte) {
        isAdapte = act.some(function (r) { return rxAdapte.test(String(r.NomFrais || '')); }) ? 1 : 0;
      }
    }
    if (isAdapte) cdpCount = '';

    var emailsAll = '';
    var emSet = emailsByPass.get(p);
    if (emSet && emSet.size) emailsAll = Array.from(emSet).join('; ');
    var primary = pickPrimaryEmail_(emailsAll);

    var photoExpDate = id.PhotoExpireLe || '';
    var photoStr = photoStatusWithCutoff(photoExpDate, ageBracket, isAdapte);

    out.push({
      'Passeport #': p,
      'Nom': nom,
      'Prénom': prenom,
      'NomComplet': fullName,
      'Courriels': emailsAll,
      'Saison': saison,
      'PS': '',
      'Adresse': adresse,
      'Ville': ville,
      'Province': province,
      'CodePostal': codePostal,
      'Téléphone1': tel1,
      'Téléphone2': tel2,
      'DateNaissance': dob,
      'Age': age,
      'AgeBracket': ageBracket,
      'Genre': genre,
      'Nom du parent 1': parent1,
      'Nom du parent 2': parent2,
      'CourrielParent1': mail1,
      'CourrielParent2': mail2,
      'U': (String(programBand || '').match(/^U\d+/i) ? String(programBand).toUpperCase() : ''),
      'typeMembre': typeMem,
      'isCoach': isCoach ? 'TRUE' : 'FALSE',
      'isAdapte': isAdapte ? '1' : '0',
      'ProgramBand': programBand,
      'hasInscription': hasInscription ? 'TRUE' : 'FALSE',
      'hasCamp': hasCamp ? 'TRUE' : 'FALSE',
      'CDP': cdpCount,
      'DerniereMaj': new Date().toISOString().slice(0, 10),
      'PhotoExpireLe': photoExpDate,
      'PhotoStr': photoStr,
      'InscriptionsJSON': JSON.stringify(_simplifyInscriptions(act)),
      'ArticlesJSON': JSON.stringify(_simplifyArticles(act))
    });
  });
  return { header: Object.keys(out[0] || {}), rows: out };
}




/** =========================================
 *  JOUEURS — INCR (par passeports)
 *  ========================================= */
function updateJoueursForPassports_(ss, touchedPassports) {
  var saison = readParam_(ss, 'SEASON_LABEL') || '';
  var adapteCsv = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || 'adapté,adapte,adapte';
  var rxAdapte = _compileKeywordsToRegex_(adapteCsv);

  var normP = (typeof normalizePassport8_ === 'function')
    ? normalizePassport8_
    : function (p) { return (p == null || p === '') ? '' : String(p).replace(/\D/g, '').padStart(8, '0'); };


  // --- helpers Age (locaux à la fonction)
  function _seasonYear_() {
    var y = parseInt(String(readParam_(ss, 'SEASON_YEAR') || ''), 10);
    if (!isNaN(y) && y > 1900) return y;
    var lbl = String(readParam_(ss, 'SEASON_LABEL') || '');
    var m = lbl.match(/(20\d{2})/);
    return m ? parseInt(m[1], 10) : (new Date()).getFullYear();
  }
  function _birthYearFrom_(v) {
    if (!v && v !== 0) return null;
    if (v instanceof Date && !isNaN(+v)) return v.getFullYear();
    var s = String(v || '').trim(); if (!s) return null;
    var m1 = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if (m1) return +m1[1];
    var m2 = s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/); if (m2) return +m2[3];
    return null;
  }
  var __SY = _seasonYear_();


  // Normalise l’ensemble des passeports touchés
  var rawSet = _toPassportSet_(touchedPassports);
  var touchedSet = new Set(Array.from(rawSet).map(normP).filter(Boolean));
  if (!touchedSet.size) return;

  function pickPrimaryEmail_(emailsStr) {
    var arr = String(emailsStr || '').split(/[;]/).map(function (s) { return s.trim(); }).filter(Boolean);
    var bad = /noreply|no-reply|invalid|test|example/i;
    for (var i = 0; i < arr.length; i++) { if (!bad.test(arr[i])) return arr[i]; }
    return arr[0] || '';
  }
  function pickAgeBracketFromLedgerRows_(rows) {
    var set = {};
    for (var i = 0; i < rows.length; i++) { var b = rows[i].ProgramBand || ''; if (b) set[b] = 1; }
    if (set['U4-U8']) return 'U4-U8';
    if (set['U9-U12']) return 'U9-U12';
    if (set['U13-U18']) return 'U13-U18';
    if (set['Adulte']) return 'Adulte';
    return '';
  }
  function makePS_(p, s) { return (p ? String(p).padStart(8, '0') : '') + '|' + String(s || ''); }

  // I/O
  var ledger = readSheetAsObjects_(ss.getId(), 'ACHATS_LEDGER');

  var shInscName = _pickSheetName_(ss, SHEETS.INSCRIPTIONS_FINAL, SHEETS.INSCRIPTIONS);
  var shArtName = _pickSheetName_(ss, SHEETS.ARTICLES_FINAL, SHEETS.ARTICLES);
  var inscF = readSheetAsObjects_(ss.getId(), shInscName);
  var artF = readSheetAsObjects_(ss.getId(), shArtName);

  // Identité: INSCRIPTIONS_FINAL -> ARTICLES_FINAL
  var idByPass = {};
  (inscF.rows || []).forEach(function (r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p || !touchedSet.has(p)) return;
    if (!idByPass[p]) {
      idByPass[p] = {
        Nom: r['Nom de famille'] || r['Nom'] || '',
        Prenom: r['Prénom'] || r['Prenom'] || '',
        DateNaissance: r['Date de naissance'] || r['DateNaissance'] || '',
        Genre: r['Identité de genre'] || r['Sexe'] || r['Genre'] || '',
        Adresse: r['Adresse'] || '',
        Ville: r['Ville'] || '',
        Province: r['Province'] || '',
        CodePostal: r['Code postal'] || r['CodePostal'] || '',
        Téléphone1: r['Téléphone1'] || r['Tel1'] || r['Téléphone'] || '',
        Téléphone2: r['Téléphone2'] || r['Tel2'] || '',
        PhotoExpireLe: r['PhotoExpireLe'] || '',
        isAdapte: String(r['Programme adapté'] || r['Programme'] || '').match(rxAdapte) ? 1 : 0,
        typeMembre: r['Type de membre'] || r['typeMembre'] || ''
      };
    } else {
      var id = idByPass[p];
      if (!id.Genre && (r['Identité de genre'] || r['Sexe'] || r['Genre'])) id.Genre = r['Identité de genre'] || r['Sexe'] || r['Genre'] || '';
      if (!id.PhotoExpireLe && r['PhotoExpireLe']) id.PhotoExpireLe = r['PhotoExpireLe'];
      if (!id.isAdapte) id.isAdapte = String(r['Programme adapté'] || r['Programme'] || '').match(rxAdapte) ? 1 : 0;
      if (!id.typeMembre && (r['Type de membre'] || r['typeMembre'])) id.typeMembre = r['Type de membre'] || r['typeMembre'] || '';
    }
  });
  (artF.rows || []).forEach(function (r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p || !touchedSet.has(p)) return;
    if (!idByPass[p]) {
      idByPass[p] = {
        Nom: r['Nom de famille'] || r['Nom'] || '',
        Prenom: r['Prénom'] || r['Prenom'] || '',
        DateNaissance: r['Date de naissance'] || r['DateNaissance'] || r['Naissance'] || '',
        Genre: r['Sexe'] || r['Genre'] || r['Identité de genre'] || '',
        Adresse: r['Adresse'] || '',
        Ville: r['Ville'] || '',
        Province: r['Province'] || '',
        CodePostal: r['Code postal'] || r['CodePostal'] || '',
        Téléphone1: r['Téléphone1'] || r['Tel1'] || r['Téléphone'] || '',
        Téléphone2: r['Téléphone2'] || r['Tel2'] || '',
        PhotoExpireLe: r['PhotoExpireLe'] || '',
        isAdapte: 0,
        typeMembre: r['Type de membre'] || r['typeMembre'] || ''
      };
    } else {
      var id = idByPass[p];
      if (!id.Nom) id.Nom = r['Nom de famille'] || r['Nom'] || '';
      if (!id.Prenom) id.Prenom = r['Prénom'] || r['Prenom'] || '';
      if (!id.DateNaissance) id.DateNaissance = r['Date de naissance'] || r['DateNaissance'] || r['Naissance'] || '';
      if (!id.Genre) id.Genre = r['Sexe'] || r['Genre'] || r['Identité de genre'] || '';
      if (!id.PhotoExpireLe) id.PhotoExpireLe = r['PhotoExpireLe'] || '';
      if (!id.Adresse) id.Adresse = r['Adresse'] || '';
      if (!id.Ville) id.Ville = r['Ville'] || '';
      if (!id.Province) id.Province = r['Province'] || '';
      if (!id.CodePostal) id.CodePostal = r['Code postal'] || r['CodePostal'] || '';
      if (!id.Téléphone1) id.Téléphone1 = r['Téléphone1'] || r['Tel1'] || r['Téléphone'] || '';
      if (!id.Téléphone2) id.Téléphone2 = r['Téléphone2'] || r['Tel2'] || '';
      if (!id.typeMembre && (r['Type de membre'] || r['typeMembre'])) id.typeMembre = r['Type de membre'] || r['typeMembre'] || '';
    }
  });

  // === NEW === MEMBRES_GLOBAL → priorité pour PhotoExpireLe si vide
  var mgSheet = ss.getSheetByName(readParam_(ss, 'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL');
  if (mgSheet) {
    var mg = mgSheet.getDataRange().getValues(), H = mg[0] || [];
    var idxP = H.indexOf('Passeport #'); if (idxP < 0) idxP = H.indexOf('Passeport');
    var idxPhoto = H.indexOf('PhotoExpireLe');
    if (idxP >= 0 && idxPhoto >= 0) {
      for (var i = 1; i < mg.length; i++) {
        var pid = normP(mg[i][idxP]); if (!pid || !touchedSet.has(pid)) continue;
        var ph = mg[i][idxPhoto];
        if (ph && idByPass[pid] && !idByPass[pid].PhotoExpireLe) idByPass[pid].PhotoExpireLe = ph;
      }
    }
  }

  // Activités par passeport (actifs, saison courante, non ignorés)
  var actByP = new Map();
  (ledger.rows || []).forEach(function (r) {
    if (r['Saison'] !== saison) return;
    if ((Number(r['Status']) || 0) !== 1) return;
    if ((Number(r['isIgnored']) || 0) === 1) return;
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p || !touchedSet.has(p)) return;
    if (!actByP.has(p)) actByP.set(p, []);
    var tags = (r['Tags'] ? String(r['Tags']).split(',').map(function (x) { return x.trim(); }).filter(Boolean) : []);
    actByP.get(p).push({
      Type: r['Type'],
      Tags: tags,
      Audience: r['Audience'] || '',
      ProgramBand: r['ProgramBand'] || '',
      NomFrais: r['NomFrais'] || r['Nom du frais'] || r['Frais'] || r['Produit'] || ''
    });
  });

  // Courriels indexés
  var emailsByPass = new Map();
  function _pushEmail(p, eCsv) {
    if (!eCsv) return;
    var set = emailsByPass.get(p); if (!set) { set = new Set(); emailsByPass.set(p, set); }
    String(eCsv).split(/[;]/).forEach(function (x) { x = x.trim(); if (x) set.add(x); });
  }
  (inscF.rows || []).forEach(function (r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p || !touchedSet.has(p)) return;
    var e = (typeof collectEmailsFromRow_ === 'function')
      ? collectEmailsFromRow_(r, 'Courriel,Parent 1 - Courriel,Parent 2 - Courriel')
      : (r['Courriel'] || r['Parent 1 - Courriel'] || r['Parent 2 - Courriel'] || '');
    _pushEmail(p, e);
  });
  (artF.rows || []).forEach(function (r) {
    var p = normP(r['Passeport #'] || r['Passeport'] || ''); if (!p || !touchedSet.has(p)) return;
    _pushEmail(p, r['Courriel']);
  });

  // Build nouvelles lignes
  var newRows = [];
  touchedSet.forEach(function (p) {
    var id = idByPass[p] || {};
    var nom = id.Nom || '';
    var pren = id.Prenom || '';
    var dna = id.DateNaissance || '';
    var genre = id.Genre || '';
    var emailsAll = Array.from(emailsByPass.get(p) || []).join('; ');
    var primary = pickPrimaryEmail_(emailsAll);
    var act = actByP.get(p) || [];
    var programBand = (function () {
      var set = {}; act.forEach(function (r) { if (r.ProgramBand) set[r.ProgramBand] = 1; });
      if (set['U4-U8']) return 'U4-U8';
      if (set['U9-U12']) return 'U9-U12';
      if (set['U13-U18']) return 'U13-U18';
      if (set['Adulte']) return 'Adulte';
      return '';
    })();
    var ageBracket = pickAgeBracketFromLedgerRows_(act);

    var hasInscription = act.some(function (r) { return r.Type === 'INSCRIPTION' && r.Tags.indexOf('inscription_normale') >= 0; });
    var hasCamp = act.some(function (r) { return r.Tags.indexOf('camp') >= 0; });
    var isCoach = act.some(function (r) { return r.Tags.indexOf('coach') >= 0; });
    var isAdapte = !!(act.some(function (r) { return r.Tags.indexOf('adapte') >= 0; }) || id.isAdapte);

    // CDP (U9-U12)
    var cdpCount = '';
    if (act.some(function (r) { return r.Audience === 'Joueur' && r.Tags.indexOf('inscription_normale') >= 0 && r.ProgramBand === 'U9-U12'; })) {
      cdpCount = act.filter(function (r) { return r.ProgramBand === 'U9-U12' && r.Tags.indexOf('cdp') >= 0; }).length || 0;
    }

    // === FIX === statut photo avec ss + fallback MEMBRES_GLOBAL déjà appliqué dans idByPass
    var photoStr = computePhotoStatusByYear_(id.PhotoExpireLe, ageBracket, isAdapte, ss);

    var age = (function () {
      var a = parseInt(String(id.Age || ''), 10);
      if (!isNaN(a) && a > 0 && a < 99) return a;
      var by = _birthYearFrom_(dna);
      if (!by) return '';
      var val = __SY - by;
      return (val > 0 && val < 99) ? val : '';
    })();


    newRows.push({
      'Passeport #': p,
      'Nom': nom,
      'Prénom': pren,
      'DateNaissance': dna,
      'Genre': genre,
      'Courriels': emailsAll,
      'isAdapte': isAdapte ? 1 : 0,
      'cdpCount': cdpCount,
      'hasCamp': hasCamp ? 'Oui' : 'Non',
      'hasInscription': hasInscription ? 'Oui' : 'Non',
      'PhotoExpireLe': id.PhotoExpireLe || '',
      'PhotoStr': photoStr,
      'InscriptionsJSON': JSON.stringify(_simplifyInscriptions(act)),
      'ArticlesJSON': JSON.stringify(_simplifyArticles(act)),
      'Saison': saison,
      'PS': makePS_(p, saison),
      'CourrielPrimaire': primary,
      'AgeBracket': ageBracket,
      'Age': age,
      'typeMembre': (isCoach ? 'Entraîneur' : (ageBracket === 'Adulte' ? 'Adulte' : 'Joueur')),
      'isCoach': isCoach ? 1 : 0,
      'DerniereMaj': new Date()
    });
  });

  // Upsert léger : purge lignes touchées, puis append
  var joueurs = readSheetAsObjects_(ss.getId(), 'JOUEURS');
  var existing = joueurs.rows || [];
  var kept = existing.filter(function (r) {
    var p = normP(r['Passeport #'] || '');
    return !touchedSet.has(p);
  });
  writeObjectsToSheet_(ss, 'JOUEURS', kept, [
    'Passeport #', 'Nom', 'Prénom', 'DateNaissance', 'Genre', 'Courriels',
    'Age',                              // <— AJOUT ICI
    'isAdapte', 'cdpCount', 'hasCamp', 'hasInscription',
    'PhotoExpireLe', 'PhotoStr', 'InscriptionsJSON', 'ArticlesJSON',
    'Saison', 'PS', 'CourrielPrimaire', 'AgeBracket', 'typeMembre', 'isCoach', 'DerniereMaj'
  ]);

  if (newRows.length) appendObjectsToSheet_(ss, 'JOUEURS', newRows);
}

/**
 * Détermine la date de cutoff à utiliser pour l’évaluation des photos.
 * Priorité:
 * 1) PARAMS.RETRO_PHOTO_WARN_ABS_DATE (date absolue, ex: "2026-03-01")
 * 2) PARAMS.RETRO_PHOTO_WARN_BEFORE_MMDD + année de saison (ex: "03-01" → 2026-03-01)
 * 3) 31 décembre de l’année de saison
 */
function _getPhotoCutoffDate_(ss) {
  var abs = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';
  if (abs) {
    var d = new Date(abs);
    if (!isNaN(+d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
  }

  var saisonLabel = readParam_(ss, 'SEASON_LABEL') || '';
  var seasonYear = parseSeasonYear_(saisonLabel) || (new Date()).getFullYear();
  var mmdd = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD) || ''; // ex "03-01"
  if (mmdd && /^\d{2}-\d{2}$/.test(mmdd)) {
    var mm = +mmdd.slice(0, 2), dd = +mmdd.slice(3, 5);
    return new Date(seasonYear, mm - 1, dd, 23, 59, 59, 999);
  }

  // défaut: fin d’année de saison
  return new Date(seasonYear, 11, 31, 23, 59, 59, 999);
}

/** vrai si la photo n’est pas requise selon le secteur (U4–U7) */
function _isPhotoNotRequiredBracket_(ageBracket) {
  var s = String(ageBracket || '').toUpperCase();
  // couvre "U4-U8", "U4U8", "U4", etc. → on ne requiert que U4–U7
  return /\bU4\b|\bU5\b|\bU6\b|\bU7\b/.test(s.replace(/[\s]/g, ''));
}

/**
 * Nouvelle logique demandée:
 * - cutoff: date variable (paramétrable)
 * - si exp < cutoff  → "Invalide (expirée)"
 * - si secteur non requis (U4–U7) ou Adapté → "Non requis"
 * - si aucune photo → "Invalide (aucune photo)"
 * - si exp >= cutoff → "Valide"
 */
function computePhotoStatusByYear_(d, ageBracket, isAdapte, ss) {
  // 1) non requis (Adapté ou U4–U7)
  if (isAdapte === true || /^1$|^true$/i.test(String(isAdapte)) || _isPhotoNotRequiredBracket_(ageBracket)) {
    return 'Non requis';
  }

  // 2) aucune photo
  if (!d) return 'Invalide (aucune photo)';

  // 3) comparer à cutoff
  try {
    var dt = (d instanceof Date) ? d : new Date(d);
    if (isNaN(+dt)) return 'Invalide (aucune photo)';
    var cutoff = _getPhotoCutoffDate_(ss);
    return (dt < cutoff) ? 'Invalide (expirée)' : 'Valide';
  } catch (e) {
    return 'Invalide (aucune photo)';
  }
}

// Alias compat (vieux nom)
function PhotoStatusByYear_(d, ageBracket, isAdapte) {
  return computePhotoStatusByYear_(d, ageBracket, isAdapte);
}

/* ==================
 * Helpers locaux
 * ================== */
function _compileCsvToSet_(csv) {
  return new Set(String(csv || '').split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean));
}
function _compileKeywordsToRegex_(csv) {
  var arr = String(csv || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  if (!arr.length) return null;
  var esc = arr.map(function (s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); });
  return new RegExp('(?:^|\\b)(' + esc.join('|') + ')(?:\\b|$)', 'i');
}
function _feeIgnoredLocal_(name, ignoreSet) {
  var key = _feeKey_(name);
  return ignoreSet.has(key);
}
function _feeKey_(name) {
  return String(name || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}
function _toPassportSet_(arrOrSet) {
  if (!arrOrSet) return new Set();
  if (arrOrSet instanceof Set) return arrOrSet;
  if (Array.isArray(arrOrSet)) return new Set(arrOrSet.map(function (x) { return String(x).trim(); }));
  // objet map {p:1} ?
  var out = new Set();
  Object.keys(arrOrSet).forEach(function (k) { if (arrOrSet[k]) out.add(String(k).trim()); });
  return out;
}
function _indexBy_(rows, keyFn) {
  var o = {}; (rows || []).forEach(function (r) { var k = keyFn(r); if (k) o[k] = r; }); return o;
}
function _addEmails_(map, p, csv) {
  var arr = String(csv || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean);
  if (!map[p]) map[p] = [];
  arr.forEach(function (e) { if (map[p].indexOf(e) === -1) map[p].push(e); });
}
function _isCdpName_(name) {
  var s = _feeKey_(name);
  return s.indexOf('cdp') !== -1;
}
function _cdpCountFromName_(name) {
  var s = _feeKey_(name);
  if (/\b2\b/.test(s) || /2\s*entrainement/.test(s) || /2\s*entrainements/.test(s)) return 2;
  if (/\b1\b/.test(s) || /1\s*entrainement/.test(s) || /1\s*entrainements/.test(s)) return 1;
  return 1; // CDP sans chiffre explicite => au moins 1
}

// --- Helpers JSON (hors boucle: mets-les en haut de la fonction) ---

function normName_(x) { return (x && (x.NomFrais || x.name || x.nom || '')) || ''; }
function feeGroup_(exclusiveGroupByItem, name) {
  var k = (typeof _feeKey_ === 'function') ? _feeKey_(name || '') : String(name || '');
  if (exclusiveGroupByItem) {
    if (typeof exclusiveGroupByItem.get === 'function') {
      var g = exclusiveGroupByItem.get(k);
      if (g) return g;
    } else if (exclusiveGroupByItem[k]) {
      return exclusiveGroupByItem[k];
    }
  }
  return ''; // pas mappé
}

function _isIgnoredLedgerRow_(r) {
  if (String(r.isIgnored || r.Ignore || r['Ignorer'] || '').toLowerCase() === 'true') return true;
  const tags = String(r.Tags || '').toLowerCase();
  if (/\bignore\b/.test(tags)) return true;
  const tm = String(r['Type de membre'] || r.typeMembre || '').toUpperCase();
  if (tm === 'ADULTE') return true;
  const ab = String(r.AgeBracket || r.ProgramBand || '').toUpperCase();
  if (ab === 'ADULTE') return true;
  return false;
}



function parseJSONSafe_(s) { try { return JSON.parse(String(s || '') || '[]'); } catch (_) { return []; } }

function NORM(s) { try { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase(); } catch (_) { return String(s || '').toUpperCase(); } }

// Récupère un nom d’item quel que soit le champ utilisé
function getName_(x) {
  if (!x || typeof x !== 'object') return '';
  var cand = x.NomFrais || x['Nom du frais'] || x.Frais || x.Produit || x.Item || x.Name || x.nom || x.name || '';
  return String(cand);
}

function tagsLower_(r) { return String(r['Tags'] || '').toLowerCase(); }

// U7 ou U8 (évite 17/18) — pour MapKey "u-07", "u-08"
const RE_MK_U7U8 = /\bu[-_ ]?0?(7|8)\b/i;

// "2e séance" (2e / 2eme / 2 ieme / deuxième) + "séance" (accents retirés par NORM)
function looksSecondSession_(X) { return /(2E|2EME|2\s*IEME|DEUXIEME)\s*(SEANCE|SEAN)\b/.test(X); }

function mkLower_(r) { return String(r['MapKey'] || '').toLowerCase(); }


// Status/isIgnored/Saison tolérants (clé insensible au casing / absente)
function get_(x, k) { if (!x) return undefined; return x[k] ?? x[k.toLowerCase()] ?? x[k.toUpperCase()]; }

function isActiveForSeason_(x, saisonLbl) {
  var st = Number(get_(x, 'Status') || get_(x, 'status') || 1);       // si absent → actif
  var ign = Number(get_(x, 'isIgnored') || get_(x, 'ignored') || 0);   // si absent → non ignoré
  var sx = get_(x, 'Saison');                                        // si absent → OK
  var okSeason = !sx || !saisonLbl || sx === saisonLbl;
  return (st === 1) && !ign && okSeason;
}

// Détection **lexicale**
function isU7U8SecteurName_(name) {
  var X = norm_(name);
  // U7-U8, U7/U8, U7U8, U7 – U8, etc.
  return /\bU\s*7\s*[-/ ]\s*U\s*8\b/.test(X) || /\bU7U8\b/.test(X);
}

function isU7U8SecondSessionName_(name) {
  var X = norm_(name);
  // “U7-U8” + “2e séance” (tolère 2E / 2EME / DEUXIEME)
  var hasU = (/\bU\s*7\s*[-/ ]\s*U\s*8\b/.test(X) || /\bU7U8\b/.test(X));
  var has2 = (/(^|\s)(2E|2EME|DEUXIEME)(\s+|[-/])SEANCE\b/.test(X));
  return hasU && has2;
}




/*************************************************
 *  RÈGLES — versions rapides (FULL & INCR)
 *  - S’appuie sur ACHATS_LEDGER + JOUEURS
 *  - Ecrit ERREURS en batch
 *************************************************/



/* =========================
 * Implémentation interne
 * ========================= */
/* =========================
 * Implémentation interne (FAST)
 * ========================= */
function _rulesBuildErrorsFast_(ss, touchedSet) {
  var saison = readParam_(ss, 'SEASON_LABEL') || '';

  // Helpers
  function NORM(s) {
    try { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase(); }
    catch (_) { return String(s || '').toUpperCase(); }
  }
  function uMaxFromBracket(br) {
    var m = String(br || '').match(/U\s*(\d+)\s*-\s*U?\s*(\d+)/i);
    if (m) return parseInt(m[2], 10);
    var m2 = String(br || '').match(/U\s*(\d+)/i);
    return m2 ? parseInt(m2[1], 10) : 0;
  }
  function uRangeFromBracket_(br) { // NEW: range {min,max} depuis AgeBracket
    var s = String(br || '');
    var m1 = s.match(/U\s*0?(\d+)\s*[-–]\s*U?\s*0?(\d+)/i);
    if (m1) return { min: parseInt(m1[1], 10), max: parseInt(m1[2], 10) };
    var m2 = s.match(/U\s*0?(\d+)/i);
    if (m2) { var x = parseInt(m2[1], 10); return { min: x, max: x }; }
    return null;
  }
  function _getExclusiveGroup(exclusive, key, name) {
    if (!exclusive) return '';
    if (typeof exclusive.get === 'function') {
      var g = exclusive.get(key);
      if (g) return g;
    }
    var fk = (typeof _feeKey_ === 'function') ? _feeKey_(name || '') : (name || '');
    return exclusive[key] || exclusive[fk] || exclusive[name] || '';
  }
  function mkLower_(r) { return String(r['MapKey'] || '').toLowerCase(); }
  function tagsLower_(r) { return String(r['Tags'] || '').toLowerCase(); }
  function looksSecondSession_(X) { // X = NORM(name)
    return /(2E|2EME|2\s*IEME|DEUXIEME)\s*(SEANCE|SEAN)\b/.test(X);
  }

  // === NEW: charger le catalogue MAPPINGS (AllowOrphan, Umin/Umax, etc.) ===
  // (exposé par rules.js → loadArticlesCatalog_(ss).match(name) renvoie {Umin,Umax,ExclusiveGroup,AllowOrphan,...})
  var catalog = (typeof loadArticlesCatalog_ === 'function') ? loadArticlesCatalog_(ss) : null; // :contentReference[oaicite:0]{index=0}

  // Data (1 lecture JOUEURS + 1 lecture LEDGER)
  var ledger = readSheetAsObjects_(ss.getId(), SHEETS.ACHATS_LEDGER);
  var joueurs = readSheetAsObjects_(ss.getId(), SHEETS.JOUEURS);


  // Compile rules (regex camp, exclusive map, etc.) — BESOIN AVANT l’agrégation
  var rules = _getCompiledRulesOrFallback_(ss) || {};

  // Index joueurs
  var jByPass = {};
  (joueurs.rows || []).forEach(function (r) {
    var p = String(r['Passeport #'] || r['Passeport'] || '').trim();
    if (!p) return;
    if (touchedSet && !touchedSet.has(p)) return;
    jByPass[p] = r;
  });

  // Agrégation ledger (actifs, bonne saison)
  // packs: { ART:{feeKey->{name,count}}, INSC:{...}, ANY_CAMP, HAS_U7U8_INSC, HAS_U7U8_SECOND, HAS_ADAPTE }
  var activeByPass = {};
  (ledger.rows || []).forEach(function (r) {
    if (r['Saison'] !== saison) return;
    var p = String(r['Passeport #'] || r['Passeport'] || '').trim(); if (!p) return;
    if (touchedSet && !touchedSet.has(p)) return;

    var status = Number(r['Status']) || 0;
    var ign = Number(r['isIgnored']) || 0;
    var type = r['Type'] === 'INSCRIPTION' ? 'INSC' : 'ART';
    var name = r['NomFrais'] || '';
    var k = (typeof _feeKey_ === 'function') ? _feeKey_(name) : String(name);

    if (!activeByPass[p]) {
      activeByPass[p] = {
        ART: {}, INSC: {},
        ANY_CAMP: false,
        HAS_U7U8_INSC: false,
        HAS_U7U8_SECOND: false,
        HAS_ADAPTE: false
      };
    }

    if (status === 1 && !ign) {
      var bucket = activeByPass[p][type];
      bucket[k] = bucket[k] || { name: name, count: 0 };
      bucket[k].count++;

      // ---- Détections directes depuis ACHATS_LEDGER ----
      var mk = mkLower_(r);     // ex: "u-07-masculin-saison-automne-hiver", "u7-u8-2e-seance-automne-hiver"
      var tg = tagsLower_(r);   // ex: "u4u8,inscription_normale"
      var Xn = NORM(name);

      // A) Inscription U7/U8 : MapKey contient u-07 ou u-08 (évite U17/U18)
      if (type === 'INSC' && (mk.indexOf('u-07') > -1 || mk.indexOf('u-08') > -1)) {
        activeByPass[p].HAS_U7U8_INSC = true;
      }

      // B) Adapté (via tag)
      if (type === 'INSC' && tg.indexOf('adapte') > -1) {
        activeByPass[p].HAS_ADAPTE = true;
      }

      // C) 2e séance U7-U8 : MapKey prioritaire, sinon fallback lexical
      if (type !== 'INSC') {
        if (/u-?7-?u-?8-2e-seance/.test(mk)) {
          activeByPass[p].HAS_U7U8_SECOND = true;
        } else {
          var hasU7orU8 = /\bU\s*0?7\b/.test(Xn) || /\bU\s*0?8\b/.test(Xn) || /\bU7U8\b/.test(Xn) || /\bU\s*7\s*[-/ ]\s*U?\s*8\b/.test(Xn);
          if (hasU7orU8 && looksSecondSession_(Xn)) {
            activeByPass[p].HAS_U7U8_SECOND = true;
          }
        }
      }

      // D) Camp : tags ou regex rules.rxCamp
      if (!activeByPass[p].ANY_CAMP) {
        if (tg.indexOf('camp_selection') > -1 || tg.indexOf('camp') > -1) {
          activeByPass[p].ANY_CAMP = true;
        } else if (rules.rxCamp && rules.rxCamp.test(name)) {
          activeByPass[p].ANY_CAMP = true;
        }
      }
    }
  });

  // Compile rules (regex camp, exclusive map, etc.)
  var rules = _getCompiledRulesOrFallback_(ss) || {};

  // Build erreurs
  var errors = [];
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  Object.keys(activeByPass).forEach(function (p) {
    var packs = activeByPass[p];
    var j = jByPass[p] || {};
    var nom = j['Nom'] || '';
    var prenom = j['Prénom'] || j['Prenom'] || '';
    var display = (prenom + ' ' + nom).trim();
    var saisonLbl = saison;

    // failsafe Adapté
    if (!packs.HAS_ADAPTE) {
      var ja = String(j['isAdapte'] || '').toLowerCase();
      if (ja === '1' || ja === 'true' || ja === 'oui') packs.HAS_ADAPTE = true;
    }

// --- (B) ARTICLE ORPHELIN — aucune INSCRIPTION active, sauf AllowOrphan ---
(function () {
  var hasInscription = Object.keys(packs.INSC).length > 0;
  if (hasInscription) return;
  if (!packs.ART || !catalog || typeof catalog.match !== 'function') return;

  // helper local: est-ce un "camp" ?
  function isCampName_(name) {
    if (!name) return false;
    if (rules && rules.rxCamp && rules.rxCamp.test(name)) return true;
    return /CAMP/i.test(String(name));
  }

  // calc âge joueur (réutilisé pour skip 13–18)
  function birthYearFromStr_(s) {
    var m = String(s || '').match(/(\d{4})/);
    return m ? parseInt(m[1], 10) : null;
  }
  var sy = parseSeasonYear_(saisonLbl) || (new Date()).getFullYear();
  var by = birthYearFromStr_(j['DateNaissance'] || j['Naissance'] || j['Année de naissance'] || j['Annee de naissance'] || '');
  var uExact = (by ? (sy - by) : null); // ex.: 2025 - 2012 = 13

  Object.keys(packs.ART).forEach(function (k) {
    var rec = packs.ART[k];           // { name, count }
    var match = catalog.match(rec.name);
    var allow = !!(match && match.AllowOrphan === true);

    // ⛔ IMPORTANT: si c'est un CAMP (et qu'on est en 13–18), on laisse le cas
    // être traité uniquement par la règle ciblée U13U18_CAMP_SEUL (plus bas).
    if (isCampName_(rec.name)) {
      var inU13U18 = false;
      if (uExact != null) {
        inU13U18 = (uExact >= 13 && uExact <= 18);
      } else {
        var rng = uRangeFromBracket_(j['AgeBracket'] || j['U'] || '');
        if (rng) {
          // si le range chevauche 13–18, on assume 13–18
          inU13U18 = !(rng.max < 13 || rng.min > 18);
        }
      }
      if (inU13U18) return; // 💡 skip: pas d’ARTICLE_ORPHELIN pour les camps 13–18
    }

    if (!allow) {
      errors.push(_errRow_({
        passeport: p, nom: nom, prenom: prenom, display: display,
        category: 'ARTICLES',
        code: 'ARTICLE_ORPHELIN',
        level: 'error',
        element: rec.name,
        message: 'Article sans inscription active',
        meta: { allowOrphan: allow, mapping: match || null },
        saison: saisonLbl, date: todayStr
      }));
    }
  });
})();


    // === (C) ÉLIGIBILITÉ Umin/Umax — priorité à l’âge exact (MAPPINGS) ===
    (function () {
      if (!packs.ART || !catalog || typeof catalog.match !== 'function') return;

      // 1) Essaie l’âge exact: year(saison) - year(naissance)
      function birthYearFromStr_(s) {
        var m = String(s || '').match(/(\d{4})/);
        return m ? parseInt(m[1], 10) : null;
      }
      var sy = parseSeasonYear_(saisonLbl) || (new Date()).getFullYear(); // helper dispo dans ta base
      var by = birthYearFromStr_(j['DateNaissance'] || j['Naissance'] || j['Année de naissance'] || j['Annee de naissance'] || '');
      var uExact = (by ? (sy - by) : null); // ex.: 2025 - 2013 = 12

      // 2) Fallback éventuel: range U depuis AgeBracket/U (si uExact indisponible)
      var membreRange = uExact != null
        ? { min: uExact, max: uExact }                 // priorité à l’âge exact
        : uRangeFromBracket_(j['AgeBracket'] || j['U'] || '');  // sinon range dérivé

      Object.keys(packs.ART).forEach(function (k) {
        var rec = packs.ART[k];
        var match = catalog.match(rec.name);
        if (!match) return; // pas mappé → on ne tranche pas

        var umin = (match.Umin != null) ? Number(match.Umin) : null;
        var umax = (match.Umax != null) ? Number(match.Umax) : null;
        if (umin == null && umax == null) return; // produit non borné → rien à faire

        // Si on n’a vraiment aucune info d’âge membre → prudence: on NE FLAG pas.
        if (!membreRange) return;

        // Intersection stricte produit↔membre (avec priorité uExact si disponible)
        var ok =
          (umin == null || membreRange.max >= umin) &&
          (umax == null || membreRange.min <= umax);

        if (!ok) {
          errors.push(_errRow_({
            passeport: p, nom: nom, prenom: prenom, display: display,
            category: 'ARTICLES',
            code: 'ARTICLE_INELIGIBLE_U',
            level: 'error',
            element: rec.name,
            message: 'Produit U' + (umin || '') + (umax ? ('–U' + umax) : '') +
              ' — membre ' + (uExact != null ? ('U' + uExact) : (j['AgeBracket'] || j['U'] || 'U?')),
            meta: { membreUExact: uExact, membreU: j['AgeBracket'] || j['U'] || '', Umin: umin, Umax: umax, mapping: match || null },
            saison: saisonLbl, date: todayStr
          }));
        }
      });
    })();

    // 1) DUPLICAT (ARTICLES)
    Object.keys(packs.ART).forEach(function (k) {
      var rec = packs.ART[k];
      if (rec.count > 1) {
        errors.push(_errRow_({
          passeport: p, nom: nom, prenom: prenom, display: display,
          category: 'ARTICLES', code: 'DUPLICAT', level: 'warn',
          element: rec.name, message: 'Article en double détecté',
          meta: { count: rec.count }, saison: saisonLbl, date: todayStr
        }));
      }
    });

    // 2) EXCLUSIVITE (articles exclusifs)
    var countsByGroup = {};
    Object.keys(packs.ART).forEach(function (k) {
      var rec = packs.ART[k];
      var grp = _getExclusiveGroup(rules.exclusiveGroupByItem, k, rec.name);
      if (!grp) return;
      countsByGroup[grp] = (countsByGroup[grp] || 0) + 1;
    });
    Object.keys(countsByGroup).forEach(function (grp) {
      if (countsByGroup[grp] > 1) {
        errors.push(_errRow_({
          passeport: p, nom: nom, prenom: prenom, display: display,
          category: 'ARTICLES', code: 'EXCLUSIVITE', level: 'error',
          element: grp, message: 'Conflit d’articles exclusifs (groupe: ' + grp + ')',
          meta: { group: grp, count: countsByGroup[grp] }, saison: saisonLbl, date: todayStr
        }));
      }
    });

    // 3) U13U18_CAMP_SEUL : camp actif mais aucune inscription active (13–18)
    var hasInscription = Object.keys(packs.INSC).length > 0;
    var hasCamp = packs.ANY_CAMP === true;

    if (hasCamp && !hasInscription) {
      var dna = j['DateNaissance'] || j['Naissance'] || '';
      var curY = parseSeasonYear_(saisonLbl) || (new Date()).getFullYear();
      var by = _extractBirthYearLoose_(dna);
      var age = by ? (curY - by) : 0;
      if (!age || (age >= 13 && age <= 18)) {
        errors.push(_errRow_({
          passeport: p, nom: nom, prenom: prenom, display: display,
          category: 'INSCRIPTIONS',
          code: 'U13U18_CAMP_SEUL',
          level: 'warn',
          element: 'Camp',
          message: 'Inscrit à un camp de sélection mais pas à la saison',
          meta: {}, saison: saisonLbl, date: todayStr
        }));
      }
    }

    // 4) U9–U12 SANS CDP (non-Adapté)
    (function () {
      var hasCdp = false;

      Object.keys(packs.ART).forEach(function (k) {
        var rec = packs.ART[k];
        var grp = _getExclusiveGroup(rules.exclusiveGroupByItem, k, rec.name);
        if (grp === 'CDP_ENTRAINEMENT') hasCdp = true;
        if (!grp) {
          var X = NORM(rec.name || '');
          if (/CDP/.test(X)) hasCdp = true;
        }
      });

      var isAdapte = packs.HAS_ADAPTE;
      var uMax = uMaxFromBracket(j.AgeBracket || '');
      var inU9U12 = (uMax >= 9 && uMax <= 12);

      if (inU9U12 && !isAdapte && !hasCdp) {
        errors.push(_errRow_({
          passeport: p, nom: nom, prenom: prenom, display: display,
          category: 'INSCRIPTIONS',
          code: 'U9_12_SANS_CDP',
          level: 'warn',
          element: '',
          message: 'U9–U12 sans CDP',
          meta: { U: j.AgeBracket || '' },
          saison: saisonLbl,
          date: todayStr
        }));
      }
    })();

    // 5) U7–U8 SANS 2e SÉANCE — basé sur MapKey (U-07 / U-08) + 2e séance MapKey
    (function () {
      if (!packs.HAS_U7U8_INSC) return;   // pas inscrit U7/8
      if (packs.HAS_ADAPTE) return;       // Adapté -> on ignore
      if (packs.HAS_U7U8_SECOND) return;  // a déjà la 2e séance

      errors.push(_errRow_({
        passeport: p, nom: nom, prenom: prenom, display: display,
        category: 'INSCRIPTIONS',
        code: 'U7_8_SANS_2E_SEANCE',
        level: 'warn',
        element: '',
        message: 'U7–U8 sans 2e séance',
        meta: { secteur: 'U7-U8' },
        saison: saisonLbl,
        date: todayStr
      }));
    })();

  }); // fin passeports

  // Log synthèse (unique)
  Logger.log(JSON.stringify({
    u7u8_sans_2e: errors.filter(e => e.code === 'U7_8_SANS_2E_SEANCE').length,
    u9_12_sans_cdp: errors.filter(e => e.code === 'U9_12_SANS_CDP').length,
    camp_seul: errors.filter(e => e.code === 'U13U18_CAMP_SEUL').length,
    article_orphelin: errors.filter(e => e.code === 'ARTICLE_ORPHELIN').length,
    article_ineligible_u: errors.filter(e => e.code === 'ARTICLE_INELIGIBLE_U').length
  }));

  return {
    header: _erreursHeader_(ss),
    errors: errors,
    ledgerCount: (ledger.rows || []).length,
    joueursCount: (joueurs.rows || []).length
  };
}



/** Public adapters (FAST) — exposés pour le serveur */
function runEvaluateRulesFast_(ssOpt) {
  var ss = ssOpt || getSeasonSpreadsheet_(getSeasonId_());
  // FULL: touchedSet = null
  return _rulesBuildErrorsFast_(ss, null);
}

function evaluateSeasonRulesIncrFast_(passports, ssOpt) {
  var ss = ssOpt || getSeasonSpreadsheet_(getSeasonId_());
  // INCR: touchedSet = Set<string> (8 chiffres)
  var set = new Set((passports || []).map(function (p) { return String(p || '').replace(/\D/g, '').padStart(8, '0'); }));
  return _rulesBuildErrorsFast_(ss, set);
}

/* Rendez-les accessibles via LIB et global (défensif) */
try {
  if (typeof LIB !== 'undefined' && LIB) {
    LIB.runEvaluateRulesFast_ = runEvaluateRulesFast_;
    LIB.evaluateSeasonRulesIncrFast_ = evaluateSeasonRulesIncrFast_;
    // expose aussi les privés si tu veux un fallback direct :
    LIB._rulesBuildErrorsFast_ = _rulesBuildErrorsFast_;
  } else {
    // exporte vers le global Apps Script (si pas de namespace LIB)
    this.runEvaluateRulesFast_ = runEvaluateRulesFast_;
    this.evaluateSeasonRulesIncrFast_ = evaluateSeasonRulesIncrFast_;
    this._rulesBuildErrorsFast_ = _rulesBuildErrorsFast_;
  }
} catch (_e) { /* no-op */ }


/* =========================
 *  E/S ERREURS — batch
 * ========================= */
function _rulesClearErreursSheet_(ss) {
  var sh = ss.getSheetByName('ERREURS') || ss.insertSheet('ERREURS');
  sh.clear(); // garde formatage; on réécrit un header cohérent
  appendImportLog_(ss, 'RULES_CLEAR_FULL', 'ERREURS reset (append=FALSE, no filter)');
}

function _rulesWriteFull_(ss, errs, header) {
  writeObjectsToSheet_(ss, 'ERREURS', errs, header);
}

function _rulesUpsertForPassports_(ss, newErrs, touchedSet, header) {
  var cur = readSheetAsObjects_(ss.getId(), 'ERREURS');
  var kept = (cur.rows || []).filter(function (r) {
    var p = String(r['Passeport #'] || r['Passeport'] || '').trim();
    return !touchedSet.has(p);
  });
  // rewrite + append
  writeObjectsToSheet_(ss, 'ERREURS', kept, cur.header || header);
  if (newErrs.length) appendObjectsToSheet_(ss, 'ERREURS', newErrs);
}

/* =========================
 *  Helpers erreurs & règles
 * ========================= */
function _erreursHeader_(ss) {
  // Essaie de réutiliser l’en-tête existant si présent
  var cur = readSheetAsObjects_(ss.getId(), 'ERREURS');
  if (cur && cur.header && cur.header.length) return cur.header;
  // Sinon header par défaut (adapter si besoin à ta feuille)
  return [
    'Passeport #', 'Nom', 'Prénom', 'Affichage', 'Catégorie', 'Code', 'Niveau',
    'Saison', 'Élément', 'Message', 'Meta', 'Date'
  ];
}

function _errRow_(o) {
  return {
    'Passeport #': o.passeport || '',
    'Nom': o.nom || '',
    'Prénom': o.prenom || '',
    'Affichage': o.display || '',
    'Catégorie': o.category || '',
    'Code': o.code || '',
    'Niveau': o.level || 'warn',
    'Saison': o.saison || '',
    'Élément': o.element || '',
    'Message': o.message || '',
    'Meta': JSON.stringify(o.meta || {}),
    'Date': o.date || ''
  };
}

// Tente de charger un ruleset précompilé si présent; sinon minimalistes
function _getCompiledRulesOrFallback_(ss) {
  var out = {
    ignoreFees: new Set(),
    rxCamp: _compileKeywordsToRegex_(readParam_(ss, PARAM_KEYS.RETRO_CAMP_KEYWORDS) || 'camp de selection u13,camp selection u13,camp u13'),
    exclusiveGroupByItem: new Map()
  };
  try {
    if (typeof loadRetroRulesFast_ === 'function') {
      var r = loadRetroRulesFast_(ss);
      if (r.ignoreFees) out.ignoreFees = r.ignoreFees;
      if (r.rxCamp) out.rxCamp = r.rxCamp;
      if (r.exclusiveGroupByItem) out.exclusiveGroupByItem = r.exclusiveGroupByItem;
      return out;
    }
  } catch (e) { }
  // Fallback exclusifs depuis MAPPING si tu as une fonction
  try {
    if (typeof loadExclusiveGroupsMapping_ === 'function') {
      var mapping = loadExclusiveGroupsMapping_(ss); // [{group:'CDP_ENTRAINEMENT', items:[...]}]
      mapping.forEach(function (g) {
        (g.items || []).forEach(function (name) {
          out.exclusiveGroupByItem.set(_feeKey_(name), g.group);
        });
      });
    }
  } catch (e) { }
  return out;
}


// Alias pour compatibilité : certaines branches appellent refreshJoueursPhotoStr_,
// or la vraie implémentation s’appelle refreshPhotoStrInJoueurs_.
function refreshJoueursPhotoStr_(ssOrId) {
  return refreshPhotoStrInJoueurs_(ssOrId);
}
