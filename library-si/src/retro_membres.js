/**
 * retro_members.gs — v0.7
 * Export "Rétro - Membres" compatible avec l’ancien format, mais:
 *  - lit les feuilles finals v0.7 (SHEETS.INSCRIPTIONS / SHEETS.ARTICLES)
 *  - respecte CANCELLED/EXCLUDE_FROM_EXPORT
 *  - ignore certains frais (RETRO_IGNORE_FEES_CSV)
 *  - détecte Soccer Adapté (inscriptions ET articles) => "Adapté"=1
 *  - calcule CDP via MAPPINGS (ExclusiveGroup="CDP_ENTRAINEMENT") + fallback par textes
 *  - détecte "camp sélection U13" (paramétrable)
 *  - "Muté" via une feuille de mutation configurable
 *  - (optionnel) colonne Photo selon date d’expiration + cutoff paramétrés
 *
 * retro_members.gs — v0.11 (incrémental)
* - Garde la logique de build existante (buildRetroMembresRows)
* - ✨ Ajoute le filtrage des exports par passeports «touchés»
* - options.onlyPassports (Array/Set)
* - à défaut: DocumentProperties.LAST_TOUCHED_PASSPORTS (CSV)
* - Force la col. A (Passeport) en texte et normalise en 8 chars si helpers dispo
* - Suffixe le nom de fichier par _INCR si filtrage appliqué

 * Fonctions exposées:
 *  - buildRetroMembresRows(seasonSheetId) -> {header:[], rows:[[]], nbCols:int}
 *  - writeRetroMembresSheet(seasonSheetId) -> écrit "Rétro - Membres"
 *  - exportRetroMembresXlsxToDrive(seasonSheetId) -> XLSX dans dossier paramétré
 */

/* ===================== Param keys (nouvelles) ===================== */
if (typeof PARAM_KEYS === 'undefined') { var PARAM_KEYS = {}; }
PARAM_KEYS.RETRO_IGNORE_FEES_CSV = PARAM_KEYS.RETRO_IGNORE_FEES_CSV || 'RETRO_IGNORE_FEES_CSV';
PARAM_KEYS.RETRO_ADAPTE_KEYWORDS = PARAM_KEYS.RETRO_ADAPTE_KEYWORDS || 'RETRO_ADAPTE_KEYWORDS';
PARAM_KEYS.RETRO_CAMP_KEYWORDS = PARAM_KEYS.RETRO_CAMP_KEYWORDS || 'RETRO_CAMP_KEYWORDS';
PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID = PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID || 'RETRO_EXPORTS_FOLDER_ID';
PARAM_KEYS.RETRO_MUTATION_SHEET = PARAM_KEYS.RETRO_MUTATION_SHEET || 'RETRO_MUTATION_SHEET';
PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL = PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL || 'RETRO_PHOTO_INCLUDE_COL';
PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL = PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL || 'RETRO_PHOTO_EXPIRY_COL';
PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE = PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE || 'RETRO_PHOTO_WARN_ABS_DATE';
PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD = PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD || 'RETRO_PHOTO_WARN_BEFORE_MMDD';
PARAM_KEYS.RETRO_EXCLUDE_MEMBER_IF_ONLY_IGN = PARAM_KEYS.RETRO_EXCLUDE_MEMBER_IF_ONLY_IGN || 'RETRO_EXCLUDE_MEMBER_IF_ONLY_IGN';
PARAM_KEYS.RETRO_RULES_JSON = PARAM_KEYS.RETRO_RULES_JSON || 'RETRO_RULES_JSON';



// Cache global partagé (idempotent)
var __retroRulesCache = (typeof __retroRulesCache !== 'undefined')
  ? __retroRulesCache
  : { at: 0, data: null };

function loadRetroRules_(ss) {
  // 1) Si la version centralisée est dispo (server_rules.js), on l’utilise.
  if (typeof SR_loadRetroRules_ === 'function') {
    return SR_loadRetroRules_(ss);
  }

  // 2) Fallback local avec cache 5 minutes
  var now = Date.now();
  if (__retroRulesCache.data && (now - __retroRulesCache.at) < 5 * 60 * 1000) {
    return __retroRulesCache.data;
  }

  // PARAM direct
  var raw = readParam_(ss, PARAM_KEYS.RETRO_RULES_JSON) || '';

  // Feuille "RETRO_RULES_JSON" si vide
  if (!raw) {
    var shJson = ss.getSheetByName('RETRO_RULES_JSON');
    if (shJson && shJson.getLastRow() >= 1 && shJson.getLastColumn() >= 1) {
      var vals = shJson.getDataRange().getDisplayValues();
      var pieces = [];
      for (var i = 0; i < vals.length; i++) {
        for (var j = 0; j < vals[i].length; j++) {
          var cell = vals[i][j];
          if (cell != null && String(cell).trim() !== '') pieces.push(String(cell));
        }
      }
      raw = pieces.join('\n');
      appendImportLog_(ss, 'RETRO_RULES_JSON_SHEET_READ', 'chars=' + raw.length);
    }
  }

  // Parse JSON
  var rules = [];
  if (raw) {
    try {
      var arr = JSON.parse(raw);
      rules = Array.isArray(arr) ? arr : [];
    } catch (e) {
      appendImportLog_(ss, 'RETRO_RULES_JSON_PARSE_FAIL', String(e));
    }
  }

  // Defaults si rien
  if (!rules.length) {
    var ignoreCsv = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV) || 'senior,u-sé,adulte,ligue';
    var adapteCsv = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || 'adapté,adapte';
    var campCsv = readParam_(ss, PARAM_KEYS.RETRO_CAMP_KEYWORDS) || 'camp de sélection u13,camp selection u13,camp u13';
    var photoOn = (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) || 'FALSE').toUpperCase() === 'TRUE';
    var photoCol = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL) || '';
    var warnMmDd = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD) || '03-01';
    var absDate = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';

    rules = [
      {
        id: 'ignore_fees', enabled: true, scope: 'both',
        when: { field: 'Nom du frais', contains_any: ignoreCsv.split(',') },
        action: { type: 'ignore_row' }
      },
      {
        id: 'adapte_flag', enabled: true, scope: 'both',
        when: { field: 'Nom du frais', contains_any: adapteCsv.split(',') },
        action: { type: 'set_member_field', field: 'adapte', value: 1 }
      },
      {
        id: 'cdp_2', enabled: true, scope: 'articles',
        when: { catalog_exclusive_group: 'CDP_ENTRAINEMENT', text_contains_any: ['2', '2 entrainements'] },
        action: { type: 'set_member_field_max', field: 'cdp', value: 2 }
      },
      {
        id: 'cdp_1', enabled: true, scope: 'articles',
        when: { catalog_exclusive_group: 'CDP_ENTRAINEMENT' },
        action: { type: 'set_member_field_max', field: 'cdp', value: 1 }
      },
      {
        id: 'camp_u13', enabled: true, scope: 'articles',
        when: { field: 'Nom du frais', contains_any: campCsv.split(',') },
        action: { type: 'set_member_field', field: 'camp', value: 'Oui' }
      }
    ];
    if (photoOn && photoCol) {
      rules.push({
        id: 'photo_policy', enabled: true, scope: 'member',
        action: { type: 'compute_photo', expiry_col: photoCol, warn_mmdd: warnMmDd, abs_date: absDate }
      });
    }
    appendImportLog_(ss, 'RETRO_RULES_JSON_FALLBACK', 'using PARAMS-derived defaults');
  }

  __retroRulesCache = { at: now, data: rules };
  return rules;
}


/* ===================== Helpers bas niveau ===================== */

if (typeof CONTROL_COLS === 'undefined') {
  var CONTROL_COLS = { ROW_HASH: 'ROW_HASH', CANCELLED: 'CANCELLED', EXCLUDE_FROM_EXPORT: 'EXCLUDE_FROM_EXPORT', LAST_MODIFIED_AT: 'LAST_MODIFIED_AT' };
}

// ——— Birth year robuste: gère Date, ISO, FR, tout ce qui contient un yyyy ———
function _extractBirthYearLoose_(v) {
  if (!v) return 0;
  if (v instanceof Date) return v.getFullYear();
  var s = String(v).trim();
  // Cherche un groupe de 4 chiffres entre 1900 et 2099
  var m = s.match(/(19\d{2}|20\d{2})/);
  if (m) {
    var y = parseInt(m[1], 10);
    if (y >= 1900 && y <= 2099) return y;
  }
  return 0;
}

// --- Photo: constantes + helpers communs
const PHOTO_NOT_REQUIRED = 'Non requise';

function _rm_truthy_(v) {
  var s = String(v || '').trim().toUpperCase();
  return (s === '1' || s === 'TRUE' || s === 'OUI' || s === 'YES');
}
function _rm_absWarnDate_(ss) {
  var raw = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';
  if (!raw) return null;
  var d = new Date(raw);
  return isNaN(+d) ? null : d;
}
function _rm_needPhotoByFields_(age, isAdapte, hasInscription) {
  var a = parseInt(String(age || ''), 10);
  if (!isNaN(a) && a < 8) return false;
  if (_rm_truthy_(isAdapte)) return false;
  if (!_rm_truthy_(hasInscription)) return false;
  return true;
}
function _rm_statusFromFields_(ss, expDate, age, isAdapte, hasInscription) {
  if (!_rm_needPhotoByFields_(age, isAdapte, hasInscription)) return PHOTO_NOT_REQUIRED;
  if (!expDate) return 'Aucune photo';
  var d = (expDate instanceof Date) ? expDate : new Date(expDate);
  if (isNaN(+d)) return 'Aucune photo';
  var abs = _rm_absWarnDate_(ss);
  if (abs && d < abs) return 'Expirée';
  return 'Valide';
}


// --- AJOUT: helper strict, indépendant du vieux _computePhotoCell_ ---
function _rm_getAbsWarnDate_(ss) {
  var raw = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';
  if (!raw) return null;
  try {
    var d = new Date(raw);
    return isNaN(+d) ? null : d;
  } catch (e) { return null; }
}

function _rm_needsPhoto_(age, isAdapte, hasInscription) {
  var a = parseInt(String(age || ''), 10);
  var u8 = (!isNaN(a) ? a : 0) >= 8;
  var ada = (String(isAdapte || '').trim() === '1' || String(isAdapte || '').toLowerCase() === 'true');
  return hasInscription && u8 && !ada;
}

function _rm_photoStrStrict_(ss, photoExpireLe, age, isAdapte, hasInscription) {
  // 1) Besoin?
  if (!_rm_needsPhoto_(age, isAdapte, hasInscription)) {
    // Exigence: champ VIDE pour non-inscrits; et plus largement on préfère "vide" quand non requis.
    return '';
  }
  // 2) A-t-il une photo?
  if (!photoExpireLe) return 'Aucune photo';

  // 3) Expirée vs ABS cutoff
  var abs = _rm_getAbsWarnDate_(ss);
  try {
    var exp = (photoExpireLe instanceof Date) ? photoExpireLe : new Date(photoExpireLe);
    if (isNaN(+exp)) return 'Aucune photo';
    if (abs && exp < abs) return 'Expirée';
    return 'Valide';
  } catch (e) {
    return 'Aucune photo';
  }
}


function _nrm_(s) {
  s = String(s == null ? '' : s);
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { }
  return s;
}
function _nrmLower_(s) { return _nrm_(s).toLowerCase().trim(); }
function _csvEsc_(v) { v = v == null ? '' : String(v).replace(/"/g, '""'); return /[",\n;]/.test(v) ? ('"' + v + '"') : v; }

function _isActiveRow_(r) {
  var can = String(r[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true';
  var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT] || '').toLowerCase() === 'true';
  var st = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
  return !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
}
function _feeIgnored_(name, patternsCsv) {
  var raw = _nrmLower_(name || '');
  if (!raw) return false;
  var pats = String(patternsCsv || '').split(',').map(function (x) { return _nrmLower_(x); }).filter(Boolean);
  for (var i = 0; i < pats.length; i++) { if (raw.indexOf(pats[i]) !== -1) return true; }
  return false;
}
function _containsAny_(raw, csv) {
  var s = _nrmLower_(raw || '');
  return String(csv || '').split(',').map(function (x) { return _nrmLower_(x); }).filter(Boolean).some(function (p) { return s.indexOf(p) !== -1; });
}
function _safeDate_(v) {
  if (!v) return null;
  try { return (v instanceof Date) ? v : new Date(v); } catch (e) { return null; }
}
function _yyyy_mm_dd_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}


/* ===================== Lecture "mutation" ===================== */
function _loadMutationsSet_(ss) {
  var sheetName = readParam_(ss, PARAM_KEYS.RETRO_MUTATION_SHEET) || 'Mutation';
  var sh = ss.getSheetByName(sheetName);
  var set = {};
  if (!sh || sh.getLastRow() < 2) return set;
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (var i = 0; i < vals.length; i++) {
    var p = (vals[i][0] || '').toString().trim();
    if (p) set[p] = true;
  }
  return set;
}

/* ===================== PHOTO logic (optionnel) ===================== */
function _computePhotoCell_(ss, row) {
  var include = (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) || 'FALSE').toUpperCase() === 'TRUE';
  if (!include) return ''; // reste 100% compatible par défaut

  var col = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL) || '';
  if (!col || !(col in row)) return '';

  var exp = _safeDate_(row[col]);
  if (!exp) return 'Aucune photo';
  var today = new Date();
  if (exp < today) return 'Expirée (' + _yyyy_mm_dd_(exp) + ')';

  // alerte "Expire bientôt"
  var absWarn = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';
  var warnMmDd = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD) || '03-01';

  var saisonLabel = readParam_(ss, 'SEASON_LABEL') || (row['Saison'] || '');
  var seasonYear = parseSeasonYear_(saisonLabel);

  if (absWarn) {
    var abs = _safeDate_(absWarn);
    if (abs && exp <= abs) return 'Expire bientôt (' + _yyyy_mm_dd_(exp) + ')';
  }
  if (exp.getFullYear() === seasonYear) {
    // règle: expiration durant l’année => changement avant la saison estivale
    return 'Expire bientôt (' + _yyyy_mm_dd_(exp) + ')';
  }
  return _yyyy_mm_dd_(exp); // simple date
}

/* ===================== Retro Rules — loader + exécution ===================== */

/** Charge la config de règles:
 *  1) PARAMS.RETRO_RULES_JSON (JSON inline)
 *  2) onglet feuille nommé "RETRO_RULES_JSON" (JSON dans A1 ou multi-lignes concaténées)
 *  (fallback) dérive depuis d'autres PARAMS si rien n’est trouvé
 */


function _r_normLower_(s) { return _nrmLower_(s); }
function _r_getFieldText_(row, field) {
  if (!field) return '';
  var v = row[field];
  if (v == null) return '';
  return String(v);
}
function _r_matchWhen_(when, row, feeName, catalogItem) {
  if (!when) return true;
  if (when.field && when.contains_any) {
    var txt = _r_getFieldText_(row, when.field);
    var low = _r_normLower_(txt);
    var arr = [].concat(when.contains_any || []).map(_r_normLower_).filter(Boolean);
    if (!arr.some(function (s) { return low.indexOf(s) !== -1; })) return false;
  }
  if (when.text_contains_any) {
    var low = _r_normLower_(feeName || '');
    var arr = [].concat(when.text_contains_any || []).map(_r_normLower_).filter(Boolean);
    if (!arr.some(function (s) { return low.indexOf(s) !== -1; })) return false;
  }
  if (when.catalog_exclusive_group) {
    if (!catalogItem) return false;
    if (String(catalogItem.ExclusiveGroup || '') !== String(when.catalog_exclusive_group)) return false;
  }
  return true;
}

/** Applique toutes les règles "row" sur une ligne source (inscriptions/articles).
 *  Retourne { skip:true } si l’action "ignore_row" a été tirée.
 *  Peut modifier l’objet member (adapte/cdp/camp…).
 */
function applyRetroRowRules_(rules, scope, row, member, ctx) {
  var feeName = (row['Nom du frais'] || row['Frais'] || row['Produit'] || '');
  var item = ctx.catalog.match ? ctx.catalog.match(feeName) : null;
  var skip = false;

  rules.forEach(function (rule) {
    if (!rule || !rule.enabled) return;
    if (!(rule.scope === 'both' || rule.scope === scope)) return;
    if (!_r_matchWhen_(rule.when, row, feeName, item)) return;

    var a = rule.action || {};
    if (a.type === 'ignore_row') {
      skip = true;
    } else if (a.type === 'set_member_field') {
      member[a.field] = a.value;
    } else if (a.type === 'set_member_field_max') {
      var cur = member[a.field];
      if (cur == null || cur === '') member[a.field] = a.value;
      else {
        var n1 = parseFloat(cur), n2 = parseFloat(a.value);
        if (!isNaN(n1) && !isNaN(n2)) member[a.field] = Math.max(n1, n2);
        else member[a.field] = a.value; // fallback
      }
    }
  });

  return { skip: skip };
}

/** Applique les règles "member" (post-agrégation).
 *  Peut écrire member.__photoStr pour que la feuille affiche la valeur.
 */
function applyRetroMemberRules_(rules, member, ctx) {
  rules.forEach(function (rule) {
    if (!rule || !rule.enabled) return;
    if (rule.scope !== 'member') return;
    var a = rule.action || {};
    if (a.type === 'compute_photo') {
      var tmpRow = {};
      tmpRow[a.expiry_col] = (member.__rowForPhoto && member.__rowForPhoto[a.expiry_col]) ? member.__rowForPhoto[a.expiry_col] : '';
      member.__photoStr = _computePhotoCell_(ctx.ss, tmpRow);
    }
  });
}

/* ===================== Construire les lignes "Rétro - Membres" ===================== */
function buildRetroMembresRows(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var art  = readSheetAsObjects_(ss.getId(), SHEETS.ARTICLES);

  // === MEMBRES_GLOBAL (prioritaire pour la Photo)
  var shMemName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBRES_GLOBAL) ? SHEETS.MEMBRES_GLOBAL : 'MEMBRES_GLOBAL';
  var mem = readSheetAsObjects_(ss.getId(), shMemName);
  var indexMemByPassport = {};
  if (mem && mem.rows && mem.rows.length) {
    var pCol = 'Passeport #';
    var Hmem = mem.header || [];
    if (Hmem.indexOf('Passeport #') < 0 && Hmem.indexOf('Passeport') >= 0) pCol = 'Passeport';
    mem.rows.forEach(function (r) {
      var p = (r[pCol] != null) ? String(r[pCol]).trim() : '';
      if (p) indexMemByPassport[p] = r;
    });
  }

  // Params
  var ignoreFees  = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV) || 'senior,u-se,adulte,ligue';
  var adapteKeys  = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || 'adapté,adapte,adapte';
  var campKeys    = readParam_(ss, PARAM_KEYS.RETRO_CAMP_KEYWORDS) || 'camp de selection u13,camp selection u13,camp u13';
  var includePhoto = (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) || 'FALSE').toUpperCase() === 'TRUE';

  // Map membres
  var members = {}; // { passeport, nom, prenom, dateNaissance, genre, emails[], adapte, cdp, camp, inscription:boolean }
  function ensureMember_(p, seed) {
    var k = String(p || '').trim(); if (!k) return null;
    if (!members[k]) {
      members[k] = {
        passeport: k, nom: '', prenom: '', dateNaissance: '', genre: '',
        emails: [], adapte: undefined, cdp: undefined, camp: undefined, inscription: false
      };
    }
    if (seed) {
      var m = members[k];
      if (!m.nom && seed.nom) m.nom = seed.nom;
      if (!m.prenom && seed.prenom) m.prenom = seed.prenom;
      if (!m.dateNaissance && seed.dateNaissance) m.dateNaissance = seed.dateNaissance;
      if (!m.genre && seed.genre) m.genre = seed.genre;
    }
    return members[k];
  }

  // --- Clés et drapeaux partagés (portée: toute la fonction)
  function RM_key_(pass) {
    var p = String(pass || '').trim();
    if (typeof normalizePassportToText8_ === 'function') {
      try { var p8 = normalizePassportToText8_(p); if (p8) return p8; } catch(_) {}
    }
    return p;
  }

  var hasU9U12Inscription = Object.create(null); // clé normalisée -> true si inscription U9-U12 (réelle)
  var isEliteInscription  = Object.create(null); // clé normalisée -> true si inscription élite

  var ELITE_KEYWORDS = ['D1+','D1','PLSJQ','LDP','ÉLITE','ELITE','SÉLECTION','SELECTION'];

  function _rm_extractU_(s) {
    var m = String(s||'').match(/\bU\s*-?\s*(\d{1,2})\b/i);
    return m ? parseInt(m[1],10) : 0;
  }
  function _rm_isEliteByName_(feeName){
    var n = String(feeName||'').toUpperCase();
    return ELITE_KEYWORDS.some(function(k){ return n.indexOf(k) >= 0; });
  }
  function _rm_hasEliteTag_(row){
    var t = row && (row['Tags'] || row['Tag'] || '');
    if (Array.isArray(t)) t = t.join(',');
    try {
      if (typeof t === 'string' && t.trim().startsWith('[')) t = JSON.parse(t).join(',');
    } catch(_) {}
    return String(t||'').toLowerCase().indexOf('inscription_elite') >= 0;
  }

  // --- INSCRIPTIONS (actives selon _isActiveRow_)
  var inscAct = insc.rows.filter(_isActiveRow_);
  inscAct.forEach(function (r) {
    var pass = r['Passeport #']; if (!pass) return;
    var feeName = r['Nom du frais'] || r['Frais'] || r['Produit'] || '';
    if (_feeIgnored_(feeName, ignoreFees)) return;

    var k = RM_key_(pass);

    // Drapeaux élite + U9-U12 basés sur l'INSCRIPTION (pas les articles)
    if (_rm_hasEliteTag_(r) || _rm_isEliteByName_(feeName)) {
      isEliteInscription[k] = true;
    }
    var uNum = _rm_extractU_(feeName);
    if (uNum >= 9 && uNum <= 12) {
      hasU9U12Inscription[k] = true;
    }

    var m = ensureMember_(pass, {
      nom: r['Nom'] || '',
      prenom: r['Prénom'] || r['Prenom'] || '',
      dateNaissance: r['Date de naissance'] || r['Naissance'] || '',
      genre: (r['Identité de genre'] || '').toString().trim().toUpperCase().charAt(0)
    });
    if (!m) return;
    m.inscription = true;

    // e-mails via INSCRIPTIONS
    var emails = (typeof collectEmailsFromRow_ === 'function')
      ? collectEmailsFromRow_(r, readParam_(ss, 'TO_FIELDS_INSCRIPTIONS') || 'Courriel,Parent 1 - Courriel,Parent 2 - Courriel')
      : [r['Courriel'], r['Parent 1 - Courriel'], r['Parent 2 - Courriel']].filter(Boolean).join(',');
    if (emails) emails.split(',').forEach(function (e) { e = e.trim(); if (e && members[pass].emails.indexOf(e) === -1) members[pass].emails.push(e); });

    // Adapté si mot-clef DANS L’INSCRIPTION (pas dans articles)
    if (_containsAny_(feeName, adapteKeys)) m.adapte = 1;
  });

  // --- ARTICLES (actifs selon _isActiveRow_)
  var artAct = art.rows.filter(_isActiveRow_);
  artAct.forEach(function (a) {
    var pass = a['Passeport #']; if (!pass) return;

    var feeName = a['Nom du frais'] || a['Frais'] || a['Produit'] || '';
    if (_feeIgnored_(feeName, ignoreFees)) return;

    var m = ensureMember_(pass, {
      nom: a['Nom'] || '',
      prenom: a['Prénom'] || a['Prenom'] || '',
      dateNaissance: a['Date de naissance'] || a['Naissance'] || '',
      genre: (a['Identité de genre'] || '').toString().trim().toUpperCase().charAt(0)
    });
    if (!m) return;

    // Fallback courriel via ARTICLES si rien via INSCRIPTIONS
    if (m.emails.length === 0 && a['Courriel']) {
      var ea = String(a['Courriel']).trim();
      if (ea) m.emails.push(ea);
    }

    // CDP par heuristique texte (si on n’a rien de centralisé)
    var s = _nrmLower_(feeName);
    var isCdp = s.indexOf('cdp') !== -1;
    if (isCdp) {
      if (/\b2\b/.test(s) || /2\s*entrainement/.test(s) || /2\s*entrainements/.test(s)) {
        m.cdp = 2;
      } else if (/\b1\b/.test(s) || /1\s*entrainement/.test(s) || /1\s*entrainements/.test(s)) {
        m.cdp = Math.max(m.cdp || 0, 1);
      } else {
        m.cdp = Math.max(m.cdp || 0, 1);
      }
    }

    // Camp (clé paramétrable)
    if (_containsAny_(feeName, campKeys)) m.camp = 'Oui';

    // NOTE: pas de détection "adapté" via ARTICLES (inexistant par design)
  });

  // U9–U12 => CDP=0 **uniquement** si: inscription U9-U12 active, non-élite, non-adapté, et aucun CDP acheté
  Object.keys(members).forEach(function(k){
    var m = members[k];
    var key = RM_key_(m.passeport || k);

    var isAdapt = (m.adapte === 1) || (m.adapte === '1') || (String(m.adapte||'').toLowerCase() === 'true');
    var hasU9U12 = !!hasU9U12Inscription[key];
    var isElite  = !!isEliteInscription[key];

    if (m.inscription && hasU9U12 && !isElite && !isAdapt) {
      if (m.cdp === undefined || m.cdp === null || m.cdp === '') m.cdp = 0;
    }
  });

  // Index INSCRIPTIONS pour fallback Photo
  var indexByPassport = {};
  inscAct.forEach(function (r) { indexByPassport[String(r['Passeport #'] || '').trim()] = r; });

  // Photo: priorité MEMBRES_GLOBAL, fallback INSCRIPTIONS
  Object.keys(members).forEach(function (k) {
    var m = members[k];
    m.__rowForPhoto = indexMemByPassport[m.passeport] || indexByPassport[m.passeport] || {};
    // plus de applyRetroMemberRules_ ici; on calculera Photo à l’écriture
  });

  // --- JOUEURS: index par passeport pour PhotoStr (prioritaire)
  var jTab = readSheetAsObjects_(ss.getId(), SHEETS.JOUEURS);
  var J = jTab.rows || [], HJ = jTab.header || [];
  var J_BY_P8 = {};
  for (var i = 0; i < J.length; i++) {
    var r = J[i] || {};
    var p8 = _normP8_(r['Passeport #'] || r['Passeport'] || '');
    if (p8) J_BY_P8[p8] = r;
  }

  // ——— Construction des lignes (header complet, peu de colonnes remplies) ———
  var HEADER = [
    "Identifiant unique", "Code", "Nom", "Prénom", "Date de naissance",
    "Genre(M pour Masculin ou F pour Féminin)", "Langue", "Courriels", "Adresse", "Ville",
    "Code Postal", "Domicile Téléphone", "Mobile Téléphone", "Travail Téléphone",
    "Parent 1 Nom", "Parent 1 Courriels", "Parent 1 Domicile Téléphone", "Parent 1 Mobile Téléphone", "Parent 1 Travail Téléphone",
    "Parent 2 Nom", "Parent 2 Courriels", "Parent 2 Domicile Téléphone", "Parent 2 Mobile Téléphone", "Parent 2 Travail Téléphonique",
    "Autre Nom", "Autre Courriels", "Autre Domicile Téléphone", "Autre Mobile Téléphone", "Autre Travail Téléphonique",
    "Position", "Établissement scolaire", "Fiche d'employé", "Specimen Chèque", "Filtration Policière", "Respect et Sport",
    "S3", "S7", "Théorie A+B", "Diplôme C", "Adapté", "CDP", "Euroclass", "Camp", "Muté", "École", "InscritE25"
  ];
  var header = HEADER.slice();
  if (includePhoto) header.push('Photo');

  var rows = [];
  Object.keys(members).forEach(function (k) {
    var m = members[k];
    var row = new Array(header.length); for (var i = 0; i < row.length; i++) row[i] = '';

    row[0] = (typeof normalizePassportToText8_ === 'function')
      ? normalizePassportToText8_(m.passeport)
      : String(m.passeport || '');
    // row[1] Code -> vide
    row[2] = m.nom;
    row[3] = m.prenom;
    row[4] = m.dateNaissance;
    row[5] = m.genre;
    row[7] = (m.emails && m.emails.length) ? m.emails.join('; ') : '';

    // Adapté, CDP, Camp
    row[39] = (typeof m.adapte !== 'undefined' ? m.adapte : '');
    row[40] = (typeof m.cdp !== 'undefined' ? m.cdp : '');
    row[42] = (typeof m.camp !== 'undefined' ? m.camp : '');

    // Photo (optionnelle)
    if (includePhoto) {
      var wrote = false;
      var p8 = _normP8_(m.passeport);
      var jRow = J_BY_P8[p8] || null;

      // 1) PhotoStr direct depuis JOUEURS (si présent)
      if (jRow && String(jRow['PhotoStr'] || '').trim() !== '') {
        row[header.length - 1] = String(jRow['PhotoStr']).trim();
        wrote = true;
      }

      // 2) Sinon, calcule proprement via JOUEURS (sans date brute)
      if (!wrote && jRow) {
        var ageJ = jRow['Age'] || '';
        var adaJ = jRow['isAdapte'] || '';
        var hasInJ = jRow['hasInscription'] || ''; // TRUE/FALSE ou Oui/Non — _rm_truthy_ gère
        var expJ = jRow['PhotoExpireLe'] || '';
        row[header.length - 1] = _rm_statusFromFields_(ss, expJ, ageJ, adaJ, hasInJ);
        wrote = true;
      }

      // 3) Sinon, fallback MEMBRES_GLOBAL/INSCRIPTIONS (strict, pas de dates)
      if (!wrote) {
        var saisonYear = parseSeasonYear_(readParam_(ss, 'SEASON_LABEL') || '');
        var by = _extractBirthYearLoose_(m.dateNaissance);
        var age = (by && saisonYear) ? (saisonYear - by) : '';
        var isAdapte = (m.adapte === 1 || m.adapte === '1' || String(m.adapte || '').toLowerCase() === 'true');
        var hasInscription = !!m.inscription;

        var expDate = '';
        if (m.__rowForPhoto) {
          expDate = m.__rowForPhoto['PhotoExpireLe'] || m.__rowForPhoto['Photo Expire Le'] || '';
        }
        row[header.length - 1] = _rm_statusFromFields_(ss, expDate, age, isAdapte, hasInscription);
        wrote = true;
      }
    }

    rows.push(row);
  });

  return { header: header, rows: rows, nbCols: header.length };
}



/* ===================== Filtrage incrémental (touchés) ===================== */
function _rm_norm_passport_(v) { var s = String(v == null ? '' : v).trim(); try { if (typeof normalizePassportToText8_ === 'function') return normalizePassportToText8_(s); } catch (_) { } try { if (typeof normalizePassportPlain8_ === 'function') return normalizePassportPlain8_(s); } catch (_) { } return s; }
function _rm_readTouchedPassportSet_(options) {
  options = options || {};
  var set = {};

  // 1) options.onlyPassports
  var list = options.onlyPassports;
  if (list && typeof list.forEach === 'function') {
    list.forEach(function (p) { var t = _rm_norm_passport_(p); if (t) set[t] = true; });
  }

  // 2) Fallback: DocumentProperties.LAST_TOUCHED_PASSPORTS (JSON ou CSV)
  if (!Object.keys(set).length) {
    try {
      var raw = (PropertiesService.getDocumentProperties().getProperty('LAST_TOUCHED_PASSPORTS') || '').trim();
      if (raw) {
        var arr = (raw.charAt(0) === '[') ? JSON.parse(raw) : raw.split(',');
        arr.forEach(function (p) { var t = _rm_norm_passport_(p); if (t) set[t] = true; });
      }
    } catch (_) { }
  }
  return set;
}


function _rm_filterRowsByPassports_(rows, touchedSet) { var keys = Object.keys(touchedSet || {}); if (!keys.length) return rows; var out = []; for (var i = 0; i < rows.length; i++) { var row = rows[i]; var p = _rm_norm_passport_(row && row[0]); if (p && touchedSet[p]) out.push(row); } return out; }

/* ===================== Écriture feuille & export XLSX ===================== */

function writeRetroMembresSheet(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var res = buildRetroMembresRows(seasonSheetId);
  var sh = ss.getSheetByName('Rétro - Membres') || ss.insertSheet('Rétro - Membres');
  sh.clearContents();

  sh.getRange(1, 1, 1, res.header.length).setValues([res.header]);

  if (res.rows.length) {
    sh.getRange(2, 1, res.rows.length, res.nbCols).setValues(res.rows);
    sh.autoResizeColumns(1, res.nbCols);
    // Passeport en texte
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).setNumberFormat('@');
    // NEW: Photo en texte (si présente)
    var photoIdx = res.header.indexOf('Photo');
    if (photoIdx >= 0 && sh.getLastRow() > 1) {
      sh.getRange(2, photoIdx + 1, sh.getLastRow() - 1, 1).setNumberFormat('@');
    }
  }

  appendImportLog_(ss, 'RETRO_MEMBRES_SHEET_OK', 'rows=' + res.rows.length);
  return res.rows.length;
}


/** Export XLSX rapide (avec filtrage incrémental optionnel) */
function exportRetroMembresXlsxToDrive(seasonSheetId, options) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);

  // Param dédié (FULL par défaut pour capter les photos modifiées)
  var windowDays = parseInt(readParam_(ss, 'RETRO_EXPORT_LAST_DAYS_MEMBRES') || '0', 10);
  var cutoffDate = (windowDays > 0) ? new Date(Date.now() - windowDays * 86400000) : null;

  var incrOn = String(readParam_(ss, 'INCREMENTAL_ON') || '1').toLowerCase();
  var allowIncr = (incrOn === '1' || incrOn === 'true' || incrOn === 'yes' || incrOn === 'oui');

  var srcKey = (typeof PARAM_KEYS !== 'undefined' && PARAM_KEYS.RETRO_MEMBRES_READ_SRC)
    ? PARAM_KEYS.RETRO_MEMBRES_READ_SRC : 'RETRO_MEMBRES_READ_SOURCE';
  var srcParam = String(readParam_(ss, srcKey) || 'LEGACY').toUpperCase();
  var useJoueurs = (srcParam === 'JOUEURS' && typeof buildRetroMembresRowsFromJoueurs_ === 'function');

  var res = useJoueurs ? buildRetroMembresRowsFromJoueurs_(seasonSheetId)
                       : buildRetroMembresRows(seasonSheetId);
  var header = (res && res.header) || [];
  var nbCols = (res && res.nbCols) || header.length || 1;
  var rowsAll = (res && res.rows) || [];

  try { if (typeof appendImportLog_ === 'function') appendImportLog_(ss, 'RETRO_MEMBRES_SOURCE', JSON.stringify({ source: useJoueurs ? 'JOUEURS' : 'LEGACY' })); } catch(e){}

  var rows = rowsAll, filtered = false;

  if (windowDays > 0) {
    // Lit directement ACHATS_LEDGER et garde tous les passeports qui ont AU MOINS UN achat dans la fenêtre
    var led = readSheetAsObjects_(ss.getId(), 'ACHATS_LEDGER') || { header: [], rows: [] };
    function pick(obj, keys) { for (var i=0;i<keys.length;i++){ var k=keys[i]; if (Object.prototype.hasOwnProperty.call(obj,k)) return k; } return null; }
    var sample = (led.rows && led.rows[0]) ? led.rows[0] : {};
    var COL_PASS = pick(sample, ['Passeport #','Passeport','Passport','PS_Passport']);
    var COL_DATE = pick(sample, ['Date de la facture','Date Facture','Date facture','DateFacture','Date']);
    if (COL_PASS && COL_DATE) {
      var recent = new Set();
      for (var i=0;i<(led.rows||[]).length;i++){
        var L = led.rows[i];
        var d = parseFlexibleDate_(L[COL_DATE]);
        if (d && d >= cutoffDate) {
          var p8 = String(L[COL_PASS]||'').replace(/\D/g,'').padStart(8,'0');
          if (p8) recent.add(p8);
        }
      }
      rows = rowsAll.filter(function(r){
        var p = String(r[0]||'').replace(/\D/g,'').padStart(8,'0');
        return recent.has(p);
      });
      filtered = (rows.length !== rowsAll.length);
    }
  } else if (allowIncr) {
    // Si tu préfères vraiment toujours FULL pour membres, mets allowIncr=false ici
    var touched = _rm_readTouchedPassportSet_(options);
    rows = _rm_filterRowsByPassports_(rowsAll, touched);
    filtered = rows.length !== rowsAll.length;
  }

  var temp = SpreadsheetApp.create('Export temporaire - Retro Membres');
  var tmp = temp.getSheets()[0];
  tmp.setName('Export');

  var all = [header].concat(rows);
  if (typeof normalizePassportToText8_ === 'function') {
    for (var r=1;r<all.length;r++){ if(all[r]&&all[r].length) all[r][0]=normalizePassportToText8_(all[r][0]); }
  }
  if (all.length){
    tmp.getRange(1,1,all.length,nbCols).setValues(all);
    if (all.length>1) tmp.getRange(2,1,all.length-1,1).setNumberFormat('@');
    var photoIdx = header.indexOf('Photo');
    if (photoIdx>=0) tmp.getRange(2,photoIdx+1,all.length-1,1).setNumberFormat('@');
  }
  SpreadsheetApp.flush();

  var url='https://docs.google.com/spreadsheets/d/'+temp.getId()+'/export?format=xlsx';
  var blob=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}}).getBlob();
  var ts=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd-HHmm');
  blob.setName('Export_Retro_Membres_'+ts+(filtered?'_INCR':'')+'.xlsx');

  var folderId=readParam_(ss,PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID)||'';
  var dest=folderId?DriveApp.getFolderById(folderId):DriveApp.getRootFolder();
  var xlsx=dest.createFile(blob);

  DriveApp.getFileById(temp.getId()).setTrashed(true);
  try{ appendImportLog_(ss,(useJoueurs?'RETRO_MEMBRES_XLSX_OK_FAST_J':'RETRO_MEMBRES_XLSX_OK_FAST'),
    xlsx.getName()+' -> '+dest.getName()+' (rows='+rows.length+', filtered='+filtered+')'); }catch(e){}
  return { fileId:xlsx.getId(), name:xlsx.getName(), rows:rows.length, filtered:filtered };

  function parseFlexibleDate_(v){
    if(!v) return null;
    if(v instanceof Date && !isNaN(+v)) return v;
    if (typeof v === 'number') {
      var d = new Date(1899, 11, 30); return new Date(d.getTime() + v*86400000);
    }
    var s=String(v).trim();
    var d2=new Date(s); if(!isNaN(+d2)) return d2;
    var m=s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
    if(m){ var dd=+m[1],MM=+m[2]-1,yyyy=+m[3]; var hh=+(m[4]||'0'),mi=+(m[5]||'0'),ss=+(m[6]||'0');
      var d3=new Date(yyyy,MM,dd,hh,mi,ss); if(!isNaN(+d3)) return d3; }
    return null;
  }
}




// --- Règles FAST (réutilise ton computeRulesFromAggregates_ si déjà collé)
function runEvaluateRulesFast_() {
  var seasonId = getSeasonId_();
  var ss = SpreadsheetApp.openById(seasonId);
  var res = computeRulesFromAggregates_(ss, null); // lit JOUEURS + LEDGER et calcule

  var shE = ss.getSheetByName('SUIVI_ERREURS') || ss.insertSheet('SUIVI_ERREURS');
  shE.clearContents();
  var header = res.header || ['Passeport #', 'PS', 'Courriel', 'CodeErreur', 'Message', 'Saison', 'CreatedAt'];
  shE.getRange(1, 1, 1, header.length).setValues([header]);
  if (res.rows && res.rows.length) {
    shE.getRange(2, 1, res.rows.length, header.length).setValues(res.rows);
  }
  appendImportLog_(ss, 'RULES_FULL_OK', { count: (res.rows || []).length });
  return { count: (res.rows || []).length };
}

// --- Alias: quelle que soit la version legacy appelée, on route sur FAST
function runEvaluateRules() {
  return runEvaluateRulesFast_();
}

// --- No-op pour désarmer une ancienne référence si elle réapparaît
function _rulesClearErreursSheet_() {
  // volontairement vide : les nouvelles règles gèrent l’effacement/écriture d’un coup
}



/* ========== Exposition facultative via Library ========== */
if (typeof Library !== 'undefined') {
  Library.buildRetroMembresRows = buildRetroMembresRows;
  Library.writeRetroMembresSheet = writeRetroMembresSheet;
  Library.exportRetroMembresXlsxToDrive = exportRetroMembresXlsxToDrive;
}


/** Build rétro-membres à partir de JOUEURS (rapide, inclut articles-seulement) */
/** Rétro-membres — version FAST depuis JOUEURS
 * - Reprend le HEADER complet de la legacy buildRetroMembresRows()
 * - Inclut membres “articles-seulement” (JOUEURS les contient déjà)
 * - Adapté = 1 (aucun mail, mais export OK)
 * - CDP: utilise j.cdpCount ; défaut 0 pour U9–U12 si non-adapté et cdpCount vide
 * - Camp: “Oui” si j.hasCamp==true
 * - Photo: priorité MEMBRES_GLOBAL, sinon fallback INSCRIPTIONS ; si j.PhotoStr existe on l’utilise directement
 */
// === Helpers spécifiques rétro_membres ===
const CDP_MAPKEY_1 = 'u9-u12-1-entrainement-cdp-par-semaine-automne-hiver';
const CDP_MAPKEY_2 = 'u9-u12-2-entrainements-cdp-par-semaine-automne-hiver';

function _normP8_(p) {
  return (p == null || p === '') ? '' : String(p).replace(/\D/g, '').padStart(8, '0');
}
function _genreToMF_(g) {
  var s = String(g || '').trim();
  if (!s) return '';
  var c = s.toUpperCase().charAt(0);
  return (c === 'M' || c === 'F') ? c : '';
}
/** LEDGER → CDP (0/1/2/'') par MapKey, Status=1 et isIgnored=0, saison courante */
function _computeCDPFromLedgerByMapKey_(ledgerRows, saison, passport8) {
  var p8 = _normP8_(passport8);
  if (!p8) return '';
  var val = 0;
  (ledgerRows || []).forEach(function (r) {
    if (String(r['Saison'] || '') !== String(saison || '')) return;
    if ((Number(r['Status']) || 0) !== 1) return;
    if ((Number(r['isIgnored']) || 0) === 1) return;
    var rp = _normP8_(r['Passeport #'] || r['Passeport'] || '');
    if (rp !== p8) return;
    var mk = String(r['MapKey'] || '').toLowerCase();
    if (mk === CDP_MAPKEY_2) val = Math.max(val, 2);
    else if (mk === CDP_MAPKEY_1) val = Math.max(val, 1);
  });
  return (val === 0) ? '' : val;
}

/* ===================== Construire les lignes "Rétro - Membres" (source = JOUEURS) ===================== */
function buildRetroMembresRowsFromJoueurs_(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var saison = readParam_(ss, 'SEASON_LABEL') || '';
  var includePhoto = (String(readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) || 'FALSE').toUpperCase() === 'TRUE');

  var joueurs = readSheetAsObjects_(ss.getId(), SHEETS.JOUEURS);
  var ledger  = readSheetAsObjects_(ss.getId(), SHEETS.ACHATS_LEDGER);
  var J = joueurs.rows || [];
  var L = ledger.rows || [];

  // --- Helper clés passeport (normalisation TXT-8 si dispo)
  function RM_key_(pass) {
    var p = String(pass || '').trim();
    if (typeof normalizePassportToText8_ === 'function') {
      try { var p8 = normalizePassportToText8_(p); if (p8) return p8; } catch(_) {}
    }
    return p;
  }

  // --- Détecteurs "élite" et U-num depuis libellés / tags
  var ELITE_KEYWORDS = ['D1+','D1','PLSJQ','LDP','ÉLITE','ELITE','SÉLECTION','SELECTION'];

  function RM_extractU_(s) {
    var m = String(s||'').match(/\bU\s*-?\s*(\d{1,2})\b/i);
    return m ? parseInt(m[1],10) : 0;
  }
  function RM_isEliteByName_(feeName){
    var n = String(feeName||'').toUpperCase();
    return ELITE_KEYWORDS.some(function(k){ return n.indexOf(k) >= 0; });
  }
  function RM_tagsToString_(t) {
    if (!t) return '';
    if (Array.isArray(t)) return t.join(',');
    try {
      if (typeof t === 'string' && t.trim().startsWith('[')) return JSON.parse(t).join(',');
    } catch(_) {}
    return String(t);
  }
  function RM_hasEliteTag_(row){
    var t = RM_tagsToString_(row['Tags'] || row['Tag'] || row['tags'] || '');
    return t.toLowerCase().indexOf('inscription_elite') >= 0;
  }
  // Heuristique: repérer une "INSCRIPTION" (vs article) dans le ledger
  function RM_isInscriptionLedgerRow_(row, feeName) {
    var t = RM_tagsToString_(row['Tags'] || row['Tag'] || row['tags'] || '').toLowerCase();
    var n = String(feeName || '').toLowerCase();
    // Priorité aux tags explicites
    if (t.indexOf('inscription') >= 0) return true;
    // Heuristiques usuelles côté club
    if (/\binscription\b/.test(n)) return true;
    if (/\bsaison\b/.test(n)) return true;            // ex: "Saison Automne-Hiver"
    if (/\bligue\b/.test(n)) return true;
    // Plusieurs articles d'entraînement contiennent "U11/U12", alors on exige un mot-clé "saison/inscription/ligue"
    return false;
  }

  // --- Index MEMBRES_GLOBAL pour la Photo (prioritaire)
  var shMemName = (typeof SHEETS !== 'undefined' && SHEETS.MEMBRES_GLOBAL) ? SHEETS.MEMBRES_GLOBAL : 'MEMBRES_GLOBAL';
  var mem = readSheetAsObjects_(ss.getId(), shMemName);
  var indexMemByP8 = {};
  if (mem && mem.rows && mem.rows.length) {
    var pCol = (mem.header || []).indexOf('Passeport #') >= 0 ? 'Passeport #' :
               (mem.header || []).indexOf('Passeport') >= 0 ? 'Passeport' : 'Passeport #';
    mem.rows.forEach(function (r) {
      var p8 = _normP8_(r[pCol] || '');
      if (p8) indexMemByP8[p8] = r;
    });
  }

  // --- Construire les maps: inscription U9-U12 réelle + élite (par passeport) depuis le LEDGER
  var hasU9U12Inscription = Object.create(null); // clé passeport normalisée -> true
  var isEliteInscription  = Object.create(null); // clé passeport normalisée -> true

  // Filtrer le ledger sur la saison ciblée si colonne Saison dispo; sinon on garde tout et on laisse _computeCDPFromLedgerByMapKey_ filtrer pour CDP.
  L.forEach(function (row) {
    var pass = row['Passeport #'] || row['Passeport'] || '';
    if (!pass) return;
    var k = RM_key_(pass);

    var feeName = row['Nom du frais'] || row['Frais'] || row['Produit'] || row['Produit/Service'] || '';

    // Élites
    if (RM_hasEliteTag_(row) || RM_isEliteByName_(feeName)) {
      isEliteInscription[k] = true;
    }

    // U9-U12 "vraie INSCRIPTION" (pas article d'entraînement)
    if (RM_isInscriptionLedgerRow_(row, feeName)) {
      var u = RM_extractU_(feeName);
      if (u >= 9 && u <= 12) {
        hasU9U12Inscription[k] = true;
      }
    }
  });

  // — En-tête rétro (identique à ta version JOUEURS) —
  var HEADER = [
    "Identifiant unique", "Code", "Nom", "Prénom", "Date de naissance",
    "Genre(M pour Masculin ou F pour Féminin)", "Langue", "Courriels", "Adresse", "Ville",
    "Code Postal", "Domicile Téléphone", "Mobile Téléphone", "Travail Téléphone",
    "Parent 1 Nom", "Parent 1 Courriels", "Parent 1 Domicile Téléphone", "Parent 1 Mobile Téléphone", "Parent 1 Travail Téléphone",
    "Parent 2 Nom", "Parent 2 Courriels", "Parent 2 Domicile Téléphone", "Parent 2 Mobile Téléphone", "Parent 2 Travail Téléphone",
    "Autre Nom", "Autre Courriels", "Autre Domicile Téléphone", "Autre Mobile Téléphone", "Autre Travail Téléphone",
    "Position", "Établissement scolaire", "Adapté", "CDP", "Camp"
  ];
  var header = HEADER.slice();
  if (includePhoto) header.push('Photo');

  var IDX = {}; header.forEach(function (h, i) { IDX[h] = i; });

  var rows = [];
  for (var i = 0; i < J.length; i++) {
    var rJ = J[i] || {};
    var pass = rJ['Passeport #'] || rJ['Passeport'] || '';
    if (!pass) continue;

    var row = new Array(header.length); for (var k = 0; k < row.length; k++) row[k] = '';

    // Identité de base
    if (IDX["Identifiant unique"] >= 0) row[IDX["Identifiant unique"]] =
      (typeof normalizePassportToText8_ === 'function') ? normalizePassportToText8_(pass) : _normP8_(pass);
    if (IDX["Nom"] >= 0) row[IDX["Nom"]] = rJ['Nom'] || '';
    if (IDX["Prénom"] >= 0) row[IDX["Prénom"]] = rJ['Prénom'] || rJ['Prenom'] || '';
    if (IDX["Date de naissance"] >= 0) row[IDX["Date de naissance"]] = rJ['DateNaissance'] || rJ['Naissance'] || '';
    if (IDX["Genre(M pour Masculin ou F pour Féminin)"] >= 0) row[IDX["Genre(M pour Masculin ou F pour Féminin)"]] =
      _genreToMF_(rJ['Genre'] || rJ['Identité de genre'] || rJ['Sexe']);
    if (IDX["Courriels"] >= 0) row[IDX["Courriels"]] = rJ['Courriels'] || '';

    // Adapté
    var isAdapteRaw = String(rJ['isAdapte'] || '').trim();
    var isAdapte = (isAdapteRaw === '1' || /^true$/i.test(isAdapteRaw));
    if (IDX["Adapté"] >= 0) row[IDX["Adapté"]] = isAdapte ? 1 : '';

    // CDP
    var cdp = _computeCDPFromLedgerByMapKey_(L, saison, pass); // CDP 1/2 centralisé si dispo

    // Fallback « U9-U12 non-adapté => 0 » durci pour exiger une vraie INSCRIPTION U9-U12 et non-élite
    if (cdp === '' || cdp == null) {
      var band = String(rJ['AgeBracket'] || rJ['ProgramBand'] || '').toUpperCase();
      var isU9U12 = /U9-?U12/.test(band);
      if (!isU9U12) {
        var ageNum = Number(String(rJ['Age'] || '').toString().replace(',', '.'));
        if (!isNaN(ageNum) && ageNum >= 9 && ageNum <= 12) isU9U12 = true;
      }
      if (!isU9U12) {
        (function () {
          function _extractBirthYearLoose_(v) {
            if (v instanceof Date && !isNaN(+v)) return v.getFullYear();
            var s = String(v); var m = s.match(/\b(19|20)\d{2}\b/);
            if (m) return +m[0];
            m = s.match(/\b(\d{1,2})[\/\-](\d{1,2})[\/\-]((19|20)\d{2})\b/);
            return m ? +m[3] : 0;
          }
          function _parseSeasonYear_(label) {
            var m = String(label || '').match(/\b(19|20)\d{2}\b/);
            return m ? +m[0] : (new Date()).getFullYear();
          }
          var by = _extractBirthYearLoose_(rJ['DateNaissance'] || rJ['Naissance'] || '');
          if (by) {
            var sy = _parseSeasonYear_(saison);
            var ageCalc = sy - by;
            if (ageCalc >= 9 && ageCalc <= 12) isU9U12 = true;
          }
        })();
      }

      // 🔒 Le 0 n'est posé que si: non adapté + (9–12 par âge/band) + **inscription U9-U12 réelle** + **non-élite**
      var kPass = RM_key_(pass);
      if (!isAdapte && isU9U12 && hasU9U12Inscription[kPass] && !isEliteInscription[kPass]) {
        cdp = 0;
      }
    }
    if (IDX["CDP"] >= 0) row[IDX["CDP"]] = (cdp === '' ? '' : cdp);

    // Camp (depuis JOUEURS)
    var hasCamp = String(rJ['hasCamp'] || '').toUpperCase();
    if (IDX["Camp"] >= 0) row[IDX["Camp"]] = (hasCamp === 'TRUE' || hasCamp === 'OUI') ? 'Oui' : '';

    // Photo
    if (includePhoto && IDX["Photo"] >= 0) {
      var wrotePhoto = false;
      var strJ = rJ['PhotoStr'] || '';
      if (String(strJ).trim() !== '') { row[IDX["Photo"]] = String(strJ).trim(); wrotePhoto = true; }
      if (!wrotePhoto) {
        var ageJ = rJ['Age'] || '';
        var adaJ = rJ['isAdapte'] || '';
        var hasInJ = rJ['hasInscription'] || '';
        var expJ = rJ['PhotoExpireLe'] || '';
        row[IDX["Photo"]] = _rm_statusFromFields_(ss, expJ, ageJ, adaJ, hasInJ); wrotePhoto = true;
      }
      if (!wrotePhoto) {
        var p8 = _normP8_(pass);
        var memRow = indexMemByP8[p8] || null;
        var exp = memRow ? (memRow['PhotoExpireLe'] || memRow['Photo Expire Le'] || '') : '';
        var by2 = _extractBirthYearLoose_(row[IDX["Date de naissance"]]);
        var sy2 = parseSeasonYear_(saison);
        var ageFromDob = (by2 && sy2) ? (sy2 - by2) : '';
        var isAda2 = (row[IDX["Adapté"]] === 1);
        row[IDX["Photo"]] = _rm_statusFromFields_(ss, exp, ageFromDob, isAda2, /*hasInscription*/ '');
      }
    }

    rows.push(row);
  }

  return { header: header, rows: rows, nbCols: header.length };
}
