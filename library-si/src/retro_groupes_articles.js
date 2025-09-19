/**
* retro_groupes_articles.gs — v0.11
* - Conserve la logique existante de build (buildRetroGroupeArticlesRows)
* - ✨ Ajout export incrémental : filtrage par passeports «touchés»
* - options.onlyPassports (Array/Set)
* - sinon: DocumentProperties.LAST_TOUCHED_PASSPORTS (CSV)
* - Écrit XLSX dans le dossier paramétré; suffixe _INCR si filtré
* - Forçage col. A (Passeport) en texte; normalisation 8 caractères si helpers dispo
*/

/* ===================== Param keys ===================== */
if (typeof PARAM_KEYS === 'undefined') { var PARAM_KEYS = {}; }
PARAM_KEYS.RETRO_GART_SHEET_NAME = PARAM_KEYS.RETRO_GART_SHEET_NAME || 'RETRO_GART_SHEET_NAME';
PARAM_KEYS.RETRO_GART_EXPORTS_FOLDER_ID = PARAM_KEYS.RETRO_GART_EXPORTS_FOLDER_ID || 'RETRO_GART_EXPORTS_FOLDER_ID';

PARAM_KEYS.RETRO_GART_IGNORE_FEES_CSV = PARAM_KEYS.RETRO_GART_IGNORE_FEES_CSV || 'RETRO_GART_IGNORE_FEES_CSV';
PARAM_KEYS.RETRO_GART_ELITE_KEYWORDS = PARAM_KEYS.RETRO_GART_ELITE_KEYWORDS || 'RETRO_GART_ELITE_KEYWORDS';
PARAM_KEYS.RETRO_GART_REQUIRE_MAPPING = PARAM_KEYS.RETRO_GART_REQUIRE_MAPPING || 'RETRO_GART_REQUIRE_MAPPING';
PARAM_KEYS.RETRO_GART_REQUIRE_INSCRIPTION = PARAM_KEYS.RETRO_GART_REQUIRE_INSCRIPTION || 'RETRO_GART_REQUIRE_INSCRIPTION';
PARAM_KEYS.RETRO_DEBUG = PARAM_KEYS.RETRO_DEBUG || 'RETRO_DEBUG';

PARAM_KEYS.RETRO_RULES_JSON = PARAM_KEYS.RETRO_RULES_JSON || 'RETRO_RULES_JSON';

PARAM_KEYS.RETRO_GROUP_GROUPE_FMT = PARAM_KEYS.RETRO_GROUP_GROUPE_FMT || 'RETRO_GROUP_GROUPE_FMT';
PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT = PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT || 'RETRO_GROUP_CATEGORIE_FMT';

// Param erreurs
PARAM_KEYS.RETRO_ERRORS_SHEET_NAME = PARAM_KEYS.RETRO_ERRORS_SHEET_NAME || 'RETRO_ERRORS_SHEET_NAME';

// Adapté (pour exclure CDP0)
PARAM_KEYS.RETRO_ADAPTE_KEYWORDS = PARAM_KEYS.RETRO_ADAPTE_KEYWORDS || 'RETRO_ADAPTE_KEYWORDS';
PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS = PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS || 'RETRO_GROUP_SA_KEYWORDS';
PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID = PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID || 'RETRO_EXPORTS_FOLDER_ID';


/* ===================== Fallbacks minimes ===================== */
if (typeof CONTROL_COLS === 'undefined') {
  var CONTROL_COLS = { ROW_HASH: 'ROW_HASH', CANCELLED: 'CANCELLED', EXCLUDE_FROM_EXPORT: 'EXCLUDE_FROM_EXPORT', LAST_MODIFIED_AT: 'LAST_MODIFIED_AT' };
}

// Cache global partagé (idempotent)
var __retroRulesCache = (typeof __retroRulesCache !== 'undefined')
  ? __retroRulesCache
  : { at: 0, data: null };

function loadRetroRules_(ss) {
  if (typeof SR_loadRetroRules_ === 'function') {
    return SR_loadRetroRules_(ss);
  }
  var now = Date.now();
  if (__retroRulesCache.data && (now - __retroRulesCache.at) < 5 * 60 * 1000) return __retroRulesCache.data;

  var raw = readParam_(ss, PARAM_KEYS.RETRO_RULES_JSON) || '';
  if (!raw) {
    var shJson = ss.getSheetByName('RETRO_RULES_JSON');
    if (shJson && shJson.getLastRow() >= 1 && shJson.getLastColumn() >= 1) {
      var vals = shJson.getDataRange().getDisplayValues();
      var pieces = [];
      for (var i = 0; i < vals.length; i++) for (var j = 0; j < vals[i].length; j++) {
        var cell = vals[i][j]; if (cell != null && String(cell).trim() !== '') pieces.push(String(cell));
      }
      raw = pieces.join('\n');
      appendImportLog_(ss, 'RETRO_RULES_JSON_SHEET_READ', 'chars=' + raw.length);
    }
  }
  var rules = [];
  if (raw) {
    try { var arr = JSON.parse(raw); rules = Array.isArray(arr)?arr:[]; }
    catch(e){ appendImportLog_(ss, 'RETRO_RULES_JSON_PARSE_FAIL', String(e)); }
  }
  if (!rules.length) {
    var ignoreCsv = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV) || 'senior,u-sé,adulte,ligue';
    var adapteCsv = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || 'adapté,adapte';
    var campCsv   = readParam_(ss, PARAM_KEYS.RETRO_CAMP_KEYWORDS)   || 'camp de sélection u13,camp selection u13,camp u13';
    var photoOn   = (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) || 'FALSE').toUpperCase() === 'TRUE';
    var photoCol  = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL) || '';
    var warnMmDd  = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD) || '03-01';
    var absDate   = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';
    rules = [
      { id:'ignore_fees', enabled:true, scope:'both',
        when:{ field:'Nom du frais', contains_any: ignoreCsv.split(',') },
        action:{ type:'ignore_row' } },
      { id:'adapte_flag', enabled:true, scope:'both',
        when:{ field:'Nom du frais', contains_any: adapteCsv.split(',') },
        action:{ type:'set_member_field', field:'adapte', value:1 } },
      { id:'cdp_2', enabled:true, scope:'articles',
        when:{ catalog_exclusive_group:'CDP_ENTRAINEMENT', text_contains_any:['2','2 entrainements'] },
        action:{ type:'set_member_field_max', field:'cdp', value:2 } },
      { id:'cdp_1', enabled:true, scope:'articles',
        when:{ catalog_exclusive_group:'CDP_ENTRAINEMENT' },
        action:{ type:'set_member_field_max', field:'cdp', value:1 } },
      { id:'camp_u13', enabled:true, scope:'articles',
        when:{ field:'Nom du frais', contains_any: campCsv.split(',') },
        action:{ type:'set_member_field', field:'camp', value:'Oui' } }
    ];
    if (photoOn && photoCol) {
      rules.push({ id:'photo_policy', enabled:true, scope:'member',
        action:{ type:'compute_photo', expiry_col:photoCol, warn_mmdd:warnMmDd, abs_date:absDate }});
    }
    appendImportLog_(ss, 'RETRO_RULES_JSON_FALLBACK', 'using PARAMS-derived defaults');
  }
  __retroRulesCache = { at: now, data: rules };
  return rules;
}


/* ===================== DEBUG ===================== */
function _dbg_(on, msg, obj) { if (!on) return; if (obj === undefined) { Logger.log(msg); } else { try { Logger.log(msg + ' ' + JSON.stringify(obj)); } catch (_) { Logger.log(msg); } } }

/* ===================== Helpers genre ===================== */
function _ga_extractGenreSmart_(row) {
  var keys = [
    'Identité de genre', 'Identité de Genre', 'Identite de genre', 'Identite de Genre',
    'Genre', 'Sexe', 'Sex', 'Gender', 'F/M', 'MF', 'Gendre', 'Type',
    'Categorie', 'Catégorie', 'Catégories'
  ];
  var raw = '';
  for (var i = 0; i < keys.length; i++) {
    if (row && row.hasOwnProperty(keys[i]) && String(row[keys[i]] || '').trim() !== '') {
      raw = String(row[keys[i]]); break;
    }
  }
  if (!raw) return { label: '', initiale: '' };
  function _nrmLow(s) { try { s = String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return s.toLowerCase().trim(); }
  var n = _nrmLow(raw);

  // Supporte U11M / U11F collés OU avec espace
  if (/^(m|masculin|male|man|garcon|gar\u00e7on|homme|boy)\b/.test(n) || /\bu ?\d+\s*m\b/.test(n) || /\bu ?\d+m\b/.test(n) || /\bmasc\b/.test(n))
    return { label: 'Masculin', initiale: 'M' };
  if (/^(f|feminin|female|woman|fille|dame|girl)\b/.test(n) || /\bu ?\d+\s*f\b/.test(n) || /\bu ?\d+f\b/.test(n) || /\bfem\b/.test(n))
    return { label: 'Féminin', initiale: 'F' };
  if (/^(mixte|mix|x|non binaire|non-binaire|nb|autre)\b/.test(n))
    return { label: 'Mixte', initiale: 'X' };
  return { label: String(raw), initiale: String(raw).charAt(0).toUpperCase() };
}

/* ===================== Utils ===================== */
function _ga_nrm_(s) { s = String(s == null ? '' : s); try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return s; }
function _ga_low_(s) { return _ga_nrm_(s).toLowerCase().trim(); }
function _ga_containsAny_(txt, csv) {
  var t = _ga_low_(txt || '');
  return String(csv || '').split(',').map(_ga_low_).filter(Boolean).some(function (p) { return t.indexOf(p) !== -1; });
}
function _ga_isActiveArticle_(r) {
  var can = String(r[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true';
  var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT] || '').toLowerCase() === 'true';
  var st = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
  return !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
}
function _ga_pad2_(n) { n = Number(n || 0); return (n < 10 ? ('0' + n) : String(n)); }

/* ==== U / U2 ==== */
function _ga_deriveAgeYear_(row) {
  var dn = row['Date de naissance'] || row['Naissance'] || '';
  if (dn instanceof Date) return dn.getFullYear();
  if (dn) { var m = String(dn).match(/(19|20)\d{2}/); if (m) return parseInt(m[0], 10); }
  return null;
}
function _ga_ageCat_(birthYear, seasonYear) {
  if (!birthYear || !seasonYear) return '';
  var age = seasonYear - birthYear;
  if (age < 4 || age > 99) return '';
  return 'U' + _ga_pad2_(age); // "U09"
}
function _ga_U_(birthYear, seasonYear) {
  var u2 = _ga_ageCat_(birthYear, seasonYear);
  return u2 ? ('U' + parseInt(u2.slice(1), 10)) : ''; // "U9", "U10", ...
}
function _ga_extractUFromFeeText_(feeName) {
  var s = String(feeName || '');
  var m = s.match(/U\s*[-–]?\s*(\d{1,2})/i);
  return m ? ('U' + parseInt(m[1], 10)) : '';
}
function _ga_computeUandU2_(row, seasonYear, feeName) {
  var by = _ga_deriveAgeYear_(row);
  var U = '', U2 = '';
  if (by) {
    U2 = _ga_ageCat_(by, seasonYear);
    if (U2) U = 'U' + parseInt(U2.slice(1), 10);
  }
  return { U: U, U2: U2 };
}

/* ==== Templating ==== */
function _ga_tpl_(tpl, vars) {
  tpl = String(tpl == null ? '' : tpl);
  return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function (_, k) { return (vars && k in vars && vars[k] != null) ? String(vars[k]) : ''; });
}

/* ===== Règles ===== */
function _ga_loadRules_(ss) { if (typeof loadRetroRules_ === 'function') return loadRetroRules_(ss); return []; }
function _ga_applyRowRulesMaybeSkip_(rules, articleRow, ctx) {
  if (!rules || !rules.length || typeof applyRetroRowRules_ !== 'function') return false;
  var fakeMember = {};
  var res = applyRetroRowRules_(rules, 'articles', articleRow, fakeMember, ctx);
  return !!(res && res.skip);
}



/* ==== Normalisation passeport ==== */
function _ga_norm_passport_(ss, v) {
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (typeof normalizePassportPlain8_ === 'function') return normalizePassportPlain8_(s);
  if (/^\d+$/.test(s)) {
    var width = parseInt(readParam_(ss, 'PASSPORT_PAD_WIDTH') || '8', 10);
    if (isNaN(width) || width < 1) width = 8;
    s = (Array(width + 1).join('0') + s).slice(-width);
  }
  return s;
}

/* ==== Normalisation texte (espaces/tirets) ==== */
function _low_(s) {
  s = String(s == null ? '' : s);
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { }
  s = s
    .replace(/\u00A0/g, ' ')                 // NBSP → espace
    .replace(/[\u2010-\u2015\u2212]/g, '-')  // tous les dashes → '-'
    .replace(/\s+/g, ' ')                    // compacter
    .trim()
    .toLowerCase();
  return s;
}

/* ==== Match alias 3 étages (brut / normalisé / tokens) ==== */
function _aliasMatchExplain_(feeName, alias) {
  var feeRaw = String(feeName || '');
  var aliRaw = String(alias || '');
  var feeNorm = _low_(feeRaw);
  var aliNorm = _low_(aliRaw);

  var rawContains = !!(aliRaw && feeRaw.indexOf(aliRaw) !== -1);
  var normContains = !!(aliNorm && feeNorm.indexOf(aliNorm) !== -1);

  var tokensOk = false;
  if (!rawContains && !normContains) {
    var toks = aliNorm.split(/[^a-z0-9]+/).filter(function (t) { return t.length >= 2; });
    if (toks.length) {
      tokensOk = toks.every(function (t) { return feeNorm.indexOf(t) !== -1; });
    }
  }

  return {
    ok: rawContains || normContains || tokensOk,
    feeRaw: feeRaw,
    aliRaw: aliRaw,
    feeNorm: feeNorm,
    aliNorm: aliNorm,
    rawContains: rawContains,
    normContains: normContains,
    tokensOk: tokensOk
  };
}

/** Renvoie {passed:[], failedU:[], triedAliases:[...], matchedAliases:[...], debug:[...]} */
function _findArticleMappingCandidates_(maps, feeName, vars, DEBUG) {
  var passed = [], failedU = [], tried = [], matched = [], debug = [];
  for (var i = 0; i < maps.length; i++) {
    var m = maps[i];
    if (m.Type !== 'article') continue;
    if (!m.AliasContains) continue;

    var ali = String(m.AliasContains);
    tried.push(_low_(ali));

    var ex = _aliasMatchExplain_(feeName, ali);
    debug.push({ ali: ali, fee: feeName, rawContains: ex.rawContains, normContains: ex.normContains, tokensOk: ex.tokensOk });
    if (!ex.ok) continue; // alias ne matche pas (aucun des 3 modes)
    matched.push(_low_(ali));

    // Genre
    if (m.Genre && m.Genre !== '*' && m.Genre !== (vars.genreInitiale || '')) continue;

    // U
    var uNum = parseInt(String(vars.U || '').replace(/^U/i, ''), 10) || 0;

    if (vars.U) { var mm = String(vars.U).match(/^U(\d{1,2})$/i); if (mm) uNum = parseInt(mm[1], 10); }
    var okU = true;
    if (uNum) {
      if (m.Umin != null && uNum < m.Umin) okU = false;
      if (m.Umax != null && uNum > m.Umax) okU = false;
      if (okU) passed.push(m); else failedU.push(m);
    }
  }
  return { passed: passed, failedU: failedU, triedAliases: tried, matchedAliases: matched, debug: debug };
}

/** Rend la 1ère règle qui passe, ou null */
function _applyUnifiedMapping_(maps, feeName, vars, DEBUG) {
  var cand = _findArticleMappingCandidates_(maps, feeName, vars, DEBUG).passed;
  if (!cand.length) return null;
  var m = cand[0];
  if (m.Exclude) return { exclude: true, exclusiveGroup: m.ExclusiveGroup || '' };
  return {
    groupe: m.GroupeFmt ? _ga_tpl_(m.GroupeFmt, vars) : '',
    categorie: m.CategorieFmt ? _ga_tpl_(m.CategorieFmt, vars) : '',
    exclusiveGroup: m.ExclusiveGroup || '',
    code: m.Code || ''
  };
}

/* ===================== Unification INSCRIPTIONS + ARTICLES ===================== */
function _ga_buildMemberIndex_(ss) {
  var idx = {};
  function ensure(pRaw) {
    var k = _ga_norm_passport_(ss, pRaw);
    if (!k) return null;
    if (!idx[k]) idx[k] = { passeport: k, nom: '', prenom: '', dob: '', genreInit: '', genreLabel: '' };
    return idx[k];
  }
  function mergeFromRow(row, prefer) {
    var p = row['Passeport #']; if (!p) return;
    var m = ensure(p); if (!m) return;
    var nom = row['Nom'] || ''; var prenom = row['Prénom'] || row['Prenom'] || '';
    if (nom && !m.nom) m.nom = nom;
    if (prenom && !m.prenom) m.prenom = prenom;
    var dob = row['Date de naissance'] || row['Naissance'] || '';
    if (dob && (!m.dob || prefer === 'insc')) m.dob = dob;
    var g = _ga_extractGenreSmart_(row);
    if (g.initiale && (!m.genreInit || prefer === 'insc')) { m.genreInit = g.initiale; m.genreLabel = g.label; }
  }

  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var arts = readSheetAsObjects_(ss.getId(), SHEETS.ARTICLES);

  (insc.rows || []).forEach(function (r) { mergeFromRow(r, 'insc'); });
  (arts.rows || []).forEach(function (r) { mergeFromRow(r, 'arts'); });

  return idx;
}

/* ===================== Lecture des «touchés» & filtre ===================== */
function _ga_readTouchedPassportSet_(ss, options) {
  options = options || {};
  var set = {};

  // 1) options.onlyPassports
  var list = options.onlyPassports;
  if (list && typeof list.forEach === 'function') {
    list.forEach(function (p) { var t = _ga_norm_passport_(p); if (t) set[t] = true; });
  }

  // 2) Fallback: DocumentProperties.LAST_TOUCHED_PASSPORTS (JSON ou CSV)
  if (!Object.keys(set).length) {
    try {
      var raw = (PropertiesService.getDocumentProperties().getProperty('LAST_TOUCHED_PASSPORTS') || '').trim();
      if (raw) {
        var arr = (raw.charAt(0) === '[') ? JSON.parse(raw) : raw.split(',');
        arr.forEach(function (p) { var t = _ga_norm_passport_(p); if (t) set[t] = true; });
      }
    } catch (_) { }
  }
  return set; // possiblement vide → pas de filtrage
}




function _ga_filterRowsByPassports_(rows, touchedSet) {
  var keys = Object.keys(touchedSet || {});
  if (!keys.length) return rows;
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var p = _ga_norm_passport_(row && row[0]); // col A = Passeport
    if (p && touchedSet[p]) out.push(row);
  }
  return out;
}

/* ===================== Construction + Erreurs ===================== */
function buildRetroGroupeArticlesRows(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var DEBUG = String(readParam_(ss, PARAM_KEYS.RETRO_DEBUG) || 'FALSE').toUpperCase() === 'TRUE';

  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var art = readSheetAsObjects_(ss.getId(), SHEETS.ARTICLES);

  var rules = _ga_loadRules_(ss);
  var mappings = _loadUnifiedGroupMappings_(ss);
  _dbg_(DEBUG, '[GART] mappings loaded', { count: (mappings || []).length });

  // Filtres
  var ignoreCsv = readParam_(ss, PARAM_KEYS.RETRO_GART_IGNORE_FEES_CSV) || readParam_(ss, 'RETRO_IGNORE_FEES_CSV') || 'senior,u-sé,adulte,ligue';
  var eliteCsv = readParam_(ss, PARAM_KEYS.RETRO_GART_ELITE_KEYWORDS) || 'D1+,LDP,Ligue';
  var requireMp = (String(readParam_(ss, PARAM_KEYS.RETRO_GART_REQUIRE_MAPPING) || 'TRUE').toUpperCase() === 'TRUE');
  var requireInsc = (String(readParam_(ss, PARAM_KEYS.RETRO_GART_REQUIRE_INSCRIPTION) || 'FALSE').toUpperCase() === 'TRUE');

  // Adapté (pour exclure CDP0 warn/export)
  var adapteCsv = (readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || '') + ',' + (readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS) || '');
  adapteCsv = adapteCsv.replace(/^,|,$/g, '');

  // ⬅️ Format Catégorie pour CDP0 (paramétrable)
  var catFmtCDP = readParam_(ss, 'RETRO_GROUP_CATEGORIE_FMT_CDP') || 'CDP {{U2}} {{genreInitiale}}';

  var header = ["Identifiant unique", "Nom", "Prénom", "Date de naissance", "#", "Couleur", "Sous-groupe", "Position", "Équipe/Groupe", "Catégorie"];
  var rows = [];
  var errors = []; // {level, code, passeport, nom, prenom, feeName, message, details}

  var activeArts = (art.rows || []).filter(_ga_isActiveArticle_);
  _dbg_(DEBUG, '[GART] active articles', { count: activeArts.length });

  if (!activeArts.length) return { header: header, rows: rows, nbCols: header.length, errors: errors };

  // Saison/année
  var seasonLabel = readParam_(ss, 'SEASON_LABEL') || (activeArts[0] && activeArts[0]['Saison']) || '';
  var seasonYear = parseSeasonYear_(seasonLabel);
  _dbg_(DEBUG, '[GART] season', { label: seasonLabel, year: seasonYear });

  var ctx = { ss: ss, catalog: (typeof _loadArticlesCatalog_ === 'function' ? _loadArticlesCatalog_(ss) : { match: function () { return null; } }) };

  var memberIdx = _ga_buildMemberIndex_(ss);
  _dbg_(DEBUG, '[GART] memberIdx size', { size: Object.keys(memberIdx).length });

  // Set des passeports avec inscription active (normalisés)
  var inscActivePass = {};
  (insc.rows || []).forEach(function (r) {
    var p = _ga_norm_passport_(ss, r['Passeport #']); if (!p) return;
    var can = String(r[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true';
    var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT] || '').toLowerCase() === 'true';
    var st = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
    var active = !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
    if (active) inscActivePass[p] = true;
  });
  _dbg_(DEBUG, '[GART] active inscriptions (normalized)', { count: Object.keys(inscActivePass).length });

  // Exclusivité & Adapté
  var perPassExclusive = {}; // pass -> { groupName -> [ {feeName, code} ] }
  var perPassIsAdapte = {}; // pass -> true si un article/inscription correspond aux mots-clés "Adapté"

  activeArts.forEach(function (a, idx) {
    if (_ga_applyRowRulesMaybeSkip_(rules, a, ctx)) { _dbg_(DEBUG, '[GART] skip by rules', { i: idx }); return; }

    var feeName = a['Nom du frais'] || a['Frais'] || a['Produit'] || '';
    if (_ga_containsAny_(feeName, ignoreCsv)) { _dbg_(DEBUG, '[GART] skip ignoreCsv', { fee: feeName }); return; }
    if (_ga_containsAny_(feeName, eliteCsv)) { _dbg_(DEBUG, '[GART] skip eliteCsv', { fee: feeName }); return; }

    var passRaw = a['Passeport #'];
    var passK = _ga_norm_passport_(ss, passRaw);
    if (!passK) { _dbg_(DEBUG, '[GART] no passport, skip', { fee: feeName }); return; }

    // Article sans inscription active ?
    if (!inscActivePass[passK]) {
      errors.push({
        level: 'error', code: 'ARTICLE_WITHOUT_INSCRIPTION', passeport: passK,
        nom: (a['Nom'] || ''), prenom: (a['Prénom'] || a['Prenom'] || ''), feeName: feeName,
        message: 'Article actif sans inscription active correspondante', details: {}
      });
      _dbg_(DEBUG, '[GART] ARTICLE_WITHOUT_INSCRIPTION', { pass: passK, fee: feeName });
      if (requireInsc) return; // on ne bloque que si explicitement exigé
    }

    // Marqueur "Adapté"
    if (_ga_containsAny_(feeName, adapteCsv)) perPassIsAdapte[passK] = true;

    // Profil membre fusionné
    var m = memberIdx[passK] || {};
    var nom = (a['Nom'] || '') || m.nom || '';
    var prenom = (a['Prénom'] || a['Prenom'] || '') || m.prenom || '';
    var dob = (a['Date de naissance'] || a['Naissance'] || '') || m.dob || '';

    // U/U2 + genre — on ne déduit JAMAIS U depuis le libellé du frais
    var UU2 = _ga_computeUandU2_({ 'Date de naissance': dob, 'Naissance': dob }, seasonYear, '');
    var U = UU2.U || '';
    var U2 = UU2.U2 || '';
    if (!U2) {
      errors.push({
        level: 'error', code: 'MISSING_DOB_for_U', passeport: passK, nom: nom, prenom: prenom, feeName: feeName,
        message: 'Impossible de dériver U/U2 sans date de naissance (chaque joueur doit avoir un U).',
        details: {}
      });
      _dbg_(DEBUG, '[GART] MISSING_DOB_for_U', { pass: passK, fee: feeName });
      return;
    }

    var gA = _ga_extractGenreSmart_(a);
    var gInit = gA.initiale || m.genreInit || '';
    var gLbl = gA.label || m.genreLabel || (gInit === 'F' ? 'Féminin' : (gInit === 'M' ? 'Masculin' : (gInit === 'X' ? 'Mixte' : '')));

    var vars = { U: U, U2: U2, ageCat: U2, genreInitiale: gInit, genre: gLbl, article: feeName, saison: seasonLabel, annee: seasonYear };

    // Candidats de mapping + debug alias
    var cand = _findArticleMappingCandidates_(mappings, feeName, vars, DEBUG);
    _dbg_(DEBUG, '[GART] alias check', { fee: feeName, debugs: (cand.debug || []).slice(0, 8) });

    // Cas 1: alias matché mais U hors bornes → log AGE_OUT_OF_RANGE et ne pas exporter
    if ((!cand.passed || cand.passed.length === 0) && cand.failedU && cand.failedU.length) {
      var ranges = cand.failedU.map(function (mm) {
        var aR = []; if (mm.Umin != null) aR.push('min ' + mm.Umin); if (mm.Umax != null) aR.push('max ' + mm.Umax);
        return aR.join(', ');
      }).join(' | ');
      errors.push({
        level: 'error', code: 'AGE_OUT_OF_RANGE', passeport: passK, nom: nom, prenom: prenom, feeName: feeName,
        message: 'Âge (U) hors bornes pour cet article', details: { U: U, ranges: ranges }
      });
      _dbg_(DEBUG, '[GART] AGE_OUT_OF_RANGE', { pass: passK, fee: feeName, U: U, ranges: ranges });
      return; // pas d’export
    }

    // Cas 2: aucun alias matché → unmapped (si requireMapping)
    var hasAliasMatch = (cand.matchedAliases && cand.matchedAliases.length) || false;
    if ((!hasAliasMatch) && requireMp) {
      errors.push({
        level: 'error', code: 'ARTICLE_UNMAPPED', passeport: passK, nom: nom, prenom: prenom, feeName: feeName,
        message: 'Aucun mapping article trouvé (requireMapping=TRUE)',
        details: { tried: (cand.triedAliases || []).slice(0, 20), matchedAlias: [], U: U, genre: gInit }
      });
      _dbg_(DEBUG, '[GART] ARTICLE_UNMAPPED (no alias matched)', { fee: feeName, U: U, genre: gInit });
      return;
    }

    // Sélection du mapping applicable (passé Umin/Umax)
    var mp = null;
    if (cand.passed && cand.passed.length) {
      mp = _applyUnifiedMapping_(mappings, feeName, vars, DEBUG);
      if (!mp) {
        var chosen = cand.passed[0] || {};
        var gfmt = chosen.GroupeFmt || chosen.gfmt || chosen.groupeFmt || chosen.groupFmt || readParam_(ss, PARAM_KEYS.RETRO_GROUP_GROUPE_FMT) || '{{U2}}{{genreInitiale}}';
        var cfmt = chosen.CategorieFmt || chosen.cfmt || chosen.categorieFmt || chosen.categoryFmt || readParam_(ss, PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT) || '{{U2}} {{genreInitiale}}';
        mp = {
          groupe: _ga_tpl_(gfmt, vars),
          categorie: _ga_tpl_(cfmt, vars),
          exclude: !!chosen.Exclude || !!chosen.exclude,
          exclusiveGroup: chosen.ExclusiveGroup || chosen.exclusiveGroup || '',
          code: chosen.Code || chosen.code || ''
        };
      }
    } else {
      if (!requireMp) {
        var gfmtFb = readParam_(ss, PARAM_KEYS.RETRO_GROUP_GROUPE_FMT) || '{{U2}}{{genreInitiale}}';
        var cfmtFb = readParam_(ss, PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT) || '{{U2}} {{genreInitiale}}';
        mp = { groupe: _ga_tpl_(gfmtFb, vars), categorie: _ga_tpl_(cfmtFb, vars), exclude: false, exclusiveGroup: '', code: '' };
      } else {
        return;
      }
    }

    if (mp && mp.exclude) { _dbg_(DEBUG, '[GART] excluded by mapping', { fee: feeName }); return; }

    var groupe = (mp && mp.groupe) || '';
    var categ = (mp && mp.categorie) || '';
    var exg = (mp && mp.exclusiveGroup) || '';
    var code = (mp && mp.code) || '';

    if (exg) {
      perPassExclusive[passK] = perPassExclusive[passK] || {};
      perPassExclusive[passK][exg] = perPassExclusive[passK][exg] || [];
      perPassExclusive[passK][exg].push({ feeName: feeName, code: code || feeName });
    }

    if (!groupe && !categ) { _dbg_(DEBUG, '[GART] skip empty group/category', { fee: feeName }); return; }

    var nbCols = header.length;
    var rowOut = new Array(nbCols).fill("");
    rowOut[0] = _ga_norm_passport_(ss, passK);
    rowOut[1] = nom;
    rowOut[2] = prenom;
    rowOut[3] = dob;
    // #, Couleur, Sous-groupe, Position vides
    rowOut[8] = groupe;
    rowOut[9] = categ;

    rows.push(rowOut);
    _dbg_(DEBUG, '[GART] mapped/exported', { pass: passK, fee: feeName, groupe: groupe, categorie: categ, code: code, exg: exg });
  });

  // Conflits d’exclusivité
  Object.keys(perPassExclusive).forEach(function (passK) {
    var ex = perPassExclusive[passK];
    Object.keys(ex).forEach(function (group) {
      var arr = ex[group] || [];
      var distinct = {};
      arr.forEach(function (x) { distinct[String(x.code || '')] = true; });
      var nb = Object.keys(distinct).filter(Boolean).length;
      if (nb > 1) {
        errors.push({
          level: 'error', code: 'EXCLUSIVE_CONFLICT', passeport: passK, nom: '', prenom: '', feeName: '',
          message: 'Conflit d’exclusivité: plusieurs articles du groupe ' + group, details: { group: group, items: arr }
        });
        _dbg_(true, '[GART] EXCLUSIVE_CONFLICT', { pass: passK, group: group, items: arr });
      }
    });
  });

  // CDP0 (warn) U9–U12 hors Adapté → **inclure dans l'export** avec Catégorie préfixée "CDP "
  Object.keys(inscActivePass).forEach(function (passK) {
    var m = memberIdx[passK] || {};
    var UU2m = _ga_computeUandU2_({ 'Date de naissance': m.dob, 'Naissance': m.dob }, seasonYear, '');
    var Um = UU2m.U || '';
    var U2m = UU2m.U2 || '';
    var uNum = parseInt(String(Um).replace(/^U/i, ''), 10);
    if (!(uNum >= 9 && uNum <= 12)) return;

    var isAdapte = !!perPassIsAdapte[passK];
    if (isAdapte) return;

    var hasCDP = perPassExclusive[passK] && (perPassExclusive[passK]['CDP_ENTRAINEMENT'] || perPassExclusive[passK]['CDP']);
    var count = 0;
    if (hasCDP) {
      var a1 = (perPassExclusive[passK]['CDP_ENTRAINEMENT'] || []).length;
      var a2 = (perPassExclusive[passK]['CDP'] || []).length;
      count = a1 + a2;
    }
    if (!count) {
      // Warn (historique/comm)
      errors.push({
        level: 'warn', code: 'CDP0', passeport: passK, nom: (m.nom || ''), prenom: (m.prenom || ''), feeName: '',
        message: 'Membre U9–U12 sans CDP (1/2) — hors Adapté', details: { U: Um }
      });

      // **Ligne d'export synthétique CDP0**
      var gInitM = m.genreInit || (m.genreLabel === 'Féminin' ? 'F' : (m.genreLabel === 'Masculin' ? 'M' : (m.genreLabel === 'Mixte' ? 'X' : '')));
      var groupe = (U2m || '').concat(gInitM ? gInitM : '').concat(' CDP0'); // ex: U10M CDP0

      // Catégorie avec préfixe "CDP " (format paramétrable)
      var varsCDP = {
        U: Um,
        U2: U2m,
        ageCat: U2m,
        genreInitiale: gInitM || '',
        genre: (gInitM === 'F' ? 'Féminin' : (gInitM === 'M' ? 'Masculin' : (gInitM === 'X' ? 'Mixte' : ''))),
        article: 'CDP0',
        saison: seasonLabel,
        annee: seasonYear
      };
      var categ = _ga_tpl_(catFmtCDP, varsCDP); // ex: "CDP U10 M"

      var nbCols = header.length;
      var rowOut = new Array(nbCols).fill("");
      rowOut[0] = _ga_norm_passport_(ss, passK);
      rowOut[1] = m.nom || '';
      rowOut[2] = m.prenom || '';
      rowOut[3] = m.dob || '';
      rowOut[8] = groupe;
      rowOut[9] = categ;
      rows.push(rowOut);

      // (facultatif) noter l’exclusivité avec code "CDP0" dans le groupe CDP_ENTRAINEMENT
      perPassExclusive[passK] = perPassExclusive[passK] || {};
      perPassExclusive[passK]['CDP_ENTRAINEMENT'] = perPassExclusive[passK]['CDP_ENTRAINEMENT'] || [];
      perPassExclusive[passK]['CDP_ENTRAINEMENT'].push({ feeName: 'CDP0', code: 'CDP0' });

      _dbg_(DEBUG, '[GART] CDP0 exported', { pass: passK, groupe: groupe, categorie: categ });
    }
  });

  _dbg_(DEBUG, '[GART] done', { rows: rows.length, errors: errors.length });
  return { header: header, rows: rows, nbCols: header.length, errors: errors };
}



/* ===================== Feuille de travail (facultatif) ===================== */
function writeRetroGroupeArticlesSheet(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var out = buildRetroGroupeArticlesRows(seasonSheetId);
  var sheetName = readParam_(ss, PARAM_KEYS.RETRO_GART_SHEET_NAME) || 'Rétro - Groupe Articles';
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clearContents();
  sh.getRange(1, 1, 1, out.nbCols).setValues([out.header]);
  if (out.rows.length) {
    sh.getRange(2, 1, out.rows.length, out.nbCols).setValues(out.rows);
    sh.autoResizeColumns(1, out.nbCols);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).setNumberFormat('@');
  }
  appendImportLog_(ss, 'RETRO_GART_SHEET_OK', 'rows=' + out.rows.length);
  return out.rows.length;
}

function _writeRetroErrors_(ss, errors) {
  var sheetName = readParam_(ss, PARAM_KEYS.RETRO_ERRORS_SHEET_NAME) || 'Rétro - Erreurs';
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clearContents();

  var header = ["Date", "Niveau", "Code", "Passeport", "# Nom", "Prénom", "Article", "Message", "Détails(JSON)"];
  var rows = (errors || []).map(function (e) {
    return [
      new Date(),
      e.level || 'error',
      e.code || '',
      e.passeport || '',
      e.nom || '',
      e.prenom || '',
      e.feeName || '',
      e.message || '',
      (function () { try { return JSON.stringify(e.details || {}); } catch (_) { return ''; } })()
    ];
  });

  if (!rows.length) rows.push([new Date(), 'info', 'NO_ERRORS', '', '', '', '', 'Aucune erreur', '{}']);

  sh.getRange(1, 1, 1, header.length).setValues([header]);
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.autoResizeColumns(1, header.length);
}

/* ===================== EXPORT XLSX (Groupe Articles SEUL) — avec filtre optionnel ===================== */

function exportRetroGroupeArticlesXlsxToDrive(seasonSheetId, options){
  var ss = getSeasonSpreadsheet_(seasonSheetId);

  // 0) ON/OFF incrémental via PARAMS
  var incrOn = String(readParam_(ss, 'INCREMENTAL_ON') || '1').toLowerCase();
  var allowIncr = (incrOn === '1' || incrOn === 'true' || incrOn === 'yes' || incrOn === 'oui');

  var out = buildRetroGroupeArticlesRows(seasonSheetId);

  // Filtrage incrémental (seulement si autorisé ET set non vide)
  var rows, filtered;
  if (allowIncr) {
    var touched = _ga_readTouchedPassportSet_(ss, options);
    rows = _ga_filterRowsByPassports_(out.rows, touched);
    filtered = rows.length !== out.rows.length;
  } else {
    rows = out.rows; // FULL export
    filtered = false;
  }

  // Classeur temporaire minimal
  var temp = SpreadsheetApp.create('Export temporaire - Import Retro Groupe Articles');
  var tmp = temp.getSheets()[0];
  tmp.setName('Export');

  var all = [out.header].concat(rows);
  if (typeof normalizePassportToText8_ === 'function') {
    for (var i = 1; i < all.length; i++) { if (all[i] && all[i].length) all[i][0] = normalizePassportToText8_(all[i][0]); }
  }
  if (all.length) {
    tmp.getRange(1, 1, all.length, out.nbCols).setValues(all);
    if (all.length > 1) tmp.getRange(2, 1, all.length - 1, 1).setNumberFormat('@');
  }
  SpreadsheetApp.flush();

  var url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(url, { headers:{ Authorization:'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HHmm');
  blob.setName('Export_Retro_Groupe_Articles_' + ts + (filtered ? '_INCR' : '') + '.xlsx');

  var folderId = readParam_(ss, PARAM_KEYS.RETRO_GART_EXPORTS_FOLDER_ID)
              || readParam_(ss, PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID) || '';
  var dest = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = dest.createFile(blob);

  DriveApp.getFileById(temp.getId()).setTrashed(true);
  appendImportLog_(ss, 'RETRO_GART_XLSX_OK', file.getName() + ' -> ' + dest.getName() +
                        ' (rows=' + rows.length + ', filtered=' + filtered + ')');
  return { fileId:file.getId(), name:file.getName(), rows: rows.length, filtered: filtered };
}

/* ========== Exposition facultative via Library ========== */
if (typeof Library !== 'undefined') {
Library.buildRetroGroupeArticlesRows = buildRetroGroupeArticlesRows;
Library.writeRetroGroupeArticlesSheet = writeRetroGroupeArticlesSheet;
Library.exportRetroGroupeArticlesXlsxToDrive = exportRetroGroupeArticlesXlsxToDrive;
}

/** Export Groupes-Articles (lignes additionnelles) depuis LEDGER
 * - Prend seulement les mappings Type=article SANS ExclusiveGroup (additionnels)
 * - Respecte RETRO_GART_REQUIRE_MAPPING (sinon fallback templates si tu veux)
 * - AllowOrphan: TRUE = autorise si JOUEURS absent (en pratique, JOUEURS couvre déjà articles-seulement)
 */
function buildRetroGroupesArticlesRowsFromLedger_(seasonSheetId){
  var ss = getSeasonSpreadsheet_(seasonSheetId);

  // Data
  var ledger = readSheetAsObjects_(ss.getId(), SHEETS.ACHATS_LEDGER).rows || [];
  var maps   = readSheetAsObjects_(ss.getId(), SHEETS.MAPPINGS).rows || [];
  var joueurs= readSheetAsObjects_(ss.getId(), 'JOUEURS').rows || [];

  var articleMaps = maps.filter(function(r){ return String(r['Type']||'').toLowerCase()==='article' && !String(r['Exclude']||'').trim(); });

  // Params
  var ignCsv     = (readParam_(ss, 'RETRO_GART_IGNORE_FEES_CSV') || '').toString();
  var eliteKeys  = (readParam_(ss, 'RETRO_GART_ELITE_KEYWORDS')  || readParam_(ss,'RETRO_GROUP_ELITE_KEYWORDS') || '').toString();
  var requireMap = String(readParam_(ss, 'RETRO_GART_REQUIRE_MAPPING')||'TRUE').toUpperCase()==='TRUE';

  // Helpers
  function _hasAny(s, csv){
    if (!csv) return false;
    var hay = _nrmLower_(s||'');
    return csv.split(',').some(function(w){ return hay.indexOf(_nrmLower_(w.trim()))>=0; });
  }
  function _ageU2_(ageBracket){
    var m = String(ageBracket||'').match(/U(\d+)/i);
    return m ? m[1] : '';
  }
  function _genreInit_(g){
    return (String(g||'').toUpperCase().charAt(0) || '');
  }
  function _fmt(tpl, j){
    return String(tpl||'')
      .replace(/{{\s*U2\s*}}/g, _ageU2_(j.AgeBracket))
      .replace(/{{\s*genreInitiale\s*}}/g, _genreInit_(j.Genre));
  }
  function _overlaps(br, umin, umax){
    var m = String(br||'').match(/U(\d+)\s*-\s*U?(\d+)/i);
    if (!m) return true;
    var a = Number(m[1]||0), b = Number(m[2]||0);
    return !(b < umin || a > umax);
  }

  // JOUEURS index (pour genre/age)
  var jByPass = {};
  joueurs.forEach(function(j){ var p=String(j['Passeport #']||j['Passeport']||j['PS']||'').trim(); if(p) jByPass[p]=j; });

  // Header minimal (même que Groupes)
  var HEADER = ["Identifiant unique","Catégorie","Équipe/Groupe"];
  var out = [];

  ledger.filter(_isActiveRow_).forEach(function(a){
    var pass = String(a['Passeport #']||'').trim(); if (!pass) return;
    var pass8 = (typeof normalizePassportToText8_==='function') ? normalizePassportToText8_(pass) : pass.replace(/\D/g,'').padStart(8,'0');

    var fee = a['Nom du frais'] || a['Frais'] || a['Produit'] || '';
    if (_hasAny(fee, ignCsv))    return; // ignorés
    if (_hasAny(fee, eliteKeys)) return; // élite

    var j = jByPass[pass] || {};
    var ginit = _genreInit_(j.Genre);
    var best = null, prio=-1;

    // Trouver mapping Article correspondant (non-exclusif)
    articleMaps.forEach(function(m){
      var alias = String(m['AliasContains']||'').trim();
      if (alias && _nrmLower_(fee).indexOf(_nrmLower_(alias))<0) return;

      // SANS ExclusiveGroup ici: on ne veut que les “additionnels”
      if (String(m['ExclusiveGroup']||'').trim()) return;

      var okG = (String(m['Genre']||'*')==='*' || _genreInit_(m['Genre'])===ginit);
      var okU = _overlaps(j.AgeBracket, Number(m['Umin']||0), Number(m['Umax']||99));
      if (!okG || !okU) return;

      var p = Number(m['Priority']||0);
      if (p>prio){ prio=p; best=m; }
    });

    if (!best){
      if (requireMap) return; // pas de mapping = pas d’output
      // sinon: fallback éventuel (peu recommandé) — on peut ignorer pour rester strict
      return;
    }

    var cat = _fmt(best['CategorieFmt'], j);
    var grp = _fmt(best['GroupeFmt'],    j);
    out.push([pass8, cat, grp]);
  });

  return { header: HEADER, rows: out, nbCols: HEADER.length };

}
