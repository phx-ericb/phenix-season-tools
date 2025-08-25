/**
 * retro_groupe_articles.gs — v0.8
 * - Exporte "Rétro - Groupe Articles" (10 colonnes)
 * - Unifie INSCRIPTIONS + ARTICLES (par passeport) pour compléter Nom/Prénom/DOB/Genre
 * - MAPPINGS unifiés (Type=article) pour Groupe/Catégorie (vars: U, U2, ageCat, genre/genreInitiale, article, saison, annee)
 * - Respecte CANCELLED/EXCLUDE_FROM_EXPORT + moteur de règles (RETRO_RULES_JSON)
 * - Collecte & écrit un onglet "Erreurs" (paramétrable)
 *
 * Colonnes exportées:
 *  "Identifiant unique","Nom","Prénom","Date de naissance","#","Couleur","Sous-groupe","Position","Équipe/Groupe","Catégorie"
 */

/* ===================== Param keys ===================== */
if (typeof PARAM_KEYS === 'undefined') { var PARAM_KEYS = {}; }
PARAM_KEYS.RETRO_GART_SHEET_NAME = PARAM_KEYS.RETRO_GART_SHEET_NAME || 'RETRO_GART_SHEET_NAME';
PARAM_KEYS.RETRO_GART_EXPORTS_FOLDER_ID = PARAM_KEYS.RETRO_GART_EXPORTS_FOLDER_ID || 'RETRO_GART_EXPORTS_FOLDER_ID';

PARAM_KEYS.RETRO_GART_IGNORE_FEES_CSV = PARAM_KEYS.RETRO_GART_IGNORE_FEES_CSV || 'RETRO_GART_IGNORE_FEES_CSV';     // sinon RETRO_IGNORE_FEES_CSV
PARAM_KEYS.RETRO_GART_ELITE_KEYWORDS = PARAM_KEYS.RETRO_GART_ELITE_KEYWORDS || 'RETRO_GART_ELITE_KEYWORDS';
PARAM_KEYS.RETRO_GART_REQUIRE_MAPPING = PARAM_KEYS.RETRO_GART_REQUIRE_MAPPING || 'RETRO_GART_REQUIRE_MAPPING';      // TRUE => ne sortir que les articles mappés

PARAM_KEYS.RETRO_RULES_JSON = PARAM_KEYS.RETRO_RULES_JSON || 'RETRO_RULES_JSON';                 // moteur de règles partagé

// Param erreurs
PARAM_KEYS.RETRO_ERRORS_SHEET_NAME = PARAM_KEYS.RETRO_ERRORS_SHEET_NAME || 'RETRO_ERRORS_SHEET_NAME';

// Adapté (pour exclure CDP0)
PARAM_KEYS.RETRO_ADAPTE_KEYWORDS = PARAM_KEYS.RETRO_ADAPTE_KEYWORDS || 'RETRO_ADAPTE_KEYWORDS';
PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS = PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS || 'RETRO_GROUP_SA_KEYWORDS';

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
  if (/^(m|masculin|male|man|garcon|gar\u00e7on|homme|boy)\b/.test(n) || /\bu ?\d+\s*m\b/.test(n) || /\bmasc\b/.test(n))
    return { label: 'Masculin', initiale: 'M' };
  if (/^(f|feminin|female|woman|fille|dame|girl)\b/.test(n) || /\bu ?\d+\s*f\b/.test(n) || /\bfem\b/.test(n))
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
  if (dn) {
    var m = String(dn).match(/(19|20)\d{2}/);
    if (m) return parseInt(m[0], 10);
  }
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
  var s = String(feeName || '').toUpperCase();
  var m = s.match(/U\s*[-–]?\s*(\d{1,2})/);
  return m ? ('U' + parseInt(m[1], 10)) : '';
}
function _ga_computeUandU2_(row, seasonYear, feeName) {
  var by = _ga_deriveAgeYear_(row);
  var U = '', U2 = '';
  if (by) {
    U2 = _ga_ageCat_(by, seasonYear);
    if (U2) U = 'U' + parseInt(U2.slice(1), 10);
  }
  if (!U) {
    var uTxt = _ga_extractUFromFeeText_(feeName);
    if (uTxt) {
      U = uTxt;
      var n = parseInt(uTxt.replace(/\D/g, ''), 10);
      if (!isNaN(n)) U2 = 'U' + _ga_pad2_(n);
    }
  }
  return { U: U, U2: U2 };
}

function _ga_tpl_(tpl, vars) {
  tpl = String(tpl == null ? '' : tpl);
  return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function (_, k) { return (vars && k in vars && vars[k] != null) ? String(vars[k]) : ''; });
}

/* ===== Règles ===== */
function _ga_loadRules_(ss) {
  if (typeof loadRetroRules_ === 'function') return loadRetroRules_(ss);
  return [];
}
function _ga_applyRowRulesMaybeSkip_(rules, articleRow, ctx) {
  if (!rules || !rules.length || typeof applyRetroRowRules_ !== 'function') return false;
  var fakeMember = {};
  var res = applyRetroRowRules_(rules, 'articles', articleRow, fakeMember, ctx);
  return !!(res && res.skip);
}

/* ===== Lecture MAPPINGS unifiés (incl. ExclusiveGroup) ===== */
function _loadUnifiedGroupMappings_(ss) {
  var sh = ss.getSheetByName(SHEETS.MAPPINGS);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var data = sh.getDataRange().getValues();
  var H = (data[0] || []).map(function (h) { return String(h || '').trim(); });
  function idx(k) { var i = H.indexOf(k); return i < 0 ? null : i; }
  var iType = idx('Type'), iAli = idx('AliasContains'), iUmin = idx('Umin'), iUmax = idx('Umax'),
    iGen = idx('Genre'), iG = idx('GroupeFmt'), iC = idx('CategorieFmt'), iEx = idx('Exclude'),
    iPr = idx('Priority'), iX = idx('ExclusiveGroup'), iCode = idx('Code');
  if (iType == null || iAli == null) return out;

  for (var r = 1; r < data.length; r++) {
    var row = data[r] || [];
    if (!row.some(function (c) { return String(c || '').trim(); })) continue;
    out.push({
      Type: String(row[iType] || '').toLowerCase(),                // member | article
      AliasContains: String(row[iAli] || ''),
      Umin: isNaN(parseInt(row[iUmin], 10)) ? null : parseInt(row[iUmin], 10),
      Umax: isNaN(parseInt(row[iUmax], 10)) ? null : parseInt(row[iUmax], 10),
      Genre: String(row[iGen] || '*').toUpperCase(),
      GroupeFmt: String(row[iG] || ''),
      CategorieFmt: String(row[iC] || ''),
      Exclude: String(row[iEx] || '').toLowerCase() === 'true',
      Priority: isNaN(parseInt(row[iPr], 10)) ? 100 : parseInt(row[iPr], 10),
      ExclusiveGroup: String(row[iX] || '').trim(),
      Code: String(row[iCode] || '').trim()
    });
  }
  // priorité: plus grand d'abord, puis Alias plus long
  out.sort(function (a, b) {
    if (a.Priority !== b.Priority) return b.Priority - a.Priority;
    return (b.AliasContains || '').length - (a.AliasContains || '').length;
  });
  return out;
}

function _low_(s) { s = String(s == null ? '' : s); try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return s.toLowerCase().trim(); }

/** Renvoie {passed:[], failedU:[]} pour un fee donné (type filtré + genre/alias ok) */
function _findArticleMappingCandidates_(maps, feeName, vars) {
  var s = _low_(feeName || '');
  var passed = [], failedU = [];
  for (var i = 0; i < maps.length; i++) {
    var m = maps[i]; if (m.Type !== 'article') continue;
    if (!m.AliasContains) continue;
    if (s.indexOf(_low_(m.AliasContains)) === -1) continue;
    if (m.Genre && m.Genre !== '*' && m.Genre !== (vars.genreInitiale || '')) continue;

    // filtre U
    var uNum = 0; if (vars.U) { var mm = String(vars.U).match(/^U(\d{1,2})$/i); if (mm) uNum = parseInt(mm[1], 10); }
    var okU = true;
    if (m.Umin != null && (!uNum || uNum < m.Umin)) okU = false;
    if (m.Umax != null && (!uNum || uNum > m.Umax)) okU = false;

    if (okU) passed.push(m); else failedU.push(m);
  }
  return { passed: passed, failedU: failedU };
}

/** Rend la 1ère règle qui passe, ou null; inclut exclusiveGroup pour analytique */
function _applyUnifiedMapping_(maps, feeName, vars) {
  var cand = _findArticleMappingCandidates_(maps, feeName, vars).passed;
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
  function ensure(p) {
    var k = String(p || '').trim(); if (!k) return null;
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

/* ===================== Construction + Erreurs ===================== */
function buildRetroGroupeArticlesRows(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var art = readSheetAsObjects_(ss.getId(), SHEETS.ARTICLES);

  var rules = _ga_loadRules_(ss);
  var mappings = _loadUnifiedGroupMappings_(ss);

  // Filtres
  var ignoreCsv = readParam_(ss, PARAM_KEYS.RETRO_GART_IGNORE_FEES_CSV) || readParam_(ss, 'RETRO_IGNORE_FEES_CSV') || 'senior,u-sé,adulte,ligue';
  var eliteCsv = readParam_(ss, PARAM_KEYS.RETRO_GART_ELITE_KEYWORDS) || 'D1+,LDP,Ligue';
  var requireMp = (String(readParam_(ss, PARAM_KEYS.RETRO_GART_REQUIRE_MAPPING) || 'TRUE').toUpperCase() === 'TRUE');

  // Adapté (pour exclure CDP0 warn)
  var adapteCsv = (readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || '') + ',' + (readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS) || '');
  adapteCsv = adapteCsv.replace(/^,|,$/g, '');

  var header = ["Identifiant unique", "Nom", "Prénom", "Date de naissance", "#", "Couleur", "Sous-groupe", "Position", "Équipe/Groupe", "Catégorie"];
  var rows = [];
  var errors = []; // {level: 'error'|'warn', code, passeport, nom, prenom, feeName, message, details}

  var activeArts = (art.rows || []).filter(_ga_isActiveArticle_);
  if (!activeArts.length) return { header: header, rows: rows, nbCols: header.length, errors: errors };

  // Saison/année
  var seasonLabel = readParam_(ss, 'SEASON_LABEL') || (activeArts[0] && activeArts[0]['Saison']) || '';
  var seasonYear = parseSeasonYear_(seasonLabel);

  var ctx = { ss: ss, catalog: (typeof _loadArticlesCatalog_ === 'function' ? _loadArticlesCatalog_(ss) : { match: function () { return null; } }) };

  var memberIdx = _ga_buildMemberIndex_(ss);

  // Set des passeports avec inscription active (pour "article sans inscription")
  var inscActivePass = {};
  (insc.rows || []).forEach(function (r) {
    var p = r['Passeport #']; if (!p) return;
    var can = String(r[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true';
    var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT] || '').toLowerCase() === 'true';
    var st = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
    var active = !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
    if (active) inscActivePass[String(p).trim()] = true;
  });

  var _normPass = (typeof normalizePassportPlain8_ === 'function')
    ? normalizePassportPlain8_
    : function (v) {
      var s = String(v == null ? '' : v).trim();
      if (!s) return '';
      if (/^\d+$/.test(s)) {
        var width = parseInt(readParam_(ss, 'PASSPORT_PAD_WIDTH') || '8', 10);
        if (isNaN(width) || width < 1) width = 8;
        s = (Array(width + 1).join('0') + s).slice(-width);
      }
      return s;
    };

  // Pour l’exclusivité & CDP0
  var perPassExclusive = {}; // pass -> { groupName -> [ {feeName, code} ] }
  var perPassIsAdapte = {}; // pass -> true si un de ses articles/inscriptions correspond aux mots-clés "Adapté"

  activeArts.forEach(function (a) {
    if (_ga_applyRowRulesMaybeSkip_(rules, a, ctx)) return;

    var feeName = a['Nom du frais'] || a['Frais'] || a['Produit'] || '';
    if (_ga_containsAny_(feeName, ignoreCsv)) return;
    if (_ga_containsAny_(feeName, eliteCsv)) return;

    var pass = a['Passeport #']; if (!pass) return;
    var passK = String(pass).trim();

    // Article sans inscription active ?
    if (!inscActivePass[passK]) {
      errors.push({
        level: 'error', code: 'ARTICLE_WITHOUT_INSCRIPTION', passeport: passK,
        nom: (a['Nom'] || ''), prenom: (a['Prénom'] || a['Prenom'] || ''), feeName: feeName,
        message: 'Article actif sans inscription active correspondante', details: {}
      });
      return; // on n’écrit pas la ligne
    }

    // Marqueur "Adapté" si repéré via articles actifs
    if (_ga_containsAny_(feeName, adapteCsv)) perPassIsAdapte[passK] = true;

    // Compléter depuis index membre
    var m = memberIdx[passK] || {};
    var nom = (a['Nom'] || '') || m.nom || '';
    var prenom = (a['Prénom'] || a['Prenom'] || '') || m.prenom || '';
    var dob = (a['Date de naissance'] || a['Naissance'] || '') || m.dob || '';

    // U/U2 + genre
    var UU2 = _ga_computeUandU2_({ 'Date de naissance': dob, 'Naissance': dob }, seasonYear, feeName);
    var U = UU2.U || '';
    var U2 = UU2.U2 || '';
    var gA = _ga_extractGenreSmart_(a);
    var gInit = gA.initiale || m.genreInit || '';
    var gLbl = gA.label || m.genreLabel || (gInit === 'F' ? 'Féminin' : (gInit === 'M' ? 'Masculin' : (gInit === 'X' ? 'Mixte' : '')));

    var vars = { U: U, U2: U2, ageCat: U2, genreInitiale: gInit, genre: gLbl, article: feeName, saison: seasonLabel, annee: seasonYear };

    // Candidats de mapping (pour AGE_OUT_OF_RANGE)
    var cands = _findArticleMappingCandidates_(mappings, feeName, vars);
    if (!cands.passed.length && cands.failedU.length) {
      // au moins un alias matche mais U est hors bornes
      var ranges = cands.failedU.map(function (m) {
        var a = []; if (m.Umin != null) a.push('min ' + m.Umin); if (m.Umax != null) a.push('max ' + m.Umax);
        return a.join(', ');
      }).join(' | ');
      errors.push({
        level: 'error', code: 'AGE_OUT_OF_RANGE', passeport: passK, nom: nom, prenom: prenom, feeName: feeName,
        message: 'Âge (U) hors bornes pour cet article', details: { U: U, ranges: ranges }
      });
      // on continue, car on peut décider de ne pas écrire si requireMp==true (ci-dessous)
    }

    // Application du mapping principal
    var mp = _applyUnifiedMapping_(mappings, feeName, vars);
    if (mp && mp.exclude) return;

    var groupe = (mp && mp.groupe) || '';
    var categ = (mp && mp.categorie) || '';
    var exg = (mp && mp.exclusiveGroup) || '';
    var code = (mp && mp.code) || '';

    // Exclusivité: accumule par passeport & groupe
    if (exg) {
      perPassExclusive[passK] = perPassExclusive[passK] || {};
      perPassExclusive[passK][exg] = perPassExclusive[passK][exg] || [];
      perPassExclusive[passK][exg].push({ feeName: feeName, code: code || feeName });
    }

    if (!mp && requireMp) return; // si exigé: uniquement les mappés

    if (!groupe && !categ) return; // rien à écrire

    var nbCols = header.length;
    var rowOut = new Array(nbCols).fill("");
    rowOut[0] = _normPass(pass);
    rowOut[1] = nom;
    rowOut[2] = prenom;
    rowOut[3] = dob;
    // #, Couleur, Sous-groupe, Position vides
    rowOut[8] = groupe;
    rowOut[9] = categ;

    rows.push(rowOut);
  });

  // Conflits d’exclusivité (ex.: CDP1+CDP2)
  Object.keys(perPassExclusive).forEach(function (passK) {
    var ex = perPassExclusive[passK];
    Object.keys(ex).forEach(function (group) {
      var arr = ex[group] || [];
      // conflit si >=2 codes distincts
      var distinct = {};
      arr.forEach(function (x) { distinct[String(x.code || '')] = true; });
      var nb = Object.keys(distinct).filter(Boolean).length;
      if (nb > 1) {
        errors.push({
          level: 'error', code: 'EXCLUSIVE_CONFLICT', passeport: passK, nom: '', prenom: '', feeName: '',
          message: 'Conflit d’exclusivité: plusieurs articles du groupe ' + group, details: { group: group, items: arr }
        });
      }
    });
  });

  // CDP0 (warn) pour U9–U12 non-Adapté
  // On regarde exclusivité "CDP" sur 9..12 : si aucune trace, warning
  Object.keys(inscActivePass).forEach(function (passK) {
    // U via membre index + (fallback articles) : on refait un U rapide via memberIdx
    var m = memberIdx[passK] || {};
    var UU2 = _ga_computeUandU2_({ 'Date de naissance': m.dob, 'Naissance': m.dob }, seasonYear, '');
    var U = UU2.U || '';
    var uNum = parseInt(String(U).replace(/^U/i, ''), 10);
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
      errors.push({
        level: 'warn', code: 'CDP0', passeport: passK, nom: (m.nom || ''), prenom: (m.prenom || ''), feeName: '',
        message: 'Membre U9–U12 sans CDP (1/2) — hors Adapté', details: { U: U }
      });
    }
  });

  return { header: header, rows: rows, nbCols: header.length, errors: errors };
}

/* ===================== Écriture feuilles ===================== */

function writeRetroGroupeArticlesSheet(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var out = buildRetroGroupeArticlesRows(seasonSheetId);

  // --- Data
  var sheetName = readParam_(ss, PARAM_KEYS.RETRO_GART_SHEET_NAME) || 'Rétro - Groupe Articles';
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clearContents();
  sh.getRange(1, 1, 1, out.nbCols).setValues([out.header]);
  if (out.rows.length) {
    sh.getRange(2, 1, out.rows.length, out.nbCols).setValues(out.rows);
    sh.autoResizeColumns(1, out.nbCols);
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 1).setNumberFormat('@');
      sh.getRange('A:A').setNumberFormat('@');
    }
  }
  appendImportLog_(ss, 'RETRO_GART_SHEET_OK', 'rows=' + out.rows.length);

  // --- Erreurs
  _writeRetroErrors_(ss, out.errors);

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

/** Export XLSX rapide — Rétro Groupe Articles */
function exportRetroGroupeArticlesXlsxToDrive(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var out = buildRetroGroupeArticlesRows(seasonSheetId); // {header, rows, nbCols, errors}

  // 1) Temp minimal
  var temp = SpreadsheetApp.create('Export temporaire - Retro Groupe Articles');
  var tmp = temp.getSheets()[0];
  tmp.setName('Export');

  // 2) Écriture header + data
  var all = [out.header].concat(out.rows);

  // Normalise Passeport -> texte si helper dispo
  if (typeof normalizePassportToText8_ === 'function') {
    for (var i = 1; i < all.length; i++) {
      all[i][0] = normalizePassportToText8_(all[i][0]);
    }
  }
  if (all.length) {
    tmp.getRange(1, 1, all.length, out.nbCols).setValues(all);
    if (all.length > 1) tmp.getRange(2, 1, all.length - 1, 1).setNumberFormat('@');
  }
  SpreadsheetApp.flush();

  // 3) Export XLSX
  var url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HHmm');
  blob.setName('Export_Retro_Groupe_Articles_' + ts + '.xlsx');

  // 4) Destination
  var folderId = readParam_(ss, PARAM_KEYS.RETRO_GART_EXPORTS_FOLDER_ID)
    || readParam_(ss, PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID) || '';
  var dest = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = dest.createFile(blob);

  // 5) Nettoyage + log
  DriveApp.getFileById(temp.getId()).setTrashed(true);
  appendImportLog_(ss, 'RETRO_GART_XLSX_OK_FAST', file.getName() + ' -> ' + dest.getName() + ' (rows=' + out.rows.length + ')');

  // 6) Écrit/MAJ l’onglet Erreurs
  _writeRetroErrors_(ss, out.errors);

  return { fileId: file.getId(), name: file.getName(), rows: out.rows.length, errors: out.errors.length };
}

/* ========== Exposition facultative via Library ========== */
if (typeof Library !== 'undefined') {
  Library.buildRetroGroupeArticlesRows = buildRetroGroupeArticlesRows;
  Library.writeRetroGroupeArticlesSheet = writeRetroGroupeArticlesSheet;
  Library.exportRetroGroupeArticlesXlsxToDrive = exportRetroGroupeArticlesXlsxToDrive;
}
