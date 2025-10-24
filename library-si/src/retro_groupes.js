/**
* retro_groupes.gs — v0.14
* - Post-process fort et tolérant :
*   • Si un passeport a une ligne Adapté → on NE GARDE QUE les lignes Adapté (toutes les autres sont purgées).
*   • Si un passeport n’a aucun « CDP » et qu’on détecte U9–U12 (formats variés) → préfixe Catégorie par "CDP ".
* - Logs de stats détaillées.
* - Reste inchangé pour le build « inscriptions » (structure/colonnes).
* - Compatible normalizePassportToText8_ / normalizePassportPlain8_
*/


/* ===================== Param keys ===================== */
if (typeof PARAM_KEYS === 'undefined') { var PARAM_KEYS = {}; }
PARAM_KEYS.RETRO_GROUP_SHEET_NAME = PARAM_KEYS.RETRO_GROUP_SHEET_NAME || 'RETRO_GROUP_SHEET_NAME';
PARAM_KEYS.RETRO_GROUP_EXPORTS_FOLDER_ID = PARAM_KEYS.RETRO_GROUP_EXPORTS_FOLDER_ID || 'RETRO_GROUP_EXPORTS_FOLDER_ID';

PARAM_KEYS.RETRO_IGNORE_FEES_CSV = PARAM_KEYS.RETRO_IGNORE_FEES_CSV || 'RETRO_IGNORE_FEES_CSV';
PARAM_KEYS.RETRO_GROUP_ELITE_KEYWORDS = PARAM_KEYS.RETRO_GROUP_ELITE_KEYWORDS || 'RETRO_GROUP_ELITE_KEYWORDS';

PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS = PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS || 'RETRO_GROUP_SA_KEYWORDS';
PARAM_KEYS.RETRO_GROUP_SA_GROUPE_LABEL = PARAM_KEYS.RETRO_GROUP_SA_GROUPE_LABEL || 'RETRO_GROUP_SA_GROUPE_LABEL';
PARAM_KEYS.RETRO_GROUP_SA_CATEG_LABEL = PARAM_KEYS.RETRO_GROUP_SA_CATEG_LABEL || 'RETRO_GROUP_SA_CATEG_LABEL';

PARAM_KEYS.RETRO_GROUP_GROUPE_FMT = PARAM_KEYS.RETRO_GROUP_GROUPE_FMT || 'RETRO_GROUP_GROUPE_FMT';
PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT = PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT || 'RETRO_GROUP_CATEGORIE_FMT';

PARAM_KEYS.RETRO_RULES_JSON = PARAM_KEYS.RETRO_RULES_JSON;


/* ===================== Cache règles ===================== */
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
    try { var arr = JSON.parse(raw); rules = Array.isArray(arr) ? arr : []; }
    catch (e) { appendImportLog_(ss, 'RETRO_RULES_JSON_PARSE_FAIL', String(e)); }
  }
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


/* ===================== Helpers ===================== */

function _rg_buildElitePassportSet_(ss) {
  var eliteCsv = readParam_(ss, PARAM_KEYS.RETRO_GROUP_ELITE_KEYWORDS) || 'D1+,CFP,LDP,Ligue,Ligue 2,Ligue 3';
  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var set = {};
  var rows = (insc.rows || []).filter(_rg_isActiveRow_);
  rows.forEach(function (r) {
    var fee = r['Nom du frais'] || r['Frais'] || r['Produit'] || '';
    if (_rg_containsAny_(fee, eliteCsv)) {
      var p = String(r['Passeport #'] || r['Passeport'] || '').trim();
      if (p) {
        var p8 = (typeof normalizePassportPlain8_ === 'function')
          ? normalizePassportPlain8_(p)
          : p.replace(/\D/g, '').padStart(8, '0');
        set[p8] = true;
      }
    }
  });
  return set;
}
// Passeports ayant une INSCRIPTION élite active (pas les articles)
function _rg_buildEliteInscPassportSet_(ss) {
  var eliteCsv = readParam_(ss, PARAM_KEYS.RETRO_GROUP_ELITE_KEYWORDS) || 'D1+,CFP,LDP,Ligue,Ligue 2,Ligue 3';
  var led = readSheetAsObjects_(ss.getId(), SHEETS.ACHATS_LEDGER);
  var set = {};
  var rows = (led.rows || []);

  function _truthy_(v){ var s=String(v||'').trim().toUpperCase(); return (s==='TRUE'||s==='OUI'||s==='YES'||s==='1'); }
  function _containsAny_(s,csv){
    if (!csv) return false;
    try { s = String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,''); } catch(e){}
    s = String(s||'').toLowerCase();
    return csv.split(',').some(function(w){ return s.indexOf(String(w||'').trim().toLowerCase()) >= 0; });
  }
  function _p8_(x){
    var s = String(x||'').trim(); if (!s) return '';
    return (typeof normalizePassportToText8_ === 'function') ? normalizePassportToText8_(s) : s.replace(/\D/g,'').padStart(8,'0');
  }

  for (var i=0;i<rows.length;i++){
    var r = rows[i];
    if (String(r['Type']||'').toUpperCase() !== 'INSCRIPTION') continue; // ← seulement INSCRIPTION
    if ((Number(r['Status'])||0) !== 1) continue;
    if (Number(r['isIgnored'])||0) continue;
    var name = r['NomFrais'] || r['Nom du frais'] || r['Frais'] || r['Produit'] || '';
    if (!_containsAny_(name, eliteCsv)) continue;

    var p8 = _p8_(r['Passeport #'] || r['Passeport'] || r['PS']);
    if (p8) set[p8] = true;
  }
  return set;
}


function _rg_nrm_(s) { s = String(s == null ? '' : s); try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return s; }
function _rg_low_(s) { return _rg_nrm_(s).toLowerCase().trim(); }
function _rg_pad2_(n) { n = Number(n || 0); return (n < 10 ? ('0' + n) : String(n)); }

function _rg_isActiveRow_(r) {
  var can = String(r[CONTROL_COLS.CANCELLED] || '').toLowerCase() === 'true';
  var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT] || '').toLowerCase() === 'true';
  var st = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
  return !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
}
function _rg_containsAny_(txt, csv) {
  var t = _rg_low_(txt || '');
  return String(csv || '').split(',').map(_rg_low_).filter(Boolean).some(function (p) { return t.indexOf(p) !== -1; });
}
function _rg_tpl_(tpl, vars) {
  tpl = String(tpl == null ? '' : tpl);
  return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function (_, k) { return (vars && k in vars && vars[k] != null) ? String(vars[k]) : ''; });
}
function _rg_csvEsc_(v) { v = v == null ? '' : String(v).replace(/"/g, '""'); return /[",\n;]/.test(v) ? ('"' + v + '"') : v; }
function _rg_genreInitiale_(row) {
  var g = (row['Identité de genre'] || row['Identité de Genre'] || row['Genre'] || row['Sexe'] || '').toString().trim().toUpperCase();
  return g ? g.charAt(0) : '';
}
function _tpl_(tpl, vars) { tpl = String(tpl == null ? '' : tpl); return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function (_, k) { return (vars && k in vars && vars[k] != null) ? String(vars[k]) : ''; }); }
function _low_(s) { try { s = String(s == null ? '' : s).normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return s.toLowerCase().trim(); }


/* ====== U & U2 (DOB + fallback libellé article) ====== */
function _rg_deriveBirthYearFromRow_(row) {
  var dn = row['Date de naissance'] || row['Naissance'] || '';
  if (dn instanceof Date) return dn.getFullYear();
  if (dn) {
    var m = String(dn).match(/(19|20)\d{2}/);
    if (m) return parseInt(m[0], 10);
  }
  return null;
}
function _rg_ageCat_(birthYear, seasonYear) {
  if (!birthYear || !seasonYear) return '';
  var age = seasonYear - birthYear;
  if (age < 4 || age > 99) return '';
  return 'U' + _rg_pad2_(age); // "U09"
}
function _rg_U_(birthYear, seasonYear) {
  var u2 = _rg_ageCat_(birthYear, seasonYear);
  return u2 ? ('U' + parseInt(u2.slice(1), 10)) : ''; // "U9", "U10", ...
}
function _rg_extractUFromArticle_(articleName) {
  var s = String(articleName || '').toUpperCase();
  var m = s.match(/U\s*[-–]?\s*(\d{1,2})/);
  return m ? ('U' + parseInt(m[1], 10)) : '';
}
function _rg_computeUandU2_(row, seasonYear, feeName) {
  var by = _rg_deriveBirthYearFromRow_(row);
  var U = '', U2 = '';
  if (by) {
    U2 = _rg_ageCat_(by, seasonYear);
    if (U2) U = 'U' + parseInt(U2.slice(1), 10);
  }
  if (!U) {
    var uTxt = _rg_extractUFromArticle_(feeName);
    if (uTxt) {
      U = uTxt;
      var n = parseInt(uTxt.replace(/\D/g, ''), 10);
      if (!isNaN(n)) U2 = 'U' + _rg_pad2_(n);
    }
  }
  return { U: U, U2: U2 };
}

/* ====== Fallback extraction "U.. genre" depuis article ====== */
function _rg_extractFromArticlePair_(articleName) {
  var s = String(articleName || '');
  var re = /U[-\s]?(\d{2}).*?(F[ée]minin|M[âa]sculin)/i;
  var m = s.match(re);
  if (m) {
    var u = 'U' + m[1];
    var g = m[2].toUpperCase().charAt(0);
    return { U: u, genreInitiale: g };
  }
  return null;
}

/* ===================== Règles (réutilise le moteur des membres) ===================== */
function _rg_loadRules_(ss) {
  if (typeof loadRetroRules_ === 'function') return loadRetroRules_(ss);
  return [];
}
function _rg_applyRowRulesMaybeSkip_(rules, row, ctx) {
  if (!rules || !rules.length || typeof applyRetroRowRules_ !== 'function') return false;
  var fakeMember = {};
  var res = applyRetroRowRules_(rules, 'inscriptions', row, fakeMember, ctx);
  return !!(res && res.skip);
}


/** Choisit la règle qui matche avec la PRIORITÉ la plus haute.
 *  À priorité égale, une règle Exclude gagne sur une non-Exclude.
 */
function _applyUnifiedMapping_(maps, type, feeName, vars) {
  var s = _low_(feeName || '');
  var best = null;
  var bestPrio = -1;

  for (var i = 0; i < maps.length; i++) {
    var m = maps[i];
    if (m.Type !== type) continue;
    if (!m.AliasContains) continue;
    if (s.indexOf(_low_(m.AliasContains)) === -1) continue;

    // Genre
    if (m.Genre && m.Genre !== '*' && m.Genre !== (vars.genreInitiale || '')) continue;
    // U (via vars.U)
    if (m.Umin != null || m.Umax != null) {
      var uNum = 0;
      if (vars.U) { var mm = String(vars.U).match(/^U(\d{1,2})$/i); if (mm) uNum = parseInt(mm[1], 10); }
      if (!uNum) continue;
      if (m.Umin != null && uNum < m.Umin) continue;
      if (m.Umax != null && uNum > m.Umax) continue;
    }

    var pr = Number(m.Priority || 0);
    var isEx = !!m.Exclude;

    if (best == null
      || pr > bestPrio
      || (pr === bestPrio && isEx && !best.exclude)) {
      best = isEx ? { exclude: true } : {
        groupe: m.GroupeFmt ? _tpl_(m.GroupeFmt, vars) : '',
        categorie: m.CategorieFmt ? _tpl_(m.CategorieFmt, vars) : ''
      };
      bestPrio = pr;
      best.exclude = !!isEx;
    }
  }
  return best;
}


/* ===================== Construction des lignes (INSCRIPTIONS) ===================== */
function buildRetroGroupesRows(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);

  var rules = _rg_loadRules_(ss);
  var mappings = _loadUnifiedGroupMappings_(ss);

  // (2) Ignorés & élite
  var ignoreCsv = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV) || 'senior,u-sé,adulte,ligue';
  var eliteCsv  = readParam_(ss, PARAM_KEYS.RETRO_GROUP_ELITE_KEYWORDS) || 'D1+,CFP,LDP,Ligue,Ligue 2,Ligue 3';
  var saCsv     = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS) || 'soccer adapté,soccer adapte';
  var saGrp     = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_GROUPE_LABEL) || 'Adapté (tous)';
  var saCat     = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_CATEG_LABEL)  || 'Adapté';

  var grpFmtDef = readParam_(ss, PARAM_KEYS.RETRO_GROUP_GROUPE_FMT)      || '{{U}}{{genreInitiale}}';
  var catFmtDef = readParam_(ss, PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT)    || '{{U}} {{genreInitiale}}';

  var header = ["Identifiant unique", "Nom", "Prénom", "Date de naissance", "#", "Couleur", "Sous-groupe", "Position", "Équipe/Groupe", "Catégorie"];
  var rows = [];

  var active = (insc.rows || []).filter(_rg_isActiveRow_);
  if (!active.length) return { header: header, rows: rows, nbCols: header.length };

  var seasonLabel = readParam_(ss, 'SEASON_LABEL') || (active[0] && active[0]['Saison']) || '';
  var seasonYear  = parseSeasonYear_(seasonLabel);

  // normalisation passeport – garde le même util que le reste du fichier
  var _normPass = (typeof normalizePassportPlain8_ === 'function')
    ? normalizePassportPlain8_
    : function (v) { return String(v == null ? '' : v); };

  active.forEach(function (r) {
    // (1) Règles
    var ctx = { ss: ss, catalog: (typeof _loadArticlesCatalog_ === 'function' ? _loadArticlesCatalog_(ss) : { match: function () { return null; } }) };
    if (_rg_applyRowRulesMaybeSkip_(rules, r, ctx)) return;

    var pass = r['Passeport #'] || r['Passeport'] || r['PS'];
    if (!pass) return;

    var feeName = r['Nom du frais'] || r['Frais'] || r['Produit'] || '';

    // (2) Ignorés & élite
    //    - ignoreCsv : filtre large (ex: adultes)
    //    - eliteCsv  : ICI on est sur la feuille INSCRIPTIONS → ne skippe QUE si c’est une INSCRIPTION élite (article élite ≠ concerné)
    if (_rg_containsAny_(feeName, ignoreCsv)) return;
    if (_rg_containsAny_(feeName, eliteCsv))  return;  // ← on n’ignore que ces inscriptions élite

    // (3) Soccer adapté ?
    if (_rg_containsAny_(feeName, saCsv)) {
      rows.push([
        _normPass(pass),
        (r['Nom'] || ''),
        (r['Prénom'] || r['Prenom'] || ''),
        (r['Date de naissance'] || r['Naissance'] || ''),
        "", "", "", "",
        saGrp,
        saCat
      ]);
      return;
    }

    // (4) U/U2 + genre
    var UU2 = _rg_computeUandU2_(r, seasonYear, feeName);
    var U  = UU2.U  || '';
    var U2 = UU2.U2 || '';
    var gi = _rg_genreInitiale_(r) || '';

    var vars = {
      U: U,
      U2: U2,
      ageCat: U2, // alias
      genreInitiale: gi,
      genre: (gi === 'F' ? 'Féminin' : (gi === 'M' ? 'Masculin' : (gi === 'X' ? 'Mixte' : ''))),
      saison: seasonLabel,
      annee: seasonYear,
      article: feeName
    };

    // (5) MAPPINGS saisonniers (prioritaires) — respecte Priority & Exclude via _applyUnifiedMapping_
    var mapRes = _applyUnifiedMapping_(mappings, 'member', feeName, vars);
    if (mapRes && mapRes.exclude) return;

    var groupe = (mapRes && mapRes.groupe)    || '';
    var categ  = (mapRes && mapRes.categorie) || '';

    // (6) Fallback: extraction depuis libellé article, sinon formats défaut
    if (!groupe || !categ) {
      var fromArt = _rg_extractFromArticlePair_(feeName);
      var Ux  = fromArt && fromArt.U              ? fromArt.U              : U;
      var gix = fromArt && fromArt.genreInitiale  ? fromArt.genreInitiale  : gi;

      // Régénérer U2 si besoin
      var U2x = U2;
      if (!U2x && Ux) {
        var n = parseInt(String(Ux).replace(/\D/g, ''), 10);
        if (!isNaN(n)) U2x = 'U' + _rg_pad2_(n);
      }

      var v2 = {
        U: Ux, U2: U2x, ageCat: U2x,
        genreInitiale: gix,
        genre: (gix === 'F' ? 'Féminin' : (gix === 'M' ? 'Masculin' : (gix === 'X' ? 'Mixte' : ''))),
        saison: seasonLabel, annee: seasonYear, article: feeName
      };
      if (!groupe) groupe = _rg_tpl_(grpFmtDef, v2);
      if (!categ)  categ  = _rg_tpl_(catFmtDef, v2);
    }

    // (7) Si on n'a toujours rien -> skip
    if (!groupe && !categ) return;

    rows.push([
      _normPass(pass),
      (r['Nom'] || ''),
      (r['Prénom'] || r['Prenom'] || ''),
      (r['Date de naissance'] || r['Naissance'] || ''),
      "", "", "", "",
      groupe,
      categ
    ]);
  });

  return { header: header, rows: rows, nbCols: header.length };
}



/* ===================== Post-processing CDP0 & Adapté ===================== */
/** Règles fortes appliquées juste avant d’écrire/exporter, quel que soit le builder. */
function _postProcessRowsForCDP0AndAdapted_(rows, eliteSet) {
  eliteSet = eliteSet || {};
  if (!rows || !rows.length) return { rows: [], stats: { adaptPassports: 0, removedOnAdapt: 0, prefixedCDP0: 0, passports: 0 } };

  // Index par passeport
  var byP = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i]; if (!r || !r.length) continue;
    var p = String(r[0] || '').trim(); if (!p) continue;
    (byP[p] = byP[p] || []).push(r);
  }

  function _nrm(s) { try { return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase(); } catch (e) { return String(s || '').toLowerCase(); } }
  function _hasAdaptMark(r) {
    var g = _nrm(r[8]), c = _nrm(r[9]);
    return (g.indexOf('adapte') >= 0 || c.indexOf('adapte') >= 0);
  }
  function _hasCDPMark(r) {
    var g = _nrm(r[8]), c = _nrm(r[9]);
    return (g.indexOf('cdp') >= 0 || c.indexOf('cdp') >= 0);
  }
  function _extractUNumFromText(txt) {
    var m = String(txt || '').toUpperCase().match(/U\s*[-]?\s*(\d{1,2})/);
    return m ? Number(m[1]) : 0;
  }

  var removedOnAdapt = 0, prefixedCDP0 = 0, adaptPassports = 0;
  var out = [];

  Object.keys(byP).forEach(function (p) {
    var lines = byP[p];
    var hasAdapt = lines.some(_hasAdaptMark);
    if (hasAdapt) {
      adaptPassports++;
      lines = lines.filter(_hasAdaptMark);
      removedOnAdapt += (byP[p].length - lines.length);
    }

    // CDP0 : si U9–U12, aucune ligne "CDP", ET pas élite → on préfixe
    var anyCDP = lines.some(_hasCDPMark);
    if (!anyCDP) {
      // déduire U depuis les textes déjà présents
      var Utexts = [];
      lines.forEach(function (r) { Utexts.push(String(r[8] || '')); Utexts.push(String(r[9] || '')); });
      var U_num = 0;
      for (var i = 0; i < Utexts.length; i++) { var u = _extractUNumFromText(Utexts[i]); if (u > 0) { U_num = u; break; } }

      if (U_num >= 9 && U_num <= 12) {
        // ⛔ élite ? → on ne force PAS CDP0
        if (eliteSet[p] === true) {
          // no-op
        } else {
          // prefix Catégorie par "CDP "
          lines.forEach(function (r) {
            var cat = String(r[9] || '');
            if (cat && _nrm(cat).indexOf('cdp') < 0) {
              r[9] = 'CDP ' + cat;
              prefixedCDP0++;
            }
          });
        }
      }
    }

    Array.prototype.push.apply(out, lines);
  });

  return { rows: out, stats: { adaptPassports: adaptPassports, removedOnAdapt: removedOnAdapt, prefixedCDP0: prefixedCDP0, passports: Object.keys(byP).length } };
}


/* ===================== FILTRAGE «passeports touchés» ===================== */
function _normalizePassportText8_(v) {
  var s = String(v == null ? '' : v).trim();
  try { if (typeof normalizePassportToText8_ === 'function') return normalizePassportToText8_(s); } catch (_) { }
  try { if (typeof normalizePassportPlain8_ === 'function') return normalizePassportPlain8_(s); } catch (_) { }
  return s;
}
function _readTouchedPassportSet_(ss, options) {
  options = options || {};
  var set = {};

  // 1) {onlyPassports: Array|Set}
  var list = options.onlyPassports;
  if (list && typeof list.forEach === 'function') {
    list.forEach(function (p) { var t = _normalizePassportText8_(p); if (t) set[t] = true; });
  }

  // 2) Fallback: DocumentProperties.LAST_TOUCHED_PASSPORTS (JSON ou CSV)
  if (!Object.keys(set).length) {
    try {
      var raw = (PropertiesService.getDocumentProperties().getProperty('LAST_TOUCHED_PASSPORTS') || '').trim();
      if (raw) {
        var arr = (raw.charAt(0) === '[') ? JSON.parse(raw) : raw.split(',');
        arr.forEach(function (p) { var t = _normalizePassportText8_(p); if (t) set[t] = true; });
      }
    } catch (_) { /* ignore */ }
  }
  return set; // peut être vide -> aucun filtrage
}

function _filterRowsByPassports_(rows, touchedSet) {
  var keys = Object.keys(touchedSet || {});
  if (!keys.length) return rows; // aucun filtrage
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var p = _normalizePassportText8_(row && row[0]); // col A = Passeport
    if (p && touchedSet[p]) out.push(row);
  }
  return out;
}


/* ===================== ÉCRITURE FEUILLE ===================== */
function writeRetroGroupesSheet(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var out = buildRetroGroupesRows(seasonSheetId);

  // Post-process fort avant écriture
var eliteInscSet = _rg_buildEliteInscPassportSet_(ss);
  var pp = _postProcessRowsForCDP0AndAdapted_(merged || [], eliteInscSet);
  appendImportLog_(ss, 'RETRO_GROUPES_POSTPROC_SHEET', JSON.stringify(pp.stats));

  var sheetName = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SHEET_NAME) || 'Rétro - Groupes';
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  sh.clearContents();
  sh.getRange(1, 1, 1, out.nbCols).setValues([out.header]);
  if (pp.rows.length) {
    sh.getRange(2, 1, pp.rows.length, out.nbCols).setValues(pp.rows);
    sh.autoResizeColumns(1, out.nbCols);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).setNumberFormat('@');
  }
  appendImportLog_(ss, 'RETRO_GROUPES_SHEET_OK', 'rows=' + pp.rows.length);
  return pp.rows.length;
}


/* ===================== EXPORT XLSX (Groupes SEUL) — avec filtre optionnel ===================== */
function exportRetroGroupesXlsxToDrive(seasonSheetId, options) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var out = buildRetroGroupesRows(seasonSheetId);

  // Post-process fort
var eliteInscSet = _rg_buildEliteInscPassportSet_(ss);
  var pp = _postProcessRowsForCDP0AndAdapted_(merged || [], eliteInscSet);

  appendImportLog_(ss, 'RETRO_GROUPES_POSTPROC_XLSX', JSON.stringify(pp.stats));



  // Filtrage incrémental
  var touched = _readTouchedPassportSet_(ss, options);
  var rows = _filterRowsByPassports_(pp.rows, touched);
  var filtered = rows.length !== pp.rows.length;

  // Temp minimal
  var temp = SpreadsheetApp.create('Export temporaire - Import Retro Groupes');
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
  var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HHmm');
  blob.setName('Export_Retro_Groupes_' + ts + (filtered ? '_INCR' : '') + '.xlsx');

  var folderId = readParam_(ss, PARAM_KEYS.RETRO_GROUP_EXPORTS_FOLDER_ID)
    || readParam_(ss, PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID) || '';
  var dest = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = dest.createFile(blob);

  DriveApp.getFileById(temp.getId()).setTrashed(true);
  appendImportLog_(ss, 'RETRO_GROUPES_XLSX_OK_FAST', file.getName() + ' -> ' + dest.getName() + ' (rows=' + rows.length + ')');
  return { fileId: file.getId(), name: file.getName(), rows: rows.length, filtered: filtered };
}


/** ===================== EXPORT XLSX Groupes ALL (Groupes + GroupeArticles) — avec filtre optionnel ===================== */
function exportRetroGroupesAllXlsxToDrive(seasonSheetId, options) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);

  // Param dédié pour Groupes, fallback sur RETRO_EXPORT_LAST_DAYS (global), sinon 0
  var groupsDaysRaw = readParam_(ss, 'RETRO_EXPORT_LAST_DAYS_GROUPES');
  var fallbackDays  = readParam_(ss, 'RETRO_EXPORT_LAST_DAYS'); // optionnel
  var windowDays = parseInt((groupsDaysRaw != null && groupsDaysRaw !== '') ? groupsDaysRaw : (fallbackDays || '0'), 10);
  var cutoffDate = (windowDays > 0) ? new Date(Date.now() - windowDays * 86400000) : null;

  var incrOn = String(readParam_(ss, 'INCREMENTAL_ON') || '1').toLowerCase();
  var allowIncr = (incrOn === '1' || incrOn === 'true' || incrOn === 'yes' || incrOn === 'oui');

  // Sources
  var base = buildRetroGroupesRows(seasonSheetId);
  var addl = (typeof buildRetroGroupeArticlesRows === 'function')
    ? buildRetroGroupeArticlesRows(seasonSheetId)
    : { header: base.header, rows: [], nbCols: base.nbCols };

  var header = (base && base.header) || (addl && addl.header) || [];
  var nbCols = header.length;

  function pick(obj, keys){ for (var i=0;i<keys.length;i++){ var k=keys[i]; if (Object.prototype.hasOwnProperty.call(obj,k)) return k; } return null; }
  function parseFlexibleDate_(v){
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(+v)) return v;
    if (typeof v === 'number') { var d = new Date(1899, 11, 30); return new Date(d.getTime() + v*86400000); }
    var s = String(v).trim();
    var d3 = new Date(s); if (!isNaN(+d3)) return d3;
    var m = s.match(/^(\d{2})[\/\-](\d{2})[\/\-](\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) { var dd=+m[1],MM=+m[2]-1,yyyy=+m[3]; var hh=+(m[4]||'0'),mi=+(m[5]||'0'),ss=+(m[6]||'0');
      var d4 = new Date(yyyy,MM,dd,hh,mi,ss); if (!isNaN(+d4)) return d4; }
    return null;
  }
  function normalizeP8_(p){ return String(p||'').replace(/\D/g,'').padStart(8,'0'); }

  var baseRows = base && base.rows ? base.rows.slice() : [];
  var addlRows = addl && addl.rows ? addl.rows.slice() : [];
  var filteredByWindow = false;

  if (windowDays > 0) {
    var led = readSheetAsObjects_(ss.getId(), 'ACHATS_LEDGER') || { header: [], rows: [] };
    var sample = (led.rows && led.rows[0]) ? led.rows[0] : {};
    var COL_PASS = pick(sample, ['Passeport #','Passeport','Passport','PS_Passport']);
    var COL_DATE = pick(sample, ['Date de la facture','Date Facture','Date facture','DateFacture','Date']);

    if (COL_PASS && COL_DATE) {
      var recent = new Set();
      for (var i=0;i<(led.rows||[]).length;i++){
        var L = led.rows[i];
        var d = parseFlexibleDate_(L[COL_DATE]);
        if (d && d >= cutoffDate) {
          var p8 = normalizeP8_(L[COL_PASS]);
          if (p8) recent.add(p8);
        }
      }
      function filterRowsOnP0_(arr){
        return (arr||[]).filter(function(r){
          var p = normalizeP8_(r && r[0]);
          return p && recent.has(p);
        });
      }
      baseRows = filterRowsOnP0_(baseRows);
      addlRows = filterRowsOnP0_(addlRows);
      filteredByWindow = true;

      try { appendImportLog_(ss, 'RETRO_GROUPES_WINDOW_DIAG', JSON.stringify({ windowDays: windowDays, recentSize: recent.size })); } catch(e){}
    } else {
      try { appendImportLog_(ss, 'RETRO_GROUPES_WINDOW_SKIP', JSON.stringify({ reason: 'missing_cols' })); } catch(e){}
    }
  }

  var merged = [].concat(baseRows, addlRows);
  var pp = _postProcessRowsForCDP0AndAdapted_(merged || []);
  try { appendImportLog_(ss, 'RETRO_GROUPES_POSTPROC_ALL_MERGED', JSON.stringify(pp.stats)); } catch(e){}

  var rowsFiltered = pp.rows;
  var filtered = false;

  if (!filteredByWindow && allowIncr) {
    var touched = _readTouchedPassportSet_(ss, options);
    rowsFiltered = _filterRowsByPassports_(pp.rows, touched);
    filtered = (rowsFiltered.length !== pp.rows.length);
  } else {
    filtered = filteredByWindow && (rowsFiltered.length !== (base.rows.length + addl.rows.length));
  }

  var temp = SpreadsheetApp.create('Export temporaire - Import Retro Groupes All');
  var tmp = temp.getSheets()[0];
  tmp.setName('Export');

  var all = [header].concat(rowsFiltered);
  if (typeof normalizePassportToText8_ === 'function') {
    for (var r = 1; r < all.length; r++) {
      if (all[r] && all[r].length) all[r][0] = normalizePassportToText8_(all[r][0]);
    }
  }
  if (all.length) {
    tmp.getRange(1, 1, all.length, nbCols).setValues(all);
    if (all.length > 1) tmp.getRange(2, 1, all.length - 1, 1).setNumberFormat('@');
  }
  SpreadsheetApp.flush();

  var url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HHmm');
  blob.setName('Export_Retro_Groupes_All_' + ts + (filtered ? '_INCR' : '') + '.xlsx');

  var folderId = readParam_(ss, PARAM_KEYS.RETRO_GROUP_EXPORTS_FOLDER_ID)
              || readParam_(ss, PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID) || '';
  var dest = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = dest.createFile(blob);

  DriveApp.getFileById(temp.getId()).setTrashed(true);
  try { appendImportLog_(ss,'RETRO_GROUPES_ALL_XLSX_OK',
      file.getName()+' -> '+dest.getName()+' (rows='+rowsFiltered.length+', filtered='+filtered+')'); } catch(e){}
  return { fileId:file.getId(), name:file.getName(), rows:rowsFiltered.length, filtered:filtered };
}



/* ========== Exposition facultative via Library ========== */
if (typeof Library !== 'undefined') {
  Library.buildRetroGroupesRows = buildRetroGroupesRows;
  Library.writeRetroGroupesSheet = writeRetroGroupesSheet;
  Library.exportRetroGroupesXlsxToDrive = exportRetroGroupesXlsxToDrive;
  Library.exportRetroGroupesAllXlsxToDrive = exportRetroGroupesAllXlsxToDrive;
}


/** Export Groupes (principal) depuis JOUEURS, avec override par articles exclusifs du LEDGER
 *  (Conservé; pas utilisé par les exports ci-dessus, mais dispo si besoin.)
 */
function buildRetroGroupesRowsFromJoueursLedger_(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var saisonLbl = readParam_(ss, 'SEASON_LABEL') || '';

  // === Data ===
  var joueurs = readSheetAsObjects_(ss.getId(), 'JOUEURS').rows || [];
  var ledger  = readSheetAsObjects_(ss.getId(), SHEETS.ACHATS_LEDGER).rows || [];
  var maps    = readSheetAsObjects_(ss.getId(), SHEETS.MAPPINGS).rows || [];

  var memberMaps  = maps.filter(function (r) { return String(r['Type'] || '').toLowerCase() === 'member'  && !String(r['Exclude'] || '').trim(); });
  var articleMaps = maps.filter(function (r) { return String(r['Type'] || '').toLowerCase() === 'article' && !String(r['Exclude'] || '').trim(); });

  // Params
  var eliteKeys = (readParam_(ss, 'RETRO_GROUP_ELITE_KEYWORDS') || '').toString().trim();
  if (!eliteKeys) eliteKeys = 'D1+,CFP,LDP,Ligue 2,Ligue 3'; // fallback strict
  var saLabelG  = readParam_(ss, 'RETRO_GROUP_SA_GROUPE_LABEL') || 'Adapté (tous)';
  var saLabelC  = readParam_(ss, 'RETRO_GROUP_SA_CATEG_LABEL')  || 'Adapté';
  var saCsv     = (readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS) || 'soccer adapté,soccer adapte').toString();

  // Helpers
  function _nrmLower_(s) { try { s = String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) { } return String(s||'').toLowerCase(); }
  function _ageU2_(ageBracket) { var m = String(ageBracket || '').match(/U(\d+)/i); return m ? ('U' + m[1]) : ''; }
  function _uNum_(ageBracket)  { var m = String(ageBracket || '').match(/U(\d+)/i); return m ? Number(m[1]) : 0; }
  function _genreInit_(g)      { return (String(g || '').toUpperCase().charAt(0) || ''); }
  function _fmt(tpl, j) {
    return String(tpl || '')
      .replace(/{{\s*U2\s*}}/g, _ageU2_(j.AgeBracket))
      .replace(/{{\s*genreInitiale\s*}}/g, _genreInit_(j.Genre));
  }
  function _overlaps(br, umin, umax) {
    var m = String(br || '').match(/U(\d+)\s*-\s*U?(\d+)/i);
    if (!m) return true;
    var a = Number(m[1] || 0), b = Number(m[2] || 0);
    return !(b < umin || a > umax);
  }
  function _hasAny(s, csv) {
    if (!csv) return false;
    var hay = _nrmLower_(s || '');
    return csv.split(',').some(function (w) { return hay.indexOf(_nrmLower_(String(w).trim())) >= 0; });
  }
  function _isTrue_(v) {
    var s = String(v == null ? '' : v).trim().toLowerCase();
    return (s === '1' || s === 'true' || s === 'oui' || s === 'yes');
  }

  // Index LEDGER actifs par passeport (mêmes critères que le reste : Status=1, !isIgnored, Saison=SEASON_LABEL)
  var legByPass = {};
  ledger.forEach(function (a) {
    if (String(a['Saison'] || '') !== saisonLbl) return;
    if ((Number(a['Status']) || 0) !== 1) return;
    if ((Number(a['isIgnored']) || 0) === 1) return;
    var p = String(a['Passeport #'] || a['Passeport'] || a['PS'] || '').trim(); if (!p) return;
    (legByPass[p] = legByPass[p] || []).push(a);
  });

  // Sets dérivés du LEDGER (INSCRIPTION uniquement)
  var eliteInscByPass  = {}; // passeports avec INSCRIPTION élite active
  var normalInscByPass = {}; // passeports avec INSCRIPTION "saison" active ET non-élite

  Object.keys(legByPass).forEach(function (p) {
    var L = legByPass[p];
    for (var i = 0; i < L.length; i++) {
      var t = String(L[i]['Type'] || '').toUpperCase();
      if (t !== 'INSCRIPTION') continue;
      var fee = L[i]['NomFrais'] || L[i]['Nom du frais'] || L[i]['Frais'] || L[i]['Produit'] || '';
      var isElite = _hasAny(fee, eliteKeys);
      var isSeason = /saison/i.test(fee);
      if (isElite) {
        eliteInscByPass[p] = true;
      } else if (isSeason) {
        normalInscByPass[p] = true;
      }
    }
  });

  // Détection Adapté robuste (champs JOUEUR + mots-clés dans LEDGER)
  function _isAdaptedJ_(j) {
    if (_isTrue_(j.isAdapte) || _isTrue_(j['Adapté']) || _isTrue_(j['Programme adapté']) || _isTrue_(j['Adapte'])) return true;
    var p = String(j['Passeport #'] || j['Passeport'] || '').trim();
    var L = legByPass[p] || [];
    for (var i = 0; i < L.length; i++) {
      var fee = L[i]['Nom du frais'] || L[i]['Frais'] || L[i]['Produit'] || '';
      if (_hasAny(fee, saCsv)) return true;
    }
    return false;
  }

  // Sélection mapping member le + prioritaire
  function _selectMemberMap_(j) {
    var g = _genreInit_(j.Genre);
    var best = null, prio = -1;
    memberMaps.forEach(function (m) {
      var okG = (String(m['Genre'] || '*') === '*' || _genreInit_(m['Genre']) === g);
      var okU = _overlaps(j.AgeBracket, Number(m['Umin'] || 0), Number(m['Umax'] || 99));
      if (!okG || !okU) return;
      var p = Number(m['Priority'] || 0);
      if (p > prio) { prio = p; best = m; }
    });
    return best;
  }

  // Overrides exclusifs (famille ExclusiveGroup) basés sur les ARTICLES (on ignore les ARTICLES élite)
  function _selectExclusiveOverrides_(j) {
    var p = String(j['Passeport #'] || j['Passeport'] || '').trim();
    var L = legByPass[p] || [];
    if (!L.length) return {};
    var g = _genreInit_(j.Genre);
    var mapByFamily = {};
    L.forEach(function (a) {
      if (String(a['Type'] || '').toUpperCase() === 'INSCRIPTION') return; // overrides = articles
      var fee = a['Nom du frais'] || a['Frais'] || a['Produit'] || '';
      if (_hasAny(fee, eliteKeys)) return; // un article élite de détection ne crée pas d’override
      articleMaps.forEach(function (m) {
        var alias = String(m['AliasContains'] || '').trim();
        if (alias && _nrmLower_(fee).indexOf(_nrmLower_(alias)) < 0) return;
        var okG = (String(m['Genre'] || '*') === '*' || _genreInit_(m['Genre']) === g);
        var okU = _overlaps(j.AgeBracket, Number(m['Umin'] || 0), Number(m['Umax'] || 99));
        var fam = String(m['ExclusiveGroup'] || '').trim();
        if (!okG || !okU || !fam) return;
        var pz = Number(m['Priority'] || 0);
        if (!mapByFamily[fam] || pz > mapByFamily[fam].prio) {
          mapByFamily[fam] = { map: m, prio: pz };
        }
      });
    });
    return mapByFamily;
  }

  // Header export
  var HEADER = ["Identifiant unique", "Catégorie", "Équipe/Groupe"];
  var out = [];

  joueurs.forEach(function (j) {
    var pass = j['Passeport #'] || j['Passeport'] || j['PS'] || '';
    pass = (typeof normalizePassportToText8_ === 'function') ? normalizePassportToText8_(pass) : String(pass || '').replace(/\D/g, '').padStart(8, '0');
    if (!pass) return;

    // Adapté : sortie immédiate
    if (_isAdaptedJ_(j)) {
      out.push([pass, saLabelC, saLabelG]);
      return;
    }

    // ⛔️ INSCRIPTION élite active → pas de ligne dans rétro_groupes
    if (eliteInscByPass[pass] === true) {
      return;
    }

    // Base via mapping member
    var mm  = _selectMemberMap_(j);
    var cat = mm ? _fmt(mm['CategorieFmt'], j) : _fmt(readParam_(ss, 'RETRO_GROUP_CATEGORIE_FMT') || '{{U2}} {{genreInitiale}}', j);
    var grp = mm ? _fmt(mm['GroupeFmt'], j)    : _fmt(readParam_(ss, 'RETRO_GROUP_GROUPE_FMT')   || '{{U2}}{{genreInitiale}}', j);

    // Overrides exclusifs via ARTICLES
    var fams = _selectExclusiveOverrides_(j);
    Object.keys(fams).forEach(function (f) {
      var m = fams[f].map;
      if (m) { cat = _fmt(m['CategorieFmt'], j); grp = _fmt(m['GroupeFmt'], j); }
    });

    // CDP0 (Ledger) — si U9–U12, aucun override "CDP", ET inscription "saison" normale active → préfixer Catégorie
    var u = _uNum_(j.AgeBracket);
    var hasCDPOverride = Object.keys(fams).some(function (f) { return String(f || '').toUpperCase().indexOf('CDP') >= 0; });
    if (u >= 9 && u <= 12 && !hasCDPOverride && normalInscByPass[pass] === true) {
      if (cat.toUpperCase().indexOf('CDP ') !== 0) cat = 'CDP ' + cat;
    }

    out.push([pass, cat, grp]);
  });

  return { header: HEADER, rows: out, nbCols: HEADER.length };
}
