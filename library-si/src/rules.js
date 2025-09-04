/** rules.js — v0.8.1 (lib)
 * - Exclusion robuste des frais "Entraîneurs" des règles joueurs.
 * - Wrapper d'ignore (utilise SR_isIgnoredFeeRetro_ si dispo, sinon fallback local).
 * - Zéro dépendance à buildArticleKey_, clés recalculées localement.
 * - Fallbacks prudents pour exécuter en bibliothèque.
 */

/* ====== Fallbacks légers (seulement si absents dans l'environnement) ====== */
if (typeof CONTROL_COLS === 'undefined') {
  var CONTROL_COLS = { CANCELLED: '__cancelled', EXCLUDE_FROM_EXPORT: '__exclude_from_export' };
}
if (typeof normalizePassportPlain8_ !== 'function') {
  function normalizePassportPlain8_(p) { return String(p||'').replace(/\D/g,'').slice(-8).padStart(8,'0'); }
}
if (typeof normalizePassportToText8_ !== 'function') {
  function normalizePassportToText8_(p) { return normalizePassportPlain8_(p); }
}
if (typeof deriveUFromRow_ !== 'function') {
  function parseSeasonYear_(s){ var m=(String(s||'').match(/(20\d{2})/)); return m?parseInt(m[1],10):(new Date()).getFullYear(); }
  function birthYearFromRow_(row){
    var y=row['Année de naissance']||row['Annee de naissance']||row['Annee']||'';
    if (y && /^\d{4}$/.test(String(y))) return parseInt(y,10);
    var dob=row['Date de naissance']||''; if (dob){
      var s=String(dob), m=s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if(m) return parseInt(m[1],10);
      var m2=s.match(/(\d{1,2})[-/](\d{1,2})[-/](\d{4})/); if(m2) return parseInt(m2[3],10);
    }
    return null;
  }
  function computeUForYear_(by, sy){ if(!by||!sy) return null; var u=sy-by; return (u>=4&&u<=21)?('U'+u):null; }
  function deriveUFromRow_(row){
    var cat=row['Catégorie']||row['Categorie']||''; if (cat && /^U\d{1,2}/i.test(cat)) return cat.toUpperCase().replace(/\s+/g,'');
    var U = computeUForYear_(birthYearFromRow_(row), parseSeasonYear_(row['Saison']||'')); return U||'';
  }
}
if (typeof listActiveOccurrencesForPassport_ !== 'function') {
  function listActiveOccurrencesForPassport_(ss, sheetName, pass){ return []; } // debug-friendly no-op
}

/* ====== Normalisations communes ====== */
function RL_norm_(s){
  return String(s||'').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}

/* ====== Ignore rétro (fallback local) ====== */
function RL_isIgnoredFeeRetroLocal_(ss, fee){
  var v = RL_norm_(fee); if (!v) return false;
  var csv1 = (typeof readParam_==='function' ? (readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV)||'') : '');
  var csv2 = (typeof readParam_==='function' ? (readParam_(ss, 'RETRO_COACH_FEES_CSV')||'') : '');
  var toks = (csv1 + (csv1&&csv2?',':'') + csv2).split(',').map(RL_norm_).filter(Boolean);
  if (toks.indexOf(v) >= 0) return true;      // exact
  for (var i=0;i<toks.length;i++)             // contains
    if (v.indexOf(toks[i]) >= 0) return true;
  if (/(entraineur|entra[îi]neur|coach)/i.test(String(fee||''))) return true; // filet lexical
  return false;
}
/** Wrapper: si SR_isIgnoredFeeRetro_ existe (serveur), on l’utilise; sinon fallback local */
function isIgnoredFeeRetro_(ss, fee){
  try { if (typeof SR_isIgnoredFeeRetro_ === 'function') return SR_isIgnoredFeeRetro_(ss, fee); } catch(_){}
  return RL_isIgnoredFeeRetroLocal_(ss, fee);
}

/* ====== Détection "Coach" (lib) ====== */
function getCoachKeywordsCsv_(ss) {
  var csv = '';
  try { if (typeof readParam_ === 'function') csv = readParam_(ss, 'RETRO_COACH_FEES_CSV') || ''; } catch(_){}
  if (!csv) csv = 'Entraîneurs, Entraineurs, Entraîneur, Entraineur, Coach, Coaches';
  return csv;
}
function isCoachFeeByName_(ss, rawName) {
  var v = RL_norm_(rawName||''); if (!v) return false;
  var toks = getCoachKeywordsCsv_(ss).split(',').map(RL_norm_).filter(Boolean);
  if (toks.indexOf(v) >= 0) return true;
  for (var i=0;i<toks.length;i++) if (v.indexOf(toks[i]) >= 0) return true;
  if (/(entraineur|entra[îi]neur|coach)/i.test(String(rawName||''))) return true;
  return false;
}
function isCoachMember_(ss, row) {
  var name = (row && (row['Nom du frais']||row['Frais']||row['Produit'])) || '';
  return isCoachFeeByName_(ss, name);
}

/* ====== Clés d’agrégat ====== */
function _psKey_(row) {
  var p8 = normalizePassportPlain8_(row['Passeport #'] || row['Passeport']);
  var s  = String(row['Saison'] || '');
  return p8 + '||' + s;
}
function _articleDupKey_(row, match, rawFrais) {
  var base;
  if (match && match.Code) base = 'CODE:' + String(match.Code);
  else if (match && match.ExclusiveGroup) base = 'EXG:' + String(match.ExclusiveGroup);
  else base = 'LBL:' + String(rawFrais||'').toLowerCase().replace(/\s+/g,' ').trim();
  return _psKey_(row) + '||' + base;
}

/* ====== Catalogue (mappings ARTICLES) ====== */
function loadArticlesCatalog_(ss){
  var sh = getSheetOrCreate_(ss, SHEETS.MAPPINGS);
  var items = [];
  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues();
    var H = (data[0]||[]).map(function(h){ return String(h||'').trim(); });
    function idx(k){ var i=H.indexOf(k); return i<0 ? null : i; }

    // entête moderne
    var iType = idx('Type'), iAli = idx('AliasContains');
    if (iType != null && iAli != null) {
      var iUmin = idx('Umin'), iUmax = idx('Umax'), iCode = idx('Code'), iExGrp = idx('ExclusiveGroup');
      for (var r=1; r<data.length; r++){
        var row = data[r]||[];
        if (String(row[iType]||'').trim().toLowerCase() !== 'article') continue;
        var alias = String(row[iAli]||'').trim(); if (!alias) continue;
        var umin = (iUmin==null) ? null : parseInt(row[iUmin]||'',10); if (isNaN(umin)) umin=null;
        var umax = (iUmax==null) ? null : parseInt(row[iUmax]||'',10); if (isNaN(umax)) umax=null;
        var code = (iCode==null) ? '' : String(row[iCode]||'').trim();
        var excl = (iExGrp==null) ? '' : String(row[iExGrp]||'').trim();
        items.push({ Code: code, AliasContains: alias, Umin: umin, Umax: umax, ExclusiveGroup: excl });
      }
      return {
        items: items,
        match: function(raw){
          raw = String(raw||'').toLowerCase();
          for (var i=0;i<items.length;i++){
            var a = items[i].AliasContains.toLowerCase();
            if (a && raw.indexOf(a) !== -1) return items[i];
          }
          return null;
        }
      };
    }

    // compat ancien
    var iAli2 = idx('Alias') != null ? idx('Alias') : idx('AliasContains');
    var iUmin2 = idx('Umin'), iUmax2 = idx('Umax'), iCode2 = idx('Code'), iExGrp2 = idx('ExclusiveGroup');
    for (var r2=1; r2<data.length; r2++){
      var row2 = data[r2]||[];
      var alias2 = String(row2[iAli2]||'').trim(); if (!alias2) continue;
      var umin2 = (iUmin2==null) ? null : parseInt(row2[iUmin2]||'',10); if (isNaN(umin2)) umin2=null;
      var umax2 = (iUmax2==null) ? null : parseInt(row2[iUmax2]||'',10); if (isNaN(umax2)) umax2=null;
      var code2 = (iCode2==null) ? '' : String(row2[iCode2]||'').trim();
      var excl2 = (iExGrp2==null) ? '' : String(row2[iExGrp2]||'').trim();
      items.push({ Code: code2, AliasContains: alias2, Umin: umin2, Umax: umax2, ExclusiveGroup: excl2 });
    }
  }
  return {
    items: items,
    match: function(raw){
      raw = String(raw||'').toLowerCase();
      for (var i=0;i<items.length;i++){
        var a = items[i].AliasContains.toLowerCase();
        if (a && raw.indexOf(a) !== -1) return items[i];
      }
      return null;
    }
  };
}

/* ========================================================================== */
/*                           ÉVALUATION DES RÈGLES                             */
/* ========================================================================== */
function evaluateSeasonRules(seasonSheetId, filterPassports) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  ensureCoreSheets_(ss);

  var rulesOn = (readParam_(ss, PARAM_KEYS.RULES_ON) || 'TRUE').toUpperCase() === 'TRUE';
  if (!rulesOn) { appendImportLog_(ss, 'RULES_SKIP', 'RULES_ON=FALSE'); return {found:0, errors:0, warns:0, filtered:false}; }

  var dryRun = (readParam_(ss, PARAM_KEYS.RULES_DRY_RUN) || 'FALSE').toUpperCase() === 'TRUE';
  var threshold = (readParam_(ss, PARAM_KEYS.RULES_SEVERITY_THRESHOLD) || 'warn').toLowerCase();
  var appendMode = (readParam_(ss, PARAM_KEYS.RULES_APPEND) || 'FALSE').toUpperCase() === 'TRUE';

  // Lecture des tables
  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var art  = readSheetAsObjects_(ss.getId(), SHEETS.ARTICLES);

  // Filtre optionnel par passeports (normalisés en 8)
  var filterSet = null;
  if (Array.isArray(filterPassports) && filterPassports.length) {
    filterSet = Object.create(null);
    filterPassports.forEach(function(p){
      var raw = (p == null ? '' : String(p)).trim();
      if (!raw) return;
      try { raw = normalizePassportPlain8_(raw); } catch(_){}
      if (raw) filterSet[raw] = true;
    });
  }
  function rowPassInFilter_(row){
    if (!filterSet) return true;
    var pass = row['Passeport #'] || row['Passeport'] || '';
    var p8 = (typeof normalizePassportPlain8_ === 'function') ? normalizePassportPlain8_(pass) : String(pass||'').trim();
    return !!filterSet[p8];
  }
  function isActive_(r){
    var can = String(r[CONTROL_COLS.CANCELLED]||'').toLowerCase()==='true';
    var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT]||'').toLowerCase()==='true';
    var st  = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
    return !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
  }

  // Lignes actives + filtre passeport
  var inscAct = (insc.rows || []).filter(isActive_).filter(rowPassInFilter_);
  var artAct  = (art.rows  || []).filter(isActive_).filter(rowPassInFilter_);

  // === Exclusion globale des Coachs pour l'évaluation des règles joueurs
  var inscPlay = inscAct.filter(function(r){ return !isCoachMember_(ss, r); });
  var artPlay  = artAct.filter(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    return !isCoachFeeByName_(ss, raw);
  });

  // Catalogue
  var catalog = loadArticlesCatalog_(ss);

  // Feuille ERREURS
  var shErr = getSheetOrCreate_(ss, SHEETS.ERREURS,
    ['Passeport','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']
  );

  if (!appendMode && !dryRun) {
    if (!filterSet) {
      shErr.clearContents();
      shErr.getRange(1,1,1,12).setValues([['Passeport','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']]);
      appendImportLog_(ss, 'RULES_CLEAR_FULL', 'ERREURS reset (append=FALSE, no filter)');
    } else {
      var last = shErr.getLastRow();
      if (last >= 2) {
        var rng = shErr.getRange(2,1,last-1,12).getValues();
        var toDel = [];
        for (var i=0;i<rng.length;i++){
          var pass = String(rng[i][0]||'').trim();
          if (pass && filterSet[pass]) toDel.push(2+i);
        }
        toDel.sort(function(a,b){return b-a;}).forEach(function(r){ shErr.deleteRow(r); });
        appendImportLog_(ss, 'RULES_CLEAR_FILTERED', 'ERREURS ciblé passeports=' + Object.keys(filterSet).length + ', deleted=' + toDel.length);
      }
    }
  }
  shErr.getRange('A:A').setNumberFormat('@'); // passeport textuel

  function dict_(){ return Object.create(null); }
  function shouldWrite_(sev) {
    var sevRank = { warn:1, error:2 };
    return (sevRank[(sev||'warn')] || 1) >= (sevRank[threshold] || 1);
  }
  var errBuf = [];
  var found = 0, errors = 0, warns = 0;
  var mapInscByKey = dict_();

  // Adapté (fallback local)
  function isAdapteMember_(row) {
    try {
      var raw = (row['Nom du frais'] || row['Frais'] || row['Produit'] || '').toString();
      var txt = raw.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
      var baseCsv = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || '';
      var extraCsv = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS) || '';
      var csv = (baseCsv + ',' + extraCsv).replace(/^,|,$/g, '');
      var keys = csv.split(',').map(function (k) {
        k = (k || '').trim();
        try { k = k.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (_) {}
        return k.toLowerCase();
      }).filter(Boolean);
      for (var i = 0; i < keys.length; i++) if (keys[i] && txt.indexOf(keys[i]) !== -1) return true;
    } catch (_) {}
    return false;
  }

  function writeErr_(sev, scope, type, r, msg, ctxObj) {
    if (dryRun) return;
    if (!shouldWrite_(sev)) return;
    found++; if ((sev||'').toLowerCase()==='error') errors++; else warns++;
    var passRaw = r['Passeport #'] || r['Passeport'] || '';
    var passTxt = normalizePassportToText8_(passRaw);
    errBuf.push([
      passTxt,
      r['Nom'] || '',
      (r['Prénom'] || r['Prenom'] || ''),
      (((r['Prénom']||r['Prenom']||'') + ' ' + (r['Nom']||'')).trim()),
      scope,
      type,
      sev,
      r['Saison'] || '',
      (r['Nom du frais'] || r['Frais'] || r['Produit'] || ''),
      msg || '',
      (typeof jsonCompact_==='function'?jsonCompact_(ctxObj||{}):JSON.stringify(ctxObj||{})),
      new Date()
    ]);
  }

  /* ===== (0) Orphelins d’articles ===== */
  var setInscPS = dict_();
  inscPlay.forEach(function(r){ setInscPS[_psKey_(r)] = true; });
  artPlay.forEach(function(a){
    var k = _psKey_(a);
    if (!setInscPS[k]) writeErr_('warn','ARTICLES','ARTICLE_ORPHELIN', a, 'Article sans inscription correspondante', { key:k });
  });

  /* ===== (1) Éligibilité U vs article ===== */
  artPlay.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (!item) { writeErr_('warn','ARTICLES','ARTICLE_INCONNU', a, 'Article non reconnu (non mappé) – ignoré en export', { libelle: raw }); return; }
    var U = deriveUFromRow_(a);
    var uNum = parseInt(String(U).replace(/^U/i,''),10);
    if (!uNum || isNaN(uNum)) return;
    if ((item.Umin && uNum < item.Umin) || (item.Umax && uNum > item.Umax)) {
      writeErr_('error','ARTICLES','ELIGIBILITE_U', a, 'Article non éligible pour ' + U, { raw: raw, code: item.Code, U: U, Umin:item.Umin, Umax:item.Umax });
    }
  });

  /* ===== (2) Doublons exacts ===== */
  var dupCount = dict_();
  artPlay.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    var k = _articleDupKey_(a, item, raw);
    dupCount[k] = (dupCount[k]||0) + 1;
  });
  artPlay.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    var k = _articleDupKey_(a, item, raw);
    if (dupCount[k] > 1) writeErr_('warn','ARTICLES','DUPLICAT', a, 'Article en double détecté', { code: (item&&item.Code)||'', count: dupCount[k] });
  });

  /* ===== (3) Exclusivité par groupe ===== */
  var mapByPassSeasonGroup = dict_();
  artPlay.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (!item || !item.ExclusiveGroup) return;
    var k = _psKey_(a) + '||' + item.ExclusiveGroup;
    mapByPassSeasonGroup[k] = (mapByPassSeasonGroup[k]||0) + 1;
  });
  artPlay.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (!item || !item.ExclusiveGroup) return;
    var k = _psKey_(a) + '||' + item.ExclusiveGroup;
    if (mapByPassSeasonGroup[k] > 1) {
      writeErr_('error','ARTICLES','EXCLUSIVITE', a, 'Conflit d’articles exclusifs (groupe: ' + item.ExclusiveGroup + ')', { group:item.ExclusiveGroup, count: mapByPassSeasonGroup[k] });
    }
  });

  /* ===== (4) Doublons d'INSCRIPTIONS (même passeport+saison+frais normalisé) ===== */
  function buildInscDupKey_(r){
    var fee = String(r['Nom du frais']||r['Frais']||r['Produit']||'').toLowerCase().replace(/\s+/g,' ').trim();
    return _psKey_(r) + '||' + fee;
  }
  inscPlay.forEach(function(r){
    var k = buildInscDupKey_(r);
    mapInscByKey[k] = (mapInscByKey[k]||0) + 1;
  });
  inscPlay.forEach(function(r){
    var k = buildInscDupKey_(r);
    if (mapInscByKey[k] > 1) writeErr_('warn','INSCRIPTIONS','INSCRIPTION_DUPLICAT', r, 'Inscription en double détectée (même clé)', { key:k, count: mapInscByKey[k] });
  });

  /* ===== (5) U9–U12 sans CDP ===== */
  var hasCdp = dict_();
  artPlay.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (item && String(item.ExclusiveGroup||'') === 'CDP_ENTRAINEMENT') hasCdp[_psKey_(a)] = true;
  });
  inscPlay.forEach(function(r){
    var f = (r['Nom du frais'] || r['Frais'] || r['Produit'] || '');         // <-- bug corrigé: f défini
    if (isIgnoredFeeRetro_(ss, f)) return;                                    // garde-fou
    if (isAdapteMember_(r)) return;
    var uNum = parseInt(String(deriveUFromRow_(r)||'').replace(/^U/i,''),10);
    if (uNum>=9 && uNum<=12 && !hasCdp[_psKey_(r)]) {
      var arts = listActiveOccurrencesForPassport_(ss, SHEETS.ARTICLES, r['Passeport #']);
      writeErr_('warn','INSCRIPTIONS','U9_12_SANS_CDP', r, 'U9–U12 sans CDP', { U: deriveUFromRow_(r), articlesActifs: arts });
    }
  });

  /* ===== (6) U7–U8 sans 2e séance ===== */
  var hasU7U8Second = dict_();
  artPlay.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    var U = deriveUFromRow_(a);
    var uNum = parseInt(String(U||'').replace(/^U/i,''),10);
    if (!uNum || isNaN(uNum) || (uNum !== 7 && uNum !== 8)) return;

    var matchByMapping =
      (item && (String(item.ExclusiveGroup||'') === 'U7U8_2E_SEANCE' || String(item.Code||'') === 'U7U8_2E_SEANCE'));

    function NORM(s){ return String(s||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
    var N = NORM(raw);
    var matchByName = (/(^|\s)2\s*E(\s|$)/.test(N) && /SEANCE/.test(N)) || /DEUXIEME\s+SEANCE/.test(N);

    if (matchByMapping || matchByName) hasU7U8Second[_psKey_(a)] = true;
  });
  inscPlay.forEach(function(r){
    if (isAdapteMember_(r)) return;
    var U = deriveUFromRow_(r);
    var uNum = parseInt(String(U||'').replace(/^U/i,''),10);
    if (uNum === 7 || uNum === 8) {
      if (!hasU7U8Second[_psKey_(r)]) {
        var arts = listActiveOccurrencesForPassport_(ss, SHEETS.ARTICLES, r['Passeport #']);
        writeErr_('warn','INSCRIPTIONS','U7_8_SANS_2E_SEANCE', r, 'U7–U8 sans 2e séance', { U: U, articlesActifs: arts });
      }
    }
  });

  // Écriture batch
  if (!dryRun && errBuf.length) {
    var W = 12;
    var start = shErr.getLastRow() + 1;
    shErr.insertRowsAfter(shErr.getLastRow(), errBuf.length);
    var rows = errBuf.map(function(r){ r = (r||[]).slice(0, W); while (r.length < W) r.push(''); return r; });
    shErr.getRange(start, 1, rows.length, W).setValues(rows);
  }

  appendImportLog_(ss, 'RULES_DONE', JSON.stringify({
    dryRun: dryRun, append: appendMode, filtered: !!filterSet,
    inscScanned: inscAct.length, artScanned: artAct.length, written: errBuf.length,
    inscPlay: inscPlay.length, artPlay: artPlay.length
  }));

  return { found: found, errors: errors, warns: warns, dryRun: dryRun, filtered: !!filterSet, written: errBuf.length };
}
