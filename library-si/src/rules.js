/** rules.js — v0.8.2 (lib, blindé Spreadsheet)
 * - Exclusion robuste des frais "Entraîneurs" des règles joueurs.
 * - Wrapper d'ignore (utilise SR_isIgnoredFeeRetro_ si dispo, sinon fallback local).
 * - Zéro dépendance à buildArticleKey_, clés recalculées localement.
 * - Fallbacks prudents pour exécuter en bibliothèque.
 * - NEW: coercion sûre en Spreadsheet via ensureSpreadsheet_ (si absent).
 */

/* ====== Shim: obtenir un vrai Spreadsheet peu importe l’input ====== */
if (typeof extractSpreadsheetId_ !== 'function') {
  function extractSpreadsheetId_(s) {
    var m = String(s||'').match(/[-\w]{25,}/);
    if (m && m[0]) return m[0];
    throw new Error('extractSpreadsheetId_: ID/URL invalide: ' + s);
  }
}
if (typeof _debugType_ !== 'function') {
  function _debugType_(x) {
    if (x == null) return 'null/undefined';
    if (typeof x === 'string') return 'string';
    if (typeof x.getSheetByName === 'function') return 'Spreadsheet';
    if (typeof x.getId === 'function') return 'DriveFile(id=' + x.getId() + ')';
    return Object.prototype.toString.call(x);
  }
}
if (typeof ensureSpreadsheet_ !== 'function') {
  function ensureSpreadsheet_(ssOrId) {
    // déjà Spreadsheet ?
    if (ssOrId && typeof ssOrId.getSheetByName === 'function') return ssOrId;
    // DriveFile ?
    if (ssOrId && typeof ssOrId.getId === 'function') return SpreadsheetApp.openById(ssOrId.getId());
    // string (id|url) ?
    if (typeof ssOrId === 'string' && ssOrId.trim()) return SpreadsheetApp.openById(extractSpreadsheetId_(ssOrId));
    // saison courante si helpers dispos
    if (typeof getSeasonSpreadsheet_ === 'function' && typeof getSeasonId_ === 'function') {
      return getSeasonSpreadsheet_(getSeasonId_());
    }
    if (typeof getSeasonId_ === 'function') return SpreadsheetApp.openById(getSeasonId_());
    var active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
    throw new Error('ensureSpreadsheet_: impossible d’obtenir un Spreadsheet valide (input=' + _debugType_(ssOrId) + ')');
  }
}

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
      var iAllow = idx('AllowOrphan'); // 👈 nouvelle colonne facultative
      for (var r=1; r<data.length; r++){
        var row = data[r]||[];
        if (String(row[iType]||'').trim().toLowerCase() !== 'article') continue;
        var alias = String(row[iAli]||'').trim(); if (!alias) continue;
        var umin = (iUmin==null) ? null : parseInt(row[iUmin]||'',10); if (isNaN(umin)) umin=null;
        var umax = (iUmax==null) ? null : parseInt(row[iUmax]||'',10); if (isNaN(umax)) umax=null;
        var code = (iCode==null) ? '' : String(row[iCode]||'').trim();
        var excl = (iExGrp==null) ? '' : String(row[iExGrp]||'').trim();
        var allow = (iAllow==null) ? false : String(row[iAllow]||'').toUpperCase()==='TRUE';
        items.push({ Code: code, AliasContains: alias, Umin: umin, Umax: umax, ExclusiveGroup: excl, AllowOrphan: allow });
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
    var iAllow2 = idx('AllowOrphan'); // 👈 si dispo, on lit aussi
    for (var r2=1; r2<data.length; r2++){
      var row2 = data[r2]||[];
      var alias2 = String(row2[iAli2]||'').trim(); if (!alias2) continue;
      var umin2 = (iUmin2==null) ? null : parseInt(row2[iUmin2]||'',10); if (isNaN(umin2)) umin2=null;
      var umax2 = (iUmax2==null) ? null : parseInt(row2[iUmax2]||'',10); if (isNaN(umax2)) umax2=null;
      var code2 = (iCode2==null) ? '' : String(row2[iCode2]||'').trim();
      var excl2 = (iExGrp2==null) ? '' : String(row2[iExGrp2]||'').trim();
      var allow2 = (iAllow2==null) ? false : String(row2[iAllow2]||'').toUpperCase()==='TRUE';
      items.push({ Code: code2, AliasContains: alias2, Umin: umin2, Umax: umax2, ExclusiveGroup: excl2, AllowOrphan: allow2 });
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
  // --- CHANGEMENT CRUCIAL ICI: coercion sûre ---
  var ss = ensureSpreadsheet_(seasonSheetId);   // <— au lieu de getSeasonSpreadsheet_(...)

  if (typeof ensureCoreSheets_ === 'function') ensureCoreSheets_(ss);

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
    ['Passeport #','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']
  );

  if (!appendMode) {
    if (!filterSet) {
      shErr.clearContents();
      shErr.getRange(1,1,1,12).setValues([['Passeport #','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']]);
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

  // Passeport en texte, mais sans casser si la colonne a des cellules “spéciales”
  try {
    var last = shErr.getLastRow();
    if (last <= 1) {
      shErr.getRange(1, 1).setNumberFormat('@');      // juste l’entête si vide
    } else {
      shErr.getRange(2, 1, last - 1, 1).setNumberFormat('@'); // lignes réelles
    }
  } catch (e) {
    appendImportLog_(ss, 'RULES_WARN_NUMFORMAT', String(e));
  }

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
function _splitKeys_(csv){
  return String(csv||'')
    .split(',')
    .map(function(s){
      return String(s||'').trim()
        .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
        .toLowerCase();
    })
    .filter(Boolean);
}
function _hay_(s){
  return String(s||'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase();
}
function _matchAny_(hay, keys){
  for (var i=0;i<keys.length;i++) if (hay.indexOf(keys[i]) !== -1) return true;
  return false;
}

var setInscPS = dict_();
inscPlay.forEach(function(r){ setInscPS[_psKey_(r)] = true; });

// NEW: construit la config d’allow-orphan (PARAMS + RULES_JSON)
var allowCfg = _buildAllowOrphanConfig_(ss);

// NEW: mots-clés de camp (Nom du frais)
var campKeys = _splitKeys_((typeof readParam_==='function' ? readParam_(ss, 'RETRO_CAMP_KEYWORDS') : '') || '');

artPlay.forEach(function(a){
  // clé passeport+saison
  var k = _psKey_(a);

  // Si une inscription correspondante existe → pas orphelin
  if (setInscPS[k]) return;

  // Libellé + mapping
  var raw  = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
  var item = catalog.match(raw); // peut être null si non mappé

  // --- (A) Cas spécial: CAMP orphelin => code U13U18_CAMP_SEUL --- LEGACY --- MAINTENANT DANS RULES_FAST
  //     (simple: basé UNIQUEMENT sur RETRO_CAMP_KEYWORDS et l'âge U13–U18)
  // if (campKeys.length) {
  //   var hay = _hay_(raw);
  //   if (_matchAny_(hay, campKeys)) {
  //     var U = deriveUFromRow_(a);
  //     var uNum = parseInt(String(U||'').replace(/^U/i,''),10);
  //     if (uNum && !isNaN(uNum) && uNum >= 13 && uNum <= 18) {
  //       var arts = listActiveOccurrencesForPassport_(ss, SHEETS.ARTICLES, a['Passeport #']);
  //       writeErr_(
  //         'warn','INSCRIPTIONS','U13U18_CAMP_SEUL', a,
  //         'Inscription au camp de sélection sans inscription à la saison',
  //         { U: U, camp:true, saison:false, key:k, raw:raw, code:(item&&item.Code)||'', articlesActifs: arts }
  //       );
  //       return; // on a émis le cas camp-seul; ne pas retomber sur ARTICLE_ORPHELIN
  //     }
  //   }
  // }

  // --- (B) Tolérance: AllowOrphan (MAPPINGS / PARAMS / RULES_JSON)
  if ((item && item.AllowOrphan === true) || isAllowedOrphan_(ss, a, item, raw, allowCfg)) return;

  // --- (C) Orphelin générique

  // Si c’est un CAMP (détecté par campKeys), on NE le marque pas orphelin ici:
// la détection est gérée par la version fast avec U13U18_CAMP_SEUL.
if (campKeys.length) {
  var hay2 = _hay_(raw);
  if (_matchAny_(hay2, campKeys)) return;
}

  writeErr_(
    'warn','ARTICLES','ARTICLE_ORPHELIN', a,
    'Article sans inscription correspondante',
    { key:k, code:(item&&item.Code)||'', group:(item&&item.ExclusiveGroup)||'', raw:raw }
  );
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
  var raw  = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
  var item = catalog.match(raw);

  // 1) Mapping explicite (comme avant)
  var byExGroup = (item && String(item.ExclusiveGroup||'') === 'CDP_ENTRAINEMENT');
  var byCode    = (item && String(item.Code||'') === 'CDP_ENTRAINEMENT');

  // 2) Fallback lexical robuste (si mapping pas calé)
  //    On considère CDP présent si libellé contient "cdp" ET (entrain|séance|training)
  var n = RL_norm_(raw); // minuscule, sans accents
  var byLex = (n.indexOf('cdp') >= 0) && (
               n.indexOf('entrain') >= 0 || n.indexOf('seance') >= 0 || n.indexOf('training') >= 0
             );

  if (byExGroup || byCode || byLex) hasCdp[_psKey_(a)] = true;
});

inscPlay.forEach(function(r){
  if (isAdapteMember_(r)) return;
  var U = deriveUFromRow_(r);
  var uNum = parseInt(String(U||'').replace(/^U/i,''),10);
  if (uNum>=9 && uNum<=12 && !hasCdp[_psKey_(r)]) {
    var arts = listActiveOccurrencesForPassport_(ss, SHEETS.ARTICLES, r['Passeport #']);
    writeErr_('warn','INSCRIPTIONS','U9_12_SANS_CDP', r, 'U9–U12 sans CDP', { U: U, articlesActifs: arts });
  }
});

  /* ===== (6) U7–U8 sans 2e séance ===== */
var hasU7U8Second = dict_();
artPlay.forEach(function(a){
  var raw  = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
  var item = catalog.match(raw);
  var U = deriveUFromRow_(a);
  var uNum = parseInt(String(U||'').replace(/^U/i,''),10);
  if (!uNum || isNaN(uNum) || (uNum !== 7 && uNum !== 8)) return;

  var matchByMapping =
    (item && (String(item.ExclusiveGroup||'') === 'U7U8_2E_SEANCE' || String(item.Code||'') === 'U7U8_2E_SEANCE'));

  function NORM(s){ return String(s||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
  var N = NORM(raw);

  // 🔁 plus tolérant: (2e|2ème|deuxième) + (SEANCE|MATCH)  OU  mention "SAMEDI"
  var matchByName =
      (/(^|\b)(2\s*E|2[EÈ]ME|DEUXIEME|DEUXIÈME)\b/.test(N) && /(SEANCE|MATCH)/.test(N))
      || /SAMEDI/.test(N);

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

/* global _rulesBuildErrorsFast_, _rulesWriteFull_, _rulesUpsertForPassports_, _toPassportSet_, getSeasonId_, getSeasonSpreadsheet_, ensureSpreadsheet_, SHEETS */

/** INCR (FAST) – construit les erreurs pour une liste de passeports */
function _rulesBuildErrorsIncrFast_(passports, ss) {
  ss = ss || (typeof getSeasonSpreadsheet_ === 'function'
      ? getSeasonSpreadsheet_(getSeasonId_())
      : ensureSpreadsheet_(SpreadsheetApp.getActiveSpreadsheet()));
  var set = (typeof _toPassportSet_ === 'function')
      ? _toPassportSet_(passports)
      : new Set((passports || []).map(function (x) { return String(x || '').trim(); }).filter(Boolean));
  return _rulesBuildErrorsFast_(ss, set);
}

/** FULL (FAST) – build + write ERREURS (remplace le legacy côté lib) */
function runEvaluateRulesFast_(ss) {
  ss = ss || (typeof getSeasonSpreadsheet_ === 'function'
      ? getSeasonSpreadsheet_(getSeasonId_())
      : ensureSpreadsheet_(SpreadsheetApp.getActiveSpreadsheet()));
  var res = _rulesBuildErrorsFast_(ss, /*touchedSet*/ null);
  _rulesWriteFull_(ss, res.errors, res.header);
  return { written: res.errors.length, ledger: res.ledgerCount, joueurs: res.joueursCount };
}




/** Écrit ERREURS (FULL) en une passe, header fiable */
function _rulesWriteFull_(ss, rows, header) {
  ss = ensureSpreadsheet_(ss);
  var name = (typeof SHEETS !== 'undefined' && SHEETS.ERREURS) ? SHEETS.ERREURS : 'ERREURS';
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);

  var H = Array.isArray(header) && header.length ? header.slice() : [
    'Passeport #','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt'
  ];

  sh.clearContents();
  sh.getRange(1, 1, 1, H.length).setValues([H]);

  if (rows && rows.length) {
    var data = rows.map(function (r) {
      var a = (r || []).slice(0, H.length);
      while (a.length < H.length) a.push('');
      return a;
    });
    sh.getRange(2, 1, data.length, H.length).setValues(data);
  }

  // Passeport en texte
  try {
    var h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
    var c = h.indexOf('Passeport #'); if (c < 0) c = h.indexOf('Passeport');
    if (c >= 0 && sh.getLastRow() >= 2) {
      sh.getRange(2, c + 1, sh.getLastRow() - 1, 1).setNumberFormat('@');
    }
  } catch (_e) { /* no-op */ }
}

/** UPSERT ciblé par passeports (INCR) dans ERREURS */
function _rulesUpsertForPassports_(ss, newRows, touchedSet, header) {
  ss = ensureSpreadsheet_(ss);
  var name = (typeof SHEETS !== 'undefined' && SHEETS.ERREURS) ? SHEETS.ERREURS : 'ERREURS';
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);

  var H = Array.isArray(header) && header.length ? header.slice() : [
    'Passeport #','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, H.length).setValues([H]);
  } else {
    // s’assurer que l’entête correspond (on n’étire pas la structure ici)
    var curH = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
    if (curH.join('¦') !== H.join('¦')) {
      // réécrire proprement
      var V = (sh.getLastRow() > 1) ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [];
      sh.clearContents();
      sh.getRange(1, 1, 1, H.length).setValues([H]);
      if (V.length) {
        // best-effort pour recoller (on tronque/pad)
        var fixed = V.map(function (r) {
          var a = (r || []).slice(0, H.length);
          while (a.length < H.length) a.push('');
          return a;
        });
        sh.getRange(2, 1, fixed.length, H.length).setValues(fixed);
      }
    }
  }

  // Construire set des passeports touchés
  var passSet = (touchedSet && typeof touchedSet.forEach === 'function')
    ? touchedSet
    : new Set((newRows || []).map(function (r) {
        return String((r && r[0]) || '').replace(/\D/g, '').slice(-8).padStart(8, '0');
      }).filter(Boolean));

  // Trouver la colonne Passeport
  var Hnow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];
  var cP = Hnow.indexOf('Passeport #'); if (cP < 0) cP = Hnow.indexOf('Passeport');

  if (cP >= 0 && sh.getLastRow() >= 2 && passSet.size) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    var toDel = [];
    for (var i = 0; i < data.length; i++) {
      var p = String(data[i][cP] || '').replace(/\D/g, '').slice(-8).padStart(8, '0');
      if (p && passSet.has(p)) toDel.push(2 + i);
    }
    toDel.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  }

  if (newRows && newRows.length) {
    var rows = newRows.map(function (r) {
      var a = (r || []).slice(0, H.length);
      while (a.length < H.length) a.push('');
      return a;
    });
    var start = sh.getLastRow() + 1;
    sh.insertRowsAfter(sh.getLastRow(), rows.length);
    sh.getRange(start, 1, rows.length, H.length).setValues(rows);
  }

  // Passeport en texte
  try {
    var c = Hnow.indexOf('Passeport #'); if (c < 0) c = Hnow.indexOf('Passeport');
    if (c >= 0 && sh.getLastRow() >= 2) {
      sh.getRange(2, c + 1, sh.getLastRow() - 1, 1).setNumberFormat('@');
    }
  } catch (_e2) { /* no-op */ }
}
