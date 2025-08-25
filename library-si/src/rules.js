/**
 * v0.7 — rules.gs
 * - Jointure INSCRIPTIONS + ARTICLES (actifs)
 * - Règles:
 *   0) Orphelins: ARTICLE sans INSCRIPTION correspondante (warn)
 *   1) Éligibilité âge/U ↔ article (via MAPPINGS.ARTICLES Umin/Umax)
 *   2) Doublons exacts sous même Code article
 *   3) Exclusivité: un seul article par groupe exclusif
 *   4) Doublons d'INSCRIPTIONS (même clé)
 *   // ===== (5) U9–U12 sans CDP (warning) =====
 * - Log dans ERREURS avec Nom/Prénom/NomComplet/Saison/Frais
 */

function evaluateSeasonRules(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  ensureCoreSheets_(ss);

  var rulesOn = (readParam_(ss, PARAM_KEYS.RULES_ON) || 'TRUE').toUpperCase() === 'TRUE';
  if (!rulesOn) { appendImportLog_(ss, 'RULES_SKIP', 'RULES_ON=FALSE'); return {found:0, errors:0, warns:0}; }

  var dryRun = (readParam_(ss, PARAM_KEYS.RULES_DRY_RUN) || 'FALSE').toUpperCase() === 'TRUE';
  var threshold = (readParam_(ss, PARAM_KEYS.RULES_SEVERITY_THRESHOLD) || 'warn').toLowerCase();
  var appendMode = (readParam_(ss, PARAM_KEYS.RULES_APPEND) || 'FALSE').toUpperCase() === 'TRUE';

  // Lecture des tables
  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var art  = readSheetAsObjects_(ss.getId(), SHEETS.ARTICLES);

  function isActive_(r){
    var can = String(r[CONTROL_COLS.CANCELLED]||'').toLowerCase()==='true';
    var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT]||'').toLowerCase()==='true';
    var st  = (r['Statut de l\'inscription'] || r['Statut'] || '').toString().toLowerCase();
    return !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
  }
  var inscAct = insc.rows.filter(isActive_);
  var artAct  = art.rows.filter(isActive_);

  // MAPPINGS: ARTICLES
  var catalog = loadArticlesCatalog_(ss); // { items:[{Code, Umin, Umax, AliasContains, ExclusiveGroup}], match(rawName)->item|null }

  // ERREURS (création/clear si pas append). En dry-run, on n'écrira rien.
  var shErr = getSheetOrCreate_(ss, SHEETS.ERREURS, ['Passeport','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']);
  if (!appendMode && !dryRun) { shErr.clearContents(); shErr.getRange(1,1,1,12).setValues([['Passeport','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']]); }

  function shouldWrite_(sev) {
    var sevRank = { warn:1, error:2 };
    return (sevRank[(sev||'warn')] || 1) >= (sevRank[threshold] || 1);
  }
  function writeErr_(sev, scope, type, r, msg, ctxObj) {
    if (dryRun) return;            // pas d'écriture en DRY_RUN
    if (!shouldWrite_(sev)) return;
    shErr.appendRow([
      r['Passeport #'] || '',
      r['Nom'] || '',
      (r['Prénom'] || r['Prenom'] || ''),
      (((r['Prénom']||r['Prenom']||'') + ' ' + (r['Nom']||'')).trim()),
      scope,
      type,
      sev,
      r['Saison'] || '',
      (r['Nom du frais'] || r['Frais'] || r['Produit'] || ''),
      msg || '',
      jsonCompact_(ctxObj||{}),
      new Date()
    ]);
  }

  var found=0, warns=0, errors=0;

  // ===== (0) ARTICLE ORPHELIN : article actif sans inscription active correspondante =====
  var keyCols = getKeyColsFromParams_(ss);
  var setInscKeys = {};
  inscAct.forEach(function(r){ setInscKeys[ buildKeyFromRow_(r, keyCols) ] = true; });
  artAct.forEach(function(a){
    var k = buildKeyFromRow_(a, keyCols);
    if (!setInscKeys[k]) {
      found++; warns++;
      writeErr_('warn', 'ARTICLES', 'ARTICLE_ORPHELIN', a, 'Article sans inscription correspondante', { key: k });
    }
  });

  // ===== (1) Éligibilité âge/U ↔ article =====
  artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (!item) return; // pas mappé => pas d'éligibilité à contrôler
    var U = deriveUFromRow_(a);
    var uNum = parseInt(String(U).replace(/^U/i,''),10);
    if (!uNum || isNaN(uNum)) return;

    if ((item.Umin && uNum < item.Umin) || (item.Umax && uNum > item.Umax)) {
      var sev = 'error';
      writeErr_(sev, 'ARTICLES', 'ELIGIBILITE_U', a, 'Article non éligible pour ' + U, { raw: raw, code: item.Code, U: U, Umin:item.Umin, Umax:item.Umax });
      found++; (sev==='error'?errors++:warns++);
    }
  });

  // ===== (2) Doublons exacts sous même code article/passeport/saison =====
  var mapByPassSeasonCode = {};
  artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    var code = item ? item.Code : raw.toLowerCase(); // à défaut, raw
    var k = (a['Passeport #']||'') + '||' + (a['Saison']||'') + '||' + code;
    mapByPassSeasonCode[k] = (mapByPassSeasonCode[k] || 0) + 1;
  });
  artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    var code = item ? item.Code : raw.toLowerCase();
    var k = (a['Passeport #']||'') + '||' + (a['Saison']||'') + '||' + code;
    if (mapByPassSeasonCode[k] > 1) {
      writeErr_('warn', 'ARTICLES', 'DUPLICAT', a, 'Article en double détecté', { code: code, count: mapByPassSeasonCode[k] });
      found++; warns++;
    }
  });

  // ===== (3) Exclusivité (un seul article par ExclusiveGroup) =====
  var mapByPassSeasonGroup = {};
  artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (!item || !item.ExclusiveGroup) return;
    var k = (a['Passeport #']||'') + '||' + (a['Saison']||'') + '||' + item.ExclusiveGroup;
    mapByPassSeasonGroup[k] = (mapByPassSeasonGroup[k] || 0) + 1;
  });
  artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (!item || !item.ExclusiveGroup) return;
    var k = (a['Passeport #']||'') + '||' + (a['Saison']||'') + '||' + item.ExclusiveGroup;
    if (mapByPassSeasonGroup[k] > 1) {
      writeErr_('error', 'ARTICLES', 'EXCLUSIVITE', a, 'Conflit d’articles exclusifs (groupe: ' + item.ExclusiveGroup + ')', { group: item.ExclusiveGroup, count: mapByPassSeasonGroup[k] });
      found++; errors++;
    }
  });

  // ===== (4) Doublons d'INSCRIPTIONS (même clé) =====
  var mapInscByKey = {};
  inscAct.forEach(function(r){
    var k = buildKeyFromRow_(r, keyCols);
    mapInscByKey[k] = (mapInscByKey[k] || 0) + 1;
  });
  inscAct.forEach(function(r){
    var k = buildKeyFromRow_(r, keyCols);
    if (mapInscByKey[k] > 1) {
      writeErr_('warn', 'INSCRIPTIONS', 'INSCRIPTION_DUPLICAT', r, 'Inscription en double détectée (même clé)', { key: k, count: mapInscByKey[k] });
      found++; warns++;
    }
  });

  // ===== (5) U9–U12 sans CDP (warning) =====
// Marquer les (passeport||saison) qui ont un article de groupe exclusif CDP

  var hasCdp = {};
  artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (item && String(item.ExclusiveGroup||'') === 'CDP_ENTRAINEMENT') {
      var k = (a['Passeport #']||'') + '||' + (a['Saison']||'');
      hasCdp[k] = true;
    }
  });

  inscAct.forEach(function(r){
    var U = deriveUFromRow_(r);
    var uNum = parseInt(String(U||'').replace(/^U/i,''),10);
    if (!uNum || uNum < 9 || uNum > 12) return;
    var k = (r['Passeport #']||'') + '||' + (r['Saison']||'');
    if (hasCdp[k]) return;

    // Contexte : liste des articles actifs du passeport
    var arts = listActiveOccurrencesForPassport_(ss, SHEETS.ARTICLES, r['Passeport #']);
    writeErr_('warn', 'INSCRIPTIONS', 'U9_12_SANS_CDP', r, 'U9–U12 sans CDP', { U: U, articlesActifs: arts });
    found++; warns++;
  });

  // Résumé
  appendImportLog_(ss, dryRun ? 'RULES_OK_DRYRUN' : 'RULES_OK', JSON.stringify({found:found, errors:errors, warns:warns}));
  return {found:found, errors:errors, warns:warns, dryRun: dryRun};
}

/** Charge le catalogue d’articles depuis MAPPINGS (section "ARTICLES") */
/** Charge le catalogue d’articles :
 *  - PRIORITAIRE : feuille MAPPINGS "entête unifiée" (Type=article)
 *  - FALLBACK   : ancienne section "ARTICLES" dans MAPPINGS (header sur la ligne suivante)
 */
function loadArticlesCatalog_(ss) {
  var sh = ss.getSheetByName(SHEETS.MAPPINGS);
  var items = [];
  if (sh && sh.getLastRow() >= 2) {
    var data = sh.getDataRange().getValues();
    var H = (data[0]||[]).map(function(h){ return String(h||'').trim(); });
    function idx(k){ var i=H.indexOf(k); return i<0 ? null : i; }

    // --- Chemin "entête unifiée"
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

    // --- Fallback : ancienne section "ARTICLES"
    for (var i=0;i<data.length;i++){
      if (String(data[i][0]).toUpperCase().trim() === 'ARTICLES') {
        var header = (i+1 < data.length) ? data[i+1] : null;
        if (!header) break;
        var hIdx = {}; header.forEach(function(h,j){ hIdx[String(h).trim()] = j; });
        for (var r=i+2; r<data.length; r++){
          var row = data[r]||[];
          var code = (row[hIdx['Code']]||'').toString().trim();
          var alias = (row[hIdx['AliasContains']]||'').toString().trim();
          var umin = parseInt(row[hIdx['Umin']]||'',10); if (isNaN(umin)) umin = null;
          var umax = parseInt(row[hIdx['Umax']]||'',10); if (isNaN(umax)) umax = null;
          var excl = (row[hIdx['ExclusiveGroup']]||'').toString().trim();
          if (!code && !alias) break;
          items.push({ Code: code, AliasContains: alias, Umin: umin, Umax: umax, ExclusiveGroup: excl });
        }
        break;
      }
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

