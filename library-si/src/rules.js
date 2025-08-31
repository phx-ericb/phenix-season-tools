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

// Clé de jointure "Passeport||Saison" – même normalisation des deux côtés
function _psKey_(row) {
  var p8 = normalizePassportPlain8_(row['Passeport #'] || row['Passeport']);
  var s  = String(row['Saison'] || '');
  return p8 + '||' + s;
}

// Clé "Passeport||Saison||base" pour les doublons d'articles
function _articleDupKey_(row, match, rawFrais) {
  var base;
  if (match && match.Code) base = 'CODE:' + String(match.Code);
  else if (match && match.ExclusiveGroup) base = 'EXG:' + String(match.ExclusiveGroup);
  else base = 'LBL:' + String(rawFrais||'').toLowerCase().replace(/\s+/g,' ').trim();
  return _psKey_(row) + '||' + base;
}


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

  // Helpers "dictionnaire sans prototype"
  function dict_(){ return Object.create(null); }

  // Helper : renvoie true si la ligne correspond au programme "Adapté"
function isAdapteMember_(row){
  function norm(s){ s=String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s.toUpperCase(); }
  var cat = norm(row['Catégorie']||row['Categorie']||'');
  var frais = norm(row['Nom du frais']||row['Frais']||row['Produit']||'');
  if (cat.indexOf('ADAPTE') !== -1) return true;
  if (frais.indexOf('ADAPTE') !== -1) return true;

  // Optionnel : si tu veux le baser explicitement sur MAPPINGS (Type=member, AliasContains='Adapté')
  try{
    var mapObj = readSheetAsObjects_(ss.getId(), 'MAPPINGS'); // ss est dispo dans evaluateSeasonRules
    var maps = (mapObj && mapObj.rows) ? mapObj.rows : [];
    var txt = norm((row['Catégorie']||row['Categorie']||'') + ' ' + (row['Nom du frais']||row['Frais']||row['Produit']||''));
    for (var i=0;i<maps.length;i++){
      var mp = maps[i];
      if (String(mp['Type']||'').toLowerCase()!=='member') continue;
      var alias = norm(mp['AliasContains']||mp['Alias']||'');
      if (alias && alias.indexOf('ADAPTE')!==-1){
        if (txt.indexOf(alias)!==-1 || txt.indexOf('ADAPTE')!==-1) return true;
      }
    }
  }catch(e){}
  return false;
}


  // TOUJOURS initialiser avant usage
  var setInscPS = dict_();
  var dupCount = dict_();
  var mapByPassSeasonGroup = dict_();
  var mapInscByKey = dict_();
  var hasCdp = dict_();
  var hasU7U8Second = dict_();

  // Garde-fou si readSheetAsObjects_ renvoie qqch de falsy
  var inscAct = (insc.rows || []).filter(isActive_);
  var artAct  = (art.rows  || []).filter(isActive_);


  // MAPPINGS: ARTICLES
  var catalog = loadArticlesCatalog_(ss); // { items:[{Code, Umin, Umax, AliasContains, ExclusiveGroup}], match(rawName)->item|null }

// ERREURS (création/clear si pas append). En dry-run, on n'écrira rien.
var shErr = getSheetOrCreate_(ss, SHEETS.ERREURS, ['Passeport','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']);
if (!appendMode && !dryRun) {
  shErr.clearContents();
  shErr.getRange(1,1,1,12).setValues([['Passeport','Nom','Prénom','NomComplet','Scope','Type','Severite','Saison','Frais','Message','Contexte','CreatedAt']]);
}

// IMPORTANT : colonne A en texte, pour conserver les zéros
shErr.getRange('A:A').setNumberFormat('@');




  function shouldWrite_(sev) {
    var sevRank = { warn:1, error:2 };
    return (sevRank[(sev||'warn')] || 1) >= (sevRank[threshold] || 1);
  }

  var errBuf = []; // buffer d'écriture en batch

function writeErr_(sev, scope, type, r, msg, ctxObj) {
  if (dryRun) return;
  if (!shouldWrite_(sev)) return;

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
    jsonCompact_(ctxObj||{}),
    new Date()
  ]);
}




  var found=0, warns=0, errors=0;


// ===== (0) ARTICLE ORPHELIN : article actif sans inscription active correspondante =====
// IMPORTANT : la jointure se fait sur Passeport+Saison (pas sur 'Frais')
  inscAct.forEach(function(r){ setInscPS[_psKey_(r)] = true; });
  artAct.forEach(function(a){
    var k = _psKey_(a);
    if (!setInscPS || !setInscPS[k]) { // <- défensif
      writeErr_('warn','ARTICLES','ARTICLE_ORPHELIN', a, 'Article sans inscription correspondante', { key:k });
    }
  });



// ===== (1) Éligibilité âge/U ↔ article =====
artAct.forEach(function(a){
  var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
  var item = catalog.match(raw);
  if (!item) {
    writeErr_('warn','ARTICLES','ARTICLE_INCONNU', a, 'Article non reconnu (non mappé) – ignoré en export', { libelle: raw });
    return;
  }
  var U = deriveUFromRow_(a);
  var uNum = parseInt(String(U).replace(/^U/i,''),10);
  if (!uNum || isNaN(uNum)) return;

  if ((item.Umin && uNum < item.Umin) || (item.Umax && uNum > item.Umax)) {
    var sev = 'error';
    writeErr_(sev, 'ARTICLES', 'ELIGIBILITE_U', a, 'Article non éligible pour ' + U, { raw: raw, code: item.Code, U: U, Umin:item.Umin, Umax:item.Umax });
    found++; (sev==='error'?errors++:warns++);
  }
});



// ===== (2) Doublons exacts sous même base (Code, sinon ExclusiveGroup, sinon Libellé normalisé)
 artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    var k = _articleDupKey_(a, item, raw);
    dupCount[k] = ((dupCount && dupCount[k]) || 0) + 1;
  });
  artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    var k = _articleDupKey_(a, item, raw);
    if (dupCount && dupCount[k] > 1) {
      writeErr_('warn','ARTICLES','DUPLICAT', a, 'Article en double détecté', { code: (item&&item.Code)||'', count: dupCount[k] });
    }
  });


// ===== (3) Exclusivité (un seul article par ExclusiveGroup) =====
artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (!item || !item.ExclusiveGroup) return;
    var k = _psKey_(a) + '||' + item.ExclusiveGroup; // <- normalisé
    mapByPassSeasonGroup[k] = ((mapByPassSeasonGroup && mapByPassSeasonGroup[k]) || 0) + 1;
  });
  artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (!item || !item.ExclusiveGroup) return;
    var k = _psKey_(a) + '||' + item.ExclusiveGroup;
    if (mapByPassSeasonGroup && mapByPassSeasonGroup[k] > 1) {
      writeErr_('error','ARTICLES','EXCLUSIVITE', a, 'Conflit d’articles exclusifs (groupe: ' + item.ExclusiveGroup + ')', { group:item.ExclusiveGroup, count: mapByPassSeasonGroup[k] });
    }
  });

// ===== (4) Doublons d'INSCRIPTIONS (même clé) =====
 var keyCols = getKeyColsFromParams_(ss); // ex: ["Passeport #","Saison"]
  function buildInscDupKey_(r){
    return buildArticleKey_(r) + '||' + String(r['Nom du frais']||'').trim();
  }
  inscAct.forEach(function(r){
    var k = buildInscDupKey_(r);
    mapInscByKey[k] = ((mapInscByKey && mapInscByKey[k]) || 0) + 1;
  });
  inscAct.forEach(function(r){
    var k = buildInscDupKey_(r);
    if (mapInscByKey && mapInscByKey[k] > 1) {
      writeErr_('warn','INSCRIPTIONS','INSCRIPTION_DUPLICAT', r, 'Inscription en double détectée (même clé)', { key:k, count: mapInscByKey[k] });
    }
  });



  // ===== (5) U9–U12 sans CDP (warning) =====
// Marquer les (Passeport||Saison) qui ont un article de groupe exclusif CDP
artAct.forEach(function(a){
    var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
    var item = catalog.match(raw);
    if (item && String(item.ExclusiveGroup||'') === 'CDP_ENTRAINEMENT') {
      hasCdp[_psKey_(a)] = true;
    }
  });
  inscAct.forEach(function(r){
    if (isAdapteMember_(r)) return;
    var uNum = parseInt(String(deriveUFromRow_(r)||'').replace(/^U/i,''),10);
    if (uNum>=9 && uNum<=12 && !(hasCdp && hasCdp[_psKey_(r)])) {
      var arts = listActiveOccurrencesForPassport_(ss, SHEETS.ARTICLES, r['Passeport #']);
      writeErr_('warn','INSCRIPTIONS','U9_12_SANS_CDP', r, 'U9–U12 sans CDP', { U: deriveUFromRow_(r), articlesActifs: arts });
    }
  });

  // ===== (6) U7–U8 sans 2e séance (warning) =====
// Idée: marquer (Passeport||Saison) qui ont un article "2e séance", puis loguer ceux qui n'en ont pas.
// On accepte 2 sources de vérité :
//  - via MAPPINGS (ExclusiveGroup ou Code == 'U7U8_2E_SEANCE')
//  - fallback regex sur le libellé ('2e' + 'séance' ou 'deuxieme séance'), sans accents



artAct.forEach(function(a){
  var raw = (a['Nom du frais'] || a['Frais'] || a['Produit'] || '').toString();
  var item = catalog.match(raw);
  var U = deriveUFromRow_(a);
  var uNum = parseInt(String(U||'').replace(/^U/i,''),10);

  // Ne marque que si l'article est pour U7 ou U8
  if (!uNum || isNaN(uNum) || (uNum !== 7 && uNum !== 8)) return;

  var matchByMapping =
    (item && (String(item.ExclusiveGroup||'') === 'U7U8_2E_SEANCE' || String(item.Code||'') === 'U7U8_2E_SEANCE'));

  // fallback : robustifier la détection par texte brut (sans accents)
  function norm(s){ return String(s||'').toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
  var N = norm(raw);
  var matchByName =
    (/(^|\s)2\s*E(\s|$)/.test(N) && /SEANCE/.test(N)) || /DEUXIEME\s+SEANCE/.test(N);

  if (matchByMapping || matchByName) {
    hasU7U8Second[_psKey_(a)] = true;
  }
});

inscAct.forEach(function(r){
   if (isAdapteMember_(r)) return; 
  var U = deriveUFromRow_(r);
  var uNum = parseInt(String(U||'').replace(/^U/i,''),10);
  if (uNum === 7 || uNum === 8) {
    if (!(hasU7U8Second && hasU7U8Second[_psKey_(r)])) {
      var arts = listActiveOccurrencesForPassport_(ss, SHEETS.ARTICLES, r['Passeport #']);
      writeErr_('warn','INSCRIPTIONS','U7_8_SANS_2E_SEANCE', r, 'U7–U8 sans 2e séance', { U: U, articlesActifs: arts });
    }
  }
});


  // Résumé
  // appendImportLog_(ss, dryRun ? 'RULES_OK_DRYRUN' : 'RULES_OK', JSON.stringify({found:found, errors:errors, warns:warns}));

if (!dryRun && errBuf.length) {
  var W = 12; // largeur d'en-tête ERREURS
  var start = shErr.getLastRow() + 1;
  shErr.insertRowsAfter(shErr.getLastRow(), errBuf.length);
  // pad/troncature défensive
  var rows = errBuf.map(r => {
    r = (r || []).slice(0, W);
    while (r.length < W) r.push('');
    return r;
  });
  shErr.getRange(start, 1, rows.length, W).setValues(rows);
}
  
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

