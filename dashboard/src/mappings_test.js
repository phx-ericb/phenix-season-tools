/** Test/validation des MAPPINGS unifiés.
 * - Détecte "no match", duplications, exclusions et conflits d'exclusivité.
 * - Champs requis (GroupeFmt/CategorieFmt) pour Type=member.
 * - Calcule U = SEASON_YEAR - année_naissance (fallback: "Uxx" dans libellé).
 * - Ignore les adultes (U > RETRO_MEMBER_MAX_U, défaut 18) dans le TEST uniquement.
 */

// --- helper genre robuste (couvre "Identité de genre", "Genre", etc.)
function _mt_extractGenreSmart_(row){
  var keys = [
    'Identité de genre','Identité de Genre','Identite de genre','Identite de Genre',
    'Genre','Sexe','Sex','Gender','F/M','MF','Gendre','Type',
    'Categorie','Catégorie','Catégories'
  ];
  var raw = '';
  for (var i=0;i<keys.length;i++){
    if (row && row.hasOwnProperty(keys[i]) && String(row[keys[i]]||'').trim()!==''){
      raw = String(row[keys[i]]); break;
    }
  }
  if (!raw) return { label:'', initiale:'' };
  function _nrmLow(s){ try{ s=String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,''); }catch(e){} return s.toLowerCase().trim(); }
  var n = _nrmLow(raw);
  if (/^(m|masculin|male|man|garcon|gar\u00e7on|homme|boy)\b/.test(n) || /\bu ?\d+\s*m\b/.test(n) || /\bmasc\b/.test(n))
    return { label:'Masculin', initiale:'M' };
  if (/^(f|feminin|female|woman|fille|dame|girl)\b/.test(n) || /\bu ?\d+\s*f\b/.test(n) || /\bfem\b/.test(n))
    return { label:'Féminin', initiale:'F' };
  if (/^(mixte|mix|x|non binaire|non-binaire|nb|autre)\b/.test(n))
    return { label:'Mixte', initiale:'X' };
  return { label:String(raw), initiale:String(raw).charAt(0).toUpperCase() };
}

function testMappings(seasonSpreadsheetId, opts) {
  var _seasonYearCache = null;
  var _maxUCache = null;

  opts = opts || {};
  var maxPreviewPerBucket = opts.maxPreviewPerBucket || 10;

  var ss = SpreadsheetApp.openById(seasonSpreadsheetId);

  // Feuille mappings à nom fixe
  var shMap = ss.getSheetByName('MAPPINGS');
  if (!shMap) return { ok:false, error:'Feuille MAPPINGS introuvable' };

  var map = _readUnifiedMappings_(shMap);
  if (map.errors.length) return { ok:false, error:'Entête MAPPINGS invalide: '+map.errors.join(', ') };

  var insc = _readActive_(ss, 'INSCRIPTIONS');
  var arts = _readActive_(ss, 'ARTICLES');

  // Index utilitaire sur articles
  var byArticleName = {};
  arts.forEach(function(a){
    var raw = (_articleNameFromRow_(a) || '').toString().trim();
    if (!raw) return;
    byArticleName[raw] = (byArticleName[raw] || 0) + 1;
  });
  var uniqueArticles = Object.keys(byArticleName).sort();

  var out = {
    ok:true,
    stats: { uniqueArticles: uniqueArticles.length, inscriptions: insc.length, articles: arts.length },
    articleBuckets: { matched:[], noMatch:[], multiMatch:[] }, // multiMatch gardé à titre de visu si tu veux
    memberBuckets:  { matched:[], noMatch:[], multiMatch:[], missingFmt:[] },
    warnings: [],
    errors: [] // on y met aussi les conflits d'exclusivité
  };

  out.debug = { seasonYear: null, maxU: null, adultsSkipped: 0 };

  // ---------- ARTICLES ----------
  uniqueArticles.forEach(function(raw){
    var matches = _matchArticle_(map.rows, raw);

    if (matches.length === 0) {
      out.articleBuckets.noMatch.push({ raw: raw, count: byArticleName[raw] });
      return;
    }

    // Si une exclusion matche, on ne traite que celle-là (pas de duplication)
    if (matches.length === 1 && matches[0].Exclude) {
      out.articleBuckets.matched.push({ raw: raw, match: _minify(matches[0]) });
      return;
    }

    // Conflit d'exclusivité ?
    var conflict = _exclusiveConflict(matches);
    if (conflict) {
      out.errors.push({
        type: 'exclusive_conflict_article',
        raw: raw,
        group: conflict.group,
        matches: conflict.matches.map(_minify)
      });
      // on ne génère pas de duplications en cas de conflit d'exclusivité
      return;
    }

    // Plusieurs matches → duplication logique (on pousse tous les rendus)
    if (matches.length > 1) {
      out.articleBuckets.multiMatch.push({ raw: raw, matches: matches.map(_minify) });
    }
    matches.forEach(function(m){
      out.articleBuckets.matched.push({ raw: raw, match: _minify(m) });
    });
  });

  // ---------- MEMBERS ----------
  var maxU = _maxUCache != null ? _maxUCache : (_maxUCache = _getMaxU_(ss));
  insc.forEach(function(r){
    var U = _deriveUFromRow_(r, ss);
    var uNum = U ? parseInt(String(U).replace(/^U/i,''),10) : null;
    var articleName = _articleNameFromRow_(r);
    var ginfo = _mt_extractGenreSmart_(r);
    var genreInit = ginfo.initiale; // "M" | "F" | "X" | ""

    var seasonYear = _seasonYearCache || (_seasonYearCache = _getSeasonYear_(ss));
    out.debug.seasonYear = seasonYear;
    out.debug.maxU = (maxU == null ? null : maxU);

    // Ignore les adultes (en test) si borne param
    if (uNum != null && maxU != null && uNum > maxU) {
      out.debug.adultsSkipped = (out.debug.adultsSkipped || 0) + 1;
      return; // on n’ajoute rien dans noMatch
    }

    var matches = _matchMember_(map.rows, { U: uNum, genre: genreInit, articleName: articleName });

    if (matches.length === 0) {
      if (out.memberBuckets.noMatch.length < maxPreviewPerBucket) {
        out.memberBuckets.noMatch.push({ passeport: r['Passeport #'], U: U, genre: genreInit, article: articleName });
      }
      return;
    }

    if (matches.length === 1 && matches[0].Exclude) {
      // Exclusion prioritaire: on ne produit rien pour cette ligne
      return;
    }

    // Conflit d'exclusivité ?
    var conflict = _exclusiveConflict(matches);
    if (conflict) {
      out.errors.push({
        type: 'exclusive_conflict_member',
        passeport: r['Passeport #'],
        U: U,
        genre: genreInit,
        article: articleName,
        group: conflict.group,
        matches: conflict.matches.map(_minify)
      });
      return;
    }

    // Duplication: preview pour chaque match
    if (matches.length > 1) {
      if (out.memberBuckets.multiMatch.length < maxPreviewPerBucket) {
        out.memberBuckets.multiMatch.push({ passeport: r['Passeport #'], U: U, genre: genreInit, matches: matches.map(_minify) });
      }
    }

    for (var i=0; i<matches.length; i++){
      var m = matches[i];

      // Champs requis pour Type=member
      if (m.Type === 'member') {
        var miss = [];
        if (!m.GroupeFmt)    miss.push('GroupeFmt');
        if (!m.CategorieFmt) miss.push('CategorieFmt');
        if (miss.length) {
          if (out.memberBuckets.missingFmt.length < maxPreviewPerBucket) {
            out.memberBuckets.missingFmt.push({ U: U, genre: genreInit, mapping: _minify(m), missing: miss });
          }
          continue;
        }
      }

      var rendered = {
        groupe: _renderFmt_(m.GroupeFmt, r, U, genreInit),
        categorie: _renderFmt_(m.CategorieFmt, r, U, genreInit)
      };

      if (out.memberBuckets.matched.length < maxPreviewPerBucket) {
        out.memberBuckets.matched.push({ U: U, genre: genreInit, mapping: _minify(m), rendered: rendered });
      }
    }
  });

  // Sanity CDP (si tu veux conserver cet avertissement)
  var hasCDP = map.rows.some(function(x){ return x.Type==='article' && /^(CDP|CDP_ENTRAINEMENT)$/i.test(String(x.ExclusiveGroup||'')); });
  if (!hasCDP) out.warnings.push('Aucun mapping Type=article avec ExclusiveGroup=CDP_ENTRAINEMENT/CDP — la règle U9–U12 sans CDP ne pourra jamais se déclencher.');

  return out;

  // ===== Helpers =====

  function _exclusiveConflict(hits){
    // Conflit s’il y a au moins 2 matches partageant le même group non‑vide
    var by = {};
    hits.forEach(function(h){
      var g = (h.ExclusiveGroup||'').trim();
      if (!g) return;
      by[g] = by[g] || [];
      by[g].push(h);
    });
    var gKeys = Object.keys(by);
    for (var i=0;i<gKeys.length;i++){
      var k = gKeys[i];
      if (by[k].length > 1) return { group: k, matches: by[k] };
    }
    return null;
  }

  function _getSeasonYear_(ss){
    try {
      var sh = ss.getSheetByName('PARAMS'); if (sh) {
        var v = sh.getDataRange().getValues();
        var head = (v[0]||[]).map(String);
        var ik = head.findIndex(function(h){ return /^(clé|cle|key)$/i.test(h); });
        var iv = head.findIndex(function(h){ return /^(valeur|value)$/i.test(h); });
        if (ik>-1 && iv>-1) {
          var byKey={};
          for (var r=1;r<v.length;r++){ var k=v[r][ik]; if(!k) continue; byKey[String(k)] = v[r][iv]; }
          var sy = parseInt(byKey['SEASON_YEAR'], 10);
          if (!isNaN(sy) && sy>1900 && sy<3000) return sy;
          var sl = String(byKey['SEASON_LABEL']||'');
          var m = sl.match(/(20\d{2})/); if (m) return parseInt(m[1],10);
        }
      }
      var title = ss.getName();
      var mt = title && title.match(/(20\d{2})/); if (mt) return parseInt(mt[1],10);
    } catch(e){}
    return (new Date()).getFullYear();
  }

  function _getMaxU_(ss){
    try {
      var sh = ss.getSheetByName('PARAMS'); if (!sh) return 18;
      var v = sh.getDataRange().getValues(); if (!v.length) return 18;
      var head = (v[0]||[]).map(String);
      var ik = head.findIndex(function(h){ return /^(clé|cle|key)$/i.test(h); });
      var iv = head.findIndex(function(h){ return /^(valeur|value)$/i.test(h); });
      if (ik>-1 && iv>-1) {
        for (var r=1; r<v.length; r++){
          if (String(v[r][ik]) === 'RETRO_MEMBER_MAX_U') {
            var n = parseInt(v[r][iv], 10);
            if (!isNaN(n) && n > 0) return n;
          }
        }
      }
    } catch(e){}
    return 18;
  }

  function _extractBirthYear_(row){
    var cand = [
      row['Année de naissance'], row['Annee de naissance'], row['Annee'],
      row['Birth Year'], row['Naissance'], row['Date de naissance'], row['Date naissance']
    ];
    for (var i=0;i<cand.length;i++){
      var v = cand[i];
      if (v instanceof Date) return v.getFullYear();
      var s = String(v||'').trim();
      if (!s) continue;
      var m = s.match(/(\d{4})/);
      if (m) { var y = parseInt(m[1],10); if (y>1900 && y<3000) return y; }
    }
    return null;
  }

  function _articleNameFromRow_(r){
    return r['Nom du produit'] || r['Nom du frais'] || r['Frais'] || r['Produit'] || '';
  }

  function _readUnifiedMappings_(sh){
    var data = sh.getDataRange().getValues();
    var H = (data[0]||[]).map(String);
    var need = ['Type','AliasContains','Umin','Umax','Genre','GroupeFmt','CategorieFmt','Exclude','Priority','Code','ExclusiveGroup'];
    var miss = need.filter(function(k){ return H.indexOf(k)===-1; });
    if (miss.length) return { rows:[], errors: miss };

    function idx(k){ return H.indexOf(k); }

    var rows = [];
    for (var r=1; r<data.length; r++){
      var row = data[r]||[];
      var type = String(row[idx('Type')]||'').trim().toLowerCase();
      if (!type) continue;
      var rec = {
        Type: (type==='article'?'article':(type==='member'?'member':type)),
        AliasContains: String(row[idx('AliasContains')]||'').trim(),
        Umin: _toInt(row[idx('Umin')]),
        Umax: _toInt(row[idx('Umax')]),
        Genre: String(row[idx('Genre')]||'*').trim().toUpperCase(),
        GroupeFmt: String(row[idx('GroupeFmt')]||'').trim(),
        CategorieFmt: String(row[idx('CategorieFmt')]||'').trim(),
        Exclude: _toBool(row[idx('Exclude')]),
        Priority: _toInt(row[idx('Priority')]) || 0,
        Code: String(row[idx('Code')]||'').trim(),
        ExclusiveGroup: String(row[idx('ExclusiveGroup')]||'').trim()
      };
      rows.push(rec);
    }
    // tri global : Priority desc puis AliasContains plus long d'abord
    rows.sort(function(a,b){
      if (a.Priority!==b.Priority) return b.Priority - a.Priority;
      return (b.AliasContains||'').length - (a.AliasContains||'').length;
    });
    return { rows: rows, errors: [] };
  }

  function _toInt(v){ var n=parseInt(v,10); return isNaN(n)?null:n; }
  function _toBool(v){ var s=String(v||'').toLowerCase(); return (s==='true'||s==='1'||s==='oui'||s==='yes'); }

  function _norm(s){
    // supprime accents + lower
    return String(s||'')
      .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
      .toLowerCase();
  }

  function _matchArticle_(rows, rawName){
    var raw = _norm(rawName);
    var hits = rows.filter(function(x){
      if (x.Type!=='article') return false;
      // on garde les Exclude pour prioriser ensuite
      if (x.AliasContains && raw.indexOf(_norm(x.AliasContains))===-1) return false;
      return true;
    });
    // Exclusion prioritaire
    var excl = hits.filter(function(h){ return h.Exclude; });
    if (excl.length) return [excl[0]];
    return hits; // 0..n (duplication gérée plus haut)
  }

  function _matchMember_(rows, ctx){
    var hits = rows.filter(function(x){
      if (x.Type!=='member') return false;
      if (x.Umin!=null && (ctx.U==null || ctx.U < x.Umin)) return false;
      if (x.Umax!=null && (ctx.U==null || ctx.U > x.Umax)) return false;
      if (x.Genre && x.Genre!=='*' && ctx.genre && x.Genre!==ctx.genre) return false;
      if (x.AliasContains) {
        var a = _norm(x.AliasContains);
        var raw = _norm(ctx.articleName||'');
        if (a && raw.indexOf(a)===-1) return false;
      }
      return true;
    });
    var excl = hits.filter(function(h){ return h.Exclude; });
    if (excl.length) return [excl[0]];
    return hits; // 0..n
  }

function _u2FromU_(U) {
  var s = (U || '').toString().trim();         // ex: "U7" ou "U10" ou "7"
  if (!s) return '';
  var m = s.match(/^[Uu]\s*(\d{1,2})$/) || s.match(/^(\d{1,2})$/);
  var d = m ? m[1] : '';
  if (!d) return '';
  return 'U' + (d.length === 1 ? '0' + d : d); // -> "U07" ou "U10"
}


function _renderFmt_(fmt, row, U, genreInit) {
  if (!fmt) return '';
  var g = _mt_extractGenreSmart_(row);
  var init = (genreInit==='F' ? 'F' :
             (genreInit==='M' ? 'M' :
             (genreInit==='X' ? 'X' : '')));
  var U2 = _u2FromU_(U);

  // Remplit {{genre}} (label), {{genreInitiale}} (M/F/X), {{U}}, {{U2}}, etc.
  return String(fmt)
    .replace(/{{\s*U2\s*}}/g, U2 || '')
    .replace(/{{\s*U\s*}}/g, U || '')
    .replace(/{{\s*genreInitiale\s*}}/g, init || g.initiale || '')
    .replace(/{{\s*genre\s*}}/g, g.label || '')
    .replace(/{{\s*article\s*}}/g, String(_articleNameFromRow_(row) || ''))
    .replace(/{{\s*saison\s*}}/g, String(row['Saison'] || ''))
    .replace(/{{\s*annee\s*}}/g, String((row['Saison'] || '').toString().replace(/.*(\d{4}).*$/,'$1')));
}

  function _deriveUFromRow_(r, ss){
    // 0) direct
    var u0 = r['Catégorie'] || r['Catégories'] || r['U'];
    if (u0) return String(u0);

    // 1) SEASON_YEAR - année_naissance
    var seasonYear = _seasonYearCache || (_seasonYearCache = _getSeasonYear_(ss));
    var by = _extractBirthYear_(r);
    if (by) {
      var udiff = seasonYear - by;
      if (udiff >= 4 && udiff <= 99) return 'U' + udiff;
    }

    // 2) fallback: "Uxx" dans libellé
    var txt = String(_articleNameFromRow_(r) || r['Catégorie texte'] || '').toUpperCase();
    var m = txt.match(/U\s*(\d{1,2})/);
    if (m) return 'U' + m[1];

    return '';
  }

  function _readActive_(ss, sheetName){
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return [];
    var data = sh.getDataRange().getValues();
    if (data.length<2) return [];
    var H = data[0];
    var rows = [];
    for (var i=1;i<data.length;i++){
      var row = {};
      for (var j=0;j<H.length;j++) row[H[j]] = data[i][j];
      // exclusions "soft" (export fera les vraies règles)
      if (row['CANCELLED']===true || row['EXCLUDE_FROM_EXPORT']===true) continue;
      rows.push(row);
    }
    return rows;
  }

  function _minify(m){
    return {
      Type:m.Type, AliasContains:m.AliasContains, Umin:m.Umin, Umax:m.Umax,
      Genre:m.Genre, Priority:m.Priority, Code:m.Code, ExclusiveGroup:m.ExclusiveGroup,
      GroupeFmt: m.GroupeFmt ? '✓' : '', CategorieFmt: m.CategorieFmt ? '✓' : '',
      Exclude: m.Exclude ? true : undefined
    };
  }
}
