/**
 * ===== Suivi Annulations (ACHATS_LEDGER + JOUEURS) =====
 * - Liste toutes les lignes annulées (inscriptions & articles)
 * - Affiche, pour chaque ligne annulée, les frais actifs restants
 * - Toggle "Corrigé" persistant (PropertiesService), clé par saison
 *
 * API exposées côté UI:
 *   - API_listAnnulations()
 *   - API_toggleAnnulationCorrige(key, val)
 *   - API_AN_invalidateCache()
 */

// ==== GLOBAL HELPERS (fallbacks si absents) ====
// -> évite "ReferenceError: AN_hdrMap_ is not defined" + calme VS Code/TS

var AN_hdrMap_ = this.AN_hdrMap_ || function(hdr){
  var m = {};
  for (var i=0;i<hdr.length;i++){
    var k = String(hdr[i]||'').replace(/\s+/g,' ').trim().toLowerCase();
    // normalise accents pour matcher les alias
    k = k.replace(/[éèêë]/g,'e').replace(/[àâ]/g,'a').replace(/[îï]/g,'i')
         .replace(/[ôö]/g,'o').replace(/[ûü]/g,'u');
    m[k] = i;
  }
  return m;
};

var AN_guessPassportCol_ = this.AN_guessPassportCol_ || function(rows){
  if (!rows || rows.length<2) return -1;
  var bestIdx = -1, bestScore = -1;
  for (var c=0;c<rows[0].length;c++){
    var score=0, seen=0;
    for (var r=1;r<Math.min(rows.length,200);r++){
      var v = rows[r][c]; if (v==null || v==='') continue; seen++;
      var s = String(v).trim().replace(/\s+/g,'');
      if (/^\d{6,12}$/.test(s)) score++;
      if (/^[A-Z]?\d{6,12}$/i.test(s)) score += 0.5;
    }
    if (seen>0){
      var ratio = score/seen;
      if (ratio>bestScore){ bestScore=ratio; bestIdx=c; }
    }
  }
  return (bestScore>=0.5) ? bestIdx : -1;
};

var AN_guessDateCol_ = this.AN_guessDateCol_ || function(rows){
  if (!rows || rows.length<2) return -1;
  for (var c=0;c<rows[0].length;c++){
    var hits=0, seen=0;
    for (var r=1;r<Math.min(rows.length,200);r++){
      var v = rows[r][c]; if (v==null || v==='') continue; seen++;
      if (v instanceof Date) hits++;
      else if (/^\d{4}-\d{2}-\d{2}$/.test(String(v))) hits++;
      else if (/^\d{1,2}[/\-]\d{1,2}[/\-]\d{2,4}$/.test(String(v))) hits++;
    }
    if (seen>0 && hits/seen>=0.5) return c;
  }
  return -1;
};

var AN_pickColSmart_ = this.AN_pickColSmart_ || function(H, rows, aliases, opts){
  opts = opts || {};
  // 1) alias exact (normalisé)
  for (var i=0;i<aliases.length;i++){
    var ak = String(aliases[i]||'').toLowerCase().trim()
      .replace(/[éèêë]/g,'e').replace(/[àâ]/g,'a')
      .replace(/[îï]/g,'i').replace(/[ôö]/g,'o').replace(/[ûü]/g,'u');
    if (H.hasOwnProperty(ak)) return H[ak];
  }
  // 2) contains
  var keys = Object.keys(H);
  for (var j=0;j<aliases.length;j++){
    var needle = String(aliases[j]||'').toLowerCase().trim()
      .replace(/[éèêë]/g,'e').replace(/[àâ]/g,'a')
      .replace(/[îï]/g,'i').replace(/[ôö]/g,'o').replace(/[ûü]/g,'u');
    for (var k=0;k<keys.length;k++){
      if (keys[k].indexOf(needle)>=0) return H[keys[k]];
    }
  }
  // 3) heuristiques
  if (opts.wantPassport){
    var g = AN_guessPassportCol_(rows);
    if (g>=0) return g;
  }
  if (opts.isDate){
    var g2 = AN_guessDateCol_(rows);
    if (g2>=0) return g2;
  }
  // 4) fallback safe
  return 0;
};

var AN_now_ = this.AN_now_ || function(){
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
};


function API_listAnnulations() {
  var seasonId = getSeasonId_();
  var cache = CacheService.getDocumentCache();
  var CK = 'AN_PAYLOAD_' + seasonId;
  var hit = cache.get(CK);
  if (hit) try { return JSON.parse(hit); } catch (_) {}

  var payload = AN_buildPayload_(seasonId);

  cache.put(CK, JSON.stringify(payload), 60); // 60s cache
  return payload;
}

function API_toggleAnnulationCorrige(key, val) {
  key = String(key || '');
  var seasonId = getSeasonId_();
  var dp = PropertiesService.getDocumentProperties();
  var RAW = dp.getProperty('AN_DONE_' + seasonId);
  var map = {};
  if (RAW) { try { map = JSON.parse(RAW) || {}; } catch (_) {} }
  if (val) map[key] = 1; else delete map[key];
  dp.setProperty('AN_DONE_' + seasonId, JSON.stringify(map));
  // invalider cache
  try { CacheService.getDocumentCache().remove('AN_PAYLOAD_' + seasonId); } catch (_){}
  return { ok: true };
}

function API_AN_invalidateCache() {
  try { CacheService.getDocumentCache().remove('AN_PAYLOAD_' + getSeasonId_()); } catch (_){}
  return { ok: true };
}

/** ===== Impl ===== */
function AN_buildPayload_(seasonId) {
  var ss = getSeasonSpreadsheet_(seasonId);

  // Feuilles nécessaires
  var shJ   = AN_getSheetAny_(ss, ['JOUEURS','Joueurs','Players']);
  var shIns = AN_getSheetAny_(ss, ['INSCRIPTIONS','inscriptions','Inscriptions']);
  var shArt = AN_getSheetAny_(ss, ['ARTICLES','articles','Articles']);

  if (!shJ) throw new Error('Feuille JOUEURS introuvable.');
  if (!shIns && !shArt) {
    return { ok:true, updatedAt: AN_now_(), counts:{total:0,aCorriger:0,corriges:0}, rows:[] };
  }

  // Persistance "Corrigé"
  var dp = PropertiesService.getDocumentProperties();
  var RAW = dp.getProperty('AN_DONE_' + seasonId);
  var doneMap = {};
  if (RAW) { try { doneMap = JSON.parse(RAW) || {}; } catch (_) {} }

  /* =================== JOUEURS: profil + ACTIFS via JSON =================== */
  var mapJ = {};           // pass -> {passeport, nom, prenom, dob}
  var activesByPass = {};  // pass -> [{type:'INSCRIPTION'|'ARTICLE', label:'...'}, ...]

  (function readJoueurs(){
    var J = shJ.getDataRange().getValues();
    if (!J || J.length<2) return;
    var H = AN_hdrMap_(J[0]);

    var cPass  = AN_pickColSmart_(H, J, ['Passeport #','Passeport','ID','passeport','id','passport','player id','member id'], {wantPassport:true});
    var cNom   = AN_pickColSmart_(H, J, ['Nom','Last Name','Family Name','nom'], {});
    var cPre   = AN_pickColSmart_(H, J, ['Prénom','Prenom','First Name','prenom','prénom'], {});
    var cDob   = AN_pickColSmart_(H, J, ['Naissance','Date de naissance','DOB','Birthdate','dob'], {isDate:true});
    var cInsJS = AN_pickColSmart_(H, J, ['InscriptionsJSON','inscriptionsjson','Inscriptions JSON','inscriptions_json'], {});
    var cArtJS = AN_pickColSmart_(H, J, ['ArticlesJSON','articlesjson','Articles JSON','articles_json'], {});

    for (var i=1;i<J.length;i++){
      var r = J[i];
      var pass = AN_toPlainPass_(r[cPass]); // "00123456"
      if (!pass) continue;

      mapJ[pass] = {
        passeport: pass,
        nom: r[cNom] || '',
        prenom: r[cPre] || '',
        dob: AN_fmtDateShort_(r[cDob])
      };

      var list = [];
      list = list.concat(AN_extractActivesFromJSON_(r[cInsJS], 'INSCRIPTION'));
      list = list.concat(AN_extractActivesFromJSON_(r[cArtJS], 'ARTICLE'));

      // dédoublonne par "type|label"
      if (list.length){
        var dedup = {};
        var out = [];
        for (var k=0;k<list.length;k++){
          var it = list[k]; var key = (it.type||'') + '|' + (it.label||'');
          if (!dedup[key]) { dedup[key]=1; out.push(it); }
        }
        activesByPass[pass] = out;
      }
    }
  })();

  /* =================== ANNULÉS: depuis INSCRIPTIONS + ARTICLES =================== */
  // pass|labelLower -> {pass, type, label, date}
  var canceledByKey = {};

  function isCancelledStatus_(s){
    s = String(s||'').toLowerCase();
    return s.indexOf('annul')>=0 || s.indexOf('cancel')>=0 || s.indexOf('refund')>=0 || s.indexOf('rembours')>=0;
  }

  function ingestCancelled(sheet, typeTag){
    if (!sheet) return;
    var A = sheet.getDataRange().getValues();
    if (!A || A.length<2) return;
    var H = AN_hdrMap_(A[0]);

    var cPass  = AN_pickColSmart_(H, A, ['Passeport #','Passeport','ID','passeport','id','passport','player id','member id'], {wantPassport:true});
    var cLabel = AN_pickColSmart_(H, A, ['Nom du frais','Nom','Élément','Element','Label','Fee','Item','Description','nom'], {});
    // attention aux apostrophes droites vs courbes
    var cStat  = AN_pickColSmart_(H, A, [
      "Statut de l'inscription","Statut de l’inscription","Statut inscription",
      'Statut','Status','statut','status','État','etat'
    ], {});
    var cDate  = AN_pickColSmart_(H, A, [
      "Date d'annulation","Date d’annulation",'Annulé le','Annulation','Cancelled At','cancelled at','Cancelled','Refunded At'
    ], {isDate:true});

    for (var i=1;i<A.length;i++){
      var r = A[i];
      var pass  = AN_toPlainPass_(r[cPass]);
      if (!pass) continue;

      var label = String(r[cLabel]||'').trim();
      if (!label) continue;

      var stat = r[cStat];
      var dateTxt = AN_fmtDateShort_(r[cDate]);

      if (isCancelledStatus_(stat) || (dateTxt && dateTxt !== '')) {
        var key = pass + '|' + label.toLowerCase();
        var cur = canceledByKey[key];
        // garde la date la plus récente si doublon
        if (!cur || String(dateTxt) > String(cur.date||'')) {
          canceledByKey[key] = { pass: pass, type: typeTag, label: label, date: dateTxt };
        }
      }
    }
  }

  ingestCancelled(shIns, 'INSCRIPTION');
  ingestCancelled(shArt, 'ARTICLE');

  /* =================== Assemblage final =================== */
  var rows = [];
  var keys = Object.keys(canceledByKey);
  for (var i=0;i<keys.length;i++){
    var it = canceledByKey[keys[i]];
    var prof = mapJ[it.pass] || { passeport: it.pass, nom:'', prenom:'', dob:'' };
    var act  = activesByPass[it.pass] || [];
    var rowKey = [it.pass, it.type||'', it.label||'', it.date||'-'].join('|');

    rows.push({
      key: rowKey,
      passeport: it.pass,
      nom: prof.nom || '',
      prenom: prof.prenom || '',
      dob: prof.dob || '',
      cancelled: { type: it.type || '', label: it.label || '', date: it.date || '' },
      active: act,                 // [{type,label},...] (depuis JOUEURS JSON)
      corrige: !!doneMap[rowKey]
    });
  }

  // Tri lisible
rows.sort(function(a,b){
  var da = String(a.cancelled.date||'');
  var db = String(b.cancelled.date||'');
  return db.localeCompare(da); // date décroissante
});


  var stats = {
    total: rows.length,
    aCorriger: rows.filter(function(x){ return !x.corrige; }).length,
    corriges:  rows.filter(function(x){ return  x.corrige; }).length
  };

  return {
    ok: true,
    updatedAt: AN_now_(),
    counts: stats,
    rows: rows
  };
}

/* ===== Helpers spécifiques à ce patch (si pas déjà dans le fichier) ===== */

// Utilise ta lib si dispo; sinon fallback local "plain 8"
function AN_toPlainPass_(val){
  try {
    if (typeof normalizePassportPlain8_ === 'function') {
      return normalizePassportPlain8_(val); // "00123456"
    }
  } catch(_){}
  var s = String(val==null?'':val).trim();
  if (!s) return '';
  if (s[0] === "'") s = s.slice(1);
  s = s.replace(/\s+/g,'');
  if (/^\d+$/.test(s) && s.length<8) s = ('00000000'+s).slice(-8);
  return s;
}

// Parse les JSON d'actifs dans JOUEURS et retourne [{type,label}]
function AN_extractActivesFromJSON_(cellValue, typeTag){
  var out = [];
  if (!cellValue) return out;
  var txt = String(cellValue).trim();
  if (!/^[\[\{]/.test(txt)) return out; // pas du JSON
  try {
    var data = JSON.parse(txt);
    if (Array.isArray(data)){
      for (var i=0;i<data.length;i++){
        var it = data[i];
        var label = null;
        if (typeof it === 'string') label = it.trim();
        else if (it && typeof it === 'object') {
          label = it.Produit || it.produit || it.Label || it.label || it.Nom || it.nom || null;
        }
        if (label) out.push({ type: typeTag, label: String(label).trim() });
      }
    } else if (data && typeof data === 'object') {
      var label2 = data.Produit || data.produit || data.Label || data.label || data.Nom || data.nom || null;
      if (label2) out.push({ type: typeTag, label: String(label2).trim() });
    }
  } catch(_){}
  return out;
}


/* === petits helpers ajoutés === */
function AN_getSheetAny_(ss, names){
  for (var i=0;i<names.length;i++){
    var sh = ss.getSheetByName(names[i]);
    if (sh) return sh;
  }
  return null;
}


/* ===== Helpers ===== */
function AN_idxHeaders_(hdr){
  var m = {}; for (var i=0;i<hdr.length;i++) m[String(hdr[i]||'').trim().toLowerCase()] = i;
  return m;
}
function AN_pickCol_(H, aliases){
  for (var i=0;i<aliases.length;i++){
    var k = String(aliases[i]).trim().toLowerCase();
    if (H.hasOwnProperty(k)) return H[k];
  }
  // par défaut: -1 (tolérance; l’appelant doit gérer)
  return -1;
}
function AN_normPass_(v){
  var s = String(v==null?'':v).trim();
  if (!s) return '';
  if (s[0] === "'") s = s.slice(1);
  s = s.replace(/\s+/g,'');
  if (/^\d+$/.test(s) && s.length < 8) s = ('00000000' + s).slice(-8);
  return s;
}
function AN_fmtDateShort_(d){
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var s = String(d||'').trim(); if (!s) return '';
  // si déjà texte "yyyy-mm-dd" on renvoie tel quel
  return s;
}
function AN_keyDate_(d){
  var s = AN_fmtDateShort_(d); return s || '-';
}

// Pré-chauffe (utilisé au boot UI)
function API_AN_warm(seasonId){
  try {
    // Memoize via DocumentCache (déjà implémenté dans API_listAnnulations)
    return API_listAnnulations();
  } catch (e) {
    return { ok:false, error:String(e && e.message || e) };
  }
}
