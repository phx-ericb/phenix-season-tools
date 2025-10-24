/** ======================== JOUEURS – API UI =========================
 * - Lecture JOUEURS (mapping tolérant + heuristique passeport)
 * - Pagination + filtres + recherche
 * - Détail joueur (activités ACHATS_LEDGER)
 * - Caches cohérents (ScriptCache)
 */

/* ---------- Helpers ---------- */
var JO_norm = this.JO_norm || function (s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().replace(/[^a-z0-9]+/g,'').trim();
};

// 👉 Utilise les helpers partagés si dispo (évite les divergences entre modules)
// 👉 Résolution de la saison robuste:
//  - 1) getSeasonId_() si dispo
//  - 2) ACTIVE_SEASON_ID (nouveau flux server_seasons)
//  - 3) SEASON_SHEET_ID (legacy)
var JO_resolveSeasonId_ = this.ER_resolveSeasonId_ || function (overrideId) {
  if (overrideId) return overrideId;

  if (typeof getSeasonId_ === 'function') {
    var sid = null;
    try { sid = getSeasonId_(); } catch(_) {}
    if (sid) return sid;
  }

  var props = PropertiesService.getScriptProperties();
  var sidProp = props.getProperty('ACTIVE_SEASON_ID') || props.getProperty('SEASON_SHEET_ID');
  if (!sidProp) throw new Error('Aucune saison active (ACTIVE_SEASON_ID manquant).');
  return sidProp;
};

var JO_openSeasonSpreadsheet_ = this.ER_openSeasonSpreadsheet_ || function (seasonId) {
  if (typeof getSeasonSpreadsheet_ === 'function') { var ss=getSeasonSpreadsheet_(seasonId); if (ss) return ss; }
  return SpreadsheetApp.openById(seasonId);
};

function JO_normPass_(v){
  if (typeof normalizePassportPlain8_ === 'function') return normalizePassportPlain8_(v);
  if (typeof normalizePassportToText8_ === 'function') {
    var t=normalizePassportToText8_(v);
    return (t && t[0]==="'")?t.slice(1):t;
  }
  var s=String(v==null?'':v).trim();
  if (!s) return '';
  if (s[0]==="'") s=s.slice(1);
  s = s.replace(/\s+/g,'');
  if (/^\d+$/.test(s) && s.length<8) s=('00000000'+s).slice(-8);
  return s;
}

/* ---------- Index JOUEURS ---------- */
function JO_buildHeaderIndex_(headers){
  var map = {}; headers.forEach(function(h,i){ map[JO_norm(h)] = i; });
  function pick(){ for (var i=0;i<arguments.length;i++){ var k=JO_norm(arguments[i]); if (k in map) return map[k]; } return -1; }
  return {
    Passeport: pick('Passeport','Passeport #','Passport','ID','Player ID','Member ID','Passeport#','ID membre'),
    Nom: pick('Nom','Last name','Famille'),
    Prenom: pick('Prénom','Prenom','First name'),
    Band: pick('Band','ProgramBand','ProgrammeBand','Program Band','U','AgeBracket'),
    Adapte: pick('Adapté','Adapte','Adapté?','Adapte?','isAdapte'),
    Photo: pick('Photo','PhotoStr','Statut Photo','Photo status','Photo Statut'),
    PhotoDate: pick('Photo (expire)','PhotoExpireLe','Expire Photo','Photo Expire','Expiration photo','Photo Date','Date Photo'),
    Courriels: pick('Courriels','Courriel','Email','Emails','E-mail')
  };
}
function JO_val_(row, idx, key){ var i=idx[key]; return i>=0 ? row[i] : ''; }

/* ---------- Heuristique si l’entête “Passeport” est introuvable ---------- */
function JO_guessPassportCol_(values){
  if (!values || values.length<2) return -1;
  var best = { col:-1, score:-1 };
  var rows = Math.min(values.length-1, 300);
  var cols = Math.min(values[0].length, 40);
  for (var c=0;c<cols;c++){
    var seen=0, hits=0;
    for (var r=1;r<=rows;r++){
      var v = values[r][c];
      if (v==null || v==='') continue;
      seen++;
      var s = String(v).trim().replace(/\s+/g,'');
      if (/^'?([A-Za-z]?\d{6,12})$/.test(s)) hits++;
    }
    if (seen>0){
      var score = hits/seen;
      if (score>best.score){ best={ col:c, score:score }; }
    }
  }
  return (best.score>=0.5) ? best.col : -1;
}

/* ---------- Lecture JOUEURS (avec cache + fallback heuristique) ---------- */
function JO_cacheKey_(sid){ return 'JO_ALL_v2::'+sid; }
function JO_readAll_(seasonId){
  var sid = JO_resolveSeasonId_(seasonId);
  var cache = CacheService.getScriptCache(); var key = JO_cacheKey_(sid);
  var hit = cache.get(key); if (hit) { try { return JSON.parse(hit); } catch(_){ } }

  var ss = JO_openSeasonSpreadsheet_(sid);
  var sh = ss.getSheetByName('JOUEURS'); if (!sh) throw new Error("Onglet 'JOUEURS' introuvable.");
  var values = sh.getDataRange().getValues(); if (!values || !values.length) return { headers: [], rows: [] };
  var headers = (values[0]||[]).map(String);
  var idx = JO_buildHeaderIndex_(headers);

  // ⛑️ Fallback si l’entête “Passeport” n’a pas été trouvée
  if (idx.Passeport < 0){
    var guessed = JO_guessPassportCol_(values);
    if (guessed >= 0){
      // injecte la “découverte” pour cette lecture
      var tmp = {}; for (var k in idx) tmp[k] = idx[k];
      tmp.Passeport = guessed;
      idx = tmp;
    }
  }

  var rows = [];
  for (var r=1;r<values.length;r++){
    var row = values[r]; if (!row || row.every(function(c){ return c===''||c==null; })) continue;
    var pass = JO_normPass_(JO_val_(row,idx,'Passeport')); if (!pass) continue;
    var pd = JO_val_(row,idx,'PhotoDate');
    var pdOut = (pd instanceof Date)
      ? Utilities.formatDate(pd, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(pd||'');
    rows.push({
      passeport: pass,
      nom: String(JO_val_(row,idx,'Nom')||''),
      prenom: String(JO_val_(row,idx,'Prenom')||''),
      band: String(JO_val_(row,idx,'Band')||''),
      adapte: String(JO_val_(row,idx,'Adapte')||''),
      photo: String(JO_val_(row,idx,'Photo')||''),
      photoDate: pdOut,
      courriels: String(JO_val_(row,idx,'Courriels')||'')
    });
  }

  var payload = { headers: headers, rows: rows };
  try { cache.put(key, JSON.stringify(payload), 120); } catch(_){}
  return payload;
}

/* ---------- LEDGER & APIs: inchangés (sauf cache key déjà v2 plus haut) ---------- */
function LEDGER_cacheKey_(sid){ return 'LED_ALL_'+sid; }
function LEDGER_readAll_(seasonId){
  var sid = JO_resolveSeasonId_(seasonId);
  var cache = CacheService.getScriptCache(); var key=LEDGER_cacheKey_(sid);
  var hit = cache.get(key); if (hit) { try { return JSON.parse(hit); } catch(_){ } }

  var ss = JO_openSeasonSpreadsheet_(sid);
  var sh = ss.getSheetByName('ACHATS_LEDGER'); if (!sh) return { headers:[], rows:[] };
  var values = sh.getDataRange().getValues(); if (!values || !values.length) return { headers:[], rows:[] };
  var headers = values[0].map(String);
  var H = {}; headers.forEach(function(h,i){ H[JO_norm(h)] = i; });
  function col(){ for (var i=0;i<arguments.length;i++){ var k=JO_norm(arguments[i]); if (k in H) return H[k]; } return -1; }
  var I = { Passeport: col('Passeport','Passeport #','ID','Passport'), Type: col('Type','type'), Nom: col('Nom','Nom du frais','Label','Description'), Band: col('Band','U','AgeBracket'), Tags: col('Tags'), Date: col('Date','Date de la facture') };

  var rows = [];
  for (var r=1;r<values.length;r++){
    var row = values[r]; if (!row) continue;
    var p = row[I.Passeport]; if (p===''||p==null) continue;
    rows.push({
      passeport: JO_normPass_(p),
      type: String(I.Type>=0 ? row[I.Type] : ''),
      nom: String(I.Nom>=0 ? row[I.Nom] : ''),
      band: String(I.Band>=0 ? row[I.Band] : ''),
      tags: String(I.Tags>=0 ? row[I.Tags] : ''),
      date: (I.Date>=0 && row[I.Date] instanceof Date)
        ? Utilities.formatDate(row[I.Date], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(I.Date>=0 ? row[I.Date] : '')
    });
  }
  var out = { headers: headers, rows: rows };
  try { cache.put(key, JSON.stringify(out), 600); } catch(_){}
  return out;
}
function LEDGER_byPassport_(seasonId, pass){
  var wanted = JO_normPass_(pass);
  var all = LEDGER_readAll_(seasonId);
  return (all.rows||[]).filter(function(r){ return JO_normPass_(r.passeport) === wanted; });
}

/* ---------- APIs ---------- */
function API_listJoueursPage(opt){
  opt = opt || {};
  var offset = Math.max(0, +opt.offset || 0);
  var limit  = Math.min(500, Math.max(1, +opt.limit || 50));
  var search = String(opt.search||'').toLowerCase();
  var bandF  = String(opt.band||'').toLowerCase();
  var adapteF= String(opt.adapte||'');
  var photoF = String(opt.photo||'');

  var data = JO_readAll_(); var rows = (data && data.rows) ? data.rows.slice() : [];

  function normBoolish(v){ var s=String(v||'').toLowerCase(); return (s==='1'||s==='oui'||s==='true'||s==='vrai'||s==='yes')?1:(s==='0'||s==='non'||s==='false'||s==='faux'||s==='no')?0:null; }
  function photoStatus(r){
    var s=String(r.photo||'').toLowerCase(); var d = new Date(r.photoDate);
    var now=new Date();
    if (!r.photo || /^\s*$/.test(String(r.photo))) return 'missing';
    if (s.includes('expir')||s.includes('échue')||s.includes('echue')) return 'expired';
    if (!isNaN(+d)){ if (d<now) return 'expired'; var diff=(d-now)/(24*3600*1000); if (diff<=30) return 'soon'; return 'ok'; }
    return s||'ok';
  }

  var out = rows.filter(function(r){
    if (search){ var hay=[r.passeport,r.nom,r.prenom,r.courriels].join(' ').toLowerCase(); if (!hay.includes(search)) return false; }
    if (bandF){ if (String(r.band||'').toLowerCase()!==bandF) return false; }
    if (adapteF){ var v=normBoolish(r.adapte); if (adapteF==='1' && v!==1) return false; if (adapteF==='0' && v!==0) return false; }
    if (photoF){ var ps=photoStatus(r); if (photoF!==ps) return false; }
    return true;
  });

  var total = out.length; var page = out.slice(offset, offset+limit);
  return { total: total, limit: limit, rows: page };
}

function API_getJoueurDetailLoose(passeport){
  var sid = JO_resolveSeasonId_();
  var wanted = JO_normPass_(passeport);

  var data = JO_readAll_(); var rows = data.rows || []; var pos = -1;
  for (var i=0;i<rows.length;i++){ if (JO_normPass_(rows[i].passeport)===wanted){ pos=i; break; } }
  if (pos<0) return { notFound:true, tried:[wanted, "'"+wanted], sid:sid };

  var activites = LEDGER_byPassport_(sid, wanted);
  return { joueur: rows[pos], activites: activites, sid: sid };
}

// 👉 Warmer JO non-bloquant avec garde + cache court
function API_JO_warm() {
  if (typeof _getFlag_ === 'function' && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
  return { ok:false, busy:true, at: Date.now() };
}

  var sid;
  try { sid = JO_resolveSeasonId_(); } catch (e) {
    // pas de saison → on ne bloque pas, l’UI pourra réessayer
    return { ok:false, reason:'no-season', at:Date.now() };
  }

  var cache = CacheService.getScriptCache();
  var K_WARM = 'JO_WARM_v3_' + sid;
  var hit = cache.get(K_WARM);
  if (hit) { return { ok:true, cached:true, at:Date.now() }; }

  // anti-rafale: si un warm est déjà en vol, on répond "busy" (l’UI réessaie plus tard)
  var GUARD = 'JO_WARM_INFLIGHT_' + sid;
  if (cache.get(GUARD)) {
    return { ok:false, busy:true, at:Date.now() };
  }
  try { cache.put(GUARD, '1', 15); } catch(_) {}

  // petit lock pour sérialiser l’ouverture du classeur; si indisponible → busy, pas de blocage
  var lock = null;
  try {
    
if (_getFlag_ && _getFlag_('PHENIX_IMPORT_LOCK') === '1') {
  return { ok:false, busy:true, reason:'import-running' };
}


    // Charge JOUEURS (met en cache 120 s dans JO_readAll_)
    var a = JO_readAll_(sid);
    // Préchauffe aussi le LEDGER (cache 600 s)
    LEDGER_readAll_(sid);

    // marque le warm comme “fini” pour 60 s (assez pour le boot)
    try { cache.put(K_WARM, '1', 60); } catch(_){}
    return { ok:true, warmed:true, rows:(a && a.rows ? a.rows.length : 0), at:Date.now() };

  } catch (e) {
    return { ok:false, error:String(e), at:Date.now() };

  } finally {
    try { cache.remove(GUARD); } catch(_){}
    if (lock) { try { lock.releaseLock(); } catch(_){ } }
  }
}

// 👉 Pense à invalider aussi les clés du warm
function API_JO_invalidateCache(seasonId){
  var sid = JO_resolveSeasonId_(seasonId);
  var c = CacheService.getScriptCache();
  c.remove(JO_cacheKey_(sid));
  c.remove(LEDGER_cacheKey_(sid));
  c.remove('JO_WARM_v3_' + sid);
  c.remove('JO_WARM_INFLIGHT_' + sid);
  return { ok:true };
}