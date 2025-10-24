/** =========================================================
 * server_inscriptions_overview.js — Vue d’ensemble Inscriptions
 * Source : ACHATS_LEDGER (lignes actives seulement)
 * Jointure JOUEURS pour Genre / U / ProgramBand si manquants
 * Sortie : par Secteur -> Type=Inscription -> M/F/Total
 * Cache 60s (DocumentCache)
 * ========================================================= */

/* ---------- Utils ---------- */
function OV_normAccents_(s){
  return String(s||'')
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toLowerCase().replace(/\s+/g,' ').trim();
}
function OV_hdrMap_(hdr){
  var m={}; for (var i=0;i<hdr.length;i++){ m[OV_normAccents_(hdr[i])] = i; } return m;
}
function OV_pick_(H, aliases, def){
  for (var i=0;i<aliases.length;i++){
    var k = OV_normAccents_(aliases[i]||'');
    if (H.hasOwnProperty(k)) return H[k];
    // légère tolérance (contains)
    var keys = Object.keys(H);
    for (var j=0;j<keys.length;j++){
      if (keys[j].indexOf(k)>=0) return H[keys[j]];
    }
  }
  return (def==null ? -1 : def);
}
function OV_pass8_(v){
  var s = String(v==null?'':v).trim();
  if (!s) return '';
  if (s[0]==="'") s = s.slice(1);
  s = s.replace(/\s+/g,'');
  if (/^\d+$/.test(s) && s.length<8) s = ('00000000'+s).slice(-8);
  return s;
}

/* ---------- Secteur ---------- */
function OV_sectorFrom_(uTxt, bandTxt, feeLabel){
  var u = String(uTxt||'').toUpperCase();
  var band = String(bandTxt||'').toUpperCase();
  var lbl = String(feeLabel||'').toUpperCase();

  // Élite prioritaire
  if (band.indexOf('ELITE')>=0 || lbl.indexOf('PLSJQ')>=0 || lbl.indexOf('LDP')>=0 || lbl.indexOf('ÉLITE')>=0 || lbl.indexOf('ELITE')>=0){
    return 'Élite';
  }
  // Essayer d'extraire le nombre d'U
  var m = u.match(/U?\s*(\d{1,2})/i);
  var x = m ? +m[1] : NaN;
  if (!isNaN(x)){
    if (x>=4 && x<=8)  return 'U4-U8';
    if (x>=9 && x<=12) return 'U9-U12';
    if (x>=13 && x<=18) return 'U13-U18';
  }
  // Fallback par band
  if (band.indexOf('U4')>=0 || band.indexOf('U5')>=0 || band.indexOf('U8')>=0) return 'U4-U8';
  if (band.indexOf('U9')>=0 || band.indexOf('U12')>=0) return 'U9-U12';
  if (band.indexOf('U13')>=0 || band.indexOf('U18')>=0) return 'U13-U18';

  return 'U9-U12'; // défaut
}

/* ---------- JOUEURS (enrichissement) ---------- */
function OV_readJoueursQuick_(ss){
  var sh = ss.getSheetByName('JOUEURS'); if (!sh) return {};
  var n = sh.getLastRow(); if (n<2) return {};
  var values = sh.getDataRange().getValues();
  var H = OV_hdrMap_(values[0]);
  var cPass = OV_pick_(H, ['passeport #','passeport','passport','id']);
  var cGenre= OV_pick_(H, ['genre','sex','sexe']);
  var cU    = OV_pick_(H, ['u','categorie','catégorie','agebracket','age bracket']);
  var cBand = OV_pick_(H, ['programband','band','programme','program']);
  var map = {};
  for (var r=1;r<values.length;r++){
    var row = values[r];
    var p = OV_pass8_(row[cPass]); if (!p) continue;
    var g = String(row[cGenre]||'').toUpperCase();
    map[p] = {
      genre: g.startsWith('F') ? 'F' : (g.startsWith('M') ? 'M' : ''),
      u: String(row[cU]||''),
      band: String(row[cBand]||'')
    };
  }
  return map;
}

/* ---------- ACHATS_LEDGER ---------- */
function OV_readLedgerInscriptionRows_(ss){
  var sh = ss.getSheetByName('ACHATS_LEDGER'); if (!sh) return [];
  var n = sh.getLastRow(); if (n<2) return [];
  var values = sh.getDataRange().getDisplayValues();
  var H = OV_hdrMap_(values[0]);

  var cPass   = OV_pick_(H, ['passeport #','passeport','passport','id']);
  var cActive = OV_pick_(H, ['actif','active','isactive','status']); // 1 | true
  var cType   = OV_pick_(H, ['type','category','categorie','catégorie']);
  var cLabel  = OV_pick_(H, ['nom du frais','nom frais','fee','item','description','article','libellé','label']);
  var cGenre  = OV_pick_(H, ['genre','sex','sexe'], -1);
  var cU      = OV_pick_(H, ['u','categorie','catégorie','agebracket','age bracket'], -1);
  var cBand   = OV_pick_(H, ['programband','band','programme','program'], -1);

  var out = [];
  for (var r=1;r<values.length;r++){
    var row = values[r];
    var activeRaw = (cActive>=0 ? String(row[cActive]).trim() : '1');
    var isActive = (activeRaw === '1' || /^true$/i.test(activeRaw));
    if (!isActive) continue;

    var t = String(cType>=0 ? row[cType] : '').toUpperCase();
    var label = String(cLabel>=0 ? row[cLabel] : '').trim();
    if (!(t.indexOf('INSCRIPTION')>=0 || /inscription/i.test(label))) continue;

    out.push({
      pass: OV_pass8_(row[cPass]),
      genre: (cGenre>=0 ? String(row[cGenre]||'') : '').toUpperCase(),
      u:     (cU>=0 ? String(row[cU]||'') : ''),
      band:  (cBand>=0 ? String(row[cBand]||'') : ''),
      label: label
    });
  }
  return out;
}

/* ---------- Agrégation ---------- */
function OV_aggregate_(rows, joueursMap){
  // { [sector]: { totals:{M,F,total}, fees:{ [label]:{M,F,total} } } }
  var S = {};
  function ensure(sec){
    if (!S[sec]) S[sec] = { totals:{M:0,F:0,total:0}, fees:{} };
    return S[sec];
  }

  for (var i=0;i<rows.length;i++){
    var it = rows[i]; if (!it.pass) continue;
    var jo = joueursMap[it.pass] || {};
    var genre = it.genre || jo.genre || '';
    genre = (genre==='F'||genre==='M') ? genre : '';

    var sec = OV_sectorFrom_(it.u || jo.u, it.band || jo.band, it.label);
    var b = ensure(sec);
    var feeKey = it.label || 'Inscription';
    if (!b.fees[feeKey]) b.fees[feeKey] = { M:0, F:0, total:0 };

    b.fees[feeKey].total++;
    b.totals.total++;
    if (genre==='M'){ b.fees[feeKey].M++; b.totals.M++; }
    else if (genre==='F'){ b.fees[feeKey].F++; b.totals.F++; }
  }
  return S;
}

/* ---------- API ---------- */
function API_OV_getSummary(seasonId){
  var cache = CacheService.getDocumentCache();
  var sid = seasonId || (typeof getSeasonId_==='function' ? getSeasonId_() : '');
  var CK = 'OV_INS_SUMMARY_' + sid;

  try {
    var hit = cache.get(CK);
    if (hit) return JSON.parse(hit);
  } catch (_){}

  var ss = (typeof getSeasonSpreadsheet_==='function') ? getSeasonSpreadsheet_(sid) : SpreadsheetApp.openById(sid);

  var joueursMap = OV_readJoueursQuick_(ss);
  var rows = OV_readLedgerInscriptionRows_(ss);
  var sectorsAgg = OV_aggregate_(rows, joueursMap);

  var sectors = [
    { id:'U4-U8',   data: sectorsAgg['U4-U8']   || { totals:{M:0,F:0,total:0}, fees:{} } },
    { id:'U9-U12',  data: sectorsAgg['U9-U12']  || { totals:{M:0,F:0,total:0}, fees:{} } },
    { id:'U13-U18', data: sectorsAgg['U13-U18'] || { totals:{M:0,F:0,total:0}, fees:{} } },
    { id:'Élite',   data: sectorsAgg['Élite']   || { totals:{M:0,F:0,total:0}, fees:{} } }
  ];

  var g = {M:0,F:0,total:0};
  for (var i=0;i<sectors.length;i++){
    var t = sectors[i].data.totals;
    g.M += t.M; g.F += t.F; g.total += t.total;
  }

  var payload = {
    ok: true,
    updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
    sectors: sectors,
    grandTotals: g
  };

  try { cache.put(CK, JSON.stringify(payload), 60); } catch(_){}
  return payload;
}

function API_OV_invalidateCache(){
  try {
    var sid = (typeof getSeasonId_==='function') ? getSeasonId_() : '';
    CacheService.getDocumentCache().remove('OV_INS_SUMMARY_' + sid);
  } catch(_){}
  return { ok:true };
}
