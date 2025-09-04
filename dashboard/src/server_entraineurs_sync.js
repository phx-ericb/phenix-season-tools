/** ===== Entraîneurs: vue matérialisée INSCRIPTIONS_ENTRAINEURS =====
 * - Additif: ne modifie rien ailleurs.
 * - À appeler après import global et/ou import membres.
 * - Détection coach: utilise isCoachMember_(ss,row) si présent (lib), sinon fallback robuste.
 * - MEMBRES_GLOBAL: utilise getMembresIndex_(ss) si présent (lib), sinon fallback local.
 */

// ---------------- Fallbacks légers (au cas où la lib n’est pas chargée) ----------------
if (typeof SHEETS === 'undefined') {
  var SHEETS = { INSCRIPTIONS:'INSCRIPTIONS' };
}
if (typeof getSeasonSpreadsheet_ !== 'function') {
  function getSeasonSpreadsheet_(id){ return SpreadsheetApp.openById(id); }
}
if (typeof normalizePassportPlain8_ !== 'function') {
  function normalizePassportPlain8_(v){
    var s = String(v == null ? '' : v).replace(/\D/g,'').trim();
    if (!s) return '';
    s = s.slice(-8);
    while (s.length < 8) s = '0' + s;
    return s;
  }
}
if (typeof readParam_ !== 'function') {
  function readParam_(ss, key){
    var sh = ss.getSheetByName('PARAMS');
    if (!sh || sh.getLastRow() < 1) return '';
    var data = sh.getRange(1,1,sh.getLastRow(),2).getDisplayValues();
    for (var i=0;i<data.length;i++) if ((data[i][0]+'').trim()===key) return (data[i][1]+'').trim();
    return '';
  }
}

// MEMBRES_GLOBAL index: privilégie la version lib, sinon fallback local (sans cache)
if (typeof getMembresIndex_ !== 'function') {
  function getMembresIndex_(ss){
    var name = readParam_(ss,'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL';
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return { byPass:{} };
    var V = sh.getDataRange().getValues();
    var H = V[0].map(String); var col={}; H.forEach(function(h,i){ col[h]=i; });
    function gi(n){ return (typeof col[n]==='number') ? col[n] : -1; }
    var cPass=gi('Passeport'), cPh=gi('PhotoExpireLe'), cInv=gi('PhotoInvalide'),
        cCas=(gi('CasierExpiré')>=0?gi('CasierExpiré'):gi('CasierExpire')),
        cUpd=gi('LastUpdate'), cDue=gi('PhotoInvalideDuesLe'), cStat=gi('StatutMembre');
    var byPass = Object.create(null);
    for (var r=1;r<V.length;r++){
      var p8 = normalizePassportPlain8_(V[r][cPass]); if (!p8) continue;
      var photoInv = String(V[r][cInv]||'').toLowerCase();
      var casVal   = String(V[r][cCas]||'').toLowerCase();
      byPass[p8] = {
        Passeport: p8,
        PhotoExpireLe: (cPh>=0? String(V[r][cPh]||'') : ''),
        PhotoInvalide: (photoInv==='true'||photoInv==='1') ? 1 : 0,
        CasierExpire:  (casVal==='true'||casVal==='1') ? 1 : 0,
        PhotoInvalideDuesLe: (cDue>=0? String(V[r][cDue]||'') : ''),
        LastUpdate: (cUpd>=0? String(V[r][cUpd]||'') : ''),
        StatutMembre: (cStat>=0? String(V[r][cStat]||'') : '')
      };
    }
    return { byPass: byPass };
  }
}

// ---------------- Helpers locaux ----------------
function _se_norm_(s){
  return String(s||'').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
function _se_isActiveRow_(r){
  var can = String(r['__cancelled']||r['Annulé']||r['Annule']||r['Cancelled']||'').toLowerCase()==='true';
  var exc = String(r['__exclude_from_export']||r['EXCLUDE_FROM_EXPORT']||'').toLowerCase()==='true';
  var st  = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
  return !can && !exc && st!=='annulé' && st!=='annule' && st!=='cancelled';
}
function _se_pass8_(raw){
  try{ return normalizePassportPlain8_(raw); }catch(_){ return String(raw||'').trim(); }
}
function _se_ensureSheetWithHeader_(ss, name, headers){
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  } else {
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (lr < 1 || lc < headers.length) {
      sh.clear();
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    } else {
      // Vérifie l'entête, sinon la réécrit proprement
      var cur = sh.getRange(1,1,1,headers.length).getDisplayValues()[0].map(String);
      var need = headers.some(function(h,i){ return (cur[i]||'') !== h; });
      if (need) sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
  }
  // Colonne Passeport en texte
  sh.getRange('A:A').setNumberFormat('@');
  return sh;
}

// Statut photo depuis une date d’expiration (YYYY-MM-DD)
function _se_computePhotoStatus_(ss, expStr){
  var seasonYear = Number(readParam_(ss, 'SEASON_YEAR') || new Date().getFullYear());
  var mmdd = (readParam_(ss, 'PHOTO_INVALID_FROM_MMDD') || '04-01').trim();
  var cutoffNextJan1 = (seasonYear + 1) + '-01-01';
  var dueDate = seasonYear + '-' + mmdd;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var exp = String(expStr||'').trim();
  // si aucune date -> invalide (à renouv/échue selon dueDate)
  if (!exp) return (today >= dueDate ? 'ÉCHUE' : 'À RENOUVELER');
  // si expire avant le 1er janv de l’année suivante -> invalide
  var invalid = (exp < cutoffNextJan1);
  if (!invalid) return 'OK';
  return (today >= dueDate ? 'ÉCHUE' : 'À RENOUVELER');
}

function _se_casierLabel_(flag){
  var s = String(flag==null?'':flag).toLowerCase();
  var on = (s==='1'||s==='true'||s==='oui'||s==='yes');
  return on ? 'EXPIRÉ' : 'OK';
}

// Détection “coach” : priorise la fonction lib si dispo; sinon param CSV + filet lexical
function _isCoachMemberSafe_(ss, row){
  try { if (typeof isCoachMember_ === 'function') return !!isCoachMember_(ss, row); } catch(_){}
  var fee = (row && (row['Nom du frais']||row['Frais']||row['Produit'])) || '';
  var v = _se_norm_(fee);
  if (!v) return false;
  var csv = (typeof readParam_==='function') ? (readParam_(ss,'RETRO_COACH_FEES_CSV')||'') : '';
  var toks = csv.split(',').map(_se_norm_).filter(Boolean);
  if (toks.length){
    if (toks.indexOf(v)>=0) return true;
    for (var i=0;i<toks.length;i++) if (v.indexOf(toks[i])>=0) return true;
  }
  return /(entraineur|entra[îi]neur|coach)/i.test(String(fee||''));
}

// ---------------- Sync principal ----------------
function SR_syncInscriptionsEntraineurs_(ss){
  // 1) Lire INSCRIPTIONS (finales)
  var insc = (typeof readSheetAsObjects_==='function')
    ? readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS)
    : { rows: ss.getSheetByName(SHEETS.INSCRIPTIONS).getDataRange().getDisplayValues().slice(1) /* fallback très brut */ };

  // 2) Index MEMBRES_GLOBAL (via lib si possible -> cache)
  var mg = getMembresIndex_(ss); // { byPass: {...} }

  // 3) Filtrer coachs actifs -> 1 ligne par passeport (dernier vu)
  var coachByPass = Object.create(null);
  (insc.rows||[]).forEach(function(r){
    if (!_se_isActiveRow_(r)) return;
    if (!_isCoachMemberSafe_(ss, r)) return;

    var p8 = _se_pass8_(r['Passeport #']||r['Passeport']);
    if (!p8) return;

    coachByPass[p8] = {
      Passeport:  p8,
      Nom:        r['Nom']||'',
      Prenom:     r['Prénom']||r['Prenom']||'',
      NomComplet: (((r['Prénom']||r['Prenom']||'') + ' ' + (r['Nom']||'')).trim()),
      Saison:     r['Saison']||'',
      Frais:      String(r['Nom du frais'] || r['Frais'] || r['Produit'] || '')
    };
  });

  // 4) Construire la sortie (sans U)
  var HEADERS = [
    'Passeport','Nom','Prénom','NomComplet','Saison','Frais',
    'Exp_Photo','StatutPhoto','Casier','Membre_LastUpdate','SyncedAt'
  ];

  var out = [];
  Object.keys(coachByPass).sort().forEach(function(p8){
    var base = coachByPass[p8];
    var info = (mg && mg.byPass && mg.byPass[p8]) ? mg.byPass[p8] : {};
    var expPhoto = info.PhotoExpireLe || '';
    out.push([
      base.Passeport,
      base.Nom,
      base.Prenom,
      base.NomComplet,
      base.Saison,
      base.Frais,
      expPhoto,
      _se_computePhotoStatus_(ss, expPhoto),
      _se_casierLabel_(info.CasierExpire || 0),
      info.LastUpdate || '',
      new Date()
    ]);
  });

  // 5) Écrire/rafraîchir la feuille
  var sh = _se_ensureSheetWithHeader_(ss, 'INSCRIPTIONS_ENTRAINEURS', HEADERS);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2,1,last-1, sh.getLastColumn()).clearContent();
  if (out.length){
    sh.getRange(2,1,out.length, out[0].length).setValues(out);
  }

  return { total: out.length };
}

// ---------------- Runner (UI / post-import) ----------------
function runSyncInscriptionsEntraineurs(seasonSheetId){
  return _wrap('runSyncInscriptionsEntraineurs', function(){
    var sid = seasonSheetId || (typeof getSeasonId_==='function' ? getSeasonId_() : null);
    if (!sid) throw new Error('seasonSheetId manquant');
    var ss = getSeasonSpreadsheet_(sid);
    var res = SR_syncInscriptionsEntraineurs_(ss);
    appendImportLog_(ss, 'COACHS_SYNC_DONE', 'rows='+res.total);
    return res;
  });
}
