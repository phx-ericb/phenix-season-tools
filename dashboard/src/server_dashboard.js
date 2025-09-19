/** ---------- CONFIG ---------- */
const DASH_SHEETS = {
  INSCRIPTIONS : 'INSCRIPTIONS',
  ARTICLES     : 'ARTICLES',
  ERREURS      : 'ERREURS',
  OUTBOX       : 'MAIL_OUTBOX',
  IMPORT_LOG   : 'IMPORT_LOG'
};

function openSeasonSpreadsheet_(){
  // Utilise TA fonction existante fournie dans le projet
  const ssId = getSeasonId_();
  return SpreadsheetApp.openById(ssId);
}

/** Renvoie le nombre de lignes non vides (hors en-tête) d’une feuille si elle existe. */
function countRows_(ss, name){
  const sh = ss.getSheetByName(name);
  if(!sh) return 0;
  const last = sh.getLastRow();
  if(last <= 1) return 0;
  return last - 1; // on assume en-tête à la ligne 1
}

/** Lis un tableau (2D) avec header + rows. Renvoie {headers, rows} */
function readTable_(ss, name){
  const sh = ss.getSheetByName(name);
  if(!sh) return { headers: [], rows: [] };
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if(lr < 1 || lc < 1) return { headers: [], rows: [] };
  const values = sh.getRange(1,1,lr,lc).getValues();
  if(values.length === 0) return { headers: [], rows: [] };
  const headers = (values.shift() || []).map(h => String(h||'').trim());
  return { headers, rows: values };
}

function findCol_(headers, candidates){
  const set = headers.map(h => h.toString().trim().toLowerCase());
  for (const name of candidates){
    const i = set.indexOf(name.toLowerCase());
    if(i >= 0) return i;
  }
  return -1;
}

/**
 * KPI pour l’accueil (onglet Saison).
 */
function getDashboardMetrics(){
  const ss = openSeasonSpreadsheet_();

  const inscriptionsTotal = countRows_(ss, DASH_SHEETS.INSCRIPTIONS);
  const articlesTotal     = countRows_(ss, DASH_SHEETS.ARTICLES);


  // NOUVEAU: compte JOUEURS + LEDGER
  const joueursTotal = countRows_(ss, 'JOUEURS');
  const ledgerTotal  = countRows_(ss, 'ACHATS_LEDGER');

  // NOUVEAU: stats Photo depuis JOUEURS
  const jTab = readTable_(ss, 'JOUEURS'); // {headers, rows}
  const h = jTab.headers || [];
  const iPhotoStr = h.indexOf('PhotoStr');
  let photoExp = 0, photoSoon = 0;
  if (iPhotoStr >= 0) {
    jTab.rows.forEach(r => {
      const s = String(r[iPhotoStr]||'').toLowerCase();
      if (s.indexOf('expirée') !== -1 || s.indexOf('expiree') !== -1) photoExp++;
      else if (s.indexOf('expire bientôt') !== -1 || s.indexOf('expire bientot') !== -1) photoSoon++;
    });
  }


  // ERREURS
  let erreursTotal = 0, erreursDernierImport = null, errSubtitle = 'Dernier import';
  const errT = readTable_(ss, DASH_SHEETS.ERREURS);
  erreursTotal = errT.rows.length;

  if (errT.rows.length > 0){
    const dateCol = findCol_(errT.headers, ['date','timestamp','time','datetime']);
    if (dateCol >= 0){
      const map = {};
      errT.rows.forEach(r=>{
        const cell = r[dateCol];
        const iso = (cell instanceof Date)
          ? cell.toISOString().slice(0,10)
          : String(cell||'').trim();
        map[iso] = (map[iso]||0) + 1;
      });
      const dates = Object.keys(map).sort();
      const last = dates[dates.length-1];
      erreursDernierImport = map[last];
      errSubtitle = last ? `Le ${last}` : 'Dernier import';
    }
  }

  // OUTBOX
  let outboxPending = 0, outboxSubtitle = 'Prêts à l’envoi';
  const outT = readTable_(ss, DASH_SHEETS.OUTBOX);
  if (outT.rows.length){
    const statusCol = findCol_(outT.headers, ['status','etat','state']);
    const sentAtCol = findCol_(outT.headers, ['sent_at','envoye_le','sent','date_envoi']);
    outboxPending = outT.rows.filter(r=>{
      const st = statusCol >= 0 ? String(r[statusCol]||'').toUpperCase() : '';
      const sent = sentAtCol >= 0 ? r[sentAtCol] : '';
      const isSent = (st === 'SENT' || st === 'ENVOYE' || (sent && String(sent).trim() !== ''));
      return !isSent;
    }).length;
  }

  return {
    inscriptionsTotal,
    inscriptionsSubtitle: 'Total cumul.',
    articlesTotal,
    articlesSubtitle: 'Actifs',
    erreursTotal,
    erreursDernierImport,
    erreursSubtitle: errSubtitle,
    outboxPending,
    outboxSubtitle
  };
}

/**
 * Activité récente (10 lignes) : essaie IMPORT_LOG puis fallback ERREURS.
 */
// function getRecentActivity(){
//   const ss = openSeasonSpreadsheet_();

//   // IMPORT_LOG prioritaire
//   const logT = readTable_(ss, DASH_SHEETS.IMPORT_LOG);
//   if (logT.rows.length){
//     const cDate = findCol_(logT.headers, ['date','timestamp','time','datetime']);
//     const cType = findCol_(logT.headers, ['type','level','categorie']);
//     const cMsg  = findCol_(logT.headers, ['message','details','detail','info','texte']);
//     return logT.rows.slice(-10).reverse().map(r=>({
//       date: formatDate_(cDate>=0 ? r[cDate] : ''),
//       type: cType>=0 ? String(r[cType]||'') : '',
//       details: cMsg>=0 ? String(r[cMsg]||'') : ''
//     }));
//   }

//   // Fallback ERREURS
//   const errT = readTable_(ss, DASH_SHEETS.ERREURS);
//   if (errT.rows.length){
//     const cDate = findCol_(errT.headers, ['date','timestamp','time','datetime']);
//     const cType = findCol_(errT.headers, ['type','code','error','erreur']);
//     const cMsg  = findCol_(errT.headers, ['details','message','info','texte']);
//     return errT.rows.slice(-10).reverse().map(r=>({
//       date: formatDate_(cDate>=0 ? r[cDate] : ''),
//       type: cType>=0 ? String(r[cType]||'') : 'ERREUR',
//       details: cMsg>=0 ? String(r[cMsg]||'') : ''
//     }));
//   }

//   return [];
// }

function formatDate_(v){
  if (v instanceof Date) {
    const pad = n=>('0'+n).slice(-2);
    return `${v.getFullYear()}-${pad(v.getMonth()+1)}-${pad(v.getDate())} ${pad(v.getHours())}:${pad(v.getMinutes())}`;
  }
  return String(v||'');
}

function getMembreFlagsMetrics() {
  var ss = getSeasonSpreadsheet_(getSeasonId_());
  var sh = ss.getSheetByName(readParam_ ? readParam_(ss, 'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL' : 'MEMBRES_GLOBAL');
  if (!sh || sh.getLastRow() < 2) return { photosInvalides: 0, casiersExpires: 0, total: 0 };

  var vals = sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  var header = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var col = {}; header.forEach(function(h,i){ col[String(h)]=i; });

  var ciPhotoInv = col['PhotoInvalide'] ?? -1;
  var ciCasier   = col['CasierExpiré'] ?? -1;

  var photosInvalides = 0, casiersExpires = 0, total = vals.length;
  for (var i=0;i<vals.length;i++){
    if (ciPhotoInv>=0 && Number(vals[i][ciPhotoInv])===1) photosInvalides++;
    if (ciCasier>=0   && Number(vals[i][ciCasier])===1)   casiersExpires++;
  }
  return { photosInvalides: photosInvalides, casiersExpires: casiersExpires, total: total };
}

/** KPIs par type (entraineurs | joueurs) basés sur MEMBRES_GLOBAL,
 *  avec la source "inscriptions" pour déterminer les joueurs réellement inscrits.
 */
function getKpiPhotosCasierByType(type /* 'entraineurs' | 'joueurs' */) {
  var seasonId = getSeasonId_();
  var ss = SpreadsheetApp.openById(seasonId);

  var seasonYear = Number(readParam_(ss, 'SEASON_YEAR') || new Date().getFullYear());
  var invalidFrom = (readParam_(ss, 'PHOTO_INVALID_FROM_MMDD') || '04-01').trim(); // ex "04-01"
  var cutoffNextJan1 = (seasonYear + 1) + '-01-01';        // ex "2026-01-01"
  var seasonInvalidDate = seasonYear + '-' + invalidFrom;  // ex "2025-04-01"
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // 1) Construire l’ensemble des passeports "inscrits", selon le type
  var passportsSet = new Set();

  if (type === 'entraineurs') {
    // À partir de ENTRAINEURS_ROLES (unique par passeport)
    var shR = ss.getSheetByName('ENTRAINEURS_ROLES');
    if (shR && shR.getLastRow() > 1) {
      var R = shR.getDataRange().getValues();
      var h = R[0];
      var ciPass = h.indexOf('Passeport');
      for (var i=1;i<R.length;i++){
        var p = normalizePassportPlain8_(R[i][ciPass]);
        if (p) passportsSet.add(p);
      }
    }
  } else {
    // JOUEURS : on lit la feuille finale "inscriptions"
    var shJ = ss.getSheetByName('inscriptions');
    if (shJ && shJ.getLastRow() > 1) {
      var J = shJ.getDataRange().getValues();
      var hj = J[0]; 
      var jp = hj.indexOf('Passeport');      // colonne attendue
      var js = hj.indexOf('Statut');         // optionnel : si tu veux filtrer "inscrit", "actif", etc.
      for (var j=1;j<J.length;j++){
        if (jp < 0) break;
        // si tu veux forcer un statut, décommente la ligne suivante:
        // if (js >= 0 && !/inscrit|actif/i.test(String(J[j][js]||''))) continue;
        var pj = normalizePassportPlain8_(J[j][jp]);
        if (pj) passportsSet.add(pj);
      }
    }
  }

  // 2) Parcours MEMBRES_GLOBAL pour les passeports retenus
  var sh = ss.getSheetByName(readParam_(ss,'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL');
  if (!sh || sh.getLastRow() < 2) return { photosInvalides:0, dues:0, casiersExpires:0, total:0 };

  var V = sh.getDataRange().getValues();
  var H = V[0];
  var cPass  = H.indexOf('Passeport'),
      cPhoto = H.indexOf('PhotoExpireLe'),
      cCas   = H.indexOf('CasierExpiré');

  var photosInvalides = 0, dues = 0, casiersExpires = 0, total = 0;

  for (var r=1; r<V.length; r++){
    var p = normalizePassportPlain8_(V[r][cPass]);
    if (!p) continue;

    // Si on a un set (entraineurs ou joueurs) -> filtrer
    if (passportsSet.size && !passportsSet.has(p)) continue;

    total++;
    var exp   = String(V[r][cPhoto] || '');
    var inval = (exp && exp < cutoffNextJan1) ? 1 : 0; // même règle que l’import
    if (inval) {
      if (today >= seasonInvalidDate) photosInvalides++;
      else dues++; // invalide mais "à renouveler" (due) à partir du 1er avril
    }

    var cas = Number(V[r][cCas] || 0);
    if (cas === 1) casiersExpires++;
  }

  return { photosInvalides: photosInvalides, dues: dues, casiersExpires: casiersExpires, total: total };
}
