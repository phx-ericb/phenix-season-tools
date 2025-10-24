/** =========================================================
 * server_ententes.js — Module "Ententes de paiement"
 * Source de vérité : ACHATS_LEDGER (pivot)
 * Registre échéanciers : feuille "Ententes"
 * ========================================================= */

/* ================== UTILITAIRES COMPAT / FALLBACKS ================== */
// Normalisation clé (accents -> ASCII, casse, ponctuation)
var EN_norm = this.ER_norm || function (s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '').trim();
};

// Résolution d'ID saison (aligne avec server_errors)
var EN_resolveSeasonId_ = this.ER_resolveSeasonId_ || function (overrideId) {
  if (overrideId) return overrideId;
  if (typeof getSeasonId_ === 'function') { var sid = getSeasonId_(); if (sid) return sid; }
  var sidProp = PropertiesService.getScriptProperties().getProperty('SEASON_SHEET_ID');
  if (sidProp) return sidProp;
  throw new Error("Aucun ID de saison : passe un seasonId, implémente getSeasonId_(), ou définis SEASON_SHEET_ID.");
};

// Ouverture du classeur saison (aligne avec server_errors)
var EN_openSeasonSpreadsheet_ = this.ER_openSeasonSpreadsheet_ || function (seasonId) {
  if (typeof getSeasonSpreadsheet_ === 'function') { var ss = getSeasonSpreadsheet_(seasonId); if (ss) return ss; }
  return SpreadsheetApp.openById(seasonId);
};

// Passeport : préférer ta normalisation centrale si dispo
function EN_normPass_(v) {
  if (typeof normalizePassportPlain8_ === 'function') return normalizePassportPlain8_(v);
  if (typeof normalizePassportToText8_ === 'function') {
    var t = normalizePassportToText8_(v);
    return (t && t[0]==="'") ? t.slice(1) : t;
  }
  var s = String(v == null ? '' : v).trim(); if (!s) return '';
  if (s[0] === "'") s = s.slice(1);
  if (/^\d+$/.test(s) && s.length<8) s = ('00000000'+s).slice(-8);
  return s;
}

// parse monnaie FR/EN tolérant
function EN_parseCurrency_(str) {
  if (str == null) return 0;
  var s = String(str).trim().replace(/\u202f|\s/g, '');
  var lastC = s.lastIndexOf(','), lastD = s.lastIndexOf('.');
  if (lastC > -1 && lastD > -1) {
    if (lastC > lastD) s = s.replace(/\./g, '').replace(/,/g, '.'); // 1.234,56
    else s = s.replace(/,/g, '');                                   // 1,234.56
  } else if (lastC > -1) {
    s = s.replace(/,/g, '.'); // 1234,56
  }
  s = s.replace(/[^0-9.]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function EN_todayYmd_() {
  var d = new Date();
  var mm = String(d.getMonth() + 1).padStart(2, '0');
  var dd = String(d.getDate()).padStart(2, '0');
  return d.getFullYear() + '-' + mm + '-' + dd;
}

/* ================== FEUILLE "ENTENTES" ================== */
var ENTENTES_SHEET_NAME = 'Ententes';
var ENT_HEADERS = [
  'Passeport #','Prénom','Nom','DDN','Catégorie',
  'Date paiement prévu','Montant prévu','Montant reçu',
  'Paid Baseline','Statut','Commentaire','Entré par','Date entrée'
];

function ENT_bootstrapSheet_(ss) {
  var sh = ss.getSheetByName(ENTENTES_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(ENTENTES_SHEET_NAME);
    sh.getRange(1,1,1,ENT_HEADERS.length).setValues([ENT_HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }
  // S’assurer que toutes les colonnes existent (sinon ajout à droite)
  var h = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(String);
  var missing = ENT_HEADERS.filter(function(x){ return h.indexOf(x) < 0; });
  if (missing.length) {
    sh.getRange(1, h.length+1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

/* ================== LECTURE PIVOT ACHATS_LEDGER ================== */
function EN_buildHeaderIndex_(headers) {
  var map = {}; headers.forEach(function(h,i){ map[EN_norm(h)] = i; });
  function pick(){ for (var i=0;i<arguments.length;i++){ var k=EN_norm(arguments[i]); if (k in map) return map[k]; } return -1; }
  return {
    Passeport: pick('Passeport','Passeport #','Passport','ID'),
    Prenom:    pick('Prénom','Prenom','First name','FirstName','First'),
    Nom:       pick('Nom','Nom de famille','Last name','LastName','Surname','Family Name'),
    Dob:       pick('DDN','Date de naissance','Date of Birth','DOB'),
    Due:       pick('AmountDue','Montant dû','Montant du','Due'),
    Paid:      pick('AmountPaid','Montant payé','Paid'),
    Rem:       pick('AmountRestant','Montant dû restant','Remaining','Balance','Solde'),
    Active:    pick('Status','Actif','Active','IsActive'),
    PayStatus: pick('PaymentStatus','Statut paiement','Statut'),
FeeLabel:  pick(
  'NomFrais','Nom Frais','Nom_du_frais','Nom des frais',
  'Nom du frais','Frais','Article','Nom de l\'article',
  'Fee','Fee Name','Item','Item Name','Description','Libellé'
)
  };
}



function ENT_readBalancesFromLedger_(ss) {
  var sh = ss.getSheetByName('ACHATS_LEDGER');
  if (!sh) return [];
  var n = sh.getLastRow(), m = sh.getLastColumn();
  if (n < 2) return [];
  var values = sh.getDataRange().getDisplayValues();
  var H = EN_buildHeaderIndex_(values[0].map(String));
  var outMap = {}; // pass -> agg

  for (var r=1; r<values.length; r++) {
    var row = values[r];

    // --- garder seulement les frais actifs ---
    var activeRaw = (H.Active >= 0 ? String(row[H.Active]).trim() : '1');
    if (!(activeRaw === '1' || activeRaw.toLowerCase() === 'true')) continue;

    var pass = EN_normPass_(row[H.Passeport]);
    if (!pass) continue;

    var agg = outMap[pass] || (outMap[pass] = {
      passport: pass,
      firstName: String(row[H.Prenom] || ''),
      lastName:  String(row[H.Nom] || ''),
      dob:       String(row[H.Dob] || ''),
      due: 0, paid: 0, remaining: 0,
      statuses: {},
      fees: [] // libellés uniques
    });

    agg.due       += EN_parseCurrency_(row[H.Due]);
    agg.paid      += EN_parseCurrency_(row[H.Paid]);
    agg.remaining += EN_parseCurrency_(row[H.Rem]);

    var pst = (H.PayStatus>=0) ? String(row[H.PayStatus]||'').trim() : '';
    if (pst) agg.statuses[pst] = true;

var lbl = (H.FeeLabel>=0) ? String(row[H.FeeLabel]||'').trim() : '';
if (lbl && agg.fees.indexOf(lbl) < 0) agg.fees.push(lbl);

  }

  return Object.keys(outMap).map(function(k){
    var a = outMap[k];
    var due  = +(a.due.toFixed(2));
    var paid = +(a.paid.toFixed(2));
    var rem  = +(a.remaining.toFixed(2));
    var outstanding = Math.max(due - paid, rem);
    return {
      passport: a.passport,
      firstName: a.firstName,
      lastName:  a.lastName,
      dob:       a.dob,
      totalDue:  due,
      totalPaid: paid,
      outstanding: +(outstanding.toFixed(2)),
      activeFees: a.fees || [],
      needsPlan: (outstanding > 0) || !!a.statuses.Partial || !!a.statuses.Unpaid
    };
  });
}


/* ================== LECTURE ÉCHÉANCES EXISTANTES ================== */
function ENT_readAll_(ss) {
  var sh = ENT_bootstrapSheet_(ss);
  var n = sh.getLastRow(), m = sh.getLastColumn();
  if (n < 2) return [];
  var data = sh.getRange(2,1,n-1,m).getDisplayValues();
  var Hrow = sh.getRange(1,1,1,m).getDisplayValues()[0].map(String);
  var idx = {}; Hrow.forEach(function(h,i){ idx[h] = i; });
  function ci(name){ return (idx[name] != null ? idx[name] : -999) + 1; }

  var cPass = ci('Passeport #'),
      cFn   = ci('Prénom'),
      cLn   = ci('Nom'),
      cDob  = ci('DDN'),
      cDate = ci('Date paiement prévu'),
      cAmt  = ci('Montant prévu'),
      cRec  = ci('Montant reçu'),
      cBase = ci('Paid Baseline'),
      cStat = ci('Statut'),
      cCom  = ci('Commentaire');

  var out = [];
  for (var i=0;i<data.length;i++){
    var r = data[i];
    out.push({
      rowNum: i+2,
      passport: EN_normPass_(r[cPass-1]),
      firstName: r[cFn-1]||'',
      lastName:  r[cLn-1]||'',
      dob:       r[cDob-1]||'',
      dateStr:   r[cDate-1]||'', // JJ/MM/AAAA
      amount:    EN_parseCurrency_(r[cAmt-1]),
      received:  EN_parseCurrency_(r[cRec-1]),
      baseline:  EN_parseCurrency_(r[cBase-1]),
      status:    r[cStat-1]||'',
      comment:   r[cCom-1]||''
    });
  }
  return out;
}

/* ================== API: LISTE (fusion ledger + ententes) ================== */

function ENT_enrichNamesFromJoueurs_(ss, arr) {
  try {
    var sh = ss.getSheetByName('JOUEURS'); if (!sh) return;
    var n=sh.getLastRow(), m=sh.getLastColumn(); if (n<2) return;
    var data = sh.getDataRange().getDisplayValues();
    var heads = data.shift().map(String);
    var norm = heads.map(EN_norm);
    function find(){ for (var i=0;i<arguments.length;i++){ var k=EN_norm(arguments[i]); var j=norm.indexOf(k); if (j>=0) return j; } return -1; }
    var iP = find('Passeport','Passeport #','Passport');
    var iFn= find('Prénom','Prenom','First name','FirstName');
    var iLn= find('Nom','Last name','LastName');
    if (iP<0) return;
    var map = {};
    data.forEach(function(r){
      var p = EN_normPass_(r[iP]); if (!p) return;
      map[p] = {fn:r[iFn]||'', ln:r[iLn]||''};
    });
    arr.forEach(function(b){
      if (b && map[b.passport]) {
        if (!b.firstName) b.firstName = map[b.passport].fn;
        if (!b.lastName ) b.lastName  = map[b.passport].ln;
      }
    });
  } catch(_){}
}


function API_ententes_list(seasonId) {
  var sid = EN_resolveSeasonId_(seasonId);
  var ss  = EN_openSeasonSpreadsheet_(sid);

  var balances = ENT_readBalancesFromLedger_(ss);
  ENT_enrichNamesFromJoueurs_(ss, balances); // <— fallback noms

  var ents = ENT_readAll_(ss);
  // index ententes par passeport
  var byP = {};
  ents.forEach(function(e){
    (byP[e.passport] = byP[e.passport] || []).push(e);
  });
  Object.keys(byP).forEach(function(p){
    byP[p].sort(function(a,b){
      var da = String(a.dateStr||'').split('/').reverse().join('');
      var db = String(b.dateStr||'').split('/').reverse().join('');
      return da<db?-1:da>db?1:0;
    });
  });

  // fusion & filtrage
  var out = balances
    .filter(function(b){ return b.needsPlan || (byP[b.passport] && byP[b.passport].length); })
    .map(function(b){
      var ent = byP[b.passport] || [];
      var scheduled = ent.reduce(function(s,e){ return s + (e.amount||0); }, 0);
return {
  passport: b.passport,
  firstName: b.firstName,
  lastName:  b.lastName,
  dob:       b.dob,
  outstanding: b.outstanding,
  totalDue:  b.totalDue,
  totalPaid: b.totalPaid,
  activeFees: b.activeFees || [],   // ← ajout
  ent: ent,
  scheduled: +(scheduled.toFixed(2)),
  gap: +((b.outstanding - scheduled).toFixed(2))
};

    });

  return { ok:true, members: out };
}

/* ================== API: CRÉATION D’UN PLAN ================== */
// payload = { passport, payments:[{date:'JJ/MM/AAAA', amount}], comment }
function API_ententes_savePlan(seasonId, payload) {
  var sid = EN_resolveSeasonId_(seasonId);
  var ss  = EN_openSeasonSpreadsheet_(sid);
  var sh  = ENT_bootstrapSheet_(ss);

  var byPass = {};
  ENT_readBalancesFromLedger_(ss).forEach(function(b){ byPass[b.passport] = b; });

  var me = Session.getActiveUser().getEmail() || 'n/a';
  var now = new Date();

  var p = EN_normPass_(payload && payload.passport);
  if (!p) throw new Error('Passeport manquant');
  var base = byPass[p] || { firstName:'', lastName:'', dob:'', totalPaid:0 };

  var rows = [];
  var paidBaseline = base.totalPaid || 0;
  (payload.payments||[]).forEach(function(pay){
    if (!pay || !pay.date) return;
    var amt = +(+pay.amount||0).toFixed(2); if (amt<=0) return;
    rows.push([
      p, base.firstName, base.lastName, base.dob, '',
      pay.date, amt,
      0,                   // Montant reçu
      paidBaseline,        // Paid Baseline (snapshot)
      'Ouvert',
      payload.comment||'',
      me, now
    ]);
  });

  if (rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,ENT_HEADERS.length).setValues(rows);
  return { ok:true, added: rows.length };
}

/* ================== API: MISE À JOUR D’UNE LIGNE ================== */
function API_ententes_updateItem(seasonId, rowNum, patch) {
  var sid = EN_resolveSeasonId_(seasonId);
  var ss  = EN_openSeasonSpreadsheet_(sid);
  var sh  = ENT_bootstrapSheet_(ss);
  var H   = sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0].map(String);
  function ci(name){ return H.indexOf(name)+1; }

  if (patch.date)         sh.getRange(rowNum, ci('Date paiement prévu')).setValue(patch.date);
  if (patch.amount!=null) sh.getRange(rowNum, ci('Montant prévu')).setValue(+(+patch.amount).toFixed(2));
  if (patch.comment!=null)sh.getRange(rowNum, ci('Commentaire')).setValue(String(patch.comment||''));
  return { ok:true };
}

/* ================== API: SUPPRESSION D’UNE LIGNE ================== */
function API_ententes_deleteItem(seasonId, rowNum) {
  var sid = EN_resolveSeasonId_(seasonId);
  var ss  = EN_openSeasonSpreadsheet_(sid);
  var sh  = ENT_bootstrapSheet_(ss);
  sh.deleteRow(rowNum);
  return { ok:true };
}

/* ================== API: RÉCONCILIATION DELTA-ONLY + STATUTS ================== */
function API_ententes_detectAndSync(seasonId) {
  var sid = EN_resolveSeasonId_(seasonId);
  var ss  = EN_openSeasonSpreadsheet_(sid);
  var sh  = ENT_bootstrapSheet_(ss);

  var n = sh.getLastRow(), m = sh.getLastColumn();
  if (n < 2) return { ok:true, updated:0 };

  var Hrow = sh.getRange(1,1,1,m).getDisplayValues()[0].map(String);
  function ci(name){ return Hrow.indexOf(name)+1; }

  var range = sh.getRange(2,1,n-1,m);
  var data  = range.getDisplayValues();
  var vals  = range.getValues(); // écriture batch

  var cPass = ci('Passeport #')-1,
      cDate = ci('Date paiement prévu')-1,
      cAmt  = ci('Montant prévu')-1,
      cRec  = ci('Montant reçu')-1,
      cBase = ci('Paid Baseline')-1,
      cStat = ci('Statut')-1;

  // Totaux payés actuels (ledger)
  var paidNow = {};
  ENT_readBalancesFromLedger_(ss).forEach(function(b){ paidNow[b.passport] = b.totalPaid || 0; });

  // Regrouper lignes par passeport
  var byP = {};
  for (var i=0;i<data.length;i++){
    var p = EN_normPass_(data[i][cPass]);
    if (!p) continue;
    (byP[p] = byP[p] || []).push({ i:i });
  }
  // Trier par date JJ/MM/AAAA
  Object.keys(byP).forEach(function(p){
    byP[p].sort(function(a,b){
      var da = String(data[a.i][cDate]||'').split('/').reverse().join('');
      var db = String(data[b.i][cDate]||'').split('/').reverse().join('');
      return da<db?-1:da>db?1:0;
    });
  });

  var todayYmd = EN_todayYmd_();
  var updates = 0;

  Object.keys(byP).forEach(function(p){
    var rows = byP[p];
    var totPaidNow = paidNow[p] || 0;

    for (var k=0;k<rows.length;k++){
      var iRow = rows[k].i;

      var amt  = EN_parseCurrency_(data[iRow][cAmt]);
      var rec  = EN_parseCurrency_(data[iRow][cRec]);
      var base = EN_parseCurrency_(data[iRow][cBase]);

      var delta = Math.max(0, totPaidNow - base);
      if (delta > 0) {
        var dueRem = Math.max(0, amt - rec);
        var assign = Math.min(dueRem, delta);
        if (assign > 0) {
          vals[iRow][cRec]  = +(rec + assign).toFixed(2);
          vals[iRow][cBase] = +(base + assign).toFixed(2);
          delta -= assign;
          updates++;
        }
      }

      // Statut (après éventuelle assignation)
      var newRec = EN_parseCurrency_(vals[iRow][cRec]);
      var dateJJ = String(data[iRow][cDate]||'');
      var ymd = '';
      if (dateJJ && dateJJ.indexOf('/')>-1) {
        var parts = dateJJ.split('/');
        ymd = parts[2] + '-' + String(parts[1]).padStart(2,'0') + '-' + String(parts[0]).padStart(2,'0');
      }
      var status = 'Ouvert';
      if (newRec >= amt - 1e-9) status = 'Payé';
      else if (newRec > 0) status = 'Partiel';
      if (status !== 'Payé' && ymd && ymd < todayYmd) status = 'En retard';

      vals[iRow][cStat] = status;
    }
  });

  range.setValues(vals);
  return { ok:true, updated: updates };
}
