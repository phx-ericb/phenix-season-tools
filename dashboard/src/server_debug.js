/** Reconstruit JOUEURS + ACHATS_LEDGER sans relancer d'import (mode FULL) */
function debugRebuildAggregatesFull() {
  var seasonId = (typeof readParamValue === 'function' && readParamValue('SEASON_ID')) || (typeof getSeasonId_ === 'function' && getSeasonId_());
  if (!seasonId) throw new Error('seasonId introuvable (PARAMS.SEASON_ID ou getSeasonId_).');

  if (!LIB || typeof LIB.runPostImportAugmentations !== 'function') {
    throw new Error('LIB.runPostImportAugmentations indisponible (vérifie l’export dans index.js).');
  }

  // isFull = true, isDryRun = false → on écrit vraiment
  LIB.runPostImportAugmentations(seasonId, [], { isFull: true, isDryRun: false });

var ss = SpreadsheetApp.openById(seasonId);
  var msg = _checkAggregates_(ss);
  Logger.log(msg);
  try { Browser.msgBox(msg); } catch (_) {}
}
/** Reconstruit partiellement (mode INCR) à partir d’un échantillon de passeports */
/** INCR debug: passeports ciblés (Array) ou un nombre à échantillonner */
function debugRebuildAggregatesIncr(passportsOrCount) {
  var seasonId = (typeof readParamValue === 'function' && readParamValue('SEASON_ID')) || (typeof getSeasonId_ === 'function' && getSeasonId_());
  if (!seasonId) throw new Error('seasonId introuvable.');
  var ss = SpreadsheetApp.openById(seasonId);
  var saison = readParam_(ss, 'SEASON_LABEL') || '';

  var normP = (typeof normalizePassport8_ === 'function')
    ? normalizePassport8_
    : function(p){ return (p == null || p === '') ? '' : String(p).replace(/\D/g,'').padStart(8,'0'); };

  // 1) Construire la liste "touched"
  var touched = [];
  if (Array.isArray(passportsOrCount) && passportsOrCount.length) {
    touched = passportsOrCount.map(normP);
  } else {
    var n = Number(passportsOrCount)||40;
    touched = _pickSomePassportsFallback_(ss, saison, n).map(normP);
  }
  touched = touched.filter(Boolean);
  if (!touched.length) { Logger.log('Aucun passeport sélectionné.'); return; }

  // 2) INCR: LEDGER -> JOUEURS
  if (typeof LIB !== 'undefined' && LIB && typeof LIB.runPostImportAugmentations === 'function') {
    LIB.runPostImportAugmentations(seasonId, touched, { isFull: false, isDryRun: false });
  } else {
    // fallback direct si la LIB n’est pas chargée
    if (typeof updateAchatsLedgerForPassports_ !== 'function' || typeof updateJoueursForPassports_ !== 'function') {
      throw new Error('INCR manquants (updateAchatsLedgerForPassports_/updateJoueursForPassports_)');
    }
    updateAchatsLedgerForPassports_(ss, touched);
    updateJoueursForPassports_(ss, touched);
  }

  var msg = (_checkAggregates_ ? _checkAggregates_(ss) : 'INCR exécuté') + ' | touched=' + touched.length;
  Logger.log(msg);
  try { Browser.msgBox(msg); } catch (_) {}
}

/** Fallback si _pickSomePassports_ n’existe pas : prend N passeports de la saison courante */
function _pickSomePassportsFallback_(ss, saison, n){
  n = n||40;
  var out = [];
  // Priorité JOUEURS (rapide), sinon LEDGER
  var shJ = ss.getSheetByName('JOUEURS');
  if (shJ) {
    var v = shJ.getDataRange().getValues(); var H=v[0]; var cP=H.indexOf('Passeport #'); var cS=H.indexOf('Saison');
    for (var r=1; r<v.length && out.length<n; r++){
      if (!v[r][cP]) continue;
      if (cS>=0 && saison && v[r][cS] !== saison) continue;
      out.push(String(v[r][cP]).padStart(8,'0'));
    }
    if (out.length) return out;
  }
  var shL = ss.getSheetByName('ACHATS_LEDGER');
  if (shL){
    var L = shL.getDataRange().getValues(); var HL=L[0]; var cLP=HL.indexOf('Passeport #'); var cLS=HL.indexOf('Saison');
    var seen = {};
    for (var r=1; r<L.length && out.length<n; r++){
      if (!L[r][cLP]) continue;
      if (cLS>=0 && saison && L[r][cLS] !== saison) continue;
      var p = String(L[r][cLP]).padStart(8,'0');
      if (!seen[p]){ seen[p]=1; out.push(p); }
    }
  }
  return out;
}


/** Prend N passeports distincts depuis INSCRIPTIONS (ou "Passeport #") */
function _pickSomePassports_(ssId, n) {
var ss = SpreadsheetApp.openById(seasonId);
 var sh = ss.getSheetByName('INSCRIPTIONS') || ss.getSheetByName('Inscriptions');
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  var hdr = vals[0];
  var ci = hdr.indexOf('Passeport #'); if (ci < 0) ci = hdr.indexOf('Passeport');
  if (ci < 0) return [];

  var out = [], seen = {};
  for (var r = 1; r < vals.length && out.length < (n || 50); r++) {
    var p = String(vals[r][ci] || '').trim();
    if (p && !seen[p]) { seen[p] = 1; out.push(p); }
  }
  return out;
}
/** Vérifie que les 2 feuilles sont bien peuplées (au moins 1 ligne de données) */
function _checkAggregates_(ss) {
  var a = ss.getSheetByName('ACHATS_LEDGER'), j = ss.getSheetByName('JOUEURS');
  var al = a ? a.getLastRow() : 0, jl = j ? j.getLastRow() : 0;
  var ok = (al > 1 && jl > 1);
  return 'ACHATS_LEDGER rows='+al+' | JOUEURS rows='+jl+' | OK='+(ok ? 'YES' : 'NO');
}
