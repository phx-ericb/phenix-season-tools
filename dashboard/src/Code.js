/** =========================================================================
 *  Dashboard runners pour la librairie Phénix (v1.0)
 *  - AUCUN appel à getActiveSpreadsheet() (dashboard non lié à un Sheets)
 *  - Passe toujours l’ID du classeur saison à la librairie
 *  - Fournit les endpoints UI: getRecentActivity(), getDashboardMetrics()
 *  - Logger unifié appendImportLog_ (1 ou 3 arguments)
 * ========================================================================= */
/* ============================== Code.js — v1.1 ==============================
 * - Aligne les runners UI sur les nouveaux exporteurs «*ToDrive»
 * - Ajoute une variante incrémentale qui passe directement onlyPassports
 * - Le flow runImportAndExports déclenche déjà l’incrémental automatiquement
 *   car les exporteurs lisent LAST_TOUCHED_PASSPORTS s’il n’y a pas d’options.
 */


function seedSeasonYearOnce() {
  var id = getSeasonId_(); 
  setParamValue('SEASON_YEAR', 2025);
  setParamValue('RETRO_MEMBER_MAX_U', 18); // optionnel
}

/// Alias vers la lib (ajuste "SI" si ton alias est différent)
var LIB = SI && SI.Library ? SI.Library : null;

/** ======================== Config de la cible saison ======================== */
/** 1) Option A : définis la constante ci-dessous et basta */
var SEASON_SHEET_ID = '1IVVHi17Jyo8jvWtrSuenbPW8IyEZqlY1bXx-WbnXPkk'; // ← colle l'ID du classeur saison ici ou laisse vide

/** 2) Option B : stocke l’ID en Script Property une fois pour toutes
 *    exécute setSeasonSheetIdOnce() une seule fois puis laisse SEASON_SHEET_ID vide
 */
function setSeasonSheetIdOnce() {
  var id = '1IVVHi17Jyo8jvWtrSuenbPW8IyEZqlY1bXx-WbnXPkk';
  var props = PropertiesService.getScriptProperties();
  props.setProperty('PHENIX_SEASON_SHEET_ID', id);
  props.setProperty('ACTIVE_SEASON_ID', id);

  // Ajoute au registre si absent (pour l’UI)
  var ss = SpreadsheetApp.openById(id);
  var list = JSON.parse(props.getProperty('SEASONS_JSON') || '[]');
  if (!list.some(function(s){ return s.id === id; })) {
    list.push({ id:id, title:ss.getName(), url:ss.getUrl() });
    props.setProperty('SEASONS_JSON', JSON.stringify(list));
  }
}

/** Récupère l’ID du classeur saison (constante > ScriptProperty) */
function getSeasonId_() {
  var props = PropertiesService.getScriptProperties();
  var id =
    props.getProperty('ACTIVE_SEASON_ID') ||
    (SEASON_SHEET_ID && String(SEASON_SHEET_ID).trim()) ||
    props.getProperty('PHENIX_SEASON_SHEET_ID') ||
    props.getProperty('SEASON_SPREADSHEET_ID');

  if (!id) {
    throw new Error(
      "Aucun ID de classeur saison. Définis SEASON_SHEET_ID " +
      "ou exécute setSeasonSheetIdOnce() / clique 'Définir active' dans l'UI."
    );
  }
  return String(id).trim();
}

/** ============================ Runners principaux ============================ */
/** Importer/mettre à jour les données (utilise la lib) */
function runImporterDonneesSaison() {
  if (!LIB || typeof LIB.importerDonneesSaison !== 'function') {
    throw new Error('Fonction importerDonneesSaison indisponible dans la lib.');
  }
  return LIB.importerDonneesSaison(getSeasonId_());
}

/** Export XLSX — Rétro : Membres (auto: utilise LAST_TOUCHED_PASSPORTS si présent) */
function runExportRetroMembres() {
  if (!LIB || typeof LIB.exportRetroMembresXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroMembresXlsxToDrive indisponible dans la lib.');
  }
  // Fallback auto: sans options, la lib lira LAST_TOUCHED_PASSPORTS → incrémental
  return LIB.exportRetroMembresXlsxToDrive(getSeasonId_());
}

/** Export XLSX — Rétro : Membres (FORCÉ incrémental) */
function runExportRetroMembresIncr() {
  if (!LIB || typeof LIB.exportRetroMembresXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroMembresXlsxToDrive indisponible dans la lib.');
  }
  var list = (typeof getLastTouchedPassports_==='function') ? getLastTouchedPassports_() : [];
  return LIB.exportRetroMembresXlsxToDrive(getSeasonId_(), { onlyPassports: list });
}
/** Export XLSX — Rétro : Groupes (ALL = Groupes + GroupeArticles) */
function runExportRetroGroupes() {
  if (!LIB || typeof LIB.exportRetroGroupesAllXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroGroupesAllXlsxToDrive indisponible dans la lib.');
  }
  return LIB.exportRetroGroupesAllXlsxToDrive(getSeasonId_());
}

/** Export XLSX — Rétro : Groupes (ALL) FORCÉ incrémental */
function runExportRetroGroupesIncr() {
  if (!LIB || typeof LIB.exportRetroGroupesAllXlsxToDrive !== 'function') {
    throw new Error('Fonction exportRetroGroupesAllXlsxToDrive indisponible dans la lib.');
  }
  var list = (typeof getLastTouchedPassports_==='function') ? getLastTouchedPassports_() : [];
  return LIB.exportRetroGroupesAllXlsxToDrive(getSeasonId_(), { onlyPassports: list });
}


/** Lecture simple d’un param dans PARAMS (utilisée par runImportAndExports pour DRY_RUN) */
function readParamValue(key) {
  var ss = SpreadsheetApp.openById(getSeasonId_());
  var sh = ss.getSheetByName('PARAMS');
  if (!sh || sh.getLastRow() < 2) return '';
  var vals = sh.getRange(2,1, sh.getLastRow()-1, 2).getValues(); // Col A=Clé, B=Valeur
  for (var i=0;i<vals.length;i++) {
    if (String(vals[i][0]||'') === key) return String(vals[i][1]||'');
  }
  return '';
}


/** Appliquer les règles (remplit ERREURS) */
function runEvaluateRules() {
  if (!LIB || typeof LIB.evaluateSeasonRules !== 'function') {
    throw new Error('Fonction evaluateSeasonRules indisponible dans la lib.');
  }
  return LIB.evaluateSeasonRules(getSeasonId_());
}

// ============ Code.js ============
// Runner unique : Import -> Exports rétro (Membres + Groupes ALL) -> Log
// (on garde les exports ici SEULEMENT pour éviter les doublons)
function runImportAndExports(){
  var out = { ok:false, steps: [] };
  appendImportLog_({ type:'RUN_IMPORT_START', details:'Déclenchement via dashboard' });

  try {
    // 1) Import
    if (!LIB || typeof LIB.importerDonneesSaison !== 'function') {
      throw new Error('Fonction importerDonneesSaison indisponible dans la lib.');
    }
    var impRes = LIB.importerDonneesSaison(getSeasonId_());
    out.steps.push({ step: 'import', res: impRes });

    // 2) Exports rétro — auto (incr si LAST_TOUCHED_PASSPORTS existe)
    try {
      var mRes = runExportRetroMembres();
      appendImportLog_({ type:'EXPORT_RETRO_MEMBRES_OK', details: JSON.stringify(mRes) });
      out.steps.push({ step:'export_membres', res:mRes });
    } catch(eM){ appendImportLog_({ type:'EXPORT_RETRO_MEMBRES_FAIL', details: String(eM) }); }

    try {
      var gRes = runExportRetroGroupes();
      appendImportLog_({ type:'EXPORT_RETRO_GROUPES_ALL_OK', details: JSON.stringify(gRes) });
      out.steps.push({ step:'export_groupes_all', res:gRes });
    } catch(eG){ appendImportLog_({ type:'EXPORT_RETRO_GROUPES_ALL_FAIL', details: String(eG) }); }


  var ss = SpreadsheetApp.openById(getSeasonId_());
try {
  var x = SR_syncInscriptionsEntraineurs_(ss);
  appendImportLog_(ss, 'COACHS_SYNC_DONE', 'rows=' + (x && x.total || 0));
} catch(e) {
  appendImportLog_(ss, 'COACHS_SYNC_FAIL', String(e));
}



    out.ok = true;
    return out;
  } finally {
    appendImportLog_({ type:'RUN_IMPORT_END', details:'Terminé' });
  }
  
}

/** Code.gs — wrappers exposés à l’UI */

// Import du fichier le plus récent dans le dossier Validation_Membres → upsert MEMBRES_GLOBAL
function runImportValidationMembres() {
  return importValidationMembresToGlobal_(getSeasonId_());
}

// Exports entraîneurs (membres + groupes)
function runExportEntraineursMembres() {
  return exportRetroEntraineursMembresXlsxToDrive(getSeasonId_(), {});
}
function runExportEntraineursGroupes() {
  return exportRetroEntraineursGroupesXlsxToDrive(getSeasonId_(), {});
}
function getMembreFlagsMetrics() {
  var ss = SpreadsheetApp.openById(getSeasonId_());
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

/** ============================== Outils debug =============================== */
function debugRetroFns() {
  if (!LIB) { Logger.log('LIB indisponible'); return; }
  Logger.log('typeof exportRetroMembresXlsxToDrive         = %s', typeof LIB.exportRetroMembresXlsxToDrive);
  Logger.log('typeof exportRetroGroupesAllXlsxToDrive      = %s', typeof LIB.exportRetroGroupesAllXlsxToDrive);
  Logger.log('typeof importerDonneesSaison                 = %s', typeof LIB.importerDonneesSaison);
  Logger.log('typeof evaluateSeasonRules                   = %s', typeof LIB.evaluateSeasonRules);
  Logger.log('typeof sendPendingOutbox                     = %s', typeof LIB.sendPendingOutbox);
}

function debugSeasonId() {
  var id = getSeasonId_();
  try {
    var f = DriveApp.getFileById(id);
    Logger.log('Drive OK: name=%s, mime=%s, url=%s, trashed=%s', f.getName(), f.getMimeType(), f.getUrl(), f.isTrashed());
  } catch (e) {
    Logger.log('DriveApp.getFileById FAILED: %s', e);
  }
  try {
    var ss = SpreadsheetApp.openById(id);
    Logger.log('Spreadsheet OK: title=%s, sheets=%s', ss.getName(), ss.getSheets().map(function(s){return s.getName();}).join(', '));
  } catch (e) {
    Logger.log('SpreadsheetApp.openById FAILED: %s', e);
  }
}

function debug_tailImportLog(n) {
  n = n || 30;
  var ss = SpreadsheetApp.openById(getSeasonId_());
  var sh = ss.getSheetByName('IMPORT_LOG');
  if (!sh || sh.getLastRow() < 2) { Logger.log('IMPORT_LOG vide.'); return; }
  var last = sh.getLastRow();
  var start = Math.max(2, last - n + 1);
  var vals = sh.getRange(start, 1, last - start + 1, Math.min(sh.getLastColumn(), 3)).getValues();
  vals.forEach(function(r){ Logger.log('%s | %s | %s', r[0], r[1], r[2]); });
}

/** ============================ Utilitaires UI ============================ */
/** Écrit une ligne dans IMPORT_LOG (Date | Type | Détails). Accepte 1 ou 3 arguments. */
function appendImportLog_(a, b, c){
  var ss = SpreadsheetApp.openById(getSeasonId_());
  var name = 'IMPORT_LOG';
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,3).setValues([['Date','Type','Détails']]);
  }

  var type, details;

  // Signature 1: appendImportLog_({ type, details })
  if (arguments.length === 1 && a && typeof a === 'object') {
    type = a.type || 'INFO';
    details = a.details || '';
  }
  // Signature 2: appendImportLog_(ssOrAnything, type, details) — on ignore le 1er
  else if (arguments.length >= 3) {
    type = b || 'INFO';
    details = c || '';
  }
  // Fallback
  else {
    type = 'INFO';
    details = String(a == null ? '' : a);
  }

  sh.appendRow([ new Date(), type, details ]);
}





/** KPI rapides pour l’accueil */
function getDashboardMetrics() {
  var ss = SpreadsheetApp.openById(getSeasonId_());

  function countRows_(name){
    var sh = ss.getSheetByName(name);
    return (sh && sh.getLastRow() > 1) ? (sh.getLastRow() - 1) : 0;
    }

  var m = {
    inscriptionsTotal: countRows_(SHEETS.INSCRIPTIONS),
    articlesTotal:     countRows_(SHEETS.ARTICLES),
    erreursTotal:      countRows_(SHEETS.ERREURS),
    outboxPending:     countRows_(SHEETS.MAIL_OUTBOX)
  };

  // Sous-titres par défaut (l’UI a des fallbacks)
  m.inscriptionsSubtitle = 'Total cumul.';
  m.articlesSubtitle     = 'Actifs';
  m.erreursSubtitle      = 'Total';
  m.outboxSubtitle       = 'Prêts à l’envoi';

  return m;
}

/** Écriture simple d’un param dans PARAMS (sans dépendre des helpers internes de la lib) */
function setParamValue(key, value) {
  var ss = SpreadsheetApp.openById(getSeasonId_());
  var sh = ss.getSheetByName('PARAMS') || ss.insertSheet('PARAMS');
  if (sh.getLastRow() < 1) sh.getRange(1,1,1,4).setValues([['Clé','Valeur','Type','Description']]);

  var last = sh.getLastRow();
  if (last < 2) { sh.appendRow([key, value, '', '']); return; }

  var keys = sh.getRange(2,1,last-1,1).getValues().map(function(r){return String(r[0]||'');});
  var row = -1;
  for (var i=0;i<keys.length;i++){ if (keys[i] === key) { row = 2+i; break; } }
  if (row === -1) sh.appendRow([key, value, '', '']);
  else sh.getRange(row,2).setValue(value);
}

/** Mini smoke test complet : import → règles → export Membres */
function smoke_test() {
  if (!LIB) throw new Error('Librairie indisponible');
  var id = getSeasonId_();
  Logger.log('---- RUN importerDonneesSaison ----');
  if (typeof LIB.importerDonneesSaison === 'function') Logger.log(LIB.importerDonneesSaison(id));
  Logger.log('---- RUN evaluateSeasonRules ----');
  if (typeof LIB.evaluateSeasonRules === 'function') Logger.log(JSON.stringify(LIB.evaluateSeasonRules(id)));
  Logger.log('---- RUN exportRetroMembresXlsxToDrive ----');
  Logger.log(JSON.stringify(LIB.exportRetroMembresXlsxToDrive(id)));
  Logger.log('---- TAIL IMPORT_LOG ----');
  debug_tailImportLog(40);
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


/** Liste unique des entraîneurs (rôles agrégés) + statuts photo/casier */
function getEntraineursAggreges() {
  var seasonId = getSeasonId_();
  var ss = SpreadsheetApp.openById(seasonId);

  var seasonYear = Number(readParam_(ss, 'SEASON_YEAR') || new Date().getFullYear());
  var invalidFrom = (readParam_(ss, 'PHOTO_INVALID_FROM_MMDD') || '04-01').trim();
  var cutoffNextJan1 = (seasonYear + 1) + '-01-01';
  var dueDate = seasonYear + '-' + invalidFrom;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var shR = ss.getSheetByName('ENTRAINEURS_ROLES');
  if (!shR || shR.getLastRow() < 2) return [];
  var R = shR.getDataRange().getValues();
  var h = R[0];
  var ciPass = h.indexOf('Passeport'),
      ciRole = h.indexOf('Role'),
      ciCat  = h.indexOf('Categorie'),
      ciEq   = h.indexOf('Equipe');

  var grouped = {}; // pass -> {Passeport, Roles[]}
  for (var i=1;i<R.length;i++){
    var p = normalizePassportPlain8_(R[i][ciPass]);
    if (!p) continue;
    var piece = [
      String(R[i][ciRole]||'').trim(),
      String(R[i][ciCat]||'').trim(),
      String(R[i][ciEq]||'').trim()
    ].filter(Boolean).join(' | ');
    (grouped[p] = grouped[p] || {Passeport:p, Roles:[]}).Roles.push(piece);
  }

  // jointure MEMBRES_GLOBAL
  var shMG = ss.getSheetByName(readParam_(ss,'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL');
  var V = shMG ? shMG.getDataRange().getValues() : [];
  var H = V[0] || [];
  var cPass = H.indexOf('Passeport'),
      cNom  = H.indexOf('Nom'),
      cPre  = H.indexOf('Prenom'),
      cDOB  = H.indexOf('DateNaissance'),
      cGen  = H.indexOf('Genre'),
      cPh   = H.indexOf('PhotoExpireLe'),
      cCas  = H.indexOf('CasierExpiré');

  var out = [];
  Object.keys(grouped).forEach(function(p){
    var nom='', prenom='', dob='', genre='', photo='', cas=0;

    // lookup linéaire (si volume ↑ on optimisera en map)
    for (var r=1; r<V.length; r++){
      if (normalizePassportPlain8_(V[r][cPass]) === p) {
        nom   = String(V[r][cNom]||''); 
        prenom= String(V[r][cPre]||'');
        dob   = String(V[r][cDOB]||'');
        genre = String(V[r][cGen]||'');
        photo = String(V[r][cPh]||'');
        cas   = Number(V[r][cCas]||0);
        break;
      }
    }

    var invalide = (photo && photo < cutoffNextJan1) ? 1 : 0;
    var statutPhoto = 'OK';
    if (invalide) {
      statutPhoto = (today >= dueDate) ? 'ÉCHUE' : ('À RENOUVELER dès ' + dueDate);
    }
    var statutCasier = (cas === 1) ? 'EXPIRÉ' : 'OK';

    out.push({
      Passeport: p,
      Prenom: prenom, Nom: nom, DateNaissance: dob, Genre: genre,
      PhotoExpireLe: photo, StatutPhoto: statutPhoto,
      Casier: statutCasier,
      Roles: grouped[p].Roles
    });
  });

  out.sort(function(a,b){
    var A=(a.Nom||'')+(a.Prenom||''), B=(b.Nom||'')+(b.Prenom||'');
    return A.localeCompare(B,'fr');
  });
  return out;
}
