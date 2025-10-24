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

/** Écriture simple d’un param dans PARAMS (sans dépendre des helpers internes de la lib) */
function setParamValue(key, value) {
  var ss = getSeasonSpreadsheet_(getSeasonId_());
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

/** ======================== SEASON CONTEXT (HARDENED) ======================== */

var __SEASON_CTX__ = __SEASON_CTX__ || { id: null, ss: null, ts: 0 };

function setSeasonId_(seasonId) {
  seasonId = String(seasonId || '').trim();
  if (!seasonId) throw new Error('setSeasonId_: seasonId vide');
  __SEASON_CTX__.id = seasonId;
  __SEASON_CTX__.ts = Date.now();
  try { CacheService.getScriptCache().put('SEASON_ID', seasonId, 300); } catch (_){}
  try { PropertiesService.getScriptProperties().setProperty('SEASON_ID', seasonId); } catch (_){}
  try { PropertiesService.getScriptProperties().setProperty('ACTIVE_SEASON_ID', seasonId); } catch (_){}
  return seasonId;
}

/** IMPORTANT: ne touche PAS aux feuilles ici. Jamais.
 *  Résolution prioritaire harmonisée (constante → ACTIVE_SEASON_ID → PHENIX_SEASON_SHEET_ID → SEASON_ID legacy)
 */
function getSeasonId_() {
  // 1) Contexte mémoire
  if (__SEASON_CTX__.id) return __SEASON_CTX__.id;

  var props = PropertiesService.getScriptProperties();
  var id = null;

  // 2) Constante (si fournie)
  if (!id && SEASON_SHEET_ID && String(SEASON_SHEET_ID).trim()) id = String(SEASON_SHEET_ID).trim();

  // 3) ACTIVE_SEASON_ID (pilotée par l’UI)
  if (!id) try { id = String(props.getProperty('ACTIVE_SEASON_ID') || '').trim(); } catch (_){}

  // 4) PHENIX_SEASON_SHEET_ID (setSeasonSheetIdOnce)
  if (!id) try { id = String(props.getProperty('PHENIX_SEASON_SHEET_ID') || '').trim(); } catch (_){}

  // 5) SEASON_ID (legacy/cache)
  if (!id) {
    try { id = String(CacheService.getScriptCache().get('SEASON_ID') || '').trim(); } catch (_){}
    if (!id) try { id = String(props.getProperty('SEASON_ID') || '').trim(); } catch (_){}
  }

  if (!id) {
    throw new Error(
      "Aucun ID de classeur saison. Définis SEASON_SHEET_ID " +
      "ou exécute setSeasonSheetIdOnce() / clique 'Définir active' dans l'UI."
    );
  }

  __SEASON_CTX__.id = id;
  __SEASON_CTX__.ts = Date.now();
  return id;
}

/** Open with exponential backoff + léger lock pour éviter l’overlap */
function openSpreadsheetWithBackoff_(id, opt) {
  opt = opt || {};
  var maxAttempts = Math.max(3, opt.maxAttempts || 8);
  var base = Math.max(200, opt.baseSleepMs || 600);
  var lastErr = null;

  // Lock court: sérialise l’ouverture entre threads
  var lock = null;
  try {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(2000)) Utilities.sleep(250);
  } catch (_){}

  try {
    for (var a = 1; a <= maxAttempts; a++) {
      try {
        // petit "poke" Drive pour éviter un cold-start
        try { DriveApp.getFileById(id).getId(); } catch (_){}
        var ss = SpreadsheetApp.openById(id);
        // réchauffe une lecture non-mutante
        try { var _ = ss.getSheets()[0].getMaxRows(); } catch(_){}
        return ss;
      } catch (e) {
        lastErr = e;
        var msg = String(e && e.message || e);
        if (/Permission|not found|Invalid/i.test(msg)) throw e; // non retryable
        var wait = Math.min(5000, base * Math.pow(1.6, a - 1));
        Utilities.sleep(wait + Math.floor(Math.random()*200)); // jitter
      }
    }
    throw new Error('openSpreadsheetWithBackoff_: échec après ' + maxAttempts + ' essais — ' + String(lastErr));
  } finally {
    try { lock && lock.releaseLock(); } catch(_){}
  }
}

/** Toujours utiliser ce point d’entrée unique pour obtenir le classeur saison */
function getSeasonSpreadsheet_(seasonIdOpt) {
  var wanted = String(seasonIdOpt || getSeasonId_() || '').trim();
  if (!wanted) throw new Error('getSeasonSpreadsheet_: seasonId vide');

  var ctx = __SEASON_CTX__;
  if (ctx.ss && ctx.id === wanted) return ctx.ss;

  var ss = openSpreadsheetWithBackoff_(wanted, { maxAttempts: 8, baseSleepMs: 600 });
  ctx.id = wanted;
  ctx.ss = ss;
  ctx.ts = Date.now();

  // Persistance légère
  try { CacheService.getScriptCache().put('SEASON_ID', wanted, 300); } catch (_){}
  try { PropertiesService.getScriptProperties().setProperty('SEASON_ID', wanted); } catch (_){}
  try { PropertiesService.getScriptProperties().setProperty('ACTIVE_SEASON_ID', wanted); } catch (_){}

  return ss;
}

function ensureSeasonContext_() {
  var id = getSeasonId_();
  var ss = getSeasonSpreadsheet_(id);
  return { id: id, ss: ss };
}

/** ======================== PARAMS helpers ======================== */

/** Lecture simple d’un param dans PARAMS (utilisée par runImportAndExports pour DRY_RUN) */
function readParamValue(key) {
  var ss = getSeasonSpreadsheet_(getSeasonId_());
  return readParam_(ss, key);
}

/** Lit un paramètre depuis PARAMS (fallback DocumentProperties) */
function readParam_(ss, key) {
  ss = ensureSpreadsheet_(ss);
  var sh = ss.getSheetByName('PARAMS');
  if (!sh || sh.getLastRow() < 2) {
    // fallback éventuel : Document Properties
    var v = PropertiesService.getDocumentProperties().getProperty(String(key));
    return v ? String(v) : '';
  }
  var vals = sh.getRange(2,1, sh.getLastRow()-1, 2).getValues(); // Col A=Clé, B=Valeur
  for (var i=0;i<vals.length;i++) {
    if (String(vals[i][0]||'') === key) return String(vals[i][1]||'');
  }
  return '';
}

/** Garantit qu’on a un Spreadsheet */
function ensureSpreadsheet_(ss) {
  if (ss && typeof ss.getId === 'function') return ss;
  return getSeasonSpreadsheet_(getSeasonId_());
}

// Utils communs
function _ok(data, message) { return { ok: true, data: data || null, message: message || '' }; }
function _err(e) { var msg = (e && e.stack) ? String(e.stack) : String(e); return { ok: false, error: msg }; }
function _registry_() { return PropertiesService.getScriptProperties(); }

function _wrap(action, fn, opts) {
  opts = opts || {};              // { lock: false|true }
  var lock = null;
  if (opts.lock) {
    lock = LockService.getScriptLock();
    lock.waitLock(20 * 1000);
  }
  try {
    var out = fn();
    _auditLog_(action, { ok: true });
    return out;
  } catch (e) {
    _auditLog_(action, { ok: false, error: String(e) });
    return _err(e);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function _auditLog_(action, details) {
  try {
    // Option A : si tu veux un journal dans un classeur dédié, stocke son ID dans Script Properties
    // (clé DASHBOARD_AUDIT_SHEET_ID). Sinon on logge juste en console.
    var props = PropertiesService.getScriptProperties();
    var auditId = props.getProperty('DASHBOARD_AUDIT_SHEET_ID');

    if (!auditId) {
      // Pas de classeur d’audit configuré → log console et on sort
      console.log('AUDIT', action, JSON.stringify(details || {}));
      return;
    }

    var ss = SpreadsheetApp.openById(auditId);
    var sh = ss.getSheetByName('DASHBOARD_AUDIT') || ss.insertSheet('DASHBOARD_AUDIT');
    if (sh.getLastRow() === 0) sh.appendRow(['When','User','Action','Details']);
    sh.appendRow([ new Date(), Session.getActiveUser().getEmail() || 'anon', action, JSON.stringify(details || {}) ]);
  } catch (e) {
    // Surtout ne jamais faire planter l’appel serveur à cause de l’audit
    console.log('AUDIT_FAIL', action, String(e));
  }
}

/** Feuille utilitaire */
function getSheetOrCreate_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}

// PARAM schema partagé (utilisé par server_params)
var PARAM_SCHEMA = {
  DRIVE_FOLDER_IMPORTS:'FOLDER_ID', FILE_PATTERN_INSCRIPTIONS:'STRING', FILE_PATTERN_ARTICLES:'STRING',
  DRY_RUN:'BOOLEAN', MOVE_CONVERTED_TO_ARCHIVE:'BOOLEAN', INCREMENTAL_ON:'BOOLEAN', KEY_COLS:'STRING', SEASON_LABEL:'STRING',
  ANNULATIONS_ACTIVE:'BOOLEAN', STATUS_COL_INSCRIPTIONS:'STRING', STATUS_COL_ARTICLES:'STRING', STATUS_CANCEL_VALUES:'STRING', STATUS_CANCEL_DATE_COL:'STRING',
  MAIL_FROM:'STRING', MAIL_BATCH_MAX:'NUMBER', MAIL_TO_NEW_INSCRIPTIONS:'STRING', MAIL_CC_NEW_INSCRIPTIONS:'STRING', TO_FIELDS_INSCRIPTIONS:'STRING',
  MAIL_TEMPLATE_INSCRIPTION_NEW_SUBJECT:'STRING', MAIL_TEMPLATE_INSCRIPTION_NEW_BODY:'STRING',
  MAIL_TO_SUMMARY_U4U8:'STRING', MAIL_CC_SUMMARY_U4U8:'STRING', MAIL_TO_SUMMARY_U9U12:'STRING', MAIL_CC_SUMMARY_U9U12:'STRING',
  MAIL_TO_SUMMARY_U13U18:'STRING', MAIL_CC_SUMMARY_U13U18:'STRING', MAIL_TEMPLATE_SUMMARY_SUBJECT:'STRING', MAIL_TEMPLATE_SUMMARY_BODY:'STRING',
  RULES_ON:'BOOLEAN', RULES_SEVERITY_THRESHOLD:'STRING', RULES_APPEND:'BOOLEAN', RULES_DRY_RUN:'BOOLEAN', RULES_MAX_ROWS:'NUMBER',
  EXPORT_MODIFIED_SINCE_DAYS:'NUMBER', RETRO_IGNORE_FEES_CSV:'STRING', RETRO_ADAPTE_KEYWORDS:'STRING', RETRO_CAMP_KEYWORDS:'STRING',
  RETRO_MUTATION_SHEET:'STRING', RETRO_EXPORTS_FOLDER_ID:'FOLDER_ID',
  RETRO_PHOTO_INCLUDE_COL:'BOOLEAN', RETRO_PHOTO_EXPIRY_COL:'STRING', RETRO_PHOTO_WARN_ABS_DATE:'STRING',
  RETRO_PHOTO_WARN_BEFORE_MMDD:'STRING', RETRO_EXCLUDE_MEMBER_IF_ONLY_IGN:'BOOLEAN',
  RETRO_GROUP_SHEET_NAME:'STRING', RETRO_GROUP_EXPORTS_FOLDER_ID:'FOLDER_ID', RETRO_GROUP_ELITE_KEYWORDS:'STRING', RETRO_GROUP_SA_KEYWORDS:'STRING',
  RETRO_GROUP_SA_GROUPE_LABEL:'STRING', RETRO_GROUP_SA_CATEG_LABEL:'STRING',
  RETRO_GROUP_GROUPE_FMT:'STRING', RETRO_GROUP_CATEGORIE_FMT:'STRING',
  RETRO_GART_SHEET_NAME:'STRING', RETRO_GART_EXPORTS_FOLDER_ID:'FOLDER_ID', RETRO_GART_IGNORE_FEES_CSV:'STRING',
  RETRO_GART_ELITE_KEYWORDS:'STRING', RETRO_GART_REQUIRE_MAPPING:'BOOLEAN',
  SEASON_YEAR:'NUMBER',            // ex. 2025 (prioritaire sur le parsing de SEASON_LABEL)
  RETRO_MEMBER_MAX_U:'NUMBER',     // optionnel: ex. 18 → >=19 traité comme adulte pour la validation
};
function _coerceByType_(key, val) {
  var t = PARAM_SCHEMA[key] || 'STRING';
  if (val === null || val === undefined) return '';
  if (t === 'BOOLEAN') return (String(val).toLowerCase() === 'true' || val === true);
  if (t === 'NUMBER') return Number(val);
  return String(val);
}
function _upsertParam_(ss, key, value, type, desc) {
  var sh = ss.getSheetByName('PARAMS') || ss.insertSheet('PARAMS');
  if (sh.getLastRow() === 0) sh.appendRow(['Clé','Valeur','Type','Description']);
  var v = sh.getDataRange().getValues();
  for (var i=1; i<v.length; i++) {
    if (v[i][0] === key) {
      sh.getRange(i+1, 2).setValue(value);
      if (type) sh.getRange(i+1, 3).setValue(type);
      if (desc) sh.getRange(i+1, 4).setValue(desc);
      return;
    }
  }
  sh.appendRow([key, value, type||'', desc||'']);
}

/** Divers utilitaires */
function getSheet_(ss, name, createIfMissing) {
  return createIfMissing ? getSheetOrCreate_(ss, name) : ss.getSheetByName(name);
}
function norm_(s){
  return String(s||'').trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
function isIgnoredFeeRetro_(ss, fee){
  var v = norm_(fee);
  if (!v) return false;
  var csv = (typeof readParam_==='function' ? (readParam_(ss,'RETRO_IGNORE_FEES_CSV')||'') : '');
  var toks = csv.split(',').map(norm_).filter(Boolean);

  if (toks.indexOf(v) >= 0) return true;        // exact
  for (var i=0;i<toks.length;i++)               // contains
    if (v.indexOf(toks[i]) >= 0) return true;

  // filet si le param n'était pas encore rempli
  if (/(entraineur|entra[îi]neur|coach)/i.test(String(fee||''))) return true;
  return false;
}

// Backward-compat: si certains appels utilisent encore __openByIdWithRetry__
if (typeof __openByIdWithRetry__ !== 'function') {
  function __openByIdWithRetry__(id, attempts) {
    // si on me demande la saison active, renvoie le handle existant
    try {
      var sid = getSeasonId_();
      if (sid && String(id) === String(sid)) return getSeasonSpreadsheet_(sid);
    } catch (_){}
    // sinon open direct (ou via openSpreadsheetWithBackoff_ si dispo)
    return (typeof openSpreadsheetWithBackoff_ === 'function')
      ? openSpreadsheetWithBackoff_(id, { maxAttempts: Math.max(3, attempts || 5), baseSleepMs: 600 })
      : SpreadsheetApp.openById(id);
  }
}
