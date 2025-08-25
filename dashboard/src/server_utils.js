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


function getSeasonSpreadsheet_(seasonId) { return SpreadsheetApp.openById(seasonId); }

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
  RETRO_GROUP_SA_GROUPE_LABEL:'STRING', RETRO_GROUP_SA_CATEG_LABEL:'STRING', RETRO_GROUP_GROUPE_FMT:'STRING', RETRO_GROUP_CATEGORIE_FMT:'STRING',
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
// Helpers PARAMS
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
function _inferSeasonLabelFromTitle_(title) {
  var m = title.match(/^(.+?)(?:\s*-\s*Saison)?$/i);
  return (m ? m[1] : title).trim();
}
