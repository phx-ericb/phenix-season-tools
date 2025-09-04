// --- Helpers "ignore frais (RÉTRO)" ---
// normalise (trim, minuscule, sans accents)
function SR_norm_(s) {
  return String(s || '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

// Vrai si le "Nom du frais" doit être ignoré selon RETRO_IGNORE_FEES_CSV (+ RETRO_COACH_FEES_CSV)
function SR_isIgnoredFeeRetro_(ss, fee) {
  var v = SR_norm_(fee);
  if (!v) return false;

  var baseCsv  = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV) || '';
  var coachCsv = readParam_(ss, 'RETRO_COACH_FEES_CSV') || '';
  var toks = (baseCsv + (baseCsv&&coachCsv?',':'') + coachCsv)
    .split(',').map(SR_norm_).filter(Boolean);

  if (toks.indexOf(v) >= 0) return true;            // exact
  for (var i = 0; i < toks.length; i++)             // contains
    if (v.indexOf(toks[i]) >= 0) return true;

  // filet lexical si paramétrage incomplet
  if (/(entraineur|entra[îi]neur|coach)/i.test(String(fee || ''))) return true;

  return false;
}


function getRules(seasonId) {
  return _wrap('getRules', function(){
    var ss = getSeasonSpreadsheet_(seasonId);
    var sh = ss.getSheetByName('RETRO_RULES_JSON');
    var text = sh && sh.getLastRow()>0 ? String(sh.getRange(1,1).getValue()||'') : '';
    if (!text) {
      var p=ss.getSheetByName('PARAMS'); if (p) { var values=p.getDataRange().getValues();
        for (var i=1;i<values.length;i++) if (values[i][0]==='RETRO_RULES_JSON') { text=String(values[i][1]||''); break; }
      }
    }
    var parsed=null, ok=true, err=''; if (text && String(text).trim()) { try { parsed=JSON.parse(text);} catch(e){ ok=false; err=String(e);} }
    return _ok({ jsonText:text, parsed:parsed, parsedOk:ok, error:err });
  });
}
function setRules(seasonId, jsonText) {
  return _wrap('setRules', function(){
    try { if (jsonText && jsonText.trim()) JSON.parse(jsonText); } catch(e) { throw new Error('Invalid JSON: ' + e); }
    var sh = getSeasonSpreadsheet_(seasonId).getSheetByName('RETRO_RULES_JSON') || getSeasonSpreadsheet_(seasonId).insertSheet('RETRO_RULES_JSON');
    sh.clear(); sh.getRange(1,1).setValue(jsonText);
    return _ok(null,'Rules saved');
  });
}
// ---------- server_rules.js ----------
// Cache global partagé entre tous les fichiers (5 minutes)
var __retroRulesCache = (typeof __retroRulesCache !== 'undefined')
  ? __retroRulesCache
  : { at: 0, data: null };

/**
 * Charge les règles rétro depuis PARAM ou la feuille 'RETRO_RULES_JSON'.
 * Renvoie toujours un Array (éventuellement vide). Cache 5 min.
 */
function SR_loadRetroRules_(ss) {
  var now = Date.now();
  if (__retroRulesCache.data && (now - __retroRulesCache.at) < 5 * 60 * 1000) {
    return __retroRulesCache.data;
  }

  // 1) PARAM direct
  var raw = readParam_(ss, PARAM_KEYS.RETRO_RULES_JSON) || '';

  // 2) Feuille "RETRO_RULES_JSON" si vide
  if (!raw) {
    var shJson = ss.getSheetByName('RETRO_RULES_JSON');
    if (shJson && shJson.getLastRow() >= 1 && shJson.getLastColumn() >= 1) {
      var vals = shJson.getDataRange().getDisplayValues();
      var pieces = [];
      for (var i = 0; i < vals.length; i++) {
        for (var j = 0; j < vals[i].length; j++) {
          var cell = vals[i][j];
          if (cell != null && String(cell).trim() !== '') pieces.push(String(cell));
        }
      }
      raw = pieces.join('\n');
      appendImportLog_(ss, 'RETRO_RULES_JSON_SHEET_READ', 'chars=' + raw.length);
    }
  }

  // 3) Parse JSON si présent
  if (raw) {
    try {
      var arr = JSON.parse(raw);
      var rulesFromJson = Array.isArray(arr) ? arr : [];
      __retroRulesCache = { at: now, data: rulesFromJson };
      return rulesFromJson;
    } catch (e) {
      appendImportLog_(ss, 'RETRO_RULES_JSON_PARSE_FAIL', String(e));
    }
  }

  // 4) Fallback: règles dérivées des PARAMS (par défauts raisonnables)
  var ignoreCsv = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV) || 'senior,u-sé,adulte,ligue';
  var adapteCsv = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || 'adapté,adapte';
  var campCsv   = readParam_(ss, PARAM_KEYS.RETRO_CAMP_KEYWORDS)   || 'camp de sélection u13,camp selection u13,camp u13';
  var photoOn   = (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) || 'FALSE').toUpperCase() === 'TRUE';
  var photoCol  = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL) || '';
  var warnMmDd  = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD) || '03-01';
  var absDate   = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';

  var rules = [
    { id:'ignore_fees', enabled:true, scope:'both',
      when:{ field:'Nom du frais', contains_any: ignoreCsv.split(',') },
      action:{ type:'ignore_row' } },
    { id:'adapte_flag', enabled:true, scope:'both',
      when:{ field:'Nom du frais', contains_any: adapteCsv.split(',') },
      action:{ type:'set_member_field', field:'adapte', value:1 } },
    // CDP via catalogue
    { id:'cdp_2', enabled:true, scope:'articles',
      when:{ catalog_exclusive_group:'CDP_ENTRAINEMENT', text_contains_any:['2','2 entrainements'] },
      action:{ type:'set_member_field_max', field:'cdp', value:2 } },
    { id:'cdp_1', enabled:true, scope:'articles',
      when:{ catalog_exclusive_group:'CDP_ENTRAINEMENT' },
      action:{ type:'set_member_field_max', field:'cdp', value:1 } },
    { id:'camp_u13', enabled:true, scope:'articles',
      when:{ field:'Nom du frais', contains_any: campCsv.split(',') },
      action:{ type:'set_member_field', field:'camp', value:'Oui' } }
  ];
  if (photoOn && photoCol) {
    rules.push({ id:'photo_policy', enabled:true, scope:'member',
      action:{ type:'compute_photo', expiry_col:photoCol, warn_mmdd:warnMmDd, abs_date:absDate }});
  }
  appendImportLog_(ss, 'RETRO_RULES_JSON_FALLBACK', 'using PARAMS-derived defaults');

  __retroRulesCache = { at: now, data: rules };
  return rules;
}
