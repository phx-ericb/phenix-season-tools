/* ======================= dashboard/src/server_rules.js ======================= */
/* FAST STRICT — aucun fallback legacy, logs explicites, zéro récursion. */

function _getSeasonSS_() {
  return getSeasonSpreadsheet_(getSeasonId_());
}

/* -------- Bridges pour écrire ERREURS (utilise LIB si dispo, sinon fallback) -------- */
function _rulesWriteFull_bridge(ss, errors, header) {
  try {
    if (typeof LIB !== 'undefined' && LIB && typeof LIB._rulesWriteFull_ === 'function') {
      return LIB._rulesWriteFull_(ss, errors, header);
    }
    if (typeof _rulesWriteFull_ === 'function') {
      return _rulesWriteFull_(ss, errors, header);
    }
  } catch (e) {
    // si l'impl de la lib jette, on tombera sur le fallback ci-dessous
  }

  // Fallback minimal : (ré)écrit entièrement ERREURS
  var sh = ss.getSheetByName('ERREURS') || ss.insertSheet('ERREURS');
  sh.clearContents();
  var H = Array.isArray(header) && header.length ? header
        : ['Passeport #','PS','Courriel','Type','Message','Saison','Frais','CreatedAt'];
  sh.getRange(1, 1, 1, H.length).setValues([H]);

  if (Array.isArray(errors) && errors.length) {
    sh.getRange(2, 1, errors.length, H.length).setValues(errors);
  }
}

function _rulesUpsertForPassports_bridge(ss, newErrors, touchedSet, header) {
  try {
    if (typeof LIB !== 'undefined' && LIB && typeof LIB._rulesUpsertForPassports_ === 'function') {
      return LIB._rulesUpsertForPassports_(ss, newErrors, touchedSet, header);
    }
    if (typeof _rulesUpsertForPassports_ === 'function') {
      return _rulesUpsertForPassports_(ss, newErrors, touchedSet, header);
    }
  } catch (e) { /* tombe sur fallback */ }

  // Fallback : supprime les lignes des passeports touchés puis append newErrors
  var sh = ss.getSheetByName('ERREURS') || ss.insertSheet('ERREURS');
  if (sh.getLastRow() < 1) {
    var H = Array.isArray(header) && header.length ? header
          : ['Passeport #','PS','Courriel','Type','Message','Saison','Frais','CreatedAt'];
    sh.getRange(1,1,1,H.length).setValues([H]);
  }
  var Hdr = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0] || [];
  var cP = Hdr.indexOf('Passeport #'); if (cP < 0) cP = Hdr.indexOf('Passeport');

  if (cP >= 0 && touchedSet && touchedSet.size && sh.getLastRow() > 1) {
    var rng = sh.getRange(2,1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
    var rowsToDel = [];
    for (var i=0;i<rng.length;i++) {
      var p = String(rng[i][cP]||'').replace(/\D/g,'').padStart(8,'0');
      if (touchedSet.has(p)) rowsToDel.push(2+i);
    }
    rowsToDel.sort(function(a,b){ return b-a; }).forEach(function(r){ sh.deleteRow(r); });
  }

  if (Array.isArray(newErrors) && newErrors.length) {
    sh.insertRowsAfter(sh.getLastRow(), newErrors.length);
    sh.getRange(sh.getLastRow()-newErrors.length+1, 1, newErrors.length, (Hdr.length||newErrors[0].length))
      .setValues(newErrors);
  }
}

/* --------- Sélection stricte du builder FAST (+ logging route) --------- */

function _pickFastBuilder_(ss) {
  // Priorité: LIB._rulesBuildErrorsFast_
  if (typeof LIB !== 'undefined' && LIB && typeof LIB._rulesBuildErrorsFast_ === 'function') {
    appendImportLog_(ss, 'RULES_FAST_ROUTE', { route: 'LIB._rulesBuildErrorsFast_' });
    return {
      build: function(touchedSet) { return LIB._rulesBuildErrorsFast_(ss, touchedSet || null); },
      writeFull: function(errs, header) { return LIB._rulesWriteFull_(ss, errs, header); },
      upsert: function(errs, touchedSet, header) { return LIB._rulesUpsertForPassports_(ss, errs, touchedSet, header); }
    };
  }

  // Repli strict: global _rulesBuildErrorsFast_ (défini dans library-si/src/source.js)
  if (typeof _rulesBuildErrorsFast_ === 'function') {
    appendImportLog_(ss, 'RULES_FAST_ROUTE', { route: '_rulesBuildErrorsFast_' });
    return {
      build: function(touchedSet) { return _rulesBuildErrorsFast_(ss, touchedSet || null); },
      writeFull: function(errs, header) { return _rulesWriteFull_(ss, errs, header); },
      upsert: function(errs, touchedSet, header) { return _rulesUpsertForPassports_(ss, errs, touchedSet, header); }
    };
  }

  appendImportLog_(ss, 'RULES_FAST_ROUTE_FAIL', { reason: 'no-fast-builder' });
  throw new Error('FAST rules builder introuvable (LIB._rulesBuildErrorsFast_ / _rulesBuildErrorsFast_)');
}

/* ----------------------------- FULL (FAST STRICT) ---------------------------- */

function runEvaluateRules() {
  var ss = getSeasonSpreadsheet_(getSeasonId_());

  if (LIB && typeof LIB.runEvaluateRulesFast_ === 'function') {
    appendImportLog_(ss, 'RULES_FAST_ROUTE', JSON.stringify({ path: 'LIB.runEvaluateRulesFast_' }));
    return LIB.runEvaluateRulesFast_(ss);
  }

  if (typeof _rulesBuildErrorsFast_ === 'function' && typeof _rulesWriteFull_ === 'function') {
    appendImportLog_(ss, 'RULES_FAST_ROUTE', JSON.stringify({ path: '_rulesBuildErrorsFast_ + _rulesWriteFull_' }));
    var res = _rulesBuildErrorsFast_(ss, null);
    return _rulesWriteFull_(ss, res.errors, res.header);
  }

  appendImportLog_(ss, 'RULES_FAST_ROUTE_FAIL', JSON.stringify({ reason: 'no-fast-builder' }));
  throw new Error('FAST rules builder introuvable (LIB.runEvaluateRulesFast_ / _rulesBuildErrorsFast_)');
}

function runEvaluateRulesFast_() {
  var ss = getSeasonSpreadsheet_(getSeasonId_());

  if (typeof LIB !== 'undefined' && LIB && typeof LIB._rulesBuildErrorsFast_ === 'function') {
    appendImportLog_(ss, 'RULES_FAST_ROUTE', { path: 'LIB._rulesBuildErrorsFast_' });
    var res = LIB._rulesBuildErrorsFast_(ss); // FULL => touchedSet=null
    _rulesWriteFull_(ss, res.errors, res.header);
    appendImportLog_(ss, 'RULES_DONE', { written: res.errors.length, ledRows: res.ledgerCount, joueursRows: res.joueursCount });
    return { written: res.errors.length };
  }

  appendImportLog_(ss, 'RULES_FAST_ROUTE_FAIL', { reason: 'no-fast-builder' });
  throw new Error('FAST rules builder introuvable (LIB._rulesBuildErrorsFast_)');
}

/* ---------------------------- INCR (FAST STRICT) ----------------------------- */

function evaluateSeasonRulesIncr(passports, ss) {
  ss = ss || getSeasonSpreadsheet_(getSeasonId_());
  var set = (typeof _toPassportSet_ === 'function') ? _toPassportSet_(passports) : new Set((passports||[]).map(String));

  if (LIB && typeof LIB._rulesBuildErrorsIncrFast_ === 'function') {
    appendImportLog_(ss, 'RULES_FAST_ROUTE', JSON.stringify({ path: 'LIB._rulesBuildErrorsIncrFast_' }));
    var res = LIB._rulesBuildErrorsIncrFast_(passports, ss);
    var up = (LIB && typeof LIB._rulesUpsertForPassports_ === 'function') ? LIB._rulesUpsertForPassports_ : _rulesUpsertForPassports_;
    if (typeof up !== 'function') throw new Error('_rulesUpsertForPassports_ manquant');
    return up(ss, res.errors, set, res.header);
  }

  if (typeof _rulesBuildErrorsFast_ === 'function' && typeof _rulesUpsertForPassports_ === 'function') {
    appendImportLog_(ss, 'RULES_FAST_ROUTE', JSON.stringify({ path: '_rulesBuildErrorsFast_ + _rulesUpsertForPassports_' }));
    var res2 = _rulesBuildErrorsFast_(ss, set);
    return _rulesUpsertForPassports_(ss, res2.errors, set, res2.header);
  }

  appendImportLog_(ss, 'RULES_FAST_ROUTE_FAIL', JSON.stringify({ reason: 'no-fast-builder' }));
  throw new Error('FAST INCR rules introuvable (LIB._rulesBuildErrorsIncrFast_ / _rulesBuildErrorsFast_)');
}

function evaluateSeasonRulesIncr(passports, ss) {
  ss = ss || getSeasonSpreadsheet_(getSeasonId_());

  if (typeof LIB !== 'undefined' && LIB && typeof LIB._rulesBuildErrorsIncrFast === 'function') {
    appendImportLog_(ss, 'RULES_FAST_ROUTE', { path: 'LIB._rulesBuildErrorsIncrFast' });
    var res = LIB._rulesBuildErrorsIncrFast(passports, ss);
    _rulesUpsertForPassports_(ss, res.errors, _toPassportSet_(passports), res.header);
    appendImportLog_(ss, 'RULES_INCR_DONE', { touched: passports.length, written: res.errors.length });
    return { written: res.errors.length };
  }

  appendImportLog_(ss, 'RULES_FAST_ROUTE_FAIL', { reason: 'no-fast-builder-incr' });
  throw new Error('FAST INCR builder introuvable (LIB._rulesBuildErrorsIncrFast)');
}


/* ----------------------------- utilitaire public ----------------------------- */

function runEvaluateRulesIncrFromLastTouched(ssOrId) {
  var ss = ssOrId || _getSeasonSS_();
  var raw = PropertiesService.getDocumentProperties().getProperty('LAST_TOUCHED_PASSPORTS') || '[]';
  var list = (raw[0] === '[' ? JSON.parse(raw) : raw.split(',')).map(function(x){ return String(x||'').trim(); }).filter(Boolean);
  return evaluateSeasonRulesIncrFast_(list, ss);
}

/* ======================= FIN server_rules.js (FAST STRICT) =================== */
