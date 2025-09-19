/* ============================== server_flow.js =============================== */

function __ping_aug__() {
  var seasonId = getSeasonId_();
  var ok = (typeof runPostImportAugmentations_ === 'function');
  if (!ok) throw new Error('Hook not visible');
  // Toujours (re)construire les pivots, même en dry-run (safe: écrit dans le classeur)
  runPostImportAugmentations_(seasonId, [], { isFull: true, isDryRun: false });
}

/* ----------------------------- INCR PIPELINE -------------------------------- */

function runImportRulesExportsIncr() {
  var seasonId = (typeof getSeasonId_ === 'function') ? getSeasonId_() : null;
  var ss = getSeasonSpreadsheet_(seasonId);
  var ctx = (typeof startImportRun_ === 'function')
    ? startImportRun_({ seasonId: seasonId, source: 'dashboard-incr' })
    : { runId: Utilities.getUuid() };

  try {
    // (0) Import fichiers (scan/convert → STAGING → INSCRIPTIONS/ARTICLES)
    var imp = (typeof runImporterDonneesSaison === 'function')
      ? runImporterDonneesSaison(seasonId)  // via LIB si dispo
      : null;

    // (1) Passeports "touchés" → uniques + format 8 chiffres
    var touchedRaw = (typeof _getTouchedPassportsArray_ === 'function')
      ? _getTouchedPassportsArray_(ss, seasonId)
      : [];
    var touched = Array.from(new Set(touchedRaw || []))
      .map(function (p) { return String(p || '').replace(/\D/g, '').padStart(8, '0'); })
      .filter(Boolean);

    // (2) AUGMENTATIONS (toujours actives) : build ACHATS_LEDGER + JOUEURS
    try {
      var augOpts = { isFull: false, isDryRun: false }; // on force l’écriture des pivots
      if (touched.length) {
        // NOTE: l’API attend un Set pour cibler les passeports
        runPostImportAugmentations_(seasonId, new Set(touched), augOpts);
        appendImportLog_(ss, 'AUG_INCR_PRE_OK', JSON.stringify({ via: 'SHIM', count: touched.length }));
      } else {
        appendImportLog_(ss, 'AUG_INCR_PRE_SKIP', JSON.stringify({ reason: 'no-passports' }));
      }
    } catch (eAugPre) {
      appendImportLog_(ss, 'AUG_INCR_PRE_FAIL', String(eAugPre));
    }

    // (3) RÈGLES (INCR) → le moteur FAST clear/écrit ERREURS lui-même
    var rRules = null;
    try {
      if (touched.length && typeof evaluateSeasonRulesIncr === 'function') {
        rRules = evaluateSeasonRulesIncr(touched, ss);
        appendImportLog_(ss, 'RULES_INCR_DONE', rRules);
      } else {
        appendImportLog_(ss, 'RULES_INCR_SKIP', { reason: touched.length ? 'engine-missing' : 'no-passports' });
      }
    } catch (eRules) {
      appendImportLog_(ss, 'RULES_INCR_FAIL', String(eRules));
    }

    // (4) EXPORTS INCR (si tu as un exporteur incrémental)
    try {
      var combined = _boolParam_ ? _boolParam_('RETRO_COMBINED_XLSX', false) : false;
      if (typeof runRetroExportsIncr === 'function') {
        var rExp = runRetroExportsIncr(touched, { combined: combined });
        appendImportLog_(ss, 'EXPORTS_INCR_OK', rExp);
      } else {
        appendImportLog_(ss, 'EXPORTS_INCR_SKIP', 'no-exporter');
      }
    } catch (eExp) {
      appendImportLog_(ss, 'EXPORTS_INCR_FAIL', String(eExp));
    }

    // (5) MAILS (AFTER) — DRY_RUN géré par le worker
    try {
      var stage = String(readParam_(ss, 'MAIL_STAGE') || 'AFTER').trim().toUpperCase();
      if (stage === 'AFTER' && typeof runMailPipelineSelected_ === 'function') {
        runMailPipelineSelected_(ss, 'AFTER');
      }
    } catch (e) {
      appendImportLog_(ss, 'MAIL_PIPELINE_FAIL', String(e));
    }

    return { ok: true, import: imp, rules: rRules };
  } catch (e) {
    appendImportLog_(ss, 'FLOW_INCR_FAIL', String(e));
    return { ok: false, error: String(e) };
  } finally {
    if (typeof endImportRun_ === 'function') endImportRun_(ctx);
  }
}


/* ----------------------------- FULL PIPELINE -------------------------------- */

function runImportRulesExportsFull() {
  var seasonId = getSeasonId_();
  var ss = getSeasonSpreadsheet_(seasonId);
  var ctx = startImportRun_({ seasonId: seasonId, source: 'dashboard-full' });

  try {
    // 0) Import fichiers (scan/convert → STAGING → INSCRIPTIONS/ARTICLES)
    var imp = runImporterDonneesSaison();

    // 1) AUGMENTATIONS (FULL) — rebuild pivots ACHATS_LEDGER + JOUEURS
    try {
      var augOpts = { isFull: true, isDryRun: false }; // toujours écrire les pivots
      runPostImportAugmentations_(seasonId, [], augOpts);
      appendImportLog_(ss, 'AUG_FULL_PRE_OK', JSON.stringify({ via: 'SHIM' }));
    } catch (eAugPre) {
      appendImportLog_(ss, 'AUG_FULL_PRE_FAIL', String(eAugPre));
    }

    // 2) RÈGLES (FULL). Le moteur FAST clear/écrit ERREURS lui-même
    var rFull;
    try {
      if (typeof runEvaluateRules === 'function') {
        rFull = runEvaluateRules(); // alias vers FAST si redirigé
      } else if (typeof runEvaluateRulesFast_ === 'function') {
        rFull = runEvaluateRulesFast_();
      } else {
        rFull = { ok: false, error: 'No rules engine available' };
      }
      appendImportLog_(ss, 'RULES_FULL_DONE', rFull);
    } catch (eRules) {
      if (typeof runEvaluateRulesFast_ === 'function') {
        appendImportLog_(ss, 'RULES_FALLBACK_FAST', String(eRules));
        rFull = runEvaluateRulesFast_();
        appendImportLog_(ss, 'RULES_FULL_DONE', rFull);
      } else {
        appendImportLog_(ss, 'RULES_FULL_FAIL', String(eRules));
        throw eRules;
      }
    }

    // 3) EXPORTS (FULL garantis)
    var mRes = null, gRes = null;
    try {
      mRes = runExportRetroMembres();   // lit JOUEURS si RETRO_MEMBRES_READ_SOURCE=JOUEURS
      appendImportLog_(ss, 'EXPORT_MEMBRES_OK', mRes);
    } catch (eM) {
      appendImportLog_(ss, 'EXPORT_MEMBRES_FAIL', String(eM));
    }
    try {
      gRes = runExportRetroGroupes();   // buildRetroGroupesAllRows_ ou ton append existant
      appendImportLog_(ss, 'EXPORT_GROUPES_OK', gRes);
    } catch (eG) {
      appendImportLog_(ss, 'EXPORT_GROUPES_FAIL', String(eG));
    }

    // 4) MAILS (AFTER)
    try {
      var stage = String(readParam_(ss, 'MAIL_STAGE') || 'AFTER').trim().toUpperCase();
      if (stage === 'AFTER' && typeof runMailPipelineSelected_ === 'function') {
        runMailPipelineSelected_(ss, 'AFTER'); // worker dry-run-aware
      }
    } catch (e) {
      appendImportLog_(ss, 'MAIL_PIPELINE_FAIL', String(e));
    }

    return { ok: true, import: imp, rules: rFull, membres: mRes, groupes: gRes };

  } catch (e) {
    appendImportLog_(ss, 'FLOW_FULL_FAIL', String(e));
    return { ok: false, error: String(e) };

  } finally {
    endImportRun_(ctx);
  }
}

/* ------------------------------- HELPERS ------------------------------------ */

function _boolParam_(key, defVal) {
  var v = String(readParamValue(key) || '').trim().toLowerCase();
  if (!v) return !!defVal;
  return v === '1' || v === 'true' || v === 'yes' || v === 'oui';
}

/** Route unique vers l’implémentation des augmentations (LIB/global), peu importe le nom */
function runPostImportAugmentations_(seasonId, touched, opts) {
  if (typeof runPostImportAugmentations === 'function') {
    return runPostImportAugmentations(seasonId, touched, opts);
  }
  if (typeof LIB !== 'undefined' && LIB) {
    if (typeof LIB.runPostImportAugmentations === 'function') {
      return LIB.runPostImportAugmentations(seasonId, touched, opts);
    }
    if (typeof LIB.runPostImportAugmentations_ === 'function') {
      return LIB.runPostImportAugmentations_(seasonId, touched, opts);
    }
  }
  throw new Error('runPostImportAugmentations introuvable');
}
