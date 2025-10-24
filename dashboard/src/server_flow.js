/* ============================== server_flow.js =============================== */

function __ping_aug__() {
  var seasonId = getSeasonId_();
  var ok = (typeof runPostImportAugmentations_ === 'function');
  if (!ok) throw new Error('Hook not visible');
  // Toujours (re)construire les pivots, même en dry-run (safe: écrit dans le classeur)
  runPostImportAugmentations_(seasonId, [], { isFull: true, isDryRun: false });
}

function bindActiveSpreadsheet_(ss) {
  try { SpreadsheetApp.setActiveSpreadsheet(ss); } catch (_) {}
}

function _setFlag_(k, v) {
  PropertiesService.getScriptProperties().setProperty(k, String(v));
}
function _getFlag_(k) {
  return PropertiesService.getScriptProperties().getProperty(k);
}

/** Lock global pour sérialiser FULL/INCR */
function acquireImportLock_(ms) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(Math.max(500, Math.min(ms || 15000, 20000)))) {
    throw new Error('Import déjà en cours (lock global indisponible)');
  }
  _setFlag_('PHENIX_IMPORT_LOCK', '1'); // visible par les endpoints de lecture
  return lock;
}
function releaseImportLock_(lock) {
  try { _setFlag_('PHENIX_IMPORT_LOCK', '0'); } catch(_) {}
  try { lock && lock.releaseLock(); } catch(_) {}
}


function runImportRulesExportsAuto(){
  var list = (typeof getLastTouchedPassports_==='function')
    ? getLastTouchedPassports_()
    : (typeof _getTouchedPassportsArray_==='function'
        ? _getTouchedPassportsArray_()
        : []);
  if (Array.isArray(list) && list.length > 0) {
    return runImportRulesExportsIncr();   // INCR par défaut s’il y a du contenu
  }
  return runImportRulesExportsFull();     // sinon full
}



/* ----------------------------- INCR PIPELINE -------------------------------- */
function runImportRulesExportsIncr() {
  var seasonId = (typeof getSeasonId_ === 'function') ? getSeasonId_() : null;
  var lockGlobal = acquireImportLock_(15000);               // aligne avec FULL (était 8000 chez toi)
  var ss = getSeasonSpreadsheet_(seasonId);
  bindActiveSpreadsheet_(ss);

  // ⚠️ Hints legacy (comme le FULL)
  try {
    PropertiesService.getScriptProperties().setProperty('PHENIX_IMPORT_ACTIVE_SID', seasonId);
    PropertiesService.getScriptProperties().setProperty('PHENIX_IMPORT_RUNNING', '1');
  } catch(_) {}

  var ctx = (typeof startImportRun_ === 'function')
    ? startImportRun_({ seasonId: seasonId, source: 'dashboard-incr' })
    : { runId: Utilities.getUuid(), ss: ss };

  try {
    // (0) Import
    var imp = (typeof runImporterDonneesSaison === 'function')
      ? runImporterDonneesSaison(seasonId)
      : null;

    // (0.5) VM refresh (inchangé)
    try {
      var vm = (typeof API_VM_fullRefresh === 'function')
        ? API_VM_fullRefresh(false)
        : (typeof API_VM_refreshSeasonSubset === 'function'
            ? API_VM_refreshSeasonSubset()
            : { ok:false, error:'API_VM_* not found' });
      appendImportLog_(ss, 'VM_REFRESH_INCR_OK', vm);
    } catch (eVM) {
      appendImportLog_(ss, 'VM_REFRESH_INCR_FAIL', String(eVM));
    }

    // (1) Passeports touchés — 3 sources, dans cet ordre:
    //    A) ce que l'import retourne (si tu le fais retourner)
    //    B) DocumentProperties (fallback)
    //    C) Diff rapide (fallback-final)
    var touched = [];
    try {
      if (imp && Array.isArray(imp.touchedPassports)) touched = imp.touchedPassports.slice();
    } catch(_) {}

    if (!touched.length && typeof _getTouchedPassportsArray_ === 'function') {
      touched = _getTouchedPassportsArray_() || [];
    }

    if (!touched.length) {
      touched = _computeTouchedPassportsFallback_(ss) || [];   // ⇒ code plus bas
      appendImportLog_(ss, 'TOUCHED_FALLBACK_USED', { count: touched.length });
    }

    // normalise (txt-8)
    touched = Array.from(new Set((touched||[])
      .map(function (p){ return String(p||'').replace(/\D/g,'').padStart(8,'0'); })
      .filter(Boolean)));

    // expose pour d’autres exporteurs legacy
    try {
      PropertiesService.getDocumentProperties()
        .setProperty('LAST_TOUCHED_PASSPORTS', JSON.stringify(touched));
    } catch(_) {}

    appendImportLog_(ss, 'TOUCHED_INCR', { count: touched.length });

    // (2) AUGMENTATIONS
    try {
      var augOpts = { isFull: false, isDryRun: false };
      if (touched.length) {
        runPostImportAugmentations_(seasonId, new Set(touched), augOpts);
        appendImportLog_(ss, 'AUG_INCR_PRE_OK', { count: touched.length });
      } else {
        appendImportLog_(ss, 'AUG_INCR_PRE_SKIP', { reason: 'no-passports' });
      }
    } catch (eAugPre) {
      appendImportLog_(ss, 'AUG_INCR_PRE_FAIL', String(eAugPre));
    }

    // (3) RÈGLES INCR
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

    // (4) EXPORTS INCR
    try {
      var combined = _boolParam_ ? _boolParam_('RETRO_COMBINED_XLSX', false) : false;
      var rExp = (typeof runRetroExportsIncr === 'function')
        ? runRetroExportsIncr(touched, { combined: combined })
        : { ok:false, note:'no-exporter' };
      appendImportLog_(ss, 'EXPORTS_INCR_OK', rExp);
    } catch (eExp) {
      appendImportLog_(ss, 'EXPORTS_INCR_FAIL', String(eExp));
    }

    // (5) MAILS AFTER
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
  try { PropertiesService.getScriptProperties().setProperty('PHENIX_IMPORT_RUNNING', '0'); } catch(_) {}
  releaseImportLock_(lockGlobal);
  try { endImportRun_(ctx); } catch(_) {}

  // ✅ Trie le log après import pour lisibilité
  try {
    var logSheet = ss.getSheetByName('IMPORT_LOG');
    if (logSheet) logSheet.sort(1, true);
  } catch (eSort) {
    Logger.log('Erreur tri IMPORT_LOG: ' + String(eSort));
  }
}
}

/**
 * Fallback “diff” si personne ne fournit la liste touchée.
 * Ici on calcule un hash par passeport à partir des colonnes utiles aux exports rétro,
 * on compare au snapshot précédent, et on renvoie les passeports modifiés.
 */
function _computeTouchedPassportsFallback_(ss) {
  try {
    // 1) Lis les données minimales nécessaires aux 3 exports (à adapter si besoin)
    //    - JOUEURS (passeport, nom, prénom, sexe, dob, email, etc.)
    //    - ASSIGN / GROUPES (affectations)
    var joueurs = (typeof getJoueursSheetValues_==='function')
      ? getJoueursSheetValues_(ss) : []; // [[...]] avec passeport en col 0
    var assign  = (typeof getAssignSheetValues_==='function')
      ? getAssignSheetValues_(ss)  : []; // [[...]] avec passeport en col 0

    var nowMap = {};
    function h(s){ return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, s).map(function(b){return (b+256).toString(16).slice(-2)}).join(''); }

    // 2) Construis un “fingerprint” par passeport (concat de champs clés)
    joueurs.forEach(function(r){
      var p = String(r && r[0] || '').replace(/\D/g,'').padStart(8,'0');
      if (!p) return;
      var sig = JSON.stringify([r[1],r[2],r[3],r[4],r[5],r[6]]); // ex. nom/prénom/sexe/dob/email/statuts...
      nowMap[p] = h('J|'+sig+'|'+(nowMap[p]||''));              // préfixe J
    });
    assign.forEach(function(r){
      var p = String(r && r[0] || '').replace(/\D/g,'').padStart(8,'0');
      if (!p) return;
      var sig = JSON.stringify([r[1],r[2],r[3]]);               // ex. Catégorie, Groupe, Rôle...
      nowMap[p] = h('A|'+sig+'|'+(nowMap[p]||''));              // préfixe A
    });

    // 3) Compare au snapshot précédent
    var props = PropertiesService.getDocumentProperties();
    var prevRaw = props.getProperty('RETRO_HASH_SNAPSHOT') || '{}';
    var prev = {};
    try { prev = JSON.parse(prevRaw); } catch(_) {}
    var touched = [];
    Object.keys(nowMap).forEach(function(p){
      if (nowMap[p] !== prev[p]) touched.push(p);
    });

    // 4) Sauvegarde le snapshot
    props.setProperty('RETRO_HASH_SNAPSHOT', JSON.stringify(nowMap));

    return touched;
  } catch(e) {
    appendImportLog_(ss, 'TOUCHED_FALLBACK_FAIL', String(e));
    return [];
  }
}


/* ----------------------------- FULL PIPELINE -------------------------------- */
function runImportRulesExportsFull() {
  var seasonId = getSeasonId_();
  var lockGlobal = acquireImportLock_(15000);
  var ss = getSeasonSpreadsheet_(seasonId);

  bindActiveSpreadsheet_(ss);
  try {
    PropertiesService.getScriptProperties().setProperty('PHENIX_IMPORT_ACTIVE_SID', seasonId);
    PropertiesService.getScriptProperties().setProperty('PHENIX_IMPORT_RUNNING', '1');
  } catch(_) {}

  var ctx = startImportRun_({ seasonId: seasonId, source: 'dashboard-full' });

  var imp = null, rFull = null, mRes = null, gRes = null;
  try {
    // 0) Import principal (met à jour JOUEURS)
    imp = (typeof runImporterDonneesSaison === 'function') ? runImporterDonneesSaison(seasonId) : null;

    // 1) AUGMENTATIONS
    try {
      runPostImportAugmentations_(seasonId, [], { isFull: true, isDryRun: false });
      appendImportLog_(ss, 'AUG_FULL_PRE_OK', { via: 'SHIM' });
    } catch (eAugPre) {
      appendImportLog_(ss, 'AUG_FULL_PRE_FAIL', String(eAugPre));
    }

    // 2) RÈGLES FULL
    try {
      if (typeof runEvaluateRules === 'function') {
        rFull = runEvaluateRules();
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

    // ✅ 3) MISE À JOUR PHOTOS (avant export)
    try {
      const centralId = (typeof getCentralValidationSpreadsheetId_ === 'function')
        ? getCentralValidationSpreadsheetId_()
        : null;
      if (typeof syncMembresGlobalSubsetFromCentral_ === 'function' && centralId) {
        syncMembresGlobalSubsetFromCentral_(seasonId, centralId);
      }
      if (typeof applySeasonMembresToJoueurs_ === 'function') {
        applySeasonMembresToJoueurs_(seasonId);
      }
      if (typeof refreshPhotoStrInJoueurs_ === 'function') {
        refreshPhotoStrInJoueurs_(seasonId);
      } else if (typeof refreshPhotoStrInJoueurs__ === 'function') {
        refreshPhotoStrInJoueurs__(seasonId); // fallback
      }
      appendImportLog_(ss, 'VM_REFRESH_FULL_OK', { ok: true });
    } catch (eRefresh) {
      appendImportLog_(ss, 'VM_REFRESH_FULL_FAIL', String(eRefresh));
    }

    // 4) EXPORTS
    try { mRes = runExportRetroMembres(); appendImportLog_(ss, 'EXPORT_MEMBRES_OK', mRes); }
    catch (eM) { appendImportLog_(ss, 'EXPORT_MEMBRES_FAIL', String(eM)); }
    try { gRes = runExportRetroGroupes(); appendImportLog_(ss, 'EXPORT_GROUPES_OK', gRes); }
    catch (eG) { appendImportLog_(ss, 'EXPORT_GROUPES_FAIL', String(eG)); }

    // 5) MAILS
    try {
      var stage = String(readParam_(ss, 'MAIL_STAGE') || 'AFTER').trim().toUpperCase();
      if (stage === 'AFTER' && typeof runMailPipelineSelected_ === 'function') {
        runMailPipelineSelected_(ss, 'AFTER');
      }
    } catch (e) {
      appendImportLog_(ss, 'MAIL_PIPELINE_FAIL', String(e));
    }

  } catch (e) {
    appendImportLog_(ss, 'FLOW_FULL_FAIL', String(e));
    return { ok: false, error: String(e) };

  } finally {
    try { PropertiesService.getScriptProperties().setProperty('PHENIX_IMPORT_RUNNING', '0'); } catch(_) {}
    releaseImportLock_(lockGlobal);
    try { endImportRun_(ctx); } catch(_) {}
  }

  return { ok: true, import: imp, rules: rFull, membres: mRes, groupes: gRes };
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
