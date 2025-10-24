/** ================= ERREURS – API UI ===================
 * - Mapping d'entêtes tolérant ("Passeport #", accents, casse)
 * - Normalisation passeport via normalizePassportPlain8_ si dispo
 * - Agrégation par groupe pour l'UI + compteurs par niveau
 * - Cache court (DocumentCache, 120s) + API d’invalidation
 * - Contrat d’API harmonisé : { ok:true, ... , error?:string }
 * - Tolérance aux alias de feuille: ERREURS / Erreurs / Errors
 */

/* ---------- Utilitaires partagés (fallbacks si absents) ---------- */

// Normalise une étiquette (accents -> ASCII, casse, ponctuation)
var ER_norm = this.ER_norm || function (s) {
  return String(s || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toLowerCase().replace(/[^a-z0-9]+/g, '').trim();
};

// ID de saison : utilise getSeasonId_() si dispo, sinon ScriptProperties
var ER_resolveSeasonId_ = this.ER_resolveSeasonId_ || function (overrideId) {
  if (overrideId) return overrideId;
  if (typeof getSeasonId_ === 'function') { var sid = getSeasonId_(); if (sid) return sid; }
  var sidProp = PropertiesService.getScriptProperties().getProperty('SEASON_SHEET_ID');
  if (sidProp) return sidProp;
  throw new Error("Aucun ID de saison : passe un seasonId, implémente getSeasonId_(), ou définis SEASON_SHEET_ID.");
};

// Ouvre le classeur de saison : utilise getSeasonSpreadsheet_() si dispo
var ER_openSeasonSpreadsheet_ = this.ER_openSeasonSpreadsheet_ || function (seasonId) {
  if (typeof getSeasonSpreadsheet_ === 'function') { var ss = getSeasonSpreadsheet_(seasonId); if (ss) return ss; }
  return SpreadsheetApp.openById(seasonId);
};

// Normalise passeport
function ER_normPass_(v) {
  if (typeof normalizePassportPlain8_ === 'function') return normalizePassportPlain8_(v);
  if (typeof normalizePassportToText8_ === 'function') {
    var t = normalizePassportToText8_(v);
    return (t && t[0] === "'") ? t.slice(1) : t;
  }
  var s = String(v == null ? '' : v).trim();
  if (!s) return '';
  if (s[0] === "'") s = s.slice(1);
  if (/^\d+$/.test(s) && s.length < 8) s = ('00000000' + s).slice(-8);
  return s;
}

// Normalise un niveau (err/warn/info/other)
function ER_normLevel_(v) {
  var t = String(v || '').toLowerCase().trim();
  if (t === 'err' || t === 'error' || t === 'erreur' || t === 'fatal' || t === 'critique') return 'err';
  if (t === 'warn' || t === 'warning' || t === 'avertissement') return 'warn';
  if (t === 'info' || t === 'information') return 'info';
  if (!t) return 'other';
  if (t === 'other') return 'other';
  return t;
}

/* ---------- Lecture & mapping ERREURS ---------- */

function ER_buildHeaderIndex_(headers) {
  var map = {};
  headers.forEach(function (h, i) { map[ER_norm(h)] = i; });
  function pick() { for (var i = 0; i < arguments.length; i++) { var k = ER_norm(arguments[i]); if (k in map) return map[k]; } return -1; }
  return {
    Passeport: pick('Passeport', 'Passeport #', 'Passport', 'ID'),
    Nom: pick('Nom', 'Last name', 'Famille'),
    Prenom: pick('Prénom', 'Prenom', 'First name'),
    Affichage: pick('Affichage', 'Display'),
    Categorie: pick('Catégorie', 'Categorie', 'Category'),
    Code: pick('Code'),
    Niveau: pick('Niveau', 'Level', 'Severity'),
    Saison: pick('Saison', 'Season'),
    Element: pick('Élément', 'Element'),
    Message: pick('Message'),
    Meta: pick('Meta', 'JSON', 'Data'),
    Date: pick('Date')
  };
}
function ER_val_(row, idx, key) { var i = idx[key]; return i >= 0 ? row[i] : ''; }

function ER_getSheetByAliases_(ss, aliases) {
  for (var i = 0; i < aliases.length; i++) {
    var sh = ss.getSheetByName(aliases[i]);
    if (sh) return sh;
  }
  return null;
}

function ER_readErrorsRaw_(seasonId) {
  var sid = ER_resolveSeasonId_(seasonId);
  var ss = ER_openSeasonSpreadsheet_(sid);
  var sh = ER_getSheetByAliases_(ss, ['ERREURS', 'Erreurs', 'Errors']);
  if (!sh) return { headers: [], rows: [] };

  var values = sh.getDataRange().getValues();
  if (!values || !values.length) return { headers: [], rows: [] };

  var headers = values[0].map(String);
  var idx = ER_buildHeaderIndex_(headers);

  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row || row.every(function (c) { return c === '' || c == null; })) continue;

    var metaRaw = ER_val_(row, idx, 'Meta'), metaObj = {};
    if (metaRaw && typeof metaRaw === 'string') {
      try { metaObj = JSON.parse(metaRaw); }
      catch (e) { metaObj = { _raw: String(metaRaw) }; }
    }

    var dc = ER_val_(row, idx, 'Date');
    var dateOut = (dc instanceof Date) ? dc : String(dc || '');

    out.push({
      passeport: ER_normPass_(ER_val_(row, idx, 'Passeport')),
      nom: String(ER_val_(row, idx, 'Nom') || ''),
      prenom: String(ER_val_(row, idx, 'Prenom') || ''),
      affichage: String(ER_val_(row, idx, 'Affichage') || ''),
      categorie: String(ER_val_(row, idx, 'Categorie') || ''),
      code: String(ER_val_(row, idx, 'Code') || ''),
      niveau: ER_normLevel_(ER_val_(row, idx, 'Niveau')),
      saison: String(ER_val_(row, idx, 'Saison') || ''),
      element: String(ER_val_(row, idx, 'Element') || ''),
      message: String(ER_val_(row, idx, 'Message') || ''),
      meta: metaObj,
      date: dateOut
    });
  }
  return { headers: headers, rows: out };
}

/* ---------- Agrégation pour l’UI ---------- */

function ER_groupKey_(row) { return [row.categorie || '', row.code || '', row.element || '', ER_normLevel_(row.niveau) || ''].join('||'); }
function ER_groupTitle_(row) { if (row.element) return row.element; if (row.categorie && row.code) return row.categorie + ' — ' + row.code; return row.code || row.categorie || 'Autres'; }

function ER_aggregateForUi_(rows) {
  var groups = {}, counts = { err: 0, warn: 0, info: 0, other: 0 };
  rows.forEach(function (row) {
    var sev = ER_normLevel_(row.niveau);
    var k = ER_groupKey_(row);
    if (!groups[k]) groups[k] = {
      key: k,
      titre: ER_groupTitle_(row),
      categorie: row.categorie || '',
      code: row.code || '',
      element: row.element || '',
      niveau: sev || '',
      items: []
    };
    groups[k].items.push(row);
    if (sev === 'err') counts.err++;
    else if (sev === 'warn') counts.warn++;
    else if (sev === 'info') counts.info++;
    else counts.other++;
  });
  var rank = { err: 0, warn: 1, info: 2, other: 3 };
  var list = Object.keys(groups).map(function (k) { var g = groups[k]; g.total = g.items.length; return g; })
    .sort(function (a, b) {
      var ra = (rank[a.niveau] != null) ? rank[a.niveau] : 9, rb = (rank[b.niveau] != null) ? rank[b.niveau] : 9;
      return (ra !== rb) ? (ra - rb) : (b.total - a.total);
    });
  return { groups: list, counts: counts, total: rows.length };
}

/* ---------- Mail outbox (stats historiques) ---------- */

function ER_mailOutboxStats_(seasonId) {
  var sid = ER_resolveSeasonId_(seasonId);
  var ss = ER_openSeasonSpreadsheet_(sid);
  var sh = ss.getSheetByName('MAIL_OUTBOX');
  if (!sh) return { totalErrors: 0, byCode: {} };
  var last = sh.getLastRow(); if (last < 2) return { totalErrors: 0, byCode: {} };
  var types = sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return (r[0] || '').toString().trim(); });
  var byCode = {}, total = 0;
  types.forEach(function (t) { if (!t || t === 'INSCRIPTION_NEW') return; byCode[t] = (byCode[t] || 0) + 1; total++; });
  return { totalErrors: total, byCode: byCode };
}

/* ---------- Cache & helpers payload ---------- */

function ER_emptyPayload_(sid) {
  return {
    seasonSheetId: sid || '',
    updatedAt: new Date().toISOString(),
    total: 0,
    counts: { err: 0, warn: 0, info: 0, other: 0 },
    groups: [],
    mailOutbox: { totalErrors: 0, byCode: {} },
    historicErrorTotal: 0
  };
}
function ER_cacheKey_(sid) { return 'ER_UI_v3::' + sid; }

/* ---------- Build payload (tolérant, jamais throw) ---------- */

function ER_getUiPayload_(seasonId) {
  var sid;
  try {
    sid = ER_resolveSeasonId_(seasonId);
  } catch (e) {
    var p0 = ER_emptyPayload_('');
    p0.ok = true;
    p0.error = 'season_id_missing: ' + String((e && e.message) || e);
    return p0;
  }

  var cache = CacheService.getDocumentCache();
  var key = ER_cacheKey_(sid);

  // Try cache
  try {
    var hit = cache.get(key);
    if (hit) {
      var cached = JSON.parse(hit);
      if (cached && Array.isArray(cached.groups)) return cached;
    }
  } catch (_) {}

  // Build fresh
  var payload;
  try {
    var raw  = ER_readErrorsRaw_(sid);
    var aggr = ER_aggregateForUi_(raw.rows);
    var mo   = ER_mailOutboxStats_(sid);
    payload = {
      seasonSheetId: sid,
      updatedAt: new Date().toISOString(),
      total: aggr.total,
      counts: aggr.counts,
      groups: aggr.groups,
      mailOutbox: mo,
      historicErrorTotal: mo.totalErrors
    };
  } catch (e) {
    payload = ER_emptyPayload_(sid);
    payload.error = String((e && e.message) || e);
  }

  payload.ok = true;
  try {
    if (payload && Array.isArray(payload.groups)) cache.put(key, JSON.stringify(payload), 120);
  } catch (_) {}
  return payload;
}

/** ===== API exposées ===== */

function API_listErreurs(seasonSheetId) {
  
  var p = ER_getUiPayload_(seasonSheetId);
  if (p && p._error) { p.error = p._error; delete p._error; }
  if (!('ok' in p)) p.ok = true;
  if (!p.seasonSheetId) { try { p.seasonSheetId = ER_resolveSeasonId_(seasonSheetId); } catch (_) {} }
  if (!p || !Array.isArray(p.groups)) {
    var sid = ''; try { sid = ER_resolveSeasonId_(seasonSheetId); } catch (_) {}
    var empty = ER_emptyPayload_(sid);
    empty.ok = true;
    empty.error = (p && p.error) ? p.error : 'payload_validation_failed';
    return empty;
  }
  return p;
}

function API_ER_invalidateCache(seasonSheetId) {
  var sid = ER_resolveSeasonId_(seasonSheetId);
  var key = ER_cacheKey_(sid);
  try { CacheService.getDocumentCache().remove(key); } catch (_) {}
  try { CacheService.getScriptCache().remove(key); } catch (_) {}
  return { ok: true, sid: sid };
}

// Helpers UI
function API_ER_warm(seasonSheetId) { return API_listErreurs(seasonSheetId); }

function API_listErreurs_debug(seasonSheetId) {
  try {
    var sid;
    try {
      sid = ER_resolveSeasonId_(seasonSheetId);
    } catch (e) {
      return {
        ok: false,
        sid: '',
        error: 'season_id_missing: ' + String((e && e.message) || e),
        hint: 'Passe seasonSheetId à API_listErreurs / API_ER_invalidateCache, ou définis SEASON_SHEET_ID / getSeasonId_().'
      };
    }
    var raw = ER_readErrorsRaw_(sid);
    var ag  = ER_aggregateForUi_(raw.rows);
    return {
      ok: true,
      sid: sid,
      headers: raw.headers,
      sample: (raw.rows || []).slice(0, 5),
      counts: ag.counts,
      total: ag.total
    };
  } catch (e) {
    return { ok: false, error: String((e && e.message) || e) };
  }
}
