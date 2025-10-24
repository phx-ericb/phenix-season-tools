function getSeasonList() {
  return _wrap('getSeasonList', function(){
    var cache = CacheService.getScriptCache();
    var CK = 'SEASON_LIST_v2';
    var hit = cache.get(CK);
    if (hit) { try { return _ok(JSON.parse(hit)); } catch(_){ cache.remove(CK); } }

    var props = _registry_();
    var list = [];
    try { list = JSON.parse(props.getProperty('SEASONS_JSON') || '[]'); } catch(_){ list = []; }

    // Validation Drive non-bloquante (petit lock); si indisponible, on sert la liste telle quelle
    var cleaned = [];
    var validated = false;
    var lock = null;
    try {
      lock = LockService.getScriptLock();
      if (lock.tryLock(500)) {
        list.forEach(function(s){
          try { DriveApp.getFileById(s.id).getId(); cleaned.push(s); } catch(e){}
        });
        validated = true;
      } else {
        cleaned = list.slice(); // pas de validation sous contention
      }
    } finally {
      if (lock) { try { lock.releaseLock(); } catch(_){ } }
    }

    if (validated && cleaned.length !== list.length) {
      props.setProperty('SEASONS_JSON', JSON.stringify(cleaned));
    }

    // Active
    var active = props.getProperty('ACTIVE_SEASON_ID');
    if (!active || !cleaned.some(function(s){ return s.id === active; })) {
      active = cleaned.length ? cleaned[0].id : null;
      if (active) props.setProperty('ACTIVE_SEASON_ID', active);
    }
    cleaned.forEach(function(s){ s.isActive = (s.id === active); });

    // cache 5 min — l’UI lit ça ultra-rapidement
    try { cache.put(CK, JSON.stringify(cleaned), 300); } catch(_){}

    return _ok(cleaned);
  });
}

function setActiveSeason(seasonId, opts) {
  return _wrap('setActiveSeason', function(){
    opts = opts || {};
    var props = _registry_();
    var list = JSON.parse(props.getProperty('SEASONS_JSON') || '[]');
    var found = list.find(function(s){ return s.id === seasonId; });
    if (!found) {
      var ssInfo = SpreadsheetApp.openById(seasonId);
      found = { id: ssInfo.getId(), title: ssInfo.getName(), url: ssInfo.getUrl() };
      list.push(found); props.setProperty('SEASONS_JSON', JSON.stringify(list));
    }
    props.setProperty('ACTIVE_SEASON_ID', seasonId);
    if (opts.updateSeasonLabel !== false) {
      var ss = SpreadsheetApp.openById(seasonId);
      _upsertParam_(ss, 'SEASON_LABEL', (opts.seasonLabel || _inferSeasonLabelFromTitle_(ss.getName())), 'string', 'Libellé humain de la saison (affichage/exports).');
    }
    return _ok({ seasonId: seasonId, info: found }, 'Active season set');
  });
}
function registerExistingSeason(seasonId, makeActive, seasonLabel) {
  return _wrap('registerExistingSeason', function(){
    var props = _registry_();
    var ss = SpreadsheetApp.openById(seasonId);
    var info = { id:ss.getId(), title:ss.getName(), url:ss.getUrl() };
    var list = JSON.parse(props.getProperty('SEASONS_JSON') || '[]');
    if (!list.some(function(s){ return s.id === info.id; })) { list.push(info); props.setProperty('SEASONS_JSON', JSON.stringify(list)); }
    if (makeActive) { props.setProperty('ACTIVE_SEASON_ID', info.id); _upsertParam_(ss,'SEASON_LABEL', (seasonLabel||_inferSeasonLabelFromTitle_(ss.getName())), 'string','Libellé humain de la saison (affichage/exports).'); }
    return _ok(info, 'Saison enregistrée' + (makeActive ? ' et active' : ''));
  });
}
function cloneSeason(srcSeasonId, newTitle) {
  return _wrap('cloneSeason', function(){
    if (!srcSeasonId || !newTitle) throw new Error('Missing source or title');
    var srcFile = DriveApp.getFileById(srcSeasonId);
    var parent = srcFile.getParents().hasNext() ? srcFile.getParents().next() : null;
    var copy = parent ? srcFile.makeCopy(newTitle, parent) : srcFile.makeCopy(newTitle);
    var ss = SpreadsheetApp.openById(copy.getId());
    var info = { id:ss.getId(), title:ss.getName(), url:ss.getUrl() };
    var list = JSON.parse(_registry_().getProperty('SEASONS_JSON') || '[]');
    list.push(info); _registry_().setProperty('SEASONS_JSON', JSON.stringify(list));
    return _ok(info, 'Season cloned');
  });
}

function API_setActiveSeason(seasonId) {
  seasonId = String(seasonId || '').trim();
  if (!seasonId) return { ok:false, error:'seasonId vide' };

  // persiste + prime le cache hardened
  PropertiesService.getScriptProperties().setProperty('ACTIVE_SEASON_ID', seasonId);
  try { setSeasonId_(seasonId); } catch(_) {}

  return { ok:true, seasonId: seasonId, at: Date.now() };
}
