/**
 * retro_groupes.gs — v0.8
 * - Construit l'export "Import Rétro - Groupes" (10 colonnes)
 * - Paramétrable: patterns d'ignorés, élite, soccer adapté, gabarits Groupe/Catégorie
 * - Support d’un mapping saisonnier via MAPPINGS (section "GROUPES")
 * - Respecte CANCELLED/EXCLUDE_FROM_EXPORT et le moteur de règles (RETRO_RULES_JSON)
 * - U / U2 robustes: DOB -> U/U2, fallback extraction libellé article
 *
 * Colonnes exportées:
 *  "Identifiant unique","Nom","Prénom","Date de naissance","#","Couleur","Sous-groupe","Position","Équipe/Groupe","Catégorie"
 */

/* ===================== Param keys ===================== */
if (typeof PARAM_KEYS === 'undefined') { var PARAM_KEYS = {}; }
PARAM_KEYS.RETRO_GROUP_SHEET_NAME        = PARAM_KEYS.RETRO_GROUP_SHEET_NAME        || 'RETRO_GROUP_SHEET_NAME';
PARAM_KEYS.RETRO_GROUP_EXPORTS_FOLDER_ID = PARAM_KEYS.RETRO_GROUP_EXPORTS_FOLDER_ID || 'RETRO_GROUP_EXPORTS_FOLDER_ID';

PARAM_KEYS.RETRO_IGNORE_FEES_CSV         = PARAM_KEYS.RETRO_IGNORE_FEES_CSV         || 'RETRO_IGNORE_FEES_CSV';
PARAM_KEYS.RETRO_GROUP_ELITE_KEYWORDS    = PARAM_KEYS.RETRO_GROUP_ELITE_KEYWORDS    || 'RETRO_GROUP_ELITE_KEYWORDS';

PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS       = PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS       || 'RETRO_GROUP_SA_KEYWORDS';
PARAM_KEYS.RETRO_GROUP_SA_GROUPE_LABEL   = PARAM_KEYS.RETRO_GROUP_SA_GROUPE_LABEL   || 'RETRO_GROUP_SA_GROUPE_LABEL';
PARAM_KEYS.RETRO_GROUP_SA_CATEG_LABEL    = PARAM_KEYS.RETRO_GROUP_SA_CATEG_LABEL    || 'RETRO_GROUP_SA_CATEG_LABEL';

PARAM_KEYS.RETRO_GROUP_GROUPE_FMT        = PARAM_KEYS.RETRO_GROUP_GROUPE_FMT        || 'RETRO_GROUP_GROUPE_FMT';
PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT     = PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT     || 'RETRO_GROUP_CATEGORIE_FMT';

PARAM_KEYS.RETRO_RULES_JSON              = PARAM_KEYS.RETRO_RULES_JSON              || 'RETRO_RULES_JSON';

/* ===================== Helpers ===================== */
function _rg_nrm_(s){ s=String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s; }
function _rg_low_(s){ return _rg_nrm_(s).toLowerCase().trim(); }

function _rg_isActiveRow_(r){
  var can = String(r[CONTROL_COLS.CANCELLED]||'').toLowerCase()==='true';
  var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT]||'').toLowerCase()==='true';
  var st  = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
  return !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
}
function _rg_containsAny_(txt, csv){
  var t=_rg_low_(txt||'');
  return String(csv||'').split(',').map(_rg_low_).filter(Boolean).some(function(p){return t.indexOf(p)!==-1;});
}
function _rg_tpl_(tpl, vars){
  tpl = String(tpl==null?'':tpl);
  return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function(_,k){ return (vars && k in vars && vars[k]!=null)? String(vars[k]) : ''; });
}
function _rg_csvEsc_(v){ v=v==null?'':String(v).replace(/"/g,'""'); return /[",\n;]/.test(v)?('"'+v+'"'):v; }
function _rg_genreInitiale_(row){
  var g = (row['Identité de genre']||row['Identité de Genre']||row['Genre']||row['Sexe']||'').toString().trim().toUpperCase();
  return g ? g.charAt(0) : '';
}

/* ====== U & U2 (DOB + fallback libellé article) ====== */
function _rg_pad2_(n){ n=Number(n||0); return (n<10?('0'+n):String(n)); }
function _rg_deriveBirthYearFromRow_(row){
  var dn = row['Date de naissance'] || row['Naissance'] || '';
  if (dn instanceof Date) return dn.getFullYear();
  if (dn) {
    var m = String(dn).match(/(19|20)\d{2}/);
    if (m) return parseInt(m[0],10);
  }
  return null;
}
function _rg_ageCat_(birthYear, seasonYear){
  if (!birthYear || !seasonYear) return '';
  var age = seasonYear - birthYear;
  if (age < 4 || age > 99) return '';
  return 'U'+_rg_pad2_(age); // "U09"
}
function _rg_U_(birthYear, seasonYear){
  var u2 = _rg_ageCat_(birthYear, seasonYear);
  return u2 ? ('U'+parseInt(u2.slice(1),10)) : ''; // "U9", "U10", ...
}
function _rg_extractUFromArticle_(articleName){
  var s = String(articleName||'').toUpperCase();
  var m = s.match(/U\s*[-–]?\s*(\d{1,2})/);
  return m ? ('U'+parseInt(m[1],10)) : '';
}
function _rg_computeUandU2_(row, seasonYear, feeName){
  var by = _rg_deriveBirthYearFromRow_(row);
  var U='', U2='';
  if (by){
    U2 = _rg_ageCat_(by, seasonYear);
    if (U2) U = 'U'+parseInt(U2.slice(1),10);
  }
  if (!U){
    var uTxt = _rg_extractUFromArticle_(feeName);
    if (uTxt){
      U = uTxt;
      var n=parseInt(uTxt.replace(/\D/g,''),10);
      if (!isNaN(n)) U2='U'+_rg_pad2_(n);
    }
  }
  return { U:U, U2:U2 };
}

/* ====== Fallback extraction "U.. genre" depuis article ====== */
function _rg_extractFromArticlePair_(articleName){
  var s = String(articleName||'');
  var re = /U[-\s]?(\d{2}).*?(F[ée]minin|M[âa]sculin)/i;
  var m = s.match(re);
  if (m) {
    var u = 'U'+m[1];
    var g = m[2].toUpperCase().charAt(0);
    return { U:u, genreInitiale:g };
  }
  return null;
}

/* ===================== Règles (réutilise le moteur des membres) ===================== */
function _rg_loadRules_(ss){
  if (typeof loadRetroRules_ === 'function') return loadRetroRules_(ss);
  return [];
}
function _rg_applyRowRulesMaybeSkip_(rules, row, ctx){
  if (!rules || !rules.length || typeof applyRetroRowRules_ !== 'function') return false;
  var fakeMember = {};
  var res = applyRetroRowRules_(rules, 'inscriptions', row, fakeMember, ctx);
  return !!(res && res.skip);
}

/** Charge la table MAPPINGS unifiée (header unique). */
function _loadUnifiedGroupMappings_(ss){
  var sh = ss.getSheetByName(SHEETS.MAPPINGS);
  var out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  var data = sh.getDataRange().getValues();
  var H = (data[0]||[]).map(function(h){ return String(h||'').trim(); });
  function idx(k){ var i=H.indexOf(k); return i<0?null:i; }
  var iType=idx('Type'), iAli=idx('AliasContains'), iUmin=idx('Umin'), iUmax=idx('Umax'),
      iGen=idx('Genre'), iG=idx('GroupeFmt'), iC=idx('CategorieFmt'), iEx=idx('Exclude'), iPr=idx('Priority');
  if (iType==null || iAli==null) return out;

  for (var r=1; r<data.length; r++){
    var row=data[r]||[];
    if (!row.some(function(c){return String(c||'').trim();})) continue;
    out.push({
      Type: String(row[iType]||'').toLowerCase(),                // member | article
      AliasContains: String(row[iAli]||''),
      Umin: isNaN(parseInt(row[iUmin],10))?null:parseInt(row[iUmin],10),
      Umax: isNaN(parseInt(row[iUmax],10))?null:parseInt(row[iUmax],10),
      Genre: String(row[iGen]||'*').toUpperCase(),
      GroupeFmt: String(row[iG]||''),
      CategorieFmt: String(row[iC]||''),
      Exclude: String(row[iEx]||'').toLowerCase()==='true',
      Priority: isNaN(parseInt(row[iPr],10))?100:parseInt(row[iPr],10)
    });
  }
  // tri par priorité croissante
  out.sort(function(a,b){ return a.Priority - b.Priority; });
  return out;
}

function _tpl_(tpl, vars){
  tpl = String(tpl==null?'':tpl);
  return tpl.replace(/{{\s*([\w.]+)\s*}}/g, function(_,k){ return (vars && k in vars && vars[k]!=null)? String(vars[k]) : ''; });
}

function _low_(s){ s=String(s==null?'':s); try{s=s.normalize('NFD').replace(/[\u0300-\u036f]/g,'');}catch(e){} return s.toLowerCase().trim(); }

/** Applique la première règle qui matche pour le type demandé. */
function _applyUnifiedMapping_(maps, type, feeName, vars){
  var s = _low_(feeName||'');
  for (var i=0;i<maps.length;i++){
    var m = maps[i];
    if (m.Type !== type) continue;
    if (!m.AliasContains) continue;
    if (s.indexOf(_low_(m.AliasContains)) === -1) continue;

    // Genre
    if (m.Genre && m.Genre !== '*' && m.Genre !== (vars.genreInitiale||'')) continue;
    // U (utilise vars.U)
    if (m.Umin != null || m.Umax != null){
      var uNum = 0;
      if (vars.U){ var mm = String(vars.U).match(/^U(\d{1,2})$/i); if (mm) uNum = parseInt(mm[1],10); }
      if (!uNum) continue;
      if (m.Umin != null && uNum < m.Umin) continue;
      if (m.Umax != null && uNum > m.Umax) continue;
    }
    if (m.Exclude) return { exclude:true };
    return {
      groupe:    m.GroupeFmt    ? _tpl_(m.GroupeFmt, vars)    : '',
      categorie: m.CategorieFmt ? _tpl_(m.CategorieFmt, vars) : ''
    };
  }
  return null;
}

/* ===================== Construction des lignes ===================== */
function buildRetroGroupesRows(seasonSheetId){
  var ss   = getSeasonSpreadsheet_(seasonSheetId);
  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);

  var rules     = _rg_loadRules_(ss);
  var mappings  = _loadUnifiedGroupMappings_(ss);

  var ignoreCsv = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV)      || 'senior,u-sé,adulte,ligue';
  var eliteCsv  = readParam_(ss, PARAM_KEYS.RETRO_GROUP_ELITE_KEYWORDS) || 'D1+,LDP,Ligue';
  var saCsv     = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_KEYWORDS)    || 'soccer adapté,soccer adapte';
  var saGrp     = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_GROUPE_LABEL)|| 'Adapté (tous)';
  var saCat     = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SA_CATEG_LABEL) || 'Adapté';

  var grpFmtDef = readParam_(ss, PARAM_KEYS.RETRO_GROUP_GROUPE_FMT)     || '{{U}}{{genreInitiale}}';
  var catFmtDef = readParam_(ss, PARAM_KEYS.RETRO_GROUP_CATEGORIE_FMT)  || '{{U}} {{genreInitiale}}';

  var header = ["Identifiant unique", "Nom", "Prénom", "Date de naissance", "#", "Couleur", "Sous-groupe", "Position", "Équipe/Groupe", "Catégorie"];
  var rows   = [];

  var active = insc.rows.filter(_rg_isActiveRow_);
  if (!active.length) return { header: header, rows: rows, nbCols: header.length };

  var seasonLabel = readParam_(ss, 'SEASON_LABEL') || (active[0] && active[0]['Saison']) || '';
  var seasonYear  = parseSeasonYear_(seasonLabel);
  var _normPass = (typeof normalizePassportPlain8_ === 'function')
  ? normalizePassportPlain8_
  : function(v){ return String(v == null ? '' : v); };

  active.forEach(function(r){
    // (1) Règles
    var ctx = { ss:ss, catalog: (typeof _loadArticlesCatalog_==='function' ? _loadArticlesCatalog_(ss) : {match:function(){return null;}}) };
    if (_rg_applyRowRulesMaybeSkip_(rules, r, ctx)) return;

    var pass = r['Passeport #']; if (!pass) return;

    var feeName = r['Nom du frais'] || r['Frais'] || r['Produit'] || '';
    // (2) Ignorés & élite
    if (_rg_containsAny_(feeName, ignoreCsv)) return;
    if (_rg_containsAny_(feeName, eliteCsv))  return;

    // (3) Soccer adapté?
    if (_rg_containsAny_(feeName, saCsv)) {
      rows.push([
        _normPass(pass),
        (r['Nom']||''),
        (r['Prénom']||r['Prenom']||''),
        (r['Date de naissance']||r['Naissance']||''),
        "", "", "", "",
        saGrp,
        saCat
      ]);
      return;
    }

    // (4) U/U2 + genre
    var UU2 = _rg_computeUandU2_(r, seasonYear, feeName);
    var U   = UU2.U || '';
    var U2  = UU2.U2 || '';
    var gi  = _rg_genreInitiale_(r) || '';
    var vars = {
      U: U,
      U2: U2,
      ageCat: U2, // alias
      genreInitiale: gi,
      genre: (gi==='F'?'Féminin':(gi==='M'?'Masculin':(gi==='X'?'Mixte':''))),
      saison: seasonLabel,
      annee: seasonYear,
      article: feeName
    };

    // (5) MAPPINGS saisonniers (prioritaires)
    var mapRes = _applyUnifiedMapping_(mappings, 'member', feeName, vars);
    if (mapRes && mapRes.exclude) return;

    var groupe = (mapRes && mapRes.groupe) || '';
    var categ  = (mapRes && mapRes.categorie) || '';

    // (6) Fallback: extraction depuis libellé article, sinon fmt défaut
    if (!groupe || !categ) {
      var fromArt = _rg_extractFromArticlePair_(feeName);
      var Ux  = fromArt && fromArt.U  ? fromArt.U  : U;
      var gix = fromArt && fromArt.genreInitiale ? fromArt.genreInitiale : gi;
      // Régénère aussi U2 si on a seulement U depuis libellé
      var U2x = U2;
      if (!U2x && Ux){ var n=parseInt(String(Ux).replace(/\D/g,''),10); if(!isNaN(n)) U2x='U'+_rg_pad2_(n); }
      var v2  = { U:Ux, U2:U2x, ageCat:U2x, genreInitiale:gix, genre:(gix==='F'?'Féminin':(gix==='M'?'Masculin':(gix==='X'?'Mixte':''))), saison:seasonLabel, annee:seasonYear, article:feeName };
      if (!groupe) groupe = _rg_tpl_(grpFmtDef, v2);
      if (!categ)  categ  = _rg_tpl_(catFmtDef, v2);
    }

    // (7) Si on n'a toujours rien -> skip
    if (!groupe && !categ) return;

    rows.push([
      _normPass(pass),
      (r['Nom']||''),
      (r['Prénom']||r['Prenom']||''),
      (r['Date de naissance']||r['Naissance']||''),
      "", "", "", "",
      groupe,
      categ
    ]);
  });

  return { header: header, rows: rows, nbCols: header.length };
}

/* ===================== Écriture feuille & export XLSX ===================== */

function writeRetroGroupesSheet(seasonSheetId){
  var ss  = getSeasonSpreadsheet_(seasonSheetId);
  var out = buildRetroGroupesRows(seasonSheetId);

  var sheetName = readParam_(ss, PARAM_KEYS.RETRO_GROUP_SHEET_NAME) || 'Import Rétro - Groupes';
  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

  sh.clearContents();
  sh.getRange(1,1,1,out.nbCols).setValues([out.header]);
  if (out.rows.length) {
    sh.getRange(2,1,out.rows.length,out.nbCols).setValues(out.rows);
    sh.autoResizeColumns(1, out.nbCols);
    if (sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,1).setNumberFormat('@'); // Passeport texte
  }
  appendImportLog_(ss, 'RETRO_GROUPES_SHEET_OK', 'rows='+out.rows.length);
  return out.rows.length;
}

/** Export XLSX rapide — Rétro Groupes */
function exportRetroGroupesXlsxToDrive(seasonSheetId){
  var ss  = getSeasonSpreadsheet_(seasonSheetId);
  var out = buildRetroGroupesRows(seasonSheetId); // {header, rows, nbCols}

  // 1) Classeur temporaire minimal
  var temp = SpreadsheetApp.create('Export temporaire - Import Retro Groupes');
  var tmp  = temp.getSheets()[0];
  tmp.setName('Export');

  // 2) Écriture header + data
  var all = [out.header].concat(out.rows);

  // Normalise Passeport -> texte 'XXXXXXXX si helper dispo
  if (typeof normalizePassportToText8_ === 'function') {
    for (var i = 1; i < all.length; i++) {
      all[i][0] = normalizePassportToText8_(all[i][0]);
    }
  }
  if (all.length) {
    tmp.getRange(1, 1, all.length, out.nbCols).setValues(all);
    if (all.length > 1) tmp.getRange(2, 1, all.length - 1, 1).setNumberFormat('@'); // A en texte
  }
  SpreadsheetApp.flush();

  // 3) Export XLSX
  var url  = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(url, { headers:{ Authorization:'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  var ts   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HHmm');
  blob.setName('Export_Retro_Groupes_' + ts + '.xlsx');

  // 4) Destination
  var folderId = readParam_(ss, PARAM_KEYS.RETRO_GROUP_EXPORTS_FOLDER_ID)
              || readParam_(ss, PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID) || '';
  var dest = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var file = dest.createFile(blob);

  // 5) Nettoyage + log
  DriveApp.getFileById(temp.getId()).setTrashed(true);
  appendImportLog_(ss, 'RETRO_GROUPES_XLSX_OK_FAST', file.getName() + ' -> ' + dest.getName() + ' (rows=' + out.rows.length + ')');

  return { fileId:file.getId(), name:file.getName(), rows: out.rows.length };
}

/* ========== Exposition facultative via Library ========== */
if (typeof Library !== 'undefined') {
  Library.buildRetroGroupesRows = buildRetroGroupesRows;
  Library.writeRetroGroupesSheet = writeRetroGroupesSheet;
  Library.exportRetroGroupesXlsxToDrive = exportRetroGroupesXlsxToDrive;
}
