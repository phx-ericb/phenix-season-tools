/**
 * retro_members.gs — v0.7
 * Export "Rétro - Membres" compatible avec l’ancien format, mais:
 *  - lit les feuilles finals v0.7 (SHEETS.INSCRIPTIONS / SHEETS.ARTICLES)
 *  - respecte CANCELLED/EXCLUDE_FROM_EXPORT
 *  - ignore certains frais (RETRO_IGNORE_FEES_CSV)
 *  - détecte Soccer Adapté (inscriptions ET articles) => "Adapté"=1
 *  - calcule CDP via MAPPINGS (ExclusiveGroup="CDP_ENTRAINEMENT") + fallback par textes
 *  - détecte "camp sélection U13" (paramétrable)
 *  - "Muté" via une feuille de mutation configurable
 *  - (optionnel) colonne Photo selon date d’expiration + cutoff paramétrés
 *
 * Fonctions exposées:
 *  - buildRetroMembresRows(seasonSheetId) -> {header:[], rows:[[]], nbCols:int}
 *  - writeRetroMembresSheet(seasonSheetId) -> écrit "Rétro - Membres"
 *  - exportRetroMembresXlsxToDrive(seasonSheetId) -> XLSX dans dossier paramétré
 */

/* ===================== Param keys (nouvelles) ===================== */
if (typeof PARAM_KEYS === 'undefined') { var PARAM_KEYS = {}; }
PARAM_KEYS.RETRO_IGNORE_FEES_CSV            = PARAM_KEYS.RETRO_IGNORE_FEES_CSV            || 'RETRO_IGNORE_FEES_CSV';
PARAM_KEYS.RETRO_ADAPTE_KEYWORDS            = PARAM_KEYS.RETRO_ADAPTE_KEYWORDS            || 'RETRO_ADAPTE_KEYWORDS';
PARAM_KEYS.RETRO_CAMP_KEYWORDS              = PARAM_KEYS.RETRO_CAMP_KEYWORDS              || 'RETRO_CAMP_KEYWORDS';
PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID          = PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID          || 'RETRO_EXPORTS_FOLDER_ID';
PARAM_KEYS.RETRO_MUTATION_SHEET             = PARAM_KEYS.RETRO_MUTATION_SHEET             || 'RETRO_MUTATION_SHEET';

PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL          = PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL          || 'RETRO_PHOTO_INCLUDE_COL';     // TRUE/FALSE
PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL           = PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL           || 'RETRO_PHOTO_EXPIRY_COL';      // header côté INSCRIPTIONS
PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE        = PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE        || 'RETRO_PHOTO_WARN_ABS_DATE';   // YYYY-MM-DD (optionnel)
PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD     = PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD     || 'RETRO_PHOTO_WARN_BEFORE_MMDD';// ex "03-01" (1er mars)
PARAM_KEYS.RETRO_EXCLUDE_MEMBER_IF_ONLY_IGN = PARAM_KEYS.RETRO_EXCLUDE_MEMBER_IF_ONLY_IGN || 'RETRO_EXCLUDE_MEMBER_IF_ONLY_IGN';
PARAM_KEYS.RETRO_RULES_JSON                 = PARAM_KEYS.RETRO_RULES_JSON                 || 'RETRO_RULES_JSON';

/* ===================== Helpers bas niveau ===================== */

// ——— Birth year robuste: gère Date, ISO, FR, tout ce qui contient un yyyy ———
function _extractBirthYearLoose_(v){
  if (!v) return 0;
  if (v instanceof Date) return v.getFullYear();
  var s = String(v).trim();
  // Cherche un groupe de 4 chiffres entre 1900 et 2099
  var m = s.match(/(19\d{2}|20\d{2})/);
  if (m) {
    var y = parseInt(m[1], 10);
    if (y >= 1900 && y <= 2099) return y;
  }
  return 0;
}

function _nrm_(s){
  s = String(s == null ? '' : s);
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g,''); } catch(e){}
  return s;
}
function _nrmLower_(s){ return _nrm_(s).toLowerCase().trim(); }
function _csvEsc_(v){ v = v==null?'':String(v).replace(/"/g,'""'); return /[",\n;]/.test(v)?('"'+v+'"'):v; }

function _isActiveRow_(r){
  var can = String(r[CONTROL_COLS.CANCELLED]||'').toLowerCase()==='true';
  var exc = String(r[CONTROL_COLS.EXCLUDE_FROM_EXPORT]||'').toLowerCase()==='true';
  var st  = (r["Statut de l'inscription"] || r['Statut'] || '').toString().toLowerCase();
  return !can && !exc && st !== 'annulé' && st !== 'annule' && st !== 'cancelled';
}
function _feeIgnored_(name, patternsCsv){
  var raw = _nrmLower_(name||'');
  if (!raw) return false;
  var pats = String(patternsCsv||'').split(',').map(function(x){return _nrmLower_(x);}).filter(Boolean);
  for (var i=0;i<pats.length;i++){ if (raw.indexOf(pats[i]) !== -1) return true; }
  return false;
}
function _containsAny_(raw, csv){
  var s = _nrmLower_(raw||'');
  return String(csv||'').split(',').map(function(x){return _nrmLower_(x);}).filter(Boolean).some(function(p){ return s.indexOf(p)!==-1; });
}
function _safeDate_(v){
  if (!v) return null;
  try { return (v instanceof Date) ? v : new Date(v); } catch(e){ return null; }
}
function _yyyy_mm_dd_(d){
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/* ===================== Catalogue ARTICLES (mêmes règles que rules.js) ===================== */
function _loadArticlesCatalog_(ss) {
  var sh = ss.getSheetByName(SHEETS.MAPPINGS);
  var items = [];
  if (sh) {
    var last = sh.getLastRow();
    if (last > 1) {
      var data = sh.getRange(1,1,last,Math.max(7, sh.getLastColumn())).getValues();
      for (var i=0;i<data.length;i++){
        if (String(data[i][0]).toUpperCase().trim() === 'ARTICLES') {
          var header = (i+1 < data.length) ? data[i+1] : null;
          if (!header) break;
          var hIdx = {}; header.forEach(function(h,j){ hIdx[String(h).trim()] = j; });
          for (var r=i+2; r<data.length; r++){
            var row = data[r];
            var code = (row[hIdx['Code']]||'').toString().trim();
            var alias = (row[hIdx['AliasContains']]||'').toString().trim();
            var excl = (row[hIdx['ExclusiveGroup']]||'').toString().trim();
            if (!code && !alias) break;
            items.push({ Code: code, AliasContains: alias, ExclusiveGroup: excl });
          }
          break;
        }
      }
    }
  }
  function match_(raw){
    var s=_nrmLower_(raw||'');
    for (var i=0;i<items.length;i++){
      var a=_nrmLower_(items[i].AliasContains||'');
      if (!a) continue;
      if (s.indexOf(a)!==-1) return items[i];
    }
    return null;
  }
  return { items: items, match: match_ };
}

/* ===================== Lecture "mutation" ===================== */
function _loadMutationsSet_(ss){
  var sheetName = readParam_(ss, PARAM_KEYS.RETRO_MUTATION_SHEET) || 'Mutation';
  var sh = ss.getSheetByName(sheetName);
  var set = {};
  if (!sh || sh.getLastRow()<2) return set;
  var vals = sh.getRange(2,1, sh.getLastRow()-1, 1).getDisplayValues();
  for (var i=0;i<vals.length;i++){
    var p = (vals[i][0]||'').toString().trim();
    if (p) set[p]=true;
  }
  return set;
}

/* ===================== PHOTO logic (optionnel) ===================== */
function _computePhotoCell_(ss, row){
  var include = (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) || 'FALSE').toUpperCase()==='TRUE';
  if (!include) return ''; // reste 100% compatible par défaut

  var col = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_EXPIRY_COL) || '';
  if (!col || !(col in row)) return '';

  var exp = _safeDate_(row[col]);
  if (!exp) return 'Aucune photo';
  var today = new Date();
  if (exp < today) return 'Expirée ('+_yyyy_mm_dd_(exp)+')';

  // alerte "Expire bientôt"
  var absWarn  = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_ABS_DATE) || '';
  var warnMmDd = readParam_(ss, PARAM_KEYS.RETRO_PHOTO_WARN_BEFORE_MMDD) || '03-01';

  var saisonLabel = readParam_(ss, 'SEASON_LABEL') || (row['Saison']||'');
  var seasonYear = parseSeasonYear_(saisonLabel);

  if (absWarn) {
    var abs = _safeDate_(absWarn);
    if (abs && exp <= abs) return 'Expire bientôt ('+_yyyy_mm_dd_(exp)+')';
  }
  if (exp.getFullYear() === seasonYear) {
    // règle: expiration durant l’année => changement avant la saison estivale
    return 'Expire bientôt ('+_yyyy_mm_dd_(exp)+')';
  }
  return _yyyy_mm_dd_(exp); // simple date
}

/* ===================== Retro Rules — loader + exécution ===================== */

/** Charge la config de règles:
 *  1) PARAMS.RETRO_RULES_JSON (JSON inline)
 *  2) onglet feuille nommé "RETRO_RULES_JSON" (JSON dans A1 ou multi-lignes concaténées)
 *  (fallback) dérive depuis d'autres PARAMS si rien n’est trouvé
 */
function loadRetroRules_(ss){
  // 1) PARAM direct
  var raw = readParam_(ss, PARAM_KEYS.RETRO_RULES_JSON) || '';
  if (!raw) {
    // 2) Feuille "RETRO_RULES_JSON"
    var shJson = ss.getSheetByName('RETRO_RULES_JSON');
    if (shJson && shJson.getLastRow() >= 1 && shJson.getLastColumn() >= 1) {
      var vals = shJson.getDataRange().getDisplayValues();
      // a) si A1 contient tout le JSON, on le prend; sinon on concatène toutes les cellules avec des retours
      var pieces = [];
      for (var i=0;i<vals.length;i++){
        for (var j=0;j<vals[i].length;j++){
          var cell = vals[i][j];
          if (cell != null && String(cell).trim() !== '') pieces.push(String(cell));
        }
      }
      raw = pieces.join('\n');
      appendImportLog_(ss, 'RETRO_RULES_JSON_SHEET_READ', 'chars='+raw.length);
    }
  }
  if (raw) {
    try {
      var arr = JSON.parse(raw);
      return (Array.isArray(arr) ? arr : []);
    } catch(e) {
      appendImportLog_(ss, 'RETRO_RULES_JSON_PARSE_FAIL', String(e));
    }
  }

  // Fallback: déduire depuis tes PARAMS actuels
  var ignoreCsv = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV) || 'senior,u-sé,adulte,ligue';
  var adapteCsv = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || 'adapté,adapte,adapte';
  var campCsv   = readParam_(ss, PARAM_KEYS.RETRO_CAMP_KEYWORDS)   || 'camp de sélection u13,camp selection u13,camp u13';
  var photoOn   = (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL)||'FALSE').toUpperCase()==='TRUE';
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
    // CDP via catalogue: 2 séances avant 1 séance
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
  return rules;
}

function _r_normLower_(s){ return _nrmLower_(s); }
function _r_getFieldText_(row, field){
  if (!field) return '';
  var v = row[field];
  if (v == null) return '';
  return String(v);
}
function _r_matchWhen_(when, row, feeName, catalogItem){
  if (!when) return true;
  if (when.field && when.contains_any) {
    var txt = _r_getFieldText_(row, when.field);
    var low = _r_normLower_(txt);
    var arr = [].concat(when.contains_any||[]).map(_r_normLower_).filter(Boolean);
    if (!arr.some(function(s){ return low.indexOf(s)!==-1; })) return false;
  }
  if (when.text_contains_any) {
    var low = _r_normLower_(feeName||'');
    var arr = [].concat(when.text_contains_any||[]).map(_r_normLower_).filter(Boolean);
    if (!arr.some(function(s){ return low.indexOf(s)!==-1; })) return false;
  }
  if (when.catalog_exclusive_group) {
    if (!catalogItem) return false;
    if (String(catalogItem.ExclusiveGroup||'') !== String(when.catalog_exclusive_group)) return false;
  }
  return true;
}

/** Applique toutes les règles "row" sur une ligne source (inscriptions/articles).
 *  Retourne { skip:true } si l’action "ignore_row" a été tirée.
 *  Peut modifier l’objet member (adapte/cdp/camp…).
 */
function applyRetroRowRules_(rules, scope, row, member, ctx){
  var feeName = (row['Nom du frais'] || row['Frais'] || row['Produit'] || '');
  var item = ctx.catalog.match ? ctx.catalog.match(feeName) : null;
  var skip = false;

  rules.forEach(function(rule){
    if (!rule || !rule.enabled) return;
    if (!(rule.scope==='both' || rule.scope===scope)) return;
    if (!_r_matchWhen_(rule.when, row, feeName, item)) return;

    var a = rule.action || {};
    if (a.type === 'ignore_row') {
      skip = true;
    } else if (a.type === 'set_member_field') {
      member[a.field] = a.value;
    } else if (a.type === 'set_member_field_max') {
      var cur = member[a.field];
      if (cur == null || cur === '') member[a.field] = a.value;
      else {
        var n1 = parseFloat(cur), n2 = parseFloat(a.value);
        if (!isNaN(n1) && !isNaN(n2)) member[a.field] = Math.max(n1, n2);
        else member[a.field] = a.value; // fallback
      }
    }
  });

  return { skip: skip };
}

/** Applique les règles "member" (post-agrégation).
 *  Peut écrire member.__photoStr pour que la feuille affiche la valeur.
 */
function applyRetroMemberRules_(rules, member, ctx){
  rules.forEach(function(rule){
    if (!rule || !rule.enabled) return;
    if (rule.scope !== 'member') return;
    var a = rule.action || {};
    if (a.type === 'compute_photo') {
      var tmpRow = {};
      tmpRow[a.expiry_col] = (member.__rowForPhoto && member.__rowForPhoto[a.expiry_col]) ? member.__rowForPhoto[a.expiry_col] : '';
      member.__photoStr = _computePhotoCell_(ctx.ss, tmpRow);
    }
  });
}

/* ===================== Construire les lignes "Rétro - Membres" ===================== */
function buildRetroMembresRows(seasonSheetId){
  var ss   = getSeasonSpreadsheet_(seasonSheetId);
  var insc = readSheetAsObjects_(ss.getId(), SHEETS.INSCRIPTIONS);
  var art  = readSheetAsObjects_(ss.getId(), SHEETS.ARTICLES);

  var catalog = _loadArticlesCatalog_(ss);
  var mutated = _loadMutationsSet_(ss);
  var rules   = loadRetroRules_(ss);
  var ctx     = { ss:ss, catalog:catalog };

  var ignoreFees = readParam_(ss, PARAM_KEYS.RETRO_IGNORE_FEES_CSV) || 'senior,u-se,adulte,ligue';
  var adapteKeys = readParam_(ss, PARAM_KEYS.RETRO_ADAPTE_KEYWORDS) || 'adapté,adapte,adapte';
  var campKeys   = readParam_(ss, PARAM_KEYS.RETRO_CAMP_KEYWORDS)   || 'camp de selection u13,camp selection u13,camp u13';
  var exclOnlyIgn = (readParam_(ss, PARAM_KEYS.RETRO_EXCLUDE_MEMBER_IF_ONLY_IGN) || 'TRUE').toUpperCase()==='TRUE';

  // map membres (clé = passeport)
  var members = {}; // { passeport, nom, prenom, dateNaissance, genre, emails[], facture, adapte, cdp, camp, inscription:boolean }

  function ensureMember_(p, seed){
    var k = String(p||'').trim(); if (!k) return null;
    if (!members[k]) {
      members[k] = {
        passeport: k,
        nom: '',
        prenom: '',
        dateNaissance: '',
        genre: '',
        emails: [],
        facture: '',
        adapte: undefined,
        cdp: undefined,
        camp: undefined,
        inscription: false
      };
    }
    if (seed) {
      var m = members[k];
      if (!m.nom && seed.nom) m.nom = seed.nom;
      if (!m.prenom && seed.prenom) m.prenom = seed.prenom;
      if (!m.dateNaissance && seed.dateNaissance) m.dateNaissance = seed.dateNaissance;
      if (!m.genre && seed.genre) m.genre = seed.genre;
    }
    return members[k];
  }

  // --- INSCRIPTIONS (actives)
  var inscAct = insc.rows.filter(_isActiveRow_);
  inscAct.forEach(function(r){
    var pass = r['Passeport #']; if (!pass) return;

    var feeName = r['Nom du frais'] || r['Frais'] || r['Produit'] || '';
    if (_feeIgnored_(feeName, ignoreFees)) return; // ignore cette ligne

    var m = ensureMember_(pass, {
      nom: r['Nom'] || '',
      prenom: r['Prénom'] || r['Prenom'] || '',
      dateNaissance: r['Date de naissance'] || r['Naissance'] || '',
      genre: (r['Identité de genre']||'').toString().trim().toUpperCase().charAt(0)
    });
    if (!m) return;
    m.inscription = true;

    // règles "row" (scope inscriptions)
    var rr = applyRetroRowRules_(rules, 'inscriptions', r, m, ctx);
    if (rr.skip) return;

    // e-mails
    var emails = (typeof collectEmailsFromRow_ === 'function')
      ? collectEmailsFromRow_(r, readParam_(ss, 'TO_FIELDS_INSCRIPTIONS') || 'Courriel,Parent 1 - Courriel,Parent 2 - Courriel')
      : [r['Courriel'], r['Parent 1 - Courriel'], r['Parent 2 - Courriel']].filter(Boolean).join(',');
    if (emails) {
      emails.split(',').forEach(function(e){ e=e.trim(); if (e && m.emails.indexOf(e)===-1) m.emails.push(e); });
    }

    // facture (col paramétrable, fallback "Date de facture")
    var facCol = (insc.headers.indexOf('Date de facture')>=0) ? 'Date de facture' : null;
    if (facCol && r[facCol]) m.facture = r[facCol];

    // Adapté si mot-clef dans le frais (en plus des règles, pour compat immédiate)
    if (_containsAny_(feeName, adapteKeys)) m.adapte = 1;
  });

  // --- ARTICLES (actifs)
  var artAct = art.rows.filter(_isActiveRow_);
  artAct.forEach(function(a){
    var pass = a['Passeport #']; if (!pass) return;

    var feeName = a['Nom du frais'] || a['Frais'] || a['Produit'] || '';
    if (_feeIgnored_(feeName, ignoreFees)) return;

    var m = ensureMember_(pass, {
      nom: a['Nom'] || '',
      prenom: a['Prénom'] || a['Prenom'] || '',
      dateNaissance: a['Date de naissance'] || a['Naissance'] || '',
      genre: (a['Identité de genre']||'').toString().trim().toUpperCase().charAt(0)
    });
    if (!m) return;

    // règles "row" (scope articles)
    var ra = applyRetroRowRules_(rules, 'articles', a, m, ctx);
    if (ra.skip) return;

    // Adapté aussi via articles (fallback rapide)
    if (_containsAny_(feeName, adapteKeys)) m.adapte = 1;

    // CDP: via MAPPINGS (ExclusiveGroup='CDP_ENTRAINEMENT') -> 1 ou 2; sinon fallback texte
      // --- CDP: détection robuste (sans dépendre du "catalogue ARTICLES")
    // Normalise texte: accents enlevés + minuscules
    var s = _nrmLower_(feeName);

    // On considère CDP s'il y a "cdp" quelque part
    var isCdp = s.indexOf('cdp') !== -1;

    if (isCdp) {
      // 2 entraînement(s)
      if (/\b2\b/.test(s) || /2\s*entrainement/.test(s) || /2\s*entrainements/.test(s)) {
        m.cdp = 2;
      }
      // 1 entraînement
      else if (/\b1\b/.test(s) || /1\s*entrainement/.test(s) || /1\s*entrainements/.test(s)) {
        m.cdp = Math.max(m.cdp || 0, 1);
      }
      // Si CDP sans chiffre explicite → au moins 1
      else {
        m.cdp = Math.max(m.cdp || 0, 1);
      }
    }


    // Camp (clé paramétrable)
    if (_containsAny_(feeName, campKeys)) m.camp = 'Oui';
  });

 // U9–U12 → défaut CDP=0 si non défini et non "Adapté"
var seasonYear = parseSeasonYear_(readParam_(ss, 'SEASON_LABEL') || '');
var currentYear = seasonYear || (new Date()).getFullYear();

Object.keys(members).forEach(function(k){
  var m = members[k];

  // année de naissance robuste
  var by = _extractBirthYearLoose_(m.dateNaissance);
  var age = by ? (currentYear - by) : 0;

  // adapt? considère 1 / "1" / TRUE (string)
  var isAdapt = (m.adapte === 1) ||
                (m.adapte === '1') ||
                (String(m.adapte||'').toLowerCase() === 'true');

  if (age >= 9 && age <= 12 && !isAdapt) {
    if (typeof m.cdp === 'undefined' || m.cdp === null || m.cdp === '') {
      m.cdp = 0;
    }
  }
});


  // si demandé: exclure les membres qui n’ont QUE des frais ignorés
  if (exclOnlyIgn) {
    // On a déjà filtré ligne par ligne → les membres présents ont ≥ 1 ligne utile.
  }

  // Préparer index INSCRIPTIONS par passeport (pour la photo)
  var indexByPassport = {};
  inscAct.forEach(function(r){ indexByPassport[String(r['Passeport #']||'').trim()] = r; });

  // Appliquer les règles "member" (photo, etc.) et assigner la row photo
  Object.keys(members).forEach(function(k){
    var m = members[k];
    m.__rowForPhoto = indexByPassport[m.passeport] || {};
    applyRetroMemberRules_(rules, m, ctx);
  });

  // ——— Construction des lignes (même entête/ordre que l’export existant) ———
  var HEADER = [
    "Identifiant unique", "Code", "Nom", "Prénom", "Date de naissance",
    "Genre(M pour Masculin ou F pour Féminin)", "Langue", "Courriels", "Adresse", "Ville",
    "Code Postal", "Domicile Téléphone", "Mobile Téléphone", "Travail Téléphone",
    "Parent 1 Nom", "Parent 1 Courriels", "Parent 1 Domicile Téléphone", "Parent 1 Mobile Téléphone", "Parent 1 Travail Téléphone",
    "Parent 2 Nom", "Parent 2 Courriels", "Parent 2 Domicile Téléphone", "Parent 2 Mobile Téléphone", "Parent 2 Travail Téléphonique",
    "Autre Nom", "Autre Courriels", "Autre Domicile Téléphone", "Autre Mobile Téléphone", "Autre Travail Téléphonique",
    "Position", "Établissement scolaire", "Fiche d'employé", "Specimen Chèque", "Filtration Policière", "Respect et Sport",
    "S3", "S7", "Théorie A+B", "Diplôme C", "Adapté", "CDP", "Euroclass", "Camp", "Muté", "École", "InscritE25"
  ];
  var includePhoto = (readParam_(ss, PARAM_KEYS.RETRO_PHOTO_INCLUDE_COL) || 'FALSE').toUpperCase()==='TRUE';
  var header = HEADER.slice();
  if (includePhoto) header.push('Photo'); // colonne optionnelle à la fin

  var rows = [];
  Object.keys(members).forEach(function(k){
    var m = members[k];
    var row = new Array(header.length); for (var i=0;i<row.length;i++) row[i]='';

row[0] = (typeof normalizePassportToText8_ === 'function')
  ? normalizePassportToText8_(m.passeport)
  : String(m.passeport || '');
    // row[1] Code -> vide
    row[2]  = m.nom;
    row[3]  = m.prenom;
    row[4]  = m.dateNaissance;
    row[5]  = m.genre;
    row[7]  = (m.emails && m.emails.length) ? m.emails.join('; ') : '';

    // Adapté, CDP, Camp
    row[39] = (typeof m.adapte !== 'undefined' ? m.adapte : '');
    row[40] = (typeof m.cdp    !== 'undefined' ? m.cdp    : '');
    row[42] = (typeof m.camp   !== 'undefined' ? m.camp   : '');

    // Muté
    row[43] = mutated[m.passeport] ? 'Oui' : 'Non';

    // Photo (optionnel)
    if (includePhoto) {
      row[header.length-1] = (m.__photoStr != null) ? m.__photoStr : _computePhotoCell_(ss, m.__rowForPhoto || {});
    }

    rows.push(row);
  });

  return { header: header, rows: rows, nbCols: header.length };
}

/* ===================== Écriture feuille & export XLSX ===================== */

function writeRetroMembresSheet(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);
  var res = buildRetroMembresRows(seasonSheetId);

  var sh = ss.getSheetByName('Rétro - Membres') || ss.insertSheet('Rétro - Membres');
  // reset
  sh.clearContents();
  sh.getRange(1,1,1,res.header.length).setValues([res.header]);
  if (res.rows.length) {
    sh.getRange(2,1,res.rows.length,res.nbCols).setValues(res.rows);
    sh.autoResizeColumns(1, res.nbCols);
    if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,1).setNumberFormat('@'); // passeport en texte
  }
  appendImportLog_(ss, 'RETRO_MEMBRES_SHEET_OK', 'rows='+res.rows.length);
  return res.rows.length;
}

/** Export XLSX rapide: évite copyTo, écrit directement dans un classeur temporaire minimal. */
function exportRetroMembresXlsxToDrive(seasonSheetId) {
  var ss = getSeasonSpreadsheet_(seasonSheetId);

  // 1) Construit les données (sans écrire dans "Rétro - Membres")
  var res = buildRetroMembresRows(seasonSheetId); // {header, rows, nbCols}

  // 2) Classeur temporaire minimal
  var temp = SpreadsheetApp.create('Export temporaire - Retro Membres');
  var tmp = temp.getSheets()[0];
  tmp.setName('Export');

  // 3) Ecriture en un seul setValues (header + data)
  var all = [res.header].concat(res.rows);
  if (all.length) {
    tmp.getRange(1, 1, all.length, res.nbCols).setValues(all);
    // Passeport texte (ceinture+bretelles: on garde aussi l’apostrophe côté build si dispo)
    if (all.length > 1) {
      tmp.getRange(2, 1, all.length - 1, 1).setNumberFormat('@');
    }
  }
  SpreadsheetApp.flush();

  // 4) Export XLSX (une seule feuille, pas de formatting lourd)
  var url = 'https://docs.google.com/spreadsheets/d/' + temp.getId() + '/export?format=xlsx';
  var blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob();
  var fileName = 'Export_Retro_Membres_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd-HHmm') + '.xlsx';
  blob.setName(fileName);

  var folderId = readParam_(ss, PARAM_KEYS.RETRO_EXPORTS_FOLDER_ID) || '';
  var dest = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  var xlsx = dest.createFile(blob);

  // 5) Nettoyage
  DriveApp.getFileById(temp.getId()).setTrashed(true);

  appendImportLog_(ss, 'RETRO_MEMBRES_XLSX_OK_FAST', xlsx.getName() + ' -> ' + dest.getName() + ' (rows=' + res.rows.length + ')');
  return { fileId: xlsx.getId(), name: xlsx.getName(), rows: res.rows.length };
}


/* ========== Exposition facultative via Library ========== */
if (typeof Library !== 'undefined') {
  Library.buildRetroMembresRows = buildRetroMembresRows;
  Library.writeRetroMembresSheet = writeRetroMembresSheet;
  Library.exportRetroMembresXlsxToDrive = exportRetroMembresXlsxToDrive;
}
