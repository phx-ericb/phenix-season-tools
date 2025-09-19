/** Liste unique des entraîneurs (rôles agrégés) + statuts photo/casier */
function getEntraineursAggreges() {
  var seasonId = getSeasonId_();
  var ss = SpreadsheetApp.openById(seasonId);

  var seasonYear = Number(readParam_(ss, 'SEASON_YEAR') || new Date().getFullYear());
  var invalidFrom = (readParam_(ss, 'PHOTO_INVALID_FROM_MMDD') || '04-01').trim();
  var cutoffNextJan1 = (seasonYear + 1) + '-01-01';
  var dueDate = seasonYear + '-' + invalidFrom;
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var shR = ss.getSheetByName('ENTRAINEURS_ROLES');
  if (!shR || shR.getLastRow() < 2) return [];
  var R = shR.getDataRange().getValues();
  var h = R[0];
  var ciPass = h.indexOf('Passeport'),
      ciRole = h.indexOf('Role'),
      ciCat  = h.indexOf('Categorie'),
      ciEq   = h.indexOf('Equipe');

  var grouped = {}; // pass -> {Passeport, Roles[]}
  for (var i=1;i<R.length;i++){
    var p = normalizePassportPlain8_(R[i][ciPass]);
    if (!p) continue;
    var piece = [
      String(R[i][ciRole]||'').trim(),
      String(R[i][ciCat]||'').trim(),
      String(R[i][ciEq]||'').trim()
    ].filter(Boolean).join(' | ');
    (grouped[p] = grouped[p] || {Passeport:p, Roles:[]}).Roles.push(piece);
  }

  // jointure MEMBRES_GLOBAL
  var shMG = ss.getSheetByName(readParam_(ss,'SHEET_MEMBRES_GLOBAL') || 'MEMBRES_GLOBAL');
  var V = shMG ? shMG.getDataRange().getValues() : [];
  var H = V[0] || [];
  var cPass = H.indexOf('Passeport'),
      cNom  = H.indexOf('Nom'),
      cPre  = H.indexOf('Prenom'),
      cDOB  = H.indexOf('DateNaissance'),
      cGen  = H.indexOf('Genre'),
      cPh   = H.indexOf('PhotoExpireLe'),
      cCas  = H.indexOf('CasierExpiré');

  var out = [];
  Object.keys(grouped).forEach(function(p){
    var nom='', prenom='', dob='', genre='', photo='', cas=0;

    // lookup linéaire (si volume ↑ on optimisera en map)
    for (var r=1; r<V.length; r++){
      if (normalizePassportPlain8_(V[r][cPass]) === p) {
        nom   = String(V[r][cNom]||''); 
        prenom= String(V[r][cPre]||'');
        dob   = String(V[r][cDOB]||'');
        genre = String(V[r][cGen]||'');
        photo = String(V[r][cPh]||'');
        cas   = Number(V[r][cCas]||0);
        break;
      }
    }

    var invalide = (photo && photo < cutoffNextJan1) ? 1 : 0;
    var statutPhoto = 'OK';
    if (invalide) {
      statutPhoto = (today >= dueDate) ? 'ÉCHUE' : ('À RENOUVELER dès ' + dueDate);
    }
    var statutCasier = (cas === 1) ? 'EXPIRÉ' : 'OK';

    out.push({
      Passeport: p,
      Prenom: prenom, Nom: nom, DateNaissance: dob, Genre: genre,
      PhotoExpireLe: photo, StatutPhoto: statutPhoto,
      Casier: statutCasier,
      Roles: grouped[p].Roles
    });
  });

  out.sort(function(a,b){
    var A=(a.Nom||'')+(a.Prenom||''), B=(b.Nom||'')+(b.Prenom||'');
    return A.localeCompare(B,'fr');
  });
  return out;
}
