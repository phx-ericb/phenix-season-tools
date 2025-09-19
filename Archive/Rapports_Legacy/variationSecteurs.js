/**
 * Détermine le "jour actuel" pour 2025 en lisant l’onglet "2025".
 * On considère que la première inscription (colonne G) correspond au jour 1.
 * La fonction calcule le nombre de jours entre la première et la dernière date.
 */
function determinerJourActuel2025() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet2025 = ss.getSheetByName("2025");
  if (!sheet2025) {
    Logger.log("Onglet '2025' introuvable.");
    return 0;
  }
  var data = sheet2025.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("Pas de données dans l'onglet 2025.");
    return 0;
  }
  
  var dates = [];
  // Colonne G a l'index 6
  for (var i = 1; i < data.length; i++) {
    var dateVal = data[i][6];
    if (!dateVal) continue;
    var d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (!isNaN(d.getTime())) {
      dates.push(d);
    }
  }
  if (dates.length === 0) {
    Logger.log("Aucune date valide trouvée dans l'onglet 2025.");
    return 0;
  }
  dates.sort(function(a, b) { return a - b; });
  var premierDate = dates[0];
  var derniereDate = dates[dates.length - 1];
  // Calcul du nombre de jours écoulés (ajouter 1 pour que le premier jour soit 1)
  var diffMs = derniereDate - premierDate;
  var diffDays = Math.floor(diffMs / (1000 * 3600 * 24)) + 1;
  Logger.log("Premier jour: " + premierDate + ", Dernier jour: " + derniereDate + ", Jour actuel: " + diffDays);
  return diffDays;
}

/**
 * Génère une synthèse de variation pour la saison Ete.
 * Pour chaque secteur, elle affiche :
 *   - Le cumul d’inscriptions 2025 au "jour actuel" (déterminé via l’onglet "2025")
 *   - La variation sur 15 jours (différence entre la valeur au jour actuel et celle 15 jours plus tôt, dans 2025)
 *   - La variation sur 7 jours (même principe pour 7 jours)
 *   - L’écart par rapport aux mêmes jours pour 2023 et 2024
 *   - La part (%) de chaque secteur dans le total 2025
 *
 * On suppose que pour chaque secteur, il existe un onglet nommé "Progression Ete <secteur>"
 * contenant un tableau avec les colonnes : ["Jour", "2023", "2024", "2025"] (avec la première ligne en en-tête).
 */
function genererSyntheseVariationEte() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var jourAnalyse = determinerJourActuel2025();
  if (jourAnalyse <= 0) {
    SpreadsheetApp.getUi().alert("Impossible de déterminer le jour actuel à partir de l'onglet 2025.");
    return;
  }
  
  // Liste des secteurs à traiter – adaptez cette liste selon vos besoins
  var secteurs = ["U4-U8", "U9-U12", "U13-U18", "Adapté", "Senior", "Élite"];
  
  // Créer ou réinitialiser la feuille de synthèse
  var syntheseSheetName = "Synthese Variation Ete";
  var syntheseSheet = ss.getSheetByName(syntheseSheetName) || ss.insertSheet(syntheseSheetName);
  syntheseSheet.clear();
  
  // En-tête du tableau de synthèse
  var header = ["Secteur", "Jour Analyse", "Inscriptions 2025", "Variation 15j", "Variation 7j", "Écart vs 2023", "Écart vs 2024", "% du Total"];
  var output = [header];
  
  var total2025Global = 0;
  var resultParSecteur = {};
  
  secteurs.forEach(function(secteur) {
    var sheetName = "Progression Ete " + secteur;
    var sheetProg = ss.getSheetByName(sheetName);
    if (!sheetProg) {
      Logger.log("Feuille " + sheetName + " introuvable.");
      return;
    }
    var data = sheetProg.getDataRange().getValues();
    if (data.length < 2) return; // Pas de données
    
    // On suppose que la première ligne est l'en-tête: ["Jour", "2023", "2024", "2025"]
    // Chercher la ligne dont la première colonne vaut le jour d'analyse.
    var targetRow = null;
    for (var i = 1; i < data.length; i++) {
      if (parseInt(data[i][0], 10) === jourAnalyse) {
        targetRow = data[i];
        break;
      }
    }
    // Si le jour d'analyse n'est pas trouvé, on prend la dernière ligne disponible pour ce secteur.
    if (!targetRow) {
      targetRow = data[data.length - 1];
      // On utilisera le "jour" disponible dans ce secteur
      var jourAnalyseLocal = parseInt(targetRow[0], 10);
      Logger.log("Pour le secteur " + secteur + ", jour d'analyse " + jourAnalyse + " n'existe pas, on prend " + jourAnalyseLocal);
      jourAnalyse = jourAnalyseLocal;
    }
    
    // Valeurs de 2025, 2023, 2024 à la ligne cible
    var inscriptions2025 = parseInt(targetRow[3], 10) || 0;
    var inscriptions2023 = parseInt(targetRow[1], 10) || 0;
    var inscriptions2024 = parseInt(targetRow[2], 10) || 0;
    
    // Variation sur 15 jours
    var variation15 = 0;
    if (jourAnalyse > 15) {
      var targetRow15 = null;
      for (var i = 1; i < data.length; i++) {
        if (parseInt(data[i][0], 10) === (jourAnalyse - 15)) {
          targetRow15 = data[i];
          break;
        }
      }
      var inscriptions2025_15 = targetRow15 ? (parseInt(targetRow15[3], 10) || 0) : 0;
      variation15 = inscriptions2025 - inscriptions2025_15;
    }
    
    // Variation sur 7 jours
    var variation7 = 0;
    if (jourAnalyse > 7) {
      var targetRow7 = null;
      for (var i = 1; i < data.length; i++) {
        if (parseInt(data[i][0], 10) === (jourAnalyse - 7)) {
          targetRow7 = data[i];
          break;
        }
      }
      var inscriptions2025_7 = targetRow7 ? (parseInt(targetRow7[3], 10) || 0) : 0;
      variation7 = inscriptions2025 - inscriptions2025_7;
    }
    
    var ecart2023 = inscriptions2025 - inscriptions2023;
    var ecart2024 = inscriptions2025 - inscriptions2024;
    
    resultParSecteur[secteur] = {
      inscriptions2025: inscriptions2025,
      variation15: variation15,
      variation7: variation7,
      ecart2023: ecart2023,
      ecart2024: ecart2024
    };
    total2025Global += inscriptions2025;
  });
  
  // Calculer le pourcentage de chaque secteur par rapport au total 2025
  for (var secteur in resultParSecteur) {
    var r = resultParSecteur[secteur];
    var pct = total2025Global ? (r.inscriptions2025 / total2025Global * 100).toFixed(1) + "%" : "0%";
    output.push([secteur, jourAnalyse, r.inscriptions2025, r.variation15, r.variation7, r.ecart2023, r.ecart2024, pct]);
  }
  
  // Écrire le tableau final dans la feuille de synthèse
  syntheseSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  syntheseSheet.autoResizeColumns(1, output[0].length);
  
  //SpreadsheetApp.getUi().alert("Synthèse de variation pour la saison Ete générée avec succès.");
}
