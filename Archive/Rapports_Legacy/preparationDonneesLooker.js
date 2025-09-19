/**
 * Génère les tables de données pour Looker Studio à partir de la synthèse
 * des inscriptions de type JOUEURS.
 * Pour les années 2019-2022, seules les totaux sont disponibles.
 * Pour 2023-2025, seules les inscriptions dont le secteur figure dans la table
 * de standardisation (type JOUEURS) sont prises en compte, avec répartition par genre.
 * 
 * Deux onglets sont générés :
 * - "Looker Global Joueurs" : agrégation globale par groupe (Total, Ete, Automne Hiver)
 * - "Looker Secteurs Joueurs" : agrégation par secteur pour les années 2023-2025
 */
function genererDonneesLookerJoueurs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Récupérer la synthèse des inscriptions
  var syntheseSheet = ss.getSheetByName("Synthèse Inscriptions Standardisées");
  if (!syntheseSheet) {
    SpreadsheetApp.getUi().alert("La feuille 'Synthèse Inscriptions Standardisées' est introuvable.");
    return;
  }
  var syntheseData = syntheseSheet.getDataRange().getValues();
  if (syntheseData.length < 2) {
    SpreadsheetApp.getUi().alert("Pas de données suffisantes dans la synthèse.");
    return;
  }
  
  // 2. Récupérer la table de standardisation pour connaître les secteurs de type JOUEURS
  var standardSheet = ss.getSheetByName("Standardisation");
  if (!standardSheet) {
    SpreadsheetApp.getUi().alert("L'onglet 'Standardisation' est introuvable.");
    return;
  }
  var standardData = standardSheet.getDataRange().getValues(); // En-tête en ligne 1
  var joueursSecteurs = {};
  for (var i = 1; i < standardData.length; i++) {
    var row = standardData[i];
    var type = row[2] ? row[2].toString().trim().toUpperCase() : "";
    if (type === "JOUEURS") {
      var secteur = row[3] ? row[3].toString().trim() : "";
      if (secteur !== "") {
        joueursSecteurs[secteur] = true;
      }
    }
  }
  
  // 3. Agréger les données de la synthèse.
  // On suppose que la synthèse possède au moins 6 colonnes :
  // A: Année, B: Saison, C: Secteur, D: Féminin, E: Masculin, F: Total
  // Optionnellement, la colonne G (index 6) contient le statut.
  var aggregatedGlobal = { "Total": {}, "Ete": {}, "Automne Hiver": {} };
  var aggregatedBySector = {};  // pour les inscriptions new (2023-2025)
  
  for (var i = 1; i < syntheseData.length; i++) {
    var row = syntheseData[i];
    // Si la ligne contient une colonne de statut (colonne G), l'ignorer si "Annulé"
    if (row.length >= 7) {
      var statut = row[6] ? row[6].toString().trim() : "";
      if (statut === "Annulé") continue;
    }
    
    var anneeStr = row[0];
    var annee = parseInt(anneeStr, 10);
    var saison = row[1];
    // Récupérer et normaliser le secteur
    var secteur = row[2] ? row[2].toString().trim() : "";
    if (secteur.toLowerCase().indexOf("u-sé") !== -1) {
      secteur = "Senior";
    }
    
    var feminin = parseInt(row[3], 10) || 0;
    var masculin = parseInt(row[4], 10) || 0;
    var total = parseInt(row[5], 10) || 0;
    
    // Normaliser la saison
    var saisonLower = saison ? saison.toString().toLowerCase() : "";
    var normalizedSeason = "";
    if (saisonLower.indexOf("été") !== -1 || saisonLower.indexOf("ete") !== -1) {
      normalizedSeason = "Ete";
    } else if (saisonLower.indexOf("automne") !== -1 || saisonLower.indexOf("hiver") !== -1) {
      normalizedSeason = "Automne Hiver";
    } else {
      normalizedSeason = saison;
    }
    
    var isNew = (annee >= 2023);
    
    // Agrégation globale : 
    if (!isNew || (isNew && joueursSecteurs[secteur])) {
      // Groupe "Total"
      if (!aggregatedGlobal["Total"][anneeStr]) {
        aggregatedGlobal["Total"][anneeStr] = { feminin: 0, masculin: 0, total: 0 };
      }
      aggregatedGlobal["Total"][anneeStr].feminin += feminin;
      aggregatedGlobal["Total"][anneeStr].masculin += masculin;
      aggregatedGlobal["Total"][anneeStr].total += total;
      
      // Groupe spécifique (Ete ou Automne Hiver)
      if (!aggregatedGlobal[normalizedSeason][anneeStr]) {
        aggregatedGlobal[normalizedSeason][anneeStr] = { feminin: 0, masculin: 0, total: 0 };
      }
      aggregatedGlobal[normalizedSeason][anneeStr].feminin += feminin;
      aggregatedGlobal[normalizedSeason][anneeStr].masculin += masculin;
      aggregatedGlobal[normalizedSeason][anneeStr].total += total;
    }
    
    // Agrégation par secteur : pour les inscriptions new (2023-2025)
    if (isNew && joueursSecteurs[secteur]) {
      if (!aggregatedBySector[secteur]) {
        aggregatedBySector[secteur] = { "Total": {}, "Ete": {}, "Automne Hiver": {} };
      }
      // Groupe "Total"
      if (!aggregatedBySector[secteur]["Total"][anneeStr]) {
        aggregatedBySector[secteur]["Total"][anneeStr] = { feminin: 0, masculin: 0, total: 0 };
      }
      aggregatedBySector[secteur]["Total"][anneeStr].feminin += feminin;
      aggregatedBySector[secteur]["Total"][anneeStr].masculin += masculin;
      aggregatedBySector[secteur]["Total"][anneeStr].total += total;
      
      // Groupe spécifique (Ete ou Automne Hiver)
      if (!aggregatedBySector[secteur][normalizedSeason][anneeStr]) {
        aggregatedBySector[secteur][normalizedSeason][anneeStr] = { feminin: 0, masculin: 0, total: 0 };
      }
      aggregatedBySector[secteur][normalizedSeason][anneeStr].feminin += feminin;
      aggregatedBySector[secteur][normalizedSeason][anneeStr].masculin += masculin;
      aggregatedBySector[secteur][normalizedSeason][anneeStr].total += total;
    }
  }
  
  // 4. Créer (ou réinitialiser) l'onglet "Looker Global Joueurs"
  var globalSheetName = "Looker Global Joueurs";
  var globalSheet = ss.getSheetByName(globalSheetName);
  if (!globalSheet) {
    globalSheet = ss.insertSheet(globalSheetName);
  } else {
    globalSheet.clear();
  }
  
  // Préparer la table globale : Colonnes : Groupe | Année | Féminin | Masculin | Total
  var globalOutput = [["Groupe", "Année", "Féminin", "Masculin", "Total"]];
  for (var group in aggregatedGlobal) {
    var groupData = aggregatedGlobal[group];
    var anneesKeys = Object.keys(groupData).sort(function(a, b) {
      return parseInt(a, 10) - parseInt(b, 10);
    });
    for (var j = 0; j < anneesKeys.length; j++) {
      var annee = anneesKeys[j];
      var stats = groupData[annee];
      globalOutput.push([group, annee, stats.feminin, stats.masculin, stats.total]);
    }
  }
  globalSheet.getRange(1, 1, globalOutput.length, globalOutput[0].length).setValues(globalOutput);
  globalSheet.autoResizeColumns(1, globalOutput[0].length);
  
  // 5. Créer (ou réinitialiser) l'onglet "Looker Secteurs Joueurs"
  var secteursSheetName = "Looker Secteurs Joueurs";
  var secteursSheet = ss.getSheetByName(secteursSheetName);
  if (!secteursSheet) {
    secteursSheet = ss.insertSheet(secteursSheetName);
  } else {
    secteursSheet.clear();
  }
  
  // Préparer la table par secteur : Colonnes : Secteur | Groupe | Année | Féminin | Masculin | Total
  var secteursOutput = [["Secteur", "Groupe", "Année", "Féminin", "Masculin", "Total"]];
  for (var secteur in aggregatedBySector) {
    var sectorData = aggregatedBySector[secteur];
    for (var group in sectorData) {
      var groupData = sectorData[group];
      var anneesKeys = Object.keys(groupData).sort(function(a, b) {
        return parseInt(a, 10) - parseInt(b, 10);
      });
      for (var j = 0; j < anneesKeys.length; j++) {
        var annee = anneesKeys[j];
        var stats = groupData[annee];
        secteursOutput.push([secteur, group, annee, stats.feminin, stats.masculin, stats.total]);
      }
    }
  }
  secteursSheet.getRange(1, 1, secteursOutput.length, secteursOutput[0].length).setValues(secteursOutput);
  secteursSheet.autoResizeColumns(1, secteursOutput[0].length);
  
  // SpreadsheetApp.getUi().alert("Les données pour Looker Studio ont été générées avec succès.");
}
