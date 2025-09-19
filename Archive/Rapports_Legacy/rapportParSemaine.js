function genererTableauxInscriptionsHebdo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Liste des années à traiter (feuilles existantes)
  var anneeOnglets = ["2019", "2020", "2019", "2021", "2022", "2023", "2024", "2025"];
  // Remplacer "2019" par "2019" dans l'exemple si besoin, ici on garde votre liste originale ou une liste correcte.
  // Pour cet exemple, on utilisera :
  anneeOnglets = ["2019", "2020", "2021", "2022", "2023", "2024", "2025"];
  
  // Les deux saisons à traiter
  var saisons = ["Été", "Automne-hiver"];
  
  // Structure pour stocker les inscriptions par date pour chaque année et saison.
  // seasonalCounts[year][saison][dateKey] = nombre d'inscriptions
  var seasonalCounts = {};
  anneeOnglets.forEach(function(year) {
    seasonalCounts[year] = {};
    saisons.forEach(function(season) {
      seasonalCounts[year][season] = {};
    });
  });
  
  // Lecture des feuilles et comptage par date (en fonction du format ancien ou nouveau)
  anneeOnglets.forEach(function(year) {
    var sheet = ss.getSheetByName(year);
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    // Ancien format (2019-2022) : 
    //   - Date en colonne A (index 0)
    //   - Total inscriptions en colonne B (index 1)
    //   - Saison en colonne C (index 2)
    // Nouveau format (2023 et après) : 
    //   - Chaque ligne représente 1 inscription
    //   - Date en colonne G (index 6)
    //   - Saison en colonne H (index 7)
    //   - Année en colonne I (index 8) qui doit correspondre à l'onglet courant
    //   - Statut en colonne J (index 9) : si "Annulé", l'inscription est ignorée.
    var isAncienFormat = parseInt(year, 10) < 2023;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dateFacture, count, saisonBrut;
      if (isAncienFormat) {
        dateFacture = row[0];
        count = parseInt(row[1], 10) || 0;
        saisonBrut = row[2] ? row[2].toString().trim() : "";
      } else {
        // Vérifier que la colonne "Année" correspond à l'année de l'onglet
        var anneeInscription = row[8] ? row[8].toString().trim() : "";
        if (anneeInscription !== year) continue;
        // Vérifier le statut : ignorer la ligne si "Annulé" (colonne J, index 9)
        var statut = row[9] ? row[9].toString().trim() : "";
        if (statut === "Annulé") continue;
        dateFacture = row[6];
        count = 1; // Chaque ligne correspond à une inscription individuelle
        saisonBrut = row[7] ? row[7].toString().trim() : "";
      }
      if (!dateFacture) continue;
      var dateObj = (dateFacture instanceof Date) ? dateFacture : new Date(dateFacture);
      if (isNaN(dateObj.getTime())) continue;
      
      // Formater la date en "yyyy-mm-dd"
      var yyyy = dateObj.getFullYear();
      var mm = ("0" + (dateObj.getMonth() + 1)).slice(-2);
      var dd = ("0" + dateObj.getDate()).slice(-2);
      var dateKey = yyyy + "-" + mm + "-" + dd;
      
      // Normaliser la saison (on ne traite que "Été" et "Automne-hiver")
      var saisonLower = saisonBrut.toLowerCase();
      var seasonNormalized = "";
      if (saisonLower.indexOf("été") !== -1 || saisonLower.indexOf("ete") !== -1) {
        seasonNormalized = "Été";
      } else if (saisonLower.indexOf("automne") !== -1 || saisonLower.indexOf("hiver") !== -1) {
        seasonNormalized = "Automne-hiver";
      } else {
        continue; // inscription ignorée si la saison n'est pas reconnue
      }
      
      if (!seasonalCounts[year][seasonNormalized][dateKey]) {
        seasonalCounts[year][seasonNormalized][dateKey] = 0;
      }
      seasonalCounts[year][seasonNormalized][dateKey] += count;
    }
  });
  
  // Regrouper par semaine pour chaque année et chaque saison.
  // On définit : jours 1 à 42 => semaine = Math.floor((jour-1)/7)+1, et les jours >=43 => semaine 7.
  var weeklyCounts = {};
  anneeOnglets.forEach(function(year) {
    weeklyCounts[year] = {};
    saisons.forEach(function(season) {
      weeklyCounts[year][season] = {1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0};
      var datesObj = seasonalCounts[year][season];
      var dateKeys = Object.keys(datesObj);
      if (dateKeys.length === 0) return;
      dateKeys.sort();
      dateKeys.forEach(function(dKey, index) {
        var jourRelatif = index + 1;
        var semaine = (jourRelatif <= 42) ? Math.floor((jourRelatif - 1) / 7) + 1 : 7;
        weeklyCounts[year][season][semaine] += datesObj[dKey];
      });
    });
  });
  
  // Calculer les totaux par semaine pour chaque année en additionnant les deux saisons.
  var weeklyCountsTotal = {};
  anneeOnglets.forEach(function(year) {
    weeklyCountsTotal[year] = {1:0, 2:0, 3:0, 4:0, 5:0, 6:0, 7:0};
    saisons.forEach(function(season) {
      if (!weeklyCounts[year][season]) return;
      for (var w = 1; w <= 7; w++) {
        weeklyCountsTotal[year][w] += weeklyCounts[year][season][w];
      }
    });
  });
  
  // Libellés des semaines
  var semaineLabels = [
    "Semaine 1",
    "Semaine 2",
    "Semaine 3",
    "Semaine 4",
    "Semaine 5",
    "Semaine 6",
    "Semaine 7 et les suivantes"
  ];
  
  // Fonction pour générer uniquement le tableau dans une feuille donnée.
  function ecrireTableau(nomFeuille, weeklyDataSource) {
    var reportSheet = ss.getSheetByName(nomFeuille);
    if (!reportSheet) {
      reportSheet = ss.insertSheet(nomFeuille);
    } else {
      reportSheet.clear();
    }
    
    // Construction du tableau : première colonne "Semaine", puis une colonne par année.
    var header = ["Semaine"].concat(anneeOnglets);
    var tableData = [header];
    for (var i = 0; i < semaineLabels.length; i++) {
      var row = [semaineLabels[i]];
      anneeOnglets.forEach(function(year) {
        var val = (weeklyDataSource[year] && weeklyDataSource[year][i+1]) ? weeklyDataSource[year][i+1] : 0;
        row.push(val);
      });
      tableData.push(row);
    }
    
    // Écrire le tableau dans la feuille
    var numRows = tableData.length;
    var numCols = tableData[0].length;
    reportSheet.getRange(1, 1, numRows, numCols).setValues(tableData);
    reportSheet.autoResizeColumns(1, numCols);
  }
  
  // Générer les tableaux de données (sans graphiques)
  var weeklyDataEte = {};
  anneeOnglets.forEach(function(year) {
    weeklyDataEte[year] = weeklyCounts[year]["Été"];
  });
  ecrireTableau("Graphique par semaine - Été", weeklyDataEte);
  
  var weeklyDataAutomne = {};
  anneeOnglets.forEach(function(year) {
    weeklyDataAutomne[year] = weeklyCounts[year]["Automne-hiver"];
  });
  ecrireTableau("Graphique par semaine - Automne-hiver", weeklyDataAutomne);
  
  ecrireTableau("Graphique par semaine - Total", weeklyCountsTotal);
  
//  SpreadsheetApp.getUi().alert("Rapport hebdomadaire (tableaux) généré avec succès.");
}
