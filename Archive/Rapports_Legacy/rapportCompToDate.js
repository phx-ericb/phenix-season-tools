/**
 * Fonction principale qui génère pour chaque saison (Été et Automne-hiver)
 * un onglet de comparaison entre la dernière année (détectée automatiquement)
 * et chaque année historique.
 * Le tableau présente, pour chaque semaine (1 à 7, avec "Semaine 7 et suivantes")
 * la valeur de l'année historique, la valeur de la dernière année, la différence absolue
 * et la variation en pourcentage, ainsi qu'une ligne "Total to date" (cumul).
 */
function genererComparaisonsDerniereAnnee() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Liste des onglets années (à adapter si nécessaire)
  var anneeOnglets = ["2019", "2020", "2019", "2021", "2022", "2023", "2024", "2025"];
  // Pour la cohérence, on peut aussi rechercher les feuilles dont le nom est un nombre...
  // Ici, nous utilisons un tableau prédéfini.
  
  // Détecter la dernière année en prenant le maximum
  var latestYear = Math.max.apply(null, anneeOnglets.map(function(y){ return parseInt(y, 10); })).toString();
  
  var seasons = ["Été", "Automne-hiver"];
  
  // Pour chaque saison, pour chaque année historique (inférieure à la dernière année)
  seasons.forEach(function(season) {
    anneeOnglets.forEach(function(year) {
      if (parseInt(year, 10) >= parseInt(latestYear, 10)) return;
      
      // Calculer les données hebdomadaires pour l'année historique et pour la dernière année
      var dataHist = computeWeeklyData(year, season);
      var dataLatest = computeWeeklyData(latestYear, season);
      
      // Préparer le tableau
      // Ligne d'en-tête : ["Semaine", "Année historique", "Dernière année", "Diff (Nb)", "Variation (%)"]
      var header = ["Semaine", year, latestYear, "Diff (Nb)", "Variation (%)"];
      var semaineLabels = [
        "Semaine 1",
        "Semaine 2",
        "Semaine 3",
        "Semaine 4",
        "Semaine 5",
        "Semaine 6",
        "Semaine 7 et suivantes"
      ];
      var tableData = [];
      tableData.push(header);
      
      // Pour les 7 semaines
      for (var w = 1; w <= 7; w++) {
        var histCount = dataHist.weekly[w] || 0;
        var latestCount = dataLatest.weekly[w] || 0;
        var diff = latestCount - histCount;
        var variation = (histCount === 0) ? "N/A" : (diff / histCount * 100).toFixed(2) + "%";
        tableData.push([semaineLabels[w-1], histCount, latestCount, diff, variation]);
      }
      
      // Ligne supplémentaire pour le total cumulatif (total to date)
      var cumulativeHist = dataHist.cumulative[7] || 0;
      var cumulativeLatest = dataLatest.cumulative[7] || 0;
      var cumDiff = cumulativeLatest - cumulativeHist;
      var cumVariation = (cumulativeHist === 0) ? "N/A" : (cumDiff / cumulativeHist * 100).toFixed(2) + "%";
      tableData.push(["Total to date", cumulativeHist, cumulativeLatest, cumDiff, cumVariation]);
      
      // Nom de l'onglet de comparaison, par exemple "Comparaison 2025 vs 2019 - Été"
      var sheetName = "Comparaison " + latestYear + " vs " + year + " - " + season;
      var compSheet = ss.getSheetByName(sheetName);
      if (!compSheet) {
        compSheet = ss.insertSheet(sheetName);
      } else {
        compSheet.clear();
      }
      
      // Écriture du tableau à partir de A1
      compSheet.getRange(1, 1, tableData.length, tableData[0].length).setValues(tableData);
      compSheet.autoResizeColumns(1, tableData[0].length);
    });
  });
}

/**
 * Fonction utilitaire qui calcule pour une année et une saison donnée :
 * - Les inscriptions par semaine (selon la règle : les 42 premiers enregistrements répartis
 *   sur 6 semaines de 7, le reste en semaine 7)
 * - Les totaux cumulés par semaine (total to date)
 * Pour les années antérieures à 2023, on utilise le format agrégé (date en colonne A, total en B, saison en C).
 * Pour les années 2023 et plus, on parcourt les inscriptions individuelles (date en colonne G, saison en H, année en I)
 * et on ne prend en compte que les inscriptions dont la date de facture est antérieure ou égale à aujourd'hui.
 *
 * @param {string} year - Année à traiter (ex: "2025")
 * @param {string} season - "Été" ou "Automne-hiver"
 * @return {Object} - {weekly: {1: count, ..., 7: count}, cumulative: {1: total, ..., 7: total}}
 */
function computeWeeklyData(year, season) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(year);
  if (!sheet) return {weekly: {}, cumulative: {}};
  
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return {weekly: {}, cumulative: {}};
  
  var isOldFormat = parseInt(year, 10) < 2023;
  var today = new Date();
  var dateCounts = {}; // agrégation des inscriptions par date (clé "yyyy-mm-dd")
  
  // Parcourir les données en ignorant l'en-tête
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowSeason = "";
    var dateVal;
    var count = 0;
    
    if (isOldFormat) {
      // Ancien format : Date en A (index 0), total en B (index 1), saison en C (index 2)
      dateVal = row[0];
      count = parseInt(row[1], 10) || 0;
      rowSeason = row[2] ? row[2].toString().trim() : "";
    } else {
      // Nouveau format : Date en G (index 6), saison en H (index 7), année en I (index 8)
      var rowYear = row[8] ? row[8].toString().trim() : "";
      if (rowYear !== year) continue;
      dateVal = row[6];
      // Ne prendre en compte que les inscriptions jusqu'à aujourd'hui
      var dateObjCheck = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
      if (dateObjCheck > today) continue;
      count = 1;
      rowSeason = row[7] ? row[7].toString().trim() : "";
    }
    
    // Normalisation de la saison
    var rowSeasonLower = rowSeason.toLowerCase();
    var normalizedSeason = "";
    if (rowSeasonLower.indexOf("été") !== -1 || rowSeasonLower.indexOf("ete") !== -1) {
      normalizedSeason = "Été";
    } else if (rowSeasonLower.indexOf("automne") !== -1 || rowSeasonLower.indexOf("hiver") !== -1) {
      normalizedSeason = "Automne-hiver";
    }
    if (normalizedSeason !== season) continue;
    
    if (!dateVal) continue;
    var dateObj = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    if (isNaN(dateObj.getTime())) continue;
    
    // Formater la date en "yyyy-mm-dd"
    var yyyy = dateObj.getFullYear();
    var mm = ("0" + (dateObj.getMonth() + 1)).slice(-2);
    var dd = ("0" + dateObj.getDate()).slice(-2);
    var dateKey = yyyy + "-" + mm + "-" + dd;
    
    if (!dateCounts[dateKey]) dateCounts[dateKey] = 0;
    dateCounts[dateKey] += count;
  }
  
  // Trier les dates et les regrouper en semaines selon la règle :
  // les 42 premières inscriptions (en ordre chronologique) sont réparties sur 6 semaines (7 par semaine),
  // le reste est attribué à la semaine 7.
  var dateKeys = Object.keys(dateCounts);
  dateKeys.sort();
  var weekly = {1:0,2:0,3:0,4:0,5:0,6:0,7:0};
  dateKeys.forEach(function(dateKey, index) {
    var order = index + 1;
    var week = (order <= 42) ? Math.floor((order - 1) / 7) + 1 : 7;
    weekly[week] += dateCounts[dateKey];
  });
  
  // Calcul des totaux cumulés par semaine
  var cumulative = {};
  var cum = 0;
  for (var w = 1; w <= 7; w++) {
    cum += weekly[w];
    cumulative[w] = cum;
  }
  
  return {weekly: weekly, cumulative: cumulative};
}
