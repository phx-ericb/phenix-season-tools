/*******************************
 * Génération de la synthèse des inscriptions standardisées avec Saison,
 * regroupées par Secteur. Pour les années 2023,2024,2025, la synthèse
 * sépare par genre (Féminin/Masculin). Pour les années 2019-2022, on ne
 * fait qu'un total par saison (secteur par défaut = "Total").
 *******************************/
function genererSyntheseInscriptionsStandardisees() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Charger le mapping depuis l'onglet "Standardisation"
  var standardSheet = ss.getSheetByName("Standardisation");
  if (!standardSheet) {
    SpreadsheetApp.getUi().alert("L'onglet Standardisation est introuvable.");
    return;
  }
  var standardData = standardSheet.getDataRange().getValues(); // En-tête en ligne 1
  // Stocker dans mapping l'objet standardisé complet : {standard, type, secteur}
  var mapping = {}; // mapping[original] = { standard: ..., type: ..., secteur: ... }
  for (var i = 1; i < standardData.length; i++) {
    var row = standardData[i];
    var original = row[0] ? row[0].toString().trim().toLowerCase() : "";
    var standard = row[1] ? row[1].toString().trim() : "";
    var type = row[2] ? row[2].toString().trim() : "";
    var secteur = row[3] ? row[3].toString().trim() : "";
    if (original !== "" && secteur !== "") {
      mapping[original] = {
        standard: standard,
        type: type,
        secteur: secteur
      };
    }
  }
  
  // 2. Définir les onglets à traiter
  var anneeOnglets = ["2019", "2020", "2021", "2022", "2023", "2024", "2025"];
  var oldYears = ["2019", "2020", "2021", "2022"];
  
  // summary[key] = { feminin: compteur, masculin: compteur, total: compteur }
  // La clé sera "Année_Saison_Secteur" où la saison est normalisée.
  var summary = {};
  
  // 3. Parcourir les onglets et agréger les inscriptions
  for (var j = 0; j < anneeOnglets.length; j++) {
    var sheetName = anneeOnglets[j];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) continue;
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    if (data.length < 2) continue; // Pas de données
    
    if (oldYears.indexOf(sheetName) !== -1) {
      // Pour les années 2019-2022 : colonnes : A: Date de facture, B: Total inscription, C: Saison
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        // Total inscription en colonne B
        var totalInscription = row[1] ? parseInt(row[1].toString().trim(), 10) : 0;
        // Saison en colonne C
        var saisonBrut = row[2] ? row[2].toString().trim() : "";
        var saison = "";
        var saisonLower = saisonBrut.toLowerCase();
        if (saisonLower.indexOf("été") !== -1 || saisonLower.indexOf("ete") !== -1) {
          saison = "Été";
        } else if (saisonLower.indexOf("automne") !== -1 || saisonLower.indexOf("hiver") !== -1) {
          saison = "Automne Hiver";
        } else {
          saison = saisonBrut;
        }
        // Pour ces onglets, il n'y a pas de secteur ni de genre,
        // on fixe le secteur à "Total" par défaut.
        var secteur = "Total";
        var annee = sheetName; // L'année est déduite du nom de l'onglet.
        var keySummary = annee + "_" + saison + "_" + secteur;
        if (!summary[keySummary]) {
          summary[keySummary] = { feminin: 0, masculin: 0, total: 0 };
        }
        summary[keySummary].total += totalInscription;
      }
    } else {
      // Pour les années 2023,2024,2025 :
      // Colonnes attendues :
      // A: Passeport #, B: Prénom, C: Nom, D: Date de naissance, E: Identité de genre,
      // F: Nom du frais, G: Date de facture, H: Saison, I: Année, J: Statut.
      for (var i = 1; i < data.length; i++) {
        var row = data[i];
        // Ignorer les inscriptions annulées (colonne J, index 9)
        var statut = row[9] ? row[9].toString().trim() : "";
        if (statut === "Annulé") continue;
        
        // Nom du frais en colonne F (index 5)
        var originalItem = row[5] ? row[5].toString().trim().toLowerCase() : "";
        if (originalItem === "") continue;
        
var standardEntry = mapping[originalItem];
if (!standardEntry) {
  for (var key in mapping) {
    if (originalItem.indexOf(key) !== -1) {
      standardEntry = mapping[key];
      break;
    }
  }
}

// Ici, vérifie absolument que standardEntry existe
if (!standardEntry) {
  Logger.log("Aucun mapping trouvé pour l'item : " + originalItem + ", ligne ignorée.");
  continue; // saute cette ligne, car pas d'entrée trouvée
}

       
        
        // Récupérer l'année : colonne I (index 8) s'il y a une valeur, sinon le nom de l'onglet
        var annee = row[8] ? row[8].toString().trim() : sheetName;
        // Récupérer et normaliser la saison : colonne H (index 7)
        var saisonBrut = row[7] ? row[7].toString().trim() : "";
        var saison = "";
        var saisonLower = saisonBrut.toLowerCase();
        if (saisonLower.indexOf("été") !== -1 || saisonLower.indexOf("ete") !== -1) {
          saison = "Été";
        } else if (saisonLower.indexOf("automne") !== -1 || saisonLower.indexOf("hiver") !== -1) {
          saison = "Automne Hiver";
        } else {
          saison = saisonBrut;
        }
        // Récupérer l'identité de genre : colonne E (index 4)
        var genre = row[4] ? row[4].toString().trim() : "";
        // La clé de regroupement se base sur Année, Saison et Secteur (standardisé)
        var keySummary = annee + "_" + saison + "_" + standardEntry.secteur;
        if (!summary[keySummary]) {
          summary[keySummary] = { feminin: 0, masculin: 0, total: 0 };
        }
        summary[keySummary].total++;
        var genreLower = genre.toLowerCase();
        if (genreLower.indexOf("féminin") !== -1) {
          summary[keySummary].feminin++;
        } else if (genreLower.indexOf("masculin") !== -1) {
          summary[keySummary].masculin++;
        }
      }
    }
  }
  
  // 4. Construire le tableau de sortie avec en-têtes :
  // Année | Saison | Secteur | Féminin | Masculin | Total
  var output = [];
  output.push(["Année", "Saison", "Secteur", "Féminin", "Masculin", "Total"]);
  
  var keys = Object.keys(summary);
  keys.sort(function(a, b) {
    var partsA = a.split("_");
    var partsB = b.split("_");
    var yearA = parseInt(partsA[0], 10);
    var yearB = parseInt(partsB[0], 10);
    if (yearA !== yearB) return yearA - yearB;
    if (partsA[1] !== partsB[1]) return partsA[1].localeCompare(partsB[1]);
    return partsA[2].localeCompare(partsB[2]);
  });
  
  for (var i = 0; i < keys.length; i++) {
    var parts = keys[i].split("_");
    var stats = summary[keys[i]];
    output.push([parts[0], parts[1], parts[2], stats.feminin, stats.masculin, stats.total]);
  }
  
  // 5. Écriture du résultat dans un onglet de synthèse
  var destSheet = ss.getSheetByName("Synthèse Inscriptions Standardisées");
  if (!destSheet) {
    destSheet = ss.insertSheet("Synthèse Inscriptions Standardisées");
  } else {
    destSheet.clearContents();
  }
  
  destSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  destSheet.autoResizeColumns(1, output[0].length);
  
//  SpreadsheetApp.getUi().alert("Synthèse générée avec succès.");
}
