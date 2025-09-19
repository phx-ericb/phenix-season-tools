/**
 * Détermine le "jour actuel" pour 2025 en lisant l’onglet "2025".
 * On considère que la première inscription (colonne G) correspond au jour 1.
 * La fonction calcule le nombre de jours entre la première et la dernière date,
 * en excluant les lignes dont le statut (colonne J) est "Annulé".
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
  // Colonne G (index 6) : Date de facture, colonne J (index 9) : Statut
  for (var i = 1; i < data.length; i++) {
    var status = data[i][9] ? data[i][9].toString().trim() : "";
    if (status === "Annulé") continue;
    
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
  var diffMs = derniereDate - premierDate;
  var diffDays = Math.floor(diffMs / (1000 * 3600 * 24));
  Logger.log("Premier jour: " + premierDate + ", Dernier jour: " + derniereDate + ", Jour actuel: " + diffDays);
  return diffDays;
}

/**
 * Génère des tableaux de progression journalière des inscriptions par secteur,
 * pour une saison donnée (par exemple "Ete" ou "Automne Hiver"), en utilisant
 * les données brutes des onglets 2023, 2024 et 2025 (format nouveau).
 *
 * Pour chaque secteur reconnu (selon la table de standardisation de l'onglet "Standardisation"
 * pour le type JOUEURS), le script calcule pour chaque transaction le "jour relatif" (offset)
 * par rapport au premier jour d'inscription de la saison et en déduit le cumul des inscriptions jour par jour.
 *
 * Un onglet est créé pour chaque secteur (nommé "Progression [saisonCible] [secteur]"),
 * avec pour colonnes : "Jour", "2023", "2024", "2025".
 *
 * Les lignes dont le statut (colonne J) est "Annulé" sont ignorées.
 */
function genererTableauxProgressionParSecteur() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Paramètres
  var annees = ["2023", "2024", "2025"];
  var saisonCible = "Ete"; // Modifier à "Automne Hiver" pour l'autre saison
  
  // 1. Charger la table de standardisation et en extraire la liste des secteurs de type JOUEURS
  var standardSheet = ss.getSheetByName("Standardisation");
  if (!standardSheet) {
    SpreadsheetApp.getUi().alert("L'onglet 'Standardisation' est introuvable.");
    return;
  }
  var standardData = standardSheet.getDataRange().getValues();
  var joueursSecteurs = {};
  for (var i = 1; i < standardData.length; i++) {
    var row = standardData[i];
    var type = row[2] ? row[2].toString().trim().toUpperCase() : "";
    if (type === "JOUEURS") {
      var secteurKey = row[3] ? row[3].toString().trim() : "";
      if (secteurKey !== "") {
        joueursSecteurs[secteurKey] = true;
      }
    }
  }
  Logger.log("Secteurs joueurs identifiés : " + JSON.stringify(joueursSecteurs));
  
  // 2. Parcourir les onglets des années et regrouper les transactions par secteur.
  // Format attendu dans chaque onglet (2023, 2024, 2025) :
  //   - Colonne G (index 6) : Date de la facture
  //   - Colonne H (index 7) : Saison (ex. "Ete")
  //   - Colonne I (index 8) : Année (ex. "2023")
  //   - Colonne F (index 5) : Nom du frais
  //   - Colonne J (index 9) : Statut (ex. "Annulé")
  var dataParSecteur = {}; // dataParSecteur[secteur][annee] = array of transactions { date: Date }
  
  annees.forEach(function(annee) {
    var sheet = ss.getSheetByName(annee);
    if (!sheet) {
      Logger.log("Onglet " + annee + " introuvable.");
      return;
    }
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log("Pas de données dans " + annee);
      return;
    }
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Vérifier que la colonne I correspond à l'année (si renseignée)
      var anneeRow = row[8] ? row[8].toString().trim() : "";
      if (anneeRow !== "" && anneeRow !== annee) continue;
      
      // Ignorer les lignes annulées
      var statut = row[9] ? row[9].toString().trim() : "";
      if (statut === "Annulé") continue;
      
      // Normaliser la saison
      var saisonRow = row[7] ? row[7].toString().trim() : "";
      var normalizedSaison = saisonRow.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
      if (normalizedSaison.indexOf("ete") !== -1) {
        normalizedSaison = "Ete";
      } else if (normalizedSaison.indexOf("automne") !== -1 || normalizedSaison.indexOf("hiver") !== -1) {
        normalizedSaison = "Automne Hiver";
      }
      if (normalizedSaison !== saisonCible) continue;
      
      // Déterminer le secteur à partir du nom du frais en utilisant la table de standardisation
      var nomFrais = row[5] ? row[5].toString().trim() : "";
      var secteur = determineSecteur(nomFrais, standardData);
      // Autoriser "Élite" ou "Senior" même s'ils ne figurent pas dans la liste extraite
      if (secteur === "" || (secteur !== "Élite" && secteur !== "Senior" && !joueursSecteurs[secteur])) continue;
      
      // Extraire la date (colonne G)
      var dateFacture = row[6];
      if (!dateFacture) continue;
      var dateObj = (dateFacture instanceof Date) ? dateFacture : new Date(dateFacture);
      if (isNaN(dateObj.getTime())) continue;
      
      Logger.log("Transaction pour " + annee + ": secteur=" + secteur + ", date=" + dateObj);
      
      if (!dataParSecteur[secteur]) {
        dataParSecteur[secteur] = {};
        annees.forEach(function(a) { dataParSecteur[secteur][a] = []; });
      }
      dataParSecteur[secteur][annee].push({ date: dateObj });
    }
  });
  
  Logger.log("Données par secteur: " + JSON.stringify(dataParSecteur));
  
  // 3. Pour chaque secteur, pour chaque année, calculer l'offset (jour relatif) et le cumul.
  for (var secteur in dataParSecteur) {
    var progressionByYear = {};
    var maxJourSecteur = 0;
    
    annees.forEach(function(annee) {
      var transactions = dataParSecteur[secteur][annee];
      transactions.sort(function(a, b) {
        return a.date - b.date;
      });
      if (transactions.length > 0) {
        var startDate = transactions[0].date;
        transactions.forEach(function(tx) {
          var diffMs = tx.date - startDate;
          tx.offset = Math.floor(diffMs / (1000 * 3600 * 24)) + 1;
        });
        var cumMap = {};
        var cumul = 0;
        var maxJour = transactions[transactions.length - 1].offset;
        if (maxJour > maxJourSecteur) maxJourSecteur = maxJour;
        var idx = 0;
        for (var jour = 1; jour <= maxJour; jour++) {
          var countJour = 0;
          while (idx < transactions.length && transactions[idx].offset === jour) {
            countJour++;
            idx++;
          }
          cumul += countJour;
          cumMap[jour] = cumul;
        }
        progressionByYear[annee] = cumMap;
      } else {
        progressionByYear[annee] = {};
      }
    });
    
    // 4. Remplir les jours manquants jusqu'au maximum observé pour ce secteur.
    var filledProgression = {};
    annees.forEach(function(annee) {
      filledProgression[annee] = {};
      var lastVal = 0;
      for (var jour = 1; jour <= maxJourSecteur; jour++) {
        if (progressionByYear[annee][jour] !== undefined) {
          lastVal = progressionByYear[annee][jour];
        }
        filledProgression[annee][jour] = lastVal;
      }
    });
    
    // 5. Construire le tableau final : "Jour", puis colonnes pour chaque année.
    var header = ["Jour"].concat(annees);
    var output = [header];
    for (var jour = 1; jour <= maxJourSecteur; jour++) {
      var rowOut = [jour];
      annees.forEach(function(annee) {
        rowOut.push(filledProgression[annee][jour] || 0);
      });
      output.push(rowOut);
    }
    
    // 6. Écrire le tableau dans un onglet nommé "Progression [saisonCible] [secteur]"
    var sheetName = "Progression " + saisonCible + " " + secteur;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    sheet.autoResizeColumns(1, output[0].length);
    
    Logger.log("Tableau généré pour " + sheetName);
  }
  
  // Optionnel : Afficher une alerte à la fin
  // SpreadsheetApp.getUi().alert("Tableaux de progression par secteur générés pour la saison " + saisonCible + ".");
}

/**
 * Détermine le secteur à partir du nom du frais en comparant avec la table de standardisation.
 * La table (à partir de la ligne 2) doit contenir les colonnes suivantes :
 *   - Colonne A : Original
 *   - Colonne B : Standard
 *   - Colonne C : Type
 *   - Colonne D : Secteur
 * Si le nom du frais (en minuscules) contient une des valeurs "Original", retourne
 * la valeur correspondante de la colonne "Secteur". Si aucune correspondance n'est trouvée,
 * retombe sur une logique par défaut.
 */
function determineSecteur(nomFrais, standardData) {
  if (!nomFrais) return "";
  var s = nomFrais.toLowerCase();
  
  // Parcourir la table de standardisation (à partir de la 2e ligne)
  if (standardData) {
    for (var i = 1; i < standardData.length; i++) {
      var original = standardData[i][0] ? standardData[i][0].toString().trim().toLowerCase() : "";
      var secteur = standardData[i][3] ? standardData[i][3].toString().trim() : "";
      if (original !== "" && secteur !== "") {
        if (s.indexOf(original) !== -1) {
          return secteur;
        }
      }
    }
  }
  
  // Logique par défaut si aucune correspondance trouvée :
  if (s.indexOf("ldp") !== -1 || s.indexOf("d1+") !== -1) return "Élite";
  if (s.indexOf("adapté") !== -1) return "Adapté";
  if (s.indexOf("u-sé") !== -1 || s.indexOf("adulte") !== -1 || s.indexOf("ligue") !== -1 || s.indexOf("o-") !== -1) {
    return "Senior";
  }
  
  var match = s.match(/u-?\s*(\d{1,2})/);
  if (match) {
    var num = parseInt(match[1], 10);
    if (num >= 13 && num <= 19) return "U13-U18";
    if (num >= 9 && num <= 12) return "U9-U12";
    if (num >= 4 && num <= 8) return "U4-U8";
  }
  return "";
}
