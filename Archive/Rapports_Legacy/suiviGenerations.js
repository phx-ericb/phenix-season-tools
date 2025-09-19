/*******************************************************
 * Fonction principale : Génère l'onglet "Suivi Générations"
 * avec colonnes de variations entre saisons et années.
 *******************************************************/
function generateSuiviGenerations() {
  //=== 1) Paramètres et références de base
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // On définit explicitement les années à traiter
  var targetYears = [2023, 2024, 2025];
  
  // On récupère les feuilles correspondant à ces années
  var annualSheets = [];
  targetYears.forEach(function(yr) {
    var sheet = ss.getSheetByName(yr.toString());
    if (sheet) {
      annualSheets.push({year: yr, sheet: sheet});
    }
  });
  
  // Tri par année décroissante (pour l'affichage dans le tableau final)
  annualSheets.sort(function(a, b) {
    return b.year - a.year;
  });
  
  // Déterminer l'année la plus récente (pour filtrer U4–U19)
  var newestYear = Math.max.apply(null, targetYears);
  
  //=== 2) Structure d'agrégation
  // aggregate[key] = {
  //   birthYear: <année de naissance>,
  //   genre: <"F" ou "M">,
  //   inscriptions: {
  //     <année>: { ete: 0, automne: 0, categorie: "Uxx" }
  //   }
  // }
  var aggregate = {};
  
  //=== 3) Lecture de chaque feuille annuelle et agrégation
  annualSheets.forEach(function(item) {
    var year = item.year;
    var sheet = item.sheet;
    var data = sheet.getDataRange().getValues();
    
    // On suppose que la 1re ligne est l'en-tête
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      
      // Indices de colonnes (à adapter si besoin)
      var birthColIndex = 3;   // Date/Année de naissance
      var genreColIndex  = 4;   // Genre
      var saisonColIndex = 7;   // Saison (Été, Automne, Automne Hiver, etc.)
      var statutColIndex = 9;   // Statut (Annulé ?)
      
      // Vérif du statut : on ignore si "annulé"
      var statut = row[statutColIndex];
      if (statut && statut.toString().toLowerCase().indexOf("annulé") !== -1) {
        continue; 
      }
      
      // Récup de l'année de naissance
      var birthValue = row[birthColIndex];
      var birthYear;
      if (birthValue instanceof Date) {
        // Si c'est un objet Date, on prend l'année
        birthYear = birthValue.getFullYear();
      } else {
        // Sinon on parse comme un nombre
        birthYear = parseInt(birthValue, 10);
      }
      if (!birthYear || isNaN(birthYear)) {
        continue; // Donnée invalide
      }
      
      // Récup du genre
      var genre = row[genreColIndex] ? row[genreColIndex].toString().trim() : "";
      if (!genre) {
        continue; // Genre non renseigné
      }
      
      // Récup de la saison
      var saison = row[saisonColIndex] ? row[saisonColIndex].toString().trim().toLowerCase() : "";
      // On essaie de reconnaître « été » ou « automne/hiver »
      var seasonKey = "";
      if (saison.indexOf("été") !== -1 || saison.indexOf("ete") !== -1) {
        seasonKey = "ete";
      } else if (saison.indexOf("automne") !== -1 || saison.indexOf("hiver") !== -1) {
        seasonKey = "automne";
      } else {
        // Saison non reconnue => on ignore
        continue;
      }
      
      // Calcul de la catégorie : U + (année - annéeNaissance)
      var catNum = year - birthYear;
      // On considère seulement les catégories U4 à U19
      var categorie = (catNum >= 4 && catNum <= 19) ? "U" + catNum : "";
      
      // Clé d'agrégation
      var key = birthYear + "_" + genre;
      if (!aggregate[key]) {
        aggregate[key] = {
          birthYear: birthYear,
          genre: genre,
          inscriptions: {}
        };
      }
      if (!aggregate[key].inscriptions[year]) {
        aggregate[key].inscriptions[year] = {
          ete: 0,
          automne: 0,
          categorie: categorie
        };
      }
      
      // Incrément du compteur
      if (seasonKey === "ete") {
        aggregate[key].inscriptions[year].ete++;
      } else if (seasonKey === "automne") {
        aggregate[key].inscriptions[year].automne++;
      }
    }
  });
  
  //=== 4) Préparation de l'onglet "Suivi Générations"
  var suiviSheetName = "Suivi Générations";
  var suiviSheet = ss.getSheetByName(suiviSheetName);
  if (!suiviSheet) {
    suiviSheet = ss.insertSheet(suiviSheetName);
  } else {
    suiviSheet.clear();
  }
  
  //=== 5) Construction de l'entête
  // On ajoute 5 colonnes par année :
  // - Catégorie <year>
  // - Été <year>
  // - Automne <year>
  // - Var E->A <year>
  // - Var E <year>->E <year précédent>
  var header = ["Année de naissance", "Genre"];
  
  annualSheets.forEach(function(item, i) {
    var y = item.year;
    header.push("Catégorie " + y);
    header.push("Été " + y);
    header.push("Automne " + y);
    header.push("Var E->A " + y);  // Variation Été->Automne (même année)
    
    // Variation Été Y -> Été (Y-1) si l'année suivante existe dans annualSheets
    // (i < annualSheets.length - 1) signifie qu'il y a une "année précédente" plus loin dans le tableau
    if (i < annualSheets.length - 1) {
      var olderYear = annualSheets[i+1].year;
      header.push("Var E " + y + "->E " + olderYear);
    } else {
      // Pour la dernière année (la plus ancienne), on mettra "N/A"
      header.push("Var E " + y + "->E ???");
    }
  });
  
  //=== 6) Construction des lignes
  var rows = [];
  var keys = Object.keys(aggregate);
  
  // On trie les clés d'abord par année de naissance, puis par genre
  keys.sort(function(a, b) {
    var aData = aggregate[a];
    var bData = aggregate[b];
    if (aData.birthYear === bData.birthYear) {
      return aData.genre.localeCompare(bData.genre);
    }
    return aData.birthYear - bData.birthYear;
  });
  
  // Parcours de chaque génération (annéeNaissance + genre)
  keys.forEach(function(key) {
    var obj = aggregate[key];
    var birthYear = obj.birthYear;
    
    // Vérifier si la génération est éligible pour l'année la plus récente (U4–U19)
    var catNumForNewest = newestYear - birthYear;
    if (catNumForNewest < 4 || catNumForNewest > 19) {
      // On exclut cette génération du tableau
      return;
    }
    
    // Préparation de la ligne
    var row = [birthYear, obj.genre];
    
    // On va stocker temporairement les données de chaque année pour calculer les variations
    // yearData[i] = { categorie, ete, automne }
    var yearData = [];
    annualSheets.forEach(function(item, i) {
      var y = item.year;
      var insc = obj.inscriptions[y];
      if (insc) {
        yearData[i] = {
          categorie: insc.categorie,
          ete: insc.ete,
          automne: insc.automne
        };
      } else {
        yearData[i] = {
          categorie: "",
          ete: 0,
          automne: 0
        };
      }
    });
    
    // Construire la partie "colonnes" pour chaque année
    annualSheets.forEach(function(item, i) {
      var d = yearData[i];
      
      // 1) Catégorie, Été, Automne
      row.push(d.categorie);
      row.push(d.ete);
      row.push(d.automne);
      
      // 2) Variation Été->Automne (même année)
      var varEtoA = computePctVar(d.automne, d.ete); 
      row.push(varEtoA);
      
      // 3) Variation Été Y -> Été (Y-1)
      if (i < annualSheets.length - 1) {
        // l'année "précédente" est à l'index i+1 (car on a trié en ordre décroissant)
        var dPrev = yearData[i+1];
        var varEtoPrev = computePctVar(d.ete, dPrev.ete);
        row.push(varEtoPrev);
      } else {
        // pour la dernière année (la plus ancienne), on n'a pas de référence
        row.push("N/A");
      }
    });
    
    rows.push(row);
  });
  
  //=== 7) Écriture dans l’onglet "Suivi Générations"
  // Écriture de l'entête
  suiviSheet.getRange(1, 1, 1, header.length).setValues([header]);
  
  // Écriture des données
  if (rows.length > 0) {
    suiviSheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  }
  
  // Ajustement automatique des largeurs de colonnes
  suiviSheet.autoResizeColumns(1, header.length);
}


/*******************************************************
 * Fonction utilitaire : calcule la variation en %
 * newVal : nouvelle valeur
 * oldVal : ancienne valeur (référence)
 *******************************************************/
function computePctVar(newVal, oldVal) {
  if (!oldVal || oldVal === 0) {
    return "N/A";
  }
  var diff = newVal - oldVal;
  var pct = (diff / oldVal) * 100;
  // On arrondit à l'entier près (ou garde 1 décimale si besoin)
  return Math.round(pct) + "%";
}
