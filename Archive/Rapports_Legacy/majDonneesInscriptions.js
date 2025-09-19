function genererRapport2025() {
  // Accès au classeur cible (celui où le script est lancé) et à l'onglet "2025"
  var targetSS = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = targetSS.getSheetByName("2025");
  if (!targetSheet) {
    targetSheet = targetSS.insertSheet("2025");
  } else {
    targetSheet.clearContents();
  }
  
  // Ouvrir le classeur source et accéder à l'onglet "INSCRIPTIONS"
  var sourceSS = SpreadsheetApp.openById("1xYOQf0kPt8vKKCHpMrNSQYS6kS7SGX_dehjkpdQLo5w");
  var sourceSheet = sourceSS.getSheetByName("INSCRIPTIONS");
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("L'onglet INSCRIPTIONS est introuvable dans la feuille source.");
    return;
  }
  
  // Récupérer toutes les données du classeur source
  var sourceData = sourceSheet.getDataRange().getValues();
  var output = [];
  
  // Ajout de la ligne d'en-tête
  var header = [
    "Passeport #", 
    "Prénom", 
    "Nom", 
    "Date de naissance", 
    "Identité de genre", 
    "Nom du frais", 
    "Date de la facture", 
    "Saison", 
    "Année", 
    "Statut de l'inscription"
  ];
  output.push(header);
  
  // Parcourir les lignes du source (en ignorant la première ligne d'en-tête)
  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    // Création d'une nouvelle ligne avec les colonnes demandées
    // Index source (0-based) :
    // A (Passeport #)        -> row[0]
    // B (Prénom)             -> row[1]
    // C (Nom)                -> row[2]
    // D (Date de naissance)  -> row[3]
    // F (Identité de genre)  -> row[5]
    // L (Nom du frais)       -> row[11]
    // Z (Date de la facture) -> row[25]
    // Saison                 -> "Été"
    // Année                  -> "2025"
    // W (Statut de l'inscription) -> row[22]
    var newRow = [];
    newRow.push(row[0]);    // Passeport #
    newRow.push(row[1]);    // Prénom
    newRow.push(row[2]);    // Nom
    newRow.push(row[3]);    // Date de naissance
    newRow.push(row[5]);    // Identité de genre
    newRow.push(row[11]);   // Nom du frais
    newRow.push(row[25]);   // Date de la facture
    newRow.push("Été");     // Saison
    newRow.push("2025");    // Année
    newRow.push(row[22]);   // Statut de l'inscription
    output.push(newRow);
  }
  
  // Écriture des données dans l'onglet cible "2025" à partir de la cellule A1
  targetSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
  
  // Forcer la première colonne (Passeport #) en format texte à partir de la 2ème ligne (les en-têtes restent inchangées)
  if (output.length > 1) {
    targetSheet.getRange(2, 1, output.length - 1, 1).setNumberFormat("@");
  }
  
  // Optionnel : ajuster la largeur des colonnes
  targetSheet.autoResizeColumns(1, output[0].length);
  
//  SpreadsheetApp.getUi().alert("Rapport '2025' généré avec succès.");
}
