function onOpen() {
  var ui = SpreadsheetApp.getUi();
 ui.createMenu("Rapports Inscriptions")
  .addItem("Mise à jour des rapports", "genererRapportComplet")
//.addItem("Générer Graphiques Joueurs", "genererGraphiquesJoueurs")
  //.addItem("Générer Rapport Hebdo", "genererTableauxInscriptionsHebdo")  
  //.addItem("Générer Tableau Variation", "genererSyntheseVariationEte")    
//.addItem("Mise à jour comparaisons", "genererComparaisonsDerniereAnnee")   
  //.addItem("Préparer les données pour Looker", "genererDonneesLookerJoueurs")     
  .addToUi();
}
