/**************************************
 * Main - Générer le rapport complet
 **************************************/
function genererRapportComplet() {
  genererRapport2025();
  genererSyntheseInscriptionsStandardisees();
  genererTableauxInscriptionsHebdo();  
  genererDonneesLookerJoueurs();
  genererTableauxProgressionParSecteur();
  genererSyntheseVariationEte();
  generateSuiviGenerations();
  SpreadsheetApp.getUi().alert("Rapport complet généré avec succès.");
}
