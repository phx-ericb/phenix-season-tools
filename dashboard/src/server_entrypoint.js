function doGet() {
  return HtmlService.createTemplateFromFile('ui').evaluate()
    .setTitle('Suivi Inscriptions – Dashboard v0.1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
