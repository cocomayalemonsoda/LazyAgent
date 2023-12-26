function showGuidelines() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('GuidelinesPage')
    .setWidth(450)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'How to use the script');
}
