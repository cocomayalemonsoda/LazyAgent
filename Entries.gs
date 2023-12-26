function entriesHowTo() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('EntriesPage')
    .setWidth(450)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'How to submit entries');
}
