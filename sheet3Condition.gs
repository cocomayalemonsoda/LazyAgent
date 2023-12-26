function onEdit(e) {
  var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet3"); // Change to your sheet name
  
  if (e && e.range) {
    var range = e.range;
    var column = range.getColumn();
    var row = range.getRow();
    
    if (column >= 2 && column <= 5) {
      // If "YES" is entered, reset background color for all cells in the row
      if (e.value && e.value.toString().toUpperCase() === "YES") {
        for (var col = 2; col <= 5; col++) {
          sheet3.getRange(row, col).setBackground(null);
        }

        var yesCount = 0;
        for (var col = 2; col <= 5; col++) {
          if (sheet3.getRange(row, col).getValue() && sheet3.getRange(row, col).getValue().toString().toUpperCase() === "YES") {
            yesCount++;
          }
        }

        if (yesCount > 1) {
          for (var col = 2; col <= 5; col++) {
            if (col !== column) {
              sheet3.getRange(row, col).setValue(null).setBackground('#CCCCCC');
            }
          }
        }
      } else {
        // If "YES" is removed, reset background color
        sheet3.getRange(row, column).setBackground(null);
      }
    }
  }
}
