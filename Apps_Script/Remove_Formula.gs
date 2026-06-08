function removeFormulaMultipleSheets() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Add all required sheet names here
  var sheetNames = [
    "Final-25-26",
    "Room-25-26",
    "Verify-25-26"
  ];


  sheetNames.forEach(function(name) {

    var sheet = ss.getSheetByName(name);

    if (sheet) {

      var range = sheet.getDataRange();

      var formulas = range.getFormulas();
      var values = range.getValues();


      for (var r = 0; r < formulas.length; r++) {

        for (var c = 0; c < formulas[r].length; c++) {

          // Convert only formula cells to values
          if (formulas[r][c] != "") {

            sheet
              .getRange(r + 1, c + 1)
              .setValue(values[r][c]);

          }

        }

      }
    }

  });

}
