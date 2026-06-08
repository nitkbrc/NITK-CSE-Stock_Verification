// BRC - Extract entire rows from "Accession" containing "FAR" to "Acc-FAR"
// Keeps first 3 rows of Acc-FAR unchanged

function copy_row_format() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var source = ss.getSheetByName("Accession"); 
  var target = ss.getSheetByName("Acc-FAR");

  var data = source.getDataRange().getValues();
  var lastColumn = source.getLastColumn();


  // ---------- Clear target sheet from Row 4 onwards only ----------
  var lastTargetRow = target.getLastRow();

  if (lastTargetRow > 3) {
    target
      .getRange(
        4,
        1,
        lastTargetRow - 3,
        target.getLastColumn()
      )
      .clear();
  }


  // ---------- Retain Column Widths ----------
  for (var c = 1; c <= lastColumn; c++) {
    target.setColumnWidth(
      c,
      source.getColumnWidth(c)
    );
  }


  // ---------- Copy Matching Rows ----------
  var targetRow = 4;   // Start after top 3 rows


  for (var i = 0; i < data.length; i++) {

    // Check Column K contains FAR
    if (data[i][10].toString().includes("FAR")) {


      source
        .getRange(i + 1, 1, 1, lastColumn)
        .copyTo(
          target.getRange(targetRow, 1),
          {contentsOnly:false}
        );


      // Retain row height
      target.setRowHeight(
        targetRow,
        source.getRowHeight(i + 1)
      );


      targetRow++;
    }
  }
}
