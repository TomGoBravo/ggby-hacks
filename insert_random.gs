function insert_random_onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Random')
      .addItem('Set Active Range Randomly', 'fillActiveRangeWithRandomNumbers')
      .addToUi();
}


/** created by google search for 'google sheets active range with random numbers appscript' */
function fillActiveRangeWithRandomNumbers() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeRange = spreadsheet.getActiveRange();

  if (activeRange) {
    var numRows = activeRange.getNumRows();
    var numColumns = activeRange.getNumColumns();
    var values = [];

    // Define the range for random numbers (e.g., between 1 and 100)
    var min = 100000;
    var max = 999999;

    for (var i = 0; i < numRows; i++) {
      var row = [];
      for (var j = 0; j < numColumns; j++) {
        // Generate a random integer between min and max (inclusive)
        var randomNumber = Math.floor(Math.random() * (max - min + 1)) + min;
        row.push(randomNumber);
      }
      values.push(row);
    }
    activeRange.setValues(values);
  } else {
    Browser.msgBox("No active range selected.", "Please select a range of cells before running this script.", Browser.Buttons.OK);
  }
}
