/**
 * Combines data from a specified column across all sheets
 * whose names start with 'FY'.
 *
 * @param {string} columnLetter - The letter of the column to combine.
 * @returns {Array} - An array containing combined data from the specified column.
 */
function getCombinedColumn(columnLetter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Get the active spreadsheet
  var sheets = ss.getSheets(); // Get all sheets in the spreadsheet
  var combinedData = []; // Initialize an array to store combined data

  // Ensure columnLetter is always defined
  columnLetter = columnLetter || 'A';
  Logger.log("Column letter: " + columnLetter);

  // Iterate through each sheet in the spreadsheet
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();
    
    // Check if the sheet name starts with 'FY'
    if (sheetName.startsWith('FY')) {
      var lastRow = sheet.getLastRow(); // Get the last row with data in the sheet
      Logger.log("Processing sheet: " + sheetName + ", lastRow: " + lastRow);
      
      // Ensure there's data beyond the header row
      if (lastRow > 1) {
        try {
          // Log the range string
          var rangeStr = columnLetter + '2:' + columnLetter + lastRow;
          Logger.log("Range string: " + rangeStr);
          
          // Get the range of data
          var range = sheet.getRange(rangeStr);
          // Get values from the range
          var values = range.getValues();
          
          // Filter out rows that are completely empty
          values = values.filter(function(row) {
            return row.some(function(cell) {
              return cell !== '';
            });
          });

          // Flatten the filtered array and concatenate the values to the combinedData array
          combinedData = combinedData.concat(values.flat());
        } catch (e) {
          // Log an error message if there's an issue getting the range
          Logger.log("Error getting range from sheet: " + sheetName + " - " + e.message);
        }
      } else {
        // Log a message if the sheet has no data beyond the header row
        Logger.log("Sheet " + sheetName + " has no data beyond the header row.");
      }
    } else {
      Logger.log("Skipping sheet: " + sheetName + " (does not start with 'FY')");
    }
  }
  
  // Log the combined data for debugging purposes
  Logger.log("Combined Data: " + combinedData.join(", "));
  return combinedData; // Return the combined data
}

/**
 * Clears the content of the TRANSFORM tab from row 2 downwards, except for columns K and L,
 * sets formulas in A2 to Q2 (skipping K2 and L2), and then updates it with combined data.
 */
function updateTransformTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transformSheet = ss.getSheetByName('TRANSFORM'); // Get the TRANSFORM tab
  var lastRow = transformSheet.getLastRow();
  
  // Step 1: Clear existing content from row 2 onwards, from column A to Z, except for columns K and L
  var columnsToClear = 'ABCDEFGHIJMNOPQRSTUVWXYZ'.split('');
  for (var i = 0; i < columnsToClear.length; i++) {
    var columnLetter = columnsToClear[i];
    transformSheet.getRange(columnLetter + '2:' + columnLetter + lastRow).clearContent();
  }
  Logger.log("Cleared content in TRANSFORM tab from row 2 onwards, except for columns K and L.");

  // Step 2: Place formulas in A2 to Q2, skipping K2 and L2
  var columnLetters = {
    'A': 'O',  // Column A in TRANSFORM will use column A in FY tabs
    'B': 'P',  // Column B in TRANSFORM will use column B in FY tabs
    'C': 'V',  // Column C in TRANSFORM will use column C in FY tabs
    'D': 'W',  // Column D in TRANSFORM will use column D in FY tabs
    'E': 'Q',  // Column E in TRANSFORM will use column E in FY tabs
    'F': 'R',  // Column F in TRANSFORM will use column F in FY tabs
    'G': 'S',  // Column G in TRANSFORM will use column G in FY tabs
    'H': 'T',  // Column H in TRANSFORM will use column H in FY tabs
    'J': 'AC',  // Column J in TRANSFORM will use column J in FY tabs
    'M': 'X',  // Column M in TRANSFORM will use column M in FY tabs
    'N': 'Y',  // Column N in TRANSFORM will use column N in FY tabs
    'O': 'Z',  // Column O in TRANSFORM will use column O in FY tabs
    'P': 'AA',  // Column P in TRANSFORM will use column P in FY tabs
    'Q': 'AD'   // Column Q in TRANSFORM will use column Q in FY tabs
  };

  for (var column in columnLetters) {
    var formula = '=ARRAYFORMULA(getCombinedColumn("' + columnLetters[column] + '"))';
    transformSheet.getRange(column + '2').setFormula(formula);
  }
  Logger.log("Placed formulas in A2 to Q2, skipping K2 and L2.");

  // Step 3: Log final update status
  Logger.log("Updated TRANSFORM tab with new data.");
}
