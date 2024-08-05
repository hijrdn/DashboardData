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
Logger.log("Column Letter: " + columnLetter);

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
          // Get values from the range, flatten the array, and filter out empty values
          var values = range.getValues().flat().filter(function(value) {
            return value !== '';
          });
          
          // Concatenate the values to the combinedData array
          combinedData = combinedData.concat(values);
        } catch (e) {
          // Log an error message if there's an issue getting the range
          Logger.log("Error getting range from sheet: " + sheetName + " - " + e.message);
        }
      } else {
        // Log a message if the sheet has no data beyond the header row
        Logger.log("Sheet " + sheetName + " has no data beyond the header row.");
      }
    }
  }
  
  // Log the combined data for debugging purposes
  Logger.log("Combined Data: " + combinedData.join(", "));
  return combinedData; // Return the combined data
}

/**
 * Clears the content of the TRANSFORM tab from row 2 down before updating.
 *
 * @param {string} columnLetter - The letter of the column to combine.
 */
function updateTransformTab(columnLetter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transformSheet = ss.getSheetByName('TRANSFORM'); // Get the TRANSFORM tab
  var lastRow = transformSheet.getLastRow();

// Clear existing content from row 2 and below, columns A to Z
transformSheet.getRange('A2:Z' + lastRow).clearContent();

var combinedData = getCombinedColumn(columnLetter); // Get combined data from FY tabs
  for (var i = 0; i < combinedData.length; i++) {
    transformSheet.getRange(columnLetter + (i + 2)).setValue(combinedData[i]); // Update TRANSFORM tab
  }
}
