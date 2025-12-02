/**
 * Consolidates partner emails from columns W, Z, AC, and AF into column AK.
 * Appends "@google.com" to each LDAP found.
 * Starts from row 3.
 * 
 * @version 1.2
 * @date 2025-12-02
 * @change Added logic to handle multiple LDAPs per cell and remove duplicates.
 */
function processPartnerEmails() {
  var sheetName = 'Consolidate by Partner';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // SpreadsheetApp.getUi().alert('Sheet "' + sheetName + '" not found.');
    console.error('Sheet "' + sheetName + '" not found.');
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    return; // No data to process
  }
  
  // Column indices (1-based)
  // W = 23, Z = 26, AC = 29, AF = 32, AK = 37
  var colW = 23;
  var colZ = 26;
  var colAC = 29;
  var colAF = 32;
  var colAK = 37;
  
  var range = sheet.getRange(3, 1, lastRow - 2, colAK);
  var data = range.getValues();
  var outputEmails = [];
  
  // Regular expression to split by common separators: space, comma, slash, semicolon
  var separatorRegex = /[\s,\/;]+/;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var ldapsSet = new Set();
    
    // Columns to check: W, Z, AC, AF (0-based indices)
    var columnsToCheck = [colW - 1, colZ - 1, colAC - 1, colAF - 1];
    
    columnsToCheck.forEach(function(colIndex) {
      var cellValue = row[colIndex];
      if (cellValue && cellValue.toString().trim() !== '' && cellValue.toString().trim() !== '#N/A') {
        var parts = cellValue.toString().split(separatorRegex);
        parts.forEach(function(part) {
          var trimmedPart = part.trim();
          if (trimmedPart !== '' && trimmedPart !== '#N/A') {
            ldapsSet.add(trimmedPart + '@google.com');
          }
        });
      }
    });
    
    var sortedLdaps = Array.from(ldapsSet).sort();
    outputEmails.push([sortedLdaps.join(', ')]);
  }
  
  // Write to column AK starting from row 3
  sheet.getRange(3, colAK, outputEmails.length, 1).setValues(outputEmails);
  
  // SpreadsheetApp.getUi().alert('Processed ' + outputEmails.length + ' rows.');
  console.log('Processed ' + outputEmails.length + ' rows.');
}
