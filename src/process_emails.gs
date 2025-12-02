/**
 * Consolidates partner emails from columns W, Z, AC, and AF into column AK.
 * Appends "@google.com" to each LDAP found.
 * Starts from row 3.
 * 
 * @version 1.1
 * @date 2025-12-02
 * @change Removed UI alerts, added console logs and versioning.
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
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var ldaps = [];
    
    // Check columns W, Z, AC, AF (0-based index in data array)
    // Indices in data array are colIndex - 1
    var valW = row[colW - 1];
    var valZ = row[colZ - 1];
    var valAC = row[colAC - 1];
    var valAF = row[colAF - 1];
    
    if (valW && valW.toString().trim() !== '' && valW.toString().trim() !== '#N/A') ldaps.push(valW.toString().trim() + '@google.com');
    if (valZ && valZ.toString().trim() !== '' && valZ.toString().trim() !== '#N/A') ldaps.push(valZ.toString().trim() + '@google.com');
    if (valAC && valAC.toString().trim() !== '' && valAC.toString().trim() !== '#N/A') ldaps.push(valAC.toString().trim() + '@google.com');
    if (valAF && valAF.toString().trim() !== '' && valAF.toString().trim() !== '#N/A') ldaps.push(valAF.toString().trim() + '@google.com');
    
    outputEmails.push([ldaps.join(', ')]);
  }
  
  // Write to column AK starting from row 3
  sheet.getRange(3, colAK, outputEmails.length, 1).setValues(outputEmails);
  
  // SpreadsheetApp.getUi().alert('Processed ' + outputEmails.length + ' rows.');
  console.log('Processed ' + outputEmails.length + ' rows.');
}
