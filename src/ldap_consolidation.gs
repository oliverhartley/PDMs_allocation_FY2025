/**
 * @OnlyCurrentDoc
 */

/**
 * Consolidates partner data by LDAP from the 'Consolidate by Partner' sheet 
 * into the 'Consolidate by ldap' sheet.
 * @version 1.0
 * @date 2025-12-02
 * @change Initial version.
 */
function consolidateByLdap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = 'Consolidate by Partner';
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    console.error('Sheet "' + sourceSheetName + '" not found.');
    return;
  }
  
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 3) {
    console.log('No data to process in "' + sourceSheetName + '".');
    return;
  }
  
  // Column indices (1-based)
  var colA = 1;  // Partner Name
  var colAK = 37; // Emails
  
  var range = sourceSheet.getRange(3, 1, lastRow - 2, colAK);
  var data = range.getValues();
  
  var ldapMap = {};
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var partnerName = row[colA - 1];
    var emails = row[colAK - 1];
    
    if (emails && emails.toString().trim() !== '') {
      var emailArray = emails.toString().split(', ');
      emailArray.forEach(function(email) {
        var trimmedEmail = email.trim();
        if (trimmedEmail !== '') {
          if (!ldapMap[trimmedEmail]) {
            ldapMap[trimmedEmail] = new Set();
          }
          ldapMap[trimmedEmail].add(partnerName);
        }
      });
    }
  }
  
  var targetSheetName = 'Consolidate by ldap';
  var targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
    console.log('Sheet "' + targetSheetName + '" created.');
  }
  
  targetSheet.clearContents();
  targetSheet.appendRow(['ldap', 'Partner']);
  
  var outputData = [];
  for (var email in ldapMap) {
    var partners = Array.from(ldapMap[email]).sort().join(', ');
    outputData.push([email, partners]);
  }
  
  // Sort by email address
  outputData.sort(function(a, b) {
    return a[0].localeCompare(b[0]);
  });
  
  if (outputData.length > 0) {
    targetSheet.getRange(2, 1, outputData.length, 2).setValues(outputData);
  }
  
  console.log('Consolidation by LDAP complete. Processed ' + Object.keys(ldapMap).length + ' unique emails.');
}
