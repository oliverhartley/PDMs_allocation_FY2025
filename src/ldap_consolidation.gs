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
/**
 * Consolidates partner data by LDAP from the 'Consolidate by Partner' sheet 
 * into the 'Consolidate by ldap' sheet, linking to partner files in Google Drive.
 * @version 1.3
 * @date 2025-12-02
 * @change Updated partner name column to AH (34).
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
  var colAH = 34; // Partner Name
  var colAK = 37; // Emails
  
  var range = sourceSheet.getRange(3, 1, lastRow - 2, colAK);
  var data = range.getValues();
  
  var ldapMap = {};
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var partnerName = row[colAH - 1];
    var emails = row[colAK - 1];
    
    if (emails && emails.toString().trim() !== '' && partnerName && partnerName.toString().trim() !== '') {
      var emailArray = emails.toString().split(', ');
      emailArray.forEach(function(email) {
        var trimmedEmail = email.trim();
        if (trimmedEmail !== '') {
          if (!ldapMap[trimmedEmail]) {
            ldapMap[trimmedEmail] = new Set();
          }
          ldapMap[trimmedEmail].add(partnerName.toString().trim());
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
  
  var driveFolderId = '1GT-A2Hkg75uXxQF0FYCKROXW8rBw_XjC';
  var folder = DriveApp.getFolderById(driveFolderId);
  var files = folder.getFiles();
  var cachedFiles = [];
  while (files.hasNext()) {
    var file = files.next();
    cachedFiles.push({ name: file.getName(), url: file.getUrl() });
  }
  
  var outputData = [];
  var richTextOutput = [];
  
  var sortedEmails = Object.keys(ldapMap).sort();
  
  for (var j = 0; j < sortedEmails.length; j++) {
    var email = sortedEmails[j];
    var partners = Array.from(ldapMap[email]).sort();
    var partnerLinks = [];
    
    var fullText = '';
    var linkRanges = [];
    
    for (var k = 0; k < partners.length; k++) {
      var linkInfo = getPartnerFileLink(partners[k], cachedFiles);
      if (k > 0) {
        fullText += ', ';
      }
      var start = fullText.length;
      fullText += linkInfo.name;
      var end = fullText.length;
      if (linkInfo.url) {
        linkRanges.push({start: start, end: end, url: linkInfo.url});
      }
    }
    
    var builder = SpreadsheetApp.newRichTextValue().setText(fullText);
    linkRanges.forEach(function(range) {
      builder.setLinkUrl(range.start, range.end, range.url);
    });
    richTextOutput.push([builder.build()]);
    outputData.push([email]);
  }
  
  if (outputData.length > 0) {
    targetSheet.getRange(2, 1, outputData.length, 1).setValues(outputData);
    targetSheet.getRange(2, 2, richTextOutput.length, 1).setRichTextValues(richTextOutput);
  }
  
  console.log('Consolidation by LDAP complete. Processed ' + sortedEmails.length + ' unique emails.');
}

/**
 * Searches for a partner file in cached files and returns link info.
 * @version 1.3
 * @date 2025-12-02
 * @change Updated to use cached files for performance.
 * @param {string} partnerName The name of the partner.
 * @param {Array} cachedFiles Array of cached file objects {name, url}.
 * @return {object} An object with name and url or just name.
 */
function getPartnerFileLink(partnerName, cachedFiles) {
  var partnerNameLower = partnerName.toLowerCase();
  
  for (var i = 0; i < cachedFiles.length; i++) {
    var file = cachedFiles[i];
    if (file.name.toLowerCase().indexOf(partnerNameLower) !== -1) {
      return { name: file.name, url: file.url };
    }
  }
  return { name: partnerName };
}
