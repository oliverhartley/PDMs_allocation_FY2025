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
 * @version 1.1
 * @date 2025-12-02
 * @change Added Google Drive file linking for partner names.
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
  
  var outputData = [];
  var richTextOutput = [];
  
  var sortedEmails = Object.keys(ldapMap).sort();
  
  for (var j = 0; j < sortedEmails.length; j++) {
    var email = sortedEmails[j];
    var partners = Array.from(ldapMap[email]).sort();
    var partnerLinks = [];
    
    for (var k = 0; k < partners.length; k++) {
      partnerLinks.push(getPartnerFileLink(partners[k], folder));
    }
    
    outputData.push([email, '']); // Placeholder for rich text
    
    var combinedRichText = SpreadsheetApp.newRichTextValue().setText(partners[0]); // Fallback
    if (partnerLinks.length > 0) {
      var builder = SpreadsheetApp.newRichTextValue();
      partnerLinks.forEach(function(linkInfo, index) {
        if (index > 0) {
          builder.setText(builder.getText() + ', ');
        }
        if (linkInfo.url) {
          builder.setText(builder.getText() + linkInfo.name);
          // This is not a direct way to build combined rich text with multiple links, 
          // Apps Script RichTextValueBuilder is limited. We'll use formula method.
        } else {
          builder.setText(builder.getText() + linkInfo.name);
        }
      });
      // Instead of rich text, use HYPERLINK formula
      var formulas = partners.map(function(partner) {
        var linkInfo = getPartnerFileLink(partner, folder);
        if (linkInfo.url) {
          return '=HYPERLINK("' + linkInfo.url + '", "' + linkInfo.name + '")';
        } else {
          return '"' + linkInfo.name + '"';
        }
      });
      outputData[j][1] = '=JOIN(", ", ARRAYFORMULA({' + formulas.join(';') + '}))';
    }
  }
  
  if (outputData.length > 0) {
    targetSheet.getRange(2, 1, outputData.length, 2).setValues(outputData);
  }
  
  console.log('Consolidation by LDAP complete. Processed ' + sortedEmails.length + ' unique emails.');
}

/**
 * Searches for a partner file in Drive and returns link info.
 * @version 1.1
 * @date 2025-12-02
 * @change Updated file name search pattern.
 * @param {string} partnerName The name of the partner.
 * @param {Folder} folder The Google Drive folder to search in.
 * @return {object} An object with name and url or just name.
 */
function getPartnerFileLink(partnerName, folder) {
  var searchFileName = partnerName + ' - Partner Dashboard';
  var files = folder.getFilesByName(searchFileName);
  if (files.hasNext()) {
    var file = files.next();
    return { name: file.getName(), url: file.getUrl() };
  } else {
    // Fallback search without suffix
    files = folder.getFilesByName(partnerName);
    if (files.hasNext()) {
      var file = files.next();
      return { name: file.getName(), url: file.getUrl() };
    }
  }
  return { name: partnerName };
}
