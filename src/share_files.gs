/**
 * @OnlyCurrentDoc
 */

/**
 * Shares partner files with the corresponding managers (LDAPs).
 * @version 1.0
 * @date 2025-12-02
 * @change Initial version.
 */
function sharePartnerFiles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Consolidate by ldap';
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.error('Sheet "' + sheetName + '" not found.');
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    console.log('No data to process.');
    return;
  }
  
  // Get all data from column A (emails) and column B (rich text links)
  var emailRange = sheet.getRange(2, 1, lastRow - 1, 1);
  var linkRange = sheet.getRange(2, 2, lastRow - 1, 1);
  
  var emails = emailRange.getValues();
  var richTextValues = linkRange.getRichTextValues();
  
  var filesShared = 0;
  
  for (var i = 0; i < emails.length; i++) {
    var email = emails[i][0];
    var richText = richTextValues[i][0];
    
    if (email && email.toString().trim() !== '' && richText) {
      var runs = richText.getRuns();
      for (var j = 0; j < runs.length; j++) {
        var url = runs[j].getLinkUrl();
        if (url) {
          try {
            var fileId = extractFileIdFromUrl(url);
            if (fileId) {
              var file = DriveApp.getFileById(fileId);
              file.addEditor(email.toString().trim());
              console.log('Shared file "' + file.getName() + '" with ' + email);
              filesShared++;
            }
          } catch (e) {
            console.error('Error sharing file at URL ' + url + ' with ' + email + ': ' + e.message);
          }
        }
      }
    }
  }
  
  console.log('Finished sharing files. Total shared: ' + filesShared);
}

/**
 * Extracts the file ID from a Google Drive URL.
 * @param {string} url The Google Drive URL.
 * @return {string|null} The file ID or null if not found.
 */
function extractFileIdFromUrl(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}
