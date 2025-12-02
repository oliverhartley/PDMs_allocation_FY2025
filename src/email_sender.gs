/**
 * @OnlyCurrentDoc
 */

/**
 * Sends emails to partners requesting domain updates.
 * @version 1.0
 * @date 2025-12-02
 * @change Initial version.
 * @param {boolean} testMode - If true, sends only the first email to oliverhartley@google.com.
 */
function sendDomainUpdateEmails(testMode) {
  var sheetName = 'Consolidate by Partner';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.error('Sheet "' + sheetName + '" not found.');
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    console.log('No data to process.');
    return; // No data to process
  }
  
  // Column indices (1-based)
  var colA = 1;  // Partner Name
  var colAI = 35; // Domain
  var colAK = 37; // Emails
  
  var range = sheet.getRange(3, 1, lastRow - 2, colAK);
  var data = range.getValues();
  var emailsSent = 0;
  
  var subject = '[Action Needed] - Completar Dominio (Columna AI) en "PDMs allocation - FY2025"';
  var sheetUrl = 'https://docs.google.com/spreadsheets/d/1XUVbK_VsV-9SsUzfp8YwUF2zJr3rMQ1ANJyQWdtagos/edit?gid=1288523495#gid=1288523495';
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var partnerName = row[colA - 1];
    var domain = row[colAI - 1];
    var emails = row[colAK - 1];
    
    if ((domain == null || domain.toString().trim() === '' || domain.toString().trim() === '#N/A') && (emails && emails.toString().trim() !== '')) {
      var recipients = emails.toString().trim();
      var firstLdap = recipients.split('@')[0]; // Assuming the first email is the primary contact
      
      var body = 'Hola ' + firstLdap + ',\n\n' +
                 'Te escribo en relación al archivo "PDMs allocation - FY2025": ' + sheetUrl + '\n\n' +
                 'Ya que apareces como responsable del Partner: ' + partnerName + ', te pido por favor que completes la Columna AI, llamada "Domain", con el dominio web de ese partner (por ejemplo: Partner.com).\n\n' +
                 '¡Muchas gracias por tu ayuda!\n\n' +
                 'Saludos,\n' +
                 'Oliver';
               if (testMode) {
        GmailApp.sendEmail('oliverhartley@google.com', subject, body, {cc: 'jcarrique@google.com'});
        console.log('Test email sent to oliverhartley@google.com for partner: ' + partnerName);
        emailsSent++;
        break; // Send only one email in test mode
      } else {
        GmailApp.sendEmail(recipients, subject, body, {cc: 'jcarrique@google.com'});
        console.log('Email sent to ' + recipients + ' for partner: ' + partnerName);
        emailsSent++;
      }
    }
  }
  
  console.log('Finished processing. Emails sent: ' + emailsSent);
}

/**
 * @version 1.0
 * @date 2025-12-02
 * @change Initial version.
 */
function sendDomainUpdateEmailsReal() {
  sendDomainUpdateEmails(false);
}

/**
 * @version 1.0
 * @date 2025-12-02
 * @change Initial version.
 */
function sendDomainUpdateEmailTest() {
  sendDomainUpdateEmails(true);
}
