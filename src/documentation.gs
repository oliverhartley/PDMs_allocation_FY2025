/**
 * @OnlyCurrentDoc
 */

/**
 * Project: PDMs Allocation FY2025
 * 
 * Purpose:
 * This project is designed to organize the email database to send DRP (Digital Rights Management) files to partners.
 * It connects to the spreadsheet ID: 1XUVbK_VsV-9SsUzfp8YwUF2zJr3rMQ1ANJyQWdtagos
 * and is managed via Google Apps Script Project ID: 1V4jym9MNp7DP430jZZQ89VelFjII7NMPjqBcBgUZW5xA5maTtJo0phNo
 * 
 * Main Functions:
 * - Organize partner email database.
 * - Prepare and send DRP files to partners.
 * 
 * Setup:
 * - Clasp is used for local development.
 * - Git is used for version control.
 * 
 * @version 1.0
 * @date 2025-12-02
 * @change Initial version.
 */

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('DRP Management')
      .addItem('Show Documentation', 'showDocumentation')
      .addItem('Process Partner Emails', 'processPartnerEmails')
      .addSeparator()
      .addItem('Send Domain Update Emails (Test)', 'sendDomainUpdateEmailTest')
      .addItem('Send Domain Update Emails (REAL)', 'sendDomainUpdateEmailsReal')
      .addSeparator()
      .addItem('Consolidate by LDAP', 'consolidateByLdap')
      .addToUi();
}

/**
 * @version 1.0
 * @date 2025-12-02
 * @change Initial version.
 */
function showDocumentation() {
  var html = HtmlService.createHtmlOutputFromFile('Documentation_HTML')
      .setTitle('Project Documentation')
      .setWidth(600)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Project Documentation');
}
