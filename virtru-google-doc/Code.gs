/**
 *  Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
      .addItem('Start', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}



/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.  Using a template allows externalization of CSS
 * and JS files. 
 */

function showSidebar() {  
  var template = HtmlService.createTemplateFromFile('virtruSidebar');
  var ui = template.evaluate().setTitle('Virtru Protect & Share');
  DocumentApp.getUi().showSidebar(ui);
}


/*
 * Allows use of external files for CSS.
 *
 * @param {string} fileName Name of the file to import.
 */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}


function doGet(request) {
  return HtmlService.createTemplateFromFile('virtruSidebar')
      .evaluate();
}




/**
 * Grabs & returns the name of the document.
 */

function getFileName() {
  var doc = DocumentApp.getActiveDocument();
  var title = doc.getName();
  return title;
}


/**
 * Gets document content as a blob, then converts to bytes,
 * base64-encodes, and returns that data to pass to client.
 *
 * @return {string} blob64 The base64'd PDF blob data.
 */
function createPDF() {
  var docBlob = DocumentApp.getActiveDocument().getBlob();
  docBlob.setName(doc.getName() + '.pdf');
  var blobB64 = Utilities.base64Encode(docBlob.getBytes());
  return blobB64;
}

/**
 * Generates the email to the list of authorized users.
 * Email body is dictated by `emailHTML.html` file. 
 * Adds the pdf.tdf3.html attachment.
 *
 * @param {string} cipherText  The HTML formatted ciphertext of the file to be added as an attachment.
 * @param {array}  recipients  The users granted access to this file; will receive an email with attachment.
 * @param {string} userMessage The custom messaging the file owner entered when sharing.
 */

function sendEmail(cipherText, recipients, userMessage) {

  // Get email address of file owner and assign attachment title.
  var fileOwner = Session.getActiveUser().getEmail();
  var fileName = DocumentApp.getActiveDocument().getName() + ".pdf.tdf3.html";
  
  // Provide a basic email body for recipients who do not support HTML.
  var emailBody = fileOwner + " has shared the encrypted file " + fileName + " with you.\r\n\r\nIt\'s attached below; please download to open in Virtru\'s Secure Reader.";
  
  // Assign values to variables in emailHTML.html template.
  var htmlContent = HtmlService.createTemplateFromFile('emailHTML');
  htmlContent.fileOwner = fileOwner;
  htmlContent.fileName = fileName;
  htmlContent.userMessage = userMessage;
  
  // Create subject line based on filename and owner email address. 
  var subject = fileOwner + ' has shared a secure file: "' + fileName + '"';
  
  // Convert ciphertext string to HTML blob.
  var blob = Utilities.newBlob(cipherText, 'text/html', fileName);
  
  
  // Send the email with the tdf.html blob as attachment.
  MailApp.sendEmail(recipients, subject, emailBody, {
    name: fileOwner,
    attachments: [blob],
    htmlBody: htmlContent.evaluate().getContent()
  });
}
