var TEMPLATE_ID = {
  'one-way': '17Iv5CK3VksoTgvpMK42go25GqimF6P-vuKF9_q5cAQc',
  'mutual': '115JGvcpt43uTrg99tLMwLtyJJNmZikebzAEoUfMKfi0'
};
var TEMPLATE_NAME = '_help.com-nda-generator-template-';
var TITLE = 'NDA Generator by Help.com';

/**
 * The initialization method that is run as soon as the web app loads. (Native)
 * Load skeleton.html and create a web view from it. Additionally, set the title.
 *
 * @method doGet
 * @return {HtmlOutput} The web view built.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('skeleton').setTitle('NDA Generator by Help.com');
}

/**
 * Check to see if a template exists in the active user's drive. If so, get the
 * template Doc and its URL. Otherwise, return false.
 *
 * @method isTemplateFound
 * @param {String} type - Type of NDA.
 * @return {Object|Boolean} If the template exists, the Doc and the URL. If not, false.
 */
function isTemplateFound(type) {
  var ndaTemplateSearch = DriveApp.getFilesByName(TEMPLATE_NAME + type);
  if (ndaTemplateSearch.hasNext()) {
    var template = ndaTemplateSearch.next();
    return {
      template: template,
      url: template.getUrl()
    };
  } else {
    return false;
  }
}

/**
 * Copies an empty NDA template from Help.com's drive into the active user's drive.
 * It then spits back the URL of the copy.
 *
 * @method copyTemplate
 * @param {String} type - Type of NDA.
 * @return {String} The URL of the copied NDA template.
 */
function copyTemplate(type) {
  return DriveApp.getFileById(TEMPLATE_ID[type]).makeCopy(TEMPLATE_NAME + type).getUrl();
}

/**
 * Take the NDA template from the active user's drive and generate an NDA with the
 * desired specifications. The fields that are replaced are:
 * {USERNAME}, {USERTITLE},
 * {RECIPIENTNAME}, {RECIPIENTTITLE}, {RECIPIENTADDRESS},
 * and {DATE}. The NDA is saved as nda-recipient-name.
 *
 * The IDs of each HTML input field directly correspond to the field.
 *
 * @method generateNda
 * @param {Object} details - The required details that should be in the NDA.
 * @return {Object} If the NDA template exists, spit back the generated NDA's URL and the details. If not, return an error.
 */
function generateNda(details) {
  var filename = 'nda-' + details.recipientName.replace(' ', '-').toLowerCase();

  var ndaTemplateSearch = isTemplateFound(details.ndaType);
  var ndaFile = false;
  if (ndaTemplateSearch) {
    ndaFile = ndaTemplateSearch.template.makeCopy(filename);
  }

  if (ndaFile) {
    var ndaId = ndaFile.getId();
    var ndaUrl = ndaFile.getUrl();
    var ndaDoc = DocumentApp.openById(ndaId);
    var ndaBody = ndaDoc.getBody();

    ndaBody.replaceText('{USERNAME}', details.userName);
    ndaBody.replaceText('{USERTITLE}', details.userTitle);
    
    ndaBody.replaceText('{RECIPIENTNAME}', details.recipientName);
    switch (details.recipientType) {
      case 'individual':
        ndaBody.replaceText('{RECIPIENTTITLE}', details.recipientTitle);
        ndaBody.replaceText('{RECIPIENTTYPE}', 'an individual');
        ndaBody.replaceText('{RECIPIENTADDRESS}', '');
        ndaBody.replaceText('{RECIPIENTSIGNER}', details.recipientName);
        break;
      case 'corporation':
        ndaBody.replaceText('{RECIPIENTTITLE}', '');
        ndaBody.replaceText('{RECIPIENTTYPE}', 'a corporation');
        ndaBody.replaceText('{RECIPIENTADDRESS}', 'with its principal address at ' + details.recipientAddress);
        ndaBody.replaceText('{RECIPIENTSIGNER}', details.recipientSigner);
        break;
      case 'llc':
        ndaBody.replaceText('{RECIPIENTTITLE}', '');
        ndaBody.replaceText('{RECIPIENTTYPE}', 'a limited liability company');
        ndaBody.replaceText('{RECIPIENTADDRESS}', 'with its principal address at ' + details.recipientAddress);
        ndaBody.replaceText('{RECIPIENTSIGNER}', details.recipientSigner);
        break;
    }
    
    ndaBody.replaceText('{DATE}', details.date);

    ndaDoc.saveAndClose();

    DriveApp.getFileById(ndaId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
    return { url: ndaUrl, details: details };
  } else {
    // Template not found.
    return { error: 404 };
  }
}

/**
 * Send an E-Mail to the recipient with a link to the NDA Doc.
 *
 * @method sendEmail
 * @param {Object} details - Specifications for the email.
 */
function sendEmail(details) {
  MailApp.sendEmail({
    to: details.recipientEmail,
    subject: details.subject,
    htmlBody: '<p>' + details.body.replace(/\r\n|\r|\n/g, '<br/>') + '</p>',
    name: details.userName
  });
}
