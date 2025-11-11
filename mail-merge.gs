// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/*
Copyright 2022 Martin Hawksey

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/
 
/**
 * @OnlyCurrentDoc
*/
 
/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
 */

/*
Copied to mail_merge_config.gs
const RECIPIENT_COL  = "Email";
const EMAIL_SENT_COL = "Email Sent";
const EMAIL_SUBJECTS = Object.freeze(["mail merge test 1", "mail merge test 2"]);
const EMAIL_SENDER = "communications@fruitbowlfriends.org"

var emailMenuItems = [];

// Based on https://stackoverflow.com/a/64402332/341400
(function() {
  for (var i = 0; i < EMAIL_SUBJECTS.length; ++i) {
    let subject = EMAIL_SUBJECTS[i];
    let functionName = 'sendEmail_handler_' + i;
    emailMenuItems.push([subject, functionName]);
    this[functionName] = () => sendEmails(subject);
  }
})();
*/

function mail_merge_onOpen() {
      var ui = SpreadsheetApp.getUi();
      var menu = ui.createMenu("Mail Merge");
      emailMenuItems.forEach( O => menu.addItem(O[0], O[1]) );
      menu.addToUi();
    }

    
/**
 * Finds the first row containing all given values and returns it, the data below it and offset
 * @param List of values found on the header row
 * @param data array
 */
function findHeaderAndData(expectedHeaders, data) {
  for (var i = 0; i < data.length; ++i) {
    const row = data[i];
    if (expectedHeaders.every(v => row.indexOf(v) >= 0)) {
      return [row, data.slice(i + 1), i];
    }
  }
  Browser.msgBox("Can't find row containing headers: " + expectedHeaders.join());
}

/**
 * Sends emails from sheet data.
 * @param {string} subjectLine (optional) for the email draft message
 * @param {Sheet} sheet to read data from
*/
function sendEmails(subjectLine, sheet=SpreadsheetApp.getActiveSheet()) {
  if (EMAIL_SENDER) {
    const sessionEmailAddress = Session.getActiveUser().getEmail();
    if (sessionEmailAddress != EMAIL_SENDER) {
      Browser.msgBox("Mail must be sent by " + EMAIL_SENDER + ", not " + sessionEmailAddress);
      return;
    }
  }
  let emailSentColumnHeader = EMAIL_SENT_COL;
  if (subjectLine) {
    emailSentColumnHeader = EMAIL_SENT_COL + "[" + subjectLine + "]";
  } else {
    subjectLine = Browser.inputBox("Mail Merge", 
                                      "Type or copy/paste the subject line of the Gmail " +
                                      "draft message you would like to mail merge with:",
                                      Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == ""){ 
    // If no subject line, finishes up
    return;
    }
  }

  // Gets the draft Gmail message to use as a template
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);

  // Gets the data from the passed sheet
  const dataRange = sheet.getDataRange();
  
  // Fetches displayed values for each row in the Range HT Andrew Roberts 
  // https://mashe.hawksey.info/2020/04/a-bulk-email-mail-merge-with-gmail-and-google-sheets-solution-evolution-using-v8/#comment-187490
  // @see https://developers.google.com/apps-script/reference/spreadsheet/range#getdisplayvalues
  const dataRaw = dataRange.getDisplayValues();

  const [heads, data, headsOffset] = findHeaderAndData([RECIPIENT_COL, emailSentColumnHeader], dataRaw);
 
  // Gets the index of the column named 'Email Status' (Assumes header names are unique)
  // @see http://ramblings.mcpher.com/Home/excelquirks/gooscript/arrayfunctions
  const emailSentColIdx = heads.indexOf(emailSentColumnHeader);
  
  // Converts 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // For a pretty version, see https://mashe.hawksey.info/?p=17869/#comment-184945
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  // Creates an array of email recipients
  const recipients = [];

  // Loops through all the rows of data, populating recipients.
  obj.forEach(function(row, rowIdx){
    // Only sends emails if email_sent cell is blank and not hidden by a filter
    if (row[emailSentColumnHeader] == '') {
      recipients.push(row[RECIPIENT_COL]);
    }
  });

  let send_yes_no = Browser.msgBox("Send " + recipients.length +
      " emails to " + recipients.join(", ") + ".", Browser.Buttons.YES_NO);
  if (send_yes_no != "yes") {
    return;
  }

  obj.forEach(function(row, rowIdx){
     if (row[emailSentColumnHeader] == '') {
      let result = 'unknown error';
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object)
        // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
        // Uncomment advanced parameters as needed (see docs for limitations)
        MailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody: msgObj.html,
          // bcc: 'a.bcc@email.com',
          // cc: 'a.cc@email.com',
          // from: 'an.alias@email.com',
          // name: 'name of the sender',
          // replyTo: 'a.reply@email.com',
          // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
          attachments: emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages
        });
        // Edits cell to record email sent date
        result = new Date();
      } catch(e) {
        // modify cell to record error
        result = e.message;
      }
      // Updates the cell with new data
      sheet.getRange(2 + headsOffset + rowIdx, emailSentColIdx+1).setValue(result);
    }
  });
  
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subject_line to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subject_line){
    try {
      // get drafts
      const drafts = GmailApp.getDrafts();
      // filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      // get the message object
      const msg = draft.getMessage();

      // Handles inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      // Gets all attachments and inline image attachments
      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true,includeAttachments:false});
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody(); 

      // Creates an inline image object with the image name as key 
      // (can't rely on image index as array based on insert order)
      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj) ,{});

      //Regexp searches for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      //Initiates the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);

      return {message: {subject: subject_line, text: msg.getPlainBody(), html:htmlBody}, 
              attachments: attachments, inlineImages: inlineImagesObj };
    } catch(e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subject_line to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subject_line){
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
  
  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data) {
    // We have two templates one for plain text and the html body
    // Stringifing the object means we can do a global replace
    let template_string = JSON.stringify(template);

    // Token replacement
    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return  JSON.parse(template_string);
  }

  /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
  */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  };
}
