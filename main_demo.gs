// ======== SETUP BEFORE USING =========
// step 1: create a google form
// step 2: copy this code (everything)
// step 3: paste into: your google form > more (the 3 dots) > Apps Script > (you'll see the code editor)
//         make use you are logged-in WITH the google account 
//         you wish to send reminder email from (!important), usually your organization's comms acct.
// step 4: fill in the VARIABLES session bellow
// step 5: select "init_1" from the function dropdown (top bar right to Debug)
//         click Run Button (left to Debug)
//         this lets you grant necessary permissions to run this script (click advanced > trust)
//         NOTE: in the warning page the "developer's email address" is instead that of the form owner
//         feel free to run multiple times to make sure permission is granted
// step 6: select "init_2" from the same function dropdown
//         click Run Button (left to Debug)
//         click ONLY ONCE (!important), unless you want to add multiple triggers to a function
//         you can remove triggers manually (left sidebar clock icon), or by running removeAllTriggers()

// >> made with <3 by Carson. github.com/carsonjan/ <<

// ======== VARIABLES ===========

// A) limit number of responses:
// 1) max number of submissions allowed, set to -1 to disable
//    this value is read every time when a new form is submitted
const RESPONSE_LIMIT = 30; 
const MESSAGE_FORM_FULL = "sorry, event sign-up is full"; // only applies when RESPONSE_LIMIT is set. Leave blank ("") to disable.

// B) open/close form at datetime:
// NOTE: input in format "YYYY/MM/DD HH:mm"
//       use 24h time format. prefix with 0 when needed. 
//       time in minutes has max +- 15min delay (Google policy)
//       time is in script timezone. to change, to go sidebar > settings (the gear button)
//       leave blank ("") to disable.
const OPEN_DATETIME = "2125/11/04 13:03";
const CLOSE_DATETIME = "2125/11/14 13:03";

const MESSAGE_BEFORE_OPEN = "event sign-up will open on 4th, 13:03"; // only applies when OPEN_DATETIME is set. Leave blank ("") to disable.
const MESSAGE_AFTER_CLOSE = "sorry, event sign-up is closed"; // only applies when CLOSE_DATETIME is set. Leave blank ("") to disable.

// C) send reminder email:
// REQUIRES: form setting record_email_address=TRUE, OR from has a question with title including (case insensitive) "email", "e-mail", OR "e mail". If multiple questions applies the first with non-empty answer will be used.

// 1) send email at datetime, see (B) NOTE for format guidelines
const EMAIL_SEND_AT = "2125/11/20 13:03";
// 2) the email to send reminder email "to:". Usually your current google account's address.
//    this is not the address of your respondents (that automatically goes to "bcc:"). Surround with "
const EMAIL_ADDR_SELF = "example@gmail.com";
// 3) the title/subject of the reminder email. Surround with " 
const REMINDER_EMAIL_TITLE = "Event reminder";
// 4) the top half (before event details) of the reminder email. Surround with ` .
//    line breaks (return key) will reflect on your email body
const REMINDER_EMAIL_TOP = `Hello,

This is a reminder of your upcoming event. Details as follows:`;
// 5) the bottom half (after event details) of the reminder email. Surround with ` .
//    line breaks (return key) will reflect on your email body
const REMINDER_EMAIL_BOTTOM = `If you cannot attend the event, please reply to this email so we can allocate your spot to another attendee.

Please feel free to reach out if you have any questions. We look forward to seeing you.

Best Regards,
Example`;

// !! remember to go back and do step 6 and 7

// ======= CODE ===========

// request permission from google services required to run the script
function init_1() {
  try {
    FormApp.getActiveForm();
    MailApp.getRemainingDailyQuota();
    ScriptApp.getAuthorizationInfo();
  } catch (e) {
    // do nothing
  }
  Logger.log(`init_1 finished. Run init2().`);
}

// add event triggers
function init_2() {
  if (OPEN_DATETIME != "") {
    FormApp.getActiveForm().setAcceptingResponses(false);
    Logger.log("Form closed by using OPEN_DATETIME in init2()");
    ScriptApp.newTrigger("openForm")
      .timeBased()
      .at(new Date(OPEN_DATETIME + ":00"))
      .create();
    Logger.log(`Trigger to open form at ${OPEN_DATETIME} added by init2()`);
    if (MESSAGE_BEFORE_OPEN != "") FormApp.getActiveForm().setCustomClosedFormMessage(MESSAGE_BEFORE_OPEN);
  }

  if (CLOSE_DATETIME != "") {
    ScriptApp.newTrigger("closeForm")
      .timeBased()
      .at(new Date(CLOSE_DATETIME + ":00"))
      .create();
    Logger.log(`Trigger to close form at ${CLOSE_DATETIME} added by init2()`);
  }

  if (EMAIL_SEND_AT != "") {
    ScriptApp.newTrigger("sendReminderEmail")
      .timeBased()
      .at(new Date(EMAIL_SEND_AT + ":00"))
      .create();
    Logger.log(`Trigger to send reminder email at ${EMAIL_SEND_AT} added by init2()`);
    if (EMAIL_ADDR_SELF == "") Logger.log(`WARNING: you must set EMAIL_ADDR_SELF variable`);
  }

  if (RESPONSE_LIMIT != -1) {
    ScriptApp.newTrigger("limitResponses")
      .forForm(FormApp.getActiveForm())
      .onFormSubmit()
      .create();
    Logger.log(`Trigger to limit form to ${RESPONSE_LIMIT} responses added by init2()`);
  }
  Logger.log(`To remove all triggers, run removeAllTriggers() function`);
  Logger.log(`All init finished. FORM IS READY TO USE.`);
}

function removeAllTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log("All triggers have been removed.");
}

// limit the total number of responses on this google form.
function limitResponses() {
  if (RESPONSE_LIMIT == -1) return;

  const form = FormApp.getActiveForm();
  const responses = form.getResponses(); 
  const currentCount = responses.length;
  
  Logger.log(`Form submitted. Current response count: ${currentCount}`);

  if (currentCount >= RESPONSE_LIMIT) {
    form.setAcceptingResponses(false);
    form.setCustomClosedFormMessage(MESSAGE_FORM_FULL);
    Logger.log(`Limit of ${RESPONSE_LIMIT} reached. The form has been closed.`);
  }
}

// open the form for response
function openForm() {
  const form = FormApp.getActiveForm();
  form.setAcceptingResponses(true);
  Logger.log("form opened");
}

// close the form for response
function closeForm() {
  const form = FormApp.getActiveForm();
  form.setAcceptingResponses(false);
  if (MESSAGE_AFTER_CLOSE != "") form.setCustomClosedFormMessage(MESSAGE_AFTER_CLOSE);
  Logger.log("form closed");
}

// send reminder email
// REQUIRES: form setting record_email_address=TRUE, OR from has a question with title including (case insensitive) "email", "e-mail", OR "e mail". If multiple questions applies the first with non-empty answer will be used.
function sendReminderEmail() {
  const form = FormApp.getActiveForm();
  const responses = form.getResponses();
  
  // get responses email address
  let responsesEmail = [];
  for (const response of responses) {
    if (form.collectsEmail()) {
      responsesEmail.push(response.getRespondentEmail());
    } else {
      const itemResponses = response.getItemResponses();
      // for each question
      outerloop: for (const itemResponse of itemResponses) {
        const itemTitle = itemResponse.getItem().getTitle().toLowerCase();
        const emailStrings = ["email", "e-mail", "e mail"];
        // for each variant of title "email"
        for (const emailString of emailStrings) {
          const emailAddress = itemResponse.getResponse().trim();
          if (itemTitle.includes(emailString) && emailAddress != "") {
            responsesEmail.push(emailAddress);
            break outerloop;
          }
        }
      }
    }
  }

  // abort if Apps Script has insufficient email sending quota 
  let emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  if (emailQuotaRemaining < responsesEmail.length) {
    const msg = `Apps Script has insufficient email sending quota today. Expected=${responsesEmail.length}, RemainingQuota=${emailQuotaRemaining}`;
    throw new Error(msg);
  }

  // build email to be sent
  const formTitle = form.getTitle();
  const formIntro = form.getDescription();
  const bccChain = responsesEmail.join(",");

  const emailBody = REMINDER_EMAIL_TOP + `
-------------------
${formTitle}
===================
${formIntro}
-------------------

` + REMINDER_EMAIL_BOTTOM;

  // send email
  MailApp.sendEmail({
    to: EMAIL_ADDR_SELF,
    subject: REMINDER_EMAIL_TITLE,
    body: emailBody,
    bcc: bccChain
  });
  
  // log result
  emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  try {
    Logger.log(`Event Reminder Email Sent to ${responsesEmail.length} addresses. 
  Event Title: ${formTitle}. 
  Apps Script remaining send email quota today: ${emailQuotaRemaining}
  Bcc List: ${bccChain}`);
  } catch (e) { // if bccChain is too long
    Logger.log(`Event Reminder Email Sent to ${responsesEmail.length} addresses. 
  Event Title: ${formTitle}. 
  Apps Script remaining send email quota today: ${emailQuotaRemaining}
  Bcc List: (chain is too long, not displayed)`);
  }
}
