// ======== SETUP BEFORE USING =========
// step 1: create a google form
// step 2: copy this code (everything)
// step 3: paste into: your google form > more (the 3 dots) > Apps Script > (you'll see the code editor)
//         make use you are logged-in WITH the google account 
//         you wish to send reminder email from (!important), usually your organization's comms acct.
// step 4: fill in the VARIABLES session bellow
// step 5: select "init" from the function dropdown (top bar right to Debug)
//         click Run Button (left to Debug)
//         this lets you grant necessary permissions to run this script (click advanced > trust)
//         NOTE: in the warning page the "developer's email address" is instead that of the form owner
//         this also add all requested triggers, to remove all triggers, run removeAllTriggers()

// >> made with <3 by Carson. github.com/carsonjan/ <<
// version: v0.2.2

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
// REQUIRES: form setting record_email_address=TRUE, OR from has a question with title including (case insensitive) "email", "e-mail", OR "e mail". 
// If multiple questions applies the first with non-empty answer will be used.

// 1) send email at datetime, see (B) NOTE for format guidelines
const EMAIL_SEND_AT = "2125/11/20 13:03";
// 2) the title/subject of the reminder email. Surround with " 
const REMINDER_EMAIL_TITLE = "Event reminder";
// 3) the top half (before event details) of the reminder email. Surround with ` .
//    line breaks (return key) will reflect on your email body
const REMINDER_EMAIL_TOP = `Hello,

This is a reminder of your upcoming event. Details as follows:`;
// 4) embed form title and description in the email if this value is true. options: [true, false]
const EMBED_EVENT_DETAILS = true;
// 5) the bottom half (after event details) of the reminder email. Surround with ` .
//    line breaks (return key) will reflect on your email body
const REMINDER_EMAIL_BOTTOM = `If you cannot attend the event, please reply to this email so we can allocate your spot to another attendee.

Please feel free to reach out if you have any questions. We look forward to seeing you.

Best Regards,
Example`;

// !! remember to go back and do step 6 and 7

// ======= CODE (do not modify below unless you know what you're doing) =======
// ============================================================================
const DATE_REGEX = /\d{4}\/(0[1-9]|1[0-2])\/\d{2} \d{2}:\d{2}/; // do not modify
const USER_EMAIL_ADDR = Session.getActiveUser().getEmail(); // do not modify

// request permission from google services required to run the script and add event triggers
function init() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);

  FormApp.getActiveForm().setPublished(true);
  FormApp.getActiveForm().setAcceptingResponses(false);

  if (OPEN_DATETIME != "") {
    checkDateFormat_(OPEN_DATETIME, "OPEN_DATETIME");
    FormApp.getActiveForm().setAcceptingResponses(false);
    Logger.log("Form closed by using OPEN_DATETIME in init()");
    ScriptApp.newTrigger("openForm_")
      .timeBased()
      .at(new Date(OPEN_DATETIME + ":00"))
      .create();
    Logger.log(`Trigger to open form at ${OPEN_DATETIME} added by init()`);
    if (MESSAGE_BEFORE_OPEN != "") setClosedFormMessage_(MESSAGE_BEFORE_OPEN);
  }

  if (CLOSE_DATETIME != "") {
    checkDateFormat_(CLOSE_DATETIME, "CLOSE_DATETIME");
    ScriptApp.newTrigger("closeForm_")
      .timeBased()
      .at(new Date(CLOSE_DATETIME + ":00"))
      .create();
    Logger.log(`Trigger to close form at ${CLOSE_DATETIME} added by init()`);
  }

  if (EMAIL_SEND_AT != "") {
    checkDateFormat_(EMAIL_SEND_AT, "EMAIL_SEND_AT");
    ScriptApp.newTrigger("sendReminderEmail_")
      .timeBased()
      .at(new Date(EMAIL_SEND_AT + ":00"))
      .create();
    Logger.log(`Trigger to send reminder email at ${EMAIL_SEND_AT} added by init()`);
    Logger.log(`> NOTE: Reminder Email will be sent from ${USER_EMAIL_ADDR}`);
  }

  if (RESPONSE_LIMIT != -1) {
    if (RESPONSE_LIMIT <= 0) throw new Error("RESPONSE_LIMIT cannot be less than or equal to 0");
    ScriptApp.newTrigger("limitResponses_")
      .forForm(FormApp.getActiveForm())
      .onFormSubmit()
      .create();
    Logger.log(`Trigger to limit form to ${RESPONSE_LIMIT} responses added by init()`);
  }

  Logger.log(`> NOTE: To remove all triggers, run removeAllTriggers() function`);
  Logger.log(`>> All init finished. FORM IS READY TO USE <<`);
  Logger.log(FormApp.getActiveForm().getPublishedUrl());
}

function removeAllTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log("All triggers have been removed.");
}

function checkDateFormat_(str, varName) {
  if (DATE_REGEX.test(str)) return;
  const msg = `Invalid Datetime Format for variable ${varName}. 
  Use YYYY/MM/DD HH:mm in 24 hour format
  prepend numbers with 0 when needed
  > all triggers now removed, please fix format and rerun init() <`;
  removeAllTriggers();
  throw new Error(msg);
}

function setClosedFormMessage_(msg) {
  const form = FormApp.getActiveForm();
  const wasOpen = form.isAcceptingResponses();

  if (!wasOpen) form.setAcceptingResponses(true);
  form.setCustomClosedFormMessage(msg);
  if (!wasOpen) form.setAcceptingResponses(false);
  
  Logger.log(`Custom closed form message set to: ${msg}`);
}

// limit the total number of responses on this google form.
function limitResponses_() {
  if (RESPONSE_LIMIT == -1) return;

  const form = FormApp.getActiveForm();
  const responses = form.getResponses(); 
  const currentCount = responses.length;
  
  Logger.log(`Form submitted. Current response count: ${currentCount}`);

  if (currentCount >= RESPONSE_LIMIT) {
    form.setAcceptingResponses(false);
    if (MESSAGE_FORM_FULL != "") setClosedFormMessage_(MESSAGE_FORM_FULL);
    Logger.log(`Limit of ${RESPONSE_LIMIT} reached. The form has been closed.`);
  }
}

// open the form for response
function openForm_() {
  const form = FormApp.getActiveForm();
  form.setAcceptingResponses(true);
  Logger.log("form opened");
}

// close the form for response
function closeForm_() {
  const form = FormApp.getActiveForm();
  form.setAcceptingResponses(false);
  if (MESSAGE_AFTER_CLOSE != "") setClosedFormMessage_(MESSAGE_AFTER_CLOSE);
  Logger.log("form closed");
}

// send reminder email
// REQUIRES: form setting record_email_address=TRUE, OR from has a question with title including (case insensitive) "email", "e-mail", OR "e mail". If multiple questions applies the first with non-empty answer will be used.
function sendReminderEmail_() {
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

  const formDetails = EMBED_EVENT_DETAILS ? `
-------------------
${formTitle}
===================
${formIntro}
-------------------

` : "";

  const emailBody = REMINDER_EMAIL_TOP + formDetails + REMINDER_EMAIL_BOTTOM;

  // send email
  MailApp.sendEmail({
    to: USER_EMAIL_ADDR,
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