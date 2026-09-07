//  YSE HR AUTOMATION SCRIPT
//  Automates: Interview Invitations | Offer Letters | Onboarding


//  CONFIGURATION  <-  UPDATE THESE BEFORE USING

const HR_CONFIG = {

  // The tab name of your current intake sheet (bottom tab in Google Sheets)
  SHEET_NAME: '2026 (February)',   // <- update each new intake cycle

  // YSE HR email details
  HR_EMAIL:   'yse.hr@ysenpo.org',
  HR_NAME:    'YSE HR Team',
  ORG_NAME:   'Youth Society for Education',
  WEBSITE:    'www.youthsocietyforeducation.org',

  // Column positions (1 = Column A).
  // Based on your 2026 (February) sheet layout.
  // If you add/remove columns, update the numbers below.
  COL: {
    NO:                         1,   // A
    NAME:                       2,   // B
    POSITION:                   3,   // C
    DEPARTMENT:                 4,   // D
    EMAIL:                      5,   // E
    BRANCH:                     6,   // F
    // Columns 7–10: Interviewer scores (Seint, Kaung Khant, Chan, Noe)
    FINAL:                      11,  // K
    INTERVIEW_ASSIGNED_TO:      12,  // L
    INTERVIEW_INVITATION_STATUS:13,  // M  <- Trigger 1: type "Sent"
    INTERVIEW_DATE:             14,  // N
    INTERVIEW_REMARK:           15,  // O
    INTERVIEW_STATUS:           16,  // P
    INTERVIEW_STATUS_REMARK:    17,  // Q
    OFFER_LETTER_STATUS:        18,  // R
    OFFER_STATUS:               19,  // S
    ONBOARDING_STATUS:          20,  // T  <- Trigger 3: type "Done"
    FACEBOOK:                   21,  // U
    TELEGRAM:                   22,  // V
  },

  // Google Drive File IDs for document templates and static attachments.
  
  // HOW TO FIND A FILE ID:
  //   1. Open the file in Google Drive
  //   2. Look at the URL: https://drive.google.com/file/d/FILE_ID_HERE/view
  //   3. Copy the long string between /d/ and /view  — that is the FILE ID

  // For Google Docs templates, open the doc and copy from:
  //   https://docs.google.com/document/d/FILE_ID_HERE/edit
  
  DRIVE_FILES: {
    // Paste the Google Doc template ID for the offer letter
    OFFER_LETTER_TEMPLATE:       'PASTE_OFFER_LETTER_GOOGLE_DOC_ID_HERE',
    // Paste the Google Doc template ID for the onboarding packet
    ONBOARDING_PACKET_TEMPLATE:  'PASTE_ONBOARDING_PACKET_GOOGLE_DOC_ID_HERE',
    // Paste the PDF file IDs for static attachments
    EMPLOYEE_HANDBOOK:           'PASTE_EMPLOYEE_HANDBOOK_PDF_ID_HERE',
    BYLAWS:                      'PASTE_BYLAWS_PDF_ID_HERE',
    ACKNOWLEDGEMENT:             'PASTE_ACKNOWLEDGEMENT_PDF_ID_HERE',
  }
};


// Menu

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🎓 YSE HR Tools')
    .addItem('📧 Send Interview Invitation',  'openInterviewDialogManual')
    .addItem('📄 Send Offer Letter',           'openOfferDialog')
    .addItem('📦 Send Onboarding Package',     'openOnboardingDialogManual')
    .addSeparator()
    .addItem('⚙️  First-Time Setup (run once)', 'setupTriggers')
    .addToUi();
}


//  FIRST-TIME SETUP (installs the automatic edit trigger)

function setupTriggers() {
  // Remove any existing trigger with the same handler (avoids duplicates)
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create a new installable onEdit trigger
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    '✅  Setup Complete!\n\n' +
    'Automation is now active. The system will automatically open a\n' +
    'dialog whenever you:\n\n' +
    '  • Set "Interview Invitation Status" to "Sent"\n' +
    '  • Set "Onboarding Packet Preparation Status" to "Done"\n\n' +
    'You can also trigger any step manually via the "🎓 YSE HR Tools" menu.'
  );
}


//  INSTALLABLE ON-EDIT TRIGGER
//  Fires automatically whenever a cell is edited

function onEditInstallable(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== HR_CONFIG.SHEET_NAME) return;

  var row   = e.range.getRow();
  var col   = e.range.getColumn();
  var value = String(e.value || '').trim();

  if (row <= 1) return;  // skip header row

  // TRIGGER 1 — Interview Invitation Status = "Sent"
  if (col === HR_CONFIG.COL.INTERVIEW_INVITATION_STATUS && value === 'Sent') {
    openInterviewDialog(row);
  }

  // TRIGGER 3 — Onboarding Packet Preparation Status = "Done"
  if (col === HR_CONFIG.COL.ONBOARDING_STATUS && value === 'Done') {
    openOnboardingDialog(row);
  }
}


//  SHARED HELPER — reads one candidate row from the sheet

function getRowData(row) {
  var sheet  = SpreadsheetApp.getActiveSpreadsheet()
                 .getSheetByName(HR_CONFIG.SHEET_NAME);
  var values = sheet.getRange(row, 1, 1, 25).getValues()[0];
  var c      = HR_CONFIG.COL;

  var branch = values[c.BRANCH - 1];
  var branchDisplay = branch ? branch + ' Branch' : 'Yangon Branch';

  return {
    row:          row,
    name:         String(values[c.NAME        - 1] || '').trim(),
    position:     String(values[c.POSITION    - 1] || '').trim(),
    department:   String(values[c.DEPARTMENT  - 1] || '').trim(),
    email:        String(values[c.EMAIL       - 1] || '').trim(),
    branch:       branchDisplay,
    rawBranch:    String(branch || '').trim(),
    assignedTo:   String(values[c.INTERVIEW_ASSIGNED_TO - 1] || '').trim(),
  };
}


//  PROCESS 1 INTERVIEW INVITATION

// Opens dialog automatically (called by trigger)
function openInterviewDialog(row) {
  var data     = getRowData(row);
  var template = HtmlService.createTemplateFromFile('InterviewDialog');
  template.data   = data;
  template.config = HR_CONFIG;
  var html = template.evaluate().setWidth(560).setHeight(680);
  SpreadsheetApp.getUi().showModalDialog(
    html, '📧 Interview Invitation — ' + data.name
  );
}

// Opens dialog from the menu (uses the currently selected row)
function openInterviewDialogManual() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName(HR_CONFIG.SHEET_NAME);
  var row   = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please click on a candidate\'s row first, then try again.');
    return;
  }
  openInterviewDialog(row);
}

// Called by InterviewDialog.html when HR clicks "Send"

function sendInterviewEmail(fd) {
  var data = getRowData(fd.row);
  if (!data.email) throw new Error('No email address found for this candidate.');

  var subject =
    'YSE Interview with ' + data.name +
    ' [' + fd.position + '] ' +
    fd.interviewDay + ', ' + fd.interviewDate;

  var body =
    'Dear ' + data.name + ',\n\n' +
    'Congratulations! \n' +
    'After carefully reviewing your application, we are pleased to inform you that you have been shortlisted for an interview with the core recruitment team.\n\n' +
    'You have impressed us with your personal stories and experiences, as well as your passion for the ' + fd.position + ', and we want to help you grow that passion. ' +
    'You are now one step closer to being a part of the ' +
    (data.department ? data.department + ' Department' : 'YSE') +
    ' of the YSE ' + data.branch + '.\n\n' +
    'The interview will last around 15 to 30 minutes, and we would like to identify the most fitting ' + fd.position + ' for the Youth Society for Education.\n\n' +
    'We would like to set up an online interview for you on ' +
    fd.interviewDay + ', ' + fd.interviewDate + ', at ' + fd.interviewTime +
    ', via Google Meet. Please be prepared for your interview.\n\n' +
    'YSE Interview with ' + data.name + ' [' + fd.position + ']\n' +
    fd.interviewDay + ', ' + fd.interviewDate + ' · ' + fd.interviewTime + '–' + fd.interviewEndTime + '\n' +
    'Time zone: ' + fd.timezone + '\n' +
    'Google Meet joining info\n' +
    'Video call link: ' + fd.meetLink + '\n\n' +
    'Please reply to this email if you accept this Interview no later than ' + fd.replyDeadline + '. ' +
    'If you want to cancel the interview for any reason, please email us and cancel at least one hour ahead of the interview. ' +
    'Failure to comply on time would mean giving up your opportunity for the interview.\n\n' +
    '--\nSincerely,\nYSE HR Team\n' + HR_CONFIG.ORG_NAME + '\n' +
    HR_CONFIG.HR_EMAIL + '\n' + HR_CONFIG.WEBSITE;

  GmailApp.sendEmail(data.email, subject, body, {
    name:    HR_CONFIG.HR_NAME,
    replyTo: HR_CONFIG.HR_EMAIL,
    cc:      fd.ccHr ? HR_CONFIG.HR_EMAIL : '',
  });

  // Write the interview date back into the sheet
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(HR_CONFIG.SHEET_NAME)
    .getRange(fd.row, HR_CONFIG.COL.INTERVIEW_DATE)
    .setValue(fd.interviewDay + ', ' + fd.interviewDate + ' at ' + fd.interviewTime);

  return { success: true };
}


//  PROCESS 2 OFFER LETTER

function openOfferDialog() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName(HR_CONFIG.SHEET_NAME);
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please click on a candidate\'s row first, then try again.');
    return;
  }

  var data     = getRowData(row);
  var template = HtmlService.createTemplateFromFile('OfferDialog');
  template.data   = data;
  template.config = HR_CONFIG;
  var html = template.evaluate().setWidth(560).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(
    html, '📄 Offer Letter — ' + data.name
  );
}

// Called by OfferDialog.html when HR clicks "Send"

function sendOfferEmail(fd) {
  var data = getRowData(fd.row);
  if (!data.email) throw new Error('No email address found for this candidate.');

  var subject = 'YSE Offer Letter — ' + data.name + ' [' + fd.position + ']';

  var body =
    'Dear ' + data.name + ',\n\n' +
    'Congratulations on completing the application and interview processes! ' +
    'After thorough consideration of your application and interview, we have decided to offer you a position as a ' +
    fd.position + ' of the ' + fd.branch + ' starting from ' + fd.startDate + '.\n\n' +
    'You are truly a dedicated candidate who seeks not only to improve your own skills but also to help your colleagues and contribute back to the community. ' +
    'We, from the YSE HR team, would like to extend our congratulations on your accomplishments and welcome you to the YSE.\n\n' +
    'Please reply to this email if you accept the position no later than ' + fd.replyDeadline + '. ' +
    'If you accept our offer, it would be helpful to reply to our email by providing your Facebook and Telegram account (Name and Profile link).\n\n' +
    '--\nSincerely,\nYSE HR Team\n' + HR_CONFIG.ORG_NAME + '\n' +
    HR_CONFIG.HR_EMAIL + '\n' + HR_CONFIG.WEBSITE;

  var attachments = [];

  // Generate PDF offer letter from Google Docs template (if set up)

  if (fd.attachPdf) {
    var templateId = HR_CONFIG.DRIVE_FILES.OFFER_LETTER_TEMPLATE;
    if (templateId && templateId.indexOf('PASTE_') === -1) {
      try {
        attachments.push(generateOfferLetterPdf(fd, data));
      } catch(err) {
        Logger.log('Offer letter PDF error: ' + err.message);
        throw new Error('Could not generate offer letter PDF: ' + err.message);
      }
    }
  }

  GmailApp.sendEmail(data.email, subject, body, {
    name:        HR_CONFIG.HR_NAME,
    replyTo:     HR_CONFIG.HR_EMAIL,
    cc:          fd.ccHr ? HR_CONFIG.HR_EMAIL : '',
    attachments: attachments,
  });

  // Mark offer letter as sent in the sheet

  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(HR_CONFIG.SHEET_NAME)
    .getRange(fd.row, HR_CONFIG.COL.OFFER_LETTER_STATUS)
    .setValue('Sent');

  return { success: true };
}

function generateOfferLetterPdf(fd, data) {
  var templateFile = DriveApp.getFileById(HR_CONFIG.DRIVE_FILES.OFFER_LETTER_TEMPLATE);
  var copy         = templateFile.makeCopy('Offer Letter - ' + data.name, DriveApp.getRootFolder());
  var doc          = DocumentApp.openById(copy.getId());
  var body         = doc.getBody();

  // Replace placeholders <- your Google Doc template should contain these exact tags
  body.replaceText('\\{\\{DATE\\}\\}',       fd.letterDate  || '');
  body.replaceText('\\{\\{NAME\\}\\}',       data.name      || '');
  body.replaceText('\\{\\{POSITION\\}\\}',   fd.position    || '');
  body.replaceText('\\{\\{BRANCH\\}\\}',     fd.branch      || '');
  body.replaceText('\\{\\{START_DATE\\}\\}', fd.startDate   || '');

  doc.saveAndClose();

  var pdf = DriveApp.getFileById(copy.getId())
              .getAs('application/pdf')
              .setName('YSE Offer Letter - ' + data.name + '.pdf');

  DriveApp.getFileById(copy.getId()).setTrashed(true);  // clean up temp copy
  return pdf;
}


//  PROCESS 3 <- ONBOARDING PACKAGE

// Opens dialog automatically (called by trigger)

function openOnboardingDialog(row) {
  var data     = getRowData(row);
  var template = HtmlService.createTemplateFromFile('OnboardingDialog');
  template.data   = data;
  template.config = HR_CONFIG;
  var html = template.evaluate().setWidth(580).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(
    html, '📦 Onboarding Package — ' + data.name
  );
}

// Opens dialog from the menu (uses the currently selected row)

function openOnboardingDialogManual() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet()
                .getSheetByName(HR_CONFIG.SHEET_NAME);
  var row = sheet.getActiveCell().getRow();
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Please click on a candidate\'s row first, then try again.');
    return;
  }
  openOnboardingDialog(row);
}

// Called by OnboardingDialog.html when HR clicks "Send"

function sendOnboardingEmail(fd) {
  var data = getRowData(fd.row);
  if (!data.email) throw new Error('No email address found for this candidate.');

  var subject = 'Welcome to YSE — Onboarding Package for ' + data.name;

  var body =
    'Dear ' + data.name + ',\n\n' +
    'First of all, Congratulations and Welcome to YSE Team. ' +
    'The documents of YSE Employee Handbooks, By-laws for officers and Officer Onboarding Packet are attached below. ' +
    'Please read carefully and make sure to understand the rules and regulations practices in YSE. ' +
    'And please sign the acknowledgement, onboarding packet, and By-laws, then send us back no later than ' + fd.replyDeadline + '.\n\n' +
    'We also would like to give your YSE mail account provided by YSE, which you will have to use for further organizational contact and Google Drive access.\n' +
    'YSE account username - ' + fd.yseEmail    + '\n' +
    'Password - '             + fd.ysePassword + '\n\n' +
    'If you have any difficulties, please feel free to contact the HR Department.\n\n' +
    '--\nSincerely,\nYSE HR Team\n' + HR_CONFIG.ORG_NAME + '\n' +
    HR_CONFIG.HR_EMAIL + '\n' + HR_CONFIG.WEBSITE;

  var attachments = [];
  var files       = HR_CONFIG.DRIVE_FILES;

  // Attach Employee Handbook PDF
  if (files.EMPLOYEE_HANDBOOK && files.EMPLOYEE_HANDBOOK.indexOf('PASTE_') === -1) {
    try {
      attachments.push(
        DriveApp.getFileById(files.EMPLOYEE_HANDBOOK)
          .getAs('application/pdf')
          .setName('YSE Employee Handbook.pdf')
      );
    } catch(e) { Logger.log('Handbook: ' + e.message); }
  }

  // Attach Bylaws PDF
  if (files.BYLAWS && files.BYLAWS.indexOf('PASTE_') === -1) {
    try {
      attachments.push(
        DriveApp.getFileById(files.BYLAWS)
          .getAs('application/pdf')
          .setName('YSE Bylaws.pdf')
      );
    } catch(e) { Logger.log('Bylaws: ' + e.message); }
  }

  // Attach Acknowledgement of Receipt PDF
  if (files.ACKNOWLEDGEMENT && files.ACKNOWLEDGEMENT.indexOf('PASTE_') === -1) {
    try {
      attachments.push(
        DriveApp.getFileById(files.ACKNOWLEDGEMENT)
          .getAs('application/pdf')
          .setName('YSE Acknowledgement of Receipt.pdf')
      );
    } catch(e) { Logger.log('Acknowledgement: ' + e.message); }
  }

  // Generate personalised Onboarding Packet PDF
  if (files.ONBOARDING_PACKET_TEMPLATE && files.ONBOARDING_PACKET_TEMPLATE.indexOf('PASTE_') === -1) {
    try {
      attachments.push(generateOnboardingPacketPdf(fd, data));
    } catch(e) { Logger.log('Onboarding packet: ' + e.message); }
  }

  GmailApp.sendEmail(data.email, subject, body, {
    name:        HR_CONFIG.HR_NAME,
    replyTo:     HR_CONFIG.HR_EMAIL,
    cc:          fd.ccHr ? HR_CONFIG.HR_EMAIL : '',
    attachments: attachments,
  });

  return { success: true };
}

function generateOnboardingPacketPdf(fd, data) {
  var templateFile = DriveApp.getFileById(HR_CONFIG.DRIVE_FILES.ONBOARDING_PACKET_TEMPLATE);
  var copy         = templateFile.makeCopy('Onboarding Packet - ' + data.name, DriveApp.getRootFolder());
  var doc          = DocumentApp.openById(copy.getId());
  var body         = doc.getBody();

  // Replace placeholders — your Google Doc template should contain these exact tags
  body.replaceText('\\{\\{OFFICER_NAME\\}\\}',        data.name             || '');
  body.replaceText('\\{\\{DEPARTMENT\\}\\}',           data.department       || '');
  body.replaceText('\\{\\{POSITION\\}\\}',             data.position         || '');
  body.replaceText('\\{\\{PERSONAL_EMAIL\\}\\}',       data.email            || '');
  body.replaceText('\\{\\{START_DATE\\}\\}',           fd.startDate          || '');
  body.replaceText('\\{\\{YSE_EMAIL\\}\\}',            fd.yseEmail           || '');
  body.replaceText('\\{\\{YSE_PASSWORD\\}\\}',         fd.ysePassword        || '');
  body.replaceText('\\{\\{SUPERVISOR_NAME\\}\\}',      fd.supervisorName     || '');
  body.replaceText('\\{\\{SUPERVISOR_POSITION\\}\\}',  fd.supervisorPosition || '');
  body.replaceText('\\{\\{SUPERVISOR_PHONE\\}\\}',     fd.supervisorPhone    || '');
  body.replaceText('\\{\\{SUPERVISOR_EMAIL\\}\\}',     fd.supervisorEmail    || '');

  doc.saveAndClose();

  var pdf = DriveApp.getFileById(copy.getId())
              .getAs('application/pdf')
              .setName('YSE Onboarding Packet - ' + data.name + '.pdf');

  DriveApp.getFileById(copy.getId()).setTrashed(true);
  return pdf;
}