const SPREADSHEET_ID = '1DlV-ZJOy6HqPMeBqO8-j1bvhh8cBi84aoagMJ8aNoD4';

const INTAKE_SHEET = 'CV Intake';
const CONFIG_SHEET = 'CV Automation Config';
const RESPONSE_SHEET = 'Candidate Interest Responses';
const EMPLOYMENT_HISTORY_SHEET = 'Employment History';

const MAX_NEW_CVS_PER_RUN = 4;

function setupCvAutomation() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');

  if (!apiKey) {
    throw new Error(
      'OPENAI_API_KEY is missing. Add it under Project Settings → Script Properties.'
    );
  }

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'scanCvFolders')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('scanCvFolders')
    .timeBased()
    .everyMinutes(5)
    .create();

  scanCvFolders();
}

function scanCvFolders() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) return;

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const intakeSheet = ss.getSheetByName(INTAKE_SHEET);
    const configSheet = ss.getSheetByName(CONFIG_SHEET);

    if (!intakeSheet || !configSheet) {
      throw new Error('CV Intake or CV Automation Config sheet is missing.');
    }

    const lastRow = intakeSheet.getLastRow();

    const processedIds = new Set(
      lastRow > 1
        ? intakeSheet
            .getRange(2, 1, lastRow - 1, 1)
            .getDisplayValues()
            .flat()
            .filter(String)
        : []
    );

    const config = configSheet.getDataRange().getDisplayValues().slice(1);

    let processedThisRun = 0;

    for (const row of config) {
      if (processedThisRun >= MAX_NEW_CVS_PER_RUN) break;

      const position = row[0];
      const folderId = row[1];
      const screeningFocus = row[3];
      const verifyOnCall = row[4];
      const active = row[5];

      if (
        !position ||
        !folderId ||
        String(active).toLowerCase() !== 'yes'
      ) {
        continue;
      }

      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();

      while (
        files.hasNext() &&
        processedThisRun < MAX_NEW_CVS_PER_RUN
      ) {
        const file = files.next();

        if (processedIds.has(file.getId())) continue;

        processCvFile_(
          file,
          position,
          screeningFocus,
          verifyOnCall,
          intakeSheet
        );

        processedIds.add(file.getId());
        processedThisRun++;
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function processCvFile_(
  file,
  position,
  screeningFocus,
  verifyOnCall,
  sheet
) {
  const now = new Date();
  let openAiFileId = null;

  try {
    const blob = prepareCvBlob_(file);
    openAiFileId = uploadFileToOpenAI_(blob);

    const cv = extractCvData_(
      openAiFileId,
      position,
      screeningFocus,
      verifyOnCall
    );

    const interestLink = createCandidateInterestLink_(
      file.getId(),
      cv.candidate_name || '',
      position
    );

    appendIntakeRow_(sheet, [
      file.getId(),                      // A Intake ID
      file.getName(),                    // B CV File Name
      file.getUrl(),                     // C CV Link
      position,                          // D Position
      cv.candidate_name || '',           // E Candidate Name
      cv.mobile_number || '',            // F Mobile Number
      cv.candidate_location || '',        // G Candidate Location
      cv.total_experience || '',          // H Total Experience
      cv.relevant_experience || '',       // I Relevant Experience
      cv.current_last_designation || '', // J Current / Last Designation
      cv.current_last_company || '',     // K Current / Last Company
      cv.education || '',                // L Education
      'Pending',                         // M Interest Status
      interestLink,                      // N Interest Form Link
      'Not Sent',                        // O WhatsApp Status
      'Pending Candidate Response',      // P Screening Result
      'CV parsed automatically. Awaiting candidate interest.', // Q Screening Summary
      cv.cv_missing_or_concern || '',    // R Missing / Concern
      cv.cv_followup_questions || '',    // S Call Questions
      '',                                // T Recruiter
      '',                                // U Source
      now,                               // V Added On
      now,                               // W Last Processed On
      'Not Added',                       // X Candidate Tracker Status
      ''                                 // Y Notes
    ]);

    saveEmploymentHistory_(
      file.getId(),
      cv.candidate_name || '',
      cv.work_history || []
    );
  } catch (error) {
    appendIntakeRow_(sheet, [
      file.getId(),                      // A
      file.getName(),                    // B
      file.getUrl(),                     // C
      position,                          // D
      '', '', '', '', '', '', '', '',   // E:L
      'Error',                           // M
      '',                                // N
      'Not Sent',                        // O
      'Error',                           // P
      '', '', '', '', '',                // Q:U
      now,                               // V
      now,                               // W
      'Not Added',                       // X
      'CV processing error: ' + error.message // Y
    ]);

    console.error(
      'CV processing failed: ' +
      file.getName() +
      ' | ' +
      error.message
    );
  } finally {
    if (openAiFileId) deleteOpenAIFile_(openAiFileId);
  }
}

function prepareCvBlob_(file) {
  const mime = file.getMimeType();

  if (mime === MimeType.GOOGLE_DOCS) {
    return file
      .getAs(MimeType.PDF)
      .setName(file.getName() + '.pdf');
  }

  return file.getBlob().setName(file.getName());
}

function uploadFileToOpenAI_(blob) {
  const apiKey = getOpenAiKey_();

  const response = UrlFetchApp.fetch(
    'https://api.openai.com/v1/files',
    {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + apiKey
      },
      payload: {
        purpose: 'user_data',
        file: blob
      },
      muteHttpExceptions: true
    }
  );

  const status = response.getResponseCode();
  const text = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(
      'OpenAI file upload failed (' + status + '): ' + text
    );
  }

  const json = JSON.parse(text);

  if (!json.id) {
    throw new Error('OpenAI did not return a File ID.');
  }

  return json.id;
}

function extractCvData_(
  openAiFileId,
  position,
  screeningFocus,
  verifyOnCall
) {
  const apiKey = getOpenAiKey_();

  const model =
    PropertiesService
      .getScriptProperties()
      .getProperty('OPENAI_MODEL') ||
    'gpt-5.6-luna';

  const prompt = `
You are extracting recruitment information from a candidate CV.

POSITION BEING CONSIDERED:
${position}

JOB-RELEVANT SCREENING FOCUS:
${screeningFocus}

INFORMATION THAT MAY LATER NEED TO BE VERIFIED:
${verifyOnCall}

RULES:

1. Extract only information actually supported by the CV.
2. Never invent missing information.
3. If information is unavailable, return an empty string.
4. Do NOT make a final hiring or rejection decision.
5. Do NOT infer or evaluate age, gender, religion, caste, marital status,
   health, disability, ethnicity, or other protected/personal characteristics.
6. Concerns must relate only to job-relevant information that is missing,
   unclear, or needs verification.
7. Relevant experience means experience reasonably relevant to the position.
8. For mobile number:
   - return digits only
   - if a clear Indian +91 number is present, return the final 10 digits.
9. Follow-up questions must be concise.
10. Maximum 5 follow-up questions.
11. Do not penalize a candidate simply because information is missing from the CV.
    Flag it for verification instead.
12. Keep extracted information concise and practical for an HR recruiter.
13. Extract the candidate's complete employment history, newest/current role first.
14. For each employment entry, capture company, role, start date, end date, tenure, and location only when supported by the CV.
15. Preserve dates as written when exact month/year is unavailable. Use "Present" or "Current" only when the CV clearly says so.
16. Calculate or summarize tenure only when the source supports it. If the CV itself states a tenure that conflicts with dates, preserve the supported dates and mention the conflict in notes rather than silently correcting it.
17. Do not include education, personal details, protected characteristics, or non-employment activities in work_history unless they are clearly presented as employment.
`;

  const schema = {
    type: 'object',
    properties: {
      candidate_name: { type: 'string' },
      mobile_number: { type: 'string' },
      candidate_location: { type: 'string' },
      total_experience: { type: 'string' },
      relevant_experience: { type: 'string' },
      current_last_designation: { type: 'string' },
      current_last_company: { type: 'string' },
      education: { type: 'string' },
      work_history: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            company: { type: 'string' },
            role: { type: 'string' },
            start_date: { type: 'string' },
            end_date: { type: 'string' },
            tenure: { type: 'string' },
            location: { type: 'string' },
            notes: { type: 'string' }
          },
          required: [
            'company',
            'role',
            'start_date',
            'end_date',
            'tenure',
            'location',
            'notes'
          ],
          additionalProperties: false
        }
      },
      cv_missing_or_concern: { type: 'string' },
      cv_followup_questions: { type: 'string' }
    },
    required: [
      'candidate_name',
      'mobile_number',
      'candidate_location',
      'total_experience',
      'relevant_experience',
      'current_last_designation',
      'current_last_company',
      'education',
      'work_history',
      'cv_missing_or_concern',
      'cv_followup_questions'
    ],
    additionalProperties: false
  };

  const payload = {
    model: model,
    store: false,
    reasoning: {
      effort: 'none'
    },
    input: [
      {
        role: 'user',
        content: [
          {
            type: 'input_file',
            file_id: openAiFileId
          },
          {
            type: 'input_text',
            text: prompt
          }
        ]
      }
    ],
    text: {
      format: {
        type: 'json_schema',
        name: 'cv_intake',
        strict: true,
        schema: schema
      }
    },
    max_output_tokens: 2400
  };

  const response = UrlFetchApp.fetch(
    'https://api.openai.com/v1/responses',
    {
      method: 'post',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );

  const status = response.getResponseCode();
  const text = response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(
      'OpenAI CV extraction failed (' + status + '): ' + text
    );
  }

  const responseJson = JSON.parse(text);
  const outputText = getOutputText_(responseJson);

  if (!outputText) {
    throw new Error('No structured CV data returned by OpenAI.');
  }

  return JSON.parse(outputText);
}

function saveEmploymentHistory_(intakeId, candidateName, workHistory) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(EMPLOYMENT_HISTORY_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(EMPLOYMENT_HISTORY_SHEET);
    sheet.hideSheet();
  }

  const headers = [
    'Intake ID',
    'Candidate Name',
    'Sequence',
    'Company',
    'Role',
    'Start Date',
    'End Date',
    'Tenure',
    'Location',
    'Display Line'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existingHeaders = sheet
      .getRange(1, 1, 1, headers.length)
      .getDisplayValues()[0];

    if (existingHeaders.join('|') !== headers.join('|')) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }

  const lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    const ids = sheet
      .getRange(2, 1, lastRow - 1, 1)
      .getDisplayValues()
      .flat();

    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i] || '').trim() === String(intakeId || '').trim()) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  if (!Array.isArray(workHistory) || workHistory.length === 0) return;

  const rows = workHistory
    .filter(job => job && (job.company || job.role))
    .map((job, index) => {
      const company = String(job.company || '').trim();
      const role = String(job.role || '').trim();
      const startDate = String(job.start_date || '').trim();
      const endDate = String(job.end_date || '').trim();
      const tenure = String(job.tenure || '').trim();
      const location = String(job.location || '').trim();

      const dateRange = [startDate, endDate].filter(String).join(' - ');
      const roleLine = [company, role].filter(String).join(' — ');
      const details = [dateRange, tenure].filter(String).join(' · ');
      const displayLine = details ? roleLine + ' · ' + details : roleLine;

      return [
        intakeId,
        candidateName,
        index + 1,
        company,
        role,
        startDate,
        endDate,
        tenure,
        location,
        displayLine
      ];
    });

  if (rows.length) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length)
      .setValues(rows);
  }
}

function getOutputText_(response) {
  if (response.output_text) return response.output_text;
  if (!response.output) return '';

  for (const item of response.output) {
    if (!item.content) continue;

    for (const content of item.content) {
      if (
        content.type === 'output_text' &&
        content.text
      ) {
        return content.text;
      }
    }
  }

  return '';
}

function deleteOpenAIFile_(fileId) {
  try {
    UrlFetchApp.fetch(
      'https://api.openai.com/v1/files/' + fileId,
      {
        method: 'delete',
        headers: {
          Authorization: 'Bearer ' + getOpenAiKey_()
        },
        muteHttpExceptions: true
      }
    );
  } catch (error) {
    console.warn(
      'Could not delete temporary OpenAI file: ' + fileId
    );
  }
}

function getOpenAiKey_() {
  const key = PropertiesService
    .getScriptProperties()
    .getProperty('OPENAI_API_KEY');

  if (!key) {
    throw new Error('OPENAI_API_KEY is not configured.');
  }

  return key;
}

function getCandidateFormUrl_() {
  const url = PropertiesService
    .getScriptProperties()
    .getProperty('CANDIDATE_FORM_URL');

  if (!url) {
    throw new Error(
      'CANDIDATE_FORM_URL is missing from Script Properties.'
    );
  }

  return url;
}

function createCandidateInterestLink_(
  intakeId,
  candidateName,
  position
) {
  if (!intakeId || !candidateName || !position) {
    return '';
  }

  const formUrl = getCandidateFormUrl_();

  return (
    formUrl +
    '?id=' + encodeURIComponent(intakeId) +
    '&name=' + encodeURIComponent(candidateName) +
    '&position=' + encodeURIComponent(position)
  );
}

function generateCandidateInterestLinks() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(INTAKE_SHEET);

  if (!sheet) {
    throw new Error('CV Intake sheet is missing.');
  }

  const formUrl = getCandidateFormUrl_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const data = sheet
    .getRange(2, 1, lastRow - 1, 25)
    .getValues();

  data.forEach(function(row, index) {
    const intakeId = row[0];       // A
    const position = row[3];       // D
    const candidateName = row[4];  // E
    const existingLink = String(row[13] || ''); // N

    const hasCorrectFormLink =
      existingLink.startsWith(formUrl);

    if (
      intakeId &&
      candidateName &&
      position &&
      !hasCorrectFormLink
    ) {
      const uniqueLink =
        createCandidateInterestLink_(
          intakeId,
          candidateName,
          position
        );

      sheet
        .getRange(index + 2, 14)
        .setValue(uniqueLink);
    }
  });
}

function doGet(e) {

  // -----------------------------------------
  // META WHATSAPP WEBHOOK VERIFICATION
  // -----------------------------------------
  if (
    e &&
    e.parameter &&
    e.parameter['hub.mode'] === 'subscribe'
  ) {

    const verifyToken =
      PropertiesService
        .getScriptProperties()
        .getProperty('WHATSAPP_WEBHOOK_VERIFY_TOKEN');

    const receivedToken =
      e.parameter['hub.verify_token'] || '';

    const challenge =
      e.parameter['hub.challenge'] || '';

    if (
      verifyToken &&
      receivedToken === verifyToken
    ) {
      return ContentService
        .createTextOutput(challenge)
        .setMimeType(ContentService.MimeType.TEXT);
    }

    return ContentService
      .createTextOutput('Verification failed')
      .setMimeType(ContentService.MimeType.TEXT);
  }


  // -----------------------------------------
  // EXISTING CANDIDATE INTEREST FORM
  // -----------------------------------------

  const template =
    HtmlService.createTemplateFromFile('CandidateForm');

  template.intakeId =
    e && e.parameter
      ? (e.parameter.id || '')
      : '';

  template.candidateName =
    e && e.parameter
      ? (e.parameter.name || '')
      : '';

  template.position =
    e && e.parameter
      ? (e.parameter.position || '')
      : '';

  template.formAction =
    ScriptApp.getService().getUrl() ||
    getCandidateFormUrl_();

  return template
    .evaluate()
    .setTitle('Brynex Career Interest Form');
}
function doPost(e) {

  try {

    // -----------------------------------------
    // WHATSAPP WEBHOOK POST
    // -----------------------------------------

    const contentType =
      e &&
      e.postData &&
      e.postData.type
        ? String(e.postData.type).toLowerCase()
        : '';

    if (
      contentType.includes('application/json')
    ) {

      const payload =
        JSON.parse(e.postData.contents || '{}');

      handleWhatsAppWebhook_(payload);

      const debugSheet =
  SpreadsheetApp
    .openById(SPREADSHEET_ID)
    .getSheetByName('Webhook Debug');

if (debugSheet) {
  debugSheet.appendRow([
    new Date(),
    '',
    '',
    '',
    '',
    JSON.stringify(payload)
  ]);
}

      return ContentService
        .createTextOutput('EVENT_RECEIVED')
        .setMimeType(ContentService.MimeType.TEXT);
    }


    // -----------------------------------------
    // CANDIDATE INTEREST FORM POST
    // -----------------------------------------

    if (!e || !e.parameter) {
      return buildMessagePage_(
        'Submission could not be processed.',
        'Please reopen the recruitment link and try again.'
      );
    }

    const data = e.parameter;

    if (!data.intakeId) {
      return buildMessagePage_(
        'Candidate reference is missing.',
        'Please reopen the original recruitment link and try again.'
      );
    }

    const ss =
      SpreadsheetApp.openById(SPREADSHEET_ID);

    const responseSheet =
      ss.getSheetByName(RESPONSE_SHEET);

    const intakeSheet =
      ss.getSheetByName(INTAKE_SHEET);

    if (!responseSheet || !intakeSheet) {
      throw new Error(
        'Required recruitment sheet is missing.'
      );
    }

    const lastRow =
      intakeSheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(
        'Candidate could not be found.'
      );
    }

    const ids =
      intakeSheet
        .getRange(2, 1, lastRow - 1, 1)
        .getDisplayValues()
        .flat();

    const index =
      ids.indexOf(String(data.intakeId));

    if (index === -1) {
      throw new Error(
        'Candidate could not be found.'
      );
    }

    const intakeRow =
      index + 2;

    const candidateName =
      intakeSheet
        .getRange(intakeRow, 5)
        .getValue();

    const mobileNumber =
      intakeSheet
        .getRange(intakeRow, 6)
        .getValue();

    const position =
      intakeSheet
        .getRange(intakeRow, 4)
        .getValue();

    const now = new Date();

    responseSheet.appendRow([
      now,                           // A Timestamp
      data.intakeId || '',           // B Intake ID
      candidateName || '',           // C Candidate Name
      mobileNumber || '',            // D Mobile Number
      position || '',                // E Position
      data.interested || '',         // F Interested?
      data.currentLocation || '',    // G Current Location
      data.designation || '',        // H Current Role
      data.totalExperience || '',    // I Total Experience
      data.relevantExperience || '', // J Relevant Experience
      data.currentSalary || '',      // K Current Salary
      data.expectedSalary || '',     // L Expected Salary
      data.noticePeriod || '',       // M Notice / Joining
      data.relocate || '',           // N Relocate?
      data.preferredLocation || '',  // O Preferred Location
      data.remarks || '',            // P Remarks
      'New'                          // Q Processing Status
    ]);

    intakeSheet
      .getRange(intakeRow, 13)
      .setValue(
        data.interested || 'Submitted'
      );

    intakeSheet
      .getRange(intakeRow, 23)
      .setValue(now);

    return buildMessagePage_(
      'Thank you.',
      'Your response has been submitted successfully. Our recruitment team will review your profile and contact you if required.'
    );

  } catch (error) {

    console.error(
      'doPost error: ' +
      error.message
    );

    return ContentService
      .createTextOutput('OK')
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function buildMessagePage_(title, message) {
  const safeTitle = escapeHtml_(title);
  const safeMessage = escapeHtml_(message);

  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <title>${safeTitle}</title>
      </head>
      <body style="
        font-family:Arial,sans-serif;
        max-width:620px;
        margin:60px auto;
        padding:20px;
        text-align:center;
        line-height:1.5;
      ">
        <h2>${safeTitle}</h2>
        <p>${safeMessage}</p>
      </body>
    </html>
  `);
}

function escapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
/**
 * WHATSAPP CLOUD API
 * Phase 1: controlled connection test
 */

function testWhatsAppConnection() {

  const props = PropertiesService.getScriptProperties();

  const testNumber =
    props.getProperty('WHATSAPP_TEST_TO');

  const formUrl =
    props.getProperty('CANDIDATE_FORM_URL');

  if (!testNumber) {
    throw new Error(
      'WHATSAPP_TEST_TO is missing from Script Properties.'
    );
  }

  if (!formUrl) {
    throw new Error(
      'CANDIDATE_FORM_URL is missing from Script Properties.'
    );
  }

  const testLink =
    formUrl +
    '?id=TEST-CANDIDATE' +
    '&name=' + encodeURIComponent('Test Candidate') +
    '&position=' + encodeURIComponent('ASM');

  const result = sendWhatsAppTemplate_(
    testNumber,
    'Test Candidate',
    'ASM',
    testLink
  );

  console.log(JSON.stringify(result));
}


/**
 * Sends the approved Brynex recruitment template.
 */
function sendWhatsAppTemplate_(
  mobileNumber,
  candidateName,
  position,
  interestFormLink
) {

  const props = PropertiesService.getScriptProperties();

  const phoneNumberId =
    props.getProperty('WHATSAPP_PHONE_NUMBER_ID');

  const token =
    props.getProperty('WHATSAPP_ACCESS_TOKEN');

  const templateName =
    props.getProperty('WHATSAPP_TEMPLATE_NAME');

  const language =
    props.getProperty('WHATSAPP_TEMPLATE_LANGUAGE') ||
    'en_US';

  if (!phoneNumberId) {
    throw new Error(
      'WHATSAPP_PHONE_NUMBER_ID is missing.'
    );
  }

  if (!token) {
    throw new Error(
      'WHATSAPP_ACCESS_TOKEN is missing.'
    );
  }

  if (!templateName) {
    throw new Error(
      'WHATSAPP_TEMPLATE_NAME is missing.'
    );
  }

  const cleanNumber =
    normalizeWhatsAppNumber_(mobileNumber);

  if (!cleanNumber) {
    throw new Error(
      'Invalid WhatsApp mobile number.'
    );
  }

  const endpoint =
    'https://graph.facebook.com/v26.0/' +
    phoneNumberId +
    '/messages';

  const payload = {

    messaging_product: 'whatsapp',

    recipient_type: 'individual',

    to: cleanNumber,

    type: 'template',

    template: {

      name: templateName,

      language: {
        code: language
      },

      components: [

        {
          type: 'body',

          parameters: [

            {
              type: 'text',
              text: String(candidateName || '')
            },

            {
              type: 'text',
              text: String(position || '')
            },

            {
              type: 'text',
              text: String(interestFormLink || '')
            }

          ]
        }

      ]
    }
  };

  const response = UrlFetchApp.fetch(
    endpoint,
    {

      method: 'post',

      contentType: 'application/json',

      headers: {
        Authorization: 'Bearer ' + token
      },

      payload: JSON.stringify(payload),

      muteHttpExceptions: true

    }
  );

  const status =
    response.getResponseCode();

  const text =
    response.getContentText();

  let result;

  try {
    result = JSON.parse(text);
  } catch (error) {
    result = {
      rawResponse: text
    };
  }

  if (status < 200 || status >= 300) {

    throw new Error(
      'WhatsApp API failed (' +
      status +
      '): ' +
      text
    );
  }

  return result;
}


/**
 * Converts Indian mobile numbers into WhatsApp API format.
 *
 * 9876543210  -> 919876543210
 * +91...      -> 91...
 */
function normalizeWhatsAppNumber_(number) {

  let digits =
    String(number || '')
      .replace(/\D/g, '');

  if (!digits) {
    return '';
  }

  if (
    digits.length === 10 &&
    /^[6-9]/.test(digits)
  ) {

    digits = '91' + digits;

  }

  if (
    digits.length === 12 &&
    digits.startsWith('91')
  ) {

    return digits;

  }

  return digits;
}
/**
 * Sends WhatsApp for ONE specific CV Intake row.
 * Use this for controlled testing before enabling automation.
 */
function sendWhatsAppForCvIntakeRow(rowNumber) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(INTAKE_SHEET);

  if (!sheet) {
    throw new Error('CV Intake sheet is missing.');
  }

  if (!rowNumber || rowNumber < 2) {
    throw new Error('Invalid CV Intake row number.');
  }

  const row = sheet
    .getRange(rowNumber, 1, 1, 25)
    .getValues()[0];

  const candidateName = row[4];   // E
  const mobileNumber = row[5];    // F
  const position = row[3];        // D
  const interestStatus = row[12]; // M
  const formLink = row[13];       // N
  const whatsappStatus = row[14]; // O

  if (!candidateName) {
    throw new Error('Candidate name is missing.');
  }

  if (!mobileNumber) {
    throw new Error('Candidate mobile number is missing.');
  }

  if (!position) {
    throw new Error('Position is missing.');
  }

  if (!formLink) {
    throw new Error('Candidate interest form link is missing.');
  }

  if (
    String(interestStatus).toLowerCase() !== 'pending'
  ) {
    throw new Error(
      'Candidate Interest Status is not Pending.'
    );
  }

  if (
    String(whatsappStatus).toLowerCase() === 'sent'
  ) {
    throw new Error(
      'WhatsApp has already been sent to this candidate.'
    );
  }

  try {

    const result = sendWhatsAppTemplate_(
      mobileNumber,
      candidateName,
      position,
      formLink
    );

    // O = WhatsApp Status
    sheet
      .getRange(rowNumber, 15)
      .setValue('Sent');

    // W = Last Processed On
    sheet
      .getRange(rowNumber, 23)
      .setValue(new Date());

    console.log(
      'WhatsApp sent to ' +
      candidateName +
      ' | ' +
      JSON.stringify(result)
    );

    return result;

  } catch (error) {

    sheet
      .getRange(rowNumber, 15)
      .setValue('Failed');

    sheet
      .getRange(rowNumber, 25)
      .setValue(
        'WhatsApp error: ' + error.message
      );

      throw error;
  }
}
function testRealCandidateWhatsApp() {
  sendWhatsAppForCvIntakeRow(5);
}
function sendPendingWhatsAppMessages() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(INTAKE_SHEET);

  if (!sheet) {
    throw new Error('CV Intake sheet is missing.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 25).getValues();

  // Count pending rows by mobile number
  const mobileCounts = {};

  data.forEach(row => {
    const mobile = normalizeWhatsAppNumber_(row[5]);
    const interestStatus = String(row[12] || '').trim();
    const whatsappStatus = String(row[14] || '').trim();

    if (
      mobile &&
      interestStatus === 'Pending' &&
      whatsappStatus === 'Not Sent'
    ) {
      mobileCounts[mobile] = (mobileCounts[mobile] || 0) + 1;
    }
  });

  data.forEach((row, index) => {

    const rowNumber = index + 2;

    const candidateName = row[4];
    const mobileNumber = row[5];
    const position = row[3];
    const interestStatus = String(row[12] || '').trim();
    const formLink = row[13];
    const whatsappStatus = String(row[14] || '').trim();

    if (
      !candidateName ||
      !mobileNumber ||
      !position ||
      !formLink
    ) {
      return;
    }

    if (interestStatus !== 'Pending') return;
    if (whatsappStatus !== 'Not Sent') return;

    const cleanMobile = normalizeWhatsAppNumber_(mobileNumber);

    // Prevent sending multiple messages to same mobile automatically
    if (mobileCounts[cleanMobile] > 1) {

      sheet
        .getRange(rowNumber, 15)
        .setValue('Review - Duplicate Mobile');

      return;
    }

    try {

      sendWhatsAppTemplate_(
        mobileNumber,
        candidateName,
        position,
        formLink
      );

      sheet
        .getRange(rowNumber, 15)
        .setValue('Sent');

      sheet
        .getRange(rowNumber, 23)
        .setValue(new Date());

    } catch (error) {

      sheet
        .getRange(rowNumber, 15)
        .setValue('Failed');

      sheet
        .getRange(rowNumber, 25)
        .setValue(
          'WhatsApp error: ' + error.message
        );
    }

  });
}
/**
 * Processes new candidate interest responses for ASM.
 *
 * Output in CV Intake:
 * P = Screening Result
 * Q = Screening Summary
 * R = Missing / Concern
 * S = Call Questions
 *
 * Candidate Interest Responses:
 * Q = Processed?
 */
function processNewCandidateInterestResponses() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const responseSheet = ss.getSheetByName(RESPONSE_SHEET);
  const intakeSheet = ss.getSheetByName(INTAKE_SHEET);

  if (!responseSheet || !intakeSheet) {
    throw new Error('Required recruitment sheet is missing.');
  }

  const responseLastRow = responseSheet.getLastRow();
  const intakeLastRow = intakeSheet.getLastRow();

  if (responseLastRow < 2 || intakeLastRow < 2) return;

  const responses = responseSheet
    .getRange(2, 1, responseLastRow - 1, 17)
    .getValues();

  const intakeData = intakeSheet
    .getRange(2, 1, intakeLastRow - 1, 25)
    .getValues();

  // Build Intake ID → row lookup
  const intakeMap = {};

  intakeData.forEach((row, index) => {
    const intakeId = String(row[0] || '').trim();

    if (intakeId) {
      intakeMap[intakeId] = {
        rowNumber: index + 2,
        values: row
      };
    }
  });

  let processedCount = 0;

  responses.forEach((response, index) => {

    const responseRow = index + 2;

    const intakeId = String(response[1] || '').trim();     // B
    const position = String(response[4] || '').trim();     // E
    const interested = String(response[5] || '').trim();   // F
    const processed = String(response[16] || '').trim();   // Q

    // Only process genuinely new rows
    if (processed !== 'New') return;

    const intakeRecord = intakeMap[intakeId];

    if (!intakeRecord) {
      responseSheet
        .getRange(responseRow, 17)
        .setValue('Error - Intake ID not found');

      return;
    }

    const intakeRow = intakeRecord.rowNumber;
    const cv = intakeRecord.values;
    const intakePosition = String(cv[3] || '').trim().toUpperCase();
const responsePosition = String(position || '').trim().toUpperCase();
const interestAnswer = String(interested || '').trim().toLowerCase();

// Ignore old / malformed responses
if (
  !['yes', 'no'].includes(interestAnswer) ||
  responsePosition !== intakePosition
) {

  responseSheet
    .getRange(responseRow, 17)
    .setValue('Ignored - Invalid legacy response');

  return;
}

    // Candidate said No
if (interestAnswer === 'no') {

      intakeSheet.getRange(intakeRow, 16).setValue('Not Interested');
      intakeSheet.getRange(intakeRow, 17)
        .setValue('Candidate indicated that they are not interested in proceeding.');
      intakeSheet.getRange(intakeRow, 18).clearContent();
      intakeSheet.getRange(intakeRow, 19).clearContent();
      intakeSheet.getRange(intakeRow, 23).setValue(new Date());

      responseSheet
        .getRange(responseRow, 17)
        .setValue('Processed');

      processedCount++;
      return;
    }

    let result;

switch (intakePosition) {

  case 'ASM':
    result = screenAsmCandidate_(cv, response);
    break;

  case 'SALES':
    result = screenSalesCandidate_(cv, response);
    break;

  default:
    return;
}

    // CV Intake P:S
    intakeSheet
      .getRange(intakeRow, 16, 1, 4)
      .setValues([[
        result.screeningResult,
        result.summary,
        result.concerns,
        result.callQuestions
      ]]);

    // W = Last Processed On
    intakeSheet
      .getRange(intakeRow, 23)
      .setValue(new Date());

    // Candidate Interest Responses Q
    responseSheet
      .getRange(responseRow, 17)
      .setValue('Processed');

    processedCount++;
  });

  console.log(
    'Candidate screening completed. Responses processed: ' +
    processedCount
  );
}


/**
 * ASM screening rules.
 * This is decision support only.
 */
function screenAsmCandidate_(cv, response) {

  // ---- CV Intake values ----
  const cvLocation = String(cv[6] || '').trim();          // G
  const cvTotalExperience = String(cv[7] || '').trim();   // H
  const cvRelevantExperience = String(cv[8] || '').trim();// I
  const cvDesignation = String(cv[9] || '').trim();       // J
  const cvCompany = String(cv[10] || '').trim();          // K
  const cvExistingConcerns = String(cv[17] || '').trim(); // R

  // ---- Candidate response values ----
  const currentLocation = String(response[6] || '').trim();     // G
  const currentRole = String(response[7] || '').trim();         // H
  const totalExpText = String(response[8] || '').trim();        // I
  const relevantExpText = String(response[9] || '').trim();     // J
  const currentSalaryText = String(response[10] || '').trim();  // K
  const expectedSalaryText = String(response[11] || '').trim(); // L
  const noticeText = String(response[12] || '').trim();         // M
  const relocate = String(response[13] || '').trim();           // N
  const preferredLocation = String(response[14] || '').trim();  // O
  const candidateRemarks = String(response[15] || '').trim();   // P

  const totalExp = parseYears_(totalExpText);
  const relevantExp = parseYears_(relevantExpText);

  const currentSalary = parseMoney_(currentSalaryText);
  const expectedSalary = parseMoney_(expectedSalaryText);

  const noticeDays = parseNoticeDays_(noticeText);

  const concerns = [];
  const questions = [];
  const positives = [];

  let screeningResult = 'Strong Match';

  // =====================================================
  // 1. RELEVANT EXPERIENCE
  // Minimum ASM requirement = 1 year
  // =====================================================

  if (relevantExp === null) {

    concerns.push(
      'Relevant retail experience could not be confirmed from the candidate response.'
    );

    questions.push(
      'Please confirm your total relevant retail experience.'
    );

    screeningResult = 'Review';

  } else if (relevantExp < 1) {

    concerns.push(
      'Relevant retail experience is below the 1-year ASM minimum.'
    );

    screeningResult = 'Weak Match';

  } else {

    positives.push(
      relevantExp + ' year(s) of relevant experience reported'
    );
  }


  // =====================================================
  // 2. LEADERSHIP / TEAM HANDLING
  // Preferred, not compulsory
  // =====================================================

  const leadershipText = (
    cvRelevantExperience + ' ' +
    cvDesignation + ' ' +
    currentRole
  ).toLowerCase();

  const leadershipKeywords = [
    'manager',
    'assistant manager',
    'supervisor',
    'team leader',
    'team handling',
    'team management',
    'team coordination',
    'team support',
    'department manager',
    'leadership',
    'supporting junior'
  ];

  const hasLeadership = leadershipKeywords.some(
    keyword => leadershipText.includes(keyword)
  );

  if (hasLeadership) {

    positives.push('leadership/team exposure indicated');

  } else {

    concerns.push(
      'Team-handling exposure is not clearly established.'
    );

    questions.push(
      'What team size have you handled, supervised or supported?'
    );
  }


  // =====================================================
  // 3. KPI / TARGET EXPOSURE
  // Preferred, not compulsory
  // =====================================================

  const kpiText = (
    cvRelevantExperience + ' ' +
    cvExistingConcerns
  ).toLowerCase();

  const kpiKeywords = [
    'kpi',
    'target',
    'conversion',
    'sales performance',
    'sales target',
    'performance',
    'mis'
  ];

  const hasKpiExposure = kpiKeywords.some(
    keyword => kpiText.includes(keyword)
  );

  if (hasKpiExposure) {

    positives.push('KPI/target exposure indicated');

  } else {

    concerns.push(
      'KPI and target-management exposure is not clearly established.'
    );

    questions.push(
      'Which retail KPIs and sales targets have you handled?'
    );
  }


  // =====================================================
  // 4. NOTICE PERIOD
  // <=30 preferred
  // 31–60 review
  // >60 concern
  // =====================================================

  if (noticeDays === null) {

    concerns.push('Notice period / joining availability needs confirmation.');

    questions.push(
      'What is your earliest possible joining date?'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }

  } else if (noticeDays <= 30) {

    positives.push(
      noticeDays === 0
        ? 'available to join immediately'
        : noticeDays + '-day notice period'
    );

  } else if (noticeDays <= 60) {

    concerns.push(
      'Notice period is ' + noticeDays + ' days and needs review.'
    );

    questions.push(
      'Can the notice period be reduced?'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }

  } else {

    concerns.push(
      'Notice period exceeds 60 days.'
    );

    questions.push(
      'Is earlier joining possible?'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }
  }


  // =====================================================
  // 5. SALARY FIT
  //
  // Small stores: ₹20,000–₹25,000
  // Kochi / Trivandrum: up to ₹40,000
  // Salary never automatically rejects candidate.
  // =====================================================

  const locationText =
    (preferredLocation + ' ' + currentLocation).toLowerCase();

  const highBudgetLocation =
    isKochiOrTrivandrum_(locationText);

  const smallStoreLocation =
    isSmallStoreLocation_(locationText);

  if (expectedSalary === null) {

    concerns.push('Expected salary is not clearly available.');

    questions.push(
      'Please confirm your expected monthly salary.'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }

  } else if (highBudgetLocation) {

    if (expectedSalary <= 40000) {

      positives.push(
        'expected salary ₹' +
        expectedSalary.toLocaleString('en-IN') +
        ' is within Kochi/Trivandrum ASM ceiling'
      );

    } else {

      concerns.push(
        'Expected salary ₹' +
        expectedSalary.toLocaleString('en-IN') +
        ' is above the current Kochi/Trivandrum ASM range.'
      );

      questions.push(
        'Is there flexibility in your expected salary?'
      );

      if (screeningResult === 'Strong Match') {
        screeningResult = 'Review';
      }
    }

  } else if (smallStoreLocation) {

    if (expectedSalary <= 25000) {

      positives.push(
        'expected salary ₹' +
        expectedSalary.toLocaleString('en-IN') +
        ' is within the small-store ASM range'
      );

    } else {

      concerns.push(
        'Expected salary ₹' +
        expectedSalary.toLocaleString('en-IN') +
        ' exceeds the ₹25,000 small-store ASM ceiling.'
      );

      questions.push(
        'Would you consider the salary range applicable to the preferred store?'
      );

      if (screeningResult === 'Strong Match') {
        screeningResult = 'Review';
      }
    }

  } else {

    if (expectedSalary > 25000) {

      concerns.push(
        'Salary fit depends on the final store/location assignment.'
      );

      questions.push(
        'Which store locations are you willing to consider?'
      );

      if (screeningResult === 'Strong Match') {
        screeningResult = 'Review';
      }

    } else {

      positives.push(
        'expected salary is within the standard ASM range'
      );
    }
  }


  // =====================================================
  // 6. LOCATION / RELOCATION
  // =====================================================

  if (!preferredLocation) {

    concerns.push('Preferred work location is missing.');

    questions.push(
      'Which Brynex locations are you willing to work in?'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }
  }

  if (!relocate) {

    concerns.push('Relocation willingness is not confirmed.');

    questions.push(
      'Are you willing to relocate if required?'
    );
  }


  // =====================================================
  // 7. CV / FORM EXPERIENCE CONSISTENCY
  // =====================================================

  const cvYears = parseYears_(cvTotalExperience);

  if (
    cvYears !== null &&
    totalExp !== null &&
    Math.abs(cvYears - totalExp) >= 2
  ) {

    concerns.push(
      'Total experience stated in the form differs materially from the CV.'
    );

    questions.push(
      'Please confirm your correct total work experience.'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }
  }


  // =====================================================
  // SUMMARY
  // =====================================================

  let summaryParts = [];

  if (relevantExp !== null) {
    summaryParts.push(
      relevantExp + ' year(s) relevant experience'
    );
  }

  if (currentRole) {
    summaryParts.push(
      'current role: ' + currentRole
    );
  }

  if (preferredLocation) {
    summaryParts.push(
      'preferred location: ' + preferredLocation
    );
  }

  if (expectedSalary !== null) {
    summaryParts.push(
      'expected salary ₹' +
      expectedSalary.toLocaleString('en-IN')
    );
  }

  if (noticeDays !== null) {
    summaryParts.push(
      noticeDays === 0
        ? 'immediate joining'
        : noticeDays + '-day notice'
    );
  }

  if (positives.length > 0) {
    summaryParts.push(
      positives.join('; ')
    );
  }

  const summary =
    summaryParts.length > 0
      ? summaryParts.join('. ') + '.'
      : 'Candidate response received and requires recruiter review.';


  // Remove duplicate questions
  const uniqueQuestions = [...new Set(questions)];

  const callQuestions = uniqueQuestions.length
    ? uniqueQuestions
        .map((q, i) => (i + 1) + '. ' + q)
        .join('\n')
    : 'No major clarification required. Recruiter may proceed with normal screening call.';

  const concernText = concerns.length
    ? [...new Set(concerns)].join(' ')
    : 'No major concerns identified from the available CV and candidate response.';

  return {
    screeningResult: screeningResult,
    summary: summary,
    concerns: concernText,
    callQuestions: callQuestions
  };
}


/**
 * Extract first usable year value.
 */
function parseYears_(value) {

  if (
    value === null ||
    value === undefined ||
    String(value).trim() === ''
  ) {
    return null;
  }

  const text = String(value)
    .toLowerCase()
    .replace(/,/g, '.');

  const match = text.match(/\d+(\.\d+)?/);

  if (!match) return null;

  return Number(match[0]);
}


/**
 * Convert salary text to numeric monthly salary.
 */
function parseMoney_(value) {

  if (
    value === null ||
    value === undefined ||
    String(value).trim() === ''
  ) {
    return null;
  }

  let text = String(value)
    .toLowerCase()
    .trim();

  // Examples: 25k, 25 k
  const kMatch = text.match(/(\d+(?:\.\d+)?)\s*k/);

  if (kMatch) {
    return Math.round(Number(kMatch[1]) * 1000);
  }

  const digits = text.replace(/[^\d.]/g, '');

  if (!digits) return null;

  return Math.round(Number(digits));
}


/**
 * Convert notice period into approximate days.
 */
function parseNoticeDays_(value) {

  if (
    value === null ||
    value === undefined ||
    String(value).trim() === ''
  ) {
    return null;
  }

  const text = String(value).toLowerCase().trim();

  if (
    text.includes('immediate') ||
    text.includes('immediately') ||
    text === '0'
  ) {
    return 0;
  }

  const match = text.match(/\d+(\.\d+)?/);

  if (!match) return null;

  const number = Number(match[0]);

  if (text.includes('month')) {
    return Math.round(number * 30);
  }

  if (text.includes('week')) {
    return Math.round(number * 7);
  }

  return Math.round(number);
}


/**
 * Kochi / Trivandrum higher ASM salary locations.
 */
function isKochiOrTrivandrum_(text) {

  const locations = [
    'kochi',
    'ernakulam',
    'edappally',
    'edapally',
    'mg road',
    'trivandrum',
    'thiruvananthapuram'
  ];

  return locations.some(
    location => text.includes(location)
  );
}


/**
 * Current smaller-store locations.
 */
function isSmallStoreLocation_(text) {

  const locations = [
    'kottayam',
    'perumbavoor',
    'thrissur',
    'palakkad',
    'edappal',
    'kottakkal',
    'manjeri',
    'perinthalmanna',
    'calicut',
    'kozhikode',
    'vadakara',
    'kannur',
    'kalpetta'
  ];

  return locations.some(
    location => text.includes(location)
  );
}
function syncShortlistReviewQueue() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const intakeSheet = ss.getSheetByName(INTAKE_SHEET);
  const responseSheet = ss.getSheetByName(RESPONSE_SHEET);
  const queueSheet = ss.getSheetByName('Shortlist Review Queue');

  if (!intakeSheet || !responseSheet || !queueSheet) {
    throw new Error('Required recruitment sheet is missing.');
  }

  const intakeLastRow = intakeSheet.getLastRow();
  const responseLastRow = responseSheet.getLastRow();
  const queueLastRow = queueSheet.getLastRow();

  if (intakeLastRow < 2) return;

  const intakeData = intakeSheet
    .getRange(2, 1, intakeLastRow - 1, 25)
    .getValues();

  const responseData = responseLastRow > 1
    ? responseSheet
        .getRange(2, 1, responseLastRow - 1, 17)
        .getValues()
    : [];

  // Latest candidate response by Intake ID
  const latestResponse = {};

  responseData.forEach(row => {
    const intakeId = String(row[1] || '').trim();

    if (intakeId) {
      latestResponse[intakeId] = row;
    }
  });

  // Existing queue rows by Intake ID
  const queueMap = {};

  if (queueLastRow > 1) {

    const queueIds = queueSheet
      .getRange(2, 1, queueLastRow - 1, 1)
      .getValues()
      .flat();

    queueIds.forEach((id, index) => {

      if (id) {
        queueMap[String(id).trim()] = index + 2;
      }
    });
  }

  intakeData.forEach(cv => {

    const intakeId = String(cv[0] || '').trim();

    if (!intakeId) return;

    const screeningResult =
      String(cv[15] || '').trim();

    // Human review only for these
    if (
      screeningResult !== 'Strong Match' &&
      screeningResult !== 'Review'
    ) {
      return;
    }

    const response = latestResponse[intakeId];

    const candidateName = cv[4] || '';
    const position = cv[3] || '';

    const location = response
      ? response[6] || cv[6] || ''
      : cv[6] || '';

    const relevantExperience = response
      ? response[9] || cv[8] || ''
      : cv[8] || '';

    const expectedSalary = response
      ? response[11] || ''
      : '';

    const noticePeriod = response
      ? response[12] || ''
      : '';

    const screeningSummary = cv[16] || '';
    const concerns = cv[17] || '';

    const rowValues = [
      intakeId,
      candidateName,
      position,
      location,
      relevantExperience,
      expectedSalary,
      noticePeriod,
      screeningResult,
      screeningSummary,
      concerns
    ];

    if (queueMap[intakeId]) {

      // Update automated columns only.
      // Human Decision / Reviewed By / Review Date remain untouched.
      queueSheet
        .getRange(queueMap[intakeId], 1, 1, 10)
        .setValues([rowValues]);

    } else {

      queueSheet.appendRow([
        ...rowValues,
        '', // Human Decision
        '', // Reviewed By
        ''  // Review Date
      ]);
    }
  });

  console.log('Shortlist Review Queue synced.');
}
function screenSalesCandidate_(cv, response) {

  const cvLocation = String(cv[6] || '').trim();
  const cvTotalExperience = String(cv[7] || '').trim();
  const cvRelevantExperience = String(cv[8] || '').trim();
  const cvDesignation = String(cv[9] || '').trim();

  const currentLocation = String(response[6] || '').trim();
  const currentRole = String(response[7] || '').trim();
  const totalExpText = String(response[8] || '').trim();
  const relevantExpText = String(response[9] || '').trim();
  const expectedSalaryText = String(response[11] || '').trim();
  const noticeText = String(response[12] || '').trim();
  const relocate = String(response[13] || '').trim();
  const preferredLocation = String(response[14] || '').trim();

  const totalExp = parseYears_(totalExpText);
  const relevantExp = parseYears_(relevantExpText);
  const expectedSalary = parseMoney_(expectedSalaryText);
  const noticeDays = parseNoticeDays_(noticeText);

  const positives = [];
  const concerns = [];
  const questions = [];

  let screeningResult = 'Strong Match';

  // 1. Minimum relevant experience = 1 year
  if (relevantExp === null) {

    concerns.push(
      'Relevant sales/retail experience needs confirmation.'
    );

    questions.push(
      'Please confirm your relevant sales or retail experience.'
    );

    screeningResult = 'Review';

  } else if (relevantExp < 1) {

    concerns.push(
      'Relevant sales/retail experience is below the 1-year requirement.'
    );

    screeningResult = 'Weak Match';

  } else {

    positives.push(
      relevantExp + ' year(s) relevant sales/retail experience'
    );
  }

  // 2. Customer-facing / sales exposure
  const salesText = (
    cvRelevantExperience + ' ' +
    cvDesignation + ' ' +
    currentRole
  ).toLowerCase();

  const salesKeywords = [
    'sales',
    'customer',
    'retail',
    'store',
    'billing',
    'cashier',
    'customer service',
    'sales executive',
    'sales associate'
  ];

  const hasSalesExposure = salesKeywords.some(
    keyword => salesText.includes(keyword)
  );

  if (hasSalesExposure) {

    positives.push('customer-facing/sales exposure indicated');

  } else {

    concerns.push(
      'Direct customer-facing sales exposure is not clearly established.'
    );

    questions.push(
      'Please describe your customer-facing sales experience.'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }
  }

  // 3. Fashion/apparel exposure - preferred only
  const fashionKeywords = [
    'fashion',
    'apparel',
    'garment',
    'menswear',
    'womenswear',
    'textile',
    'clothing'
  ];

  const hasFashionExposure = fashionKeywords.some(
    keyword => salesText.includes(keyword)
  );

  if (hasFashionExposure) {
    positives.push('fashion/apparel exposure indicated');
  }

  // 4. Salary range ₹13,000–₹15,000
  if (expectedSalary === null) {

    concerns.push('Expected salary needs confirmation.');

    questions.push(
      'Please confirm your expected monthly salary.'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }

  } else if (expectedSalary <= 15000) {

    positives.push(
      'expected salary ₹' +
      expectedSalary.toLocaleString('en-IN') +
      ' is within Sales budget'
    );

  } else {

    concerns.push(
      'Expected salary ₹' +
      expectedSalary.toLocaleString('en-IN') +
      ' is above the ₹15,000 Sales budget.'
    );

    questions.push(
      'Is there flexibility in your expected salary?'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }
  }

  // 5. Notice period
  if (noticeDays === null) {

    concerns.push(
      'Notice period / joining availability needs confirmation.'
    );

    questions.push(
      'What is your earliest possible joining date?'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }

  } else if (noticeDays <= 30) {

    positives.push(
      noticeDays === 0
        ? 'available to join immediately'
        : noticeDays + '-day notice period'
    );

  } else {

    concerns.push(
      'Notice period is ' + noticeDays + ' days.'
    );

    questions.push(
      'Can you join within 30 days?'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }
  }

  // 6. Preferred location
  if (!preferredLocation) {

    concerns.push(
      'Preferred work location is not confirmed.'
    );

    questions.push(
      'Which Brynex locations are you willing to work in?'
    );
  }

  if (!relocate) {

    questions.push(
      'Are you open to working at another nearby location if required?'
    );
  }

  // 7. CV/form consistency
  const cvYears = parseYears_(cvTotalExperience);

  if (
    cvYears !== null &&
    totalExp !== null &&
    Math.abs(cvYears - totalExp) >= 2
  ) {

    concerns.push(
      'Total experience stated in the form differs materially from the CV.'
    );

    questions.push(
      'Please confirm your correct total work experience.'
    );

    if (screeningResult === 'Strong Match') {
      screeningResult = 'Review';
    }
  }

  // Summary
  const summaryParts = [];

  if (relevantExp !== null) {
    summaryParts.push(
      relevantExp + ' year(s) relevant experience'
    );
  }

  if (currentRole) {
    summaryParts.push(
      'current role: ' + currentRole
    );
  }

  if (preferredLocation) {
    summaryParts.push(
      'preferred location: ' + preferredLocation
    );
  }

  if (expectedSalary !== null) {
    summaryParts.push(
      'expected salary ₹' +
      expectedSalary.toLocaleString('en-IN')
    );
  }

  if (noticeDays !== null) {
    summaryParts.push(
      noticeDays === 0
        ? 'immediate joining'
        : noticeDays + '-day notice'
    );
  }

  if (positives.length) {
    summaryParts.push(positives.join('; '));
  }

  const summary =
    summaryParts.join('. ') + '.';

  const concernText = concerns.length
    ? [...new Set(concerns)].join(' ')
    : 'No major concerns identified from the available CV and candidate response.';

  const uniqueQuestions = [...new Set(questions)];

  const callQuestions = uniqueQuestions.length
    ? uniqueQuestions
        .map((q, i) => (i + 1) + '. ' + q)
        .join('\n')
    : 'No major clarification required. Recruiter may proceed with normal screening call.';

  return {
    screeningResult,
    summary,
    concerns: concernText,
    callQuestions
  };
}
function syncInterviewSchedulingQueue() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const shortlistSheet =
    ss.getSheetByName('Shortlist Review Queue');

  const intakeSheet =
    ss.getSheetByName(INTAKE_SHEET);

  const interviewSheet =
    ss.getSheetByName('Interview Scheduling Queue');

  if (!shortlistSheet || !intakeSheet || !interviewSheet) {
    throw new Error('Required recruitment sheet is missing.');
  }

  const shortlistLastRow = shortlistSheet.getLastRow();
  const intakeLastRow = intakeSheet.getLastRow();
  const interviewLastRow = interviewSheet.getLastRow();

  if (shortlistLastRow < 2) return;

  const shortlistData = shortlistSheet
    .getRange(2, 1, shortlistLastRow - 1, 13)
    .getValues();

  const intakeData = intakeLastRow > 1
    ? intakeSheet
        .getRange(2, 1, intakeLastRow - 1, 25)
        .getValues()
    : [];

  // Intake ID → CV Intake row
  const intakeMap = {};

  intakeData.forEach(row => {
    const intakeId = String(row[0] || '').trim();

    if (intakeId) {
      intakeMap[intakeId] = row;
    }
  });

  // Existing interview queue IDs
  const existingIds = new Set();

  if (interviewLastRow > 1) {

    interviewSheet
      .getRange(2, 1, interviewLastRow - 1, 1)
      .getValues()
      .flat()
      .forEach(id => {

        if (id) {
          existingIds.add(String(id).trim());
        }
      });
  }

  shortlistData.forEach(row => {

    const intakeId = String(row[0] || '').trim();
    const candidateName = row[1] || '';
    const position = row[2] || '';
    const location = row[3] || '';
    const decision = String(row[10] || '').trim();

    if (!intakeId) return;

    // Only human-approved candidates
    if (decision !== 'Proceed to Interview') return;

    // Avoid duplicate queue entries
    if (existingIds.has(intakeId)) return;

    const cv = intakeMap[intakeId];

    if (!cv) return;

    const mobileNumber = cv[5] || '';

interviewSheet.appendRow([
  intakeId,        // A Intake ID
  candidateName,   // B Candidate Name
  position,        // C Position
  mobileNumber,    // D Mobile Number
  location,        // E Location

  '',              // F Slot 1 Date
  '',              // G Slot 1 Time
  '',              // H Slot 2 Date
  '',              // I Slot 2 Time

  '',              // J Interviewer
  '',              // K Interview Mode
  '',              // L Interview Details

  '',              // M Candidate Reply
  '',              // N Confirmed Date
  '',              // O Confirmed Time

  'To Schedule',   // P Interview Status
  'Not Sent',      // Q Invite Status
  '',              // R Invite Sent On
  ''               // S Notes
]);

    existingIds.add(intakeId);
  });

  console.log('Interview Scheduling Queue synced.');
}
function handleWhatsAppWebhook_(payload) {

  if (!payload || !payload.entry) {
    return;
  }

  const ss =
    SpreadsheetApp.openById(SPREADSHEET_ID);

  const debugSheet =
    ss.getSheetByName('Webhook Debug');

  const cache =
    CacheService.getScriptCache();


  payload.entry.forEach(entry => {

    const changes = entry.changes || [];

    changes.forEach(change => {

      const value = change.value || {};
      const messages = value.messages || [];

      // Ignore sent / delivered / read status events
      if (!messages.length) {
        return;
      }


      messages.forEach(message => {

        const messageId =
          String(message.id || '').trim();

        const mobileNumber =
          String(message.from || '').trim();

        let reply = '';
        let buttonPayload = '';


        // --------------------------------
        // TEMPLATE QUICK-REPLY BUTTON
        // --------------------------------

        if (
          message.type === 'button' &&
          message.button
        ) {

          reply =
            String(
              message.button.text || ''
            )
            .trim()
            .toLowerCase();

          buttonPayload =
            String(
              message.button.payload || ''
            )
            .trim();
        }


        // --------------------------------
        // INTERACTIVE BUTTON FALLBACK
        // --------------------------------

        else if (
          message.type === 'interactive' &&
          message.interactive &&
          message.interactive.button_reply
        ) {

          reply =
            String(
              message.interactive
                .button_reply
                .title || ''
            )
            .trim()
            .toLowerCase();

          buttonPayload =
            String(
              message.interactive
                .button_reply
                .id || ''
            )
            .trim();
        }


        // --------------------------------
        // NORMAL TEXT FALLBACK
        // --------------------------------

        else if (
          message.type === 'text' &&
          message.text
        ) {

          reply =
            String(
              message.text.body || ''
            )
            .trim()
            .toLowerCase();
        }


        else {
          return;
        }


        if (!mobileNumber || !reply) {
          return;
        }


        // --------------------------------
        // DUPLICATE PROTECTION
        // --------------------------------

        if (messageId) {

          const cacheKey =
            'WA_PROCESSED_' + messageId;

          const lock =
            LockService.getScriptLock();

          try {

            lock.waitLock(5000);

            // Same WhatsApp event already handled
            if (cache.get(cacheKey)) {
              return;
            }

            // Mark it before processing
            cache.put(
              cacheKey,
              'processing',
              21600 // 6 hours
            );

          } finally {

            lock.releaseLock();
          }


          try {

            // Debug only real candidate messages
            if (debugSheet) {

              debugSheet.appendRow([
                new Date(),
                message.type || '',
                mobileNumber,
                reply,
                buttonPayload,
                JSON.stringify(message)
              ]);
            }


            processInterviewReply_(
              mobileNumber,
              reply,
              buttonPayload
            );


            cache.put(
              cacheKey,
              'processed',
              21600
            );

          } catch (error) {

            // Allow Meta retry if processing genuinely failed
            cache.remove(cacheKey);

            throw error;
          }

        }


        // Fallback for any message with no ID
        else {

          processInterviewReply_(
            mobileNumber,
            reply,
            buttonPayload
          );
        }

      });
    });
  });
}

function processInterviewReply_(mobileNumber, reply, buttonPayload) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Interview Scheduling Queue');

  if (!sheet) {
    throw new Error('Interview Scheduling Queue is missing.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet
    .getRange(2, 1, lastRow - 1, 21)
    .getValues();

  let action = '';
  let intakeId = '';

  // Preferred method: read the quick-reply button payload
  if (buttonPayload) {

    const parts = String(buttonPayload).split('|');

    action = String(parts[0] || '').trim().toUpperCase();
    intakeId = String(parts[1] || '').trim();
  }

  // Fallback if WhatsApp only sends button text
  if (!action) {

    const replyText = String(reply || '').trim().toLowerCase();

    if (replyText.includes('slot 1')) {
      action = 'SLOT1';
    } else if (replyText.includes('slot 2')) {
      action = 'SLOT2';
    } else if (replyText.includes('another')) {
      action = 'RESCHEDULE';
    }
  }

  let rowNumber = null;
  let rowData = null;

  // First try matching by Intake ID
  if (intakeId) {

    for (let i = data.length - 1; i >= 0; i--) {

      const rowIntakeId =
        String(data[i][0] || '').trim();

      if (rowIntakeId === intakeId) {

        rowNumber = i + 2;
        rowData = data[i];
        break;
      }
    }
  }

  // Fallback: match by mobile number
  if (!rowNumber) {

    const incomingMobile =
      normalizeMobile_(mobileNumber);

    for (let i = data.length - 1; i >= 0; i--) {

      const rowMobile =
        normalizeMobile_(data[i][3]);

      const inviteStatus =
        String(data[i][16] || '').trim();

      if (
        rowMobile === incomingMobile &&
        inviteStatus === 'Sent'
      ) {

        rowNumber = i + 2;
        rowData = data[i];
        break;
      }
    }
  }

  if (!rowNumber || !rowData) {
    console.log('No matching interview row found.');
    return;
  }

  // Candidate asks for another time
  if (action === 'RESCHEDULE') {

    sheet.getRange(rowNumber, 13)
      .setValue('Need Another Time');

    sheet.getRange(rowNumber, 16)
      .setValue('Reschedule');

    return;
  }

  let slotId = '';
  let candidateReply = '';

  if (action === 'SLOT1') {

    slotId = String(rowData[19] || '').trim();
    candidateReply = 'Confirm Slot 1';

  } else if (action === 'SLOT2') {

    slotId = String(rowData[20] || '').trim();
    candidateReply = 'Confirm Slot 2';

  } else {

    console.log(
      'Unknown interview reply: ' +
      reply +
      ' / ' +
      buttonPayload
    );

    return;
  }

  if (!slotId) {
    throw new Error('Selected interview slot ID is missing.');
  }

  const slot =
    getInterviewSlotById_(ss, slotId);

  if (!slot) {
    throw new Error(
      'Interview slot not found: ' +
      slotId
    );
  }

  updateConfirmedInterview_(
    sheet,
    rowNumber,
    slot,
    candidateReply
  );

  sendInterviewConfirmationText_(
    rowData[3],
    slot
  );
}

function updateConfirmedInterview_(
  sheet,
  rowNumber,
  slot,
  candidateReply
) {

  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(10000);

    // J:Q written together in one update
    sheet
      .getRange(rowNumber, 10, 1, 8)
      .setValues([[
        slot.interviewer,   // J Interviewer
        slot.mode,          // K Interview Mode
        slot.details,       // L Interview Details
        candidateReply,     // M Candidate Reply
        slot.date,          // N Confirmed Date
        slot.time,          // O Confirmed Time
        'Scheduled',        // P Interview Status
        'Confirmed'         // Q Invite Status
      ]]);

    SpreadsheetApp.flush();

  } finally {

    lock.releaseLock();
  }
}

function sendInterviewConfirmationText_(
  mobileNumber,
  slot
) {

  const props =
    PropertiesService.getScriptProperties();

  const phoneNumberId =
    props.getProperty(
      'WHATSAPP_PHONE_NUMBER_ID'
    );

  const token =
    props.getProperty(
      'WHATSAPP_ACCESS_TOKEN'
    );

  const cleanNumber =
    normalizeWhatsAppNumber_(
      mobileNumber
    );

  let message =
    'Your Brynex interview is confirmed.\n\n' +
    'Date: ' + slot.dateText + '\n' +
    'Time: ' + slot.timeText + '\n' +
    'Mode: ' + slot.mode;

  if (slot.details) {
    message +=
      '\nDetails: ' + slot.details;
  }

  const payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: cleanNumber,
    type: 'text',

    text: {
      body: message
    }
  };

  const url =
    'https://graph.facebook.com/v26.0/' +
    phoneNumberId +
    '/messages';

  const response =
    UrlFetchApp.fetch(url, {
      method: 'post',

      headers: {
        Authorization:
          'Bearer ' + token
      },

      contentType: 'application/json',

      payload:
        JSON.stringify(payload),

      muteHttpExceptions: true
    });

  const status =
    response.getResponseCode();

  if (status < 200 || status >= 300) {

    console.error(
      'Interview confirmation message failed: ' +
      response.getContentText()
    );
  }
}

function sendInterviewInviteForRow_(rowNumber) {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const queueSheet =
    ss.getSheetByName('Interview Scheduling Queue');

  if (!queueSheet) {
    throw new Error('Interview Scheduling Queue is missing.');
  }

  const row = queueSheet
    .getRange(rowNumber, 1, 1, 21)
    .getValues()[0];

  const intakeId = String(row[0] || '').trim();
  const candidateName = String(row[1] || '').trim();
  const position = String(row[2] || '').trim();
  const mobile = row[3];

  const inviteStatus =
    String(row[16] || '').trim();

  const slot1Id =
    String(row[19] || '').trim();

  const slot2Id =
    String(row[20] || '').trim();

  if (!intakeId) throw new Error('Intake ID missing.');
  if (!candidateName) throw new Error('Candidate name missing.');
  if (!mobile) throw new Error('Mobile number missing.');
  if (!position) throw new Error('Position missing.');

  if (!slot1Id || !slot2Id) {
    throw new Error('Interview slots have not been assigned.');
  }

  if (inviteStatus.toLowerCase() === 'sent') {
    throw new Error('Interview invitation already sent.');
  }

  const slot1 =
    getInterviewSlotById_(ss, slot1Id);

  const slot2 =
    getInterviewSlotById_(ss, slot2Id);

  if (!slot1 || !slot2) {
    throw new Error('Assigned interview slot could not be found.');
  }

  if (Number(slot1.available) <= 0) {
    throw new Error('Slot 1 is full.');
  }

  if (Number(slot2.available) <= 0) {
    throw new Error('Slot 2 is full.');
  }

  const result =
    sendInterviewWhatsAppTemplate_(
      mobile,
      candidateName,
      position,

      slot1.dateText,
      slot1.timeText,
      slot1.mode,

      slot2.dateText,
      slot2.timeText,
      slot2.mode,

      intakeId
    );

  // Q Invite Status
  queueSheet
    .getRange(rowNumber, 17)
    .setValue('Sent');

  // R Invite Sent On
  queueSheet
    .getRange(rowNumber, 18)
    .setValue(new Date());

  return result;
}


function sendInterviewWhatsAppTemplate_(
  mobileNumber,
  candidateName,
  position,
  slot1Date,
  slot1Time,
  slot2Date,
  slot2Time,
  mode,
  interviewDetails,
  intakeId
) {

  const props =
    PropertiesService.getScriptProperties();

  const phoneNumberId =
    props.getProperty('WHATSAPP_PHONE_NUMBER_ID');

  const token =
    props.getProperty('WHATSAPP_ACCESS_TOKEN');

  const templateName =
    props.getProperty(
      'WHATSAPP_INTERVIEW_TEMPLATE_NAME'
    ) || 'brynex_interview_slot_confirmation';

  const language =
    props.getProperty(
      'WHATSAPP_INTERVIEW_TEMPLATE_LANGUAGE'
    ) || 'en';

  const cleanNumber =
    normalizeWhatsAppNumber_(mobileNumber);

  const payload = {
    messaging_product: 'whatsapp',
    recipient_type: 'individual',
    to: cleanNumber,

    type: 'template',

    template: {
      name: templateName,

      language: {
        code: language
      },

      components: [

        {
          type: 'body',

          parameters: [
            { type: 'text', text: candidateName },
            { type: 'text', text: position },
            { type: 'text', text: slot1Date },
            { type: 'text', text: slot1Time },
            { type: 'text', text: slot2Date },
            { type: 'text', text: slot2Time },
            { type: 'text', text: mode },
            { type: 'text', text: interviewDetails }
          ]
        },

        {
          type: 'button',
          sub_type: 'quick_reply',
          index: '0',

          parameters: [
            {
              type: 'payload',
              payload: 'SLOT1|' + intakeId
            }
          ]
        },

        {
          type: 'button',
          sub_type: 'quick_reply',
          index: '1',

          parameters: [
            {
              type: 'payload',
              payload: 'SLOT2|' + intakeId
            }
          ]
        },

        {
          type: 'button',
          sub_type: 'quick_reply',
          index: '2',

          parameters: [
            {
              type: 'payload',
              payload: 'RESCHEDULE|' + intakeId
            }
          ]
        }
      ]
    }
  };

  const url =
    'https://graph.facebook.com/v26.0/' +
    phoneNumberId +
    '/messages';

  const response =
    UrlFetchApp.fetch(url, {
      method: 'post',

      headers: {
        Authorization: 'Bearer ' + token
      },

      contentType: 'application/json',

      payload: JSON.stringify(payload),

      muteHttpExceptions: true
    });

  const status =
    response.getResponseCode();

  const body =
    response.getContentText();

  if (status < 200 || status >= 300) {
    throw new Error(
      'Interview WhatsApp API failed (' +
      status +
      '): ' +
      body
    );
  }

  return JSON.parse(body);
}
function assignInterviewSlotsFromMaster() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const masterSheet =
    ss.getSheetByName('Interview Slot Master');

  const queueSheet =
    ss.getSheetByName('Interview Scheduling Queue');

  if (!masterSheet || !queueSheet) {
    throw new Error('Interview sheets are missing.');
  }

  const masterLastRow = masterSheet.getLastRow();
  const queueLastRow = queueSheet.getLastRow();

  if (masterLastRow < 2 || queueLastRow < 2) return;

  const masterData =
    masterSheet
      .getRange(2, 1, masterLastRow - 1, 10)
      .getValues();

  // Only active slots with availability
  const availableSlots = masterData
    .filter(row => {

      const slotId = String(row[0] || '').trim();
      const date = row[1];
      const time = row[2];
      const available = Number(row[8] || 0);
      const active = String(row[9] || '').trim();

      return (
        slotId &&
        date &&
        time &&
        available > 0 &&
        active === 'Yes'
      );
    })
    .sort((a, b) => {

      const aDateTime =
        combineInterviewDateTime_(a[1], a[2]);

      const bDateTime =
        combineInterviewDateTime_(b[1], b[2]);

      return aDateTime - bDateTime;
    });

  if (availableSlots.length < 2) {
    console.log('Less than 2 interview slots available.');
    return;
  }

  const queueData =
    queueSheet
      .getRange(2, 1, queueLastRow - 1, 21)
      .getValues();

  queueData.forEach((row, index) => {

    const rowNumber = index + 2;

    const interviewStatus =
      String(row[15] || '').trim(); // P

    const inviteStatus =
      String(row[16] || '').trim(); // Q

    const existingSlot1Id =
      String(row[19] || '').trim(); // T

    const existingSlot2Id =
      String(row[20] || '').trim(); // U

    if (interviewStatus !== 'To Schedule') return;
    if (inviteStatus !== 'Not Sent') return;

    // Already assigned
    if (existingSlot1Id || existingSlot2Id) return;

    const slot1 = availableSlots[0];
    const slot2 = availableSlots[1];

    // F-I = Date/Time options
    queueSheet
      .getRange(rowNumber, 6, 1, 4)
      .setValues([[
        slot1[1], // Slot 1 Date
        slot1[2], // Slot 1 Time
        slot2[1], // Slot 2 Date
        slot2[2]  // Slot 2 Time
      ]]);

    // T-U = technical Slot IDs
    queueSheet
      .getRange(rowNumber, 20, 1, 2)
      .setValues([[
        slot1[0],
        slot2[0]
      ]]);
  });

  console.log('Interview slots assigned.');
}


function combineInterviewDateTime_(dateValue, timeValue) {

  const date = new Date(dateValue);
  const time = new Date(timeValue);

  date.setHours(
    time.getHours(),
    time.getMinutes(),
    0,
    0
  );

  return date;
}
function sendPendingInterviewInvites() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheet =
    ss.getSheetByName('Interview Scheduling Queue');

  if (!sheet) {
    throw new Error('Interview Scheduling Queue is missing.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const data =
    sheet
      .getRange(2, 1, lastRow - 1, 21)
      .getValues();

  let sentCount = 0;

  data.forEach((row, index) => {

    const rowNumber = index + 2;

    const interviewStatus =
      String(row[15] || '').trim(); // P

    const inviteStatus =
      String(row[16] || '').trim(); // Q

    const slot1Id =
      String(row[19] || '').trim(); // T

    const slot2Id =
      String(row[20] || '').trim(); // U

    if (interviewStatus !== 'To Schedule') return;

    if (inviteStatus !== 'Not Sent') return;

    // Don't send until two slots are assigned
    if (!slot1Id || !slot2Id) return;

    try {

      sendInterviewInviteForRow_(rowNumber);

      sentCount++;

    } catch (error) {

      // S = Notes
      sheet
        .getRange(rowNumber, 19)
        .setValue(
          'Interview invite error: ' +
          error.message
        );

      console.error(
        'Row ' +
        rowNumber +
        ': ' +
        error.message
      );
    }
  });

  console.log(
    'Interview invitations sent: ' +
    sentCount
  );
}
function testInterviewInvite() {
  sendInterviewInviteForRow_(2);
}
function getInterviewSlotById_(ss, slotId) {

  const sheet =
    ss.getSheetByName('Interview Slot Master');

  if (!sheet) {
    throw new Error('Interview Slot Master is missing.');
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return null;

  const data =
    sheet
      .getRange(2, 1, lastRow - 1, 10)
      .getValues();

  const slot = data.find(row =>
    String(row[0] || '').trim() ===
    String(slotId || '').trim()
  );

  if (!slot) return null;

  const tz = Session.getScriptTimeZone();

  let dateText = '';

  if (slot[1] instanceof Date) {
    dateText = Utilities.formatDate(
      slot[1],
      tz,
      'dd MMM yyyy'
    );
  } else {
    dateText = String(slot[1] || '').trim();
  }

  let timeText = '';

  if (slot[2] instanceof Date) {
    timeText = Utilities.formatDate(
      slot[2],
      tz,
      'hh:mm a'
    );
  } else {
    timeText = String(slot[2] || '').trim();
  }

  return {
    id: slot[0],
    date: slot[1],
    time: slot[2],
    dateText: dateText,
    timeText: timeText,
    interviewer: slot[3] || '',
    mode: slot[4] || '',
    details: slot[5] || '',
    capacity: Number(slot[6] || 0),
    confirmed: Number(slot[7] || 0),
    available: Number(slot[8] || 0),
    active: String(slot[9] || '').trim()
  };
}
function normalizeMobile_(number) {
  const digits = String(number || '').replace(/\D/g, '');

  // For Indian numbers coming from WhatsApp as 91XXXXXXXXXX
  if (digits.length > 10) {
    return digits.slice(-10);
  }

  return digits;
}
function appendIntakeRow_(sheet, rowValues) {
  const ids = sheet
    .getRange(2, 1, sheet.getMaxRows() - 1, 1)
    .getDisplayValues()
    .flat();

  let targetRow = 2;

  for (let i = 0; i < ids.length; i++) {
    if (!ids[i]) {
      targetRow = i + 2;
      break;
    }
  }

  sheet
    .getRange(targetRow, 1, 1, rowValues.length)
    .setValues([rowValues]);
}