/***************************************************************
 *  FULL AUTOMATED TRAINING QC SYSTEM
 *  Calendar → Transcript → Gemini 2.5 Flash Lite → Google Sheet
 ***************************************************************/

const CONFIG = {
  GEMINI_API_KEY: "AIzaSyB5iRKoVAblr3Sj0Diu0Mqdzdv5rpcg2KA",   // <<-- REQUIRED
  SHEET_ID: "15Z_VEjGlSuxF6fD4J1m6x6E7u-p8HVBM0Z_rxAGxn4I",
  SHEET_TAB_NAME: "QC Results",
  TOPICS_TAB_NAME: "Topics",
  MODEL: "gemini-2.5-flash-lite",
  TEMPERATURE: 0.2,
  LOOKBACK_DAYS: 7,
  MIN_DURATION_MINS: 20,
  MIN_CHAR_COUNT: 500,
  MAX_EVENT_RESULTS: 250
};

/***************************************************************
 *                ENTRY POINT (Daily Trigger)
 ***************************************************************/
function runDailyQC() {
  Logger.log("=== QC SCRIPT STARTED ===");

  const meetings = fetchCalendarMeetings();
  Logger.log("Meetings found: " + meetings.length);

  meetings.forEach(event => {
    try {
      processMeeting(event);
    } catch (err) {
      logStatus({
        title: event.summary || "Untitled",
        status: "ERROR",
        reason: "Script failed: " + err,
        fileId: ""
      });
    }
  });

  Logger.log("=== QC SCRIPT FINISHED ===");
}

/***************************************************************
 *                STEP A — FETCH MEETINGS FROM CALENDAR
 ***************************************************************/
function fetchCalendarMeetings() {
  const now = new Date();
  const past = new Date();
  past.setDate(now.getDate() - CONFIG.LOOKBACK_DAYS);

  const events = Calendar.Events.list("primary", {
    timeMin: past.toISOString(),
    timeMax: now.toISOString(),
    maxResults: CONFIG.MAX_EVENT_RESULTS,
    singleEvents: true,
    orderBy: "startTime"
  });

  return events.items || [];
}

/***************************************************************
 *                PROCESS EACH MEETING
 ***************************************************************/
function processMeeting(event) {
  Logger.log("Processing: " + event.summary);

  if (!event.attachments || event.attachments.length === 0) {
    return logStatus({
      title: event.summary,
      status: "SKIPPED",
      reason: "No attachments"
    });
  }

  const { transcriptFile, notesFile } = identifyFiles(event.attachments);

  if (!transcriptFile) {
    return logStatus({
      title: event.summary,
      status: "SKIPPED",
      reason: "Transcript not found"
    });
  }

  // Duplicate check
  if (isAlreadyProcessed(transcriptFile.fileId)) {
    return logStatus({
      title: event.summary,
      status: "SKIPPED",
      reason: "Duplicate transcript",
      fileId: transcriptFile.fileId
    });
  }

  // Read transcript
  const transcriptText = readDriveFile(transcriptFile.fileId);
  const notesText = notesFile ? readDriveFile(notesFile.fileId) : "";

  // Language check
  if (!detectLanguage(event.summary)) {
    return logStatus({
      title: event.summary,
      status: "SKIPPED",
      reason: "Meeting title missing Hindi/English"
    });
  }

  // Duration check
  const duration = extractDuration(transcriptText);
  if (duration < CONFIG.MIN_DURATION_MINS) {
    return logStatus({
      title: event.summary,
      status: "SKIPPED",
      reason: "Short duration (<20 mins)",
      fileId: transcriptFile.fileId
    });
  }

  // Character count
  if (transcriptText.length < CONFIG.MIN_CHAR_COUNT) {
    return logStatus({
      title: event.summary,
      status: "SKIPPED",
      reason: "Not enough transcript content",
      fileId: transcriptFile.fileId
    });
  }

  // Load Topics from sheet
  const topicContext = lookupTopics(event.summary);

  // RUN GEMINI ANALYSIS
  const aiResult = runGeminiAnalysis(
    transcriptText,
    notesText,
    topicContext,
    detectLanguage(event.summary)
  );

  // Save to Sheet
  saveQCResult(event, aiResult, transcriptFile.fileId, transcriptText.length, duration);
}

/***************************************************************
 *     IDENTIFY TRANSCRIPT + NOTES FROM ATTACHMENTS
 ***************************************************************/
function identifyFiles(attachments) {
  let transcript = null;
  let notes = null;

  attachments.forEach(att => {
    const title = att.title.toLowerCase();

    if (title.includes("notes")) notes = att;
    else if (title.includes("transcript")) transcript = att;
  });

  // fallback: ANY doc if transcript missing
  if (!transcript) {
    transcript = attachments.find(a =>
      a.mimeType === "application/vnd.google-apps.document"
    );
  }

  return { transcriptFile: transcript, notesFile: notes };
}

/***************************************************************
 *     READ GOOGLE DRIVE FILE CONTENT
 ***************************************************************/
function readDriveFile(fileId) {
  try {
    const file = DriveApp.getFileById(fileId);

    if (file.getMimeType() === "application/vnd.google-apps.document") {
      return DocumentApp.openById(fileId).getBody().getText();
    }
    return file.getBlob().getDataAsString();
  } catch (err) {
    Logger.log("Error reading file " + fileId);
    return "";
  }
}

/***************************************************************
 *     LANGUAGE CHECK BASED ON TITLE
 ***************************************************************/
function detectLanguage(title) {
  const t = title.toLowerCase();
  return t.includes("hindi") || t.includes("english");
}

/***************************************************************
 *     EXTRACT TRANSCRIPT DURATION USING TIMESTAMPS
 ***************************************************************/
function extractDuration(text) {
  const matches = [...text.matchAll(/(\d{1,2}):(\d{2})/g)];
  if (matches.length < 2) return 0;

  const first = matches[0];
  const last = matches[matches.length - 1];

  const start = parseInt(first[1]) * 60 + parseInt(first[2]);
  const end = parseInt(last[1]) * 60 + parseInt(last[2]);

  return end - start;
}

/***************************************************************
 *     CHECK IF TRANSCRIPT ALREADY PROCESSED
 ***************************************************************/
function isAlreadyProcessed(fileId) {
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_TAB_NAME);
  const data = sheet.getDataRange().getValues();

  return data.some(r => r[15] === fileId); // File ID Column
}

/***************************************************************
 *     LOAD EXPECTED TOPICS FROM TOPICS SHEET
 ***************************************************************/
function lookupTopics(title) {
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.TOPICS_TAB_NAME);
  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    const keyword = rows[i][0];
    if (keyword && title.toLowerCase().includes(keyword.toLowerCase())) {
      return rows[i][1];
    }
  }
  return "";
}

/***************************************************************
 *     CALL GEMINI 2.5 FLASH LITE
 ***************************************************************/
function runGeminiAnalysis(transcript, notes, expectedTopics, lang) {
  const prompt = `
You are an AI QC auditor.

LANGUAGE: ${lang}
EXPECTED TOPICS: ${expectedTopics}

TRANSCRIPT:
${transcript}

NOTES:
${notes}

TASK:
Return strictly valid JSON with:

{
  "trainer_intro": "Yes/No",
  "tone_professional": "Yes/No",
  "qa_handled": "Yes/No",
  "topics_found": [],
  "topics_missing": [],
  "topic_coverage": 0,
  "ai_score": 0,
  "summary": ""
}
`;

  try {
    const response = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/" +
      CONFIG.MODEL +
      ":generateContent?key=" +
      CONFIG.GEMINI_API_KEY,
      {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: CONFIG.TEMPERATURE }
        })
      }
    );

    const result = JSON.parse(response.getContentText());
    const text = result.candidates[0].content.parts[0].text;

    return JSON.parse(text); // must be valid JSON

  } catch (err) {
    Logger.log("AI JSON ERROR: " + err);
    return {
      trainer_intro: "ERROR",
      tone_professional: "ERROR",
      qa_handled: "ERROR",
      topics_found: [],
      topics_missing: [],
      topic_coverage: 0,
      ai_score: 0,
      summary: "AI failed to generate valid JSON"
    };
  }
}

/***************************************************************
 *     SAVE RESULT TO GOOGLE SHEET
 ***************************************************************/
function saveQCResult(event, ai, fileId, charCount, duration) {
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_TAB_NAME);

  sheet.appendRow([
    new Date(),
    event.summary,
    "SUCCESS",
    ai.summary,
    detectLanguage(event.summary),
    duration,
    charCount,
    ai.trainer_intro,
    ai.tone_professional,
    ai.qa_handled,
    "", // topics expected optional
    ai.topics_found.join(", "),
    ai.topics_missing.join(", "),
    ai.topic_coverage,
    ai.ai_score,
    fileId
  ]);
}

/***************************************************************
 *     LOG SKIPPED / ERROR STATUS
 ***************************************************************/
function logStatus(obj) {
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(CONFIG.SHEET_TAB_NAME);

  sheet.appendRow([
    new Date(),
    obj.title,
    obj.status,
    obj.reason,
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    obj.fileId || ""
  ]);
}
