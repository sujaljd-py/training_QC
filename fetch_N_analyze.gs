/**
 * COMBINED CONFIGURATION
 */
const CONFIG = {
  // Sheet Configuration
  SPREADSHEET_ID: '15Z_VEjGlSuxF6fD4J1m6x6E7u-p8HVBM0Z_rxAGxn4I',
  FETCH_LOGS_TAB: 'Fetch Logs',
  QC_RESULTS_TAB: 'QC Results',
  
  // Calendar Fetch Settings
  CALENDAR_ID: 'primary',
  DAYS_TO_LOOK_BACK: 14,
  INCLUDE_REGEX: /(Hindi|English)/i,
  EXCLUDE_REGEX: /with/i,
  MIN_LENGTH_CHARS: 500,
  WORDS_PER_MINUTE: 130,
  
  // Gemini QC Settings
  GEMINI_API_KEY: '-', // 🔥 PASTE YOUR GEMINI API KEY HERE
  GEMINI_MODEL: 'gemini-2.0-flash-lite-preview-02-05',
  TITLE_CLEAN_REGEX: /(\s+in\s+(Hindi|English).*)|(\s*\(.*\))/gi
};

/**
 * ⭐⭐⭐ MAIN FUNCTION - RUN THIS ONE ⭐⭐⭐
 * Complete Workflow:
 * 1. Fetch training records from Google Calendar
 * 2. Write to Fetch Logs
 * 3. Run QC Analysis on SUCCESS records
 */
function runCompleteQCWorkflow() {
  Logger.log("🚀 ========== STARTING COMPLETE QC WORKFLOW ==========\n");
  
  const cleanSheetId = extractSheetId(CONFIG.SPREADSHEET_ID);
  if (!cleanSheetId) {
    Logger.log("❌ CONFIG ERROR: Valid Spreadsheet ID required.");
    return;
  }
  
  const ss = SpreadsheetApp.openById(cleanSheetId);
  
  // ==================== STEP 1: FETCH FROM CALENDAR ====================
  Logger.log("📅 STEP 1: Fetching training records from Google Calendar...\n");
  
  const newRecords = fetchMeetingTranscripts(ss);
  
  if (newRecords.length === 0) {
    Logger.log("✅ No new training records found. Everything is up to date!\n");
    return;
  }
  
  Logger.log(`✅ Fetched ${newRecords.length} new records and wrote to Fetch Logs\n`);
  
  // ==================== STEP 2: RUN QC ANALYSIS ====================
  Logger.log("🤖 STEP 2: Running QC Analysis on SUCCESS records...\n");
  
  const resultSheet = setupResultSheet(ss);
  const processedQCIds = getProcessedDocIds(resultSheet);
  
  let analyzedCount = 0;
  
  for (let i = 0; i < newRecords.length; i++) {
    const record = newRecords[i];
    
    // Only analyze SUCCESS records that haven't been analyzed yet
    if (record.status === "SUCCESS" && !processedQCIds.includes(record.docId)) {
      Logger.log(`[${i + 1}/${newRecords.length}] 🤖 Analyzing: "${record.title}"`);
      analyzeRecord(record, ss, resultSheet);
      analyzedCount++;
    } else if (record.status !== "SUCCESS") {
      Logger.log(`[${i + 1}/${newRecords.length}] ⏭️  Skipping: "${record.title}" - Status: ${record.status}`);
    } else {
      Logger.log(`[${i + 1}/${newRecords.length}] ⏭️  Skipping: "${record.title}" - Already analyzed`);
    }
  }
  
  Logger.log(`\n✅ ========== WORKFLOW COMPLETE ==========`);
  Logger.log(`   📥 Fetched & Logged: ${newRecords.length} records`);
  Logger.log(`   🤖 QC Analyzed: ${analyzedCount} records`);
}

// ============================================================================
// PART 1: CALENDAR FETCH FUNCTIONS
// ============================================================================

/**
 * Fetch meeting transcripts from Google Calendar
 * Returns array of new records that were logged
 */
function fetchMeetingTranscripts(ss) {
  const now = new Date();
  const startDate = new Date();
  startDate.setDate(now.getDate() - CONFIG.DAYS_TO_LOOK_BACK);
  const futureDate = new Date();
  futureDate.setFullYear(now.getFullYear() + 2);

  Logger.log(`   Date Range: ${startDate.toDateString()} to Future`);

  // 1. Get or create Fetch Logs sheet
  let sheet = ss.getSheetByName(CONFIG.FETCH_LOGS_TAB);
  if (!sheet) {
    Logger.log(`   Creating new tab: ${CONFIG.FETCH_LOGS_TAB}`);
    sheet = ss.insertSheet(CONFIG.FETCH_LOGS_TAB);
    sheet.appendRow(["Doc ID", "Title", "Status", "Date", "Duration (Mins)", "Link"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }

  // 2. Get existing Doc IDs to avoid duplicates
  const existingData = sheet.getDataRange().getValues();
  const existingDocIds = existingData.slice(1).map(row => row[0]).filter(id => id !== "");

  // 3. Fetch Calendar Events
  let events = [];
  try {
    const response = Calendar.Events.list(CONFIG.CALENDAR_ID, {
      timeMin: startDate.toISOString(),
      timeMax: futureDate.toISOString(),
      singleEvents: true,
      orderBy: 'startTime'
    });
    events = response.items;
  } catch (e) {
    Logger.log(`   ❌ Could not access Calendar: ${e.message}`);
    return [];
  }

  if (!events || events.length === 0) {
    Logger.log("   ⚠️ No events found in Calendar for this date range.");
    return [];
  }

  Logger.log(`   Found ${events.length} total calendar events. Filtering...`);

  let rowsToLog = [];
  let newRecords = [];

  // 4. Process each event
  events.forEach(event => {
    const title = event.summary || "No Title";

    // Filter 1: Exclude
    if (CONFIG.EXCLUDE_REGEX.test(title)) return;

    // Filter 2: Include
    if (!CONFIG.INCLUDE_REGEX.test(title)) return;

    // Get event date
    let eventDate = new Date(event.start.dateTime || event.start.date);
    let finalDuration = 0;
    let status = "NOT FOUND / MISSING";
    let docId = "N/A";
    let fileLink = "N/A";

    // Calculate duration from calendar
    if (event.start.dateTime && event.end.dateTime) {
      const start = new Date(event.start.dateTime);
      const end = new Date(event.end.dateTime);
      finalDuration = Math.round((end - start) / 1000 / 60);
    }

    // Check for transcript attachment
    if (event.attachments && event.attachments.length > 0) {
      for (const file of event.attachments) {
        if (file.mimeType === 'application/vnd.google-apps.document') {
          let fileId = extractFileId(file);
          if (!fileId) continue;

          // Skip if already logged
          if (existingDocIds.includes(fileId)) continue;

          try {
            const fullText = exportDocAsText(fileId);
            const totalLength = fullText.length;
            const wordCount = fullText.split(/\s+/).length;

            if (totalLength < CONFIG.MIN_LENGTH_CHARS) continue;

            // FOUND VALID TRANSCRIPT
            docId = fileId;
            fileLink = file.fileUrl;
            status = "SUCCESS";

            // Estimate duration from word count if needed
            if (finalDuration === 0) {
              const estMins = Math.round(wordCount / CONFIG.WORDS_PER_MINUTE);
              if (estMins > 5) {
                finalDuration = estMins;
              }
            }

            break; // Stop looking for more docs in this event
          } catch (e) {
            // Continue to next attachment
          }
        }
      }
    }

    // Only log if we found a valid doc that's not already logged
    if (docId !== "N/A" && !existingDocIds.includes(docId)) {
      const formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      rowsToLog.push([
        docId,
        title,
        status,
        formattedDate,
        finalDuration,
        fileLink
      ]);

      newRecords.push({
        docId: docId,
        title: title,
        status: status,
        date: formattedDate,
        duration: finalDuration,
        link: fileLink
      });

      Logger.log(`   ✅ Found: "${title}" (${status})`);
    }
  });

  // 5. Write to sheet
  if (rowsToLog.length > 0) {
    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;
    sheet.getRange(nextRow, 1, rowsToLog.length, rowsToLog[0].length).setValues(rowsToLog);
    Logger.log(`   💾 Wrote ${rowsToLog.length} new records to Fetch Logs`);
  }

  return newRecords;
}

// ============================================================================
// PART 2: QC ANALYSIS FUNCTIONS
// ============================================================================

/**
 * Setup QC Results sheet
 */
function setupResultSheet(ss) {
  let resultSheet = ss.getSheetByName(CONFIG.QC_RESULTS_TAB);
  
  if (!resultSheet) {
    resultSheet = ss.insertSheet(CONFIG.QC_RESULTS_TAB);
    const headers = [
      "Doc ID", "Date", "Title", "Duration", "Link", 
      "Intro (Y/N)", "Intro Proof", "Greeting (Y/N)", "Greeting Proof", 
      "Tone Professional?", "Topic Coverage %", "Missing Topics", 
      "Summary", "QC Score", "Input Tokens", "Output Tokens"
    ];
    resultSheet.appendRow(headers);
    resultSheet.setFrozenRows(1);
    resultSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  
  return resultSheet;
}

/**
 * Get processed Doc IDs from QC Results
 */
function getProcessedDocIds(resultSheet) {
  const data = resultSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => row[0]).filter(id => id !== "");
}

/**
 * Analyze a single record with Gemini
 */
function analyzeRecord(record, ss, resultSheet) {
  const docId = record.docId;
  const fullTitle = String(record.title);
  
  // Step 1: Find criteria tab
  let cleanTitle = fullTitle.replace(CONFIG.TITLE_CLEAN_REGEX, "").trim();
  cleanTitle = cleanTitle.replace(/-\s*$/, "").trim();
  
  Logger.log(`      📋 Looking for criteria: "${cleanTitle}"`);
  
  const criteriaSheet = ss.getSheetByName(cleanTitle);
  
  if (!criteriaSheet) {
    Logger.log(`      ❌ Criteria tab not found`);
    resultSheet.appendRow([
      docId, record.date, fullTitle, record.duration, record.link,
      "SKIPPED", `Tab Not Found: ${cleanTitle}`, 
      "", "", "", "", "", "", 0, 0, 0
    ]);
    return;
  }
  
  // Step 2: Get topics from criteria sheet
  const criteriaData = criteriaSheet.getDataRange().getValues();
  const topicList = criteriaData.slice(1)
    .filter(r => r[0] !== "")
    .map(r => `- TOPIC: "${r[0]}"\n  REQUIREMENT: ${r[1]}`)
    .join("\n");
  
  const totalTopics = criteriaData.length - 1;
  Logger.log(`      ✅ Found ${totalTopics} topics`);
  
  // Step 3: Download transcript
  let transcriptText = "";
  try {
    transcriptText = exportDocAsText(docId);
    Logger.log(`      📄 Transcript downloaded (${transcriptText.length} chars)`);
  } catch (e) {
    Logger.log(`      ❌ Transcript download failed: ${e.message}`);
    resultSheet.appendRow([
      docId, record.date, fullTitle, record.duration, record.link,
      "ERROR", `Download Failed: ${e.message}`, 
      "", "", "", "", "", "", 0, 0, 0
    ]);
    return;
  }
  
  // Step 4: Build prompt
  const prompt = `
You are a Quality Control Auditor.
Analyze the training transcript against the Required Topics.

CONTEXT:
- Training: "${cleanTitle}"
- Duration: ${record.duration} minutes
- Total Required Topics: ${totalTopics}

REQUIRED TOPICS:
${topicList}

TASK:
1. Intro: Did the trainer state their name?
2. Greeting: Did they greet attendees?
3. Tone: Is it professional?
4. COVERAGE ANALYSIS:
   - Calculate Coverage % = (Count of Covered Topics / ${totalTopics}) * 100.
   - List ONLY topics that are MISSING.
5. Summary: Brief summary.
6. Score: 1-10.

TRANSCRIPT:
${transcriptText.substring(0, 70000)}

OUTPUT JSON:
{
  "intro_yes_no": "Yes/No",
  "intro_proof": "String",
  "greeting_yes_no": "Yes/No",
  "greeting_proof": "String",
  "tone_professional": "Yes/No",
  "coverage_percentage": number,
  "missing_topics_list": "String",
  "summary_text": "String",
  "qc_score": number
}
`;
  
  // Step 5: Call Gemini API
  try {
    Logger.log(`      🧠 Calling Gemini API...`);
    const apiResponse = callGeminiAPI(prompt);
    const ai = apiResponse.json;
    const usage = apiResponse.usage;
    
    resultSheet.appendRow([
      docId,
      record.date,
      fullTitle,
      record.duration,
      record.link,
      ai.intro_yes_no,
      ai.intro_proof,
      ai.greeting_yes_no,
      ai.greeting_proof,
      ai.tone_professional,
      ai.coverage_percentage + "%",
      ai.missing_topics_list,
      ai.summary_text,
      ai.qc_score,
      usage.promptTokenCount,
      usage.candidatesTokenCount
    ]);
    
    Logger.log(`      ✅ Score: ${ai.qc_score}/10, Coverage: ${ai.coverage_percentage}%`);
    
  } catch (e) {
    Logger.log(`      ❌ AI Analysis failed: ${e.message}`);
    resultSheet.appendRow([
      docId, record.date, fullTitle, record.duration, record.link,
      "ERROR", `AI Error: ${e.message}`, 
      "", "", "", "", "", "", 0, 0, 0
    ]);
  }
}

/**
 * Call Gemini API
 */
function callGeminiAPI(promptText) {
  if (!CONFIG.GEMINI_API_KEY) {
    throw new Error("Gemini API Key not configured!");
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${CONFIG.GEMINI_MODEL}:generateContent?key=${CONFIG.GEMINI_API_KEY}`;

  const payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: {
      temperature: 0.2,
      responseMimeType: "application/json"
    }
  };

  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error(`API Error: ${json.error ? json.error.message : "Unknown"}`);
  }

  let rawText = json.candidates[0].content.parts[0].text;
  rawText = rawText.replace(/```json/g, "").replace(/```/g, "").trim();

  return {
    json: JSON.parse(rawText),
    usage: json.usageMetadata || { promptTokenCount: 0, candidatesTokenCount: 0 }
  };
}

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

function extractSheetId(urlOrId) {
  if (!urlOrId || urlOrId.includes("YOUR_SPREADSHEET")) return null;
  const match = urlOrId.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : urlOrId;
}

function extractFileId(fileObj) {
  if (fileObj.fileId) return fileObj.fileId;
  const match = fileObj.fileUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (match) return match[1];
  return null;
}

function exportDocAsText(fileId) {
  const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${fileId}&exportFormat=txt`;
  const options = {
    method: "GET",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  if (response.getResponseCode() !== 200) throw new Error("Export Failed");
  return response.getContentText();
}
