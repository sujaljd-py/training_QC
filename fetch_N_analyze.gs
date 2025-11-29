/**
 * QC CONFIGURATION
 */
const QC_CONFIG = {
  // 1. PASTE YOUR SHEET ID
  SHEET_ID: '15Z_VEjGlSuxF6fD4J1m6x6E7u-p8HVBM0Z_rxAGxn4I', 
  
  // 2. PASTE YOUR GEMINI API KEY
  API_KEY: 'AIzaSyCJ3RNKXhaRfd0SF6i5SPy00nUPQf3Uz90', 
  
  // 3. MODEL: Setting to Gemini 2.0 Flash Lite (Preview)
  // If this fails, revert to 'gemini-1.5-flash'
  MODEL_STRING: 'gemini-2.0-flash-lite-preview-02-05', 
  
  LOGS_TAB: 'Fetch Logs',
  RESULTS_TAB: 'QC Results',
  
  // Regex to clean titles (Removes "in Hindi", "in English", brackets)
  TITLE_CLEAN_REGEX: /(\s+in\s+(Hindi|English).*)|(\s*\(.*\))/gi
};

function runQCAnalysis() {
  const ss = SpreadsheetApp.openById(QC_CONFIG.SHEET_ID);
  const logSheet = ss.getSheetByName(QC_CONFIG.LOGS_TAB);
  let resultSheet = ss.getSheetByName(QC_CONFIG.RESULTS_TAB);

  // 1. Setup Result Sheet if missing
  if (!resultSheet) {
    resultSheet = ss.insertSheet(QC_CONFIG.RESULTS_TAB);
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

  // 2. Read Logs
  const logData = logSheet.getDataRange().getValues();
  const headers = logData[0];
  
  // --- FIX: MATCH EXACT HEADERS FROM FETCH LOGS ---
  const col = {
    docId: headers.indexOf("Doc ID"),
    title: headers.indexOf("Title"),       // Changed from "Training Title" to "Title"
    status: headers.indexOf("Status"),
    date: headers.indexOf("Date"),         // Changed from "Training Date" to "Date"
    duration: headers.indexOf("Duration (Mins)"),
    link: headers.indexOf("Link")          // Changed from "File Link" to "Link"
  };

  // Check if columns were found
  if (col.docId === -1 || col.title === -1) {
    Logger.log("❌ ERROR: Could not find required columns in 'Fetch Logs'.");
    Logger.log("Found Headers: " + headers.join(", "));
    return;
  }

  const processedIds = resultSheet.getDataRange().getValues().map(r => r[0]);

  // 3. Loop through Logs
  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const docId = row[col.docId];
    const status = row[col.status];
    let fullTitle = row[col.title]; // Use let so we can modify if needed

    // SAFETY CHECK: If title is missing/undefined, skip row
    if (!fullTitle) continue;
    
    // Check validity
    if (!docId || docId === "Doc ID") continue;
    if (status !== "SUCCESS") continue;
    if (processedIds.includes(docId)) continue; 

    // Convert to string safely (just in case it's interpreted as something else)
    fullTitle = String(fullTitle);

    Logger.log(`\n🤖 STARTING QC: "${fullTitle}"`);

    // --- STEP A: FIND CRITERIA TAB ---
    let cleanTitle = fullTitle.replace(QC_CONFIG.TITLE_CLEAN_REGEX, "").trim();
    cleanTitle = cleanTitle.replace(/-\s*$/, "").trim(); 

    Logger.log(`   Looking for Tab: "${cleanTitle}"`);
    
    const criteriaSheet = ss.getSheetByName(cleanTitle);

    if (!criteriaSheet) {
      Logger.log(`   ❌ Criteria Tab not found. Skipping.`);
      resultSheet.appendRow([
        docId, row[col.date], fullTitle, row[col.duration], row[col.link],
        "SKIPPED", "Tab Not Found: " + cleanTitle, 
        "", "", "", "", "", "", 0, 0, 0
      ]);
      continue;
    }

    // --- STEP B: PREPARE TOPICS ---
    const criteriaData = criteriaSheet.getDataRange().getValues();
    const topicList = criteriaData.slice(1)
      .filter(r => r[0] !== "")
      .map(r => `- TOPIC: "${r[0]}"\n  REQUIREMENT: ${r[1]}`)
      .join("\n");

    Logger.log(`   ✅ Found Criteria (${criteriaData.length - 1} topics).`);

    // --- STEP C: DOWNLOAD TRANSCRIPT ---
    let transcriptText = "";
    try {
      transcriptText = exportDocAsText(docId);
    } catch (e) {
      Logger.log(`   ❌ Transcript Download Error: ${e.message}`);
      continue;
    }

    // --- STEP D: PROMPT ---
    const prompt = `
      You are a Quality Control Auditor.
      Analyze the training transcript against the Required Topics.
      
      CONTEXT:
      - Training: "${cleanTitle}"
      - Duration: ${row[col.duration]} minutes
      
      REQUIRED TOPICS:
      ${topicList}

      TASK:
      1. Intro: Did the trainer state their name?
      2. Greeting: Did they greet attendees?
      3. Tone: Is it professional?
      4. COVERAGE ANALYSIS:
         - Calculate Coverage % = (Count of Covered Topics / Total Topics) * 100.
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

    // --- STEP E: CALL API ---
    try {
      const apiResponse = callGeminiAPI(prompt);
      const ai = apiResponse.json;
      const usage = apiResponse.usage;

      resultSheet.appendRow([
        docId,
        row[col.date],
        fullTitle,
        row[col.duration],
        row[col.link],
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
      
      Logger.log(`   ✅ QC Saved. Score: ${ai.qc_score}`);

    } catch (e) {
      Logger.log(`   ❌ AI Error: ${e.message}`);
      resultSheet.appendRow([
        docId, row[col.date], fullTitle, row[col.duration], row[col.link],
        "ERROR", e.message, "", "", "", "", "", "", 0, 0, 0
      ]);
    }
  }
}

function callGeminiAPI(promptText) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${QC_CONFIG.MODEL_STRING}:generateContent?key=${QC_CONFIG.API_KEY}`;

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
