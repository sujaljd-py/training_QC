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
  DAYS_TO_LOOK_BACK: 6,
  INCLUDE_REGEX: /(Hindi|English)/i,
  EXCLUDE_REGEX: /with/i,
  MIN_LENGTH_CHARS: 500,
  WORDS_PER_MINUTE: 130,
  
  // Gemini QC Settings
  GEMINI_API_KEY: '--', // 🔥 PASTE YOUR GEMINI API KEY HERE
  GEMINI_MODEL: 'gemini-2.0-flash-lite-preview-02-05',
  TITLE_CLEAN_REGEX: /(\s+in\s+(Hindi|English).*)|(\s*\(.*\))/gi,

   
  // Email Notification Settings
  NOTIFICATION_EMAIL: 'sujal.jadhv@ezeetechnosys.com', // 🔥 CHANGE THIS TO YOUR EMAIL
  SEND_EMAIL_NOTIFICATIONS: true, // Set to false to disable emails
};


/**
 * 📧 SEND QC COMPLETION EMAIL
 * Call this at the end of runCompleteQCWorkflow() or runWeeklyQCWorkflow()
 * 
 * @param {Object} stats - Statistics object with counts
 * @param {string} status - 'SUCCESS' or 'FAILED'
 * @param {string} errorMessage - Optional error message if failed
 */
function sendQCCompletionEmail(stats, status = 'SUCCESS', errorMessage = '') {
  // Check if email notifications are enabled
  if (!CONFIG.SEND_EMAIL_NOTIFICATIONS) {
    Logger.log('📧 Email notifications disabled in CONFIG');
    return;
  }
  
  if (!CONFIG.NOTIFICATION_EMAIL || CONFIG.NOTIFICATION_EMAIL === 'your.email@example.com') {
    Logger.log('⚠️  Email notification skipped: No valid email configured');
    return;
  }
  
  // Ensure stats object exists with default values
  if (!stats) {
    stats = {};
  }
  
  // Set defaults for all stats
  stats.totalScanned = stats.totalScanned || 0;
  stats.passedValidation = stats.passedValidation || 0;
  stats.transcriptsRead = stats.transcriptsRead || 0;
  stats.qcAnalyzed = stats.qcAnalyzed || 0;
  stats.totalInputTokens = stats.totalInputTokens || 0;
  stats.totalOutputTokens = stats.totalOutputTokens || 0;
  
  try {
    const sheetUrl = `https://docs.google.com/spreadsheets/d/${extractSheetId(CONFIG.SPREADSHEET_ID)}`;
    const currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'EEEE, MMMM dd, yyyy');
    const currentTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'hh:mm a');
    
    // Determine subject based on status
    let subject = '';
    let statusEmoji = '';
    
    if (status === 'SUCCESS') {
      subject = `✅ Weekly QC Report - ${currentDate}`;
      statusEmoji = '✅';
    } else {
      subject = `❌ QC Workflow Failed - ${currentDate}`;
      statusEmoji = '❌';
    }
    
    // Build HTML email body
    let htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      background-color: #f4f4f4;
      margin: 0;
      padding: 20px;
    }
    .container {
      max-width: 600px;
      margin: 0 auto;
      background-color: #ffffff;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
      overflow: hidden;
    }
    .header {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      padding: 30px;
      text-align: center;
    }
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .header p {
      margin: 5px 0 0 0;
      opacity: 0.9;
      font-size: 14px;
    }
    .content {
      padding: 30px;
    }
    .status-badge {
      display: inline-block;
      padding: 8px 16px;
      border-radius: 20px;
      font-weight: 600;
      font-size: 14px;
      margin-bottom: 20px;
    }
    .status-success {
      background-color: #d4edda;
      color: #155724;
    }
    .status-failed {
      background-color: #f8d7da;
      color: #721c24;
    }
    .stats-table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
      background-color: #f9f9f9;
      border-radius: 8px;
      overflow: hidden;
    }
    .stats-table tr {
      border-bottom: 1px solid #e0e0e0;
    }
    .stats-table tr:last-child {
      border-bottom: none;
    }
    .stats-table td {
      padding: 15px;
      font-size: 14px;
    }
    .stats-table td:first-child {
      font-weight: 600;
      color: #555;
      width: 60%;
    }
    .stats-table td:last-child {
      text-align: right;
      font-weight: 700;
      color: #667eea;
      font-size: 16px;
    }
    .token-row {
      background-color: #fff3cd;
    }
    .cta-button {
      display: inline-block;
      padding: 12px 30px;
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
      color: white;
      text-decoration: none;
      border-radius: 25px;
      font-weight: 600;
      margin: 20px 0;
      text-align: center;
    }
    .cta-button:hover {
      opacity: 0.9;
    }
    .footer {
      background-color: #f9f9f9;
      padding: 20px;
      text-align: center;
      font-size: 12px;
      color: #777;
      border-top: 1px solid #e0e0e0;
    }
    .error-box {
      background-color: #f8d7da;
      border-left: 4px solid #dc3545;
      padding: 15px;
      margin: 15px 0;
      border-radius: 4px;
      color: #721c24;
    }
    .highlight {
      font-weight: 700;
      color: #667eea;
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- Header -->
    <div class="header">
      <h1>${statusEmoji} Training QC Workflow Report</h1>
      <p>${currentDate} • ${currentTime}</p>
    </div>
    
    <!-- Content -->
    <div class="content">
      <!-- Status Badge -->
      <div class="status-badge ${status === 'SUCCESS' ? 'status-success' : 'status-failed'}">
        ${status === 'SUCCESS' ? '✅ Workflow Completed Successfully' : '❌ Workflow Failed'}
      </div>
      
      ${status === 'FAILED' ? `
      <!-- Error Message -->
      <div class="error-box">
        <strong>Error Details:</strong><br>
        ${errorMessage || 'Unknown error occurred'}
      </div>
      ` : ''}
      
      <!-- Summary Statistics -->
      <h2 style="color: #333; margin-top: 20px;">📊 Weekly Summary</h2>
      
      <table class="stats-table">
        <tr>
          <td>📅 Total Training Sessions Scanned</td>
          <td>${stats.totalScanned || 0}</td>
        </tr>
        <tr>
          <td>✅ Passed Validation Conditions</td>
          <td>${stats.passedValidation || 0}</td>
        </tr>
        <tr>
          <td>📄 Transcripts Successfully Read</td>
          <td>${stats.transcriptsRead || 0}</td>
        </tr>
        <tr>
          <td>🤖 QC Analyses Completed</td>
          <td>${stats.qcAnalyzed || 0}</td>
        </tr>
        <tr class="token-row">
          <td>🔤 Total Input Tokens Used</td>
          <td>${stats.totalInputTokens || 0}</td>
        </tr>
        <tr class="token-row">
          <td>🔤 Total Output Tokens Used</td>
          <td>${stats.totalOutputTokens || 0}</td>
        </tr>
      </table>
      
      ${status === 'SUCCESS' && stats.qcAnalyzed > 0 ? `
      <!-- Additional Insights -->
      <h3 style="color: #333; margin-top: 25px;">📈 Key Insights</h3>
      <ul style="color: #555; line-height: 1.8;">
        <li><span class="highlight">${stats.passedValidation}</span> out of <span class="highlight">${stats.totalScanned}</span> sessions had valid transcripts</li>
        <li>Average tokens per analysis: <span class="highlight">${Math.round((stats.totalInputTokens + stats.totalOutputTokens) / stats.qcAnalyzed)}</span></li>
        <li>Success rate: <span class="highlight">${Math.round((stats.qcAnalyzed / stats.totalScanned) * 100)}%</span></li>
      </ul>
      ` : ''}
      
      <!-- CTA Button -->
      <div style="text-align: center; margin-top: 30px;">
        <a href="${sheetUrl}" class="cta-button">
          📊 View Complete QC Results
        </a>
      </div>
      
      <p style="color: #777; font-size: 13px; text-align: center; margin-top: 15px;">
        Click the button above to access the full spreadsheet with detailed QC results
      </p>
    </div>
    
    <!-- Footer -->
    <div class="footer">
      <p>🤖 Automated Training QC System</p>
      <p>Yanolja Cloud Solution Training Hub</p>
      <p style="margin-top: 10px;">
        <a href="${sheetUrl}" style="color: #667eea; text-decoration: none;">View Spreadsheet</a>
      </p>
    </div>
  </div>
</body>
</html>
    `;
    
    // Plain text version (fallback)
    let plainBody = `
TRAINING QC WORKFLOW REPORT
${status === 'SUCCESS' ? '✅ COMPLETED SUCCESSFULLY' : '❌ FAILED'}
Date: ${currentDate}
Time: ${currentTime}

${status === 'FAILED' ? `ERROR: ${errorMessage}\n` : ''}

WEEKLY SUMMARY:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📅 Total Training Sessions Scanned: ${stats.totalScanned || 0}
✅ Passed Validation Conditions: ${stats.passedValidation || 0}
📄 Transcripts Successfully Read: ${stats.transcriptsRead || 0}
🤖 QC Analyses Completed: ${stats.qcAnalyzed || 0}
🔤 Total Input Tokens Used: ${stats.totalInputTokens || 0}
🔤 Total Output Tokens Used: ${stats.totalOutputTokens || 0}

📊 VIEW COMPLETE RESULTS:
${sheetUrl}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
🤖 Automated Training QC System
Yanolja Cloud Solution Training Hub
    `;
    
    // Send email
    MailApp.sendEmail({
      to: CONFIG.NOTIFICATION_EMAIL,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: 'Training QC System'
    });
    
    Logger.log(`✅ Email notification sent to ${CONFIG.NOTIFICATION_EMAIL}`);
    
  } catch (e) {
    Logger.log(`❌ Failed to send email notification: ${e.message}`);
  }
}

/**
 * 📊 EXAMPLE: How to use this in your runCompleteQCWorkflow()
 * Add this code at the END of your runCompleteQCWorkflow() function
 */
/*

// At the end of runCompleteQCWorkflow(), before the final Logger.log:

  // Prepare stats for email
  const emailStats = {
    totalScanned: events.length,
    passedValidation: processedCount,
    transcriptsRead: processedCount,
    qcAnalyzed: analyzedCount,
    totalInputTokens: 0,  // You can track this if needed
    totalOutputTokens: 0  // You can track this if needed
  };
  
  // Send success email
  sendQCCompletionEmail(emailStats, 'SUCCESS');

  Logger.log(`\n✅ ========== WORKFLOW COMPLETE ==========`);
  Logger.log(`   📥 New records fetched & logged: ${processedCount}`);
  Logger.log(`   🤖 New records analyzed: ${analyzedCount}`);

*/



/**
 * ⭐⭐⭐ MAIN FUNCTION - RUN THIS ONE ⭐⭐⭐
 * Sequential Processing:
 * For each calendar event:
 *   1. Validate transcript
 *   2. Write to Fetch Logs
 *   3. Run QC Analysis immediately
 *   4. Write to QC Results
 *   5. Move to next one
 */
function runCompleteQCWorkflow() {
  Logger.log("🚀 ========== STARTING SEQUENTIAL QC WORKFLOW ==========\n");
  
  const cleanSheetId = extractSheetId(CONFIG.SPREADSHEET_ID);
  if (!cleanSheetId) {
    Logger.log("❌ CONFIG ERROR: Valid Spreadsheet ID required.");
    return;
  }
  
  const ss = SpreadsheetApp.openById(cleanSheetId);
  
  // Setup sheets
  const logSheet = setupLogSheet(ss);
  const resultSheet = setupResultSheet(ss);
  
  // Get existing IDs to avoid duplicates
  const existingLogIds = getExistingLogIds(logSheet);
  const processedQCIds = getProcessedDocIds(resultSheet);
  
  Logger.log(`📋 Already logged: ${existingLogIds.length} records`);
  Logger.log(`📋 Already analyzed: ${processedQCIds.length} records\n`);
  
  // ==================== FETCH CALENDAR EVENTS ====================
  const now = new Date();
  const startDate = new Date();
  startDate.setDate(now.getDate() - CONFIG.DAYS_TO_LOOK_BACK);
  const futureDate = new Date();
  futureDate.setFullYear(now.getFullYear() + 2);

  Logger.log(`📅 Fetching calendar events from ${startDate.toDateString()}...\n`);

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
    Logger.log(`❌ Could not access Calendar: ${e.message}`);
    return;
  }

  if (!events || events.length === 0) {
    Logger.log("⚠️ No events found in Calendar.");
    return;
  }

  Logger.log(`Found ${events.length} total calendar events. Processing...\n`);

  let processedCount = 0;
  let analyzedCount = 0;

  // ==================== PROCESS EACH EVENT SEQUENTIALLY ====================
  for (let eventIndex = 0; eventIndex < events.length; eventIndex++) {
    const event = events[eventIndex];
    const title = event.summary || "No Title";

    // FILTER 1: Exclude regex
    if (CONFIG.EXCLUDE_REGEX.test(title)) {
      continue;
    }

    // FILTER 2: Include regex
    if (!CONFIG.INCLUDE_REGEX.test(title)) {
      continue;
    }

    Logger.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    Logger.log(`[${eventIndex + 1}/${events.length}] 📝 Processing: "${title}"`);
    Logger.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);

    // Get event metadata
    let eventDate = new Date(event.start.dateTime || event.start.date);
    let finalDuration = 0;
    
    // Calculate duration from calendar
    if (event.start.dateTime && event.end.dateTime) {
      const start = new Date(event.start.dateTime);
      const end = new Date(event.end.dateTime);
      finalDuration = Math.round((end - start) / 1000 / 60);
    }

    // Check for attachments
    if (!event.attachments || event.attachments.length === 0) {
      Logger.log(`   ⚠️  No attachments found. Skipping.`);
      continue;
    }

    Logger.log(`   📎 Found ${event.attachments.length} attachment(s). Checking...`);

    // Look for valid transcript
    let validTranscriptFound = false;

    for (const file of event.attachments) {
      if (file.mimeType !== 'application/vnd.google-apps.document') {
        continue;
      }

      Logger.log(`   📄 Checking document: "${file.title}"`);

      const fileId = extractFileId(file);
      if (!fileId) {
        Logger.log(`      ❌ Could not extract file ID. Skipping.`);
        continue;
      }

      // STEP 1: CHECK IF ALREADY LOGGED
      if (existingLogIds.includes(fileId)) {
        Logger.log(`      ⏭️  Already logged in Fetch Logs. Skipping.`);
        continue;
      }

      // STEP 2: VALIDATE TRANSCRIPT
      let transcriptText = "";
      try {
        transcriptText = exportDocAsText(fileId);
      } catch (e) {
        Logger.log(`      ❌ Failed to download: ${e.message}`);
        continue;
      }

      const totalLength = transcriptText.length;
      const wordCount = transcriptText.split(/\s+/).length;

      Logger.log(`      📏 Length: ${totalLength} chars, ${wordCount} words`);

      // Validate minimum length
      if (totalLength < CONFIG.MIN_LENGTH_CHARS) {
        Logger.log(`      ⚠️  Too short (min: ${CONFIG.MIN_LENGTH_CHARS}). Skipping.`);
        continue;
      }

      // Estimate duration if needed
      if (finalDuration === 0) {
        const estMins = Math.round(wordCount / CONFIG.WORDS_PER_MINUTE);
        if (estMins > 5) {
          finalDuration = estMins;
        }
      }

      const formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      const fileLink = file.fileUrl;

      Logger.log(`      ✅ VALID TRANSCRIPT FOUND!`);
      validTranscriptFound = true;

      // ========== STEP 3: WRITE TO FETCH LOGS ==========
      Logger.log(`\n   📝 STEP 1: Writing to Fetch Logs...`);
      
      logSheet.appendRow([
        fileId,
        title,
        "SUCCESS",
        formattedDate,
        finalDuration,
        fileLink
      ]);
      
      existingLogIds.push(fileId); // Add to array to prevent re-processing
      processedCount++;
      
      Logger.log(`      ✅ Written to Fetch Logs (Row ${logSheet.getLastRow()})`);

      // ========== STEP 4: RUN QC ANALYSIS IMMEDIATELY ==========
      Logger.log(`\n   🤖 STEP 2: Running QC Analysis...`);

      // Check if already analyzed
      if (processedQCIds.includes(fileId)) {
        Logger.log(`      ⏭️  Already analyzed. Skipping QC.`);
      } else {
        const record = {
          docId: fileId,
          title: title,
          status: "SUCCESS",
          date: formattedDate,
          duration: finalDuration,
          link: fileLink,
          transcriptText: transcriptText // Pass the already-downloaded transcript
        };

        analyzeRecord(record, ss, resultSheet);
        processedQCIds.push(fileId);
        analyzedCount++;
      }

      Logger.log(`   ✅ COMPLETE for this event!\n`);

      break; // Found valid transcript, move to next event
    }

    if (!validTranscriptFound) {
      Logger.log(`   ❌ No valid transcript found for this event.\n`);
    }
  }
  
  Logger.log(`\n✅ ========== WORKFLOW COMPLETE ==========`);
  Logger.log(`   📥 New records fetched & logged: ${processedCount}`);
  Logger.log(`   🤖 New records analyzed: ${analyzedCount}`);

  sendQCCompletionEmail({
    totalScanned: events.length,
    passedValidation: processedCount,
    transcriptsRead: processedCount,
    qcAnalyzed: analyzedCount,
    totalInputTokens: 0,
    totalOutputTokens: 0
  }, 'SUCCESS');
}

// ============================================================================
// SETUP FUNCTIONS
// ============================================================================

function setupLogSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.FETCH_LOGS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.FETCH_LOGS_TAB);
    sheet.appendRow(["Doc ID", "Title", "Status", "Date", "Duration (Mins)", "Link"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
  return sheet;
}

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

function getExistingLogIds(logSheet) {
  const data = logSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => row[0]).filter(id => id !== "");
}

function getProcessedDocIds(resultSheet) {
  const data = resultSheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => row[0]).filter(id => id !== "");
}

// ============================================================================
// QC ANALYSIS FUNCTION
// ============================================================================

function analyzeRecord(record, ss, resultSheet) {
  const docId = record.docId;
  const fullTitle = String(record.title);
  
  // Step 1: Find criteria tab
  let cleanTitle = fullTitle.replace(CONFIG.TITLE_CLEAN_REGEX, "").trim();
  cleanTitle = cleanTitle.replace(/-\s*$/, "").trim();
  
  Logger.log(`      📋 Looking for criteria tab: "${cleanTitle}"`);
  
  const criteriaSheet = ss.getSheetByName(cleanTitle);
  
  if (!criteriaSheet) {
    Logger.log(`      ❌ Criteria tab "${cleanTitle}" not found`);
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
  Logger.log(`      ✅ Found ${totalTopics} required topics`);
  
  // Step 3: Use transcript (already downloaded during validation)
  let transcriptText = record.transcriptText;
  
  if (!transcriptText) {
    // Fallback: download again if not passed
    Logger.log(`      📄 Downloading transcript...`);
    try {
      transcriptText = exportDocAsText(docId);
    } catch (e) {
      Logger.log(`      ❌ Download failed: ${e.message}`);
      resultSheet.appendRow([
        docId, record.date, fullTitle, record.duration, record.link,
        "ERROR", `Download Failed: ${e.message}`, 
        "", "", "", "", "", "", 0, 0, 0
      ]);
      return;
    }
  }
  
  Logger.log(`      📄 Transcript ready (${transcriptText.length} chars)`);
  
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
1. Intro: Did the trainer(attendee: Yanolja Cloud Solution Training Hub) state their name?
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
    
    Logger.log(`      ✅ QC Complete - Score: ${ai.qc_score}/10, Coverage: ${ai.coverage_percentage}%`);
    Logger.log(`      📊 Tokens: Input ${usage.promptTokenCount}, Output ${usage.candidatesTokenCount}`);
    
  } catch (e) {
    Logger.log(`      ❌ AI Analysis failed: ${e.message}`);
    resultSheet.appendRow([
      docId, record.date, fullTitle, record.duration, record.link,
      "ERROR", `AI Error: ${e.message}`, 
      "", "", "", "", "", "", 0, 0, 0
    ]);
  }
}

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
