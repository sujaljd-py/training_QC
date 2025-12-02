/**
 * OPTIMIZED TRAINING QC CONFIGURATION
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
  
  // Gemini QC Settings
  GEMINI_API_KEY: '--', // 🔥 PASTE YOUR GEMINI API KEY HERE
  GEMINI_MODEL: 'gemini-2.0-flash-lite-preview-02-05',
  TITLE_CLEAN_REGEX: /(\s+in\s+(Hindi|English).*)|(\s*\(.*\))/gi,
   
  // Email Notification Settings
  NOTIFICATION_EMAIL: 'sujal.jadhv@ezeetechnosys.com',
  SEND_EMAIL_NOTIFICATIONS: true,
};

/**
 * ⭐ MAIN WORKFLOW FUNCTION
 */
function runCompleteQCWorkflow() {
  Logger.log("🚀 ========== STARTING TRAINING QC WORKFLOW ==========\n");
  
  const cleanSheetId = extractSheetId(CONFIG.SPREADSHEET_ID);
  if (!cleanSheetId) {
    sendQCCompletionEmail(null, 'FAILED', 'Invalid Spreadsheet ID in CONFIG');
    return;
  }
  
  let ss;
  try {
    ss = SpreadsheetApp.openById(cleanSheetId);
  } catch (e) {
    sendQCCompletionEmail(null, 'FAILED', `Cannot open spreadsheet: ${e.message}`);
    return;
  }
  
  const logSheet = setupLogSheet(ss);
  const resultSheet = setupResultSheet(ss);
  
  const existingLogIds = getExistingIds(logSheet);
  const processedQCIds = getExistingIds(resultSheet);
  
  Logger.log(`📋 Already logged: ${existingLogIds.length} records`);
  Logger.log(`📋 Already analyzed: ${processedQCIds.length} records\n`);
  
  // Fetch calendar events
  const now = new Date();
  const startDate = new Date();
  startDate.setDate(now.getDate() - CONFIG.DAYS_TO_LOOK_BACK);
  const futureDate = new Date();
  futureDate.setFullYear(now.getFullYear() + 2);

  Logger.log(`📅 Fetching events from ${startDate.toDateString()}...\n`);

  let events = [];
  try {
    const response = Calendar.Events.list(CONFIG.CALENDAR_ID, {
      timeMin: startDate.toISOString(),
      timeMax: futureDate.toISOString(),
      singleEvents: true,
      orderBy: 'startTime'
    });
    events = response.items || [];
  } catch (e) {
    const errorMsg = `Calendar access failed: ${e.message}`;
    Logger.log(`❌ ${errorMsg}`);
    sendQCCompletionEmail(null, 'FAILED', errorMsg);
    return;
  }

  if (events.length === 0) {
    Logger.log("⚠️ No events found in Calendar.");
    sendQCCompletionEmail({ totalScanned: 0, passedValidation: 0, transcriptsRead: 0, qcAnalyzed: 0 }, 'SUCCESS');
    return;
  }

  Logger.log(`Found ${events.length} calendar events. Processing...\n`);

  let processedCount = 0;
  let analyzedCount = 0;
  let totalInputTokens = 0;
  let totalOutputTokens = 0;

  // Process each event
  for (let i = 0; i < events.length; i++) {
    const event = events[i];
    const title = event.summary || "No Title";

    if (CONFIG.EXCLUDE_REGEX.test(title) || !CONFIG.INCLUDE_REGEX.test(title)) {
      continue;
    }

    Logger.log(`\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);
    Logger.log(`[${i + 1}/${events.length}] 📝 Processing: "${title}"`);
    Logger.log(`━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`);

    const eventDate = new Date(event.start.dateTime || event.start.date);
    const formattedDate = Utilities.formatDate(eventDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

    if (!event.attachments || event.attachments.length === 0) {
      Logger.log(`   ⚠️  No attachments. Skipping.`);
      continue;
    }

    Logger.log(`   📎 Found ${event.attachments.length} attachment(s)`);

    let validTranscriptFound = false;

    for (const file of event.attachments) {
      if (file.mimeType !== 'application/vnd.google-apps.document') continue;

      Logger.log(`   📄 Checking: "${file.title}"`);

      const fileId = extractFileId(file);
      if (!fileId) {
        Logger.log(`      ❌ Could not extract file ID`);
        continue;
      }

      if (existingLogIds.includes(fileId)) {
        Logger.log(`      ⏭️  Already logged. Skipping.`);
        continue;
      }

      // Download and validate transcript
      let transcriptText = "";
      try {
        transcriptText = exportDocAsText(fileId);
      } catch (e) {
        Logger.log(`      ❌ Download failed: ${e.message}`);
        continue;
      }

      if (transcriptText.length < CONFIG.MIN_LENGTH_CHARS) {
        Logger.log(`      ⚠️  Too short (${transcriptText.length} chars). Skipping.`);
        continue;
      }

      // Extract duration from transcript timestamps
      const duration = extractDurationFromTranscript(transcriptText);
      Logger.log(`      ⏱️  Duration: ${duration} minutes (from timestamps)`);

      const fileLink = file.fileUrl;
      validTranscriptFound = true;

      // Check if training is too short (likely no-show or cancelled)
      if (duration < 20) {
        Logger.log(`      ⚠️  SHORT TRAINING (${duration} mins) - Likely No Show`);
        Logger.log(`\n   📝 Writing to Fetch Logs with SKIPPED status...`);
        
        logSheet.appendRow([fileId, title, "SKIPPED - Short Training/No Show", formattedDate, duration, fileLink]);
        existingLogIds.push(fileId);
        processedCount++;
        
        // Also write to QC Results with No Show status
        resultSheet.appendRow([
          fileId,                    // Doc ID
          formattedDate,             // Date
          title,                     // Title
          duration,                  // Duration
          fileLink,                  // Link
          "No Show",                 // Status
          "No Show",                 // Intro (Y/N)
          "No Show",                 // Intro Proof
          "No Show",                 // Greeting (Y/N)
          "No Show",                 // Greeting Proof
          "No Show",                 // Tone Professional?
          "No Show",                 // Topic Coverage %
          "No Show",                 // Missing Topics
          `Training too short (${duration} mins) - Likely no show or cancelled`, // Summary
          0,                         // QC Score
          0,                         // Input Tokens
          0                          // Output Tokens
        ]);
        
        Logger.log(`      ✅ Logged as No Show (Saved API tokens!)`);
        Logger.log(`   ✅ COMPLETE!\n`);
        break;
      }

      Logger.log(`      ✅ VALID TRANSCRIPT!`);

      // Write to Fetch Logs
      Logger.log(`\n   📝 Writing to Fetch Logs...`);
      logSheet.appendRow([fileId, title, "SUCCESS", formattedDate, duration, fileLink]);
      existingLogIds.push(fileId);
      processedCount++;
      Logger.log(`      ✅ Logged (Row ${logSheet.getLastRow()})`);

      // Run QC Analysis
      Logger.log(`\n   🤖 Running QC Analysis...`);
      
      if (processedQCIds.includes(fileId)) {
        Logger.log(`      ⏭️  Already analyzed. Skipping QC.`);
      } else {
        const record = {
          docId: fileId,
          title: title,
          date: formattedDate,
          duration: duration,
          link: fileLink,
          transcriptText: transcriptText
        };

        const tokens = analyzeRecord(record, ss, resultSheet);
        if (tokens) {
          totalInputTokens += tokens.input;
          totalOutputTokens += tokens.output;
        }
        processedQCIds.push(fileId);
        analyzedCount++;
      }

      Logger.log(`   ✅ COMPLETE!\n`);
      break;
    }

    if (!validTranscriptFound) {
      Logger.log(`   ❌ No valid transcript found.\n`);
    }
  }
  
  Logger.log(`\n✅ ========== WORKFLOW COMPLETE ==========`);
  Logger.log(`   📥 New records logged: ${processedCount}`);
  Logger.log(`   🤖 New records analyzed: ${analyzedCount}`);
  Logger.log(`   🔤 Total tokens used: ${totalInputTokens + totalOutputTokens}`);

  sendQCCompletionEmail({
    totalScanned: events.length,
    passedValidation: processedCount,
    transcriptsRead: processedCount,
    qcAnalyzed: analyzedCount,
    totalInputTokens: totalInputTokens,
    totalOutputTokens: totalOutputTokens
  }, 'SUCCESS');
}

/**
 * 🕒 EXTRACT DURATION FROM TRANSCRIPT TIMESTAMPS
 */
function extractDurationFromTranscript(transcript) {
  // Match timestamps in format: HH:MM:SS, HH:MM, MM:SS, or [HH:MM:SS], [HH:MM], [MM:SS]
  const timestampRegex = /\[?(\d{1,2}):(\d{2})(?::(\d{2}))?\]?/g;
  const matches = [...transcript.matchAll(timestampRegex)];
  
  if (matches.length < 2) {
    Logger.log(`      ⚠️  Could not find timestamps. Defaulting to 90 mins.`);
    return 90;
  }
  
  // Convert timestamp to minutes
  function toMinutes(hours, mins, secs = 0) {
    return parseInt(hours) * 60 + parseInt(mins) + parseInt(secs) / 60;
  }
  
  // Get first and last timestamp
  const firstMatch = matches[0];
  const lastMatch = matches[matches.length - 1];
  
  const startMins = toMinutes(firstMatch[1], firstMatch[2], firstMatch[3] || 0);
  const endMins = toMinutes(lastMatch[1], lastMatch[2], lastMatch[3] || 0);
  
  const duration = Math.round(endMins - startMins);
  
  // Validate duration (should be between 10 and 300 minutes)
  if (duration < 10 || duration > 300) {
    Logger.log(`      ⚠️  Calculated duration (${duration} mins) seems invalid. Using 90 mins.`);
    return 90;
  }
  
  return duration;
}

/**
 * 🤖 ANALYZE RECORD WITH GEMINI
 */
function analyzeRecord(record, ss, resultSheet) {
  const docId = record.docId;
  const fullTitle = String(record.title);
  
  let cleanTitle = fullTitle.replace(CONFIG.TITLE_CLEAN_REGEX, "").trim();
  cleanTitle = cleanTitle.replace(/-\s*$/, "").trim();
  
  Logger.log(`      📋 Looking for criteria: "${cleanTitle}"`);
  
  const criteriaSheet = ss.getSheetByName(cleanTitle);
  
  if (!criteriaSheet) {
    Logger.log(`      ❌ Criteria tab not found`);
    resultSheet.appendRow([
      docId, record.date, fullTitle, record.duration, record.link,
      "SKIPPED", "Tab Not Found: " + cleanTitle, 
      "", "", "", "", "", "", 0, 0, 0
    ]);
    return null;
  }
  
  const criteriaData = criteriaSheet.getDataRange().getValues();
  const topicList = criteriaData.slice(1)
    .filter(r => r[0] !== "")
    .map(r => `- "${r[0]}"`)
    .join("\n");
  
  const totalTopics = criteriaData.length - 1;
  Logger.log(`      ✅ Found ${totalTopics} required topics`);
  
  const transcriptText = record.transcriptText;
  Logger.log(`      📄 Transcript ready (${transcriptText.length} chars)`);
  
  const prompt = `You are a Quality Control Auditor for hotel technology training sessions.

TRAINING CONTEXT:
These are daily training sessions for Yanolja Cloud Solution Training Hub's hotel tech systems including:
- PMS (Property Management System)
- Channel Manager
- Booking Engine
- POS (Point of Sale)

The goal is to ensure every new joinee understands the key features and their practical usage.

SESSION DETAILS:
- Training: "${cleanTitle}"
- Duration: ${record.duration} minutes
- Total Required Topics: ${totalTopics}

REQUIRED TOPICS TO COVER:
${topicList}

EVALUATION CRITERIA:
1. **Intro Check**: Did someone from Yanolja Cloud Solution Training Hub specifically introduce themselves by stating their name at the beginning?
   - Look for phrases like "I am [Name] from Yanolja Cloud Solution Training Hub" or "My name is [Name], I'm from Yanolja"
   - The introduction must include BOTH the trainer's name AND mention of Yanolja Cloud Solution Training Hub
   - Just saying "Hello" or starting the training without proper introduction is NOT sufficient

2. **Greeting Check**: Did they greet attendees/participants?

3. **Tone Assessment**: Is the overall tone professional and clear?

4. **Topic Coverage Analysis**:
   - For EACH required topic, check if it was explained in the transcript
   - The explanation doesn't need to be exhaustive - even basic coverage counts as "covered"
   - Look for any mention or explanation of the feature/concept, not necessarily perfect detail
   - A topic is COVERED if the trainer discussed or demonstrated it in ANY meaningful way
   - A topic is MISSING only if it was completely absent or just briefly mentioned without any explanation
   
   Calculate: Coverage % = (Number of Covered Topics / ${totalTopics}) × 100
   
   List ONLY the topics that are genuinely MISSING (not explained at all)

5. **Summary**: Provide a 2-3 sentence summary of the training session's effectiveness

6. **QC Score**: Rate 1-10 based on:
   - Introduction and professionalism (2 points)
   - Topic coverage completeness (5 points)
   - Clarity and tone (3 points)

TRANSCRIPT:
${transcriptText.substring(0, 70000)}

IMPORTANT INSTRUCTIONS:
- Be lenient with topic coverage - if a topic is mentioned and explained even briefly, mark it as covered
- Focus on whether concepts were communicated, not whether they were explained perfectly
- Only list topics as missing if they are truly absent from the discussion
- Consider practical demonstrations and Q&A as valid topic coverage

OUTPUT STRICT JSON FORMAT:
{
  "intro_yes_no": "Yes" or "No",
  "intro_proof": "Exact quote showing trainer's name introduction with Yanolja mention, or 'Not found'",
  "greeting_yes_no": "Yes" or "No",
  "greeting_proof": "Brief quote or description showing greeting",
  "tone_professional": "Yes" or "No",
  "coverage_percentage": numeric value (0-100),
  "missing_topics_list": "Comma-separated list of ONLY genuinely missing topics, or 'None' if all covered",
  "summary_text": "2-3 sentence effectiveness summary",
  "qc_score": numeric value (1-10)
}`;
  
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
      "SUCCESS",
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
    Logger.log(`      📊 Tokens: ${usage.promptTokenCount} input, ${usage.candidatesTokenCount} output`);
    
    return {
      input: usage.promptTokenCount,
      output: usage.candidatesTokenCount
    };
    
  } catch (e) {
    Logger.log(`      ❌ AI Analysis failed: ${e.message}`);
    resultSheet.appendRow([
      docId, record.date, fullTitle, record.duration, record.link,
      "ERROR", `AI Error: ${e.message}`, 
      "", "", "", "", "", "", 0, 0, 0
    ]);
    return null;
  }
}

/**
 * 🌐 CALL GEMINI API
 */
function callGeminiAPI(promptText) {
  if (!CONFIG.GEMINI_API_KEY || CONFIG.GEMINI_API_KEY === '--') {
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

/**
 * 📧 SEND COMPLETION EMAIL
 */
function sendQCCompletionEmail(stats, status = 'SUCCESS', errorMessage = '') {
  if (!CONFIG.SEND_EMAIL_NOTIFICATIONS) {
    Logger.log('📧 Email notifications disabled');
    return;
  }
  
  if (!CONFIG.NOTIFICATION_EMAIL || CONFIG.NOTIFICATION_EMAIL === 'your.email@example.com') {
    Logger.log('⚠️  No valid email configured');
    return;
  }
  
  stats = stats || {};
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
    
    const subject = status === 'SUCCESS' 
      ? `✅ Weekly QC Report - ${currentDate}` 
      : `❌ QC Workflow Failed - ${currentDate}`;
    
    const statusEmoji = status === 'SUCCESS' ? '✅' : '❌';
    
    const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f4f4; margin: 0; padding: 20px; }
    .container { max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); overflow: hidden; }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
    .header h1 { margin: 0; font-size: 24px; font-weight: 600; }
    .header p { margin: 5px 0 0 0; opacity: 0.9; font-size: 14px; }
    .content { padding: 30px; }
    .status-badge { display: inline-block; padding: 8px 16px; border-radius: 20px; font-weight: 600; font-size: 14px; margin-bottom: 20px; }
    .status-success { background-color: #d4edda; color: #155724; }
    .status-failed { background-color: #f8d7da; color: #721c24; }
    .stats-table { width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #f9f9f9; border-radius: 8px; overflow: hidden; }
    .stats-table tr { border-bottom: 1px solid #e0e0e0; }
    .stats-table tr:last-child { border-bottom: none; }
    .stats-table td { padding: 15px; font-size: 14px; }
    .stats-table td:first-child { font-weight: 600; color: #555; width: 60%; }
    .stats-table td:last-child { text-align: right; font-weight: 700; color: #667eea; font-size: 16px; }
    .token-row { background-color: #fff3cd; }
    .cta-button { display: inline-block; padding: 12px 30px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; text-decoration: none; border-radius: 25px; font-weight: 600; margin: 20px 0; text-align: center; }
    .footer { background-color: #f9f9f9; padding: 20px; text-align: center; font-size: 12px; color: #777; border-top: 1px solid #e0e0e0; }
    .error-box { background-color: #f8d7da; border-left: 4px solid #dc3545; padding: 15px; margin: 15px 0; border-radius: 4px; color: #721c24; }
    .highlight { font-weight: 700; color: #667eea; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>${statusEmoji} Training QC Workflow Report</h1>
      <p>${currentDate} • ${currentTime}</p>
    </div>
    
    <div class="content">
      <div class="status-badge ${status === 'SUCCESS' ? 'status-success' : 'status-failed'}">
        ${status === 'SUCCESS' ? '✅ Workflow Completed Successfully' : '❌ Workflow Failed'}
      </div>
      
      ${status === 'FAILED' ? `<div class="error-box"><strong>Error:</strong><br>${errorMessage || 'Unknown error'}</div>` : ''}
      
      <h2 style="color: #333; margin-top: 20px;">📊 Weekly Summary</h2>
      
      <table class="stats-table">
        <tr><td>📅 Total Sessions Scanned</td><td>${stats.totalScanned}</td></tr>
        <tr><td>✅ Passed Validation</td><td>${stats.passedValidation}</td></tr>
        <tr><td>📄 Transcripts Read</td><td>${stats.transcriptsRead}</td></tr>
        <tr><td>🤖 QC Analyses Completed</td><td>${stats.qcAnalyzed}</td></tr>
        <tr class="token-row"><td>🔤 Input Tokens</td><td>${stats.totalInputTokens}</td></tr>
        <tr class="token-row"><td>🔤 Output Tokens</td><td>${stats.totalOutputTokens}</td></tr>
      </table>
      
      ${status === 'SUCCESS' && stats.qcAnalyzed > 0 ? `
      <h3 style="color: #333; margin-top: 25px;">📈 Key Insights</h3>
      <ul style="color: #555; line-height: 1.8;">
        <li><span class="highlight">${stats.passedValidation}</span> of <span class="highlight">${stats.totalScanned}</span> sessions had valid transcripts</li>
        <li>Average tokens per analysis: <span class="highlight">${Math.round((stats.totalInputTokens + stats.totalOutputTokens) / stats.qcAnalyzed)}</span></li>
        <li>Success rate: <span class="highlight">${Math.round((stats.qcAnalyzed / stats.totalScanned) * 100)}%</span></li>
      </ul>
      ` : ''}
      
      <div style="text-align: center; margin-top: 30px;">
        <a href="${sheetUrl}" class="cta-button">📊 View Complete QC Results</a>
      </div>
    </div>
    
    <div class="footer">
      <p>🤖 Automated Training QC System</p>
      <p>Yanolja Cloud Solution Training Hub</p>
    </div>
  </div>
</body>
</html>`;
    
    const plainBody = `
TRAINING QC WORKFLOW REPORT
${status === 'SUCCESS' ? '✅ COMPLETED' : '❌ FAILED'}
${currentDate} • ${currentTime}

${status === 'FAILED' ? `ERROR: ${errorMessage}\n` : ''}
SUMMARY:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Sessions Scanned: ${stats.totalScanned}
Passed Validation: ${stats.passedValidation}
Transcripts Read: ${stats.transcriptsRead}
QC Analyses: ${stats.qcAnalyzed}
Input Tokens: ${stats.totalInputTokens}
Output Tokens: ${stats.totalOutputTokens}

📊 VIEW RESULTS: ${sheetUrl}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`;
    
    MailApp.sendEmail({
      to: CONFIG.NOTIFICATION_EMAIL,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: 'Training QC System'
    });
    
    Logger.log(`✅ Email sent to ${CONFIG.NOTIFICATION_EMAIL}`);
    
  } catch (e) {
    Logger.log(`❌ Email failed: ${e.message}`);
  }
}

// ============================================================================
// SETUP & HELPER FUNCTIONS
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
  let sheet = ss.getSheetByName(CONFIG.QC_RESULTS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.QC_RESULTS_TAB);
    const headers = [
      "Doc ID", "Date", "Title", "Duration", "Link", "Status",
      "Intro (Y/N)", "Intro Proof", "Greeting (Y/N)", "Greeting Proof", 
      "Tone Professional?", "Topic Coverage %", "Missing Topics", 
      "Summary", "QC Score", "Input Tokens", "Output Tokens"
    ];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  return sheet;
}

function getExistingIds(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => row[0]).filter(id => id !== "");
}

function extractSheetId(urlOrId) {
  if (!urlOrId || urlOrId.includes("YOUR_SPREADSHEET")) return null;
  const match = urlOrId.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : urlOrId;
}

function extractFileId(fileObj) {
  if (fileObj.fileId) return fileObj.fileId;
  const match = fileObj.fileUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
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
