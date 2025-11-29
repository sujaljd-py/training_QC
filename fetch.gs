/**
 * CONFIGURATION
 */
const CONFIG = {
  SPREADSHEET_ID: '15Z_VEjGlSuxF6fD4J1m6x6E7u-p8HVBM0Z_rxAGxn4I', // <--- PASTE YOUR SHEET ID HERE
  SHEET_TAB_NAME: 'FetchLogs',
  
  CALENDAR_ID: 'primary',
  DAYS_TO_LOOK_BACK: 14, 
  INCLUDE_REGEX: /(Hindi|English)/i, 
  EXCLUDE_REGEX: /with/i,
  MIN_LENGTH_CHARS: 500 
};

function logMeetingTranscripts() {
  const now = new Date();
  
  // 1. Setup Dates
  const startDate = new Date();
  startDate.setDate(now.getDate() - CONFIG.DAYS_TO_LOOK_BACK);
  const futureDate = new Date();
  futureDate.setFullYear(now.getFullYear() + 2); 

  Logger.log(`=== STARTING FETCH & LOG PROCESS ===`);

  // 2. FETCH EVENTS
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
    Logger.log(`❌ ERROR: Could not list events. Check Permissions.`);
    return;
  }

  if (!events || events.length === 0) {
    Logger.log("No events found.");
    return;
  }

  // 3. INITIALIZE DATA ARRAY
  // We will push all valid rows here
  let rowsToLog = [];
  let totalValidSessions = 0;

  // 4. PROCESS EVENTS
  events.forEach(event => {
    const title = event.summary || ""; 
    
    // Filters
    if (CONFIG.EXCLUDE_REGEX.test(title)) return;
    if (!CONFIG.INCLUDE_REGEX.test(title)) return;
    if (!event.attachments || event.attachments.length === 0) return;

    for (const file of event.attachments) {
      
      // Look for Google Docs
      if (file.mimeType === 'application/vnd.google-apps.document') {

        // Get ID
        let fileId = file.fileId;
        if (!fileId) {
             const match = file.fileUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
             if (match) fileId = match[1];
        }

        if (!fileId) continue;

        try {
          // DOWNLOAD TEXT
          const fullText = exportDocAsText(fileId);
          const totalLength = fullText.length;

          // Check Length
          if (totalLength < CONFIG.MIN_LENGTH_CHARS) continue;

          // --- GATHER DATA POINTS ---

          // A. Duration Estimate
          const wordCount = fullText.split(/\s+/).length;
          const estimatedDuration = Math.round(wordCount / 130);

          // B. Attendees
          let attendeeList = "Unknown";
          if (event.attendees && event.attendees.length > 0) {
             attendeeList = event.attendees
               .map(a => a.displayName || a.email) // Use Name, fallback to Email
               .join(', ');
          }

          // C. Event Date
          const eventDate = new Date(event.start.dateTime || event.start.date);

          // PUSH TO DATA ARRAY
          rowsToLog.push([
            eventDate,           // Date
            title,               // Title
            estimatedDuration,   // Duration (Mins)
            totalLength,         // Total Transcript Characters
            fileId,              // Document ID
            attendeeList,        // Attendee Names
            file.title,          // Transcript File Name
            file.fileUrl         // Transcript File Link
          ]);

          Logger.log(`✅ QUEUED: ${title} (${estimatedDuration} mins)`);
          totalValidSessions++;

        } catch (e) {
          Logger.log(`❌ ERROR processing ${title}: ${e.message}`);
        }
      }
    }
  });

  // 5. WRITE TO SHEET
  if (rowsToLog.length > 0) {
    saveToSheet(rowsToLog);
  } else {
    Logger.log("No valid training sessions found to log.");
  }
  
  // FINAL SUMMARY
  Logger.log(`\n==============================================`);
  Logger.log(`SUMMARY: ${totalValidSessions} Sessions Logged to Sheet.`);
  Logger.log(`==============================================`);
}

/**
 * HELPER: Write data to Google Sheet
 */
function saveToSheet(newRows) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_TAB_NAME);

  // Create Sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_TAB_NAME);
    // Add Headers
    sheet.appendRow([
      "Training Date", 
      "Training Title", 
      "Duration (Mins)", 
      "Char Count", 
      "Doc ID", 
      "Attendees", 
      "File Name", 
      "File Link"
    ]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 8).setFontWeight("bold");
    Logger.log(`Created new tab: ${CONFIG.SHEET_TAB_NAME}`);
  }

  // Get next empty row
  const lastRow = sheet.getLastRow();
  const nextRow = lastRow + 1;

  // Write Batch
  sheet.getRange(nextRow, 1, newRows.length, newRows[0].length).setValues(newRows);
  Logger.log(`Written ${newRows.length} rows to spreadsheet.`);
}

/**
 * HELPER: Downloads a Google Doc as plain text.
 */
function exportDocAsText(fileId) {
  const url = `https://docs.google.com/feeds/download/documents/export/Export?id=${fileId}&exportFormat=txt`;
  
  const options = {
    method: "GET",
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`Failed to download text (HTTP ${response.getResponseCode()})`);
  }
  
  return response.getContentText();
}
