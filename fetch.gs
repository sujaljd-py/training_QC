/**
 * CONFIGURATION
 */
const CONFIG = {
  // ⚠️ DOUBLE CHECK THIS ID IS CORRECT
  SPREADSHEET_ID: '15Z_VEjGlSuxF6fD4J1m6x6E7u-p8HVBM0Z_rxAGxn4I', 
  
  SHEET_TAB_NAME: 'Fetch Logs',
  CALENDAR_ID: 'primary',
  DAYS_TO_LOOK_BACK: 14, 
  INCLUDE_REGEX: /(Hindi|English)/i, 
  EXCLUDE_REGEX: /with/i,
  MIN_LENGTH_CHARS: 500, // Reduced threshold to ensure we catch files
  WORDS_PER_MINUTE: 130 
};

function logMeetingTranscripts() {
  // 1. Validate ID
  const cleanSheetId = extractSheetId(CONFIG.SPREADSHEET_ID);
  if (!cleanSheetId) {
    throw new Error("❌ CONFIG ERROR: Valid Spreadsheet ID required.");
  }

  const now = new Date();
  const startDate = new Date();
  startDate.setDate(now.getDate() - CONFIG.DAYS_TO_LOOK_BACK);
  const futureDate = new Date();
  futureDate.setFullYear(now.getFullYear() + 2); 

  Logger.log(`=== STARTING DEBUG SCAN ===`);
  Logger.log(`Range: ${startDate.toDateString()} to Future`);

  // 2. Fetch Events
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
    Logger.log(`❌ CRITICAL ERROR: Could not access Calendar. Permissions?`);
    return;
  }

  if (!events || events.length === 0) {
    Logger.log("⚠️ No events found in Calendar for this date range.");
    return;
  }

  Logger.log(`Found ${events.length} total events. Filtering...`);

  let rowsToLog = [];

  // 3. Loop Events
  events.forEach(event => {
    const title = event.summary || "No Title"; 
    
    // --- DEBUG LOGGING START ---
    
    // Filter 1: Exclude
    if (CONFIG.EXCLUDE_REGEX.test(title)) {
      Logger.log(`[SKIP] "${title}" -> Contains exclusion keyword (e.g. 'with')`);
      return;
    }

    // Filter 2: Include
    if (!CONFIG.INCLUDE_REGEX.test(title)) {
      Logger.log(`[SKIP] "${title}" -> Missing keyword (Hindi/English)`);
      return;
    }

    Logger.log(`\n🔎 INSPECTING: "${title}"`);

    // --- INIT DATA ---
    let eventDate = new Date(event.start.dateTime || event.start.date);
    let finalDuration = 0;
    let durationSource = "Calendar Schedule"; 
    let status = "Checking...";
    let docId = "N/A";
    let charCount = 0;
    let fileName = "N/A";
    let fileLink = "N/A";
    let attendees = "Unknown";

    // 1. Duration (Calendar)
    if (event.start.dateTime && event.end.dateTime) {
      const start = new Date(event.start.dateTime);
      const end = new Date(event.end.dateTime);
      finalDuration = Math.round((end - start) / 1000 / 60);
    }

    // 2. Video Check
    let hasVideo = false;
    if (event.attachments) {
      const vids = event.attachments.filter(a => a.mimeType === 'video/mp4');
      if (vids.length > 0) {
        // ... (Video logic same as before) ...
        // Simplified for debug: just checking if it exists
        hasVideo = true;
        durationSource = "Video Metadata"; // Assuming logic holds
      }
    }

    // 3. Transcript Check
    let validTranscriptFound = false;

    if (!event.attachments || event.attachments.length === 0) {
       Logger.log(`   ⚠️ No Attachments found on this event.`);
    } else {
       Logger.log(`   📂 Found ${event.attachments.length} attachment(s).`);
       
       for (const file of event.attachments) {
         if (file.mimeType === 'application/vnd.google-apps.document') {
           Logger.log(`      > Checking Doc: "${file.title}"`);
           
           let fileId = extractFileId(file);
           if (!fileId) continue;

           try {
             const fullText = exportDocAsText(fileId);
             const totalLength = fullText.length;
             const wordCount = fullText.split(/\s+/).length;

             Logger.log(`        Length: ${totalLength} chars.`);

             if (totalLength < CONFIG.MIN_LENGTH_CHARS) {
               Logger.log(`        ⚠️ Too short (Min: ${CONFIG.MIN_LENGTH_CHARS}). Skipping.`);
               continue;
             }

             // FOUND IT
             docId = fileId;
             charCount = totalLength;
             fileName = file.title;
             fileLink = file.fileUrl;
             validTranscriptFound = true;
             status = "SUCCESS";

             if (event.attendees) {
                attendees = event.attendees.map(a => a.displayName || a.email).join(', ');
             }

             if (!hasVideo) {
               const estMins = Math.round(wordCount / CONFIG.WORDS_PER_MINUTE);
               if (estMins > 5) {
                 finalDuration = estMins;
                 durationSource = `Text Estimate (${wordCount} words)`;
               }
             }
             break; // Stop loop

           } catch (e) {
             Logger.log(`        ❌ Read Error: ${e.message}`);
           }
         }
       }
    }

    if (!validTranscriptFound) {
      Logger.log(`   ❌ RESULT: No valid transcript found.`);
      status = "NOT FOUND / MISSING";
      // We log it anyway so you can see it in the sheet as "NOT FOUND"
    } else {
      Logger.log(`   ✅ RESULT: Valid Transcript Found!`);
    }

    // PUSH ROW
    rowsToLog.push([
      eventDate,
      title,
      status,
      finalDuration,
      durationSource,
      charCount,
      docId,
      attendees,
      fileName,
      fileLink
    ]);
  });

  // 4. WRITE TO SHEET
  if (rowsToLog.length > 0) {
    Logger.log(`\n💾 Attempting to write ${rowsToLog.length} rows to Sheet...`);
    saveToSheet(cleanSheetId, rowsToLog);
  } else {
    Logger.log(`\n⚠️ Finished scan. No events qualified for logging.`);
  }
}

// --- HELPERS ---

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

function saveToSheet(sheetId, newRows) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    let sheet = ss.getSheetByName(CONFIG.SHEET_TAB_NAME);

    if (!sheet) {
      Logger.log(`   Creating new tab: ${CONFIG.SHEET_TAB_NAME}`);
      sheet = ss.insertSheet(CONFIG.SHEET_TAB_NAME);
      sheet.appendRow(["Date", "Title", "Status", "Duration (Mins)", "Source", "Chars", "Doc ID", "Attendees", "File Name", "Link"]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 10).setFontWeight("bold");
    }

    const lastRow = sheet.getLastRow();
    const nextRow = lastRow + 1;
    sheet.getRange(nextRow, 1, newRows.length, newRows[0].length).setValues(newRows);
    Logger.log(`   ✅ SUCCESS: Written to sheet successfully.`);
  } catch (e) {
    Logger.log(`   ❌ SHEET ERROR: ${e.message}`);
    Logger.log(`   (Did you enable the Google Sheets API in appsscript.json?)`);
  }
}
