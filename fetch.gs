/**
 * CONFIGURATION
 */
const CONFIG = {
  CALENDAR_ID: 'primary',
  DAYS_TO_LOOK_BACK: 14, 
  INCLUDE_REGEX: /(Hindi|English)/i, 
  EXCLUDE_REGEX: /with/i,
  MIN_LENGTH_CHARS: 500 // Min chars to count as a valid session
};

function logMeetingTranscripts() {
  const now = new Date();
  
  // 1. Date Setup
  const startDate = new Date();
  startDate.setDate(now.getDate() - CONFIG.DAYS_TO_LOOK_BACK);
  const futureDate = new Date();
  futureDate.setFullYear(now.getFullYear() + 2); 

  Logger.log(`=== SCANNING FOR COMPLETED SESSIONS ===`);
  Logger.log(`Range: ${startDate.toDateString()} to Now (including future scheduled dates)`);

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

  // --- COUNTER INITIALIZATION ---
  let totalValidSessions = 0;

  events.forEach(event => {
    const title = event.summary || ""; 
    
    // Filters
    if (CONFIG.EXCLUDE_REGEX.test(title)) return;
    if (!CONFIG.INCLUDE_REGEX.test(title)) return;
    if (!event.attachments || event.attachments.length === 0) return;

    let headerPrinted = false;
    let eventHasValidTranscript = false; // Flag for this specific event

    for (const file of event.attachments) {
      
      // Look for Google Docs
      if (file.mimeType === 'application/vnd.google-apps.document') {

        if (!headerPrinted) {
          Logger.log(`\n--------------------------------------------------`);
          Logger.log(`EVENT: "${title}"`); 
          Logger.log(`DATE:  ${event.start.dateTime || event.start.date}`);
          headerPrinted = true;
        }

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
          if (totalLength < CONFIG.MIN_LENGTH_CHARS) {
            Logger.log(`   ⚠️ SKIPPED FILE: "${file.title}" (Too short: ${totalLength} chars)`);
            continue;
          }

          // Generate Preview
          const previewLines = fullText.split('\n')
            .filter(line => line.trim() !== '') 
            .slice(0, 3)                        
            .join('\n     ');                   

          Logger.log(`   ✅ VALID TRANSCRIPT: "${file.title}"`);
          Logger.log(`     📏 LENGTH: ${totalLength} chars`);
          Logger.log(`     📝 PREVIEW:\n     ${previewLines}`);

          // Mark this event as valid
          eventHasValidTranscript = true;

        } catch (e) {
          Logger.log(`     ❌ EXPORT ERROR: ${e.message}`);
        }
      }
    }

    // Increment Counter if at least one valid transcript was found for this event
    if (eventHasValidTranscript) {
      totalValidSessions++;
    }
  });
  
  // --- FINAL SUMMARY ---
  Logger.log(`\n==============================================`);
  Logger.log(`📊 SUMMARY REPORT`);
  Logger.log(`----------------------------------------------`);
  Logger.log(`Total Events Scanned: ${events.length}`);
  Logger.log(`Total Valid Sessions: ${totalValidSessions}`);
  Logger.log(`(Matches Title + Has Transcript + Sufficient Length)`);
  Logger.log(`==============================================`);
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
