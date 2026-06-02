// ═══════════════════════════════════════════════════════════════
// Meeting Poll — Google Apps Script Backend
// Paste this entire file into script.google.com → New project
// Then: Deploy → New deployment → Web app
//   Execute as: Me
//   Who has access: Anyone
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = 'MeetingPollVotes';

// Column layout in the sheet
const COL = {
  TIMESTAMP: 1,
  POLL_ID:   2,
  POLL_TITLE:3,
  VOTER:     4,
  SLOT_ISO:  5,
  SLOT_LABEL:6
};

// ── Routing ──────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action  = payload.action;

    if (action === 'ping')      return respond({ ok: true, message: 'Meeting Poll backend is alive.' });
    if (action === 'saveVote')  return respond(handleSaveVote(payload));
    if (action === 'getVotes')  return respond(handleGetVotes(payload));

    return respond({ error: 'Unknown action: ' + action });
  } catch (err) {
    return respond({ error: err.message });
  }
}

// Also handle GET so a browser tab visit doesn't 404
function doGet(e) {
  return respond({ ok: true, message: 'Meeting Poll backend running. Use POST requests.' });
}

// ── Helpers ───────────────────────────────────────────────────
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
           || SpreadsheetApp.create('MeetingPollVotes');

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'Poll ID', 'Poll Title', 'Voter Name', 'Slot ISO', 'Slot Label']);
    sheet.setFrozenRows(1);
    // Format header row
    const header = sheet.getRange(1, 1, 1, 6);
    header.setFontWeight('bold');
    header.setBackground('#e8f0fb');
  }
  return sheet;
}

// ── Save a vote ───────────────────────────────────────────────
// payload: { pollId, pollTitle, voterName, slots: [isoString, ...] }
// Each selected slot gets its own row. If the voter already exists
// for this poll, their old votes are removed and replaced (idempotent).
function handleSaveVote(payload) {
  const { pollId, pollTitle, voterName, slots } = payload;
  if (!pollId)    throw new Error('Missing pollId');
  if (!voterName) throw new Error('Missing voterName');
  if (!Array.isArray(slots)) throw new Error('slots must be an array');

  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  // Delete existing rows for this voter + poll (rows are 1-indexed; row 1 is header)
  // Iterate backwards so row deletions don't shift indices
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][COL.POLL_ID - 1])  === String(pollId) &&
        String(data[i][COL.VOTER - 1])    === String(voterName)) {
      sheet.deleteRow(i + 1);
    }
  }

  // Append a row for each selected slot
  const ts = new Date().toISOString();
  slots.forEach(slotIso => {
    sheet.appendRow([ts, pollId, pollTitle || '', voterName, slotIso, '']);
  });

  return { ok: true, saved: slots.length };
}

// ── Get votes for a poll ──────────────────────────────────────
// payload: { pollId }
// Returns: { votes: [{ voter, slotIso }, ...] }
function handleGetVotes(payload) {
  const { pollId } = payload;
  if (!pollId) throw new Error('Missing pollId');

  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();

  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (String(row[COL.POLL_ID - 1]) === String(pollId)) {
      result.push({
        voter:   String(row[COL.VOTER    - 1]),
        slotIso: String(row[COL.SLOT_ISO - 1])
      });
    }
  }

  return { ok: true, votes: result };
}
