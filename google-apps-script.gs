/**
 * STOODANT · Class 10 Early Access → Google Sheet
 * 1. Open your "STOODANT Leads" Google Sheet → Extensions → Apps Script.
 * 2. Delete everything there, paste this whole file, click Save.
 * 3. Deploy → New deployment → type "Web app".
 *      Execute as: Me      Who has access: Anyone
 * 4. Click Deploy, allow access, copy the Web app URL (ends with /exec).
 * 5. Paste that URL into SCRIPT_URL at the bottom of index.html.
 */

const SHEET_NAME   = 'Leads';
const NOTIFY_EMAIL = 'Jiteshsain20@gmail.com';   

const HEADERS = [
  'Timestamp', 'Lead ID', 'Status', 'Counsellor', 'Filled by', 'Student', 'Board', 'City', 'School',
  'Subjects', 'Struggles', 'Last exam %', 'Target %', 'Format', 'Time slots',
  'Parent name', 'WhatsApp', 'Consent', 'utm_source', 'utm_medium', 'utm_campaign',
  'utm_content', 'fbclid', 'Page', 'First contact at', 'Demo date', 'Notes'
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const d  = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#F8B82E');
    }
    const clean = v => String(v == null ? '' : v).replace(/^[=+\-@]/, "'$&").slice(0, 300);
    sh.appendRow([
      new Date(), d.leadId, 'New', '', d.filledBy, d.studentName, d.board, d.city, d.school,
      d.subjects, d.struggles, d.lastBand, d.target, d.format, d.slots,
      d.parentName, d.phone, d.consent, d.utm_source, d.utm_medium, d.utm_campaign,
      d.utm_content, d.fbclid, d.page, '', '', ''
    ].map((v, i) => i === 0 ? v : clean(v)));

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(NOTIFY_EMAIL, 'New Class 10 lead: ' + d.studentName + ' (' + d.format + ')',
        [d.parentName, d.phone, d.board + ', ' + d.city, d.subjects, 'Target ' + d.target, d.slots, 'Ref ' + d.leadId].join('\n'));
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Run once from the editor to create the header row before your first lead.
function setup() {
  doPost({ postData: { contents: JSON.stringify({ leadId: 'TEST-ROW', studentName: 'Test', format: 'Test', consent: 'Yes' }) } });
}
