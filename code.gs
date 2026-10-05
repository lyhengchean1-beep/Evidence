const SHEET_NAME = 'Evidence';                 // tab with headers: No, Date, Type, Before, After, Remark, Notes
const FOLDER_ID = 'PASTE_DRIVE_FOLDER_ID_HERE'; // Drive folder that stores the photos
const PUBLIC_LINKS = true;                      // true = anyone with the link can view the photos

function doGet() {
  return json({ ok: true, message: 'Probation evidence endpoint is running.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const d = JSON.parse(e.postData.contents);
    if (!d.type) throw new Error('Type is required.');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found.');

    // Next No = last row's No + 1
    const last = sheet.getLastRow();
    const no = last < 2 ? 1 : (Number(sheet.getRange(last, 1).getValue()) || last - 1) + 1;

    const folder = DriveApp.getFolderById(FOLDER_ID);
    const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd-HHmmss');
    const beforeUrl = saveImage(d.before, folder, 'No' + no + '-before-' + stamp);
    const afterUrl = saveImage(d.after, folder, 'No' + no + '-after-' + stamp);

    sheet.appendRow([no, new Date(), safe(d.type), beforeUrl, afterUrl, safe(d.remark), safe(d.notes)]);
    sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('yyyy-mm-dd hh:mm');

    return json({ ok: true, no: no });
  } catch (err) {
    return json({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function saveImage(dataUrl, folder, name) {
  if (!dataUrl) return '';
  const parts = dataUrl.split(',');
  if (parts.length !== 2 || parts[0].indexOf('data:image/') !== 0) throw new Error('Invalid image data.');
  const mime = parts[0].slice(5, parts[0].indexOf(';'));
  const ext = mime.split('/')[1].replace('jpeg', 'jpg');
  const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), mime, name + '.' + ext);
  const file = folder.createFile(blob);
  if (PUBLIC_LINKS) file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// Stops text starting with = + - @ from being evaluated as a formula
function safe(v) {
  v = String(v || '');
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}