/**
 * Survey response receiver (Google Apps Script web app).
 *
 * Accepts a response either as GET ?d=<json>&cb=<callback> (JSONP, so the
 * survey page can confirm the row was saved) or as POST with a JSON body.
 * Appends one row per response to the "responses" sheet, which is the first
 * sheet of the spreadsheet. Stores no IP address, name or other identifier.
 * A repeated response_id is ignored, so a resent response cannot create a
 * duplicate row.
 */
const SHEET_NAME = 'responses';
const COLUMNS = [
  'response_id', 'received_at', 'version', 'duration_sec', 'scenario_order',
  'consent', 'age18',
  'D1', 'D2', 'D3', 'D4', 'D5',
  'B_M', 'B_H', 'B_E',
  'S_M', 'S_H', 'S_E',
  'C_M1', 'C_M2', 'C_H1', 'C_H2', 'C_E1', 'C_E2',
  'G1', 'G2', 'G3', 'G4', 'AC1', 'G5', 'G6', 'G7', 'G8',
  'F_M', 'F_H', 'F_E',
  'K1', 'K2', 'K3',
  'K4', 'K5', 'K6', 'K7', 'K8', 'K9', 'K10', 'K11'
];
const ID_PATTERN = /^[0-9a-f-]{16,40}$/i;
const VALUE_PATTERN = /^(T|F|DK|[0-9]{1,2})$/;
const CALLBACK_PATTERN = /^[A-Za-z_][A-Za-z0-9_]{0,40}$/;

function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = CALLBACK_PATTERN.test(p.cb || '') ? p.cb : '';
  if (!p.d) return reply({ ok: true, status: 'running' }, cb);
  let data;
  try { data = JSON.parse(p.d); } catch (err) { return reply({ ok: false, error: 'bad_json' }, cb); }
  return reply(save(data), cb);
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return reply({ ok: false, error: 'bad_json' }); }
  return reply(save(data));
}

function save(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return { ok: false, error: 'busy' };
  }
  try {
    const id = String((data && data.response_id) || '');
    if (!ID_PATTERN.test(id)) return { ok: false, error: 'bad_id' };

    const sheet = getSheet();
    const last = sheet.getLastRow();
    if (last > 1) {
      const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
      if (ids.indexOf(id) !== -1) return { ok: true, duplicate: true };
    }

    const row = COLUMNS.map(function (col) {
      if (col === 'response_id') return id;
      if (col === 'received_at') return new Date();
      const v = data[col];
      if (v === undefined || v === null) return '';
      const s = String(v).slice(0, 20);
      if (col === 'version' || col === 'scenario_order') return s.replace(/[^0-9A-Za-z.\-]/g, '');
      if (col === 'duration_sec') return Number(s) || '';
      return VALUE_PATTERN.test(s) ? s : '';
    });
    sheet.appendRow(row);
    return { ok: true, row: sheet.getLastRow() };
  } catch (err) {
    return { ok: false, error: 'server_error', detail: String(err).slice(0, 200) };
  } finally {
    lock.releaseLock();
  }
}

/** Uses the "responses" sheet; if missing, renames the first sheet when it is empty, else adds one. */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const first = ss.getSheets()[0];
    if (first && first.getLastRow() === 0) {
      sheet = first.setName(SHEET_NAME);
    } else {
      sheet = ss.insertSheet(SHEET_NAME, 0);
    }
  }
  if (sheet.getLastRow() === 0) sheet.appendRow(COLUMNS);
  return sheet;
}

/** Run once from the editor (select testSave, press Run) to check that rows can be written. */
function testSave() {
  const result = save({ response_id: Utilities.getUuid(), version: 'test' });
  Logger.log(JSON.stringify(result));
}

function reply(obj, cb) {
  const json = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
