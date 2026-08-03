/**
 * RSVP backend for the Anis & Atiq invitation card.
 *
 * ── SETUP ──────────────────────────────────────────────────────────
 *  1. Create a new Google Sheet. You do NOT need to rename anything or
 *     type any headers — running setup() below does all of that.
 *
 *  2. Copy the Sheet's id from its URL — the long string between
 *     /d/ and /edit :
 *        docs.google.com/spreadsheets/d/THIS_PART_HERE/edit
 *     Paste it into SHEET_ID below, between the quote marks.
 *
 *  3. Extensions → Apps Script. Delete everything in the editor,
 *     paste this whole file, and save (the disk icon).
 *
 *  4. In the toolbar dropdown that says "doGet", choose "setup"
 *     and press Run. Approve the permissions when asked.
 *     Check the Sheet: you should now have a tab named RSVP with
 *     headers in row 1, and one test row you can delete.
 *
 *  5. Deploy → New deployment → gear icon → Web app
 *        Execute as:      Me
 *        Who has access:  Anyone     ← must be "Anyone", NOT
 *                                      "Anyone with Google account"
 *     Deploy, then copy the Web app URL ending in /exec
 *
 *  6. In index.html find:   const RSVP_API = "";
 *     Paste the URL between the quotes and push the file.
 *
 * ── AFTER EDITING THIS FILE ────────────────────────────────────────
 *  Deploy → Manage deployments → pencil → Version: New version → Deploy.
 *  The /exec URL does not change.
 * ───────────────────────────────────────────────────────────────────
 */

const SHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';
const TAB      = 'RSVP';
const HEADERS  = ['Masa', 'Nama', 'Bilangan', 'Hadir', 'Ucapan'];

/**
 * Finds the RSVP tab, creating it if needed. If the spreadsheet only has
 * its original default tab (Sheet1 / Helaian1) and nothing else, that tab
 * is renamed rather than leaving an empty one behind.
 */
function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(TAB);

  if (!sh) {
    const all = ss.getSheets();
    if (all.length === 1 && all[0].getLastRow() === 0) {
      sh = all[0].setName(TAB);          // reuse the empty default tab
    } else {
      sh = ss.insertSheet(TAB);
    }
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Run this once from the editor to create the tab and confirm it works. */
function setup() {
  const sh = sheet_();
  sh.appendRow([new Date(), 'Ujian', 2, 'Ya', 'Baris ujian — boleh padam']);
  Logger.log('OK. Tab "%s" ready in "%s". Delete the test row when you like.',
             sh.getName(), sh.getParent().getName());
}

/** Returns the attendee list, newest first. */
function doGet() {
  const rows = sheet_().getDataRange().getValues();
  rows.shift();                       // drop the header row

  const out = rows
    .filter(r => String(r[1]).trim() !== '')
    .map(r => ({
      nama:   String(r[1]),
      pax:    Number(r[2]) || 1,
      hadir:  r[3] === true || String(r[3]).toLowerCase() === 'ya',
      ucapan: String(r[4] || '')
    }))
    .reverse();

  return json_(out);
}

/** Records one RSVP submission. */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    const nama = String(d.nama || '').trim().slice(0, 80);
    if (nama.length < 2) {
      return json_({ ok: false, error: 'Nama tidak sah' });
    }

    const hadir = d.hadir === true;
    const pax   = hadir ? Math.min(30, Math.max(1, Number(d.pax) || 1)) : 0;

    sheet_().appendRow([
      new Date(),
      nama,
      pax,
      hadir ? 'Ya' : 'Tidak',
      String(d.ucapan || '').trim().slice(0, 220)
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
