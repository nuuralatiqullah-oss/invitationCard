/**
 * RSVP backend for the Anis & Atiq invitation card.
 *
 * SETUP (about 5 minutes, free):
 *  1. Create a new Google Sheet. Name the first tab: RSVP
 *  2. Row 1 headers, left to right:
 *       Masa | Nama | Telefon | Bilangan | Hadir | Ucapan
 *  3. Extensions → Apps Script. Delete the placeholder, paste this file.
 *  4. Replace SHEET_ID below with the long id from your Sheet's URL:
 *       docs.google.com/spreadsheets/d/<THIS PART>/edit
 *  5. Deploy → New deployment → type "Web app"
 *       Execute as:        Me
 *       Who has access:    Anyone
 *  6. Copy the /exec URL it gives you and paste it into RSVP_API
 *     at the top of the <script> block in the invitation HTML.
 *
 * Re-deploy (Manage deployments → edit → new version) after any edit here.
 */

const SHEET_ID = 'PASTE_YOUR_SHEET_ID_HERE';
const TAB      = 'RSVP';

function sheet_() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB);
}

function json_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Returns the attendee list, newest first. */
function doGet() {
  const rows = sheet_().getDataRange().getValues();
  rows.shift(); // drop header

  const out = rows
    .filter(r => String(r[1]).trim() !== '')
    .map(r => ({
      nama:   String(r[1]),
      pax:    Number(r[3]) || 1,
      hadir:  r[4] === true || String(r[4]).toLowerCase() === 'ya',
      ucapan: String(r[5] || '')
      // Telefon is deliberately NOT returned — it stays private in the Sheet.
    }))
    .reverse();

  return json_(out);
}

/** Records one RSVP submission. */
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);

    const nama = String(d.nama || '').trim().slice(0, 80);
    const tel  = String(d.tel  || '').trim().slice(0, 25);
    if (nama.length < 2 || tel.length < 7) {
      return json_({ ok: false, error: 'Nama atau telefon tidak sah' });
    }

    const hadir = d.hadir === true;
    const pax   = hadir ? Math.min(30, Math.max(1, Number(d.pax) || 1)) : 0;

    sheet_().appendRow([
      new Date(),
      nama,
      "'" + tel,                              // leading quote keeps the 0 prefix
      pax,
      hadir ? 'Ya' : 'Tidak',
      String(d.ucapan || '').trim().slice(0, 220)
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
