/**
 * Cool Finds: saves stars (favorites) and 👍/👎 votes from the live site
 * (https://dorefan.github.io/cool-finds/) into the "Site saves" tab of this sheet.
 *
 * Setup (once):
 *  1. Paste this file into Extensions > Apps Script of the Cool Finds sheet.
 *  2. Project Settings (gear) > Script properties > add PASSCODE = a passcode you choose.
 *  3. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *  4. Copy the web app URL (ends in /exec).
 * The passcode lives only here, never in the public website code.
 */
const TAB = 'Site saves';
const HEAD = ['Kind', 'Key', 'Vote', 'Name', 'Price', 'Shop', 'Category', 'Link', 'Why it is cool', 'Saved at', 'Data'];

function doGet(e) {
  let res;
  try { res = handle_(e.parameter || {}); } catch (err) { res = { ok: false, error: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  if (!pass || p.pass !== pass) return { ok: false, error: 'wrong passcode' };
  const sh = tab_();
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues() : [];

  if (p.action === 'list') {
    const favs = {}, votes = {};
    rows.forEach(r => {
      if (!r[1]) return;
      let d = {}; try { d = JSON.parse(r[10] || '{}'); } catch (err) {}
      if (r[0] === 'fav') favs[r[1]] = d;
      if (r[0] === 'vote' && (r[2] === 'like' || r[2] === 'nope')) votes[r[1]] = r[2];
    });
    return { ok: true, favs, votes };
  }

  const kind = p.kind, key = String(p.key || '').slice(0, 200);
  if ((kind !== 'fav' && kind !== 'vote') || !key) return { ok: false, error: 'bad request' };
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const idx = rows.findIndex(r => r[0] === kind && r[1] === key);
    if (p.action === 'delete') {
      if (idx >= 0) sh.deleteRow(idx + 2);
      return { ok: true };
    }
    if (p.action === 'set') {
      const d = JSON.parse(p.data || '{}');
      const vote = kind === 'vote' ? (d.vote === 'like' || d.vote === 'nope' ? d.vote : '') : '';
      if (kind === 'vote' && !vote) return { ok: false, error: 'bad vote' };
      const row = [kind, key, vote, d.n || d.name || '', d.p || d.price || '', d.s || d.shop || '',
        d.category || d.c || '', d.u || d.link || '', d.w || d.why || '', new Date(), JSON.stringify(d).slice(0, 40000)];
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, HEAD.length).setValues([row]);
      else sh.appendRow(row);
      return { ok: true };
    }
    return { ok: false, error: 'bad action' };
  } finally { lock.releaseLock(); }
}

function tab_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
