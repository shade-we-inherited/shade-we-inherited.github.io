/**
 * Shade We Inherited: backend for the tree pin map.
 * New pins arrive by POST and are added to the Sheet unapproved. The public map (GET) only receives rows
 * whose "approved" box is ticked, so nothing appears until the owner checks it.
 * Each pin carries a private edit token chosen by the contributor's browser; only its SHA-256 hash is stored.
 * Edits put the pin back to unapproved while the last approved version (column "live") stays on the map.
 * Run setup() once, then deploy as a web app (Execute as: Me, Who has access: Anyone).
 */
const SHEET_NAME = 'pins';
const HEAD = ['timestamp', 'lat', 'lon', 'tree', 'story', 'status', 'name', 'country', 'approved', 'token_hash', 'removed', 'live', 'edited'];
const C = {}; HEAD.forEach((h, i) => C[h] = i);          // column index (0-based)
const FIELDS = ['lat', 'lon', 'tree', 'story', 'status', 'name', 'country'];

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) { ensureHead_(); Logger.log('Already set up: ' + SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getUrl()); return; }
  const ss = SpreadsheetApp.create('Shade We Inherited pins');
  const sh = ss.getSheets()[0].setName(SHEET_NAME);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(5, 420);
  props.setProperty('SHEET_ID', ss.getId());
  Logger.log('Sheet: ' + ss.getUrl());
}

function sheet_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName(SHEET_NAME); }
function ensureHead_() {
  const sh = sheet_(); const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  if (cur.length < HEAD.length || HEAD.some((h, i) => cur[i] !== h)) sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
}
function clean_(v, n) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n);
  if (/^[=+\-@]/.test(s)) s = "'" + s;          // keep spreadsheet formulas out
  return s;
}
function hash_(t) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(t)).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function unq_(s) { return String(s == null ? '' : s).replace(/^'/, ''); }
function fields_(p) {
  const lat = Number(p.lat), lon = Number(p.lon);
  if (!isFinite(lat) || !isFinite(lon) || Math.abs(lat) > 90 || Math.abs(lon) > 180) return null;
  const story = clean_(p.story, 300); if (story.length < 3) return null;
  return { lat: Math.round(lat * 1e5) / 1e5, lon: Math.round(lon * 1e5) / 1e5, tree: clean_(p.tree, 80), story: story,
           status: ['standing', 'gone', 'unsure'].indexOf(p.status) >= 0 ? p.status : 'unsure', name: clean_(p.name, 40), country: clean_(p.country, 50) };
}
function findRow_(sh, h) {
  const col = sh.getRange(2, C.token_hash + 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < col.length; i++) if (col[i][0] === h) return i + 2;
  return 0;
}
function notify_(subject, f, sh) {
  try {
    MailApp.sendEmail(Session.getEffectiveUser().getEmail(), subject,
      f.story + '\n\n' + [f.name, f.country].filter(String).join(', ') + '\nhttps://www.google.com/maps?q=' + f.lat + ',' + f.lon +
      '\n\nTick "approved" to show it on the map:\n' + sh.getParent().getUrl());
  } catch (err) {}
}
function publicPin_(r) {
  return { lat: r[C.lat], lon: r[C.lon], tree: unq_(r[C.tree]), story: unq_(r[C.story]), status: r[C.status], name: unq_(r[C.name]), country: unq_(r[C.country]),
           pid: String(r[C.token_hash] || '').slice(0, 12) };
}

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.website) return json_({ ok: true });     // honeypot field filled by bots
    const action = p.action || 'add';
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    try {
      ensureHead_();
      const sh = sheet_();
      if (action === 'add') {
        const f = fields_(p); if (!f) return json_({ ok: false });
        const h = p.token && String(p.token).length >= 20 ? hash_(p.token) : '';
        if (h && findRow_(sh, h)) return json_({ ok: true });      // duplicate send
        const row = [new Date(), f.lat, f.lon, f.tree, f.story, f.status, f.name, f.country, false, h, false, '', ''];
        sh.appendRow(row);
        sh.getRange(sh.getLastRow(), C.approved + 1).insertCheckboxes();
        notify_('New tree pin: ' + (f.tree || 'a tree'), f, sh);
      } else if (action === 'edit' || action === 'remove') {
        if (!p.token || String(p.token).length < 20) return json_({ ok: false });
        const n = findRow_(sh, hash_(p.token)); if (!n) return json_({ ok: false });
        const rng = sh.getRange(n, 1, 1, HEAD.length); const r = rng.getValues()[0];
        if (r[C.removed] === true) return json_({ ok: false });
        if (action === 'remove') {
          sh.getRange(n, C.removed + 1).setValue(true); sh.getRange(n, C.edited + 1).setValue(new Date());
        } else {
          const f = fields_(p); if (!f) return json_({ ok: false });
          if (r[C.approved] === true) {                           // keep the approved version on the map until the edit is checked
            const live = {}; FIELDS.forEach(k => live[k] = r[C[k]]); r[C.live] = JSON.stringify(live);
          }
          FIELDS.forEach(k => r[C[k]] = f[k]); r[C.approved] = false; r[C.edited] = new Date();
          rng.setValues([r]);
          notify_('Edited tree pin: ' + (f.tree || 'a tree'), f, sh);
        }
      }
      CacheService.getScriptCache().remove('pins');
      return json_({ ok: true });
    } finally { lock.releaseLock(); }
  } catch (err) { return json_({ ok: false }); }
}

function doGet(e) {
  const q = (e && e.parameter) || {};
  if (q.action === 'mine' && q.token) {                      // a contributor loading their own pin to edit
    const sh = sheet_(); const n = findRow_(sh, hash_(q.token));
    if (!n) return json_({ ok: false });
    const r = sh.getRange(n, 1, 1, HEAD.length).getValues()[0];
    if (r[C.removed] === true) return json_({ ok: false, removed: true });
    const pin = publicPin_(r); pin.approved = r[C.approved] === true;
    return json_({ ok: true, pin: pin });
  }
  const cache = CacheService.getScriptCache(); const hit = cache.get('pins');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const rows = sheet_().getDataRange().getValues().slice(1);
  const out = [];
  rows.forEach(r => {
    if (r[C.removed] === true) return;
    if (r[C.approved] === true) { out.push(publicPin_(r)); return; }
    if (r[C.live]) {                                          // edited since approval: show the last approved version
      try { const l = JSON.parse(r[C.live]); const pin = publicPin_(r); FIELDS.forEach(k => pin[k] = unq_(l[k])); pin.lat = l.lat; pin.lon = l.lon; out.push(pin); } catch (err) {}
    }
  });
  const s = JSON.stringify(out);
  cache.put('pins', s, 60);
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
