/**
 * Shade We Inherited: backend for the tree pin map.
 * New pins arrive by POST and are added to the Sheet unapproved. The public map (GET) only receives rows
 * whose "approved" box is ticked, so nothing appears until the owner checks it.
 * Run setup() once, then deploy as a web app (Execute as: Me, Who has access: Anyone).
 */
const SHEET_NAME = 'pins';
const HEAD = ['timestamp', 'lat', 'lon', 'tree', 'story', 'status', 'name', 'country', 'approved'];

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) { Logger.log('Already set up: ' + SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getUrl()); return; }
  const ss = SpreadsheetApp.create('Shade We Inherited pins');
  const sh = ss.getSheets()[0].setName(SHEET_NAME);
  sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
  sh.setFrozenRows(1);
  sh.setColumnWidth(5, 420);
  props.setProperty('SHEET_ID', ss.getId());
  Logger.log('Sheet: ' + ss.getUrl());
}

function sheet_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName(SHEET_NAME); }
function clean_(v, n) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n);
  if (/^[=+\-@]/.test(s)) s = "'" + s;          // keep spreadsheet formulas out
  return s;
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (p.website) return json_({ ok: true });     // honeypot field filled by bots
    const lat = Number(p.lat), lon = Number(p.lon);
    const story = clean_(p.story, 300);
    if (!isFinite(lat) || !isFinite(lon) || Math.abs(lat) > 90 || Math.abs(lon) > 180 || story.length < 3) return json_({ ok: false });
    const status = ['standing', 'gone', 'unsure'].indexOf(p.status) >= 0 ? p.status : 'unsure';
    const lock = LockService.getScriptLock(); lock.waitLock(10000);
    const sh = sheet_();
    sh.appendRow([new Date(), Math.round(lat * 1e5) / 1e5, Math.round(lon * 1e5) / 1e5, clean_(p.tree, 80), story, status,
                  clean_(p.name, 40), clean_(p.country, 50), false]);
    sh.getRange(sh.getLastRow(), 9).insertCheckboxes();
    lock.releaseLock();
    try {
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'New tree pin: ' + (clean_(p.tree, 80) || 'a tree'),
        story + '\n\n' + [clean_(p.name, 40), clean_(p.country, 50)].filter(String).join(', ') +
        '\nhttps://www.google.com/maps?q=' + lat + ',' + lon + '\n\nTick "approved" to show it on the map:\n' + sh.getParent().getUrl());
    } catch (err) {}
    return json_({ ok: true });
  } catch (err) { return json_({ ok: false }); }
}

function doGet() {
  const cache = CacheService.getScriptCache(); const hit = cache.get('pins');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const rows = sheet_().getDataRange().getValues().slice(1).filter(r => r[8] === true);
  const out = JSON.stringify(rows.map(r => ({ lat: r[1], lon: r[2], tree: r[3], story: String(r[4]).replace(/^'/, ''), status: r[5], name: r[6], country: r[7] })));
  cache.put('pins', out, 60);
  return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
}
