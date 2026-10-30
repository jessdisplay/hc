// In-memory stand-ins for the Google services Code.gs uses, so it can be run and tested with node.
// They behave like Google's where the script depends on it (a yyyy-mm-dd string typed into a cell
// becomes that day; copies land in the folder given; replies sit in the thread) and record every
// email so a test can see exactly what went out.
'use strict';

const BRISBANE = 10 * 3600 * 1000;  // no daylight saving in Queensland
let nextId = 1;
const id = (p) => (p + String(nextId++).padStart(4, '0') + 'x'.repeat(28)).slice(0, 33);

function fakeDate(nowIso) {
  const Real = Date;
  class FakeDate extends Real {
    constructor(...a) { if (a.length) super(...a); else super(FakeDate.now()); }
    static now() { return Real.parse(FakeDate.at); }
  }
  FakeDate.at = nowIso;
  FakeDate.parse = Real.parse; FakeDate.UTC = Real.UTC;
  return FakeDate;
}

function formatDate(d, tz, fmt) {
  const shift = tz === 'Australia/Brisbane' ? BRISBANE : 0;
  const t = new Date(d.getTime() + shift);
  const M = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
             'October', 'November', 'December'];
  const p = (n) => String(n).padStart(2, '0');
  return fmt.replace(/yyyy|MMMM|MMM|MM|dd|d|HH|mm/g, (k) => ({
    yyyy: t.getUTCFullYear(), MMMM: M[t.getUTCMonth()], MMM: M[t.getUTCMonth()].slice(0, 3),
    MM: p(t.getUTCMonth() + 1), dd: p(t.getUTCDate()), d: t.getUTCDate(), HH: p(t.getUTCHours()),
    mm: p(t.getUTCMinutes())
  })[k]);
}

class Blob {
  constructor(bytes, type, name) { this.bytes = bytes; this.type = type; this.name = name; }
  setName(n) { this.name = n; return this; }
  getName() { return this.name; }
  copyBlob() { return new Blob(this.bytes, this.type, this.name); }
}

// ---------------------------------------------------------------- Drive
class Drive {
  constructor(clock) { this.items = new Map(); this.clock = clock; }
  folder(name, parent) {
    const f = { kind: 'folder', id: id('fo'), name, parent, drive: this };
    this.items.set(f.id, f);
    return new Folder(f);
  }
  file(name, mime, parent, extra) {
    const f = Object.assign({ kind: 'file', id: id('fi'), name, mime, parent, created: new this.clock(),
                              description: '', drive: this }, extra || {});
    this.items.set(f.id, f);
    return new File(f);
  }
  get(fid, kind) {
    const f = this.items.get(fid);
    if (!f || f.kind !== kind) throw new Error('No item with the given ID could be found: ' + fid);
    return f;
  }
}
const iter = (arr) => { let i = 0; return { hasNext: () => i < arr.length, next: () => arr[i++] }; };
class Folder {
  constructor(f) { this.f = f; }
  getId() { return this.f.id; } getName() { return this.f.name; }
  getUrl() { return 'https://drive.google.com/drive/folders/' + this.f.id; }
  createFolder(n) { return this.f.drive.folder(n, this.f.id); }
  createFile(blob) { return this.f.drive.file(blob.name, blob.type, this.f.id, { bytes: blob.bytes }); }
  getFilesByType(mime) {
    return iter([...this.f.drive.items.values()]
      .filter((x) => x.kind === 'file' && x.parent === this.f.id && x.mime === mime).map((x) => new File(x)));
  }
}
class File {
  constructor(f) { this.f = f; }
  getId() { return this.f.id; } getName() { return this.f.name; }
  getUrl() {
    return this.f.mime === 'application/vnd.google-apps.document'
      ? 'https://docs.google.com/document/d/' + this.f.id + '/edit' : 'https://drive.google.com/file/d/' + this.f.id + '/view';
  }
  getDateCreated() { return this.f.created; }
  getDescription() { return this.f.description; }
  setDescription(s) { this.f.description = s; return this; }
  getParents() { return iter([new Folder(this.f.drive.get(this.f.parent, 'folder'))]); }
  makeCopy(name, folder) {
    return this.f.drive.file(name, this.f.mime, folder.getId(), { from: this.f.id, fields: this.f.fields });
  }
  getAs(mime) { return new Blob('pdf of ' + this.f.name, mime, this.f.name + '.pdf'); }
}

// ---------------------------------------------------------------- Sheets
class Range {
  constructor(sheet, r, c, nr, nc) { Object.assign(this, { sheet, r, c, nr, nc }); }
  cells(fn) {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(fn(this.sheet.cell(this.r + i, this.c + j)));
      out.push(row);
    }
    return out;
  }
  getValues() { return this.cells((x) => (x.v === undefined ? '' : x.v)); }
  getValue() { return this.getValues()[0][0]; }
  getDisplayValues() {
    return this.cells((x) => (x.v instanceof Date ? formatDate(x.v, 'Australia/Brisbane', 'd MMM yyyy')
      : x.v === undefined ? '' : String(x.v)));
  }
  getFormulas() { return this.cells((x) => x.formula || ''); }
  getRichTextValues() { return this.cells((x) => ({ getLinkUrl: () => x.link || null })); }
  setValues(rows) {
    if (rows.length !== this.nr || rows.some((r) => r.length !== this.nc)) {
      throw new Error('The number of rows or columns in the data does not match the range.');
    }
    rows.forEach((row, i) => row.forEach((v, j) => this.sheet.put(this.r + i, this.c + j, v)));
    return this;
  }
  setValue(v) { this.sheet.put(this.r, this.c, v); return this; }
  insertCheckboxes() { this.cells((x) => { x.v = false; x.checkbox = true; }); return this; }
  getRow() { return this.r; } getNumRows() { return this.nr; }
}
for (const m of ['setNumberFormat', 'setFontSize', 'setFontWeight', 'setFontColor', 'setBackground',
                 'setWrap', 'setDataValidation']) {
  Range.prototype[m] = function () { return this; };
}

class Sheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.grid = new Map(); this.maxRows = 1000; }
  cell(r, c) {
    const k = r + ',' + c;
    if (!this.grid.has(k)) this.grid.set(k, {});
    return this.grid.get(k);
  }
  put(r, c, v) {
    const x = this.cell(r, c);
    // Sheets reads typed yyyy-mm-dd text as that day, at midnight in the sheet's time zone.
    x.v = typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v) ? new this.ss.clock(v + 'T00:00:00+10:00') : v;
  }
  getName() { return this.name; }
  setName(n) { this.ss.rename(this, n); return this; }
  getRange(r, c, nr, nc) {
    if (r < 1 || c < 1) throw new Error('Those rows are out of bounds.');
    return new Range(this, r, c, nr || 1, nc || 1);
  }
  getLastRow() {
    let last = 0;
    for (const [k, x] of this.grid) if (x.v !== undefined && x.v !== '' && x.v !== null) last = Math.max(last, +k.split(',')[0]);
    return last;
  }
  getLastColumn() {
    let last = 0;
    for (const [k, x] of this.grid) if (x.v !== undefined && x.v !== '') last = Math.max(last, +k.split(',')[1]);
    return last;
  }
  getMaxRows() { return this.maxRows; }
  appendRow(row) { const r = this.getLastRow() + 1; row.forEach((v, j) => this.put(r, j + 1, v)); return this; }
  copyTo(ss) {
    const s = ss.insertSheet('Copy of ' + this.name);
    for (const [k, x] of this.grid) s.grid.set(k, Object.assign({}, x));
    return s;
  }
  setFrozenRows() { return this; }
  getActiveRange() { return this.ss.selection; }
}

class Spreadsheet {
  constructor(clock) { this.clock = clock; this.sheets = []; this.toasts = []; this.tz = null; }
  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error('A sheet with the name "' + name + '" already exists.');
    const s = new Sheet(this, name);
    this.sheets.push(s);
    return s;
  }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  rename(sheet, n) {
    if (this.getSheetByName(n)) throw new Error('A sheet with the name "' + n + '" already exists.');
    sheet.name = n;
  }
  setSpreadsheetTimeZone(tz) { this.tz = tz; }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/SHEET/edit'; }
  toast(msg) { this.toasts.push(msg); }
  select(sheet, r, nr) { this.active = sheet; this.selection = sheet.getRange(r, 1, nr || 1, 1); }
}

// ---------------------------------------------------------------- Gmail and Mail
class Gmail {
  constructor(me, clock) { this.me = me; this.clock = clock; this.sent = []; this.threads = new Map(); this.labels = {}; this.drafts = []; }
  sendEmail(to, subject, body, opts) { this.sent.push({ to, subject, body, opts: opts || {} }); }
  createDraft(to, subject, body, opts) {
    const gm = this;
    return {
      send() {
        const t = new Thread(gm);
        gm.threads.set(t.id, t);
        const m = t.add(gm.me, to, subject, body, (opts && opts.attachments) || []);
        gm.sent.push({ to, subject, body, opts: opts || {}, thread: t.id });
        return m;
      }
    };
  }
  getThreadById(tid) { return this.threads.get(tid) || null; }
  getUserLabelByName(n) { return this.labels[n] || null; }
  createLabel(n) { return (this.labels[n] = { name: n }); }
}
class Thread {
  constructor(gm) { this.gm = gm; this.id = id('th'); this.messages = []; this.labels = []; }
  add(from, to, subject, body, attachments, at) {
    const t = this;
    const m = { from, to, subject, body, attachments, date: at || new this.gm.clock() };
    this.messages.push(m);
    return {
      getThread: () => t, getFrom: () => m.from, getDate: () => m.date,
      getAttachments: () => m.attachments
    };
  }
  getId() { return this.id; }
  addLabel(l) { this.labels.push(l.name); return this; }
  getMessages() {
    return this.messages.map((m) => ({ getFrom: () => m.from, getDate: () => m.date, getAttachments: () => m.attachments }));
  }
  createDraftReply(body) { this.gm.drafts.push({ thread: this.id, body }); }
}

// ---------------------------------------------------------------- the whole Google
function google(opts) {
  const me = opts.me || 'jesse@riseos.care';
  const clock = fakeDate(opts.now);
  const ss = new Spreadsheet(clock);
  const drive = new Drive(clock);
  const gmail = new Gmail('Jesse <' + me + '>', clock);
  const mail = [];
  const alerts = [];
  const triggers = [];
  const fetches = [];
  let fetchFails = null;  // a function(url) -> status code, to make Drive's API refuse
  const trigger = (handler) => {
    const b = { handler };
    const chain = new Proxy({}, { get: (_, k) => (k === 'create' ? () => { triggers.push(b); return b; } : (...a) => { b[k] = a; return chain; }) });
    return chain;
  };
  const globals = {
    Date: clock,
    SpreadsheetApp: {
      getActive: () => ss,
      getActiveSheet: () => ss.active,
      getUi: () => ({
        alert: (m) => alerts.push(m),
        showSidebar: (h) => alerts.push('sidebar:' + h.title),
        createMenu: () => { const m = { addItem: () => m, addSeparator: () => m, addToUi: () => m }; return m; }
      }),
      newDataValidation: () => { const b = { requireValueInList: () => b, build: () => ({}) }; return b; }
    },
    DriveApp: {
      createFolder: (n) => drive.folder(n, 'root'),
      getFolderById: (fid) => new Folder(drive.get(fid, 'folder')),
      getFileById: (fid) => new File(drive.get(fid, 'file'))
    },
    UrlFetchApp: {
      fetch(url, o) {
        fetches.push({ url, o });
        const code = fetchFails ? fetchFails(url) : 200;
        const fid = url.match(/files\/([\w-]+)/)[1];
        if (code !== 200) return { getResponseCode: () => code, getContentText: () => '{"error":"refused"}' };
        const src = drive.get(fid, 'file');
        if (/\/copy\?/.test(url)) {
          const body = JSON.parse(o.payload);
          const doc = drive.file(body.name, body.mimeType, body.parents[0], { from: fid });
          return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ id: doc.getId() }) };
        }
        return { getResponseCode: () => 200, getBlob: () => new Blob('docx of ' + src.name, 'docx', 'export') };
      }
    },
    ScriptApp: {
      getOAuthToken: () => 'token',
      getProjectTriggers: () => triggers.map((t) => ({ getHandlerFunction: () => t.handler, t })),
      deleteTrigger: (x) => triggers.splice(triggers.indexOf(x.t), 1),
      newTrigger: trigger,
      WeekDay: { MONDAY: 'MONDAY' }
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => me }) },
    Utilities: { formatDate },
    GmailApp: gmail,
    MailApp: { sendEmail: (m) => mail.push(m) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
    HtmlService: { createHtmlOutput: (h) => ({ html: h, setTitle(t) { this.title = t; return this; } }) },
    console
  };
  return { globals, ss, drive, gmail, mail, alerts, triggers, fetches, clock, Blob,
           failFetch: (fn) => { fetchFails = fn; } };
}

module.exports = { google, Blob };
