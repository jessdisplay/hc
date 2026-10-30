/**
 * Harmony Care forms :: send forms to people to sign, and see them come back
 *
 * Lives inside the Harmony Care checklist, opened as a Google Sheet. Adds a
 * 'Harmony forms' menu:
 *
 *   Set up                    first time only: People, Settings, Signing templates and Forms log tabs,
 *                             the two Drive folders, and the hourly, daily and Monday timers
 *   Add the selected people   a Drive folder each; a worker also gets their own staff file tab and a
 *                             line on Team at a glance
 *   Send forms to someone     pick a person and their forms; a test copy of every email comes to you first
 *   Send the selected rows    on 'Sent and returned': send tested rows, if the Send panel was closed
 *   Make signing templates    a Google Doc of every form, for placing eSignature fields once
 *   Check for returned forms  what the hourly timer does, now
 *   Mark reviewed today       stamps today's date as 'Last reviewed' for the document on the selected row
 *   Email me what is due      what the Monday email says, now
 *
 * Two ways a form goes out. A form whose signing template has its signature fields placed goes by
 * Google eSignature: the script makes the person's copy, you press Request signature in it (Google
 * has no way for a script to press it), and the signed PDF Google files beside the copy is picked up
 * by the hourly check. Every other form goes by email as a PDF and a Word file, and the signed copy
 * they reply with is filed the same way. Either way 'Sent and returned' gets the row, the person's
 * folder gets the file, and the people in Settings get an email when it is back.
 *
 * Every email goes out from the account that ran Set up, never from the person pressing the button.
 * Pressing Send puts it on a queue that a one-minute timer (running as that account) empties. That
 * way every reply lands in the one inbox the hourly check reads, whoever in Harmony sent the form.
 *
 * Install: Extensions > Apps Script, replace Code.gs with this file, save, reload the sheet, then
 * Harmony forms > Set up. Google asks once for permission to use Drive, Gmail and this sheet.
 */

var TZ = 'Australia/Brisbane';
var TABS = {
  people: 'People', settings: 'Settings', templates: 'Signing templates', log: 'Forms log',
  tracker: 'Sent and returned', staff: 'Staff file (copy per worker)', team: 'Team at a glance',
  checklist: 'Checklist', everything: 'Everything'
};
var ROLES = ['Participant', 'Worker', 'Practitioner or partner'];
var SEND_TABS = {
  'Participant': 'Send to participants', 'Worker': 'Send to workers',
  'Practitioner or partner': 'Send to practitioners'
};
var LABEL = 'Harmony forms';  // every email thread the script starts carries this Gmail label
var ESIGN = 'Google eSignature', EMAIL = 'Email';
var MARK = 'Harmony forms: ';  // on a file's description: the hourly check has dealt with it

// People: one row per person, typed in by Harmony.
var PEOPLE_HEAD = 4;
var PEOPLE = ['Name', 'Email', 'Role', 'Signs for them (guardian or representative)', 'Their email',
              'Start date', 'Active', 'Folder', 'Notes'];
var P = index_(PEOPLE);

// Sent and returned: the workbook's own columns A to K, then the seven this script keeps.
var TRACKER_HEAD = 5, TRACKER_FIRST = 6;
var TRACKER = ['Date sent', 'Sent to', 'Who they are', 'Email', 'Form', 'They need to', 'Returned on',
               'Approved by', 'Approved on', 'Status', 'Notes',
               'Copy sent (opens it)', 'How it went out', 'Email thread', 'Signed copy', 'Last chased',
               'Send step', 'Test copy to'];
// Send step: 'Test' (queued) > 'Test sent' > 'Send' (queued) > 'Sent', or 'Not sent: why'.
var T = index_(TRACKER);

var TEMPLATE_HEAD = 4;
var TEMPLATES = ['Form', 'Goes to', 'Signing template (open it to place the fields)',
                 'Signature fields placed', 'Original'];
var TP = index_(TEMPLATES);

var SETTINGS = [
  ['Send emails as', 'Harmony Care', 'The name people see the emails come from.'],
  ['Test copies go to', '', 'Where a test copy goes if the sender gives no address in the Send panel.'],
  ['Emails go out from', '', 'Set up fills this: the account that ran it. Every email goes from here.'],
  ['Tell these people when a form comes back', '', 'Comma separated. They also get the Monday email.'],
  ["Folder for each person's files", '', 'Set up makes it. Each person gets a folder inside it.'],
  ['Signing templates folder', '', 'Set up makes it. The signing templates live here.'],
  ['Chase after (days)', 14, 'A form not back after this many days gets a reminder drafted in Gmail.']
];

var TIME_BUDGET_MS = 4.5 * 60 * 1000;  // Apps Script stops a run at 6 minutes
var TIMERS = ['sendQueued', 'checkReturns', 'dailyChase', 'mondayDigest'];


/* ------------------------------------------------------------------ menu and set up */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Harmony forms')
    .addItem('Send forms to someone', 'openSender')
    .addItem('Send the selected rows', 'sendSelectedRows')
    .addItem('Add the selected people', 'addSelectedPeople')
    .addSeparator()
    .addItem('Check for returned forms', 'checkReturns')
    .addItem('Mark reviewed today', 'markReviewedToday')
    .addItem('Email me what is due', 'mondayDigest')
    .addSeparator()
    .addItem('Make signing templates', 'makeTemplates')
    .addItem('Set up', 'setup')
    .addToUi();
}

function setup() {
  var ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  var me = Session.getEffectiveUser().getEmail();

  var settings = ensureTab_(ss, TABS.settings, 'Settings',
    'What the forms menu uses. Change a value in column B; leave the labels in column A as they are.',
    3, ['Setting', 'Value', 'What it does']);
  var labels = settings.getRange(4, 1, Math.max(settings.getLastRow() - 3, 1), 1).getValues()
    .map(function (r) { return r[0]; });
  SETTINGS.forEach(function (row) {
    if (labels.indexOf(row[0]) === -1) settings.appendRow(row);
  });
  var before = String(setting_('Emails go out from') || '');
  putSetting_('Emails go out from', me);
  if (!setting_('Test copies go to')) putSetting_('Test copies go to', me);
  if (!setting_('Tell these people when a form comes back')) {
    putSetting_('Tell these people when a form comes back', me);
  }
  if (!folderId_(setting_("Folder for each person's files"))) {
    putSetting_("Folder for each person's files",
                DriveApp.createFolder("Harmony Care - each person's forms").getUrl());
  }
  if (!folderId_(setting_('Signing templates folder'))) {
    putSetting_('Signing templates folder', DriveApp.createFolder('Harmony Care - signing templates').getUrl());
  }
  var shared = shareFolders_(ss);

  var people = ensureTab_(ss, TABS.people, 'People',
    'Everyone Harmony sends forms to: staff, participants, practitioners. Type a name, email and role, ' +
    'then select the row and choose Harmony forms > Add the selected people. A participant who does not ' +
    'sign for themselves: put the guardian or representative and their email in the next two columns.',
    PEOPLE_HEAD, PEOPLE);
  var max = Math.max(people.getMaxRows(), 200);
  people.getRange(PEOPLE_HEAD + 1, P['Role'] + 1, max - PEOPLE_HEAD, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(ROLES, true).build());
  people.getRange(PEOPLE_HEAD + 1, P['Active'] + 1, max - PEOPLE_HEAD, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());

  ensureTab_(ss, TABS.templates, 'Signing templates',
    'A Google Doc of each form that goes out. Open one, choose Tools > eSignature, place a Signature field ' +
    '(and Name and Date) for each person who signs, then tick column D. From then on that form goes by ' +
    'Google eSignature. A form without a ticked template goes by email for them to sign and send back.',
    TEMPLATE_HEAD, TEMPLATES);
  ensureTab_(ss, TABS.log, 'Forms log', 'What the forms menu and its timers did, newest at the bottom.',
    3, ['When', 'What', 'Detail']);

  var tracker = tab_(TABS.tracker);
  var head = tracker.getRange(TRACKER_HEAD, 1, 1, TRACKER.length);
  var have = head.getValues()[0];
  for (var i = 0; i < 11; i++) {
    if (have[i] !== TRACKER[i]) throw new Error("'Sent and returned' has changed: column " + (i + 1) +
                                                " should be '" + TRACKER[i] + "'.");
  }
  head.setValues([TRACKER]);
  tracker.getRange(TRACKER_HEAD, 12, 1, TRACKER.length - 11)
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#38553E').setWrap(true);

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (TIMERS.indexOf(t.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendQueued').timeBased().everyMinutes(1).create();
  ScriptApp.newTrigger('checkReturns').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('dailyChase').timeBased().everyDays(1).atHour(8).inTimezone(TZ).create();
  ScriptApp.newTrigger('mondayDigest').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8)
    .inTimezone(TZ).create();

  var moved = before && before.toLowerCase() !== me.toLowerCase();
  log_('Set up', 'Tabs, folders and timers ready. Emails go out from ' + me +
       (moved ? '; ' + before + "'s timers stop at their next run." : '.'));
  ui_('Set up is done. Emails now go out from ' + me + '.' +
      (moved ? ' ' + before + ' no longer sends anything.' : '') +
      (shared.length ? '\n\nThe two forms folders are now shared with: ' + shared.join(', ') + '.' : '') +
      '\n\nNext: type people into the People tab, then Harmony forms > Make signing templates.');
}


/* ------------------------------------------------------------------ people */

/** The menu item: a folder for every selected person; a worker also gets a staff file tab. */
function addSelectedPeople() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== TABS.people) return ui_('Select the people on the People tab first.');
  var range = sheet.getActiveRange();
  var done = [], problems = [];
  for (var r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    if (r <= PEOPLE_HEAD) continue;
    try {
      var person = person_(r);
      if (!person.name) continue;
      setUpPerson_(person, true);
      done.push(person.name);
      var who = signers_(person);
      if (who.problems.length) problems.push('Row ' + r + ': ' + who.problems.join(' ') + ' Forms cannot go to them yet.');
    } catch (err) {
      problems.push('Row ' + r + ': ' + err.message);
      log_('Add person failed', 'Row ' + r + ': ' + err.message);
    }
  }
  ui_((done.length ? 'Ready: ' + done.join(', ') + '.' : 'Nobody new was added.') +
      (problems.length ? '\n\nNot done:\n' + problems.join('\n') : ''));
}

function person_(row) {
  var v = tab_(TABS.people).getRange(row, 1, 1, PEOPLE.length).getValues()[0];
  return {
    row: row,
    name: String(v[P['Name']] || '').trim(),
    email: String(v[P['Email']] || '').trim(),
    role: String(v[P['Role']] || '').trim(),
    guardian: String(v[P['Signs for them (guardian or representative)']] || '').trim(),
    guardianEmail: String(v[P['Their email']] || '').trim(),
    active: String(v[P['Active']] || 'Yes').trim() !== 'No',
    folder: folderId_(v[P['Folder']])
  };
}

/** Their folder, and for a worker their staff file tab. 'strict' (the menu item) also stops on a full
 * Team at a glance; sending a form only logs it, because a full team list is no reason not to send. */
function setUpPerson_(person, strict) {
  if (ROLES.indexOf(person.role) === -1) {
    throw new Error(person.name + ' needs a role: ' + ROLES.join(', ') + '.');
  }
  if (!person.folder) {
    var parent = DriveApp.getFolderById(requireFolder_("Folder for each person's files"));
    var folder = parent.createFolder(person.name + ' (' + person.role + ')');
    tab_(TABS.people).getRange(person.row, P['Folder'] + 1).setValue(folder.getUrl());
    person.folder = folder.getId();
    log_('Folder made', person.name + ': ' + folder.getUrl());
  }
  if (person.role === 'Worker') {
    try {
      staffTab_(person);
    } catch (err) {
      if (strict) throw err;
      log_('Staff file tab not made', person.name + ': ' + err.message);
    }
  }
  return person;
}

/** A copy of the staff file tab named for the worker, and their name on Team at a glance. */
function staffTab_(person) {
  var ss = SpreadsheetApp.getActive();
  var name = sheetName_(person.name);
  if (!ss.getSheetByName(name)) {
    var copy = tab_(TABS.staff).copyTo(ss).setName(name);
    var labels = copy.getRange(1, 1, 20, 1).getValues();
    var start = tab_(TABS.people).getRange(person.row, P['Start date'] + 1).getValue();
    var fill = { 'Worker name': person.name, 'Role': 'Worker', 'Start date': start };
    labels.forEach(function (l, i) {
      if (Object.prototype.hasOwnProperty.call(fill, l[0]) && fill[l[0]] !== '') {
        copy.getRange(i + 1, 2).setValue(fill[l[0]]);
      }
    });
    log_('Staff file tab made', name);
  }
  var team = tab_(TABS.team);
  var names = team.getRange(5, 1, 15, 1).getValues().map(function (r) { return String(r[0]).trim(); });
  if (names.indexOf(name) !== -1) return;
  var free = names.indexOf('');
  if (free === -1) {
    throw new Error('Team at a glance is full (15 workers). Add a row to it, then add ' + person.name + ' again.');
  }
  team.getRange(5 + free, 1).setValue(name);
}

/** Who signs for this person, and anything wrong with the addresses. */
function signers_(person) {
  var to = [], problems = [];
  [[person.email, person.name], [person.guardianEmail, person.guardian || 'their guardian']].forEach(function (p) {
    if (!p[0]) return;
    if (!isEmail_(p[0])) problems.push("'" + p[0] + "' for " + p[1] + ' is not an email address.');
    else to.push(p[0]);
  });
  if (!to.length && !problems.length) problems.push(person.name + ' has no email address on the People tab.');
  return { to: to, problems: problems };
}


/* ------------------------------------------------------------------ sending */

function openSender() {
  var html = HtmlService.createHtmlOutput(SIDEBAR).setTitle('Send forms');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** For the sidebar: everyone active on the People tab. */
function sidebarPeople() {
  var sheet = tab_(TABS.people);
  var last = sheet.getLastRow();
  if (last <= PEOPLE_HEAD) return [];
  var out = [];
  for (var r = PEOPLE_HEAD + 1; r <= last; r++) {
    var p = person_(r);
    if (p.name && p.active) out.push({ row: r, name: p.name, role: p.role });
  }
  return out;
}

/** For the sidebar: the forms for this person's role, grouped as on its 'Send to' tab. */
function formsFor(row) {
  var person = person_(row);
  var tabName = SEND_TABS[person.role];
  if (!tabName) throw new Error(person.name + ' needs a role on the People tab before forms can go out.');
  var templates = templates_();
  var groups = sendGroups_(tabName).map(function (g) {
    return {
      name: g.name,
      forms: g.forms.map(function (f) {
        var t = templates[f.id];
        return { id: f.id, title: f.title, need: f.need, how: t && t.ready ? ESIGN : EMAIL };
      })
    };
  });
  var who = signers_(person);
  return { person: person.name, role: person.role, to: who.to, problems: who.problems, groups: groups,
           testTo: testTo_(), from: setting_('Emails go out from') };
}

/**
 * Makes the person's copy of each chosen form, writes its row in 'Sent and returned' (not dated until
 * it really goes), and queues a test copy of every email to testTo. Nothing reaches the person here.
 */
function prepare(row, ids, testTo) {
  var person = setUpPerson_(person_(row), false);
  var who = signers_(person);
  if (who.problems.length) throw new Error(who.problems.join(' '));
  var forms = sendGroups_(SEND_TABS[person.role]).reduce(function (all, g) { return all.concat(g.forms); }, []);
  var templates = templates_();
  testTo = String(testTo || testTo_()).trim();
  if (!isEmail_(testTo)) throw new Error('Give an email address for the test copy.');
  PropertiesService.getUserProperties().setProperty('TEST_TO', testTo);  // remembered for this person
  var folder = DriveApp.getFolderById(person.folder);
  var stamp = Utilities.formatDate(new Date(), TZ, 'd MMM yyyy');
  var out = [];
  ids.forEach(function (id) {
    var form = forms.filter(function (f) { return f.id === id; })[0];
    if (!form) return out.push({ title: id, error: 'That form is not on the ' + person.role + ' list.' });
    var name = form.title + ' - ' + person.name + ' (' + stamp + ')';
    try {
      var t = templates[id];
      var item = { title: form.title, to: who.to, how: t && t.ready ? ESIGN : EMAIL };
      var copyId = item.how === ESIGN
        ? DriveApp.getFileById(t.id).makeCopy(name, folder).getId()
        : toDoc_(id, name, person.folder);
      item.copyUrl = DriveApp.getFileById(copyId).getUrl();
      item.trackerRow = trackRow_({
        to: person.name, who: person.role, email: who.to.join(', '), form: form.title, need: form.need,
        copy: item.copyUrl, how: item.how, step: item.how === EMAIL ? 'Test' : '', testTo: testTo
      });
      if (item.how === EMAIL) item.tested = testTo;
      out.push(item);
      log_('Prepared', name + ' (' + item.how + ')');
    } catch (err) {
      log_('Prepare failed', name + ': ' + err.message);
      out.push({ title: form.title, error: err.message });
    }
  });
  return out;
}

/**
 * Queues the emails prepare() tested; the one-minute timer sends them. A row goes once its test copy
 * has gone, and only once, so pressing twice sends once.
 */
function sendNow(rows) {
  var tracker = tab_(TABS.tracker);
  return rows.map(function (r) {
    var v = tracker.getRange(r, 1, 1, TRACKER.length).getValues()[0];
    var step = v[T['Send step']];
    if (v[T['How it went out']] !== EMAIL) return { row: r, error: 'Row ' + r + ' is not an email form.' };
    if (v[T['Date sent']] || step === 'Sent') return { row: r, queued: true, note: 'Already sent.' };
    if (step === 'Send') return { row: r, queued: true, note: 'Already on its way.' };
    if (step !== 'Test sent') {
      return { row: r, error: v[T['Form']] + ': its test copy has not gone yet. Press Send again in a minute.' };
    }
    tracker.getRange(r, T['Send step'] + 1).setValue('Send');
    return { row: r, queued: true };
  });
}

/** The menu item, for when the Send panel was closed: queues the selected rows of 'Sent and returned'. */
function sendSelectedRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== TABS.tracker) return ui_("Select the rows to send on 'Sent and returned' first.");
  var range = sheet.getActiveRange(), rows = [];
  for (var r = Math.max(range.getRow(), TRACKER_FIRST); r < range.getRow() + range.getNumRows(); r++) rows.push(r);
  var res = sendNow(rows);
  var bad = res.filter(function (x) { return x.error; });
  ui_((res.length - bad.length) + ' on the way; they go out within a minute.' +
      (bad.length ? '\n\nNot sent:\n' + bad.map(function (x) { return x.error; }).join('\n') : ''));
}

/**
 * Every minute, as the account that ran Set up: sends queued test copies, then queued emails. A failure
 * stays on its row as 'Not sent: why' and is told to Settings' people; nothing is retried blindly.
 */
function sendQueued(e) {
  if (!timerOwner_(e)) return;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var tracker = tab_(TABS.tracker);
    var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
    var rows = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues();
    rows.forEach(function (v, i) {
      var step = v[T['Send step']], r = TRACKER_FIRST + i;
      if (step !== 'Test' && step !== 'Send') return;
      var cell = tracker.getRange(r, T['Send step'] + 1);
      try {
        var person = byName_(v[T['Sent to']]);
        var mail = email_(person, { title: v[T['Form']], need: v[T['They need to']] },
                          fileId_(v[T['Copy sent (opens it)']]));
        var opts = { name: setting_('Send emails as'), attachments: mail.attachments };
        if (step === 'Test') {
          GmailApp.sendEmail(v[T['Test copy to']], '[Test copy for ' + person.name + '] ' + mail.subject,
                             mail.body, opts);
          cell.setValue('Test sent');
          return;
        }
        var who = signers_(person);
        if (who.problems.length) throw new Error(who.problems.join(' '));
        var thread = GmailApp.createDraft(who.to.join(','), mail.subject, mail.body, opts).send().getThread();
        thread.addLabel(GmailApp.getUserLabelByName(LABEL) || GmailApp.createLabel(LABEL));
        date_(tracker.getRange(r, T['Date sent'] + 1), new Date());
        tracker.getRange(r, T['Email thread'] + 1).setValue(thread.getId());
        cell.setValue('Sent');
        log_('Sent', v[T['Form']] + ' to ' + who.to.join(', '));
      } catch (err) {
        cell.setValue('Not sent: ' + err.message);
        log_('Send failed', 'Row ' + r + ': ' + err.message);
        tell_('Not sent: ' + v[T['Form']] + ' to ' + v[T['Sent to']],
              v[T['Form']] + ' to ' + v[T['Sent to']] + ' did not go out: ' + err.message + '\n\n' +
              "Fix it, then make it again from Send forms. Row " + r + " of 'Sent and returned'.\n" +
              SpreadsheetApp.getActive().getUrl());
      }
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * The timers run as one account only: the one in Settings' 'Emails go out from', which ran Set up
 * last. Any other account's timer, left over from an earlier Set up, removes itself instead of
 * sending from the wrong address or reading the wrong inbox. 'e' is the timer's event; a menu click
 * has none.
 */
function timerOwner_(e) {
  var owner = String(setting_('Emails go out from') || '').trim().toLowerCase();
  var me = Session.getEffectiveUser().getEmail().toLowerCase();
  if (!owner || owner === me) return true;
  if (e && e.triggerUid) {
    ScriptApp.getProjectTriggers().forEach(function (t) {
      if (TIMERS.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
    });
    log_('Timers stopped', me + ' no longer sends; ' + owner + ' does.');
  }
  return false;
}

/**
 * Everyone who can edit the sheet can send forms, so each needs the two folders the copies go in.
 * Run Set up again after adding someone to the sheet. Returns who was newly added.
 */
function shareFolders_(ss) {
  var editors = ss.getEditors().map(function (u) { return u.getEmail(); }).filter(isEmail_);
  var added = [];
  ["Folder for each person's files", 'Signing templates folder'].forEach(function (label) {
    var folder = DriveApp.getFolderById(requireFolder_(label));
    var has = folder.getEditors().map(function (u) { return u.getEmail().toLowerCase(); });
    var owner = folder.getOwner();  // none for a folder in a shared drive
    if (owner) has.push(owner.getEmail().toLowerCase());
    editors.forEach(function (email) {
      if (has.indexOf(email.toLowerCase()) !== -1) return;
      try {
        folder.addEditor(email);
        if (added.indexOf(email) === -1) added.push(email);
      } catch (err) {
        log_('Folder not shared', label + ' with ' + email + ': ' + err.message);
      }
    });
  });
  return added;
}

/** The test address to offer: the last one this person used, else their own, else Settings'. */
function testTo_() {
  return PropertiesService.getUserProperties().getProperty('TEST_TO') ||
         Session.getActiveUser().getEmail() || setting_('Test copies go to');
}

/** The sidebar's 'I have requested the signature' button: dates the eSignature row. */
function markRequested(row) {
  var cell = tab_(TABS.tracker).getRange(row, T['Date sent'] + 1);
  if (!cell.getValue()) date_(cell, new Date());
  return true;
}

function email_(person, form, copyId) {
  var doc = DriveApp.getFileById(copyId);
  var first = (person.guardian && !person.email ? person.guardian : person.name).split(' ')[0];
  var need = String(form.need || 'sign').replace(/\.$/, '');
  return {
    subject: 'Harmony Care: ' + form.title,
    body: 'Hi ' + first + ',\n\n' +
          'Attached is the ' + form.title + ' from Harmony Care, as a PDF and as a Word file.\n\n' +
          'What to do: ' + need + '.\n\n' +
          'When it is done, reply to this email with the signed copy attached. ' +
          'A clear photo of each signed page is fine.\n\n' +
          'Any questions, just reply here.\n\nThank you,\nHarmony Care',
    attachments: [doc.getAs('application/pdf').setName(form.title + '.pdf'),
                  docxOf_(copyId, form.title)]
  };
}


/* ------------------------------------------------------------------ coming back */

/** Hourly: files each signed copy that came back, marks its row, and tells Settings' people. */
function checkReturns(e) {
  if (!timerOwner_(e)) {
    if (!e) ui_('Returned forms are checked from ' + setting_('Emails go out from') +
                "'s account, where the replies arrive. It happens every hour by itself.");
    return;
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;  // the last run is still going
  try {
    var tracker = tab_(TABS.tracker);
    var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
    var rows = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues();
    var claimed = {};
    rows.forEach(function (v) { var id = fileId_(v[T['Signed copy']]); if (id) claimed[id] = true; });
    var me = Session.getEffectiveUser().getEmail().toLowerCase();
    var back = 0;
    // eSignature rows waiting in the same folder share that folder's new PDFs
    var waiting = {};
    rows.forEach(function (v, i) {
      if (!v[T['Date sent']] || v[T['Returned on']] || v[T['Signed copy']]) return;
      var r = TRACKER_FIRST + i;
      try {
        if (v[T['How it went out']] === EMAIL && v[T['Email thread']]) {
          back += fromEmail_(tracker, r, v, me, claimed) ? 1 : 0;
        } else if (v[T['How it went out']] === ESIGN) {
          var copy = DriveApp.getFileById(fileId_(v[T['Copy sent (opens it)']]));
          var folder = copy.getParents().next();
          (waiting[folder.getId()] = waiting[folder.getId()] || { folder: folder, rows: [] })
            .rows.push({ r: r, v: v, name: copy.getName() });
        }
      } catch (err) {
        log_('Check failed', 'Row ' + r + ': ' + err.message);
      }
    });
    Object.keys(waiting).forEach(function (id) {
      back += fromFolder_(tracker, waiting[id], claimed);
    });
    if (back) log_('Came back', back + ' form' + (back > 1 ? 's' : ''));
  } finally {
    lock.releaseLock();
  }
}

function fromEmail_(tracker, r, v, me, claimed) {
  var thread = GmailApp.getThreadById(v[T['Email thread']]);
  if (!thread) return false;
  var sent = v[T['Date sent']];
  var replies = thread.getMessages().filter(function (m) {
    return m.getFrom().toLowerCase().indexOf(me) === -1 && day_(m.getDate()) >= day_(sent);
  });
  if (!replies.length) return false;
  var files = [];
  replies.forEach(function (m) {
    m.getAttachments({ includeInlineImages: false }).forEach(function (a) { files.push(a); });
  });
  var notes = String(v[T['Notes']] || '');
  if (!files.length) {
    if (notes.indexOf('Replied without') === -1) {
      tracker.getRange(r, T['Notes'] + 1).setValue(
        (notes ? notes + ' ' : '') + 'Replied without the signed form attached (' + day_(new Date()) + ').');
      var said = replies.map(function (m) { return String(m.getPlainBody() || '').trim().slice(0, 600); })
        .filter(String).join('\n---\n');
      tell_('Reply without the form: ' + v[T['Form']] + ' from ' + v[T['Sent to']],
            v[T['Sent to']] + ' replied about ' + v[T['Form']] + ' but did not attach the signed form. ' +
            'They wrote:\n\n' + (said || '(nothing)') + '\n\n' +
            'Reply to them at ' + v[T['Email']] + ', or ask ' + setting_('Emails go out from') +
            ' to answer in the thread (Gmail label: ' + LABEL + ').');
    }
    return false;
  }
  var folder = DriveApp.getFileById(fileId_(v[T['Copy sent (opens it)']])).getParents().next();
  var saved = files.map(function (a, n) {
    var ext = (a.getName().match(/\.[A-Za-z0-9]+$/) || [''])[0];
    var name = v[T['Form']] + ' - ' + v[T['Sent to']] + ' - signed' + (files.length > 1 ? ' ' + (n + 1) : '') + ext;
    return folder.createFile(a.copyBlob().setName(name))
      .setDescription(MARK + 'signed copy from email, Sent and returned row ' + r);
  });
  saved.forEach(function (f) { claimed[f.getId()] = true; });
  returned_(tracker, r, v, saved[0].getUrl(), saved.length > 1 ? saved.length + ' files, all in their folder.' : '');
  return true;
}

/**
 * Google files a completed eSignature PDF beside the document it was requested from. A new PDF in
 * the folder whose name holds the copy's name is that copy's. A new PDF that matches no copy goes to
 * the only waiting row if there is one, and otherwise to a person to match by hand.
 */
function fromFolder_(tracker, group, claimed) {
  var since = group.rows.map(function (w) { return day_(w.v[T['Date sent']]); }).sort()[0];
  var fresh = [];
  var it = group.folder.getFilesByType('application/pdf');
  while (it.hasNext()) {
    var f = it.next();
    if (claimed[f.getId()] || (f.getDescription() || '').indexOf(MARK) === 0) continue;
    if (day_(f.getDateCreated()) >= since) fresh.push(f);
  }
  var n = 0;
  var open = group.rows.slice();
  fresh.forEach(function (f) {
    var match = open.filter(function (w) {
      return f.getName().toLowerCase().indexOf(w.name.toLowerCase()) !== -1;
    })[0];
    if (!match && open.length === 1) match = open[0];
    if (!match) {
      claimed[f.getId()] = true;
      f.setDescription(MARK + 'signed file not matched to a form; match it by hand');  // asked once, not hourly
      log_('Unmatched signed file', f.getName() + ' ' + f.getUrl());
      tell_('A signed file to match: ' + f.getName(),
            'A new signed PDF arrived in ' + group.folder.getName() + ' and more than one form there is ' +
            'waiting, so it was not matched.\n\nOpen it: ' + f.getUrl() + '\n\nType the date in ' +
            "'Returned on' on the right row of 'Sent and returned', and paste the link in 'Signed copy'.");
      return;
    }
    claimed[f.getId()] = true;
    open.splice(open.indexOf(match), 1);
    returned_(tracker, match.r, match.v, f.getUrl(), '');
    n++;
  });
  return n;
}

function returned_(tracker, r, v, url, note) {
  date_(tracker.getRange(r, T['Returned on'] + 1), new Date());
  tracker.getRange(r, T['Signed copy'] + 1).setValue(url);
  if (note) {
    var notes = String(v[T['Notes']] || '');
    tracker.getRange(r, T['Notes'] + 1).setValue((notes ? notes + ' ' : '') + note);
  }
  tell_('Signed and back: ' + v[T['Form']] + ' from ' + v[T['Sent to']],
        v[T['Sent to']] + ' has returned ' + v[T['Form']] + '.\n\nOpen it: ' + url + '\n\n' +
        "Check it, then type who approved it and when in 'Sent and returned' (row " + r + ').\n' +
        SpreadsheetApp.getActive().getUrl());
}


/* ------------------------------------------------------------------ chasing and dates */

/**
 * Daily: a form out longer than 'Chase after (days)' gets a reminder drafted in its Gmail thread (never
 * sent by the script), at most once a week, and one email lists them all.
 */
function dailyChase(e) {
  if (!timerOwner_(e)) return;
  var tracker = tab_(TABS.tracker);
  var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
  var rows = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues();
  var days = Number(setting_('Chase after (days)')) || 14;
  var today = day_(new Date());
  var lines = [];
  rows.forEach(function (v, i) {
    var sent = v[T['Date sent']], r = TRACKER_FIRST + i;
    if (!sent || v[T['Returned on']] || !v[T['Form']]) return;
    if (daysBetween_(day_(sent), today) < days) return;
    var chased = v[T['Last chased']];
    if (chased && daysBetween_(day_(chased), today) < 7) return;
    try {
      var how = v[T['How it went out']];
      if (how === EMAIL && v[T['Email thread']]) {
        var thread = GmailApp.getThreadById(v[T['Email thread']]);
        if (thread) {
          thread.createDraftReply('Hi, just a reminder about the ' + v[T['Form']] + ' we sent on ' +
            pretty_(sent) + '. Could you sign it and reply with the signed copy ' +
            'when you have a moment?\n\nThank you,\nHarmony Care');
        }
        lines.push('- ' + v[T['Form']] + ', ' + v[T['Sent to']] + ': a reminder is waiting in Gmail drafts.');
      } else {
        lines.push('- ' + v[T['Form']] + ', ' + v[T['Sent to']] + ': open the copy and send a reminder from ' +
                   'its eSignature panel. ' + v[T['Copy sent (opens it)']]);
      }
      date_(tracker.getRange(r, T['Last chased'] + 1), new Date());
    } catch (err) {
      log_('Chase failed', 'Row ' + r + ': ' + err.message);
    }
  });
  if (lines.length) {
    tell_(lines.length + ' form' + (lines.length > 1 ? 's' : '') + ' to chase',
          'Out for ' + days + ' days or more and not back yet:\n\n' + lines.join('\n') + '\n\n' +
          SpreadsheetApp.getActive().getUrl());
    log_('Chased', lines.length + ' forms');
  }
}

/** Monday 8 am, or from the menu: what is overdue or due soon on the Everything tab, and forms out. */
function mondayDigest(e) {
  if (e && !timerOwner_(e)) return;  // from the menu anyone may ask for it
  var sheet = tab_(TABS.everything);
  var head = 5;
  var values = sheet.getRange(head, 1, Math.max(sheet.getLastRow() - head + 1, 1), 12).getDisplayValues();
  var col = index_(values[0]);
  ['What (opens the file)', 'Kind', 'Next due', 'Status', 'Practice Standard'].forEach(function (h) {
    if (!(h in col)) throw new Error("The Everything tab has changed: no '" + h + "' column.");
  });
  var now = [], soon = [];
  values.slice(1).forEach(function (v) {
    if (!v[col['Kind']]) return;
    var s = v[col['Status']];
    var line = '- ' + v[col['What (opens the file)']] + ' (' + v[col['Practice Standard']] + ')' +
               (v[col['Next due']] ? ', due ' + v[col['Next due']] : '') + ': ' + s;
    if (/^(Overdue|Missing|Expired)/.test(s)) now.push(line);
    else if (/^(Due within|Renew within)/.test(s)) soon.push(line);
  });
  var tracker = tab_(TABS.tracker);
  var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
  var out = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues()
    .filter(function (v) { return v[T['Date sent']] && !v[T['Returned on']]; }).length;
  var approve = tracker.getRange(TRACKER_FIRST, 1, last - TRACKER_FIRST + 1, TRACKER.length).getValues()
    .filter(function (v) { return v[T['Returned on']] && !v[T['Approved on']]; }).length;
  tell_('Harmony Care this week: ' + now.length + ' overdue or missing, ' + soon.length + ' due soon',
        'Overdue or missing (' + now.length + '):\n' + (now.join('\n') || '- nothing') + '\n\n' +
        'Due in the next 30 days (' + soon.length + '):\n' + (soon.join('\n') || '- nothing') + '\n\n' +
        'Forms out and not back: ' + out + '. Back and waiting for approval: ' + approve + '.\n\n' +
        SpreadsheetApp.getActive().getUrl());
}

/** The menu item: today's date in 'Last reviewed' for the document on the selected row. */
function markReviewedToday() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveRange().getRow();
  var checklist = tab_(TABS.checklist);
  var head = checklist.getRange(5, 1, 1, checklist.getLastColumn()).getValues()[0];
  var lastCol = head.indexOf('Last reviewed (type the date)') + 1;
  if (!lastCol) throw new Error("The Checklist tab has changed: no 'Last reviewed (type the date)' column.");
  var no = sheet.getRange(row, 1).getValue();
  if ([TABS.checklist, TABS.everything].indexOf(sheet.getName()) === -1 || typeof no !== 'number') {
    return ui_('Select a document row (one with a number in column A) on the Everything or Checklist tab.');
  }
  var target = row;
  if (sheet.getName() === TABS.everything) {
    var nos = checklist.getRange(1, 1, checklist.getLastRow(), 1).getValues();
    target = 0;
    for (var i = 0; i < nos.length; i++) if (nos[i][0] === no) { target = i + 1; break; }
    if (!target) throw new Error('Document No. ' + no + ' is not on the Checklist tab.');
  }
  date_(checklist.getRange(target, lastCol), new Date());
  var title = checklist.getRange(target, 4).getValue();
  log_('Reviewed', 'No. ' + no + ', ' + title);
  SpreadsheetApp.getActive().toast('No. ' + no + ' marked reviewed today. Next due and Status update by themselves.',
                                   'Harmony forms');
}


/* ------------------------------------------------------------------ signing templates */

/** A Google Doc of every form on the three 'Send to' tabs, once each, listed on Signing templates. */
function makeTemplates() {
  var started = Date.now();
  var folderId = requireFolder_('Signing templates folder');
  var sheet = tab_(TABS.templates);
  var have = templates_();
  var forms = {};
  Object.keys(SEND_TABS).forEach(function (role) {
    sendGroups_(SEND_TABS[role]).forEach(function (g) {
      g.forms.forEach(function (f) {
        (forms[f.id] = forms[f.id] || { title: f.title, roles: [] }).roles.push(role);
      });
    });
  });
  var made = 0, left = 0;
  Object.keys(forms).forEach(function (id) {
    if (have[id]) return;
    if (Date.now() - started > TIME_BUDGET_MS) { left++; return; }
    var f = forms[id];
    try {
      var docId = toDoc_(id, f.title, folderId);
      var r = Math.max(sheet.getLastRow(), TEMPLATE_HEAD) + 1;
      sheet.getRange(r, 1, 1, TEMPLATES.length).setValues([[
        f.title, f.roles.join(', '), DriveApp.getFileById(docId).getUrl(), false,
        'https://drive.google.com/file/d/' + id + '/view']]);
      sheet.getRange(r, TP['Signature fields placed'] + 1).insertCheckboxes();
      made++;
    } catch (err) {
      log_('Template failed', f.title + ': ' + err.message);
    }
  });
  log_('Signing templates', made + ' made' + (left ? ', ' + left + ' still to make' : ''));
  ui_(made + ' signing template' + (made === 1 ? '' : 's') + ' made.' +
      (left ? ' ' + left + ' still to make: run Make signing templates again.' : '') +
      '\n\nOpen each one, place the eSignature fields, and tick its box on the Signing templates tab.');
}

/** {original file id: {id: template doc id, ready: fields placed}} */
function templates_() {
  var sheet = tab_(TABS.templates);
  var last = sheet.getLastRow();
  var out = {};
  if (last <= TEMPLATE_HEAD) return out;
  sheet.getRange(TEMPLATE_HEAD + 1, 1, last - TEMPLATE_HEAD, TEMPLATES.length).getValues().forEach(function (v) {
    var original = fileId_(v[TP['Original']]), doc = fileId_(v[TP['Signing template (open it to place the fields)']]);
    if (original && doc) out[original] = { id: doc, ready: v[TP['Signature fields placed']] === true };
  });
  return out;
}


/* ------------------------------------------------------------------ the workbook's own tabs */

/** [{name, forms: [{id, title, need}]}] from a 'Send to' tab, the way the tab groups them. */
function sendGroups_(name) {
  var sheet = tab_(name);
  var head = 5;
  var last = sheet.getLastRow();
  if (last <= head) return [];
  var range = sheet.getRange(head + 1, 1, last - head, 3);
  var values = range.getValues(), links = range.getRichTextValues(), formulas = range.getFormulas();
  var groups = [];
  values.forEach(function (v, i) {
    if (typeof v[0] !== 'number' && v[0] && !v[1]) return groups.push({ name: String(v[0]), forms: [] });
    if (typeof v[0] !== 'number' || !groups.length) return;
    var url = links[i][1] && links[i][1].getLinkUrl();
    if (!url) url = (String(formulas[i][1]).match(/HYPERLINK\("([^"]+)"/i) || [])[1];
    var id = fileId_(url);
    if (id) groups[groups.length - 1].forms.push({ id: id, title: String(v[1]), need: String(v[2] || '') });
  });
  return groups.filter(function (g) { return g.forms.length; });
}

/** Writes a new row in 'Sent and returned' (left undated) and returns its row number. */
function trackRow_(x) {
  var tracker = tab_(TABS.tracker);
  var last = Math.max(tracker.getLastRow(), TRACKER_FIRST);
  var forms = tracker.getRange(TRACKER_FIRST, T['Form'] + 1, last - TRACKER_FIRST + 1, 1).getValues();
  var r = TRACKER_FIRST + forms.length;
  for (var i = 0; i < forms.length; i++) if (!forms[i][0]) { r = TRACKER_FIRST + i; break; }
  tracker.getRange(r, T['Sent to'] + 1, 1, 5).setValues([[x.to, x.who, x.email, x.form, x.need]]);
  tracker.getRange(r, T['Copy sent (opens it)'] + 1, 1, 2).setValues([[x.copy, x.how]]);
  tracker.getRange(r, T['Send step'] + 1, 1, 2).setValues([[x.step, x.step ? x.testTo : '']]);
  return r;
}


/* ------------------------------------------------------------------ Drive */

/** Copies a Word or PDF form into a Google Doc. Drive's own API does the conversion. */
function toDoc_(fileId, name, folderId) {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId +
                              '/copy?supportsAllDrives=true&fields=id', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: JSON.stringify({ name: name, mimeType: 'application/vnd.google-apps.document', parents: [folderId] })
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('Google could not make a Google Doc of ' + name + ' (' + res.getResponseCode() + ': ' +
                    res.getContentText().slice(0, 200) + ')');
  }
  return JSON.parse(res.getContentText()).id;
}

/** The Google Doc as a Word file, for people who would rather fill it in on a computer. */
function docxOf_(docId, title) {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + docId + '/export?mimeType=' +
    encodeURIComponent('application/vnd.openxmlformats-officedocument.wordprocessingml.document'), {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('Google could not make a Word copy of ' + title + '.');
  return res.getBlob().setName(title + '.docx');
}

function fileId_(url) {
  var m = String(url || '').match(/\/d\/([\w-]{20,})|[?&]id=([\w-]{20,})/);
  return m ? (m[1] || m[2]) : '';
}

function folderId_(url) {
  var m = String(url || '').match(/\/folders\/([\w-]{20,})/);
  return m ? m[1] : '';
}

function requireFolder_(label) {
  var id = folderId_(setting_(label));
  if (!id) throw new Error("Settings: '" + label + "' needs a Drive folder link. Run Harmony forms > Set up.");
  return id;
}


/* ------------------------------------------------------------------ small things */

function tab_(name) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error("There is no '" + name + "' tab. Run Harmony forms > Set up, or check its name.");
  return sheet;
}

function ensureTab_(ss, name, title, intro, headRow, head) {
  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1).setValue(title).setFontSize(14).setFontWeight('bold').setFontColor('#38553E');
  sheet.getRange(2, 1).setValue(intro).setFontColor('#5F564C');
  sheet.getRange(headRow, 1, 1, head.length).setValues([head])
    .setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#38553E').setWrap(true);
  sheet.setFrozenRows(headRow);
  return sheet;
}

function setting_(label) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(TABS.settings);
  if (!sheet || sheet.getLastRow() < 4) return '';
  var rows = sheet.getRange(4, 1, sheet.getLastRow() - 3, 2).getValues();
  for (var i = 0; i < rows.length; i++) if (rows[i][0] === label) return rows[i][1];
  return '';
}

function putSetting_(label, value) {
  var sheet = tab_(TABS.settings);
  var rows = sheet.getRange(4, 1, Math.max(sheet.getLastRow() - 3, 1), 1).getValues();
  for (var i = 0; i < rows.length; i++) {
    if (rows[i][0] === label) return sheet.getRange(4 + i, 2).setValue(value);
  }
  throw new Error("Settings has no '" + label + "' row.");
}

/** Everyone on the Settings line hears about it; a bad address there is logged, never fatal. */
function tell_(subject, body) {
  var to = String(setting_('Tell these people when a form comes back') || '')
    .split(',').map(function (s) { return s.trim(); }).filter(isEmail_);
  if (!to.length) return log_('Not told', subject + ' (no address in Settings)');
  try {
    MailApp.sendEmail({ to: to.join(','), subject: subject, body: body, name: setting_('Send emails as') });
  } catch (err) {
    log_('Notice failed', subject + ': ' + err.message);
  }
}

function log_(what, detail) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(TABS.log);
  if (sheet) sheet.appendRow([Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'), what, detail]);
}

function ui_(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (err) { log_('Message', message); }  // timers have no screen
}

function byName_(name) {
  var sheet = tab_(TABS.people);
  var names = sheet.getRange(PEOPLE_HEAD + 1, 1, Math.max(sheet.getLastRow() - PEOPLE_HEAD, 1), 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === String(name).trim()) return person_(PEOPLE_HEAD + 1 + i);
  }
  throw new Error(name + ' is no longer on the People tab.');
}

/** Dates go in as yyyy-mm-dd text, which Sheets reads as that day whatever the script's time zone. */
function date_(range, d) {
  range.setValue(day_(d)).setNumberFormat('d mmm yyyy');
}

function day_(d) {
  return d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d || '').slice(0, 10);
}

function pretty_(d) {
  var key = day_(d);
  return Utilities.formatDate(new Date(key + 'T12:00:00Z'), 'UTC', 'd MMMM');
}

function daysBetween_(a, b) {
  return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 86400000);
}

function isEmail_(s) {
  return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(s || '').trim());
}

/** A tab name Team at a glance can find: no characters Sheets refuses, no straight apostrophe. */
function sheetName_(name) {
  return String(name).replace(/[\[\]\*\?:\/\\]/g, ' ').replace(/'/g, '’').replace(/\s+/g, ' ').trim()
    .slice(0, 90);
}

function index_(names) {
  var out = {};
  names.forEach(function (n, i) { out[n] = i; });
  return out;
}


/* ------------------------------------------------------------------ the sidebar */

var SIDEBAR = [
  '<!doctype html><html><head><base target="_top"><style>',
  'body{font:13px/1.45 Arial,sans-serif;color:#2E2A26;margin:12px}',
  'h2{font-size:15px;color:#38553E;margin:0 0 8px}h3{font-size:12px;color:#38553E;margin:14px 0 4px}',
  'select,button{font:inherit;width:100%;padding:6px;margin:4px 0}',
  'button{background:#38553E;color:#fff;border:0;border-radius:4px;cursor:pointer}',
  'button:disabled{background:#9EAE8F;cursor:default}',
  'label{display:flex;gap:6px;align-items:flex-start;margin:3px 0}',
  '.how{font-size:11px;color:#5F564C}.bad{color:#8A4B2A}.ok{color:#38553E}',
  '.item{border-top:1px solid #DFD5C4;padding:6px 0}a{color:#38553E}',
  '</style></head><body>',
  '<h2>Send forms</h2>',
  '<select id="person"><option value="">Choose a person</option></select>',
  '<div id="forms"></div>',
  '<label class="how" style="display:block">Test copy to<input id="testTo" type="email" style="width:100%;font:inherit;padding:5px"></label>',
  '<button id="prepare" disabled>Make their copies and send me a test</button>',
  '<div id="result"></div>',
  '<button id="send" disabled style="display:none">Send the emails for real</button>',
  '<script>',
  'var $=function(id){return document.getElementById(id)};',
  'function el(tag,text,cls){var e=document.createElement(tag);if(text)e.textContent=text;if(cls)e.className=cls;return e}',
  'function fail(err){$("result").textContent="";$("result").appendChild(el("p",String(err.message||err),"bad"));',
  '  $("prepare").disabled=false}',
  'google.script.run.withFailureHandler(fail).withSuccessHandler(function(people){',
  '  people.forEach(function(p){var o=el("option",p.name+" ("+(p.role||"no role")+")");o.value=p.row;$("person").appendChild(o)})',
  '}).sidebarPeople();',
  '$("person").onchange=function(){',
  '  $("forms").textContent="";$("result").textContent="";$("send").style.display="none";$("prepare").disabled=true;',
  '  if(!this.value)return;',
  '  google.script.run.withFailureHandler(fail).withSuccessHandler(function(d){',
  '    if(d.problems.length){$("forms").appendChild(el("p",d.problems.join(" "),"bad"));return}',
  '    $("forms").appendChild(el("p","Goes to "+d.to.join(", ")+", from "+d.from+". A test of every email comes to you first.","how"));',
  '    if(!$("testTo").value)$("testTo").value=d.testTo||"";',
  '    d.groups.forEach(function(g){$("forms").appendChild(el("h3",g.name));',
  '      g.forms.forEach(function(f){var l=el("label");var c=el("input");c.type="checkbox";c.value=f.id;',
  '        c.onchange=function(){$("prepare").disabled=!document.querySelector("#forms input:checked")};',
  '        var t=el("span");t.appendChild(el("span",f.title));t.appendChild(el("br"));',
  '        t.appendChild(el("span",(f.how==="Email"?"By email, they sign and reply":"By Google eSignature")+". "+f.need,"how"));',
  '        l.appendChild(c);l.appendChild(t);$("forms").appendChild(l)})})',
  '  }).formsFor(Number(this.value))};',
  '$("prepare").onclick=function(){',
  '  var ids=[].map.call(document.querySelectorAll("#forms input:checked"),function(c){return c.value});',
  '  this.disabled=true;$("result").textContent="Making copies...";',
  '  google.script.run.withFailureHandler(fail).withSuccessHandler(function(items){',
  '    $("result").textContent="";var rows=[];',
  '    items.forEach(function(it){var d=el("div",null,"item");d.appendChild(el("b",it.title));',
  '      if(it.error){d.appendChild(el("p",it.error,"bad"))}',
  '      else if(it.how==="Email"){d.appendChild(el("p","A test copy goes to "+it.tested+" within a minute. Check it, then press Send.","how"));rows.push(it.trackerRow)}',
  '      else{var a=el("a","Open their copy");a.href=it.copyUrl;a.target="_blank";d.appendChild(a);',
  '        d.appendChild(el("p","In it: Tools > eSignature > Request signature, to "+it.to.join(", ")+".","how"));',
  '        var b=el("button","I have requested the signature");b.onclick=function(){b.disabled=true;',
  '          google.script.run.withFailureHandler(function(e){b.disabled=false;fail(e)})',
  '            .withSuccessHandler(function(){b.textContent="Dated in Sent and returned"}).markRequested(it.trackerRow)};',
  '        d.appendChild(b)}',
  '      $("result").appendChild(d)});',
  '    if(rows.length){$("send").style.display="block";$("send").disabled=false;',
  '      $("send").onclick=function(){$("send").disabled=true;',
  '        google.script.run.withFailureHandler(fail).withSuccessHandler(function(res){',
  '          var bad=res.filter(function(x){return x.error});',
  '          if(bad.length)$("send").disabled=false;',
  '          $("result").appendChild(el("p",(res.length-bad.length)+" on the way; they go out within a minute."+(bad.length?" "+bad.map(function(x){return x.error}).join(" "):""),bad.length?"bad":"ok"))',
  '        }).sendNow(rows)}}',
  '  }).prepare(Number($("person").value),ids,$("testTo").value)};',
  '</script></body></html>'
].join('\n');
