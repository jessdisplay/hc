/* Harmony's compliance tool. The page holds no data: it signs in by emailed code, then reads and writes
   through the script behind the Sheet. Every value reaches the screen as text, never as markup. */
(function () {
  'use strict';
  var API = window.HC_TOOL_API;
  var main = document.getElementById('main'), nav = document.getElementById('nav');
  var state = { token: null, email: '', today: '', checklist: null, forms: null, notes: null, book: null, team: null, file: {}, bulk: null, filter: 'all', find: '', role: '', ticked: {}, working: null };
  var ROLE_WORD = { 'Participant': 'Participants', 'Worker': 'Workers', 'Practitioner or partner': 'Practitioners and partners' };

  function store(key, value) {
    try {
      if (value === undefined) return localStorage.getItem(key);
      if (value === null) localStorage.removeItem(key); else localStorage.setItem(key, value);
    } catch (e) { return null; }
    return null;
  }

  function h(tag, attrs) {
    var el = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      var v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === 'on') Object.keys(v).forEach(function (ev) { el.addEventListener(ev, v[ev]); });
      else if (k === 'class') el.className = v;
      else if (k in el && k !== 'list') el[k] = v;
      else el.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(el, arguments[i]);
    return el;
  }
  function add(el, child) {
    if (child === null || child === undefined || child === false) return;
    if (Array.isArray(child)) child.forEach(function (c) { add(el, c); });
    else el.appendChild(child.nodeType ? child : document.createTextNode(String(child)));
  }
  function page() {
    main.textContent = '';
    for (var i = 0; i < arguments.length; i++) add(main, arguments[i]);
  }
  function plural(n, word, many) { return n + ' ' + (n === 1 ? word : (many || word + 's')); }
  function day(iso, withYear) {
    if (!iso) return '';
    var d = new Date(iso + 'T12:00:00Z');
    return d.toLocaleDateString('en-AU', { day: 'numeric', month: 'short', year: withYear === false ? undefined : 'numeric', timeZone: 'UTC' });
  }
  function offFile(url, words) { return url ? h('a', { href: url, target: '_blank', rel: 'noopener noreferrer' }, words) : words; }

  function call(action, body) {
    var send = Object.assign({ action: action, token: state.token }, body || {});
    return fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify(send), redirect: 'follow' })
      .then(function (r) { if (!r.ok) throw new Error('net'); return r.json(); }, function () { throw new Error('net'); })
      .catch(function (e) {
        if (e.message === 'net' || e instanceof SyntaxError) throw new Error("We couldn't reach your list. Check your connection and try again.");
        throw e;
      })
      .then(function (res) {
        if (res.signedOut) { signedOut(); throw new Error(res.error); }
        if (!res.ok) throw new Error(res.error || 'That did not work. Try again in a moment.');
        return res;
      });
  }
  function took(res) {
    if (res.today) state.today = res.today;
    if (res.forms) state.forms = res.forms;
    if (res.checklist) state.checklist = res.checklist;
    if (res.notes) state.notes = res.notes;
    return res;
  }
  /** Runs a change, shows what went wrong in place, and redraws with what came back. */
  function act(button, action, body, done) {
    var word = button && button.textContent;
    if (button) { button.disabled = true; button.textContent = 'Saving…'; }
    return call(action, body).then(took).then(function (res) {
      if (done) done(res);
      show();
    }).catch(function (e) {
      if (button) { button.disabled = false; button.textContent = word; }
      if (state.token) alertLine(e.message);
    });
  }
  function alertLine(text) {
    var old = document.getElementById('problem');
    if (old) old.remove();
    var line = h('p', { id: 'problem', class: 'note bad', role: 'alert' }, text);
    main.insertBefore(line, main.firstChild);
    window.scrollTo(0, 0);
  }

  /* ---- signing in */

  function signedOut() {
    state.token = null;
    store('hc-tool-token', null);
    state.forms = state.checklist = null;
    drawNav();
    signIn();
  }
  function signIn(asked, email, problem) {
    drawNav();
    var input = h('input', { type: asked ? 'text' : 'email', required: true, value: asked ? '' : (email || store('hc-tool-email') || ''),
      autocomplete: asked ? 'one-time-code' : 'email', inputMode: asked ? 'numeric' : 'email' });
    var go = h('button', { class: 'btn', type: 'submit' }, asked ? 'Sign in' : 'Email me a code');
    var form = h('form', { class: 'form', on: { submit: function (e) {
      e.preventDefault();
      if (!API) return signIn(asked, email, "This isn't switched on yet.");
      go.disabled = true;
      if (!asked) {
        var address = input.value.trim();
        store('hc-tool-email', address);
        fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action: 'code', email: address }) })
          .then(function () { signIn(true, address); }, function () { signIn(false, address, "We couldn't reach your list. Check your connection and try again."); });
        return;
      }
      call('signin', { email: email, code: input.value }).then(function (res) {
        state.token = res.token; state.email = res.email;
        store('hc-tool-token', res.token);
        start();
      }).catch(function (err) { signIn(true, email, err.message); });
    } } },
      h('label', null, asked ? 'The code we emailed to ' + email : 'Your email address', input), go);
    page(h('div', { class: 'signin' },
      h('h1', null, asked ? 'Check your email' : 'Sign in'),
      h('p', { class: 'say' }, asked ? 'It can take a minute to arrive, and works for 10 minutes. If nothing comes, this address may not be set up yet.' : 'We will email you a code. There is no password to remember.'),
      problem && h('p', { class: 'note bad', role: 'alert' }, problem), form,
      asked && h('p', { class: 'say' }, h('button', { class: 'link', type: 'button', on: { click: function () { signIn(false, email); } } }, 'Use a different address'))));
    input.focus();
  }

  /* ---- the frame */

  function drawNav() {
    [].slice.call(nav.querySelectorAll('.tabs,.me')).forEach(function (n) { n.remove(); });
    if (!state.token) return;
    var at = location.hash.slice(1).split('/')[0] || 'forms';
    if (at === 'person') at = 'forms';
    var tab = function (key, word) { return h('a', { href: '#' + key, 'aria-current': at === key ? 'page' : null }, word); };
    nav.appendChild(h('nav', { class: 'tabs', 'aria-label': 'Sections' }, tab('forms', 'Send and sign'), tab('documents', 'Documents'), tab('today', 'Today'), tab('checklist', 'Checklist'), tab('full', 'Full checklist'),
      tab('staff', 'Staff files'), tab('planner', 'Year planner'), tab('rules', 'Rules')));
    nav.appendChild(h('div', { class: 'me' }, state.email, h('button', { class: 'link', type: 'button', on: { click: function () {
      call('signout').catch(function () {}).then(signedOut);
    } } }, 'Sign out')));
  }
  function start() {
    drawNav();
    page(h('p', { class: 'say' }, 'Loading your list…'));
    call('load').then(took).then(show).catch(function (e) {
      if (state.token) page(h('p', { class: 'note bad', role: 'alert' }, e.message),
        h('p', null, h('button', { class: 'btn', type: 'button', on: { click: start } }, 'Try again')));
    });
  }
  function show() {
    if (!state.token) return;
    if (!state.forms || !state.checklist) return start();
    drawNav();
    var at = location.hash.slice(1).split('/');
    if (at[0] === 'checklist') checklistPage();
    else if (at[0] === 'full') fullPage();
    else if (at[0] === 'documents') documentsPage();
    else if (at[0] === 'today') todayPage();
    else if (at[0] === 'person') personPage(Number(at[1]));
    else if (at[0] === 'staff') staffPage(Number(at[1]));
    else if (at[0] === 'planner') plannerPage();
    else if (at[0] === 'rules') rulesPage();
    else formsPage(at[1] || 'list');
  }

  /* ---- reading the data */

  function bucket(status) {
    if (/^(Overdue|Expired)/.test(status)) return 'overdue';
    if (status === 'Missing') return 'missing';
    if (/^(Due within|Renew within)/.test(status)) return 'soon';
    if (/^(In date|Done|Not due yet)/.test(status)) return 'ok';
    return 'none';
  }
  function statusClass(status) {
    var b = bucket(status);
    return b === 'overdue' || b === 'missing' ? 'bad' : b === 'soon' ? 'soon' : b === 'ok' ? 'good' : 'quiet';
  }
  function allItems() {
    var out = [];
    state.checklist.modules.forEach(function (m) { m.standards.forEach(function (s) { out = out.concat(s.items); }); });
    return out;
  }
  function counts(items, of) {
    var c = {};
    items.forEach(function (i) { var k = of(i); c[k] = (c[k] || 0) + 1; });
    return c;
  }
  var STANDING = {
    ready: { word: 'Getting ready', cls: 'soon' }, failed: { word: 'Did not send', cls: 'bad' }, out: { word: 'With them', cls: 'quiet' },
    chase: { word: 'Chase up', cls: 'bad' }, read: { word: 'Sent, nothing to return', cls: 'good' },
    back: { word: 'Back, to check', cls: 'soon' }, done: { word: 'Done', cls: 'good' }
  };
  function sendsOf(name) { return state.forms.sends.filter(function (s) { return s.to === name; }); }
  /** The newest send of a form to a person: a form can go twice, and the latest is where it stands. */
  function latest(sends, title) {
    var mine = sends.filter(function (s) { return s.form === title; });
    return mine.length ? mine[mine.length - 1] : null;
  }

  /* ---- Today */

  function todayPage() {
    var items = allItems(), c = counts(items, function (i) { return bucket(i.status); });
    var f = counts(state.forms.sends, function (s) { return s.standing; });
    var line = function (n, words, href, note) {
      return h('p', { class: 'line' }, h('b', { class: n ? 'bad' : 'quiet' }, n), h('a', { href: href }, words), note && h('span', { class: 'say' }, note));
    };
    var when = new Date(state.today + 'T12:00:00Z').toLocaleDateString('en-AU', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC' });
    var next = items.filter(function (i) { return i.due && i.due >= state.today; }).sort(function (a, b) { return a.due < b.due ? -1 : 1; }).slice(0, 5);
    var people = state.forms.people.filter(function (p) { return p.active; }).length;
    var go = function (filter) { return '#checklist/' + filter; };
    page(h('h1', null, 'Today'), h('p', { class: 'lede' }, when),
      h('h2', null, 'Send and sign'),
      line(listTotal(), 'documents and forms go out to people', '#forms'),
      line(f.out || 0, 'with people now', '#forms/people'),
      line(f.chase || 0, 'to chase', '#forms/people'),
      line(f.back || 0, 'signed and back, to check', '#forms/people'),
      f.failed ? line(f.failed, 'that did not send', '#forms/people') : null,
      backSection(),
      h('h2', null, 'Checklist'),
      line(c.overdue || 0, 'overdue', go('overdue'), (c.overdue || 0) > 20 ? 'Most are one review round: documents approved in October 2023' : null),
      line(c.missing || 0, 'missing documents', go('missing')),
      line(c.soon || 0, 'due in the next 30 days', go('soon')),
      h('h2', null, 'Coming up'),
      next.length ? next.map(function (i) {
        return h('p', { class: 'line' }, h('span', null, offFile(i.url, i.title)), h('span', { class: 'say' }, day(i.due)));
      }) : h('p', { class: 'say' }, 'Nothing with a date ahead.'),
      h('h2', null, 'People'),
      people ? h('p', { class: 'say' }, plural(people, 'person', 'people') + ' on your list. ', h('a', { href: '#forms/people' }, 'See who has what'))
        : h('p', { class: 'say' }, 'Nobody is on your list yet. ', h('a', { href: '#forms/people' }, 'Add your first person'), ' to start sending.'),
      state.notes && state.notes.attention.length ? [h('h2', null, 'Worth knowing'), state.notes.attention.map(function (n) {
        return h('details', { class: 'area' }, h('summary', null, h('span', null, n.title)), h('div', { class: 'body' }, h('p', { class: 'say outcome' }, n.text)));
      })] : null,
      state.notes && state.notes.files.length ? [h('h2', null, 'The audited files'), state.notes.files.map(function (n) {
        return h('p', { class: 'line' }, h('span', null, offFile(n.url, n.title)), h('span', { class: 'say' }, n.text));
      })] : null);
  }

  /** What has come back and is waiting on someone here, named, so nobody has to go looking. */
  function backSection() {
    var back = state.forms.sends.filter(function (s) { return s.standing === 'back'; });
    if (!back.length) return null;
    return [h('h2', null, 'Signed and back'), back.map(function (s) {
      var p = state.forms.people.filter(function (x) { return x.name === s.to; })[0];
      return h('p', { class: 'line' }, h('span', null, p ? h('a', { href: '#person/' + p.row }, s.form + ', from ' + s.to) : s.form + ', from ' + s.to),
        h('span', { class: 'say' }, 'Back ' + day(s.back)));
    })];
  }

  /* ---- Checklist */

  function checklistPage() {
    var filter = location.hash.split('/')[1] || 'all';
    var word = state.find.trim().toLowerCase();
    var keep = function (i) {
      if (filter !== 'all' && bucket(i.status) !== filter) return false;
      return !word || (i.title + ' ' + i.who + ' ' + i.when).toLowerCase().indexOf(word) !== -1;
    };
    var items = allItems(), c = counts(items, function (i) { return bucket(i.status); });
    var stat = function (key, label, n, bad) {
      return h('a', { class: 'stat' + (bad && n ? ' bad' : ''), href: '#checklist/' + key, 'aria-pressed': filter === key ? 'true' : 'false' }, h('b', null, n), h('span', null, label));
    };
    var search = h('input', { type: 'search', placeholder: 'Find a document', 'aria-label': 'Find a document', value: state.find,
      on: { input: function () { state.find = search.value; drawList(); } } });
    var list = h('div');
    function drawList() {
      list.textContent = '';
      var word2 = state.find.trim().toLowerCase();
      word = word2;
      var shown = 0;
      state.checklist.modules.forEach(function (m) {
        var blocks = m.standards.map(function (s) {
          var rows = s.items.filter(keep);
          if (filter !== 'all' || word) { if (!rows.length) return null; } else if (!s.items.length) return null;
          shown += rows.length;
          var cc = counts(s.items, function (i) { return bucket(i.status); });
          var say = [cc.overdue && cc.overdue + ' overdue', cc.missing && cc.missing + ' missing', cc.soon && cc.soon + ' due soon'].filter(Boolean).join(', ');
          return h('details', { class: 'area', open: filter !== 'all' || !!word },
            h('summary', null, h('span', null, s.name), h('span', { class: 'say' + (cc.overdue || cc.missing ? ' bad' : '') }, say || plural(s.items.length, 'item'))),
            h('div', { class: 'body' },
              s.outcome && h('p', { class: 'say outcome' }, s.outcome.split('\n')[0]),
              s.indicators.length ? h('details', { class: 'more' }, h('summary', null, plural(s.indicators.length, 'indicator') + ' to be demonstrated'), h('ul', null, s.indicators.map(function (t) { return h('li', null, t); }))) : null,
              itemTable(rows),
              (s.also || []).map(function (a) { return h('p', { class: 'say' }, 'Also meets this standard: ', offFile(a.url, a.title)); })));
        }).filter(Boolean);
        if (blocks.length) { list.appendChild(h('h2', null, m.name)); blocks.forEach(function (b) { list.appendChild(b); }); }
      });
      if (!shown && (filter !== 'all' || word)) list.appendChild(h('p', { class: 'say' }, 'Nothing matches that.'));
    }
    page(h('h1', null, 'Checklist'),
      h('p', { class: 'lede' }, 'Every area, by Practice Standard: what you must hold, when it is needed and whether it is in date.'),
      h('div', { class: 'stats' }, stat('all', 'everything', items.length), stat('overdue', 'overdue', c.overdue || 0, true), stat('missing', 'missing', c.missing || 0, true),
        stat('soon', 'due soon', c.soon || 0), stat('ok', 'in date', c.ok || 0), stat('none', 'to record', c.none || 0)),
      h('div', { class: 'tools' }, search), list);
    drawList();
  }
  function itemTable(rows) {
    return h('table', { class: 'stack' },
      h('thead', null, h('tr', null, ['What', 'Who it is for', 'When it is needed', 'How often', 'Next due', 'Status'].map(function (t) { return h('th', null, t); }))),
      h('tbody', null, rows.map(function (i) {
        var who = /participant/i.test(i.who) ? 'who-participant' : /organisation/i.test(i.who) ? 'who-organisation' : /worker|key personnel|drivers/i.test(i.who) && i.details ? 'who-worker' : null;
        return h('tr', null,
          h('td', null, offFile(i.url, (i.no ? i.no + '. ' : '') + i.title), i.details && h('span', { class: 'sub' }, i.details), i.record ? recordMore(i.record) : null),
          h('td', { class: who, 'data-h': 'For' }, i.who), h('td', { 'data-h': 'When' }, i.when), h('td', { 'data-h': 'How often' }, i.often),
          h('td', { 'data-h': 'Next due' }, day(i.due)),
          h('td', { 'data-h': 'Status' }, h('span', { class: statusClass(i.status) }, i.status),
            i.no && i.status !== 'Missing' ? h('span', { class: 'doit' }, h('button', { class: 'link', type: 'button', on: { click: function (e) {
              if (window.confirm('Mark No. ' + i.no + ', ' + i.title + ', as reviewed today?')) act(e.target, 'reviewed', { no: i.no });
            } } }, 'Reviewed today')) : null));
      })));
  }

  function recordMore(r) {
    var line = function (label, value) { return value ? h('li', null, label + ': ' + value) : null; };
    return h('details', { class: 'more' }, h('summary', null, 'More about this document'), h('ul', null,
      line('Version', r.version), line('Approved by', r.approvedBy), line('Approved', r.approved), line('Last reviewed', r.reviewed ? day(r.reviewed) || r.reviewed : ''),
      line('Completed by', r.completedBy), line('Signed by', r.signedBy), line('Kept', r.kept), line('For the next audit', r.notes), line("Harmony's rule", r.rule)));
  }

  /* ---- The long version */

  /** Every Practice Standard on one page, in the same five columns as the 2023 checklist. */
  function fullPage() {
    var line = function (i) { return h('p', null, offFile(i.url, (i.no ? i.no + '. ' : '') + i.title)); };
    var rows = [];
    state.checklist.modules.forEach(function (m) {
      rows.push(h('tr', { class: 'band' }, h('td', { colSpan: 5 }, m.name)));
      m.standards.forEach(function (s) {
        rows.push(h('tr', null,
          h('td', { class: 'stdname', 'data-h': 'Practice Standard' }, s.name),
          h('td', { class: 'out', 'data-h': 'Outcome' }, s.outcome),
          h('td', { class: 'ind', 'data-h': 'Indicators to be demonstrated' }, s.indicators.map(function (t) { return h('p', null, t); })),
          h('td', { 'data-h': 'Policies and other documents' }, s.items.filter(function (i) { return !i.evidence; }).map(line),
            (s.also || []).map(function (a) { return h('p', null, 'Also meets this standard: ', offFile(a.url, a.title)); })),
          h('td', { 'data-h': 'Evidence/Information to be provided by Harmony Care' }, s.items.filter(function (i) { return i.evidence; }).map(line))));
      });
    });
    page(h('p', { class: 'row no-print' }, h('button', { class: 'btn ghost', type: 'button', on: { click: function () { window.print(); } } }, 'Print or save as a PDF')),
      h('div', { class: 'maphead' }, 'Harmony Care Support Services, Compliance Checklist'),
      h('table', { class: 'map' }, h('thead', null, h('tr', null, ['Practice Standard', 'Outcome', 'Indicators to be demonstrated', 'Policies and other documents', 'Evidence/Information to be provided by Harmony Care'].map(function (t) { return h('th', null, t); }))),
        h('tbody', null, rows)));
  }

  /* ---- Documents */

  var SHELVES = [
    ['Policies and procedures', ['Policy and procedure', 'Policy', 'Statement']],
    ['Forms and agreements', ['Form', 'Agreement', 'Checklist', 'Template']],
    ['Registers', ['Register']], ['Plans', ['Plan']], ['Records and charts', ['Record', 'Chart']],
    ['Position descriptions', ['Position description']], ['Guides and Easy Read', ['Guidance', 'Easy Read']],
    ['Business papers', ['Business paper']]
  ];
  function documentsPage() {
    var all = allItems().filter(function (i) { return i.record || i.status === 'Missing'; });
    var typeOf = function (i) { return i.record ? i.record.type : 'Business paper'; };
    var filter = location.hash.split('/')[1] || 'all', list = h('div');
    var c = counts(all, function (i) { return bucket(i.status); });
    var stat = function (key, label, n, bad) {
      return h('a', { class: 'stat' + (bad && n ? ' bad' : ''), href: '#documents/' + key, 'aria-pressed': filter === key ? 'true' : 'false' }, h('b', null, n), h('span', null, label));
    };
    var search = h('input', { type: 'search', placeholder: 'Find a document', 'aria-label': 'Find a document', value: state.find, on: { input: function () { state.find = search.value; draw(); } } });
    function draw() {
      list.textContent = '';
      var word = state.find.trim().toLowerCase(), shown = 0;
      SHELVES.forEach(function (shelf) {
        var rows = all.filter(function (i) {
          return shelf[1].indexOf(typeOf(i)) !== -1 && (filter === 'all' || bucket(i.status) === filter) && (!word || i.title.toLowerCase().indexOf(word) !== -1);
        }).sort(function (a, b) { return a.title < b.title ? -1 : 1; });
        if (!rows.length) return;
        shown += rows.length;
        list.appendChild(h('h2', null, shelf[0] + ': ' + rows.length));
        list.appendChild(h('table', { class: 'stack shelf' }, h('thead', null, h('tr', null, ['Document', 'Version', 'Approved', 'Next review', 'Status'].map(function (t) { return h('th', null, t); }))),
          h('tbody', null, rows.map(function (i) {
            var r = i.record || {};
            return h('tr', null, h('td', null, offFile(i.url, i.title), r.notes ? h('details', { class: 'more' }, h('summary', null, 'For the next audit'), h('p', { class: 'say' }, r.notes)) : null),
              h('td', { 'data-h': 'Version' }, r.version), h('td', { 'data-h': 'Approved' }, [r.approved, r.approvedBy].filter(Boolean).join(', ')), h('td', { 'data-h': 'Next review' }, day(i.due)),
              h('td', { 'data-h': 'Status' }, h('span', { class: statusClass(i.status) }, i.status), i.no && i.status !== 'Missing' ? h('span', { class: 'doit' }, h('button', { class: 'link', type: 'button', on: { click: function (e) {
                if (window.confirm('Mark ' + i.title + ' as reviewed today?')) act(e.target, 'reviewed', { no: i.no });
              } } }, 'Reviewed today')) : null));
          }))));
      });
      if (!shown) list.appendChild(h('p', { class: 'say' }, 'Nothing matches that.'));
    }
    page(h('h1', null, 'Documents'), h('p', { class: 'lede' }, 'Every policy, procedure, form, register and plan from the audited set. Each opens in Google Docs. Procedures sit inside their policy.'),
      h('div', { class: 'stats' }, stat('all', 'documents', all.length), stat('overdue', 'due for review', c.overdue || 0, true), stat('missing', 'missing', c.missing || 0, true), stat('soon', 'due soon', c.soon || 0), stat('ok', 'in date', c.ok || 0), stat('none', 'no date', c.none || 0)),
      h('div', { class: 'tools' }, search), list);
    draw();
  }

  /* ---- Send and sign */

  function tally(sends) {
    var c = counts(sends, function (s) { return s.standing; });
    return { out: (c.out || 0), chase: c.chase || 0, back: c.back || 0, done: (c.done || 0) + (c.read || 0), failed: c.failed || 0, ready: c.ready || 0 };
  }
  function listTotal() {
    return Object.keys(ROLE_WORD).reduce(function (n, role) {
      return n + state.forms.lists[role].reduce(function (m, g) { return m + g.forms.length; }, 0);
    }, 0);
  }
  function formsPage(view) {
    var f = state.forms, all = tally(f.sends);
    var plain = function (n, word, bad) { return h('span', { class: 'stat' + (bad && n ? ' bad' : '') }, h('b', null, n), h('span', null, word)); };
    var pick = function (key, href, word) { return view === key ? h('b', null, word) : h('a', { href: href }, word); };
    page(h('h1', null, 'Send and sign'), h('p', { class: 'lede' }, 'Everything that goes out to people: what they do with it, when, and who signs. Then whether it has come back.'),
      h('div', { class: 'stats' }, plain(listTotal(), 'go out'), plain(all.out, 'with people'), plain(all.chase, 'to chase', true), plain(all.back, 'back to check'), plain(all.done, 'done'), all.failed ? plain(all.failed, 'did not send', true) : null),
      h('div', { class: 'tools' }, h('p', { class: 'say' }, pick('list', '#forms', 'What goes out'), '  |  ', pick('people', '#forms/people', 'Who has what'), '  |  ', pick('office', '#forms/office', 'Kept in the office')),
        f.completed ? h('p', { class: 'say' }, offFile(f.completed, 'Open every completed document')) : null),
      view === 'people' ? peopleList() : view === 'office' ? officeList() : whatGoesOut());
  }
  function officeList() {
    var wrap = h('div', null, h('p', { class: 'say' }, 'Loading…'));
    needBook(function (book) {
      wrap.textContent = '';
      wrap.appendChild(h('p', { class: 'say' }, 'Registers, records and plans Harmony fills in itself. Nobody outside needs to be sent these.'));
      book.office.forEach(function (g) {
        wrap.appendChild(h('h3', null, g.name));
        wrap.appendChild(h('table', { class: 'stack goes' }, h('thead', null, h('tr', null, ['Document', 'When it is needed', 'Who fills it in', 'Who signs'].map(function (t) { return h('th', null, t); }))),
          h('tbody', null, g.forms.map(function (x) { return h('tr', null, h('td', null, offFile(x.url, x.title)), h('td', { 'data-h': 'When' }, x.when), h('td', { 'data-h': 'Fills it in' }, x.fills), h('td', { 'data-h': 'Signs' }, x.signs)); }))));
      });
    }, wrap);
    return wrap;
  }
  function needBook(draw, wrap) {
    if (state.book) return draw(state.book);
    call('book').then(function (res) { state.book = res.data; draw(state.book); }).catch(function (e) { wrap.textContent = ''; wrap.appendChild(h('p', { class: 'note bad' }, e.message)); });
  }
  function peopleList() {
    var f = state.forms, wrap = h('div');
    var open = h('button', { class: 'btn ghost', type: 'button', on: { click: function () { wrap.insertBefore(personForm(), wrap.firstChild); open.remove(); } } }, 'Add a person');
    var many = h('button', { class: 'btn', type: 'button', on: { click: function () { wrap.insertBefore(bulkPanel(), wrap.firstChild); many.remove(); } } }, 'Add many people');
    var search = h('input', { type: 'search', placeholder: 'Find a person', 'aria-label': 'Find a person', value: state.find, on: { input: function () { state.find = search.value; draw(); } } });
    var pick = h('select', { 'aria-label': 'Who they are', on: { change: function () { state.role = pick.value; draw(); } } },
      h('option', { value: '' }, 'Everyone'), Object.keys(ROLE_WORD).map(function (r) { return h('option', { value: r, selected: state.role === r }, ROLE_WORD[r]); }));
    var table = h('div');
    function draw() {
      table.textContent = '';
      if (state.bulk) {
        table.appendChild(h('div', { class: 'note' }, h('p', null, plural(state.bulk.added, 'person', 'people') + ' added. Their folders appear within a few minutes.'),
          state.bulk.skipped.length ? h('ul', null, state.bulk.skipped.map(function (x) { return h('li', { class: 'bad' }, x.name + ': ' + x.why); })) : null));
      }
      var word = state.find.trim().toLowerCase();
      var rows = f.people.filter(function (p) { return (!state.role || p.role === state.role) && (!word || p.name.toLowerCase().indexOf(word) !== -1); });
      if (!f.people.length) return table.appendChild(h('p', { class: 'say' }, 'Add the people Harmony sends forms to: workers, participants and practitioners.'));
      if (!rows.length) return table.appendChild(h('p', { class: 'say' }, 'Nobody matches that.'));
      table.appendChild(h('table', { class: 'stack' },
        h('thead', null, h('tr', null, h('th', null, 'Person'), h('th', null, 'Who they are'), ['With them', 'To chase', 'Back to check', 'Done'].map(function (t) { return h('th', { class: 'num' }, t); }))),
        h('tbody', null, rows.map(function (p) {
          var t = tally(sendsOf(p.name));
          var n = function (v, label, bad) { return h('td', { class: 'num', 'data-h': label }, v ? h('span', { class: bad ? 'bad' : '' }, v) : h('span', { class: 'quiet' }, '0')); };
          return h('tr', null, h('td', null, h('a', { href: '#person/' + p.row }, p.name), !p.active && h('span', { class: 'sub' }, 'Not active'), p.problems.length ? h('span', { class: 'sub bad' }, p.problems.join(' ')) : null),
            h('td', { 'data-h': 'Who' }, p.role), n(t.out, 'With them'), n(t.chase, 'To chase', true), n(t.back, 'Back to check'), n(t.done, 'Done'));
        }))));
    }
    wrap.appendChild(h('div', { class: 'tools' }, many, open, search, pick));
    wrap.appendChild(table);
    draw();
    return wrap;
  }
  /** A list from Excel, a CSV or a paste, shown and checked here before anything is saved. */
  function bulkPanel() {
    var panel = h('section', { class: 'note' }), preview = h('div');
    var rows = [], fallback = h('select', { 'aria-label': 'If your list does not say who they are', on: { change: read } }, h('option', { value: '' }, 'Choose'), Object.keys(ROLE_WORD).map(function (r) { return h('option', { value: r }, ROLE_WORD[r]); }));
    function read() {
      preview.textContent = '';
      if (!rows.length) return;
      var known = state.forms.people.map(function (p) { return p.name; });
      var people = window.HCPeople.toPeople(rows, fallback.value, known);
      var good = people.filter(function (x) { return !x.problem; }), bad = people.length - good.length;
      var add = h('button', { class: 'btn', type: 'button', disabled: !good.length, on: { click: function () {
        act(add, 'people', { people: good.map(function (x) { return { name: x.name, email: x.email, role: x.role, guardian: x.guardian, guardianEmail: x.guardianEmail, start: x.start }; }) },
          function (res) { state.bulk = res; });
      } } }, 'Add ' + plural(good.length, 'person', 'people'));
      preview.appendChild(h('p', { class: 'row' }, add, bad ? h('span', { class: 'bad' }, plural(bad, 'line') + ' to fix. The rest can go in now.') : h('span', { class: 'good' }, 'All good.')));
      preview.appendChild(h('table', { class: 'stack' }, h('thead', null, h('tr', null, ['Name', 'Email', 'Who they are', 'Signs for them', 'Their email', 'Check'].map(function (t) { return h('th', null, t); }))),
        h('tbody', null, people.map(function (x) { return h('tr', null, h('td', null, x.name), h('td', { 'data-h': 'Email' }, x.email), h('td', { 'data-h': 'Who' }, x.role), h('td', { 'data-h': 'Signs' }, x.guardian), h('td', { 'data-h': 'Their email' }, x.guardianEmail),
          h('td', { 'data-h': 'Check' }, h('span', { class: x.problem ? 'bad' : 'good' }, x.problem || 'Ready'))); }))));
    }
    var paste = h('textarea', { rows: 4, 'aria-label': 'Paste from Excel', placeholder: 'Or copy the rows in Excel and paste them here', on: { input: function () { rows = window.HCPeople.parseText(paste.value); read(); } } });
    var problem = h('p', { class: 'bad', role: 'alert' });
    var file = h('input', { type: 'file', accept: '.xlsx,.csv,.txt', 'aria-label': 'Choose a file', on: { change: function () {
      var f = file.files[0];
      problem.textContent = '';
      if (!f) return;
      var xlsx = /\.xlsx$/i.test(f.name);
      var done = function (r) { rows = r; read(); };
      var fail = function () { problem.textContent = "We couldn't read that file. Save it as CSV, or copy the rows and paste them below."; };
      if (xlsx) f.arrayBuffer().then(window.HCPeople.readXlsx).then(done).catch(fail);
      else f.text().then(function (t) { done(window.HCPeople.parseText(t)); }).catch(fail);
    } } });
    panel.appendChild(h('h3', null, 'Add many people'));
    panel.appendChild(h('p', null, 'Upload an Excel or CSV file with a heading row. We look for their name, email, who they are, and who signs for them with that person\u2019s email. ', h('a', { href: 'people-template.csv', download: 'Harmony people list.csv' }, 'Download a blank list')));
    panel.appendChild(h('div', { class: 'tools' }, file));
    panel.appendChild(problem);
    panel.appendChild(h('div', { class: 'form wide' }, paste, h('label', null, 'If your list does not say who they are', fallback)));
    panel.appendChild(preview);
    return panel;
  }
  function personForm(p) {
    var field = function (label, input) { return h('label', null, label, input); };
    var name = h('input', { required: true, maxLength: 120, value: p ? p.name : '', readOnly: !!p });
    var email = h('input', { type: 'email', maxLength: 254, value: p ? p.email : '' });
    var role = h('select', { required: true }, h('option', { value: '' }, 'Choose'), Object.keys(ROLE_WORD).map(function (r) { return h('option', { value: r, selected: !!p && p.role === r }, r); }));
    var guardian = h('input', { maxLength: 120, value: p ? p.guardian : '' });
    var guardianEmail = h('input', { type: 'email', maxLength: 254, value: p ? p.guardianEmail : '' });
    var active = h('input', { type: 'checkbox', checked: !p || p.active });
    var save = h('button', { class: 'btn', type: 'submit' }, p ? 'Save' : 'Add them');
    return h('form', { class: 'form', on: { submit: function (e) {
      e.preventDefault();
      act(save, 'person', { row: p ? p.row : 0, name: name.value, email: email.value, role: role.value, guardian: guardian.value, guardianEmail: guardianEmail.value, active: active.checked },
        function (res) { if (!p) location.hash = '#person/' + res.person; });
    } } },
      field('Their name', name), field('Their email', email), field('Who they are', role),
      field('Signs for them (guardian or representative, if any)', guardian), field("That person's email", guardianEmail),
      p ? h('label', { class: 'row' }, active, 'Still with Harmony') : null, h('div', { class: 'row' }, save));
  }
  /** Every document and form that goes out, in the order it goes, whether or not anyone is on the list yet. */
  function whatGoesOut() {
    var wrap = h('div'), list = h('div');
    var search = h('input', { type: 'search', placeholder: 'Find a document or form', 'aria-label': 'Find a document or form', value: state.find, on: { input: function () { state.find = search.value; draw(); } } });
    var pick = h('select', { 'aria-label': 'Who it goes to', on: { change: function () { state.role = pick.value; draw(); } } },
      h('option', { value: '' }, 'Everyone'), Object.keys(ROLE_WORD).map(function (r) { return h('option', { value: r, selected: state.role === r }, ROLE_WORD[r]); }));
    function draw() {
      list.textContent = '';
      var word = state.find.trim().toLowerCase(), shown = 0;
      Object.keys(ROLE_WORD).forEach(function (role) {
        if (state.role && state.role !== role) return;
        var groups = state.forms.lists[role].map(function (g) {
          return { name: g.name, forms: g.forms.filter(function (x) { return !word || (x.title + ' ' + x.need).toLowerCase().indexOf(word) !== -1; }) };
        }).filter(function (g) { return g.forms.length; });
        if (!groups.length) return;
        var n = groups.reduce(function (m, g) { return m + g.forms.length; }, 0);
        shown += n;
        list.appendChild(h('h2', null, ROLE_WORD[role] + ': ' + n));
        groups.forEach(function (g) {
          list.appendChild(h('h3', null, g.name));
          list.appendChild(h('table', { class: 'stack goes' },
            h('thead', null, h('tr', null, ['Document', 'They need to', 'When it is needed', 'Who fills it in', 'Who signs', 'Where it stands'].map(function (t) { return h('th', null, t); }))),
            h('tbody', null, g.forms.map(function (form) { return goesRow(role, form); }))));
        });
      });
      if (!shown) list.appendChild(h('p', { class: 'say' }, 'Nothing matches that.'));
    }
    wrap.appendChild(h('div', { class: 'tools' }, search, pick));
    wrap.appendChild(list);
    draw();
    return wrap;
  }
  function goesRow(role, form) {
    var sends = state.forms.sends.filter(function (s) { return s.form === form.title; }), t = tally(sends);
    var say = sends.length ? [t.out + t.chase ? (t.out + t.chase) + ' with people' : '', t.chase ? t.chase + ' to chase' : '', t.back ? t.back + ' back' : '', t.done ? t.done + ' done' : ''].filter(Boolean).join(', ') : 'Not sent yet';
    var people = state.forms.people.filter(function (p) { return p.active && p.role === role; });
    var send = people.length ? h('select', { 'aria-label': 'Send ' + form.title + ' to', on: { change: function (e) {
      var row = Number(e.target.value);
      if (!row) return;
      (state.ticked[row] = state.ticked[row] || {})[form.id] = true;
      location.hash = '#person/' + row;
    } } }, h('option', { value: '' }, 'Send to'), people.map(function (p) { return h('option', { value: p.row }, p.name); }))
      : h('a', { href: '#forms/people' }, 'Add a person to send it');
    return h('tr', null, h('td', null, offFile(form.url, form.title)), h('td', { 'data-h': 'They need to' }, form.need), h('td', { 'data-h': 'When' }, form.when),
      h('td', { 'data-h': 'Fills it in' }, form.fills), h('td', { 'data-h': 'Signs' }, form.signs),
      h('td', { 'data-h': 'Where it stands' }, h('span', { class: t.chase ? 'bad' : 'quiet' }, say), h('span', { class: 'doit' }, send)));
  }

  /* ---- one person */

  function personPage(row) {
    var p = state.forms.people.filter(function (x) { return x.row === row; })[0];
    if (!p) return page(h('p', { class: 'say' }, 'That person is not on your list. ', h('a', { href: '#forms/people' }, 'Back to your list')));
    var sends = sendsOf(p.name), ticked = state.ticked[row] = state.ticked[row] || {};
    var bar = h('div', { class: 'tickbar' }), editing = h('div');
    var count = function () { return Object.keys(ticked).filter(function (k) { return ticked[k]; }); };
    function drawBar() {
      bar.textContent = '';
      var ids = count();
      bar.hidden = !ids.length;
      if (!ids.length) return;
      var email = h('button', { class: 'btn', type: 'button', disabled: !!p.problems.length, on: { click: function () {
        act(email, 'prepare', { person: row, ids: ids }, function (res) { state.ticked[row] = {}; state.working = { row: row, items: res.items }; });
      } } }, 'Send by email');
      var hand = h('button', { class: 'btn ghost', type: 'button', on: { click: function () {
        act(hand, 'handed', { person: row, ids: ids }, function () { state.ticked[row] = {}; });
      } } }, 'Handed over in person');
      bar.appendChild(h('b', null, plural(ids.length, 'ticked', 'ticked')));
      bar.appendChild(email); bar.appendChild(hand);
      if (p.problems.length) bar.appendChild(h('span', { class: 'bad' }, 'Fix their email first.'));
    }
    var boxes = [];
    var groups = state.forms.lists[p.role] || [];
    var content = groups.map(function (g) {
      var open = g.forms.filter(function (form) { var s = latest(sends, form.title); return !s || s.standing === 'failed'; });
      return h('section', null, h('div', { class: 'row' }, h('h3', null, g.name),
        open.length > 1 ? h('button', { class: 'link say', type: 'button', on: { click: function () {
          var all = open.every(function (x) { return ticked[x.id]; });
          open.forEach(function (x) { ticked[x.id] = !all; });
          boxes.forEach(function (b) { b.box.checked = !!ticked[b.id]; });
          drawBar();
        } } }, 'Tick all not sent') : null),
        h('table', { class: 'stack person' }, h('thead', null, h('tr', null, h('th', null, ''), h('th', null, 'Form'), h('th', null, 'When it is needed'), h('th', null, 'Where it stands'))),
          h('tbody', null, g.forms.map(function (form) { return formRow(form, latest(sends, form.title), ticked, boxes, drawBar); }))));
    });
    var w = state.working && state.working.row === row ? state.working.items : null;
    page(h('a', { class: 'back', href: '#forms/people' }, 'Who has what'), h('h1', null, p.name),
      h('p', { class: 'lede' }, [p.role, p.email, p.guardian ? 'Signs: ' + p.guardian + (p.guardianEmail ? ', ' + p.guardianEmail : '') : ''].filter(Boolean).join('  |  ')),
      p.problems.length ? h('p', { class: 'note bad' }, p.problems.join(' ')) : null,
      h('p', { class: 'row say' }, h('button', { class: 'link', type: 'button', on: { click: function () { editing.textContent = ''; editing.appendChild(personForm(p)); } } }, 'Change their details'),
        p.folder ? offFile(p.folder, 'Open their folder') : h('span', null, 'Their folder is being made'),
        p.role === 'Worker' ? h('a', { href: '#staff/' + p.row }, 'Their staff file') : null), editing,
      w ? testPanel(p, w) : null, content, bar);
    drawBar();
  }
  function formRow(form, send, ticked, boxes, drawBar) {
    var free = !send || send.standing === 'failed';
    var box = free ? h('input', { type: 'checkbox', checked: !!ticked[form.id], 'aria-label': 'Tick ' + form.title, on: { change: function () { ticked[form.id] = box.checked; drawBar(); } } }) : null;
    if (box) boxes.push({ id: form.id, box: box });
    var st = send ? STANDING[send.standing] : null;
    var acts = [];
    var button = function (word, action, undo) {
      return h('button', { class: 'link', type: 'button', on: { click: function (e) { act(e.target, action, { row: send.row, undo: undo }); } } }, word);
    };
    if (send) {
      if (send.standing === 'out' || send.standing === 'chase') acts.push(button('Mark as back', 'back'));
      if (send.standing === 'back') { acts.push(button('Approve', 'approve')); acts.push(button('Not back yet', 'back', true)); }
      if (send.standing === 'done') acts.push(button('Not approved yet', 'approve', true));
      if (send.signed) acts.push(offFile(send.signed, 'Open the signed copy'));
      else if (send.standing === 'out' || send.standing === 'chase' || send.standing === 'back') acts.push(addSigned(send));
      if (send.standing === 'ready' && send.how === 'Google eSignature') {
        acts.push(offFile(send.copy, 'Open their copy'));
        acts.push(button('I have requested the signature', 'requested'));
      }
    }
    var line = !send ? 'Not sent' : send.standing === 'failed' ? send.step : send.standing === 'ready' ? (send.step === 'Test sent' ? 'Test sent, ready to send' : send.step === 'Send' ? 'On its way' : send.step === 'Test' ? 'Test on its way' : 'Copy made') : st.word;
    var dates = !send ? '' : [send.sent && (send.how === 'By hand' ? 'Handed over ' : 'Sent ') + day(send.sent), send.back && 'Back ' + day(send.back), send.approved && 'Approved by ' + send.approvedBy + ', ' + day(send.approved)].filter(Boolean).join('. ');
    return h('tr', null, h('td', { class: 'tick' }, box),
      h('td', null, offFile(form.url, form.title), h('span', { class: 'sub' }, form.need + (form.how === 'Google eSignature' ? '. By Google eSignature' : ''))),
      h('td', { 'data-h': 'When' }, form.when),
      h('td', { 'data-h': 'Where it stands' }, h('span', { class: st ? st.cls : 'quiet' }, line), dates && h('span', { class: 'sub' }, dates), send && send.notes ? h('span', { class: 'sub' }, send.notes) : null, acts.length ? h('span', { class: 'doit' }, acts) : null));
  }
  /** A form that came back on paper: scan or photograph it, and it files like one that came by email. */
  function addSigned(send) {
    var label, input = h('input', { type: 'file', class: 'away', accept: '.pdf,.jpg,.jpeg,.png,.heic,.doc,.docx', 'aria-label': 'Add the signed copy of ' + send.form, on: { change: function () {
      var f = input.files[0];
      if (!f) return;
      if (f.size > 15000000) return alertLine('That file is too big. Keep it under 15 MB.');
      label.firstChild.textContent = 'Adding…';
      var reader = new FileReader();
      reader.onload = function () { act(null, 'signed', { row: send.row, name: f.name, data: String(reader.result).split(',')[1] }); };
      reader.onerror = function () { alertLine("We couldn't read that file."); };
      reader.readAsDataURL(f);
    } } });
    label = h('label', { class: 'link' }, 'Add the signed copy', input);
    return label;
  }
  function testPanel(p, items) {
    var bad = items.filter(function (i) { return i.error; });
    var mail = items.filter(function (i) { return !i.error && i.how === 'Email'; });
    var rows = mail.map(function (i) { return i.trackerRow; });
    var go = h('button', { class: 'btn', type: 'button', on: { click: function () {
      act(go, 'send', { rows: rows }, function (res) {
        var no = res.sent.filter(function (x) { return x.error; });
        state.working = no.length ? state.working : null;
        if (no.length) setTimeout(function () { alertLine(no[0].error); }, 0);
      });
    } } }, 'Send to ' + p.name.split(' ')[0]);
    return h('div', { class: 'note' },
      mail.length ? h('p', null, 'A test copy is on its way to you, within a minute. Check it looks right, then send it.') : null,
      bad.length ? h('ul', null, bad.map(function (i) { return h('li', { class: 'bad' }, i.title + ': ' + i.error); })) : null,
      mail.length ? h('p', { class: 'row' }, go, h('button', { class: 'link', type: 'button', on: { click: function () { state.working = null; show(); } } }, 'Not yet')) : null,
      items.some(function (i) { return i.how === 'Google eSignature'; }) ? h('p', null, 'For the eSignature forms below: open their copy, choose Tools, eSignature, Request signature, then tell us here.') : null);
  }

  /* ---- Staff files */

  function waiting(promise, draw) {
    page(h('p', { class: 'say' }, 'Loading…'));
    promise.then(draw).catch(function (e) { if (state.token) page(h('p', { class: 'note bad', role: 'alert' }, e.message)); });
  }
  function staffPage(row) {
    if (row) return waiting(call('staffFile', { person: row }), function (res) { staffFile(row, res.data); });
    waiting(call('team'), function (res) {
      var rows = res.data;
      page(h('h1', null, 'Staff files'), h('p', { class: 'lede' }, 'What each worker\u2019s file must hold, and what runs out next.'),
        rows.length ? h('table', { class: 'stack' }, h('thead', null, h('tr', null, h('th', null, 'Worker'), h('th', { class: 'num' }, 'Not recorded yet'), h('th', { class: 'num' }, 'Expired'), h('th', { class: 'num' }, 'Held'), h('th', null, 'Runs out next'))),
          h('tbody', null, rows.map(function (w) {
            return h('tr', null, h('td', null, h('a', { href: '#staff/' + w.row }, w.name)), h('td', { class: 'num', 'data-h': 'Not recorded' }, w.missing), h('td', { class: 'num', 'data-h': 'Expired' }, h('span', { class: w.expired ? 'bad' : 'quiet' }, w.expired)), h('td', { class: 'num', 'data-h': 'Held' }, w.held),
              h('td', { 'data-h': 'Next' }, w.next ? w.nextItem + ', ' + day(w.next) : ''));
          }))) : h('p', { class: 'say' }, 'No workers yet. ', h('a', { href: '#forms/people' }, 'Add them'), ' and each gets a staff file here.'));
    });
  }
  function staffFile(row, file) {
    var draw = function (data) { staffFile(row, data); };
    var flag = function (x) {
      var box = h('input', { type: 'checkbox', checked: x.yes, on: { change: function () {
        box.disabled = true;
        call('staffFlag', { person: row, sheetRow: x.row, yes: box.checked }).then(function (res) { draw(res.data); }).catch(function (e) { alertLine(e.message); box.disabled = false; });
      } } });
      return h('label', { class: 'row flag' }, box, x.label);
    };
    var date = function (i, field, value) {
      var input = h('input', { type: 'date', value: value, 'aria-label': (field === 'done' ? 'Date done, ' : 'Expiry, ') + i.item, on: { change: function () {
        input.disabled = true;
        call('staffDate', { person: row, sheetRow: i.row, field: field, value: input.value }).then(function (res) { draw(res.data); }).catch(function (e) { alertLine(e.message); input.disabled = false; });
      } } });
      return input;
    };
    var needed = file.items.filter(function (i) { return i.status !== 'Not needed'; }), left = file.items.length - needed.length, groups = [];
    needed.forEach(function (i) { var g = groups[groups.length - 1]; if (g && g.name === i.group) g.items.push(i); else groups.push({ name: i.group, items: [i] }); });
    var cls = function (s) { return s === 'Expired' ? 'bad' : s === 'Renew soon' ? 'soon' : s === 'In date' || s === 'Held' ? 'good' : 'quiet'; };
    page(h('a', { class: 'back', href: '#staff' }, 'Staff files'), h('h1', null, file.name),
      h('p', { class: 'row say' }, offFile(file.folder, 'Open their folder')),
      h('details', { class: 'more' }, h('summary', null, 'What applies to this worker'), h('div', { class: 'flags' }, file.flags.map(flag))),
      groups.map(function (g) {
        return [h('h3', null, g.name), h('table', { class: 'stack' }, h('thead', null, h('tr', null, ['What', 'Renew', 'Date done', 'Expiry or next due', 'Status'].map(function (t) { return h('th', null, t); }))),
          h('tbody', null, g.items.map(function (i) {
            return h('tr', null, h('td', null, i.item, h('span', { class: 'sub' }, i.keep), i.form && i.form !== 'None' ? h('span', { class: 'sub' }, offFile(i.formUrl, i.form)) : null), h('td', { 'data-h': 'Renew' }, i.renew),
              h('td', { 'data-h': 'Date done' }, date(i, 'done', i.done)), h('td', { 'data-h': 'Expiry' }, i.expiryWorkedOut ? day(i.expiry) : date(i, 'expiry', i.expiry)), h('td', { 'data-h': 'Status' }, h('span', { class: cls(i.status) }, i.status)));
          })))];
      }),
      left ? h('p', { class: 'say' }, plural(left, 'item') + ' not needed for this worker. Change that under What applies to this worker.') : null);
  }

  /* ---- Year planner and Rules */

  function plannerPage() {
    var draw = function () {
      var jobs = state.book.planner;
      page(h('h1', null, 'Year planner'), h('p', { class: 'lede' }, 'The dated jobs for the next 12 months: internal audits, reviews and renewals.'),
        h('table', { class: 'stack' }, h('thead', null, h('tr', null, ['Due', 'What', 'Open', 'Why', 'Done'].map(function (t) { return h('th', null, t); }))),
          h('tbody', null, jobs.map(function (j) {
            var late = j.due && !j.done && j.due < state.today;
            return h('tr', null, h('td', { 'data-h': 'Due' }, h('span', { class: late ? 'bad' : '' }, j.due ? day(j.due) : j.every)), h('td', null, j.what), h('td', { 'data-h': 'Open' }, offFile(j.url, j.open)), h('td', { 'data-h': 'Why' }, j.why),
              h('td', { 'data-h': 'Done' }, j.done ? [h('span', { class: 'good' }, day(j.done)), h('span', { class: 'doit' }, plannerButton(j, true, 'Not done yet'))] : j.due ? plannerButton(j, false, 'Done today') : null));
          }))));
    };
    if (state.book) return draw();
    waiting(call('book'), function (res) { state.book = res.data; draw(); });
  }
  function plannerButton(j, undo, word) {
    return h('button', { class: 'link', type: 'button', on: { click: function (e) {
      e.target.disabled = true;
      call('plannerDone', { row: j.row, undo: undo }).then(took).then(function (res) { state.book.planner = res.data.planner; show(); }).catch(function (err) { alertLine(err.message); e.target.disabled = false; });
    } } }, word);
  }
  function rulesPage() {
    var draw = function () {
      page(h('h1', null, 'Rules'), h('p', { class: 'lede' }, 'The rules behind the renewal dates, and where each comes from.'),
        state.book.rules.map(function (r) {
          var line = function (label, text) { return text ? h('p', { class: 'say outcome' }, h('b', null, label + ': '), text) : null; };
          return h('details', { class: 'area' }, h('summary', null, h('span', null, r.item)), h('div', { class: 'body' }, line('Who it applies to', r.who), line('Valid for, or renew', r.valid), line('Evidence to keep', r.keep), line('Recent change', r.change),
            r.source ? h('p', { class: 'say' }, offFile(r.url, r.source)) : null));
        }));
    };
    if (state.book) return draw();
    waiting(call('book'), function (res) { state.book = res.data; draw(); });
  }

  window.addEventListener('hashchange', function () { state.find = ''; state.bulk = null; window.scrollTo(0, 0); show(); });
  state.token = store('hc-tool-token');
  state.email = store('hc-tool-email') || '';
  if (state.token) start(); else signIn();
})();
