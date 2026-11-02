/* The sign page. A person opens the link from their email, sees their forms, and fills in and signs each one here,
   with no account. The form comes from the script as a PDF; what they add is stamped on it in this browser
   and the finished PDF goes back to be filed. Harmony countersign on the same page, from the tool. */
(function () {
  'use strict';
  var API = window.HC_TOOL_API, main = document.getElementById('main');
  var key = decodeURIComponent(location.hash.slice(1));
  var staffRow = Number((key.match(/^staff\/(\d+)$/) || [])[1]) || 0;
  var INK = [0.12, 0.165, 0.267];
  var state = { who: null, forms: [], ink: null, name: '' };

  function h(tag, attrs) {
    var el = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) {
      var v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === 'on') Object.keys(v).forEach(function (ev) { el.addEventListener(ev, v[ev]); });
      else if (k === 'class') el.className = v;
      else if (k in el) el[k] = v;
      else el.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(el, arguments[i]);
    return el;
  }
  function add(el, c) {
    if (c === null || c === undefined || c === false) return;
    if (Array.isArray(c)) c.forEach(function (x) { add(el, x); });
    else el.appendChild(typeof c === 'object' ? c : document.createTextNode(String(c)));
  }
  function page() {
    main.textContent = '';
    for (var i = 0; i < arguments.length; i++) add(main, arguments[i]);
    window.scrollTo(0, 0);
    main.focus({ preventScroll: true });
  }
  function today() {
    var d = new Date();
    return ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + '/' + d.getFullYear();
  }
  function session() {
    try { return localStorage.getItem('hc-tool-token'); } catch (e) { return null; }
  }
  function call(action, body) {
    return fetch(API, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, redirect: 'follow',
      body: JSON.stringify(Object.assign({ action: action }, body)) })
      .then(function (r) { if (!r.ok) throw new Error('net'); return r.json(); })
      .catch(function () { throw new Error('We could not reach Harmony Care. Check your connection and try again.'); })
      .then(function (res) { if (!res.ok) throw new Error(res.error || 'That did not work. Try again in a moment.'); return res; });
  }
  function bytesOf(b64) {
    var text = atob(b64), out = new Uint8Array(text.length);
    for (var i = 0; i < text.length; i++) out[i] = text.charCodeAt(i);
    return out;
  }
  function base64Of(bytes) {
    var parts = [];
    for (var i = 0; i < bytes.length; i += 32766) parts.push(btoa(String.fromCharCode.apply(null, bytes.subarray(i, i + 32766))));
    return parts.join('');
  }
  function stopped(message, retry) {
    page(h('p', { class: 'note bad', role: 'alert' }, message), retry ? h('p', { class: 'doit' }, h('button', { class: 'btn bigbtn', type: 'button', on: { click: retry } }, 'Try again')) : null);
  }

  /* ---- their list */

  function load() {
    page(h('p', { class: 'say' }, 'Opening your forms…'));
    call('signOpen', { link: key }).then(function (res) {
      state.who = res;
      state.name = res.signer;
      state.forms = res.forms;
      list();
    }).catch(function (e) { stopped(e.message, /reach/.test(e.message) ? load : null); });
  }
  function list() {
    var open = state.forms.filter(function (f) { return !f.done; });
    if (!state.forms.length) return page(h('h1', null, 'Nothing to sign'), h('p', { class: 'lede' }, 'Harmony Care has nothing waiting for you. If you were expecting a form, reply to their email.'));
    if (!open.length) {
      return page(h('h1', null, 'All done, thank you ' + state.who.first), h('p', { class: 'lede' }, 'Harmony Care has your signed ' + (state.forms.length === 1 ? 'form' : 'forms') + ' and there is nothing else to do. You can close this page.'));
    }
    page(h('h1', null, 'Hi ' + state.who.first),
      h('p', { class: 'lede' }, open.length === 1 ? 'Harmony Care has one form for you. It shows you where to fill in and sign.' : 'Harmony Care has ' + open.length + ' forms for you. Each one shows you where to fill in and sign.'),
      h('ul', { class: 'pack' }, state.forms.map(function (f) {
        return h('li', null, h('button', { class: 'packitem', type: 'button', disabled: f.done, on: { click: function () { openForm(f); } } },
          h('span', null, f.title, h('small', null, f.need)), h('span', { class: 'state' }, f.done ? 'Done' : 'Open')));
      })));
  }
  function openForm(f) {
    page(h('p', { class: 'say', role: 'status' }, 'Opening ' + f.title + '…'));
    call('signForm', { link: key, row: f.row }).then(function (res) {
      return showForm({ row: f.row, title: f.title, need: f.need, pdf: res.pdf, sha: res.sha });
    }).catch(function (e) { stopped(e.message, function () { openForm(f); }); });
  }

  /* ---- one form */

  function loadPdfJs() {
    return import('./vendor/pdf.min.mjs').then(function (lib) {
      lib.GlobalWorkerOptions.workerSrc = new URL('vendor/pdf.worker.min.mjs', location.href).href;
      return lib;
    });
  }

  function showForm(form) {
    var original = bytesOf(form.pdf), OPS;
    return loadPdfJs().then(function (lib) {
      OPS = lib.OPS;
      return lib.getDocument({ data: original.slice(), isEvalSupported: false }).promise;
    }).then(function (doc) {
      var sheets = [];
      var next = function (n) {
        if (n > doc.numPages) return sheets;
        return doc.getPage(n).then(function (p) {
          return Promise.all([p.getTextContent(), p.getOperatorList()]).then(function (got) {
            var view = p.getViewport({ scale: 1 });
            var cells = HCFind.cellsOf(got[1].fnArray, got[1].argsArray, OPS, view.width, view.height);
            var spots = HCFind.findSpots(got[0].items, view.width, cells, HCFind.fieldsOf(got[1].fnArray, got[1].argsArray, OPS));
            spots.forEach(function (x) { x.page = n; });
            sheets.push({ n: n, pdf: p, w: view.width, h: view.height, spots: spots });
            return next(n + 1);
          });
        });
      };
      return next(1);
    }).then(function (sheets) { drawForm(form, sheets, original); });
  }

  function drawForm(form, sheets, original) {
    var kinds = staffRow ? ['harmony', 'supervisor'] : state.who.kinds;
    var every = [], marks = [];
    sheets.forEach(function (s) { s.spots.forEach(function (x) { x.sheet = s; every.push(x); }); });
    var musts = HCFind.mine(every, kinds);
    musts.forEach(function (x) { x.must = true; });
    var needsSign = /sign|initial/i.test(form.need || 'sign');
    var dock = h('div', { class: 'in' }), holder = h('div', { class: 'dock' }, dock);
    var problem = null;

    function spotEl(x) {
      var s = x.sheet;
      var style = function (el) {
        el.style.left = (x.x / s.w * 100) + '%';
        el.style.top = ((s.h - x.y - x.h) / s.h * 100) + '%';
        el.style.width = (x.w / s.w * 100) + '%';
        el.style.height = (x.h / s.h * 100) + '%';
        return el;
      };
      var theirs = x.type === 'sign' && !x.must;
      var cls = 'spot ' + (x.must ? 'must' : theirs ? 'other' : x.type === 'tick' ? 'tick' : 'write');
      var word = x.must ? 'Sign here' : theirs ? forWhom(x) : x.type === 'tick' ? 'Tick box' : x.type === 'date' ? 'Date' : x.type === 'name' ? 'Your name' : 'Write here';
      x.el = style(h('button', { class: cls, type: 'button', 'aria-label': word, on: { click: function (e) { e.stopPropagation(); tap(x); } } }, x.must || theirs ? word : null));
      return x.el;
    }
    function spare(x) { return !x.who || kinds.indexOf(x.who) !== -1; }
    /** Every signature box shows who it is for, so nobody wonders about the ones that are not theirs. */
    function forWhom(x) {
      if (spare(x)) return 'Sign if needed';
      return { harmony: 'Harmony Care', supervisor: 'Harmony Care', representative: 'Representative', client: 'Client', worker: 'Worker', practitioner: 'Practitioner' }[x.who];
    }
    function paint(x) {
      if (!x.el) return;
      var s = x.sheet;
      x.el.textContent = '';
      x.el.classList.toggle('on', !!x.value);
      if (!x.value) { if (x.must) x.el.textContent = 'Sign here'; else if (x.type === 'sign' && !x.free) x.el.textContent = forWhom(x); return; }
      if (x.value.png) { x.el.appendChild(h('img', { src: x.value.png, alt: 'Signature' })); return; }
      var mark = h('span', { class: tall(x) ? 'mark many' : 'mark' }, x.value.text);
      mark.style.fontSize = (x.type === 'tick' ? x.h : 10) / s.w * 100 + 'cqw';
      x.el.appendChild(mark);
    }
    function set(x, value) {
      x.value = value;
      if (marks.indexOf(x) === -1 && value) marks.push(x);
      if (!value) marks.splice(marks.indexOf(x), 1);
      paint(x);
      drawDock();
    }
    /** Signing also fills the date and printed name beside it, so nobody writes them twice. */
    function signed(x) {
      set(x, { png: state.ink.png, ratio: state.ink.ratio });
      var date = HCFind.partner(x.sheet.spots, x, 'date'), name = HCFind.partner(x.sheet.spots, x, 'name');
      if (date) set(date, { text: today() });
      if (name && state.name) set(name, { text: state.name });
    }
    function tap(x) {
      problem = null;
      if (x.must || x.type === 'sign') {
        var sign = function () {
          if (x.value) return askSign(function () { signed(x); }, function () { set(x, null); });
          if (state.ink) return signed(x);
          askSign(function () { signed(x); });
        };
        return x.must || x.value || x.free || spare(x) ? sign() : askOther(forWhom(x), sign);
      }
      if (x.type === 'tick') return set(x, x.value ? null : { text: '✓' });
      askText(x.type === 'date' ? 'The date' : x.type === 'name' ? 'Your name' : 'What goes here', x.value ? x.value.text : x.type === 'date' ? today() : x.type === 'name' ? state.name : '', function (text) { set(x, text ? { text: text } : null); }, x.value || x.free ? function () { removeFree(x); } : null, x);
    }
    function removeFree(x) {
      set(x, null);
      if (x.free && x.el) { x.el.remove(); x.sheet.spots.splice(x.sheet.spots.indexOf(x), 1); }
    }
    /** A form with nothing marked still works: tap the place, and choose what goes there. */
    function tapPage(s, e, layer) {
      var box = layer.getBoundingClientRect();
      var px = (e.clientX - box.left) / box.width * s.w, py = s.h - (e.clientY - box.top) / box.height * s.h;
      var x = { type: 'text', free: true, sheet: s, x: px, y: py - 5, w: Math.min(s.w - px - 10, 220), h: 14 };
      var place = function () { s.spots.push(x); layer.appendChild(spotEl(x)); };
      var addWords = function () { askText('What goes here', '', function (text) { if (!text) return; place(); set(x, { text: text }); }); };
      var addInk = function () {
        var go = function () { x.type = 'sign'; x.y = py - 26; x.h = 34; x.w = Math.min(s.w - px - 10, 200); place(); signed(x); };
        if (state.ink) go(); else askSign(go);
      };
      askChoice(addWords, addInk);
    }

    function nextMust() { return musts.filter(function (x) { return !x.value; })[0]; }
    function signatures() { return marks.filter(function (x) { return x.value.png; }).length; }
    function ready() { return !nextMust() && (!needsSign || signatures() > 0); }
    function drawDock() {
      dock.textContent = '';
      var left = musts.filter(function (x) { return !x.value; }).length;
      var words = problem ? h('span', { class: 'bad', role: 'alert' }, problem)
        : left ? h('span', null, left === 1 ? 'One signature to go' : left + ' signatures to go')
        : musts.length ? h('span', null, 'Signed. Check it, then finish.')
        : h('span', null, needsSign && !signatures() ? 'Tap where you sign on the form.' : 'Tap anything to fill it in.');
      dock.appendChild(words);
      dock.appendChild(ready()
        ? h('button', { class: 'btn bigbtn', type: 'button', on: { click: finish } }, 'Finish')
        : left ? h('button', { class: 'btn bigbtn', type: 'button', on: { click: function () {
          var x = nextMust();
          x.el.scrollIntoView({ block: 'center', behavior: 'smooth' });
          x.el.classList.remove('now');
          void x.el.offsetWidth;
          x.el.classList.add('now');
        } } }, 'Next') : null);
    }

    function finish(e) {
      var button = e.target;
      var go = function () {
        button.disabled = true;
        button.textContent = 'Saving…';
        stamp(original, marks, form).then(function (bytes) {
          var pdf = base64Of(bytes);
          return staffRow ? call('counterDone', { token: session(), row: staffRow, pdf: pdf, name: state.name })
            : call('signDone', { link: key, row: form.row, pdf: pdf, name: state.name });
        }).then(function () {
          holder.remove();
          if (staffRow) return page(h('h1', null, 'Signed and approved'), h('p', { class: 'lede' }, form.title + ' for ' + form.person + ' is filed.'), h('p', null, h('a', { href: '../#forms/people' }, 'Back to the tool')));
          state.forms.forEach(function (f) { if (f.row === form.row) f.done = true; });
          list();
        }).catch(function (err) {
          problem = err.message;
          drawDock();
        });
      };
      if (state.name) return go();
      askText('Your full name', '', function (text) { if (text) { state.name = text; go(); } });
    }

    var stage = h('div', { class: 'sheets' }, sheets.map(function (s) {
      var canvas = h('canvas', { 'aria-label': 'Page ' + s.n + ' of ' + form.title });
      var layer = h('div', { class: 'layer', on: { click: function (e) { if (e.target === layer) tapPage(s, e, layer); } } }, s.spots.filter(function (x) {
        return x.type !== 'sign' || x.must || !staffRow || kinds.indexOf(x.who) !== -1;  // on a returned copy the others have signed
      }).map(spotEl));
      var sheet = h('div', { class: 'leaf' }, canvas, layer);
      sheet.style.aspectRatio = s.w + ' / ' + s.h;
      sheet.style.containerType = 'inline-size';
      s.box = sheet;
      s.canvas = canvas;
      return sheet;
    }));

    page(h('button', { class: 'back', type: 'button', on: { click: function () {
      holder.remove();
      if (staffRow) location.href = '../#forms/people'; else list();
    } } }, staffRow ? 'Back to the tool' : 'Back to your forms'), h('h1', { class: 'formhead' }, form.title),
      h('p', { class: 'say' }, 'Tap a box to tick it. Tap any space on the form to type in it.'), stage);
    document.body.appendChild(holder);
    drawDock();
    drawWhenSeen(sheets);
  }

  /** Long forms draw their pages as they come near the screen, and let go of the far ones, so a phone copes. */
  function drawWhenSeen(sheets) {
    var scale = Math.min(window.devicePixelRatio || 1, 2);
    var watch = new IntersectionObserver(function (seen) {
      seen.forEach(function (entry) {
        var s = sheets.filter(function (x) { return x.box === entry.target; })[0];
        if (entry.isIntersecting && !s.drawn) {
          s.drawn = true;
          var view = s.pdf.getViewport({ scale: entry.target.clientWidth / s.w * scale });
          s.canvas.width = view.width;
          s.canvas.height = view.height;
          s.pdf.render({ canvasContext: s.canvas.getContext('2d'), viewport: view });
        } else if (!entry.isIntersecting && s.drawn && sheets.length > 12) {
          s.drawn = false;
          s.canvas.width = 0;
        }
      });
    }, { rootMargin: '900px 0px' });
    sheets.forEach(function (s) { watch.observe(s.box); });
  }

  /* ---- the questions */

  function dialog(title, body) {
    var d = h('dialog', { 'aria-label': title, on: { close: function () { d.remove(); } } }, h('h2', null, title), body);
    document.body.appendChild(d);
    d.showModal();
    return d;
  }
  function askOther(whose, yes) {
    var d = dialog('This box is for: ' + whose, [h('p', { class: 'say' }, 'Sign it only if that is you.'), h('div', { class: 'doit' },
      h('button', { class: 'btn', type: 'button', on: { click: function () { d.close(); yes(); } } }, 'It is mine'),
      h('button', { class: 'linkbtn', type: 'button', on: { click: function () { d.close(); } } }, 'Leave it'))]);
  }
  function askChoice(words, ink) {
    var d = dialog('What goes here?', h('div', { class: 'doit' },
      h('button', { class: 'btn', type: 'button', on: { click: function () { d.close(); words(); } } }, 'Words'),
      h('button', { class: 'btn', type: 'button', on: { click: function () { d.close(); ink(); } } }, 'My signature'),
      h('button', { class: 'linkbtn', type: 'button', on: { click: function () { d.close(); } } }, 'Cancel')));
  }
  function tall(x) { return !!x.cell && x.h >= 26; }
  /** The most letters a box can hold at the smallest size the stamp will go to. */
  function roomFor(x) {
    if (!x || !x.cell) return 200;
    return Math.min(2000, Math.max(4, Math.floor((x.w - 2) / 3) * Math.max(1, Math.floor(x.h / 7.2))));
  }
  function askText(title, value, done, remove, box) {
    var input = box && tall(box) ? h('textarea', { rows: 4, value: value, 'aria-label': title, maxLength: roomFor(box) })
      : h('input', { type: 'text', value: value, 'aria-label': title, maxLength: roomFor(box) });
    var save = function () { var text = input.value.trim(); d.close(); done(text); };
    input.addEventListener('keydown', function (e) { if (e.key === 'Enter' && input.tagName === 'INPUT') { e.preventDefault(); save(); } });
    var d = dialog(title, [input, h('div', { class: 'doit' },
      h('button', { class: 'btn', type: 'button', on: { click: save } }, 'Save'),
      remove ? h('button', { class: 'linkbtn', type: 'button', on: { click: function () { d.close(); remove(); } } }, 'Remove') : null,
      h('button', { class: 'linkbtn', type: 'button', on: { click: function () { d.close(); } } }, 'Cancel'))]);
  }
  function inkFrom(canvas) {
    var ctx = canvas.getContext('2d'), w = canvas.width, h2 = canvas.height, data = ctx.getImageData(0, 0, w, h2).data;
    var left = w, right = 0, top = h2, bottom = 0;
    for (var y = 0; y < h2; y++) for (var x = 0; x < w; x++) {
      if (data[(y * w + x) * 4 + 3] > 10) { left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y); }
    }
    if (right <= left) return null;
    var cut = document.createElement('canvas');
    cut.width = right - left + 8;
    cut.height = bottom - top + 8;
    cut.getContext('2d').drawImage(canvas, left - 4, top - 4, cut.width, cut.height, 0, 0, cut.width, cut.height);
    return { png: cut.toDataURL('image/png'), ratio: cut.width / cut.height };
  }
  function typedInk(name) {
    var c = document.createElement('canvas');
    c.width = 900;
    c.height = 220;
    var ctx = c.getContext('2d');
    ctx.fillStyle = '#1F2A44';
    ctx.font = 'italic 96px Georgia, "Times New Roman", serif';
    ctx.textBaseline = 'middle';
    ctx.fillText(name, 20, 110, 860);
    return inkFrom(c);
  }
  function askSign(done, remove) {
    var drawing = true, drew = false;
    var name = h('input', { type: 'text', value: state.name, autocomplete: 'name', maxLength: 120, 'aria-label': 'Your full name' });
    var canvas = h('canvas', { 'aria-label': 'Draw your signature here' });
    var hint = h('span', { class: 'hint' }, 'Draw your signature here');
    var pad = h('div', { class: 'pad' }, canvas, h('div', { class: 'line' }), hint);
    var typed = h('input', { type: 'text', 'aria-label': 'Type your signature', hidden: true, maxLength: 120, placeholder: 'Type your name' });
    var wrong = h('p', { class: 'bad', role: 'alert', hidden: true });
    var drawTab = h('button', { type: 'button', 'aria-pressed': 'true', on: { click: function () { mode(true); } } }, 'Draw');
    var typeTab = h('button', { type: 'button', 'aria-pressed': 'false', on: { click: function () { mode(false); } } }, 'Type');
    function mode(draw) {
      drawing = draw;
      pad.hidden = !draw;
      typed.hidden = draw;
      drawTab.setAttribute('aria-pressed', String(draw));
      typeTab.setAttribute('aria-pressed', String(!draw));
      if (!draw && !typed.value) typed.value = name.value;
    }
    var save = function () {
      var who = name.value.trim();
      if (!who) { wrong.textContent = 'Your full name is needed.'; wrong.hidden = false; return; }
      var ink = drawing ? (drew ? inkFrom(canvas) : null) : typedInk(typed.value.trim());
      if (!ink) { wrong.textContent = drawing ? 'Draw your signature, or choose Type.' : 'Type your name.'; wrong.hidden = false; return; }
      state.name = who;
      state.ink = ink;
      d.close();
      done();
    };
    var d = dialog('Your signature', [
      h('label', null, 'Your full name', name), h('div', { class: 'tabs' }, drawTab, typeTab), pad, typed, wrong,
      h('p', { class: 'legal' }, 'By pressing Sign, I agree this is my signature on this form, and that signing here counts the same as signing on paper.'),
      h('div', { class: 'doit' }, h('button', { class: 'btn', type: 'button', on: { click: save } }, 'Sign'),
        h('button', { class: 'linkbtn', type: 'button', on: { click: function () {
          drew = false;
          canvas.getContext('2d').clearRect(0, 0, canvas.width, canvas.height);
          hint.hidden = false;
        } } }, 'Clear'),
        remove ? h('button', { class: 'linkbtn', type: 'button', on: { click: function () { d.close(); remove(); } } }, 'Remove') : null,
        h('button', { class: 'linkbtn', type: 'button', on: { click: function () { d.close(); } } }, 'Cancel'))]);
    var ratio = window.devicePixelRatio || 1, box = canvas.getBoundingClientRect();
    canvas.width = box.width * ratio;
    canvas.height = box.height * ratio;
    var ctx = canvas.getContext('2d');
    ctx.scale(ratio, ratio);
    ctx.lineWidth = 2.2;
    ctx.lineCap = ctx.lineJoin = 'round';
    ctx.strokeStyle = '#1F2A44';
    var down = false, at = function (e) { var b = canvas.getBoundingClientRect(); return [e.clientX - b.left, e.clientY - b.top]; };
    canvas.addEventListener('pointerdown', function (e) {
      down = true;
      canvas.setPointerCapture(e.pointerId);
      var p = at(e);
      ctx.beginPath();
      ctx.moveTo(p[0], p[1]);
      ctx.lineTo(p[0] + 0.1, p[1] + 0.1);
      ctx.stroke();
      drew = true;
      hint.hidden = true;
      e.preventDefault();
    });
    canvas.addEventListener('pointermove', function (e) {
      if (!down) return;
      var p = at(e);
      ctx.lineTo(p[0], p[1]);
      ctx.stroke();
    });
    ['pointerup', 'pointercancel'].forEach(function (n) { canvas.addEventListener(n, function () { down = false; }); });
  }

  /* ---- stamping the PDF */

  /** Standard PDF fonts cannot show every script, so anything outside them goes on as a picture of the words. */
  function plain(text) { return /^[\x20-\x7E\xA0-\xFF‘’“”–—]*$/.test(text); }
  function wordsPng(text, size) {
    var c = document.createElement('canvas'), ctx = c.getContext('2d');
    ctx.font = (size * 4) + 'px Helvetica, Arial, sans-serif';
    c.width = Math.ceil(ctx.measureText(text).width) + 8;
    c.height = size * 6;
    ctx.font = (size * 4) + 'px Helvetica, Arial, sans-serif';
    ctx.fillStyle = '#1F2A44';
    ctx.textBaseline = 'middle';
    ctx.fillText(text, 2, c.height / 2);
    return { url: c.toDataURL('image/png'), w: c.width / 4, h: c.height / 4 };
  }
  function drawWords(doc, page, font, text, x, y, size, max) {
    var color = PDFLib.rgb(INK[0], INK[1], INK[2]);
    if (plain(text)) {
      var s = size;
      while (s > 6 && font.widthOfTextAtSize(text, s) > max) s -= 0.5;
      page.drawText(text, { x: x, y: y, size: s, font: font, color: color });
      return Promise.resolve();
    }
    var pic = wordsPng(text, size);
    return doc.embedPng(pic.url).then(function (img) {
      var fit = Math.min(1, max / pic.w);
      page.drawImage(img, { x: x, y: y - 3, width: pic.w * fit, height: pic.h * fit });
    });
  }
  /** A long answer wraps inside its cell, smaller if it has to. */
  function drawLines(page, font, x) {
    var size = 10.5, lines = [];
    var wrap = function (s) {
      var out = [];
      x.value.text.split('\n').forEach(function (para) {
        var line = '';
        para.split(' ').forEach(function (word) {
          var next = line ? line + ' ' + word : word;
          if (line && font.widthOfTextAtSize(next, s) > x.w - 2) { out.push(line); line = word; } else line = next;
        });
        out.push(line);
      });
      return out;
    };
    do { size -= 0.5; lines = wrap(size); } while (size > 6 && lines.length * size * 1.2 > x.h);
    lines.forEach(function (line, i) {
      page.drawText(line, { x: x.x + 1, y: x.y + x.h - size * 1.2 * (i + 1), size: size, font: font, color: PDFLib.rgb(INK[0], INK[1], INK[2]) });
    });
    return null;
  }
  /** In a measured cell the ink stays in the cell. On a ruled line it may rise above the line, as a pen does. */
  function fitInk(x) {
    var ratio = x.value.ratio, h2 = x.cell ? x.h : Math.max(x.h, 22) + 4, w = h2 * ratio;
    if (w > x.w) { w = x.w; h2 = w / ratio; }
    return { x: x.x + 1, y: x.cell ? x.y + (x.h - h2) / 2 : x.y, width: w, height: h2 };
  }
  function stamp(original, marks, form) {
    return PDFLib.PDFDocument.load(original, { ignoreEncryption: true }).then(function (doc) {
      return doc.embedFont(PDFLib.StandardFonts.Helvetica).then(function (font) {
        return doc.embedFont(PDFLib.StandardFonts.HelveticaBold).then(function (bold) {
          var pages = doc.getPages(), color = PDFLib.rgb(INK[0], INK[1], INK[2]);
          var steps = marks.reduce(function (chain, x) {
            return chain.then(function () {
              var p = pages[x.sheet.n - 1];
              if (x.value.png) return doc.embedPng(x.value.png).then(function (img) { p.drawImage(img, fitInk(x)); });
              if (x.type === 'tick') {
                var l = x.x + x.w * 0.2, b = x.y + x.h * 0.45, m = x.x + x.w * 0.42, n = x.y + x.h * 0.2, r = x.x + x.w * 0.85, t = x.y + x.h * 0.85;
                p.drawLine({ start: { x: l, y: b }, end: { x: m, y: n }, thickness: 1.3, color: color });
                p.drawLine({ start: { x: m, y: n }, end: { x: r, y: t }, thickness: 1.3, color: color });
                return null;
              }
              if (tall(x) && plain(x.value.text.replace(/\n/g, ' '))) return drawLines(p, font, x);
              return drawWords(doc, p, font, x.value.text.replace(/\n/g, ' '), x.x + 1, x.y + (x.cell ? Math.max((x.h - 10) / 2, 1) + 2 : 3), 10, x.cell ? x.w - 2 : Math.max(x.w, 60));
            });
          }, Promise.resolve());
          return steps.then(function () { return record(doc, font, bold, form); }).then(function () { return doc.save(); });
        });
      });
    });
  }
  /** The last page says who signed, when, and exactly which document it was. */
  function record(doc, font, bold, form) {
    var page = doc.addPage([595, 842]), y = 780, color = PDFLib.rgb(0.2, 0.2, 0.2);
    var now = new Date();
    var when = now.toLocaleString('en-AU', { dateStyle: 'long', timeStyle: 'short' }) + ' (' + Intl.DateTimeFormat().resolvedOptions().timeZone + ')';
    page.drawText('Signing record', { x: 50, y: y, size: 16, font: bold, color: color });
    y -= 36;
    var rows = [['Form', form.title], ['Signed by', state.name], ['Signed on', when],
      ['How', staffRow ? 'Signed for Harmony Care in the Harmony Care tool, signed in as ' + (form.signer || 'staff') : 'Signed on Harmony Care’s sign page, from a private link emailed to them'],
      ['Agreed', 'They agreed this is their signature and that it counts the same as signing on paper'],
      ['Document', 'SHA-256 of the form as they opened it:'], ['', form.sha || '']];
    return rows.reduce(function (chain, r) {
      return chain.then(function () {
        if (r[0]) page.drawText(r[0], { x: 50, y: y, size: 10, font: bold, color: color });
        var done = drawWords(doc, page, font, r[1], 130, y, 9.5, 415);
        y -= 22;
        return done;
      });
    }, Promise.resolve());
  }

  /* ---- start */

  function start() {
    if (staffRow) {
      if (!session()) return stopped('Sign in to the tool first, then open this form from there.');
      page(h('p', { class: 'say' }, 'Opening the form…'));
      return call('counterForm', { token: session(), row: staffRow }).then(function (res) {
        state.name = '';
        return showForm({ row: staffRow, title: res.title, person: res.person, signer: res.signer, need: 'Sign', pdf: res.pdf, sha: res.sha });
      }).catch(function (e) { stopped(e.message, start); });
    }
    if (!key) return stopped('This link is not complete. Open it from your email again.');
    load();
  }
  start();
})();
