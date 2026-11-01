/* Reads a form's PDF text and works out what a person does on it: where to tick, write and sign.
   Harmony's forms come from many templates, so nothing here is set up per form. It takes pdf.js text items
   and returns spots in PDF points (origin bottom left). No page or network here, so it runs in the tests. */
(function (root) {
  'use strict';
  var BOX = /[☐☑☒□▢⬜]/;
  var RUN = /(?:[….·]{2,}\/){2}[….·]{2,}|[….·]{6,}|…{3,}|_{5,}/g;
  var DATED = /\//;
  var ROLES = [
    ['worker', /client.?s worker/i],
    ['', /interpreter/i],
    ['representative', /signed by the representative|executed by the representative|advocates?\s*\/\s*guardian|family\s*\/\s*guardian|supporter|guardian.?s signature|parent.*consent/i],
    ['harmony', /on behalf|services\)?,? by:?$|by:$|organisation representative|authorised harmony|executed by silp|key management|service provider/i],
    ['supervisor', /supervisor|appraiser|manager|principal|assessor|confirmed by/i],
    ['client', /signed by the client|executed by the client|^client|^the person\b|client.?s signature|client \/ client/i],
    ['worker', /signed by the worker|worker|staff member|staff attending|employee|appraisee|applicant|tester/i],
    ['practitioner', /practitioner|prescriber|developer|trainer|health professional|\bprofession\b/i]
  ];

  function itemsOf(textItems) {
    return textItems.filter(function (i) { return i.str && i.str.trim(); }).map(function (i) {
      return { s: i.str, x: i.transform[4], y: i.transform[5], w: i.width, h: Math.abs(i.transform[3]) || i.height || 10 };
    });
  }

  function linesOf(items) {
    var sorted = items.slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    var lines = [];
    sorted.forEach(function (it) {
      var line = lines[lines.length - 1];
      if (line && Math.abs(line.y - it.y) < 3) line.items.push(it);
      else lines.push({ y: it.y, items: [it] });
    });
    lines.forEach(function (l) {
      l.items.sort(function (a, b) { return a.x - b.x; });
      l.text = l.items.map(function (i) { return i.s; }).join(' ').replace(/\s+/g, ' ').trim();
      l.left = l.items[0].x;
    });
    return lines;
  }

  /** A role, '' for a named party that is nobody's job here, or null when the words say nothing. */
  function role(text) {
    for (var i = 0; i < ROLES.length; i++) if (ROLES[i][1].test(text)) return ROLES[i][0];
    return null;
  }

  function blanksOn(line) {
    var out = [];
    line.items.forEach(function (it) {
      var m, re = new RegExp(RUN.source, 'g');
      while ((m = re.exec(it.s))) {
        var share = it.w / it.s.length;
        var blank = { x: it.x + m.index * share, w: m[0].length * share, y: it.y, h: it.h, dated: DATED.test(m[0]) };
        if (blank.w >= 24) out.push(blank);
      }
    });
    return out;
  }

  /** The role of a signature spot: its own label first, then the lines just under it, then the heading above. */
  function roleFor(lines, index) {
    var own = role(lines[index].text);
    if (own !== null) return own;
    var i;
    for (i = index + 1; i <= index + 2 && i < lines.length; i++) {
      if (lines[i - 1].y - lines[i].y > 34) break;
      var below = lines[i].text;
      if (below.length < 70 && !/please print/i.test(below)) { var r = role(below); if (r !== null) return r; }
    }
    for (i = index - 1; i >= 0 && i >= index - 8; i--) {
      if (lines[i].text.length < 90) { var up = role(lines[i].text); if (up !== null) return up; }
    }
    return '';
  }

  /** Table cells, from the rectangles the PDF draws them with: pdf.js operator list in, boxes in points out. */
  function cellsOf(fns, args, OPS, width, height) {
    var ctm = [1, 0, 0, 1, 0, 0], stack = [], seen = {}, cells = [];
    var times = function (m, n) {
      return [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
              m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
    };
    var skip = {};
    skip[OPS.moveTo] = skip[OPS.lineTo] = 2;
    skip[OPS.curveTo] = 6;
    skip[OPS.curveTo2] = skip[OPS.curveTo3] = 4;
    fns.forEach(function (fn, i) {
      if (fn === OPS.save) stack.push(ctm);
      else if (fn === OPS.restore) ctm = stack.pop() || ctm;
      else if (fn === OPS.transform) ctm = times(ctm, args[i]);
      else if (fn === OPS.constructPath) {
        var at = 0, c = args[i][1];
        args[i][0].forEach(function (op) {
          if (op !== OPS.rectangle) { at += skip[op] || 0; return; }
          var x1 = ctm[0] * c[at] + ctm[2] * c[at + 1] + ctm[4], y1 = ctm[1] * c[at] + ctm[3] * c[at + 1] + ctm[5];
          var x2 = ctm[0] * (c[at] + c[at + 2]) + ctm[2] * (c[at + 1] + c[at + 3]) + ctm[4];
          var y2 = ctm[1] * (c[at] + c[at + 2]) + ctm[3] * (c[at + 1] + c[at + 3]) + ctm[5];
          at += 4;
          var box = { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) };
          var key = [box.x, box.y, box.w, box.h].map(Math.round).join();
          var inPage = box.y > 34 && box.y + box.h < height - 34 && box.w < width - 60;
          if (box.w >= 36 && box.h >= 12 && inPage && !seen[key]) { seen[key] = true; cells.push(box); }
        });
      }
    });
    return oneOfEach(cells);
  }

  function shared(a, b) {
    var w = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x), h = Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y);
    return w > 0 && h > 0 ? w * h : 0;
  }

  /** A cell's shading and its border are two rectangles for one cell: keep the larger. */
  function oneOfEach(cells) {
    var kept = [];
    cells.slice().sort(function (a, b) { return b.w * b.h - a.w * a.h; }).forEach(function (c) {
      var area = c.w * c.h;
      if (!kept.some(function (k) { return shared(k, c) > area * 0.9 && area > k.w * k.h * 0.5; })) kept.push(c);
    });
    return kept;
  }

  function wordsIn(cell, items) {
    return items.filter(function (i) {
      var cx = i.x + i.w / 2, cy = i.y + i.h / 3;
      return cx > cell.x && cx < cell.x + cell.w && cy > cell.y && cy < cell.y + cell.h;
    }).sort(function (a, b) { return b.y - a.y || a.x - b.x; });
  }

  function textOf(items) {
    return items.map(function (i) { return i.s; }).join(' ').replace(/\s+/g, ' ').trim();
  }

  /** The widest stretch of a one-line cell with nothing in it, and the words to its left. */
  function roomIn(cell, words) {
    if (words[0].y - words[words.length - 1].y > 4) return null;
    var row = words.slice().sort(function (a, b) { return a.x - b.x; }), at = cell.x + 2, best = null;
    row.concat([{ x: cell.x + cell.w - 2, w: 0 }]).forEach(function (i) {
      if (!best || i.x - at > best.w) best = { x: at, w: i.x - at };
      at = Math.max(at, i.x + i.w);
    });
    best.before = textOf(row.filter(function (i) { return i.x < best.x; }));
    return best;
  }

  function kindOf(label) {
    if (/signature|\bsigned by\b|^signed$|^sign\b|\binitials\b|^initial(\s*\d)?$|^initial of\b/i.test(label) && !/^date/i.test(label)) return 'sign';
    if (/\bdate\b/i.test(label)) return 'date';
    if (/\bname\b/i.test(label)) return 'name';
    return 'text';
  }

  /** Every blank in a table: an empty cell, or the room left in one, named by the label beside it or the heading above it.
      A column of cells under one heading is a group. */
  function cellSpots(cells, items, lines) {
    var inside = cells.map(function (c) { return wordsIn(c, items); });
    var word = inside.map(textOf);
    var overlapY = function (a, b) { return Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) > Math.min(a.h, b.h) * 0.6; };
    var overlapX = function (cell, o) { return Math.min(cell.x + cell.w, o.x + o.w) - Math.max(cell.x, o.x) > cell.w * 0.6; };
    // a label with a blank of its own, or a line of tick boxes, is not the heading of the cells below it
    var heading = function (j) {
      if (word[j].length >= 60 || BOX.test(word[j])) return false;
      var room = roomIn(cells[j], inside[j]);
      return !(room && /:$/.test(word[j]) && room.w >= 60);
    };
    var out = [];
    cells.forEach(function (cell, i) {
      var left = -1, up = -1, column = -1, rowStart = -1, dated = false;
      cells.forEach(function (o, j) {
        if (j === i || !word[j]) return;
        if (overlapY(o, cell) && Math.abs(o.x + o.w - cell.x) < 4) left = j;
        if (overlapY(o, cell) && o.x < cell.x && (rowStart === -1 || o.x < cells[rowStart].x)) rowStart = j;
        if (overlapY(o, cell) && /^date\b/i.test(word[j])) dated = true;
        if (overlapX(cell, o) && o.y > cell.y) {
          if (up === -1 || o.y < cells[up].y) up = j;
          var same = Math.abs(o.x - cell.x) < 4 && Math.abs(o.w - cell.w) < 6;
          if (same && heading(j) && (column === -1 || o.y < cells[column].y)) column = j;
        }
      });
      var beside = left !== -1 && word[left].length < 60 ? word[left] : '';
      if (up === -1 || !heading(up)) up = column;
      if (up !== column && cells[up].y - cell.y > 120) up = -1;
      var above = up !== -1 ? word[up] : '';
      var box = { x: cell.x + 2, y: cell.y + 2, w: cell.w - 4, h: Math.max(cell.h - 4, 9) };
      var label = beside;
      if (word[i]) {
        var room = roomIn(cell, inside[i]);
        var last = inside[i][inside[i].length - 1];
        var blankBeside = cells.some(function (o, j) { return !word[j] && overlapY(o, cell) && Math.abs(cell.x + cell.w - o.x) < 4; });
        var overhead = blankBeside || cells.some(function (o, j) { return o.y < cell.y && !word[j] && overlapX(o, cell); });
        if (room && /:$/.test(room.before) && room.before.length < 60 && room.w >= 60) label = room.before;
        else if (room && room.w >= 90 && beside && kindOf(beside) === 'sign') label = beside;
        else if (kindOf(word[i]) === 'sign' && word[i].length < 30 && last.y - cell.y >= 18 && !overhead) {
          // a label at the top of a tall cell, with the blank under it
          out.push({ type: 'sign', x: box.x, y: box.y, w: box.w, h: last.y - cell.y - 4, cell: true, who: role(word[i]) !== null ? role(word[i]) : roleFor(lines, 0) });
          return;
        } else return;
        box.x = room.x + 4;
        box.w = room.w - 8;
        above = '';
      }
      if (!label && !above) return;
      var kind = kindOf(label);
      if (kind === 'text' && above) kind = kindOf(above);
      if (kind === 'text' && dated && /client|representative/.test(role(label))) kind = 'sign';  // "The person: ____  Date: ____"
      var line = 0;
      while (line < lines.length - 1 && lines[line].y > cell.y + cell.h) line++;
      var own = role(label), row = rowStart !== -1 ? role(word[rowStart]) : null, head = role(above);
      var spot = { type: kind, x: box.x, y: box.y, w: box.w, h: box.h, cell: true };
      if (kind === 'sign') {
        spot.who = own !== null ? own : row !== null ? row : head !== null ? head : roleFor(lines, Math.max(line - 1, 0));
        if (!label) spot.group = up;
        if (/initial/i.test(label || above)) spot.initials = true;
      }
      out.push(spot);
    });
    return out;
  }

  /** The same blank found twice is one spot: the surest kind wins, and a measured cell beats a guess from the words. */
  function oneSpotEach(spots) {
    var rank = { tick: 0, sign: 1, date: 2, name: 3, text: 4 }, kept = [];
    spots.slice().sort(function (a, b) { return rank[a.type] - rank[b.type] || (b.cell ? 1 : 0) - (a.cell ? 1 : 0); }).forEach(function (s) {
      var small = function (k) { return Math.min(s.w * s.h, k.w * k.h) / 2; };
      if (s.type === 'tick' || !kept.some(function (k) { return k.type !== 'tick' && shared(k, s) > small(k); })) kept.push(s);
    });
    return kept.sort(function (a, b) { return Math.abs(a.y - b.y) > 4 ? b.y - a.y : a.x - b.x; });
  }

  /** @param {Array} textItems pdf.js items for one page @param {number} width page width in points @param {Array} [cells] from cellsOf */
  function findSpots(textItems, width, cells) {
    var items = itemsOf(textItems), lines = linesOf(items), spots = [];
    var right = (width || 595) - 36;
    var signed = {};
    var measured = cellSpots(cells || [], items, lines);
    // a label a measured box already answers, beside it or in the column under it
    var answered = function (it) {
      return measured.some(function (s) {
        if (s.type !== 'sign') return false;
        var beside = s.x >= it.x && s.y < it.y + it.h && s.y + s.h > it.y - 4;
        var under = s.y < it.y && s.x < it.x + it.w && s.x + s.w > it.x;
        return beside || under;
      });
    };

    lines.forEach(function (line) {
      line.items.forEach(function (it) {
        for (var i = 0; i < it.s.length; i++) {
          if (!BOX.test(it.s.charAt(i))) continue;
          var share = it.w / it.s.length;
          var size = Math.min(it.h * 1.1, 13);
          spots.push({ type: 'tick', x: it.x + i * share, y: it.y - 1, w: size, h: size });
        }
      });
    });

    lines.forEach(function (line, index) {
      blanksOn(line).forEach(function (b) {
        var before = line.items.filter(function (i) { return i.x + 2 < b.x; }).map(function (i) { return i.s.replace(/[…._]/g, ''); }).join(' ');
        var next = lines[index + 1];
        var near = next && line.y - next.y < 26 && line.y - next.y > 3;
        var under = near ? next.text : '';
        var box = { x: b.x, y: b.y - 1, w: b.w, h: Math.max(b.h + 4, 14) };
        if (b.dated || /date:?\s*$/i.test(before)) { box.type = 'date'; spots.push(box); return; }
        if (/^signature\b/i.test(under) || /(signature|initials):?\s*$/i.test(before)) {
          box.type = 'sign';
          box.initials = /initials:?\s*$/i.test(before);
          box.who = roleFor(lines, near && /^signature\b/i.test(under) ? index + 1 : index);
          box.h = Math.max(box.h, 26);
          signed[index] = signed[index + 1] = true;
          spots.push(box);
          return;
        }
        if (/^name\b/i.test(under) || /name:?\s*$/i.test(before)) { box.type = 'name'; spots.push(box); return; }
        box.type = 'text';
        spots.push(box);
      });
    });

    // a signature with no line to write on: a table cell beside or under its label, or a line lost to a page break
    lines.forEach(function (line, index) {
      if (signed[index]) return;
      line.items.forEach(function (it, n) {
        var s = it.s.trim();
        if (!/signature/i.test(s) || /^signatures:?$/i.test(s) || s.length > 44 || /[.]$/.test(s) || /^section/i.test(s) || answered(it)) return;
        var after = line.items[n + 1];
        var room = (after ? after.x : right) - (it.x + it.w);
        var spot = { type: 'sign', who: roleFor(lines, index), h: 26 };
        if (room > 90) { spot.x = it.x + it.w + 8; spot.w = Math.min(room - 12, 220); spot.y = it.y - 6; }
        else { spot.x = it.x; spot.w = Math.max(it.w, 110); spot.y = it.y - 30; }
        spots.push(spot);
      });
    });

    // 'Date:' with nothing after it, usually beside a signature
    lines.forEach(function (line) {
      var last = line.items[line.items.length - 1];
      if (!/^date:$/i.test(last.s.trim()) || right - last.x - last.w < 40) return;
      spots.push({ type: 'date', x: last.x + last.w + 4, y: last.y - 1, w: Math.min(right - last.x - last.w, 110), h: 14 });
    });
    return oneSpotEach(spots.concat(measured));
  }

  /** In a column of rows a person needs only the first: the rest stay there for anyone who needs them. */
  function firstOfEach(spots) {
    var seen = {};
    return spots.filter(function (s) {
      if (s.group === undefined) return true;
      var key = s.page + ':' + s.group;
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
  }

  /** What a form really asks for. Initials count only where there is no signature to give, a table of rows only where there is
      no signature line, and more than three of anything is a record that repeats: its first will do. */
  function asked(spots) {
    var full = spots.filter(function (s) { return !s.initials; });
    full = full.length ? full : spots;
    var lines = full.filter(function (s) { return s.group === undefined; });
    var rows = lines.length ? lines : firstOfEach(full);
    return rows.length > 3 ? rows.slice(0, 1) : rows;
  }

  /** The spots this signer must sign: the first of their roles the form has, or those it names no one for. */
  function mine(spots, kinds) {
    var signs = spots.filter(function (s) { return s.type === 'sign'; });
    for (var i = 0; i < kinds.length; i++) {
      var own = signs.filter(function (s) { return s.who === kinds[i]; });
      if (own.length) return asked(own);
    }
    return asked(signs.filter(function (s) { return !s.who; }));
  }

  /** The empty date or name that belongs with a signature: its own row or just under it, never a date further along a register. */
  function partner(spots, sign, type) {
    var best = null, far = Infinity;
    spots.forEach(function (o) {
      if (o.type !== type || o.value) return;
      if (o.y > sign.y + sign.h + 12 || o.y + o.h < sign.y - 48) return;
      var down = sign.y - o.y;
      var along = o.x >= sign.x + sign.w ? o.x - sign.x - sign.w : sign.x >= o.x + o.w ? sign.x - o.x - o.w : 0;
      if (type === 'date' && Math.abs(down) < 8 && along > 60) return;
      var d = Math.abs(down) * 4 + along;
      if (d < far) { far = d; best = o; }
    });
    return best;
  }

  var api = { findSpots: findSpots, cellsOf: cellsOf, mine: mine, partner: partner };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.HCFind = api;
})(this);
