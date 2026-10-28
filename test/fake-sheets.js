/* A small stand-in for Google Sheets / Apps Script, so Code.gs can run in a browser for testing.
   Only covers what Code.gs uses. Not used on the real screen. */
(function (g) {
  function pad(n) { return String(n).padStart(2, '0'); }
  function show(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
    if (v === true) return 'TRUE';
    if (v === false) return 'FALSE';
    return String(v);
  }
  function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }

  function Range(sheet, r, c, nr, nc) { this.s = sheet; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  Range.prototype.cells = function (fn) {
    var out = [];
    for (var i = 0; i < this.nr; i++) { var row = []; for (var j = 0; j < this.nc; j++) row.push(fn(this.r + i, this.c + j)); out.push(row); }
    return out;
  };
  Range.prototype.getValues = function () { var s = this.s; return this.cells(function (r, c) { return s.get(r, c); }); };
  Range.prototype.getDisplayValues = function () { var s = this.s; return this.cells(function (r, c) { return show(s.get(r, c)); }); };
  Range.prototype.getDisplayValue = function () { return show(this.s.get(this.r, this.c)); };
  Range.prototype.getValue = function () { return this.s.get(this.r, this.c); };
  Range.prototype.setValues = function (v) {
    if (v.length !== this.nr || v[0].length !== this.nc) throw new Error('setValues: size mismatch ' + v.length + 'x' + v[0].length + ' vs ' + this.nr + 'x' + this.nc);
    for (var i = 0; i < this.nr; i++) for (var j = 0; j < this.nc; j++) this.s.set(this.r + i, this.c + j, v[i][j]);
    return this;
  };
  Range.prototype.setValue = function (v) { for (var i = 0; i < this.nr; i++) for (var j = 0; j < this.nc; j++) this.s.set(this.r + i, this.c + j, v); return this; };
  Range.prototype.clearContent = function () { return this.setValue(''); };
  Range.prototype.insertCheckboxes = function () { for (var i = 0; i < this.nr; i++) if (this.s.get(this.r + i, this.c) === '') this.s.set(this.r + i, this.c, false); return this; };
  ['setFontWeight', 'setBackground', 'setFontSize', 'setNumberFormat', 'setWrap', 'setVerticalAlignment', 'setFontStyle', 'setFontColor']
    .forEach(function (k) { Range.prototype[k] = function () { return this; }; });

  function Sheet(name) { this.name = name; this.rows = []; }
  Sheet.prototype.getName = function () { return this.name; };
  Sheet.prototype.get = function (r, c) { var row = this.rows[r - 1]; var v = row ? row[c - 1] : undefined; return v === undefined ? '' : v; };
  Sheet.prototype.set = function (r, c, v) {
    if (r < 1 || c < 1) throw new Error('Bad cell ' + r + ',' + c);
    while (this.rows.length < r) this.rows.push([]);
    this.rows[r - 1][c - 1] = v instanceof Date ? new Date(v.getTime()) : v;
  };
  Sheet.prototype.getLastRow = function () {
    for (var i = this.rows.length - 1; i >= 0; i--) if ((this.rows[i] || []).some(function (v) { return v !== '' && v !== undefined && v !== null; })) return i + 1;
    return 0;
  };
  Sheet.prototype.getLastColumn = function () { var m = 0; this.rows.forEach(function (r) { if (r && r.length > m) m = r.length; }); return m; };
  Sheet.prototype.getRange = function (a, b, c, d) {
    if (typeof a === 'string') {
      var m = a.match(/^([A-Z]+)(\d*)(?::([A-Z]+)(\d*))?$/);
      var c1 = colNum(m[1]), r1 = m[2] ? +m[2] : 1, c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] : (m[2] ? r1 : Math.max(1, this.getLastRow()));
      return new Range(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    if (c === 0 || d === 0) throw new Error('getRange: zero rows or columns');
    return new Range(this, a, b, c || 1, d || 1);
  };
  Sheet.prototype.getDataRange = function () { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); };
  Sheet.prototype.insertRowAfter = function (r) { this.rows.splice(r, 0, []); };
  Sheet.prototype.deleteRow = function (r) { if (r < 1 || r > this.rows.length) throw new Error('deleteRow out of range: ' + r); this.rows.splice(r - 1, 1); };
  Sheet.prototype.appendRow = function (v) { var r = this.getLastRow() + 1; for (var j = 0; j < v.length; j++) this.set(r, j + 1, v[j]); };
  Sheet.prototype.setFrozenRows = function () {};
  Sheet.prototype.setColumnWidth = function () {};
  // a readable dump for checking the template layout
  Sheet.prototype.dump = function () { return this.rows.slice(0, this.getLastRow()).map(function (r) { return (r || []).map(show); }); };

  var book = { sheets: [new Sheet('Sheet1')] };
  var ss = {
    getSheetByName: function (n) { return book.sheets.filter(function (s) { return s.name === n; })[0] || null; },
    insertSheet: function (n) { var s = new Sheet(n); book.sheets.push(s); return s; },
    getSheets: function () { return book.sheets.slice(); },
    deleteSheet: function (s) { book.sheets.splice(book.sheets.indexOf(s), 1); }
  };
  g.SpreadsheetApp = { getActiveSpreadsheet: function () { return ss; }, getUi: function () { throw new Error('no UI in tests'); } };
  g.LockService = { getScriptLock: function () { return { waitLock: function () {}, releaseLock: function () {} }; } };
  g.Logger = { log: function () {} };
  g.__BOOK__ = ss;
})(window);
