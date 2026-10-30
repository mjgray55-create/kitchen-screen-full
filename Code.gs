/**
 * Kitchen screen, Sheet edition: server side.
 * Lives inside the household's own "Kitchen" Google Sheet (Extensions → Apps Script).
 * Everything the screen shows is kept on the sheet's tabs, so it can also be read and
 * changed from a phone or a computer.
 *
 * The tabs are made automatically the first time the screen opens (or from the
 * "Kitchen screen" menu in the sheet), so a blank copy of the template just works.
 */

var DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
var TIERS = {
  basic: { meals: true, shopping: true, stock: false, recipes: false, looks: false },
  expanded: { meals: true, shopping: true, stock: true, recipes: true, looks: false },
  primo: { meals: true, shopping: true, stock: true, recipes: true, looks: true }
};
var DEFAULT_HEADINGS = ['Meat & Fish', 'Ready Meals & Batch Cooking', 'Vegetables & Fruit', 'Snacks, Breakfasts & Lunches', 'Bread & Baking', 'Other'];
var CELL_LIMIT = 49000;   // Google Sheets holds up to 50,000 characters in one cell

/* ---------- web app entry, and the menu in the sheet ---------- */

function doGet() {
  ensureTabs_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kitchen')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Kitchen screen')
    .addItem('Set up or repair the tabs', 'setUpTabs')
    .addToUi();
}

/** Makes any missing tab. Never changes or removes anything already there. */
function setUpTabs() {
  var made = ensureTabs_();
  var msg = made.length ? 'Added: ' + made.join(', ') + '.' : 'All the tabs are already there.';
  try { SpreadsheetApp.getUi().alert('Kitchen screen', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(msg); }
}

/* ---------- small helpers ---------- */

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) { ensureTabs_(); s = ss_().getSheetByName(name); }
  if (!s) throw new Error('Can\'t find the "' + name + '" tab in the Kitchen sheet.');
  return s;
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function num_(v) {
  var n = parseFloat(String(v).replace(',', '.'));
  return isNaN(n) ? null : n;
}
function norm_(s) { return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim(); }
function yes_(v) { return /^(y|yes|true|on|1)$/i.test(String(v).trim()); }

/* ---------- the tabs, and what a blank one looks like ---------- */

var TABS = [
  { name: 'Menu', head: ['Day', 'Lunch', 'Dinner', 'Notes', "Who's cooking"], widths: [110, 260, 260, 220, 140], rows: DAYS.map(function (d) { return [d, '', '', '', '']; }) },
  { name: 'Meals', head: ['Meal', 'What goes with it', 'Recipe', 'Notes'], widths: [220, 260, 200, 260] },
  { name: 'Shopping', head: ['Got it', 'Item', 'How many', 'Where', 'Added by'], widths: [70, 260, 110, 160, 140], checkbox: true },
  { name: 'Fridge', head: ['Leftover', 'Portions', 'Date', 'Notes'], widths: [240, 90, 110, 260], dateCol: 3 },
  { name: 'Freezer', freezer: true },
  { name: 'Consumed', head: ['Item', 'Qty', 'Unit', 'Date', 'Freezer', 'Note'], widths: [240, 70, 100, 110, 80, 200], dateCol: 4 },
  { name: 'Recipes', head: ['Recipe', 'Section', 'From', 'Ingredients (one per line)', 'Method (one step per line)', 'Notes', 'Imported (leave alone)'],
    widths: [200, 140, 160, 280, 360, 220, 140], wrap: true },
  { name: 'Birthdays', head: ['Name', 'Date (e.g. March 4)', 'Who (Person, Dog, Cat or Anniversary)'], widths: [200, 160, 260], textCol: 2 },
  { name: 'Settings', settings: true }
];

/* Settings tab: one row per setting, found by the label in column A, so the rows can be moved around. */
var SETTINGS = [
  ['Household name', 'household', '', 'Optional. Shows in the browser tab.'],
  ['Package', 'tier', 'Basic', 'Basic, Expanded or Primo. Picking one on the screen also sets the five rows below.'],
  ['Meals list', 'f:meals', 'Yes', 'Yes or No'],
  ['Shopping list', 'f:shopping', 'Yes', 'Yes or No'],
  ['Fridge & freezer', 'f:stock', 'No', 'Yes or No. Freezer headings are the heading rows on the Freezer tab.'],
  ['Recipe book', 'f:recipes', 'No', 'Yes or No'],
  ['Seasons, holidays & birthdays', 'f:looks', 'No', 'Yes or No. Birthdays and anniversaries go on the Birthdays tab.'],
  ['Look', 'theme', 'Dark', 'Dark, Light, or Follow the tablet'],
  ['Who cooks', 'cooks', 'Takeaway, Eating out', 'Separate with commas. A single name shows as "Sam\'s cooking".'],
  ['Freezers', 'freezers', 'KF = Kitchen freezer', 'Code = name, separated by semicolons, e.g. KF = Kitchen freezer; UR = Basement upright'],
  ['Falling decorations', 'decorations', 'All year', 'All year, Only on special days, or Off'],
  ['Set up on the tablet', 'setupDone', 'No', 'Yes once the first-time setup on the screen is finished.']
];
var THEMES = { dark: 'Dark', light: 'Light', auto: 'Follow the tablet' };
var DECOS = { always: 'All year', special: 'Only on special days', off: 'Off' };
var TIER_NAMES = { basic: 'Basic', expanded: 'Expanded', primo: 'Primo' };

function ensureTabs_() {
  var ss = ss_(), made = [];
  TABS.forEach(function (t) {
    if (ss.getSheetByName(t.name)) return;
    var s = ss.insertSheet(t.name);
    made.push(t.name);
    if (t.freezer) return buildFreezer_(s);
    if (t.settings) return buildSettings_(s);
    s.getRange(1, 1, 1, t.head.length).setValues([t.head]).setFontWeight('bold').setBackground('#efe7dc');
    s.setFrozenRows(1);
    t.widths.forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
    if (t.rows) s.getRange(2, 1, t.rows.length, t.head.length).setValues(t.rows);
    if (t.wrap) s.getRange('A:G').setWrap(true).setVerticalAlignment('top');
    if (t.dateCol) s.getRange(2, t.dateCol, 999, 1).setNumberFormat('yyyy-mm-dd');
    if (t.textCol) s.getRange(2, t.textCol, 999, 1).setNumberFormat('@');
  });
  // a brand-new spreadsheet starts with an empty "Sheet1": tidy it away once the real tabs exist
  var extra = ss.getSheetByName('Sheet1');
  if (extra && made.length && ss.getSheets().length > 1 && extra.getLastRow() === 0) ss.deleteSheet(extra);
  return made;
}

/* Freezer: title, "Last updated", a header row, then each heading on its own row with its items below it. */
function buildFreezer_(s) {
  s.getRange('A1').setValue('Freezer').setFontSize(14).setFontWeight('bold');
  s.getRange('B2:C2').setValues([['Last updated', '']]);
  s.getRange('C2').setNumberFormat('yyyy-mm-dd');
  s.getRange(3, 1, 1, 7).setValues([['', 'Item', 'Qty', 'Unit', 'Date in', 'Freezer', 'Comment']])
    .setFontWeight('bold').setBackground('#efe7dc');
  s.setFrozenRows(3);
  [200, 240, 60, 100, 110, 80, 220].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
  DEFAULT_HEADINGS.forEach(function (h, i) { styleHeading_(s.getRange(4 + i, 1, 1, 7)); s.getRange(4 + i, 1).setValue(h); });
  s.getRange('E:E').setNumberFormat('yyyy-mm-dd');
}
function styleHeading_(range) { range.setFontWeight('bold').setBackground('#f6efe6'); }

function buildSettings_(s) {
  s.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'What it means']]).setFontWeight('bold').setBackground('#efe7dc');
  s.setFrozenRows(1);
  s.getRange(2, 1, SETTINGS.length, 3).setValues(SETTINGS.map(function (x) { return [x[0], x[2], x[3]]; }));
  s.getRange(2, 2, SETTINGS.length, 1).setNumberFormat('@');
  s.getRange(2, 3, SETTINGS.length, 1).setFontStyle('italic').setFontColor('#7a736b');
  [220, 260, 520].forEach(function (w, i) { s.setColumnWidth(i + 1, w); });
}

/* ---------- everything the screen shows, in one call ---------- */

function getAll() {
  var settings = readSettings_();
  var frz = readFreezer_();
  settings.headings = frz.categories.map(function (c) { return c.name; });
  var freezers = {};
  settings.freezers.forEach(function (f) { freezers[f.code] = f.name; });
  return {
    settings: settings,
    menu: readMenu_(),
    meals: readMeals_(),
    fridge: readFridge_(),
    cooks: settings.cooks,
    freezer: frz,
    consumed: readConsumed_(12),
    shopping: readShopping_(),
    freezers: freezers,
    loadedAt: new Date().toISOString()
  };
}

/* ---------- Settings ---------- */

function settingsRows_() {
  var s = sheet_('Settings');
  var last = Math.max(1, s.getLastRow());
  var v = s.getRange(1, 1, last, 2).getDisplayValues();
  var rows = {};
  v.forEach(function (r, i) { var k = norm_(r[0]); if (k && !rows[k]) rows[k] = { row: i + 1, value: String(r[1]).trim() }; });
  return { sheet: s, rows: rows };
}

function readSettings_() {
  var got = settingsRows_().rows;
  function val(label, dflt) { var x = got[norm_(label)]; return x && x.value !== '' ? x.value : dflt; }
  var out = { features: {}, lastBackup: '' };
  SETTINGS.forEach(function (d) {
    var v = val(d[0], d[2]), key = d[1];
    if (key.indexOf('f:') === 0) out.features[key.slice(2)] = yes_(v);
    else if (key === 'tier') out.tier = /primo/i.test(v) ? 'primo' : /expand|full/i.test(v) ? 'expanded' : 'basic';
    else if (key === 'theme') out.theme = /light/i.test(v) ? 'light' : /follow|auto|tablet/i.test(v) ? 'auto' : 'dark';
    else if (key === 'decorations') out.decorations = /off|never|no/i.test(v) ? 'off' : /special/i.test(v) ? 'special' : 'always';
    else if (key === 'cooks') out.cooks = v.split(',').map(function (c) { return c.trim(); }).filter(Boolean);
    else if (key === 'freezers') out.freezers = parseFreezers_(v);
    else if (key === 'setupDone') out.setupDone = yes_(v);
    else out[key] = v;
  });
  if (!out.freezers.length) out.freezers = [{ code: 'KF', name: 'Kitchen freezer' }];
  out.occasions = readBirthdays_();
  return out;
}

/* "KF = Kitchen freezer; UR = Basement upright" → [{ code, name }] */
function parseFreezers_(v) {
  return String(v || '').split(/[;\n]/).map(function (p) {
    var m = p.match(/^\s*([^=:]+?)\s*[=:]\s*(.+?)\s*$/);
    return m ? { code: m[1].toUpperCase().replace(/\s+/g, ''), name: m[2] } : null;
  }).filter(Boolean);
}

/**
 * Save the Settings page from the screen.
 * renames: [[oldHeading, newHeading], ...] for freezer headings that were renamed.
 */
function saveSettings(s, renames) {
  return withLock_(function () {
    var st = settingsRows_(), sh = st.sheet;
    var f = s.features || {};
    var out = {
      household: s.household || '',
      tier: TIER_NAMES[s.tier] || 'Basic',
      'f:meals': f.meals ? 'Yes' : 'No', 'f:shopping': f.shopping ? 'Yes' : 'No', 'f:stock': f.stock ? 'Yes' : 'No',
      'f:recipes': f.recipes ? 'Yes' : 'No', 'f:looks': f.looks ? 'Yes' : 'No',
      theme: THEMES[s.theme] || 'Dark',
      cooks: (s.cooks || []).join(', '),
      freezers: (s.freezers || []).map(function (x) { return x.code + ' = ' + x.name; }).join('; '),
      decorations: DECOS[s.decorations] || 'All year',
      setupDone: s.setupDone ? 'Yes' : 'No'
    };
    SETTINGS.forEach(function (d) {
      var hit = st.rows[norm_(d[0])];
      var row = hit ? hit.row : sh.getLastRow() + 1;
      if (!hit) sh.getRange(row, 1, 1, 3).setValues([[d[0], '', d[3]]]);
      sh.getRange(row, 2).setNumberFormat('@').setValue(out[d[1]]);
    });
    if (s.occasions) writeBirthdays_(s.occasions);
    if (s.headings) syncHeadings_(s.headings, renames || []);
    return getAll();
  });
}

/* Make the Freezer tab's heading rows match the Settings page: rename, add at the end, drop empty ones. */
function syncHeadings_(headings, renames) {
  var m = freezerMap_(), s = m.sheet;
  renames.forEach(function (r) {
    var c = m.cats.filter(function (x) { return x.name === r[0]; })[0];
    if (c && r[1]) s.getRange(c.row, 1).setValue(r[1]);
  });
  m = freezerMap_();
  // remove headings that were taken off, bottom first, only when nothing is filed under them
  m.cats.slice().reverse().forEach(function (c) {
    if (headings.indexOf(c.name) < 0 && !c.items.length && c.row > m.headerRow) s.deleteRow(c.row);
  });
  m = freezerMap_();
  var have = m.cats.map(function (c) { return c.name; });
  headings.forEach(function (h) {
    if (have.indexOf(h) >= 0) return;
    var r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, 7).clearContent();
    styleHeading_(s.getRange(r, 1, 1, 7));
    s.getRange(r, 1).setValue(h);
    have.push(h);
  });
}

/* ---------- Birthdays (Name | Date | Who) ---------- */

function readBirthdays_() {
  var s = ss_().getSheetByName('Birthdays');
  if (!s || s.getLastRow() < 2) return [];
  return s.getRange(2, 1, s.getLastRow() - 1, 3).getDisplayValues()
    .map(function (r) { return { name: String(r[0]).trim(), date: String(r[1]).trim(), who: String(r[2]).trim() || 'Person' }; })
    .filter(function (b) { return b.name && b.date; });
}

function writeBirthdays_(list) {
  var s = sheet_('Birthdays');
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 3).clearContent();
  if (!list.length) return;
  s.getRange(2, 2, list.length, 1).setNumberFormat('@');
  s.getRange(2, 1, list.length, 3).setValues(list.map(function (o) { return [o.name, o.date, o.who || 'Person']; }));
}

/* ---------- Menu ---------- */

function readMenu_() {
  var v = sheet_('Menu').getRange(2, 1, 7, 5).getDisplayValues();
  return v.map(function (r, i) {
    return { day: DAYS[i], lunch: r[1], dinner: r[2], notes: r[3], cook: r[4] };
  });
}

var SLOT_COL = { lunch: 2, dinner: 3, notes: 4, cook: 5 };

function setMenu(dayIndex, slot, value) {
  var col = SLOT_COL[slot];
  if (!col || dayIndex < 0 || dayIndex > 6) throw new Error('Unknown menu slot.');
  return withLock_(function () {
    sheet_('Menu').getRange(2 + dayIndex, col).setValue(value || '');
    return getAll();
  });
}

/** Kick the can: a planned lunch or dinner moves to another day.
    mode: 'move' (other day is empty), 'swap' the two days, 'join' (add alongside),
    or 'replace' what's there (e.g. last week's meal on a day that has gone). */
function kickMenu(fromDay, toDay, slot, expected, mode) {
  var col = { lunch: 2, dinner: 3 }[slot];
  if (!col || fromDay < 0 || fromDay > 6 || toDay < 0 || toDay > 6 || fromDay === toDay) throw new Error('Unknown menu slot.');
  return withLock_(function () {
    var menu = sheet_('Menu');
    var a = menu.getRange(2 + fromDay, col), b = menu.getRange(2 + toDay, col);
    var av = String(a.getDisplayValue()).trim(), bv = String(b.getDisplayValue()).trim();
    if (av !== String(expected || '').trim()) throw new Error('That meal changed on the sheet — refreshing.');
    if (!av) return getAll();
    if (mode === 'swap') { a.setValue(bv); b.setValue(av); }
    else if (mode === 'join') { b.setValue(bv ? bv + ' + ' + av : av); a.setValue(''); }
    else if (mode === 'replace') { b.setValue(av); a.setValue(''); }
    else {
      if (bv) throw new Error('Something is already planned that day — refreshing.');
      b.setValue(av); a.setValue('');
    }
    return getAll();
  });
}

function clearMenu() {
  return withLock_(function () {
    sheet_('Menu').getRange(2, 2, 7, 4).clearContent();
    return getAll();
  });
}

/* ---------- Meals ---------- */

function readMeals_() {
  var s = sheet_('Meals');
  var last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 4).getDisplayValues()
    .map(function (r, i) { return { row: i + 2, name: r[0], contents: r[1], recipe: r[2], notes: r[3] }; })
    .filter(function (m) { return m.name; });
}

function addMeal(name, contents, recipe, notes) {
  name = String(name || '').trim();
  if (!name) throw new Error('The meal needs a name.');
  return withLock_(function () {
    sheet_('Meals').appendRow([name, contents || '', recipe || '', notes || '']);
    return getAll();
  });
}

function deleteMeal(row, name) {
  return withLock_(function () {
    var s = sheet_('Meals');
    s.deleteRow(checkRow_(s, row, 1, name));
    return getAll();
  });
}

/* ---------- Freezer (title, Last updated, header row, headings on their own rows) ---------- */

function freezerMap_() {
  var s = sheet_('Freezer');
  var v = s.getDataRange().getDisplayValues();
  var headerRow = -1, updatedRow = -1;
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][1]).trim() === 'Item') { headerRow = i; break; }
    if (/^last updated/i.test(String(v[i][1]).trim())) updatedRow = i;
  }
  if (headerRow < 0) throw new Error('Can\'t find the Item / Qty header row on the Freezer tab.');
  var cats = [], cur = null;
  for (var r = headerRow + 1; r < v.length; r++) {
    var a = String(v[r][0]).trim(), b = String(v[r][1]).trim();
    if (a && !b) { cur = { name: a, row: r + 1, items: [] }; cats.push(cur); continue; }
    if (!b) continue;
    if (!cur) { cur = { name: 'Other', row: headerRow + 1, items: [] }; cats.push(cur); }
    cur.items.push({
      row: r + 1, name: b, qty: v[r][2], unit: v[r][3], date: v[r][4],
      f: String(v[r][5]).trim().toUpperCase(), comment: v[r][6]
    });
  }
  return {
    sheet: s, cats: cats, headerRow: headerRow + 1,
    updated: updatedRow >= 0 ? v[updatedRow][2] : '', updatedRow: updatedRow >= 0 ? updatedRow + 1 : -1
  };
}

function readFreezer_() {
  var m = freezerMap_();
  return { updated: m.updated, categories: m.cats.map(function (c) { return { name: c.name, items: c.items }; }) };
}

function touchFreezer_(m) {
  if (m.updatedRow > 0) m.sheet.getRange(m.updatedRow, 3).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
}

/** Take some out of the freezer: lowers the count and logs it on Consumed. */
function useFreezer(row, name, amount, note) {
  amount = num_(amount);
  if (!amount || amount <= 0) throw new Error('How much was used?');
  return withLock_(function () {
    takeFromFreezer_(freezerMap_(), row, name, amount, note);
    return getAll();
  });
}

function takeFromFreezer_(m, row, name, amount, note) {
  var s = m.sheet;
  row = checkRow_(s, row, 2, name);
  var vals = s.getRange(row, 1, 1, 7).getDisplayValues()[0];
  var have = num_(vals[2]);
  logConsumed_(name, amount, vals[3], vals[5], note);
  if (have === null || have - amount <= 0.0001) {
    s.deleteRow(row);
  } else {
    s.getRange(row, 3).setValue(Math.round((have - amount) * 100) / 100);
  }
  touchFreezer_(m);
}

/** Correct a count without logging anything as eaten. */
function setFreezerQty(row, name, qty) {
  qty = num_(qty);
  if (qty === null || qty < 0) throw new Error('That isn\'t a number.');
  return withLock_(function () {
    var m = freezerMap_();
    row = checkRow_(m.sheet, row, 2, name);
    if (qty === 0) m.sheet.deleteRow(row);
    else m.sheet.getRange(row, 3).setValue(qty);
    touchFreezer_(m);
    return getAll();
  });
}

/** Put something new in the freezer, at the end of its heading. */
function addFreezer(category, item, qty, unit, f, comment) {
  item = String(item || '').trim();
  if (!item) throw new Error('What are you putting in?');
  return withLock_(function () {
    putInFreezer_(category, item, qty, unit, f, comment);
    return getAll();
  });
}

function putInFreezer_(category, item, qty, unit, f, comment) {
  var m = freezerMap_();
  var s = m.sheet;
  var cat = category ? m.cats.filter(function (c) { return c.name === category; })[0] : m.cats[0];
  if (!cat) throw new Error('No "' + category + '" heading on the Freezer tab.');
  var after = cat.items.length ? cat.items[cat.items.length - 1].row : cat.row;
  s.insertRowAfter(after);
  var r = after + 1;
  var q = num_(qty);
  s.getRange(r, 1, 1, 7)
    .setFontWeight('normal').setBackground(null)
    .setValues([['', item, q === null ? '' : q, unit || '', new Date(), String(f || '').toUpperCase(), comment || '']]);
  s.getRange(r, 5).setNumberFormat('yyyy-mm-dd');
  touchFreezer_(m);
}

/* ---------- Fridge (leftovers) ---------- */

function readFridge_() {
  var s = sheet_('Fridge');
  var last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 4).getDisplayValues()
    .map(function (r, i) { return { row: i + 2, name: String(r[0]).trim(), portions: r[1], date: r[2], notes: r[3] }; })
    .filter(function (x) { return x.name; });
}

function addLeftover(item, portions, note) {
  item = String(item || '').trim();
  var p = num_(portions);
  if (!item) throw new Error('What\'s left over?');
  if (!p || p <= 0) throw new Error('How many portions?');
  return withLock_(function () {
    var s = sheet_('Fridge');
    var r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, 4).setValues([[item, p, new Date(), note || '']]);
    s.getRange(r, 3).setNumberFormat('yyyy-mm-dd');
    return getAll();
  });
}

/** Correct the portion count. Zero takes it off the list. */
function setLeftoverPortions(row, name, portions) {
  var p = num_(portions);
  if (p === null || p < 0) throw new Error('That isn\'t a number.');
  return withLock_(function () {
    var s = sheet_('Fridge');
    row = checkRow_(s, row, 1, name);
    if (p === 0) s.deleteRow(row); else s.getRange(row, 2).setValue(p);
    return getAll();
  });
}

function removeLeftover(row, name) { return setLeftoverPortions(row, name, 0); }

/** Leftover is getting old: put it in the freezer instead. */
function moveLeftoverToFreezer(row, name, category, f) {
  return withLock_(function () {
    var s = sheet_('Fridge');
    row = checkRow_(s, row, 1, name);
    var p = num_(s.getRange(row, 2).getDisplayValue());
    putInFreezer_(category, name, p === null ? 1 : p, 'portions', f, 'leftovers, from the fridge');
    s.deleteRow(checkRow_(s, row, 1, name));
    return getAll();
  });
}

/* ---------- "Eaten" on the menu ----------
   A lunch or dinner cell holds one or more things joined by " + ".
   Each can be tagged:  "Bean soup (leftovers) ×2"   "Sausage rolls (freezer)"
   Untagged things are meals and never touch any counts.
   Once eaten, the thing gets a "✓ " in front so it can't be counted twice. */

function parseItem_(raw) {
  var m = String(raw).trim().match(/^(✓\s*)?(.*?)(?:\s*\((leftovers|freezer)\))?(?:\s*×\s*(\d+(?:\.\d+)?))?$/);
  return { raw: String(raw).trim(), eaten: !!m[1], name: m[2].trim(), kind: m[3] || 'meal', qty: m[4] ? parseFloat(m[4]) : 1 };
}

function eatMenuItem(dayIndex, slot, raw, amount) {
  var col = { lunch: 2, dinner: 3 }[slot];
  if (!col || dayIndex < 0 || dayIndex > 6) throw new Error('Unknown menu slot.');
  return withLock_(function () {
    var menu = sheet_('Menu');
    var cell = menu.getRange(2 + dayIndex, col);
    var parts = String(cell.getDisplayValue()).split(/\s+\+\s+/);
    var idx = -1;
    for (var i = 0; i < parts.length; i++) if (parts[i].trim() === String(raw).trim()) { idx = i; break; }
    if (idx < 0) throw new Error('That isn\'t on the menu any more — refreshing.');
    var it = parseItem_(parts[idx]);
    if (it.eaten) return getAll();
    var qty = num_(amount) || it.qty;
    if (it.kind === 'leftovers') {
      var f = sheet_('Fridge'), v = readFridge_().filter(function (x) { return norm_(x.name) === norm_(it.name); })[0];
      if (!v) throw new Error('There\'s no "' + it.name + '" in the fridge list.');
      var have = num_(v.portions);
      if (have === null || have - qty <= 0.0001) f.deleteRow(v.row); else f.getRange(v.row, 2).setValue(have - qty);
    } else if (it.kind === 'freezer') {
      var m = freezerMap_(), hit = null;
      m.cats.forEach(function (c) { c.items.forEach(function (x) { if (!hit && norm_(x.name) === norm_(it.name)) hit = x; }); });
      if (!hit) throw new Error('There\'s no "' + it.name + '" on the freezer list.');
      takeFromFreezer_(m, hit.row, hit.name, qty, slot);
    }
    parts[idx] = '✓ ' + parts[idx].trim();
    cell.setValue(parts.join(' + '));
    return getAll();
  });
}

/* ---------- Consumed ---------- */

function logConsumed_(name, qty, unit, f, note) {
  var s = sheet_('Consumed');
  var r = s.getLastRow() + 1;
  s.getRange(r, 1, 1, 6).setValues([[name, qty, unit || '', new Date(), f || '', note || '']]);
  s.getRange(r, 4).setNumberFormat('yyyy-mm-dd');
}

function readConsumed_(n) {
  var s = sheet_('Consumed');
  var last = s.getLastRow();
  if (last < 2) return [];
  var v = s.getRange(2, 1, last - 1, 6).getDisplayValues().filter(function (r) { return r[0]; });
  return v.slice(-n).reverse().map(function (r) {
    return { name: r[0], qty: r[1], unit: r[2], date: r[3], f: r[4], comment: r[5] };
  });
}

/* ---------- Shopping ---------- */

function readShopping_() {
  var s = sheet_('Shopping');
  var last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 5).getValues()
    .map(function (r, i) {
      return { row: i + 2, done: r[0] === true, item: String(r[1]).trim(), qty: String(r[2]), where: String(r[3]), by: String(r[4]) };
    })
    .filter(function (x) { return x.item; });
}

function addShopping(item, qty, where) {
  item = String(item || '').trim();
  if (!item) throw new Error('What do you need?');
  return withLock_(function () {
    var s = sheet_('Shopping');
    var r = s.getLastRow() + 1;
    s.getRange(r, 1, 1, 5).setValues([[false, item, qty || '', where || '', 'Kitchen screen']]);
    s.getRange(r, 1).insertCheckboxes();
    return getAll();
  });
}

function setShoppingDone(row, item, done) {
  return withLock_(function () {
    var s = sheet_('Shopping');
    s.getRange(checkRow_(s, row, 2, item), 1).setValue(!!done);
    return getAll();
  });
}

function clearTicked() {
  return withLock_(function () {
    var s = sheet_('Shopping');
    var last = s.getLastRow();
    if (last >= 2) {
      var v = s.getRange(2, 1, last - 1, 1).getValues();
      for (var i = v.length - 1; i >= 0; i--) if (v[i][0] === true) s.deleteRow(i + 2);
    }
    return getAll();
  });
}

/* ---------- Recipes (one row each; ingredients and method one per line) ---------- */

function getRecipes() {
  var s = sheet_('Recipes');
  var last = s.getLastRow();
  if (last < 2) return [];
  return s.getRange(2, 1, last - 1, 7).getValues()
    .map(function (r, i) {
      var x = { id: i + 2, title: String(r[0]).trim(), section: String(r[1]).trim() || 'Recipes', source: String(r[2]),
        ingredients: String(r[3]).split('\n'), method: String(r[4]).split('\n'), notes: String(r[5]) };
      if (String(r[6]).trim()) x.html = String(r[6]);
      return x;
    })
    .filter(function (x) { return x.title; });
}

/** Add a recipe (no id) or save changes to one (id = its row; was = its name before editing). */
function saveRecipe(r) {
  var title = String(r.title || '').trim();
  if (!title) throw new Error('The recipe needs a name.');
  var row = [title, String(r.section || '').trim() || 'Recipes', r.source || '',
    [].concat(r.ingredients || []).join('\n').trim(), [].concat(r.method || []).join('\n').trim(), r.notes || '', ''];
  row.forEach(function (c) { if (String(c).length > CELL_LIMIT) throw new Error('That recipe is too long for one cell of the sheet. Try splitting it in two.'); });
  return withLock_(function () {
    var s = sheet_('Recipes');
    var at = r.id ? checkRow_(s, r.id, 1, r.was || title) : s.getLastRow() + 1;
    s.getRange(at, 1, 1, 7).setValues([row]);
    return at;
  });
}

function deleteRecipe(id, title) {
  return withLock_(function () {
    var s = sheet_('Recipes');
    s.deleteRow(checkRow_(s, id, 1, title));
    return true;
  });
}

/** Recipes read from an HTML recipe file on the tablet. Skips any whose name is already in the book. */
function importRecipes(list) {
  return withLock_(function () {
    var s = sheet_('Recipes');
    var have = getRecipes().map(function (r) { return norm_(r.title); });
    var rows = [];
    (list || []).forEach(function (r) {
      var t = String(r.title || '').trim();
      if (!t || have.indexOf(norm_(t)) >= 0 || String(r.html).length > CELL_LIMIT) return;
      have.push(norm_(t));
      rows.push([t, r.section || 'Recipes', '', '', '', '', r.html]);
    });
    if (rows.length) s.getRange(s.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
    return rows.length;
  });
}

/* ---------- safety: someone may have edited the sheet since the screen loaded ---------- */

function checkRow_(s, row, col, expected) {
  expected = String(expected).trim();
  if (row >= 1 && row <= s.getLastRow() && String(s.getRange(row, col).getDisplayValue()).trim() === expected) return row;
  var v = s.getRange(1, col, Math.max(1, s.getLastRow()), 1).getDisplayValues();
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === expected) return i + 1;
  throw new Error('"' + expected + '" isn\'t there any more — someone may have just changed the sheet. Refreshing.');
}
