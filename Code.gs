/**
 * AARF Karaoke & Event Rentals - booking backend (Google Apps Script)
 *
 * Sheets:
 *   "Requests" -> every booking request from the website. You change the Status
 *                 column (B) with the dropdown: New / Accepted / Hold / Declined.
 *   "Dates"    -> the dates the website shows as booked or hold.
 *                 Rows added by this script say "request" in column C.
 *                 Rows you type yourself are never touched by the script.
 *                 Status can be: booked, hold, or unavailable (a date you are closed).
 *                 Column D is the package. Each package is booked separately, so Standard
 *                 can be booked while Pro and Premium are still open. Leave D blank to
 *                 block ALL packages (for example a date you are closed).
 *
 * What happens when you change Status in "Requests":
 *   Accepted -> the date becomes BOOKED on the website, and a reply message is written
 *   Hold     -> the date becomes ON HOLD on the website, and a reply message is written
 *   Declined -> the date is freed (if nothing else uses it), and a reply message is written
 *   New      -> the date is freed (if nothing else uses it), and the reply is cleared
 *
 * Run setup() once from the editor (and again after pasting a new version).
 */

// Put your email here to get a message for every new request. Leave '' to turn off.
const NOTIFY_EMAIL = '';
const BUSINESS = 'AARF Karaoke & Event Rentals';

const DATES_SHEET = 'Dates';
const REQUESTS_SHEET = 'Requests';
const STATUSES = ['New', 'Accepted', 'Hold', 'Declined'];
// Added to the confirmation reply.
const REMINDER = 'A ₱1,000 security deposit is collected upon delivery. It is refundable upon pick up, as long as there is no damage and no missing equipment. Delivery is free within 5 km, then ₱30 for every additional km.';
const PACKAGES_SHEET = 'Packages';
const ADDONS_SHEET = 'Add-ons';
const SETTINGS_SHEET = 'Settings';
const SPECIAL_SHEET = 'Special Prices';
const AVAIL_SHEET = 'Availability by Date';
const SITE_TZ = 'Asia/Manila'; // used to decide what "today" is for the booking window

// Requests columns (1-based)
const C = { RECEIVED: 1, STATUS: 2, NAME: 3, PHONE: 4, PKG: 5, PRICE: 6, DATE: 7, TIME: 8, HOURS: 9, VENUE: 10, DELIVERY: 11, NOTES: 12, REPLY: 13, SMS: 14, EVENT: 15, ADDONS: 16, TOTAL: 17, LOC: 18 };

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupContent(ss); // the Packages and Add-ons sheets you can edit
  setupSpecial(ss); // the Special Prices sheet (holiday or date-specific rates)
  setupAvailability(ss); // the Availability by Date sheet (turn packages / add-ons off on certain dates)
  setupSettings(ss); // the Settings sheet (header text)

  let d = ss.getSheetByName(DATES_SHEET);
  if (!d) d = ss.insertSheet(DATES_SHEET);
  d.getRange(1, 1, 1, 4).setValues([['Date (YYYY-MM-DD)', 'Status (booked, hold or unavailable)', 'Source (auto)', 'Package (blank = all packages)']]);
  d.getRange(2, 4, 1000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(PACKAGES_SHEET).getRange('B2:B30'), true).setAllowInvalid(false).build());
  d.getRange('A:A').setNumberFormat('@');
  d.setFrozenRows(1);
  d.getRange(2, 2, 1000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['booked', 'hold', 'unavailable'], true).setAllowInvalid(false).build());

  let r = ss.getSheetByName(REQUESTS_SHEET);
  if (!r) r = ss.insertSheet(REQUESTS_SHEET);
  r.getRange(1, 1, 1, 18).setValues([['Received', 'Status', 'Name', 'Mobile', 'Package', 'Price', 'Event date', 'Start', 'Hours', 'Venue', 'Delivery', 'Notes', 'Reply to send', 'Tap to text', 'Event type', 'Add-ons', 'Total', 'Location (map link)']]);
  r.getRange('G:H').setNumberFormat('@'); // Event date + Start stay plain text so Sheets does not turn "8:00 AM" into a time value
  r.setFrozenRows(1);
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false).build();
  r.getRange(2, C.STATUS, 1000, 1).setDataValidation(rule);
  r.setColumnWidth(C.REPLY, 380);
  clearCache();
}

/* ---------------- website: load dates ---------------- */
// The page asks "did anything change?" every few seconds. The answer is kept in a short-lived cache so those
// checks are instant, and the cache is cleared the moment you edit the sheet, so changes still show within seconds.
const CACHE_KEY = 'site-payload', CACHE_SECONDS = 60;
function clearCache() { try { CacheService.getScriptCache().remove(CACHE_KEY); } catch (err) {} }

function doGet(e) {
  const cache = CacheService.getScriptCache();
  let payload = null;
  try { const hit = cache.get(CACHE_KEY); if (hit) payload = JSON.parse(hit); } catch (err) {}
  if (!payload) {
    const days = getDays(), packages = getPackages(), addons = getAddons(), settings = getSettings(), used = getAddonUse(addons), pused = getPackageUse(), special = getSpecialPrices();
    // Availability by Date: a package set to No becomes "unavailable" for that package on that date; add-ons set to No go in "off"
    const avail = getAvailability(), addonTitles = (addons || []).map(function (x) { return x.title; }), off = {};
    Object.keys(avail).forEach(function (d) {
      Object.keys(avail[d]).forEach(function (t) {
        if (addonTitles.indexOf(t) > -1) { off[d] = off[d] || {}; off[d][t] = 1; }
        else { days[d] = days[d] || {}; if (!days[d]['*'] && !days[d][t]) days[d][t] = 'unavailable'; }
      });
    });
    payload = { ok: true, days: days, packages: packages, addons: addons, settings: settings, used: used, pused: pused, special: special, off: off, v: hash(JSON.stringify([days, packages, addons, settings, used, pused, special, off])) };
    try { cache.put(CACHE_KEY, JSON.stringify(payload), CACHE_SECONDS); } catch (err) {}
  }
  if (e && e.parameter && e.parameter.check) return json({ ok: true, v: payload.v }); // cheap "did anything change?" check
  return json(payload);
}

/* ---------------- Packages and Add-ons sheets (you edit these) ---------------- */
function setupContent(ss) {
  let p = ss.getSheetByName(PACKAGES_SHEET);
  if (!p) p = ss.insertSheet(PACKAGES_SHEET);
  p.getRange(1, 1, 1, 9).setValues([['Show (Yes/No)', 'Package name', 'Promo price', 'Regular price', 'Note under the price', 'Image (optional: file in images folder, web link, or Google Drive link)', 'Included items (one per line: Alt+Enter for a new line)', 'Available (Yes/No)', 'Quantity available (how many customers can pick)']]);
  p.setFrozenRows(1);
  if (p.getLastRow() < 2) {
    p.getRange(2, 1, 3, 9).setValues([
      ['Yes', 'Standard Package', 1200, 1500, 'per event, up to 22 hrs', '', ['Sony ULT Tower 7 420W Speaker', 'Xiaomi 32 inch LED TV', 'Platinum XL SD with 23k Songs', 'Platinum jukebox keypad', 'Behringer wired mic & 1 Platinum wired mic', 'Song book 23k songs'].join('\n'), 'Yes', 1],
      ['Yes', 'Pro Package', 1500, 1800, 'per event, up to 22 hrs', '', ['JBL Partybox 710 800W Speaker', 'Xiaomi 32 inch LED TV', 'Platinum XL SD with 23k Songs', 'Platinum jukebox keypad', 'JBL Partybox Wireless Mic 2', 'Song book 23k songs', '1 High Chair'].join('\n'), 'Yes', 1],
      ['Yes', 'Premium Package', 1800, 2300, 'per event, up to 22 hrs', '', ['JBL Partybox 720 800W Speaker with built-in Battery', 'Xiaomi 43 inch QLED TV 4k Display', 'Platinum Piano 4k - 4k Output Display', 'JBL Easysing Wireless Mic with AI', 'Song book 23k songs', '2 High bar Chair'].join('\n'), 'Yes', 1]
    ]);
  }
  p.getRange('C:D').setNumberFormat('#,##0');
  p.getRange('G:G').setWrap(true);
  p.setColumnWidth(2, 160); p.setColumnWidth(6, 220); p.setColumnWidth(7, 380);
  p.getRange(2, 1, 30, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
  p.getRange(2, 8, 30, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
  p.setColumnWidth(8, 130); p.setColumnWidth(9, 190);
  p.getRange('I:I').setNumberFormat('0');

  let a = ss.getSheetByName(ADDONS_SHEET);
  if (!a) a = ss.insertSheet(ADDONS_SHEET);
  a.getRange(1, 1, 1, 7).setValues([['Show (Yes/No)', 'Add-on name', 'Promo price', 'Regular price', 'Note', 'Available (Yes/No)', 'Quantity available (how many customers can pick)']]);
  a.setFrozenRows(1);
  if (a.getLastRow() < 2) {
    a.getRange(2, 1, 3, 7).setValues([
      ['Yes', 'Shure SM58 Professional Dynamic Vocal Microphone', 150, 200, '', 'Yes', 1],
      ['Yes', 'Tables', '', '', 'Rental tables for your guests', 'No', ''],
      ['Yes', 'Chairs', '', '', 'Rental chairs for your guests', 'No', '']
    ]);
  }
  a.getRange('C:D').setNumberFormat('#,##0');
  a.setColumnWidth(2, 300); a.setColumnWidth(5, 220); a.setColumnWidth(7, 190);
  a.getRange('G:G').setNumberFormat('0');
  a.getRange(2, 1, 30, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
  a.getRange(2, 6, 30, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
}

// Settings sheet: [key used by the website, label shown in the sheet, starting text, help text]
const SETTINGS_ROWS = [
  ['eyebrow', 'Header small text', 'Reserve Now!', 'The small gold text above the title.'],
  ['title', 'Header title', BUSINESS, 'The big title at the top of the page.'],
  ['subtitle', 'Header description', 'Videoke sets for birthdays, debuts, reunions and fiestas. Pick a package and date, then send your request.', 'The sentence under the title.'],
  ['image', 'Header background image', '', 'Optional. A web link, a Google Drive share link (set sharing to "Anyone with the link"), or a file in the images folder, like images/header.jpg. Leave blank for the default design.'],
  ['color', 'Header background color', '', 'Optional. A color code like #1b1740. Shown when there is no image, and behind it.'],
  ['dark', 'Image darkness (0-90)', '55', 'How dark the shade over the image is, so the white text stays easy to read. 0 = no shade, 90 = very dark.'],
  ['step1', 'Step 1 title', 'Pick a package', 'The heading of the packages section.'],
  ['step1note', 'Step 1 note', 'All prices shown are promo prices.', 'The line under the Step 1 heading.'],
  ['addonsTitle', 'Add-ons title', 'Add-ons', 'The heading above the add-ons table.'],
  ['addonsNote', 'Add-ons note', 'Tick the add-ons you want, then choose how many. It is included in your total rental fee.', 'The line under the add-ons heading.'],
  ['minDays', 'Earliest booking date (days from today)', '0', 'How soon customers can book. 0 = from today. Example: 3 = the earliest date is 3 days from today. Dates before that are greyed out.'],
  ['maxDays', 'Latest booking date (days from today)', '', 'How far ahead customers can book. Example: 30 = only the next 30 days are open. Blank = no limit (the calendar shows about 12 months).'],
  ['step2', 'Step 2 title', 'Pick a date', 'The heading of the calendar section.'],
  ['calNote', 'Calendar note', 'Each date shows how many packages are left. Pick a package above to check yours.', 'The line above the calendar, shown before a package is chosen.'],
  ['step3', 'Step 3 title', 'Your details and request', 'The heading of the form section.'],
  ['reqTitle', 'Your request title', 'Your request', 'The heading above the request summary.'],
  ['sendBtn', 'Send button text', 'Send booking request', 'The text on the send button.'],
  ['locNote', 'Location note', 'By sending this request, you allow your location to be shared with {name}.', 'The line under the send button. {name} becomes the business name from config.js.'],
  ['gcashTitle', 'GCash title', 'Pay with GCash', 'The heading of the payment section.'],
  ['gcashNote', 'GCash note', 'Already confirmed? Pay here. The reservation fee is included in your total rental fee.', 'The first line in the payment section.'],
  ['gcashHint', 'GCash hint', 'Open GCash, tap Send Money, and enter the exact amount of your total rental fee.', 'The line under the GCash number.'],
  ['remTitle', 'Reminder title', 'Reminder', 'The heading of the reminder box.'],
  ['rem1', 'Reminder 1', 'The reservation fee is included in your total rental fee. It is not an extra charge.', 'First reminder line. Type - to hide it.'],
  ['rem2', 'Reminder 2', 'A ₱1,000 security deposit is collected upon delivery. It is refundable upon pick up, as long as there is no damage and no missing equipment.', 'Second reminder line. Type - to hide it.'],
  ['rem3', 'Reminder 3', 'Delivery is free within 5 km. Beyond 5 km, it is ₱30 for every additional kilometer.', 'Third reminder line. Type - to hide it.'],
  ['rem4', 'Reminder 4', '', 'Optional extra reminder line.'],
  ['rem5', 'Reminder 5', '', 'Optional extra reminder line.'],
  ['phone', 'Phone number (call or text)', '0999 711 4240', 'Shown to customers as "You can also call or text". Change it here when your number changes. Type - to hide.'],
  ['whatsapp', 'WhatsApp number', '09997114240', 'Number for the WhatsApp button. Change it here if you get a new number. Blank = no button.'],
  ['viber', 'Viber number', '09997114240', 'Number for the Viber button. Blank = no button.'],
  ['messenger', 'Messenger page ID or username', '1368190183045322', 'Your Facebook page ID or username for the Messenger button. Blank = no button.'],
  ['gcashNumber', 'GCash number', '09997114240', 'The GCash number customers see and can copy. Blank = hides the GCash section.'],
  ['gcashName', 'GCash account name', '', 'Optional. Shown next to the GCash number.'],
  ['showWhatsapp', 'Show WhatsApp button', 'Yes', 'Choose Yes or No from the dropdown.'],
  ['showViber', 'Show Viber button', 'Yes', 'Choose Yes or No from the dropdown.'],
  ['showMessenger', 'Show Messenger button', 'Yes', 'Choose Yes or No from the dropdown.'],
  ['showGcash', 'Show GCash payment section', 'Yes', 'Choose Yes or No from the dropdown.'],
  ['showCall', 'Show call or text line', 'Yes', 'Choose Yes or No from the dropdown.'],
  ['acceptRequests', 'Accept booking requests (Send button)', 'Yes', 'Choose No to pause online booking. The Send button turns off and customers see the paused message below.'],
  ['pausedMsg', 'Paused message', 'Online booking is paused right now. Please message us to book.', 'Shown above the Send button when Accept booking requests is No.'],
  ['footer', 'Footer note', 'Your date is not reserved until we reply and confirm.', 'The small line at the very bottom of the page.'],
  ['doneTitle', 'Request sent title', 'Request sent!', 'The heading customers see after sending.'],
  ['doneMsg', 'Request sent message', 'Thank you, {customer}. We will contact you at {phone} to confirm your date for {date}.', 'The message after sending. {customer}, {phone} and {date} are filled in automatically.'],
  ['doneNote', 'Request sent note', 'Your date is not reserved until we reply and confirm.', 'The second line after sending. Type - to hide it.'],
  ['doneShot', 'Payment screenshot message', 'Please send screenshots of your payment so we can verify and confirm your reservation. Thank you!', 'The bold line after sending. Type - to hide it.']
];

const TOGGLE_KEYS = ['showWhatsapp', 'showViber', 'showMessenger', 'showGcash', 'showCall', 'acceptRequests'];

// Creates the Settings sheet, or adds any rows it is missing. Your own text is never overwritten.
function setupSettings(ss) {
  let t = ss.getSheetByName(SETTINGS_SHEET);
  if (!t) t = ss.insertSheet(SETTINGS_SHEET);
  if (t.getLastRow() < 1) t.getRange(1, 1, 1, 3).setValues([['Setting (do not rename)', 'Your text', 'What it changes']]);
  t.setFrozenRows(1);
  t.getRange('A:C').setNumberFormat('@');
  const have = {};
  if (t.getLastRow() > 1) t.getRange(2, 1, t.getLastRow() - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim().toLowerCase()] = true; });
  const add = SETTINGS_ROWS.filter(function (r) { return !have[r[1].toLowerCase()] && !have[r[0].toLowerCase()]; })
    .map(function (r) { return [r[1], r[2], r[3]]; });
  if (add.length) t.getRange(t.getLastRow() + 1, 1, add.length, 3).setValues(add);
  // The add-ons note used to say "Tick an add-on...". Update that old text once; any other text you wrote is kept.
  if (t.getLastRow() > 1) {
    t.getRange(2, 1, t.getLastRow() - 1, 2).getValues().forEach(function (r, i) {
      const lab = String(r[0]).trim().toLowerCase();
      if ((lab === 'add-ons note' || lab === 'addonsnote') && /^\s*(tick an?\s*add|choose how many of each)/i.test(String(r[1]))) {
        t.getRange(i + 2, 2).setValue('Tick the add-ons you want, then choose how many. It is included in your total rental fee.');
      }
    });
  }
  t.setColumnWidth(1, 190); t.setColumnWidth(2, 460); t.setColumnWidth(3, 340);
  t.getRange('B:B').setWrap(true);
  // Yes/No dropdown on every on/off switch row
  const yn = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build();
  const keyOf = {};
  SETTINGS_ROWS.forEach(function (r) { keyOf[r[1].toLowerCase()] = r[0]; keyOf[r[0].toLowerCase()] = r[0]; });
  if (t.getLastRow() > 1) {
    t.getRange(2, 1, t.getLastRow() - 1, 1).getValues().forEach(function (r, i) {
      if (TOGGLE_KEYS.indexOf(keyOf[String(r[0]).trim().toLowerCase()]) > -1) t.getRange(i + 2, 2).setDataValidation(yn);
    });
  }
}

// Reads the Settings sheet and gives the website {key: text}. Works with the labels in column A.
function getSettings() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  if (!sh || sh.getLastRow() < 2) return null; // the website then uses its built-in text
  const keyOf = {};
  SETTINGS_ROWS.forEach(function (r) { keyOf[r[1].toLowerCase()] = r[0]; keyOf[r[0].toLowerCase()] = r[0]; });
  const out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    const k = keyOf[String(r[0] || '').trim().toLowerCase()];
    if (k) out[k] = String(r[1] == null ? '' : r[1]).trim();
  });
  return out;
}

function money(v) {
  const n = Number(String(v == null ? '' : v).replace(/[^\d.]/g, ''));
  return n ? String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '';
}

function getPackages() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PACKAGES_SHEET);
  if (!sh || sh.getLastRow() < 2) return null; // the website then uses config.js
  const out = [];
  sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(function (r) {
    const title = String(r[1] || '').trim();
    if (!title || String(r[0]).trim().toLowerCase() === 'no') return;
    out.push({
      title: title, price: money(r[2]), regular: money(r[3]), note: String(r[4] || '').trim(),
      available: String(r[7]).trim().toLowerCase() !== 'no', // blank = Yes; No = shown but greyed out and not tappable
      qty: Math.max(1, Math.min(50, Math.floor(Number(r[8])) || 1)), // how many can be booked on the same date; blank = 1
      image: String(r[5] || '').trim(),
      items: String(r[6] || '').split(/\r?\n/).map(function (s) { return s.trim(); }).filter(Boolean)
    });
  });
  return out.length ? out : null;
}

function getAddons() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADDONS_SHEET);
  if (!sh) return null;
  const out = [];
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().forEach(function (r) {
      const title = String(r[1] || '').trim();
      if (!title || String(r[0]).trim().toLowerCase() === 'no') return;
      const qty = Math.max(1, Math.min(50, Math.floor(Number(r[6])) || 1)); // blank = 1
      out.push({
        title: title, price: money(r[2]), regular: money(r[3]), note: String(r[4] || '').trim(), qty: qty,
        available: String(r[5]).trim().toLowerCase() !== 'no' && !!money(r[2])
      });
    });
  }
  return out;
}

/* ---------------- Availability by Date sheet (1 row per date, 1 column per package / add-on) ---------------- */
function setupAvailability(ss) {
  let sh = ss.getSheetByName(AVAIL_SHEET);
  if (!sh) sh = ss.insertSheet(AVAIL_SHEET);
  const pk = (getPackages() || []).map(function (x) { return x.title; });
  const ad = (getAddons() || []).map(function (x) { return x.title; });
  const have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (v) { return String(v || '').trim(); }) : [];
  if (!have[0]) { have[0] = 'Date (YYYY-MM-DD)'; sh.getRange(1, 1).setValue(have[0]); }
  // Add a column for every package and add-on that has none yet. Existing columns and your Yes/No choices are never touched.
  pk.concat(ad).forEach(function (t) {
    if (have.indexOf(t) < 0) { have.push(t); sh.getRange(1, have.length).setValue(t); }
  });
  const n = have.length;
  sh.setFrozenRows(1); sh.setFrozenColumns(1);
  sh.getRange('A:A').setNumberFormat('@');
  sh.setColumnWidth(1, 150);
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setWrap(true);
  sh.getRange('A1').setNote('One row per date, like 2026-12-25. Under each package or add-on, pick No to turn it off on that date. Blank or Yes = available as usual.');
  if (n > 1) {
    const yn = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).setAllowInvalid(false).build();
    for (let c = 2; c <= n; c++) {
      sh.getRange(2, c, 500, 1).setDataValidation(yn);
      sh.setColumnWidth(c, 140);
      sh.getRange(1, c).setBackground(pk.indexOf(have[c - 1]) > -1 ? '#dbeafe' : ad.indexOf(have[c - 1]) > -1 ? '#fef3c7' : '#eeeeee');
    }
    sh.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No').setBackground('#f4c7c3').setFontColor('#a61c00').setRanges([sh.getRange(2, 2, 500, n - 1)]).build()]);
  }
}

// { 'YYYY-MM-DD': { 'Package or add-on title': true } } for every cell set to No
function getAvailability() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(AVAIL_SHEET);
  const out = {};
  if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() < 2) return out;
  const vals = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  const heads = vals[0].map(function (v) { return String(v || '').trim(); });
  for (let i = 1; i < vals.length; i++) {
    const d = normDate(vals[i][0]);
    if (!d) continue;
    for (let c = 1; c < heads.length; c++) {
      if (heads[c] && String(vals[i][c]).trim().toLowerCase() === 'no') { out[d] = out[d] || {}; out[d][heads[c]] = true; }
    }
  }
  return out;
}

/* ---------------- Special Prices sheet (a different price on certain dates) ---------------- */
function setupSpecial(ss) {
  let sp = ss.getSheetByName(SPECIAL_SHEET);
  if (!sp) sp = ss.insertSheet(SPECIAL_SHEET);
  sp.getRange(1, 1, 1, 6).setValues([['From date (YYYY-MM-DD)', 'To date (optional, for a range)', 'Package (blank = all packages)', 'Price for these dates (PHP)', 'OR add / subtract (PHP, example 300 or -200)', 'Label shown to customers (example: Holiday rate)']]);
  sp.setFrozenRows(1);
  sp.getRange('A:B').setNumberFormat('@');
  sp.getRange('D:E').setNumberFormat('#,##0');
  sp.getRange(2, 3, 300, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName(PACKAGES_SHEET).getRange('B2:B30'), true).setAllowInvalid(false).build());
  sp.getRange('A1').setNote('Example: 2026-12-25. For one day, leave the To date blank.');
  sp.getRange('D1').setNote('Fills in the exact price for that package on those dates. If this is filled, the add/subtract column is ignored.');
  sp.getRange('E1').setNote('Use this instead to raise or lower the normal price. Example: 300 makes it 300 more, -200 makes it 200 less. Handy with "all packages".');
  sp.getRange('F1').setNote('The small tag customers see beside the price on those dates. Blank shows "Special rate".');
  sp.setColumnWidth(1, 160); sp.setColumnWidth(2, 190); sp.setColumnWidth(3, 190); sp.setColumnWidth(4, 190); sp.setColumnWidth(5, 230); sp.setColumnWidth(6, 260);
  sp.getRange(1, 1, 1, 6).setWrap(true);
}

// [{from, to, pkg ('' = all), price (number or null), add (number or null), label}]
function getSpecialPrices() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SPECIAL_SHEET);
  const out = [];
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(function (r) {
    const from = normDate(r[0]);
    if (!from) return;
    let to = normDate(r[1]) || from;
    if (to < from) to = from;
    const price = Math.round(Number(String(r[3] == null ? '' : r[3]).replace(/[^\d.]/g, ''))) || null;
    const addRaw = String(r[4] == null ? '' : r[4]).replace(/[^\d.\-]/g, '');
    const add = addRaw !== '' && !isNaN(Number(addRaw)) ? Math.round(Number(addRaw)) : null;
    if (price === null && add === null) return;
    out.push({ from: from, to: to, pkg: String(r[2] || '').trim(), price: price, add: add, label: String(r[5] || '').trim() });
  });
  return out;
}

// The rule that applies to a package on a date: a rule for that package beats an "all packages" rule; a lower row beats a higher one.
function specialFor(rules, title, date) {
  let best = null, bestScore = -1;
  rules.forEach(function (r) {
    if (date < r.from || date > r.to) return;
    if (r.pkg && r.pkg !== title) return;
    const score = r.pkg ? 1 : 0;
    if (score >= bestScore) { best = r; bestScore = score; }
  });
  return best;
}

// Add-ons already taken on each date, so the website can show what is left.
// Counts requests marked Accepted or Hold. (Add 'new' to the list to also count requests you have not answered yet.)
const STOCK_STATUSES = ['accepted', 'hold'];
function getAddonUse(addons) {
  const out = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_SHEET);
  if (!sh || sh.getLastRow() < 2 || !addons || !addons.length) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, C.ADDONS).getValues().forEach(function (r) {
    if (STOCK_STATUSES.indexOf(String(r[C.STATUS - 1]).trim().toLowerCase()) < 0) return;
    const date = normDate(r[C.DATE - 1]), text = String(r[C.ADDONS - 1] || '').trim();
    if (!date || !text) return;
    addons.forEach(function (a) {
      // the Add-ons column looks like: "Shure SM58 ... x2, Tables x3" (no x means 1)
      const m = text.match(new RegExp('(?:^|, )' + a.title.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?: x(\\d+))?(?=, |$)'));
      if (!m) return;
      out[date] = out[date] || {};
      out[date][a.title] = (out[date][a.title] || 0) + (m[1] ? parseInt(m[1], 10) : 1);
    });
  });
  return out;
}

// Packages already taken on each date: { 'YYYY-MM-DD': { 'Package title': count } }. Counts the same statuses as add-ons.
function getPackageUse() {
  const out = {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_SHEET);
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, C.DATE).getValues().forEach(function (r) {
    if (STOCK_STATUSES.indexOf(String(r[C.STATUS - 1]).trim().toLowerCase()) < 0) return;
    const date = normDate(r[C.DATE - 1]), pk = String(r[C.PKG - 1]).trim();
    if (!date || !pk) return;
    out[date] = out[date] || {};
    out[date][pk] = (out[date][pk] || 0) + 1;
  });
  return out;
}

// How many of a package can be booked on one date (the Quantity column on the Packages sheet; 1 if blank or not found).
function packageCap(title) {
  const list = getPackages() || [];
  const m = list.filter(function (x) { return x.title === title; })[0];
  return m ? m.qty : 1;
}

function getDays() {
  // { 'YYYY-MM-DD': { '*': 'unavailable' }  or  { 'Standard Package': 'booked', 'Pro Package': 'hold' } }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATES_SHEET);
  const days = {};
  if (sheet && sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    rows.forEach(function (row) {
      const d = normDate(row[0]);
      const status = String(row[1] || '').trim().toLowerCase();
      const pk = String(row[3] || '').trim() || '*';
      if (!d || !(status === 'booked' || status === 'hold' || status === 'unavailable')) return;
      days[d] = days[d] || {};
      if (days[d][pk] !== 'booked') days[d][pk] = status;
    });
  }
  return days;
}

function hash(s) {
  let h = 0;
  for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return String(h);
}

/* ---------------- website: save a request ---------------- */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = JSON.parse(e.postData.contents);

    if (p.website) return json({ ok: true }); // honeypot: bots fill this, humans don't

    const name = clean(p.name), phone = clean(p.phone), date = clean(p.date);
    if (!name || !phone || !/^\d{4}-\d{2}-\d{2}$/.test(date) || !clean(p.pkg)) {
      return json({ ok: false, error: 'invalid' });
    }

    // Booking window (Settings sheet): refuse dates that are too soon or too far ahead
    const cfgS = getSettings() || {};
    const dayKey = function (n) { const d = new Date(); d.setDate(d.getDate() + n); return Utilities.formatDate(d, SITE_TZ, 'yyyy-MM-dd'); };
    const minD = Math.max(0, parseInt(cfgS.minDays, 10) || 0), maxD = parseInt(cfgS.maxDays, 10);
    if (date < dayKey(minD) || (!isNaN(maxD) && maxD >= 0 && date > dayKey(maxD))) return json({ ok: false, error: 'closed' });

    // Price check: refuse if the price the customer saw is no longer the right price for that date
    const pkRow = (getPackages() || []).filter(function (x) { return x.title === clean(p.pkg); })[0];
    if (pkRow) {
      const base = Number(String(pkRow.price).replace(/[^\d]/g, '')) || 0;
      const sp = specialFor(getSpecialPrices(), pkRow.title, date);
      const expect = sp ? (sp.price !== null ? sp.price : Math.max(0, base + (sp.add || 0))) : base;
      const sent = Number(String(p.price).replace(/[^\d]/g, '')) || 0;
      if (sent !== expect) return json({ ok: false, error: 'price', expected: expect });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dSheet = ss.getSheetByName(DATES_SHEET);
    if (dSheet && dSheet.getLastRow() > 1) {
      const rows = dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < rows.length; i++) {
        const st = String(rows[i][1] || '').trim().toLowerCase();
        const rowPkg = String(rows[i][3] || '').trim();
        if (normDate(rows[i][0]) === date && (st === 'booked' || st === 'hold' || st === 'unavailable') && (!rowPkg || rowPkg === clean(p.pkg))) {
          return json({ ok: false, error: 'taken' });
        }
      }
    }

    // Availability by Date: refuse a package or add-on that is switched off for this date
    const av = getAvailability()[date] || {};
    if (av[clean(p.pkg)]) return json({ ok: false, error: 'taken' });

    // Package stock: refuse if every set of this package is already taken for this date
    const pkgUse = (getPackageUse()[date] || {})[clean(p.pkg)] || 0;
    if (pkgUse >= packageCap(clean(p.pkg))) return json({ ok: false, error: 'taken' });

    // Add-on stock: refuse if someone already took the last ones for this date
    if (p.addonQty && typeof p.addonQty === 'object') {
      const list = getAddons() || [], use = (getAddonUse(list)[date]) || {};
      for (const t in p.addonQty) {
        const a = list.filter(function (x) { return x.title === t; })[0], q = Math.floor(Number(p.addonQty[t])) || 0;
        if (q < 1) continue;
        if (av[t]) return json({ ok: false, error: 'addon', item: t });
        if (!a || !a.available || q > a.qty - (use[t] || 0)) return json({ ok: false, error: 'addon', item: t });
      }
    }

    // Customer location (they allowed it in the browser): saved as a Google Maps link
    let loc = '';
    if (p.lat != null && p.lng != null && p.lat !== '' && p.lng !== '') {
      const lat = Number(p.lat), lng = Number(p.lng);
      if (isFinite(lat) && isFinite(lng) && Math.abs(lat) <= 90 && Math.abs(lng) <= 180) loc = 'https://www.google.com/maps?q=' + lat.toFixed(6) + ',' + lng.toFixed(6);
    }

    const rSheet = ss.getSheetByName(REQUESTS_SHEET);
    rSheet.appendRow([
      new Date(), 'New', name, phone, clean(p.pkg), clean(p.price),
      date, clean(p.time), clean(p.hours), clean(p.venue),
      p.delivery ? 'Yes' : 'No (pickup)', clean(p.notes), '', '', clean(p.events), clean(p.addons), clean(p.total), loc
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail(
        NOTIFY_EMAIL,
        'New booking request: ' + date + ' - ' + name,
        'Name: ' + name + '\nMobile: ' + phone + '\nPackage: ' + clean(p.pkg) +
        '\nDate: ' + date + '\nStart: ' + clean(p.time) + ' (' + clean(p.hours) + ' hours)' +
        '\nVenue: ' + clean(p.venue) + '\nDelivery: ' + (p.delivery ? 'Yes' : 'No') +
        '\nEvent: ' + clean(p.events) + '\nAdd-ons: ' + clean(p.addons) + '\nTotal: ' + clean(p.total) + '\nLocation: ' + (loc || 'not shared') + '\nNotes: ' + clean(p.notes)
      );
    }
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------- you change a Status: update calendar + write reply ---------------- */
// Runs by itself whenever you edit the sheet (a "simple trigger", no setup needed).
function onEdit(e) {
  if (!e || !e.range) return;
  clearCache(); // any edit in the spreadsheet: the website should see it on its next check
  const sheet = e.range.getSheet();
  if (sheet.getName() !== REQUESTS_SHEET) return;
  if (e.range.getColumn() > C.STATUS || e.range.getLastColumn() < C.STATUS) return;

  const ss = e.source;
  const first = Math.max(2, e.range.getRow());
  const last = e.range.getLastRow();
  const dates = {};

  for (let row = first; row <= last; row++) {
    const v = sheet.getRange(row, 1, 1, 17).getValues()[0];
    const date = normDate(v[C.DATE - 1]);
    if (!date) continue;
    dates[date] = true;
    writeReply(sheet, row, v, date);
  }
  Object.keys(dates).forEach(function (d) { syncDate(ss, d); });
  clearCache();
}

// Work out the date's status from all requests, then update the "Dates" sheet.
function syncDate(ss, date) {
  const rSheet = ss.getSheetByName(REQUESTS_SHEET);
  const want = {}; // package -> 'booked' | 'hold'
  const acc = {}, hld = {}; // package -> how many accepted / on hold for this date
  if (rSheet.getLastRow() > 1) {
    const rows = rSheet.getRange(2, 1, rSheet.getLastRow() - 1, 7).getValues();
    rows.forEach(function (r) {
      if (normDate(r[C.DATE - 1]) !== date) return;
      const pk = String(r[C.PKG - 1]).trim();
      const st = String(r[C.STATUS - 1]).trim().toLowerCase();
      if (!pk) return;
      if (st === 'accepted') acc[pk] = (acc[pk] || 0) + 1;
      else if (st === 'hold') hld[pk] = (hld[pk] || 0) + 1;
    });
  }
  // The date is only marked for a package once ALL its sets are taken (Quantity on the Packages sheet; 1 = as before)
  Object.keys(acc).concat(Object.keys(hld)).forEach(function (pk) {
    if (want[pk]) return;
    const a = acc[pk] || 0, total = a + (hld[pk] || 0), cap = packageCap(pk);
    if (total >= cap) want[pk] = a >= cap ? 'booked' : 'hold';
  });

  const dSheet = ss.getSheetByName(DATES_SHEET);
  const n = dSheet.getLastRow() - 1;
  const rows = n > 0 ? dSheet.getRange(2, 1, n, 4).getValues() : [];
  const typed = {}; // rows you typed yourself are never touched
  rows.forEach(function (r) {
    if (normDate(r[0]) === date && String(r[2]).trim() !== 'request') typed[String(r[3]).trim()] = true;
  });
  const have = {};
  for (let i = rows.length - 1; i >= 0; i--) { // bottom-up so deleting keeps row numbers right
    if (normDate(rows[i][0]) !== date || String(rows[i][2]).trim() !== 'request') continue;
    const pk = String(rows[i][3]).trim();
    if (want[pk] && !have[pk]) {
      have[pk] = true;
      if (String(rows[i][1]).trim() !== want[pk]) dSheet.getRange(i + 2, 2).setValue(want[pk]);
    } else {
      dSheet.deleteRow(i + 2);
    }
  }
  Object.keys(want).forEach(function (pk) {
    if (!have[pk] && !typed[pk] && !typed['']) dSheet.appendRow([date, want[pk], 'request', pk]);
  });
}

// Run once after upgrading: rebuilds the auto rows in "Dates" from your existing requests.
function rebuildDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const r = ss.getSheetByName(REQUESTS_SHEET);
  const seen = {};
  if (r.getLastRow() > 1) {
    r.getRange(2, C.DATE, r.getLastRow() - 1, 1).getValues().forEach(function (v) {
      const d = normDate(v[0]);
      if (d) seen[d] = true;
    });
  }
  Object.keys(seen).forEach(function (d) { syncDate(ss, d); });
  clearCache();
}

function writeReply(sheet, row, v, date) {
  const status = String(v[C.STATUS - 1]).trim().toLowerCase();
  const name = String(v[C.NAME - 1]).trim();
  const nice = prettyDate(date);
  let msg = '';

  if (status === 'accepted') {
    msg = 'Hi ' + name + '! Your booking is CONFIRMED.\n' +
      'Package: ' + v[C.PKG - 1] + ' (' + v[C.PRICE - 1] + ')\n' +
      (v[C.ADDONS - 1] ? 'Add-ons: ' + v[C.ADDONS - 1] + '\n' : '') +
      (v[C.TOTAL - 1] ? 'Total rental fee: ' + v[C.TOTAL - 1] + ' (reservation fee included)\n' : '') +
      'Date: ' + nice + ', ' + timeText(v[C.TIME - 1]) + ' for ' + v[C.HOURS - 1] + ' hours\n' +
      'Venue: ' + v[C.VENUE - 1] + '\n' +
      REMINDER + '\n' +
      'Thank you for choosing ' + BUSINESS + '!';
  } else if (status === 'hold') {
    msg = 'Hi ' + name + '! We are holding ' + nice + ' for you (' + v[C.PKG - 1] + '). ' +
      'Please reply to confirm so we can lock in your booking. Thank you! - ' + BUSINESS;
  } else if (status === 'declined') {
    msg = 'Hi ' + name + '! Sorry, we are not available on ' + nice + '. ' +
      'You are welcome to book another date. Thank you! - ' + BUSINESS;
  }

  const replyCell = sheet.getRange(row, C.REPLY);
  const smsCell = sheet.getRange(row, C.SMS);
  if (!msg) { replyCell.clearContent(); smsCell.clearContent(); return; }

  replyCell.setValue(msg).setWrap(true);
  const phone = String(v[C.PHONE - 1]).replace(/[^\d+]/g, '');
  if (phone) {
    const url = 'sms:' + phone + '?body=' + encodeURIComponent(msg);
    smsCell.setFormula('=HYPERLINK("' + url + '","Text them")');
  } else {
    smsCell.clearContent();
  }
}

/* ---------------- helpers ---------------- */
// Turns a Date or text into 'YYYY-MM-DD' (or '' if it is not a date).
function normDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v == null ? '' : v).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

// Shows the Start cell as "8:00 AM" even if Sheets stored it as a time value (which prints as Dec 30 1899).
function timeText(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'h:mm a');
  return String(v == null ? '' : v).trim();
}

// Run ONCE to repair old rows: turns the Start column into plain text like "8:00 AM".
function fixStartTimes() {
  const r = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUESTS_SHEET);
  const n = r.getLastRow() - 1;
  if (n < 1) return;
  const rng = r.getRange(2, C.TIME, n, 1);
  const vals = rng.getValues().map(function (row) { return [timeText(row[0])]; });
  rng.setNumberFormat('@').setValues(vals);
}

function prettyDate(s) {
  const p = s.split('-');
  return Utilities.formatDate(new Date(+p[0], +p[1] - 1, +p[2]), Session.getScriptTimeZone(), 'EEEE, MMMM d, yyyy');
}

// Trim, limit length, and stop spreadsheet formulas from being injected.
function clean(v) {
  let s = String(v == null ? '' : v).trim().slice(0, 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
