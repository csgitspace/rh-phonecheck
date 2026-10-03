/**
 * Roy-Hart HS — First Period Phone Check
 * Standalone Google Apps Script web app.
 *
 * This script owns nothing but itself. On first run it creates its own data
 * spreadsheet; the original collection sheet is optional and only ever read,
 * so view-only access to it is enough.
 *
 * Script Properties:
 *   SHEET_ID         the data spreadsheet this app creates and writes to
 *   SOURCE_SHEET_ID  optional. A collection sheet to import teacher tabs from.
 *
 * Sheets in the data spreadsheet:
 *   Settings   Setting | Value | Notes
 *   Teachers   Email | Name | Section | Room | Role | Active
 *   Roster     Student Name | Student ID | Grade | Section | Active
 *   Statuses   Status | Short Label | Category | Color | Order | Active
 *   Log        Key | Timestamp | Date | Section | Teacher Email | Teacher Name |
 *              Student ID | Student Name | Status | Note
 *
 * Deploy: Execute as ME, access = anyone at royhart.org.
 */

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

var SHEETS = {
  SETTINGS: 'Settings',
  TEACHERS: 'Teachers',
  ROSTER:   'Roster',
  STATUSES: 'Statuses',
  LOG:      'Log',
  GRID:     'Grid View'
};

var SYSTEM_SHEETS = [
  SHEETS.SETTINGS, SHEETS.TEACHERS, SHEETS.ROSTER,
  SHEETS.STATUSES, SHEETS.LOG, SHEETS.GRID
];

var TEACHER_HEADERS = ['Email', 'Name', 'Section', 'Room', 'Role', 'Active'];
var ROSTER_HEADERS  = ['Student Name', 'Student ID', 'Grade', 'Section', 'Active'];
var LOG_HEADERS = [
  'Key', 'Timestamp', 'Date', 'Section', 'Teacher Email', 'Teacher Name',
  'Student ID', 'Student Name', 'Status', 'Note'
];

var PROP_DATA_ID = 'SHEET_ID';
var PROP_SOURCE_ID = 'SOURCE_SHEET_ID';

var DEFAULT_SETTINGS = [
  ['SCHOOL_NAME',           'Royalton-Hartland High School', 'Shown in the app header.'],
  ['PERIOD_LABEL',          'First Period',                  'Which period does the check.'],
  ['ADMIN_EMAILS',          '',                              'Comma-separated. These people see every section, all student detail, and the admin console.'],
  ['QUICK_FILL_STATUS',     'Does not have a phone',         'What the "fill the rest of the class" button applies. Must match a row on the Statuses tab.'],
  ['ALLOW_TEACHER_ROSTER_EDITS', 'TRUE',                     'TRUE lets teachers add and remove students on their own roster.'],
  ['STUDENT_DETAIL_FOR_ALL','FALSE',                         'TRUE lets every signed-in staff member see student-level history. FALSE limits it to admins and each teacher\u2019s own students.'],
  ['LOCK_TIME',             '',                              'Optional. e.g. 10:30 \u2014 after this time teachers can no longer change today\u2019s entries. Blank = no lock.'],
  ['REMINDER_TIME',         '9',                             'Hour (0-23) for the missing-submission reminder email.'],
  ['DIGEST_RECIPIENTS',     '',                              'Comma-separated addresses for the daily admin digest. Blank = no digest.'],
  ['SKIP_DATES',            '',                              'Comma-separated yyyy-MM-dd dates with no school. Excluded from reminders and expected-day math.'],
  ['LOGO_URL',              'https://files.smartsites.parentsquare.com/4898/img_pd_102121_mxvilf.png', 'Header logo.']
];

var DEFAULT_STATUSES = [
  ['Phone is in classroom lockbox',  'Lockbox',         'Compliant', '#63419A', 1, true],
  ['Phone is not in the building',   'Not in building', 'Compliant', '#4A2F78', 2, true],
  ['Does not have a phone',          'No phone',        'Compliant', '#AC9AC9', 3, true],
  ['Phone kept by student',          'Kept phone',      'Attention', '#B23A2E', 4, true],
  ['Absent from first period',       'Absent',          'Neutral',   '#636263', 5, true]
];

// ---------------------------------------------------------------------------
// Spreadsheet access
// ---------------------------------------------------------------------------

function props_() { return PropertiesService.getScriptProperties(); }

/**
 * Per-execution memo. Globals live for the length of one request, so anything
 * read more than once in a single call is fetched from Sheets exactly once.
 */
var MEMO_ = {};
function memo_(key, fn) {
  if (MEMO_[key] === undefined) MEMO_[key] = fn();
  return MEMO_[key];
}

function isReady_() { return !!props_().getProperty(PROP_DATA_ID); }

function getSS_() {
  return memo_('ss', function () {
    var id = props_().getProperty(PROP_DATA_ID);
    if (!id) throw new Error('Phone Check has not been set up yet. Open the admin console and run setup.');
    try {
      return SpreadsheetApp.openById(id);
    } catch (err) {
      throw new Error('The data spreadsheet (' + id + ') could not be opened. It may have been deleted. ' +
                      'Clear the SHEET_ID script property to build a fresh one.');
    }
  });
}

function getDataUrl_() {
  var id = props_().getProperty(PROP_DATA_ID);
  return id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '';
}

/** The original collection sheet. Opened read-only; view access is enough. */
function sourceSS_() {
  var id = props_().getProperty(PROP_SOURCE_ID);
  if (!id) throw new Error('No collection sheet is linked yet. Paste its link in the admin console first.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  return memo_('sh:' + name, function () {
    var sh = getSS_().getSheetByName(name);
    if (!sh) throw new Error('Missing sheet: ' + name + '. Run setup again from the admin console.');
    return sh;
  });
}

function tz_() {
  return memo_('tz', function () { return getSS_().getSpreadsheetTimeZone() || 'America/New_York'; });
}
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function fmtDate_(d) { return Utilities.formatDate(d, tz_(), 'yyyy-MM-dd'); }

function ui_() {
  try { return SpreadsheetApp.getUi(); } catch (err) { return null; }
}

function readObjects_(name) {
  return memo_('rows:' + name, function () { return readObjectsUncached_(name); });
}

function readObjectsUncached_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (row.join('') === '') continue;
    var obj = { _row: r + 1 };
    for (var c = 0; c < headers.length; c++) obj[headers[c]] = row[c];
    out.push(obj);
  }
  return out;
}

function truthy_(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === 'yes' || s === 'y' || s === 'x' || s === '1' || s === '\u2713';
}

function activeFlag_(v) { return (v === '' || v === null || v === undefined) ? true : truthy_(v); }

// ---------------------------------------------------------------------------
// Settings
// ---------------------------------------------------------------------------

function getSettings_() {
  var cache = CacheService.getScriptCache().get('settings');
  if (cache) return JSON.parse(cache);
  var map = {};
  DEFAULT_SETTINGS.forEach(function (r) { map[r[0]] = r[1]; });
  readObjects_(SHEETS.SETTINGS).forEach(function (r) {
    if (r['Setting']) map[String(r['Setting']).trim()] = String(r['Value']).trim();
  });
  CacheService.getScriptCache().put('settings', JSON.stringify(map), 300);
  return map;
}

function clearCaches_() {
  CacheService.getScriptCache().removeAll(['settings', 'statuses', 'teachers', 'roster']);
  MEMO_ = {};
  bumpReportVersion_();
}

/**
 * Report payloads are cached per range. There is no wildcard delete in
 * CacheService, so the cache key carries a version that changes whenever the
 * underlying data does; stale entries simply expire.
 */
function reportVersion_() {
  return props_().getProperty('REPORT_V') || '1';
}
function bumpReportVersion_() {
  try { props_().setProperty('REPORT_V', String(Date.now())); } catch (err) {}
}

/**
 * Teachers and Roster are read on nearly every request but change rarely, so
 * they are kept in CacheService as plain rows. Write paths use readObjects_
 * directly, since those need the underlying row numbers.
 */
function cachedRows_(key, sheetName, headers) {
  return memo_('cached:' + key, function () {
    var cached = CacheService.getScriptCache().get(key);
    if (cached) {
      try { return JSON.parse(cached); } catch (err) {}
    }
    var rows = readObjects_(sheetName).map(function (r) {
      var o = {};
      headers.forEach(function (h) { o[h] = (r[h] === undefined || r[h] === null) ? '' : r[h]; });
      return o;
    });
    try {
      var json = JSON.stringify(rows);
      if (json.length < 95000) CacheService.getScriptCache().put(key, json, 900);
    } catch (err2) {}
    return rows;
  });
}

function teachers_() { return cachedRows_('teachers', SHEETS.TEACHERS, TEACHER_HEADERS); }
function roster_()   { return cachedRows_('roster',   SHEETS.ROSTER,   ROSTER_HEADERS); }

// ---------------------------------------------------------------------------
// Identity
// ---------------------------------------------------------------------------

function currentEmail_() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch (err) {}
  if (!email) { try { email = Session.getEffectiveUser().getEmail() || ''; } catch (err2) {} }
  return String(email).trim().toLowerCase();
}

/** True for whoever deployed the app. Always an admin, so setup can bootstrap. */
function isOwner_() {
  try {
    return currentEmail_() === String(Session.getEffectiveUser().getEmail()).trim().toLowerCase();
  } catch (err) { return false; }
}

function identity_() {
  var email = currentEmail_();
  if (!email) throw new Error('Google could not confirm who you are. Sign in with your royhart.org account and reload.');

  var owner = isOwner_();
  if (!isReady_()) {
    return { email: email, name: email.split('@')[0], isAdmin: owner, isOwner: owner, sections: [], isKnown: owner };
  }

  var settings = getSettings_();
  var adminList = String(settings.ADMIN_EMAILS || '')
    .split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(String);

  var rows = teachers_().filter(function (t) {
    return String(t['Email']).trim().toLowerCase() === email && activeFlag_(t['Active']);
  });

  var isAdmin = owner || adminList.indexOf(email) > -1 ||
    rows.some(function (t) { return String(t['Role']).trim().toLowerCase() === 'admin'; });

  var name = rows.length ? String(rows[0]['Name']).trim() : email.split('@')[0];
  var sections = rows.map(function (t) { return String(t['Section']).trim(); }).filter(String);
  if (isAdmin) sections = allSections_();

  return {
    email: email,
    name: name || email,
    isAdmin: isAdmin,
    isOwner: owner,
    sections: sections,
    isKnown: rows.length > 0 || isAdmin
  };
}

function assertAdmin_() {
  var me = identity_();
  if (!me.isAdmin) throw new Error('That action is limited to administrators.');
  return me;
}

function allSections_() {
  var seen = {};
  teachers_().forEach(function (t) {
    var s = String(t['Section']).trim();
    if (s && activeFlag_(t['Active'])) seen[s] = true;
  });
  roster_().forEach(function (r) {
    var s = String(r['Section']).trim();
    if (s) seen[s] = true;
  });
  return Object.keys(seen).sort();
}

// ---------------------------------------------------------------------------
// Web app entry points
// ---------------------------------------------------------------------------

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page ? String(e.parameter.page) : '').toLowerCase();
  if (!isReady_()) page = 'admin';

  var file = 'CheckIn';
  if (page === 'report' || page === 'trends') file = 'Report';
  else if (page === 'admin' || page === 'setup') file = 'Admin';

  var t = HtmlService.createTemplateFromFile(file);
  t.appUrl = webAppUrl_();
  return t.evaluate()
    .setTitle('Phone Check \u2014 Roy-Hart')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function webAppUrl_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (err) { return ''; }
}

// ---------------------------------------------------------------------------
// Statuses
// ---------------------------------------------------------------------------

function getStatuses_() {
  var cached = CacheService.getScriptCache().get('statuses');
  if (cached) return JSON.parse(cached);
  var rows = readObjects_(SHEETS.STATUSES)
    .filter(function (r) { return r['Status'] && activeFlag_(r['Active']); })
    .map(function (r) {
      return {
        name:     String(r['Status']).trim(),
        short:    String(r['Short Label'] || r['Status']).trim(),
        category: String(r['Category'] || 'Compliant').trim(),
        color:    String(r['Color'] || '#63419A').trim(),
        order:    Number(r['Order'] || 99)
      };
    })
    .sort(function (a, b) { return a.order - b.order; });
  CacheService.getScriptCache().put('statuses', JSON.stringify(rows), 300);
  return rows;
}

// ---------------------------------------------------------------------------
// Log helpers
// ---------------------------------------------------------------------------

function logKey_(date, section, studentKey) { return date + '|' + section + '|' + studentKey; }

function readLog_() {
  return memo_('log:full', function () {
    var sh = sheet_(SHEETS.LOG);
    var last = sh.getLastRow();
    if (last < 2) return [];
    return mapLogRows_(sh.getRange(2, 1, last - 1, LOG_HEADERS.length).getValues(), 2);
  });
}

/**
 * Reads only the end of the log. Rows are appended in date order, so the last
 * few school days sit at the bottom of the sheet — no reason to pull 90,000
 * rows to answer "what did this class report today?". The window widens until
 * it reaches back past sinceDate, and falls back to a full read if it has to.
 */
function readLogTail_(sinceDate) {
  if (!sinceDate) return readLog_();
  return memo_('log:tail:' + sinceDate, function () {
    var sh = sheet_(SHEETS.LOG);
    var last = sh.getLastRow();
    if (last < 2) return [];

    var span = 3000;
    while (true) {
      var start = Math.max(2, last - span + 1);
      var rows = mapLogRows_(sh.getRange(start, 1, last - start + 1, LOG_HEADERS.length).getValues(), start);
      var earliest = '';
      for (var i = 0; i < rows.length; i++) {
        if (!earliest || rows[i].date < earliest) earliest = rows[i].date;
      }
      if (start === 2 || (earliest && earliest <= sinceDate)) return rows;
      span *= 4;
    }
  });
}

function mapLogRows_(values, startRow) {
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    out.push({
      row: startRow + i,
      key: String(r[0]),
      timestamp: r[1],
      date: (r[2] instanceof Date) ? fmtDate_(r[2]) : String(r[2]).trim(),
      section: String(r[3]).trim(),
      teacherEmail: String(r[4]).trim(),
      teacherName: String(r[5]).trim(),
      studentId: String(r[6]).trim(),
      studentName: String(r[7]).trim(),
      status: String(r[8]).trim(),
      note: String(r[9] || '').trim()
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Client API — check-in page
// ---------------------------------------------------------------------------

function getBootstrap() {
  var me = identity_();
  var settings = getSettings_();
  var statuses = getStatuses_();

  var quick = String(settings.QUICK_FILL_STATUS || '').trim();
  var quickStatus = statuses.filter(function (s) { return s.name === quick; })[0] || statuses[0] || null;

  return {
    user: me,
    today: today_(),
    statuses: statuses,
    quickFill: quickStatus ? { name: quickStatus.name, short: quickStatus.short } : null,
    schoolName: settings.SCHOOL_NAME,
    periodLabel: settings.PERIOD_LABEL,
    logoUrl: settings.LOGO_URL,
    allowRosterEdits: truthy_(settings.ALLOW_TEACHER_ROSTER_EDITS) || me.isAdmin,
    appUrl: webAppUrl_(),
    allSections: me.isAdmin ? allSections_() : me.sections
  };
}

function getSectionDay(section, dateStr) {
  var me = identity_();
  section = String(section || '').trim();
  dateStr = String(dateStr || today_()).trim();
  assertSectionAccess_(me, section);

  var roster = roster_()
    .filter(function (r) {
      return String(r['Section']).trim() === section &&
             String(r['Student Name']).trim() && activeFlag_(r['Active']);
    })
    .sort(function (a, b) {
      return String(a['Student Name']).localeCompare(String(b['Student Name']));
    });

  var existing = {}, prior = {};
  var yesterday = previousSchoolDay_(dateStr);
  var since = (yesterday && yesterday < dateStr) ? yesterday : dateStr;
  readLogTail_(since).forEach(function (row) {
    if (row.section !== section) return;
    if (row.date === dateStr) existing[row.key] = row;
    if (row.date === yesterday) prior[row.key] = row;
  });

  var students = roster.map(function (r) {
    var sk = String(r['Student ID'] || '').trim() || String(r['Student Name']).trim();
    var hit = existing[logKey_(dateStr, section, sk)];
    var was = prior[logKey_(yesterday, section, sk)];
    return {
      id: String(r['Student ID'] || '').trim(),
      name: String(r['Student Name']).trim(),
      grade: String(r['Grade'] || '').trim(),
      status: hit ? hit.status : '',
      note: hit ? hit.note : '',
      previous: was ? was.status : ''
    };
  });

  var submittedAt = null;
  Object.keys(existing).forEach(function (k) {
    var ts = existing[k].timestamp;
    if (ts && (!submittedAt || ts > submittedAt)) submittedAt = ts;
  });

  return {
    section: section,
    date: dateStr,
    dateLabel: Utilities.formatDate(parseDate_(dateStr), tz_(), 'EEEE, MMMM d, yyyy'),
    students: students,
    locked: isLocked_(dateStr),
    submittedAt: submittedAt ? Utilities.formatDate(new Date(submittedAt), tz_(), 'h:mm a') : '',
    previousDate: yesterday
  };
}

/**
 * One round trip instead of two. The check-in page used to boot, wait, then ask
 * for the roster; this hands back both together and everything it needs is
 * already memoized from the first read.
 */
function getCheckInPayload(section, dateStr) {
  var boot = getBootstrap();
  if (!boot.allSections.length) return { boot: boot, day: null };

  section = String(section || '').trim();
  if (!section || boot.allSections.indexOf(section) === -1) section = boot.allSections[0];

  return { boot: boot, day: getSectionDay(section, dateStr || boot.today) };
}

function parseDate_(dateStr) {
  var p = String(dateStr).split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12, 0, 0);
}

function previousSchoolDay_(dateStr) {
  var skip = String(getSettings_().SKIP_DATES || '').split(',').map(function (s) { return s.trim(); });
  var d = parseDate_(dateStr);
  for (var i = 0; i < 10; i++) {
    d.setDate(d.getDate() - 1);
    var iso = fmtDate_(d);
    if (d.getDay() !== 0 && d.getDay() !== 6 && skip.indexOf(iso) === -1) return iso;
  }
  return '';
}

function isLocked_(dateStr) {
  var lock = String(getSettings_().LOCK_TIME || '').trim();
  if (!lock) return false;
  if (dateStr !== today_()) return dateStr < today_();
  var parts = lock.split(':');
  var cutoff = new Date();
  cutoff.setHours(Number(parts[0] || 23), Number(parts[1] || 0), 0, 0);
  return new Date() > cutoff;
}

function assertSectionAccess_(me, section) {
  if (me.isAdmin) return;
  if (me.sections.indexOf(section) === -1) {
    throw new Error('You are not listed as the teacher for ' + section + '. Ask the tech office to add you.');
  }
}

function submitCheckIn(payload) {
  var me = identity_();
  var section = String(payload.section || '').trim();
  var dateStr = String(payload.date || today_()).trim();
  assertSectionAccess_(me, section);

  if (isLocked_(dateStr) && !me.isAdmin) {
    throw new Error('Entries for ' + dateStr + ' are closed. Contact the main office to change them.');
  }

  var valid = {};
  getStatuses_().forEach(function (s) { valid[s.name] = true; });

  var entries = (payload.entries || []).filter(function (e) { return e && e.status; });
  entries.forEach(function (e) {
    if (!valid[e.status]) throw new Error('Unrecognized status: ' + e.status);
  });
  if (!entries.length) throw new Error('Nothing to save yet \u2014 mark at least one student.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sh = sheet_(SHEETS.LOG);

    // Existing rows for this date are near the bottom of the sheet, so the tail
    // read is enough to tell an edit from a first submission.
    var keyIndex = {};
    readLogTail_(dateStr).forEach(function (r) { keyIndex[r.key] = r.row; });

    var now = new Date();
    var appends = [];
    entries.forEach(function (e) {
      var sk = String(e.id || '').trim() || String(e.name || '').trim();
      var key = logKey_(dateStr, section, sk);
      var row = [key, now, dateStr, section, me.email, me.name,
                 String(e.id || '').trim(), String(e.name || '').trim(),
                 String(e.status).trim(), String(e.note || '').trim().slice(0, 300)];
      if (keyIndex[key]) sh.getRange(keyIndex[key], 1, 1, LOG_HEADERS.length).setValues([row]);
      else appends.push(row);
    });

    if (appends.length) {
      sh.getRange(sh.getLastRow() + 1, 1, appends.length, LOG_HEADERS.length).setValues(appends);
    }
  } finally {
    lock.releaseLock();
  }

  bumpReportVersion_();
  MEMO_ = {};
  return { ok: true, saved: entries.length, at: Utilities.formatDate(new Date(), tz_(), 'h:mm a') };
}

// ---------------------------------------------------------------------------
// Roster management
// ---------------------------------------------------------------------------

/**
 * Adds students to a section. Teachers can do this for their own sections when
 * ALLOW_TEACHER_ROSTER_EDITS is on; admins can do it anywhere.
 * One student per line. "Name, ID, Grade" on a line also works.
 */
function addStudents(section, namesText) {
  var me = identity_();
  section = String(section || '').trim();
  assertSectionAccess_(me, section);
  if (!me.isAdmin && !truthy_(getSettings_().ALLOW_TEACHER_ROSTER_EDITS)) {
    throw new Error('Roster changes are handled by the office. Send them the name and they will add it.');
  }

  var lines = String(namesText || '').split(/[\r\n]+/)
    .map(function (s) { return s.trim(); }).filter(String);
  if (!lines.length) throw new Error('Type at least one student name.');

  var existing = {};
  readObjectsUncached_(SHEETS.ROSTER).forEach(function (r) {
    if (String(r['Section']).trim() === section) {
      existing[String(r['Student Name']).trim().toLowerCase()] = r._row;
    }
  });

  var sh = sheet_(SHEETS.ROSTER);
  var rows = [], reactivated = 0;
  lines.forEach(function (line) {
    var parts = line.split(/[\t,]/).map(function (s) { return s.trim(); });
    var name = parts[0];
    if (!name) return;
    var key = name.toLowerCase();
    if (existing[key]) {
      if (typeof existing[key] === 'number') { sh.getRange(existing[key], 5).setValue(true); reactivated++; }
      return;
    }
    existing[key] = true;
    rows.push([name, parts[1] || '', parts[2] || '', section, true]);
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, ROSTER_HEADERS.length).setValues(rows);
  }
  clearCaches_();
  return { added: rows.length, reactivated: reactivated };
}

function setStudentActive(section, name, active) {
  var me = identity_();
  section = String(section || '').trim();
  assertSectionAccess_(me, section);
  if (!me.isAdmin && !truthy_(getSettings_().ALLOW_TEACHER_ROSTER_EDITS)) {
    throw new Error('Roster changes are handled by the office.');
  }
  var target = readObjectsUncached_(SHEETS.ROSTER).filter(function (r) {
    return String(r['Section']).trim() === section &&
           String(r['Student Name']).trim() === String(name).trim();
  })[0];
  if (!target) throw new Error('Could not find ' + name + ' in ' + section + '.');
  sheet_(SHEETS.ROSTER).getRange(target._row, 5).setValue(!!active);
  clearCaches_();
  return { ok: true };
}

function listRoster(section) {
  var me = identity_();
  section = String(section || '').trim();
  if (section) assertSectionAccess_(me, section);
  return roster_()
    .filter(function (r) {
      if (!String(r['Student Name']).trim()) return false;
      if (section && String(r['Section']).trim() !== section) return false;
      if (!me.isAdmin && me.sections.indexOf(String(r['Section']).trim()) === -1) return false;
      return true;
    })
    .map(function (r) {
      return {
        name: String(r['Student Name']).trim(),
        id: String(r['Student ID'] || '').trim(),
        grade: String(r['Grade'] || '').trim(),
        section: String(r['Section']).trim(),
        active: activeFlag_(r['Active'])
      };
    })
    .sort(function (a, b) {
      return a.section.localeCompare(b.section) || a.name.localeCompare(b.name);
    });
}

// ---------------------------------------------------------------------------
// Teacher management (admin)
// ---------------------------------------------------------------------------

function listTeachers() {
  assertAdmin_();
  return teachers_()
    .filter(function (t) { return String(t['Section']).trim() || String(t['Email']).trim(); })
    .map(function (t) {
      return {
        email: String(t['Email'] || '').trim(),
        name: String(t['Name'] || '').trim(),
        section: String(t['Section'] || '').trim(),
        room: String(t['Room'] || '').trim(),
        role: String(t['Role'] || 'Teacher').trim(),
        active: activeFlag_(t['Active'])
      };
    })
    .sort(function (a, b) { return a.section.localeCompare(b.section) || a.name.localeCompare(b.name); });
}

function addTeacher(t) {
  assertAdmin_();
  var email = String(t.email || '').trim().toLowerCase();
  var section = String(t.section || '').trim();
  var role = String(t.role || 'Teacher').trim();

  if (role.toLowerCase() !== 'admin' && !section) {
    throw new Error('Section is required for a teacher \u2014 it links them to a roster. Admins can be added without one.');
  }
  if (!email) throw new Error('An email address is required.');
  if (email.indexOf('@') === -1) throw new Error('That email address does not look right.');

  var dupe = readObjectsUncached_(SHEETS.TEACHERS).filter(function (r) {
    return String(r['Section']).trim() === section &&
           String(r['Email']).trim().toLowerCase() === email;
  })[0];
  if (dupe) throw new Error((section || 'That person') + ' is already on the list.');

  sheet_(SHEETS.TEACHERS).appendRow([
    email, String(t.name || '').trim() || section, section,
    String(t.room || '').trim(), role, true
  ]);
  clearCaches_();
  return { ok: true };
}

function updateTeacher(original, changes) {
  assertAdmin_();
  var target = readObjectsUncached_(SHEETS.TEACHERS).filter(function (r) {
    return String(r['Section']).trim() === String(original.section || '').trim() &&
           String(r['Email']).trim().toLowerCase() === String(original.email || '').trim().toLowerCase();
  })[0];
  if (!target) throw new Error('That row no longer exists. Reload and try again.');

  var row = [
    changes.email   !== undefined ? String(changes.email).trim().toLowerCase() : String(target['Email']).trim(),
    changes.name    !== undefined ? String(changes.name).trim()                : String(target['Name']).trim(),
    changes.section !== undefined ? String(changes.section).trim()             : String(target['Section']).trim(),
    changes.room    !== undefined ? String(changes.room).trim()                : String(target['Room']).trim(),
    changes.role    !== undefined ? String(changes.role).trim()                : String(target['Role']).trim(),
    changes.active  !== undefined ? !!changes.active                           : activeFlag_(target['Active'])
  ];
  sheet_(SHEETS.TEACHERS).getRange(target._row, 1, 1, TEACHER_HEADERS.length).setValues([row]);
  clearCaches_();
  return { ok: true };
}

// ---------------------------------------------------------------------------
// CSV import
// ---------------------------------------------------------------------------

function matchHeader_(raw, aliases) {
  var k = String(raw).trim().toLowerCase().replace(/[^a-z]/g, '');
  for (var canon in aliases) {
    if (aliases[canon].indexOf(k) > -1) return canon;
  }
  return null;
}

function parseCsvRows_(text, aliases, required) {
  var trimmed = String(text || '').replace(/^\uFEFF/, '').trim();
  if (!trimmed) throw new Error('That file looks empty.');

  var rows;
  try { rows = Utilities.parseCsv(trimmed); }
  catch (err) { throw new Error('That file could not be read as CSV. Export it again as Comma Separated Values.'); }
  if (!rows.length) throw new Error('That file looks empty.');

  var map = {};
  rows[0].forEach(function (h, i) {
    var canon = matchHeader_(h, aliases);
    if (canon && map[canon] === undefined) map[canon] = i;
  });

  var missing = required.filter(function (c) { return map[c] === undefined; });
  if (missing.length) {
    throw new Error('The file needs a column for: ' + missing.join(', ') +
                    '. Columns found: ' + rows[0].join(', ') + '.');
  }

  var out = [];
  for (var r = 1; r < rows.length; r++) {
    if (rows[r].join('').trim() === '') continue;
    var obj = {};
    for (var canon2 in map) obj[canon2] = String(rows[r][map[canon2]] || '').trim();
    out.push(obj);
  }
  return out;
}

var TEACHER_ALIASES = {
  email:   ['email', 'emailaddress', 'teacheremail', 'staffemail', 'googleaccount', 'username'],
  name:    ['name', 'teacher', 'teachername', 'staffname', 'displayname'],
  section: ['section', 'tab', 'homeroom', 'firstperiod', 'period', 'sectionname', 'class', 'course'],
  room:    ['room', 'roomnumber', 'location'],
  role:    ['role', 'type', 'accesslevel'],
  active:  ['active', 'enabled', 'status']
};

var ROSTER_ALIASES = {
  name:    ['studentname', 'name', 'student', 'fullname', 'lastfirst', 'studentlastfirst'],
  id:      ['studentid', 'id', 'studentnumber', 'localid', 'stateid'],
  grade:   ['grade', 'gradelevel', 'gl', 'yearlevel'],
  section: ['section', 'tab', 'homeroom', 'firstperiod', 'period', 'teacher', 'sectionname', 'class', 'course'],
  active:  ['active', 'enabled', 'status']
};

/** mode: 'merge' updates matching sections and adds new ones. 'replace' clears first. */
function importTeachersCsv(text, mode) {
  assertAdmin_();
  var rows = parseCsvRows_(text, TEACHER_ALIASES, ['section']);
  var sh = sheet_(SHEETS.TEACHERS);

  if (mode === 'replace' && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, TEACHER_HEADERS.length).clearContent();
  }

  var index = {};
  readObjectsUncached_(SHEETS.TEACHERS).forEach(function (r) {
    index[String(r['Section']).trim().toLowerCase()] = r._row;
  });

  var added = 0, updated = 0, appends = [];
  rows.forEach(function (r) {
    var section = r.section;
    if (!section) return;
    var out = [
      String(r.email || '').trim().toLowerCase(),
      r.name || section,
      section,
      r.room || '',
      r.role || 'Teacher',
      (r.active === undefined || r.active === '') ? true : truthy_(r.active)
    ];
    var at = index[section.toLowerCase()];
    if (typeof at === 'number') { sh.getRange(at, 1, 1, TEACHER_HEADERS.length).setValues([out]); updated++; }
    else { appends.push(out); index[section.toLowerCase()] = true; added++; }
  });

  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, TEACHER_HEADERS.length).setValues(appends);
  }
  clearCaches_();
  return { added: added, updated: updated, total: rows.length };
}

/** defaultSection fills in for any row whose section column is blank. */
function importRosterCsv(text, mode, defaultSection) {
  assertAdmin_();
  var rows = parseCsvRows_(text, ROSTER_ALIASES, ['name']);
  var sh = sheet_(SHEETS.ROSTER);

  if (mode === 'replace' && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, ROSTER_HEADERS.length).clearContent();
  }

  var index = {};
  readObjectsUncached_(SHEETS.ROSTER).forEach(function (r) {
    index[(String(r['Section']).trim() + '|' + String(r['Student Name']).trim()).toLowerCase()] = r._row;
  });

  var added = 0, updated = 0, skipped = 0, appends = [];
  rows.forEach(function (r) {
    var name = r.name;
    var section = r.section || String(defaultSection || '').trim();
    if (!name || !section) { skipped++; return; }
    var out = [
      name, r.id || '', r.grade || '', section,
      (r.active === undefined || r.active === '') ? true : truthy_(r.active)
    ];
    var k = (section + '|' + name).toLowerCase();
    var at = index[k];
    if (typeof at === 'number') { sh.getRange(at, 1, 1, ROSTER_HEADERS.length).setValues([out]); updated++; }
    else { appends.push(out); index[k] = true; added++; }
  });

  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, ROSTER_HEADERS.length).setValues(appends);
  }
  clearCaches_();
  return { added: added, updated: updated, skipped: skipped, total: rows.length };
}

function getCsvTemplate(kind) {
  if (kind === 'teachers') {
    return 'Email,Name,Section,Room,Role,Active\n' +
           'escott@royhart.org,Scott,Scott,214,Teacher,TRUE\n' +
           'jfeocco@royhart.org,Feocco,Feocco,118,Teacher,TRUE\n' +
           'principal@royhart.org,Building Admin,,,Admin,TRUE\n';
  }
  return 'Student Name,Student ID,Grade,Section,Active\n' +
         'Doe John,102345,10,Scott,TRUE\n' +
         'Roe Jane,102346,11,Scott,TRUE\n';
}

// ---------------------------------------------------------------------------
// Client API — report page
// ---------------------------------------------------------------------------

/** Thin wrapper kept for the daily digest and any direct calls. */
function getReportData(opts) {
  var out = buildReport_(opts);
  out.user = identity_();
  return out;
}

/**
 * Everything the report page draws, minus the student table. Cached across
 * requests: the payload is identical for every viewer, so one build serves the
 * whole building until someone submits.
 */
function getReportSummary(opts) {
  opts = opts || {};
  var days = Number(opts.days || 30);
  var section = String(opts.section || '').trim();
  var me = identity_();
  var cache = CacheService.getScriptCache();
  var v = reportVersion_();
  var sumKey = 'rs|' + v + '|' + days + '|' + section;

  var hit = cache.get(sumKey);
  if (hit) {
    try {
      var cachedOut = JSON.parse(hit);
      cachedOut.user = me;
      cachedOut.cached = true;
      return cachedOut;
    } catch (err) {}
  }

  var full = buildReport_({ days: days, section: section });
  var students = full.byStudent;
  delete full.byStudent;

  putCache_(cache, sumKey, full, 600);
  putCache_(cache, 'sd|' + v + '|' + days + '|' + section, students, 600);

  full.user = me;
  return full;
}

/** The student table, fetched after the page has already painted. */
function getStudentDetail(opts) {
  opts = opts || {};
  var days = Number(opts.days || 30);
  var section = String(opts.section || '').trim();
  var me = identity_();
  var showAll = truthy_(getSettings_().STUDENT_DETAIL_FOR_ALL) || me.isAdmin;

  var cache = CacheService.getScriptCache();
  var key = 'sd|' + reportVersion_() + '|' + days + '|' + section;
  var rows = null;
  var hit = cache.get(key);
  if (hit) { try { rows = JSON.parse(hit); } catch (err) {} }
  if (rows === null) {
    rows = buildReport_({ days: days, section: section }).byStudent;
    putCache_(cache, key, rows, 600);
  }

  if (!showAll) {
    rows = rows.filter(function (r) { return me.sections.indexOf(r.section) > -1; });
  }
  return { rows: rows, scope: showAll ? 'all' : 'own' };
}

function putCache_(cache, key, value, seconds) {
  try {
    var json = JSON.stringify(value);
    if (json.length < 95000) cache.put(key, json, seconds);
  } catch (err) {}
}

function buildReport_(opts) {
  opts = opts || {};
  var settings = getSettings_();
  var statuses = getStatuses_();
  var days = Number(opts.days || 30);
  var sectionFilter = String(opts.section || '').trim();

  var cutoff = '';
  if (days > 0) {
    var d = new Date();
    d.setDate(d.getDate() - Math.round(days * 1.45));
    cutoff = fmtDate_(d);
  }

  var full = cutoff ? readLogTail_(cutoff) : readLog_();
  var log = full.filter(function (r) {
    if (cutoff && r.date < cutoff) return false;
    if (sectionFilter && r.section !== sectionFilter) return false;
    return true;
  });

  var statusMeta = {};
  statuses.forEach(function (s) { statusMeta[s.name] = s; });

  var totals = { entries: log.length, compliant: 0, attention: 0, neutral: 0 };
  var byStatus = {};
  statuses.forEach(function (s) { byStatus[s.name] = 0; });

  log.forEach(function (r) {
    if (byStatus[r.status] === undefined) byStatus[r.status] = 0;
    byStatus[r.status]++;
    var cat = statusMeta[r.status] ? statusMeta[r.status].category : 'Compliant';
    if (cat === 'Attention') totals.attention++;
    else if (cat === 'Neutral') totals.neutral++;
    else totals.compliant++;
  });
  var graded = totals.compliant + totals.attention;
  totals.complianceRate = graded ? Math.round((totals.compliant / graded) * 1000) / 10 : null;

  var dayMap = {};
  log.forEach(function (r) {
    if (!dayMap[r.date]) dayMap[r.date] = { date: r.date, total: 0, attention: 0, compliant: 0, counts: {} };
    var d2 = dayMap[r.date];
    d2.total++;
    d2.counts[r.status] = (d2.counts[r.status] || 0) + 1;
    var cat = statusMeta[r.status] ? statusMeta[r.status].category : 'Compliant';
    if (cat === 'Attention') d2.attention++;
    else if (cat !== 'Neutral') d2.compliant++;
  });
  var trend = Object.keys(dayMap).sort().map(function (k) {
    var x = dayMap[k];
    var g = x.compliant + x.attention;
    x.complianceRate = g ? Math.round((x.compliant / g) * 1000) / 10 : null;
    x.label = Utilities.formatDate(parseDate_(x.date), tz_(), 'M/d');
    return x;
  });
  if (days > 0) trend = trend.slice(-days);

  var teacherBySection = {};
  teachers_().forEach(function (t) {
    var s = String(t['Section']).trim();
    if (s) teacherBySection[s] = { name: String(t['Name']).trim(), email: String(t['Email']).trim() };
  });

  var sectionMap = {};
  log.forEach(function (r) {
    if (!sectionMap[r.section]) {
      sectionMap[r.section] = {
        section: r.section,
        teacher: teacherBySection[r.section] ? teacherBySection[r.section].name : '',
        total: 0, attention: 0, compliant: 0, counts: {}, days: {}
      };
    }
    var s = sectionMap[r.section];
    s.total++;
    s.days[r.date] = true;
    s.counts[r.status] = (s.counts[r.status] || 0) + 1;
    var cat = statusMeta[r.status] ? statusMeta[r.status].category : 'Compliant';
    if (cat === 'Attention') s.attention++;
    else if (cat !== 'Neutral') s.compliant++;
  });
  var totalDays = trend.length;
  var bySection = Object.keys(sectionMap).sort().map(function (k) {
    var s = sectionMap[k];
    var g = s.compliant + s.attention;
    s.complianceRate = g ? Math.round((s.compliant / g) * 1000) / 10 : null;
    s.daysSubmitted = Object.keys(s.days).length;
    s.daysExpected = totalDays;
    delete s.days;
    return s;
  });

  var today = today_();
  var submittedToday = {};
  full.forEach(function (r) { if (r.date === today) submittedToday[r.section] = true; });
  var expected = allSections_();
  var missing = expected.filter(function (s) { return !submittedToday[s]; }).map(function (s) {
    return {
      section: s,
      teacher: teacherBySection[s] ? teacherBySection[s].name : '',
      email: teacherBySection[s] ? teacherBySection[s].email : ''
    };
  });

  var studentMap = {};
  log.forEach(function (r) {
    var key = (r.studentId || r.studentName) + '|' + r.section;
    if (!studentMap[key]) {
      studentMap[key] = { name: r.studentName, id: r.studentId, section: r.section,
                          total: 0, attention: 0, counts: {}, lastDate: '', lastStatus: '' };
    }
    var s = studentMap[key];
    s.total++;
    s.counts[r.status] = (s.counts[r.status] || 0) + 1;
    if (statusMeta[r.status] && statusMeta[r.status].category === 'Attention') s.attention++;
    if (r.date > s.lastDate) { s.lastDate = r.date; s.lastStatus = r.status; }
  });
  var byStudent = Object.keys(studentMap).map(function (k) { return studentMap[k]; })
    .sort(function (a, b) {
      if (b.attention !== a.attention) return b.attention - a.attention;
      return a.name.localeCompare(b.name);
    });

  return {
    generatedAt: Utilities.formatDate(new Date(), tz_(), 'MMM d, h:mm a'),
    schoolName: settings.SCHOOL_NAME,
    periodLabel: settings.PERIOD_LABEL,
    logoUrl: settings.LOGO_URL,
    appUrl: webAppUrl_(),
    statuses: statuses,
    range: { days: days, section: sectionFilter },
    sections: expected,
    totals: totals,
    byStatus: byStatus,
    trend: trend,
    bySection: bySection,
    today: { date: today, expected: expected.length, submitted: expected.length - missing.length, missing: missing },
    byStudent: byStudent
  };
}

function exportCsv(opts) {
  opts = opts || {};
  var me = identity_();
  var settings = getSettings_();
  var showAll = truthy_(settings.STUDENT_DETAIL_FOR_ALL) || me.isAdmin;

  var days = Number(opts.days || 30);
  var cutoff = '';
  if (days > 0) {
    var d = new Date();
    d.setDate(d.getDate() - Math.round(days * 1.45));
    cutoff = fmtDate_(d);
  }
  var sectionFilter = String(opts.section || '').trim();

  var rows = (cutoff ? readLogTail_(cutoff) : readLog_()).filter(function (r) {
    if (cutoff && r.date < cutoff) return false;
    if (sectionFilter && r.section !== sectionFilter) return false;
    if (!showAll && me.sections.indexOf(r.section) === -1) return false;
    return true;
  });

  var out = [['Date', 'Section', 'Teacher', 'Student ID', 'Student Name', 'Status', 'Note'].join(',')];
  rows.forEach(function (r) {
    out.push([r.date, r.section, r.teacherName, r.studentId, r.studentName, r.status, r.note]
      .map(function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(','));
  });
  return out.join('\n');
}

// ---------------------------------------------------------------------------
// Client API — admin console
// ---------------------------------------------------------------------------

function getAdminState() {
  var email = currentEmail_();
  var owner = isOwner_();
  var ready = isReady_();

  var state = {
    ready: ready,
    isOwner: owner,
    isAdmin: owner,
    email: email,
    name: email.split('@')[0],
    appUrl: webAppUrl_(),
    dataUrl: getDataUrl_(),
    dataName: '',
    sourceLinked: !!props_().getProperty(PROP_SOURCE_ID),
    sourceName: '',
    sourceError: '',
    triggersInstalled: false,
    counts: { teachers: 0, sections: 0, students: 0, entries: 0 },
    sections: [],
    settings: [],
    statuses: [],
    logoUrl: 'https://files.smartsites.parentsquare.com/4898/img_pd_102121_mxvilf.png',
    schoolName: 'Royalton-Hartland High School'
  };

  if (!ready) return state;

  var me = identity_();
  state.isAdmin = me.isAdmin;
  state.name = me.name;
  if (!me.isAdmin) return state;

  var settings = getSettings_();
  state.schoolName = settings.SCHOOL_NAME;
  state.logoUrl = settings.LOGO_URL;
  state.dataName = getSS_().getName();
  state.statuses = getStatuses_();
  state.sections = allSections_();

  try {
    state.triggersInstalled = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getHandlerFunction() === 'sendMissingReminders';
    });
  } catch (err) {}

  if (state.sourceLinked) {
    try { state.sourceName = sourceSS_().getName(); }
    catch (err2) { state.sourceError = 'Linked, but it will not open. Check that the file is still shared with you.'; }
  }

  var teachers = teachers_();
  var roster = roster_();
  state.counts = {
    teachers: teachers.filter(function (t) { return String(t['Email']).trim(); }).length,
    sections: state.sections.length,
    students: roster.filter(function (r) { return activeFlag_(r['Active']); }).length,
    entries: Math.max(0, sheet_(SHEETS.LOG).getLastRow() - 1)
  };

  state.settings = DEFAULT_SETTINGS.map(function (row) {
    return { key: row[0], value: settings[row[0]] === undefined ? row[1] : settings[row[0]], note: row[2] };
  });

  return state;
}

/** State, teachers, and roster in a single round trip. */
function getAdminPayload() {
  var state = getAdminState();
  if (!state.ready || !state.isAdmin) return { state: state, teachers: [], roster: [] };
  return { state: state, teachers: listTeachers(), roster: listRoster('') };
}

/** Creates the data spreadsheet. Only the person who deployed the app can run it. */
function runSetup() {
  if (!isOwner_()) throw new Error('Only the person who deployed this app can run setup.');
  var p = props_();
  var created = false;
  var ss;

  if (p.getProperty(PROP_DATA_ID)) {
    ss = getSS_();
  } else {
    var now = new Date();
    var startYear = now.getMonth() >= 6 ? now.getFullYear() : now.getFullYear() - 1;
    ss = SpreadsheetApp.create('Phone Check Data \u2014 ' + startYear + '-' + (startYear + 1));
    p.setProperty(PROP_DATA_ID, ss.getId());
    created = true;
    try { ss.setSpreadsheetTimeZone(Session.getScriptTimeZone()); } catch (err) {}
  }

  ensureSheet_(ss, SHEETS.SETTINGS, ['Setting', 'Value', 'Notes'], DEFAULT_SETTINGS);
  ensureSheet_(ss, SHEETS.TEACHERS, TEACHER_HEADERS, []);
  ensureSheet_(ss, SHEETS.ROSTER, ROSTER_HEADERS, []);
  ensureSheet_(ss, SHEETS.STATUSES, ['Status', 'Short Label', 'Category', 'Color', 'Order', 'Active'], DEFAULT_STATUSES);
  ensureSheet_(ss, SHEETS.LOG, LOG_HEADERS, []);

  var log = ss.getSheetByName(SHEETS.LOG);
  log.setFrozenRows(1);
  try { log.hideColumns(1); } catch (err2) {}

  [SHEETS.SETTINGS, SHEETS.TEACHERS, SHEETS.ROSTER, SHEETS.STATUSES].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setBackground('#63419A').setFontColor('#FFFFFF').setFontWeight('bold');
    sh.autoResizeColumns(1, sh.getLastColumn());
  });

  var stray = ss.getSheetByName('Sheet1');
  if (stray && ss.getSheets().length > 1 && stray.getLastRow() === 0) ss.deleteSheet(stray);

  if (created) {
    // Give the data spreadsheet its own Phone Check menu.
    try { ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create(); } catch (err3) {}
    // Seed the deployer as an admin so the console stays reachable.
    try {
      var settingsSh = ss.getSheetByName(SHEETS.SETTINGS);
      var rows = settingsSh.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (String(rows[i][0]).trim() === 'ADMIN_EMAILS' && !String(rows[i][1]).trim()) {
          settingsSh.getRange(i + 1, 2).setValue(currentEmail_());
          break;
        }
      }
    } catch (err4) {}
  }

  clearCaches_();
  return { created: created, url: ss.getUrl(), name: ss.getName() };
}

function ensureSheet_(ss, name, headers, seedRows) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (seedRows && seedRows.length) {
      sh.getRange(2, 1, seedRows.length, seedRows[0].length).setValues(seedRows);
    }
  } else if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sh;
}

function saveSettings(map) {
  assertAdmin_();
  var sh = sheet_(SHEETS.SETTINGS);
  var values = sh.getDataRange().getValues();
  var index = {};
  for (var i = 1; i < values.length; i++) index[String(values[i][0]).trim()] = i + 1;

  var appends = [];
  Object.keys(map).forEach(function (k) {
    var v = String(map[k]);
    if (index[k]) sh.getRange(index[k], 2).setValue(v);
    else {
      var note = (DEFAULT_SETTINGS.filter(function (r) { return r[0] === k; })[0] || ['', '', ''])[2];
      appends.push([k, v, note]);
    }
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, 3).setValues(appends);

  clearCaches_();
  return { ok: true };
}

/** Links a collection sheet to import from. View access is enough. */
function setSourceSheet(raw) {
  assertAdmin_();
  var match = String(raw || '').match(/[-\w]{25,}/);
  if (!match) throw new Error('Paste the full spreadsheet link, or just the long ID from the middle of it.');
  var id = match[0];
  var name;
  try { name = SpreadsheetApp.openById(id).getName(); }
  catch (err) { throw new Error('That spreadsheet will not open. Make sure it is shared with you \u2014 view access is enough.'); }
  props_().setProperty(PROP_SOURCE_ID, id);
  return { ok: true, name: name };
}

/**
 * Reads the linked collection sheet and builds sections, rosters, and any past
 * marks from it. Strictly read-only: nothing is written back to that file.
 */
function importFromSource() {
  assertAdmin_();
  var src = sourceSS_();

  var rosterSh = sheet_(SHEETS.ROSTER);
  var teacherSh = sheet_(SHEETS.TEACHERS);
  var logSh = sheet_(SHEETS.LOG);

  var existingStudents = {};
  readObjectsUncached_(SHEETS.ROSTER).forEach(function (r) {
    existingStudents[String(r['Student Name']).trim() + '|' + String(r['Section']).trim()] = true;
  });
  var existingTeachers = {};
  readObjectsUncached_(SHEETS.TEACHERS).forEach(function (t) { existingTeachers[String(t['Section']).trim()] = true; });
  var existingKeys = {};
  readLog_().forEach(function (r) { existingKeys[r.key] = true; });

  var statusNames = {};
  getStatuses_().forEach(function (s) { statusNames[s.name.toLowerCase()] = s.name; });

  var newStudents = [], newTeachers = [], newLog = [], tabs = [];

  src.getSheets().forEach(function (sh) {
    var name = sh.getName().trim();
    if (SYSTEM_SHEETS.indexOf(name) > -1) return;
    if (name.toLowerCase().indexOf('zz') === 0) return;
    if (sh.getLastRow() < 3 || sh.getLastColumn() < 2) return;
    tabs.push(name);

    var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var dateRow = values[0], statusRow = values[1];

    // Merged date headers only carry a value in their first cell; carry it across.
    var colDate = [], running = '';
    for (var c = 0; c < dateRow.length; c++) {
      var cell = dateRow[c];
      if (cell instanceof Date) running = fmtDate_(cell);
      else if (String(cell).trim()) {
        var parsed = new Date(String(cell).trim());
        running = isNaN(parsed.getTime()) ? '' : fmtDate_(parsed);
      }
      colDate[c] = running;
    }

    if (!existingTeachers[name]) {
      newTeachers.push(['', name, name, '', 'Teacher', true]);
      existingTeachers[name] = true;
    }

    for (var r = 2; r < values.length; r++) {
      var student = String(values[r][0]).trim();
      if (!student) continue;
      if (!existingStudents[student + '|' + name]) {
        newStudents.push([student, '', '', name, true]);
        existingStudents[student + '|' + name] = true;
      }
      for (var cc = 1; cc < values[r].length; cc++) {
        if (!truthy_(values[r][cc])) continue;
        var d = colDate[cc];
        var st = statusNames[String(statusRow[cc]).trim().toLowerCase()];
        if (!d || !st) continue;
        var key = logKey_(d, name, student);
        if (existingKeys[key]) continue;
        existingKeys[key] = true;
        newLog.push([key, new Date(), d, name, '', 'Imported', '', student, st, 'Imported from ' + name]);
      }
    }
  });

  if (newTeachers.length) teacherSh.getRange(teacherSh.getLastRow() + 1, 1, newTeachers.length, TEACHER_HEADERS.length).setValues(newTeachers);
  if (newStudents.length) rosterSh.getRange(rosterSh.getLastRow() + 1, 1, newStudents.length, ROSTER_HEADERS.length).setValues(newStudents);
  if (newLog.length) logSh.getRange(logSh.getLastRow() + 1, 1, newLog.length, LOG_HEADERS.length).setValues(newLog);

  clearCaches_();
  return {
    sourceName: src.getName(),
    tabs: tabs.length,
    sections: newTeachers.length,
    students: newStudents.length,
    entries: newLog.length
  };
}

function installTriggers() {
  assertAdmin_();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'sendMissingReminders' || fn === 'sendDailyDigest') ScriptApp.deleteTrigger(t);
  });
  var hour = Number(getSettings_().REMINDER_TIME || 9);
  ScriptApp.newTrigger('sendMissingReminders').timeBased().atHour(hour).everyDays(1).create();
  ScriptApp.newTrigger('sendDailyDigest').timeBased().atHour(Math.min(23, hour + 1)).everyDays(1).create();
  return { ok: true, hour: hour };
}

/** Rebuilds a wide Student x Date grid in the data spreadsheet, for printing. */
function rebuildGridView() {
  assertAdmin_();
  var ss = getSS_();
  var sh = ss.getSheetByName(SHEETS.GRID) || ss.insertSheet(SHEETS.GRID);
  sh.clear();

  var log = readLog_();
  if (!log.length) { sh.getRange(1, 1).setValue('No entries yet.'); return { rows: 0, days: 0 }; }

  var dates = {}, students = {};
  log.forEach(function (r) {
    dates[r.date] = true;
    students[r.section + '\u001f' + r.studentName] = true;
  });
  var dateList = Object.keys(dates).sort();
  var studentList = Object.keys(students).sort();
  var index = {};
  log.forEach(function (r) { index[r.date + '\u001f' + r.section + '\u001f' + r.studentName] = r.status; });

  var shortByName = {};
  getStatuses_().forEach(function (s) { shortByName[s.name] = s.short; });

  var header = ['Section', 'Student'].concat(dateList.map(function (d) {
    return Utilities.formatDate(parseDate_(d), tz_(), 'M/d');
  }));
  var rows = studentList.map(function (k) {
    var parts = k.split('\u001f');
    var row = [parts[0], parts[1]];
    dateList.forEach(function (d) {
      var st = index[d + '\u001f' + parts[0] + '\u001f' + parts[1]];
      row.push(st ? (shortByName[st] || st) : '');
    });
    return row;
  });

  sh.getRange(1, 1, 1, header.length).setValues([header])
    .setBackground('#63419A').setFontColor('#FFFFFF').setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  return { rows: rows.length, days: dateList.length };
}

// ---------------------------------------------------------------------------
// Menu on the data spreadsheet (installed by setup)
// ---------------------------------------------------------------------------

function onOpen() {
  var ui = ui_();
  if (!ui) return;
  ui.createMenu('Phone Check')
    .addItem('Show the app links', 'showLinks')
    .addItem('Rebuild printable grid', 'rebuildGridView')
    .addItem('Email teachers who haven\u2019t submitted', 'sendMissingReminders')
    .addItem('Send admin digest now', 'sendDailyDigest')
    .addToUi();
}

function showLinks() {
  var ui = ui_();
  var url = webAppUrl_();
  var msg = (url
    ? 'Teacher check-in:\n' + url +
      '\n\nTrends report:\n' + url + '?page=report' +
      '\n\nAdmin console:\n' + url + '?page=admin'
    : 'Deploy the script first: Deploy \u2192 New deployment \u2192 Web app.');
  if (ui) ui.alert('Phone Check links', msg, ui.ButtonSet.OK);
  else Logger.log(msg);
}

// ---------------------------------------------------------------------------
// Email
// ---------------------------------------------------------------------------

function isSchoolDay_(dateStr) {
  var d = parseDate_(dateStr);
  if (d.getDay() === 0 || d.getDay() === 6) return false;
  var skip = String(getSettings_().SKIP_DATES || '').split(',').map(function (s) { return s.trim(); });
  return skip.indexOf(dateStr) === -1;
}

function sendMissingReminders() {
  if (!isReady_()) return 0;
  var today = today_();
  if (!isSchoolDay_(today)) return 0;

  var submitted = {};
  readLogTail_(today).forEach(function (r) { if (r.date === today) submitted[r.section] = true; });

  var url = webAppUrl_();
  var settings = getSettings_();
  var sent = 0;

  teachers_().forEach(function (t) {
    var section = String(t['Section']).trim();
    var email = String(t['Email']).trim();
    if (!section || !email) return;
    if (!activeFlag_(t['Active'])) return;
    if (String(t['Role']).trim().toLowerCase() === 'admin') return;
    if (submitted[section]) return;

    MailApp.sendEmail({
      to: email,
      subject: 'Phone check not submitted \u2014 ' + Utilities.formatDate(new Date(), tz_(), 'MMM d'),
      htmlBody: reminderEmail_(String(t['Name']).trim(), section, url, settings)
    });
    sent++;
  });
  return sent;
}

function reminderEmail_(name, section, url, settings) {
  return '' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F3F0F8;padding:24px 12px;font-family:Arial,Helvetica,sans-serif;">' +
    '<tr><td align="center">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:560px;background:#FFFFFF;border-radius:14px;overflow:hidden;">' +
        '<tr><td style="background:#63419A;padding:18px 24px;color:#FFFFFF;font-size:18px;font-weight:bold;">' + settings.PERIOD_LABEL + ' Phone Check</td></tr>' +
        '<tr><td style="padding:24px;color:#333333;font-size:15px;line-height:1.6;">' +
          '<p style="margin:0 0 14px;">Good morning' + (name ? ' ' + name : '') + ',</p>' +
          '<p style="margin:0 0 14px;">Today\u2019s phone check for <strong>' + section + '</strong> hasn\u2019t come in yet. It takes about a minute.</p>' +
          '<p style="margin:0 0 22px;"><a href="' + url + '" style="display:inline-block;background:#63419A;color:#FFFFFF;text-decoration:none;padding:12px 26px;border-radius:999px;font-weight:bold;">Open the check-in</a></p>' +
          '<p style="margin:0;color:#636263;font-size:13px;">Already done it another way? Reply and let the tech office know so we can fix the list.</p>' +
        '</td></tr>' +
      '</table>' +
    '</td></tr>' +
  '</table>';
}

function sendDailyDigest() {
  if (!isReady_()) return;
  var settings = getSettings_();
  var to = String(settings.DIGEST_RECIPIENTS || '').split(',').map(function (s) { return s.trim(); }).filter(String);
  if (!to.length) return;
  var today = today_();
  if (!isSchoolDay_(today)) return;

  var data = getReportData({ days: 30 });
  var rows = data.statuses.map(function (s) {
    var n = 0;
    data.trend.forEach(function (d) { if (d.date === today) n = d.counts[s.name] || 0; });
    return '<tr><td style="padding:8px 12px;border-bottom:1px solid #EEE;color:#333;">' + s.name + '</td>' +
           '<td align="right" style="padding:8px 12px;border-bottom:1px solid #EEE;color:#333;font-weight:bold;">' + n + '</td></tr>';
  }).join('');

  var missing = data.today.missing.length
    ? data.today.missing.map(function (m) { return m.section + (m.teacher ? ' (' + m.teacher + ')' : ''); }).join(', ')
    : 'None \u2014 every section reported.';

  var html = '' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#F3F0F8;padding:24px 12px;font-family:Arial,Helvetica,sans-serif;">' +
    '<tr><td align="center"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background:#FFFFFF;border-radius:14px;overflow:hidden;">' +
      '<tr><td style="background:#63419A;padding:18px 24px;color:#FFFFFF;font-size:18px;font-weight:bold;">Phone Check \u2014 ' + Utilities.formatDate(new Date(), tz_(), 'EEEE, MMM d') + '</td></tr>' +
      '<tr><td style="padding:24px;color:#333333;font-size:15px;line-height:1.6;">' +
        '<p style="margin:0 0 6px;"><strong>' + data.today.submitted + ' of ' + data.today.expected + '</strong> sections reported.</p>' +
        '<p style="margin:0 0 18px;color:#636263;font-size:14px;">Still out: ' + missing + '</p>' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #EEE;border-radius:10px;">' + rows + '</table>' +
        '<p style="margin:20px 0 0;"><a href="' + data.appUrl + '?page=report" style="display:inline-block;background:#63419A;color:#FFFFFF;text-decoration:none;padding:11px 24px;border-radius:999px;font-weight:bold;">Open the full report</a></p>' +
      '</td></tr>' +
    '</table></td></tr>' +
  '</table>';

  MailApp.sendEmail({ to: to.join(','), subject: 'Phone check digest \u2014 ' + Utilities.formatDate(new Date(), tz_(), 'MMM d'), htmlBody: html });
}
