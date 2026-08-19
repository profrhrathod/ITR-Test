/**
 * INDUSTRY INTERNSHIP TEST — Google Sheet backend
 * ---------------------------------------------------------------
 * SETUP:
 * 1. Create a new Google Sheet (any name, any starting tabs — this
 *    script will create the tabs it needs automatically).
 * 2. Extensions -> Apps Script. Delete any starter code and paste
 *    this whole file in.
 * 3. Click "Deploy" -> "New deployment" -> select type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 4. Click Deploy, authorize the permissions it asks for, then copy
 *    the Web app URL (it ends in /exec).
 * 5. Paste that URL into the API_URL constant near the top of the
 *    test's HTML file, then save and reopen the HTML file.
 *
 * Tabs created automatically:
 *   Config    -> Key | Value
 *   Questions -> ID | Domain | Question | OptionA | OptionB | OptionC | OptionD | Correct
 *   Roster    -> Roll | Name | Enrolment | Class
 *   Results   -> Roll | Name | Enrolment | Class | Score | OutOf | Percent | SubmittedAt | Answers
 *   Notices   -> Notice
 */

const SHEET_CONFIG = 'Config';
const SHEET_QUESTIONS = 'Questions';
const SHEET_ROSTER = 'Roster';
const SHEET_RESULTS = 'Results';
const SHEET_NOTICES = 'Notices';

function doGet(e) {
  try {
    const action = e.parameter.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let data;
    if (action === 'config') data = getConfig(ss);
    else if (action === 'questions') data = getQuestions(ss);
    else if (action === 'roster') data = getRoster(ss);
    else if (action === 'results') data = getResults(ss);
    else if (action === 'result') data = getResult(ss, e.parameter.roll, e.parameter.enrolment);
    else if (action === 'notices') data = getNotices(ss);
    else if (action === 'all') data = getAll(ss); // single round-trip: loads everything the app needs on startup
    else data = { error: 'Unknown action: ' + action };
    return jsonOut(data);
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function getAll(ss) {
  return {
    config: getConfig(ss),
    questions: getQuestions(ss),
    roster: getRoster(ss),
    notices: getNotices(ss)
  };
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let result;
    if (action === 'saveConfig') result = saveConfig(ss, body.data);
    else if (action === 'saveQuestions') result = saveQuestions(ss, body.data);
    else if (action === 'saveRoster') result = saveRoster(ss, body.data);
    else if (action === 'addResult') result = addResult(ss, body.data);
    else if (action === 'saveNotices') result = saveNotices(ss, body.data);
    else result = { error: 'Unknown action: ' + action };
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ---------------- Config (key/value rows) ---------------- */
function getConfig(ss) {
  const sh = getOrCreateSheet(ss, SHEET_CONFIG, ['Key', 'Value']);
  const rows = sh.getDataRange().getValues();
  const cfg = {};
  for (let i = 1; i < rows.length; i++) {
    const k = rows[i][0];
    if (!k) continue;
    cfg[k] = rows[i][1];
  }
  // These keys are always compared as text on the client (passwords/codes/dates),
  // so force them back to strings here in case Sheets auto-converted an
  // all-digit value (e.g. a numeric password) into a Number when it was saved.
  ['adminPass', 'testCode', 'testHeading', 'year', 'testDate', 'startAt', 'endAt'].forEach(k => {
    if (cfg[k] !== undefined && cfg[k] !== '') cfg[k] = String(cfg[k]);
  });
  return cfg;
}
function saveConfig(ss, data) {
  const sh = getOrCreateSheet(ss, SHEET_CONFIG, ['Key', 'Value']);
  sh.clearContents();
  sh.appendRow(['Key', 'Value']);
  const keys = Object.keys(data || {});
  keys.forEach((k, i) => {
    const rowNum = i + 2;
    const val = data[k];
    if (typeof val === 'string') {
      // Force this cell to plain-text format BEFORE writing, so an all-digit
      // string (like a numeric password) is stored as text, not auto-converted
      // into a Number by Sheets.
      sh.getRange(rowNum, 2).setNumberFormat('@');
    }
    sh.getRange(rowNum, 1).setValue(k);
    sh.getRange(rowNum, 2).setValue(val);
  });
  return { ok: true };
}

/* ---------------- Questions ---------------- */
function getQuestions(ss) {
  const sh = getOrCreateSheet(ss, SHEET_QUESTIONS,
    ['ID', 'Domain', 'Question', 'OptionA', 'OptionB', 'OptionC', 'OptionD', 'Correct']);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[2]) continue; // no question text -> skip blank row
    out.push({
      id: String(r[0] || ('q' + i)),
      domain: r[1] || '',
      text: r[2],
      options: [r[3], r[4], r[5], r[6]],
      correct: Number(r[7])
    });
  }
  return out;
}
function saveQuestions(ss, data) {
  const sh = getOrCreateSheet(ss, SHEET_QUESTIONS,
    ['ID', 'Domain', 'Question', 'OptionA', 'OptionB', 'OptionC', 'OptionD', 'Correct']);
  sh.clearContents();
  sh.appendRow(['ID', 'Domain', 'Question', 'OptionA', 'OptionB', 'OptionC', 'OptionD', 'Correct']);
  (data || []).forEach(q => sh.appendRow(
    [q.id, q.domain, q.text, q.options[0], q.options[1], q.options[2], q.options[3], q.correct]
  ));
  return { ok: true };
}

/* ---------------- Roster ---------------- */
function getRoster(ss) {
  const sh = getOrCreateSheet(ss, SHEET_ROSTER, ['Roll', 'Name', 'Enrolment', 'Class']);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({ roll: String(r[0]), name: r[1] || '', enrolment: String(r[2] || ''), cls: r[3] || '' });
  }
  return out;
}
function saveRoster(ss, data) {
  const sh = getOrCreateSheet(ss, SHEET_ROSTER, ['Roll', 'Name', 'Enrolment', 'Class']);
  sh.clearContents();
  sh.appendRow(['Roll', 'Name', 'Enrolment', 'Class']);
  (data || []).forEach(s => sh.appendRow([s.roll, s.name, s.enrolment, s.cls || '']));
  return { ok: true };
}

/* ---------------- Results ---------------- */
const RESULTS_HEADER = ['Roll', 'Name', 'Enrolment', 'Class', 'Score', 'OutOf', 'Percent', 'SubmittedAt', 'TabSwitches', 'Answers'];

function ensureResultsSchema(sh) {
  // Migration for sheets created before the TabSwitches column existed:
  // insert it before the Answers column, shifting Answers right (no data lost).
  const header = sh.getRange(1, 1, 1, RESULTS_HEADER.length).getValues()[0];
  if (header[8] !== 'TabSwitches') {
    sh.insertColumnBefore(9);
    sh.getRange(1, 9).setValue('TabSwitches');
  }
}

function getResults(ss) {
  const sh = getOrCreateSheet(ss, SHEET_RESULTS, RESULTS_HEADER);
  ensureResultsSchema(sh);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({
      roll: String(r[0]), name: r[1] || '', enrolment: String(r[2] || ''), cls: r[3] || '',
      score: Number(r[4]), outOf: Number(r[5]), percent: Number(r[6]),
      submittedAt: r[7] instanceof Date ? r[7].toISOString() : r[7],
      tabSwitches: Number(r[8]) || 0,
      answers: safeParse(r[9])
    });
  }
  return out;
}
function getResult(ss, roll, enrolment) {
  const results = getResults(ss);
  const found = results.find(r => String(r.roll) === String(roll) && String(r.enrolment) === String(enrolment));
  return { found: !!found, result: found || null };
}
function addResult(ss, data) {
  const sh = getOrCreateSheet(ss, SHEET_RESULTS, RESULTS_HEADER);
  ensureResultsSchema(sh);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.roll) && String(rows[i][2]) === String(data.enrolment)) {
      return { ok: false, error: 'This Roll No. / Enrolment No. has already submitted the test.' };
    }
  }
  sh.appendRow([
    data.roll, data.name, data.enrolment, data.cls || '',
    data.score, data.outOf, data.percent, data.submittedAt,
    data.tabSwitches || 0,
    JSON.stringify(data.answers || [])
  ]);
  return { ok: true };
}
function safeParse(s) {
  try { return JSON.parse(s); } catch (e) { return []; }
}

/* ---------------- Notices ---------------- */
function getNotices(ss) {
  const sh = getOrCreateSheet(ss, SHEET_NOTICES, ['Notice']);
  const rows = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const v = rows[i][0];
    if (v) out.push(String(v));
  }
  return out;
}
function saveNotices(ss, data) {
  const sh = getOrCreateSheet(ss, SHEET_NOTICES, ['Notice']);
  sh.clearContents();
  sh.appendRow(['Notice']);
  (data || []).forEach(n => sh.appendRow([n]));
  return { ok: true };
}

