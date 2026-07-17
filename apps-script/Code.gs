/**
 * מילון מושגים ותרגול — Industry 4.0
 * Google Apps Script — Web App backend על גבי Google Sheets
 *
 * התקנה: ראו apps-script/README.md
 */

const SHEET_TERMS = 'terms';
const SHEET_QUESTIONS = 'questions';
const STATUS_PENDING = 'ממתין';
const STATUS_APPROVED = 'מאושר';
const STATUS_REJECTED = 'נדחה';

const TERM_HEADERS = ['id', 'he', 'en', 'short', 'long', 'ex', 'topic', 'week', 'status', 'addedBy', 'timestamp'];
const QUESTION_HEADERS = ['id', 'term', 'q', 'opt1', 'opt2', 'opt3', 'opt4', 'correct', 'exp', 'status', 'addedBy', 'timestamp'];

/** הרץ פעם אחת כדי ליצור את הגיליונות והכותרות.
 *  (בלי alert — הרצה מתוך העורך אינה תומכת בחלונות UI ועלולה להיתקע) */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_TERMS, TERM_HEADERS);
  ensureSheet_(ss, SHEET_QUESTIONS, QUESTION_HEADERS);
  Logger.log('הגיליונות מוכנים: %s, %s. המשך ל-Deploy → New deployment → Web app.', SHEET_TERMS, SHEET_QUESTIONS);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function rows_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const head = data.shift();
  return data.map(function (r) {
    const o = {};
    head.forEach(function (h, i) { o[h] = r[i]; });
    return o;
  });
}

/** סטטוס מהגיליון -> סטטוס לאפליקציה. ריק/כל דבר שאינו "מאושר" = ממתין */
function statusOut_(v) {
  const s = String(v || '').trim();
  if (s === STATUS_APPROVED || s.toLowerCase() === 'approved' || s === 'TRUE' || v === true) return 'approved';
  if (s === STATUS_REJECTED || s.toLowerCase() === 'rejected') return 'rejected';
  return 'pending';
}

/** קריאת כל הנתונים — נקרא ע"י האפליקציה בטעינה */
function doGet(e) {
  try {
    const terms = rows_(SHEET_TERMS)
      .filter(function (t) { return t.id; })
      .map(function (t) {
        return {
          id: String(t.id), he: t.he, en: t.en, short: t.short, long: t.long,
          ex: t.ex, topic: t.topic || 'כללי', week: t.week || 1,
          status: statusOut_(t.status), addedBy: t.addedBy || ''
        };
      });
    const questions = rows_(SHEET_QUESTIONS)
      .filter(function (q) { return q.id; })
      .map(function (q) {
        return {
          id: String(q.id), term: String(q.term), q: q.q,
          opts: [q.opt1, q.opt2, q.opt3, q.opt4],
          correct: Number(q.correct) || 0, exp: q.exp,
          status: statusOut_(q.status), addedBy: q.addedBy || ''
        };
      });
    return json_({ ok: true, terms: terms, questions: questions });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** הוספה ואישור */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = body.action;

    if (action === 'addTerm') {
      const t = body.data || {};
      if (!t.he) return json_({ ok: false, error: 'חסר מונח בעברית' });
      const id = 't' + Date.now();
      ensureSheet_(ss, SHEET_TERMS, TERM_HEADERS).appendRow([
        id, t.he, t.en || '', t.short || '', t.long || '', t.ex || '',
        t.topic || 'כללי', t.week || 1, STATUS_PENDING, body.addedBy || '', new Date()
      ]);
      return json_({ ok: true, id: id });
    }

    if (action === 'addQuestion') {
      const q = body.data || {};
      if (!q.q) return json_({ ok: false, error: 'חסרה שאלה' });
      const opts = q.opts || [];
      const id = 'q' + Date.now();
      ensureSheet_(ss, SHEET_QUESTIONS, QUESTION_HEADERS).appendRow([
        id, q.term || '', q.q, opts[0] || '', opts[1] || '', opts[2] || '', opts[3] || '',
        Number(q.correct) || 0, q.exp || '', STATUS_PENDING, body.addedBy || '', new Date()
      ]);
      return json_({ ok: true, id: id });
    }

    /** אישור/ביטול — מוגן בקוד מנהל (Script Property בשם ADMIN_KEY) */
    if (action === 'setStatus') {
      const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
      if (!key || body.adminKey !== key) {
        return json_({ ok: false, error: 'קוד מנהל שגוי' });
      }
      const sheetName = body.type === 'term' ? SHEET_TERMS : SHEET_QUESTIONS;
      const sh = ss.getSheetByName(sheetName);
      if (!sh) return json_({ ok: false, error: 'גיליון לא נמצא' });
      const data = sh.getDataRange().getValues();
      const head = data[0];
      const idCol = head.indexOf('id');
      const stCol = head.indexOf('status');
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idCol]) === String(body.id)) {
          const newStatus = body.status === 'approved' ? STATUS_APPROVED
            : body.status === 'rejected' ? STATUS_REJECTED
            : STATUS_PENDING;
          sh.getRange(i + 1, stCol + 1).setValue(newStatus);
          return json_({ ok: true });
        }
      }
      return json_({ ok: false, error: 'מזהה לא נמצא' });
    }

    /** טעינה חד-פעמית של נתוני הבסיס לגיליון (מדלג על מזהים קיימים) — מוגן בקוד מנהל */
    if (action === 'importSeed') {
      const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
      if (!key || body.adminKey !== key) return json_({ ok: false, error: 'קוד מנהל שגוי' });
      const status = body.status === 'pending' ? STATUS_PENDING : STATUS_APPROVED;
      const tSheet = ensureSheet_(ss, SHEET_TERMS, TERM_HEADERS);
      const qSheet = ensureSheet_(ss, SHEET_QUESTIONS, QUESTION_HEADERS);
      const haveT = {}, haveQ = {};
      rows_(SHEET_TERMS).forEach(function (r) { haveT[String(r.id)] = true; });
      rows_(SHEET_QUESTIONS).forEach(function (r) { haveQ[String(r.id)] = true; });
      let added = 0;
      (body.terms || []).forEach(function (t) {
        if (haveT[String(t.id)]) return;
        tSheet.appendRow([t.id, t.he, t.en || '', t.short || '', t.long || '', t.ex || '',
          t.topic || 'כללי', t.week || 1, status, 'בסיס', new Date()]);
        added++;
      });
      (body.questions || []).forEach(function (q) {
        if (haveQ[String(q.id)]) return;
        const o = q.opts || [];
        qSheet.appendRow([q.id, q.term || '', q.q, o[0] || '', o[1] || '', o[2] || '', o[3] || '',
          Number(q.correct) || 0, q.exp || '', status, 'בסיס', new Date()]);
        added++;
      });
      return json_({ ok: true, added: added });
    }

    return json_({ ok: false, error: 'פעולה לא מוכרת' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
