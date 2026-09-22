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
const STATUS_REVISE = 'לתיקון';

const TERM_HEADERS = ['id', 'he', 'en', 'short', 'long', 'ex', 'topic', 'week', 'status', 'addedBy', 'note', 'timestamp'];
const QUESTION_HEADERS = ['id', 'term', 'q', 'opt1', 'opt2', 'opt3', 'opt4', 'correct', 'exp', 'status', 'addedBy', 'note', 'timestamp'];

/** הרץ פעם אחת כדי ליצור את הגיליונות והכותרות.
 *  (בלי alert — הרצה מתוך העורך אינה תומכת בחלונות UI ועלולה להיתקע) */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, SHEET_TERMS, TERM_HEADERS);
  ensureSheet_(ss, SHEET_QUESTIONS, QUESTION_HEADERS);
  Logger.log('הגיליונות מוכנים: %s, %s. המשך ל-Deploy → New deployment → Web app.', SHEET_TERMS, SHEET_QUESTIONS);
}

/** מוודא שהגיליון קיים ושכל העמודות הנדרשות קיימות בו (מוסיף חסרות בסוף) */
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return sh;
  }
  // מיגרציה: הוספת עמודות חדשות לגיליון קיים, בלי לגעת בנתונים
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const missing = headers.filter(function (h) { return head.indexOf(h) === -1; });
  if (missing.length) {
    sh.getRange(1, head.length + 1, 1, missing.length)
      .setValues([missing]).setFontWeight('bold');
  }
  return sh;
}

/** הוספת שורה לפי שמות העמודות בפועל — עמיד לשינויי סדר ולעמודות שנוספו */
function appendByHeaders_(sh, obj) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = head.map(function (h) {
    return Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : '';
  });
  sh.appendRow(row);
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
  if (s === STATUS_REVISE || s.toLowerCase() === 'revise') return 'revise';
  return 'pending';
}

/** סטטוס מהאפליקציה -> הערך שנכתב בגיליון */
function statusIn_(s) {
  return s === 'approved' ? STATUS_APPROVED
    : s === 'rejected' ? STATUS_REJECTED
    : s === 'revise' ? STATUS_REVISE
    : STATUS_PENDING;
}

/** מיפוי מזהה -> מספר שורה, בקריאת עמודת המזהים בלבד (ולא כל הגיליון) */
function idRows_(sh) {
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idCol = head.indexOf('id');
  const last = sh.getLastRow();
  const map = {};
  if (idCol !== -1 && last >= 2) {
    const col = sh.getRange(2, idCol + 1, last - 1, 1).getValues();
    for (let i = 0; i < col.length; i++) {
      const v = String(col[i][0]).trim();
      if (v && !map[v]) map[v] = i + 2;      // מספר השורה בגיליון
    }
  }
  return { head: head, map: map, last: last };
}

/** עדכון סטטוס (והערה) לפריט אחד או לרבים — קריאה אחת וכתיבה אחת לכל גיליון.
 *  כך אישור של עשרים פריטים עולה כמעט כמו אישור של אחד, וזה מה שחוסך למרצה
 *  את ההמתנה: Apps Script עונה באיטיות, ולכן מספר הקריאות הוא מה שקובע. */
function applyStatuses_(ss, items) {
  const done = [], missing = [];
  ['term', 'question'].forEach(function (type) {
    const isTerm = type === 'term';
    const list = items.filter(function (it) { return (it.type === 'term') === isTerm; });
    if (!list.length) return;
    const sh = ensureSheet_(ss, isTerm ? SHEET_TERMS : SHEET_QUESTIONS,
      isTerm ? TERM_HEADERS : QUESTION_HEADERS);
    const found = idRows_(sh);
    const stCol = found.head.indexOf('status') + 1;
    const noteCol = found.head.indexOf('note') + 1;
    const rows = found.last - 1;
    if (rows < 1 || !stCol) {
      list.forEach(function (it) { missing.push(it.id); });
      return;
    }
    const statuses = sh.getRange(2, stCol, rows, 1).getValues();
    const notes = noteCol ? sh.getRange(2, noteCol, rows, 1).getValues() : null;
    let wroteStatus = false, wroteNote = false;
    list.forEach(function (it) {
      const row = found.map[String(it.id).trim()];
      if (!row) { missing.push(it.id); return; }
      statuses[row - 2][0] = statusIn_(it.status);
      wroteStatus = true;
      if (notes && typeof it.note === 'string') { notes[row - 2][0] = it.note; wroteNote = true; }
      done.push(it.id);
    });
    if (wroteStatus) sh.getRange(2, stCol, rows, 1).setValues(statuses);
    if (wroteNote) sh.getRange(2, noteCol, rows, 1).setValues(notes);
  });
  return { done: done, missing: missing };
}

/** קריאת כל השורות מהגיליונות, ללא סינון (לשימוש פנימי בלבד) */
function readAll_() {
  const terms = rows_(SHEET_TERMS)
    .filter(function (t) { return t.id; })
    .map(function (t) {
      return {
        id: String(t.id), he: t.he, en: t.en, short: t.short, long: t.long,
        ex: t.ex, topic: t.topic || 'כללי', week: t.week || 1,
        status: statusOut_(t.status), addedBy: t.addedBy || '', note: t.note || ''
      };
    });
  const questions = rows_(SHEET_QUESTIONS)
    .filter(function (q) { return q.id; })
    .map(function (q) {
      return {
        id: String(q.id), term: String(q.term), q: q.q,
        opts: [q.opt1, q.opt2, q.opt3, q.opt4],
        correct: Number(q.correct) || 0, exp: q.exp,
        status: statusOut_(q.status), addedBy: q.addedBy || '', note: q.note || ''
      };
    });
  return { terms: terms, questions: questions };
}

/** השוואת שמות: רווחים כפולים ורישיות לא אמורים לנתק אדם מההגשה שלו */
function normName_(s) {
  return String(s || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

/** האם השם 'me' מופיע ברשימת המגישים של הפריט */
function ownedBy_(item, me) {
  const target = normName_(me);
  if (!target) return false;
  return String(item.addedBy || '')
    .split(/[,،;\/|]+/)
    .some(function (n) { return normName_(n) === target; });
}

/**
 * קריאה ע"י האפליקציה.
 * ברירת מחדל מחזירה אך ורק פריטים מאושרים — כך שתוכן שטרם אושר
 * (כולל התשובות הנכונות של שאלות ממתינות) אינו יוצא מהשרת.
 * פרמטר me מחזיר בנוסף את הפריטים שאותו אדם רשום בהם.
 */
function doGet(e) {
  try {
    const me = (e && e.parameter && e.parameter.me) ? e.parameter.me : '';
    const all = readAll_();
    const keep = function (x) { return x.status === 'approved' || ownedBy_(x, me); };
    return json_({
      ok: true,
      terms: all.terms.filter(keep),
      questions: all.questions.filter(keep)
    });
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
      appendByHeaders_(ensureSheet_(ss, SHEET_TERMS, TERM_HEADERS), {
        id: id, he: t.he, en: t.en || '', short: t.short || '', long: t.long || '',
        ex: t.ex || '', topic: t.topic || 'כללי', week: t.week || 1,
        status: STATUS_PENDING, addedBy: body.addedBy || '', note: '', timestamp: new Date()
      });
      return json_({ ok: true, id: id });
    }

    if (action === 'addQuestion') {
      const q = body.data || {};
      if (!q.q) return json_({ ok: false, error: 'חסרה שאלה' });
      const opts = q.opts || [];
      const id = 'q' + Date.now();
      appendByHeaders_(ensureSheet_(ss, SHEET_QUESTIONS, QUESTION_HEADERS), {
        id: id, term: q.term || '', q: q.q,
        opt1: opts[0] || '', opt2: opts[1] || '', opt3: opts[2] || '', opt4: opts[3] || '',
        correct: Number(q.correct) || 0, exp: q.exp || '',
        status: STATUS_PENDING, addedBy: body.addedBy || '', note: '', timestamp: new Date()
      });
      return json_({ ok: true, id: id });
    }

    /** רשימה מלאה כולל ממתינים — למסך האישור בלבד, מוגן בקוד מנהל */
    if (action === 'listAll') {
      const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
      if (!key || body.adminKey !== key) return json_({ ok: false, error: 'קוד מנהל שגוי' });
      const all = readAll_();
      return json_({ ok: true, terms: all.terms, questions: all.questions });
    }

    /** אישור/ביטול — פריט אחד (setStatus) או כמה בבת אחת (setStatusBulk, עם
     *  items: [{type,id,status,note}]). מוגן בקוד מנהל (Script Property ADMIN_KEY).
     *  הערת המרצה (סיבת דחייה / מה לתקן) נשמרת גם לצפיית המגישים. */
    if (action === 'setStatus' || action === 'setStatusBulk') {
      const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
      if (!key || body.adminKey !== key) {
        return json_({ ok: false, error: 'קוד מנהל שגוי' });
      }
      const items = action === 'setStatus'
        ? [{ type: body.type, id: body.id, status: body.status, note: body.note }]
        : (body.items || []);
      if (!items.length) return json_({ ok: false, error: 'לא נשלחו פריטים' });
      // נעילה — שתי כתיבות במקביל על אותה עמודה היו דורסות זו את זו
      const lock = LockService.getScriptLock();
      try { lock.waitLock(25000); }
      catch (e) { return json_({ ok: false, error: 'השרת עסוק כרגע, נסו שוב' }); }
      let res;
      try { res = applyStatuses_(ss, items); } finally { lock.releaseLock(); }
      if (action === 'setStatus') {
        return res.done.length ? json_({ ok: true }) : json_({ ok: false, error: 'מזהה לא נמצא' });
      }
      return json_({ ok: true, done: res.done, missing: res.missing });
    }

    /** עריכה. שני מסלולים:
     *  - מגישים: מותר לפריט שטרם הוכרע — "ממתין" או "לתיקון" — ולעולם לא לתוכן
     *    שאושר או נדחה. בפריט שממתין נדרשת גם התאמת שם, כדי שאדם לא יערוך
     *    הגשה של אחר. השליחה מחזירה את הפריט ל"ממתין" (לתור האישור).
     *  - מנהל (adminKey תקין): מותר לערוך כל פריט בכל סטטוס. הסטטוס והמגישים
     *    אינם משתנים — תיקון של המרצה במאושר משאיר אותו מאושר. */
    if (action === 'updateItem') {
      const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
      const isAdmin = !!body.adminKey && !!key && body.adminKey === key;
      if (body.adminKey && !isAdmin) return json_({ ok: false, error: 'קוד מנהל שגוי' });
      const isTerm = body.type === 'term';
      const sheetName = isTerm ? SHEET_TERMS : SHEET_QUESTIONS;
      const sh = ensureSheet_(ss, sheetName, isTerm ? TERM_HEADERS : QUESTION_HEADERS);
      const data = sh.getDataRange().getValues();
      const head = data[0];
      const idCol = head.indexOf('id');
      const stCol = head.indexOf('status');
      const byCol0 = head.indexOf('addedBy');
      const d = body.data || {};
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idCol]) !== String(body.id)) continue;
        const cur = statusOut_(data[i][stCol]);
        if (!isAdmin) {
          if (cur !== 'revise' && cur !== 'pending') {
            return json_({ ok: false, error: 'ניתן לערוך רק הגשה שטרם הוכרעה' });
          }
          if (cur === 'pending') {
            const owner = byCol0 === -1 ? '' : data[i][byCol0];
            if (!ownedBy_({ addedBy: owner }, body.me)) {
              return json_({ ok: false, error: 'ניתן לערוך רק הגשה שאתם רשומים בה' });
            }
          }
        }
        const vals = isTerm
          ? { he: d.he, en: d.en, short: d.short, long: d.long, ex: d.ex, topic: d.topic, week: d.week }
          : { term: d.term, q: d.q, opt1: (d.opts || [])[0], opt2: (d.opts || [])[1],
              opt3: (d.opts || [])[2], opt4: (d.opts || [])[3], correct: Number(d.correct) || 0, exp: d.exp };
        Object.keys(vals).forEach(function (k) {
          const c = head.indexOf(k);
          if (c !== -1 && vals[k] !== undefined) sh.getRange(i + 1, c + 1).setValue(vals[k]);
        });
        if (!isAdmin && body.addedBy) {
          const byCol = head.indexOf('addedBy');
          if (byCol !== -1) sh.getRange(i + 1, byCol + 1).setValue(body.addedBy);
        }
        const tsCol = head.indexOf('timestamp');
        if (tsCol !== -1) sh.getRange(i + 1, tsCol + 1).setValue(new Date());
        if (!isAdmin) sh.getRange(i + 1, stCol + 1).setValue(STATUS_PENDING);  // חוזר לתור האישור
        return json_({ ok: true });
      }
      return json_({ ok: false, error: 'מזהה לא נמצא' });
    }

    /** מחיקת הגשה ע"י המגישים — הדרך הנכונה להסיר כפל הגשה.
     *  מותר רק לפריט שטרם אושר ("ממתין" או "לתיקון") ורק למי שרשום בו.
     *  תוכן שאושר או נדחה לעולם אינו נמחק מהאתר. מנהל עם קוד תקין מוחק הכול. */
    if (action === 'deleteItem') {
      const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
      const isAdmin = !!body.adminKey && !!key && body.adminKey === key;
      if (body.adminKey && !isAdmin) return json_({ ok: false, error: 'קוד מנהל שגוי' });
      const isTerm = body.type === 'term';
      const sheetName = isTerm ? SHEET_TERMS : SHEET_QUESTIONS;
      const sh = ensureSheet_(ss, sheetName, isTerm ? TERM_HEADERS : QUESTION_HEADERS);
      const data = sh.getDataRange().getValues();
      const head = data[0];
      const idCol = head.indexOf('id');
      const stCol = head.indexOf('status');
      const byCol = head.indexOf('addedBy');
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][idCol]) !== String(body.id)) continue;
        if (!isAdmin) {
          const cur = statusOut_(data[i][stCol]);
          if (cur !== 'pending' && cur !== 'revise') {
            return json_({ ok: false, error: 'ניתן למחוק רק הגשה שהמרצה טרם אישר' });
          }
          if (!ownedBy_({ addedBy: byCol === -1 ? '' : data[i][byCol] }, body.me)) {
            return json_({ ok: false, error: 'ניתן למחוק רק הגשה שאתם רשומים בה' });
          }
        }
        sh.deleteRow(i + 1);
        return json_({ ok: true });
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
        appendByHeaders_(tSheet, {
          id: t.id, he: t.he, en: t.en || '', short: t.short || '', long: t.long || '',
          ex: t.ex || '', topic: t.topic || 'כללי', week: t.week || 1,
          status: status, addedBy: 'בסיס', note: '', timestamp: new Date()
        });
        added++;
      });
      (body.questions || []).forEach(function (q) {
        if (haveQ[String(q.id)]) return;
        const o = q.opts || [];
        appendByHeaders_(qSheet, {
          id: q.id, term: q.term || '', q: q.q,
          opt1: o[0] || '', opt2: o[1] || '', opt3: o[2] || '', opt4: o[3] || '',
          correct: Number(q.correct) || 0, exp: q.exp || '',
          status: status, addedBy: 'בסיס', note: '', timestamp: new Date()
        });
        added++;
      });
      return json_({ ok: true, added: added });
    }

    return json_({ ok: false, error: 'פעולה לא מוכרת' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
