/**
 * שיוך חד-פעמי של שאלות למושגים + השבת מושגי בסיס חסרים.
 *
 * רקע: 11 שאלות מאושרות נשמרו עם עמודת term ריקה, ולכן עד לתיקון בקוד הן
 * לא הופיעו בתרגול מסונן לפי נושא, וגם הקישור "למושג" שלהן ריק. בנוסף,
 * שני מושגי בסיס — "תעשייה 5.0" ו"האינטרנט של הדברים" — חסרים בגיליון,
 * אף ששאלות מאושרות (q2, q7, q11 ועוד) מפנות אליהם.
 *
 * מה הסקריפט עושה:
 *   1. ממלא את עמודת term לפי המיפוי RELINK_MAP — רק בשורות שהעמודה ריקה
 *      בהן. לעולם אינו דורס שיוך קיים.
 *   2. מוסיף את המושגים החסרים (RESTORE_TERMS) כ"ממתין" — האישור נשאר
 *      בידי המרצה במסך ?approve. מדלג על מזהה שכבר קיים.
 *
 * שאלות שנשארות ללא שיוך (אין להן מושג מתאים במילון, והן יופיעו בתרגול
 * תחת "כללי"): שאלת ה-AI למכרזים, שתי שאלות GOTRACK ושאלת ניצולת הצי.
 * אם יתווספו למילון מושגים כמו "כלכלה שיתופית" או "ניצולת צי" — אפשר
 * לקשר אותן ידנית בעמודת term.
 *
 * הרצה: הדביקו כקובץ נוסף בעורך Apps Script, בחרו relinkOnce ולחצו Run.
 *       אין צורך בפרסום מחדש (Deploy).
 * מומלץ לגבות לפני: בגיליון → קובץ → יצירת עותק.
 *
 * להרצה יבשה (רק דוח, בלי לשנות) — שנו ל-true:
 */
const RELINK_DRY_RUN = false;

/** שאלה (id) → מושג (id). כל היעדים קיימים ומאושרים בגיליון. */
const RELINK_MAP = {
  // חשש ההנהלה מול חשש העובדים בהטמעת אוטומציה → התנגדות לשינוי
  'q1785312588352': 't1785705104925',
  // גיבוי מערכות וקריסה בהפסקת חשמל → רציפות תפקוד
  'q1785312779536': 'continuity',
  // בדיקות FAT לפני Go-Live → FAT
  'q1785505074770': 't1785704409772',
  // מלכודת הקוסטומיזציה ב-ERP/WMS → מלכודת הקסטומיזציה
  'q1785505208082': 't1785705392147',
  // שלב ה-Go-Live — אין מושג Go-Live במילון; FAT הוא הקרוב ביותר (אותה הרצאה).
  // אם מעדיפים להשאיר תחת "כללי" — מחקו את השורה.
  'q1785505342603': 't1785704409772',
  // "טכנולוגיה נבנית מהר, אמון נבנה לאט" → התנגדות לשינוי
  'q1785505450407': 't1785705104925',
  // "הר הברזל" — אגירת מלאי מראש על הרצף JIT/JIC → מלאי ליתר ביטחון (JIC)
  'q1785864812065': 't1785794079111'
};

/** מושגי הבסיס שאבדו — התוכן זהה ל-SEED שבאפליקציה, נכנס כ"ממתין" */
const RESTORE_TERMS = [
  { id: 'industry50', he: 'תעשייה 5.0', en: 'Industry 5.0', topic: 'תעשייה 5.0', week: 1,
    short: 'השלב הבא — העברת המיקוד מהטכנולוגיה לתפיסה הוליסטית: ממשק אדם־מכונה, חוסן וקיימות.',
    long: 'אם 4.0 היא הווה (וחלקה כבר עבר), 5.0 חלקה הווה ורובה עתיד. לא \'האדם במרכז\' במובן HR, אלא ההבנה שבלי שינוי הרגלים ויכולות אנושיות — הטכנולוגיה של 4.0 לא תמצה את הפוטנציאל שלה.',
    ex: 'הטמעת מערכות WMS וליקוט קולי אצל עובד שצריך ללמוד לעבוד מולן.' },
  { id: 'iot', he: 'האינטרנט של הדברים', en: 'Internet of Things (IoT)', topic: 'מבוא / הגדרות יסוד', week: 1,
    short: 'רשת של מכשירים פיזיים מחוברים שאוספים ומחליפים נתונים בזמן אמת.',
    long: 'אחת מאבני היסוד של תעשייה 4.0 — החיבור בין מערכות ומכונות שמאפשר להן \'לדבר\' זו עם זו ולפעול על בסיס נתונים משותפים.',
    ex: 'חיישנים במכונית או במחסן שמדווחים סטטוס למערכת מרכזית.' }
];

function relinkOnce() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];

  // 1. שיוך שאלות: מילוי עמודת term בשורות שבהן היא ריקה
  const qSh = ensureSheet_(ss, SHEET_QUESTIONS, QUESTION_HEADERS);
  const qData = qSh.getDataRange().getValues();
  const qHead = qData[0];
  const qIdCol = qHead.indexOf('id');
  const qTermCol = qHead.indexOf('term');
  const done = {};
  for (let i = 1; i < qData.length; i++) {
    const id = String(qData[i][qIdCol]);
    const target = RELINK_MAP[id];
    if (!target) continue;
    done[id] = true;
    const cur = String(qData[i][qTermCol] || '').trim();
    if (cur) { out.push('• ' + id + ': כבר משויך ל-"' + cur + '" — לא נגעתי'); continue; }
    if (!RELINK_DRY_RUN) qSh.getRange(i + 1, qTermCol + 1).setValue(target);
    out.push('• ' + id + ' → ' + target);
  }
  Object.keys(RELINK_MAP).forEach(function (id) {
    if (!done[id]) out.push('• ' + id + ': לא נמצא בגיליון!');
  });

  // 2. השבת מושגי בסיס חסרים כ"ממתין"
  const haveT = {};
  rows_(SHEET_TERMS).forEach(function (r) { haveT[String(r.id)] = true; });
  const tSh = ensureSheet_(ss, SHEET_TERMS, TERM_HEADERS);
  RESTORE_TERMS.forEach(function (t) {
    if (haveT[t.id]) { out.push('• מושג "' + t.he + '" (' + t.id + ') כבר קיים — דילגתי'); return; }
    if (!RELINK_DRY_RUN) appendByHeaders_(tSh, {
      id: t.id, he: t.he, en: t.en, short: t.short, long: t.long,
      ex: t.ex, topic: t.topic, week: t.week,
      status: STATUS_PENDING, addedBy: 'בסיס', note: '', timestamp: new Date()
    });
    out.push('• הוספתי מושג "' + t.he + '" (' + t.id + ') כ"ממתין" — לאישור המרצה');
  });

  const msg = (RELINK_DRY_RUN ? '— הרצה יבשה, לא שונה דבר —\n' : '') + out.join('\n');
  Logger.log(msg);
  return msg;   // מוצג ב-Execution log
}
