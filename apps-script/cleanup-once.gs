/**
 * ניקוי חד-פעמי של תוכן שלא הוגש על ידי סטודנטים.
 *
 * מה הפונקציה עושה:
 *   1. מוחקת לגמרי את פריטי המאמר (addedBy מכיל "פרדוקס").
 *   2. מחזירה את פריטי הבסיס (addedBy = "בסיס") לסטטוס "ממתין",
 *      רושמת אותך כמגיש, ומוסיפה הערה שמסבירה שהם נוצרו בזמן בניית האתר.
 *
 * הרצה: הדביקו כקובץ נוסף בעורך Apps Script, מלאו את MY_NAME למטה,
 *       בחרו בתפריט הפונקציות את cleanupOnce ולחצו Run. פעם אחת בלבד.
 *       אין צורך בפרסום מחדש (Deploy).
 *
 * לפני ההרצה מומלץ לגבות: בגיליון → קובץ → יצירת עותק.
 */

// ⬇⬇ מלאו כאן את שמכם המלא, כפי שתרצו שיופיע כמגיש ⬇⬇
const MY_NAME = 'מלאו כאן את שמכם';

const CLEANUP_NOTE = 'נוצר בזמן בניית האתר כתוכן התחלתי, לפני פתיחת ההגשות לסטודנטים.';

function cleanupOnce() {
  if (!MY_NAME || MY_NAME.indexOf('מלאו') === 0) {
    throw new Error('יש למלא את MY_NAME בראש הקובץ לפני ההרצה.');
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const report = [];

  ['terms', 'questions'].forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) { report.push(name + ': גיליון ריק'); return; }

    const data = sh.getDataRange().getValues();
    const head = data[0];
    const byCol = head.indexOf('addedBy');
    const stCol = head.indexOf('status');
    const noteCol = head.indexOf('note');
    if (byCol === -1 || stCol === -1) { report.push(name + ': חסרות עמודות addedBy/status'); return; }

    let deleted = 0, reset = 0;

    // עוברים מלמטה למעלה כדי שמחיקת שורה לא תזיז את השורות שטרם נבדקו
    for (let i = data.length - 1; i >= 1; i--) {
      const by = String(data[i][byCol] || '').trim();

      if (by.indexOf('פרדוקס') !== -1) {          // פריטי המאמר — מחיקה
        sh.deleteRow(i + 1);
        deleted++;
      } else if (by === 'בסיס') {                  // פריטי הבסיס — החזרה לאישור
        sh.getRange(i + 1, stCol + 1).setValue('ממתין');
        sh.getRange(i + 1, byCol + 1).setValue(MY_NAME);
        if (noteCol !== -1) sh.getRange(i + 1, noteCol + 1).setValue(CLEANUP_NOTE);
        reset++;
      }
    }
    report.push(name + ': נמחקו ' + deleted + ', הוחזרו לאישור ' + reset);
  });

  const msg = report.join('\n');
  Logger.log(msg);
  return msg;   // מוצג ב-Execution log
}
