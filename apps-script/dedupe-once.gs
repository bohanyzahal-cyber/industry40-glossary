/**
 * מחיקת כפילויות חד-פעמית.
 *
 * מזהה מושגים בעלי אותו שם ושאלות בעלות אותו נוסח (בהתעלם מרווחים,
 * גרשיים וסימני פיסוק), ומשאיר מכל קבוצה את השורה ה**חדשה ביותר** לפי
 * עמודת timestamp — כלומר את התיקון האחרון ולא את הגרסה שהוחזרה.
 * שורות שסטטוסן "נדחה" אינן נחשבות ואינן נמחקות.
 *
 * הרצה: הדביקו כקובץ נוסף בעורך Apps Script, בחרו את dedupeOnce ולחצו Run.
 *       אין צורך בפרסום מחדש (Deploy).
 *
 * מומלץ לגבות לפני: בגיליון → קובץ → יצירת עותק.
 *
 * להרצה יבשה (רק דוח, בלי למחוק) — שנו ל-true:
 */
const DRY_RUN = false;

function dedupeOnce() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];

  [{ sheet: 'terms', keyCol: 'he' },
   { sheet: 'questions', keyCol: 'q' }].forEach(function (cfg) {

    const sh = ss.getSheetByName(cfg.sheet);
    if (!sh || sh.getLastRow() < 2) { out.push(cfg.sheet + ': גיליון ריק'); return; }

    const data = sh.getDataRange().getValues();
    const head = data[0];
    const kCol = head.indexOf(cfg.keyCol);
    const stCol = head.indexOf('status');
    const tsCol = head.indexOf('timestamp');
    const byCol = head.indexOf('addedBy');
    if (kCol === -1) { out.push(cfg.sheet + ': לא נמצאה עמודת ' + cfg.keyCol); return; }

    // אוספים את השורות לפי מפתח מנורמל
    const groups = {};
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][stCol] || '').trim();
      if (status === 'נדחה') continue;                       // נדחו — לא נוגעים
      const key = normKey_(data[i][kCol]);
      if (!key) continue;
      const ts = tsCol === -1 ? 0 : toTime_(data[i][tsCol]);
      (groups[key] = groups[key] || []).push({ row: i + 1, ts: ts, status: status,
                                               by: byCol === -1 ? '' : data[i][byCol],
                                               label: String(data[i][kCol]) });
    }

    // בכל קבוצה עם יותר משורה אחת — משאירים את החדשה ביותר
    const toDelete = [];
    Object.keys(groups).forEach(function (k) {
      const rows = groups[k];
      if (rows.length < 2) return;
      rows.sort(function (a, b) { return b.ts - a.ts; });     // חדש → ישן
      const keep = rows[0];
      out.push('• "' + keep.label.slice(0, 40) + '" (' + keep.by + ') — ' +
               rows.length + ' עותקים; נשמר: שורה ' + keep.row + ' [' + keep.status + ']');
      rows.slice(1).forEach(function (r) {
        out.push('    נמחקת שורה ' + r.row + ' [' + r.status + ']');
        toDelete.push(r.row);
      });
    });

    // מוחקים מלמטה למעלה כדי שמספרי השורות לא יזוזו
    toDelete.sort(function (a, b) { return b - a; });
    if (!DRY_RUN) toDelete.forEach(function (r) { sh.deleteRow(r); });

    out.push(cfg.sheet + ': ' + (DRY_RUN ? 'זוהו ' : 'נמחקו ') + toDelete.length + ' כפילויות');
  });

  const msg = (DRY_RUN ? '— הרצה יבשה, לא נמחק דבר —\n' : '') + out.join('\n');
  Logger.log(msg);
  return msg;   // מוצג ב-Execution log
}

/** אותה נורמליזציה שבה משתמשת האפליקציה כדי לחסום כפילויות */
function normKey_(s) {
  return String(s || '').replace(/[\s"'׳״.,־–—]+/g, ' ').trim().toLowerCase();
}

function toTime_(v) {
  if (v instanceof Date) return v.getTime();
  const d = new Date(v);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}
