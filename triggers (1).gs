// ============================================================
// triggers.gs — הגדרת טריגרים אוטומטיים
// הרץ setupTriggers() פעם אחת בלבד מעורך הקוד
// ============================================================

function setupTriggers() {
  // מחיקת טריגרים קיימים למניעת כפילויות
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // סינכרון קלנדר כל לילה בשעה 02:00
  ScriptApp.newTrigger('syncCalendar')
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .inTimezone('Asia/Jerusalem')
    .create();

  Logger.log('✅ טריגר לילי הוגדר: syncCalendar בשעה 02:00 (שעון ישראל)');
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log('🗑️ כל הטריגרים נמחקו');
}
