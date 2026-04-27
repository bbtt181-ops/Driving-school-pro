// ============================================================
// calendar_sync.gs — שאיבת שיעורים מגוגל קלנדר
// רץ אוטומטית כל לילה + אפשר להריץ ידנית
// ============================================================

const SS_ID = '14V83uNXOm3FO7mt62gClVJy5DVlKMJ9Sp6I9nfUzRjM';
const ALERT_EMAIL = 'bbtt181@gmail.com';

// ─────────────────────────────────────────
// פונקציה ראשית — נקראת מהטריגר
// ─────────────────────────────────────────
function syncCalendar() {
  const cal  = CalendarApp.getDefaultCalendar();
  const now  = new Date();
  const from = new Date(now.getTime() - 25 * 60 * 60 * 1000); // 25 שעות אחורה
  const to   = now;

  const events = cal.getEvents(from, to);
  Logger.log('📅 נמצאו ' + events.length + ' אירועים לסריקה');

  let synced = 0, skipped = 0, needsReview = 0;

  events.forEach(event => {
    const result = _processEvent(event);
    if (result === 'synced')  synced++;
    if (result === 'skipped') skipped++;
    if (result === 'review')  needsReview++;
  });

  Logger.log(`✅ סנכרון הושלם | נרשמו: ${synced} | דולגו: ${skipped} | דורשים בירור: ${needsReview}`);

  if (needsReview > 0) {
    _sendReviewAlert(needsReview);
  }
}

// ─────────────────────────────────────────
// עיבוד אירוע בודד
// ─────────────────────────────────────────
function _processEvent(event) {
  const eventId = event.getId();
  const title   = event.getTitle().trim();
  const start   = event.getStartTime();
  const end     = event.getEndTime();
  const duration = Math.round((end - start) / 60000); // דקות

  // בדיקה שהאירוע לא נרשם כבר
  if (_lessonExists(eventId)) return 'skipped';

  // זיהוי סוג שיעור לפי משך
  const lessonType = _classifyDuration(duration);
  if (!lessonType) return 'skipped'; // לא שיעור נהיגה

  // חיפוש תלמיד לפי שם
  const lookup = findStudentByName(title);

  if (lookup.match === 'exact') {
    const status = _detectStatus(title);
    _writeLesson(event, lookup.student, lessonType, 'קלנדר-אוטו', status);
    return 'synced';
  }

  if (lookup.match === 'multiple') {
    // כמה תלמידים עם אותו שם — רושם לבירור
    _writePendingLesson(event, title, lessonType, 'דורש בירור');
    return 'review';
  }

  if (lookup.match === 'none') {
    // שם לא מוכר — יכול להיות אירוע אישי, מדלגים
    return 'skipped';
  }
}

// ─────────────────────────────────────────
// זיהוי סטטוס שיעור מכותרת האירוע
// דוגמאות: "ירון אלמוג — בוטל" / "ירון אלמוג — לא הגיע"
// ─────────────────────────────────────────
function _detectStatus(title) {
  const t = title.toLowerCase();
  if (t.includes('בוטל') || t.includes('ביטול')) return 'בוטל';
  if (t.includes('לא הגיע') || t.includes('לא הגיעה')) return 'לא הגיע';
  return 'בוצע';
}

// ─────────────────────────────────────────
// סיווג משך → סוג שיעור
// ─────────────────────────────────────────
function _classifyDuration(minutes) {
  if (minutes >= 35 && minutes <= 55)   return 'בודד 40 דק';
  if (minutes >= 80 && minutes <= 100)  return 'שעה וחצי 90 דק';
  if (minutes >= 110 && minutes <= 130) return 'כפול 120 דק';
  return null; // לא שיעור נהיגה
}

// ─────────────────────────────────────────
// רישום שיעור מזוהה לגיליון שיעורים
// ─────────────────────────────────────────
function _writeLesson(event, student, lessonType, source, status) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('שיעורים');

  const date     = Utilities.formatDate(event.getStartTime(), 'Asia/Jerusalem', 'dd/MM/yyyy');
  const time     = Utilities.formatDate(event.getStartTime(), 'Asia/Jerusalem', 'HH:mm');
  const price    = student.price;

  sh.appendRow([
    event.getId(),
    student.internalId,
    student.fullName,
    date, time,
    lessonType,
    price,
    status,
    source
  ]);

  Logger.log(`✅ נרשם שיעור: ${student.fullName} | ${date} ${time} | ${lessonType} | ${price}₪`);
}

// ─────────────────────────────────────────
// רישום שיעור שדורש בירור (שם כפול)
// ─────────────────────────────────────────
function _writePendingLesson(event, title, lessonType, status) {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName('שיעורים');

  const date = Utilities.formatDate(event.getStartTime(), 'Asia/Jerusalem', 'dd/MM/yyyy');
  const time = Utilities.formatDate(event.getStartTime(), 'Asia/Jerusalem', 'HH:mm');

  sh.appendRow([
    event.getId(),
    '?',
    title + ' (כפילות שם)',
    date, time,
    lessonType,
    0,
    status,
    'קלנדר-אוטו'
  ]);
}

// ─────────────────────────────────────────
// בדיקה שהאירוע לא נרשם כבר
// ─────────────────────────────────────────
function _lessonExists(eventId) {
  const ss   = SpreadsheetApp.openById(SS_ID);
  const sh   = ss.getSheetByName('שיעורים');
  const data = sh.getRange('A2:A' + sh.getLastRow()).getValues();
  return data.some(row => row[0] === eventId);
}

// ─────────────────────────────────────────
// שליחת התראה על שיעורים שדורשים בירור
// ─────────────────────────────────────────
function _sendReviewAlert(count) {
  const shUrl = `https://docs.google.com/spreadsheets/d/${SS_ID}/edit`;
  GmailApp.sendEmail(
    ALERT_EMAIL,
    `⚠️ ${count} שיעורים דורשים זיהוי ידני`,
    '',
    {
      htmlBody: `
        <div dir="rtl" style="font-family:Arial;font-size:14px;">
          <h3>שלום ברוך,</h3>
          <p>נמצאו <strong>${count} שיעורים</strong> שלא זוהו אוטומטית (שם כפול).</p>
          <p>יש לפתוח את גיליון השיעורים ולעדכן ידנית.</p>
          <a href="${shUrl}" style="background:#1B3A5C;color:white;padding:10px 20px;border-radius:6px;text-decoration:none;">
            פתח גיליון שיעורים
          </a>
        </div>
      `
    }
  );
}

// ─────────────────────────────────────────
// הרצה ידנית — מה-48 שעות האחרונות
// ─────────────────────────────────────────
function syncCalendarManual() {
  const cal  = CalendarApp.getDefaultCalendar();
  const now  = new Date();
  const from = new Date(now.getTime() - 48 * 60 * 60 * 1000);

  const events = cal.getEvents(from, now);
  Logger.log('▶️ סנכרון ידני — ' + events.length + ' אירועים');

  events.forEach(event => _processEvent(event));
  Logger.log('✅ סנכרון ידני הושלם');
}
